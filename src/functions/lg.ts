/**
 * LG reconciliation — HTTP surface for the ingest pipeline (GOAL.md §4 F1/F9/F10,
 * GOAL-2 G3/G5).
 *
 *   POST /api/lg/runs                 — upload a transaction breakdown (one or more
 *                                       multipart "file" fields — multiple files pool
 *                                       into ONE combined run — or the raw bytes with
 *                                       ?filename=…); parses it and stores the run.
 *                                       Header/file-level failures come back as 422.
 *   GET  /api/lg/runs                 — list stored runs (paged).
 *   GET  /api/lg/runs/{id}            — one run with its summary and (capped) errors.
 *   GET  /api/lg/runs/{id}/matched    — the run's cleared sets (offset-paged, G3).
 *   GET  /api/lg/runs/{id}/exceptions — the run's classified exceptions (offset-paged, G3).
 *   GET  /api/lg/runs/{id}/export     — the authoritative two-sheet .xlsx statement (G5).
 *
 * Runs are file uploads and immutable, so they don't fit registerCrudRoutes —
 * handlers are hand-rolled and exported for test/lg/api.test.ts.
 */

import { createHash, randomUUID } from 'node:crypto';
import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { lgRunDetails, lgRuns } from '../data/repositories';
import { recordAudit } from '../helpers/audit';
import { requireRole } from '../helpers/auth';
import { badRequest, created, error, json, notFound } from '../helpers/json';
import { computeBranchBalances, deriveAsOf } from '../lg/balance';
import { detectExceptions } from '../lg/exceptions';
import { buildStatementWorkbook } from '../lg/export';
import { validateGlUpload } from '../lg/glGuard';
import { detectFormat, ingestFiles } from '../lg/ingest';
import { matchPostings } from '../lg/match';
import { reconcile } from '../lg/reconcile';
import { classifyRegisterExceptions } from '../lg/registerExceptions';
import { matchRegister } from '../lg/registerMatch';
import { extractStatedBalance, reconcileRegister } from '../lg/registerReconcile';
import { reconcileStatement } from '../lg/statementReconcile';
import { computeSheetBalances } from '../lg/sheetBalances';
import { explainRun } from '../lg/provenance';
import {
    ChequeOutcome,
    ChequeState,
    ExceptionSummary,
    GL_CATALOG,
    LedgerRow,
    LgException,
    LgRun,
    LgRunDetailChunk,
    LgRunFile,
    MatchedSet,
    MatchSummary,
    OutstandingItem,
    ParseError,
    Reconciliation,
    resolveGlCode,
} from '../shared/models';

/** Cosmos documents are capped at 2MB — store only the first errors plus the total count. */
const MAX_STORED_ERRORS = 100;

/** Same 2MB cap: outstanding items beyond this are counted but not stored on the run. */
const MAX_STORED_OUTSTANDING = 500;

/** Same 2MB cap for per-branch balances (defensive — branch counts are normally small). */
const MAX_STORED_BALANCES = 500;

/** Items per lgRunDetails chunk document — keeps each chunk far below the 2MB cap. */
const DETAIL_CHUNK_SIZE = 250;

/**
 * Byte budget per chunk document. Cosmos rejects requests over 2MB ("Request size
 * is too large"); item counts alone don't bound bytes because one matched set can
 * carry very many legs, so chunks are ALSO split by serialized size.
 */
const DETAIL_CHUNK_MAX_BYTES = 1_500_000;

/**
 * Legs stored per side of a matched set. A busy sub-account FIFO-chains thousands
 * of postings into one component — storing every leg would blow the document cap.
 * `creditLegCount`/`debitLegCount` keep the true totals, so the cap is visible.
 */
const MAX_STORED_SET_LEGS = 50;

/**
 * Ceiling on stored matched sets / exceptions per run (env-overridable). Totals are
 * always reported (`matchedSetCount` / `exceptionCount`), so truncation is visible,
 * never silent (GOAL-2 §6).
 */
function maxStoredDetailItems(): number {
    return Number(process.env.LG_MAX_DETAIL_ITEMS) || 20_000;
}

/** Caps a set's stored legs; the leg counts on the set keep the true totals. */
export function trimMatchedSet(set: MatchedSet): MatchedSet {
    if (set.creditLegs.length <= MAX_STORED_SET_LEGS && set.debitLegs.length <= MAX_STORED_SET_LEGS) {
        return set;
    }
    return {
        ...set,
        creditLegs: set.creditLegs.slice(0, MAX_STORED_SET_LEGS),
        debitLegs: set.debitLegs.slice(0, MAX_STORED_SET_LEGS),
    };
}

/** Splits items into chunks bounded by BOTH item count and serialized byte size. */
export function buildDetailChunks<T>(
    items: T[],
    maxItems = DETAIL_CHUNK_SIZE,
    maxBytes = DETAIL_CHUNK_MAX_BYTES
): T[][] {
    const chunks: T[][] = [];
    let chunk: T[] = [];
    let chunkBytes = 0;
    for (const item of items) {
        const itemBytes = Buffer.byteLength(JSON.stringify(item));
        if (chunk.length > 0 && (chunk.length >= maxItems || chunkBytes + itemBytes > maxBytes)) {
            chunks.push(chunk);
            chunk = [];
            chunkBytes = 0;
        }
        chunk.push(item);
        chunkBytes += itemBytes;
    }
    if (chunk.length > 0) {
        chunks.push(chunk);
    }
    return chunks;
}

/** Writes detail items as ordered chunk documents; returns how many items were stored. */
async function storeDetailChunks(
    runId: string,
    kind: LgRunDetailChunk['kind'],
    items: MatchedSet[] | LgException[] | ChequeOutcome[] | LedgerRow[]
): Promise<number> {
    const capped = items.slice(0, maxStoredDetailItems());
    const stored = kind === 'matchedSets' ? (capped as MatchedSet[]).map(trimMatchedSet) : capped;
    const chunks = buildDetailChunks<MatchedSet | LgException | ChequeOutcome | LedgerRow>(stored);
    for (let seq = 0; seq < chunks.length; seq++) {
        await lgRunDetails.create({ runId, kind, seq, items: chunks[seq] as LgRunDetailChunk['items'] });
    }
    return stored.length;
}

/** Reads every stored detail item of one kind for a run, in chunk order. */
async function readDetailItems<T extends MatchedSet | LgException | ChequeOutcome | LedgerRow>(
    runId: string,
    kind: LgRunDetailChunk['kind']
): Promise<T[]> {
    const chunks = await lgRunDetails.query<LgRunDetailChunk>({
        query: 'SELECT * FROM c WHERE c.runId = @runId AND c.kind = @kind ORDER BY c.seq ASC',
        parameters: [
            { name: '@runId', value: runId },
            { name: '@kind', value: kind },
        ],
    });
    return chunks.flatMap((chunk) => chunk.items as T[]);
}

/**
 * Upload ceiling; overridable via LG_MAX_UPLOAD_BYTES (tests/ops). Defaults to the
 * Functions host's own 100MB request-body cap — a larger value here would never be
 * reached in the cloud. The prod breakdown is ~57MB.
 */
function maxUploadBytes(): number {
    return Number(process.env.LG_MAX_UPLOAD_BYTES) || 100 * 1024 * 1024;
}

/** Strict calendar-checked ISO date (rejects '2026-02-31' and '2026-13-01'). */
function isIsoDate(text: string): boolean {
    const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) {
        return false;
    }
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const date = new Date(Date.UTC(y, mo - 1, d));
    return date.getUTCFullYear() === y && date.getUTCMonth() === mo - 1 && date.getUTCDate() === d;
}

/** File/header-level codes that mean nothing was mappable — reject, don't store. */
const REJECT_CODES = new Set([
    'MISSING_HEADER',
    'EMPTY_INPUT',
    'UNSUPPORTED_FORMAT',
    'MIXED_MODE',
    'INCOMPLETE_REGISTER_INPUT',
    'UNKNOWN_GL',
    'GL_MISMATCH',
]);

interface Upload {
    buffer: Buffer;
    filename?: string;
}

/**
 * Accepts multipart form-data (one or MORE "file" fields — a multi-file upload
 * becomes one combined run) or the raw file bytes as the body (single file).
 */
async function readUpload(request: HttpRequest): Promise<Upload[]> {
    const contentType = request.headers.get('content-type') ?? '';
    if (contentType.includes('multipart/form-data')) {
        const form = await request.formData();
        const uploads: Upload[] = [];
        for (const entry of form.getAll('file')) {
            if (typeof entry === 'string') {
                continue; // a text field named "file" is not an upload
            }
            uploads.push({ buffer: Buffer.from(await entry.arrayBuffer()), filename: entry.name || undefined });
        }
        return uploads;
    }
    return [
        {
            buffer: Buffer.from(await request.arrayBuffer()),
            filename: request.query.get('filename') ?? undefined,
        },
    ];
}

export async function createLgRun(request: HttpRequest): Promise<HttpResponseInit> {
    const { user, response } = await requireRole(request, 'staff');
    if (response) {
        return response;
    }
    const asOfParam = request.query.get('asOf');
    if (asOfParam && !isIsoDate(asOfParam)) {
        return badRequest('asOf must be a valid ISO date (yyyy-mm-dd)');
    }
    // GOAL-7: which catalog GL the user picked. Mandatory — the picker sends it.
    // Validated cheaply, before the (potentially large) upload is read.
    const glParam = request.query.get('gl');
    if (!glParam) {
        return badRequest(`Pass ?gl= with the GL picked at upload (one of: ${Object.keys(GL_CATALOG).join(', ')})`);
    }
    const glCode = resolveGlCode(glParam);
    if (!glCode) {
        return error(422, `GL "${glParam}" is not in the GL catalog`, {
            errors: [
                {
                    code: 'UNKNOWN_GL',
                    message: `GL "${glParam}" is not in the GL catalog — known GLs: ${Object.keys(GL_CATALOG).join(', ')}`,
                },
            ],
        });
    }
    const glDef = GL_CATALOG[glCode];
    let uploads: Upload[];
    try {
        uploads = await readUpload(request);
    } catch {
        return badRequest('The request body could not be read — check the multipart encoding');
    }
    if (uploads.length === 0 || uploads.every((u) => u.buffer.length === 0)) {
        return badRequest(
            'Send the breakdown as one or more multipart "file" fields or as the raw request body (with ?filename=…)'
        );
    }
    const empty = uploads.find((u) => u.buffer.length === 0);
    if (empty) {
        return badRequest(`The uploaded file "${empty.filename ?? '(unnamed)'}" is empty`);
    }
    const totalBytes = uploads.reduce((sum, u) => sum + u.buffer.length, 0);
    if (totalBytes > maxUploadBytes()) {
        return error(413, `The upload exceeds the ${Math.floor(maxUploadBytes() / (1024 * 1024))}MB limit`);
    }
    // Per-file identity, kept on the run for provenance when there is more than one file.
    const files: LgRunFile[] = uploads.map((u) => ({
        filename: u.filename,
        format: detectFormat(u.buffer, u.filename),
        sha256: createHash('sha256').update(u.buffer).digest('hex'),
        bytes: u.buffer.length,
    }));
    const filename =
        uploads.length === 1 ? uploads[0].filename : files.map((f) => f.filename ?? '(unnamed)').join(' + ');
    const format = files[0].format;
    const result = await ingestFiles(uploads.map((u) => ({ buffer: u.buffer, filename: u.filename })));
    // Header/file-level problems mean no rows were mapped — reject rather than store an empty run (F1).
    if (result.errors.some((e) => REJECT_CODES.has(e.code))) {
        return error(422, 'The file is not a recognisable transaction breakdown', { errors: result.errors });
    }

    // GOAL-7 §3: the upload must BE the GL the user said it is (family + embedded GL code).
    const glErrors = validateGlUpload(glDef, result);
    if (glErrors.length > 0) {
        return error(422, 'The upload does not match the picked GL', { errors: glErrors });
    }

    // F3 + F4 run at ingest, while the postings are in memory (they are not persisted).
    // A run whose rows all failed row-level parsing has no postings — skip matching
    // rather than storing a meaningless '1970-01-01' summary.
    const asOf = asOfParam ?? deriveAsOf(result.postings);
    const balances = computeBranchBalances(result.postings, asOf);
    const errors: ParseError[] = [...result.errors];

    let match:
        | { outstanding: OutstandingItem[]; matchedSets: MatchedSet[]; summary: MatchSummary }
        | undefined;
    let reconciliation: Reconciliation | undefined;
    let exceptions: { exceptions: LgException[]; summary: ExceptionSummary } | undefined;
    let outcomes: ChequeOutcome[] | undefined;
    let registerFields: Pick<LgRun, 'chequeCount' | 'chequesByState' | 'preWindowChequeCount'> | undefined;
    let ledgerRows: LedgerRow[] | undefined;

    if (result.postings.length > 0) {
        if (result.mode === 'register') {
            // GOAL-3: two-legged GL↔register matching; the stated EoD balance
            // drives the reconciliation block so the extract gap surfaces.
            const registerMatch = matchRegister(result.postings, result.cheques ?? [], { asOf });
            const stated = extractStatedBalance(result.postings);
            if (stated.error) {
                errors.push(stated.error);
            }
            reconciliation = reconcileRegister(stated.statedFils, balances, registerMatch, { asOf });
            exceptions = classifyRegisterExceptions(registerMatch, reconciliation.byBranch[0]?.extractGapFils);
            match = registerMatch;
            outcomes = registerMatch.outcomes;
            const chequesByState: Partial<Record<ChequeState, number>> = {};
            for (const outcome of outcomes) {
                chequesByState[outcome.state] = (chequesByState[outcome.state] ?? 0) + 1;
            }
            registerFields = {
                chequeCount: (result.cheques ?? []).length,
                chequesByState,
                preWindowChequeCount: chequesByState.PRE_WINDOW ?? 0,
            };
        } else if (result.mode === 'statement') {
            // GOAL-8: no matching, no outstanding, no exceptions. Reconcile the derived
            // net against the stated EoD; the tie-out gap (when non-zero) surfaces via
            // the `extractGap` explained figure (flagged) and reconciliation.balanced —
            // no synthetic exception is needed (and the real VAT files tie out exactly).
            const stated = extractStatedBalance(result.postings);
            if (stated.error) {
                errors.push(stated.error);
            }
            reconciliation = reconcileStatement(stated.statedFils, balances, { asOf });
            ledgerRows = result.postings.map((p) => ({
                rowNumber: p.rowNumber,
                postDate: p.postDate,
                transactionDate: p.transactionDate,
                description: p.logDescription,
                branchNumber: p.branchNumber,
                journalNumber: p.journalNumber,
                direction: p.direction,
                amountBhdFils: p.amountBhdFils,
                statedEodFils: p.statedEodFils,
            }));
        } else {
            match = matchPostings(result.postings, { asOf });
            // F5: Difference & Balanced per branch, from the full (uncapped) outstanding list.
            reconciliation = reconcile(balances, match.outstanding, { asOf });
            // F6 (GOAL-2 G2): classify the outstanding items into reviewer-facing exceptions.
            exceptions = detectExceptions(match.outstanding);
        }
    }

    // GOAL-5: per-sheet balance reference + a plain-language basis/assessment for
    // every headline number. Both are derived from figures already computed above
    // (they never re-derive money), so they tie to the screen by construction and
    // are saved on the run as the reviewer's reference.
    const sheetBalances = computeSheetBalances(result);
    const explanations = explainRun({
        mode: result.mode,
        summary: result.summary,
        asOf,
        balances,
        sheetBalances,
        reconciliation,
        matching: match?.summary,
        exceptionsSummary: exceptions?.summary,
        chequeCount: registerFields?.chequeCount,
        chequesByState: registerFields?.chequesByState,
    });

    // Input identity (GOAL.md §5 determinism). Single file: hash of the raw bytes,
    // unchanged from pre-multi-file runs. Multi-file: hash of the SORTED per-file
    // hashes, so the same set of files dedupes regardless of upload order.
    const inputSha256 =
        uploads.length === 1
            ? files[0].sha256
            : createHash('sha256')
                  .update(
                      files
                          .map((f) => f.sha256)
                          .sort()
                          .join('')
                  )
                  .digest('hex');
    // Same-input detection (GOAL.md §5 re-runnability): link, but still store the re-run.
    const duplicates = await lgRuns.query<{ id: string }>({
        query: 'SELECT c.id FROM c WHERE c.inputSha256 = @hash ORDER BY c.createdAt ASC',
        parameters: [{ name: '@hash', value: inputSha256 }],
    });

    // The run id is minted up front so the detail chunks can be written BEFORE the
    // run document: if a chunk write fails, no half-ingested run becomes visible to
    // clients (any already-written chunks reference a runId that no run document
    // carries — inert leftovers, invisible to the by-run queries).
    const runId = randomUUID();
    // G3: matched sets + exceptions (+ GOAL-3 cheque outcomes) are chunked into
    // lgRunDetails (capped, totals visible).
    if (match && match.matchedSets.length > 0) {
        await storeDetailChunks(runId, 'matchedSets', match.matchedSets);
    }
    if (exceptions && exceptions.exceptions.length > 0) {
        await storeDetailChunks(runId, 'exceptions', exceptions.exceptions);
    }
    if (outcomes && outcomes.length > 0) {
        await storeDetailChunks(runId, 'cheques', outcomes);
    }
    if (ledgerRows && ledgerRows.length > 0) {
        await storeDetailChunks(runId, 'ledger', ledgerRows);
    }
    const run = await lgRuns.create({
        id: runId,
        filename,
        format,
        inputSha256,
        // Per-file provenance only for multi-file uploads — single-file run docs
        // keep their historic shape.
        ...(uploads.length > 1 ? { files } : {}),
        duplicateOf: duplicates[0]?.id,
        uploadedBy: user.id,
        mode: result.mode,
        glCode: glDef.code,
        summary: result.summary,
        errorCount: errors.length,
        errors: errors.slice(0, MAX_STORED_ERRORS),
        asOf,
        balances: balances.slice(0, MAX_STORED_BALANCES),
        balancesCount: balances.length,
        matching: match?.summary,
        outstandingCount: match ? match.outstanding.length : 0,
        outstanding: match ? match.outstanding.slice(0, MAX_STORED_OUTSTANDING) : [],
        reconciliation,
        matchedSetCount: match ? match.matchedSets.length : 0,
        exceptionCount: exceptions ? exceptions.summary.total : 0,
        exceptionsSummary: exceptions?.summary,
        sheetBalances,
        explanations,
        ...registerFields,
        ledgerRowCount: ledgerRows ? ledgerRows.length : undefined,
    });
    await recordAudit(user.id, 'lg.breakdown.ingested', 'lgRun', run.id, {
        filename,
        fileCount: uploads.length,
        inputSha256,
        mode: result.mode,
        glCode: glDef.code,
        dataRows: result.summary.dataRows,
        parsed: result.summary.parsed,
        netFils: result.summary.netFils,
        asOf,
        outstandingCount: match ? match.outstanding.length : 0,
        matchedSetCount: match ? match.matchedSets.length : 0,
        exceptionCount: exceptions ? exceptions.summary.total : 0,
        chequeCount: registerFields?.chequeCount,
        balanced: reconciliation?.balanced,
        sheetCount: sheetBalances.length,
        explainedFigureCount: explanations.length,
        duplicateOf: duplicates[0]?.id,
    });
    return created(run);
}

export async function listLgRuns(request: HttpRequest): Promise<HttpResponseInit> {
    const { response } = await requireRole(request, 'staff');
    if (response) {
        return response;
    }
    const result = await lgRuns.list({
        maxItems: Number(request.query.get('maxItems')) || undefined,
        continuationToken: request.query.get('continuationToken') ?? undefined,
    });
    return json(result);
}

export async function getLgRun(request: HttpRequest): Promise<HttpResponseInit> {
    const { response } = await requireRole(request, 'staff');
    if (response) {
        return response;
    }
    const run = await lgRuns.get(request.params.id);
    return run ? json(run) : notFound('Run not found');
}

/** Shared offset-paged reader over a run's stored detail items (G3). */
function detailListHandler(kind: LgRunDetailChunk['kind'], totalOf: (run: LgRun) => number) {
    return async (request: HttpRequest): Promise<HttpResponseInit> => {
        const { response } = await requireRole(request, 'staff');
        if (response) {
            return response;
        }
        const run = await lgRuns.get(request.params.id);
        if (!run) {
            return notFound('Run not found');
        }
        const offset = Math.max(0, Number(request.query.get('offset')) || 0);
        const maxItems = Math.min(Math.max(1, Number(request.query.get('maxItems')) || 100), 1000);
        const stored = await readDetailItems(run.id, kind);
        return json({
            items: stored.slice(offset, offset + maxItems),
            offset,
            maxItems,
            /** How many items are actually stored (≤ total when the cap truncated). */
            storedCount: stored.length,
            /** The true engine total — storedCount < total means visible truncation. */
            total: totalOf(run),
        });
    };
}

export const listLgRunMatched = detailListHandler('matchedSets', (run) => run.matchedSetCount ?? 0);
export const listLgRunExceptions = detailListHandler('exceptions', (run) => run.exceptionCount ?? 0);
/** GOAL-3 R8: per-cheque outcomes of a register-mode run, offset-paged. */
export const listLgRunCheques = detailListHandler('cheques', (run) => run.chequeCount ?? 0);
/** GOAL-8: the persisted statement ledger rows of a statement-mode run, offset-paged. */
export const listLgRunLedger = detailListHandler('ledger', (run) => run.ledgerRowCount ?? 0);

/**
 * G5: the authoritative export — a two-sheet workbook (statement + Mismatched) built
 * server-side from the stored run + its full exception set, per GOAL.md §2.2/§2.3.
 * Scope: ?entity=&gl=&branch= select the reconciliation block; all optional when the
 * run has exactly one block.
 */
export async function exportLgRun(request: HttpRequest): Promise<HttpResponseInit> {
    const { response } = await requireRole(request, 'staff');
    if (response) {
        return response;
    }
    const run = await lgRuns.get(request.params.id);
    if (!run) {
        return notFound('Run not found');
    }
    const blocks = run.reconciliation?.byBranch ?? [];
    if (blocks.length === 0) {
        return badRequest('This run has no reconciliation to export');
    }
    // GOAL-8 §2(G): a statement run has no outstanding-items population to export.
    if (run.mode === 'statement') {
        return badRequest('This GL is a running-balance statement with no outstanding-items export — see the Ledger and Basis tabs.');
    }
    const entity = request.query.get('entity');
    const gl = request.query.get('gl');
    const branch = request.query.get('branch');

    // GOAL-3 R9: register-mode runs carry ONE consolidated GL-level block;
    // ?branch= narrows the statement SECTIONS, never the block.
    if (run.mode === 'register') {
        const exceptions = await readDetailItems<LgException>(run.id, 'exceptions');
        const outcomes = await readDetailItems<ChequeOutcome>(run.id, 'cheques');
        const branchFilter = branch ?? undefined;
        if (branchFilter !== undefined && !outcomes.some((o) => (o.issuedBranch ?? '') === branchFilter)) {
            return notFound('No register cheques belong to the requested branch');
        }
        const buffer = await buildStatementWorkbook(run, blocks[0], exceptions, outcomes, branchFilter);
        const scope = branchFilter !== undefined ? `Branch-${branchFilter}` : 'Consolidated';
        const filename = `GL-Recon_${scope}_${run.reconciliation?.asOf ?? run.asOf ?? 'draft'}.xlsx`;
        return {
            status: 200,
            body: new Uint8Array(buffer),
            headers: {
                'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'content-disposition': `attachment; filename="${filename}"`,
            },
        };
    }

    const matches = blocks.filter(
        (b) =>
            (entity === null || b.entity === entity) &&
            (gl === null || b.gl === gl) &&
            (branch === null || b.branchNumber === branch)
    );
    if (matches.length === 0) {
        return notFound('No reconciliation block matches the requested branch');
    }
    if (matches.length > 1) {
        return badRequest(
            'The run covers several branches — pass ?branch= (and ?entity=/?gl= if needed) to pick one',
            { branches: matches.map((b) => ({ entity: b.entity, gl: b.gl, branchNumber: b.branchNumber })) }
        );
    }
    const exceptions = await readDetailItems<LgException>(run.id, 'exceptions');
    const buffer = await buildStatementWorkbook(run, matches[0], exceptions);
    const filename = `GL-Recon_Branch-${matches[0].branchNumber}_${run.reconciliation?.asOf ?? run.asOf ?? 'draft'}.xlsx`;
    return {
        status: 200,
        body: new Uint8Array(buffer),
        headers: {
            'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'content-disposition': `attachment; filename="${filename}"`,
        },
    };
}

// authLevel is 'anonymous' — NOT because these routes are unprotected, but because
// the SPA (Azure Static Web App) calls this Function App directly at its public URL
// (VITE_API_BASE points straight at *.azurewebsites.net; there is no SWA managed
// backend to inject a function key). A browser can't hold a function key, so
// authLevel:'function' makes the HOST reject every request with an empty-body 401
// before any handler runs. Identity/role are enforced in-app by requireRole() via
// the x-user-email header (placeholder until Entra ID — see helpers/auth.ts).
app.http('lg-runs-create', { route: 'lg/runs', methods: ['POST'], authLevel: 'anonymous', handler: createLgRun });
app.http('lg-runs-list', { route: 'lg/runs', methods: ['GET'], authLevel: 'anonymous', handler: listLgRuns });
app.http('lg-runs-get', { route: 'lg/runs/{id}', methods: ['GET'], authLevel: 'anonymous', handler: getLgRun });
app.http('lg-runs-matched', {
    route: 'lg/runs/{id}/matched',
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: listLgRunMatched,
});
app.http('lg-runs-exceptions', {
    route: 'lg/runs/{id}/exceptions',
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: listLgRunExceptions,
});
app.http('lg-runs-cheques', {
    route: 'lg/runs/{id}/cheques',
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: listLgRunCheques,
});
app.http('lg-runs-ledger', {
    route: 'lg/runs/{id}/ledger',
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: listLgRunLedger,
});
app.http('lg-runs-export', {
    route: 'lg/runs/{id}/export',
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: exportLgRun,
});
