/**
 * LG reconciliation — HTTP surface for the ingest pipeline (GOAL.md §4 F1/F9/F10,
 * GOAL-2 G3/G5).
 *
 *   POST /api/lg/runs                 — upload a transaction breakdown (multipart "file"
 *                                       field, or the raw bytes with ?filename=…); parses
 *                                       it and stores the run. Header/file-level failures
 *                                       come back as 422.
 *   GET  /api/lg/runs                 — list stored runs (paged).
 *   GET  /api/lg/runs/{id}            — one run with its summary and (capped) errors.
 *   GET  /api/lg/runs/{id}/matched    — the run's cleared sets (offset-paged, G3).
 *   GET  /api/lg/runs/{id}/exceptions — the run's classified exceptions (offset-paged, G3).
 *   GET  /api/lg/runs/{id}/export     — the authoritative two-sheet .xlsx statement (G5).
 *
 * Runs are file uploads and immutable, so they don't fit registerCrudRoutes —
 * handlers are hand-rolled and exported for test/lg/api.test.ts.
 */

import { createHash } from 'node:crypto';
import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { lgRunDetails, lgRuns } from '../data/repositories';
import { recordAudit } from '../helpers/audit';
import { requireRole } from '../helpers/auth';
import { badRequest, created, error, json, notFound } from '../helpers/json';
import { computeBranchBalances, deriveAsOf } from '../lg/balance';
import { detectExceptions } from '../lg/exceptions';
import { buildStatementWorkbook } from '../lg/export';
import { detectFormat, ingest } from '../lg/ingest';
import { matchPostings } from '../lg/match';
import { reconcile } from '../lg/reconcile';
import { LgException, LgRun, LgRunDetailChunk, MatchedSet } from '../shared/models';

/** Cosmos documents are capped at 2MB — store only the first errors plus the total count. */
const MAX_STORED_ERRORS = 100;

/** Same 2MB cap: outstanding items beyond this are counted but not stored on the run. */
const MAX_STORED_OUTSTANDING = 500;

/** Same 2MB cap for per-branch balances (defensive — branch counts are normally small). */
const MAX_STORED_BALANCES = 500;

/** Items per lgRunDetails chunk document — keeps each chunk far below the 2MB cap. */
const DETAIL_CHUNK_SIZE = 250;

/**
 * Ceiling on stored matched sets / exceptions per run (env-overridable). Totals are
 * always reported (`matchedSetCount` / `exceptionCount`), so truncation is visible,
 * never silent (GOAL-2 §6).
 */
function maxStoredDetailItems(): number {
    return Number(process.env.LG_MAX_DETAIL_ITEMS) || 20_000;
}

/** Writes detail items as ordered chunk documents; returns how many items were stored. */
async function storeDetailChunks(
    runId: string,
    kind: LgRunDetailChunk['kind'],
    items: MatchedSet[] | LgException[]
): Promise<number> {
    const capped = items.slice(0, maxStoredDetailItems());
    for (let seq = 0; seq * DETAIL_CHUNK_SIZE < capped.length; seq++) {
        await lgRunDetails.create({
            runId,
            kind,
            seq,
            items: capped.slice(seq * DETAIL_CHUNK_SIZE, (seq + 1) * DETAIL_CHUNK_SIZE),
        });
    }
    return capped.length;
}

/** Reads every stored detail item of one kind for a run, in chunk order. */
async function readDetailItems<T extends MatchedSet | LgException>(
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
const REJECT_CODES = new Set(['MISSING_HEADER', 'EMPTY_INPUT', 'UNSUPPORTED_FORMAT']);

interface Upload {
    buffer: Buffer;
    filename?: string;
}

/** Accepts either multipart form-data (a "file" field) or the raw file bytes as the body. */
async function readUpload(request: HttpRequest): Promise<Upload> {
    const contentType = request.headers.get('content-type') ?? '';
    if (contentType.includes('multipart/form-data')) {
        const form = await request.formData();
        const file = form.get('file');
        if (!file || typeof file === 'string') {
            return { buffer: Buffer.alloc(0) };
        }
        return { buffer: Buffer.from(await file.arrayBuffer()), filename: file.name || undefined };
    }
    return {
        buffer: Buffer.from(await request.arrayBuffer()),
        filename: request.query.get('filename') ?? undefined,
    };
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
    let upload: Upload;
    try {
        upload = await readUpload(request);
    } catch {
        return badRequest('The request body could not be read — check the multipart encoding');
    }
    const { buffer, filename } = upload;
    if (buffer.length === 0) {
        return badRequest(
            'Send the breakdown as a multipart "file" field or as the raw request body (with ?filename=…)'
        );
    }
    if (buffer.length > maxUploadBytes()) {
        return error(413, `The file exceeds the ${Math.floor(maxUploadBytes() / (1024 * 1024))}MB upload limit`);
    }
    const format = detectFormat(buffer, filename);
    const result = await ingest(buffer, { filename, format });
    // Header/file-level problems mean no rows were mapped — reject rather than store an empty run (F1).
    if (result.errors.some((e) => REJECT_CODES.has(e.code))) {
        return error(422, 'The file is not a recognisable transaction breakdown', { errors: result.errors });
    }

    // F3 + F4 run at ingest, while the postings are in memory (they are not persisted).
    // A run whose rows all failed row-level parsing has no postings — skip matching
    // rather than storing a meaningless '1970-01-01' summary.
    const asOf = asOfParam ?? deriveAsOf(result.postings);
    const balances = computeBranchBalances(result.postings, asOf);
    const match = result.postings.length > 0 ? matchPostings(result.postings, { asOf }) : undefined;
    // F5: Difference & Balanced per branch, from the full (uncapped) outstanding list.
    const reconciliation = match ? reconcile(balances, match.outstanding, { asOf }) : undefined;
    // F6 (GOAL-2 G2): classify the outstanding items into reviewer-facing exceptions.
    const exceptions = match ? detectExceptions(match.outstanding) : undefined;

    const inputSha256 = createHash('sha256').update(buffer).digest('hex');
    // Same-input detection (GOAL.md §5 re-runnability): link, but still store the re-run.
    const duplicates = await lgRuns.query<{ id: string }>({
        query: 'SELECT c.id FROM c WHERE c.inputSha256 = @hash ORDER BY c.createdAt ASC',
        parameters: [{ name: '@hash', value: inputSha256 }],
    });

    const run = await lgRuns.create({
        filename,
        format,
        inputSha256,
        duplicateOf: duplicates[0]?.id,
        uploadedBy: user.id,
        summary: result.summary,
        errorCount: result.errors.length,
        errors: result.errors.slice(0, MAX_STORED_ERRORS),
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
    });
    // G3: matched sets + exceptions are chunked into lgRunDetails (capped, totals visible).
    if (match && match.matchedSets.length > 0) {
        await storeDetailChunks(run.id, 'matchedSets', match.matchedSets);
    }
    if (exceptions && exceptions.exceptions.length > 0) {
        await storeDetailChunks(run.id, 'exceptions', exceptions.exceptions);
    }
    await recordAudit(user.id, 'lg.breakdown.ingested', 'lgRun', run.id, {
        filename,
        inputSha256,
        dataRows: result.summary.dataRows,
        parsed: result.summary.parsed,
        netFils: result.summary.netFils,
        asOf,
        outstandingCount: match ? match.outstanding.length : 0,
        matchedSetCount: match ? match.matchedSets.length : 0,
        exceptionCount: exceptions ? exceptions.summary.total : 0,
        balanced: reconciliation?.balanced,
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
    const entity = request.query.get('entity');
    const gl = request.query.get('gl');
    const branch = request.query.get('branch');
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

app.http('lg-runs-create', { route: 'lg/runs', methods: ['POST'], authLevel: 'function', handler: createLgRun });
app.http('lg-runs-list', { route: 'lg/runs', methods: ['GET'], authLevel: 'function', handler: listLgRuns });
app.http('lg-runs-get', { route: 'lg/runs/{id}', methods: ['GET'], authLevel: 'function', handler: getLgRun });
app.http('lg-runs-matched', {
    route: 'lg/runs/{id}/matched',
    methods: ['GET'],
    authLevel: 'function',
    handler: listLgRunMatched,
});
app.http('lg-runs-exceptions', {
    route: 'lg/runs/{id}/exceptions',
    methods: ['GET'],
    authLevel: 'function',
    handler: listLgRunExceptions,
});
app.http('lg-runs-export', {
    route: 'lg/runs/{id}/export',
    methods: ['GET'],
    authLevel: 'function',
    handler: exportLgRun,
});
