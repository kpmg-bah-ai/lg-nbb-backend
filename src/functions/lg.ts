/**
 * LG reconciliation — HTTP surface for the ingest pipeline (GOAL.md §4 F1/F9/F10).
 *
 *   POST /api/lg/runs      — upload a transaction breakdown (multipart "file" field,
 *                            or the raw bytes with ?filename=…); parses it and stores
 *                            the run. Header/file-level failures come back as 422.
 *   GET  /api/lg/runs      — list stored runs (paged).
 *   GET  /api/lg/runs/{id} — one run with its summary and (capped) errors.
 *
 * Runs are file uploads and immutable, so they don't fit registerCrudRoutes —
 * handlers are hand-rolled and exported for test/lg/api.test.ts. Matching and
 * reconciliation endpoints (F4/F5) will attach to the same run resource.
 */

import { createHash } from 'node:crypto';
import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { lgRuns } from '../data/repositories';
import { recordAudit } from '../helpers/audit';
import { requireRole } from '../helpers/auth';
import { badRequest, created, error, json, notFound } from '../helpers/json';
import { computeBranchBalances, deriveAsOf } from '../lg/balance';
import { detectFormat, ingest } from '../lg/ingest';
import { matchPostings } from '../lg/match';
import { reconcile } from '../lg/reconcile';

/** Cosmos documents are capped at 2MB — store only the first errors plus the total count. */
const MAX_STORED_ERRORS = 100;

/** Same 2MB cap: outstanding items beyond this are counted but not stored on the run. */
const MAX_STORED_OUTSTANDING = 500;

/** Same 2MB cap for per-branch balances (defensive — branch counts are normally small). */
const MAX_STORED_BALANCES = 500;

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
    });
    await recordAudit(user.id, 'lg.breakdown.ingested', 'lgRun', run.id, {
        filename,
        inputSha256,
        dataRows: result.summary.dataRows,
        parsed: result.summary.parsed,
        netFils: result.summary.netFils,
        asOf,
        outstandingCount: match ? match.outstanding.length : 0,
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

app.http('lg-runs-create', { route: 'lg/runs', methods: ['POST'], authLevel: 'function', handler: createLgRun });
app.http('lg-runs-list', { route: 'lg/runs', methods: ['GET'], authLevel: 'function', handler: listLgRuns });
app.http('lg-runs-get', { route: 'lg/runs/{id}', methods: ['GET'], authLevel: 'function', handler: getLgRun });
