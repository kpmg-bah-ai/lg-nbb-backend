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
import { detectFormat, ingest } from '../lg/ingest';

/** Cosmos documents are capped at 2MB — store only the first errors plus the total count. */
const MAX_STORED_ERRORS = 100;

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
    const { buffer, filename } = await readUpload(request);
    if (buffer.length === 0) {
        return badRequest(
            'Send the breakdown as a multipart "file" field or as the raw request body (with ?filename=…)'
        );
    }
    const format = detectFormat(buffer, filename);
    const result = await ingest(buffer, { filename, format });
    // Header/file-level problems mean no rows were mapped — reject rather than store an empty run (F1).
    if (result.errors.some((e) => e.code === 'MISSING_HEADER' || e.code === 'EMPTY_INPUT')) {
        return error(422, 'The file is not a recognisable transaction breakdown', { errors: result.errors });
    }
    const run = await lgRuns.create({
        filename,
        format,
        inputSha256: createHash('sha256').update(buffer).digest('hex'),
        uploadedBy: user.id,
        summary: result.summary,
        errorCount: result.errors.length,
        errors: result.errors.slice(0, MAX_STORED_ERRORS),
    });
    await recordAudit(user.id, 'lg.breakdown.ingested', 'lgRun', run.id, {
        filename,
        inputSha256: run.inputSha256,
        dataRows: result.summary.dataRows,
        parsed: result.summary.parsed,
        netFils: result.summary.netFils,
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
