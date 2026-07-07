import { HttpRequest } from '@azure/functions';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

jest.mock('../../src/data/repositories', () => ({
    users: { get: jest.fn() },
    auditLogs: { create: jest.fn() },
    lgRuns: { create: jest.fn(), get: jest.fn(), list: jest.fn(), query: jest.fn() },
    lgRunDetails: { create: jest.fn(), query: jest.fn() },
}));

import { auditLogs, lgRunDetails, lgRuns, users } from '../../src/data/repositories';
import {
    createLgRun,
    exportLgRun,
    getLgRun,
    listLgRunExceptions,
    listLgRunMatched,
    listLgRuns,
} from '../../src/functions/lg';
import { LgRun, User } from '../../src/shared/models';

const mockUserGet = users.get as jest.Mock;
const mockRunCreate = lgRuns.create as jest.Mock;
const mockRunGet = lgRuns.get as jest.Mock;
const mockRunList = lgRuns.list as jest.Mock;
const mockRunQuery = lgRuns.query as jest.Mock;
const mockAuditCreate = auditLogs.create as jest.Mock;
const mockDetailCreate = lgRunDetails.create as jest.Mock;
const mockDetailQuery = lgRunDetails.query as jest.Mock;

const fixture = (name: string) => readFileSync(join(__dirname, '..', 'fixtures', 'lg', name));

function staffUser(): User {
    return {
        id: 'u1',
        displayName: 'Staff User',
        email: 'staff@kpmg.com',
        role: 'staff',
        active: true,
        createdAt: '2026-01-01T00:00:00Z',
        updatedAt: '2026-01-01T00:00:00Z',
    };
}

interface FakeRequestOptions {
    headers?: Record<string, string>;
    query?: Record<string, string>;
    params?: Record<string, string>;
    body?: Buffer;
    form?: FormData;
}

function fakeRequest(options: FakeRequestOptions = {}): HttpRequest {
    const headers = options.headers ?? {};
    const query = options.query ?? {};
    const body = options.body ?? Buffer.alloc(0);
    return {
        headers: { get: (name: string) => headers[name.toLowerCase()] ?? null },
        query: { get: (name: string) => query[name] ?? null },
        params: options.params ?? {},
        arrayBuffer: async () => body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength),
        formData: async () => options.form,
    } as unknown as HttpRequest;
}

beforeEach(() => {
    mockUserGet.mockResolvedValue(staffUser());
    mockRunCreate.mockImplementation(async (doc: Partial<LgRun>) => ({
        ...doc,
        id: doc.id ?? 'run-1',
        createdAt: '2026-07-01T00:00:00Z',
        updatedAt: '2026-07-01T00:00:00Z',
    }));
    mockRunQuery.mockResolvedValue([]);
    mockAuditCreate.mockResolvedValue(undefined);
    mockDetailCreate.mockImplementation(async (doc: Record<string, unknown>) => ({ ...doc, id: 'chunk-1' }));
    mockDetailQuery.mockResolvedValue([]);
});

afterEach(() => {
    delete process.env.LG_MAX_UPLOAD_BYTES;
});

describe('POST lg/runs (F1 upload + F9 persistence + F10 auth)', () => {
    it('rejects anonymous requests with 401 and stores nothing', async () => {
        const response = await createLgRun(fakeRequest({ body: fixture('balanced-sample.csv') }));
        expect(response.status).toBe(401);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('rejects an empty body with 400', async () => {
        const response = await createLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' } }));
        expect(response.status).toBe(400);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('ingests a raw-body csv upload, stores the run and audits it', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.format).toBe('csv');
        expect(run.uploadedBy).toBe('u1');
        expect(run.inputSha256).toMatch(/^[0-9a-f]{64}$/);
        expect(run.summary.parsed).toBe(6);
        expect(run.summary.netFils).toBe(0);
        expect(run.errorCount).toBe(0);
        expect(mockAuditCreate).toHaveBeenCalledWith(
            expect.objectContaining({ actor: 'u1', action: 'lg.breakdown.ingested', entityType: 'lgRun', entityId: run.id })
        );
    });

    it('accepts a multipart form-data upload with a "file" field', async () => {
        const form = new FormData();
        form.append('file', new File([fixture('balanced-sample.xlsx')], 'balanced-sample.xlsx'));
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1', 'content-type': 'multipart/form-data; boundary=test' },
                form,
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.format).toBe('xlsx');
        expect(run.filename).toBe('balanced-sample.xlsx');
        expect(run.summary.parsed).toBe(6);
    });

    it('rejects a file with missing required headers with 422 and stores nothing', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'bad.csv' },
                body: Buffer.from('foo,bar\n1,2\n'),
            })
        );
        expect(response.status).toBe(422);
        const details = (response.jsonBody as { details: { errors: { code: string }[] } }).details;
        expect(details.errors.some((e) => e.code === 'MISSING_HEADER')).toBe(true);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('computes and stores F3 balances and F4 matching results with the run', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.asOf).toMatch(/^\d{4}-\d{2}-\d{2}$/);
        expect(run.balances!.length).toBeGreaterThan(0);
        // The balanced fixture nets to zero, so the F5 tie-out identity gives zero here too.
        expect(run.balances!.reduce((s, b) => s + b.balanceFils, 0)).toBe(0);
        expect(run.matching!.asOf).toBe(run.asOf);
        expect(run.matching!.netOutstandingFils).toBe(0);
        expect(run.outstandingCount).toBe(run.matching!.outstandingCount);
        expect(run.outstanding!.length).toBe(run.outstandingCount);
        // F5: internally-derived balances always tie out — Balanced, Difference 0.
        expect(run.reconciliation!.balanced).toBe(true);
        expect(run.reconciliation!.totalAbsDifferenceFils).toBe(0);
        expect(run.reconciliation!.byBranch.length).toBeGreaterThan(0);
        expect(run.reconciliation!.byBranch.every((b) => b.differenceFils === 0)).toBe(true);
    });

    it('honours an explicit ?asOf= review date', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv', asOf: '2020-01-01' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.asOf).toBe('2020-01-01');
        // Every posting in the fixture is after 2020, so nothing is on the balance yet —
        // and matching must operate on that same (empty) population: nothing outstanding.
        expect(run.balances).toEqual([]);
        expect(run.matching!.netOutstandingFils).toBe(0);
        expect(run.matching!.matchedFils).toBe(0);
        expect(run.outstanding).toEqual([]);
    });

    it('rejects a malformed asOf with 400', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv', asOf: 'June 2026' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(400);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('rejects an impossible calendar asOf with 400', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv', asOf: '2026-02-31' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(400);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('returns 400 (not 500) when the multipart body cannot be parsed', async () => {
        const response = await createLgRun(
            fakeRequest({ headers: { 'x-user-id': 'u1', 'content-type': 'multipart/form-data' } })
        );
        expect(response.status).toBe(400);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('caps stored outstanding items but reports the full count', async () => {
        const header = 'entity,Branch Number,gl,Post Date,Log description,ccy,Amount (BHD),Journal Number';
        const rows = Array.from({ length: 510 }, (_, i) => `BH,1,D2810085,2023-01-08,020050 DEBIT POSTING,BHD,1.000,J${i}`);
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'many-unmatched.csv' },
                body: Buffer.from([header, ...rows].join('\n')),
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.outstandingCount).toBe(510);
        expect(run.outstanding).toHaveLength(500);
        expect(run.matching!.outstandingCount).toBe(510);
    });

    it('links a re-upload of the same bytes via duplicateOf', async () => {
        mockRunQuery.mockResolvedValue([{ id: 'run-0' }]);
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(201);
        expect((response.jsonBody as LgRun).duplicateOf).toBe('run-0');
    });

    it('rejects an oversized upload with 413 and stores nothing', async () => {
        process.env.LG_MAX_UPLOAD_BYTES = '16';
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(413);
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('rejects a legacy .xls upload with 422', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'legacy.xls' },
                body: Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00]),
            })
        );
        expect(response.status).toBe(422);
        const details = (response.jsonBody as { details: { errors: { code: string }[] } }).details;
        expect(details.errors[0].code).toBe('UNSUPPORTED_FORMAT');
        expect(mockRunCreate).not.toHaveBeenCalled();
    });

    it('caps stored row errors but reports the full count', async () => {
        const header = 'entity,Branch Number,gl,Post Date,Log description,ccy,Amount (BHD),Journal Number';
        const badRow = 'BH,1,D2810085,2023-01-08,020050 DEBIT POSTING,BHD,not-a-number,J1';
        const csv = [header, ...Array.from({ length: 150 }, () => badRow)].join('\n');
        const response = await createLgRun(
            fakeRequest({ headers: { 'x-user-id': 'u1' }, query: { filename: 'bad-rows.csv' }, body: Buffer.from(csv) })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.errorCount).toBe(150);
        expect(run.errors).toHaveLength(100);
        expect(run.summary.parsed).toBe(0);
        // With zero postings there is nothing to balance or match — no fake summaries.
        expect(run.balances).toEqual([]);
        expect(run.matching).toBeUndefined();
        expect(run.outstanding).toEqual([]);
        expect(run.outstandingCount).toBe(0);
    });
});

describe('POST lg/runs — F6 detail storage (G2/G3)', () => {
    it('stores matched-set chunks for a balanced upload and reports the counts', async () => {
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'balanced-sample.csv' },
                body: fixture('balanced-sample.csv'),
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.matchedSetCount).toBeGreaterThan(0);
        expect(run.matching!.matchedSetCount).toBe(run.matchedSetCount);
        // Fully balanced ⇒ nothing outstanding ⇒ zero exceptions, and no exception chunks.
        expect(run.exceptionCount).toBe(0);
        // Chunks are written under the SAME id the run document is created with.
        expect(mockDetailCreate).toHaveBeenCalledWith(
            expect.objectContaining({ runId: run.id, kind: 'matchedSets', seq: 0 })
        );
        expect(mockDetailCreate).not.toHaveBeenCalledWith(expect.objectContaining({ kind: 'exceptions' }));
    });

    it('classifies duplicates and near-miss amounts at ingest (end-to-end G2)', async () => {
        const header = 'entity,Branch Number,gl,Account Number,Post Date,Log description,ccy,Amount (BHD),Journal Number';
        const rows = [
            'BH,1,D2810085,ACC-T,2025-03-14,020030 BGL CR POSTING,BHD,-0.555,JR', // retry twin 1
            'BH,1,D2810085,ACC-T,2025-03-14,020030 BGL CR POSTING,BHD,-0.555,JR', // retry twin 2
            'BH,1,D2810085,ACC-A,2025-01-01,020050 DEBIT POSTING,BHD,10.000,J1', // near-miss debit
            'BH,1,D2810085,ACC-B,2025-01-05,020030 BGL CR POSTING,BHD,-10.500,J2', // near-miss credit
        ];
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'classified.csv' },
                body: Buffer.from([header, ...rows].join('\n')),
            })
        );
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        expect(run.exceptionCount).toBe(4);
        expect(run.exceptionsSummary!.byReason.DUPLICATE).toBe(2);
        expect(run.exceptionsSummary!.byReason.AMOUNT_MISMATCH).toBe(2);
        expect(mockDetailCreate).toHaveBeenCalledWith(
            expect.objectContaining({ runId: run.id, kind: 'exceptions', seq: 0 })
        );
    });

    it('chunks large detail sets and honours the LG_MAX_DETAIL_ITEMS cap visibly', async () => {
        process.env.LG_MAX_DETAIL_ITEMS = '300';
        const header = 'entity,Branch Number,gl,Post Date,Log description,ccy,Amount (BHD),Journal Number';
        const rows = Array.from(
            { length: 400 },
            (_, i) => `BH,1,D2810085,2023-01-08,020050 DEBIT POSTING,BHD,${i + 1}.000,J${i}`
        );
        const response = await createLgRun(
            fakeRequest({
                headers: { 'x-user-id': 'u1' },
                query: { filename: 'many.csv' },
                body: Buffer.from([header, ...rows].join('\n')),
            })
        );
        delete process.env.LG_MAX_DETAIL_ITEMS;
        expect(response.status).toBe(201);
        const run = response.jsonBody as LgRun;
        // The true total is reported even though storage was capped at 300 (250 + 50).
        expect(run.exceptionCount).toBe(400);
        const exceptionChunks = mockDetailCreate.mock.calls
            .map(([doc]) => doc as { kind: string; seq: number; items: unknown[] })
            .filter((doc) => doc.kind === 'exceptions');
        expect(exceptionChunks.map((c) => c.items.length)).toEqual([250, 50]);
    });
});

describe('detail chunk sizing (Cosmos 2MB regression)', () => {
    it('splits chunks by serialized byte size, not just item count', async () => {
        const { buildDetailChunks } = await import('../../src/functions/lg');
        // Four ~600KB items with a 1.5MB budget → 2 per chunk at most.
        const fat = Array.from({ length: 4 }, (_, i) => ({ i, blob: 'x'.repeat(600_000) }));
        const chunks = buildDetailChunks(fat as never[], 250, 1_500_000);
        expect(chunks.map((c) => c.length)).toEqual([2, 2]);
        // A single item over the budget still ships (alone) rather than looping forever.
        const single = buildDetailChunks([{ blob: 'x'.repeat(2_000_000) }] as never[], 250, 1_500_000);
        expect(single).toHaveLength(1);
    });

    it('still honours the item-count bound', async () => {
        const { buildDetailChunks } = await import('../../src/functions/lg');
        const chunks = buildDetailChunks(Array.from({ length: 501 }, (_, i) => ({ i })) as never[], 250, 1_500_000);
        expect(chunks.map((c) => c.length)).toEqual([250, 250, 1]);
    });

    it('trims mega-set legs for storage while keeping the true counts visible', async () => {
        const { trimMatchedSet } = await import('../../src/functions/lg');
        const leg = (rowNumber: number) => ({
            postDate: '2025-01-01', direction: 'debit' as const, originalFils: 1000, matchedFils: 1000,
            journalNumber: `J${rowNumber}`, rowNumber,
        });
        const mega = {
            entity: 'BH', gl: 'D2810085', branchNumber: '1', accountNumber: 'ACC-BUSY',
            matchedFils: 120_000,
            creditLegCount: 60, debitLegCount: 120,
            creditLegs: Array.from({ length: 60 }, (_, i) => ({ ...leg(i), direction: 'credit' as const })),
            debitLegs: Array.from({ length: 120 }, (_, i) => leg(1000 + i)),
            firstCreditDate: '2025-01-01', finalDebitDate: '2025-02-01', settledDays: 31, fullyCleared: true,
        };
        const trimmed = trimMatchedSet(mega);
        expect(trimmed.creditLegs).toHaveLength(50);
        expect(trimmed.debitLegs).toHaveLength(50);
        expect(trimmed.creditLegCount).toBe(60); // truncation stays visible
        expect(trimmed.debitLegCount).toBe(120);
        expect(trimmed.matchedFils).toBe(120_000); // figures untouched
        // Small sets pass through unchanged (same reference — no copy churn).
        const small = { ...mega, creditLegs: mega.creditLegs.slice(0, 2), debitLegs: mega.debitLegs.slice(0, 2) };
        expect(trimMatchedSet(small)).toBe(small);
    });

    it('writes detail chunks BEFORE the run document so a chunk failure leaves no half-run', async () => {
        mockDetailCreate.mockRejectedValueOnce(new Error('Request size is too large'));
        await expect(
            createLgRun(
                fakeRequest({
                    headers: { 'x-user-id': 'u1' },
                    query: { filename: 'balanced-sample.csv' },
                    body: fixture('balanced-sample.csv'),
                })
            )
        ).rejects.toThrow('Request size is too large');
        // The failure happened before lgRuns.create — no orphan run document.
        expect(mockRunCreate).not.toHaveBeenCalled();
    });
});

describe('GET lg/runs/{id}/matched and /exceptions (G3)', () => {
    const chunks = (kind: string) => [
        { id: 'k0', runId: 'run-1', kind, seq: 0, items: [{ n: 1 }, { n: 2 }, { n: 3 }] },
        { id: 'k1', runId: 'run-1', kind, seq: 1, items: [{ n: 4 }, { n: 5 }] },
    ];

    it('rejects anonymous requests with 401', async () => {
        const response = await listLgRunMatched(fakeRequest({ params: { id: 'run-1' } }));
        expect(response.status).toBe(401);
    });

    it('returns 404 for an unknown run', async () => {
        mockRunGet.mockResolvedValue(undefined);
        const response = await listLgRunExceptions(
            fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'nope' } })
        );
        expect(response.status).toBe(404);
    });

    it('pages matched sets with offset/maxItems and reports total vs storedCount', async () => {
        mockRunGet.mockResolvedValue({ id: 'run-1', matchedSetCount: 9 });
        mockDetailQuery.mockResolvedValue(chunks('matchedSets'));
        const response = await listLgRunMatched(
            fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' }, query: { offset: '1', maxItems: '2' } })
        );
        expect(response.status).toBe(200);
        const body = response.jsonBody as { items: { n: number }[]; total: number; storedCount: number; offset: number };
        expect(body.items.map((i) => i.n)).toEqual([2, 3]);
        expect(body.offset).toBe(1);
        expect(body.storedCount).toBe(5);
        expect(body.total).toBe(9); // storedCount < total ⇒ the cap is visible, never silent
        expect(mockDetailQuery).toHaveBeenCalledWith(
            expect.objectContaining({
                parameters: expect.arrayContaining([{ name: '@kind', value: 'matchedSets' }]),
            })
        );
    });

    it('pages exceptions the same way', async () => {
        mockRunGet.mockResolvedValue({ id: 'run-1', exceptionCount: 5 });
        mockDetailQuery.mockResolvedValue(chunks('exceptions'));
        const response = await listLgRunExceptions(
            fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' }, query: { offset: '3' } })
        );
        expect(response.status).toBe(200);
        const body = response.jsonBody as { items: { n: number }[]; total: number };
        expect(body.items.map((i) => i.n)).toEqual([4, 5]);
        expect(body.total).toBe(5);
    });
});

describe('GET lg/runs/{id}/export (G5)', () => {
    const reconBlock = {
        entity: 'BH',
        gl: 'D2810085',
        branchNumber: '1',
        glBalanceFils: 46_500,
        outstandingNetFils: 46_500,
        outstandingCount: 2,
        oldCount: 1,
        oldFils: 125_000,
        currentCount: 1,
        currentFils: 78_500,
        differenceFils: 0,
        difference: 0,
        balanced: true,
    };

    it('rejects anonymous requests with 401', async () => {
        const response = await exportLgRun(fakeRequest({ params: { id: 'run-1' } }));
        expect(response.status).toBe(401);
    });

    it('returns 404 for an unknown run', async () => {
        mockRunGet.mockResolvedValue(undefined);
        const response = await exportLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'x' } }));
        expect(response.status).toBe(404);
    });

    it('returns 400 when the run has no reconciliation', async () => {
        mockRunGet.mockResolvedValue({ id: 'run-1' });
        const response = await exportLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' } }));
        expect(response.status).toBe(400);
    });

    it('requires ?branch= when the run covers several branches', async () => {
        mockRunGet.mockResolvedValue({
            id: 'run-1',
            reconciliation: {
                byBranch: [reconBlock, { ...reconBlock, branchNumber: '2' }],
            },
        });
        const response = await exportLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' } }));
        expect(response.status).toBe(400);
        const details = (response.jsonBody as { details: { branches: unknown[] } }).details;
        expect(details.branches).toHaveLength(2);
    });

    it('returns 404 when no block matches the requested branch', async () => {
        mockRunGet.mockResolvedValue({ id: 'run-1', reconciliation: { byBranch: [reconBlock] } });
        const response = await exportLgRun(
            fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' }, query: { branch: '99' } })
        );
        expect(response.status).toBe(404);
    });

    it('streams the two-sheet workbook with download headers', async () => {
        mockRunGet.mockResolvedValue({
            id: 'run-1',
            asOf: '2026-06-30',
            reconciliation: { asOf: '2026-06-30', byBranch: [reconBlock] },
        });
        mockDetailQuery.mockResolvedValue([]);
        const response = await exportLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' } }));
        expect(response.status).toBe(200);
        const headers = response.headers as Record<string, string>;
        expect(headers['content-type']).toContain('spreadsheetml');
        expect(headers['content-disposition']).toContain('GL-Recon_Branch-1_2026-06-30.xlsx');
        expect((response.body as Uint8Array).byteLength).toBeGreaterThan(0);
    });
});

describe('GET lg/runs (list)', () => {
    it('rejects anonymous requests with 401', async () => {
        const response = await listLgRuns(fakeRequest());
        expect(response.status).toBe(401);
    });

    it('returns the stored runs with paging passed through', async () => {
        mockRunList.mockResolvedValue({ items: [{ id: 'run-1' }], continuationToken: 'next' });
        const response = await listLgRuns(
            fakeRequest({ headers: { 'x-user-id': 'u1' }, query: { maxItems: '10', continuationToken: 'tok' } })
        );
        expect(response.status).toBe(200);
        expect(mockRunList).toHaveBeenCalledWith({ maxItems: 10, continuationToken: 'tok' });
        expect((response.jsonBody as { items: unknown[] }).items).toHaveLength(1);
    });
});

describe('GET lg/runs/{id}', () => {
    it('returns the run when it exists', async () => {
        mockRunGet.mockResolvedValue({ id: 'run-1' });
        const response = await getLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'run-1' } }));
        expect(response.status).toBe(200);
        expect(mockRunGet).toHaveBeenCalledWith('run-1');
    });

    it('returns 404 for an unknown run', async () => {
        mockRunGet.mockResolvedValue(undefined);
        const response = await getLgRun(fakeRequest({ headers: { 'x-user-id': 'u1' }, params: { id: 'nope' } }));
        expect(response.status).toBe(404);
    });
});
