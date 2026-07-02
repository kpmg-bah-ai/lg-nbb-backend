import { HttpRequest } from '@azure/functions';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

jest.mock('../../src/data/repositories', () => ({
    users: { get: jest.fn() },
    auditLogs: { create: jest.fn() },
    lgRuns: { create: jest.fn(), get: jest.fn(), list: jest.fn() },
}));

import { auditLogs, lgRuns, users } from '../../src/data/repositories';
import { createLgRun, getLgRun, listLgRuns } from '../../src/functions/lg';
import { LgRun, User } from '../../src/shared/models';

const mockUserGet = users.get as jest.Mock;
const mockRunCreate = lgRuns.create as jest.Mock;
const mockRunGet = lgRuns.get as jest.Mock;
const mockRunList = lgRuns.list as jest.Mock;
const mockAuditCreate = auditLogs.create as jest.Mock;

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
        id: 'run-1',
        createdAt: '2026-07-01T00:00:00Z',
        updatedAt: '2026-07-01T00:00:00Z',
    }));
    mockAuditCreate.mockResolvedValue(undefined);
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
            expect.objectContaining({ actor: 'u1', action: 'lg.breakdown.ingested', entityType: 'lgRun', entityId: 'run-1' })
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
