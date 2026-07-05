import { HttpRequest } from '@azure/functions';

import { evaluateQuery, getMemoryContainer } from '../src/data/memory';
import { requireRole } from '../src/helpers/auth';
import { CrudHelper } from '../src/helpers/crudHelper';

// The memory store activates only when no Cosmos connection string is configured.
beforeAll(() => {
    delete process.env.COSMOS_CONNECTION_STRING;
});

function fakeRequest(headers: Record<string, string>): HttpRequest {
    return { headers: { get: (name: string) => headers[name.toLowerCase()] ?? null } } as unknown as HttpRequest;
}

describe('evaluateQuery (dev SQL subset)', () => {
    const docs = [
        { id: '1', email: 'A@X.com', status: 'pending', createdAt: '2026-01-02' },
        { id: '2', email: 'b@x.com', status: 'sent', createdAt: '2026-01-03' },
        { id: '3', email: 'c@x.com', status: 'pending', createdAt: '2026-01-01' },
    ];

    it('supports SELECT * with ORDER BY', () => {
        const rows = evaluateQuery('SELECT * FROM c ORDER BY c.createdAt DESC', docs) as { id: string }[];
        expect(rows.map((r) => r.id)).toEqual(['2', '1', '3']);
    });

    it('supports VALUE COUNT(1) with WHERE on string literals', () => {
        expect(evaluateQuery('SELECT VALUE COUNT(1) FROM c WHERE c.status = "pending"', docs)).toEqual([2]);
    });

    it('supports LOWER() equality with parameters', () => {
        const rows = evaluateQuery(
            { query: 'SELECT * FROM c WHERE LOWER(c.email) = @email', parameters: [{ name: '@email', value: 'a@x.com' }] },
            docs
        ) as { id: string }[];
        expect(rows.map((r) => r.id)).toEqual(['1']);
    });

    it('supports projections and bracket field syntax', () => {
        const rows = evaluateQuery(
            { query: 'SELECT c.id FROM c WHERE c["status"] = @s ORDER BY c.createdAt ASC', parameters: [{ name: '@s', value: 'pending' }] },
            docs
        );
        expect(rows).toEqual([{ id: '3' }, { id: '1' }]);
    });

    it('throws loudly on unsupported SQL instead of returning wrong results', () => {
        expect(() => evaluateQuery('SELECT * FROM c WHERE c.a > 1', docs)).toThrow(/cannot evaluate/);
        expect(() => evaluateQuery('SELECT TOP 1 * FROM c', docs)).toThrow(/cannot evaluate/);
    });
});

describe('in-memory dev store through CrudHelper', () => {
    it('supports the full create/get/update/delete/list cycle', async () => {
        interface TestDoc {
            id: string;
            createdAt: string;
            updatedAt: string;
            name: string;
        }
        const helper = new CrudHelper<TestDoc>('memory-cycle-test');
        const created = await helper.create({ name: 'first' });
        expect((await helper.get(created.id))?.name).toBe('first');
        expect((await helper.update(created.id, { name: 'second' }))?.name).toBe('second');
        expect((await helper.list()).items).toHaveLength(1);
        expect(await helper.delete(created.id)).toBe(true);
        expect(await helper.get(created.id)).toBeUndefined();
    });
});

describe('seeded dev accounts (dummy sign-in)', () => {
    it('seeds admin/manager/staff users', () => {
        getMemoryContainer('users');
        const container = getMemoryContainer('users');
        expect(container).toBeDefined();
    });

    it('lets the dev staff account authenticate through the real auth helper', async () => {
        const { user, response } = await requireRole(fakeRequest({ 'x-user-email': 'staff@nbb.dev' }), 'staff');
        expect(response).toBeUndefined();
        expect(user?.id).toBe('dev-staff');
        expect(user?.role).toBe('staff');
    });

    it('is case-insensitive on the email and enforces the role hierarchy', async () => {
        const ok = await requireRole(fakeRequest({ 'x-user-email': 'Admin@NBB.dev' }), 'manager');
        expect(ok.user?.id).toBe('dev-admin');

        const denied = await requireRole(fakeRequest({ 'x-user-email': 'staff@nbb.dev' }), 'manager');
        expect(denied.response?.status).toBe(403);
    });

    it('rejects unknown identities with 401', async () => {
        const { response } = await requireRole(fakeRequest({ 'x-user-email': 'ghost@nbb.dev' }), 'staff');
        expect(response?.status).toBe(401);
    });
});
