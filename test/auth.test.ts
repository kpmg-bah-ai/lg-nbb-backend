import { HttpRequest } from '@azure/functions';

jest.mock('../src/data/repositories', () => ({
    users: { get: jest.fn(), query: jest.fn() },
}));

import { users } from '../src/data/repositories';
import { getCurrentUser, outranksOrEquals, requireRole } from '../src/helpers/auth';
import { User } from '../src/shared/models';

const mockGet = users.get as jest.Mock;
const mockQuery = users.query as jest.Mock;

function fakeRequest(headers: Record<string, string>): HttpRequest {
    return { headers: { get: (name: string) => headers[name.toLowerCase()] ?? null } } as unknown as HttpRequest;
}

function makeUser(overrides: Partial<User> = {}): User {
    return {
        id: 'u1',
        displayName: 'Test User',
        email: 'test@kpmg.com',
        role: 'staff',
        active: true,
        createdAt: '2026-01-01T00:00:00Z',
        updatedAt: '2026-01-01T00:00:00Z',
        ...overrides,
    };
}

describe('outranksOrEquals (role hierarchy)', () => {
    it('admin outranks manager outranks staff', () => {
        expect(outranksOrEquals('admin', 'manager')).toBe(true);
        expect(outranksOrEquals('admin', 'staff')).toBe(true);
        expect(outranksOrEquals('manager', 'staff')).toBe(true);
    });

    it('lower roles do not pass higher checks', () => {
        expect(outranksOrEquals('staff', 'manager')).toBe(false);
        expect(outranksOrEquals('manager', 'admin')).toBe(false);
    });

    it('a role passes its own check', () => {
        expect(outranksOrEquals('staff', 'staff')).toBe(true);
        expect(outranksOrEquals('admin', 'admin')).toBe(true);
    });
});

describe('getCurrentUser', () => {
    it('resolves via x-user-id', async () => {
        const user = makeUser();
        mockGet.mockResolvedValue(user);
        expect(await getCurrentUser(fakeRequest({ 'x-user-id': 'u1' }))).toBe(user);
        expect(mockGet).toHaveBeenCalledWith('u1');
    });

    it('resolves via x-user-email, case-insensitively', async () => {
        const user = makeUser();
        mockQuery.mockResolvedValue([user]);
        expect(await getCurrentUser(fakeRequest({ 'x-user-email': 'Test@KPMG.com' }))).toBe(user);
        const spec = mockQuery.mock.calls[0][0];
        expect(spec.parameters).toEqual([{ name: '@email', value: 'test@kpmg.com' }]);
    });

    it('returns undefined without identity headers', async () => {
        expect(await getCurrentUser(fakeRequest({}))).toBeUndefined();
    });
});

describe('requireRole', () => {
    it('rejects anonymous requests with 401', async () => {
        const { user, response } = await requireRole(fakeRequest({}), 'staff');
        expect(user).toBeUndefined();
        expect(response?.status).toBe(401);
    });

    it('rejects deactivated accounts with 403', async () => {
        mockGet.mockResolvedValue(makeUser({ active: false }));
        const { response } = await requireRole(fakeRequest({ 'x-user-id': 'u1' }), 'staff');
        expect(response?.status).toBe(403);
    });

    it('rejects staff calling a manager endpoint with 403', async () => {
        mockGet.mockResolvedValue(makeUser({ role: 'staff' }));
        const { response } = await requireRole(fakeRequest({ 'x-user-id': 'u1' }), 'manager');
        expect(response?.status).toBe(403);
    });

    it('lets an admin through a manager check', async () => {
        const admin = makeUser({ role: 'admin' });
        mockGet.mockResolvedValue(admin);
        const { user, response } = await requireRole(fakeRequest({ 'x-user-id': 'u1' }), 'manager');
        expect(response).toBeUndefined();
        expect(user).toBe(admin);
    });

    it('lets a user through a check at their own level', async () => {
        mockGet.mockResolvedValue(makeUser({ role: 'manager' }));
        const { user, response } = await requireRole(fakeRequest({ 'x-user-id': 'u1' }), 'manager');
        expect(response).toBeUndefined();
        expect(user?.role).toBe('manager');
    });
});
