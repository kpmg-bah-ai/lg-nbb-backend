import { HttpRequest } from '@azure/functions';

import { badRequest, created, error, forbidden, json, noContent, notFound, readJson, unauthorized } from '../src/helpers/json';

function fakeRequest(jsonImpl: () => Promise<unknown>): HttpRequest {
    return { json: jsonImpl } as unknown as HttpRequest;
}

describe('readJson', () => {
    it('returns the parsed body', async () => {
        const body = await readJson<{ a: number }>(fakeRequest(async () => ({ a: 1 })));
        expect(body).toEqual({ a: 1 });
    });

    it('returns undefined on invalid or missing JSON instead of throwing', async () => {
        const body = await readJson(
            fakeRequest(async () => {
                throw new SyntaxError('Unexpected end of JSON input');
            })
        );
        expect(body).toBeUndefined();
    });
});

describe('response helpers', () => {
    it('json wraps a body with a default 200', () => {
        expect(json({ ok: true })).toEqual({ status: 200, jsonBody: { ok: true } });
    });

    it('created returns 201', () => {
        expect(created({ id: 'x' }).status).toBe(201);
    });

    it('noContent returns a bare 204', () => {
        expect(noContent()).toEqual({ status: 204 });
    });

    it('error uses a consistent envelope and only includes details when given', () => {
        expect(error(500, 'boom')).toEqual({ status: 500, jsonBody: { error: 'boom' } });
        expect(error(409, 'conflict', { id: 'x' })).toEqual({
            status: 409,
            jsonBody: { error: 'conflict', details: { id: 'x' } },
        });
    });

    it('shortcut helpers use the right status codes', () => {
        expect(badRequest('x').status).toBe(400);
        expect(unauthorized().status).toBe(401);
        expect(forbidden().status).toBe(403);
        expect(notFound().status).toBe(404);
    });
});
