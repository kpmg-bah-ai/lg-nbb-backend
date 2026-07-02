import { HttpRequest, HttpResponseInit } from '@azure/functions';

/** Parses a JSON request body; returns undefined when the body is missing or invalid. */
export async function readJson<T>(request: HttpRequest): Promise<T | undefined> {
    try {
        return (await request.json()) as T;
    } catch {
        return undefined;
    }
}

export function json<T>(body: T, status = 200): HttpResponseInit {
    return { status, jsonBody: body };
}

export function created<T>(body: T): HttpResponseInit {
    return json(body, 201);
}

export function noContent(): HttpResponseInit {
    return { status: 204 };
}

/** Consistent error envelope: { error: message, details? } */
export function error(status: number, message: string, details?: unknown): HttpResponseInit {
    return { status, jsonBody: details === undefined ? { error: message } : { error: message, details } };
}

export function badRequest(message: string, details?: unknown): HttpResponseInit {
    return error(400, message, details);
}

export function notFound(message = 'Not found'): HttpResponseInit {
    return error(404, message);
}

export function unauthorized(message = 'Authentication required'): HttpResponseInit {
    return error(401, message);
}

export function forbidden(message = 'Insufficient permissions'): HttpResponseInit {
    return error(403, message);
}
