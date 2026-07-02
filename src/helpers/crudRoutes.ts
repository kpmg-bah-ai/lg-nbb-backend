import { SqlQuerySpec } from '@azure/cosmos';
import { app, HttpRequest, HttpResponseInit } from '@azure/functions';

import { Role, User } from '../shared/models';
import { requireRole } from './auth';
import { BaseDocument, CrudHelper } from './crudHelper';
import { badRequest, created, json, noContent, notFound, readJson } from './json';

export type CrudAction = 'list' | 'get' | 'create' | 'update' | 'delete';

export interface CrudRouteOptions<T extends BaseDocument> {
    /** URL segment, e.g. 'clients' → /api/clients and /api/clients/{id}. */
    route: string;
    requiredFields: (keyof T & string)[];
    /** Query-string params that filter the list endpoint by equality, e.g. ?projectId=…&status=… */
    filterFields?: (keyof T & string)[];
    /** Fills in defaults on create, after required-field validation. */
    applyDefaults?: (body: Partial<T>) => Partial<T>;
    /** Append-only containers (audit logs): no update/delete endpoints. */
    immutable?: boolean;
    /** Minimum role per action. Defaults: read = staff, create/update/delete = manager. */
    auth?: Partial<Record<'read' | 'create' | 'update' | 'delete', Role>>;
    /** Entity-specific rule evaluated after the role check. Return a response to deny. */
    authorize?: (
        actor: User,
        action: CrudAction,
        ctx: { id?: string; body?: Partial<T> }
    ) => Promise<HttpResponseInit | undefined> | HttpResponseInit | undefined;
}

function buildFilterQuery(request: HttpRequest, filterFields: string[]): SqlQuerySpec | undefined {
    const clauses: string[] = [];
    const parameters: { name: string; value: string }[] = [];
    for (const field of filterFields) {
        const value = request.query.get(field);
        if (value) {
            clauses.push(`c["${field}"] = @${field}`);
            parameters.push({ name: `@${field}`, value });
        }
    }
    if (clauses.length === 0) {
        return undefined;
    }
    return { query: `SELECT * FROM c WHERE ${clauses.join(' AND ')} ORDER BY c.createdAt DESC`, parameters };
}

/** Registers the standard REST endpoints for one container, with role-based access control. */
export function registerCrudRoutes<T extends BaseDocument>(helper: CrudHelper<T>, options: CrudRouteOptions<T>): void {
    const { route, requiredFields, filterFields = [], applyDefaults, immutable, authorize } = options;
    const minRole: Record<CrudAction, Role> = {
        list: options.auth?.read ?? 'staff',
        get: options.auth?.read ?? 'staff',
        create: options.auth?.create ?? 'manager',
        update: options.auth?.update ?? 'manager',
        delete: options.auth?.delete ?? 'manager',
    };

    async function guard(
        request: HttpRequest,
        action: CrudAction,
        ctx: { id?: string; body?: Partial<T> } = {}
    ): Promise<HttpResponseInit | undefined> {
        const { user, response } = await requireRole(request, minRole[action]);
        if (response) {
            return response;
        }
        if (authorize) {
            return authorize(user, action, ctx);
        }
        return undefined;
    }

    app.http(`${route}-list`, {
        route,
        methods: ['GET'],
        authLevel: 'function',
        handler: async (request) => {
            const denied = await guard(request, 'list');
            if (denied) {
                return denied;
            }
            const result = await helper.list({
                query: buildFilterQuery(request, filterFields),
                maxItems: Number(request.query.get('maxItems')) || undefined,
                continuationToken: request.query.get('continuationToken') ?? undefined,
            });
            return json(result);
        },
    });

    app.http(`${route}-create`, {
        route,
        methods: ['POST'],
        authLevel: 'function',
        handler: async (request) => {
            const body = await readJson<Partial<T>>(request);
            if (!body) {
                return badRequest('Request body must be JSON');
            }
            const denied = await guard(request, 'create', { body });
            if (denied) {
                return denied;
            }
            const missing = requiredFields.filter((field) => body[field] === undefined || body[field] === '');
            if (missing.length > 0) {
                return badRequest('Missing required fields', { missing });
            }
            const payload = applyDefaults ? applyDefaults(body) : body;
            const doc = await helper.create(payload as Omit<T, keyof BaseDocument> & { id?: string });
            return created(doc);
        },
    });

    app.http(`${route}-get`, {
        route: `${route}/{id}`,
        methods: ['GET'],
        authLevel: 'function',
        handler: async (request) => {
            const denied = await guard(request, 'get', { id: request.params.id });
            if (denied) {
                return denied;
            }
            const doc = await helper.get(request.params.id);
            return doc ? json(doc) : notFound(`${route} item not found`);
        },
    });

    if (immutable) {
        return;
    }

    app.http(`${route}-update`, {
        route: `${route}/{id}`,
        methods: ['PATCH', 'PUT'],
        authLevel: 'function',
        handler: async (request) => {
            const changes = await readJson<Partial<T>>(request);
            if (!changes) {
                return badRequest('Request body must be JSON');
            }
            const denied = await guard(request, 'update', { id: request.params.id, body: changes });
            if (denied) {
                return denied;
            }
            delete changes.id;
            const updated = await helper.update(request.params.id, changes);
            return updated ? json(updated) : notFound(`${route} item not found`);
        },
    });

    app.http(`${route}-delete`, {
        route: `${route}/{id}`,
        methods: ['DELETE'],
        authLevel: 'function',
        handler: async (request) => {
            const denied = await guard(request, 'delete', { id: request.params.id });
            if (denied) {
                return denied;
            }
            const deleted = await helper.delete(request.params.id);
            return deleted ? noContent() : notFound(`${route} item not found`);
        },
    });
}
