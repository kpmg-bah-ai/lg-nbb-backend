/**
 * In-memory DEV store — used automatically when COSMOS_CONNECTION_STRING is not set,
 * so the app can be exercised end-to-end (sign-in, uploads, statement) without a
 * database. Data lives for the process lifetime only. NEVER a production path.
 *
 * It implements just the slice of the Cosmos `Container` surface that CrudHelper
 * uses, with a deliberately tiny SQL evaluator covering the query shapes this
 * codebase issues (SELECT * / VALUE COUNT(1) / projections, WHERE equality with
 * @params, LOWER(c.field), ORDER BY). Anything else throws loudly.
 *
 * The `users` container is seeded with three dev accounts so the placeholder
 * header auth works out of the box:
 *   admin@nbb.dev / manager@nbb.dev / staff@nbb.dev
 */

import { SqlQuerySpec } from '@azure/cosmos';

type Doc = Record<string, unknown> & { id: string };

interface Condition {
    field: string;
    lower: boolean;
    value: unknown;
}

function parseWhere(where: string | undefined, params: Map<string, unknown>): Condition[] {
    if (!where) {
        return [];
    }
    return where.split(/\s+AND\s+/i).map((clause) => {
        const m = clause
            .trim()
            .match(/^(LOWER\()?c(?:\.(\w+)|\["(\w+)"\])\)? = (@\w+|"[^"]*"|'[^']*')$/i);
        if (!m) {
            throw new Error(`Memory store cannot evaluate WHERE clause: ${clause}`);
        }
        const field = m[2] ?? m[3];
        const raw = m[4];
        const value = raw.startsWith('@') ? params.get(raw) : raw.slice(1, -1);
        return { field, lower: Boolean(m[1]), value };
    });
}

/** Evaluates the small SQL subset this app issues against the stored documents. */
export function evaluateQuery(spec: SqlQuerySpec | string, docs: Doc[]): unknown[] {
    const query = (typeof spec === 'string' ? spec : spec.query).replace(/\s+/g, ' ').trim();
    const params = new Map<string, unknown>(
        (typeof spec === 'object' ? spec.parameters ?? [] : []).map((p) => [p.name, p.value])
    );
    const m = query.match(
        /^SELECT (VALUE COUNT\(1\)|\*|c\.\w+(?: *, *c\.\w+)*) FROM c(?: WHERE (.+?))?(?: ORDER BY c\.(\w+) (ASC|DESC))?$/i
    );
    if (!m) {
        throw new Error(`Memory store cannot evaluate query: ${query}`);
    }
    const [, select, where, orderField, orderDir] = m;

    const conditions = parseWhere(where, params);
    let rows = docs.filter((doc) =>
        conditions.every(({ field, lower, value }) => {
            const cell = doc[field];
            return lower
                ? String(cell ?? '').toLowerCase() === String(value ?? '').toLowerCase()
                : cell === value;
        })
    );

    if (orderField) {
        const dir = orderDir.toUpperCase() === 'DESC' ? -1 : 1;
        rows = [...rows].sort((a, b) => {
            const av = a[orderField] as string | number | undefined;
            const bv = b[orderField] as string | number | undefined;
            return av === bv ? 0 : (av ?? '') < (bv ?? '') ? -dir : dir;
        });
    }

    if (/^VALUE COUNT/i.test(select)) {
        return [rows.length];
    }
    if (select === '*') {
        return rows;
    }
    const fields = select.split(',').map((f) => f.trim().replace(/^c\./, ''));
    return rows.map((doc) => Object.fromEntries(fields.map((f) => [f, doc[f]])));
}

function notFound(): Error & { code: number } {
    const err = new Error('Entity not found') as Error & { code: number };
    err.code = 404;
    return err;
}

/** The subset of the Cosmos Container surface CrudHelper touches. */
export class MemoryContainer {
    private readonly docs = new Map<string, Doc>();

    readonly items = {
        create: async (doc: Doc) => {
            const stored = { ...doc };
            this.docs.set(doc.id, stored);
            return { resource: { ...stored } };
        },
        upsert: async (doc: Doc) => {
            const stored = { ...doc };
            this.docs.set(doc.id, stored);
            return { resource: { ...stored } };
        },
        query: (spec: SqlQuerySpec | string) => ({
            fetchAll: async () => ({ resources: evaluateQuery(spec, [...this.docs.values()]) }),
            fetchNext: async () => ({
                resources: evaluateQuery(spec, [...this.docs.values()]),
                continuationToken: undefined,
            }),
        }),
    };

    item(id: string) {
        return {
            read: async () => {
                const doc = this.docs.get(id);
                if (!doc) {
                    throw notFound();
                }
                return { resource: { ...doc } };
            },
            replace: async (doc: Doc) => {
                if (!this.docs.has(id)) {
                    throw notFound();
                }
                const stored = { ...doc };
                this.docs.set(id, stored);
                return { resource: { ...stored } };
            },
            delete: async () => {
                if (!this.docs.delete(id)) {
                    throw notFound();
                }
            },
        };
    }

    seed(docs: Doc[]): void {
        for (const doc of docs) {
            this.docs.set(doc.id, { ...doc });
        }
    }
}

const containers = new Map<string, MemoryContainer>();
let warned = false;

/** Dev sign-in accounts seeded into the in-memory `users` container. */
export const DEV_USERS = [
    { id: 'dev-admin', displayName: 'Dev Admin', email: 'admin@nbb.dev', role: 'admin' },
    { id: 'dev-manager', displayName: 'Dev Manager', email: 'manager@nbb.dev', role: 'manager' },
    { id: 'dev-staff', displayName: 'Dev Staff', email: 'staff@nbb.dev', role: 'staff' },
] as const;

export function getMemoryContainer(containerId: string): MemoryContainer {
    if (!warned) {
        warned = true;
        console.warn(
            '[dev] COSMOS_CONNECTION_STRING is not set — using the in-memory store (data resets on restart). ' +
                'Dev sign-ins: admin@nbb.dev / manager@nbb.dev / staff@nbb.dev'
        );
    }
    let container = containers.get(containerId);
    if (!container) {
        container = new MemoryContainer();
        containers.set(containerId, container);
        if (containerId === 'users') {
            const now = new Date().toISOString();
            container.seed(
                DEV_USERS.map((user) => ({ ...user, active: true, createdBy: 'dev-seed', createdAt: now, updatedAt: now }))
            );
        }
    }
    return container;
}
