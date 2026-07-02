import { Container, SqlParameter, SqlQuerySpec } from '@azure/cosmos';
import { randomUUID } from 'node:crypto';

import { getContainer } from '../data/cosmos';

export interface BaseDocument {
    id: string;
    createdAt: string;
    updatedAt: string;
}

export interface ListResult<T> {
    items: T[];
    continuationToken?: string;
}

export interface ListOptions {
    /** Cosmos SQL query; defaults to SELECT * ordered by createdAt desc. */
    query?: SqlQuerySpec;
    maxItems?: number;
    continuationToken?: string;
}

function isNotFound(error: unknown): boolean {
    return (error as { code?: number })?.code === 404;
}

/**
 * Generic Cosmos DB CRUD repository. One instance per container, e.g.:
 *
 *   const tasks = new CrudHelper<WorkflowTask>('tasks');
 *   const created = await tasks.create({ name: 'VAT return', ... });
 *
 * Documents are partitioned by id unless a partitionKeyPath is supplied.
 */
export class CrudHelper<T extends BaseDocument> {
    constructor(
        private readonly containerId: string,
        private readonly partitionKeyPath = '/id',
        private readonly getPartitionKey: (item: T) => string = (item) => item.id
    ) {}

    private container(): Promise<Container> {
        return getContainer(this.containerId, this.partitionKeyPath);
    }

    async create(data: Omit<T, keyof BaseDocument> & { id?: string }): Promise<T> {
        const now = new Date().toISOString();
        const doc = { ...data, id: data.id ?? randomUUID(), createdAt: now, updatedAt: now } as unknown as T;
        const container = await this.container();
        const { resource } = await container.items.create<T>(doc);
        return resource as T;
    }

    async get(id: string, partitionKey?: string): Promise<T | undefined> {
        const container = await this.container();
        try {
            const { resource } = await container.item(id, partitionKey ?? id).read<T>();
            return resource ?? undefined;
        } catch (error) {
            if (isNotFound(error)) {
                return undefined;
            }
            throw error;
        }
    }

    async list(options: ListOptions = {}): Promise<ListResult<T>> {
        const query = options.query ?? { query: 'SELECT * FROM c ORDER BY c.createdAt DESC' };
        const container = await this.container();
        const iterator = container.items.query<T>(query, {
            maxItemCount: options.maxItems ?? 50,
            continuationToken: options.continuationToken,
        });
        const response = await iterator.fetchNext();
        return { items: response.resources, continuationToken: response.continuationToken };
    }

    /** Runs an arbitrary SQL query and returns all results (no paging). */
    async query<R = T>(spec: SqlQuerySpec): Promise<R[]> {
        const container = await this.container();
        const { resources } = await container.items.query<R>(spec).fetchAll();
        return resources;
    }

    /** Counts documents matching an optional WHERE clause, e.g. count("c.status = @s", [{ name: '@s', value: 'active' }]). */
    async count(where?: string, parameters?: SqlParameter[]): Promise<number> {
        const results = await this.query<number>({
            query: `SELECT VALUE COUNT(1) FROM c${where ? ` WHERE ${where}` : ''}`,
            parameters,
        });
        return results[0] ?? 0;
    }

    /** Shallow-merges changes into the stored document. Returns undefined if it doesn't exist. */
    async update(id: string, changes: Partial<Omit<T, keyof BaseDocument>>, partitionKey?: string): Promise<T | undefined> {
        const existing = await this.get(id, partitionKey);
        if (!existing) {
            return undefined;
        }
        const updated = { ...existing, ...changes, id, updatedAt: new Date().toISOString() } as T;
        const container = await this.container();
        const { resource } = await container
            .item(id, partitionKey ?? this.getPartitionKey(existing))
            .replace<T>(updated, { accessCondition: { type: 'IfMatch', condition: (existing as { _etag?: string })._etag } });
        return resource as T;
    }

    async upsert(doc: Omit<T, 'createdAt' | 'updatedAt'> & Partial<BaseDocument>): Promise<T> {
        const now = new Date().toISOString();
        const full = { createdAt: now, ...doc, updatedAt: now } as unknown as T;
        const container = await this.container();
        const { resource } = await container.items.upsert<T>(full);
        return resource as T;
    }

    /** Returns false if the document didn't exist. */
    async delete(id: string, partitionKey?: string): Promise<boolean> {
        const container = await this.container();
        try {
            await container.item(id, partitionKey ?? id).delete();
            return true;
        } catch (error) {
            if (isNotFound(error)) {
                return false;
            }
            throw error;
        }
    }
}
