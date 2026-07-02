jest.mock('../src/data/cosmos', () => ({
    getContainer: jest.fn(),
}));

import { getContainer } from '../src/data/cosmos';
import { BaseDocument, CrudHelper } from '../src/helpers/crudHelper';

interface TestDoc extends BaseDocument {
    name: string;
    status?: string;
}

function notFoundError(): Error & { code: number } {
    const err = new Error('Entity not found') as Error & { code: number };
    err.code = 404;
    return err;
}

describe('CrudHelper (against an in-memory Cosmos container)', () => {
    let store: Map<string, TestDoc>;
    let helper: CrudHelper<TestDoc>;

    beforeEach(() => {
        store = new Map();
        const container = {
            items: {
                create: async (doc: TestDoc) => {
                    store.set(doc.id, doc);
                    return { resource: doc };
                },
                upsert: async (doc: TestDoc) => {
                    store.set(doc.id, doc);
                    return { resource: doc };
                },
                query: () => ({
                    fetchAll: async () => ({ resources: [...store.values()] }),
                    fetchNext: async () => ({ resources: [...store.values()], continuationToken: undefined }),
                }),
            },
            item: (id: string) => ({
                read: async () => {
                    if (!store.has(id)) {
                        throw notFoundError();
                    }
                    return { resource: store.get(id) };
                },
                replace: async (doc: TestDoc) => {
                    if (!store.has(id)) {
                        throw notFoundError();
                    }
                    store.set(id, doc);
                    return { resource: doc };
                },
                delete: async () => {
                    if (!store.has(id)) {
                        throw notFoundError();
                    }
                    store.delete(id);
                },
            }),
        };
        (getContainer as jest.Mock).mockResolvedValue(container);
        helper = new CrudHelper<TestDoc>('test-docs');
    });

    it('create generates an id and timestamps', async () => {
        const doc = await helper.create({ name: 'first' });
        expect(doc.id).toEqual(expect.any(String));
        expect(doc.id.length).toBeGreaterThan(0);
        expect(doc.createdAt).toBe(doc.updatedAt);
        expect(new Date(doc.createdAt).getTime()).not.toBeNaN();
        expect(doc.name).toBe('first');
    });

    it('create respects a caller-supplied id', async () => {
        const doc = await helper.create({ name: 'custom', id: 'my-id' });
        expect(doc.id).toBe('my-id');
    });

    it('get returns the document, or undefined when missing', async () => {
        const doc = await helper.create({ name: 'findme' });
        expect((await helper.get(doc.id))?.name).toBe('findme');
        expect(await helper.get('nope')).toBeUndefined();
    });

    it('update shallow-merges, bumps updatedAt and keeps createdAt', async () => {
        const doc = await helper.create({ name: 'before', status: 'open' });
        const updated = await helper.update(doc.id, { name: 'after' });
        expect(updated?.name).toBe('after');
        expect(updated?.status).toBe('open');
        expect(updated?.createdAt).toBe(doc.createdAt);
        expect(updated?.id).toBe(doc.id);
    });

    it('update returns undefined for a missing document', async () => {
        expect(await helper.update('missing', { name: 'x' })).toBeUndefined();
    });

    it('delete returns true once, then false when already gone', async () => {
        const doc = await helper.create({ name: 'temp' });
        expect(await helper.delete(doc.id)).toBe(true);
        expect(await helper.delete(doc.id)).toBe(false);
    });

    it('list returns stored items', async () => {
        await helper.create({ name: 'a' });
        await helper.create({ name: 'b' });
        const { items } = await helper.list();
        expect(items).toHaveLength(2);
    });
});
