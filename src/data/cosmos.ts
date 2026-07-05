import { Container, CosmosClient } from '@azure/cosmos';

import { getMemoryContainer } from './memory';

const DATABASE_ID = process.env.COSMOS_DATABASE_NAME || 'workflow-tracker';

let client: CosmosClient | undefined;
const containerCache = new Map<string, Promise<Container>>();

function getClient(): CosmosClient {
    if (!client) {
        const connectionString = process.env.COSMOS_CONNECTION_STRING;
        if (!connectionString) {
            throw new Error('COSMOS_CONNECTION_STRING app setting is not configured');
        }
        client = new CosmosClient(connectionString);
    }
    return client;
}

/**
 * Returns a container handle, creating the database/container on first use
 * (per process) so local dev and UAT work without manual provisioning.
 *
 * DEV fallback: with no COSMOS_CONNECTION_STRING configured, an in-memory store
 * (seeded with dev sign-in accounts) is used so the app runs end-to-end locally.
 */
export function getContainer(containerId: string, partitionKeyPath = '/id'): Promise<Container> {
    if (!process.env.COSMOS_CONNECTION_STRING) {
        return Promise.resolve(getMemoryContainer(containerId) as unknown as Container);
    }
    let cached = containerCache.get(containerId);
    if (cached === undefined) {
        cached = (async () => {
            const { database } = await getClient().databases.createIfNotExists({ id: DATABASE_ID });
            const { container } = await database.containers.createIfNotExists({
                id: containerId,
                partitionKey: { paths: [partitionKeyPath] },
            });
            return container;
        })();
        containerCache.set(containerId, cached);
    }
    return cached;
}
