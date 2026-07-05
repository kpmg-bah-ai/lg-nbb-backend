import { app } from '@azure/functions';
import { getContainer } from '../data/cosmos';
import { CONTAINER_IDS, notifications, users } from '../data/repositories';
import { recordAudit } from '../helpers/audit';
import { requireRole } from '../helpers/auth';
import { badRequest, created, error, json, readJson } from '../helpers/json';
import { Role } from '../shared/models';

// NOTE: these live under /api/system/* because the Functions host reserves the
// built-in /admin/* route space — an app route of admin/… fails registration with
// "route conflicts with one or more built in routes" and never comes up.

// One-shot bootstrap: creates the database and every container. Master-key only.
app.http('admin-init-db', {
    route: 'system/init-db',
    methods: ['POST'],
    authLevel: 'admin',
    handler: async () => {
        for (const id of CONTAINER_IDS) {
            await getContainer(id);
        }
        return json({
            database: process.env.COSMOS_DATABASE_NAME || 'workflow-tracker',
            containers: CONTAINER_IDS,
        });
    },
});

interface BootstrapUserRequest {
    displayName: string;
    email: string;
    role: Role;
}

// Dev-team provisioning (WT-1): creates the first admin/manager accounts with the
// master key, before any in-app user exists. Day-to-day provisioning goes through
// /api/users with the role rules.
app.http('admin-bootstrap-user', {
    route: 'system/users',
    methods: ['POST'],
    authLevel: 'admin',
    handler: async (request) => {
        const body = await readJson<BootstrapUserRequest>(request);
        if (!body) {
            return badRequest('Request body must be JSON');
        }
        const missing = ['displayName', 'email', 'role'].filter((f) => !body[f as keyof BootstrapUserRequest]);
        if (missing.length > 0) {
            return badRequest('Missing required fields', { missing });
        }
        if (!['admin', 'manager', 'staff'].includes(body.role)) {
            return badRequest('role must be admin, manager or staff');
        }
        const existing = await users.count('LOWER(c.email) = @email', [
            { name: '@email', value: body.email.toLowerCase() },
        ]);
        if (existing > 0) {
            return error(409, 'A user with this email already exists');
        }
        const user = await users.create({
            displayName: body.displayName,
            email: body.email,
            role: body.role,
            createdBy: 'dev-team',
            active: true,
        });
        await recordAudit('dev-team', 'user.bootstrapped', 'user', user.id, { role: user.role });
        return created(user);
    },
});

// System check for admin users: Cosmos reachability and per-container document counts.
app.http('admin-health', {
    route: 'system/health',
    methods: ['GET'],
    authLevel: 'function',
    handler: async (request) => {
        const { response } = await requireRole(request, 'admin');
        if (response) {
            return response;
        }
        const containers: Record<string, { ok: boolean; documents?: number; error?: string }> = {};
        let healthy = true;
        for (const id of CONTAINER_IDS) {
            try {
                const container = await getContainer(id);
                const { resources } = await container.items
                    .query<number>({ query: 'SELECT VALUE COUNT(1) FROM c' })
                    .fetchAll();
                containers[id] = { ok: true, documents: resources[0] ?? 0 };
            } catch (err) {
                healthy = false;
                containers[id] = { ok: false, error: err instanceof Error ? err.message : String(err) };
            }
        }
        const pendingNotifications = await notifications.count('c.status = "pending"').catch(() => -1);
        return json({
            status: healthy ? 'healthy' : 'degraded',
            database: process.env.COSMOS_DATABASE_NAME || 'workflow-tracker',
            containers,
            pendingNotifications,
            checkedAt: new Date().toISOString(),
        });
    },
});
