import { users } from '../data/repositories';
import { outranksOrEquals } from '../helpers/auth';
import { registerCrudRoutes } from '../helpers/crudRoutes';
import { badRequest, forbidden } from '../helpers/json';

registerCrudRoutes(users, {
    route: 'users',
    requiredFields: ['displayName', 'email', 'role'],
    filterFields: ['role', 'email'],
    auth: { read: 'manager', create: 'manager', update: 'manager', delete: 'manager' },
    applyDefaults: (body) => ({ active: true, ...body }),
    // Provisioning chain (WT-1/2): admins create managers (and anything else);
    // managers create staff only. Managers can't touch admin/manager accounts.
    authorize: async (actor, action, ctx) => {
        if (action === 'create') {
            const targetRole = ctx.body?.role;
            if (targetRole && !['admin', 'manager', 'staff'].includes(targetRole)) {
                return badRequest('role must be admin, manager or staff');
            }
            if (actor.role === 'admin') {
                return undefined;
            }
            return targetRole === 'staff' ? undefined : forbidden('Managers can only create staff accounts');
        }
        if ((action === 'update' || action === 'delete') && ctx.id) {
            const target = await users.get(ctx.id);
            if (target && actor.role !== 'admin' && outranksOrEquals(target.role, 'manager')) {
                return forbidden('Only admins can modify manager or admin accounts');
            }
            const newRole = ctx.body?.role;
            if (newRole && newRole !== 'staff' && actor.role !== 'admin') {
                return forbidden('Only admins can grant manager or admin roles');
            }
        }
        return undefined;
    },
});
