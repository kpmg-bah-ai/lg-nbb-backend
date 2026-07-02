import { notifications } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

registerCrudRoutes(notifications, {
    route: 'notifications',
    requiredFields: ['kind', 'recipient'],
    filterFields: ['kind', 'recipient', 'taskId', 'status'],
    applyDefaults: (body) => ({ status: 'pending' as const, ...body }),
});
