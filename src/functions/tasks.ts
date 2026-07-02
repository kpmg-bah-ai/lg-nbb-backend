import { tasks } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

registerCrudRoutes(tasks, {
    route: 'tasks',
    requiredFields: ['name', 'clientEntity', 'assignedTo', 'dueDate'],
    filterFields: ['projectId', 'clientId', 'assignedTo', 'status'],
    // Managers define tasks (WT-7); staff update progress on their own (WT-19).
    auth: { read: 'staff', create: 'manager', update: 'staff', delete: 'manager' },
    applyDefaults: (body) => ({ status: 'not_started' as const, submittedBy: '', ...body }),
});
