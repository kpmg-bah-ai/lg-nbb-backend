import { templates } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

registerCrudRoutes(templates, {
    route: 'templates',
    requiredFields: ['name', 'type'],
    filterFields: ['type'],
    // Engagement templates are manager-only territory (WT-6).
    auth: { read: 'manager', create: 'manager', update: 'manager', delete: 'manager' },
    applyDefaults: (body) => ({ tasks: [], ...body }),
});
