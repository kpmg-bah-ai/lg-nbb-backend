import { projects } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

registerCrudRoutes(projects, {
    route: 'projects',
    requiredFields: ['name', 'clientId', 'type'],
    filterFields: ['clientId', 'type', 'status', 'period'],
    applyDefaults: (body) => ({ status: 'active' as const, ...body }),
});
