import { clients } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

registerCrudRoutes(clients, {
    route: 'clients',
    requiredFields: ['name', 'type'],
    filterFields: ['type', 'recurrence'],
    applyDefaults: (body) => ({ allocatedUserIds: [], ...body }),
});
