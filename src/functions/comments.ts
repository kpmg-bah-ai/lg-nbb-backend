import { comments } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

registerCrudRoutes(comments, {
    route: 'comments',
    requiredFields: ['taskId', 'author', 'text'],
    filterFields: ['taskId', 'projectId', 'author', 'kind'],
    // Staff comment on and update progress for their own tasks (WT-17/19).
    auth: { read: 'staff', create: 'staff', update: 'staff', delete: 'manager' },
    applyDefaults: (body) => ({ kind: 'comment' as const, ...body }),
});
