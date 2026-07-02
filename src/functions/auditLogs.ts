import { auditLogs } from '../data/repositories';
import { registerCrudRoutes } from '../helpers/crudRoutes';

// Append-only: the audit trail (WT-33) must not be editable or deletable.
registerCrudRoutes(auditLogs, {
    route: 'audit-logs',
    requiredFields: ['actor', 'action', 'entityType', 'entityId'],
    filterFields: ['actor', 'action', 'entityType', 'entityId'],
    auth: { read: 'manager', create: 'manager' },
    immutable: true,
});
