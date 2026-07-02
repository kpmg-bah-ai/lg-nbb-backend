import { CrudHelper } from '../helpers/crudHelper';
import { AppNotification, AuditLogEntry, LgRun, User } from '../shared/models';

export const users = new CrudHelper<User>('users');
export const auditLogs = new CrudHelper<AuditLogEntry>('auditLogs');
export const notifications = new CrudHelper<AppNotification>('notifications');
/** LG/MCQ reconciliation runs (GOAL.md §4 F9). */
export const lgRuns = new CrudHelper<LgRun>('lgRuns');

/** Every container the app uses — drives the init-db bootstrap. */
export const CONTAINER_IDS = ['users', 'auditLogs', 'notifications', 'lgRuns'];
