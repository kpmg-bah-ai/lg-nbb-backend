import { CrudHelper } from '../helpers/crudHelper';
import { AppNotification, AuditLogEntry, LgRun, LgRunDetailChunk, User } from '../shared/models';

export const users = new CrudHelper<User>('users');
export const auditLogs = new CrudHelper<AuditLogEntry>('auditLogs');
export const notifications = new CrudHelper<AppNotification>('notifications');
/** GL reconciliation runs (GOAL.md §4 F9). */
export const lgRuns = new CrudHelper<LgRun>('lgRuns');
/** Per-run matched sets + exceptions, chunked (GOAL-2 G3 — too large for the run doc). */
export const lgRunDetails = new CrudHelper<LgRunDetailChunk>('lgRunDetails');

/** Every container the app uses — drives the init-db bootstrap. */
export const CONTAINER_IDS = ['users', 'auditLogs', 'notifications', 'lgRuns', 'lgRunDetails'];
