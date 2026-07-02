import { CrudHelper } from '../helpers/crudHelper';
import {
    AuditLogEntry,
    Client,
    EngagementTemplate,
    NotificationRecord,
    Project,
    TaskComment,
    User,
    WorkflowTask,
} from '../shared/models';

export const users = new CrudHelper<User>('users');
export const clients = new CrudHelper<Client>('clients');
export const templates = new CrudHelper<EngagementTemplate>('templates');
export const projects = new CrudHelper<Project>('projects');
export const tasks = new CrudHelper<WorkflowTask>('tasks');
export const comments = new CrudHelper<TaskComment>('comments');
export const auditLogs = new CrudHelper<AuditLogEntry>('audit-logs');
export const notifications = new CrudHelper<NotificationRecord>('notifications');

/** Every container the app uses — drives the init-db bootstrap. */
export const CONTAINER_IDS = ['users', 'clients', 'templates', 'projects', 'tasks', 'comments', 'audit-logs', 'notifications'];
