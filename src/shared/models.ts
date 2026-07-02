import { BaseDocument } from '../helpers/crudHelper';

// ---------- Users (WT-1, WT-2) ----------

/**
 * admin   — super-manager: provisions manager accounts, system checks (WT-1).
 * manager — full portfolio visibility, creates staff accounts, BRD manager rights (WT-2/4).
 * staff   — sees only allocated clients, works own tasks (WT-3/5).
 */
export type Role = 'admin' | 'manager' | 'staff';

export interface User extends BaseDocument {
    displayName: string;
    email: string;
    role: Role;
    /** Admins provision managers; managers provision staff. */
    createdBy?: string;
    active: boolean;
}

// ---------- Clients (WT-3, WT-11) ----------

export type ClientType = 'advisory' | 'accounting';
export type RecurrenceInterval = 'monthly' | 'quarterly' | 'annually';

export interface Client extends BaseDocument {
    name: string;
    type: ClientType;
    /** Recurring (accounting) clients only. */
    recurrence?: RecurrenceInterval;
    /** Staff user ids allocated to this client — drives staff dashboard scoping. */
    allocatedUserIds: string[];
}

// ---------- Engagement templates (WT-6, WT-13) ----------

export interface TaskDefinition {
    name: string;
    /** Position in the sequential chain (blocker system, WT-8). */
    order: number;
    /** Accounting: "due by Day N of the month" — actual dates auto-calculated. */
    relativeDueDay?: number;
    /** Advisory: due date offset in days from project start. */
    durationDays?: number;
    escalationOffsetDays?: number;
    escalationContacts?: string[];
    /** Manager-approved parallel step: doesn't wait for the previous task (WT-9). */
    allowParallelWithPrevious?: boolean;
}

export interface EngagementTemplate extends BaseDocument {
    name: string;
    type: ClientType;
    tasks: TaskDefinition[];
}

// ---------- Projects / engagements (WT-10, WT-12) ----------

export type ProjectStatus = 'active' | 'closed';

export interface Project extends BaseDocument {
    name: string;
    clientId: string;
    type: ClientType;
    templateId?: string;
    /** Accounting cycle this run covers, e.g. '2026-06'. */
    period?: string;
    status: ProjectStatus;
    startDate?: string;
    endDate?: string;
    /** Closing locks the project to preserve the historical timeline. */
    closedAt?: string;
    closedBy?: string;
}

// ---------- Tasks (WT-7, WT-8, WT-16, WT-18) ----------

export type TaskStatus = 'not_started' | 'in_progress' | 'completed';

export interface WorkflowTask extends BaseDocument {
    name: string;
    projectId?: string;
    clientId?: string;
    /** Entity/client display name shown on the task receipt. */
    clientEntity: string;
    assignedTo: string;
    /** Who sent/uploaded the task. */
    submittedBy: string;
    /** Position in the project's sequential chain. */
    order?: number;
    startDate?: string;
    dueDate: string;
    status: TaskStatus;
    escalationDate?: string;
    escalationContacts?: string[];
    /** Manager override allowing this task to start in parallel (WT-9). */
    parallelAllowed?: boolean;
    /** Single sign-off: timestamped completion (WT-16). */
    signedOffAt?: string;
    signedOffBy?: string;
}

// ---------- Comments & progress updates (WT-17, WT-19) ----------

export type CommentKind = 'comment' | 'progress';

export interface TaskComment extends BaseDocument {
    taskId: string;
    projectId?: string;
    author: string;
    text: string;
    kind: CommentKind;
}

// ---------- Audit trail (WT-33) ----------

export interface AuditLogEntry extends BaseDocument {
    actor: string;
    /** e.g. 'task.signed_off', 'project.closed'. */
    action: string;
    entityType: string;
    entityId: string;
    details?: Record<string, unknown>;
}

// ---------- Notifications / escalation matrix (WT-25…WT-30) ----------

export type NotificationKind =
    | 'nudge_t_minus_2'
    | 'amber_deadline'
    | 'red_t_plus_1'
    | 'critical_t_plus_3'
    | 'sign_off'
    | 'daily_summary';

export type NotificationStatus = 'pending' | 'sent' | 'failed';

export interface NotificationRecord extends BaseDocument {
    kind: NotificationKind;
    recipient: string;
    taskId?: string;
    projectId?: string;
    status: NotificationStatus;
    sentAt?: string;
}
