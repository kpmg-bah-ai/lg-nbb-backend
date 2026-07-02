import { WorkflowTask } from '../shared/models';

export type RagStatus = 'green' | 'amber' | 'red';

const MS_PER_DAY = 86_400_000;

function startOfDayUtc(date: Date): number {
    return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

/** Whole days from today until the due date; negative when past due. */
export function daysUntilDue(dueDate: string, today = new Date()): number {
    return Math.floor((startOfDayUtc(new Date(dueDate)) - startOfDayUtc(today)) / MS_PER_DAY);
}

/**
 * RAG health (WT-21/26/27): Green = on track or completed, Amber = deadline day /
 * approaching (within the 2-day nudge window), Red = past due.
 */
export function deriveRag(task: Pick<WorkflowTask, 'status' | 'dueDate'>, today = new Date()): RagStatus {
    if (task.status === 'completed') {
        return 'green';
    }
    if (!task.dueDate) {
        return 'green';
    }
    const days = daysUntilDue(task.dueDate, today);
    if (days < 0) {
        return 'red';
    }
    return days <= 2 ? 'amber' : 'green';
}
