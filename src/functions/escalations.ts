import { app, InvocationContext, Timer } from '@azure/functions';

import { notifications, tasks } from '../data/repositories';
import { daysUntilDue } from '../helpers/rag';
import { NotificationKind, WorkflowTask } from '../shared/models';

/**
 * Escalation matrix (WT-25/27/28): T-2 nudge to the assignee, T+1 red and T+3
 * critical escalations to the escalation contacts. Amber on deadline day is
 * derived live by the dashboard (WT-26), not stored. Records are written as
 * 'pending' — the Outlook/email dispatcher picks them up later.
 */
export async function escalationScan(timer: Timer, context: InvocationContext): Promise<void> {
    const open = await tasks.query<WorkflowTask>({
        query: 'SELECT * FROM c WHERE c.status != "completed" AND IS_DEFINED(c.dueDate)',
    });
    let createdCount = 0;
    for (const task of open) {
        const days = daysUntilDue(task.dueDate);
        let kind: NotificationKind | undefined;
        if (days === 2) {
            kind = 'nudge_t_minus_2';
        } else if (days === -1) {
            kind = 'red_t_plus_1';
        } else if (days === -3) {
            kind = 'critical_t_plus_3';
        }
        if (!kind) {
            continue;
        }
        const recipients = kind === 'nudge_t_minus_2' ? [task.assignedTo] : task.escalationContacts?.length ? task.escalationContacts : ['manager'];
        for (const recipient of recipients) {
            const already = await notifications.count('c.taskId = @taskId AND c.kind = @kind AND c.recipient = @recipient', [
                { name: '@taskId', value: task.id },
                { name: '@kind', value: kind },
                { name: '@recipient', value: recipient },
            ]);
            if (already > 0) {
                continue;
            }
            await notifications.create({
                kind,
                recipient,
                taskId: task.id,
                projectId: task.projectId,
                status: 'pending' as const,
            });
            createdCount++;
        }
    }
    context.log(`Escalation scan: ${open.length} open tasks checked, ${createdCount} notification(s) queued`);
}

// Daily at 06:00 UTC.
app.timer('escalation-scan', {
    schedule: '0 0 6 * * *',
    handler: escalationScan,
});
