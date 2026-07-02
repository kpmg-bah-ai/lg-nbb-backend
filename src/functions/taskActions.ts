import { app, HttpResponseInit } from '@azure/functions';

import { comments, notifications, projects, tasks } from '../data/repositories';
import { recordAudit } from '../helpers/audit';
import { requireRole } from '../helpers/auth';
import { badRequest, error, json, notFound, readJson } from '../helpers/json';
import { WorkflowTask } from '../shared/models';

/**
 * Sequential blocker (WT-8): a task can't start/complete while an earlier task in the
 * same project is incomplete — unless a manager allowed it to run in parallel (WT-9).
 * Returns an error response when blocked, undefined when clear.
 */
async function checkBlockers(task: WorkflowTask): Promise<HttpResponseInit | undefined> {
    if (!task.projectId) {
        return undefined;
    }
    const project = await projects.get(task.projectId);
    if (project?.status === 'closed') {
        return error(409, 'Project is closed and locked');
    }
    if (task.order === undefined || task.parallelAllowed) {
        return undefined;
    }
    const blockers = await tasks.query<WorkflowTask>({
        query: 'SELECT * FROM c WHERE c.projectId = @projectId AND c["order"] < @order AND c.status != "completed"',
        parameters: [
            { name: '@projectId', value: task.projectId },
            { name: '@order', value: task.order },
        ],
    });
    if (blockers.length > 0) {
        return error(409, 'Earlier tasks must be signed off first', {
            blockedBy: blockers.map((b) => ({ id: b.id, name: b.name, order: b.order, status: b.status })),
        });
    }
    return undefined;
}

// Staff picks up a task (WT-15/19).
app.http('tasks-start', {
    route: 'tasks/{id}/start',
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request) => {
        const { user, response } = await requireRole(request, 'staff');
        if (response) {
            return response;
        }
        const task = await tasks.get(request.params.id);
        if (!task) {
            return notFound('Task not found');
        }
        if (task.status !== 'not_started') {
            return badRequest(`Task is already ${task.status}`);
        }
        const blocked = await checkBlockers(task);
        if (blocked) {
            return blocked;
        }
        const updated = await tasks.update(task.id, {
            status: 'in_progress' as const,
            startDate: task.startDate ?? new Date().toISOString(),
        });
        await recordAudit(user.id, 'task.started', 'task', task.id, { name: task.name });
        return json(updated);
    },
});

// Single sign-off: timestamps completion, optional comment travels with it (WT-16/17).
app.http('tasks-sign-off', {
    route: 'tasks/{id}/sign-off',
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request) => {
        const { user, response } = await requireRole(request, 'staff');
        if (response) {
            return response;
        }
        const body = (await readJson<{ comment?: string }>(request)) ?? {};
        const task = await tasks.get(request.params.id);
        if (!task) {
            return notFound('Task not found');
        }
        if (task.status === 'completed') {
            return badRequest('Task is already signed off');
        }
        const blocked = await checkBlockers(task);
        if (blocked) {
            return blocked;
        }
        const updated = await tasks.update(task.id, {
            status: 'completed' as const,
            signedOffAt: new Date().toISOString(),
            signedOffBy: user.id,
        });
        if (body.comment) {
            await comments.create({
                taskId: task.id,
                projectId: task.projectId,
                author: user.id,
                text: body.comment,
                kind: 'comment' as const,
            });
        }
        await recordAudit(user.id, 'task.signed_off', 'task', task.id, { name: task.name });
        // Manager dashboard/daily-summary feed (WT-29); email dispatch comes later.
        await notifications.create({
            kind: 'sign_off' as const,
            recipient: 'manager',
            taskId: task.id,
            projectId: task.projectId,
            status: 'pending' as const,
        });
        return json(updated);
    },
});

// Manager override: allow a task to run in parallel with earlier ones (WT-9).
app.http('tasks-allow-parallel', {
    route: 'tasks/{id}/allow-parallel',
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request) => {
        const { user, response } = await requireRole(request, 'manager');
        if (response) {
            return response;
        }
        const task = await tasks.get(request.params.id);
        if (!task) {
            return notFound('Task not found');
        }
        const updated = await tasks.update(task.id, { parallelAllowed: true });
        await recordAudit(user.id, 'task.parallel_allowed', 'task', task.id, { name: task.name });
        return json(updated);
    },
});
