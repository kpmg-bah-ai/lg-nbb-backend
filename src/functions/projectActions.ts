import { app } from '@azure/functions';

import { clients, projects, tasks, templates } from '../data/repositories';
import { recordAudit } from '../helpers/audit';
import { requireRole } from '../helpers/auth';
import { addDays, resolveDueDate } from '../helpers/dates';
import { badRequest, created, error, json, notFound, readJson } from '../helpers/json';
import { WorkflowTask } from '../shared/models';

interface FromTemplateRequest {
    templateId: string;
    clientId: string;
    name?: string;
    /** Advisory: project start date (defaults to today). */
    startDate?: string;
    /** Accounting: cycle the run covers, e.g. '2026-06'. */
    period?: string;
    assignedTo: string;
    submittedBy?: string;
}

// Spin up a project (and its task chain) from an engagement template (WT-6/12).
app.http('projects-from-template', {
    route: 'projects/from-template',
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request) => {
        const { user, response } = await requireRole(request, 'manager');
        if (response) {
            return response;
        }
        const body = await readJson<FromTemplateRequest>(request);
        if (!body) {
            return badRequest('Request body must be JSON');
        }
        const missing = ['templateId', 'clientId', 'assignedTo'].filter((f) => !body[f as keyof FromTemplateRequest]);
        if (missing.length > 0) {
            return badRequest('Missing required fields', { missing });
        }
        const [template, client] = await Promise.all([templates.get(body.templateId), clients.get(body.clientId)]);
        if (!template) {
            return badRequest('Unknown templateId');
        }
        if (!client) {
            return badRequest('Unknown clientId');
        }
        const startDate = body.startDate ?? new Date().toISOString().slice(0, 10);
        const project = await projects.create({
            name: body.name ?? `${client.name} — ${template.name}${body.period ? ` (${body.period})` : ''}`,
            clientId: client.id,
            type: template.type,
            templateId: template.id,
            period: body.period,
            status: 'active' as const,
            startDate,
        });
        const createdTasks: WorkflowTask[] = [];
        for (const def of [...template.tasks].sort((a, b) => a.order - b.order)) {
            const dueDate = resolveDueDate(def, startDate, body.period);
            createdTasks.push(
                await tasks.create({
                    name: def.name,
                    projectId: project.id,
                    clientId: client.id,
                    clientEntity: client.name,
                    assignedTo: body.assignedTo,
                    submittedBy: body.submittedBy ?? '',
                    order: def.order,
                    dueDate,
                    status: 'not_started' as const,
                    escalationDate: addDays(dueDate, def.escalationOffsetDays ?? 1),
                    escalationContacts: def.escalationContacts,
                    parallelAllowed: def.allowParallelWithPrevious,
                })
            );
        }
        await recordAudit(user.id, 'project.created_from_template', 'project', project.id, {
            templateId: template.id,
            taskCount: createdTasks.length,
        });
        return created({ project, tasks: createdTasks });
    },
});

// Close & lock a project once the final deliverable is signed off (WT-10/34).
app.http('projects-close', {
    route: 'projects/{id}/close',
    methods: ['POST'],
    authLevel: 'function',
    handler: async (request) => {
        const { user, response } = await requireRole(request, 'manager');
        if (response) {
            return response;
        }
        const body = (await readJson<{ force?: boolean }>(request)) ?? {};
        const project = await projects.get(request.params.id);
        if (!project) {
            return notFound('Project not found');
        }
        if (project.status === 'closed') {
            return badRequest('Project is already closed');
        }
        const openCount = await tasks.count('c.projectId = @projectId AND c.status != "completed"', [
            { name: '@projectId', value: project.id },
        ]);
        if (openCount > 0 && !body.force) {
            return error(409, `${openCount} task(s) are not signed off yet — pass force: true to close anyway`);
        }
        const updated = await projects.update(project.id, {
            status: 'closed' as const,
            closedAt: new Date().toISOString(),
            closedBy: user.id,
        });
        await recordAudit(user.id, 'project.closed', 'project', project.id, { openTasksAtClose: openCount });
        return json(updated);
    },
});
