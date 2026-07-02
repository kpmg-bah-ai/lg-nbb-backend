import { app } from '@azure/functions';

import { projects, tasks } from '../data/repositories';
import { requireRole } from '../helpers/auth';
import { json } from '../helpers/json';
import { deriveRag, RagStatus } from '../helpers/rag';
import { WorkflowTask } from '../shared/models';

// Aggregated metrics for the manager dashboard (WT-22).
app.http('dashboard-metrics', {
    route: 'dashboard/metrics',
    methods: ['GET'],
    authLevel: 'function',
    handler: async (request) => {
        const { response } = await requireRole(request, 'manager');
        if (response) {
            return response;
        }
        const today = new Date().toISOString().slice(0, 10);
        const weekAgo = new Date(Date.now() - 7 * 86_400_000).toISOString();
        const [activeProjects, openTasks, overdueTasks, signOffsLast7Days] = await Promise.all([
            projects.count('c.status = "active"'),
            tasks.count('c.status != "completed"'),
            tasks.count('c.status != "completed" AND c.dueDate < @today', [{ name: '@today', value: today }]),
            tasks.count('c.signedOffAt >= @weekAgo', [{ name: '@weekAgo', value: weekAgo }]),
        ]);
        return json({ activeProjects, openTasks, overdueTasks, signOffsLast7Days });
    },
});

// Tasks with derived RAG health, filterable by project/assignee (WT-21/23/24).
app.http('dashboard-rag', {
    route: 'dashboard/rag',
    methods: ['GET'],
    authLevel: 'function',
    handler: async (request) => {
        const { response } = await requireRole(request, 'manager');
        if (response) {
            return response;
        }
        const clauses: string[] = [];
        const parameters: { name: string; value: string }[] = [];
        for (const field of ['projectId', 'clientId', 'assignedTo'] as const) {
            const value = request.query.get(field);
            if (value) {
                clauses.push(`c.${field} = @${field}`);
                parameters.push({ name: `@${field}`, value });
            }
        }
        const items = await tasks.query<WorkflowTask>({
            query: `SELECT * FROM c${clauses.length ? ` WHERE ${clauses.join(' AND ')}` : ''} ORDER BY c.dueDate ASC`,
            parameters,
        });
        const withRag = items.map((task) => ({ ...task, rag: deriveRag(task) }));
        const summary: Record<RagStatus, number> = { green: 0, amber: 0, red: 0 };
        for (const task of withRag) {
            summary[task.rag] += 1;
        }
        return json({ summary, tasks: withRag });
    },
});
