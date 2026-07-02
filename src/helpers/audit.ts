import { auditLogs } from '../data/repositories';
import { AuditLogEntry } from '../shared/models';

/** Appends to the audit trail (WT-33). Never throws — auditing must not break the main operation. */
export async function recordAudit(
    actor: string,
    action: string,
    entityType: string,
    entityId: string,
    details?: Record<string, unknown>
): Promise<AuditLogEntry | undefined> {
    try {
        return await auditLogs.create({ actor, action, entityType, entityId, details });
    } catch {
        return undefined;
    }
}
