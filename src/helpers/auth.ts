import { HttpRequest, HttpResponseInit } from '@azure/functions';

import { users } from '../data/repositories';
import { Role, User } from '../shared/models';
import { forbidden, unauthorized } from './json';

/** admin outranks manager outranks staff — a minimum-role check passes for higher roles. */
const ROLE_RANK: Record<Role, number> = { staff: 1, manager: 2, admin: 3 };

export function outranksOrEquals(role: Role, minRole: Role): boolean {
    return ROLE_RANK[role] >= ROLE_RANK[minRole];
}

/**
 * Resolves the caller from the x-user-id or x-user-email header.
 * Placeholder identity until Entra ID / App Service auth is wired in — at that
 * point this becomes a lookup against the validated token's claims.
 */
export async function getCurrentUser(request: HttpRequest): Promise<User | undefined> {
    const userId = request.headers.get('x-user-id');
    if (userId) {
        return users.get(userId);
    }
    const email = request.headers.get('x-user-email');
    if (email) {
        const matches = await users.query<User>({
            query: 'SELECT * FROM c WHERE LOWER(c.email) = @email',
            parameters: [{ name: '@email', value: email.toLowerCase() }],
        });
        return matches[0];
    }
    return undefined;
}

export interface AuthResult {
    user?: User;
    /** Set when the request must be rejected — return it as-is. */
    response?: HttpResponseInit;
}

export async function requireRole(request: HttpRequest, minRole: Role): Promise<AuthResult> {
    const user = await getCurrentUser(request);
    if (!user) {
        return { response: unauthorized('Authentication required: send an x-user-id or x-user-email header') };
    }
    if (!user.active) {
        return { response: forbidden('Account is deactivated') };
    }
    if (!outranksOrEquals(user.role, minRole)) {
        return { response: forbidden(`Requires ${minRole} role`) };
    }
    return { user };
}
