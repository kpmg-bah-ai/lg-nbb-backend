/**
 * GOAL-3 R1 — sheet-role detection and run-mode resolution (GOAL-3 §4.1). Pure.
 *
 * A worksheet plays exactly one role: `breakdown` (the original 24-col schema),
 * `ledgerStatement` (Nostro/BGL Credit/Debit extract), `register` (the c0…c60
 * cheque-register ETL extract) or `unknown`. The three schemas are disjoint by
 * construction — the breakdown needs `Branch Number`+`Amount (BHD)`, the
 * statement needs `Branch`+`Transaction Credit/Debit Amount`, the register
 * needs the c*-columns — so the first match is the only possible match.
 *
 * A workbook's mode follows from its sheet roles: register + ledger statement
 * ⇒ `register`; breakdown sheets ⇒ `breakdown`; both families together or half
 * a register input ⇒ a loud ParseError, never a guess.
 */

import { LgRunMode, ParseError, RawRow, SheetRole } from '../shared/models';
import { findHeaderRow } from './parse';
import { findRegisterHeaderRow } from './registerParse';
import { findStatementHeaderRow } from './statementParse';

/** Classifies one worksheet by scanning its leading rows for a known header. */
export function detectSheetRole(rows: RawRow[]): SheetRole {
    if (findHeaderRow(rows) >= 0) {
        return 'breakdown';
    }
    if (findRegisterHeaderRow(rows) >= 0) {
        return 'register';
    }
    if (findStatementHeaderRow(rows) >= 0) {
        return 'ledgerStatement';
    }
    return 'unknown';
}

export interface ModeResolution {
    mode?: LgRunMode;
    error?: ParseError;
}

/** Resolves the run mode from the workbook's sheet roles (GOAL-3 §4.1). */
export function resolveMode(roles: SheetRole[]): ModeResolution {
    const hasBreakdown = roles.includes('breakdown');
    const hasRegister = roles.includes('register');
    const hasStatement = roles.includes('ledgerStatement');

    if (hasBreakdown && (hasRegister || hasStatement)) {
        return {
            error: {
                code: 'MIXED_MODE',
                message:
                    'The upload mixes both input families (transaction-breakdown sheets alongside ' +
                    'register/ledger-statement sheets) — upload each family separately',
            },
        };
    }
    if (hasRegister && hasStatement) {
        return { mode: 'register' };
    }
    if (hasRegister) {
        return {
            error: {
                code: 'INCOMPLETE_REGISTER_INPUT',
                message:
                    'The upload carries a cheque-register sheet but no Credit/Debit ledger-statement ' +
                    'sheets — the GL extract is required to reconcile against the register ' +
                    '(it can be a separate file in the same upload)',
            },
        };
    }
    if (hasStatement) {
        return {
            error: {
                code: 'INCOMPLETE_REGISTER_INPUT',
                message:
                    'The upload carries ledger-statement sheet(s) but no cheque-register sheet — ' +
                    'the register is required to identify cheque issuances and payments ' +
                    '(it can be a separate file in the same upload)',
            },
        };
    }
    if (hasBreakdown) {
        return { mode: 'breakdown' };
    }
    return {};
}
