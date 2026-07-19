/**
 * GOAL-5 — the per-sheet balance reference.
 *
 * "Show the balance of all the amounts per sheet." Every worksheet the ingest
 * pipeline touched is totalled on its own: how many rows it contributed, its Σ
 * credits, Σ debits and signed net, and — for register-family ledger sheets that
 * state one — the End Date EoD balance printed on the sheet's final-day rows.
 * The result is stored on the run (`sheetBalances`) so a reviewer sees, at a
 * glance and without re-opening the source workbook, what each sheet adds up to
 * and whether the sheets tie to one another.
 *
 * Pure function — the same IngestResult always yields the same, deterministically
 * ordered, balances (GOAL-2 §6). Money is integer fils throughout; credits/debits
 * are magnitudes, `netFils` is signed (debit +, credit −), matching the engine's
 * sign convention.
 */

import { ParseError, ParsedPosting, RegisterCheque, SheetBalance, SheetRoleContribution, filsToBhd } from '../shared/models';
import { fmtBhd } from './export';

/** Just the fields of an IngestResult this module needs (avoids a circular import). */
export interface SheetBalanceInput {
    mode: 'breakdown' | 'register' | 'statement';
    postings: ParsedPosting[];
    cheques?: RegisterCheque[];
    errors: ParseError[];
}

/** Fallback label for a posting/cheque with no recorded sheet name (single-sheet csv). */
const UNNAMED_SHEET = '(sheet)';

interface Acc {
    sheet: string;
    role: SheetRoleContribution;
    parsedRows: number;
    creditCount: number;
    debitCount: number;
    creditFils: number;
    debitFils: number;
    chequeCount: number;
    chequeFils: number;
    /** Stated EoD keyed by the sheet's latest post date, to resolve the final-day figure. */
    latestStatedDate?: string;
    latestStatedFils?: number;
    inconsistentStated: boolean;
}

/** Deterministic display order: ledger sheets, then register, then skipped; each group by name. */
const ROLE_ORDER: Record<SheetRoleContribution, number> = { ledger: 0, breakdown: 0, register: 1, skipped: 2 };

export function computeSheetBalances(input: SheetBalanceInput): SheetBalance[] {
    const byName = new Map<string, Acc>();
    // register AND statement sheets are ledger sheets (the statedEodFils capture
    // below runs for role 'ledger'); only breakdown stays 'breakdown'.
    const ledgerRole: SheetRoleContribution = input.mode === 'breakdown' ? 'breakdown' : 'ledger';

    const accFor = (sheet: string, role: SheetRoleContribution): Acc => {
        let acc = byName.get(sheet);
        if (!acc) {
            acc = {
                sheet,
                role,
                parsedRows: 0,
                creditCount: 0,
                debitCount: 0,
                creditFils: 0,
                debitFils: 0,
                chequeCount: 0,
                chequeFils: 0,
                inconsistentStated: false,
            };
            byName.set(sheet, acc);
        }
        return acc;
    };

    for (const p of input.postings) {
        const acc = accFor(p.sheet ?? UNNAMED_SHEET, ledgerRole);
        acc.parsedRows++;
        const magnitude = Math.abs(p.amountBhdFils);
        if (p.direction === 'debit') {
            acc.debitCount++;
            acc.debitFils += magnitude;
        } else {
            acc.creditCount++;
            acc.creditFils += magnitude;
        }
        // Stated EoD comes from the final-day rows; track the latest date and flag disagreement.
        if (p.statedEodFils !== undefined) {
            if (acc.latestStatedDate === undefined || p.postDate > acc.latestStatedDate) {
                acc.latestStatedDate = p.postDate;
                acc.latestStatedFils = p.statedEodFils;
                acc.inconsistentStated = false;
            } else if (p.postDate === acc.latestStatedDate && p.statedEodFils !== acc.latestStatedFils) {
                acc.inconsistentStated = true;
            }
        }
    }

    for (const c of input.cheques ?? []) {
        const acc = accFor(c.sheet ?? UNNAMED_SHEET, 'register');
        // A sheet holding cheques is a register sheet even if it also (never) had postings.
        acc.role = 'register';
        acc.chequeCount++;
        acc.chequeFils += Math.abs(c.amountFils);
    }

    for (const e of input.errors) {
        if (e.code === 'SHEET_SKIPPED' && e.sheet && !byName.has(e.sheet)) {
            accFor(e.sheet, 'skipped');
        }
    }

    const balances: SheetBalance[] = [...byName.values()].map((acc) => {
        const netFils = acc.debitFils - acc.creditFils;
        const statedEodFils =
            acc.role === 'ledger' && !acc.inconsistentStated && acc.latestStatedFils !== undefined
                ? // The file states a credit balance as a positive magnitude — negate into
                  // engine-signed fils (matches registerReconcile.extractStatedBalance).
                  -acc.latestStatedFils
                : undefined;
        return {
            sheet: acc.sheet,
            role: acc.role,
            parsedRows: acc.parsedRows,
            creditCount: acc.creditCount,
            debitCount: acc.debitCount,
            creditFils: acc.creditFils,
            debitFils: acc.debitFils,
            netFils,
            ...(acc.role === 'register' ? { chequeCount: acc.chequeCount, chequeFils: acc.chequeFils } : {}),
            ...(statedEodFils !== undefined ? { statedEodFils } : {}),
            basis: describeSheet(acc, netFils, statedEodFils),
        };
    });

    return balances.sort(
        (a, b) => ROLE_ORDER[a.role] - ROLE_ORDER[b.role] || a.sheet.localeCompare(b.sheet)
    );
}

/** One human sentence explaining what this sheet is and how its balance was derived. */
function describeSheet(acc: Acc, netFils: number, statedEodFils?: number): string {
    if (acc.role === 'register') {
        return (
            `Cheque register: ${acc.chequeCount.toLocaleString()} instrument row(s) totalling ` +
            `BHD ${fmtBhd(acc.chequeFils)}. Not a GL posting sheet — this is the instrument ` +
            `population the ledger is reconciled against, so it carries no debit/credit balance.`
        );
    }
    if (acc.role === 'skipped') {
        return 'Worksheet skipped at ingest — it matched no known schema, so it contributes no amounts.';
    }
    const dir = netFils === 0 ? 'nets to zero' : netFils > 0 ? 'net debit' : 'net credit';
    let text =
        `Ledger extract: ${acc.parsedRows.toLocaleString()} posting(s) — ` +
        `Σ debits BHD ${fmtBhd(acc.debitFils)} − Σ credits BHD ${fmtBhd(acc.creditFils)} = ` +
        `${dir} BHD ${fmtBhd(Math.abs(netFils))} (signed ${filsToBhd(netFils).toFixed(3)}).`;
    if (statedEodFils !== undefined) {
        text +=
            ` The sheet states an End Date EoD balance of BHD ${fmtBhd(Math.abs(statedEodFils))} ` +
            `on its final-day rows (the ledger's own closing figure).`;
    } else if (acc.inconsistentStated) {
        text += ' Final-day rows disagreed on the stated EoD balance, so none is shown.';
    }
    return text;
}
