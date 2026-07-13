/**
 * GOAL-3 Task 0/12 — read-only calibration harness for the register-family workbook.
 *
 * Usage:
 *   npx tsc && node dist/scripts/calibrate-register.js "<path-to-workbook.xlsx>" [headers-out.json]
 *
 * Prints, per worksheet: name, row count, and the first non-blank row (the header).
 * For the register sheet (67-column ETL extract) it also prints the status-code
 * distribution — codes only, never payee/beneficiary data. When a second argument
 * is given, the header rows are written there as JSON (headers only, no data rows),
 * for committing as backend/test/fixtures/lg/register-headers.json.
 *
 * The workbook itself is confidential and must never be committed (GOAL-3 header note).
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { filsToBhd, RawRow } from '../src/shared/models';
import { computeBranchBalances } from '../src/lg/balance';
import { ingest, sheetsFromXlsx } from '../src/lg/ingest';
import { matchRegister, statementOutstanding } from '../src/lg/registerMatch';
import { classifyRegisterExceptions } from '../src/lg/registerExceptions';
import { extractStatedBalance, reconcileRegister } from '../src/lg/registerReconcile';

function isBlankRow(row: RawRow): boolean {
    return row.every((c) => c === null || c === undefined || String(c).trim() === '');
}

function cellText(c: unknown): string {
    if (c === null || c === undefined) {
        return '';
    }
    if (c instanceof Date) {
        return c.toISOString().slice(0, 10);
    }
    return String(c);
}

async function main(): Promise<void> {
    const path = process.argv[2];
    if (!path) {
        console.error('Usage: node dist/scripts/calibrate-register.js <workbook.xlsx> [headers-out.json]');
        process.exit(1);
    }
    const headersOut = process.argv[3];
    const buffer = readFileSync(path);
    const sheets = await sheetsFromXlsx(buffer);

    const snapshot: Record<string, string[]> = {};
    for (const sheet of sheets) {
        const nonBlank = sheet.rows.filter((r) => !isBlankRow(r));
        const headerRow = nonBlank[0] ?? [];
        const headers = headerRow.map(cellText);
        snapshot[sheet.name ?? '(unnamed)'] = headers;

        console.log(`\n=== Sheet "${sheet.name}" — ${sheet.rows.length} rows (${nonBlank.length} non-blank) ===`);
        console.log(`Header (${headers.length} cols):`);
        headers.forEach((h, i) => console.log(`  [${i}] ${JSON.stringify(h)}`));

        // Register sheet heuristic: many c*-prefixed columns → print status distribution
        // (codes only) and the fill counts of the ops working columns.
        const cCols = headers.filter((h) => /^c\d+/i.test(h.trim())).length;
        if (cCols > 20) {
            const statusIdx = headers.findIndex((h) => /status/i.test(h) || /^c12\b/i.test(h.trim()));
            if (statusIdx >= 0) {
                const dist = new Map<string, number>();
                for (const row of nonBlank.slice(1)) {
                    const v = cellText(row[statusIdx]).trim();
                    dist.set(v, (dist.get(v) ?? 0) + 1);
                }
                console.log(`Status distribution (col [${statusIdx}] ${JSON.stringify(headers[statusIdx])}):`);
                [...dist.entries()].sort().forEach(([k, n]) => console.log(`  ${JSON.stringify(k)}: ${n}`));
            }
            for (const name of ['ops remark', 'ops journal', 'ops date', 'credit ref', 'debit ref']) {
                const idx = headers.findIndex((h) => h.trim().toLowerCase() === name);
                if (idx >= 0) {
                    const filled = nonBlank.slice(1).filter((r) => cellText(r[idx]).trim() !== '').length;
                    console.log(`Column ${JSON.stringify(headers[idx])} filled: ${filled}`);
                }
            }
        }
    }

    if (headersOut) {
        writeFileSync(headersOut, JSON.stringify(snapshot, null, 2) + '\n', 'utf8');
        console.log(`\nHeader snapshot written to ${headersOut} (headers only — no data rows).`);
    }

    // ---- GOAL-3 Task 12: full-pipeline calibration against §4.7 targets ----
    const bhd = (fils: number) => filsToBhd(fils).toFixed(3);
    console.log('\n=== GOAL-3 §4.7 calibration (actual vs target) ===');
    const result = await ingest(buffer, { filename: path });
    console.log(`mode: ${result.mode} (target register)`);
    console.log(`postings parsed: ${result.postings.length}; cheques: ${result.cheques?.length ?? 0}`);
    const parseErrors = new Map<string, number>();
    for (const e of result.errors) {
        parseErrors.set(e.code, (parseErrors.get(e.code) ?? 0) + 1);
    }
    console.log(`parse errors: ${JSON.stringify([...parseErrors.entries()])}`);
    if (result.mode !== 'register') {
        return;
    }

    const match = matchRegister(result.postings, result.cheques ?? []);
    const { outcomes } = match;
    const issuanceMatched = outcomes.filter((o) => o.issuanceRowNumber !== undefined).length;
    const paymentKey = outcomes.filter((o) => o.paymentRowNumbers && o.matchedVia !== 'BATCH_REF').length;
    const paymentBatch = outcomes.filter((o) => o.matchedVia === 'BATCH_REF').length;
    const preWindowPaid = outcomes.filter((o) => o.state === 'PRE_WINDOW' && o.paymentRowNumbers).length;
    console.log(`asOf: ${match.summary.asOf}`);
    console.log(`cheques with issuance credit in window: ${issuanceMatched} (target 9,921)`);
    console.log(`cheques with payment debit via KEY/variant: ${paymentKey} (target 8,178)`);
    console.log(`cheques paid via batch Ref.# resolution: ${paymentBatch} (automation win — no workbook flag)`);
    console.log(`issued pre-window but paid in-window: ${preWindowPaid} (target 223)`);

    const byState = new Map<string, number>();
    for (const o of outcomes) {
        byState.set(o.state, (byState.get(o.state) ?? 0) + 1);
    }
    console.log(`outcomes by state: ${JSON.stringify([...byState.entries()].sort())}`);
    console.log(`ops-PAID remarks parsed: ${outcomes.filter((o) => o.opsPaid).length} (README measured 1,788 PAID variants)`);

    const statement = statementOutstanding(match);
    const statementFils = statement.reduce((s, o) => s + o.outstandingFils, 0);
    console.log(`statement outstanding: ${statement.length} cheques, Σ ${bhd(statementFils)} (target 213 / 2,246,157.310)`);

    const balances = computeBranchBalances(result.postings, match.summary.asOf);
    const stated = extractStatedBalance(result.postings);
    if (stated.error) {
        console.log(`stated-balance error: ${stated.error.code} — ${stated.error.message}`);
    }
    const reconciliation = reconcileRegister(stated.statedFils, balances, match);
    const block = reconciliation.byBranch[0];
    console.log(`stated balance: ${block.statedBalanceFils !== undefined ? bhd(-block.statedBalanceFils) : '(none)'} credit (target 2,233,751.100)`);
    console.log(`derived balance: ${bhd(-(block.derivedBalanceFils ?? 0))} credit (target 2,267,752.966)`);
    console.log(`extract gap: ${bhd(Math.abs(block.extractGapFils ?? 0))} (target 34,001.866)`);
    // D_biz = S − O_stmt = (S − O_all) + (O_all − O_stmt) = difference + classified.
    const businessDifference = block.differenceFils + (block.classifiedFils ?? 0);
    console.log(`difference GL vs sections: ${bhd(Math.abs(businessDifference))} (target 12,406.210)`);
    console.log(`residual (S − O_all): ${bhd(Math.abs(block.residualFils ?? 0))}; classified: ${bhd(Math.abs(block.classifiedFils ?? 0))}`);
    console.log(`sections: old ${block.oldCount} / ${bhd(block.oldFils)} · current ${block.currentCount} / ${bhd(block.currentFils)}`);

    const gap = block.extractGapFils;
    const exceptions = classifyRegisterExceptions(match, gap);
    console.log(`exceptions: ${JSON.stringify(Object.entries(exceptions.summary.byReason).sort())}`);

    // ---- Diagnostics for target misses ----
    // (1) Cheques the workbook counts outstanding but the batch pass cleared:
    //     PAID_VIA_BATCH ∧ never register-matched ∧ not ops-PAID.
    const batchBeyondWorkbook = outcomes.filter(
        (o) =>
            o.state === 'PAID_VIA_BATCH' &&
            o.matchedPostDate === undefined &&
            o.matchedJournal === undefined &&
            !o.opsPaid
    );
    console.log(
        `batch-cleared cheques the workbook still counts outstanding: ${batchBeyondWorkbook.length}, ` +
            `Σ ${bhd(batchBeyondWorkbook.reduce((s, o) => s + o.amountFils, 0))} ` +
            `(expected to explain the 213-vs-actual statement delta)`
    );
    for (const o of batchBeyondWorkbook.slice(0, 5)) {
        console.log(`  cheque #${o.chequeNumber} ${bhd(o.amountFils)} issued ${o.issuedDate} jrnl ${o.issuedJournal}`);
    }

    // (2) Credits with an issuance journal+amount hit but no date match (the missing 9,921st).
    const issuanceJA = new Map<string, string[]>();
    for (const c of result.cheques ?? []) {
        if (c.issuedJournal) {
            const k = JSON.stringify([c.issuedJournal, c.amountFils]);
            issuanceJA.set(k, [...(issuanceJA.get(k) ?? []), `${c.issuedPostDate} (chq ${c.chequeNumber})`]);
        }
    }
    const matchedCreditRows = new Set(outcomes.filter((o) => o.issuanceRowNumber).map((o) => o.issuanceRowNumber));
    let nearMisses = 0;
    for (const p of result.postings) {
        if (p.direction !== 'credit' || matchedCreditRows.has(p.rowNumber)) {
            continue;
        }
        const k = JSON.stringify([p.journalNumber, Math.abs(p.amountBhdFils)]);
        const registerDates = issuanceJA.get(k);
        if (registerDates && nearMisses < 5) {
            nearMisses++;
            console.log(
                `  near-miss credit: row ${p.rowNumber} txn ${p.transactionDate} post ${p.postDate} ` +
                    `jrnl ${p.journalNumber} ${bhd(Math.abs(p.amountBhdFils))} vs register ${registerDates.join('; ')}`
            );
        }
    }

    // (3) Unparseable Ops Date shapes (values only — no names).
    const badOps = new Set<string>();
    const opsIdx = (snapshot['Sheet1'] ?? []).findIndex((h) => h.trim().toLowerCase() === 'ops date');
    if (opsIdx >= 0) {
        const registerSheet = sheets.find((s) => (s.name ?? '') === 'Sheet1');
        for (const row of registerSheet?.rows.slice(1) ?? []) {
            const v = row[opsIdx];
            if (typeof v === 'string' && v.trim() !== '' && !/^\d{2}\/\d{2}\/\d{4}$/.test(v.trim())) {
                badOps.add(v.trim());
                if (badOps.size >= 10) {
                    break;
                }
            }
        }
        console.log(`ops-date shapes not matching dd/mm/yyyy (sample): ${JSON.stringify([...badOps])}`);
    }
}

main().catch((err) => {
    console.error(err);
    process.exit(1);
});
