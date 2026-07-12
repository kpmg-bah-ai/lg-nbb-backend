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
import { RawRow } from '../src/shared/models';
import { sheetsFromXlsx } from '../src/lg/ingest';

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
}

main().catch((err) => {
    console.error(err);
    process.exit(1);
});
