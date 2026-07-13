// TEMPORARY verification harness for review claim #1 — DELETE AFTER RUN.
import * as ExcelJS from 'exceljs';
import { readFileSync } from 'node:fs';

import { ingestFiles } from '../src/lg/ingest';
import { matchRegister } from '../src/lg/registerMatch';

const CSV = String.raw`C:\Users\FATEMA~1\AppData\Local\Temp\claude\c--Users-fatemaahmed-OneDrive---KPMG-Documents-coding-nbb\fcee38f8-c556-413b-8e99-7be4e75382fb\scratchpad\register.sanitized.csv`;

const STATEMENT_HEADER = [
    'Transaction Date', 'Posting Date', 'Nostro/BGL Account', 'Journal Number',
    'Transaction Credit Amount', 'Transaction Debit Amount', 'Branch', 'End Date EoD Balance', 'Detailed Description',
];
const REGISTER_HEADER = [
    'c2_instr_type', 'c6_chq_no', 'c9_amount', 'c10_payee_name', 'c16_issued_date',
    'c17_issued_post_dt', 'c19_issued_jrnl_no', 'c25_matchd_post_dt', 'c27_matchd_jrnl_no',
];

async function workbookBuffer(sheets: { name: string; rows: unknown[][] }[]): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();
    for (const sheet of sheets) {
        workbook.addWorksheet(sheet.name).addRows(sheet.rows);
    }
    return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function ledger(): Promise<Buffer> {
    return workbookBuffer([
        { name: 'Credit', rows: [STATEMENT_HEADER, ['2025-03-10', '2025-03-11', '99801000', 'J5001', 100, null, '001', 2730, '']] },
        { name: 'Debit', rows: [STATEMENT_HEADER, ['2025-04-01', null, '99801000', 'J6001', null, -100, '001', 2730, '']] },
    ]);
}

describe('claim1: sanitized register CSV vs raw register xlsx', () => {
    it('SANITIZE ON: real sanitized csv pooled with ledger', async () => {
        const result = await ingestFiles([
            { buffer: await ledger(), filename: 'ledger.sanitized.xlsx' },
            { buffer: readFileSync(CSV), filename: 'register.sanitized.csv' },
        ]);
        console.log('mode:', result.mode);
        console.log('errors:', JSON.stringify(result.errors, null, 2));
        console.log('cheques:', JSON.stringify(result.cheques, null, 2));
        const match = matchRegister(result.postings, result.cheques ?? []);
        console.log('matchedSets:', match.matchedSets.length);
        console.log('outcomes:', JSON.stringify(match.outcomes.map((o) => ({
            chequeNumber: o.chequeNumber, state: o.state, ageBucket: o.ageBucket,
            issuedDate: o.issuedDate, matchedPostDate: o.matchedPostDate,
        })), null, 2));
        console.log('outstanding:', JSON.stringify(match.outstanding.map((o) => ({
            direction: o.direction, reason: o.reason, outstanding: o.outstanding, ageBucket: o.ageBucket,
        })), null, 2));
    });

    it('SANITIZE OFF: same register as raw xlsx with date cells', async () => {
        const register = await workbookBuffer([
            {
                name: 'Sheet1',
                rows: [REGISTER_HEADER, ['PO', 1001, 100, 'ALI HASSAN',
                    new Date(Date.UTC(2025, 2, 10)), new Date(Date.UTC(2025, 2, 10)), 'J5001',
                    new Date(Date.UTC(2025, 3, 1)), 'J6001']],
            },
        ]);
        const result = await ingestFiles([
            { buffer: await ledger(), filename: 'ledger.xlsx' },
            { buffer: register, filename: 'register.xlsx' },
        ]);
        console.log('mode:', result.mode);
        console.log('errors:', JSON.stringify(result.errors));
        const match = matchRegister(result.postings, result.cheques ?? []);
        console.log('matchedSets:', match.matchedSets.length);
        console.log('outcomes:', JSON.stringify(match.outcomes.map((o) => ({
            chequeNumber: o.chequeNumber, state: o.state,
            issuedDate: o.issuedDate, matchedPostDate: o.matchedPostDate,
        }))));
        console.log('outstanding:', JSON.stringify(match.outstanding.map((o) => ({
            direction: o.direction, reason: o.reason, outstanding: o.outstanding,
        }))));
    });
});
