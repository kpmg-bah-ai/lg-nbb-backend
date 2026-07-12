/**
 * GOAL-3 Task 1 — generates test/fixtures/lg/register-sample.xlsx, the crafted
 * register-family fixture whose arithmetic is hand-verified in GOAL-3.md §7:
 *
 *   Σ credits 3,675.500 · Σ debits 910.000 · derived credit balance 2,765.500
 *   stated EoD 2,730.000 · extract gap 35.500
 *   statement outstanding: CHQ 1011 (old, 30.000) + 1002/1006/1009 (current, 895.500)
 *
 * Headers are the REAL strings frozen by Task 0 (test/fixtures/lg/register-headers.json),
 * including the quirks: duplicate "Transaction Date" (cols 0 & 20), blank header col 8,
 * Debit sheet `count` col 5 + `reconciled` col 32 after six blank headers.
 *
 * Usage: npx tsc && node dist/scripts/make-register-sample.js
 */

import { join } from 'node:path';
import * as ExcelJS from 'exceljs';

const OUT = join(__dirname, '..', '..', 'test', 'fixtures', 'lg', 'register-sample.xlsx');

const LEDGER_COMMON = [
    'Transaction Date', // 0
    'Posting Date', // 1
    'Nostro/BGL Account', // 2
    'Journal Number', // 3
    'New Ref', // 4
    '', // 5 — per-sheet: 'MGR CHQ Number' (Credit) / 'count' (Debit)
    'Account Name', // 6
    'Transaction Description', // 7
    '', // 8 — blank header in the real file
    'Cheque Number', // 9
    'Transaction Credit Amount', // 10
    'Transaction Debit Amount', // 11
    'Transaction Type', // 12
    'Teller', // 13
    'Branch', // 14
    'End Date EoD Balance', // 15
    'Previous EoD Balance', // 16
    'Detailed Description', // 17
    'Outlet', // 18
    'Authorization', // 19
    'Transaction Date', // 20 — duplicate header in the real file
    'Transaction Time', // 21
    'RRN', // 22
    'Card Number', // 23
    'Balance', // 24
    'Sequence Number', // 25
];

const REGISTER_HEADERS = [
    'etl_date', 'Credit Ref', 'Debit Ref', 'c0_bank_code', 'c1_inst_no', 'c2_instr_type',
    'c3_inv_cat', 'c4_cheque_type', 'c5_account_no', 'c6_chq_no', 'c7_dup_no', 'c8_dup_flag',
    'c9_amount', 'c10_payee_name', 'c11_memb_no', 'c12_status', 'Ops Remark', 'Ops Journal',
    'Ops Date', 'c13_issued_branch', 'c14_issued_teller', 'c15_issued_trml', 'c16_issued_date',
    'c17_issued_post_dt', 'c18_issued_tran_cd', 'c19_issued_jrnl_no', 'c20_issued_batch',
    'c21_matchd_branch', 'c22_matchd_teller', 'c23_matchd_trml', 'c24_matchd_date',
    'c25_matchd_post_dt', 'c26_matchd_tran_cd', 'c27_matchd_jrnl_no', 'c28_matchd_batch',
    'c29_stpd_rsn_cd', 'c30_stpd_comment', 'c31_cncld_date', 'c32_password', 'c33_prefix',
    'c34_chk_digit', 'c35_chk_dt_flag', 'c36_issue_time', 'c37_purchaser_name',
    'c38_beneficiary_name', 'c39_beneficiary_adrs', 'c40_receive_branch', 'c41_beneficiary_father',
    'c42_beneficiary_bplace', 'c43_beneficiary_id_no', 'c44_beneficiary_tel_no', 'c45_forged_flag',
    'c46_chqm_pur_name_2', 'c47_chqm_pur_add_2', 'c48_currency', 'c49_match_ref_no',
    'c50_drawee_branch', 'c51_issued_supvisor', 'c52_matched_supvisor', 'c53_commision',
    'c54_exch_rate', 'c55_service_amt', 'c56_matched_time', 'c57_applicant_tel_no',
    'c58_ben_id_type', 'c59_app_id_type', 'c60_app_id_no',
];

const d = (iso: string) => new Date(`${iso}T00:00:00Z`);
const STATED_EOD = 2730.0;
const PREV_EOD = 2600.0;
const SENTINEL_DATE = d('1901-01-01');

interface LedgerRow {
    txn: string;
    journal: number;
    amount: number; // positive; sign handled per sheet column
    branch: string;
    type: string;
    chq?: number;
    desc?: string;
    detailed?: string;
}

// Credit sheet — cheque issuances + non-issuance credits. Σ = 3,675.500.
const CREDITS: LedgerRow[] = [
    { txn: '2025-01-01', journal: 999999999, amount: 1000.0, branch: '001', type: 'BGL CR POSTING', desc: 'OPENING TAKE ON' },
    { txn: '2025-03-10', journal: 5001, amount: 100.0, branch: '001', type: 'DD FROM DEP A/C', chq: 1001 },
    { txn: '2025-05-05', journal: 5002, amount: 250.5, branch: '001', type: 'DD FROM DEP A/C', chq: 1002 },
    { txn: '2025-07-01', journal: 5004, amount: 500.0, branch: '001', type: 'DD FROM DEP A/C', chq: 1004 },
    { txn: '2025-09-09', journal: 5005, amount: 600.0, branch: '001', type: 'DD FROM GL A/C', chq: 1005 },
    { txn: '2025-09-09', journal: 5005, amount: 600.0, branch: '001', type: 'DD FROM GL A/C', chq: 1006 },
    { txn: '2025-11-11', journal: 5007, amount: 120.0, branch: '002', type: 'DD FROM DEP A/C', chq: 1007 },
    { txn: '2025-11-11', journal: 5008, amount: 80.0, branch: '002', type: 'DD FROM DEP A/C', chq: 1008 },
    { txn: '2026-01-20', journal: 5009, amount: 45.0, branch: '002', type: 'DD FROM DEP A/C', chq: 1009 },
    { txn: '2025-12-20', journal: 5010, amount: 300.0, branch: '002', type: 'DD FROM DEP A/C', chq: 1010 },
    { txn: '2025-01-10', journal: 5011, amount: 30.0, branch: '001', type: 'DD FROM DEP A/C', chq: 1011 },
    { txn: '2025-08-15', journal: 9100, amount: 50.0, branch: '002', type: 'COR CHQ Redeep', desc: 'COR CHQ REDEEP' },
];

// Debit sheet — encashments (held NEGATIVE in the file). Σ |amounts| = 910.000.
const DEBITS: LedgerRow[] = [
    { txn: '2025-04-01', journal: 6001, amount: 100.0, branch: '001', type: 'NPB CHQ DEPOSIT', desc: 'CHQ PRESENTED' },
    { txn: '2025-10-10', journal: 6005, amount: 600.0, branch: '001', type: 'NPB CHQ DEPOSIT', desc: 'CHQ PRESENTED' },
    { txn: '2025-12-01', journal: 8001, amount: 200.0, branch: '002', type: 'DEBIT POSTING', detailed: 'DEBIT POSTING-20-Ref.# 5007,Ref.# 5008' },
    { txn: '2026-01-30', journal: 9999, amount: 10.0, branch: '002', type: 'DEBIT POSTING', detailed: 'DEBIT POSTING-20-MISC' },
];

interface RegisterRow {
    chq: number;
    amount: number;
    status: string;
    issuedDate: Date | number; // one raw Excel serial exercises F2 coercion
    issuedPost: Date;
    issuedJournal: number;
    matchedPost?: Date; // undefined ⇒ sentinel 1901-01-01 + journal 0
    matchedJournal?: number;
    branch: string;
    payee?: string;
    currency?: string;
    opsRemark?: string;
    opsJournal?: number;
    opsDate?: string; // dd/mm/yyyy TEXT — the real file's landmine
}

const REGISTER: RegisterRow[] = [
    { chq: 1001, amount: 100.0, status: '02', issuedDate: d('2025-03-10'), issuedPost: d('2025-03-10'), issuedJournal: 5001, matchedPost: d('2025-04-01'), matchedJournal: 6001, branch: '001', payee: 'PAYEE ONE' },
    { chq: 1002, amount: 250.5, status: '01', issuedDate: d('2025-05-05'), issuedPost: d('2025-05-05'), issuedJournal: 5002, branch: '001', currency: '48' },
    // 1003: legacy status-05, issued pre-window; c16 as RAW Excel serial (45458 = 2024-06-15).
    { chq: 1003, amount: 75.25, status: '05', issuedDate: 45458, issuedPost: d('2024-06-15'), issuedJournal: 4003, branch: '001' },
    { chq: 1004, amount: 500.0, status: '01', issuedDate: d('2025-07-01'), issuedPost: d('2025-07-01'), issuedJournal: 5004, branch: '001', opsRemark: 'PAID', opsJournal: 7001, opsDate: '15/08/2025' },
    // 1005/1006: legitimate key collision (same issued post date + journal + amount).
    { chq: 1005, amount: 600.0, status: '02', issuedDate: d('2025-09-09'), issuedPost: d('2025-09-09'), issuedJournal: 5005, matchedPost: d('2025-10-10'), matchedJournal: 6005, branch: '001', payee: 'JOHN DOE  �' },
    { chq: 1006, amount: 600.0, status: '01', issuedDate: d('2025-09-09'), issuedPost: d('2025-09-09'), issuedJournal: 5005, branch: '001' },
    { chq: 1007, amount: 120.0, status: '01', issuedDate: d('2025-11-11'), issuedPost: d('2025-11-11'), issuedJournal: 5007, branch: '002' },
    { chq: 1008, amount: 80.0, status: '01', issuedDate: d('2025-11-11'), issuedPost: d('2025-11-11'), issuedJournal: 5008, branch: '002' },
    { chq: 1009, amount: 45.0, status: '01', issuedDate: d('2026-01-20'), issuedPost: d('2026-01-20'), issuedJournal: 5009, branch: '002' },
    { chq: 1010, amount: 300.0, status: '02', issuedDate: d('2025-12-20'), issuedPost: d('2025-12-20'), issuedJournal: 5010, matchedPost: d('2026-01-25'), matchedJournal: 6010, branch: '002' },
    { chq: 1011, amount: 30.0, status: '01', issuedDate: d('2025-01-10'), issuedPost: d('2025-01-10'), issuedJournal: 5011, branch: '001' },
];

function ledgerRow(r: LedgerRow, credit: boolean, seq: number, width: number): (string | number | Date | null)[] {
    const row: (string | number | Date | null)[] = new Array(width).fill(null);
    row[0] = d(r.txn);
    row[1] = d(r.txn);
    row[2] = '99801000';
    row[3] = r.journal;
    row[6] = 'SAMPLE CUSTOMER NAME';
    row[7] = r.desc ?? r.type;
    row[9] = r.chq ?? null;
    row[10] = credit ? r.amount : null;
    row[11] = credit ? null : -r.amount;
    row[12] = r.type;
    row[13] = 'T123';
    row[14] = r.branch;
    row[15] = STATED_EOD;
    row[16] = PREV_EOD;
    row[17] = r.detailed ?? '';
    row[20] = d(r.txn);
    row[21] = '06:00:00';
    row[25] = seq;
    return row;
}

function registerRow(r: RegisterRow): (string | number | Date | null)[] {
    const row: (string | number | Date | null)[] = new Array(REGISTER_HEADERS.length).fill(null);
    row[0] = d('2026-02-03'); // etl_date
    // Credit Ref / Debit Ref deliberately left blank: the engine must recompute
    // keys from primitives, never trust the workbook's concatenated strings.
    row[5] = 'PO'; // c2_instr_type
    row[9] = r.chq; // c6_chq_no
    row[12] = r.amount; // c9_amount
    row[13] = r.payee ?? ''; // c10_payee_name
    row[15] = r.status; // c12_status
    row[16] = r.opsRemark ?? ''; // Ops Remark
    row[17] = r.opsJournal ?? null; // Ops Journal
    row[18] = r.opsDate ?? ''; // Ops Date (dd/mm/yyyy TEXT)
    row[19] = r.branch; // c13_issued_branch
    row[22] = r.issuedDate; // c16_issued_date
    row[23] = r.issuedPost; // c17_issued_post_dt
    row[25] = r.issuedJournal; // c19_issued_jrnl_no
    row[31] = r.matchedPost ?? SENTINEL_DATE; // c25_matchd_post_dt
    row[33] = r.matchedJournal ?? 0; // c27_matchd_jrnl_no
    row[43] = 'SAMPLE PURCHASER'; // c37_purchaser_name
    row[44] = 'SAMPLE BENEFICIARY'; // c38_beneficiary_name
    row[54] = r.currency ?? ''; // c48_currency
    return row;
}

/** A visually blank padding row (single space so the cell physically exists). */
function blankRow(width: number): (string | null)[] {
    const row: (string | null)[] = new Array(width).fill(null);
    row[0] = ' ';
    return row;
}

async function main(): Promise<void> {
    const workbook = new ExcelJS.Workbook();

    const credit = workbook.addWorksheet('Credit');
    const creditHeaders = [...LEDGER_COMMON];
    creditHeaders[5] = 'MGR CHQ Number';
    credit.addRow(creditHeaders);
    CREDITS.forEach((r, i) => credit.addRow(ledgerRow(r, true, i + 1, creditHeaders.length)));

    const register = workbook.addWorksheet('Sheet1');
    register.addRow(REGISTER_HEADERS);
    REGISTER.forEach((r) => register.addRow(registerRow(r)));
    for (let i = 0; i < 3; i++) {
        register.addRow(blankRow(REGISTER_HEADERS.length));
    }

    const debit = workbook.addWorksheet('Debit');
    const debitHeaders = [...LEDGER_COMMON, '', '', '', '', '', '', 'reconciled'];
    debitHeaders[5] = 'count';
    debit.addRow(debitHeaders);
    DEBITS.forEach((r, i) => debit.addRow(ledgerRow(r, false, i + 1, debitHeaders.length)));
    for (let i = 0; i < 3; i++) {
        debit.addRow(blankRow(debitHeaders.length));
    }

    await workbook.xlsx.writeFile(OUT);
    console.log(`Fixture written to ${OUT}`);
    console.log(`Credits: ${CREDITS.length} rows, Σ ${CREDITS.reduce((s, r) => s + r.amount, 0).toFixed(3)}`);
    console.log(`Debits: ${DEBITS.length} rows, Σ ${DEBITS.reduce((s, r) => s + r.amount, 0).toFixed(3)}`);
    console.log(`Register: ${REGISTER.length} cheques`);
}

main().catch((err) => {
    console.error(err);
    process.exit(1);
});
