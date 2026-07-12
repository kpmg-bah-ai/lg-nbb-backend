import * as ExcelJS from 'exceljs';
import { computeBranchBalances } from '../../src/lg/balance';
import { detectExceptions } from '../../src/lg/exceptions';
import { buildStatementWorkbook, fmtBhd, fmtIsoDate, MISMATCHED_SHEET, STATEMENT_SHEET } from '../../src/lg/export';
import { matchPostings } from '../../src/lg/match';
import { reconcile } from '../../src/lg/reconcile';
import { LgRun } from '../../src/shared/models';
import { makePosting } from './helpers';

const ASOF = '2026-06-30';

/**
 * Runs the full engine over crafted postings the way ingest does, so the export
 * is asserted against the SAME stored figures the screen renders (ties exactly).
 */
function buildFixtureRun() {
    const postings = [
        // A fully cleared 1:1 pair — appears in matched sets, never on the statement.
        makePosting({ amountBhdFils: -5_000_000, postDate: '2025-01-10', accountNumber: 'ACC-CLR' }),
        makePosting({ amountBhdFils: 5_000_000, postDate: '2025-02-01', accountNumber: 'ACC-CLR' }),
        // The unmatched debit (> 1 year old at the review date) — Section A.
        makePosting({ amountBhdFils: 125_000, postDate: '2022-01-10', accountNumber: 'ACC-OLD', journalNumber: 'J-OLD' }),
        // The unmatched credit (< 1 year) — Section B.
        makePosting({ amountBhdFils: -78_500, postDate: '2026-02-10', accountNumber: 'ACC-CUR', journalNumber: 'J-CUR' }),
    ];
    const balances = computeBranchBalances(postings, ASOF);
    const match = matchPostings(postings, { asOf: ASOF });
    const reconciliation = reconcile(balances, match.outstanding, { asOf: ASOF });
    const { exceptions } = detectExceptions(match.outstanding);
    const run = {
        id: 'run-golden',
        asOf: ASOF,
        reconciliation,
        matching: match.summary,
    } as unknown as LgRun;
    return { run, recon: reconciliation.byBranch[0], exceptions };
}

async function loadWorkbook(buffer: Buffer): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
    return workbook;
}

describe('buildStatementWorkbook (G5) — golden-file layout', () => {
    it('produces the two sheets of GOAL.md §2.3 with the reference layout and tying figures', async () => {
        const { run, recon, exceptions } = buildFixtureRun();
        const workbook = await loadWorkbook(await buildStatementWorkbook(run, recon, exceptions));

        // Sheet names — statement first, mismatched second (§2.3).
        expect(workbook.worksheets.map((w) => w.name)).toEqual([STATEMENT_SHEET, MISMATCHED_SHEET]);

        const ws = workbook.getWorksheet(STATEMENT_SHEET)!;
        // Title + identity block.
        expect(ws.getCell('A1').value).toBe('GL Reconciliation — Outstanding Items Statement');
        expect(ws.getCell('A2').value).toBe('Branch: 1');
        expect(ws.getCell('A3').value).toBe('Entity: BH  ·  GL: D2810085  ·  Review Date: 30 Jun 2026');

        // Reconciliation block — labels exactly as the reference sample (incl. the sic).
        expect(ws.getCell('K2').value).toBe('GL Balance');
        expect(ws.getCell('L2').value).toBe(fmtBhd(recon.glBalanceFils));
        expect(ws.getCell('L2').value).toBe('46.500'); // 125.000 DR − 78.500 CR
        expect(ws.getCell('K3').value).toBe('Total (OLD Item + MCQ)');
        expect(ws.getCell('L3').value).toBe('203.500'); // 125.000 + 78.500 (Σ|outstanding|)
        expect(ws.getCell('K4').value).toBe('Diffrence');
        expect(ws.getCell('L4').value).toBe('0.000');
        expect(ws.getCell('K5').value).toBe('Status');
        expect(ws.getCell('L5').value).toBe('Balanced');

        // Section A — Old Items.
        expect(ws.getCell('A7').value).toBe("Old Items Outstanding – Old Manager's Checks");
        expect(ws.getCell('A8').value).toBe('No.');
        expect(ws.getCell('B8').value).toBe('Amount (BHD)');
        expect(ws.getCell('F8').value).toBe('Date of Transfer to Old Items');
        expect(ws.getCell('B9').value).toBe('125.000'); // the old unmatched debit
        expect(ws.getCell('C9').value).toBe('10 Jan 2022');
        expect(ws.getCell('H9').value).toBe('ACC-OLD');
        expect(ws.getCell('A10').value).toBe('Subtotal');
        expect(ws.getCell('B10').value).toBe(fmtBhd(recon.oldFils));

        // Section B — current outstanding.
        expect(ws.getCell('A12').value).toBe('Outstanding MCQ  (Less than 1 year)');
        expect(ws.getCell('A13').value).toBe('No.');
        expect(ws.getCell('B14').value).toBe('78.500'); // the current unmatched credit
        expect(ws.getCell('A15').value).toBe('Subtotal');
        expect(ws.getCell('B15').value).toBe(fmtBhd(recon.currentFils));
    });

    it('lists every exception on the Mismatched sheet — an unmatched debit is always there', async () => {
        const { run, recon, exceptions } = buildFixtureRun();
        const workbook = await loadWorkbook(await buildStatementWorkbook(run, recon, exceptions));
        const ws = workbook.getWorksheet(MISMATCHED_SHEET)!;

        expect(ws.getCell('A1').value).toBe('Reconciling Exceptions / Mismatched Items');
        expect(ws.getCell('A2').value).toBe('ID');
        expect(ws.getCell('I2').value).toBe('Reason / Required Action');

        // Two exceptions → rows 3 and 4; the unmatched debit is surfaced with its reason.
        const types = [ws.getCell('B3').value, ws.getCell('B4').value];
        expect(types).toContain('Unmatched Debit');
        expect(types).toContain('Unmatched Credit');
        const debitRow = ws.getCell('B3').value === 'Unmatched Debit' ? 3 : 4;
        expect(ws.getCell(`C${debitRow}`).value).toBe('DEBIT');
        expect(ws.getCell(`D${debitRow}`).value).toBe('125.000');
        expect(ws.getCell(`F${debitRow}`).value).toBe('J-OLD');
        expect(String(ws.getCell(`I${debitRow}`).value)).toContain('no matching credit');
        expect(ws.getCell('A5').value).toBeNull(); // exactly the two exceptions, nothing more
    });

    it('shows Not Balanced and the exact difference when the GL balance diverges', async () => {
        const { run, recon, exceptions } = buildFixtureRun();
        // Simulate an externally supplied balance 1 BHD off (GOAL.md §9.6).
        const skewed = { ...recon, glBalanceFils: recon.glBalanceFils + 1000, differenceFils: 1000, balanced: false };
        const workbook = await loadWorkbook(await buildStatementWorkbook(run, skewed, exceptions));
        const ws = workbook.getWorksheet(STATEMENT_SHEET)!;

        expect(ws.getCell('L4').value).toBe('1.000'); // never rounded away
        expect(ws.getCell('L5').value).toBe('Not Balanced');
    });

    it('formats deterministically without locale APIs', () => {
        expect(fmtBhd(72_501_861)).toBe('72,501.861');
        expect(fmtBhd(-1_334_921)).toBe('-1,334.921');
        expect(fmtBhd(555)).toBe('0.555');
        expect(fmtBhd(0)).toBe('0.000');
        expect(fmtIsoDate('2026-06-30')).toBe('30 Jun 2026');
        expect(fmtIsoDate('2022-09-14')).toBe('14 Sep 2022');
        expect(fmtIsoDate('not-a-date')).toBe('not-a-date');
    });
});

describe('buildStatementWorkbook — register mode (GOAL-3 R9)', () => {
    const AS_OF = '2026-02-03';

    async function buildRegisterRun() {
        const { readFileSync } = await import('node:fs');
        const { join } = await import('node:path');
        const { ingest } = await import('../../src/lg/ingest');
        const { matchRegister } = await import('../../src/lg/registerMatch');
        const { extractStatedBalance, reconcileRegister } = await import('../../src/lg/registerReconcile');
        const { classifyRegisterExceptions } = await import('../../src/lg/registerExceptions');

        const buffer = readFileSync(join(__dirname, '..', 'fixtures', 'lg', 'register-sample.xlsx'));
        const result = await ingest(buffer, { filename: 'register-sample.xlsx' });
        const match = matchRegister(result.postings, result.cheques!, { asOf: AS_OF });
        const balances = computeBranchBalances(result.postings, AS_OF);
        const stated = extractStatedBalance(result.postings);
        const reconciliation = reconcileRegister(stated.statedFils, balances, match, { asOf: AS_OF });
        const { exceptions } = classifyRegisterExceptions(match, reconciliation.byBranch[0].extractGapFils);
        const run = {
            id: 'run-register',
            mode: 'register',
            asOf: AS_OF,
            reconciliation,
            matching: match.summary,
        } as unknown as LgRun;
        return { run, recon: reconciliation.byBranch[0], exceptions, outcomes: match.outcomes };
    }

    it('fills the statement with real cheque attributes and the decomposed block', async () => {
        const { run, recon, exceptions, outcomes } = await buildRegisterRun();
        const workbook = await loadWorkbook(await buildStatementWorkbook(run, recon, exceptions, outcomes));
        expect(workbook.worksheets.map((w) => w.name)).toEqual([STATEMENT_SHEET, MISMATCHED_SHEET]);

        const ws = workbook.getWorksheet(STATEMENT_SHEET)!;
        expect(ws.getCell('A2').value).toBe('Branch: (all branches)');

        // The decomposed GL-level block — Task-1 numbers exactly.
        const block = [2, 3, 4, 5, 6, 7, 8, 9].map((r) => [ws.getCell(`K${r}`).value, ws.getCell(`L${r}`).value]);
        expect(block).toEqual([
            ['GL Balance', '2,730.000'],
            ['Total (OLD Item + MCQ)', '925.500'],
            ['Diffrence', '1,804.500'],
            ['Classified exceptions (Sheet 2)', '1,840.000'],
            ['Unexplained residual', '35.500'],
            ['Derived balance (from postings)', '2,765.500'],
            ['Ledger extract gap', '35.500'],
            ['Status', 'Not Balanced'],
        ]);

        // Section A: CHQ 1011 with its REAL cheque number and issuance date.
        expect(ws.getCell('A12').value).toBe('No.');
        expect(ws.getCell('B13').value).toBe('001'); // branch
        expect(ws.getCell('C13').value).toBe('30.000');
        expect(ws.getCell('D13').value).toBe('10 Jan 2025');
        expect(ws.getCell('E13').value).toBe('1011');
        expect(ws.getCell('A14').value).toBe('Subtotal');
        expect(ws.getCell('C14').value).toBe('30.000');

        // Section B: 1002, 1006 (branch 001) then 1009 (002); subtotal 895.500.
        const chqCells = [18, 19, 20].map((r) => ws.getCell(`E${r}`).value);
        expect(chqCells).toEqual(['1002', '1006', '1009']);
        expect(ws.getCell('A21').value).toBe('Subtotal');
        expect(ws.getCell('C21').value).toBe('895.500');

        // Ops-PAID 1004 and batch-cleared 1007/1008 never appear as statement lines.
        const allValues: string[] = [];
        ws.eachRow((row) => row.eachCell((cell) => allValues.push(String(cell.value))));
        expect(allValues).not.toContain('1004');
        expect(allValues).not.toContain('1007');

        // Sheet 2 carries the register taxonomy incl. the extract gap.
        const wsExc = workbook.getWorksheet(MISMATCHED_SHEET)!;
        const types: string[] = [];
        wsExc.eachRow((row) => types.push(String(row.getCell(2).value)));
        expect(types).toEqual(
            expect.arrayContaining(['Ledger Extract Gap', 'Ops Paid — Register Lag', 'Non-Issuance Credit'])
        );
    });

    it('?branch narrows the sections and recomputes their subtotals; the block stays GL-level', async () => {
        const { run, recon, exceptions, outcomes } = await buildRegisterRun();
        const workbook = await loadWorkbook(await buildStatementWorkbook(run, recon, exceptions, outcomes, '002'));
        const ws = workbook.getWorksheet(STATEMENT_SHEET)!;
        expect(ws.getCell('A2').value).toBe('Branch: 002');
        expect(String(ws.getCell('A4').value)).toMatch(/GL-level/);
        expect(ws.getCell('L2').value).toBe('2,730.000'); // block untouched
        // Only CHQ 1009 (branch 002) remains outstanding; Section A is empty.
        expect(ws.getCell('A13').value).toBe('Subtotal');
        expect(ws.getCell('C13').value).toBe('0.000');
        expect(ws.getCell('E17').value).toBe('1009');
        expect(ws.getCell('C18').value).toBe('45.000');
    });
});
