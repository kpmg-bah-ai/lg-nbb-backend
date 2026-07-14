/**
 * LG reconciliation — the authoritative Excel export (GOAL.md §4 F8, GOAL-2 G5).
 *
 * Builds the two-sheet workbook of GOAL.md §2.3 server-side from the stored run:
 *   Sheet 1 — the GL's outstanding-items statement in its catalog layout
 *           (GL_CATALOG[glCodeOf(run)].statementLabels, GOAL-7 §6): the cheque
 *           two-section layout for register-mode GLs, or the suspense-fragment
 *           sections for breakdown-mode GLs. Both carry the header block with
 *           GL Balance / Total / Diffrence (sic) / Status and Section A (Old) +
 *           Section B (< 1 year) with subtotals.
 *   Sheet 2 "Mismatched"   — the full F6 exception list with reason codes.
 *
 * Figures come from the run's stored reconciliation block and detail items — the
 * same numbers the screen shows — so the export ties exactly (GOAL.md §5). Line
 * items are the run's exceptions (every outstanding item is one exception), split
 * by age bucket. Columns the ledger does not carry (CHQ #, Remitting Bank,
 * reviewer Comment/Status — GOAL.md §9.3) are placeholders, as on screen.
 *
 * Formatting is hand-rolled (no locale APIs) so the same input always produces
 * the same cells regardless of runtime ICU (GOAL-2 §6 determinism).
 */

import * as ExcelJS from 'exceljs';
import {
    BranchReconciliation,
    ChequeOutcome,
    ExplainedFigure,
    GL_CATALOG,
    glCodeOf,
    LgException,
    LgExceptionReason,
    LgRun,
    SheetBalance,
    SheetRoleContribution,
} from '../shared/models';

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Integer fils → "1,234.567" (3 dp, thousands separators). Display only. */
export function fmtBhd(fils: number): string {
    const sign = fils < 0 ? '-' : '';
    const abs = Math.abs(fils);
    const whole = Math.floor(abs / 1000)
        .toString()
        .replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return `${sign}${whole}.${String(abs % 1000).padStart(3, '0')}`;
}

/** ISO yyyy-mm-dd → "05 Jul 2026"; anything else passes through unchanged. */
export function fmtIsoDate(iso: string): string {
    const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) {
        return iso;
    }
    return `${m[3]} ${MONTHS[Number(m[2]) - 1]} ${m[1]}`;
}

export const EXC_TYPE_LABEL: Record<LgExceptionReason, string> = {
    UNMATCHED_DEBIT: 'Unmatched Debit',
    UNMATCHED_CREDIT: 'Unmatched Credit',
    PARTIALLY_MATCHED_DEBIT: 'Partially Matched Debit',
    PARTIALLY_MATCHED_CREDIT: 'Partially Matched Credit',
    DUPLICATE: 'Duplicate Posting',
    AMOUNT_MISMATCH: 'Amount Mismatch',
    // GOAL-3 register family:
    NON_ISSUANCE_CREDIT: 'Non-Issuance Credit',
    UNRESOLVED_BATCH_DEBIT: 'Unresolved Batch Debit',
    UNMATCHED_LEDGER_DEBIT: 'Unmatched Ledger Debit',
    REGISTER_PAID_NO_LEDGER_DEBIT: 'Register Paid — No Ledger Debit',
    REGISTER_LAG_OPS_PAID: 'Ops Paid — Register Lag',
    KEY_COLLISION: 'Key Collision',
    EXTRACT_GAP: 'Ledger Extract Gap',
};

export const MISMATCHED_SHEET = 'Mismatched';
export const BALANCES_SHEET = 'Balances & Basis';

const SHEET_ROLE_LABEL: Record<SheetRoleContribution, string> = {
    ledger: 'GL ledger',
    breakdown: 'GL breakdown',
    register: 'Cheque register',
    skipped: 'Skipped',
};

const FIGURE_GROUP_LABEL: Record<ExplainedFigure['group'], string> = {
    input: 'Input population',
    balance: 'GL balance',
    matching: 'Matching',
    reconciliation: 'Reconciliation',
    exceptions: 'Exceptions',
    sheet: 'Per-sheet',
};

/**
 * GOAL-5: the reference sheet. Two tables saved into the exported workbook so the
 * balance-per-sheet and the how/why behind every number travel WITH the statement:
 *   1. Per-sheet balances — each worksheet's Σ credits / Σ debits / net (+ stated EoD).
 *   2. Explained figures — every headline number with its basis (how) and assessment (why).
 */
export function appendBalancesBasisSheet(workbook: ExcelJS.Workbook, run: LgRun): void {
    const ws = workbook.addWorksheet(BALANCES_SHEET);
    ws.columns = [28, 16, 10, 18, 18, 18, 18, 70].map((width) => ({ width }));

    ws.getCell('A1').value = 'Per-Sheet Balances & Number Basis (reference)';
    ws.getCell('A1').font = { bold: true, size: 13 };
    ws.getCell('A2').value = `Input: ${run.filename ?? '(unnamed)'}  ·  Review Date: ${fmtIsoDate(
        run.reconciliation?.asOf ?? run.asOf ?? ''
    )}  ·  SHA-256: ${(run.inputSha256 ?? '').slice(0, 12)}…`;

    // ── Table 1: per-sheet balances ───────────────────────────────────────────
    let row = 4;
    ws.getCell(`A${row}`).value = 'Balance of all amounts, per worksheet';
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    const balHeaders = ['Worksheet', 'Role', 'Rows', 'Σ Credits (BHD)', 'Σ Debits (BHD)', 'Net (BHD)', 'Stated EoD (BHD)', 'Basis'];
    ws.getRow(row).values = balHeaders;
    ws.getRow(row).font = { bold: true };
    row++;
    const sheetBalances: SheetBalance[] = run.sheetBalances ?? [];
    for (const sb of sheetBalances) {
        ws.getRow(row).values = [
            sb.sheet,
            SHEET_ROLE_LABEL[sb.role],
            sb.role === 'register' ? (sb.chequeCount ?? 0) : sb.parsedRows,
            sb.role === 'register' ? '' : fmtBhd(sb.creditFils),
            sb.role === 'register' ? '' : fmtBhd(sb.debitFils),
            sb.role === 'register' ? fmtBhd(sb.chequeFils ?? 0) : fmtBhd(sb.netFils),
            sb.statedEodFils !== undefined ? fmtBhd(sb.statedEodFils) : '',
            sb.basis,
        ];
        row++;
    }
    if (sheetBalances.length === 0) {
        ws.getCell(`A${row}`).value = '(no per-sheet balances recorded for this run)';
        row++;
    }

    // ── Table 2: explained figures (basis + assessment) ───────────────────────
    row += 1;
    ws.getCell(`A${row}`).value = 'Every reported number — how we got it, and why it matters';
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    ws.getRow(row).values = ['Figure', 'Value', 'Flag', 'Group', 'Basis (how)', '', '', 'Assessment (why)'];
    ws.getRow(row).font = { bold: true };
    // The Basis column spans B..G visually; keep header in E for clarity.
    row++;
    for (const f of run.explanations ?? []) {
        ws.getRow(row).values = [
            f.label,
            f.display,
            f.flag ? '⚠' : '',
            FIGURE_GROUP_LABEL[f.group],
            f.basis,
            '',
            '',
            f.assessment,
        ];
        if (f.flag) {
            ws.getRow(row).font = { bold: true };
        }
        row++;
    }
    if ((run.explanations ?? []).length === 0) {
        ws.getCell(`A${row}`).value = '(no explained figures recorded for this run)';
    }
}

/**
 * GOAL-3 R9: the register-mode statement. Line items are the OUTSTANDING
 * cheques themselves (real CHQ #, issuance date, payee, ops comments — §9.3
 * answered), grouped per branch inside the two age sections; the block is
 * GL-level (the file states one EoD per GL) and decomposes the Difference into
 * classified exceptions and the unexplained residual. `branchFilter` narrows
 * the SECTIONS to one branch (subtotals recomputed from the visible lines so
 * the sheet stays internally consistent); the block always stays GL-level.
 */
function buildRegisterStatement(
    ws: ExcelJS.Worksheet,
    recon: BranchReconciliation,
    outcomes: ChequeOutcome[],
    reviewDate: string,
    labels: (typeof GL_CATALOG)[keyof typeof GL_CATALOG]['statementLabels'],
    branchFilter?: string
): void {
    const lines = outcomes
        .filter((o) => o.state === 'OUTSTANDING')
        .filter((o) => branchFilter === undefined || (o.issuedBranch ?? '') === branchFilter)
        .sort(
            (a, b) =>
                (a.issuedBranch ?? '').localeCompare(b.issuedBranch ?? '') ||
                (a.issuedDate ?? '').localeCompare(b.issuedDate ?? '') ||
                (a.chequeNumber ?? '').localeCompare(b.chequeNumber ?? '')
        );
    const oldLines = lines.filter((o) => o.ageBucket === 'old');
    const currentLines = lines.filter((o) => o.ageBucket !== 'old');
    const oldFils = oldLines.reduce((s, o) => s + o.amountFils, 0);
    const currentFils = currentLines.reduce((s, o) => s + o.amountFils, 0);
    const comment = (o: ChequeOutcome) =>
        o.opsRemark ? `${o.opsRemark}${o.opsJournal ? ` · jrnl ${o.opsJournal}` : ''}${o.opsDate ? ` · ${fmtIsoDate(o.opsDate)}` : ''}` : '';

    ws.columns = [6, 10, 16, 14, 14, 24, 18, 26, 14, 18, 28, 16].map((width) => ({ width }));
    ws.getCell('A1').value = labels.statementTitle;
    ws.getCell('A1').font = { bold: true };
    ws.getCell('A2').value = `Branch: ${branchFilter ?? '(all branches)'}`;
    ws.getCell('A3').value = `Entity: ${recon.entity}  ·  GL: ${recon.gl}  ·  Review Date: ${fmtIsoDate(reviewDate)}`;
    if (branchFilter !== undefined) {
        ws.getCell('A4').value =
            'Sections filtered to one branch; the reconciliation block remains GL-level (consolidated).';
    }

    // GL-level reconciliation block (credit balances shown as magnitudes).
    const glMagFils = Math.abs(recon.glBalanceFils);
    const totalFils = oldFils + currentFils;
    const blockRows: [string, string][] = [
        ['GL Balance', fmtBhd(glMagFils)],
        [labels.totalLabel, fmtBhd(totalFils)],
        ['Diffrence', fmtBhd(glMagFils - totalFils)], // sic — matching the reference sample
        ['Classified exceptions (Sheet 2)', fmtBhd(Math.abs(recon.classifiedFils ?? 0))],
        ['Unexplained residual', fmtBhd(Math.abs(recon.residualFils ?? 0))],
        ['Derived balance (from postings)', fmtBhd(Math.abs(recon.derivedBalanceFils ?? 0))],
        ['Ledger extract gap', fmtBhd(Math.abs(recon.extractGapFils ?? 0))],
        ['Status', recon.balanced ? 'Balanced' : 'Not Balanced'],
    ];
    blockRows.forEach(([label, value], i) => {
        ws.getCell(`K${2 + i}`).value = label;
        ws.getCell(`L${2 + i}`).value = value;
    });

    let row = 11;
    ws.getCell(`A${row}`).value = labels.oldTitle;
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    ws.getRow(row).values = [
        'No.',
        'Branch',
        'Amount (BHD)',
        'Issuance Date',
        'CHQ. #',
        'Payee',
        'Review Date',
        'Comment',
        'Register Status',
        'Processed/Returned',
    ];
    ws.getRow(row).font = { bold: true };
    row++;
    oldLines.forEach((o, i) => {
        ws.getRow(row).values = [
            i + 1,
            o.issuedBranch ?? '',
            fmtBhd(o.amountFils),
            fmtIsoDate(o.issuedDate ?? ''),
            o.chequeNumber ?? '',
            o.payee ?? '',
            fmtIsoDate(reviewDate),
            comment(o),
            o.status ?? '',
            '',
        ];
        row++;
    });
    ws.getCell(`A${row}`).value = 'Subtotal';
    ws.getCell(`A${row}`).font = { bold: true };
    ws.getCell(`C${row}`).value = fmtBhd(oldFils);
    ws.getCell(`C${row}`).font = { bold: true };
    row += 2;

    ws.getCell(`A${row}`).value = labels.currentTitle;
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    ws.getRow(row).values = [
        'No.',
        'Branch',
        'Amount (BHD)',
        'Issuance Date',
        'CHQ. #',
        'Payee',
        'Review Date',
        'Comment',
        'Register Status',
    ];
    ws.getRow(row).font = { bold: true };
    row++;
    currentLines.forEach((o, i) => {
        ws.getRow(row).values = [
            i + 1,
            o.issuedBranch ?? '',
            fmtBhd(o.amountFils),
            fmtIsoDate(o.issuedDate ?? ''),
            o.chequeNumber ?? '',
            o.payee ?? '',
            fmtIsoDate(reviewDate),
            comment(o),
            o.status ?? '',
        ];
        row++;
    });
    ws.getCell(`A${row}`).value = 'Subtotal';
    ws.getCell(`A${row}`).font = { bold: true };
    ws.getCell(`C${row}`).value = fmtBhd(currentFils);
    ws.getCell(`C${row}`).font = { bold: true };
}

/** Builds the per-branch two-sheet workbook; returns the xlsx bytes. */
export async function buildStatementWorkbook(
    run: LgRun,
    recon: BranchReconciliation,
    exceptions: LgException[],
    /** Register-mode runs: the per-cheque outcomes that feed the statement lines. */
    outcomes?: ChequeOutcome[],
    /** Register-mode runs: narrow the statement sections to one branch. */
    branchFilter?: string
): Promise<Buffer> {
    const reviewDate = run.reconciliation?.asOf ?? run.asOf ?? '';
    const labels = GL_CATALOG[glCodeOf(run)].statementLabels;
    const isRegister = run.mode === 'register' && outcomes !== undefined;
    const branchExceptions = isRegister
        ? exceptions
        : exceptions.filter(
              (e) => e.entity === recon.entity && e.gl === recon.gl && e.branchNumber === recon.branchNumber
          );
    const oldItems = branchExceptions.filter((e) => e.ageBucket === 'old');
    const currentItems = branchExceptions.filter((e) => e.ageBucket === 'current');

    const workbook = new ExcelJS.Workbook();

    if (isRegister) {
        buildRegisterStatement(workbook.addWorksheet(labels.sheetName), recon, outcomes!, reviewDate, labels, branchFilter);
        appendExceptionSheet(workbook, branchExceptions);
        appendBalancesBasisSheet(workbook, run);
        const bytes = await workbook.xlsx.writeBuffer();
        return Buffer.from(bytes as ArrayBuffer);
    }

    // ── Sheet 1: the breakdown (TCS) statement — suspense/FIFO fragment columns.
    // Only breakdown runs reach this path, so the columns describe fragments, never
    // cheques (GOAL-7 §6: TCS never in cheque columns).
    const ws = workbook.addWorksheet(labels.sheetName);
    ws.columns = [6, 16, 14, 20, 18, 12, 8, 14, 80].map((width) => ({ width }));

    ws.getCell('A1').value = labels.statementTitle;
    ws.getCell('A1').font = { bold: true };
    ws.getCell('A2').value = `Branch: ${recon.branchNumber}`;
    ws.getCell('A3').value = `Entity: ${recon.entity}  ·  GL: ${recon.gl}  ·  Review Date: ${fmtIsoDate(reviewDate)}`;

    // Reconciliation block (right of the header, like the reference sample).
    ws.getCell('K2').value = 'GL Balance';
    ws.getCell('L2').value = fmtBhd(recon.glBalanceFils);
    ws.getCell('K3').value = labels.totalLabel;
    ws.getCell('L3').value = fmtBhd(recon.oldFils + recon.currentFils);
    ws.getCell('K4').value = 'Diffrence'; // sic — matching the reference sample
    ws.getCell('L4').value = fmtBhd(recon.differenceFils);
    ws.getCell('K5').value = 'Status';
    ws.getCell('L5').value = recon.balanced ? 'Balanced' : 'Not Balanced';

    const sectionHeaders = [
        'No.',
        'Amount (BHD)',
        'Post Date',
        'Account Number',
        'Journal Number',
        'Log Code',
        'DR/CR',
        'Review Date',
        'Reason / Comment',
    ];
    const fragmentRow = (item: LgException, i: number): (string | number)[] => [
        i + 1,
        fmtBhd(item.outstandingFils),
        fmtIsoDate(item.postDate),
        item.accountNumber ?? '',
        item.journalNumber,
        item.logCode ?? '',
        item.direction === 'debit' ? 'DEBIT' : 'CREDIT',
        fmtIsoDate(reviewDate),
        item.message,
    ];

    let row = 7;
    ws.getCell(`A${row}`).value = labels.oldTitle;
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    ws.getRow(row).values = sectionHeaders;
    ws.getRow(row).font = { bold: true };
    row++;
    oldItems.forEach((item, i) => {
        ws.getRow(row).values = fragmentRow(item, i);
        row++;
    });
    ws.getCell(`A${row}`).value = 'Subtotal';
    ws.getCell(`A${row}`).font = { bold: true };
    ws.getCell(`B${row}`).value = fmtBhd(recon.oldFils);
    ws.getCell(`B${row}`).font = { bold: true };
    row += 2;

    ws.getCell(`A${row}`).value = labels.currentTitle;
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    ws.getRow(row).values = sectionHeaders;
    ws.getRow(row).font = { bold: true };
    row++;
    currentItems.forEach((item, i) => {
        ws.getRow(row).values = fragmentRow(item, i);
        row++;
    });
    ws.getCell(`A${row}`).value = 'Subtotal';
    ws.getCell(`A${row}`).font = { bold: true };
    ws.getCell(`B${row}`).value = fmtBhd(recon.currentFils);
    ws.getCell(`B${row}`).font = { bold: true };

    appendExceptionSheet(workbook, branchExceptions);
    appendBalancesBasisSheet(workbook, run);

    const bytes = await workbook.xlsx.writeBuffer();
    return Buffer.from(bytes as ArrayBuffer);
}

/** Sheet 2: the mismatched / exception items (shared by both statement modes). */
function appendExceptionSheet(workbook: ExcelJS.Workbook, exceptions: LgException[]): void {
    const wsExc = workbook.addWorksheet(MISMATCHED_SHEET);
    wsExc.columns = [10, 22, 10, 16, 14, 20, 18, 32, 80].map((width) => ({ width }));
    wsExc.getCell('A1').value = 'Reconciling Exceptions / Mismatched Items';
    wsExc.getCell('A1').font = { bold: true };
    wsExc.getRow(2).values = [
        'ID',
        'Type',
        'Direction',
        'Amount (BHD)',
        'Post Date',
        'Journal Number',
        'Account Number',
        'Log Description',
        'Reason / Required Action',
    ];
    wsExc.getRow(2).font = { bold: true };
    exceptions.forEach((exc, i) => {
        wsExc.getRow(3 + i).values = [
            `ROW-${exc.rowNumber}`,
            EXC_TYPE_LABEL[exc.reason],
            exc.direction === 'debit' ? 'DEBIT' : 'CREDIT',
            fmtBhd(exc.outstandingFils),
            fmtIsoDate(exc.postDate),
            exc.journalNumber,
            exc.accountNumber ?? '—',
            `${exc.logCode ?? '——'} · source row ${exc.rowNumber}${exc.sheet ? ` · ${exc.sheet}` : ''}`,
            exc.message,
        ];
    });
}
