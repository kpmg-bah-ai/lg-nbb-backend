/**
 * LG reconciliation — the authoritative Excel export (GOAL.md §4 F8, GOAL-2 G5).
 *
 * Builds the two-sheet workbook of GOAL.md §2.3 server-side from the stored run:
 *   Sheet 1 "MCQ+OLD ITEM" — the outstanding-items statement in the reference
 *           sample's layout (§2.2): header block with GL Balance / Total /
 *           Diffrence (sic) / Status, Section A (Old Items) and Section B
 *           (Outstanding MCQ, < 1 year) with subtotals.
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
import { BranchReconciliation, LgException, LgExceptionReason, LgRun } from '../shared/models';

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

export const STATEMENT_SHEET = 'MCQ+OLD ITEM';
export const MISMATCHED_SHEET = 'Mismatched';

/** Builds the per-branch two-sheet workbook; returns the xlsx bytes. */
export async function buildStatementWorkbook(
    run: LgRun,
    recon: BranchReconciliation,
    exceptions: LgException[]
): Promise<Buffer> {
    const reviewDate = run.reconciliation?.asOf ?? run.asOf ?? '';
    const branchExceptions = exceptions.filter(
        (e) => e.entity === recon.entity && e.gl === recon.gl && e.branchNumber === recon.branchNumber
    );
    const oldItems = branchExceptions.filter((e) => e.ageBucket === 'old');
    const currentItems = branchExceptions.filter((e) => e.ageBucket === 'current');

    const workbook = new ExcelJS.Workbook();

    // ── Sheet 1: the statement ────────────────────────────────────────────────
    const ws = workbook.addWorksheet(STATEMENT_SHEET);
    ws.columns = [6, 16, 14, 20, 18, 26, 14, 18, 36, 18, 28, 16].map((width) => ({ width }));

    ws.getCell('A1').value = 'GL Reconciliation — Outstanding Items Statement';
    ws.getCell('A1').font = { bold: true };
    ws.getCell('A2').value = `Branch: ${recon.branchNumber}`;
    ws.getCell('A3').value = `Entity: ${recon.entity}  ·  GL: ${recon.gl}  ·  Review Date: ${fmtIsoDate(reviewDate)}`;

    // Reconciliation block (right of the header, like the reference sample).
    ws.getCell('K2').value = 'GL Balance';
    ws.getCell('L2').value = fmtBhd(recon.glBalanceFils);
    ws.getCell('K3').value = 'Total (OLD Item + MCQ)';
    ws.getCell('L3').value = fmtBhd(recon.oldFils + recon.currentFils);
    ws.getCell('K4').value = 'Diffrence'; // sic — matching the reference sample
    ws.getCell('L4').value = fmtBhd(recon.differenceFils);
    ws.getCell('K5').value = 'Status';
    ws.getCell('L5').value = recon.balanced ? 'Balanced' : 'Not Balanced';

    let row = 7;
    ws.getCell(`A${row}`).value = "Old Items Outstanding – Old Manager's Checks";
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    const oldHeaders = [
        'No.',
        'Amount (BHD)',
        'Issuance Date',
        'CHQ. #',
        'Remitting Bank',
        'Date of Transfer to Old Items',
        'Review Date',
        'Debit Account',
        'Comment',
        'Processed/Returned',
        'Status',
    ];
    ws.getRow(row).values = oldHeaders;
    ws.getRow(row).font = { bold: true };
    row++;
    oldItems.forEach((item, i) => {
        ws.getRow(row).values = [
            i + 1,
            fmtBhd(item.outstandingFils),
            fmtIsoDate(item.postDate),
            `ROW ${item.rowNumber}${item.sheet ? ` · ${item.sheet}` : ''}`,
            '',
            '',
            fmtIsoDate(reviewDate),
            item.accountNumber ?? '',
            item.message,
            '',
            '',
        ];
        row++;
    });
    ws.getCell(`A${row}`).value = 'Subtotal';
    ws.getCell(`A${row}`).font = { bold: true };
    ws.getCell(`B${row}`).value = fmtBhd(recon.oldFils);
    ws.getCell(`B${row}`).font = { bold: true };
    row += 2;

    ws.getCell(`A${row}`).value = 'Outstanding MCQ  (Less than 1 year)';
    ws.getCell(`A${row}`).font = { bold: true };
    row++;
    const currentHeaders = ['No.', 'Amount (BHD)', 'Issuance Date', 'CHQ. #', 'Remitting Bank', 'Review Date', 'Comment', 'Status'];
    ws.getRow(row).values = currentHeaders;
    ws.getRow(row).font = { bold: true };
    row++;
    currentItems.forEach((item, i) => {
        ws.getRow(row).values = [
            i + 1,
            fmtBhd(item.outstandingFils),
            fmtIsoDate(item.postDate),
            `ROW ${item.rowNumber}${item.sheet ? ` · ${item.sheet}` : ''}`,
            '',
            fmtIsoDate(reviewDate),
            item.message,
            '',
        ];
        row++;
    });
    ws.getCell(`A${row}`).value = 'Subtotal';
    ws.getCell(`A${row}`).font = { bold: true };
    ws.getCell(`B${row}`).value = fmtBhd(recon.currentFils);
    ws.getCell(`B${row}`).font = { bold: true };

    // ── Sheet 2: the mismatched / exception items ────────────────────────────
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
    branchExceptions.forEach((exc, i) => {
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

    const bytes = await workbook.xlsx.writeBuffer();
    return Buffer.from(bytes as ArrayBuffer);
}
