/**
 * GOAL-8 — integration proof against the REAL data/VAT files: both GLs ingest as
 * statement mode, auto-identify from their embedded Nostro/BGL Account, and tie
 * out exactly (derived net == −stated End-Date EoD, gap 0). The asserted figures
 * were computed from the files with exceljs and independently re-verified — do
 * not re-derive them; a mismatch means the pipeline regressed.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { ingest } from '../../src/lg/ingest';
import { computeBranchBalances } from '../../src/lg/balance';
import { extractStatedBalance } from '../../src/lg/registerReconcile';
import { reconcileStatement } from '../../src/lg/statementReconcile';
import { detectGlFromUpload } from '../../src/lg/glGuard';

const DATA = join(__dirname, '..', '..', '..', 'data', 'VAT');
const cases = [
    { file: 'INPUT VAT RECEIVABLE MUBASHER - BHD_2023.xlsx',  gl: '8828010400010000', dataRows: 537,  parsed: 468,  netFils: 12_367_348, branches: 8 },
    { file: 'OUTPUT VAT PAYABLE MUBASHER - BHD_2023.xlsx',    gl: '8828010500010000', dataRows: 2615, parsed: 2578, netFils: 21_911_069, branches: 9 },
];

describe.each(cases)('real VAT file $gl', (c) => {
    it('ingests as statement mode, auto-detects the GL, and ties out', async () => {
        const result = await ingest(readFileSync(join(DATA, c.file)), { filename: c.file });
        expect(result.mode).toBe('statement');
        expect(result.summary.dataRows).toBe(c.dataRows);
        expect(result.summary.parsed).toBe(c.parsed);
        expect(result.summary.netFils).toBe(c.netFils);      // engine-signed derived net
        expect(result.summary.branches).toHaveLength(c.branches);
        expect(detectGlFromUpload(result)).toEqual({ glCode: c.gl });

        const balances = computeBranchBalances(result.postings);
        const stated = extractStatedBalance(result.postings);
        expect(stated.statedFils).toBe(c.netFils);           // −(file stated) == derived
        const recon = reconcileStatement(stated.statedFils, balances);
        expect(recon.byBranch[0].extractGapFils).toBe(0);
        expect(recon.balanced).toBe(true);
    });
});
