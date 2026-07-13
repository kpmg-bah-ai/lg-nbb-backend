import { explainRun, ExplainInput } from '../../src/lg/provenance';
import { SheetBalance } from '../../src/shared/models';

const sheetBalances: SheetBalance[] = [
    {
        sheet: 'Credit',
        role: 'ledger',
        parsedRows: 3,
        creditCount: 3,
        debitCount: 0,
        creditFils: 10_000,
        debitFils: 0,
        netFils: -10_000,
        statedEodFils: -2_000,
        basis: 'Ledger extract …',
    },
    {
        sheet: 'Debit',
        role: 'ledger',
        parsedRows: 2,
        creditCount: 0,
        debitCount: 2,
        creditFils: 0,
        debitFils: 7_000,
        netFils: 7_000,
        basis: 'Ledger extract …',
    },
    {
        sheet: 'Sheet1',
        role: 'register',
        parsedRows: 0,
        creditCount: 0,
        debitCount: 0,
        creditFils: 0,
        debitFils: 0,
        netFils: 0,
        chequeCount: 5,
        chequeFils: 12_000,
        basis: 'Cheque register …',
    },
];

function registerInput(): ExplainInput {
    return {
        mode: 'register',
        summary: { dataRows: 5, parsed: 5, debitCount: 2, creditCount: 3, netFils: -3_000, currencies: ['BHD'], branches: ['1'] },
        asOf: '2026-02-03',
        balances: [],
        sheetBalances,
        reconciliation: {
            asOf: '2026-02-03',
            toleranceFils: 1,
            balanced: false,
            totalAbsDifferenceFils: 500,
            byBranch: [
                {
                    entity: 'BH',
                    gl: '99801000',
                    branchNumber: '',
                    glBalanceFils: -2_000,
                    outstandingNetFils: -1_500,
                    outstandingCount: 4,
                    oldCount: 1,
                    oldFils: 900,
                    currentCount: 3,
                    currentFils: 600,
                    differenceFils: -500,
                    difference: -0.5,
                    balanced: false,
                    statedBalanceFils: -2_000,
                    derivedBalanceFils: -3_000,
                    extractGapFils: -1_000,
                    classifiedFils: 0,
                    residualFils: -500,
                },
            ],
        },
        exceptionsSummary: { total: 6, byReason: { NON_ISSUANCE_CREDIT: 4, UNMATCHED_LEDGER_DEBIT: 2 } },
        chequeCount: 5,
        chequesByState: { OUTSTANDING: 4, PAID: 1 },
    };
}

describe('explainRun (GOAL-5 number provenance)', () => {
    it('explains every headline register figure with a basis and an assessment', () => {
        const figures = explainRun(registerInput());
        const byKey = Object.fromEntries(figures.map((f) => [f.key, f]));

        expect(byKey.glBalance).toBeDefined();
        expect(byKey.derivedBalance).toBeDefined();
        expect(byKey.extractGap).toBeDefined();
        expect(byKey.difference).toBeDefined();
        expect(byKey.residual).toBeDefined();

        for (const f of figures) {
            expect(f.basis.length).toBeGreaterThan(0);
            expect(f.assessment.length).toBeGreaterThan(0);
            expect(f.display.length).toBeGreaterThan(0);
        }
    });

    it('records the GL balance value and where it came from', () => {
        const glBalance = explainRun(registerInput()).find((f) => f.key === 'glBalance')!;
        expect(glBalance.valueFils).toBe(-2_000);
        expect(glBalance.group).toBe('balance');
        expect(glBalance.basis).toMatch(/stated/i);
    });

    it('flags the extract gap and the unexplained residual when the run does not balance', () => {
        const figures = explainRun(registerInput());
        expect(figures.find((f) => f.key === 'extractGap')!.flag).toBe(true);
        expect(figures.find((f) => f.key === 'residual')!.flag).toBe(true);
    });

    it('derives the difference from the GL balance and outstanding (inputs are traceable)', () => {
        const difference = explainRun(registerInput()).find((f) => f.key === 'difference')!;
        expect(difference.inputs).toEqual(expect.arrayContaining(['glBalance', 'outstandingNet']));
    });

    it('emits a per-sheet figure for each stored sheet balance', () => {
        const sheetFigures = explainRun(registerInput()).filter((f) => f.group === 'sheet');
        expect(sheetFigures.map((f) => f.sheet)).toEqual(expect.arrayContaining(['Credit', 'Debit', 'Sheet1']));
    });

    it('explains breakdown-mode runs from the matching + reconciliation summary', () => {
        const figures = explainRun({
            mode: 'breakdown',
            summary: { dataRows: 4, parsed: 4, debitCount: 2, creditCount: 2, netFils: 0, currencies: ['BHD'], branches: ['1', '2'] },
            balances: [
                { entity: 'BH', gl: 'D2810085', branchNumber: '1', balanceFils: 500, balance: 0.5, postingCount: 2 },
                { entity: 'BH', gl: 'D2810085', branchNumber: '2', balanceFils: -500, balance: -0.5, postingCount: 2 },
            ],
            sheetBalances: [],
            matching: {
                asOf: '2026-01-15',
                matchKey: ['entity', 'gl', 'branchNumber', 'accountNumber'],
                matchedFils: 4_000,
                outstandingCount: 2,
                outstandingDebitFils: 500,
                outstandingCreditFils: 500,
                netOutstandingFils: 0,
                oldCount: 0,
                currentCount: 2,
                matchedSetCount: 1,
                fullyClearedSetCount: 1,
                byBranch: [],
            },
            reconciliation: {
                asOf: '2026-01-15',
                toleranceFils: 1,
                balanced: true,
                totalAbsDifferenceFils: 0,
                byBranch: [],
            },
            exceptionsSummary: { total: 2, byReason: { UNMATCHED_DEBIT: 1, UNMATCHED_CREDIT: 1 } },
        });
        const keys = figures.map((f) => f.key);
        expect(keys).toEqual(expect.arrayContaining(['glBalanceTotal', 'matched', 'outstandingNet', 'totalDifference', 'exceptions']));
        expect(figures.every((f) => f.basis && f.assessment)).toBe(true);
    });
});
