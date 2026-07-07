import { detectExceptions } from '../../src/lg/exceptions';
import { matchPostings } from '../../src/lg/match';
import { makePosting } from './helpers';

const K = 1_000_000; // 1k BHD in fils

/** Convenience: engine → classifier, the way ingest runs them. */
function classify(postings: ReturnType<typeof makePosting>[], band?: number) {
    const { outstanding } = matchPostings(postings);
    return { outstanding, ...detectExceptions(outstanding, band === undefined ? {} : { mismatchBandFils: band }) };
}

describe('detectExceptions (G2) — nothing is ever dropped', () => {
    it('a debit with no matching credit is always surfaced (the GOAL.md §3 requirement)', () => {
        const debit = makePosting({ amountBhdFils: 125_000, postDate: '2024-06-15' });
        const { exceptions, summary } = classify([debit]);

        expect(exceptions).toHaveLength(1);
        expect(exceptions[0].reason).toBe('UNMATCHED_DEBIT');
        expect(exceptions[0].rowNumber).toBe(debit.rowNumber);
        expect(exceptions[0].outstandingFils).toBe(125_000);
        expect(exceptions[0].message).toContain('no matching credit');
        expect(summary.total).toBe(1);
        expect(summary.byReason.UNMATCHED_DEBIT).toBe(1);
    });

    it('emits exactly one exception per outstanding item (exceptions ⊇ outstanding)', () => {
        const postings = [
            makePosting({ amountBhdFils: -11 * K, accountNumber: 'A', postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: 9 * K, accountNumber: 'A', postDate: '2025-02-01' }),
            makePosting({ amountBhdFils: 700, accountNumber: 'B' }),
            makePosting({ amountBhdFils: -300, accountNumber: 'C' }),
        ];
        const { outstanding, exceptions } = classify(postings);

        expect(exceptions).toHaveLength(outstanding.length);
        // Amounts pass through untouched — Σ outstanding is preserved exactly.
        expect(exceptions.reduce((s, e) => s + e.outstandingFils, 0)).toBe(
            outstanding.reduce((s, o) => s + o.outstandingFils, 0)
        );
        expect(new Set(exceptions.map((e) => e.rowNumber))).toEqual(new Set(outstanding.map((o) => o.rowNumber)));
    });

    it('describes a partial residual with the offset and remaining amounts', () => {
        const { exceptions } = classify([
            makePosting({ amountBhdFils: -11 * K, postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: 9 * K, postDate: '2025-02-01' }),
        ]);
        expect(exceptions).toHaveLength(1);
        expect(exceptions[0].reason).toBe('PARTIALLY_MATCHED_CREDIT');
        expect(exceptions[0].message).toContain('9000.000');
        expect(exceptions[0].message).toContain('2000.000');
    });
});

describe('detectExceptions (G2) — DUPLICATE classification', () => {
    it('flags twin postings sharing journal, date, amount, account and direction', () => {
        const twinA = makePosting({ amountBhdFils: -555_000, postDate: '2025-03-14', journalNumber: 'JNL-RETRY' });
        const twinB = makePosting({ amountBhdFils: -555_000, postDate: '2025-03-14', journalNumber: 'JNL-RETRY' });
        const { exceptions, summary } = classify([twinA, twinB]);

        expect(exceptions).toHaveLength(2);
        expect(exceptions.every((e) => e.reason === 'DUPLICATE')).toBe(true);
        expect(exceptions[0].relatedRowNumbers).toEqual([exceptions[1].rowNumber]);
        expect(exceptions[1].relatedRowNumbers).toEqual([exceptions[0].rowNumber]);
        expect(exceptions[0].message).toContain('system retry');
        expect(summary.byReason.DUPLICATE).toBe(2);
    });

    it('does not flag postings that differ in journal number', () => {
        const { exceptions } = classify([
            makePosting({ amountBhdFils: -555_000, postDate: '2025-03-14', journalNumber: 'J1' }),
            makePosting({ amountBhdFils: -555_000, postDate: '2025-03-14', journalNumber: 'J2' }),
        ]);
        expect(exceptions.every((e) => e.reason === 'UNMATCHED_CREDIT')).toBe(true);
    });

    it('links all members of a triple retry', () => {
        const make = () => makePosting({ amountBhdFils: 250_000, postDate: '2025-05-01', journalNumber: 'J-3X' });
        const { exceptions } = classify([make(), make(), make()]);
        expect(exceptions.every((e) => e.reason === 'DUPLICATE')).toBe(true);
        expect(exceptions[0].relatedRowNumbers).toHaveLength(2);
    });
});

describe('detectExceptions (G2) — AMOUNT_MISMATCH classification', () => {
    // A near-miss pair always sits in different match-key groups (e.g. different
    // account numbers): inside one group, FIFO offsetting would have consumed one
    // side entirely. The classifier therefore scans the branch scope.
    it('flags the unambiguous near-miss pair within the band and cross-links it', () => {
        const debit = makePosting({ amountBhdFils: 10_000, postDate: '2025-01-01', accountNumber: 'ACC-A' });
        const credit = makePosting({ amountBhdFils: -10_500, postDate: '2025-01-05', accountNumber: 'ACC-B' });
        const { exceptions, summary } = classify([debit, credit]);

        expect(exceptions).toHaveLength(2);
        expect(exceptions.every((e) => e.reason === 'AMOUNT_MISMATCH')).toBe(true);
        const dExc = exceptions.find((e) => e.direction === 'debit')!;
        const cExc = exceptions.find((e) => e.direction === 'credit')!;
        expect(dExc.relatedRowNumbers).toEqual([credit.rowNumber]);
        expect(cExc.relatedRowNumbers).toEqual([debit.rowNumber]);
        expect(dExc.message).toContain('0.500'); // the difference, to the fil
        expect(summary.byReason.AMOUNT_MISMATCH).toBe(2);
    });

    it('leaves the pair unmatched when the gap exceeds the band', () => {
        const { exceptions } = classify([
            makePosting({ amountBhdFils: 10_000, accountNumber: 'ACC-A' }),
            makePosting({ amountBhdFils: -20_000, accountNumber: 'ACC-B' }),
        ]);
        expect(exceptions.map((e) => e.reason).sort()).toEqual(['UNMATCHED_CREDIT', 'UNMATCHED_DEBIT']);
    });

    it('honours a custom band', () => {
        const wide = classify(
            [
                makePosting({ amountBhdFils: 10_000, accountNumber: 'ACC-A' }),
                makePosting({ amountBhdFils: -20_000, accountNumber: 'ACC-B' }),
            ],
            15_000
        );
        expect(wide.exceptions.every((e) => e.reason === 'AMOUNT_MISMATCH')).toBe(true);
    });

    it('does not flag when the pair is ambiguous (two candidate debits)', () => {
        const { exceptions } = classify([
            makePosting({ amountBhdFils: 10_000, accountNumber: 'ACC-A' }),
            makePosting({ amountBhdFils: 10_100, accountNumber: 'ACC-B' }),
            makePosting({ amountBhdFils: -10_050, accountNumber: 'ACC-C' }),
        ]);
        expect(exceptions).toHaveLength(3);
        expect(exceptions.some((e) => e.reason === 'AMOUNT_MISMATCH')).toBe(false);
    });

    it('never mismatches across branches', () => {
        const { exceptions } = classify([
            makePosting({ amountBhdFils: 10_000, branchNumber: '1', accountNumber: 'ACC-A' }),
            makePosting({ amountBhdFils: -10_500, branchNumber: '2', accountNumber: 'ACC-B' }),
        ]);
        expect(exceptions.every((e) => e.reason !== 'AMOUNT_MISMATCH')).toBe(true);
    });

    it('excludes partial residuals from mismatch pairing', () => {
        const { exceptions } = classify([
            // ACC-A: 11k credit partially offset by a 9k debit → 2k partial residual
            makePosting({ amountBhdFils: -11_000, accountNumber: 'ACC-A', postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: 9_000, accountNumber: 'ACC-A', postDate: '2025-01-10' }),
            // ACC-B: a whole debit within the band of that residual
            makePosting({ amountBhdFils: 2_500, accountNumber: 'ACC-B' }),
        ]);
        // The residual is PARTIALLY_MATCHED — it already consumed a real counter-leg
        // and must not be re-explained as a keying error.
        expect(exceptions.some((e) => e.reason === 'AMOUNT_MISMATCH')).toBe(false);
        expect(exceptions.map((e) => e.reason).sort()).toEqual(['PARTIALLY_MATCHED_CREDIT', 'UNMATCHED_DEBIT']);
    });

    it('prefers DUPLICATE over AMOUNT_MISMATCH', () => {
        const twinA = makePosting({ amountBhdFils: -10_500, postDate: '2025-03-14', journalNumber: 'J-R', accountNumber: 'ACC-T' });
        const twinB = makePosting({ amountBhdFils: -10_500, postDate: '2025-03-14', journalNumber: 'J-R', accountNumber: 'ACC-T' });
        const debit = makePosting({ amountBhdFils: 10_000, postDate: '2025-01-01', accountNumber: 'ACC-D' });
        const { exceptions } = classify([twinA, twinB, debit]);

        const credits = exceptions.filter((e) => e.direction === 'credit');
        expect(credits).toHaveLength(2);
        expect(credits.every((e) => e.reason === 'DUPLICATE')).toBe(true);
        // With the twins claimed by DUPLICATE, no unambiguous credit remains — the
        // lone debit stays a plain unmatched item rather than a guessed mismatch.
        expect(exceptions.find((e) => e.direction === 'debit')!.reason).toBe('UNMATCHED_DEBIT');
    });
});
