import { ParsedPosting } from '../../src/shared/models';

let rowCounter = 0;

/**
 * Builds a valid ParsedPosting for balance/matching tests. Direction follows the
 * sign of amountBhdFils (positive = debit) unless overridden.
 */
export function makePosting(overrides: Partial<ParsedPosting> & { amountBhdFils: number }): ParsedPosting {
    const fils = overrides.amountBhdFils;
    const debit = fils >= 0;
    return {
        entity: 'BH',
        branchNumber: '1',
        gl: 'D2810085',
        accountNumber: 'ACC-1',
        postDate: '2026-01-15',
        logDescription: debit ? '020050 DEBIT POSTING' : '020030 BGL CR POSTING',
        logCode: debit ? '020050' : '020030',
        currency: 'BHD',
        amountBhd: fils / 1000,
        direction: debit ? 'debit' : 'credit',
        journalNumber: 'J1',
        rowNumber: ++rowCounter,
        ...overrides,
    };
}

export function signedTotalFils(postings: ParsedPosting[]): number {
    return postings.reduce((sum, p) => sum + p.amountBhdFils, 0);
}
