import {
    GL_CATALOG,
    GlCode,
    glCodeOf,
    normalizeGlText,
    resolveGlCode,
} from '../../src/shared/models';

describe('GL catalog (GOAL-7 §1)', () => {
    it('seeds the register/breakdown GLs with the right families', () => {
        expect(Object.keys(GL_CATALOG).sort()).toEqual(
            ['8828010400010000', '8828010500010000', '99801000', 'D2810085']
        );
        expect(GL_CATALOG['99801000'].mode).toBe('register');
        expect(GL_CATALOG['D2810085'].mode).toBe('breakdown');
    });

    it('resolves every documented spelling to its canonical code', () => {
        // MGR: ledger states 0000000099801000, register c5 states 0099801000 (GL doc §8).
        expect(resolveGlCode('99801000')).toBe('99801000');
        expect(resolveGlCode('0099801000')).toBe('99801000');
        expect(resolveGlCode('0000000099801000')).toBe('99801000');
        // TCS: breakdown rows carry D2810085; the balance sheet strips the D (GL doc §5).
        expect(resolveGlCode('D2810085')).toBe('D2810085');
        expect(resolveGlCode('2810085')).toBe('D2810085');
        expect(resolveGlCode('d2810085')).toBe('D2810085');
        expect(resolveGlCode(' D2810085 ')).toBe('D2810085');
    });

    it('returns undefined for unknown or empty spellings', () => {
        expect(resolveGlCode('12345678')).toBeUndefined();
        expect(resolveGlCode('')).toBeUndefined();
        expect(resolveGlCode(undefined)).toBeUndefined();
    });

    it('every variant in the catalog resolves back to its own GL (no cross-GL collisions)', () => {
        for (const def of Object.values(GL_CATALOG)) {
            for (const variant of def.codeVariants) {
                expect(resolveGlCode(variant)).toBe(def.code);
            }
        }
    });

    it('normalizeGlText strips leading zeros from digit strings only', () => {
        expect(normalizeGlText('0000000099801000')).toBe('99801000');
        expect(normalizeGlText('d2810085')).toBe('D2810085'); // letters block zero-stripping
    });

    it('glCodeOf derives legacy runs from their mode', () => {
        expect(glCodeOf({ glCode: 'D2810085', mode: 'register' })).toBe('D2810085'); // explicit wins
        expect(glCodeOf({ mode: 'register' })).toBe('99801000');
        expect(glCodeOf({ mode: 'breakdown' })).toBe('D2810085');
        expect(glCodeOf({})).toBe('D2810085'); // pre-GOAL-3 runs have no mode ⇒ breakdown
    });

    it('exception subsets match each family taxonomy', () => {
        expect(GL_CATALOG['99801000'].exceptionReasons).toEqual(
            expect.arrayContaining(['UNRESOLVED_BATCH_DEBIT', 'EXTRACT_GAP', 'REGISTER_LAG_OPS_PAID'])
        );
        expect(GL_CATALOG['99801000'].exceptionReasons).not.toContain('DUPLICATE');
        expect(GL_CATALOG['D2810085'].exceptionReasons).toEqual(
            expect.arrayContaining(['UNMATCHED_DEBIT', 'DUPLICATE', 'AMOUNT_MISMATCH'])
        );
        expect(GL_CATALOG['D2810085'].exceptionReasons).not.toContain('EXTRACT_GAP');
    });

    it('sentinels: MGR carries the register conventions, TCS carries none (GL docs §4)', () => {
        expect(GL_CATALOG['99801000'].sentinels).toEqual({
            neverPaidDate: '1901-01-01',
            takeOnJournal: '999999999',
            noRegisterHitChequeNumber: '0',
        });
        expect(GL_CATALOG['D2810085'].sentinels).toEqual({});
    });

    it('the type-level union matches the catalog keys', () => {
        const codes: GlCode[] = ['99801000', 'D2810085', '8828010400010000', '8828010500010000'];
        expect(codes.every((c) => GL_CATALOG[c].code === c)).toBe(true);
    });
});

describe('GOAL-8: VAT statement GLs', () => {
    it('seeds the two VAT GLs as statement-mode logics', () => {
        expect(Object.keys(GL_CATALOG).sort()).toEqual(
            ['8828010400010000', '8828010500010000', '99801000', 'D2810085']
        );
        const input = GL_CATALOG['8828010400010000'];
        const output = GL_CATALOG['8828010500010000'];
        expect(input.mode).toBe('statement');
        expect(output.mode).toBe('statement');
        expect(input.matchModel).toBe('none');
        expect(input.balanceModel).toBe('derivedEqualsStated');
        expect(input.statement).toBe('runningBalance');
        expect(input.branchNav).toBe('consolidated');
        expect(input.requiredSheetRoles).toEqual(['ledgerStatement']);
    });

    it('resolves each VAT account (the Nostro/BGL Account) to its GL', () => {
        expect(resolveGlCode('8828010400010000')).toBe('8828010400010000');
        expect(resolveGlCode('8828010500010000')).toBe('8828010500010000');
        // A different 16-digit account is unknown.
        expect(resolveGlCode('8828010300010000')).toBeUndefined();
    });

    it('marks the VAT PII surface (Account Name / Teller / Detailed Description)', () => {
        expect(GL_CATALOG['8828010400010000'].piiColumns).toEqual(
            expect.arrayContaining(['Account Name', 'Teller', 'Detailed Description'])
        );
    });
});
