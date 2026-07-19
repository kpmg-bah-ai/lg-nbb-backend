import { detectGlFromUpload, validateGlUpload } from '../../src/lg/glGuard';
import { IngestResult } from '../../src/lg/ingest';
import { GL_CATALOG } from '../../src/shared/models';
import { makePosting } from './helpers';

/** Minimal IngestResult around a set of postings. */
function ingestResult(mode: 'breakdown' | 'register' | 'statement', gls: string[]): IngestResult {
    const postings = gls.map((gl, i) => makePosting({ amountBhdFils: 1000 * (i + 1), gl }));
    return {
        mode,
        postings,
        errors: [],
        summary: {
            dataRows: postings.length,
            parsed: postings.length,
            debitCount: postings.length,
            creditCount: 0,
            netFils: postings.reduce((s, p) => s + p.amountBhdFils, 0),
            currencies: ['BHD'],
            branches: ['1'],
        },
    };
}

describe('validateGlUpload (GOAL-7 §3 guardrail)', () => {
    it('passes a TCS breakdown picked as TCS', () => {
        expect(validateGlUpload(GL_CATALOG.D2810085, ingestResult('breakdown', ['D2810085']))).toEqual([]);
    });

    it('passes an MGR register family picked as MGR — statement rows carry the padded account', () => {
        expect(
            validateGlUpload(GL_CATALOG['99801000'], ingestResult('register', ['0000000099801000', '99801000']))
        ).toEqual([]);
    });

    it('rejects a family mismatch with GL_MISMATCH naming both GLs', () => {
        const errors = validateGlUpload(GL_CATALOG['99801000'], ingestResult('breakdown', ['D2810085']));
        expect(errors).toHaveLength(1);
        expect(errors[0].code).toBe('GL_MISMATCH');
        expect(errors[0].message).toContain('99801000');
        expect(errors[0].message).toContain('D2810085'); // the "did you mean" hint
    });

    it('rejects rows that identify the OTHER catalog GL with GL_MISMATCH', () => {
        const errors = validateGlUpload(GL_CATALOG.D2810085, ingestResult('breakdown', ['0099801000']));
        expect(errors).toHaveLength(1);
        expect(errors[0].code).toBe('GL_MISMATCH');
        expect(errors[0].message).toContain('99801000');
    });

    it('flags rows whose GL value resolves to no catalog GL with UNKNOWN_GL', () => {
        const errors = validateGlUpload(GL_CATALOG.D2810085, ingestResult('breakdown', ['D2810085', 'X9999999']));
        expect(errors).toHaveLength(1);
        expect(errors[0].code).toBe('UNKNOWN_GL');
        expect(errors[0].message).toContain('X9999999');
    });

    it('reports each distinct offending GL value once, not once per row', () => {
        const errors = validateGlUpload(
            GL_CATALOG.D2810085,
            ingestResult('breakdown', ['X1', 'X1', 'X1', 'D2810085'])
        );
        expect(errors).toHaveLength(1);
    });

    it('a family mismatch short-circuits content checks (one clear error, not a cascade)', () => {
        const errors = validateGlUpload(GL_CATALOG.D2810085, ingestResult('register', ['0000000099801000']));
        expect(errors).toHaveLength(1);
        expect(errors[0].code).toBe('GL_MISMATCH');
    });
});

describe('detectGlFromUpload (GOAL-8 content auto-detect)', () => {
    it('detects the VAT GL from an ingested statement upload', () => {
        const result = ingestResult('statement', ['8828010400010000', '8828010400010000']);
        expect(detectGlFromUpload(result)).toEqual({ glCode: '8828010400010000' });
    });

    it('reports an unknown embedded account', () => {
        const result = ingestResult('statement', ['8828019999990000']);
        expect(detectGlFromUpload(result)).toEqual({ unknown: ['8828019999990000'] });
    });

    it('reports ambiguity when rows carry more than one catalog GL', () => {
        const result = ingestResult('statement', ['8828010400010000', '8828010500010000']);
        expect(detectGlFromUpload(result)).toEqual({
            ambiguous: expect.arrayContaining(['8828010400010000', '8828010500010000']),
        });
    });

    it('returns an empty detection when nothing is embedded', () => {
        expect(detectGlFromUpload(ingestResult('statement', []))).toEqual({});
    });
});
