/**
 * GOAL-7 §3 — validates an ingested upload against the GL the user picked. Pure;
 * returns file-level ParseErrors (empty array = pass).
 *
 * Two independent checks:
 *  1. FAMILY — the workbook's resolved mode must equal the GL's mode. A TCS
 *     breakdown uploaded as MGR (or vice versa) is a wrong pick, not a bad file,
 *     so the message suggests the GL that DOES match what was uploaded.
 *  2. CONTENT — every distinct posting.gl value must resolve to the picked GL.
 *     A value resolving to ANOTHER catalog GL is GL_MISMATCH; a non-empty value
 *     resolving to NO catalog GL is UNKNOWN_GL. Each distinct value reports once.
 *
 * Sheet-role completeness (register needs ledger + register sheets, etc.) is
 * already enforced upstream by resolveMode() — MIXED_MODE / INCOMPLETE_REGISTER_INPUT
 * reject before this guard runs, so requiredSheetRoles needs no re-check here.
 */

import { GL_CATALOG, GlCode, GlDefinition, ParseError, resolveGlCode } from '../shared/models';
import { IngestResult } from './ingest';

export function validateGlUpload(gl: GlDefinition, result: IngestResult): ParseError[] {
    if (result.mode !== gl.mode) {
        const other = Object.values(GL_CATALOG).find((d) => d.mode === result.mode);
        return [
            {
                code: 'GL_MISMATCH',
                message:
                    `The upload parses as the ${result.mode} family, but GL ${gl.code} (${gl.name}) ` +
                    `expects ${gl.mode} input` +
                    (other ? ` — did you mean GL ${other.code} (${other.name})?` : ''),
            },
        ];
    }

    const errors: ParseError[] = [];
    const seen = new Set<string>();
    for (const posting of result.postings) {
        const raw = posting.gl.trim();
        if (!raw || seen.has(raw)) {
            continue;
        }
        seen.add(raw);
        const resolved = resolveGlCode(raw);
        if (resolved === undefined) {
            errors.push({
                code: 'UNKNOWN_GL',
                message:
                    `The rows carry GL "${raw}", which is not in the GL catalog — ` +
                    `known GLs: ${Object.keys(GL_CATALOG).join(', ')}`,
            });
        } else if (resolved !== gl.code) {
            errors.push({
                code: 'GL_MISMATCH',
                message: `The rows identify GL ${resolved}, but GL ${gl.code} (${gl.name}) was picked at upload`,
            });
        }
    }
    return errors;
}

export interface GlDetection {
    glCode?: GlCode;
    /** More than one catalog GL appeared in the rows (should not happen for VAT). */
    ambiguous?: GlCode[];
    /** Non-empty embedded GL values that resolve to nothing. */
    unknown?: string[];
}

/** Resolves the GL a parsed upload belongs to from its embedded content (GOAL-8 auto-detect). */
export function detectGlFromUpload(result: IngestResult): GlDetection {
    const found = new Set<GlCode>();
    const unknown = new Set<string>();
    for (const p of result.postings) {
        const raw = p.gl.trim();
        if (!raw) {
            continue;
        }
        const code = resolveGlCode(raw);
        if (code) {
            found.add(code);
        } else {
            unknown.add(raw);
        }
    }
    if (found.size === 1) {
        return { glCode: [...found][0] };
    }
    if (found.size > 1) {
        return { ambiguous: [...found] };
    }
    if (unknown.size > 0) {
        return { unknown: [...unknown] };
    }
    return {};
}
