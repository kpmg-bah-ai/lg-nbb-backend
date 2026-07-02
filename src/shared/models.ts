import { BaseDocument } from '../helpers/crudHelper';

// ---------- Users (WT-1, WT-2) ----------

/**
 * admin   — super-manager: provisions manager accounts, system checks (WT-1).
 * manager — full portfolio visibility, creates staff accounts, BRD manager rights (WT-2/4).
 * staff   — sees only allocated clients, works own tasks (WT-3/5).
 */
export type Role = 'admin' | 'manager' | 'staff';

export interface User extends BaseDocument {
    displayName: string;
    email: string;
    role: Role;
    /** Admins provision managers; managers provision staff. */
    createdBy?: string;
    active: boolean;
}

// ---------- Audit trail (WT-33) ----------

export interface AuditLogEntry extends BaseDocument {
    actor: string;
    /** e.g. 'task.signed_off', 'project.closed'. */
    action: string;
    entityType: string;
    entityId: string;
    details?: Record<string, unknown>;
}



/**
 * LG / Manager's-Cheque GL reconciliation — domain model (GOAL.md §2, §4 F2).
 *
 * The input is a raw GL "transaction breakdown" for a suspense/clearing account
 * (e.g. the SWIFT inter-system account behind Letters of Guarantee and Manager's
 * Cheques). Each spreadsheet row is a single GL posting. We normalise those rows
 * into `LedgerPosting`s that the matching / reconciliation engine (later slices)
 * can work with.
 *
 * Money: the sample data is BHD, which has 3 decimal places (fils). We keep every
 * amount as an integer number of **fils** so arithmetic never touches floats
 * (GOAL.md §6). `amountBhd` is a convenience decimal for display only.
 */

/** A single raw cell as produced by the xlsx/csv reader. */
export type RawCell = string | number | boolean | Date | null | undefined;

/** A raw row: header row is row 0, data rows follow. Cells are 0-indexed. */
export type RawRow = RawCell[];

/** Debit = money out of the account (positive amount); credit = money in (negative). */
export type PostingDirection = 'debit' | 'credit';

/** The canonical fields we extract from the breakdown. Mapped from headers by alias. */
export type CanonicalField =
    | 'entity'
    | 'branchNumber'
    | 'sbu'
    | 'level6'
    | 'level3'
    | 'level0'
    | 'glDesc'
    | 'glName'
    | 'gl'
    | 'accountNumber'
    | 'postDate'
    | 'postTime'
    | 'valueDate'
    | 'source'
    | 'logDescription'
    | 'currency'
    | 'amountFcy'
    | 'amountLcy'
    | 'amountBhd'
    | 'journalNumber'
    | 'sequence'
    | 'userId'
    | 'username';

/** A normalised GL posting. `amount*Fils` are signed integer fils (positive = debit). */
export interface LedgerPosting {
    entity: string;
    branchNumber: string;
    sbu?: string;
    level6?: string;
    level3?: string;
    level0?: string;
    glDesc?: string;
    glName?: string;
    gl: string;
    accountNumber?: string;
    /** ISO date (yyyy-mm-dd). */
    postDate: string;
    /** Raw time string as supplied, e.g. '06:53:19.46'. */
    postTime?: string;
    valueDate?: string;
    source?: string;
    logDescription: string;
    /** Leading 6-digit posting-type code parsed from logDescription, e.g. '020050'. */
    logCode?: string;
    currency: string;
    amountFcyFils?: number;
    amountLcyFils?: number;
    /** The reconciling amount, in signed integer fils (positive = debit, negative = credit). */
    amountBhdFils: number;
    /** Convenience decimal (amountBhdFils / 1000) — for display only, never for maths. */
    amountBhd: number;
    direction: PostingDirection;
    journalNumber: string;
    sequence?: string;
    userId?: string;
    username?: string;
    /** 1-based source row number (data rows only, header excluded) for traceability. */
    rowNumber: number;
}

/** Codes used on parse errors/warnings so callers/tests can assert precisely. */
export type ParseErrorCode =
    | 'MISSING_HEADER'
    | 'EMPTY_INPUT'
    | 'MISSING_FIELD'
    | 'BAD_AMOUNT'
    | 'BAD_DATE'
    | 'ZERO_AMOUNT';

export interface ParseError {
    code: ParseErrorCode;
    message: string;
    /** 1-based data-row number; omitted for file/header-level problems. */
    row?: number;
    field?: CanonicalField;
}

export interface ParseSummary {
    dataRows: number;
    parsed: number;
    debitCount: number;
    creditCount: number;
    /** Sum of signed amountBhdFils across parsed postings (≈0 ⇒ debits balance credits). */
    netFils: number;
    currencies: string[];
    branches: string[];
}

export interface ParseResult {
    postings: LedgerPosting[];
    errors: ParseError[];
    summary: ParseSummary;
}

/** BHD (and the sample data) use 3 decimal places. */
export const AMOUNT_SCALE = 1000;

export function filsToBhd(fils: number): number {
    return fils / AMOUNT_SCALE;
}
