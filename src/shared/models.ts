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

// ---------- Workflow tasks (WT-13, WT-21) ----------

export type TaskStatus = 'not_started' | 'in_progress' | 'completed';

/** A template task inside a project workflow; due dates resolve via helpers/dates.ts (WT-13). */
export interface TaskDefinition {
    name: string;
    order: number;
    /** "Day N of the month" for accounting-period tasks. */
    relativeDueDay?: number;
    /** Day offset from the project start date for advisory tasks. */
    durationDays?: number;
}

/** A concrete task instance; RAG health derives from status + dueDate (helpers/rag.ts, WT-21). */
export interface WorkflowTask extends BaseDocument {
    name: string;
    status: TaskStatus;
    dueDate?: string;
    assignedTo?: string;
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

// ---------- Notifications ----------

/** Queued outbound notification. Only `status` is contracted so far (admin health counts pending). */
export interface AppNotification extends BaseDocument {
    status: 'pending' | 'sent' | 'failed';
    recipient?: string;
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
export interface LedgerPosting extends BaseDocument {
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
    /** Worksheet the row came from (multi-sheet workbooks, GOAL.md §2.3). */
    sheet?: string;
}

/**
 * A posting as parsed from the file: no storage identity yet. `id`/timestamps are
 * assigned by CrudHelper if/when a later slice persists postings as LedgerPosting docs.
 */
export type ParsedPosting = Omit<LedgerPosting, keyof BaseDocument>;

/** Codes used on parse errors/warnings so callers/tests can assert precisely. */
export type ParseErrorCode =
    | 'MISSING_HEADER'
    | 'EMPTY_INPUT'
    | 'UNSUPPORTED_FORMAT'
    | 'SHEET_SKIPPED'
    | 'MISSING_FIELD'
    | 'BAD_AMOUNT'
    | 'BAD_DATE'
    | 'ZERO_AMOUNT';

// ParseError/ParseSummary/ParseResult are value objects embedded in an LgRun (or held
// in memory) — they are never stored as standalone documents, so no BaseDocument.
export interface ParseError {
    code: ParseErrorCode;
    message: string;
    /** 1-based data-row number; omitted for file/header-level problems. */
    row?: number;
    field?: CanonicalField;
    /** Worksheet the problem relates to (multi-sheet workbooks). */
    sheet?: string;
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
    postings: ParsedPosting[];
    errors: ParseError[];
    summary: ParseSummary;
}

// ---------- F3: GL balance ----------

/** Per-branch GL balance derived from the postings as at `asOf` (GOAL.md §4 F3). */
export interface BranchBalance {
    entity: string;
    gl: string;
    branchNumber: string;
    /** Signed integer fils (positive = net debit). */
    balanceFils: number;
    /** Display-only decimal — never for maths. */
    balance: number;
    postingCount: number;
    firstPostDate?: string;
    lastPostDate?: string;
}

// ---------- F4: matching ----------

/** Aging per the sample statement: "Old Items" (> 1 year) vs "Less than 1 year". */
export type AgeBucket = 'old' | 'current';

export type OutstandingReason =
    | 'UNMATCHED_DEBIT'
    | 'UNMATCHED_CREDIT'
    | 'PARTIALLY_MATCHED_DEBIT'
    | 'PARTIALLY_MATCHED_CREDIT';

/**
 * A posting (or the unmatched remainder of one) left over after debit↔credit pairing
 * (GOAL.md §3/§4 F4). Partial fragments arise from one-to-many clears, e.g. an 11k
 * credit offset by 2k + 9k debits leaves nothing; offset by 9k alone leaves a 2k
 * PARTIALLY_MATCHED_CREDIT fragment.
 */
export interface OutstandingItem {
    entity: string;
    gl: string;
    branchNumber: string;
    accountNumber?: string;
    postDate: string;
    direction: PostingDirection;
    /** |amount| of the source posting, integer fils. */
    originalFils: number;
    /** Still-unmatched |amount|, integer fils (≤ originalFils; < ⇒ partial). */
    outstandingFils: number;
    /** Display-only decimal of outstandingFils. */
    outstanding: number;
    logCode?: string;
    journalNumber: string;
    sequence?: string;
    /** Source data-row number (traceability, GOAL.md §5). */
    rowNumber: number;
    sheet?: string;
    ageBucket: AgeBucket;
    reason: OutstandingReason;
}

export interface MatchSummary {
    /** Review date balances/aging were computed against (yyyy-mm-dd). */
    asOf: string;
    /** Fields whose combination pairs debits with credits (§9.2 — default, confirm). */
    matchKey: string[];
    /** |amount| fils successfully paired debit↔credit. */
    matchedFils: number;
    outstandingCount: number;
    outstandingDebitFils: number;
    outstandingCreditFils: number;
    /** Σ signed outstanding — equals Σ signed postings (the F5 tie-out identity). */
    netOutstandingFils: number;
    oldCount: number;
    currentCount: number;
    byBranch: { branchNumber: string; outstandingCount: number; outstandingFils: number }[];
}

/**
 * A persisted reconciliation run (GOAL.md §4 F9): the uploaded breakdown's identity
 * (hash), parse summary and errors, plus the F3 balances and F4 matching results
 * computed at ingest. Postings are not persisted — they are re-derivable from the
 * input, and 550k rows would blow Cosmos document limits.
 */
export interface LgRun extends BaseDocument {
    filename?: string;
    format: 'xlsx' | 'csv';
    /** SHA-256 of the uploaded bytes — same input ⇒ same hash (GOAL.md §5 determinism). */
    inputSha256: string;
    /** id of an earlier run with the same inputSha256, when one exists. */
    duplicateOf?: string;
    /** User id of the uploader (F9 audit / F10 access control). */
    uploadedBy: string;
    summary: ParseSummary;
    /** Total number of parse errors; `errors` itself is capped when stored. */
    errorCount: number;
    errors: ParseError[];
    /** Review date used for balances and aging (defaults to the latest post date). */
    asOf?: string;
    /** F3: per-branch balances as at asOf; capped when stored (see balancesCount). */
    balances?: BranchBalance[];
    balancesCount?: number;
    /** F4: matching summary; `outstanding` is capped when stored (see outstandingCount). */
    matching?: MatchSummary;
    outstandingCount?: number;
    outstanding?: OutstandingItem[];
}

/** BHD (and the sample data) use 3 decimal places. */
export const AMOUNT_SCALE = 1000;

export function filsToBhd(fils: number): number {
    return fils / AMOUNT_SCALE;
}
