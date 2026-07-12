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

/**
 * Which parsing family produced a run's postings (GOAL-3 §4.1 / GOAL.md §9.11 —
 * both families stay in scope). Runs stored before GOAL-3 have no `mode` and are
 * breakdown runs by definition.
 */
export type LgRunMode = 'breakdown' | 'register';

/** Role a worksheet plays within an uploaded workbook (GOAL-3 §4.1). */
export type SheetRole = 'breakdown' | 'ledgerStatement' | 'register' | 'unknown';

/** How a cheque's payment leg was resolved (GOAL-3 §4.4). */
export type MatchedVia = 'KEY' | 'POSTING_DATE_VARIANT' | 'BATCH_REF';

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

    // ---- Register-family (ledger-statement layout) extensions — GOAL-3 §4.2 ----
    /** Transaction Date — the key-leg date for GL↔register matching (GOAL-3 §4.4). */
    transactionDate?: string;
    /** Cheque Number column (sparse; display/disambiguation only). */
    chequeNumber?: string;
    /** e.g. 'DD FROM DEP A/C', 'DEBIT POSTING'. */
    transactionType?: string;
    teller?: string;
    accountName?: string;
    /** Carries batch `Ref.#` journal lists on DEBIT POSTING rows (GOAL-3 §4.4 pass 4). */
    detailedDescription?: string;
    /** Stated End Date EoD Balance on this row, signed integer fils (GOAL-3 §4.5). */
    statedEodFils?: number;
    statedPrevEodFils?: number;
    /** Manual disposition from the Debit sheet's `reconciled` column (file §5.2). */
    reconciledNote?: string;
}

// ---------- GOAL-3: cheque register (second input family) ----------

/** One cheque-register row, normalised (register-family runs only; GOAL-3 §4.3). */
export interface RegisterCheque {
    instrument?: string;
    chequeNumber?: string;
    amountFils: number;
    payee?: string;
    /** '01'…'05', '91', '92' — text, zero-padded; display + classification only. */
    status?: string;
    /** Authoritative issuance date — drives aging (GOAL.md §9.4). */
    issuedDate?: string;
    /** Issuance key leg (GOAL-3 §4.4 pass 1). */
    issuedPostDate?: string;
    issuedBranch?: string;
    /** Issuance key leg. */
    issuedJournal?: string;
    /** Payment key leg; the sentinel 1901-01-01 means "never paid" ⇒ undefined. */
    matchedPostDate?: string;
    /** Payment key leg; journal '0' means "never paid" ⇒ undefined. */
    matchedJournal?: string;
    stopReason?: string;
    cancelDate?: string;
    purchaser?: string;
    beneficiary?: string;
    /** ISO-numeric currency text mapped ('48' ⇒ 'BHD'). */
    currency?: string;
    opsRemark?: string;
    /** opsRemark normalises to PAID (case/whitespace variants). */
    opsPaid: boolean;
    opsJournal?: string;
    /** Parsed from dd/mm/yyyy TEXT via the column-scoped parser (GOAL.md §11.3). */
    opsDate?: string;
    rowNumber: number;
    sheet?: string;
}

export type ChequeState =
    | 'PAID' // both legs matched in the ledger (KEY / variant)
    | 'PAID_VIA_BATCH' // payment resolved from a batch debit's Ref.# list
    | 'OPS_PAID' // reviewer disposition says paid; ledger debit absent (register lag)
    | 'OUTSTANDING' // issuance credit in window, no payment evidence
    | 'REGISTER_MATCHED_NO_DEBIT' // register says paid; no in-window ledger debit
    | 'PRE_WINDOW' // no issuance credit in the ledger window (legacy 05 population)
    | 'STOPPED'; // status 04 with stop/cancel evidence, unmatched

/** Per-cheque outcome (stored chunked as detail kind 'cheques'; GOAL-3 R8). */
export interface ChequeOutcome extends RegisterCheque {
    state: ChequeState;
    matchedVia?: MatchedVia;
    /** Outstanding cheques only, keyed off issuedDate (GOAL-3 §4.4 aging). */
    ageBucket?: AgeBucket;
    /** The (date, journal, amount) key bucket had more than one occupant. */
    keyCollision?: boolean;
    /** Linked GL credit row (issuance). */
    issuanceRowNumber?: number;
    /** Linked GL debit row(s) (payment / batch share). */
    paymentRowNumbers?: number[];
}

/** Statement attributes carried onto an outstanding cheque's OutstandingItem (GOAL-3 §4.4). */
export interface ChequeAttributes {
    chequeNumber?: string;
    payee?: string;
    issuedDate?: string;
    status?: string;
    purchaser?: string;
    opsRemark?: string;
    registerRowNumber: number;
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
    | 'ZERO_AMOUNT'
    // GOAL-3 register family:
    | 'AMBIGUOUS_DIRECTION' // both amount columns non-zero on one ledger-statement row
    | 'MIXED_MODE' // breakdown and register families in one workbook — rejected
    | 'INCOMPLETE_REGISTER_INPUT' // register without ledger sheets (or vice versa)
    | 'INCONSISTENT_STATED_BALANCE'; // final-day rows disagree on End Date EoD Balance

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
    /** Register-family: the cheque this outstanding issuance credit belongs to (GOAL-3 §4.4). */
    cheque?: ChequeAttributes;
    /** Register-family: `Ref.#` journals parsed from a batch debit's Detailed Description. */
    batchRefs?: string[];
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
    /** Cleared sets emitted by the engine (GOAL-2 G1). */
    matchedSetCount: number;
    /** Sets whose every leg was fully consumed (no partial residual). */
    fullyClearedSetCount: number;
    byBranch: {
        branchNumber: string;
        outstandingCount: number;
        outstandingFils: number;
        matchedSetCount: number;
    }[];
}

// ---------- F6a: matched (cleared) sets (GOAL-2 G1) ----------

/** One posting's participation in a matched set. */
export interface MatchedLeg {
    postDate: string;
    direction: PostingDirection;
    /** |amount| of the source posting, integer fils. */
    originalFils: number;
    /** Portion of this leg offset within this set, integer fils (≤ originalFils). */
    matchedFils: number;
    journalNumber: string;
    sequence?: string;
    logCode?: string;
    /** Source data-row number (traceability, GOAL.md §5). */
    rowNumber: number;
    sheet?: string;
}

/**
 * A cleared set: the connected component of debit↔credit offsets within one
 * match-key group — how an instrument (or an account-FIFO chain, pending §9.2)
 * actually cleared. Covers 1:1, 1:N and M:N uniformly. A set containing a
 * partially-consumed leg is `fullyCleared: false`; the residual of that leg is a
 * separate OutstandingItem, never netted silently.
 */
export interface MatchedSet {
    entity: string;
    gl: string;
    branchNumber: string;
    accountNumber?: string;
    /** Σ offset fils inside this set (each offset counted once). */
    matchedFils: number;
    /**
     * True leg totals. A busy account can FIFO-chain thousands of postings into one
     * component, so stored `creditLegs`/`debitLegs` are capped (Cosmos 2MB) — counts
     * greater than the stored array length mean visible truncation, never silent.
     */
    creditLegCount: number;
    debitLegCount: number;
    creditLegs: MatchedLeg[];
    debitLegs: MatchedLeg[];
    /** Earliest credit post date in the set. */
    firstCreditDate: string;
    /** Latest debit post date in the set. */
    finalDebitDate: string;
    /** finalDebitDate − firstCreditDate in days (negative when a debit pre-dates its credit). */
    settledDays: number;
    /** Every participating leg fully consumed — nothing from this set is outstanding. */
    fullyCleared: boolean;
    /** Register-family: the cheque this set cleared (GOAL-3 §4.4). */
    chequeNumber?: string;
    /** Register-family: how the payment leg was resolved. */
    matchedVia?: MatchedVia;
}

// ---------- F6b: reconciliation exceptions (GOAL-2 G2) ----------

/**
 * OutstandingReason plus the classified cases: breakdown-mode DUPLICATE /
 * AMOUNT_MISMATCH (GOAL.md §3) and the register-family taxonomy (GOAL-3 §4.6).
 */
export type LgExceptionReason =
    | OutstandingReason
    | 'DUPLICATE'
    | 'AMOUNT_MISMATCH'
    // GOAL-3 register family:
    | 'NON_ISSUANCE_CREDIT' // credit with no register issuance match (take-on, transfers, redeems)
    | 'UNRESOLVED_BATCH_DEBIT' // batch debit whose Ref.# allocation failed or was partial
    | 'UNMATCHED_LEDGER_DEBIT' // debit with no register payment match and no refs
    | 'REGISTER_PAID_NO_LEDGER_DEBIT' // register says paid; the ledger window lacks the debit
    | 'REGISTER_LAG_OPS_PAID' // ops-PAID but register still unmatched (statement excludes)
    | 'KEY_COLLISION' // informational: key bucket shared by several instruments/legs
    | 'EXTRACT_GAP'; // run-level: derived balance ≠ stated EoD balance

/**
 * A reconciling exception: every outstanding item becomes exactly one exception
 * (exceptions ⊇ outstanding — nothing dropped), with the reason upgraded to
 * DUPLICATE / AMOUNT_MISMATCH where the classifier finds those patterns.
 */
export interface LgException {
    entity: string;
    gl: string;
    branchNumber: string;
    accountNumber?: string;
    postDate: string;
    direction: PostingDirection;
    /** |amount| of the source posting, integer fils. */
    originalFils: number;
    /** Still-unmatched |amount|, integer fils. */
    outstandingFils: number;
    logCode?: string;
    journalNumber: string;
    sequence?: string;
    rowNumber: number;
    sheet?: string;
    ageBucket: AgeBucket;
    reason: LgExceptionReason;
    /** Human-readable finding for reviewers (GOAL.md §3 — displayed, never hidden). */
    message: string;
    /** Source rows of related postings (duplicate twins / mismatch counterpart). */
    relatedRowNumbers?: number[];
}

export interface ExceptionSummary {
    total: number;
    byReason: Partial<Record<LgExceptionReason, number>>;
}

/**
 * A chunk of per-run detail rows stored in the `lgRunDetails` container (GOAL-2
 * G3/§8.3): matched sets and exceptions are too numerous for the run document
 * (Cosmos 2MB), so they are chunked and paged. `seq` orders the chunks.
 */
export interface LgRunDetailChunk extends BaseDocument {
    runId: string;
    kind: 'matchedSets' | 'exceptions' | 'cheques';
    seq: number;
    items: MatchedSet[] | LgException[] | ChequeOutcome[];
}

// ---------- F5: reconciliation & Difference ----------

/** One branch's reconciliation block (GOAL.md §2.2): GL Balance vs Σ outstanding. */
export interface BranchReconciliation {
    entity: string;
    gl: string;
    branchNumber: string;
    /** F3 balance, signed integer fils. */
    glBalanceFils: number;
    /** Σ signed outstanding fils (debits +, credits −). */
    outstandingNetFils: number;
    outstandingCount: number;
    oldCount: number;
    /** Σ |outstanding| fils of Old Items — the statement's Section A subtotal. */
    oldFils: number;
    currentCount: number;
    /** Σ |outstanding| fils of current (< 1 year) items — Section B subtotal. */
    currentFils: number;
    /** glBalanceFils − outstandingNetFils. Never rounded away (GOAL.md §5). */
    differenceFils: number;
    /** Display-only decimal of differenceFils. */
    difference: number;
    /** |differenceFils| ≤ tolerance (register mode: |residualFils| ≤ tolerance). */
    balanced: boolean;

    // ---- Register-family decomposition (GOAL-3 §4.5) — all signed integer fils ----
    /** The file's stated End Date EoD Balance (credit balances negative). */
    statedBalanceFils?: number;
    /** Balance derived from the postings (F3). */
    derivedBalanceFils?: number;
    /** derivedBalanceFils − statedBalanceFils — the ledger-extract gap. */
    extractGapFils?: number;
    /** Σ signed outstanding classified as exceptions (everything but outstanding cheques). */
    classifiedFils?: number;
    /** differenceFils − classifiedFils — the unexplained reviewer-chase number. */
    residualFils?: number;
}

export interface Reconciliation {
    asOf?: string;
    /** GOAL.md §6: balanced ⇔ |difference| ≤ 0.001 BHD, i.e. 1 fil, by default. */
    toleranceFils: number;
    /** Every branch balanced. */
    balanced: boolean;
    /** Σ |differenceFils| across branches — the headline audit figure. */
    totalAbsDifferenceFils: number;
    byBranch: BranchReconciliation[];
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
    /** F5: per-branch Difference / Balanced status derived from balances + outstanding. */
    reconciliation?: Reconciliation;
    /** F6a: total cleared sets (the sets themselves live in `lgRunDetails`, capped). */
    matchedSetCount?: number;
    /** F6b: total exceptions (the items live in `lgRunDetails`, capped). */
    exceptionCount?: number;
    exceptionsSummary?: ExceptionSummary;

    // ---- GOAL-3 register family ----
    /** Input family; runs stored before GOAL-3 have no mode and are breakdown runs. */
    mode?: LgRunMode;
    /** Register rows parsed (outcomes live in `lgRunDetails` kind 'cheques', capped). */
    chequeCount?: number;
    chequesByState?: Partial<Record<ChequeState, number>>;
    /** Cheques with no issuance credit in the ledger window (legacy population). */
    preWindowChequeCount?: number;
}

/** BHD (and the sample data) use 3 decimal places. */
export const AMOUNT_SCALE = 1000;

export function filsToBhd(fils: number): number {
    return fils / AMOUNT_SCALE;
}
