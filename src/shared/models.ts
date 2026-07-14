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
    /** The cheque's outcome state — lets stored items drive the statement filter. */
    state?: ChequeState;
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
    | 'INCONSISTENT_STATED_BALANCE' // final-day rows disagree on End Date EoD Balance
    // GOAL-7 GL catalog:
    | 'UNKNOWN_GL' // the picked GL code (or a GL value in the rows) is not in the catalog
    | 'GL_MISMATCH'; // the upload's family or embedded GL code contradicts the picked GL

// ParseError/ParseSummary/ParseResult are value objects embedded in an LgRun (or held
// in memory) — they are never stored as standalone documents, so no BaseDocument.
export interface ParseError {
    code: ParseErrorCode;
    message: string;
    /** 1-based data-row number; omitted for file/header-level problems. */
    row?: number;
    /** Source field the problem relates to (canonical breakdown fields, or register/statement field names). */
    field?: string;
    /** Worksheet the problem relates to (multi-sheet workbooks). */
    sheet?: string;
    /**
     * Raw content of the wrong cell (trimmed, capped) — present only when a
     * NON-EMPTY value failed its column's declared format. Empty cells are
     * missing, never "wrong", and carry no value.
     */
    value?: string;
    /** Spreadsheet column the wrong value sat in, as the Excel letter (e.g. 'K'). */
    column?: string;
    /** That column's header text as written in the source file. */
    columnHeader?: string;
    /**
     * True when the row STILL parsed into a posting/cheque and the wrong value
     * was only set aside — tracking, not exclusion. Absent when the problem
     * excluded the row (or is file/header-level).
     */
    rowParsed?: boolean;
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
    /** Register-family: manual disposition from the Debit sheet's `reconciled` column. */
    reconciledNote?: string;
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

// ---------- GOAL-7: GL catalog ----------

/** Canonical GL codes the app knows — the discriminant persisted on LgRun.glCode. */
export type GlCode = '99801000' | 'D2810085';

/** How the dashboard scopes a GL's statement (GOAL-7 "Two scoping concepts"). */
export type GlBranchNav = 'branchSections' | 'consolidated';

/**
 * One first-class GL. Everything mode-, layout- or copy-specific hangs off this
 * definition so adding a GL is a catalog entry + a GL_UI entry — never a new screen.
 * Facts are lifted from the GOAL-6 docs (GL-99801000-MC-PAYABLE.md /
 * GL-D2810085-TCS-SWIFT.md); update those docs first, then this seed.
 */
export interface GlDefinition {
    code: GlCode;
    /** Spellings that identify this GL in source files (resolved via resolveGlCode). */
    codeVariants: string[];
    name: string;
    shortName: string;
    /** Parsing family this GL's uploads must resolve to (detect.ts). */
    mode: LgRunMode;
    /** Worksheet roles an upload must contain (resolveMode already enforces the combination). */
    requiredSheetRoles: SheetRole[];
    /** Header fingerprint that identifies this GL's sheets — documents detect.ts, asserted in tests. */
    fingerprint: string[];
    entity: string;
    currency: string;
    branchNav: GlBranchNav;
    /** DR/CR derivation: split Credit/Debit amount columns vs the sign of one amount column. */
    directionModel: 'columnSplit' | 'amountSign';
    /** Two-legged GL↔register key matching vs per-account FIFO offsetting. */
    matchModel: 'registerTwoLegged' | 'fifoByAccount';
    /** Stated-EoD with gap/residual decomposition vs derived == stated tie-out. */
    balanceModel: 'statedEodDecomposition' | 'derivedEqualsStated';
    /** Statement layout key — resolved by the frontend GL_UI registry and the export labels. */
    statement: 'chequeTwoSection' | 'suspenseFragments';
    /** Exception reasons this GL's pipeline can emit (drives UI badge subsets + tests). */
    exceptionReasons: LgExceptionReason[];
    /** Data sentinels the parsers honour. Empty object = none: absence expresses "missing". */
    sentinels: {
        neverPaidDate?: string;
        takeOnJournal?: string;
        noRegisterHitChequeNumber?: string;
    };
    /** Source column headers treated as PII for this GL (browser sanitizer + upload copy). */
    piiColumns: string[];
    /** Export workbook labels (lg/export.ts) — the de-hardcoded 'MCQ+OLD ITEM' family. */
    statementLabels: {
        sheetName: string;
        statementTitle: string;
        oldTitle: string;
        currentTitle: string;
        totalLabel: string;
    };
}

export const GL_CATALOG: Record<GlCode, GlDefinition> = {
    '99801000': {
        code: '99801000',
        codeVariants: ['99801000', '0099801000', '0000000099801000'],
        name: "MC PAYABLE — Manager's Cheques",
        shortName: 'MGR 99801000',
        mode: 'register',
        requiredSheetRoles: ['ledgerStatement', 'register'],
        fingerprint: ['Branch', 'Transaction Credit Amount', 'Transaction Debit Amount', 'c0_bank_code…c60'],
        entity: 'BH',
        currency: 'BHD',
        branchNav: 'branchSections',
        directionModel: 'columnSplit',
        matchModel: 'registerTwoLegged',
        balanceModel: 'statedEodDecomposition',
        statement: 'chequeTwoSection',
        exceptionReasons: [
            'NON_ISSUANCE_CREDIT',
            'UNRESOLVED_BATCH_DEBIT',
            'UNMATCHED_LEDGER_DEBIT',
            'REGISTER_PAID_NO_LEDGER_DEBIT',
            'REGISTER_LAG_OPS_PAID',
            'KEY_COLLISION',
            'EXTRACT_GAP',
        ],
        sentinels: {
            neverPaidDate: '1901-01-01',
            takeOnJournal: '999999999',
            noRegisterHitChequeNumber: '0',
        },
        // GL-99801000-MC-PAYABLE.md §8 "PII surface".
        piiColumns: [
            'Teller',
            'Detailed Description',
            'c10_payee_name',
            'c11_memb_no',
            'c14_issued_teller',
            'c22_matchd_teller',
            'c37_purchaser_name',
            'c38_beneficiary_name',
            'c39_beneficiary_adrs',
            'c43_beneficiary_id_no',
            'c44_beneficiary_tel_no',
            'c46_chqm_pur_name_2',
            'c51_issued_supvisor',
            'c57_applicant_tel_no',
            'RRN',
            'Card Number',
        ],
        statementLabels: {
            sheetName: 'MCQ+OLD ITEM',
            statementTitle: "GL Reconciliation — Outstanding Manager's Cheques (register-based)",
            oldTitle: "Old Items Outstanding – Old Manager's Checks",
            currentTitle: 'Outstanding MCQ  (Less than 1 year)',
            totalLabel: 'Total (OLD Item + MCQ)',
        },
    },
    D2810085: {
        code: 'D2810085',
        codeVariants: ['D2810085', '2810085'],
        name: 'Inter System Account — SWIFT (TCS)',
        shortName: 'TCS D2810085',
        mode: 'breakdown',
        requiredSheetRoles: ['breakdown'],
        fingerprint: ['Branch Number', 'Amount (BHD)'],
        entity: 'BH',
        currency: 'BHD',
        branchNav: 'consolidated',
        directionModel: 'amountSign',
        matchModel: 'fifoByAccount',
        balanceModel: 'derivedEqualsStated',
        statement: 'suspenseFragments',
        exceptionReasons: [
            'UNMATCHED_DEBIT',
            'UNMATCHED_CREDIT',
            'PARTIALLY_MATCHED_DEBIT',
            'PARTIALLY_MATCHED_CREDIT',
            'DUPLICATE',
            'AMOUNT_MISMATCH',
        ],
        sentinels: {},
        // GL-D2810085-TCS-SWIFT.md §5 "PII surface".
        piiColumns: ['Account Number', 'User ID', 'Username'],
        statementLabels: {
            sheetName: 'SUSPENSE OUTSTANDING',
            statementTitle: 'GL Reconciliation — Outstanding Suspense Fragments (FIFO)',
            oldTitle: 'Old Items Outstanding – Aged Suspense Fragments',
            currentTitle: 'Outstanding Fragments  (Less than 1 year)',
            totalLabel: 'Total (Old + Current Fragments)',
        },
    },
};

/** Normalises a GL spelling: trim, uppercase, strip leading zeros from all-digit text. */
export function normalizeGlText(text: string): string {
    const t = text.trim().toUpperCase();
    return /^\d+$/.test(t) ? t.replace(/^0+(?=\d)/, '') : t;
}

/** Resolves any documented spelling ('0000000099801000', '2810085', …) to its catalog code. */
export function resolveGlCode(text: string | undefined | null): GlCode | undefined {
    if (!text) {
        return undefined;
    }
    const norm = normalizeGlText(String(text));
    for (const def of Object.values(GL_CATALOG)) {
        if (def.codeVariants.some((v) => normalizeGlText(v) === norm)) {
            return def.code;
        }
    }
    return undefined;
}

/** GL of a stored run; legacy pre-GOAL-7 runs derive from their mode (GOAL-3 default: breakdown). */
export function glCodeOf(run: Pick<LgRun, 'glCode' | 'mode'>): GlCode {
    return run.glCode ?? (run.mode === 'register' ? '99801000' : 'D2810085');
}

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

/** Per-file provenance for a run created from a multi-file upload. */
export interface LgRunFile {
    filename?: string;
    format: 'xlsx' | 'csv';
    /** SHA-256 of this file's raw bytes. */
    sha256: string;
    bytes: number;
}

// ---------- GOAL-5: per-sheet balance reference + number provenance ----------

/** What a worksheet contributed to the run (GOAL-5 §per-sheet balance). */
export type SheetRoleContribution = 'ledger' | 'register' | 'breakdown' | 'skipped';

/**
 * The balance of all amounts on ONE worksheet — computed at ingest and stored on
 * the run as an at-a-glance, per-sheet reference (GOAL-5). Answers "what does each
 * sheet add up to, and does it tie?" without re-opening the source workbook.
 * All *Fils fields are integer fils; credit/debit are magnitudes, `netFils` is
 * signed (debit +, credit −).
 */
export interface SheetBalance {
    /** Worksheet name as it appeared in the input ("file › sheet" for multi-file uploads). */
    sheet: string;
    /** How this sheet was classified and used. */
    role: SheetRoleContribution;
    /** Rows from this sheet that parsed into postings (ledger/breakdown) — the counted population. */
    parsedRows: number;
    creditCount: number;
    debitCount: number;
    /** Σ |amount| fils of credit postings on this sheet. */
    creditFils: number;
    /** Σ |amount| fils of debit postings on this sheet. */
    debitFils: number;
    /** creditFils/debitFils signed: Σ (debit − credit) fils — the sheet's net movement. */
    netFils: number;
    /** register sheets: cheque rows parsed. */
    chequeCount?: number;
    /** register sheets: Σ cheque |amount| fils. */
    chequeFils?: number;
    /** ledger sheets that state one: the End Date EoD balance on the sheet's final-day rows, signed fils. */
    statedEodFils?: number;
    /** Reviewer-readable note: what this sheet is and how its balance was derived. */
    basis: string;
}

/**
 * One reported number with the story behind it (GOAL-5). Every headline figure the
 * run surfaces carries a `basis` (HOW it was derived — the formula/source) and an
 * `assessment` (WHY it matters / what a reviewer should conclude). Stored on the run
 * and rendered next to the number so no figure is unexplained.
 */
export interface ExplainedFigure {
    /** Stable machine key, e.g. 'glBalance', 'sumCredits', 'residual'. */
    key: string;
    /** Human label, e.g. "GL closing balance". */
    label: string;
    /** Signed integer fils for money figures; omitted for pure counts. */
    valueFils?: number;
    /** Raw integer for count figures (rows, cheques, branches). */
    count?: number;
    /** Pre-formatted display string (BHD 3-dp with sign, or the integer). */
    display: string;
    /** HOW the number was derived — the formula or source. */
    basis: string;
    /** WHY it matters / what it means for the reconciliation — the reviewer assessment. */
    assessment: string;
    /** Keys of the component figures this one is derived from (drill-through). */
    inputs?: string[];
    /** Display grouping. */
    group: 'input' | 'balance' | 'matching' | 'reconciliation' | 'exceptions' | 'sheet';
    /** Set for per-sheet figures. */
    sheet?: string;
    /** True when this figure signals a control break a reviewer must act on. */
    flag?: boolean;
}

/**
 * A persisted reconciliation run (GOAL.md §4 F9): the uploaded breakdown's identity
 * (hash), parse summary and errors, plus the F3 balances and F4 matching results
 * computed at ingest. Postings are not persisted — they are re-derivable from the
 * input, and 550k rows would blow Cosmos document limits.
 */
export interface LgRun extends BaseDocument {
    /** Display name of the input: the file's name, or "a.xlsx + b.xlsx" for multi-file uploads. */
    filename?: string;
    format: 'xlsx' | 'csv';
    /**
     * SHA-256 identity of the input — same input ⇒ same hash (GOAL.md §5 determinism).
     * Single file: hash of the raw bytes (unchanged from pre-multi-file runs).
     * Multi-file: hash of the sorted per-file hashes, so the same SET of files
     * dedupes regardless of upload order.
     */
    inputSha256: string;
    /** Per-file provenance; present only when the run came from a multi-file upload. */
    files?: LgRunFile[];
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
    // ---- GOAL-7 GL catalog ----
    /** Which catalog GL the user picked at upload; legacy runs derive via glCodeOf(). */
    glCode?: GlCode;
    /** Register rows parsed (outcomes live in `lgRunDetails` kind 'cheques', capped). */
    chequeCount?: number;
    chequesByState?: Partial<Record<ChequeState, number>>;
    /** Cheques with no issuance credit in the ledger window (legacy population). */
    preWindowChequeCount?: number;

    // ---- GOAL-5: per-sheet balance reference + number provenance ----
    /** Balance of all amounts per worksheet — the saved reference (GOAL-5). */
    sheetBalances?: SheetBalance[];
    /** Description/assessment of how every headline number was derived and why (GOAL-5). */
    explanations?: ExplainedFigure[];
}

/** BHD (and the sample data) use 3 decimal places. */
export const AMOUNT_SCALE = 1000;

export function filsToBhd(fils: number): number {
    return fils / AMOUNT_SCALE;
}
