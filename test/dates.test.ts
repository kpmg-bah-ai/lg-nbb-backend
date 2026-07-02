import { addDays, daysInPeriod, resolveDueDate } from '../src/helpers/dates';

describe('addDays', () => {
    it('adds days across month boundaries', () => {
        expect(addDays('2026-01-30', 3)).toBe('2026-02-02');
    });

    it('supports negative offsets', () => {
        expect(addDays('2026-06-01', -1)).toBe('2026-05-31');
    });
});

describe('daysInPeriod', () => {
    it('knows month lengths', () => {
        expect(daysInPeriod('2026-06')).toBe(30);
        expect(daysInPeriod('2026-07')).toBe(31);
    });

    it('handles February and leap years', () => {
        expect(daysInPeriod('2026-02')).toBe(28);
        expect(daysInPeriod('2028-02')).toBe(29);
    });
});

describe('resolveDueDate (WT-13)', () => {
    it('uses "Day N of the month" for accounting periods', () => {
        expect(resolveDueDate({ name: 'Reconciliation', order: 1, relativeDueDay: 5 }, '2026-06-01', '2026-06')).toBe('2026-06-05');
    });

    it('clamps relative due days to the period length', () => {
        expect(resolveDueDate({ name: 'Close', order: 2, relativeDueDay: 31 }, '2026-02-01', '2026-02')).toBe('2026-02-28');
    });

    it('uses day offsets from the start date for advisory tasks', () => {
        expect(resolveDueDate({ name: 'Draft', order: 1, durationDays: 10 }, '2026-06-01')).toBe('2026-06-11');
    });

    it('falls back to the start date when nothing is defined', () => {
        expect(resolveDueDate({ name: 'Kickoff', order: 0 }, '2026-06-01')).toBe('2026-06-01');
    });

    it('ignores relativeDueDay without a period and uses durationDays instead', () => {
        expect(resolveDueDate({ name: 'X', order: 1, relativeDueDay: 5, durationDays: 2 }, '2026-06-01')).toBe('2026-06-03');
    });
});
