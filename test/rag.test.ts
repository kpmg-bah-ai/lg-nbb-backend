import { daysUntilDue, deriveRag } from '../src/helpers/rag';

const today = new Date('2026-06-11T10:00:00Z');

describe('daysUntilDue', () => {
    it('returns positive days for future due dates', () => {
        expect(daysUntilDue('2026-06-14', today)).toBe(3);
    });

    it('returns 0 on deadline day regardless of time of day', () => {
        expect(daysUntilDue('2026-06-11', today)).toBe(0);
    });

    it('returns negative days when past due', () => {
        expect(daysUntilDue('2026-06-10', today)).toBe(-1);
        expect(daysUntilDue('2026-06-08', today)).toBe(-3);
    });
});

describe('deriveRag (WT-21/26/27)', () => {
    it('completed tasks are green even when past due', () => {
        expect(deriveRag({ status: 'completed', dueDate: '2026-06-01' }, today)).toBe('green');
    });

    it('tasks without a due date are green', () => {
        expect(deriveRag({ status: 'in_progress', dueDate: undefined as unknown as string }, today)).toBe('green');
    });

    it('past-due tasks are red', () => {
        expect(deriveRag({ status: 'in_progress', dueDate: '2026-06-10' }, today)).toBe('red');
    });

    it('deadline day is amber', () => {
        expect(deriveRag({ status: 'not_started', dueDate: '2026-06-11' }, today)).toBe('amber');
    });

    it('approaching deadline (within 2 days) is amber', () => {
        expect(deriveRag({ status: 'in_progress', dueDate: '2026-06-13' }, today)).toBe('amber');
    });

    it('more than 2 days out is green', () => {
        expect(deriveRag({ status: 'in_progress', dueDate: '2026-06-14' }, today)).toBe('green');
    });
});
