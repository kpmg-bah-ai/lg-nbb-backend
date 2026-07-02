import { TaskDefinition } from '../shared/models';

const MS_PER_DAY = 86_400_000;

/** Adds whole days to an ISO date and returns the date part (YYYY-MM-DD). */
export function addDays(isoDate: string, days: number): string {
    return new Date(new Date(isoDate).getTime() + days * MS_PER_DAY).toISOString().slice(0, 10);
}

/** Number of days in a 'YYYY-MM' period, used to clamp relative due days like "Day 31". */
export function daysInPeriod(period: string): number {
    const [year, month] = period.split('-').map(Number);
    return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/**
 * Resolves a template task's concrete due date (WT-13):
 * accounting tasks use "Day N of the month" against the project period,
 * advisory tasks use a day offset from the project start date.
 */
export function resolveDueDate(def: TaskDefinition, startDate: string, period?: string): string {
    if (def.relativeDueDay !== undefined && period) {
        const day = Math.min(def.relativeDueDay, daysInPeriod(period));
        return `${period}-${String(day).padStart(2, '0')}`;
    }
    if (def.durationDays !== undefined) {
        return addDays(startDate, def.durationDays);
    }
    return startDate;
}
