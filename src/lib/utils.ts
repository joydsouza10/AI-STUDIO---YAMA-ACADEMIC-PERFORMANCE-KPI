import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function calculateScore(record: any, weights: any) {
  const submissionRate = record.totalSubmissions / record.totalEnrolled || 0;
  const passRate = record.passes / record.totalSubmissions || 0;
  const attendanceRate = record.attendanceRate / 100 || 0;
  const meqSatisfaction = record.meqSatisfaction / 100 || 0;

  return (
    submissionRate * weights.submissionWeight +
    passRate * weights.passWeight +
    attendanceRate * weights.attendanceWeight +
    meqSatisfaction * weights.meqWeight
  ) * 100;
}
