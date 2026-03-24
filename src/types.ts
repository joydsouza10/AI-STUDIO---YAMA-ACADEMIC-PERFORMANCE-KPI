export interface User {
  id: number;
  email: string;
  role: 'ADMIN' | 'MANAGER' | 'LECTURER';
  name: string;
  lecturerId?: number;
}

export interface KPIRecord {
  id: number;
  lecturerId: number;
  lecturerName: string;
  moduleId: number;
  moduleName: string;
  programmeName: string;
  course?: string;
  level: string; // 4, 5, 6
  group: string; // Evening, Weekday, Weekend
  term: string;
  intake: string;
  totalEnrolled: number;
  totalSubmissions: number;
  nonSubmissions: number;
  passes: number;
  fails: number;
  attendanceRate: number;
  meqSatisfaction: number;
  meqResponseRate: number;
  department?: string;
}

export interface KPIWeights {
  submissionWeight: number;
  passWeight: number;
  attendanceWeight: number;
  meqWeight: number;
}
