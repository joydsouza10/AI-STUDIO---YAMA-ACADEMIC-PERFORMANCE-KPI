import React, { useState, useEffect } from 'react';
import { 
  LayoutDashboard, 
  Users, 
  Upload, 
  Settings, 
  LogOut, 
  TrendingUp, 
  AlertTriangle, 
  CheckCircle2,
  ChevronRight,
  ChevronDown,
  BarChart3,
  PieChart as PieChartIcon,
  FileSpreadsheet,
  Cloud,
  RefreshCw,
  Sparkles,
  Download,
  Search,
  ArrowUp,
  ArrowDown
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { GoogleGenAI, Type } from "@google/genai";
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  LineChart,
  Line,
  Cell,
  PieChart,
  Pie
} from 'recharts';
import * as XLSX from 'xlsx';
import Markdown from 'react-markdown';
import { cn, calculateScore } from './lib/utils';
import { User, KPIRecord, KPIWeights } from './types';
import KPIFrameworkPage from './KPIFrameworkPage';

const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      const base64 = result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
};

const apiFetch = (url: string, options: RequestInit = {}) => {
  const token = localStorage.getItem('token');
  const headers = new Headers(options.headers);
  if (token) {
    headers.set('Authorization', `Bearer ${token}`);
  }
  return fetch(url, { ...options, headers });
};

// --- Hooks ---

const useKPI = () => {
  const [records, setRecords] = useState<KPIRecord[]>([]);
  const [weights, setWeights] = useState<KPIWeights | null>(null);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [filters, setFilters] = useState({
    programme: '',
    intake: '',
    level: '',
    module: '',
    lecturer: '',
    group: '',
    course: ''
  });

  const fetchData = async (silent = false) => {
    if (!silent) setLoading(true);
    setRefreshing(true);
    try {
      const [recRes, weightRes] = await Promise.all([
        apiFetch('/api/kpi/institution'),
        apiFetch('/api/kpi/weights')
      ]);
      if (!recRes.ok || !weightRes.ok) {
        throw new Error('Failed to fetch KPI data');
      }
      const recData = await recRes.json();
      const weightData = await weightRes.json();
      setRecords(Array.isArray(recData) ? recData : []);
      setWeights(weightData);
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  };

  useEffect(() => {
    fetchData();
    const interval = setInterval(() => fetchData(true), 5 * 60 * 1000); // Auto-refresh silently every 5 minutes
    return () => clearInterval(interval);
  }, []);

  const filteredRecords = records.filter(r => {
    const cleanModuleName = r.moduleName.split(' - ')[0].split(' (')[0].trim();
    return (
      (!filters.programme || r.programmeName === filters.programme) &&
      (!filters.intake || r.intake === filters.intake) &&
      (!filters.level || r.level === filters.level) &&
      (!filters.module || cleanModuleName === filters.module) &&
      (!filters.lecturer || r.lecturerName === filters.lecturer) &&
      (!filters.group || r.group === filters.group) &&
      (!filters.course || r.course === filters.course)
    );
  });

  return { records, filteredRecords, weights, loading, refreshing, filters, setFilters, refresh: () => fetchData(true) };
};

// --- Components ---

const GlobalFilters = ({ records, filters, setFilters, refresh, loading }: any) => {
  const programmes = Array.from(new Set(records.map((r: any) => r.programmeName)))
    .filter((p: any) => p && (p.includes('BM') || p.includes('BME')))
    .sort();
  const intakes = Array.from(new Set(records.map((r: any) => r.intake)))
    .filter((i: any) => i && i.includes('October Term 1'))
    .sort();
  const levels = Array.from(new Set(records.map((r: any) => r.level))).filter(Boolean).sort();
  const modules = Array.from(new Set(records.map((r: any) => r.moduleName.split(' - ')[0].split(' (')[0].trim()))).filter(Boolean).sort();
  const lecturers = Array.from(new Set(records.map((r: any) => r.lecturerName))).filter(Boolean).sort();
  const groups = Array.from(new Set(records.map((r: any) => r.group))).filter(Boolean).sort();

  const FilterSelect = ({ label, value, options, onChange }: any) => (
    <div className="space-y-1">
      <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{label}</label>
      <select 
        value={value} 
        onChange={(e) => onChange(e.target.value)}
        className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
      >
        <option value="">All {label}s</option>
        {options.map((opt: string) => (
          <option key={opt} value={opt}>{opt}</option>
        ))}
      </select>
    </div>
  );

  return (
    <div className="flex flex-col xl:flex-row gap-4 p-5 bg-white border border-slate-100 rounded-2xl shadow-sm mb-8">
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 flex-1">
        <FilterSelect label="Programme" value={filters.programme} options={programmes} onChange={(v: any) => setFilters({...filters, programme: v})} />
        <FilterSelect label="Intake" value={filters.intake} options={intakes} onChange={(v: any) => setFilters({...filters, intake: v})} />
        <FilterSelect label="Level" value={filters.level} options={levels} onChange={(v: any) => setFilters({...filters, level: v})} />
        <FilterSelect label="Module" value={filters.module} options={modules} onChange={(v: any) => setFilters({...filters, module: v})} />
        <FilterSelect label="Lecturer" value={filters.lecturer} options={lecturers} onChange={(v: any) => setFilters({...filters, lecturer: v})} />
        <FilterSelect label="Group" value={filters.group} options={groups} onChange={(v: any) => setFilters({...filters, group: v})} />
      </div>
      {refresh && (
        <div className="flex items-end justify-end mt-4 xl:mt-0 xl:pl-4 xl:border-l border-slate-100">
          <button 
            onClick={refresh}
            disabled={loading}
            className="flex items-center justify-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 bg-indigo-50 rounded-lg hover:bg-indigo-100 transition-colors h-[38px] disabled:opacity-50"
          >
            <RefreshCw size={16} className={loading ? "animate-spin" : ""} />
            Refresh Data
          </button>
        </div>
      )}
    </div>
  );
};

const SidebarItem = ({ icon: Icon, label, active, onClick }: any) => (
  <button
    onClick={onClick}
    className={cn(
      "flex items-center w-full gap-3 px-4 py-3 text-sm font-medium transition-colors rounded-lg",
      active 
        ? "bg-indigo-600 text-white shadow-lg shadow-indigo-200" 
        : "text-slate-500 hover:bg-slate-100 hover:text-slate-900"
    )}
  >
    <Icon size={20} />
    {label}
  </button>
);

const KPICard = ({ title, value, subtitle, icon: Icon, color }: any) => (
  <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
    <div className="flex items-start justify-between">
      <div>
        <p className="text-sm font-medium text-slate-500">{title}</p>
        <h3 className="mt-1 text-2xl font-bold text-slate-900">{value}</h3>
        {subtitle && <p className="mt-1 text-xs text-slate-400">{subtitle}</p>}
      </div>
      <div className={cn("p-3 rounded-xl", color)}>
        <Icon size={24} className="text-white" />
      </div>
    </div>
  </div>
);

// --- Pages ---

const LoginPage = ({ onLogin }: { onLogin: (user: User) => void }) => {
  const [email, setEmail] = useState('admin@institution.edu');
  const [password, setPassword] = useState('admin123');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError('');
    try {
      const res = await apiFetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password }),
      });
      const data = await res.json();
      if (res.ok) {
        localStorage.setItem('token', data.token);
        onLogin(data.user);
      }
      else setError(data.error);
    } catch (err) {
      setError('Connection failed');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div id="login-page-container" className="flex items-center justify-center min-h-screen bg-slate-50">
      <motion.div 
        id="login-card"
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-md p-8 bg-white shadow-2xl rounded-3xl"
      >
        <div className="flex flex-col items-center mb-8">
          <div className="p-4 mb-4 bg-indigo-600 rounded-2xl">
            <LayoutDashboard size={32} className="text-white" />
          </div>
          <h1 className="text-2xl font-bold text-slate-900">Academic KPI Portal</h1>
          <p className="text-slate-500">Sign in to access your dashboard</p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block mb-1 text-sm font-medium text-slate-700">Email Address</label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full px-4 py-2 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-transparent outline-none"
              required
            />
          </div>
          <div>
            <label className="block mb-1 text-sm font-medium text-slate-700">Password</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full px-4 py-2 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-transparent outline-none"
              required
            />
          </div>
          {error && <p className="text-sm text-red-500">{error}</p>}
          <button
            type="submit"
            disabled={loading}
            className="w-full py-3 font-semibold text-white transition-all bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:opacity-50"
          >
            {loading ? 'Signing in...' : 'Sign In'}
          </button>
        </form>
      </motion.div>
    </div>
  );
};

const exportToCSV = (records: KPIRecord[], filename: string) => {
  if (records.length === 0) return;
  const worksheet = XLSX.utils.json_to_sheet(records);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const ExecutiveOverview = ({ user }: { user: User }) => {
  const { records, filteredRecords, weights, loading, refreshing, filters, setFilters, refresh } = useKPI();

  if (loading || !weights) return <div className="p-8 text-slate-500">Loading dashboard...</div>;

  const totalEnrolled = filteredRecords.reduce((acc, r) => acc + r.totalEnrolled, 0);
  const totalSubmissions = filteredRecords.reduce((acc, r) => acc + r.totalSubmissions, 0);
  const totalPasses = filteredRecords.reduce((acc, r) => acc + r.passes, 0);
  const totalFails = filteredRecords.reduce((acc, r) => acc + r.fails, 0);
  const totalNonSubmissions = filteredRecords.reduce((acc, r) => acc + r.nonSubmissions, 0);
  
  const avgSubmissionRate = (totalSubmissions / totalEnrolled) * 100 || 0;
  const avgPassRate = (totalPasses / totalSubmissions) * 100 || 0;
  const avgAttendance = filteredRecords.length ? filteredRecords.reduce((acc, r) => acc + r.attendanceRate, 0) / filteredRecords.length : 0;
  const avgSatisfaction = filteredRecords.length ? filteredRecords.reduce((acc, r) => acc + r.meqSatisfaction, 0) / filteredRecords.length : 0;

  // Level comparison data
  const levels = ['4', '5', '6'];
  const levelData = levels.map(lvl => {
    const lvlRecs = filteredRecords.filter(r => String(r.level) === lvl);
    const enrolled = lvlRecs.reduce((acc, r) => acc + r.totalEnrolled, 0);
    const subs = lvlRecs.reduce((acc, r) => acc + r.totalSubmissions, 0);
    const passes = lvlRecs.reduce((acc, r) => acc + r.passes, 0);
    const fails = lvlRecs.reduce((acc, r) => acc + r.fails, 0);
    const nonSubs = lvlRecs.reduce((acc, r) => acc + r.nonSubmissions, 0);
    const att = lvlRecs.length ? lvlRecs.reduce((acc, r) => acc + r.attendanceRate, 0) / lvlRecs.length : 0;

    return {
      name: `Level ${lvl}`,
      submissionRate: Math.round((subs / enrolled) * 100) || 0,
      passRate: Math.round((passes / subs) * 100) || 0,
      attendanceRate: Math.round(att),
      passes,
      fails,
      nonSubmissions: nonSubs
    };
  });

  // Top Lecturers Data
  const lecturerGroups = filteredRecords.reduce((acc: any, r) => {
    if (!acc[r.lecturerId]) {
      acc[r.lecturerId] = {
        name: r.lecturerName,
        totalEnrolled: 0,
        totalSubmissions: 0,
        totalPasses: 0,
        totalMeq: 0,
        moduleCount: 0
      };
    }
    const group = acc[r.lecturerId];
    group.totalEnrolled += r.totalEnrolled;
    group.totalSubmissions += r.totalSubmissions;
    group.totalPasses += r.passes;
    group.totalMeq += r.meqSatisfaction;
    group.moduleCount += 1;
    return acc;
  }, {});

  const topLecturers = Object.values(lecturerGroups).map((l: any) => {
    const passRate = (l.totalPasses / l.totalSubmissions) * 100 || 0;
    const submissionRate = (l.totalSubmissions / l.totalEnrolled) * 100 || 0;
    const avgMeq = l.totalMeq / l.moduleCount;
    const score = (passRate * 0.5 + submissionRate * 0.3 + avgMeq * 0.2);
    return { name: l.name, score: Math.round(score), passRate: Math.round(passRate) };
  }).sort((a, b) => b.score - a.score).slice(0, 8);

  return (
    <div className="space-y-8">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Executive Overview</h1>
          <p className="text-slate-500">Institutional performance summary and level comparison</p>
        </div>
        <button
          onClick={() => exportToCSV(filteredRecords, 'executive_overview_kpi_data.csv')}
          className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-colors shadow-sm"
        >
          <Download className="w-4 h-4" />
          <span>Export Data</span>
        </button>
      </div>

      <GlobalFilters records={records} filters={filters} setFilters={setFilters} refresh={refresh} loading={refreshing} />

      <div className="grid grid-cols-1 gap-6 md:grid-cols-2 lg:grid-cols-4 xl:grid-cols-7">
        <KPICard title="Total Enrolled" value={totalEnrolled.toLocaleString()} icon={Users} color="bg-blue-500" />
        <KPICard title="Submission Rate" value={`${Math.round(avgSubmissionRate)}%`} icon={TrendingUp} color="bg-indigo-500" />
        <KPICard title="Pass Rate" value={`${Math.round(avgPassRate)}%`} icon={CheckCircle2} color="bg-emerald-500" />
        <KPICard title="Total Fails" value={totalFails.toLocaleString()} icon={AlertTriangle} color="bg-red-500" />
        <KPICard title="Non-Submissions" value={totalNonSubmissions.toLocaleString()} icon={AlertTriangle} color="bg-amber-500" />
        <KPICard title="Satisfaction" value={`${Math.round(avgSatisfaction)}%`} icon={Sparkles} color="bg-violet-500" />
        <KPICard title="Attendance" value={`${Math.round(avgAttendance)}%`} icon={Users} color="bg-cyan-500" />
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <h3 className="mb-6 text-lg font-semibold text-slate-900">Level Performance Comparison</h3>
          <div className="h-[350px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={levelData}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} />
                <YAxis axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} />
                <Bar dataKey="submissionRate" name="Submission %" fill="#6366f1" radius={[4, 4, 0, 0]} />
                <Bar dataKey="passRate" name="Pass %" fill="#10b981" radius={[4, 4, 0, 0]} />
                <Bar dataKey="attendanceRate" name="Attendance %" fill="#06b6d4" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <h3 className="mb-6 text-lg font-semibold text-slate-900">Pass vs Fail vs Non-Submission per Level</h3>
          <div className="h-[350px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={levelData}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} />
                <YAxis axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} />
                <Bar dataKey="passes" name="Passes" stackId="a" fill="#10b981" />
                <Bar dataKey="fails" name="Fails" stackId="a" fill="#ef4444" />
                <Bar dataKey="nonSubmissions" name="Non-Submissions" stackId="a" fill="#f59e0b" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
        <h3 className="mb-6 text-lg font-semibold text-slate-900">Top Performing Lecturers (Effectiveness Score)</h3>
        <div className="h-[350px]">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={topLecturers} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
              <XAxis type="number" domain={[0, 100]} axisLine={false} tickLine={false} />
              <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} width={150} tick={{ fontSize: 11 }} />
              <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} />
              <Bar dataKey="score" name="Effectiveness Score" fill="#6366f1" radius={[0, 4, 4, 0]} />
              <Bar dataKey="passRate" name="Pass Rate %" fill="#10b981" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
        <h3 className="mb-6 text-lg font-semibold text-slate-900">Pass Rate Trend across Levels</h3>
        <div className="h-[300px]">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={levelData}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis dataKey="name" axisLine={false} tickLine={false} />
              <YAxis axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} />
              <Line type="monotone" dataKey="passRate" stroke="#6366f1" strokeWidth={3} dot={{ r: 6, fill: '#6366f1' }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
};

const LevelAnalysisPage = () => {
  const { records, filteredRecords, weights, loading, refreshing, filters, setFilters, refresh } = useKPI();

  if (loading || !weights) return <div className="p-8 text-slate-500">Loading level analysis...</div>;

  const levels = ['4', '5', '6'];
  
  return (
    <div className="space-y-8">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Level Analysis (L4, L5, L6)</h1>
          <p className="text-slate-500">Deep dive into performance metrics per academic level</p>
        </div>
        <button
          onClick={() => exportToCSV(filteredRecords, 'level_analysis_kpi_data.csv')}
          className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-colors shadow-sm"
        >
          <Download className="w-4 h-4" />
          <span>Export Data</span>
        </button>
      </div>

      <GlobalFilters records={records} filters={filters} setFilters={setFilters} refresh={refresh} loading={refreshing} />

      <div className="grid grid-cols-1 gap-8">
        {levels.map(lvl => {
          const lvlRecs = filteredRecords.filter(r => String(r.level) === lvl);
          if (lvlRecs.length === 0) return null;

          const totalEnrolled = lvlRecs.reduce((acc, r) => acc + r.totalEnrolled, 0);
          const totalSubmissions = lvlRecs.reduce((acc, r) => acc + r.totalSubmissions, 0);
          const totalPasses = lvlRecs.reduce((acc, r) => acc + r.passes, 0);
          const totalFails = lvlRecs.reduce((acc, r) => acc + r.fails, 0);
          const totalNonSubs = lvlRecs.reduce((acc, r) => acc + r.nonSubmissions, 0);
          const avgAttendance = lvlRecs.reduce((acc, r) => acc + r.attendanceRate, 0) / lvlRecs.length;
          const avgSatisfaction = lvlRecs.reduce((acc, r) => acc + r.meqSatisfaction, 0) / lvlRecs.length;
          const avgMeqResponse = lvlRecs.reduce((acc, r) => acc + r.meqResponseRate, 0) / lvlRecs.length;

          const submissionRate = (totalSubmissions / totalEnrolled) * 100 || 0;
          const passRate = (totalPasses / totalSubmissions) * 100 || 0;
          const atRiskRate = ((totalFails + totalNonSubs) / totalEnrolled) * 100 || 0;

          const performanceIndex = (passRate * 0.5 + submissionRate * 0.3 + avgAttendance * 0.2);

          // Lecturer stats for this level
          const lvlLecturerGroups = lvlRecs.reduce((acc: any, r) => {
            if (!acc[r.lecturerId]) {
              acc[r.lecturerId] = { name: r.lecturerName, passes: 0, subs: 0 };
            }
            acc[r.lecturerId].passes += r.passes;
            acc[r.lecturerId].subs += r.totalSubmissions;
            return acc;
          }, {});

          const lvlLecturerStats = Object.values(lvlLecturerGroups).map((l: any) => ({
            name: l.name,
            passRate: Math.round((l.passes / l.subs) * 100) || 0
          })).sort((a, b) => b.passRate - a.passRate).slice(0, 5);

          return (
            <div key={lvl} className="p-8 bg-white border border-slate-100 rounded-3xl shadow-sm space-y-8">
              <div className="flex items-center justify-between border-b border-slate-50 pb-6">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-xl font-bold shadow-lg shadow-indigo-100">
                    {lvl}
                  </div>
                  <div>
                    <h2 className="text-xl font-bold text-slate-900">Level {lvl} Metrics</h2>
                    <p className="text-sm text-slate-500">Performance Index: <span className="font-bold text-indigo-600">{Math.round(performanceIndex)}</span></p>
                  </div>
                </div>
                <div className="text-right">
                  <div className={cn(
                    "px-4 py-2 rounded-xl text-sm font-bold inline-block",
                    performanceIndex > 80 ? "bg-emerald-50 text-emerald-700" :
                    performanceIndex > 60 ? "bg-amber-50 text-amber-700" : "bg-red-50 text-red-700"
                  )}>
                    {performanceIndex > 80 ? 'High Performing' : performanceIndex > 60 ? 'Stable' : 'Critical Attention'}
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-6">
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-400 uppercase">Total Students</p>
                  <p className="text-2xl font-bold text-slate-900">{totalEnrolled}</p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-400 uppercase">Submission Rate</p>
                  <p className="text-2xl font-bold text-slate-900">{Math.round(submissionRate)}%</p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-400 uppercase">Pass Rate</p>
                  <p className="text-2xl font-bold text-slate-900">{Math.round(passRate)}%</p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-400 uppercase">Attendance</p>
                  <p className="text-2xl font-bold text-slate-900">{Math.round(avgAttendance)}%</p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-400 uppercase text-red-500">At-Risk %</p>
                  <p className="text-2xl font-bold text-red-600">{Math.round(atRiskRate)}%</p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs font-bold text-slate-400 uppercase">Satisfaction</p>
                  <p className="text-2xl font-bold text-slate-900">{Math.round(avgSatisfaction)}%</p>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                <div className="space-y-4">
                  <h4 className="text-sm font-bold text-slate-900 uppercase tracking-wider">Module Breakdown</h4>
                  <div className="space-y-3">
                    {lvlRecs.slice(0, 5).map(r => (
                      <div key={r.id} className="flex items-center justify-between p-4 bg-slate-50 rounded-xl">
                        <div className="overflow-hidden">
                          <p className="text-sm font-bold text-slate-900 truncate">{r.moduleName}</p>
                          <p className="text-xs text-slate-500">{r.lecturerName}</p>
                        </div>
                        <div className="text-right flex items-center gap-4">
                          <div className="text-center">
                            <p className="text-[10px] font-bold text-slate-400 uppercase">Pass %</p>
                            <p className="text-sm font-bold text-indigo-600">{Math.round((r.passes / r.totalSubmissions) * 100) || 0}%</p>
                          </div>
                          <div className="text-center">
                            <p className="text-[10px] font-bold text-slate-400 uppercase">Risk</p>
                            <p className={cn(
                              "text-sm font-bold",
                              (r.fails + r.nonSubmissions) / r.totalEnrolled > 0.3 ? "text-red-600" : "text-emerald-600"
                            )}>{Math.round(((r.fails + r.nonSubmissions) / r.totalEnrolled) * 100)}%</p>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
                <div className="bg-slate-50 rounded-2xl p-6">
                  <h4 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4">Top Lecturers (Pass Rate)</h4>
                  <div className="h-[200px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={lvlLecturerStats} layout="vertical">
                        <XAxis type="number" hide />
                        <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} width={100} tick={{ fontSize: 10 }} />
                        <Tooltip />
                        <Bar dataKey="passRate" fill="#6366f1" radius={[0, 4, 4, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <div className="grid grid-cols-2 gap-4 mt-6">
                    <div className="text-center space-y-1">
                       <div className="text-2xl font-bold text-indigo-600">{Math.round(avgMeqResponse)}%</div>
                       <p className="text-[10px] font-bold text-slate-400 uppercase">MEQ Response</p>
                    </div>
                    <div className="text-center space-y-1">
                       <div className="text-2xl font-bold text-emerald-600">{Math.round(avgSatisfaction)}%</div>
                       <p className="text-[10px] font-bold text-slate-400 uppercase">Satisfaction</p>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

const LecturerPerformancePage = () => {
  const { records, filteredRecords, weights, loading, refreshing, filters, setFilters, refresh } = useKPI();
  const [aiLoading, setAiLoading] = useState(false);
  const [aiReport, setAiReport] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'stats' | 'matrix' | 'ai' | 'modules'>('stats');
  const [sortBy, setSortBy] = useState<'effectiveness' | 'passRate'>('effectiveness');
  const [moduleSortBy, setModuleSortBy] = useState<'lecturer' | 'module' | 'passRate' | 'submissionRate' | 'satisfaction'>('passRate');
  const [moduleSortOrder, setModuleSortOrder] = useState<'asc' | 'desc'>('desc');
  const [moduleSearch, setModuleSearch] = useState('');
  const [expandedRows, setExpandedRows] = useState<Set<string>>(new Set());
  const [lecturerSummaries, setLecturerSummaries] = useState<Record<string, string>>({});
  const [loadingSummaries, setLoadingSummaries] = useState<Record<string, boolean>>({});

  const toggleRow = (id: string) => {
    const newExpanded = new Set(expandedRows);
    if (newExpanded.has(id)) {
      newExpanded.delete(id);
    } else {
      newExpanded.add(id);
    }
    setExpandedRows(newExpanded);
  };

  const generateLecturerSummary = async (lecturer: any) => {
    setLoadingSummaries(prev => ({ ...prev, [lecturer.id]: true }));
    try {
      const ai = new GoogleGenAI({ apiKey: (import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || '') });
      
      const prompt = `
You are an Academic Performance Analyst. I am providing you with performance data for a specific lecturer.

Please provide a brief, personalized summary of the lecturer's performance based on their KPIs, highlighting strengths and areas for development. Keep it concise (2-3 short paragraphs).

Lecturer Data:
${JSON.stringify(lecturer, null, 2)}

Official KPI Targets to evaluate against:
- Student-Focused B3 Metrics: >90% submission rate, >75% first-time pass rate, >85% module pass rate, >80% continuation (UG Y1), >75% completion (UG), >60% MEQ response rate, >85% satisfaction.
- Curriculum Development (OfS B4): 80% asynchronous engagement, AI ethical use, ESG/DEI themes, live projects.
- Quality Assurance: <1% academic misconduct, timely feedback/marking, UCA IV compliance.

Focus on their specific module pass rates, submission rates, and satisfaction scores compared to the targets.
`;

      const response = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: prompt,
      });

      setLecturerSummaries(prev => ({ ...prev, [lecturer.id]: response.text || 'No summary generated.' }));
    } catch (err: any) {
      setLecturerSummaries(prev => ({ ...prev, [lecturer.id]: 'Failed to generate summary: ' + (err.message || 'Unknown error') }));
    } finally {
      setLoadingSummaries(prev => ({ ...prev, [lecturer.id]: false }));
    }
  };

  if (loading || !weights) return <div className="p-8 text-slate-500">Loading lecturer performance...</div>;

  // Group by lecturer
  const lecturerGroups = filteredRecords.reduce((acc: any, r) => {
    if (!acc[r.lecturerId]) {
      acc[r.lecturerId] = {
        id: r.lecturerId,
        name: r.lecturerName,
        department: r.department || 'N/A',
        modules: [],
        totalEnrolled: 0,
        totalSubmissions: 0,
        totalPasses: 0,
        totalFails: 0,
        totalNonSubs: 0,
        totalAttendance: 0,
        totalMeq: 0,
        totalMeqResponse: 0,
        moduleCount: 0
      };
    }
    const group = acc[r.lecturerId];
    group.modules.push(r);
    group.totalEnrolled += r.totalEnrolled;
    group.totalSubmissions += r.totalSubmissions;
    group.totalPasses += r.passes;
    group.totalFails += r.fails;
    group.totalNonSubs += r.nonSubmissions;
    group.totalAttendance += r.attendanceRate;
    group.totalMeq += r.meqSatisfaction;
    group.totalMeqResponse += r.meqResponseRate;
    group.moduleCount += 1;
    if (r.department && group.department === 'N/A') group.department = r.department;
    return acc;
  }, {});

  const lecturerStats = Object.values(lecturerGroups).map((group: any) => {
    const avgAttendance = group.totalAttendance / group.moduleCount;
    const avgMeq = group.totalMeq / group.moduleCount;
    const avgMeqResponse = group.totalMeqResponse / group.moduleCount;
    const submissionRate = (group.totalSubmissions / group.totalEnrolled) * 100 || 0;
    const passRate = (group.totalPasses / group.totalSubmissions) * 100 || 0;
    const failRate = (group.totalFails / group.totalSubmissions) * 100 || 0;
    const nonSubmissionRate = (group.totalNonSubs / group.totalEnrolled) * 100 || 0;
    
    // Lecturer Effectiveness Score: (0.5 × Pass %) + (0.3 × Submission %) + (0.2 × Satisfaction %)
    const effectivenessScore = (passRate * 0.5 + submissionRate * 0.3 + avgMeq * 0.2);

    return {
      ...group,
      avgAttendance,
      avgMeq,
      avgMeqResponse,
      submissionRate,
      passRate,
      failRate,
      nonSubmissionRate,
      effectivenessScore
    };
  });

  const sortedLecturerStats = [...lecturerStats].sort((a, b) => {
    if (sortBy === 'effectiveness') return b.effectivenessScore - a.effectivenessScore;
    if (sortBy === 'passRate') return b.passRate - a.passRate;
    return 0;
  });

  // Performance Matrix Data
  const allModules = Array.from(new Set(filteredRecords.map(r => r.moduleName))).sort();
  
  const lecturerLevelPerformance = sortedLecturerStats.map((l: any) => {
    const levels: Record<string, { passes: number; subs: number }> = {};
    l.modules.forEach((m: any) => {
      const lvl = `L${m.level}`;
      if (!levels[lvl]) levels[lvl] = { passes: 0, subs: 0 };
      levels[lvl].passes += m.passes;
      levels[lvl].subs += m.totalSubmissions;
    });
    const result: any = { name: l.name };
    ['L4', 'L5', 'L6'].forEach(lvl => {
      if (levels[lvl]) {
        result[lvl] = Math.round((levels[lvl].passes / levels[lvl].subs) * 100) || 0;
      }
    });
    return result;
  });

  const matrixData = lecturerStats.map((l: any) => {
    const moduleMap: Record<string, any> = {};
    l.modules.forEach((m: any) => {
      moduleMap[m.moduleName] = {
        passRate: (m.passes / m.totalSubmissions) * 100 || 0,
        level: m.level
      };
    });
    return {
      lecturer: l.name,
      modules: moduleMap
    };
  });

  const generateAIReport = async () => {
    setAiLoading(true);
    setAiReport(null);
    try {
      if (!lecturerStats || lecturerStats.length === 0) {
        throw new Error('No lecturer data provided for analysis.');
      }

      const ai = new GoogleGenAI({ apiKey: (import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || '') });
      
      const prompt = `
You are a Senior Academic Strategy Consultant and Data Scientist. I am providing you with aggregated performance data for individual lecturers.

Please provide a highly concise, analytical, and critical strategic report evaluating lecturers against the official BM/BME Programme KPI Framework.

Lecturer Data:
${JSON.stringify(lecturerStats, null, 2)}

Official KPI Targets to evaluate against:
- Student-Focused B3 Metrics: >90% submission rate, >75% first-time pass rate, >85% module pass rate, >80% continuation (UG Y1), >75% completion (UG), >60% MEQ response rate, >85% satisfaction.
- Curriculum Development (OfS B4): 80% asynchronous engagement, AI ethical use, ESG/DEI themes, live projects.
- Quality Assurance: <1% academic misconduct, timely feedback/marking, UCA IV compliance.

Your report MUST include:
1. Critical Performance Audit against KPIs: Identify the top 3 and bottom 3 lecturers based explicitly on the KPI targets above. Be critical—analyze WHY the bottom performers are failing (is it low engagement, high difficulty, or poor satisfaction?).
2. Predictive Analytics: Based on current trends (Pass Rates vs. MEQ), predict which lecturers or modules are likely to see a decline in performance in the next term if no intervention occurs.
3. Professional Development & People Management: Recommend specific CPDs, mentoring assignments, or Advance HE fellowship targets for specific lecturers based on their performance.
4. Executive Summary: 3 concise, high-level strategic actions for the Academic Board.

Format: Use sharp markdown headings, data-driven tables, and bullet points. Avoid fluff. Be direct and analytical.
`;

      const response = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: prompt,
      });

      setAiReport(response.text);
    } catch (err: any) {
      alert(err.message || 'Failed to generate report');
    } finally {
      setAiLoading(false);
    }
  };

  const moduleKPIs = filteredRecords.map(r => {
    const passRate = r.totalSubmissions > 0 ? (r.passes / r.totalSubmissions) * 100 : 0;
    const submissionRate = r.totalEnrolled > 0 ? (r.totalSubmissions / r.totalEnrolled) * 100 : 0;
    return {
      id: r.id,
      lecturer: r.lecturerName,
      module: r.moduleName,
      level: r.level,
      passRate,
      submissionRate,
      satisfaction: r.meqSatisfaction,
      enrolled: r.totalEnrolled
    };
  });

  const filteredModuleKPIs = moduleKPIs.filter(m => 
    m.lecturer.toLowerCase().includes(moduleSearch.toLowerCase()) || 
    m.module.toLowerCase().includes(moduleSearch.toLowerCase())
  ).sort((a, b) => {
    let valA = a[moduleSortBy];
    let valB = b[moduleSortBy];
    if (typeof valA === 'string' && typeof valB === 'string') {
      return moduleSortOrder === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
    }
    return moduleSortOrder === 'asc' ? (valA as number) - (valB as number) : (valB as number) - (valA as number);
  });

  const handleModuleSort = (key: 'lecturer' | 'module' | 'passRate' | 'submissionRate' | 'satisfaction') => {
    if (moduleSortBy === key) {
      setModuleSortOrder(moduleSortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setModuleSortBy(key);
      setModuleSortOrder('desc');
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex items-end justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Lecturer Performance Analysis</h1>
          <p className="text-slate-500">Detailed KPI analysis and effectiveness ranking</p>
        </div>
        <button
          onClick={generateAIReport}
          disabled={aiLoading || lecturerStats.length === 0}
          className="flex items-center gap-2 px-6 py-3 text-sm font-semibold text-white transition-all bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:opacity-50"
        >
          {aiLoading ? <RefreshCw size={18} className="animate-spin" /> : <Sparkles size={18} />}
          {aiLoading ? 'Analyzing Staff...' : 'Staff AI Insights'}
        </button>
      </div>

      <GlobalFilters records={records} filters={filters} setFilters={setFilters} refresh={refresh} loading={refreshing} />

      <div className="flex gap-4 border-b border-slate-100 mb-6">
        <button
          onClick={() => setActiveTab('stats')}
          className={cn(
            "pb-4 text-sm font-semibold transition-all border-b-2",
            activeTab === 'stats' ? "border-indigo-600 text-indigo-600" : "border-transparent text-slate-400 hover:text-slate-600"
          )}
        >
          Performance Stats
        </button>
        <button
          onClick={() => setActiveTab('matrix')}
          className={cn(
            "pb-4 text-sm font-semibold transition-all border-b-2",
            activeTab === 'matrix' ? "border-indigo-600 text-indigo-600" : "border-transparent text-slate-400 hover:text-slate-600"
          )}
        >
          Performance Matrix
        </button>
        <button
          onClick={() => setActiveTab('ai')}
          className={cn(
            "pb-4 text-sm font-semibold transition-all border-b-2 flex items-center gap-2",
            activeTab === 'ai' ? "border-indigo-600 text-indigo-600" : "border-transparent text-slate-400 hover:text-slate-600"
          )}
        >
          <Sparkles size={16} />
          AI Strategic Report
        </button>
        <button
          onClick={() => setActiveTab('modules')}
          className={cn(
            "pb-4 text-sm font-semibold transition-all border-b-2",
            activeTab === 'modules' ? "border-indigo-600 text-indigo-600" : "border-transparent text-slate-400 hover:text-slate-600"
          )}
        >
          Individual Module KPIs
        </button>
      </div>

      {activeTab === 'stats' && (
        <>
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-lg font-semibold text-slate-900">Lecturer Performance Overview</h3>
            <div className="flex items-center gap-3">
              <span className="text-sm text-slate-500">Sort by:</span>
              <select 
                value={sortBy} 
                onChange={(e) => setSortBy(e.target.value as any)}
                className="bg-white border border-slate-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
              >
                <option value="effectiveness">Effectiveness Score</option>
                <option value="passRate">Average Pass Rate</option>
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
              <h3 className="mb-6 text-lg font-semibold text-slate-900">Lecturer vs Pass Rate</h3>
              <div className="h-[300px]">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={lecturerStats.slice(0, 10)}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10 }} />
                    <YAxis axisLine={false} tickLine={false} />
                    <Tooltip />
                    <Bar dataKey="passRate" name="Pass Rate %" fill="#6366f1" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
              <h3 className="mb-6 text-lg font-semibold text-slate-900">Pass Rate by Level per Lecturer</h3>
              <div className="h-[300px]">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={lecturerLevelPerformance.slice(0, 10)}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10 }} />
                    <YAxis axisLine={false} tickLine={false} />
                    <Tooltip />
                    <Bar dataKey="L4" name="Level 4 %" fill="#6366f1" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="L5" name="Level 5 %" fill="#10b981" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="L6" name="Level 6 %" fill="#f59e0b" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          <div className="overflow-hidden bg-white border border-slate-100 rounded-2xl shadow-sm mt-8">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50">
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 w-10"></th>
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500">Lecturer</th>
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500">Department</th>
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500">Level & Module Breakdown</th>
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 text-center">Avg Pass Rate</th>
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 text-center">Effectiveness</th>
                    <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 text-right">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedLecturerStats.map((l: any) => (
                    <React.Fragment key={l.id}>
                      <tr 
                        className="hover:bg-slate-50 transition-colors cursor-pointer"
                        onClick={() => toggleRow(l.id)}
                      >
                        <td className="px-6 py-4 text-slate-400">
                          {expandedRows.has(l.id) ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                        </td>
                        <td className="px-6 py-4">
                          <div className="font-medium text-slate-900">{l.name}</div>
                          <div className="text-[10px] text-slate-400 font-bold uppercase">{l.moduleCount} Modules</div>
                        </td>
                        <td className="px-6 py-4">
                          <div className="text-sm text-slate-600">{l.department}</div>
                        </td>
                        <td className="px-6 py-4">
                          <div className="flex flex-wrap gap-2">
                            {l.modules.map((m: any, idx: number) => (
                              <div key={idx} className="px-2 py-1 bg-slate-50 border border-slate-100 rounded-lg text-[10px] flex items-center gap-2">
                                <span className="px-1 bg-indigo-100 text-indigo-700 rounded text-[8px] font-black">L{m.level}</span>
                                <span className="font-bold text-slate-700">{m.moduleName}</span>
                                <span className={cn(
                                  "font-black",
                                  (m.passes/m.totalSubmissions) >= 0.85 ? "text-emerald-600" : 
                                  (m.passes/m.totalSubmissions) >= 0.75 ? "text-amber-600" : "text-red-600"
                                )}>
                                  {Math.round((m.passes/m.totalSubmissions)*100)}%
                                </span>
                              </div>
                            ))}
                          </div>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <span className="text-sm font-bold text-slate-700">{Math.round(l.passRate)}%</span>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <span className="text-sm font-bold text-indigo-600">{Math.round(l.effectivenessScore)}</span>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <span className={cn(
                            "px-2 py-1 text-xs font-medium rounded-full",
                            l.effectivenessScore > 80 ? "bg-emerald-100 text-emerald-700" : 
                            l.effectivenessScore > 60 ? "bg-amber-100 text-amber-700" : "bg-red-100 text-red-700"
                          )}>
                            {l.effectivenessScore > 80 ? 'High' : l.effectivenessScore > 60 ? 'Average' : 'Low'}
                          </span>
                        </td>
                      </tr>
                      {expandedRows.has(l.id) && (
                        <tr className="bg-slate-50/50">
                          <td colSpan={7} className="px-6 py-6 border-t border-slate-100">
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                              <div className="p-4 bg-white border border-slate-200 rounded-xl shadow-sm">
                                <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">Lecturer Profile</h4>
                                <div className="space-y-4">
                                  <div>
                                    <p className="text-[10px] text-slate-400 uppercase font-bold mb-1">Department</p>
                                    <p className="text-sm font-medium text-slate-900">{l.department}</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] text-slate-400 uppercase font-bold mb-1">Total Teaching Load</p>
                                    <p className="text-sm font-medium text-slate-900">{l.totalEnrolled} Enrolled Students</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] text-slate-400 uppercase font-bold mb-1">Total Modules</p>
                                    <p className="text-sm font-medium text-slate-900">{l.moduleCount} Modules</p>
                                  </div>
                                </div>
                              </div>
                              <div className="md:col-span-2 p-4 bg-white border border-slate-200 rounded-xl shadow-sm">
                                <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">Assigned Modules Details</h4>
                                <div className="overflow-x-auto">
                                  <table className="w-full text-left text-sm">
                                    <thead>
                                      <tr className="text-slate-500 border-b border-slate-100">
                                        <th className="pb-2 font-medium">Module</th>
                                        <th className="pb-2 font-medium text-center">Level</th>
                                        <th className="pb-2 font-medium text-center">Enrolled</th>
                                        <th className="pb-2 font-medium text-center">Pass Rate</th>
                                        <th className="pb-2 font-medium text-center">Satisfaction</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-50">
                                      {l.modules.map((m: any, idx: number) => (
                                        <tr key={idx}>
                                          <td className="py-2 font-medium text-slate-900">{m.moduleName}</td>
                                          <td className="py-2 text-center text-slate-600">L{m.level}</td>
                                          <td className="py-2 text-center text-slate-600">{m.totalEnrolled}</td>
                                          <td className="py-2 text-center">
                                            <span className={cn(
                                              "font-bold",
                                              (m.passes/m.totalSubmissions) >= 0.85 ? "text-emerald-600" : 
                                              (m.passes/m.totalSubmissions) >= 0.75 ? "text-amber-600" : "text-red-600"
                                            )}>
                                              {Math.round((m.passes/m.totalSubmissions)*100)}%
                                            </span>
                                          </td>
                                          <td className="py-2 text-center">
                                            <span className={cn(
                                              "font-bold",
                                              m.meqSatisfaction >= 85 ? "text-emerald-600" : 
                                              m.meqSatisfaction >= 75 ? "text-amber-600" : "text-red-600"
                                            )}>
                                              {Math.round(m.meqSatisfaction)}%
                                            </span>
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            </div>
                            
                            <div className="mt-6 p-5 bg-indigo-50/50 border border-indigo-100 rounded-xl shadow-sm">
                              <div className="flex items-center justify-between mb-4">
                                <h4 className="text-sm font-bold text-indigo-900 flex items-center gap-2">
                                  <Sparkles size={16} className="text-indigo-600" />
                                  AI Performance Summary
                                </h4>
                                {!lecturerSummaries[l.id] && !loadingSummaries[l.id] && (
                                  <button
                                    onClick={() => generateLecturerSummary(l)}
                                    className="text-xs font-semibold text-indigo-600 hover:text-indigo-700 bg-white px-3 py-1.5 rounded-lg border border-indigo-200 shadow-sm transition-colors"
                                  >
                                    Generate Summary
                                  </button>
                                )}
                              </div>
                              
                              {loadingSummaries[l.id] ? (
                                <div className="flex items-center gap-3 text-sm text-indigo-600/70 py-4">
                                  <div className="w-4 h-4 border-2 border-indigo-600/30 border-t-indigo-600 rounded-full animate-spin" />
                                  Analyzing performance data...
                                </div>
                              ) : lecturerSummaries[l.id] ? (
                                <div className="text-sm text-slate-700 prose prose-sm max-w-none prose-p:leading-relaxed prose-headings:text-indigo-900 prose-a:text-indigo-600">
                                  <Markdown>{lecturerSummaries[l.id]}</Markdown>
                                </div>
                              ) : (
                                <p className="text-sm text-slate-500 italic">
                                  Click "Generate Summary" to get an AI-powered analysis of {l.name}'s performance against KPI targets.
                                </p>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
      {activeTab === 'matrix' && (
        <div className="overflow-hidden bg-white border border-slate-100 rounded-3xl shadow-sm">
          <div className="p-6 border-b border-slate-100">
            <h3 className="text-lg font-bold text-slate-900">Lecturer-Module Performance Matrix</h3>
            <p className="text-sm text-slate-500">Mapping lecturers across levels and modules (Values: Pass Rate %)</p>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse min-w-[800px]">
              <thead>
                <tr className="bg-slate-50">
                  <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 border-r border-slate-200 sticky left-0 bg-slate-50 z-10">Lecturer</th>
                  {allModules.map(m => (
                    <th key={m} className="px-4 py-4 text-[10px] font-bold tracking-wider uppercase text-slate-500 text-center min-w-[120px]">
                      <div className="truncate w-24 mx-auto" title={m}>{m}</div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {matrixData.map((row: any) => (
                  <tr key={row.lecturer} className="hover:bg-slate-50 transition-colors">
                    <td className="px-6 py-4 font-medium text-slate-900 border-r border-slate-200 sticky left-0 bg-white z-10">
                      {row.lecturer}
                    </td>
                    {allModules.map(m => {
                      const data = row.modules[m];
                      return (
                        <td key={m} className="px-4 py-4 text-center">
                          {data ? (
                            <div className="space-y-1">
                              <div className={cn(
                                "text-sm font-bold",
                                data.passRate >= 85 ? "text-emerald-600" : 
                                data.passRate >= 75 ? "text-amber-600" : "text-red-600"
                              )}>
                                {Math.round(data.passRate)}%
                              </div>
                              <div className="text-[9px] font-bold text-slate-400 uppercase">Lvl {data.level}</div>
                            </div>
                          ) : (
                            <span className="text-slate-200">—</span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {activeTab === 'ai' && (
        <div className="space-y-6">
          <div className="p-8 bg-white border border-slate-100 rounded-3xl shadow-sm text-center">
            <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center mx-auto mb-4">
              <Sparkles size={32} />
            </div>
            <h3 className="text-xl font-bold text-slate-900">AI Strategic Performance Report</h3>
            <p className="text-slate-500 max-w-md mx-auto mt-2">
              Generate a comprehensive analysis of lecturer effectiveness, risk identification, and strategic resource allocation recommendations.
            </p>
            <button
              onClick={generateAIReport}
              disabled={aiLoading || lecturerStats.length === 0}
              className="mt-6 flex items-center gap-2 px-8 py-3 text-sm font-semibold text-white transition-all bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:opacity-50 mx-auto"
            >
              {aiLoading ? <RefreshCw size={18} className="animate-spin" /> : <Sparkles size={18} />}
              {aiLoading ? 'Generating Strategic Analysis...' : 'Generate Strategic Report'}
            </button>
          </div>

          {aiReport && (
            <motion.div 
              initial={{ opacity: 0, y: 20 }} 
              animate={{ opacity: 1, y: 0 }} 
              className="p-10 bg-white border border-slate-100 rounded-3xl shadow-sm prose prose-slate max-w-none"
            >
              <div className="markdown-body">
                <Markdown>{aiReport}</Markdown>
              </div>
            </motion.div>
          )}
        </div>
      )}

      {activeTab === 'modules' && (
        <div className="space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <h3 className="text-lg font-semibold text-slate-900">Individual Module KPIs</h3>
            <div className="flex flex-col sm:flex-row items-center gap-4">
              <div className="flex items-center gap-2">
                <span className="text-sm text-slate-500 whitespace-nowrap">Sort by:</span>
                <select
                  value={`${moduleSortBy}-${moduleSortOrder}`}
                  onChange={(e) => {
                    const [by, order] = e.target.value.split('-');
                    setModuleSortBy(by as any);
                    setModuleSortOrder(order as any);
                  }}
                  className="bg-white border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                >
                  <option value="passRate-desc">Pass Rate (High to Low)</option>
                  <option value="passRate-asc">Pass Rate (Low to High)</option>
                  <option value="submissionRate-desc">Submission Rate (High to Low)</option>
                  <option value="submissionRate-asc">Submission Rate (Low to High)</option>
                  <option value="satisfaction-desc">Satisfaction (High to Low)</option>
                  <option value="satisfaction-asc">Satisfaction (Low to High)</option>
                  <option value="lecturer-asc">Lecturer (A-Z)</option>
                  <option value="lecturer-desc">Lecturer (Z-A)</option>
                  <option value="module-asc">Module (A-Z)</option>
                  <option value="module-desc">Module (Z-A)</option>
                </select>
              </div>
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                <input
                  type="text"
                  placeholder="Search lecturer or module..."
                  value={moduleSearch}
                  onChange={(e) => setModuleSearch(e.target.value)}
                  className="pl-9 pr-4 py-2 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 outline-none w-full sm:w-64"
                />
              </div>
            </div>
          </div>

          <div className="overflow-hidden bg-white border border-slate-100 rounded-2xl shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50">
                    <th 
                      className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 cursor-pointer hover:bg-slate-100"
                      onClick={() => handleModuleSort('lecturer')}
                    >
                      <div className="flex items-center gap-2">
                        Lecturer
                        {moduleSortBy === 'lecturer' && (moduleSortOrder === 'asc' ? <ArrowUp size={14} /> : <ArrowDown size={14} />)}
                      </div>
                    </th>
                    <th 
                      className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 cursor-pointer hover:bg-slate-100"
                      onClick={() => handleModuleSort('module')}
                    >
                      <div className="flex items-center gap-2">
                        Module
                        {moduleSortBy === 'module' && (moduleSortOrder === 'asc' ? <ArrowUp size={14} /> : <ArrowDown size={14} />)}
                      </div>
                    </th>
                    <th 
                      className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 cursor-pointer hover:bg-slate-100 text-center"
                      onClick={() => handleModuleSort('passRate')}
                    >
                      <div className="flex items-center justify-center gap-2">
                        Pass Rate
                        {moduleSortBy === 'passRate' && (moduleSortOrder === 'asc' ? <ArrowUp size={14} /> : <ArrowDown size={14} />)}
                      </div>
                    </th>
                    <th 
                      className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 cursor-pointer hover:bg-slate-100 text-center"
                      onClick={() => handleModuleSort('submissionRate')}
                    >
                      <div className="flex items-center justify-center gap-2">
                        Submission Rate
                        {moduleSortBy === 'submissionRate' && (moduleSortOrder === 'asc' ? <ArrowUp size={14} /> : <ArrowDown size={14} />)}
                      </div>
                    </th>
                    <th 
                      className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 cursor-pointer hover:bg-slate-100 text-center"
                      onClick={() => handleModuleSort('satisfaction')}
                    >
                      <div className="flex items-center justify-center gap-2">
                        Satisfaction
                        {moduleSortBy === 'satisfaction' && (moduleSortOrder === 'asc' ? <ArrowUp size={14} /> : <ArrowDown size={14} />)}
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredModuleKPIs.map((m: any) => (
                    <tr key={m.id} className="hover:bg-slate-50 transition-colors">
                      <td className="px-6 py-4">
                        <div className="font-medium text-slate-900">{m.lecturer}</div>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-2">
                          <span className="px-2 py-1 bg-indigo-50 text-indigo-700 rounded-lg text-[10px] font-black">L{m.level}</span>
                          <span className="text-sm text-slate-700 font-medium">{m.module}</span>
                        </div>
                      </td>
                      <td className="px-6 py-4 text-center">
                        <span className={cn(
                          "text-sm font-bold",
                          m.passRate >= 85 ? "text-emerald-600" : m.passRate >= 75 ? "text-amber-600" : "text-red-600"
                        )}>
                          {Math.round(m.passRate)}%
                        </span>
                      </td>
                      <td className="px-6 py-4 text-center">
                        <span className={cn(
                          "text-sm font-bold",
                          m.submissionRate >= 90 ? "text-emerald-600" : m.submissionRate >= 80 ? "text-amber-600" : "text-red-600"
                        )}>
                          {Math.round(m.submissionRate)}%
                        </span>
                      </td>
                      <td className="px-6 py-4 text-center">
                        <span className={cn(
                          "text-sm font-bold",
                          m.satisfaction >= 85 ? "text-emerald-600" : m.satisfaction >= 75 ? "text-amber-600" : "text-red-600"
                        )}>
                          {Math.round(m.satisfaction)}%
                        </span>
                      </td>
                    </tr>
                  ))}
                  {filteredModuleKPIs.length === 0 && (
                    <tr>
                      <td colSpan={5} className="px-6 py-8 text-center text-slate-500">
                        No modules found matching your search.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const GroupAnalysisPage = () => {
  const { records, filteredRecords, weights, loading, refreshing, filters, setFilters, refresh } = useKPI();

  if (loading || !weights) return <div className="p-8 text-slate-500">Loading group analysis...</div>;

  const groups = Array.from(new Set(records.map(r => r.group))).filter(Boolean);
  const groupStats = groups.map(grp => {
    const grpRecs = filteredRecords.filter(r => r.group === grp);
    const enrolled = grpRecs.reduce((acc, r) => acc + r.totalEnrolled, 0);
    const subs = grpRecs.reduce((acc, r) => acc + r.totalSubmissions, 0);
    const passes = grpRecs.reduce((acc, r) => acc + r.passes, 0);
    const fails = grpRecs.reduce((acc, r) => acc + r.fails, 0);
    const nonSubs = grpRecs.reduce((acc, r) => acc + r.nonSubmissions, 0);
    const att = grpRecs.length ? grpRecs.reduce((acc, r) => acc + r.attendanceRate, 0) / grpRecs.length : 0;

    return {
      name: grp,
      submissionRate: Math.round((subs / enrolled) * 100) || 0,
      passRate: Math.round((passes / subs) * 100) || 0,
      attendanceRate: Math.round(att),
      passes,
      fails,
      nonSubmissions: nonSubs
    };
  });

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Group & Delivery Mode Analysis</h1>
        <p className="text-slate-500">Performance comparison between Evening, Weekday, and Weekend groups</p>
      </div>

      <GlobalFilters records={records} filters={filters} setFilters={setFilters} refresh={refresh} loading={refreshing} />

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <h3 className="mb-6 text-lg font-semibold text-slate-900">Group vs Submission %</h3>
          <div className="h-[300px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={groupStats}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} />
                <YAxis axisLine={false} tickLine={false} />
                <Tooltip />
                <Bar dataKey="submissionRate" name="Submission %" fill="#6366f1" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <h3 className="mb-6 text-lg font-semibold text-slate-900">Pass / Fail / Non-submission by Group</h3>
          <div className="h-[300px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={groupStats}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} />
                <YAxis axisLine={false} tickLine={false} />
                <Tooltip />
                <Bar dataKey="passes" name="Passes" stackId="a" fill="#10b981" />
                <Bar dataKey="fails" name="Fails" stackId="a" fill="#ef4444" />
                <Bar dataKey="nonSubmissions" name="Non-Submissions" stackId="a" fill="#f59e0b" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
        <h3 className="mb-6 text-lg font-semibold text-slate-900">Attendance Comparison by Group</h3>
        <div className="h-[300px]">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={groupStats}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis dataKey="name" axisLine={false} tickLine={false} />
              <YAxis axisLine={false} tickLine={false} />
              <Tooltip />
              <Bar dataKey="attendanceRate" name="Attendance %" fill="#06b6d4" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
};

const RiskInterventionPage = () => {
  const { records, filteredRecords, weights, loading, refreshing, filters, setFilters, refresh } = useKPI();

  if (loading || !weights) return <div className="p-8 text-slate-500">Loading risk analysis...</div>;

  const atRiskModules = filteredRecords.map(r => {
    const atRiskRatio = (r.fails + r.nonSubmissions) / r.totalEnrolled;
    const difficultyIndex = 1 - (r.passes / r.totalSubmissions || 0);
    const engagementRisk = r.meqResponseRate < 20 && (r.passes / r.totalSubmissions) < 0.5;

    return {
      ...r,
      atRiskRatio,
      difficultyIndex,
      engagementRisk
    };
  }).sort((a, b) => b.atRiskRatio - a.atRiskRatio);

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Risk & Intervention Dashboard</h1>
        <p className="text-slate-500">Early warning system for at-risk students and modules</p>
      </div>

      <GlobalFilters records={records} filters={filters} setFilters={setFilters} refresh={refresh} loading={refreshing} />

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <KPICard title="High Risk Modules" value={atRiskModules.filter(m => m.atRiskRatio > 0.3).length} icon={AlertTriangle} color="bg-red-500" />
        <KPICard title="Avg. Difficulty Index" value={(atRiskModules.reduce((acc, m) => acc + m.difficultyIndex, 0) / atRiskModules.length).toFixed(2)} icon={TrendingUp} color="bg-amber-500" />
        <KPICard title="Engagement Alerts" value={atRiskModules.filter(m => m.engagementRisk).length} icon={AlertTriangle} color="bg-orange-500" />
      </div>

      <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
        <h3 className="mb-6 text-lg font-semibold text-slate-900">At-Risk Student Hotspots</h3>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50">
                <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500">Module</th>
                <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500">Lecturer</th>
                <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 text-center">At-Risk Ratio</th>
                <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 text-center">Difficulty</th>
                <th className="px-6 py-4 text-xs font-semibold tracking-wider uppercase text-slate-500 text-right">Intervention</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {atRiskModules.slice(0, 10).map((m) => (
                <tr key={m.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-6 py-4">
                    <div className="font-medium text-slate-900">{m.moduleName}</div>
                    <div className="text-xs text-slate-400">Level {m.level}</div>
                  </td>
                  <td className="px-6 py-4 text-sm text-slate-600">{m.lecturerName}</td>
                  <td className="px-6 py-4 text-center">
                    <div className="flex items-center justify-center gap-2">
                      <div className="w-16 h-2 bg-slate-100 rounded-full overflow-hidden">
                        <div className={cn("h-full", m.atRiskRatio > 0.4 ? "bg-red-500" : m.atRiskRatio > 0.2 ? "bg-amber-500" : "bg-emerald-500")} style={{ width: `${m.atRiskRatio * 100}%` }} />
                      </div>
                      <span className="text-xs font-bold">{Math.round(m.atRiskRatio * 100)}%</span>
                    </div>
                  </td>
                  <td className="px-6 py-4 text-center text-sm font-bold text-slate-600">{m.difficultyIndex.toFixed(2)}</td>
                  <td className="px-6 py-4 text-right">
                    <span className={cn(
                      "px-2 py-1 text-xs font-medium rounded-full",
                      m.atRiskRatio > 0.4 ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                    )}>
                      {m.atRiskRatio > 0.4 ? 'Urgent' : 'Monitor'}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const UploadPage = ({ onComplete }: { onComplete: () => void }) => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [result, setResult] = useState<any>(null);
  
  // Google Drive State
  const [driveFiles, setDriveFiles] = useState<any[]>([]);
  const [googleTokens, setGoogleTokens] = useState<any>(null);
  const [fetchingDrive, setFetchingDrive] = useState(false);

  // OneDrive State
  const [oneDriveFiles, setOneDriveFiles] = useState<any[]>([]);
  const [msTokens, setMsTokens] = useState<any>(null);
  const [fetchingOneDrive, setFetchingOneDrive] = useState(false);
  const [fileFilter, setFileFilter] = useState<'all' | 'excel' | 'csv' | 'json'>('all');

  useEffect(() => {
    const gTokens = localStorage.getItem('google_drive_tokens');
    if (gTokens) {
      try { setGoogleTokens(JSON.parse(gTokens)); } catch (e) { console.error(e); }
    }

    const mTokens = localStorage.getItem('ms_onedrive_tokens');
    if (mTokens) {
      try { setMsTokens(JSON.parse(mTokens)); } catch (e) { console.error(e); }
    }

    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'GOOGLE_AUTH_SUCCESS') {
        const tokens = event.data.tokens;
        setGoogleTokens(tokens);
        localStorage.setItem('google_drive_tokens', JSON.stringify(tokens));
      }
      if (event.data?.type === 'MS_AUTH_SUCCESS') {
        const tokens = event.data.tokens;
        setMsTokens(tokens);
        localStorage.setItem('ms_onedrive_tokens', JSON.stringify(tokens));
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  useEffect(() => {
    if (googleTokens) fetchDriveFiles();
  }, [googleTokens]);

  useEffect(() => {
    if (msTokens) fetchOneDriveFiles();
  }, [msTokens]);

  const fetchDriveFiles = async () => {
    setFetchingDrive(true);
    try {
      const res = await apiFetch('/api/drive/files', {
        headers: { 'x-google-tokens': JSON.stringify(googleTokens) }
      });
      if (res.ok) {
        const data = await res.json();
        setDriveFiles(Array.isArray(data) ? data : []);
      }
    } catch (err) {
      console.error(err);
    } finally {
      setFetchingDrive(false);
    }
  };

  const fetchOneDriveFiles = async () => {
    if (!msTokens) return;
    setFetchingOneDrive(true);
    try {
      const res = await apiFetch('/api/onedrive/files', {
        headers: { 'x-ms-tokens': JSON.stringify(msTokens) }
      });
      if (res.ok) {
        const data = await res.json();
        setOneDriveFiles(Array.isArray(data) ? data : []);
      } else if (res.status === 401 && msTokens.refresh_token) {
        // Try to refresh
        const refreshRes = await apiFetch('/api/auth/ms/refresh', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ refresh_token: msTokens.refresh_token })
        });
        if (refreshRes.ok) {
          const newTokens = await refreshRes.json();
          const updatedTokens = { ...msTokens, ...newTokens };
          setMsTokens(updatedTokens);
          localStorage.setItem('ms_onedrive_tokens', JSON.stringify(updatedTokens));
          // Retry fetch
          fetchOneDriveFiles();
        } else {
          // Refresh failed, disconnect
          handleDisconnectOneDrive();
        }
      }
    } catch (err) {
      console.error(err);
    } finally {
      setFetchingOneDrive(false);
    }
  };

  const handleConnectDrive = async () => {
    try {
      const res = await apiFetch('/api/auth/google/url');
      if (!res.ok) throw new Error('Failed to get auth URL');
      const { url } = await res.json();
      const authWindow = window.open(url, 'google_auth', 'width=600,height=700');
      if (!authWindow) {
        alert('Popup blocked. Please allow popups for this site to connect your account.');
      }
    } catch (err) {
      console.error(err);
      alert('Could not connect to Google Drive. Please ensure Client ID/Secret are configured.');
    }
  };

  const handleConnectOneDrive = async () => {
    try {
      const res = await apiFetch('/api/auth/ms/url');
      if (!res.ok) throw new Error('Failed to get auth URL');
      const { url } = await res.json();
      const authWindow = window.open(url, 'ms_auth', 'width=600,height=700');
      if (!authWindow) {
        alert('Popup blocked. Please allow popups for this site to connect your account.');
      }
    } catch (err) {
      console.error(err);
      alert('Could not connect to OneDrive. Please ensure Client ID/Secret are configured.');
    }
  };

  const handleDisconnectDrive = () => {
    setGoogleTokens(null);
    setDriveFiles([]);
    localStorage.removeItem('google_drive_tokens');
  };

  const handleDisconnectOneDrive = () => {
    setMsTokens(null);
    setOneDriveFiles([]);
    localStorage.removeItem('ms_onedrive_tokens');
  };

  const handleImportFromDrive = async (fileId: string) => {
    setUploading(true);
    try {
      const res = await apiFetch('/api/drive/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId, tokens: JSON.stringify(googleTokens) }),
      });
      const data = await res.json();
      setResult(data);
    } catch (err) {
      console.error(err);
    } finally {
      setUploading(false);
    }
  };

  const handleImportFromOneDrive = async (fileId: string) => {
    setUploading(true);
    try {
      const res = await apiFetch('/api/onedrive/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId, tokens: JSON.stringify(msTokens) }),
      });
      const data = await res.json();
      setResult(data);
    } catch (err) {
      console.error(err);
    } finally {
      setUploading(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) setFile(e.target.files[0]);
  };

  const KPI_ALIASES: Record<string, string[]> = {
  programme: ['programme', 'program', 'course', 'programme name', 'dept', 'department', 'school', 'faculty', 'class'],
  module: ['module', 'subject', 'module name', 'unit', 'code', 'module code', 'row labels', 'item', 'title'],
  lecturer: ['lecturer', 'teacher', 'instructor', 'lecturer name', 'staff', 'academic', 'tutor', 'professor', 'name', 'assigned to'],
  term: ['term', 'semester', 'period', 'academic year', 'year', 'session', 'trimester'],
  intake: ['intake', 'batch', 'cohort', 'month', 'start date', 'period'],
  level: ['level', 'lvl', 'stage', 'year of study', 'level of study'],
  group: ['group', 'delivery mode', 'mode', 'session', 'class', 'delivery group'],
  course: ['course', 'programme', 'degree', 'course name'],
  totalenrolled: ['totalenrolled', 'enrolled', 'students', 'total students', 'total enrolled', 'count', 'size', 'enrolment', 'no. of students'],
  totalsubmissions: ['totalsubmissions', 'submissions', 'total submissions', 'total submission', 'submitted', 'submission count', 'average of submission %', 'total'],
  nonsubmissions: ['nonsubmissions', 'non-submissions', 'not submitted', 'absent', 'non-submission count'],
  passes: ['passes', 'pass', 'passed', 'success', 'passed count', 'achieved'],
  fails: ['fails', 'fail', 'failed', 'failure', 'failed count', 'not achieved'],
  attendancerate: ['attendancerate', 'attendance', 'avg attendance', 'attendance rate', 'presence', 'attendance %', 'at %'],
  meqsatisfaction: ['meqsatisfaction', 'meq', 'satisfaction', 'student satisfaction', 'meq satisfaction', 'feedback', 'rating', 'score', 'satisfaction %'],
  meqresponserate: ['meq response rate', 'response rate', 'meq %', 'feedback rate', 'meq response %']
};

const handleUpload = async () => {
    if (!file) return;
    setUploading(true);
    setResult(null);
    
    try {
      const isStandardFormat = file.name.toLowerCase().endsWith('.xlsx') || 
                               file.name.toLowerCase().endsWith('.xls') || 
                               file.name.toLowerCase().endsWith('.csv') || 
                               file.name.toLowerCase().endsWith('.json');

      const data = await new Promise<Uint8Array>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
          if (!e.target?.result) reject(new Error('Failed to read file'));
          else resolve(new Uint8Array(e.target.result as ArrayBuffer));
        };
        reader.onerror = () => reject(new Error('File reading error'));
        reader.readAsArrayBuffer(file);
      });

      let json: any[];

      if (isStandardFormat) {
        if (file.name.toLowerCase().endsWith('.json')) {
          const text = new TextDecoder().decode(data);
          json = JSON.parse(text);
        } else {
          const workbook = XLSX.read(data, { type: 'array' });
          let allData: any[] = [];
          
          // Check all sheets
          for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
            if (rows.length === 0) continue;

            let headerRowIndex = -1;
            let mapping: Record<string, number> = {};
            let metadata: Record<string, any> = {};

            // 1. Scan for metadata (Key: Value) in the first 20 rows
            for (let i = 0; i < Math.min(rows.length, 20); i++) {
              const row = rows[i];
              if (!Array.isArray(row)) continue;
              row.forEach((cell, cellIdx) => {
                if (typeof cell === 'string' && cell.includes(':')) {
                  const [key, ...valParts] = cell.split(':');
                  const val = valParts.join(':').trim();
                  const cleanKey = key.trim().toLowerCase();
                  
                  Object.entries(KPI_ALIASES).forEach(([targetKey, aliases]) => {
                    if (aliases.includes(cleanKey)) {
                      metadata[targetKey] = val;
                    }
                  });
                } else if (typeof cell === 'string' && cellIdx < row.length - 1) {
                  // Check if next cell is the value
                  const nextCell = row[cellIdx + 1];
                  const cleanKey = cell.trim().toLowerCase();
                  Object.entries(KPI_ALIASES).forEach(([targetKey, aliases]) => {
                    if (aliases.includes(cleanKey) && nextCell !== undefined && nextCell !== null) {
                      metadata[targetKey] = nextCell;
                    }
                  });
                }
              });
            }

            // 2. Find the header row
            for (let i = 0; i < Math.min(rows.length, 20); i++) {
              const row = rows[i];
              if (!Array.isArray(row)) continue;
              
              const tempMapping: Record<string, number> = {};
              let matches = 0;

              Object.keys(KPI_ALIASES).forEach(targetKey => {
                const aliases = KPI_ALIASES[targetKey];
                const colIndex = row.findIndex(cell => 
                  cell && aliases.includes(String(cell).trim().toLowerCase())
                );
                if (colIndex !== -1) {
                  tempMapping[targetKey] = colIndex;
                  matches++;
                }
              });

              if (matches >= 3) { // Lowered threshold to 3 for better detection
                headerRowIndex = i;
                mapping = tempMapping;
                break;
              }
            }

            if (headerRowIndex !== -1) {
              const sheetJson = rows.slice(headerRowIndex + 1).map(row => {
                const obj: any = { ...metadata };
                Object.keys(KPI_ALIASES).forEach(key => {
                  if (mapping[key] !== undefined) {
                    obj[key] = row[mapping[key]];
                  }
                });
                // Try to find department
                const deptIndex = rows[headerRowIndex].findIndex(cell => 
                  cell && (String(cell).trim().toLowerCase() === 'department' || String(cell).trim().toLowerCase() === 'dept')
                );
                if (deptIndex !== -1) obj.department = row[deptIndex];
                return obj;
              }).filter(obj => {
                // Keep row if it has at least some data
                const values = Object.values(obj);
                return values.some(v => v !== undefined && v !== null && v !== '');
              });
              allData = [...allData, ...sheetJson];
            }
          }

          if (allData.length === 0) {
            // Final fallback: standard sheet_to_json on first sheet
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            json = XLSX.utils.sheet_to_json(firstSheet);
          } else {
            json = allData;
          }
        }
      } else {
        // Smart Extraction for non-standard formats (PDF, Images, etc.)
        if (file.size > 15 * 1024 * 1024) {
          throw new Error('File is too large for smart extraction. Please upload a file smaller than 15MB.');
        }

        const base64 = await fileToBase64(file);
        
        const ai = new GoogleGenAI({ apiKey: (import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || '') });
        
        const prompt = `
          Extract academic performance data from the provided file. 
          The data should be returned as a JSON array of objects.
          
          CRITICAL: 
          1. Extract the LECTURER name accurately. Look for columns like "Lecturer", "Staff", "Assigned To", or names in the rows.
          2. Clean the MODULE name. Remove any group suffixes like "- Group 1" or "(G1)".
          
          Each object MUST have these fields:
          - programme (string)
          - module (string)
          - lecturer (string)
          - term (string)
          - intake (string)
          - totalEnrolled (number)
          - totalSubmissions (number)
          - passes (number)
          - fails (number)
          - attendanceRate (number, 0-100)
          - meqSatisfaction (number, 0-100)
          - department (string, optional)

          If a field is missing, try to infer it or use null. 
          Return ONLY the JSON array.
        `;

        try {
          const response = await ai.models.generateContent({
            model: "gemini-3-flash-preview",
            contents: {
              parts: [
                { text: prompt },
                {
                  inlineData: {
                    data: base64,
                    mimeType: file.type || 'application/pdf'
                  }
                }
              ]
            },
            config: {
              responseMimeType: "application/json"
            }
          });

          json = JSON.parse(response.text || '[]');
        } catch (err: any) {
          if (err.message === 'Load failed' || err.message.includes('fetch')) {
            throw new Error('Network error during smart extraction. The file might be too large or your connection was interrupted.');
          }
          throw err;
        }
      }

      if (!Array.isArray(json) || json.length === 0) {
        throw new Error('The file appears to be empty or invalid.');
      }

      const requiredColumns = Object.keys(KPI_ALIASES);
      const firstRow = json[0];
      
      const columnMapping: Record<string, string> = {};
      const missingColumns: string[] = [];

      requiredColumns.forEach(targetCol => {
        const aliases = KPI_ALIASES[targetCol];
        const foundKey = Object.keys(firstRow).find(actualKey => 
          aliases.includes(actualKey.trim().toLowerCase()) || actualKey.trim().toLowerCase() === targetCol
        );
        
        if (foundKey) {
          columnMapping[targetCol] = foundKey;
        } else {
          missingColumns.push(targetCol);
        }
      });

      // Only require lecturer and module as essential
      const essential = ['lecturer', 'module'];
      const missingEssential = essential.filter(k => !columnMapping[k]);
      
      let finalData = [];

      if (missingEssential.length > 0) {
        // Try Smart Parsing with Gemini
        setUploading(true);
        const columns = Object.keys(firstRow).join(', ');
        try {
          const ai = new GoogleGenAI({ apiKey: (import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || '') });
          const sampleData = JSON.stringify(json.slice(0, 10));
          
          const prompt = `
            I have an academic spreadsheet with these columns: [${columns}].
            Here is a sample of the data: ${sampleData}.
            
            This data might be in a "wide" format where modules are column headers, or it might be a student list.
            Please transform this into a standardized KPI format (one row per module/lecturer).
            
            CRITICAL:
            1. Ensure the LECTURER name is extracted for every row. If it's in a header or a specific column, propagate it to all relevant rows.
            2. Clean the MODULE name. Remove any group suffixes (e.g., "Module A - Group 1" becomes "Module A").
            
            Return a JSON array of objects with these keys:
            - module (The name of the module/subject)
            - lecturer (The lecturer name, or "Department Staff" if not found)
            - programme (The programme/course name if available)
            - course (The course name if available)
            - level (The academic level: 4, 5, or 6)
            - group (The delivery group: Evening, Weekday, or Weekend)
            - intake (The intake/batch/cohort if available)
            - totalenrolled (Total number of students in that module)
            - totalsubmissions (Number of students who submitted/sat for exam)
            - nonsubmissions (Number of students who did not submit)
            - passes (Number of students who passed)
            - fails (Number of students who failed)
            - attendancerate (Average attendance % for that module)
            - meqsatisfaction (Average satisfaction score 0-100)
            - meqresponserate (MEQ response rate %)
            
            Only return the JSON array. No other text.
          `;

          let response;
          try {
            response = await ai.models.generateContent({
              model: "gemini-3-flash-preview",
              contents: prompt,
              config: { responseMimeType: "application/json" }
            });
          } catch (err: any) {
            if (err.message === 'Load failed' || err.message.includes('fetch')) {
              throw new Error('Network error during AI parsing. Your connection was interrupted.');
            }
            throw err;
          }

          const smartData = JSON.parse(response.text || '[]');
          if (Array.isArray(smartData) && smartData.length > 0) {
            finalData = smartData;
          } else {
            throw new Error('Smart parsing failed to extract data.');
          }
        } catch (aiErr) {
          console.error('Smart parsing error:', aiErr);
          throw new Error(`Missing essential columns: ${missingEssential.join(', ')}. Standard mapping failed and smart parsing could not resolve the format. Found columns: ${columns}`);
        }
      } else {
        finalData = json.map(row => {
          const newRow: any = {};
          requiredColumns.forEach(col => {
            if (columnMapping[col]) {
              newRow[col] = row[columnMapping[col]];
            }
          });
          const deptKey = Object.keys(row).find(k => k.trim().toLowerCase() === 'department' || k.trim().toLowerCase() === 'dept');
          if (deptKey) newRow.department = row[deptKey];
          return newRow;
        });
      }

      // Type validation and cleaning for all rows
      const cleanedData = [];
      for (let i = 0; i < finalData.length; i++) {
        const row = finalData[i];
        
        // Skip rows that are completely empty or missing both essential fields
        const hasModule = row.module && String(row.module).trim() !== '';
        const hasLecturer = row.lecturer && String(row.lecturer).trim() !== '';
        
        if (!hasModule || !hasLecturer) {
          console.warn(`Skipping row ${i + 1} due to missing essential data (Module/Lecturer)`);
          continue;
        }

        const numericFields = ['totalenrolled', 'totalsubmissions', 'nonsubmissions', 'passes', 'fails', 'attendancerate', 'meqsatisfaction', 'meqresponserate'];
        let rowValid = true;
        
        for (const field of numericFields) {
          const val = row[field];
          if (val !== undefined && val !== null && val !== '' && isNaN(Number(val))) {
            // If it's a percentage string like "85%", try to parse it
            if (typeof val === 'string' && val.includes('%')) {
              row[field] = parseFloat(val.replace('%', ''));
            } else {
              console.warn(`Invalid numeric value at row ${i + 1} for ${field}: ${val}`);
              row[field] = 0; // Default to 0 instead of crashing
            }
          }
        }
        
        cleanedData.push(row);
      }

      if (cleanedData.length === 0) {
        throw new Error('No valid data rows found. Please ensure your file has at least one row with both a Lecturer and a Module name.');
      }

      let res;
      try {
        res = await apiFetch('/api/upload/performance', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: cleanedData }),
        });
      } catch (err: any) {
        if (err.message === 'Load failed' || err.message.includes('fetch')) {
          throw new Error('Network error during upload. The server might be restarting or your connection was interrupted.');
        }
        throw err;
      }
      
      const resData = await res.json();
      if (!res.ok) throw new Error(resData.error || 'Upload failed');
      setResult(resData);
      setFile(null);
    } catch (err: any) {
      console.error(err);
      setResult({ error: err.message || 'Failed to process file' });
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-8">
      <div className="flex justify-between items-end">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Import Performance Data</h1>
          <p className="text-slate-500">Upload Excel files or link from Cloud Storage</p>
        </div>
        <div className="flex gap-2">
          {!googleTokens && (
            <button 
              onClick={handleConnectDrive}
              className="flex items-center gap-2 px-4 py-2 text-sm font-medium bg-white border border-slate-200 rounded-lg text-slate-600 hover:bg-slate-50"
            >
              <Cloud size={16} />
              Connect Google
            </button>
          )}
          {!msTokens && (
            <button 
              onClick={handleConnectOneDrive}
              className="flex items-center gap-2 px-4 py-2 text-sm font-medium bg-white border border-slate-200 rounded-lg text-slate-600 hover:bg-slate-50"
            >
              <Cloud size={16} />
              Connect OneDrive
            </button>
          )}
          {(googleTokens || msTokens) && (
            <button 
              onClick={() => { fetchDriveFiles(); fetchOneDriveFiles(); }}
              className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 hover:bg-indigo-50 rounded-lg"
            >
              <RefreshCw size={16} className={(fetchingDrive || fetchingOneDrive) ? "animate-spin" : ""} />
              Refresh All
            </button>
          )}
        </div>
      </div>

      <div className="flex items-center gap-4 p-4 bg-white border border-slate-100 rounded-2xl shadow-sm">
        <span className="text-sm font-medium text-slate-500">Filter Cloud Files:</span>
        <div className="flex gap-2">
          {(['all', 'excel', 'csv', 'json'] as const).map((type) => (
            <button
              key={type}
              onClick={() => setFileFilter(type)}
              className={cn(
                "px-3 py-1 text-xs font-medium rounded-full transition-all",
                fileFilter === type 
                  ? "bg-indigo-600 text-white" 
                  : "bg-slate-100 text-slate-600 hover:bg-slate-200"
              )}
            >
              {type.toUpperCase()}
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="p-8 border-2 border-dashed border-slate-200 rounded-3xl bg-white flex flex-col items-center justify-center text-center">
          <div className="p-4 mb-4 bg-indigo-50 text-indigo-600 rounded-2xl">
            <Upload size={32} />
          </div>
          <h3 className="text-lg font-semibold text-slate-900">Universal Upload</h3>
          <p className="mb-6 text-sm text-slate-500">Excel, CSV, JSON, PDF, or Images</p>
          
          <input 
            type="file" 
            id="file-upload" 
            className="hidden" 
            accept=".xlsx,.xls,.csv,.json,.pdf,.png,.jpg,.jpeg"
            onChange={handleFileChange}
          />
          <label 
            htmlFor="file-upload"
            className="px-6 py-2 text-sm font-semibold text-white bg-indigo-600 rounded-xl cursor-pointer hover:bg-indigo-700 transition-all"
          >
            {file ? file.name : 'Choose File'}
          </label>

          {file && (
            <button
              onClick={handleUpload}
              disabled={uploading}
              className="mt-4 text-indigo-600 text-sm font-medium hover:underline"
            >
              {uploading ? 'Processing...' : 'Start Import'}
            </button>
          )}
        </div>

        {/* Google Drive Section */}
        <div className="p-8 border border-slate-100 rounded-3xl bg-white flex flex-col">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-emerald-50 text-emerald-600 rounded-lg">
                <Cloud size={20} />
              </div>
              <h3 className="text-lg font-semibold text-slate-900">Google Drive</h3>
            </div>
            {googleTokens && (
              <button 
                onClick={handleDisconnectDrive}
                className="text-xs text-red-500 hover:underline"
              >
                Disconnect
              </button>
            )}
          </div>
          
          {!googleTokens ? (
            <div className="flex-1 flex flex-col items-center justify-center text-center py-8">
              <p className="text-sm text-slate-500 mb-4">Connect Google Drive</p>
              <button 
                onClick={handleConnectDrive}
                className="px-6 py-2 text-sm font-semibold text-indigo-600 border border-indigo-200 rounded-xl hover:bg-indigo-50"
              >
                Connect
              </button>
            </div>
          ) : (
            <div className="flex-1 overflow-y-auto max-h-[200px] space-y-2 pr-2">
              {driveFiles
                .filter(f => {
                  if (fileFilter === 'all') return true;
                  const ext = f.name.toLowerCase().split('.').pop();
                  if (fileFilter === 'excel') return ext === 'xlsx' || ext === 'xls';
                  return ext === fileFilter;
                })
                .length > 0 ? (
                driveFiles
                  .filter(f => {
                    if (fileFilter === 'all') return true;
                    const ext = f.name.toLowerCase().split('.').pop();
                    if (fileFilter === 'excel') return ext === 'xlsx' || ext === 'xls';
                    return ext === fileFilter;
                  })
                  .map(f => (
                  <div key={f.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                    <div className="overflow-hidden">
                      <p className="text-sm font-medium text-slate-900 truncate">{f.name}</p>
                      <p className="text-[10px] text-slate-400">Modified: {new Date(f.modifiedTime).toLocaleDateString()}</p>
                    </div>
                    <button 
                      onClick={() => handleImportFromDrive(f.id)}
                      disabled={uploading}
                      className="p-2 text-indigo-600 hover:bg-indigo-100 rounded-lg transition-colors"
                    >
                      <Upload size={16} />
                    </button>
                  </div>
                ))
              ) : (
                <p className="text-sm text-slate-500 text-center py-8">No files found</p>
              )}
            </div>
          )}
        </div>

        {/* OneDrive Section */}
        <div className="p-8 border border-slate-100 rounded-3xl bg-white flex flex-col">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-blue-50 text-blue-600 rounded-lg">
                <Cloud size={20} />
              </div>
              <h3 className="text-lg font-semibold text-slate-900">OneDrive</h3>
            </div>
            {msTokens && (
              <button 
                onClick={handleDisconnectOneDrive}
                className="text-xs text-red-500 hover:underline"
              >
                Disconnect
              </button>
            )}
          </div>
          
          {!msTokens ? (
            <div className="flex-1 flex flex-col items-center justify-center text-center py-8">
              <p className="text-sm text-slate-500 mb-4">Connect OneDrive</p>
              <button 
                onClick={handleConnectOneDrive}
                className="px-6 py-2 text-sm font-semibold text-indigo-600 border border-indigo-200 rounded-xl hover:bg-indigo-50"
              >
                Connect
              </button>
            </div>
          ) : (
            <div className="flex-1 overflow-y-auto max-h-[200px] space-y-2 pr-2">
              {oneDriveFiles
                .filter(f => {
                  if (fileFilter === 'all') return true;
                  const ext = f.name.toLowerCase().split('.').pop();
                  if (fileFilter === 'excel') return ext === 'xlsx' || ext === 'xls';
                  return ext === fileFilter;
                })
                .length > 0 ? (
                oneDriveFiles
                  .filter(f => {
                    if (fileFilter === 'all') return true;
                    const ext = f.name.toLowerCase().split('.').pop();
                    if (fileFilter === 'excel') return ext === 'xlsx' || ext === 'xls';
                    return ext === fileFilter;
                  })
                  .map(f => (
                  <div key={f.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                    <div className="overflow-hidden">
                      <p className="text-sm font-medium text-slate-900 truncate">{f.name}</p>
                      <p className="text-[10px] text-slate-400">Modified: {new Date(f.lastModifiedDateTime).toLocaleDateString()}</p>
                    </div>
                    <button 
                      onClick={() => handleImportFromOneDrive(f.id)}
                      disabled={uploading}
                      className="p-2 text-indigo-600 hover:bg-indigo-100 rounded-lg transition-colors"
                    >
                      <Upload size={16} />
                    </button>
                  </div>
                ))
              ) : (
                <p className="text-sm text-slate-500 text-center py-8">No files found</p>
              )}
            </div>
          )}
        </div>
      </div>

      {result && (
        <motion.div 
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className={cn(
            "p-6 border rounded-2xl flex items-center gap-4",
            result.error ? "bg-red-50 border-red-100" : "bg-emerald-50 border-emerald-100"
          )}
        >
          <div className={cn(
            "p-2 text-white rounded-lg",
            result.error ? "bg-red-500" : "bg-emerald-500"
          )}>
            {result.error ? <AlertTriangle size={20} /> : <CheckCircle2 size={20} />}
          </div>
          <div>
            <h4 className={cn(
              "font-semibold",
              result.error ? "text-red-900" : "text-emerald-900"
            )}>
              {result.error ? "Import Failed" : "Import Successful"}
            </h4>
            <p className={cn(
              "text-sm",
              result.error ? "text-red-700" : "text-emerald-700"
            )}>
              {result.error || `Successfully imported ${result.count} performance records.`}
            </p>
          </div>
          {!result.error && (
            <button 
              onClick={onComplete}
              className="ml-auto px-4 py-2 text-sm font-medium bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors"
            >
              View Analytics
            </button>
          )}
        </motion.div>
      )}

      <div className="p-6 bg-slate-50 rounded-2xl border border-slate-100">
        <h4 className="font-semibold text-slate-900 mb-4">Required Column Headers:</h4>
        <div className="grid grid-cols-2 md:grid-cols-3 gap-4 text-sm text-slate-600">
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> lecturer</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> module</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> programme</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> term</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> intake</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> totalEnrolled</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> totalSubmissions</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> passes</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> fails</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> attendanceRate</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> meqSatisfaction</div>
          <div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-indigo-400" /> department (opt)</div>
        </div>
      </div>
    </div>
  );
};

const SettingsPage = () => {
  const [weights, setWeights] = useState<KPIWeights>({
    submissionWeight: 0.25,
    passWeight: 0.25,
    attendanceWeight: 0.25,
    meqWeight: 0.25
  });
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    apiFetch('/api/kpi/weights')
      .then(res => {
        if (!res.ok) throw new Error('Failed to fetch weights');
        return res.json();
      })
      .then(setWeights)
      .catch(console.error);
  }, []);

  const handleSave = async () => {
    setSaving(true);
    try {
      await apiFetch('/api/kpi/weights', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(weights),
      });
      alert('Weights updated successfully');
    } finally {
      setSaving(false);
    }
  };

  const totalWeight: number = (Object.values(weights) as any[]).reduce((a: number, b: number) => a + (Number(b) || 0), 0);

  return (
    <div className="max-w-2xl space-y-8">
      <div>
        <h1 className="text-2xl font-bold text-slate-900">KPI Weight Configuration</h1>
        <p className="text-slate-500">Adjust the importance of each metric in the overall score calculation</p>
      </div>

      <div className="p-8 bg-white border border-slate-100 rounded-3xl shadow-sm space-y-6">
        {Object.entries(weights).map(([key, value]) => (
          <div key={key}>
            <div className="flex justify-between mb-2">
              <label className="text-sm font-medium text-slate-700 capitalize">
                {key.replace('Weight', '')} Weight
              </label>
              <span className="text-sm font-bold text-indigo-600">{Math.round((value as number) * 100)}%</span>
            </div>
            <input
              type="range"
              min="0"
              max="1"
              step="0.05"
              value={value as number}
              onChange={(e) => setWeights({ ...weights, [key]: parseFloat(e.target.value) })}
              className="w-full h-2 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-indigo-600"
            />
          </div>
        ))}

        <div className="pt-6 border-t border-slate-100">
          <div className="flex justify-between items-center mb-6">
            <span className="text-sm font-medium text-slate-500">Total Weight</span>
            <span className={cn(
              "text-lg font-bold",
              Math.abs(totalWeight - 1) < 0.01 ? "text-emerald-600" : "text-red-600"
            )}>
              {Math.round(totalWeight * 100)}%
            </span>
          </div>
          <button
            onClick={handleSave}
            disabled={saving || Math.abs(totalWeight - 1) > 0.01}
            className="w-full py-3 font-semibold text-white bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:opacity-50 transition-all"
          >
            {saving ? 'Saving...' : 'Save Configuration'}
          </button>
          {Math.abs(totalWeight - 1) > 0.01 && (
            <p className="mt-2 text-xs text-center text-red-500">Total weight must equal 100%</p>
          )}
        </div>
      </div>
    </div>
  );
};

const AIInsightsPage = () => {
  const [report, setReport] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const generateReport = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await apiFetch('/api/kpi/raw-data');
      if (!res.ok) throw new Error('Failed to fetch raw data');
      const records = await res.json();

      if (!records || records.length === 0) {
        throw new Error('No data available to analyze. Please import data first.');
      }

      const header = "Lecturer,Module,Programme,Term,Intake,Enrolled,Submissions,Passes,Fails,AttendanceRate,MEQSatisfaction\n";
      const csvData = records.map((r: any) => 
        `"${r.lecturerName}","${r.moduleName}","${r.programmeName}","${r.term}","${r.intake}",${r.totalEnrolled},${r.totalSubmissions},${r.passes},${r.fails},${r.attendanceRate},${r.meqSatisfaction}`
      ).join('\n');

      const fullCsv = header + csvData;

      const ai = new GoogleGenAI({ apiKey: (import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || '') });
      
      const prompt = `
You are a Senior Academic Strategy Consultant and Data Scientist. I am providing you with academic performance data.

Please provide a highly concise, analytical, and critical strategic report evaluating the institution against the official BM/BME Programme KPI Framework.

Official KPI Targets to evaluate against:
- Student-Focused B3 Metrics: >90% submission rate, >75% first-time pass rate, >85% module pass rate, >80% continuation (UG Y1), >75% completion (UG), >60% MEQ response rate, >85% satisfaction.
- Curriculum Development (OfS B4): 80% asynchronous engagement, AI ethical use, ESG/DEI themes, live projects.
- Quality Assurance: <1% academic misconduct, timely feedback/marking, UCA IV compliance.

Data Structure:
- Level-Wise Analytics (Levels 4, 5, and 6)
- Module-Wise Analytics
- Lecturer-Wise Analytics

Your report MUST include:
1. Critical Performance Audit against KPIs: Identify the highest and lowest performing levels and modules based explicitly on the KPI targets above. Be critical—analyze the root causes of underperformance (e.g., is Level 4 attendance impacting Level 5 pass rates?).
2. Predictive Analytics: Based on current trends (Submission Rates vs. Pass Rates), predict which modules are at risk of failing to meet quality standards in the next academic cycle.
3. Strategic Resource Allocation & Professional Development: Recommend where to reallocate teaching resources, provide intensive pedagogical support, or assign mentoring/CPD targets.
4. Executive Summary: 3 concise, high-level strategic actions for the Academic Board.

Format: Use sharp markdown headings, data-driven tables, and bullet points. Avoid fluff. Be direct and analytical.

Here is the CSV data:
${fullCsv}
`;

      const response = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: prompt,
      });

      setReport(response.text);
    } catch (err: any) {
      setError(err.message || 'Failed to generate report.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto space-y-8">
      <div className="flex items-end justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 flex items-center gap-2">
            <Sparkles className="text-indigo-600" />
            AI Analytics Report
          </h1>
          <p className="text-slate-500">Generate comprehensive insights using Gemini AI</p>
        </div>
        <button
          onClick={generateReport}
          disabled={loading}
          className="flex items-center gap-2 px-6 py-3 text-sm font-semibold text-white transition-all bg-indigo-600 rounded-xl hover:bg-indigo-700 disabled:opacity-50"
        >
          {loading ? (
            <>
              <RefreshCw size={18} className="animate-spin" />
              Analyzing Data...
            </>
          ) : (
            <>
              <Sparkles size={18} />
              Generate Report
            </>
          )}
        </button>
      </div>

      {error && (
        <div className="p-4 bg-red-50 text-red-700 rounded-xl border border-red-100">
          {error}
        </div>
      )}

      {report && (
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="p-8 bg-white border border-slate-100 rounded-3xl shadow-sm prose prose-slate max-w-none"
        >
          <div className="markdown-body">
            <Markdown>{report}</Markdown>
          </div>
        </motion.div>
      )}

      {!report && !loading && !error && (
        <div className="flex flex-col items-center justify-center p-12 text-center bg-white border border-slate-100 rounded-3xl border-dashed">
          <div className="p-4 mb-4 bg-indigo-50 rounded-2xl">
            <Sparkles size={32} className="text-indigo-600" />
          </div>
          <h3 className="text-lg font-semibold text-slate-900">Ready to Analyze</h3>
          <p className="max-w-md mt-2 text-sm text-slate-500">
            Click the button above to generate a comprehensive analytics report covering Level-Wise, Module-Wise, and Lecturer-Wise performance metrics.
          </p>
        </div>
      )}
    </div>
  );
};

// --- Main App ---

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [activeTab, setActiveTab] = useState('executive');
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    apiFetch('/api/auth/me')
      .then(async res => {
        if (!res.ok) return;
        const data = await res.json();
        if (data.user) setUser(data.user);
      })
      .catch(console.error)
      .finally(() => setLoading(false));
  }, []);

  const handleLogout = async () => {
    await apiFetch('/api/auth/logout', { method: 'POST' });
    localStorage.removeItem('token');
    setUser(null);
  };

  if (loading) return <div className="flex items-center justify-center min-h-screen">Loading...</div>;

  if (!user) return <LoginPage onLogin={setUser} />;

  return (
    <div id="app-root-container" className="flex min-h-screen bg-slate-50 font-sans">
      {/* Sidebar */}
      <aside id="app-sidebar" className="fixed inset-y-0 left-0 w-64 p-6 bg-white border-r border-slate-100">
        <div id="sidebar-logo" className="flex items-center gap-3 mb-10 px-2">
          <div className="p-2 bg-indigo-600 rounded-xl">
            <LayoutDashboard size={24} className="text-white" />
          </div>
          <span className="text-xl font-bold text-slate-900 tracking-tight">KPI Portal</span>
        </div>

        <nav id="sidebar-nav" className="space-y-1">
          <SidebarItem 
            id="nav-executive"
            icon={LayoutDashboard} 
            label="Executive Overview" 
            active={activeTab === 'executive'} 
            onClick={() => setActiveTab('executive')} 
          />
          <SidebarItem 
            id="nav-level"
            icon={TrendingUp} 
            label="Level Analysis" 
            active={activeTab === 'level'} 
            onClick={() => setActiveTab('level')} 
          />
          <SidebarItem 
            id="nav-lecturers"
            icon={Users} 
            label="Lecturer Performance" 
            active={activeTab === 'lecturers'} 
            onClick={() => setActiveTab('lecturers')} 
          />
          <SidebarItem 
            id="nav-group"
            icon={BarChart3} 
            label="Group Analysis" 
            active={activeTab === 'group'} 
            onClick={() => setActiveTab('group')} 
          />
          <SidebarItem 
            id="nav-risk"
            icon={AlertTriangle} 
            label="Risk & Intervention" 
            active={activeTab === 'risk'} 
            onClick={() => setActiveTab('risk')} 
          />
          {user.role !== 'LECTURER' && (
            <SidebarItem 
              id="nav-upload"
              icon={Upload} 
              label="Import Data" 
              active={activeTab === 'upload'} 
              onClick={() => setActiveTab('upload')} 
            />
          )}
          <SidebarItem 
            id="nav-ai-insights"
            icon={Sparkles} 
            label="AI Insights" 
            active={activeTab === 'ai-insights'} 
            onClick={() => setActiveTab('ai-insights')} 
          />
          <SidebarItem 
            id="nav-kpi-framework"
            icon={CheckCircle2} 
            label="KPI Framework" 
            active={activeTab === 'kpi-framework'} 
            onClick={() => setActiveTab('kpi-framework')} 
          />
          {user.role === 'ADMIN' && (
            <SidebarItem 
              id="nav-settings"
              icon={Settings} 
              label="Settings" 
              active={activeTab === 'settings'} 
              onClick={() => setActiveTab('settings')} 
            />
          )}
        </nav>

        <div id="sidebar-footer" className="absolute bottom-6 left-6 right-6">
          <div id="user-profile-card" className="p-4 mb-4 bg-slate-50 rounded-2xl border border-slate-100">
            <div className="flex items-center gap-3">
              <div id="user-avatar" className="w-10 h-10 rounded-full bg-indigo-100 flex items-center justify-center text-indigo-600 font-bold">
                {user.name.charAt(0)}
              </div>
              <div id="user-info" className="overflow-hidden">
                <p className="text-sm font-semibold text-slate-900 truncate">{user.name}</p>
                <p className="text-xs text-slate-500 capitalize">{user.role.toLowerCase()}</p>
              </div>
            </div>
          </div>
          <button 
            id="logout-btn"
            onClick={handleLogout}
            className="flex items-center w-full gap-3 px-4 py-3 text-sm font-medium text-red-500 transition-colors rounded-lg hover:bg-red-50"
          >
            <LogOut size={20} />
            Logout
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main id="main-content-area" className="flex-1 ml-64 p-10">
        <AnimatePresence mode="wait">
          <motion.div
            key={activeTab}
            initial={{ opacity: 0, x: 10 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -10 }}
            transition={{ duration: 0.2 }}
          >
            {activeTab === 'executive' && <ExecutiveOverview user={user} />}
            {activeTab === 'level' && <LevelAnalysisPage />}
            {activeTab === 'lecturers' && <LecturerPerformancePage />}
            {activeTab === 'group' && <GroupAnalysisPage />}
            {activeTab === 'risk' && <RiskInterventionPage />}
            {activeTab === 'upload' && <UploadPage onComplete={() => setActiveTab('executive')} />}
            {activeTab === 'ai-insights' && <AIInsightsPage />}
            {activeTab === 'kpi-framework' && <KPIFrameworkPage />}
            {activeTab === 'settings' && <SettingsPage />}
          </motion.div>
        </AnimatePresence>
      </main>
    </div>
  );
}
