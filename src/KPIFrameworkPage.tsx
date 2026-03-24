import React from 'react';
import { CheckCircle2, BookOpen, ShieldCheck, Users, GraduationCap } from 'lucide-react';

const KPIFrameworkPage = () => {
  return (
    <div className="space-y-8 max-w-5xl">
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Lecturer KPI Framework</h1>
        <p className="text-slate-500">Official performance parameters and expectations for BM/BME Programme</p>
      </div>

      <div className="grid grid-cols-1 gap-6">
        {/* B3 Metrics */}
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-blue-100 text-blue-600 rounded-lg">
              <CheckCircle2 size={20} />
            </div>
            <h2 className="text-lg font-bold text-slate-900">Student-Focused B3 Metrics</h2>
          </div>
          <ul className="space-y-3 text-slate-600">
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Achieve &gt;90% student submission rate.</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Maintain first-time pass rate &gt;75%.</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Maintain &gt;85% module pass rate.</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Support continuation rates &gt;80% (UG Y1) and &gt;75% (OUG Y1).</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Contribute to retention rates &gt;85% (Years 2–3), &gt;90% (Years 3–4).</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Support completion rates &gt;75% (UG), &gt;65% (OUG).</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Ensure student progression &gt;60% (UG), &gt;45% (OUG).</li>
            <li className="flex items-start gap-2"><span className="text-blue-500 mt-1">•</span> Achieve &gt;60% MEQ response rate and &gt;85% satisfaction at each level.</li>
          </ul>
        </div>

        {/* Curriculum Development */}
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-indigo-100 text-indigo-600 rounded-lg">
              <BookOpen size={20} />
            </div>
            <h2 className="text-lg font-bold text-slate-900">Curriculum Development & Enhancement (OfS B4)</h2>
          </div>
          <ul className="space-y-3 text-slate-600">
            <li className="flex items-start gap-2"><span className="text-indigo-500 mt-1">•</span> Contribute in review of the assigned module contents, assignment brief and asynchronous contents.</li>
            <li className="flex items-start gap-2"><span className="text-indigo-500 mt-1">•</span> Maintain at least 80% student engagement in asynchronous learning.</li>
            <li className="flex items-start gap-2"><span className="text-indigo-500 mt-1">•</span> Introduce AI and its ethical use in the class with the module taught/led.</li>
            <li className="flex items-start gap-2"><span className="text-indigo-500 mt-1">•</span> Introduce and implement ESG, DEI and sustainability themes in the module.</li>
            <li className="flex items-start gap-2"><span className="text-indigo-500 mt-1">•</span> Introduce and implement a live project in the module lead (Live briefs, live experiential events, guest speakers, career events).</li>
          </ul>
        </div>

        {/* Quality Assurance */}
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-emerald-100 text-emerald-600 rounded-lg">
              <ShieldCheck size={20} />
            </div>
            <h2 className="text-lg font-bold text-slate-900">Quality Assurance, Academic Rigour & Integrity</h2>
          </div>
          <ul className="space-y-3 text-slate-600">
            <li className="flex items-start gap-2"><span className="text-emerald-500 mt-1">•</span> Enhance awareness on academic integrity and minimise the number of academic misconduct cases to less than 1% of the population.</li>
            <li className="flex items-start gap-2"><span className="text-emerald-500 mt-1">•</span> Ensure professional conduct and adherence to classroom schedule.</li>
            <li className="flex items-start gap-2"><span className="text-emerald-500 mt-1">•</span> Ensure comprehensive feedback is provided on formative and summative submissions on time.</li>
            <li className="flex items-start gap-2"><span className="text-emerald-500 mt-1">•</span> Ensure marking is on time and within deadline.</li>
            <li className="flex items-start gap-2"><span className="text-emerald-500 mt-1">•</span> Ensure compliance with UCA IV processes on time and within stipulated quality.</li>
            <li className="flex items-start gap-2"><span className="text-emerald-500 mt-1">•</span> To reduce the number of academic misconduct cases by 50% compared to previous cohort/intake.</li>
          </ul>
        </div>

        {/* People Management */}
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-amber-100 text-amber-600 rounded-lg">
              <Users size={20} />
            </div>
            <h2 className="text-lg font-bold text-slate-900">People Management, Administrative & Operational</h2>
          </div>
          <ul className="space-y-3 text-slate-600">
            <li className="flex items-start gap-2"><span className="text-amber-500 mt-1">•</span> Become a Buddy and Mentor at least one new colleague.</li>
            <li className="flex items-start gap-2"><span className="text-amber-500 mt-1">•</span> To contribute equally in the employability events and exhibitions etc.</li>
            <li className="flex items-start gap-2"><span className="text-amber-500 mt-1">•</span> Assist CD in operational matters including data management, analysis, reports, internal exam boards, townhall meetings, events and inductions based on business needs.</li>
            <li className="flex items-start gap-2"><span className="text-amber-500 mt-1">•</span> Maintain appropriate covers as requested.</li>
          </ul>
        </div>

        {/* Professional Development */}
        <div className="p-6 bg-white border border-slate-100 rounded-2xl shadow-sm">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-purple-100 text-purple-600 rounded-lg">
              <GraduationCap size={20} />
            </div>
            <h2 className="text-lg font-bold text-slate-900">Professional Development</h2>
          </div>
          <ul className="space-y-3 text-slate-600">
            <li className="flex items-start gap-2"><span className="text-purple-500 mt-1">•</span> To ensure at least 6 CPDS are performed throughout the year.</li>
            <li className="flex items-start gap-2"><span className="text-purple-500 mt-1">•</span> To mentor new colleagues on internal processes, best practices and classroom management methods.</li>
            <li className="flex items-start gap-2"><span className="text-purple-500 mt-1">•</span> To apply and get approved for Advance HE fellowship and/or subject relevant accreditation or professional certifications.</li>
          </ul>
        </div>
      </div>
    </div>
  );
};

export default KPIFrameworkPage;
