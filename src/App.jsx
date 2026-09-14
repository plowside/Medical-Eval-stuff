import React, { useState, useEffect } from 'react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword, updateProfile, signOut, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, collection, onSnapshot, writeBatch, doc, updateDoc } from 'firebase/firestore';
import { Activity, Globe, LogOut, CheckCircle, Clock, AlertCircle, Shield, Upload, Eye, Lock, UserPlus, Search, Download, RotateCcw } from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
  measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const translations = {
  en: {
    title: 'Medical Eval System',
    langSwitch: 'عربي',
    logout: 'Sign Out',
    loginTitle: 'System Login',
    registerTitle: 'Trainee Registration',
    email: 'Email Address',
    password: 'Password',
    name: 'Full Name',
    authBtn: 'Authenticate',
    regBtn: 'Create Account',
    switchToReg: 'Need an account? Register',
    switchToLogin: 'Already have an account? Login',
    uploadTitle: 'Upload Schedule',
    excelDesc: 'Import structural Excel file (.xlsx)',
    selectBtn: 'Select File',
    exportBtn: 'Export Report',
    tracker: 'Evaluation Tracker',
    search: 'Search by name or email...',
    evaluator: 'Evaluator',
    evaluatee: 'Evaluatee',
    status: 'Status',
    deadline: 'Deadline',
    actions: 'Actions',
    total: 'Total',
    completed: 'Completed',
    pending: 'Pending',
    missed: 'Missed',
    view: 'View',
    submitted: 'Evaluation Submitted',
    noTasks: 'No active evaluations assigned to your email.',
    unlockBtn: 'Unlock & Reset',
    dir: 'ltr',
    qsConsultant: { q1: 'Q1: Reports & Results', q2: 'Q2: Emergencies & Communication', q3: 'Q3: Knowledge & Productivity' },
    qsTrainee: { q1: 'Q1: Mentorship & Guidance', q2: 'Q2: Communication & Respect', q3: 'Q3: Clinical Teaching Quality' }
  },
  ar: {
    title: 'نظام التقييم الطبي',
    langSwitch: 'English',
    logout: 'خروج',
    loginTitle: 'تسجيل الدخول',
    registerTitle: 'تسجيل المتدربين',
    email: 'البريد الإلكتروني',
    password: 'كلمة المرور',
    name: 'الاسم الكامل',
    authBtn: 'دخول',
    regBtn: 'إنشاء حساب',
    switchToReg: 'ليس لديك حساب؟ سجل هنا',
    switchToLogin: 'لديك حساب بالفعل؟ تسجيل الدخول',
    uploadTitle: 'تحميل الجدول',
    excelDesc: 'استيراد ملف Excel',
    selectBtn: 'تحديد ملف',
    exportBtn: 'تصدير التقرير',
    tracker: 'متتبع التقييم',
    search: 'البحث بالاسم أو البريد...',
    evaluator: 'المُقيِّم',
    evaluatee: 'المُقيَّم',
    status: 'الحالة',
    deadline: 'الموعد',
    actions: 'إجراءات',
    total: 'الإجمالي',
    completed: 'مكتمل',
    pending: 'انتظار',
    missed: 'مفقود',
    view: 'عرض',
    submitted: 'تم الإرسال',
    noTasks: 'لا توجد تقييمات نشطة مخصصة لبريدك.',
    unlockBtn: 'إلغاء القفل وإعادة الضبط',
    dir: 'rtl',
    qsConsultant: { q1: 'س1: التقارير والنتائج', q2: 'س2: حالات الطوارئ والتواصل', q3: 'س3: المعرفة والإنتاجية' },
    qsTrainee: { q1: 'س1: التوجيه والإرشاد', q2: 'س2: التواصل والاحترام', q3: 'س3: جودة التدريب السريري' }
  }
};

const StatusBadge = ({ status }) => {
  const styles = { Completed: 'bg-green-100 text-green-700', Pending: 'bg-yellow-100 text-yellow-700', Missed: 'bg-red-100 text-red-700' };
  const icons = { Completed: <CheckCircle size={14} className="mx-1 shrink-0" />, Pending: <Clock size={14} className="mx-1 shrink-0" />, Missed: <AlertCircle size={14} className="mx-1 shrink-0" /> };
  return (
    <span className={`inline-flex items-center px-2 py-1 rounded-full text-[10px] sm:text-xs font-medium whitespace-nowrap ${styles[status]}`}>
      {icons[status]} <span className="truncate">{status}</span>
    </span>
  );
};

const ScoreSelector = ({ value, onChange }) => (
  <div className="flex gap-2 sm:gap-3 w-full sm:w-auto mt-2 sm:mt-0 justify-between sm:justify-start">
    {[1, 2, 3, 4].map(v => (
      <button
        key={v}
        onClick={() => onChange(v)}
        className={`w-10 h-10 sm:w-12 sm:h-12 rounded-xl text-sm sm:text-base font-bold transition-all shadow-sm ${value === v ? 'bg-blue-600 text-white ring-2 ring-blue-600 ring-offset-1' : 'bg-gray-50 text-gray-600 border border-gray-200 hover:bg-gray-100'}`}
      >
        {v}
      </button>
    ))}
  </div>
);

const ScoreViewer = ({ value }) => (
  <div className="flex gap-2 sm:gap-3 w-full sm:w-auto mt-2 sm:mt-0 justify-between sm:justify-start pointer-events-none">
    {[1, 2, 3, 4].map(v => (
      <div key={v} className={`w-10 h-10 sm:w-12 sm:h-12 rounded-xl flex items-center justify-center text-sm sm:text-base font-bold ${value === v ? 'bg-blue-600 text-white shadow-sm' : 'bg-gray-50 text-gray-300 border border-gray-100'}`}>
        {v}
      </div>
    ))}
  </div>
);

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginError, setLoginError] = useState('');
  const [isRegistering, setIsRegistering] = useState(false);
  const [showModal, setShowModal] = useState(false);
  const [viewingResult, setViewingResult] = useState(null);
  const [activeTask, setActiveTask] = useState(null);
  const [lang, setLang] = useState('en');
  const [scores, setScores] = useState({ q1: null, q2: null, q3: null });
  const [tasks, setTasks] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');

  const t = translations[lang];
  const toggleLang = () => setLang(prev => prev === 'en' ? 'ar' : 'en');

  useEffect(() => {
    document.title = t.title;
  }, [t.title]);

  useEffect(() => {
    const unsubscribeAuth = onAuthStateChanged(auth, (user) => {
      if (user) {
        const email = user.email.toLowerCase() || '';
        const name = user.displayName || email.split('@')[0];
        let role = 'user';
        let title = 'Staff';

        if (email.startsWith('admin')) {
          role = 'admin';
          title = 'Admin Office';
        }

        setCurrentUser({ uid: user.uid, email, role, name, title });
      } else {
        setCurrentUser(null);
      }
      setAuthLoading(false);
    });
    return () => unsubscribeAuth();
  }, []);

  useEffect(() => {
    if (!currentUser) return;
    const unsubscribeDb = onSnapshot(collection(db, 'tasks'), (snapshot) => {
      const fetchedTasks = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const now = new Date();
      
      const updatedTasks = fetchedTasks.map(task => {
        if (task.status === 'Pending' && new Date(task.deadlineRaw) < now) {
          return { ...task, status: 'Missed' };
        }
        return task;
      });
      
      updatedTasks.sort((a, b) => new Date(a.deadlineRaw) - new Date(b.deadlineRaw));
      setTasks(updatedTasks);
    }, (error) => {
      console.error(error);
    });
    return () => unsubscribeDb();
  }, [currentUser]);

  const handleAuth = async (e) => {
    e.preventDefault();
    setLoginError('');
    const email = e.target.email.value.trim().toLowerCase();
    const password = e.target.password.value;
    
    try {
      if (isRegistering) {
        const name = e.target.name.value.trim();
        const res = await createUserWithEmailAndPassword(auth, email, password);
        await updateProfile(res.user, { displayName: name });
        setCurrentUser(prev => ({ ...prev, name }));
      } else {
        await signInWithEmailAndPassword(auth, email, password);
      }
    } catch (error) {
      setLoginError(error.message);
    }
  };

  const handleLogout = () => signOut(auth);

  const parseExcelDate = (dateStr) => {
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return new Date();
      return d;
    } catch {
      return new Date();
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        const batch = writeBatch(db);
        let currentConsultant = null;
        let currentDateStr = null;

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length < 5) continue;
          if (row[0]) currentDateStr = row[0].toString().trim();
          
          const role = row[2]?.toString().trim();
          const name = row[3]?.toString().trim();
          const email = row[4]?.toString().trim().toLowerCase();
          
          if (!role || !name || !email) continue;

          let baseDate = currentDateStr ? parseExcelDate(currentDateStr) : new Date();
          baseDate.setHours(23, 59, 59, 999);
          
          const deadlineObj = new Date(baseDate.getTime() + 48 * 60 * 60 * 1000);
          const deadlineStr = deadlineObj.toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: 'numeric' });
          const deadlineRaw = deadlineObj.toISOString();

          if (role === 'Consultant') {
            currentConsultant = { name, email };
          } else if (role === 'Trainee' && currentConsultant) {
            const taskKey1 = `${currentConsultant.email}_${email}_${deadlineRaw}`;
            const existingTask1 = tasks.find(t => t.taskKey === taskKey1);
            
            if (!existingTask1) {
              const tRef1 = doc(collection(db, 'tasks'));
              batch.set(tRef1, {
                evaluator: currentConsultant.name,
                evaluatorEmail: currentConsultant.email,
                evaluatee: name,
                evaluateeEmail: email,
                status: 'Pending',
                deadline: deadlineStr,
                deadlineRaw,
                type: 'consultant_to_trainee',
                taskKey: taskKey1
              });
            }

            const taskKey2 = `${email}_${currentConsultant.email}_${deadlineRaw}`;
            const existingTask2 = tasks.find(t => t.taskKey === taskKey2);
            
            if (!existingTask2) {
              const tRef2 = doc(collection(db, 'tasks'));
              batch.set(tRef2, {
                evaluator: 'Anonymous',
                evaluatorEmail: email,
                evaluatee: currentConsultant.name,
                evaluateeEmail: currentConsultant.email,
                status: 'Pending',
                deadline: deadlineStr,
                deadlineRaw,
                type: 'trainee_to_consultant',
                taskKey: taskKey2
              });
            }
          }
        }
        await batch.commit();
      } catch (error) {
        console.error(error);
        alert('Upload failed.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = null;
  };

  const submitEval = async () => {
    if (!activeTask || !scores.q1 || !scores.q2 || !scores.q3) return;
    const totalRaw = scores.q1 + scores.q2 + scores.q3;
    const avgScore = Math.round((totalRaw / 12) * 100);
    const taskRef = doc(db, 'tasks', activeTask.id);
    try {
      await updateDoc(taskRef, { status: 'Completed', score: avgScore, scoresDetail: scores });
      setShowModal(false);
      setActiveTask(null);
      setScores({ q1: null, q2: null, q3: null });
    } catch (e) {
      console.error(e);
    }
  };

  const resetEval = async (taskId) => {
    if (!window.confirm('Are you sure you want to reset this evaluation? All scores will be deleted and it will be returned to the user as Pending.')) return;
    try {
      await updateDoc(doc(db, 'tasks', taskId), {
        status: 'Pending',
        score: null,
        scoresDetail: null
      });
      setViewingResult(null);
    } catch (e) {
      console.error(e);
    }
  };

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Evaluation Report');

    sheet.columns = [
      { header: 'Evaluator', key: 'evaluator', width: 30 },
      { header: 'Evaluatee', key: 'evaluatee', width: 30 },
      { header: 'Status', key: 'status', width: 15 },
      { header: 'Deadline', key: 'deadline', width: 25 },
      { header: 'Overall Score %', key: 'score', width: 18 },
      { header: 'Q1 (1-4)', key: 'q1', width: 12 },
      { header: 'Q2 (1-4)', key: 'q2', width: 12 },
      { header: 'Q3 (1-4)', key: 'q3', width: 12 }
    ];

    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
    sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    tasks.forEach(t => {
      const row = sheet.addRow({
        evaluator: t.type === 'trainee_to_consultant' ? `${t.evaluatorEmail.split('@')[0]} (Anon)` : t.evaluator,
        evaluatee: t.evaluatee,
        status: t.status,
        deadline: t.deadline,
        score: t.score !== undefined && t.score !== null ? `${t.score}%` : 'N/A',
        q1: t.scoresDetail?.q1 !== undefined && t.scoresDetail?.q1 !== null ? t.scoresDetail.q1 : 'N/A',
        q2: t.scoresDetail?.q2 !== undefined && t.scoresDetail?.q2 !== null ? t.scoresDetail.q2 : 'N/A',
        q3: t.scoresDetail?.q3 !== undefined && t.scoresDetail?.q3 !== null ? t.scoresDetail.q3 : 'N/A'
      });

      const statusCell = row.getCell('status');
      statusCell.font = { bold: true };
      if (t.status === 'Completed') {
        statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBF7D0' } };
        statusCell.font = { color: { argb: 'FF166534' } };
      } else if (t.status === 'Pending') {
        statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF08A' } };
        statusCell.font = { color: { argb: 'FF854D0E' } };
      } else if (t.status === 'Missed') {
        statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFECACA' } };
        statusCell.font = { color: { argb: 'FF991B1B' } };
      }

      if (t.type === 'consultant_to_trainee') {
        row.getCell('evaluator').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
        row.getCell('evaluatee').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
      } else {
        row.getCell('evaluator').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } }; 
        row.getCell('evaluatee').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } }; 
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), "Evaluation_Report.xlsx");
  };

  const renderTopBar = () => (
    <div className="flex flex-col sm:flex-row justify-between items-center p-4 border-b border-gray-200 bg-white shadow-sm gap-4" dir={t.dir}>
      <div className="flex items-center gap-3 w-full sm:w-auto justify-center sm:justify-start">
        <div className="bg-blue-600 text-white p-2 rounded-lg shrink-0"><Activity size={20} /></div>
        <div className="text-center sm:text-left min-w-0">
          <h1 className="font-bold text-gray-900 text-base sm:text-lg truncate">{t.title}</h1>
          <p className="text-xs text-gray-500 truncate">{currentUser.title} • {currentUser.name}</p>
        </div>
      </div>
      <div className="flex items-center justify-center gap-4 w-full sm:w-auto border-t sm:border-t-0 pt-3 sm:pt-0 border-gray-100">
        <button onClick={toggleLang} className="flex items-center justify-center gap-2 text-gray-600 hover:text-gray-900 text-sm font-medium w-full sm:w-auto">
          <Globe size={16} /><span>{t.langSwitch}</span>
        </button>
        <button onClick={handleLogout} className="flex items-center justify-center gap-2 text-gray-600 hover:text-red-600 border-l border-gray-200 pl-4 text-sm font-medium transition-colors w-full sm:w-auto">
          <LogOut size={16} /><span>{t.logout}</span>
        </button>
      </div>
    </div>
  );

  if (authLoading) return <div className="min-h-screen flex items-center justify-center font-bold text-gray-500">Loading...</div>;

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4" dir={t.dir}>
        <div className="max-w-md w-full bg-white rounded-3xl shadow-xl overflow-hidden">
          <div className="bg-blue-600 p-6 sm:p-8 text-center text-white">
            <div className="inline-block bg-white/20 p-3 sm:p-4 rounded-full mb-3 sm:mb-4">
              {isRegistering ? <UserPlus size={28} /> : <Lock size={28} />}
            </div>
            <h1 className="text-xl sm:text-2xl font-bold mb-1 sm:mb-2">{isRegistering ? t.registerTitle : t.loginTitle}</h1>
          </div>
          <div className="p-6 sm:p-8">
            <form onSubmit={handleAuth} className="space-y-4 sm:space-y-5">
              {loginError && <div className="p-3 bg-red-50 text-red-600 rounded-lg text-xs sm:text-sm text-center">{loginError}</div>}
              {isRegistering && (
                <div>
                  <label className="block text-xs sm:text-sm font-medium text-gray-700 mb-1">{t.name}</label>
                  <input type="text" name="name" required className="w-full px-3 py-2 sm:px-4 sm:py-3 rounded-xl border border-gray-300 focus:ring-2 focus:ring-blue-600 outline-none text-sm" />
                </div>
              )}
              <div>
                <label className="block text-xs sm:text-sm font-medium text-gray-700 mb-1">{t.email}</label>
                <input type="email" name="email" required className="w-full px-3 py-2 sm:px-4 sm:py-3 rounded-xl border border-gray-300 focus:ring-2 focus:ring-blue-600 outline-none text-sm" />
              </div>
              <div>
                <label className="block text-xs sm:text-sm font-medium text-gray-700 mb-1">{t.password}</label>
                <input type="password" name="password" required className="w-full px-3 py-2 sm:px-4 sm:py-3 rounded-xl border border-gray-300 focus:ring-2 focus:ring-blue-600 outline-none text-sm" />
              </div>
              <button type="submit" className="w-full py-2.5 sm:py-3 bg-blue-600 text-white rounded-xl font-bold hover:bg-blue-700 shadow-md text-sm sm:text-base">
                {isRegistering ? t.regBtn : t.authBtn}
              </button>
            </form>
            <div className="mt-4 sm:mt-6 text-center">
              <button onClick={() => { setIsRegistering(!isRegistering); setLoginError(''); }} className="text-xs sm:text-sm text-blue-600 hover:underline font-medium">
                {isRegistering ? t.switchToLogin : t.switchToReg}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (currentUser.role === 'admin') {
    const filteredTasks = tasks.filter(t => 
      t.evaluator.toLowerCase().includes(searchTerm.toLowerCase()) ||
      t.evaluatee.toLowerCase().includes(searchTerm.toLowerCase()) ||
      t.evaluatorEmail?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      t.evaluateeEmail?.toLowerCase().includes(searchTerm.toLowerCase())
    );

    return (
      <div className="min-h-screen bg-gray-50/50 pb-10" dir={t.dir}>
        {renderTopBar()}
        
        {viewingResult && (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4" dir="ltr">
            <div className="bg-white rounded-2xl w-full max-w-lg p-5 sm:p-6 shadow-2xl overflow-y-auto max-h-[90vh]">
              <div className="flex justify-between items-start mb-4 sm:mb-6 border-b border-gray-100 pb-4">
                <div className="pr-4">
                  <h2 className="text-lg sm:text-xl font-bold text-gray-900">Report Details</h2>
                  <p className="text-xs sm:text-sm text-gray-500 break-all mt-1">From: {viewingResult?.type === 'trainee_to_consultant' ? viewingResult?.evaluatorEmail.split('@')[0] : viewingResult?.evaluator} <br/>To: {viewingResult?.evaluatee}</p>
                </div>
                <div className="text-right shrink-0 bg-blue-50 px-4 py-2 rounded-xl">
                  <span className="text-2xl sm:text-3xl font-bold text-blue-600">{viewingResult?.score}%</span>
                  <p className="text-[10px] sm:text-xs text-blue-600/80 font-medium">Overall Score</p>
                </div>
              </div>
              <div className="space-y-4 sm:space-y-5 mb-6 sm:mb-8">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-3 bg-gray-50 rounded-xl">
                  <span className="text-xs sm:text-sm font-medium text-gray-700">{viewingResult?.type === 'consultant_to_trainee' ? t.qsConsultant.q1 : t.qsTrainee.q1}</span>
                  <ScoreViewer value={viewingResult?.scoresDetail?.q1} />
                </div>
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-3 bg-gray-50 rounded-xl">
                  <span className="text-xs sm:text-sm font-medium text-gray-700">{viewingResult?.type === 'consultant_to_trainee' ? t.qsConsultant.q2 : t.qsTrainee.q2}</span>
                  <ScoreViewer value={viewingResult?.scoresDetail?.q2} />
                </div>
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-3 bg-gray-50 rounded-xl">
                  <span className="text-xs sm:text-sm font-medium text-gray-700">{viewingResult?.type === 'consultant_to_trainee' ? t.qsConsultant.q3 : t.qsTrainee.q3}</span>
                  <ScoreViewer value={viewingResult?.scoresDetail?.q3} />
                </div>
              </div>
              <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
                <button onClick={() => setViewingResult(null)} className="w-full px-4 py-2.5 sm:py-3 border border-gray-300 rounded-xl text-gray-700 font-medium hover:bg-gray-50 text-sm sm:text-base transition-colors">Close</button>
                <button onClick={() => resetEval(viewingResult.id)} className="w-full px-4 py-2.5 sm:py-3 bg-red-50 text-red-600 rounded-xl font-medium hover:bg-red-100 flex items-center justify-center gap-2 text-sm sm:text-base transition-colors">
                  <RotateCcw size={16} /> <span className="truncate">{t.unlockBtn}</span>
                </button>
              </div>
            </div>
          </div>
        )}

        <div className="p-4 sm:p-6 max-w-7xl mx-auto space-y-4 sm:space-y-8">
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4 sm:p-6 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div className="flex items-center gap-3 sm:gap-4">
              <div className="bg-blue-50 p-2 sm:p-4 rounded-xl text-blue-600 shrink-0"><Upload size={20} className="sm:w-6 sm:h-6" /></div>
              <div>
                <h3 className="font-bold text-sm sm:text-lg text-gray-900">{t.uploadTitle}</h3>
                <p className="text-[10px] sm:text-sm text-gray-500">{t.excelDesc}</p>
              </div>
            </div>
            <div className="flex flex-col sm:flex-row gap-2 sm:gap-3 w-full md:w-auto mt-2 md:mt-0">
              <button onClick={exportToExcel} className="flex items-center justify-center gap-2 px-4 py-2.5 sm:py-3 border border-gray-300 text-gray-700 rounded-xl font-medium hover:bg-gray-50 text-xs sm:text-sm w-full sm:w-auto transition-colors">
                <Download size={14} />
                <span className="truncate">{t.exportBtn}</span>
              </button>
              <input type="file" id="schedule-upload" className="hidden" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} />
              <label htmlFor="schedule-upload" className="flex items-center justify-center px-4 py-2.5 sm:py-3 bg-blue-600 text-white rounded-xl font-medium hover:bg-blue-700 cursor-pointer text-xs sm:text-sm w-full sm:w-auto transition-colors shadow-sm">
                <span className="truncate">{t.selectBtn}</span>
              </label>
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-6">
            <div className="bg-blue-50 p-3 sm:p-6 rounded-xl sm:rounded-2xl">
              <div className="text-xl sm:text-4xl font-bold text-blue-700 mb-0.5 sm:mb-1">{tasks.length}</div>
              <div className="text-blue-600/80 text-[10px] sm:text-sm font-medium truncate">{t.total}</div>
            </div>
            <div className="bg-green-50 p-3 sm:p-6 rounded-xl sm:rounded-2xl">
              <div className="text-xl sm:text-4xl font-bold text-green-700 mb-0.5 sm:mb-1">{tasks.filter(x => x.status === 'Completed').length}</div>
              <div className="text-green-600/80 text-[10px] sm:text-sm font-medium truncate">{t.completed}</div>
            </div>
            <div className="bg-yellow-50 p-3 sm:p-6 rounded-xl sm:rounded-2xl">
              <div className="text-xl sm:text-4xl font-bold text-yellow-700 mb-0.5 sm:mb-1">{tasks.filter(x => x.status === 'Pending').length}</div>
              <div className="text-yellow-600/80 text-[10px] sm:text-sm font-medium truncate">{t.pending}</div>
            </div>
            <div className="bg-red-50 p-3 sm:p-6 rounded-xl sm:rounded-2xl">
              <div className="text-xl sm:text-4xl font-bold text-red-700 mb-0.5 sm:mb-1">{tasks.filter(x => x.status === 'Missed').length}</div>
              <div className="text-red-600/80 text-[10px] sm:text-sm font-medium truncate">{t.missed}</div>
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-4 sm:p-6 border-b border-gray-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3 sm:gap-4">
              <div className="flex items-center gap-2">
                <Activity size={18} className="text-blue-600 sm:w-5 sm:h-5" />
                <h3 className="font-bold text-sm sm:text-lg">{t.tracker}</h3>
              </div>
              <div className="relative w-full sm:w-64">
                <Search size={14} className={`absolute top-1/2 -translate-y-1/2 text-gray-400 sm:w-4 sm:h-4 ${t.dir === 'rtl' ? 'right-3' : 'left-3'}`} />
                <input 
                  type="text" 
                  placeholder={t.search} 
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className={`w-full py-2 bg-gray-50 border border-gray-200 rounded-lg text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-blue-600 transition-shadow ${t.dir === 'rtl' ? 'pr-9 pl-3' : 'pl-9 pr-3'}`}
                />
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[500px]" dir={t.dir}>
                <thead>
                  <tr className="border-b border-gray-100 bg-gray-50/50">
                    <th className="p-3 sm:p-4 text-[10px] sm:text-sm font-bold text-gray-700">{t.evaluator}</th>
                    <th className="p-3 sm:p-4 text-[10px] sm:text-sm font-bold text-gray-700">{t.evaluatee}</th>
                    <th className="p-3 sm:p-4 text-[10px] sm:text-sm font-bold text-gray-700">{t.status}</th>
                    <th className="p-3 sm:p-4 text-[10px] sm:text-sm font-bold text-gray-700">{t.deadline}</th>
                    <th className="p-3 sm:p-4 text-[10px] sm:text-sm font-bold text-gray-700">{t.actions}</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTasks.length === 0 ? (
                    <tr><td colSpan="5" className="p-4 text-center text-gray-400 text-sm">No tasks found.</td></tr>
                  ) : (
                    filteredTasks.map((task) => (
                      <tr key={task.id} className="border-b border-gray-50 hover:bg-gray-50/50">
                        <td className="p-3 sm:p-4 text-xs sm:text-sm text-gray-600 font-medium">
                          {task.type === 'trainee_to_consultant' ? `${task.evaluatorEmail.split('@')[0]} (Anon)` : task.evaluator}
                        </td>
                        <td className="p-3 sm:p-4 text-xs sm:text-sm text-gray-600">{task.evaluatee}</td>
                        <td className="p-3 sm:p-4"><StatusBadge status={task.status} /></td>
                        <td className="p-3 sm:p-4 text-[10px] sm:text-xs text-gray-400 font-mono whitespace-nowrap">{task.deadline}</td>
                        <td className="p-3 sm:p-4">
                          <button 
                            onClick={() => setViewingResult(task)}
                            disabled={task.status !== 'Completed'} 
                            className={`inline-flex items-center justify-center gap-1 p-1.5 sm:px-3 sm:py-1.5 rounded-lg text-xs sm:text-sm font-medium transition-colors ${task.status === 'Completed' ? 'text-blue-600 hover:bg-blue-50' : 'text-gray-300 cursor-not-allowed'}`}
                          >
                            <Eye size={14} className="sm:w-4 sm:h-4" /> <span className="hidden sm:inline">{t.view}</span>
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const myTasks = tasks.filter(t => t.evaluatorEmail === currentUser.email);
  const isConsultant = myTasks.some(t => t.type === 'consultant_to_trainee');
  const isTrainee = myTasks.some(t => t.type === 'trainee_to_consultant');
  const pendingTasks = myTasks.filter(t => t.status === 'Pending');
  const completedTasks = myTasks.filter(t => t.status !== 'Pending');

  const isFormComplete = scores.q1 !== null && scores.q2 !== null && scores.q3 !== null;
  const currentTotal = (scores.q1 || 0) + (scores.q2 || 0) + (scores.q3 || 0);
  const currentPercent = isFormComplete ? Math.round((currentTotal / 12) * 100) : 0;

  return (
    <div className="min-h-screen bg-gray-50/50 pb-10" dir={t.dir}>
      {renderTopBar()}
      
      {showModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4" dir="ltr">
          <div className="bg-white rounded-2xl w-full max-w-lg p-5 sm:p-6 shadow-2xl overflow-y-auto max-h-[90vh]">
            <div className="flex justify-between items-start mb-4 sm:mb-6 border-b border-gray-100 pb-4">
              <div className="pr-4">
                <h2 className="text-lg sm:text-xl font-bold text-gray-900">Evaluate</h2>
                <p className="text-xs sm:text-sm text-gray-500 break-all mt-1">{activeTask?.evaluatee}</p>
              </div>
              <div className="text-right shrink-0 bg-blue-50 px-4 py-2 rounded-xl">
                <span className={`text-2xl sm:text-3xl font-bold ${isFormComplete ? 'text-blue-600' : 'text-gray-300'}`}>
                  {isFormComplete ? `${currentPercent}%` : '--'}
                </span>
                <p className={`text-[10px] sm:text-xs font-medium ${isFormComplete ? 'text-blue-600/80' : 'text-gray-400'}`}>Score</p>
              </div>
            </div>
            <div className="space-y-6 sm:space-y-8 mb-6 sm:mb-8">
              <div className="flex flex-col gap-2">
                <span className="text-xs sm:text-sm font-medium text-gray-700">{activeTask?.type === 'consultant_to_trainee' ? t.qsConsultant.q1 : t.qsTrainee.q1}</span>
                <ScoreSelector value={scores.q1} onChange={(v) => setScores(p => ({...p, q1: v}))} />
              </div>
              <div className="flex flex-col gap-2">
                <span className="text-xs sm:text-sm font-medium text-gray-700">{activeTask?.type === 'consultant_to_trainee' ? t.qsConsultant.q2 : t.qsTrainee.q2}</span>
                <ScoreSelector value={scores.q2} onChange={(v) => setScores(p => ({...p, q2: v}))} />
              </div>
              <div className="flex flex-col gap-2">
                <span className="text-xs sm:text-sm font-medium text-gray-700">{activeTask?.type === 'consultant_to_trainee' ? t.qsConsultant.q3 : t.qsTrainee.q3}</span>
                <ScoreSelector value={scores.q3} onChange={(v) => setScores(p => ({...p, q3: v}))} />
              </div>
            </div>
            <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
              <button onClick={() => { setShowModal(false); setActiveTask(null); setScores({ q1: null, q2: null, q3: null }); }} className="w-full px-4 py-2.5 sm:py-3 border border-gray-300 rounded-xl text-gray-700 font-medium hover:bg-gray-50 text-sm sm:text-base transition-colors">Cancel</button>
              <button onClick={submitEval} disabled={!isFormComplete} className={`w-full px-4 py-2.5 sm:py-3 rounded-xl font-medium text-sm sm:text-base transition-colors shadow-sm ${isFormComplete ? 'bg-blue-600 text-white hover:bg-blue-700' : 'bg-gray-100 text-gray-400 cursor-not-allowed'}`}>Submit</button>
            </div>
          </div>
        </div>
      )}

      {(!isConsultant && !isTrainee) ? (
        <div className="p-6 sm:p-8 max-w-7xl mx-auto text-center mt-10 sm:mt-20">
          <h2 className="text-lg sm:text-2xl font-bold text-gray-500">{t.noTasks}</h2>
        </div>
      ) : (
        <div className="p-4 sm:p-8 max-w-7xl mx-auto">
          {isTrainee && (
            <div className="bg-blue-50 border border-blue-100 p-3 sm:p-4 rounded-xl mb-4 sm:mb-8 flex gap-2 sm:gap-3 text-blue-800 items-start md:items-center">
              <Shield size={16} className="shrink-0 mt-0.5 md:mt-0 sm:w-5 sm:h-5" />
              <p className="text-[10px] sm:text-sm font-medium">Your evaluations of consultants are strictly anonymous.</p>
            </div>
          )}
          
          {pendingTasks.length > 0 && <h2 className="text-lg sm:text-2xl font-bold text-gray-900 mb-3 sm:mb-6">{t.pending}</h2>}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-6 mb-6 sm:mb-8">
            {pendingTasks.map(task => (
              <div key={task.id} className="bg-white p-4 sm:p-6 rounded-2xl shadow-sm border border-gray-100 flex flex-col hover:border-blue-200 transition-colors">
                <div className="flex justify-between items-start mb-4 sm:mb-6">
                  <div className="pr-2">
                    <div className="text-[10px] sm:text-xs font-bold text-gray-400 uppercase mb-0.5 sm:mb-1">{task.type === 'trainee_to_consultant' ? 'Evaluate Consultant' : 'Trainee'}</div>
                    <div className="text-sm sm:text-lg font-bold text-gray-900 break-words">{task.evaluatee}</div>
                    <div className="text-[10px] sm:text-xs text-gray-400 font-mono mt-1.5">{task.deadline}</div>
                  </div>
                  <StatusBadge status={task.status} />
                </div>
                <div className="mt-auto">
                  <button onClick={() => { setActiveTask(task); setShowModal(true); }} className="w-full py-2.5 sm:py-3 bg-blue-600 text-white rounded-xl text-sm sm:text-base font-medium hover:bg-blue-700 transition-colors shadow-sm">
                    <span className="truncate">{task.type === 'trainee_to_consultant' ? 'Evaluate Anonymously' : 'Evaluate'}</span>
                  </button>
                </div>
              </div>
            ))}
          </div>

          {completedTasks.length > 0 && <h2 className="text-lg sm:text-2xl font-bold text-gray-900 mb-3 sm:mb-6">{t.completed}</h2>}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-6">
            {completedTasks.map(task => (
              <div key={task.id} className="bg-white p-4 sm:p-6 rounded-2xl shadow-sm border border-gray-100 flex flex-col opacity-75">
                <div className="flex justify-between items-start mb-4 sm:mb-6">
                  <div className="pr-2">
                    <div className="text-[10px] sm:text-xs font-bold text-gray-400 uppercase mb-0.5 sm:mb-1">{task.type === 'trainee_to_consultant' ? 'Consultant' : 'Trainee'}</div>
                    <div className="text-sm sm:text-lg font-bold text-gray-900 break-words">{task.evaluatee}</div>
                  </div>
                  <StatusBadge status={task.status} />
                </div>
                <div className="mt-auto pt-3 sm:pt-4 border-t border-gray-50">
                  <div className="w-full py-2 text-center text-[10px] sm:text-sm text-gray-500 font-medium bg-gray-50 rounded-xl">
                    {t.submitted} (Score: {task.score}%)
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}