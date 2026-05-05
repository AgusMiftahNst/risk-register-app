/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { 
  LayoutDashboard, 
  ShieldAlert, 
  Search,
  FileSearch, 
  BarChart3, 
  Settings2, 
  RefreshCcw, 
  RotateCw,
  ClipboardList, 
  Map, 
  Users, 
  FileText,
  Menu,
  X,
  ChevronRight,
  Plus,
  Minus,
  Edit2,
  Edit,
  Check,
  CheckCircle,
  Eye,
  EyeOff,
  Trash2,
  Lock,
  Target,
  LogOut,
  FileSpreadsheet,
  FileJson,
  Download,
  ExternalLink,
  ShieldCheck
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import html2canvas from 'html2canvas';

// --- Types ---

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errMessage = error instanceof Error ? error.message : String(error);
  const errInfo: FirestoreErrorInfo = {
    error: errMessage,
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
    },
    operationType,
    path
  };

  if (errMessage.includes('Quota limit exceeded')) {
    alert("Kapasitas (Quota) Firebase Gratis harian telah habis. Data tidak dapat disimpan/dibaca sampai reset otomatis besok pagi. Silakan hubungi pengembang jika ini sering terjadi.");
  }

  console.error('Firestore Error Detail: ', JSON.stringify(errInfo, null, 2));
  throw new Error(JSON.stringify(errInfo));
}

interface MenuItem {
  id: number;
  title: string;
  icon: React.ElementType;
}

// --- Components ---

/**
 * A debounced input component that uses local state to avoid cursor jumps
 * during real-time Firestore updates and supports browser undo/redo.
 */
function EditableTextarea({ value, onChange, disabled, placeholder, className, rows = 3, autoFocus }: { 
  value: string; 
  onChange: (val: string) => void; 
  disabled?: boolean; 
  placeholder?: string; 
  className?: string;
  rows?: number;
  autoFocus?: boolean;
}) {
  const [localValue, setLocalValue] = useState(value);

  useEffect(() => {
    setLocalValue(value);
  }, [value]);

  const handleBlur = () => {
    if (localValue !== value) {
      onChange(localValue);
    }
  };

  return (
    <textarea
      autoFocus={autoFocus}
      className={className}
      rows={rows}
      value={localValue}
      disabled={disabled}
      placeholder={placeholder}
      onChange={(e) => setLocalValue(e.target.value)}
      onBlur={handleBlur}
    />
  );
}

function EditableInput({ value, onChange, disabled, placeholder, className, type = "text", onFocus, onKeyDown, maxLength, pattern, filter }: { 
  value: string; 
  onChange: (val: string) => void; 
  disabled?: boolean; 
  placeholder?: string; 
  className?: string;
  type?: string;
  onFocus?: () => void;
  onKeyDown?: (e: React.KeyboardEvent) => void;
  maxLength?: number;
  pattern?: string;
  filter?: (val: string) => string;
}) {
  const [localValue, setLocalValue] = useState(value);

  useEffect(() => {
    setLocalValue(value);
  }, [value]);

  const handleBlur = () => {
    if (localValue !== value) {
      onChange(localValue);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    let val = e.target.value;
    if (filter) val = filter(val);
    setLocalValue(val);
  };

  return (
    <input
      type={type}
      className={className}
      value={localValue}
      disabled={disabled}
      placeholder={placeholder}
      maxLength={maxLength}
      pattern={pattern}
      onChange={handleChange}
      onBlur={handleBlur}
      onFocus={onFocus}
      onKeyDown={onKeyDown}
    />
  );
}

// --- Menu Definitions ---
const MENU_ITEMS_BASE: MenuItem[] = [
  { id: 1, title: 'I. PENETAPAN KONTEKS', icon: LayoutDashboard },
  { id: 2, title: 'II. IDENTIFIKASI RISIKO', icon: ShieldAlert },
  { id: 3, title: 'III. ANALISIS RISIKO', icon: BarChart3 },
  { id: 4, title: 'IV. EVALUASI RISIKO', icon: Settings2 },
  { id: 5, title: 'V. RENCANA PENANGANAN (RTP)', icon: RefreshCcw },
  { id: 6, title: 'VI. KOMUNIKASI PENGENDALIAN', icon: ClipboardList },
  { id: 7, title: 'VII. RENCANA MONITORING PI', icon: FileText },
  { id: 8, title: 'VIII. PETA RISIKO (HEATMAP)', icon: Map },
  { id: 9, title: 'IX. MONITORING KETERJADIAN', icon: Users },
  { id: 10, title: 'X. DOKUMEN FINAL', icon: ShieldCheck },
  { id: 11, title: 'XI. MANAJEMEN AKUN', icon: LayoutDashboard },
];

import { initializeApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, signOut, GoogleAuthProvider, signInWithPopup, signInAnonymously } from 'firebase/auth';
import { getFirestore, collection, onSnapshot, query, where, doc, getDoc, setDoc, getDocs, deleteDoc, updateDoc, getDocFromServer, writeBatch } from 'firebase/firestore';
import { db, auth, googleProvider } from './lib/firebase';

export default function App() {
  const [user, setUser] = useState<{username: string; role: string; uid: string} | null>(() => {
    const savedUser = localStorage.getItem('isman_user');
    if (savedUser) {
      try {
        return JSON.parse(savedUser);
      } catch (e) {
        localStorage.removeItem('isman_user');
      }
    }
    return null;
  });
  const [viewingUser, setViewingUser] = useState<{username: string; role: string; uid: string} | null>(null);
  const [loading, setLoading] = useState(true);
  const [activeMenu, setActiveMenu] = useState<number>(0);
  const [selectedRiskType, setSelectedRiskType] = useState<'strategis' | 'operasional' | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [accounts, setAccounts] = useState<any[]>([]);
  const [showLogoutConfirm, setShowLogoutConfirm] = useState(false);
  const [showChangePassword, setShowChangePassword] = useState(false);
  const [newPassword, setNewPassword] = useState('');
  const [confirmNewPassword, setConfirmNewPassword] = useState('');
  const [changePassError, setChangePassError] = useState('');
  const [changePassSuccess, setChangePassSuccess] = useState('');
  const [isExporting, setIsExporting] = useState(false);

  const menuItems = useMemo(() => {
    if (!user) return [];
    let items = [...MENU_ITEMS_BASE];
    
    if (user.role === 'Administrator' || user.role === 'Operator') {
      items.unshift({ id: 0, title: 'MONITORING PROGRESS', icon: LayoutDashboard });
    }
    
    // Only Administrators can see Account Management
    if (user.role !== 'Administrator') {
      items = items.filter(i => i.id !== 11);
    }
    
    return items;
  }, [user]);

  // Set initial active menu
  useEffect(() => {
    if (user) {
      if (user.role === 'Administrator' || user.role === 'Operator') {
        setActiveMenu(0);
      } else {
        setActiveMenu(1);
      }
    }
  }, [user]);

  useEffect(() => {
    if (user && !viewingUser) {
      setViewingUser(user);
    }
  }, [user, viewingUser]);

  const isActuallyReadOnly = useMemo(() => {
    if (!user || !viewingUser) return false;
    // Administrator can always edit
    if (user.role === 'Administrator') return false;
    // Everyone can edit their own data
    if (user.uid === viewingUser.uid) return false;
    // Otherwise, it's read-only (e.g. Operator viewing another user)
    return true;
  }, [user, viewingUser]);

  // Listen to Auth State
  useEffect(() => {
    const initDefaultAdmin = async () => {
      // Use sessionStorage to only check once per browser session
      if (sessionStorage.getItem('isman_admin_checked')) return;
      
      try {
        const q = query(collection(db, 'accounts'), where('username', '==', 'admin'));
        const snapshot = await getDocs(q);
        if (snapshot.empty) {
          const adminId = 'default_admin';
          await setDoc(doc(db, 'accounts', adminId), {
            username: 'admin',
            password: 'admin123',
            role: 'Administrator',
            createdAt: new Date().toISOString(),
            uid: adminId
          });
        }
        sessionStorage.setItem('isman_admin_checked', 'true');
      } catch (error) {
        console.error("Failed to init admin:", error);
      }
    };
    initDefaultAdmin();

    const unsubscribe = onAuthStateChanged(auth, async (firebaseUser) => {
      if (firebaseUser && !firebaseUser.isAnonymous) {
        try {
          // Fetch user role from Firestore
          const userDoc = await getDoc(doc(db, 'accounts', firebaseUser.uid));
          if (userDoc.exists()) {
            const data = userDoc.data();
            const displayName = firebaseUser.displayName || firebaseUser.email?.split('@')[0] || data.username || 'User';

            // Auto-upgrade developer email to Admin if they are currently just User
            if (firebaseUser.email === 'agusateng090800@gmail.com' && data.role !== 'Administrator') {
              await setDoc(doc(db, 'accounts', firebaseUser.uid), { role: 'Administrator' }, { merge: true });
              setUser({ username: displayName, role: 'Administrator', uid: firebaseUser.uid });
            } else {
              setUser({ username: displayName, role: data.role, uid: firebaseUser.uid });
            }
          } else {
            // New User (Google)
            const newUser = { 
              username: firebaseUser.displayName || firebaseUser.email?.split('@')[0] || 'User',
              role: firebaseUser.email === 'agusateng090800@gmail.com' ? 'Administrator' : 'User',
              createdAt: new Date().toISOString(),
              uid: firebaseUser.uid
            };
            await setDoc(doc(db, 'accounts', firebaseUser.uid), newUser);
            setUser({ username: newUser.username, role: newUser.role, uid: firebaseUser.uid });
          }
        } catch (error) {
          console.error("Auth sync error:", error);
        }
      } else {
        // Not signed in with Google, check manual session
        const savedUser = localStorage.getItem('isman_user');
        if (savedUser) {
          try {
            const parsed = JSON.parse(savedUser);
            // Re-verify the account exists in Firestore
            const q = query(collection(db, 'accounts'), where('username', '==', parsed.username));
            const querySnapshot = await getDocs(q);
            if (!querySnapshot.empty) {
              const accData = querySnapshot.docs[0];
              setUser({ username: accData.data().username, role: accData.data().role, uid: accData.id });
            } else {
              // Fallback for case-insensitivity during session restore
              const allSnap = await getDocs(collection(db, 'accounts'));
              const docMatch = allSnap.docs.find(d => (d.data().username || '').toLowerCase() === parsed.username.toLowerCase());
              if (docMatch) {
                setUser({ username: docMatch.data().username, role: docMatch.data().role, uid: docMatch.id });
              } else {
                setUser(null);
                localStorage.removeItem('isman_user');
              }
            }
          } catch (e) {
            setUser(null);
          }
        } else {
          setUser(null);
        }
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // Listen to Accounts for Admin
  useEffect(() => {
    if (user?.role === 'Administrator') {
      const q = collection(db, 'accounts');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const accs = snapshot.docs.map(doc => ({ ...doc.data(), uid: doc.id }));
        setAccounts(accs);
      }, (error) => {
        console.error('Firestore List Error (accounts):', error);
      });
      return () => unsubscribe();
    }
  }, [user]);

  // Migration logic for legacy data
  useEffect(() => {
    if (!user) return;
    
    // Use sessionStorage to only run migration check once per browser session per user
    const migrationSessionKey = `isman_migrated_${user.uid}`;
    if (sessionStorage.getItem(migrationSessionKey)) return;

    const migrateData = async () => {
      try {
        const batch = writeBatch(db);
        let hasChanges = false;

        const risksRef = collection(db, 'risk_identification');
        const q = query(risksRef, where('createdByUid', '==', user.uid));
        const snapshot = await getDocs(q);
        
        snapshot.docs.forEach(d => {
          if (!d.data().riskType) {
            batch.update(d.ref, { riskType: 'strategis' });
            hasChanges = true;
          }
        });

        const otherCollections = ['risk_context', 'monitoring_plan_pi', 'monitoring_communication', 'risk_occurrence_monitoring'];
        for (const col of otherCollections) {
          if (col === 'risk_context') {
             const legacyId = `risk_context_${user.uid}`;
             const newId = `risk_context_${user.uid}_strategis`;
             const snap = await getDoc(doc(db, col, legacyId));
             if (snap.exists()) {
               const newSnap = await getDoc(doc(db, col, newId));
               const isShell = newSnap.exists() && (!newSnap.data().informasiLain || newSnap.data().informasiLain === '-');
               if (!newSnap.exists() || isShell) {
                 batch.set(doc(db, col, newId), snap.data());
                 hasChanges = true;
               }
             }
             continue;
          }
          
          const colRef = collection(db, col);
          const colQ = query(colRef, where('createdByUid', '==', user.uid));
          const colSnap = await getDocs(colQ);
          colSnap.docs.forEach(d => {
            if (!d.data().riskType) {
              batch.update(d.ref, { riskType: 'strategis' });
              hasChanges = true;
            }
          });
        }
        
        if (hasChanges) {
          await batch.commit();
          console.log(`Executed migration batch for ${user.username}`);
        }
        
        // Mark as migrated for this session
        sessionStorage.setItem(migrationSessionKey, 'true');
      } catch (err) {
        console.error('Migration error:', err);
      }
    };

    migrateData();
  }, [user?.uid]);

  const handleExportAll = async (format: 'pdf' | 'excel') => {
    if (!viewingUser || !selectedRiskType) return;
    setIsExporting(true);
    try {
      const riskType = selectedRiskType;
      const targetUid = viewingUser.uid;
      const typeLabel = riskType === 'strategis' ? 'STRATEGIS' : 'OPERASIONAL';
      const typeCode = riskType === 'strategis' ? 'RSO' : 'ROO';
      const themeHex = riskType === 'strategis' ? '1E3A8A' : '059669';
      const themeColor = riskType === 'strategis' ? [30, 58, 138] : [5, 150, 105];
      
      const [contextSnap, risksSnap, piPlansSnap, monCommsSnap, occurrencesSnap] = await Promise.all([
        getDoc(doc(db, 'risk_context', `risk_context_${targetUid}_${riskType}`)),
        getDocs(query(collection(db, 'risk_identification'), where('createdByUid', '==', targetUid))),
        getDocs(query(collection(db, 'monitoring_plan_pi'), where('createdByUid', '==', targetUid))),
        getDocs(query(collection(db, 'monitoring_communication'), where('createdByUid', '==', targetUid))),
        getDocs(query(collection(db, 'risk_occurrence_monitoring'), where('createdByUid', '==', targetUid)))
      ]);

      const ctx = contextSnap.data() || {};
      const filterByType = (doc: any) => (doc.data().riskType || 'strategis') === riskType;

      const risksRaw = risksSnap.docs.filter(filterByType).map(d => ({ ...d.data(), id: d.id }));
      const piPlans = piPlansSnap.docs.filter(filterByType).map(d => d.data());
      const monComms = monCommsSnap.docs.filter(filterByType).map(d => d.data());
      const occurrences = occurrencesSnap.docs.filter(filterByType).map(d => d.data());
      
      const risks = risksRaw.map((r: any) => {
        const dS = (r.dampakScores || []).slice(0, ctx.participantCount || 5).filter((v: number) => v > 0);
        const kS = (r.kemungkinanScores || []).slice(0, ctx.participantCount || 5).filter((v: number) => v > 0);
        
        const avgD = dS.length > 0 ? parseFloat((dS.reduce((a: number, b: number) => a + b, 0) / dS.length).toFixed(2)) : 0;
        const avgK = kS.length > 0 ? parseFloat((kS.reduce((a: number, b: number) => a + b, 0) / kS.length).toFixed(2)) : 0;
        const riskLevel = getRiskLevel(avgD, avgK);
        return { 
          ...r, avgD, avgK, score: avgD * avgK, riskLevel,
          resD: parseFloat(r.residualDampak || 0), resK: parseFloat(r.residualKemungkinan || 0),
          resScore: parseFloat(r.residualDampak || 0) * parseFloat(r.residualKemungkinan || 0),
          resLevel: getRiskLevel(parseFloat(r.residualDampak || 0), parseFloat(r.residualKemungkinan || 0))
        };
      });

      const filename = `Laporan_Risiko_${typeCode}_${viewingUser.username}_${new Date().toISOString().split('T')[0]}`;

      if (format === 'excel') {
        const workbook = new ExcelJS.Workbook();
        const styleTbl = (sheet: ExcelJS.Worksheet, rows: number, cols: number, startR = 1) => {
          for (let r = startR; r < startR + rows; r++) {
            for (let c = 1; c <= cols; c++) {
              const cell = sheet.getCell(r, c);
              // Avoid styling if the row is effectively empty across the defined columns
              const row = sheet.getRow(r);
              let hasValue = false;
              for(let i=1; i<=cols; i++) {
                if (row.getCell(i).value !== null && row.getCell(i).value !== undefined && row.getCell(i).value !== '') {
                  hasValue = true;
                  break;
                }
              }
              if (!hasValue && r > startR) continue; // Keep header border even if empty, but skip body empty rows

              cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              cell.alignment = { vertical: 'middle', wrapText: true };
            }
          }
        };
        const styleHdr = (sheet: ExcelJS.Worksheet, row: number, cols: number) => {
          for (let c = 1; c <= cols; c++) {
            const cell = sheet.getCell(row, c);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: themeHex } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        };

        const ws1 = workbook.addWorksheet('1. Konteks');
        ws1.columns = [
          { width: 5 }, { width: 45 }, // Left table (No, Content)
          { width: 5 }, { width: 45 }, { width: 15 } // Right table (No, Content, Target)
        ];
        
        ws1.mergeCells('A1:E1');
        const titleCell = ws1.getCell('A1');
        titleCell.value = `I. PENETAPAN KONTEKS RISIKO ${typeLabel} OPD`;
        titleCell.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: themeHex } };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

        const ctxM = [
          ['Nama Pemerintah Daerah', ctx.namaPemda || '-'],
          ['Tahun Penilaian', ctx.tahunPenilaian || '-'],
          ['Periode yang dinilai', ctx.periodeRenstra || '-'],
          ['Urusan Pemerintahan', ctx.urusanPemerintahan || '-'],
          ['OPD yang Dinilai', ctx.opdDinilai || '-'],
          ['Sumber Data', ctx.sumberData || '-'],
          [riskType === 'operasional' ? 'Tujuan Operasional' : 'Tujuan Strategis', ctx.tujuanStrategis || '-']
        ];

        ctxM.forEach((m, i) => {
          const row = ws1.getRow(i + 3);
          row.getCell(1).value = m[0];
          row.getCell(1).font = { bold: true };
          row.getCell(2).value = ':';
          row.getCell(2).alignment = { horizontal: 'center' };
          ws1.mergeCells(i + 3, 3, i + 3, 5);
          row.getCell(3).value = m[1];
          row.getCell(3).alignment = { wrapText: true };
        });

        let curR = 12;
        // Group 1: Program/Sasaran (Left) & Kegiatan Utama/IKU Sasaran (Right)
        const row1SubHdr = ws1.getRow(curR);
        row1SubHdr.getCell(1).value = 'No'; row1SubHdr.getCell(2).value = riskType === 'operasional' ? 'Program' : 'Sasaran Strategis';
        row1SubHdr.getCell(3).value = 'No'; row1SubHdr.getCell(4).value = riskType === 'operasional' ? 'Kegiatan Utama' : 'IKU Sasaran OPD'; row1SubHdr.getCell(5).value = riskType === 'operasional' ? '' : 'Target';
        [1, 2].forEach(c => styleHdr(ws1, curR, c));
        [3, 4, 5].forEach(c => styleHdr(ws1, curR, c));
        curR += 1;

        const sasaranList = ctx.sasaran || [];
        const ikuSasaranList = ctx.ikuSasaran || [];
        const maxRow1 = Math.max(sasaranList.length, ikuSasaranList.length, 1);
        
        for (let i = 0; i < maxRow1; i++) {
          const r = ws1.getRow(curR + i);
          if (sasaranList[i]) {
            r.getCell(1).value = i + 1;
            r.getCell(2).value = sasaranList[i];
            [1, 2].forEach(c => {
              r.getCell(c).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              r.getCell(c).alignment = { vertical: 'middle', wrapText: true };
            });
          }
          if (ikuSasaranList[i]) {
            r.getCell(3).value = i + 1;
            r.getCell(4).value = ikuSasaranList[i].name;
            r.getCell(5).value = ikuSasaranList[i].target;
            [3, 4, 5].forEach(c => {
              r.getCell(c).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              r.getCell(c).alignment = { vertical: 'middle', wrapText: true };
            });
          }
        }
        curR += maxRow1 + 1;

        // Group 2: Subkegiatan/Program (Left) & Indikator Keluaran/IKU Program (Right)
        const row2SubHdr = ws1.getRow(curR);
        row2SubHdr.getCell(1).value = 'No'; row2SubHdr.getCell(2).value = riskType === 'operasional' ? 'Subkegiatan' : 'Program Strategis';
        row2SubHdr.getCell(3).value = 'No'; row2SubHdr.getCell(4).value = riskType === 'operasional' ? 'Indikator Keluaran' : 'IKU Program OPD'; 
        if (riskType !== 'operasional') row2SubHdr.getCell(5).value = 'Target';
        [1, 2].forEach(c => styleHdr(ws1, curR, c));
        [3, 4, 5].forEach(c => styleHdr(ws1, curR, c));
        curR += 1;

        const programList = ctx.program || [];
        const ikuProgramList = ctx.ikuProgram || [];
        const maxRow2 = Math.max(programList.length, ikuProgramList.length, 1);

        for (let i = 0; i < maxRow2; i++) {
          const r = ws1.getRow(curR + i);
          if (programList[i]) {
            r.getCell(1).value = i + 1;
            r.getCell(2).value = programList[i];
            [1, 2].forEach(c => {
              r.getCell(c).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              r.getCell(c).alignment = { vertical: 'middle', wrapText: true };
            });
          }
          if (ikuProgramList[i]) {
            r.getCell(3).value = i + 1;
            r.getCell(4).value = ikuProgramList[i].name;
            if (riskType !== 'operasional') r.getCell(5).value = ikuProgramList[i].target;
            [3, 4, 5].forEach(c => {
              if (riskType === 'operasional' && c === 5) return;
              r.getCell(c).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
              r.getCell(c).alignment = { vertical: 'middle', wrapText: true };
            });
          }
        }
        curR += maxRow2;

        // Informasi Lain
        ws1.getRow(curR).getCell(1).value = 'Informasi Lain';
        ws1.getRow(curR).getCell(1).font = { bold: true };
        ws1.getRow(curR).getCell(2).value = ':';
        ws1.mergeCells(curR, 3, curR, 5);
        ws1.getRow(curR).getCell(3).value = ctx.informasiLain || '-';
        ws1.getRow(curR).getCell(3).alignment = { wrapText: true };

        curR += 2;
        // Assessment Rows Subtitle
        ws1.mergeCells(`A${curR}:E${curR}`);
        const subTitleCell = ws1.getCell(`A${curR}`);
        subTitleCell.value = riskType === 'operasional' 
          ? 'Program, Kegiatan, Subkegiatan, dan Keluaran/Hasil Subkegiatan yang akan dilakukan penilaian risiko'
          : 'Tujuan, Sasaran, Program Strategis, IKU Program yang akan dilakukan penilaian risiko';
        subTitleCell.font = { bold: true };
        subTitleCell.alignment = { horizontal: 'left' };
        curR += 1;

        // Assessment Rows
        const aHeader = ws1.getRow(curR);
        const aCols = [
          'No', 
          riskType === 'operasional' ? 'Program' : 'Tujuan Strategis', 
          riskType === 'operasional' ? 'Kegiatan' : 'Sasaran Strategis', 
          riskType === 'operasional' ? 'Subkegiatan' : 'Program Strategis', 
          riskType === 'operasional' ? 'Keluaran/Hasil Subkegiatan' : 'IKU Program'
        ];
        aCols.forEach((col, i) => {
          aHeader.getCell(i + 1).value = col;
        });
        styleHdr(ws1, curR, aCols.length);
        const aRows = ctx.assessmentRows || [];
        aRows.forEach((row: any, i: number) => {
          const r = ws1.getRow(curR + 1 + i);
          r.getCell(1).value = i + 1;
          r.getCell(1).alignment = { horizontal: 'center' };
          r.getCell(2).value = row.tujuan;
          r.getCell(3).value = row.sasaran;
          r.getCell(4).value = row.program;
          r.getCell(5).value = row.iku;
        });
        styleTbl(ws1, aRows.length + 1, aCols.length, curR);

        curR += (aRows.length || 1) + 4;
        const sigX = 3;
        ws1.getCell(curR, sigX).value = `${ctx.ttdTempat || 'Paniai'}, ${ctx.ttdBulan || '-'}`;
        ws1.getCell(curR + 1, sigX).value = ctx.ttdJabatan || '-';
        ws1.getCell(curR + 2, sigX).value = ctx.ttdKabupaten || 'Kabupaten Paniai';
        ws1.getCell(curR + 7, sigX).value = ctx.ttdNama || '(Nama)';
        ws1.getCell(curR + 7, sigX).font = { bold: true, underline: true };
        ws1.getCell(curR + 8, sigX).value = `Pangkat: ${ctx.ttdPangkat || '-'}`;
        ws1.getCell(curR + 9, sigX).value = `NIP. ${ctx.ttdNip || '-'}`;

        const ws2 = workbook.addWorksheet('2. Identifikasi');
        const ws2Hdr = ['No', riskType === 'operasional' ? 'Subkegiatan' : 'Tujuan / Sasaran', riskType === 'operasional' ? 'Indikator Keluaran' : 'Indikator Kinerja', 'Risiko (Uraian)', 'Risiko (Kode)', 'Pemilik', 'Sebab (Uraian)', 'Sebab (Sumber)', 'Control (C/UC)', 'Dampak (Akibat)', 'Dampak (Pihak)'];
        ws2.getRow(1).values = ws2Hdr;
        styleHdr(ws2, 1, ws2Hdr.length);
        
        let ws2RowCount = 1;
        risks.forEach((r, i) => {
          const srsRaw = r.subRows || [];
          const srs = (srsRaw.length > 0 && (srsRaw[0].sebabUraian?.trim() || srsRaw[0].dampakUraian?.trim()))
            ? srsRaw 
            : [{
                sebabUraian: r.sebabUraian || '',
                sebabSumber: r.sebabSumber || '',
                control: r.control || '',
                dampakUraian: r.dampakUraian || '',
                dampakPihak: r.dampakPihak || ''
              }];
          
          srs.forEach((sub: any) => {
            ws2.addRow([
              i + 1, 
              r.tujuan || '-', 
              r.indikator || '-', 
              r.risikoUraian || '-', 
              r.risikoKode || '-', 
              r.pemilik || '-', 
              sub.sebabUraian || '-', 
              sub.sebabSumber || '-', 
              sub.control || '-', 
              sub.dampakUraian || '-', 
              sub.dampakPihak || '-'
            ]);
            ws2RowCount++;
          });
        });
        styleTbl(ws2, ws2RowCount, ws2Hdr.length); 
        ws2.columns = ws2Hdr.map(() => ({ width: 30 }));
        ws2.getColumn(1).width = 5;

        const ws3 = workbook.addWorksheet('3. Analisis');
        const pCount = ctx.participantCount || 5;
        const pNames = Array.from({ length: pCount }, (_, k) => `P${k + 1}`);

        // Define column header labels
        const col1 = [
          'No',
          'Analisis Risiko Teridentifikasi',
          'Dampak (Impact)', ...Array.from({ length: pCount - 1 }, () => ''),
          'Rata-rata Dampak',
          'Probabilitas (S)', ...Array.from({ length: pCount - 1 }, () => ''),
          'Rata-rata Kemungkinan',
          'Score (DxS)',
          'Level'
        ];
        ws3.getRow(1).values = col1;
        ws3.getRow(2).values = [
          '', '', 
          ...pNames,
          '',
          ...pNames,
          '', '', ''
        ];

        // Headers Styling
        styleHdr(ws3, 1, col1.length);
        styleHdr(ws3, 2, col1.length);

        // Merging
        ws3.mergeCells(1, 1, 2, 1); // No
        ws3.mergeCells(1, 2, 2, 2); // Identifikasi
        ws3.mergeCells(1, 3, 1, 3 + pCount - 1); // Impact Group
        ws3.mergeCells(1, 3 + pCount, 2, 3 + pCount); // Avg D
        ws3.mergeCells(1, 3 + pCount + 1, 1, 3 + 2 * pCount); // Prob Group
        ws3.mergeCells(1, 3 + 2 * pCount + 1, 2, 3 + 2 * pCount + 1); // Avg S
        ws3.mergeCells(1, 3 + 2 * pCount + 2, 2, 3 + 2 * pCount + 2); // Score
        ws3.mergeCells(1, 3 + 2 * pCount + 3, 2, 3 + 2 * pCount + 3); // Level

        risks.forEach((r, i) => {
          const ds = r.dampakScores || [];
          const ks = r.kemungkinanScores || [];
          const rowData = [i + 1, `${r.risikoKode}: ${r.risikoUraian}`];
          // Dampak individual
          for (let k = 0; k < pCount; k++) rowData.push(ds[k] || 0);
          rowData.push(r.avgD.toFixed(2));
          // Prob individual
          for (let k = 0; k < pCount; k++) rowData.push(ks[k] || 0);
          rowData.push(r.avgK.toFixed(2));
          // Final results
          rowData.push(r.score.toFixed(2), r.riskLevel.label);
          ws3.addRow(rowData);
        });
        styleTbl(ws3, risks.length + 2, col1.length);
        ws3.columns = col1.map((_, idx) => ({ width: idx === 1 ? 45 : 10 }));
        ws3.getColumn(1).width = 5;

        const ws4 = workbook.addWorksheet('4. Risiko Aktual');
        const col4_1 = ['No', 'Risiko Teridentifikasi', 'Risiko Awal', '', '', '', 'Pengendalian Saat Ini', 'Celah (Gap)', 'Risiko Sisa (Residual)', '', '', ''];
        const col4_2 = ['', '', 'D', 'K', 'Skor', 'Level', '', '', 'D', 'S', 'Skor', 'Level'];
        
        ws4.getRow(1).values = col4_1;
        ws4.getRow(2).values = col4_2;
        
        styleHdr(ws4, 1, col4_1.length);
        styleHdr(ws4, 2, col4_2.length);

        ws4.mergeCells(1, 1, 2, 1); // No
        ws4.mergeCells(1, 2, 2, 2); // Risiko
        ws4.mergeCells(1, 3, 1, 6); // Risiko Awal
        ws4.mergeCells(1, 7, 2, 7); // Pengendalian
        ws4.mergeCells(1, 8, 2, 8); // Celah
        ws4.mergeCells(1, 9, 1, 12); // Risiko Sisa

        risks.forEach((r, i) => ws4.addRow([
          i + 1, 
          `${r.risikoKode}: ${r.risikoUraian}`, 
          r.avgD.toFixed(2), 
          r.avgK.toFixed(2), 
          r.score.toFixed(2), 
          r.riskLevel.label,
          r.rtpControl,
          r.rtpGap,
          r.resD.toFixed(2), 
          r.resK.toFixed(2), 
          r.resScore.toFixed(2), 
          r.resLevel.label
        ]));
        
        styleTbl(ws4, risks.length + 2, col4_1.length);
        ws4.columns = col4_1.map((_, idx) => ({ width: idx === 1 || idx === 6 || idx === 7 ? 40 : 10 }));
        ws4.getColumn(1).width = 5;

        const ws5 = workbook.addWorksheet('5. RTP');
        const ws5Hdr = ['No', 'Risiko (Residual)', 'Pengendalian Sisa', 'Celah Sisa', 'Rencana Tindak (RTP) Baru', 'Penanggung Jawab (PJ)', 'Deadline'];
        ws5.getRow(1).values = ws5Hdr;
        styleHdr(ws5, 1, ws5Hdr.length);
        risks.forEach((r, i) => {
          if (r.resScore > 0 && r.resLevel.level >= 3) {
            ws5.addRow([
              i + 1, 
              `${r.risikoKode}: ${r.risikoUraian}`, 
              r.rtpControl, 
              r.rtpGap, 
              r.rtpAction, 
              r.rtpPJ, 
              r.rtpDeadline
            ]);
          }
        });
        styleTbl(ws5, ws5.rowCount, ws5Hdr.length);
        ws5.columns = ws5Hdr.map((_, idx) => ({ width: idx === 1 || idx === 2 || idx === 3 || idx === 4 ? 35 : 15 }));
        ws5.getColumn(1).width = 5;

        const ws6 = workbook.addWorksheet('6. Komunikasi');
        const ws6Hdr = ['No', 'Kegiatan Pengendalian', 'Media', 'Penyedia Informasi', 'Penerima Informasi', 'Rencana Waktu', 'Realisasi Waktu', 'Keterangan'];
        ws6.getRow(1).values = ws6Hdr;
        styleHdr(ws6, 1, ws6Hdr.length);
        risks.filter(r => r.resScore > 0 && r.resLevel.level >= 3).forEach((r, i) => {
          ws6.addRow([
            i + 1, 
            r.rtpAction || '-', 
            r.commMedia || '-', 
            r.commProvider || '-', 
            r.commReceiver || '-', 
            r.commPlanTime || '-', 
            r.commRealTime || '-', 
            r.commNotes || '-'
          ]);
        });
        styleTbl(ws6, ws6.rowCount, ws6Hdr.length);
        ws6.columns = ws6Hdr.map(() => ({ width: 25 }));
        ws6.getColumn(1).width = 5;

        const ws7 = workbook.addWorksheet('7. PI Monitoring');
        const ws7Hdr = ['No', 'Kegiatan Pengendalian', 'Bentuk/Metode Pemantauan', 'Penanggung Jawab', 'Rencana Waktu', 'Realisasi Waktu', 'Keterangan'];
        ws7.getRow(1).values = ws7Hdr;
        styleHdr(ws7, 1, ws7Hdr.length);
        risks.filter(r => r.resScore > 0 && r.resLevel.level >= 3).forEach((r, i) => {
          ws7.addRow([
            i + 1, 
            r.rtpAction || '-', 
            r.monMethod || '-', 
            r.monPJ || '-', 
            r.monPlanTime || '-', 
            r.monRealTime || '-', 
            r.monNotes || '-'
          ]);
        });
        styleTbl(ws7, ws7.rowCount, ws7Hdr.length);
        ws7.columns = ws7Hdr.map(() => ({ width: 25 }));
        ws7.getColumn(1).width = 5;

        const ws9 = workbook.addWorksheet('9. Kejadian');
        const ws9Hdr = ['No', 'Risiko Teridentifikasi', 'Kode', 'Tanggal Kejadian', 'Sebab Actual', 'Dampak Actual', 'Keterangan Kejadian', 'RTP', 'Rencana RTP', 'Realisasi RTP', 'Keterangan RTP'];
        ws9.getRow(1).values = ws9Hdr;
        styleHdr(ws9, 1, ws9Hdr.length);
        risks.forEach((r, i) => {
          ws9.addRow([
            i + 1, 
            r.risikoUraian, 
            r.risikoKode, 
            r.eventDate || '-', 
            r.eventCause || '-', 
            r.eventImpact || '-', 
            r.eventNotes || '-', 
            r.rtpAction || '-', 
            r.rtpPlanDate || '-', 
            r.rtpRealDate || '-', 
            r.rtpNotesContent || '-'
          ]);
        });
        styleTbl(ws9, ws9.rowCount, ws9Hdr.length);
        ws9.columns = ws9Hdr.map(() => ({ width: 25 }));
        ws9.getColumn(1).width = 5;

        saveAs(new Blob([await workbook.xlsx.writeBuffer()]), `${filename}.xlsx`);
      } else {
        const docPDF = new jsPDF('l', 'mm', 'a4');
        const pW = docPDF.internal.pageSize.getWidth();
        const pH = docPDF.internal.pageSize.getHeight();
        const m = 15;
        
        const hdrPDF = (title: string) => {
          docPDF.setFontSize(14).setFont('helvetica', 'bold');
          const t = `${title} ${typeLabel} OPD`;
          const tw = docPDF.getTextWidth(t);
          docPDF.text(t, (pW - tw) / 2, 15);
          docPDF.setDrawColor(themeColor[0], themeColor[1], themeColor[2]).setLineWidth(0.8).line(m, 18, pW - m, 18);
        };

        const sigPDF = (y: number) => {
          if (y > pH - 65) { docPDF.addPage(); y = 30; }
          const x = pW - 90;
          docPDF.setFontSize(10).setFont('helvetica', 'normal').text(`${ctx.ttdTempat || 'Paniai'}, ${ctx.ttdBulan || '-'}`, x, y);
          y += 6; docPDF.setFont('helvetica', 'bold').text(`${ctx.ttdJabatan || '-'}`, x, y);
          y += 6; docPDF.text(`${ctx.ttdKabupaten || 'Kabupaten Paniai'}`, x, y);
          y += 20; docPDF.text(`${ctx.ttdNama || '(Nama)'}`, x, y);
          docPDF.line(x, y + 1.2, x + docPDF.getTextWidth(ctx.ttdNama || '(Nama)'), y + 1.2);
          y += 6; docPDF.setFont('helvetica', 'normal').text(`Pangkat: ${ctx.ttdPangkat || '-'}`, x, y);
          y += 5; docPDF.text(`NIP. ${ctx.ttdNip || '-'}`, x, y);
        };

        const tS: any = {
          fontSize: 7, cellPadding: 1.5, valign: 'middle',
          headStyles: { fillColor: themeColor, textColor: 255, fontStyle: 'bold', halign: 'center', fontSize: 7.5 },
          margin: { top: 20, left: m, right: m, bottom: 15 },
          styles: { textColor: [30, 41, 59], lineColor: [200, 200, 200], lineWidth: 0.1, overflow: 'linebreak' },
          alternateRowStyles: { fillColor: [249, 250, 251] }
        };

        hdrPDF('I. PENETAPAN KONTEKS RISIKO');
        let y = 22;
        const ctxRowsCommon = [
          ['Nama Pemerintah Daerah', `${ctx.namaPemda || '-'}`],
          ['Tahun Penilaian', `${ctx.tahunPenilaian || '-'}`],
          ['Periode yang dinilai', `${ctx.periodeRenstra || '-'}`],
          ['Urusan Pemerintahan', `${ctx.urusanPemerintahan || '-'}`],
          ['OPD yang Dinilai', `${ctx.opdDinilai || '-'}`],
          ['Sumber Data', `${ctx.sumberData || '-'}`],
        ];

        ctxRowsCommon.forEach(p => {
          docPDF.setFontSize(8).setFont('helvetica', 'bold').text(p[0], m, y);
          docPDF.text(':', m + 50, y);
          docPDF.setFont('helvetica', 'normal').text(p[1], m + 53, y);
          y += 4.2;
        });

        // Multiline support for Tujuan with aligned wrapping
        const tujuanLabel = riskType === 'operasional' ? 'Tujuan Operasional' : 'Tujuan Strategis';
        const tujuanVal = `${ctx.tujuanStrategis || '-'}`;
        docPDF.setFont('helvetica', 'bold').text(tujuanLabel, m, y);
        docPDF.text(':', m + 50, y);
        const splitTujuan = docPDF.splitTextToSize(tujuanVal, pW - m - 53 - m);
        docPDF.setFont('helvetica', 'normal').text(splitTujuan, m + 53, y);
        y += (splitTujuan.length * 4) + 1.5;

        // Group 1: Program/Sasaran (Left) & Kegiatan Utama/IKU Sasaran (Right)
        if (y > pH - 40) { docPDF.addPage(); y = 25; }
        const colW = (pW - (2 * m) - 6) / 2; 
        const gutter = 3;

        const startYGroup1 = y;
        // Tabel 1 (Sasaran/Program)
        autoTable(docPDF, {
          startY: startYGroup1,
          ...tS,
          margin: { left: m },
          tableWidth: colW,
          head: [['NO', riskType === 'operasional' ? 'PROGRAM' : 'SASARAN STRATEGIS']],
          body: (ctx.sasaran || []).map((s: string, i: number) => [i + 1, s]),
          columnStyles: { 0: { cellWidth: 8, halign: 'center' } }
        });
        const finalY1L = (docPDF as any).lastAutoTable.finalY || startYGroup1;

        // Tabel 2 (IKU Sasaran / Kegiatan Utama)
        autoTable(docPDF, {
          startY: startYGroup1,
          ...tS,
          margin: { left: m + colW + 6 },
          tableWidth: colW,
          head: [riskType === 'operasional' ? ['NO', 'KEGIATAN UTAMA'] : ['NO', 'IKU SASARAN OPD', 'TARGET']],
          body: (ctx.ikuSasaran || []).map((v: any, i: number) => {
            if (riskType === 'operasional') return [i + 1, v.name || '-'];
            return [i + 1, v.name || '-', v.target || '-'];
          }),
          columnStyles: { 0: { cellWidth: 8, halign: 'center' }, 2: { cellWidth: 15, halign: 'center' } }
        });
        const finalY1R = (docPDF as any).lastAutoTable.finalY || startYGroup1;

        y = Math.max(finalY1L, finalY1R) + gutter;

        // Group 2: Subkegiatan/Program (Left) & Indikator Keluaran/IKU Program (Right)
        const estH2 = Math.max((ctx.program || []).length, (ctx.ikuProgram || []).length) * 6 + 10;
        if (y + estH2 > pH - 15) { docPDF.addPage(); y = 25; }
        const startYGroup2 = y;

        autoTable(docPDF, {
          startY: startYGroup2,
          ...tS,
          margin: { left: m },
          tableWidth: colW,
          head: [['NO', riskType === 'operasional' ? 'SUBKEGIATAN' : 'PROGRAM STRATEGIS']],
          body: (ctx.program || []).map((p: string, i: number) => [i + 1, p]),
          columnStyles: { 0: { cellWidth: 8, halign: 'center' } }
        });
        const finalY2L = (docPDF as any).lastAutoTable.finalY || startYGroup2;

        autoTable(docPDF, {
          startY: startYGroup2,
          ...tS,
          margin: { left: m + colW + 6 },
          tableWidth: colW,
          head: [riskType === 'operasional' ? ['NO', 'INDIKATOR KELUARAN'] : ['NO', 'IKU PROGRAM OPD', 'TARGET']],
          body: (ctx.ikuProgram || []).map((v: any, i: number) => {
            if (riskType === 'operasional') return [i + 1, v.name || '-'];
            return [i + 1, v.name || '-', v.target || '-'];
          }),
          columnStyles: { 0: { cellWidth: 8, halign: 'center' }, ...(riskType === 'operasional' ? {} : { 2: { cellWidth: 15, halign: 'center' } }) }
        });
        const finalY2R = (docPDF as any).lastAutoTable.finalY || startYGroup2;
        y = Math.max(finalY2L, finalY2R); 

        if (y > pH - 10) { docPDF.addPage(); y = 15; }
        y += 0.5;
        docPDF.setFontSize(8).setFont('helvetica', 'bold').text('Informasi Lain', m, y);
        docPDF.text(':', m + 50, y);
        const splitInfo = docPDF.splitTextToSize(`${ctx.informasiLain || '-'}`, pW - m - 53 - m);
        docPDF.setFont('helvetica', 'normal').text(splitInfo, m + 53, y);
        y += (splitInfo.length * 4) + 1.5;

        const stText = riskType === 'operasional' 
          ? 'Program, Kegiatan, Subkegiatan, dan Keluaran/Hasil Subkegiatan yang akan dilakukan penilaian risiko'
          : 'Tujuan, Sasaran, Program Strategis, IKU Program yang akan dilakukan penilaian risiko';
        
        if (y > pH - 25) { docPDF.addPage(); y = 20; }
        docPDF.setFontSize(8.5).setFont('helvetica', 'bold');
        const stTextSplit = docPDF.splitTextToSize(stText, pW - m - m);
        docPDF.text(stTextSplit, m, y);
        y += (stTextSplit.length * 4) + 2;

        const aRows = ctx.assessmentRows || [];
        autoTable(docPDF, {
          startY: y,
          ...tS,
          head: [[
            'NO', 
            riskType === 'operasional' ? 'PROGRAM' : 'TUJUAN STRATEGIS', 
            riskType === 'operasional' ? 'KEGIATAN' : 'SASARAN STRATEGIS', 
            riskType === 'operasional' ? 'SUBKEGIATAN' : 'PROGRAM STRATEGIS', 
            riskType === 'operasional' ? 'KELUARAN/HASIL SUBKEGIATAN' : 'IKU PROGRAM'
          ]],
          body: aRows.map((r: any, i: number) => [i + 1, r.tujuan, r.sasaran, r.program, r.iku]),
          columnStyles: { 0: { cellWidth: 8, halign: 'center' } }
        });

        sigPDF((docPDF as any).lastAutoTable.finalY + 12);


        const participantCount = ctx.participantCount || 5;
        const pHeaders = Array.from({ length: participantCount }, (_, k) => `P${k + 1}`);

        const pdfPages = [
          { 
            t: 'II. IDENTIFIKASI RISIKO', 
            h: [['NO', riskType === 'operasional' ? 'SUBKEGIATAN' : 'TUJUAN / SASARAN', riskType === 'operasional' ? 'INDIKATOR KELUARAN' : 'INDIKATOR', 'RISIKO (URAIAN)', 'KODE', 'PEMILIK', 'SEBAB', 'SUMBER', 'C/UC', 'AKIBAT', 'PIHAK']], 
            b: risks.flatMap((r, i) => {
              const srsRaw = r.subRows || [];
              const srs = (srsRaw.length > 0 && (srsRaw[0].sebabUraian?.trim() || srsRaw[0].dampakUraian?.trim()))
                ? srsRaw 
                : [{ 
                    sebabUraian: r.sebabUraian || '', 
                    sebabSumber: r.sebabSumber || '', 
                    control: r.control || '', 
                    dampakUraian: r.dampakUraian || '', 
                    dampakPihak: r.dampakPihak || '' 
                  }];
              return srs.map((sub: any) => [
                i+1, 
                r.tujuan || '-', 
                r.indikator || '-', 
                r.risikoUraian || '-', 
                r.risikoKode || '-', 
                r.pemilik || '-', 
                sub.sebabUraian || '-', 
                sub.sebabSumber || '-', 
                sub.control || '-', 
                sub.dampakUraian || '-', 
                sub.dampakPihak || '-'
              ]);
            }),
            cWidths: { 0: { cellWidth: 8 }, 4: { cellWidth: 15 }, 8: { cellWidth: 12 } }
          },
          { 
            t: 'III. ANALISIS RISIKO (INHEREN)', 
            h: [
              [
                { content: 'NO', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'ANALISIS RISIKO TERIDENTIFIKASI', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'DAMPAK (IMPACT)', colSpan: participantCount, styles: { halign: 'center' } },
                { content: 'RATA-RATA DAMPAK', rowSpan: 2, styles: { halign: 'center', valign: 'middle', fillColor: [40, 40, 40] } },
                { content: 'PROBABILITAS (S)', colSpan: participantCount, styles: { halign: 'center' } },
                { content: 'RATA-RATA KEMUNGKINAN', rowSpan: 2, styles: { halign: 'center', valign: 'middle', fillColor: [40, 40, 40] } },
                { content: 'SKOR', rowSpan: 2, styles: { halign: 'center', valign: 'middle', fillColor: [6, 78, 59] } },
                { content: 'LEVEL', rowSpan: 2, styles: { halign: 'center', valign: 'middle', fillColor: [6, 78, 59] } }
              ],
              [
                ...Array.from({ length: participantCount }, (_, k) => ({ content: `P${k + 1}`, styles: { halign: 'center', fontSize: 6 } })),
                ...Array.from({ length: participantCount }, (_, k) => ({ content: `P${k + 1}`, styles: { halign: 'center', fontSize: 6 } }))
              ]
            ], 
            b: risks.map((r, i) => {
              const ds = r.dampakScores || [];
              const ks = r.kemungkinanScores || [];
              const row: any[] = [
                i + 1, 
                { content: `${r.risikoKode}\n${(r.risikoUraian || '').toUpperCase()}`, styles: { fontSize: 6.5 } }
              ];
              for(let k=0; k<participantCount; k++) row.push(ds[k] || 0);
              row.push({ content: r.avgD.toFixed(2), styles: { fontStyle: 'bold' } });
              for(let k=0; k<participantCount; k++) row.push(ks[k] || 0);
              row.push({ content: r.avgK.toFixed(2), styles: { fontStyle: 'bold' } });
              row.push({ content: r.score.toFixed(2), styles: { fontStyle: 'bold' } }, r.riskLevel.label);
              return row;
            }),
            cWidths: { 
              0: { cellWidth: 7 }, 
              1: { cellWidth: 40 }
            }
          },
          { 
            t: 'IV. EVALUASI RISIKO (SISA)', 
            h: [
              [
                { content: 'NO', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'RISIKO TERIDENTIFIKASI', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'RISIKO AWAL', colSpan: 4, styles: { halign: 'center', fillColor: [51, 65, 85] } },
                { content: 'PENGENDALIAN SAAT INI', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'CELAH (GAP)', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'RISIKO SISA (RESIDUAL)', colSpan: 4, styles: { halign: 'center', fillColor: [51, 65, 85] } }
              ],
              [
                { content: 'D', styles: { halign: 'center' } },
                { content: 'K', styles: { halign: 'center' } },
                { content: 'SKOR', styles: { halign: 'center' } },
                { content: 'LEVEL', styles: { halign: 'center' } },
                { content: 'D', styles: { halign: 'center' } },
                { content: 'S', styles: { halign: 'center' } },
                { content: 'SKOR', styles: { halign: 'center' } },
                { content: 'LEVEL', styles: { halign: 'center' } }
              ]
            ],
            b: risks.map((r, i) => [
              i + 1, 
              `${r.risikoKode}: ${r.risikoUraian}`, 
              r.avgD.toFixed(2), 
              r.avgK.toFixed(2), 
              r.score.toFixed(2), 
              r.riskLevel.label,
              r.rtpControl,
              r.rtpGap,
              r.resD.toFixed(2), 
              r.resK.toFixed(2), 
              r.resScore.toFixed(2), 
              r.resLevel.label
            ]),
            cWidths: { 0: { cellWidth: 7 }, 1: { cellWidth: 35 }, 4: { cellWidth: 15 } }
          },
          { 
            t: 'V. RENCANA PENANGANAN RISIKO (RTP)', 
            h: [['NO', 'RISIKO (RESIDUAL)', 'PENGENDALIAN SISA', 'CELAH SISA', 'RENCANA TINDAK (RTP) BARU', 'PJ', 'DEADLINE']], 
            b: risks.filter(r => r.resScore > 0 && r.resLevel.level >= 3).map((r, i) => [
              i + 1, 
              `${r.risikoKode}: ${r.risikoUraian}`, 
              r.rtpControl, 
              r.rtpGap, 
              r.rtpAction, 
              r.rtpPJ, 
              r.rtpDeadline
            ]),
            cWidths: { 
              0: { cellWidth: 7 }, 
              1: { cellWidth: 30 }, 
              2: { cellWidth: 30 }, 
              3: { cellWidth: 30 }, 
              4: { cellWidth: 35 } 
            }
          }
        ];

        pdfPages.forEach(ph => { 
          docPDF.addPage(); 
          hdrPDF(ph.t); 
          autoTable(docPDF, { 
            ...tS, 
            head: ph.h, 
            body: ph.b,
            columnStyles: ph.cWidths as any
          }); 
        });

        // Add remaining pages that were potentially missing or cut off
        const extraPages = [
          { 
            t: 'VI. KOMUNIKASI PENGENDALIAN', 
            h: [['NO', 'KEGIATAN PENGENDALIAN', 'MEDIA / SARANA', 'PENYEDIA', 'PENERIMA', 'RENCANA', 'REALISASI', 'KET']], 
            b: risks.filter(r => r.resScore > 0 && r.resLevel.level >= 3).map((r, i) => [
              i + 1, 
              r.rtpAction || '-', 
              r.commMedia || '-', 
              r.commProvider || '-', 
              r.commReceiver || '-', 
              r.commPlanTime || '-', 
              r.commRealTime || '-', 
              r.commNotes || '-'
            ]),
            cWidths: { 0: { cellWidth: 8 }, 1: { cellWidth: 40 } }
          },
          { 
            t: 'VII. RENCANA MONITORING PI', 
            h: [['NO', 'KEGIATAN PENGENDALIAN', 'METODE PEMANTAUAN', 'PJ', 'RENCANA', 'REALISASI', 'KETERANGAN']], 
            b: risks.filter(r => r.resScore > 0 && r.resLevel.level >= 3).map((r, i) => [
              i + 1, 
              r.rtpAction || '-', 
              r.monMethod || '-', 
              r.monPJ || '-', 
              r.monPlanTime || '-', 
              r.monRealTime || '-', 
              r.monNotes || '-'
            ]),
            cWidths: { 0: { cellWidth: 8 }, 1: { cellWidth: 40 } }
          },
          { 
            t: 'IX. MONITORING KETERJADIAN RISIKO', 
            h: [
              [
                { content: 'NO', rowSpan: 2 }, 
                { content: 'URAIAN RISIKO', rowSpan: 2 }, 
                { content: 'KODE', rowSpan: 2 },
                { content: 'KEJADIAN RISIKO', colSpan: 3, styles: { halign: 'center' } },
                { content: 'KET', rowSpan: 2 },
                { content: 'RTP', rowSpan: 2 },
                { content: 'RENCANA', rowSpan: 2 },
                { content: 'REALISASI', rowSpan: 2 },
                { content: 'KET', rowSpan: 2 }
              ],
              ['TANGGAL', 'SEBAB', 'DAMPAK']
            ], 
            b: risks.map((r, i) => [
              i + 1, 
              r.risikoUraian, 
              r.risikoKode, 
              r.eventDate || '-', 
              r.eventCause || '-', 
              r.eventImpact || '-', 
              r.eventNotes || '-', 
              r.rtpAction || '-', 
              r.rtpPlanDate || '-', 
              r.rtpRealDate || '-', 
              r.rtpNotesContent || '-'
            ]),
            cWidths: { 0: { cellWidth: 7 }, 1: { cellWidth: 30 } }
          }
        ];

        extraPages.forEach(ph => {
          docPDF.addPage();
          hdrPDF(ph.t);
          autoTable(docPDF, { ...tS, head: ph.h, body: ph.b });
        });

        docPDF.addPage(); 
        hdrPDF('VIII. PETA RISIKO (HEATMAP)');
        
        try {
          const startX = m;
          const startY = 30;
          const gridSize = 100;
          const cellSize = gridSize / 5;
          
          // Draw Heatmap Grid
          for (let row = 1; row <= 5; row++) {
            for (let col = 1; col <= 5; col++) {
              const rL = getRiskLevel(6 - row, col);
              let fill = [220, 220, 220]; // Default
              if (rL.label === 'Sangat Tinggi') fill = [220, 38, 38]; // red-600
              else if (rL.label === 'Tinggi') fill = [245, 158, 11]; // amber-500
              else if (rL.label === 'Sedang') fill = [250, 204, 21]; // yellow-400
              else if (rL.label === 'Rendah') fill = [34, 197, 94]; // green-500
              
              docPDF.setFillColor(fill[0], fill[1], fill[2]);
              docPDF.rect(startX + (col - 1) * cellSize, startY + (row - 1) * cellSize, cellSize, cellSize, 'F');
              docPDF.setDrawColor(255, 255, 255);
              docPDF.rect(startX + (col - 1) * cellSize, startY + (row - 1) * cellSize, cellSize, cellSize, 'S');
            }
          }
          
          // Axis Labels
          docPDF.setFontSize(8).setFont('helvetica', 'bold').setTextColor(100, 100, 100);
          docPDF.text('IMPACT (DAMPAK)', startX - 5, startY + gridSize / 2, { angle: 90, align: 'center' });
          docPDF.text('PROBABILITY (KEMUNGKINAN)', startX + gridSize / 2, startY + gridSize + 8, { align: 'center' });
          
          // Numbers on Axis
          for (let i = 1; i <= 5; i++) {
            docPDF.text(`${6 - i}`, startX - 3, startY + (i - 1) * cellSize + cellSize / 2 + 2);
            docPDF.text(`${i}`, startX + (i - 1) * cellSize + cellSize / 2, startY + gridSize + 4);
          }

          // Plot Risks as Dots
          risks.forEach((r, idx) => {
            const d = Math.max(1, Math.min(5, Math.ceil(r.avgD)));
            const k = Math.max(1, Math.min(5, Math.ceil(r.avgK)));
            
            const dotX = startX + (k - 1) * cellSize + cellSize / 2;
            const dotY = startY + (5 - d) * cellSize + cellSize / 2;
            
            docPDF.setFillColor(15, 23, 42); // slate-900
            docPDF.circle(dotX, dotY, 3, 'F');
            docPDF.setTextColor(255, 255, 255);
            docPDF.setFontSize(6);
            docPDF.text(`${idx + 1}`, dotX, dotY + 1.5, { align: 'center' });
          });

          // Right side: List and Legend
          const listX = startX + gridSize + 15;
          docPDF.setTextColor(50, 50, 50);
          docPDF.setFontSize(10).setFont('helvetica', 'bold').text('DAFTAR TITIK RISIKO', listX, startY + 5);
          
          const listRows = risks.map((r, i) => [i + 1, r.risikoKode, r.risikoUraian]);
          autoTable(docPDF, {
            startY: startY + 8,
            margin: { left: listX },
            tableWidth: pW - listX - m,
            styles: { fontSize: 7, cellPadding: 1 },
            head: [['No', 'Kode', 'Uraian']],
            body: listRows,
            columnStyles: { 0: { cellWidth: 8 }, 1: { cellWidth: 15 } }
          });

          // Legend
          let legendY = (docPDF as any).lastAutoTable.finalY + 10;
          if (legendY > pH - 30) { docPDF.addPage(); legendY = 30; }
          docPDF.setFontSize(9).setFont('helvetica', 'bold').text('KETERANGAN LEVEL', listX, legendY);
          
          const legends = [
            { l: 'Sangat Tinggi', c: [220, 38, 38] },
            { l: 'Tinggi', c: [245, 158, 11] },
            { l: 'Sedang', c: [250, 204, 21] },
            { l: 'Rendah', c: [34, 197, 94] }
          ];
          
          legends.forEach((lg, i) => {
            const ly = legendY + 5 + (i * 5);
            docPDF.setFillColor(lg.c[0], lg.c[1], lg.c[2]);
            docPDF.rect(listX, ly - 3, 4, 4, 'F');
            docPDF.setTextColor(80, 80, 80);
            docPDF.setFontSize(8).setFont('helvetica', 'normal').text(lg.l, listX + 6, ly);
          });

        } catch(e) {
          console.error('Heatmap PDF generation failed:', e);
          docPDF.setFontSize(12).text('Gagal memproses visualisasi heatmap ke PDF.', m, 40);
        }

        docPDF.save(`${filename}.pdf`);
      }
      alert('Ekspor Berhasil!');
    } catch (err: any) {
      console.error('Export error:', err);
      alert('Gagal mengekspor data: ' + err.message);
    } finally {
      setIsExporting(false);
    }
  };

  const handleLogin = async (userData?: any) => {
    if (userData) {
      // Normalize role casing to match UI expectations
      let role = userData.role || 'User';
      if (role.toLowerCase() === 'operator') role = 'Operator';
      if (role.toLowerCase() === 'administrator') role = 'Administrator';
      
      const u = { username: userData.username, role: role, uid: userData.uid };
      localStorage.setItem('isman_user', JSON.stringify(u));
      setUser(u);
      return;
    }
    try {
      localStorage.removeItem('isman_user');
      await signInWithPopup(auth, googleProvider);
    } catch (error) {
      console.error("Login failed", error);
    }
  };

  const handleLogout = () => {
    setShowLogoutConfirm(true);
  };

  const confirmLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error("Logout error:", error);
    }
    // Always clear local state regardless of Firebase auth state
    localStorage.removeItem('isman_user');
    setUser(null);
    setViewingUser(null);
    setShowLogoutConfirm(false);
    setActiveMenu(0);
    // Hard reset for cleanest state
    window.location.href = window.location.origin;
  };

  const handleChangePassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setChangePassError('');
    setChangePassSuccess('');

    if (newPassword !== confirmNewPassword) {
      setChangePassError('Konfirmasi password tidak cocok!');
      return;
    }

    if (newPassword.length < 6) {
      setChangePassError('Password minimal 6 karakter!');
      return;
    }

    try {
      if (!user?.uid) return;
      await updateDoc(doc(db, 'accounts', user.uid), {
        password: newPassword,
        updatedAt: new Date().toISOString()
      });
      setChangePassSuccess('Password berhasil diperbarui!');
      setNewPassword('');
      setConfirmNewPassword('');
      setTimeout(() => {
        setShowChangePassword(false);
        setChangePassSuccess('');
      }, 2000);
    } catch (err: any) {
      setChangePassError('Gagal memperbarui password: ' + err.message);
    }
  };

  if (loading) {
    return (
      <div className="h-screen bg-slate-900 flex items-center justify-center">
        <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin"></div>
      </div>
    );
  }

  if (!user) {
    return <LoginPage onLogin={handleLogin} accounts={accounts} />;
  }

  if (user && !selectedRiskType && user.role !== 'Administrator') {
    return (
      <div className="min-h-screen bg-slate-900 flex flex-col items-center justify-center p-6 bg-[radial-gradient(circle_at_top_right,_var(--tw-gradient-stops))] from-blue-900/20 via-slate-900 to-slate-900">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="w-full max-w-2xl bg-white rounded-2xl shadow-2xl overflow-hidden border border-slate-100"
        >
          <div className="p-8 text-center border-b border-slate-100 bg-slate-50">
            <h1 className="text-3xl font-black text-blue-600 tracking-tighter uppercase italic">ISMAN</h1>
            <p className="text-slate-600 text-[10px] font-black uppercase tracking-tight -mt-1 mb-4">Integrated Risk Management System</p>
            <h2 className="text-xl font-bold text-slate-900 uppercase italic">Pilih Jenis Risiko</h2>
            <p className="text-slate-500 text-xs mt-1">Silahkan pilih kategori penilaian risiko untuk dilanjutkan</p>
          </div>
          
          <div className="p-8 grid grid-cols-1 md:grid-cols-2 gap-6 pb-12">
            <button 
              id="select-strategis"
              onClick={() => setSelectedRiskType('strategis')}
              className="group p-8 rounded-xl border-2 border-slate-100 hover:border-blue-500 hover:bg-blue-50 transition-all text-left flex flex-col items-center text-center gap-4"
            >
              <div className="w-16 h-16 rounded-full bg-blue-100 flex items-center justify-center group-hover:scale-110 transition-transform">
                <Target className="w-8 h-8 text-blue-600" />
              </div>
              <div>
                <h3 className="font-black text-slate-900 text-lg uppercase">Risiko Strategis</h3>
                <p className="text-xs text-slate-500 mt-2">Penilaian risiko terkait pencapaian sasaran strategis instansi</p>
              </div>
            </button>

            <button 
              id="select-operasional"
              onClick={() => setSelectedRiskType('operasional')}
              className="group p-8 rounded-xl border-2 border-slate-100 hover:border-emerald-500 hover:bg-emerald-50 transition-all text-left flex flex-col items-center text-center gap-4"
            >
              <div className="w-16 h-16 rounded-full bg-emerald-100 flex items-center justify-center group-hover:scale-110 transition-transform">
                <Settings2 className="w-8 h-8 text-emerald-600" />
              </div>
              <div>
                <h3 className="font-black text-slate-900 text-lg uppercase">Risiko Operasional</h3>
                <p className="text-xs text-slate-500 mt-2">Penilaian risiko terkait proses bisnis dan kegiatan operasional</p>
              </div>
            </button>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-slate-50 font-sans text-slate-900 overflow-hidden">
      {/* Sidebar */}
      <motion.aside 
        initial={false}
        animate={{ width: isSidebarOpen ? 300 : 80 }}
        className="bg-slate-900 text-slate-300 flex flex-col border-r border-slate-800"
      >
        <div className="p-6 flex items-center justify-between">
          {isSidebarOpen && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="font-bold text-xl tracking-tight text-white flex items-center gap-2"
            >
              <div className="flex flex-col">
                <span className="bg-gradient-to-r from-white to-amber-400 bg-clip-text text-transparent font-black tracking-tighter leading-none">ISMAN</span>
                <span className="text-[6px] text-slate-500 font-bold tracking-widest leading-none mt-0.5">INTEGRATED RISK MANAGEMENT</span>
              </div>
            </motion.div>
          )}
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className="p-2 hover:bg-slate-800 rounded-lg transition-colors"
          >
            {isSidebarOpen ? <X size={20} /> : <Menu size={20} />}
          </button>
        </div>

        <nav className="flex-1 overflow-y-auto px-4 py-2 space-y-1">
          {menuItems.map((item) => (
            <button
              key={item.id}
              onClick={() => setActiveMenu(item.id)}
              className={`w-full flex items-center gap-3 p-3 rounded-lg transition-all text-sm font-medium ${
                activeMenu === item.id 
                  ? 'bg-blue-600 text-white shadow-lg shadow-blue-900/20' 
                  : 'hover:bg-slate-800 hover:text-white text-slate-400'
              }`}
            >
              <item.icon size={20} className="shrink-0" />
              {isSidebarOpen && (
                <span className="truncate text-left leading-tight">
                  {item.title}
                </span>
              )}
            </button>
          ))}
        </nav>

        <div className="p-4 border-t border-slate-800 space-y-3">
          <div className="flex items-center gap-3 overflow-hidden">
            <div className="w-8 h-8 rounded-full bg-blue-500 flex items-center justify-center text-white font-bold shrink-0">
              {user.username.charAt(0).toUpperCase()}
            </div>
            {isSidebarOpen && (
              <div className="overflow-hidden flex-1">
                <p className="text-sm font-black text-white truncate uppercase tracking-tighter">{user.username}</p>
                <p className="text-xs text-slate-500 truncate">{user.role}</p>
              </div>
            )}
          </div>
          
          <div className={`flex gap-2 ${isSidebarOpen ? 'flex-row' : 'flex-col items-center'}`}>
            <button 
              onClick={() => setShowChangePassword(true)}
              className="p-2 text-slate-500 hover:text-blue-400 hover:bg-slate-800 rounded-lg transition-all flex-1 flex justify-center items-center gap-2"
              title="Ganti Password"
            >
              <Lock size={18} />
              {isSidebarOpen && <span className="text-[10px] uppercase font-black tracking-widest">Sandi</span>}
            </button>
            <button 
              onClick={handleLogout}
              className="p-2 text-slate-500 hover:text-red-400 hover:bg-slate-800 rounded-lg transition-all flex-1 flex justify-center items-center gap-2"
              title="Keluar / Logout"
            >
              <RefreshCcw size={18} className="rotate-180" />
              {isSidebarOpen && <span className="text-[10px] uppercase font-black tracking-widest">Keluar</span>}
            </button>
          </div>
        </div>
      </motion.aside>

      {/* Main Content */}
       <main className="flex-1 flex flex-col overflow-hidden">
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 shrink-0">
          <div className="flex items-center gap-4">
            <h2 className="text-sm font-semibold text-slate-500 uppercase tracking-widest flex items-center gap-2">
              {menuItems.find(m => m.id === activeMenu)?.title}
              {selectedRiskType && (
                <div className="flex items-center gap-2">
                  <span className="text-blue-600">
                    ({selectedRiskType === 'operasional' ? 'OPERASIONAL' : 'STRATEGIS'})
                  </span>
                  {user.role !== 'Administrator' && (
                    <button 
                      onClick={() => setSelectedRiskType(null)}
                      className="p-1 hover:bg-blue-50 text-blue-400 hover:text-blue-600 rounded-lg transition-all"
                      title="Ganti Mode Risiko"
                    >
                      <RotateCw size={14} />
                    </button>
                  )}
                </div>
              )}
            </h2>
            {viewingUser && user && viewingUser.uid !== user.uid && (
              <div className="flex items-center gap-2 bg-blue-50 border border-blue-100 px-3 py-1 rounded-full">
                <div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse"></div>
                <span className="text-[10px] font-black text-blue-600 uppercase tracking-tighter">Viewing As: {viewingUser.username}</span>
                <button 
                  onClick={() => setViewingUser(user)}
                  className="ml-1 text-blue-400 hover:text-blue-600 font-black text-[10px] uppercase"
                >
                  [STOP]
                </button>
              </div>
            )}
          </div>
          <div className="flex items-center gap-4">
            <div className="flex bg-slate-100 p-1 rounded-lg gap-1 border border-slate-200">
               <button 
                 disabled={isExporting}
                 onClick={() => handleExportAll('excel')}
                 className="flex items-center gap-2 px-3 py-1.5 text-xs font-bold uppercase tracking-wider text-emerald-700 hover:bg-emerald-50 rounded transition-all disabled:opacity-50"
               >
                 <FileSpreadsheet size={14} /> EXCEL
               </button>
               <button 
                 disabled={isExporting}
                 onClick={() => handleExportAll('pdf')}
                 className="flex items-center gap-2 px-3 py-1.5 text-xs font-bold uppercase tracking-wider text-red-700 hover:bg-red-50 rounded transition-all disabled:opacity-50"
               >
                 <FileJson size={14} /> PDF
               </button>
            </div>

            {!isActuallyReadOnly && (
              <>
                {activeMenu === 2 ? null : activeMenu !== 0 && activeMenu !== 8 && activeMenu !== 10 && (
                  <button className="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 transition-colors flex items-center gap-2">
                    <Plus size={16} />
                    Tambah Data
                  </button>
                )}
              </>
            )}
          </div>
        </header>

        <section className="flex-1 overflow-y-auto p-8">
          <AnimatePresence mode="wait">
            <motion.div
              key={activeMenu}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
              className="max-w-7xl mx-auto"
            >
              {activeMenu === 0 ? (
                <MonitoringProgressView user={user!} onSelectUser={(u, rt) => { setViewingUser(u); setSelectedRiskType(rt); setActiveMenu(1); }} />
              ) : activeMenu === 1 ? (
                <ContextSettingView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 2 ? (
                <RiskIdentificationView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 3 ? (
                <RiskAnalysisView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 4 ? (
                <RiskResidualView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 5 ? (
                <RiskTreatmentView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 6 ? (
                <MonitoringCommunicationView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 7 ? (
                <MonitoringPlanPIView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 8 ? (
                <RiskMapView user={viewingUser!} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 9 ? (
                <RiskOccurrenceMonitoringView user={viewingUser!} isReadOnly={isActuallyReadOnly} riskType={selectedRiskType || 'strategis'} />
              ) : activeMenu === 10 ? (
                <FinalDocumentView user={viewingUser!} isAdmin={user!.role === 'Administrator'} isOperator={user!.role === 'Operator'} />
              ) : activeMenu === 11 ? (
                <AccountManagementView accounts={accounts} />
              ) : (
                <PlaceholderView title={menuItems.find(m => m.id === activeMenu)?.title || ''} />
              )}
            </motion.div>
          </AnimatePresence>
        </section>
      </main>

      {/* Modals */}
      <AnimatePresence>
        {showLogoutConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center"
            >
              <div className="w-16 h-16 bg-red-50 rounded-full flex items-center justify-center mx-auto mb-4">
                <RefreshCcw size={32} className="text-red-500 rotate-180" />
              </div>
              <h3 className="text-xl font-bold text-slate-900 mb-2">Konfirmasi Keluar</h3>
              <p className="text-slate-500 text-sm mb-8 font-medium italic leading-relaxed">
                "Apakah anda yakin ingin keluar dari sistem ISMAN ini?"
              </p>
              <div className="flex gap-3">
                <button 
                  onClick={() => setShowLogoutConfirm(false)}
                  className="flex-1 px-4 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold text-sm hover:bg-slate-200 transition-colors"
                >
                  Batal
                </button>
                <button 
                  onClick={confirmLogout}
                  className="flex-1 px-4 py-3 bg-red-600 text-white rounded-xl font-bold text-sm hover:bg-red-700 transition-colors"
                >
                  Ya, Keluar
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {showChangePassword && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl p-8 max-w-md w-full"
            >
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl font-bold text-slate-900 flex items-center gap-2 italic uppercase tracking-tight">
                  <Lock size={20} className="text-blue-600" />
                  Ganti Password
                </h3>
                <button onClick={() => setShowChangePassword(false)} className="p-2 hover:bg-slate-100 rounded-lg text-slate-400">
                  <X size={20} />
                </button>
              </div>

              <form onSubmit={handleChangePassword} className="space-y-4">
                {changePassError && (
                  <div className="p-3 bg-red-50 border border-red-100 text-red-500 text-xs font-bold rounded-lg flex items-center gap-2">
                    <ShieldAlert size={14} /> {changePassError}
                  </div>
                )}
                {changePassSuccess && (
                  <div className="p-3 bg-emerald-50 border border-emerald-100 text-emerald-600 text-xs font-bold rounded-lg flex items-center gap-2">
                    <Check size={14} /> {changePassSuccess}
                  </div>
                )}

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Password Baru</label>
                  <input 
                    type="password"
                    value={newPassword}
                    onChange={(e) => setNewPassword(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-4 text-sm outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all font-bold"
                    placeholder="Minimal 6 karakter"
                    required
                  />
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Konfirmasi Password Baru</label>
                  <input 
                    type="password"
                    value={confirmNewPassword}
                    onChange={(e) => setConfirmNewPassword(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-4 text-sm outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all font-bold"
                    placeholder="Ulangi password baru"
                    required
                  />
                </div>

                <button 
                  type="submit"
                  className="w-full bg-blue-600 text-white font-black uppercase tracking-widest py-4 rounded-xl shadow-lg shadow-blue-600/30 hover:bg-blue-700 hover:-translate-y-0.5 active:translate-y-0 transition-all text-sm mt-4"
                >
                  Simpan Password Baru
                </button>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

// --- Login Component ---

function MonitoringProgressView({ user, onSelectUser }: { user: any, onSelectUser: (u: any, rt: 'strategis' | 'operasional') => void }) {
  const [accounts, setAccounts] = useState<any[]>([]);
  const [stats, setStats] = useState<Record<string, any>>({});
  const [activeTooltip, setActiveTooltip] = useState<{accId: string, menuKey: string, msg: string} | null>(null);
  const [loading, setLoading] = useState(true);
  const [filterRiskType, setFilterRiskType] = useState<'strategis' | 'operasional'>('strategis');
  const [risksState, setRisksState] = useState<any[]>([]);
  const [contextsState, setContextsState] = useState<any[]>([]);
  const [finalDocsState, setFinalDocsState] = useState<any[]>([]);
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc' | 'name'>('name');
  const [isExporting, setIsExporting] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  const handleBulkExcel = async () => {
    setIsExporting(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const riskType = filterRiskType;
      const typeLabel = riskType === 'strategis' ? 'STRATEGIS' : 'OPERASIONAL';
      const themeHex = riskType === 'strategis' ? '1E3A8A' : '059669';

      const styleHdr = (sheet: any, row: number, cols: number) => {
        for (let c = 1; c <= cols; c++) {
          const cell = sheet.getCell(row, c);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: themeHex } };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      };

      const styleTbl = (sheet: any, startRow: number, rowCount: number, colCount: number) => {
        for (let r = startRow; r < startRow + rowCount; r++) {
          for (let c = 1; c <= colCount; c++) {
            const cell = sheet.getCell(r, c);
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', wrapText: true };
          }
        }
      };

      // 1. Ringkasan Sheet
      const ws1 = workbook.addWorksheet('RINGKASAN');
      const ws1Hdr = ['NO', 'NAMA OPD', 'TIPE', 'PROGRESS %'];
      ws1.getRow(1).values = ws1Hdr;
      styleHdr(ws1, 1, 4);
      sortedAccounts.forEach((acc, i) => {
        const s = stats[acc.uid] || { percent: 0 };
        ws1.addRow([i + 1, acc.username, typeLabel, `${s.percent}%`]);
      });
      styleTbl(ws1, 2, sortedAccounts.length, 4);
      ws1.columns = [{ width: 5 }, { width: 50 }, { width: 15 }, { width: 15 }];

      // 2. Identifikasi Sheet
      const ws2 = workbook.addWorksheet('IDENTIFIKASI');
      const ws2Hdr = ['NAMA OPD', 'NO', 'KODE', 'URAIAN RISIKO', 'TUJUAN / SASARAN', 'INDIKATOR', 'PEMILIK', 'SEBAB', 'SUMBER', 'C/UC', 'DAMPAK', 'PIHAK'];
      ws2.getRow(1).values = ws2Hdr;
      styleHdr(ws2, 1, ws2Hdr.length);
      let r2 = 2;
      sortedAccounts.forEach(acc => {
        const userRisks = risksState.filter(data => data.createdByUid === acc.uid && (data.riskType || 'strategis') === riskType);
        userRisks.forEach((risiko, i) => {
          const srs = (risiko.subRows && risiko.subRows.length > 0) 
            ? risiko.subRows 
            : [{ 
                sebabUraian: risiko.sebabUraian || '', 
                sebabSumber: risiko.sebabSumber || '', 
                control: risiko.control || '', 
                dampakUraian: risiko.dampakUraian || '', 
                dampakPihak: risiko.dampakPihak || '' 
              }];
          srs.forEach((sub: any) => {
            ws2.addRow([acc.username, i + 1, risiko.risikoKode, risiko.risikoUraian, risiko.tujuan, risiko.indikator, risiko.pemilik, sub.sebabUraian, sub.sebabSumber, sub.control, sub.dampakUraian, sub.dampakPihak]);
            r2++;
          });
        });
      });
      styleTbl(ws2, 2, r2 - 2, ws2Hdr.length);
      ws2.columns = ws2Hdr.map((_, i) => ({ width: i === 0 || i === 3 || i === 4 ? 40 : 15 }));

      // 3. Analisis Sheet
      const ws3 = workbook.addWorksheet('ANALISIS');
      const ws3Hdr = ['NAMA OPD', 'NO', 'KODE', 'URAIAN RISIKO', 'RATA DAMPAK', 'RATA KEMUNGKINAN', 'SKOR', 'LEVEL'];
      ws3.getRow(1).values = ws3Hdr;
      styleHdr(ws3, 1, ws3Hdr.length);
      let r3 = 2;
      sortedAccounts.forEach(acc => {
        const userRisks = risksState.filter(data => data.createdByUid === acc.uid && (data.riskType || 'strategis') === riskType);
        userRisks.forEach((risiko, i) => {
          const dVal = risiko.dampakScores ? risiko.dampakScores.reduce((a: number, b: number) => a + b, 0) / risiko.dampakScores.length : 0;
          const kVal = risiko.kemungkinanScores ? risiko.kemungkinanScores.reduce((a: number, b: number) => a + b, 0) / risiko.kemungkinanScores.length : 0;
          const lv = getRiskLevel(dVal, kVal);
          ws3.addRow([acc.username, i + 1, risiko.risikoKode, risiko.risikoUraian, dVal.toFixed(2), kVal.toFixed(2), (dVal * kVal).toFixed(2), lv.label]);
          r3++;
        });
      });
      styleTbl(ws3, 2, r3 - 2, ws3Hdr.length);
      ws3.columns = ws3Hdr.map((_, i) => ({ width: i === 0 || i === 3 ? 40 : 15 }));

      // 4. Residual Sheet
      const ws4 = workbook.addWorksheet('RESIDUAL & RTP');
      const ws4Hdr = ['NAMA OPD', 'NO', 'KODE', 'URAIAN RISIKO', 'RESIDUAL D', 'RESIDUAL K', 'RESKORE', 'LEVEL', 'KENDALI', 'GAP', 'AKSI RTP', 'PJ', 'DEADLINE'];
      ws4.getRow(1).values = ws4Hdr;
      styleHdr(ws4, 1, ws4Hdr.length);
      let r4 = 2;
      sortedAccounts.forEach(acc => {
        const userRisks = risksState.filter(data => data.createdByUid === acc.uid && (data.riskType || 'strategis') === riskType);
        userRisks.forEach((risiko, i) => {
          const rd = parseFloat(risiko.residualDampak || 0);
          const rk = parseFloat(risiko.residualKemungkinan || 0);
          const lv = getRiskLevel(rd, rk);
          ws4.addRow([acc.username, i + 1, risiko.risikoKode, risiko.risikoUraian, rd.toFixed(2), rk.toFixed(2), (rd * rk).toFixed(2), lv.label, risiko.rtpControl, risiko.rtpGap, risiko.rtpAction, risiko.rtpPJ, risiko.rtpDeadline]);
          r4++;
        });
      });
      styleTbl(ws4, 2, r4 - 2, ws4Hdr.length);
      ws4.columns = ws4Hdr.map((_, i) => ({ width: i === 0 || i === 3 || i === 10 ? 40 : 15 }));

      saveAs(new Blob([await workbook.xlsx.writeBuffer()]), `MASSAL_DATA_${typeLabel}_${new Date().toISOString().split('T')[0]}.xlsx`);
    } catch (err: any) {
      console.error('Bulk Excel error:', err);
      alert('Gagal Ekspor Massal: ' + err.message);
    } finally {
      setIsExporting(false);
    }
  };

  const handleBulkPDF = async () => {
    setIsExporting(true);
    try {
      const riskType = filterRiskType;
      const typeLabel = riskType === 'strategis' ? 'STRATEGIS' : 'OPERASIONAL';
      const themeColor: [number, number, number] = riskType === 'strategis' ? [30, 58, 138] : [5, 150, 105];

      const doc = new jsPDF('l', 'mm', 'a4');
      const pW = doc.internal.pageSize.getWidth();
      
      doc.setFontSize(16).setFont('helvetica', 'bold');
      doc.text(`LAPORAN KONSOLIDASI RISIKO ${typeLabel}`, pW / 2, 15, { align: 'center' });
      doc.setFontSize(8).setFont('helvetica', 'normal');
      doc.text(`Dicetak pada: ${new Date().toLocaleString('id-ID')}`, pW / 2, 22, { align: 'center' });

      // Summary Table
      autoTable(doc, {
        startY: 30,
        head: [['NO', 'NAMA OPD', 'PROGRESS %', 'JUMLAH RISIKO']],
        body: sortedAccounts.map((acc, i) => {
          const s = stats[acc.uid] || { percent: 0, total: 0 };
          return [i + 1, acc.username, `${s.percent}%`, s.total];
        }),
        headStyles: { fillColor: themeColor },
        styles: { fontSize: 8, halign: 'center' },
        columnStyles: { 1: { halign: 'left' } }
      });

      // Consolidated Identification Table
      doc.addPage();
      doc.setFontSize(12).text(`KONSOLIDASI IDENTIFIKASI RISIKO - ${typeLabel}`, 15, 15);
      
      const body: any[] = [];
      sortedAccounts.forEach(acc => {
        const userRisks = risksState.filter(data => data.createdByUid === acc.uid && (data.riskType || 'strategis') === riskType);
        userRisks.forEach((risiko) => {
          const srs = (risiko.subRows && risiko.subRows.length > 0) 
            ? risiko.subRows 
            : [{ 
                sebabUraian: risiko.sebabUraian || '', 
                dampakUraian: risiko.dampakUraian || '' 
              }];

          srs.forEach((sub: any) => {
            body.push([
              acc.username,
              risiko.risikoKode,
              risiko.risikoUraian,
              sub.sebabUraian || '-',
              sub.dampakUraian || '-'
            ]);
          });
        });
      });

      autoTable(doc, {
        startY: 20,
        head: [['NAMA OPD', 'KODE', 'URAIAN RISIKO', 'SEBAB', 'DAMPAK']],
        body: body,
        headStyles: { fillColor: themeColor },
        styles: { fontSize: 6 },
        columnStyles: { 
          0: { cellWidth: 35 }, 
          1: { cellWidth: 20 },
          2: { cellWidth: 60 },
          3: { cellWidth: 70 },
          4: { cellWidth: 70 }
        }
      });

      // Residual Summary Table
      doc.addPage();
      doc.setFontSize(12).text(`KONSOLIDASI EVALUASI & RTP - ${typeLabel}`, 15, 15);
      
      const resBody: any[] = [];
      sortedAccounts.forEach(acc => {
        const userRisks = risksState.filter(data => data.createdByUid === acc.uid && (data.riskType || 'strategis') === riskType);
        userRisks.forEach((risiko) => {
          const rd = parseFloat(risiko.residualDampak || 0);
          const rk = parseFloat(risiko.residualKemungkinan || 0);
          const lv = getRiskLevel(rd, rk);
          resBody.push([
            acc.username,
            risiko.risikoKode,
            (rd * rk).toFixed(2),
            lv.label,
            risiko.rtpAction || '-'
          ]);
        });
      });

      autoTable(doc, {
        startY: 20,
        head: [['NAMA OPD', 'KODE', 'SKOR SISA', 'LEVEL SISA', 'RENCANA TINDAK (RTP)']],
        body: resBody,
        headStyles: { fillColor: themeColor },
        styles: { fontSize: 7 },
        columnStyles: { 0: { cellWidth: 50 }, 4: { cellWidth: 100 } }
      });

      doc.save(`MASSAL_LAPORAN_${typeLabel}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err: any) {
      console.error('Bulk PDF error:', err);
      alert('Gagal Ekspor Massal PDF: ' + err.message);
    } finally {
      setIsExporting(false);
    }
  };

  useEffect(() => {
    setLoading(true);
    // Listen to accounts
    const unsubAccounts = onSnapshot(collection(db, 'accounts'), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ ...doc.data(), uid: doc.id, id: doc.id })) as any[];
      const filtered = user.role === 'Operator' ? data.filter(a => a.role !== 'Administrator') : data;
      filtered.sort((a, b) => (a.username || '').localeCompare(b.username || ''));
      setAccounts(filtered);
      setLoading(false);
    }, (error) => {
      console.error('Accounts load error:', error);
      setLoading(false);
    });

    // Listen to risks
    const unsubRisks = onSnapshot(collection(db, 'risk_identification'), (snapshot) => {
      const risks = snapshot.docs.map(d => ({ ...d.data(), id: d.id }));
      setRisksState(risks);
    }, (error) => {
      console.error('Risks load error:', error);
    });

    // Listen to contexts
    const unsubContexts = onSnapshot(collection(db, 'risk_context'), (snapshot) => {
      const contexts = snapshot.docs.map(d => ({ ...d.data(), id: d.id }));
      setContextsState(contexts);
    }, (error) => {
      console.error('Contexts load error:', error);
    });

    // Listen to final documents
    const unsubFinalDocs = onSnapshot(collection(db, 'final_documents'), (snapshot) => {
      const docs = snapshot.docs.map(d => ({ ...d.data(), id: d.id }));
      setFinalDocsState(docs);
    }, (error) => {
      console.error('FinalDocs load error:', error);
    });

    return () => {
      unsubAccounts();
      unsubRisks();
      unsubContexts();
      unsubFinalDocs();
    };
  }, [user.role]);

  useEffect(() => {
    if (accounts.length === 0) return;

    const allStats: Record<string, any> = {};
    
    accounts.forEach(acc => {
      // Helper to match risk level logic used in components
      const getScoreAndLevel = (d: any, type: 'init' | 'res') => {
        let dVal = 0, kVal = 0;
        if (type === 'init') {
          const activeD = (d.dampakScores || []).filter((v: number) => v > 0);
          const activeK = (d.kemungkinanScores || []).filter((v: number) => v > 0);
          dVal = activeD.length > 0 ? parseFloat((activeD.reduce((a: number, b: number) => a + b, 0) / activeD.length).toFixed(2)) : 0;
          kVal = activeK.length > 0 ? parseFloat((activeK.reduce((a: number, b: number) => a + b, 0) / activeK.length).toFixed(2)) : 0;
        } else {
          dVal = parseFloat(d.residualDampak || 0);
          kVal = parseFloat(d.residualKemungkinan || 0);
        }
        const score = dVal * kVal;
        const level = getRiskLevel(dVal, kVal);
        return { score, level: level.level, label: level.label };
      };

      const baseRisks = risksState.filter(data => {
        const rt = data.riskType || 'strategis';
        return data.createdByUid === acc.uid && rt === filterRiskType;
      });

      // Default sort (for Menu II, III, VI, VII, IX)
      const sortedByDefault = [...baseRisks].sort((a: any, b: any) => {
        if (a.order !== undefined && b.order !== undefined) {
          return a.order - b.order;
        }
        const dateA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
        const dateB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
        return dateA - dateB;
      });
      
      const contextDocId = `risk_context_${acc.uid}_${filterRiskType}`;
      const contextData = contextsState.find(d => d.id === contextDocId);
      const reasons: Record<string, string> = {};

      const getRiskName = (r: any, idx: number) => r.risikoKode?.trim() || `R${idx + 1}`;

      // Menu I: Context
      const missingCtx: string[] = [];
      if (!contextData) {
        missingCtx.push('Data Konteks Belum Ada');
      } else {
        if (!(contextData.tujuanStrategis?.trim())) missingCtx.push('Tujuan');
        if (!contextData.sasaran || contextData.sasaran.length === 0 || !contextData.sasaran.some((s: string) => s.trim())) missingCtx.push('Sasaran');
        if (!contextData.program || contextData.program.length === 0 || !contextData.program.some((p: string) => p.trim())) missingCtx.push('Program');
        if (!contextData.ikuProgram || contextData.ikuProgram.length === 0 || !contextData.ikuProgram.some((i: any) => i.name?.trim())) missingCtx.push('IKU');
        if (!(contextData.ttdNama?.trim())) missingCtx.push('Nama TTD');
        if (!(contextData.ttdNip?.trim())) missingCtx.push('NIP');
        if (!(contextData.ttdJabatan?.trim())) missingCtx.push('Jabatan');
      }
      const hasContext = (!!contextData) && missingCtx.length === 0;
      if (!hasContext) reasons.context = "Kekurangan Menu I: " + missingCtx.join(', ');

      // Menu II: Identifikasi (Sorted by Default)
      const missingIdent: string[] = [];
      if (sortedByDefault.length === 0) {
        missingIdent.push('Belum ada risiko');
      } else {
        sortedByDefault.forEach((r, idx) => {
          const m: string[] = [];
          if (!r.risikoUraian?.trim()) m.push('Uraian');
          if (!r.risikoKode?.trim()) m.push('Kode');
          if (!r.pemilik?.trim()) m.push('Pemilik');
          
          const sRowsRaw = r.subRows || [];
          // Only use subRows if they actually have content, otherwise fallback to root fields
          const hasFilledSubRows = sRowsRaw.some((s: any) => 
            s.sebabUraian?.trim() || 
            s.sebabSumber?.trim() || 
            s.control?.trim() || 
            s.dampakUraian?.trim() || 
            s.dampakPihak?.trim()
          );
          const sRows = hasFilledSubRows ? sRowsRaw : [];

          if (sRows.length === 0) {
            // Fallback for old data or if somehow missing subRows
            if (!r.sebabUraian?.trim()) m.push('Sebab');
            if (!r.sebabSumber?.trim()) m.push('Sumber');
            if (!r.control?.trim()) m.push('Kendali');
            if (!r.dampakUraian?.trim()) m.push('Dampak');
            if (!r.dampakPihak?.trim()) m.push('Pihak');
          } else {
            sRows.forEach((sub: any, sIdx: number) => {
              // Only check subrows that have at least one field partially filled, to avoid nagging about extra empty subrows
              const isPartiallyFilled = sub.sebabUraian?.trim() || sub.dampakUraian?.trim();
              if (isPartiallyFilled) {
                const sm: string[] = [];
                if (!sub.sebabUraian?.trim()) sm.push('Sebab');
                if (!sub.sebabSumber?.trim()) sm.push('Sumber');
                if (!sub.control?.trim()) sm.push('Kendali');
                if (!sub.dampakUraian?.trim()) sm.push('Dampak');
                if (!sub.dampakPihak?.trim()) sm.push('Pihak');
                if (sm.length > 0) {
                  const suffix = sRows.length > 1 ? ` (B${sIdx + 1})` : '';
                  sm.forEach(field => m.push(field + suffix));
                }
              }
            });
          }
          
          if (m.length > 0) missingIdent.push(`${getRiskName(r, idx)}: ${m.join(', ')}`);
        });
      }
      const hasIdentification = (sortedByDefault.length > 0) && missingIdent.length === 0;
      if (!hasIdentification) reasons.identification = "Kekurangan Menu II: " + missingIdent.slice(0, 2).join('; ') + (missingIdent.length > 2 ? '...' : '');

      // Menu III: Analisis (Sorted by Default)
      const missingAnal: string[] = [];
      if (sortedByDefault.length === 0) {
        missingAnal.push('Data Kosong');
      } else {
        sortedByDefault.forEach((r, idx) => {
          const dScores = r.dampakScores || [];
          const kScores = r.kemungkinanScores || [];
          const hasD = dScores.some((s: number) => s > 0);
          const hasK = kScores.some((s: number) => s > 0);
          if (!hasD || !hasK) {
            missingAnal.push(`${getRiskName(r, idx)} blm dinilai`);
          }
        });
      }
      const hasAnalysis = (sortedByDefault.length > 0) && missingAnal.length === 0;
      if (!hasAnalysis) reasons.analysis = "Kekurangan Menu III: " + missingAnal.join(', ');

      // Menu IV: Residual (SORTED BY INIT RISK LEVEL/SCORE)
      const missingRes: string[] = [];
      if (baseRisks.length === 0) {
        missingRes.push('Data Kosong');
      } else {
        const sortedForRes = [...baseRisks].map(r => {
          const stats = getScoreAndLevel(r, 'init');
          return { ...r, stats };
        }).sort((a, b) => {
          if (b.stats.level !== a.stats.level) return b.stats.level - a.stats.level;
          return b.stats.score - a.stats.score;
        });

        sortedForRes.forEach((r, idx) => {
          const m: string[] = [];
          if (!(parseFloat(r.residualDampak || 0) > 0)) m.push('Nilai Dampak');
          if (!(parseFloat(r.residualKemungkinan || 0) > 0)) m.push('Nilai Kemunk.');
          if (!r.rtpControl?.trim()) m.push('RTP Kendali');
          if (m.length > 0) missingRes.push(`${getRiskName(r, idx)}: ${m.join(', ')}`);
        });
      }
      const hasResidual = (baseRisks.length > 0) && missingRes.length === 0;
      if (!hasResidual) reasons.residual = "Kekurangan Menu IV: " + missingRes.slice(0, 2).join('; ');

      // Menu V: Treatment (FILTERED resScore > 0, SORTED BY RES LEVEL/SCORE)
      const missingTreat: string[] = [];
      const treatmentRisks = baseRisks.map(r => {
        const stats = getScoreAndLevel(r, 'res');
        return { ...r, stats };
      }).filter(r => r.stats.score > 0 && r.stats.level >= 3);

      if (treatmentRisks.length === 0 && baseRisks.length > 0) {
        // Technically if no risks have residual score, Menu V is empty, but we might want to flag that Menu IV isn't done
        // If hasResidual is false, treatment might not be reachable.
      } else {
        treatmentRisks.sort((a, b) => {
          if (b.stats.level !== a.stats.level) return b.stats.level - a.stats.level;
          return b.stats.score - a.stats.score;
        });

        treatmentRisks.forEach((r, idx) => {
          const m: string[] = [];
          if (!r.rtpAction?.trim()) m.push('Aksi');
          if (!r.rtpPJ?.trim()) m.push('PJ');
          if (!r.rtpDeadline?.trim()) m.push('Deadline');
          if (m.length > 0) missingTreat.push(`${getRiskName(r, idx)}: ${m.join(', ')}`);
        });
      }
      const hasTreatment = (baseRisks.length > 0) && hasAnalysis && hasResidual && (treatmentRisks.length === 0 || (treatmentRisks.length > 0 && missingTreat.length === 0));
      if (!hasTreatment) reasons.treatment = (baseRisks.length > 0 && hasAnalysis && hasResidual && treatmentRisks.length === 0) ? "Lengkap: Tidak ada risiko yang memerlukan RTP" : "Kekurangan Menu V: " + missingTreat.slice(0, 2).join('; ');

      // Menu VI: Komunikasi (Filtered, Sorted by Default)
      const missingComm: string[] = [];
      const commRisks = sortedByDefault.filter(r => {
        const stats = getScoreAndLevel(r, 'res');
        return stats.score > 0 && stats.level >= 3;
      });

      if (commRisks.length === 0 && baseRisks.length > 0) {
        // Empty
      } else {
        commRisks.forEach((r, idx) => {
          const m: string[] = [];
          if (!r.commMedia?.trim()) m.push('Media');
          if (!r.commProvider?.trim()) m.push('Penyedia');
          if (!r.commReceiver?.trim()) m.push('Penerima');
          if (!r.commPlanTime?.trim()) m.push('Rencana');
          if (m.length > 0) missingComm.push(`${getRiskName(r, idx)}: ${m.join(', ')}`);
        });
      }
      const hasComm = (baseRisks.length > 0) && hasAnalysis && hasResidual && (commRisks.length === 0 || (commRisks.length > 0 && missingComm.length === 0));
      if (!hasComm) reasons.comm = (baseRisks.length > 0 && hasAnalysis && hasResidual && commRisks.length === 0) ? "Lengkap: Tidak ada risiko yang memerlukan komunikasi" : "Kekurangan Menu VI: " + missingComm.slice(0, 2).join('; ');

      // Menu VII: Monitoring PI (Filtered, Sorted by Default)
      const missingPi: string[] = [];
      const piRisks = sortedByDefault.filter(r => {
        const stats = getScoreAndLevel(r, 'res');
        return stats.score > 0 && stats.level >= 3;
      });

      if (piRisks.length === 0 && baseRisks.length > 0) {
        // Empty
      } else {
        piRisks.forEach((r, idx) => {
          const m: string[] = [];
          if (!r.monMethod?.trim()) m.push('Metode');
          if (!r.monPJ?.trim()) m.push('PJ');
          if (!r.monPlanTime?.trim()) m.push('Rencana');
          if (m.length > 0) missingPi.push(`${getRiskName(r, idx)}: ${m.join(', ')}`);
        });
      }
      const hasPiPlan = (baseRisks.length > 0) && hasAnalysis && hasResidual && (piRisks.length === 0 || (piRisks.length > 0 && missingPi.length === 0));
      if (!hasPiPlan) reasons.pi = (baseRisks.length > 0 && hasAnalysis && hasResidual && piRisks.length === 0) ? "Lengkap: Tidak ada risiko yang memerlukan Monitoring PI" : "Kekurangan Menu VII: " + missingPi.slice(0, 2).join('; ');

      // Menu VIII: Heatmap (Automatic if Analysis & Residual done)
      const hasHeatmap = (baseRisks.length > 0) && hasAnalysis && hasResidual;

      // Menu IX: Keterjadian (All risks, default sort)
      const missingOcc: string[] = [];
      if (sortedByDefault.length === 0) {
        missingOcc.push('Daftar Risiko Kosong');
      } else {
        sortedByDefault.forEach((r, idx) => {
          const m: string[] = [];
          if (!r.eventDate?.trim()) m.push('Tgl');
          if (!r.eventImpact?.trim()) m.push('Dampak');
          if (!r.eventCause?.trim()) m.push('Sebab');
          if (!r.rtpRealDate?.trim()) m.push('Realisasi RTP');
          if (!r.rtpNotesContent?.trim()) m.push('Ket');
          if (m.length > 0) missingOcc.push(`${getRiskName(r, idx)}: ${m.join(', ')}`);
        });
      }
      const hasOccurrence = (sortedByDefault.length > 0) && missingOcc.length === 0;
      if (!hasOccurrence) reasons.occurrence = "Kekurangan Menu IX: " + missingOcc.slice(0, 2).join('; ');

      const hasFinalDoc = finalDocsState.some(d => (d.id === acc.uid || d.uid === acc.uid) && (d.status === 'verified' || d.status === 'Verified'));

      const allSteps = [
        hasContext, 
        hasIdentification, 
        hasAnalysis, 
        hasResidual, 
        hasTreatment, 
        hasComm, 
        hasPiPlan, 
        hasHeatmap,
        hasOccurrence,
        hasFinalDoc
      ];
      const completedCount = allSteps.filter(Boolean).length;
      let finalPercent = Math.round((completedCount / 10) * 100);
      
      // Force 100% if verified
      if (hasFinalDoc) {
        finalPercent = 100;
      }

      allStats[acc.uid] = {
        total: baseRisks.length,
        percent: finalPercent,
        reasons: reasons,
        checks: {
          context: hasFinalDoc || hasContext,
          identification: hasFinalDoc || hasIdentification,
          analysis: hasFinalDoc || hasAnalysis,
          residual: hasFinalDoc || hasResidual,
          treatment: hasFinalDoc || hasTreatment,
          comm: hasFinalDoc || hasComm,
          pi: hasFinalDoc || hasPiPlan,
          heatmap: hasFinalDoc || hasHeatmap,
          occurrence: hasFinalDoc || hasOccurrence,
          finalDoc: hasFinalDoc
        }
      };
    });
    setStats(allStats);
  }, [accounts, risksState, contextsState, finalDocsState, filterRiskType, user.role]);

  const sortedAccounts = useMemo(() => {
    let filtered = [...accounts].filter(acc => (acc.role || '').toLowerCase() !== 'administrator');
    
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(acc => (acc.username || '').toLowerCase().includes(q));
    }

    if (sortOrder === 'name') {
      return filtered.sort((a, b) => (a.username || '').localeCompare(b.username || ''));
    }
    return filtered.sort((a, b) => {
      const percentA = stats[a.uid]?.percent || 0;
      const percentB = stats[b.uid]?.percent || 0;
      return sortOrder === 'desc' ? percentB - percentA : percentA - percentB;
    });
  }, [accounts, stats, sortOrder, searchQuery]);

  const handleDownloadReport = async () => {
    setIsExporting(true);
    try {
      const { jsPDF } = await import('jspdf');
      const autoTable = (await import('jspdf-autotable')).default;
      
      const doc = new jsPDF('p', 'mm', 'a4');
      const timestamp = new Date().toLocaleString('id-ID');
      
      doc.setFontSize(16);
      doc.text('LAPORAN MONITORING PROGRESS PENGISIAN', 105, 15, { align: 'center' });
      doc.setFontSize(10);
      doc.text(`Kategori: ${filterRiskType.toUpperCase()} | Waktu: ${timestamp}`, 105, 22, { align: 'center' });

      const tableData = sortedAccounts
        .filter(acc => {
          const role = (acc.role || '').toLowerCase();
          // Exclude admin and operator, only show user
          return role === 'user';
        })
        .map((acc, idx) => {
          const s = stats[acc.uid] || { percent: 0 };
          return [
            idx + 1,
            acc.username || '-',
            `${s.percent || 0}%`
          ];
        });

      autoTable(doc, {
        startY: 30,
        head: [['No', 'Nama OPD', 'Progress %']],
        body: tableData,
        theme: 'grid',
        headStyles: { fillColor: [51, 65, 85], textColor: 255, fontStyle: 'bold', halign: 'center' },
        columnStyles: {
          0: { cellWidth: 15, halign: 'center' },
          1: { cellWidth: 'auto' },
          2: { cellWidth: 35, halign: 'center' }
        },
        styles: { fontSize: 9 }
      });

      doc.save(`Progress_Monitoring_${filterRiskType}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err) {
      console.error('Export error:', err);
      alert('Gagal mengunduh laporan');
    } finally {
      setIsExporting(false);
    }
  };

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex flex-wrap gap-4 justify-between items-center">
          <div className="flex items-center gap-4">
            <h4 className="font-bold text-slate-700 uppercase italic text-sm">Monitoring Progress Pengisian</h4>
            <div className="flex bg-slate-200 p-0.5 rounded-lg text-[9px] font-black">
              <button 
                onClick={() => setFilterRiskType('strategis')}
                className={`px-3 py-1 rounded-md transition-all ${filterRiskType === 'strategis' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
              >
                STRATEGIS
              </button>
              <button 
                onClick={() => setFilterRiskType('operasional')}
                className={`px-3 py-1 rounded-md transition-all ${filterRiskType === 'operasional' ? 'bg-white text-emerald-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
              >
                OPERASIONAL
              </button>
            </div>
            
            <div className="relative ml-2">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={12} />
              <input 
                type="text"
                placeholder="Cari OPD..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="pl-8 pr-4 py-1.5 bg-white border border-slate-200 rounded-lg text-[10px] font-bold outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500 w-48 transition-all"
              />
              {searchQuery && (
                <button 
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X size={10} />
                </button>
              )}
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            <div className="flex bg-slate-100 p-1 rounded-lg gap-1 border border-slate-200">
               <button 
                 onClick={handleBulkExcel}
                 disabled={isExporting}
                 className="flex items-center gap-2 px-3 py-1 text-[9px] font-black uppercase tracking-wider text-emerald-700 hover:bg-emerald-50 rounded transition-all disabled:opacity-50"
                 title="Download Seluruh Data OPD (Excel)"
               >
                 <Download size={14} /> MASSAL EXCEL
               </button>
               <button 
                 onClick={handleBulkPDF}
                 disabled={isExporting}
                 className="flex items-center gap-2 px-3 py-1 text-[9px] font-black uppercase tracking-wider text-red-700 hover:bg-red-50 rounded transition-all disabled:opacity-50"
                 title="Download Seluruh Data OPD (PDF)"
               >
                 <Download size={14} /> MASSAL PDF
               </button>
            </div>
            
            <div className="flex items-center gap-2 bg-white border border-slate-200 px-3 py-1 rounded-lg">
              <span className="text-[10px] font-bold text-slate-400">SORT:</span>
              <select 
                className="text-[10px] font-bold text-slate-600 bg-transparent outline-none cursor-pointer"
                value={sortOrder}
                onChange={(e) => setSortOrder(e.target.value as any)}
              >
                <option value="name">Nama (A-Z)</option>
                <option value="desc">Progress (Tinggi-Rendah)</option>
                <option value="asc">Progress (Rendah-Tinggi)</option>
              </select>
            </div>

            <button 
              onClick={handleDownloadReport}
              disabled={isExporting}
              className="flex items-center gap-2 px-4 py-1.5 bg-slate-900 shadow-lg shadow-slate-200 text-white rounded-lg text-[10px] font-bold uppercase transition-all hover:bg-slate-800 disabled:opacity-50"
            >
              <Download size={14} />
              {isExporting ? 'Exporting...' : 'Unduh Laporan'}
            </button>
          </div>
        </div>
        <div className="p-6">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {sortedAccounts.map((acc) => {
              const s = stats[acc.uid] || { total: 0, percent: 0, checks: {}, reasons: {} };
              return (
                <div 
                  key={acc.uid} 
                  className="p-5 border border-slate-200 rounded-2xl hover:shadow-xl hover:border-blue-200 transition-all group bg-white flex flex-col h-full relative"
                >
                  {/* Tooltip Overlay */}
                  {activeTooltip && activeTooltip.accId === acc.uid && (
                    <div className="absolute inset-x-0 -top-12 z-50 flex justify-center px-4">
                      <div className="bg-slate-900 text-white text-[10px] py-2 px-3 rounded-lg shadow-xl animate-in fade-in zoom-in-95 duration-200 flex items-center gap-2 max-w-[250px]">
                        <div className="w-1.5 h-1.5 rounded-full bg-amber-400 shrink-0" />
                        <span className="leading-tight">{activeTooltip.msg}</span>
                      </div>
                    </div>
                  )}

                  <div className="flex items-start justify-between mb-4">
                      <div className="flex items-center gap-3">
                      <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center text-slate-400 group-hover:bg-blue-600 group-hover:text-white transition-colors capitalize font-black text-xl">
                        {acc.username?.charAt(0)}
                      </div>
                      <div>
                        <div className="text-[8px] font-black text-slate-400 uppercase tracking-widest mb-0.5">Nama OPD</div>
                        <h5 className="font-bold text-slate-800 text-base uppercase tracking-tighter">{acc.username}</h5>
                        <span className={`text-[9px] px-2 py-0.5 rounded-full font-black uppercase tracking-tighter ${
                          acc.role === 'Administrator' ? 'bg-red-100 text-red-600' :
                          acc.role === 'Operator' ? 'bg-purple-100 text-purple-600' :
                          'bg-blue-100 text-blue-600'
                        }`}>
                          {acc.role}
                        </span>
                      </div>
                    </div>
                    <div className="text-right">
                      <div className={`text-2xl font-black leading-none ${s.percent === 100 ? 'text-emerald-600' : 'text-blue-600'}`}>
                        {s.percent}%
                      </div>
                      <div className="text-[8px] font-bold text-slate-400 uppercase">Progres</div>
                    </div>
                  </div>
                  
                  <div className="space-y-4 flex-1">
                     <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                        <motion.div 
                          initial={{ width: 0 }}
                          animate={{ width: `${s.percent}%` }}
                          className={`h-full transition-all duration-1000 ${s.percent === 100 ? 'bg-emerald-500' : 'bg-blue-500'}`}
                        ></motion.div>
                     </div>
  
                     <div className="grid grid-cols-2 gap-2">
                        <div 
                          onMouseEnter={() => !s.checks.context && setActiveTooltip({ accId: acc.uid, menuKey: 'context', msg: s.reasons.context })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.context ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu I</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.context ? 'text-emerald-700' : 'text-slate-400'}`}>Konteks</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.identification && setActiveTooltip({ accId: acc.uid, menuKey: 'ident', msg: s.reasons.identification })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.identification ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu II</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.identification ? 'text-emerald-700' : 'text-slate-400'}`}>Identifikasi</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.analysis && setActiveTooltip({ accId: acc.uid, menuKey: 'analysis', msg: s.reasons.analysis })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.analysis ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu III</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.analysis ? 'text-emerald-700' : 'text-slate-400'}`}>Analisis</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.residual && setActiveTooltip({ accId: acc.uid, menuKey: 'residual', msg: s.reasons.residual })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.residual ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu IV</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.residual ? 'text-emerald-700' : 'text-slate-400'}`}>Risiko Aktual</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.treatment && setActiveTooltip({ accId: acc.uid, menuKey: 'treatment', msg: s.reasons.treatment })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.treatment ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu V</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.treatment ? 'text-emerald-700' : 'text-slate-400'}`}>RTP Penanganan</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.comm && setActiveTooltip({ accId: acc.uid, menuKey: 'comm', msg: s.reasons.comm })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.comm ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu VI</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.comm ? 'text-emerald-700' : 'text-slate-400'}`}>Komunikasi</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.pi && setActiveTooltip({ accId: acc.uid, menuKey: 'pi', msg: s.reasons.pi })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.pi ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu VII</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.pi ? 'text-emerald-700' : 'text-slate-400'}`}>Monitoring PI</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.heatmap && setActiveTooltip({ accId: acc.uid, menuKey: 'heatmap', msg: "Selesaikan Analisis & Residual untuk melihat Peta Risiko" })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.heatmap ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu VIII</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.heatmap ? 'text-emerald-700' : 'text-slate-400'}`}>Peta Risiko</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.occurrence && setActiveTooltip({ accId: acc.uid, menuKey: 'occ', msg: s.reasons.occurrence })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.occurrence ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu IX</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.occurrence ? 'text-emerald-700' : 'text-slate-400'}`}>Monitoring</span>
                           </div>
                        </div>
                        <div 
                          onMouseEnter={() => !s.checks.finalDoc && setActiveTooltip({ accId: acc.uid, menuKey: 'final', msg: "Dokumen Final belum diunggah atau belum diverifikasi Admin" })}
                          onMouseLeave={() => setActiveTooltip(null)}
                          className={`p-2 rounded-lg border flex flex-col gap-1 transition-all ${s.checks.finalDoc ? 'bg-emerald-50 border-emerald-100' : 'bg-slate-50 border-slate-100 opacity-50 hover:bg-slate-100'}`}
                        >
                           <span className="text-[8px] font-black uppercase tracking-tight text-slate-400 font-mono">Menu X</span>
                           <div className="flex items-center gap-1">
                              <span className={`text-[9px] font-bold ${s.checks.finalDoc ? 'text-emerald-700' : 'text-slate-400'}`}>Dokumen Final</span>
                           </div>
                        </div>
                     </div>

                     <button 
                       onClick={() => onSelectUser(acc, filterRiskType)}
                       className="w-full mt-6 py-2.5 rounded-xl bg-slate-50 border border-slate-200 text-slate-600 text-xs font-black uppercase tracking-widest hover:bg-blue-600 hover:text-white hover:border-blue-600 transition-all flex items-center justify-center gap-2"
                     >
                       {user.role === 'Administrator' ? (
                         <>
                           <Edit2 size={14} />
                           Pilih & Edit Data ({filterRiskType === 'operasional' ? 'Opr.' : 'Str.'})
                         </>
                       ) : (
                         <>
                           <Eye size={14} />
                           Pilih & Lihat Data ({filterRiskType === 'operasional' ? 'Opr.' : 'Str.'})
                         </>
                       )}
                     </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

function LoginPage({ onLogin, accounts }: { onLogin: (userData: any) => void, accounts: any[] }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const handleManualLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setIsLoggingIn(true);
    
    try {
      const cleanUsername = username.trim();
      
      // First try to find by username only
      let found: any = null;
      
      // 1. Try exact match (including casing)
      const q = query(
        collection(db, 'accounts'), 
        where('username', '==', cleanUsername)
      );
      const querySnapshot = await getDocs(q);
      
      if (!querySnapshot.empty) {
        const docData = querySnapshot.docs[0];
        found = { ...docData.data(), uid: docData.id };
      } else {
        // 2. Fallback: Check all accounts case-insensitively
        const allSnap = await getDocs(collection(db, 'accounts'));
        const docMatch = allSnap.docs.find(d => (d.data().username || '').toLowerCase() === cleanUsername.toLowerCase());
        if (docMatch) {
          found = { ...docMatch.data(), uid: docMatch.id };
        }
      }
      
      if (found) {
        if (found.password === password) {
          onLogin(found);
        } else {
          setError('Password yang Anda masukkan salah!');
        }
      } else {
        setError('Username tidak terdaftar!');
      }
    } catch (err: any) {
      console.error(err);
      setError('Gagal masuk: ' + (err.message || 'Koneksi database bermasalah'));
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleGoogleLogin = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (err: any) {
      setError('Gagal masuk dengan Google: ' + err.message);
    }
  };

  return (
    <div className="min-h-screen bg-slate-900 flex flex-col items-center justify-center p-6 bg-[radial-gradient(circle_at_top_right,_var(--tw-gradient-stops))] from-blue-900/20 via-slate-900 to-slate-900">
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-md bg-white rounded-2xl shadow-2xl overflow-hidden"
      >
        <div className="p-8 bg-slate-50 border-b border-slate-100 text-center">
          <h1 className="text-3xl font-black text-blue-600 tracking-tighter uppercase italic">ISMAN</h1>
          <p className="text-slate-600 text-[10px] font-black uppercase tracking-tight -mt-1">Integrated Risk Management System</p>
          <p className="text-slate-400 text-[10px] mt-2 font-bold uppercase tracking-widest">Pemerintah Provinsi Papua Tengah</p>
        </div>
        
        <form onSubmit={handleManualLogin} className="p-8 space-y-6">
          {error && (
            <motion.p initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="text-xs text-red-500 font-bold bg-red-50 p-3 rounded-lg border border-red-100 flex items-center gap-2">
              <ShieldAlert size={14} /> {error}
            </motion.p>
          )}

          <div className="space-y-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Username</label>
            <input 
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="w-full bg-slate-50 border border-slate-200 rounded-xl p-4 text-sm outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all font-bold"
              placeholder="Masukkan username"
            />
          </div>

          <div className="space-y-2">
            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Password</label>
            <div className="relative">
              <input 
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-xl p-4 pr-12 text-sm outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all font-bold"
                placeholder="Masukkan password"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors"
              >
                {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </div>

          <button 
            type="submit"
            disabled={isLoggingIn}
            className="w-full bg-blue-600 text-white font-black uppercase tracking-widest py-4 rounded-xl shadow-lg shadow-blue-600/30 hover:bg-blue-700 hover:-translate-y-0.5 active:translate-y-0 disabled:opacity-50 disabled:translate-y-0 transition-all text-sm flex items-center justify-center gap-2"
          >
            {isLoggingIn ? (
              <>
                <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                Memproses...
              </>
            ) : "Masuk dengan Akun Sistem"}
          </button>

          <div className="relative">
            <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-slate-200"></div></div>
            <div className="relative flex justify-center text-xs uppercase"><span className="bg-white px-2 text-slate-400 font-bold">Atau</span></div>
          </div>

          <div className="text-center space-y-4">
            <button 
              type="button"
              onClick={handleGoogleLogin}
              className="w-full flex items-center justify-center gap-3 bg-white border border-slate-200 text-slate-700 font-bold py-4 rounded-xl shadow-sm hover:bg-slate-50 hover:border-slate-300 transition-all text-sm"
            >
              <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-5 h-5" />
              Masuk dengan Google
            </button>
            <p className="text-slate-400 text-[10px] uppercase font-black tracking-widest leading-relaxed px-4">Gunakan akun Google untuk sinkronisasi otomatis antar perangkat</p>
          </div>
        </form>
        
        <div className="p-6 bg-slate-50 border-t border-slate-100 text-center">
          <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest leading-relaxed">
            Pemerintah Provinsi Papua Tengah<br/>Badan Pengawasan Daerah
          </p>
        </div>
      </motion.div>
    </div>
  );
}

// --- Sub-views ---

function AccountManagementView({ accounts }: { accounts: any[] }) {
  const [newAcc, setNewAcc] = useState({ username: '', password: '', role: 'Administrator' });
  const [editingUid, setEditingUid] = useState<string | null>(null);
  const [editForm, setEditForm] = useState({ username: '', password: '', role: '' });
  const [confirmDelete, setConfirmDelete] = useState<string | null>(null);

  const handleAdd = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    
    if (!newAcc.username || !newAcc.password) {
      setError('Username dan password wajib diisi');
      return;
    }

    const cleanUsername = newAcc.username.trim();
    
    if (!cleanUsername || !newAcc.password) {
      setError('Username dan password wajib diisi');
      return;
    }

    // Check if username already exists case-insensitively using current accounts prop
    const isDuplicate = accounts.some(acc => acc.username.toLowerCase() === cleanUsername.toLowerCase());
    if (isDuplicate) {
      setError('Username sudah digunakan! Gunakan nama lain.');
      return;
    }
    
    const newUid = doc(collection(db, 'accounts')).id;
    try {
      const payload = {
        username: cleanUsername,
        password: newAcc.password,
        role: newAcc.role,
        createdAt: new Date().toISOString(),
        uid: newUid
      };
      
      await setDoc(doc(db, 'accounts', newUid), payload);
      console.log('Success creating account:', newUid);
      
      setNewAcc({ username: '', password: '', role: 'Administrator' });
      alert('Akun permanen berhasil dibuat!');
    } catch (error: any) {
      console.error('Create account error:', error);
      alert('Gagal membuat akun: ' + (error.message || 'Error tidak diketahui'));
    }
  };

  const [error, setError] = useState('');

  const [confirmDeleteData, setConfirmDeleteData] = useState<{uid: string, username: string} | null>(null);

  const handleDelete = async (uid: string, username: string) => {
    if (username === 'admin') {
      alert('Akun admin utama tidak dapat dihapus!');
      return;
    }
    setConfirmDeleteData({ uid, username });
  };

  const commitDeleteAccount = async () => {
    if (!confirmDeleteData) return;
    const { uid } = confirmDeleteData;
    try {
      await deleteDoc(doc(db, 'accounts', uid));
      setConfirmDeleteData(null);
      alert('Akun berhasil dihapus permanen.');
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `accounts/${uid}`);
    }
  };

  const startEditing = (acc: any) => {
    setEditingUid(acc.uid);
    setEditForm({ username: acc.username, password: acc.password, role: acc.role });
  };

  const saveEdit = async () => {
    if (!editingUid) return;
    try {
      const cleanUsername = editForm.username.trim();

      // Duplication check for other accounts
      const isDuplicate = accounts.some(acc => 
        acc.uid !== editingUid && 
        acc.username.toLowerCase() === cleanUsername.toLowerCase()
      );

      if (isDuplicate) {
        alert('Username sudah digunakan oleh akun lain!');
        return;
      }

      await setDoc(doc(db, 'accounts', editingUid), {
        ...editForm,
        username: cleanUsername
      }, { merge: true });
      setEditingUid(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `accounts/${editingUid}`);
    }
  };

  const capitalizeAllUsernames = async () => {
    if (!window.confirm("Apakah anda yakin ingin mengubah seluruh username menjadi HURUF KAPITAL?")) return;
    try {
      const promises = accounts.map(acc => {
        return updateDoc(doc(db, 'accounts', acc.uid), {
          username: acc.username.toUpperCase(),
          updatedAt: new Date().toISOString()
        });
      });
      await Promise.all(promises);
      alert('Berhasil mengkapitalisasi seluruh username.');
    } catch (error) {
      console.error('Capitalize error:', error);
      alert('Gagal: ' + error);
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 bg-slate-900 text-white flex justify-between items-center">
          <h4 className="font-bold text-xs uppercase tracking-widest italic">Tambah Akun Baru</h4>
          <button 
            type="button"
            onClick={capitalizeAllUsernames}
            className="text-[9px] font-black uppercase bg-white/10 text-white px-3 py-1 rounded border border-white/20 hover:bg-white/20 transition-all"
          >
            Kapitalisasi Semua Username
          </button>
        </div>
        <form onSubmit={handleAdd} className="p-6 grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
          {error && <div className="col-span-full text-red-500 text-[10px] font-bold bg-red-50 p-2 rounded">{error}</div>}
          <div className="space-y-1">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Username</label>
            <input 
              className="w-full bg-slate-50 border border-slate-200 p-2 rounded outline-none focus:border-blue-500 text-xs" 
              value={newAcc.username}
              onChange={e => setNewAcc({...newAcc, username: e.target.value})}
              placeholder="Username"
            />
          </div>
          <div className="space-y-1">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Password</label>
            <input 
              type="text"
              className="w-full bg-slate-50 border border-slate-200 p-2 rounded outline-none focus:border-blue-500 text-xs" 
              value={newAcc.password}
              onChange={e => setNewAcc({...newAcc, password: e.target.value})}
              placeholder="Password"
            />
          </div>
          <div className="space-y-1">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Role</label>
            <select 
              className="w-full bg-slate-50 border border-slate-200 p-2 rounded outline-none focus:border-blue-500 text-xs"
              value={newAcc.role}
              onChange={e => setNewAcc({...newAcc, role: e.target.value})}
            >
              <option>Administrator</option>
              <option>Operator</option>
              <option>User</option>
            </select>
          </div>
          <button 
            type="submit"
            className="bg-blue-600 text-white py-2 px-4 rounded font-bold uppercase text-[10px] tracking-widest hover:bg-blue-700 transition-all flex items-center justify-center gap-2"
          >
            <Plus size={14} /> Buat Akun
          </button>
        </form>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 bg-slate-50 border-b border-slate-100 italic">
          <h4 className="font-bold text-slate-600 text-xs uppercase tracking-widest flex items-center gap-2">
            <Users size={14} /> Daftar Seluruh Akun
          </h4>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-xs text-left">
            <thead className="bg-slate-50 uppercase text-[10px] font-black text-slate-400 tracking-wider">
              <tr>
                <th className="px-6 py-4 w-12">No</th>
                <th className="px-6 py-4">Username</th>
                <th className="px-6 py-4">Password</th>
                <th className="px-6 py-4">Role</th>
                <th className="px-6 py-4">Terdaftar Pada</th>
                <th className="px-6 py-4 text-center">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {accounts.map((acc, idx) => (
                <tr key={acc.uid || acc.username} className={`hover:bg-slate-50/50 transition-colors ${editingUid === acc.uid ? 'bg-blue-50/50' : ''}`}>
                  <td className="px-6 py-4 font-bold text-slate-400">{idx + 1}</td>
                  <td className="px-6 py-4">
                    {editingUid === acc.uid ? (
                      <input 
                        className="bg-white border border-slate-200 p-1 rounded w-full outline-none focus:border-blue-500" 
                        value={editForm.username}
                        onChange={e => setEditForm({...editForm, username: e.target.value})}
                        disabled={acc.username === 'admin'}
                      />
                    ) : (
                      <div className="flex items-center gap-2">
                        <div className="w-6 h-6 rounded-full bg-blue-100 flex items-center justify-center text-blue-600 font-bold text-[10px]">
                          {acc.username.charAt(0).toUpperCase()}
                        </div>
                        <span className="font-black text-slate-900 uppercase tracking-tighter">{acc.username}</span>
                      </div>
                    )}
                  </td>
                  <td className="px-6 py-4">
                    {editingUid === acc.uid ? (
                      <input 
                        className="bg-white border border-slate-200 p-1 rounded w-full outline-none focus:border-blue-500 font-mono" 
                        value={editForm.password}
                        onChange={e => setEditForm({...editForm, password: e.target.value})}
                      />
                    ) : (
                      <span className="font-mono text-slate-500">{acc.password}</span>
                    )}
                  </td>
                  <td className="px-6 py-4">
                    {editingUid === acc.uid ? (
                      <select 
                        className="bg-white border border-slate-200 p-1 rounded w-full outline-none focus:border-blue-500"
                        value={editForm.role}
                        onChange={e => setEditForm({...editForm, role: e.target.value})}
                        disabled={acc.username === 'admin'}
                      >
                        <option>Administrator</option>
                        <option>Operator</option>
                        <option>User</option>
                      </select>
                    ) : (
                      <span className="bg-slate-100 px-2 py-1 rounded text-[10px] font-bold text-slate-600 uppercase italic">
                        {acc.role}
                      </span>
                    )}
                  </td>
                  <td className="px-6 py-4 text-slate-400">{new Date(acc.createdAt).toLocaleDateString('id-ID')}</td>
                  <td className="px-6 py-4 text-center">
                    <div className="flex items-center justify-center gap-3">
                      {editingUid === acc.uid ? (
                        <>
                          <button onClick={saveEdit} className="text-green-600 hover:text-green-700" title="Simpan">
                            <Check size={18} />
                          </button>
                          <button onClick={() => setEditingUid(null)} className="text-slate-400 hover:text-slate-600" title="Batal">
                            <X size={18} />
                          </button>
                        </>
                      ) : (
                        <>
                          <button onClick={() => startEditing(acc)} className="text-blue-500 hover:text-blue-700" title="Edit">
                            <Edit2 size={16} />
                          </button>
                          <button 
                            onClick={() => handleDelete(acc.uid, acc.username)}
                            className={`transition-all flex items-center justify-center gap-1 font-black uppercase text-[9px] px-3 py-1.5 rounded-full shadow-sm ${
                              acc.username === 'admin' 
                                ? 'opacity-20 cursor-not-allowed text-slate-400 bg-slate-100' 
                                : 'text-red-500 bg-red-50 hover:bg-red-100 border border-red-100'
                            }`}
                            disabled={acc.username === 'admin'}
                          >
                            <Trash2 size={12} /> Hapus Akun
                          </button>
                        </>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {confirmDeleteData && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center border border-slate-100"
            >
              <div className="w-16 h-16 bg-red-50 rounded-full flex items-center justify-center mx-auto mb-4">
                <Trash2 size={32} className="text-red-500" />
              </div>
              <h4 className="text-xl font-bold text-slate-900 mb-2 uppercase italic tracking-tight font-black">Hapus Akun?</h4>
              <p className="text-slate-500 text-[11px] mb-8 font-medium italic leading-relaxed">
                Apakah anda yakin ingin menghapus akun <strong>{confirmDeleteData.username}</strong>? Data pengisian user ini akan tetap ada di database tapi akses masuk akan dicabut.
              </p>
              <div className="flex gap-3">
                <button 
                  onClick={() => setConfirmDeleteData(null)}
                  className="flex-1 px-4 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold text-[10px] uppercase tracking-widest hover:bg-slate-200 transition-colors"
                >
                  Batal
                </button>
                <button 
                  onClick={commitDeleteAccount}
                  className="flex-1 px-4 py-3 bg-red-600 text-white rounded-xl font-black text-[10px] uppercase tracking-widest hover:bg-red-700 transition-colors shadow-lg shadow-red-100"
                >
                  Ya, Hapus
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function PlaceholderView({ title }: { title: string }) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-24 flex flex-col items-center justify-center text-center space-y-4 shadow-sm">
      <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center text-slate-400">
        <FileText size={32} />
      </div>
      <div>
        <h3 className="text-xl font-bold text-slate-900">{title}</h3>
        <p className="text-slate-500 mt-1 max-w-md">
          Halaman modul <strong>{title}</strong> siap diisi data sesuai petunjuk teknis.
        </p>
      </div>
      <button className="text-blue-600 font-medium hover:underline flex items-center gap-1">
        Buka Petunjuk Pengisian <ChevronRight size={16} />
      </button>
    </div>
  );
}

const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

function ScoreTable({ 
  title, 
  type, 
  color, 
  rows, 
  participantCount, 
  updateScores, 
  onUpdateParticipantCount,
  isReadOnly
}: { 
  title: string, 
  type: 'dampak' | 'kemungkinan', 
  color: string, 
  rows: any[], 
  participantCount: number,
  updateScores: (id: string, type: 'dampak' | 'kemungkinan', scores: number[]) => Promise<void>,
  onUpdateParticipantCount: (newCount: number) => void,
  isReadOnly?: boolean
}) {
  const calculateAvg = (scores?: number[]) => {
    if (!scores || scores.length === 0) return '0.00';
    // Only average the scores for the current participant count that are non-zero (active voters)
    const activeScores = scores.slice(0, participantCount).filter(v => v > 0);
    if (activeScores.length === 0) return '0.00';
    return (activeScores.reduce((a, b) => a + b, 0) / activeScores.length).toFixed(2);
  };

  const currentParticipants = alphabet.slice(0, participantCount);

  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
      <div className={`px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center`}>
        <div className="flex items-center gap-4">
          <h4 className={`font-bold text-slate-700 uppercase border-l-4 ${color} pl-3 italic`}>
            Formulir Perhitungan Rata-rata Skala {title}
          </h4>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1 bg-white border border-slate-200 rounded-lg p-1 shadow-sm">
            <span className="text-[9px] font-black uppercase text-slate-400 px-2">Peserta: {participantCount}</span>
            {!isReadOnly && (
              <>
                <button 
                  onClick={() => onUpdateParticipantCount(Math.max(1, participantCount - 1))}
                  disabled={participantCount <= 1}
                  className="w-6 h-6 flex items-center justify-center rounded bg-slate-100 hover:bg-slate-200 disabled:opacity-30 transition-colors"
                >
                  <Minus size={12} />
                </button>
                <button 
                  onClick={() => onUpdateParticipantCount(Math.min(26, participantCount + 1))}
                  disabled={participantCount >= 26}
                  className="w-6 h-6 flex items-center justify-center rounded bg-slate-100 hover:bg-slate-200 disabled:opacity-30 transition-colors"
                >
                  <Plus size={12} />
                </button>
              </>
            )}
          </div>
          <span className="text-[10px] bg-slate-100 px-2 py-1 rounded font-bold uppercase tracking-tight italic">Skor 1-5</span>
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-[11px] text-left min-w-[800px]">
          <thead className="bg-slate-50 text-slate-500 border-b border-slate-200 uppercase text-[9px] font-black">
            <tr>
              <th className="px-3 py-3 w-10" rowSpan={2}>No</th>
              <th className="px-3 py-3" rowSpan={2}>Kode & Uraian Risiko</th>
              <th className="px-3 py-3 text-center border-l border-slate-200" colSpan={participantCount}>Skala Penilaian Peserta</th>
              <th className="px-3 py-3 text-center bg-slate-100 uppercase" rowSpan={2}>
                {type === 'dampak' ? 'Rata-rata Dampak' : 'Rata-rata Kemungkinan'}
              </th>
            </tr>
            <tr className="border-t border-slate-200">
              {currentParticipants.map(p => (
                <th key={p} className="px-1 py-2 text-center border-l border-slate-200 w-10">{p}</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {rows.map((row, idx) => {
              const baseScores = (type === 'dampak' ? row.dampakScores : row.kemungkinanScores) || [];
              // Ensure we have enough score entries for display
              const scores = [...baseScores];
              while(scores.length < participantCount) scores.push(0);
              
              return (
                <tr key={row.id} className="hover:bg-slate-50/50 transition-colors">
                  <td className="px-3 py-3 text-slate-400 text-center">{idx + 1}</td>
                  <td className="px-3 py-3">
                    <p className="font-bold text-blue-600 mb-0.5">{row.risikoKode || 'RSO.26.XX.YY.ZZ'}</p>
                    <p className="text-slate-600 whitespace-pre-wrap leading-tight mb-2">{row.risikoUraian || '(Uraian belum diisi di Menu 2)'}</p>
                    
                    {(() => {
                      const sRows = row.subRows && row.subRows.length > 0 ? row.subRows : [
                        { sebabUraian: row.sebabUraian || '', dampakUraian: row.dampakUraian || '' }
                      ];
                      
                      const hasContent = sRows.some((s: any) => s.sebabUraian?.trim() || s.dampakUraian?.trim());
                      if (!hasContent) return null;

                      return (
                        <div className="space-y-1 mt-2">
                          {sRows.map((sub: any, sIdx: number) => (
                            <div key={sIdx} className="bg-slate-50 p-1.5 rounded text-[8px] border border-slate-100">
                              <p className="text-slate-500 italic"><span className="font-bold opacity-50">Sebab {sIdx + 1}:</span> {sub.sebabUraian || '-'}</p>
                              <p className="text-red-700/70"><span className="font-bold opacity-50">Dampak {sIdx + 1}:</span> {sub.dampakUraian || '-'}</p>
                            </div>
                          ))}
                        </div>
                      );
                    })()}
                  </td>
                  {currentParticipants.map((_, sIdx) => {
                    const s = scores[sIdx] || 0;
                    return (
                      <td key={sIdx} className="px-1 py-3 text-center border-l border-slate-200">
                        <input 
                          type="number" 
                          min="0" 
                          max="5"
                          className="w-8 text-center bg-white border border-slate-100 rounded outline-none focus:border-blue-300 transition-shadow disabled:opacity-50"
                          value={s === 0 ? '' : s}
                          disabled={isReadOnly}
                          onChange={(e) => {
                            const val = parseInt(e.target.value) || 0;
                            const newScores = [...scores];
                            newScores[sIdx] = Math.min(Math.max(val, 0), 5);
                            updateScores(row.id, type, newScores.slice(0, Math.max(newScores.length, participantCount)));
                          }}
                        />
                      </td>
                    );
                  })}
                  <td className="px-3 py-3 text-center bg-slate-50 font-black text-blue-800 text-sm">
                    {calculateAvg(scores)}
                  </td>
                </tr>
              );
            })}
            {rows.length === 0 && (
              <tr>
                <td colSpan={participantCount + 3} className="px-6 py-8 text-center text-slate-400 italic">Belum ada risiko teridentifikasi di Menu 2.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function RiskAnalysisView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [participantCount, setParticipantCount] = useState(5);

  const contextId = `risk_context_${user.uid}_${riskType}`;

  // Listen to Risk Identification Data
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id }));
      data.sort((a: any, b: any) => {
        if (a.order !== undefined && b.order !== undefined) return a.order - b.order;
        const dateA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
        const dateB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
        return dateA - dateB;
      });
      setRows(data);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid]);

  // Listen to Context for participant count
  useEffect(() => {
    const docRef = doc(db, 'risk_context', contextId);
    const unsubscribe = onSnapshot(docRef, (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        if (data.participantCount !== undefined) {
          setParticipantCount(data.participantCount);
        }
      }
    });
    return () => unsubscribe();
  }, [contextId]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');
  const [pasteType, setPasteType] = useState<'dampak' | 'kemungkinan' | 'total'>('total');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;

      setLoading(true);
      setShowPasteModal(false);

      const batch = writeBatch(db);
      let count = 0;

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 2) continue;

        const [kode, ...scores] = cols;
        const targetRow = rows.find(r => r.risikoKode === kode.trim());
        if (!targetRow) continue;

        const parsedScores = scores.map(s => parseInt(s) || 0);

        if (pasteType === 'total') {
          // Legacy/Total logic: assumes Impact scores followed by Probability scores
          const dScores = parsedScores.slice(0, participantCount);
          const pScores = parsedScores.slice(participantCount, participantCount * 2);
          batch.update(doc(db, 'risk_identification', targetRow.id), {
            dampakScores: dScores,
            kemungkinanScores: pScores,
            updatedAt: new Date().toISOString()
          });
        } else {
          const field = pasteType === 'dampak' ? 'dampakScores' : 'kemungkinanScores';
          batch.update(doc(db, 'risk_identification', targetRow.id), {
            [field]: parsedScores.slice(0, participantCount),
            updatedAt: new Date().toISOString()
          });
        }
        count++;
      }

      if (count > 0) {
        await batch.commit();
      }

      setPasteData('');
      alert(`Berhasil mengimpor ${count} skor ${pasteType === 'total' ? 'analisis' : pasteType}.`);
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const openPasteModal = (type: typeof pasteType) => {
    setPasteType(type);
    setPasteData('');
    setShowPasteModal(true);
  };

  const updateScores = async (id: string, type: 'dampak' | 'kemungkinan', scores: number[]) => {
    if (isReadOnly) return;
    try {
      const field = type === 'dampak' ? 'dampakScores' : 'kemungkinanScores';
      await updateDoc(doc(db, 'risk_identification', id), {
        [field]: scores,
        updatedAt: new Date().toISOString()
      });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, `risk_identification/${id}`);
    }
  };

  const updateParticipantCount = async (newCount: number) => {
    try {
      await setDoc(doc(db, 'risk_context', contextId), {
        participantCount: newCount,
        updatedAt: new Date().toISOString()
      }, { merge: true });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.WRITE, `risk_context/${contextId}`);
    }
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  return (
    <div className="space-y-12">
      <div className="flex justify-between items-center bg-slate-50 p-4 rounded-xl border border-slate-200 shadow-sm mb-6">
        <div>
          <h3 className="font-black text-xs uppercase tracking-widest text-slate-700">Analisis Risiko (DxS)</h3>
          <p className="text-[10px] text-slate-500 font-medium italic">Matriks Pergub No 67 Thn 2023</p>
        </div>
      </div>

      <div className="space-y-8">
        <div className="space-y-2">
          {!isReadOnly && (
            <div className="flex justify-end px-2">
              <button 
                onClick={() => openPasteModal('dampak')}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 text-red-700 border border-red-200 rounded-lg hover:bg-red-100 transition-all text-[9px] font-black uppercase tracking-wider"
              >
                <ClipboardList size={14} /> Paste Skor Dampak (Excel)
              </button>
            </div>
          )}
          <ScoreTable 
            title="Dampak" 
            type="dampak" 
            color="border-red-500" 
            rows={rows} 
            participantCount={participantCount}
            updateScores={updateScores}
            onUpdateParticipantCount={updateParticipantCount}
            isReadOnly={isReadOnly}
          />
        </div>

        <div className="space-y-2">
          {!isReadOnly && (
            <div className="flex justify-end px-2">
              <button 
                onClick={() => openPasteModal('kemungkinan')}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-lg hover:bg-blue-100 transition-all text-[9px] font-black uppercase tracking-wider"
              >
                <ClipboardList size={14} /> Paste Skor Kemungkinan (Excel)
              </button>
            </div>
          )}
          <ScoreTable 
            title="Kemungkinan" 
            type="kemungkinan" 
            color="border-blue-500" 
            rows={rows} 
            participantCount={participantCount}
            updateScores={updateScores}
            onUpdateParticipantCount={updateParticipantCount}
            isReadOnly={isReadOnly}
          />
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm pt-4">
        <div className="px-6 py-4 border-b border-slate-100 bg-emerald-900 text-white flex justify-between items-center">
          <h4 className="font-bold text-[10px] uppercase border-l-2 border-emerald-400 pl-3 italic tracking-wider">Kertas Kerja Hasil Analisis Risiko (DxS)</h4>
          <span className="text-[9px] font-mono italic opacity-60">Matriks Pergub No 67 Thn 2023</span>
        </div>

        {showPasteModal && (
          <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className={`p-4 ${pasteType === 'dampak' ? 'bg-red-600' : pasteType === 'kemungkinan' ? 'bg-blue-600' : 'bg-emerald-600'} flex items-center justify-between`}>
                <div>
                  <h3 className="text-white font-black text-xs uppercase tracking-widest">
                    Paste Skor {pasteType === 'total' ? 'Analisis' : (pasteType === 'dampak' ? 'Dampak' : 'Kemungkinan')} dari Excel
                  </h3>
                  <p className="text-white/80 text-[9px] font-bold mt-0.5">
                    Panduan: Copy kolom Kode dan Nilai Skor di Excel (Kode | P1 | P2 | ...)
                  </p>
                </div>
                <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
              </div>
              <div className="p-6 flex-1">
                <textarea 
                  autoFocus
                  placeholder="Paste di sini..."
                  className={`w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:ring-4 ${pasteType === 'dampak' ? 'focus:border-red-500 focus:ring-red-500/5' : pasteType === 'kemungkinan' ? 'focus:border-blue-500 focus:ring-blue-500/5' : 'focus:border-emerald-500 focus:ring-emerald-500/5'} resize-none font-mono`}
                  value={pasteData}
                  onChange={(e) => setPasteData(e.target.value)}
                ></textarea>
              </div>
              <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
                <button 
                  onClick={processImport} 
                  disabled={!pasteData.trim()} 
                  className={`px-8 py-2 ${pasteType === 'dampak' ? 'bg-red-600 hover:bg-red-700' : pasteType === 'kemungkinan' ? 'bg-blue-600 hover:bg-blue-700' : 'bg-emerald-600 hover:bg-emerald-700'} text-white font-black text-[10px] uppercase rounded-lg shadow-lg tracking-widest`}
                >
                  Impor
                </button>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[1200px]">
            <thead className="bg-slate-900 text-white uppercase text-[9px] font-black tracking-widest">
              <tr>
                <th className="px-4 py-4 w-12 border-b border-slate-800" rowSpan={2}>No</th>
                <th className="px-4 py-4 border-b border-slate-800" rowSpan={2}>Analisis Risiko Teridentifikasi</th>
                <th className="px-4 py-3 text-center border-b border-slate-800 border-l border-slate-700" colSpan={participantCount}>Dampak (Impact)</th>
                <th className="px-4 py-3 text-center border-b border-slate-800 border-l border-slate-700 bg-slate-800/50" rowSpan={2}>Avg D</th>
                <th className="px-4 py-3 text-center border-b border-slate-800 border-l border-slate-700" colSpan={participantCount}>Probabilitas (S)</th>
                <th className="px-4 py-3 text-center border-b border-slate-800 border-l border-slate-700 bg-slate-800/50" rowSpan={2}>Avg S</th>
                <th className="px-4 py-3 text-center border-b border-slate-800 border-l border-slate-700 bg-emerald-800" rowSpan={2}>Score (DxS)</th>
                <th className="px-4 py-3 text-center border-b border-slate-800 border-l border-slate-700 bg-emerald-800" rowSpan={2}>Level</th>
              </tr>
              <tr className="bg-slate-800/50">
                {Array.from({ length: participantCount }).map((_, i) => (
                  <th key={`d-${i}`} className="px-1 py-2 text-center text-[8px] border-l border-slate-700 w-8">P{i+1}</th>
                ))}
                {Array.from({ length: participantCount }).map((_, i) => (
                  <th key={`k-${i}`} className="px-1 py-2 text-center text-[8px] border-l border-slate-700 w-8">P{i+1}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {rows.map((row, idx) => {
                const dScores = (row.dampakScores || []).slice(0, participantCount);
                const kScores = (row.kemungkinanScores || []).slice(0, participantCount);
                
                const activeD = dScores.filter((v: any) => v > 0);
                const activeK = kScores.filter((v: any) => v > 0);
                
                const avgDampak = activeD.length > 0 ? parseFloat((activeD.reduce((a: any, b: any) => a + b, 0) / activeD.length).toFixed(2)) : 0;
                const avgKemungkinan = activeK.length > 0 ? parseFloat((activeK.reduce((a: any, b: any) => a + b, 0) / activeK.length).toFixed(2)) : 0;

                const risk = getRiskLevel(avgDampak, avgKemungkinan);
                const score = avgDampak * avgKemungkinan;

                return (
                  <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                    <td className="px-4 py-6 text-center font-bold text-slate-400 border-r border-slate-100">{idx + 1}</td>
                    <td className="px-4 py-6 border-r border-slate-100 min-w-[200px]">
                      <p className="text-blue-600 font-black mb-1 text-[9px]">{row.risikoKode || (riskType === 'operasional' ? 'ROO...' : 'RSO...')}</p>
                      <p className="text-slate-700 font-medium leading-relaxed uppercase text-[10px] mb-2">{row.risikoUraian || '(Uraian belum diisi)'}</p>
                      
                      {(() => {
                        const sRows = row.subRows && row.subRows.length > 0 ? row.subRows : [
                          { sebabUraian: row.sebabUraian || '', dampakUraian: row.dampakUraian || '' }
                        ];
                        
                        // Only render if there's actual content in at least one field
                        const hasContent = sRows.some((s: any) => s.sebabUraian?.trim() || s.dampakUraian?.trim());
                        if (!hasContent) return null;

                        return (
                          <div className="space-y-1">
                            {sRows.map((sub: any, sIdx: number) => (
                              <div key={sIdx} className="bg-slate-50 p-1.5 rounded text-[8px] border border-slate-100">
                                <p className="text-slate-500 italic"><span className="font-bold opacity-50">Sebab {sIdx + 1}:</span> {sub.sebabUraian || '-'}</p>
                                <p className="text-red-700/70"><span className="font-bold opacity-50">Dampak {sIdx + 1}:</span> {sub.dampakUraian || '-'}</p>
                              </div>
                            ))}
                          </div>
                        );
                      })()}
                    </td>
                    
                    {/* Dampak Details */}
                    {Array.from({ length: participantCount }).map((_, i) => (
                      <td key={`dv-${i}`} className="px-1 py-6 text-center border-l border-slate-100 text-slate-500 font-mono text-[10px]">
                        {dScores[i] || 0}
                      </td>
                    ))}
                    <td className="px-2 py-6 text-center bg-slate-50 font-bold border-l border-slate-200 text-blue-600">{avgDampak.toFixed(2)}</td>

                    {/* Kemungkinan Details */}
                    {Array.from({ length: participantCount }).map((_, i) => (
                      <td key={`kv-${i}`} className="px-1 py-6 text-center border-l border-slate-100 text-slate-500 font-mono text-[10px]">
                        {kScores[i] || 0}
                      </td>
                    ))}
                    <td className="px-2 py-6 text-center bg-slate-50 font-bold border-l border-slate-200 text-blue-600">{avgKemungkinan.toFixed(2)}</td>

                    <td className="px-4 py-6 text-center bg-emerald-50 border-l border-slate-200 font-black text-slate-900 text-sm">
                      {score > 0 ? score.toFixed(2) : '-'}
                    </td>
                    <td className="px-4 py-6 text-center border-l border-slate-200">
                      {score > 0 ? (
                        <span className={`${risk.color} px-3 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest shadow-sm inline-block min-w-[100px]`}>
                          {risk.label}
                        </span>
                      ) : '-'}
                    </td>
                  </tr>
                );
              })}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={6} className="px-6 py-8 text-center text-slate-400 italic">Belum ada risiko teridentifikasi di Menu 2.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// Matrix logic based on the user's provided PNG (Matriks Analisa Risiko Pergub 67 Tahun 2023)
const getRiskLevel = (d: number, k: number) => {
  // Rounds to nearest floor/ceil for matrix mapping (1-5 scale)
  const D = Math.ceil(d);
  const K = Math.ceil(k);
  
  // Matrix definitions
  if (K === 5) {
    if (D <= 2) return { label: 'Tinggi', color: 'bg-amber-500 text-white', level: 3 };
    return { label: 'Sangat Tinggi', color: 'bg-red-600 text-white', level: 4 };
  }
  if (K === 4) {
    if (D === 1) return { label: 'Sedang', color: 'bg-yellow-400 text-slate-900', level: 2 };
    if (D === 2 || D === 3) return { label: 'Tinggi', color: 'bg-amber-500 text-white', level: 3 };
    return { label: 'Sangat Tinggi', color: 'bg-red-600 text-white', level: 4 };
  }
  if (K === 3) {
    if (D <= 2) return { label: 'Sedang', color: 'bg-yellow-400 text-slate-900', level: 2 };
    if (D === 3 || D === 4) return { label: 'Tinggi', color: 'bg-amber-500 text-white', level: 3 };
    return { label: 'Sangat Tinggi', color: 'bg-red-600 text-white', level: 4 };
  }
  if (K === 2) {
    if (D <= 2) return { label: 'Rendah', color: 'bg-green-500 text-white', level: 1 }; // (1,2) and (2,2) are now Green (Rendah)
    if (D === 3) return { label: 'Sedang', color: 'bg-yellow-400 text-slate-900', level: 2 };
    return { label: 'Tinggi', color: 'bg-amber-500 text-white', level: 3 };
  }
  if (K === 1) {
    if (D <= 2) return { label: 'Rendah', color: 'bg-green-500 text-white', level: 1 };
    if (D === 3 || D === 4) return { label: 'Sedang', color: 'bg-yellow-400 text-slate-900', level: 2 };
    return { label: 'Tinggi', color: 'bg-amber-500 text-white', level: 3 };
  }
  return { label: 'N/A', color: 'bg-slate-100', level: 0 };
};


function RiskResidualView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [participantCount, setParticipantCount] = useState(5);
  const contextId = `risk_context_${user.uid}_${riskType}`;

  // Listen to Context for participant count
  useEffect(() => {
    const docRef = doc(db, 'risk_context', contextId);
    const unsubscribe = onSnapshot(docRef, (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        if (data.participantCount !== undefined) {
          setParticipantCount(data.participantCount);
        }
      }
    });
    return () => unsubscribe();
  }, [contextId]);

  // Listen to Risk Identification Data
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const d = doc.data();
        const activeD = (d.dampakScores || []).slice(0, participantCount).filter((v: any) => v > 0);
        const activeK = (d.kemungkinanScores || []).slice(0, participantCount).filter((v: any) => v > 0);
        const avgD = activeD.length > 0 ? (activeD.reduce((a: any, b: any) => a + b, 0) / activeD.length) : 0;
        const avgK = activeK.length > 0 ? (activeK.reduce((a: any, b: any) => a + b, 0) / activeK.length) : 0;
        const initScore = parseFloat(avgD.toFixed(2)) * parseFloat(avgK.toFixed(2));
        const initRisk = getRiskLevel(avgD, avgK);

        return { 
          ...d, 
          id: doc.id,
          avgD,
          avgK,
          initScore,
          initRiskLabel: initRisk.label,
          initRiskLevel: initRisk.level
        };
      });

      // SORT: Primarily by initial level, then by score descending
      data.sort((a, b) => {
        if (b.initRiskLevel !== a.initRiskLevel) return b.initRiskLevel - a.initRiskLevel;
        return b.initScore - a.initScore;
      });

      setRows(data);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid, participantCount, riskType]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;
      setLoading(true);
      setShowPasteModal(false);

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 2) continue;

        const [kode, control, gap, d, k] = cols;
        const targetRow = rows.find(r => r.risikoKode === kode.trim());
        if (!targetRow) continue;

        await updateDoc(doc(db, 'risk_identification', targetRow.id), {
          rtpControl: (control || '').trim(),
          rtpGap: (gap || '').trim(),
          residualDampak: parseFloat(d) || 0,
          residualKemungkinan: parseFloat(k) || 0,
          updatedAt: new Date().toISOString()
        });
      }
      setPasteData('');
      alert('Berhasil mengimpor data penilaian residual.');
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const updateResidualField = async (id: string, field: string, value: any) => {
    if (isReadOnly) return;
    try {
      await updateDoc(doc(db, 'risk_identification', id), {
        [field]: value,
        updatedAt: new Date().toISOString()
      });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, `risk_identification/${id}`);
    }
  };

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
          <h4 className="font-bold text-slate-700 uppercase border-l-4 border-blue-500 pl-3 italic text-xs">Identifikasi Risiko Aktual & Risiko Sisa</h4>
          {!isReadOnly && (
            <button 
              onClick={() => setShowPasteModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-all text-[10px] font-bold uppercase tracking-wider shadow-sm"
            >
              <ClipboardList size={14} /> Paste Penilaian Residual (Excel)
            </button>
          )}
        </div>

        {showPasteModal && (
          <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className="p-4 bg-blue-600 flex items-center justify-between">
                <div>
                  <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste Penilaian Residual dari Excel</h3>
                  <p className="text-blue-100 text-[9px] font-bold mt-0.5">Kolom: Kode | Pengendalian | Celah | D | K</p>
                </div>
                <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
              </div>
              <div className="p-6 flex-1">
                <textarea 
                  autoFocus
                  placeholder="Paste di sini..."
                  className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:border-blue-500 resize-none font-mono"
                  value={pasteData}
                  onChange={(e) => setPasteData(e.target.value)}
                ></textarea>
              </div>
              <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
                <button onClick={processImport} disabled={!pasteData.trim()} className="px-8 py-2 bg-blue-600 text-white font-black text-[10px] uppercase rounded-lg">Impor</button>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-[10px] text-left">
            <thead className="bg-slate-900 text-white uppercase font-bold tracking-wider text-[9px]">
              <tr>
                <th className="px-2 py-4 border-r border-slate-800 w-10" rowSpan={2}>No</th>
                <th className="px-3 py-4 border-r border-slate-800" rowSpan={2}>Risiko Teridentifikasi</th>
                <th className="px-2 py-4 border-r border-slate-800 text-center bg-slate-800" colSpan={4}>Risiko Awal</th>
                <th className="px-3 py-4 border-r border-slate-800" rowSpan={2}>Pengendalian Saat Ini</th>
                <th className="px-3 py-4 border-r border-slate-800" rowSpan={2}>Celah (Gap)</th>
                <th className="px-2 py-4 border-r border-slate-800 text-center bg-slate-800" colSpan={4}>Risiko Sisa (Residual)</th>
              </tr>
              <tr className="bg-slate-800 text-[8px]">
                {/* Initial Risk Headers */}
                <th className="px-2 py-2 text-center border-r border-slate-700 w-10 italic opacity-70">D</th>
                <th className="px-2 py-2 text-center border-r border-slate-700 w-10 italic opacity-70">K</th>
                <th className="px-2 py-2 text-center border-r border-slate-700 w-14">Skor</th>
                <th className="px-2 py-2 text-center border-r border-slate-700 w-20 tracking-tighter">Level</th>
                {/* Residual Risk Headers */}
                <th className="px-2 py-2 text-center border-r border-slate-700 w-12">D</th>
                <th className="px-2 py-2 text-center border-r border-slate-700 w-12">S</th>
                <th className="px-2 py-2 text-center border-r border-slate-700 w-16">Skor</th>
                <th className="px-2 py-2 text-center w-24 tracking-tighter">Level</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {rows.map((row, idx) => {
                const resD = parseFloat(row.residualDampak || 0);
                const resK = parseFloat(row.residualKemungkinan || 0);
                const resScore = resD * resK;
                const resRisk = getRiskLevel(resD, resK);

                return (
                  <tr key={row.id} className="hover:bg-slate-50 transition-colors align-top">
                    <td className="px-2 py-4 font-bold text-slate-400 text-center">{idx + 1}</td>
                    <td className="px-3 py-4 min-w-[200px]">
                      <p className="font-black text-blue-600 text-[9px] mb-1">{row.risikoKode || (riskType === 'operasional' ? 'ROO...' : 'RSO...')}</p>
                      <p className="font-medium text-slate-700 uppercase leading-tight mb-2">{row.risikoUraian || '(Belum diisi)'}</p>
                      
                      {(() => {
                        const sRows = row.subRows && row.subRows.length > 0 ? row.subRows : [
                          { sebabUraian: row.sebabUraian || '', dampakUraian: row.dampakUraian || '' }
                        ];
                        
                        // Only render if there's actual content in at least one field
                        const hasContent = sRows.some((s: any) => s.sebabUraian?.trim() || s.dampakUraian?.trim());
                        if (!hasContent) return null;

                        return (
                          <div className="space-y-1">
                            {sRows.map((sub: any, sIdx: number) => (
                              <div key={sIdx} className="bg-slate-50 p-1.5 rounded text-[8px] border border-slate-100 italic">
                                 <p className="text-slate-500 font-bold opacity-50 uppercase text-[7px] mb-0.5">Sebab {sIdx + 1}</p>
                                 <p className="text-slate-600 mb-1">{sub.sebabUraian || '-'}</p>
                                 <p className="text-red-800 font-bold opacity-30 uppercase text-[7px] mb-0.5">Dampak {sIdx + 1}</p>
                                 <p className="text-red-600">{sub.dampakUraian || '-'}</p>
                              </div>
                            ))}
                          </div>
                        );
                      })()}
                    </td>
                    <td className="px-2 py-4 text-center border-r border-slate-100 bg-slate-50/20 text-slate-500 italic">
                      {row.avgD?.toFixed(2) || '0.00'}
                    </td>
                    <td className="px-2 py-4 text-center border-r border-slate-100 bg-slate-50/20 text-slate-500 italic">
                      {row.avgK?.toFixed(2) || '0.00'}
                    </td>
                    <td className="px-2 py-4 text-center border-r border-slate-100 bg-slate-100/50">
                      <span className="font-black text-slate-900">{row.initScore.toFixed(2)}</span>
                    </td>
                    <td className="px-2 py-4 text-center border-r border-slate-100 bg-slate-100/50">
                      <span className="text-[8px] font-black uppercase text-slate-400">{row.initRiskLabel}</span>
                    </td>
                    <td className="px-3 py-4">
                      <div className="flex flex-col gap-1">
                        <select 
                          className="w-full bg-white border border-slate-200 rounded text-[8px] p-0.5 outline-none"
                          value={row.rtpControl === 'Tidak Ada' ? 'Tidak Ada' : 'Manual'}
                          onChange={async (e) => {
                            if (e.target.value === 'Tidak Ada') {
                              // Auto fill everything as requested
                              try {
                                await updateDoc(doc(db, 'risk_identification', row.id), {
                                  rtpControl: 'Tidak Ada',
                                  rtpGap: '-',
                                  residualDampak: row.avgD || 0,
                                  residualKemungkinan: row.avgK || 0,
                                  updatedAt: new Date().toISOString()
                                });
                              } catch (err) {
                                console.error(err);
                              }
                            } else {
                              updateResidualField(row.id, 'rtpControl', '');
                            }
                          }}
                        >
                          <option value="Manual">Input Manual</option>
                          <option value="Tidak Ada">Tidak Ada</option>
                        </select>
                        <EditableTextarea 
                          className="w-full bg-slate-50 border border-slate-100 rounded p-1 outline-none resize-none text-[9px] focus:border-blue-300 disabled:opacity-50" 
                          rows={3}
                          placeholder="..."
                          value={row.rtpControl || ''}
                          disabled={isReadOnly || row.rtpControl === 'Tidak Ada'}
                          onChange={(val) => updateResidualField(row.id, 'rtpControl', val)}
                        />
                      </div>
                    </td>
                    <td className="px-3 py-4">
                      <EditableTextarea 
                        className={`w-full border rounded p-1 outline-none resize-none text-[9px] focus:border-red-300 transition-all ${row.rtpControl === 'Tidak Ada' ? 'bg-black text-white cursor-not-allowed border-black' : 'bg-slate-50 border-slate-100 text-red-600'}`} 
                        rows={3}
                        placeholder="..."
                        value={row.rtpControl === 'Tidak Ada' ? 'N/A (Tidak Ada Pengendalian)' : (row.rtpGap || '')}
                        disabled={isReadOnly || row.rtpControl === 'Tidak Ada'}
                        onChange={(val) => updateResidualField(row.id, 'rtpGap', val)}
                      />
                    </td>
                    <td className="px-2 py-4 text-center border-l border-slate-100 bg-slate-50/30">
                      <EditableInput 
                        type="number" 
                        className={`w-10 text-center border border-slate-200 rounded outline-none font-bold text-blue-600 ${row.rtpControl === 'Tidak Ada' ? 'bg-slate-100 cursor-not-allowed' : 'bg-white'}`}
                        value={row.rtpControl === 'Tidak Ada' ? (row.avgD || 0).toFixed(2) : (row.residualDampak || '')}
                        disabled={isReadOnly || row.rtpControl === 'Tidak Ada'}
                        onChange={(val) => updateResidualField(row.id, 'residualDampak', val)}
                      />
                    </td>
                    <td className="px-2 py-4 text-center border-l border-slate-100 bg-slate-50/30">
                      <EditableInput 
                        type="number"
                        className={`w-10 text-center border border-slate-200 rounded outline-none font-bold text-blue-600 ${row.rtpControl === 'Tidak Ada' ? 'bg-slate-100 cursor-not-allowed' : 'bg-white'}`}
                        value={row.rtpControl === 'Tidak Ada' ? (row.avgK || 0).toFixed(2) : (row.residualKemungkinan || '')}
                        disabled={isReadOnly || row.rtpControl === 'Tidak Ada'}
                        onChange={(val) => updateResidualField(row.id, 'residualKemungkinan', val)}
                      />
                    </td>
                    <td className="px-2 py-4 text-center border-l border-slate-100 bg-slate-50/50 font-black text-slate-900">
                      {resScore > 0 ? resScore.toFixed(2) : '-'}
                    </td>
                    <td className="px-2 py-4 text-center border-l border-slate-100 bg-slate-50/50">
                      {resScore > 0 ? (
                        <span className={`${resRisk.color} px-2 py-1 rounded text-[8px] font-black uppercase tracking-tighter shadow-sm`}>
                          {resRisk.label}
                        </span>
                      ) : '-'}
                    </td>
                  </tr>
                );
              })}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={9} className="px-6 py-12 text-center text-slate-400 italic">Silahkan isi Analisis Risiko di Menu 3 terlebih dahulu.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function RiskTreatmentView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Listen to Risk Identification Data
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const d = doc.data();
        
        // Calculate Residual Risk for sorting
        const resD = parseFloat(d.residualDampak || 0);
        const resK = parseFloat(d.residualKemungkinan || 0);
        const resScore = resD * resK;
        const resRisk = getRiskLevel(resD, resK);

        return { 
          ...d, 
          id: doc.id,
          resScore,
          resRiskLabel: resRisk.label,
          resRiskLevel: resRisk.level
        };
      });

      // FILTER: Only risks assessed in Menu 4 with "Tinggi" or "Sangat Tinggi" level
      const filtered = data.filter(r => r.resScore > 0 && r.resRiskLevel >= 3);

      // SORT: Primarily by new residual level, then by score descending
      filtered.sort((a, b) => {
        if (b.resRiskLevel !== a.resRiskLevel) return b.resRiskLevel - a.resRiskLevel;
        return b.resScore - a.resScore;
      });

      setRows(filtered);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;
      setLoading(true);
      setShowPasteModal(false);

      const batch = writeBatch(db);
      let count = 0;

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 2) continue;

        const [kode, action, pj, deadline] = cols;
        const targetRow = rows.find(r => r.risikoKode === kode.trim());
        if (!targetRow) continue;

        batch.update(doc(db, 'risk_identification', targetRow.id), {
          rtpAction: (action || '').trim(),
          rtpPJ: (pj || '').trim(),
          rtpDeadline: (deadline || '').trim(),
          updatedAt: new Date().toISOString()
        });
        count++;
      }
      
      if (count > 0) {
        await batch.commit();
      }
      
      setPasteData('');
      alert(`Berhasil mengimpor ${count} data rencana penanganan.`);
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const updateTreatmentField = async (id: string, field: string, value: string) => {
    if (isReadOnly) return;
    try {
      await updateDoc(doc(db, 'risk_identification', id), {
        [field]: value,
        updatedAt: new Date().toISOString()
      });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, `risk_identification/${id}`);
    }
  };

  if (loading) {
    return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600"></div></div>;
  }

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
          <h4 className="font-bold text-slate-700 uppercase border-l-4 border-indigo-500 pl-3 italic text-xs">Rencana Tindak Pengendalian (RTP) - Berdasarkan Risiko Residual</h4>
          {!isReadOnly && (
            <button 
              onClick={() => setShowPasteModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-all text-[10px] font-bold uppercase tracking-wider"
            >
              <ClipboardList size={14} /> Paste RTP dari Excel
            </button>
          )}
        </div>

        {showPasteModal && (
          <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className="p-4 bg-indigo-600 flex items-center justify-between">
                <div>
                  <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste RTP dari Excel</h3>
                  <p className="text-indigo-100 text-[9px] font-bold mt-0.5">Kolom: Kode | Rencana RTP | PJ | Deadline</p>
                </div>
                <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
              </div>
              <div className="p-6 flex-1">
                <textarea 
                  autoFocus
                  placeholder="Paste di sini..."
                  className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:border-indigo-500 resize-none font-mono"
                  value={pasteData}
                  onChange={(e) => setPasteData(e.target.value)}
                ></textarea>
              </div>
              <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
                <button onClick={processImport} disabled={!pasteData.trim()} className="px-8 py-2 bg-indigo-600 text-white font-black text-[10px] uppercase rounded-lg">Impor</button>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-[10px] text-left">
            <thead className="bg-slate-900 text-white uppercase font-bold tracking-wider">
              <tr>
                <th className="px-2 py-4 border-r border-slate-800">No</th>
                <th className="px-3 py-4 border-r border-slate-800">Risiko (Residual)</th>
                <th className="px-3 py-4 border-r border-slate-800">Pengendalian yang sudah ada</th>
                <th className="px-3 py-4 border-r border-slate-800">Sisa Celah</th>
                <th className="px-3 py-4 border-r border-slate-800">Rencana Tindak (RTP) Baru</th>
                <th className="px-3 py-4 border-r border-slate-800">PJ</th>
                <th className="px-3 py-4">Deadline</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {rows.map((row, idx) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors align-top">
                  <td className="px-2 py-4 font-bold text-slate-400">{idx + 1}</td>
                  <td className="px-3 py-4">
                    <div className="flex items-center gap-2 mb-1">
                      <span className="font-black text-blue-600 text-[9px]">{row.risikoKode || (riskType === 'operasional' ? 'ROO...' : 'RSO...')}</span>
                      <span className="px-1.5 py-0.5 bg-slate-100 text-[7px] font-black rounded uppercase">Score: {row.resScore.toFixed(2)}</span>
                    </div>
                    <p className="font-medium text-slate-700 uppercase text-[9px] leading-tight line-clamp-2">{row.risikoUraian}</p>
                  </td>
                  <td className="px-3 py-4">
                    <p className="text-[9px] italic text-slate-500 bg-slate-50 p-2 rounded border border-slate-100">{row.rtpControl || '-'}</p>
                  </td>
                  <td className="px-3 py-4">
                    <p className="text-[9px] italic text-red-400 bg-red-50/30 p-2 rounded border border-red-50">{row.rtpGap || '-'}</p>
                  </td>
                  <td className="px-3 py-4">
                    <EditableTextarea 
                      className="w-full bg-white border border-slate-200 rounded p-1 outline-none resize-none text-[9px] focus:border-blue-500" 
                      rows={3}
                      placeholder="..."
                      value={row.rtpAction || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateTreatmentField(row.id, 'rtpAction', val)}
                    />
                  </td>
                  <td className="px-3 py-4">
                    <EditableInput 
                      className="w-full bg-white border border-slate-200 rounded p-1 outline-none text-[9px] font-medium focus:border-blue-500"
                      placeholder="PJ..."
                      value={row.rtpPJ || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateTreatmentField(row.id, 'rtpPJ', val)}
                    />
                  </td>
                  <td className="px-3 py-4">
                    <EditableInput 
                      type="date"
                      className="w-full bg-white border border-slate-200 rounded p-1 outline-none text-[9px] font-bold text-slate-600 appearance-none cursor-pointer"
                      value={row.rtpDeadline || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateTreatmentField(row.id, 'rtpDeadline', val)}
                    />
                  </td>
                </tr>
              ))}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center text-slate-400 italic">Silahkan isi penilaian Risiko Aktual di Menu 4 terlebih dahulu.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function RiskMapView({ user, riskType }: { user: any, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Listen to Risk Identification Data for plotting
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const d = doc.data();
        
        // Use residual if available, otherwise fallback to avg initial
        const hasResidual = d.residualDampak && d.residualKemungkinan;
        
        const dScore = hasResidual 
          ? parseFloat(d.residualDampak)
          : (d.dampakScores ? d.dampakScores.reduce((a: any, b: any) => a + b, 0) / d.dampakScores.length : 0);
          
        const kScore = hasResidual
          ? parseFloat(d.residualKemungkinan)
          : (d.kemungkinanScores ? d.kemungkinanScores.reduce((a: any, b: any) => a + b, 0) / d.kemungkinanScores.length : 0);

        return {
          ...d,
          id: doc.id,
          avgDampak: Math.max(1, Math.min(5, Math.ceil(dScore))),
          avgKemungkinan: Math.max(1, Math.min(5, Math.ceil(kScore))),
          isResidual: !!hasResidual
        };
      });
      setRows(data);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid, user.role]);

  const cells = Array.from({ length: 25 }, (_, i) => {
    const row = 5 - Math.floor(i / 5); // Impact (D)
    const col = (i % 5) + 1; // Likelihood (K)
    const risk = getRiskLevel(row, col);
    
    return { row, col, color: risk.color };
  });

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  return (
    <div className="max-w-6xl mx-auto space-y-8" id="heatmap-view-root">
      <div className="bg-white p-8 rounded-2xl border border-slate-200 shadow-xl" id="heatmap-container">
        <h3 className="text-center font-black text-2xl uppercase tracking-tighter mb-8 italic">Heatmap Profil Risiko {riskType === 'operasional' ? 'Operasional' : 'Strategis'}</h3>
        
        <div className="flex flex-col lg:flex-row gap-8">
          <div className="flex gap-4 flex-1">
            {/* Label Y */}
            <div className="flex flex-col justify-between py-12">
              <span className="text-[10px] font-black uppercase -rotate-90 origin-center whitespace-nowrap text-slate-400">Impact (Dampak)</span>
            </div>
            
            <div className="flex-1">
              <div className="grid grid-cols-5 gap-1 aspect-square border-4 border-slate-900 bg-slate-900 relative">
                {cells.map((cell, idx) => (
                  <div key={idx} className={`${cell.color} border border-slate-900/10 flex items-center justify-center relative shadow-inner overflow-hidden`}>
                    <div className="flex flex-wrap gap-1 p-1 justify-center">
                      {rows.filter(r => r.avgDampak === cell.row && r.avgKemungkinan === cell.col).map((r, rIdx) => (
                        <div 
                          key={r.id} 
                          className="w-6 h-6 bg-slate-900 text-white rounded-full flex items-center justify-center border-2 border-white shadow-xl text-[8px] font-bold"
                          title={`${r.risikoKode}: ${r.risikoUraian}`}
                        >
                          {rIdx + 1}
                        </div>
                      ))}
                    </div>
                    <span className="absolute bottom-1 right-1 opacity-20 text-[8px] font-bold">{cell.col},{cell.row}</span>
                  </div>
                ))}
              </div>
              {/* Label X */}
              <div className="mt-4 flex justify-center">
                <span className="text-[10px] font-black uppercase text-slate-400 tracking-widest">Probability (Kemungkinan)</span>
              </div>
            </div>
          </div>

          {/* Legend & List */}
          <div className="w-full lg:w-72 space-y-4">
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
              <h5 className="text-[10px] font-bold uppercase mb-3 text-slate-500">Daftar Titik Risiko</h5>
              <div className="space-y-2 max-h-96 overflow-y-auto pr-2">
                {rows.map((r, idx) => (
                  <div key={r.id} className="text-[9px] flex gap-2 items-start bg-white p-2 rounded border border-slate-200">
                     <span className="bg-slate-900 text-white w-4 h-4 rounded-full flex-shrink-0 flex items-center justify-center font-bold">{idx + 1}</span>
                     <div>
                       <p className="font-bold text-blue-600 leading-none mb-1">{r.risikoKode || (riskType === 'operasional' ? 'ROO...' : 'RSO...')}</p>
                       <p className="text-slate-500 line-clamp-2">{r.risikoUraian || 'N/A'}</p>
                       <p className="text-[8px] font-black mt-1 text-slate-400 tracking-tighter">POSISI: {r.avgKemungkinan}, {r.avgDampak}</p>
                     </div>
                  </div>
                ))}
                {rows.length === 0 && <p className="text-[10px] text-slate-400 italic">Belum ada data.</p>}
              </div>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
              <h5 className="text-[10px] font-bold uppercase mb-2 text-slate-500">Keterangan Level</h5>
              <div className="space-y-2">
                <div className="flex items-center gap-2"><div className="w-3 h-3 bg-red-600 rounded-sm" /> <span className="text-[9px] uppercase font-bold text-slate-600">Sangat Tinggi</span></div>
                <div className="flex items-center gap-2"><div className="w-3 h-3 bg-amber-500 rounded-sm" /> <span className="text-[9px] uppercase font-bold text-slate-600">Tinggi</span></div>
                <div className="flex items-center gap-2"><div className="w-3 h-3 bg-yellow-400 rounded-sm" /> <span className="text-[9px] uppercase font-bold text-slate-600">Sedang</span></div>
                <div className="flex items-center gap-2"><div className="w-3 h-3 bg-green-500 rounded-sm border border-slate-200" /> <span className="text-[9px] uppercase font-bold text-slate-600">Rendah</span></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function SuggestionInput({ value, onChange, suggestions, disabled, placeholder, className }: { value: string, onChange: (val: string) => void, suggestions: string[], disabled?: boolean, placeholder?: string, className?: string }) {
  const [show, setShow] = useState(false);
  const [filtered, setFiltered] = useState<string[]>([]);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (value && show) {
      const q = value.toLowerCase();
      const matches = suggestions.filter(s => s.toLowerCase().includes(q) && s.toLowerCase() !== q);
      setFiltered(matches.slice(0, 5));
    } else {
      setFiltered([]);
    }
  }, [value, suggestions, show]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setShow(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Tab' && filtered.length > 0) {
      e.preventDefault();
      onChange(filtered[0]);
      setShow(false);
    }
  };

  return (
    <div ref={containerRef} className="relative w-full">
      <EditableInput 
        value={value}
        onChange={(val) => {
          onChange(val);
          setShow(true);
        }}
        disabled={disabled}
        placeholder={placeholder}
        className={className}
        onFocus={() => setShow(true)}
        onKeyDown={handleKeyDown}
      />
      {show && filtered.length > 0 && (
        <div className="absolute z-[100] left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-xl overflow-hidden animate-in fade-in slide-in-from-top-1 duration-150">
          {filtered.map((s, i) => (
            <button
              key={i}
              type="button"
              onClick={() => {
                onChange(s);
                setShow(false);
              }}
              className="w-full px-3 py-2 text-left text-[10px] font-bold text-slate-600 hover:bg-blue-50 hover:text-blue-600 transition-colors border-b border-slate-50 last:border-0 uppercase tracking-tighter"
            >
              {s}
            </button>
          ))}
          <div className="p-1 px-3 bg-slate-50 text-[8px] text-slate-400 font-bold uppercase tracking-widest text-right">Tekan TAB untuk auto-fill</div>
        </div>
      )}
    </div>
  );
}

function RiskIdentificationView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [contextData, setContextData] = useState<any>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<{id: string, type: 'row' | 'sub', rowId?: string, subRows?: any[], subIdx?: number} | null>(null);

  const contextId = `risk_context_${user.uid}_${riskType}`;

  // Listen to Risk Identification
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id }));
      // Sort by order field primarily, then createdAt
      data.sort((a: any, b: any) => {
        if (a.order !== undefined && b.order !== undefined) {
          return a.order - b.order;
        }
        const dateA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
        const dateB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
        return dateA - dateB;
      });
      setRows(data);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid, riskType]);

  // Listen to Context for manual sync data
  useEffect(() => {
    const docRef = doc(db, 'risk_context', contextId);
    const unsubscribe = onSnapshot(docRef, (snapshot) => {
      if (snapshot.exists()) {
        setContextData(snapshot.data());
      }
    });
    return () => unsubscribe();
  }, [contextId]);

  const addRow = useCallback(async (initTujuan = '', initIndikator = '', customOrder?: number) => {
    if (isReadOnly) return;
    const path = 'risk_identification';
    try {
      const newDocRef = doc(collection(db, path));
      const orderValue = customOrder !== undefined 
        ? customOrder 
        : (rows.length > 0 ? Math.max(...rows.map(r => r.order || 0)) + 1 : 1);
      
      const prefix = riskType === 'operasional' ? 'ROO' : 'RSO';
      
      // Calculate auto-increment code segment
      let lastSeq = 0;
      rows.forEach(r => {
        const segs = (r.risikoKode || '').split('.');
        if (segs.length === 5) {
          const seq = parseInt(segs[4]);
          if (!isNaN(seq) && seq > lastSeq) lastSeq = seq;
        }
      });
      const nextSeq = (lastSeq + 1).toString().padStart(2, '0');
      
      // Default segments: PREFIX . 26 . 01 . 01 . NEXT_SEQ
      let baseP1 = '01';
      let baseP2 = '01';
      if (rows.length > 0) {
        const firstRowParts = (rows[0].risikoKode || '').split('.');
        if (firstRowParts.length >= 5) {
          baseP1 = firstRowParts[2] || '01';
          baseP2 = firstRowParts[3] || '01';
        }
      }

      const initialCode = `${prefix}.26.${baseP1}.${baseP2}.${nextSeq}`;
      
      await setDoc(newDocRef, {
        tujuan: initTujuan || '',
        indikator: initIndikator || '',
        risikoUraian: '',
        risikoKode: initialCode,
        pemilik: '',
        // Initialize with one sub-row
        subRows: [
          {
            id: doc(collection(db, 'temp')).id,
            sebabUraian: '',
            sebabSumber: 'Internal',
            control: 'C',
            dampakUraian: '',
            dampakPihak: ''
          }
        ],
        dampakScores: [0, 0, 0, 0, 0],
        kemungkinanScores: [0, 0, 0, 0, 0],
        order: orderValue,
        createdBy: user?.username || 'Unknown',
        createdByUid: user?.uid || '',
        riskType: riskType, 
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    } catch (error: any) {
      console.error('Error adding document:', error);
      alert('Gagal menambah data: ' + error.message);
      handleFirestoreError(error, OperationType.CREATE, path);
    }
  }, [rows, user?.username, user?.uid, riskType]);

  useEffect(() => {
    const handleAddRow = () => addRow();
    window.addEventListener('add-risk-row', handleAddRow);
    return () => window.removeEventListener('add-risk-row', handleAddRow);
  }, [addRow]);

  const commitDeleteRow = async () => {
    if (!deleteTarget || isReadOnly) return;
    const { id } = deleteTarget;
    try {
      await deleteDoc(doc(db, 'risk_identification', id));
      setDeleteTarget(null);
    } catch (err: any) {
      console.error('Error deleting document:', err);
      alert('Gagal menghapus: ' + err.message);
      handleFirestoreError(err, OperationType.DELETE, `risk_identification/${id}`);
    }
  };

  const commitRemoveSubRow = async () => {
    if (!deleteTarget || isReadOnly) return;
    const { rowId, subRows, subIdx } = deleteTarget;
    if (!rowId || !subRows || subIdx === undefined) return;

    const newSubRows = subRows.filter((_, i) => i !== subIdx);
    try {
      await updateDoc(doc(db, 'risk_identification', rowId), {
        subRows: newSubRows,
        updatedAt: new Date().toISOString()
      });
      setDeleteTarget(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `risk_identification/${rowId}`);
    }
  };

  const deleteRow = (id: string) => {
    setDeleteTarget({ id, type: 'row' });
  };

  const removeSubRow = (rowId: string, subRows: any[], subIdx: number) => {
    setDeleteTarget({ id: `${rowId}_${subIdx}`, type: 'sub', rowId, subRows, subIdx });
  };

  // Handle sub-row operations
  const updateSubRowField = async (rowId: string, subRows: any[], subIdx: number, field: string, value: any) => {
    if (isReadOnly) return;
    const newSubRows = [...subRows];
    newSubRows[subIdx] = { ...newSubRows[subIdx], [field]: value };
    try {
      await updateDoc(doc(db, 'risk_identification', rowId), {
        subRows: newSubRows,
        updatedAt: new Date().toISOString()
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `risk_identification/${rowId}`);
    }
  };

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');
  const [groupPasteTarget, setGroupPasteTarget] = useState<{tujuan: string, indikator: string, order: number} | null>(null);
  const [causePasteRowId, setCausePasteRowId] = useState<string | null>(null);
  const [pasteDataCause, setPasteDataCause] = useState('');

  const processImportCause = async () => {
    if (!pasteDataCause.trim() || !causePasteRowId) return;
    try {
      const rawRows = pasteDataCause.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;

      setLoading(true);
      const targetRow = rows.find(r => r.id === causePasteRowId);
      if (!targetRow) {
        setLoading(false);
        return;
      }

      let newSubRows = [...(targetRow.subRows || [])];
      // If there's only one subrow and it's mostly empty, remove it before appending
      if (newSubRows.length === 1 && !newSubRows[0].sebabUraian?.trim() && !newSubRows[0].dampakUraian?.trim()) {
        newSubRows = [];
      }

      for (const line of rawRows) {
        const cols = line.split('\t');
        const [uraian, sumber, control, akibat, pihak] = cols;
        newSubRows.push({
          sebabUraian: (uraian || '').trim(),
          sebabSumber: (sumber || '').trim(),
          control: (control || '').trim().toUpperCase(),
          dampakUraian: (akibat || '').trim(),
          dampakPihak: (pihak || '').trim()
        });
      }

      await updateDoc(doc(db, 'risk_identification', causePasteRowId), {
        subRows: newSubRows,
        updatedAt: new Date().toISOString()
      });

      setPasteDataCause('');
      setCausePasteRowId(null);
      alert('Berhasil mengimpor sebab/dampak.');
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const handleImportFromClipboard = async () => {
    if (isReadOnly) return;
    setShowPasteModal(true);
  };

  const processImport = async () => {
    if (!pasteData.trim()) return;
    
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;

      setLoading(true);
      setShowPasteModal(false);
      const prefix = riskType === 'operasional' ? 'ROO' : 'RSO';
      
      let baseP1 = '01';
      let baseP2 = '01';
      if (rows.length > 0) {
        const firstRowParts = (rows[0].risikoKode || '').split('.');
        if (firstRowParts.length >= 5) {
          baseP1 = firstRowParts[2] || '01';
          baseP2 = firstRowParts[3] || '01';
        }
      }

      let lastSeq = 0;
      rows.forEach(r => {
        const segs = (r.risikoKode || '').split('.');
        if (segs.length === 5) {
          const seq = parseInt(segs[4]);
          if (!isNaN(seq) && seq > lastSeq) lastSeq = seq;
        }
      });

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 3) continue;

        const [risikoUraian, pemilik, sebabUraian, sebabSumber, control, dampakUraian, dampakPihak] = cols;

        lastSeq++;
        const nextSeq = lastSeq.toString().padStart(2, '0');
        const risikoKode = `${prefix}.26.${baseP1}.${baseP2}.${nextSeq}`;

        const newDocRef = doc(collection(db, 'risk_identification'));
        const currentOrder = groupPasteTarget ? (groupPasteTarget.order + (lastSeq * 0.001)) : (rows.length + lastSeq);

        await setDoc(newDocRef, {
          tujuan: groupPasteTarget ? groupPasteTarget.tujuan : '',
          indikator: groupPasteTarget ? groupPasteTarget.indikator : '',
          risikoUraian: (risikoUraian || '').trim(),
          risikoKode: risikoKode,
          pemilik: (pemilik || '').trim(),
          subRows: [{
            id: doc(collection(db, 'temp')).id,
            sebabUraian: (sebabUraian || '').trim(),
            sebabSumber: (sebabSumber || '').trim(),
            control: (control || '').trim().toUpperCase(),
            dampakUraian: (dampakUraian || '').trim(),
            dampakPihak: (dampakPihak || '').trim()
          }],
          order: currentOrder,
          riskType,
          createdByUid: user.uid,
          username: user.username,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        });
      }

      setPasteData('');
      setGroupPasteTarget(null);
      alert(`Berhasil mengimpor data.`);
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const addSubRow = async (rowId: string, subRows: any[]) => {
    if (isReadOnly) return;
    const newSubRows = [
      ...subRows,
      {
        sebabUraian: '',
        sebabSumber: '',
        control: '',
        dampakUraian: '',
        dampakPihak: ''
      }
    ];
    try {
      await updateDoc(doc(db, 'risk_identification', rowId), {
        subRows: newSubRows,
        updatedAt: new Date().toISOString()
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `risk_identification/${rowId}`);
    }
  };
  const runSync = async () => {
    if (isReadOnly) return;
    
    if (!contextData || !contextData.assessmentRows || contextData.assessmentRows.length === 0) {
      alert(`Menu 1 (Konfigurasi Konteks) masih kosong. Pastikan tabel di bagian bawah Menu 1 sudah diisi dengan ${riskType === 'operasional' ? 'Subkegiatan dan Indikator Keluaran' : 'Sasaran Strategis dan IKU'} agar dapat disinkronkan ke Menu 2.`);
      return;
    }

    // Filter valid rows from Menu 1 bottom table
    const contextPairsList = (contextData.assessmentRows || [])
        .filter((r: any) => {
          if (riskType === 'operasional') {
            return (r.program || '').trim() || (r.iku || '').trim();
          }
          return (r.sasaran || '').trim() || (r.iku || '').trim();
        })
        .map((r: any) => {
          const tujuanVal = riskType === 'operasional' ? (r.program || '').trim() : (r.sasaran || '').trim();
          const ikuVal = (r.iku || '').trim();
          
          return {
            tujuan: tujuanVal,
            iku: ikuVal,
            key: `${tujuanVal.toLowerCase()}|${ikuVal.toLowerCase()}`
          };
        });
    
    if (contextPairsList.length === 0) {
      alert("Tabel di Menu 1 belum memiliki data yang cukup (Subkegiatan/IKU) untuk disinkronkan.");
      return;
    }

    if (!window.confirm("Apakah Anda yakin ingin melakukan sinkronisasi otomatis?\n\nSinkronisasi akan menambahkan baris baru di Menu 2 berdasarkan data dari Menu 1. Baris yang sudah ada tidak akan dihapus.")) {
      return;
    }

    setIsSyncing(true);
    try {
      let batch = writeBatch(db);
      let batchCount = 0;
      
      const existingKeys = new Set(rows.map(r => 
        `${(r.tujuan || '').trim().toLowerCase()}|${(r.indikator || '').trim().toLowerCase()}`
      ));

      let addedCount = 0;
      let lastSeq = 0;
      rows.forEach(r => {
        const segs = (r.risikoKode || '').split('.');
        if (segs.length === 5) {
          const seq = parseInt(segs[4]);
          if (!isNaN(seq) && seq > lastSeq) lastSeq = seq;
        }
      });
      
      const currentMaxOrder = rows.length > 0 ? Math.max(...rows.map(r => r.order || 0)) : 0;
      const prefix = riskType === 'operasional' ? 'ROO' : 'RSO';
      const year = '26'; // Default for the app's current context
      
      let baseP1 = '01';
      let baseP2 = '01';
      if (rows.length > 0) {
        const firstRowParts = (rows[0].risikoKode || '').split('.');
        if (firstRowParts.length >= 5) {
          baseP1 = firstRowParts[2] || '01';
          baseP2 = firstRowParts[3] || '01';
        }
      }

      for (const p of contextPairsList) {
        if (!existingKeys.has(p.key) && p.key !== '|') {
          const newDocRef = doc(collection(db, 'risk_identification'));
          addedCount++;
          const nextSeq = (lastSeq + addedCount).toString().padStart(2, '0');
          const risikoKode = `${prefix}.${year}.${baseP1}.${baseP2}.${nextSeq}`;

          batch.set(newDocRef, {
            tujuan: p.tujuan,
            indikator: p.iku,
            risikoUraian: '',
            risikoKode: risikoKode,
            pemilik: '',
            subRows: [{
              sebabUraian: '',
              sebabSumber: 'Internal',
              control: 'C',
              dampakUraian: '',
              dampakPihak: ''
            }],
            dampakScores: [0, 0, 0, 0, 0],
            kemungkinanScores: [0, 0, 0, 0, 0],
            order: currentMaxOrder + addedCount,
            riskType,
            username: user.username || user.email || 'User',
            createdBy: user.username || user.email || 'User',
            createdByUid: user.uid,
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          });
          
          existingKeys.add(p.key);
          batchCount++;
          if (batchCount >= 400) {
            await batch.commit();
            batch = writeBatch(db);
            batchCount = 0;
          }
        }
      }

      if (batchCount > 0) {
        await batch.commit();
      }

      if (addedCount > 0) {
        alert(`Sinkronisasi berhasil! Menambah ${addedCount} data baru ke Menu 2.`);
      } else {
        alert("Semua data dari Menu 1 sudah ada di Menu 2. Tidak ada data baru yang ditambahkan.");
      }
    } catch (err) {
      console.error("Sync Error:", err);
      alert("Terjadi kesalahan saat sinkronisasi: " + (err instanceof Error ? err.message : String(err)));
    } finally {
      setIsSyncing(false);
    }
  };

  const updateField = async (id: string, field: string, value: string) => {
    const path = `risk_identification/${id}`;
    try {
      await setDoc(doc(db, 'risk_identification', id), { [field]: value }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, path);
    }
  };

  // Filter rows based on role (redundant if query is already filtered, but kept for extra safety)
  const displayRows = rows;

  // Helper to handle risk code segment updates
  const updateRiskCodeSegment = async (id: string, currentFullCode: string, segmentIndex: number, newValue: string) => {
    // Only allow digits and max 2 digits
    let val = newValue.replace(/\D/g, '').slice(0, 2);
    // Pad to 2 digits for internal state
    const normalizedVal = val.padStart(2, '0');
    
    const prefix = riskType === 'operasional' ? 'ROO' : 'RSO';
    const defaultCode = `${prefix}.26.01.01.01`;
    let parts = (currentFullCode || defaultCode).split('.');
    
    if (parts.length < 5) {
      parts = [prefix, '26', '01', '01', '01'];
    }

    const partIndex = segmentIndex + 2; 
    parts[partIndex] = normalizedVal;
    const newFullCode = parts.join('.');
    
    // Propagation logic: If first row changes segment 2 or 3 (index 0 or 1), apply to all rows
    if (rows.length > 0 && id === rows[0].id && segmentIndex < 2) {
      try {
        const batch = writeBatch(db);
        rows.forEach(r => {
          let rParts = (r.risikoKode || defaultCode).split('.');
          if (rParts.length < 5) rParts = [prefix, '26', '01', '01', '01'];
          rParts[segmentIndex + 2] = normalizedVal;
          const rNewCode = rParts.join('.');
          
          batch.update(doc(db, 'risk_identification', r.id), { 
            risikoKode: rNewCode,
            updatedAt: new Date().toISOString()
          });
        });
        await batch.commit();
      } catch (err) {
        console.error("Propagation error:", err);
      }
    } else {
      await updateField(id, 'risikoKode', newFullCode);
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-900 text-white flex justify-between items-center">
          <h4 className="font-bold text-[10px] uppercase italic tracking-widest">Formulir Kertas Kerja Identifikasi Risiko {riskType === 'operasional' ? 'Operasional' : 'Strategis'} OPD</h4>
          <span className="text-[10px] text-slate-400">Tahun 2026 {user.role === 'User' ? `(Milik ${user.username})` : '(Izin Luas)'}</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-[10px] text-left border-collapse">
            <thead className="bg-slate-50 text-slate-500 border-b border-slate-200 font-bold uppercase text-center text-[9px]">
              <tr>
                <th className="px-1 py-3 bg-red-50 text-red-600 border-r border-slate-200" rowSpan={2}>Aksi</th>
                <th className="px-1 py-3 border-r border-slate-200" rowSpan={2}>No</th>
                <th className="px-2 py-3 border-r border-slate-200 min-w-[150px]" rowSpan={2}>{riskType === 'operasional' ? 'Subkegiatan' : 'Tujuan/Sasaran Strategis'}</th>
                <th className="px-2 py-3 border-r border-slate-200 min-w-[150px]" rowSpan={2}>{riskType === 'operasional' ? 'Indikator Keluaran' : 'Indikator Kinerja'}</th>
                <th className="px-2 py-3 border-r border-slate-200" colSpan={2}>Risiko</th>
                <th className="px-2 py-3 border-r border-slate-200 min-w-[120px]" rowSpan={2}>Pemilik</th>
                <th className="px-2 py-3 border-r border-slate-200" colSpan={2}>Sebab</th>
                <th className="px-1 py-3 border-r border-slate-200 w-16" rowSpan={2}>C/UC</th>
                <th className="px-2 py-3 border-r border-slate-200" colSpan={2}>Dampak</th>
              </tr>
              <tr className="border-t border-slate-200 bg-slate-50/50">
                <th className="px-2 py-2 border-r border-slate-200 font-black min-w-[200px]">Uraian</th>
                <th className="px-2 py-2 border-r border-slate-200 font-black w-32">Kode</th>
                <th className="px-2 py-2 border-r border-slate-200 font-black min-w-[200px]">Uraian</th>
                <th className="px-2 py-2 border-r border-slate-200 font-black min-w-[100px]">Sumber</th>
                <th className="px-2 py-2 border-r border-slate-200 font-black min-w-[200px]">Uraian Akibat</th>
                <th className="px-2 py-2 font-black min-w-[120px]">Pihak Terkena</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {displayRows.map((row: any, idx: number) => {
                const subRows = row.subRows && row.subRows.length > 0 ? row.subRows : [
                  {
                    sebabUraian: row.sebabUraian || '',
                    sebabSumber: row.sebabSumber || '',
                    control: row.control || '',
                    dampakUraian: row.dampakUraian || '',
                    dampakPihak: row.dampakPihak || ''
                  }
                ];

                return (
                  <React.Fragment key={row.id}>
                    {subRows.map((sub: any, sIdx: number) => (
                      <tr key={`${row.id}-${sIdx}`} className="group hover:bg-slate-50 transition-colors align-top">
                        {sIdx === 0 && (
                          <>
                            <td className="px-1 py-4 text-center border-r border-slate-100 align-middle" rowSpan={subRows.length}>
                              {!isReadOnly && (
                                <button 
                                  type="button"
                                  onClick={() => {
                                    if (row && row.id) {
                                      deleteRow(row.id);
                                    }
                                  }}
                                  className="text-red-500 hover:text-red-700 transition-colors p-1"
                                  title="Hapus Baris"
                                >
                                  <X size={16} />
                                </button>
                              )}
                            </td>
                            <td className="px-2 py-4 border-r border-slate-100 text-center font-bold align-middle" rowSpan={subRows.length}>{idx + 1}</td>
                            <td className="px-2 py-4 border-r border-slate-100 align-middle bg-slate-50/30" rowSpan={subRows.length}>
                              <div className="flex flex-col gap-2">
                                <EditableTextarea 
                                  className="w-full bg-transparent p-0 outline-none resize-none disabled:text-slate-500 cursor-not-allowed" 
                                  rows={3} 
                                  value={row.tujuan} 
                                  onChange={val => updateField(row.id, 'tujuan', val)}
                                  disabled={true}
                                />
                                {!isReadOnly && (
                                  <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-all">
                                    <button 
                                      type="button"
                                      onClick={() => addRow(row.tujuan, row.indikator, (row.order || 0) + 0.001)}
                                      className="self-start text-[8px] font-black uppercase bg-blue-50 text-blue-600 px-2 py-1 rounded hover:bg-blue-100 transition-all flex items-center gap-1 shadow-sm border border-blue-100"
                                    >
                                      <Plus size={8} /> Tambah Risiko
                                    </button>
                                    <button 
                                      type="button"
                                      onClick={() => {
                                        setGroupPasteTarget({ tujuan: row.tujuan, indikator: row.indikator, order: row.order || 0 });
                                        setShowPasteModal(true);
                                      }}
                                      className="self-start text-[8px] font-black uppercase bg-slate-50 text-slate-600 px-2 py-1 rounded hover:bg-slate-100 transition-all flex items-center gap-1 shadow-sm border border-slate-200"
                                    >
                                      <ClipboardList size={8} /> Paste Risiko
                                    </button>
                                  </div>
                                )}
                              </div>
                            </td>
                            <td className="px-2 py-4 border-r border-slate-100 align-middle bg-slate-50/30" rowSpan={subRows.length}>
                              <EditableTextarea 
                                className="w-full bg-transparent p-0 outline-none resize-none disabled:text-slate-500 cursor-not-allowed" 
                                rows={3} 
                                value={row.indikator} 
                                onChange={val => updateField(row.id, 'indikator', val)}
                                disabled={true}
                              />
                            </td>
                            <td className="px-2 py-4 border-r border-slate-100 align-middle" rowSpan={subRows.length}>
                              <EditableTextarea 
                                className="w-full bg-transparent p-0 outline-none resize-none disabled:text-slate-500" 
                                rows={3} 
                                value={row.risikoUraian} 
                                onChange={val => updateField(row.id, 'risikoUraian', val)}
                                disabled={isReadOnly}
                              />
                            </td>
                            <td className="px-2 py-4 border-r border-slate-100 font-mono text-blue-600 font-bold bg-slate-50/50 text-center align-middle" rowSpan={subRows.length}>
                              {(() => {
                                const prefix = riskType === 'operasional' ? 'ROO' : 'RSO';
                                const defaultCode = `${prefix}.26.01.01.01`;
                                const codeParts = (row.risikoKode || defaultCode).split('.');
                                const p1 = codeParts[2] || '01'; // Segment 2
                                const p2 = codeParts[3] || '01'; // Segment 3
                                const p3 = codeParts[4] || '01'; // Segment 4 (Auto)

                                // Locked values for non-first rows
                                const firstRow = rows[0];
                                const firstRowParts = (firstRow?.risikoKode || defaultCode).split('.');
                                const p1Locked = firstRowParts[2] || '01';
                                const p2Locked = firstRowParts[3] || '01';
                                
                                return (
                                  <div className="flex flex-col items-center gap-1 group/code">
                                    <span className="text-[8px] text-slate-400 uppercase font-black tracking-widest leading-none mb-0.5">{prefix}.26</span>
                                    <div className="flex items-center justify-center gap-0.5 text-[10px]">
                                      <EditableInput 
                                        type="text"
                                        maxLength={2}
                                        filter={v => v.replace(/\D/g, '')}
                                        value={idx === 0 ? p1 : p1Locked}
                                        onChange={(val) => updateRiskCodeSegment(row.id, row.risikoKode, 0, val)}
                                        className={`w-6 border border-slate-200 rounded text-center outline-none py-0.5 transition-all ${idx === 0 ? 'bg-white focus:border-blue-400 focus:ring-1 focus:ring-blue-100 hover:border-blue-200' : 'bg-slate-100 text-slate-500 font-normal opacity-70 cursor-not-allowed'}`}
                                        disabled={isReadOnly || idx > 0}
                                        placeholder="01"
                                      />
                                      <span className="text-slate-400 select-none">.</span>
                                      <EditableInput 
                                        type="text"
                                        maxLength={2}
                                        filter={v => v.replace(/\D/g, '')}
                                        value={idx === 0 ? p2 : p2Locked}
                                        onChange={(val) => updateRiskCodeSegment(row.id, row.risikoKode, 1, val)}
                                        className={`w-6 border border-slate-200 rounded text-center outline-none py-0.5 transition-all ${idx === 0 ? 'bg-white focus:border-blue-400 focus:ring-1 focus:ring-blue-100 hover:border-blue-200' : 'bg-slate-100 text-slate-500 font-normal opacity-70 cursor-not-allowed'}`}
                                        disabled={isReadOnly || idx > 0}
                                        placeholder="01"
                                      />
                                      <span className="text-slate-400 select-none">.</span>
                                      <EditableInput 
                                        type="text"
                                        maxLength={2}
                                        filter={v => v.replace(/\D/g, '')}
                                        value={p3}
                                        onChange={(val) => updateRiskCodeSegment(row.id, row.risikoKode, 2, val)}
                                        className="w-7 bg-white border border-blue-400 text-blue-700 rounded text-center font-black outline-none py-0.5 focus:border-blue-600 focus:ring-1 focus:ring-blue-100 shadow-sm transition-all hover:bg-white"
                                        disabled={isReadOnly}
                                        placeholder="01"
                                      />
                                    </div>
                                  </div>
                                );
                              })()}
                            </td>
                            <td className="px-2 py-4 border-r border-slate-100 align-middle" rowSpan={subRows.length}>
                              <SuggestionInput 
                                suggestions={Array.from(new Set(rows.map(r => String(r.pemilik || '')).filter(Boolean)))}
                                value={row.pemilik} 
                                onChange={val => updateField(row.id, 'pemilik', val)}
                                disabled={isReadOnly}
                                className="w-full bg-transparent p-0 outline-none font-bold text-slate-700 disabled:text-slate-500 uppercase tracking-tighter"
                              />
                            </td>
                          </>
                        )}
                        <td className="px-2 py-4 border-r border-slate-100 italic relative group/sub">
                          <EditableTextarea 
                            className="w-full bg-transparent p-0 outline-none resize-none disabled:text-slate-500" 
                            rows={3} 
                            value={sub.sebabUraian} 
                            onChange={val => updateSubRowField(row.id, subRows, sIdx, 'sebabUraian', val)}
                            disabled={isReadOnly}
                          />
                          {!isReadOnly && sIdx === subRows.length - 1 && (
                            <div className="absolute -bottom-1 right-1 flex gap-1 opacity-0 group-hover/sub:opacity-100 transition-all z-10">
                              <button 
                                onClick={() => setCausePasteRowId(row.id)}
                                className="text-[7px] font-black uppercase text-slate-600 bg-white border border-slate-200 px-1 py-0.5 rounded shadow-sm hover:bg-slate-50 flex items-center gap-0.5"
                                title="Paste Multiple Causes from Excel"
                              >
                                <ClipboardList size={8} /> Paste
                              </button>
                              <button 
                                onClick={() => addSubRow(row.id, subRows)}
                                className="text-[7px] font-black uppercase text-blue-600 bg-blue-50 border border-blue-100 px-1 py-0.5 rounded shadow-sm hover:bg-blue-100 flex items-center gap-0.5"
                              >
                                <Plus size={8} /> Tambah Sebab
                              </button>
                            </div>
                          )}
                          {!isReadOnly && subRows.length > 1 && (
                            <button 
                              onClick={() => removeSubRow(row.id, subRows, sIdx)}
                              className="absolute -bottom-1 left-1 text-[7px] font-black uppercase text-red-600 bg-red-50 px-1 py-0.5 rounded shadow-sm opacity-0 group-hover/sub:opacity-100 transition-all hover:bg-red-100 flex items-center gap-0.5 z-10"
                            >
                              <X size={8} /> Hapus
                            </button>
                          )}
                        </td>
                        <td className="px-2 py-4 border-r border-slate-100 text-center">
                          <SuggestionInput 
                            suggestions={Array.from(new Set(rows.flatMap(r => (r.subRows || []).map((s: any) => String(s.sebabSumber || ''))).filter(Boolean)))}
                            value={sub.sebabSumber} 
                            onChange={val => updateSubRowField(row.id, subRows, sIdx, 'sebabSumber', val)}
                            disabled={isReadOnly}
                            className="w-full bg-transparent p-0 outline-none text-center font-bold text-slate-700 disabled:text-slate-400 uppercase tracking-tighter"
                          />
                        </td>
                        <td className="px-1 py-4 border-r border-slate-100 text-center font-black align-middle">
                          <div className="flex flex-col gap-1 items-center justify-center min-h-full">
                            <select 
                              className={`w-full bg-white border rounded text-[10px] py-1.5 px-1 outline-none transition-all shadow-sm font-bold ${sub.control === 'C' ? 'border-green-200 text-green-700' : sub.control === 'UC' ? 'border-red-200 text-red-700' : 'border-slate-200 text-slate-600'}`}
                              value={sub.control === 'C' || sub.control === 'UC' ? sub.control : ''}
                              onChange={(e) => updateSubRowField(row.id, subRows, sIdx, 'control', e.target.value)}
                              disabled={isReadOnly}
                            >
                              <option value="">-</option>
                              <option value="C">C</option>
                              <option value="UC">UC</option>
                            </select>
                          </div>
                        </td>
                        <td className="px-2 py-4 border-r border-slate-100 text-red-600 bg-red-50/10">
                          <EditableTextarea 
                            className="w-full bg-transparent p-0 outline-none resize-none disabled:text-slate-500 font-bold leading-relaxed" 
                            rows={3} 
                            value={sub.dampakUraian} 
                            onChange={val => updateSubRowField(row.id, subRows, sIdx, 'dampakUraian', val)}
                            disabled={isReadOnly}
                          />
                        </td>
                        <td className="px-2 py-4">
                          <SuggestionInput 
                            suggestions={Array.from(new Set(rows.flatMap(r => (r.subRows || []).map((s: any) => String(s.dampakPihak || ''))).filter(Boolean)))}
                            value={sub.dampakPihak} 
                            onChange={val => updateSubRowField(row.id, subRows, sIdx, 'dampakPihak', val)}
                            disabled={isReadOnly}
                            className="w-full bg-transparent p-0 outline-none font-bold text-slate-700 disabled:text-slate-400 uppercase tracking-tighter"
                          />
                        </td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="p-4 bg-slate-50 border-t border-slate-100 flex justify-between items-center">
           {!isReadOnly && (
             <div className="flex gap-4">
               {/* Sync button remains */}
               <button 
                 onClick={runSync}
                 disabled={isSyncing}
                 className={`text-[10px] font-bold px-4 py-2 rounded-lg flex items-center gap-2 uppercase tracking-widest border transition-all ${isSyncing ? 'bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed' : 'text-emerald-600 hover:bg-emerald-50 border-emerald-100'}`}
               >
                 {isSyncing ? (
                   <RotateCw size={14} className="animate-spin" />
                 ) : (
                   <RotateCw size={14} />
                 )}
                 {isSyncing ? 'Sinkronisasi...' : 'Sinkronisasi dari Menu 1'}
               </button>
             </div>
           )}
           <p className="text-[9px] text-slate-400 italic">
             {isReadOnly ? 'Mode Lihat Data (Hanya Baca)' : 'Klik pada teks untuk mengubah isi. Gunakan sinkronisasi untuk menarik data Tujuan & Indikator dari Menu 1.'}
           </p>
        </div>
      </div>

      <AnimatePresence>
        {deleteTarget && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center"
            >
              <div className="w-16 h-16 bg-red-50 rounded-full flex items-center justify-center mx-auto mb-4">
                <Trash2 size={32} className="text-red-500" />
              </div>
              <h3 className="text-xl font-bold text-slate-900 mb-2 font-black uppercase italic tracking-tight">Konfirmasi Hapus</h3>
              <p className="text-slate-500 text-[11px] mb-8 font-medium italic leading-relaxed">
                {deleteTarget.type === 'row' 
                  ? "Apakah anda yakin ingin menghapus baris risiko ini? Data yang terkait di menu lain juga akan terdampak."
                  : "Apakah anda yakin ingin menghapus baris sebab/dampak ini?"}
              </p>
              <div className="flex gap-3">
                <button 
                  onClick={() => setDeleteTarget(null)}
                  className="flex-1 px-4 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold tracking-widest uppercase text-[10px] hover:bg-slate-200 transition-colors"
                >
                  Batal
                </button>
                <button 
                  onClick={deleteTarget.type === 'row' ? commitDeleteRow : commitRemoveSubRow}
                  className="flex-1 px-4 py-3 bg-red-600 text-white rounded-xl font-black tracking-widest uppercase text-[10px] hover:bg-red-700 transition-colors shadow-lg shadow-red-100"
                >
                  Ya, Hapus
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {showPasteModal && (
        <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 flex flex-col max-h-[80vh] animate-in zoom-in-95 duration-200">
            <div className="p-4 bg-blue-600 flex items-center justify-between">
              <div>
                <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste Data Risiko dari Excel</h3>
                <p className="text-blue-100 text-[9px] font-bold mt-0.5">
                  {groupPasteTarget 
                    ? `Menambah risiko untuk: ${groupPasteTarget.tujuan.substring(0, 50)}...` 
                    : 'Copy baris di Excel (Risiko, Pemilik, Sebab, Sumber, C/UC, Akibat, Pihak)'}
                </p>
                <p className="text-blue-200 text-[8px] font-medium mt-1 uppercase tracking-wider">
                  Kolom: Risiko | Pemilik | Sebab | Sumber | C/UC | Akibat | Pihak
                </p>
              </div>
              <button 
                onClick={() => {
                  setShowPasteModal(false);
                  setGroupPasteTarget(null);
                }} 
                className="text-white/80 hover:text-white hover:bg-white/10 p-1.5 rounded-lg transition-all"
              >
                <X size={18} />
              </button>
            </div>
            <div className="p-6 flex-1 overflow-y-auto">
              <EditableTextarea 
                autoFocus
                placeholder="Tekan Ctrl + V di sini untuk menempelkan data dari Excel..."
                className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] font-mono focus:border-blue-500 focus:ring-4 focus:ring-blue-500/5 outline-none transition-all resize-none"
                value={pasteData}
                onChange={(val) => setPasteData(val)}
              />
              <div className="mt-4 flex items-start gap-3 bg-blue-50 p-3 rounded-lg border border-blue-100">
                <div className="bg-blue-100 p-1.5 rounded-md text-blue-600">
                  <ShieldAlert size={14} />
                </div>
                <div className="text-[9px] text-blue-800 font-bold leading-relaxed">
                  Sistem mencocokkan kolom secara otomatis berdasarkan TAB. Pastikan urutan kolom sesuai dengan panduan di atas agar data masuk dengan benar.
                </div>
              </div>
            </div>
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-3">
              <button 
                onClick={() => {
                  setShowPasteModal(false);
                  setGroupPasteTarget(null);
                }}
                className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase tracking-widest hover:bg-slate-200 rounded-lg transition-all"
              >
                Batal
              </button>
              <button 
                onClick={processImport}
                disabled={!pasteData.trim()}
                className="px-8 py-2 bg-blue-600 text-white font-black text-[10px] uppercase tracking-widest hover:bg-blue-700 rounded-lg shadow-lg shadow-blue-500/20 transition-all disabled:opacity-50 disabled:grayscale"
              >
                Proses Impor
              </button>
            </div>
          </div>
        </div>
      )}

      {causePasteRowId && (
        <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 flex flex-col max-h-[80vh] animate-in zoom-in-95 duration-200">
            <div className="p-4 bg-slate-800 flex items-center justify-between">
              <div>
                <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste Sebab & Dampak dari Excel</h3>
                <p className="text-slate-300 text-[9px] font-bold mt-0.5">Copy kolom: Uraian Sebab | Sumber | C/UC | Uraian Akibat | Pihak Terkena</p>
              </div>
              <button onClick={() => setCausePasteRowId(null)} className="text-white/80 hover:text-white hover:bg-white/10 p-1.5 rounded-lg transition-all">
                <X size={18} />
              </button>
            </div>
            <div className="p-6 flex-1">
              <EditableTextarea 
                autoFocus
                placeholder="Paste kolom Sebab, Sumber, C/UC, Akibat, Pihak di sini..."
                className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] font-mono focus:border-slate-500 focus:ring-4 focus:ring-slate-500/5 outline-none transition-all resize-none"
                value={pasteDataCause}
                onChange={(val) => setPasteDataCause(val)}
              />
              <div className="mt-4 flex items-start gap-3 bg-slate-50 p-3 rounded-lg border border-slate-200">
                <div className="text-[9px] text-slate-500 font-bold leading-relaxed italic">
                  * Baris kosong di Excel akan dilewati. Jika risiko hanya memiliki satu sub-baris kosong, maka akan diganti dengan data paste. Jika sudah ada isinya, akan ditambahkan ke bawah.
                </div>
              </div>
            </div>
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-3">
              <button 
                onClick={() => setCausePasteRowId(null)}
                className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase tracking-widest hover:bg-slate-200 rounded-lg transition-all"
              >
                Batal
              </button>
              <button 
                onClick={processImportCause}
                disabled={!pasteDataCause.trim()}
                className="px-8 py-2 bg-slate-900 text-white font-black text-[10px] uppercase tracking-widest hover:bg-black rounded-lg shadow-lg transition-all disabled:opacity-50"
              >
                Impor Sebab
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ContextSettingView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const storageKey = `risk_context_${user.uid}_${riskType}`;
  const [formData, setFormData] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const docRef = doc(db, 'risk_context', storageKey);
    const unsubscribe = onSnapshot(docRef, (snapshot) => {
      if (snapshot.exists()) {
        setFormData(snapshot.data());
      } else {
        // Initial data if none exists
        const initial = {
          namaPemda: 'Pemerintah Provinsi Papua Tengah',
          tahunPenilaian: '2026',
          periodeRenstra: 'Periode Renstra Tahun 2025-2029',
          urusanPemerintahan: 'Pengawasan Pembangunan Daerah',
          opdDinilai: 'Inspektorat Provinsi Papua Tengah',
          sumberData: 'Renstra Inspektorat Provinsi Papua Tengah Tahun 2025-2029',
          tujuanStrategis: 'Mewujudkan Penguatan APIP dalam melaksanakan Tugas Pokok dan Fungsi demi tercapainya Tata Kelola Pemerintahan yang baik, bersih dan transparan yang menerapkan Sistem Administrasi Publik berbasis Digitalisasi Teknologi',
          informasiLain: '-',
          ttdTempat: 'Paniai',
          ttdBulan: 'April 2026',
          ttdJabatan: 'Kepala Dinas XXX',
          ttdNama: '(Nama)',
          ttdPangkat: '(Pangkat)',
          ttdNip: '(NIP)',
          sasaran: [
            'Peningkatan transparansi dan akuntabilitas dalam pengelolaan keuangan daerah guna mencegah korupsi dan penyalahgunaan anggaran',
            'Penerapan sistem pengawasan berbasis teknologi seperti e-audit dan e-monitoring terhadap anggaran dan proyek pemerintah.',
            'Mendorong peran aktif masyarakat dalam pengawasan pemerintahan, melalui sistem pelaporan yang mudah diakses.'
          ],
          ikuSasaran: [
            { name: 'Indeks Pelayanan Publik', target: '1,11' },
            { name: 'Indeks Integritas Nasional', target: '64' }
          ],
          program: [
            'Program penunjang urusan pemerintahan daerah',
            'Program penyelenggaraan pengawasan',
            'Program perumusan kebijakan, pendampingan dan asistensi'
          ],
          ikuProgram: [
            { name: '', target: '' },
            { name: '', target: '' },
            { name: '', target: '' }
          ],
          assessmentRows: [{ tujuan: '', sasaran: '', program: '', iku: '' }],
          footerVenue: 'Kabupaten Paniai',
          footerDate: 'April 2026',
          createdByUid: user.uid
        };
        setDoc(docRef, initial).catch(e => handleFirestoreError(e, OperationType.WRITE, `risk_context/${storageKey}`));
        setFormData(initial);
      }
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.GET, `risk_context/${storageKey}`);
    });
    return () => unsubscribe();
  }, [storageKey, user.uid]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');
  const [pasteType, setPasteType] = useState<'sasaran' | 'ikuSasaran' | 'program' | 'ikuProgram' | 'assessment'>('assessment');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;

      setLoading(true);
      setShowPasteModal(false);
      
      const updates: any = {};

      if (pasteType === 'sasaran') {
        const newData = rawRows.map(line => line.trim()).filter(Boolean);
        updates.sasaran = [...(formData.sasaran || []), ...newData];
      } else if (pasteType === 'ikuSasaran') {
        const newData = rawRows.map(line => {
          const cols = line.split('\t');
          return { name: (cols[0] || '').trim(), target: (cols[1] || '').trim() };
        }).filter(item => item.name);
        updates.ikuSasaran = [...(formData.ikuSasaran || []), ...newData];
      } else if (pasteType === 'program') {
        const newData = rawRows.map(line => line.trim()).filter(Boolean);
        updates.program = [...(formData.program || []), ...newData];
      } else if (pasteType === 'ikuProgram') {
        const newData = rawRows.map(line => {
          const cols = line.split('\t');
          return { name: (cols[0] || '').trim(), target: (cols[1] || '').trim() };
        }).filter(item => item.name);
        updates.ikuProgram = [...(formData.ikuProgram || []), ...newData];
      } else {
        const newAssessmentRows = [...formData.assessmentRows];
        for (const line of rawRows) {
          const cols = line.split('\t');
          if (cols.length < 1) continue;
          const [tujuan, sasaran, program, iku] = cols;
          newAssessmentRows.push({
            tujuan: (tujuan || '').trim(),
            sasaran: (sasaran || '').trim(),
            program: (program || '').trim(),
            iku: (iku || '').trim()
          });
        }
        updates.assessmentRows = newAssessmentRows;
      }

      await updateData(updates);
      setPasteData('');
      alert('Berhasil mengimpor data.');
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const openPasteModal = (type: typeof pasteType) => {
    setPasteType(type);
    setPasteData('');
    setShowPasteModal(true);
  };

  const updateData = async (updates: any) => {
    try {
      await setDoc(doc(db, 'risk_context', storageKey), updates, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `risk_context/${storageKey}`);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    if (isReadOnly) return;
    updateData({ [e.target.name]: e.target.value });
  };

  if (loading || !formData) return <div className="text-center py-10 text-slate-400">Memuat konteks...</div>;

  const TableHeader = ({ title, colSpan = 1, rightLabel = '' }: { title: string, colSpan?: number, rightLabel?: string }) => (
    <thead className="bg-slate-50 border-y border-slate-200">
      <tr className="text-[10px] font-black uppercase tracking-widest text-slate-500">
        <th className="px-4 py-2 border-r border-slate-200 w-12 text-center">No</th>
        <th className="px-4 py-2 text-left" colSpan={colSpan}>{title}</th>
        {rightLabel && <th className="px-4 py-2 w-32 text-center border-l border-slate-200">{rightLabel}</th>}
      </tr>
    </thead>
  );

  return (
    <div className="space-y-8 bg-white p-12 rounded-lg border border-slate-200 shadow-xl max-w-5xl mx-auto text-xs">
      {/* Header Info */}
      <div className="text-right space-y-1 mb-8 invisible">
        <p>Lampiran 5</p>
        <p>Form 2.b</p>
      </div>

      <div className="flex justify-between items-center mb-12 italic border-b-2 border-slate-900 pb-4">
        <h1 className="font-black text-sm uppercase tracking-[0.2em]">
          PENETAPAN KONTEKS RISIKO {riskType === 'operasional' ? 'OPERASIONAL' : 'STRATEGIS'} OPD
        </h1>
      </div>

      {showPasteModal && (
        <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
            <div className="p-4 bg-blue-600 flex items-center justify-between">
              <div>
                <h3 className="text-white font-black text-xs uppercase tracking-widest">
                  Paste {pasteType === 'sasaran' ? (riskType === 'operasional' ? 'Program' : 'Sasaran') : 
                         pasteType === 'ikuSasaran' ? (riskType === 'operasional' ? 'Kegiatan Utama' : 'IKU Sasaran') :
                         pasteType === 'program' ? (riskType === 'operasional' ? 'Subkegiatan' : 'Program Strategis') :
                         pasteType === 'ikuProgram' ? (riskType === 'operasional' ? 'Keluaran' : 'IKU Program') : 'Assessment'} dari Excel
                </h3>
                <p className="text-blue-100 text-[9px] font-bold mt-0.5">
                  {pasteType === 'sasaran' || pasteType === 'program' ? 'Panduan: Copy list data di Excel (Satu kolom saja)' :
                   pasteType === 'ikuSasaran' || pasteType === 'ikuProgram' ? 'Panduan: Copy 2 kolom di Excel (Uraian | Target)' :
                   'Panduan: Copy 4 kolom di Excel (Tujuan | Sasaran | Program | IKU)'}
                </p>
              </div>
              <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
            </div>
            <div className="p-6 flex-1">
              <textarea 
                autoFocus
                placeholder="Paste di sini..."
                className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:border-blue-500 resize-none font-mono"
                value={pasteData}
                onChange={(e) => setPasteData(e.target.value)}
              ></textarea>
            </div>
            <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
              <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
              <button onClick={processImport} disabled={!pasteData.trim()} className="px-8 py-2 bg-blue-600 text-white font-black text-[10px] uppercase rounded-lg">Impor</button>
            </div>
          </div>
        </div>
      )}

      <div className="space-y-1">
        <div className="grid grid-cols-3 gap-2">
          <div className="font-bold">Nama Pemda</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableInput value={formData.namaPemda} onChange={val => updateData({ namaPemda: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-400 flex-1 outline-none font-bold disabled:border-transparent" />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2">
          <div className="font-bold">Tahun Penilaian</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableInput value={formData.tahunPenilaian} onChange={val => updateData({ tahunPenilaian: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-400 flex-1 outline-none font-bold disabled:border-transparent" />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2">
          <div className="font-bold">Periode yang dinilai</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableInput value={formData.periodeRenstra} onChange={val => updateData({ periodeRenstra: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-400 flex-1 outline-none disabled:border-transparent" />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2">
          <div className="font-bold">Urusan Pemerintahan</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableInput value={formData.urusanPemerintahan} onChange={val => updateData({ urusanPemerintahan: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-400 flex-1 outline-none disabled:border-transparent" />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2">
          <div className="font-bold">OPD yang Dinilai</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableInput value={formData.opdDinilai} onChange={val => updateData({ opdDinilai: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-400 flex-1 outline-none disabled:border-transparent" />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2">
          <div className="font-bold">Sumber Data</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableInput value={formData.sumberData} onChange={val => updateData({ sumberData: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-900 flex-1 outline-none font-bold disabled:border-transparent" />
          </div>
        </div>
        <div className="grid grid-cols-3 gap-2 pt-4">
          <div className="font-bold">{riskType === 'operasional' ? 'Tujuan Operasional' : 'Tujuan Strategis'}</div>
          <div className="col-span-2 flex gap-2">
            <span>:</span> 
            <EditableTextarea value={formData.tujuanStrategis} onChange={val => updateData({ tujuanStrategis: val })} disabled={isReadOnly} className="border border-slate-200 p-2 rounded flex-1 outline-none text-[11px] leading-relaxed disabled:border-slate-100 disabled:text-slate-500" rows={3} />
          </div>
        </div>
      </div>

      {/* Tables Section */}
      <div className="space-y-6 border border-slate-900 mt-6">
        <table className="w-full border-collapse">
          <TableHeader title={riskType === 'operasional' ? 'Program' : 'Sasaran Strategis'} />
          <tbody>
            {(formData.sasaran || []).map((s: string, idx: number) => (
              <tr key={idx} className="border-b border-slate-200 last:border-0 grow group">
                <td className="px-4 py-2 border-r border-slate-200 text-center">{idx + 1}</td>
                <td className="px-4 py-2" colSpan={2}>
                  <div className="flex gap-2 items-start">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500" 
                      rows={2} 
                      value={s} 
                      disabled={isReadOnly}
                      onChange={(val) => {
                        const next = [...formData.sasaran];
                        next[idx] = val;
                        updateData({ sasaran: next });
                      }}
                    />
                    {!isReadOnly && (
                      <button 
                        onClick={() => {
                          const label = riskType === 'operasional' ? "program" : "sasaran strategis";
                          if (window.confirm(`Hapus ${label} ini?`)) {
                            const next = formData.sasaran.filter((_: any, i: number) => i !== idx);
                            updateData({ sasaran: next });
                          }
                        }}
                        className="opacity-0 group-hover:opacity-100 text-red-300 hover:text-red-500 transition-all p-1"
                      >
                        <X size={14} />
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
            {!isReadOnly && (
              <tr className="bg-slate-50/50">
                <td colSpan={3} className="px-4 py-2">
                  <div className="flex gap-4">
                    <button 
                      onClick={() => updateData({ sasaran: [...(formData.sasaran || []), ''] })}
                      className="flex items-center gap-2 text-[10px] font-bold text-blue-600 hover:text-blue-700 hover:bg-blue-50 px-2 py-1 rounded transition-all uppercase tracking-wider"
                    >
                      <Plus size={12} /> Tambah {riskType === 'operasional' ? 'Program' : 'Sasaran Strategis'}
                    </button>
                    <button 
                      onClick={() => openPasteModal('sasaran')}
                      className="flex items-center gap-2 text-[10px] font-bold text-slate-600 hover:text-slate-700 hover:bg-slate-100 px-2 py-1 rounded transition-all uppercase tracking-wider border border-slate-200"
                    >
                      <ClipboardList size={12} /> Paste {riskType === 'operasional' ? 'Program' : 'Sasaran'} (Excel)
                    </button>
                  </div>
                </td>
              </tr>
            )}
          </tbody>

          <TableHeader title={riskType === 'operasional' ? 'Kegiatan Utama' : 'IKU Sasaran OPD'} rightLabel={riskType === 'operasional' ? '' : '2026'} />
          <tbody>
            {(formData.ikuSasaran || []).map((i: any, idx: number) => (
              <tr key={idx} className="border-b border-slate-200 last:border-0 group">
                <td className="px-4 py-2 border-r border-slate-200 text-center">{idx + 1}</td>
                <td className="px-4 py-2" colSpan={riskType === 'operasional' ? 2 : 1}>
                  <div className="flex gap-2 items-center">
                    <EditableInput 
                      className="w-full bg-transparent outline-none disabled:text-slate-500" 
                      value={i.name} 
                      disabled={isReadOnly}
                      onChange={(val) => {
                        const next = [...formData.ikuSasaran];
                        next[idx] = { ...next[idx], name: val };
                        updateData({ ikuSasaran: next });
                      }}
                    />
                    {riskType === 'operasional' && !isReadOnly && (
                      <button 
                        onClick={() => {
                          if (window.confirm(`Hapus kegiatan utama ini?`)) {
                            const next = formData.ikuSasaran.filter((_: any, i: number) => i !== idx);
                            updateData({ ikuSasaran: next });
                          }
                        }}
                        className="opacity-0 group-hover:opacity-100 text-red-300 hover:text-red-500 transition-all p-1"
                      >
                        <X size={14} />
                      </button>
                    )}
                  </div>
                </td>
                {riskType !== 'operasional' && (
                  <td className="px-4 py-2 text-center border-l border-slate-200 font-bold">
                    <div className="flex gap-2 items-center">
                      <EditableInput 
                        className="w-full text-center bg-transparent outline-none disabled:text-slate-500" 
                        value={i.target} 
                        disabled={isReadOnly}
                        onChange={(val) => {
                          const next = [...formData.ikuSasaran];
                          next[idx] = { ...next[idx], target: val };
                          updateData({ ikuSasaran: next });
                        }}
                      />
                      {!isReadOnly && (
                        <button 
                          onClick={() => {
                            if (window.confirm(`Hapus IKU sasaran ini?`)) {
                              const next = formData.ikuSasaran.filter((_: any, i: number) => i !== idx);
                              updateData({ ikuSasaran: next });
                            }
                          }}
                          className="opacity-0 group-hover:opacity-100 text-red-300 hover:text-red-500 transition-all p-1"
                        >
                          <X size={14} />
                        </button>
                      )}
                    </div>
                  </td>
                )}
              </tr>
            ))}
            {!isReadOnly && (
              <tr className="bg-slate-50/50">
                <td colSpan={3} className="px-4 py-2">
                  <div className="flex gap-4">
                    <button 
                      onClick={() => updateData({ ikuSasaran: [...(formData.ikuSasaran || []), { name: '', target: '' }] })}
                      className="flex items-center gap-2 text-[10px] font-bold text-blue-600 hover:text-blue-700 hover:bg-blue-50 px-2 py-1 rounded transition-all uppercase tracking-wider"
                    >
                      <Plus size={12} /> Tambah {riskType === 'operasional' ? 'Kegiatan Utama' : 'IKU Sasaran OPD'}
                    </button>
                    <button 
                      onClick={() => openPasteModal('ikuSasaran')}
                      className="flex items-center gap-2 text-[10px] font-bold text-slate-600 hover:text-slate-700 hover:bg-slate-100 px-2 py-1 rounded transition-all uppercase tracking-wider border border-slate-200"
                    >
                      <ClipboardList size={12} /> Paste {riskType === 'operasional' ? 'Kegiatan Utama' : 'IKU'} (Excel)
                    </button>
                  </div>
                </td>
              </tr>
            )}
          </tbody>

          <TableHeader title={riskType === 'operasional' ? 'Subkegiatan' : 'Program Strategis'} />
          <tbody>
            {(formData.program || []).map((p: string, idx: number) => (
              <tr key={idx} className="border-b border-slate-200 last:border-0 group">
                <td className="px-4 py-2 border-r border-slate-200 text-center">{idx + 1}</td>
                <td className="px-4 py-2" colSpan={2}>
                  <div className="flex gap-2 items-center">
                    <EditableInput 
                      className="w-full bg-transparent outline-none disabled:text-slate-500" 
                      value={p} 
                      disabled={isReadOnly}
                      onChange={(val) => {
                        const next = [...formData.program];
                        next[idx] = val;
                        updateData({ program: next });
                      }}
                    />
                    {!isReadOnly && (
                      <button 
                        onClick={() => {
                          const label = riskType === 'operasional' ? "subkegiatan" : "program strategis";
                          if (window.confirm(`Hapus ${label} ini?`)) {
                            const next = formData.program.filter((_: any, i: number) => i !== idx);
                            updateData({ program: next });
                          }
                        }}
                        className="opacity-0 group-hover:opacity-100 text-red-300 hover:text-red-500 transition-all p-1"
                      >
                        <X size={14} />
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
            {!isReadOnly && (
              <tr className="bg-slate-50/50">
                <td colSpan={3} className="px-4 py-2">
                  <div className="flex gap-4">
                    <button 
                      onClick={() => updateData({ program: [...(formData.program || []), ''] })}
                      className="flex items-center gap-2 text-[10px] font-bold text-blue-600 hover:text-blue-700 hover:bg-blue-50 px-2 py-1 rounded transition-all uppercase tracking-wider"
                    >
                      <Plus size={12} /> Tambah {riskType === 'operasional' ? 'Subkegiatan' : 'Program Strategis'}
                    </button>
                    <button 
                      onClick={() => openPasteModal('program')}
                      className="flex items-center gap-2 text-[10px] font-bold text-slate-600 hover:text-slate-700 hover:bg-slate-100 px-2 py-1 rounded transition-all uppercase tracking-wider border border-slate-200"
                    >
                      <ClipboardList size={12} /> Paste {riskType === 'operasional' ? 'Subkegiatan' : 'Program'} (Excel)
                    </button>
                  </div>
                </td>
              </tr>
            )}
          </tbody>

          <TableHeader title={riskType === 'operasional' ? 'Indikator Keluaran' : 'IKU Program OPD'} rightLabel="2026" />
          <tbody>
            {(formData.ikuProgram || []).map((i: any, idx: number) => (
              <tr key={idx} className="border-b border-slate-200 last:border-0 h-8 group">
                <td className="px-4 py-2 border-r border-slate-200 text-center">{idx + 1}</td>
                <td className="px-4 py-2">
                  <EditableInput 
                    className="w-full bg-transparent outline-none disabled:text-slate-500" 
                    placeholder="..." 
                    value={i.name || ''}
                    disabled={isReadOnly}
                    onChange={(val) => {
                      const next = [...(formData.ikuProgram || [])];
                      next[idx] = { ...next[idx], name: val };
                      updateData({ ikuProgram: next });
                    }}
                  />
                </td>
                <td className="px-4 py-2 border-l border-slate-200">
                  <div className="flex gap-2 items-center">
                    <EditableInput 
                      className="w-full text-center bg-transparent outline-none disabled:text-slate-500" 
                      placeholder="..." 
                      value={i.target || ''}
                      disabled={isReadOnly}
                      onChange={(val) => {
                        const next = [...(formData.ikuProgram || [])];
                        next[idx] = { ...next[idx], target: val };
                        updateData({ ikuProgram: next });
                      }}
                    />
                    {!isReadOnly && (
                      <button 
                        onClick={() => {
                          const label = riskType === 'operasional' ? "indikator keluaran" : "IKU program";
                          if (window.confirm(`Hapus ${label} ini?`)) {
                            const next = (formData.ikuProgram || []).filter((_: any, i: number) => i !== idx);
                            updateData({ ikuProgram: next });
                          }
                        }}
                        className="opacity-0 group-hover:opacity-100 text-red-300 hover:text-red-500 transition-all p-1"
                      >
                        <X size={14} />
                      </button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
            {!isReadOnly && (
              <tr className="bg-slate-50/50">
                <td colSpan={3} className="px-4 py-2">
                  <div className="flex gap-4">
                    <button 
                      onClick={() => updateData({ ikuProgram: [...(formData.ikuProgram || []), { name: '', target: '' }] })}
                      className="flex items-center gap-2 text-[10px] font-bold text-blue-600 hover:text-blue-700 hover:bg-blue-50 px-2 py-1 rounded transition-all uppercase tracking-wider"
                    >
                      <Plus size={12} /> Tambah {riskType === 'operasional' ? 'Indikator Keluaran' : 'IKU Program OPD'}
                    </button>
                    <button 
                      onClick={() => openPasteModal('ikuProgram')}
                      className="flex items-center gap-2 text-[10px] font-bold text-slate-600 hover:text-slate-700 hover:bg-slate-100 px-2 py-1 rounded transition-all uppercase tracking-wider border border-slate-200"
                    >
                      <ClipboardList size={12} /> Paste {riskType === 'operasional' ? 'Indikator Keluaran' : 'IKU Program'} (Excel)
                    </button>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="grid grid-cols-3 gap-2 mt-4">
        <div className="font-bold">Informasi lain</div>
        <div className="col-span-2 flex gap-2">
          <span>:</span> 
          <EditableInput value={formData.informasiLain} onChange={val => updateData({ informasiLain: val })} disabled={isReadOnly} className="border-b border-dotted border-slate-400 flex-1 outline-none disabled:border-transparent disabled:text-slate-500" />
        </div>
      </div>

      {/* The Bottom Assessment Section Table */}
      <div className="mt-12 space-y-4">
        <div className="flex justify-between items-end">
          <h4 className="font-black text-[10px] uppercase tracking-widest text-slate-500 bg-slate-50 p-2 border-l-4 border-slate-900">
            {riskType === 'operasional' 
              ? 'Program, Kegiatan, Subkegiatan, dan Keluaran/Hasil Subkegiatan yang akan dilakukan penilaian risiko'
              : 'Tujuan, Sasaran, Program Strategis, IKU Program yang akan dilakukan penilaian risiko'
            }
          </h4>
          {!isReadOnly && (
            <div className="flex gap-2">
              <button 
                onClick={() => openPasteModal('assessment')}
                className="text-[10px] bg-white text-slate-700 border border-slate-200 px-3 py-1.5 rounded font-bold uppercase flex items-center gap-2 hover:bg-slate-50 transition-colors"
              >
                <ClipboardList size={12} /> Paste Assessment (Excel)
              </button>
              <button 
                onClick={() => updateData({ assessmentRows: [...(formData.assessmentRows || []), { tujuan: '', sasaran: '', program: '', iku: '' }] })}
                className="text-[10px] bg-slate-900 text-white px-3 py-1.5 rounded font-black uppercase flex items-center gap-2 hover:bg-slate-800 transition-colors"
              >
                <Plus size={12} /> Tambah Baris
              </button>
            </div>
          )}
        </div>
        
        <div className="border border-slate-900 overflow-hidden shadow-lg">
          <table className="w-full text-[10px] text-left border-collapse">
            <thead className="bg-slate-900 text-white uppercase font-bold tracking-wider text-[9px]">
              <tr className="divide-x divide-slate-800 text-center">
                <th className="px-2 py-3 w-10">No</th>
                <th className="px-4 py-3 w-32">{riskType === 'operasional' ? 'Program' : 'Tujuan Strategis'}</th>
                <th className="px-4 py-3 w-48">{riskType === 'operasional' ? 'Kegiatan' : 'Sasaran Strategis'}</th>
                <th className="px-4 py-3 w-48">{riskType === 'operasional' ? 'Subkegiatan' : 'Program Strategis'}</th>
                <th className="px-4 py-3">{riskType === 'operasional' ? 'Keluaran/Hasil Subkegiatan' : 'IKU Program'}</th>
                <th className={`px-2 py-3 w-12 bg-red-900 ${isReadOnly ? 'hidden' : ''}`}>Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-300">
              {(formData.assessmentRows || []).map((row: any, idx: number) => (
                <tr key={idx} className="align-top divide-x divide-slate-200">
                  <td className="px-2 py-4 text-center font-bold bg-slate-100/50">{idx + 1}</td>
                  <td className="px-2 py-4 bg-slate-50/50">
                    <select 
                      className="w-full bg-transparent outline-none border-none p-1 font-bold h-full min-h-[60px] disabled:opacity-50" 
                      value={row.tujuan}
                      disabled={isReadOnly}
                      onChange={(e) => {
                        const next = [...formData.assessmentRows];
                        next[idx].tujuan = e.target.value;
                        updateData({ assessmentRows: next });
                      }}
                    >
                      <option value="">- {riskType === 'operasional' ? 'Pilih Program' : 'Pilih Tujuan Strategis'} -</option>
                      {riskType === 'operasional' ? (
                        (formData.sasaran || []).map((s: string, sIdx: number) => (
                          <option key={sIdx} value={s}>{s}</option>
                        ))
                      ) : (
                        <option value={formData.tujuanStrategis}>{formData.tujuanStrategis}</option>
                      )}
                    </select>
                  </td>
                  <td className="px-2 py-4">
                    <select 
                      className="w-full bg-transparent outline-none border-none p-1 h-full min-h-[60px] disabled:opacity-50" 
                      value={row.sasaran}
                      disabled={isReadOnly}
                      onChange={(e) => {
                        const next = [...formData.assessmentRows];
                        next[idx].sasaran = e.target.value;
                        updateData({ assessmentRows: next });
                      }}
                    >
                      <option value="">- {riskType === 'operasional' ? 'Pilih Kegiatan' : 'Pilih Sasaran Strategis'} -</option>
                      {riskType === 'operasional' ? (
                        (formData.ikuSasaran || []).map((i: any, iIdx: number) => (
                          <option key={iIdx} value={i.name}>{i.name}</option>
                        ))
                      ) : (
                        (formData.sasaran || []).map((s: string, sIdx: number) => (
                          <option key={sIdx} value={s}>{s}</option>
                        ))
                      )}
                    </select>
                  </td>
                  <td className="px-2 py-4 bg-slate-50/50">
                    <select 
                      className="w-full bg-transparent outline-none border-none p-1 h-full min-h-[60px] disabled:opacity-50" 
                      value={row.program}
                      disabled={isReadOnly}
                      onChange={(e) => {
                        const next = [...formData.assessmentRows];
                        next[idx].program = e.target.value;
                        updateData({ assessmentRows: next });
                      }}
                    >
                      <option value="">- {riskType === 'operasional' ? 'Pilih Subkegiatan' : 'Pilih Program Strategis'} -</option>
                      {(formData.program || []).map((p: string, pIdx: number) => (
                        <option key={pIdx} value={p}>{p}</option>
                      ))}
                    </select>
                  </td>
                  <td className="px-2 py-4">
                    <select 
                      className="w-full bg-transparent outline-none border-none p-1 h-full min-h-[60px] disabled:opacity-50" 
                      value={row.iku}
                      disabled={isReadOnly}
                      onChange={(e) => {
                        const next = [...formData.assessmentRows];
                        next[idx].iku = e.target.value;
                        updateData({ assessmentRows: next });
                      }}
                    >
                      <option value="">- {riskType === 'operasional' ? 'Pilih Keluaran/Hasil Subkegiatan' : 'Pilih IKU Program'} -</option>
                      {(formData.ikuProgram || []).map((i: any, iIdx: number) => (
                        <option key={iIdx} value={i.name}>{i.name}</option>
                      ))}
                    </select>
                  </td>
                  {!isReadOnly && (
                    <td className="px-2 py-4 text-center">
                      <button 
                        onClick={() => {
                          const next = formData.assessmentRows.filter((_: any, i: number) => i !== idx);
                          updateData({ assessmentRows: next });
                        }}
                        className="text-red-500 hover:text-red-700 transition-colors"
                        title="Hapus Baris"
                      >
                        <X size={16} />
                      </button>
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Signature Section */}
      <div className="mt-16 flex justify-end">
        <div className="w-72 text-center space-y-0.5">
          <p className="mb-20">
            <EditableInput 
              value={formData.ttdTempat} 
              onChange={val => updateData({ ttdTempat: val })} 
              disabled={isReadOnly}
              className="w-24 border-b border-slate-200 text-center outline-none bg-transparent hover:bg-slate-50 transition-colors placeholder:text-slate-300 disabled:border-transparent" 
              placeholder="Tempat..." 
            />, 
            <EditableInput 
              value={formData.ttdBulan} 
              onChange={val => updateData({ ttdBulan: val })} 
              disabled={isReadOnly}
              className="w-32 border-b border-slate-200 text-center outline-none bg-transparent hover:bg-slate-50 transition-colors ml-1 placeholder:text-slate-300 disabled:border-transparent" 
              placeholder="Bulan Tahun..." 
            />
            <br />
            <EditableInput 
              value={formData.ttdJabatan} 
              onChange={val => updateData({ ttdJabatan: val })} 
              disabled={isReadOnly}
              className="w-full mt-2 font-bold outline-none text-center bg-transparent hover:bg-slate-50 transition-colors placeholder:text-slate-300" 
              placeholder="Jabatan Penandatangan..."
            />
            <br />
            <EditableInput 
              value={formData.ttdKabupaten || 'Kabupaten Paniai'} 
              onChange={val => updateData({ ttdKabupaten: val })} 
              disabled={isReadOnly}
              className="w-full text-center outline-none bg-transparent hover:bg-blue-50 border-b border-transparent focus:border-blue-200 transition-all font-semibold italic text-slate-600 placeholder:text-slate-300 disabled:border-transparent" 
              placeholder="Kabupaten..." 
            />
          </p>
          <div className="space-y-1">
            <p className="font-bold underline decoration-2 h-6">
              <EditableInput 
                value={formData.ttdNama} 
                onChange={val => updateData({ ttdNama: val })} 
                disabled={isReadOnly}
                className="w-full text-center bg-transparent outline-none placeholder:text-slate-300" 
                placeholder="NAMA LENGKAP"
              />
            </p>
            <p className="text-[10px]">
              <EditableInput 
                value={formData.ttdPangkat} 
                onChange={val => updateData({ ttdPangkat: val })} 
                disabled={isReadOnly}
                className="w-full text-center bg-transparent outline-none placeholder:text-slate-200" 
                placeholder="Pangkat/Golongan"
              />
            </p>
            <p className="font-mono text-[9px] text-slate-500">
              NIP. <EditableInput 
                value={formData.ttdNip} 
                onChange={val => updateData({ ttdNip: val })} 
                disabled={isReadOnly}
                className="w-32 bg-transparent outline-none placeholder:text-slate-300 text-left" 
                placeholder="19xxxxxxxxxxxxxx"
              />
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

function MonitoringCommunicationView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Listen to Risk Identification Data (same logic as Menu 5)
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const d = doc.data();
        
        // Use Residual Risk Assessment
        const resD = parseFloat(d.residualDampak || 0);
        const resK = parseFloat(d.residualKemungkinan || 0);
        const resScore = resD * resK;
        const resRisk = getRiskLevel(resD, resK);

        return { 
          ...d, 
          id: doc.id,
          resScore,
          resRiskLabel: resRisk.label,
          resRiskLevel: resRisk.level
        };
      });

      // Show risks that have been assessed for residual risk with "Tinggi" or "Sangat Tinggi" level
      const filtered = data.filter(r => r.resScore > 0 && r.resRiskLevel >= 3);

      // SORT: Primarily by residual level, then by score descending
      filtered.sort((a, b) => {
        if (b.resRiskLevel !== a.resRiskLevel) return b.resRiskLevel - a.resRiskLevel;
        return b.resScore - a.resScore;
      });

      setRows(filtered);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid, user.role]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;
      setLoading(true);
      setShowPasteModal(false);

      const batch = writeBatch(db);
      let count = 0;

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 2) continue;

        const [kode, media, prov, recv, pTime, rTime, notes] = cols;
        const targetRow = rows.find(r => r.risikoKode === kode.trim());
        if (!targetRow) continue;

        batch.update(doc(db, 'risk_identification', targetRow.id), {
          commMedia: (media || '').trim(),
          commProvider: (prov || '').trim(),
          commReceiver: (recv || '').trim(),
          commPlanTime: (pTime || '').trim(),
          commRealTime: (rTime || '').trim(),
          commNotes: (notes || '').trim(),
          updatedAt: new Date().toISOString()
        });
        count++;
      }

      if (count > 0) {
        await batch.commit();
      }

      setPasteData('');
      alert(`Berhasil mengimpor ${count} data komunikasi.`);
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const updateCommField = async (id: string, field: string, value: string) => {
    try {
      await updateDoc(doc(db, 'risk_identification', id), {
        [field]: value,
        updatedAt: new Date().toISOString()
      });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, `risk_identification/${id}`);
    }
  };

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
          <h4 className="font-bold text-slate-700 uppercase italic">Komunikasi Pengendalian</h4>
          {!isReadOnly && (
            <button 
              onClick={() => setShowPasteModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-all text-[10px] font-bold uppercase tracking-wider"
            >
              <ClipboardList size={14} /> Paste Komunikasi dari Excel
            </button>
          )}
        </div>

        {showPasteModal && (
          <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className="p-4 bg-indigo-600 flex items-center justify-between">
                <div>
                  <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste Komunikasi dari Excel</h3>
                  <p className="text-indigo-100 text-[9px] font-bold mt-0.5">Kolom: Kode | Media | Prov | Recv | Plan | Real | Notes</p>
                </div>
                <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
              </div>
              <div className="p-6 flex-1">
                <textarea 
                  autoFocus
                  placeholder="Paste di sini..."
                  className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:border-indigo-500 resize-none font-mono"
                  value={pasteData}
                  onChange={(e) => setPasteData(e.target.value)}
                ></textarea>
              </div>
              <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
                <button onClick={processImport} disabled={!pasteData.trim()} className="px-8 py-2 bg-indigo-600 text-white font-black text-[10px] uppercase rounded-lg">Impor</button>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-[11px] text-left border-collapse">
            <thead className="bg-slate-50 text-slate-500 uppercase text-[9px] font-black border-b border-slate-200">
              <tr>
                <th className="px-3 py-4 w-8 border-r border-slate-200">No</th>
                <th className="px-3 py-4 w-64 border-r border-slate-200">Kegiatan Pengendalian yang Dibutuhkan</th>
                <th className="px-3 py-4 w-48 border-r border-slate-200">Media/Bentuk Sarana Pengkomunikasian</th>
                <th className="px-3 py-4 w-40 border-r border-slate-200">Penyedia Informasi</th>
                <th className="px-3 py-4 w-40 border-r border-slate-200">Penerima Informasi</th>
                <th className="px-3 py-4 w-32 border-r border-slate-200">Rencana Waktu Pelaksanaan</th>
                <th className="px-3 py-4 w-32 border-r border-slate-200">Realisasi Waktu Pelaksanaan</th>
                <th className="px-3 py-4">Keterangan</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {rows.map((row, idx) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors align-top">
                  <td className="px-3 py-4 font-bold text-slate-400 border-r border-slate-200 text-center">{idx + 1}</td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <p className="font-medium text-slate-700 leading-relaxed italic">{row.rtpAction || '(RTP belum diisi di Menu 5)'}</p>
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500 text-[9px]" 
                      rows={2}
                      placeholder="..."
                      value={row.commMedia || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateCommField(row.id, 'commMedia', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableInput 
                      className="w-full bg-transparent outline-none disabled:text-slate-500 text-[9px]"
                      placeholder="..."
                      value={row.commProvider || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateCommField(row.id, 'commProvider', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableInput 
                      className="w-full bg-transparent outline-none disabled:text-slate-500 text-[9px]"
                      placeholder="..."
                      value={row.commReceiver || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateCommField(row.id, 'commReceiver', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableInput 
                      type="date"
                      className="w-full bg-transparent outline-none font-bold text-slate-600 cursor-pointer disabled:opacity-50 text-[9px]"
                      value={row.commPlanTime || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateCommField(row.id, 'commPlanTime', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500 text-[9px]" 
                      rows={2}
                      placeholder="..."
                      value={row.commRealTime || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateCommField(row.id, 'commRealTime', val)}
                    />
                  </td>
                  <td className="px-3 py-4">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500 text-[9px]" 
                      rows={2}
                      placeholder="..."
                      value={row.commNotes || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateCommField(row.id, 'commNotes', val)}
                    />
                  </td>
                </tr>
              ))}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={8} className="px-6 py-8 text-center text-slate-400 italic">Belum ada risiko High/Very High teridentifikasi.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function MonitoringPlanPIView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const d = doc.data();
        
        // Use Residual Risk Assessment
        const resD = parseFloat(d.residualDampak || 0);
        const resK = parseFloat(d.residualKemungkinan || 0);
        const resScore = resD * resK;
        const resRisk = getRiskLevel(resD, resK);

        return { 
          ...d, 
          id: doc.id,
          resScore,
          resRiskLabel: resRisk.label,
          resRiskLevel: resRisk.level
        };
      });

      // Show risks that have been assessed for residual risk with "Tinggi" or "Sangat Tinggi" level
      const filtered = data.filter(r => r.resScore > 0 && r.resRiskLevel >= 3);
      
      filtered.sort((a, b) => {
        if (b.resRiskLevel !== a.resRiskLevel) return b.resRiskLevel - a.resRiskLevel;
        return b.resScore - a.resScore;
      });

      setRows(filtered);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid, user.role]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;
      setLoading(true);
      setShowPasteModal(false);

      const batch = writeBatch(db);
      let count = 0;

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 2) continue;

        const [kode, method, pj, pTime, rTime, notes] = cols;
        const targetRow = rows.find(r => r.risikoKode === kode.trim());
        if (!targetRow) continue;

        batch.update(doc(db, 'risk_identification', targetRow.id), {
          monMethod: (method || '').trim(),
          monPJ: (pj || '').trim(),
          monPlanTime: (pTime || '').trim(),
          monRealTime: (rTime || '').trim(),
          monNotes: (notes || '').trim(),
          updatedAt: new Date().toISOString()
        });
        count++;
      }

      if (count > 0) {
        await batch.commit();
      }

      setPasteData('');
      alert(`Berhasil mengimpor ${count} data rencana monitoring.`);
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const updateMonField = async (id: string, field: string, value: string) => {
    try {
      await updateDoc(doc(db, 'risk_identification', id), {
        [field]: value,
        updatedAt: new Date().toISOString()
      });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, `risk_identification/${id}`);
    }
  };

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
          <h4 className="font-bold text-slate-700 uppercase italic">RENCANA MONITORING PI</h4>
          {!isReadOnly && (
            <button 
              onClick={() => setShowPasteModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-all text-[10px] font-bold uppercase tracking-wider"
            >
              <ClipboardList size={14} /> Paste Monitoring dari Excel
            </button>
          )}
        </div>

        {showPasteModal && (
          <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className="p-4 bg-slate-600 flex items-center justify-between">
                <div>
                  <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste Monitoring dari Excel</h3>
                  <p className="text-slate-100 text-[9px] font-bold mt-0.5">Kolom: Kode | Metode | PJ | Plan | Real | Notes</p>
                </div>
                <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
              </div>
              <div className="p-6 flex-1">
                <textarea 
                  autoFocus
                  placeholder="Paste di sini..."
                  className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:border-slate-500 resize-none font-mono"
                  value={pasteData}
                  onChange={(e) => setPasteData(e.target.value)}
                ></textarea>
              </div>
              <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
                <button onClick={processImport} disabled={!pasteData.trim()} className="px-8 py-2 bg-slate-600 text-white font-black text-[10px] uppercase rounded-lg">Impor</button>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-[11px] text-left border-collapse">
            <thead className="bg-slate-50 text-slate-500 uppercase text-[9px] font-black border-b border-slate-200">
              <tr>
                <th className="px-3 py-4 w-8 border-r border-slate-200">No</th>
                <th className="px-3 py-4 w-64 border-r border-slate-200">Kegiatan Pengendalian yang Dibutuhkan</th>
                <th className="px-3 py-4 w-48 border-r border-slate-200">Bentuk/Metode Pemantauan yang Diperlukan</th>
                <th className="px-3 py-4 w-40 border-r border-slate-200">Penanggung Jawab Pemantauan</th>
                <th className="px-3 py-4 w-32 border-r border-slate-200">Rencana Waktu Pelaksanaan Pemantauan</th>
                <th className="px-3 py-4 w-32 border-r border-slate-200">Realisasi Waktu Pelaksanaan</th>
                <th className="px-3 py-4">Keterangan</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {rows.map((row, idx) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors align-top">
                  <td className="px-3 py-4 font-bold text-slate-400 border-r border-slate-200 text-center">{idx + 1}</td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <p className="font-medium text-slate-700 leading-relaxed italic">{row.rtpAction || '(RTP belum diisi di Menu 5)'}</p>
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500" 
                      rows={2}
                      placeholder="Input metode..."
                      value={row.monMethod || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateMonField(row.id, 'monMethod', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableInput 
                      className="w-full bg-transparent outline-none disabled:text-slate-500"
                      placeholder="Input PJ..."
                      value={row.monPJ || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateMonField(row.id, 'monPJ', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableInput 
                      type="date"
                      className="w-full bg-transparent outline-none font-bold text-slate-600 cursor-pointer disabled:opacity-50"
                      value={row.monPlanTime || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateMonField(row.id, 'monPlanTime', val)}
                    />
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200">
                    <EditableInput 
                      type="date"
                      className="w-full bg-transparent outline-none font-bold text-slate-600 cursor-pointer disabled:opacity-50"
                      value={row.monRealTime || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateMonField(row.id, 'monRealTime', val)}
                    />
                  </td>
                  <td className="px-3 py-4">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500" 
                      rows={2}
                      placeholder="Input keterangan..."
                      value={row.monNotes || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateMonField(row.id, 'monNotes', val)}
                    />
                  </td>
                </tr>
              ))}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={7} className="px-6 py-8 text-center text-slate-400 italic">Belum ada risiko High/Very High teridentifikasi.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function RiskOccurrenceMonitoringView({ user, isReadOnly, riskType }: { user: any, isReadOnly?: boolean, riskType: 'strategis' | 'operasional' }) {
  const [rows, setRows] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Listen to Risk Identification Data
  useEffect(() => {
    const baseQuery = collection(db, 'risk_identification');
    const q = query(baseQuery, 
      where('createdByUid', '==', user.uid),
      where('riskType', '==', riskType)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const d = doc.data();
        const resD = parseFloat(d.residualDampak || 0);
        const resK = parseFloat(d.residualKemungkinan || 0);
        const resScore = resD * resK;
        const resRisk = getRiskLevel(resD, resK);

        return { 
          ...d, 
          id: doc.id,
          resScore,
          resRiskLabel: resRisk.label,
          resRiskLevel: resRisk.level
        };
      });

      // SORT: Primarily by residual level, then by score descending
      data.sort((a, b) => {
        if (b.resRiskLevel !== a.resRiskLevel) return b.resRiskLevel - a.resRiskLevel;
        return (b.resScore || 0) - (a.resScore || 0);
      });

      setRows(data);
      setLoading(false);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'risk_identification');
    });
    return () => unsubscribe();
  }, [user.uid, user.role]);

  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pasteData, setPasteData] = useState('');

  const processImport = async () => {
    if (!pasteData.trim()) return;
    try {
      const rawRows = pasteData.split(/\r?\n/).filter(line => line.trim() !== '');
      if (rawRows.length === 0) return;
      setLoading(true);
      setShowPasteModal(false);

      const batch = writeBatch(db);
      let count = 0;

      for (const line of rawRows) {
        const cols = line.split('\t');
        if (cols.length < 2) continue;

        const [kode, date, cause, impact, eNotes, rtpPlan, rtpReal, rtpNotes] = cols;
        const targetRow = rows.find(r => r.risikoKode === kode.trim());
        if (!targetRow) continue;

        batch.update(doc(db, 'risk_identification', targetRow.id), {
          eventDate: (date || '').trim(),
          eventCause: (cause || '').trim(),
          eventImpact: (impact || '').trim(),
          eventNotes: (eNotes || '').trim(),
          rtpPlanDate: (rtpPlan || '').trim(),
          rtpRealDate: (rtpReal || '').trim(),
          rtpNotesContent: (rtpNotes || '').trim(),
          updatedAt: new Date().toISOString()
        });
        count++;
      }

      if (count > 0) {
        await batch.commit();
      }

      setPasteData('');
      alert(`Berhasil mengimpor ${count} data monitoring keterjadian.`);
    } catch (err) {
      console.error('Import error:', err);
      alert('Gagal mengimpor data.');
    } finally {
      setLoading(false);
    }
  };

  const updateField = async (id: string, field: string, value: string) => {
    if (isReadOnly) return;
    try {
      await updateDoc(doc(db, 'risk_identification', id), {
        [field]: value,
        updatedAt: new Date().toISOString()
      });
    } catch (err: any) {
      handleFirestoreError(err, OperationType.UPDATE, `risk_identification/${id}`);
    }
  };

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
          <h4 className="font-bold text-slate-700 uppercase italic text-sm">Monitoring Keterjadian Risiko</h4>
          {!isReadOnly && (
            <button 
              onClick={() => setShowPasteModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-all text-[10px] font-bold uppercase tracking-wider"
            >
              <ClipboardList size={14} /> Paste Keterjadian dari Excel
            </button>
          )}
        </div>

        {showPasteModal && (
          <div className="fixed inset-0 z-[200] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 flex flex-col max-h-[80vh]">
              <div className="p-4 bg-slate-600 flex items-center justify-between">
                <div>
                  <h3 className="text-white font-black text-xs uppercase tracking-widest">Paste Keterjadian dari Excel</h3>
                  <p className="text-slate-100 text-[9px] font-bold mt-0.5">Kolom: Kode | Uraian | Waktu | Catatan</p>
                </div>
                <button onClick={() => setShowPasteModal(false)} className="text-white/80 hover:text-white"><X size={18} /></button>
              </div>
              <div className="p-6 flex-1">
                <textarea 
                  autoFocus
                  placeholder="Paste di sini..."
                  className="w-full h-64 bg-slate-50 border-2 border-dashed border-slate-300 rounded-xl p-4 text-[10px] outline-none focus:border-slate-500 resize-none font-mono"
                  value={pasteData}
                  onChange={(e) => setPasteData(e.target.value)}
                ></textarea>
              </div>
              <div className="p-4 bg-slate-50 border-t flex justify-end gap-3">
                <button onClick={() => setShowPasteModal(false)} className="px-6 py-2 text-slate-500 font-bold text-[10px] uppercase">Batal</button>
                <button onClick={processImport} disabled={!pasteData.trim()} className="px-8 py-2 bg-slate-600 text-white font-black text-[10px] uppercase rounded-lg">Impor</button>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-[10px] text-left border-collapse min-w-[1200px]">
            <thead className="bg-slate-50 text-slate-500 uppercase text-[9px] font-black border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 w-8 border-r border-slate-200" rowSpan={2}>No</th>
                <th className="px-3 py-3 w-48 border-r border-slate-200" rowSpan={2}>"Risiko" yang Teridentifikasi</th>
                <th className="px-3 py-3 w-24 border-r border-slate-200" rowSpan={2}>Kode Risiko</th>
                <th className="px-3 py-3 border-r border-slate-200 text-center" colSpan={3}>Kejadian Risiko</th>
                <th className="px-3 py-3 w-32 border-r border-slate-200" rowSpan={2}>Keterangan</th>
                <th className="px-3 py-3 w-48 border-r border-slate-200" rowSpan={2}>RTP</th>
                <th className="px-3 py-3 w-32 border-r border-slate-200" rowSpan={2}>Rencana Pelaksanaan RTP</th>
                <th className="px-3 py-3 w-32 border-r border-slate-200" rowSpan={2}>Realisasi Pelaksanaan RTP</th>
                <th className="px-3 py-3 w-32" rowSpan={2}>Keterangan</th>
              </tr>
              <tr className="border-b border-slate-200">
                <th className="px-2 py-2 w-32 border-r border-slate-200 bg-slate-50/50 text-center">Tanggal terjadi</th>
                <th className="px-2 py-2 w-40 border-r border-slate-200 bg-slate-50/50 text-center">Sebab</th>
                <th className="px-2 py-2 w-40 border-r border-slate-200 bg-slate-50/50 text-center">Dampak</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {rows.map((row, idx) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors align-top">
                  <td className="px-3 py-4 font-bold text-slate-400 border-r border-slate-200 text-center">{idx + 1}</td>
                  <td className="px-3 py-4 border-r border-slate-200 font-medium text-slate-700 italic leading-relaxed">
                    <p className="mb-2">{row.risikoUraian || '(Belum diisi)'}</p>
                    {(() => {
                      const sRows = row.subRows && row.subRows.length > 0 ? row.subRows : [
                        { sebabUraian: row.sebabUraian || '', dampakUraian: row.dampakUraian || '' }
                      ];
                      
                      const hasContent = sRows.some((s: any) => s.sebabUraian?.trim() || s.dampakUraian?.trim());
                      if (!hasContent) return null;

                      return (
                        <div className="space-y-1 mt-1 not-italic">
                          {sRows.map((sub: any, sIdx: number) => (
                            <div key={sIdx} className="bg-slate-50 p-1 rounded text-[7px] border border-slate-100">
                              <p className="text-slate-400"><span className="font-bold opacity-70">Identified Sebab:</span> {sub.sebabUraian || '-'}</p>
                              <p className="text-red-600/50"><span className="font-bold opacity-70">Identified Dampak:</span> {sub.dampakUraian || '-'}</p>
                            </div>
                          ))}
                        </div>
                      );
                    })()}
                  </td>
                  <td className="px-3 py-4 border-r border-slate-200 font-black text-blue-600">
                    {row.risikoKode || '-'}
                  </td>
                  {/* Kejadian Risiko - Tanggal */}
                  <td className="px-2 py-4 border-r border-slate-200">
                    <EditableInput 
                      type="date"
                      className="w-full bg-transparent outline-none cursor-pointer font-bold text-slate-600 disabled:opacity-50"
                      value={row.eventDate || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'eventDate', val)}
                    />
                  </td>
                  {/* Kejadian Risiko - Sebab */}
                  <td className="px-2 py-4 border-r border-slate-200">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500"
                      rows={2}
                      placeholder="..."
                      value={row.eventCause || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'eventCause', val)}
                    />
                  </td>
                  {/* Kejadian Risiko - Dampak */}
                  <td className="px-2 py-4 border-r border-slate-200">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500"
                      rows={2}
                      placeholder="..."
                      value={row.eventImpact || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'eventImpact', val)}
                    />
                  </td>
                  {/* Keterangan (Kejadian) */}
                  <td className="px-2 py-4 border-r border-slate-200">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500"
                      rows={2}
                      placeholder="..."
                      value={row.eventNotes || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'eventNotes', val)}
                    />
                  </td>
                  {/* RTP */}
                  <td className="px-3 py-4 border-r border-slate-200 text-slate-600 italic">
                    {row.rtpAction || '(Belum diisi RTP)'}
                  </td>
                  {/* Rencana Pelaksanaan RTP */}
                  <td className="px-2 py-4 border-r border-slate-200">
                    <EditableInput 
                      type="date"
                      className="w-full bg-transparent outline-none cursor-pointer font-bold text-slate-600 disabled:opacity-50"
                      value={row.rtpPlanDate || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'rtpPlanDate', val)}
                    />
                  </td>
                  {/* Realisasi Pelaksanaan RTP */}
                  <td className="px-2 py-4 border-r border-slate-200">
                    <EditableInput 
                      type="date"
                      className="w-full bg-transparent outline-none cursor-pointer font-bold text-slate-600 disabled:opacity-50"
                      value={row.rtpRealDate || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'rtpRealDate', val)}
                    />
                  </td>
                  {/* Keterangan (RTP) */}
                  <td className="px-2 py-4">
                    <EditableTextarea 
                      className="w-full bg-transparent outline-none resize-none disabled:text-slate-500"
                      rows={2}
                      placeholder="..."
                      value={row.rtpNotesContent || ''}
                      disabled={isReadOnly}
                      onChange={(val) => updateField(row.id, 'rtpNotesContent', val)}
                    />
                  </td>
                </tr>
              ))}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={11} className="px-6 py-8 text-center text-slate-400 italic">Belum ada data risiko teridentifikasi.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function FinalDocumentView({ user, isAdmin, isOperator }: { user: any, isAdmin: boolean, isOperator: boolean }) {
  const [docLink, setDocLink] = useState('');
  const [masterLink, setMasterLink] = useState('');
  const [opdLinks, setOpdLinks] = useState<Record<string, { driveLink: string; uploadLink: string }>>({});
  const [allUsers, setAllUsers] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingOpdId, setEditingOpdId] = useState<string | null>(null);
  const [editValue, setEditValue] = useState('');
  const [finalDocs, setFinalDocs] = useState<any[]>([]);
  const [activeTab, setActiveTab] = useState<'verification' | 'config'>('verification');
  const isSuper = isAdmin || isOperator;

  useEffect(() => {
    // Listen to Master Storage Link (Global Fallback)
    const unsubConfig = onSnapshot(doc(db, 'system_config', 'final_doc_storage'), (snap) => {
      if (snap.exists()) {
        setMasterLink(snap.data().link || '');
      }
    });

    // Listen to Per-OPD Storage Links
    const unsubOpdLinks = onSnapshot(collection(db, 'system_config', 'storage_links', 'opds'), (snap) => {
      const mapping: Record<string, { driveLink: string; uploadLink: string }> = {};
      snap.forEach(d => {
        const data = d.data();
        mapping[d.id] = {
          driveLink: data.driveLink || data.link || '', // Fallback to old 'link' field if it exists
          uploadLink: data.uploadLink || ''
        };
      });
      setOpdLinks(mapping);
    });

    if (isSuper) {
      // Listen to all users (OPDs) to facilitate link mapping
      const unsubUsers = onSnapshot(collection(db, 'accounts'), (snap) => {
        setAllUsers(snap.docs.map(d => ({ ...d.data(), uid: d.id })));
      });

      const unsub = onSnapshot(collection(db, 'final_documents'), (snap) => {
        setFinalDocs(snap.docs.map(d => ({ ...d.data(), id: d.id })));
        setLoading(false);
      }, (error) => console.error('FinalDocs super error:', error));
      
      return () => { 
        unsub(); 
        unsubConfig(); 
        unsubOpdLinks(); 
        unsubUsers(); 
      };
    } else {
      const unsub = onSnapshot(doc(db, 'final_documents', user.uid), (snap) => {
        if (snap.exists()) {
          setFinalDocs([{ ...snap.data(), id: snap.id }]);
        } else {
          setFinalDocs([]);
        }
        setLoading(false);
      }, (error) => console.error('FinalDocs user error:', error));
      return () => { unsub(); unsubConfig(); unsubOpdLinks(); };
    }
  }, [user.uid, isSuper]);

  const handleUpdateOpdLink = async (targetUid: string, driveLink: string, uploadLink: string) => {
    setSaving(true);
    try {
      await setDoc(doc(db, 'system_config', 'storage_links', 'opds', targetUid), {
        driveLink,
        uploadLink,
        updatedAt: new Date().toISOString(),
        updatedBy: auth.currentUser?.uid || 'system'
      });
      setEditingOpdId(null);
      alert('Konfigurasi link OPD berhasil diperbarui');
    } catch (err) {
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  const currentOpdLinks = opdLinks[user.uid] || { driveLink: masterLink, uploadLink: '' };
  const uploadLinkToUse = currentOpdLinks.uploadLink || currentOpdLinks.driveLink;

  const handleUpdateMasterLink = async (link: string) => {
    if (!isAdmin) return;
    setSaving(true);
    try {
      await setDoc(doc(db, 'system_config', 'final_doc_storage'), {
        link,
        updatedAt: new Date().toISOString(),
        updatedBy: auth.currentUser?.uid || 'system'
      });
      setMasterLink(link);
      alert('Link folder penyimpanan berhasil diperbarui');
    } catch (err) {
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  const [hasConfirmedUpload, setHasConfirmedUpload] = useState(false);
  const [showDeclaration, setShowDeclaration] = useState(false);

  const handleConfirmUpload = async () => {
    if (!hasConfirmedUpload) {
      alert("Silakan centang pernyataan bahwa Anda telah mengunggah dokumen.");
      return;
    }
    
    setSaving(true);
    const targetUid = user.uid;
    try {
      await setDoc(doc(db, 'final_documents', targetUid), {
        status: 'pending',
        updatedAt: new Date().toISOString(),
        updatedBy: targetUid,
        username: user.username,
        submissionType: 'Google Drive Confirmation',
        note: null // Clear previous notes on re-submission
      }, { merge: true });
      setShowDeclaration(false);
      alert('Berhasil mengonfirmasi. Admin akan segera memverifikasi dokumen Anda di folder Drive.');
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, `final_documents/${targetUid}`);
    } finally {
      setSaving(false);
    }
  };

  const handleSubmit = async (targetUid: string = user.uid, link: string = docLink) => {
    // Legacy support for link-based submission (admins can still edit if needed, or I remove it)
    if (!link.trim()) return alert('Data tidak boleh kosong');
    setSaving(true);
    const docPath = `final_documents/${targetUid}`;
    try {
      await setDoc(doc(db, 'final_documents', targetUid), {
        link,
        status: isSuper ? 'verified' : 'pending',
        updatedAt: new Date().toISOString(),
        updatedBy: auth.currentUser?.uid || 'system',
        username: isSuper ? (finalDocs.find(d => d.id === targetUid)?.username || 'User') : user.username
      }, { merge: true });
      if (!isSuper) setDocLink(link);
      setEditingId(null);
      alert('Berhasil memperbarui data');
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, docPath);
    } finally {
      setSaving(false);
    }
  };

  const handleVerify = async (uid: string, status: 'verified' | 'rejected') => {
    let note = '';
    if (status === 'rejected') {
      note = prompt('Masukkan alasan penolakan (opsional):') || '';
    }

    setSaving(true);
    try {
      await setDoc(doc(db, 'final_documents', uid), {
        status,
        verifiedAt: new Date().toISOString(),
        verifiedBy: auth.currentUser?.uid || 'system',
        note: note || null
      }, { merge: true });
    } catch (err) {
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  if (loading) return <div className="flex justify-center p-12"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-slate-600"></div></div>;

  if (isSuper) {
    return (
      <div className="space-y-6">
        {/* Tab Navigation for Admin/Operator */}
        {isAdmin && (
          <div className="flex gap-2 p-1 bg-slate-100 rounded-2xl w-fit mb-6">
            <button 
              onClick={() => setActiveTab('verification')}
              className={`px-6 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest transition-all ${activeTab === 'verification' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
              Verifikasi Dokumen
            </button>
            <button 
              onClick={() => setActiveTab('config')}
              className={`px-6 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest transition-all ${activeTab === 'config' ? 'bg-white text-slate-800 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
              Konfigurasi Folder
            </button>
          </div>
        )}

        {(activeTab === 'config' && isAdmin) ? (
          <div className="space-y-6">
             <div className="bg-gradient-to-br from-slate-900 to-indigo-950 rounded-3xl p-8 shadow-xl border border-slate-800">
                <div className="flex items-center gap-4 mb-6">
                  <div className="p-3 bg-blue-500/20 text-blue-400 rounded-2xl border border-blue-500/30">
                    <Settings2 size={28} />
                  </div>
                  <div>
                    <h4 className="text-lg font-bold text-white uppercase tracking-tight italic">Pengaturan Link Folder per OPD</h4>
                    <p className="text-slate-400 text-xs">Masing-masing OPD akan diarahkan ke folder Google Drive spesifik yang Anda tentukan di sini.</p>
                  </div>
                </div>

                <div className="bg-white/5 rounded-2xl overflow-hidden border border-white/10">
                  <table className="w-full text-left">
                    <thead className="bg-white/5 border-b border-white/10">
                      <tr>
                        <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Nama OPD / Unit Kerja</th>
                        <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Konfigurasi Link</th>
                        <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/5">
                      {allUsers.filter(u => u.role !== 'admin' && u.role !== 'superadmin').map(u => {
                        const isEditing = editingOpdId === u.uid;
                        const currentData = opdLinks[u.uid] || { driveLink: '', uploadLink: '' };

                        return (
                          <tr key={u.uid} className="hover:bg-white/5 transition-colors">
                            <td className="px-6 py-4">
                              <p className="text-sm font-bold text-white uppercase">{u.username || u.name}</p>
                              <p className="text-[10px] text-slate-500 font-mono italic">{u.uid}</p>
                            </td>
                            <td className="px-6 py-4">
                              <div className="flex flex-col gap-3 min-w-[300px]">
                                {/* Upload Link / Google Form */}
                                <div className="space-y-1">
                                  <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest block">1. Link Form/Upload (Untuk User)</label>
                                  <div className="flex items-center gap-2">
                                    <input 
                                      type="text"
                                      className={`w-full bg-slate-800/50 border border-slate-700 rounded-lg px-3 py-1.5 text-white text-[10px] outline-none focus:ring-1 focus:ring-blue-500 transition-all ${!isEditing ? 'opacity-50 cursor-not-allowed border-transparent bg-transparent' : ''}`}
                                      placeholder={isEditing ? "https://docs.google.com/forms/..." : "Belum diatur"}
                                      defaultValue={typeof currentData === 'string' ? '' : currentData.uploadLink}
                                      id={`upload-link-${u.uid}`}
                                      disabled={!isEditing}
                                    />
                                  </div>
                                </div>
                                {/* Drive Link / Folder Storage */}
                                <div className="space-y-1">
                                  <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest block">2. Link Folder Drive (Untuk Admin)</label>
                                  <div className="flex items-center gap-2">
                                    <input 
                                      type="text"
                                      className={`w-full bg-slate-800/50 border border-slate-700 rounded-lg px-3 py-1.5 text-white text-[10px] outline-none focus:ring-1 focus:ring-blue-500 transition-all ${!isEditing ? 'opacity-50 cursor-not-allowed border-transparent bg-transparent' : ''}`}
                                      placeholder={isEditing ? "https://drive.google.com/..." : "Belum diatur"}
                                      defaultValue={typeof currentData === 'string' ? currentData : currentData.driveLink}
                                      id={`drive-link-${u.uid}`}
                                      disabled={!isEditing}
                                    />
                                    {!isEditing && (typeof currentData !== 'string' ? currentData.driveLink : currentData) && (
                                      <a href={ (typeof currentData !== 'string' ? currentData.driveLink : currentData).startsWith('http') ? (typeof currentData !== 'string' ? currentData.driveLink : currentData) : `https://${(typeof currentData !== 'string' ? currentData.driveLink : currentData)}`} target="_blank" rel="noreferrer" className="text-blue-400 hover:text-blue-300">
                                        <ExternalLink size={14} />
                                      </a>
                                    )}
                                  </div>
                                </div>
                              </div>
                            </td>
                            <td className="px-6 py-4 text-right">
                              <div className="flex justify-end gap-2">
                                {isEditing ? (
                                  <>
                                    <button 
                                      onClick={() => setEditingOpdId(null)}
                                      className="px-3 py-2 bg-slate-700 hover:bg-slate-600 text-white rounded-lg font-bold text-[10px] uppercase tracking-widest transition-all"
                                    >
                                      Batal
                                    </button>
                                    <button 
                                      onClick={() => {
                                        const uLink = (document.getElementById(`upload-link-${u.uid}`) as HTMLInputElement)?.value;
                                        const dLink = (document.getElementById(`drive-link-${u.uid}`) as HTMLInputElement)?.value;
                                        handleUpdateOpdLink(u.uid, dLink, uLink);
                                      }}
                                      disabled={saving}
                                      className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-bold text-[10px] uppercase tracking-widest transition-all shadow-lg shadow-emerald-200 disabled:opacity-50"
                                    >
                                      SIMPAN
                                    </button>
                                  </>
                                ) : (
                                  <button 
                                    onClick={() => setEditingOpdId(u.uid)}
                                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-emerald-400 rounded-lg font-bold text-[10px] uppercase tracking-widest transition-all border border-emerald-500/30 flex items-center gap-2"
                                  >
                                    <Edit size={12} /> {(typeof currentData === 'string' ? currentData : currentData.driveLink) ? 'EDIT LINK' : 'SET LINK'}
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
          </div>
        ) : (
          <div className="bg-white rounded-3xl border border-slate-200 overflow-hidden shadow-sm p-8">
            <div className="flex justify-between items-start mb-8">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-blue-100 text-blue-600 rounded-2xl">
                  <ShieldCheck size={32} />
                </div>
                <div>
                  <h3 className="text-xl font-bold text-slate-800 tracking-tight uppercase">Verifikasi Dokumen Final OPD</h3>
                  <p className="text-sm text-slate-500">Monitor dan validasi hasil unggahan dokumen final dari seluruh unit kerja.</p>
                </div>
              </div>
            </div>

            <div className="overflow-hidden border border-slate-100 rounded-2xl">
              <table className="w-full text-left">
              <thead className="bg-slate-50 border-b border-slate-100">
                <tr>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">OPD / Unit Kerja</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Link / Folder Drive</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-center">Status</th>
                  <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {finalDocs.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="px-6 py-12 text-center text-slate-400 italic text-sm">Belum ada dokumen yang diunggah</td>
                  </tr>
                ) : (
                  finalDocs.map(d => (
                    <tr key={d.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-5">
                        <p className="text-sm font-bold text-slate-700">{d.username || 'User ' + d.id.slice(0, 5)}</p>
                        <p className="text-[10px] text-slate-400 font-mono italic">{d.id}</p>
                      </td>
                      <td className="px-6 py-5">
                        <div className="flex flex-col gap-1">
                          {(() => {
                            const links = opdLinks[d.id] || { driveLink: masterLink };
                            const driveLink = links.driveLink;
                            return driveLink ? (
                              <div className="flex flex-col gap-1.5">
                                <div className="flex items-center gap-2">
                                  <img src="https://upload.wikimedia.org/wikipedia/commons/1/12/Google_Drive_icon_%282020%29.svg" className="w-3.5 h-3.5" alt="G-Drive" />
                                  <a href={driveLink.startsWith('http') ? driveLink : `https://${driveLink}`} target="_blank" rel="noreferrer" className="text-blue-600 hover:underline text-[10px] font-black uppercase tracking-widest">Buka Drive OPD</a>
                                </div>
                                {links.uploadLink && (
                                  <div className="flex items-center gap-2 opacity-60">
                                    <ExternalLink size={12} className="text-slate-400" />
                                    <a href={links.uploadLink.startsWith('http') ? links.uploadLink : `https://${links.uploadLink}`} target="_blank" rel="noreferrer" className="text-slate-600 hover:underline text-[9px] font-bold">Link Form User</a>
                                  </div>
                                )}
                              </div>
                            ) : <span className="text-slate-300 text-[10px] italic">Link Belum Diatur</span>
                          })()}
                          {d.updatedAt && <p className="text-[10px] text-slate-400 font-mono italic">Dikonfirmasi: {new Date(d.updatedAt).toLocaleString()}</p>}
                        </div>
                      </td>
                      <td className="px-6 py-5 text-center">
                        <span className={`px-2 py-1 rounded-md text-[10px] font-black uppercase tracking-tighter ${
                          d.status === 'verified' ? 'bg-emerald-100 text-emerald-700 border border-emerald-200' : 
                          d.status === 'rejected' ? 'bg-rose-100 text-rose-700 border border-rose-200' : 'bg-yellow-100 text-yellow-700 border border-yellow-200'
                        }`}>
                          {d.status === 'verified' ? 'TERVALIDASI' : d.status === 'rejected' ? 'DITOLAK' : 'MENUNGGU'}
                        </span>
                        {d.note && (
                          <p className="mt-1 text-[9px] text-rose-500 bg-rose-50 p-1 rounded border border-rose-100 font-medium max-w-[150px]">
                            Ket: {d.note}
                          </p>
                        )}
                      </td>
                      <td className="px-6 py-5 text-right space-x-2">
                        <div className="flex justify-end gap-2">
                          <button 
                            onClick={() => handleVerify(d.id, 'verified')}
                            disabled={d.status === 'verified' || saving}
                            className={`p-2 rounded-xl transition-all shadow-sm active:scale-95 ${d.status === 'verified' ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-emerald-100'}`}
                            title="Setujui/Verifikasi"
                          >
                            <Check size={18} strokeWidth={3} />
                          </button>
                          <button 
                            onClick={() => handleVerify(d.id, 'rejected')}
                            disabled={d.status === 'rejected' || saving}
                            className={`p-2 rounded-xl transition-all shadow-sm active:scale-95 ${d.status === 'rejected' ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-rose-600 text-white hover:bg-rose-700 shadow-rose-100'}`}
                            title="Tolak"
                          >
                            <X size={18} strokeWidth={3} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
          
          <div className="mt-8 p-4 bg-amber-50 rounded-2xl border border-amber-100 flex gap-4 items-start">
             <div className="p-2 bg-white text-amber-600 rounded-lg shadow-sm shrink-0">
               <ShieldCheck size={20} />
             </div>
             <div>
               <h5 className="font-bold text-amber-800 text-sm">Petunjuk {isAdmin ? 'Administrator' : 'Operator'}</h5>
               <p className="text-amber-700 text-[11px] leading-relaxed">
                 {isAdmin 
                   ? "Anda hanya memiliki akses untuk memvalidasi dokumen. Jika link dokumen salah, silakan minta Operator untuk melakukan pembaruan link." 
                   : "Anda dapat memvalidasi dokumen atau memperbarui link secara langsung jika terjadi kesalahan input oleh OPD."}
               </p>
             </div>
          </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className="bg-white rounded-3xl border border-slate-200 overflow-hidden shadow-sm">
        <div className="p-12 text-center space-y-6">
          <div className="w-20 h-20 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mx-auto border-4 border-blue-50">
            <ShieldCheck size={40} />
          </div>
          
          <div className="space-y-2">
            <h3 className="text-2xl font-bold text-slate-800 uppercase tracking-tight italic">Penyerahan Dokumen Final</h3>
            <p className="text-slate-500 max-w-lg mx-auto leading-relaxed text-sm">
              Pastikan Anda telah mengunggah semua dokumen laporan final yang telah ditandatangani ke folder Google Drive yang disediakan di bawah ini.
            </p>
          </div>

          <div className="max-w-2xl mx-auto pt-4">
            {currentOpdLinks.driveLink || currentOpdLinks.uploadLink ? (
              <div className="p-8 bg-slate-50 border border-slate-100 rounded-[32px] space-y-6 transition-all hover:shadow-xl hover:shadow-slate-200/50">
                <div className="space-y-4">
                  <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Alur Penyerahan Dokumen</span>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col items-center gap-3 text-center group hover:border-blue-300 transition-all">
                      <div className="w-8 h-8 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center font-black text-[10px] ring-4 ring-blue-50/50">1</div>
                      <img src="https://upload.wikimedia.org/wikipedia/commons/1/12/Google_Drive_icon_%282020%29.svg" className="w-10 h-10" alt="G-Drive" />
                      <div>
                        <p className="text-xs font-bold text-slate-700 uppercase tracking-tight">Upload Laporan</p>
                        <p className="text-[9px] text-slate-400 italic mt-1 leading-tight">Gunakan link berikut untuk mengunggah dokumen Anda</p>
                      </div>
                      <a 
                        href={uploadLinkToUse.startsWith('http') ? uploadLinkToUse : `https://${uploadLinkToUse}`} 
                        target="_blank" 
                        rel="noreferrer"
                        className="w-full group flex items-center justify-center gap-2 bg-slate-900 py-3 rounded-xl text-white font-black text-[10px] hover:bg-blue-600 transition-all uppercase tracking-widest shadow-lg shadow-slate-200"
                      >
                        UPLOAD DI SINI <ExternalLink size={12} className="group-hover:translate-x-0.5 group-hover:-translate-y-0.5 transition-transform" />
                      </a>
                    </div>

                    {(() => {
                      const docData = finalDocs.find(d => d.id === user.uid);
                      const status = docData?.status || 'none';
                      const isSubmitted = status === 'pending' || status === 'verified';

                      return (
                        <div className={`p-6 rounded-2xl border flex flex-col items-center gap-3 text-center transition-all ${saving || isSubmitted ? 'bg-slate-50 border-slate-100' : 'bg-white border-slate-200 hover:border-emerald-300 shadow-sm'}`}>
                          <div className={`w-8 h-8 rounded-full flex items-center justify-center font-black text-[10px] ring-4 ${isSubmitted ? 'bg-slate-100 text-slate-400 ring-slate-100/50' : 'bg-emerald-50 text-emerald-600 ring-emerald-50/50'}`}>2</div>
                          <div className={`w-10 h-10 rounded-full flex items-center justify-center ${isSubmitted ? 'bg-slate-100 text-slate-300' : 'bg-emerald-50 text-emerald-500'}`}>
                            {isSubmitted ? <ShieldCheck size={24} /> : <CheckCircle size={24} />}
                          </div>
                          <div>
                            <p className="text-xs font-bold text-slate-700 uppercase tracking-tight">{isSubmitted ? 'Telah Dikonfirmasi' : 'Konfirmasi Selesai'}</p>
                            <p className="text-[9px] text-slate-400 italic mt-1 leading-tight mb-4">
                              {status === 'verified' ? 'Dokumen sudah divalidasi oleh admin.' : 
                               status === 'pending' ? 'Menunggu admin memverifikasi folder Anda.' : 
                               'Admin akan mulai memverifikasi setelah Anda mengonfirmasi'}
                            </p>
                          </div>

                          {!isSubmitted && (
                            <button 
                              onClick={() => setShowDeclaration(true)}
                              disabled={saving || isSubmitted}
                              className={`w-full flex items-center justify-center gap-2 py-3 rounded-xl font-black text-[10px] transition-all uppercase tracking-widest shadow-lg disabled:cursor-not-allowed ${
                                isSubmitted ? 'bg-slate-200 text-slate-400 shadow-none' : 'bg-blue-600 text-white hover:bg-blue-700 shadow-blue-100'
                              }`}
                            >
                              KONFIRMASI SEKARANG
                            </button>
                          )}

                          {isSubmitted && (
                            <button 
                              disabled
                              className="w-full flex items-center justify-center gap-2 py-3 rounded-xl font-black text-[10px] transition-all uppercase tracking-widest bg-slate-200 text-slate-400"
                            >
                              SUDAH DIKONFIRMASI
                            </button>
                          )}
                        </div>
                      );
                    })()}
                  </div>
                </div>

                <AnimatePresence>
                  {showDeclaration && (
                    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
                      <motion.div 
                        initial={{ scale: 0.95, opacity: 0 }}
                        animate={{ scale: 1, opacity: 1 }}
                        exit={{ scale: 0.95, opacity: 0 }}
                        className="bg-white rounded-3xl shadow-2xl p-8 max-w-sm w-full text-center border border-slate-100"
                      >
                        <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mx-auto mb-6">
                          <CheckCircle size={32} className="text-blue-500" />
                        </div>
                        <h3 className="text-xl font-bold text-slate-900 mb-2 uppercase italic tracking-tight">Konfirmasi Final</h3>
                        <p className="text-slate-500 text-[11px] mb-6 font-medium italic leading-relaxed">
                          "Apakah Anda yakin data yang diupload sudah benar? Data tidak dapat diubah setelah konfirmasi."
                        </p>

                        <div className="flex items-start gap-3 text-left bg-emerald-50 p-4 rounded-2xl border border-emerald-100 mb-6 cursor-pointer hover:bg-emerald-100 transition-colors" onClick={() => setHasConfirmedUpload(!hasConfirmedUpload)}>
                          <div className={`mt-0.5 w-5 h-5 rounded border-2 flex items-center justify-center transition-colors ${hasConfirmedUpload ? 'bg-emerald-600 border-emerald-600' : 'bg-white border-emerald-300'}`}>
                            {hasConfirmedUpload && <Check size={14} className="text-white" />}
                          </div>
                          <span className="text-[10px] text-emerald-800 font-bold leading-tight uppercase tracking-tight">SAYA MENYATAKAN BAHWA SUDAH UPLOAD DOKUMEN FINAL KE FOLDER GOOGLE DRIVE DENGAN BENAR</span>
                        </div>

                        <div className="flex gap-3">
                          <button 
                            onClick={() => { setShowDeclaration(false); setHasConfirmedUpload(false); }}
                            className="flex-1 px-4 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold text-[11px] uppercase tracking-wider hover:bg-slate-200 transition-colors"
                          >
                            Batal
                          </button>
                          <button 
                            onClick={handleConfirmUpload}
                            disabled={saving || !hasConfirmedUpload}
                            className="flex-[1.5] px-4 py-3 bg-emerald-600 text-white rounded-xl font-black text-[11px] uppercase tracking-wider hover:bg-emerald-700 transition-colors shadow-lg shadow-emerald-100 disabled:opacity-50"
                          >
                            {saving ? 'MEMPROSES...' : 'YA, KONFIRMASI'}
                          </button>
                        </div>
                      </motion.div>
                    </div>
                  )}
                </AnimatePresence>

                {(() => {
                  const docData = finalDocs.find(d => d.id === user.uid);
                  if (!docData) return null; // Only show status section if they've confirmed at least once

                  const status = docData.status;

                  return (
                    <div className="pt-8 border-t border-slate-200/50 flex flex-col items-center gap-4 animate-in fade-in slide-in-from-bottom-4 duration-500">
                      <div className="flex items-center gap-1.5">
                         <div className="w-1.5 h-1.5 rounded-full bg-slate-300" />
                         <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Status Verifikasi</span>
                         <div className="w-1.5 h-1.5 rounded-full bg-slate-300" />
                      </div>
                      
                      <div className={`px-10 py-4 rounded-[24px] border-2 flex items-center gap-3 font-black uppercase tracking-widest text-sm shadow-xl ${
                        status === 'verified' ? 'bg-emerald-50 border-emerald-200 text-emerald-700 shadow-emerald-100/50' :
                        status === 'rejected' ? 'bg-rose-50 border-rose-200 text-rose-700 shadow-rose-100/50' :
                        'bg-blue-50 border-blue-200 text-blue-700 shadow-blue-100/50'
                      }`}>
                         {status === 'verified' ? <Check size={20} className="stroke-[3]" /> : status === 'rejected' ? <X size={20} className="stroke-[3]" /> : <div className="w-3 h-3 bg-blue-500 rounded-full animate-ping" />}
                         {status === 'verified' ? 'Sudah Terverifikasi' : status === 'rejected' ? 'Ditolak / Perlu Revisi' : 'Menunggu Verifikasi'}
                      </div>
                      
                      {docData.note && status === 'rejected' && (
                        <div className="p-4 bg-rose-50 border border-rose-100 text-rose-700 rounded-2xl text-[10px] font-black uppercase tracking-wider max-w-sm italic text-center leading-relaxed">
                          Catatan Admin: "{docData.note}"
                        </div>
                      )}

                      {status === 'pending' && (
                        <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest text-center animate-pulse">
                          Dokumen Anda sedang dalam antrean pemeriksaan
                        </p>
                      )}
                      
                      {status === 'verified' && (
                        <div className="flex flex-col items-center gap-1">
                          <p className="text-[10px] text-emerald-600 font-black uppercase tracking-widest">Proses Selesai</p>
                          <p className="text-[10px] text-slate-400 font-medium italic">Laporan final telah divalidasi oleh Verifikator.</p>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>
            ) : (
              <div className="p-12 bg-amber-50 border border-amber-100 rounded-[32px] text-center space-y-4">
                <div className="w-16 h-16 bg-white rounded-2xl shadow-sm flex items-center justify-center text-amber-500 mx-auto">
                  <ShieldAlert size={32} />
                </div>
                <div className="space-y-1">
                  <p className="font-bold text-amber-900">Folder Belum Tersedia</p>
                  <p className="text-xs text-amber-700 leading-relaxed max-w-xs mx-auto">
                    Admin belum mengatur folder Google Drive khusus untuk OPD Anda. Silakan hubungi Administrator untuk meminta link folder.
                  </p>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}


