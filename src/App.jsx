import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from 'recharts';
import { 
  Activity, Droplets, Wind, Thermometer, RotateCw, PlusCircle, 
  Download, LayoutDashboard, LineChart as ChartIcon, FolderPlus, Trash2, FolderOpen,
  Table as TableIcon, Users, Activity as ActivityIcon, Edit3,
  ChevronDown, ChevronUp, ChevronRight, Settings, LogOut, Cpu, Database, Folder,
  Menu, X, Check
} from 'lucide-react';
import './index.css';
import './form.css';
import './tabs.css';
import './table.css';

const getElapsedHours = (job, dataPointTimestamp) => {
  if (!job || !dataPointTimestamp) return 0;
  
  let startTimeMs = null;
  
  // Find the earliest timestamp among all data points in this job
  if (job.data && job.data.length > 0) {
    let minTimeMs = Infinity;
    job.data.forEach(row => {
      if (row.timestamp) {
        const t = new Date(row.timestamp).getTime();
        if (!isNaN(t) && t < minTimeMs) {
          minTimeMs = t;
        }
      }
    });
    if (minTimeMs !== Infinity) {
      startTimeMs = minTimeMs;
    }
  }
  
  // Fallback to Job ID or Job creation time if no data points have valid timestamps
  if (!startTimeMs) {
    if (job.id && typeof job.id === 'string' && job.id.startsWith('job-')) {
      const idNum = parseInt(job.id.replace('job-', ''), 10);
      if (!isNaN(idNum)) {
        startTimeMs = idNum;
      }
    }
    if (!startTimeMs && job.createdAt) {
      const d = new Date(job.createdAt);
      if (!isNaN(d.getTime())) {
        startTimeMs = d.getTime();
      }
    }
  }
  
  if (!startTimeMs) return 0;
  
  const recordTimeMs = new Date(dataPointTimestamp).getTime();
  if (isNaN(recordTimeMs)) return 0;
  
  const diffMs = recordTimeMs - startTimeMs;
  const hours = diffMs / (1000 * 60 * 60);
  return Math.max(0, parseFloat(hours.toFixed(1)));
};

const CustomTooltip = ({ active, payload, label }) => {
  if (active && payload && payload.length) {
    const rowData = payload[0].payload;
    const hrText = rowData && rowData.cultureHour !== undefined ? ` (ชั่วโมงที่ ${rowData.cultureHour})` : '';
    return (
      <div className="custom-tooltip">
        <p className="custom-tooltip-label">{label}{hrText}</p>
        {payload.map((entry, index) => (
          <p key={index} style={{ color: entry.color, margin: 0, fontSize: '14px', fontWeight: 600 }}>
            {entry.name}: {typeof entry.value === 'number' ? entry.value.toFixed(2) : entry.value || '-'}
          </p>
        ))}
      </div>
    );
  }
  return null;
};

// Initial defaults
const defaultMachine = { id: 'm1', name: 'Bioreactor 1' };
const defaultJob = {
  id: 'job-' + Date.now(),
  machineId: 'm1',
  name: 'Default Session',
  createdAt: new Date().toLocaleString(),
  data: []
};

function App() {
  // Helper: normalize time strings or Date to HH:MM for <input type="time">
  const pad2 = (n) => String(n).padStart(2, '0');
  const toHHMM = (v) => {
    if (!v) return '';
    if (v instanceof Date) return `${pad2(v.getHours())}:${pad2(v.getMinutes())}`;
    if (typeof v === 'string') {
      // try match HH:MM or HH:MM:SS
      const m = v.match(/(\d{1,2}):(\d{2})/);
      if (m) return `${pad2(parseInt(m[1], 10))}:${pad2(parseInt(m[2], 10))}`;
      // try Date.parse fallback
      const d = new Date(v);
      if (!isNaN(d.getTime())) return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
    }
    return '';
  };
  // Authentication State
  const [userRole, setUserRole] = useState(() => {
    return localStorage.getItem('bioprocess-role') || null; // 'admin' | 'customer' | null
  });
  const [activeCustomerJobId, setActiveCustomerJobId] = useState(() => {
    return localStorage.getItem('bioprocess-customer-job-id') || null;
  });

  // PWA Install Prompt State
  const [deferredPrompt, setDeferredPrompt] = useState(null);
  const [isStandalone, setIsStandalone] = useState(false);

  // Unique client ID for active users status tracking
  const [clientId] = useState(() => {
    let id = sessionStorage.getItem('bioprocess-client-id');
    if (!id) {
      id = 'client-' + Math.random().toString(36).substring(2, 11);
      try {
        sessionStorage.setItem('bioprocess-client-id', id);
      } catch (e) {
        // ignore storage errors
      }
    }
    return id;
  });

  const [activeUsers, setActiveUsers] = useState([]);

  useEffect(() => {
    const checkStandalone = window.navigator.standalone === true || 
                            window.matchMedia('(display-mode: standalone)').matches;
    setIsStandalone(checkStandalone);

    const handleBeforeInstallPrompt = (e) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallApp = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    console.log(`User response to install: ${outcome}`);
    setDeferredPrompt(null);
  };

  useEffect(() => {
    if (userRole) {
      localStorage.setItem('bioprocess-role', userRole);
    } else {
      localStorage.removeItem('bioprocess-role');
    }
  }, [userRole]);

  useEffect(() => {
    if (activeCustomerJobId) {
      localStorage.setItem('bioprocess-customer-job-id', activeCustomerJobId);
    } else {
      localStorage.removeItem('bioprocess-customer-job-id');
    }
  }, [activeCustomerJobId]);

  // Global View State
  const [currentAppView, setCurrentAppView] = useState('monitoring'); // 'monitoring' | 'customers'
  const [isInstrumentsExpanded, setIsInstrumentsExpanded] = useState(true);
  const [isSessionsExpanded, setIsSessionsExpanded] = useState(true);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  // Load States from Backend
  const [machines, setMachines] = useState([]);
  const [currentMachineId, setCurrentMachineId] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [currentJobId, setCurrentJobId] = useState(null);
  const [customers, setCustomers] = useState([]);

  const [activeTab, setActiveTab] = useState('dashboard'); // 'dashboard' | 'combined' | 'table'
  // Theme: 'dark' | 'light'
  const [theme, setTheme] = useState(() => {
    return localStorage.getItem('bioprocess-theme') || 'dark';
  });
  // Customer notice modal state
  const [showCustomerNotice, setShowCustomerNotice] = useState(false);
  const [pendingJobCode, setPendingJobCode] = useState(null);
  const [dontShowAgain, setDontShowAgain] = useState(() => {
    try { return localStorage.getItem('bioprocess-customer-notice-ack') === 'true'; } catch (e) { return false; }
  });
  const [showCustomerBanner, setShowCustomerBanner] = useState(false);
  // Customer Login Warning Toast State
  const [showCustomerToast, setShowCustomerToast] = useState(false);
  const [isToastHiding, setIsToastHiding] = useState(false);
  // Replay recorded data (playback) state
  const [isReplay, setIsReplay] = useState(false);
  const [replayIndex, setReplayIndex] = useState(0);
  const [replayVisibleData, setReplayVisibleData] = useState([]);

  const [formData, setFormData] = useState({
    temp_set: 38.0, temp_read: 38.0,
    ph_set: 7.00, ph_read: 7.00,
    do_set: 50, do_read: 50,
    agit_set: 200, agit_read: 200,
    air_set: 2.0, air_read: 2.0,
    remark: '',
    date: new Date().toLocaleDateString('en-CA'),
    time: toHHMM(new Date())
  });

  const [visibleParameters, setVisibleParameters] = useState({
    temp: true,
    ph: true,
    do: true,
    agit: true,
    air: true
  });

  const [customerFormData, setCustomerFormData] = useState({
    companyName: '',
    machineId: '',
    email: ''
  });

  const [selectedCustomerJobs, setSelectedCustomerJobs] = useState({});

  // Fetch Database from Backend Helper
  const fetchDB = async (shouldAutoSelect = false) => {
    try {
      const role = userRole || 'guest';
      const params = new URLSearchParams({
        clientId,
        role,
        machineId: currentMachineId || '',
        jobId: currentJobId || ''
      });
      const res = await fetch(`/api/db?${params.toString()}`);
      if (res.ok) {
        const data = await res.json();
        setMachines(data.machines);
        setJobs(data.jobs);
        setCustomers(data.customers);
        if (data.activeUsers) {
          setActiveUsers(data.activeUsers);
        }
        
        if (shouldAutoSelect) {
          const savedMachineId = localStorage.getItem('bioprocess-current-machine');
          const savedJobId = localStorage.getItem('bioprocess-current-job');
          
          const targetMachineId = savedMachineId && data.machines.some(m => m.id === savedMachineId) 
            ? savedMachineId 
            : (data.machines[0]?.id || null);
            
          const targetJobId = savedJobId && data.jobs.some(j => j.id === savedJobId)
            ? savedJobId
            : (data.jobs.filter(j => j.machineId === targetMachineId)[0]?.id || null);
            
          setCurrentMachineId(targetMachineId);
          setCurrentJobId(targetJobId);
        }
      }
    } catch (e) {
      console.error("Error fetching database:", e);
    }
  };

  const handleAutoCustomerLogin = async (jobCode) => {
    try {
      const role = userRole || 'guest';
      const params = new URLSearchParams({
        clientId,
        role,
        machineId: currentMachineId || '',
        jobId: jobCode
      });
      const res = await fetch(`/api/db?${params.toString()}`);
      if (res.ok) {
        const data = await res.json();
        const targetJob = data.jobs.find(j => j.id === jobCode);
        if (!targetJob) {
          alert('ไม่พบรหัสงานนี้ในระบบ หรือ ลิงก์ไม่ถูกต้อง');
          return;
        }

        // Expiry check
        if (targetJob.expiresAt && new Date() > new Date(targetJob.expiresAt)) {
          alert('สิทธิ์การเข้าใช้งานเซสชันนี้หมดอายุแล้ว');
          return;
        }

        // Apply fresh DB
        setMachines(data.machines);
        setJobs(data.jobs);
        setCustomers(data.customers);
        if (data.activeUsers) {
          setActiveUsers(data.activeUsers);
        }

        const ack = localStorage.getItem('bioprocess-customer-notice-ack') === 'true';
        if (ack) {
          setActiveCustomerJobId(jobCode);
          setUserRole('customer');
          setCurrentAppView('monitoring');
        } else {
          setPendingJobCode(jobCode);
          setShowCustomerNotice(true);
        }
      }
    } catch (err) {
      console.error("Error auto logging in customer:", err);
    }
  };

  useEffect(() => {
    fetchDB(true);

    const params = new URLSearchParams(window.location.search);
    const urlJobId = params.get('job');
    if (urlJobId) {
      window.history.replaceState({}, document.title, window.location.pathname);
      handleAutoCustomerLogin(urlJobId);
    }
  }, []);

  // Poll server for real-time synchronization across multiple users
  useEffect(() => {
    const interval = setInterval(() => {
      fetchDB(false);
    }, 5000);
    return () => clearInterval(interval);
  }, [currentMachineId, currentJobId, userRole]);

  // Active Customer Session Expiry Check
  useEffect(() => {
    if (userRole === 'customer' && activeCustomerJobId && jobs.length > 0) {
      const activeJob = jobs.find(j => j.id === activeCustomerJobId);
      if (activeJob && activeJob.expiresAt && new Date() > new Date(activeJob.expiresAt)) {
        alert("ระยะเวลาการเข้าใช้งานเซสชันนี้หมดอายุแล้ว ระบบจะนำคุณออกจากระบบโดยอัตโนมัติ");
        setUserRole(null);
        setActiveCustomerJobId(null);
        setCurrentAppView('monitoring');
      }
    }
  }, [jobs, userRole, activeCustomerJobId]);

  // Save current active selections locally in browser
  useEffect(() => {
    if (currentMachineId) {
      localStorage.setItem('bioprocess-current-machine', currentMachineId);
    }
  }, [currentMachineId]);

  useEffect(() => {
    if (currentJobId) {
      localStorage.setItem('bioprocess-current-job', currentJobId);
    }
  }, [currentJobId]);

  // Apply theme to document and persist
  useEffect(() => {
    try {
      document.documentElement.setAttribute('data-theme', theme);
      localStorage.setItem('bioprocess-theme', theme);
    } catch (e) {
      // server-side or non-browser env ignore
    }
  }, [theme]);

  useEffect(() => {
    try { localStorage.setItem('bioprocess-customer-notice-ack', dontShowAgain ? 'true' : 'false'); } catch (e) {}
  }, [dontShowAgain]);

  // Ensure customer form has a valid machineId selected
  useEffect(() => {
    if (!customerFormData.machineId && machines.length > 0) {
      setCustomerFormData(prev => ({ ...prev, machineId: machines[0].id }));
    }
  }, [machines, customerFormData.machineId]);

  // Derived state
  const jobsForMachine = jobs.filter(j => j.machineId === currentMachineId);
  const currentMachine = machines.find(m => m.id === currentMachineId);
  const currentJob = jobs.find(j => j.id === currentJobId) || jobsForMachine[0];
  const currentJobData = currentJob?.data || [];
  // Choose which data set to render: live/full stored data or replayed slice
  const displayData = isReplay ? replayVisibleData : currentJobData;
  const chartData = displayData.map((row, idx) => ({
    ...row,
    originalIndex: idx,
    cultureHour: getElapsedHours(currentJob, row.timestamp),
    temp_read: row.temp_read !== undefined ? row.temp_read : row.temp,
    temp_set: row.temp_set !== undefined ? row.temp_set : row.temp,
    ph_read: row.ph_read !== undefined ? row.ph_read : row.ph,
    ph_set: row.ph_set !== undefined ? row.ph_set : row.ph,
    do_read: row.do_read !== undefined ? row.do_read : row.do,
    do_set: row.do_set !== undefined ? row.do_set : row.do,
    agit_read: row.agit_read !== undefined ? row.agit_read : row.agit,
    agit_set: row.agit_set !== undefined ? row.agit_set : row.agit,
    air_read: row.air_read !== undefined ? row.air_read : row.air,
    air_set: row.air_set !== undefined ? row.air_set : row.air,
    remark: row.remark !== undefined ? row.remark : ''
  })).sort((a, b) => {
    const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
    const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
    return aTime - bTime;
  });
  // Sorting state for table
  const [sortField, setSortField] = useState('timestamp'); // default sort by timestamp
  const [sortAsc, setSortAsc] = useState(false);

  const toggleSort = (field) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const getSortedRows = () => {
    const rows = [...chartData];
    rows.sort((a, b) => {
      if (sortField === 'date' || sortField === 'time' || sortField === 'timestamp' || sortField === 'cultureHour') {
        const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
        const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
        return sortAsc ? aTime - bTime : bTime - aTime;
      }
      const aVal = a[sortField] !== undefined ? a[sortField] : '';
      const bVal = b[sortField] !== undefined ? b[sortField] : '';
      if (aVal < bVal) return sortAsc ? -1 : 1;
      if (aVal > bVal) return sortAsc ? 1 : -1;
      return 0;
    });
    return rows;
  };
  const getEditingRowCultureHour = () => {
    if (!editingRowData) return 0;
    const { date, time } = editingRowData;
    if (date && time) {
      const dt = new Date(`${date}T${time}`);
      if (!isNaN(dt.getTime())) {
        return getElapsedHours(currentJob, dt.toISOString());
      }
      const dt2 = new Date(`${date} ${time}`);
      if (!isNaN(dt2.getTime())) {
        return getElapsedHours(currentJob, dt2.toISOString());
      }
    }
    return getElapsedHours(currentJob, editingRowData.timestamp);
  };
  const lastDataPoint = chartData.length > 0 ? chartData[chartData.length - 1] : null;

  // Enforce customer constraints: lock into active customer job and machine
  useEffect(() => {
    if (userRole === 'customer' && activeCustomerJobId) {
      setCurrentJobId(activeCustomerJobId);
      const targetJob = jobs.find(j => j.id === activeCustomerJobId);
      if (targetJob && targetJob.machineId) {
        setCurrentMachineId(targetJob.machineId);
      }
    }
  }, [userRole, activeCustomerJobId, jobs]);

  // Show transient banner for customers on login (unless they've chosen don't-show-again)
  useEffect(() => {
    if (userRole === 'customer') {
      try {
        const ack = localStorage.getItem('bioprocess-customer-notice-ack') === 'true';
        if (!ack) {
          setShowCustomerBanner(true);
          const t = setTimeout(() => setShowCustomerBanner(false), 6000);
          return () => clearTimeout(t);
        }
      } catch (e) {
        // ignore
      }
    }
    return undefined;
  }, [userRole, activeCustomerJobId]);

  // Customer Warning Toast Effect
  useEffect(() => {
    if (userRole === 'customer') {
      setShowCustomerToast(true);
      setIsToastHiding(false);
      
      const hideTimer = setTimeout(() => {
        setIsToastHiding(true);
      }, 14400);

      const closeTimer = setTimeout(() => {
        setShowCustomerToast(false);
        setIsToastHiding(false);
      }, 15000);

      return () => {
        clearTimeout(hideTimer);
        clearTimeout(closeTimer);
      };
    } else {
      setShowCustomerToast(false);
      setIsToastHiding(false);
    }
  }, [userRole]);

  const dismissCustomerToast = () => {
    setIsToastHiding(true);
    setTimeout(() => {
      setShowCustomerToast(false);
      setIsToastHiding(false);
    }, 600);
  };

  const handleCopyShareLink = async (job) => {
    let currentExpiryText = "ไม่มีวันหมดอายุ";
    if (job.expiresAt) {
      const expDate = new Date(job.expiresAt);
      if (expDate > new Date()) {
        currentExpiryText = `หมดอายุวันที่ ${expDate.toLocaleString('th-TH')}`;
      } else {
        currentExpiryText = `หมดอายุแล้วเมื่อ ${expDate.toLocaleString('th-TH')}`;
      }
    }

    const inputHours = prompt(
      `ระบุระยะเวลาที่ลูกค้าสามารถเข้าใช้งานลิงก์นี้ได้ (เป็นชั่วโมง)\n\n• ป้อนตัวเลข เช่น 24 (สำหรับ 1 วัน), 168 (สำหรับ 1 สัปดาห์)\n• เว้นว่าง หรือ ป้อน 0 เพื่อไม่จำกัดระยะเวลาการใช้งาน\n\nสถานะหมดอายุปัจจุบัน: ${currentExpiryText}`,
      ""
    );

    if (inputHours === null) return; // User clicked Cancel

    let expiresAt = null;
    const hours = parseFloat(inputHours);

    if (!isNaN(hours) && hours > 0) {
      expiresAt = new Date(Date.now() + hours * 60 * 60 * 1000).toISOString();
    }

    try {
      const res = await fetch(`/api/jobs/${job.id}/expiry`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ expiresAt })
      });

      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);

        // Copy direct link to clipboard
        const loginUrl = `${window.location.origin}/?job=${job.id}`;
        await navigator.clipboard.writeText(loginUrl);

        if (expiresAt) {
          const expDate = new Date(expiresAt);
          alert(`คัดลอกลิงก์สำหรับลูกค้าเข้าใช้งานเรียบร้อยแล้ว!\n\nลิงก์: ${loginUrl}\n\n⚠️ ลิงก์นี้จะหมดอายุวันที่: ${expDate.toLocaleString('th-TH')}`);
        } else {
          alert(`คัดลอกลิงก์สำหรับลูกค้าเข้าใช้งานเรียบร้อยแล้ว!\n\nลิงก์: ${loginUrl}\n\n(ลิงก์นี้ไม่มีวันหมดอายุ)`);
        }
      } else {
        alert("ไม่สามารถกำหนดเวลาหมดอายุได้ กรุณาลองใหม่อีกครั้ง");
      }
    } catch (err) {
      console.error("Error setting session expiry:", err);
      alert("เกิดข้อผิดพลาดในการตั้งค่าเวลาหมดอายุ");
    }
  };

  // Reset replay when job changes or jobs list updates
  useEffect(() => {
    if (!currentJobId) {
      setIsReplay(false);
      setReplayIndex(0);
      setReplayVisibleData([]);
    } else {
      // If replay active but job changed, restart replay
      if (isReplay) {
        setReplayIndex(0);
        setReplayVisibleData([]);
      }
    }
  }, [currentJobId, jobs]);

  useEffect(() => {
    if (currentJob && currentJob.machineId !== currentMachineId) {
      const firstJobForMachine = jobsForMachine[0];
      setCurrentJobId(firstJobForMachine ? firstJobForMachine.id : null);
    }
  }, [currentMachineId, currentJob, jobsForMachine]);

  // Pre-populate manual entry form with the last data point of the active session
  useEffect(() => {
    // determine last timestamp -> date and time fields
    const lastTs = lastDataPoint && (lastDataPoint.timestamp || lastDataPoint.time) ? (lastDataPoint.timestamp || lastDataPoint.time) : null;
    const lastDate = lastTs ? new Date(lastTs) : null;
    const lastDateValid = lastDate && !isNaN(lastDate.getTime());
    if (lastDataPoint) {
      setFormData({
        temp_set: lastDataPoint.temp_set,
        temp_read: lastDataPoint.temp_read,
        ph_set: lastDataPoint.ph_set,
        ph_read: lastDataPoint.ph_read,
        do_set: lastDataPoint.do_set,
        do_read: lastDataPoint.do_read,
        agit_set: lastDataPoint.agit_set,
        agit_read: lastDataPoint.agit_read,
        air_set: lastDataPoint.air_set,
        air_read: lastDataPoint.air_read,
        remark: '',
        date: new Date().toLocaleDateString('en-CA'),
        time: toHHMM(new Date())
      });
    } else {
      setFormData({
        temp_set: 37.0,
        temp_read: 37.0,
        ph_set: 7.00,
        ph_read: 7.00,
        do_set: 50,
        do_read: 50,
        agit_set: 200,
        agit_read: 200,
        air_set: 2.0,
        air_read: 2.0,
        remark: '',
        date: new Date().toISOString().slice(0,10),
        time: toHHMM(new Date())
      });
    }
  }, [currentJobId, lastDataPoint === null]);

  // Helper to apply updated DB state
  const applyDBUpdate = (data) => {
    setMachines(data.machines);
    setJobs(data.jobs);
    setCustomers(data.customers);
  };

  // Actions
  const handleMachineChange = async (e) => {
    const value = e.target.value;
    if (value === 'ADD_NEW') {
      const name = prompt("Enter a name for the new Machine/Instrument:");
      if (name && name.trim()) {
        try {
          const res = await fetch('/api/machines', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: name.trim() })
          });
          if (res.ok) {
            const data = await res.json();
            applyDBUpdate(data);
            
            const newMachine = data.machines[data.machines.length - 1];
            if (newMachine) {
              setCurrentMachineId(newMachine.id);
              const jobRes = await fetch('/api/jobs', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ machineId: newMachine.id, name: 'Session 1' })
              });
              if (jobRes.ok) {
                const jobData = await jobRes.json();
                applyDBUpdate(jobData);
                const newJob = jobData.jobs.find(j => j.machineId === newMachine.id);
                if (newJob) {
                  setCurrentJobId(newJob.id);
                }
              }
            }
          }
        } catch (err) {
          console.error(err);
        }
      }
    } else {
      setCurrentMachineId(value);
    }
  };

  const renameMachine = async () => {
    const activeMachine = machines.find(m => m.id === currentMachineId);
    if (!activeMachine) return;
    const newName = prompt("Enter new name for the machine:", activeMachine.name);
    if (newName && newName.trim() && newName.trim() !== activeMachine.name) {
      try {
        const res = await fetch(`/api/machines/${currentMachineId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: newName.trim() })
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const deleteMachine = async () => {
    const activeMachine = machines.find(m => m.id === currentMachineId);
    if (!activeMachine) return;
    
    if (machines.length <= 1) {
      alert("You must keep at least one machine.");
      return;
    }

    if (window.confirm(`Are you sure you want to delete "${activeMachine.name}"? All its sessions and data will be permanently deleted.`)) {
      try {
        const res = await fetch(`/api/machines/${currentMachineId}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
          
          const nextMachineId = data.machines[0]?.id || null;
          setCurrentMachineId(nextMachineId);
          
          const nextMachineJobs = data.jobs.filter(j => j.machineId === nextMachineId);
          setCurrentJobId(nextMachineJobs.length > 0 ? nextMachineJobs[0].id : null);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const createNewJob = async () => {
    if (!currentMachineId) return alert("Please select or create a machine first.");
    const jobName = prompt("Enter a name for the new session:", `Session ${jobsForMachine.length + 1}`);
    if (jobName && jobName.trim()) {
      try {
        const res = await fetch('/api/jobs', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ machineId: currentMachineId, name: jobName.trim() })
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
          
          const newJob = data.jobs.find(j => j.machineId === currentMachineId && j.name === jobName.trim());
          if (newJob) {
            setCurrentJobId(newJob.id);
          }
          // Auto mode removed
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const deleteJob = async (id, e) => {
    e.stopPropagation();
    if (window.confirm("Are you sure you want to delete this session? All its data will be lost.")) {
      try {
        const res = await fetch(`/api/jobs/${id}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
          if (currentJobId === id) {
            const remainingForMachine = data.jobs.filter(j => j.machineId === currentMachineId);
            setCurrentJobId(remainingForMachine.length > 0 ? remainingForMachine[0].id : null);
          }
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  // Replay (playback) logic: append stored points to visible slice
  useEffect(() => {
    if (!isReplay) return undefined;
    if (!currentJobData || currentJobData.length === 0) {
      setIsReplay(false);
      return undefined;
    }

    let interval = null;
    interval = setInterval(() => {
      setReplayIndex((ri) => {
        const next = ri + 1;
        if (currentJobData[ri]) {
          setReplayVisibleData(prev => [...prev, currentJobData[ri]]);
        }
        if (next >= currentJobData.length) {
          // Finish replay
          clearInterval(interval);
          setIsReplay(false);
          return currentJobData.length; // set index beyond end
        }
        return next;
      });
    }, 2000);

    return () => clearInterval(interval);
  }, [isReplay, currentJobData]);

  const deleteDataPoint = async (index) => {
    if (window.confirm("Delete this recorded value?")) {
      try {
        const res = await fetch(`/api/jobs/${currentJobId}/data/${index}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  // Row Editing States
  const [editingRowIndex, setEditingRowIndex] = useState(null);
  const [editingRowData, setEditingRowData] = useState(null);

  const startEditRow = (index, row) => {
    setEditingRowIndex(index);
    setEditingRowData({
      timestamp: row.timestamp || new Date().toISOString(),
      date: row.date || (row.timestamp ? new Date(row.timestamp).toISOString().slice(0, 10) : ''),
      time: row.time || (row.timestamp ? new Date(row.timestamp).toTimeString().slice(0, 5) : ''),
      temp_set: row.temp_set !== undefined ? row.temp_set : 0,
      temp_read: row.temp_read !== undefined ? row.temp_read : 0,
      ph_set: row.ph_set !== undefined ? row.ph_set : 0,
      ph_read: row.ph_read !== undefined ? row.ph_read : 0,
      do_set: row.do_set !== undefined ? row.do_set : 0,
      do_read: row.do_read !== undefined ? row.do_read : 0,
      agit_set: row.agit_set !== undefined ? row.agit_set : 0,
      agit_read: row.agit_read !== undefined ? row.agit_read : 0,
      air_set: row.air_set !== undefined ? row.air_set : 0,
      air_read: row.air_read !== undefined ? row.air_read : 0,
      remark: row.remark !== undefined ? row.remark : ''
    });
  };

  const handleEditChange = (field, value) => {
    setEditingRowData(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const cancelEditRow = () => {
    setEditingRowIndex(null);
    setEditingRowData(null);
  };

  const saveEditRow = async () => {
    if (editingRowIndex === null || !editingRowData) return;
    try {
      const { date, time } = editingRowData;
      let timestamp = editingRowData.timestamp || new Date().toISOString();
      if (date && time) {
        const dateParts = date.split('-');
        const timeParts = time.split(':');
        if (dateParts.length === 3 && timeParts.length >= 2) {
          const year = parseInt(dateParts[0], 10);
          const month = parseInt(dateParts[1], 10) - 1;
          const day = parseInt(dateParts[2], 10);
          const hours = parseInt(timeParts[0], 10);
          const minutes = parseInt(timeParts[1], 10);
          const localDate = new Date(year, month, day, hours, minutes, 0);
          if (!isNaN(localDate.getTime())) {
            timestamp = localDate.toISOString();
          }
        }
      }

      const payload = {
        ...editingRowData,
        timestamp
      };

      const res = await fetch(`/api/jobs/${currentJobId}/data/${editingRowIndex}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        cancelEditRow();
      } else {
        alert("Failed to save changes. Please try again.");
      }
    } catch (err) {
      console.error("Error saving edited row:", err);
      alert("An error occurred while saving.");
    }
  };

  const clearAllData = async () => {
    if (window.confirm("Are you sure you want to delete ALL data in this session?")) {
      try {
        const res = await fetch(`/api/jobs/${currentJobId}/clear`, {
          method: 'POST'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const handleInputChange = (e) => {
    const { name, value, type } = e.target;
    setFormData(prev => ({ 
      ...prev, 
      [name]: type === 'number' ? (parseFloat(value) || 0) : value 
    }));
  };

  const handleManualSubmit = (e) => {
    e.preventDefault();
    const now = new Date();
    const inputDate = formData.date || now.toLocaleDateString('en-CA');
    const inputTime = formData.time || toHHMM(now);
    addDataPoint(
      formData.temp_set, formData.temp_read,
      formData.ph_set, formData.ph_read,
      formData.do_set, formData.do_read,
      formData.agit_set, formData.agit_read,
      formData.air_set, formData.air_read,
      formData.remark,
      inputTime,
      inputDate
    );
    // Reset remark and update date/time to now
    const resetNow = new Date();
    setFormData(prev => ({ 
      ...prev, 
      remark: '',
      date: resetNow.toLocaleDateString('en-CA'),
      time: toHHMM(resetNow)
    }));
  };

  const addDataPoint = async (
    temp_set, temp_read,
    ph_set, ph_read,
    do_set, do_read,
    agit_set, agit_read,
    air_set, air_read,
    remark = '',
    time = '',
    date = ''
  ) => {
    if (!currentJobId) return;
    
    // Construct local timestamp on client side to avoid server-side timezone shifts
    let timestamp = new Date().toISOString();
    if (date && time) {
      const dateParts = date.split('-');
      const timeParts = time.split(':');
      if (dateParts.length === 3 && timeParts.length >= 2) {
        const year = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1;
        const day = parseInt(dateParts[2], 10);
        const hours = parseInt(timeParts[0], 10);
        const minutes = parseInt(timeParts[1], 10);
        const localDate = new Date(year, month, day, hours, minutes, 0);
        if (!isNaN(localDate.getTime())) {
          timestamp = localDate.toISOString();
        }
      }
    }

    try {
      const res = await fetch(`/api/jobs/${currentJobId}/data`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          temp_set, temp_read,
          ph_set, ph_read,
          do_set, do_read,
          agit_set, agit_read,
          air_set, air_read,
          remark,
          time,
          date,
          timestamp
        })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
      }
    } catch (err) {
      console.error(err);
    }
  };

  // Customer Management Actions
  const handleCustomerInputChange = (e) => {
    const { name, value } = e.target;
    setCustomerFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleAddCustomer = async (e) => {
    e.preventDefault();
    if (!customerFormData.companyName.trim()) return alert("Please enter a company name.");
    
    try {
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          companyName: customerFormData.companyName.trim(),
          machineId: customerFormData.machineId,
          email: customerFormData.email.trim()
        })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        setCustomerFormData(prev => ({ ...prev, companyName: '', email: '' }));
      }
    } catch (err) {
      console.error(err);
    }
  };

  const deleteCustomer = async (id) => {
    if (window.confirm("Are you sure you want to delete this customer?")) {
      try {
        const res = await fetch(`/api/customers/${id}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const getMachineName = (id) => {
    const m = machines.find(m => m.id === id);
    return m ? m.name : 'Unknown Machine';
  };

  // Export
  const exportToCSV = () => {
    if (!currentJob || currentJob.data.length === 0) {
      alert("No data to export!");
      return;
    }
    const headers = [
      'Date',
      'Time',
      'Culture Hour (Hr)',
      'Temp SV (C)', 'Temp PV (C)', 
      'pH SV', 'pH PV', 
      'DO SV (%)', 'DO PV (%)', 
      'Agit SV (RPM)', 'Agit PV (RPM)', 
      'Air Flow SV (L/M)', 'Air Flow PV (L/M)',
      'Remarks'
    ];
    const csvRows = [headers.join(',')];
    const rowsToExport = getSortedRows().length > 0 ? getSortedRows() : currentJob.data;
    rowsToExport.forEach(row => {
      const dateVal = row.date || (row.timestamp ? (new Date(row.timestamp).toISOString().slice(0, 10)) : '');
      const timeVal = row.time || (row.timestamp ? (new Date(row.timestamp).toTimeString().slice(0, 5)) : '');
      const hrVal = row.cultureHour !== undefined ? row.cultureHour : getElapsedHours(currentJob, row.timestamp);
      const t_s = row.temp_set !== undefined ? row.temp_set : row.temp;
      const t_r = row.temp_read !== undefined ? row.temp_read : row.temp;
      const p_s = row.ph_set !== undefined ? row.ph_set : row.ph;
      const p_r = row.ph_read !== undefined ? row.ph_read : row.ph;
      const d_s = row.do_set !== undefined ? row.do_set : row.do;
      const d_r = row.do_read !== undefined ? row.do_read : row.do;
      const ag_s = row.agit_set !== undefined ? row.agit_set : row.agit;
      const ag_r = row.agit_read !== undefined ? row.agit_read : row.agit;
      const ai_s = row.air_set !== undefined ? row.air_set : row.air;
      const ai_r = row.air_read !== undefined ? row.air_read : row.air;
      const rem = row.remark !== undefined ? row.remark : '';

      const values = [
        `"${dateVal}"`, 
        `"${timeVal}"`, 
        hrVal,
        t_s, t_r, 
        p_s, p_r, 
        d_s, d_r, 
        ag_s, ag_r, 
        ai_s, ai_r,
        `"${rem.replace(/"/g, '""')}"`
      ];
      csvRows.push(values.join(','));
    });
    const csvContent = csvRows.join('\n');
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const safeName = currentJob.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    link.setAttribute('download', `${safeName}_data.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const exportToExcel = () => {
    if (!currentJob || currentJob.data.length === 0) {
      alert("No data to export!");
      return;
    }

    const rowsToExport = getSortedRows().length > 0 ? getSortedRows() : currentJob.data;
    
    const sheetData = rowsToExport.map(row => {
      return {
        'Date': row.date || (row.timestamp ? (new Date(row.timestamp).toISOString().slice(0, 10)) : ''),
        'Time': row.time || (row.timestamp ? (new Date(row.timestamp).toTimeString().slice(0, 5)) : ''),
        'Culture Hour (Hr)': row.cultureHour !== undefined ? row.cultureHour : getElapsedHours(currentJob, row.timestamp),
        'Temp SV (°C)': row.temp_set !== undefined ? row.temp_set : row.temp,
        'Temp PV (°C)': row.temp_read !== undefined ? row.temp_read : row.temp,
        'pH SV': row.ph_set !== undefined ? row.ph_set : row.ph,
        'pH PV': row.ph_read !== undefined ? row.ph_read : row.ph,
        'DO SV (%)': row.do_set !== undefined ? row.do_set : row.do,
        'DO PV (%)': row.do_read !== undefined ? row.do_read : row.do,
        'Agit SV (RPM)': row.agit_set !== undefined ? row.agit_set : row.agit,
        'Agit PV (RPM)': row.agit_read !== undefined ? row.agit_read : row.agit,
        'Air Flow SV (L/M)': row.air_set !== undefined ? row.air_set : row.air,
        'Air Flow PV (L/M)': row.air_read !== undefined ? row.air_read : row.air,
        'Remarks': row.remark !== undefined ? row.remark : ''
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Bioprocess Records");
    
    // Auto-fit column widths
    const max_widths = [];
    sheetData.forEach(row => {
      Object.keys(row).forEach((key, colIndex) => {
        const value = row[key] ? String(row[key]) : "";
        const length = Math.max(value.length, key.length) + 2;
        max_widths[colIndex] = Math.max(max_widths[colIndex] || 0, length);
      });
    });
    worksheet['!cols'] = max_widths.map(w => ({ wch: w }));

    const safeName = currentJob.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    XLSX.writeFile(workbook, `${safeName}_data.xlsx`);
  };

  const exportToPDF = () => {
    if (!currentJob) return;
    window.print();
  };

  // Auto Simulation removed

  if (!userRole) {
    return (
      <div className="login-overlay">
        <div className="glass-panel login-card" style={{ maxWidth: '450px', width: '95%', padding: '2.5rem', margin: 'auto' }}>
          <h1 style={{ textAlign: 'center', marginBottom: '0.5rem', background: 'linear-gradient(to right, var(--accent-blue), var(--accent-purple))', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', fontWeight: 700 }}>DBMS</h1>
          <p style={{ textAlign: 'center', color: 'var(--text-secondary)', fontSize: '0.9rem', marginBottom: '2rem' }}>เครื่องมือบันทึกและจัดเก็บข้อมูล</p>


          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            {/* Admin Login Block */}
            <div style={{ padding: '1.25rem', border: '1px solid var(--border-color)', borderRadius: '12px', background: 'rgba(15, 23, 42, 0.3)' }}>
              <h3 style={{ marginBottom: '1rem', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '1.05rem' }}>🔐 สำหรับแอดมิน (Admin)</h3>
              <form onSubmit={async (e) => {
                e.preventDefault();
                const password = e.target.adminPassword.value;
                try {
                  const res = await fetch('/api/settings/verify-password', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ password })
                  });
                  if (res.ok) {
                    setUserRole('admin');
                  } else {
                    alert('รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง');
                  }
                } catch (err) {
                  console.error('Error verifying password:', err);
                  alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้');
                }
              }} style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                <input 
                  type="password" 
                  name="adminPassword" 
                  placeholder="รหัสผ่านแอดมิน" 
                  style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'rgba(15, 23, 42, 0.5)', color: 'white' }}
                />
                <button type="submit" className="submit-btn" style={{ width: '100%', height: '40px' }}>เข้าสู่ระบบแอดมิน</button>
              </form>
            </div>

            {/* Customer Login Block */}
            <div style={{ padding: '1.25rem', border: '1px solid var(--border-color)', borderRadius: '12px', background: 'rgba(15, 23, 42, 0.3)' }}>
              <h3 style={{ marginBottom: '1rem', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '1.05rem' }}>👥 สำหรับลูกค้า (Customer)</h3>
              {/* Static customer notice removed per request */}
              <form onSubmit={async (e) => {
                e.preventDefault();
                const jobCode = e.target.jobCode.value.trim();

                // Fetch fresh DB from server to ensure we have the latest jobs
                try {
                  const res = await fetch('/api/db');
                  if (!res.ok) throw new Error('Failed to fetch DB');
                  const data = await res.json();
                  const jobExists = data.jobs.find(j => j.id === jobCode);
                  if (!jobExists) {
                    alert('ไม่พบรหัสงานนี้ในระบบ กรุณาตรวจสอบรหัสอีกครั้ง');
                    return;
                  }

                  if (jobExists.expiresAt && new Date() > new Date(jobExists.expiresAt)) {
                    alert('สิทธิ์การเข้าใช้งานเซสชันนี้หมดอายุแล้ว');
                    return;
                  }

                  try {
                    const ack = localStorage.getItem('bioprocess-customer-notice-ack') === 'true';
                    if (ack) {
                      // Apply fresh DB and grant access
                      setMachines(data.machines);
                      setJobs(data.jobs);
                      setCustomers(data.customers);
                      setActiveCustomerJobId(jobCode);
                      setUserRole('customer');
                      setCurrentAppView('monitoring');
                      return;
                    }
                  } catch (err) {
                    // ignore storage errors
                  }

                  // show confirmation notice modal before granting access
                  setPendingJobCode(jobCode);
                  setShowCustomerNotice(true);
                } catch (err) {
                  console.error(err);
                  alert('เกิดข้อผิดพลาดขณะตรวจสอบรหัสงาน กรุณาลองใหม่');
                }
              }} style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                <input 
                  type="text" 
                  name="jobCode" 
                  placeholder="ป้อนรหัสงานของคุณ (เช่น job-...)" 
                  style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'rgba(15, 23, 42, 0.5)', color: 'white' }}
                />
                <button type="submit" className="submit-btn" style={{ width: '100%', height: '40px', background: 'linear-gradient(135deg, var(--accent-blue), #2563eb)' }}>เข้าดูข้อมูลงาน</button>
              </form>
            </div>
          </div>
        </div>
        {showCustomerNotice && (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
            <div style={{ width: 520, maxWidth: '95%', background: 'var(--panel-bg)', border: '1px solid var(--border-color)', borderRadius: 12, padding: '1.25rem' }}>
              <h3 style={{ marginBottom: '0.5rem' }}>⚠️ แจ้งเตือนการเข้าใช้งาน</h3>
              <div style={{ color: 'var(--text-secondary)', marginBottom: '1rem' }}>
                รหัสผ่านนี้ใช้เข้าดูได้เฉพาะข้อมูลในส่วนงานของลูกค้าที่ได้รับอนุญาตเท่านั้น.
                ข้อมูลทั้งหมดจะถูกจัดเก็บเป็นความลับ และระบบจะสำรองข้อมูลไว้เป็นเวลา 7 วันหลังเสร็จงาน.
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '0.5rem', marginTop: '8px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-secondary)' }}>
                  <input type="checkbox" checked={dontShowAgain} onChange={(e) => setDontShowAgain(e.target.checked)} />
                  อย่าขึ้นเตือนอีก
                </label>
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem' }}>
                  <button onClick={() => { setShowCustomerNotice(false); setPendingJobCode(null); }} className="submit-btn" style={{ background: 'transparent', border: '1px solid var(--border-color)', color: 'var(--text-primary)' }}>ยกเลิก</button>
                  <button onClick={async () => {
                    if (dontShowAgain) {
                      try { localStorage.setItem('bioprocess-customer-notice-ack', 'true'); } catch (e) {}
                    }
                    if (pendingJobCode) {
                      try {
                        const res = await fetch('/api/db');
                        if (res.ok) {
                          const data = await res.json();
                          setMachines(data.machines);
                          setJobs(data.jobs);
                          setCustomers(data.customers);
                        }
                      } catch (e) {
                        console.error('Failed to refresh DB before customer access', e);
                      }
                      setActiveCustomerJobId(pendingJobCode);
                      setUserRole('customer');
                      setCurrentAppView('monitoring');
                    }
                    setShowCustomerNotice(false);
                    setPendingJobCode(null);
                  }} className="submit-btn">ยอมรับและเข้าสู่ระบบ</button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }


  return (
    <div className="app-layout">
      {/* Customer Warning Toast Notification */}
      {showCustomerToast && (
        <div className={`login-toast-container ${isToastHiding ? 'hiding' : ''}`}>
          <div className="login-toast">
            <div className="login-toast-header">
              <span>⚠️ แจ้งเตือนการใช้งานระบบ</span>
              <button className="login-toast-close" onClick={dismissCustomerToast}>&times;</button>
            </div>
            <div className="login-toast-body">
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>ระบบนี้จัดทำขึ้นเพื่อบันทึกและจัดเก็บข้อมูลสำหรับการปฏิบัติงานและการจัดทำรายงานเท่านั้น</span>
              </div>
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>ID และรหัสผ่านเป็นข้อมูลสำคัญ ห้ามเปิดเผยหรือแชร์ให้บุคคลที่ไม่เกี่ยวข้อง</span>
              </div>
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>สิทธิ์การเข้าใช้งานระบบถูกกำหนดเฉพาะผู้ได้รับอนุญาตเท่านั้น</span>
              </div>
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>ข้อมูลจะถูกสำรองไว้ 7 วัน นับจากวันที่บันทึกข้อมูลล่าสุด</span>
              </div>
            </div>
          </div>
        </div>
      )}
      {showCustomerNotice && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
          <div style={{ width: 520, maxWidth: '95%', background: 'var(--panel-bg)', border: '1px solid var(--border-color)', borderRadius: 12, padding: '1.25rem' }}>
            <h3 style={{ marginBottom: '0.5rem' }}>⚠️ แจ้งเตือนการเข้าใช้งาน</h3>
            <div style={{ color: 'var(--text-secondary)', marginBottom: '1rem' }}>
              รหัสผ่านนี้ใช้เข้าดูได้เฉพาะข้อมูลในส่วนงานของลูกค้าที่ได้รับอนุญาตเท่านั้น.
              ข้อมูลทั้งหมดจะถูกจัดเก็บเป็นความลับ และระบบจะสำรองข้อมูลไว้เป็นเวลา 7 วันหลังเสร็จงาน.
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '0.5rem', marginTop: '8px' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-secondary)' }}>
                <input type="checkbox" checked={dontShowAgain} onChange={(e) => setDontShowAgain(e.target.checked)} />
                อย่าขึ้นเตือนอีก
              </label>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem' }}>
                <button onClick={() => { setShowCustomerNotice(false); setPendingJobCode(null); }} className="submit-btn" style={{ background: 'transparent', border: '1px solid var(--border-color)', color: 'var(--text-primary)' }}>ยกเลิก</button>
                <button onClick={async () => {
                  if (dontShowAgain) {
                      try { localStorage.setItem('bioprocess-customer-notice-ack', 'true'); } catch (e) {}
                    }
                    if (pendingJobCode) {
                      // Refresh DB before granting access to ensure currentJobId/machine are set
                      try {
                        const res = await fetch('/api/db');
                        if (res.ok) {
                          const data = await res.json();
                          setMachines(data.machines);
                          setJobs(data.jobs);
                          setCustomers(data.customers);
                        }
                      } catch (e) {
                        console.error('Failed to refresh DB before customer access', e);
                      }
                      setActiveCustomerJobId(pendingJobCode);
                      setUserRole('customer');
                      setCurrentAppView('monitoring');
                    }
                  setShowCustomerNotice(false);
                  setPendingJobCode(null);
                }} className="submit-btn">ยอมรับและเข้าสู่ระบบ</button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Mobile Header Bar */}
      <div className="mobile-topbar">
        <button 
          onClick={() => setIsMobileMenuOpen(true)}
          style={{ background: 'transparent', border: 'none', color: 'white', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
        >
          <Menu size={24} />
        </button>
        <span className="mobile-topbar-title">
          <ActivityIcon size={20} style={{ color: '#00f0ff', filter: 'drop-shadow(0 0 6px rgba(0, 240, 255, 0.5))' }} />
          BIOPROCESS
        </span>
        <span style={{ fontSize: '0.8rem', color: '#84b2bc', fontWeight: 600 }}>
          {currentMachine?.name || 'Bioprocess'}
        </span>
      </div>

      {/* Backdrop overlay for mobile menu */}
      {isMobileMenuOpen && (
        <div 
          className="mobile-overlay" 
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}

      {/* SIDEBAR */}
      <aside className={`sidebar ${isMobileMenuOpen ? 'open' : ''}`}>
        {/* Header */}
        <div className="sidebar-header">
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', width: '100%', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
              <ActivityIcon size={28} style={{ color: '#00f0ff', filter: 'drop-shadow(0 0 6px rgba(0, 240, 255, 0.5))' }} />
              <div>
                <div className="sidebar-logo-text">DBMS</div>
                <div className="sidebar-logo-subtext">SYSTEM ULTRA</div>
              </div>
            </div>
            {/* Close button on mobile sidebar header */}
            <button 
              onClick={() => setIsMobileMenuOpen(false)}
              style={{ background: 'transparent', border: 'none', color: '#84b2bc', cursor: 'pointer', display: 'none' }}
              className="mobile-close-btn"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Menu */}
        <div className="sidebar-menu">
          {userRole === 'admin' ? (
            <>
              {/* Active Users Status Component */}
              <div className="sidebar-active-users" style={{ padding: '0 1.25rem 1rem 1.25rem', borderBottom: '1px solid rgba(255, 255, 255, 0.05)', marginBottom: '0.5rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '0.5rem', fontWeight: 700 }}>
                  <span style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: '#10b981', display: 'inline-block', boxShadow: '0 0 8px #10b981' }} className="status-dot"></span>
                  ผู้ใช้งานออนไลน์ ({activeUsers.length})
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  {activeUsers.map((u, i) => {
                    const isSelf = u.clientId === clientId;
                    const mName = machines.find(m => m.id === u.machineId)?.name || '';
                    const jName = jobs.find(j => j.id === u.jobId)?.name || '';
                    return (
                      <div key={u.clientId || i} style={{ fontSize: '0.8rem', padding: '6px 8px', borderRadius: '6px', background: isSelf ? 'rgba(0, 240, 255, 0.03)' : 'rgba(255, 255, 255, 0.02)', border: isSelf ? '1px solid rgba(0, 240, 255, 0.15)' : '1px solid rgba(255, 255, 255, 0.05)' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 600, color: isSelf ? '#00f0ff' : 'var(--text-primary)' }}>
                          <span>
                            {u.role === 'admin' ? '🔐 แอดมิน' : '👥 ลูกค้า'}
                            {isSelf && ' (คุณ)'}
                          </span>
                          <span style={{ fontSize: '0.7rem', color: '#10b981', fontWeight: 600 }}>Active</span>
                        </div>
                        {(mName || jName) && (
                          <div style={{ fontSize: '0.7rem', color: '#84b2bc', marginTop: '2px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            กำลังดู: {mName} {jName && `› ${jName}`}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Category: Monitoring */}
              <div className="sidebar-menu-header">Monitoring</div>

              {/* Menu Dashboard */}
              <div 
                className={`sidebar-menu-item ${currentAppView === 'monitoring' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('monitoring');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <LayoutDashboard size={18} />
                  Dashboard / Monitor
                </span>
              </div>

              {/* Instrument Accordion */}
              <div 
                className="sidebar-menu-item"
                onClick={() => setIsInstrumentsExpanded(!isInstrumentsExpanded)}
              >
                <span className="sidebar-menu-link">
                  <Cpu size={18} />
                  Instruments
                </span>
                <span className="sidebar-menu-arrow">
                  {isInstrumentsExpanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                </span>
              </div>
              
              {isInstrumentsExpanded && (
                <div className="sidebar-submenu">
                  {machines.map(m => (
                    <div 
                      key={m.id}
                      className={`sidebar-submenu-item ${m.id === currentMachineId && currentAppView === 'monitoring' ? 'active' : ''}`}
                      onClick={() => {
                        setCurrentAppView('monitoring');
                        setCurrentMachineId(m.id);
                        setIsMobileMenuOpen(false);
                        // Filter and auto-select first job for this machine
                        const machineJobs = jobs.filter(j => j.machineId === m.id);
                        if (machineJobs.length > 0) {
                          setCurrentJobId(machineJobs[0].id);
                        } else {
                          setCurrentJobId(null);
                        }
                      }}
                    >
                      <span style={{ display: 'flex', alignItems: 'center', gap: '4px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '140px' }}>
                        {m.id === currentMachineId && currentAppView === 'monitoring' && <span className="sidebar-submenu-dot" />}
                        {m.name}
                      </span>
                      {m.id === currentMachineId && (
                        <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }} onClick={e => e.stopPropagation()}>
                          <Edit3 size={13} style={{ cursor: 'pointer', opacity: 0.8 }} onClick={renameMachine} title="Rename" />
                          {machines.length > 1 && (
                            <Trash2 size={13} style={{ cursor: 'pointer', color: '#ef4444', opacity: 0.8 }} onClick={deleteMachine} title="Delete" />
                          )}
                        </div>
                      )}
                    </div>
                  ))}
                  <div 
                    className="sidebar-submenu-item" 
                    style={{ color: '#00f0ff', fontWeight: 600 }}
                    onClick={() => {
                      setIsMobileMenuOpen(false);
                      handleMachineChange({ target: { value: 'ADD_NEW' } });
                    }}
                  >
                    + Add Instrument...
                  </div>
                </div>
              )}

              {/* Sessions Accordion */}
              {currentAppView === 'monitoring' && currentMachineId && (
                <>
                  <div 
                    className="sidebar-menu-item"
                    onClick={() => setIsSessionsExpanded(!isSessionsExpanded)}
                  >
                    <span className="sidebar-menu-link">
                      <FolderOpen size={18} />
                      Sessions
                    </span>
                    <span className="sidebar-menu-arrow">
                      {isSessionsExpanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                    </span>
                  </div>
                  
                  {isSessionsExpanded && (
                    <div className="sidebar-submenu">
                      {jobsForMachine.map(job => (
                        <div 
                          key={job.id}
                          className={`sidebar-submenu-item ${job.id === currentJobId ? 'active' : ''}`}
                          onClick={() => {
                            setCurrentJobId(job.id);
                            setIsMobileMenuOpen(false);
                          }}
                          style={{ flexDirection: 'column', alignItems: 'flex-start', gap: '2px', padding: '6px 8px' }}
                        >
                          <div style={{ display: 'flex', width: '100%', justifyContent: 'space-between', alignItems: 'center' }}>
                            <span style={{ display: 'flex', alignItems: 'center', fontWeight: job.id === currentJobId ? '600' : 'normal', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '140px' }}>
                              {job.id === currentJobId && <span className="sidebar-submenu-dot" />}
                              {job.name}
                            </span>
                            <div style={{ display: 'flex', gap: '8px' }} onClick={e => e.stopPropagation()}>
                              <Trash2 size={13} style={{ cursor: 'pointer', color: '#ef4444', opacity: 0.8 }} onClick={(e) => deleteJob(job.id, e)} title="Delete Session" />
                            </div>
                          </div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', width: '100%', fontSize: '0.7rem', color: '#688d96', marginTop: '2px' }}>
                            <span>{job.createdAt.split(',')[0]}</span>
                            <div style={{ display: 'flex', gap: '6px' }}>
                              <span 
                                style={{ color: '#10b981', cursor: 'pointer', background: 'rgba(16, 185, 129, 0.08)', padding: '0 4px', borderRadius: '3px', fontWeight: 600 }}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleCopyShareLink(job);
                                }}
                                title="Click to Copy Direct Login Link for Customer"
                              >
                                ลิงก์แชร์
                              </span>
                              <span 
                                style={{ color: '#00f0ff', cursor: 'pointer', background: 'rgba(0, 240, 255, 0.05)', padding: '0 4px', borderRadius: '3px' }}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  navigator.clipboard.writeText(job.id);
                                  alert(`คัดลอกรหัสงานเรียบร้อย: ${job.id}`);
                                }}
                                title="Click to copy Code"
                              >
                                Code: {job.id.substring(4, 9)}..
                              </span>
                            </div>
                          </div>
                        </div>
                      ))}
                      <div 
                        className="sidebar-submenu-item"
                        style={{ color: '#00f0ff', fontWeight: 600 }}
                        onClick={() => {
                          setIsMobileMenuOpen(false);
                          createNewJob();
                        }}
                      >
                        + Add Session...
                      </div>
                    </div>
                  )}
                </>
              )}

              {/* Category: Management */}
              <div className="sidebar-menu-header">Management</div>

              {/* Menu Customer Database */}
              <div 
                className={`sidebar-menu-item ${currentAppView === 'customers' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('customers');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <Users size={18} />
                  Customer Database
                </span>
                <span className="sidebar-badge">{customers.length}</span>
              </div>

              {/* Menu System Settings */}
              <div 
                className={`sidebar-menu-item ${currentAppView === 'settings' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('settings');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <Settings size={18} />
                  System Settings
                </span>
              </div>
            </>
          ) : (
            /* Customer Portal Navigation */
            <div style={{ padding: '0 1.25rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <div style={{ borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: '1rem' }}>
                <div style={{ fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: '0.25rem' }}>👤 Customer Portal</div>
                <div style={{ fontSize: '0.85rem', fontWeight: 600, color: '#00f0ff', wordBreak: 'break-all' }}>ID: {activeCustomerJobId}</div>
              </div>

              <div>
                <div style={{ fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Instrument</div>
                <div style={{ fontSize: '0.95rem', fontWeight: 600, color: '#ffffff', marginTop: '0.25rem' }}>{currentMachine?.name}</div>
              </div>

              <div>
                <div style={{ fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Session</div>
                <div style={{ fontSize: '0.95rem', fontWeight: 600, color: '#ffffff', marginTop: '0.25rem' }}>{currentJob?.name}</div>
              </div>

              {showCustomerBanner && (
                <div style={{ padding: '12px', borderRadius: '8px', background: 'rgba(255,255,255,0.02)', borderLeft: '3px solid #f59e0b', fontSize: '0.8rem', lineHeight: 1.4, color: '#cbdce0' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                    <div>
                      <div style={{ fontWeight: 700, color: '#f59e0b', marginBottom: '4px' }}>⚠️ บันทึกข้อมูลและสิทธิ์</div>
                      <div>• สิทธิ์ในการอ่านข้อมูลเท่านั้น</div>
                      <div>• ข้อมูลถูกเก็บเป็นความลับ</div>
                      <div>• สำรองข้อมูล 7 วันหลังจบงาน</div>
                    </div>
                    <button onClick={() => setShowCustomerBanner(false)} style={{ background: 'transparent', border: 'none', color: '#84b2bc', cursor: 'pointer' }}>✕</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="sidebar-footer">
          <button 
            className="sidebar-logout-btn"
            onClick={() => {
              if (window.confirm("ยืนยันการออกจากระบบ?")) {
                setUserRole(null);
                setActiveCustomerJobId(null);
                setIsMobileMenuOpen(false);
              }
            }}
          >
            <LogOut size={16} />
            Logout
          </button>
        </div>
      </aside>

      {/* MAIN CONTENT */}
      <main className="main-content">
        {/* Print-only Header */}
        <div className="print-only-header">
          <h1>Bioprocess Data Report</h1>
          <p><strong>Session Name:</strong> {currentJob?.name || '-'}</p>
          <p><strong>Instrument / Machine:</strong> {currentMachine?.name || '-'}</p>
          <p><strong>Date Generated:</strong> {new Date().toLocaleString('th-TH')}</p>
        </div>
        {currentAppView === 'customers' ? (
          /* CUSTOMER DATA VIEW */
          <div className="customers-view">
            <header className="dashboard-header">
              <h2>Customer Database</h2>
            </header>

            <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1rem' }}>
              <h2 className="form-title">Add New Customer</h2>
              <form onSubmit={handleAddCustomer} className="customer-add-form">
                <div className="form-group" style={{ flex: 2 }}>
                  <label>Company / Customer Name</label>
                  <input
                    type="text"
                    name="companyName"
                    placeholder="บริษัท/ชื่อลูกค้า"
                    value={customerFormData.companyName}
                    onChange={handleCustomerInputChange}
                  />
                </div>

                <div className="form-group" style={{ flex: 1 }}>
                  <label>Email (อีเมล)</label>
                  <input
                    type="email"
                    name="email"
                    placeholder="customer@example.com"
                    value={customerFormData.email}
                    onChange={handleCustomerInputChange}
                    style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', height: '40px' }}
                  />
                </div>

                <div className="form-group" style={{ flex: 1 }}>
                  <label>Instrument / Machine Used</label>
                  <select
                    name="machineId"
                    value={customerFormData.machineId}
                    onChange={handleCustomerInputChange}
                    className="machine-dropdown"
                    style={{ padding: '10px', height: '42px', backgroundImage: 'none' }}
                  >
                    {machines.map(m => (
                      <option key={m.id} value={m.id}>{m.name}</option>
                    ))}
                  </select>
                </div>

                <button type="submit" className="submit-btn" style={{ height: '42px' }}>
                  <Users size={18} style={{ marginRight: '8px' }} /> Add Customer
                </button>
              </form>
            </div>

            <div className="glass-panel" style={{ padding: '2rem' }}>
              <div className="table-header-controls">
                <h2 className="chart-title">Customer List</h2>
                <span style={{ color: 'var(--text-secondary)' }}>Total: {customers.length}</span>
              </div>
              <div className="data-table-container">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Company / Customer Name</th>
                      <th>Email</th>
                      <th>Machine Assigned</th>
                      <th>Session & Share Link</th>
                      <th>Date Added</th>
                      <th style={{ width: '80px', textAlign: 'center' }}>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {customers.length === 0 ? (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>
                          No customers recorded yet.
                        </td>
                      </tr>
                    ) : (
                      customers.map((c) => (
                        <tr key={c.id}>
                          <td style={{ fontWeight: 600 }}>{c.companyName}</td>
                          <td style={{ color: 'var(--text-secondary)' }}>{c.email || '-'}</td>
                          <td>
                            <span style={{ background: 'rgba(59, 130, 246, 0.2)', color: 'var(--accent-blue)', padding: '4px 8px', borderRadius: '4px', fontSize: '0.85rem' }}>
                              {getMachineName(c.machineId)}
                            </span>
                          </td>
                          <td>
                            {(() => {
                              const custJobs = jobs.filter(j => j.machineId === c.machineId);
                              if (custJobs.length === 0) {
                                return <span style={{ color: 'var(--text-secondary)', fontSize: '0.85rem' }}>ไม่มีเซสชันในเครื่องนี้</span>;
                              }
                              const selectedJobId = selectedCustomerJobs[c.id] || custJobs[0].id;
                              return (
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                  <select
                                    value={selectedJobId}
                                    onChange={(e) => setSelectedCustomerJobs(prev => ({ ...prev, [c.id]: e.target.value }))}
                                    className="machine-dropdown"
                                    style={{ padding: '4px 8px', height: '32px', fontSize: '0.85rem', width: 'auto', display: 'inline-block', minWidth: '120px', backgroundImage: 'none' }}
                                  >
                                    {custJobs.map(j => (
                                      <option key={j.id} value={j.id}>{j.name}</option>
                                    ))}
                                  </select>
                                  <button
                                    className="submit-btn"
                                    style={{
                                      padding: '4px 10px',
                                      height: '32px',
                                      fontSize: '0.85rem',
                                      background: 'linear-gradient(135deg, #10b981, #059669)',
                                      border: 'none',
                                      borderRadius: '6px',
                                      color: 'white',
                                      cursor: 'pointer',
                                      fontWeight: 600,
                                      display: 'inline-flex',
                                      alignItems: 'center',
                                      gap: '4px'
                                    }}
                                    onClick={() => {
                                      const targetJob = custJobs.find(j => j.id === selectedJobId) || custJobs[0];
                                      handleCopyShareLink(targetJob);
                                    }}
                                  >
                                    ลิงก์แชร์
                                  </button>
                                </div>
                              );
                            })()}
                          </td>
                          <td style={{ color: 'var(--text-secondary)' }}>{c.createdAt}</td>
                          <td style={{ textAlign: 'center' }}>
                            <button
                              className="delete-row-btn"
                              onClick={() => deleteCustomer(c.id)}
                              title="Delete customer"
                            >
                              <Trash2 size={16} />
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
        ) : currentAppView === 'settings' ? (
          /* SYSTEM SETTINGS VIEW */
          <div className="settings-view">
            <header className="dashboard-header">
              <h2>System Settings</h2>
            </header>

            <div className="glass-panel" style={{ padding: '2rem', maxWidth: '500px' }}>
              <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <Settings size={20} color="var(--accent-blue)" /> ตั้งค่ารหัสผ่านผู้ดูแลระบบ (Change Admin Password)
              </h3>
              
              <form onSubmit={async (e) => {
                e.preventDefault();
                const currentPassword = e.target.currentPassword.value;
                const newPassword = e.target.newPassword.value;
                const confirmPassword = e.target.confirmPassword.value;

                if (!currentPassword || !newPassword || !confirmPassword) {
                  alert('กรุณากรอกข้อมูลให้ครบถ้วน');
                  return;
                }

                if (newPassword !== confirmPassword) {
                  alert('รหัสผ่านใหม่และยืนยันรหัสผ่านใหม่ไม่ตรงกัน');
                  return;
                }

                try {
                  const res = await fetch('/api/settings/update-password', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ currentPassword, newPassword })
                  });
                  const result = await res.json();
                  if (res.ok) {
                    alert('เปลี่ยนรหัสผ่านสำเร็จเรียบร้อยแล้ว');
                    e.target.reset();
                  } else {
                    alert(`ผิดพลาด: ${result.error}`);
                  }
                } catch (err) {
                  console.error('Error changing password:', err);
                  alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์เพื่อเปลี่ยนรหัสผ่านได้');
                }
              }} className="data-form" style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                
                <div className="form-group" style={{ width: '100%' }}>
                  <label>รหัสผ่านปัจจุบัน (Current Password)</label>
                  <input
                    type="password"
                    name="currentPassword"
                    placeholder="ป้อนรหัสผ่านปัจจุบัน"
                    style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'rgba(15, 23, 42, 0.5)', color: 'white', width: '100%' }}
                  />
                </div>

                <div className="form-group" style={{ width: '100%' }}>
                  <label>รหัสผ่านใหม่ (New Password)</label>
                  <input
                    type="password"
                    name="newPassword"
                    placeholder="ป้อนรหัสผ่านใหม่"
                    style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'rgba(15, 23, 42, 0.5)', color: 'white', width: '100%' }}
                  />
                </div>

                <div className="form-group" style={{ width: '100%' }}>
                  <label>ยืนยันรหัสผ่านใหม่ (Confirm New Password)</label>
                  <input
                    type="password"
                    name="confirmPassword"
                    placeholder="ป้อนยืนยันรหัสผ่านใหม่"
                    style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'rgba(15, 23, 42, 0.5)', color: 'white', width: '100%' }}
                  />
                </div>

                <button type="submit" className="submit-btn" style={{ width: '100%', height: '42px', marginTop: '0.5rem' }}>
                  บันทึกการตั้งค่า (Update Password)
                </button>
              </form>
            </div>
          </div>
        ) : (
          /* MONITORING VIEW */
          <>
            <header className="dashboard-header">
              <div style={{ display: 'flex', alignItems: 'center', gap: '2rem' }}>
                <h2>{currentJob?.name || 'No Session Selected'}</h2>
                
                {currentJob && (
                  <div className="nav-tabs">
                    <button 
                      className={`nav-tab ${activeTab === 'dashboard' ? 'active' : ''}`}
                      onClick={() => setActiveTab('dashboard')}
                    >
                      <LayoutDashboard size={18} /> Dashboard
                    </button>
                    <button 
                      className={`nav-tab ${activeTab === 'combined' ? 'active' : ''}`}
                      onClick={() => setActiveTab('combined')}
                    >
                      <ChartIcon size={18} /> Combined Graph
                    </button>
                    <button 
                      className={`nav-tab ${activeTab === 'table' ? 'active' : ''}`}
                      onClick={() => setActiveTab('table')}
                    >
                      <TableIcon size={18} /> Data Table
                    </button>
                  </div>
                )}
              </div>

              <div className="header-controls" style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
                <button
                  className={`theme-toggle-btn`}
                  onClick={() => setTheme(prev => prev === 'dark' ? 'light' : 'dark')}
                  title="Toggle theme"
                  style={{ margin: 0 }}
                >
                  {theme === 'dark' ? '🌙' : '☀️'}
                </button>
                {userRole === 'admin' && (
                  <button
                    className={`toggle-btn ${isReplay ? 'active' : ''}`}
                    onClick={() => {
                      if (isReplay) {
                        setIsReplay(false);
                        setReplayIndex(0);
                        setReplayVisibleData([]);
                      } else {
                        // start replay from stored data
                        setReplayVisibleData([]);
                        setReplayIndex(0);
                        setIsReplay(true);
                      }
                    }}
                    disabled={!currentJob || currentJobData.length === 0}
                    style={{ margin: 0 }}
                  >
                    {isReplay ? 'Replay ON' : 'Replay'}
                  </button>
                )}
                <button className="export-btn" onClick={exportToExcel} disabled={!currentJob} style={{ margin: 0 }}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export Excel
                </button>
                <button className="export-btn" onClick={exportToPDF} disabled={!currentJob} style={{ margin: 0 }}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export PDF
                </button>
                <button className="export-btn" onClick={exportToCSV} disabled={!currentJob} style={{ margin: 0 }}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export CSV
                </button>
                {/* Auto simulation removed */}
              </div>
            </header>

            {currentJob && userRole === 'admin' && (
              /* Manual Input Form */
              <div className="glass-panel form-container">
                <h2 className="form-title">Manual Data Entry (Record Data {currentMachine?.name || 'เครื่องมือ'})</h2>
                <form onSubmit={handleManualSubmit} className="data-form">
                  <div className="form-group-container">
                    <div className="form-group">
                      <label>TEMP (°C)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.01" name="temp_set" value={formData.temp_set} onChange={handleInputChange} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.01" name="temp_read" value={formData.temp_read} onChange={handleInputChange} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>pH</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.01" name="ph_set" value={formData.ph_set} onChange={handleInputChange} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.01" name="ph_read" value={formData.ph_read} onChange={handleInputChange} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>DO (%)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="1" name="do_set" value={formData.do_set} onChange={handleInputChange} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="1" name="do_read" value={formData.do_read} onChange={handleInputChange} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>AGIT (RPM)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="1" name="agit_set" value={formData.agit_set} onChange={handleInputChange} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="1" name="agit_read" value={formData.agit_read} onChange={handleInputChange} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>AIR (L/M)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.1" name="air_set" value={formData.air_set} onChange={handleInputChange} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.1" name="air_read" value={formData.air_read} onChange={handleInputChange} />
                        </div>
                      </div>
                    </div>
                    {/* Date and Time manual inputs */}
                    <div className="form-group" style={{ minWidth: '180px' }}>
                      <label>📅 DATE (วันที่บันทึก)</label>
                      <input 
                        type="date" 
                        name="date" 
                        value={formData.date} 
                        onChange={handleInputChange}
                        style={{ 
                          width: '100%', 
                          padding: '10px 14px', 
                          borderRadius: '8px', 
                          border: '1px solid var(--border-color)', 
                          background: 'rgba(15, 23, 42, 0.5)', 
                          color: 'white',
                          fontFamily: 'inherit',
                          fontSize: '0.95rem'
                        }}
                      />
                    </div>
                    <div className="form-group" style={{ minWidth: '180px' }}>
                      <label>⏰ TIME (เวลาที่บันทึก)</label>
                      <input 
                        type="time" 
                        name="time" 
                        value={formData.time} 
                        onChange={handleInputChange}
                        style={{ 
                          width: '100%', 
                          padding: '10px 14px', 
                          borderRadius: '8px', 
                          border: '1px solid var(--border-color)', 
                          background: 'rgba(15, 23, 42, 0.5)', 
                          color: 'white',
                          fontFamily: 'inherit',
                          fontSize: '0.95rem'
                        }}
                      />
                    </div>

                    {/* Remarks/Notes full-width input */}
                    <div className="form-group" style={{ flex: '1 1 100%' }}>
                      <label>REMARKS / NOTES (บันทึกข้อความ)</label>
                      <input 
                        type="text" 
                        name="remark" 
                        placeholder="ระบุหมายเหตุหรือข้อความบันทึกที่นี่ (เช่น ปรับค่าอัตราไหล, ตรวจสภาพโพรบ)..." 
                        value={formData.remark} 
                        onChange={handleInputChange}
                        style={{ 
                          width: '100%', 
                          padding: '10px 14px', 
                          borderRadius: '8px', 
                          border: '1px solid var(--border-color)', 
                          background: 'rgba(15, 23, 42, 0.5)', 
                          color: 'white',
                          textAlign: 'left',
                          fontFamily: 'inherit',
                          fontSize: '0.95rem'
                        }} 
                      />
                    </div>

                    {/* Date and Time are set automatically to current values */}
                  </div>
                  <button type="submit" className="submit-btn" style={{ minWidth: '160px' }}>
                    <PlusCircle size={18} style={{ marginRight: '8px' }} /> Add Record
                  </button>
                </form>
              </div>
            )}

            {!currentJob ? (
              <div className="glass-panel empty-state">
                <h2>Select or create a session to start.</h2>
              </div>
                ) : currentJobData.length === 0 ? (
              <div className="glass-panel empty-state">
                <FolderOpen size={48} opacity={0.5} />
                <h2>No data recorded yet</h2>
                <p>Use the manual entry form above to add your first record.</p>
              </div>
            ) : (
              activeTab === 'dashboard' ? (
                <>
                  {/* Real-time Metrics Grid */}
                  <div className="metrics-grid">
                    <div className="glass-panel metric-card">
                      <Thermometer color="var(--accent-red)" size={24} style={{ marginBottom: '10px' }} />
                      <div className="metric-title">Temperature</div>
                      <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div className="metric-value" style={{ color: 'var(--accent-red)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                          {typeof (lastDataPoint.temp_read !== undefined ? lastDataPoint.temp_read : lastDataPoint?.temp) === 'number' ? ((lastDataPoint.temp_read !== undefined ? lastDataPoint.temp_read : lastDataPoint?.temp) || 0).toFixed(2) : '-'}<span className="metric-unit">°C</span>
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {typeof (lastDataPoint.temp_set !== undefined ? lastDataPoint.temp_set : lastDataPoint?.temp) === 'number' ? ((lastDataPoint.temp_set !== undefined ? lastDataPoint.temp_set : lastDataPoint?.temp) || 0).toFixed(2) : '-'}°C
                        </div>
                      </div>
                    </div>
                    <div className="glass-panel metric-card">
                      <Droplets color="var(--accent-blue)" size={24} style={{ marginBottom: '10px' }} />
                      <div className="metric-title">pH Level</div>
                      <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div className="metric-value" style={{ color: 'var(--accent-blue)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                          {typeof (lastDataPoint.ph_read !== undefined ? lastDataPoint.ph_read : lastDataPoint?.ph) === 'number' ? ((lastDataPoint.ph_read !== undefined ? lastDataPoint.ph_read : lastDataPoint?.ph) || 0).toFixed(2) : '-'}
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {typeof (lastDataPoint.ph_set !== undefined ? lastDataPoint.ph_set : lastDataPoint?.ph) === 'number' ? ((lastDataPoint.ph_set !== undefined ? lastDataPoint.ph_set : lastDataPoint?.ph) || 0).toFixed(2) : '-'}
                        </div>
                      </div>
                    </div>
                    <div className="glass-panel metric-card">
                      <Activity color="var(--accent-green)" size={24} style={{ marginBottom: '10px' }} />
                      <div className="metric-title">Dissolved Oxygen</div>
                      <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div className="metric-value" style={{ color: 'var(--accent-green)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                          {lastDataPoint.do_read !== undefined ? lastDataPoint.do_read : lastDataPoint.do}<span className="metric-unit">%</span>
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {lastDataPoint.do_set !== undefined ? lastDataPoint.do_set : lastDataPoint.do}%
                        </div>
                      </div>
                    </div>
                    <div className="glass-panel metric-card">
                      <RotateCw color="var(--accent-yellow)" size={24} style={{ marginBottom: '10px' }} />
                      <div className="metric-title">Agitation</div>
                      <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div className="metric-value" style={{ color: 'var(--accent-yellow)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                          {lastDataPoint.agit_read !== undefined ? lastDataPoint.agit_read : lastDataPoint.agit}<span className="metric-unit">RPM</span>
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {lastDataPoint.agit_set !== undefined ? lastDataPoint.agit_set : lastDataPoint.agit} RPM
                        </div>
                      </div>
                    </div>
                    <div className="glass-panel metric-card">
                      <Wind color="var(--accent-purple)" size={24} style={{ marginBottom: '10px' }} />
                      <div className="metric-title">Air Flow</div>
                      <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div className="metric-value" style={{ color: 'var(--accent-purple)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                          {typeof (lastDataPoint.air_read !== undefined ? lastDataPoint.air_read : lastDataPoint?.air) === 'number' ? ((lastDataPoint.air_read !== undefined ? lastDataPoint.air_read : lastDataPoint?.air) || 0).toFixed(1) : '-'}<span className="metric-unit">L/M</span>
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {typeof (lastDataPoint.air_set !== undefined ? lastDataPoint.air_set : lastDataPoint?.air) === 'number' ? ((lastDataPoint.air_set !== undefined ? lastDataPoint.air_set : lastDataPoint?.air) || 0).toFixed(1) : '-'} L/M
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Charts Grid */}
                  <div className="charts-grid">
                    <div className="glass-panel chart-container">
                      <div className="chart-header">
                        <h2 className="chart-title" style={{ color: 'var(--accent-red)' }}>Temperature Trend</h2>
                      </div>
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="time" stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                          <YAxis domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{fontSize: 12}} tickFormatter={(val) => typeof val === 'number' ? val.toFixed(2) : val} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line type="monotone" dataKey="temp_read" name="TEMP PV (Read)" stroke="var(--accent-red)" strokeWidth={3} dot={true} activeDot={{ r: 8 }} />
                          <Line type="monotone" dataKey="temp_set" name="TEMP SV (Set)" stroke="var(--accent-red)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>

                    <div className="glass-panel chart-container">
                      <div className="chart-header">
                        <h2 className="chart-title" style={{ color: 'var(--accent-blue)' }}>pH Trend</h2>
                      </div>
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="time" stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                          <YAxis domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line type="monotone" dataKey="ph_read" name="pH PV (Read)" stroke="var(--accent-blue)" strokeWidth={3} dot={true} activeDot={{ r: 8 }} />
                          <Line type="monotone" dataKey="ph_set" name="pH SV (Set)" stroke="var(--accent-blue)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>

                    <div className="glass-panel chart-container">
                      <div className="chart-header">
                        <h2 className="chart-title" style={{ color: 'var(--accent-green)' }}>DO & Agitation Trend</h2>
                      </div>
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="time" stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                          <YAxis yAxisId="left" stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                          <YAxis yAxisId="right" orientation="right" domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line yAxisId="left" type="monotone" dataKey="do_read" name="DO PV (%)" stroke="var(--accent-green)" strokeWidth={3} dot={true} />
                          <Line yAxisId="left" type="monotone" dataKey="do_set" name="DO SV (%)" stroke="var(--accent-green)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                          <Line yAxisId="right" type="monotone" dataKey="agit_read" name="AGIT PV (RPM)" stroke="var(--accent-yellow)" strokeWidth={3} dot={true} />
                          <Line yAxisId="right" type="monotone" dataKey="agit_set" name="AGIT SV (RPM)" stroke="var(--accent-yellow)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </>
                ) : activeTab === 'combined' ? (
                <div className="glass-panel" style={{ height: '650px', padding: '2rem' }}>
                  <div className="chart-header" style={{ marginBottom: '1.5rem', flexDirection: 'column', alignItems: 'flex-start', gap: '1rem' }}>
                    <h2 className="chart-title">Combined Bioprocess Overview</h2>
                    
                    {/* Toggle Selector Buttons */}
                    <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', width: '100%', padding: '0.5rem', background: 'rgba(15, 23, 42, 0.3)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                      <button 
                        className="toggle-btn"
                        style={{ 
                          borderColor: 'var(--accent-red)', 
                          color: visibleParameters.temp ? 'white' : 'var(--accent-red)', 
                          background: visibleParameters.temp ? 'rgba(239, 68, 68, 0.2)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, temp: !prev.temp }))}
                      >
                        TEMP ({visibleParameters.temp ? 'ON' : 'OFF'})
                      </button>
                      <button 
                        className="toggle-btn"
                        style={{ 
                          borderColor: 'var(--accent-blue)', 
                          color: visibleParameters.ph ? 'white' : 'var(--accent-blue)', 
                          background: visibleParameters.ph ? 'rgba(59, 130, 246, 0.2)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, ph: !prev.ph }))}
                      >
                        pH ({visibleParameters.ph ? 'ON' : 'OFF'})
                      </button>
                      <button 
                        className="toggle-btn"
                        style={{ 
                          borderColor: 'var(--accent-green)', 
                          color: visibleParameters.do ? 'white' : 'var(--accent-green)', 
                          background: visibleParameters.do ? 'rgba(16, 185, 129, 0.2)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, do: !prev.do }))}
                      >
                        DO ({visibleParameters.do ? 'ON' : 'OFF'})
                      </button>
                      <button 
                        className="toggle-btn"
                        style={{ 
                          borderColor: 'var(--accent-yellow)', 
                          color: visibleParameters.agit ? 'white' : 'var(--accent-yellow)', 
                          background: visibleParameters.agit ? 'rgba(245, 158, 11, 0.2)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, agit: !prev.agit }))}
                      >
                        AGIT ({visibleParameters.agit ? 'ON' : 'OFF'})
                      </button>
                      <button 
                        className="toggle-btn"
                        style={{ 
                          borderColor: 'var(--accent-purple)', 
                          color: visibleParameters.air ? 'white' : 'var(--accent-purple)', 
                          background: visibleParameters.air ? 'rgba(139, 92, 246, 0.2)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, air: !prev.air }))}
                      >
                        AIR ({visibleParameters.air ? 'ON' : 'OFF'})
                      </button>
                    </div>
                  </div>
                  <ResponsiveContainer width="100%" height="80%">
                    <LineChart data={chartData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                      <XAxis dataKey="time" stroke="var(--text-secondary)" tick={{fontSize: 12}} />
                      <YAxis yAxisId="left" stroke="var(--text-secondary)" tick={{fontSize: 12}} label={{ value: 'TEMP / pH / DO / AIR', angle: -90, position: 'insideLeft', fill: 'var(--text-secondary)' }} />
                      <YAxis yAxisId="right" orientation="right" domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{fontSize: 12}} label={{ value: 'AGIT (RPM)', angle: 90, position: 'insideRight', fill: 'var(--text-secondary)' }} />
                      <Tooltip content={<CustomTooltip />} />
                      <Legend wrapperStyle={{ fontSize: '14px', paddingTop: '10px' }} />
                      {visibleParameters.temp && <Line yAxisId="left" type="monotone" dataKey="temp_read" name="TEMP PV (°C)" stroke="var(--accent-red)" strokeWidth={3} dot={true} />}
                      {visibleParameters.temp && <Line yAxisId="left" type="monotone" dataKey="temp_set" name="TEMP SV (°C)" stroke="var(--accent-red)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.ph && <Line yAxisId="left" type="monotone" dataKey="ph_read" name="pH PV" stroke="var(--accent-blue)" strokeWidth={3} dot={true} />}
                      {visibleParameters.ph && <Line yAxisId="left" type="monotone" dataKey="ph_set" name="pH SV" stroke="var(--accent-blue)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.do && <Line yAxisId="left" type="monotone" dataKey="do_read" name="DO PV (%)" stroke="var(--accent-green)" strokeWidth={3} dot={true} />}
                      {visibleParameters.do && <Line yAxisId="left" type="monotone" dataKey="do_set" name="DO SV (%)" stroke="var(--accent-green)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.agit && <Line yAxisId="right" type="monotone" dataKey="agit_read" name="AGIT PV (RPM)" stroke="var(--accent-yellow)" strokeWidth={3} dot={true} />}
                      {visibleParameters.agit && <Line yAxisId="right" type="monotone" dataKey="agit_set" name="AGIT SV (RPM)" stroke="var(--accent-yellow)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.air && <Line yAxisId="left" type="monotone" dataKey="air_read" name="AIR PV (L/M)" stroke="var(--accent-purple)" strokeWidth={3} dot={true} />}
                      {visibleParameters.air && <Line yAxisId="left" type="monotone" dataKey="air_set" name="AIR SV (L/M)" stroke="var(--accent-purple)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              ) : (
                <div className="glass-panel" style={{ padding: '2rem' }}>
                  <div className="table-header-controls">
                    <h2 className="chart-title">Recorded Data</h2>
                    {userRole === 'admin' && (
                      <button 
                        className="delete-job-btn" 
                        onClick={clearAllData}
                        style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', background: 'rgba(239, 68, 68, 0.1)', padding: '8px 16px', borderRadius: '8px' }}
                      >
                        <Trash2 size={16} /> Clear All Records
                      </button>
                    )}
                  </div>
                  <div className="data-table-container">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'center', cursor: 'pointer' }} onClick={() => toggleSort('date')}>
                            Date {sortField === 'date' ? (sortAsc ? '▲' : '▼') : ''}
                          </th>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'center', cursor: 'pointer' }} onClick={() => toggleSort('time')}>
                            Time {sortField === 'time' ? (sortAsc ? '▲' : '▼') : ''}
                          </th>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'center', cursor: 'pointer' }} onClick={() => toggleSort('cultureHour')}>
                            ชั่วโมงที่ {sortField === 'cultureHour' ? (sortAsc ? '▲' : '▼') : ''}
                          </th>
                          <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-red)' }}>Temp (°C)</th>
                          <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-blue)' }}>pH</th>
                          <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-green)' }}>DO (%)</th>
                          <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-yellow)' }}>AGIT (RPM)</th>
                          <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-purple)' }}>AIR (L/M)</th>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'left', padding: '12px', minWidth: '150px' }}>Remarks / บันทึก</th>
                          {userRole === 'admin' && <th rowSpan="2" style={{ width: '80px', textAlign: 'center', verticalAlign: 'middle' }}>Action</th>}
                        </tr>
                        <tr>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                          <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                        </tr>
                      </thead>
                      <tbody>
                        {getSortedRows().map((row, index) => (
                          editingRowIndex === row.originalIndex ? (
                            <tr key={index} style={{ background: 'rgba(59, 130, 246, 0.05)' }}>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="date" 
                                  value={editingRowData.date} 
                                  onChange={(e) => handleEditChange('date', e.target.value)} 
                                  style={{ padding: '6px 4px', borderRadius: '4px', border: '1px solid var(--accent-blue)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '115px' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="time" 
                                  value={editingRowData.time} 
                                  onChange={(e) => handleEditChange('time', e.target.value)} 
                                  style={{ padding: '6px 4px', borderRadius: '4px', border: '1px solid var(--accent-blue)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '85px' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px', color: 'var(--text-secondary)' }}>
                                {getEditingRowCultureHour()} ชม.
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="0.01" 
                                  value={editingRowData.temp_set} 
                                  onChange={(e) => handleEditChange('temp_set', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '45px', textAlign: 'center' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="0.01" 
                                  value={editingRowData.temp_read} 
                                  onChange={(e) => handleEditChange('temp_read', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-red)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '45px', textAlign: 'center', fontWeight: 600 }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="0.01" 
                                  value={editingRowData.ph_set} 
                                  onChange={(e) => handleEditChange('ph_set', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '50px', textAlign: 'center' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="0.01" 
                                  value={editingRowData.ph_read} 
                                  onChange={(e) => handleEditChange('ph_read', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-blue)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '50px', textAlign: 'center', fontWeight: 600 }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="1" 
                                  value={editingRowData.do_set} 
                                  onChange={(e) => handleEditChange('do_set', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '40px', textAlign: 'center' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="1" 
                                  value={editingRowData.do_read} 
                                  onChange={(e) => handleEditChange('do_read', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-green)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '40px', textAlign: 'center', fontWeight: 600 }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="1" 
                                  value={editingRowData.agit_set} 
                                  onChange={(e) => handleEditChange('agit_set', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '50px', textAlign: 'center' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="1" 
                                  value={editingRowData.agit_read} 
                                  onChange={(e) => handleEditChange('agit_read', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-yellow)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '50px', textAlign: 'center', fontWeight: 600 }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="0.1" 
                                  value={editingRowData.air_set} 
                                  onChange={(e) => handleEditChange('air_set', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '45px', textAlign: 'center' }}
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input 
                                  type="number" 
                                  step="0.1" 
                                  value={editingRowData.air_read} 
                                  onChange={(e) => handleEditChange('air_read', parseFloat(e.target.value) || 0)} 
                                  style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-purple)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '45px', textAlign: 'center', fontWeight: 600 }}
                                />
                              </td>
                              <td style={{ textAlign: 'left', padding: '6px' }}>
                                <input 
                                  type="text" 
                                  value={editingRowData.remark} 
                                  onChange={(e) => handleEditChange('remark', e.target.value)} 
                                  style={{ padding: '6px 8px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'white', fontSize: '0.85rem', width: '100%', minWidth: '100px' }}
                                />
                              </td>
                              {userRole === 'admin' && (
                                <td style={{ textAlign: 'center', padding: '6px' }}>
                                  <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                    <button 
                                      className="save-row-btn" 
                                      onClick={saveEditRow}
                                      title="Save changes"
                                      style={{ background: 'transparent', border: 'none', color: 'var(--accent-green)', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                                    >
                                      <Check size={16} />
                                    </button>
                                    <button 
                                      className="cancel-row-btn" 
                                      onClick={cancelEditRow}
                                      title="Cancel editing"
                                      style={{ background: 'transparent', border: 'none', color: 'var(--accent-red)', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                                    >
                                      <X size={16} />
                                    </button>
                                  </div>
                                </td>
                              )}
                            </tr>
                          ) : (
                            <tr key={index}>
                              <td style={{ textAlign: 'center' }}>{row.date || (row.timestamp ? (new Date(row.timestamp).toISOString().slice(0,10)) : '')}</td>
                              <td style={{ textAlign: 'center' }}>{row.time || (row.timestamp ? (new Date(row.timestamp).toTimeString().slice(0,5)) : '')}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.cultureHour !== undefined ? `${row.cultureHour} ชม.` : '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.temp_set === 'number' ? row.temp_set.toFixed(1) : '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--accent-red)', fontWeight: 600 }}>{typeof row.temp_read === 'number' ? row.temp_read.toFixed(1) : '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.ph_set === 'number' ? row.ph_set.toFixed(2) : '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--accent-blue)', fontWeight: 600 }}>{typeof row.ph_read === 'number' ? row.ph_read.toFixed(2) : '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.do_set || '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--accent-green)', fontWeight: 600 }}>{row.do_read || '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.agit_set || '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--accent-yellow)', fontWeight: 600 }}>{row.agit_read || '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.air_set === 'number' ? row.air_set.toFixed(1) : '-'}</td>
                              <td style={{ textAlign: 'center', color: 'var(--accent-purple)', fontWeight: 600 }}>{typeof row.air_read === 'number' ? row.air_read.toFixed(1) : '-'}</td>
                              <td style={{ textAlign: 'left', padding: '12px', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>{row.remark || '-'}</td>
                              {userRole === 'admin' && (
                                <td style={{ textAlign: 'center' }}>
                                  <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                    <button 
                                      className="edit-row-btn" 
                                      onClick={() => startEditRow(row.originalIndex, row)}
                                      title="Edit this record"
                                      style={{ background: 'transparent', border: 'none', color: 'var(--accent-blue)', cursor: 'pointer', display: 'flex', alignItems: 'center', opacity: 0.7 }}
                                      onMouseEnter={e => e.currentTarget.style.opacity = 1}
                                      onMouseLeave={e => e.currentTarget.style.opacity = 0.7}
                                    >
                                      <Edit3 size={16} />
                                    </button>
                                    <button 
                                      className="delete-row-btn" 
                                      onClick={() => deleteDataPoint(row.originalIndex)}
                                      title="Delete this record"
                                      style={{ margin: '0' }}
                                    >
                                      <Trash2 size={16} />
                                    </button>
                                  </div>
                                </td>
                              )}
                            </tr>
                          )
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )
            )}
          </>
        )}
      </main>
    </div>
  );
}

export default App;
