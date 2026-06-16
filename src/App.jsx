import React, { useState, useEffect } from 'react';
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from 'recharts';
import { 
  Activity, Droplets, Wind, Thermometer, RotateCw, PlusCircle, 
  Download, LayoutDashboard, LineChart as ChartIcon, FolderPlus, Trash2, FolderOpen,
  Table as TableIcon, Users, Activity as ActivityIcon, Edit3
} from 'lucide-react';
import './index.css';
import './form.css';
import './tabs.css';
import './table.css';

const CustomTooltip = ({ active, payload, label }) => {
  if (active && payload && payload.length) {
    return (
      <div className="custom-tooltip">
        <p className="custom-tooltip-label">{label}</p>
        {payload.map((entry, index) => (
          <p key={index} style={{ color: entry.color, margin: 0, fontSize: '14px', fontWeight: 600 }}>
            {entry.name}: {entry.value.toFixed(2)}
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
    return sessionStorage.getItem('bioprocess-role') || null; // 'admin' | 'customer' | null
  });
  const [activeCustomerJobId, setActiveCustomerJobId] = useState(() => {
    return sessionStorage.getItem('bioprocess-customer-job-id') || null;
  });

  useEffect(() => {
    if (userRole) {
      sessionStorage.setItem('bioprocess-role', userRole);
    } else {
      sessionStorage.removeItem('bioprocess-role');
    }
  }, [userRole]);

  useEffect(() => {
    if (activeCustomerJobId) {
      sessionStorage.setItem('bioprocess-customer-job-id', activeCustomerJobId);
    } else {
      sessionStorage.removeItem('bioprocess-customer-job-id');
    }
  }, [activeCustomerJobId]);

  // Global View State
  const [currentAppView, setCurrentAppView] = useState('monitoring'); // 'monitoring' | 'customers'

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
    date: '',
    time: ''
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

  // Fetch Database from Backend Helper
  const fetchDB = async (shouldAutoSelect = false) => {
    try {
      const res = await fetch('/api/db');
      if (res.ok) {
        const data = await res.json();
        setMachines(data.machines);
        setJobs(data.jobs);
        setCustomers(data.customers);
        
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

  useEffect(() => {
    fetchDB(true);
  }, []);

  // Poll server for real-time synchronization across multiple users
  useEffect(() => {
    const interval = setInterval(() => {
      fetchDB(false);
    }, 5000);
    return () => clearInterval(interval);
  }, []);

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
  const chartData = displayData.map(row => ({
    ...row,
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
  }));
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
      const aVal = a[sortField] || a.time || '';
      const bVal = b[sortField] || b.time || '';
      const aDate = new Date(aVal);
      const bDate = new Date(bVal);
      if (!isNaN(aDate.getTime()) && !isNaN(bDate.getTime())) {
        return sortAsc ? aDate - bDate : bDate - aDate;
      }
      // fallback string compare
      if (aVal < bVal) return sortAsc ? -1 : 1;
      if (aVal > bVal) return sortAsc ? 1 : -1;
      return 0;
    });
    return rows;
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
        date: lastDateValid ? lastDate.toISOString().slice(0,10) : new Date().toISOString().slice(0,10),
        time: toHHMM(lastDataPoint.time) || (lastDateValid ? toHHMM(lastDate) : toHHMM(new Date()))
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
        if (isCloud) {
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
    const curDate = now.toISOString().slice(0,10);
    const curTime = toHHMM(now);
    addDataPoint(
      formData.temp_set, formData.temp_read,
      formData.ph_set, formData.ph_read,
      formData.do_set, formData.do_read,
      formData.agit_set, formData.agit_read,
      formData.air_set, formData.air_read,
      formData.remark,
      curTime,
      curDate
    );
    // Reset remark field
    setFormData(prev => ({ ...prev, remark: '' }));
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
          date
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
      'Timestamp', 
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
      const tstamp = row.timestamp || row.time || '';
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
        `"${tstamp}"`, 
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
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const safeName = currentJob.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    link.setAttribute('download', `${safeName}_data.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Auto Simulation removed

  if (!userRole) {
    return (
      <div className="login-overlay">
        <div className="glass-panel login-card" style={{ maxWidth: '450px', width: '95%', padding: '2.5rem', margin: 'auto' }}>
          <h1 style={{ textAlign: 'center', marginBottom: '0.5rem', background: 'linear-gradient(to right, var(--accent-blue), var(--accent-purple))', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', fontWeight: 700 }}>Bioprocess System</h1>
          <p style={{ textAlign: 'center', color: 'var(--text-secondary)', fontSize: '0.9rem', marginBottom: '2rem' }}>เครื่องมือบันทึกและจัดเก็บข้อมูล</p>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            {/* Admin Login Block */}
            <div style={{ padding: '1.25rem', border: '1px solid var(--border-color)', borderRadius: '12px', background: 'rgba(15, 23, 42, 0.3)' }}>
              <h3 style={{ marginBottom: '1rem', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '1.05rem' }}>🔐 สำหรับแอดมิน (Admin)</h3>
              <form onSubmit={(e) => {
                e.preventDefault();
                const password = e.target.adminPassword.value;
                if (password === 'admin123') { // Simple default password
                  setUserRole('admin');
                } else {
                  alert('รหัสผ่านไม่ถูกต้อง (รหัสเริ่มต้น: admin123)');
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
      {/* SIDEBAR */}
      <aside className="sidebar">
        <div className="sidebar-header" style={{ paddingBottom: '1rem' }}>
          <h1 style={{ marginBottom: '1.5rem' }}>Bioprocess</h1>
          
          {/* Main App Navigation (Admin only) */}
          {userRole === 'admin' ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginBottom: '1rem' }}>
              <button 
                className={`nav-tab ${currentAppView === 'monitoring' ? 'active' : ''}`}
                onClick={() => setCurrentAppView('monitoring')}
                style={{ justifyContent: 'flex-start', padding: '10px 16px' }}
              >
                <ActivityIcon size={18} /> Monitoring & Logging
              </button>
              <button 
                className={`nav-tab ${currentAppView === 'customers' ? 'active' : ''}`}
                onClick={() => setCurrentAppView('customers')}
                style={{ justifyContent: 'flex-start', padding: '10px 16px' }}
              >
                <Users size={18} /> Customer Data
              </button>
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.25rem', marginBottom: '0.5rem' }}>
              <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>👤 Customer Portal</span>
              <span style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--accent-blue)', wordBreak: 'break-all' }}>
                ID: {activeCustomerJobId}
              </span>
              {showCustomerBanner && (
                <div style={{ marginTop: '10px', padding: '10px', borderRadius: '8px', background: 'linear-gradient(90deg, rgba(255,244,229,0.06), rgba(255,250,240,0.03))', borderLeft: '4px solid var(--accent-yellow)', color: 'var(--text-primary)', fontSize: '0.9rem', lineHeight: 1.3 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '8px' }}>
                    <div>
                      <div style={{ fontWeight: 700, marginBottom: '6px' }}>⚠️ ข้อความแจ้งเตือนการเข้าถึงและบันทึกข้อมูล</div>
                      <div style={{ color: 'var(--text-secondary)' }}>• สิทธิ์การเข้าถึง: รหัสผ่านนี้สามารถใช้เข้าดูได้เฉพาะข้อมูลในส่วนงานของลูกค้าที่ได้รับอนุญาตเท่านั้น</div>
                      <div style={{ color: 'var(--text-secondary)' }}>• การรักษาความลับ: ข้อมูลทั้งหมดจะถูกจัดเก็บและรักษาไว้เป็นความลับสูงสุด</div>
                      <div style={{ color: 'var(--text-secondary)' }}>• การสำรองข้อมูล: ระบบจะทำการสำรองข้อมูลไว้เป็นเวลา 7 วัน หลังจากเสร็จสิ้นงาน</div>
                    </div>
                    <button onClick={() => setShowCustomerBanner(false)} style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', fontSize: '16px' }}>✕</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        
        {/* Context Menu: Only show if in monitoring view */}
        {currentAppView === 'monitoring' && (
          <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
            {userRole === 'admin' ? (
              <div style={{ padding: '0 1.5rem 1rem 1.5rem', borderBottom: '1px solid var(--border-color)' }}>
                <div className="machine-selector" style={{ border: 'none', padding: 0, margin: 0 }}>
                  <label style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>
                    Select Instrument / Machine
                  </label>
                  <div style={{ display: 'flex', gap: '0.5rem', marginBottom: '1rem' }}>
                    <select 
                      className="machine-dropdown" 
                      value={currentMachineId} 
                      onChange={handleMachineChange}
                      style={{ margin: 0, flex: 1 }}
                    >
                      {machines.map(m => (
                        <option key={m.id} value={m.id}>{m.name}</option>
                      ))}
                      <option value="ADD_NEW">+ Add New Machine...</option>
                    </select>
                    <button 
                      onClick={renameMachine} 
                      title="Rename selected machine"
                      style={{ 
                        background: 'rgba(255,255,255,0.05)', 
                        border: '1px solid var(--border-color)', 
                        borderRadius: '8px', 
                        color: 'var(--text-primary)', 
                        padding: '8px', 
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        transition: 'all 0.2s'
                      }}
                    >
                      <Edit3 size={16} />
                    </button>
                    <button 
                      onClick={deleteMachine} 
                      title="Delete selected machine"
                      style={{ 
                        background: 'rgba(239, 68, 68, 0.1)', 
                        border: '1px solid rgba(239, 68, 68, 0.2)', 
                        borderRadius: '8px', 
                        color: 'var(--accent-red)', 
                        padding: '8px', 
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        transition: 'all 0.2s'
                      }}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                </div>
                <button className="new-job-btn" onClick={createNewJob}>
                  <FolderPlus size={18} />
                  New Session
                </button>
              </div>
            ) : (
              // Customer Read-Only Machine/Session Card
              <div style={{ padding: '1rem 1.5rem', borderBottom: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>เครื่องมือที่ใช้ (Instrument)</span>
                <span style={{ fontSize: '1rem', fontWeight: 600, color: 'var(--text-primary)' }}>{currentMachine?.name}</span>
                <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', textTransform: 'uppercase', letterSpacing: '0.05em', marginTop: '0.5rem' }}>เซสชันการทดสอบ (Session)</span>
                <span style={{ fontSize: '1rem', fontWeight: 600, color: 'var(--text-primary)' }}>{currentJob?.name}</span>
              </div>
            )}
            
            {userRole === 'admin' && (
              <div className="job-list">
                {jobsForMachine.length === 0 ? (
                  <div style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>
                    No sessions. Create one above.
                  </div>
                ) : (
                  jobsForMachine.map((job) => (
                    <div 
                      key={job.id} 
                      className={`job-item ${job.id === currentJobId ? 'active' : ''}`}
                      onClick={() => {
                          setCurrentJobId(job.id);
                        }}
                    >
                      <div className="job-name">{job.name}</div>
                      <div className="job-date">{job.createdAt}</div>
                      
                      {/* Copy code control for Admins */}
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '0.5rem', borderTop: '1px solid rgba(255,255,255,0.05)', paddingTop: '0.5rem' }}>
                        <span style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>
                          Code: <code style={{ color: 'var(--accent-blue)', background: 'rgba(59, 130, 246, 0.1)', padding: '2px 4px', borderRadius: '4px', cursor: 'pointer' }} onClick={(e) => {
                            e.stopPropagation();
                            navigator.clipboard.writeText(job.id);
                            alert(`คัดลอกรหัสงานเรียบร้อย: ${job.id}`);
                          }} title="Click to copy Job ID">
                            {job.id.substring(0, 10)}...
                          </code>
                        </span>
                        <button className="delete-job-btn" style={{ margin: 0 }} onClick={(e) => deleteJob(job.id, e)} title="Delete Session">
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
        )}

        {/* Logout Button */}
        <div style={{ padding: '1.5rem', borderTop: '1px solid var(--border-color)' }}>
          <button 
            onClick={() => {
              if (window.confirm("ยืนยันการออกจากระบบ?")) {
                setUserRole(null);
                setActiveCustomerJobId(null);
              }
            }} 
            style={{ 
              width: '100%', 
              padding: '10px', 
              borderRadius: '8px', 
              border: '1px solid rgba(239, 68, 68, 0.2)', 
              background: 'rgba(239, 68, 68, 0.1)', 
              color: 'var(--accent-red)', 
              fontWeight: 600, 
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '0.5rem',
              transition: 'all 0.2s'
            }}
          >
            ออกจากระบบ (Logout)
          </button>
        </div>
      </aside>

      {/* MAIN CONTENT */}
      <main className="main-content">
        {currentAppView === 'customers' ? (
          /* CUSTOMER DATA VIEW */
          <div className="customers-view">
            <header className="dashboard-header">
              <h2>Customer Database</h2>
            </header>

            <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1rem' }}>
              <h2 className="form-title">Add New Customer</h2>
              <form onSubmit={handleAddCustomer} className="data-form" style={{ display: 'flex', gap: '1rem', alignItems: 'flex-end' }}>
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
                      <th>Date Added</th>
                      <th style={{ width: '80px', textAlign: 'center' }}>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {customers.length === 0 ? (
                      <tr>
                        <td colSpan="5" style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>
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

              <div className="header-controls">
                <button
                  className={`theme-toggle-btn`}
                  onClick={() => setTheme(prev => prev === 'dark' ? 'light' : 'dark')}
                  title="Toggle theme"
                  style={{ marginRight: '12px' }}
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
                    style={{ marginRight: '8px' }}
                  >
                    {isReplay ? 'Replay ON' : 'Replay'}
                  </button>
                )}
                <button className="export-btn" onClick={exportToCSV} disabled={!currentJob}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export CSV
                </button>
                {/* Auto simulation removed */}
              </div>
            </header>

            {currentJob && userRole === 'admin' && (
              /* Manual Input Form */
              <div className="glass-panel form-container">
                <h2 className="form-title">Manual Data Entry (กรอกค่าทดสอบของ {currentMachine?.name || 'เครื่องมือ'})</h2>
                <form onSubmit={handleManualSubmit} className="data-form">
                  <div className="form-group-container">
                    <div className="form-group">
                      <label>TEMP (°C)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.1" name="temp_set" value={formData.temp_set} onChange={handleInputChange} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.1" name="temp_read" value={formData.temp_read} onChange={handleInputChange} />
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
                          {(lastDataPoint.temp_read !== undefined ? lastDataPoint.temp_read : lastDataPoint.temp).toFixed(1)}<span className="metric-unit">°C</span>
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {(lastDataPoint.temp_set !== undefined ? lastDataPoint.temp_set : lastDataPoint.temp).toFixed(1)}°C
                        </div>
                      </div>
                    </div>
                    <div className="glass-panel metric-card">
                      <Droplets color="var(--accent-blue)" size={24} style={{ marginBottom: '10px' }} />
                      <div className="metric-title">pH Level</div>
                      <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                        <div className="metric-value" style={{ color: 'var(--accent-blue)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                          {(lastDataPoint.ph_read !== undefined ? lastDataPoint.ph_read : lastDataPoint.ph).toFixed(2)}
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {(lastDataPoint.ph_set !== undefined ? lastDataPoint.ph_set : lastDataPoint.ph).toFixed(2)}
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
                          {(lastDataPoint.air_read !== undefined ? lastDataPoint.air_read : lastDataPoint.air).toFixed(1)}<span className="metric-unit">L/M</span>
                        </div>
                        <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          SV (Set): {(lastDataPoint.air_set !== undefined ? lastDataPoint.air_set : lastDataPoint.air).toFixed(1)} L/M
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
                          <YAxis domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{fontSize: 12}} />
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
                          <tr key={index}>
                            <td style={{ textAlign: 'center' }}>{row.date || (row.timestamp ? (new Date(row.timestamp).toISOString().slice(0,10)) : '')}</td>
                            <td style={{ textAlign: 'center' }}>{row.time || (row.timestamp ? (new Date(row.timestamp).toTimeString().slice(0,5)) : '')}</td>
                            <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.temp_set.toFixed(1)}</td>
                            <td style={{ textAlign: 'center', color: 'var(--accent-red)', fontWeight: 600 }}>{row.temp_read.toFixed(1)}</td>
                            <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.ph_set.toFixed(2)}</td>
                            <td style={{ textAlign: 'center', color: 'var(--accent-blue)', fontWeight: 600 }}>{row.ph_read.toFixed(2)}</td>
                            <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.do_set}</td>
                            <td style={{ textAlign: 'center', color: 'var(--accent-green)', fontWeight: 600 }}>{row.do_read}</td>
                            <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.agit_set}</td>
                            <td style={{ textAlign: 'center', color: 'var(--accent-yellow)', fontWeight: 600 }}>{row.agit_read}</td>
                            <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.air_set.toFixed(1)}</td>
                            <td style={{ textAlign: 'center', color: 'var(--accent-purple)', fontWeight: 600 }}>{row.air_read.toFixed(1)}</td>
                            <td style={{ textAlign: 'left', padding: '12px', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>{row.remark || '-'}</td>
                            {userRole === 'admin' && (
                              <td style={{ textAlign: 'center' }}>
                                <button 
                                  className="delete-row-btn" 
                                  onClick={() => deleteDataPoint(index)}
                                  title="Delete this record"
                                  style={{ margin: '0 auto' }}
                                >
                                  <Trash2 size={16} />
                                </button>
                              </td>
                            )}
                          </tr>
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
