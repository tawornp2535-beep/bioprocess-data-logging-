import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { initializeApp, cert } from 'firebase-admin/app';
import { getFirestore, FieldValue } from 'firebase-admin/firestore';

// Load environment variables
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DB_FILE = path.join(__dirname, 'db.json');

const app = express();
app.use(cors());
app.use(express.json());

// Helper: normalize time to HH:MM (24-hour)
const pad2 = (n) => String(n).padStart(2, '0');
const normalizeTimeHHMM = (input) => {
  if (!input) return null;
  if (typeof input === 'string') {
    const m = input.match(/(\d{1,2}):(\d{2})/);
    if (m) return `${pad2(parseInt(m[1], 10))}:${pad2(parseInt(m[2], 10))}`;
  }
  if (input instanceof Date) {
    return `${pad2(input.getHours())}:${pad2(input.getMinutes())}`;
  }
  return null;
};

// Helper: normalize date+time (or timestamp string) to ISO datetime
const normalizeDateTime = (dateInput, timeInput) => {
  // If a full timestamp string provided
  if (dateInput && typeof dateInput === 'string' && timeInput === undefined) {
    const d = new Date(dateInput);
    if (!isNaN(d.getTime())) return d.toISOString();
  }

  // If separate date and time provided
  if (dateInput && timeInput && typeof dateInput === 'string' && typeof timeInput === 'string') {
    const dt = new Date(`${dateInput}T${timeInput}`);
    if (!isNaN(dt.getTime())) return dt.toISOString();
    const dt2 = new Date(`${dateInput} ${timeInput}`);
    if (!isNaN(dt2.getTime())) return dt2.toISOString();
  }

  // If timeInput alone given (assume today)
  if (timeInput && typeof timeInput === 'string' && !dateInput) {
    const today = new Date();
    const parts = timeInput.match(/(\d{1,2}):(\d{2})/);
    if (parts) {
      const hh = parseInt(parts[1], 10);
      const mm = parseInt(parts[2], 10);
      const d = new Date(today.getFullYear(), today.getMonth(), today.getDate(), hh, mm);
      return d.toISOString();
    }
  }

  return null;
};

// ──────────────────────────────────────────────────────────
// Firebase Firestore Setup
// ──────────────────────────────────────────────────────────
let isCloud = false;
let db = null; // Firestore database reference

// Try to initialize Firebase Admin SDK
const SERVICE_ACCOUNT_PATH = process.env.FIREBASE_SERVICE_ACCOUNT_PATH || '';
const SERVICE_ACCOUNT_JSON = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || '';

if (SERVICE_ACCOUNT_JSON) {
  try {
    const serviceAccount = JSON.parse(SERVICE_ACCOUNT_JSON);
    initializeApp({
      credential: cert(serviceAccount)
    });
    db = getFirestore('default');
    isCloud = true;
    console.log('✅ Successfully connected to Firebase Firestore (Cloud Database) via Environment Variable.');
  } catch (err) {
    console.error('❌ Failed to connect to Firebase Firestore via Env. Falling back to local db.json. Error:', err.message);
  }
} else if (SERVICE_ACCOUNT_PATH && fs.existsSync(path.resolve(__dirname, SERVICE_ACCOUNT_PATH))) {
  try {
    const serviceAccount = JSON.parse(
      fs.readFileSync(path.resolve(__dirname, SERVICE_ACCOUNT_PATH), 'utf-8')
    );
    initializeApp({
      credential: cert(serviceAccount)
    });
    db = getFirestore('default');
    isCloud = true;
    console.log('✅ Successfully connected to Firebase Firestore (Cloud Database).');
  } catch (err) {
    console.error('❌ Failed to connect to Firebase Firestore. Falling back to local db.json. Error:', err.message);
  }
} else {
  console.log('ℹ️  No FIREBASE_SERVICE_ACCOUNT_PATH/JSON found or file missing. Using local db.json file storage.');
  if (SERVICE_ACCOUNT_PATH) {
    console.log(`   Looked for: ${path.resolve(__dirname, SERVICE_ACCOUNT_PATH)}`);
  }
}

// ──────────────────────────────────────────────────────────
// Firestore Helper Functions
// ──────────────────────────────────────────────────────────

// Collection references
const MACHINES_COL = 'machines';
const JOBS_COL = 'jobs';
const CUSTOMERS_COL = 'customers';

// Get all documents from a Firestore collection
const getCollection = async (collectionName) => {
  const snapshot = await db.collection(collectionName).get();
  return snapshot.docs.map(doc => doc.data());
};

// ──────────────────────────────────────────────────────────
// Local JSON DB Helpers (fallback)
// ──────────────────────────────────────────────────────────
const readLocalDB = () => {
  if (!fs.existsSync(DB_FILE)) {
    const defaultData = {
      machines: [{ id: 'm1', name: 'Bioreactor 1' }],
      jobs: [
        {
          id: 'job-initial',
          machineId: 'm1',
          name: 'Default Session',
          createdAt: new Date().toLocaleString(),
          data: []
        }
      ],
      customers: []
    };
    fs.writeFileSync(DB_FILE, JSON.stringify(defaultData, null, 2), 'utf-8');
    return defaultData;
  }
  try {
    const content = fs.readFileSync(DB_FILE, 'utf-8');
    return JSON.parse(content);
  } catch (e) {
    console.error('Error reading db.json, returning empty template:', e);
    return { machines: [], jobs: [], customers: [] };
  }
};

const writeLocalDB = (data) => {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing db.json:', e);
  }
};

// ──────────────────────────────────────────────────────────
// Unified Data Fetch
// ──────────────────────────────────────────────────────────
const getDB = async () => {
  if (isCloud) {
    try {
      let machines = await getCollection(MACHINES_COL);
      let jobs = await getCollection(JOBS_COL);
      let customers = await getCollection(CUSTOMERS_COL);

      // Seed cloud database if empty
      if (machines.length === 0) {
        const defaultMachine = { id: 'm1', name: 'Bioreactor 1' };
        const defaultJob = {
          id: 'job-initial',
          machineId: 'm1',
          name: 'Default Session',
          createdAt: new Date().toLocaleString(),
          data: []
        };
        await db.collection(MACHINES_COL).doc('m1').set(defaultMachine);
        await db.collection(JOBS_COL).doc('job-initial').set(defaultJob);

        machines = [defaultMachine];
        jobs = [defaultJob];
      }
      return { machines, jobs, customers };
    } catch (e) {
      console.error('Error reading Firestore, falling back to local file:', e);
      return readLocalDB();
    }
  } else {
    return readLocalDB();
  }
};

// ──────────────────────────────────────────────────────────
// Settings Helpers
// ──────────────────────────────────────────────────────────
const SETTINGS_COL = 'settings';

const getSettings = async () => {
  if (isCloud) {
    try {
      const doc = await db.collection(SETTINGS_COL).doc('system').get();
      if (doc.exists) {
        return doc.data();
      }
    } catch (e) {
      console.error('Error reading settings from Firestore:', e);
    }
  }
  
  const localDB = readLocalDB();
  if (!localDB.settings) {
    localDB.settings = { adminPassword: 'admin123' };
    writeLocalDB(localDB);
  }
  
  if (isCloud) {
    try {
      await db.collection(SETTINGS_COL).doc('system').set(localDB.settings);
    } catch (e) {
      console.error('Failed to seed settings in Firestore:', e);
    }
  }
  return localDB.settings;
};

const saveSettings = async (settings) => {
  if (isCloud) {
    try {
      await db.collection(SETTINGS_COL).doc('system').set(settings);
    } catch (e) {
      console.error('Error writing settings to Firestore:', e);
    }
  }
  const localDB = readLocalDB();
  localDB.settings = settings;
  writeLocalDB(localDB);
};

// ──────────────────────────────────────────────────────────
// API Routes
// ──────────────────────────────────────────────────────────

// In-memory active users tracker
const activeUsers = {};

app.get('/api/db', async (req, res) => {
  const { clientId, role, machineId, jobId } = req.query;
  
  if (clientId) {
    activeUsers[clientId] = {
      role: role || 'guest',
      machineId: machineId || '',
      jobId: jobId || '',
      lastActive: Date.now()
    };
  }

  // Clean up users who haven't sent a heartbeat/poll in 12 seconds
  const now = Date.now();
  for (const cid in activeUsers) {
    if (now - activeUsers[cid].lastActive > 12000) {
      delete activeUsers[cid];
    }
  }

  const activeUsersList = Object.entries(activeUsers).map(([cid, u]) => ({
    clientId: cid,
    role: u.role,
    machineId: u.machineId,
    jobId: u.jobId
  }));

  const dbData = await getDB();
  res.json({
    ...dbData,
    activeUsers: activeUsersList
  });
});

// ── Settings ─────────────────────────────────────────────

app.post('/api/settings/verify-password', async (req, res) => {
  const { password } = req.body;
  const settings = await getSettings();
  if (password === settings.adminPassword) {
    return res.json({ success: true });
  } else {
    return res.status(400).json({ error: 'Incorrect password' });
  }
});

app.post('/api/settings/update-password', async (req, res) => {
  const { currentPassword, newPassword } = req.body;
  const settings = await getSettings();
  if (currentPassword !== settings.adminPassword) {
    return res.status(400).json({ error: 'รหัสผ่านปัจจุบันไม่ถูกต้อง' });
  }
  if (!newPassword || newPassword.trim() === '') {
    return res.status(400).json({ error: 'รหัสผ่านใหม่ห้ามเป็นค่าว่าง' });
  }
  settings.adminPassword = newPassword.trim();
  await saveSettings(settings);
  res.json({ success: true });
});

// ── Machines ─────────────────────────────────────────────

app.post('/api/machines', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }

  const id = 'm-' + Date.now();
  const newMachine = { id, name: name.trim() };

  if (isCloud) {
    try {
      await db.collection(MACHINES_COL).doc(id).set(newMachine);
    } catch (e) {
      console.error('Firestore error creating machine:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.machines.push(newMachine);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

app.put('/api/machines/:id', async (req, res) => {
  const { id } = req.params;
  const { name } = req.body;
  if (!name || !name.trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }

  if (isCloud) {
    try {
      await db.collection(MACHINES_COL).doc(id).update({ name: name.trim() });
    } catch (e) {
      console.error('Firestore error renaming machine:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.machines = localDB.machines.map(m => m.id === id ? { ...m, name: name.trim() } : m);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

app.delete('/api/machines/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      // Delete machine
      await db.collection(MACHINES_COL).doc(id).delete();
      // Delete all jobs for this machine
      const jobsSnapshot = await db.collection(JOBS_COL).where('machineId', '==', id).get();
      const batch1 = db.batch();
      jobsSnapshot.docs.forEach(doc => batch1.delete(doc.ref));
      await batch1.commit();
      // Delete all customers for this machine
      const customersSnapshot = await db.collection(CUSTOMERS_COL).where('machineId', '==', id).get();
      const batch2 = db.batch();
      customersSnapshot.docs.forEach(doc => batch2.delete(doc.ref));
      await batch2.commit();
    } catch (e) {
      console.error('Firestore error deleting machine:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.machines = localDB.machines.filter(m => m.id !== id);
    localDB.jobs = localDB.jobs.filter(j => j.machineId !== id);
    localDB.customers = localDB.customers.filter(c => c.machineId !== id);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

// ── Jobs (Sessions) ──────────────────────────────────────

app.post('/api/jobs', async (req, res) => {
  const { machineId, name } = req.body;
  if (!machineId || !name || !name.trim()) {
    return res.status(400).json({ error: 'Machine ID and name are required' });
  }

  const newJob = {
    id: 'job-' + Date.now(),
    machineId,
    name: name.trim(),
    createdAt: new Date().toLocaleString(),
    data: []
  };

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(newJob.id).set(newJob);
    } catch (e) {
      console.error('Firestore error creating session:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.jobs.push(newJob);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

app.delete('/api/jobs/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).delete();
    } catch (e) {
      console.error('Firestore error deleting session:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.jobs = localDB.jobs.filter(j => j.id !== id);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

app.put('/api/jobs/:id/expiry', async (req, res) => {
  const { id } = req.params;
  const { expiresAt } = req.body;

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).update({ expiresAt });
    } catch (e) {
      console.error('Firestore error updating session expiry:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.expiresAt = expiresAt;
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

// ── Data Points ──────────────────────────────────────────

app.post('/api/jobs/:id/data', async (req, res) => {
  const { id } = req.params;
  const {
    temp_set, temp_read,
    ph_set, ph_read,
    do_set, do_read,
    agit_set, agit_read,
    air_set, air_read,
    remark
  } = req.body;

  console.log(`POST /api/jobs/${id}/data payload:`, req.body);

  const newDataPoint = {
    timestamp: req.body.timestamp || normalizeDateTime(req.body.date, req.body.time) || normalizeDateTime(undefined, req.body.time) || new Date().toISOString(),
    date: null,
    time: normalizeTimeHHMM(req.body.time) || normalizeTimeHHMM(new Date()),
    temp_set: parseFloat(temp_set) || 0,
    temp_read: parseFloat(temp_read) || 0,
    ph_set: parseFloat(ph_set) || 0,
    ph_read: parseFloat(ph_read) || 0,
    do_set: parseFloat(do_set) || 0,
    do_read: parseFloat(do_read) || 0,
    agit_set: parseFloat(agit_set) || 0,
    agit_read: parseFloat(agit_read) || 0,
    air_set: parseFloat(air_set) || 0,
    air_read: parseFloat(air_read) || 0,
    remark: remark || ''
  };

  // Derive date from timestamp
  try {
    const dt = new Date(newDataPoint.timestamp);
    if (!isNaN(dt.getTime())) {
      newDataPoint.date = dt.toISOString().slice(0, 10);
    } else {
      newDataPoint.date = newDataPoint.timestamp ? String(newDataPoint.timestamp).slice(0, 10) : (new Date().toISOString().slice(0, 10));
    }
  } catch (e) {
    newDataPoint.date = new Date().toISOString().slice(0, 10);
  }

  if (isCloud) {
    try {
      // Use FieldValue.arrayUnion to add data point to the job's data array
      await db.collection(JOBS_COL).doc(id).update({
        data: FieldValue.arrayUnion(newDataPoint)
      });
    } catch (e) {
      console.error('Firestore error pushing data point:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.data.push(newDataPoint);
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

app.delete('/api/jobs/:id/data/:index', async (req, res) => {
  const { id, index } = req.params;
  const idx = parseInt(index, 10);

  if (isNaN(idx) || idx < 0) {
    return res.status(400).json({ error: 'Invalid index' });
  }

  if (isCloud) {
    try {
      const jobDoc = await db.collection(JOBS_COL).doc(id).get();
      if (jobDoc.exists) {
        const jobData = jobDoc.data();
        if (jobData.data && idx < jobData.data.length) {
          jobData.data.splice(idx, 1);
          await db.collection(JOBS_COL).doc(id).update({ data: jobData.data });
        }
      }
    } catch (e) {
      console.error('Firestore error deleting data point:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job && idx < job.data.length) {
      job.data.splice(idx, 1);
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

app.put('/api/jobs/:id/data/:index', async (req, res) => {
  const { id, index } = req.params;
  const idx = parseInt(index, 10);
  const updatedPoint = req.body;

  if (isNaN(idx) || idx < 0) {
    return res.status(400).json({ error: 'Invalid index' });
  }

  const cleanPoint = {
    timestamp: updatedPoint.timestamp || normalizeDateTime(updatedPoint.date, updatedPoint.time) || new Date().toISOString(),
    date: updatedPoint.date || new Date().toISOString().slice(0, 10),
    time: normalizeTimeHHMM(updatedPoint.time) || normalizeTimeHHMM(new Date()),
    temp_set: parseFloat(updatedPoint.temp_set) || 0,
    temp_read: parseFloat(updatedPoint.temp_read) || 0,
    ph_set: parseFloat(updatedPoint.ph_set) || 0,
    ph_read: parseFloat(updatedPoint.ph_read) || 0,
    do_set: parseFloat(updatedPoint.do_set) || 0,
    do_read: parseFloat(updatedPoint.do_read) || 0,
    agit_set: parseFloat(updatedPoint.agit_set) || 0,
    agit_read: parseFloat(updatedPoint.agit_read) || 0,
    air_set: parseFloat(updatedPoint.air_set) || 0,
    air_read: parseFloat(updatedPoint.air_read) || 0,
    remark: updatedPoint.remark || ''
  };

  if (isCloud) {
    try {
      const jobDoc = await db.collection(JOBS_COL).doc(id).get();
      if (jobDoc.exists) {
        const jobData = jobDoc.data();
        if (jobData.data && idx < jobData.data.length) {
          jobData.data[idx] = cleanPoint;
          await db.collection(JOBS_COL).doc(id).update({ data: jobData.data });
        }
      }
    } catch (e) {
      console.error('Firestore error updating data point:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job && idx < job.data.length) {
      job.data[idx] = cleanPoint;
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

app.post('/api/jobs/:id/clear', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).update({ data: [] });
    } catch (e) {
      console.error('Firestore error clearing data:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.data = [];
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

// ── Customers ────────────────────────────────────────────

app.post('/api/customers', async (req, res) => {
  const { companyName, machineId, email } = req.body;
  if (!companyName || !companyName.trim()) {
    return res.status(400).json({ error: 'Company name is required' });
  }

  const newCustomer = {
    id: 'cust-' + Date.now(),
    companyName: companyName.trim(),
    machineId,
    email: email && typeof email === 'string' && email.trim() !== '' ? email.trim() : undefined,
    createdAt: new Date().toLocaleDateString()
  };

  if (isCloud) {
    try {
      // Remove undefined fields for Firestore (Firestore doesn't accept undefined)
      const cleanCustomer = JSON.parse(JSON.stringify(newCustomer));
      await db.collection(CUSTOMERS_COL).doc(newCustomer.id).set(cleanCustomer);
    } catch (e) {
      console.error('Firestore error creating customer:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.customers.push(newCustomer);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

app.delete('/api/customers/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await db.collection(CUSTOMERS_COL).doc(id).delete();
    } catch (e) {
      console.error('Firestore error deleting customer:', e);
    }
  } else {
    const localDB = readLocalDB();
    localDB.customers = localDB.customers.filter(c => c.id !== id);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

// ──────────────────────────────────────────────────────────
// Serve static frontend files from 'dist' directory
// ──────────────────────────────────────────────────────────
const distPath = path.join(__dirname, 'dist');
app.use(express.static(distPath));

// Fallback for SPA routing: send index.html for non-API requests
app.get('/*splat', (req, res, next) => {
  if (req.path.startsWith('/api')) {
    return next();
  }
  const indexHtml = path.join(distPath, 'index.html');
  if (fs.existsSync(indexHtml)) {
    res.sendFile(indexHtml);
  } else {
    res.status(404).send('Frontend build (dist/) not found. Please run "npm run build" first.');
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n🚀 Backend server running on http://localhost:${PORT}`);
  console.log(`   Storage: ${isCloud ? '☁️  Firebase Firestore (Cloud)' : '💾 Local db.json'}`);
  console.log('');
});
