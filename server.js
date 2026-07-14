import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { initializeApp, cert } from 'firebase-admin/app';
import { getFirestore, FieldValue } from 'firebase-admin/firestore';
import { getStorage } from 'firebase-admin/storage';

// Load environment variables
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DB_FILE = path.join(__dirname, 'db.json');
const UPLOADS_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));
app.use('/uploads', express.static(UPLOADS_DIR));

// Keep track of recent requests in memory to check pings/uptime robots
const recentRequests = [];

// Request logging middleware
app.use((req, res, next) => {
  // Invalidate database cache on write operations (POST, PUT, DELETE, PATCH)
  if (req.method !== 'GET') {
    invalidateCache();
  }

  const userAgent = req.headers['user-agent'] || 'Unknown';
  recentRequests.push({
    time: new Date().toISOString(),
    method: req.method,
    url: req.url,
    userAgent,
    ip: req.headers['x-forwarded-for'] || req.ip || req.socket.remoteAddress
  });
  if (recentRequests.length > 40) {
    recentRequests.shift();
  }

  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  if (req.body && Object.keys(req.body).length) {
    // Avoid logging large file uploads/image uploads if any
    const bodyLog = { ...req.body };
    if (bodyLog.imageData) bodyLog.imageData = '[BASE64_IMAGE_DATA]';
    console.log(`  Body:`, JSON.stringify(bodyLog));
  }
  next();
});


// Helper: normalize time to HH:MM (24-hour)
const pad2 = (n) => String(n).padStart(2, '0');
const normalizeTimeHHMM = (input) => {
  if (!input) return null;
  if (typeof input === 'string') {
    const m = input.match(/(\d{1,2}):(\d{2})/);
    if (m) return `${pad2(parseInt(m[1], 10))}:${pad2(parseInt(m[2], 10))}`;
  }
  if (input instanceof Date) {
    // Offset by +7 hours to get ICT time
    const ictTime = new Date(input.getTime() + 7 * 60 * 60 * 1000);
    return `${pad2(ictTime.getUTCHours())}:${pad2(ictTime.getUTCMinutes())}`;
  }
  return null;
};

// Helper to parse date (YYYY-MM-DD) and time (HH:MM) in ICT (UTC+7) timezone
const parseICTDateTime = (dateStr, timeStr) => {
  let year, month, day, hh = 0, mm = 0;

  if (dateStr) {
    const dParts = dateStr.split('-');
    if (dParts.length === 3) {
      year = parseInt(dParts[0], 10);
      month = parseInt(dParts[1], 10) - 1;
      day = parseInt(dParts[2], 10);
    }
  }

  // If no valid date was parsed, use "today" in ICT
  if (year === undefined) {
    const ictNow = new Date(Date.now() + 7 * 60 * 60 * 1000);
    year = ictNow.getUTCFullYear();
    month = ictNow.getUTCMonth();
    day = ictNow.getUTCDate();
  }

  if (timeStr) {
    const tParts = timeStr.match(/(\d{1,2}):(\d{2})/);
    if (tParts) {
      hh = parseInt(tParts[1], 10);
      mm = parseInt(tParts[2], 10);
    }
  }

  const utcMs = Date.UTC(year, month, day, hh, mm, 0) - (7 * 60 * 60 * 1000);
  return new Date(utcMs);
};

// Helper: normalize date+time (or timestamp string) to ISO datetime
const normalizeDateTime = (dateInput, timeInput) => {
  // If a full timestamp string provided
  if (dateInput && typeof dateInput === 'string' && timeInput === undefined) {
    if (dateInput.includes('T') && (dateInput.endsWith('Z') || dateInput.includes('+') || dateInput.includes('-'))) {
      const d = new Date(dateInput);
      if (!isNaN(d.getTime())) return d.toISOString();
    }
    const d = parseICTDateTime(dateInput, undefined);
    if (!isNaN(d.getTime())) return d.toISOString();
  }

  // If separate date and time provided (or time alone)
  if (dateInput || timeInput) {
    const d = parseICTDateTime(dateInput, timeInput);
    if (!isNaN(d.getTime())) return d.toISOString();
  }

  return null;
};

// ──────────────────────────────────────────────────────────
// Firebase Firestore Setup
// ──────────────────────────────────────────────────────────
let isCloud = false;
let db = null; // Firestore database reference
let dbCache = null; // Cache database object
let machinesCache = null;
let machinesLastFetched = 0;
const MACHINES_TTL_MS = 300000; // 5 minutes cache for machines (rarely changes)

let customersCache = null;
let customersLastFetched = 0;
const CUSTOMERS_TTL_MS = 300000; // 5 minutes cache for customers (rarely changes)

let feedbacksCache = null;
let feedbacksLastFetched = 0;
const FEEDBACKS_TTL_MS = 300000; // 5 minutes cache for feedbacks (rarely changes)

let jobsCache = null;
let jobsLastFetched = 0;
const JOBS_TTL_MS = 45000; // 45 seconds cache for jobs (updates during runs)

let hasMigratedTimezones = false;

const invalidateCache = () => {
  dbCache = null;
  machinesCache = null;
  machinesLastFetched = 0;
  customersCache = null;
  customersLastFetched = 0;
  feedbacksCache = null;
  feedbacksLastFetched = 0;
  jobsCache = null;
  jobsLastFetched = 0;
};
let startupError = null;
let dbReadCount = 0;
let dbWriteCount = 0;
let lastQuotaResetDate = new Date().toDateString();

const checkQuotaReset = () => {
  const today = new Date().toDateString();
  if (today !== lastQuotaResetDate) {
    dbReadCount = 0;
    dbWriteCount = 0;
    lastQuotaResetDate = today;
  }
};

const setupFirestoreProxy = (fdb) => {
  if (!fdb) return fdb;
  
  const QUOTA_FILE = path.join(__dirname, 'quota-usage.json');
  try {
    if (fs.existsSync(QUOTA_FILE)) {
      const data = JSON.parse(fs.readFileSync(QUOTA_FILE, 'utf-8'));
      if (data.date === new Date().toDateString()) {
        dbReadCount = data.reads || 0;
        dbWriteCount = data.writes || 0;
      }
    }
  } catch (e) {}

  const saveQuotaToFile = () => {
    try {
      fs.writeFileSync(QUOTA_FILE, JSON.stringify({
        date: new Date().toDateString(),
        reads: dbReadCount,
        writes: dbWriteCount
      }), 'utf-8');
    } catch (e) {}
  };

  const originalCollection = fdb.collection;
  fdb.collection = function(...args) {
    const colRef = originalCollection.apply(this, args);
    
    const originalGet = colRef.get;
    colRef.get = async function(...getArgs) {
      checkQuotaReset();
      dbReadCount += 1;
      saveQuotaToFile();
      try {
        const res = await originalGet.apply(this, getArgs);
        if (res && res.size !== undefined) {
          dbReadCount += res.size;
          saveQuotaToFile();
        }
        return res;
      } catch (err) {
        throw err;
      }
    };
    
    const originalDoc = colRef.doc;
    colRef.doc = function(...docArgs) {
      const docRef = originalDoc.apply(this, docArgs);
      
      const originalDocGet = docRef.get;
      docRef.get = async function(...dgArgs) {
        checkQuotaReset();
        dbReadCount += 1;
        saveQuotaToFile();
        return await originalDocGet.apply(this, dgArgs);
      };
      
      const originalSet = docRef.set;
      docRef.set = async function(...setArgs) {
        checkQuotaReset();
        dbWriteCount += 1;
        saveQuotaToFile();
        return await originalSet.apply(this, setArgs);
      };
      
      const originalUpdate = docRef.update;
      docRef.update = async function(...upArgs) {
        checkQuotaReset();
        dbWriteCount += 1;
        saveQuotaToFile();
        return await originalUpdate.apply(this, upArgs);
      };
      
      const originalDelete = docRef.delete;
      docRef.delete = async function(...delArgs) {
        checkQuotaReset();
        dbWriteCount += 1;
        saveQuotaToFile();
        return await originalDelete.apply(this, delArgs);
      };
      
      return docRef;
    };
    
    return colRef;
  };
  
  const originalBatch = fdb.batch;
  fdb.batch = function() {
    const batch = originalBatch.apply(this);
    const originalCommit = batch.commit;
    batch.commit = async function(...commitArgs) {
      checkQuotaReset();
      dbWriteCount += 1; 
      saveQuotaToFile();
      return await originalCommit.apply(this, commitArgs);
    };
    return batch;
  };

  return fdb;
};

// Try to initialize Firebase Admin SDK
const SERVICE_ACCOUNT_PATH = process.env.FIREBASE_SERVICE_ACCOUNT_PATH || '';
const SERVICE_ACCOUNT_JSON = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || '';

if (SERVICE_ACCOUNT_JSON) {
  try {
    const serviceAccount = JSON.parse(SERVICE_ACCOUNT_JSON);
    initializeApp({
      credential: cert(serviceAccount),
      storageBucket: process.env.FIREBASE_STORAGE_BUCKET || `${serviceAccount.project_id}.appspot.com`
    });
    db = setupFirestoreProxy(getFirestore('default'));
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
      credential: cert(serviceAccount),
      storageBucket: process.env.FIREBASE_STORAGE_BUCKET || `${serviceAccount.project_id}.appspot.com`
    });
    db = setupFirestoreProxy(getFirestore('default'));
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

// Verify Firestore connection and check for quota limits at startup
if (isCloud) {
  try {
    await db.collection('machines').limit(1).get();
    console.log('✅ Firestore connection and quota verified successfully.');
  } catch (err) {
    console.error('❌ Firestore connection test failed (possibly out of quota or offline). Falling back to local db.json. Error:', err.message);
    startupError = err.message;
    isCloud = false;
    db = null;
  }
}

let isStorageCloudAvailable = false;
if (isCloud) {
  try {
    const bucket = getStorage().bucket();
    const [exists] = await bucket.exists();
    if (exists) {
      isStorageCloudAvailable = true;
      console.log('✅ Firebase Storage bucket verified successfully.');
    } else {
      console.warn('⚠️ Firebase Storage bucket does not exist. Falling back to local disk storage.');
    }
  } catch (err) {
    console.error('❌ Firebase Storage connection test failed. Falling back to local disk storage. Error:', err.message);
  }
}

// ──────────────────────────────────────────────────────────
// Firestore Helper Functions
// ──────────────────────────────────────────────────────────

// Collection references
const MACHINES_COL = 'machines';
const JOBS_COL = 'jobs';
const CUSTOMERS_COL = 'customers';
const FEEDBACKS_COL = 'feedbacks';

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
          createdAt: new Date().toISOString(),
          data: []
        }
      ],
      customers: [],
      feedbacks: []
    };
    fs.writeFileSync(DB_FILE, JSON.stringify(defaultData, null, 2), 'utf-8');
    return defaultData;
  }
  try {
    const content = fs.readFileSync(DB_FILE, 'utf-8');
    const parsed = JSON.parse(content);
    if (!parsed.feedbacks) parsed.feedbacks = [];
    return parsed;
  } catch (e) {
    console.error('Error reading db.json, returning empty template:', e);
    return { machines: [], jobs: [], customers: [], feedbacks: [] };
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
// Timezone Migration Helper (Fixes old UTC-parsed timestamps & mismatched date/time fields)
// ──────────────────────────────────────────────────────────
const migrateDataPoints = (jobs) => {
  let migrated = false;
  jobs.forEach(job => {
    if (job.data && Array.isArray(job.data)) {
      job.data.forEach(row => {
        if (row.timestamp) {
          const dt = new Date(row.timestamp);
          if (!isNaN(dt.getTime())) {
            // Get correct ICT date and time (UTC+7)
            const ictTime = new Date(dt.getTime() + 7 * 60 * 60 * 1000);
            const correctDate = ictTime.toISOString().slice(0, 10);
            
            const pad2 = (n) => String(n).padStart(2, '0');
            const correctTime = `${pad2(ictTime.getUTCHours())}:${pad2(ictTime.getUTCMinutes())}`;

            let cleanTime = '';
            if (row.time) {
              const m = row.time.match(/(\d{1,2}):(\d{2})/);
              if (m) {
                cleanTime = `${pad2(parseInt(m[1], 10))}:${pad2(parseInt(m[2], 10))}`;
              }
            }

            const tsTime = row.timestamp.slice(11, 16);
            if (cleanTime && tsTime === cleanTime) {
              // Old UTC-parsed timestamp bug detected! Migrate timestamp first.
              const dateParts = (row.date || correctDate).split('-');
              const timeParts = cleanTime.split(':');
              if (dateParts.length === 3 && timeParts.length >= 2) {
                const year = parseInt(dateParts[0], 10);
                const month = parseInt(dateParts[1], 10) - 1;
                const day = parseInt(dateParts[2], 10);
                const hours = parseInt(timeParts[0], 10);
                const minutes = parseInt(timeParts[1], 10);
                const utcMs = Date.UTC(year, month, day, hours, minutes, 0) - (7 * 60 * 60 * 1000);
                const localDate = new Date(utcMs);
                if (!isNaN(localDate.getTime())) {
                  row.timestamp = localDate.toISOString();
                  migrated = true;
                  
                  // Re-evaluate correctDate and correctTime for the updated timestamp
                  const updatedDt = new Date(row.timestamp);
                  const updatedIctTime = new Date(updatedDt.getTime() + 7 * 60 * 60 * 1000);
                  const updatedDate = updatedIctTime.toISOString().slice(0, 10);
                  const updatedTime = `${pad2(updatedIctTime.getUTCHours())}:${pad2(updatedIctTime.getUTCMinutes())}`;
                  
                  if (row.date !== updatedDate) {
                    row.date = updatedDate;
                    migrated = true;
                  }
                  if (row.time !== updatedTime) {
                    row.time = updatedTime;
                    migrated = true;
                  }
                  return; // Go to next row
                }
              }
            }

            // If timestamp is fine, but stored date/time are incorrect
            if (row.date !== correctDate) {
              row.date = correctDate;
              migrated = true;
            }
            if (row.time !== correctTime) {
              row.time = correctTime;
              migrated = true;
            }
          }
        }
      });
    }
  });
  return migrated;
};

// ──────────────────────────────────────────────────────────
// Unified Data Fetch
// ──────────────────────────────────────────────────────────
const getDB = async () => {
  if (!isCloud) {
    const localDB = readLocalDB();
    const migrated = migrateDataPoints(localDB.jobs);
    if (migrated) {
      console.log('⚡ Detected old data points with timezone offset bugs. Migrating local db.json...');
      writeLocalDB(localDB);
    }
    dbCache = localDB;
    return localDB;
  }

  const now = Date.now();
  let machines = machinesCache;
  let jobs = jobsCache;
  let customers = customersCache;
  let feedbacks = feedbacksCache;

  try {
    // Fetch machines if expired or null (5 minutes cache)
    if (!machinesCache || (now - machinesLastFetched >= MACHINES_TTL_MS)) {
      machines = await getCollection(MACHINES_COL);
      machinesCache = machines;
      machinesLastFetched = now;
    }

    // Fetch jobs if expired or null (45 seconds cache)
    if (!jobsCache || (now - jobsLastFetched >= JOBS_TTL_MS)) {
      jobs = await getCollection(JOBS_COL);

      // Run timezone migration on loaded Firestore jobs ONLY ONCE on startup
      if (!hasMigratedTimezones) {
        const migrated = migrateDataPoints(jobs);
        if (migrated) {
          console.log('⚡ Detected old data points with timezone offset bugs. Migrating Firestore documents...');
          const batch = db.batch();
          for (const job of jobs) {
            batch.set(db.collection(JOBS_COL).doc(job.id), job);
          }
          await batch.commit();
          console.log('✅ Firestore timezone migration completed successfully.');
        }
        hasMigratedTimezones = true;
      }

      jobsCache = jobs;
      jobsLastFetched = now;
    }

    // Fetch customers if expired or null (5 minutes cache)
    if (!customersCache || (now - customersLastFetched >= CUSTOMERS_TTL_MS)) {
      customers = await getCollection(CUSTOMERS_COL);
      customersCache = customers;
      customersLastFetched = now;
    }

    // Fetch feedbacks if expired or null (5 minutes cache)
    if (!feedbacksCache || (now - feedbacksLastFetched >= FEEDBACKS_TTL_MS)) {
      feedbacks = await getCollection(FEEDBACKS_COL);
      feedbacksCache = feedbacks;
      feedbacksLastFetched = now;
    }

    // Seed cloud database if empty
    if (machines.length === 0) {
      const defaultMachine = { id: 'm1', name: 'Bioreactor 1' };
      const defaultJob = {
        id: 'job-initial',
        machineId: 'm1',
        name: 'Default Session',
        createdAt: new Date().toISOString(),
        data: []
      };
      await db.collection(MACHINES_COL).doc('m1').set(defaultMachine);
      await db.collection(JOBS_COL).doc('job-initial').set(defaultJob);

      machines = [defaultMachine];
      jobs = [defaultJob];
      machinesCache = machines;
      jobsCache = jobs;
    }

    dbCache = { machines, jobs, customers, feedbacks };
    return dbCache;
  } catch (e) {
    console.error('Error reading Firestore, falling back to local file:', e);
    const localDB = readLocalDB();
    dbCache = localDB;
    return localDB;
  }
};

// ──────────────────────────────────────────────────────────
// Settings Helpers
// ──────────────────────────────────────────────────────────
const SETTINGS_COL = 'settings';

const getSettings = async () => {
  let settings = {};
  if (isCloud) {
    try {
      const doc = await db.collection(SETTINGS_COL).doc('system').get();
      if (doc.exists) {
        settings = doc.data();
      }
    } catch (e) {
      console.error('Error reading settings from Firestore, falling back to local file:', e);
      const localDB = readLocalDB();
      settings = localDB.settings || {};
    }
  } else {
    const localDB = readLocalDB();
    settings = localDB.settings || {};
  }

  if (!settings.adminPassword) {
    settings.adminPassword = 'admin123';
  }
  
  // Ensure default developer info and VVM config are present
  const defaultAbout = {
    systemName: 'DBMS (Bioprocess Data Logging)',
    systemVersion: 'v2.4.0 (SCADA Polish)',
    developer: 'ทีมวิศวกรรมข้อมูลชีวภาพ (Bioprocess Engineering Team)',
    techStack: 'React / Vite / Node.js / GCS',
    supportEmail: 'support@bioprocess-logging.local',
    supportPhone: '+66 2 123 4567',
    vvmCalcType: 'dynamic',
    maxVolumeLiters: 5.0,
    constantVolumeLiters: 3.5,
    airUnit: 'mlmin',
    cctvUrl: 'https://www.w3schools.com/html/mov_bbb.mp4'
  };
  
  let needsWrite = false;
  for (const [k, v] of Object.entries(defaultAbout)) {
    if (settings[k] === undefined) {
      settings[k] = v;
      needsWrite = true;
    }
  }
  
  if (needsWrite) {
    if (isCloud) {
      try {
        await db.collection(SETTINGS_COL).doc('system').set(settings);
      } catch (e) {
        console.error('Failed to seed settings in Firestore:', e);
      }
    } else {
      const localDB = readLocalDB();
      localDB.settings = settings;
      writeLocalDB(localDB);
    }
  }
  
  return settings;
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

// Health check endpoint (used to keep server active via Uptime monitoring pings)
app.get('/api/health', (req, res) => {
  let firebaseProjectId = null;
  const SERVICE_ACCOUNT_JSON = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || '';
  if (SERVICE_ACCOUNT_JSON) {
    try {
      const sa = JSON.parse(SERVICE_ACCOUNT_JSON);
      firebaseProjectId = sa.project_id;
    } catch (e) {}
  }
  res.json({
    status: 'ok',
    time: new Date().toISOString(),
    isCloud,
    firebaseProjectId,
    startupError
  });
});

// Endpoint to check recent incoming pings/requests (useful for debugging UptimeRobot)
app.get('/api/pings', (req, res) => {
  res.json({
    message: "Recent HTTP requests received by server",
    uptimeRobotActive: recentRequests.some(r => r.userAgent.toLowerCase().includes('uptimerobot')),
    totalLogged: recentRequests.length,
    recentRequests
  });
});

// In-memory active users tracker
const activeUsers = {};

// ── Feedbacks ─────────────────────────────────────────────
app.post('/api/feedbacks', async (req, res) => {
  const { jobId, scores, channels, tools, suggestion, rating, comment } = req.body;
  if (!jobId) {
    return res.status(400).json({ error: 'Job ID is required' });
  }

  const dbData = await getDB();
  const job = dbData.jobs.find(j => j.id === jobId);
  const jobName = job ? job.name : 'Unknown Session';

  // Support both new format (scores) and legacy format (rating)
  let avgScore = 5;
  let normalizedScores = null;
  if (scores && typeof scores === 'object') {
    const vals = Object.values(scores).map(v => parseInt(v, 10)).filter(v => !isNaN(v) && v >= 1 && v <= 5);
    avgScore = vals.length > 0 ? parseFloat((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2)) : 5;
    normalizedScores = scores;
  } else if (rating) {
    avgScore = parseInt(rating, 10) || 5;
  }

  const newFeedback = {
    id: 'fb-' + Date.now(),
    jobId,
    jobName,
    // New structured fields
    scores: normalizedScores,          // { q1..q11 } or null for legacy
    channels: Array.isArray(channels) ? channels : [],
    tools: Array.isArray(tools) ? tools : [],
    suggestion: suggestion ? String(suggestion).trim() : '',
    avgScore,
    // Legacy fallback fields (kept for backward compat display)
    rating: avgScore,
    comment: suggestion ? String(suggestion).trim() : (comment ? String(comment).trim() : ''),
    createdAt: new Date().toISOString()
  };

  if (isCloud) {
    try {
      await db.collection(FEEDBACKS_COL).doc(newFeedback.id).set(newFeedback);
    } catch (e) {
      console.error('Firestore error saving feedback:', e);
    }
  } else {
    const localDB = readLocalDB();
    if (!localDB.feedbacks) localDB.feedbacks = [];
    localDB.feedbacks.push(newFeedback);
    writeLocalDB(localDB);
  }

  res.json(await getDB());
});

app.delete('/api/feedbacks/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await db.collection(FEEDBACKS_COL).doc(id).delete();
    } catch (e) {
      console.error('Firestore error deleting feedback:', e);
    }
  } else {
    const localDB = readLocalDB();
    if (localDB.feedbacks) {
      localDB.feedbacks = localDB.feedbacks.filter(f => f.id !== id);
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

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
    activeUsers: activeUsersList,
    isCloud
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

app.get('/api/settings', async (req, res) => {
  const settings = await getSettings();
  const publicSettings = { ...settings };
  delete publicSettings.adminPassword;
  res.json(publicSettings);
});

app.get('/api/storage-info', async (req, res) => {
  const cloudStorageFreeLimit = 5 * 1024 * 1024 * 1024; // 5 GB
  let usedBytes = 0;

  let dbData;
  try {
    dbData = await getDB();
  } catch (err) {
    dbData = { machines: [], jobs: [], customers: [], feedbacks: [] };
  }

  const machinesCount = dbData.machines ? dbData.machines.length : 0;
  const jobsCount = dbData.jobs ? dbData.jobs.length : 0;
  let dataPointsCount = 0;
  if (dbData.jobs && Array.isArray(dbData.jobs)) {
    dbData.jobs.forEach(job => {
      if (job.data && Array.isArray(job.data)) {
        dataPointsCount += job.data.length;
      }
    });
  }

  // Calculate size in bytes
  let machineImagesSize = 0;
  if (dbData.machines && Array.isArray(dbData.machines)) {
    dbData.machines.forEach(m => {
      if (m.imageData && typeof m.imageData === 'string') {
        // Approximate bytes for base64
        machineImagesSize += Math.round(m.imageData.length * 0.75);
      }
    });
  }

  // Determine if Storage is Cloud or Local
  let isStorageCloud = false;
  if (isCloud && isStorageCloudAvailable) {
    try {
      const bucket = getStorage().bucket();
      const [files] = await bucket.getFiles();
      for (const file of files) {
        usedBytes += parseInt(file.metadata.size || 0, 10);
      }
      isStorageCloud = true;
    } catch (err) {
      console.error('❌ Error getting Firebase Storage size, falling back to local storage:', err.message);
      if (err.code === 404 || err.message.includes('does not exist') || err.message.includes('notFound')) {
        console.log('ℹ️ Firebase Storage bucket not found/configured. Bypassing cloud storage for subsequent checks.');
        isStorageCloudAvailable = false;
      }
    }
  }

  if (!isStorageCloud) {
    // Local size calculation
    try {
      const getDirSize = (dirPath) => {
        let totalSize = 0;
        if (fs.existsSync(dirPath)) {
          const files = fs.readdirSync(dirPath);
          for (const file of files) {
            const filePath = path.join(dirPath, file);
            const stats = fs.statSync(filePath);
            if (stats.isFile()) {
              totalSize += stats.size;
            } else if (stats.isDirectory()) {
              totalSize += getDirSize(filePath);
            }
          }
        }
        return totalSize;
      };
      usedBytes = getDirSize(UPLOADS_DIR);
    } catch (err) {
      console.error('Error getting local storage size:', err);
    }
  }

  // Estimate DB sizes
  const machinesDbSize = machinesCount * 500;
  const jobsDbSize = jobsCount * 1000;
  const dataPointsDbSize = dataPointsCount * 400;

  return res.json({
    isCloud,
    isStorageCloud,
    firestoreQuota: {
      readsUsed: dbReadCount,
      readsLimit: 50000,
      writesUsed: dbWriteCount,
      writesLimit: 20000
    },
    storage: {
      usedBytes,
      freeLimitBytes: cloudStorageFreeLimit,
      plan: isCloud ? 'Blaze Plan (Pay as you go)' : 'Local Disk',
      rateInfo: '$0.026 per GB after free 5 GB'
    },
    totalUsed: usedBytes + machineImagesSize + machinesDbSize + jobsDbSize + dataPointsDbSize,
    categories: [
      {
        name: 'รูปภาพแนบในตารางประวัติ (Attached Images)',
        size: usedBytes,
        description: 'รูปถ่ายกิจกรรมและรายงานอัปโหลดระหว่างประวัติระบบ',
        icon: 'Camera'
      },
      {
        name: 'รูปถังปฏิกรณ์ชีวภาพ (Machine Avatars)',
        size: machineImagesSize,
        description: 'รูปแทนประเภทถังหมักและคอนฟิกเครื่องจักร',
        icon: 'Cpu'
      },
      {
        name: 'จุดข้อมูลบันทึก HMI (Data Logger Points)',
        size: dataPointsDbSize,
        description: 'ค่าพารามิเตอร์ Temp, pH, DO, RPM, AirFlow รายวินาที',
        icon: 'Activity'
      },
      {
        name: 'ข้อมูลเซสชันการรัน (Experimental Runs)',
        size: jobsDbSize,
        description: 'ประวัติเวลาการรันงานและบันทึกผู้ใช้',
        icon: 'Folder'
      },
      {
        name: 'ฐานข้อมูลตั้งค่าระบบ (Config Databases)',
        size: machinesDbSize,
        description: 'คอนฟิกตั้งค่าสเกล และระดับการตั้งสัญญาณเตือน',
        icon: 'Settings'
      }
    ]
  });
});

app.post('/api/settings/update-about', async (req, res) => {
  const { systemName, systemVersion, developer, techStack, supportEmail, supportPhone } = req.body;
  const settings = await getSettings();

  settings.systemName = (systemName || '').trim();
  settings.systemVersion = (systemVersion || '').trim();
  settings.developer = (developer || '').trim();
  settings.techStack = (techStack || '').trim();
  settings.supportEmail = (supportEmail || '').trim();
  settings.supportPhone = (supportPhone || '').trim();

  await saveSettings(settings);
  
  const publicSettings = { ...settings };
  delete publicSettings.adminPassword;
  res.json({ success: true, settings: publicSettings });
});

app.post('/api/settings/update-vvm', async (req, res) => {
  const { vvmCalcType, maxVolumeLiters, constantVolumeLiters, airUnit, cctvUrl } = req.body;
  const settings = await getSettings();

  settings.vvmCalcType = vvmCalcType || 'dynamic';
  settings.maxVolumeLiters = parseFloat(maxVolumeLiters) || 5.0;
  settings.constantVolumeLiters = parseFloat(constantVolumeLiters) || 3.5;
  settings.airUnit = airUnit || 'mlmin';
  settings.cctvUrl = cctvUrl !== undefined ? cctvUrl : '';

  await saveSettings(settings);
  
  const publicSettings = { ...settings };
  delete publicSettings.adminPassword;
  res.json({ success: true, settings: publicSettings });
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

// ── Machine Image Upload ──────────────────────────────────

app.put('/api/machines/:id/image', async (req, res) => {
  const { id } = req.params;
  const { imageData } = req.body; // base64 string or null (to remove)

  if (isCloud) {
    try {
      await db.collection(MACHINES_COL).doc(id).update({ imageData: imageData || null });
    } catch (e) {
      console.error('Firestore error updating machine image:', e);
      return res.status(500).json({ error: 'Failed to update machine image' });
    }
  } else {
    const localDB = readLocalDB();
    localDB.machines = localDB.machines.map(m =>
      m.id === id ? { ...m, imageData: imageData || null } : m
    );
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
  const { machineId, name, targetHours } = req.body;
  if (!machineId || !name || !name.trim()) {
    return res.status(400).json({ error: 'Machine ID and name are required' });
  }

  const parsedTargetHours = Number(targetHours);
  const newJob = {
    id: 'job-' + Date.now(),
    machineId,
    name: name.trim(),
    createdAt: new Date().toISOString(),
    status: 'running',
    targetHours: (!isNaN(parsedTargetHours) && parsedTargetHours > 0) ? parsedTargetHours : 48,
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

app.put('/api/jobs/:id/target-hours', async (req, res) => {
  const { id } = req.params;
  const { targetHours } = req.body;

  const hoursNum = Number(targetHours);
  if (isNaN(hoursNum) || hoursNum <= 0) {
    return res.status(400).json({ error: 'Invalid target hours value' });
  }

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).update({ targetHours: hoursNum });
    } catch (e) {
      console.error('Firestore error updating target hours:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.targetHours = hoursNum;
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

app.put('/api/jobs/:id/name', async (req, res) => {
  const { id } = req.params;
  const { name } = req.body;

  if (!name || !name.trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).update({ name: name.trim() });
    } catch (e) {
      console.error('Firestore error updating session name:', e);
      return res.status(500).json({ error: 'Failed to update session name' });
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.name = name.trim();
      writeLocalDB(localDB);
    } else {
      return res.status(404).json({ error: 'Session not found' });
    }
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

app.put('/api/jobs/:id/history-permission', async (req, res) => {
  const { id } = req.params;
  const { allowHistoryView } = req.body;

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).update({ allowHistoryView: !!allowHistoryView });
    } catch (e) {
      console.error('Firestore error updating history permission:', e);
      return res.status(500).json({ error: 'Failed to update history permission' });
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.allowHistoryView = !!allowHistoryView;
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});


app.put('/api/jobs/:id/status', async (req, res) => {
  const { id } = req.params;
  const { status } = req.body;

  if (status !== 'running' && status !== 'stopped') {
    return res.status(400).json({ error: 'Invalid status' });
  }

  if (isCloud) {
    try {
      // Check if already finished — cannot change status of finished job
      const doc = await db.collection(JOBS_COL).doc(id).get();
      if (doc.exists && doc.data().status === 'finished') {
        return res.status(400).json({ error: 'Cannot change status of a finished job' });
      }
      await db.collection(JOBS_COL).doc(id).update({ status });
    } catch (e) {
      console.error('Firestore error updating session status:', e);
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      if (job.status === 'finished') {
        return res.status(400).json({ error: 'Cannot change status of a finished job' });
      }
      job.status = status;
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

// Finish a job permanently (cannot be restarted)
app.put('/api/jobs/:id/finish', async (req, res) => {
  const { id } = req.params;
  const finishedAt = new Date().toISOString();

  if (isCloud) {
    try {
      await db.collection(JOBS_COL).doc(id).update({ status: 'finished', finishedAt });
    } catch (e) {
      console.error('Firestore error finishing job:', e);
      return res.status(500).json({ error: 'Failed to finish job' });
    }
  } else {
    const localDB = readLocalDB();
    const job = localDB.jobs.find(j => j.id === id);
    if (job) {
      job.status = 'finished';
      job.finishedAt = finishedAt;
      writeLocalDB(localDB);
    }
  }

  res.json(await getDB());
});

app.post('/api/upload', async (req, res) => {
  const { fileName, base64Data } = req.body;
  if (!base64Data) {
    return res.status(400).json({ error: 'No image data provided' });
  }

  const cleanFileName = fileName ? fileName.replace(/[^a-zA-Z0-9_.-]/g, '') : `img_${Date.now()}.jpg`;

  try {
    const matches = base64Data.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) {
      return res.status(400).json({ error: 'Invalid base64 image data format' });
    }
    const buffer = Buffer.from(matches[2], 'base64');
    const mimeType = matches[1];

    if (isCloud) {
      try {
        const bucket = getStorage().bucket();
        const fileRef = bucket.file(`uploads/${Date.now()}_${cleanFileName}`);
        
        await fileRef.save(buffer, {
          metadata: { contentType: mimeType },
          public: true
        });

        const publicUrl = `https://storage.googleapis.com/${bucket.name}/${fileRef.name}`;
        console.log(`✅ Uploaded image to Firebase Storage: ${publicUrl}`);
        return res.json({ url: publicUrl });
      } catch (err) {
        console.error('❌ Failed to upload to Firebase Storage, falling back to persistent database storage:', err.message);
        // Fallback: return base64 data directly so it gets saved in Firestore persistently
        // to prevent image loss on Render's ephemeral local disk.
        return res.json({ url: base64Data });
      }
    }

    const localPath = path.join(UPLOADS_DIR, `${Date.now()}_${cleanFileName}`);
    fs.writeFileSync(localPath, buffer);
    const localUrl = `/uploads/${path.basename(localPath)}`;
    console.log(`💾 Saved image locally (Offline): ${localUrl}`);
    return res.json({ url: localUrl });

  } catch (err) {
    console.error('❌ Error handling image upload:', err);
    return res.status(500).json({ error: 'Failed to process image upload' });
  }
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

  const timestamp = req.body.timestamp || normalizeDateTime(req.body.date, req.body.time) || normalizeDateTime(undefined, req.body.time) || new Date().toISOString();
  let date = null;
  let time = null;
  
  try {
    const dt = new Date(timestamp);
    if (!isNaN(dt.getTime())) {
      const ictTime = new Date(dt.getTime() + 7 * 60 * 60 * 1000);
      date = ictTime.toISOString().slice(0, 10);
      const pad2 = (n) => String(n).padStart(2, '0');
      time = `${pad2(ictTime.getUTCHours())}:${pad2(ictTime.getUTCMinutes())}`;
    }
  } catch (e) {
    console.error('Error deriving ICT date/time from timestamp:', e);
  }

  if (!date) {
    const ictNow = new Date(Date.now() + 7 * 60 * 60 * 1000);
    date = ictNow.toISOString().slice(0, 10);
    const pad2 = (n) => String(n).padStart(2, '0');
    time = `${pad2(ictNow.getUTCHours())}:${pad2(ictNow.getUTCMinutes())}`;
  }

  const newDataPoint = {
    timestamp,
    date,
    time,
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
    level_set: req.body.level_set !== undefined && req.body.level_set !== null ? parseFloat(req.body.level_set) : null,
    level_read: req.body.level_read !== undefined && req.body.level_read !== null ? parseFloat(req.body.level_read) : null,
    air_out_set: req.body.air_out_set !== undefined && req.body.air_out_set !== null ? parseFloat(req.body.air_out_set) : null,
    air_out_read: req.body.air_out_read !== undefined && req.body.air_out_read !== null ? parseFloat(req.body.air_out_read) : null,
    heat_set: req.body.heat_set !== undefined && req.body.heat_set !== null ? parseFloat(req.body.heat_set) : null,
    heat_read: req.body.heat_read !== undefined && req.body.heat_read !== null ? parseFloat(req.body.heat_read) : null,
    imageUrl: req.body.imageUrl || null,
    remark: remark || ''
  };

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

  const timestamp = updatedPoint.timestamp || normalizeDateTime(updatedPoint.date, updatedPoint.time) || new Date().toISOString();
  let date = null;
  let time = null;

  try {
    const dt = new Date(timestamp);
    if (!isNaN(dt.getTime())) {
      const ictTime = new Date(dt.getTime() + 7 * 60 * 60 * 1000);
      date = ictTime.toISOString().slice(0, 10);
      const pad2 = (n) => String(n).padStart(2, '0');
      time = `${pad2(ictTime.getUTCHours())}:${pad2(ictTime.getUTCMinutes())}`;
    }
  } catch (e) {
    console.error('Error deriving ICT date/time in PUT:', e);
  }

  if (!date) {
    const ictNow = new Date(Date.now() + 7 * 60 * 60 * 1000);
    date = ictNow.toISOString().slice(0, 10);
    const pad2 = (n) => String(n).padStart(2, '0');
    time = `${pad2(ictNow.getUTCHours())}:${pad2(ictNow.getUTCMinutes())}`;
  }

  const cleanPoint = {
    timestamp,
    date,
    time,
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
    level_set: updatedPoint.level_set !== undefined && updatedPoint.level_set !== null ? parseFloat(updatedPoint.level_set) : null,
    level_read: updatedPoint.level_read !== undefined && updatedPoint.level_read !== null ? parseFloat(updatedPoint.level_read) : null,
    air_out_set: updatedPoint.air_out_set !== undefined && updatedPoint.air_out_set !== null ? parseFloat(updatedPoint.air_out_set) : null,
    air_out_read: updatedPoint.air_out_read !== undefined && updatedPoint.air_out_read !== null ? parseFloat(updatedPoint.air_out_read) : null,
    heat_set: updatedPoint.heat_set !== undefined && updatedPoint.heat_set !== null ? parseFloat(updatedPoint.heat_set) : null,
    heat_read: updatedPoint.heat_read !== undefined && updatedPoint.heat_read !== null ? parseFloat(updatedPoint.heat_read) : null,
    imageUrl: updatedPoint.imageUrl || null,
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
    createdAt: new Date().toISOString()
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

// ── AI Integration (Gemini API) ───────────────────────────

app.post('/api/ai/analyze', async (req, res) => {
  const { jobId } = req.body;
  if (!jobId) {
    return res.status(400).json({ error: 'Job ID is required' });
  }

  const dbData = await getDB();
  const job = dbData.jobs.find(j => j.id === jobId);
  if (!job) {
    return res.status(404).json({ error: 'Session not found' });
  }

  const machine = dbData.machines.find(m => m.id === job.machineId);
  const machineName = machine ? machine.name : 'Unknown Machine';

  // Format data points for prompt
  let dataSummary = 'ไม่มีข้อมูลบันทึกในเซสชันนี้';
  if (job.data && job.data.length > 0) {
    // Sort chronologically
    const sortedData = [...job.data].sort((a, b) => new Date(a.timestamp || 0) - new Date(b.timestamp || 0));
    
    // Convert to a neat table representation
    const rows = sortedData.map((row, idx) => {
      // Find culture hour
      let cultureHour = 0;
      const firstTime = new Date(sortedData[0].timestamp).getTime();
      const currTime = new Date(row.timestamp).getTime();
      if (!isNaN(firstTime) && !isNaN(currTime)) {
        cultureHour = ((currTime - firstTime) / 3600000).toFixed(1);
      }
      
      const tempStr = `Temp: ${row.temp_read || 0}/${row.temp_set || 0}`;
      const phStr = `pH: ${row.ph_read || 0}/${row.ph_set || 0}`;
      const doStr = `DO: ${row.do_read || 0}/${row.do_set || 0}`;
      const agitStr = `Agit: ${row.agit_read || 0}/${row.agit_set || 0}`;
      const airStr = `Air: ${row.air_read || 0}/${row.air_set || 0}`;
      const remarkStr = row.remark ? `Remark: ${row.remark}` : '';
      
      return `ชม.ที่ ${cultureHour} (${row.time || ''}) | ${tempStr} | ${phStr} | ${doStr} | ${agitStr} | ${airStr} | ${remarkStr}`;
    });
    
    dataSummary = rows.join('\n');
  }

  const apiKey = process.env.GEMINI_API_KEY || '';

  if (!apiKey) {
    // Return high-fidelity mock data if API key is not configured
    const mockReport = `### 🤖 รายงานวิเคราะห์โดย AI (Simulated Mode)
*หมายเหตุ: คุณยังไม่ได้กำหนดค่า \`GEMINI_API_KEY\` ในระบบ ระบบจึงแสดงรายงานจำลองคุณภาพสูงนี้ให้ทดลองใช้งาน*

#### 1. บทสรุปผู้บริหาร (Executive Summary)
เซสชัน **"${job.name}"** ของเครื่องมือ **"${machineName}"** แสดงถึงกระบวนการหมักเบื้องต้นที่มีเสถียรภาพ อัตราการเติบโตของเซลล์เป็นไปตามแบบจำลองมาตรฐาน และค่าพารามิเตอร์ส่วนใหญ่ควบคุมได้ดีตามค่าเป้าหมาย (Set Value)

#### 2. การวิเคราะห์แนวโน้มพารามิเตอร์ (Trend Analysis)
*   **อุณหภูมิ (Temperature):** ควบคุมได้คงที่รอบๆ 37.0°C - 38.0°C สอดคล้องกับค่าเป้าหมายตลอดรันการผลิต
*   **ค่า pH:** มีแนวโน้มลดลงเล็กน้อยตามธรรมชาติเมื่อเกิดเมตาบอลิซึมของเซลล์ ระบบควบคุมอัตโนมัติสามารถทำงานเพื่อชดเชยการลดลงได้อย่างสมดุล
*   **ออกซิเจนละลาย (DO) & อัตราการกวน (Agitation):** เมื่อเซลล์มีความหนาแน่นสูงขึ้น (ชั่วโมงที่ 0.1-0.2) ความต้องการออกซิเจนจะเพิ่มขึ้น ส่งผลให้ระดับ DO เริ่มลดต่ำลง ซึ่งความเร็วรอบการกวน (Agitator RPM) ได้รันอยู่ที่ระดับตั้งไว้เพื่อพยุงระดับ DO

#### 3. การตรวจพบสิ่งผิดปกติ (Anomaly Detection)
*   ⚠️ **ตรวจพบค่า DO ลดลงรวดเร็ว:** ในจุดบันทึกล่าสุด ค่า DO ตกลงมาอย่างรวดเร็วเนื่องจากการใช้ออกซิเจนของเซลล์ที่โตขึ้น ควรคอยสังเกตการณ์หากค่า DO ลดต่ำกว่า 30%
*   💬 **บันทึกกิจกรรม:** มีการบันทึกการเติม *Antifoam* ที่ชั่วโมงการเพาะเลี้ยงล่าสุดเพื่อควบคุมฟอง ซึ่งถือว่าทำได้ทันเวลา

#### 4. คำแนะนำในการปรับปรุง (Optimization Recommendations)
1.  **การปรับอากาศ (Aeration):** หากระดับ DO ลดต่ำกว่า 40% ให้พิจารณาเพิ่มอัตราไหลอากาศลม (Air Flow Rate) จากเดิมขึ้นอีก 10-20% หรือจ่ายแก๊สออกซิเจนบริสุทธิ์เข้าระบบ
2.  **อัตราการกวน (Agitation):** เพิ่มรอบการกวนทีละน้อยหากระดับ DO ยังคงลดต่ำลง เพื่อเสริมอัตราการละลายออกซิเจน (Mass Transfer Coefficient, kLa)`;

    return res.json({ report: mockReport, isMock: true });
  }

  try {
    const prompt = `คุณคือผู้เชี่ยวชาญด้านกระบวนการทางชีวภาพ (Bioprocess Expert) และวิศวกรควบคุมระบบถังหมัก
วิเคราะห์ข้อมูลการรันถังหมัก (Bioreactor) ดังต่อไปนี้:
ชื่อเซสชัน: ${job.name}
เครื่องมือ: ${machineName}
เวลาที่เริ่มสร้างรัน: ${job.createdAt}

ข้อมูลตัวชี้วัดกระบวนการ (เรียงตามเวลา):
ชั่วโมงเพาะเลี้ยง (เวลาบันทึก) | อุณหภูมิ (Read/Set) | pH (Read/Set) | DO (Read/Set) | Agitation (Read/Set) | Air Flow (Read/Set) | บันทึกเพิ่มเติม
${dataSummary}

กรุณาเขียนรายงานวิเคราะห์รายละเอียดเป็นภาษาไทยในรูปแบบ Markdown ประกอบด้วยหัวข้อดังนี้:
1. **บทสรุปผู้บริหาร (Executive Summary)**: สรุปภาพรวมความสมบูรณ์ของการรันและผลผลิตโดยรวม
2. **การวิเคราะห์แนวโน้มพารามิเตอร์ (Parameter Trend Analysis)**: เจาะลึกแนวโน้มของอุณหภูมิ, pH, DO, อัตรากวน และการจ่ายลม
3. **การตรวจพบสิ่งผิดปกติ (Anomaly Detection)**: ระบุจุดเบี่ยงเบนหรือความผิดปกติใดๆ (เช่น สัญญาณพารามิเตอร์แกว่ง, ค่าตกลงรวดเร็ว, การเติมสารชดเชยล่าช้า) หรือยืนยันว่ารันปกติดี
4. **ข้อเสนอแนะในการควบคุมกระบวนการ (Optimization Recommendations)**: ข้อแนะนำที่เป็นรูปธรรมสำหรับการปรับจูนค่าพารามิเตอร์ถัดไปเพื่อเพิ่มความเสถียรและ Yield`;

    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }]
      })
    });

    if (!response.ok) {
      const errorData = await response.json();
      console.error('Gemini API Error details:', errorData);
      throw new Error(errorData.error?.message || 'Failed to call Gemini API');
    }

    const data = await response.json();
    const textResponse = data.candidates?.[0]?.content?.parts?.[0]?.text || 'ไม่สามารถดึงข้อมูลรายงานวิเคราะห์ได้';
    res.json({ report: textResponse, isMock: false });
  } catch (err) {
    console.error('Gemini API Error:', err.message);
    res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเรียกใช้ AI: ' + err.message });
  }
});

app.post('/api/ai/chat', async (req, res) => {
  const { jobId, messages, newMessage } = req.body;
  if (!jobId || !newMessage) {
    return res.status(400).json({ error: 'Job ID and newMessage are required' });
  }

  const dbData = await getDB();
  const job = dbData.jobs.find(j => j.id === jobId);
  if (!job) {
    return res.status(404).json({ error: 'Session not found' });
  }

  const machine = dbData.machines.find(m => m.id === job.machineId);
  const machineName = machine ? machine.name : 'Unknown Machine';

  // Format data points for context
  let dataSummary = 'ไม่มีข้อมูลบันทึกในเซสชันนี้';
  if (job.data && job.data.length > 0) {
    const sortedData = [...job.data].sort((a, b) => new Date(a.timestamp || 0) - new Date(b.timestamp || 0));
    const rows = sortedData.map((row, idx) => {
      let cultureHour = 0;
      const firstTime = new Date(sortedData[0].timestamp).getTime();
      const currTime = new Date(row.timestamp).getTime();
      if (!isNaN(firstTime) && !isNaN(currTime)) {
        cultureHour = ((currTime - firstTime) / 3600000).toFixed(1);
      }
      return `ชม.ที่ ${cultureHour} (${row.time || ''}) | Temp: ${row.temp_read}/${row.temp_set} | pH: ${row.ph_read}/${row.ph_set} | DO: ${row.do_read}/${row.do_set} | Agit: ${row.agit_read}/${row.agit_set} | Air: ${row.air_read}/${row.air_set} | ${row.remark || ''}`;
    });
    dataSummary = rows.join('\n');
  }

  const apiKey = process.env.GEMINI_API_KEY || '';

  if (!apiKey) {
    // Return mock response
    let mockResponse = `🤖 ผู้ช่วย AI (Simulated Mode): ขออภัยครับ ปัจจุบันระบบไม่ได้เชื่อมโยง \`GEMINI_API_KEY\` สำหรับคุยสด 
แต่จากข้อมูลจำลองกระบวนการผลิต "${job.name}" ค่า pH อยู่ที่ ${job.data[job.data.length - 1]?.ph_read || 7.0} และอุณหภูมิปัจจุบันคือ ${job.data[job.data.length - 1]?.temp_read || 37.0}°C ซึ่งเป็นค่าที่เหมาะสมสำหรับการเจริญเติบโตของจุลินทรีย์ครับ! 
(โปรดตั้งค่า GEMINI_API_KEY ในไฟล์ .env หรือ Render Dashboard เพื่อแชตกับ AI จริง)`;
    
    if (newMessage.includes('ผิดปกติ') || newMessage.toLowerCase().includes('anomaly')) {
      mockResponse = `🤖 ผู้ช่วย AI (Simulated Mode): จากการสแกนข้อมูลเบื้องต้น ค่า DO ในชั่วโมงสุดท้ายลดลงค่อนข้างเร็ว และพบการเติม Antifoam ครับ แต่อุณหภูมิและพารามิเตอร์อื่นๆ ยังคงสมบูรณ์ดีครับ`;
    } else if (newMessage.includes('pH') || newMessage.includes('ph')) {
      mockResponse = `🤖 ผู้ช่วย AI (Simulated Mode): ค่า pH ปัจจุบันรันอยู่ที่ ${job.data[job.data.length - 1]?.ph_read || 7.00} โดยมีค่าเป้าหมายคือ ${job.data[job.data.length - 1]?.ph_set || 7.00} ซึ่งทำงานสอดคล้องสัมพันธ์กันดีครับ`;
    }
    
    return res.json({ response: mockResponse, isMock: true });
  }

  try {
    // Formulate system prompt with grounding context
    const systemPrompt = `คุณคือผู้ช่วย AI ด้านกระบวนการทางชีวภาพ (Bioprocess AI Co-pilot) ประจำห้องแล็บถังหมัก
คุณคอยช่วยเหลือโอเปอเรเตอร์และตอบคำถามเกี่ยวกับรอบรันถังหมักชื่อ "${job.name}" (เครื่องมือ: "${machineName}")
นี่คือตารางข้อมูลบันทึกในถังหมักล่าสุด (เรียงตามเวลา):
ชั่วโมงเพาะเลี้ยง (เวลาบันทึก) | อุณหภูมิ (Read/Set) | pH (Read/Set) | DO (Read/Set) | Agit (Read/Set) | Air (Read/Set) | บันทึกเพิ่มเติม
${dataSummary}

กรุณาตอบคำถามของผู้ใช้งานโดยวิเคราะห์อิงจากข้อมูลด้านบน ตอบคำถามกระชับ เข้าใจง่าย และให้คำแนะนำเชิงวิชาการ/วิศวกรรมที่ถูกต้อง เป็นกันเองและเป็นมืออาชีพ ตอบเป็นภาษาไทย`;

    // Map message history to Gemini API format
    const contents = [];
    
    // Add history
    if (messages && Array.isArray(messages)) {
      messages.forEach(msg => {
        contents.push({
          role: msg.role === 'user' ? 'user' : 'model',
          parts: [{ text: msg.content }]
        });
      });
    }

    // Add current query, but prefix it with system instructions to ensure it remains grounded in the data
    const lastPromptText = `${systemPrompt}\n\nคำถามล่าสุดของผู้ใช้งาน: ${newMessage}`;
    contents.push({
      role: 'user',
      parts: [{ text: lastPromptText }]
    });

    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ contents })
    });

    if (!response.ok) {
      const errorData = await response.json();
      console.error('Gemini API Error details:', errorData);
      throw new Error(errorData.error?.message || 'Failed to call Gemini API');
    }

    const data = await response.json();
    const textResponse = data.candidates?.[0]?.content?.parts?.[0]?.text || 'ขออภัยครับ ไม่สามารถสร้างคำตอบได้ในขณะนี้';
    res.json({ response: textResponse, isMock: false });
  } catch (err) {
    console.error('Gemini API Error:', err.message);
    res.status(500).json({ error: 'เกิดข้อผิดพลาดในการเรียกใช้ AI: ' + err.message });
  }
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
