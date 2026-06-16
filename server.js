import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import mongoose from 'mongoose';
import dns from 'dns';

// Set DNS servers to Google's public DNS to handle MongoDB Atlas SRV query resolution on Windows
dns.setServers(['8.8.8.8', '8.8.4.4']);

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
    // dateInput expected YYYY-MM-DD, timeInput HH:MM or HH:MM:SS
    const dt = new Date(`${dateInput}T${timeInput}`);
    if (!isNaN(dt.getTime())) return dt.toISOString();
    // fallback try parsing with space
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

const MONGODB_URI = process.env.MONGODB_URI;
let isCloud = false;

// Connect to Cloud Database if URI is present
if (MONGODB_URI && MONGODB_URI.trim()) {
  try {
    await mongoose.connect(MONGODB_URI, {
      serverSelectionTimeoutMS: 10000,
      socketTimeoutMS: 45000,
      connectTimeoutMS: 10000,
    });
    console.log("Successfully connected to cloud MongoDB database.");
    isCloud = true;
  } catch (err) {
    console.error("Failed to connect to MongoDB. Falling back to local db.json. Error:", err);
  }
} else {
  console.log("No MONGODB_URI found. Operating with local db.json file storage.");
}

// Mongoose Schemas for Cloud DB
const machineSchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  name: { type: String, required: true }
});

const jobSchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  machineId: { type: String, required: true },
  name: { type: String, required: true },
  createdAt: { type: String, required: true },
  data: { type: Array, default: [] }
});

const customerSchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  companyName: { type: String, required: true },
  machineId: { type: String, required: true },
  email: { type: String },
  createdAt: { type: String, required: true }
});

const MachineModel = mongoose.model('Machine', machineSchema);
const JobModel = mongoose.model('Job', jobSchema);
const CustomerModel = mongoose.model('Customer', customerSchema);

// Helper to read local JSON DB
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

// Helper to write local JSON DB
const writeLocalDB = (data) => {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing db.json:', e);
  }
};

// Unified Data Fetch Helper
const getDB = async () => {
  if (isCloud) {
    try {
      let machines = await MachineModel.find({}, { _id: 0, __v: 0 }).lean();
      let jobs = await JobModel.find({}, { _id: 0, __v: 0 }).lean();
      let customers = await CustomerModel.find({}, { _id: 0, __v: 0 }).lean();

      // Seed cloud database if empty (use upsert to avoid duplicate key errors)
      if (machines.length === 0) {
        const defaultMachine = { id: 'm1', name: 'Bioreactor 1' };
        const defaultJob = {
          id: 'job-initial',
          machineId: 'm1',
          name: 'Default Session',
          createdAt: new Date().toLocaleString(),
          data: []
        };
        await MachineModel.updateOne({ id: 'm1' }, defaultMachine, { upsert: true });
        await JobModel.updateOne({ id: 'job-initial' }, defaultJob, { upsert: true });
        
        machines = await MachineModel.find({}, { _id: 0, __v: 0 }).lean();
        jobs = await JobModel.find({}, { _id: 0, __v: 0 }).lean();
      }
      return { machines, jobs, customers };
    } catch (e) {
      console.error("Error reading MongoDB, falling back to local file:", e);
      return readLocalDB();
    }
  } else {
    return readLocalDB();
  }
};

// Routes
app.get('/api/db', async (req, res) => {
  res.json(await getDB());
});

app.post('/api/machines', async (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }

  const id = 'm-' + Date.now();
  const newMachine = { id, name: name.trim() };

  if (isCloud) {
    try {
      await MachineModel.create(newMachine);
    } catch (e) {
      console.error("MongoDB error creating machine:", e);
    }
  } else {
    const db = readLocalDB();
    db.machines.push(newMachine);
    writeLocalDB(db);
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
      await MachineModel.updateOne({ id }, { name: name.trim() });
    } catch (e) {
      console.error("MongoDB error renaming machine:", e);
    }
  } else {
    const db = readLocalDB();
    db.machines = db.machines.map(m => m.id === id ? { ...m, name: name.trim() } : m);
    writeLocalDB(db);
  }

  res.json(await getDB());
});

app.delete('/api/machines/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await MachineModel.deleteOne({ id });
      await JobModel.deleteMany({ machineId: id });
      await CustomerModel.deleteMany({ machineId: id });
    } catch (e) {
      console.error("MongoDB error deleting machine:", e);
    }
  } else {
    const db = readLocalDB();
    db.machines = db.machines.filter(m => m.id !== id);
    db.jobs = db.jobs.filter(j => j.machineId !== id);
    db.customers = db.customers.filter(c => c.machineId !== id);
    writeLocalDB(db);
  }

  res.json(await getDB());
});

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
      await JobModel.create(newJob);
    } catch (e) {
      console.error("MongoDB error creating session:", e);
    }
  } else {
    const db = readLocalDB();
    db.jobs.push(newJob);
    writeLocalDB(db);
  }

  res.json(await getDB());
});

app.delete('/api/jobs/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await JobModel.deleteOne({ id });
    } catch (e) {
      console.error("MongoDB error deleting session:", e);
    }
  } else {
    const db = readLocalDB();
    db.jobs = db.jobs.filter(j => j.id !== id);
    writeLocalDB(db);
  }

  res.json(await getDB());
});

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
    // store full timestamp (ISO) if provided or inferred, keep date (YYYY-MM-DD) and time as HH:MM
    timestamp: normalizeDateTime(req.body.date, req.body.time) || normalizeDateTime(undefined, req.body.time) || new Date().toISOString(),
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

  if (isCloud) {
    try {
      // derive date (YYYY-MM-DD) from timestamp for cloud storage
      try {
        const dt = new Date(newDataPoint.timestamp);
        if (!isNaN(dt.getTime())) {
          newDataPoint.date = dt.toISOString().slice(0,10);
        } else {
          newDataPoint.date = newDataPoint.timestamp ? String(newDataPoint.timestamp).slice(0,10) : (new Date().toISOString().slice(0,10));
        }
      } catch (e) {
        newDataPoint.date = new Date().toISOString().slice(0,10);
      }
      await JobModel.updateOne({ id }, { $push: { data: newDataPoint } });
    } catch (e) {
      console.error("MongoDB error pushing data point:", e);
    }
  } else {
    const db = readLocalDB();
    const job = db.jobs.find(j => j.id === id);
    if (job) {
      // derive date (YYYY-MM-DD) from timestamp for local storage
      try {
        const dt = new Date(newDataPoint.timestamp);
        if (!isNaN(dt.getTime())) {
          newDataPoint.date = dt.toISOString().slice(0,10);
        } else {
          newDataPoint.date = newDataPoint.timestamp ? String(newDataPoint.timestamp).slice(0,10) : (new Date().toISOString().slice(0,10));
        }
      } catch (e) {
        newDataPoint.date = new Date().toISOString().slice(0,10);
      }
      job.data.push(newDataPoint);
      writeLocalDB(db);
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
      const job = await JobModel.findOne({ id });
      if (job && idx < job.data.length) {
        job.data.splice(idx, 1);
        await JobModel.updateOne({ id }, { data: job.data });
      }
    } catch (e) {
      console.error("MongoDB error deleting data point:", e);
    }
  } else {
    const db = readLocalDB();
    const job = db.jobs.find(j => j.id === id);
    if (job && idx < job.data.length) {
      job.data.splice(idx, 1);
      writeLocalDB(db);
    }
  }

  res.json(await getDB());
});

app.post('/api/jobs/:id/clear', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await JobModel.updateOne({ id }, { data: [] });
    } catch (e) {
      console.error("MongoDB error clearing data:", e);
    }
  } else {
    const db = readLocalDB();
    const job = db.jobs.find(j => j.id === id);
    if (job) {
      job.data = [];
      writeLocalDB(db);
    }
  }

  res.json(await getDB());
});

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
      await CustomerModel.create(newCustomer);
    } catch (e) {
      console.error("MongoDB error creating customer:", e);
    }
  } else {
    const db = readLocalDB();
    db.customers.push(newCustomer);
    writeLocalDB(db);
  }

  res.json(await getDB());
});

app.delete('/api/customers/:id', async (req, res) => {
  const { id } = req.params;

  if (isCloud) {
    try {
      await CustomerModel.deleteOne({ id });
    } catch (e) {
      console.error("MongoDB error deleting customer:", e);
    }
  } else {
    const db = readLocalDB();
    db.customers = db.customers.filter(c => c.id !== id);
    writeLocalDB(db);
  }

  res.json(await getDB());
});

// Serve static frontend files from 'dist' directory
const distPath = path.join(__dirname, 'dist');
app.use(express.static(distPath));

// Fallback for SPA routing: send index.html for non-API requests
app.get(/^(.*)$/, (req, res, next) => {
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
  console.log(`Backend server running on http://0.0.0.0:${PORT}`);
});
