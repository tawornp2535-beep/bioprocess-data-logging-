import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { initializeApp, cert } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SERVICE_ACCOUNT_PATH = process.env.FIREBASE_SERVICE_ACCOUNT_PATH || 'firebase-service-account.json';
const serviceAccountFullPath = path.resolve(__dirname, SERVICE_ACCOUNT_PATH);

if (!fs.existsSync(serviceAccountFullPath)) {
  process.exit(1);
}

try {
  const serviceAccount = JSON.parse(fs.readFileSync(serviceAccountFullPath, 'utf-8'));
  initializeApp({
    credential: cert(serviceAccount)
  });
} catch (err) {
  process.exit(1);
}

const db = getFirestore('default');

async function checkJobs() {
  try {
    const machinesSnapshot = await db.collection('machines').get();
    console.log('--- MACHINES IN FIRESTORE ---');
    machinesSnapshot.forEach(doc => {
      console.log(doc.id, '=>', doc.data().name);
    });

    const jobsSnapshot = await db.collection('jobs').get();
    console.log('--- JOBS IN FIRESTORE ---');
    jobsSnapshot.forEach(doc => {
      const data = doc.data();
      console.log(doc.id, '=> name:', data.name, 'machineId:', data.machineId, 'dataPoints:', data.data ? data.data.length : 0);
    });
    process.exit(0);
  } catch (error) {
    console.error(error);
    process.exit(1);
  }
}

checkJobs();
