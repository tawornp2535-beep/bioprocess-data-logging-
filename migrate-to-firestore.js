import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { initializeApp, cert } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';

// Load environment variables
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DB_FILE = path.join(__dirname, 'db.json');

const SERVICE_ACCOUNT_PATH = process.env.FIREBASE_SERVICE_ACCOUNT_PATH || 'firebase-service-account.json';
const serviceAccountFullPath = path.resolve(__dirname, SERVICE_ACCOUNT_PATH);

if (!fs.existsSync(serviceAccountFullPath)) {
  console.error(`❌ Error: Service account key file not found at: ${serviceAccountFullPath}`);
  console.log('Please download your Firebase Private Key JSON file, rename it to "firebase-service-account.json", and place it in the project directory.');
  process.exit(1);
}

if (!fs.existsSync(DB_FILE)) {
  console.error(`❌ Error: Local database file not found at: ${DB_FILE}`);
  process.exit(1);
}

// Initialize Firebase Admin
try {
  const serviceAccount = JSON.parse(fs.readFileSync(serviceAccountFullPath, 'utf-8'));
  initializeApp({
    credential: cert(serviceAccount)
  });
  console.log('✅ Connected to Firebase!');
} catch (err) {
  console.error('❌ Failed to initialize Firebase:', err.message);
  process.exit(1);
}

const db = getFirestore('default');

// Migration function
async function migrate() {
  try {
    const rawData = fs.readFileSync(DB_FILE, 'utf-8');
    const localData = JSON.parse(rawData);

    console.log('\nStarting migration...');

    // 1. Migrate Machines
    if (localData.machines && localData.machines.length > 0) {
      console.log(`\nMigrating ${localData.machines.length} machines...`);
      for (const machine of localData.machines) {
        if (!machine.id) continue;
        const docRef = db.collection('machines').doc(machine.id);
        const docSnap = await docRef.get();
        if (docSnap.exists) {
          console.log(`  - Machine ${machine.id} (${machine.name}) already exists in Firestore, skipping.`);
        } else {
          await docRef.set(machine);
          console.log(`  - Migrated machine: ${machine.name} (${machine.id})`);
        }
      }
    }

    // 2. Migrate Jobs (Sessions)
    if (localData.jobs && localData.jobs.length > 0) {
      console.log(`\nMigrating ${localData.jobs.length} sessions...`);
      for (const job of localData.jobs) {
        if (!job.id) continue;
        const docRef = db.collection('jobs').doc(job.id);
        const docSnap = await docRef.get();
        if (docSnap.exists) {
          console.log(`  - Session ${job.id} (${job.name}) already exists in Firestore, skipping.`);
        } else {
          await docRef.set(job);
          console.log(`  - Migrated session: ${job.name} (${job.id}) with ${job.data ? job.data.length : 0} data points`);
        }
      }
    }

    // 3. Migrate Customers
    if (localData.customers && localData.customers.length > 0) {
      console.log(`\nMigrating ${localData.customers.length} customers...`);
      for (const customer of localData.customers) {
        if (!customer.id) continue;
        const docRef = db.collection('customers').doc(customer.id);
        const docSnap = await docRef.get();
        if (docSnap.exists) {
          console.log(`  - Customer ${customer.id} (${customer.companyName}) already exists in Firestore, skipping.`);
        } else {
          const cleanCustomer = JSON.parse(JSON.stringify(customer));
          await docRef.set(cleanCustomer);
          console.log(`  - Migrated customer: ${customer.companyName} (${customer.id})`);
        }
      }
    }

    console.log('\n🎉 Migration completed successfully!');
    process.exit(0);
  } catch (error) {
    console.error('❌ Migration failed:', error);
    process.exit(1);
  }
}

migrate();
