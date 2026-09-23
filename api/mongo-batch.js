import { MongoClient } from 'mongodb';

const uri    = process.env.MONGO_URI;
const SECRET = process.env.BATCH_SECRET;
const DB     = 'sar-id-ops'; // DB name
const TTL_MS = 6 * 60 * 60 * 1000; // 6 hrs

let mongoClient;
async function getClient() {
  if (!mongoClient) {
    mongoClient = new MongoClient(uri, { maxPoolSize: 5, socketTimeoutMS: 45000 });
    await mongoClient.connect();
  }
  return mongoClient;
}

let cache = { data: null, builtAt: null };

async function buildCache(db) {
  console.log('[ID-OPS] Building cache...');
  const now   = new Date();
  const from  = new Date('2024-01-01').toISOString();
  const to    = now.toISOString();
  const today = now.toISOString().slice(0, 10);

  // Fetch all shipment records
  const allDocs = await db.collection('id_ops_shipments').find({
    etdDate: { $gte: from, $lte: to }
  }).toArray();

  console.log(`[ID-OPS] Total docs: ${allDocs.length}`);
  cache = { data: allDocs, builtAt: now.toISOString() };
  return cache;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-batch-secret');
  if (req.method === 'OPTIONS') return res.status(200).end();

  // ── PUSH from Apps Script ──────────────────────────────────────────────────
  if (req.method === 'POST') {
    const secret = req.headers['x-batch-secret'];
    if (!secret || secret !== SECRET) return res.status(401).json({ error: 'Unauthorized' });

    const { action, records, direction } = req.body;

    if (action === 'push') {
      if (!records || !records.length) return res.status(400).json({ error: 'No records' });
      const c   = await getClient();
      const db  = c.db(DB);
      const col = db.collection('id_ops_shipments');

      // Delete existing records for this direction then reinsert
      await col.deleteMany({ direction: direction });
      const result = await col.insertMany(records, { ordered: false });

      // Bust cache
      cache = { data: null, builtAt: null };

      return res.status(200).json({ inserted: result.insertedCount, direction });
    }

    return res.status(400).json({ error: 'Unknown action' });
  }

  // ── GET — serve dashboard data ─────────────────────────────────────────────
  if (req.method === 'GET') {
    try {
      const now = new Date();
      if (!cache.data || !cache.builtAt || (now - new Date(cache.builtAt)) > TTL_MS) {
        const c  = await getClient();
        const db = c.db(DB);
        await buildCache(db);
      }
      return res.status(200).json({ records: cache.data, builtAt: cache.builtAt });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
