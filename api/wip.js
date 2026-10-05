import { MongoClient } from 'mongodb';

const uri    = process.env.MONGO_URI;
const SECRET = process.env.BATCH_SECRET;
const DB     = 'sar-id-ops';
const TTL_MS = 6 * 60 * 60 * 1000;

let mongoClient;
async function getClient() {
  if (!mongoClient) {
    mongoClient = new MongoClient(uri, { maxPoolSize: 5, socketTimeoutMS: 45000 });
    await mongoClient.connect();
  }
  return mongoClient;
}

let cache = { data: null, builtAt: null };

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-batch-secret');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method === 'POST') {
    const secret = req.headers['x-batch-secret'];
    if (!secret || secret !== SECRET) return res.status(401).json({ error: 'Unauthorized' });

    const { action, records } = req.body;

    if (action === 'wipe') {
      const c  = await getClient();
      const db = c.db(DB);
      const result = await db.collection('th_ops_wip').deleteMany({});
      cache = { data: null, builtAt: null };
      return res.status(200).json({ deleted: result.deletedCount });
    }

    if (action === 'push') {
      if (!records || !records.length) return res.status(400).json({ error: 'No records' });
      const c  = await getClient();
      const db = c.db(DB);
      const result = await db.collection('th_ops_wip').insertMany(records, { ordered: false });
      cache = { data: null, builtAt: null };
      return res.status(200).json({ inserted: result.insertedCount });
    }

    return res.status(400).json({ error: 'Unknown action' });
  }

  if (req.method === 'GET') {
    try {
      const now = new Date();
      if (!cache.data || !cache.builtAt || (now - new Date(cache.builtAt)) > TTL_MS) {
        const c  = await getClient();
        const db = c.db(DB);
        console.log('[TH-WIP] Building cache...');
        cache.data = await db.collection('th_ops_wip').find({}).toArray();
        cache.builtAt = new Date().toISOString();
        console.log(`[TH-WIP] Total docs: ${cache.data.length}`);
      }
      return res.status(200).json({ records: cache.data, builtAt: cache.builtAt });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
