require('dotenv').config();
const { App } = require('@slack/bolt');
const express = require('express');
const path = require('path');
const https = require('https');
const Anthropic = require('@anthropic-ai/sdk');
const { Pool } = require('pg');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL && process.env.DATABASE_URL.includes('railway')
    ? { rejectUnauthorized: false }
    : false,
  connectionTimeoutMillis: 5000,
});

let slackTeamDomain = '';

// ─── In-memory cache (loaded from DB on startup) ──────────────────────────────

let ordersCache   = {};
let channelsCache = { restaurant: [], grocery: [] };

// ─── DB helpers ───────────────────────────────────────────────────────────────

let dbAvailable = false;

async function initDb() {
  if (!process.env.DATABASE_URL) {
    console.log('No DATABASE_URL — using in-memory storage only.');
    return;
  }
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS store (
        key   TEXT PRIMARY KEY,
        value TEXT NOT NULL
      )
    `);
    const o = await pool.query("SELECT value FROM store WHERE key = 'orders'");
    const c = await pool.query("SELECT value FROM store WHERE key = 'channels'");
    if (o.rows.length) ordersCache   = JSON.parse(o.rows[0].value);
    if (c.rows.length) channelsCache = JSON.parse(c.rows[0].value);
    dbAvailable = true;
    console.log('Database connected — orders:', Object.keys(ordersCache).length, 'customers');
  } catch (e) {
    console.error('Database connection failed, running without persistence:', e.message);
  }
}

async function dbSet(key, value) {
  if (!dbAvailable) return;
  try {
    await pool.query(
      `INSERT INTO store (key, value) VALUES ($1, $2)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      [key, JSON.stringify(value)]
    );
  } catch (e) {
    console.error('DB write error:', e.message);
  }
}

// ─── Storage helpers (synchronous via cache) ──────────────────────────────────

function loadOrders()   { return ordersCache; }
function loadChannels() { return channelsCache; }

function saveOrders(orders) {
  ordersCache = orders;
  dbSet('orders', orders).catch(e => console.error('DB write error (orders):', e));
}

function saveChannels(ch) {
  channelsCache = ch;
  dbSet('channels', ch).catch(e => console.error('DB write error (channels):', e));
}

function typeForChannel(channelId) {
  const ch = loadChannels();
  if (ch.restaurant.includes(channelId)) return 'restaurant';
  if (ch.grocery.includes(channelId))    return 'grocery';
  return 'unknown';
}

// ─── Item categorizer (based on weekly availability list categories) ─────────

const ITEM_CATEGORIES = [
  { category: 'Tomatoes', keywords: ['tomato', 'tomatoes', 'beefsteak', 'roma', 'plum', 'heirloom', 'cherry tom', 'grape tom', 'vine ripe', 'on vine', 'san marzano'] },
  { category: 'Herbs & Spices', keywords: ['basil', 'bay leaf', 'bay leaves', 'chervil', 'chive', 'cilantro', 'dill', 'fennel', 'garlic scape', 'lemongrass', 'marjoram', 'mint', 'oregano', 'parsley', 'rosemary', 'sage', 'tarragon', 'thyme', 'dried mushroom', 'dried pepper', 'ancho', 'chipotle', 'guajillo', 'habanero', 'serrano', 'edible flower'] },
  { category: 'Greens', keywords: ['arugula', 'bok choy', 'cabbage', 'chicory', 'collard', 'dandelion', 'escarole', 'endive', 'iceberg', 'kale', 'lettuce', 'romaine', 'spinach', 'spring mix', 'mesclun', 'swiss chard', 'chard', 'green leaf', 'red leaf', 'boston lettuce', 'baby gem'] },
  { category: 'Vegetables', keywords: ['artichoke', 'asparagus', 'string bean', 'green bean', 'wax bean', 'haricot', 'beet', 'broccoli', 'broccolini', 'broccoli rabe', 'brussels sprout', 'carrot', 'celery', 'cauliflower', 'corn', 'cucumber', 'kirby', 'eggplant', 'fava', 'kohlrabi', 'leek', 'mushroom', 'onion', 'shallot', 'pea', 'pepper', 'peppers', 'radicchio', 'radish', 'squash', 'zucchini', 'turnip', 'parsnip', 'sunchoke', 'jerusalem artichoke', 'rhubarb', 'snap pea', 'snow pea'] },
  { category: 'Potatoes', keywords: ['potato', 'potatoes', 'yam', 'sweet potato', 'fingerling', 'yukon', 'idaho', 'russet', 'red bliss'] },
  { category: 'Fruit', keywords: ['apple', 'avocado', 'banana', 'berry', 'blueberry', 'blackberry', 'raspberry', 'strawberry', 'citrus', 'clementine', 'grapefruit', 'lemon', 'lime', 'mandarin', 'orange', 'tangerine', 'mango', 'melon', 'cantaloupe', 'honeydew', 'watermelon', 'peach', 'nectarine', 'plum', 'pear', 'pineapple', 'grape ', 'kiwi', 'fig', 'pomegranate', 'passion fruit', 'papaya', 'guava', 'coconut', 'cherry '] },
  { category: 'Organic', keywords: ['organic'] },
];

function categorizeItem(itemName) {
  const lower = itemName.toLowerCase();
  for (const { category, keywords } of ITEM_CATEGORIES) {
    if (keywords.some(kw => lower.includes(kw))) return category;
  }
  return 'Other';
}

// ─── Order parser (text) ──────────────────────────────────────────────────────

function parseOrder(text) {
  const lines = text.trim().split('\n').map(l => l.trim()).filter(Boolean);
  if (lines.length < 1) return null;

  const itemPattern = /^(\d+)\s+(.+)$/;
  const items = [];
  let customerName = null;

  for (let i = 0; i < lines.length; i++) {
    const match = lines[i].match(itemPattern);
    if (match) {
      const itemName = match[2];
      items.push({ qty: parseInt(match[1], 10), item: itemName, category: categorizeItem(itemName) });
    } else if (i === 0 && items.length === 0) {
      customerName = lines[i];
    }
  }

  if (items.length === 0) return null;
  return { customerName, items };
}

// ─── Download image from Slack ────────────────────────────────────────────────

function downloadImageAsBase64(url, token) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers: { Authorization: `Bearer ${token}` } }, res => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        // Follow redirect
        return downloadImageAsBase64(res.headers.location, token).then(resolve).catch(reject);
      }
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve({
        data: Buffer.concat(chunks).toString('base64'),
        mediaType: res.headers['content-type'] || 'image/jpeg',
      }));
    });
    req.on('error', reject);
  });
}

// ─── Extract order from image via Claude ─────────────────────────────────────

async function extractOrderFromImage(imageBase64, mediaType) {
  const response = await anthropic.messages.create({
    model: 'claude-opus-4-5',
    max_tokens: 1024,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'image',
          source: { type: 'base64', media_type: mediaType, data: imageBase64 },
        },
        {
          type: 'text',
          text: `This is a handwritten produce order. Extract the customer name and all order items from the image.
Return ONLY the order in this exact plain text format with nothing else:
CustomerName
Qty Item
Qty Item

Rules:
- First line is the customer/business name
- Each following line starts with a number then the item name
- If you can't read something clearly, do your best guess
- No extra commentary, just the formatted order`,
        },
      ],
    }],
  });

  return response.content[0].text;
}

// ─── SSE broadcast ────────────────────────────────────────────────────────────

const sseClients = new Set();

function broadcast() {
  const data = `data: ${JSON.stringify(loadOrders())}\n\n`;
  for (const res of sseClients) {
    try { res.write(data); } catch { sseClients.delete(res); }
  }
}

// ─── Core message processor ───────────────────────────────────────────────────

async function processMessage(message, channelId, client) {
  if (message.subtype && message.subtype !== 'file_share') return;

  const ts = message.ts;
  const type = typeForChannel(channelId);

  // ── Handle image attachments ──
  if (message.files && message.files.length > 0) {
    for (const file of message.files) {
      if (!file.mimetype || !file.mimetype.startsWith('image/')) continue;

      try {
        console.log(`[${new Date().toLocaleTimeString()}] Reading handwritten order image from ${file.name}...`);
        const { data, mediaType } = await downloadImageAsBase64(
          file.url_private,
          process.env.SLACK_BOT_TOKEN
        );
        const extractedText = await extractOrderFromImage(data, mediaType);
        console.log(`Extracted from image:\n${extractedText}`);

        const parsed = parseOrder(extractedText);
        if (!parsed) continue;

        let senderName = 'Unknown';
        try {
          const info = await client.users.info({ user: message.user });
          senderName = info.user.real_name || info.user.profile?.display_name || info.user.name;
        } catch { /* non-fatal */ }

        const rawName = parsed.customerName || senderName;
        const orders = loadOrders();
        const customerName = findCustomerKey(orders, rawName) || rawName;
        if (!orders[customerName]) orders[customerName] = [];

        const imgTs = ts + '_img_' + file.id;
        if (orders[customerName].some(o => o.ts === imgTs || (o.ts_all || []).includes(imgTs))) continue;

        const slackLink = slackTeamDomain
          ? `https://${slackTeamDomain}.slack.com/archives/${channelId}/p${ts.replace('.', '')}`
          : null;

        const msgDate = new Date(parseFloat(ts) * 1000).toDateString();
        const existingToday = orders[customerName].find(o => new Date(o.timestamp).toDateString() === msgDate);

        if (existingToday) {
          for (const newItem of parsed.items) {
            const already = existingToday.items.some(i => i.qty === newItem.qty && i.item.toLowerCase() === newItem.item.toLowerCase());
            if (!already) existingToday.items.push(newItem);
          }
          if (!existingToday.slackLinks) existingToday.slackLinks = [existingToday.slackLink].filter(Boolean);
          if (slackLink && !existingToday.slackLinks.includes(slackLink)) existingToday.slackLinks.push(slackLink);
          existingToday.ts_all = [...(existingToday.ts_all || [existingToday.ts]), imgTs];
        } else {
          orders[customerName].push({
            id: Date.now() + Math.floor(Math.random() * 1000),
            ts: imgTs,
            ts_all: [imgTs],
            timestamp: new Date(parseFloat(ts) * 1000).toISOString(),
            channel: channelId,
            type,
            source: 'image',
            sender: senderName,
            slackLink,
            slackLinks: slackLink ? [slackLink] : [],
            items: parsed.items,
          });
        }

        saveOrders(orders);
        broadcast();
        console.log(`[${new Date().toLocaleTimeString()}] Image order saved — ${customerName}: ${parsed.items.map(i => `${i.qty}x ${i.item}`).join(', ')}`);
      } catch (e) {
        console.error('Failed to process image order:', e.message);
      }
    }
    return;
  }

  // ── Handle text messages ──
  const parsed = parseOrder(message.text || '');
  if (!parsed) return;

  let senderName = 'Unknown';
  try {
    const info = await client.users.info({ user: message.user });
    senderName = info.user.real_name || info.user.profile?.display_name || info.user.name;
  } catch { /* non-fatal */ }

  const rawName = parsed.customerName || senderName;
  const orders = loadOrders();
  const customerName = findCustomerKey(orders, rawName) || rawName;
  if (!orders[customerName]) orders[customerName] = [];

  if (orders[customerName].some(o => o.ts === ts)) return;

  const slackLink = slackTeamDomain
    ? `https://${slackTeamDomain}.slack.com/archives/${channelId}/p${ts.replace('.', '')}`
    : null;

  // If the customer name contains "add", merge into their most recent order
  const isAddOn = /\badd\b/i.test(customerName);

  if (isAddOn) {
    const baseName = extractBaseName(customerName);
    const baseKey = findBaseKey(orders, baseName);
    const baseOrders = baseKey ? orders[baseKey] : null;

    if (baseOrders && baseOrders.length > 0) {
      // Merge into the most recent order for that customer
      const mostRecent = baseOrders[baseOrders.length - 1];
      for (const newItem of parsed.items) {
        const already = mostRecent.items.some(i => i.qty === newItem.qty && i.item.toLowerCase() === newItem.item.toLowerCase());
        if (!already) mostRecent.items.push(newItem);
      }
      if (!mostRecent.slackLinks) mostRecent.slackLinks = [mostRecent.slackLink].filter(Boolean);
      if (slackLink && !mostRecent.slackLinks.includes(slackLink)) mostRecent.slackLinks.push(slackLink);
      mostRecent.ts_all = [...(mostRecent.ts_all || [mostRecent.ts]), ts];
      // Remove the empty "add" entry we created above before saving
      if (orders[customerName] && orders[customerName].length === 0) delete orders[customerName];
      saveOrders(orders);
      broadcast();
      console.log(`[${new Date().toLocaleTimeString()}] ADD-ON merged into ${baseKey}: ${parsed.items.map(i => `${i.qty}x ${i.item}`).join(', ')}`);
      return;
    }
  }

  // New order
  orders[customerName].push({
    id: Date.now() + Math.floor(Math.random() * 1000),
    ts,
    ts_all: [ts],
    timestamp: new Date(parseFloat(ts) * 1000).toISOString(),
    channel: channelId,
    type,
    source: 'text',
    sender: senderName,
    slackLink,
    slackLinks: slackLink ? [slackLink] : [],
    items: parsed.items,
  });

  saveOrders(orders);
  broadcast();
  console.log(`[${new Date().toLocaleTimeString()}] ${type.toUpperCase()} order — ${customerName}: ${parsed.items.map(i => `${i.qty}x ${i.item}`).join(', ')}`);
}

// ─── One-time migration: merge existing "add" orders into base customers ───────

function extractBaseName(customerName) {
  // Handle "Add to Farmingdale" → "Farmingdale"
  // Handle "Farmingdale add" → "Farmingdale"
  return customerName
    .replace(/^\s*add\s+to\s+/i, '')   // strip leading "add to"
    .replace(/\s*\badd\b\s*$/i, '')    // strip trailing "add"
    .trim();
}

function findBaseKey(orders, baseName) {
  const lower = baseName.toLowerCase();
  const keys = Object.keys(orders).filter(k => !/\badd\b/i.test(k)); // exclude other "add" entries

  // 1. Exact match (case-insensitive)
  const exact = keys.find(k => k.toLowerCase() === lower);
  if (exact) return exact;

  // 2. Partial match — base name is contained in customer name (e.g. "mangia" matches "mangia pizza")
  const partial = keys.find(k => k.toLowerCase().includes(lower) || lower.includes(k.toLowerCase()));
  return partial || null;
}

// Find existing customer key with lenient matching (case, spaces, typos)
function findCustomerKey(orders, name) {
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const nameNorm = norm(name);
  const keys = Object.keys(orders).filter(k => !/\badd\b/i.test(k));

  // Exact case-insensitive
  const exact = keys.find(k => k.toLowerCase() === name.toLowerCase());
  if (exact) return exact;

  // Normalized (strip punctuation/spaces) — catches "Shop Rite" vs "ShopRite"
  const stripped = keys.find(k => norm(k) === nameNorm);
  if (stripped) return stripped;

  // One contains the other (short name ≥ 4 chars to avoid false matches)
  if (nameNorm.length >= 4) {
    const partial = keys.find(k => {
      const kn = norm(k);
      return kn.includes(nameNorm) || nameNorm.includes(kn);
    });
    if (partial) return partial;
  }

  return null;
}

// On startup: merge duplicate customer entries that differ only by case/spacing
function mergeDuplicateCustomers() {
  const orders = loadOrders();
  let changed = false;
  const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const seen = {}; // normKey → canonical customer name

  for (const name of Object.keys(orders)) {
    if (/\badd\b/i.test(name)) continue;
    const key = norm(name);
    if (seen[key] && seen[key] !== name) {
      // Merge this duplicate into the canonical entry
      orders[seen[key]].push(...orders[name]);
      orders[seen[key]].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
      delete orders[name];
      changed = true;
      console.log(`Merged duplicate customer "${name}" → "${seen[key]}"`);
    } else {
      seen[key] = name;
    }
  }

  if (changed) {
    saveOrders(orders);
    console.log('Duplicate customer merge complete.');
  }
}

function backfillCategories() {
  const orders = loadOrders();
  let changed = false;

  for (const orderList of Object.values(orders)) {
    for (const order of orderList) {
      if (!order.items) continue;
      for (const item of order.items) {
        if (!item.category) {
          item.category = categorizeItem(item.item);
          changed = true;
        }
      }
    }
  }

  if (changed) {
    saveOrders(orders);
    console.log('Backfilled item categories for existing orders.');
  }
}

function migrateAddOrders() {
  const orders = loadOrders();
  let changed = false;

  for (const customerName of Object.keys(orders)) {
    if (!/\badd\b/i.test(customerName)) continue;

    const baseName = extractBaseName(customerName);
    const baseKey = findBaseKey(orders, baseName);
    if (!baseKey || orders[baseKey].length === 0) {
      console.log(`No base customer found for "${customerName}" (looking for "${baseName}")`);
      continue;
    }

    const mostRecent = orders[baseKey][orders[baseKey].length - 1];

    for (const addOrder of orders[customerName]) {
      for (const newItem of addOrder.items) {
        const already = mostRecent.items.some(i =>
          i.qty === newItem.qty && i.item.toLowerCase() === newItem.item.toLowerCase()
        );
        if (!already) mostRecent.items.push(newItem);
      }
    }

    delete orders[customerName];
    changed = true;
    console.log(`Migrated "${customerName}" into "${baseKey}"`);
  }

  if (changed) saveOrders(orders);
}

// ─── Backfill today's orders ──────────────────────────────────────────────────

async function backfillToday(client) {
  console.log("Backfilling today's orders from Slack...");
  const startOfDay = new Date();
  startOfDay.setHours(0, 0, 0, 0);
  const oldest = String(startOfDay.getTime() / 1000);

  let channels = [];
  try {
    let cursor;
    do {
      const res = await client.conversations.list({
        types: 'public_channel,private_channel,im,mpim',
        limit: 200,
        cursor,
      });
      channels = channels.concat(res.channels);
      cursor = res.response_metadata?.next_cursor;
    } while (cursor);
  } catch (e) {
    console.error('Could not list channels:', e.message);
    return;
  }

  // Collect all messages first, then process oldest-first so "add" orders
  // are always processed after their base customer order exists
  const allMessages = [];
  for (const channel of channels) {
    try {
      const history = await client.conversations.history({ channel: channel.id, oldest, limit: 200 });
      for (const msg of (history.messages || [])) {
        allMessages.push({ msg, channelId: channel.id });
      }
    } catch { /* skip channels bot can't access */ }
  }

  // Sort oldest first (ts is a unix timestamp string)
  allMessages.sort((a, b) => parseFloat(a.msg.ts) - parseFloat(b.msg.ts));

  for (const { msg, channelId } of allMessages) {
    await processMessage(msg, channelId, client);
  }

  console.log(`Backfill complete — scanned ${channels.length} channels.`);
}

// ─── Slack bot ────────────────────────────────────────────────────────────────

const slackApp = new App({
  token: process.env.SLACK_BOT_TOKEN,
  appToken: process.env.SLACK_APP_TOKEN,
  socketMode: true,
});

slackApp.message(async ({ message, client }) => {
  await processMessage(message, message.channel, client);
});

// ─── Web server ───────────────────────────────────────────────────────────────

const web = express();
web.use(express.json());
web.use(express.static(path.join(__dirname, '..', 'public')));

web.get('/api/orders', (_req, res) => res.json(loadOrders()));
web.get('/api/channels', (_req, res) => res.json(loadChannels()));

web.get('/api/orders/stream', (req, res) => {
  res.set({ 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' });
  res.flushHeaders();
  res.write(`data: ${JSON.stringify(loadOrders())}\n\n`);
  sseClients.add(res);
  req.on('close', () => sseClients.delete(res));
});

web.patch('/api/orders/:customer/:orderId/type', (req, res) => {
  const { customer, orderId } = req.params;
  const { type } = req.body;
  const orders = loadOrders();
  if (orders[customer]) {
    const order = orders[customer].find(o => o.id === parseInt(orderId, 10));
    if (order) { order.type = type; saveOrders(orders); broadcast(); }
  }
  res.json({ ok: true });
});

web.delete('/api/orders/:customer/:orderId', (req, res) => {
  const orders = loadOrders();
  const { customer, orderId } = req.params;
  if (orders[customer]) {
    orders[customer] = orders[customer].filter(o => o.id !== parseInt(orderId, 10));
    if (orders[customer].length === 0) delete orders[customer];
    saveOrders(orders); broadcast();
  }
  res.json({ ok: true });
});

web.delete('/api/orders/:customer', (req, res) => {
  const orders = loadOrders();
  delete orders[req.params.customer];
  saveOrders(orders); broadcast();
  res.json({ ok: true });
});

web.post('/api/channels', (req, res) => {
  saveChannels(req.body);
  const orders = loadOrders();
  for (const customer of Object.keys(orders)) {
    for (const order of orders[customer]) {
      order.type = typeForChannel(order.channel);
    }
  }
  saveOrders(orders); broadcast();
  res.json({ ok: true });
});

const PORT = process.env.PORT || 3000;
web.listen(PORT, () => console.log(`Dashboard → http://localhost:${PORT}`));

(async () => {
  await initDb();
  await slackApp.start();
  console.log('Slack bot connected — listening for orders and images...');
  try {
    const auth = await slackApp.client.auth.test();
    slackTeamDomain = auth.url.replace('https://', '').replace('.slack.com/', '');
    console.log(`Slack workspace: ${slackTeamDomain}`);
  } catch (e) {
    console.error('Could not get team domain:', e.message);
  }
  mergeDuplicateCustomers();
  migrateAddOrders();
  backfillCategories();
  await backfillToday(slackApp.client);
})();
