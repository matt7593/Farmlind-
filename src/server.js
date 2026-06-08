require('dotenv').config();
const { App } = require('@slack/bolt');
const express = require('express');
const fs = require('fs');
const path = require('path');
const https = require('https');
const Anthropic = require('@anthropic-ai/sdk');

const ORDERS_FILE   = path.join(__dirname, '..', 'orders.json');
const CHANNELS_FILE = path.join(__dirname, '..', 'channels.json');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

let slackTeamDomain = ''; // e.g. "farmlindproduce" → used to build slack message links

// ─── Storage helpers ──────────────────────────────────────────────────────────

function loadOrders() {
  if (!fs.existsSync(ORDERS_FILE)) return {};
  try { return JSON.parse(fs.readFileSync(ORDERS_FILE, 'utf8')); } catch { return {}; }
}

function saveOrders(orders) {
  fs.writeFileSync(ORDERS_FILE, JSON.stringify(orders, null, 2));
}

function loadChannels() {
  try { return JSON.parse(fs.readFileSync(CHANNELS_FILE, 'utf8')); } catch { return { restaurant: [], grocery: [] }; }
}

function saveChannels(ch) {
  fs.writeFileSync(CHANNELS_FILE, JSON.stringify(ch, null, 2));
}

function typeForChannel(channelId) {
  const ch = loadChannels();
  if (ch.restaurant.includes(channelId)) return 'restaurant';
  if (ch.grocery.includes(channelId))    return 'grocery';
  return 'unknown';
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
      items.push({ qty: parseInt(match[1], 10), item: match[2] });
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

        const customerName = parsed.customerName || senderName;
        const orders = loadOrders();
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

  const customerName = parsed.customerName || senderName;
  const orders = loadOrders();
  if (!orders[customerName]) orders[customerName] = [];

  if (orders[customerName].some(o => o.ts === ts)) return;

  const slackLink = slackTeamDomain
    ? `https://${slackTeamDomain}.slack.com/archives/${channelId}/p${ts.replace('.', '')}`
    : null;

  // If the customer name contains "add", merge into their most recent order
  const isAddOn = /\badd\b/i.test(customerName);

  if (isAddOn) {
    const baseName = extractBaseName(customerName);
    const baseOrders = orders[baseName];

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
      saveOrders(orders);
      broadcast();
      console.log(`[${new Date().toLocaleTimeString()}] ADD-ON merged into ${baseName}: ${parsed.items.map(i => `${i.qty}x ${i.item}`).join(', ')}`);
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

function migrateAddOrders() {
  const orders = loadOrders();
  let changed = false;

  for (const customerName of Object.keys(orders)) {
    if (!/\badd\b/i.test(customerName)) continue;

    const baseName = extractBaseName(customerName);
    if (!orders[baseName] || orders[baseName].length === 0) continue;

    const mostRecent = orders[baseName][orders[baseName].length - 1];

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
    console.log(`Migrated "add" orders from "${customerName}" into "${baseName}"`);
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
  await slackApp.start();
  console.log('Slack bot connected — listening for orders and images...');
  try {
    const auth = await slackApp.client.auth.test();
    slackTeamDomain = auth.url.replace('https://', '').replace('.slack.com/', '');
    console.log(`Slack workspace: ${slackTeamDomain}`);
  } catch (e) {
    console.error('Could not get team domain:', e.message);
  }
  migrateAddOrders();
  await backfillToday(slackApp.client);
})();
