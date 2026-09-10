/**
 * upload-emojis.js
 *
 * Uploads image/gif files as DISCORD APPLICATION EMOJIS — emojis that belong
 * to your bot itself, not to any one server. Application emojis render
 * correctly in EVERY server your bot is in (that's how bots like Quo do it),
 * unlike a normal server emoji which only works inside its home server.
 *
 * HOW TO USE
 * 1. Drop your .png / .jpg / .gif files into the emoji-assets/ folder.
 *    - Filename becomes the emoji name (letters, numbers, underscores only).
 *      "check_green.png" -> emoji name "check_green"
 *    - Max 256 KB per file. Discord recommends 128x128px.
 *    - GIFs become animated emojis automatically.
 * 2. Make sure DISCORD_TOKEN and CLIENT_ID are set in your .env (same ones
 *    the bot already uses).
 * 3. Run:
 *      node upload-emojis.js
 * 4. It uploads any NEW files it finds (skips ones already uploaded),
 *    then prints a table of ready-to-paste codes and writes them to
 *    emoji-codes.json and emoji-codes.md.
 *
 * USING THE RESULT
 * Paste the printed code straight into any embed/button text, e.g.:
 *      .setDescription('<:check_green:1234567890123456789> Registered!')
 * Animated ones look like: <a:party_parrot:1234567890123456789>
 * It will work the same way in every server the bot is added to.
 *
 * Re-running the script later with new files in emoji-assets/ only
 * uploads the new ones — existing ones are left alone.
 */

require('dotenv').config();
const fs = require('fs');
const path = require('path');

const TOKEN = process.env.DISCORD_TOKEN;
const CLIENT_ID = process.env.CLIENT_ID;
const ASSETS_DIR = path.join(__dirname, 'emoji-assets');
const OUT_JSON = path.join(__dirname, 'emoji-codes.json');
const OUT_MD = path.join(__dirname, 'emoji-codes.md');
const API = 'https://discord.com/api/v10';

const MIME_BY_EXT = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
};

function sanitizeName(rawName) {
  let name = rawName.toLowerCase().replace(/[^a-z0-9_]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
  if (name.length < 2) name = name.padEnd(2, '_');
  if (name.length > 32) name = name.slice(0, 32);
  return name;
}

async function discordRequest(method, endpoint, body) {
  const res = await fetch(`${API}${endpoint}`, {
    method,
    headers: {
      Authorization: `Bot ${TOKEN}`,
      'Content-Type': 'application/json',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let data;
  try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
  if (!res.ok) {
    const err = new Error(`Discord API ${res.status} on ${method} ${endpoint}: ${JSON.stringify(data)}`);
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

async function getApplicationId() {
  if (CLIENT_ID) return CLIENT_ID;
  const app = await discordRequest('GET', '/applications/@me');
  return app.id;
}

async function listExistingEmojis(appId) {
  const data = await discordRequest('GET', `/applications/${appId}/emojis`);
  return data.items || [];
}

async function uploadEmoji(appId, name, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mime = MIME_BY_EXT[ext];
  if (!mime) throw new Error(`Unsupported file type: ${ext}`);
  const buffer = fs.readFileSync(filePath);
  if (buffer.length > 256 * 1024) {
    throw new Error(`File too large (${(buffer.length / 1024).toFixed(0)} KB) — Discord's limit is 256 KB`);
  }
  const base64 = buffer.toString('base64');
  const image = `data:${mime};base64,${base64}`;
  return discordRequest('POST', `/applications/${appId}/emojis`, { name, image });
}

function codeFor(emoji) {
  return emoji.animated ? `<a:${emoji.name}:${emoji.id}>` : `<:${emoji.name}:${emoji.id}>`;
}

async function main() {
  if (!TOKEN) {
    console.error('Missing DISCORD_TOKEN in your .env file.');
    process.exit(1);
  }
  if (!fs.existsSync(ASSETS_DIR)) {
    fs.mkdirSync(ASSETS_DIR, { recursive: true });
  }

  const files = fs.readdirSync(ASSETS_DIR).filter((f) => MIME_BY_EXT[path.extname(f).toLowerCase()]);
  if (files.length === 0) {
    console.log(`No image/gif files found in ${ASSETS_DIR}`);
    console.log('Drop your emoji/gif files in there, then run this script again.');
    return;
  }

  const appId = await getApplicationId();
  const existing = await listExistingEmojis(appId);
  const existingByName = new Map(existing.map((e) => [e.name, e]));

  const results = [];

  for (const file of files) {
    const filePath = path.join(ASSETS_DIR, file);
    const baseName = sanitizeName(path.parse(file).name);

    if (existingByName.has(baseName)) {
      const e = existingByName.get(baseName);
      console.log(`SKIP  ${file} -> already uploaded as "${baseName}"`);
      results.push({ file, name: baseName, id: e.id, animated: !!e.animated, code: codeFor(e), status: 'already-uploaded' });
      continue;
    }

    try {
      const emoji = await uploadEmoji(appId, baseName, filePath);
      console.log(`OK    ${file} -> ${codeFor(emoji)}`);
      results.push({ file, name: baseName, id: emoji.id, animated: !!emoji.animated, code: codeFor(emoji), status: 'uploaded' });
    } catch (err) {
      console.log(`FAIL  ${file} -> ${err.message}`);
      results.push({ file, name: baseName, error: err.message, status: 'failed' });
    }
  }

  fs.writeFileSync(OUT_JSON, JSON.stringify(results, null, 2));

  const mdLines = [
    '# Emoji codes',
    '',
    'Paste any of these directly into embed/button text in your bot code.',
    '',
    '| File | Code | Preview text |',
    '|---|---|---|',
    ...results
      .filter((r) => r.code)
      .map((r) => `| ${r.file} | \`${r.code}\` | ${r.code} |`),
  ];
  if (results.some((r) => r.status === 'failed')) {
    mdLines.push('', '## Failed', '', ...results.filter((r) => r.status === 'failed').map((r) => `- ${r.file}: ${r.error}`));
  }
  fs.writeFileSync(OUT_MD, mdLines.join('\n'));

  console.log('\nDone.');
  console.log(`Codes saved to ${path.relative(__dirname, OUT_JSON)} and ${path.relative(__dirname, OUT_MD)}`);
}

main().catch((err) => {
  console.error('Script failed:', err.message);
  process.exit(1);
});
