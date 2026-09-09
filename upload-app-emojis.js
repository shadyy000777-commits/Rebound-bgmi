// scripts/upload-app-emojis.js
//
// Uploads image files as APPLICATION-owned emojis (the same mechanism bots
// like "Quo" use to show custom emoji identically in every server they're
// in — no per-server upload needed, and USE_EXTERNAL_EMOJIS isn't required).
//
// How it works:
//   Discord lets a bot APPLICATION (not a guild) own up to 2000 emojis.
//   Once uploaded, they can be used by the bot anywhere via the normal
//   custom-emoji tag format: <:name:id> (or <a:name:id> if animated).
//
// Setup:
//   1. Drop image files (png/jpg/gif/webp) into ./emoji-assets/
//      The filename (without extension) becomes the emoji name, e.g.
//      emoji-assets/check.png -> emoji named "check"
//   2. Make sure your .env has DISCORD_TOKEN set (same token the bot uses).
//   3. Run:  node scripts/upload-app-emojis.js
//   4. Copy the printed <:name:id> tags — paste them wherever you want a
//      consistent icon (e.g. into the tournament wizard's "Reactions"
//      modal for the accept/deny fields).
//
// This is a one-time REST-only script — it does NOT log the bot into the
// gateway, so it's safe to run any time, including while the bot is live.

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { REST } = require('discord.js');

const TOKEN = process.env.DISCORD_TOKEN;
const ASSETS_DIR = path.join(__dirname, '..', 'emoji-assets');
const VALID_EXT = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp']);
const MIME_BY_EXT = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
};

function isValidEmojiName(name) {
  // Discord emoji names: 2-32 chars, letters/numbers/underscores only.
  return /^\w{2,32}$/.test(name);
}

async function main() {
  if (!TOKEN) {
    console.error('❌ DISCORD_TOKEN is not set in your .env file.');
    process.exit(1);
  }

  if (!fs.existsSync(ASSETS_DIR)) {
    console.error(`❌ Folder not found: ${ASSETS_DIR}\nCreate it and add images first.`);
    process.exit(1);
  }

  const files = fs.readdirSync(ASSETS_DIR)
    .filter((f) => VALID_EXT.has(path.extname(f).toLowerCase()));

  if (!files.length) {
    console.error(`❌ No image files found in ${ASSETS_DIR}\nSupported: ${[...VALID_EXT].join(', ')}`);
    process.exit(1);
  }

  const rest = new REST({ version: '10' }).setToken(TOKEN);

  // Application emojis are owned by the app, not the bot user directly —
  // fetch the application id via the same token.
  const me = await rest.get('/oauth2/applications/@me').catch((err) => {
    console.error('❌ Could not resolve application id from token:', err.message);
    process.exit(1);
  });
  const applicationId = me.id;
  console.log(`Application: ${me.name} (${applicationId})\n`);

  const results = [];

  for (const file of files) {
    const ext = path.extname(file).toLowerCase();
    const name = path.basename(file, ext).toLowerCase().replace(/[^a-z0-9_]/g, '_');

    if (!isValidEmojiName(name)) {
      console.warn(`⚠️  Skipping "${file}" — name "${name}" isn't a valid emoji name (2-32 word chars).`);
      continue;
    }

    const filePath = path.join(ASSETS_DIR, file);
    const stats = fs.statSync(filePath);
    if (stats.size > 256 * 1024) {
      console.warn(`⚠️  Skipping "${file}" — ${(stats.size / 1024).toFixed(0)}KB exceeds the 256KB API limit (Discord's UI auto-compresses, the API doesn't).`);
      continue;
    }

    const base64 = fs.readFileSync(filePath).toString('base64');
    const dataUri = `data:${MIME_BY_EXT[ext]};base64,${base64}`;

    try {
      const emoji = await rest.post(`/applications/${applicationId}/emojis`, {
        body: { name, image: dataUri },
      });
      const tag = emoji.animated ? `<a:${emoji.name}:${emoji.id}>` : `<:${emoji.name}:${emoji.id}>`;
      results.push({ file, tag });
      console.log(`✅ Uploaded "${name}" -> ${tag}`);
    } catch (err) {
      console.error(`❌ Failed to upload "${file}": ${err.message}`);
    }
  }

  if (results.length) {
    console.log('\nDone. These tags work in every server the bot is in:');
    results.forEach((r) => console.log(`  ${r.file.padEnd(20)} ${r.tag}`));
    console.log('\nPaste a tag directly into any field that accepts a custom emoji, e.g. the');
    console.log('tournament wizard\'s "Reactions" (accept/deny) modal.');
  } else {
    console.log('\nNothing was uploaded.');
  }
}

main();
