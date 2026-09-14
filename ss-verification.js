const crypto = require('crypto');
const { EmbedBuilder } = require('discord.js');
const { getGuildStore, saveGuildStore } = require('./storage');
const { resolveLogChannel } = require('./log-channel');

// Vision-capable model on Groq's free-tier API (same key/endpoint as
// ai-chat.js). Configurable via .env in case Groq renames/retires this
// model later — no code changes needed.
const VISION_MODEL = process.env.GROQ_VISION_MODEL || 'qwen/qwen3.6-27b';

// One team member submits proof for the whole squad in a single message:
// their own follow screenshot plus one from each of 3 teammates — all 4
// showing they follow the SAME configured account, and all 4 must be
// genuinely separate screenshots (see duplicate checks below).
const REQUIRED_SCREENSHOT_COUNT = 4;

// Only these get sent off to the vision model — random file attachments
// (zip, txt, etc.) in the channel are ignored rather than misread as "no
// screenshot" screenshots.
function imageAttachments(message) {
  return [...message.attachments.values()].filter(a =>
    (a.contentType || '').startsWith('image/')
  );
}

// Catches the laziest version of reusing one screenshot for all 4 slots:
// literally the same image file (or an exact re-upload of it) attached
// more than once. Byte-identical only — this is a cheap pre-check before
// the AI call, not a replacement for the AI's own near-duplicate check
// below (which also catches a re-saved/recompressed copy).
async function hashImage(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch image for hashing: ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  return crypto.createHash('sha256').update(buf).digest('hex');
}

async function findExactDuplicates(imageUrls) {
  const hashes = await Promise.all(imageUrls.map(hashImage));
  const firstSeenAt = new Map(); // hash -> first 0-based index
  const pairs = []; // [imageNumber1, imageNumber2] (1-based, for the reply)
  hashes.forEach((hash, i) => {
    if (firstSeenAt.has(hash)) {
      pairs.push([firstSeenAt.get(hash) + 1, i + 1]);
    } else {
      firstSeenAt.set(hash, i);
    }
  });
  return pairs;
}

function buildPrompt(instagramUsername) {
  return `You are checking ${REQUIRED_SCREENSHOT_COUNT} screenshots submitted together as proof that a player AND their 3 teammates are all following the same Instagram account: "@${instagramUsername}".

You are shown ${REQUIRED_SCREENSHOT_COUNT} images, in the order they were submitted (image 1, image 2, image 3, image 4). Each one is meant to be a SEPARATE person's Instagram app, open on the profile page of "@${instagramUsername}", showing a "Following" state on the follow button.

For EACH image, decide:
- whether the username shown matches "${instagramUsername}" (a leading "@" or minor case difference is fine)
- whether the follow button shows "Following" (already followed) rather than "Follow" (not yet followed)
- whether it looks like a genuine Instagram profile screenshot — not an unrelated image, a different app, or an obviously edited/mocked-up button

Be strict: if any image's username doesn't match, or its button says "Follow" instead of "Following", or you can't clearly tell, that image does NOT count.

ALSO compare the 4 images to each other: someone might try to pass off one real screenshot as if it were 4 separate people by submitting the same screenshot more than once — possibly resized, re-compressed, lightly cropped, or with a filter/watermark added, so it isn't byte-identical anymore but is clearly the same underlying screenshot (same status bar time, same notification icons, same scroll position, same visible content). If any two (or more) of the 4 images look like the same screenshot reused rather than genuinely separate captures, set "duplicate_images" to true and this submission is NOT verified, even if each image individually looks fine.

This submission is verified ONLY if ALL ${REQUIRED_SCREENSHOT_COUNT} images individually show "@${instagramUsername}" with the "Following" state, AND none of them are duplicates/near-duplicates of each other.

Respond with ONLY a JSON object, nothing else:
{"verified": true or false, "per_image": [true or false, true or false, true or false, true or false], "duplicate_images": true or false, "reason": "<one short sentence explaining the decision, noting which image number(s) failed or were duplicates if any>"}`;
}

/**
 * Calls the Groq vision model with the submitted images. Returns
 * { ok: true, verified, perImage, duplicateImages, reason } on a
 * successful call, or { ok: false, error } if the model couldn't be
 * reached/parsed — the caller treats that as "couldn't verify right now",
 * never as a pass.
 */
async function analyzeFollowScreenshot(imageUrls, instagramUsername) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    console.warn('[ss-verify] GROQ_API_KEY not set — skipping screenshot verification.');
    return { ok: false, error: 'no_api_key' };
  }

  const content = [
    { type: 'text', text: buildPrompt(instagramUsername) },
    // Groq vision accepts image URLs directly server-side — Discord's CDN
    // attachment URLs work fine here without downloading/re-uploading.
    ...imageUrls.map(url => ({ type: 'image_url', image_url: { url } })),
  ];

  try {
    const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: VISION_MODEL,
        messages: [{ role: 'user', content }],
        max_completion_tokens: 350,
        temperature: 0,
        // NOTE: deliberately NOT using response_format: json_object here.
        // Combined with image input, this model can return Groq's
        // "json_validate_failed" 400 error instead of a completion (the
        // strict JSON-schema decoder rejects some of what the vision
        // model generates). The prompt already asks for JSON-only output,
        // so we just parse that leniently below instead.
      }),
    });

    if (!res.ok) {
      const errBody = await res.text().catch(() => '');
      console.error(`[ss-verify] Groq API returned ${res.status}:`, errBody);
      return { ok: false, error: 'api_error' };
    }

    const data = await res.json();
    const text = data?.choices?.[0]?.message?.content?.trim();
    if (!text) {
      console.error('[ss-verify] Groq API response had no text:', JSON.stringify(data));
      return { ok: false, error: 'empty_response' };
    }

    // The model sometimes wraps the JSON in a code fence or adds a stray
    // sentence before/after it despite the prompt — pull out just the
    // {...} object instead of requiring the whole response to be clean JSON.
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    let parsed;
    try {
      parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text);
    } catch (err) {
      console.error('[ss-verify] Failed to parse JSON from vision model:', text);
      return { ok: false, error: 'bad_json' };
    }

    const perImage = Array.isArray(parsed.per_image)
      ? parsed.per_image.map(v => v === true)
      : [];
    const duplicateImages = parsed.duplicate_images === true;

    return {
      ok: true,
      // Trust the model's own "verified" verdict, but also require it to
      // have actually marked all 4 individual images true and no
      // duplicates flagged — guards against the top-level flag and the
      // detailed breakdown disagreeing.
      verified: parsed.verified === true
        && perImage.length === REQUIRED_SCREENSHOT_COUNT
        && perImage.every(Boolean)
        && !duplicateImages,
      perImage,
      duplicateImages,
      reason: typeof parsed.reason === 'string' ? parsed.reason : '',
    };
  } catch (err) {
    console.error('[ss-verify] Failed to reach Groq API:', err);
    return { ok: false, error: 'network_error' };
  }
}

// Gives the configured "SS verified" role, if one is set. Never throws —
// mirrors giveVerifiedRole in verification-handlers.js so a missing role,
// missing permission, or the member having left doesn't block the rest of
// the flow.
async function giveSsVerifiedRole(message, store) {
  const roleId = store.settings && store.settings.ssVerifyRoleId;
  if (!roleId) {
    console.warn(`[ss-verify] No ssVerifyRoleId configured for guild ${message.guild.id} — run /set-ss-verify-role.`);
    return;
  }

  const role = message.guild.roles.cache.get(roleId);
  if (!role) {
    console.error(`[ss-verify] Configured role ${roleId} no longer exists in guild ${message.guild.id} — re-run /set-ss-verify-role.`);
    return;
  }

  const botMember = message.guild.members.me;
  if (!botMember.permissions.has('ManageRoles')) {
    console.error(`[ss-verify] Bot is missing the "Manage Roles" permission in guild ${message.guild.id}.`);
    return;
  }
  if (role.position >= botMember.roles.highest.position) {
    console.error(`[ss-verify] Bot's highest role is below "${role.name}" (${roleId}) in guild ${message.guild.id} — move the bot's role above it.`);
    return;
  }

  try {
    const member = message.member ?? await message.guild.members.fetch(message.author.id);
    if (!member.roles.cache.has(roleId)) {
      await member.roles.add(roleId);
    }
  } catch (err) {
    console.error(`[ss-verify] Failed to give role ${roleId} to ${message.author.id} in guild ${message.guild.id}: ${err.code ?? ''} ${err.message}`);
  }
}

function buildLogEmbed(message, instagramUsername, result, thumbnailUrl) {
  return new EmbedBuilder()
    .setTitle('📸 Instagram Follow — Squad Verified')
    .setColor(0x57F287)
    .setDescription(`${message.author} submitted ${REQUIRED_SCREENSHOT_COUNT} distinct screenshots (themself + 3 teammates) all following **@${instagramUsername}**.`)
    .addFields(
      { name: 'Model reasoning', value: result.reason || '_none given_' },
    )
    .setThumbnail(thumbnailUrl)
    .setTimestamp();
}

// Entry point, called from index.js for every message posted in the
// configured SS-verify channel. Returns true if this message was handled
// (so the caller knows not to fall through to anything else), false if it
// wasn't an SS-verify submission at all.
async function handleSsVerifyMessage(message) {
  const store = getGuildStore(message.guild.id);
  if (!store.settings) store.settings = {};

  const instagramUsername = store.settings.instagramUsername;
  if (!instagramUsername) {
    await message.reply({
      content: '⚠️ Screenshot verification isn\'t fully set up yet — an admin needs to run `/set-instagram-username` first.',
    }).catch(() => {});
    return true;
  }

  if (!store.ssVerifications) store.ssVerifications = {};
  const existing = store.ssVerifications[message.author.id];
  if (existing) {
    await message.reply({
      content: `✅ You're already verified as following **@${instagramUsername}**.`,
    }).catch(() => {});
    return true;
  }

  const images = imageAttachments(message);

  if (images.length === 0) {
    await message.reply({
      content: `📸 Post **${REQUIRED_SCREENSHOT_COUNT} screenshots together in one message** — yours plus your **3 teammates'** (each a genuinely separate screenshot) — each showing you're following **@${instagramUsername}** (the "Following" button visible), to get verified.`,
    }).catch(() => {});
    return true;
  }

  if (images.length !== REQUIRED_SCREENSHOT_COUNT) {
    await message.reply({
      content: `❌ You attached ${images.length} screenshot(s), but exactly **${REQUIRED_SCREENSHOT_COUNT}** are required — yours plus your **3 teammates'**, all following **@${instagramUsername}**. Please resend all ${REQUIRED_SCREENSHOT_COUNT} together in a single message.`,
    }).catch(() => {});
    return true;
  }

  // Cheap pre-check before spending an AI call: catches the exact same
  // image file attached more than once (e.g. reusing one person's
  // screenshot to fill all 4 slots).
  let exactDupPairs = [];
  try {
    exactDupPairs = await findExactDuplicates(images.map(a => a.url));
  } catch (err) {
    console.error('[ss-verify] Failed to hash images for duplicate check:', err);
    // Don't block verification just because hashing failed — the AI's own
    // near-duplicate check below still runs as a backstop.
  }

  if (exactDupPairs.length) {
    const pairText = exactDupPairs.map(([a, b]) => `#${a} & #${b}`).join(', ');
    await message.reply({
      content: `❌ Screenshot(s) ${pairText} are the exact same image reused. Each of the ${REQUIRED_SCREENSHOT_COUNT} screenshots must be a genuinely different screenshot (yours + 3 teammates'). Please resend all ${REQUIRED_SCREENSHOT_COUNT} together.`,
    }).catch(() => {});
    return true;
  }

  await message.channel.sendTyping().catch(() => {});

  const result = await analyzeFollowScreenshot(images.map(a => a.url), instagramUsername);

  if (!result.ok) {
    await message.reply({
      content: '⚠️ Couldn\'t run verification right now (the AI checker is temporarily unavailable). Please try again shortly, or ping a staff member if this keeps happening.',
    }).catch(() => {});
    return true;
  }

  if (!result.verified) {
    const failedNums = result.perImage
      .map((ok, i) => (ok ? null : i + 1))
      .filter(Boolean);
    const failedLine = failedNums.length
      ? `\nScreenshot(s) #${failedNums.join(', #')} didn't check out.`
      : '';
    const dupLine = result.duplicateImages
      ? '\nTwo or more of these look like the same screenshot reused — each of the 4 must be a genuinely separate screenshot.'
      : '';
    const reasonLine = result.reason ? `\n> ${result.reason}` : '';
    await message.reply({
      content: `❌ Couldn't confirm all ${REQUIRED_SCREENSHOT_COUNT} screenshots show following **@${instagramUsername}**.${reasonLine}${failedLine}${dupLine}\nMake sure every screenshot is a separate, genuine capture showing that profile with the **Following** button, then resend all ${REQUIRED_SCREENSHOT_COUNT} together.`,
    }).catch(() => {});
    return true;
  }

  store.ssVerifications[message.author.id] = {
    verifiedAt: new Date().toISOString(),
    imageUrls: images.map(a => a.url),
  };
  saveGuildStore(message.guild.id, store);

  await giveSsVerifiedRole(message, store);

  await message.react('✅').catch(() => {});
  await message.reply({
    content: `✅ Verified! You and your 3 teammates are confirmed as following **@${instagramUsername}**.`,
  }).catch(() => {});

  const logChannel = await resolveLogChannel(message.guild, store, store.settings.ssVerifyLogChannelId);
  if (logChannel) {
    try {
      await logChannel.send({ embeds: [buildLogEmbed(message, instagramUsername, result, images[0].url)] });
    } catch (err) {
      console.error('[ss-verify] Failed to post to log channel:', err);
    }
  }

  return true;
}

module.exports = { handleSsVerifyMessage, analyzeFollowScreenshot, REQUIRED_SCREENSHOT_COUNT };
