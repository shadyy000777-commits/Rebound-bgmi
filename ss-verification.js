const { EmbedBuilder } = require('discord.js');
const { getGuildStore, saveGuildStore } = require('./storage');
const { resolveLogChannel } = require('./log-channel');

// Vision-capable model on Groq's free-tier API (same key/endpoint as
// ai-chat.js). Configurable via .env in case Groq renames/retires this
// model later — no code changes needed.
const VISION_MODEL = process.env.GROQ_VISION_MODEL || 'qwen/qwen3.6-27b';

// Members must submit exactly one screenshot per required account, all in
// the same message, so the model can cross-check the full set at once
// instead of accumulating trust across separate submissions.
const REQUIRED_SCREENSHOT_COUNT = 4;

// Only these get sent off to the vision model — random file attachments
// (zip, txt, etc.) in the channel are ignored rather than misread as "no
// screenshot" screenshots.
function imageAttachments(message) {
  return [...message.attachments.values()].filter(a =>
    (a.contentType || '').startsWith('image/')
  );
}

function buildPrompt(instagramUsernames) {
  const list = instagramUsernames.map(u => `@${u}`).join(', ');
  return `You are checking ${REQUIRED_SCREENSHOT_COUNT} screenshots submitted as proof of following ALL ${REQUIRED_SCREENSHOT_COUNT} of these Instagram accounts: ${list}.

You are shown ${REQUIRED_SCREENSHOT_COUNT} images, in the order they were submitted (image 1, image 2, ...). For EACH image, work out:
- which one of the required accounts (if any) it is a profile screenshot of
- whether the follow button on that screenshot shows "Following" (already followed) rather than "Follow" (not yet followed)

Be strict:
- Each image's username must match one of the required accounts (a leading "@" or minor case difference is fine).
- Each must look like a genuine Instagram profile screenshot — not an unrelated image, a different app, or an obviously edited/mocked-up button.
- If a button says "Follow" (not "Following"/"Message"), or a username doesn't match any required account, or you can't clearly tell, do NOT count that image as satisfying that account.
- This submission is verified ONLY if all ${REQUIRED_SCREENSHOT_COUNT} required accounts are each covered by a distinct image showing "Following" — no required account missing, and the same account can't be used twice to cover two required accounts.

Respond with ONLY a JSON object, nothing else:
{"verified": true or false, "matched_usernames": ["<username read from image 1, or null>", "<image 2>", "<image 3>", "<image 4>"], "missing_accounts": ["<any required account with no matching Following screenshot>"], "reason": "<one short sentence explaining the decision>"}`;
}

/**
 * Calls the Groq vision model with the submitted images. Returns
 * { ok: true, verified, matchedUsernames, missingAccounts, reason } on a
 * successful call, or { ok: false, error } if the model couldn't be
 * reached/parsed — the caller treats that as "couldn't verify right now",
 * never as a pass.
 */
async function analyzeFollowScreenshot(imageUrls, instagramUsernames) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    console.warn('[ss-verify] GROQ_API_KEY not set — skipping screenshot verification.');
    return { ok: false, error: 'no_api_key' };
  }

  const content = [
    { type: 'text', text: buildPrompt(instagramUsernames) },
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

    return {
      ok: true,
      verified: parsed.verified === true,
      matchedUsernames: Array.isArray(parsed.matched_usernames)
        ? parsed.matched_usernames.map(u => (typeof u === 'string' ? u : null))
        : [],
      missingAccounts: Array.isArray(parsed.missing_accounts)
        ? parsed.missing_accounts.filter(u => typeof u === 'string')
        : [],
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

function buildLogEmbed(message, instagramUsernames, result, thumbnailUrl) {
  return new EmbedBuilder()
    .setTitle('📸 Instagram Follow — Verified')
    .setColor(0x57F287)
    .setDescription(`${message.author} verified as following all ${REQUIRED_SCREENSHOT_COUNT} required accounts.`)
    .addFields(
      { name: 'Required accounts', value: instagramUsernames.map(u => `@${u}`).join(', ') },
      { name: 'Matched usernames', value: result.matchedUsernames.filter(Boolean).join(', ') || '_not read_' },
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

  const usernames = store.settings.instagramUsernames;
  if (!usernames || usernames.length !== REQUIRED_SCREENSHOT_COUNT) {
    await message.reply({
      content: `⚠️ Screenshot verification isn't fully set up yet — an admin needs to run \`/set-instagram-accounts\` first (all ${REQUIRED_SCREENSHOT_COUNT} accounts).`,
    }).catch(() => {});
    return true;
  }

  if (!store.ssVerifications) store.ssVerifications = {};
  const existing = store.ssVerifications[message.author.id];
  if (existing) {
    await message.reply({
      content: `✅ You're already verified as following all ${REQUIRED_SCREENSHOT_COUNT} required accounts.`,
    }).catch(() => {});
    return true;
  }

  const images = imageAttachments(message);
  const accountList = usernames.map(u => `**@${u}**`).join(', ');

  if (images.length === 0) {
    await message.reply({
      content: `📸 Post **${REQUIRED_SCREENSHOT_COUNT} screenshots together in one message** — one for each of: ${accountList} — showing the "Following" button, to get verified.`,
    }).catch(() => {});
    return true;
  }

  if (images.length !== REQUIRED_SCREENSHOT_COUNT) {
    await message.reply({
      content: `❌ You attached ${images.length} screenshot(s), but exactly **${REQUIRED_SCREENSHOT_COUNT}** are required — one for each of: ${accountList}. Please resend all ${REQUIRED_SCREENSHOT_COUNT} together in a single message.`,
    }).catch(() => {});
    return true;
  }

  await message.channel.sendTyping().catch(() => {});

  const result = await analyzeFollowScreenshot(images.map(a => a.url), usernames);

  if (!result.ok) {
    await message.reply({
      content: '⚠️ Couldn\'t run verification right now (the AI checker is temporarily unavailable). Please try again shortly, or ping a staff member if this keeps happening.',
    }).catch(() => {});
    return true;
  }

  if (!result.verified) {
    const missingLine = result.missingAccounts.length
      ? `\nStill missing: ${result.missingAccounts.map(u => `**@${u}**`).join(', ')}`
      : '';
    const reasonLine = result.reason ? `\n> ${result.reason}` : '';
    await message.reply({
      content: `❌ Couldn't confirm you're following all ${REQUIRED_SCREENSHOT_COUNT} required accounts from those screenshots.${reasonLine}${missingLine}\nMake sure each screenshot clearly shows the matching profile page with the **Following** button, then resend all ${REQUIRED_SCREENSHOT_COUNT} together.`,
    }).catch(() => {});
    return true;
  }

  store.ssVerifications[message.author.id] = {
    verifiedAt: new Date().toISOString(),
    matchedUsernames: result.matchedUsernames,
    imageUrls: images.map(a => a.url),
  };
  saveGuildStore(message.guild.id, store);

  await giveSsVerifiedRole(message, store);

  await message.react('✅').catch(() => {});
  await message.reply({
    content: `✅ Verified! You're now confirmed as following all ${REQUIRED_SCREENSHOT_COUNT} required accounts.`,
  }).catch(() => {});

  const logChannel = await resolveLogChannel(message.guild, store, store.settings.ssVerifyLogChannelId);
  if (logChannel) {
    try {
      await logChannel.send({ embeds: [buildLogEmbed(message, usernames, result, images[0].url)] });
    } catch (err) {
      console.error('[ss-verify] Failed to post to log channel:', err);
    }
  }

  return true;
}

module.exports = { handleSsVerifyMessage, analyzeFollowScreenshot, REQUIRED_SCREENSHOT_COUNT };
