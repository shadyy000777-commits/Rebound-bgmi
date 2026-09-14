const { EmbedBuilder } = require('discord.js');
const { getGuildStore, saveGuildStore } = require('./storage');
const { resolveLogChannel } = require('./log-channel');

// Vision-capable model on Groq's free-tier API (same key/endpoint as
// ai-chat.js). Configurable via .env in case Groq renames/retires this
// model later — no code changes needed.
const VISION_MODEL = process.env.GROQ_VISION_MODEL || 'qwen/qwen3.6-27b';

// Only these get sent off to the vision model — random file attachments
// (zip, txt, etc.) in the channel are ignored rather than misread as "no
// screenshot" screenshots.
function imageAttachments(message) {
  return [...message.attachments.values()].filter(a =>
    (a.contentType || '').startsWith('image/')
  );
}

function buildPrompt(instagramUsername) {
  return `You are checking a screenshot submitted as proof of following the Instagram account "@${instagramUsername}".

Decide whether the image clearly shows the Instagram app open on that account's profile page with a "Following" state on the follow button (not a "Follow" button, which means they have NOT followed yet).

Be strict:
- The username visible in the screenshot must match "${instagramUsername}" (a leading "@" or minor case difference is fine).
- It must look like a genuine Instagram profile screenshot — not an unrelated image, a different app, or an obviously edited/mocked-up button.
- If the button says "Follow" (not "Following"/"Message"), or the username doesn't match, or you can't clearly tell, mark it NOT verified rather than guessing.

Respond with ONLY a JSON object, nothing else:
{"verified": true or false, "matched_username": "<username you read, or null>", "reason": "<one short sentence explaining the decision>"}`;
}

/**
 * Calls the Groq vision model with the submitted image(s). Returns
 * { ok: true, verified, matchedUsername, reason } on a successful call,
 * or { ok: false, error } if the model couldn't be reached/parsed — the
 * caller treats that as "couldn't verify right now", never as a pass.
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
    ...imageUrls.slice(0, 3).map(url => ({ type: 'image_url', image_url: { url } })),
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
        max_completion_tokens: 200,
        temperature: 0,
        response_format: { type: 'json_object' },
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

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (err) {
      console.error('[ss-verify] Failed to parse JSON from vision model:', text);
      return { ok: false, error: 'bad_json' };
    }

    return {
      ok: true,
      verified: parsed.verified === true,
      matchedUsername: typeof parsed.matched_username === 'string' ? parsed.matched_username : null,
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

function buildLogEmbed(message, instagramUsername, result, imageUrl) {
  return new EmbedBuilder()
    .setTitle('📸 Instagram Follow — Verified')
    .setColor(0x57F287)
    .setDescription(`${message.author} verified as following **@${instagramUsername}**.`)
    .addFields(
      { name: 'Matched username', value: result.matchedUsername || '_not read_', inline: true },
      { name: 'Model reasoning', value: result.reason || '_none given_' },
    )
    .setThumbnail(imageUrl)
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
      content: `📸 Post a screenshot showing you're **following @${instagramUsername}** on Instagram (their profile page, with the "Following" button visible) to get verified.`,
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
    const reasonLine = result.reason ? `\n> ${result.reason}` : '';
    await message.reply({
      content: `❌ Couldn't confirm you're following **@${instagramUsername}** from that screenshot.${reasonLine}\nMake sure the screenshot clearly shows their profile page with the **Following** button, then try again.`,
    }).catch(() => {});
    return true;
  }

  store.ssVerifications[message.author.id] = {
    verifiedAt: new Date().toISOString(),
    matchedUsername: result.matchedUsername,
    imageUrl: images[0].url,
  };
  saveGuildStore(message.guild.id, store);

  await giveSsVerifiedRole(message, store);

  await message.react('✅').catch(() => {});
  await message.reply({
    content: `✅ Verified! You're now confirmed as following **@${instagramUsername}**.`,
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

module.exports = { handleSsVerifyMessage, analyzeFollowScreenshot };
