const {
  EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle,
  ModalBuilder, TextInputBuilder, TextInputStyle, StringSelectMenuBuilder,
  ChannelSelectMenuBuilder, RoleSelectMenuBuilder, UserSelectMenuBuilder, ChannelType, AttachmentBuilder,
  MessageFlags, PermissionFlagsBits,
} = require('discord.js');
const ExcelJS = require('exceljs');
const { getGuildStore, saveGuildStore } = require('./storage');
const { buildGroupsEmbed, buildTournamentSlotListEmbed } = require('./embeds');
const {
  startPending: startRegPending, getPending: getRegPending,
  updatePending: updateRegPending, clearPending: clearRegPending,
} = require('./pending-tournament-registrations');

const RESTART_HINT = 'Click **Register Team** again to restart — no partial data is saved.';

const MAX_GROUPS = 60;
const DEFAULT_GROUP_CAPACITY = 20;
const MAX_GROUP_CAPACITY = 1000;
const MAX_TOTAL_SLOTS = 15000;
const MAX_ROUND = 10;
// Groups are keyed 1..MAX_GROUPS (plain numeric strings) rather than
// letters, so "Group 1", "Group 2"... display correctly everywhere they're
// already interpolated as `Group ${letter}` without needing a separate
// display-name lookup.
const GROUP_LETTERS = Array.from({ length: MAX_GROUPS }, (_, i) => String(i + 1));
const MAX_GUILD_ROLES = 250;
const MAX_GUILD_CHANNELS = 500;
const SAFETY_MARGIN = 5;

function isBanned(tournament, teamName) {
  return (tournament.bannedTeams || []).includes(teamName.toLowerCase());
}

function isDuplicateTeam(tournament, teamName) {
  return Object.values(tournament.groups)
    .some(g => g.teams.some(t => t.team.toLowerCase() === teamName.toLowerCase()));
}

// ---------------------------------------------------------------------------
// Main panel
// ---------------------------------------------------------------------------
function buildTournamentWizardPayload(store) {
  const tournament = store.tournament;

  const embed = new EmbedBuilder()
    .setTitle(tournament ? `🥇 Tournament Setup — ${tournament.name}` : '🥇 Tournament Setup')
    .setColor(tournament ? (tournament.open ? 0x57F287 : 0xED4245) : 0x5865F2);

  if (!tournament) {
    embed.setDescription('No tournament is set up yet. Click **Create Tournament** to get started — the rest of these buttons need one to exist first.');
  } else {
    const groupCount = Object.keys(tournament.groups).length;
    const teamCount = Object.values(tournament.groups).reduce((sum, g) => sum + g.teams.length, 0);
    const bannedCount = (tournament.bannedTeams || []).length;

    embed.addFields(
      { name: 'Status', value: tournament.open ? '🟢 Open' : '🔴 Closed', inline: true },
      { name: 'Groups', value: groupCount ? String(groupCount) : 'None yet', inline: true },
      { name: 'Teams Registered', value: String(teamCount), inline: true },
    );

    if (bannedCount) {
      embed.addFields({ name: '🔨 Banned Teams', value: String(bannedCount), inline: true });
    }
    if (tournament.slotManagerChannelId) {
      embed.addFields({ name: 'Slot-Manager Channel', value: `<#${tournament.slotManagerChannelId}>`, inline: true });
    }
  }

  const row1 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_create').setLabel('Create Tournament').setEmoji('➕').setStyle(ButtonStyle.Success),
    new ButtonBuilder()
      .setCustomId('tourney_wizard_toggle')
      .setLabel('Start/Pause Reg')
      .setEmoji(tournament && tournament.open ? '⏸️' : '▶️')
      .setStyle(tournament && tournament.open ? ButtonStyle.Danger : ButtonStyle.Success),
    new ButtonBuilder().setCustomId('tourney_wizard_manage_groups').setLabel('Manage Groups').setEmoji('🗂️').setStyle(ButtonStyle.Success),
  );
  const row2 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_edit_settings').setLabel('Edit Settings').setEmoji('🛠️').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_wizard_create_channels').setLabel('Create Channels').setEmoji('📺').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_wizard_ban_unban').setLabel('Ban/Unban').setEmoji('🔨').setStyle(ButtonStyle.Danger),
  );
  const row3 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_cancel_slots').setLabel('Cancel Slots').setEmoji('🗑️').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId('tourney_wizard_manual_add').setLabel('Manually Add Slot').setEmoji('📌').setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId('tourney_wizard_post_register_panel').setLabel('Post Register Panel').setEmoji('📮').setStyle(ButtonStyle.Success),
  );
  const row4 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_slot_manager_channel').setLabel('Slot-Manager channel').setEmoji('📡').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_wizard_excel_export').setLabel('MS Excel File').setEmoji('📊').setStyle(ButtonStyle.Primary),
  );
  const row5 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_delete').setLabel('Delete Tournament').setEmoji('🗑️').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId('tourney_wizard_help').setLabel('Help').setEmoji('❓').setStyle(ButtonStyle.Secondary),
  );

  return { embeds: [embed], components: [row1, row2, row3, row4, row5] };
}

// "Manage Groups" now goes straight to the group picker (see
// buildSlotListGroupSelectPayload, further down) — pick a group and its
// current slot list is shown. Nothing else to configure here manually
// since groups are auto-created from Total Slots / Teams-per-Group when
// Create Channels is hit, and each group's own channel carries its own
// admin panel (Publish Slot List / Punish Team / Result — see
// buildTournamentGroupAdminPanelPayload) for jobs scoped to that group.

// Sub-panel behind "Create Channels" — lets the admin choose between the
// quick default (Auto Channels) and picking a category + naming format
// themselves (Create Channel).
function buildCreateChannelsSubmenuPayload() {
  const embed = new EmbedBuilder()
    .setTitle('📺 Create Channels')
    .setColor(0x5865F2)
    .setDescription(
      '**Auto Channels** creates every group\'s own channel right away with the default naming (`group-<letter>`), under an auto-managed "🥇 Tournament Groups" category — one channel per group.\n\n' +
      '**Create Channel** creates exactly one category with one channel inside it (not per group) — you name both yourself.'
    );

  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_create_channels_manual').setLabel('Create Channel').setEmoji('➕').setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId('tourney_wizard_create_channels_auto').setLabel('Auto Channels').setEmoji('⚙️').setStyle(ButtonStyle.Success),
  );

  return { embeds: [embed], components: [row] };
}

// Posted automatically in a group's own channel the moment that channel is
// created (see ensureTournamentGroupChannelAndRole and the "Create
// Channels" bulk button) — the per-group counterpart to the top-level
// wizard panel, scoped to just this one group.
function buildTournamentGroupAdminPanelPayload(letter) {
  const embed = new EmbedBuilder()
    .setTitle(`🛠️ Group ${letter} — Admin Panel`)
    .setColor(0x5865F2)
    .setDescription('**Publish Slot List** posts this group\'s current teams here. **Punish Team** bans a registered team and strips their roles. **Result** marks this group\'s qualifiers and promotes them into a Round 2 group (its own role + channel), filling each Round 2 group before moving to the next.');

  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`tourney_wizard_group_publish:${letter}`).setLabel('Publish Slot List').setEmoji('📤').setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId(`tourney_wizard_group_punish:${letter}`).setLabel('Punish Team').setEmoji('🔨').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId(`tourney_wizard_group_result:${letter}`).setLabel('Result').setEmoji('🌟').setStyle(ButtonStyle.Primary),
  );

  return { embeds: [embed], components: [row] };
}

// Public panel — this is the one meant to live in a #register-style
// channel where players (not admins) click to sign their team up. It's
// just an embed + the same Register Team button the admin panel used to
// carry, but posted on its own so players never see admin controls.
function buildTournamentRegisterPanelPayload(tournament) {
  const teamCount = Object.values(tournament.groups).reduce((sum, g) => sum + g.teams.length, 0);
  const perGroupCapacity = tournament.teamsPerGroup || DEFAULT_GROUP_CAPACITY;
  // Effective max capacity for display — the smaller of the admin's overall
  // totalSlots cap (if set) and what MAX_GROUPS groups can actually hold.
  // Groups themselves are created lazily as they're needed, so this is a
  // ceiling, not a count of slots that already exist.
  const maxCapacity = Math.min(
    tournament.totalSlots || Infinity,
    MAX_GROUPS * perGroupCapacity,
  );

  const embed = new EmbedBuilder()
    .setTitle(`🥇 ${tournament.name} — Team Registration`)
    .setColor(tournament.open ? 0x57F287 : 0xED4245)
    .setDescription(
      tournament.open
        ? 'Click **Register Team** below, enter your team name, then mention your teammates. You\'ll be auto-assigned to whichever group still has room.'
        : '🔒 Registration is currently closed.'
    )
    .addFields(
      { name: 'Status', value: tournament.open ? '🟢 Open' : '🔴 Closed', inline: true },
      { name: 'Slots Filled', value: Number.isFinite(maxCapacity) ? `${teamCount}/${maxCapacity}` : String(teamCount), inline: true },
    );

  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_register_team').setLabel('Register Team').setEmoji('📝').setStyle(ButtonStyle.Success)
  );

  return { embeds: [embed], components: [row] };
}
// Quotient-style lettered settings screen — each field is edited by its own
// button (A-G) rather than one big modal, since Discord modals cap at 5
// text inputs and two of these fields (channel, role) need pickers anyway.
// Every pick saves immediately, so "Go Back" and "Save" both just return to
// the main panel — there's no separate unsaved draft to discard or commit.
function buildCreateSettingsPayload(tournament) {
  const embed = new EmbedBuilder()
    .setTitle('Enter details & Press Save')
    .setColor(0x5865F2)
    .addFields(
      { name: 'A. Registration Channel', value: tournament.registrationChannelId ? `<#${tournament.registrationChannelId}>` : 'Not-Set' },
      { name: 'B. Confirm Channel', value: tournament.confirmChannelId ? `<#${tournament.confirmChannelId}>` : 'Not-Set' },
      { name: 'C. Required Mentions', value: String(tournament.requiredMentions ?? 4) },
      { name: 'D. Teams per Group', value: tournament.teamsPerGroup ? String(tournament.teamsPerGroup) : 'Not-Set' },
      { name: 'E. Total Slots', value: tournament.totalSlots ? String(tournament.totalSlots) : 'Not-Set' },
      { name: 'F. Reactions', value: (tournament.reactions && tournament.reactions.length ? tournament.reactions : ['✅', '❌']).join(' , ') },
    );

  const row1 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_create_settings_a').setLabel('A').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_create_settings_b').setLabel('B').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_create_settings_c').setLabel('C').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_create_settings_d').setLabel('D').setStyle(ButtonStyle.Primary),
  );
  const row2 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_create_settings_e').setLabel('E').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_create_settings_f').setLabel('F').setStyle(ButtonStyle.Primary),
  );
  const row3 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_manage_rounds').setLabel('Manage Rounds').setEmoji('🏆').setStyle(ButtonStyle.Secondary),
  );
  const row4 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_create_settings_back').setLabel('Go Back').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId('tourney_create_settings_save').setLabel('Save').setStyle(ButtonStyle.Success),
  );

  return { embeds: [embed], components: [row1, row2, row3, row4] };
}


function buildHelpEmbed() {
  return new EmbedBuilder()
    .setTitle('❓ Tournament Panel Help')
    .setColor(0x5865F2)
    .setDescription([
      '**Start/Pause Reg** — open or close team registration',
      '**Manage Groups** — pick a group to view its current slot list',
      '**Create Channels** — groups auto-generate from Total Slots / Teams-per-Group if none exist yet, then pick **Auto Channels** for the default setup or **Create Channel** to choose the category and naming format yourself',
      '**Each group\'s own channel** — carries its own panel: Publish Slot List, Punish Team, and Result, scoped to just that group',
      '**Edit Settings** — rename the tournament',
      '**Create Channels** — auto-create a text channel per group',
      '**Ban/Unban** — block or unblock a team name from registering',
      '**Cancel Slots** — remove a registered team from its group',
      '**Manually Add Slot** — force-register a team into a specific group, bypassing auto-assign',
      '**Post Register Panel** — posts the public registration panel in this channel, for players to register themselves',
      '**Slot-Manager channel** — pick a channel where slot lists get published automatically, and where a self-service panel (Cancel My Slot / My Groups / Change Team Name) is posted for players',
      '**MS Excel File** — export the full slot list as a spreadsheet',
      '**Delete Tournament** — wipe everything and start over',
    ].join('\n'));
}

// ---------------------------------------------------------------------------
// Modals
// ---------------------------------------------------------------------------
function buildTournamentCreateModal() {
  return new ModalBuilder()
    .setCustomId('tourney_wizard_create_modal')
    .setTitle('Create Tournament')
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('name').setLabel('Tournament name').setStyle(TextInputStyle.Short)
          .setPlaceholder('e.g. BGMI Winter Championship').setRequired(true).setMaxLength(80)
      ),
    );
}

function buildAddGroupModal() {
  return new ModalBuilder()
    .setCustomId('tourney_wizard_group_modal')
    .setTitle('Add Group')
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('letter').setLabel(`Group number (1-${MAX_GROUPS})`).setStyle(TextInputStyle.Short)
          .setPlaceholder('1').setRequired(true).setMaxLength(2)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('capacity').setLabel('Team capacity').setStyle(TextInputStyle.Short)
          .setPlaceholder('e.g. 20').setRequired(true).setMaxLength(4)
      ),
    );
}

// Core group-creation math, shared by the "Create Channels" auto-setup path
// below. Fills GROUP_LETTERS sequentially with `perGroup`-capacity groups
// until `total` teams are covered. Returns null groups on validation error.
function computeAutoGroups(tournament, total, perGroup) {
  if (!Number.isInteger(total) || total < 1 || total > MAX_TOTAL_SLOTS) {
    return { error: `❌ Total teams must be a whole number between 1 and ${MAX_TOTAL_SLOTS}.` };
  }
  if (!Number.isInteger(perGroup) || perGroup < 1 || perGroup > MAX_GROUP_CAPACITY) {
    return { error: `❌ Teams per group must be a whole number between 1 and ${MAX_GROUP_CAPACITY}.` };
  }

  const groupsNeeded = Math.ceil(total / perGroup);
  const freeLetters = GROUP_LETTERS.filter(l => !tournament.groups[l]);

  if (groupsNeeded > freeLetters.length) {
    return {
      error: `❌ That needs **${groupsNeeded}** new group(s), but only **${freeLetters.length}** letter slot(s) are free (max ${GROUP_LETTERS.length} groups total). Raise "teams per group" or delete an unused group first.`,
    };
  }

  const createdLetters = [];
  let remaining = total;
  for (let i = 0; i < groupsNeeded; i++) {
    const letter = freeLetters[i];
    const capacity = Math.min(perGroup, remaining);
    tournament.groups[letter] = { capacity, teams: [] };
    remaining -= capacity;
    createdLetters.push(letter);
  }

  return { createdLetters, groupsNeeded };
}

// Shared by both "Create Channels" paths (auto and manual) — makes sure
// groups exist before any channel gets created, auto-generating them from
// Total Slots / Teams-per-Group. Safe to call even if some groups already
// exist (e.g. a team registered before Create Channels was ever pressed,
// auto-creating just one group) — computeAutoGroups only fills in letters
// that are still free, so existing groups are left untouched. Returns an
// error string to show the admin, or null once groups are ready.
function ensureGroupsExist(interaction, store) {
  const tournament = store.tournament;
  if (!tournament.totalSlots) {
    if (Object.keys(tournament.groups).length) return null;
    return '❌ Set **Total Slots** first — Edit Settings → Total Slots (and Teams-per-Group, if you want something other than the default).';
  }
  const perGroup = tournament.teamsPerGroup || DEFAULT_GROUP_CAPACITY;
  const groupsAlreadyCovered = Object.keys(tournament.groups).length;
  const totalGroupsWanted = Math.ceil(tournament.totalSlots / perGroup);
  const stillNeeded = totalGroupsWanted - groupsAlreadyCovered;
  if (stillNeeded <= 0) return null;

  const result = computeAutoGroups(tournament, stillNeeded * perGroup, perGroup);
  if (result.error) return result.error;
  saveGuildStore(interaction.guildId, store);
  return null;
}

// Looks up (or auto-creates, once per group) the private role scoped to
// one tournament group — the mechanism that keeps a group's channel
// visible only to teams actually registered into that group, never to the
// whole tournament. Shared by createGroupChannels (bulk "Auto Channels")
// and ensureTournamentGroupChannelAndRole (lazy, on registration) so both
// paths land on the exact same role for a given group instead of drifting.
// Never throws — a role-cap or permissions hiccup just logs and returns
// null, so it never blocks registration or channel creation outright.
async function ensureGroupRole(interaction, store, group, roleName, reason) {
  let role = group.roleId ? interaction.guild.roles.cache.get(group.roleId) : null;
  if (role) return role;

  const botMember = interaction.guild.members.me;
  if (!botMember.permissions.has('ManageRoles')) {
    console.error(`[tournament-group-role] Bot is missing the "Manage Roles" permission in guild ${interaction.guildId}.`);
    return null;
  }
  if (interaction.guild.roles.cache.size >= MAX_GUILD_ROLES - SAFETY_MARGIN) {
    console.error(`[tournament-group-role] Guild ${interaction.guildId} is at/near Discord's ${MAX_GUILD_ROLES}-role cap — skipping auto-create for "${roleName}".`);
    return null;
  }
  try {
    role = await interaction.guild.roles.create({ name: roleName, mentionable: false, reason });
    group.roleId = role.id;
    saveGuildStore(interaction.guildId, store);
    return role;
  } catch (err) {
    console.error(`[tournament-group-role] Failed to auto-create role "${roleName}" in guild ${interaction.guildId}: ${err.code ?? ''} ${err.message}`);
    return null;
  }
}

// Shared by both "Create Channels" paths — actually creates the Discord
// channel for every group that doesn't have one yet, under `parentId` (or
// no category) using `nameFormat` (a "{letter}" token gets swapped for the
// group letter; if the format doesn't include one, "-{letter}" is appended
// so names stay unique across groups). Every channel comes out private —
// only that specific group's own role can see it (auto-created here if it
// doesn't have one yet), so a team registered into Group 1 can never see
// Group 2's channel, and it can't spin up threads — same lockdown
// regardless of which path made it.
async function createGroupChannels(interaction, store, { nameFormat, parentId }) {
  const tournament = store.tournament;
  const letters = Object.keys(tournament.groups);
  const hasToken = /\{letter\}/i.test(nameFormat);

  const created = [];
  for (const letter of letters) {
    const group = tournament.groups[letter];
    if (group.channelId && interaction.guild.channels.cache.has(group.channelId)) continue;

    const role = await ensureGroupRole(interaction, store, group, `Tournament Group ${letter}`, `Auto-created for Group ${letter} tournament registration`);
    if (!role) continue; // reason already logged inside ensureGroupRole

    const overwrites = [
      { id: interaction.guild.roles.everyone.id, deny: [PermissionFlagsBits.ViewChannel] },
      {
        id: role.id,
        allow: [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages, PermissionFlagsBits.ReadMessageHistory],
        deny: [PermissionFlagsBits.CreatePublicThreads, PermissionFlagsBits.CreatePrivateThreads],
      },
    ];

    const name = hasToken ? nameFormat.replace(/\{letter\}/gi, letter) : `${nameFormat}-${letter}`;
    try {
      const channel = await interaction.guild.channels.create({
        name: name.toLowerCase(),
        type: ChannelType.GuildText,
        parent: parentId || undefined,
        permissionOverwrites: overwrites,
        reason: `Tournament group channel created by ${interaction.user.tag}`,
      });
      group.channelId = channel.id;
      created.push(`<#${channel.id}>`);
      await channel.send(buildTournamentGroupAdminPanelPayload(letter)).catch(() => {});
    } catch (err) {
      console.error(`[tournament] Failed to create channel for group ${letter}:`, err.message);
    }
  }
  saveGuildStore(interaction.guildId, store);
  return created;
}

// In-progress "Create Channel" (manual) data — Channel Format and Category
// Name are set one at a time via separate modals, so this bridges them
// until both are set and "Create Channels" is pressed. In-memory only: if
// the bot restarts mid-flow, the admin just presses the button again.
const pendingManualChannelCreation = new Map(); // key: `${guildId}:${userId}` -> { data, timer }
const MANUAL_CHANNEL_CREATION_TTL_MS = 15 * 60 * 1000;

function manualChannelCreationKey(guildId, userId) {
  return `${guildId}:${userId}`;
}

function setPendingChannelCreation(guildId, userId, fields) {
  const key = manualChannelCreationKey(guildId, userId);
  const existing = pendingManualChannelCreation.get(key);
  if (existing && existing.timer) clearTimeout(existing.timer);
  const data = { ...(existing ? existing.data : {}), ...fields };
  const timer = setTimeout(() => pendingManualChannelCreation.delete(key), MANUAL_CHANNEL_CREATION_TTL_MS);
  pendingManualChannelCreation.set(key, { data, timer });
  return data;
}

function getPendingChannelCreation(guildId, userId) {
  const entry = pendingManualChannelCreation.get(manualChannelCreationKey(guildId, userId));
  return entry ? entry.data : null;
}

function clearPendingChannelCreation(guildId, userId) {
  const key = manualChannelCreationKey(guildId, userId);
  const entry = pendingManualChannelCreation.get(key);
  if (entry && entry.timer) clearTimeout(entry.timer);
  pendingManualChannelCreation.delete(key);
}

// Panel behind "Create Channel" (manual) — set Channel Format and Category
// Name (each via its own modal), then Create Channels makes a brand-new
// category with that name and every group's channel inside it.
function buildManualChannelCreationPayload(data) {
  const ready = Boolean(data.channelFormat && data.categoryName);
  const embed = new EmbedBuilder()
    .setTitle('📺 Tournament Channel Creation')
    .setColor(ready ? 0x57F287 : 0x5865F2)
    .setDescription('Creates exactly one category with one channel inside it — not per group.')
    .addFields(
      { name: 'Channel Name', value: data.channelFormat ? `\`${data.channelFormat}\`` : '`Not Set`' },
      { name: 'Category Name', value: data.categoryName ? `\`${data.categoryName}\`` : '`Not Set`' },
      { name: 'Status', value: ready ? '✅ Ready — hit Create Channel.' : '🔒 Set both the channel name and category name to continue' },
    );

  const row1 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_manual_channels_set_format').setLabel('Channel Name').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_wizard_manual_channels_set_category').setLabel('Category Name').setStyle(ButtonStyle.Primary),
  );
  const row2 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_manual_channels_create').setLabel('Create Channel').setEmoji('✅').setStyle(ButtonStyle.Success).setDisabled(!ready),
    new ButtonBuilder().setCustomId('tourney_wizard_manual_channels_cancel').setLabel('Cancel').setEmoji('🚫').setStyle(ButtonStyle.Danger),
  );

  return { embeds: [embed], components: [row1, row2] };
}

function buildAutoGroupsModal(tournament) {
  const totalInput = new TextInputBuilder().setCustomId('total').setLabel('Total teams expected').setStyle(TextInputStyle.Short)
    .setPlaceholder('e.g. 720').setRequired(true).setMaxLength(5);
  if (tournament && tournament.totalSlots) totalInput.setValue(String(tournament.totalSlots));

  const perGroupInput = new TextInputBuilder().setCustomId('per_group').setLabel('Teams per group').setStyle(TextInputStyle.Short)
    .setPlaceholder('e.g. 20').setRequired(true).setMaxLength(4);
  if (tournament && tournament.teamsPerGroup) perGroupInput.setValue(String(tournament.teamsPerGroup));

  return new ModalBuilder()
    .setCustomId('tourney_wizard_auto_groups_modal')
    .setTitle('Auto-Create Groups')
    .addComponents(
      new ActionRowBuilder().addComponents(totalInput),
      new ActionRowBuilder().addComponents(perGroupInput),
    );
}

// Modal only collects the team name — Discord modals can't hold a user-select
// component, so the 4 player mentions are picked in a follow-up step (see
// buildMentionPlayersRow) after this is submitted.
function buildRegisterTeamModal() {
  return new ModalBuilder()
    .setCustomId('tourney_wizard_register_modal')
    .setTitle('Register Team')
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('team').setLabel('Team name').setStyle(TextInputStyle.Short)
          .setRequired(true).setMaxLength(80)
      ),
    );
}

function buildMentionPlayersRow(count = 4) {
  const menu = new UserSelectMenuBuilder()
    .setCustomId('tourney_reg_select_players')
    .setPlaceholder(`Mention the ${count} player${count === 1 ? '' : 's'} on your team`)
    .setMinValues(count)
    .setMaxValues(count);
  return new ActionRowBuilder().addComponents(menu);
}

function tourneyConfirmRow() {
  return new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_reg_confirm').setLabel('Confirm Registration').setEmoji('✅').setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId('tourney_wizard_reg_cancel').setLabel('Cancel').setEmoji('✖️').setStyle(ButtonStyle.Danger)
  );
}

// Whether any of the newly-picked player IDs is already locked into another
// team's roster in this tournament — a player can only be mentioned on one
// team at a time. Returns the conflicting user ID and their team, or null.
function findTournamentLineupConflict(tournament, selectedIds) {
  for (const group of Object.values(tournament.groups)) {
    for (const t of group.teams) {
      for (const id of t.playerIds || []) {
        if (selectedIds.includes(id)) {
          return { conflictId: id, team: t.team };
        }
      }
    }
  }
  return null;
}

function buildTeamRegPreviewEmbed(data) {
  const lineup = (data.selectedPlayerIds || []).map(id => `<@${id}>`).join(' ') || '_none selected_';
  return new EmbedBuilder()
    .setTitle('📝 Review Your Registration')
    .setColor(0xFEE75C)
    .setDescription(`**Team Name** — ${data.team}\n**Players** — ${lineup}`)
    .setFooter({ text: 'Double-check everything, then tap Confirm to lock in your slot.' });
}

function buildEditSettingsModal(tournament) {
  const nameInput = new TextInputBuilder().setCustomId('name').setLabel('Tournament name').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(80);
  if (tournament && tournament.name) nameInput.setValue(tournament.name);

  return new ModalBuilder()
    .setCustomId('tourney_wizard_edit_modal')
    .setTitle('Edit Settings')
    .addComponents(new ActionRowBuilder().addComponents(nameInput));
}

function buildBanUnbanModal() {
  return new ModalBuilder()
    .setCustomId('tourney_wizard_ban_modal')
    .setTitle('Ban / Unban Team')
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('team').setLabel('Team name').setStyle(TextInputStyle.Short)
          .setPlaceholder('Exact team name — running this again unbans it').setRequired(true).setMaxLength(80)
      ),
    );
}

function buildManualAddSlotModal() {
  return new ModalBuilder()
    .setCustomId('tourney_wizard_manual_add_modal')
    .setTitle('Manually Add Slot')
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('letter').setLabel('Group number').setStyle(TextInputStyle.Short)
          .setPlaceholder('1').setRequired(true).setMaxLength(2)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('team').setLabel('Team name').setStyle(TextInputStyle.Short)
          .setRequired(true).setMaxLength(80)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('player1').setLabel('Player 1 IGN').setStyle(TextInputStyle.Short)
          .setRequired(true).setMaxLength(40)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('player2').setLabel('Player 2 IGN').setStyle(TextInputStyle.Short)
          .setRequired(false).setMaxLength(40)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('players34').setLabel('Player 3 & 4 IGN (comma separated)').setStyle(TextInputStyle.Short)
          .setRequired(false).setMaxLength(80)
      ),
    );
}

function buildRequiredMentionsModal(tournament) {
  const input = new TextInputBuilder().setCustomId('value').setLabel('Required Mentions (1-4)').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(1).setPlaceholder('4');
  if (tournament.requiredMentions) input.setValue(String(tournament.requiredMentions));
  return new ModalBuilder().setCustomId('tourney_create_settings_d_modal').setTitle('Required Mentions')
    .addComponents(new ActionRowBuilder().addComponents(input));
}

function buildTeamsPerGroupModal(tournament) {
  const input = new TextInputBuilder().setCustomId('value').setLabel('Teams per Group').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(4).setPlaceholder('e.g. 20');
  input.setValue(String(tournament.teamsPerGroup || DEFAULT_GROUP_CAPACITY));
  return new ModalBuilder().setCustomId('tourney_create_settings_e_modal').setTitle('Teams per Group')
    .addComponents(new ActionRowBuilder().addComponents(input));
}

function buildTotalSlotsModal(tournament) {
  const input = new TextInputBuilder().setCustomId('value').setLabel('Total Slots').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(5).setPlaceholder('e.g. 15000');
  if (tournament.totalSlots) input.setValue(String(tournament.totalSlots));
  return new ModalBuilder().setCustomId('tourney_create_settings_f_modal').setTitle('Total Slots')
    .addComponents(new ActionRowBuilder().addComponents(input));
}

function buildReactionsModal(tournament) {
  const reactions = tournament.reactions && tournament.reactions.length ? tournament.reactions : ['✅', '❌'];
  const accept = new TextInputBuilder().setCustomId('accept').setLabel('Accept reaction (emoji)').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(32).setValue(reactions[0] || '✅');
  const deny = new TextInputBuilder().setCustomId('deny').setLabel('Deny reaction (emoji)').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(32).setValue(reactions[1] || '❌');
  return new ModalBuilder().setCustomId('tourney_create_settings_g_modal').setTitle('Reactions')
    .addComponents(
      new ActionRowBuilder().addComponents(accept),
      new ActionRowBuilder().addComponents(deny),
    );
}




function hasManageGuild(interaction) {
  return interaction.member.permissions.has(PermissionFlagsBits.ManageGuild);
}

// ---------------------------------------------------------------------------
// Button handler
// ---------------------------------------------------------------------------
async function handleTournamentWizardButton(interaction) {
  const store = getGuildStore(interaction.guildId);
  const id = interaction.customId;

  // Public: any player can register a team, no Manage Server needed.
  if (id === 'tourney_wizard_register_team') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    if (!store.tournament.open) {
      return interaction.reply({ content: '❌ Registration is currently closed.', flags: MessageFlags.Ephemeral });
    }
    return interaction.showModal(buildRegisterTeamModal());
  }

  // Public: continuing the register flow, also no Manage Server needed.
  if (id === 'tourney_wizard_reg_confirm') {
    return handleTourneyRegConfirm(interaction);
  }
  if (id === 'tourney_wizard_reg_cancel') {
    return handleTourneyRegCancel(interaction);
  }

  // Public: self-service slot management, posted in the Slot-Manager
  // channel — any registered player can cancel their own slot, check
  // which group they're in, or rename their own team. No Manage Server
  // permission needed for any of these.
  if (id === 'tourney_wizard_selfservice_cancel') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const entry = findUserTournamentEntry(store.tournament, interaction.user.id);
    if (!entry) {
      return interaction.reply({ content: "❌ You're not registered for this tournament.", flags: MessageFlags.Ephemeral });
    }
    const row = new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId('tourney_wizard_selfservice_cancel_confirm').setLabel('Yes, Cancel My Slot').setStyle(ButtonStyle.Danger),
      new ButtonBuilder().setCustomId('tourney_wizard_selfservice_cancel_abort').setLabel('Never Mind').setStyle(ButtonStyle.Secondary),
    );
    return interaction.reply({
      content: `⚠️ Cancel **${entry.team.team}**'s registration in Group **${entry.letter}**? This removes your tournament role(s) and frees the slot for someone else. This can't be undone.`,
      components: [row],
      flags: MessageFlags.Ephemeral,
    });
  }

  if (id === 'tourney_wizard_selfservice_cancel_abort') {
    return interaction.update({ content: "✅ No changes made — you're still registered.", components: [] });
  }

  if (id === 'tourney_wizard_selfservice_cancel_confirm') {
    if (!store.tournament) {
      return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
    }
    const tournament = store.tournament;
    const entry = findUserTournamentEntry(tournament, interaction.user.id);
    if (!entry) {
      return interaction.update({ content: "❌ You're not registered for this tournament.", components: [] });
    }

    await interaction.deferUpdate();

    const { letter, group, team } = entry;
    await removeTeamFromRound2(interaction, store, team);
    const groupRoleId = group.roleId;
    const targets = new Set([team.ownerId, ...(team.playerIds || [])].filter(Boolean));
    for (const userId of targets) {
      const member = interaction.guild.members.cache.get(userId)
        ?? await interaction.guild.members.fetch(userId).catch(() => null);
      if (!member) continue;
      if (groupRoleId) await member.roles.remove(groupRoleId).catch(() => {});
    }

    group.teams = group.teams.filter(t => t !== team);
    tournament.qualified = tournament.qualified.filter(name => name !== team.team);
    saveGuildStore(interaction.guildId, store);

    return interaction.editReply({ content: `🗑️ **${team.team}** has been removed from Group ${letter} — your tournament roles have been cleared.`, components: [] });
  }

  if (id === 'tourney_wizard_selfservice_my_groups') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const entry = findUserTournamentEntry(store.tournament, interaction.user.id);
    if (!entry) {
      return interaction.reply({ content: "❌ You're not registered for this tournament.", flags: MessageFlags.Ephemeral });
    }
    const { letter, team, idx } = entry;
    const lineup = (team.playerIds && team.playerIds.length)
      ? team.playerIds.map(id => `<@${id}>`).join(' ')
      : (team.players || []).join(' ');
    return interaction.reply({
      content: `📋 **${team.team}** is in **Group ${letter}**, Slot **${idx + 1}**.\n👥 ${lineup}`,
      flags: MessageFlags.Ephemeral,
    });
  }

  if (id === 'tourney_wizard_selfservice_change_name') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const entry = findUserTournamentEntry(store.tournament, interaction.user.id);
    if (!entry) {
      return interaction.reply({ content: "❌ You're not registered for this tournament.", flags: MessageFlags.Ephemeral });
    }
    const modal = new ModalBuilder().setCustomId('tourney_selfservice_change_name_modal').setTitle('Change Team Name');
    const input = new TextInputBuilder()
      .setCustomId('team').setLabel('New team name').setStyle(TextInputStyle.Short)
      .setValue(entry.team.team).setMaxLength(100).setRequired(true);
    modal.addComponents(new ActionRowBuilder().addComponents(input));
    return interaction.showModal(modal);
  }

  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_create') {
    if (store.tournament) {
      return interaction.reply({ content: '❌ A tournament already exists. Delete it first to create a new one.', flags: MessageFlags.Ephemeral });
    }
    return interaction.showModal(buildTournamentCreateModal());
  }

  if (id === 'tourney_create_settings_a') {
    if (!store.tournament) return interaction.update({ content: '❌ No tournament exists yet.', embeds: [], components: [] });
    const select = new ChannelSelectMenuBuilder().setCustomId('tourney_create_regchannel_select')
      .setPlaceholder('Choose the registration channel').addChannelTypes(ChannelType.GuildText);
    return interaction.update({ content: 'A. Pick the registration channel:', embeds: [], components: [new ActionRowBuilder().addComponents(select)] });
  }

  if (id === 'tourney_create_settings_b') {
    if (!store.tournament) return interaction.update({ content: '❌ No tournament exists yet.', embeds: [], components: [] });
    const select = new ChannelSelectMenuBuilder().setCustomId('tourney_create_confirmchannel_select')
      .setPlaceholder('Choose the confirm channel').addChannelTypes(ChannelType.GuildText);
    return interaction.update({ content: 'B. Pick the confirm channel:', embeds: [], components: [new ActionRowBuilder().addComponents(select)] });
  }

  if (id === 'tourney_create_settings_c') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildRequiredMentionsModal(store.tournament));
  }

  if (id === 'tourney_create_settings_d') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildTeamsPerGroupModal(store.tournament));
  }

  if (id === 'tourney_create_settings_e') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildTotalSlotsModal(store.tournament));
  }

  if (id === 'tourney_create_settings_f') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildReactionsModal(store.tournament));
  }

  if (id === 'tourney_create_settings_back' || id === 'tourney_create_settings_save') {
    // Everything on this screen saves the moment it's picked, so both
    // buttons do the same thing: drop back to the main panel.
    return interaction.update(buildTournamentWizardPayload(store));
  }

  if (id === 'tourney_wizard_manage_groups') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const payload = buildSlotListGroupSelectPayload(store.tournament);
    if (payload.error) {
      return interaction.reply({ content: payload.error, flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ ...payload, flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_edit_settings') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    return interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
  }

  if (id === 'tourney_wizard_ban_unban') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    return interaction.showModal(buildBanUnbanModal());
  }

  if (id === 'tourney_wizard_post_register_panel') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const select = new ChannelSelectMenuBuilder()
      .setCustomId('tourney_register_panel_channel_select')
      .setPlaceholder('Choose a channel to post the registration panel in')
      .addChannelTypes(ChannelType.GuildText);
    return interaction.reply({
      content: '📮 Pick a channel — players will register from there.',
      components: [new ActionRowBuilder().addComponents(select)],
      flags: MessageFlags.Ephemeral,
    });
  }

  if (id === 'tourney_wizard_manual_add') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    if (!Object.keys(store.tournament.groups).length) {
      return interaction.reply({ content: '❌ Add a group first.', flags: MessageFlags.Ephemeral });
    }
    return interaction.showModal(buildManualAddSlotModal());
  }

  if (id === 'tourney_wizard_cancel_slots') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const payload = buildCancelGroupSelectPayload(store.tournament);
    if (payload.error) {
      return interaction.reply({ content: payload.error, flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ ...payload, flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_slot_manager_channel') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const select = new ChannelSelectMenuBuilder()
      .setCustomId('tourney_slotmanager_channel_select')
      .setPlaceholder('Choose the slot-manager channel')
      .addChannelTypes(ChannelType.GuildText);
    return interaction.reply({
      content: '📡 Pick a channel — published slot lists will be posted there.',
      components: [new ActionRowBuilder().addComponents(select)],
      flags: MessageFlags.Ephemeral,
    });
  }

  if (id === 'tourney_wizard_create_channels') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ ...buildCreateChannelsSubmenuPayload(), flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_create_channels_auto') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const groupsErr = ensureGroupsExist(interaction, store);
    if (groupsErr) {
      return interaction.reply({ content: groupsErr, flags: MessageFlags.Ephemeral });
    }
    const me = interaction.guild.members.me;
    if (!me.permissions.has(PermissionFlagsBits.ManageChannels)) {
      return interaction.reply({ content: '❌ I need the **Manage Channels** permission to do that.', flags: MessageFlags.Ephemeral });
    }
    if (!me.permissions.has(PermissionFlagsBits.ManageRoles)) {
      return interaction.reply({ content: '❌ I need the **Manage Roles** permission to do that — each group gets its own private role.', flags: MessageFlags.Ephemeral });
    }

    await interaction.deferReply({ flags: MessageFlags.Ephemeral });

    // Default path shares the same "🥇 Tournament Groups" category Auto
    // Setup uses, so everything auto-created lands in one place.
    let category = store.settings && store.settings.tournamentGroupChannelsCategoryId
      ? interaction.guild.channels.cache.get(store.settings.tournamentGroupChannelsCategoryId)
      : null;
    if (!category) {
      try {
        category = await interaction.guild.channels.create({
          name: '🥇 Tournament Groups',
          type: ChannelType.GuildCategory,
          reason: 'Auto-created to hold per-group tournament channels',
        });
        if (!store.settings) store.settings = {};
        store.settings.tournamentGroupChannelsCategoryId = category.id;
        saveGuildStore(interaction.guildId, store);
      } catch (err) {
        console.error('[tournament] Failed to auto-create group category:', err.message);
      }
    }

    const created = await createGroupChannels(interaction, store, {
      nameFormat: 'group-{letter}',
      parentId: category ? category.id : null,
    });
    return interaction.editReply({
      content: created.length
        ? `✅ Created: ${created.join(', ')}`
        : 'ℹ️ Every group already has a channel (or channel creation failed — check my permissions).',
    });
  }

  if (id === 'tourney_wizard_create_channels_manual') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const me = interaction.guild.members.me;
    if (!me.permissions.has(PermissionFlagsBits.ManageChannels)) {
      return interaction.reply({ content: '❌ I need the **Manage Channels** permission to do that.', flags: MessageFlags.Ephemeral });
    }

    const data = setPendingChannelCreation(interaction.guildId, interaction.user.id, {});
    return interaction.reply({ ...buildManualChannelCreationPayload(data), flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_manual_channels_set_format') {
    const data = getPendingChannelCreation(interaction.guildId, interaction.user.id) || {};
    const modal = new ModalBuilder().setCustomId('tourney_manual_channels_format_modal').setTitle('Channel Format');
    const input = new TextInputBuilder()
      .setCustomId('value')
      .setLabel('Name for the channel')
      .setStyle(TextInputStyle.Short)
      .setPlaceholder('tournament-lobby')
      .setValue(data.channelFormat || '')
      .setRequired(true)
      .setMaxLength(80);
    modal.addComponents(new ActionRowBuilder().addComponents(input));
    return interaction.showModal(modal);
  }

  if (id === 'tourney_wizard_manual_channels_set_category') {
    const data = getPendingChannelCreation(interaction.guildId, interaction.user.id) || {};
    const modal = new ModalBuilder().setCustomId('tourney_manual_channels_categoryname_modal').setTitle('Category Name');
    const input = new TextInputBuilder()
      .setCustomId('value')
      .setLabel('Name for the new category')
      .setStyle(TextInputStyle.Short)
      .setPlaceholder('🥇 Tournament Groups')
      .setValue(data.categoryName || '')
      .setRequired(true)
      .setMaxLength(100);
    modal.addComponents(new ActionRowBuilder().addComponents(input));
    return interaction.showModal(modal);
  }

  if (id === 'tourney_wizard_manual_channels_cancel') {
    clearPendingChannelCreation(interaction.guildId, interaction.user.id);
    return interaction.update({ content: '❌ Cancelled.', embeds: [], components: [] });
  }

  if (id === 'tourney_wizard_manual_channels_create') {
    const data = getPendingChannelCreation(interaction.guildId, interaction.user.id);
    if (!data || !data.channelFormat || !data.categoryName) {
      return interaction.reply({ content: '❌ Set both Channel Format and Category Name first.', flags: MessageFlags.Ephemeral });
    }
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const me = interaction.guild.members.me;
    if (!me.permissions.has(PermissionFlagsBits.ManageChannels)) {
      return interaction.reply({ content: '❌ I need the **Manage Channels** permission to do that.', flags: MessageFlags.Ephemeral });
    }

    await interaction.deferUpdate();

    // Exactly one category, one channel — not per group, and not tied to
    // any tournament role, so it's created with default (visible-to-
    // everyone) permissions rather than gated behind a role that no
    // longer exists.
    let category;
    try {
      category = await interaction.guild.channels.create({
        name: data.categoryName,
        type: ChannelType.GuildCategory,
        reason: `Tournament category created by ${interaction.user.tag}`,
      });
    } catch (err) {
      console.error('[tournament] Failed to create category:', err.message);
      return interaction.editReply({ content: '❌ Failed to create the category — check my **Manage Channels** permission.', embeds: [], components: [] });
    }

    let channel;
    try {
      channel = await interaction.guild.channels.create({
        name: data.channelFormat.replace(/\{letter\}/gi, '').trim() || 'tournament',
        type: ChannelType.GuildText,
        parent: category.id,
        reason: `Tournament channel created by ${interaction.user.tag}`,
      });
    } catch (err) {
      console.error('[tournament] Failed to create channel:', err.message);
      return interaction.editReply({ content: `❌ Category **${category.name}** was created, but the channel failed — check my **Manage Channels** permission.`, embeds: [], components: [] });
    }

    clearPendingChannelCreation(interaction.guildId, interaction.user.id);
    return interaction.editReply({
      content: `✅ Created category **${category.name}** with <#${channel.id}>.`,
      embeds: [],
      components: [],
    });
  }

  if (id === 'tourney_wizard_excel_export') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    await interaction.deferReply({ flags: MessageFlags.Ephemeral });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Slot List');
    sheet.columns = [
      { header: 'Group', key: 'group', width: 10 },
      { header: 'Slot', key: 'slot', width: 8 },
      { header: 'Team', key: 'team', width: 30 },
      { header: 'Players', key: 'players', width: 50 },
    ];
    for (const [letter, group] of Object.entries(store.tournament.groups).sort(([a], [b]) => a.localeCompare(b))) {
      group.teams.forEach((t, idx) => {
        // Newer registrations store real Discord IDs in playerIds (from the
        // mention-based flow) — resolve those to readable tags for the
        // spreadsheet instead of dumping raw <@id> mention text. Older
        // registrations (typed IGNs, no playerIds) fall back to t.players as-is.
        const playerLabels = (t.playerIds && t.playerIds.length)
          ? t.playerIds.map(id => interaction.guild.members.cache.get(id)?.user.tag ?? `<@${id}>`)
          : (t.players || []);
        sheet.addRow({ group: letter, slot: idx + 1, team: t.team, players: playerLabels.join(', ') });
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `tournament-${(store.tournament.name || 'export').replace(/[^a-z0-9]+/gi, '-').toLowerCase()}.xlsx`;
    const attachment = new AttachmentBuilder(Buffer.from(buffer), { name: filename });
    return interaction.editReply({ content: '📊 Full slot list export:', files: [attachment] });
  }

  if (id === 'tourney_wizard_help') {
    return interaction.reply({ embeds: [buildHelpEmbed()], flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_toggle') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    store.tournament.open = !store.tournament.open;
    saveGuildStore(interaction.guildId, store);
    const payload = buildTournamentWizardPayload(store);
    await interaction.update(payload);
    await interaction.channel.send(
      store.tournament.open
        ? `✅ Registration opened for **${store.tournament.name}**. Teams can now register.`
        : `🔒 Registration for **${store.tournament.name}** is now closed.`
    ).catch(() => {});
    return;
  }

  // View Groups / Slot List / Qualify used to live on the top-level panel —
  // removed in favor of each group's own channel panel below, which covers
  // publish/punish/result scoped to that specific group.

  if (id.startsWith('tourney_wizard_group_publish:')) {
    const [, letter] = id.split(':');
    return handleTourneyGroupPublish(interaction, store, letter);
  }

  if (id.startsWith('tourney_wizard_group_punish:')) {
    const [, letter] = id.split(':');
    if (!store.tournament || !store.tournament.groups[letter]) {
      return interaction.reply({ content: '❌ That group no longer exists.', flags: MessageFlags.Ephemeral });
    }
    const payload = buildTournamentPunishSelectPayload(store.tournament, letter);
    if (payload.error) {
      return interaction.reply({ content: payload.error, flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ ...payload, flags: MessageFlags.Ephemeral });
  }

  if (id.startsWith('tourney_wizard_group_result:')) {
    const [, letter] = id.split(':');
    if (!store.tournament || !store.tournament.groups[letter]) {
      return interaction.reply({ content: '❌ That group no longer exists.', flags: MessageFlags.Ephemeral });
    }
    const payload = buildQualifySelectPayload(store.tournament, letter);
    if (payload.error) {
      return interaction.reply({ content: payload.error, flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ ...payload, flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_delete') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const row = new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId('tourney_wizard_delete_confirm').setLabel('Yes, delete it').setStyle(ButtonStyle.Danger),
      new ButtonBuilder().setCustomId('tourney_wizard_delete_cancel').setLabel('Cancel').setStyle(ButtonStyle.Secondary),
    );
    return interaction.update({
      content: `⚠️ Delete **${store.tournament.name}** and all groups/registrations? This can't be undone.`,
      embeds: [],
      components: [row],
    });
  }

  if (id === 'tourney_wizard_delete_confirm') {
    const groups = store.tournament ? Object.values(store.tournament.groups) : [];
    let cleanupFailures = 0;

    for (const group of groups) {
      if (group.channelId) {
        const channel = interaction.guild.channels.cache.get(group.channelId);
        if (channel) {
          await channel.delete('Tournament deleted').catch(() => { cleanupFailures++; });
        }
      }
      if (group.roleId) {
        const role = interaction.guild.roles.cache.get(group.roleId);
        if (role) {
          await role.delete('Tournament deleted').catch(() => { cleanupFailures++; });
        }
      }
    }

    store.tournament = null;
    saveGuildStore(interaction.guildId, store);
    const payload = buildTournamentWizardPayload(store);
    return interaction.update({
      content: cleanupFailures
        ? `🗑️ Tournament deleted. ⚠️ ${cleanupFailures} group channel/role(s) couldn't be removed automatically — check the bot's permissions.`
        : '🗑️ Tournament deleted, along with all group channels and roles.',
      ...payload,
    });
  }

  if (id === 'tourney_wizard_delete_cancel') {
    const payload = buildTournamentWizardPayload(store);
    return interaction.update({ content: '', ...payload });
  }
}

// ---------------------------------------------------------------------------
// Modal submits
// ---------------------------------------------------------------------------
async function handleTournamentCreateModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (store.tournament) {
    return interaction.reply({ content: '❌ A tournament already exists.', flags: MessageFlags.Ephemeral });
  }

  const name = interaction.fields.getTextInputValue('name').trim();
  store.tournament = {
    name, open: false, groups: {}, qualified: [], bannedTeams: [],
    slotManagerChannelId: null, registrationChannelId: null, confirmChannelId: null,
    requiredMentions: 4, teamsPerGroup: DEFAULT_GROUP_CAPACITY, totalSlots: null,
    reactions: ['✅', '❌'],
    rounds: {}, // rounds[2..MAX_ROUND] = { roleId, groupSize, groups: {} } — created lazily, see getRound()
  };
  saveGuildStore(interaction.guildId, store);

  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

async function handleAddGroupModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }

  const letter = interaction.fields.getTextInputValue('letter').trim().toUpperCase();
  const capacityRaw = interaction.fields.getTextInputValue('capacity').trim();
  const capacity = parseInt(capacityRaw, 10);

  if (!GROUP_LETTERS.includes(letter)) {
    return interaction.reply({ content: `❌ Group number must be a whole number between 1 and ${MAX_GROUPS}.`, flags: MessageFlags.Ephemeral });
  }
  if (store.tournament.groups[letter]) {
    return interaction.reply({ content: `❌ Group **${letter}** already exists.`, flags: MessageFlags.Ephemeral });
  }
  if (!Number.isInteger(capacity) || capacity < 1 || capacity > MAX_GROUP_CAPACITY) {
    return interaction.reply({ content: `❌ Capacity must be a whole number between 1 and ${MAX_GROUP_CAPACITY}.`, flags: MessageFlags.Ephemeral });
  }

  store.tournament.groups[letter] = { capacity, teams: [] };
  saveGuildStore(interaction.guildId, store);

  const payload = buildTournamentWizardPayload(store);
  await interaction.update(payload);
}

// Auto-creates as many groups as needed to cover `total` teams at `perGroup`
// capacity each, filling the next free letters in order (A, B, C...). The
// last group created absorbs whatever remainder is left over, so the
// capacities always add up to exactly `total` instead of over-provisioning.
async function handleAutoGroupsModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }

  const total = parseInt(interaction.fields.getTextInputValue('total').trim(), 10);
  const perGroup = parseInt(interaction.fields.getTextInputValue('per_group').trim(), 10);

  const result = computeAutoGroups(store.tournament, total, perGroup);
  if (result.error) {
    return interaction.reply({ content: result.error, flags: MessageFlags.Ephemeral });
  }
  const { createdLetters, groupsNeeded } = result;
  saveGuildStore(interaction.guildId, store);

  const payload = buildTournamentWizardPayload(store);
  await interaction.update(payload);
  await interaction.channel.send(
    `⚙️ Auto-created **${groupsNeeded}** group(s) — ${createdLetters.join(', ')} — covering **${total}** teams at up to **${perGroup}** per group.`
  ).catch(() => {});
}

// Public team registration, step 1 — the team-name modal. This only
// validates + stages the team name; the actual slot assignment happens
// once the player has mentioned their 4 teammates and hit Confirm (see
// handleTourneyRegSelectPlayers / handleTourneyRegConfirm below).
async function handleRegisterTeamModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;

  if (!tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }
  if (!tournament.open) {
    return interaction.reply({ content: '❌ Registration is currently closed.', flags: MessageFlags.Ephemeral });
  }

  const team = interaction.fields.getTextInputValue('team').trim();
  if (!team) {
    return interaction.reply({ content: '❌ Team name is required.', flags: MessageFlags.Ephemeral });
  }
  if (isBanned(tournament, team)) {
    return interaction.reply({ content: `❌ **${team}** is banned from registering.`, flags: MessageFlags.Ephemeral });
  }
  if (isDuplicateTeam(tournament, team)) {
    return interaction.reply({ content: `❌ A team named **${team}** is already registered.`, flags: MessageFlags.Ephemeral });
  }

  const requiredMentions = tournament.requiredMentions || 4;
  startRegPending(interaction.user.id, interaction.guildId, { team });

  return interaction.reply({
    content: `Team name set to **${team}**. Now mention the **${requiredMentions} player${requiredMentions === 1 ? '' : 's'}** on your team:`,
    components: [buildMentionPlayersRow(requiredMentions)],
    flags: MessageFlags.Ephemeral,
  });
}

// Public team registration, step 2 — player mentions picked from the
// select menu shown after step 1. Just stages the pick and shows a
// review/confirm screen; nothing is saved to data.json yet.
async function handleTourneyRegSelectPlayers(interaction) {
  const pendingEntry = getRegPending(interaction.user.id);
  if (!pendingEntry) {
    return interaction.reply({ content: `❌ Your session expired or was interrupted. ${RESTART_HINT}`, flags: MessageFlags.Ephemeral });
  }

  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  if (!tournament) {
    clearRegPending(interaction.user.id);
    return interaction.update({ content: '❌ No tournament exists yet.', embeds: [], components: [] });
  }

  const requiredMentions = tournament.requiredMentions || 4;

  // Bots can't be a playing member of a team lineup.
  const botPicked = interaction.users.find(u => u.bot);
  if (botPicked) {
    return interaction.update({
      content: `❌ ${botPicked} is a bot and can't be picked as a player. Mention ${requiredMentions} human player${requiredMentions === 1 ? '' : 's'} below:`,
      embeds: [],
      components: [buildMentionPlayersRow(requiredMentions)],
    });
  }

  // A player can only be on one team's lineup at a time in this tournament.
  const conflict = findTournamentLineupConflict(tournament, interaction.values);
  if (conflict) {
    return interaction.update({
      content: `❌ <@${conflict.conflictId}> is already registered as a player on **${conflict.team}** and can't be picked again. Mention a different lineup below:`,
      embeds: [],
      components: [buildMentionPlayersRow(requiredMentions)],
    });
  }

  updateRegPending(interaction.user.id, { selectedPlayerIds: interaction.values });
  const pending = getRegPending(interaction.user.id);

  return interaction.update({
    content: null,
    embeds: [buildTeamRegPreviewEmbed(pending.data)],
    components: [tourneyConfirmRow()],
  });
}

// ---------------------------------------------------------------------------
// Automatic group/slot assignment for public tournament registration —
// mirrors how scrims auto-assign slots into fixed-size groups, but starting
// at slot 1 (no reserved slots) and capped at MAX_GROUPS groups total.
// ---------------------------------------------------------------------------
function totalRegisteredTeams(tournament) {
  return Object.values(tournament.groups).reduce((sum, g) => sum + g.teams.length, 0);
}

// Finds (auto-creating if needed) the group the next team should land in:
// the first not-yet-full existing group, or the next new group number if
// every existing group is full. Returns null once the tournament's overall
// totalSlots cap or the MAX_GROUPS safety cap has been reached — at which
// point every group must be completely full before registration can grow.
function autoAssignGroup(tournament) {
  if (tournament.totalSlots && totalRegisteredTeams(tournament) >= tournament.totalSlots) {
    return null;
  }
  const capacity = tournament.teamsPerGroup || DEFAULT_GROUP_CAPACITY;
  for (let i = 1; i <= MAX_GROUPS; i++) {
    const key = String(i);
    const group = tournament.groups[key];
    if (!group) {
      tournament.groups[key] = { capacity, teams: [] };
      return key;
    }
    if (group.teams.length < group.capacity) {
      return key;
    }
  }
  return null;
}

// Auto-creates (once per group, same idea as scrims' giveGroupRole/
// giveGroupChannel) a role + a private text channel scoped to that role for
// the given tournament group, and returns the role (or null on failure —
// never throws, a role/channel hiccup should never block registration).
// Unlike the scrim version this channel isn't locked pending an admin
// !open — there's no tournament equivalent of that command, so the group's
// role can read and send from the moment the channel exists.
async function ensureTournamentGroupChannelAndRole(interaction, store, groupKey) {
  const tournament = store.tournament;
  const group = tournament.groups[groupKey];
  const botMember = interaction.guild.members.me;

  const role = await ensureGroupRole(interaction, store, group, `Tournament Group ${groupKey}`, `Auto-created for Group ${groupKey} tournament registration`);

  if (!group.channelId || !interaction.guild.channels.cache.has(group.channelId)) {
    if (!botMember.permissions.has('ManageChannels')) {
      console.error(`[tournament-group-channel] Bot is missing the "Manage Channels" permission in guild ${interaction.guildId}.`);
    } else if (interaction.guild.channels.cache.size >= MAX_GUILD_CHANNELS - SAFETY_MARGIN) {
      console.error(`[tournament-group-channel] Guild ${interaction.guildId} is at/near Discord's ${MAX_GUILD_CHANNELS}-channel cap — skipping auto-create for Group ${groupKey}.`);
    } else {
      try {
        let category = store.settings && store.settings.tournamentGroupChannelsCategoryId
          ? interaction.guild.channels.cache.get(store.settings.tournamentGroupChannelsCategoryId)
          : null;

        if (!category) {
          category = await interaction.guild.channels.create({
            name: '🥇 Tournament Groups',
            type: ChannelType.GuildCategory,
            reason: 'Auto-created to hold per-group tournament channels',
          });
          if (!store.settings) store.settings = {};
          store.settings.tournamentGroupChannelsCategoryId = category.id;
        }

        // Private to this one group's role only — no tournament-wide role
        // is ever added here, so a team in Group 1 can never see Group 2.
        const overwrites = [{ id: interaction.guild.roles.everyone.id, deny: [PermissionFlagsBits.ViewChannel] }];
        if (role) {
          overwrites.push({
            id: role.id,
            allow: [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages, PermissionFlagsBits.ReadMessageHistory],
            deny: [PermissionFlagsBits.CreatePublicThreads, PermissionFlagsBits.CreatePrivateThreads],
          });
        }

        const channel = await interaction.guild.channels.create({
          name: `group-${groupKey}`,
          type: ChannelType.GuildText,
          parent: category.id,
          permissionOverwrites: overwrites,
          reason: `Auto-created for Group ${groupKey} tournament registration`,
        });

        group.channelId = channel.id;
        saveGuildStore(interaction.guildId, store);
        await channel.send(buildTournamentGroupAdminPanelPayload(groupKey)).catch(() => {});
      } catch (err) {
        console.error(`[tournament-group-channel] Failed to auto-create channel for Group ${groupKey} in guild ${interaction.guildId}: ${err.code ?? ''} ${err.message}`);
      }
    }
  } else if (role) {
    // Channel already existed (e.g. made earlier by "Create Channels",
    // before this group had its own role yet) — make sure it's locked down
    // to @everyone and grants only this group's role access, so nobody
    // outside this specific group can see it.
    const existingChannel = interaction.guild.channels.cache.get(group.channelId);
    if (existingChannel && !existingChannel.permissionOverwrites.cache.has(role.id)) {
      await existingChannel.permissionOverwrites.edit(interaction.guild.roles.everyone, { ViewChannel: false }).catch(() => {});
      await existingChannel.permissionOverwrites.edit(role, {
        ViewChannel: true,
        SendMessages: true,
        ReadMessageHistory: true,
        CreatePublicThreads: false,
        CreatePrivateThreads: false,
      }).catch(() => {});
    }
  }

  return role;
}

// ---------------------------------------------------------------------------
// Round 2 — teams qualified out of their Round 1 group (via the "Result"
// button) get funneled into Round 2 groups the same way Round 1 fills up:
// first not-yet-full Round 2 group, or the next new one, capped at
// round2GroupSize teams. Each Round 2 group gets its own role + private
// channel, mirroring the Round 1 group setup above.
// ---------------------------------------------------------------------------
function findRound2Entry(tournament, ownerId) {
  if (!tournament.round2Groups) return null;
  for (const letter of Object.keys(tournament.round2Groups)) {
    const g = tournament.round2Groups[letter];
    const idx = g.teams.findIndex(t => t.ownerId === ownerId);
    if (idx !== -1) return { letter, group: g, idx, team: g.teams[idx] };
  }
  return null;
}

function autoAssignRound2Group(tournament) {
  if (!tournament.round2Groups) tournament.round2Groups = {};
  const capacity = tournament.round2GroupSize || DEFAULT_GROUP_CAPACITY;
  for (let i = 1; i <= MAX_GROUPS; i++) {
    const key = String(i);
    const group = tournament.round2Groups[key];
    if (!group) {
      tournament.round2Groups[key] = { capacity, teams: [] };
      return key;
    }
    if (group.teams.length < group.capacity) {
      return key;
    }
  }
  return null;
}

// Same shape as ensureTournamentGroupChannelAndRole, scoped to Round 2's
// own category/roles/channels so it never collides with Round 1's.
async function ensureRound2GroupChannelAndRole(interaction, store, groupKey) {
  const tournament = store.tournament;
  const group = tournament.round2Groups[groupKey];
  const botMember = interaction.guild.members.me;

  let role = group.roleId ? interaction.guild.roles.cache.get(group.roleId) : null;
  if (!role) {
    if (!botMember.permissions.has('ManageRoles')) {
      console.error(`[tournament-round2-role] Bot is missing the "Manage Roles" permission in guild ${interaction.guildId}.`);
    } else if (interaction.guild.roles.cache.size >= MAX_GUILD_ROLES - SAFETY_MARGIN) {
      console.error(`[tournament-round2-role] Guild ${interaction.guildId} is at/near Discord's ${MAX_GUILD_ROLES}-role cap — skipping auto-create for Round 2 Group ${groupKey}.`);
    } else {
      try {
        role = await interaction.guild.roles.create({
          name: `Round 2 - Group ${groupKey}`,
          mentionable: false,
          reason: `Auto-created for Round 2 Group ${groupKey}`,
        });
        group.roleId = role.id;
        saveGuildStore(interaction.guildId, store);
      } catch (err) {
        console.error(`[tournament-round2-role] Failed to auto-create role for Round 2 Group ${groupKey} in guild ${interaction.guildId}: ${err.code ?? ''} ${err.message}`);
      }
    }
  }

  if (!group.channelId || !interaction.guild.channels.cache.has(group.channelId)) {
    if (!botMember.permissions.has('ManageChannels')) {
      console.error(`[tournament-round2-channel] Bot is missing the "Manage Channels" permission in guild ${interaction.guildId}.`);
    } else if (interaction.guild.channels.cache.size >= MAX_GUILD_CHANNELS - SAFETY_MARGIN) {
      console.error(`[tournament-round2-channel] Guild ${interaction.guildId} is at/near Discord's ${MAX_GUILD_CHANNELS}-channel cap — skipping auto-create for Round 2 Group ${groupKey}.`);
    } else {
      try {
        let category = store.settings && store.settings.tournamentRound2ChannelsCategoryId
          ? interaction.guild.channels.cache.get(store.settings.tournamentRound2ChannelsCategoryId)
          : null;

        if (!category) {
          category = await interaction.guild.channels.create({
            name: '🏆 Round 2 Groups',
            type: ChannelType.GuildCategory,
            reason: 'Auto-created to hold per-group Round 2 tournament channels',
          });
          if (!store.settings) store.settings = {};
          store.settings.tournamentRound2ChannelsCategoryId = category.id;
        }

        const overwrites = [{ id: interaction.guild.roles.everyone.id, deny: [PermissionFlagsBits.ViewChannel] }];
        if (tournament.round2RoleId) {
          overwrites.push({
            id: tournament.round2RoleId,
            allow: [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages, PermissionFlagsBits.ReadMessageHistory],
            deny: [PermissionFlagsBits.CreatePublicThreads, PermissionFlagsBits.CreatePrivateThreads],
          });
        }
        if (role) {
          overwrites.push({
            id: role.id,
            allow: [PermissionFlagsBits.ViewChannel, PermissionFlagsBits.SendMessages, PermissionFlagsBits.ReadMessageHistory],
            deny: [PermissionFlagsBits.CreatePublicThreads, PermissionFlagsBits.CreatePrivateThreads],
          });
        }

        const channel = await interaction.guild.channels.create({
          name: `round2-group-${groupKey}`,
          type: ChannelType.GuildText,
          parent: category.id,
          permissionOverwrites: overwrites,
          reason: `Auto-created for Round 2 Group ${groupKey}`,
        });

        group.channelId = channel.id;
        saveGuildStore(interaction.guildId, store);
      } catch (err) {
        console.error(`[tournament-round2-channel] Failed to auto-create channel for Round 2 Group ${groupKey} in guild ${interaction.guildId}: ${err.code ?? ''} ${err.message}`);
      }
    }
  }

  return role;
}

// Strips a team's Round 2 access (role + group role) and removes it from
// whichever Round 2 group it was sitting in — used when a team is
// un-qualified (Result re-run without them), punished, or self-service
// cancelled after already being promoted to Round 2.
async function removeTeamFromRound2(interaction, store, team) {
  const tournament = store.tournament;
  const entry = findRound2Entry(tournament, team.ownerId);
  if (!entry) return;

  const { group, idx } = entry;
  group.teams.splice(idx, 1);
  saveGuildStore(interaction.guildId, store);

  const targets = new Set([team.ownerId, ...(team.playerIds || [])].filter(Boolean));
  for (const userId of targets) {
    const member = interaction.guild.members.cache.get(userId)
      ?? await interaction.guild.members.fetch(userId).catch(() => null);
    if (!member) continue;
    if (tournament.round2RoleId) await member.roles.remove(tournament.round2RoleId).catch(() => {});
    if (group.roleId) await member.roles.remove(group.roleId).catch(() => {});
  }
}

// Public team registration, step 3 — "Confirm Registration" pressed on the
// review screen. This is where the team actually gets a slot and the
// success role is handed out.
async function handleTourneyRegConfirm(interaction) {
  const pendingEntry = getRegPending(interaction.user.id);
  if (!pendingEntry) {
    return interaction.update({ content: `❌ Your session expired or was interrupted. ${RESTART_HINT}`, embeds: [], components: [] });
  }

  const { team, selectedPlayerIds = [] } = pendingEntry.data;
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;

  if (!tournament) {
    clearRegPending(interaction.user.id);
    return interaction.update({ content: '❌ No tournament exists yet.', embeds: [], components: [] });
  }
  if (!tournament.open) {
    clearRegPending(interaction.user.id);
    return interaction.update({ content: '❌ Registration is currently closed.', embeds: [], components: [] });
  }
  if (isBanned(tournament, team) || isDuplicateTeam(tournament, team)) {
    clearRegPending(interaction.user.id);
    return interaction.update({ content: `❌ **${team}** can no longer be registered. ${RESTART_HINT}`, embeds: [], components: [] });
  }

  // Belt-and-suspenders: re-check for a lineup conflict here too, in case
  // another team grabbed one of these players in the gap between the
  // player-select step and this confirm tap.
  const conflict = findTournamentLineupConflict(tournament, selectedPlayerIds);
  if (conflict) {
    clearRegPending(interaction.user.id);
    return interaction.update({
      content: `❌ <@${conflict.conflictId}> just got locked into **${conflict.team}** by someone else. ${RESTART_HINT}`,
      embeds: [],
      components: [],
    });
  }

  const letter = autoAssignGroup(tournament);

  if (!letter) {
    clearRegPending(interaction.user.id);
    return interaction.update({ content: '❌ Registration is full — every group has reached capacity.', embeds: [], components: [] });
  }

  tournament.groups[letter].teams.push({
    team,
    playerIds: selectedPlayerIds,
    players: selectedPlayerIds.map(id => `<@${id}>`),
    ownerId: interaction.user.id,
  });
  saveGuildStore(interaction.guildId, store);
  clearRegPending(interaction.user.id);

  const slotNumber = tournament.groups[letter].teams.length;

  // Registration grants access to this team's own group channel only —
  // ensureTournamentGroupChannelAndRole creates this group's channel/role
  // on the spot if "Create Channels" hasn't been run yet (e.g. a brand
  // new group auto-created by autoAssignGroup above), then every member
  // of the team gets just that one group's role. A team in Group 1 never
  // gets any role that lets it see Group 2's channel.
  let roleWarning = null;
  const groupRole = await ensureTournamentGroupChannelAndRole(interaction, store, letter).catch(() => null);
  if (groupRole) {
    const roleTargets = new Set([interaction.user.id, ...selectedPlayerIds]);
    let failures = 0;
    for (const userId of roleTargets) {
      const member = interaction.guild.members.cache.get(userId)
        ?? await interaction.guild.members.fetch(userId).catch(() => null);
      if (!member) { failures++; continue; }
      await member.roles.add(groupRole.id).catch(err => {
        failures++;
        console.error(`[tournament-group-role] Failed to add role ${groupRole.id} to ${userId} in guild ${interaction.guildId}: ${err.code ?? ''} ${err.message}`);
      });
    }
    if (failures) {
      roleWarning = `\n\n⚠️ Couldn't give the Group ${letter} role to ${failures === roleTargets.size ? 'everyone' : `${failures} player(s)`} — ask an admin to check my **Manage Roles** permission and that my role sits above <@&${groupRole.id}>.`;
    }
  } else {
    console.warn(`[tournament-group-role] Couldn't create/find a role for Group ${letter} in guild ${interaction.guildId} — check my Manage Roles permission.`);
  }

  if (tournament.confirmChannelId) {
    const confirmChannel = interaction.guild.channels.cache.get(tournament.confirmChannelId);
    if (confirmChannel) {
      const lineup = selectedPlayerIds.map(id => `<@${id}>`).join(' ');
      await confirmChannel.send(`✅ **${team}** (${interaction.user}) registered into **Group ${letter}**, Slot **${slotNumber}**.\n👥 ${lineup}`).catch(() => {});
    }
  }

  return interaction.update({
    content: null,
    embeds: [
      new EmbedBuilder()
        .setTitle('🎯 Registration Complete!')
        .setColor(0x57F287)
        .setDescription(
          `**Team** — ${team}\n` +
          `**Group** — ${letter}\n` +
          `**Slot** — ${slotNumber}\n` +
          `**Players** — ${selectedPlayerIds.map(id => `<@${id}>`).join(' ')}\n\n` +
          (roleWarning || '')
        ),
    ],
    components: [],
  });

}

// "Cancel" pressed on the review screen.
async function handleTourneyRegCancel(interaction) {
  clearRegPending(interaction.user.id);
  return interaction.update({
    content: `❌ Registration cancelled — nothing was saved. ${RESTART_HINT}`,
    embeds: [],
    components: [],
  });
}

async function handleRequiredMentionsModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }
  const value = parseInt(interaction.fields.getTextInputValue('value').trim(), 10);
  if (!Number.isInteger(value) || value < 1 || value > 4) {
    return interaction.reply({ content: '❌ Required Mentions must be a whole number between 1 and 4.', flags: MessageFlags.Ephemeral });
  }
  store.tournament.requiredMentions = value;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

async function handleTeamsPerGroupModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }
  const value = parseInt(interaction.fields.getTextInputValue('value').trim(), 10);
  if (!Number.isInteger(value) || value < 1 || value > MAX_GROUP_CAPACITY) {
    return interaction.reply({ content: `❌ Teams per Group must be a whole number between 1 and ${MAX_GROUP_CAPACITY}.`, flags: MessageFlags.Ephemeral });
  }
  store.tournament.teamsPerGroup = value;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

async function handleTotalSlotsModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }
  const value = parseInt(interaction.fields.getTextInputValue('value').trim(), 10);
  if (!Number.isInteger(value) || value < 1 || value > MAX_TOTAL_SLOTS) {
    return interaction.reply({ content: `❌ Total Slots must be a whole number between 1 and ${MAX_TOTAL_SLOTS}.`, flags: MessageFlags.Ephemeral });
  }
  store.tournament.totalSlots = value;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

async function handleReactionsModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }
  const accept = interaction.fields.getTextInputValue('accept').trim();
  const deny = interaction.fields.getTextInputValue('deny').trim();
  if (!accept || !deny) {
    return interaction.reply({ content: '❌ Both reactions are required.', flags: MessageFlags.Ephemeral });
  }
  // Accepts a plain unicode emoji or a Discord custom-emoji tag, e.g. <:name:123456789012345678>
  // (or <a:name:...> for animated) — the same format used by app-owned/application emojis.
  const isValidEmoji = (value) => /^<a?:\w{2,32}:\d{17,20}>$/.test(value) || [...value].length <= 8;
  if (!isValidEmoji(accept) || !isValidEmoji(deny)) {
    return interaction.reply({
      content: '❌ Enter a plain emoji (e.g. ✅) or a full custom-emoji tag like `<:name:123456789012345678>`.',
      flags: MessageFlags.Ephemeral,
    });
  }
  store.tournament.reactions = [accept, deny];
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

async function handleCreateRegChannelSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
  }
  store.tournament.registrationChannelId = interaction.channels.first().id;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

async function handleCreateConfirmChannelSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
  }
  store.tournament.confirmChannelId = interaction.channels.first().id;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
}

// ---------------------------------------------------------------------------
// Manage Rounds — configure Round 2..MAX_ROUND's promotion role and group
// size. Round 1 needs none of this (it's the tournament's own registration
// settings above); every later round is created lazily the first time a
// team is promoted into it, but the admin can pre-configure its role/size
// here at any point.
// ---------------------------------------------------------------------------
function getRound(tournament, roundNum) {
  if (!tournament.rounds) tournament.rounds = {};
  if (!tournament.rounds[roundNum]) {
    tournament.rounds[roundNum] = { roleId: null, groupSize: DEFAULT_GROUP_CAPACITY, groups: {} };
  }
  return tournament.rounds[roundNum];
}

// tournament.groups IS round 1 — this just picks the right container so
// the rest of the round-aware code can treat every round the same way.
function getRoundGroups(tournament, roundNum) {
  return roundNum <= 1 ? tournament.groups : getRound(tournament, roundNum).groups;
}

function buildRoundListPayload(tournament) {
  const options = [];
  for (let n = 2; n <= MAX_ROUND; n++) {
    const round = tournament.rounds && tournament.rounds[n];
    const teamCount = round ? Object.values(round.groups).reduce((sum, g) => sum + g.teams.length, 0) : 0;
    options.push({
      label: `Round ${n}`,
      description: `${round?.groupSize || DEFAULT_GROUP_CAPACITY} per group · ${teamCount} team${teamCount === 1 ? '' : 's'} promoted · ${round?.roleId ? 'role set' : 'no role set'}`,
      value: String(n),
    });
  }

  const select = new StringSelectMenuBuilder()
    .setCustomId('tourney_round_config_select')
    .setPlaceholder('Pick a round to configure')
    .addOptions(options);

  const embed = new EmbedBuilder()
    .setTitle('🏆 Manage Rounds')
    .setColor(0x5865F2)
    .setDescription(
      `Rounds 2 through ${MAX_ROUND} chain off each other — clicking **Result** in a group's channel promotes its picked teams into the next round, filling that round's groups in order. ` +
      'Pick a round below to set its promotion role and group size (both default automatically if you skip this).'
    );

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

function buildRoundDetailPayload(tournament, roundNum) {
  const round = getRound(tournament, roundNum);
  const embed = new EmbedBuilder()
    .setTitle(`🏆 Round ${roundNum} Settings`)
    .setColor(0x5865F2)
    .addFields(
      { name: 'Role', value: round.roleId ? `<@&${round.roleId}>` : 'Not-Set' },
      { name: 'Group Size', value: String(round.groupSize || DEFAULT_GROUP_CAPACITY) },
    );

  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`tourney_round_config_role:${roundNum}`).setLabel('Set Role').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId(`tourney_round_config_size:${roundNum}`).setLabel('Set Group Size').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_round_config_back').setLabel('Back').setStyle(ButtonStyle.Secondary),
  );

  return { embeds: [embed], components: [row] };
}

function buildRoundSizeModal(round, roundNum) {
  const input = new TextInputBuilder().setCustomId('value').setLabel(`Round ${roundNum} Group Size`).setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(4).setPlaceholder('e.g. 12');
  input.setValue(String(round.groupSize || DEFAULT_GROUP_CAPACITY));
  return new ModalBuilder().setCustomId(`tourney_round_size_modal:${roundNum}`).setTitle(`Round ${roundNum} Group Size`)
    .addComponents(new ActionRowBuilder().addComponents(input));
}

async function handleRoundConfigSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', embeds: [], components: [] });
  }
  const roundNum = parseInt(interaction.values[0], 10);
  await interaction.update({ content: '', ...buildRoundDetailPayload(store.tournament, roundNum) });
}

async function handleRoundRoleSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const [, roundNumStr] = interaction.customId.split(':');
  const roundNum = parseInt(roundNumStr, 10);
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
  }

  const role = interaction.roles.first();
  if (role.managed) {
    return interaction.reply({
      content: '❌ That role is managed by an integration (e.g. a bot or booster role) and can\'t be assigned manually. Pick a regular role instead.',
      flags: MessageFlags.Ephemeral,
    });
  }
  const botMember = interaction.guild.members.me;
  if (role.position >= botMember.roles.highest.position) {
    return interaction.reply({
      content: `❌ I can't assign **${role.name}** — it's positioned above my highest role. Move my bot's role above it in Server Settings → Roles, then pick it again.`,
      flags: MessageFlags.Ephemeral,
    });
  }

  getRound(store.tournament, roundNum).roleId = role.id;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildRoundDetailPayload(store.tournament, roundNum) });
}

async function handleRoundSizeModalSubmit(interaction) {
  const [, roundNumStr] = interaction.customId.split(':');
  const roundNum = parseInt(roundNumStr, 10);
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }
  const value = parseInt(interaction.fields.getTextInputValue('value').trim(), 10);
  if (!Number.isInteger(value) || value < 1 || value > MAX_GROUP_CAPACITY) {
    return interaction.reply({ content: `❌ Group Size must be a whole number between 1 and ${MAX_GROUP_CAPACITY}.`, flags: MessageFlags.Ephemeral });
  }
  getRound(store.tournament, roundNum).groupSize = value;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildRoundDetailPayload(store.tournament, roundNum) });
}

// "Create Channel" (manual) — Channel Format / Category Name modal
// submissions. Both just save into the in-memory pending state (see
// pendingManualChannelCreation below) and re-render the panel; actual
// creation happens on the "Create Channels" button once both are set.
async function handleManualChannelsFormatModalSubmit(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const value = interaction.fields.getTextInputValue('value').trim();
  if (!value) {
    return interaction.reply({ content: '❌ Channel format can\'t be empty.', flags: MessageFlags.Ephemeral });
  }
  const data = setPendingChannelCreation(interaction.guildId, interaction.user.id, { channelFormat: value });
  return interaction.update({ ...buildManualChannelCreationPayload(data) });
}

async function handleManualChannelsCategoryNameModalSubmit(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const value = interaction.fields.getTextInputValue('value').trim();
  if (!value) {
    return interaction.reply({ content: '❌ Category name can\'t be empty.', flags: MessageFlags.Ephemeral });
  }
  const data = setPendingChannelCreation(interaction.guildId, interaction.user.id, { categoryName: value });
  return interaction.update({ ...buildManualChannelCreationPayload(data) });
}


async function handleEditSettingsModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }

  const name = interaction.fields.getTextInputValue('name').trim();
  if (!name) {
    return interaction.reply({ content: '❌ Name is required.', flags: MessageFlags.Ephemeral });
  }

  store.tournament.name = name;
  saveGuildStore(interaction.guildId, store);

  const payload = buildTournamentWizardPayload(store);
  await interaction.update(payload);
}

// Toggling the same team name again flips it back — ban if not banned,
// unban if already banned. Banning also evicts them from whatever group
// they're currently sitting in, since a banned team shouldn't keep a slot.
async function handleBanUnbanModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  if (!tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }

  const teamRaw = interaction.fields.getTextInputValue('team').trim();
  if (!teamRaw) {
    return interaction.reply({ content: '❌ Team name is required.', flags: MessageFlags.Ephemeral });
  }
  const teamKey = teamRaw.toLowerCase();

  if (!tournament.bannedTeams) tournament.bannedTeams = [];
  const idx = tournament.bannedTeams.indexOf(teamKey);

  if (idx === -1) {
    tournament.bannedTeams.push(teamKey);
    let removedFrom = null;
    for (const [letter, group] of Object.entries(tournament.groups)) {
      const before = group.teams.length;
      group.teams = group.teams.filter(t => t.team.toLowerCase() !== teamKey);
      if (group.teams.length !== before) removedFrom = letter;
    }
    tournament.qualified = tournament.qualified.filter(name => name.toLowerCase() !== teamKey);
    saveGuildStore(interaction.guildId, store);
    return interaction.reply({
      content: `🔨 **${teamRaw}** is now banned from registering.${removedFrom ? ` Removed from Group ${removedFrom}.` : ''}`,
      flags: MessageFlags.Ephemeral,
    });
  }

  tournament.bannedTeams.splice(idx, 1);
  saveGuildStore(interaction.guildId, store);
  return interaction.reply({ content: `✅ **${teamRaw}** has been unbanned and can register again.`, flags: MessageFlags.Ephemeral });
}

// Admin version of team registration — picks the exact group instead of
// auto-assigning, and works even while registration is closed.
async function handleManualAddSlotModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  if (!tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }

  const letter = interaction.fields.getTextInputValue('letter').trim().toUpperCase();
  const team = interaction.fields.getTextInputValue('team').trim();
  const player1 = interaction.fields.getTextInputValue('player1').trim();
  const player2 = interaction.fields.getTextInputValue('player2').trim();
  const players34 = interaction.fields.getTextInputValue('players34').trim();

  const group = tournament.groups[letter];
  if (!group) {
    const existing = Object.keys(tournament.groups).join(', ') || 'none yet';
    return interaction.reply({ content: `❌ Group **${letter}** doesn't exist. Current groups: ${existing}.`, flags: MessageFlags.Ephemeral });
  }
  if (!team) {
    return interaction.reply({ content: '❌ Team name is required.', flags: MessageFlags.Ephemeral });
  }
  if (isBanned(tournament, team)) {
    return interaction.reply({ content: `❌ **${team}** is banned from registering.`, flags: MessageFlags.Ephemeral });
  }
  if (isDuplicateTeam(tournament, team)) {
    return interaction.reply({ content: `❌ A team named **${team}** is already registered.`, flags: MessageFlags.Ephemeral });
  }
  if (group.teams.length >= group.capacity) {
    return interaction.reply({ content: `❌ Group **${letter}** is already full (${group.capacity}/${group.capacity}).`, flags: MessageFlags.Ephemeral });
  }

  const players = [player1, player2, ...players34.split(',').map(s => s.trim())].filter(Boolean);
  group.teams.push({ team, players });
  saveGuildStore(interaction.guildId, store);

  return interaction.reply({
    content: `✅ **${team}** added to **Group ${letter}**, Slot **${group.teams.length}**.`,
    flags: MessageFlags.Ephemeral,
  });
}

// ---------------------------------------------------------------------------
// Qualify flow
// ---------------------------------------------------------------------------
function buildQualifySelectPayload(tournament, letter) {
  if (!letter) {
    return { error: "❌ Couldn't tell which group to qualify." };
  }

  const group = tournament.groups[letter];
  if (!group) {
    const existing = Object.keys(tournament.groups).join(', ') || 'none yet';
    return { error: `❌ Group **${letter}** doesn't exist. Current groups: ${existing}.` };
  }
  if (!group.teams.length) {
    return { error: `❌ Group **${letter}** has no registered teams yet.` };
  }
  if (group.teams.length > 25) {
    return { error: `❌ Group **${letter}** has ${group.teams.length} teams — Discord select menus cap at 25 options, so this group can't be shown as one list.` };
  }

  const alreadyQualified = new Set(tournament.qualified);
  const select = new StringSelectMenuBuilder()
    .setCustomId(`qualify_select_teams:${letter}`)
    .setPlaceholder(`Select qualifying teams from Group ${letter}`)
    .setMinValues(0)
    .setMaxValues(group.teams.length)
    .addOptions(group.teams.map((t, idx) => ({
      label: t.team.slice(0, 100),
      value: String(idx),
      default: alreadyQualified.has(t.team),
    })));

  const embed = new EmbedBuilder()
    .setTitle(`✅ Qualify Teams — Group ${letter}`)
    .setColor(0x5865F2)
    .setDescription('Select every team from this group that qualifies, then confirm. Already-qualified teams are pre-checked.');

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

// Picker shown after clicking "Qualify" — pick which group, then hand off
// to the existing per-group team picker (buildQualifySelectPayload).
function buildQualifyGroupSelectPayload(tournament) {
  const letters = Object.keys(tournament.groups);
  if (!letters.length) {
    return { error: '❌ No groups exist yet — add one first.' };
  }

  const select = new StringSelectMenuBuilder()
    .setCustomId('tourney_qualify_group_select')
    .setPlaceholder('Select a group to qualify teams from')
    .setMinValues(1)
    .setMaxValues(1)
    .addOptions(letters.map(letter => ({
      label: `Group ${letter}`,
      description: `${tournament.groups[letter].teams.length}/${tournament.groups[letter].capacity} teams`,
      value: letter,
    })));

  const embed = new EmbedBuilder()
    .setTitle('✅ Qualify Teams')
    .setColor(0x5865F2)
    .setDescription('Pick a group, then choose which of its teams qualify.');

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

async function handleQualifyGroupSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  const [letter] = interaction.values;

  if (!tournament || !tournament.groups[letter]) {
    return interaction.update({ content: '❌ That group no longer exists.', embeds: [], components: [] });
  }

  const payload = buildQualifySelectPayload(tournament, letter);
  if (payload.error) {
    return interaction.update({ content: payload.error, embeds: [], components: [] });
  }
  await interaction.update({ content: '', ...payload });
}

async function handleQualifySelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const [, letter] = interaction.customId.split(':');
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;

  if (!tournament || !tournament.groups[letter]) {
    return interaction.update({ content: '❌ That group no longer exists.', embeds: [], components: [] });
  }

  await interaction.deferUpdate();

  const group = tournament.groups[letter];
  const selectedIdx = new Set(interaction.values.map(v => parseInt(v, 10)));
  const selectedTeams = group.teams.filter((t, idx) => selectedIdx.has(idx));
  const selectedNames = selectedTeams.map(t => t.team);
  const selectedOwnerIds = new Set(selectedTeams.map(t => t.ownerId));

  // Re-running qualify on the same group cleanly replaces its previous
  // picks rather than piling up duplicates: drop every team from this
  // group out of the qualified list first, then add back only what's
  // selected now.
  const groupTeamNames = new Set(group.teams.map(t => t.team));
  tournament.qualified = tournament.qualified.filter(name => !groupTeamNames.has(name));
  tournament.qualified.push(...selectedNames);
  saveGuildStore(interaction.guildId, store);

  // Un-promote anyone from this group who was in Round 2 but isn't
  // selected this time (e.g. Result re-run with a smaller pick).
  for (const t of group.teams) {
    if (!selectedOwnerIds.has(t.ownerId)) {
      await removeTeamFromRound2(interaction, store, t);
    }
  }

  // Promote newly-qualified teams into Round 2 — first not-yet-full
  // Round 2 group, or the next new one, filling in order the same way
  // Round 1 registration does. Already-promoted teams (Result re-run
  // with the same picks) are left where they are.
  const promoted = [];
  const failed = [];
  for (const t of selectedTeams) {
    if (findRound2Entry(tournament, t.ownerId)) continue;

    const r2Letter = autoAssignRound2Group(tournament);
    if (!r2Letter) { failed.push(t.team); continue; }
    tournament.round2Groups[r2Letter].teams.push(t);
    saveGuildStore(interaction.guildId, store);

    const r2Role = await ensureRound2GroupChannelAndRole(interaction, store, r2Letter).catch(() => null);

    const targets = new Set([t.ownerId, ...(t.playerIds || [])].filter(Boolean));
    for (const userId of targets) {
      const member = interaction.guild.members.cache.get(userId)
        ?? await interaction.guild.members.fetch(userId).catch(() => null);
      if (!member) continue;
      if (tournament.round2RoleId) await member.roles.add(tournament.round2RoleId).catch(() => {});
      if (r2Role) await member.roles.add(r2Role.id).catch(() => {});
    }
    promoted.push(`**${t.team}** → Round 2 Group ${r2Letter}`);
  }

  const lines = [`✅ Group **${letter}** qualifiers updated: ${selectedNames.length ? selectedNames.map(t => `**${t}**`).join(', ') : '_none selected_'}`];
  if (promoted.length) lines.push(`🏆 ${promoted.join('\n🏆 ')}`);
  if (failed.length) lines.push(`⚠️ Couldn't find/create a Round 2 slot for: ${failed.map(t => `**${t}**`).join(', ')} — check my Manage Roles/Manage Channels permissions.`);
  if (promoted.length && !tournament.round2RoleId) lines.push('ℹ️ No Round 2 role is set — set one via **Manage Rounds** → Round 2 → Set Role.');

  await interaction.editReply({ content: lines.join('\n'), embeds: [], components: [] });
}

// ---------------------------------------------------------------------------
// Per-group admin panel — Publish Slot List / Punish Team (posted
// automatically in each group's own channel, see
// buildTournamentGroupAdminPanelPayload)
// ---------------------------------------------------------------------------

// "Publish Slot List" — posts the group's current slot list right into its
// own channel (where the button lives), and also mirrors it to the
// Slot-Manager channel if one's configured, same as the old top-level
// Slot List picker did.
async function handleTourneyGroupPublish(interaction, store, letter) {
  const tournament = store.tournament;
  if (!tournament || !tournament.groups[letter]) {
    return interaction.reply({ content: '❌ That group no longer exists.', flags: MessageFlags.Ephemeral });
  }

  const embed = buildTournamentSlotListEmbed(tournament, letter);
  await interaction.channel.send({ embeds: [embed] }).catch(() => {});

  if (tournament.slotManagerChannelId && tournament.slotManagerChannelId !== interaction.channelId) {
    const publishChannel = interaction.guild.channels.cache.get(tournament.slotManagerChannelId);
    if (publishChannel) await publishChannel.send({ embeds: [embed] }).catch(() => {});
  }

  return interaction.reply({
    content: `✅ Slot list published for Group **${letter}**.`,
    flags: MessageFlags.Ephemeral,
  });
}

// Team picker shown by "Punish Team" — same shape as the Qualify picker,
// just for banning instead.
function buildTournamentPunishSelectPayload(tournament, letter) {
  const group = tournament.groups[letter];
  if (!group) {
    return { error: `❌ Group **${letter}** doesn't exist.` };
  }
  if (!group.teams.length) {
    return { error: `❌ Group **${letter}** has no registered teams yet.` };
  }
  if (group.teams.length > 25) {
    return { error: `❌ Group **${letter}** has ${group.teams.length} teams — Discord select menus cap at 25 options.` };
  }

  const select = new StringSelectMenuBuilder()
    .setCustomId(`tourney_punish_select_teams:${letter}`)
    .setPlaceholder(`Select team(s) to punish from Group ${letter}`)
    .setMinValues(1)
    .setMaxValues(group.teams.length)
    .addOptions(group.teams.map((t, idx) => ({ label: t.team.slice(0, 100), value: String(idx) })));

  const embed = new EmbedBuilder()
    .setTitle(`🔨 Punish Teams — Group ${letter}`)
    .setColor(0xED4245)
    .setDescription('Select every team to punish — each is banned from re-registering, removed from this group, and loses the success role and this group\'s role.');

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

// On submit: bans the picked team name(s), evicts them from the group, and
// strips the success role + this group's role from every player on file
// for that team (owner + selected lineup). Never lets one player's role
// removal fail block the rest.
async function handleTournamentPunishSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const [, letter] = interaction.customId.split(':');
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;

  if (!tournament || !tournament.groups[letter]) {
    return interaction.update({ content: '❌ That group no longer exists.', embeds: [], components: [] });
  }

  const group = tournament.groups[letter];
  const indices = new Set(interaction.values.map(v => parseInt(v, 10)));
  const punishedTeams = group.teams.filter((t, idx) => indices.has(idx));

  if (!punishedTeams.length) {
    return interaction.update({ content: '❌ Nothing selected.', embeds: [], components: [] });
  }

  await interaction.deferUpdate();

  if (!tournament.bannedTeams) tournament.bannedTeams = [];
  const groupRoleId = group.roleId;
  const lines = [];

  for (const team of punishedTeams) {
    if (!tournament.bannedTeams.includes(team.team.toLowerCase())) {
      tournament.bannedTeams.push(team.team.toLowerCase());
    }
    await removeTeamFromRound2(interaction, store, team);
    const targets = new Set([team.ownerId, ...(team.playerIds || [])].filter(Boolean));
    for (const userId of targets) {
      const member = interaction.guild.members.cache.get(userId)
        ?? await interaction.guild.members.fetch(userId).catch(() => null);
      if (!member) continue;
      if (groupRoleId) await member.roles.remove(groupRoleId).catch(() => {});
    }
    lines.push(`🔨 **${team.team}** banned and removed from Group ${letter}.`);
  }

  group.teams = group.teams.filter(t => !punishedTeams.includes(t));
  tournament.qualified = tournament.qualified.filter(name => !punishedTeams.some(t => t.team === name));
  saveGuildStore(interaction.guildId, store);

  await interaction.editReply({ content: lines.join('\n'), embeds: [], components: [] });
}

// ---------------------------------------------------------------------------
// Slot list flow
// ---------------------------------------------------------------------------
// Picker shown after clicking "Slot List" — pick which group to view.
// Returns { error } when there's nothing to pick from yet.
function buildSlotListGroupSelectPayload(tournament) {
  const letters = Object.keys(tournament.groups);
  if (!letters.length) {
    return { error: '❌ No groups exist yet — add one first.' };
  }

  const select = new StringSelectMenuBuilder()
    .setCustomId('tourney_slotlist_select')
    .setPlaceholder('Select a group to view its slot list')
    .setMinValues(1)
    .setMaxValues(1)
    .addOptions(letters.map(letter => ({
      label: `Group ${letter}`,
      description: `${tournament.groups[letter].teams.length}/${tournament.groups[letter].capacity} teams`,
      value: letter,
    })));

  const embed = new EmbedBuilder()
    .setTitle('🔢 Tournament Slot List')
    .setColor(0x5865F2)
    .setDescription('Pick a group — its slot list is generated automatically from current registrations.');

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

async function handleSlotListSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  const [letter] = interaction.values;

  if (!tournament || !tournament.groups[letter]) {
    return interaction.update({ content: '❌ That group no longer exists.', embeds: [], components: [] });
  }

  const embed = buildTournamentSlotListEmbed(tournament, letter);
  await interaction.update({ content: '', embeds: [embed], components: [] });

  // Also publish to the configured Slot-Manager channel, if set — this is
  // what that channel is for (a running, publicly-visible copy of slot
  // lists) separate from this ephemeral admin view.
  if (tournament.slotManagerChannelId) {
    const publishChannel = interaction.guild.channels.cache.get(tournament.slotManagerChannelId);
    if (publishChannel) {
      await publishChannel.send({ embeds: [embed] }).catch(() => {});
    }
  }
}

// ---------------------------------------------------------------------------
// Cancel Slots flow
// ---------------------------------------------------------------------------
function buildCancelGroupSelectPayload(tournament) {
  const letters = Object.keys(tournament.groups).filter(l => tournament.groups[l].teams.length > 0);
  if (!letters.length) {
    return { error: '❌ No registered teams in any group yet.' };
  }

  const select = new StringSelectMenuBuilder()
    .setCustomId('tourney_cancel_group_select')
    .setPlaceholder('Select a group to cancel slots from')
    .setMinValues(1)
    .setMaxValues(1)
    .addOptions(letters.map(letter => ({
      label: `Group ${letter}`,
      description: `${tournament.groups[letter].teams.length}/${tournament.groups[letter].capacity} teams`,
      value: letter,
    })));

  const embed = new EmbedBuilder()
    .setTitle('🗑️ Cancel Slots')
    .setColor(0xED4245)
    .setDescription('Pick a group, then choose which team(s) to remove.');

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

function buildCancelTeamsSelectPayload(tournament, letter) {
  const group = tournament.groups[letter];
  if (!group) {
    return { error: `❌ Group **${letter}** doesn't exist.` };
  }
  if (!group.teams.length) {
    return { error: `❌ Group **${letter}** has no registered teams.` };
  }
  if (group.teams.length > 25) {
    return { error: `❌ Group **${letter}** has ${group.teams.length} teams — over Discord's 25-option select limit.` };
  }

  const select = new StringSelectMenuBuilder()
    .setCustomId(`cancel_select_teams:${letter}`)
    .setPlaceholder(`Select team(s) to remove from Group ${letter}`)
    .setMinValues(1)
    .setMaxValues(group.teams.length)
    .addOptions(group.teams.map((t, idx) => ({ label: t.team.slice(0, 100), value: String(idx) })));

  const embed = new EmbedBuilder()
    .setTitle(`🗑️ Cancel Slots — Group ${letter}`)
    .setColor(0xED4245)
    .setDescription('Select every team to remove, then confirm.');

  return { embeds: [embed], components: [new ActionRowBuilder().addComponents(select)] };
}

async function handleCancelGroupSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  const [letter] = interaction.values;

  if (!tournament || !tournament.groups[letter]) {
    return interaction.update({ content: '❌ That group no longer exists.', embeds: [], components: [] });
  }

  const payload = buildCancelTeamsSelectPayload(tournament, letter);
  if (payload.error) {
    return interaction.update({ content: payload.error, embeds: [], components: [] });
  }
  await interaction.update({ content: '', ...payload });
}

async function handleCancelTeamsSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const [, letter] = interaction.customId.split(':');
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;

  if (!tournament || !tournament.groups[letter]) {
    return interaction.update({ content: '❌ That group no longer exists.', embeds: [], components: [] });
  }

  const group = tournament.groups[letter];
  const removeIndices = new Set(interaction.values.map(v => parseInt(v, 10)));
  const removedTeams = group.teams.filter((t, idx) => removeIndices.has(idx));
  const removedNames = removedTeams.map(t => t.team);
  group.teams = group.teams.filter((t, idx) => !removeIndices.has(idx));

  await interaction.deferUpdate();
  for (const team of removedTeams) {
    await removeTeamFromRound2(interaction, store, team);
  }

  const removedSet = new Set(removedNames);
  tournament.qualified = tournament.qualified.filter(name => !removedSet.has(name));
  saveGuildStore(interaction.guildId, store);

  await interaction.editReply({
    content: `🗑️ Removed from Group ${letter}: ${removedNames.map(n => `**${n}**`).join(', ')}`,
    embeds: [],
    components: [],
  });
}

// ---------------------------------------------------------------------------
// Slot-Manager channel select (ChannelSelectMenu)
// ---------------------------------------------------------------------------
async function handleSlotManagerChannelSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
  }

  const channel = interaction.channels.first();
  store.tournament.slotManagerChannelId = channel.id;
  saveGuildStore(interaction.guildId, store);

  await channel.send(buildSlotSelfServicePanelPayload()).catch(() => {});

  await interaction.update({ content: `✅ Slot-Manager channel set to ${channel} — the self-service panel has been posted there.`, components: [] });
}

// Finds the team (if any) a given user belongs to — as the owner or as a
// listed player — across every group in the tournament. A player can only
// ever be on one team at a time (enforced at registration), so the first
// match is the only match.
function findUserTournamentEntry(tournament, userId) {
  for (const letter of Object.keys(tournament.groups)) {
    const group = tournament.groups[letter];
    const idx = group.teams.findIndex(t => t.ownerId === userId || (t.playerIds || []).includes(userId));
    if (idx !== -1) return { letter, group, idx, team: group.teams[idx] };
  }
  return null;
}

// Public self-service panel — posted automatically in the configured
// Slot-Manager channel. Lets a registered player cancel their own slot,
// check which group they're in, or rename their own team, without needing
// an admin.
function buildSlotSelfServicePanelPayload() {
  const embed = new EmbedBuilder()
    .setTitle('🎯 Tourney Slot Manager')
    .setColor(0x5865F2)
    .setDescription(
      '• Click **Cancel My Slot** below to cancel your slot.\n' +
      '• Click **My Groups** to see which group your team is in.\n' +
      '• Click **Change Team Name** if you want to update your team\'s name.\n\n' +
      '*Note that slot cancel is irreversible.*'
    );

  const row1 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_selfservice_cancel').setLabel('Cancel My Slot').setEmoji('🗑️').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId('tourney_wizard_selfservice_my_groups').setLabel('My Groups').setEmoji('🗂️').setStyle(ButtonStyle.Success),
  );
  const row2 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_selfservice_change_name').setLabel('Change Team Name').setEmoji('✏️').setStyle(ButtonStyle.Primary),
  );

  return { embeds: [embed], components: [row1, row2] };
}

// Submit handler for the "Change Team Name" modal above — renames the
// player's own team in place (same duplicate-name check registration
// uses) and keeps the qualified list in sync if that team already
// qualified under its old name.
async function handleSelfServiceChangeNameModalSubmit(interaction) {
  const store = getGuildStore(interaction.guildId);
  const tournament = store.tournament;
  if (!tournament) {
    return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
  }

  const entry = findUserTournamentEntry(tournament, interaction.user.id);
  if (!entry) {
    return interaction.reply({ content: "❌ You're not registered for this tournament.", flags: MessageFlags.Ephemeral });
  }

  const newName = interaction.fields.getTextInputValue('team').trim();
  if (!newName) {
    return interaction.reply({ content: '❌ Team name cannot be empty.', flags: MessageFlags.Ephemeral });
  }

  const oldName = entry.team.team;
  if (newName.toLowerCase() !== oldName.toLowerCase()) {
    const taken = Object.values(tournament.groups).some(g => g.teams.some(t => t.team.toLowerCase() === newName.toLowerCase()));
    if (taken) {
      return interaction.reply({ content: `❌ A team named **${newName}** is already registered.`, flags: MessageFlags.Ephemeral });
    }
  }

  entry.team.team = newName;
  const qIdx = tournament.qualified.indexOf(oldName);
  if (qIdx !== -1) tournament.qualified[qIdx] = newName;
  saveGuildStore(interaction.guildId, store);

  return interaction.reply({ content: `✅ Team name updated to **${newName}**.`, flags: MessageFlags.Ephemeral });
}

async function handleRegisterPanelChannelSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }

  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
  }

  const channel = interaction.channels.first();
  const me = interaction.guild.members.me;
  if (!channel.permissionsFor(me).has(['ViewChannel', 'SendMessages', 'EmbedLinks'])) {
    return interaction.update({
      content: `❌ I don't have permission to post in ${channel}. I need **View Channel**, **Send Messages**, and **Embed Links** there.`,
      components: [],
    });
  }

  await channel.send(buildTournamentRegisterPanelPayload(store.tournament));
  await interaction.update({ content: `✅ Registration panel posted in ${channel}.`, components: [] });
}

module.exports = {
  buildTournamentWizardPayload,
  buildTournamentRegisterPanelPayload,
  handleTournamentWizardButton,
  handleTournamentCreateModalSubmit,
  handleAddGroupModalSubmit,
  handleAutoGroupsModalSubmit,
  handleRegisterTeamModalSubmit,
  handleTourneyRegSelectPlayers,
  handleEditSettingsModalSubmit,
  handleBanUnbanModalSubmit,
  handleManualAddSlotModalSubmit,
  buildQualifySelectPayload,
  handleQualifySelect,
  buildQualifyGroupSelectPayload,
  handleQualifyGroupSelect,
  buildSlotListGroupSelectPayload,
  handleSlotListSelect,
  handleCancelGroupSelect,
  handleCancelTeamsSelect,
  handleSlotManagerChannelSelect,
  handleRegisterPanelChannelSelect,
  handleRequiredMentionsModalSubmit,
  handleTeamsPerGroupModalSubmit,
  handleTotalSlotsModalSubmit,
  handleReactionsModalSubmit,
  handleCreateRegChannelSelect,
  handleCreateConfirmChannelSelect,
  handleTournamentPunishSelect,
  handleManualChannelsFormatModalSubmit,
  handleManualChannelsCategoryNameModalSubmit,
  handleSelfServiceChangeNameModalSubmit,
};
