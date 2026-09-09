const {
  EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle,
  ModalBuilder, TextInputBuilder, TextInputStyle, StringSelectMenuBuilder,
  ChannelSelectMenuBuilder, RoleSelectMenuBuilder, ChannelType, AttachmentBuilder,
  MessageFlags, PermissionFlagsBits,
} = require('discord.js');
const ExcelJS = require('exceljs');
const { getGuildStore, saveGuildStore } = require('./storage');
const { buildGroupsEmbed, buildTournamentSlotListEmbed } = require('./embeds');

const GROUP_LETTERS = 'ABCDEFGHIJKL'.split('');

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

    if (groupCount) {
      const groupLines = Object.entries(tournament.groups)
        .map(([letter, g]) => `**${letter}** — ${g.teams.length}/${g.capacity}${g.channelId ? ' 📺' : ''}`)
        .join('  •  ');
      embed.addFields({ name: 'Group Capacity', value: groupLines });
    }
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

// Sub-panel behind "Manage Groups" — keeps the top-level panel matching the
// requested layout (one button) while still exposing the group tools.
function buildManageGroupsSubmenuPayload(tournament) {
  const embed = new EmbedBuilder()
    .setTitle('🗂️ Manage Groups')
    .setColor(0x5865F2)
    .setDescription('Add a group, auto-create a batch of them, view current groups, check a slot list, or qualify teams for the next stage.');

  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_wizard_add_group').setLabel('Add Group').setEmoji('➕').setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId('tourney_wizard_auto_groups').setLabel('Auto Groups').setEmoji('⚙️').setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId('tourney_wizard_view_groups').setLabel('View Groups').setEmoji('📋').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_wizard_slotlist').setLabel('Slot List').setEmoji('🔢').setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId('tourney_wizard_qualify').setLabel('Qualify').setEmoji('✅').setStyle(ButtonStyle.Primary),
  );

  return { embeds: [embed], components: [row] };
}

// Public panel — this is the one meant to live in a #register-style
// channel where players (not admins) click to sign their team up. It's
// just an embed + the same Register Team button the admin panel used to
// carry, but posted on its own so players never see admin controls.
function buildTournamentRegisterPanelPayload(tournament) {
  const groupCount = Object.keys(tournament.groups).length;
  const teamCount = Object.values(tournament.groups).reduce((sum, g) => sum + g.teams.length, 0);
  const capacity = Object.values(tournament.groups).reduce((sum, g) => sum + g.capacity, 0);

  const embed = new EmbedBuilder()
    .setTitle(`🥇 ${tournament.name} — Team Registration`)
    .setColor(tournament.open ? 0x57F287 : 0xED4245)
    .setDescription(
      tournament.open
        ? 'Click **Register Team** below and fill in your team name + player IGNs. You\'ll be auto-assigned to whichever group still has room.'
        : '🔒 Registration is currently closed.'
    )
    .addFields(
      { name: 'Status', value: tournament.open ? '🟢 Open' : '🔴 Closed', inline: true },
      { name: 'Slots Filled', value: groupCount ? `${teamCount}/${capacity}` : 'No groups yet', inline: true },
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
      { name: 'C. Success Role', value: tournament.successRoleId ? `<@&${tournament.successRoleId}>` : 'Not-Set' },
      { name: 'D. Required Mentions', value: String(tournament.requiredMentions ?? 4) },
      { name: 'E. Teams per Group', value: tournament.teamsPerGroup ? String(tournament.teamsPerGroup) : 'Not-Set' },
      { name: 'F. Total Slots', value: tournament.totalSlots ? String(tournament.totalSlots) : 'Not-Set' },
      { name: 'G. Reactions', value: (tournament.reactions && tournament.reactions.length ? tournament.reactions : ['✅', '❌']).join(' , ') },
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
    new ButtonBuilder().setCustomId('tourney_create_settings_g').setLabel('G').setStyle(ButtonStyle.Primary),
  );
  const row3 = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId('tourney_create_settings_back').setLabel('Go Back').setStyle(ButtonStyle.Danger),
    new ButtonBuilder().setCustomId('tourney_create_settings_save').setLabel('Save').setStyle(ButtonStyle.Success),
  );

  return { embeds: [embed], components: [row1, row2, row3] };
}


function buildHelpEmbed() {
  return new EmbedBuilder()
    .setTitle('❓ Tournament Panel Help')
    .setColor(0x5865F2)
    .setDescription([
      '**Start/Pause Reg** — open or close team registration',
      '**Manage Groups** — add groups, auto-create a batch, view groups, slot lists, qualify teams',
      '**Edit Settings** — rename the tournament',
      '**Create Channels** — auto-create a text channel per group',
      '**Ban/Unban** — block or unblock a team name from registering',
      '**Cancel Slots** — remove a registered team from its group',
      '**Manually Add Slot** — force-register a team into a specific group, bypassing auto-assign',
      '**Post Register Panel** — posts the public registration panel in this channel, for players to register themselves',
      '**Slot-Manager channel** — pick a channel where slot lists get published automatically',
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
        new TextInputBuilder().setCustomId('letter').setLabel('Group letter (A-L)').setStyle(TextInputStyle.Short)
          .setPlaceholder('A').setRequired(true).setMaxLength(1)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('capacity').setLabel('Team capacity').setStyle(TextInputStyle.Short)
          .setPlaceholder('e.g. 16').setRequired(true).setMaxLength(4)
      ),
    );
}

function buildAutoGroupsModal(tournament) {
  const totalInput = new TextInputBuilder().setCustomId('total').setLabel('Total teams expected').setStyle(TextInputStyle.Short)
    .setPlaceholder('e.g. 64').setRequired(true).setMaxLength(4);
  if (tournament && tournament.totalSlots) totalInput.setValue(String(tournament.totalSlots));

  const perGroupInput = new TextInputBuilder().setCustomId('per_group').setLabel('Teams per group').setStyle(TextInputStyle.Short)
    .setPlaceholder('e.g. 16').setRequired(true).setMaxLength(3);
  if (tournament && tournament.teamsPerGroup) perGroupInput.setValue(String(tournament.teamsPerGroup));

  return new ModalBuilder()
    .setCustomId('tourney_wizard_auto_groups_modal')
    .setTitle('Auto-Create Groups')
    .addComponents(
      new ActionRowBuilder().addComponents(totalInput),
      new ActionRowBuilder().addComponents(perGroupInput),
    );
}

function buildRegisterTeamModal() {
  return new ModalBuilder()
    .setCustomId('tourney_wizard_register_modal')
    .setTitle('Register Team')
    .addComponents(
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
        new TextInputBuilder().setCustomId('player3').setLabel('Player 3 IGN').setStyle(TextInputStyle.Short)
          .setRequired(false).setMaxLength(40)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId('player4').setLabel('Player 4 IGN').setStyle(TextInputStyle.Short)
          .setRequired(false).setMaxLength(40)
      ),
    );
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
        new TextInputBuilder().setCustomId('letter').setLabel('Group letter').setStyle(TextInputStyle.Short)
          .setPlaceholder('A').setRequired(true).setMaxLength(1)
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
    .setRequired(true).setMaxLength(3).setPlaceholder('e.g. 16');
  if (tournament.teamsPerGroup) input.setValue(String(tournament.teamsPerGroup));
  return new ModalBuilder().setCustomId('tourney_create_settings_e_modal').setTitle('Teams per Group')
    .addComponents(new ActionRowBuilder().addComponents(input));
}

function buildTotalSlotsModal(tournament) {
  const input = new TextInputBuilder().setCustomId('value').setLabel('Total Slots').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(4).setPlaceholder('e.g. 64');
  if (tournament.totalSlots) input.setValue(String(tournament.totalSlots));
  return new ModalBuilder().setCustomId('tourney_create_settings_f_modal').setTitle('Total Slots')
    .addComponents(new ActionRowBuilder().addComponents(input));
}

function buildReactionsModal(tournament) {
  const reactions = tournament.reactions && tournament.reactions.length ? tournament.reactions : ['✅', '❌'];
  const accept = new TextInputBuilder().setCustomId('accept').setLabel('Accept reaction (emoji)').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(10).setValue(reactions[0] || '✅');
  const deny = new TextInputBuilder().setCustomId('deny').setLabel('Deny reaction (emoji)').setStyle(TextInputStyle.Short)
    .setRequired(true).setMaxLength(10).setValue(reactions[1] || '❌');
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
    if (!store.tournament) return interaction.update({ content: '❌ No tournament exists yet.', embeds: [], components: [] });
    const select = new RoleSelectMenuBuilder().setCustomId('tourney_create_role_select').setPlaceholder('Choose the success role');
    return interaction.update({ content: 'C. Pick the success role:', embeds: [], components: [new ActionRowBuilder().addComponents(select)] });
  }

  if (id === 'tourney_create_settings_d') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildRequiredMentionsModal(store.tournament));
  }

  if (id === 'tourney_create_settings_e') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildTeamsPerGroupModal(store.tournament));
  }

  if (id === 'tourney_create_settings_f') {
    if (!store.tournament) return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    return interaction.showModal(buildTotalSlotsModal(store.tournament));
  }

  if (id === 'tourney_create_settings_g') {
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
    return interaction.reply({ ...buildManageGroupsSubmenuPayload(store.tournament), flags: MessageFlags.Ephemeral });
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
    const letters = Object.keys(store.tournament.groups);
    if (!letters.length) {
      return interaction.reply({ content: '❌ Add a group first.', flags: MessageFlags.Ephemeral });
    }
    const me = interaction.guild.members.me;
    if (!me.permissions.has(PermissionFlagsBits.ManageChannels)) {
      return interaction.reply({ content: '❌ I need the **Manage Channels** permission to do that.', flags: MessageFlags.Ephemeral });
    }

    await interaction.deferReply({ flags: MessageFlags.Ephemeral });
    const created = [];
    for (const letter of letters) {
      const group = store.tournament.groups[letter];
      if (group.channelId && interaction.guild.channels.cache.has(group.channelId)) continue;
      try {
        const channel = await interaction.guild.channels.create({
          name: `group-${letter.toLowerCase()}`,
          type: ChannelType.GuildText,
          parent: interaction.channel.parentId || undefined,
          reason: `Tournament group channel created by ${interaction.user.tag}`,
        });
        group.channelId = channel.id;
        created.push(`<#${channel.id}>`);
      } catch (err) {
        console.error(`[tournament] Failed to create channel for group ${letter}:`, err.message);
      }
    }
    saveGuildStore(interaction.guildId, store);
    return interaction.editReply({
      content: created.length
        ? `✅ Created: ${created.join(', ')}`
        : 'ℹ️ Every group already has a channel (or channel creation failed — check my permissions).',
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
        sheet.addRow({ group: letter, slot: idx + 1, team: t.team, players: (t.players || []).join(', ') });
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

  if (id === 'tourney_wizard_add_group') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    if (Object.keys(store.tournament.groups).length >= GROUP_LETTERS.length) {
      return interaction.reply({ content: `❌ Max ${GROUP_LETTERS.length} groups (A-L) reached.`, flags: MessageFlags.Ephemeral });
    }
    return interaction.showModal(buildAddGroupModal());
  }

  if (id === 'tourney_wizard_auto_groups') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    if (Object.keys(store.tournament.groups).length >= GROUP_LETTERS.length) {
      return interaction.reply({ content: `❌ Max ${GROUP_LETTERS.length} groups (A-L) reached.`, flags: MessageFlags.Ephemeral });
    }
    return interaction.showModal(buildAutoGroupsModal(store.tournament));
  }

  if (id === 'tourney_wizard_toggle') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    if (!store.tournament.open && Object.keys(store.tournament.groups).length === 0) {
      return interaction.reply({ content: '❌ Add at least one group before opening registration.', flags: MessageFlags.Ephemeral });
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

  if (id === 'tourney_wizard_view_groups') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ embeds: [buildGroupsEmbed(store.tournament)], flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_slotlist') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const payload = buildSlotListGroupSelectPayload(store.tournament);
    if (payload.error) {
      return interaction.reply({ content: payload.error, flags: MessageFlags.Ephemeral });
    }
    return interaction.reply({ ...payload, flags: MessageFlags.Ephemeral });
  }

  if (id === 'tourney_wizard_qualify') {
    if (!store.tournament) {
      return interaction.reply({ content: '❌ No tournament exists yet.', flags: MessageFlags.Ephemeral });
    }
    const payload = buildQualifyGroupSelectPayload(store.tournament);
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
    store.tournament = null;
    saveGuildStore(interaction.guildId, store);
    const payload = buildTournamentWizardPayload(store);
    return interaction.update({ content: '🗑️ Tournament deleted.', ...payload });
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
    successRoleId: null, requiredMentions: 4, teamsPerGroup: null, totalSlots: null,
    reactions: ['✅', '❌'],
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
    return interaction.reply({ content: `❌ Group letter must be a single letter A-${GROUP_LETTERS[GROUP_LETTERS.length - 1]}.`, flags: MessageFlags.Ephemeral });
  }
  if (store.tournament.groups[letter]) {
    return interaction.reply({ content: `❌ Group **${letter}** already exists.`, flags: MessageFlags.Ephemeral });
  }
  if (!Number.isInteger(capacity) || capacity < 1 || capacity > 64) {
    return interaction.reply({ content: '❌ Capacity must be a whole number between 1 and 64.', flags: MessageFlags.Ephemeral });
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

  if (!Number.isInteger(total) || total < 1 || total > 768) {
    return interaction.reply({ content: '❌ Total teams must be a whole number between 1 and 768.', flags: MessageFlags.Ephemeral });
  }
  if (!Number.isInteger(perGroup) || perGroup < 1 || perGroup > 64) {
    return interaction.reply({ content: '❌ Teams per group must be a whole number between 1 and 64.', flags: MessageFlags.Ephemeral });
  }

  const groupsNeeded = Math.ceil(total / perGroup);
  const freeLetters = GROUP_LETTERS.filter(l => !store.tournament.groups[l]);

  if (groupsNeeded > freeLetters.length) {
    return interaction.reply({
      content: `❌ That needs **${groupsNeeded}** new group(s), but only **${freeLetters.length}** letter slot(s) are free (max ${GROUP_LETTERS.length} groups total). Raise "teams per group" or delete an unused group first.`,
      flags: MessageFlags.Ephemeral,
    });
  }

  const createdLetters = [];
  let remaining = total;
  for (let i = 0; i < groupsNeeded; i++) {
    const letter = freeLetters[i];
    const capacity = Math.min(perGroup, remaining);
    store.tournament.groups[letter] = { capacity, teams: [] };
    remaining -= capacity;
    createdLetters.push(letter);
  }
  saveGuildStore(interaction.guildId, store);

  const payload = buildTournamentWizardPayload(store);
  await interaction.update(payload);
  await interaction.channel.send(
    `⚙️ Auto-created **${groupsNeeded}** group(s) — ${createdLetters.join(', ')} — covering **${total}** teams at up to **${perGroup}** per group.`
  ).catch(() => {});
}

// Public team registration — auto-assigns to the first (alphabetically)
// group that still has a free slot, rather than making the player pick one.
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

  const players = ['player1', 'player2', 'player3', 'player4']
    .map(id => interaction.fields.getTextInputValue(id).trim())
    .filter(Boolean);

  const requiredMentions = tournament.requiredMentions || 4;
  if (players.length < requiredMentions) {
    return interaction.reply({
      content: `❌ This tournament requires at least **${requiredMentions}** player IGN(s) — you gave ${players.length}.`,
      flags: MessageFlags.Ephemeral,
    });
  }

  const letter = Object.keys(tournament.groups)
    .sort()
    .find(l => tournament.groups[l].teams.length < tournament.groups[l].capacity);

  if (!letter) {
    return interaction.reply({ content: '❌ Every group is full right now — ask an admin to open more groups.', flags: MessageFlags.Ephemeral });
  }

  tournament.groups[letter].teams.push({ team, players });
  saveGuildStore(interaction.guildId, store);

  const slotNumber = tournament.groups[letter].teams.length;

  if (tournament.confirmChannelId) {
    const confirmChannel = interaction.guild.channels.cache.get(tournament.confirmChannelId);
    if (confirmChannel) {
      await confirmChannel.send(`✅ **${team}** (${interaction.user}) registered into **Group ${letter}**, Slot **${slotNumber}**.`).catch(() => {});
    }
  }

  if (tournament.successRoleId) {
    await interaction.member.roles.add(tournament.successRoleId).catch(() => {});
  }

  return interaction.reply({
    content: `✅ **${team}** registered into **Group ${letter}**, Slot **${slotNumber}**.`,
    flags: MessageFlags.Ephemeral,
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
  if (!Number.isInteger(value) || value < 1 || value > 64) {
    return interaction.reply({ content: '❌ Teams per Group must be a whole number between 1 and 64.', flags: MessageFlags.Ephemeral });
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
  if (!Number.isInteger(value) || value < 1 || value > 768) {
    return interaction.reply({ content: '❌ Total Slots must be a whole number between 1 and 768.', flags: MessageFlags.Ephemeral });
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

async function handleCreateRoleSelect(interaction) {
  if (!hasManageGuild(interaction)) {
    return interaction.reply({ content: '❌ You need the **Manage Server** permission to do that.', flags: MessageFlags.Ephemeral });
  }
  const store = getGuildStore(interaction.guildId);
  if (!store.tournament) {
    return interaction.update({ content: '❌ No tournament exists yet.', components: [] });
  }
  store.tournament.successRoleId = interaction.roles.first().id;
  saveGuildStore(interaction.guildId, store);
  await interaction.update({ content: '', ...buildCreateSettingsPayload(store.tournament) });
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

  const group = tournament.groups[letter];
  const selectedTeams = interaction.values.map(v => group.teams[parseInt(v, 10)]?.team).filter(Boolean);

  // Re-running qualify on the same group cleanly replaces its previous
  // picks rather than piling up duplicates: drop every team from this
  // group out of the qualified list first, then add back only what's
  // selected now.
  const groupTeamNames = new Set(group.teams.map(t => t.team));
  tournament.qualified = tournament.qualified.filter(name => !groupTeamNames.has(name));
  tournament.qualified.push(...selectedTeams);
  saveGuildStore(interaction.guildId, store);

  await interaction.update({
    content: `✅ Group **${letter}** qualifiers updated: ${selectedTeams.length ? selectedTeams.map(t => `**${t}**`).join(', ') : '_none selected_'}`,
    embeds: [],
    components: [],
  });
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
  const removedNames = group.teams.filter((t, idx) => removeIndices.has(idx)).map(t => t.team);
  group.teams = group.teams.filter((t, idx) => !removeIndices.has(idx));

  const removedSet = new Set(removedNames);
  tournament.qualified = tournament.qualified.filter(name => !removedSet.has(name));
  saveGuildStore(interaction.guildId, store);

  await interaction.update({
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

  await interaction.update({ content: `✅ Slot-Manager channel set to ${channel}.`, components: [] });
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
  handleCreateRoleSelect,
};
