const { SlashCommandBuilder, PermissionFlagsBits, MessageFlags } = require('discord.js');
const { getGuildStore, saveGuildStore } = require('./storage');

// Accepts a bare handle or a full profile URL/@-mention and normalizes
// down to just the username, since people will copy-paste either.
function extractUsername(raw) {
  let value = raw.trim();
  const urlMatch = value.match(/instagram\.com\/([A-Za-z0-9._]+)/i);
  if (urlMatch) value = urlMatch[1];
  value = value.replace(/^@/, '');
  return value;
}

module.exports = {
  data: new SlashCommandBuilder()
    .setName('set-instagram-accounts')
    .setDescription('Set the 4 Instagram accounts members must all follow for screenshot verification')
    .addStringOption(opt =>
      opt.setName('account1').setDescription('First Instagram handle (e.g. reboundesports)').setRequired(true))
    .addStringOption(opt =>
      opt.setName('account2').setDescription('Second Instagram handle').setRequired(true))
    .addStringOption(opt =>
      opt.setName('account3').setDescription('Third Instagram handle').setRequired(true))
    .addStringOption(opt =>
      opt.setName('account4').setDescription('Fourth Instagram handle').setRequired(true))
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  async execute(interaction) {
    const usernames = [1, 2, 3, 4].map(n => extractUsername(interaction.options.getString(`account${n}`)));

    const invalid = usernames.filter(u => !/^[A-Za-z0-9._]{1,30}$/.test(u));
    if (invalid.length) {
      return interaction.reply({
        content: `❌ These don't look like valid Instagram usernames: ${invalid.join(', ')}. Use the handle itself (e.g. \`reboundesports\`) or a full profile link for each.`,
        flags: MessageFlags.Ephemeral,
      });
    }

    const distinct = new Set(usernames.map(u => u.toLowerCase()));
    if (distinct.size !== 4) {
      return interaction.reply({
        content: '❌ All 4 accounts need to be different — you entered a duplicate.',
        flags: MessageFlags.Ephemeral,
      });
    }

    const store = getGuildStore(interaction.guildId);
    if (!store.settings) store.settings = {};
    store.settings.instagramUsernames = usernames;
    saveGuildStore(interaction.guildId, store);

    await interaction.reply({
      content: `✅ Screenshot verification now requires following all 4: ${usernames.map(u => `**@${u}**`).join(', ')}.`,
      flags: MessageFlags.Ephemeral,
    });
  },
};
