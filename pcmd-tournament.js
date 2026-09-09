const { getGuildStore } = require('./storage');
const { buildTournamentWizardPayload } = require('./tournament-wizard-handlers');

module.exports = {
  name: 'tournament',
  aliases: ['tourney'],
  description: 'Post the tournament setup panel — create a tournament, add/auto-create groups, register teams, view slot lists (usage: !tournament)',
  adminOnly: true,

  async execute(message) {
    const store = getGuildStore(message.guildId);
    const payload = buildTournamentWizardPayload(store);
    await message.channel.send(payload);
  },
};
