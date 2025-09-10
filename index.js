// Minimal Microsoft Teams Bot (Node.js + Express)
const express = require('express');
const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration
} = require('botbuilder');
require('dotenv').config();

// Echo bot with 'help' and 'card' commands
class EchoBot extends ActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      const text = (context.activity.text || '').trim().toLowerCase();
      if (text === 'help') {
        await context.sendActivity('Try typing anything, or "card" to see an Adaptive Card.');
      } else if (text === 'card') {
        await context.sendActivity({
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: {
              '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
              'type': 'AdaptiveCard',
              'version': '1.4',
              'body': [
                { 'type': 'TextBlock', 'text': 'Hello from Adaptive Card!', 'weight': 'Bolder', 'size': 'Medium' },
                { 'type': 'TextBlock', 'text': 'This card was sent by your Azure-hosted bot.' }
              ]
            }
          }]
        });
      } else {
        await context.sendActivity(`You said: ${context.activity.text}`);
      }
      await next();
    });
    this.onMembersAdded(async (context, next) => {
      for (const m of context.activity.membersAdded ?? []) {
        if (m.id !== context.activity.recipient.id) {
          await context.sendActivity('Hello! I am your Teams bot. Type "help" to see options.');
        }
      }
      await next();
    });
  }
}

// Bot Framework auth
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppType: process.env.MicrosoftAppType || 'SingleTenant',
  MicrosoftTenantId: process.env.MicrosoftTenantId
});
const bfa = createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);
const adapter = new CloudAdapter(bfa);
const bot = new EchoBot();

// Express server
const app = express();
const port = process.env.PORT || 3978;
app.use(express.json());

// Health check
app.get('/', (req, res) => res.status(200).send('OK'));

// REQUIRED endpoint for Teams/Bot Framework
app.post('/api/messages', (req, res) => {
  adapter.process(req, res, (context) => bot.run(context));
});

app.listen(port, () => console.log(`Bot listening on port ${port}`));
