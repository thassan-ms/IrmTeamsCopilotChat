// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  MemoryStorage,
} from "botbuilder";

// This bot's main dialog.
//import { TeamsBot } from "./teamsBot";
import config from "./config";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(`The bot encountered unhandled error (fixed):\n ${error}`);
  await context.sendActivity(`${error.stack}`)
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
//const bot = new TeamsBot();

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

import { Application, ConversationHistory, DefaultPromptManager, DefaultTurnState, AI, AzureOpenAIPlanner } from '@microsoft/teams-ai';
import path from "path";

// eslint-disable-next-line @typescript-eslint/no-empty-interface
interface ConversationState {
  riskyUser: string;
  alertsList: {[key: string]: any};
  insightsList: {[key: string]: any};
}
type ApplicationTurnState = DefaultTurnState<ConversationState>;

// Create AI components
const apiKey = config.openAIKey
const openAIEndpoint = config.openAIEndpoint
const model2 = 'gpt-35-turbo-16k'

const planner = new AzureOpenAIPlanner({
  endpoint: openAIEndpoint,
  apiKey: apiKey,
  defaultModel: model2,
  logRequests: true,
  apiVersion: "2023-07-01-preview"
})
// const moderator = new OpenAIModerator({
//   apiKey: apiKey,
//   endpoint: openAIEndpoint,
//   model: model2,
//   moderate: 'both'
// });
const promptManager = new DefaultPromptManager<ApplicationTurnState>(path.join(__dirname, './prompts' ));

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
  storage,
  ai: {
      planner,
      //moderator,
      promptManager,
      prompt: 'chat',
      history: {
          assistantHistoryType: 'text'
      }
  }
});

interface EntityData {
  riskyUser: string; // <- populated by GPT
}

const getAlertsForUser = (user: string) => {
  const alertsData = {
    "diego": {
      "alertId": "1",
      "severity": "high",
      "riskScore": "100",
    },
    "tasmiha":{
      "alertId": "2",
      "severity": "medium",
      "riskScore": "50",
    },
    "moise": {
      "alertId": "3",
      "severity": "low",
      "riskScore": "10",
    }
  }

  return alertsData[user] ?? "{Unknown}"
}

app.ai.action(AI.RateLimitedActionName, async (context, state, data) => {
  await context.sendActivity(`Your request was rate limited: ${JSON.stringify(data)}`);
  return false;
});

app.ai.action(AI.FlaggedInputActionName, async (context, state, data) => {
  await context.sendActivity(`I'm sorry your message was flagged: ${JSON.stringify(data)}`);
  return false;
});

app.ai.action(AI.FlaggedOutputActionName, async (context, state, data) => {
  await context.sendActivity(`I'm not allowed to talk about such things.`);
  return false;
});

app.ai.action("RetrieveAlerts", async (context, state, data: EntityData) => {
  await context.sendActivity("Retrieving alerts for user: " + data.riskyUser);
  state.conversation.value.riskyUser = data.riskyUser
  state.conversation.value.alertsList = getAlertsForUser(data.riskyUser)
  return false;
});

app.ai.action("SummarizeAlert", async (context, state, data: EntityData) => {
  await context.sendActivity("Summarizing alert: " + data.riskyUser);
  state.conversation.value.riskyUser = data.riskyUser
  state.conversation.value.alertsList = getAlertsForUser(data.riskyUser)
  await app.ai.chain(context, state, 'summarize');
  return false
});

app.message('/history', async (context, state) => {
  const history = ConversationHistory.toString(state, 2000, '\n\n');
  await context.sendActivity(history);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await app.run(context);
  });
});
