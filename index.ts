// Import required packages
import * as restify from "restify";
import * as fs from 'fs';
import { promisify } from 'util';

const readFileAsync = promisify(fs.readFile);

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  MemoryStorage,
  CardFactory,
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

  if (error.message.includes("Cannot read properties of undefined (reading '0')")) {
    await context.sendActivity("Your request was throttled, please try again.")
  }
  else {
    // Send a message to the user
    await context.sendActivity(`The bot encountered unhandled error (fixed):\n ${error}`);
    await context.sendActivity(`${error.stack}`)
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
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

import { Application, ConversationHistory, DefaultPromptManager, DefaultTurnState, AI, AzureOpenAIPlanner, DefaultUserState, DefaultTempState } from '@microsoft/teams-ai';
import path from "path";

// eslint-disable-next-line @typescript-eslint/no-empty-interface
interface ConversationState {
  riskyUser: string;
  subscribedUser: string;
  alertsList: AlertsData[];
  insightsList: {[key: string]: any};
  subscribers: string[];
}
type UserState = DefaultUserState;

interface TempState extends DefaultTempState {
  alertsList: AlertsData[];
}

interface AlertsData {
  UserPrincipalName: string;
  SequentialActivities: {[key: string]: any}[];
  ComparativeActivities: {[key: string]: any}[];
}

type ApplicationTurnState = DefaultTurnState<ConversationState, UserState, TempState>;

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
  subscribedUser: string;
  summary: string;
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
  state.conversation.value.riskyUser = data.riskyUser
  await readJsonFile(data.riskyUser)
    .then((jsonData) => {
      state.conversation.value.alertsList = jsonData.length > 0 ? jsonData : []
      console.log("Alerts found: " + JSON.stringify(state.conversation.value.alertsList))
    })
    .catch((error) => {
      console.error(error);
    });
  if (state.conversation.value.alertsList.length > 0) {
    state.temp.value.alertsList = state.conversation.value.alertsList
    await app.ai.chain(context, state, 'summarize');
  }
  else {
    await context.sendActivity("No alerts found for user: " + data.riskyUser);
  }
  return false;
});

app.ai.action("SummarizeAlert", async (context, state, data: EntityData) => {
  state.conversation.value.riskyUser = data.riskyUser
  await readJsonFile(data.riskyUser)
    .then((jsonData) => {
      state.conversation.value.alertsList = jsonData.length > 0 ? jsonData : []
      console.log("Alerts found:" + JSON.stringify(state.conversation.value.alertsList))
      //console.log("Alerts found")
    })
    .catch((error) => {
      console.error(error);
    });
  if (state.conversation.value.alertsList.length > 0) {
    state.temp.value.alertsList = state.conversation.value.alertsList
    await app.ai.chain(context, state, 'summarize');
  }
  else {
    await context.sendActivity("No alerts found for user: " + data.riskyUser);
  }
  return false
});

app.ai.action("SetupUserReminder", async (context, state, data: EntityData) => {
  if (!state.conversation.value.subscribers) {
    state.conversation.value.subscribers = [];
  }
  state.conversation.value.subscribers.push(data.subscribedUser);
  await context.sendActivity("Subscribing to alerts for user: " + data.subscribedUser);
  return true;
});

app.ai.action("RemoveUserReminder", async (context, state, data: EntityData) => {
  const index = state.conversation.value.subscribers.indexOf(data.subscribedUser);
  if (index !== -1) {
    // Element found in the subscribers array, remove it
    state.conversation.value.subscribers.splice(index, 1);
    await context.sendActivity("Removed user reminder for: " + data.subscribedUser);
  } else {
    await context.sendActivity("User reminder not found for: " + data.subscribedUser);
  }
  return true;
});

app.ai.action("DisplayReminderUserList", async (context, state, data: EntityData) => {
  await context.sendActivity("You are currently subscribed to reminders for the following users: " + state.conversation.value.subscribers);
  return true;
});

app.ai.action("displayAdaptiveCardWithSummary", async (context, state, data: EntityData) => {
  const card = {
    "type": "AdaptiveCard",
    "body": [
        {
            "type": "TextBlock",
            "size": "Medium",
            "weight": "Bolder",
            "text": `Summary of alerts for user: ${state.conversation.value.riskyUser}`
        },
        {
            "type": "TextBlock",
            "text": data.summary,
            "wrap": true
        }
    ],
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.5"
  }
  await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
  return true;
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

async function readJsonFile(riskyUser: string): Promise<AlertsData[]> {
  const filePath = 'data/alerts.json';

  try {
    const fileContent = await readFileAsync(filePath, 'utf8');
    const jsonData = JSON.parse(fileContent).filter((item) => {
      return item.UserPrincipalName.toLowerCase().includes(riskyUser.toLowerCase());
    });
  
    return jsonData;
  } catch (error) {
    throw new Error(`Error reading JSON file: ${error}`);
  }
}