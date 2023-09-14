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

// Create the bot that will handle incoming messages.
const conversationReferences = {};

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
  ai: 
  {
    planner,
    //moderator,
    promptManager,
    prompt: 'chat',
    history: {
        assistantHistoryType: 'text'
    }
  },
  botAppId: process.env.BOT_ID,
  adapter: adapter
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

let cachedAlertData: AlertsData[] = [];
let loadedAlertsFromDb: boolean = false;
app.ai.action("RetrieveAlerts", async (context, state, data: EntityData) => {
  state.conversation.value.riskyUser = data.riskyUser
  if (!loadedAlertsFromDb) {
    await readJsonFile(data.riskyUser)
      .then((jsonData) => {
        loadedAlertsFromDb = true;
        state.conversation.value.alertsList = jsonData.length > 0 ? jsonData : []
        cachedAlertData.push(...jsonData);
        console.log("Alerts found: " + JSON.stringify(state.conversation.value.alertsList))
      })
      .catch((error) => {
        console.error(error);
      });
  }
  
  state.conversation.value.alertsList = cachedAlertData.filter((item) => {
    return item.UserPrincipalName.toLowerCase().includes(data.riskyUser.toLowerCase());
  });
  
  if (state.conversation.value.alertsList.length > 0) {
    console.log("Alerts found for " + data.riskyUser + " : " + JSON.stringify(state.conversation.value.alertsList))
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
  if (!loadedAlertsFromDb) {
    await readJsonFile(data.riskyUser)
    .then((jsonData) => {
      loadedAlertsFromDb = true;
      state.conversation.value.alertsList = jsonData.length > 0 ? jsonData : []
      cachedAlertData.push(...jsonData);
      console.log("Alerts found:" + JSON.stringify(state.conversation.value.alertsList))
    })
    .catch((error) => {
      console.error(error);
    });
  }

  state.conversation.value.alertsList = cachedAlertData.filter((item) => {
    return item.UserPrincipalName.toLowerCase().includes(data.riskyUser.toLowerCase());
  });
  
  if (state.conversation.value.alertsList.length > 0) {
    console.log("Alerts found for " + data.riskyUser + " : " + JSON.stringify(state.conversation.value.alertsList))
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
    conversationReferences[context.activity.conversation.id] = TurnContext.getConversationReference(context.activity);
    await app.run(context);
  });
});

// Listen for incoming notifications and send proactive messages to users.
server.post('/api/notify', async (req, res) => {
  console.log(JSON.stringify(conversationReferences));
  for (const conversationReference of Object.values(conversationReferences)) {
    await app.continueConversationAsync(conversationReference, async (context) => {
      const userName:string = req.body.key[0].UserPrincipalName;
      const card = {
        "type": "AdaptiveCard",
        "body": [
          {
            "type": "Image",
            "url" : "data:image/png;base64,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",
            "size": "medium", // Y
          },
          {
                "type": "TextBlock",
                "size": "Medium",
                "weight": "Bolder",
                "text": `New alert reported for : ${userName}`
            }
          ],
        "actions": [
              {
                  type: 'Action.Submit',
                  title: 'Summarize alert',
                  data: {
                    "msteams": {
                      "type": "messageBack",
                      "displayText": `Summarize Alert for : ${userName}`,
                      "text": `Summarize Alert for : ${userName}`,
                  },
                }
              },
          ],
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5"
      }

      //await context.sendActivity("Received new alert for user: " + req.body.key[0].UserPrincipalName);
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
      cachedAlertData.push(...req.body.key);
    });
  }

  res.setHeader('Content-Type', 'text/html');
  res.writeHead(200);
  res.write('<html><body><h1>Proactive messages have been sent.</h1></body></html>');
  res.end();
});

async function readJsonFile(riskyUser: string): Promise<AlertsData[]> {
  const filePath = 'data/alerts.json';

  try {
    const fileContent = await readFileAsync(filePath, 'utf8');
    const jsonData = JSON.parse(fileContent);
  
    return jsonData;
  } catch (error) {
    throw new Error(`Error reading JSON file: ${error}`);
  }
}