// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required packages
import * as path from 'path';
import * as restify from 'restify';

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory,
    MemoryStorage,
    TurnContext
} from 'botbuilder';

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.BOT_ID,
        MicrosoftAppPassword: process.env.BOT_PASSWORD,
        MicrosoftAppType: 'MultiTenant'
    })
);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo test your bot in Teams, sideload the app manifest.json within Teams Apps.');
});

import {
    Application,
    DefaultTurnState,
    DefaultConversationState,
    DefaultUserState,
    DefaultTempState,
    DefaultPromptManager,
    AzureOpenAIPlanner,
    ConversationHistory
} from '@microsoft/teams-ai';

type ApplicationTurnState = DefaultTurnState<DefaultConversationState, DefaultUserState, DefaultTempState>;

const planner = new AzureOpenAIPlanner({
    apiKey: process.env.OPENAI_API_KEY,
    endpoint: process.env.OPENAI_ENDPOINT,
    defaultModel: 'davinciThreeTeamsBot', //this needs to be the name of the openai deployment you have, not the model name
    logRequests: true,
})

const promptManager = new DefaultPromptManager<ApplicationTurnState>(path.join(__dirname, '../src/prompts'));

// Define storage and application
const storage = new MemoryStorage();

// Define the application instance
const app = new Application<ApplicationTurnState>({
    storage,
    ai: {
        planner,
        promptManager,
        prompt: 'chat', // refers to the `src/prompts/chat` folder
        history: {
            assistantHistoryType: 'text',
            maxTurns: 10, 
            maxTokens: 1000
        }
    }
});

// Send '/clearhistory' to the bot in Teams to clear the current converstion history with the LLM AI model
app.message('/clearhistory', async (context: TurnContext, state: ApplicationTurnState) => {
    ConversationHistory.clear(state);
    await context.sendActivity('The AI conversation history is cleared')
})

// Send '/history' to the bot in Teams to view the current conversation history with the LLM AI model
app.message('/history', async (context: TurnContext, state: ApplicationTurnState) => {
    const history = ConversationHistory.toString(state, 2000, '\n\n');
    await context.sendActivity(history);
});

// Listen for incoming server requests.
server.post('/api/messages', async (req, res) => {
    // Route received a request to adapter for processing
    await adapter.process(req, res as any, async (context) => {
        // Dispatch to application for routing
        await app.run(context);
    });
});
