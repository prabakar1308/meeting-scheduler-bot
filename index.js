require("dotenv").config();
const express = require("express");
const crypto = require("crypto");
const axios = require("axios");
const {
  CloudAdapter,
  MemoryStorage,
  ConversationState,
  UserState,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  CardFactory,
} = require("botbuilder");
const { Client } = require("@microsoft/microsoft-graph-client");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3978;

// OAuth Configuration
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;
const TENANT_ID = process.env.TENANT_ID || "common"; // Use 'common' for multi-tenant
const REDIRECT_URI =
  process.env.OAUTH_REDIRECT_URI ||
  `https://meeting-scheduler-bot-i5oj.onrender.com/api/oauth/callback`;

const SCOPES = "User.Read Calendars.ReadWrite offline_access";

console.log("🔧 OAuth Configuration:");
console.log("   App ID:", MICROSOFT_APP_ID ? "✓ Set" : "✗ Missing");
console.log("   App Password:", MICROSOFT_APP_PASSWORD ? "✓ Set" : "✗ Missing");
console.log("   Tenant ID:", TENANT_ID);
console.log("   Redirect URI:", REDIRECT_URI);

// Bot Framework setup
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: MICROSOFT_APP_ID,
  MicrosoftAppPassword: MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error(`\n❌ [onTurnError] Error: ${error.message}`);
  console.error(error.stack);

  await context.sendActivity(
    "⚠️ Sorry, something went wrong. Please try again."
  );
  await conversationState.delete(context);
};

// State management
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// In-memory storage for OAuth states and tokens
// In production, use Redis or a database
const oauthStateStore = new Map();
const tokenStore = new Map();

// Clean up old states (older than 10 minutes)
setInterval(() => {
  const tenMinutesAgo = Date.now() - 10 * 60 * 1000;
  for (const [state, data] of oauthStateStore.entries()) {
    if (data.timestamp < tenMinutesAgo) {
      oauthStateStore.delete(state);
    }
  }
}, 5 * 60 * 1000);

// Helper function to generate OAuth URL
function generateOAuthUrl(userId, conversationId) {
  const state = crypto.randomBytes(16).toString("hex");

  // Store state with user info
  oauthStateStore.set(state, {
    userId,
    conversationId,
    timestamp: Date.now(),
  });

  const authUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    `client_id=${MICROSOFT_APP_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&response_mode=query` +
    `&scope=${encodeURIComponent(SCOPES)}` +
    `&state=${state}`;

  console.log(`🔗 Generated OAuth URL for user: ${userId}`);
  return authUrl;
}

// Helper function to exchange code for token
async function exchangeCodeForToken(code) {
  try {
    const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append("client_id", MICROSOFT_APP_ID);
    params.append("scope", SCOPES);
    params.append("code", code);
    params.append("redirect_uri", REDIRECT_URI);
    params.append("grant_type", "authorization_code");
    params.append("client_secret", MICROSOFT_APP_PASSWORD);

    const response = await axios.post(tokenUrl, params, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });

    return response.data;
  } catch (error) {
    console.error(
      "❌ Token exchange error:",
      error.response?.data || error.message
    );
    throw error;
  }
}

// Helper function to get Graph client
function getGraphClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}

// Bot class
class MeetingSchedulerBot {
  constructor(conversationState, userState) {
    this.conversationState = conversationState;
    this.userState = userState;
    this.conversationData =
      conversationState.createProperty("conversationData");
    this.userData = userState.createProperty("userData");
  }

  async run(context) {
    await this.onTurn(context);
    await this.conversationState.saveChanges(context, false);
    await this.userState.saveChanges(context, false);
  }

  async onTurn(context) {
    const activityType = context.activity.type;

    if (activityType === "message") {
      await this.handleMessage(context);
    } else if (activityType === "conversationUpdate") {
      await this.handleConversationUpdate(context);
    }
  }

  async handleMessage(context) {
    const text = (context.activity.text || "").toLowerCase().trim();
    const userId = context.activity.from.id;
    const conversationId = context.activity.conversation.id;

    console.log(
      `\n📩 Message from ${context.activity.from.name}: "${context.activity.text}"`
    );

    // Check if user has a valid token
    const userToken = tokenStore.get(userId);
    const hasToken = userToken && userToken.access_token;

    console.log(`   Has token: ${hasToken}`);

    if (!hasToken) {
      // User needs to authenticate
      if (
        text.includes("login") ||
        text.includes("signin") ||
        text === "hi" ||
        text === "hello"
      ) {
        await this.sendSignInCard(context, userId, conversationId);
      } else {
        await context.sendActivity(
          "👋 Welcome! Please sign in to use this bot.\n\n" +
            "Type 'login' to get started."
        );
      }
      return;
    }

    // User is authenticated - handle commands
    console.log("✅ User is authenticated");

    if (text.includes("calendar") || text.includes("events")) {
      await this.listCalendarEvents(context, userId);
    } else if (text.includes("schedule") || text.includes("create")) {
      await context.sendActivity(
        "📅 Let's schedule a meeting! (Feature coming soon)"
      );
    } else if (text.includes("logout") || text.includes("signout")) {
      tokenStore.delete(userId);
      await context.sendActivity("✅ You've been signed out successfully.");
    } else if (text.includes("help")) {
      await this.sendHelpMessage(context);
    } else {
      await context.sendActivity(
        "I'm here to help! Try:\n\n" +
          "• `calendar` - View your calendar events\n" +
          "• `schedule` - Create a new meeting\n" +
          "• `logout` - Sign out\n" +
          "• `help` - Show help"
      );
    }
  }

  async handleConversationUpdate(context) {
    if (context.activity.membersAdded) {
      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            "👋 **Welcome to Meeting Scheduler Bot!**\n\n" +
              "I can help you manage your calendar and schedule meetings.\n\n" +
              "Type **login** to get started."
          );
        }
      }
    }
  }

  async sendSignInCard(context, userId, conversationId) {
    const authUrl = generateOAuthUrl(userId, conversationId);

    const signInCard = CardFactory.signinCard(
      "Sign in to Microsoft",
      authUrl,
      "Please sign in to continue"
    );

    await context.sendActivity({ attachments: [signInCard] });

    console.log("🔐 Sign-in card sent");
  }

  async sendHelpMessage(context) {
    await context.sendActivity(
      "**📚 Available Commands:**\n\n" +
        "**Calendar Management:**\n" +
        "• `calendar` or `events` - View your upcoming calendar events\n" +
        "• `schedule` or `create` - Schedule a new meeting\n\n" +
        "**Account:**\n" +
        "• `logout` - Sign out from your account\n" +
        "• `help` - Show this help message"
    );
  }

  async listCalendarEvents(context, userId) {
    try {
      const userToken = tokenStore.get(userId);
      const client = getGraphClient(userToken.access_token);

      await context.sendActivity("🔍 Fetching your calendar events...");

      const events = await client
        .api("/me/calendar/events")
        .top(10)
        .select("subject,start,end,organizer")
        .orderby("start/dateTime")
        .get();

      if (events.value.length === 0) {
        await context.sendActivity(
          "📅 You have no upcoming events in your calendar."
        );
        return;
      }

      let message = "📅 **Your Upcoming Events:**\n\n";
      events.value.forEach((event, index) => {
        const startDate = new Date(event.start.dateTime);
        const endDate = new Date(event.end.dateTime);

        message += `**${index + 1}. ${event.subject}**\n`;
        message += `   📅 ${startDate.toLocaleDateString()} ${startDate.toLocaleTimeString(
          [],
          { hour: "2-digit", minute: "2-digit" }
        )}\n`;
        message += `   ⏱️ Duration: ${Math.round(
          (endDate - startDate) / 60000
        )} minutes\n`;
        if (event.organizer?.emailAddress?.name) {
          message += `   👤 Organizer: ${event.organizer.emailAddress.name}\n`;
        }
        message += "\n";
      });

      await context.sendActivity(message);
      console.log("✅ Calendar events sent");
    } catch (error) {
      console.error("❌ Error fetching calendar:", error);

      if (error.statusCode === 401) {
        tokenStore.delete(userId);
        await context.sendActivity(
          "⚠️ Your session has expired. Please type 'login' to sign in again."
        );
      } else {
        await context.sendActivity(
          "❌ Sorry, I couldn't fetch your calendar events. Please try again later."
        );
      }
    }
  }
}

const bot = new MeetingSchedulerBot(conversationState, userState);

// Bot messaging endpoint
app.post("/api/messages", async (req, res) => {
  console.log("\n" + "=".repeat(60));
  console.log("📨 Incoming message request");

  try {
    await adapter.process(req, res, async (context) => {
      await bot.run(context);
    });
  } catch (error) {
    console.error("❌ Error in /api/messages:", error);
    if (!res.headersSent) {
      res.status(500).json({ error: error.message });
    }
  }
});

// OAuth callback endpoint
app.get("/api/oauth/callback", async (req, res) => {
  console.log("\n🔐 OAuth callback received");

  const { code, state, error } = req.query;

  if (error) {
    console.error("❌ OAuth error:", error);
    return res.send(`
      <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>❌ Authentication Failed</h1>
          <p>Error: ${error}</p>
          <p>Please close this window and try again in Teams.</p>
        </body>
      </html>
    `);
  }

  // Verify state
  const stateData = oauthStateStore.get(state);
  if (!stateData) {
    console.error("❌ Invalid or expired state");
    return res.send(`
      <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>❌ Authentication Failed</h1>
          <p>Invalid or expired session. Please try again.</p>
        </body>
      </html>
    `);
  }

  try {
    // Exchange code for token
    console.log("🔄 Exchanging code for token...");
    const tokenData = await exchangeCodeForToken(code);

    // Store token for user
    tokenStore.set(stateData.userId, {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: Date.now() + tokenData.expires_in * 1000,
    });

    console.log("✅ Token stored for user:", stateData.userId);

    // Clean up state
    oauthStateStore.delete(state);

    // Send success page
    res.send(`
      <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>✅ Sign-in Successful!</h1>
          <p>You can now close this window and return to Teams.</p>
          <p>Try typing <strong>"calendar"</strong> to see your events!</p>
          <script>
            setTimeout(() => window.close(), 3000);
          </script>
        </body>
      </html>
    `);
  } catch (error) {
    console.error("❌ Token exchange failed:", error);
    res.send(`
      <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>❌ Authentication Failed</h1>
          <p>Could not complete sign-in. Please try again.</p>
        </body>
      </html>
    `);
  }
});

// Health check
app.get("/", (req, res) => {
  res.json({
    status: "✅ Bot is running with OAuth",
    timestamp: new Date().toISOString(),
  });
});

app.get("/health", (req, res) => {
  res.json({
    status: "healthy",
    uptime: process.uptime(),
  });
});

// Start server
app.listen(PORT, () => {
  console.log("\n" + "=".repeat(60));
  console.log("🤖 Meeting Scheduler Bot with OAuth Started");
  console.log("=".repeat(60));
  console.log(`📍 Port: ${PORT}`);
  console.log(
    `💬 Messages: https://meeting-scheduler-bot-i5oj.onrender.com/api/messages`
  );
  console.log(`🔐 OAuth Callback: ${REDIRECT_URI}`);
  console.log("=".repeat(60) + "\n");
});
