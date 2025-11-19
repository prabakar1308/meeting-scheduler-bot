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
  MessageFactory,
  ActionTypes,
} = require("botbuilder");
const { Client } = require("@microsoft/microsoft-graph-client");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3978;

// ============================================================================
// OAUTH CONFIGURATION
// ============================================================================

const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;
const TENANT_ID = process.env.TENANT_ID || "common";

const BASE_URL =
  process.env.BASE_URL || "https://meeting-scheduler-bot-i5oj.onrender.com";
const REDIRECT_URI = `${BASE_URL}/api/oauth/callback`;

const SCOPES = "User.Read Calendars.ReadWrite offline_access";

console.log("\n🔧 OAuth Configuration:");
console.log("   App ID:", MICROSOFT_APP_ID ? "✓ Configured" : "✗ MISSING");
console.log(
  "   App Password:",
  MICROSOFT_APP_PASSWORD ? "✓ Configured" : "✗ MISSING"
);
console.log("   Tenant ID:", TENANT_ID);
console.log("   Redirect URI:", REDIRECT_URI);

// ============================================================================
// BOT FRAMEWORK SETUP
// ============================================================================

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: MICROSOFT_APP_ID,
  MicrosoftAppPassword: MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error(`\n❌ [Bot Error]: ${error.message}`);
  console.error(error.stack);

  try {
    await context.sendActivity(
      "⚠️ Sorry, something went wrong. Please try again."
    );
  } catch (err) {
    console.error("Failed to send error message:", err);
  }
};

// ============================================================================
// STATE MANAGEMENT
// ============================================================================

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// ============================================================================
// IN-MEMORY STORAGE
// ============================================================================

const oauthStateStore = new Map();
const tokenStore = new Map();

// Clean up expired states
setInterval(() => {
  const tenMinutesAgo = Date.now() - 10 * 60 * 1000;
  for (const [state, data] of oauthStateStore.entries()) {
    if (data.timestamp < tenMinutesAgo) {
      oauthStateStore.delete(state);
    }
  }
}, 5 * 60 * 1000);

// ============================================================================
// OAUTH HELPER FUNCTIONS
// ============================================================================

function generateOAuthUrl(userId) {
  const state = crypto.randomBytes(16).toString("hex");

  oauthStateStore.set(state, {
    userId,
    timestamp: Date.now(),
  });

  const authUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    `client_id=${encodeURIComponent(MICROSOFT_APP_ID)}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&response_mode=query` +
    `&scope=${encodeURIComponent(SCOPES)}` +
    `&state=${state}` +
    `&prompt=select_account`;

  console.log(`🔗 Generated OAuth URL for user: ${userId}`);
  return authUrl;
}

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

    console.log("🔄 Exchanging authorization code for token...");

    const response = await axios.post(tokenUrl, params, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });

    console.log("✅ Token exchange successful");
    return response.data;
  } catch (error) {
    console.error("❌ Token exchange error:");
    console.error("   Status:", error.response?.status);
    console.error("   Data:", error.response?.data);
    throw error;
  }
}

async function getValidToken(userId) {
  const tokenData = tokenStore.get(userId);

  if (!tokenData) {
    return null;
  }

  const now = Date.now();
  const expiresAt = tokenData.expires_at || 0;

  if (now >= expiresAt - 5 * 60 * 1000) {
    // Token expired
    tokenStore.delete(userId);
    return null;
  }

  return tokenData.access_token;
}

function getGraphClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}

// ============================================================================
// BOT CLASS
// ============================================================================

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
    const userName = context.activity.from.name || "User";

    console.log(`\n📩 Message from ${userName}: "${context.activity.text}"`);

    // Check authentication
    const accessToken = await getValidToken(userId);
    const isAuthenticated = !!accessToken;

    console.log(`   Authentication: ${isAuthenticated ? "✅ Yes" : "❌ No"}`);

    // Handle unauthenticated users
    if (!isAuthenticated) {
      if (
        text.includes("login") ||
        text.includes("signin") ||
        text.includes("start") ||
        text.includes("hello") ||
        text.includes("hi")
      ) {
        await this.sendSignInMessage(context, userId);
      } else {
        await context.sendActivity(
          "👋 **Welcome to Meeting Scheduler Bot!**\n\n" +
            "To get started, please sign in to your Microsoft account.\n\n" +
            "Type **login** to continue."
        );
      }
      return;
    }

    // Handle authenticated commands
    if (
      text.includes("calendar") ||
      text.includes("events") ||
      text.includes("list")
    ) {
      await this.listCalendarEvents(context, userId, accessToken);
    } else if (
      text.includes("profile") ||
      text.includes("me") ||
      text.includes("whoami")
    ) {
      await this.showUserProfile(context, userId, accessToken);
    } else if (text.includes("logout") || text.includes("signout")) {
      await this.logout(context, userId);
    } else if (text.includes("help")) {
      await this.sendHelpMessage(context, true);
    } else {
      await context.sendActivity(
        `You said: "${context.activity.text}"\n\n` +
          "**Available commands:**\n" +
          "• `calendar` - View your calendar\n" +
          "• `profile` - Show your profile\n" +
          "• `logout` - Sign out\n" +
          "• `help` - Show help"
      );
    }

    console.log("   ✅ Response sent");
  }

  async handleConversationUpdate(context) {
    if (context.activity.membersAdded) {
      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          console.log(`   👤 New member: ${member.name || member.id}`);

          await context.sendActivity(
            "👋 **Welcome to Meeting Scheduler Bot!**\n\n" +
              "I can help you manage your Microsoft Calendar.\n\n" +
              "Type **login** to get started!"
          );
        }
      }
    }
  }

  async sendSignInMessage(context, userId) {
    const authUrl = generateOAuthUrl(userId);

    // Create a simple message with a link instead of a signin card
    const message = MessageFactory.text(
      `🔐 **Please sign in to continue**\n\n` +
        `Click here to sign in: [Sign in to Microsoft](${authUrl})\n\n` +
        `After signing in, close the browser window and return here.`
    );

    await context.sendActivity(message);
    console.log("   🔐 Sign-in link sent");
  }

  async sendHelpMessage(context, isAuthenticated) {
    let helpText = "**📚 Meeting Scheduler Bot - Help**\n\n";

    if (isAuthenticated) {
      helpText +=
        "**📅 Available Commands:**\n" +
        "• `calendar` or `events` - View your upcoming calendar events\n" +
        "• `profile` or `me` - Show your profile information\n" +
        "• `logout` - Sign out from your account\n";
    } else {
      helpText +=
        "**🔐 Getting Started:**\n" +
        "• `login` - Sign in to your Microsoft account\n\n" +
        "Once signed in, you can manage your calendar!";
    }

    await context.sendActivity(helpText);
  }

  async showUserProfile(context, userId, accessToken) {
    try {
      const client = getGraphClient(accessToken);
      await context.sendActivity("🔍 Fetching your profile...");

      const user = await client.api("/me").get();

      const profileText =
        `👤 **Your Profile**\n\n` +
        `**Name:** ${user.displayName || "N/A"}\n` +
        `**Email:** ${user.mail || user.userPrincipalName || "N/A"}\n` +
        `**Job Title:** ${user.jobTitle || "N/A"}\n` +
        `**Office:** ${user.officeLocation || "N/A"}`;

      await context.sendActivity(profileText);
      console.log("   ✅ Profile displayed");
    } catch (error) {
      console.error("❌ Error fetching profile:", error);
      await this.handleGraphError(context, userId, error);
    }
  }

  async listCalendarEvents(context, userId, accessToken) {
    try {
      const client = getGraphClient(accessToken);
      await context.sendActivity("🔍 Fetching your calendar events...");

      const events = await client
        .api("/me/calendar/events")
        .top(10)
        .select("subject,start,end,organizer,location")
        .filter(`start/dateTime ge '${new Date().toISOString()}'`)
        .orderby("start/dateTime")
        .get();

      if (events.value.length === 0) {
        await context.sendActivity("📅 You have no upcoming events.");
        return;
      }

      let message = `📅 **Your Upcoming Events (${events.value.length}):**\n\n`;

      events.value.forEach((event, index) => {
        const startDate = new Date(event.start.dateTime);
        const endDate = new Date(event.end.dateTime);
        const duration = Math.round((endDate - startDate) / 60000);

        message += `**${index + 1}. ${event.subject}**\n`;
        message += `   📅 ${startDate.toLocaleDateString()} at ${startDate.toLocaleTimeString(
          [],
          { hour: "2-digit", minute: "2-digit" }
        )}\n`;
        message += `   ⏱️ ${duration} minutes\n`;

        if (event.location?.displayName) {
          message += `   📍 ${event.location.displayName}\n`;
        }

        message += "\n";
      });

      await context.sendActivity(message);
      console.log("   ✅ Calendar events sent");
    } catch (error) {
      console.error("❌ Error fetching calendar:", error);
      await this.handleGraphError(context, userId, error);
    }
  }

  async logout(context, userId) {
    tokenStore.delete(userId);

    await context.sendActivity(
      "✅ **Signed Out Successfully**\n\n" +
        "You've been signed out. Type **login** to sign in again."
    );

    console.log(`   🔓 User signed out`);
  }

  async handleGraphError(context, userId, error) {
    if (
      error.statusCode === 401 ||
      error.code === "InvalidAuthenticationToken"
    ) {
      tokenStore.delete(userId);
      await context.sendActivity(
        "⚠️ Your session has expired. Please type **login** to sign in again."
      );
    } else {
      await context.sendActivity(
        "❌ Sorry, I encountered an error. Please try again later."
      );
    }
  }
}

const bot = new MeetingSchedulerBot(conversationState, userState);

// ============================================================================
// API ENDPOINTS
// ============================================================================

app.post("/api/messages", async (req, res) => {
  console.log("\n📨 Incoming message");

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

app.get("/api/oauth/callback", async (req, res) => {
  console.log("\n🔐 OAuth callback received");

  const { code, state, error, error_description } = req.query;

  if (error) {
    console.error("❌ OAuth error:", error, error_description);
    return res.send(generateErrorPage(error_description || error));
  }

  const stateData = oauthStateStore.get(state);
  if (!stateData) {
    console.error("❌ Invalid or expired OAuth state");
    return res.send(
      generateErrorPage("Invalid or expired session. Please try again.")
    );
  }

  try {
    const tokenData = await exchangeCodeForToken(code);

    const now = Date.now();
    tokenStore.set(stateData.userId, {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: now + tokenData.expires_in * 1000,
    });

    console.log("✅ Token stored for user:", stateData.userId);
    oauthStateStore.delete(state);

    res.send(generateSuccessPage());
  } catch (error) {
    console.error("❌ Token exchange failed:", error);
    res.send(
      generateErrorPage("Failed to complete sign-in. Please try again.")
    );
  }
});

app.get("/", (req, res) => {
  res.json({
    status: "✅ Bot is running",
    timestamp: new Date().toISOString(),
  });
});

app.get("/health", (req, res) => {
  res.json({
    status: "healthy",
    uptime: process.uptime(),
  });
});

// ============================================================================
// HTML GENERATORS
// ============================================================================

function generateSuccessPage() {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sign-in Successful</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      margin: 0;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .container {
      background: white;
      padding: 40px;
      border-radius: 12px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.2);
      text-align: center;
      max-width: 400px;
    }
    .icon { font-size: 64px; margin-bottom: 20px; }
    h1 { color: #2d3748; margin: 0 0 10px 0; }
    p { color: #4a5568; line-height: 1.6; }
    .button {
      background: #667eea;
      color: white;
      border: none;
      padding: 12px 24px;
      border-radius: 6px;
      font-size: 16px;
      cursor: pointer;
      margin-top: 20px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">✅</div>
    <h1>Sign-in Successful!</h1>
    <p>You've been authenticated successfully.</p>
    <p>Close this window and return to Microsoft Teams.</p>
    <p><strong>Type "calendar"</strong> to see your events!</p>
    <button class="button" onclick="window.close()">Close Window</button>
  </div>
  <script>setTimeout(() => window.close(), 5000);</script>
</body>
</html>
  `;
}

function generateErrorPage(errorMessage) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sign-in Failed</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      margin: 0;
      background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    }
    .container {
      background: white;
      padding: 40px;
      border-radius: 12px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.2);
      text-align: center;
      max-width: 400px;
    }
    .icon { font-size: 64px; margin-bottom: 20px; }
    h1 { color: #2d3748; margin: 0 0 10px 0; }
    p { color: #4a5568; line-height: 1.6; }
    .error { background: #fee; padding: 12px; border-radius: 6px; margin: 20px 0; color: #c53030; font-size: 14px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">❌</div>
    <h1>Sign-in Failed</h1>
    <div class="error">${errorMessage}</div>
    <p>Please close this window and try again from Teams.</p>
  </div>
</body>
</html>
  `;
}

// ============================================================================
// START SERVER
// ============================================================================

app.listen(PORT, () => {
  console.log("\n" + "=".repeat(60));
  console.log("🤖 Meeting Scheduler Bot with Manual OAuth");
  console.log("=".repeat(60));
  console.log(`📍 Port: ${PORT}`);
  console.log(`💬 Messages: ${BASE_URL}/api/messages`);
  console.log(`🔐 OAuth Callback: ${REDIRECT_URI}`);
  console.log("=".repeat(60) + "\n");
});
