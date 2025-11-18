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
  MessageFactory,
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

// Get the base URL from environment or construct it
const BASE_URL =
  process.env.BASE_URL || "https://meeting-scheduler-bot-i5oj.onrender.com";
const REDIRECT_URI = `${BASE_URL}/api/oauth/callback`;

// Scopes we need from Microsoft Graph
const SCOPES =
  "User.Read Calendars.ReadWrite Calendars.ReadWrite.Shared offline_access";

console.log("\n🔧 OAuth Configuration:");
console.log("   App ID:", MICROSOFT_APP_ID ? "✓ Configured" : "✗ MISSING");
console.log(
  "   App Password:",
  MICROSOFT_APP_PASSWORD ? "✓ Configured" : "✗ MISSING"
);
console.log("   Tenant ID:", TENANT_ID);
console.log("   Redirect URI:", REDIRECT_URI);
console.log("   Scopes:", SCOPES);

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
      "⚠️ Oops! Something went wrong. Please try again."
    );
    await conversationState.delete(context);
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
// IN-MEMORY STORAGE FOR OAUTH
// NOTE: In production, use Redis or a database
// ============================================================================

const oauthStateStore = new Map(); // Stores OAuth state → user mapping
const tokenStore = new Map(); // Stores userId → access token
const conversationReferenceStore = new Map(); // Stores userId → conversation reference

// Clean up expired OAuth states every 5 minutes
setInterval(() => {
  const tenMinutesAgo = Date.now() - 10 * 60 * 1000;
  for (const [state, data] of oauthStateStore.entries()) {
    if (data.timestamp < tenMinutesAgo) {
      oauthStateStore.delete(state);
      console.log(`🧹 Cleaned up expired OAuth state: ${state}`);
    }
  }
}, 5 * 60 * 1000);

// ============================================================================
// OAUTH HELPER FUNCTIONS
// ============================================================================

/**
 * Generate OAuth authorization URL
 */
function generateOAuthUrl(userId, conversationReference) {
  const state = crypto.randomBytes(16).toString("hex");

  // Store state with user info for callback verification
  oauthStateStore.set(state, {
    userId,
    conversationReference,
    timestamp: Date.now(),
  });

  // Build Microsoft OAuth URL
  const authUrl =
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    `client_id=${encodeURIComponent(MICROSOFT_APP_ID)}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&response_mode=query` +
    `&scope=${encodeURIComponent(SCOPES)}` +
    `&state=${state}` +
    `&prompt=select_account`; // Always show account picker

  console.log(`🔗 Generated OAuth URL for user: ${userId}`);
  return authUrl;
}

/**
 * Exchange authorization code for access token
 */
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

/**
 * Refresh access token using refresh token
 */
async function refreshAccessToken(refreshToken) {
  try {
    const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

    const params = new URLSearchParams();
    params.append("client_id", MICROSOFT_APP_ID);
    params.append("scope", SCOPES);
    params.append("refresh_token", refreshToken);
    params.append("grant_type", "refresh_token");
    params.append("client_secret", MICROSOFT_APP_PASSWORD);

    const response = await axios.post(tokenUrl, params, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });

    console.log("✅ Token refreshed successfully");
    return response.data;
  } catch (error) {
    console.error("❌ Token refresh error:", error.response?.data);
    throw error;
  }
}

/**
 * Get valid access token for user (refresh if expired)
 */
async function getValidToken(userId) {
  const tokenData = tokenStore.get(userId);

  if (!tokenData) {
    return null;
  }

  // Check if token is expired or about to expire (5 min buffer)
  const now = Date.now();
  const expiresAt = tokenData.expires_at || 0;

  if (now >= expiresAt - 5 * 60 * 1000) {
    // Token expired or about to expire, refresh it
    if (tokenData.refresh_token) {
      try {
        console.log("🔄 Refreshing expired token for user:", userId);
        const newTokenData = await refreshAccessToken(tokenData.refresh_token);

        // Update stored token
        tokenStore.set(userId, {
          access_token: newTokenData.access_token,
          refresh_token: newTokenData.refresh_token || tokenData.refresh_token,
          expires_at: now + newTokenData.expires_in * 1000,
        });

        return newTokenData.access_token;
      } catch (error) {
        console.error("❌ Failed to refresh token:", error);
        // Clear invalid token
        tokenStore.delete(userId);
        return null;
      }
    } else {
      // No refresh token, clear and require re-auth
      tokenStore.delete(userId);
      return null;
    }
  }

  return tokenData.access_token;
}

/**
 * Create Microsoft Graph client
 */
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
    // Store conversation reference for proactive messaging
    const userId = context.activity.from.id;
    conversationReferenceStore.set(userId, {
      conversationReference: context.activity.getConversationReference(),
      serviceUrl: context.activity.serviceUrl,
    });

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

    console.log(
      `\n📩 Message from ${userName} (${userId}): "${context.activity.text}"`
    );

    // Check if user has valid token
    const accessToken = await getValidToken(userId);
    const isAuthenticated = !!accessToken;

    console.log(
      `   Authentication status: ${
        isAuthenticated ? "✅ Authenticated" : "❌ Not authenticated"
      }`
    );

    // Handle unauthenticated users
    if (!isAuthenticated) {
      if (
        text.includes("login") ||
        text.includes("signin") ||
        text.includes("start") ||
        text.includes("hello") ||
        text.includes("hi")
      ) {
        await this.sendSignInCard(context);
      } else {
        await context.sendActivity(
          "👋 Welcome to **Meeting Scheduler Bot**!\n\n" +
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
      text.includes("create") ||
      text.includes("schedule") ||
      text.includes("new meeting")
    ) {
      await this.createMeeting(context, userId, accessToken);
    } else if (
      text.includes("profile") ||
      text.includes("me") ||
      text.includes("whoami")
    ) {
      await this.showUserProfile(context, userId, accessToken);
    } else if (
      text.includes("logout") ||
      text.includes("signout") ||
      text.includes("sign out")
    ) {
      await this.logout(context, userId);
    } else if (text.includes("help")) {
      await this.sendHelpMessage(context, true);
    } else {
      await context.sendActivity(
        `You said: "${context.activity.text}"\n\n` +
          "Type **help** to see what I can do!"
      );
    }

    console.log("   ✅ Response sent");
  }

  async handleConversationUpdate(context) {
    if (context.activity.membersAdded) {
      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          console.log(`   👤 New member: ${member.name || member.id}`);

          const card = CardFactory.heroCard(
            "👋 Welcome to Meeting Scheduler Bot!",
            "I can help you manage your Microsoft Calendar and schedule meetings.",
            [],
            [
              { type: "imBack", title: "🔐 Sign In", value: "login" },
              { type: "imBack", title: "❓ Help", value: "help" },
            ]
          );

          await context.sendActivity({ attachments: [card] });
        }
      }
    }
  }

  async sendSignInCard(context) {
    const userId = context.activity.from.id;
    const conversationRef = context.activity.getConversationReference();

    const authUrl = generateOAuthUrl(userId, conversationRef);

    const card = CardFactory.signinCard(
      "Sign in to Microsoft",
      authUrl,
      "Click below to sign in with your Microsoft account"
    );

    await context.sendActivity({ attachments: [card] });
    await context.sendActivity(
      "💡 **Tip:** After signing in, close the browser window and return here."
    );

    console.log("   🔐 Sign-in card sent");
  }

  async sendHelpMessage(context, isAuthenticated) {
    let helpText = "**📚 Meeting Scheduler Bot - Help**\n\n";

    if (isAuthenticated) {
      helpText +=
        "**📅 Calendar Commands:**\n" +
        "• `calendar` or `events` - View your upcoming calendar events\n" +
        "• `create` or `schedule` - Schedule a new meeting\n" +
        "• `profile` or `me` - Show your profile information\n\n" +
        "**🔐 Account:**\n" +
        "• `logout` - Sign out from your account\n";
    } else {
      helpText +=
        "**🔐 Getting Started:**\n" +
        "• `login` - Sign in to your Microsoft account\n\n" +
        "Once signed in, you'll be able to:\n" +
        "• View your calendar events\n" +
        "• Schedule meetings\n" +
        "• Manage your calendar\n";
    }

    await context.sendActivity(helpText);
  }

  async showUserProfile(context, userId, accessToken) {
    try {
      const client = getGraphClient(accessToken);

      await context.sendActivity("🔍 Fetching your profile...");

      const user = await client.api("/me").get();

      const profileCard = CardFactory.thumbnailCard(
        user.displayName || "User",
        user.jobTitle || "No job title",
        [
          `📧 ${user.mail || user.userPrincipalName}`,
          user.mobilePhone ? `📱 ${user.mobilePhone}` : "",
          user.officeLocation ? `🏢 ${user.officeLocation}` : "",
        ].filter(Boolean)
      );

      await context.sendActivity({ attachments: [profileCard] });
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
        await context.sendActivity(
          "📅 You have no upcoming events in your calendar."
        );
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
        message += `   ⏱️ Duration: ${duration} minutes\n`;

        if (event.location?.displayName) {
          message += `   📍 Location: ${event.location.displayName}\n`;
        }

        if (event.organizer?.emailAddress?.name) {
          message += `   👤 Organizer: ${event.organizer.emailAddress.name}\n`;
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

  async createMeeting(context, userId, accessToken) {
    // Simple demo - in production, use dialogs for multi-turn conversation
    await context.sendActivity(
      "📅 **Create a Meeting**\n\n" +
        "This is a demo feature. In a full implementation, I would:\n" +
        "1. Ask for meeting title\n" +
        "2. Ask for date and time\n" +
        "3. Ask for attendees\n" +
        "4. Create the event in your calendar\n\n" +
        "For now, try typing `calendar` to see your existing events!"
    );
  }

  async logout(context, userId) {
    tokenStore.delete(userId);
    conversationReferenceStore.delete(userId);

    await context.sendActivity(
      "✅ **Signed Out Successfully**\n\n" +
        "You've been signed out. Type **login** to sign in again."
    );

    console.log(`   🔓 User ${userId} signed out`);
  }

  async handleGraphError(context, userId, error) {
    if (
      error.statusCode === 401 ||
      error.code === "InvalidAuthenticationToken"
    ) {
      // Token invalid, clear it
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

/**
 * Bot messaging endpoint
 */
app.post("/api/messages", async (req, res) => {
  console.log("\n" + "=".repeat(60));
  console.log("📨 Incoming message");

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

/**
 * OAuth callback endpoint - handles redirect from Microsoft
 */
app.get("/api/oauth/callback", async (req, res) => {
  console.log("\n🔐 OAuth callback received");

  const { code, state, error, error_description } = req.query;

  // Handle OAuth errors
  if (error) {
    console.error("❌ OAuth error:", error, error_description);
    return res.send(generateErrorPage(error_description || error));
  }

  // Validate state parameter
  const stateData = oauthStateStore.get(state);
  if (!stateData) {
    console.error("❌ Invalid or expired OAuth state");
    return res.send(
      generateErrorPage(
        "Invalid or expired session. Please try signing in again."
      )
    );
  }

  try {
    // Exchange code for tokens
    const tokenData = await exchangeCodeForToken(code);

    // Store tokens for user
    const now = Date.now();
    tokenStore.set(stateData.userId, {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: now + tokenData.expires_in * 1000,
    });

    console.log("✅ Token stored for user:", stateData.userId);

    // Clean up OAuth state
    oauthStateStore.delete(state);

    // Send success page
    res.send(generateSuccessPage());

    // Optional: Send proactive message to user in Teams
    // This requires additional setup with conversation reference
  } catch (error) {
    console.error("❌ Token exchange failed:", error);
    res.send(
      generateErrorPage("Failed to complete sign-in. Please try again.")
    );
  }
});

/**
 * Health check endpoint
 */
app.get("/", (req, res) => {
  res.json({
    status: "✅ Bot is running",
    timestamp: new Date().toISOString(),
    oauth: "Manual OAuth configured",
  });
});

app.get("/health", (req, res) => {
  res.json({
    status: "healthy",
    uptime: process.uptime(),
    memory: process.memoryUsage(),
  });
});

// ============================================================================
// HTML PAGE GENERATORS
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
    .success-icon {
      font-size: 64px;
      margin-bottom: 20px;
    }
    h1 {
      color: #2d3748;
      margin: 0 0 10px 0;
    }
    p {
      color: #4a5568;
      line-height: 1.6;
    }
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
    <div class="success-icon">✅</div>
    <h1>Sign-in Successful!</h1>
    <p>You've been authenticated successfully.</p>
    <p>You can now close this window and return to Microsoft Teams.</p>
    <p><strong>Try typing "calendar"</strong> to see your events!</p>
    <button class="button" onclick="window.close()">Close Window</button>
  </div>
  <script>
    // Auto-close after 5 seconds
    setTimeout(() => window.close(), 5000);
  </script>
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
    .error-icon {
      font-size: 64px;
      margin-bottom: 20px;
    }
    h1 {
      color: #2d3748;
      margin: 0 0 10px 0;
    }
    p {
      color: #4a5568;
      line-height: 1.6;
    }
    .error-message {
      background: #fee;
      padding: 12px;
      border-radius: 6px;
      margin: 20px 0;
      color: #c53030;
      font-size: 14px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="error-icon">❌</div>
    <h1>Sign-in Failed</h1>
    <div class="error-message">${errorMessage}</div>
    <p>Please close this window and try signing in again from Microsoft Teams.</p>
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
  console.log(`🌐 Base URL: ${BASE_URL}`);
  console.log(`💬 Messages: ${BASE_URL}/api/messages`);
  console.log(`🔐 OAuth Callback: ${REDIRECT_URI}`);
  console.log("\n✅ Configuration Check:");
  console.log(`   App ID: ${MICROSOFT_APP_ID ? "✓" : "✗ MISSING"}`);
  console.log(`   App Password: ${MICROSOFT_APP_PASSWORD ? "✓" : "✗ MISSING"}`);
  console.log(`   Tenant ID: ${TENANT_ID}`);
  console.log("=".repeat(60) + "\n");

  if (!MICROSOFT_APP_ID || !MICROSOFT_APP_PASSWORD) {
    console.error("⚠️  WARNING: Missing required environment variables!");
    console.error(
      "   Please set MICROSOFT_APP_ID and MICROSOFT_APP_PASSWORD\n"
    );
  }
});
