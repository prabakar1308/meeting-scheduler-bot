const { ActivityHandler, MessageFactory, CardFactory } = require("botbuilder");
const { ScheduleDialog } = require("./dialogs/scheduleDialog");
const { DialogSet, DialogTurnStatus } = require("botbuilder-dialogs");
const { AuthService } = require("./services/authService");

class MeetingSchedulerBot extends ActivityHandler {
  constructor(conversationState, userState) {
    super();

    this.conversationState = conversationState;
    this.userState = userState;
    this.authService = new AuthService();

    // Create state property accessors
    this.dialogState = this.conversationState.createProperty("DialogState");
    this.userProfile = this.userState.createProperty("UserProfile");

    // Create dialog set
    this.dialogs = new DialogSet(this.dialogState);

    // Add schedule dialog
    this.scheduleDialog = new ScheduleDialog(
      "scheduleDialog",
      this.userProfile
    );
    this.dialogs.add(this.scheduleDialog);

    this.onMessage(async (context, next) => {
      const userProfile = await this.userProfile.get(context, {});
      const text = context.activity.text?.trim();

      console.log(`\n💬 Message from ${context.activity.from.name}: "${text}"`);
      console.log(`   Has token: ${!!userProfile.accessToken}`);

      // Debug commands (remove in production!)
      if (text && text.toLowerCase().startsWith("debug:")) {
        console.log("   🐛 Debug command detected");
        await this.handleDebugCommand(context, text.toLowerCase(), userProfile);
        await next();
        return;
      }

      // Handle magic code from OAuth (format: auth:CODE or just the code)
      if (
        text &&
        (text.toLowerCase().startsWith("auth:") || this.looksLikeAuthCode(text))
      ) {
        const code = text.toLowerCase().startsWith("auth:")
          ? text.substring(5).trim()
          : text.trim();

        await this.completeAuthentication(context, code, userProfile);
        await next();
        return;
      }

      // Check if user has valid token
      if (!userProfile.accessToken) {
        console.log("❌ No access token found, initiating authentication");
        await this.initiateAuthentication(context, userProfile);
        await next();
        return;
      }

      // Check if token is expired and refresh if needed
      try {
        console.log("🔄 Checking token validity...");
        const tokens = await this.authService.ensureValidToken(userProfile);

        if (tokens.accessToken !== userProfile.accessToken) {
          // Token was refreshed, update profile
          console.log("✅ Token refreshed successfully");
          userProfile.accessToken = tokens.accessToken;
          userProfile.refreshToken = tokens.refreshToken;
          userProfile.expiresOn = tokens.expiresOn;
          await this.userProfile.set(context, userProfile);
        } else {
          console.log("✅ Token is still valid");
        }
      } catch (error) {
        console.error("❌ Token validation/refresh failed:", error.message);
        // Clear tokens and re-authenticate
        userProfile.accessToken = null;
        userProfile.refreshToken = null;
        userProfile.expiresOn = null;
        await this.userProfile.set(context, userProfile);

        await context.sendActivity(
          "Your session has expired. Please authenticate again."
        );
        await this.initiateAuthentication(context, userProfile);
        await next();
        return;
      }

      // Token is valid, proceed with dialog
      console.log("✅ User authenticated, processing message...");
      const dialogContext = await this.dialogs.createContext(context);
      const results = await dialogContext.continueDialog();

      if (results.status === DialogTurnStatus.empty) {
        await dialogContext.beginDialog("scheduleDialog");
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      const welcomeText =
        "👋 Hello! I'm your Meeting Scheduler Assistant.\n\n" +
        "I can help you schedule meetings in your Outlook calendar using natural language.\n\n" +
        "📝 Examples:\n" +
        '• "Schedule a team sync tomorrow at 2 PM for 1 hour"\n' +
        '• "Book a meeting with john@example.com next Monday"\n' +
        '• "Create a planning session on Friday at 3 PM"\n\n' +
        "To get started, I'll need access to your calendar.";

      for (let member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(MessageFactory.text(welcomeText));

          // Check if already authenticated
          const userProfile = await this.userProfile.get(context, {});
          if (!userProfile.accessToken) {
            await context.sendActivity("Let me help you authenticate...");
            await this.initiateAuthentication(context, userProfile);
          } else {
            await context.sendActivity(
              `Welcome back${
                userProfile.displayName ? ", " + userProfile.displayName : ""
              }! How can I help you today?`
            );
          }
        }
      }

      await next();
    });
  }

  // Helper to identify if text looks like an authorization code
  looksLikeAuthCode(text) {
    // OAuth codes are typically long alphanumeric strings
    // Check if it's a long string without spaces and contains both letters and numbers
    if (!text || text.length < 20) return false;
    if (text.includes(" ")) return false;

    const hasLetters = /[a-zA-Z]/.test(text);
    const hasNumbers = /[0-9]/.test(text);
    const isAlphanumeric = /^[a-zA-Z0-9._-]+$/.test(text);

    return hasLetters && hasNumbers && isAlphanumeric && text.length > 30;
  }

  async initiateAuthentication(context, userProfile) {
    // Generate a unique state for this auth request
    const state = `${context.activity.from.id}_${Date.now()}`;

    try {
      const authUrl = await this.authService.getAuthCodeUrl(state);

      console.log("🔗 Generated auth URL for user:", context.activity.from.id);

      const card = CardFactory.heroCard(
        "🔐 Authentication Required",
        "I need permission to access your Outlook calendar to schedule meetings.",
        [],
        [
          {
            type: "openUrl",
            title: "Sign in with Microsoft",
            value: authUrl,
          },
        ]
      );

      await context.sendActivity({ attachments: [card] });

      // Provide alternative instruction
      await context.sendActivity(
        "**Instructions:**\n" +
          '1. Click the "Sign in with Microsoft" button above\n' +
          "2. Complete the login in your browser\n" +
          "3. Copy the code from the success page\n" +
          "4. Return here and send: `auth:YOUR_CODE`\n\n" +
          '💡 Tip: You can also just paste the code directly without "auth:"'
      );
    } catch (error) {
      console.error("❌ Error generating auth URL:", error);
      await context.sendActivity(
        "❌ Sorry, I encountered an error setting up authentication. " +
          "Please make sure OAuth is configured correctly.\n\n" +
          `Error: ${error.message}`
      );
    }
  }

  async completeAuthentication(context, code, userProfile) {
    console.log("\n🔐 Completing authentication...");
    console.log(`   Code length: ${code.length} characters`);
    console.log(`   User: ${context.activity.from.name}`);

    try {
      await context.sendActivity("🔄 Processing your authentication...");

      // Exchange code for tokens
      console.log("📤 Exchanging code for tokens...");
      const tokens = await this.authService.acquireTokenByCode(code);

      console.log("✅ Tokens acquired successfully");
      console.log(`   Access token length: ${tokens.accessToken.length}`);
      console.log(`   Has refresh token: ${!!tokens.refreshToken}`);
      console.log(`   Expires on: ${tokens.expiresOn}`);

      // Store tokens in user profile
      userProfile.accessToken = tokens.accessToken;
      userProfile.refreshToken = tokens.refreshToken;
      userProfile.expiresOn = tokens.expiresOn;
      userProfile.authenticatedAt = new Date().toISOString();

      // Save to state
      await this.userProfile.set(context, userProfile);

      console.log("💾 Tokens saved to user profile");

      // Get user info to personalize
      try {
        console.log("📤 Fetching user profile from Graph API...");
        const axios = require("axios");
        const response = await axios.get(
          "https://graph.microsoft.com/v1.0/me",
          {
            headers: { Authorization: `Bearer ${tokens.accessToken}` },
          }
        );

        userProfile.displayName = response.data.displayName;
        userProfile.email =
          response.data.mail || response.data.userPrincipalName;
        await this.userProfile.set(context, userProfile);

        console.log(
          `✅ User profile retrieved: ${userProfile.displayName} (${userProfile.email})`
        );

        await context.sendActivity(
          `✅ **Authentication Successful!**\n\n` +
            `Welcome, **${response.data.displayName}**!\n\n` +
            `📧 Email: ${userProfile.email}\n` +
            `🔑 Access granted to your calendar\n\n` +
            `I'm ready to help you schedule meetings. What would you like to do?`
        );
      } catch (graphError) {
        console.error("⚠️  Could not fetch user profile:", graphError.message);
        await context.sendActivity(
          "✅ **Authentication Successful!**\n\n" +
            "I can now access your calendar. How can I help you schedule a meeting?"
        );
      }

      // Important: Save state changes immediately
      await this.userState.saveChanges(context, false);
      await this.conversationState.saveChanges(context, false);

      console.log("✅ Authentication completed successfully\n");
    } catch (error) {
      console.error("❌ Authentication failed:", error);
      console.error("   Error name:", error.name);
      console.error("   Error message:", error.message);
      console.error("   Error code:", error.errorCode);

      await context.sendActivity(
        "❌ **Authentication Failed**\n\n" +
          `Error: ${error.message}\n\n` +
          "Please try the following:\n" +
          "1. Make sure you copied the entire code\n" +
          "2. Try authenticating again (codes expire quickly)\n" +
          "3. Check that your OAuth configuration is correct\n\n" +
          "Type anything to get a new authentication link."
      );
    }
  }

  async run(context) {
    await super.run(context);

    // Save any state changes - CRITICAL!
    await this.conversationState.saveChanges(context, false);
    await this.userState.saveChanges(context, false);
  }

  async handleDebugCommand(context, command, userProfile) {
    const cmd = command.substring(6).trim(); // Remove 'debug:'

    switch (cmd) {
      case "state":
        const stateInfo = {
          hasToken: !!userProfile.accessToken,
          hasRefreshToken: !!userProfile.refreshToken,
          displayName: userProfile.displayName || "Not set",
          email: userProfile.email || "Not set",
          authenticatedAt: userProfile.authenticatedAt || "Never",
          tokenLength: userProfile.accessToken
            ? userProfile.accessToken.length
            : 0,
          expiresOn: userProfile.expiresOn || "Not set",
        };

        await context.sendActivity(
          "🔍 **Debug: User State**\n\n" +
            `Has Access Token: ${stateInfo.hasToken ? "✅ Yes" : "❌ No"}\n` +
            `Has Refresh Token: ${
              stateInfo.hasRefreshToken ? "✅ Yes" : "❌ No"
            }\n` +
            `Display Name: ${stateInfo.displayName}\n` +
            `Email: ${stateInfo.email}\n` +
            `Authenticated At: ${stateInfo.authenticatedAt}\n` +
            `Token Length: ${stateInfo.tokenLength} chars\n` +
            `Expires On: ${stateInfo.expiresOn}`
        );
        console.log("Debug state:", JSON.stringify(stateInfo, null, 2));
        break;

      case "token":
        if (!userProfile.accessToken) {
          await context.sendActivity(
            "❌ No access token available. Please authenticate first."
          );
          break;
        }

        await context.sendActivity("🔍 Testing access token...");

        const GraphService = require("./services/graphService").GraphService;
        const graphService = new GraphService();

        // Test token validity
        const validation = await graphService.validateToken(
          userProfile.accessToken
        );

        if (validation.valid) {
          await context.sendActivity(
            `✅ **Token is Valid**\n\n` +
              `User: ${validation.user.displayName}\n` +
              `Email: ${
                validation.user.mail || validation.user.userPrincipalName
              }`
          );

          // Test calendar access
          const calendarTest = await graphService.testCalendarAccess(
            userProfile.accessToken
          );

          if (calendarTest.hasAccess) {
            await context.sendActivity(
              `✅ **Calendar Access Granted**\n\n` +
                `Calendars found: ${calendarTest.calendars.length}`
            );
          } else {
            await context.sendActivity(
              `❌ **Calendar Access Denied**\n\n` +
                `Error: ${calendarTest.error}\n\n` +
                `Make sure Calendars.ReadWrite permission is granted in Azure Portal.`
            );
          }
        } else {
          await context.sendActivity(
            `❌ **Token is Invalid**\n\n` +
              `Error: ${validation.error}\n\n` +
              `Please re-authenticate by typing anything to get a new auth link.`
          );

          // Clear invalid token
          userProfile.accessToken = null;
          userProfile.refreshToken = null;
          await this.userProfile.set(context, userProfile);
        }
        break;

      case "clear":
        userProfile.accessToken = null;
        userProfile.refreshToken = null;
        userProfile.expiresOn = null;
        userProfile.displayName = null;
        userProfile.email = null;
        userProfile.authenticatedAt = null;
        await this.userProfile.set(context, userProfile);
        await this.userState.saveChanges(context, false);

        await context.sendActivity(
          "🗑️ All tokens and user data cleared. Type anything to re-authenticate."
        );
        console.log("✅ User state cleared");
        break;

      case "help":
        await context.sendActivity(
          "🛠️ **Debug Commands:**\n\n" +
            "`debug:state` - Show authentication state\n" +
            "`debug:token` - Validate token and test calendar access\n" +
            "`debug:clear` - Clear all tokens\n" +
            "`debug:help` - Show this help\n\n" +
            "**Normal Commands:**\n" +
            "`auth:CODE` - Complete authentication with code"
        );
        break;

      default:
        await context.sendActivity(
          "❓ Unknown debug command. Try `debug:help`"
        );
    }
  }
}

module.exports.MeetingSchedulerBot = MeetingSchedulerBot;
