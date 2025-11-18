require("dotenv").config();
const express = require("express");
const {
  CloudAdapter,
  MemoryStorage,
  ConversationState,
  UserState,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require("botbuilder");
const { MeetingSchedulerBot } = require("./bot");
const { AuthService } = require("./services/authService");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3978;

// Create ConfigurationServiceClientCredentialFactory
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
});

// Create BotFrameworkAuthentication instance
const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);

// Update adapter to use botFrameworkAuthentication
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Error handling
adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity("Please check the console for more details.");

  // Clear conversation state to avoid stuck states
  await conversationState.delete(context);
};

// Create conversation and user state
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create auth service
const authService = new AuthService();

// Create the bot
const bot = new MeetingSchedulerBot(conversationState, userState);

// Store for tracking OAuth state (in production, use Redis or similar)
const oauthStateStore = new Map();

// Clean up old OAuth states periodically (older than 10 minutes)
setInterval(() => {
  const tenMinutesAgo = Date.now() - 10 * 60 * 1000;
  for (const [state, data] of oauthStateStore.entries()) {
    if (data.timestamp < tenMinutesAgo) {
      oauthStateStore.delete(state);
      console.log(`🧹 Cleaned up expired OAuth state: ${state}`);
    }
  }
}, 5 * 60 * 1000); // Run every 5 minutes

// Bot messaging endpoint
// app.post("/api/messages", async (req, res) => {
//   const fromName = req.body.from?.name || "Unknown User";
//   const messageText = req.body.text || "[No text]";

//   console.log(
//     `\n📨 Message from ${fromName}: ${messageText.substring(0, 50)}${
//       messageText.length > 50 ? "..." : ""
//     }`
//   );

//   try {
//     await adapter.process(req, res, async (context) => {
//       // Store OAuth state for this conversation
//       if (context.activity.from && context.activity.conversation) {
//         const state = `${context.activity.from.id}_${Date.now()}`;
//         oauthStateStore.set(state, {
//           userId: context.activity.from.id,
//           conversationId: context.activity.conversation.id,
//           serviceUrl: context.activity.serviceUrl,
//           timestamp: Date.now(),
//         });
//       }

//       await bot.run(context);
//     });
//   } catch (error) {
//     console.error("❌ Error processing message:", error);
//     if (!res.headersSent) {
//       res.status(500).json({
//         error: "Internal server error",
//         message: error.message,
//       });
//     }
//   }
// });

// Main bot endpoint
app.post("/api/messages", async (req, res) => {
  console.log("\n" + "=".repeat(60));
  console.log("📨 Incoming Request");
  console.log("   From:", req.body.from?.name || "Unknown");
  console.log("   Type:", req.body.type);

  try {
    await adapter.process(req, res, async (context) => {
      await bot.run(context);
    });
  } catch (error) {
    console.error("\n❌ ERROR in /api/messages:");
    console.error("   Message:", error.message);
    console.error("   Stack:", error.stack);

    if (!res.headersSent) {
      res.status(500).json({
        error: "Internal server error",
        message: error.message,
      });
    }
  }
});

// OAuth callback endpoint
app.get("/api/oauth/callback", async (req, res) => {
  const code = req.query.code;
  const error = req.query.error;
  const state = req.query.state;
  const errorDescription = req.query.error_description;

  console.log("\n🔐 OAuth Callback Received");
  console.log(
    `   Code: ${code ? "✓ Present (" + code.length + " chars)" : "✗ Missing"}`
  );
  console.log(`   State: ${state || "None"}`);
  console.log(`   Error: ${error || "None"}`);

  // Handle OAuth error
  if (error) {
    console.error(`❌ OAuth Error: ${error} - ${errorDescription}`);
    return res.status(400).send(generateErrorPage(error, errorDescription));
  }

  // Validate authorization code
  if (!code) {
    console.error("❌ No authorization code received");
    return res
      .status(400)
      .send(
        generateErrorPage(
          "missing_code",
          "No authorization code was received from Microsoft"
        )
      );
  }

  // Retrieve OAuth state data
  const stateData = state ? oauthStateStore.get(state) : null;

  if (stateData) {
    console.log(`✅ Found OAuth state for user: ${stateData.userId}`);
    // Clean up used state
    oauthStateStore.delete(state);
  } else {
    console.warn("⚠️  OAuth state not found or expired");
  }

  // Return success page with code for user to paste
  console.log("✅ Authorization successful, displaying code to user");
  res.send(generateSuccessPage(code, state));
});

// Token exchange endpoint (for proactive authentication completion)
app.post("/api/oauth/exchange", async (req, res) => {
  const { code, userId } = req.body;

  console.log(`\n🔄 Token exchange requested for user: ${userId}`);

  if (!code) {
    return res.status(400).json({
      success: false,
      error: "Authorization code is required",
    });
  }

  try {
    // Exchange code for tokens
    const tokens = await authService.acquireTokenByCode(code);

    console.log("✅ Tokens acquired successfully");

    // In production, you would:
    // 1. Store tokens securely (encrypted in database)
    // 2. Associate with user ID
    // 3. Send proactive message to user confirming auth

    res.json({
      success: true,
      message: "Authentication completed successfully",
      expiresOn: tokens.expiresOn,
    });
  } catch (error) {
    console.error("❌ Token exchange failed:", error.message);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// Health check endpoint
app.get("/health", (req, res) => {
  const health = {
    status: "healthy",
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    environment: {
      nodeVersion: process.version,
      port: PORT,
      oauthConfigured: !!process.env.MICROSOFT_APP_ID,
    },
    oauth: {
      clientId: process.env.MICROSOFT_APP_ID ? "configured" : "not configured",
      redirectUri: process.env.OAUTH_REDIRECT_URI || "not configured",
    },
  };

  res.json(health);
});

// Root endpoint with documentation
app.get("/", (req, res) => {
  res.send(generateHomePage());
});

// Test endpoint to verify OAuth URL generation
app.get("/api/oauth/test", async (req, res) => {
  try {
    const testState = "test_" + Date.now();
    const authUrl = await authService.getAuthCodeUrl(testState);
    res.json({
      success: true,
      authUrl: authUrl,
      state: testState,
      message:
        "OAuth URL generated successfully. Use this to test authentication flow.",
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
      message: "Failed to generate OAuth URL. Check your configuration.",
    });
  }
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    error: "Not Found",
    message: `Route ${req.method} ${req.path} not found`,
    availableEndpoints: [
      "POST /api/messages",
      "GET /api/oauth/callback",
      "POST /api/oauth/exchange",
      "GET /api/oauth/test",
      "GET /health",
      "GET /",
    ],
  });
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error("Express error:", error);
  res.status(500).json({
    error: "Internal Server Error",
    message: error.message,
  });
});

// Start server
const server = app.listen(PORT, () => {
  console.log("\n" + "═".repeat(70));
  console.log("  🤖 MEETING SCHEDULER BOT - PRODUCTION SERVER");
  console.log("═".repeat(70));
  console.log("\n📊 Server Information:");
  console.log(`   Port: ${PORT}`);
  console.log(`   Environment: ${process.env.NODE_ENV || "development"}`);
  console.log(`   Node Version: ${process.version}`);

  console.log("\n🔗 Endpoints:");
  console.log(`   Bot Messages:    http://localhost:${PORT}/api/messages`);
  console.log(
    `   OAuth Callback:  http://localhost:${PORT}/api/oauth/callback`
  );
  console.log(`   Health Check:    http://localhost:${PORT}/health`);
  console.log(`   Documentation:   http://localhost:${PORT}`);

  console.log("\n🔐 OAuth Configuration:");
  if (process.env.MICROSOFT_APP_ID) {
    console.log(
      `   ✅ Client ID: ${process.env.MICROSOFT_APP_ID.substring(0, 8)}...`
    );
    console.log(`   ✅ Client Secret: ${"*".repeat(20)}`);
    console.log(`   ✅ Redirect URI: ${process.env.OAUTH_REDIRECT_URI}`);
  } else {
    console.log(
      "   ⚠️  OAuth not configured (set MICROSOFT_APP_ID and MICROSOFT_APP_PASSWORD)"
    );
  }

  console.log("\n📱 Bot Framework Emulator Setup:");
  console.log("   1. Open Bot Framework Emulator");
  console.log('   2. Click "Open Bot"');
  console.log(`   3. Bot URL: http://localhost:${PORT}/api/messages`);
  console.log("   4. For local testing: Leave credentials empty");
  console.log("   5. For OAuth testing: Enter App ID and Password");

  console.log("\n💡 Quick Test:");
  console.log(`   Visit: http://localhost:${PORT}/api/oauth/test`);
  console.log("   This will generate a test OAuth URL\n");

  console.log("═".repeat(70));
  console.log("✅ Server is ready to accept connections");
  console.log("Press Ctrl+C to stop\n");
});

// Graceful shutdown
process.on("SIGINT", () => {
  console.log("\n\n🛑 Shutting down gracefully...");
  server.close(() => {
    console.log("✅ Server closed");
    process.exit(0);
  });
});

process.on("SIGTERM", () => {
  console.log("\n\n🛑 SIGTERM received, shutting down...");
  server.close(() => {
    console.log("✅ Server closed");
    process.exit(0);
  });
});

// Handle uncaught exceptions
process.on("uncaughtException", (error) => {
  console.error("💥 Uncaught Exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("💥 Unhandled Rejection at:", promise, "reason:", reason);
});

// Helper function to generate error page
function generateErrorPage(error, description) {
  return `
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Authentication Error</title>
                <style>
                    body {
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        min-height: 100vh;
                        margin: 0;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        padding: 20px;
                    }
                    .container {
                        background: white;
                        padding: 40px;
                        border-radius: 12px;
                        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
                        text-align: center;
                        max-width: 500px;
                        width: 100%;
                    }
                    .error-icon {
                        font-size: 64px;
                        margin-bottom: 20px;
                    }
                    h1 {
                        color: #d32f2f;
                        margin: 0 0 10px 0;
                        font-size: 24px;
                    }
                    .error-code {
                        background: #ffebee;
                        padding: 15px;
                        border-radius: 8px;
                        margin: 20px 0;
                        color: #c62828;
                        font-size: 14px;
                        word-break: break-word;
                    }
                    p {
                        color: #666;
                        line-height: 1.6;
                        margin: 15px 0;
                    }
                    button {
                        background: #667eea;
                        color: white;
                        border: none;
                        padding: 12px 30px;
                        border-radius: 6px;
                        cursor: pointer;
                        font-size: 16px;
                        margin-top: 20px;
                        transition: background 0.3s;
                    }
                    button:hover {
                        background: #5568d3;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="error-icon">❌</div>
                    <h1>Authentication Failed</h1>
                    <div class="error-code">
                        <strong>Error:</strong> ${error}<br>
                        ${
                          description
                            ? `<strong>Details:</strong> ${description}`
                            : ""
                        }
                    </div>
                    <p>Please close this window and try authenticating again in your chat.</p>
                    <button onclick="window.close()">Close Window</button>
                </div>
            </body>
        </html>
    `;
}

// Helper function to generate success page
function generateSuccessPage(code, state) {
  return `
        <!DOCTYPE html>
        <html>
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Authentication Successful</title>
                <style>
                    * { box-sizing: border-box; }
                    body {
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        min-height: 100vh;
                        margin: 0;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        padding: 20px;
                    }
                    .container {
                        background: white;
                        padding: 40px;
                        border-radius: 12px;
                        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
                        text-align: center;
                        max-width: 600px;
                        width: 100%;
                        animation: slideUp 0.5s ease-out;
                    }
                    @keyframes slideUp {
                        from { transform: translateY(30px); opacity: 0; }
                        to { transform: translateY(0); opacity: 1; }
                    }
                    .success-icon {
                        font-size: 64px;
                        margin-bottom: 20px;
                        animation: scaleIn 0.6s ease-out;
                    }
                    @keyframes scaleIn {
                        from { transform: scale(0); }
                        to { transform: scale(1); }
                    }
                    h1 {
                        color: #333;
                        margin: 0 0 15px 0;
                        font-size: 28px;
                    }
                    .subtitle {
                        color: #666;
                        margin-bottom: 30px;
                        font-size: 16px;
                    }
                    .instructions {
                        background: #e3f2fd;
                        padding: 25px;
                        border-radius: 8px;
                        margin: 25px 0;
                        text-align: left;
                        border-left: 4px solid #2196f3;
                    }
                    .instructions h3 {
                        margin: 0 0 15px 0;
                        color: #1976d2;
                        font-size: 18px;
                    }
                    .instructions ol {
                        margin: 10px 0;
                        padding-left: 25px;
                    }
                    .instructions li {
                        margin: 10px 0;
                        color: #424242;
                        line-height: 1.6;
                    }
                    .code-container {
                        background: #f5f5f5;
                        padding: 20px;
                        border-radius: 8px;
                        margin: 25px 0;
                        border: 2px dashed #667eea;
                    }
                    .code-label {
                        font-weight: 600;
                        color: #333;
                        margin-bottom: 12px;
                        font-size: 14px;
                    }
                    .code {
                        font-family: 'Courier New', monospace;
                        font-size: 14px;
                        color: #667eea;
                        font-weight: bold;
                        word-break: break-all;
                        padding: 15px;
                        background: white;
                        border-radius: 6px;
                        cursor: pointer;
                        user-select: all;
                        transition: background 0.3s;
                    }
                    .code:hover {
                        background: #f0f0f0;
                    }
                    .example {
                        font-size: 13px;
                        color: #999;
                        margin-top: 15px;
                        font-style: italic;
                    }
                    .example code {
                        background: #f5f5f5;
                        padding: 2px 6px;
                        border-radius: 3px;
                        color: #e83e8c;
                        font-style: normal;
                    }
                    .btn-container {
                        display: flex;
                        gap: 12px;
                        justify-content: center;
                        margin-top: 30px;
                        flex-wrap: wrap;
                    }
                    button {
                        border: none;
                        padding: 14px 28px;
                        border-radius: 6px;
                        cursor: pointer;
                        font-size: 15px;
                        font-weight: 500;
                        transition: all 0.3s;
                        min-width: 140px;
                    }
                    .btn-copy {
                        background: #667eea;
                        color: white;
                    }
                    .btn-copy:hover {
                        background: #5568d3;
                        transform: translateY(-2px);
                        box-shadow: 0 4px 8px rgba(102, 126, 234, 0.4);
                    }
                    .btn-close {
                        background: #e0e0e0;
                        color: #333;
                    }
                    .btn-close:hover {
                        background: #d0d0d0;
                    }
                    .success-msg {
                        display: none;
                        color: #4caf50;
                        margin-top: 15px;
                        font-weight: 600;
                        animation: fadeIn 0.3s;
                    }
                    @keyframes fadeIn {
                        from { opacity: 0; }
                        to { opacity: 1; }
                    }
                    .highlight {
                        background: #fff59d;
                        padding: 2px 6px;
                        border-radius: 3px;
                        font-weight: bold;
                        color: #f57f17;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="success-icon">✅</div>
                    <h1>Authentication Successful!</h1>
                    <p class="subtitle">Your Microsoft account has been verified</p>
                    
                    <div class="instructions">
                        <h3>📋 Complete Setup (3 Steps)</h3>
                        <ol>
                            <li><strong>Copy the code below</strong> by clicking on it or using the "Copy Code" button</li>
                            <li><strong>Return to your bot conversation</strong> in the emulator or chat</li>
                            <li><strong>Send this message:</strong> <span class="highlight">auth:</span> followed by the code</li>
                        </ol>
                    </div>
                    
                    <div class="code-container">
                        <div class="code-label">🔑 Your Authentication Code:</div>
                        <div class="code" id="authCode" onclick="copyCode()" title="Click to select">${code}</div>
                    </div>
                    
                    <div class="example">
                        💡 Example message to send: <code>auth:${code.substring(
                          0,
                          30
                        )}...</code>
                    </div>
                    
                    <div class="btn-container">
                        <button class="btn-copy" onclick="copyCode()">
                            📋 Copy Code with Prefix
                        </button>
                        <button class="btn-close" onclick="window.close()">
                            Close Window
                        </button>
                    </div>
                    
                    <div class="success-msg" id="successMsg">
                        ✓ Code copied! Now paste it in your bot chat.
                    </div>
                </div>
                
                <script>
                    function copyCode() {
                        const codeElement = document.getElementById('authCode');
                        const code = codeElement.textContent;
                        const fullCommand = 'auth:' + code;
                        
                        // Try modern clipboard API first
                        if (navigator.clipboard && navigator.clipboard.writeText) {
                            navigator.clipboard.writeText(fullCommand)
                                .then(() => {
                                    showSuccess();
                                })
                                .catch(err => {
                                    console.error('Clipboard API failed:', err);
                                    fallbackCopy(fullCommand);
                                });
                        } else {
                            fallbackCopy(fullCommand);
                        }
                    }
                    
                    function fallbackCopy(text) {
                        const textArea = document.createElement('textarea');
                        textArea.value = text;
                        textArea.style.position = 'fixed';
                        textArea.style.left = '-999999px';
                        textArea.style.top = '-999999px';
                        document.body.appendChild(textArea);
                        textArea.focus();
                        textArea.select();
                        
                        try {
                            const successful = document.execCommand('copy');
                            if (successful) {
                                showSuccess();
                            } else {
                                alert('Copy failed. Please manually copy: ' + text);
                            }
                        } catch (err) {
                            console.error('Fallback copy failed:', err);
                            alert('Please manually copy the code:\\n\\n' + text);
                        }
                        
                        document.body.removeChild(textArea);
                    }
                    
                    function showSuccess() {
                        const successMsg = document.getElementById('successMsg');
                        successMsg.style.display = 'block';
                        
                        setTimeout(() => {
                            successMsg.style.display = 'none';
                        }, 3000);
                    }
                    
                    // Auto-select code on page load for easy copying
                    window.addEventListener('load', () => {
                        const codeElement = document.getElementById('authCode');
                        const selection = window.getSelection();
                        const range = document.createRange();
                        range.selectNodeContents(codeElement);
                        selection.removeAllRanges();
                        selection.addRange(range);
                        
                        console.log('Auth code ready. State:', '${
                          state || "none"
                        }');
                    });
                </script>
            </body>
        </html>
    `;
}
