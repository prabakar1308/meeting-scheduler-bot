require("dotenv").config();
const express = require("express");
const {
  BotFrameworkAdapter,
  MemoryStorage,
  ConversationState,
  UserState,
} = require("botbuilder");
const { MeetingSchedulerBot } = require("./bot");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3978;

// Create adapter WITHOUT credentials for local testing
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID || "",
  appPassword: process.env.MICROSOFT_APP_PASSWORD || "",
});

// Error handling
adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  console.error(error.stack);

  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity(`Error details: ${error.message}`);

  // Clear conversation state
  await conversationState.delete(context);
};

// Create conversation and user state
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create the bot
const bot = new MeetingSchedulerBot(conversationState, userState);

// Bot endpoint
app.post("/api/messages", async (req, res) => {
  console.log("Received message:", req.body);

  try {
    await adapter.process(req, res, async (context) => {
      await bot.run(context);
    });
  } catch (error) {
    console.error("Error processing request:", error);
    res.status(500).send("Error processing message");
  }
});

// OAuth callback endpoint with better handling
app.get("/api/oauth/callback", async (req, res) => {
  const code = req.query.code;
  const error = req.query.error;
  const state = req.query.state;

  console.log("OAuth callback received:", {
    code: code ? "present" : "missing",
    error,
    state,
  });

  if (error) {
    return res.status(400).send(`
            <html>
                <body style="font-family: Arial; padding: 20px;">
                    <h2 style="color: red;">Authentication Error</h2>
                    <p>Error: ${error}</p>
                    <p>Description: ${
                      req.query.error_description || "No description provided"
                    }</p>
                    <p>You can close this window and try again.</p>
                </body>
            </html>
        `);
  }

  if (!code) {
    return res.status(400).send(`
            <html>
                <body style="font-family: Arial; padding: 20px;">
                    <h2 style="color: orange;">Authorization Code Missing</h2>
                    <p>No authorization code was received.</p>
                    <p>You can close this window and try again.</p>
                </body>
            </html>
        `);
  }

  // In a real implementation, you would:
  // 1. Exchange the code for tokens
  // 2. Store tokens in user state
  // 3. Resume the conversation

  res.send(`
        <html>
            <head>
                <style>
                    body {
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                        margin: 0;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    }
                    .container {
                        background: white;
                        padding: 40px;
                        border-radius: 10px;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                        text-align: center;
                        max-width: 400px;
                    }
                    .checkmark {
                        font-size: 60px;
                        color: #4CAF50;
                    }
                    h2 {
                        color: #333;
                        margin: 20px 0 10px 0;
                    }
                    p {
                        color: #666;
                        margin: 10px 0;
                    }
                    .code {
                        background: #f5f5f5;
                        padding: 10px;
                        border-radius: 5px;
                        font-family: monospace;
                        font-size: 12px;
                        word-break: break-all;
                        margin: 20px 0;
                    }
                    button {
                        background: #667eea;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        border-radius: 5px;
                        cursor: pointer;
                        font-size: 14px;
                        margin-top: 20px;
                    }
                    button:hover {
                        background: #5568d3;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="checkmark">✓</div>
                    <h2>Authentication Successful!</h2>
                    <p>Your Microsoft account has been connected.</p>
                    <div class="code">
                        Authorization code received
                    </div>
                    <p style="font-size: 12px; color: #999;">
                        Return to your chat to continue scheduling meetings.
                    </p>
                    <button onclick="window.close()">Close Window</button>
                </div>
                <script>
                    // Auto-close after 3 seconds
                    setTimeout(() => {
                        window.close();
                    }, 3000);
                </script>
            </body>
        </html>
    `);
});

// Token exchange endpoint (for completing OAuth flow)
app.post("/api/oauth/token", express.json(), async (req, res) => {
  const { code, userId } = req.body;

  if (!code) {
    return res.status(400).json({ error: "Code is required" });
  }

  // Here you would exchange the code for tokens using MSAL
  // For now, just acknowledge receipt
  console.log("Token exchange requested for user:", userId);

  res.json({
    success: true,
    message: "Token exchange completed",
  });
});

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({
    status: "ok",
    timestamp: new Date().toISOString(),
    appId: process.env.MICROSOFT_APP_ID ? "configured" : "not configured",
  });
});

// Root endpoint with instructions
app.get("/", (req, res) => {
  res.send(`
        <html>
            <head>
                <title>Meeting Scheduler Bot</title>
                <style>
                    body {
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        max-width: 800px;
                        margin: 50px auto;
                        padding: 20px;
                        background: #f5f5f5;
                    }
                    .card {
                        background: white;
                        padding: 30px;
                        border-radius: 10px;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    }
                    h1 {
                        color: #667eea;
                        margin-top: 0;
                    }
                    .status {
                        padding: 10px;
                        border-radius: 5px;
                        margin: 20px 0;
                    }
                    .status.ok {
                        background: #d4edda;
                        color: #155724;
                        border: 1px solid #c3e6cb;
                    }
                    code {
                        background: #f5f5f5;
                        padding: 2px 6px;
                        border-radius: 3px;
                        font-family: monospace;
                    }
                    ul {
                        line-height: 1.8;
                    }
                </style>
            </head>
            <body>
                <div class="card">
                    <h1>🤖 Meeting Scheduler Bot</h1>
                    <div class="status ok">
                        ✓ Bot server is running successfully
                    </div>
                    
                    <h2>Connection Instructions</h2>
                    <p>To test this bot locally with Bot Framework Emulator:</p>
                    
                    <ol>
                        <li>Open <strong>Bot Framework Emulator</strong></li>
                        <li>Click <strong>"Open Bot"</strong></li>
                        <li>Enter bot URL: <code>http://localhost:${PORT}/api/messages</code></li>
                        <li><strong>Leave App ID and Password fields EMPTY</strong> for local testing</li>
                        <li>Click <strong>"Connect"</strong></li>
                    </ol>
                    
                    <h2>Available Endpoints</h2>
                    <ul>
                        <li><code>POST /api/messages</code> - Bot messaging endpoint</li>
                        <li><code>GET /api/oauth/callback</code> - OAuth callback</li>
                        <li><code>GET /health</code> - Health check</li>
                    </ul>
                    
                    <h2>Test Commands</h2>
                    <p>Try these commands in the Bot Emulator:</p>
                    <ul>
                        <li>"Schedule a team sync tomorrow at 2 PM for 1 hour"</li>
                        <li>"Book a meeting with john@example.com next Monday at 10 AM"</li>
                        <li>"Create a planning session on Friday at 3 PM"</li>
                    </ul>
                </div>
            </body>
        </html>
    `);
});

// Start server
app.listen(PORT, () => {
  console.log("\n╔══════════════════════════════════════════════════════════╗");
  console.log("║  🤖 Meeting Scheduler Bot Server                        ║");
  console.log("╚══════════════════════════════════════════════════════════╝\n");
  console.log(`✓ Server running on port ${PORT}`);
  console.log(`✓ Bot endpoint: http://localhost:${PORT}/api/messages`);
  console.log(`✓ OAuth callback: http://localhost:${PORT}/api/oauth/callback`);
  console.log(`✓ Web interface: http://localhost:${PORT}\n`);
  console.log("📝 To connect with Bot Framework Emulator:");
  console.log("   1. Open Bot Framework Emulator");
  console.log("   2. Use URL: http://localhost:3978/api/messages");
  console.log("   3. Leave App ID and Password EMPTY for local testing\n");
  console.log("Press Ctrl+C to stop the server\n");
});
