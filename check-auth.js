require("dotenv").config();

console.log("\n🔍 Checking OAuth Configuration\n");
console.log("═".repeat(70));

console.log("\n📋 Environment Variables:");
console.log(
  `   MICROSOFT_APP_ID: ${
    process.env.MICROSOFT_APP_ID ? "✅ Set" : "❌ Missing"
  }`
);
console.log(
  `   MICROSOFT_APP_PASSWORD: ${
    process.env.MICROSOFT_APP_PASSWORD ? "✅ Set" : "❌ Missing"
  }`
);
console.log(
  `   MICROSOFT_APP_TENANT_ID: ${
    process.env.MICROSOFT_APP_TENANT_ID || "common (default)"
  }`
);
console.log(
  `   OAUTH_REDIRECT_URI: ${process.env.OAUTH_REDIRECT_URI || "Not set"}`
);

if (!process.env.MICROSOFT_APP_ID || !process.env.MICROSOFT_APP_PASSWORD) {
  console.log("\n❌ Missing required credentials\n");
  process.exit(1);
}

console.log("\n🔐 OAuth Configuration:");
const tenantId = process.env.MICROSOFT_APP_TENANT_ID || "common";
const authorizeUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;
const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

console.log(`   Tenant: ${tenantId}`);
console.log(`   Authorize URL: ${authorizeUrl}`);
console.log(`   Token URL: ${tokenUrl}`);

console.log("\n📝 Required Permissions:");
console.log("   ✓ Calendars.ReadWrite (Delegated)");
console.log("   ✓ User.Read (Delegated)");
console.log("   ✓ offline_access (Delegated)");

console.log("\n📍 Redirect URI Configuration:");
console.log("   In .env file:");
console.log(`      ${process.env.OAUTH_REDIRECT_URI}`);
console.log("   Should match in Azure Portal exactly:");
console.log(`      ${process.env.OAUTH_REDIRECT_URI}`);

console.log("\n✅ Next Steps:");
console.log("   1. Go to Azure Portal");
console.log("   2. App registrations → Your app → API permissions");
console.log("   3. Verify all three permissions are listed");
console.log('   4. Click "Grant admin consent" if you see the button');
console.log("   5. In Bot Emulator, run: debug:clear");
console.log("   6. Re-authenticate with fresh login");

console.log("\n" + "═".repeat(70) + "\n");

// Try to get auth URL
const { AuthService } = require("./services/authService");
const authService = new AuthService();

if (authService.isConfigured) {
  console.log("🔗 Generating test auth URL...\n");

  authService
    .getAuthCodeUrl("test_" + Date.now())
    .then((url) => {
      console.log("   Test URL generated successfully!");
      console.log(`   Length: ${url.length} characters`);
      console.log(`\n   Click to test: ${url}\n`);
    })
    .catch((err) => {
      console.error("   ❌ Failed to generate URL:", err.message);
    });
} else {
  console.log("⚠️  Auth service not properly configured\n");
}
