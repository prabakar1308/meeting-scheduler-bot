const msal = require("@azure/msal-node");
const axios = require("axios");

class AuthService {
  constructor() {
    this.msalConfig = {
      auth: {
        clientId: process.env.MICROSOFT_APP_ID,
        authority: `https://login.microsoftonline.com/${
          process.env.MICROSOFT_APP_TENANT_ID || "common"
        }`,
        clientSecret: process.env.MICROSOFT_APP_PASSWORD,
      },
    };

    this.pca = new msal.ConfidentialClientApplication(this.msalConfig);
    this.redirectUri = process.env.OAUTH_REDIRECT_URI;
    this.scopes = [
      "Calendars.ReadWrite",
      "User.Read",
      "offline_access",
      "Mail.Send",
    ];
  }

  getAuthCodeUrl() {
    const authCodeUrlParameters = {
      scopes: this.scopes,
      redirectUri: this.redirectUri,
    };

    return this.pca.getAuthCodeUrl(authCodeUrlParameters);
  }

  async acquireTokenByCode(code) {
    const tokenRequest = {
      code: code,
      scopes: this.scopes,
      redirectUri: this.redirectUri,
    };

    try {
      const response = await this.pca.acquireTokenByCode(tokenRequest);
      return {
        accessToken: response.accessToken,
        refreshToken: response.refreshToken,
        expiresOn: response.expiresOn,
      };
    } catch (error) {
      console.error("Error acquiring token:", error);
      throw error;
    }
  }

  async acquireTokenByRefreshToken(refreshToken) {
    const refreshTokenRequest = {
      refreshToken: refreshToken,
      scopes: this.scopes,
    };

    try {
      const response = await this.pca.acquireTokenByRefreshToken(
        refreshTokenRequest
      );
      return {
        accessToken: response.accessToken,
        refreshToken: response.refreshToken || refreshToken,
        expiresOn: response.expiresOn,
      };
    } catch (error) {
      console.error("Error refreshing token:", error);
      throw error;
    }
  }

  isTokenExpired(expiresOn) {
    if (!expiresOn) return true;
    const now = new Date();
    const expiry = new Date(expiresOn);
    // Add 5 minute buffer
    return expiry.getTime() - now.getTime() < 5 * 60 * 1000;
  }

  async ensureValidToken(userProfile) {
    if (!userProfile.accessToken) {
      throw new Error("No access token available");
    }

    if (this.isTokenExpired(userProfile.expiresOn)) {
      if (!userProfile.refreshToken) {
        throw new Error("Token expired and no refresh token available");
      }

      const newTokens = await this.acquireTokenByRefreshToken(
        userProfile.refreshToken
      );
      return newTokens;
    }

    return {
      accessToken: userProfile.accessToken,
      refreshToken: userProfile.refreshToken,
      expiresOn: userProfile.expiresOn,
    };
  }

  async getAccessToken(userProfile) {
    if (userProfile) {
      // Ensure the token is valid for the user
      const validToken = await this.ensureValidToken(userProfile);
      return validToken.accessToken;
    }

    // Fallback to application token if no user context is provided
    const result = await this.pca.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    return result.accessToken;
  }
}

module.exports.AuthService = AuthService;
