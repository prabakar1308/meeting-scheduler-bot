const axios = require("axios");

class GraphService {
  constructor() {
    this.graphBaseUrl = "https://graph.microsoft.com/v1.0";
  }

  async createMeeting(accessToken, meetingDetails) {
    const { title, startTime, durationMinutes, attendees, location, body } =
      meetingDetails;

    // Calculate end time
    const endTime = new Date(
      new Date(startTime).getTime() + durationMinutes * 60000
    );

    // Prepare event object
    const event = {
      subject: title,
      body: {
        contentType: "HTML",
        content: body || "Meeting scheduled via Bot",
      },
      start: {
        dateTime: startTime,
        timeZone: "UTC",
      },
      end: {
        dateTime: endTime.toISOString(),
        timeZone: "UTC",
      },
      location: location
        ? {
            displayName: location,
          }
        : undefined,
      attendees:
        attendees && attendees.length > 0
          ? attendees.map((email) => ({
              emailAddress: {
                address: email,
                name: email.split("@")[0],
              },
              type: "required",
            }))
          : [],
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
    };

    try {
      const response = await axios.post(
        `${this.graphBaseUrl}/me/events?sendUpdates=all`,
        event,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      return response.data;
    } catch (error) {
      console.error("Graph API Error:", error.response?.data || error.message);
      throw new Error(
        error.response?.data?.error?.message || "Failed to create meeting"
      );
    }
  }

  async getCalendarEvents(accessToken, startDate, endDate) {
    try {
      const response = await axios.get(`${this.graphBaseUrl}/me/calendarview`, {
        params: {
          startDateTime: startDate.toISOString(),
          endDateTime: endDate.toISOString(),
        },
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      return response.data.value;
    } catch (error) {
      console.error("Graph API Error:", error.response?.data || error.message);
      throw new Error("Failed to retrieve calendar events");
    }
  }

  async findAvailableSlots(accessToken, attendees, duration, preferredDate) {
    const startTime = new Date(preferredDate);
    startTime.setHours(9, 0, 0, 0); // Start at 9 AM

    const endTime = new Date(preferredDate);
    endTime.setHours(17, 0, 0, 0); // End at 5 PM

    const requestBody = {
      schedules: ["your-email@outlook.com", ...attendees],
      startTime: {
        dateTime: startTime.toISOString(),
        timeZone: "UTC",
      },
      endTime: {
        dateTime: endTime.toISOString(),
        timeZone: "UTC",
      },
      availabilityViewInterval: duration,
    };

    try {
      const response = await axios.post(
        `${this.graphBaseUrl}/me/calendar/getSchedule`,
        requestBody,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      return response.data.value;
    } catch (error) {
      console.error("Graph API Error:", error.response?.data || error.message);
      return null;
    }
  }

  async getUserProfile(accessToken) {
    try {
      const response = await axios.get(`${this.graphBaseUrl}/me`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      return response.data;
    } catch (error) {
      console.error("Graph API Error:", error.response?.data || error.message);
      throw new Error("Failed to get user profile");
    }
  }

  async validateToken(accessToken) {
    console.log("\n🔍 Validating access token...");
    console.log(`   Token length: ${accessToken?.length || 0}`);

    try {
      const response = await axios.get(`${this.graphBaseUrl}/me`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      console.log("   ✅ Token is valid");
      console.log(
        `   User: ${response.data.displayName} (${
          response.data.mail || response.data.userPrincipalName
        })`
      );

      return {
        valid: true,
        user: response.data,
      };
    } catch (error) {
      console.error("   ❌ Token validation failed");
      console.error("   Status:", error.response?.status);
      console.error("   Error:", error.response?.data?.error?.message);

      return {
        valid: false,
        error: error.response?.data?.error?.message || error.message,
      };
    }
  }

  async testCalendarAccess(accessToken) {
    console.log("\n🔍 Testing calendar access...");

    try {
      const response = await axios.get(`${this.graphBaseUrl}/me/calendars`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      console.log("   ✅ Calendar access granted");
      console.log(`   Calendars found: ${response.data.value.length}`);

      return {
        hasAccess: true,
        calendars: response.data.value,
      };
    } catch (error) {
      console.error("   ❌ Calendar access failed");
      console.error("   Status:", error.response?.status);
      console.error("   Error:", error.response?.data?.error?.message);

      return {
        hasAccess: false,
        error: error.response?.data?.error?.message || error.message,
      };
    }
  }
}

module.exports.GraphService = GraphService;
