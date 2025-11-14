const {
  WaterfallDialog,
  ComponentDialog,
  TextPrompt,
  ConfirmPrompt,
  DialogTurnStatus,
} = require("botbuilder-dialogs");
const { GraphService } = require("../services/graphService");
const { NLPService } = require("../services/nlpService");
const { AuthService } = require("../services/authService");

const TEXT_PROMPT = "TextPrompt";
const CONFIRM_PROMPT = "ConfirmPrompt";
const WATERFALL_DIALOG = "WaterfallDialog";

class ScheduleDialog extends ComponentDialog {
  constructor(id, userProfileAccessor) {
    super(id);

    this.userProfileAccessor = userProfileAccessor;
    this.graphService = new GraphService();
    this.nlpService = new NLPService();
    this.authService = new AuthService();

    this.addDialog(new TextPrompt(TEXT_PROMPT));
    this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
    this.addDialog(
      new WaterfallDialog(WATERFALL_DIALOG, [
        this.gatherMeetingDetails.bind(this),
        this.parseMeetingDetails.bind(this),
        this.confirmMeeting.bind(this),
        this.scheduleMeeting.bind(this),
      ])
    );

    this.initialDialogId = WATERFALL_DIALOG;
  }

  async gatherMeetingDetails(stepContext) {
    stepContext.values.meetingDetails = {};

    // Check if user already provided details in their initial message
    const userMessage = stepContext.context.activity.text;

    console.log("\n📝 Gathering meeting details...");
    console.log(`   User message: "${userMessage}"`);

    // Try to parse the initial message
    if (userMessage && this.looksLikeMeetingRequest(userMessage)) {
      console.log(
        "   ✅ Message looks like a meeting request, trying to parse..."
      );

      try {
        const meetingDetails = await this.nlpService.parseMeetingRequest(
          userMessage
        );
        console.log(
          "   ✅ Parsed successfully:",
          JSON.stringify(meetingDetails, null, 2)
        );

        // Store the parsed details and skip to confirmation
        stepContext.values.meetingDetails = meetingDetails;
        return await stepContext.next(userMessage); // Pass to next step
      } catch (error) {
        console.log("   ⚠️  Could not parse automatically:", error.message);
        // Fall through to prompt
      }
    }

    // If we couldn't parse or no details provided, prompt the user
    return await stepContext.prompt(TEXT_PROMPT, {
      prompt:
        "Please describe the meeting you want to schedule. Include:\n" +
        "• Meeting title\n" +
        "• Date and time\n" +
        "• Duration\n" +
        "• Attendees (optional)\n\n" +
        'Example: "Schedule a team sync meeting tomorrow at 2 PM for 1 hour with john@example.com"',
    });
  }

  async parseMeetingDetails(stepContext) {
    // If we already have meeting details from step 1, skip parsing
    if (
      stepContext.values.meetingDetails &&
      stepContext.values.meetingDetails.title
    ) {
      console.log("   ℹ️  Using pre-parsed meeting details");
      const confirmMessage = this.formatMeetingDetails(
        stepContext.values.meetingDetails
      );
      return await stepContext.prompt(CONFIRM_PROMPT, {
        prompt: `I've extracted the following meeting details:\n\n${confirmMessage}\n\nWould you like me to schedule this meeting?`,
      });
    }

    // Otherwise, parse the user's response
    const userInput = stepContext.result;

    console.log("\n🔍 Parsing meeting details...");
    console.log(`   Input: "${userInput}"`);

    try {
      // Parse the natural language input
      const meetingDetails = await this.nlpService.parseMeetingRequest(
        userInput
      );
      stepContext.values.meetingDetails = meetingDetails;

      console.log("   ✅ Parsed successfully");
      console.log(JSON.stringify(meetingDetails, null, 2));

      // Format details for confirmation
      const confirmMessage = this.formatMeetingDetails(meetingDetails);

      return await stepContext.prompt(CONFIRM_PROMPT, {
        prompt: `I've extracted the following meeting details:\n\n${confirmMessage}\n\nWould you like me to schedule this meeting?`,
      });
    } catch (error) {
      console.error("   ❌ Failed to parse:", error.message);
      await stepContext.context.sendActivity(
        "❌ I couldn't understand the meeting details. Please try again with more specific information.\n\n" +
          "Make sure to include:\n" +
          "• Meeting title or purpose\n" +
          "• Date (today, tomorrow, or specific date)\n" +
          "• Time (e.g., 2 PM, 14:00)\n\n" +
          'Example: "Team standup tomorrow at 9 AM for 30 minutes"'
      );
      return await stepContext.endDialog();
    }
  }

  async confirmMeeting(stepContext) {
    if (stepContext.result) {
      console.log("   ✅ User confirmed, proceeding to schedule...");
      return await stepContext.next();
    } else {
      console.log("   ❌ User cancelled");
      await stepContext.context.sendActivity(
        "Meeting cancelled. Let me know when you want to schedule something else!"
      );
      return await stepContext.endDialog();
    }
  }

  async scheduleMeeting(stepContext) {
    const meetingDetails = stepContext.values.meetingDetails;

    console.log("\n📅 Scheduling meeting...");
    console.log(JSON.stringify(meetingDetails, null, 2));

    try {
      // Get user profile with access token
      const userProfile = await this.userProfileAccessor.get(
        stepContext.context,
        {}
      );

      if (!userProfile.accessToken) {
        console.error("   ❌ No access token found");
        await stepContext.context.sendActivity(
          "❌ Your session has expired. Please authenticate again."
        );
        return await stepContext.endDialog();
      }

      console.log(
        "   📤 Creating meeting via Graph API...",
        userProfile.accessToken
      );

      const accessToken = await this.authService.getAccessToken(userProfile);

      // Create the meeting using Graph API
      const meeting = await this.graphService.createMeeting(
        accessToken,
        meetingDetails
      );

      console.log("   ✅ Meeting created successfully");
      console.log(`   Meeting ID: ${meeting.id}`);
      console.log(`   Web link: ${meeting.webLink}`);

      // Format success message
      const startTime = new Date(meetingDetails.startTime);
      const attendeesList =
        meetingDetails.attendees && meetingDetails.attendees.length > 0
          ? `\n👥 Attendees: ${meetingDetails.attendees.join(", ")}`
          : "";

      await stepContext.context.sendActivity(
        `✅ **Meeting Scheduled Successfully!**\n\n` +
          `📋 **${meetingDetails.title}**\n` +
          `📅 ${startTime.toLocaleDateString()} at ${startTime.toLocaleTimeString()}\n` +
          `⏱️ Duration: ${meetingDetails.durationMinutes} minutes${attendeesList}\n\n` +
          `🔗 [View in Outlook](${meeting.webLink})`
      );

      // Ask if they want to schedule another
      await stepContext.context.sendActivity(
        "Would you like to schedule another meeting? Just describe it!"
      );
    } catch (error) {
      console.error("   ❌ Error scheduling meeting:", error);
      console.error("   Error details:", error.response?.data || error.message);

      await stepContext.context.sendActivity(
        `❌ **Failed to schedule meeting**\n\n` +
          `Error: ${error.message}\n\n` +
          "Please check:\n" +
          "• Your calendar permissions are granted\n" +
          "• The meeting details are valid\n" +
          "• You have an active internet connection"
      );
    }

    return await stepContext.endDialog();
  }

  looksLikeMeetingRequest(text) {
    if (!text) return false;

    const lowerText = text.toLowerCase();

    // Check for scheduling keywords
    const hasScheduleKeyword =
      /\b(schedule|create|book|set up|plan|arrange)\b/i.test(lowerText);
    const hasMeetingKeyword =
      /\b(meeting|call|sync|session|appointment|event)\b/i.test(lowerText);
    const hasTimeReference =
      /\b(today|tomorrow|next|at|am|pm|:\d{2}|\d{1,2}\s*(am|pm))\b/i.test(
        lowerText
      );

    // Should have at least schedule/meeting keywords and time reference
    return (hasScheduleKeyword || hasMeetingKeyword) && hasTimeReference;
  }

  formatMeetingDetails(details) {
    let formatted = `**📋 Title:** ${details.title}\n`;

    const startTime = new Date(details.startTime);
    formatted += `**📅 Date:** ${startTime.toLocaleDateString()}\n`;
    formatted += `**🕐 Time:** ${startTime.toLocaleTimeString()}\n`;
    formatted += `**⏱️ Duration:** ${details.durationMinutes} minutes\n`;

    if (details.attendees && details.attendees.length > 0) {
      formatted += `**👥 Attendees:** ${details.attendees.join(", ")}\n`;
    }

    if (details.location) {
      formatted += `**📍 Location:** ${details.location}\n`;
    }

    return formatted;
  }
}

module.exports.ScheduleDialog = ScheduleDialog;
