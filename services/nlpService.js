const { parse } = require("dotenv");

class NLPService {
  constructor() {
    // Simple patterns for POC (in production, use actual NLP libraries like compromise or natural)
    this.timePatterns = [
      /\bat\s+(\d{1,2})\s*(am|pm|AM|PM)/i, // "at 2 PM"
      /(\d{1,2})\s*(am|pm|AM|PM)/i, // "2 PM"
      /\bat\s+(\d{1,2}):(\d{2})\s*(am|pm|AM|PM)?/i, // "at 2:30 PM"
      /(\d{1,2}):(\d{2})\s*(am|pm|AM|PM)?/i, // "14:30" or "2:30 PM"
      /\bat\s+(\d{1,2}):(\d{2})/i, // "at 14:30"
      /(\d{1,2}):(\d{2})/, // "14:30"
    ];

    this.datePatterns = [
      { pattern: /\btomorrow\b/i, offset: 1 },
      { pattern: /\btoday\b/i, offset: 0 },
      {
        pattern:
          /\bnext\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i,
        type: "weekday",
      },
      {
        pattern:
          /\bon\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i,
        type: "weekday",
      },
      { pattern: /(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/, type: "date" },
      { pattern: /(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?/, type: "date" },
    ];

    this.durationPatterns = [
      /\bfor\s+(\d+)\s*hours?\b/i,
      /\bfor\s+(\d+)\s*h\b/i,
      /\bfor\s+(\d+)\s*mins?\b/i,
      /\bfor\s+(\d+)\s*minutes?\b/i,
      /(\d+)\s*hours?\b/i,
      /(\d+)\s*h\b/i,
      /(\d+)\s*mins?\b/i,
      /(\d+)\s*minutes?\b/i,
    ];

    this.emailPattern = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
  }

  async parseMeetingRequest(text) {
    console.log("\n🔍 NLP: Parsing meeting request...");
    console.log(`   Input: "${text}"`);

    try {
      // Extract meeting title (usually the first meaningful phrase)
      const title = this.extractTitle(text);
      console.log(`   Title: "${title}"`);

      // Extract date and time
      const dateTime = this.extractDateTime(text);
      console.log(`   DateTime: ${dateTime.toISOString()}`);

      // Validate datetime
      if (!dateTime || isNaN(dateTime.getTime())) {
        throw new Error("Could not parse valid date/time");
      }

      // Check if date is in the past
      const now = new Date();
      if (dateTime < now) {
        console.warn(
          "   ⚠️  Warning: Meeting time is in the past, adjusting..."
        );
        // If time is in the past, assume they meant tomorrow
        dateTime.setDate(dateTime.getDate() + 1);
        console.log(`   Adjusted to: ${dateTime.toISOString()}`);
      }

      // Extract duration
      const duration = this.extractDuration(text) || 60; // Default 60 minutes
      console.log(`   Duration: ${duration} minutes`);

      // Extract attendees
      const attendees = this.extractAttendees(text);
      console.log(
        `   Attendees: ${attendees.length > 0 ? attendees.join(", ") : "none"}`
      );

      // Extract location
      const location = this.extractLocation(text);
      if (location) {
        console.log(`   Location: "${location}"`);
      }

      if (!title) {
        throw new Error("Could not extract meeting title");
      }

      const result = {
        title,
        startTime: dateTime.toISOString(),
        durationMinutes: duration,
        attendees,
        location,
        body: `Meeting scheduled via AI Assistant\n\nOriginal request: ${text}`,
      };

      console.log("   ✅ Parsing complete");
      return result;
    } catch (error) {
      console.error("   ❌ Parsing failed:", error.message);
      throw error;
    }
  }

  extractTitle(text) {
    // Remove common phrases and extract the main subject
    let cleanText = text;

    // Try to find explicit title patterns first
    const titlePatterns = [
      /(?:schedule|create|book|set up)\s+(?:a\s+)?(?:meeting\s+)?(?:for\s+)?["']?([^"']+?)["']?(?:\s+(?:meeting|tomorrow|today|on|at|with|for|next))/i,
      /(?:meeting\s+(?:about|for|regarding))\s+([^,]+?)(?:\s+(?:tomorrow|today|on|at|with|for|next))/i,
      /(?:a\s+)?([a-z\s]+(?:meeting|sync|call|session|standup))(?:\s+(?:tomorrow|today|on|at|with|for|next))/i,
    ];

    for (const pattern of titlePatterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        let title = match[1].trim();
        // Clean up
        title = title.replace(/^(a|an|the)\s+/i, "");
        title = title.replace(/\s+meeting$/i, " Meeting");
        return title.charAt(0).toUpperCase() + title.slice(1);
      }
    }

    // Fallback: extract between keywords
    cleanText = text
      .replace(/^(schedule|create|set up|book)\s+/gi, "")
      .replace(/^(a|an|the)\s+/gi, "")
      .replace(/\s+(tomorrow|today|next week|on|at|for|with|from).*$/gi, "")
      .replace(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g, "")
      .trim();

    if (cleanText.length > 3 && cleanText.length < 100) {
      // Capitalize first letter
      cleanText = cleanText.charAt(0).toUpperCase() + cleanText.slice(1);
      // Ensure it ends with "Meeting" if it doesn't have a meaningful ending
      if (!/(meeting|call|sync|session|standup)$/i.test(cleanText)) {
        cleanText += " Meeting";
      }
      return cleanText;
    }

    // Last resort
    return "Team Meeting";
  }

  extractDateTime(text) {
    const now = new Date();
    let date = new Date(now);
    date.setHours(0, 0, 0, 0); // Reset to start of day

    let hours = 14; // Default 2 PM
    let minutes = 0;
    let timeFound = false;

    console.log("   📅 Extracting date/time from:", text);

    // Extract date first
    let dateFound = false;
    for (const datePattern of this.datePatterns) {
      if (datePattern.pattern.test(text)) {
        const match = text.match(datePattern.pattern);
        console.log(`   Matched date pattern: ${datePattern.pattern}`);

        if (datePattern.type === "weekday") {
          const weekday = match[1].toLowerCase();
          const targetDay = this.getWeekdayNumber(weekday);
          if (targetDay !== -1) {
            const currentDay = date.getDay();
            let daysToAdd = targetDay - currentDay;
            if (daysToAdd <= 0) daysToAdd += 7;
            date.setDate(date.getDate() + daysToAdd);
            dateFound = true;
            console.log(
              `   Found weekday: ${weekday}, adding ${daysToAdd} days`
            );
          }
        } else if (datePattern.type === "date") {
          const month = parseInt(match[1]) - 1;
          const day = parseInt(match[2]);
          const year = match[3] ? parseInt(match[3]) : date.getFullYear();
          date = new Date(year, month, day);
          dateFound = true;
          console.log(`   Found date: ${month + 1}/${day}/${year}`);
        } else if (datePattern.offset !== undefined) {
          date.setDate(date.getDate() + datePattern.offset);
          dateFound = true;
          console.log(
            `   Found relative date, offset: ${datePattern.offset} days`
          );
        }
        break;
      }
    }

    if (!dateFound) {
      console.log("   No specific date found, using today");
    }

    // Extract time
    for (const timePattern of this.timePatterns) {
      const match = text.match(timePattern);
      if (match) {
        console.log(`   Matched time pattern: ${timePattern}`);
        console.log(`   Time match groups:`, match);

        hours = parseInt(match[1]);
        if (match[2] && parseInt(match[2]) < 60) {
          minutes = parseInt(match[2]);
        } else minutes = 0;

        // Handle AM/PM
        const meridiem = match[2];
        if (meridiem) {
          const meridiemLower = meridiem.toLowerCase();
          if (meridiemLower === "pm" && hours < 12) {
            hours += 12;
          } else if (meridiemLower === "am" && hours === 12) {
            hours = 0;
          }
        }

        timeFound = true;
        console.log(
          `   Extracted time: ${hours}:${minutes} (meridiem: ${
            meridiem || "none"
          })`
        );
        break;
      }
    }

    if (!timeFound) {
      console.log(`   No time found, using default: ${hours}:${minutes}`);
    }

    // Set the time on the date
    date.setHours(hours, minutes, 0, 0);

    // Validate the resulting date
    if (isNaN(date.getTime())) {
      console.error("   ❌ Invalid date generated");
      // Fallback to tomorrow at 2 PM
      date = new Date();
      date.setDate(date.getDate() + 1);
      date.setHours(14, 0, 0, 0);
      console.log("   Using fallback date:", date.toISOString());
    } else {
      console.log(`   ✅ Final date/time: ${date.toISOString()}`);
    }

    return date;
  }

  extractDuration(text) {
    console.log("   ⏱️  Extracting duration...");

    for (const pattern of this.durationPatterns) {
      const match = text.match(pattern);
      if (match) {
        const value = parseInt(match[1]);
        console.log(`   Matched duration pattern: ${pattern}, value: ${value}`);

        // Check if pattern is for hours or minutes
        const patternStr = pattern.source.toLowerCase();
        const isHours =
          patternStr.includes("hour") || patternStr.includes("\\bh\\b");

        const minutes = isHours ? value * 60 : value;
        console.log(
          `   Duration: ${minutes} minutes (${
            isHours ? "hours" : "minutes"
          } input)`
        );
        return minutes;
      }
    }

    console.log("   No duration found, using default: 60 minutes");
    return null;
  }

  extractAttendees(text) {
    const emails = text.match(this.emailPattern);
    return emails || [];
  }

  extractLocation(text) {
    // Simple location extraction (look for "at" or "in" followed by location name)
    const locationMatch = text.match(
      /(?:at|in)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*(?:\s+Room)?)/
    );
    return locationMatch ? locationMatch[1] : null;
  }

  getWeekdayNumber(weekday) {
    const days = {
      sunday: 0,
      sun: 0,
      monday: 1,
      mon: 1,
      tuesday: 2,
      tue: 2,
      tues: 2,
      wednesday: 3,
      wed: 3,
      thursday: 4,
      thu: 4,
      thur: 4,
      thurs: 4,
      friday: 5,
      fri: 5,
      saturday: 6,
      sat: 6,
    };
    return days[weekday.toLowerCase()] ?? -1;
  }
}

module.exports.NLPService = NLPService;
