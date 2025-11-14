// services/meetingIntelligence.js
// Add this to enhance your bot with AI-powered features

const { GraphService } = require("./graphService");

class MeetingIntelligence {
  constructor() {
    this.graphService = new GraphService();
  }

  /**
   * Find optimal meeting time based on attendee availability
   */
  async suggestOptimalTime(accessToken, attendees, duration, preferredDate) {
    const startDate = new Date(preferredDate);
    startDate.setHours(0, 0, 0, 0);

    const endDate = new Date(preferredDate);
    endDate.setHours(23, 59, 59, 999);

    try {
      // Get calendar view for all attendees
      const schedules = await this.graphService.findAvailableSlots(
        accessToken,
        attendees,
        duration,
        preferredDate
      );

      if (!schedules) {
        return this.getDefaultTimeSlots(preferredDate);
      }

      // Analyze availability and suggest best times
      const suggestions = this.analyzeAvailability(schedules, duration);
      return suggestions;
    } catch (error) {
      console.error("Error finding optimal time:", error);
      return this.getDefaultTimeSlots(preferredDate);
    }
  }

  /**
   * Analyze schedules to find best meeting times
   */
  analyzeAvailability(schedules, durationMinutes) {
    const suggestions = [];
    const workingHours = this.getWorkingHours();

    // Find common free slots
    for (let hour = workingHours.start; hour < workingHours.end; hour++) {
      for (let minute = 0; minute < 60; minute += 30) {
        const slotStart = new Date();
        slotStart.setHours(hour, minute, 0, 0);

        const slotEnd = new Date(slotStart);
        slotEnd.setMinutes(slotEnd.getMinutes() + durationMinutes);

        // Check if all attendees are free
        const isFree = this.isSlotAvailable(schedules, slotStart, slotEnd);

        if (isFree) {
          suggestions.push({
            startTime: slotStart.toISOString(),
            endTime: slotEnd.toISOString(),
            confidence: this.calculateConfidence(hour, minute),
          });
        }

        if (suggestions.length >= 3) break;
      }
      if (suggestions.length >= 3) break;
    }

    return suggestions;
  }

  /**
   * Check if time slot is available for all attendees
   */
  isSlotAvailable(schedules, slotStart, slotEnd) {
    for (const schedule of schedules) {
      const busyPeriods = schedule.scheduleItems || [];

      for (const period of busyPeriods) {
        const periodStart = new Date(period.start.dateTime);
        const periodEnd = new Date(period.end.dateTime);

        // Check for overlap
        if (slotStart < periodEnd && slotEnd > periodStart) {
          return false;
        }
      }
    }
    return true;
  }

  /**
   * Calculate confidence score for suggested time
   */
  calculateConfidence(hour, minute) {
    // Prefer times between 10 AM - 4 PM
    let confidence = 100;

    if (hour < 10 || hour >= 16) confidence -= 20;
    if (hour === 12 || hour === 13) confidence -= 10; // Lunch time
    if (minute === 0) confidence += 5; // On the hour

    return Math.max(0, Math.min(100, confidence));
  }

  /**
   * Get default working hours
   */
  getWorkingHours() {
    return { start: 9, end: 17 }; // 9 AM to 5 PM
  }

  /**
   * Get default time slot suggestions
   */
  getDefaultTimeSlots(date) {
    const suggestions = [];
    const times = [
      { hour: 10, minute: 0 },
      { hour: 14, minute: 0 },
      { hour: 15, minute: 30 },
    ];

    times.forEach((time) => {
      const slotStart = new Date(date);
      slotStart.setHours(time.hour, time.minute, 0, 0);

      suggestions.push({
        startTime: slotStart.toISOString(),
        confidence: 80,
      });
    });

    return suggestions;
  }

  /**
   * Detect meeting conflicts
   */
  async detectConflicts(accessToken, startTime, endTime) {
    try {
      const events = await this.graphService.getCalendarEvents(
        accessToken,
        new Date(startTime),
        new Date(endTime)
      );

      return events.filter((event) => {
        const eventStart = new Date(event.start.dateTime);
        const eventEnd = new Date(event.end.dateTime);
        const proposedStart = new Date(startTime);
        const proposedEnd = new Date(endTime);

        return proposedStart < eventEnd && proposedEnd > eventStart;
      });
    } catch (error) {
      console.error("Error detecting conflicts:", error);
      return [];
    }
  }

  /**
   * Generate smart meeting title suggestions
   */
  generateTitleSuggestions(context, attendees) {
    const suggestions = [];

    // Based on attendees count
    if (attendees.length === 0) {
      suggestions.push("Personal Time Block", "Focus Session", "Planning Time");
    } else if (attendees.length === 1) {
      suggestions.push("1:1 Meeting", "Sync", "Check-in");
    } else if (attendees.length <= 5) {
      suggestions.push(
        "Team Sync",
        "Project Discussion",
        "Collaboration Session"
      );
    } else {
      suggestions.push("Team Meeting", "All Hands", "Group Discussion");
    }

    return suggestions;
  }

  /**
   * Estimate meeting duration based on context
   */
  estimateDuration(title, attendees) {
    const titleLower = title.toLowerCase();

    // Quick meetings
    if (titleLower.includes("standup") || titleLower.includes("quick")) {
      return 15;
    }

    // 1:1 meetings
    if (titleLower.includes("1:1") || titleLower.includes("one-on-one")) {
      return 30;
    }

    // Reviews or planning
    if (titleLower.includes("review") || titleLower.includes("planning")) {
      return 60;
    }

    // Default based on attendee count
    if (attendees.length <= 2) return 30;
    if (attendees.length <= 5) return 45;
    return 60;
  }

  /**
   * Format time suggestions for user display
   */
  formatTimeSuggestions(suggestions) {
    if (!suggestions || suggestions.length === 0) {
      return "No available time slots found.";
    }

    let formatted = "🕐 **Suggested meeting times:**\n\n";

    suggestions.forEach((suggestion, index) => {
      const startTime = new Date(suggestion.startTime);
      const confidence = suggestion.confidence || 0;
      const emoji = confidence >= 90 ? "✅" : confidence >= 70 ? "👍" : "📅";

      formatted += `${emoji} Option ${
        index + 1
      }: ${startTime.toLocaleString()} `;
      formatted += `(${confidence}% optimal)\n`;
    });

    return formatted;
  }
}

module.exports.MeetingIntelligence = MeetingIntelligence;
