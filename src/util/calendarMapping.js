/**
 * Shared Calendar Mapping Utilities
 *
 * Provider-agnostic functions shared between Google Calendar and Outlook Calendar
 * mapping modules. Eliminates duplication across gcalMapping.js and outlookMapping.js.
 */

import { parseRange, getNormalizedTimestamp, strictTimestampRegex } from "./dates";
import { getBlockContentByUid } from "./roamApi";
import { uidRegex } from "./regex";

/**
 * Clean a Roam block title for an external calendar provider.
 * Removes Roam-specific syntax but preserves TODO/DONE based on user preference.
 *
 * @param {string} title - The Roam block title to clean
 * @param {object} options
 * @param {string[]} options.triggerTagsToRemove - Specific trigger tags to remove (null = remove all hashtags)
 * @param {string} options.defaultTag - The provider's default tag name (e.g., "Google calendar" or "Outlook calendar")
 * @param {Function} options.getCheckboxFormat - Function returning "bracket" or "roam"
 * @returns {string} Cleaned title
 */
export const cleanTitleForCalendar = (title, { triggerTagsToRemove = null, defaultTag, getCheckboxFormat }) => {
  if (!title) return "";

  let cleaned = title;

  // Remove bullet points at the beginning
  cleaned = cleaned.replace(/^[•\-]\s*/, "");

  // Get user's checkbox format preference
  const checkboxFormat = getCheckboxFormat();

  if (checkboxFormat === "bracket") {
    cleaned = cleaned.replace(/^\{\{\[\[TODO\]\]\}\}\s*/g, "[ ] ");
    cleaned = cleaned.replace(/^\{\{\[\[DONE\]\]\}\}\s*/g, "[x] ");
  } else {
    cleaned = cleaned.replace(/\{\{\[\[TODO\]\]\}\}/g, "[[TODO]]");
    cleaned = cleaned.replace(/\{\{\[\[DONE\]\]\}\}/g, "[[DONE]]");
  }

  // Protect backtick content
  const backtickContent = [];
  cleaned = cleaned.replace(/`([^`]+)`/g, (match, content) => {
    backtickContent.push(content);
    return `__BACKTICK_${backtickContent.length - 1}__`;
  });

  // Remove hashtags - either specific trigger tags or all hashtags
  if (triggerTagsToRemove !== null) {
    const tagsToRemove = [...new Set([...triggerTagsToRemove, defaultTag])];
    for (const tag of tagsToRemove) {
      if (!tag || !tag.trim()) continue;
      const escapedTag = tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      // Remove #[[tag]] format
      cleaned = cleaned.replace(new RegExp(`#\\[\\[${escapedTag}\\]\\]`, "gi"), "");
      // Remove #tag format (only if tag has no spaces)
      if (!tag.includes(" ")) {
        cleaned = cleaned.replace(new RegExp(`#${escapedTag}(?=\\s|$)`, "gi"), "");
      }
      // Also remove [[tag]] format (page reference style) for trigger tags
      cleaned = cleaned.replace(new RegExp(`\\[\\[${escapedTag}\\]\\]`, "gi"), "");
    }
  } else {
    cleaned = cleaned.replace(/#\[\[([^\]]+)\]\]/g, "");
    cleaned = cleaned.replace(/#([^\s]+)/g, "");
  }

  // Remove page references except [[TODO]] and [[DONE]]
  cleaned = cleaned.replace(/\[\[(?!TODO\]\]|DONE\]\])([^\]]+)\]\]/g, "$1");

  // Remove block embeds
  cleaned = cleaned.replace(/\{\{embed:\s*\(\([a-zA-Z0-9_-]+\)\)\}\}/g, "");

  // Remove other Roam syntax
  cleaned = cleaned.replace(/\{\{[^}]+\}\}/g, "");

  // Restore backtick content
  cleaned = cleaned.replace(/__BACKTICK_(\d+)__/g, (match, index) => {
    return `\`${backtickContent[parseInt(index)]}\``;
  });

  cleaned = cleaned.replace(/\s+/g, " ").trim();

  return cleaned || "(No title)";
};

/**
 * Convert [[TODO]], [[DONE]], [ ], or [x] in calendar title to Roam format.
 * Shared between Google Calendar and Outlook.
 *
 * @param {string} title - Calendar event title
 * @returns {string} Title with converted TODO/DONE syntax
 */
export const convertCalTodoToRoam = (title) => {
  if (!title) return title;
  let converted = title;

  converted = converted.replace(/^\[\[TODO\]\]\s*/g, "{{[[TODO]]}} ");
  converted = converted.replace(/^\[\[DONE\]\]\s*/g, "{{[[DONE]]}} ");
  converted = converted.replace(/^\[\s*\]\s*/g, "{{[[TODO]]}} ");
  converted = converted.replace(/^\[x\]\s*/g, "{{[[DONE]]}} ");

  return converted;
};

/**
 * Build Roam block content from a calendar event.
 * Shared logic for importing events from any calendar provider.
 *
 * @param {object} options
 * @param {string} options.title - Event title (subject)
 * @param {boolean} options.isAllDay - Whether the event is all-day
 * @param {Date|null} options.startDate - Parsed start date (for timed events)
 * @param {Date|null} options.endDate - Parsed end date (for timed events)
 * @param {object} options.calendarConfig - Calendar configuration
 * @param {boolean|null} options.hadOriginalTimeRange - Whether original had time range
 * @param {string} options.defaultTag - Default tag name (e.g., "Google calendar" or "Outlook calendar")
 * @returns {string} Roam block content
 */
export const buildRoamContentFromCalEvent = ({
  title: rawTitle,
  isAllDay,
  startDate,
  endDate,
  calendarConfig,
  hadOriginalTimeRange = null,
  defaultTag,
}) => {
  let content = "";

  let title = rawTitle || "(No title)";
  title = convertCalTodoToRoam(title);

  // Check if the title already contains a timestamp
  const titleHasTimeRange = parseRange(title) !== null;
  const titleHasTimestamp =
    titleHasTimeRange || getNormalizedTimestamp(title, strictTimestampRegex) !== null;

  // Add time for timed (non-all-day) events if title doesn't already have one
  if (!isAllDay && startDate && !titleHasTimestamp) {
    const hours = startDate.getHours();
    const minutes = startDate.getMinutes();
    const timeStr = `${hours}:${String(minutes).padStart(2, "0")}`;

    const shouldIncludeEndTime =
      hadOriginalTimeRange === true || hadOriginalTimeRange === null;

    if (shouldIncludeEndTime && endDate) {
      const endHours = endDate.getHours();
      const endMinutes = endDate.getMinutes();
      const endTimeStr = `${endHours}:${String(endMinutes).padStart(2, "0")}`;

      const durationMs = endDate.getTime() - startDate.getTime();
      const isDefaultDuration = durationMs === 3600000; // 1 hour

      if (
        hadOriginalTimeRange === true ||
        (hadOriginalTimeRange === null && !isDefaultDuration)
      ) {
        content += `${timeStr}-${endTimeStr} `;
      } else {
        content += `${timeStr} `;
      }
    } else {
      content += `${timeStr} `;
    }
  }

  content += title;

  // Add trigger tag
  const customTag = calendarConfig.triggerTags?.[0]?.trim();
  const tagToAdd = customTag || defaultTag;
  content += tagToAdd.includes(" ") ? ` #[[${tagToAdd}]]` : ` #${tagToAdd}`;

  return content;
};

/**
 * Extract block references from a title and resolve their content.
 * Used when syncing from Roam to an external calendar to build a
 * "block references legend" in the description.
 *
 * @param {string} title - Roam block title
 * @returns {Array<{ref: string, content: string}>} Resolved block references
 */
export const extractBlockReferences = (title) => {
  if (!title) return [];

  const blockRefLegend = [];
  uidRegex.lastIndex = 0;
  const matches = Array.from(title.matchAll(uidRegex));
  for (const match of matches) {
    const refUid = match[0].slice(2, -2);
    const resolvedContent = getBlockContentByUid(refUid);
    if (resolvedContent) {
      blockRefLegend.push({ ref: match[0], content: resolvedContent });
    }
  }

  return blockRefLegend;
};

/**
 * Build description text for an external calendar event from Roam data.
 * Handles block reference legends and Roam block links.
 *
 * @param {string} existingDescription - Existing description text
 * @param {Array<{ref: string, content: string}>} blockRefLegend - Block references
 * @param {string|null} roamUid - Roam block UID (for adding link)
 * @param {Function} cleanTitleFn - Provider-specific title cleaning function
 * @returns {string} Built description
 */
export const buildCalendarDescription = (existingDescription, blockRefLegend, roamUid, cleanTitleFn) => {
  let description = existingDescription || "";

  // Remove old block references section and Roam link
  description = description
    .replace(/\n*---\nBlock references:[\s\S]*?(?=\n---\nRoam block:|$)/s, "")
    .trim();
  description = description.replace(/\n*---\nRoam block:.*$/s, "").trim();

  // Add block references legend
  if (blockRefLegend.length > 0) {
    description += "\n\n---\nBlock references:";
    for (const { ref, content } of blockRefLegend) {
      const cleanedContent = cleanTitleFn(content);
      description += `\n${ref} = ${cleanedContent}`;
    }
  }

  // Add Roam block link
  if (roamUid) {
    const graphName = window.roamAlphaAPI?.graph?.name;
    if (graphName) {
      const roamLink = `https://roamresearch.com/#/app/${graphName}/page/${roamUid}`;
      description += `\n\n---\nRoam block: ${roamLink}`;
    }
  }

  return description;
};

/**
 * Determine if an FC event should be synced based on its tags.
 * Logic is identical between Google Calendar and Outlook.
 *
 * @param {object} fcEvent - FullCalendar event
 * @param {array} connectedCalendars - Array of connected calendar configs
 * @returns {object|null} Calendar config to sync to, or null if no match
 */
export const findCalendarForEvent = (fcEvent, connectedCalendars) => {
  const eventTags = fcEvent.extendedProps?.eventTags || [];

  for (const calendar of connectedCalendars) {
    if (!calendar.syncEnabled) continue;
    if (calendar.syncDirection === "import") continue;

    for (const eventTag of eventTags) {
      const tagName = eventTag.name?.toLowerCase();

      // Check displayName first
      if (calendar.displayName && calendar.displayName.toLowerCase() === tagName) {
        return calendar;
      }

      // Check trigger tags
      if (calendar.triggerTags && calendar.triggerTags.length > 0) {
        if (calendar.triggerTags.some((trigger) => trigger.toLowerCase() === tagName)) {
          return calendar;
        }

        // Check tag pages/aliases
        if (eventTag.pages && Array.isArray(eventTag.pages)) {
          for (const page of eventTag.pages) {
            if (calendar.triggerTags.some((trigger) => trigger.toLowerCase() === page.toLowerCase())) {
              return calendar;
            }
          }
        }
      }
    }
  }

  return null;
};

/**
 * Check if an event has any sync trigger tags.
 *
 * @param {object} fcEvent - FullCalendar event
 * @param {array} connectedCalendars - Array of connected calendar configs
 * @returns {boolean}
 */
export const hasSyncTriggerTag = (fcEvent, connectedCalendars) => {
  return findCalendarForEvent(fcEvent, connectedCalendars) !== null;
};

/**
 * Parse HTML description from a calendar provider into an array of Roam block contents.
 * Shared between Google Calendar and Outlook.
 *
 * @param {string} htmlDescription - HTML or text description
 * @returns {string[]} Array of block contents
 */
export const parseCalDescriptionToBlocks = (htmlDescription) => {
  if (!htmlDescription) return [];

  let text = htmlDescription;

  // Remove Roam link section
  text = text.replace(/\n*---\n*Roam block:.*$/s, "").trim();
  text = text.replace(/\n*---\n*Block references:[\s\S]*?(?=\n---\n|$)/s, "").trim();

  // Convert HTML to text
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/p>/gi, "\n");
  text = text.replace(/<\/div>/gi, "\n");
  text = text.replace(/<li[^>]*>/gi, "\n• ");
  text = text.replace(/<\/li>/gi, "");
  text = text.replace(/<a\s+[^>]*href=["']([^"']+)["'][^>]*>([^<]*)<\/a>/gi, "[$2]($1)");
  text = text.replace(/<[^>]*>/g, "");

  // Decode HTML entities
  text = text.replace(/&nbsp;/g, " ");
  text = text.replace(/&amp;/g, "&");
  text = text.replace(/&lt;/g, "<");
  text = text.replace(/&gt;/g, ">");
  text = text.replace(/&quot;/g, '"');
  text = text.replace(/&#39;/g, "'");
  text = text.replace(/&apos;/g, "'");

  const lines = text
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  return lines;
};

/**
 * Parse calendar event metadata into Roam child blocks.
 * Handles Location and Attendees which are common to both providers.
 *
 * @param {object} options
 * @param {string} options.location - Event location
 * @param {Array} options.attendees - Array of attendee objects
 * @param {Function} options.getAttendeeName - Function to extract display name from attendee object
 * @param {Array} [options.attachments] - Optional array of attachment objects (GCal-specific)
 * @param {Function} [options.formatAttachment] - Optional function to format an attachment
 * @returns {string[]} Array of block contents for metadata
 */
export const parseCalMetadataToBlocks = ({ location, attendees, getAttendeeName, attachments, formatAttachment }) => {
  const blocks = [];

  if (location) {
    blocks.push(`Location:: ${location}`);
  }

  if (attendees && attendees.length > 0) {
    const attendeesList = attendees
      .map((attendee) => `[[${getAttendeeName(attendee)}]]`)
      .join(", ");
    blocks.push(`Attendees:: ${attendeesList}`);
  }

  if (attachments && attachments.length > 0 && formatAttachment) {
    const attachmentLinks = attachments.map(formatAttachment).join(", ");
    blocks.push(`Attachments:: ${attachmentLinks}`);
  }

  return blocks;
};

/**
 * Validate and normalize a start date, with fallback.
 *
 * @param {*} startInput - Raw start date (Date, string, etc.)
 * @param {*} fallbackDate - Fallback date source (e.g., fcEvent.date)
 * @returns {Date} Valid Date object
 */
export const normalizeStartDate = (startInput, fallbackDate = null) => {
  let startDate = startInput;
  if (!(startDate instanceof Date)) {
    startDate = new Date(startDate);
  }
  if (isNaN(startDate.getTime())) {
    if (fallbackDate) {
      startDate = new Date(fallbackDate);
    }
    if (isNaN(startDate.getTime())) {
      startDate = new Date();
    }
  }
  return startDate;
};

/**
 * Validate and normalize an end date.
 *
 * @param {*} endInput - Raw end date
 * @returns {Date|null} Valid Date object or null
 */
export const normalizeEndDate = (endInput) => {
  if (!endInput) return null;
  let endDate = endInput;
  if (!(endDate instanceof Date)) {
    endDate = new Date(endDate);
  }
  if (isNaN(endDate.getTime())) {
    return null;
  }
  return endDate;
};

/**
 * Build default end date/time when no end is specified.
 *
 * @param {Date} startDate - Event start date
 * @param {boolean} isAllDay - Whether the event is all-day
 * @returns {Date} Default end date
 */
export const buildDefaultEndDate = (startDate, isAllDay) => {
  const defaultEnd = new Date(startDate);
  if (isAllDay) {
    defaultEnd.setDate(defaultEnd.getDate() + 1);
  } else {
    defaultEnd.setHours(defaultEnd.getHours() + 1);
  }
  return defaultEnd;
};
