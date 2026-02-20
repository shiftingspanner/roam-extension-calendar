/**
 * Outlook Calendar Event Mapping Utilities
 *
 * Handles conversion between Microsoft Graph events and FullCalendar events
 */

import { DateTime } from "luxon";
import { getTagFromName } from "../models/EventTag";
import { OutlookSyncStatus } from "../models/OutlookSyncMetadata";
import { parseRange, getNormalizedTimestamp, strictTimestampRegex } from "./dates";
import {
  getOutlookUseOriginalColors,
  getOutlookCheckboxFormat,
  getOutlookConnectedCalendars,
} from "../services/outlookCalendarService";
import { getBlockContentByUid } from "./roamApi";
import { uidRegex } from "./regex";

// Outlook category name to hex color mapping
const OUTLOOK_CATEGORY_COLORS = {
  "Red category": "#e74c3c",
  "Orange category": "#e67e22",
  "Yellow category": "#f1c40f",
  "Green category": "#2ecc71",
  "Blue category": "#3498db",
  "Purple category": "#9b59b6",
};

// Outlook calendar color names to hex mapping
const OUTLOOK_CALENDAR_COLORS = {
  auto: "#0078d4",
  lightBlue: "#3498db",
  lightGreen: "#2ecc71",
  lightOrange: "#e67e22",
  lightGray: "#95a5a6",
  lightYellow: "#f1c40f",
  lightTeal: "#1abc9c",
  lightPink: "#e91e63",
  lightBrown: "#795548",
  lightRed: "#e74c3c",
  maxColor: "#0078d4",
};

/**
 * Parse Outlook datetime (with timezone) to a JS Date-compatible ISO string
 * Outlook sends: { dateTime: "2025-01-15T10:00:00.0000000", timeZone: "Pacific Standard Time" }
 */
const parseOutlookDateTime = (outlookDateTime) => {
  if (!outlookDateTime || !outlookDateTime.dateTime) return null;

  const tz = outlookDateTime.timeZone || "UTC";
  // luxon can parse IANA and Windows timezone names
  const dt = DateTime.fromISO(outlookDateTime.dateTime, { zone: tz });

  if (!dt.isValid) {
    // Fallback: try treating the datetime as UTC
    return new Date(outlookDateTime.dateTime + "Z").toISOString();
  }

  return dt.toISO();
};

/**
 * Convert a Microsoft Graph event to a FullCalendar event
 * @param {object} outlookEvent - Microsoft Graph event object
 * @param {object} calendarConfig - Connected calendar configuration
 * @returns {object} FullCalendar event object
 */
export const outlookEventToFCEvent = (outlookEvent, calendarConfig) => {
  const isAllDay = outlookEvent.isAllDay === true;

  // Determine which tag to use
  let eventTag;
  if (calendarConfig.showAsSeparateTag) {
    const tagName = calendarConfig.displayName || calendarConfig.name;
    eventTag = getTagFromName(tagName);
  }

  if (!eventTag) {
    eventTag = getTagFromName("Outlook calendar");
  }

  const eventTags = eventTag ? [eventTag] : [];

  // Check for TODO/DONE in title
  const title = outlookEvent.subject || "";
  if (title.match(/^\[\[TODO\]\]/) || title.match(/^\[\s*\]/)) {
    const todoTag = getTagFromName("TODO");
    if (todoTag) eventTags.push(todoTag);
  } else if (title.match(/^\[\[DONE\]\]/) || title.match(/^\[x\]/)) {
    const doneTag = getTagFromName("DONE");
    if (doneTag) eventTags.push(doneTag);
  }

  // Determine color
  let eventColor;
  if (getOutlookUseOriginalColors()) {
    if (outlookEvent.categories && outlookEvent.categories.length > 0) {
      eventColor = OUTLOOK_CATEGORY_COLORS[outlookEvent.categories[0]];
    }
    if (!eventColor && calendarConfig.color) {
      eventColor =
        OUTLOOK_CALENDAR_COLORS[calendarConfig.color] || calendarConfig.color;
    }
    if (!eventColor) {
      eventColor = eventTag?.color || "#0078d4";
    }
  } else {
    eventColor = eventTag?.color || "#0078d4";
  }

  // Parse dates
  let start, end;
  if (isAllDay) {
    // For all-day events, extract just the date part
    start = outlookEvent.start.dateTime.split("T")[0];
    end = outlookEvent.end.dateTime.split("T")[0];
  } else {
    start = parseOutlookDateTime(outlookEvent.start);
    end = parseOutlookDateTime(outlookEvent.end);
  }

  // Extract description text from body
  let description = "";
  if (outlookEvent.body) {
    if (outlookEvent.body.contentType === "text") {
      description = outlookEvent.body.content || "";
    } else {
      // HTML - store as-is, will be parsed when needed
      description = outlookEvent.body.content || "";
    }
  }

  const fcEvent = {
    id: `outlook-${outlookEvent.id}`,
    title: outlookEvent.subject || "(No title)",
    start,
    end,
    allDay: isAllDay,
    classNames: ["fc-event-outlook"],
    extendedProps: {
      eventTags,
      isRef: false,
      hasTime: !isAllDay,
      // Outlook-specific metadata
      outlookId: outlookEvent.id,
      outlookCalendarId: calendarConfig.id,
      outlookCalendarName:
        calendarConfig.displayName || calendarConfig.name,
      outlookChangeKey: outlookEvent.changeKey,
      outlookUpdated: outlookEvent.lastModifiedDateTime,
      description,
      location: outlookEvent.location?.displayName || "",
      syncStatus: OutlookSyncStatus.OUTLOOK_ONLY,
      isOutlookEvent: true,
      // Original Outlook data for reference
      outlookEventData: {
        webLink: outlookEvent.webLink,
        organizer: outlookEvent.organizer,
        attendees: outlookEvent.attendees,
        recurrence: outlookEvent.recurrence,
        importance: outlookEvent.importance,
        showAs: outlookEvent.showAs,
        categories: outlookEvent.categories,
        bodyContentType: outlookEvent.body?.contentType,
      },
    },
    color: eventColor,
    editable: calendarConfig.syncDirection !== "import",
    display: "block",
  };

  return fcEvent;
};

/**
 * Convert a FullCalendar/Roam event to a Microsoft Graph event
 * @param {object} fcEvent - FullCalendar event object
 * @param {string} calendarId - Target Outlook Calendar ID
 * @param {string} roamUid - Optional Roam block UID to add link to description
 * @returns {object} Microsoft Graph event resource
 */
export const fcEventToOutlookEvent = (fcEvent, calendarId, roamUid = null) => {
  let title = fcEvent.title;
  if (roamUid && fcEvent.extendedProps?.hasInfosInChildren) {
    const parentContent = getBlockContentByUid(roamUid);
    if (parentContent) {
      title = parentContent;
    }
  }

  // Extract block references for description
  const blockRefLegend = [];
  if (title) {
    uidRegex.lastIndex = 0;
    const matches = Array.from(title.matchAll(uidRegex));
    for (const match of matches) {
      const refUid = match[0].slice(2, -2);
      const resolvedContent = getBlockContentByUid(refUid);
      if (resolvedContent) {
        blockRefLegend.push({ ref: match[0], content: resolvedContent });
      }
    }
  }

  // Clean title
  const connectedCalendars = getOutlookConnectedCalendars();
  const calendarConfig = connectedCalendars.find((c) => c.id === calendarId);
  const triggerTags = calendarConfig?.triggerTags || [];
  title = cleanTitleForOutlook(title, triggerTags);

  const isAllDay = fcEvent.allDay || !fcEvent.extendedProps?.hasTime;
  const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

  const outlookEvent = {
    subject: title,
  };

  // Build description
  let description = fcEvent.extendedProps?.description || "";

  description = description
    .replace(/\n*---\nBlock references:[\s\S]*?(?=\n---\nRoam block:|$)/s, "")
    .trim();
  description = description.replace(/\n*---\nRoam block:.*$/s, "").trim();

  if (blockRefLegend.length > 0) {
    description += "\n\n---\nBlock references:";
    for (const { ref, content } of blockRefLegend) {
      const cleanedContent = cleanTitleForOutlook(content);
      description += `\n${ref} = ${cleanedContent}`;
    }
  }

  if (roamUid) {
    const graphName = window.roamAlphaAPI?.graph?.name;
    if (graphName) {
      const roamLink = `https://roamresearch.com/#/app/${graphName}/page/${roamUid}`;
      description += `\n\n---\nRoam block: ${roamLink}`;
    }
  }

  if (description) {
    outlookEvent.body = {
      contentType: "text",
      content: description,
    };
  }

  // Handle dates
  let startDate = fcEvent.start;
  if (!(startDate instanceof Date)) {
    startDate = new Date(startDate);
  }
  if (isNaN(startDate.getTime())) {
    if (fcEvent.date) {
      startDate = new Date(fcEvent.date);
    }
    if (isNaN(startDate.getTime())) {
      startDate = new Date();
    }
  }

  if (isAllDay) {
    outlookEvent.isAllDay = true;
    outlookEvent.start = {
      dateTime: formatDateForOutlook(startDate),
      timeZone: userTimeZone,
    };
  } else {
    outlookEvent.isAllDay = false;
    outlookEvent.start = {
      dateTime: formatDateTimeForOutlook(startDate),
      timeZone: userTimeZone,
    };
  }

  // Handle end time
  let endDate = fcEvent.end;
  if (endDate) {
    if (!(endDate instanceof Date)) {
      endDate = new Date(endDate);
    }
    if (isNaN(endDate.getTime())) {
      endDate = null;
    }
  }

  if (endDate) {
    if (isAllDay) {
      outlookEvent.end = {
        dateTime: formatDateForOutlook(endDate),
        timeZone: userTimeZone,
      };
    } else {
      outlookEvent.end = {
        dateTime: formatDateTimeForOutlook(endDate),
        timeZone: userTimeZone,
      };
    }
  } else {
    if (isAllDay) {
      const defaultEnd = new Date(startDate);
      defaultEnd.setDate(defaultEnd.getDate() + 1);
      outlookEvent.end = {
        dateTime: formatDateForOutlook(defaultEnd),
        timeZone: userTimeZone,
      };
    } else {
      const defaultEnd = new Date(startDate);
      defaultEnd.setHours(defaultEnd.getHours() + 1);
      outlookEvent.end = {
        dateTime: formatDateTimeForOutlook(defaultEnd),
        timeZone: userTimeZone,
      };
    }
  }

  return outlookEvent;
};

/**
 * Clean a Roam block title for Outlook Calendar
 * Removes Roam-specific syntax but preserves TODO/DONE based on user preference
 * @param {string} title - The Roam block title to clean
 * @param {string[]} triggerTagsToRemove - Optional array of trigger tags to remove
 */
export const cleanTitleForOutlook = (title, triggerTagsToRemove = null) => {
  if (!title) return "";

  let cleaned = title;

  cleaned = cleaned.replace(/^[•\-]\s*/, "");

  const checkboxFormat = getOutlookCheckboxFormat();

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
    const tagsToRemove = [
      ...new Set([...triggerTagsToRemove, "Outlook calendar"]),
    ];
    for (const tag of tagsToRemove) {
      if (!tag || !tag.trim()) continue;
      const escapedTag = tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      cleaned = cleaned.replace(
        new RegExp(`#\\[\\[${escapedTag}\\]\\]`, "gi"),
        ""
      );
      if (!tag.includes(" ")) {
        cleaned = cleaned.replace(
          new RegExp(`#${escapedTag}(?=\\s|$)`, "gi"),
          ""
        );
      }
      cleaned = cleaned.replace(
        new RegExp(`\\[\\[${escapedTag}\\]\\]`, "gi"),
        ""
      );
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
 * Format a date for Outlook (all-day events)
 * Outlook all-day format: "YYYY-MM-DDT00:00:00.0000000"
 */
export const formatDateForOutlook = (date) => {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}T00:00:00.0000000`;
};

/**
 * Format a datetime for Outlook
 * Outlook datetime format: "YYYY-MM-DDTHH:mm:ss.0000000"
 */
export const formatDateTimeForOutlook = (date) => {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");
  const seconds = String(d.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}.0000000`;
};

/**
 * Merge Outlook event data with existing FC event
 */
export const mergeOutlookDataToFCEvent = (
  fcEvent,
  outlookEvent,
  calendarConfig
) => {
  const updated = { ...fcEvent };

  updated.title = outlookEvent.subject || updated.title;

  const isAllDay = outlookEvent.isAllDay === true;
  if (isAllDay) {
    updated.start = outlookEvent.start.dateTime.split("T")[0];
    updated.end = outlookEvent.end.dateTime.split("T")[0];
  } else {
    updated.start = parseOutlookDateTime(outlookEvent.start);
    updated.end = parseOutlookDateTime(outlookEvent.end);
  }
  updated.allDay = isAllDay;

  let description = "";
  if (outlookEvent.body) {
    description = outlookEvent.body.content || "";
  }

  updated.extendedProps = {
    ...updated.extendedProps,
    outlookChangeKey: outlookEvent.changeKey,
    outlookUpdated: outlookEvent.lastModifiedDateTime,
    description,
    location: outlookEvent.location?.displayName || "",
    syncStatus: OutlookSyncStatus.SYNCED,
    outlookEventData: {
      webLink: outlookEvent.webLink,
      organizer: outlookEvent.organizer,
      attendees: outlookEvent.attendees,
      recurrence: outlookEvent.recurrence,
      importance: outlookEvent.importance,
      showAs: outlookEvent.showAs,
      categories: outlookEvent.categories,
      bodyContentType: outlookEvent.body?.contentType,
    },
  };

  return updated;
};

/**
 * Check if two events represent the same Outlook calendar entry
 */
export const isSameOutlookEvent = (fcEvent, outlookEvent) => {
  if (fcEvent.extendedProps?.outlookId === outlookEvent.id) {
    return true;
  }

  if (fcEvent.id === `outlook-${outlookEvent.id}`) {
    return true;
  }

  return false;
};

/**
 * Determine if an FC event should be synced to Outlook based on its tags
 * @param {object} fcEvent - FullCalendar event
 * @param {array} connectedCalendars - Array of connected Outlook calendar configs
 * @returns {object|null} Calendar config to sync to, or null if no match
 */
export const findOutlookCalendarForEvent = (fcEvent, connectedCalendars) => {
  const eventTags = fcEvent.extendedProps?.eventTags || [];

  for (const calendar of connectedCalendars) {
    if (!calendar.syncEnabled) continue;
    if (calendar.syncDirection === "import") continue;

    for (const eventTag of eventTags) {
      const tagName = eventTag.name?.toLowerCase();

      if (
        calendar.displayName &&
        calendar.displayName.toLowerCase() === tagName
      ) {
        return calendar;
      }

      if (calendar.triggerTags && calendar.triggerTags.length > 0) {
        if (
          calendar.triggerTags.some(
            (trigger) => trigger.toLowerCase() === tagName
          )
        ) {
          return calendar;
        }

        if (eventTag.pages && Array.isArray(eventTag.pages)) {
          for (const page of eventTag.pages) {
            if (
              calendar.triggerTags.some(
                (trigger) => trigger.toLowerCase() === page.toLowerCase()
              )
            ) {
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
 * Check if an event has any Outlook sync trigger tags
 */
export const hasOutlookSyncTriggerTag = (fcEvent, connectedCalendars) => {
  return findOutlookCalendarForEvent(fcEvent, connectedCalendars) !== null;
};

/**
 * Convert [[TODO]], [[DONE]], [ ], or [x] in Outlook title to Roam format
 */
export const convertOutlookTodoToRoam = (title) => {
  if (!title) return title;
  let converted = title;

  converted = converted.replace(/^\[\[TODO\]\]\s*/g, "{{[[TODO]]}} ");
  converted = converted.replace(/^\[\[DONE\]\]\s*/g, "{{[[DONE]]}} ");
  converted = converted.replace(/^\[\s*\]\s*/g, "{{[[TODO]]}} ");
  converted = converted.replace(/^\[x\]\s*/g, "{{[[DONE]]}} ");

  return converted;
};

/**
 * Extract Roam block content from Outlook event
 * Used when importing an Outlook event to Roam
 * @param {object} outlookEvent - Microsoft Graph event
 * @param {object} calendarConfig - Calendar configuration
 * @param {boolean} hadOriginalTimeRange - If true, the original had a time range
 */
export const outlookEventToRoamContent = (
  outlookEvent,
  calendarConfig,
  hadOriginalTimeRange = null
) => {
  let content = "";

  let title = outlookEvent.subject || "(No title)";
  title = convertOutlookTodoToRoam(title);

  const titleHasTimeRange = parseRange(title) !== null;
  const titleHasTimestamp =
    titleHasTimeRange ||
    getNormalizedTimestamp(title, strictTimestampRegex) !== null;

  // Add time for timed (non-all-day) events
  if (!outlookEvent.isAllDay && outlookEvent.start?.dateTime && !titleHasTimestamp) {
    const startDt = parseOutlookDateTime(outlookEvent.start);
    const startDate = new Date(startDt);
    const hours = startDate.getHours();
    const minutes = startDate.getMinutes();
    const timeStr = `${hours}:${String(minutes).padStart(2, "0")}`;

    const shouldIncludeEndTime =
      hadOriginalTimeRange === true || hadOriginalTimeRange === null;

    if (shouldIncludeEndTime && outlookEvent.end?.dateTime) {
      const endDt = parseOutlookDateTime(outlookEvent.end);
      const endDate = new Date(endDt);
      const endHours = endDate.getHours();
      const endMinutes = endDate.getMinutes();
      const endTimeStr = `${endHours}:${String(endMinutes).padStart(2, "0")}`;

      const durationMs = endDate.getTime() - startDate.getTime();
      const isDefaultDuration = durationMs === 3600000;

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
  const tagToAdd = customTag || "Outlook calendar";
  content += tagToAdd.includes(" ")
    ? ` #[[${tagToAdd}]]`
    : ` #${tagToAdd}`;

  return content;
};

/**
 * Parse HTML description from Outlook into an array of Roam block contents
 * @param {string} htmlDescription - HTML description from Outlook event body
 * @returns {string[]} Array of block contents
 */
export const parseOutlookDescriptionToBlocks = (htmlDescription) => {
  if (!htmlDescription) return [];

  let text = htmlDescription;

  // Remove Roam link section
  text = text.replace(/\n*---\n*Roam block:.*$/s, "").trim();
  text = text
    .replace(/\n*---\n*Block references:[\s\S]*?(?=\n---\n|$)/s, "")
    .trim();

  // Convert HTML to text
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/p>/gi, "\n");
  text = text.replace(/<\/div>/gi, "\n");
  text = text.replace(/<li[^>]*>/gi, "\n• ");
  text = text.replace(/<\/li>/gi, "");
  text = text.replace(
    /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([^<]*)<\/a>/gi,
    "[$2]($1)"
  );
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
 * Parse Outlook event metadata into Roam child blocks
 * @param {object} event - FullCalendar event with extendedProps
 * @returns {string[]} Array of block contents for metadata
 */
export const parseOutlookMetadataToBlocks = (event) => {
  const blocks = [];
  const extendedProps = event.extendedProps || {};
  const outlookEventData = extendedProps.outlookEventData || {};

  if (extendedProps.location) {
    blocks.push(`Location:: ${extendedProps.location}`);
  }

  if (outlookEventData.attendees && outlookEventData.attendees.length > 0) {
    const attendeesList = outlookEventData.attendees
      .map((attendee) => {
        const emailAddr = attendee.emailAddress || {};
        const displayName = emailAddr.name || emailAddr.address;
        return emailAddr.name
          ? `[[${emailAddr.name}]]`
          : `[[${emailAddr.address}]]`;
      })
      .join(", ");
    blocks.push(`Attendees:: ${attendeesList}`);
  }

  return blocks;
};

/**
 * Parse all Outlook event data into an array of Roam child blocks
 */
export const parseOutlookDataToRoamBlocks = (event) => {
  const blocks = [];

  const descriptionBlocks = parseOutlookDescriptionToBlocks(
    event.extendedProps?.description
  );
  blocks.push(...descriptionBlocks);

  const metadataBlocks = parseOutlookMetadataToBlocks(event);
  blocks.push(...metadataBlocks);

  return blocks;
};
