/**
 * Outlook Calendar Event Mapping Utilities
 *
 * Handles conversion between Microsoft Graph events and FullCalendar events.
 * Delegates shared logic to calendarMapping.js to avoid duplication with Google Calendar.
 */

import { DateTime } from "luxon";
import { getTagFromName } from "../models/EventTag";
import { OutlookSyncStatus } from "../models/OutlookSyncMetadata";
import {
  getOutlookUseOriginalColors,
  getOutlookCheckboxFormat,
  getOutlookConnectedCalendars,
} from "../services/outlookCalendarService";
import { getBlockContentByUid } from "./roamApi";
import {
  cleanTitleForCalendar,
  convertCalTodoToRoam,
  buildRoamContentFromCalEvent,
  extractBlockReferences,
  buildCalendarDescription,
  findCalendarForEvent as sharedFindCalendarForEvent,
  hasSyncTriggerTag as sharedHasSyncTriggerTag,
  parseCalDescriptionToBlocks,
  parseCalMetadataToBlocks,
  normalizeStartDate,
  normalizeEndDate,
  buildDefaultEndDate,
} from "./calendarMapping";

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
export const parseOutlookDateTime = (outlookDateTime) => {
  if (!outlookDateTime || !outlookDateTime.dateTime) return null;

  const tz = outlookDateTime.timeZone || "UTC";
  const dt = DateTime.fromISO(outlookDateTime.dateTime, { zone: tz });

  if (!dt.isValid) {
    return new Date(outlookDateTime.dateTime + "Z").toISOString();
  }

  return dt.toISO();
};

/**
 * Convert a Microsoft Graph event to a FullCalendar event
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
    start = outlookEvent.start.dateTime.split("T")[0];
    end = outlookEvent.end.dateTime.split("T")[0];
  } else {
    start = parseOutlookDateTime(outlookEvent.start);
    end = parseOutlookDateTime(outlookEvent.end);
  }

  // Extract description text from body
  let description = "";
  if (outlookEvent.body) {
    description = outlookEvent.body.content || "";
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
 * Convert a FullCalendar/Roam event to a Microsoft Graph event.
 * Uses shared utilities for block refs, description building, and date normalization.
 */
export const fcEventToOutlookEvent = (fcEvent, calendarId, roamUid = null) => {
  let title = fcEvent.title;
  if (roamUid && fcEvent.extendedProps?.hasInfosInChildren) {
    const parentContent = getBlockContentByUid(roamUid);
    if (parentContent) {
      title = parentContent;
    }
  }

  // Extract block references using shared utility
  const blockRefLegend = extractBlockReferences(title);

  // Clean title using Outlook-specific cleaner
  const connectedCalendars = getOutlookConnectedCalendars();
  const calendarConfig = connectedCalendars.find((c) => c.id === calendarId);
  const triggerTags = calendarConfig?.triggerTags || [];
  title = cleanTitleForOutlook(title, triggerTags);

  const isAllDay = fcEvent.allDay || !fcEvent.extendedProps?.hasTime;
  const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

  const outlookEvent = {
    subject: title,
  };

  // Build description using shared utility
  const description = buildCalendarDescription(
    fcEvent.extendedProps?.description || "",
    blockRefLegend,
    roamUid,
    (content) => cleanTitleForOutlook(content)
  );

  if (description) {
    outlookEvent.body = {
      contentType: "text",
      content: description,
    };
  }

  // Normalize dates using shared utilities
  const startDate = normalizeStartDate(fcEvent.start, fcEvent.date);
  let endDate = normalizeEndDate(fcEvent.end);

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

  if (!endDate) {
    endDate = buildDefaultEndDate(startDate, isAllDay);
  }

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

  return outlookEvent;
};

/**
 * Clean a Roam block title for Outlook Calendar.
 * Delegates to shared cleanTitleForCalendar with Outlook-specific defaults.
 */
export const cleanTitleForOutlook = (title, triggerTagsToRemove = null) => {
  return cleanTitleForCalendar(title, {
    triggerTagsToRemove,
    defaultTag: "Outlook calendar",
    getCheckboxFormat: getOutlookCheckboxFormat,
  });
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
 * Determine if an FC event should be synced to Outlook based on its tags.
 * Delegates to shared findCalendarForEvent.
 */
export const findOutlookCalendarForEvent = sharedFindCalendarForEvent;

/**
 * Check if an event has any Outlook sync trigger tags.
 * Delegates to shared hasSyncTriggerTag.
 */
export const hasOutlookSyncTriggerTag = sharedHasSyncTriggerTag;

/**
 * Convert [[TODO]], [[DONE]], [ ], or [x] in Outlook title to Roam format.
 * Delegates to shared convertCalTodoToRoam.
 */
export const convertOutlookTodoToRoam = convertCalTodoToRoam;

/**
 * Extract Roam block content from Outlook event.
 * Delegates to shared buildRoamContentFromCalEvent with Outlook field mapping.
 */
export const outlookEventToRoamContent = (
  outlookEvent,
  calendarConfig,
  hadOriginalTimeRange = null
) => {
  const isAllDay = outlookEvent.isAllDay === true;
  let startDate = null;
  let endDate = null;

  if (!isAllDay && outlookEvent.start?.dateTime) {
    const startDt = parseOutlookDateTime(outlookEvent.start);
    startDate = startDt ? new Date(startDt) : null;
  }
  if (!isAllDay && outlookEvent.end?.dateTime) {
    const endDt = parseOutlookDateTime(outlookEvent.end);
    endDate = endDt ? new Date(endDt) : null;
  }

  return buildRoamContentFromCalEvent({
    title: outlookEvent.subject,
    isAllDay,
    startDate,
    endDate,
    calendarConfig,
    hadOriginalTimeRange,
    defaultTag: "Outlook calendar",
  });
};

/**
 * Parse HTML description from Outlook into an array of Roam block contents.
 * Delegates to shared parseCalDescriptionToBlocks.
 */
export const parseOutlookDescriptionToBlocks = parseCalDescriptionToBlocks;

/**
 * Parse Outlook event metadata into Roam child blocks.
 * Delegates to shared parseCalMetadataToBlocks with Outlook-specific field accessors.
 */
export const parseOutlookMetadataToBlocks = (event) => {
  const extendedProps = event.extendedProps || {};
  const outlookEventData = extendedProps.outlookEventData || {};

  return parseCalMetadataToBlocks({
    location: extendedProps.location,
    attendees: outlookEventData.attendees,
    getAttendeeName: (attendee) => {
      const emailAddr = attendee.emailAddress || {};
      return emailAddr.name || emailAddr.address;
    },
  });
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
