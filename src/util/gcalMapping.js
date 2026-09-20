/**
 * Google Calendar Event Mapping Utilities
 *
 * Handles conversion between Google Calendar events and FullCalendar events.
 * Delegates shared logic to calendarMapping.js to avoid duplication with Outlook.
 */

import { getTagFromName } from "../models/EventTag";
import { SyncStatus } from "../models/SyncMetadata";
import { getUseOriginalColors, getCheckboxFormat, getConnectedCalendars } from "../services/googleCalendarService";
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

// Google Calendar event colorId to hex color mapping
// See: https://developers.google.com/calendar/api/v3/reference/colors
const GCAL_EVENT_COLORS = {
  "1": "#a4bdfc", // Lavender
  "2": "#7ae7bf", // Sage
  "3": "#dbadff", // Grape
  "4": "#ff887c", // Flamingo
  "5": "#fbd75b", // Banana
  "6": "#ffb878", // Tangerine
  "7": "#46d6db", // Peacock
  "8": "#e1e1e1", // Graphite
  "9": "#5484ed", // Blueberry
  "10": "#51b749", // Basil
  "11": "#dc2127", // Tomato
};

// ============================================
// Google Tasks Detection Utilities
// ============================================

/**
 * Pattern to detect Google Tasks in calendar event descriptions
 * Tasks created in Google Calendar have a URL like: https://tasks.google.com/task/{taskId}
 */
const TASKS_URL_PATTERN = /https:\/\/tasks\.google\.com\/task\/([a-zA-Z0-9_-]+)/;

/**
 * Check if a Google Calendar event is actually a Google Task
 * Tasks appear in Calendar with a generic description containing a tasks.google.com link
 * @param {object} gcalEvent - Google Calendar event object
 * @returns {boolean} True if the event is a Google Task
 */
export const isGCalTask = (gcalEvent) => {
  return TASKS_URL_PATTERN.test(gcalEvent.description || "");
};

/**
 * Extract the Google Task ID from a calendar event's description
 * @param {object} gcalEvent - Google Calendar event object
 * @returns {string|null} Task ID if found, null otherwise
 */
export const extractTaskIdFromEvent = (gcalEvent) => {
  const match = (gcalEvent.description || "").match(TASKS_URL_PATTERN);
  return match ? match[1] : null;
};

/**
 * Convert a Google Calendar event to a FullCalendar event
 * @param {object} gcalEvent - Google Calendar event object
 * @param {object} calendarConfig - Connected calendar configuration
 * @returns {object} FullCalendar event object
 */
export const gcalEventToFCEvent = (gcalEvent, calendarConfig) => {
  const isAllDay = !gcalEvent.start.dateTime;

  // Determine which tag to use based on showAsSeparateTag
  let eventTag;
  if (calendarConfig.showAsSeparateTag) {
    // Calendar has its own separate tag - use displayName
    const tagName = calendarConfig.displayName || calendarConfig.name;
    eventTag = getTagFromName(tagName);
  }

  // Fall back to main "Google calendar" tag
  if (!eventTag) {
    eventTag = getTagFromName("Google calendar");
  }

  // Build event tags array
  const eventTags = eventTag ? [eventTag] : [];

  // For Google Tasks, add TODO or DONE tag based on task status
  const taskData = gcalEvent._taskData;
  if (taskData) {
    const statusTag = taskData.status === "completed"
      ? getTagFromName("DONE")
      : getTagFromName("TODO");
    if (statusTag) {
      eventTags.push(statusTag);
    }
  } else {
    // For non-task GCal events, check if title has checkbox markers
    const title = gcalEvent.summary || "";
    if (title.match(/^\[\[TODO\]\]/) || title.match(/^\[\s*\]/)) {
      const todoTag = getTagFromName("TODO");
      if (todoTag) {
        eventTags.push(todoTag);
      }
    } else if (title.match(/^\[\[DONE\]\]/) || title.match(/^\[x\]/)) {
      const doneTag = getTagFromName("DONE");
      if (doneTag) {
        eventTags.push(doneTag);
      }
    }
  }

  // Determine color - use original GCal color if setting enabled, otherwise use tag color
  let eventColor;
  if (getUseOriginalColors()) {
    // Priority: event's colorId > calendar's backgroundColor > tag color > default
    if (gcalEvent.colorId && GCAL_EVENT_COLORS[gcalEvent.colorId]) {
      eventColor = GCAL_EVENT_COLORS[gcalEvent.colorId];
    } else if (calendarConfig.backgroundColor) {
      eventColor = calendarConfig.backgroundColor;
    } else {
      eventColor = eventTag?.color || "#4285f4";
    }
  } else {
    eventColor = eventTag?.color || "#4285f4";
  }

  const fcEvent = {
    id: `gcal-${gcalEvent.id}`, // Prefix to distinguish from Roam UIDs
    title: gcalEvent.summary || "(No title)",
    start: isAllDay ? gcalEvent.start.date : gcalEvent.start.dateTime,
    end: isAllDay ? gcalEvent.end.date : gcalEvent.end.dateTime,
    allDay: isAllDay,
    classNames: ["fc-event-gcal"],
    extendedProps: {
      eventTags,
      isRef: false,
      hasTime: !isAllDay,
      // GCal-specific metadata
      gCalId: gcalEvent.id,
      gCalCalendarId: calendarConfig.id,
      gCalCalendarName: calendarConfig.displayName || calendarConfig.name, // Display name for calendar
      gCalEtag: gcalEvent.etag,
      gCalUpdated: gcalEvent.updated,
      description: gcalEvent.description || "",
      location: gcalEvent.location || "",
      attachments: gcalEvent.attachments || [],
      syncStatus: SyncStatus.GCAL_ONLY,
      isGCalEvent: true,
      // Google Task data (if this event is a task enriched by taskService)
      _taskData: gcalEvent._taskData || null,
      // Original GCal data for reference
      gCalEventData: {
        htmlLink: gcalEvent.htmlLink,
        creator: gcalEvent.creator,
        organizer: gcalEvent.organizer,
        attendees: gcalEvent.attendees,
        recurrence: gcalEvent.recurrence,
        recurringEventId: gcalEvent.recurringEventId,
        status: gcalEvent.status,
      },
    },
    color: eventColor,
    editable: calendarConfig.syncDirection !== "import",
    // Don't set url property - it causes FullCalendar to navigate on click
    // Store htmlLink in extendedProps instead for manual access
    display: "block",
  };

  return fcEvent;
};

/**
 * Convert a FullCalendar/Roam event to a Google Calendar event
 * @param {object} fcEvent - FullCalendar event object
 * @param {string} calendarId - Target Google Calendar ID
 * @param {string} roamUid - Optional Roam block UID to add link to description
 * @returns {object} Google Calendar event resource
 */
export const fcEventToGCalEvent = (fcEvent, calendarId, roamUid = null) => {
  // For events with children, use only the parent block content
  let title = fcEvent.title;
  if (roamUid && fcEvent.extendedProps?.hasInfosInChildren) {
    const parentContent = getBlockContentByUid(roamUid);
    if (parentContent) {
      title = parentContent;
    }
  }

  // Extract block references using shared utility
  const blockRefLegend = extractBlockReferences(title);

  // Clean title using GCal-specific cleaner
  const connectedCalendars = getConnectedCalendars();
  const calendarConfig = connectedCalendars.find((c) => c.id === calendarId);
  const triggerTags = calendarConfig?.triggerTags || [];
  title = cleanTitleForGCal(title, triggerTags);
  const isAllDay = fcEvent.allDay || !fcEvent.extendedProps?.hasTime;

  const gcalEvent = {
    summary: title,
  };

  // Build description using shared utility
  const description = buildCalendarDescription(
    fcEvent.extendedProps?.description || "",
    blockRefLegend,
    roamUid,
    (content) => cleanTitleForGCal(content)
  );

  if (description) {
    gcalEvent.description = description;
  }

  // Normalize dates using shared utilities
  const startDate = normalizeStartDate(fcEvent.start, fcEvent.date);
  let endDate = normalizeEndDate(fcEvent.end);

  // Handle start time
  if (isAllDay) {
    gcalEvent.start = { date: formatDateForGCal(startDate) };
  } else {
    gcalEvent.start = {
      dateTime: formatDateTimeForGCal(startDate),
      timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    };
  }

  // Handle end time
  if (!endDate) {
    endDate = buildDefaultEndDate(startDate, isAllDay);
  }

  if (isAllDay) {
    gcalEvent.end = { date: formatDateForGCal(endDate) };
  } else {
    gcalEvent.end = {
      dateTime: formatDateTimeForGCal(endDate),
      timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    };
  }

  return gcalEvent;
};

/**
 * Clean a Roam block title for Google Calendar.
 * Delegates to shared cleanTitleForCalendar with GCal-specific defaults.
 */
export const cleanTitleForGCal = (title, triggerTagsToRemove = null) => {
  return cleanTitleForCalendar(title, {
    triggerTagsToRemove,
    defaultTag: "Google calendar",
    getCheckboxFormat,
  });
};

/**
 * Format a date for Google Calendar (all-day events)
 * Format: YYYY-MM-DD
 */
export const formatDateForGCal = (date) => {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

/**
 * Format a datetime for Google Calendar
 * Format: YYYY-MM-DDTHH:mm:ss
 */
export const formatDateTimeForGCal = (date) => {
  const d = new Date(date);
  return d.toISOString();
};

/**
 * Merge GCal event data with existing FC event
 * Used when updating a synced event
 */
export const mergeGCalDataToFCEvent = (fcEvent, gcalEvent, calendarConfig) => {
  const updated = { ...fcEvent };

  // Update basic properties
  updated.title = gcalEvent.summary || updated.title;

  const isAllDay = !gcalEvent.start.dateTime;
  updated.start = isAllDay ? gcalEvent.start.date : gcalEvent.start.dateTime;
  updated.end = isAllDay
    ? gcalEvent.end?.date
    : gcalEvent.end?.dateTime || null;
  updated.allDay = isAllDay;

  // Update extended props
  updated.extendedProps = {
    ...updated.extendedProps,
    gCalEtag: gcalEvent.etag,
    gCalUpdated: gcalEvent.updated,
    description: gcalEvent.description || "",
    location: gcalEvent.location || "",
    attachments: gcalEvent.attachments || [],
    syncStatus: SyncStatus.SYNCED,
    // Store GCal event data for access to htmlLink and other metadata
    gCalEventData: {
      htmlLink: gcalEvent.htmlLink,
      creator: gcalEvent.creator,
      organizer: gcalEvent.organizer,
      attendees: gcalEvent.attendees,
      recurrence: gcalEvent.recurrence,
      recurringEventId: gcalEvent.recurringEventId,
      status: gcalEvent.status,
    },
  };

  return updated;
};

/**
 * Check if two events represent the same calendar entry
 */
export const isSameEvent = (fcEvent, gcalEvent) => {
  // Check by GCal ID stored in extendedProps
  if (fcEvent.extendedProps?.gCalId === gcalEvent.id) {
    return true;
  }

  // Check by prefixed ID
  if (fcEvent.id === `gcal-${gcalEvent.id}`) {
    return true;
  }

  return false;
};

/**
 * Determine if an FC event should be synced based on its tags.
 * Delegates to shared findCalendarForEvent.
 */
export const findCalendarForEvent = sharedFindCalendarForEvent;

/**
 * Check if an event has any sync trigger tags.
 * Delegates to shared hasSyncTriggerTag.
 */
export const hasSyncTriggerTag = sharedHasSyncTriggerTag;

/**
 * Convert [[TODO]], [[DONE]], [ ], or [x] in GCal title to Roam format.
 * Delegates to shared convertCalTodoToRoam.
 */
export const convertGCalTodoToRoam = convertCalTodoToRoam;

/**
 * Extract Roam block content from GCal event.
 * Delegates to shared buildRoamContentFromCalEvent with GCal-specific field mapping.
 */
export const gcalEventToRoamContent = (gcalEvent, calendarConfig, hadOriginalTimeRange = null) => {
  const isAllDay = !gcalEvent.start.dateTime;
  return buildRoamContentFromCalEvent({
    title: gcalEvent.summary,
    isAllDay,
    startDate: gcalEvent.start.dateTime ? new Date(gcalEvent.start.dateTime) : null,
    endDate: gcalEvent.end?.dateTime ? new Date(gcalEvent.end.dateTime) : null,
    calendarConfig,
    hadOriginalTimeRange,
    defaultTag: "Google calendar",
  });
};

/**
 * Parse HTML description from GCal into an array of Roam block contents.
 * Delegates to shared parseCalDescriptionToBlocks.
 */
export const parseGCalDescriptionToBlocks = parseCalDescriptionToBlocks;

/**
 * Parse GCal event metadata into Roam child blocks.
 * Delegates to shared parseCalMetadataToBlocks with GCal-specific field accessors.
 */
export const parseGCalMetadataToBlocks = (event) => {
  const extendedProps = event.extendedProps || {};
  const gCalEventData = extendedProps.gCalEventData || {};

  return parseCalMetadataToBlocks({
    location: extendedProps.location,
    attendees: gCalEventData.attendees,
    getAttendeeName: (attendee) => attendee.displayName || attendee.email,
    attachments: extendedProps.attachments,
    formatAttachment: (attachment) => {
      const title = attachment.title || attachment.fileUrl || "Attachment";
      const url = attachment.fileUrl || attachment.iconLink;
      return `[${title}](${url})`;
    },
  });
};

/**
 * Parse all GCal event data into an array of Roam child blocks
 * Combines description parsing and metadata extraction
 * @param {object} event - FullCalendar event with extendedProps
 * @returns {string[]} Array of all child block contents to create
 */
export const parseGCalDataToRoamBlocks = (event) => {
  const blocks = [];

  // Parse description into blocks
  const descriptionBlocks = parseGCalDescriptionToBlocks(
    event.extendedProps?.description
  );
  blocks.push(...descriptionBlocks);

  // Add metadata blocks
  const metadataBlocks = parseGCalMetadataToBlocks(event);
  blocks.push(...metadataBlocks);

  return blocks;
};

export default {
  // Task detection
  isGCalTask,
  extractTaskIdFromEvent,
  // Event mapping
  gcalEventToFCEvent,
  fcEventToGCalEvent,
  cleanTitleForGCal,
  formatDateForGCal,
  formatDateTimeForGCal,
  mergeGCalDataToFCEvent,
  isSameEvent,
  findCalendarForEvent,
  hasSyncTriggerTag,
  gcalEventToRoamContent,
  convertGCalTodoToRoam,
  // GCal to Roam parsing
  parseGCalDescriptionToBlocks,
  parseGCalMetadataToBlocks,
  parseGCalDataToRoamBlocks,
};
