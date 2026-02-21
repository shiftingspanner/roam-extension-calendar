/**
 * Shared Sync Helpers
 *
 * Provider-agnostic helper functions shared between Google Calendar sync
 * and Outlook Calendar sync. Eliminates exact code duplication.
 */

import {
  getBlockContentByUid,
  updateBlock,
  createChildBlock,
  getParentBlock,
  deleteBlockIfNoChild,
  isExistingNode,
  getTreeByUid,
  getEventDateFromBlock,
  blockHasCalendarTag,
  addTagToBlock,
} from "../util/roamApi";

import { parseRange, dateToISOString } from "../util/dates";
import { acquireSyncLock, releaseSyncLock } from "./syncLockService";
import { Toaster, Position, Intent } from "@blueprintjs/core";
import { areEventsDuplicate } from "./deduplicationService";
import { getCalendarUidFromPage } from "../util/data";
import { startDateRegex, untilDateRegex, roamDateRegex } from "../util/regex";
import { rangeEndAttribute } from "../index";

/**
 * Helper to detect if a block content contains TODO marker.
 */
export const hasTodoMarker = (content) => {
  return content && content.includes("{{[[TODO]]}}");
};

/**
 * Find child blocks containing date information (start:: or end::/until::).
 *
 * @param {string} parentUid - Parent block UID
 * @returns {object} Object with startBlock and endBlock info
 */
export const findDateChildBlocks = (parentUid) => {
  const tree = getTreeByUid(parentUid);
  if (!tree || !tree[0] || !tree[0].children) {
    return { startBlock: null, endBlock: null };
  }

  let startBlock = null;
  let endBlock = null;

  for (const child of tree[0].children) {
    const content = child.string || "";

    if (startDateRegex) {
      startDateRegex.lastIndex = 0;
      if (startDateRegex.test(content)) {
        startBlock = { uid: child.uid, content };
      }
    }

    if (untilDateRegex) {
      untilDateRegex.lastIndex = 0;
      if (untilDateRegex.test(content)) {
        endBlock = { uid: child.uid, content };
      }
    }

    roamDateRegex.lastIndex = 0;
    if (roamDateRegex.test(content) && !startBlock && !endBlock) {
      if (
        rangeEndAttribute &&
        content.toLowerCase().includes(rangeEndAttribute.toLowerCase())
      ) {
        endBlock = { uid: child.uid, content };
      }
    }
  }

  return { startBlock, endBlock };
};

/**
 * Update a child block's date reference.
 *
 * @param {string} blockUid - Block UID to update
 * @param {string} currentContent - Current block content
 * @param {Date} newDate - New date to set
 */
export const updateChildBlockDate = async (blockUid, currentContent, newDate) => {
  const newRoamDate = window.roamAlphaAPI.util.dateToPageTitle(newDate);
  roamDateRegex.lastIndex = 0;
  const matchingDates = currentContent.match(roamDateRegex);

  if (matchingDates && matchingDates.length) {
    const currentDateStr = matchingDates[0].replace("[[", "").replace("]]", "");
    const newContent = currentContent.replace(currentDateStr, newRoamDate);
    await updateBlock(blockUid, newContent);
  }
};

/**
 * Format event date/time for display in toast notifications.
 *
 * @param {string} startDateTime - ISO date or datetime string
 * @param {string} endDateTime - ISO date or datetime string (optional)
 * @returns {string} Formatted date/time string
 */
export const formatEventDateTime = (startDateTime, endDateTime) => {
  if (!startDateTime) return "";

  const start = new Date(startDateTime);
  const hasTime = startDateTime.includes("T");

  const dateFormat = new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
  });

  const timeFormat = new Intl.DateTimeFormat("en-US", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });

  let result = dateFormat.format(start);

  if (hasTime) {
    result += " at " + timeFormat.format(start);

    if (endDateTime) {
      const end = new Date(endDateTime);
      if (start.toDateString() !== end.toDateString()) {
        result += " - " + dateFormat.format(end) + " at " + timeFormat.format(end);
      } else {
        result += " - " + timeFormat.format(end);
      }
    }
  } else if (endDateTime) {
    const end = new Date(endDateTime);
    if (start.toDateString() !== end.toDateString()) {
      result += " - " + dateFormat.format(end);
    }
  }

  return result;
};

/**
 * Create a generic sync result object.
 *
 * @param {object} options
 * @param {string} options.deletedKeyName - Key name for "deleted from provider" (e.g., "deletedFromGCal")
 * @returns {object} Sync result
 */
export const createGenericSyncResult = (deletedKeyName = "deletedFromProvider") => ({
  imported: [],
  exported: [],
  updated: [],
  conflicts: [],
  errors: [],
  [deletedKeyName]: [],
  deletedFromRoam: [],
});

/**
 * Show sync result toast notification (provider-agnostic).
 *
 * @param {object} result - Sync result
 * @param {string} blockUid - Block UID for fetching content
 * @param {string} providerName - Calendar provider display name (e.g., "Google Calendar", "Outlook")
 */
export const showGenericSyncResultToast = (result, blockUid, providerName) => {
  const toaster = Toaster.create({ position: Position.TOP });

  if (result.success) {
    const blockContent = getBlockContentByUid(blockUid) || "Event";
    const actionText = result.action === "created" ? "created in" : "updated in";
    const eventTitle =
      blockContent.length > 50 ? blockContent.substring(0, 47) + "..." : blockContent;

    const dateTimeStr = formatEventDateTime(result.eventStart, result.eventEnd);
    const dateTimeInfo = dateTimeStr ? ` (${dateTimeStr})` : "";

    toaster.show({
      message: `"${eventTitle}" ${actionText} ${providerName}: ${result.calendarName}${dateTimeInfo}`,
      intent: Intent.SUCCESS,
      icon: "tick-circle",
      timeout: 4000,
    });
  } else {
    let intent = Intent.DANGER;
    let icon = "error";

    if (
      result.error.includes("not determine event date") ||
      result.error.includes("Block not found")
    ) {
      intent = Intent.WARNING;
      icon = "warning-sign";
    }

    toaster.show({
      message: result.error,
      intent: intent,
      icon: icon,
      timeout: 5000,
    });
  }
};

/**
 * Sync a block to the default calendar of a given provider.
 * Shared orchestration logic for syncBlockToDefaultCalendar / syncBlockToDefaultOutlookCalendar.
 *
 * @param {object|string} blockContextOrUid - Block context or UID
 * @param {object} options
 * @param {Function} options.getConnectedCalendars - Function to get connected calendars
 * @param {Function} options.syncEventFn - Function(roamUid, fcEvent, calendarId) to sync event
 * @param {string} options.providerName - Provider display name for error messages
 * @returns {object} Sync result
 */
export const syncBlockToDefaultProviderCalendar = async (blockContextOrUid, {
  getConnectedCalendars,
  syncEventFn,
  providerName,
}) => {
  try {
    const blockUid =
      typeof blockContextOrUid === "string"
        ? blockContextOrUid
        : blockContextOrUid["block-uid"];

    const blockContent = getBlockContentByUid(blockUid);

    if (!blockContent) {
      return { success: false, error: "Block not found or empty" };
    }

    const calendars = getConnectedCalendars();
    if (!calendars || calendars.length === 0) {
      return {
        success: false,
        error: `No ${providerName} connected. Please configure ${providerName} first.`,
      };
    }

    const defaultCalendar = calendars.find(
      (cal) => cal.syncEnabled && cal.syncDirection !== "import"
    );
    if (!defaultCalendar) {
      return {
        success: false,
        error: `No ${providerName.toLowerCase()} calendar available for sync. Please enable sync for at least one calendar.`,
      };
    }

    const eventDate = getEventDateFromBlock(blockUid);
    if (!eventDate) {
      return {
        success: false,
        error: "Could not determine event date. Block must be in a Daily Note Page or contain a date reference.",
      };
    }

    const eventDateStr = dateToISOString(eventDate);
    const rangeInfo = parseRange(blockContent);

    const fcEvent = {
      id: blockUid,
      title: blockContent,
      start: rangeInfo ? `${eventDateStr}T${rangeInfo.range.start}` : eventDateStr,
      end: rangeInfo && rangeInfo.range.end ? `${eventDateStr}T${rangeInfo.range.end}` : null,
      extendedProps: { eventTags: [] },
    };

    // Add calendar tag to block if not present
    if (!blockHasCalendarTag(blockUid, defaultCalendar)) {
      let tagToAdd = null;
      if (defaultCalendar.triggerTags && defaultCalendar.triggerTags.length > 0) {
        tagToAdd = defaultCalendar.triggerTags[0];
      } else if (defaultCalendar.displayName) {
        tagToAdd = defaultCalendar.displayName;
      }

      if (tagToAdd) {
        await addTagToBlock(blockUid, tagToAdd);
      }
    }

    const result = await syncEventFn(blockUid, fcEvent, defaultCalendar.id);

    if (result.success) {
      return {
        success: true,
        action: result.action,
        calendarName: defaultCalendar.name,
        eventStart: fcEvent.start,
        eventEnd: fcEvent.end,
        ...(result.gCalId ? { gCalId: result.gCalId } : {}),
        ...(result.outlookId ? { outlookId: result.outlookId } : {}),
      };
    } else if (result.skipped) {
      return { success: false, error: "Sync already in progress for this block" };
    } else {
      return { success: false, error: friendlyErrorMessage(result.error, providerName) };
    }
  } catch (error) {
    console.error(`[${providerName}Sync] Error syncing block:`, error);
    return {
      success: false,
      error: friendlyErrorMessage(error.message, providerName),
    };
  }
};

/**
 * Convert raw error messages to user-friendly messages.
 *
 * @param {string} errorMessage - Raw error message
 * @param {string} providerName - Provider name for messages
 * @returns {string} User-friendly error message
 */
const friendlyErrorMessage = (errorMessage, providerName) => {
  if (!errorMessage) return "Unknown error";

  if (
    errorMessage.includes("Failed to fetch") ||
    errorMessage.includes("NetworkError") ||
    errorMessage.includes("network")
  ) {
    return `Unable to connect to ${providerName}. Please check your internet connection.`;
  } else if (errorMessage.includes("401") || errorMessage.includes("Unauthorized")) {
    return `${providerName} authentication expired. Please reconnect your calendar.`;
  } else if (errorMessage.includes("403") || errorMessage.includes("Forbidden")) {
    return `Permission denied. Please check your ${providerName} permissions.`;
  } else if (errorMessage.includes("404")) {
    return "Calendar not found. The calendar may have been deleted.";
  }

  return errorMessage;
};

/**
 * Handle multi-day event child blocks during import or update.
 * Creates or updates end-date child blocks as needed.
 *
 * @param {string} blockUid - The event block UID
 * @param {Date} startDate - Event start date
 * @param {Date|null} endDateObj - Event end date (exclusive for all-day events - already adjusted)
 * @param {boolean} isAllDay - Whether this is an all-day event
 */
export const handleMultiDayChildBlocks = async (blockUid, startDate, endDateObj) => {
  if (!endDateObj || !rangeEndAttribute) return;

  if (startDate.toDateString() !== endDateObj.toDateString()) {
    const endDateStr = window.roamAlphaAPI.util.dateToPageTitle(endDateObj);
    const endBlockContent = `${rangeEndAttribute}:: [[${endDateStr}]]`;
    await createChildBlock(blockUid, endBlockContent, "first");
  }
};

/**
 * Move a block to the correct Daily Note Page if its date changed.
 *
 * @param {string} roamUid - Block UID to move
 * @param {string} newDnpUid - Target DNP UID
 */
export const moveBlockToDnpIfNeeded = async (roamUid, newDnpUid) => {
  const currentParentUid = getParentBlock(roamUid);
  if (!currentParentUid) return;

  let currentDnpUid = getParentBlock(currentParentUid);
  if (!currentDnpUid) {
    currentDnpUid = currentParentUid;
  }

  const currentDateFromUid = window.roamAlphaAPI.util.pageTitleToDate(
    window.roamAlphaAPI.pull("[:node/title]", [":block/uid", currentDnpUid])?.[":node/title"]
  );

  if (currentDateFromUid) {
    const currentDnpDateStr = window.roamAlphaAPI.util.dateToPageUid(currentDateFromUid);

    if (newDnpUid !== currentDnpDateStr) {
      const newCalendarBlockUid = await getCalendarUidFromPage(newDnpUid);
      await window.roamAlphaAPI.moveBlock({
        location: {
          "parent-uid": newCalendarBlockUid,
          order: "last",
        },
        block: { uid: roamUid },
      });

      deleteBlockIfNoChild(currentParentUid);
    }
  }
};

/**
 * Update child blocks with start/end dates during a calendar-to-Roam update.
 *
 * @param {string} roamUid - Block UID
 * @param {Date} startDateObj - Calendar event start date
 * @param {Date|null} endDateObj - Calendar event end date (already adjusted for exclusive all-day)
 */
export const updateDateChildBlocks = async (roamUid, startDateObj, endDateObj) => {
  const { startBlock, endBlock } = findDateChildBlocks(roamUid);

  const isMultiDayEvent =
    endDateObj && startDateObj.toDateString() !== endDateObj.toDateString();

  if (startBlock) {
    await updateChildBlockDate(startBlock.uid, startBlock.content, startDateObj);
  }

  if (endBlock) {
    if (isMultiDayEvent && endDateObj) {
      await updateChildBlockDate(endBlock.uid, endBlock.content, endDateObj);
    }
  } else if (isMultiDayEvent && endDateObj && rangeEndAttribute) {
    const endDateStr = window.roamAlphaAPI.util.dateToPageTitle(endDateObj);
    const endBlockContent = `${rangeEndAttribute}:: [[${endDateStr}]]`;
    await createChildBlock(roamUid, endBlockContent, "first");
  }
};

/**
 * Find matching events in an external calendar for deduplication.
 *
 * @param {object} fcEvent - FullCalendar event from Roam
 * @param {string} calendarId - Calendar ID
 * @param {Function} getEventsFn - Function(calendarId, startOfDay, endOfDay) to fetch events
 * @param {Function} getRoamUidByExternalIdFn - Function(externalId) to check if already linked
 * @param {Function} getSubjectFn - Function(externalEvent) to get the event's title/subject
 * @param {Function} getStartFn - Function(externalEvent) to get start datetime string
 * @param {Function} getEndFn - Function(externalEvent) to get end datetime string
 * @param {Function} isCancelledFn - Function(externalEvent) to check if cancelled
 * @returns {Promise<array>} Array of matching external events
 */
export const findMatchingExternalEvents = async (
  fcEvent,
  calendarId,
  { getEventsFn, getRoamUidByExternalIdFn, getSubjectFn, getStartFn, getEndFn, isCancelledFn, getIdFn }
) => {
  try {
    const eventDate = new Date(fcEvent.start);
    const startOfDay = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate());
    const endOfDay = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate() + 1);

    const externalEvents = await getEventsFn(calendarId, startOfDay, endOfDay);

    const roamEventForComparison = {
      id: fcEvent.id,
      summary: fcEvent.title,
      start: fcEvent.start,
      end: fcEvent.end,
    };

    const matches = externalEvents.filter((extEvent) => {
      if (isCancelledFn(extEvent)) return false;

      const externalId = getIdFn(extEvent);
      const existingRoamUid = getRoamUidByExternalIdFn(externalId);
      if (existingRoamUid) return false;

      const extForComparison = {
        id: externalId,
        summary: getSubjectFn(extEvent),
        start: getStartFn(extEvent),
        end: getEndFn(extEvent),
      };

      return areEventsDuplicate(roamEventForComparison, extForComparison);
    });

    return matches;
  } catch (error) {
    console.error("[Sync] Error finding matching events:", error);
    return [];
  }
};
