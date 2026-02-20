/**
 * Outlook Sync Service - Handles two-way synchronization between Roam and Outlook Calendar
 *
 * Independent sync orchestrator mirroring syncService.js but using Outlook APIs.
 * Reuses existing provider-agnostic services: syncLockService, deduplicationService.
 */

import {
  getOutlookEvents,
  createOutlookEvent,
  updateOutlookEvent as updateOutlookEventApi,
  deleteOutlookEvent as deleteOutlookEventApi,
  getOutlookConnectedCalendars,
  updateOutlookConnectedCalendar,
} from "./outlookCalendarService";

import {
  loadOutlookSyncMetadata,
  getOutlookSyncMetadata,
  saveOutlookSyncMetadata,
  updateOutlookSyncMetadata,
  deleteOutlookSyncMetadata,
  getRoamUidByOutlookId,
  createOutlookSyncMetadata,
  OutlookSyncStatus,
  determineOutlookSyncStatus,
} from "../models/OutlookSyncMetadata";

import {
  fcEventToOutlookEvent,
  outlookEventToFCEvent,
  findOutlookCalendarForEvent,
  outlookEventToRoamContent,
  mergeOutlookDataToFCEvent,
  cleanTitleForOutlook,
} from "../util/outlookMapping";

import {
  getBlockContentByUid,
  updateBlock,
  createChildBlock,
  deleteBlock,
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
 * Helper to detect if a block content contains TODO marker
 */
const hasTodoMarker = (content) => {
  return content && content.includes("{{[[TODO]]}}");
};

/**
 * Helper to extract end date from Outlook event for storage
 */
const getOutlookEventEndDateString = (outlookEvent) => {
  if (!outlookEvent) return null;

  const endDateTime = outlookEvent.end?.dateTime;
  if (!endDateTime) return null;

  const endDate = new Date(endDateTime);

  // For all-day events, Outlook end date is exclusive (next day), subtract 1 day
  if (outlookEvent.isAllDay) {
    endDate.setDate(endDate.getDate() - 1);
  }

  return endDate.toISOString().split("T")[0];
};

/**
 * Sync result object
 */
export const createOutlookSyncResult = () => ({
  imported: [],
  exported: [],
  updated: [],
  conflicts: [],
  errors: [],
  deletedFromOutlook: [],
  deletedFromRoam: [],
});

/**
 * Sync a single Roam event to Outlook Calendar
 * @param {string} roamUid - Roam block UID
 * @param {object} fcEvent - FullCalendar event object
 * @param {string} calendarId - Target Outlook Calendar ID
 * @returns {object} Sync result
 */
export const syncEventToOutlook = async (roamUid, fcEvent, calendarId) => {
  if (!acquireSyncLock(roamUid)) {
    return { success: false, error: "Already syncing", skipped: true };
  }

  try {
    const metadata = getOutlookSyncMetadata(roamUid);
    const outlookEvent = fcEventToOutlookEvent(fcEvent, calendarId, roamUid);

    if (metadata && metadata.outlookId) {
      // Update existing event
      const result = await updateOutlookEventApi(metadata.outlookId, outlookEvent);

      await updateOutlookSyncMetadata(roamUid, {
        outlookUpdated: result.lastModifiedDateTime,
        changeKey: result.changeKey,
        roamUpdated: Date.now(),
        lastSync: Date.now(),
      });

      return { success: true, action: "updated", outlookId: result.id };
    } else {
      // Double-check metadata
      const freshMetadata = getOutlookSyncMetadata(roamUid);
      if (freshMetadata && freshMetadata.outlookId) {
        const result = await updateOutlookEventApi(
          freshMetadata.outlookId,
          outlookEvent
        );

        await updateOutlookSyncMetadata(roamUid, {
          outlookUpdated: result.lastModifiedDateTime,
          changeKey: result.changeKey,
          roamUpdated: Date.now(),
          lastSync: Date.now(),
        });

        return { success: true, action: "updated", outlookId: result.id };
      }

      // Create new event
      const result = await createOutlookEvent(calendarId, outlookEvent);

      const blockContent = getBlockContentByUid(roamUid);
      const isTodo = hasTodoMarker(blockContent);

      const eventEndDate = fcEvent.end
        ? new Date(fcEvent.end).toISOString().split("T")[0]
        : new Date(fcEvent.start).toISOString().split("T")[0];

      const hadOriginalTimeRange = parseRange(blockContent) !== null;

      await saveOutlookSyncMetadata(
        roamUid,
        createOutlookSyncMetadata({
          outlookId: result.id,
          outlookCalendarId: calendarId,
          changeKey: result.changeKey,
          outlookUpdated: result.lastModifiedDateTime,
          roamUpdated: Date.now(),
          lastSync: Date.now(),
          eventEndDate,
          isTodo,
          hadOriginalTimeRange,
        })
      );

      return { success: true, action: "created", outlookId: result.id };
    }
  } catch (error) {
    console.error("[OutlookSync] Error syncing event:", error);
    return { success: false, error: error.message };
  } finally {
    releaseSyncLock(roamUid);
  }
};

/**
 * Delete a synced event from Outlook Calendar
 */
export const deleteEventFromOutlook = async (roamUid) => {
  try {
    const metadata = getOutlookSyncMetadata(roamUid);

    if (metadata && metadata.outlookId) {
      await deleteOutlookEventApi(metadata.outlookId);
      await deleteOutlookSyncMetadata(roamUid);
      return { success: true };
    }

    return { success: false, error: "No sync metadata found" };
  } catch (error) {
    console.error("[OutlookSync] Error deleting event:", error);
    return { success: false, error: error.message };
  }
};

/**
 * Fetch events from Outlook Calendar for a date range
 */
export const fetchOutlookEventsForRange = async (
  calendarId,
  startDate,
  endDate,
  calendarConfig
) => {
  try {
    const outlookEvents = await getOutlookEvents(
      calendarId,
      startDate,
      endDate
    );

    return outlookEvents.map((outlookEvent) =>
      outlookEventToFCEvent(outlookEvent, calendarConfig)
    );
  } catch (error) {
    console.error("[OutlookSync] Error fetching events:", error);
    return [];
  }
};

/**
 * Perform incremental sync for an Outlook calendar
 */
export const incrementalOutlookSync = async (calendarConfig) => {
  const result = createOutlookSyncResult();
  const { id: calendarId, lastSyncTime } = calendarConfig;

  try {
    const now = new Date();
    const timeMin = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const timeMax = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);

    const options = {};
    if (lastSyncTime) {
      options.updatedMin = new Date(lastSyncTime);
    }

    const outlookEvents = await getOutlookEvents(
      calendarId,
      timeMin,
      timeMax,
      options
    );

    for (const outlookEvent of outlookEvents) {
      try {
        await processOutlookEvent(outlookEvent, calendarConfig, result);
      } catch (error) {
        result.errors.push({
          eventId: outlookEvent.id,
          error: error.message,
        });
      }
    }

    updateOutlookConnectedCalendar(calendarId, {
      lastSyncTime: Date.now(),
    });
  } catch (error) {
    console.error("[OutlookSync] Error during incremental sync:", error);
    result.errors.push({ error: error.message });
  }

  return result;
};

/**
 * Process a single Outlook event during sync
 */
const processOutlookEvent = async (outlookEvent, calendarConfig, result) => {
  const roamUid = getRoamUidByOutlookId(outlookEvent.id);

  if (outlookEvent.isCancelled) {
    if (roamUid) {
      result.deletedFromOutlook.push({
        outlookId: outlookEvent.id,
        roamUid,
      });
      await deleteOutlookSyncMetadata(roamUid);
    }
    return;
  }

  if (!roamUid) {
    result.imported.push({
      outlookEvent,
      calendarConfig,
      status: "pending",
    });
    return;
  }

  const metadata = getOutlookSyncMetadata(roamUid);
  const syncStatus = determineOutlookSyncStatus(metadata, outlookEvent);

  switch (syncStatus) {
    case OutlookSyncStatus.CONFLICT:
      result.conflicts.push({
        roamUid,
        outlookEvent,
        metadata,
        calendarConfig,
      });
      break;

    case OutlookSyncStatus.PENDING: {
      const outlookUpdated = new Date(
        outlookEvent.lastModifiedDateTime
      ).getTime();
      const roamUpdated = metadata.roamUpdated || metadata.lastSync;

      if (outlookUpdated > roamUpdated) {
        result.updated.push({
          direction: "outlook-to-roam",
          roamUid,
          outlookEvent,
          calendarConfig,
        });
      } else {
        result.updated.push({
          direction: "roam-to-outlook",
          roamUid,
          metadata,
          calendarConfig,
        });
      }
      break;
    }

    case OutlookSyncStatus.SYNCED:
      break;
  }
};

/**
 * Apply sync results - import Outlook events to Roam
 */
export const applyOutlookImport = async (outlookEvent, calendarConfig) => {
  try {
    const isAllDay = outlookEvent.isAllDay === true;
    let eventStart;
    if (isAllDay) {
      eventStart = outlookEvent.start.dateTime.split("T")[0];
    } else {
      eventStart = outlookEvent.start.dateTime;
    }
    const startDate = new Date(eventStart);
    const dnpUid = window.roamAlphaAPI.util.dateToPageUid(startDate);

    const content = outlookEventToRoamContent(outlookEvent, calendarConfig);
    const newBlockUid = await createChildBlock(dnpUid, content);

    if (newBlockUid) {
      // Handle multi-day events
      const outlookEndDate = outlookEvent.end?.dateTime;
      if (outlookEndDate && rangeEndAttribute) {
        let endDateObj = new Date(outlookEndDate);

        if (isAllDay) {
          endDateObj = new Date(endDateObj.getTime() - 24 * 60 * 60 * 1000);
        }

        if (startDate.toDateString() !== endDateObj.toDateString()) {
          const endDateStr =
            window.roamAlphaAPI.util.dateToPageTitle(endDateObj);
          const endBlockContent = `${rangeEndAttribute}:: [[${endDateStr}]]`;
          await createChildBlock(newBlockUid, endBlockContent, "first");
        }
      }

      const isTodo = hasTodoMarker(content);
      const eventEndDate = getOutlookEventEndDateString(outlookEvent);

      let hadOriginalTimeRange = false;
      if (!isAllDay && outlookEvent.start?.dateTime && outlookEvent.end?.dateTime) {
        const sd = new Date(outlookEvent.start.dateTime);
        const ed = new Date(outlookEvent.end.dateTime);
        const durationMs = ed.getTime() - sd.getTime();
        hadOriginalTimeRange = durationMs !== 3600000;
      }

      await saveOutlookSyncMetadata(
        newBlockUid,
        createOutlookSyncMetadata({
          outlookId: outlookEvent.id,
          outlookCalendarId: calendarConfig.id,
          changeKey: outlookEvent.changeKey,
          outlookUpdated: outlookEvent.lastModifiedDateTime,
          roamUpdated: Date.now(),
          eventEndDate,
          isTodo,
          hadOriginalTimeRange,
        })
      );
    }

    return { success: true, roamUid: newBlockUid };
  } catch (error) {
    console.error("[OutlookSync] Error importing event:", error);
    return { success: false, error: error.message };
  }
};

/**
 * Apply sync results - update Roam from Outlook
 */
export const applyOutlookToRoamUpdate = async (
  roamUid,
  outlookEvent,
  calendarConfig
) => {
  try {
    if (!isExistingNode(roamUid)) {
      console.log(
        "[OutlookSync] Roam block no longer exists, cleaning up:",
        roamUid
      );
      await deleteOutlookSyncMetadata(roamUid);
      return await applyOutlookImport(outlookEvent, calendarConfig);
    }

    const metadata = getOutlookSyncMetadata(roamUid);
    const hadOriginalTimeRange = metadata?.hadOriginalTimeRange || false;

    const currentRoamContent = getBlockContentByUid(roamUid);
    const newContent = outlookEventToRoamContent(
      outlookEvent,
      calendarConfig,
      hadOriginalTimeRange
    );

    const triggerTags = [
      ...(calendarConfig.triggerTags || []),
      "Outlook calendar",
    ];
    const cleanedRoamTitle = cleanTitleForOutlook(
      currentRoamContent,
      triggerTags
    );
    const cleanedOutlookTitle = cleanTitleForOutlook(newContent, triggerTags);

    if (cleanedRoamTitle !== cleanedOutlookTitle) {
      await updateBlock(roamUid, newContent);
    }

    // Handle date changes and block location
    const isAllDay = outlookEvent.isAllDay === true;
    let outlookStartDate;
    if (isAllDay) {
      outlookStartDate = outlookEvent.start.dateTime.split("T")[0];
    } else {
      outlookStartDate = outlookEvent.start.dateTime;
    }
    const newEventDate = new Date(outlookStartDate);
    const newDnpUid = window.roamAlphaAPI.util.dateToPageUid(newEventDate);

    // Handle child blocks with start/end dates
    const { startBlock, endBlock } = findDateChildBlocks(roamUid);

    let outlookStartDateObj = new Date(outlookStartDate);
    let outlookEndDateObj = outlookEvent.end?.dateTime
      ? new Date(outlookEvent.end.dateTime)
      : null;

    if (isAllDay && outlookEndDateObj) {
      outlookEndDateObj = new Date(
        outlookEndDateObj.getTime() - 24 * 60 * 60 * 1000
      );
    }

    const isMultiDayEvent =
      outlookEndDateObj &&
      outlookStartDateObj.toDateString() !== outlookEndDateObj.toDateString();

    if (startBlock) {
      await updateChildBlockDate(
        startBlock.uid,
        startBlock.content,
        outlookStartDateObj
      );
    }

    if (endBlock) {
      if (isMultiDayEvent && outlookEndDateObj) {
        await updateChildBlockDate(
          endBlock.uid,
          endBlock.content,
          outlookEndDateObj
        );
      }
    } else if (isMultiDayEvent && outlookEndDateObj && rangeEndAttribute) {
      const endDateStr =
        window.roamAlphaAPI.util.dateToPageTitle(outlookEndDateObj);
      const endBlockContent = `${rangeEndAttribute}:: [[${endDateStr}]]`;
      await createChildBlock(roamUid, endBlockContent, "first");
    }

    // Move block to correct DNP if date changed
    const currentParentUid = getParentBlock(roamUid);
    if (currentParentUid) {
      let currentDnpUid = getParentBlock(currentParentUid);
      if (!currentDnpUid) {
        currentDnpUid = currentParentUid;
      }

      const currentDateFromUid = window.roamAlphaAPI.util.pageTitleToDate(
        window.roamAlphaAPI.pull("[:node/title]", [
          ":block/uid",
          currentDnpUid,
        ])?.[":node/title"]
      );

      if (currentDateFromUid) {
        const currentDnpDateStr =
          window.roamAlphaAPI.util.dateToPageUid(currentDateFromUid);

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
    }

    const updatedContent = getBlockContentByUid(roamUid);
    const isTodo = hasTodoMarker(updatedContent);
    const eventEndDate = getOutlookEventEndDateString(outlookEvent);

    await updateOutlookSyncMetadata(roamUid, {
      outlookUpdated: outlookEvent.lastModifiedDateTime,
      changeKey: outlookEvent.changeKey,
      roamUpdated: Date.now(),
      lastSync: Date.now(),
      eventEndDate,
      isTodo,
    });

    return { success: true };
  } catch (error) {
    console.error("[OutlookSync] Error updating Roam from Outlook:", error);
    return { success: false, error: error.message };
  }
};

/**
 * Find child blocks containing date information
 */
const findDateChildBlocks = (parentUid) => {
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
 * Update a child block's date reference
 */
const updateChildBlockDate = async (blockUid, currentContent, newDate) => {
  const newRoamDate = window.roamAlphaAPI.util.dateToPageTitle(newDate);
  roamDateRegex.lastIndex = 0;
  const matchingDates = currentContent.match(roamDateRegex);

  if (matchingDates && matchingDates.length) {
    const currentDateStr = matchingDates[0]
      .replace("[[", "")
      .replace("]]", "");
    const newContent = currentContent.replace(currentDateStr, newRoamDate);
    await updateBlock(blockUid, newContent);
  }
};

/**
 * Full sync for all connected Outlook calendars
 */
export const fullOutlookSync = async () => {
  const calendars = getOutlookConnectedCalendars();
  const results = [];

  for (const calendar of calendars) {
    if (!calendar.syncEnabled) continue;

    const result = await incrementalOutlookSync(calendar);
    results.push({
      calendarId: calendar.id,
      calendarName: calendar.name,
      ...result,
    });
  }

  return results;
};

/**
 * Format event date/time for display in toast
 */
const formatEventDateTime = (startDateTime, endDateTime) => {
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
        result +=
          " - " + dateFormat.format(end) + " at " + timeFormat.format(end);
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
 * Show Outlook sync result toast notification
 */
export const showOutlookSyncResultToast = (result, blockUid) => {
  const toaster = Toaster.create({ position: Position.TOP });

  if (result.success) {
    const blockContent = getBlockContentByUid(blockUid) || "Event";
    const actionText =
      result.action === "created" ? "created in" : "updated in";
    const eventTitle =
      blockContent.length > 50
        ? blockContent.substring(0, 47) + "..."
        : blockContent;

    const dateTimeStr = formatEventDateTime(result.eventStart, result.eventEnd);
    const dateTimeInfo = dateTimeStr ? ` (${dateTimeStr})` : "";

    toaster.show({
      message: `"${eventTitle}" ${actionText} Outlook: ${result.calendarName}${dateTimeInfo}`,
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
 * Sync a block to the default Outlook Calendar
 * Called from block context menu or command palette
 */
export const syncBlockToDefaultOutlookCalendar = async (blockContextOrUid) => {
  try {
    const blockUid =
      typeof blockContextOrUid === "string"
        ? blockContextOrUid
        : blockContextOrUid["block-uid"];

    const blockContent = getBlockContentByUid(blockUid);

    if (!blockContent) {
      return {
        success: false,
        error: "Block not found or empty",
      };
    }

    const calendars = getOutlookConnectedCalendars();
    if (!calendars || calendars.length === 0) {
      return {
        success: false,
        error: "No Outlook Calendar connected. Please configure Outlook Calendar first.",
      };
    }

    const defaultCalendar = calendars.find(
      (cal) => cal.syncEnabled && cal.syncDirection !== "import"
    );
    if (!defaultCalendar) {
      return {
        success: false,
        error: "No Outlook calendar available for sync. Please enable sync for at least one calendar.",
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
      start: rangeInfo
        ? `${eventDateStr}T${rangeInfo.range.start}`
        : eventDateStr,
      end:
        rangeInfo && rangeInfo.range.end
          ? `${eventDateStr}T${rangeInfo.range.end}`
          : null,
      extendedProps: {
        eventTags: [],
      },
    };

    // Add calendar tag to block if not present
    if (!blockHasCalendarTag(blockUid, defaultCalendar)) {
      let tagToAdd = null;
      if (
        defaultCalendar.triggerTags &&
        defaultCalendar.triggerTags.length > 0
      ) {
        tagToAdd = defaultCalendar.triggerTags[0];
      } else if (defaultCalendar.displayName) {
        tagToAdd = defaultCalendar.displayName;
      }

      if (tagToAdd) {
        await addTagToBlock(blockUid, tagToAdd);
      }
    }

    const result = await syncEventToOutlook(
      blockUid,
      fcEvent,
      defaultCalendar.id
    );

    if (result.success) {
      return {
        success: true,
        action: result.action,
        calendarName: defaultCalendar.name,
        outlookId: result.outlookId,
        eventStart: fcEvent.start,
        eventEnd: fcEvent.end,
      };
    } else if (result.skipped) {
      return {
        success: false,
        error: "Sync already in progress for this block",
      };
    } else {
      let errorMessage = result.error || "Unknown error";

      if (
        errorMessage.includes("Failed to fetch") ||
        errorMessage.includes("NetworkError") ||
        errorMessage.includes("network")
      ) {
        errorMessage =
          "Unable to connect to Outlook Calendar. Please check your internet connection.";
      } else if (
        errorMessage.includes("401") ||
        errorMessage.includes("Unauthorized")
      ) {
        errorMessage =
          "Outlook Calendar authentication expired. Please reconnect your calendar.";
      } else if (
        errorMessage.includes("403") ||
        errorMessage.includes("Forbidden")
      ) {
        errorMessage =
          "Permission denied. Please check your Outlook Calendar permissions.";
      }

      return {
        success: false,
        error: errorMessage,
      };
    }
  } catch (error) {
    console.error("[OutlookSync] Error syncing block:", error);

    let errorMessage = error.message;
    if (
      errorMessage.includes("Failed to fetch") ||
      errorMessage.includes("NetworkError")
    ) {
      errorMessage =
        "Unable to connect to Outlook Calendar. Please check your internet connection.";
    }

    return {
      success: false,
      error: errorMessage,
    };
  }
};

/**
 * Find matching Outlook events for a Roam event
 */
export const findMatchingOutlookEvents = async (fcEvent, calendarId) => {
  try {
    const eventDate = new Date(fcEvent.start);
    const startOfDay = new Date(
      eventDate.getFullYear(),
      eventDate.getMonth(),
      eventDate.getDate()
    );
    const endOfDay = new Date(
      eventDate.getFullYear(),
      eventDate.getMonth(),
      eventDate.getDate() + 1
    );

    const outlookEvents = await getOutlookEvents(
      calendarId,
      startOfDay,
      endOfDay
    );

    const roamEventForComparison = {
      id: fcEvent.id,
      summary: fcEvent.title,
      start: fcEvent.start,
      end: fcEvent.end,
    };

    const matches = outlookEvents.filter((outlookEvent) => {
      if (outlookEvent.isCancelled) return false;

      const existingRoamUid = getRoamUidByOutlookId(outlookEvent.id);
      if (existingRoamUid) return false;

      // Compare using subject field for Outlook events
      const outlookForComparison = {
        id: outlookEvent.id,
        summary: outlookEvent.subject,
        start: outlookEvent.start?.dateTime,
        end: outlookEvent.end?.dateTime,
      };

      return areEventsDuplicate(roamEventForComparison, outlookForComparison);
    });

    return matches;
  } catch (error) {
    console.error("[OutlookSync] Error finding matching events:", error);
    return [];
  }
};

/**
 * Link a Roam event to an existing Outlook event
 */
export const linkEventToExistingOutlook = async (
  roamUid,
  fcEvent,
  existingOutlookEvent,
  calendarId
) => {
  if (!acquireSyncLock(roamUid)) {
    return { success: false, error: "Already syncing", skipped: true };
  }

  try {
    const metadata = getOutlookSyncMetadata(roamUid);
    if (metadata && metadata.outlookId) {
      return { success: false, error: "Event already synced" };
    }

    // Update Outlook event description with Roam link
    const graphName = window.roamAlphaAPI?.graph?.name;
    let bodyContent = existingOutlookEvent.body?.content || "";

    bodyContent = bodyContent.replace(/\n*---\nRoam block:.*$/s, "").trim();

    if (graphName) {
      const roamLink = `https://roamresearch.com/#/app/${graphName}/page/${roamUid}`;
      bodyContent += `\n\n---\nRoam block: ${roamLink}`;
    }

    await updateOutlookEventApi(existingOutlookEvent.id, {
      subject: existingOutlookEvent.subject,
      body: {
        contentType: existingOutlookEvent.body?.contentType || "text",
        content: bodyContent,
      },
      start: existingOutlookEvent.start,
      end: existingOutlookEvent.end,
    });

    const eventEndDate = fcEvent.end
      ? new Date(fcEvent.end).toISOString().split("T")[0]
      : new Date(fcEvent.start).toISOString().split("T")[0];

    const blockContent = getBlockContentByUid(roamUid);
    const isTodo = hasTodoMarker(blockContent);
    const hadOriginalTimeRange = parseRange(blockContent) !== null;

    await saveOutlookSyncMetadata(
      roamUid,
      createOutlookSyncMetadata({
        outlookId: existingOutlookEvent.id,
        outlookCalendarId: calendarId,
        changeKey: existingOutlookEvent.changeKey,
        outlookUpdated: existingOutlookEvent.lastModifiedDateTime,
        roamUpdated: Date.now(),
        lastSync: Date.now(),
        eventEndDate,
        isTodo,
        hadOriginalTimeRange,
      })
    );

    return {
      success: true,
      action: "linked",
      outlookId: existingOutlookEvent.id,
    };
  } catch (error) {
    console.error("[OutlookSync] Error linking event:", error);
    return { success: false, error: error.message };
  } finally {
    releaseSyncLock(roamUid);
  }
};
