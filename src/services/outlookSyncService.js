/**
 * Outlook Sync Service - Handles two-way synchronization between Roam and Outlook Calendar
 *
 * Independent sync orchestrator using Outlook APIs.
 * Shared helpers are imported from syncHelpers.js to avoid duplication with syncService.js.
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
  parseOutlookDateTime,
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

import {
  hasTodoMarker,
  findDateChildBlocks,
  updateChildBlockDate,
  showGenericSyncResultToast,
  syncBlockToDefaultProviderCalendar,
  findMatchingExternalEvents,
} from "./syncHelpers";

import { getCalendarUidFromPage } from "../util/data";
import { rangeEndAttribute } from "../index";

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

    // Handle child blocks with start/end dates using shared helpers
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
 * Show Outlook sync result toast notification.
 * Delegates to shared showGenericSyncResultToast.
 */
export const showOutlookSyncResultToast = (result, blockUid) => {
  showGenericSyncResultToast(result, blockUid, "Outlook");
};

/**
 * Sync a block to the default Outlook Calendar.
 * Delegates to shared syncBlockToDefaultProviderCalendar.
 */
export const syncBlockToDefaultOutlookCalendar = async (blockContextOrUid) => {
  return syncBlockToDefaultProviderCalendar(blockContextOrUid, {
    getConnectedCalendars: getOutlookConnectedCalendars,
    syncEventFn: syncEventToOutlook,
    providerName: "Outlook Calendar",
  });
};

/**
 * Find matching Outlook events for a Roam event.
 * Delegates to shared findMatchingExternalEvents.
 */
export const findMatchingOutlookEvents = async (fcEvent, calendarId) => {
  return findMatchingExternalEvents(fcEvent, calendarId, {
    getEventsFn: getOutlookEvents,
    getRoamUidByExternalIdFn: getRoamUidByOutlookId,
    getSubjectFn: (e) => e.subject,
    getStartFn: (e) => e.start?.dateTime,
    getEndFn: (e) => e.end?.dateTime,
    isCancelledFn: (e) => e.isCancelled,
    getIdFn: (e) => e.id,
  });
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
