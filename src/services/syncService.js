/**
 * Sync Service - Handles two-way synchronization between Roam and Google Calendar
 *
 * Sync algorithm:
 * 1. Fetch GCal events updated since lastSyncTime
 * 2. For each GCal event:
 *    - Find matching Roam event by gCalId
 *    - No match → gcal-only (import prompt)
 *    - Match → compare timestamps:
 *      - gCalUpdated > roamUpdated → Update Roam
 *      - roamUpdated > gCalUpdated → Update GCal
 *      - Both changed → Mark conflict
 * 3. For Roam events with trigger tags:
 *    - Has gCalId → Update GCal
 *    - No gCalId + should sync → Create in GCal
 * 4. Update lastSyncTime
 */

import {
  getEvents,
  createEvent,
  updateEvent as updateGCalEvent,
  deleteEvent as deleteGCalEvent,
  getConnectedCalendars,
  updateConnectedCalendar,
} from "./googleCalendarService";

import {
  loadSyncMetadata,
  getSyncMetadata,
  saveSyncMetadata,
  updateSyncMetadata,
  deleteSyncMetadata,
  getRoamUidByGCalId,
  createSyncMetadata,
  SyncStatus,
  determineSyncStatus,
} from "../models/SyncMetadata";

import {
  fcEventToGCalEvent,
  gcalEventToFCEvent,
  findCalendarForEvent,
  gcalEventToRoamContent,
  mergeGCalDataToFCEvent,
  cleanTitleForGCal,
} from "../util/gcalMapping";

import { enrichEventsWithTaskData } from "./taskService";

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

import {
  acquireSyncLock,
  releaseSyncLock,
} from "./syncLockService";

import { Toaster, Position, Intent } from "@blueprintjs/core";

import { areEventsDuplicate } from "./deduplicationService";

import {
  hasTodoMarker,
  findDateChildBlocks,
  updateChildBlockDate,
  formatEventDateTime,
  showGenericSyncResultToast,
  syncBlockToDefaultProviderCalendar,
  moveBlockToDnpIfNeeded,
  updateDateChildBlocks,
  findMatchingExternalEvents,
} from "./syncHelpers";

/**
 * Helper to extract end date from event for storage
 * Returns ISO date string (YYYY-MM-DD)
 */
const getEventEndDateString = (gcalEvent) => {
  if (!gcalEvent) return null;

  const endDateTime = gcalEvent.end?.dateTime || gcalEvent.end?.date;
  if (!endDateTime) return null;

  const endDate = new Date(endDateTime);

  // For all-day events, GCal end date is exclusive (next day), subtract 1 day
  if (!gcalEvent.end?.dateTime && gcalEvent.end?.date) {
    endDate.setDate(endDate.getDate() - 1);
  }

  return endDate.toISOString().split("T")[0];
};
import { getCalendarUidFromPage } from "../util/data";
import { rangeEndAttribute } from "../index";

/**
 * Sync result object
 */
export const createSyncResult = () => ({
  imported: [],
  exported: [],
  updated: [],
  conflicts: [],
  errors: [],
  deletedFromGCal: [],
  deletedFromRoam: [],
});

/**
 * Sync a single Roam event to Google Calendar
 * @param {string} roamUid - Roam block UID
 * @param {object} fcEvent - FullCalendar event object
 * @param {string} calendarId - Target Google Calendar ID
 * @returns {object} Sync result
 */
export const syncEventToGCal = async (roamUid, fcEvent, calendarId) => {
  // Acquire lock to prevent duplicate syncs across multiple component instances
  if (!acquireSyncLock(roamUid)) {
    return { success: false, error: "Already syncing", skipped: true };
  }

  try {
    const metadata = getSyncMetadata(roamUid);
    const gcalEvent = fcEventToGCalEvent(fcEvent, calendarId, roamUid);
    console.log("[syncEventToGCal] Block UID:", roamUid);
    console.log("[syncEventToGCal] Existing metadata:", metadata);

    if (metadata && metadata.gCalId) {
      // Update existing event
      const result = await updateGCalEvent(calendarId, metadata.gCalId, gcalEvent);

      await updateSyncMetadata(roamUid, {
        gCalUpdated: result.updated,
        etag: result.etag,
        roamUpdated: Date.now(),
        lastSync: Date.now(),
      });

      return { success: true, action: "updated", gCalId: result.id };
    } else {
      // Double-check metadata again in case another instance just created it
      // This handles the race where two instances both saw "no metadata" before acquiring the lock
      const freshMetadata = getSyncMetadata(roamUid);
      if (freshMetadata && freshMetadata.gCalId) {
        // Update instead of create
        const result = await updateGCalEvent(calendarId, freshMetadata.gCalId, gcalEvent);

        await updateSyncMetadata(roamUid, {
          gCalUpdated: result.updated,
          etag: result.etag,
          roamUpdated: Date.now(),
          lastSync: Date.now(),
        });

        return { success: true, action: "updated", gCalId: result.id };
      }

      // Create new event
      console.log("[syncEventToGCal] Creating new event in calendar:", calendarId);
      console.log("[syncEventToGCal] Event data:", gcalEvent);
      const result = await createEvent(calendarId, gcalEvent);
      console.log("[syncEventToGCal] Event created successfully:", result);

      // Check if the Roam block has TODO marker
      const blockContent = getBlockContentByUid(roamUid);
      const isTodo = hasTodoMarker(blockContent);

      // Extract end date for cleanup purposes
      const eventEndDate = fcEvent.end
        ? new Date(fcEvent.end).toISOString().split("T")[0]
        : new Date(fcEvent.start).toISOString().split("T")[0];

      // Detect if original event had a time range (not just start time)
      // Check for time ranges in block content (e.g., "13:00-14:00")
      const hadOriginalTimeRange = parseRange(blockContent) !== null;

      await saveSyncMetadata(
        roamUid,
        createSyncMetadata({
          gCalId: result.id,
          gCalCalendarId: calendarId,
          etag: result.etag,
          gCalUpdated: result.updated,
          roamUpdated: Date.now(),
          lastSync: Date.now(),
          eventEndDate,
          isTodo,
          hadOriginalTimeRange,
        })
      );

      return { success: true, action: "created", gCalId: result.id };
    }
  } catch (error) {
    console.error("Error syncing event to GCal:", error);
    return { success: false, error: error.message };
  } finally {
    // Always release the lock, even if there was an error
    releaseSyncLock(roamUid);
  }
};

/**
 * Delete a synced event from Google Calendar
 * @param {string} roamUid - Roam block UID
 */
export const deleteEventFromGCal = async (roamUid) => {
  try {
    const metadata = getSyncMetadata(roamUid);

    if (metadata && metadata.gCalId) {
      await deleteGCalEvent(metadata.gCalCalendarId, metadata.gCalId);
      await deleteSyncMetadata(roamUid);
      return { success: true };
    }

    return { success: false, error: "No sync metadata found" };
  } catch (error) {
    console.error("Error deleting event from GCal:", error);
    return { success: false, error: error.message };
  }
};

/**
 * Fetch events from Google Calendar for a date range
 * Automatically enriches Google Tasks with their actual notes/description
 * @param {string} calendarId - Calendar ID
 * @param {Date} startDate - Start of range
 * @param {Date} endDate - End of range
 * @param {object} calendarConfig - Calendar configuration
 * @returns {array} Array of FC events
 */
export const fetchGCalEventsForRange = async (
  calendarId,
  startDate,
  endDate,
  calendarConfig
) => {
  try {
    let gcalEvents = await getEvents(calendarId, startDate, endDate);

    // Filter out cancelled events
    gcalEvents = gcalEvents.filter((event) => event.status !== "cancelled");

    // Enrich any Google Tasks with their actual notes/description
    // This replaces the placeholder description with the real task data
    gcalEvents = await enrichEventsWithTaskData(gcalEvents, startDate, endDate);

    return gcalEvents.map((gcalEvent) => gcalEventToFCEvent(gcalEvent, calendarConfig));
  } catch (error) {
    console.error("Error fetching GCal events:", error);
    return [];
  }
};

/**
 * Perform incremental sync for a calendar
 * Fetches only events updated since lastSyncTime
 * @param {object} calendarConfig - Calendar configuration
 * @returns {object} Sync result
 */
export const incrementalSync = async (calendarConfig) => {
  const result = createSyncResult();
  const { id: calendarId, lastSyncTime } = calendarConfig;

  try {
    // Fetch events updated since last sync
    const now = new Date();
    const timeMin = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000); // Last 30 days
    const timeMax = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000); // Next 90 days

    const options = {};
    if (lastSyncTime) {
      options.updatedMin = new Date(lastSyncTime);
      options.showDeleted = true;
    }

    const gcalEvents = await getEvents(calendarId, timeMin, timeMax, options);

    for (const gcalEvent of gcalEvents) {
      try {
        await processGCalEvent(gcalEvent, calendarConfig, result);
      } catch (error) {
        result.errors.push({
          eventId: gcalEvent.id,
          error: error.message,
        });
      }
    }

    // Update last sync time
    updateConnectedCalendar(calendarId, {
      lastSyncTime: Date.now(),
    });
  } catch (error) {
    console.error("Error during incremental sync:", error);
    result.errors.push({ error: error.message });
  }

  return result;
};

/**
 * Process a single GCal event during sync
 */
const processGCalEvent = async (gcalEvent, calendarConfig, result) => {
  const roamUid = getRoamUidByGCalId(gcalEvent.id);

  if (gcalEvent.status === "cancelled") {
    // Event was deleted in GCal
    if (roamUid) {
      result.deletedFromGCal.push({
        gCalId: gcalEvent.id,
        roamUid,
      });
      await deleteSyncMetadata(roamUid);
    }
    return;
  }

  if (!roamUid) {
    // New event from GCal, not yet imported
    result.imported.push({
      gCalEvent: gcalEvent,
      calendarConfig,
      status: "pending",
    });
    return;
  }

  // Event exists in both systems - check for conflicts
  const metadata = getSyncMetadata(roamUid);
  const syncStatus = determineSyncStatus(metadata, gcalEvent);

  switch (syncStatus) {
    case SyncStatus.CONFLICT:
      result.conflicts.push({
        roamUid,
        gCalEvent,
        metadata,
        calendarConfig,
      });
      break;

    case SyncStatus.PENDING:
      const gCalUpdated = new Date(gcalEvent.updated).getTime();
      const roamUpdated = metadata.roamUpdated || metadata.lastSync;

      if (gCalUpdated > roamUpdated) {
        // GCal is newer - update Roam
        result.updated.push({
          direction: "gcal-to-roam",
          roamUid,
          gCalEvent,
          calendarConfig,
        });
      } else {
        // Roam is newer - update GCal
        result.updated.push({
          direction: "roam-to-gcal",
          roamUid,
          metadata,
          calendarConfig,
        });
      }
      break;

    case SyncStatus.SYNCED:
      // No action needed
      break;
  }
};

/**
 * Apply sync results - import GCal events to Roam
 */
export const applyImport = async (gcalEvent, calendarConfig) => {
  try {
    const eventStart = gcalEvent.start.dateTime || gcalEvent.start.date;
    const startDate = new Date(eventStart);
    const dnpUid = window.roamAlphaAPI.util.dateToPageUid(startDate);

    const content = gcalEventToRoamContent(gcalEvent, calendarConfig);
    const newBlockUid = await createChildBlock(dnpUid, content);

    if (newBlockUid) {
      // Check if this is a multi-day event and create end date child block
      const gcalEndDate = gcalEvent.end?.dateTime || gcalEvent.end?.date;
      if (gcalEndDate && rangeEndAttribute) {
        const isAllDayEvent = !gcalEvent.start.dateTime;
        let endDateObj = new Date(gcalEndDate);

        // For all-day events, GCal end date is exclusive (next day)
        if (isAllDayEvent) {
          endDateObj = new Date(endDateObj.getTime() - 24 * 60 * 60 * 1000);
        }

        // Check if it's actually a multi-day event
        if (startDate.toDateString() !== endDateObj.toDateString()) {
          const endDateStr = window.roamAlphaAPI.util.dateToPageTitle(endDateObj);
          const endBlockContent = `${rangeEndAttribute}:: [[${endDateStr}]]`;
          await createChildBlock(newBlockUid, endBlockContent, "first");
        }
      }

      // Check if the imported content has TODO marker
      const isTodo = hasTodoMarker(content);

      // Extract end date for cleanup purposes
      const eventEndDate = getEventEndDateString(gcalEvent);

      // For new imports, check if the event has a non-default duration
      // If it does, mark it as having a time range so future syncs preserve it
      let hadOriginalTimeRange = false;
      if (gcalEvent.start.dateTime && gcalEvent.end?.dateTime) {
        const startDate = new Date(gcalEvent.start.dateTime);
        const endDate = new Date(gcalEvent.end.dateTime);
        const durationMs = endDate.getTime() - startDate.getTime();
        const isDefaultDuration = durationMs === 3600000; // 1 hour
        hadOriginalTimeRange = !isDefaultDuration;
      }

      await saveSyncMetadata(
        newBlockUid,
        createSyncMetadata({
          gCalId: gcalEvent.id,
          gCalCalendarId: calendarConfig.id,
          etag: gcalEvent.etag,
          gCalUpdated: gcalEvent.updated,
          roamUpdated: Date.now(),
          eventEndDate,
          isTodo,
          hadOriginalTimeRange,
        })
      );
    }

    return { success: true, roamUid: newBlockUid };
  } catch (error) {
    console.error("Error importing GCal event:", error);
    return { success: false, error: error.message };
  }
};

// findDateChildBlocks and updateChildBlockDate are now imported from syncHelpers

/**
 * Apply sync results - update Roam from GCal
 */
export const applyGCalToRoamUpdate = async (roamUid, gcalEvent, calendarConfig) => {
  try {
    // Check if the Roam block still exists (it may have been deleted)
    if (!isExistingNode(roamUid)) {
      console.log("Roam block no longer exists, cleaning up stale metadata:", roamUid);
      await deleteSyncMetadata(roamUid);
      // Re-import the event as a new block
      return await applyImport(gcalEvent, calendarConfig);
    }

    // Get metadata to check if original event had a time range
    const metadata = getSyncMetadata(roamUid);
    const hadOriginalTimeRange = metadata?.hadOriginalTimeRange || false;

    // Get current Roam content and proposed new content from GCal
    const currentRoamContent = getBlockContentByUid(roamUid);
    const newContent = gcalEventToRoamContent(gcalEvent, calendarConfig, hadOriginalTimeRange);

    // Compare cleaned titles to detect real changes (excluding trigger tags)
    // Always include "Google calendar" in the list of tags to strip for comparison
    const triggerTags = [...(calendarConfig.triggerTags || []), "Google calendar"];
    const cleanedRoamTitle = cleanTitleForGCal(currentRoamContent, triggerTags);
    const cleanedGCalTitle = cleanTitleForGCal(newContent, triggerTags);

    // Only update the block if there's an actual content change (not just trigger tag difference)
    if (cleanedRoamTitle !== cleanedGCalTitle) {
      await updateBlock(roamUid, newContent);
    }

    // Check if the date changed and move the block if needed
    const gcalStartDate = gcalEvent.start.dateTime || gcalEvent.start.date;
    const gcalEndDate = gcalEvent.end?.dateTime || gcalEvent.end?.date;
    const newEventDate = new Date(gcalStartDate);
    const newDnpUid = window.roamAlphaAPI.util.dateToPageUid(newEventDate);

    // =========================================================================
    // Handle child blocks with start:: and end:: dates (multi-day events)
    // =========================================================================
    const { startBlock, endBlock } = findDateChildBlocks(roamUid);

    // Check if this is a multi-day event in GCal
    const isAllDayEvent = !gcalEvent.start.dateTime;
    let gcalStartDateObj = new Date(gcalStartDate);
    let gcalEndDateObj = gcalEndDate ? new Date(gcalEndDate) : null;

    // For all-day events, GCal end date is exclusive (next day), so we need to subtract 1 day
    // to get the actual last day of the event
    if (isAllDayEvent && gcalEndDateObj) {
      gcalEndDateObj = new Date(gcalEndDateObj.getTime() - 24 * 60 * 60 * 1000);
    }

    const isMultiDayEvent =
      gcalEndDateObj &&
      gcalStartDateObj.toDateString() !== gcalEndDateObj.toDateString();

    // Update start child block if it exists
    if (startBlock) {
      await updateChildBlockDate(startBlock.uid, startBlock.content, gcalStartDateObj);
    }

    // Update end/until child block if it exists
    if (endBlock) {
      if (isMultiDayEvent && gcalEndDateObj) {
        await updateChildBlockDate(endBlock.uid, endBlock.content, gcalEndDateObj);
      }
    } else if (isMultiDayEvent && gcalEndDateObj && rangeEndAttribute) {
      // No end block exists but the event is multi-day - create one
      const endDateStr = window.roamAlphaAPI.util.dateToPageTitle(gcalEndDateObj);
      const endBlockContent = `${rangeEndAttribute}:: [[${endDateStr}]]`;
      await createChildBlock(roamUid, endBlockContent, "first");
    }

    // =========================================================================
    // Handle block location (move to correct DNP if start date changed)
    // =========================================================================
    const currentParentUid = getParentBlock(roamUid);
    if (currentParentUid) {
      // Get the DNP UID that contains the current parent
      // The parent might be a "Calendar" block, so we need to check its parent too
      let currentDnpUid = getParentBlock(currentParentUid);
      // If no grandparent, the parent itself might be the DNP
      if (!currentDnpUid) {
        currentDnpUid = currentParentUid;
      }

      // Check if the event date is different from current location
      const currentDateFromUid = window.roamAlphaAPI.util.pageTitleToDate(
        window.roamAlphaAPI.pull("[:node/title]", [":block/uid", currentDnpUid])?.[":node/title"]
      );

      if (currentDateFromUid) {
        const currentDnpDateStr = window.roamAlphaAPI.util.dateToPageUid(currentDateFromUid);

        if (newDnpUid !== currentDnpDateStr) {
          // Date changed - move the block to the new DNP
          console.log("GCal event date changed, moving block from", currentDnpDateStr, "to", newDnpUid);

          const newCalendarBlockUid = await getCalendarUidFromPage(newDnpUid);
          await window.roamAlphaAPI.moveBlock({
            location: {
              "parent-uid": newCalendarBlockUid,
              order: "last",
            },
            block: { uid: roamUid },
          });

          // Clean up empty calendar block in old location
          deleteBlockIfNoChild(currentParentUid);
        }
      }
    }

    // Update metadata with current TODO status and end date
    const updatedContent = getBlockContentByUid(roamUid);
    const isTodo = hasTodoMarker(updatedContent);
    const eventEndDate = getEventEndDateString(gcalEvent);

    await updateSyncMetadata(roamUid, {
      gCalUpdated: gcalEvent.updated,
      etag: gcalEvent.etag,
      roamUpdated: Date.now(),
      lastSync: Date.now(),
      eventEndDate,
      isTodo,
    });

    return { success: true };
  } catch (error) {
    console.error("Error updating Roam from GCal:", error);
    return { success: false, error: error.message };
  }
};

/**
 * Apply sync results - update GCal from Roam
 */
export const applyRoamToGCalUpdate = async (roamUid, fcEvent, calendarId) => {
  return syncEventToGCal(roamUid, fcEvent, calendarId);
};

/**
 * Resolve a conflict by choosing a side
 * @param {object} conflict - Conflict object from sync result
 * @param {string} resolution - "roam" | "gcal" | "both"
 */
export const resolveConflict = async (conflict, resolution, calendarConfig) => {
  const { roamUid, gCalEvent, metadata } = conflict;

  try {
    switch (resolution) {
      case "gcal":
        // Keep GCal version, update Roam
        return await applyGCalToRoamUpdate(roamUid, gCalEvent, calendarConfig);

      case "roam":
        // Keep Roam version, update GCal
        const blockContent = getBlockContentByUid(roamUid);
        // Would need to reconstruct FC event from block content
        // For now, just update metadata to mark as synced
        await updateSyncMetadata(roamUid, {
          lastSync: Date.now(),
          roamUpdated: Date.now(),
        });
        return { success: true, action: "kept-roam" };

      case "both":
        // Keep both - create duplicate in Roam
        return await applyImport(gCalEvent, calendarConfig);

      default:
        return { success: false, error: "Invalid resolution" };
    }
  } catch (error) {
    console.error("Error resolving conflict:", error);
    return { success: false, error: error.message };
  }
};

/**
 * Full sync for all connected calendars
 */
export const fullSync = async () => {
  const calendars = getConnectedCalendars();
  const results = [];

  for (const calendar of calendars) {
    if (!calendar.syncEnabled) continue;

    const result = await incrementalSync(calendar);
    results.push({
      calendarId: calendar.id,
      calendarName: calendar.name,
      ...result,
    });
  }

  return results;
};

/**
 * Check if a Roam event needs to be synced
 * @param {string} roamUid - Roam block UID
 * @param {object} fcEvent - FullCalendar event
 */
export const checkNeedsSync = (roamUid, fcEvent) => {
  const metadata = getSyncMetadata(roamUid);

  if (!metadata) {
    // Not synced yet - check if has trigger tag
    const calendars = getConnectedCalendars();
    const targetCalendar = findCalendarForEvent(fcEvent, calendars);
    return targetCalendar !== null;
  }

  // Already synced - check if updated since last sync
  return metadata.roamUpdated > metadata.lastSync;
};

/**
 * Get sync status for a Roam event
 */
export const getEventSyncStatus = (roamUid) => {
  const metadata = getSyncMetadata(roamUid);

  if (!metadata) {
    return { synced: false, status: SyncStatus.LOCAL_ONLY };
  }

  return {
    synced: true,
    status: metadata.roamUpdated > metadata.lastSync ? SyncStatus.PENDING : SyncStatus.SYNCED,
    gCalId: metadata.gCalId,
    lastSync: metadata.lastSync,
  };
};

/**
 * Sync a block to the default (first enabled) Google Calendar
 * Called from block context menu or command palette
 * @param {object|string} blockContextOrUid - Block context object or block UID string
 * @returns {object} Result with success status and message
 */
// formatEventDateTime is now imported from syncHelpers

/**
 * Show sync result toast notification.
 * Delegates to shared showGenericSyncResultToast.
 */
export const showSyncResultToast = (result, blockUid) => {
  showGenericSyncResultToast(result, blockUid, "Google Calendar");
};

/**
 * Sync a block to the default Google Calendar.
 * Delegates to shared syncBlockToDefaultProviderCalendar.
 */
export const syncBlockToDefaultCalendar = async (blockContextOrUid) => {
  return syncBlockToDefaultProviderCalendar(blockContextOrUid, {
    getConnectedCalendars,
    syncEventFn: syncEventToGCal,
    providerName: "Google Calendar",
  });
};

/**
 * Find matching GCal events for a Roam event.
 * Delegates to shared findMatchingExternalEvents.
 */
export const findMatchingGCalEvents = async (fcEvent, calendarId) => {
  return findMatchingExternalEvents(fcEvent, calendarId, {
    getEventsFn: getEvents,
    getRoamUidByExternalIdFn: getRoamUidByGCalId,
    getSubjectFn: (e) => e.summary,
    getStartFn: (e) => e.start?.dateTime || e.start?.date,
    getEndFn: (e) => e.end?.dateTime || e.end?.date,
    isCancelledFn: (e) => e.status === "cancelled",
    getIdFn: (e) => e.id,
  });
};

/**
 * Link a Roam event to an existing Google Calendar event (without creating new)
 * @param {string} roamUid - Roam block UID
 * @param {object} fcEvent - FullCalendar event object
 * @param {object} existingGCalEvent - Existing GCal event to link to
 * @param {string} calendarId - Google Calendar ID
 * @returns {Promise<object>} Link result
 */
export const linkEventToExistingGCal = async (roamUid, fcEvent, existingGCalEvent, calendarId) => {
  // Acquire lock to prevent duplicate operations
  if (!acquireSyncLock(roamUid)) {
    return { success: false, error: "Already syncing", skipped: true };
  }

  try {
    // Check if already linked
    const metadata = getSyncMetadata(roamUid);
    if (metadata && metadata.gCalId) {
      return { success: false, error: "Event already synced" };
    }

    // Update the GCal event description to include Roam block link
    const graphName = window.roamAlphaAPI?.graph?.name;
    let description = existingGCalEvent.description || "";

    // Remove any existing Roam block links (in case of re-linking)
    description = description.replace(/\n*---\nRoam block:.*$/s, "").trim();

    if (graphName) {
      const roamLink = `https://roamresearch.com/#/app/${graphName}/page/${roamUid}`;
      description += `\n\n---\nRoam block: ${roamLink}`;
    }

    // Update GCal event with Roam link in description
    // IMPORTANT: Must include start/end/summary to avoid "Missing end time" error
    const updatePayload = {
      summary: existingGCalEvent.summary,
      description: description,
      start: existingGCalEvent.start,
      end: existingGCalEvent.end,
    };

    await updateGCalEvent(calendarId, existingGCalEvent.id, updatePayload);

    // Prepare metadata
    const eventEndDate = fcEvent.end
      ? new Date(fcEvent.end).toISOString().split("T")[0]
      : new Date(fcEvent.start).toISOString().split("T")[0];

    const blockContent = getBlockContentByUid(roamUid);
    const isTodo = hasTodoMarker(blockContent);
    const hadOriginalTimeRange = parseRange(blockContent) !== null;

    // Save sync metadata to establish the link
    await saveSyncMetadata(
      roamUid,
      createSyncMetadata({
        gCalId: existingGCalEvent.id,
        gCalCalendarId: calendarId,
        etag: existingGCalEvent.etag,
        gCalUpdated: existingGCalEvent.updated,
        roamUpdated: Date.now(),
        lastSync: Date.now(),
        eventEndDate,
        isTodo,
        hadOriginalTimeRange,
      })
    );

    console.log(`[Sync] Linked Roam block ${roamUid} to existing GCal event ${existingGCalEvent.id}`);
    return { success: true, action: "linked", gCalId: existingGCalEvent.id };
  } catch (error) {
    console.error("Error linking event to existing GCal:", error);
    return { success: false, error: error.message };
  } finally {
    // Always release the lock
    releaseSyncLock(roamUid);
  }
};

export default {
  syncEventToGCal,
  deleteEventFromGCal,
  fetchGCalEventsForRange,
  incrementalSync,
  fullSync,
  applyImport,
  applyGCalToRoamUpdate,
  applyRoamToGCalUpdate,
  resolveConflict,
  checkNeedsSync,
  getEventSyncStatus,
  createSyncResult,
  syncBlockToDefaultCalendar,
  showSyncResultToast,
  findMatchingGCalEvents,
  linkEventToExistingGCal,
};
