/**
 * OutlookSyncMetadata - Handles sync metadata storage for Outlook Calendar
 *
 * Sync metadata is stored in extension storage, NOT in Roam blocks.
 * Maps Roam block UIDs to their corresponding Outlook event IDs.
 */

import { extensionStorage } from "..";
import { removeTagsFromBlock, isExistingNode } from "../util/roamApi";

const STORAGE_KEY = "outlook-sync-metadata";

// In-memory cache of sync metadata
let syncMetadataCache = null;

/**
 * Get all trigger tags from all connected Outlook calendars
 * @returns {string[]} Array of all trigger tags (including displayNames)
 */
const getAllOutlookTriggerTags = () => {
  // Lazy import to avoid circular dependency
  const { getOutlookConnectedCalendars } = require("../services/outlookCalendarService");
  const calendars = getOutlookConnectedCalendars();
  const allTags = [];

  for (const calendar of calendars) {
    if (calendar.displayName) {
      allTags.push(calendar.displayName);
    }
    if (calendar.triggerTags && calendar.triggerTags.length > 0) {
      allTags.push(...calendar.triggerTags);
    }
  }

  return [...new Set(allTags)];
};

/**
 * Remove all Outlook calendar trigger tags from a Roam block
 * @param {string} roamUid - Roam block UID
 */
const removeOutlookTriggerTagsFromBlock = (roamUid) => {
  if (!isExistingNode(roamUid)) {
    return;
  }

  const allTriggerTags = getAllOutlookTriggerTags();
  if (allTriggerTags.length > 0) {
    removeTagsFromBlock(roamUid, allTriggerTags);
  }
};

/**
 * Sync metadata structure for a single event
 */
export const createOutlookSyncMetadata = ({
  outlookId,
  outlookCalendarId,
  changeKey = null,
  outlookUpdated = null,
  roamUpdated = null,
  lastSync = Date.now(),
  eventEndDate = null,
  isTodo = false,
  hadOriginalTimeRange = false,
}) => ({
  outlookId,
  outlookCalendarId,
  changeKey,
  outlookUpdated,
  roamUpdated,
  lastSync,
  eventEndDate,
  isTodo,
  hadOriginalTimeRange,
});

/**
 * Load all sync metadata from storage
 */
export const loadOutlookSyncMetadata = () => {
  if (syncMetadataCache !== null) {
    return syncMetadataCache;
  }

  try {
    const stored = extensionStorage.get(STORAGE_KEY);
    if (typeof stored === "string") {
      syncMetadataCache = JSON.parse(stored);
    } else {
      syncMetadataCache = stored || {};
    }
    console.log(
      `[OutlookSync] Loaded ${Object.keys(syncMetadataCache).length} entries from storage`
    );
    return syncMetadataCache;
  } catch (error) {
    console.error("[OutlookSync] Failed to load sync metadata:", error);
    syncMetadataCache = {};
    return syncMetadataCache;
  }
};

/**
 * Save all sync metadata to storage
 */
const persistOutlookSyncMetadata = () => {
  try {
    const serialized = JSON.stringify(syncMetadataCache);
    extensionStorage.set(STORAGE_KEY, serialized);
  } catch (error) {
    console.error("[OutlookSync] Failed to persist sync metadata:", error);
  }
};

/**
 * Get sync metadata for a specific Roam block
 * @param {string} roamUid - Roam block UID
 * @returns {object|null} Sync metadata or null if not found
 */
export const getOutlookSyncMetadata = (roamUid) => {
  const allMetadata = loadOutlookSyncMetadata();
  return allMetadata[roamUid] || null;
};

/**
 * Get Roam UID by Outlook event ID
 * @param {string} outlookId - Outlook event ID
 * @returns {string|null} Roam block UID or null if not found
 */
export const getRoamUidByOutlookId = (outlookId) => {
  const allMetadata = loadOutlookSyncMetadata();
  for (const [roamUid, metadata] of Object.entries(allMetadata)) {
    if (metadata.outlookId === outlookId) {
      return roamUid;
    }
  }
  return null;
};

/**
 * Save sync metadata for a Roam block
 * @param {string} roamUid - Roam block UID
 * @param {object} metadata - Sync metadata
 */
export const saveOutlookSyncMetadata = async (roamUid, metadata) => {
  loadOutlookSyncMetadata();
  syncMetadataCache[roamUid] = metadata;
  persistOutlookSyncMetadata();
  return roamUid;
};

/**
 * Update specific fields in sync metadata
 * @param {string} roamUid - Roam block UID
 * @param {object} updates - Fields to update
 */
export const updateOutlookSyncMetadata = async (roamUid, updates) => {
  const existing = getOutlookSyncMetadata(roamUid);

  if (existing) {
    const updatedMetadata = { ...existing, ...updates };
    await saveOutlookSyncMetadata(roamUid, updatedMetadata);
    return updatedMetadata;
  }

  return null;
};

/**
 * Delete sync metadata for a Roam block
 * @param {string} roamUid - Roam block UID
 */
export const deleteOutlookSyncMetadata = async (roamUid) => {
  loadOutlookSyncMetadata();

  if (syncMetadataCache[roamUid]) {
    delete syncMetadataCache[roamUid];
    persistOutlookSyncMetadata();
    return true;
  }

  return false;
};

/**
 * Check if a Roam block is synced with Outlook
 * @param {string} roamUid - Roam block UID
 */
export const isOutlookSynced = (roamUid) => {
  return getOutlookSyncMetadata(roamUid) !== null;
};

/**
 * Get the Outlook event ID from a Roam block's metadata
 * @param {string} roamUid - Roam block UID
 */
export const getOutlookIdFromEvent = (roamUid) => {
  const metadata = getOutlookSyncMetadata(roamUid);
  return metadata ? metadata.outlookId : null;
};

/**
 * Get all synced events for a specific Outlook calendar
 * @param {string} calendarId - Outlook Calendar ID
 */
export const getSyncedEventsForOutlookCalendar = (calendarId) => {
  const allMetadata = loadOutlookSyncMetadata();
  const result = {};

  for (const [roamUid, metadata] of Object.entries(allMetadata)) {
    if (metadata.outlookCalendarId === calendarId) {
      result[roamUid] = metadata;
    }
  }

  return result;
};

/**
 * Clear all Outlook sync metadata
 * Also removes trigger tags from all synced blocks to prevent auto-resync.
 */
export const clearAllOutlookSyncMetadata = () => {
  loadOutlookSyncMetadata();

  for (const roamUid of Object.keys(syncMetadataCache)) {
    removeOutlookTriggerTagsFromBlock(roamUid);
  }

  syncMetadataCache = {};
  persistOutlookSyncMetadata();
  console.log(
    `[OutlookSync] Cleared all sync metadata and removed trigger tags`
  );
};

/**
 * Sync status types
 */
export const OutlookSyncStatus = {
  SYNCED: "synced",
  PENDING: "pending",
  CONFLICT: "conflict",
  LOCAL_ONLY: "local-only",
  OUTLOOK_ONLY: "outlook-only",
};

/**
 * Determine sync status by comparing timestamps
 */
export const determineOutlookSyncStatus = (metadata, outlookEvent) => {
  if (!metadata) {
    return OutlookSyncStatus.LOCAL_ONLY;
  }

  if (!outlookEvent) {
    return OutlookSyncStatus.LOCAL_ONLY;
  }

  const outlookUpdated = new Date(outlookEvent.lastModifiedDateTime).getTime();
  const roamUpdated = metadata.roamUpdated || metadata.lastSync;

  // Both modified since last sync
  if (outlookUpdated > metadata.lastSync && roamUpdated > metadata.lastSync) {
    return OutlookSyncStatus.CONFLICT;
  }

  // Outlook is newer
  if (outlookUpdated > roamUpdated) {
    return OutlookSyncStatus.PENDING;
  }

  // Roam is newer
  if (roamUpdated > outlookUpdated) {
    return OutlookSyncStatus.PENDING;
  }

  return OutlookSyncStatus.SYNCED;
};

/**
 * Get storage statistics for Outlook sync metadata
 * @returns {object} { eventCount, todoCount, estimatedBytes }
 */
export const getOutlookStorageStats = () => {
  const allMetadata = loadOutlookSyncMetadata();
  const entries = Object.entries(allMetadata);

  let todoCount = 0;
  for (const [, metadata] of entries) {
    if (metadata.isTodo) {
      todoCount++;
    }
  }

  const estimatedBytes = entries.length * 200;

  return {
    eventCount: entries.length,
    todoCount,
    estimatedBytes,
  };
};

/**
 * Cleanup old sync metadata for past events
 * Removes metadata for events that ended more than N days ago,
 * unless the event still has TODO status.
 * @param {number} daysThreshold - Days after which to cleanup (default: 90 days)
 * @returns {object} { removedCount, keptTodoCount }
 */
export const cleanupOldOutlookMetadata = (daysThreshold = 90) => {
  loadOutlookSyncMetadata();

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const thresholdDate = new Date(
    today.getTime() - daysThreshold * 24 * 60 * 60 * 1000
  );

  let removedCount = 0;
  let keptTodoCount = 0;
  const toRemove = [];

  for (const [roamUid, metadata] of Object.entries(syncMetadataCache)) {
    if (!metadata.eventEndDate) {
      continue;
    }

    const endDate = new Date(metadata.eventEndDate);

    if (endDate < thresholdDate) {
      if (metadata.isTodo) {
        keptTodoCount++;
        continue;
      }

      toRemove.push(roamUid);
    }
  }

  for (const roamUid of toRemove) {
    removeOutlookTriggerTagsFromBlock(roamUid);
    delete syncMetadataCache[roamUid];
    removedCount++;
  }

  if (removedCount > 0) {
    persistOutlookSyncMetadata();
    console.log(
      `[OutlookSync] Cleaned up ${removedCount} old entries, kept ${keptTodoCount} TODOs`
    );
  }

  return { removedCount, keptTodoCount };
};

/**
 * Cleanup ALL past Outlook events (manual cleanup)
 * @returns {object} { removedCount }
 */
export const cleanupAllPastOutlookMetadata = () => {
  loadOutlookSyncMetadata();

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  let removedCount = 0;
  const toRemove = [];

  for (const [roamUid, metadata] of Object.entries(syncMetadataCache)) {
    if (!metadata.eventEndDate) {
      continue;
    }

    const endDate = new Date(metadata.eventEndDate);

    if (endDate < today) {
      toRemove.push(roamUid);
    }
  }

  for (const roamUid of toRemove) {
    removeOutlookTriggerTagsFromBlock(roamUid);
    delete syncMetadataCache[roamUid];
    removedCount++;
  }

  if (removedCount > 0) {
    persistOutlookSyncMetadata();
    console.log(`[OutlookSync] Removed ${removedCount} past event entries`);
  }

  return { removedCount };
};
