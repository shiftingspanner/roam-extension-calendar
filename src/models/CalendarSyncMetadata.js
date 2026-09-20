/**
 * CalendarSyncMetadata - Shared base for calendar sync metadata storage.
 *
 * Provides a factory function to create provider-specific metadata modules,
 * eliminating duplication between SyncMetadata.js and OutlookSyncMetadata.js.
 *
 * Both Google Calendar and Outlook Calendar store sync metadata with
 * identical patterns: a storage key, in-memory cache, and CRUD operations
 * mapping Roam block UIDs to external event IDs.
 */

import { extensionStorage } from "..";
import { removeTagsFromBlock, isExistingNode } from "../util/roamApi";

/**
 * Sync status types - shared across all providers.
 */
export const SyncStatus = {
  SYNCED: "synced",
  PENDING: "pending",
  CONFLICT: "conflict",
  LOCAL_ONLY: "local-only",
};

/**
 * Determine sync status by comparing timestamps.
 *
 * @param {object} metadata - Sync metadata for the event
 * @param {number} externalUpdatedTimestamp - External event's last-modified timestamp (ms)
 * @param {number} lastSync - Last sync timestamp from metadata
 * @param {number} roamUpdated - Roam block last-modified timestamp from metadata
 * @returns {string} One of the SyncStatus values
 */
export const determineSyncStatus = (metadata, externalUpdatedTimestamp) => {
  if (!metadata) {
    return SyncStatus.LOCAL_ONLY;
  }

  if (externalUpdatedTimestamp === null || externalUpdatedTimestamp === undefined) {
    return SyncStatus.LOCAL_ONLY;
  }

  const roamUpdated = metadata.roamUpdated || metadata.lastSync;

  // Both modified since last sync
  if (externalUpdatedTimestamp > metadata.lastSync && roamUpdated > metadata.lastSync) {
    return SyncStatus.CONFLICT;
  }

  // External is newer
  if (externalUpdatedTimestamp > roamUpdated) {
    return SyncStatus.PENDING;
  }

  // Roam is newer
  if (roamUpdated > externalUpdatedTimestamp) {
    return SyncStatus.PENDING;
  }

  return SyncStatus.SYNCED;
};

/**
 * Create a provider-specific sync metadata module.
 *
 * @param {object} options
 * @param {string} options.storageKey - Extension storage key (e.g., "gcal-sync-metadata")
 * @param {string} options.logPrefix - Log prefix (e.g., "[SyncMetadata]" or "[OutlookSync]")
 * @param {Function} options.getConnectedCalendars - Function to get connected calendars
 * @param {string} options.externalIdField - Field name for external ID (e.g., "gCalId" or "outlookId")
 * @param {string} options.calendarIdField - Field name for calendar ID (e.g., "gCalCalendarId" or "outlookCalendarId")
 * @returns {object} Metadata module with all CRUD operations
 */
export const createSyncMetadataModule = ({
  storageKey,
  logPrefix,
  getConnectedCalendars,
  externalIdField,
  calendarIdField,
}) => {
  // In-memory cache
  let syncMetadataCache = null;

  /**
   * Get all trigger tags from all connected calendars
   */
  const getAllTriggerTags = () => {
    const calendars = getConnectedCalendars();
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
   * Remove all calendar trigger tags from a Roam block
   */
  const removeTriggerTagsFromBlock = (roamUid) => {
    if (!isExistingNode(roamUid)) return;

    const allTriggerTags = getAllTriggerTags();
    if (allTriggerTags.length > 0) {
      removeTagsFromBlock(roamUid, allTriggerTags);
    }
  };

  /**
   * Load all sync metadata from storage
   */
  const loadMetadata = () => {
    if (syncMetadataCache !== null) {
      return syncMetadataCache;
    }

    try {
      const stored = extensionStorage.get(storageKey);
      if (typeof stored === "string") {
        syncMetadataCache = JSON.parse(stored);
      } else {
        syncMetadataCache = stored || {};
      }
      console.log(
        `${logPrefix} Loaded ${Object.keys(syncMetadataCache).length} entries from storage`
      );
      return syncMetadataCache;
    } catch (error) {
      console.error(`${logPrefix} Failed to load sync metadata:`, error);
      syncMetadataCache = {};
      return syncMetadataCache;
    }
  };

  /**
   * Persist all sync metadata to storage
   */
  const persistMetadata = () => {
    try {
      const serialized = JSON.stringify(syncMetadataCache);
      extensionStorage.set(storageKey, serialized);
    } catch (error) {
      console.error(`${logPrefix} Failed to persist sync metadata:`, error);
    }
  };

  /**
   * Get sync metadata for a specific Roam block
   */
  const getMetadata = (roamUid) => {
    const allMetadata = loadMetadata();
    return allMetadata[roamUid] || null;
  };

  /**
   * Get Roam UID by external event ID
   */
  const getRoamUidByExternalId = (externalId) => {
    const allMetadata = loadMetadata();
    for (const [roamUid, metadata] of Object.entries(allMetadata)) {
      if (metadata[externalIdField] === externalId) {
        return roamUid;
      }
    }
    return null;
  };

  /**
   * Save sync metadata for a Roam block
   */
  const saveMetadata = async (roamUid, metadata) => {
    loadMetadata();
    syncMetadataCache[roamUid] = metadata;
    persistMetadata();
    return roamUid;
  };

  /**
   * Update specific fields in sync metadata
   */
  const updateMetadata = async (roamUid, updates) => {
    const existing = getMetadata(roamUid);

    if (existing) {
      const updatedMetadata = { ...existing, ...updates };
      await saveMetadata(roamUid, updatedMetadata);
      return updatedMetadata;
    }

    return null;
  };

  /**
   * Delete sync metadata for a Roam block
   */
  const deleteMetadata = async (roamUid) => {
    loadMetadata();

    if (syncMetadataCache[roamUid]) {
      delete syncMetadataCache[roamUid];
      persistMetadata();
      return true;
    }

    return false;
  };

  /**
   * Check if a Roam block is synced
   */
  const isSynced = (roamUid) => {
    return getMetadata(roamUid) !== null;
  };

  /**
   * Get external event ID from a Roam block
   */
  const getExternalIdFromEvent = (roamUid) => {
    const metadata = getMetadata(roamUid);
    return metadata ? metadata[externalIdField] : null;
  };

  /**
   * Get all synced events for a specific calendar
   */
  const getSyncedEventsForCalendar = (calendarId) => {
    const allMetadata = loadMetadata();
    const result = {};

    for (const [roamUid, metadata] of Object.entries(allMetadata)) {
      if (metadata[calendarIdField] === calendarId) {
        result[roamUid] = metadata;
      }
    }

    return result;
  };

  /**
   * Clear all sync metadata and remove trigger tags.
   */
  const clearAllMetadata = () => {
    loadMetadata();

    for (const roamUid of Object.keys(syncMetadataCache)) {
      removeTriggerTagsFromBlock(roamUid);
    }

    syncMetadataCache = {};
    persistMetadata();
    console.log(`${logPrefix} Cleared all sync metadata and removed trigger tags`);
  };

  /**
   * Get storage statistics
   */
  const getStorageStats = () => {
    const allMetadata = loadMetadata();
    const entries = Object.entries(allMetadata);

    let todoCount = 0;
    for (const [, metadata] of entries) {
      if (metadata.isTodo) {
        todoCount++;
      }
    }

    const estimatedBytes = entries.length * 200;

    return { eventCount: entries.length, todoCount, estimatedBytes };
  };

  /**
   * Cleanup old sync metadata for past events.
   * Removes metadata for events that ended more than N days ago,
   * unless the event still has TODO status.
   */
  const cleanupOldMetadata = (daysThreshold = 90) => {
    loadMetadata();

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const thresholdDate = new Date(today.getTime() - daysThreshold * 24 * 60 * 60 * 1000);

    let removedCount = 0;
    let keptTodoCount = 0;
    const toRemove = [];

    for (const [roamUid, metadata] of Object.entries(syncMetadataCache)) {
      if (!metadata.eventEndDate) continue;

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
      removeTriggerTagsFromBlock(roamUid);
      delete syncMetadataCache[roamUid];
      removedCount++;
    }

    if (removedCount > 0) {
      persistMetadata();
      console.log(`${logPrefix} Cleaned up ${removedCount} old entries, kept ${keptTodoCount} TODOs`);
    }

    return { removedCount, keptTodoCount };
  };

  /**
   * Cleanup ALL past events (manual cleanup).
   */
  const cleanupAllPastMetadata = () => {
    loadMetadata();

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    let removedCount = 0;
    const toRemove = [];

    for (const [roamUid, metadata] of Object.entries(syncMetadataCache)) {
      if (!metadata.eventEndDate) continue;

      const endDate = new Date(metadata.eventEndDate);

      if (endDate < today) {
        toRemove.push(roamUid);
      }
    }

    for (const roamUid of toRemove) {
      removeTriggerTagsFromBlock(roamUid);
      delete syncMetadataCache[roamUid];
      removedCount++;
    }

    if (removedCount > 0) {
      persistMetadata();
      console.log(`${logPrefix} Removed ${removedCount} past event entries`);
    }

    return { removedCount };
  };

  return {
    loadMetadata,
    getMetadata,
    getRoamUidByExternalId,
    saveMetadata,
    updateMetadata,
    deleteMetadata,
    isSynced,
    getExternalIdFromEvent,
    getSyncedEventsForCalendar,
    clearAllMetadata,
    getStorageStats,
    cleanupOldMetadata,
    cleanupAllPastMetadata,
  };
};
