/**
 * OutlookSyncMetadata - Outlook Calendar sync metadata storage
 *
 * Maps Roam block UIDs to their corresponding Outlook event IDs.
 * Delegates core CRUD logic to the shared CalendarSyncMetadata module.
 */

import {
  createSyncMetadataModule,
  SyncStatus as SharedSyncStatus,
  determineSyncStatus as sharedDetermineSyncStatus,
} from "./CalendarSyncMetadata";

// Lazy-load connector to avoid circular dependency
const getOutlookCalendars = () => {
  const { getOutlookConnectedCalendars } = require("../services/outlookCalendarService");
  return getOutlookConnectedCalendars();
};

// Create provider-specific module using shared base
const outlookMetadata = createSyncMetadataModule({
  storageKey: "outlook-sync-metadata",
  logPrefix: "[OutlookSync]",
  getConnectedCalendars: getOutlookCalendars,
  externalIdField: "outlookId",
  calendarIdField: "outlookCalendarId",
});

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
 * Sync status types (includes Outlook-specific OUTLOOK_ONLY)
 */
export const OutlookSyncStatus = {
  ...SharedSyncStatus,
  OUTLOOK_ONLY: "outlook-only",
};

/**
 * Determine sync status by comparing timestamps (Outlook-compatible wrapper)
 */
export const determineOutlookSyncStatus = (metadata, outlookEvent) => {
  if (!metadata) return OutlookSyncStatus.LOCAL_ONLY;
  if (!outlookEvent) return OutlookSyncStatus.LOCAL_ONLY;
  const outlookUpdated = new Date(outlookEvent.lastModifiedDateTime).getTime();
  return sharedDetermineSyncStatus(metadata, outlookUpdated);
};

// Export CRUD operations via shared module with Outlook-specific names
export const loadOutlookSyncMetadata = outlookMetadata.loadMetadata;
export const getOutlookSyncMetadata = outlookMetadata.getMetadata;
export const getRoamUidByOutlookId = outlookMetadata.getRoamUidByExternalId;
export const saveOutlookSyncMetadata = outlookMetadata.saveMetadata;
export const updateOutlookSyncMetadata = outlookMetadata.updateMetadata;
export const deleteOutlookSyncMetadata = outlookMetadata.deleteMetadata;
export const isOutlookSynced = outlookMetadata.isSynced;
export const getOutlookIdFromEvent = outlookMetadata.getExternalIdFromEvent;
export const getSyncedEventsForOutlookCalendar = outlookMetadata.getSyncedEventsForCalendar;
export const clearAllOutlookSyncMetadata = outlookMetadata.clearAllMetadata;
export const getOutlookStorageStats = outlookMetadata.getStorageStats;
export const cleanupOldOutlookMetadata = outlookMetadata.cleanupOldMetadata;
export const cleanupAllPastOutlookMetadata = outlookMetadata.cleanupAllPastMetadata;
