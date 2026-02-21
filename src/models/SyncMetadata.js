/**
 * SyncMetadata - Google Calendar sync metadata storage
 *
 * Maps Roam block UIDs to their corresponding GCal event IDs.
 * Delegates core CRUD logic to the shared CalendarSyncMetadata module.
 */

import { getConnectedCalendars } from "../services/googleCalendarService";
import {
  createSyncMetadataModule,
  SyncStatus as SharedSyncStatus,
  determineSyncStatus as sharedDetermineSyncStatus,
} from "./CalendarSyncMetadata";

// Create provider-specific module using shared base
const gcalMetadata = createSyncMetadataModule({
  storageKey: "gcal-sync-metadata",
  logPrefix: "[SyncMetadata]",
  getConnectedCalendars,
  externalIdField: "gCalId",
  calendarIdField: "gCalCalendarId",
});

/**
 * Sync metadata structure for a single event
 */
export const createSyncMetadata = ({
  gCalId,
  gCalCalendarId,
  etag = null,
  gCalUpdated = null,
  roamUpdated = null,
  lastSync = Date.now(),
  eventEndDate = null,
  isTodo = false,
  hadOriginalTimeRange = false,
}) => ({
  gCalId,
  gCalCalendarId,
  etag,
  gCalUpdated,
  roamUpdated,
  lastSync,
  eventEndDate,
  isTodo,
  hadOriginalTimeRange,
});

/**
 * Sync status types (includes GCal-specific GCAL_ONLY)
 */
export const SyncStatus = {
  ...SharedSyncStatus,
  GCAL_ONLY: "gcal-only",
};

/**
 * Determine sync status by comparing timestamps (GCal-compatible wrapper)
 */
export const determineSyncStatus = (metadata, gCalEvent) => {
  if (!metadata) return SyncStatus.LOCAL_ONLY;
  if (!gCalEvent) return SyncStatus.LOCAL_ONLY;
  const gCalUpdated = new Date(gCalEvent.updated).getTime();
  return sharedDetermineSyncStatus(metadata, gCalUpdated);
};

// Export CRUD operations via shared module
export const loadSyncMetadata = gcalMetadata.loadMetadata;
export const getSyncMetadata = gcalMetadata.getMetadata;
export const getRoamUidByGCalId = gcalMetadata.getRoamUidByExternalId;
export const saveSyncMetadata = gcalMetadata.saveMetadata;
export const updateSyncMetadata = gcalMetadata.updateMetadata;
export const deleteSyncMetadata = gcalMetadata.deleteMetadata;
export const isSynced = gcalMetadata.isSynced;
export const getGCalIdFromEvent = gcalMetadata.getExternalIdFromEvent;
export const getSyncedEventsForCalendar = gcalMetadata.getSyncedEventsForCalendar;
export const clearAllSyncMetadata = gcalMetadata.clearAllMetadata;
export const getStorageStats = gcalMetadata.getStorageStats;
export const cleanupOldMetadata = gcalMetadata.cleanupOldMetadata;
export const cleanupAllPastMetadata = gcalMetadata.cleanupAllPastMetadata;

export default {
  createSyncMetadata,
  loadSyncMetadata,
  getSyncMetadata,
  getRoamUidByGCalId,
  saveSyncMetadata,
  updateSyncMetadata,
  deleteSyncMetadata,
  isSynced,
  getGCalIdFromEvent,
  getSyncedEventsForCalendar,
  clearAllSyncMetadata,
  determineSyncStatus,
  SyncStatus,
  getStorageStats,
  cleanupOldMetadata,
  cleanupAllPastMetadata,
};
