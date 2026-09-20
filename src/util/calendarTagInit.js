/**
 * Shared Calendar Tag Initialization
 *
 * Provides a generic function to initialize EventTags for connected calendars,
 * eliminating duplication between initializeGCalTags and initializeOutlookTags.
 */

import { Colors } from "@blueprintjs/core";
import { EventTag, getTagFromName } from "../models/EventTag";

/**
 * Initialize EventTags for a calendar provider's connected calendars.
 *
 * Calendars with showAsSeparateTag=false are grouped under the main provider tag.
 * Calendars with showAsSeparateTag=true get their own EventTag.
 *
 * @param {object} options
 * @param {array|null} options.calendarsOverride - Override calendars (for async safety)
 * @param {Function} options.getConnectedCalendars - Function to get connected calendars
 * @param {string} options.mainTagName - Main tag name (e.g., "Google calendar" or "Outlook calendar")
 * @param {string} options.providerPrefix - Prefix for tag properties (e.g., "gCal" or "outlook")
 * @param {string} options.defaultColor - Default color for new tags
 * @param {array} options.mapOfTags - Reference to the mapOfTags array
 * @param {Function} options.getStoredTagInfos - Function to get stored tag display settings
 * @param {boolean} [options.applyOriginalColors=false] - Whether to apply original calendar colors
 * @param {Function} [options.getUseOriginalColors] - Function to check if original colors are enabled
 */
export const initializeCalendarProviderTags = ({
  calendarsOverride = null,
  getConnectedCalendars,
  mainTagName,
  providerPrefix,
  defaultColor,
  mapOfTags,
  getStoredTagInfos,
  applyOriginalColors = false,
  getUseOriginalColors = null,
}) => {
  const connectedCalendars = calendarsOverride || getConnectedCalendars();
  if (!connectedCalendars || !connectedCalendars.length) return;

  const mainTag = getTagFromName(mainTagName);
  if (!mainTag) {
    console.warn(`Main '${mainTagName}' tag not found`);
    return;
  }

  // Determine property names based on provider
  // GCal uses: isGCalTag, gCalCalendarId, gCalCalendarIds, disabledCalendarIds
  // Outlook uses: isOutlookTag, outlookCalendarId, outlookCalendarIds, disabledOutlookCalendarIds
  const isProviderTagKey = `is${providerPrefix.charAt(0).toUpperCase() + providerPrefix.slice(1)}Tag`;
  const calendarIdKey = `${providerPrefix}CalendarId`;
  const calendarIdsKey = `${providerPrefix}CalendarIds`;
  const disabledIdsKey = `disabled${providerPrefix.charAt(0).toUpperCase() + providerPrefix.slice(1)}CalendarIds`;

  // Remove separate tags that are no longer configured as separate
  const separateCalendarNames = connectedCalendars
    .filter((cal) => cal.showAsSeparateTag)
    .map((cal) => cal.displayName || cal.name);

  for (let i = mapOfTags.length - 1; i >= 0; i--) {
    const tag = mapOfTags[i];
    if (
      tag[isProviderTagKey] &&
      tag[calendarIdKey] &&
      !separateCalendarNames.includes(tag.name)
    ) {
      mapOfTags.splice(i, 1);
    }
  }

  // Initialize arrays for the main tag
  mainTag[calendarIdsKey] = [];
  mainTag[disabledIdsKey] = [];

  for (const calendarConfig of connectedCalendars) {
    if (calendarConfig.showAsSeparateTag) {
      const tagName = calendarConfig.displayName || calendarConfig.name;
      let existingTag = getTagFromName(tagName);

      if (!existingTag) {
        const pages = [tagName];
        if (calendarConfig.triggerTags && calendarConfig.triggerTags.length > 0) {
          pages.push(...calendarConfig.triggerTags);
        }

        const newTag = new EventTag({
          name: tagName,
          color: defaultColor,
          ...getStoredTagInfos(tagName),
          pages: pages,
          [isProviderTagKey]: true,
          [calendarIdKey]: calendarConfig.id,
          isToDisplay: true,
          isToDisplayInSb: true,
        });
        mapOfTags.push(newTag);
      } else {
        existingTag[calendarIdKey] = calendarConfig.id;
        existingTag[isProviderTagKey] = true;

        if (calendarConfig.triggerTags && calendarConfig.triggerTags.length > 0) {
          const currentPages = existingTag.pages || [existingTag.name];
          const newPages = [
            ...new Set([...currentPages, ...calendarConfig.triggerTags]),
          ];
          existingTag.updatePages(newPages);
        }
      }
    } else {
      mainTag[calendarIdsKey].push(calendarConfig.id);

      if (!calendarConfig.syncEnabled) {
        mainTag[disabledIdsKey].push(calendarConfig.id);
      }

      if (calendarConfig.triggerTags && calendarConfig.triggerTags.length > 0) {
        const currentPages = mainTag.pages || [mainTagName];
        const newPages = [
          ...new Set([...currentPages, ...calendarConfig.triggerTags]),
        ];
        mainTag.updatePages(newPages);
      }
    }
  }

  // Apply original calendar colors if enabled
  if (applyOriginalColors && getUseOriginalColors && getUseOriginalColors()) {
    let defaultCalendarColor = null;

    for (const calendarConfig of connectedCalendars) {
      if (!calendarConfig.syncEnabled || !calendarConfig.backgroundColor)
        continue;

      if (calendarConfig.showAsSeparateTag) {
        const tagName = calendarConfig.displayName || calendarConfig.name;
        const tag = getTagFromName(tagName);
        if (tag) {
          tag.setColor(calendarConfig.backgroundColor);
        }
      } else if (calendarConfig.isDefault) {
        defaultCalendarColor = calendarConfig.backgroundColor;
      }
    }

    if (defaultCalendarColor && mainTag) {
      mainTag.setColor(defaultCalendarColor);
    }
  }
};
