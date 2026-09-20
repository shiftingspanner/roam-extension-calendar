/**
 * OutlookEventInfo - Reusable component to display Outlook Calendar event information
 * Used for synced events and matched-but-unsynced events
 */

import { Icon, Tooltip } from "@blueprintjs/core";
import { parseHtmlToReact } from "../util/htmlParser";
import OutlookCalendarIconSvg from "../services/outlook-calendar.svg";

const OutlookEventInfo = ({
  calendarName,
  location,
  attendees,
  description,
  webLink,
  showClickableCalendar = false,
  onCalendarClick,
}) => {
  return (
    <div className="fc-outlook-event-info">
      {/* Calendar name */}
      <Tooltip content="View in Outlook Calendar" position="top">
        {calendarName && (
          <div
            className={
              showClickableCalendar
                ? "fc-outlook-calendar-source fc-outlook-calendar-source-clickable"
                : "fc-outlook-calendar-source"
            }
            onClick={showClickableCalendar ? onCalendarClick : undefined}
            style={showClickableCalendar ? { cursor: "pointer" } : undefined}
          >
            <OutlookCalendarIconSvg
              className="fc-outlook-icon-small"
              style={{ width: "16px", height: "16px" }}
            />
            <span>{calendarName}</span>
          </div>
        )}
      </Tooltip>

      {/* Location */}
      {location && (
        <div className="fc-outlook-location">
          <Icon icon="map-marker" size={12} />
          <span>{location}</span>
        </div>
      )}

      {/* Attendees */}
      {attendees && attendees.length > 0 && (
        <div className="fc-outlook-attendees">
          <Icon icon="people" size={12} />
          <span>
            {attendees
              .slice(0, 3)
              .map((a) => a.emailAddress?.name || a.emailAddress?.address)
              .join(", ")}
            {attendees.length > 3 && ` +${attendees.length - 3} more`}
          </span>
        </div>
      )}

      {/* Description */}
      {description && (
        <div className="fc-outlook-description">
          {typeof description === "string"
            ? parseHtmlToReact(description)
            : description}
        </div>
      )}

      {/* Web link */}
      {webLink && (
        <div className="fc-outlook-weblink">
          <a
            href={webLink}
            target="_blank"
            rel="noopener noreferrer"
            onClick={(e) => e.stopPropagation()}
          >
            <Icon icon="share" size={12} />
            <span>Open in Outlook</span>
          </a>
        </div>
      )}
    </div>
  );
};

export default OutlookEventInfo;
