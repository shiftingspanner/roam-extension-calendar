/**
 * OutlookConfigDialog - Configuration dialog for Outlook Calendar integration
 * Mirrors GCalConfigDialog pattern but uses MSAL auth and Microsoft Graph API
 */

import {
  Button,
  Callout,
  Card,
  Classes,
  Dialog,
  FormGroup,
  HTMLSelect,
  Icon,
  InputGroup,
  Spinner,
  Switch,
  Toaster,
  Position,
  Intent,
} from "@blueprintjs/core";
import { useState, useEffect } from "react";
import {
  authenticateOutlook,
  signOutOutlook,
  isOutlookAuthenticated,
  listOutlookCalendars,
  getOutlookConnectedCalendars,
  saveOutlookConnectedCalendars,
  updateOutlookConnectedCalendar,
  getOutlookUseOriginalColors,
  setOutlookUseOriginalColors,
  getOutlookCheckboxFormat,
  setOutlookCheckboxFormat,
  getOutlookClientId,
  setOutlookClientId,
  onOutlookAuthStateChange,
  DEFAULT_OUTLOOK_CALENDAR_CONFIG,
} from "../services/outlookCalendarService";
import { initializeOutlookTags, mapOfTags } from "../index";
import { getTagFromName } from "../models/EventTag";
import { updateStoredTags } from "../util/data";
import {
  getOutlookStorageStats,
  cleanupAllPastOutlookMetadata,
  clearAllOutlookSyncMetadata,
} from "../models/OutlookSyncMetadata";
import { invalidateAllEventsCache } from "../services/eventCacheService";

const OutlookConfigDialog = ({ isOpen, onClose }) => {
  const [isConnected, setIsConnected] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [configChanged, setConfigChanged] = useState(false);
  const [clientId, setClientIdState] = useState("");

  const showToast = (message, intent = Intent.PRIMARY) => {
    const toaster = Toaster.create({ position: Position.TOP });
    toaster.show({ message, intent, timeout: 3000 });
  };

  // Calendars
  const [availableCalendars, setAvailableCalendars] = useState([]);
  const [calendarConfigs, setCalendarConfigs] = useState([]);

  // Settings
  const [useOriginalColors, setUseOriginalColorsState] = useState(false);
  const [checkboxFormat, setCheckboxFormatState] = useState("roam");

  // Sync stats
  const [syncStats, setSyncStats] = useState({ eventCount: 0, todoCount: 0 });

  // Confirmation dialogs
  const [confirmReinitSync, setConfirmReinitSync] = useState(false);

  // Handler for connect button
  const handleConnect = async () => {
    if (!clientId.trim()) {
      setError("Please enter an Azure AD Client ID first.");
      return;
    }

    // Save the client ID
    setOutlookClientId(clientId.trim());

    setIsLoading(true);
    setError("");

    try {
      await authenticateOutlook();
    } catch (err) {
      console.error(err);
      setIsLoading(false);
      const errorMessage = err?.message || "";
      if (
        errorMessage.includes("popup") ||
        errorMessage.includes("blocked")
      ) {
        setError(
          "Connection failed. Please check if your browser is blocking popups."
        );
      } else if (
        errorMessage.includes("cancelled") ||
        errorMessage.includes("user_cancelled")
      ) {
        setError("Authentication was cancelled.");
      } else {
        setError(`Failed to connect to Outlook: ${errorMessage}`);
      }
    }
  };

  // Load initial state
  useEffect(() => {
    if (isOpen) {
      loadInitialState();
    }
  }, [isOpen]);

  // Listen for auth state changes
  useEffect(() => {
    const unsubscribe = onOutlookAuthStateChange(async (authenticated) => {
      setIsConnected(authenticated);
      if (authenticated) {
        setConfigChanged(true);
        if (isOpen) {
          try {
            await new Promise((resolve) => setTimeout(resolve, 500));
            await fetchAvailableCalendars();
          } finally {
            setIsLoading(false);
          }
        }
      }
    });
    return unsubscribe;
  }, [isOpen]);

  const loadInitialState = async () => {
    setIsLoading(true);
    setError("");

    try {
      const authenticated = isOutlookAuthenticated();
      setIsConnected(authenticated);

      setClientIdState(getOutlookClientId() || "");

      const calConfigs = getOutlookConnectedCalendars();
      setCalendarConfigs(calConfigs);

      setUseOriginalColorsState(getOutlookUseOriginalColors());
      setCheckboxFormatState(getOutlookCheckboxFormat());

      setSyncStats(getOutlookStorageStats());

      if (authenticated) {
        await fetchAvailableCalendars();
      }
    } catch (err) {
      setError("Failed to load configuration");
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const fetchAvailableCalendars = async (retryCount = 0) => {
    try {
      const calendars = await listOutlookCalendars();

      setAvailableCalendars(calendars);
      setError("");

      // Initialize configs for any new calendars
      const existingConfigs = getOutlookConnectedCalendars();
      const existingIds = new Set(existingConfigs.map((c) => c.id));

      const newConfigs = [...existingConfigs];
      for (const cal of calendars) {
        if (!existingIds.has(cal.id)) {
          newConfigs.push({
            ...DEFAULT_OUTLOOK_CALENDAR_CONFIG,
            id: cal.id,
            name: cal.name,
            displayName: cal.name,
            triggerTags: [],
            syncEnabled: false,
            isDefault: newConfigs.length === 0,
            color: cal.color || null,
          });
        } else {
          const existingIndex = newConfigs.findIndex((c) => c.id === cal.id);
          if (existingIndex !== -1) {
            newConfigs[existingIndex].name = cal.name;
            if (
              !newConfigs[existingIndex].displayName ||
              newConfigs[existingIndex].displayName ===
                newConfigs[existingIndex].name
            ) {
              newConfigs[existingIndex].displayName = cal.name;
            }
            newConfigs[existingIndex].color = cal.color || null;
          }
        }
      }

      const availableIds = new Set(calendars.map((c) => c.id));
      const filteredConfigs = newConfigs.filter((c) =>
        availableIds.has(c.id)
      );

      saveOutlookConnectedCalendars(filteredConfigs);
      setCalendarConfigs(filteredConfigs);
    } catch (err) {
      console.error("Failed to fetch Outlook calendars:", err);

      if (retryCount < 2) {
        const delay = Math.pow(2, retryCount) * 1000;
        console.log(
          `Retrying in ${delay}ms (attempt ${retryCount + 1}/2)...`
        );
        await new Promise((resolve) => setTimeout(resolve, delay));
        return fetchAvailableCalendars(retryCount + 1);
      }

      setError("Failed to load calendars. Please try disconnecting and reconnecting.");
    }
  };

  const handleDisconnect = async () => {
    try {
      await signOutOutlook();
      setIsConnected(false);
      setAvailableCalendars([]);
      setCalendarConfigs([]);
      setConfigChanged(true);
      showToast("Disconnected from Outlook Calendar", Intent.WARNING);
    } catch (err) {
      setError("Failed to disconnect");
    }
  };

  const handleCalendarToggle = (calendarId, enabled) => {
    updateOutlookConnectedCalendar(calendarId, { syncEnabled: enabled });
    setCalendarConfigs(getOutlookConnectedCalendars());
    setConfigChanged(true);
  };

  const handleDisplayNameChange = (calendarId, displayName) => {
    updateOutlookConnectedCalendar(calendarId, { displayName });
    setCalendarConfigs(getOutlookConnectedCalendars());
    setConfigChanged(true);
  };

  const handleTriggerTagsChange = (calendarId, triggerTagsStr) => {
    const triggerTags = triggerTagsStr
      .split(",")
      .map((t) => t.trim())
      .filter((t) => t);
    updateOutlookConnectedCalendar(calendarId, { triggerTags });
    setCalendarConfigs(getOutlookConnectedCalendars());
    setConfigChanged(true);
  };

  const handleSyncDirectionChange = (calendarId, syncDirection) => {
    updateOutlookConnectedCalendar(calendarId, { syncDirection });
    setCalendarConfigs(getOutlookConnectedCalendars());
    setConfigChanged(true);
  };

  const handleDefaultChange = (calendarId) => {
    calendarConfigs.forEach((cal) => {
      updateOutlookConnectedCalendar(cal.id, {
        isDefault: cal.id === calendarId,
      });
    });
    setCalendarConfigs(getOutlookConnectedCalendars());
    setConfigChanged(true);
  };

  const handleSeparateTagToggle = (calendarId, showAsSeparateTag) => {
    updateOutlookConnectedCalendar(calendarId, { showAsSeparateTag });
    setCalendarConfigs(getOutlookConnectedCalendars());
    setConfigChanged(true);
  };

  const handleUseOriginalColorsChange = (enabled) => {
    setOutlookUseOriginalColors(enabled);
    setUseOriginalColorsState(enabled);
    setConfigChanged(true);
  };

  const handleCheckboxFormatChange = (format) => {
    setOutlookCheckboxFormat(format);
    setCheckboxFormatState(format);
    setConfigChanged(true);
  };

  const handleClearSyncData = () => {
    clearAllOutlookSyncMetadata();
    setSyncStats(getOutlookStorageStats());
    setConfigChanged(true);
    showToast("Outlook sync data cleared", Intent.WARNING);
    setConfirmReinitSync(false);
  };

  const handleCleanupPastEvents = () => {
    const result = cleanupAllPastOutlookMetadata();
    setSyncStats(getOutlookStorageStats());
    showToast(
      `Cleaned up ${result.removedCount} past Outlook events`,
      Intent.SUCCESS
    );
  };

  const handleClose = () => {
    if (configChanged) {
      // Reinitialize tags to reflect new calendar configs
      if (typeof initializeOutlookTags === "function") {
        initializeOutlookTags();
      }
      invalidateAllEventsCache();
    }
    onClose({ shouldRemountCalendar: configChanged });
  };

  return (
    <Dialog
      isOpen={isOpen}
      onClose={handleClose}
      title="Outlook Calendar Settings"
      className="fc-outlook-config-dialog"
      style={{ width: "750px" }}
      canOutsideClickClose={false}
    >
      <div className={Classes.DIALOG_BODY}>
        {error && (
          <Callout intent={Intent.DANGER} style={{ marginBottom: "15px" }}>
            {error}
          </Callout>
        )}

        {/* Client ID Section */}
        <Card style={{ marginBottom: "15px" }}>
          <h4>Azure AD Client ID</h4>
          <p style={{ fontSize: "12px", color: "#888", marginBottom: "10px" }}>
            Register an app in the{" "}
            <a
              href="https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"
              target="_blank"
              rel="noopener noreferrer"
            >
              Azure Portal
            </a>{" "}
            to get a Client ID. Set redirect URI as SPA type with your origin URL.
            Required permissions: Calendars.ReadWrite, User.Read.
          </p>
          <FormGroup>
            <InputGroup
              placeholder="Enter Azure AD Client ID..."
              value={clientId}
              onChange={(e) => setClientIdState(e.target.value)}
              disabled={isConnected}
            />
          </FormGroup>
        </Card>

        {/* Connection Section */}
        <Card style={{ marginBottom: "15px" }}>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <div>
              <h4>Connection</h4>
              <p
                style={{
                  color: isConnected ? "#0F9960" : "#888",
                  fontSize: "12px",
                }}
              >
                {isConnected
                  ? "Connected to Outlook Calendar"
                  : "Not connected"}
              </p>
            </div>
            <div>
              {isLoading ? (
                <Spinner size={20} />
              ) : isConnected ? (
                <Button
                  intent={Intent.DANGER}
                  text="Disconnect"
                  onClick={handleDisconnect}
                  minimal
                />
              ) : (
                <Button
                  intent={Intent.PRIMARY}
                  text="Connect Outlook"
                  onClick={handleConnect}
                  disabled={!clientId.trim()}
                />
              )}
            </div>
          </div>
        </Card>

        {/* Calendars Section */}
        {isConnected && calendarConfigs.length > 0 && (
          <Card style={{ marginBottom: "15px" }}>
            <h4>Calendars</h4>
            <div
              style={{
                maxHeight: "300px",
                overflowY: "auto",
              }}
            >
              <table
                className="bp5-html-table bp5-html-table-condensed"
                style={{ width: "100%" }}
              >
                <thead>
                  <tr>
                    <th>Enabled</th>
                    <th>Calendar</th>
                    <th>Display Name</th>
                    <th>Trigger Tags</th>
                    <th>Sync</th>
                    <th>Default</th>
                    <th>Separate Tag</th>
                  </tr>
                </thead>
                <tbody>
                  {calendarConfigs.map((cal) => (
                    <tr key={cal.id}>
                      <td>
                        <Switch
                          checked={cal.syncEnabled}
                          onChange={(e) =>
                            handleCalendarToggle(
                              cal.id,
                              e.target.checked
                            )
                          }
                          style={{ marginBottom: 0 }}
                        />
                      </td>
                      <td>
                        <span
                          style={{
                            fontSize: "12px",
                            maxWidth: "120px",
                            display: "inline-block",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}
                          title={cal.name}
                        >
                          {cal.name}
                        </span>
                      </td>
                      <td>
                        <InputGroup
                          value={cal.displayName}
                          onChange={(e) =>
                            handleDisplayNameChange(
                              cal.id,
                              e.target.value
                            )
                          }
                          small
                          placeholder="Display name..."
                        />
                      </td>
                      <td>
                        <InputGroup
                          value={(cal.triggerTags || []).join(", ")}
                          onChange={(e) =>
                            handleTriggerTagsChange(
                              cal.id,
                              e.target.value
                            )
                          }
                          small
                          placeholder="tag1, tag2..."
                        />
                      </td>
                      <td>
                        <HTMLSelect
                          value={cal.syncDirection || "both"}
                          onChange={(e) =>
                            handleSyncDirectionChange(
                              cal.id,
                              e.target.value
                            )
                          }
                          minimal
                          small
                        >
                          <option value="both">Both</option>
                          <option value="import">Import</option>
                          <option value="export">Export</option>
                        </HTMLSelect>
                      </td>
                      <td>
                        <input
                          type="radio"
                          name="outlookDefaultCalendar"
                          checked={cal.isDefault}
                          onChange={() => handleDefaultChange(cal.id)}
                        />
                      </td>
                      <td>
                        <Switch
                          checked={cal.showAsSeparateTag || false}
                          onChange={(e) =>
                            handleSeparateTagToggle(
                              cal.id,
                              e.target.checked
                            )
                          }
                          style={{ marginBottom: 0 }}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        )}

        {/* Display Settings */}
        {isConnected && (
          <Card style={{ marginBottom: "15px" }}>
            <h4>Display Settings</h4>
            <Switch
              label="Use original Outlook calendar colors"
              checked={useOriginalColors}
              onChange={(e) =>
                handleUseOriginalColorsChange(e.target.checked)
              }
            />
            <FormGroup
              label="Checkbox format in Outlook"
              style={{ marginTop: "10px" }}
            >
              <HTMLSelect
                value={checkboxFormat}
                onChange={(e) =>
                  handleCheckboxFormatChange(e.target.value)
                }
              >
                <option value="roam">
                  {"[[TODO]] / [[DONE]]"}
                </option>
                <option value="bracket">{"[ ] / [x]"}</option>
              </HTMLSelect>
            </FormGroup>
          </Card>
        )}

        {/* Sync Data */}
        {isConnected && (
          <Card>
            <h4>Sync Data</h4>
            <p style={{ fontSize: "12px", color: "#888" }}>
              {syncStats.eventCount} synced events (
              {syncStats.todoCount} TODOs)
            </p>
            <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
              <Button
                small
                text="Cleanup past events"
                onClick={handleCleanupPastEvents}
              />
              <Button
                small
                intent={Intent.DANGER}
                text="Clear all sync data"
                onClick={() => setConfirmReinitSync(true)}
              />
            </div>

            {/* Confirmation dialog */}
            <Dialog
              isOpen={confirmReinitSync}
              onClose={() => setConfirmReinitSync(false)}
              title="Clear Outlook Sync Data?"
              style={{ width: "400px" }}
            >
              <div className={Classes.DIALOG_BODY}>
                <p>
                  This will remove all sync metadata. Synced events will
                  remain in both Roam and Outlook, but the link between
                  them will be lost.
                </p>
              </div>
              <div className={Classes.DIALOG_FOOTER}>
                <div className={Classes.DIALOG_FOOTER_ACTIONS}>
                  <Button
                    text="Cancel"
                    onClick={() => setConfirmReinitSync(false)}
                  />
                  <Button
                    intent={Intent.DANGER}
                    text="Clear"
                    onClick={handleClearSyncData}
                  />
                </div>
              </div>
            </Dialog>
          </Card>
        )}
      </div>

      <div className={Classes.DIALOG_FOOTER}>
        <div className={Classes.DIALOG_FOOTER_ACTIONS}>
          <Button text="Close" onClick={handleClose} />
        </div>
      </div>
    </Dialog>
  );
};

export default OutlookConfigDialog;
