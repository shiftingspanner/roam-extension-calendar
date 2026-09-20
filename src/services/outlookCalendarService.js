/**
 * Outlook Calendar Service
 * Handles MSAL authentication, token management, and Microsoft Graph API interactions
 */

import { extensionStorage } from "..";
import { PublicClientApplication } from "@azure/msal-browser";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

// Default Client ID - users can override in settings
const DEFAULT_CLIENT_ID = "";

const GRAPH_SCOPES = ["Calendars.ReadWrite", "User.Read"];

// Service state
let msalInstance = null;
let msalAccount = null;
let authStateListeners = [];
let tokenRefreshInterval = null;
let onlineHandlers = [];

/**
 * Storage keys for Outlook data
 */
export const OUTLOOK_STORAGE_KEYS = {
  CLIENT_ID: "outlook-client-id",
  ACCESS_TOKEN: "outlook-access-token",
  TOKEN_EXPIRY: "outlook-token-expiry",
  CONNECTED_CALENDARS: "outlook-connected-calendars",
  USE_ORIGINAL_COLORS: "outlook-use-original-colors",
  CHECKBOX_FORMAT: "outlook-checkbox-format",
};

/**
 * Default calendar configuration
 */
export const DEFAULT_OUTLOOK_CALENDAR_CONFIG = {
  id: "",
  name: "",
  displayName: "",
  triggerTags: [],
  showAsSeparateTag: false,
  isDefault: true,
  syncEnabled: true,
  syncDirection: "both",
  lastSyncTime: 0,
  color: null,
};

/**
 * Get the configured client ID (user override or default)
 */
const getClientId = () => {
  const userClientId = extensionStorage.get(OUTLOOK_STORAGE_KEYS.CLIENT_ID);
  return userClientId || DEFAULT_CLIENT_ID;
};

/**
 * Set the client ID
 */
export const setOutlookClientId = (clientId) => {
  extensionStorage.set(OUTLOOK_STORAGE_KEYS.CLIENT_ID, clientId);
};

/**
 * Get the client ID
 */
export const getOutlookClientId = () => {
  return getClientId();
};

/**
 * Initialize or re-initialize the MSAL instance
 */
const initMsalInstance = async (clientId) => {
  if (!clientId) {
    console.warn("[Outlook Auth] No Client ID configured");
    return null;
  }

  const msalConfig = {
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };

  msalInstance = new PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Check for existing accounts
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalAccount = accounts[0];
  }

  return msalInstance;
};

/**
 * Notify all listeners of auth state change
 */
const notifyAuthStateChange = (isAuthenticated) => {
  authStateListeners.forEach((listener) => listener(isAuthenticated));
};

/**
 * Initialize Outlook Calendar service
 * Should be called on extension load
 */
export const initOutlookCalendarService = async () => {
  const clientId = getClientId();
  if (!clientId) {
    console.log("[Outlook Auth] No Client ID configured - skipping initialization");
    return false;
  }

  const savedToken = extensionStorage.get(OUTLOOK_STORAGE_KEYS.ACCESS_TOKEN);
  const tokenExpiry = extensionStorage.get(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY);

  // If offline and have stored credentials, enable cache-only mode
  if (!navigator.onLine && savedToken) {
    console.log("[Outlook Auth] Offline with stored credentials - enabling cache-only mode");
    notifyAuthStateChange(true);

    const onlineHandler = async () => {
      console.log("[Outlook Auth] Back online - completing initialization");
      window.removeEventListener("online", onlineHandler);
      const index = onlineHandlers.indexOf(onlineHandler);
      if (index > -1) onlineHandlers.splice(index, 1);

      try {
        await initMsalInstance(clientId);
        await silentRefreshOutlook();
        startOutlookTokenRefreshMonitoring();
      } catch (error) {
        console.error("[Outlook Auth] Failed to initialize after coming online:", error);
      }
    };
    window.addEventListener("online", onlineHandler);
    onlineHandlers.push(onlineHandler);

    return true;
  }

  try {
    await initMsalInstance(clientId);

    if (savedToken && tokenExpiry && Date.now() < tokenExpiry) {
      notifyAuthStateChange(true);
      startOutlookTokenRefreshMonitoring();
      return true;
    } else if (msalAccount) {
      // Have an MSAL account, try silent token acquisition
      const refreshed = await silentRefreshOutlook();
      if (refreshed) {
        startOutlookTokenRefreshMonitoring();
      }
      return refreshed;
    }

    return false;
  } catch (error) {
    if (savedToken) {
      console.warn("[Outlook Auth] Init failed but have stored credentials - enabling cache-only mode");
      notifyAuthStateChange(true);
      return true;
    }
    console.error("[Outlook Auth] Failed to initialize:", error);
    return false;
  }
};

/**
 * Try to silently refresh the token using MSAL
 */
export const silentRefreshOutlook = async () => {
  if (!msalInstance || !msalAccount) {
    return false;
  }

  try {
    const silentRequest = {
      scopes: GRAPH_SCOPES,
      account: msalAccount,
    };

    const response = await msalInstance.acquireTokenSilent(silentRequest);

    if (response && response.accessToken) {
      const expiryTime = response.expiresOn
        ? response.expiresOn.getTime()
        : Date.now() + 3600 * 1000;
      extensionStorage.set(OUTLOOK_STORAGE_KEYS.ACCESS_TOKEN, response.accessToken);
      extensionStorage.set(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY, expiryTime);
      notifyAuthStateChange(true);
      console.log("[Outlook Auth] Token refreshed silently");
      return true;
    }

    return false;
  } catch (error) {
    console.warn("[Outlook Auth] Silent refresh failed:", error.message);
    return false;
  }
};

/**
 * Start monitoring token expiry and proactively refresh
 */
const startOutlookTokenRefreshMonitoring = () => {
  if (tokenRefreshInterval) {
    clearInterval(tokenRefreshInterval);
  }

  tokenRefreshInterval = setInterval(async () => {
    const tokenExpiry = extensionStorage.get(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY);
    if (!tokenExpiry) return;

    // Refresh if token expires within 10 minutes
    const tenMinutes = 10 * 60 * 1000;
    if (Date.now() > tokenExpiry - tenMinutes) {
      console.log("[Outlook Auth] Proactive token refresh (expires soon)");
      try {
        await silentRefreshOutlook();
      } catch (error) {
        console.error("[Outlook Auth] Proactive refresh failed:", error);
      }
    }
  }, 5 * 60 * 1000); // Check every 5 minutes

  console.log("[Outlook Auth] Token refresh monitoring started");
};

/**
 * Stop monitoring token expiry
 */
export const stopOutlookTokenRefreshMonitoring = () => {
  if (tokenRefreshInterval) {
    clearInterval(tokenRefreshInterval);
    tokenRefreshInterval = null;
    console.log("[Outlook Auth] Token refresh monitoring stopped");
  }
};

/**
 * Cleanup all event listeners (call on extension unload)
 */
export const cleanupOutlookEventListeners = () => {
  onlineHandlers.forEach((handler) => {
    window.removeEventListener("online", handler);
  });
  onlineHandlers = [];
  console.log("[Outlook Auth] Cleaned up all event listeners");
};

/**
 * Check if user is currently authenticated with Outlook
 */
export const isOutlookAuthenticated = () => {
  const savedToken = extensionStorage.get(OUTLOOK_STORAGE_KEYS.ACCESS_TOKEN);
  const tokenExpiry = extensionStorage.get(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY);

  if (savedToken && Date.now() < (tokenExpiry || 0)) {
    return true;
  }

  // Offline with saved token
  if (!navigator.onLine && savedToken) {
    return true;
  }

  // Check MSAL account
  if (msalInstance) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      return true;
    }
  }

  return false;
};

/**
 * Get Outlook connection status
 */
export const getOutlookConnectionStatus = async () => {
  if (!isOutlookAuthenticated()) {
    return {
      isConnected: false,
      isOffline: false,
      reason: "not_authenticated",
      message: "Outlook Calendar not connected",
    };
  }

  if (!navigator.onLine) {
    return {
      isConnected: false,
      isOffline: true,
      reason: "offline",
      message: "You are offline",
    };
  }

  try {
    await getOutlookAccessToken();
    return {
      isConnected: true,
      isOffline: false,
      reason: null,
      message: "Connected",
    };
  } catch (error) {
    return {
      isConnected: false,
      isOffline: true,
      reason: "api_error",
      message: "Outlook Calendar connection failed",
    };
  }
};

/**
 * Request user authentication via MSAL popup
 * IMPORTANT: Must be called from a user click handler
 */
export const authenticateOutlook = async () => {
  const clientId = getClientId();
  if (!clientId) {
    throw new Error("No Azure AD Client ID configured. Please set one in the Outlook Calendar settings.");
  }

  if (!msalInstance) {
    await initMsalInstance(clientId);
  }

  if (!msalInstance) {
    throw new Error("Failed to initialize MSAL. Please check your Client ID.");
  }

  try {
    console.log("[Outlook Auth] Starting popup auth...");
    const loginResponse = await msalInstance.loginPopup({
      scopes: GRAPH_SCOPES,
      prompt: "select_account",
    });

    if (loginResponse && loginResponse.account) {
      msalAccount = loginResponse.account;

      // Now acquire a token
      const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        account: msalAccount,
      });

      if (tokenResponse && tokenResponse.accessToken) {
        const expiryTime = tokenResponse.expiresOn
          ? tokenResponse.expiresOn.getTime()
          : Date.now() + 3600 * 1000;
        extensionStorage.set(OUTLOOK_STORAGE_KEYS.ACCESS_TOKEN, tokenResponse.accessToken);
        extensionStorage.set(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY, expiryTime);
        notifyAuthStateChange(true);
        startOutlookTokenRefreshMonitoring();

        console.log("[Outlook Auth] Authentication complete!");
        return tokenResponse;
      }
    }

    throw new Error("Authentication failed - no token received");
  } catch (error) {
    console.error("[Outlook Auth] Authentication error:", error);
    throw error;
  }
};

/**
 * Sign out from Outlook
 */
export const signOutOutlook = async () => {
  stopOutlookTokenRefreshMonitoring();

  if (msalInstance && msalAccount) {
    try {
      await msalInstance.logoutPopup({
        account: msalAccount,
      });
    } catch (error) {
      console.error("[Outlook Auth] Error during logout:", error);
    }
  }

  msalAccount = null;

  extensionStorage.set(OUTLOOK_STORAGE_KEYS.ACCESS_TOKEN, null);
  extensionStorage.set(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY, null);

  notifyAuthStateChange(false);
};

/**
 * Add listener for authentication state changes
 */
export const onOutlookAuthStateChange = (callback) => {
  authStateListeners.push(callback);
  return () => {
    authStateListeners = authStateListeners.filter((cb) => cb !== callback);
  };
};

/**
 * Get access token, refreshing if needed
 */
export const getOutlookAccessToken = async () => {
  const tokenExpiry = extensionStorage.get(OUTLOOK_STORAGE_KEYS.TOKEN_EXPIRY);

  // Check if token is about to expire (within 5 minutes)
  if (tokenExpiry && Date.now() > tokenExpiry - 5 * 60 * 1000) {
    const refreshed = await silentRefreshOutlook();
    if (!refreshed) {
      throw new Error("Token expired and refresh failed");
    }
  }

  const token = extensionStorage.get(OUTLOOK_STORAGE_KEYS.ACCESS_TOKEN);
  if (!token) {
    throw new Error("No Outlook access token available");
  }

  return token;
};

// ============================================
// Microsoft Graph API Helper
// ============================================

/**
 * Make an authenticated request to Microsoft Graph API
 */
const graphFetch = async (url, options = {}) => {
  const token = await getOutlookAccessToken();
  const fullUrl = url.startsWith("http") ? url : `${GRAPH_BASE_URL}${url}`;

  const response = await fetch(fullUrl, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const error = new Error(
      errorData.error?.message || `Graph API error: ${response.status}`
    );
    error.status = response.status;
    error.code = errorData.error?.code;
    throw error;
  }

  // DELETE returns no content
  if (response.status === 204) {
    return null;
  }

  return response.json();
};

// ============================================
// Calendar API Methods
// ============================================

/**
 * List all calendars accessible by the user
 */
export const listOutlookCalendars = async () => {
  try {
    const data = await graphFetch("/me/calendars?$top=100");
    return data.value || [];
  } catch (error) {
    console.error("[Outlook] Error listing calendars:", error);
    throw error;
  }
};

/**
 * Get events from a specific calendar (uses calendarView to expand recurrences)
 * @param {string} calendarId - Calendar ID
 * @param {Date|string} timeMin - Start of date range
 * @param {Date|string} timeMax - End of date range
 * @param {object} options - Additional options
 */
export const getOutlookEvents = async (calendarId, timeMin, timeMax, options = {}) => {
  try {
    const startDateTime =
      timeMin instanceof Date ? timeMin.toISOString() : timeMin;
    const endDateTime =
      timeMax instanceof Date ? timeMax.toISOString() : timeMax;

    const selectFields = [
      "id",
      "subject",
      "body",
      "start",
      "end",
      "location",
      "attendees",
      "organizer",
      "isAllDay",
      "categories",
      "recurrence",
      "lastModifiedDateTime",
      "changeKey",
      "webLink",
      "isCancelled",
      "showAs",
      "importance",
    ].join(",");

    let url;
    if (options.updatedMin) {
      // Incremental sync - fetch events updated since a given time
      const updatedMin =
        options.updatedMin instanceof Date
          ? options.updatedMin.toISOString()
          : options.updatedMin;
      url =
        `/me/calendars/${calendarId}/events` +
        `?$filter=lastModifiedDateTime ge ${updatedMin}` +
        `&$select=${selectFields}` +
        `&$top=${options.maxResults || 250}` +
        `&$orderby=start/dateTime`;
    } else {
      // Calendar view - expands recurring events into instances
      url =
        `/me/calendars/${calendarId}/calendarView` +
        `?startDateTime=${startDateTime}` +
        `&endDateTime=${endDateTime}` +
        `&$select=${selectFields}` +
        `&$top=${options.maxResults || 250}` +
        `&$orderby=start/dateTime`;
    }

    const data = await graphFetch(url);
    const events = data.value || [];

    // Filter out cancelled events
    return events.filter((e) => !e.isCancelled);
  } catch (error) {
    console.error("[Outlook] Error fetching events:", error);
    throw error;
  }
};

/**
 * Create a new event in Outlook Calendar
 * @param {string} calendarId - Calendar ID
 * @param {object} event - Event data (Microsoft Graph format)
 */
export const createOutlookEvent = async (calendarId, event) => {
  try {
    const result = await graphFetch(`/me/calendars/${calendarId}/events`, {
      method: "POST",
      body: JSON.stringify(event),
    });
    return result;
  } catch (error) {
    console.error("[Outlook] Error creating event:", error);
    throw error;
  }
};

/**
 * Update an existing event in Outlook Calendar
 * @param {string} eventId - Event ID
 * @param {object} event - Updated event data
 */
export const updateOutlookEvent = async (eventId, event) => {
  try {
    const result = await graphFetch(`/me/events/${eventId}`, {
      method: "PATCH",
      body: JSON.stringify(event),
    });
    return result;
  } catch (error) {
    console.error("[Outlook] Error updating event:", error);
    throw error;
  }
};

/**
 * Delete an event from Outlook Calendar
 * @param {string} eventId - Event ID
 */
export const deleteOutlookEvent = async (eventId) => {
  try {
    await graphFetch(`/me/events/${eventId}`, {
      method: "DELETE",
    });
    return true;
  } catch (error) {
    console.error("[Outlook] Error deleting event:", error);
    throw error;
  }
};

/**
 * Get a single event by ID
 * @param {string} eventId - Event ID
 */
export const getOutlookEventById = async (eventId) => {
  try {
    return await graphFetch(`/me/events/${eventId}`);
  } catch (error) {
    console.error("[Outlook] Error getting event:", error);
    throw error;
  }
};

// ============================================
// Connected Calendars Management
// ============================================

/**
 * Get connected Outlook calendars configuration
 */
export const getOutlookConnectedCalendars = () => {
  const calendars = extensionStorage.get(OUTLOOK_STORAGE_KEYS.CONNECTED_CALENDARS);
  if (!calendars) return [];

  const parsedCalendars = JSON.parse(calendars);

  return parsedCalendars.map((cal) => ({
    ...DEFAULT_OUTLOOK_CALENDAR_CONFIG,
    ...cal,
    showAsSeparateTag: cal.showAsSeparateTag ?? false,
  }));
};

/**
 * Save connected Outlook calendars configuration
 */
export const saveOutlookConnectedCalendars = (calendars) => {
  extensionStorage.set(
    OUTLOOK_STORAGE_KEYS.CONNECTED_CALENDARS,
    JSON.stringify(calendars)
  );
};

/**
 * Add a new connected Outlook calendar
 */
export const addOutlookConnectedCalendar = (calendarConfig) => {
  const calendars = getOutlookConnectedCalendars();
  if (calendarConfig.isDefault) {
    calendars.forEach((cal) => (cal.isDefault = false));
  }
  calendars.push({ ...DEFAULT_OUTLOOK_CALENDAR_CONFIG, ...calendarConfig });
  saveOutlookConnectedCalendars(calendars);
  return calendars;
};

/**
 * Update a connected Outlook calendar configuration
 */
export const updateOutlookConnectedCalendar = (calendarId, updates) => {
  const calendars = getOutlookConnectedCalendars();
  const index = calendars.findIndex((cal) => cal.id === calendarId);
  if (index !== -1) {
    if (updates.isDefault) {
      calendars.forEach((cal) => (cal.isDefault = false));
    }
    calendars[index] = { ...calendars[index], ...updates };
    saveOutlookConnectedCalendars(calendars);
  }
  return calendars;
};

/**
 * Remove a connected Outlook calendar
 */
export const removeOutlookConnectedCalendar = (calendarId) => {
  const calendars = getOutlookConnectedCalendars().filter(
    (cal) => cal.id !== calendarId
  );
  saveOutlookConnectedCalendars(calendars);
  return calendars;
};

/**
 * Find Outlook calendar by trigger tag or displayName
 */
export const findOutlookCalendarByTag = (tagName) => {
  const calendars = getOutlookConnectedCalendars();
  const lowerTagName = tagName.toLowerCase();
  return calendars.find((cal) => {
    if (cal.displayName && cal.displayName.toLowerCase() === lowerTagName) {
      return true;
    }
    return cal.triggerTags.some((tag) => tag.toLowerCase() === lowerTagName);
  });
};

/**
 * Get default Outlook calendar
 */
export const getOutlookDefaultCalendar = () => {
  const calendars = getOutlookConnectedCalendars();
  return calendars.find((cal) => cal.isDefault) || calendars[0] || null;
};

// ============================================
// Sync Settings
// ============================================

/**
 * Get whether to use original Outlook calendar colors
 */
export const getOutlookUseOriginalColors = () => {
  return extensionStorage.get(OUTLOOK_STORAGE_KEYS.USE_ORIGINAL_COLORS) ?? false;
};

/**
 * Set whether to use original Outlook calendar colors
 */
export const setOutlookUseOriginalColors = (enabled) => {
  extensionStorage.set(OUTLOOK_STORAGE_KEYS.USE_ORIGINAL_COLORS, enabled);
};

/**
 * Get Outlook checkbox format preference
 */
export const getOutlookCheckboxFormat = () => {
  return extensionStorage.get(OUTLOOK_STORAGE_KEYS.CHECKBOX_FORMAT) ?? "roam";
};

/**
 * Set Outlook checkbox format preference
 */
export const setOutlookCheckboxFormat = (format) => {
  extensionStorage.set(OUTLOOK_STORAGE_KEYS.CHECKBOX_FORMAT, format);
};
