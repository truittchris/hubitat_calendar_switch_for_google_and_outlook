/*
 *  Hubitat Calendar Switch
 *
 *  Author: Chris Truitt
 *  Website: https://christruitt.com
 *  GitHub:  https://github.com/truittchris
 *  Namespace: truittchris
 *
 *  Summary
 *  - Handles OAuth for Google and Microsoft.
 *  - Creates child devices that evaluate calendar events and control their own switch state.
 *  - Polls providers and delivers a normalized event list to each child.
 *
 *  Architectural rule
 *  - This app must not implement include/exclude keyword parsing or matching.
 *  - All match rules live on the child device (driver preferences).
 */

import groovy.json.JsonSlurper
import groovy.transform.Field
import java.net.URLEncoder

@Field static final String APP_VERSION = "1.0.8"
@Field static final String APP_NAME = "Hubitat Calendar Switch"

// IMPORTANT: must exactly match the Driver "Name" in Hubitat Drivers Code
@Field static final String CHILD_DRIVER_NAMESPACE = "truittchris"
@Field static final String CHILD_DRIVER_NAME = "Hubitat Calendar Switch - Control Device"

definition(
    name: "Hubitat Calendar Switch",
    namespace: "truittchris",
    author: "Chris Truitt",
    description: "Connects to Google and Microsoft calendars via OAuth and delivers events to child switches.",
    category: "Convenience",
    singleInstance: true,
    iconUrl: "",
    iconX2Url: "",
    iconX3Url: "",
    oauth: true
)

preferences {
    page(name: "mainPage")
    page(name: "doAddSwitchPage")
    page(name: "doRemoveSwitchPage")
    page(name: "doUpdateAllNowPage")
    page(name: "doTestSelectedPage")
}

mappings {
    path("/oauth/microsoft") { action: [GET: "microsoftOAuthCallback"] }
    path("/oauth/google")    { action: [GET: "googleOAuthCallback"] }
}

// -----------------------------------------------------------------------------
// Lifecycle
// -----------------------------------------------------------------------------

def installed() {
    ensureAccessToken()
    logBasic("Installed v${APP_VERSION}")
    initialize()
}

def updated() {
    ensureAccessToken()
    logBasic("Updated v${APP_VERSION}")
    unschedule()
    initialize()

    // Clear one-shot UI inputs so the page does not feel "sticky".
    app.updateSetting("newSwitchName", [value: "", type: "text"])
    app.updateSetting("newSwitchCalendarId", [value: "", type: "text"])
}

private void initialize() {
    runEvery1Minute("pollAllChildren")
    runIn(5, "pollAllChildren")
}

// -----------------------------------------------------------------------------
// UI
// -----------------------------------------------------------------------------

def mainPage() {
    dynamicPage(name: "mainPage", title: "${APP_NAME} (v${APP_VERSION})", install: true, uninstall: true) {

        if (state?.uiMsg) {
            section("Status") {
                paragraph("${safeString(state.uiMsg)}")
                if (state?.uiMsgAt) paragraph("Time: ${safeString(state.uiMsgAt)}")
            }
        }

        section("Overview") {
            paragraph("This app connects to Google or Microsoft using OAuth and creates calendar-driven control switches. Your switch turns on during matching calendar events and turns off when the event window ends.")
        }

        section("Setup") {
            paragraph("Authorization opens in a new browser tab. When you see success, close that tab and return here.")

            paragraph("Provider connections")

            input "gClientId", "text", title: "Google client ID", required: false
            input "gClientSecret", "password", title: "Google client secret", required: false
            paragraph(renderProviderStatus("google"))
            href url: googleAuthorizeUrl(), style: "external", required: false, title: "Authorize Google", description: "Opens in a new tab"
            if (state?.gToken?.refresh_token) {
                input "gDisconnect", "button", title: "Disconnect Google", submitOnChange: true
            }

            paragraph(" ")

            input "msClientId", "text", title: "Microsoft client ID", required: false
            input "msClientSecret", "password", title: "Microsoft client secret", required: false
            input "msTenant", "text", title: "Tenant", description: "Usually 'common'. Leave as-is unless you know you need something else.", defaultValue: "common", required: false
            paragraph(renderProviderStatus("microsoft"))
            href url: microsoftAuthorizeUrl(), style: "external", required: false, title: "Authorize Microsoft", description: "Opens in a new tab"
            if (state?.msToken?.refresh_token) {
                input "msDisconnect", "button", title: "Disconnect Microsoft", submitOnChange: true
            }

            paragraph("Add a calendar switch")
            paragraph("Add one switch per rule. Example: Work Busy, Kids Sports, Do Not Disturb.")

            input "newSwitchProvider", "enum", title: "Provider", required: false, options: ["google": "Google", "microsoft": "Microsoft"], submitOnChange: true
            input "newSwitchName", "text", title: "Switch name", required: false

            if ((settings?.newSwitchProvider ?: "") == "google") {
                input "newSwitchCalendarId", "text", title: "Google calendar ID", description: "Optional. Leave blank for primary.", required: false
            } else if ((settings?.newSwitchProvider ?: "") == "microsoft") {
                input "newSwitchCalendarId", "text", title: "Microsoft calendar", description: "Optional. Leave blank for primary.", required: false
            }

            // Action page (reliable) instead of a button handler
            href name: "goAddSwitch", page: "doAddSwitchPage", title: "Add switch", description: "Creates the child switch device now"

            paragraph("After adding a switch, open the device and set Match words and Ignore words (plus timing and filters).")

            paragraph("Your switches")

            def children = (getChildDevices() ?: []).sort { it.displayName?.toLowerCase() }
            if (children.isEmpty()) {
                paragraph("No switches yet.")
            } else {
                children.each { cd ->
                    String provider = cd.getDataValue("provider") ?: ""
                    String calId = cd.getDataValue("calendarId") ?: ""
                    String dv = safeString(cd.currentValue("driverVersion"))
                    String lastPoll = safeString(cd.currentValue("lastPoll"))
                    String lastErr = safeString(cd.currentValue("lastError"))

                    String line = "${cd.displayName} - ${provider}" + (calId ? " (${calId})" : "")
                    if (dv) line += " - driver ${dv}"
                    if (lastPoll) line += " - last update ${lastPoll}"
                    if (lastErr) line += " - error: ${lastErr}"

                    paragraph(line)
                    href url: "/device/edit/${cd.id}", title: "Open ${cd.displayName}", description: "Set rules and run Fetch now"
                }

                input "switchToRemove", "enum", title: "Remove a switch", required: false, options: children.collectEntries { [(it.deviceNetworkId): it.displayName] }
                href name: "goRemoveSwitch", page: "doRemoveSwitchPage", title: "Remove selected switch", description: "Deletes the selected child device"
            }
        }

        section("Options") {
            input "fetchIntervalMinutes", "number", title: "Fetch new events every (minutes)", defaultValue: 5, required: true
            input "fetchLookBackHours", "number", title: "Fetch window - hours back", defaultValue: 24, required: true
            input "fetchLookAheadHours", "number", title: "Fetch window - hours ahead", defaultValue: 168, required: true

            href name: "goUpdateAll", page: "doUpdateAllNowPage", title: "Update all switches now", description: "Fetches provider events and re-evaluates all switches"
            paragraph("Performance note: switches re-evaluate every minute using the most recently fetched events. If you change match rules on a device, use Fetch now on that device for an immediate refresh.")
        }

        section("Test") {
            def children = (getChildDevices() ?: []).sort { it.displayName?.toLowerCase() }
            if (!children.isEmpty()) {
                input "testSwitchDni", "enum", title: "Switch to test", required: false, options: children.collectEntries { [(it.deviceNetworkId): it.displayName] }
                href name: "goTestSelected", page: "doTestSelectedPage", title: "Test selected switch", description: "Runs a refresh for the selected switch"
            } else {
                paragraph("Add a switch first.")
            }

            paragraph("Last test time: ${safeString(state?.lastTestTime)}")
            paragraph("Last test result: ${safeString(state?.lastTestResult)}")
        }

        section("Advanced") {
            paragraph("These settings are optional and only needed in special cases.")
            input "logLevel", "enum", title: "Logging", required: true, defaultValue: "off", options: ["off": "Off (recommended)", "basic": "Basic", "debug": "Debug"]
        }

        section("Support") {
            paragraph("Name: ${APP_NAME}")
            paragraph("Author: Chris Truitt")
            paragraph("Website: https://christruitt.com")
            paragraph("GitHub: https://github.com/truittchris")
            paragraph("Support development: https://christruitt.com/tip-jar")
            paragraph("When requesting help, enable Debug logging and include a screenshot of this page.")
        }
    }
}

// Action pages (reliable server-side execution)
def doAddSwitchPage() {
    handleAddSwitch()
    return mainPage()
}

def doRemoveSwitchPage() {
    handleRemoveSwitch()
    return mainPage()
}

def doUpdateAllNowPage() {
    state.uiMsg = "Updating all switches now..."
    state.uiMsgAt = isoNow()
    pollAllChildren(true)
    state.uiMsg = "Update started. Check device events/logs for results."
    state.uiMsgAt = isoNow()
    return mainPage()
}

def doTestSelectedPage() {
    handleTestSelected()
    return mainPage()
}

// Button handler kept only for disconnect buttons
void appButtonHandler(String buttonName) {
    switch (buttonName) {
        case "gDisconnect":
            disconnectProvider("google")
            break
        case "msDisconnect":
            disconnectProvider("microsoft")
            break
        default:
            break
    }
}

// -----------------------------------------------------------------------------
// Switch creation / removal / test
// -----------------------------------------------------------------------------

private void handleAddSwitch() {
    state.uiMsg = ""
    state.uiMsgAt = isoNow()

    String provider = (settings?.newSwitchProvider ?: "").toString()
    if (!provider) {
        state.uiMsg = "Select a provider before adding a switch."
        logWarn(state.uiMsg)
        return
    }

    if (provider == "google" && !state?.gToken?.refresh_token) {
        state.uiMsg = "Google is not connected. Connect Google first."
        logWarn(state.uiMsg)
        return
    }

    if (provider == "microsoft" && !state?.msToken?.refresh_token) {
        state.uiMsg = "Microsoft is not connected. Connect Microsoft first."
        logWarn(state.uiMsg)
        return
    }

    String label = (settings?.newSwitchName ?: "Calendar Switch").toString().trim()
    if (!label) label = "Calendar Switch"

    String calId = (settings?.newSwitchCalendarId ?: "").toString().trim()
    if (!calId) calId = "primary"

    String dni = "ct-cal-${app.id}-${now()}"

    try {
        def child = addChildDevice(CHILD_DRIVER_NAMESPACE, CHILD_DRIVER_NAME, dni, [label: label, isComponent: true])
        child.updateDataValue("provider", provider)
        child.updateDataValue("calendarId", calId)
        child.updateDataValue("parentAppId", app.id.toString())

        state.uiMsg = "Added switch '${label}' (${provider}, ${calId})."
        state.uiMsgAt = isoNow()
        logBasic(state.uiMsg)

        runIn(1, "pollAllChildren")
    } catch (Exception e) {
        state.uiMsg = "Failed to add switch. ${e.message}"
        state.uiMsgAt = isoNow()
        logWarn(state.uiMsg)

        // Common root cause: driver name mismatch
        logWarn("Expected child driver name exactly: '${CHILD_DRIVER_NAME}' in namespace '${CHILD_DRIVER_NAMESPACE}'.")
    }
}

private void handleRemoveSwitch() {
    state.uiMsg = ""
    state.uiMsgAt = isoNow()

    String dni = (settings?.switchToRemove ?: "").toString()
    if (!dni) {
        state.uiMsg = "Select a switch to remove."
        return
    }

    def cd = getChildDevice(dni)
    if (!cd) {
        state.uiMsg = "Could not find the selected switch."
        return
    }

    try {
        deleteChildDevice(dni)
        state.uiMsg = "Removed switch '${cd.displayName}'."
        state.uiMsgAt = isoNow()
        logBasic(state.uiMsg)
        app.updateSetting("switchToRemove", [value: "", type: "enum"])
    } catch (Exception e) {
        state.uiMsg = "Failed to remove '${cd?.displayName}': ${e.message}"
        state.uiMsgAt = isoNow()
        logWarn(state.uiMsg)
    }
}

private void handleTestSelected() {
    String dni = (settings?.testSwitchDni ?: "").toString()
    state.lastTestTime = isoNow()

    if (!dni) {
        state.lastTestResult = "No switch selected"
        return
    }

    try {
        pollChild(dni)
        state.lastTestResult = "Triggered update"
        state.uiMsg = "Test started for selected switch."
        state.uiMsgAt = isoNow()
    } catch (Exception e) {
        state.lastTestResult = "Failed: ${e.message}"
        state.uiMsg = "Test failed: ${e.message}"
        state.uiMsgAt = isoNow()
        logWarn(state.uiMsg)
    }
}

private void disconnectProvider(String provider) {
    if (provider == "google") {
        state.gToken = null
        state.gPkceVerifier = null
        state.gOauthNonce = null
        state.gLastAuth = null
        state.gLastResult = null
        state.lastGFetchMs = null
        state.uiMsg = "Disconnected Google."
        state.uiMsgAt = isoNow()
        logBasic(state.uiMsg)
    }

    if (provider == "microsoft") {
        state.msToken = null
        state.msPkceVerifier = null
        state.msOauthNonce = null
        state.msLastAuth = null
        state.msLastResult = null
        state.lastMsFetchMs = null
        state.uiMsg = "Disconnected Microsoft."
        state.uiMsgAt = isoNow()
        logBasic(state.uiMsg)
    }
}

// -----------------------------------------------------------------------------
// Polling and child delivery
// -----------------------------------------------------------------------------

def pollAllChildren(Boolean force = false) {
    def children = getChildDevices() ?: []
    if (children.isEmpty()) return

    long minSeconds = safeLong(settings?.fetchIntervalMinutes, 5L) * 60L

    boolean msNeeded = children.any { (it.getDataValue("provider") == "microsoft") }
    boolean gNeeded  = children.any { (it.getDataValue("provider") == "google") }

    Map msResult = null
    Map gResult = null

    if (msNeeded) {
        if (force || shouldFetchProvider("microsoft", minSeconds) || !state?.msLastResult) {
            msResult = fetchMicrosoftEvents(force)
            state.msLastResult = msResult
            state.lastMsFetchMs = now()
        } else {
            msResult = state.msLastResult
        }
    }

    if (gNeeded) {
        if (force || shouldFetchProvider("google", minSeconds) || !state?.gLastResult) {
            gResult = fetchGoogleEvents(force)
            state.gLastResult = gResult
            state.lastGFetchMs = now()
        } else {
            gResult = state.gLastResult
        }
    }

    children.each { cd ->
        String provider = cd.getDataValue("provider")
        String calId = cd.getDataValue("calendarId") ?: "primary"

        Map providerMeta = [
            provider: provider,
            calendarId: calId,
            fetchedAt: isoNow(),
            hubTimeZone: location?.timeZone?.ID
        ]

        try {
            Map result = (provider == "microsoft") ? (msResult ?: [events: [], error: "No recent fetch"]) : (gResult ?: [events: [], error: "No recent fetch"])
            List events = (result?.events instanceof List) ? result.events : []

            if (provider == "google") {
                events = events.findAll { (it?.calendarId ?: "primary") == calId }
            }

            if (result?.error) providerMeta.error = result.error

            cd.evaluateEvents(providerMeta, events)
        } catch (Exception e) {
            logWarn("Error delivering events to ${cd?.displayName}: ${e.message}")
            try {
                providerMeta.error = e.message
                cd.evaluateEvents(providerMeta, [])
            } catch (Exception ignored) { }
        }
    }
}

boolean shouldFetchProvider(String provider, long minSeconds) {
    long lastMs = 0L
    if (provider == "microsoft") lastMs = safeLong(state?.lastMsFetchMs, 0L)
    if (provider == "google")    lastMs = safeLong(state?.lastGFetchMs, 0L)
    return (now() - lastMs) >= (minSeconds * 1000L)
}

// Called by child device when a user wants an immediate refresh.
void childRequestFetch(String dni, String reason = "") {
    Long last = safeLong(state?.lastChildRequestMs, 0L)
    if ((now() - last) < 3000L) return
    state.lastChildRequestMs = now()
    pollChild(dni)
}

// Called by child driver refresh() or Test.
void pollChild(String dni) {
    def cd = getChildDevice(dni)
    if (!cd) return

    String provider = cd.getDataValue("provider")
    Map result = (provider == "microsoft") ? fetchMicrosoftEvents(true) : fetchGoogleEvents(true)

    if (provider == "microsoft") {
        state.msLastResult = result
        state.lastMsFetchMs = now()
    } else if (provider == "google") {
        state.gLastResult = result
        state.lastGFetchMs = now()
    }

    Map providerMeta = [
        provider: provider,
        calendarId: cd.getDataValue("calendarId") ?: "primary",
        fetchedAt: isoNow(),
        hubTimeZone: location?.timeZone?.ID
    ]

    if (result?.error) providerMeta.error = result.error

    List events = (result?.events instanceof List) ? result.events : []
    if (provider == "google") {
        String calId = providerMeta.calendarId
        events = events.findAll { (it?.calendarId ?: "primary") == calId }
    }

    cd.evaluateEvents(providerMeta, events)
}

// -----------------------------------------------------------------------------
// OAuth helper - common
// -----------------------------------------------------------------------------

private void ensureAccessToken() {
    if (!state.accessToken) {
        try {
            state.accessToken = createAccessToken()
        } catch (Exception e) {
            logWarn("Unable to create access token. Ensure OAuth is enabled for the app in Hubitat: ${e.message}")
        }
    }
}

private String oauthStateFor(String provider) {
    String nonce = ""
    if (provider == "google") nonce = (state?.gOauthNonce ?: "").toString()
    if (provider == "microsoft") nonce = (state?.msOauthNonce ?: "").toString()

    String noncePart = nonce ? "&nonce=${urlEnc(nonce)}" : ""
    return "${getHubUID()}/apps/${app.id}/oauth/${provider}?access_token=${state.accessToken}${noncePart}"
}

private String renderProviderStatus(String provider) {
    Map tok = (provider == "microsoft") ? state?.msToken : state?.gToken
    if (!tok?.access_token) return "Status: Not connected"

    long exp = safeLong(tok?.expires_at, 0L)
    String expStr = exp ? new Date(exp).toString() : "unknown"
    String last = (provider == "microsoft") ? (state?.msLastAuth ?: "") : (state?.gLastAuth ?: "")

    return "Status: Connected\nToken expires: ${expStr}" + (last ? "\nLast authorization: ${last}" : "")
}

// -----------------------------------------------------------------------------
// OAuth - Google
// -----------------------------------------------------------------------------

private String googleAuthorizeUrl() {
    if (!settings?.gClientId) return ""

    ensureAccessToken()

    String verifier = randomString(64)
    String challenge = base64UrlSha256(verifier)
    state.gPkceVerifier = verifier
    state.gOauthNonce = randomString(20)

    String redirectUri = "https://cloud.hubitat.com/oauth/stateredirect"
    String scope = URLEncoder.encode("https://www.googleapis.com/auth/calendar.readonly", "UTF-8")
    String stateParam = URLEncoder.encode(oauthStateFor("google"), "UTF-8")

    return "https://accounts.google.com/o/oauth2/v2/auth?client_id=${settings.gClientId}" +
        "&redirect_uri=${URLEncoder.encode(redirectUri, 'UTF-8')}" +
        "&response_type=code" +
        "&scope=${scope}" +
        "&access_type=offline" +
        "&prompt=consent" +
        "&code_challenge=${challenge}" +
        "&code_challenge_method=S256" +
        "&state=${stateParam}"
}

def googleOAuthCallback() {
    try {
        String code = params?.code
        if (!code) {
            logWarn("Google OAuth callback missing code")
            return renderOAuthHtml(true, false, "Google authorization failed", "No code returned from Google.")
        }

        Map token = exchangeGoogleCodeForToken(code)
        if (token?.access_token) {
            long expiresIn = safeLong(token?.expires_in, 3600L)
            token.expires_at = now() + (expiresIn * 1000L)
            state.gToken = token
            state.gLastAuth = isoNow()
            state.gOauthNonce = null
            logBasic("Stored Google token. Expires in ${expiresIn}s")
            return renderOAuthHtml(true, true, "Google connected", "Google authorization succeeded. You can close this tab and return to Hubitat.")
        }

        String detail = token?._detail ? "\n\nDetails: ${token._detail}" : ""
        logWarn("Google token exchange failed.${detail}")
        return renderOAuthHtml(true, false, "Google authorization failed", "Google did not return an access token.${detail}")
    } catch (Exception e) {
        logWarn("Google OAuth callback exception: ${e.message}")
        return renderOAuthHtml(true, false, "Google authorization failed", e.message)
    }
}

private Map exchangeGoogleCodeForToken(String code) {
    String verifier = (state?.gPkceVerifier ?: "").toString()
    String redirectUri = "https://cloud.hubitat.com/oauth/stateredirect"

    Map body = [
        code: code,
        client_id: settings?.gClientId,
        client_secret: settings?.gClientSecret,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
        code_verifier: verifier
    ]

    def req = [
        uri: "https://oauth2.googleapis.com/token",
        requestContentType: "application/x-www-form-urlencoded",
        contentType: "application/json",
        body: body
    ]

    Map token = [:]
    Integer status = null
    String raw = null

    try {
        httpPost(req) { resp ->
            status = resp?.status as Integer
            if (resp?.data instanceof Map) token = (Map) resp.data
            else raw = resp?.data?.toString()
            if (!raw) {
                try { raw = resp?.body?.toString() } catch (ignored) { }
            }
        }
    } catch (Exception e) {
        return [error: "exception", _detail: e.message]
    }

    if (status != null && status != 200) {
        return [error: "HTTP ${status}", _detail: (raw ?: token?.toString() ?: "")]
    }

    return (token instanceof Map) ? token : [:]
}

private String googleAccessToken() {
    Map tok = (state?.gToken instanceof Map) ? (Map) state.gToken : null
    if (!tok) return null

    long expAt = safeLong(tok?.expires_at, 0L)
    if (expAt && expAt > (now() + 60000L)) return tok.access_token

    String refreshToken = tok.refresh_token
    if (!refreshToken) return tok.access_token

    Map body = [
        client_id: settings?.gClientId,
        client_secret: settings?.gClientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token"
    ]

    def req = [
        uri: "https://oauth2.googleapis.com/token",
        requestContentType: "application/x-www-form-urlencoded",
        contentType: "application/json",
        body: body
    ]

    Map refreshed = [:]
    Integer status = null

    httpPost(req) { resp ->
        status = resp?.status as Integer
        if (resp?.data instanceof Map) refreshed = (Map) resp.data
    }

    if (status != null && status != 200) {
        logWarn("Google refresh failed: HTTP ${status}")
        return tok.access_token
    }

    if (refreshed?.access_token) {
        long expiresIn = safeLong(refreshed?.expires_in, 3600L)
        tok.access_token = refreshed.access_token
        tok.expires_in = expiresIn
        tok.expires_at = now() + (expiresIn * 1000L)
        state.gToken = tok
        logDebug("Refreshed Google token")
        return tok.access_token
    }

    return tok.access_token
}

private Map fetchGoogleEvents(Boolean force = false) {
    if (!state?.gToken?.refresh_token) return [events: [], error: "Google not connected"]

    String token = googleAccessToken()
    if (!token) return [events: [], error: "Google token unavailable"]

    Integer backHrs = safeInt(settings?.fetchLookBackHours, 24)
    Integer aheadHrs = safeInt(settings?.fetchLookAheadHours, 168)

    Date timeMin = new Date(now() - (backHrs * 3600L * 1000L))
    Date timeMax = new Date(now() + (aheadHrs * 3600L * 1000L))

    String timeMinStr = timeMin.format("yyyy-MM-dd'T'HH:mm:ssXXX", location?.timeZone)
    String timeMaxStr = timeMax.format("yyyy-MM-dd'T'HH:mm:ssXXX", location?.timeZone)

    String calId = "primary"
    String url = "https://www.googleapis.com/calendar/v3/calendars/${URLEncoder.encode(calId, 'UTF-8')}/events" +
        "?singleEvents=true&orderBy=startTime" +
        "&timeMin=${URLEncoder.encode(timeMinStr, 'UTF-8')}" +
        "&timeMax=${URLEncoder.encode(timeMaxStr, 'UTF-8')}"

    def req = [
        uri: url,
        headers: [Authorization: "Bearer ${token}"]
    ]

    Map result = [events: []]

    httpGet(req) { resp ->
        if (resp?.status == 200) {
            List items = (resp?.data?.items instanceof List) ? resp.data.items : []

            List events = items.collect { Map it ->
                boolean isAllDay = (it?.start?.date != null)
                String start = isAllDay ? (it?.start?.date?.toString() + "T00:00:00" + _tzSuffix()) : it?.start?.dateTime?.toString()
                String end = isAllDay ? (it?.end?.date?.toString() + "T00:00:00" + _tzSuffix()) : it?.end?.dateTime?.toString()

                [
                    provider: "google",
                    calendarId: calId,
                    id: it?.id,
                    title: it?.summary ?: "",
                    description: it?.description ?: "",
                    location: it?.location ?: "",
                    organizer: it?.organizer?.email ?: "",
                    categories: "",
                    start: start,
                    end: end,
                    isAllDay: isAllDay,
                    isBusy: ((it?.transparency ?: "opaque").toString() != "transparent"),
                    isPrivate: ((it?.visibility ?: "default").toString() == "private")
                ]
            }

            result.events = events
        } else {
            result.error = "Google fetch failed: HTTP ${resp?.status}"
        }
    }

    return result
}

private String _tzSuffix() {
    def tz = location?.timeZone ?: TimeZone.getTimeZone("UTC")
    return new Date().format("XXX", tz)
}

// -----------------------------------------------------------------------------
// OAuth - Microsoft
// -----------------------------------------------------------------------------

private String microsoftAuthorizeUrl() {
    if (!settings?.msClientId) return ""

    ensureAccessToken()

    String tenant = (settings?.msTenant ?: "common").toString().trim()
    if (!tenant) tenant = "common"

    String verifier = randomString(64)
    String challenge = base64UrlSha256(verifier)
    state.msPkceVerifier = verifier
    state.msOauthNonce = randomString(20)

    String redirectUri = "https://cloud.hubitat.com/oauth/stateredirect"
    String scope = URLEncoder.encode("offline_access Calendars.Read", "UTF-8")
    String stateParam = URLEncoder.encode(oauthStateFor("microsoft"), "UTF-8")

    return "https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize" +
        "?client_id=${settings.msClientId}" +
        "&response_type=code" +
        "&redirect_uri=${URLEncoder.encode(redirectUri, 'UTF-8')}" +
        "&response_mode=query" +
        "&scope=${scope}" +
        "&code_challenge=${challenge}" +
        "&code_challenge_method=S256" +
        "&state=${stateParam}"
}

def microsoftOAuthCallback() {
    try {
        String code = params?.code
        if (!code) {
            logWarn("Microsoft OAuth callback missing code")
            return renderOAuthHtml(true, false, "Microsoft authorization failed", "No code returned from Microsoft.")
        }

        Map token = exchangeMicrosoftCodeForToken(code)
        if (token?.access_token) {
            long expiresIn = safeLong(token?.expires_in, 3600L)
            token.expires_at = now() + (expiresIn * 1000L)
            state.msToken = token
            state.msLastAuth = isoNow()
            state.msOauthNonce = null
            logBasic("Stored Microsoft token. Expires in ${expiresIn}s")
            return renderOAuthHtml(true, true, "Microsoft connected", "Microsoft authorization succeeded. You can close this tab and return to Hubitat.")
        }

        String detail = token?._detail ? "\n\nDetails: ${token._detail}" : ""
        logWarn("Microsoft token exchange failed.${detail}")
        return renderOAuthHtml(true, false, "Microsoft authorization failed", "Microsoft did not return an access token.${detail}")
    } catch (Exception e) {
        logWarn("Microsoft OAuth callback exception: ${e.message}")
        return renderOAuthHtml(true, false, "Microsoft authorization failed", e.message)
    }
}

private Map exchangeMicrosoftCodeForToken(String code) {
    String tenant = (settings?.msTenant ?: "common").toString().trim()
    if (!tenant) tenant = "common"

    String verifier = (state?.msPkceVerifier ?: "").toString()
    String redirectUri = "https://cloud.hubitat.com/oauth/stateredirect"

    Map body = [
        client_id: settings?.msClientId,
        client_secret: settings?.msClientSecret,
        code: code,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
        code_verifier: verifier,
        scope: "offline_access Calendars.Read"
    ]

    def req = [
        uri: "https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token",
        requestContentType: "application/x-www-form-urlencoded",
        contentType: "application/json",
        body: body
    ]

    Map token = [:]
    Integer status = null
    String raw = null

    try {
        httpPost(req) { resp ->
            status = resp?.status as Integer
            if (resp?.data instanceof Map) token = (Map) resp.data
            else raw = resp?.data?.toString()
            if (!raw) {
                try { raw = resp?.body?.toString() } catch (ignored) { }
            }
        }
    } catch (Exception e) {
        return [error: "exception", _detail: e.message]
    }

    if (status != null && status != 200) {
        return [error: "HTTP ${status}", _detail: (raw ?: token?.toString() ?: "")]
    }

    if ((token?.error ?: "").toString()) {
        return [error: token.error, _detail: (token?.error_description ?: token?.toString() ?: "")]
    }

    return (token instanceof Map) ? token : [:]
}

private String microsoftAccessToken() {
    Map tok = (state?.msToken instanceof Map) ? (Map) state.msToken : null
    if (!tok) return null

    long expAt = safeLong(tok?.expires_at, 0L)
    if (expAt && expAt > (now() + 60000L)) return tok.access_token

    String refreshToken = tok.refresh_token
    if (!refreshToken) return tok.access_token

    String tenant = (settings?.msTenant ?: "common").toString().trim()
    if (!tenant) tenant = "common"

    Map body = [
        client_id: settings?.msClientId,
        client_secret: settings?.msClientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
        scope: "offline_access Calendars.Read"
    ]

    def req = [
        uri: "https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token",
        requestContentType: "application/x-www-form-urlencoded",
        contentType: "application/json",
        body: body
    ]

    Map refreshed = [:]
    Integer status = null

    httpPost(req) { resp ->
        status = resp?.status as Integer
        if (resp?.data instanceof Map) refreshed = (Map) resp.data
    }

    if (status != null && status != 200) {
        logWarn("Microsoft refresh failed: HTTP ${status}")
        return tok.access_token
    }

    if (refreshed?.access_token) {
        long expiresIn = safeLong(refreshed?.expires_in, 3600L)
        tok.access_token = refreshed.access_token
        tok.expires_in = expiresIn
        tok.expires_at = now() + (expiresIn * 1000L)
        state.msToken = tok
        logDebug("Refreshed Microsoft token")
        return tok.access_token
    }

    return tok.access_token
}

private Map fetchMicrosoftEvents(Boolean force = false) {
    if (!state?.msToken?.refresh_token) return [events: [], error: "Microsoft not connected"]

    String token = microsoftAccessToken()
    if (!token) return [events: [], error: "Microsoft token unavailable"]

    Integer backHrs = safeInt(settings?.fetchLookBackHours, 24)
    Integer aheadHrs = safeInt(settings?.fetchLookAheadHours, 168)

    Date start = new Date(now() - (backHrs * 3600L * 1000L))
    Date end = new Date(now() + (aheadHrs * 3600L * 1000L))

    String startStr = start.format("yyyy-MM-dd'T'HH:mm:ssXXX", TimeZone.getTimeZone("UTC"))
    String endStr = end.format("yyyy-MM-dd'T'HH:mm:ssXXX", TimeZone.getTimeZone("UTC"))

    String url = "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${URLEncoder.encode(startStr, 'UTF-8')}&endDateTime=${URLEncoder.encode(endStr, 'UTF-8')}"

    def req = [
        uri: url,
        headers: [Authorization: "Bearer ${token}", Prefer: "outlook.timezone=\"${location?.timeZone?.ID ?: 'UTC'}\""]
    ]

    Map result = [events: []]

    httpGet(req) { resp ->
        if (resp?.status == 200) {
            List items = (resp?.data?.value instanceof List) ? resp.data.value : []

            List events = items.collect { Map it ->
                boolean isAllDay = safeBool(it?.isAllDay)
                boolean isPrivate = ((it?.sensitivity ?: "normal").toString() == "private")

                [
                    provider: "microsoft",
                    calendarId: "primary",
                    id: it?.id,
                    title: it?.subject ?: "",
                    description: it?.bodyPreview ?: "",
                    location: it?.location?.displayName ?: "",
                    organizer: it?.organizer?.emailAddress?.address ?: "",
                    categories: (it?.categories instanceof List) ? ((List) it.categories).join(",") : "",
                    start: it?.start?.dateTime ?: "",
                    end: it?.end?.dateTime ?: "",
                    isAllDay: isAllDay,
                    isBusy: ((it?.showAs ?: "busy").toString() != "free"),
                    isPrivate: isPrivate
                ]
            }

            result.events = events
        } else {
            result.error = "Microsoft fetch failed: HTTP ${resp?.status}"
        }
    }

    return result
}

// -----------------------------------------------------------------------------
// OAuth render helpers
// -----------------------------------------------------------------------------

private def renderOAuthHtml(boolean attemptClose, boolean success, String title, String message) {
    String status = success ? "Success" : "Error"

    // Do NOT reload the opener. Navigate it to a clean URL to avoid 431/414 loops.
    String closeJs = attemptClose ? """
  <script>
    (function() {
      try {
        if (window.opener && window.opener.location) {
          try {
            var origin = window.opener.location.origin || "";
            window.opener.location.href = origin + "/installedapp/configure/${app.id}/mainPage";
          } catch (e1) {}
        }
        setTimeout(function() {
          try { window.close(); } catch (e2) {}
        }, 150);
      } catch (e3) {}
    })();
  </script>
""" : ""

    String html = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 30px; }
    .card { max-width: 720px; padding: 18px 20px; border: 1px solid #ddd; border-radius: 10px; }
    .status { font-size: 14px; color: #666; margin-bottom: 8px; }
    h1 { margin: 0 0 8px 0; font-size: 20px; }
    p { margin: 0 0 14px 0; line-height: 1.35; white-space: pre-wrap; }
    .hint { font-size: 13px; color: #444; }
  </style>
</head>
<body>
  <div class="card">
    <div class="status">${status}</div>
    <h1>${escapeHtml(title)}</h1>
    <p>${escapeHtml(message)}</p>
    <div class="hint">If this tab does not close automatically, close it and return to Hubitat.</div>
  </div>
  ${closeJs}
</body>
</html>
"""
    render contentType: "text/html", data: html
}

private String escapeHtml(String s) {
    if (s == null) return ""
    return s.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;")
            .replace("'", "&#39;")
}

// -----------------------------------------------------------------------------
// Utility
// -----------------------------------------------------------------------------

private String isoNow() {
    def tz = location?.timeZone ?: TimeZone.getTimeZone("UTC")
    return new Date().format("yyyy-MM-dd'T'HH:mm:ssXXX", tz)
}

private long safeLong(def v, long dflt = 0L) {
    try {
        if (v == null) return dflt
        if (v instanceof Number) return ((Number) v).longValue()
        String s = v.toString().trim()
        return s ? Long.parseLong(s) : dflt
    } catch (Exception ignored) {
        return dflt
    }
}

private int safeInt(def v, int dflt = 0) {
    try {
        if (v == null) return dflt
        if (v instanceof Number) return ((Number) v).intValue()
        String s = v.toString().trim()
        return s ? Integer.parseInt(s) : dflt
    } catch (Exception ignored) {
        return dflt
    }
}

private boolean safeBool(def v) {
    if (v == null) return false
    if (v instanceof Boolean) return (Boolean) v
    return v.toString().toLowerCase() in ["true", "1", "yes", "y"]
}

private String safeString(def v) {
    return (v == null) ? "" : v.toString()
}

private String randomString(int len) {
    String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
    String out = ""
    Random r = new Random()
    for (int i = 0; i < len; i++) out += alphabet.charAt(r.nextInt(alphabet.length()))
    return out
}

private String base64UrlSha256(String s) {
    byte[] bytes = java.security.MessageDigest.getInstance("SHA-256").digest(s.getBytes("UTF-8"))
    return bytes.encodeBase64().toString().replace("+", "-").replace("/", "_").replace("=", "")
}

private String urlEnc(String s) {
    return URLEncoder.encode(s ?: "", "UTF-8")
}

// Logging helpers
private boolean isBasicEnabled() { return (settings?.logLevel ?: "off") in ["basic", "debug"] }
private boolean isDebugEnabled() { return (settings?.logLevel ?: "off") == "debug" }

private void logBasic(String msg) { if (isBasicEnabled()) log.info "[${APP_NAME}] ${msg}" }
private void logDebug(String msg) { if (isDebugEnabled()) log.debug "[${APP_NAME}] ${msg}" }
private void logWarn(String msg)  { log.warn  "[${APP_NAME}] ${msg}" }