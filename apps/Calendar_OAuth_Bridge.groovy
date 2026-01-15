/**
 * truittchris Calendar OAuth Bridge
 * Parent app (OAuth + polling + child switch orchestration)
 *
 * v0.1.2
 *
 * What it does
 * - OAuth 2.0 connect to:
 *   - Google Calendar (read-only)
 *   - Microsoft 365 / Outlook via Microsoft Graph (read-only)
 * - Polls primary calendar, evaluates events against per-switch rules
 * - Drives child switch devices ON during matching events (with before/after offsets)
 * - Publishes upcoming matching events (next 3) to each child device (Option A)
 *
 * Important Hubitat note
 * - You must enable OAuth for this app in Apps Code (the OAuth toggle) before authentication will work.
 * - The redirect URI to register at Google/Microsoft is always:
 *   https://cloud.hubitat.com/oauth/stateredirect
 *
 * v0.1.2 changes
 * - Fix "Create switch" UX: after a successful create/save/delete, the app returns to the Child switches list
 *   so required fields do not clear and block navigation.
 */

import groovy.json.JsonSlurper
import java.net.URLEncoder
import java.net.URLDecoder

def appVersion() { "0.1.3" }

definition(
    name: "Calendar OAuth Bridge",
    namespace: "truittchris",
    author: "Chris Truitt",
    description: "OAuth connect to Google or Microsoft calendar and drive child switches from calendar events.",
    category: "Convenience",
    iconUrl: "https://s3.amazonaws.com/smartapp-icons/Convenience/Cat-Convenience.png",
    iconX2Url: "https://s3.amazonaws.com/smartapp-icons/Convenience/Cat-Convenience@2x.png",
    iconX3Url: "https://s3.amazonaws.com/smartapp-icons/Convenience/Cat-Convenience@2x.png",
    singleInstance: false
)

preferences {
    page(name: "mainPage")
    page(name: "authenticationPage")
    page(name: "devicesPage")
    page(name: "editDevicePage")
    page(name: "removePage")
    page(name: "authenticationResetPage")
}

mappings {
    path("/callback") { action: [GET: "callback"] }
}

/* ---------------------------
 * Lifecycle
 * --------------------------*/

def installed() {
    logInfo("Installed")
    initialize()
}

def updated() {
    logInfo("Updated")
    unsubscribe()
    unschedule()
    initialize()
}

def initialize() {
    // Normalize credential whitespace
    if (settings.gaClientID && settings.gaClientID != settings.gaClientID.trim()) {
        app.updateSetting("gaClientID", [type: "text", value: settings.gaClientID.trim()])
    }
    if (settings.gaClientSecret && settings.gaClientSecret != settings.gaClientSecret.trim()) {
        app.updateSetting("gaClientSecret", [type: "text", value: settings.gaClientSecret.trim()])
    }
    if (settings.msClientId && settings.msClientId != settings.msClientId.trim()) {
        app.updateSetting("msClientId", [type: "text", value: settings.msClientId.trim()])
    }
    if (settings.msClientSecret && settings.msClientSecret != settings.msClientSecret.trim()) {
        app.updateSetting("msClientSecret", [type: "password", value: settings.msClientSecret.trim()])
    }
    if (settings.msTenant && settings.msTenant != settings.msTenant.trim()) {
        app.updateSetting("msTenant", [type: "text", value: settings.msTenant.trim()])
    }

    // Ensure base structures
    state.switches = (state.switches instanceof Map) ? state.switches : [:]
    state.nextIndex = (state.nextIndex instanceof Number) ? (state.nextIndex as Integer) : 1
    state.goToDevicesPage = (state.goToDevicesPage == true)

    // Ensure OAuth token for callback access
    oauthEnabled()

    // Set label if desired
    if (settings.appName) {
        app.updateLabel(settings.appName)
    }

    schedulePolling()
}

def uninstalled() {
    try {
        revokeAccess()
    } catch (e) {
        logWarn("Revoke error: ${e}")
    }
    try {
        getChildDevices()?.each { deleteChildDevice(it.deviceNetworkId) }
    } catch (e) {
        logWarn("Child cleanup error: ${e}")
    }
}

/* ---------------------------
 * UI pages
 * --------------------------*/

def mainPage() {
    def oauthOk = oauthEnabled()
    def authorized = authTokenValid("mainPage")

    dynamicPage(name: "mainPage", title: "Calendar OAuth Bridge v${appVersion()}", install: true, uninstall: false) {

        section("Options") {
            input name: "appName", type: "text", title: "Name this app", required: true, defaultValue: "Calendar OAuth Bridge", submitOnChange: true
            input name: "provider", type: "enum", title: "Provider", required: true, submitOnChange: true,
                options: ["google": "Google Calendar", "microsoft": "Microsoft 365 / Outlook (Graph)"]
            input name: "isDebugEnabled", type: "bool", title: "Enable debug logging?", required: false, defaultValue: false, submitOnChange: true
            input name: "debugAuth", type: "bool", title: "Debug authentication?", required: false, defaultValue: false, submitOnChange: true
        }

        section("Polling") {
            input name: "pollSeconds", type: "number", title: "Poll interval (seconds)", required: true, defaultValue: 60
            input name: "windowHours", type: "number", title: "Lookahead window (hours)", required: true, defaultValue: 24
            input name: "lookbackHours", type: "number", title: "Lookback window (hours)", required: true, defaultValue: 6
            input "runNow", "button", title: "Poll now", submitOnChange: true
            if (state.lastPoll) {
                paragraph("Last poll: ${new Date(state.lastPoll as Long)}")
            }
            if (state.lastError) {
                paragraph("Last error: ${state.lastError}")
            }
        }

        section("Authentication") {
            if (!oauthOk) {
                paragraph("OAuth is not enabled for this app. Open Apps Code and enable the OAuth toggle for this app, then return here.")
            } else {
                paragraph("Redirect URI to register at the provider:\n${getRedirectURL()}")
                paragraph("Authorized: ${authorized ? "yes" : "no"}")
                href(name: "auth", page: "authenticationPage", title: "Provider authorization", description: "Authenticate, re-authenticate, or reset")
            }
        }

        section("Child switches") {
            href(name: "devices", page: "devicesPage", title: "Manage child switches", description: "Add or edit switches and rules")
            paragraph("Configured switches: ${state.switches?.size() ?: 0}")
        }

        section("Removal") {
            href(name: "remove", page: "removePage", title: "Remove this app", description: "Uninstall and remove child devices")
        }
    }
}

def authenticationPage() {
    def oauthOk = oauthEnabled()
    def authorized = authTokenValid("authenticationPage")

    dynamicPage(name: "authenticationPage", title: "Provider authorization", install: false, uninstall: false, nextPage: "mainPage") {
        if (!oauthOk) {
            section("OAuth not enabled") {
                paragraph(oAuthInstructions())
            }
            return
        }

        section("Credentials") {
            if (!settings.provider) {
                paragraph("Select a provider on the main page first.")
                return
            }

            if (settings.provider == "google") {
                paragraph("Enter your Google OAuth Client ID and Client Secret. OAuth client must be type: Web application.")
                input "gaClientID", "text", title: "Google Client ID", required: true, submitOnChange: true
                input "gaClientSecret", "text", title: "Google Client Secret", required: true, submitOnChange: true
            } else if (settings.provider == "microsoft") {
                paragraph("Enter your Microsoft Entra app registration Client ID and Client Secret. Tenant can be 'common', 'organizations', 'consumers', or a specific tenant ID.")
                input "msTenant", "text", title: "Tenant", required: false, defaultValue: "common", submitOnChange: true
                input "msClientId", "text", title: "Client ID", required: true, submitOnChange: true
                input "msClientSecret", "password", title: "Client Secret", required: true, submitOnChange: true
            }
        }

        section("Authorize") {
            if (!credentialsPresent()) {
                paragraph("Enter credentials above to enable the authorization link.")
            } else if (!authorized) {
                paragraph(authenticationInstructions())
                href url: getOAuthInitUrl(), style: "external", required: true, title: "Authenticate", description: "Tap to start the OAuth process"
            } else {
                paragraph("Authentication is complete.")
            }
            href "authenticationResetPage", title: "Reset authorization", description: "Revoke tokens (best-effort) and clear stored tokens"
        }
    }
}

def authenticationResetPage() {
    revokeAccess()
    atomicState.authToken = null
    atomicState.refreshToken = null
    atomicState.tokenExpires = null
    atomicState.scopesAuthorized = null
    atomicState.provider = null
    state.lastError = null

    dynamicPage(name: "authenticationResetPage", title: "Authorization reset", install: false, uninstall: false, nextPage: "authenticationPage") {
        section("Done") {
            paragraph("Authorization has been reset. Return to the authorization page and authenticate again.")
        }
    }
}

def devicesPage() {
    dynamicPage(name: "devicesPage", title: "Child switches", install: false, uninstall: false, nextPage: "mainPage") {
        if (state.deviceSaveMsg) {
            section("Status") {
                paragraph(state.deviceSaveMsg)
            }
            state.deviceSaveMsg = null
        }
        if (state.lastError) {
            section("Last error") {
                paragraph(state.lastError)
            }
        }

        section("Add") {
            href(name: "add", page: "editDevicePage", title: "Add a new switch", description: "Create a new child device and set rules", params: [dni: "NEW"])
        }
        section("Existing") {
            if (!state.switches || state.switches.size() == 0) {
                paragraph("No child switches configured.")
            } else {
                state.switches.keySet().sort().each { dni ->
                    def cfg = state.switches[dni] ?: [:]
                    def label = cfg?.label ?: dni
                    href(
                        name: "edit_${dni}",
                        page: "editDevicePage",
                        title: label,
                        description: "Edit rules",
                        params: [dni: dni]
                    )
                }
            }
        }
    }
}

def editDevicePage(params) {
    // UX fix: after a successful create/save/delete button click, return to list page immediately.
    if (state.goToDevicesPage == true) {
        state.goToDevicesPage = false
        return devicesPage()
    }

    def dni = params?.dni ?: "NEW"
    def isNew = (dni == "NEW")
    def cfg = (!isNew && state.switches?.get(dni) instanceof Map) ? (state.switches[dni] as Map) : [:]

    // Track which DNI is being edited for button handling
    state.editingDni = dni

    dynamicPage(name: "editDevicePage", title: isNew ? "Add switch" : "Edit switch", install: false, uninstall: false, nextPage: "devicesPage") {
        if (state.lastError) {
            section("Last error") {
                paragraph(state.lastError)
            }
        }
        if (state.deviceSaveMsg) {
            section("Status") {
                paragraph(state.deviceSaveMsg)
            }
        }

        section("Device") {
            input name: "tmpLabel", type: "text", title: "Switch name", required: true, defaultValue: (isNew ? "" : (cfg.label ?: ""))
            input name: "tmpBeforeMin", type: "number", title: "Minutes before start to turn on", required: true, defaultValue: (cfg.beforeMin != null ? cfg.beforeMin : 0)
            input name: "tmpAfterMin", type: "number", title: "Minutes after end to turn off", required: true, defaultValue: (cfg.afterMin != null ? cfg.afterMin : 0)
        }

        section("Match rules") {
            input name: "tmpInclude", type: "text", title: "Include keywords (comma-separated, optional)", required: false, defaultValue: (cfg.include ?: "")
            input name: "tmpExclude", type: "text", title: "Exclude keywords (comma-separated, optional)", required: false, defaultValue: (cfg.exclude ?: "")
            input name: "tmpRequireBusy", type: "bool", title: "Require busy/opaque events", required: false, defaultValue: (cfg.requireBusy ?: false)
            input name: "tmpIncludeAllDay", type: "bool", title: "Allow all-day events", required: false, defaultValue: (cfg.includeAllDay ?: false)
            input name: "tmpIncludePrivate", type: "bool", title: "Allow private events", required: false, defaultValue: (cfg.includePrivate != null ? cfg.includePrivate : true)
        }

        section("Actions") {
            input "btnSaveSwitch", "button", title: isNew ? "Create switch" : "Save changes", submitOnChange: true
            if (!isNew) {
                input "btnDeleteSwitch", "button", title: "Delete switch", submitOnChange: true
            }
        }
    }
}

def removePage() {
    dynamicPage(name: "removePage", title: "Remove Calendar OAuth Bridge", install: false, uninstall: true) {
        section() {
            paragraph("Uninstalling will revoke access (best-effort), remove all child devices, and delete all configurations.")
        }
    }
}

/* ---------------------------
 * Buttons
 * --------------------------*/

def appButtonHandler(btn) {
    switch (btn) {
        case "runNow":
            poll()
            break
        case "btnSaveSwitch":
            saveSwitchFromTemp()
            break
        case "btnDeleteSwitch":
            deleteSwitch(state.editingDni as String)
            break
    }
}

/* ---------------------------
 * Child device management
 * --------------------------*/

private void saveSwitchFromTemp() {
    def editing = (state.editingDni ?: "NEW") as String
    def isNew = (editing == "NEW")

    def label = (settings.tmpLabel ?: "").toString().trim()
    if (!label) {
        state.lastError = "Switch name is required."
        state.deviceSaveMsg = null
        logWarn(state.lastError)
        return
    }

    Integer beforeMin = safeInt(settings.tmpBeforeMin, 0)
    Integer afterMin = safeInt(settings.tmpAfterMin, 0)
    beforeMin = Math.max(0, beforeMin)
    afterMin = Math.max(0, afterMin)

    def include = (settings.tmpInclude ?: "").toString()
    def exclude = (settings.tmpExclude ?: "").toString()

    def requireBusy = (settings.tmpRequireBusy == true)
    def includeAllDay = (settings.tmpIncludeAllDay == true)
    def includePrivate = (settings.tmpIncludePrivate != null) ? (settings.tmpIncludePrivate == true) : true

    // Copy-on-write update to avoid nested-state persistence edge cases
    Map switches = (state.switches instanceof Map) ? new LinkedHashMap(state.switches as Map) : new LinkedHashMap()

    String dni
    if (isNew) {
        dni = nextDni()
        try {
            createChildSwitch(dni, label)
        } catch (e) {
            state.lastError = "Unable to create child device. Verify the 'Calendar Control Switch' driver is installed. Details: ${e}"
            state.deviceSaveMsg = null
            logWarn(state.lastError)
            return
        }
    } else {
        dni = editing
        def child = getChildDevice(dni)
        if (!child) {
            try {
                createChildSwitch(dni, label)
            } catch (e) {
                state.lastError = "Unable to re-create missing child device. Details: ${e}"
                state.deviceSaveMsg = null
                logWarn(state.lastError)
                return
            }
        } else {
            child.setLabel(label)
        }
    }

    switches[dni] = [
        dni: dni,
        label: label,
        beforeMin: beforeMin,
        afterMin: afterMin,
        include: include,
        exclude: exclude,
        requireBusy: requireBusy,
        includeAllDay: includeAllDay,
        includePrivate: includePrivate
    ]

    state.switches = switches
    state.lastError = null
    state.deviceSaveMsg = (isNew ? "Switch created: ${label}" : "Switch saved: ${label}")
    logInfo("Saved switch: ${label} (${dni})")

    // UX: return to list page after save
    state.goToDevicesPage = true
}


private void deleteSwitch(String dni) {
    if (!dni || dni == "NEW") return
    def cfg = (state.switches instanceof Map) ? (state.switches as Map).get(dni) : null

    try {
        deleteChildDevice(dni)
    } catch (e) {
        logWarn("Delete child failed (${dni}): ${e}")
    }

    try {
        Map switches = (state.switches instanceof Map) ? new LinkedHashMap(state.switches as Map) : new LinkedHashMap()
        switches.remove(dni)
        state.switches = switches
    } catch (ignored) { }

    clearTempDeviceInputs()
    state.lastError = null
    state.deviceSaveMsg = "Switch deleted: ${cfg?.label ?: dni}"
    state.goToDevicesPage = true

    logInfo("Deleted switch: ${cfg?.label ?: dni}")
}


private String nextDni() {
    Integer idx = (state.nextIndex instanceof Number) ? (state.nextIndex as Integer) : 1
    String dni = "calOauthSwitch-${app.id}-${idx}"
    state.nextIndex = idx + 1
    return dni
}

private void createChildSwitch(String dni, String label) {
    addChildDevice("truittchris", "Calendar Control Switch", dni, [label: label, isComponent: false])
}

private void clearTempDeviceInputs() {
    ["tmpLabel","tmpBeforeMin","tmpAfterMin","tmpInclude","tmpExclude","tmpRequireBusy","tmpIncludeAllDay","tmpIncludePrivate"].each { k ->
        try { app.removeSetting(k) } catch (ignored) { }
    }
}

/* ---------------------------
 * Polling
 * --------------------------*/

def schedulePolling() {
    Integer seconds = safeInt(settings.pollSeconds, 60)
    if (seconds < 30) seconds = 30
    if (seconds > 3600) seconds = 3600

    runIn(seconds, "poll", [overwrite: true])
}

def poll() {
    try {
        if (!authTokenValid("poll")) {
            state.lastError = "Not authorized. Complete provider authorization."
            logWarn(state.lastError)
            return
        }

        def events = fetchEvents()
        evaluateAndDrive(events)
        state.lastPoll = now()
        state.lastError = null
    } catch (e) {
        state.lastError = e.toString()
        logWarn("Poll failed: ${e}")
    } finally {
        schedulePolling()
    }
}

def childRefresh(String dni) {
    try {
        if (!authTokenValid("childRefresh")) return
        def events = fetchEvents()
        evaluateAndDrive(events, dni)
        state.lastPoll = now()
        state.lastError = null
    } catch (e) {
        state.lastError = e.toString()
        logWarn("Child refresh failed: ${e}")
    }
}

/* ---------------------------
 * Evaluate
 * --------------------------*/

private void evaluateAndDrive(List<Map> events, String onlyDni = null) {
    if (!(state.switches instanceof Map) || state.switches.size() == 0) return
    def nowMs = now()

    def sorted = (events ?: [])
        .findAll { it?.startMs != null && it?.endMs != null }
        .sort { a, b -> (a.startMs as Long) <=> (b.startMs as Long) }

    state.switches.keySet().each { dni ->
        if (onlyDni && dni != onlyDni) return

        def cfg = state.switches[dni] ?: [:]
        def child = getChildDevice(dni)
        if (!child) return

        Integer beforeMin = safeInt(cfg.beforeMin, 0)
        Integer afterMin = safeInt(cfg.afterMin, 0)

        def incList = csvToList(cfg.include)
        def excList = csvToList(cfg.exclude)

        boolean requireBusy = (cfg.requireBusy == true)
        boolean includeAllDay = (cfg.includeAllDay == true)
        boolean includePrivate = (cfg.includePrivate != null) ? (cfg.includePrivate == true) : true

        boolean active = false
        Integer activeCount = 0
        Map firstActive = null

        List<Map> upcoming = []

        sorted.each { ev ->
            if (!includeAllDay && ev.isAllDay == true) return
            if (requireBusy && ev.busyFlag == false) return
            if (!includePrivate && ev.privateFlag == true) return
            if (!matchesKeywords(ev.title ?: "", incList, excList)) return

            Long onAt = (ev.startMs as Long) - (beforeMin * 60_000L)
            Long offAt = (ev.endMs as Long) + (afterMin * 60_000L)

            if (nowMs >= onAt && nowMs < offAt) {
                active = true
                activeCount = activeCount + 1
                if (!firstActive) firstActive = ev
                return
            }

            if (nowMs < onAt && upcoming.size() < 3) {
                upcoming << [
                    provider: ev.provider ?: settings.provider,
                    title: ev.title ?: "",
                    onAtMs: onAt,
                    offAtMs: offAt,
                    startMs: ev.startMs,
                    endMs: ev.endMs
                ]
            }
        }

                def nextEv = (upcoming && upcoming.size() > 0) ? upcoming[0] : null
        def upcomingSummary = upcoming.collect { ev -> "${formatInstant(ev.onAtMs as Long)} – ${(ev.title ?: "Busy")}" }.join(" | ")

Map meta = [
            activeCount: activeCount,
            provider: firstActive?.provider ?: settings.provider,
            title: firstActive?.title ?: "",
            startMs: firstActive?.startMs,
            endMs: firstActive?.endMs,
            nextEvents: upcoming,
            upcomingSummary: upcomingSummary,
            nextTitle: nextEv?.title,
            nextStartMs: nextEv?.startMs,
            nextEndMs: nextEv?.endMs,
            nextOnAtMs: nextEv?.onAtMs,
            nextOffAtMs: nextEv?.offAtMs,
            lastPoll: state.lastPoll,
            lastError: state.lastError
        ]

        child.setCalendarState(active, meta)

        if (isDebugEnabled) {
            logDebug("Device ${child.displayName} -> ${active ? "ON" : "OFF"} (matches: ${activeCount}, upcoming: ${upcoming.size()})")
        }
    }
}

private boolean matchesKeywords(String title, List<String> includeList, List<String> excludeList) {
    def t = (title ?: "").toString().toLowerCase()

    if (excludeList?.size()) {
        for (String ex : excludeList) {
            if (ex && t.contains(ex.toLowerCase())) return false
        }
    }

    if (!includeList || includeList.size() == 0) return true

    for (String inc : includeList) {
        if (inc && t.contains(inc.toLowerCase())) return true
    }
    return false
}

/* ---------------------------
 * Fetch events
 * --------------------------*/

private List<Map> fetchEvents() {
    def tz = location?.timeZone ?: TimeZone.getTimeZone("UTC")
    Long nowMs = now()
    Long lookbackMs = safeInt(settings.lookbackHours, 6) * 60L * 60L * 1000L
    Long lookaheadMs = safeInt(settings.windowHours, 24) * 60L * 60L * 1000L

    Long startMs = nowMs - lookbackMs
    Long endMs = nowMs + lookaheadMs

    if (settings.provider == "google") {
        return fetchGoogleEvents(startMs, endMs, tz)
    } else if (settings.provider == "microsoft") {
        return fetchMicrosoftEvents(startMs, endMs, tz)
    }
    return []
}

private List<Map> fetchGoogleEvents(Long startMs, Long endMs, TimeZone tz) {
    def timeMin = toIsoRfc3339(startMs)
    def timeMax = toIsoRfc3339(endMs)

    def uri = "https://www.googleapis.com"
    def path = "/calendar/v3/calendars/primary/events"
    def queryParams = [
        timeMin: timeMin,
        timeMax: timeMax,
        singleEvents: "true",
        orderBy: "startTime",
        maxResults: "250"
    ]

    def resp = apiGet("fetchGoogleEvents", uri, path, queryParams, [:])
    if (!(resp instanceof Map)) return []

    def items = resp?.items ?: []
    List<Map> out = []
    items.each { ev -> out << normalizeGoogleEvent(ev, tz) }
    return out
}

private Map normalizeGoogleEvent(def ev, TimeZone tz) {
    def startRaw = ev?.start?.dateTime ?: ev?.start?.date
    def endRaw = ev?.end?.dateTime ?: ev?.end?.date

    Long startMs = parseGoogleToMillis(startRaw, tz)
    Long endMs = parseGoogleToMillis(endRaw, tz)

    boolean isAllDay = (ev?.start?.date && !ev?.start?.dateTime)

    def transparency = String.valueOf(ev?.transparency ?: "")
    boolean busyFlag = !(transparency?.toLowerCase() == "transparent")

    def visibility = String.valueOf(ev?.visibility ?: "")
    boolean privateFlag = (visibility?.toLowerCase() == "private")

    return [
        provider: "google",
        id: ev?.id,
        title: ev?.summary ?: "",
        startMs: startMs,
        endMs: endMs,
        isAllDay: isAllDay,
        busyFlag: busyFlag,
        privateFlag: privateFlag
    ]
}

private Long parseGoogleToMillis(String s, TimeZone tz) {
    if (!s) return null

    // All-day date only (yyyy-MM-dd)
    if (s.size() == 10 && s[4] == '-' && s[7] == '-') {
        def sdf = new java.text.SimpleDateFormat("yyyy-MM-dd")
        sdf.setTimeZone(tz)
        return sdf.parse(s).time
    }

    // RFC3339 dateTime
    def patterns = [
        "yyyy-MM-dd'T'HH:mm:ss.SSSXXX",
        "yyyy-MM-dd'T'HH:mm:ssXXX",
        "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'",
        "yyyy-MM-dd'T'HH:mm:ss'Z'"
    ]
    for (p in patterns) {
        try {
            def sdf = new java.text.SimpleDateFormat(p)
            if (p.endsWith("'Z'")) sdf.setTimeZone(TimeZone.getTimeZone("UTC"))
            return sdf.parse(s).time
        } catch (ignored) { }
    }

    // Fallback: strip fractional seconds if present
    try {
        def trimmed = s.replaceAll("\\.\\d+","")
        def sdf = new java.text.SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssXXX")
        return sdf.parse(trimmed).time
    } catch (e) {
        return null
    }
}

private List<Map> fetchMicrosoftEvents(Long startMs, Long endMs, TimeZone tz) {
    def startIso = toIsoLocal(startMs, tz)
    def endIso = toIsoLocal(endMs, tz)

    def uri = "https://graph.microsoft.com"
    def path = "/v1.0/me/calendarView"
    def queryParams = [
        startDateTime: startIso,
        endDateTime: endIso,
        "\$top": "250",
        "\$select": "id,subject,start,end,isAllDay,showAs,sensitivity"
    ]
    def headersExtra = [
        "Prefer": "outlook.timezone=\"${tz.getID()}\""
    ]

    def resp = apiGet("fetchMicrosoftEvents", uri, path, queryParams, headersExtra)
    if (!(resp instanceof Map)) return []

    def items = resp?.value ?: []
    List<Map> out = []
    items.each { ev -> out << normalizeMicrosoftEvent(ev, tz) }
    return out
}

private Map normalizeMicrosoftEvent(def ev, TimeZone tz) {
    def startStr = ev?.start?.dateTime
    def endStr = ev?.end?.dateTime

    Long startMs = parseGraphLocalToMillis(startStr, tz)
    Long endMs = parseGraphLocalToMillis(endStr, tz)

    boolean busyFlag = (String.valueOf(ev?.showAs ?: "")?.toLowerCase() != "free")
    boolean privateFlag = (String.valueOf(ev?.sensitivity ?: "")?.toLowerCase() == "private")

    return [
        provider: "microsoft",
        id: ev?.id,
        title: ev?.subject ?: "",
        startMs: startMs,
        endMs: endMs,
        isAllDay: (ev?.isAllDay == true),
        busyFlag: busyFlag,
        privateFlag: privateFlag
    ]
}

private Long parseGraphLocalToMillis(String s, TimeZone tz) {
    if (!s) return null
    def patterns = [
        "yyyy-MM-dd'T'HH:mm:ss.SSSSSSS",
        "yyyy-MM-dd'T'HH:mm:ss.SSS",
        "yyyy-MM-dd'T'HH:mm:ss"
    ]
    for (p in patterns) {
        try {
            def sdf = new java.text.SimpleDateFormat(p)
            sdf.setTimeZone(tz)
            return sdf.parse(s).time
        } catch (ignored) { }
    }
    try {
        def t = (s.size() >= 19) ? s.substring(0, 19) : s
        def sdf = new java.text.SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss")
        sdf.setTimeZone(tz)
        return sdf.parse(t).time
    } catch (e) {
        return null
    }
}

/* ---------------------------
 * OAuth plumbing
 * --------------------------*/

private getRedirectURL() { "https://cloud.hubitat.com/oauth/stateredirect" }

// Keep the state value simple. Do not add extra '&' parameters.
private String oauthInitState() {
    oauthEnabled()
    return "${getApiServerUrl()}/apps/${app.id}/callback?access_token=${state.accessToken}"
}

def oauthEnabled() {
    def ok = false
    if (state.accessToken) {
        ok = true
    } else {
        try {
            def accessToken = createAccessToken()
            state.accessToken = accessToken
            ok = true
        } catch (e) {
            if (e.toString().indexOf("OAuth is not enabled for this App") > -1) {
                log.error "OAuth must be enabled for this app in Apps Code."
            } else {
                log.error "${e}"
            }
            ok = false
        }
    }

    if (ok && !state.oauthInitState) {
        state.oauthInitState = oauthInitState()
    }
    return ok
}

private boolean credentialsPresent() {
    if (!settings.provider) return false
    if (settings.provider == "google") return (settings.gaClientID && settings.gaClientSecret)
    if (settings.provider == "microsoft") return (settings.msClientId && settings.msClientSecret)
    return false
}

def getOAuthInitUrl() {
    oauthEnabled()
    if (!state.oauthInitState) state.oauthInitState = oauthInitState()

    if (settings.provider == "google") {
        def oauthParams = [
            response_type: "code",
            access_type: "offline",
            prompt: "consent",
            client_id: settings.gaClientID,
            state: state.oauthInitState,
            redirect_uri: getRedirectURL(),
            scope: "https://www.googleapis.com/auth/calendar.readonly"
        ]
        return "https://accounts.google.com/o/oauth2/v2/auth?" + toQueryString(oauthParams)
    }

    if (settings.provider == "microsoft") {
        def tenant = (settings.msTenant ?: "common").toString().trim()
        def oauthParams = [
            client_id: settings.msClientId,
            response_type: "code",
            redirect_uri: getRedirectURL(),
            response_mode: "query",
            scope: "offline_access Calendars.Read",
            prompt: "select_account",
            state: state.oauthInitState
        ]
        return "https://login.microsoftonline.com/${urlEnc(tenant)}/oauth2/v2.0/authorize?" + toQueryString(oauthParams)
    }

    return ""
}

def callback() {
    def code = params.code
    def oauthState = params.state

    def decodedState = oauthState ? safeUrlDecode(oauthState) : null
    if (oauthState && state.oauthInitState && (oauthState != state.oauthInitState) && (decodedState != state.oauthInitState)) {
        // Hubitat Cloud stateredirect does not always pass the original state back to the app callback.
        logWarn("callback() state mismatch (non-fatal). received=${oauthState} decoded=${decodedState} expected=${state.oauthInitState}")
    }

    if (!code) {
        log.error "callback() missing code. Params: ${params}"
        return connectionStatus("<p>Authorization failed. Missing authorization code.</p>")
    }

    try {
        def tokenData = exchangeCodeForToken(code)
        atomicState.provider = settings.provider
        atomicState.authToken = tokenData.access_token
        if (tokenData.refresh_token) atomicState.refreshToken = tokenData.refresh_token
        atomicState.tokenExpires = now() + ((safeInt(tokenData.expires_in, 3600) - 60) * 1000L)
        atomicState.scopesAuthorized = tokenData.scope ?: ""
        state.lastError = null

        state.oauthInitState = null
        logInfo("OAuth completed for ${settings.provider}")
        connectionStatus("<p>Authorization successful.</p><p>Close this page and return to Hubitat.</p>")
    } catch (e) {
        log.error "callback() token exchange failed: ${e}"
        state.lastError = e.toString()
        connectionStatus("<p>Authorization failed during token exchange.</p><p>Close this page and try again.</p>")
    }
}

private Map exchangeCodeForToken(String code) {
    if (settings.provider == "google") {
        def tokenUrl = "https://oauth2.googleapis.com/token"
        def tokenParams = [
            code: code,
            client_id: settings.gaClientID,
            client_secret: settings.gaClientSecret,
            redirect_uri: getRedirectURL(),
            grant_type: "authorization_code"
        ]
        return tokenPost("exchangeCodeForToken", tokenUrl, tokenParams)
    }

    if (settings.provider == "microsoft") {
        def tenant = (settings.msTenant ?: "common").toString().trim()
        def tokenUrl = "https://login.microsoftonline.com/${urlEnc(tenant)}/oauth2/v2.0/token"
        def tokenParams = [
            code: code,
            client_id: settings.msClientId,
            client_secret: settings.msClientSecret,
            redirect_uri: getRedirectURL(),
            grant_type: "authorization_code",
            scope: "offline_access Calendars.Read"
        ]
        return tokenPost("exchangeCodeForToken", tokenUrl, tokenParams)
    }

    throw new RuntimeException("Provider not set")
}

private Map refreshAuthToken() {
    if (!atomicState.refreshToken) return [ok: false]

    if (settings.provider == "google") {
        def tokenUrl = "https://oauth2.googleapis.com/token"
        def tokenParams = [
            refresh_token: atomicState.refreshToken,
            client_id: settings.gaClientID,
            client_secret: settings.gaClientSecret,
            grant_type: "refresh_token"
        ]
        def data = tokenPost("refreshAuthToken", tokenUrl, tokenParams)
        if (data?.access_token) {
            atomicState.authToken = data.access_token
            atomicState.tokenExpires = now() + ((safeInt(data.expires_in, 3600) - 60) * 1000L)
            return [ok: true]
        }
        return [ok: false]
    }

    if (settings.provider == "microsoft") {
        def tenant = (settings.msTenant ?: "common").toString().trim()
        def tokenUrl = "https://login.microsoftonline.com/${urlEnc(tenant)}/oauth2/v2.0/token"
        def tokenParams = [
            refresh_token: atomicState.refreshToken,
            client_id: settings.msClientId,
            client_secret: settings.msClientSecret,
            grant_type: "refresh_token",
            scope: "offline_access Calendars.Read"
        ]
        def data = tokenPost("refreshAuthToken", tokenUrl, tokenParams)
        if (data?.access_token) {
            atomicState.authToken = data.access_token
            if (data.refresh_token) atomicState.refreshToken = data.refresh_token
            atomicState.tokenExpires = now() + ((safeInt(data.expires_in, 3600) - 60) * 1000L)
            return [ok: true]
        }
        return [ok: false]
    }

    return [ok: false]
}

private boolean authTokenValid(fromFunction) {
    if (!atomicState.authToken || !atomicState.tokenExpires) return false
    if ((atomicState.tokenExpires as Long) >= now()) return true
    return (refreshAuthToken()?.ok == true)
}

private Map tokenPost(String fromFunction, String url, Map bodyParams) {
    Map out = [:]

    httpPost([
        uri: url,
        body: bodyParams,
        requestContentType: "application/x-www-form-urlencoded",
        contentType: "application/json"
    ]) { resp ->
        if (resp?.status != 200) throw new RuntimeException("Token endpoint HTTP ${resp?.status}")
        def data = resp.data
        if (data instanceof Map) out = data
        else if (data != null) out = (new JsonSlurper().parseText(data.toString()) as Map)
    }

    if (debugAuth) logDebug("tokenPost ${fromFunction} keys=${out?.keySet()}", "auth")
    return out
}

def revokeAccess() {
    // Best-effort revoke for Google only.
    if (settings.provider == "google" && atomicState.authToken) {
        try {
            def uri = "https://accounts.google.com/o/oauth2/revoke?token=${atomicState.authToken}"
            httpGet(uri) { resp -> }
        } catch (e) {
            logWarn("Google revoke failed: ${e}")
        }
    }
}

/* ---------------------------
 * Provider API wrappers
 * --------------------------*/

def apiGet(fromFunction, uri, path, queryParams, Map headersExtra = [:]) {
    if (!authTokenValid(fromFunction)) return null

    def headers = [
        "Authorization": "Bearer ${atomicState.authToken}",
        "Content-Type": "application/json"
    ]
    headersExtra?.each { k, v -> headers[k] = v }

    def apiParams = [
        uri: uri,
        path: path,
        headers: headers,
        query: queryParams
    ]

    try {
        def apiResponse
        httpGet(apiParams) { resp -> apiResponse = resp.data }
        return apiResponse
    } catch (e) {
        // If 401, refresh and retry once
        try {
            if (e?.response?.status == 401 && refreshAuthToken()?.ok == true) {
                return apiGet(fromFunction, uri, path, queryParams, headersExtra)
            }
        } catch (ignored) { }
        log.error "apiGet - ${fromFunction} ${path} error: ${e}"
        return null
    }
}

/* ---------------------------
 * Helpers
 * --------------------------*/

private String toIsoRfc3339(Long ms) {
    def d = new Date(ms)
    return d.format("yyyy-MM-dd'T'HH:mm:ss'Z'", TimeZone.getTimeZone("UTC"))
}

private String toIsoLocal(Long ms, TimeZone tz) {
    def d = new Date(ms)
    return d.format("yyyy-MM-dd'T'HH:mm:ss", tz)
}

private List<String> csvToList(def s) {
    if (!s) return []
    return s.toString().split(",").collect { it.toString().trim() }.findAll { it }
}

private Integer safeInt(def v, Integer fallback) {
    try {
        if (v == null) return fallback
        if (v instanceof Number) return (v as Number).intValue()
        return Integer.parseInt(v.toString())
    } catch (e) {
        return fallback
    }
}

private String urlEnc(String s) { URLEncoder.encode(s ?: "", "UTF-8") }

private String safeUrlDecode(String s) {
    try { URLDecoder.decode(s, "UTF-8") } catch (ignored) { s }
}

def toQueryString(Map m) {
    return m.collect { k, v -> "${k}=${URLEncoder.encode(v.toString(), 'UTF-8')}" }.sort().join("&")
}

def oAuthInstructions() {
    "Steps to enable OAuth:\n1) Go to Apps Code in Hubitat.\n2) Open this app's code entry.\n3) Enable the OAuth toggle.\n4) Click Save.\n"
}

def authenticationInstructions() {
    "Tap Authenticate to open the provider login page. Complete sign-in and consent. When successful, close the tab and return to Hubitat."
}

def connectionStatus(message) {
    def html = """
        <!DOCTYPE html>
        <html>
            <head>
                <meta name="viewport" content="width=device-width, initial-scale=1">
                <title>Authorization</title>
            </head>
            <body>
                <div style="font-family: Arial, sans-serif; padding: 16px;">
                    ${message}
                </div>
            </body>
        </html>
    """
    render contentType: "text/html", data: html
}

/* ---------------------------
 * Logging
 * --------------------------*/

private void logDebug(msg, type=null) {
    if (isDebugEnabled != true) return
    if (type == "auth" && debugAuth != true) return
    log.debug "${msg}"
}
private void logInfo(msg) { log.info "${msg}" }
private void logWarn(msg) { log.warn "${msg}" }
