/*
 *  Hubitat Calendar Switch - Control Device
 *
 *  Author: Chris Truitt
 *  Website: https://christruitt.com
 *  GitHub:  https://github.com/truittchris
 *  Namespace: truittchris
 *
 *  Child driver for the Hubitat Calendar Switch app.
 *
 *  Architecture
 *  - The parent app handles OAuth, provider polling, and event fetch.
 *  - This device owns all match logic (match words, ignore words, timing, and filters).
 *  - The parent app delivers a normalized event list; this device evaluates and sets its own switch state.
 */

import groovy.transform.Field

@Field static final String DRIVER_NAME = "Hubitat Calendar Switch - Control Device"
@Field static final String DRIVER_VERSION = "1.0.5"
@Field static final String DT_FMT = "yyyy-MM-dd'T'HH:mm:ssXXX"

metadata {
    definition(
        name: DRIVER_NAME,
        namespace: "truittchris",
        author: "Chris Truitt",
        description: "Calendar driven control device managed by the Hubitat Calendar Switch app.",
        category: "Convenience"
    ) {
        capability "Switch"
        capability "Actuator"
        capability "Sensor"
        capability "Refresh"

        // User-facing commands (shown on the Commands page).
        command "applyRulesNow"
        command "fetchNow"
        command "fetchNowAndApply"
        command "testNow"

        // Informational attributes.
        attribute "provider", "string"
        attribute "activeEventTitle", "string"
        attribute "activeEventStart", "string"
        attribute "activeEventEnd", "string"
        attribute "nextEventTitle", "string"
        attribute "nextEventStart", "string"
        attribute "nextEventEnd", "string"
        attribute "upcomingSummary", "string"
        attribute "lastPoll", "string"
        attribute "lastFetch", "string"
        attribute "lastError", "string"
        attribute "lastTestTime", "string"
        attribute "lastTestResult", "string"
        attribute "driverVersion", "string"
        attribute "commandHelp", "string"
    }
}

preferences {
    section("Overview") {
        paragraph("This switch turns On when the calendar matches your rules. It turns Off when no matching event is found.\n\nSet Match words and Ignore words below, then use Test now to confirm.")
    }

    section("Match rules") {
        input name: "includeKeywords", type: "string", title: "Match words", required: false,
            description: "Optional. If set, an event must contain at least one of these words. Separate with commas or new lines."
        input name: "excludeKeywords", type: "string", title: "Ignore words", required: false,
            description: "Optional. If an event contains any ignore word, it will be ignored. Separate with commas or new lines."

        input name: "clearMatchWords", type: "button", title: "Clear match words"
        input name: "clearIgnoreWords", type: "button", title: "Clear ignore words"
        input name: "clearAllWords", type: "button", title: "Clear both"
    }

    section("Timing") {
        input name: "minutesBeforeStart", type: "number", title: "Minutes before start", defaultValue: 0, required: true,
            description: "Turns On this many minutes before the event starts."
        input name: "minutesAfterEnd", type: "number", title: "Minutes after end", defaultValue: 0, required: true,
            description: "Stays On this many minutes after the event ends."
    }

    section("Event filters") {
        input name: "onlyBusyEvents", type: "bool", title: "Only consider busy events", defaultValue: true, required: true
        input name: "allowAllDayEvents", type: "bool", title: "Allow all day events", defaultValue: false, required: true
        input name: "allowPrivateEvents", type: "bool", title: "Allow private events", defaultValue: true, required: true
    }

    section("Test") {
        paragraph("Commands are also available on the Commands tab.\n\nApply rules now: evaluates using the last fetched events.\nFetch now: asks the app to pull new events.\nFetch now and apply: fetches and evaluates.\nTest now: fetches and evaluates and updates Last Test Result.")
        input name: "applyRulesNow", type: "button", title: "Apply rules now (use last fetched events)"
        input name: "fetchNowAndApply", type: "button", title: "Fetch now and apply rules"
        input name: "testNow", type: "button", title: "Test now (fetch and apply)"
    }

    section("Advanced") {
        input name: "logLevel", type: "enum", title: "Logging", required: true, defaultValue: "off", options: [
            "off": "Off (recommended)",
            "basic": "Basic",
            "debug": "Debug"
        ]
    }

    section("Support") {
        paragraph("Driver: ${DRIVER_NAME}\nVersion: ${DRIVER_VERSION}\nAuthor: Chris Truitt\nWebsite: https://christruitt.com\nGitHub: https://github.com/truittchris\nSupport development: https://christruitt.com/tip-jar")
    }
}

// -----------------------------------------------------------------------------
// Lifecycle
// -----------------------------------------------------------------------------

void installed() {
    initialize()
}

void updated() {
    initialize()

    // After saving preferences, request a fetch so changes take effect quickly.
    requestFetchNow("preferences")

    // Apply changes immediately using the most recently fetched events (if available).
    if (state?.lastEvents instanceof List && !((List)state.lastEvents).isEmpty()) {
        reEvaluateFromCache()
    }
}

private void initialize() {
    if (device.currentValue("switch") == null) sendEvent(name: "switch", value: "off")
    sendEvent(name: "driverVersion", value: DRIVER_VERSION)
    if (device.currentValue("lastError") == null) sendEvent(name: "lastError", value: "")
    setCommandHelp()
}

private void setCommandHelp() {
    String s = "Apply rules now - evaluates using the last fetched events. " +
        "Fetch now - asks the app to pull new events. " +
        "Fetch now and apply - fetches and evaluates. " +
        "Test now - fetches and evaluates and updates Last Test Result. " +
        "Refresh - same as Fetch now and apply."
    sendEvent(name: "commandHelp", value: s)
}

// -----------------------------------------------------------------------------
// Switch capability
// -----------------------------------------------------------------------------

void on() { sendEvent(name: "switch", value: "on") }
void off() { sendEvent(name: "switch", value: "off") }

// -----------------------------------------------------------------------------
// Refresh capability
// -----------------------------------------------------------------------------

void refresh() {
    fetchNowAndApply()
}

// -----------------------------------------------------------------------------
// Preference buttons
// -----------------------------------------------------------------------------

void clearMatchWords() {
    device.updateSetting("includeKeywords", [value: "", type: "string"])
    _basic("Cleared match words")
}

void clearIgnoreWords() {
    device.updateSetting("excludeKeywords", [value: "", type: "string"])
    _basic("Cleared ignore words")
}

void clearAllWords() {
    device.updateSetting("includeKeywords", [value: "", type: "string"])
    device.updateSetting("excludeKeywords", [value: "", type: "string"])
    _basic("Cleared match words and ignore words")
}

// -----------------------------------------------------------------------------
// User-facing commands (Commands page)
// -----------------------------------------------------------------------------

void applyRulesNow() {
    reEvaluateFromCache()
}

void fetchNow() {
    requestFetchNow("fetch")
}

void fetchNowAndApply() {
    requestFetchNow("fetchApply")
    runIn(1, "reEvaluateFromCache")
}

void testNow() {
    state.testRequestedMs = now()
    sendEvent(name: "lastTestTime", value: _fmtNow())
    sendEvent(name: "lastTestResult", value: "Running")
    requestFetchNow("test")
}

// -----------------------------------------------------------------------------
// Parent coordination
// -----------------------------------------------------------------------------

private void requestFetchNow(String reason) {
    long lastReq = (state?.lastFetchRequestMs instanceof Long) ? (Long)state.lastFetchRequestMs : 0L
    if (lastReq && (now() - lastReq) < 3000L) return
    state.lastFetchRequestMs = now()

    if (!parent) {
        String msg = "Not linked to the app. Delete this device and re-add it from the Hubitat Calendar Switch app."
        sendEvent(name: "lastError", value: msg)
        _warn(msg)
        return
    }

    try {
        parent.childRequestFetch(device.deviceNetworkId, reason ?: "")
    } catch (Exception e) {
        sendEvent(name: "lastError", value: "Fetch failed: ${e.message}")
        _warn("Fetch request failed: ${e.message}")
    }
}

private void reEvaluateFromCache() {
    Map meta = (state.lastProviderMeta instanceof Map) ? (Map)state.lastProviderMeta : [:]
    List cached = (state.lastEvents instanceof List) ? (List)state.lastEvents : []

    if (!cached || cached.isEmpty()) {
        sendEvent(name: "lastError", value: "No events yet. Use Fetch now and apply, or wait for the next scheduled poll.")
        return
    }

    evaluateEvents(meta, cached)
}

// -----------------------------------------------------------------------------
// Entry point invoked by parent app
// -----------------------------------------------------------------------------

void evaluateEvents(Map providerMeta, List events) {
    setCommandHelp()

    String providerName = (providerMeta?.provider ?: "").toString()
    String error = (providerMeta?.error ?: "").toString()

    List<Map> normalizedEvents = (events ?: []).collect { (it instanceof Map) ? (Map)it : [:] }
    state.lastProviderMeta = (providerMeta instanceof Map) ? providerMeta : [:]
    state.lastEvents = normalizedEvents

    if (providerName) sendEvent(name: "provider", value: providerName)

    sendEvent(name: "lastPoll", value: _fmtNow())
    if (providerMeta?.fetchedAt) sendEvent(name: "lastFetch", value: (providerMeta.fetchedAt ?: "").toString())

    if (error) {
        sendEvent(name: "lastError", value: error)
        _warn("Provider error received: ${error}")
        _updateTestResultIfPending(false, "Provider error: ${error}")
        return
    }

    sendEvent(name: "lastError", value: "")

    Date nowDate = new Date()

    List<Map> candidates = normalizedEvents
        .findAll { Map e -> _passesEligibilityFilters(e) }
        .findAll { Map e -> _passesKeywordFilters(e) }
        .collect { Map e -> _withParsedTimes(e) }
        .findAll { Map e -> e._startDate != null && e._endDate != null }
        .sort { Map a, Map b -> a._startDate <=> b._startDate }

    Map active = _findActiveEvent(candidates, nowDate)
    Map next = _findNextEvent(candidates, nowDate)

    boolean shouldBeOn = (active != null)

    _updateActiveAttributes(active)
    _updateNextAttributes(next)
    _updateUpcomingSummary(candidates, nowDate)

    String current = (device.currentValue("switch") ?: "off").toString()
    String target = shouldBeOn ? "on" : "off"
    if (current != target) sendEvent(name: "switch", value: target)

    _updateTestResultIfPending(true, "OK - evaluated ${normalizedEvents.size()} events")

    if (isDebugEnabled()) {
        _debug("Evaluated ${normalizedEvents.size()} events, candidates=${candidates.size()}, active=${active?.title ?: "none"}, next=${next?.title ?: "none"}, switch=${target}")
    }
}

// -----------------------------------------------------------------------------
// Filtering helpers
// -----------------------------------------------------------------------------

private boolean _passesEligibilityFilters(Map e) {
    boolean isAllDay = _asBool(e?.isAllDay)
    boolean isBusy = _asBool(e?.isBusy)
    boolean isPrivate = _asBool(e?.isPrivate)

    if (!allowAllDayEvents && isAllDay) return false
    if (onlyBusyEvents && !isBusy) return false
    if (!allowPrivateEvents && isPrivate) return false

    return true
}

private boolean _passesKeywordFilters(Map e) {
    List<String> includes = _normalizeKeywords(includeKeywords)
    List<String> excludes = _normalizeKeywords(excludeKeywords)

    String haystack = _eventHaystack(e)
    boolean includeOk = (includes.isEmpty() || includes.any { kw -> haystack.contains(kw) })
    boolean excludeHit = (!excludes.isEmpty() && excludes.any { kw -> haystack.contains(kw) })

    return includeOk && !excludeHit
}

private Map _withParsedTimes(Map e) {
    Map copy = new LinkedHashMap(e)
    copy._startDate = _parseEventDate(e?.start)
    copy._endDate = _parseEventDate(e?.end)
    return copy
}

private Map _findActiveEvent(List<Map> events, Date nowDate) {
    Integer beforeMins = _asInt(minutesBeforeStart, 0)
    Integer afterMins = _asInt(minutesAfterEnd, 0)

    for (Map e : events) {
        Date start = (Date)e._startDate
        Date end = (Date)e._endDate
        if (start == null || end == null) continue

        Date windowStart = new Date(start.time - (beforeMins * 60L * 1000L))
        Date windowEnd = new Date(end.time + (afterMins * 60L * 1000L))

        if (nowDate >= windowStart && nowDate <= windowEnd) return e
    }
    return null
}

private Map _findNextEvent(List<Map> events, Date nowDate) {
    for (Map e : events) {
        Date start = (Date)e._startDate
        if (start != null && start.after(nowDate)) return e
    }
    return null
}

private void _updateActiveAttributes(Map e) {
    if (e == null) {
        sendEvent(name: "activeEventTitle", value: "")
        sendEvent(name: "activeEventStart", value: "")
        sendEvent(name: "activeEventEnd", value: "")
        return
    }
    sendEvent(name: "activeEventTitle", value: (e.title ?: "").toString())
    sendEvent(name: "activeEventStart", value: _fmtDate((Date)e._startDate))
    sendEvent(name: "activeEventEnd", value: _fmtDate((Date)e._endDate))
}

private void _updateNextAttributes(Map e) {
    if (e == null) {
        sendEvent(name: "nextEventTitle", value: "")
        sendEvent(name: "nextEventStart", value: "")
        sendEvent(name: "nextEventEnd", value: "")
        return
    }
    sendEvent(name: "nextEventTitle", value: (e.title ?: "").toString())
    sendEvent(name: "nextEventStart", value: _fmtDate((Date)e._startDate))
    sendEvent(name: "nextEventEnd", value: _fmtDate((Date)e._endDate))
}

private void _updateUpcomingSummary(List<Map> events, Date nowDate) {
    List<Map> relevant = events.findAll { Map e ->
        Date end = (Date)e._endDate
        return (end != null && end.after(new Date(nowDate.time - 60L * 1000L)))
    }

    List<Map> top = relevant.take(3)
    if (top.isEmpty()) {
        sendEvent(name: "upcomingSummary", value: "")
        return
    }

    List<String> parts = top.collect { Map e ->
        String t = (e.title ?: "").toString()
        Date s = (Date)e._startDate
        String when = (s != null) ? _fmtDate(s) : ""
        return when ? "${when} - ${t}" : t
    }

    sendEvent(name: "upcomingSummary", value: parts.join(" | "))
}

// -----------------------------------------------------------------------------
// Keyword normalization
// -----------------------------------------------------------------------------

private List<String> _normalizeKeywords(String raw) {
    if (raw == null) return []
    String cleaned = raw.toString().trim()
    if (!cleaned) return []

    List<String> tokens = cleaned.split(/[\n\r,;|]+/)
        .collect { it.toString().trim().toLowerCase() }
        .findAll { it }

    List<String> out = []
    tokens.each { String t -> if (!out.contains(t)) out << t }
    return out
}

private String _eventHaystack(Map e) {
    List<String> fields = []
    fields << (e?.title ?: "").toString()
    fields << (e?.location ?: "").toString()
    fields << (e?.organizer ?: "").toString()
    fields << (e?.description ?: "").toString()
    fields << (e?.categories ?: "").toString()
    return fields.join(" ").toLowerCase()
}

// -----------------------------------------------------------------------------
// Parsing and formatting
// -----------------------------------------------------------------------------

private Date _parseEventDate(def v) {
    if (v == null) return null
    if (v instanceof Date) return (Date)v

    try {
        return Date.parse(DT_FMT, v.toString())
    } catch (Exception ignored) {
        try {
            return Date.parse("yyyy-MM-dd'T'HH:mm:ss.SSSX", v.toString())
        } catch (Exception ignored2) {
            return null
        }
    }
}

private String _fmtNow() { return _fmtDate(new Date()) }

private String _fmtDate(Date d) {
    if (d == null) return ""
    def tz = location?.timeZone ?: TimeZone.getTimeZone("UTC")
    return d.format(DT_FMT, tz)
}

private boolean _asBool(def v) {
    if (v == null) return false
    if (v instanceof Boolean) return (Boolean)v
    return v.toString().toLowerCase() in ["true", "1", "yes", "y", "busy"]
}

private Integer _asInt(def v, Integer defaultVal) {
    try {
        if (v == null) return defaultVal
        if (v instanceof Number) return ((Number)v).intValue()
        String s = v.toString().trim()
        return s ? Integer.parseInt(s) : defaultVal
    } catch (Exception ignored) {
        return defaultVal
    }
}

// -----------------------------------------------------------------------------
// Test result handling
// -----------------------------------------------------------------------------

private void _updateTestResultIfPending(boolean ok, String details) {
    Long req = (state?.testRequestedMs instanceof Long) ? (Long)state.testRequestedMs : null
    if (!req) return

    if ((now() - req) > 60000L) {
        state.testRequestedMs = null
        return
    }

    sendEvent(name: "lastTestResult", value: details ?: (ok ? "OK" : "Failed"))
    state.testRequestedMs = null
}

// -----------------------------------------------------------------------------
// Logging
// -----------------------------------------------------------------------------

private boolean isBasicEnabled() {
    String lvl = (settings?.logLevel ?: "off").toString()
    return (lvl == "basic" || lvl == "debug")
}

private boolean isDebugEnabled() {
    String lvl = (settings?.logLevel ?: "off").toString()
    return (lvl == "debug")
}

private void _debug(String msg) { if (isDebugEnabled()) log.debug "${device.displayName}: ${msg}" }
private void _basic(String msg) { if (isBasicEnabled()) log.info "${device.displayName}: ${msg}" }
private void _warn(String msg) { log.warn "${device.displayName}: ${msg}" }
