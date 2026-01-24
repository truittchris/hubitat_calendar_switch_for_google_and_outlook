/*
 * Hubitat Calendar Switch - Control Device
 *
 * Author: Chris Truitt
 * Website: https://christruitt.com
 * GitHub: https://github.com/truittchris
 * Namespace: truittchris
 *
 * Child driver for the Hubitat Calendar Switch app.
 *
 * Architecture
 * - The parent app handles OAuth, provider polling, and event fetch.
 * - This device owns all match logic (match words, ignore words, timing, and filters).
 * - The parent app delivers a normalized event list; this device evaluates and sets its own switch state.
 */

import groovy.transform.Field
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter

@Field static final String DRIVER_NAME = "Hubitat Calendar Switch - Control Device"
@Field static final String DRIVER_VERSION = "1.0.7"
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

        command "applyRulesNow"
        command "fetchNow"
        command "fetchNowAndApply"
        command "testNow"

        command "clearMatchWords"
        command "clearIgnoreWords"
        command "clearAllWords"

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

/*
 * NOTE: Driver preferences are input-only.
 * Do not use section()/paragraph() or other app-only UI helpers in drivers.
 */
preferences {
    input name: "includeKeywords", type: "text", title: "Match words", required: false,
        description: "Optional. Event must contain at least one. Separate with commas or new lines."

    input name: "excludeKeywords", type: "text", title: "Ignore words", required: false,
        description: "Optional. If event contains any ignore word, it will be ignored. Separate with commas or new lines."

    input name: "minutesBeforeStart", type: "number", title: "Minutes before start", defaultValue: 0, required: true,
        description: "Turns On this many minutes before the event starts."

    input name: "minutesAfterEnd", type: "number", title: "Minutes after end", defaultValue: 0, required: true,
        description: "Stays On this many minutes after the event ends."

    input name: "onlyBusyEvents", type: "bool", title: "Only consider busy events", defaultValue: true, required: true
    input name: "allowAllDayEvents", type: "bool", title: "Allow all day events", defaultValue: false, required: true
    input name: "allowPrivateEvents", type: "bool", title: "Allow private events", defaultValue: true, required: true

    input name: "logLevel", type: "enum", title: "Logging", required: true, defaultValue: "off",
        options: [ "off": "Off (recommended)", "basic": "Basic", "debug": "Debug" ]
}

// -----------------------------------------------------------------------------
// Lifecycle
// -----------------------------------------------------------------------------
void installed() {
    initialize()
}

void updated() {
    initialize()
    requestFetchNow("preferences")
    if (state?.lastEvents instanceof List && !((List) state.lastEvents).isEmpty()) {
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
    String s =
        "applyRulesNow – evaluates using the last fetched events.\n" +
        "fetchNow – asks the parent app to fetch new events.\n" +
        "fetchNowAndApply – fetches and evaluates.\n" +
        "testNow – fetches and evaluates, sets Last Test Result.\n\n" +
        "clearMatchWords – clears Match words.\n" +
        "clearIgnoreWords – clears Ignore words.\n" +
        "clearAllWords – clears both word lists."
    device.updateDataValue("commandHelp", s)
    sendEvent(name: "commandHelp", value: s)
}

void refresh() {
    fetchNowAndApply()
}

// -----------------------------------------------------------------------------
// User commands
// -----------------------------------------------------------------------------
void applyRulesNow() {
    reEvaluateFromCache()
}

void fetchNow() {
    requestFetchNow("fetchNow")
}

void fetchNowAndApply() {
    requestFetchNow("fetchNowAndApply")
}

void testNow() {
    state.pendingTest = true
    state.pendingTestAt = _fmtNow()
    requestFetchNow("testNow")
}

void clearMatchWords() {
    device.updateSetting("includeKeywords", [value: "", type: "text"])
    reEvaluateFromCache()
}

void clearIgnoreWords() {
    device.updateSetting("excludeKeywords", [value: "", type: "text"])
    reEvaluateFromCache()
}

void clearAllWords() {
    device.updateSetting("includeKeywords", [value: "", type: "text"])
    device.updateSetting("excludeKeywords", [value: "", type: "text"])
    reEvaluateFromCache()
}

// -----------------------------------------------------------------------------
// Parent interaction
// -----------------------------------------------------------------------------
private void requestFetchNow(String reason) {
    try {
        parent?.childRequestFetch(device.deviceNetworkId, reason ?: "")
    } catch (Exception e) {
        _warn("Unable to request fetch from parent: ${e.message}")
        sendEvent(name: "lastError", value: "Fetch request failed: ${e.message}")
        _updateTestResultIfPending(false, "Fetch request failed: ${e.message}")
    }
}

private void reEvaluateFromCache() {
    try {
        Map meta = (state?.lastProviderMeta instanceof Map) ? (Map) state.lastProviderMeta : [:]
        List ev = (state?.lastEvents instanceof List) ? (List) state.lastEvents : []
        evaluateEvents(meta, ev)
    } catch (Exception e) {
        _warn("Re-evaluate failed: ${e.message}")
        sendEvent(name: "lastError", value: "Re-evaluate failed: ${e.message}")
        _updateTestResultIfPending(false, "Re-evaluate failed: ${e.message}")
    }
}

// -----------------------------------------------------------------------------
// Called by parent app: evaluate events and set switch state + attributes
// -----------------------------------------------------------------------------
void evaluateEvents(Map providerMeta, List events) {
    setCommandHelp()

    String providerName = (providerMeta?.provider ?: "").toString()
    String error = (providerMeta?.error ?: "").toString()

    List normalizedEvents = (events ?: []).collect { (it instanceof Map) ? (Map) it : [:] }
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

    // Stage breakdown helps diagnose why nothing becomes a candidate
    List stage1 = normalizedEvents.findAll { Map e -> _passesEligibilityFilters(e) }
    List stage2 = stage1.findAll { Map e -> _passesKeywordFilters(e) }
    List stage3 = stage2.collect { Map e -> _withParsedTimes(e) }

    List candidates = stage3
        .findAll { Map e -> e._startDate != null && e._endDate != null }
        .sort { Map a, Map b -> a._startDate <=> b._startDate }

    if (isDebugEnabled()) {
        _debug("providerMeta=${providerMeta ?: [:]}")
        _debug("events total=${normalizedEvents.size()} eligibility=${stage1.size()} keyword=${stage2.size()} parsed=${stage3.size()} candidates=${candidates.size()}")
        if (!normalizedEvents.isEmpty()) _debug("sampleEvent(0)=${normalizedEvents[0]}")
        int parsedOk = stage3.count { Map e -> e._startDate != null && e._endDate != null }
        _debug("parsed dates OK=${parsedOk}/${stage3.size()}")
    }

    Map active = _findActiveEvent(candidates, nowDate)
    Map next = _findNextEvent(candidates, nowDate)

    boolean shouldBeOn = (active != null)

    _updateActiveAttributes(active)
    _updateNextAttributes(next)
    _updateUpcomingSummary(candidates, nowDate)

    String current = (device.currentValue("switch") ?: "off").toString()
    String target = shouldBeOn ? "on" : "off"
    if (current != target) sendEvent(name: "switch", value: target)

    _updateTestResultIfPending(true, "OK – evaluated ${normalizedEvents.size()} events")

    if (isDebugEnabled()) {
        _debug("active=${active?.title ?: 'none'} next=${next?.title ?: 'none'} switch=${target}")
    }
}

// -----------------------------------------------------------------------------
// Filtering
// -----------------------------------------------------------------------------
private boolean _passesEligibilityFilters(Map e) {
    boolean allowAllDay = (settings?.allowAllDayEvents == null) ? false : (Boolean) settings.allowAllDayEvents
    boolean allowPrivate = (settings?.allowPrivateEvents == null) ? true : (Boolean) settings.allowPrivateEvents
    boolean onlyBusy = (settings?.onlyBusyEvents == null) ? true : (Boolean) settings.onlyBusyEvents

    // All day detection (either explicit flag, or date-only fields)
    boolean isAllDay = _asBool(e?.isAllDay) ||
        ((e?.start instanceof Map) && ((Map) e.start).date && !((Map) e.start).dateTime) ||
        ((e?.end instanceof Map) && ((Map) e.end).date && !((Map) e.end).dateTime)

    if (isAllDay && !allowAllDay) return false

    // Private detection
    String sensitivity = (e?.sensitivity ?: e?.visibility ?: "").toString().toLowerCase()
    boolean isPrivate = (sensitivity == "private")
    if (isPrivate && !allowPrivate) return false

    // Busy detection (Graph commonly uses showAs; Google often uses transparency)
    String showAs = (e?.showAs ?: e?.showas ?: e?.busyStatus ?: "").toString().toLowerCase()
    String transparency = (e?.transparency ?: "").toString().toLowerCase()

    boolean busy =
        (showAs in ["busy", "oof", "workingelsewhere"]) ||
        (transparency == "opaque") ||
        _asBool(e?.busy)

    if (onlyBusy && !busy) return false

    // Declined / cancelled filtering if provided
    String response = (e?.responseStatus ?: e?.response ?: e?.attendeeStatus ?: "").toString().toLowerCase()
    if (response == "declined") return false

    String status = (e?.status ?: "").toString().toLowerCase()
    if (status == "cancelled" || status == "canceled") return false

    return true
}

private boolean _passesKeywordFilters(Map e) {
    List<String> includes = _parseWords(settings?.includeKeywords)
    List<String> excludes = _parseWords(settings?.excludeKeywords)

    String haystack = _eventSearchText(e)

    if (!excludes.isEmpty()) {
        for (String w : excludes) {
            if (w && haystack.contains(w)) return false
        }
    }

    if (includes.isEmpty()) return true

    for (String w : includes) {
        if (w && haystack.contains(w)) return true
    }
    return false
}

private List<String> _parseWords(def raw) {
    if (raw == null) return []
    String s = raw.toString()
    if (!s.trim()) return []
    return s
        .split(/[\n,]+/)
        .collect { it.toString().trim().toLowerCase() }
        .findAll { it }
        .unique()
}

private String _eventSearchText(Map e) {
    List parts = []
    parts << (e?.title ?: e?.subject ?: "")
    parts << (e?.location ?: "")
    parts << (e?.organizer ?: "")
    parts << (e?.bodyPreview ?: e?.body ?: "")
    return parts.collect { it.toString().toLowerCase() }.join(" | ")
}

// -----------------------------------------------------------------------------
// Time parsing and windowing
// -----------------------------------------------------------------------------
private Map _withParsedTimes(Map e) {
    Map out = new LinkedHashMap(e ?: [:])

    Date start = _parseEventDate(e?.start)
    Date end = _parseEventDate(e?.end)

    // Some providers use startDateTime / endDateTime
    if (start == null) start = _parseEventDate(e?.startDateTime)
    if (end == null) end = _parseEventDate(e?.endDateTime)

    Integer beforeMin = _asInt(settings?.minutesBeforeStart, 0)
    Integer afterMin = _asInt(settings?.minutesAfterEnd, 0)

    if (start != null && beforeMin != 0) start = new Date(start.time - (beforeMin * 60L * 1000L))
    if (end != null && afterMin != 0) end = new Date(end.time + (afterMin * 60L * 1000L))

    out._startDate = start
    out._endDate = end
    return out
}

private Date _parseEventDate(def v) {
    if (v == null) return null
    if (v instanceof Date) return (Date) v

    // Graph-style: [dateTime: "...", timeZone: "..."] or all-day [date: "yyyy-mm-dd"]
    if (v instanceof Map) {
        Map m = (Map) v
        def dt = m.dateTime ?: m.datetime
        def tz = m.timeZone ?: m.timezone
        def d = m.date

        if (d != null && dt == null) {
            return _parseDateOnly(d.toString())
        }
        if (dt != null) {
            return _parseGraphDateTime(dt.toString(), tz?.toString())
        }
        return null
    }

    String s = v.toString().trim()
    if (!s) return null

    // Date-only
    if (s ==~ /^\d{4}-\d{2}-\d{2}$/) {
        return _parseDateOnly(s)
    }

    // ISO with offset or Z
    try { return Date.parse(DT_FMT, s) } catch (Exception ignored) {}
    try { return Date.parse("yyyy-MM-dd'T'HH:mm:ss.SSSX", s) } catch (Exception ignored) {}
    try { return Date.parse("yyyy-MM-dd'T'HH:mm:ssX", s) } catch (Exception ignored) {}

    // ISO local date-time (no offset)
    try {
        LocalDateTime ldt = LocalDateTime.parse(s, DateTimeFormatter.ISO_LOCAL_DATE_TIME)
        ZoneId zid = _safeZoneId(null)
        return Date.from(ldt.atZone(zid).toInstant())
    } catch (Exception ignored) {}

    return null
}

private Date _parseDateOnly(String s) {
    try {
        LocalDate ld = LocalDate.parse(s)
        ZoneId zid = _safeZoneId(null)
        return Date.from(ld.atStartOfDay(zid).toInstant())
    } catch (Exception ignored) {
        return null
    }
}

private Date _parseGraphDateTime(String dt, String tzId) {
    if (!dt) return null

    // If it already contains offset or Z, parse as-is
    if (dt.endsWith("Z") || (dt ==~ /.*[+-]\d\d:\d\d$/) || (dt ==~ /.*[+-]\d{4}$/)) {
        return _parseEventDate(dt)
    }

    // Otherwise interpret as local date-time in provided tz (fallback to hub tz)
    try {
        ZoneId zid = _safeZoneId(tzId)
        LocalDateTime ldt = LocalDateTime.parse(dt, DateTimeFormatter.ISO_LOCAL_DATE_TIME)
        return Date.from(ldt.atZone(zid).toInstant())
    } catch (Exception e) {
        if (isDebugEnabled()) _debug("Failed to parse Graph datetime dt='${dt}' tz='${tzId}': ${e.message}")
        return null
    }
}

private ZoneId _safeZoneId(String tzId) {
    if (tzId != null && tzId.toString().trim()) {
        try {
            return ZoneId.of(tzId.toString().trim())
        } catch (Exception ignored) {
            // Graph may send Windows TZ names; fall back to hub tz
        }
    }
    try {
        if (location?.timeZone?.ID) return ZoneId.of(location.timeZone.ID)
    } catch (Exception ignored2) {}
    return ZoneId.of("UTC")
}

// -----------------------------------------------------------------------------
// Active / next selection
// -----------------------------------------------------------------------------
private Map _findActiveEvent(List candidates, Date nowDate) {
    long nowMs = nowDate.time
    return candidates.find { Map e ->
        Date s = (Date) e._startDate
        Date en = (Date) e._endDate
        (s != null && en != null && s.time <= nowMs && nowMs < en.time)
    }
}

private Map _findNextEvent(List candidates, Date nowDate) {
    long nowMs = nowDate.time
    return candidates.find { Map e ->
        Date s = (Date) e._startDate
        (s != null && s.time > nowMs)
    }
}

// -----------------------------------------------------------------------------
// Attribute updates
// -----------------------------------------------------------------------------
private void _updateActiveAttributes(Map active) {
    if (active == null) {
        sendEvent(name: "activeEventTitle", value: "")
        sendEvent(name: "activeEventStart", value: "")
        sendEvent(name: "activeEventEnd", value: "")
        return
    }
    sendEvent(name: "activeEventTitle", value: _safeStr(active?.title ?: active?.subject ?: ""))
    sendEvent(name: "activeEventStart", value: _fmtDate((Date) active._startDate))
    sendEvent(name: "activeEventEnd", value: _fmtDate((Date) active._endDate))
}

private void _updateNextAttributes(Map next) {
    if (next == null) {
        sendEvent(name: "nextEventTitle", value: "")
        sendEvent(name: "nextEventStart", value: "")
        sendEvent(name: "nextEventEnd", value: "")
        return
    }
    sendEvent(name: "nextEventTitle", value: _safeStr(next?.title ?: next?.subject ?: ""))
    sendEvent(name: "nextEventStart", value: _fmtDate((Date) next._startDate))
    sendEvent(name: "nextEventEnd", value: _fmtDate((Date) next._endDate))
}

private void _updateUpcomingSummary(List candidates, Date nowDate) {
    // Keep it short: next 3 upcoming titles
    long nowMs = nowDate.time
    List<Map> upcoming = candidates.findAll { Map e ->
        Date s = (Date) e._startDate
        (s != null && s.time >= nowMs)
    }.take(3)

    if (upcoming.isEmpty()) {
        sendEvent(name: "upcomingSummary", value: "")
        return
    }

    String summary = upcoming.collect { Map e ->
        String t = _safeStr(e?.title ?: e?.subject ?: "")
        String whenStr = _fmtDate((Date) e._startDate)
        "${whenStr} – ${t}"
    }.join(" | ")

    sendEvent(name: "upcomingSummary", value: summary)
}

// -----------------------------------------------------------------------------
// Test result helpers
// -----------------------------------------------------------------------------
private void _updateTestResultIfPending(boolean ok, String msg) {
    if (state?.pendingTest) {
        state.pendingTest = false
        sendEvent(name: "lastTestTime", value: (state?.pendingTestAt ?: _fmtNow()).toString())
        sendEvent(name: "lastTestResult", value: (ok ? "PASS: " : "FAIL: ") + (msg ?: ""))
        state.pendingTestAt = null
    }
}

// -----------------------------------------------------------------------------
// Logging + formatting + coercion
// -----------------------------------------------------------------------------
private boolean isDebugEnabled() {
    return ((settings?.logLevel ?: "off").toString() == "debug")
}

private boolean isBasicEnabled() {
    String lvl = (settings?.logLevel ?: "off").toString()
    return (lvl == "basic" || lvl == "debug")
}

private void _debug(String m) { if (isDebugEnabled()) log.debug "${device.displayName}: ${m}" }
private void _info(String m)  { if (isBasicEnabled()) log.info "${device.displayName}: ${m}" }
private void _warn(String m)  { log.warn "${device.displayName}: ${m}" }

private String _fmtNow() { return _fmtDate(new Date()) }

private String _fmtDate(Date d) {
    if (d == null) return ""
    def tz = location?.timeZone ?: TimeZone.getTimeZone("UTC")
    return d.format(DT_FMT, tz)
}

private boolean _asBool(def v) {
    if (v == null) return false
    if (v instanceof Boolean) return (Boolean) v
    return v.toString().toLowerCase() in ["true", "1", "yes", "y", "busy"]
}

private Integer _asInt(def v, Integer defaultVal) {
    try {
        if (v == null) return defaultVal
        if (v instanceof Number) return ((Number) v).intValue()
        String s = v.toString().trim()
        return s ? Integer.parseInt(s) : defaultVal
    } catch (Exception ignored) {
        return defaultVal
    }
}

private String _safeStr(def v) {
    return (v == null) ? "" : v.toString()
}
