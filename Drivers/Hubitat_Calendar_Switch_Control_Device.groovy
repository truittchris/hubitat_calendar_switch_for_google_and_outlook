diff --git a/Drivers/Hubitat_Calendar_Switch_Control_Device.groovy b/Drivers/Hubitat_Calendar_Switch_Control_Device.groovy
index 1111111..2222222 100644
--- a/Drivers/Hubitat_Calendar_Switch_Control_Device.groovy
+++ b/Drivers/Hubitat_Calendar_Switch_Control_Device.groovy
@@ -1,12 +1,64 @@
 /* * Hubitat Calendar Switch - Control Device * * Author: Chris Truitt * Website: https://christruitt.com * GitHub: https://github.com/truittchris * Namespace: truittchris * * Child driver for the Hubitat Calendar Switch app. * * Architecture * - The parent app handles OAuth, provider polling, and event fetch. * - This device owns all match logic (match words, ignore words, timing, and filters). * - The parent app delivers a normalized event list; this device evaluates and sets its own switch state.
 */ import groovy.transform.Field
+import java.time.LocalDate
+import java.time.LocalDateTime
+import java.time.ZoneId
+import java.time.format.DateTimeFormatter
+
 @Field static final String DRIVER_NAME = "Hubitat Calendar Switch - Control Device"
-@Field static final String DRIVER_VERSION = "1.0.5"
+@Field static final String DRIVER_VERSION = "1.0.6"
 @Field static final String DT_FMT = "yyyy-MM-dd'T'HH:mm:ssXXX"
 metadata { definition( name: DRIVER_NAME, namespace: "truittchris", author: "Chris Truitt", description: "Calendar driven control device managed by the Hubitat Calendar Switch app.", category: "Convenience" ) { capability "Switch" capability "Actuator" capability "Sensor" capability "Refresh" // User-facing commands (shown on the Commands page).
 command "applyRulesNow" command "fetchNow" command "fetchNowAndApply" command "testNow" // Informational attributes.
 attribute "provider", "string" attribute "activeEventTitle", "string" attribute "activeEventStart", "string" attribute "activeEventEnd", "string" attribute "nextEventTitle", "string" attribute "nextEventStart", "string" attribute "nextEventEnd", "string" attribute "upcomingSummary", "string" attribute "lastPoll", "string" attribute "lastFetch", "string" attribute "lastError", "string" attribute "lastTestTime", "string" attribute "lastTestResult", "string" attribute "driverVersion", "string" attribute "commandHelp", "string" } }
-preferences { section("Overview") { paragraph("This switch turns On when the calendar matches your rules.
-It turns Off when no matching event is found.\n\nSet Match words and Ignore words below, then use Test now to confirm.") } section("Match rules") { input name: "includeKeywords", type: "string", title: "Match words", required: false, description: "Optional. If set, an event must contain at least one of these words. Separate with commas or new lines." input name: "excludeKeywords", type: "string", title: "Ignore words", required: false, description: "Optional.
-If an event contains any ignore word, it will be ignored.
-Separate with commas or new lines." input name: "clearMatchWords", type: "button", title: "Clear match words" input name: "clearIgnoreWords", type: "button", title: "Clear ignore words" input name: "clearAllWords", type: "button", title: "Clear both" } section("Timing") { input name: "minutesBeforeStart", type: "number", title: "Minutes before start", defaultValue: 0, required: true, description: "Turns On this many minutes before the event starts." input name: "minutesAfterEnd", type: "number", title: "Minutes after end", defaultValue: 0, required: true, description: "Stays On this many minutes after the event ends." } section("Event filters") { input name: "onlyBusyEvents", type: "bool", title: "Only consider busy events", defaultValue: true, required: true input name: "allowAllDayEvents", type: "bool", title: "Allow all day events", defaultValue: false, required: true input name: "allowPrivateEvents", type: "bool", title: "Allow private events", defaultValue: true, required: true } section("Test") { paragraph("Commands are also available on the Commands tab.\n\nApply rules now: evaluates using the last fetched events.\nFetch now: asks the app to pull new events.\nFetch now and apply: fetches and evaluates.\nTest now: fetches and evaluates and updates Last Test Result.") input name: "applyRulesNow", type: "button", title: "Apply rules now (use last fetched events)" input name: "fetchNowAndApply", type: "button", title: "Fetch now and apply rules" input name: "testNow", type: "button", title: "Test now (fetch and apply)" } section("Advanced") { input name: "logLevel", type: "enum", title: "Logging", required: true, defaultValue: "off", options: [ "off": "Off (recommended)", "basic": "Basic", "debug": "Debug" ] } section("Support") { paragraph("Driver: ${DRIVER_NAME}\nVersion: ${DRIVER_VERSION}\nAuthor: Chris Truitt\nWebsite: https://christruitt.com\nGitHub: https://github.com/truittchris\nSupport development: https://christruitt.com/tip-jar") } }
+// NOTE: Driver preferences do not support app-style UI helpers like section()/paragraph().
+// Keep this as input-only so HPM installs cleanly.
+preferences {
+    input name: "includeKeywords", type: "string", title: "Match words", required: false,
+          description: "Optional. Event must contain at least one. Separate with commas or new lines."
+    input name: "excludeKeywords", type: "string", title: "Ignore words", required: false,
+          description: "Optional. If event contains any ignore word, it will be ignored. Separate with commas or new lines."
+
+    input name: "clearMatchWords", type: "button", title: "Clear match words"
+    input name: "clearIgnoreWords", type: "button", title: "Clear ignore words"
+    input name: "clearAllWords", type: "button", title: "Clear both"
+
+    input name: "minutesBeforeStart", type: "number", title: "Minutes before start", defaultValue: 0, required: true,
+          description: "Turns On this many minutes before the event starts."
+    input name: "minutesAfterEnd", type: "number", title: "Minutes after end", defaultValue: 0, required: true,
+          description: "Stays On this many minutes after the event ends."
+
+    input name: "onlyBusyEvents", type: "bool", title: "Only consider busy events", defaultValue: true, required: true
+    input name: "allowAllDayEvents", type: "bool", title: "Allow all day events", defaultValue: false, required: true
+    input name: "allowPrivateEvents", type: "bool", title: "Allow private events", defaultValue: true, required: true
+
+    input name: "applyRulesNow", type: "button", title: "Apply rules now (use last fetched events)"
+    input name: "fetchNowAndApply", type: "button", title: "Fetch now and apply rules"
+    input name: "testNow", type: "button", title: "Test now (fetch and apply)"
+
+    input name: "logLevel", type: "enum", title: "Logging", required: true, defaultValue: "off",
+          options: [ "off": "Off (recommended)", "basic": "Basic", "debug": "Debug" ]
+}
 // ----------------------------------------------------------------------------- // Lifecycle // -----------------------------------------------------------------------------
 void installed() { initialize() } void updated() { initialize() // After saving preferences, request a fetch so changes take effect quickly.
 requestFetchNow("preferences") // Apply changes immediately using the most recently fetched events (if available).
 if (state?.lastEvents instanceof List && !((List)state.lastEvents).isEmpty()) { reEvaluateFromCache() } } private void initialize() { if (device.currentValue("switch") == null) sendEvent(name: "switch", value: "off") sendEvent(name: "driverVersion", value: DRIVER_VERSION) if (device.currentValue("lastError") == null) sendEvent(name: "lastError", value: "") setCommandHelp() } private void setCommandHelp() { String s = "Apply rules now - evaluates using the last fetched events.
@@ -12,7 +64,36 @@
 void evaluateEvents(Map providerMeta, List events) { setCommandHelp() String providerName = (providerMeta?.provider ?: "").toString() String error = (providerMeta?.error ?: "").toString() List normalizedEvents = (events ?: []).collect { (it instanceof Map) ? (Map)it : [:] } state.lastProviderMeta = (providerMeta instanceof Map) ? providerMeta : [:] state.lastEvents = normalizedEvents if (providerName) sendEvent(name: "provider", value: providerName) sendEvent(name: "lastPoll", value: _fmtNow()) if (providerMeta?.fetchedAt) sendEvent(name: "lastFetch", value: (providerMeta.fetchedAt ?: "").toString()) if (error) { sendEvent(name: "lastError", value: error) _warn("Provider error received: ${error}") _updateTestResultIfPending(false, "Provider error: ${error}") return } sendEvent(name: "lastError", value: "") Date nowDate = new Date()
- List candidates = normalizedEvents .findAll { Map e -> _passesEligibilityFilters(e) } .findAll { Map e -> _passesKeywordFilters(e) } .collect { Map e -> _withParsedTimes(e) } .findAll { Map e -> e._startDate != null && e._endDate != null } .sort { Map a, Map b -> a._startDate <=> b._startDate }
+ // Break into stages so debug logging can pinpoint where events are being dropped.
+ List stage1 = normalizedEvents.findAll { Map e -> _passesEligibilityFilters(e) }
+ List stage2 = stage1.findAll { Map e -> _passesKeywordFilters(e) }
+ List stage3 = stage2.collect { Map e -> _withParsedTimes(e) }
+ List candidates = stage3
+     .findAll { Map e -> e._startDate != null && e._endDate != null }
+     .sort { Map a, Map b -> a._startDate <=> b._startDate }
+
+ if (isDebugEnabled()) {
+     _debug("providerMeta=${providerMeta ?: [:]}")
+     _debug("events total=${normalizedEvents.size()} eligibility=${stage1.size()} keyword=${stage2.size()} parsed=${stage3.size()} candidates=${candidates.size()}")
+     if (!normalizedEvents.isEmpty()) _debug("sampleEvent(0)=${normalizedEvents[0]}")
+     int parsedOk = stage3.count { Map e -> e._startDate != null && e._endDate != null }
+     _debug("parsed dates OK=${parsedOk}/${stage3.size()}")
+ }
  Map active = _findActiveEvent(candidates, nowDate) Map next = _findNextEvent(candidates, nowDate) boolean shouldBeOn = (active != null) _updateActiveAttributes(active) _updateNextAttributes(next) _updateUpcomingSummary(candidates, nowDate) String current = (device.currentValue("switch") ?: "off").toString() String target = shouldBeOn ? "on" : "off" if (current != target) sendEvent(name: "switch", value: target) _updateTestResultIfPending(true, "OK - evaluated ${normalizedEvents.size()} events") if (isDebugEnabled()) { _debug("Evaluated ${normalizedEvents.size()} events, candidates=${candidates.size()}, active=${active?.title ?: "none"}, next=${next?.title ?: "none"}, switch=${target}") } }
@@ -12,9 +93,55 @@
 // ----------------------------------------------------------------------------- // Parsing and formatting // -----------------------------------------------------------------------------
-private Date _parseEventDate(def v) { if (v == null) return null if (v instanceof Date) return (Date)v try { return Date.parse(DT_FMT, v.toString()) } catch (Exception ignored) { try { return Date.parse("yyyy-MM-dd'T'HH:mm:ss.SSSX", v.toString()) } catch (Exception ignored2) { return null } } }
+private Date _parseEventDate(def v) {
+    if (v == null) return null
+    if (v instanceof Date) return (Date)v
+
+    // Handle Graph-style objects: [dateTime: "...", timeZone: "..."]
+    if (v instanceof Map) {
+        def dt = v.dateTime ?: v.datetime ?: v.start ?: null
+        def tz = v.timeZone ?: v.timezone ?: null
+        if (dt != null) return _parseGraphDateTime(dt.toString(), tz?.toString())
+        return null
+    }
+
+    String s = v.toString().trim()
+    if (!s) return null
+
+    // Date-only (common for all-day)
+    if (s ==~ /^\d{4}-\d{2}-\d{2}$/) {
+        try {
+            LocalDate ld = LocalDate.parse(s)
+            ZoneId zid = (location?.timeZone?.ID) ? ZoneId.of(location.timeZone.ID) : ZoneId.of("UTC")
+            return Date.from(ld.atStartOfDay(zid).toInstant())
+        } catch (Exception ignored) {
+            return null
+        }
+    }
+
+    // ISO with offset or Z, or other known patterns
+    try { return Date.parse(DT_FMT, s) } catch (Exception ignored) {}
+    try { return Date.parse("yyyy-MM-dd'T'HH:mm:ss.SSSX", s) } catch (Exception ignored) {}
+
+    // ISO local date-time (no offset)
+    try {
+        LocalDateTime ldt = LocalDateTime.parse(s, DateTimeFormatter.ISO_LOCAL_DATE_TIME)
+        ZoneId zid = (location?.timeZone?.ID) ? ZoneId.of(location.timeZone.ID) : ZoneId.of("UTC")
+        return Date.from(ldt.atZone(zid).toInstant())
+    } catch (Exception ignored) {}
+
+    return null
+}
+
+private Date _parseGraphDateTime(String dt, String tzId) {
+    if (!dt) return null
+
+    // If it already contains an offset or Z, parse as-is
+    if (dt.endsWith("Z") || (dt ==~ /.*[+-]\d\d:\d\d$/)) {
+        return _parseEventDate(dt)
+    }
+
+    // Otherwise interpret as a local date-time in the provided tz (fallback to hub tz)
+    try {
+        ZoneId zid = (tzId && tzId.trim())
+            ? ZoneId.of(tzId.trim())
+            : ((location?.timeZone?.ID) ? ZoneId.of(location.timeZone.ID) : ZoneId.of("UTC"))
+
+        LocalDateTime ldt = LocalDateTime.parse(dt, DateTimeFormatter.ISO_LOCAL_DATE_TIME)
+        return Date.from(ldt.atZone(zid).toInstant())
+    } catch (Exception e) {
+        if (isDebugEnabled()) _debug("Failed to parse Graph datetime dt='${dt}' tz='${tzId}': ${e.message}")
+        return null
+    }
+}
+
 private String _fmtNow() { return _fmtDate(new Date()) } private String _fmtDate(Date d) { if (d == null) return "" def tz = location?.timeZone ?: TimeZone.getTimeZone("UTC") return d.format(DT_FMT, tz) } private boolean _asBool(def v) { if (v == null) return false if (v instanceof Boolean) return (Boolean)v return v.toString().toLowerCase() in ["true", "1", "yes", "y", "busy"] } private Integer _asInt(def v, Integer defaultVal) { try { if (v == null) return defaultVal if (v instanceof Number) return ((Number)v).intValue() String s = v.toString().trim() return s ? Integer.parseInt(s) : defaultVal } catch (Exception ignored) { return defaultVal } }
