/**
 * truittchris Calendar Control Switch
 * Child driver for Calendar OAuth Bridge
 *
 * v0.1.3
 *
 * The parent app calls setCalendarState(onOff, meta) to:
 * - turn the switch on/off
 * - publish informative attributes for dashboards/logs
 * - publish upcoming matching events (next 3) for visibility on the device page
 */

metadata {
    definition(
        name: "Calendar Control Switch",
        namespace: "truittchris",
        author: "Chris Truitt"
    ) {
        capability "Switch"
        capability "Refresh"
        capability "Sensor"
        capability "Actuator"

        attribute "activeEventTitle", "string"
        attribute "activeEventStart", "string"
        attribute "activeEventEnd", "string"
        attribute "activeEventCount", "number"
        attribute "provider", "string"
        attribute "lastPoll", "string"
        attribute "lastError", "string"
        attribute "upcomingSummary", "string"

        // Upcoming (next 3)
        attribute "nextEvent1Title", "string"
        attribute "nextEvent1On", "string"
        attribute "nextEvent1Off", "string"
        attribute "nextEvent2Title", "string"
        attribute "nextEvent2On", "string"
        attribute "nextEvent2Off", "string"
        attribute "nextEvent3Title", "string"
        attribute "nextEvent3On", "string"
        attribute "nextEvent3Off", "string"

        command "setCalendarState", [[name: "onOff", type: "BOOL"], [name: "meta", type: "MAP"]]
    }
}

preferences {
    input name: "infoLogging", type: "bool", title: "Enable info logging?", defaultValue: true
    input name: "debugLogging", type: "bool", title: "Enable debug logging?", defaultValue: false
}

def installed() {
    sendEvent(name: "switch", value: "off")
    sendEvent(name: "activeEventCount", value: 0)
    clearUpcoming()
}

def updated() { }

def on() {
    // Manual override allowed; parent will correct on next poll
    sendEvent(name: "switch", value: "on")
    info("Manual on()")
}

def off() {
    sendEvent(name: "switch", value: "off")
    info("Manual off()")
}

def refresh() {
    debug("refresh() requested")
    try {
        parent?.childRefresh(device.deviceNetworkId)
    } catch (e) {
        sendEvent(name: "lastError", value: e.toString())
        warn("refresh() failed: ${e}")
    }
}

def setCalendarState(Boolean onOff, Map meta = [:]) {
    def desired = (onOff == true) ? "on" : "off"
    if (device.currentValue("switch") != desired) {
        sendEvent(name: "switch", value: desired)
    }

    Integer cnt = 0
    try { cnt = (meta?.activeCount instanceof Number) ? (meta.activeCount as Number).intValue() : 0 } catch (ignored) { }

    sendEvent(name: "activeEventCount", value: cnt)
    sendEvent(name: "provider", value: (meta?.provider ?: ""))

    sendEvent(name: "activeEventTitle", value: (meta?.title ?: ""))

    def startMs = meta?.startMs
    def endMs = meta?.endMs
    sendEvent(name: "activeEventStart", value: (startMs instanceof Number) ? new Date(startMs as Long).toString() : "")
    sendEvent(name: "activeEventEnd", value: (endMs instanceof Number) ? new Date(endMs as Long).toString() : "")

    // Upcoming events
    def nextEvents = (meta?.nextEvents instanceof List) ? (meta.nextEvents as List) : []
    applyUpcoming(nextEvents)

    sendEvent(name: "lastPoll", value: new Date().toString())
    sendEvent(name: "lastError", value: (meta?.lastError ?: ""))

    debug("setCalendarState(${desired}) meta keys=${meta?.keySet()}")
}

private void applyUpcoming(List nextEvents) {
    clearUpcoming()

    def items = nextEvents ?: []
    for (int i = 0; i < Math.min(items.size(), 3); i++) {
        def ev = items[i]
        def t = (ev?.title ?: "").toString()
        def onAt = (ev?.onAtMs instanceof Number) ? new Date(ev.onAtMs as Long).toString() : ""
        def offAt = (ev?.offAtMs instanceof Number) ? new Date(ev.offAtMs as Long).toString() : ""

        sendEvent(name: "nextEvent${i+1}Title", value: t)
        sendEvent(name: "nextEvent${i+1}On", value: onAt)
        sendEvent(name: "nextEvent${i+1}Off", value: offAt)
    }
}

private void clearUpcoming() {
    (1..3).each { n ->
        sendEvent(name: "nextEvent${n}Title", value: "")
        sendEvent(name: "nextEvent${n}On", value: "")
        sendEvent(name: "nextEvent${n}Off", value: "")
    }
}

private void debug(String msg) {
    if (settings.debugLogging == true) log.debug "${device.displayName}: ${msg}"
}
private void info(String msg) {
    if (settings.infoLogging != false) log.info "${device.displayName}: ${msg}"
}
private void warn(String msg) {
    log.warn "${device.displayName}: ${msg}"
}
