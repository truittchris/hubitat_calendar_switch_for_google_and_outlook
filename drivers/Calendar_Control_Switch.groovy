/**
 * truittchris Calendar Control Switch
 * Child driver for Calendar OAuth Bridge
 *
 * v0.1.0
 *
 * The parent app calls setCalendarState(onOff, meta) to:
 * - turn the switch on/off
 * - publish a few informative attributes for dashboards/logs
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

        command "setCalendarState", [[name: "onOff", type: "BOOL"], [name: "meta", type: "MAP"]]
    }
}

preferences {
    input name: "infoLogging", type: "bool", title: "Enable info logging?", defaultValue: true
    input name: "debugLogging", type: "bool", title: "Enable debug logging?", defaultValue: false
}

def installed() {
    sendEvent(name: "switch", value: "off")
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

    def title = (meta?.title ?: "")
    sendEvent(name: "activeEventTitle", value: title)

    def startMs = meta?.startMs
    def endMs = meta?.endMs
    if (startMs instanceof Number) {
        sendEvent(name: "activeEventStart", value: new Date(startMs as Long).toString())
    } else {
        sendEvent(name: "activeEventStart", value: "")
    }
    if (endMs instanceof Number) {
        sendEvent(name: "activeEventEnd", value: new Date(endMs as Long).toString())
    } else {
        sendEvent(name: "activeEventEnd", value: "")
    }

    sendEvent(name: "lastPoll", value: new Date().toString())
    sendEvent(name: "lastError", value: "")

    debug("setCalendarState(${desired}) meta=${meta}")
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
