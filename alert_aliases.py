"""Canonical alert labels and backwards-compatible aliases."""

SENSOR_MISALIGNED = "Sensor desalineado"
SENSOR_BLOCKED = "Sensor Tapado"

ALERT_NAME_ALIASES = {
    "Cámara Desalineada": SENSOR_MISALIGNED,
    "Alarma Tapado de cámara": SENSOR_BLOCKED,
}

def normalize_alert_name(value):
    """Return the current label while accepting exact legacy names."""
    return ALERT_NAME_ALIASES.get(value, value)
