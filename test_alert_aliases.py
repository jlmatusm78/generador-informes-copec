import unittest

from alert_aliases import SENSOR_BLOCKED, SENSOR_MISALIGNED, normalize_alert_name
from report_engine import RISK_WEIGHTS, normalize_alert


class AlertAliasTests(unittest.TestCase):
    def test_exact_legacy_names_use_current_labels(self):
        self.assertEqual(normalize_alert_name("Cámara Desalineada"), SENSOR_MISALIGNED)
        self.assertEqual(normalize_alert_name("Alarma Tapado de cámara"), SENSOR_BLOCKED)

    def test_report_classification_accepts_old_and_new_names(self):
        for value in ("Cámara Desalineada", "Cámara desalineada", SENSOR_MISALIGNED):
            self.assertEqual(normalize_alert(value), SENSOR_MISALIGNED)
        for value in ("Alarma Tapado de cámara", "Tapado cámara", SENSOR_BLOCKED):
            self.assertEqual(normalize_alert(value), SENSOR_BLOCKED)

    def test_risk_weights_and_unrelated_categories_are_preserved(self):
        self.assertEqual(RISK_WEIGHTS["Fatiga sin cumplimiento"], 10)
        self.assertEqual(RISK_WEIGHTS["Sin cinturón"], 6)
        self.assertEqual(RISK_WEIGHTS["Conductor fumando"], 6)
        self.assertEqual(normalize_alert("Alarma uso de celular"), "Uso celular")
        self.assertEqual(normalize_alert("Alerta de bostezo"), "Bostezo")


if __name__ == "__main__":
    unittest.main()
