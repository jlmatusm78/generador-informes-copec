import unittest
import inspect
from datetime import date

import pandas as pd

from report_engine import (
    ReportConfig,
    comparison_ranges,
    detailed_driver_ranking,
    operational_risk_summary,
    recurrence_summary,
    sensor_alert_summary,
    generate_transportista_report,
    filter_real_transportistas,
)


def frame(rows):
    return pd.DataFrame(rows, columns=["Conductor", "Tipo", "CumplimientoFatiga", "ID"])


class OperationalRiskTests(unittest.TestCase):
    def test_transportista_filter_uses_exact_generic_names(self):
        data = pd.DataFrame({"Transportista": [
            "COPEC",
            "GPS",
            "TRANSPORTES COPEC SUR",
            "LOGÍSTICA GPS CHILE",
            "TRANSPORTES PLANTA NORTE",
        ]})
        result = filter_real_transportistas(data)
        self.assertEqual(result["Transportista"].tolist(), [
            "TRANSPORTES COPEC SUR",
            "LOGÍSTICA GPS CHILE",
            "TRANSPORTES PLANTA NORTE",
        ])

    def test_weights_and_explanation(self):
        data = frame([
            ("Ana", "Fatiga", "NO", 1),
            ("Ana", "Sin cinturón", "PENDIENTE", 2),
            ("Luis", "Conductor fumando", "PENDIENTE", 3),
            ("Luis", "Fatiga", "SI", 4),
        ])
        risk = operational_risk_summary(data)
        self.assertEqual(risk["points"], 25)
        self.assertEqual(risk["level"], "ALTO")
        self.assertIn("1×10", risk["equation"])

    def test_repeated_unmet_fatigue_is_critical(self):
        data = frame([("Ana", "Fatiga", "NO", 1), ("Ana", "Fatiga", "NO", 2)])
        risk = operational_risk_summary(data)
        self.assertEqual(risk["level"], "CRÍTICO")
        self.assertEqual(risk["rule"], "reincidencia de fatiga sin cumplimiento")

    def test_recurrence_uses_two_or_more_events(self):
        data = frame([("Ana", "Uso celular", "PENDIENTE", 1), ("Ana", "Uso celular", "PENDIENTE", 2), ("Luis", "Uso celular", "PENDIENTE", 3)])
        result = recurrence_summary(data)
        self.assertEqual(result["recurrent"], 1)
        self.assertEqual(result["rate"], 50.0)

    def test_default_week_comparison_is_monday_to_sunday(self):
        config = ReportConfig(date(2026, 8, 3), date(2026, 8, 9))
        ranges = comparison_ranges(config)
        self.assertEqual(len(ranges), 4)
        self.assertEqual(ranges[-1][1].date().weekday(), 0)
        self.assertEqual(ranges[-1][2].date().weekday(), 6)

    def test_sensor_analysis_counts_affected_and_recurrent_tractos(self):
        current = pd.DataFrame({
            "Tipo": ["Sensor Tapado", "Sensor Tapado", "Sensor desalineado", "Sensor desalineado"],
            "Tracto": ["T1", "T1", "T1", "T2"],
        })
        previous = pd.DataFrame({"Tipo": ["Sensor Tapado"], "Tracto": ["T1"]})
        result = sensor_alert_summary(current, previous)
        self.assertEqual(result["total"], 4)
        self.assertEqual(result["affected"], 2)
        self.assertEqual(result["recurrent"], 1)
        self.assertEqual(result["both"], 1)
        self.assertEqual(result["priority"], "ALTA")

    def test_transportista_pdf_contains_sensor_analysis(self):
        source = inspect.getsource(generate_transportista_report)
        self.assertIn("Análisis técnico de alertas de sensores", source)
        self.assertIn("_sensor_summary_table", source)
        self.assertIn("_sensor_equipment_rows", source)
        self.assertNotIn("_write_transportista_excel", source)


if __name__ == "__main__":
    unittest.main()
