import unittest
from datetime import date

import pandas as pd

from report_engine import (
    ReportConfig,
    comparison_ranges,
    detailed_driver_ranking,
    operational_risk_summary,
    recurrence_summary,
)


def frame(rows):
    return pd.DataFrame(rows, columns=["Conductor", "Tipo", "CumplimientoFatiga", "ID"])


class OperationalRiskTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
