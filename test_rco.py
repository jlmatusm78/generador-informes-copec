import unittest
import math
import pandas as pd
from report_engine import recurrence_summary, operational_risk_summary, detailed_driver_ranking
from rco_sections import rco_summary, driver_count

class RcoTests(unittest.TestCase):
    def frame(self,names,companies=None):
        return pd.DataFrame({'Conductor':names,'Transportista':companies or ['A']*len(names),'Tipo':['Fatiga']*len(names),'CumplimientoFatiga':['NO']*len(names)})
    def test_unknowns_keep_points_without_false_critical_recurrence(self):
        d=self.frame(['Sin RCO',' sin r.c.o. '])
        risk=operational_risk_summary(d)
        self.assertEqual(risk['points'],20)
        self.assertEqual(risk['level'],'MEDIO')
        self.assertEqual(risk['repeated_fatigue_no_drivers'],0)
        self.assertTrue(math.isnan(recurrence_summary(d)['rate']))
        rank=detailed_driver_ranking(d,d,include_transportista=True)
        self.assertTrue(rank.empty)
        self.assertIn('Riesgo',rank)
    def test_missing_is_separate_from_sin_rco(self):
        s=rco_summary(self.frame(['Sin RCO',None,'Ana','Ana']))
        self.assertEqual((s['sin_rco'],s['other'],s['identified'],s['pct']),(1,1,2,25))
    def test_same_name_in_different_companies_is_not_recurrent(self):
        d=self.frame(['Ana','Ana'],['A','B'])
        self.assertEqual(driver_count(d),2)
        self.assertEqual(recurrence_summary(d)['recurrent'],0)
        self.assertEqual(operational_risk_summary(d)['repeated_fatigue_no_drivers'],0)
        old=self.frame(['Ana','Ana','Ana'],['A','A','B'])
        rank=detailed_driver_ranking(d,old,include_transportista=True)
        self.assertEqual(dict(zip(rank.Transportista,rank.Anterior)),{'A':2,'B':1})
    def test_empty_percent(self):
        self.assertIsNone(rco_summary(self.frame([]))['pct'])
