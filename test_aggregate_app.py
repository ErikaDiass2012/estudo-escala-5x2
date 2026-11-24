import unittest
from pathlib import Path

from aggregate_app import load_public_data
from openpyxl import load_workbook


class PublicAggregateAppTests(unittest.TestCase):
    def test_public_workbook_excludes_active_roster_sheet(self):
        workbook = load_workbook(
            Path(__file__).with_name("Simulador_HC_Escala_5x2.xlsx"),
            read_only=True,
        )
        try:
            self.assertNotIn("Ativos", workbook.sheetnames)
        finally:
            workbook.close()

    def test_public_loader_uses_aggregate_data_only(self):
        stores, premises, validation = load_public_data()

        self.assertEqual(len(stores), 30)
        self.assertEqual(int(stores["HC atual"].sum()), 292)
        self.assertEqual(int(stores["HC recomendado estimado"].sum()), 269)
        self.assertNotIn("Colaborador", stores.columns)
        self.assertNotIn("Cargo", stores.columns)
        self.assertEqual(validation["days_worked"], 5)
        self.assertEqual(validation["days_off"], 2)
        self.assertEqual(validation["active_count"], 269)
        self.assertEqual(validation["stores_valid"], 30)
        self.assertEqual(validation["stores_review"], 0)
        self.assertEqual(premises["Dias de baixo fluxo por semana"], 5)
        self.assertEqual(premises["Dias de fim de semana por semana"], 2)


if __name__ == "__main__":
    unittest.main()
