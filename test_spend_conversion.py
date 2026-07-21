import unittest

from openpyxl import Workbook

from process_live_ads import TARGET_ROWS, aggregate_rows, populate_summary, summary_values
from web_app import summarize_rows


class SpendConversionTest(unittest.TestCase):
    def setUp(self):
        self.rows = [{"名称": "roi", "总消耗": "415.2", "总成交金额": "830.4"}]

    def test_input_spend_is_yuan(self):
        values = summary_values(aggregate_rows(self.rows), "")

        self.assertEqual(values["总消耗金额"], 415.2)
        self.assertEqual(values["总消耗（豆）"], 4152)
        self.assertEqual(values["总成交roi"], 2)
        self.assertEqual(summarize_rows(self.rows)["spend_yuan"], 415.2)

    def test_excel_formulas_do_not_divide_input_spend(self):
        vertical = Workbook().active
        populate_summary(vertical, 2, transpose=True)
        self.assertEqual(vertical["B2"].value, "=B3*10")
        self.assertEqual(vertical["B3"].value, "=SUM('直播投放数据源'!$H$2:$H$2)")

        horizontal = Workbook().active
        populate_summary(horizontal, 2, transpose=False)
        self.assertEqual(horizontal["B2"].value, "=C2*10")
        self.assertEqual(horizontal["C2"].value, "=SUM('直播投放数据源'!$H$2:$H$2)")


if __name__ == "__main__":
    unittest.main()
