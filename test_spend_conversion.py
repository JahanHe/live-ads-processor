import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from process_live_ads import SOURCE_HEADERS, aggregate_rows, build_workbook, populate_summary, summary_values
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

    def test_long_plan_rows_merge_with_standard_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            standard = root / "standard.csv"
            plan = root / "plan.csv"
            output = root / "output.xlsx"
            with standard.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=SOURCE_HEADERS)
                writer.writeheader()
                writer.writerow({"名称": "标准订单", "总消耗": 100, "总成交金额": 10})
            with plan.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["消耗总金额", "曝光总人数", "当场成交GMV"])
                writer.writeheader()
                writer.writerow({"消耗总金额": "¥1416.5", "曝光总人数": 4700, "当场成交GMV": "¥5553.55"})

            build_workbook([standard], output, long_plan_paths=[plan])

            wb = load_workbook(output, data_only=True)
            summary = wb["数据汇总"]
            values = {summary.cell(row, 1).value: summary.cell(row, 2).value for row in range(2, summary.max_row + 1)}
            self.assertEqual(wb["直播投放数据源"].max_row, 3)
            self.assertEqual(values["总消耗金额"], 1516.5)
            self.assertEqual(values["总消耗（豆）"], 15165)
            self.assertEqual(values["曝光人数"], 4700)
            self.assertEqual(values["总成交金额"], 5563.55)


if __name__ == "__main__":
    unittest.main()
