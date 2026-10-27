import csv
import io
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from web_app import parse_screenshot_tsv, read_screenshot_metrics, screenshot_table_check


def tsv_lines(lines):
    stream = io.StringIO()
    writer = csv.writer(stream, delimiter="\t")
    writer.writerow("level page_num block_num par_num line_num word_num left top width height conf text".split())
    for index, (text, x, y, width, height) in enumerate(lines, 1):
        writer.writerow([5, 1, index, 1, 1, 1, x, y, width, height, 95, text])
    return stream.getvalue()


class ScreenshotOcrTests(unittest.TestCase):
    def test_positions_survive_unknown_headers_and_reordered_cards(self):
        lines = [
            ("数据汇总", 10, 10, 80, 20),
            ("消耗金额", 10, 60, 80, 20), ("未列入核对的指标", 210, 60, 160, 20),
            ("¥100.25", 10, 100, 100, 24), ("412", 210, 100, 50, 24),
            ("曝光人数", 10, 180, 80, 20), ("当场成交金额", 210, 180, 120, 20),
            ("2468", 10, 220, 70, 24), ("¥O", 210, 220, 36, 24),
            ("当场成交ROI", 10, 300, 120, 20), ("0", 10, 340, 20, 24),
        ]
        for scale, offset in [(1, 0), (0.65, 100), (2, 300)]:
            with self.subTest(scale=scale):
                records = parse_screenshot_tsv(tsv_lines([
                    (s, int(x * scale + offset), int(y * scale + offset), int(w * scale), int(h * scale))
                    for s, x, y, w, h in lines
                ]))
                self.assertEqual(records["spend_yuan"]["value"], 100.25)
                self.assertEqual(records["deal_amount"]["value"], 0)
                self.assertEqual(records["deal_roi"]["value"], 0)
                self.assertNotIn("order_roi", records)

    def test_direct_and_net_values_never_fill_current_session_metrics(self):
        records = parse_screenshot_tsv(tsv_lines([
            ("直接成交金额", 10, 10, 120, 20), ("¥200", 10, 50, 70, 24),
            ("净成交金额", 210, 10, 100, 20), ("¥300", 210, 50, 70, 24),
            ("当场成交金额", 410, 10, 120, 20), ("¥400", 410, 50, 70, 24),
        ]))
        self.assertEqual(records["deal_amount"]["value"], 400)

    def test_unknown_number_stays_missing_not_next_card_value(self):
        records = parse_screenshot_tsv(tsv_lines([
            ("当场成交金额", 10, 10, 120, 20), ("--", 10, 50, 30, 24),
            ("曝光人数", 10, 90, 80, 20), ("1234", 10, 130, 70, 24),
        ]))
        self.assertIsNone(records["deal_amount"]["value"])

    def test_same_name_inline_and_zero_values(self):
        records = parse_screenshot_tsv(tsv_lines([
            ("成交金额 ¥1,200.50", 10, 10, 240, 24),
            ("成交 ROI 0.0000", 10, 60, 240, 24),
            ("成交订单数 0", 10, 110, 180, 24),
        ]))
        self.assertEqual(records["deal_amount"]["value"], 1200.5)
        self.assertEqual(records["deal_roi"]["value"], 0)
        self.assertEqual(records["deal_orders"]["value"], 0)

    def test_details_below_summary_do_not_override_summary(self):
        records = parse_screenshot_tsv(tsv_lines([
            ("数据汇总", 10, 10, 80, 20),
            ("消耗金额", 10, 60, 80, 20), ("¥100", 10, 100, 70, 24),
            ("数据明细", 10, 180, 80, 20),
            ("消耗金额", 10, 220, 80, 20), ("¥20", 10, 260, 70, 24),
        ]))
        self.assertEqual(records["spend_yuan"]["value"], 100)

    @patch("web_app.tesseract_path", return_value="tesseract")
    @patch("web_app.subprocess.run")
    def test_small_images_merge_passes_without_changing_original(self, run, _):
        run.side_effect = [subprocess.CompletedProcess([], 0, tsv_lines(lines)) for lines in [
            [("消耗金额 ¥100.25", 10, 10, 240, 24), ("成交ROI 2", 10, 60, 120, 24)],
            [("当场成交金额 ¥200", 10, 10, 240, 24), ("成交ROI 3", 10, 60, 120, 24)],
        ]]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "small.png"
            Image.new("RGB", (640, 480), "white").save(path)
            original = path.read_bytes()
            records = read_screenshot_metrics(path)
            self.assertEqual(path.read_bytes(), original)
        self.assertEqual(records["spend_yuan"]["value"], 100.25)
        self.assertEqual(records["deal_amount"]["value"], 200)
        self.assertTrue(records["deal_roi"]["conflict"])
        self.assertIsNone(records["deal_roi"]["value"])
        self.assertEqual(run.call_count, 2)
        for call in run.call_args_list:
            self.assertFalse(Path(call.args[0][1]).exists())

    @patch("web_app.tesseract_path", return_value="tesseract")
    @patch("web_app.read_screenshot_metrics")
    def test_reconciliation_keeps_conflicts_missing_and_real_differences_distinct(self, read, _):
        rows = [{"总消耗": 100.25, "总成交金额": 0}]
        read.side_effect = [
            {"spend_yuan": {"value": 100.28, "label": "消耗金额"}, "deal_amount": {"value": 0, "label": "当场成交金额"}},
            {"spend_yuan": {"value": 100.28, "label": "消耗金额"}, "deal_amount": {"value": 25, "label": "当场成交金额"}},
        ]
        result = screenshot_table_check(rows, [Path("a.png"), Path("b.png")])
        by_key = {item["key"]: item for item in result["items"]}
        self.assertFalse(by_key["spend_yuan"]["match"])
        self.assertEqual(by_key["deal_amount"]["reason"], "conflict")
        self.assertIsNone(by_key["deal_amount"]["match"])
        self.assertEqual(by_key["order_roi"]["reason"], "label_missing")

    @patch("web_app.tesseract_path", return_value="tesseract")
    @patch("web_app.read_screenshot_metrics", side_effect=subprocess.TimeoutExpired("tesseract", 20))
    def test_timeout_does_not_fail_workbook_generation(self, read, _):
        result = screenshot_table_check([{"总消耗": 1}], [Path("a.png")])
        self.assertEqual(result["status"], "unavailable")
        self.assertIn("超时", result["message"])


if __name__ == "__main__":
    unittest.main()
