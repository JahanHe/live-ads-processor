import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from PIL import Image

from process_live_ads import (
    SOURCE_HEADERS,
    aggregate_rows,
    build_workbook,
    default_output_stem,
    populate_summary,
    summary_values,
)
from web_app import parse_long_plan_ocr_text, screenshot_metrics_from_text, summarize_rows


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

    def test_new_standard_backend_export_normalizes_to_source_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            standard = root / "new_standard.csv"
            with standard.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=[
                        "订单信息",
                        "当场成交ROI",
                        "订单状态",
                        "投放进度",
                        "消耗详情",
                        "实际加热时长",
                        "新增关注数",
                        "平均千次展示费用",
                        "商品曝光次数",
                        "商品点击次数",
                        "千次观看成交金额",
                        "当场成交订单数",
                        "当场成交金额",
                        "当场下单金额",
                        "当场下单ROI",
                        "曝光人数",
                        "曝光人次",
                        "观看人数",
                        "进入率（人次）",
                        "进入率（人数）",
                        "观看人次",
                        "超过1分钟观看人数",
                        "直播间平均千次展示费用",
                        "点赞数",
                        "评论数",
                        "当场下单订单数",
                    ],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "订单信息": "roi1.2 30+三线 直播加热 1787500800_1897965",
                        "当场成交ROI": "4.2801",
                        "订单状态": "已结束-主动取消",
                        "投放进度": "¥84.11 / ¥400",
                        "消耗详情": "预算消耗/使用 ¥84.2/¥400\n应退金额 ¥315.8",
                        "实际加热时长": "17分钟",
                        "新增关注数": "2",
                        "平均千次展示费用": "¥66.86",
                        "商品曝光次数": "516",
                        "商品点击次数": "10",
                        "千次观看成交金额": "¥308.21",
                        "当场成交订单数": "1",
                        "当场成交金额": "¥360",
                        "当场下单金额": "¥360",
                        "当场下单ROI": "4.2801",
                        "曝光人数": "1168",
                        "曝光人次": "1258",
                        "观看人数": "453",
                        "进入率（人次）": "45.15%",
                        "进入率（人数）": "38.78%",
                        "观看人次": "568",
                        "超过1分钟观看人数": "49",
                        "直播间平均千次展示费用": "¥66.86",
                        "点赞数": "200",
                        "评论数": "6",
                        "当场下单订单数": "1",
                    }
                )

            output = root / "output.xlsx"
            build_workbook([standard], output)

            wb = load_workbook(output, data_only=True)
            source = wb["直播投放数据源"]
            summary = wb["数据汇总"]
            values = {summary.cell(row, 1).value: summary.cell(row, 2).value for row in range(2, summary.max_row + 1)}
            self.assertEqual(source["A2"].value, "吴晓波频道")
            self.assertEqual(source["B2"].value, "2026年8月24日")
            self.assertEqual(source["C2"].value, "roi1.2 30+三线 直播加热 1787500800_1897965")
            self.assertEqual(source["D2"].value, "1787500800_1897965")
            self.assertEqual(source["G2"].value, 1020)
            self.assertEqual(source["S2"].value, 1)
            self.assertEqual(values["总消耗金额"], 84.11)
            self.assertEqual(values["总成交金额"], 360)
            self.assertEqual(values["曝光人数"], 1168)
            self.assertEqual(values["观看人数"], 453)

    def test_summary_inserts_original_screenshot_without_combined_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.csv"
            image = root / "screenshot.png"
            output = root / "output.xlsx"
            Image.new("RGB", (240, 120), "white").save(image)
            with source.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=SOURCE_HEADERS)
                writer.writeheader()
                writer.writerow({"名称": "roi", "总消耗": 100, "总成交金额": 200})

            result = build_workbook([source], output, image_paths=[image], combined_image_path=root / "unused.png")

            wb = load_workbook(output)
            self.assertIsNone(result["combined_image"])
            self.assertFalse((root / "unused.png").exists())
            self.assertEqual(len(wb["数据汇总"]._images), 1)

    def test_detail_export_is_primary_metric_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            detail = root / "detail.csv"
            with detail.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=[
                        "订单/计划",
                        "消耗金额",
                        "曝光人数",
                        "观看人数",
                        "新增关注数",
                        "评论数",
                        "点赞数",
                        "超过1分钟观看人数",
                        "当场成交金额",
                        "当场成交ROI",
                        "当场成交订单数",
                        "当场下单金额",
                        "当场下单订单数",
                    ],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "订单/计划": "roi1.2 30+三线 直播间标准订单 1787500800_1897965",
                        "消耗金额": "￥84.17",
                        "曝光人数": "1168",
                        "观看人数": "495",
                        "新增关注数": "2",
                        "评论数": "6",
                        "点赞数": "200",
                        "超过1分钟观看人数": "934",
                        "当场成交金额": "￥360",
                        "当场成交ROI": "4.277",
                        "当场成交订单数": "1",
                        "当场下单金额": "￥360",
                        "当场下单订单数": "1",
                    }
                )

            output = root / "output.xlsx"
            build_workbook([detail], output)

            wb = load_workbook(output, data_only=True)
            source = wb["直播投放数据源"]
            summary = wb["数据汇总"]
            values = {summary.cell(row, 1).value: summary.cell(row, 2).value for row in range(2, summary.max_row + 1)}
            self.assertEqual(source["A2"].value, "吴晓波频道")
            self.assertEqual(source["B2"].value, "2026年8月24日")
            self.assertEqual(source["D2"].value, "1787500800_1897965")
            self.assertEqual(values["总消耗金额"], 84.17)
            self.assertEqual(values["观看人数"], 495)
            self.assertEqual(values["超过1分钟观看人数"], 934)
            self.assertEqual(values["总成交金额"], 360)
            self.assertEqual(values["总下单金额"], 360)

    def test_latest_detail_export_allows_missing_likes_and_uses_author(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            detail = root / "latest_detail.csv"
            with detail.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=[
                        "作者",
                        "直播间场次",
                        "订单/计划",
                        "消耗金额",
                        "曝光人数",
                        "观看人数",
                        "新增关注数",
                        "评论数",
                        "超过1分钟观看人次",
                        "当场成交金额",
                        "当场成交ROI",
                        "当场成交订单数",
                        "当场下单金额",
                        "当场下单订单数",
                    ],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "作者": "金海曾易英语 教育自媒体",
                        "直播间场次": "2026年08月23日",
                        "订单/计划": "roi1.6 直播间标准订单 1787414400_220786",
                        "消耗金额": "￥203.8",
                        "曝光人数": "582",
                        "观看人数": "124",
                        "新增关注数": "0",
                        "评论数": "32",
                        "超过1分钟观看人次": "728",
                        "当场成交金额": "￥399",
                        "当场成交ROI": "1.9577",
                        "当场成交订单数": "1",
                        "当场下单金额": "￥399",
                        "当场下单订单数": "1",
                    }
                )

            output = root / "output.xlsx"
            build_workbook([detail], output)

            wb = load_workbook(output, data_only=True)
            source = wb["直播投放数据源"]
            summary = wb["数据汇总"]
            values = {summary.cell(row, 1).value: summary.cell(row, 2).value for row in range(2, summary.max_row + 1)}
            self.assertEqual(source["A2"].value, "金海曾易英语 教育自媒体")
            self.assertEqual(source["B2"].value, "2026年8月23日")
            self.assertEqual(values["总消耗金额"], 203.8)
            self.assertEqual(values["点赞次数"], 0)
            self.assertEqual(values["评论次数"], 32)
            self.assertEqual(values["超过1分钟观看人数"], 728)

    def test_default_output_name_removes_author_category_suffix(self):
        rows = [{"加热主播": "金海曾易英语 教育自媒体", "加热开始时间": "2026年8月25日"}]

        self.assertEqual(default_output_stem(rows), "2026年8月25日金海曾易英语直播投放数据")

    def test_screenshot_metric_text_parses_standard_summary(self):
        text = """
数据汇总
消耗金额 ¥509.7
直接下单金额 ¥1,197
直接下单ROI 2.3484
直接成交金额 ¥1,197
当场下单订单数 3
当场下单金额 ¥1,197
当场成交订单数 3
当场下单ROI 2.3484
当场成交金额 ¥1,197
当场成交ROI 2.3484
"""

        metrics = screenshot_metrics_from_text(text)

        self.assertEqual(metrics["spend_yuan"], 509.7)
        self.assertEqual(metrics["order_amount"], 1197)
        self.assertEqual(metrics["order_orders"], 3)
        self.assertEqual(metrics["order_roi"], 2.3484)
        self.assertEqual(metrics["deal_amount"], 1197)
        self.assertEqual(metrics["deal_orders"], 3)
        self.assertEqual(metrics["deal_roi"], 2.3484)

    def test_screenshot_metric_text_parses_ocr_layout_rows(self):
        text = """
数据汇总
freee OH                      WHER O                                            平均千次展示费用 O                                   直接下单金额 O                                        直接下单ROL 〇
¥478.53                               ¥240.22                               ¥798                                  1.6675
直接成交金额 O                                     当场下单订单数 O                                  当场下单金额 OD                                     当场成交订单数 O
¥798                                  3                                      ¥1,197                                 4
当场下单ROI O                                           当场成交金额 O                                          当场成交ROI O
2.5013                                ¥1,596                                3.3351
"""

        metrics = screenshot_metrics_from_text(text)

        self.assertEqual(metrics["spend_yuan"], 478.53)
        self.assertEqual(metrics["order_amount"], 1197)
        self.assertEqual(metrics["order_orders"], 3)
        self.assertEqual(metrics["order_roi"], 2.5013)
        self.assertEqual(metrics["deal_amount"], 1596)
        self.assertEqual(metrics["deal_orders"], 4)
        self.assertEqual(metrics["deal_roi"], 3.3351)

    def test_screenshot_metric_text_parses_same_name_short_headers(self):
        text = """
数据汇总
消耗金额                     下单金额                     下单 ROI                     下单订单数                     成交金额
¥3,120.66                   ¥39,006.25                  12.4994                     214                         ¥30,389.26
成交 ROI                    成交订单数
9.7381                      190
"""

        metrics = screenshot_metrics_from_text(text)

        self.assertEqual(metrics["spend_yuan"], 3120.66)
        self.assertEqual(metrics["order_amount"], 39006.25)
        self.assertEqual(metrics["order_roi"], 12.4994)
        self.assertEqual(metrics["order_orders"], 214)
        self.assertEqual(metrics["deal_amount"], 30389.26)
        self.assertEqual(metrics["deal_roi"], 9.7381)
        self.assertEqual(metrics["deal_orders"], 190)

    def test_detail_export_overrides_standard_order_export_when_ids_match(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            standard = root / "standard.csv"
            detail = root / "detail.csv"
            with standard.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=SOURCE_HEADERS)
                writer.writeheader()
                writer.writerow(
                    {
                        "名称": "roi 标准订单 1787500800_1897965",
                        "编号": "1787500800_1897965",
                        "总消耗": 10,
                        "总成交金额": 20,
                    }
                )
            with detail.open("w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=[
                        "订单/计划",
                        "消耗金额",
                        "曝光人数",
                        "观看人数",
                        "新增关注数",
                        "评论数",
                        "点赞数",
                        "超过1分钟观看人数",
                        "当场成交金额",
                        "当场成交ROI",
                        "当场成交订单数",
                        "当场下单金额",
                        "当场下单订单数",
                    ],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "订单/计划": "roi 数据明细 1787500800_1897965",
                        "消耗金额": "￥84.17",
                        "曝光人数": "1168",
                        "观看人数": "495",
                        "新增关注数": "2",
                        "评论数": "6",
                        "点赞数": "200",
                        "超过1分钟观看人数": "934",
                        "当场成交金额": "￥360",
                        "当场成交ROI": "4.277",
                        "当场成交订单数": "1",
                        "当场下单金额": "￥360",
                        "当场下单订单数": "1",
                    }
                )

            output = root / "output.xlsx"
            build_workbook([standard, detail], output)

            wb = load_workbook(output, data_only=True)
            source = wb["直播投放数据源"]
            summary = wb["数据汇总"]
            values = {summary.cell(row, 1).value: summary.cell(row, 2).value for row in range(2, summary.max_row + 1)}
            self.assertEqual(source.max_row, 2)
            self.assertEqual(source["C2"].value, "roi 数据明细 1787500800_1897965")
            self.assertEqual(values["总消耗金额"], 84.17)
            self.assertEqual(values["总成交金额"], 360)

    def test_long_plan_ocr_handles_noisy_full_page_capture(self):
        text = """
直播间加热效果 —- 2026-03-05 = 2026-08-13 f=}
消耗总金额               BSA               进入总人数               点跑总次数               评论总次数               新增总关注
¥81293.34             312595              73075                7476                4062                1005
电商加热效果 —«- 2026-03-05 = 2026-08-13 [=]
成交ROI               直接成交ROI              当场成交GMV             直接成交GMV             当场成交订单数            直接成交订单数
6.4698              6.4698            ¥627835.47          ¥627835.47             3369                3369
净成交订单数              净成交金额               净成交ROI               当场下单GMV             直接下单GMV             当场下单订单数
2837             ¥466428.57            4.8065            ¥746545.58          ¥746545.58             3725
直接下单订单数              商品点击人数              商品点击次数
3725                6843                17027
"""

        fields = parse_long_plan_ocr_text(text)

        self.assertEqual(fields["消耗总金额"], "81293.34")
        self.assertEqual(fields["曝光总人数"], "312595")
        self.assertEqual(fields["进入总人数"], "73075")
        self.assertEqual(fields["点赞总次数"], "7476")
        self.assertEqual(fields["评论总次数"], "4062")
        self.assertEqual(fields["新增总关注"], "1005")
        self.assertEqual(fields["成交GMV"], "627835.47")
        self.assertEqual(fields["成交订单数"], "3369")
        self.assertEqual(fields["下单GMV"], "746545.58")
        self.assertEqual(fields["下单订单数"], "3725")

    def test_long_plan_ocr_handles_cropped_live_block(self):
        text = """
直播间加热效果 —«-2026-08-12 = 2026-08-12 =)
消耗总金额                 HBA                 进入总人数                 点跑总次数                 评论总次数                 新增总关注
¥729.96                 6322                  1500                   360                    39                     17
电商加热效果 —«- 2026-03-05 = 2026-08-13 [=]
成交ROI                 直接成交ROI               当场成交GMV               直接成交GMV              当场成交订单数              直接成交订单数
6.4698                6.4698             ¥627835.47           ¥627835.47              3369                  3369
净成交订单数                净成交金额                 净成交ROI                当场下单GMV               直接下单GMV              当场下单订单数
2837              ¥466428.57             4.8065             ¥746545.58           ¥746545.58               3725
直接下单订单数               商品点击人数               商品点击次数
3725                  6843                 17027
"""

        fields = parse_long_plan_ocr_text(text)

        self.assertEqual(fields["消耗总金额"], "729.96")
        self.assertEqual(fields["曝光总人数"], "6322")
        self.assertEqual(fields["进入总人数"], "1500")
        self.assertEqual(fields["点赞总次数"], "360")
        self.assertEqual(fields["评论总次数"], "39")
        self.assertEqual(fields["新增总关注"], "17")
        self.assertEqual(fields["成交GMV"], "627835.47")
        self.assertEqual(fields["成交订单数"], "3369")
        self.assertEqual(fields["下单GMV"], "746545.58")
        self.assertEqual(fields["下单订单数"], "3725")


if __name__ == "__main__":
    unittest.main()
