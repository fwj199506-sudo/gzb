import unittest

import pandas as pd

from report_layout import (
    move_filename_column_to_end,
    normalize_product_report,
    sort_product_names,
    sort_product_rows,
)


class ReportLayoutTests(unittest.TestCase):
    def test_product_rows_follow_requested_priority_and_keep_other_order(self):
        frame = pd.DataFrame(
            {
                "产品名称": [
                    "其他乙",
                    "柏治固定收益6号私募投资基金",
                    "柏治稳鸿2号私募证券投资基金",
                    "柏治固定收益1号私募投资基金",
                    "柏治远迈私募证券投资基金",
                    "其他甲",
                ],
                "文件名": ["b", "f6", "wh", "f1", "ym", "a"],
            }
        )

        result = sort_product_rows(frame, "产品名称")

        self.assertEqual(
            result["产品名称"].tolist(),
            [
                "柏治固定收益1号私募投资基金",
                "柏治固定收益6号私募投资基金",
                "柏治远迈私募证券投资基金",
                "柏治稳鸿2号私募证券投资基金",
                "其他乙",
                "其他甲",
            ],
        )

    def test_product_name_lists_use_the_same_priority(self):
        self.assertEqual(
            sort_product_names(
                [
                    "其他",
                    "柏治固收三号",
                    "柏治稳鸿",
                    "柏治远迈",
                    "柏治固定收益2号",
                ]
            ),
            [
                "柏治固定收益2号",
                "柏治固收三号",
                "柏治远迈",
                "柏治稳鸿",
                "其他",
            ],
        )

    def test_filename_column_is_moved_to_the_end(self):
        frame = pd.DataFrame(
            {"产品名称": ["A"], "文件名": ["a.xls"], "总资产": [1.0]}
        )

        result = move_filename_column_to_end(frame)

        self.assertEqual(result.columns.tolist(), ["产品名称", "总资产", "文件名"])

    def test_normalize_product_report_applies_both_rules(self):
        frame = pd.DataFrame(
            {
                "产品名称": ["普通产品", "柏治固定收益2号"],
                "文件名": ["other.xls", "fixed.xls"],
                "总资产": [1.0, 2.0],
            }
        )

        result = normalize_product_report(frame)

        self.assertEqual(result["产品名称"].tolist(), ["柏治固定收益2号", "普通产品"])
        self.assertEqual(result.columns[-1], "文件名")

    def test_missing_product_or_filename_column_is_left_safe(self):
        frame = pd.DataFrame({"资产名称": ["债券A"], "今日市值": [10.0]})

        sorted_result = sort_product_rows(frame, "产品名称")
        column_result = move_filename_column_to_end(frame)

        pd.testing.assert_frame_equal(sorted_result, frame)
        pd.testing.assert_frame_equal(column_result, frame)


if __name__ == "__main__":
    unittest.main()
