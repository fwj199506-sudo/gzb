import io
import sys
import types
import unittest

import pandas as pd


def _install_streamlit_stub():
    streamlit = types.ModuleType("streamlit")
    streamlit.set_page_config = lambda **kwargs: None
    streamlit.title = lambda *args, **kwargs: None
    streamlit.markdown = lambda *args, **kwargs: None
    streamlit.file_uploader = lambda *args, **kwargs: []
    streamlit.info = lambda *args, **kwargs: None
    sys.modules["streamlit"] = streamlit


_install_streamlit_stub()
import app  # noqa: E402


class NamedBytesIO(io.BytesIO):
    def __init__(self, content, name):
        super().__init__(content)
        self.name = name


def _valuation_file(product_name, code_prefix):
    content = (
        "科目代码,科目名称,数量,成本,市值\n"
        "资产合计,资产合计,,,100\n"
        "资产净值,资产净值,,,90\n"
    ).encode("utf-8-sig")
    return NamedBytesIO(
        content,
        f"{code_prefix}_{product_name}_资产估值表_20260715_四级.csv",
    )


class AppReportOrderingTests(unittest.TestCase):
    def test_all_product_reports_use_requested_order_and_filename_last(self):
        uploaded_files = [
            _valuation_file("普通产品", "AAA001"),
            _valuation_file("柏治稳鸿私募证券投资基金", "AAA002"),
            _valuation_file("柏治远迈私募证券投资基金", "AAA003"),
            _valuation_file("柏治固定收益2号私募投资基金", "AAA004"),
            _valuation_file("柏治固定收益1号私募投资基金", "AAA005"),
        ]

        summary, details = app.process_valuation_files(uploaded_files)
        allocation = app.build_position_allocation(summary, details)
        validation = app.build_validation_report(summary, allocation, details)
        hierarchy = app.build_hierarchical_detail(summary, details)

        expected_products = [
            "柏治固定收益1号私募投资基金",
            "柏治固定收益2号私募投资基金",
            "柏治远迈私募证券投资基金",
            "柏治稳鸿私募证券投资基金",
            "普通产品",
        ]
        self.assertEqual(summary["产品名称"].tolist(), expected_products)
        self.assertEqual(allocation["产品名称"].tolist(), expected_products)
        self.assertEqual(validation["产品名称"].tolist(), expected_products)
        self.assertEqual(
            hierarchy.loc[hierarchy["层级"] == "产品", "行标签"].tolist(),
            expected_products,
        )
        self.assertEqual(
            details.drop_duplicates("所属产品")["所属产品"].tolist(),
            expected_products,
        )

        for frame in [summary, details, allocation, validation, hierarchy]:
            self.assertEqual(frame.columns[-1], "文件名")

    def test_cross_product_outputs_sort_product_names_and_place_files_last(self):
        summary = pd.DataFrame(
            {
                "产品名称": [
                    "普通产品",
                    "柏治稳鸿",
                    "柏治远迈",
                    "柏治固定收益2号",
                    "柏治固定收益1号",
                ],
                "估值日期": ["20260715"] * 5,
                "文件名": ["other.xls", "wh.xls", "ym.xls", "f2.xls", "f1.xls"],
                "银行存款": [0.0] * 5,
                "结算备付金及保证金": [0.0] * 5,
                "股票投资金额": [100.0] * 5,
                "债券投资金额": [0.0] * 5,
                "基金投资金额": [0.0] * 5,
                "买入返售金融资产": [0.0] * 5,
                "其他交易性金融资产": [0.0] * 5,
                "总资产": [100.0] * 5,
                "净资产": [100.0] * 5,
            }
        )
        details = pd.DataFrame(
            {
                "所属产品": summary["产品名称"],
                "估值日期": ["20260715"] * 5,
                "文件名": summary["文件名"],
                "资产类别": ["股票"] * 5,
                "代码": ["600000.SH"] * 5,
                "资产名称": ["测试股票"] * 5,
                "数量": [1.0] * 5,
                "单位成本": [100.0] * 5,
                "总成本": [100.0] * 5,
                "今日行情": [100.0] * 5,
                "今日市值": [100.0] * 5,
                "占比": [1.0] * 5,
            }
        )

        cross, comparison = app.build_cross_product_analysis(summary, details)

        self.assertEqual(
            cross.loc[0, "涉及产品"],
            "柏治固定收益1号、柏治固定收益2号、柏治远迈、柏治稳鸿、普通产品",
        )
        self.assertEqual(cross.columns[-1], "涉及文件")
        self.assertEqual(
            comparison["产品名称"].tolist(),
            ["柏治固定收益1号", "柏治固定收益2号", "柏治远迈", "柏治稳鸿", "普通产品"],
        )
        self.assertEqual(comparison.columns[-1], "文件名")


if __name__ == "__main__":
    unittest.main()
