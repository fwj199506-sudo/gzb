import unittest

from test_app_report_ordering import NamedBytesIO, app


def parse_rows(rows, total=100):
    text = '科目代码,科目名称,数量,成本,市值\n'
    text += f'资产合计,资产合计,,,{total}\n资产净值,资产净值,,,{total}\n'
    text += '\n'.join(','.join(map(str, row)) for row in rows)
    return app.process_valuation_files([NamedBytesIO(
        text.encode('utf-8-sig'), 'AAA001_柏治固定收益1号_资产估值表_20260715_四级.csv'
    )])


class BankDepositTests(unittest.TestCase):
    def check_bank(self, rows, expected=100, total=100):
        summary, details = parse_rows(rows, total)
        self.assertEqual(summary.iloc[0]['银行存款'], expected)
        bank = details[details['代码'] == '1002']
        self.assertEqual(len(bank), int(expected != 0))
        self.assertEqual(bank['今日市值'].sum(), expected)
        self.assertEqual(details['今日市值'].sum(), total)
        self.assertFalse((details['代码'] == 'RESIDUAL').any())
        allocation = app.build_position_allocation(summary, details)
        self.assertEqual(allocation.iloc[0]['现金类资产占比'], 1 if total else 0)

    def test_children_only(self):
        self.check_bank([('1002.01', '活期存款', '', 60, 60),
                         ('1002.02', '定期存款', '', 40, 40)])

    def test_top_preferred_regardless_of_order(self):
        children = [('1002.01', '活期存款', '', 60, 60),
                    ('1002.02', '定期存款', '', 40, 40)]
        for rows in [ [('1002', '银行存款', '', 100, 100)] + children,
                     children + [('1002', '银行存款', '', 100, 100)]]:
            with self.subTest(rows=rows):
                self.check_bank(rows)

    def test_intermediate_nodes_exclude_descendants(self):
        self.check_bank([('1002.01.01', '账户一', '', 60, 60),
                         ('1002.01', '活期存款', '', 60, 60),
                         ('1002.02.01', '账户二', '', 40, 40)])

    def test_sibling_prefix_is_not_parent(self):
        self.check_bank([('1002.01', '账户一', '', 60, 60),
                         ('1002.010', '账户二', '', 40, 40)])

    def test_zero_top_is_authoritative(self):
        self.check_bank([('1002', '银行存款', '', 0, 0),
                         ('1002.01', '活期存款', '', 100, 100)], expected=0, total=0)

    def test_name_only_compatibility(self):
        self.check_bank([('', '银行存款', '', 100, 100)])

    def test_blank_market_value_uses_cost(self):
        self.check_bank([('1002.01', '活期存款', '', 100, '')])

    def test_compact_codes_and_quantity_do_not_duplicate_cash(self):
        self.check_bank([('10020101', '账户一', 60, 60, 60),
                         ('100201', '活期存款', 60, 60, 60),
                         ('100202', '定期存款', 40, 40, 40)])


class FundClassificationTests(unittest.TestCase):
    def test_active_funds_are_other(self):
        for code, name, kind in [
            ('1105.01.01.000001.OF', '主动纯债基金', '公募/私募基金'),
            ('000001.OF', '主动纯债A', 'ETF/基金投资'),
            ('1105', '主动混合基金', 'ETF/基金投资'),
            ('000001.OF', '中证投资主动债券基金', '公募/私募基金'),
        ]:
            with self.subTest(name=name):
                self.assertEqual(app.classify_position_bucket(code, name, kind), '其他')

    def test_explicit_index_and_existing_buckets(self):
        for code, name, kind, expected in [
            ('510300.SH', '沪深300etf', 'ETF/基金投资', '指数类资产'),
            ('000001.OF', '沪深300指数基金', '公募/私募基金', '指数类资产'),
            ('1103.01', '信用债', '债券', '信用债'),
            ('1103.04', '测试可转债', '债券', '转债'),
            ('1102.01', '测试股票', '股票', '权益'),
            ('1002', '银行存款', '现金类资产', '现金类资产'),
            ('1108', '中证估值违约债权', '其他交易性金融资产', '其他'),
        ]:
            with self.subTest(name=name):
                self.assertEqual(app.classify_position_bucket(code, name, kind), expected)

    def test_fund_allocation_through_parser(self):
        summary, details = parse_rows([
            ('1105.01.01.000001.OF', '主动纯债A', 1, 60, 60),
            ('1105.01.01.510300.SH', '沪深300ETF', 1, 40, 40),
        ])
        allocation = app.build_position_allocation(summary, details).iloc[0]
        self.assertEqual(allocation['指数类资产占比'], 0.4)
        self.assertEqual(allocation['其他'], 0.6)
        self.assertEqual(details['今日市值'].sum(), 100)


if __name__ == '__main__':
    unittest.main()
