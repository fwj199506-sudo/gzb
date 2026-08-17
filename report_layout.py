import re

import pandas as pd


_CHINESE_NUMBER_TO_INT = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
}


def product_priority(product_name):
    """Return the requested display priority for a product name."""
    compact_name = re.sub(r"\s+", "", str(product_name or ""))
    fixed_income_match = re.search(
        r"(?:固定收益|固收)(?:第)?([1-6一二三四五六])号?",
        compact_name,
    )
    if fixed_income_match:
        number_text = fixed_income_match.group(1)
        number = int(number_text) if number_text.isdigit() else _CHINESE_NUMBER_TO_INT[number_text]
        return number - 1
    if "远迈" in compact_name:
        return 6
    if "稳鸿" in compact_name:
        return 7
    return 8


def sort_product_names(names):
    """Stable-sort unique product names using the shared report priority."""
    unique_names = list(dict.fromkeys(names))
    return sorted(unique_names, key=product_priority)


def sort_product_rows(df, product_col="产品名称"):
    """Stable-sort product rows while preserving all original columns."""
    if df is None:
        return df
    result = df.copy()
    if result.empty or product_col not in result.columns:
        return result

    sort_col = "__product_priority__"
    while sort_col in result.columns:
        sort_col = f"_{sort_col}"
    result[sort_col] = result[product_col].map(product_priority)
    return (
        result.sort_values(sort_col, kind="stable")
        .drop(columns=[sort_col])
        .reset_index(drop=True)
    )


def move_filename_column_to_end(df, filename_col="文件名"):
    """Return a copy with the requested file-name column in the last position."""
    if df is None:
        return df
    result = df.copy()
    if filename_col not in result.columns:
        return result
    columns = [column for column in result.columns if column != filename_col]
    return result[columns + [filename_col]]


def normalize_product_report(df, product_col="产品名称"):
    """Apply the shared product-row order and file-name column layout."""
    return move_filename_column_to_end(sort_product_rows(df, product_col))
