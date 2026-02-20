# -*- coding: utf-8 -*-
"""Rebuild an Excel pivot table with distinct-count clients by referrer."""

from __future__ import annotations

import argparse
import pathlib
import sys

import pythoncom
import win32com.client as win32


XL_SRC_RANGE = 1
XL_DATABASE = 1
XL_YES = 1
XL_ROW_FIELD = 1
XL_DISTINCT_COUNT = 11
XL_PIVOT_TABLE_VERSION15 = 5
XL_UP = -4162
XL_TO_LEFT = -4159


def get_header_col(ws, header_text: str):
    last_col = ws.Cells(1, ws.Columns.Count).End(XL_TO_LEFT).Column
    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value
        if value is not None and str(value).strip() == header_text:
            return col
    return None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Rebuild Excel pivot table using distinct count of account IDs."
    )
    parser.add_argument("--workbook", required=True, help="Path to xlsx workbook")
    parser.add_argument("--pivot-sheet", default="统计透视", help="Pivot sheet name")
    parser.add_argument(
        "--pivot-title",
        default="推荐人客户数统计（按资金账号非重复计数）",
        help="Pivot title text",
    )
    parser.add_argument("--table-name", default="订单明细", help="Source table name")
    parser.add_argument("--account-header", default="资金账号", help="Account header name")
    parser.add_argument("--referrer-header", default="推荐人姓名", help="Referrer header name")
    parser.add_argument("--legacy-helper-header", default="客户去重标记", help="Legacy helper header")
    args = parser.parse_args()

    workbook_path = pathlib.Path(args.workbook).expanduser().resolve()
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(workbook_path))

        source_ws = None

        for ws in list(wb.Worksheets):
            if ws.Name == args.pivot_sheet:
                ws.Delete()
                break

        for ws in list(wb.Worksheets):
            if ws.Name == args.pivot_sheet:
                continue
            acct_col = get_header_col(ws, args.account_header)
            ref_col = get_header_col(ws, args.referrer_header)
            if acct_col and ref_col:
                source_ws = ws
                break

        if source_ws is None:
            raise RuntimeError(
                f"No source sheet with headers '{args.account_header}' and '{args.referrer_header}'"
            )

        helper_col = get_header_col(source_ws, args.legacy_helper_header)
        if helper_col:
            source_ws.Columns(helper_col).Delete()

        acct_col = get_header_col(source_ws, args.account_header)
        ref_col = get_header_col(source_ws, args.referrer_header)
        if not acct_col or not ref_col:
            raise RuntimeError("Required headers missing after cleanup.")

        last_row = source_ws.Cells(source_ws.Rows.Count, acct_col).End(XL_UP).Row
        if last_row < 2:
            raise RuntimeError("No data rows found.")

        last_col = source_ws.Cells(1, source_ws.Columns.Count).End(XL_TO_LEFT).Column
        if last_col < max(acct_col, ref_col):
            last_col = max(acct_col, ref_col)

        source_range = source_ws.Range(source_ws.Cells(1, 1), source_ws.Cells(last_row, last_col))

        if source_ws.ListObjects.Count > 0:
            table = source_ws.ListObjects(1)
            table.Resize(source_range)
        else:
            table = source_ws.ListObjects.Add(XL_SRC_RANGE, source_range, None, XL_YES)

        try:
            table.Name = args.table_name
        except Exception:
            pass

        pivot_ws = wb.Worksheets.Add()
        pivot_ws.Name = args.pivot_sheet
        pivot_ws.Range("A1").Value = args.pivot_title

        pivot_cache = wb.PivotCaches().Create(XL_DATABASE, table.Range, XL_PIVOT_TABLE_VERSION15)
        pivot_table = pivot_cache.CreatePivotTable(pivot_ws.Range("A3"), "推荐人客户统计")

        pivot_row = pivot_table.PivotFields(args.referrer_header)
        pivot_row.Orientation = XL_ROW_FIELD
        pivot_row.Position = 1

        data_field = pivot_table.AddDataField(
            pivot_table.PivotFields(args.account_header),
            f"客户数(非重复{args.account_header})",
        )
        data_field.Function = XL_DISTINCT_COUNT

        pivot_table.PivotCache().RefreshOnFileOpen = True
        pivot_ws.Columns("A:B").AutoFit()

        wb.Save()
        print(f"Done: rebuilt pivot table in {workbook_path}")
        print(f"Sheet: {args.pivot_sheet}")
        print(f"Value field: Distinct Count of {args.account_header}")
        return 0
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
