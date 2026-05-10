# -*- coding: utf-8 -*-
"""
Excel safe writer - 修改 xlsx 单元格值，不破坏任何格式/条件格式/公式。
原理：通过 Excel COM 接口写入，而非 openpyxl 重写整个文件。
仅适用于 Windows + Excel 已安装。
"""
import win32com.client
from openpyxl.utils import get_column_letter


def safe_update_cells(filepath, sheet_name, updates, recalc_ranges=None):
    """
    安全更新 Excel 单元格值，保留所有原始格式。

    Parameters
    ----------
    filepath : str
        xlsx 文件路径
    sheet_name : str
        sheet 名称
    updates : dict
        {(row, col): value, ...}  — 行号和列号均从 1 开始
    recalc_ranges : list of str, optional
        需要强制重算的区域，如 ['F119:F134', 'M119:M134']

    Returns
    -------
    bool — True 表示成功
    """
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    try:
        excel.Visible = False
    except:
        pass
    try:
        excel.DisplayAlerts = False
    except:
        pass

    try:
        wb = excel.Workbooks.Open(filepath)
        ws = wb.Worksheets(sheet_name)

        for (row, col), val in updates.items():
            col_letter = get_column_letter(col)
            ws.Range(f"{col_letter}{row}").Value = val

        # 强制重算公式区域
        if recalc_ranges:
            for rng in recalc_ranges:
                ws.Range(rng).Calculate()
        else:
            ws.EnableCalculation = True

        wb.Save()
        wb.Close()
        return True
    finally:
        try:
            excel.Quit()
        except:
            pass


# 使用示例
if __name__ == "__main__":
    safe_update_cells(
        filepath=r"C:\your_file.xlsx",
        sheet_name="Sheet1",
        updates={
            (10, 3): "新的结论文本",
            (15, 5): 12345.67,
        },
        recalc_ranges=["F119:F134", "M119:M134"]
    )
