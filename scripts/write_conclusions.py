# -*- coding: utf-8 -*-
"""
高爷家云图周报结论回填脚本
通过 win32com 安全写入结论单元格，保留原始条件格式。
"""

import os
import shutil

# 从同目录引入 bundled 的 safe writer（无需外部依赖路径）
from excel_safe_writer import safe_update_cells

# 结论单元格映射（Excel 1-indexed）
CONCLUSION_CELLS = {
    "市场搜索趋势":   {(3, 2): None},           # B3
    "市场格局":       {(5, 2): None, (47, 2): None},  # B5, B47
    "品牌人群资产":   {(2, 8): None, (3, 14): None},  # H2, N3
    "人群和渠道数据": {(2, 3): None},           # C2
    "货品SPU资产新":  {(2, 10): None},          # J2
}


def write_conclusions(filepath, conclusions, output_suffix="_带结论"):
    """
    将结论写入 Excel 各 sheet 对应单元格。

    Parameters
    ----------
    filepath : str
        源 xlsx 文件路径
    conclusions : dict
        {sheet_name: {(row, col): text}, ...}
        键名必须匹配 CONCLUSION_CELLS 中的 sheet 名
    output_suffix : str
        输出文件名后缀，默认 "_带结论"

    Returns
    -------
    str — 输出文件路径
    """
    # 生成输出路径
    dirname = os.path.dirname(filepath)
    basename = os.path.basename(filepath)
    name, ext = os.path.splitext(basename)
    output_path = os.path.join(dirname, f"{name}{output_suffix}{ext}")

    # 复制源文件
    shutil.copy2(filepath, output_path)

    # 逐 sheet 写入
    for sheet_name, updates in conclusions.items():
        if not updates:
            continue
        print(f"写入 {sheet_name} ...")
        safe_update_cells(output_path, sheet_name, updates)

    print(f"完成: {output_path}")
    return output_path


def build_conclusion_map(conclusion_texts):
    """
    从简化的文本映射构建完整结论字典。
    自动按 CONCLUSION_CELLS 结构匹配位置。

    Parameters
    ----------
    conclusion_texts : dict
        简化格式：{sheet_name: text} 或 {sheet_name: {row: {col: text}}}
        对于多单元格 sheet（市场格局、品牌人群资产），
        使用 dict: {"市场格局": {(5,2): "猫砂", (47,2): "猫粮"}}

    Returns
    -------
    dict — 完整的 conclusions 字典
    """
    result = {}
    for sheet_name, value in conclusion_texts.items():
        if isinstance(value, dict):
            result[sheet_name] = value
        else:
            # 单值：取 CONCLUSION_CELLS 的第一个位置
            cells = CONCLUSION_CELLS.get(sheet_name, {})
            if cells:
                first_cell = next(iter(cells))
                result[sheet_name] = {first_cell: value}
            else:
                print(f"警告: 未知 sheet 名 '{sheet_name}'，跳过")
    return result


if __name__ == "__main__":
    # 使用示例
    conclusions = build_conclusion_map({
        "市场搜索趋势": "搜索词总结：品类大词【猫砂】环比+2.09%，TOP10中...",
        "市场格局": {
            (5, 2): "猫砂行业本周销售金额环比-2.77%，品牌侧...",
            (47, 2): "猫粮行业本周销售金额环比-6.03%，品牌侧...",
        },
        "品牌人群资产": {
            (2, 8): "人群资产整体环比增长，A2提升最大+10.85%...",
            (3, 14): "各链路GMV价值本周整体回升，环比+23.85%...",
        },
        "人群和渠道数据": "付费渠道方面，星图视频本竞品A2体量均有所增长...",
        "货品SPU资产新": "五个SPU整体A3人群环比+8.46%，其中...",
    })

    # write_conclusions("input.xlsx", conclusions)
    print("脚本就绪，结论单元格映射已加载。")
