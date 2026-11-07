# -*- coding: utf-8 -*-
"""
Generate Standard Template for Shopee Replenishment (官方仓补货模板生成器)
支持 MY/TH/PH 三大站点专属标准格式：
- 马来/泰国：N列 60天，O列 75天（最高限制75天），P列 备注，Q列 出口货值
- 菲律宾：N列 60天（公式：60*销售速度-合计，最高限制60天），删除O列75天，O列 备注，P列 出口货值
- 彻底修复泰国模板 P/Q 列 #VALUE! 报错问题
"""

import copy
import time
import openpyxl

def generate_template():
    t0 = time.time()
    src_path = r"D:\桌面文件\虾皮补货文件\官方仓补货总表.xlsx"
    wb_src = openpyxl.load_workbook(src_path, data_only=False)
    print(f"Loaded source workbook in {time.time() - t0:.2f}s")

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)  # 移除默认Sheet

    # 1. 生成【马来补货模板】 (MY)
    ws_my_src = wb_src["11.2my"]
    ws_my = wb_out.create_sheet(title="马来补货模板")
    build_my_sheet(ws_my_src, ws_my)

    # 2. 生成【泰国补货模板】 (TH) - 修复 P/Q 列 #VALUE! 问题
    ws_th_src = wb_src["11.2th"]
    ws_th = wb_out.create_sheet(title="泰国补货模板")
    build_th_sheet(ws_th_src, ws_th)

    # 3. 生成【菲律宾补货模板】 (PH) - N列60天，删除75天列，O列备注，P列出口货值
    ws_ph_src = wb_src["11.2ph"]
    ws_ph = wb_out.create_sheet(title="菲律宾补货模板")
    build_ph_sheet(ws_ph_src, ws_ph)

    # 保存模板文件
    out_path1 = r"D:\桌面文件\虾皮补货文件\官方仓补货模板.xlsx"
    out_path2 = r"D:\Accio\work3\shopee_replenishment_tool\官方仓补货模板.xlsx"

    wb_out.save(out_path1)
    wb_out.save(out_path2)
    print(f"模板生成成功！总耗时: {time.time() - t0:.2f}s")
    print(f"已保存至: {out_path1}")
    print(f"已保存至: {out_path2}")


def build_my_sheet(ws_src, ws_dst):
    """构建马来模板 (标准17列)"""
    # 列宽与行高
    for col_letter, cd in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col_letter].width = cd.width
    for r in range(1, ws_src.max_row + 1):
        if r in ws_src.row_dimensions:
            ws_dst.row_dimensions[r].height = ws_src.row_dimensions[r].height

    headers = [
        "商品名称/MY", "SKU", "仓库SKU ID", "30天销量", "在仓", "在途", "合计",
        "补货至60天销售量", "剩余可售天数", "充足补货量", "虾皮马来需补", "采购批准",
        "销售速度", "60天", "75天", "备注", "出口货值"
    ]

    for c, h in enumerate(headers, 1):
        cell_src = ws_src.cell(1, c)
        cell_dst = ws_dst.cell(1, c, value=h)
        copy_cell_style(cell_src, cell_dst)

    for r in range(2, ws_src.max_row + 1):
        sku = ws_src.cell(r, 2).value
        wh_sku = ws_src.cell(r, 3).value
        if not sku and not wh_sku:
            continue

        name_val = ws_src.cell(r, 1).value
        rem_val = ws_src.cell(r, 16).value
        exp_val = ws_src.cell(r, 17).value

        for c in range(1, 18):
            cell_src = ws_src.cell(r, c)
            cell_dst = ws_dst.cell(r, c)
            copy_cell_style(cell_src, cell_dst)

        ws_dst.cell(r, 1, value=name_val)
        ws_dst.cell(r, 2, value=sku)
        ws_dst.cell(r, 3, value=wh_sku)
        ws_dst.cell(r, 4, value=None)   # 30天销量
        ws_dst.cell(r, 5, value=None)   # 在仓
        ws_dst.cell(r, 6, value=None)   # 在途
        ws_dst.cell(r, 7, value=f'=IF(AND(ISBLANK(E{r}),ISBLANK(F{r})),"",E{r}+F{r})')
        ws_dst.cell(r, 8, value=f'=IF(OR(ISBLANK(D{r}),D{r}=""),"",2*D{r})')
        ws_dst.cell(r, 9, value=f'=IF(OR(ISBLANK(D{r}),D{r}=0,D{r}=""),"",E{r}/(D{r}/30))')
        ws_dst.cell(r, 10, value=f'=IF(AND(ISBLANK(D{r}),ISBLANK(E{r}),ISBLANK(F{r})),"",H{r}-F{r}-E{r})')
        ws_dst.cell(r, 11, value=None)  # 需补
        ws_dst.cell(r, 12, value=None)  # 采购批准
        ws_dst.cell(r, 13, value=None)  # 销售速度
        ws_dst.cell(r, 14, value=f'=IF(OR(ISBLANK(M{r}),M{r}=""),"",60*M{r}-G{r})')
        ws_dst.cell(r, 15, value=f'=IF(OR(ISBLANK(M{r}),M{r}=""),"",75*M{r}-G{r})')
        ws_dst.cell(r, 16, value=rem_val)
        ws_dst.cell(r, 17, value=exp_val)

    # 复制条件格式
    for cf in ws_src.conditional_formatting._cf_rules:
        ws_dst.conditional_formatting._cf_rules[cf] = copy.deepcopy(ws_src.conditional_formatting._cf_rules[cf])


def build_th_sheet(ws_src, ws_dst):
    """构建泰国模板 (标准17列，修复原有P/Q列 #VALUE! 报错)"""
    for col_letter, cd in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col_letter].width = cd.width
    for r in range(1, ws_src.max_row + 1):
        if r in ws_src.row_dimensions:
            ws_dst.row_dimensions[r].height = ws_src.row_dimensions[r].height

    headers = [
        "商品名称/TH", "SKU", "仓库SKU ID", "30天销量", "在仓", "在途", "合计",
        "补货至60天销售量", "剩余可售天数", "充足补货量", "虾皮泰国需补", "采购批准",
        "销售速度", "60天", "75天", "备注", "出口货值"
    ]

    for c, h in enumerate(headers, 1):
        cell_src = ws_src.cell(1, c)
        cell_dst = ws_dst.cell(1, c, value=h)
        copy_cell_style(cell_src, cell_dst)

    for r in range(2, ws_src.max_row + 1):
        sku = ws_src.cell(r, 2).value
        wh_sku = ws_src.cell(r, 3).value
        if not sku and not wh_sku:
            continue

        name_val = ws_src.cell(r, 1).value
        # 从原始 11.2th 的对应列提取真实备注与货值，避免错位
        rem_val = ws_src.cell(r, 15).value or ws_src.cell(r, 18).value
        exp_val = ws_src.cell(r, 19).value or ws_src.cell(r, 17).value

        for c in range(1, 18):
            cell_src = ws_src.cell(r, min(c, ws_src.max_column))
            cell_dst = ws_dst.cell(r, c)
            copy_cell_style(cell_src, cell_dst)

        ws_dst.cell(r, 1, value=name_val)
        ws_dst.cell(r, 2, value=sku)
        ws_dst.cell(r, 3, value=wh_sku)
        ws_dst.cell(r, 4, value=None)   # 30天销量
        ws_dst.cell(r, 5, value=None)   # 在仓
        ws_dst.cell(r, 6, value=None)   # 在途
        ws_dst.cell(r, 7, value=f'=IF(AND(ISBLANK(E{r}),ISBLANK(F{r})),"",E{r}+F{r})')
        ws_dst.cell(r, 8, value=f'=IF(OR(ISBLANK(D{r}),D{r}=""),"",2*D{r})')
        ws_dst.cell(r, 9, value=f'=IF(OR(ISBLANK(D{r}),D{r}=0,D{r}=""),"",E{r}/(D{r}/30))')
        ws_dst.cell(r, 10, value=f'=IF(AND(ISBLANK(D{r}),ISBLANK(E{r}),ISBLANK(F{r})),"",H{r}-F{r}-E{r})')
        ws_dst.cell(r, 11, value=None)  # 需补
        ws_dst.cell(r, 12, value=None)  # 采购批准
        ws_dst.cell(r, 13, value=None)  # 销售速度
        # 正确公式：60*销售速度-合计，75*销售速度-合计
        ws_dst.cell(r, 14, value=f'=IF(OR(ISBLANK(M{r}),M{r}=""),"",60*M{r}-G{r})')
        ws_dst.cell(r, 15, value=f'=IF(OR(ISBLANK(M{r}),M{r}=""),"",75*M{r}-G{r})')
        ws_dst.cell(r, 16, value=rem_val)
        ws_dst.cell(r, 17, value=exp_val)

    # 复制条件格式
    for cf in ws_src.conditional_formatting._cf_rules:
        ws_dst.conditional_formatting._cf_rules[cf] = copy.deepcopy(ws_src.conditional_formatting._cf_rules[cf])


def build_ph_sheet(ws_src, ws_dst):
    """构建菲律宾模板 (专属16列：N列为60天公式，删除O列75天，O列备注，P列出口货值)"""
    for col_letter, cd in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col_letter].width = cd.width
    for r in range(1, ws_src.max_row + 1):
        if r in ws_src.row_dimensions:
            ws_dst.row_dimensions[r].height = ws_src.row_dimensions[r].height

    headers = [
        "商品名称/PH", "SKU", "仓库SKU ID", "30天销量", "在仓", "在途", "合计",
        "补货至60天销售量", "剩余可售天数", "充足补货量", "虾皮菲律宾需补", "采购批准",
        "销售速度", "60天", "备注", "出口货值"
    ]

    for c, h in enumerate(headers, 1):
        cell_src = ws_src.cell(1, min(c, ws_src.max_column))
        cell_dst = ws_dst.cell(1, c, value=h)
        copy_cell_style(cell_src, cell_dst)

    for r in range(2, ws_src.max_row + 1):
        sku = ws_src.cell(r, 2).value
        wh_sku = ws_src.cell(r, 3).value
        if not sku and not wh_sku:
            continue

        name_val = ws_src.cell(r, 1).value
        rem_val = ws_src.cell(r, 15).value
        exp_val = ws_src.cell(r, 17).value

        for c in range(1, 17):
            cell_src = ws_src.cell(r, min(c, ws_src.max_column))
            cell_dst = ws_dst.cell(r, c)
            copy_cell_style(cell_src, cell_dst)

        ws_dst.cell(r, 1, value=name_val)
        ws_dst.cell(r, 2, value=sku)
        ws_dst.cell(r, 3, value=wh_sku)
        ws_dst.cell(r, 4, value=None)   # 30天销量
        ws_dst.cell(r, 5, value=None)   # 在仓
        ws_dst.cell(r, 6, value=None)   # 在途
        ws_dst.cell(r, 7, value=f'=IF(AND(ISBLANK(E{r}),ISBLANK(F{r})),"",E{r}+F{r})')
        ws_dst.cell(r, 8, value=f'=IF(OR(ISBLANK(D{r}),D{r}=""),"",2*D{r})')
        ws_dst.cell(r, 9, value=f'=IF(OR(ISBLANK(D{r}),D{r}=0,D{r}=""),"",E{r}/(D{r}/30))')
        ws_dst.cell(r, 10, value=f'=IF(AND(ISBLANK(D{r}),ISBLANK(E{r}),ISBLANK(F{r})),"",H{r}-F{r}-E{r})')
        ws_dst.cell(r, 11, value=None)  # 需补
        ws_dst.cell(r, 12, value=None)  # 采购批准
        ws_dst.cell(r, 13, value=None)  # 销售速度
        # N列公式：销售速度*60-合计
        ws_dst.cell(r, 14, value=f'=IF(OR(ISBLANK(M{r}),M{r}=""),"",60*M{r}-G{r})')
        # 删除75天列，O列直接为备注，P列为出口货值
        ws_dst.cell(r, 15, value=rem_val)
        ws_dst.cell(r, 16, value=exp_val)

    # 复制条件格式
    for cf in ws_src.conditional_formatting._cf_rules:
        ws_dst.conditional_formatting._cf_rules[cf] = copy.deepcopy(ws_src.conditional_formatting._cf_rules[cf])


def copy_cell_style(src, dst):
    """复制单元格样式"""
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.number_format = copy.copy(src.number_format)
        dst.protection = copy.copy(src.protection)
        dst.alignment = copy.copy(src.alignment)


if __name__ == "__main__":
    generate_template()
