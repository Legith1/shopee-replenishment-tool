# -*- coding: utf-8 -*-
"""生成双端一致性验证用的合成测试数据（不参与业务，仅用于校验算法一致性）

运行：python tests/make_fixtures.py
产物：tests/fixtures/*.xlsx
"""
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures")
ROOT = os.path.dirname(HERE)

SALES = [
    ("cast pro左手", 1),
    ("CAST X3左手 普通版", 84),
    ("CAST X3左手 半陶版", 156),
    ("cast-x 左手", 23),
    ("CAST鱼竿C602", 23),
]

INV = [
    ("29486030649_243377971511", 18, 0),
    ("54901018899_410092717433", 137, 0),
    ("54901018899_410092717435", 47, 100),
    ("22619323581_213740302622", 55, 0),
    ("40300320923_260025872750", 12, 30),
]

SPEED = [
    ("29486030649_243377971511", 0.03),
    ("54901018899_410092717433", 2.83),
    ("54901018899_410092717435", 3.80),
    ("22619323581_213740302622", 0.83),
    ("40300320923_260025872750", 1.20),
]

# sku, qty, 长, 宽, 高, 重量(kg)
BOXSPEC = [
    ("cast pro左手", 20, 46, 36, 24, 12.5),
    ("CAST X3左手 普通版", 20, 42, 34, 22, 11.0),
    ("CAST X3左手 半陶版", 50, 58, 40, 30, 31.5),   # 故意超重，触发校验3
    ("cast-x 左手", 10, 120, 40, 30, 18.0),          # 故意超长，触发校验3
    ("CAST鱼竿C602", 5, 130, 12, 12, 9.8),
]


def build_basic():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "my销量"
    ws.append(["SKU", "商品SKU", "销售数量"])
    for sku, q in SALES:
        ws.append([sku, sku, q])
    wb.save(os.path.join(OUT, "我的马帮销量导出表.xlsx"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "my库存"
    ws.append(["仓库SKU ID", "商品名", "可售数量", "待处理的IR已到达"])
    for wh, a, t in INV:
        ws.append([wh, "", a, t])
    wb.save(os.path.join(OUT, "my库存导出表.xlsx"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["仓库SKU", "销售速度"])
    for wh, s in SPEED:
        ws.append([wh, s])
    wb.save(os.path.join(OUT, "my销售速度.xlsx"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["SKU", "箱规数量", "长(cm)", "宽(cm)", "高(cm)", "重量(kg)", "备注"])
    for sku, q, l, w, h, wt in BOXSPEC:
        ws.append([sku, q, l, w, h, wt, ""])
    wb.save(os.path.join(OUT, "补货箱规表.xlsx"))


def build_stage2():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "补货申请填写表"
    ws.append(["商品名称", "规格名称", "商品ID", "Warehouse SKU ID", "外部SKU ID",
               "* 申请入库数量 (个)", "* 单件出口申报货值 (RMB)"])
    for row in [
        ("SEASIR CAST X3", "Left hand", "1001", "54901018899_410092717433", "cast-x3-left"),
        ("SEASIR M800", "Left hand", "1002", "29486030649_243377971511", "M800-left"),
        ("UNKNOWN SKU", "x", "1003", "99999999_000", "unknown"),
    ]:
        ws.append(list(row))
    wb.save(os.path.join(OUT, "my入库申请模板.xlsx"))


def build_stage3():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "已打包数量"
    ws.append(["商品名称", "规格", "数量", "仓库SKU ID", "外部SKU ID", "状态", "总打包数"]
              + ["箱子 #%d" % i for i in range(1, 25)])
    for row in [
        ("Ryobiseasir DK2000", "std", 20, "54901018899_410092717433", "CAST X3左手 普通版", "", None),
        ("Ryobiseasir DK3000", "std", 30, "54901018899_410092717435", "CAST X3左手 半陶版", "", None),
        ("overcast鱼线0.6", "std", 20, "40300320923_260025872750", "cast-x 左手", "", None),
        ("overcast鱼线1.0", "std", 60, "", "cast pro左手", "", None),
        ("NOT-IN-SUMMARY", "std", 5, "", "zzz-none", "", None),
    ]:
        ws.append(list(row))
    ws2 = wb.create_sheet("尺寸和重量")
    ws2.append(["属性"] + ["箱子 #%d" % i for i in range(1, 25)])
    for name in ["长(cm)", "宽(cm)", "高(cm)", "重量(kg)"]:
        ws2.append([name])
    wb.save(os.path.join(OUT, "打包模板.xlsx"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "马来"
    ws.append(["11.7虾皮马来补货"])
    ws.append(["商品名称", "数量", "序号", "箱数"])
    ws.append([])
    for row in [("Ryobiseasir DK2000", 20, 1, ""),
                ("Ryobiseasir DK3000", 30, "3-7", ""),
                ("overcast鱼线0.6", 20, 19, ""),
                ("overcast鱼线1.0", 60, "", "")]:
        ws.append(list(row))
    wb.save(os.path.join(OUT, "汇总补货表.xlsx"))


def build():
    os.makedirs(OUT, exist_ok=True)
    build_basic()
    build_stage2()
    build_stage3()
    print("fixtures ->", OUT)
    for f in sorted(os.listdir(OUT)):
        print("  ", f)


if __name__ == "__main__":
    build()
