# -*- coding: utf-8 -*-
"""
Shopee Official Warehouse Full-Stage Replenishment Command Line Interface (CLI)
"""

import os
import sys
import argparse
import datetime
from core import ReplenishmentEngine

def main():
    parser = argparse.ArgumentParser(description="Shopee 官方仓全流程智能补货工具 (CLI)")
    parser.add_argument("--stage", "-t", choices=["1", "2", "3", "all"], default="1", help="执行阶段: 1-补货总表, 2-入库申请表, 3-打包装箱表, all-全流程")
    parser.add_argument("--site", choices=["my", "th", "ph"], default="my", help="站点代号: my / th / ph")
    parser.add_argument("--date", help="批次日期 (如 11.7, 默认今天)")

    # 阶段一参数
    parser.add_argument("--base", "-b", help="补货底表/模板路径")
    parser.add_argument("--sales", "-s", help="马帮销量文件路径")
    parser.add_argument("--inv", "-i", help="Shopee官方仓库存文件路径")
    parser.add_argument("--rpa", "-r", help="RPA库存/销售速度文件路径")
    parser.add_argument("--box", "-k", help="补货箱规表路径")

    # 阶段二参数
    parser.add_argument("--master", help="已审批的补货总表路径")
    parser.add_argument("--inbound-tpl", help="Shopee入库申请模板路径")

    # 阶段三参数
    parser.add_argument("--summary", help="汇总补货表路径 (如 11.2虾皮补货.xlsx)")
    parser.add_argument("--pack-tpl", help="Shopee官方打包模板路径 (如 11.2虾皮马来.xlsx)")

    parser.add_argument("--output", "-o", help="输出文件路径")

    args = parser.parse_args()

    today = datetime.datetime.now()
    date_str = args.date or f"{today.month}.{today.day}"
    site_code = args.site.lower()
    box_spec_path = args.box or r"D:\桌面文件\虾皮补货文件\补货箱规表.xlsx"

    engine = ReplenishmentEngine(box_spec_path if os.path.exists(box_spec_path) else None)
    print(f"[*] 已初始化补货引擎，当前加载箱规: {len(engine.box_specs)} 条")

    # 执行阶段一
    if args.stage in ["1", "all"]:
        base_f = args.base or r"D:\桌面文件\虾皮补货文件\官方仓补货总表_新.xlsx"
        if not os.path.exists(base_f):
            base_f = r"D:\桌面文件\虾皮补货文件\官方仓补货模板.xlsx"
        sales_f = args.sales or base_f
        inv_f = args.inv or base_f
        rpa_f = args.rpa or r"C:\Users\Administrator\Downloads\_jtr_rpa_shopee_inventory-1.xlsx"

        out_f = args.output or os.path.join(os.path.dirname(os.path.abspath(base_f)), f"官方仓补货总表_{date_str}{site_code}.xlsx")
        new_sheet = f"{date_str}{site_code}"

        print(f"[*] 执行阶段一：生成补货总表 [站点: {site_code.upper()}, Sheet: {new_sheet}]...")
        sales_map = engine.load_sales_data(sales_f, sheet_name=f"{site_code}销量")
        inv_map = engine.load_shopee_inventory(inv_f, sheet_name=f"{site_code}库存")
        rpa_map = engine.load_rpa_inventory(rpa_f) if os.path.exists(rpa_f) else {}

        out_path, _ = engine.process_stage1_replenish(
            base_workbook_path=base_f,
            template_sheet_name=f"{site_code}补货模板",
            output_workbook_path=out_f,
            new_sheet_name=new_sheet,
            sales_map=sales_map,
            inv_map=inv_map,
            rpa_map=rpa_map,
            site_code=site_code
        )
        print(f"[+] 阶段一完成！已生成: {out_path}")

    # 执行阶段二
    if args.stage in ["2", "all"]:
        master_f = args.master or args.base or r"D:\桌面文件\虾皮补货文件\官方仓补货总表_新.xlsx"
        inb_tpl_f = args.inbound_tpl or r"D:\桌面文件\虾皮补货文件\my入库申请模板.xlsx"
        out_inb = args.output or os.path.join(os.path.dirname(os.path.abspath(inb_tpl_f)), f"{date_str}{site_code}_入库申请填写表_已填充.xlsx")

        print(f"[*] 执行阶段二：生成入库申请表 [站点: {site_code.upper()}]...")
        out_path, filled_cnt, _ = engine.process_stage2_inbound(
            approved_master_path=master_f,
            inbound_template_path=inb_tpl_f,
            site_code=site_code,
            output_path=out_inb
        )
        print(f"[+] 阶段二完成！已填报 {filled_cnt} 条批准SKU，生成: {out_path}")

    # 执行阶段三
    if args.stage in ["3", "all"]:
        sum_f = args.summary or r"D:\桌面文件\虾皮补货文件\11.2\11.2虾皮补货.xlsx"
        pack_tpl_f = args.pack_tpl or r"D:\桌面文件\虾皮补货文件\11.2\11.2虾皮马来.xlsx"
        out_pack = args.output or os.path.join(os.path.dirname(os.path.abspath(pack_tpl_f)), f"{date_str}虾皮{site_code}_已打包.xlsx")

        print(f"[*] 执行阶段三：打包装箱与尺寸重量计算 [站点: {site_code.upper()}]...")
        out_path, box_groups, max_box = engine.process_stage3_packing(
            summary_replenish_path=sum_f,
            shopee_packing_template_path=pack_tpl_f,
            site_code=site_code,
            output_path=out_pack
        )
        print(f"[+] 阶段三完成！共解析 {box_groups} 组装箱，分配至箱子 #1~#{max_box}，生成: {out_path}")

if __name__ == "__main__":
    main()
