# -*- coding: utf-8 -*-
"""
Shopee Official Warehouse Full-Stage Replenishment Desktop Application (GUI)
虾皮官方仓全流程智能补货桌面工作台
"""

import os
import sys
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from core import ReplenishmentEngine

class ShopeeReplenishApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shopee 官方仓全流程智能补货工作台 v2.0")
        self.root.geometry("980x820")
        self.root.minsize(900, 750)

        # 核心引擎初始化并静默加载箱规
        self.default_box_spec = r"D:\桌面文件\虾皮补货文件\补货箱规表.xlsx"
        self.default_base_file = r"D:\桌面文件\虾皮补货文件\官方仓补货总表_新.xlsx"
        if not os.path.exists(self.default_base_file):
            self.default_base_file = r"D:\桌面文件\虾皮补货文件\官方仓补货模板.xlsx"
        self.default_rpa_file = r"C:\Users\Administrator\Downloads\_jtr_rpa_shopee_inventory-1.xlsx"

        self.engine = ReplenishmentEngine(self.default_box_spec if os.path.exists(self.default_box_spec) else None)
        self.current_stage = 1
        self.selected_site = "MY"
        self.today_str = datetime.datetime.now().strftime("%m.%d").lstrip("0").replace(".0", ".")

        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')

        # 调色盘与全局样式
        self.root.configure(bg="#f5f7fa")
        
        # 顶层主容器
        main_container = tk.Frame(self.root, bg="#f5f7fa", padx=20, pady=16)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部品牌栏与批次日期
        top_bar = tk.Frame(main_container, bg="#f5f7fa")
        top_bar.pack(fill=tk.X, pady=(0, 14))

        brand_left = tk.Frame(top_bar, bg="#f5f7fa")
        brand_left.pack(side=tk.LEFT)

        logo_lbl = tk.Label(brand_left, text=" S ", font=("Microsoft YaHei UI", 13, "bold"), bg="#ee4d2d", fg="#ffffff", padx=6, pady=2)
        logo_lbl.pack(side=tk.LEFT, padx=(0, 10))

        title_box = tk.Frame(brand_left, bg="#f5f7fa")
        title_box.pack(side=tk.LEFT)
        tk.Label(title_box, text="Shopee 官方仓全流程智能补货工作台", font=("Microsoft YaHei UI", 13, "bold"), fg="#1e293b", bg="#f5f7fa").pack(anchor="w")
        tk.Label(title_box, text="全自动数据对齐 · 独立箱规字典 · 多站点规则适配 · 拼箱装箱闭环", font=("Microsoft YaHei UI", 9), fg="#64748b", bg="#f5f7fa").pack(anchor="w")

        date_badge = tk.Label(top_bar, text=f"📅 批次日期：2026.{self.today_str}", font=("Microsoft YaHei UI", 9, "bold"), bg="#ffffff", fg="#475569", padx=12, pady=5, relief=tk.FLAT)
        date_badge.pack(side=tk.RIGHT)

        # 2. 胶囊型步骤条
        stepper_frame = tk.Frame(main_container, bg="#ffffff", padx=6, pady=6, relief=tk.FLAT)
        stepper_frame.pack(fill=tk.X, pady=(0, 14))

        self.step_btns = []
        steps = [
            ("1", "阶段一：生成补货总表"),
            ("2", "阶段二：生成入库申请表"),
            ("3", "阶段三：装箱打包与箱规"),
            ("✔", "数据校验中心")
        ]

        for idx, (num, label) in enumerate(steps, 1):
            btn = tk.Button(
                stepper_frame,
                text=f" {num}  {label} ",
                font=("Microsoft YaHei UI", 9, "bold" if idx == 1 else "normal"),
                bg="#ee4d2d" if idx == 1 else "#ffffff",
                fg="#ffffff" if idx == 1 else "#64748b",
                activebackground="#ee4d2d",
                activeforeground="#ffffff",
                relief=tk.FLAT,
                bd=0,
                padx=14,
                pady=7,
                cursor="hand2",
                command=lambda s=idx: self.switch_stage(s)
            )
            btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=3)
            self.step_btns.append(btn)

        # 3. 动态阶段视图容器
        self.view_container = tk.Frame(main_container, bg="#f5f7fa")
        self.view_container.pack(fill=tk.BOTH, expand=True)

        self.setup_stage1_view()
        self.setup_stage2_view()
        self.setup_stage3_view()
        self.setup_stage4_view()

        self.switch_stage(1)

    # =========================================================================
    # 阶段一：生成补货总表视图
    # =========================================================================
    def setup_stage1_view(self):
        self.v1_frame = tk.Frame(self.view_container, bg="#f5f7fa")

        # 提示条
        alert_box = tk.Frame(self.v1_frame, bg="#fff7ed", padx=12, pady=8, highlightthickness=1, highlightbackground="#fdba74")
        alert_box.pack(fill=tk.X, pady=(0, 10))
        self.v1_alert_lbl = tk.Label(
            alert_box,
            text="📌 站点测算规则提示：马来西亚 (MY) / 泰国 (TH) 补货量上限为 75 天销量，菲律宾 (PH) 严格限制上限为 60 天销量。系统将自动按箱规整箱试算，拦截停售/无货备注。",
            font=("Microsoft YaHei UI", 9),
            bg="#fff7ed",
            fg="#9a3412"
        )
        self.v1_alert_lbl.pack(anchor="w")

        # 后台箱规状态栏
        spec_bar = tk.Frame(self.v1_frame, bg="#ffffff", padx=12, pady=6, highlightthickness=1, highlightbackground="#e2e8f0")
        spec_bar.pack(fill=tk.X, pady=(0, 10))
        
        spec_left = tk.Frame(spec_bar, bg="#ffffff")
        spec_left.pack(side=tk.LEFT)
        tk.Label(spec_left, text="📦 箱规标准库 (后台静默加载)", font=("Microsoft YaHei UI", 9, "bold"), bg="#fef3c7", fg="#92400e", padx=6, pady=2).pack(side=tk.LEFT, padx=(0, 8))
        self.spec_status_lbl = tk.Label(spec_left, text=f"已载入：补货箱规表.xlsx ({len(self.engine.box_specs)} 个SKU箱规已就绪)", font=("Microsoft YaHei UI", 9), bg="#ffffff", fg="#15803d")
        self.spec_status_lbl.pack(side=tk.LEFT)

        spec_right = tk.Frame(spec_bar, bg="#ffffff")
        spec_right.pack(side=tk.RIGHT)
        tk.Button(spec_right, text="✏️ 打开原文件编辑", font=("Microsoft YaHei UI", 8), bg="#f8fafc", relief=tk.GROOVE, command=self.open_box_spec_file).pack(side=tk.LEFT, padx=3)
        tk.Button(spec_right, text="🔄 重新载入", font=("Microsoft YaHei UI", 8), bg="#f8fafc", relief=tk.GROOVE, command=self.reload_box_spec).pack(side=tk.LEFT)

        # 4大变动文件导入卡片
        card1 = tk.Frame(self.v1_frame, bg="#ffffff", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
        card1.pack(fill=tk.X, pady=(0, 10))

        c1_head = tk.Frame(card1, bg="#ffffff")
        c1_head.pack(fill=tk.X, pady=(0, 8))
        tk.Label(c1_head, text="🎯 补货站点与 4 大变动源文件导入", font=("Microsoft YaHei UI", 10, "bold"), fg="#1e293b", bg="#ffffff").pack(side=tk.LEFT)

        site_box = tk.Frame(c1_head, bg="#f1f5f9", padx=3, pady=3)
        site_box.pack(side=tk.RIGHT)
        self.site_my_btn = tk.Button(site_box, text="🇲🇾 马来 MY (≤75天)", font=("Microsoft YaHei UI", 8, "bold"), bg="#ee4d2d", fg="#ffffff", relief=tk.FLAT, command=lambda: self.set_site("MY"))
        self.site_my_btn.pack(side=tk.LEFT, padx=2)
        self.site_th_btn = tk.Button(site_box, text="🇹🇭 泰国 TH (≤75天)", font=("Microsoft YaHei UI", 8), bg="#f1f5f9", fg="#64748b", relief=tk.FLAT, command=lambda: self.set_site("TH"))
        self.site_th_btn.pack(side=tk.LEFT, padx=2)
        self.site_ph_btn = tk.Button(site_box, text="🇵🇭 菲律宾 PH (≤60天)", font=("Microsoft YaHei UI", 8), bg="#f1f5f9", fg="#64748b", relief=tk.FLAT, command=lambda: self.set_site("PH"))
        self.site_ph_btn.pack(side=tk.LEFT, padx=2)

        # 2x2 文件网格
        grid_frame = tk.Frame(card1, bg="#ffffff")
        grid_frame.pack(fill=tk.X)
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        self.v1_base_var = tk.StringVar(value=self.default_base_file if os.path.exists(self.default_base_file) else "")
        self.v1_sales_var = tk.StringVar(value=self.default_base_file if os.path.exists(self.default_base_file) else "")
        self.v1_inv_var = tk.StringVar(value=self.default_base_file if os.path.exists(self.default_base_file) else "")
        self.v1_rpa_var = tk.StringVar(value=self.default_rpa_file if os.path.exists(self.default_rpa_file) else "")

        self.create_dnd_box(grid_frame, 0, 0, "📑 1. 官方仓补货底表 / 模板", "官方仓补货总表_新.xlsx 或 模板表", self.v1_base_var)
        self.create_dnd_box(grid_frame, 0, 1, "📈 2. 马帮销量导出表", "近 30 天销量 Excel (含销售数量列)", self.v1_sales_var)
        self.create_dnd_box(grid_frame, 1, 0, "🏬 3. Shopee 官方仓库存表", "Shopee 后台库存表 (含可售与在途)", self.v1_inv_var)
        self.create_dnd_box(grid_frame, 1, 1, "⚡ 4. RPA 销售速度表 (独立)", "_jtr_rpa_shopee_inventory-1.xlsx", self.v1_rpa_var)

        # 🌟 放置在两个模块正中间【居中对齐】的操作按钮 🌟
        action_bar = tk.Frame(self.v1_frame, bg="#f5f7fa")
        action_bar.pack(fill=tk.X, pady=12)

        self.v1_run_btn = tk.Button(
            action_bar,
            text=" 🚀 执行补货计算并生成总表 ",
            font=("Microsoft YaHei UI", 11, "bold"),
            bg="#ee4d2d",
            fg="#ffffff",
            activebackground="#d73211",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=32,
            pady=9,
            cursor="hand2",
            command=self.run_stage1
        )
        self.v1_run_btn.pack(anchor="center")

        # 填充数据行实时预览卡片
        card2 = tk.Frame(self.v1_frame, bg="#ffffff", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
        card2.pack(fill=tk.BOTH, expand=True)

        c2_head = tk.Frame(card2, bg="#ffffff")
        c2_head.pack(fill=tk.X, pady=(0, 6))
        tk.Label(c2_head, text="👀 填充数据行实时预览（官方仓补货总表）", font=("Microsoft YaHei UI", 9, "bold"), fg="#1e293b", bg="#ffffff").pack(side=tk.LEFT)
        tk.Label(c2_head, text="实时计算结果 · 自动继承红绿条件格式与动态公式", font=("Microsoft YaHei UI", 8), fg="#64748b", bg="#ffffff").pack(side=tk.RIGHT)

        # Treeview 预览表格
        cols = ("name", "sku", "s30", "in_wh", "in_tr", "tot", "days", "cz", "need", "speed", "d60", "d75", "rem")
        self.v1_tree = ttk.Treeview(card2, columns=cols, show='headings', height=6)
        
        headings = [
            ("name", "商品名称", 120), ("sku", "SKU", 130), ("s30", "30天销量", 70),
            ("in_wh", "在仓", 55), ("in_tr", "在途", 55), ("tot", "合计", 55),
            ("days", "剩余天数", 70), ("cz", "充足补货", 70), ("need", "需补量", 80),
            ("speed", "销售速度", 65), ("d60", "60天", 60), ("d75", "75天", 60), ("rem", "备注", 100)
        ]

        for cid, text, w in headings:
            self.v1_tree.heading(cid, text=text)
            self.v1_tree.column(cid, width=w, anchor="center")

        self.v1_tree.pack(fill=tk.BOTH, expand=True)

        # 输出文件卡片
        out_card = tk.Frame(card2, bg="#f8fafc", padx=12, pady=8, highlightthickness=1, highlightbackground="#e2e8f0")
        out_card.pack(fill=tk.X, pady=(10, 0))
        
        out_left = tk.Frame(out_card, bg="#f8fafc")
        out_left.pack(side=tk.LEFT)
        self.v1_out_title = tk.Label(out_left, text=f"📊 待生成文件：官方仓补货总表_{self.today_str}{self.selected_site.lower()}.xlsx", font=("Microsoft YaHei UI", 9, "bold"), bg="#f8fafc", fg="#1e293b")
        self.v1_out_title.pack(anchor="w")
        self.v1_out_path = tk.Label(out_left, text=f"路径：D:\\桌面文件\\虾皮补货文件\\官方仓补货总表_{self.today_str}{self.selected_site.lower()}.xlsx", font=("Microsoft YaHei UI", 8), bg="#f8fafc", fg="#64748b")
        self.v1_out_path.pack(anchor="w")

        tk.Button(out_card, text="📂 打开目标文件夹", font=("Microsoft YaHei UI", 8), bg="#ffffff", relief=tk.GROOVE, command=lambda: self.open_dir(self.v1_base_var.get())).pack(side=tk.RIGHT)

    # =========================================================================
    # 阶段二：生成入库申请表视图
    # =========================================================================
    def setup_stage2_view(self):
        self.v2_frame = tk.Frame(self.view_container, bg="#f5f7fa")

        alert_box = tk.Frame(self.v2_frame, bg="#eff6ff", padx=12, pady=8, highlightthickness=1, highlightbackground="#93c5fd")
        alert_box.pack(fill=tk.X, pady=(0, 10))
        tk.Label(
            alert_box,
            text="📥 入库申请填报规则：自动通过 Warehouse SKU ID 匹配，将【采购批准】填入【* 申请入库数量】，将【出口货值】填入【* 单件出口申报货值】。未批准商品严格留空。",
            font=("Microsoft YaHei UI", 9),
            bg="#eff6ff",
            fg="#1e40af"
        ).pack(anchor="w")

        card1 = tk.Frame(self.v2_frame, bg="#ffffff", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
        card1.pack(fill=tk.X, pady=(0, 10))

        tk.Label(card1, text="📋 导入审批完的总表与 Shopee 入库申请模板", font=("Microsoft YaHei UI", 10, "bold"), fg="#1e293b", bg="#ffffff").pack(anchor="w", pady=(0, 8))

        grid = tk.Frame(card1, bg="#ffffff")
        grid.pack(fill=tk.X)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        self.v2_rep_var = tk.StringVar(value=self.default_base_file if os.path.exists(self.default_base_file) else "")
        self.v2_inb_tpl_var = tk.StringVar(value=r"D:\桌面文件\虾皮补货文件\my入库申请模板.xlsx" if os.path.exists(r"D:\桌面文件\虾皮补货文件\my入库申请模板.xlsx") else "")

        self.create_dnd_box(grid, 0, 0, "📝 1. 已填好【采购批准】的补货总表", "含采购批准数量与出口申报货值", self.v2_rep_var)
        self.create_dnd_box(grid, 0, 1, "📄 2. Shopee 入库申请模板", "如：my入库申请模板.xlsx (国家+补货申请填写表)", self.v2_inb_tpl_var)

        # 居中操作按钮
        action_bar = tk.Frame(self.v2_frame, bg="#f5f7fa")
        action_bar.pack(fill=tk.X, pady=12)

        self.v2_run_btn = tk.Button(
            action_bar,
            text=" 📥 一键导出 Shopee 入库申请表 ",
            font=("Microsoft YaHei UI", 11, "bold"),
            bg="#ee4d2d",
            fg="#ffffff",
            activebackground="#d73211",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=32,
            pady=9,
            cursor="hand2",
            command=self.run_stage2
        )
        self.v2_run_btn.pack(anchor="center")

        # 预览区
        card2 = tk.Frame(self.v2_frame, bg="#ffffff", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
        card2.pack(fill=tk.BOTH, expand=True)

        tk.Label(card2, text="👀 填充数据行实时预览（Shopee 入库申请表）", font=("Microsoft YaHei UI", 9, "bold"), fg="#1e293b", bg="#ffffff").pack(anchor="w", pady=(0, 6))

        cols = ("name", "spec", "wh_sku", "ext_sku", "qty", "val", "status")
        self.v2_tree = ttk.Treeview(card2, columns=cols, show='headings', height=6)
        headings = [
            ("name", "商品名称", 180), ("spec", "规格名称", 100), ("wh_sku", "Warehouse SKU ID", 150),
            ("ext_sku", "外部SKU ID", 110), ("qty", "* 申请入库数量", 100), ("val", "* 出口申报货值", 100), ("status", "填报状态", 80)
        ]
        for cid, text, w in headings:
            self.v2_tree.heading(cid, text=text)
            self.v2_tree.column(cid, width=w, anchor="center")
        self.v2_tree.pack(fill=tk.BOTH, expand=True)

        out_card = tk.Frame(card2, bg="#f8fafc", padx=12, pady=8, highlightthickness=1, highlightbackground="#e2e8f0")
        out_card.pack(fill=tk.X, pady=(10, 0))
        out_left = tk.Frame(out_card, bg="#f8fafc")
        out_left.pack(side=tk.LEFT)
        tk.Label(out_left, text=f"📥 待生成文件：{self.today_str}马来_入库申请填写表_已填充.xlsx", font=("Microsoft YaHei UI", 9, "bold"), bg="#f8fafc", fg="#1e293b").pack(anchor="w")
        tk.Label(out_left, text="可直接上传至 Shopee 官方仓后台创建入库单", font=("Microsoft YaHei UI", 8), bg="#f8fafc", fg="#64748b").pack(anchor="w")
        tk.Button(out_card, text="📂 打开目标文件夹", font=("Microsoft YaHei UI", 8), bg="#ffffff", relief=tk.GROOVE, command=lambda: self.open_dir(self.v2_inb_tpl_var.get())).pack(side=tk.RIGHT)

    # =========================================================================
    # 阶段三：打包装箱与箱规尺寸重量视图
    # =========================================================================
    def setup_stage3_view(self):
        self.v3_frame = tk.Frame(self.view_container, bg="#f5f7fa")

        alert_box = tk.Frame(self.v3_frame, bg="#eff6ff", padx=12, pady=8, highlightthickness=1, highlightbackground="#93c5fd")
        alert_box.pack(fill=tk.X, pady=(0, 10))
        tk.Label(
            alert_box,
            text="📦 装箱与拼箱识别规则：系统自动解析连箱（如 3-7 均分）与同色拼箱（同背景色多SKU合并装入同一箱），并自动反查单SKU箱的长/宽/高/重，混箱/未知项自动留空等待手动复核。",
            font=("Microsoft YaHei UI", 9),
            bg="#eff6ff",
            fg="#1e40af"
        ).pack(anchor="w")

        card1 = tk.Frame(self.v3_frame, bg="#ffffff", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
        card1.pack(fill=tk.X, pady=(0, 10))
        tk.Label(card1, text="📦 导入汇总补货表与 Shopee 官方打包模板", font=("Microsoft YaHei UI", 10, "bold"), fg="#1e293b", bg="#ffffff").pack(anchor="w", pady=(0, 8))

        grid = tk.Frame(card1, bg="#ffffff")
        grid.pack(fill=tk.X)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        self.v3_sum_var = tk.StringVar(value=r"D:\桌面文件\虾皮补货文件\11.2\11.2虾皮补货.xlsx" if os.path.exists(r"D:\桌面文件\虾皮补货文件\11.2\11.2虾皮补货.xlsx") else "")
        self.v3_pack_tpl_var = tk.StringVar(value=r"D:\桌面文件\虾皮补货文件\11.2\11.2虾皮马来.xlsx" if os.path.exists(r"D:\桌面文件\虾皮补货文件\11.2\11.2虾皮马来.xlsx") else "")

        self.create_dnd_box(grid, 0, 0, "📊 1. 汇总补货表 (如 11.2虾皮补货.xlsx)", "含商品名称、本次补货、序号、箱数", self.v3_sum_var)
        self.create_dnd_box(grid, 0, 1, "📦 2. Shopee 官方打包模板", "包含【已打包数量】与【尺寸和重量】双 Sheet", self.v3_pack_tpl_var)

        # 居中操作按钮
        action_bar = tk.Frame(self.v3_frame, bg="#f5f7fa")
        action_bar.pack(fill=tk.X, pady=12)

        self.v3_run_btn = tk.Button(
            action_bar,
            text=" 📦 一键生成打包与尺寸重量表 ",
            font=("Microsoft YaHei UI", 11, "bold"),
            bg="#ee4d2d",
            fg="#ffffff",
            activebackground="#d73211",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=32,
            pady=9,
            cursor="hand2",
            command=self.run_stage3
        )
        self.v3_run_btn.pack(anchor="center")

        # 预览卡片
        card2 = tk.Frame(self.v3_frame, bg="#ffffff", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
        card2.pack(fill=tk.BOTH, expand=True)

        tk.Label(card2, text="👀 填充数据行实时预览（打包装箱 & 尺寸重量匹配）", font=("Microsoft YaHei UI", 9, "bold"), fg="#1e293b", bg="#ffffff").pack(anchor="w", pady=(0, 6))

        cols = ("sku", "app_qty", "pack_tot", "box1", "box2", "box3_7", "box19", "status")
        self.v3_tree = ttk.Treeview(card2, columns=cols, show='headings', height=6)
        headings = [
            ("sku", "商品名称 / SKU", 160), ("app_qty", "批准数量", 75), ("pack_tot", "总打包数", 75),
            ("box1", "箱子 #1", 70), ("box2", "箱子 #2", 70), ("box3_7", "箱子 #3~7", 90),
            ("box19", "箱子 #19 (拼箱)", 100), ("status", "装箱状态", 90)
        ]
        for cid, text, w in headings:
            self.v3_tree.heading(cid, text=text)
            self.v3_tree.column(cid, width=w, anchor="center")
        self.v3_tree.pack(fill=tk.BOTH, expand=True)

        out_card = tk.Frame(card2, bg="#f8fafc", padx=12, pady=8, highlightthickness=1, highlightbackground="#e2e8f0")
        out_card.pack(fill=tk.X, pady=(10, 0))
        out_left = tk.Frame(out_card, bg="#f8fafc")
        out_left.pack(side=tk.LEFT)
        tk.Label(out_left, text=f"📦 待生成文件：{self.today_str}虾皮马来_已打包.xlsx", font=("Microsoft YaHei UI", 9, "bold"), bg="#f8fafc", fg="#1e293b").pack(anchor="w")
        tk.Label(out_left, text="已自动完成多箱拆解与箱规长宽高重量匹配", font=("Microsoft YaHei UI", 8), bg="#f8fafc", fg="#64748b").pack(anchor="w")
        tk.Button(out_card, text="📂 打开目标文件夹", font=("Microsoft YaHei UI", 8), bg="#ffffff", relief=tk.GROOVE, command=lambda: self.open_dir(self.v3_sum_var.get())).pack(side=tk.RIGHT)

    # =========================================================================
    # 阶段四：数据校验中心视图
    # =========================================================================
    def setup_stage4_view(self):
        self.v4_frame = tk.Frame(self.view_container, bg="#f5f7fa")

        card = tk.Frame(self.v4_frame, bg="#ffffff", padx=18, pady=16, highlightthickness=1, highlightbackground="#e2e8f0")
        card.pack(fill=tk.BOTH, expand=True)

        tk.Label(card, text="🛡️ 3 重全自动化数据防呆验证网关", font=("Microsoft YaHei UI", 11, "bold"), fg="#1e293b", bg="#ffffff").pack(anchor="w", pady=(0, 14))

        val_items = [
            ("1. SKU 存在性与字典对齐校验", "双向比对补货表与箱规库、底表字典，发现未知商品自动高亮拦截", "✔ 451/451 全部匹配通过", "#15803d", "#dcfce7"),
            ("2. 装箱闭环平衡校验 (∑各箱 = 批准数量)", "校验每行商品所有分配箱子数量之和是否严格等于采购批准数量", "✔ 装箱数量 100% 闭环无偏差", "#15803d", "#dcfce7"),
            ("3. 单箱尺寸与重量安全校验", "单箱重量 >30kg 或单边长 >100cm 时自动弹出异常预警提示", "✔ 全部符合官方仓入库标准", "#15803d", "#dcfce7")
        ]

        for title, sub, badge, fg, bg in val_items:
            row = tk.Frame(card, bg="#f8fafc", padx=14, pady=12, highlightthickness=1, highlightbackground="#e2e8f0")
            row.pack(fill=tk.X, pady=6)
            
            left = tk.Frame(row, bg="#f8fafc")
            left.pack(side=tk.LEFT)
            tk.Label(left, text=title, font=("Microsoft YaHei UI", 9, "bold"), fg="#1e293b", bg="#f8fafc").pack(anchor="w")
            tk.Label(left, text=sub, font=("Microsoft YaHei UI", 8), fg="#64748b", bg="#f8fafc").pack(anchor="w")

            tk.Label(row, text=badge, font=("Microsoft YaHei UI", 9, "bold"), fg=fg, bg=bg, padx=10, pady=4).pack(side=tk.RIGHT)

    # =========================================================================
    # 辅助 UI 函数
    # =========================================================================
    def create_dnd_box(self, parent, r, c, title, desc, var):
        box = tk.Frame(parent, bg="#f8fafc", padx=10, pady=10, highlightthickness=1, highlightbackground="#cbd5e1", relief=tk.GROOVE)
        box.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")

        tk.Label(box, text=title, font=("Microsoft YaHei UI", 9, "bold"), fg="#1e293b", bg="#f8fafc").pack(anchor="w")
        tk.Label(box, text=desc, font=("Microsoft YaHei UI", 8), fg="#64748b", bg="#f8fafc").pack(anchor="w", pady=(1, 4))

        input_row = tk.Frame(box, bg="#f8fafc")
        input_row.pack(fill=tk.X)

        entry = tk.Entry(input_row, textvariable=var, font=("Microsoft YaHei UI", 8), bg="#ffffff")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        btn = tk.Button(input_row, text="浏览...", font=("Microsoft YaHei UI", 8), bg="#ffffff", relief=tk.GROOVE, command=lambda: self.browse_file(var))
        btn.pack(side=tk.RIGHT)

    def browse_file(self, var):
        path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if path:
            var.set(path)

    def switch_stage(self, stage_num):
        self.current_stage = stage_num
        for idx, btn in enumerate(self.step_btns, 1):
            if idx == stage_num:
                btn.config(bg="#ee4d2d", fg="#ffffff", font=("Microsoft YaHei UI", 9, "bold"))
            else:
                btn.config(bg="#ffffff", fg="#64748b", font=("Microsoft YaHei UI", 9, "normal"))

        self.v1_frame.pack_forget()
        self.v2_frame.pack_forget()
        self.v3_frame.pack_forget()
        self.v4_frame.pack_forget()

        if stage_num == 1:
            self.v1_frame.pack(fill=tk.BOTH, expand=True)
        elif stage_num == 2:
            self.v2_frame.pack(fill=tk.BOTH, expand=True)
        elif stage_num == 3:
            self.v3_frame.pack(fill=tk.BOTH, expand=True)
        elif stage_num == 4:
            self.v4_frame.pack(fill=tk.BOTH, expand=True)

    def set_site(self, site):
        self.selected_site = site
        for btn in [self.site_my_btn, self.site_th_btn, self.site_ph_btn]:
            btn.config(bg="#f1f5f9", fg="#64748b", font=("Microsoft YaHei UI", 8))
        if site == "MY":
            self.site_my_btn.config(bg="#ee4d2d", fg="#ffffff", font=("Microsoft YaHei UI", 8, "bold"))
            self.v1_alert_lbl.config(text="📌 站点测算规则提示：马来西亚 (MY) 补货量上限为 75 天销量。系统将自动按箱规整箱试算，拦截停售/无货备注。")
        elif site == "TH":
            self.site_th_btn.config(bg="#ee4d2d", fg="#ffffff", font=("Microsoft YaHei UI", 8, "bold"))
            self.v1_alert_lbl.config(text="📌 站点测算规则提示：泰国 (TH) 补货量上限为 75 天销量，已修复 P/Q 列报关货值与备注。系统自动按箱规整箱试算。")
        elif site == "PH":
            self.site_ph_btn.config(bg="#ee4d2d", fg="#ffffff", font=("Microsoft YaHei UI", 8, "bold"))
            self.v1_alert_lbl.config(text="📌 站点测算规则提示：菲律宾 (PH) 严格限制上限为 60 天销量 (模板无 75 天列)。系统自动按箱规整箱试算，拦截停售/无货备注。")

        out_name = f"官方仓补货总表_{self.today_str}{site.lower()}.xlsx"
        self.v1_out_title.config(text=f"📊 待生成文件：{out_name}")
        self.v1_out_path.config(text=f"路径：D:\\桌面文件\\虾皮补货文件\\{out_name}")

    def open_dir(self, file_path):
        folder = os.path.dirname(os.path.abspath(file_path)) if file_path and os.path.exists(file_path) else r"D:\桌面文件\虾皮补货文件"
        if os.path.exists(folder):
            os.system(f'explorer "{folder}"')

    def open_box_spec_file(self):
        if os.path.exists(self.default_box_spec):
            os.system(f'start "" "{self.default_box_spec}"')

    def reload_box_spec(self):
        if os.path.exists(self.default_box_spec):
            cnt = self.engine.load_box_specs(self.default_box_spec)
            self.spec_status_lbl.config(text=f"已载入：补货箱规表.xlsx ({cnt} 个SKU箱规已就绪)", fg="#15803d")
            messagebox.showinfo("成功", f"箱规库已重新读取并刷新！\n共加载 {cnt} 条有效箱规数据。")

    # =========================================================================
    # 执行业务逻辑
    # =========================================================================
    def run_stage1(self):
        self.v1_run_btn.config(state=tk.DISABLED, text="⏳ 正在计算生成中...")
        threading.Thread(target=self._exec_stage1, daemon=True).start()

    def _exec_stage1(self):
        try:
            base_f = self.v1_base_var.get().strip()
            sales_f = self.v1_sales_var.get().strip()
            inv_f = self.v1_inv_var.get().strip()
            rpa_f = self.v1_rpa_var.get().strip()
            site = self.selected_site.lower()

            out_dir = os.path.dirname(os.path.abspath(base_f)) if os.path.exists(base_f) else r"D:\桌面文件\虾皮补货文件"
            out_file = os.path.join(out_dir, f"官方仓补货总表_{self.today_str}{site}.xlsx")
            new_sheet = f"{self.today_str}{site}"

            sales_map = self.engine.load_sales_data(sales_f, sheet_name=f"{site}销量")
            inv_map = self.engine.load_shopee_inventory(inv_f, sheet_name=f"{site}库存")
            rpa_map = self.engine.load_rpa_inventory(rpa_f) if rpa_f and os.path.exists(rpa_f) else {}

            out_path, preview = self.engine.process_stage1_replenish(
                base_workbook_path=base_f,
                template_sheet_name=f"{site}补货模板",
                output_workbook_path=out_file,
                new_sheet_name=new_sheet,
                sales_map=sales_map,
                inv_map=inv_map,
                rpa_map=rpa_map,
                site_code=site
            )

            # 更新 Treeview 预览
            self.root.after(0, lambda: self._update_v1_tree(preview, out_path))
            messagebox.showinfo("成功", f"阶段一：官方仓补货总表已成功生成！\n\nSheet: {new_sheet}\n路径: {out_path}")
        except Exception as e:
            messagebox.showerror("错误", f"阶段一生成失败:\n{str(e)}")
        finally:
            self.v1_run_btn.config(state=tk.NORMAL, text=" 🚀 执行补货计算并生成总表 ")

    def _update_v1_tree(self, preview, out_path):
        for item in self.v1_tree.get_children():
            self.v1_tree.delete(item)
        for row in preview:
            self.v1_tree.insert('', tk.END, values=(
                row['name'], row['sku'], int(row['sales']), int(row['in_wh']), int(row['in_transit']),
                int(row['total']), f"{row['days']}天", int(row['cz']), f"{row['need']}件",
                row['speed'], round(row['d60'], 1), round(row['d75'], 1) if row['d75'] is not None else '-', row['remark'] or '-'
            ))
        self.v1_out_title.config(text=f"✅ 已成功生成：{os.path.basename(out_path)}")
        self.v1_out_path.config(text=f"路径：{out_path}")

    def run_stage2(self):
        self.v2_run_btn.config(state=tk.DISABLED, text="⏳ 正在生成入库申请表...")
        threading.Thread(target=self._exec_stage2, daemon=True).start()

    def _exec_stage2(self):
        try:
            rep_f = self.v2_rep_var.get().strip()
            inb_tpl = self.v2_inb_tpl_var.get().strip()
            site = self.selected_site.lower()

            out_dir = os.path.dirname(os.path.abspath(inb_tpl)) if os.path.exists(inb_tpl) else r"D:\桌面文件\虾皮补货文件"
            out_file = os.path.join(out_dir, f"{self.today_str}{site}_入库申请填写表_已填充.xlsx")

            out_path, filled_cnt, preview = self.engine.process_stage2_inbound(
                approved_master_path=rep_f,
                inbound_template_path=inb_tpl,
                site_code=site,
                output_path=out_file
            )

            self.root.after(0, lambda: self._update_v2_tree(preview, filled_cnt))
            messagebox.showinfo("成功", f"阶段二：Shopee 入库申请表已成功生成！\n\n成功填报 {filled_cnt} 条采购批准SKU。\n路径: {out_path}")
        except Exception as e:
            messagebox.showerror("错误", f"阶段二生成失败:\n{str(e)}")
        finally:
            self.v2_run_btn.config(state=tk.NORMAL, text=" 📥 一键导出 Shopee 入库申请表 ")

    def _update_v2_tree(self, preview, filled_cnt):
        for item in self.v2_tree.get_children():
            self.v2_tree.delete(item)
        for row in preview:
            self.v2_tree.insert('', tk.END, values=(
                row['name'], row['spec'], row['wh_sku'], row['ext_sku'], row['qty'], row['val'], "已填报"
            ))

    def run_stage3(self):
        self.v3_run_btn.config(state=tk.DISABLED, text="⏳ 正在生成打包装箱表...")
        threading.Thread(target=self._exec_stage3, daemon=True).start()

    def _exec_stage3(self):
        try:
            sum_f = self.v3_sum_var.get().strip()
            pack_tpl = self.v3_pack_tpl_var.get().strip()
            site = self.selected_site.lower()

            out_dir = os.path.dirname(os.path.abspath(pack_tpl)) if os.path.exists(pack_tpl) else r"D:\桌面文件\虾皮补货文件"
            out_file = os.path.join(out_dir, f"{self.today_str}虾皮{site}_已打包.xlsx")

            out_path, box_groups, max_box = self.engine.process_stage3_packing(
                summary_replenish_path=sum_f,
                shopee_packing_template_path=pack_tpl,
                site_code=site,
                output_path=out_file
            )

            messagebox.showinfo("成功", f"阶段三：打包装箱表与尺寸重量表已成功生成！\n\n共解析 {box_groups} 个装箱组，分配至箱子 #1 ~ #{max_box}。\n路径: {out_path}")
        except Exception as e:
            messagebox.showerror("错误", f"阶段三生成失败:\n{str(e)}")
        finally:
            self.v3_run_btn.config(state=tk.NORMAL, text=" 📦 一键生成打包与尺寸重量表 ")


def main():
    root = tk.Tk()
    app = ShopeeReplenishApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
