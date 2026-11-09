# -*- coding: utf-8 -*-
"""把「补货箱规表.xlsx」内嵌进 index.html —— 后续更新箱规的唯一入口

用法：
    python build_boxspec.py                          # 默认读同目录 补货箱规表.xlsx
    python build_boxspec.py --spec "D:\\桌面文件\\虾皮补货文件\\补货箱规表.xlsx"
    python build_boxspec.py --spec 新箱规.xlsx --html index.html

做的事情：
  1. 用 core.ReplenishmentEngine 的同一套解析逻辑读取箱规表（保证网页版与桌面版口径一致）
  2. 生成 JSON 数据，替换 index.html 中 /* ==BOXSPEC_START== */ ... /* ==BOXSPEC_END== */ 之间的内容
  3. 写入版本号（文件 SHA1 前 8 位）、条数、生成时间、来源文件名
  4. 同时把箱规表复制到项目根目录，供 server.py（后端模式）读取

之后刷新页面即可生效；无需改动其它代码。
"""

import argparse
import hashlib
import json
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from core import ReplenishmentEngine  # noqa: E402

START = "/* ==BOXSPEC_START== */"
END = "/* ==BOXSPEC_END== */"


def parse_specs(path):
    engine = ReplenishmentEngine(path)
    specs = {}
    for sku, info in engine.box_specs.items():
        if not sku:
            continue
        specs[sku] = {
            "qty": info.get("qty"),
            "length": info.get("length"),
            "width": info.get("width"),
            "height": info.get("height"),
            "weight": info.get("weight"),
            "note": info.get("note"),
        }
    return specs


def build_block(specs, source_name, digest):
    payload = {
        "version": digest,
        "updated": time.strftime("%Y-%m-%d %H:%M"),
        "source": source_name,
        "count": len(specs),
        "specs": specs,
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # 防止 SKU 中出现 </script> 之类字符提前闭合脚本
    body = body.replace("<", "\\u003c").replace(">", "\\u003e")
    return "const EMBEDDED_BOXSPEC = " + body + ";"


def main():
    ap = argparse.ArgumentParser(description="内嵌箱规表到 index.html")
    ap.add_argument("--spec", default=os.path.join(HERE, "补货箱规表.xlsx"), help="箱规表 xlsx 路径")
    ap.add_argument("--html", default=os.path.join(HERE, "index.html"), help="目标 index.html")
    ap.add_argument("--no-copy", action="store_true", help="不把箱规表复制到项目根目录")
    args = ap.parse_args()

    if not os.path.exists(args.spec):
        print("[错误] 找不到箱规表：%s" % args.spec)
        print("       请从 D:\\桌面文件\\虾皮补货文件\\补货箱规表.xlsx 复制到项目目录，或用 --spec 指定路径")
        return 1

    specs = parse_specs(args.spec)
    if not specs:
        print("[错误] 未能从 %s 解析出任何箱规。" % os.path.basename(args.spec))
        print("       要求：第一列为 SKU；表头包含 数量(或箱规) / 长 / 宽 / 高 / 重 列。")
        return 1

    missing_qty = [k for k, v in specs.items() if not v.get("qty")]
    with open(args.spec, "rb") as fh:
        digest = hashlib.sha1(fh.read()).hexdigest()[:8]

    block = build_block(specs, os.path.basename(args.spec), digest)

    with open(args.html, "r", encoding="utf-8") as fh:
        html = fh.read()
    if START not in html or END not in html:
        print("[错误] %s 中找不到箱规标记 %s ... %s" % (args.html, START, END))
        return 1
    head, rest = html.split(START, 1)
    _, tail = rest.split(END, 1)
    with open(args.html, "w", encoding="utf-8") as fh:
        fh.write(head + START + "\n" + block + "\n" + END + tail)

    if not args.no_copy and os.path.abspath(args.spec) != os.path.join(HERE, "补货箱规表.xlsx"):
        shutil.copyfile(args.spec, os.path.join(HERE, "补货箱规表.xlsx"))

    print("箱规已内嵌：%s" % args.html)
    print("  条数：%d 个 SKU" % len(specs))
    print("  版本：%s" % digest)
    print("  来源：%s" % os.path.basename(args.spec))
    if missing_qty:
        print("  [注意] 以下 %d 个 SKU 未维护整箱数量，需补量会留空：" % len(missing_qty))
        print("        " + "、".join(missing_qty[:20]) + (" ..." if len(missing_qty) > 20 else ""))
    print("\n完成。刷新网页（或重启 run_web.bat）即可使用新箱规。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
