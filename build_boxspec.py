# -*- coding: utf-8 -*-
"""把「补货箱规表.xlsx」「官方仓补货模板.xlsx」与「xlsx.full.min.js」内嵌进 index.html —— 后续更新与单文件打包的唯一入口

用法：
    python build_boxspec.py                          # 默认把同目录箱规表、补货模板、SheetJS 全部内嵌进 index.html
    python build_boxspec.py --spec "D:\\桌面文件\\虾皮补货文件\\补货箱规表.xlsx"
    python build_boxspec.py --spec 新箱规.xlsx --template 官方仓补货模板.xlsx --html index.html

做的事情：
  1. 用 core.ReplenishmentEngine 的同一套解析逻辑读取箱规表（保证网页版与桌面版口径一致），同时写入原文件 Base64 支持一键导出
  2. 生成 JSON 数据，替换 index.html 中 /* ==BOXSPEC_START== */ ... /* ==BOXSPEC_END== */ 之间的内容
  3. 将「官方仓补货模板.xlsx」以 Base64 形式内嵌进 /* ==TEMPLATE_START== */ ... /* ==TEMPLATE_END== */，支持免上传默认加载与一键导出
  4. 将「xlsx.full.min.js」内联进 /* ==XLSX_LIB_START== */ ... /* ==XLSX_LIB_END== */，实现彻底的单文件跨机离线分发
  5. 同时把箱规表与模板复制到项目根目录，供 server.py（后端模式）读取

之后刷新页面即可生效；发给同事只需单独发送 index.html 一个文件即可离线使用。
"""

import argparse
import base64
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
TPL_START = "/* ==TEMPLATE_START== */"
TPL_END = "/* ==TEMPLATE_END== */"
XLSX_START = "/* ==XLSX_LIB_START== */"
XLSX_END = "/* ==XLSX_LIB_END== */"


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


def build_block(specs, source_name, digest, raw_b64):
    payload = {
        "version": digest,
        "updated": time.strftime("%Y-%m-%d %H:%M"),
        "source": source_name,
        "count": len(specs),
        "b64": raw_b64,
        "specs": specs,
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # 防止 SKU 中出现 </script> 之类字符提前闭合脚本
    body = body.replace("<", "\\u003c").replace(">", "\\u003e")
    return "const EMBEDDED_BOXSPEC = " + body + ";"


def build_template_block(tpl_path):
    with open(tpl_path, "rb") as fh:
        raw = fh.read()
    digest = hashlib.sha1(raw).hexdigest()[:8]
    payload = {
        "version": digest,
        "updated": time.strftime("%Y-%m-%d %H:%M"),
        "source": os.path.basename(tpl_path),
        "b64": base64.b64encode(raw).decode("ascii"),
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return "const EMBEDDED_TEMPLATE = " + body + ";", digest


def replace_section(html, start_marker, end_marker, content):
    if start_marker not in html or end_marker not in html:
        raise ValueError("找不到标记 %s ... %s" % (start_marker, end_marker))
    head, rest = html.split(start_marker, 1)
    _, tail = rest.split(end_marker, 1)
    return head + start_marker + "\n" + content + "\n" + end_marker + tail


def main():
    ap = argparse.ArgumentParser(description="内嵌箱规表、官方仓补货模板与 SheetJS 到 index.html")
    ap.add_argument("--spec", default=os.path.join(HERE, "补货箱规表.xlsx"), help="箱规表 xlsx 路径")
    ap.add_argument("--template", default=os.path.join(HERE, "官方仓补货模板.xlsx"), help="官方仓补货模板 xlsx 路径")
    ap.add_argument("--xlsx-js", default=os.path.join(HERE, "xlsx.full.min.js"), help="SheetJS 离线脚本路径")
    ap.add_argument("--html", default=os.path.join(HERE, "index.html"), help="目标 index.html")
    ap.add_argument("--no-copy", action="store_true", help="不把箱规表/模板复制到项目根目录")
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
        spec_raw = fh.read()
    digest = hashlib.sha1(spec_raw).hexdigest()[:8]
    spec_b64 = base64.b64encode(spec_raw).decode("ascii")

    block = build_block(specs, os.path.basename(args.spec), digest, spec_b64)

    with open(args.html, "r", encoding="utf-8") as fh:
        html = fh.read()

    html = replace_section(html, START, END, block)

    tpl_digest = None
    if os.path.exists(args.template) and TPL_START in html and TPL_END in html:
        tpl_block, tpl_digest = build_template_block(args.template)
        html = replace_section(html, TPL_START, TPL_END, tpl_block)

    xlsx_embedded = False
    if os.path.exists(args.xlsx_js) and XLSX_START in html and XLSX_END in html:
        with open(args.xlsx_js, "r", encoding="utf-8") as fh:
            xlsx_code = fh.read().replace("</script>", "<\\/script>")
        html = replace_section(html, XLSX_START, XLSX_END, xlsx_code)
        xlsx_embedded = True

    with open(args.html, "w", encoding="utf-8") as fh:
        fh.write(html)

    if not args.no_copy:
        if os.path.abspath(args.spec) != os.path.join(HERE, "补货箱规表.xlsx"):
            shutil.copyfile(args.spec, os.path.join(HERE, "补货箱规表.xlsx"))
        if os.path.exists(args.template) and os.path.abspath(args.template) != os.path.join(HERE, "官方仓补货模板.xlsx"):
            shutil.copyfile(args.template, os.path.join(HERE, "官方仓补货模板.xlsx"))

    print("单文件打包已完成：%s (%.1f KB)" % (args.html, os.path.getsize(args.html) / 1024.0))
    print("  箱规库：%d 个 SKU (版本 %s，来源 %s)" % (len(specs), digest, os.path.basename(args.spec)))
    if tpl_digest:
        print("  补货模板：已内嵌 (版本 %s，来源 %s)" % (tpl_digest, os.path.basename(args.template)))
    if xlsx_embedded:
        print("  SheetJS 引擎：已内联至 HTML（支持彻底单文件离线跨机使用）")
    if missing_qty:
        print("  [注意] 以下 %d 个 SKU 未维护整箱数量，需补量会留空：" % len(missing_qty))
        print("        " + "、".join(missing_qty[:20]) + (" ..." if len(missing_qty) > 20 else ""))
    print("\n完成。刷新网页（或重启 run_web.bat）即可生效。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
