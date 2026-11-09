# -*- coding: utf-8 -*-
"""双端引擎一致性校验：Python core.py  vs  index.html 内置 JS Engine

保证「桌面版/后端」与「纯浏览器版」算出的数字完全一致。
任何一端改了算法，跑一次这个脚本就能发现分歧。

用法：
    python tests/compare_engines.py          # 需要本机有 node
"""

import json
import os
import shutil
import subprocess
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BUILD = os.path.join(HERE, ".build")
FIX = os.path.join(HERE, "fixtures")

sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

from core import ReplenishmentEngine  # noqa: E402
import make_fixtures  # noqa: E402

SITES = ["my", "th", "ph"]
PREVIEW_FIELDS = ["sku", "sales", "in_wh", "in_transit", "total", "days", "cz", "need",
                  "status", "speed", "d60", "d75"]


def norm(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return round(float(v), 4)
    return v


# ---------------------------------------------------------------- Python 侧
def run_python():
    eng = ReplenishmentEngine(os.path.join(FIX, "补货箱规表.xlsx"))
    sales = eng.load_sales_data(os.path.join(FIX, "我的马帮销量导出表.xlsx"), sheet_name="my销量")
    inv = eng.load_shopee_inventory(os.path.join(FIX, "my库存导出表.xlsx"), sheet_name="my库存")
    rpa = eng.load_rpa_inventory(os.path.join(FIX, "my销售速度.xlsx"))

    out = {"stage1": {}}
    for site in SITES:
        path = os.path.join(FIX, "py_%s.xlsx" % site)
        _, preview = eng.process_stage1_replenish(
            base_workbook_path=os.path.join(ROOT, "官方仓补货模板.xlsx"),
            template_sheet_name="%s补货模板" % site,
            output_workbook_path=path,
            new_sheet_name="11.7%s" % site,
            sales_map=sales, inv_map=inv, rpa_map=rpa, site_code=site)
        out["stage1"][site] = {"preview": preview, "file": path}

    approved = os.path.join(FIX, "已审批总表.xlsx")
    wb = openpyxl.load_workbook(out["stage1"]["my"]["file"])
    ws = wb["11.7my"]
    for r, v in [(2, 20), (3, 40), (4, 100)]:
        ws.cell(r, 12).value = v
    wb.save(approved)

    p2 = os.path.join(FIX, "py_s2.xlsx")
    _, filled, prev2 = eng.process_stage2_inbound(
        approved, os.path.join(FIX, "my入库申请模板.xlsx"), "my", p2)

    p3 = os.path.join(FIX, "py_s3.xlsx")
    _, groups, maxbox, prev3 = eng.process_stage3_packing(
        os.path.join(FIX, "汇总补货表.xlsx"), os.path.join(FIX, "打包模板.xlsx"), "my", p3,
        return_preview=True)

    val = eng.validate_system(out["stage1"]["my"]["file"], p3, "my")

    out["stage2"] = {"preview": prev2, "filled": filled, "file": p2}
    out["stage3"] = {"preview": prev3, "groups": groups, "maxbox": maxbox, "file": p3}
    out["validate"] = val
    return out


# ---------------------------------------------------------------- JS 侧
def extract_engine():
    html = open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()
    a = html.index("// ==ENGINE_START==")
    b = html.index("// ==ENGINE_END==") + len("// ==ENGINE_END==")
    xlsx_path = os.path.join(ROOT, "xlsx.full.min.js").replace("\\", "/")
    js = ("const XLSX = require(%r);\n" % xlsx_path) + html[a:b] + "\nmodule.exports = Engine;\n"
    os.makedirs(BUILD, exist_ok=True)
    path = os.path.join(BUILD, "engine_node.js")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(js)
    return path


NODE_RUNNER = r"""
const fs = require('fs');
const path = require('path');
const Engine = require(process.argv[2]);
const FIX = process.argv[3];
const ROOT = process.argv[4];
const OUTDIR = process.argv[5];
const rd = p => fs.readFileSync(p);

const specs = Engine.loadBoxSpecs(rd(path.join(FIX, '补货箱规表.xlsx')));
const out = { stage1: {} };
for (const site of ['my', 'th', 'ph']) {
  const res = Engine.stage1({
    baseBuf: rd(path.join(ROOT, '官方仓补货模板.xlsx')),
    salesBuf: rd(path.join(FIX, '我的马帮销量导出表.xlsx')),
    invBuf: rd(path.join(FIX, 'my库存导出表.xlsx')),
    rpaBuf: rd(path.join(FIX, 'my销售速度.xlsx')),
    site, date: '11.7', boxSpecs: specs
  });
  const f = path.join(OUTDIR, 'js_' + site + '.xlsx');
  fs.writeFileSync(f, Buffer.from(res.outBuf));
  out.stage1[site] = { preview: res.preview, file: f };
}

const r2 = Engine.stage2({
  repBuf: rd(path.join(FIX, '已审批总表.xlsx')),
  inbBuf: rd(path.join(FIX, 'my入库申请模板.xlsx')),
  site: 'my', date: '11.7'
});
const f2 = path.join(OUTDIR, 'js_s2.xlsx');
fs.writeFileSync(f2, Buffer.from(r2.outBuf));

const r3 = Engine.stage3({
  summaryBuf: rd(path.join(FIX, '汇总补货表.xlsx')),
  packBuf: rd(path.join(FIX, '打包模板.xlsx')),
  site: 'my', date: '11.7', boxSpecs: specs
});
const f3 = path.join(OUTDIR, 'js_s3.xlsx');
fs.writeFileSync(f3, Buffer.from(r3.outBuf));

const val = Engine.validate({
  masterBuf: rd(out.stage1.my.file), packingBuf: fs.readFileSync(f3),
  site: 'my', boxSpecs: specs
});

out.stage2 = { preview: r2.preview, filled: r2.filled, file: f2 };
out.stage3 = { preview: r3.preview, groups: r3.groups, maxbox: r3.maxBox, file: f3 };
out.validate = val;
fs.writeFileSync(path.join(OUTDIR, 'js.json'), JSON.stringify(out));
console.log('node side done');
"""


def run_js(engine_path):
    os.makedirs(BUILD, exist_ok=True)
    runner = os.path.join(BUILD, "run_js.js")
    with open(runner, "w", encoding="utf-8") as fh:
        fh.write(NODE_RUNNER)
    res = subprocess.run(
        ["node", runner, engine_path, FIX, ROOT, BUILD],
        capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError("node 执行失败：\n%s\n%s" % (res.stdout, res.stderr))
    with open(os.path.join(BUILD, "js.json"), encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------- 比对
def compare_preview(py_rows, js_rows, label, problems):
    if len(py_rows) != len(js_rows):
        problems.append("%s 行数不一致 py=%d js=%d" % (label, len(py_rows), len(js_rows)))
        return
    for a, b in zip(py_rows, js_rows):
        for f in PREVIEW_FIELDS:
            if norm(a.get(f)) != norm(b.get(f)):
                problems.append("%s row%s %s [%s] py=%r js=%r"
                                % (label, a.get("row", "?"), a.get("sku"), f, a.get(f), b.get(f)))


def compare_workbook(p1, p2, label, problems):
    w1 = openpyxl.load_workbook(p1)
    w2 = openpyxl.load_workbook(p2)
    if w1.sheetnames != w2.sheetnames:
        problems.append("%s 工作表不一致 py=%s js=%s" % (label, w1.sheetnames, w2.sheetnames))
        return
    for name in w1.sheetnames:
        a, b = w1[name], w2[name]
        for r in range(1, max(a.max_row, b.max_row) + 1):
            for c in range(1, max(a.max_column, b.max_column) + 1):
                x, y = a.cell(r, c).value, b.cell(r, c).value
                if x == y or (x in (None, "") and y in (None, "")):
                    continue
                problems.append("%s/%s 单元格 %s%d py=%r js=%r"
                                % (label, name, openpyxl.utils.get_column_letter(c), r, x, y))


def main():
    make_fixtures.build()
    if not shutil.which("node"):
        print("[跳过] 未检测到 node，无法执行 JS 引擎比对")
        return 2

    py = run_python()
    js = run_js(extract_engine())

    problems = []
    for site in SITES:
        compare_preview(py["stage1"][site]["preview"], js["stage1"][site]["preview"],
                        "阶段一[%s]" % site, problems)
        compare_workbook(py["stage1"][site]["file"], js["stage1"][site]["file"],
                         "阶段一[%s]" % site, problems)

    compare_preview(py["stage2"]["preview"], js["stage2"]["preview"], "阶段二", problems)
    compare_workbook(py["stage2"]["file"], js["stage2"]["file"], "阶段二", problems)
    compare_preview(py["stage3"]["preview"], js["stage3"]["preview"], "阶段三", problems)
    compare_workbook(py["stage3"]["file"], js["stage3"]["file"], "阶段三", problems)

    for key in ("filled", "groups", "maxbox"):
        a = py["stage2"].get(key, py["stage3"].get(key))
        b = js["stage2"].get(key, js["stage3"].get(key))
        if a != b:
            problems.append("%s 不一致 py=%r js=%r" % (key, a, b))

    for key, item in py["validate"].items():
        other = js["validate"].get(key, {})
        if item["pass"] != other.get("pass") or item["msg"] != other.get("msg"):
            problems.append("校验[%s] 不一致 py=%s js=%s" % (key, item, other))

    checks = 3 * 3 + 2 + 3
    if problems:
        print("\n❌ 发现 %d 处不一致：" % len(problems))
        for p in problems[:40]:
            print("   -", p)
        return 1
    print("\n✅ 通过：三个站点阶段一 + 阶段二 + 阶段三 + 三重校验，Python 与 JS 结果完全一致")
    print("   比对范围：%d 个站点 × 全部预览字段 + 生成 Excel 逐单元格" % len(SITES))
    return 0


if __name__ == "__main__":
    sys.exit(main())
