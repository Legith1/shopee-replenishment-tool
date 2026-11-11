# -*- coding: utf-8 -*-
"""Shopee 官方仓智能补货工具 —— 本地服务端

只有标准库依赖（除 openpyxl 外无第三方 Web 框架），复用 core.py 的业务引擎，
因此网页版与桌面版走同一套算法，且 openpyxl 会完整保留模板的公式与条件格式。

启动：
    python server.py            # 自动选端口并打开浏览器
    python server.py --port 8765 --no-browser

接口（全部返回 JSON，文件通过 base64 传入/下载器取出）：
    GET  /                       -> index.html
    GET  /api/health             -> 服务与箱规库状态
    POST /api/stage1             -> 生成补货总表
    POST /api/stage2             -> 生成入库申请表
    POST /api/stage3             -> 生成打包与尺寸重量表
    POST /api/validate           -> 三重防呆校验
    POST /api/boxspec            -> 上传/替换箱规表（后续更新入口）
    GET  /api/download/<id>      -> 下载生成的文件
"""

import argparse
import base64
import hashlib
import importlib
import json
import os
import re
import socket
import sys
import time
import traceback
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import core  # noqa: E402

APP_VERSION = "2.1.0"
RUNTIME_DIR = os.path.join(HERE, "_runtime")
RUNS_DIR = os.path.join(RUNTIME_DIR, "runs")
BOXSPEC_DIR = os.path.join(RUNTIME_DIR, "boxspec")
os.makedirs(RUNS_DIR, exist_ok=True)
os.makedirs(BOXSPEC_DIR, exist_ok=True)

# 箱规表查找顺序：用户后台上传的优先，其次随包内置
BOXSPEC_CANDIDATES = [
    os.path.join(BOXSPEC_DIR, "补货箱规表.xlsx"),
    os.path.join(HERE, "补货箱规表.xlsx"),
]

SITE_NAMES = {"my": "马来", "th": "泰国", "ph": "菲律宾"}

# 下载注册表：id -> {'path', 'name', 'ts'}
_FILES = {}


# --------------------------------------------------------------------------
# 通用工具
# --------------------------------------------------------------------------
def site_of(payload):
    site = str(payload.get("site") or "my").strip().lower()
    return site if site in SITE_NAMES else "my"


def date_of(payload):
    """批次日期：优先用页面填写的日期；未填时从上传文件名中识别（如 11.9my库存导出表.xlsx -> 11.9）；都没有才用当天。

    输出文件名以批次日期为准，避免“11.9 的补货在 11.10 上传，被命名为 11.10”。
    """
    d = str(payload.get("date") or "").strip()
    if d:
        return d
    for v in payload.values():
        names = []
        if isinstance(v, dict):
            if isinstance(v.get("name"), str):
                names.append(v["name"])
            for sub in v.values():
                if isinstance(sub, dict) and isinstance(sub.get("name"), str):
                    names.append(sub["name"])
        for n in names:
            m = re.search(r"(\d{1,2})\.(\d{1,2})", n)
            if m:
                return "%d.%d" % (int(m.group(1)), int(m.group(2)))
    now = time.localtime()
    return "%d.%d" % (now.tm_mon, now.tm_mday)


def run_dir(tag):
    path = os.path.join(RUNS_DIR, "%s_%s" % (time.strftime("%Y%m%d_%H%M%S"), tag))
    os.makedirs(path, exist_ok=True)
    return path


def decode_upload(obj, dest_dir, fallback_name):
    """{'name':..,'data':base64} -> 本地路径"""
    if not obj:
        return None
    raw = obj.get("data")
    if not raw:
        return None
    if "," in raw and raw.strip().lower().startswith("data:"):
        raw = raw.split(",", 1)[1]
    name = os.path.basename(str(obj.get("name") or fallback_name))
    path = os.path.join(dest_dir, name)
    with open(path, "wb") as fh:
        fh.write(base64.b64decode(raw))
    return path


def register_file(path, name=None):
    fid = uuid.uuid4().hex[:12]
    _FILES[fid] = {"path": path, "name": name or os.path.basename(path), "ts": time.time()}
    stale = [k for k, v in _FILES.items() if time.time() - v["ts"] > 6 * 3600]
    for k in stale:
        _FILES.pop(k, None)
    return {"id": fid, "name": _FILES[fid]["name"]}


def find_boxspec():
    for p in BOXSPEC_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def load_engine():
    """每次请求重新构建引擎并热重载 core.py，保证箱规表或规则修改后立即生效"""
    importlib.reload(core)
    return core.ReplenishmentEngine(find_boxspec())


def boxspec_status():
    path = find_boxspec()
    if not path:
        return {"count": 0, "source": "未载入", "version": "", "updated": ""}
    st = os.stat(path)
    with open(path, "rb") as fh:
        digest = hashlib.sha1(fh.read()).hexdigest()[:8]
    return {
        "count": len(load_engine().box_specs),
        "source": os.path.basename(path),
        "version": digest,
        "updated": time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime)),
    }


def count_stats(preview):
    stats = {"filled": 0, "intercepted": 0, "no_boxspec": 0, "no_demand": 0}
    for row in preview:
        key = row.get("status")
        if key in stats:
            stats[key] += 1
    stats["total"] = len(preview)
    return stats


# --------------------------------------------------------------------------
# 业务接口
# --------------------------------------------------------------------------
def api_stage1(payload):
    site = site_of(payload)
    date = date_of(payload)
    out_dir = run_dir("stage1_%s" % site)

    base = decode_upload(payload.get("base"), out_dir, "官方仓补货模板.xlsx")
    if not base:
        default_tpl = os.path.join(HERE, "官方仓补货模板.xlsx")
        if os.path.exists(default_tpl):
            base = default_tpl
    sales = decode_upload(payload.get("sales"), out_dir, "马帮销量导出表.xlsx")
    inv = decode_upload(payload.get("inv"), out_dir, "库存导出表.xlsx")
    rpa = decode_upload(payload.get("rpa"), out_dir, "销售速度.xlsx")

    missing = [n for n, p in (("补货底表", base), ("马帮销量表", sales), ("官方仓库存表", inv)) if not p]
    if missing:
        return {"ok": False, "error": "缺少必需文件：%s" % "、".join(missing)}

    engine = load_engine()
    sales_map = engine.load_sales_data(sales, sheet_name="%s销量" % site)
    inv_map = engine.load_shopee_inventory(inv, sheet_name="%s库存" % site)
    rpa_map = engine.load_rpa_inventory(rpa) if rpa else {}

    sheet = "%s%s" % (date, site)
    out_path = os.path.join(out_dir, "官方仓补货总表_%s%s.xlsx" % (date, site))
    out_path, preview = engine.process_stage1_replenish(
        base_workbook_path=base,
        template_sheet_name="%s补货模板" % site,
        output_workbook_path=out_path,
        new_sheet_name=sheet,
        sales_map=sales_map,
        inv_map=inv_map,
        rpa_map=rpa_map,
        site_code=site,
    )
    stats = count_stats(preview)
    warnings = []
    if not engine.box_specs:
        warnings.append("箱规库为空，所有需补量都会留空，请先上传箱规表")
    elif stats["no_boxspec"]:
        warnings.append("%d 行因缺少箱规未测算，已留空" % stats["no_boxspec"])

    return {
        "ok": True,
        "sheet": sheet,
        "preview": preview,
        "stats": stats,
        "warnings": warnings,
        "matched": {
            "sales": len(sales_map),
            "inventory": len(inv_map),
            "speed": len(rpa_map),
        },
        "file": register_file(out_path, "官方仓补货总表_%s%s.xlsx" % (date, site)),
    }


def api_stage1_batch(payload):
    """阶段一（批量）：一次上传 3 国家 × (库存/销量/销售速度) 共 9 个文件，输出一个文件，
    含 日期+国家 命名的 sheet（如 11.9my / 11.9th / 11.9ph）。"""
    date = date_of(payload)
    out_dir = run_dir("stage1_batch")
    engine = load_engine()

    base = decode_upload(payload.get("base"), out_dir, "官方仓补货模板.xlsx")
    if not base:
        default_tpl = os.path.join(HERE, "官方仓补货模板.xlsx")
        if os.path.exists(default_tpl):
            base = default_tpl
    if not base:
        return {"ok": False, "error": "缺少补货底表（官方仓补货模板）"}

    country_inputs, missing, matched = {}, [], {}
    for site in ("my", "th", "ph"):
        group = payload.get(site) or {}
        inv = decode_upload(group.get("inv"), out_dir, "%s库存导出表.xlsx" % site)
        sales = decode_upload(group.get("sales"), out_dir, "%s销量导出表.xlsx" % site)
        rpa = decode_upload(group.get("rpa"), out_dir, "%s销售速度表.xlsx" % site)
        if not (inv and sales and rpa):
            missing.append(site)
            continue
        sales_map = engine.load_sales_data(sales, sheet_name="%s销量" % site)
        inv_map = engine.load_shopee_inventory(inv, sheet_name="%s库存" % site)
        rpa_map = engine.load_rpa_inventory(rpa)
        country_inputs[site] = {"sales_map": sales_map, "inv_map": inv_map, "rpa_map": rpa_map}
        matched[site] = {"sales": len(sales_map), "inventory": len(inv_map), "speed": len(rpa_map)}

    if not country_inputs:
        return {"ok": False, "error": "未检测到任何国家的完整 3 个文件（库存/销量/销售速度）"}

    out_path = os.path.join(out_dir, "官方仓补货总表_%s.xlsx" % date)
    out_path, results, sheets = engine.process_stage1_batch(
        base_workbook_path=base,
        output_workbook_path=out_path,
        date=date,
        country_inputs=country_inputs,
    )
    summary = {}
    previews = {}
    for site, preview in results.items():
        summary[site] = {"sheet": sheets[site], "stats": count_stats(preview), "rows": len(preview)}
        previews[site] = preview

    warnings = []
    if missing:
        warnings.append("以下国家未上传完整 3 个文件，已跳过：%s" % "、".join(s.upper() for s in missing))
    if not engine.box_specs:
        warnings.append("箱规库为空，无箱规的 SKU 将按需补量计算")

    return {
        "ok": True,
        "sheets": sheets,
        "summary": summary,
        "previews": previews,
        "matched": matched,
        "warnings": warnings,
        "file": register_file(out_path, "官方仓补货总表_%s.xlsx" % date),
    }


def api_stage2(payload):
    site = site_of(payload)
    date = date_of(payload)
    out_dir = run_dir("stage2_%s" % site)

    rep = decode_upload(payload.get("rep"), out_dir, "官方仓补货总表.xlsx")
    inb = decode_upload(payload.get("inb"), out_dir, "入库申请模板.xlsx")
    missing = [n for n, p in (("已审批补货总表", rep), ("入库申请模板", inb)) if not p]
    if missing:
        return {"ok": False, "error": "缺少必需文件：%s" % "、".join(missing)}

    out_path = os.path.join(out_dir, "%s%s_入库申请填写表_已填充.xlsx" % (date, site))
    out_path, filled, preview = load_engine().process_stage2_inbound(rep, inb, site, out_path)
    return {
        "ok": True,
        "filled": filled,
        "preview": preview,
        "file": register_file(out_path, os.path.basename(out_path)),
    }


def api_stage3(payload):
    site = site_of(payload)
    date = date_of(payload)
    out_dir = run_dir("stage3_%s" % site)

    summary = decode_upload(payload.get("summary"), out_dir, "汇总补货表.xlsx")
    pack = decode_upload(payload.get("pack"), out_dir, "打包模板.xlsx")
    missing = [n for n, p in (("汇总补货表", summary), ("Shopee打包模板", pack)) if not p]
    if missing:
        return {"ok": False, "error": "缺少必需文件：%s" % "、".join(missing)}

    out_path = os.path.join(out_dir, "%s虾皮%s_已打包.xlsx" % (date, SITE_NAMES[site]))
    out_path, groups, max_box, preview = load_engine().process_stage3_packing(
        summary_replenish_path=summary,
        shopee_packing_template_path=pack,
        site_code=site,
        output_path=out_path,
        return_preview=True,
    )
    return {
        "ok": True,
        "groups": groups,
        "max_box": max_box,
        "preview": preview,
        "file": register_file(out_path, os.path.basename(out_path)),
    }


def api_stage2_batch(payload):
    """阶段二（批量）：1 个含 3 国家 sheet 的已审批总表 + 3 个国家入库申请模板 -> 输出 3 个填写后的申请表。"""
    date = date_of(payload)
    out_dir = run_dir("stage2_batch")
    engine = load_engine()

    rep = decode_upload(payload.get("rep"), out_dir, "已审批补货总表.xlsx")
    if not rep:
        return {"ok": False, "error": "缺少已审批补货总表（含3国家sheet）"}

    results, missing = {}, []
    for site in ("my", "th", "ph"):
        tpl = decode_upload((payload.get(site) or {}).get("inb"), out_dir, "%s入库申请模板.xlsx" % site)
        if not tpl:
            missing.append(site)
            continue
        out_path = os.path.join(out_dir, "%s%s_入库申请填写表_已填充.xlsx" % (date, site))
        out_path, filled, preview = engine.process_stage2_inbound(rep, tpl, site, out_path)
        results[site] = {
            "filled": filled,
            "preview": preview,
            "file": register_file(out_path, os.path.basename(out_path)),
        }

    if not results:
        return {"ok": False, "error": "未检测到任何国家的入库申请模板"}
    warnings = []
    if missing:
        warnings.append("以下国家缺少入库申请模板，已跳过：%s" % "、".join(s.upper() for s in missing))
    return {"ok": True, "results": results, "warnings": warnings}


def api_stage3_batch(payload):
    """阶段三（批量）：1 个汇总补货表（含各国家 sheet）+ 各国家 Shopee 打包模板 -> 每个国家一个已打包文件。

    规则：
    - 每个国家只使用自己的 sheet 与打包模板，互不混用。
    - 汇总 sheet 中没有任何序号（箱号）的国家跳过并提示，不生成全 0 的表。
    - 缺少打包模板的国家跳过并提示。
    """
    date = date_of(payload)
    out_dir = run_dir("stage3_batch")
    engine = load_engine()

    summary = decode_upload(payload.get("summary"), out_dir, "汇总补货表.xlsx")
    if not summary:
        return {"ok": False, "error": "缺少汇总补货表（含各国家 sheet）"}

    results, skipped = {}, []
    for site in ("my", "th", "ph"):
        pack = decode_upload((payload.get(site) or {}).get("pack"), out_dir, "%s打包模板.xlsx" % site)
        if not pack:
            skipped.append("%s（缺少打包模板）" % site.upper())
            continue
        if not engine.has_box_sequence(summary, site):
            skipped.append("%s（汇总表中尚无序号，请先补充序号后再打包）" % site.upper())
            continue
        out_path = os.path.join(out_dir, "%s虾皮%s_已打包.xlsx" % (date, SITE_NAMES[site]))
        out_path, groups, max_box, preview = engine.process_stage3_packing(
            summary_replenish_path=summary,
            shopee_packing_template_path=pack,
            site_code=site,
            output_path=out_path,
            return_preview=True,
        )
        results[site] = {
            "groups": groups,
            "max_box": max_box,
            "preview": preview,
            "file": register_file(out_path, os.path.basename(out_path)),
        }

    if not results:
        return {"ok": False, "error": "没有可打包的国家：%s" % "；".join(skipped) if skipped else "未检测到任何国家的打包模板"}
    warnings = []
    if skipped:
        warnings.append("以下国家已跳过：%s" % "；".join(skipped))
    return {"ok": True, "results": results, "warnings": warnings}


def api_validate(payload):
    site = site_of(payload)
    out_dir = run_dir("validate_%s" % site)
    master = decode_upload(payload.get("master"), out_dir, "官方仓补货总表.xlsx")
    packing = decode_upload(payload.get("packing"), out_dir, "已打包.xlsx")
    engine = load_engine()
    return {"ok": True, "results": engine.validate_system(master, packing, site)}


def api_boxspec(payload):
    obj = payload.get("file")
    path = decode_upload(obj, BOXSPEC_DIR, "补货箱规表.xlsx")
    if not path:
        return {"ok": False, "error": "未收到文件"}
    target = os.path.join(BOXSPEC_DIR, "补货箱规表.xlsx")
    if os.path.abspath(path) != os.path.abspath(target):
        os.replace(path, target)
    engine = load_engine()
    if not engine.box_specs:
        return {"ok": False, "error": "未能从该文件解析出箱规，请确认首列为 SKU、且存在数量/长/宽/高/重列"}
    return {"ok": True, "boxspec": boxspec_status()}


# --------------------------------------------------------------------------
# HTTP 层
# --------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "ShopeeReplenish/" + APP_VERSION

    def log_message(self, fmt, *args):  # 静默访问日志，避免控制台刷屏
        pass

    # ---------- 基础响应 ----------
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")

    def _json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _body_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8") or "{}")

    # ---------- 路由 ----------
    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path in ("/", "/index.html"):
                return self._serve_index()
            if path == "/api/health":
                return self._json({
                    "ok": True,
                    "app": "shopee-replenishment",
                    "version": APP_VERSION,
                    "boxspec": boxspec_status(),
                    "sites": [{"code": k, "name": v} for k, v in SITE_NAMES.items()],
                })
            if path.startswith("/api/download/"):
                return self._serve_download(path.rsplit("/", 1)[-1])
            return self._json({"ok": False, "error": "not found"}, 404)
        except Exception as exc:
            return self._json({"ok": False, "error": str(exc), "trace": traceback.format_exc()}, 500)

    def do_POST(self):
        path = urlparse(self.path).path
        routes = {
            "/api/stage1": api_stage1,
            "/api/stage1_batch": api_stage1_batch,
            "/api/stage2_batch": api_stage2_batch,
            "/api/stage3_batch": api_stage3_batch,
            "/api/stage2": api_stage2,
            "/api/stage3": api_stage3,
            "/api/validate": api_validate,
            "/api/boxspec": api_boxspec,
        }
        handler = routes.get(path)
        if not handler:
            return self._json({"ok": False, "error": "not found"}, 404)
        try:
            payload = self._body_json()
        except Exception as exc:
            return self._json({"ok": False, "error": "请求体解析失败：%s" % exc}, 400)
        try:
            return self._json(handler(payload))
        except Exception as exc:
            return self._json({"ok": False, "error": str(exc), "trace": traceback.format_exc()}, 500)

    # ---------- 静态与下载 ----------
    def _serve_index(self):
        page = os.path.join(HERE, "index.html")
        if not os.path.exists(page):
            return self._json({"ok": False, "error": "index.html 不存在"}, 404)
        with open(page, "rb") as fh:
            body = fh.read()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _serve_download(self, fid):
        item = _FILES.get(unquote(fid))
        if not item or not os.path.exists(item["path"]):
            return self._json({"ok": False, "error": "文件已过期，请重新生成"}, 404)
        name = item["name"]
        with open(item["path"], "rb") as fh:
            body = fh.read()
        self.send_response(200)
        self.send_header("Content-Type",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Length", str(len(body)))
        self.send_header(
            "Content-Disposition",
            "attachment; filename=\"download.xlsx\"; filename*=UTF-8''%s" % quote(name),
        )
        self._cors()
        self.end_headers()
        self.wfile.write(body)


def quote(text):
    from urllib.parse import quote as _q
    return _q(text, safe="")


def pick_port(preferred):
    for port in [preferred] + list(range(preferred + 1, preferred + 21)):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("8765-%d 端口均被占用" % (preferred + 20))


def main():
    parser = argparse.ArgumentParser(description="Shopee 补货工具本地服务")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    port = pick_port(args.port)
    url = "http://127.0.0.1:%d/" % port
    status = boxspec_status()

    print("=" * 62)
    print(" Shopee 官方仓智能补货工具  v%s" % APP_VERSION)
    print(" 地址: %s" % url)
    print(" 箱规: %s (%d 个 SKU)" % (status["source"], status["count"]))
    print(" 关闭本窗口即停止服务")
    print("=" * 62)

    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
