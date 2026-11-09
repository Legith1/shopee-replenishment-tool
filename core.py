# -*- coding: utf-8 -*-
"""
Shopee Official Warehouse Full-Stage Replenishment Engine (虾皮官方仓全流程智能补货引擎)
========================================================================================
阶段一：数据清洗对齐、需补量智能测算、保留条件格式与动态IF公式生成补货总表
阶段二：已审批总表数据自动映射至 Shopee 入库申请填写表（采购批准与出口货值）
阶段三：汇总补货表智能装箱解析（均分连箱、同色拼箱/混装）与箱规尺寸重量自动匹配
数据校验：SKU字典对齐校验、装箱闭环平衡校验、单箱重量尺寸安全校验
"""

import io
import os
import re
import copy
import math
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

_ORIG_LOAD_WORKBOOK = openpyxl.load_workbook
_MINIMAL_STYLES_XML = (
    b'<?xml version="1.0" encoding="UTF-8"?>'
    b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>'
)


def safe_load_workbook(file_path, *args, **kwargs):
    """兼容 Shopee / ERP 导出 Excel 含非法 styles.xml（如 border style="" 或 style="none"）的安全加载器"""
    try:
        return _ORIG_LOAD_WORKBOOK(file_path, *args, **kwargs)
    except Exception:
        if not isinstance(file_path, (str, os.PathLike)) or not os.path.exists(file_path):
            raise
        # 第一级修复：清洗 xl/styles.xml 中的非法 border style="" / style="none" 属性
        try:
            buf = io.BytesIO()
            with zipfile.ZipFile(file_path, 'r') as zin, zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == 'xl/styles.xml':
                        txt = data.decode('utf-8', errors='ignore')
                        txt = re.sub(r'\s+style="(?:none|)"', '', txt)
                        data = txt.encode('utf-8')
                    zout.writestr(item, data)
            buf.seek(0)
            return _ORIG_LOAD_WORKBOOK(buf, *args, **kwargs)
        except Exception:
            # 第二级兜底：替换为极简合法的 xl/styles.xml，确保数据读写不中断
            buf2 = io.BytesIO()
            with zipfile.ZipFile(file_path, 'r') as zin, zipfile.ZipFile(buf2, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = _MINIMAL_STYLES_XML if item.filename == 'xl/styles.xml' else zin.read(item.filename)
                    zout.writestr(item, data)
            buf2.seek(0)
            return _ORIG_LOAD_WORKBOOK(buf2, *args, **kwargs)

class ReplenishmentEngine:
    def __init__(self, box_spec_path=None):
        self.box_spec_path = box_spec_path
        self.box_specs = {}  # {sku_clean: {'qty': int, 'length': float, 'width': float, 'height': float, 'weight': float, 'note': str}}
        if box_spec_path and os.path.exists(box_spec_path):
            self.load_box_specs(box_spec_path)

    def load_box_specs(self, file_path):
        """加载补货箱规表（包含装量与长宽高重）"""
        self.box_spec_path = file_path
        wb = safe_load_workbook(file_path, data_only=True)
        ws = wb.active
        self.box_specs = {}

        # 扫描表头定位长宽高重列
        header_row = 1
        qty_col, l_col, w_col, h_col, wt_col, note_col = 3, 4, 5, 6, 7, 8
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(header_row, c).value or '').strip()
            if '数量' in val or '箱规' in val:
                qty_col = c
            elif '长' in val:
                l_col = c
            elif '宽' in val:
                w_col = c
            elif '高' in val:
                h_col = c
            elif '重' in val:
                wt_col = c
            elif '备注' in val:
                note_col = c

        for r in range(2, ws.max_row + 1):
            sku = ws.cell(r, 1).value
            if sku is not None:
                sku_clean = self.clean_key(sku)
                qty = ws.cell(r, qty_col).value
                length = ws.cell(r, l_col).value
                width = ws.cell(r, w_col).value
                height = ws.cell(r, h_col).value
                weight = ws.cell(r, wt_col).value
                note = ws.cell(r, note_col).value

                def to_num(v):
                    if v is None or str(v).strip() == '':
                        return None
                    try:
                        return float(v)
                    except:
                        return None

                qty_val = None
                if qty is not None and str(qty).strip() != '':
                    try:
                        qty_val = int(float(qty))
                    except:
                        qty_val = None

                self.box_specs[sku_clean] = {
                    'qty': qty_val,
                    'length': to_num(length),
                    'width': to_num(width),
                    'height': to_num(height),
                    'weight': to_num(weight),
                    'note': str(note).strip() if note else None
                }
        return len(self.box_specs)

    #: 站点 -> 工作表别名（中文优先，避免 'th'/'my' 匹配不到中文表名而误用首表）
    SITE_SHEET_ALIASES = {
        'my': ['马来', 'my', 'malaysia'],
        'th': ['泰国', 'th', 'thailand'],
        'ph': ['菲律宾', 'ph', 'philippines'],
    }

    @classmethod
    def resolve_sheet(cls, workbook, site_code, preferred=None):
        """按站点解析工作表名：精确名 > 含站点别名的模板/补货表 > 含站点别名的任意表 > 首个工作表"""
        names = list(workbook.sheetnames)
        if not names:
            return None
        if preferred and preferred in names:
            return preferred
        aliases = cls.SITE_SHEET_ALIASES.get(str(site_code).lower(), [])
        for kw in aliases:
            for n in names:
                if kw.lower() in n.lower() and ('模板' in n or '补货' in n):
                    return n
        for kw in aliases:
            for n in names:
                if kw.lower() in n.lower():
                    return n
        return names[0]

    @staticmethod
    def clean_key(val):
        """统一清洗关联键（SKU / 仓库SKU ID）"""
        if val is None:
            return ""
        s = str(val).strip()
        if s.endswith('.0') and s[:-2].isdigit():
            s = s[:-2]
        return s

    # =========================================================================
    # 阶段一：补货总表测算核心逻辑
    # =========================================================================
    def load_sales_data(self, sales_file, sheet_name=None):
        """加载马帮销量表（匹配 SKU -> 销售数量）"""
        wb = safe_load_workbook(sales_file, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

        header_row, sku_col, sales_col = 1, None, None
        for r in range(1, min(6, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if val in ['SKU', '商品SKU', '库存SKU', 'sku', '商品编码']:
                    sku_col, header_row = c, r
                elif val in ['销售数量', '30天销量', '30天销售量', '近30天销量']:
                    sales_col = c
            if sku_col and sales_col:
                break

        if not sales_col:
            for r in range(1, min(6, ws.max_row + 1)):
                for c in range(1, ws.max_column + 1):
                    val = str(ws.cell(r, c).value or '').strip()
                    if '平均' not in val and any(k in val for k in ['销售数量', '销量', '数量']):
                        sales_col = c
                        break
                if sales_col:
                    break

        sku_col = sku_col or 1
        sales_col = sales_col or 10

        sales_map = {}
        for r in range(header_row + 1, ws.max_row + 1):
            sku = self.clean_key(ws.cell(r, sku_col).value)
            raw_sales = ws.cell(r, sales_col).value
            if sku:
                try:
                    sales_val = float(raw_sales) if raw_sales is not None else 0.0
                except:
                    sales_val = 0.0
                sales_map[sku] = sales_val
        return sales_map

    def load_shopee_inventory(self, inventory_file, sheet_name=None):
        """加载Shopee官方仓库存表（匹配 仓库SKU ID -> 在仓可售数量 / 在途待处理IR）"""
        wb = safe_load_workbook(inventory_file, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

        header_row, wh_sku_col, available_col, transit_col = 1, None, None, None
        for r in range(1, min(6, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if val in ['仓库SKU ID', '仓库SKU', 'Warehouse SKU ID', 'Warehouse SKU']:
                    wh_sku_col, header_row = c, r
                elif val in ['可售数量', '在仓', 'Available']:
                    available_col = c
                elif val in ['待处理的IR已到达', '在途', '待入库', 'In Transit']:
                    transit_col = c
            if wh_sku_col and available_col:
                break

        inv_map = {}
        if wh_sku_col:
            for r in range(header_row + 1, ws.max_row + 1):
                wh_sku = self.clean_key(ws.cell(r, wh_sku_col).value)
                if wh_sku:
                    raw_avail = ws.cell(r, available_col).value if available_col else 0
                    raw_transit = ws.cell(r, transit_col).value if transit_col else 0
                    try:
                        avail_val = float(raw_avail) if raw_avail is not None else 0.0
                    except:
                        avail_val = 0.0
                    try:
                        transit_val = float(raw_transit) if raw_transit is not None else 0.0
                    except:
                        transit_val = 0.0
                    inv_map[wh_sku] = {'in_warehouse': avail_val, 'in_transit': transit_val}
        return inv_map

    def load_rpa_inventory(self, rpa_file):
        """加载RPA导出的销售速度表"""
        wb = safe_load_workbook(rpa_file, data_only=True)
        ws = wb.active
        header_row, wh_sku_col, speed_col = 1, None, None
        for r in range(1, min(6, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if val in ['仓库SKU', '仓库SKU ID', 'Warehouse SKU', 'Warehouse SKU ID']:
                    wh_sku_col, header_row = c, r
                elif '销售速度' in val or 'Sales Velocity' in val:
                    speed_col = c
            if wh_sku_col and speed_col:
                break

        rpa_map = {}
        if wh_sku_col and speed_col:
            for r in range(header_row + 1, ws.max_row + 1):
                wh_sku = self.clean_key(ws.cell(r, wh_sku_col).value)
                if wh_sku:
                    raw_speed = ws.cell(r, speed_col).value
                    try:
                        speed_val = float(raw_speed) if raw_speed is not None else 0.0
                    except:
                        speed_val = 0.0
                    rpa_map[wh_sku] = speed_val
        return rpa_map

    #: 备注拦截关键词（命中任一即整行留空，不参与补货测算）
    INTERCEPT_KEYWORDS = ['停', '不补', '超100cm', '超100', '超长', '下架', '停产', '无库存', '无货', 'dw不补']

    # 状态枚举：filled=已测算 / intercepted=备注拦截 / no_boxspec=无箱规 / no_demand=无需补货
    STATUS_FILLED = 'filled'
    STATUS_INTERCEPTED = 'intercepted'
    STATUS_NO_BOXSPEC = 'no_boxspec'
    STATUS_NO_DEMAND = 'no_demand'

    def evaluate_row(self, sku, cz_replenish, max_limit_val, remark, box_specs=None):
        """需补量测算的唯一实现。

        返回 (need_qty, status, reason)：
        - need_qty: int 或 None（None 表示该行留空）
        - status:   filled / intercepted / no_boxspec / no_demand
        - reason:   人可读的中文原因，用于前端预览与复核

        测算顺序严格遵循业务规则：
        1) 备注命中拦截词 -> 留空
        2) 充足补货量 <= 0 -> 留空
        3) 箱规缺失或箱数<=0 -> 留空
        4) 按箱规整箱试算：boxes = round(cz / 箱规)（Python round，银行家舍入）
        5) 若整箱结果超出站点上限：箱数-1 仍达标则退一箱，否则退到 min(round(cz), floor(上限))
        """
        specs = self.box_specs if box_specs is None else box_specs

        if remark is not None and str(remark).strip() != '':
            rem_str = str(remark).strip()
            for kw in self.INTERCEPT_KEYWORDS:
                if kw in rem_str:
                    return None, self.STATUS_INTERCEPTED, '备注命中拦截词「%s」，整行留空' % kw

        if cz_replenish is None or cz_replenish <= 0:
            return None, self.STATUS_NO_DEMAND, '充足补货量≤0，无需补货'

        sku_clean = self.clean_key(sku)
        box_info = specs.get(sku_clean)
        if not box_info or not box_info.get('qty'):
            return None, self.STATUS_NO_BOXSPEC, '箱规库无此 SKU 或未维护整箱数量'

        box_qty = box_info['qty']
        if box_qty <= 0:
            return None, self.STATUS_NO_BOXSPEC, '箱规整箱数量无效（≤0）'

        boxes = round(cz_replenish / box_qty)
        if boxes == 0 and cz_replenish >= (0.5 * box_qty):
            boxes = 1

        candidate = boxes * box_qty
        limit = None
        if max_limit_val is not None and not isinstance(max_limit_val, str):
            try:
                if not math.isnan(float(max_limit_val)):
                    limit = float(max_limit_val)
            except (TypeError, ValueError):
                limit = None

        capped = False
        if limit is not None and limit > 0:
            if candidate > limit:
                capped = True
                if boxes > 1 and ((boxes - 1) * box_qty) <= limit:
                    candidate = (boxes - 1) * box_qty
                else:
                    candidate = min(round(cz_replenish), math.floor(limit))

        if candidate <= 0:
            return None, self.STATUS_NO_DEMAND, '整箱试算结果≤0，无需补货'

        need_qty = int(round(candidate))
        if capped:
            reason = '整箱试算 %d 件超出站点上限 %.0f 件，已按上限平滑为 %d 件（箱规 %d）' % (
                int(boxes * box_qty), limit, need_qty, box_qty)
        else:
            reason = '整箱试算 %d 箱 × %d 件/箱 = %d 件' % (int(boxes), box_qty, need_qty)
        return need_qty, self.STATUS_FILLED, reason

    def calculate_replenishment_quantity(self, sku, cz_replenish, max_limit_val, remark):
        """智能需补量计算（保留原签名，内部统一走 evaluate_row）"""
        need, _status, _reason = self.evaluate_row(sku, cz_replenish, max_limit_val, remark)
        return need

    def process_stage1_replenish(self, base_workbook_path, template_sheet_name, output_workbook_path, new_sheet_name, sales_map, inv_map, rpa_map, site_code='my'):
        """阶段一：生成官方仓补货总表并写入公式与条件格式"""
        wb = safe_load_workbook(base_workbook_path)
        site_lower = site_code.lower()

        template_sheet_name = self.resolve_sheet(wb, site_code, template_sheet_name)

        ws_template = wb[template_sheet_name]

        if new_sheet_name in wb.sheetnames:
            wb.remove(wb[new_sheet_name])

        ws_new = wb.copy_worksheet(ws_template)
        ws_new.title = new_sheet_name

        for cf in ws_template.conditional_formatting._cf_rules:
            ws_new.conditional_formatting._cf_rules[cf] = copy.deepcopy(ws_template.conditional_formatting._cf_rules[cf])

        header_row = 1
        headers = {str(ws_new.cell(header_row, c).value or '').strip(): c for c in range(1, ws_new.max_column + 1)}

        sku_col = headers.get('SKU', 2)
        wh_sku_col = headers.get('仓库SKU ID', 3)
        sales_col = headers.get('30天销量', 4)
        in_wh_col = headers.get('在仓', 5)
        in_transit_col = headers.get('在途', 6)
        total_col = headers.get('合计', 7)
        d60_sales_col = headers.get('补货至60天销售量', 8)
        days_avail_col = headers.get('剩余可售天数', 9)
        cz_replenish_col = headers.get('充足补货量', 10)

        need_col = next((c for h, c in headers.items() if '需补' in h), 11)
        speed_col = headers.get('销售速度', 13)
        d60_col = headers.get('60天', 14)
        d75_col = headers.get('75天', 15) if site_lower != 'ph' else None
        remark_col = headers.get('备注', 15 if site_lower == 'ph' else 16)
        # 模板新增的「数量」列：回填该 SKU 在箱规表中的整箱数量（仅作核对参考，不参与测算）
        box_qty_col = headers.get('数量')

        preview_rows = []

        for r in range(header_row + 1, ws_new.max_row + 1):
            sku_val = self.clean_key(ws_new.cell(r, sku_col).value)
            wh_sku_val = self.clean_key(ws_new.cell(r, wh_sku_col).value)
            remark_val = ws_new.cell(r, remark_col).value
            name_val = ws_new.cell(r, 1).value

            if not sku_val and not wh_sku_val:
                continue

            sales_val = sales_map.get(sku_val, 0.0)
            ws_new.cell(r, sales_col).value = int(sales_val) if sales_val.is_integer() else sales_val

            inv_data = inv_map.get(wh_sku_val, {'in_warehouse': 0.0, 'in_transit': 0.0})
            in_wh = inv_data['in_warehouse']
            in_transit = inv_data['in_transit']
            ws_new.cell(r, in_wh_col).value = int(in_wh) if in_wh.is_integer() else in_wh
            ws_new.cell(r, in_transit_col).value = int(in_transit) if in_transit.is_integer() else in_transit

            speed_val = rpa_map.get(wh_sku_val, 0.0)
            ws_new.cell(r, speed_col).value = round(speed_val, 2)

            # 注意：G/H/I/J/N/O 列的空值保护公式由模板继承而来，此处不再覆盖。
            # 之前用 =E+F / =2*D 覆盖会导致空白行显示 0，破坏模板的 ISBLANK 保护。

            cur_total = in_wh + in_transit
            cur_cz = (sales_val * 2) - cur_total
            cur_d60 = (60.0 * speed_val) - cur_total
            cur_d75 = (75.0 * speed_val) - cur_total if site_lower != 'ph' else None
            max_limit = cur_d60 if site_lower == 'ph' else cur_d75

            need_qty, status, reason = self.evaluate_row(
                sku=sku_val, cz_replenish=cur_cz, max_limit_val=max_limit, remark=remark_val)
            ws_new.cell(r, need_col).value = need_qty if need_qty is not None else None

            spec = self.box_specs.get(sku_val)
            box_qty_val = spec.get('qty') if spec else None
            if box_qty_col:
                ws_new.cell(r, box_qty_col).value = box_qty_val if box_qty_val else None

            if sales_val > 0:
                days_avail = round(in_wh / (sales_val / 30.0), 1)
            else:
                days_avail = None

            preview_rows.append({
                'row': r, 'name': name_val, 'sku': sku_val, 'wh_sku': wh_sku_val,
                'sales': sales_val, 'in_wh': in_wh, 'in_transit': in_transit,
                'total': cur_total, 'days': days_avail, 'cz': round(cur_cz, 4),
                'need': need_qty if need_qty is not None else '',
                'status': status, 'reason': reason,
                'box_qty': box_qty_val if box_qty_val else None,
                'speed': speed_val, 'd60': round(cur_d60, 4),
                'd75': round(cur_d75, 4) if cur_d75 is not None else None,
                'remark': remark_val
            })

        wb._sheets = [ws_new] + [s for s in wb._sheets if s != ws_new]
        wb.save(output_workbook_path)
        return output_workbook_path, preview_rows

    # =========================================================================
    # 阶段二：生成 Shopee 入库申请表核心逻辑
    # =========================================================================
    def process_stage2_inbound(self, approved_master_path, inbound_template_path, site_code, output_path):
        """阶段二：根据审批完的补货总表生成入库申请表"""
        wb_rep = safe_load_workbook(approved_master_path, data_only=True)
        # 优先取“批次数据表”（如 11.7my），排除 *_补货模板 之类的空白模板表
        names = list(wb_rep.sheetnames)
        aliases = self.SITE_SHEET_ALIASES.get(str(site_code).lower(), [])
        cands = [n for n in names if any(k.lower() in n.lower() for k in aliases)] or names
        data_sheets = [n for n in cands if '模板' not in n]
        target_sheet = (data_sheets or cands)[-1]
        ws_rep = wb_rep[target_sheet]

        rep_headers = {str(ws_rep.cell(1, c).value or '').strip(): c for c in range(1, ws_rep.max_column + 1)}
        col_wh_sku = rep_headers.get('仓库SKU ID', 3)
        col_approved = rep_headers.get('采购批准', 12)
        col_export = rep_headers.get('出口货值', 17 if site_code.lower() != 'ph' else 16)

        rep_map = {}
        for r in range(2, ws_rep.max_row + 1):
            wh_sku_id = ws_rep.cell(r, col_wh_sku).value  # 仓库SKU ID
            approved_qty = ws_rep.cell(r, col_approved).value  # 采购批准
            export_val = ws_rep.cell(r, col_export).value  # 出口货值
            if wh_sku_id:
                rep_map[self.clean_key(wh_sku_id)] = {
                    'approved_qty': approved_qty,
                    'export_val': export_val
                }

        wb_inb = safe_load_workbook(inbound_template_path)
        ws_inb = wb_inb.active

        filled_rows = 0
        preview_data = []

        for r in range(2, ws_inb.max_row + 1):
            wh_sku_id = self.clean_key(ws_inb.cell(r, 4).value)  # col 4: Warehouse SKU ID
            prod_name = ws_inb.cell(r, 1).value
            spec_name = ws_inb.cell(r, 2).value
            ext_sku = ws_inb.cell(r, 5).value

            if wh_sku_id in rep_map:
                item = rep_map[wh_sku_id]
                app_qty = item['approved_qty']
                exp_val = item['export_val']

                qty_val = None
                if app_qty is not None and str(app_qty).strip() not in ['', '0', '#N/A']:
                    try:
                        num = int(float(app_qty))
                        if num > 0:
                            qty_val = num
                    except:
                        pass

                ws_inb.cell(r, 6).value = qty_val

                if exp_val is not None and str(exp_val).strip() not in ['', '#N/A']:
                    try:
                        ws_inb.cell(r, 7).value = float(exp_val) if not float(exp_val).is_integer() else int(float(exp_val))
                    except:
                        ws_inb.cell(r, 7).value = exp_val

                if qty_val:
                    filled_rows += 1
                    if len(preview_data) < 6:
                        preview_data.append({
                            'name': prod_name, 'spec': spec_name, 'wh_sku': wh_sku_id,
                            'ext_sku': ext_sku, 'qty': qty_val, 'val': ws_inb.cell(r, 7).value
                        })
            else:
                ws_inb.cell(r, 6).value = None

        wb_inb.save(output_path)
        return output_path, filled_rows, preview_data

    # =========================================================================
    # 阶段三：打包装箱分配与尺寸重量自动生成核心逻辑
    # =========================================================================
    def process_stage3_packing(self, summary_replenish_path, shopee_packing_template_path, site_code, output_path,
                               return_preview=False):
        """阶段三：解析汇总补货表（含均分连箱与同色拼箱），填充已打包数量与尺寸重量 Sheet

        return_preview=True 时额外返回逐行装箱预览（供 Web UI 使用），
        默认返回 3 元组以兼容既有桌面版调用。
        """
        wb_sum = safe_load_workbook(summary_replenish_path, data_only=True)
        site_name_map = {'my': '马来', 'th': '泰国', 'ph': '菲律宾'}
        target_sheet_name = site_name_map.get(site_code.lower(), site_code)
        
        matched = [s for s in wb_sum.sheetnames if target_sheet_name in s or site_code.lower() in s.lower()]
        ws_sum = wb_sum[matched[0]] if matched else wb_sum.active

        # 解析箱子分组：支持连箱 (如 3-7) 与 拼箱 (同一序号下的多行商品)
        box_groups = []
        current_group = None

        for r in range(4, ws_sum.max_row + 1):
            name = ws_sum.cell(r, 1).value
            qty = ws_sum.cell(r, 2).value
            seq = ws_sum.cell(r, 3).value
            box_cnt = ws_sum.cell(r, 4).value

            if not name and not qty:
                continue

            name_clean = str(name).strip()
            qty_num = float(qty) if qty else 0.0

            if seq is not None and str(seq).strip() != '':
                seq_str = str(seq).strip()
                boxes = self.parse_box_sequence(seq_str, box_cnt)
                current_group = {
                    'seq_str': seq_str,
                    'boxes': boxes,
                    'items': [(name_clean, qty_num)]
                }
                box_groups.append(current_group)
            else:
                if current_group:
                    current_group['items'].append((name_clean, qty_num))

        # 构建 SKU -> {box_num: qty} 的装箱分布字典
        # 同时记录每个箱子装了哪些 SKU (用于判断是否混箱)
        sku_box_map = {}  # {sku_clean: {box_idx: qty}}
        box_sku_map = {}  # {box_idx: [(sku_clean, qty)]}

        for g in box_groups:
            boxes = g['boxes']
            num_boxes = len(boxes)
            items = g['items']
            is_multi_item = len(items) > 1

            for sku_name, total_qty in items:
                sku_clean = self.clean_key(sku_name)
                if sku_clean not in sku_box_map:
                    sku_box_map[sku_clean] = {}

                qty_per_box = total_qty / num_boxes
                for b_idx in boxes:
                    sku_box_map[sku_clean][b_idx] = sku_box_map[sku_clean].get(b_idx, 0) + qty_per_box
                    if b_idx not in box_sku_map:
                        box_sku_map[b_idx] = []
                    box_sku_map[b_idx].append((sku_clean, qty_per_box))

        # 写入 Shopee 官方打包模板
        wb_pack = safe_load_workbook(shopee_packing_template_path)
        ws_qty = wb_pack['已打包数量'] if '已打包数量' in wb_pack.sheetnames else wb_pack.active
        ws_dim = wb_pack['尺寸和重量'] if '尺寸和重量' in wb_pack.sheetnames else None

        # 1. 填充【已打包数量】
        max_box_idx = max(box_sku_map.keys()) if box_sku_map else 1
        
        # 定位各箱子列 (箱子 #1 对应某列) 与【总打包数】列
        box_col_start = 8  # 默认第8列开始为 箱子 #1
        total_pack_col = 7  # 默认第7列为总打包数
        for c in range(1, ws_qty.max_column + 1):
            if self.is_box_header(ws_qty.cell(1, c).value, 1):
                box_col_start = c
                break
        for c in range(1, ws_qty.max_column + 1):
            val = str(ws_qty.cell(1, c).value or '').strip()
            if '总打包' in val or ('已打包' in val and '箱子' not in val):
                total_pack_col = c
                break

        preview_rows = []

        for r in range(2, ws_qty.max_row + 1):
            prod_name = str(ws_qty.cell(r, 1).value or '').strip()
            ext_sku = str(ws_qty.cell(r, 5).value or '').strip()
            wh_sku = str(ws_qty.cell(r, 4).value or '').strip()

            # 寻找匹配的 SKU 键
            matched_sku = None
            for key in [ext_sku, prod_name, wh_sku]:
                k_clean = self.clean_key(key)
                if k_clean in sku_box_map:
                    matched_sku = k_clean
                    break

            if not matched_sku:
                # 模糊匹配
                for k in sku_box_map:
                    if k in ext_sku or k in prod_name:
                        matched_sku = k
                        break

            if matched_sku:
                box_dist = sku_box_map[matched_sku]
                total_packed = sum(box_dist.values())
                ws_qty.cell(r, total_pack_col).value = int(total_packed) if total_packed.is_integer() else total_packed

                for b_idx, q in box_dist.items():
                    col = box_col_start + (b_idx - 1)
                    ws_qty.cell(r, col).value = int(q) if q.is_integer() else q

                preview_rows.append({
                    'name': prod_name or matched_sku, 'sku': matched_sku,
                    'ext_sku': ext_sku, 'wh_sku': wh_sku,
                    'total': int(total_packed) if float(total_packed).is_integer() else round(total_packed, 2),
                    'boxes': {str(k): (int(v) if float(v).is_integer() else round(v, 2))
                              for k, v in sorted(box_dist.items())},
                    'box_count': len(box_dist),
                    'status': 'ok' if len(box_dist) == 1 else 'mixed',
                    'reason': ('单箱整装' if len(box_dist) == 1
                               else '连箱/拼箱：均分至 %d 个箱子' % len(box_dist))
                })
            else:
                preview_rows.append({
                    'name': prod_name, 'sku': '', 'ext_sku': ext_sku, 'wh_sku': wh_sku,
                    'total': 0, 'boxes': {}, 'box_count': 0,
                    'status': 'unmatched', 'reason': '汇总补货表中未找到对应 SKU，未分配箱号'
                })

        # 2. 填充【尺寸和重量】
        if ws_dim:
            # 找到尺寸和重量的各行
            len_row, wid_row, hgt_row, wt_row = 2, 3, 4, 5
            for r in range(1, ws_dim.max_row + 1):
                val = str(ws_dim.cell(r, 1).value or '').strip()
                if '长' in val:
                    len_row = r
                elif '宽' in val:
                    wid_row = r
                elif '高' in val:
                    hgt_row = r
                elif '重' in val:
                    wt_row = r

            # 尺寸和重量表的箱子列起点：优先按表头探测，默认第2列为 箱子 #1
            dim_col_start = 2
            for c in range(1, ws_dim.max_column + 1):
                if self.is_box_header(ws_dim.cell(1, c).value, 1):
                    dim_col_start = c
                    break

            for b_idx, skus_in_box in box_sku_map.items():
                col = dim_col_start + (b_idx - 1)
                # 判断是否为单商品整箱
                if len(skus_in_box) == 1:
                    sku_clean, _ = skus_in_box[0]
                    spec = self.box_specs.get(sku_clean)
                    if spec and spec.get('length'):
                        ws_dim.cell(len_row, col).value = spec['length']
                        ws_dim.cell(wid_row, col).value = spec['width']
                        ws_dim.cell(hgt_row, col).value = spec['height']
                        ws_dim.cell(wt_row, col).value = spec['weight']
                    else:
                        ws_dim.cell(len_row, col).value = None
                        ws_dim.cell(wid_row, col).value = None
                        ws_dim.cell(hgt_row, col).value = None
                        ws_dim.cell(wt_row, col).value = None
                else:
                    # 拼箱 / 混箱：自动留空等待手动复核
                    ws_dim.cell(len_row, col).value = None
                    ws_dim.cell(wid_row, col).value = None
                    ws_dim.cell(hgt_row, col).value = None
                    ws_dim.cell(wt_row, col).value = None

        wb_pack.save(output_path)
        if return_preview:
            return output_path, len(box_groups), max_box_idx, preview_rows
        return output_path, len(box_groups), max_box_idx

    @staticmethod
    def is_box_header(val, idx):
        """精确判断表头是否为「箱子 #idx」；避免 '箱子 #1' 误匹配 '箱子 #19'"""
        s = str(val or '').replace(' ', '').replace('\u3000', '')
        return s in ('箱子#%d' % idx, '箱号#%d' % idx, '箱子%d' % idx, '箱号%d' % idx)

    @staticmethod
    def parse_box_sequence(seq_str, box_cnt=None):
        """解析箱号序号字符串，如 '1' -> [1], '3-7' -> [3,4,5,6,7]"""
        s = seq_str.replace('~', '-').replace('到', '-')
        if '-' in s:
            parts = s.split('-')
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
                return list(range(start, end + 1))
            except:
                pass
        try:
            val = int(float(s))
            if box_cnt and int(box_cnt) > 1:
                return list(range(val, val + int(box_cnt)))
            return [val]
        except:
            return [1]

    # =========================================================================
    # 数据防呆校验网关 (3重自动核验)
    # =========================================================================
    def validate_system(self, master_path=None, packing_path=None, site_code='my'):
        """执行 3 项自动化防呆核验并生成真实报告（不再返回写死结果）

        1) SKU 字典对齐：补货总表中每个 SKU 是否在箱规库中存在
        2) 装箱闭环平衡：打包表中每行 各箱数量之和 == 批准/总打包数量
        3) 单箱尺寸重量：箱规库中单箱重量 >30kg 或单边 >100cm 预警
        """
        results = {}

        # ---- 1. SKU 字典对齐 ----
        sku_unknown, sku_total = [], 0
        if master_path and os.path.exists(master_path):
            try:
                wb = safe_load_workbook(master_path, data_only=True)
                ws = wb[self.resolve_sheet(wb, site_code)]
                headers = {str(ws.cell(1, c).value or '').strip(): c for c in range(1, ws.max_column + 1)}
                col_sku = headers.get('SKU', 2)
                col_wh = headers.get('仓库SKU ID', 3)
                for r in range(2, ws.max_row + 1):
                    sku = self.clean_key(ws.cell(r, col_sku).value)
                    wh = self.clean_key(ws.cell(r, col_wh).value)
                    if not sku and not wh:
                        continue
                    sku_total += 1
                    if sku and sku not in self.box_specs:
                        sku_unknown.append(sku)
                if sku_total == 0:
                    results['sku_check'] = {'pass': False, 'msg': '补货总表中未读到任何商品行', 'detail': []}
                elif sku_unknown:
                    results['sku_check'] = {
                        'pass': False,
                        'msg': '%d/%d 个 SKU 未在箱规库中找到（%d 个），这些行需补量会留空' % (
                            len(sku_unknown), sku_total, len(sku_unknown)),
                        'detail': sku_unknown[:50]}
                else:
                    results['sku_check'] = {
                        'pass': True, 'msg': '%d/%d 个 SKU 全部命中箱规库' % (sku_total, sku_total), 'detail': []}
            except Exception as e:
                results['sku_check'] = {'pass': False, 'msg': '总表读取失败：%s' % e, 'detail': []}
        else:
            results['sku_check'] = {'pass': False, 'msg': '未提供补货总表，无法校验 SKU 字典', 'detail': []}

        # ---- 2. 装箱闭环平衡 ----
        if packing_path and os.path.exists(packing_path):
            try:
                wb_p = safe_load_workbook(packing_path, data_only=True)
                ws_p = wb_p['已打包数量'] if '已打包数量' in wb_p.sheetnames else wb_p.active
                box_start, total_col, approved_col = None, None, None
                for c in range(1, ws_p.max_column + 1):
                    val = str(ws_p.cell(1, c).value or '').strip()
                    if box_start is None and self.is_box_header(val, 1):
                        box_start = c
                    if '总打包' in val or ('已打包' in val and '箱子' not in val):
                        total_col = total_col or c
                    if '批准' in val:
                        approved_col = approved_col or c
                if box_start is None:
                    results['balance_check'] = {'pass': False, 'msg': '打包模板未找到「箱子 #1」列', 'detail': []}
                else:
                    diffs, checked = [], 0
                    for r in range(2, ws_p.max_row + 1):
                        name = str(ws_p.cell(r, 1).value or '').strip()
                        if not name:
                            continue
                        s = 0.0
                        for c in range(box_start, ws_p.max_column + 1):
                            v = ws_p.cell(r, c).value
                            if isinstance(v, (int, float)):
                                s += float(v)
                        if s <= 0:
                            continue
                        checked += 1
                        target = None
                        for c in (approved_col, total_col):
                            if c:
                                v = ws_p.cell(r, c).value
                                if isinstance(v, (int, float)):
                                    target = float(v)
                                    break
                        if target is not None and abs(s - target) > 0.001:
                            diffs.append('%s: 各箱合计 %.0f ≠ 数量 %.0f（行%d）' % (name, s, target, r))
                    results['balance_check'] = {
                        'pass': not diffs,
                        'msg': ('%d 行装箱数量 100%% 闭环无偏差' % checked) if not diffs else ('%d 行存在装箱数量偏差' % len(diffs)),
                        'detail': diffs[:50]}
            except Exception as e:
                results['balance_check'] = {'pass': False, 'msg': '打包表读取失败：%s' % e, 'detail': []}
        else:
            results['balance_check'] = {'pass': False, 'msg': '未提供打包表，无法校验装箱闭环（需先执行阶段三）', 'detail': []}

        # ---- 3. 单箱尺寸重量安全 ----
        overweight, oversize = [], []
        for sku, spec in self.box_specs.items():
            w = spec.get('weight')
            dims = [spec.get('length'), spec.get('width'), spec.get('height')]
            if w and w > 30:
                overweight.append('%s: %.1fkg' % (sku, w))
            if any(d and d > 100 for d in dims):
                oversize.append('%s: %sx%sx%s' % (sku, spec.get('length'), spec.get('width'), spec.get('height')))
        if not self.box_specs:
            results['dim_weight_check'] = {'pass': False, 'msg': '箱规库为空，无法校验尺寸重量', 'detail': []}
        elif overweight or oversize:
            results['dim_weight_check'] = {
                'pass': False,
                'msg': '存在 %d 个超重箱、%d 个超长箱，入库前需确认' % (len(overweight), len(oversize)),
                'detail': (overweight + oversize)[:50]}
        else:
            results['dim_weight_check'] = {
                'pass': True,
                'msg': '%d 个 SKU 箱规均满足 单箱≤30kg 且 单边≤100cm' % len(self.box_specs),
                'detail': []}

        return results
