# -*- coding: utf-8 -*-
import openpyxl

wb_replenish = openpyxl.load_workbook(r'D:\桌面文件\虾皮补货文件\11.2\11.2虾皮补货.xlsx', data_only=True)
ws_my = wb_replenish['马来']

groups = []
current_group = None

for r in range(4, ws_my.max_row + 1):
    name = ws_my.cell(r, 1).value
    qty = ws_my.cell(r, 2).value
    seq = ws_my.cell(r, 3).value
    box_cnt = ws_my.cell(r, 4).value
    
    if not name and not qty:
        continue
        
    if seq is not None and str(seq).strip() != '':
        current_group = {
            'seq_str': str(seq).strip(),
            'box_cnt': int(box_cnt) if box_cnt else 1,
            'items': [(str(name).strip(), float(qty) if qty else 0)]
        }
        groups.append(current_group)
    else:
        if current_group:
            current_group['items'].append((str(name).strip(), float(qty) if qty else 0))

print(f'Total box groups parsed: {len(groups)}')
for g in groups:
    items_str = ', '.join([f'{it[0]}({it[1]:.0f})' for it in g['items']])
    print(f"Seq: {g['seq_str']:6s} | Boxes: {g['box_cnt']} | Items: {items_str}")
