import tempfile
import os
import re
from datetime import datetime
from openpyxl import load_workbook

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'templates', '請求書テンプレート.xlsx')

def parse_date(date_str):
    m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date_str)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return y - 2018, mo, d
    return None, None, None

def build_invoice(data: dict) -> str:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    reiwa, mo, d = parse_date(data.get('date', ''))
    if reiwa:
        ws['Z2'] = reiwa
        ws['AB2'] = mo
        ws['AD2'] = d
    ws['A3'] = data.get('client', '')
    ws['E6'] = data.get('site', '')
    ws['E7'] = data.get('due_date', '')
    items = data.get('items', [])
    for i, item in enumerate(items[:10]):
        row = 16 + i
        ws.cell(row=row, column=3, value=item['name'])
        ws.cell(row=row, column=26, value=item['qty'])
        ws.cell(row=row, column=28, value=item['price'])
    if data.get('note'):
        ws['A29'] = data['note']
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    client_safe = re.sub(r'[\\/:*?"<>|]', '', data.get('client', '請求先'))
    filename = f'請求書_{client_safe}_{timestamp}.xlsx'
    filepath = os.path.join(tempfile.gettempdir(), filename)
    wb.save(filepath)
    return filepath
