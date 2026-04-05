import tempfile
import os
import re
from datetime import datetime
from openpyxl import load_workbook

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'templates', '給料明細テンプレート.xlsx')

def parse_period(period_str):
    m = re.match(r'(\d{4})年(\d{1,2})月', period_str)
    if m:
        return int(m.group(1)) - 2018, int(m.group(2))
    m2 = re.match(r'令和(\d+)年(\d{1,2})月', period_str)
    if m2:
        return int(m2.group(1)), int(m2.group(2))
    return None, None

def build_payslip(data: dict) -> str:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    ws['C2'] = data.get('name', '')
    ws['C3'] = data.get('emp_no', '')
    reiwa, mo = parse_period(data.get('period', ''))
    if reiwa:
        ws['D1'] = f'令和{reiwa}年{mo}月　給料支給明細書'
    ws['B6'] = data.get('work_days', '')
    ws['C6'] = data.get('attend_days', '')
    ws['D6'] = data.get('work_hours', '')
    ws['E6'] = data.get('holiday_work', '')
    ws['F6'] = data.get('paid_leave', '')
    ws['B8'] = data.get('overtime', '')
    ws['C8'] = data.get('holiday_work', '')
    ws['B11'] = data.get('base_salary', 0)
    ws['C11'] = data.get('overtime_pay', 0)
    ws['D11'] = data.get('night_pay', 0)
    ws['E11'] = data.get('director_pay', 0)
    ws['B14'] = data.get('pension', 0)
    ws['C14'] = data.get('employment_ins', 0)
    ws['D14'] = data.get('health_ins', 0)
    ws['E14'] = data.get('income_tax', 0)
    ws['F14'] = data.get('rent', 0)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    name_safe = re.sub(r'[\\/:*?"<>|]', '', data.get('name', '社員'))
    filename = f'給料明細_{name_safe}_{timestamp}.xlsx'
    filepath = os.path.join(tempfile.gettempdir(), filename)
    wb.save(filepath)
    return filepath
