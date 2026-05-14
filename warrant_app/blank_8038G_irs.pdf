import os
import io
import json
import base64
import zipfile
import re
import urllib.request
import urllib.error
import time
import threading
import uuid
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from scipy.optimize import brentq
from flask import Flask, request, jsonify, send_file, render_template, session, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import openpyxl
import pikepdf
import pandas as pd
from functools import wraps

# Load API key from .env file if present
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-change-in-production')

@app.errorhandler(413)
def request_entity_too_large(e):
    return jsonify({'error': 'PDF file is too large. Maximum size is 50MB.'}), 413

# ── Async job store for long-running warrant extraction ───────────────────────
# { job_id: { 'status': 'running'|'done'|'error', 'result': bytes|None, 'error': str|None, 'filename': str } }
_wrl_jobs = {}
_wrl_jobs_lock = threading.Lock()

# ── Password protection ───────────────────────────────────────────────────────
APP_PASSWORD = os.environ.get('APP_PASSWORD', 'bluestem2025')

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def load_sid_names():
    wb = openpyxl.load_workbook(os.path.join(os.path.dirname(__file__), 'SID_Names.xlsx'))
    ws = wb.active
    lookup = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] is not None:
            sid_num = int(row[1])
            if sid_num not in lookup:
                lookup[sid_num] = str(row[2]).strip() if row[2] else ''
    return lookup

SID_NAMES = load_sid_names()

# ── 8038 Info lookup ──────────────────────────────────────────────────────────
def load_8038_info():
    path = os.path.join(os.path.dirname(__file__), '8038_Info.xlsx')
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()
    return df

INFO_8038 = load_8038_info()

def get_sid_info(sid_num):
    """Return dict of 8038 info for a given SID number, or None."""
    rows = INFO_8038[INFO_8038['SID'] == sid_num]
    if rows.empty:
        return None
    r = rows.iloc[0]
    first_int = r['First Annual Interest Payment']
    if hasattr(first_int, 'month'):
        fi_month, fi_day = first_int.month, first_int.day
    else:
        fi_month, fi_day = 5, 1  # fallback
    raw_rate = r.get('Interest Rate', 7) if hasattr(r, 'get') else 7
    try:
        rate_val = float(raw_rate)
    except (TypeError, ValueError):
        rate_val = 7.0
    coupon = rate_val / 100 if rate_val > 1 else rate_val
    return {
        'county':    str(r['County']).strip(),
        'ein':       str(r['EIN']).strip(),
        'street':    str(r['Number and street (or P.O. box if mail is not delivered to street address)']).strip(),
        'room':      str(r['Room/Suite']).strip(),
        'city':      str(r['City, town, or post office, state, and ZIP code']).strip(),
        'name':      str(r['Name']).strip(),
        'role':      str(r['Chairman/Clerk']).strip(),
        'phone':     str(r['Phone']).strip(),
        'fi_month':  fi_month,
        'fi_day':    fi_day,
        'coupon':    coupon,
        'rate_pct':  int(rate_val),
    }

# ── 8038 Calculation helpers ──────────────────────────────────────────────────
def days_30_360(d1, d2):
    y1, m1, dd1 = d1.year, d1.month, d1.day
    y2, m2, dd2 = d2.year, d2.month, d2.day
    if dd1 == 31: dd1 = 30
    if dd2 == 31 and dd1 >= 30: dd2 = 30
    return 360*(y2-y1) + 30*(m2-m1) + (dd2-dd1)

def calc_wam_and_yield(reg_date, mat_date, face, coupon_rate, fi_month, fi_day):
    wam = days_30_360(reg_date, mat_date) / 360
    first_int = date(reg_date.year, fi_month, fi_day)
    if first_int <= reg_date:
        first_int = date(reg_date.year + 1, fi_month, fi_day)
    payments = []
    d = first_int
    while d < mat_date:
        payments.append(d)
        d = date(d.year + 1, d.month, d.day)
    cf = []
    prev = reg_date
    cum = 0
    for p in payments:
        dp = days_30_360(prev, p)
        cum += dp
        cf.append((cum, face * coupon_rate * dp / 360))
        prev = p
    dp_f = days_30_360(prev, mat_date)
    cum += dp_f
    cf.append((cum, face + face * coupon_rate * dp_f / 360))
    sr = brentq(lambda r: sum(pmt/(1+r)**(cd/180) for cd, pmt in cf) - face, 0.0001, 0.5)
    return round(wam, 3), round(sr * 2 * 100, 4)

def get_maturity_date(warrant_date, county, fund_type):
    """Sarpy/Cass: use Pro Dt from reg sheet. Douglas: warrant_date + 3/5 years."""
    years = 5 if fund_type == 'construction' else 3
    return warrant_date + relativedelta(years=years)

def identify_issuance_costs(warrants):
    """Return (issuance_cost_total, nonrefunding_total) by identifying last two payees."""
    UNDERWRITERS = {'bluestem capital partners', 'northland securities', 'northland securities inc',
                    'northland securities, inc.', 'northland securities, inc',
                    'ameritas investment company', 'ameritas investment', 'core bank',
                    'southstate|duncanwilliams', 'access bank',
                    'da davidson', 'd.a. davidson', 'd a davidson'}
    issuance = sum(w['amount'] for w in warrants
                   if any(u in w['payee'].lower() for u in UNDERWRITERS))
    total = sum(w['amount'] for w in warrants)
    return round(issuance, 2), round(total - issuance, 2)

# Canonical short → distinctive substring map used by the row-level query helper.
# Keys are lowercased short forms a user might type; values are substrings that
# appear in the Payee column for that entity. Key matching uses word boundaries,
# so "MUD" won't match "mudjacking". Values are case-insensitive substrings.
PAYEE_ALIASES = {
    'oppd': 'Omaha Public Power District',
    'mud': 'Metropolitan Utilities District',
    'ameritas': 'Ameritas Investment',
    'northland': 'Northland Securities',
    'bluestem': 'Bluestem Capital',
    'kutak': 'Kutak Rock',
    'kutak rock': 'Kutak Rock',
    'e&a': 'E & A Consulting',
    'e & a': 'E & A Consulting',
    'fullenkamp': 'Fullenkamp',
    'southstate': 'SouthState',
    'duncan williams': 'Duncan Williams',
    'access bank': 'Access Bank',
    'core bank': 'CORE BANK',
    'firnbank': 'FIRNBANK',
}

# Full/canonical names for each entity, used in the system-prompt alias list so
# the model sees readable mappings like "OPPD → Omaha Public Power District".
PAYEE_CANONICAL_NAMES = {
    'oppd': 'Omaha Public Power District',
    'mud': 'Metropolitan Utilities District',
    'ameritas': 'Ameritas Investment Company',
    'northland': 'Northland Securities, Inc.',
    'bluestem': 'Bluestem Capital Partners',
    'kutak': 'Kutak Rock LLP',
    'kutak rock': 'Kutak Rock LLP',
    'e&a': 'E & A Consulting Group, Inc.',
    'e & a': 'E & A Consulting Group, Inc.',
    'fullenkamp': 'Fullenkamp firm (various names: Fullenkamp, Doyle & Jobeun; Fullenkamp, Jobeun, Johnson & Beller; etc.)',
    'southstate': 'SouthState | Duncan Williams Securities Corp.',
    'duncan williams': 'SouthState | Duncan Williams Securities Corp.',
    'access bank': 'Access Bank',
    'core bank': 'CORE BANK',
    'firnbank': 'FIRNBANK CO.',
}

# ── Yield & WAM calculation workbooks ────────────────────────────────────────
def build_yield_wam_workbook(extracted, dt_wt, dt_reg):
    """Build yield/WAM calculation workbook. Returns (filename, bytes) or None."""
    from openpyxl import Workbook as WB
    from openpyxl.styles import Font as Fnt, Alignment as Aln, PatternFill, numbers
    from openpyxl.utils import get_column_letter

    sid_num = extracted['sid_number']
    cw = [w for w in extracted.get('construction_warrants', []) if not w.get('void')]
    gw = [w for w in extracted.get('general_warrants',      []) if not w.get('void')]
    info = get_sid_info(sid_num)
    if info is None:
        return None

    coupon   = info['coupon']
    fi_month = info['fi_month']
    fi_day   = info['fi_day']

    wb = WB()
    wb.remove(wb.active)

    hdr_fill = PatternFill('solid', fgColor='1F3864')
    hdr_font = Fnt(bold=True, color='FFFFFF', size=11)
    lbl_font = Fnt(bold=True, size=10)
    val_font = Fnt(size=10)
    bold10   = Fnt(bold=True, size=10)
    red_bold = Fnt(bold=True, size=11, color='FF0000')

    def cw_(ws, col, w):
        ws.column_dimensions[col].width = w

    def hdr_cell(ws, cell, val):
        ws[cell].value     = val
        ws[cell].fill      = hdr_fill
        ws[cell].font      = hdr_font
        ws[cell].alignment = Aln(horizontal='center', vertical='center')

    def get_first_int(reg_date):
        fi = date(reg_date.year, fi_month, fi_day)
        if fi <= reg_date:
            fi = date(reg_date.year + 1, fi_month, fi_day)
        return fi

    def build_cf_sheet(ws, total, mat_date):
        # Column widths matching example exactly
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 2
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 2
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 2
        ws.column_dimensions['G'].width = 14
        ws.column_dimensions['H'].width = 2
        ws.column_dimensions['I'].width = 14

        ws.merge_cells('A1:I1')
        ws['A1'].value     = 'CONSTRUCTION FUND WARRANT YIELD CALCULATION'
        ws['A1'].font      = Fnt(bold=True, size=12)
        ws['A1'].alignment = Aln(horizontal='center')

        # Inputs — rows 3,5,7,9,11 with blank rows between
        fi_date = get_first_int(dt_reg)
        inputs = [
            (3,  'Principal Amount',            'C3',  total,    '#,##0.00'),
            (5,  'Registration Date',           'C5',  dt_reg,   'mm/dd/yyyy'),
            (7,  'Maturity Date',               'C7',  mat_date, 'mm/dd/yyyy'),
            (9,  'Interest Rate',               'C9',  coupon,   '0.00%'),
            (11, '1st Annual Interest Payment', 'C11', fi_date,  'mm/dd/yyyy'),
        ]
        for row, lbl, vcell, val, fmt in inputs:
            ws[f'A{row}'].value = lbl;  ws[f'A{row}'].font = lbl_font
            ws[vcell].value = val;      ws[vcell].font = val_font
            ws[vcell].number_format = fmt

        # YIELD label and formula in I3 (formula set after schedule is built)
        ws['G3'].value = 'YIELD:'; ws['G3'].font = bold10
        ws['I3'].font  = red_bold
        ws['I3'].number_format = '0.0000%'

        # Column headers row 13
        for cell, lbl in [('A13','Payment Date'),('C13','Days'),
                           ('E13','Days(Cumulative)'),('G13','Payments'),('I13','PV')]:
            hdr_cell(ws, cell, lbl)

        # Build payment schedule
        pay_dates = [dt_reg]
        d = fi_date
        while d < mat_date:
            pay_dates.append(d); d = date(d.year+1, d.month, d.day)
        pay_dates.append(mat_date)

        hr = brentq(lambda r: sum(
            (total * coupon * days_30_360(pay_dates[i-1], pay_dates[i]) / 360
             if i < len(pay_dates)-1
             else total + total * coupon * days_30_360(pay_dates[i-1], pay_dates[i]) / 360)
            / (1+r)**(sum(days_30_360(pay_dates[j-1], pay_dates[j])
                          for j in range(1, i+1)) / 180)
            for i in range(1, len(pay_dates))
        ) - total, 0.0001, 0.5)

        row = 14
        prev = dt_reg; cum = 0
        for i, pd_ in enumerate(pay_dates):
            d_days = days_30_360(prev, pd_) if i > 0 else 0
            cum   += d_days
            if i == 0:
                pmt, pv = -total, 0
            elif i < len(pay_dates)-1:
                pmt = total * coupon * d_days / 360
                pv  = pmt / (1+hr)**(cum/180)
            else:
                pmt = total + total * coupon * d_days / 360
                pv  = pmt / (1+hr)**(cum/180)
            ws[f'A{row}'].value = pd_; ws[f'A{row}'].number_format = 'mm/dd/yyyy'; ws[f'A{row}'].font = val_font
            ws[f'C{row}'].value = d_days; ws[f'C{row}'].font = val_font
            ws[f'E{row}'].value = cum;    ws[f'E{row}'].font = val_font
            ws[f'G{row}'].value = round(pmt, 8); ws[f'G{row}'].number_format = '#,##0.0000'; ws[f'G{row}'].font = val_font
            ws[f'I{row}'].value = round(pv,  8); ws[f'I{row}'].number_format = '#,##0.0000'; ws[f'I{row}'].font = val_font
            prev = pd_; row += 1

        # Totals row (immediately after last payment row)
        tot_row = row
        ws[f'G{tot_row}'].value = f'=SUM(G15:G{tot_row-1})'; ws[f'G{tot_row}'].number_format = '#,##0.0000'; ws[f'G{tot_row}'].font = bold10
        ws[f'I{tot_row}'].value = f'=SUM(I15:I{tot_row-1})'; ws[f'I{tot_row}'].number_format = '#,##0.0000'; ws[f'I{tot_row}'].font = bold10
        row += 1

        # Goal seek block: skip a row between each item (matching example rows 23,25,27,28)
        row += 1  # blank row after totals
        ws[f'I{row}'].value = '180/360 Yield Target'; ws[f'I{row}'].font = lbl_font
        row += 2  # skip a row
        target_row = row
        ws[f'I{row}'].value = hr; ws[f'I{row}'].number_format = '0.00000000'; ws[f'I{row}'].font = val_font
        row += 2  # skip a row
        ws[f'I{row}'].value = 'PV Difference'; ws[f'I{row}'].font = lbl_font
        row += 1  # immediately below PV Difference
        ws[f'I{row}'].value = 0; ws[f'I{row}'].font = Fnt(size=11)

        # Set yield formula in I3 referencing the target row
        ws['I3'].value = f'=I{target_row}*2'

    def build_gf_sheet(ws, total, mat_date):
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['G'].width = 14
        ws.column_dimensions['I'].width = 14

        ws.merge_cells('A1:I1')
        ws['A1'].value     = 'GENERAL FUND WARRANT YIELD CALCULATION'
        ws['A1'].font      = Fnt(bold=True, size=12)
        ws['A1'].alignment = Aln(horizontal='center')

        inputs = [
            (3, 'Principal Amount',  'C3', total,    '#,##0.00'),
            (4, 'Registration Date', 'C4', dt_reg,   'mm/dd/yyyy'),
            (6, 'Maturity Date',     'C6', mat_date, 'mm/dd/yyyy'),
            (8, 'Interest Rate',     'C8', coupon,   '0.00%'),
        ]
        for row, lbl, vcell, val, fmt in inputs:
            ws[f'A{row}'].value = lbl;  ws[f'A{row}'].font = lbl_font
            ws[vcell].value = val;      ws[vcell].font = val_font
            ws[vcell].number_format = fmt

        wam, yld = calc_wam_and_yield(dt_reg, mat_date, total, coupon, fi_month, fi_day)
        ws['G3'].value = 'YIELD:'; ws['G3'].font = bold10
        ws['I3'].value = yld / 100
        ws['I3'].font  = red_bold
        ws['I3'].number_format = '0.0000%'

        for cell, lbl in [('A10','Payment Date'),('C10','Days'),
                           ('E10','Days(Cumulative)'),('G10','Payments'),('I10','PV')]:
            hdr_cell(ws, cell, lbl)

        d_days    = days_30_360(dt_reg, mat_date)
        pmt_final = total + total * coupon * d_days / 360
        hr        = brentq(lambda r: pmt_final / (1+r)**(d_days/180) - total, 0.0001, 0.5)
        pv_final  = pmt_final / (1+hr)**(d_days/180)

        row = 11
        for r_date, dd, cum, pmt, pv in [
            (dt_reg,   0,      0,      -total,                 0),
            (mat_date, d_days, d_days, round(pmt_final, 8),    round(pv_final, 8)),
        ]:
            ws[f'A{row}'].value = r_date; ws[f'A{row}'].number_format = 'mm/dd/yyyy'; ws[f'A{row}'].font = val_font
            ws[f'C{row}'].value = dd;     ws[f'C{row}'].font = val_font
            ws[f'E{row}'].value = cum;    ws[f'E{row}'].font = val_font
            ws[f'G{row}'].value = pmt;    ws[f'G{row}'].number_format = '#,##0.0000'; ws[f'G{row}'].font = val_font
            ws[f'I{row}'].value = pv;     ws[f'I{row}'].number_format = '#,##0.0000'; ws[f'I{row}'].font = val_font
            row += 1

        tot_row = row
        ws[f'G{tot_row}'].value = f'=SUM(G12:G{tot_row-1})'; ws[f'G{tot_row}'].number_format = '#,##0.0000'; ws[f'G{tot_row}'].font = bold10
        ws[f'I{tot_row}'].value = f'=SUM(I12:I{tot_row-1})'; ws[f'I{tot_row}'].number_format = '#,##0.0000'; ws[f'I{tot_row}'].font = bold10
        row += 1
        ws[f'I{row}'].value = 'Target';     ws[f'I{row}'].font = lbl_font; row += 1
        ws[f'I{row}'].value = hr;           ws[f'I{row}'].number_format = '0.00000000'; ws[f'I{row}'].font = val_font; row += 1
        ws[f'I{row}'].value = 'Difference'; ws[f'I{row}'].font = lbl_font

    def build_combined_sheet(ws, cf_total, gf_total, cf_mat, gf_mat):
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['E'].width = 16
        ws.column_dimensions['G'].width = 14

        ws.merge_cells('A1:G1')
        ws['A1'].value     = 'COMBINED CONSTRUCTION FUND & GENERAL FUND YIELD CALCULATION'
        ws['A1'].font      = Fnt(bold=True, size=11)
        ws['A1'].alignment = Aln(horizontal='center')

        for cell, val in [('C2','Principal'),('C3','Amount'),
                           ('E2','Yield'),   ('E3','Calculation'),
                           ('G2','Percentage'),('G3','Weighted')]:
            ws[cell].value = val; ws[cell].font = bold10
            ws[cell].alignment = Aln(horizontal='center')

        _, yld_cf = calc_wam_and_yield(dt_reg, cf_mat, cf_total, coupon, fi_month, fi_day)
        _, yld_gf = calc_wam_and_yield(dt_reg, gf_mat, gf_total, coupon, fi_month, fi_day)
        total_all = cf_total + gf_total
        pct_cf = cf_total / total_all if total_all > 0 else 0
        pct_gf = gf_total / total_all if total_all > 0 else 0

        for row, lbl, prin, yld_val, pct in [
            (4, 'GF Warrants', gf_total, yld_gf/100, pct_gf * yld_gf/100),
            (5, 'CF Warrants', cf_total, yld_cf/100, pct_cf * yld_cf/100),
        ]:
            ws[f'A{row}'].value = lbl;      ws[f'A{row}'].font = lbl_font
            ws[f'C{row}'].value = prin;     ws[f'C{row}'].number_format = '#,##0.00'; ws[f'C{row}'].font = val_font
            ws[f'E{row}'].value = yld_val;  ws[f'E{row}'].number_format = '0.0000%';  ws[f'E{row}'].font = val_font
            ws[f'G{row}'].value = pct;      ws[f'G{row}'].number_format = '0.0000%';  ws[f'G{row}'].font = val_font

        ws['A6'].value = 'TOTAL'; ws['A6'].font = bold10
        ws['C6'].value = total_all; ws['C6'].number_format = '#,##0.00'; ws['C6'].font = bold10
        ws['E7'].value = 'YIELD:'; ws['E7'].font = bold10
        ws['G7'].value = (pct_cf * yld_cf/100 + pct_gf * yld_gf/100)
        ws['G7'].number_format = '0.0000%'
        ws['G7'].font = red_bold

    # ── Sheet order: CF GoalSeek, Combined, GF GoalSeek ──────────────────────
    c_total, g_total = None, None
    mat_c,   mat_g   = None, None

    if cw:
        c_total = round(sum(w['amount'] for w in cw), 2)
        mat_c   = dt_wt + relativedelta(years=5)
        ws_cf   = wb.create_sheet('GoalSeek Construction Fund')
        build_cf_sheet(ws_cf, c_total, mat_c)

    # Only CF GoalSeek sheet is included in the yield workbook

    if not wb.sheetnames:
        return None

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return f'Yield_Calc-Goal_Seek_{sid_num}.xlsx', buf.read()


def build_wam_workbook(extracted, dt_wt, dt_reg):
    """Build remaining weighted average maturity workbook. Returns (filename, bytes) or None."""
    from openpyxl import Workbook as WB
    from openpyxl.styles import Font as Fnt, Alignment as Aln, PatternFill

    sid_num = extracted['sid_number']
    cw = [w for w in extracted.get('construction_warrants', []) if not w.get('void')]
    gw = [w for w in extracted.get('general_warrants',      []) if not w.get('void')]
    info = get_sid_info(sid_num)
    if info is None or not cw:
        return None

    coupon = info['coupon']
    wb = WB()
    wb.remove(wb.active)

    hdr_fill = PatternFill('solid', fgColor='1F3864')
    hdr_font = Fnt(bold=True, color='FFFFFF', size=10)
    lbl_font = Fnt(bold=True, size=10)
    val_font = Fnt(size=10)
    bold10   = Fnt(bold=True, size=10)

    def build_wam_sheet(ws, warrants, fund_type, mat_years):
        mat_date = dt_wt + relativedelta(years=mat_years)
        total    = round(sum(w['amount'] for w in warrants), 2)
        wam, _   = calc_wam_and_yield(dt_reg, mat_date, total, coupon,
                                       info['fi_month'], info['fi_day'])

        label = 'CONSTRUCTION FUND' if fund_type == 'cf' else 'GENERAL FUND'
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['E'].width = 14
        ws.column_dimensions['F'].width = 18

        ws.merge_cells('A1:F1')
        ws['A1'].value     = f'REMAINING WEIGHTED AVERAGE MATURITY — {label} WARRANTS'
        ws['A1'].font      = Fnt(bold=True, size=11)
        ws['A1'].alignment = Aln(horizontal='center')

        # Info rows
        for r, lbl, val, fmt in [
            (3, 'Registration Date', dt_reg,   'mm/dd/yyyy'),
            (4, 'Maturity Date',     mat_date,  'mm/dd/yyyy'),
            (5, 'Total Principal',   total,     '#,##0.00'),
            (6, 'WAM (years)',       wam,       '0.000'),
        ]:
            ws[f'A{r}'].value = lbl; ws[f'A{r}'].font = lbl_font
            ws[f'C{r}'].value = val; ws[f'C{r}'].font = val_font
            ws[f'C{r}'].number_format = fmt

        # Column headers
        for col, hdr in [('A','Warrant #'),('B','Issue Date'),('C','Maturity Date'),
                         ('D','Face Amount'),('E','Days to Mat'),('F','Weight')]:
            ws[f'{col}8'].value = hdr
            ws[f'{col}8'].fill  = hdr_fill
            ws[f'{col}8'].font  = hdr_font
            ws[f'{col}8'].alignment = Aln(horizontal='center')

        # Warrant rows
        row = 9
        for w in sorted(warrants, key=lambda x: x['warrant_number']):
            d_days = days_30_360(dt_reg, mat_date)
            weight = (w['amount'] / total) * d_days / 360 if total > 0 else 0
            ws[f'A{row}'].value = w['warrant_number']; ws[f'A{row}'].font = val_font; ws[f'A{row}'].alignment = Aln(horizontal='center')
            ws[f'B{row}'].value = dt_wt;  ws[f'B{row}'].number_format = 'mm/dd/yyyy'; ws[f'B{row}'].font = val_font
            ws[f'C{row}'].value = mat_date; ws[f'C{row}'].number_format = 'mm/dd/yyyy'; ws[f'C{row}'].font = val_font
            ws[f'D{row}'].value = w['amount']; ws[f'D{row}'].number_format = '#,##0.00'; ws[f'D{row}'].font = val_font
            ws[f'E{row}'].value = d_days; ws[f'E{row}'].font = val_font
            ws[f'F{row}'].value = round(weight, 6); ws[f'F{row}'].number_format = '0.000000'; ws[f'F{row}'].font = val_font
            row += 1

        # Totals
        ws[f'D{row}'].value = f'=SUM(D9:D{row-1})'; ws[f'D{row}'].number_format = '#,##0.00'; ws[f'D{row}'].font = bold10
        ws[f'F{row}'].value = f'=SUM(F9:F{row-1})'; ws[f'F{row}'].number_format = '0.000000'; ws[f'F{row}'].font = bold10
        row += 2
        ws[f'A{row}'].value = 'Weighted Average Maturity (years):'; ws[f'A{row}'].font = bold10
        ws.merge_cells(f'A{row}:C{row}')
        ws[f'D{row}'].value = wam; ws[f'D{row}'].number_format = '0.000'; ws[f'D{row}'].font = Fnt(bold=True, size=11)

    if cw:
        ws_cf = wb.create_sheet('Construction Fund WAM')
        build_wam_sheet(ws_cf, cw, 'cf', 5)
    if gw:
        ws_gf = wb.create_sheet('General Fund WAM')
        build_wam_sheet(ws_gf, gw, 'gf', 3)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return f'Remaining_Weighted_Average_Maturity_{sid_num}.xlsx', buf.read()


# ── 8038 PDF fill functions ───────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(__file__)
TMPL_G        = os.path.join(BASE_DIR, 'blank_8038G_irs.pdf')
TMPL_GC       = os.path.join(BASE_DIR, 'blank_8038GC_irs.pdf')
WARRANTS_PATH = os.path.join(BASE_DIR, 'Warrants_2014-2026.xlsx')


def _load_warrants_df() -> tuple:
    """
    Load the warrants history file.
    Row 1 (index 0) may contain an 'as-of' date in cell A1 — if so, skip it.
    Returns (df, date_str) where date_str is the file date or None.
    """
    if not os.path.exists(WARRANTS_PATH):
        return None, None
    try:
        # Peek at A1 to check for date
        wb_peek = openpyxl.load_workbook(WARRANTS_PATH, read_only=True, data_only=True)
        ws_peek = wb_peek.active
        a1 = ws_peek.cell(row=1, column=1).value
        wb_peek.close()

        if isinstance(a1, datetime) or (hasattr(a1, 'year') and not isinstance(a1, str)):
            # A1 is a date — header is on row 2, data from row 3
            file_date = a1.strftime('%B %-d, %Y') if hasattr(a1, 'strftime') else str(a1)
            df = pd.read_excel(WARRANTS_PATH, header=1)
        else:
            # No date prefix — header is row 1
            file_date = None
            df = pd.read_excel(WARRANTS_PATH)

        df.columns = [c.strip() for c in df.columns]
        return df, file_date
    except Exception:
        return None, None



PREPARER = {
    'name':    'Joshua P. Meyer, Esq.',
    'ptin':    'P01778308',
    'firm':    'Kutak Rock LLP',
    'address': '1650 Farnam Street, Omaha, NE 68102',
    'ein':     '47:0597598',
    'phone':   '402-346-6000',
}

def _xfa_set(xml, field_name, value):
    sc = f'<{field_name}\n/>'
    ot = f'<{field_name}\n>'
    ct = f'</{field_name}\n>'
    if sc in xml:
        return xml.replace(sc, f'{ot}{value}{ct}')
    elif ot in xml:
        sp = xml.find(ot) + len(ot)
        ep = xml.find(ct, sp)
        if ep > sp:
            return xml[:sp] + value + xml[ep:]
    # Field not present — inject before closing tag
    return xml.replace('</topmostSubform\n>', f'{ot}{value}{ct}</topmostSubform\n>')

def fill_8038G(output_fields, check_box39=True):
    """Fill 8038-G and return bytes. check_box39=False unchecks small issuer exception."""
    pdf = pikepdf.open(TMPL_G)
    xfa = pdf.Root.AcroForm.XFA
    for i in range(0, len(xfa), 2):
        if str(xfa[i]) == 'datasets':
            ds = xfa[i+1]
            break
    xml = ds.read_bytes().decode('utf-8', errors='replace')
    for k, v in output_fields.items():
        xml = _xfa_set(xml, k, v)
    # Checkboxes
    xml = _xfa_set(xml, 'c2_2', '1' if check_box39 else '0')  # box 39
    xml = _xfa_set(xml, 'c2_6', '1')  # box 43
    xml = _xfa_set(xml, 'c2_7', '1')  # box 44
    ds.write(xml.encode('utf-8'))
    buf = io.BytesIO()
    pdf.save(buf)
    return buf.getvalue()

def fill_8038GC(output_fields, check_box10=True):
    """
    Fill the single-page AcroForm 8038-GC template and return PDF bytes.
    check_box10=True: checks Single Issue box (c1_2/AP key /1) and Line 10 box (c1_3).

    The form has two c1_2 widgets:
      - Single issue   checkbox: AP/N key = /1  → check when check_box10=True
      - Consolidated   checkbox: AP/N key = /2  → always unchecked (we file single issues)
    """
    pdf = pikepdf.open(TMPL_GC)

    def decode_field_name(t_raw):
        b = bytes(t_raw)
        s = b[2:].decode('utf-16-be') if b[:2] == b'\xfe\xff' else b.decode('latin-1')
        return re.sub(r'\[\d+\]$', '', s)

    def get_ap_on_key(annot):
        """Return the /AP /N appearance key that means 'checked' for this widget."""
        ap = annot.get('/AP')
        if ap:
            n = ap.get('/N')
            if n:
                try:
                    keys = [str(k) for k in n.keys()]
                    # Return the first key that isn't /Off
                    for k in keys:
                        if k != '/Off':
                            return k
                except Exception:
                    pass
        return '/1'  # fallback

    for page in pdf.pages:
        if '/Annots' not in page:
            continue
        for aref in page.Annots:
            annot = aref
            if annot.get('/Subtype') != pikepdf.Name('/Widget'):
                continue
            t_raw = annot.get('/T')
            if t_raw is None:
                continue
            try:
                name = decode_field_name(t_raw)
            except Exception:
                continue

            ft = annot.get('/FT')
            if ft == pikepdf.Name('/Btn'):
                on_key = get_ap_on_key(annot)
                if name == 'c1_2':
                    # Single Issue = on_key /1, Consolidated = on_key /2
                    # Check Single Issue when check_box10=True, leave Consolidated unchecked
                    checked = check_box10 and on_key == '/1'
                elif name == 'c1_3':
                    # Line 10 small issuer exception — check when check_box10=True
                    checked = check_box10
                else:
                    checked = False
                pdf_val = pikepdf.Name(on_key) if checked else pikepdf.Name('/Off')
                annot['/V']  = pdf_val
                annot['/AS'] = pdf_val
            elif name in output_fields:
                annot['/V'] = pikepdf.String(str(output_fields[name]))

    pdf.Root.AcroForm['/NeedAppearances'] = pikepdf.Boolean(True)
    buf = io.BytesIO()
    pdf.save(buf)
    return buf.getvalue()


def build_combined_8038_pdf(forms):
    """
    Flatten and merge all 8038 PDFs into a single static document.
    forms = list of (filename, pdf_bytes, extra) tuples where extra holds
    fields+flags for 8038-G, or None for 8038-GC.
    Returns combined PDF bytes.
    """
    from pypdf import PdfWriter, PdfReader
    writer = PdfWriter()
    for item in forms:
        fname, pdf_bytes = item[0], item[1]
        if '8038G_CF' in fname:
            # XFA form — use reportlab overlay approach
            extra = item[2] if len(item) > 2 else {}
            fields = extra.get('fields', {})
            check_box39 = extra.get('check_box39', True)
            flat = _flatten_xfa_pdf(fields, check_box39, True)
        else:
            # AcroForm (8038-GC) — use qpdf
            flat = _flatten_acroform_pdf(pdf_bytes)
        reader = PdfReader(io.BytesIO(flat))
        for page in reader.pages:
            writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()

# ── Main 8038 builder ─────────────────────────────────────────────────────────
def build_8038_forms(extracted, dt_wt, dt_reg):
    """
    Build all required 8038 forms for a given reg sheet.
    Returns list of (filename, pdf_bytes).
    """
    sid_num  = extracted['sid_number']
    county   = extracted.get('county', 'Sarpy').strip()
    info     = get_sid_info(sid_num)
    if info is None:
        return []  # No 8038 info — skip silently

    cw = [w for w in extracted.get('construction_warrants', []) if not w.get('void')]
    gw = [w for w in extracted.get('general_warrants',      []) if not w.get('void')]

    issuer   = f'Sanitary and Improvement District No. {sid_num} of {info["county"]} County, Nebraska'
    officer  = f'{info["name"]}, {info["role"]}'
    spaces   = ' ' * 160
    # SIDs 636, 637, 638 in Douglas County: boxes unchecked when warrants issued in 2026
    uncheck_cf_boxes = (sid_num in (636, 637, 638) and dt_wt.year == 2026)
    # SIDs 603, 623, 630 (Douglas) and 325 (Sarpy): preparer section left blank
    NO_PREPARER_SIDS = {603, 623, 630, 325}
    include_preparer = sid_num not in NO_PREPARER_SIDS

    results = []

    # ── Construction fund ────────────────────────────────────────────────────
    if cw:
        c_total = round(sum(w['amount'] for w in cw), 2)
        # Maturity: Sarpy/Cass use dt_wt + 5yr; Douglas also + 5yr
        mat_c = dt_wt + relativedelta(years=5)

        # For Sarpy/Cass, Pro Dt is already computed by the Excel builder as dt_wt + 5yr
        # (same logic). Douglas also uses warrant date + 5yr.

        wt_nums  = sorted(w['warrant_number'] for w in cw)
        wt_range = f'{wt_nums[0]}/{wt_nums[-1]}'
        issuance, nonrefunding = identify_issuance_costs(cw)
        wam, yld = calc_wam_and_yield(dt_reg, mat_c, c_total, info['coupon'],
                                       info['fi_month'], info['fi_day'])

        if c_total >= 100_000:
            # 8038-G for large CF issuances
            check_box39 = not uncheck_cf_boxes
            fields = {
                'f1_01': issuer, 'f1_02': info['ein'],
                'f1_05': info['street'], 'f1_06': info['room'],
                'f1_07': info['city'],
                'f1_08': dt_reg.strftime('%m/%d/%Y'),
                'f1_09': f'Construction Fund Warrant Nos. {wt_range}',
                'f1_11': officer, 'f1_12': info['phone'],
                'f1_13': 'N/A','f1_14': 'N/A','f1_15': 'N/A','f1_16': 'N/A',
                'f1_17': 'N/A','f1_18': 'N/A','f1_19': 'N/A',
                'f1_20': 'Various Public Improvements',
                'f1_21': f'${c_total:,.2f}',
                'f1_22': mat_c.strftime('%m/%d/%Y'),
                'f1_23': f'${c_total:,.2f}', 'f1_24': f'${c_total:,.2f}',
                'f1_25': str(wam), 'f1_26': str(yld),
                'f1_27': 'N/A', 'f1_28': f'${c_total:,.2f}',
                'f1_29': f'${issuance:,.2f}',
                'f1_30': '$0.00','f1_31': '$0.00','f1_32': '$0.00','f1_33': '$0.00',
                'f1_34': f'${issuance:,.2f}', 'f1_35': f'${nonrefunding:,.2f}',
                'f1_36': 'N/A','f1_37': 'N/A','f1_38': 'N/A','f1_39': 'N/A',
                'f2_01': '$0.00','f2_02': '$0.00','f2_05': '$0.00',
                'f2_14': officer,
                'f2_15': PREPARER['name'] if include_preparer else '',
                'f2_16': PREPARER['ptin'] if include_preparer else '',
                'f2_17': PREPARER['firm'] if include_preparer else '',
                'f2_18': PREPARER['address'] if include_preparer else '',
                'f2_19': PREPARER['ein'] if include_preparer else '',
                'f2_20': PREPARER['phone'] if include_preparer else '',
            }
            pdf_bytes = fill_8038G(fields, check_box39=check_box39)
            results.append((f'SID{sid_num}_8038G_CF_{dt_reg.strftime("%m%d%Y")}.pdf', pdf_bytes,
                            {'fields': fields, 'check_box39': check_box39}))
        else:
            # 8038-GC for small CF issuances (under $100k)
            fields_gc_cf = {
                'f1_01': issuer, 'f1_02': info['ein'],
                'f1_03': info['street'], 'f1_04': info['room'],
                'f1_05': info['city'],
                'f1_06': officer, 'f1_07': info['phone'],
                'f1_08': f'{c_total:,.2f}', 'f1_09': dt_reg.strftime('%m/%d/%Y'),
                'f1_10': 'N/A','f1_11': 'N/A','f1_12': 'N/A','f1_13': 'N/A',
                'f1_14': 'N/A','f1_15': 'N/A','f1_16': 'N/A','f1_17': 'N/A',
                'f1_18': 'N/A','f1_19': 'N/A',
                'f1_20': f'{c_total:,.2f}',
                'f1_21': spaces + 'N/A',
                'f1_22': 'N/A',
                'f1_23': officer,
                'f1_24': PREPARER['name'] if include_preparer else '',
                'f1_25': PREPARER['ptin'] if include_preparer else '',
                'f1_26': PREPARER['firm'] if include_preparer else '',
                'f1_27': PREPARER['ein'] if include_preparer else '',
                'f1_28': PREPARER['address'] if include_preparer else '',
                'f1_29': PREPARER['phone'] if include_preparer else '',
            }
            pdf_bytes = fill_8038GC(fields_gc_cf, check_box10=not uncheck_cf_boxes)
            results.append((f'SID{sid_num}_8038GC_CF_{dt_reg.strftime("%m%d%Y")}.pdf', pdf_bytes))

    # ── General fund ─────────────────────────────────────────────────────────
    if gw:
        g_total  = round(sum(w['amount'] for w in gw), 2)
        wt_nums  = sorted(w['warrant_number'] for w in gw)
        wt_range = f'{wt_nums[0]}/{wt_nums[-1]}'

        # Box 10: always checked for general fund (even SID 638)
        check_box10 = True

        fields_gc = {
            'f1_01': issuer, 'f1_02': info['ein'],
            'f1_03': info['street'], 'f1_04': info['room'],
            'f1_05': info['city'],
            'f1_06': officer, 'f1_07': info['phone'],
            'f1_08': f'{g_total:,.2f}', 'f1_09': dt_reg.strftime('%m/%d/%Y'),
            'f1_10': 'N/A','f1_11': 'N/A','f1_12': 'N/A','f1_13': 'N/A',
            'f1_14': 'N/A','f1_15': 'N/A','f1_16': 'N/A','f1_17': 'N/A',
            'f1_18': 'N/A','f1_19': 'N/A',
            'f1_20': f'{g_total:,.2f}',
            'f1_21': spaces + 'N/A',
            'f1_22': 'N/A',  # Line 13 vendor EIN always N/A
            'f1_23': officer,
            # GF 8038-GC: preparer section always left blank
        }
        pdf_bytes = fill_8038GC(fields_gc, check_box10=check_box10)
        results.append((f'SID{sid_num}_8038GC_GF_{dt_reg.strftime("%m%d%Y")}.pdf', pdf_bytes))

    return results

def _anthropic_request(payload: dict) -> str:
    api_key = os.environ.get('ANTHROPIC_API_KEY', '')
    if not api_key:
        raise ValueError('ANTHROPIC_API_KEY not set')
    req = urllib.request.Request(
        'https://api.anthropic.com/v1/messages',
        data=json.dumps(payload).encode('utf-8'),
        headers={
            'Content-Type': 'application/json',
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01',
            'anthropic-beta': 'pdfs-2024-09-25'
        },
        method='POST'
    )
    for _attempt in range(5):
        try:
            with urllib.request.urlopen(req, timeout=240) as resp:
                data = json.loads(resp.read().decode('utf-8'))
            break
        except urllib.error.HTTPError as e:
            # 429: rate-limited. 529: provider overloaded. 500/502/503/504: transient
            # infrastructure failures. All are worth retrying.
            if e.code in (429, 500, 502, 503, 504, 529) and _attempt < 4:
                # Shorter backoff for 5xx (infra, not rate limit): 2s, 5s, 10s, 20s.
                # 429/529 use the longer 60s backoff since they reflect capacity limits.
                if e.code in (429, 529):
                    time.sleep(60 * (_attempt + 1))
                else:
                    time.sleep([2, 5, 10, 20][_attempt])
                continue
            # Surface the API's error body so callers see the actual reason (e.g. context too long)
            try:
                err_body = e.read().decode('utf-8', errors='replace')
            except Exception:
                err_body = ''
            raise urllib.error.HTTPError(
                e.url, e.code, f"{e.reason} — {err_body}" if err_body else e.reason,
                e.headers, None
            )
    text = data['content'][0]['text'].strip()
    if text.startswith('```'):
        text = text.split('```')[1]
        if text.startswith('json'):
            text = text[4:]
    return text.strip()

def call_anthropic(pdf_b64: str) -> dict:
    system = """You are a precise data extraction assistant. Extract warrant data from SID (Sanitary Improvement District) meeting minutes PDFs.

Return ONLY valid JSON with this exact structure:
{
  "sid_number": <integer>,
  "county": "<string - exact county name as stated, e.g. 'Sarpy' or 'Douglas'>",
  "general_warrants": [
    {"warrant_number": <int>, "amount": <float>, "payee": "<string>", "void": false}
  ],
  "construction_warrants": [
    {"warrant_number": <int>, "amount": <float>, "payee": "<string>", "void": false}
  ]
}

Rules:
- Extract the SID number from "District No. XXX" — copy ALL digits exactly as written, do not truncate (e.g. "District No. 363" -> 363, "District No. 1005" -> 1005)
- Double-check the SID number by looking for it in multiple places in the document (title, resolution text, etc.) to confirm accuracy
- Extract the county from the resolution text (e.g. "Sarpy County" -> "Sarpy", "Douglas County" -> "Douglas")
- Include VOID warrants with "void": true, amount: 0, payee: ""
- Non-void warrants: "void": false
- Preserve exact payee names as written in the document
- Expand common abbreviations in payee names: "MUD" -> "Metropolitan Utilities District", "OPPD" -> "Omaha Public Power District"
- Return ONLY the JSON object, no markdown, no explanation

CRITICAL RULE FOR GROUPED WARRANTS:
The warrant section often lists groups like: "Warrant Nos. 262 through 273, inclusive, each for $50,000.00 and Warrant No. 274 for $26,541.68 all payable to TAB Construction."
You MUST expand these into individual warrant entries. Each warrant gets its OWN individual amount:
- Wt#262: $50,000.00
- Wt#263: $50,000.00
- ... (one entry per warrant number)
- Wt#274: $26,541.68  <- the final warrant gets its own stated amount, NOT the group total

NEVER assign a group total to a single warrant. The dollar figure after "and Warrant No. X for $Y" is ALWAYS the amount for that specific last warrant only.
The consolidated narrative section (earlier in the document) may show a combined total for the group — ignore those totals for individual warrant amounts; use only the itemized warrant section."""

    text = _anthropic_request({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 4000,
        "system": system,
        "messages": [{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
            {"type": "text", "text": "Extract all warrant data from this SID meeting minutes document and return as JSON. IMPORTANT: Return a complete valid JSON object — do not truncate."}
        ]}]
    })
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"Response was truncated or malformed (try again — large PDFs occasionally hit limits). Detail: {e}")

def call_anthropic_validate(pdf_b64: str, extracted: dict) -> list:
    system = """You are a meticulous financial data auditor reviewing SID warrant registration data.

ABSOLUTE PRE-FLIGHT RULE — read this before doing anything else:
  Before you output a flag, look at the flag's message text. Find every pair of dollar amounts the message itself compares ("X vs Y", "should be X not Y", "shows X but extracted sum is Y", etc.). For each pair, strip thousands-separator commas and trailing-zero differences and parse the numbers. If the two numbers in the pair are mathematically equal, the flag is invalid and you MUST suppress it entirely. Do not output the flag, do not reword it, do not include a "note" version. Just drop it.

  This applies to formatting-only differences too: "$1,112.61" and "$1112.61" are the same number. "$361,112.61" and "$361112.61" are the same number. "$1,000" and "$1,000.00" are the same number. Comma placement and decimal trailing zeros never constitute a discrepancy.

  If after stripping formatting both numbers are equal, there is NO discrepancy and NO flag. Period.

BEFORE generating any flag, you MUST complete a silent verification step:
  - Compute the expected value yourself from the PDF
  - Compare it to the extracted value
  - Only generate a flag if there is a REAL discrepancy greater than $0.05
  - If the values match, produce NO flag — not even a note

Find ONLY these issues:

1. AMOUNT TYPOS: Extra or missing digits in a warrant amount (e.g. "$6,193.119" instead of "$6,193.19").
   Verify: does the extracted amount match the dollar figure printed on the warrant face? Flag only if clearly wrong.
   The check is on DIGITS, not formatting. "$1112.61" and "$1,112.61" both contain the digits 1-1-1-2-6-1 → identical, no typo.

2. MISSING PAYEES: A warrant has a blank or empty payee name.

3. PAYEE INCONSISTENCIES: The payee name in the itemized warrant section differs meaningfully from the consolidated narrative.
   Do NOT flag minor formatting differences (e.g. "Bluestem Capital Partners" vs "Bluestem Capital Partners, Inc." — these are the same entity).
   Only flag if the names clearly refer to different companies.

4. PAYEE TOTAL RECONCILIATION — you MUST do this for EVERY payee that appears in both sections:
   The consolidated narrative section (earlier in the document) lists each payee with a dollar amount.
   The itemized warrant resolution (later in the document) lists individual warrant amounts.
   For EVERY payee that appears in the consolidated section:
     a) Find the consolidated dollar amount stated for that payee.
     b) Sum ALL extracted warrant amounts for that payee.
     c) If |consolidated_amount − warrant_sum| > $0.05, flag it as an error.
   This check must be performed for every single payee without exception — not just fee payees.
   Never compare a single warrant amount to a consolidated group total.

5. FEE VERIFICATION:
   Advisory fees (typically Bluestem Capital Partners) and placement/underwriting fees (typically Northland Securities, Ameritas Investment, SouthState|DuncanWilliams, Access Bank, or DA Davidson & Co.) are a stated percentage of a stated base.
   The PDF explicitly states both, e.g. "advisory fees (2% of $14,517.11)".

   Step A — Identify the fee warrant:
   Only verify warrants that explicitly state a percentage. Skip flat retainer fees, fiscal agent fees, and paying agent fees — these are not percentage-based.

   Step B — Verify the math: pct × stated_base = expected_fee. Compare to extracted amount.
   Flag ONLY if |expected_fee − extracted_amount| > $0.05.

   Step C — Verify the stated base: sum the applicable warrants in the same section.
   Exclude from the base: the advisory fee warrant itself, the underwriting/placement fee warrant, any flat retainer/fiscal agent/paying agent fees.
   The underwriting/placement fee base = advisory fee base PLUS the advisory fee warrant.
   Flag ONLY if |computed_base − stated_base| > $0.05.

ANTI-PATTERN EXAMPLE — this is what NOT to do:
  BAD flag (do not produce anything like this):
    "Wt# 1187 — Amount: shows $1,112.61 in extracted data but PDF states $1,112.61. However the consolidated narrative shows TAB Construction total as $361,112.61, but extracted warrant amounts sum to $361,112.61. The individual warrant 1187 amount appears to be missing a digit — should be $1,112.61 not $1112.61."

  Why this flag is invalid:
    - "$1,112.61" vs "$1,112.61" → identical numbers
    - "$361,112.61" vs "$361,112.61" → identical numbers
    - "$1,112.61" vs "$1112.61" → identical numbers (just a comma)
  Every comparison in the message text has identical values on both sides. The pre-flight rule above requires you to suppress this flag entirely.

HARD STOP RULE: If your flag message says the calculation "appears correct", shows matching dollar amounts on both sides of any comparison, or compares two values that differ only in comma/decimal formatting, DELETE the flag entirely. Correct calculations produce no flags. The flag-message-self-check is mandatory — read your own message before output.

Return ONLY a JSON array. Empty array [] if no issues. No markdown, no explanation.
[
  {
    "severity": "error" or "warning",
    "warrant": "<warrant number(s) or N/A>",
    "field": "amount" or "payee" or "total" or "fee",
    "message": "<plain-English description of the actual discrepancy and expected value>"
  }
]"""

    text = _anthropic_request({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "system": system,
        "messages": [{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
            {"type": "text", "text": f"Audit this extracted data against the PDF:\n\n{json.dumps(extracted, indent=2)}\n\nReturn issues as JSON array."}
        ]}]
    })
    result = json.loads(text)
    if not isinstance(result, list):
        return []

    # Defense in depth: drop flags whose message text echoes the same numbers
    # back at each other rather than reporting an actual discrepancy. The
    # validator-LLM occasionally emits flags like "shows $1,112.61 but PDF
    # states $1,112.61" or "should be $1,112.61 not $1112.61" — every
    # comparison in the message has identical values on both sides.
    #
    # Heuristic: parse all dollar amounts from the flag's message, normalize
    # to numeric (so "$1,112.61" and "$1112.61" compare as equal). If every
    # distinct numeric value appears at least twice — i.e. the message is
    # purely echoing repeated values — drop the flag. Legitimate
    # discrepancy flags introduce at least one value only once.
    def _normalize_amt(s: str) -> float:
        try:
            return round(float(s.replace('$', '').replace(',', '').replace(' ', '')), 2)
        except (ValueError, AttributeError):
            return float('nan')

    _AMT_RE = re.compile(r'\$\s*[\d,]+(?:\.\d+)?')

    filtered = []
    for flag in result:
        if not isinstance(flag, dict):
            continue
        msg = flag.get('message', '') or ''
        amounts = _AMT_RE.findall(msg)
        if len(amounts) >= 2:
            parsed = [_normalize_amt(a) for a in amounts]
            parsed = [v for v in parsed if v == v]  # drop NaN
            if parsed:
                from collections import Counter
                counts = Counter(parsed)
                # Suppress when every distinct value is echoed at least twice
                # AND no value appears only once — signature of a flag where
                # both sides of each claimed comparison are identical.
                if min(counts.values()) >= 2:
                    continue
        filtered.append(flag)

    return filtered

# ── Excel: Sarpy County ───────────────────────────────────────────────────────
def build_excel_sarpy(extracted: dict, dt_wt: date, dt_reg: date) -> bytes:
    sid_num  = extracted['sid_number']
    sid_name = SID_NAMES.get(sid_num, '')
    _sid_info  = get_sid_info(sid_num)
    sid_coupon = _sid_info['coupon'] if _sid_info else 0.07
    gw = [w for w in extracted.get('general_warrants', [])      if not w.get('void')]
    cw = [w for w in extracted.get('construction_warrants', []) if not w.get('void')]

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    DATE_FMT = 'mm/dd/yy;@'
    AMT_FMT  = '#,##0.00'
    PCT_FMT  = '0%'
    bold     = Font(bold=True)
    med_bot  = Border(bottom=Side(style='medium'))

    col_widths = {'A':7.14,'B':4.99,'C':5.71,'D':8.14,'E':8.14,
                  'F':11.7,'G':6.41,'H':8.14,'I':11.99,'J':8.14,
                  'K':50.84,'L':25.41,'M':16.84}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w
    for r in range(1, 80):
        ws.row_dimensions[r].height = 12.75

    for rng in ['A1:I1','A2:I2','A3:I3','A4:I4']:
        ws.merge_cells(rng)
    ws['A1'].value = 'INSTRUCTIONS'
    ws['A2'].value = '1. Register the warrants listed below'
    ws['A3'].value = '2. Type the name of the person receiving the data and the date in the provided cells'
    ws['A4'].value = '3. Save and email the form back to Bluestem Capital Partners '
    ws['K1'].value = 'Date:'
    ws['L1'].value = '=TODAY()'
    ws['L1'].number_format = DATE_FMT
    ws['K5'].value = 'Prepared By:'
    ws['L5'].value = 'jeff@bluestemcap.com'
    ws['L5'].font = Font(color='0000D4')

    def write_section(header_row, data_start, warrants, maturity_years):
        for col in 'ABCDEFGHIJKL':
            ws[f'{col}{header_row}'].border = med_bot
            ws[f'{col}{header_row}'].font = bold
        ws[f'A{header_row}'].value = f'SID {sid_num}'
        ws[f'B{header_row}'].value = 'General WARRANTS ' if maturity_years == 3 else 'Construction WARRANTS '
        ws[f'B{header_row}'].alignment = Alignment(horizontal='left')
        ws[f'F{header_row}'].value = 'Bluestem Capital Partners'
        for col in ['F','G','H']:
            ws[f'{col}{header_row}'].number_format = AMT_FMT
        ws[f'I{header_row}'].value = ' '
        for col in ['I','J','D','E','L']:
            ws[f'{col}{header_row}'].number_format = DATE_FMT
        ws[f'K{header_row}'].value = sid_name

        col_row = header_row + 1
        col_hdrs = {'B':('Off #','right'),'C':('Wt #',None),'D':('Dt Wt',None),
                    'E':('Dt Reg',None),'F':('Amt',None),'G':('Int',None),
                    'H':('Ttl Pd',None),'I':('Dt Not',None),'J':('Dt Pd',None),
                    'K':('Drawn To',None),'L':('Pro Dt','left')}
        for col,(val,align) in col_hdrs.items():
            ws[f'{col}{col_row}'].value = val
            ws[f'{col}{col_row}'].font = bold
            if align:
                ws[f'{col}{col_row}'].alignment = Alignment(horizontal=align)
        for col in ['D','E','I','J','L']:
            ws[f'{col}{col_row}'].number_format = DATE_FMT
        for col in ['F','G','H']:
            ws[f'{col}{col_row}'].number_format = AMT_FMT

        for i, w in enumerate(warrants):
            r = data_start + i
            ws[f'A{r}'].value = f'=IF(C{r}="","",{sid_coupon})'
            ws[f'A{r}'].number_format = PCT_FMT
            ws[f'B{r}'].font = bold
            ws[f'B{r}'].alignment = Alignment(horizontal='right')
            ws[f'C{r}'].value = w['warrant_number']
            ws[f'D{r}'].value = dt_wt
            ws[f'D{r}'].number_format = DATE_FMT
            ws[f'E{r}'].value = dt_reg
            ws[f'E{r}'].number_format = DATE_FMT
            ws[f'F{r}'].value = w['amount']
            ws[f'F{r}'].number_format = AMT_FMT
            for col in ['G','H']:
                ws[f'{col}{r}'].font = bold
                ws[f'{col}{r}'].number_format = AMT_FMT
            for col in ['I','J']:
                ws[f'{col}{r}'].font = bold
                ws[f'{col}{r}'].number_format = DATE_FMT
            ws[f'K{r}'].value = w['payee']
            ws[f'L{r}'].value = date(dt_wt.year + maturity_years, dt_wt.month, dt_wt.day)
            ws[f'L{r}'].number_format = DATE_FMT
            ws[f'L{r}'].alignment = Alignment(horizontal='center')

        sum_row = data_start + len(warrants)
        ws[f'F{sum_row}'].value = f'=SUM(F{data_start}:F{sum_row-1})'
        ws[f'F{sum_row}'].number_format = AMT_FMT
        return sum_row

    # CF section: header row 8, data starts row 11
    cf_sum = write_section(8, 11, cw, 5)
    # GF section: header at row 24 minimum, or 3 rows after CF sum if CF is large
    gf_header = max(24, cf_sum + 3)
    gf_sum    = write_section(gf_header, gf_header + 3, gw, 3)

    # Signature: 2 rows below GF sum
    received_row = gf_sum + 2
    date_row     = received_row + 4
    ws.merge_cells(f'H{received_row}:K{received_row+1}')
    ws[f'H{received_row}'].value = 'Received By:'
    ws.merge_cells(f'H{date_row}:K{date_row+1}')
    ws[f'H{date_row}'].value = 'Date:'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Excel: Douglas County ─────────────────────────────────────────────────────
def build_excel_douglas(extracted: dict, dt_wt: date, dt_reg: date) -> bytes:
    sid_num = extracted['sid_number']
    _sid_info_d  = get_sid_info(sid_num)
    sid_rate_pct = _sid_info_d['rate_pct'] if _sid_info_d else 7
    gw = [w for w in extracted.get('general_warrants', [])      if not w.get('void')]
    cw = [w for w in extracted.get('construction_warrants', []) if not w.get('void')]

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    DATE_FMT = 'mm/dd/yy;@'
    AMT_FMT  = '"$"#,##0.00'
    bold     = Font(bold=True)

    col_widths = {'A':12,'B':2,'C':10,'D':2,'E':12,'F':2,
                  'G':40,'H':2,'I':14,'J':2,'K':14,'L':2,'M':8}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w
    for r in range(1, 80):
        ws.row_dimensions[r].height = 12.75

    for rng in ['A1:J1','A2:J2','A3:J3','A4:J4']:
        ws.merge_cells(rng)
    ws['A1'].value = 'INSTRUCTIONS'
    ws['A2'].value = '1. Register the warrants listed below'
    ws['A3'].value = '2. Sign and date the form'
    ws['A4'].value = '3. Email to Bluestem Capital Partners to confirm warrants have been registered'
    ws['K1'].value = 'Date:'
    ws['M1'].value = '=TODAY()'
    ws['M1'].number_format = DATE_FMT
    ws['H6'].value = 'Prepared By:'
    ws['I6'].value = 'jeff@bluestemcap.com'
    ws['I6'].font = Font(color='0000D4')

    def write_col_headers(row):
        for col, label in [('C','Warrant Number'),('E','Date of Warrant'),('G','Payee'),
                            ('I','Date Registered'),('K','Warrant Amount'),('M','Interest')]:
            ws[f'{col}{row}'].value = label
            ws[f'{col}{row}'].font = bold

    def write_section(header_row, warrants, label):
        ws[f'A{header_row}'].value = f'SID Number: {sid_num} {label}'
        ws[f'A{header_row}'].font = bold
        write_col_headers(header_row + 1)
        data_start = header_row + 3  # header + col headers + blank row

        for i, w in enumerate(warrants):
            r = data_start + i
            ws[f'A{r}'].value = 'Not Paid'
            ws[f'C{r}'].value = w['warrant_number']
            ws[f'E{r}'].value = dt_wt
            ws[f'E{r}'].number_format = DATE_FMT
            ws[f'G{r}'].value = w['payee']
            ws[f'I{r}'].value = dt_reg
            ws[f'I{r}'].number_format = DATE_FMT
            ws[f'K{r}'].value = w['amount']
            ws[f'K{r}'].number_format = AMT_FMT
            ws[f'M{r}'].value = sid_rate_pct

        sum_row = data_start + len(warrants)
        ws[f'K{sum_row}'].value = f'=SUM(K{data_start}:K{sum_row-1})'
        ws[f'K{sum_row}'].number_format = AMT_FMT
        return sum_row

    # Construction starts at row 9 if present; General follows dynamically
    if cw:
        cf_sum = write_section(9, cw, 'Construction')
        gf_start = cf_sum + 3
    else:
        gf_start = 9  # no CF section, GF starts at top

    if gw:
        gf_sum = write_section(gf_start, gw, 'General')
    else:
        gf_sum = gf_start - 1  # no GF section, point signature just below CF

    # Signature: 2 rows below GF sum
    received_row = gf_sum + 2
    date_row     = received_row + 3
    ws[f'F{received_row}'].value = 'Received By:'
    ws[f'F{date_row}'].value = 'Date:'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Route to correct builder ──────────────────────────────────────────────────
def build_excel(extracted: dict, dt_wt: date, dt_reg: date) -> bytes:
    county = extracted.get('county', 'Sarpy').strip().lower()
    if county == 'douglas':
        return build_excel_douglas(extracted, dt_wt, dt_reg)
    return build_excel_sarpy(extracted, dt_wt, dt_reg)

# ── Warrant-face PDF → Registration Log ──────────────────────────────────────

def _split_pdf_bytes(pdf_bytes: bytes, pages_per_chunk: int = 60) -> list:
    """Split PDF bytes into chunks of N pages. Returns list of bytes objects."""
    import pikepdf
    src = pikepdf.open(io.BytesIO(pdf_bytes))
    total = len(src.pages)
    chunks = []
    for start in range(0, total, pages_per_chunk):
        end = min(start + pages_per_chunk, total)
        dst = pikepdf.Pdf.new()
        for i in range(start, end):
            dst.pages.append(src.pages[i])
        buf = io.BytesIO()
        dst.save(buf)
        chunks.append(buf.getvalue())
    return chunks


def _extract_warrants_from_chunk(pdf_bytes: bytes) -> list:
    """Call Anthropic on a single PDF chunk and return list of warrant dicts."""
    system = """You are a precise data extraction assistant reading scanned physical SID warrant certificates.
Each warrant has a front face and a back registration page. Extract every warrant present.

Return ONLY valid JSON — a list of warrant objects with this exact structure:
[
  {
    "sid_number": <integer>,
    "county": "<string, e.g. 'Sarpy' or 'Douglas'>",
    "fund_type": "<'Construction' or 'General'>",
    "warrant_number": <integer>,
    "payee": "<string — exact name as printed>",
    "amount": <float>,
    "interest_rate": <float, e.g. 0.07 for 7% — read from warrant face; null if blank>,
    "warrant_date": "<MM/DD/YYYY>",
    "maturity_date": "<MM/DD/YYYY — read from 'shall become due on' line; null if not determinable>",
    "registration_date": "<MM/DD/YYYY — read from the 'Registration by County Treasurer' back page>",
    "payment_description": "<string — see Payment Description rules below; empty string if truly nothing found>"
  }
]

Rules:
- Fund type: read the large header text ('CONSTRUCTION FUND' or 'GENERAL FUND') at the top of each warrant face.
- Warrant date: the DATE field on the warrant face. Dates are often split across two printed fields, e.g. 'April 2' and '26' = 04/02/2026. Always combine month/day and year into MM/DD/YYYY.
- Maturity date: read from the 'shall become due on' line. The month and day are usually pre-printed (same as warrant date) and only the year is filled in — so if the warrant date is April 2, 2026 and the maturity year reads '29', the maturity date is 04/02/2029. If only a year is visible with no month/day, use the same month and day as the warrant date. Return null only if truly nothing is readable.
- Registration date: found on the back 'Registration by County Treasurer' page (e.g. 'April 8, 2026' = 04/08/2026).
- Interest rate: read the filled-in rate on the face (e.g. '7' = 0.07). If the line is blank, return null.
- Payment description — search ALL of these locations and combine what you find:
    1. The 'IN PAYMENT OF' line at the bottom-left of the warrant face (often handwritten in).
    2. The block just ABOVE the 'IN PAYMENT OF' label — handwritten notes here are COMMON. This is often the ONLY place the payment description is written, with the 'IN PAYMENT OF' line itself left blank.
    3. Handwritten notes that BLEED INTO, OVERLAP, sit JUST BELOW, or appear UNDERNEATH the 'SID SERVICES / OMAHA, NEBRASKA' paying-agent address block to the right. The handwriting can appear in four sub-positions inside this zone:
       (a) ACROSS the address text (overlapping pre-printed letters) — e.g. 'GF #298-305' written through 'SID SERVICES'.
       (b) BETWEEN the 'SID SERVICES' line and the 'OMAHA, NEBRASKA' line — a thin gap that often holds 'MA fee on GF', 'Under. fee on GF', etc.
       (c) UNDERNEATH the address — the small gap between the '402-504-3967' phone number line (or the OMAHA NE 68022 line) and the 'IN PAYMENT OF' label. THIS IS A VERY COMMON SPOT for handwritten invoice numbers, often appearing as a list of '#NNNNNN' values (e.g. '#505588 506696', '#181173 179487 180879'). The numbers may be space-separated, comma-separated, or written one above another. INSPECT THIS SUB-ZONE CAREFULLY — it can look like part of the printed address but is actually handwritten.
       (d) BELOW the 'OMAHA, NEBRASKA' line — the empty white space at the very bottom-right of the warrant face. THIS IS COMMONLY USED FOR LONG INVOICE LISTS, especially comma-separated forms like 'Inv 181173, 179487, 180879, 180370, 179951'. When invoice numbers don't fit elsewhere they end up here. ALWAYS check this sub-zone — it's easy to overlook because it sits in apparently blank space.
       The handwritten notes in any of these four sub-positions are real payment descriptions. They got placed there because the 'IN PAYMENT OF' area was too small or because the writer worked through the available space. LOOK VERY CAREFULLY in all four sub-positions.
    4. Any handwritten payment reference elsewhere in the bottom third of the warrant face (between the signature lines and the bottom edge).
  Common payment-description patterns to scan for in any of the four zones:
    - 'No. NNNN' or 'No. NNNN, NNNN, NNNN' — invoice/PO references (e.g. 'No. 2307', 'No. 381371', 'No. 2305, 2306, 2327').
    - '#NNNNNN' or 'Inv NNNNNN, NNNNNN, ...' — invoice numbers, sometimes a comma-separated list (e.g. '#299274', 'Inv 181173, 179487, 180879').
    - 'GF #NNN-NNN' or 'CF #NNN-NNN' — warrant range references on fee warrants (e.g. 'GF #298-305', 'CF #1030-1047', 'GF #1050-1056').
    - 'Warrant Nos. NNNN-NNNN' — verbose warrant range form on Northland fee warrants (rule will blank these later, but capture verbatim).
    - 'MA fee on GF' / 'Under. fee on GF' / 'Adv. fee on GF' — fee-allocation tags written across the address block.
    - 'PE #N' / 'Pay Estimate #N' — pay-estimate references on contractor warrants (e.g. 'PE #12 Sanitary Sewer & Storm Sewer Section I').
    - 'Inv N' followed by a project description — invoice + scope note (e.g. 'Inv 6', '5/1/26 Construction Fund Interest', '2025-2026 ASIP Fees (260968)').
    - 'Legal services <project>' — legal billing references (e.g. 'Legal services Sanitary Sewer & Storm Sewer Section I').
    - 'ARF, SWWCF, WCF Nth Install.' — installment payment references.
  Combine all handwritten content from these spots into a single payment description string. Do NOT include the pre-printed 'SID SERVICES' / 'OMAHA, NEBRASKA' address text itself — only handwritten additions.
  Before returning an empty string, you MUST confirm that you have inspected ALL FOUR zones described above for the warrant in question. Empty is only acceptable when every zone is genuinely blank. If you see any handwriting at all in any zone, transcribe it.
  Within a single PDF batch, warrants to the same payee in a consecutive sequence often share the same payment description. If a warrant is in such a sequence and the description is hard to read on one page, the consistent description from its neighbors is strong corroborating evidence — use it if the visible ink in the correct zone is consistent with that value. Do NOT fabricate a description that has no visible ink at all; only use consistency as a tiebreaker when the zone is partly legible.
  Prefix any value starting with '#' (and only '#') with 'Inv '. So '#299274' becomes 'Inv 299274', and '#505588 506696' becomes 'Inv 505588 506696' (transcribe all the numbers, prefix once). Do NOT add 'Inv ' prefix to values starting with 'No.', 'Warrant', 'GF', 'CF', 'MA', 'Under.', 'Adv.', 'PE', 'Legal', 'Inv', or any other word.
  Return empty string only if you truly find no handwritten payment description anywhere in the bottom region of the warrant face.
- Do NOT include registration back-pages as separate warrants — they belong to the preceding warrant face.
- Every warrant face in the PDF must produce a JSON entry — never skip a warrant because a field is blank or unclear. Use null for numeric fields and empty string for text fields when data is missing.
- If a warrant face appears cut off at the end of the chunk, include it with whatever data is visible.
- Return ONLY the JSON array, no markdown, no explanation."""

    pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
    text = _anthropic_request({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 8192,
        "system": system,
        "messages": [{"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
            {"type": "text", "text": "Extract all warrant data from these warrant face scans and return as a JSON array. Include every warrant present. Do not truncate."}
        ]}]
    })
    try:
        result = json.loads(text)
        return result if isinstance(result, list) else []
    except json.JSONDecodeError as e:
        # Log the bad response to help diagnose issues
        preview = text[:200].replace('\n', ' ')
        raise ValueError(f"Warrant extraction returned invalid JSON: {e}. Response preview: {preview}")


def call_anthropic_warrant_faces(pdf_b64: str) -> list:
    """
    Extract warrant data from a PDF of physical warrant faces (front + back).
    Automatically splits large PDFs into 30-page chunks to stay within
    the Anthropic API's payload size limit, then merges the results.
    Returns a list of warrant dicts ready for build_registration_log_excel().
    """
    pdf_bytes = base64.b64decode(pdf_b64)

    # Chunk if the base64-encoded size would exceed ~8MB (Anthropic API limit)
    B64_LIMIT = 24 * 1024 * 1024  # 24MB base64 (~18MB binary)
    if len(pdf_b64) <= B64_LIMIT:
        # Small enough — process in one shot
        return _extract_warrants_from_chunk(pdf_bytes)

    # Split into 30-page chunks and process each
    chunks = _split_pdf_bytes(pdf_bytes, pages_per_chunk=60)
    all_warrants = []
    seen = set()
    for i, chunk_bytes in enumerate(chunks):
        if i > 0:
            time.sleep(30)  # pause between chunks to avoid rate limiting
        warrants = _extract_warrants_from_chunk(chunk_bytes)
        for w in warrants:
            # Deduplicate by (sid_number, warrant_number) in case a warrant
            # appears at a chunk boundary and gets extracted twice
            key = (w.get('sid_number'), w.get('warrant_number'))
            if key not in seen:
                seen.add(key)
                all_warrants.append(w)
    return all_warrants


def build_registration_log_excel(warrants: list, info_8038_df) -> bytes:
    """
    Build the warrant registration log spreadsheet from extracted warrant-face data.
    Row 1: blank with date label at the end. Row 2/3/4: blank. Row 5: headers. Row 6+: data.

    Column layout (11 columns):
      A County | B SID | C Payee | D Type | E Warrants | F Amount |
      G Registration Date | H Maturity Date | I Interest | J First Interest Date |
      K Payment Description
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Warrant Registration"

    headers = [
        "County", "SID", "Payee", "Type", "Warrants",
        "Amount", "Registration Date", "Maturity Date",
        "Interest", "First Interest Date", "Payment Description",
    ]
    NUM_COLS = len(headers)  # 11

    # ── Row 1: date label in the last column (K1 with 11 cols) ────────────
    date_cell = ws.cell(row=1, column=NUM_COLS,
                        value='Date: ' + date.today().strftime('%m/%d/%Y'))
    date_cell.font = Font(name="Calibri", bold=True, size=11)

    # ── Row 5: column headers ──────────────────────────────────────────────
    HEADER_ROW = 5
    header_font   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill   = PatternFill("solid", start_color="003366")
    header_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_border = Border(
        bottom=Side(style="medium"), top=Side(style="medium"),
        left=Side(style="medium"), right=Side(style="medium")
    )
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = header_border

    ws.row_dimensions[HEADER_ROW].height = 30

    # ── Rows 6+: data ─────────────────────────────────────────────────────
    # Date columns are G (7), H (8), J (10)
    DATE_COLS = {7, 8, 10}
    # Column G of formatting:
    #   F (6) = Amount → currency, centered
    #   I (9) = Interest rate → percentage, centered
    #   B (2), D (4), E (5) = SID, Type, Warrants → centered
    #   C (3), K (11) = Payee, Payment Description → left
    data_font  = Font(name="Arial", size=10)
    align_ctr  = Alignment(horizontal="center")
    align_left = Alignment(horizontal="left")

    def parse_date(s):
        if not s:
            return None
        for fmt in ('%m/%d/%Y', '%Y-%m-%d'):
            try:
                return datetime.strptime(s, fmt).date()
            except (ValueError, TypeError):
                pass
        return None

    for r_idx, w in enumerate(warrants, HEADER_ROW + 1):
        sid_num   = w.get('sid_number')
        fund_type = (w.get('fund_type') or '').strip()
        county    = (w.get('county') or '').strip()
        is_cf     = fund_type.lower() == 'construction'

        # Interest rate: warrant face first, then 8038 lookup, then 7% default.
        # Stored as a fraction (0.07) so the percent format displays "7%".
        rate = w.get('interest_rate')
        if rate is None:
            _info = get_sid_info(sid_num)
            rate = _info['coupon'] if _info else 0.07

        # First Interest Date: Construction Fund only — General Fund stays blank.
        first_int_date = None
        if is_cf:
            rows = info_8038_df[info_8038_df['SID'] == sid_num]
            if not rows.empty:
                fi = rows.iloc[0]['First Annual Interest Payment']
                if hasattr(fi, 'month'):
                    rd = parse_date(w.get('registration_date'))
                    candidate = (date(rd.year if rd else fi.year, fi.month, fi.day)
                                 if rd else fi.date())
                    if rd and candidate <= rd:
                        candidate = date(candidate.year + 1, candidate.month, candidate.day)
                    first_int_date = candidate

        # Standing rule: Northland Securities payees always have Payment Description blanked.
        # This avoids cluttering the log with the "Warrant Nos. X-Y" reference that appears
        # on the fee-collection warrant since Northland's payment is tied to prior warrants
        # by the warrant number range alone.
        pay_desc = w.get('payment_description', '') or ''
        payee_for_rule = (w.get('payee') or '').lower()
        if 'northland' in payee_for_rule:
            pay_desc = ''

        # Maturity-date enforcement. The warrant ledger conventions are:
        #   Construction (CF) → 5 years from warrant date
        #   General (GF)      → 3 years from warrant date
        # 99.7% of CF and 99.9% of GF warrants in the historical dataset follow this
        # convention. The PDF model occasionally misreads the maturity year off blurry
        # warrant faces (e.g. confusing '31' with '29' on a smudgy CF warrant), which
        # then defaults to GF's 3-year span. Override with the convention so the
        # output log is reliable. Warrant Date is no longer in the column layout but
        # is still needed internally as the base for the maturity calculation.
        warrant_date_obj = parse_date(w.get('warrant_date'))
        maturity_date_obj = parse_date(w.get('maturity_date'))
        if warrant_date_obj is not None:
            expected_years = 5 if is_cf else 3
            try:
                expected_maturity = date(
                    warrant_date_obj.year + expected_years,
                    warrant_date_obj.month,
                    warrant_date_obj.day,
                )
            except ValueError:
                # Feb 29 leap-day rollover edge case
                expected_maturity = date(
                    warrant_date_obj.year + expected_years,
                    warrant_date_obj.month,
                    warrant_date_obj.day - 1 if warrant_date_obj.day == 29 else warrant_date_obj.day,
                )
            maturity_date_obj = expected_maturity

        row_vals = [
            county,                                  # A County
            sid_num,                                 # B SID
            w.get('payee', ''),                      # C Payee
            fund_type,                               # D Type ('General' or 'Construction')
            w.get('warrant_number'),                 # E Warrants
            w.get('amount'),                         # F Amount
            parse_date(w.get('registration_date')),  # G Registration Date
            maturity_date_obj,                       # H Maturity Date
            rate,                                    # I Interest
            first_int_date,                          # J First Interest Date
            pay_desc,                                # K Payment Description
        ]

        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font   = data_font
            cell.border = Border()
            if c_idx == 6:                  # F Amount
                cell.number_format = '$#,##0.00'
                cell.alignment = align_ctr
            elif c_idx == 9:                # I Interest rate
                cell.number_format = '0%'
                cell.alignment = align_ctr
            elif c_idx in DATE_COLS:        # G, H, J
                if val is not None:
                    cell.number_format = 'MM/DD/YYYY'
                cell.alignment = align_ctr
            elif c_idx in {2, 4, 5}:        # SID, Type, Warrants
                cell.alignment = align_ctr
            else:                            # Payee, Payment Description (left)
                cell.alignment = align_left

    # Column widths in the new order: County, SID, Payee, Type, Warrants,
    # Amount, Reg Date, Maturity Date, Interest, First Int Date, Pay Desc
    col_widths = [10, 8, 40, 14, 10, 14, 16, 14, 10, 18, 35]
    for i, w_val in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w_val

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        if request.form.get('password') == APP_PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('index'))
        error = 'Incorrect password'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/api/process', methods=['POST'])
@login_required
def process():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No PDF uploaded'}), 400

    pdf_file   = request.files['pdf']
    dt_wt_str  = request.form.get('dt_wt', '')
    dt_reg_str = request.form.get('dt_reg', '')

    if not dt_wt_str or not dt_reg_str:
        return jsonify({'error': 'Both dates are required'}), 400

    try:
        datetime.strptime(dt_wt_str,  '%Y-%m-%d')
        datetime.strptime(dt_reg_str, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'Invalid date format'}), 400

    pdf_bytes = pdf_file.read()
    pdf_b64   = base64.b64encode(pdf_bytes).decode('utf-8')

    try:
        extracted = call_anthropic(pdf_b64)
    except Exception as e:
        return jsonify({'error': f'PDF extraction failed: {str(e)}'}), 500

    extracted['sid_name'] = SID_NAMES.get(extracted.get('sid_number', 0), 'Unknown')

    # Sanity check: if SID name is unknown, look for truncation (e.g. 36 instead of 363)
    sid_num = extracted.get('sid_number', 0)
    if extracted['sid_name'] == 'Unknown' and sid_num:
        candidates = [n for n in SID_NAMES if str(n).startswith(str(sid_num)) and n != sid_num]
        if candidates:
            clist = ', '.join(str(c) for c in candidates[:3])
            extracted['sid_number_warning'] = f'SID #{sid_num} not found in lookup — did you mean {clist}?'

    # Surface SID number warning as a flag if present
    sid_warning_flags = []
    if extracted.get('sid_number_warning'):
        sid_warning_flags.append({
            'severity': 'error',
            'warrant': 'N/A',
            'field': 'sid_number',
            'message': extracted['sid_number_warning']
        })

    # Collect void warrant flags from extraction
    void_flags = []
    for w in extracted.get('construction_warrants', []) + extracted.get('general_warrants', []):
        if w.get('void'):
            void_flags.append({
                'severity': 'warning',
                'warrant': str(w['warrant_number']),
                'field': 'void',
                'message': f'Warrant No. {w["warrant_number"]} is marked VOID in the PDF and has been excluded from the registration sheet.'
            })

    try:
        flags = call_anthropic_validate(pdf_b64, extracted)
    except Exception:
        flags = []

    # Void flags first, then other validation flags (skip duplicate void flags from validator)
    all_flags = sid_warning_flags + void_flags + [f for f in flags if f.get('field') != 'void']

    return jsonify({
        'success': True,
        'data': extracted,
        'flags': all_flags,
        'dt_wt': dt_wt_str,
        'dt_reg': dt_reg_str
    })

@app.route('/api/download', methods=['POST'])
@login_required
def download():
    body       = request.get_json()
    extracted  = body.get('data', {})
    dt_wt_str  = body.get('dt_wt', '')
    dt_reg_str = body.get('dt_reg', '')

    try:
        dt_wt  = datetime.strptime(dt_wt_str,  '%Y-%m-%d').date()
        dt_reg = datetime.strptime(dt_reg_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'Invalid dates'}), 400

    try:
        xlsx_bytes = build_excel(extracted, dt_wt, dt_reg)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    sid_num      = extracted.get('sid_number', 'XXX')
    reg_date_str = dt_reg.strftime('%m-%d-%Y')
    xlsx_name    = f'SID_{sid_num}_Warrants_to_Register_{reg_date_str}.xlsx'

    # Generate 8038 PDFs
    import traceback
    pdf_error = None
    pdf_debug = {}
    try:
        # Pre-flight debug info
        sid_num_int = extracted.get('sid_number')
        info_check  = get_sid_info(sid_num_int)
        cw_raw = extracted.get('construction_warrants', [])
        gw_raw = extracted.get('general_warrants', [])
        cw_filt = [w for w in cw_raw if not w.get('void')]
        gw_filt = [w for w in gw_raw if not w.get('void')]
        pdf_debug = {
            'sid_number':        sid_num_int,
            'sid_in_info':       info_check is not None,
            'county':            info_check['county'] if info_check else None,
            'construction_raw':  len(cw_raw),
            'construction_filt': len(cw_filt),
            'construction_total': round(sum(w['amount'] for w in cw_filt), 2),
            'general_raw':       len(gw_raw),
            'general_filt':      len(gw_filt),
            'general_total':     round(sum(w['amount'] for w in gw_filt), 2),
            'dt_wt':             str(dt_wt),
            'dt_reg':            str(dt_reg),
        }
        pdf_forms = build_8038_forms(extracted, dt_wt, dt_reg)
        pdf_debug['forms_generated'] = [item[0] for item in pdf_forms]
    except Exception as e:
        pdf_forms = []
        pdf_error = traceback.format_exc()
        pdf_debug['exception'] = str(e)

    # If no 8038s generated, return zip with Excel + debug log
    if not pdf_forms:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(xlsx_name, xlsx_bytes)
            log = json.dumps(pdf_debug, indent=2)
            if pdf_error:
                log += '\n\nTRACEBACK:\n' + pdf_error
            zf.writestr('8038_DEBUG.txt', log)
        zip_buf.seek(0)
        zip_name = f'SID_{sid_num}_{reg_date_str}.zip'
        return send_file(
            zip_buf,
            mimetype='application/zip',
            as_attachment=True,
            download_name=zip_name
        )

    # Generate yield/WAM workbook if any 8038-G was produced (CF >= $100k)
    yield_wb = None
    wam_wb   = None
    cf_warrants = [w for w in extracted.get('construction_warrants', []) if not w.get('void')]
    cf_total_check = round(sum(w['amount'] for w in cf_warrants), 2)
    if cf_total_check >= 100_000:
        try:
            yield_wb = build_yield_wam_workbook(extracted, dt_wt, dt_reg)
        except Exception:
            yield_wb = None
        try:
            wam_wb = build_wam_workbook(extracted, dt_wt, dt_reg)
        except Exception:
            wam_wb = None

    # Bundle everything into a zip
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(xlsx_name, xlsx_bytes)
        for item in pdf_forms:
            zf.writestr(item[0], item[1])
        if yield_wb:
            zf.writestr(yield_wb[0], yield_wb[1])
        if wam_wb:
            zf.writestr(wam_wb[0], wam_wb[1])
    zip_buf.seek(0)

    zip_name = f'SID_{sid_num}_{reg_date_str}.zip'
    return send_file(
        zip_buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=zip_name
    )

def _run_wrl_job(job_id: str, pdf_bytes: bytes):
    """Background thread: extract warrants and build Excel, store result in _wrl_jobs."""
    try:
        pdf_b64  = base64.b64encode(pdf_bytes).decode('utf-8')
        warrants = call_anthropic_warrant_faces(pdf_b64)
        if not warrants:
            raise ValueError('No warrants found in PDF')
        xlsx_bytes = build_registration_log_excel(warrants, INFO_8038)
        first_sid  = warrants[0].get('sid_number', 'XXX')
        fname      = f'Warrant_Registration_Log_{first_sid}_{date.today().strftime("%m-%d-%Y")}.xlsx'
        with _wrl_jobs_lock:
            _wrl_jobs[job_id] = {'status': 'done', 'result': xlsx_bytes, 'filename': fname, 'error': None}
    except Exception as e:
        with _wrl_jobs_lock:
            _wrl_jobs[job_id] = {'status': 'error', 'result': None, 'filename': None, 'error': str(e)}


@app.route('/api/register-warrants', methods=['POST'])
@login_required
def register_warrants():
    """Start an async warrant extraction job. Returns immediately with a job_id."""
    if 'pdf' not in request.files:
        return jsonify({'error': 'No PDF uploaded'}), 400
    pdf_bytes = request.files['pdf'].read()
    job_id = str(uuid.uuid4())
    with _wrl_jobs_lock:
        _wrl_jobs[job_id] = {'status': 'running', 'result': None, 'filename': None, 'error': None}
    t = threading.Thread(target=_run_wrl_job, args=(job_id, pdf_bytes), daemon=True)
    t.start()
    return jsonify({'job_id': job_id})


@app.route('/api/register-warrants/status/<job_id>')
@login_required
def register_warrants_status(job_id):
    """Poll for job status. Returns {status, filename, error}."""
    with _wrl_jobs_lock:
        job = _wrl_jobs.get(job_id)
    if job is None:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify({
        'status':   job['status'],
        'filename': job['filename'],
        'error':    job['error'],
    })


@app.route('/api/register-warrants/download/<job_id>')
@login_required
def register_warrants_download(job_id):
    """Download the completed Excel file for a job."""
    with _wrl_jobs_lock:
        job = _wrl_jobs.get(job_id)
    if job is None or job['status'] != 'done':
        return jsonify({'error': 'Job not ready'}), 404
    xlsx_bytes = job['result']
    fname      = job['filename']
    # Clean up job from memory after download
    with _wrl_jobs_lock:
        _wrl_jobs.pop(job_id, None)
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=fname
    )


# ── GF Warrant Calls ──────────────────────────────────────────────────────────

def _gf_calls_select_warrants(df: pd.DataFrame, sid: int, call_date: date, cash_avail: float) -> list:
    """
    Select GF warrants eligible for a call:
      - Type == GF, CallDate is blank, Registration is set
      - Warrants with the same Number are combined (principal summed) into one row
      - Sort by Registration ASC, Warrant # ASC
      - Include in order until next warrant would exceed cash (first-fail stop)
    """
    warrants = df[
        (df['SID'] == sid) &
        (df['Type'] == 'GF') &
        (df['CallDate'].isna()) &
        (df['Registration'].notna())
    ].copy()

    # Collapse duplicate warrant numbers — sum principal, keep first row's other fields
    warrants = (
        warrants
        .sort_values(['Registration', 'Number'])
        .groupby(['SID', 'Number'], sort=False)
        .agg(
            Type=('Type', 'first'),
            Payee=('Payee', 'first'),
            Principal=('Principal', 'sum'),
            Rate=('Rate', 'first'),
            Registration=('Registration', 'first'),
            Maturity=('Maturity', 'first'),
        )
        .reset_index()
        .sort_values(['Registration', 'Number'])
    )

    running = 0.0
    included = []
    for _, w in warrants.iterrows():
        reg_d = w['Registration'].date() if hasattr(w['Registration'], 'date') else w['Registration']
        d = days_30_360(reg_d, call_date)
        interest = round(float(w['Principal']) * (float(w['Rate']) / 100) * d / 360, 2)
        total = float(w['Principal']) + interest
        if running + total <= cash_avail:
            running += total
            mat = w['Maturity']
            included.append({
                'type':         'GF',
                'warrant':      int(w['Number']),
                'payee':        w['Payee'],
                'principal':    float(w['Principal']),
                'rate':         float(w['Rate']) / 100,
                'interest':     interest,
                'registration': reg_d.strftime('%m/%d/%Y'),
                'maturity':     mat.date().strftime('%m/%d/%Y') if pd.notna(mat) else '',
            })
        else:
            break
    return included


def _gf_calls_build_workbook(sid: int, sid_name: str, county: str,
                              call_day_str: str, warrants: list):
    """Build a single GF Call Report workbook."""
    from openpyxl import Workbook as OWorkbook
    wb = OWorkbook()
    ws = wb.active
    ws.title = 'Call Report'

    col_widths = {'A': 13, 'B': 13, 'C': 13, 'D': 9.14,
                  'E': 9.14, 'F': 9.14, 'G': 13, 'H': 13, 'I': 13}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    arial12 = Font(name='Arial', size=12)
    arial10 = Font(name='Arial', size=10)
    center  = Alignment(horizontal='center')

    headers = [
        f'SANITARY AND IMPROVEMENT DISTRICT {sid} OF {county.upper()} COUNTY',
        f'"{sid_name.upper()}"',
        'WARRANT CALL REPORT',
        'WARRANTS TO BE CALLED AS OF',
        call_day_str.upper(),
        '',
    ]
    for i, text in enumerate(headers, start=1):
        ws.merge_cells(f'A{i}:I{i}')
        cell = ws.cell(row=i, column=1, value=text)
        cell.font = arial12
        cell.alignment = center

    col_headers  = ['Type', 'Warrant #', 'Payee', 'Principal', 'Rate',
                    'Interest', 'Registration', 'Maturity', 'Last Paid']
    col_formats  = [None, None, None, r'$#,##0.00', '0%', r'$#,##0.00', None, None, None]
    for col_idx, (hdr, fmt) in enumerate(zip(col_headers, col_formats), start=1):
        cell = ws.cell(row=7, column=col_idx, value=hdr)
        cell.font = arial10
        if fmt:
            cell.number_format = fmt

    for row_idx, w in enumerate(warrants, start=8):
        values  = [w['type'], w['warrant'], w['payee'], w['principal'],
                   w['rate'], w['interest'], w['registration'], w['maturity'], None]
        formats = [None, None, None, r'$#,##0.00', '0%', r'$#,##0.00', None, None, None]
        for col_idx, (val, fmt) in enumerate(zip(values, formats), start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = arial10
            if fmt:
                cell.number_format = fmt

    totals_row = 8 + len(warrants)
    total_p   = sum(w['principal'] for w in warrants)
    total_i   = sum(w['interest']  for w in warrants)
    ws.cell(row=totals_row, column=3, value=' Totals').font = arial10
    for col_idx, val in [(4, total_p), (6, total_i), (8, round(total_p + total_i, 2))]:
        cell = ws.cell(row=totals_row, column=col_idx, value=round(val, 2))
        cell.font = arial10
        cell.number_format = r'$#,##0.00'

    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize   = 9
    return wb


@app.route('/api/gf-calls', methods=['POST'])
@login_required
def gf_calls():
    """
    Accept a CashBalances xlsx/xls file + call date, return a zip of call spreadsheets.
    """
    if 'cash_file' not in request.files:
        return jsonify({'error': 'No cash balances file uploaded'}), 400

    call_date_str = request.form.get('call_date', '')
    call_day_str  = request.form.get('call_day', '')

    if not call_date_str:
        return jsonify({'error': 'Call date is required'}), 400
    try:
        call_date = date.fromisoformat(call_date_str)
    except ValueError:
        return jsonify({'error': 'Invalid call date format'}), 400

    if not call_day_str:
        # Auto-generate e.g. "TUESDAY, MAY 12, 2026"
        call_day_str = call_date.strftime('%A, %B %-d, %Y').upper()

    # Load the uploaded cash balances file
    cash_file  = request.files['cash_file']
    cash_bytes = cash_file.read()
    try:
        fname_lower = cash_file.filename.lower()
        engine = 'xlrd' if fname_lower.endswith('.xls') else 'openpyxl'
        cb_df = pd.read_excel(io.BytesIO(cash_bytes), header=None, engine=engine)
    except Exception as e:
        return jsonify({'error': f'Could not read cash balances file: {e}'}), 400

    # Parse SID rows (skip header rows at top)
    sids_data = []
    for _, row in cb_df.iterrows():
        if row[0] in (None, 'COUNTY') or pd.isna(row[0]):
            continue
        try:
            sid_num = int(row[1])
        except (ValueError, TypeError):
            continue
        sids_data.append({
            'county': str(row[0]).strip(),
            'sid':    sid_num,
            'name':   str(row[2]).strip(),
            'cash':   float(row[6] or 0) + float(row[7] or 0),
        })

    if not sids_data:
        return jsonify({'error': 'No SID rows found in cash balances file'}), 400

    # Load warrant history
    df, _ = _load_warrants_df()
    if df is None:
        return jsonify({'error': 'Warrant history file (Warrants_2014-2026.xlsx) not found on server'}), 500

    # Generate call workbooks and zip them
    zip_buf = io.BytesIO()
    generated = 0
    skipped   = []

    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for entry in sids_data:
            sid      = entry['sid']
            county   = entry['county']
            sid_name = SID_NAMES.get(sid, entry['name'])
            cash     = entry['cash']

            warrants = _gf_calls_select_warrants(df, sid, call_date, cash)

            if not warrants:
                skipped.append(sid)
                continue

            wb    = _gf_calls_build_workbook(sid, sid_name, county, call_day_str, warrants)
            fname = f"S {sid} - GF Call {call_date.strftime('%m-%d-%Y')}.xlsx"
            wb_buf = io.BytesIO()
            wb.save(wb_buf)
            zf.writestr(fname, wb_buf.getvalue())
            generated += 1

    if generated == 0:
        return jsonify({'error': 'No warrants eligible for call found for any SID'}), 400

    zip_buf.seek(0)
    zip_name = f"GF_Calls_{call_date.strftime('%m-%d-%Y')}.zip"
    return send_file(
        zip_buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=zip_name
    )


# ── Warrant Query ─────────────────────────────────────────────────────────────

def _build_warrant_context() -> str:
    """Build a rich text summary of the warrant dataset for use as AI context."""
    try:
        df, file_date = _load_warrants_df()
        if df is None:
            return "Warrant history file (Warrants_2014-2026.xlsx) not found on server."
    except Exception as e:
        return f"Could not load warrant history file: {e}"

    # Normalize column names — strip whitespace
    df.columns = [c.strip() for c in df.columns]

    def col(name):
        """Return column by name (case-insensitive, stripped), or None if missing."""
        for c in df.columns:
            if c.lower() == name.lower():
                return c
        return None

    lines = []
    lines.append("=== WARRANT DATASET SUMMARY ===")
    if file_date:
        lines.append(f"Data as of: {file_date}")
    lines.append(f"Total rows: {len(df):,}")
    lines.append(f"Columns: {', '.join(df.columns.tolist())}")

    if col('County'):
        lines.append(f"\nCounties: {sorted(df[col('County')].dropna().unique().tolist())}")
    if col('SID'):
        lines.append(f"Total SIDs: {df[col('SID')].nunique()}")
    if col('Type'):
        lines.append(f"Warrant types: {df[col('Type')].value_counts().to_dict()}")
    if col('Issued'):
        issued_range = df[col('Issued')].dropna()
        if not issued_range.empty:
            lines.append(f"Issued date range: {issued_range.min().date()} to {issued_range.max().date()}")

    if col('Principal'):
        lines.append(f"\nTotal principal (all warrants): ${df[col('Principal')].sum():,.2f}")
        if col('CallDate') and col('Registration'):
            # "Outstanding" = not yet redeemed. A warrant is outstanding if EITHER its CallDate
            # is blank OR the CallDate is in the future (call scheduled but not yet executed).
            today_ = pd.Timestamp.today().normalize()
            outstanding_mask = (
                df[col('CallDate')].isna() | (df[col('CallDate')] > today_)
            ) & df[col('Registration')].notna()
            outstanding = df[outstanding_mask]
            lines.append(
                f"Outstanding principal (not yet redeemed — CallDate blank OR CallDate in the future): "
                f"${outstanding[col('Principal')].sum():,.2f}"
            )

    # By county
    if col('County') and col('Principal'):
        lines.append("\nPrincipal by county (ALL warrants, historic + outstanding):")
        for county, grp in df.groupby(col('County')):
            lines.append(f"  {county}: ${grp[col('Principal')].sum():,.2f} ({len(grp):,} warrants)")

    # Outstanding by county — answers "what county has the most outstanding warrants"
    # distinctly from the all-warrants totals above. Uses the same outstanding
    # definition as the rest of the app: CallDate blank OR future, AND Registration set.
    if col('County') and col('Principal') and col('CallDate') and col('Registration'):
        today_ = pd.Timestamp.today().normalize()
        out_mask = (
            df[col('CallDate')].isna() | (df[col('CallDate')] > today_)
        ) & df[col('Registration')].notna()
        out_df = df[out_mask]
        if not out_df.empty:
            lines.append(
                "\nOUTSTANDING warrants by county (Registration set AND CallDate blank or future):"
            )
            out_county = out_df.groupby(col('County')).agg(
                w=(col('Principal'), 'count'),
                p=(col('Principal'), 'sum'),
            ).sort_values('p', ascending=False)
            for county, r in out_county.iterrows():
                lines.append(
                    f"  {county}: {int(r['w']):,} warrants, ${r['p']:,.2f}"
                )
            # Flag counties that have history but nothing outstanding
            all_counties = set(df[col('County')].dropna().astype(str).unique())
            out_counties = set(out_df[col('County')].dropna().astype(str).unique())
            missing = all_counties - out_counties
            if missing:
                lines.append(
                    f"  (Counties with historical warrants but NO current outstanding: "
                    f"{', '.join(sorted(missing))})"
                )
            # County × Type breakdown for outstanding
            if col('Type'):
                lines.append(
                    "\nOUTSTANDING warrants by county and type (CF = Construction Fund, GF = General Fund):"
                )
                out_ct = out_df.groupby([col('County'), col('Type')]).agg(
                    w=(col('Principal'), 'count'),
                    p=(col('Principal'), 'sum'),
                ).sort_index()
                for (county, ftype), r in out_ct.iterrows():
                    lines.append(
                        f"  {county} / {ftype}: {int(r['w']):,} warrants, ${r['p']:,.2f}"
                    )

    # By type
    if col('Type') and col('Principal'):
        lines.append("\nPrincipal by type:")
        for t, grp in df.groupby(col('Type')):
            lines.append(f"  {t}: ${grp[col('Principal')].sum():,.2f} ({len(grp):,} warrants)")


    # SID activity status — distinguishes SIDs that are currently issuing warrants
    # (have at least one registration in the current year) from SIDs that have
    # outstanding warrants but no current-year registration activity. Answers
    # questions like "which SIDs have outstanding warrants but haven't registered
    # any in 2026" without needing a separate row-level helper.
    if col('CallDate') and col('Registration') and col('SID') and col('County'):
        today_ = pd.Timestamp.today().normalize()
        current_year = today_.year
        out_mask = (
            df[col('CallDate')].isna() | (df[col('CallDate')] > today_)
        ) & df[col('Registration')].notna()
        out_df = df[out_mask]
        if not out_df.empty:
            sids_outstanding = set(
                out_df.groupby([col('County'), col('SID')]).groups.keys()
            )
            current_year_reg = df[df[col('Registration')].dt.year == current_year]
            sids_current_year = set(
                current_year_reg.groupby([col('County'), col('SID')]).groups.keys()
            )
            inactive = sorted(sids_outstanding - sids_current_year)
            active = sorted(sids_outstanding & sids_current_year)

            lines.append(
                f"\nSID activity status (current year = {current_year}):"
            )
            lines.append(
                f"  Total SIDs with outstanding warrants: {len(sids_outstanding)}"
            )
            lines.append(
                f"  SIDs with at least one warrant registered in {current_year}: "
                f"{len(sids_current_year)}"
            )
            lines.append(
                f"  ACTIVE SIDs (have outstanding AND registered in {current_year}): "
                f"{len(active)}"
            )
            lines.append(
                f"  INACTIVE-CURRENT-YEAR SIDs (have outstanding but NO {current_year} "
                f"registration): {len(inactive)}"
            )

            if inactive:
                lines.append(
                    f"\n  INACTIVE SIDs detail (outstanding but no {current_year} "
                    f"registration). 'Last reg' = most recent Registration date in "
                    f"the dataset for that SID:"
                )
                for cnty, sid in inactive:
                    sub_out = out_df[
                        (out_df[col('County')] == cnty) &
                        (out_df[col('SID')] == sid)
                    ]
                    sub_all = df[
                        (df[col('County')] == cnty) &
                        (df[col('SID')] == sid)
                    ]
                    last_reg = sub_all[col('Registration')].max()
                    last_reg_str = (
                        last_reg.strftime('%m/%d/%Y') if pd.notna(last_reg) else '(none)'
                    )
                    lines.append(
                        f"    {cnty} SID {int(sid)}: "
                        f"{len(sub_out):,} outstanding warrants, "
                        f"${sub_out[col('Principal')].sum():,.2f}, "
                        f"last reg {last_reg_str}"
                    )

            # ACTIVE SIDs detail — outstanding totals + current-year registration counts
            # for every SID with outstanding warrants. Enables filter-style questions
            # like "show SIDs with over $X outstanding registered in current year",
            # "biggest SIDs by outstanding principal", "which SIDs registered the most
            # in 2026", etc. Limited to SIDs with at least $100K outstanding to keep
            # context size reasonable; smaller SIDs are aggregated in a tail summary.
            out_per_sid = out_df.groupby(
                [col('County'), col('SID')]
            ).agg(
                out_w=(col('Principal'), 'count'),
                out_p=(col('Principal'), 'sum'),
            )
            current_year_reg = df[df[col('Registration')].dt.year == current_year]
            cy_per_sid = current_year_reg.groupby(
                [col('County'), col('SID')]
            ).agg(
                cy_w=(col('Principal'), 'count'),
                cy_p=(col('Principal'), 'sum'),
            )
            sid_combo = out_per_sid.join(cy_per_sid, how='left').fillna(0)
            sid_combo = sid_combo.sort_values('out_p', ascending=False)

            threshold = 100_000
            big_sids = sid_combo[sid_combo['out_p'] >= threshold]
            small_sids = sid_combo[sid_combo['out_p'] < threshold]

            lines.append(
                f"\n  Per-SID outstanding + {current_year} registration activity "
                f"(sorted by outstanding principal descending, SIDs with >= "
                f"${threshold:,} outstanding shown individually):"
            )
            lines.append(
                f"    Format: County SID — outstanding warrants, outstanding $; "
                f"{current_year}-registered warrants, {current_year}-registered $"
            )
            for (cnty, sid), r in big_sids.iterrows():
                cy_w = int(r['cy_w'])
                cy_p = r['cy_p']
                cy_part = (
                    f"{cy_w} {current_year}-reg warrants, ${cy_p:,.2f}"
                    if cy_w > 0 else
                    f"NO {current_year} registrations"
                )
                lines.append(
                    f"    {cnty} SID {int(sid)} — "
                    f"{int(r['out_w']):,} outstanding, ${r['out_p']:,.2f}; "
                    f"{cy_part}"
                )

            if not small_sids.empty:
                cy_small = small_sids[small_sids['cy_w'] > 0]
                lines.append(
                    f"    (Plus {len(small_sids)} smaller SIDs each with <${threshold:,} "
                    f"outstanding, totaling ${small_sids['out_p'].sum():,.2f}; "
                    f"{len(cy_small)} of them have at least one {current_year} registration.)"
                )


    # Outstanding warrants summary
    # A warrant is "outstanding" (still held, not yet redeemed) if its CallDate is blank OR
    # if the CallDate is in the future (scheduled redemption that has not yet occurred).
    if col('CallDate') and col('Registration') and col('Principal'):
        today_ = pd.Timestamp.today().normalize()
        outstanding_mask = (
            df[col('CallDate')].isna() | (df[col('CallDate')] > today_)
        ) & df[col('Registration')].notna()
        outstanding = df[outstanding_mask]
        lines.append(f"\nOutstanding warrants (not yet redeemed): {len(outstanding):,}")
        lines.append(f"Outstanding principal: ${outstanding[col('Principal')].sum():,.2f}")

    # Warrants with a future CallDate (scheduled but not yet redeemed)
    if col('CallDate') and col('Principal'):
        today = pd.Timestamp.today().normalize()
        future_called = df[df[col('CallDate')].notna() & (df[col('CallDate')] > today)].copy()
        if not future_called.empty:
            lines.append(f"\nWarrants with a future scheduled CallDate (not yet redeemed): {len(future_called):,}, principal: ${future_called[col('Principal')].sum():,.2f}")
            # Break down by holder and call date, with P+I totals
            if col('Holder'):
                fc_holders = future_called[future_called[col('Holder')].notna()].copy()
                if not fc_holders.empty:
                    # Interest may be in either InterestPaid or InterestAccrued depending on record;
                    # compute a combined interest that uses whichever is populated.
                    ip_col = col('InterestPaid')
                    ia_col = col('InterestAccrued')
                    if ip_col and ia_col:
                        fc_holders['_interest'] = (
                            fc_holders[ip_col].fillna(0) + fc_holders[ia_col].fillna(0)
                        )
                    elif ip_col:
                        fc_holders['_interest'] = fc_holders[ip_col].fillna(0)
                    elif ia_col:
                        fc_holders['_interest'] = fc_holders[ia_col].fillna(0)
                    else:
                        fc_holders['_interest'] = 0.0
                    fc_by_holder = (
                        fc_holders.groupby([col('Holder'), col('CallDate')])
                        .agg(
                            warrants=(col('Principal'), 'count'),
                            principal=(col('Principal'), 'sum'),
                            interest=('_interest', 'sum'),
                        )
                        .reset_index()
                        .sort_values(col('CallDate'))
                    )
                    lines.append("  Future-called warrants by holder and call date (P=Principal, I=Interest, P+I=total):")
                    for _, row in fc_by_holder.iterrows():
                        call_dt = row[col('CallDate')]
                        call_str = call_dt.strftime('%m/%d/%Y') if hasattr(call_dt, 'strftime') else str(call_dt)
                        pi_total = row['principal'] + row['interest']
                        lines.append(
                            f"    {row[col('Holder')]} — call date {call_str}: "
                            f"{int(row['warrants'])} warrants, P=${row['principal']:,.2f}, "
                            f"I=${row['interest']:,.2f}, P+I=${pi_total:,.2f}"
                        )

            # Break down by call date / county / SID so SID-level questions are answerable
            if col('County') and col('SID'):
                fc_sid = future_called.copy()
                ip_col2 = col('InterestPaid')
                ia_col2 = col('InterestAccrued')
                if ip_col2 and ia_col2:
                    fc_sid['_interest'] = fc_sid[ip_col2].fillna(0) + fc_sid[ia_col2].fillna(0)
                elif ip_col2:
                    fc_sid['_interest'] = fc_sid[ip_col2].fillna(0)
                elif ia_col2:
                    fc_sid['_interest'] = fc_sid[ia_col2].fillna(0)
                else:
                    fc_sid['_interest'] = 0.0
                fc_by_sid = (
                    fc_sid.groupby([col('CallDate'), col('County'), col('SID')])
                    .agg(
                        warrants=(col('Principal'), 'count'),
                        principal=(col('Principal'), 'sum'),
                        interest=('_interest', 'sum'),
                    )
                    .reset_index()
                    .sort_values([col('CallDate'), col('County'), col('SID')])
                )
                lines.append("  Future-called warrants by call date / county / SID (P=Principal, I=Interest, P+I=total):")
                for _, row in fc_by_sid.iterrows():
                    call_dt = row[col('CallDate')]
                    call_str = call_dt.strftime('%m/%d/%Y') if hasattr(call_dt, 'strftime') else str(call_dt)
                    pi_total = row['principal'] + row['interest']
                    lines.append(
                        f"    {call_str} — {row[col('County')]} SID {int(row[col('SID')])}: "
                        f"{int(row['warrants'])} warrants, P=${row['principal']:,.2f}, "
                        f"I=${row['interest']:,.2f}, P+I=${pi_total:,.2f}"
                    )

            # Break down by holder / call date / county / SID so a question like
            # "FIRNBANK warrants on 5/12/26" can be answered with the SID breakdown.
            if col('Holder') and col('County') and col('SID'):
                fc_hsid = future_called[future_called[col('Holder')].notna()].copy()
                ip_col3 = col('InterestPaid')
                ia_col3 = col('InterestAccrued')
                if ip_col3 and ia_col3:
                    fc_hsid['_interest'] = fc_hsid[ip_col3].fillna(0) + fc_hsid[ia_col3].fillna(0)
                elif ip_col3:
                    fc_hsid['_interest'] = fc_hsid[ip_col3].fillna(0)
                elif ia_col3:
                    fc_hsid['_interest'] = fc_hsid[ia_col3].fillna(0)
                else:
                    fc_hsid['_interest'] = 0.0
                fc_by_hsid = (
                    fc_hsid.groupby([col('Holder'), col('CallDate'), col('County'), col('SID')])
                    .agg(
                        warrants=(col('Principal'), 'count'),
                        principal=(col('Principal'), 'sum'),
                        interest=('_interest', 'sum'),
                    )
                    .reset_index()
                    .sort_values([col('Holder'), col('CallDate'), col('County'), col('SID')])
                )
                lines.append("  Future-called warrants by holder / call date / county / SID (P=Principal, I=Interest, P+I=total):")
                for _, row in fc_by_hsid.iterrows():
                    call_dt = row[col('CallDate')]
                    call_str = call_dt.strftime('%m/%d/%Y') if hasattr(call_dt, 'strftime') else str(call_dt)
                    pi_total = row['principal'] + row['interest']
                    lines.append(
                        f"    {row[col('Holder')]} — {call_str} — {row[col('County')]} SID {int(row[col('SID')])}: "
                        f"{int(row['warrants'])}w, P=${row['principal']:,.2f}, "
                        f"I=${row['interest']:,.2f}, P+I=${pi_total:,.2f}"
                    )

            # Distinct future call dates for quick reference
            unique_dates = sorted(future_called[col('CallDate')].dropna().unique())
            date_strs = [pd.Timestamp(d).strftime('%m/%d/%Y') for d in unique_dates]
            lines.append(f"  Distinct future call dates ({len(date_strs)}): {', '.join(date_strs)}")

    # All holders — outstanding warrants (not yet redeemed: CallDate blank OR in the future)
    if col('Holder') and col('Principal') and col('CallDate'):
        today_ = pd.Timestamp.today().normalize()
        outstanding_with_holder = df[
            (df[col('CallDate')].isna() | (df[col('CallDate')] > today_)) &
            df[col('Holder')].notna()
        ].copy()
        holder_summary = (
            outstanding_with_holder.groupby(col('Holder'))
            .agg(warrants=(col('Principal'), 'count'), principal=(col('Principal'), 'sum'))
            .sort_values('principal', ascending=False)
        )
        lines.append(
            f"\nAll holders — outstanding warrants (not yet redeemed: CallDate blank OR "
            f"CallDate in the future), {len(holder_summary)} holders:"
        )
        for holder, row in holder_summary.iterrows():
            lines.append(f"  {holder}: {int(row['warrants'])} warrants, ${row['principal']:,.2f}")

    # Top payees — outstanding warrants (not yet redeemed: CallDate blank OR in the future)
    # Limited to top 50 by principal; the full payee × expense × year cross-tab below contains
    # all payees, so the long tail is covered there without duplicating thousands of lines here.
    if col('Payee') and col('Principal') and col('CallDate'):
        today_ = pd.Timestamp.today().normalize()
        outstanding_payees = df[
            df[col('CallDate')].isna() | (df[col('CallDate')] > today_)
        ].copy()
        payee_summary = (
            outstanding_payees.groupby(col('Payee'))[col('Principal')]
            .agg(warrants='count', principal='sum')
            .sort_values('principal', ascending=False)
        )
        total_payees = len(payee_summary)
        top_payees = payee_summary.head(50)
        lines.append(
            f"\nTop 50 payees — outstanding warrants (not yet redeemed), "
            f"of {total_payees} total outstanding payees:"
        )
        for payee, row in top_payees.iterrows():
            lines.append(f"  {payee}: {int(row['warrants'])} warrants, ${row['principal']:,.2f}")

    # Called warrants by holder and year — PAST calls only (CallDate <= today).
    # Future-scheduled calls are listed separately in the 'Future-called warrants' block
    # above so the two never conflate.
    if col('Holder') and col('CallDate') and col('Principal') and col('InterestPaid'):
        today_ = pd.Timestamp.today().normalize()
        called = df[
            df[col('Holder')].notna() &
            df[col('CallDate')].notna() &
            (df[col('CallDate')] <= today_)
        ].copy()
        called['_call_year'] = called[col('CallDate')].dt.year
        called_summary = (
            called.groupby([col('Holder'), '_call_year'])
            .agg(
                warrants=(col('Principal'), 'count'),
                principal=(col('Principal'), 'sum'),
                interest=(col('InterestPaid'), 'sum')
            )
            .reset_index()
            .sort_values([col('Holder'), '_call_year'])
        )
        lines.append(
            f"\nCalled (redeemed) warrants by holder and call-year "
            f"(P=Principal paid, I=Interest paid, P+I=total paid). "
            f"Only includes past calls (CallDate <= today); upcoming calls are in the "
            f"'Future-called warrants' block above."
        )
        for holder, grp in called_summary.groupby(col('Holder')):
            year_parts = ', '.join(
                f"{int(r['_call_year'])}: {int(r['warrants'])}w, P=${r['principal']:,.2f}, "
                f"I=${r['interest']:,.2f}, P+I=${r['principal']+r['interest']:,.2f}"
                for _, r in grp.iterrows()
            )
            lines.append(f"  {holder} — {year_parts}")

    # Payee × Expense × Year issuances — all warrants (not just outstanding)
    # This cross-tab is critical: many payees (e.g. Bluestem Capital Partners) receive
    # multiple distinct expense categories. Questions like "how much in Financial Advisory
    # Fees were issued to Bluestem in 2025" MUST be answered from this breakdown, not from
    # a payee-only total.
    # Full breakdown emitted for "significant" payees (recurring vendors, multi-category
    # payees, or payees totaling >= $25k). One-off tiny vendors are summarized in one line.
    if col('Payee') and col('Issued') and col('Principal'):
        df_issued = df[df[col('Issued')].notna() & df[col('Payee')].notna()].copy()
        df_issued['_year'] = df_issued[col('Issued')].dt.year
        exp_col = col('Expense')
        if exp_col:
            # Strip whitespace so e.g. "PFA Disclosure Counsel " normalizes to "PFA Disclosure Counsel"
            df_issued['_expense'] = df_issued[exp_col].astype('string').str.strip().fillna('(unspecified)')
            df_issued.loc[df_issued['_expense'] == '', '_expense'] = '(unspecified)'
        else:
            df_issued['_expense'] = '(unspecified)'

        # Classify payees
        payee_stats = df_issued.groupby(col('Payee')).agg(
            n_expenses=('_expense', 'nunique'),
            n_warrants=(col('Principal'), 'count'),
            total=(col('Principal'), 'sum'),
        )
        significant_mask = (
            (payee_stats['n_expenses'] > 1) |
            (payee_stats['n_warrants'] >= 3) |
            (payee_stats['total'] >= 25000)
        )
        sig_payees = set(payee_stats[significant_mask].index)
        tail_payees = set(payee_stats[~significant_mask].index)

        df_sig = df_issued[df_issued[col('Payee')].isin(sig_payees)]
        payee_exp_year = (
            df_sig.groupby([col('Payee'), '_expense', '_year'])
            .agg(warrants=(col('Principal'), 'count'), principal=(col('Principal'), 'sum'))
            .reset_index()
            .sort_values([col('Payee'), '_expense', '_year'])
        )
        lines.append(
            f"\nWarrant issuances by payee, expense category, and year "
            f"(all warrants including called; {len(sig_payees)} significant payees shown in detail, "
            f"{len(tail_payees)} one-off minor payees summarized below):"
        )
        lines.append("Format: Payee — ExpenseCategory [year: Nw/$amount, ...] | ExpenseCategory [...]")
        for payee, grp in payee_exp_year.groupby(col('Payee')):
            parts = []
            for expense, g2 in grp.groupby('_expense'):
                yr_parts = ', '.join(
                    f"{int(r['_year'])}: {int(r['warrants'])}w/${r['principal']:,.2f}"
                    for _, r in g2.iterrows()
                )
                parts.append(f"{expense} [{yr_parts}]")
            lines.append(f"  {payee} — {' | '.join(parts)}")

        if tail_payees:
            tail_stats = payee_stats.loc[list(tail_payees)]
            lines.append(
                f"\nMinor one-off payees (each <3 warrants, single expense category, "
                f"<$25k total): {len(tail_payees)} payees, combined {int(tail_stats['n_warrants'].sum())} warrants, "
                f"${tail_stats['total'].sum():,.2f} total principal. "
                f"For specific questions about these payees, refer the user to the underlying dataset."
            )

    # Expense category issuances by year
    if col('Expense') and col('Issued') and col('Principal'):
        df_exp = df[df[col('Issued')].notna() & df[col('Expense')].notna()].copy()
        df_exp['_year'] = df_exp[col('Issued')].dt.year
        exp_year = (
            df_exp.groupby([col('Expense'), '_year'])
            .agg(warrants=(col('Principal'), 'count'), principal=(col('Principal'), 'sum'))
            .reset_index()
            .sort_values([col('Expense'), '_year'])
        )
        lines.append(f"\nWarrant issuances by expense category and year:")
        for expense, grp in exp_year.groupby(col('Expense')):
            year_parts = ', '.join(
                f"{int(r['_year'])}: {int(r['warrants'])} warrants (${r['principal']:,.2f})"
                for _, r in grp.iterrows()
            )
            lines.append(f"  {expense} — {year_parts}")

    # Annual interest payment schedule — pulls the payment month/day from 8038_Info.xlsx
    # and joins with outstanding warrant totals so questions like "what SIDs pay interest
    # in May" can be answered directly with principal and accrued interest per SID.
    try:
        if INFO_8038 is not None and not INFO_8038.empty and \
                col('CallDate') and col('Registration') and col('Principal') and \
                col('County') and col('SID'):
            info = INFO_8038.copy()
            info.columns = [c.strip() for c in info.columns]
            ipd_col = 'First Annual Interest Payment'
            if ipd_col in info.columns:
                info['_ipd'] = pd.to_datetime(info[ipd_col], errors='coerce')
                info['_month'] = info['_ipd'].dt.month
                info['_mmdd'] = info['_ipd'].dt.strftime('%m/%d')
                info['_county_u'] = info['County'].astype(str).str.upper().str.strip()

                today_ = pd.Timestamp.today().normalize()
                # Annual interest payments apply ONLY to CF (Construction Fund) warrants.
                # GF (General Fund) warrants do not pay annual interest, so they're excluded
                # from this schedule entirely.
                type_col = col('Type')
                outstanding_w = df[
                    (df[col('CallDate')].isna() | (df[col('CallDate')] > today_)) &
                    df[col('Registration')].notna()
                ].copy()
                if type_col:
                    outstanding_w = outstanding_w[outstanding_w[type_col] == 'CF'].copy()
                # Combined interest: accrued on blank-call warrants, paid on future-call warrants
                ip_col_is = col('InterestPaid')
                ia_col_is = col('InterestAccrued')
                if ip_col_is and ia_col_is:
                    outstanding_w['_interest'] = (
                        outstanding_w[ia_col_is].fillna(0) + outstanding_w[ip_col_is].fillna(0)
                    )
                elif ia_col_is:
                    outstanding_w['_interest'] = outstanding_w[ia_col_is].fillna(0)
                elif ip_col_is:
                    outstanding_w['_interest'] = outstanding_w[ip_col_is].fillna(0)
                else:
                    outstanding_w['_interest'] = 0.0
                outstanding_w['_county_u'] = outstanding_w[col('County')].astype(str).str.upper().str.strip()

                sid_totals = outstanding_w.groupby(['_county_u', col('SID')]).agg(
                    warrants=(col('Principal'), 'count'),
                    principal=(col('Principal'), 'sum'),
                    interest=('_interest', 'sum'),
                ).reset_index()

                # Only SIDs with CF warrants CURRENTLY outstanding appear on the schedule.
                # A SID with only GF outstanding has no annual CF interest to pay this cycle,
                # even if it has historical CF or the 8038 info lists a payment date. This
                # matches the actual interest-payment workflow: if there's no CF principal
                # outstanding, there's nothing to accrue or pay.
                cf_current_sids = set(
                    outstanding_w.groupby(['_county_u', col('SID')]).groups.keys()
                )

                merged = info.merge(
                    sid_totals,
                    on=['_county_u', col('SID')],
                    how='left'
                )
                merged = merged[
                    merged.apply(
                        lambda r: (r['_county_u'], r[col('SID')]) in cf_current_sids,
                        axis=1
                    )
                ].copy()

                months = ['January','February','March','April','May','June',
                          'July','August','September','October','November','December']
                lines.append(
                    "\nAnnual interest payment schedule (from 8038_Info.xlsx, CF warrants only). "
                    "Each SID pays its CF warrant-holders interest once per year on the date shown. "
                    "GF (General Fund) warrants do NOT pay annual interest and are excluded from "
                    "this schedule. SIDs with only GF activity do not appear here at all. "
                    "P = outstanding CF principal, I = accrued interest on those outstanding CF warrants."
                )
                for m in range(1, 13):
                    month_sids = merged[merged['_month'] == m].sort_values(['County', col('SID')])
                    if month_sids.empty:
                        continue
                    mname = months[m-1]
                    mtot_p = month_sids['principal'].fillna(0).sum()
                    mtot_i = month_sids['interest'].fillna(0).sum()
                    lines.append(
                        f"\n{mname} — {len(month_sids)} SIDs pay interest this month "
                        f"(combined P=${mtot_p:,.2f}, I=${mtot_i:,.2f}):"
                    )
                    for _, r in month_sids.iterrows():
                        ipd_str = r['_mmdd'] if pd.notna(r['_ipd']) else 'n/a'
                        if pd.notna(r['warrants']):
                            w = int(r['warrants'])
                            p = r['principal']
                            i = r['interest']
                            lines.append(
                                f"  {ipd_str} — {r['County']} SID {int(r[col('SID')])}: "
                                f"{w}w, P=${p:,.2f}, I=${i:,.2f}"
                            )
                        else:
                            # SID is in 8038 info but has no outstanding warrants right now.
                            # Still list it so the interest-schedule answer includes every
                            # SID on the calendar.
                            lines.append(
                                f"  {ipd_str} — {r['County']} SID {int(r[col('SID')])}: "
                                f"no outstanding warrants"
                            )
    except Exception as e:
        lines.append(f"\n(Interest payment schedule unavailable: {e})")

    lines.append("\n=== END SUMMARY ===")
    lines.append("\nYou have full knowledge of this dataset. Answer the user's question accurately and concisely.")
    lines.append("When referencing dollar amounts, format with commas and 2 decimal places.")
    lines.append("\nIMPORTANT: 'Outstanding' means 'not yet redeemed' — a warrant is outstanding if its CallDate is blank OR if the CallDate is in the future (a call has been scheduled but the redemption has not yet occurred). The outstanding principal figures INCLUDE warrants with a future scheduled CallDate.")
    lines.append("When answering questions about a holder's total principal held, the figure in the 'All holders' block is already correct — it includes both blank-CallDate warrants and future-called warrants. When relevant, also note any upcoming scheduled redemption dates from the 'Future-called warrants by holder and call date' block so the user has the complete picture.")
    lines.append("If the question requires specific row-level data not in this summary, say so and provide the best answer you can from the available context.")

    return '\n'.join(lines)


# Build warrant context at startup in background so first query isn't slow
_warrant_context_cache = None
_warrant_context_lock = threading.Lock()

# Cached DataFrame used by both the context builder and the row-level helper.
# Loading the 12MB Excel takes ~3s; caching it avoids re-reads on every query.
_warrants_df_cache = None
_warrants_df_lock = threading.Lock()

def _get_warrants_df_cached():
    """Return the warrants DataFrame, loading once and caching for the process lifetime."""
    global _warrants_df_cache
    with _warrants_df_lock:
        if _warrants_df_cache is None:
            try:
                df, _ = _load_warrants_df()
                if df is not None:
                    df.columns = [c.strip() for c in df.columns]
                _warrants_df_cache = df
            except Exception:
                _warrants_df_cache = None
        return _warrants_df_cache


# Cached groupby used by the amount-only helper's sum-pair detection. Grouping
# the full 193K-row DataFrame takes 600ms-1.7s; caching avoids paying that cost
# on every amount query.
_warrants_sum_groups_cache = None
_warrants_sum_groups_lock = threading.Lock()

def _get_warrants_sum_groups_cached():
    """Return a DataFrame grouped by (Payee, County, SID, Issued) with count and total."""
    global _warrants_sum_groups_cache
    with _warrants_sum_groups_lock:
        if _warrants_sum_groups_cache is None:
            df = _get_warrants_df_cached()
            if df is None or df.empty:
                return None
            try:
                df_grp = df.copy()
                df_grp['Issued'] = pd.to_datetime(df_grp['Issued'], errors='coerce')
                g = df_grp.groupby(
                    ['Payee', 'County', 'SID', 'Issued'],
                    dropna=False,
                ).agg(
                    count=('Principal', 'count'),
                    total=('Principal', 'sum'),
                ).reset_index()
                # Keep only groups with at least 2 warrants — singles can't sum
                g = g[g['count'] >= 2]
                _warrants_sum_groups_cache = g
            except Exception:
                _warrants_sum_groups_cache = None
        return _warrants_sum_groups_cache

def _get_warrant_context() -> str:
    global _warrant_context_cache
    with _warrant_context_lock:
        if _warrant_context_cache is None:
            _warrant_context_cache = _build_warrant_context()
        return _warrant_context_cache

def _warm_warrant_context():
    """Pre-build the warrant context cache at startup."""
    try:
        _get_warrants_df_cached()  # Warm DataFrame cache first
        _get_warrants_sum_groups_cached()  # Warm groupby for amount-only helper
        _get_warrant_context()
    except Exception:
        pass

threading.Thread(target=_warm_warrant_context, daemon=True).start()


def _detect_payee_from_question(q_lower: str, payee_pool) -> str:
    """
    Given a lowercased question and an iterable of candidate payee names, find the
    best matching payee using the same logic used elsewhere (full-name match,
    significant-word match, or alias map with word-boundary matching). Returns
    the matched payee name from the pool, or None.
    """
    import re as _re

    best_match = None
    best_len = 0
    generic_stop = {'omaha','public','power','district','inc','llc','the','and','of','for',
                    'inc.','llc.','ltd','group','service','services','company','corp',
                    'corporation','capital','securities','bank','trust','investment'}

    for p in payee_pool:
        p_lower = str(p).lower()
        if p_lower in q_lower and len(p_lower) > best_len:
            best_match = p
            best_len = len(p_lower)
        for word in p_lower.split():
            cleaned = word.rstrip('.,').rstrip()
            if len(cleaned) >= 4 and cleaned in q_lower and cleaned not in generic_stop:
                if len(cleaned) > best_len:
                    best_match = p
                    best_len = len(cleaned)

    # Alias map with word-boundary matching
    for alias, canonical in PAYEE_ALIASES.items():
        if _re.search(r'\b' + _re.escape(alias) + r'\b', q_lower):
            for p in payee_pool:
                if canonical.lower() in str(p).lower():
                    if len(canonical) > best_len:
                        best_match = p
                        best_len = len(canonical)
                        break
    return best_match


def _find_payee_amount_rows(question: str, max_rows: int = 20) -> str:
    """
    Detect when a question is asking whether a specific payee ever received a
    payment of a specific dollar amount (e.g. "has there ever been a payment to
    Bluestem for $3,002.85", "did we pay OPPD $5,000 in SID 366"). If so, search
    the entire DataFrame for matching rows and return a formatted block.

    Complements _find_payee_sid_rows by handling payee+amount questions where no
    SID is specified.
    """
    import re as _re

    q = question.lower()

    # Look for a dollar amount. Supports $3,002.85, $3002.85, 3,002.85, 3002.85.
    # Require either a $ sign or a decimal or comma grouping so we don't match
    # arbitrary integers like "in 365" (we're looking for money, not SIDs).
    amt_match = _re.search(
        r'\$?\s*(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+\.\d{1,2})',
        q
    )
    if not amt_match:
        return ''
    amt_str = amt_match.group(1).replace(',', '')
    try:
        amount = float(amt_str)
    except ValueError:
        return ''
    # Skip tiny amounts — likely not money references
    if amount < 1:
        return ''

    df = _get_warrants_df_cached()
    if df is None:
        return ''

    # Try to identify a payee from the question using any payee in the dataset
    all_payees = df['Payee'].dropna().unique().tolist()
    payee = _detect_payee_from_question(q, all_payees)
    if not payee:
        return ''

    # Match by principal rounded to 2 decimal places (tolerates floating point)
    matches = df[
        (df['Payee'] == payee) &
        (df['Principal'].round(2) == round(amount, 2))
    ].copy()

    if matches.empty:
        # Still return a block saying "no match found" so the model doesn't fall
        # back to the vague "can't confirm from summary" response.
        return (
            f"\n=== PAYEE+AMOUNT LOOKUP: {payee} for ${amount:,.2f} ===\n"
            f"No warrants to {payee} with principal exactly ${amount:,.2f} were found "
            f"in the dataset.\n"
            f"=== END PAYEE+AMOUNT LOOKUP ==="
        )

    # Sort by Issued date desc
    if 'Issued' in matches.columns:
        matches['Issued'] = pd.to_datetime(matches['Issued'], errors='coerce')
        matches = matches.sort_values('Issued', ascending=False, na_position='last')
    if 'CheckDate' in matches.columns:
        matches['CheckDate'] = pd.to_datetime(matches['CheckDate'], errors='coerce')
    if 'Registration' in matches.columns:
        matches['Registration'] = pd.to_datetime(matches['Registration'], errors='coerce')
    if 'CallDate' in matches.columns:
        matches['CallDate'] = pd.to_datetime(matches['CallDate'], errors='coerce')

    matches = matches.head(max_rows)

    def fmt_date(v):
        if pd.isna(v):
            return '(none)'
        try:
            return v.strftime('%m/%d/%Y')
        except Exception:
            return str(v)

    block = [
        f"\n=== PAYEE+AMOUNT LOOKUP: {payee} for ${amount:,.2f} ===",
        f"Found {len(matches)} matching warrant(s):",
        "Columns: County | SID | WarrantNo | Type | Issued | Principal | CheckDate | Expense",
    ]
    for _, r in matches.iterrows():
        wnum = int(r['Number']) if pd.notna(r.get('Number')) else '?'
        wtype = r.get('Type', '?') or '?'
        county = r.get('County', '?') or '?'
        sid = int(r['SID']) if pd.notna(r.get('SID')) else '?'
        issued = fmt_date(r.get('Issued'))
        principal = r.get('Principal')
        pstr = f"${principal:,.2f}" if pd.notna(principal) else '$?'
        check = fmt_date(r.get('CheckDate'))
        expense = r.get('Expense') or '(unspecified)'
        block.append(
            f"  {county} | SID {sid} | #{wnum} | {wtype} | Issued {issued} | "
            f"{pstr} | CheckDate {check} | {expense}"
        )
    block.append("=== END PAYEE+AMOUNT LOOKUP ===")
    return '\n'.join(block)


def _find_date_rows(question: str, max_rows_detail: int = 150) -> str:
    """
    Detect when a question is asking about warrants associated with a specific date,
    where the date could refer to Issued, Registration, CheckDate, or CallDate
    (e.g. "what warrants were registered 04/20/2026", "warrants issued 3/5/26",
    "checks dated 4/15/26"). Returns a formatted block listing matching warrants
    grouped by county/SID with totals and selected row-level detail, or empty.

    Complements the payee+SID and payee+amount helpers for date-based questions
    that don't specify a payee.
    """
    import re as _re

    q = question.lower()

    # Parse a date (MM/DD/YY, MM/DD/YYYY, M-D-YY, etc). Require 4-digit year or
    # 2-digit year with a separator so we don't match SID numbers.
    date_patterns = [
        r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b',   # MM/DD/YYYY or M-D-YYYY
        r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{2})\b',    # MM/DD/YY
    ]
    target_date = None
    for pat in date_patterns:
        m = _re.search(pat, q)
        if m:
            mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if yy < 100:
                yy += 2000
            try:
                target_date = pd.Timestamp(year=yy, month=mm, day=dd)
                break
            except (ValueError, pd.errors.OutOfBoundsDatetime):
                continue
    if target_date is None:
        return ''

    # Determine which date column the question is asking about
    if any(word in q for word in ['registered', 'registration']):
        date_col = 'Registration'
        verb = 'registered'
    elif any(word in q for word in ['issued', 'issue date']):
        date_col = 'Issued'
        verb = 'issued'
    elif any(word in q for word in ['check', 'checked', 'check date', 'checks dated']):
        date_col = 'CheckDate'
        verb = 'with check date'
    elif any(word in q for word in ['called', 'call date']):
        date_col = 'CallDate'
        verb = 'called'
    else:
        # Not clearly a date-based lookup question — skip
        return ''

    df = _get_warrants_df_cached()
    if df is None:
        return ''

    if date_col not in df.columns:
        return ''

    col_series = pd.to_datetime(df[date_col], errors='coerce')
    matches = df[col_series == target_date].copy()
    if matches.empty:
        return (
            f"\n=== DATE LOOKUP: warrants {verb} {target_date.strftime('%m/%d/%Y')} ===\n"
            f"No warrants found with {date_col} = {target_date.strftime('%m/%d/%Y')}.\n"
            f"=== END DATE LOOKUP ==="
        )

    # Normalize helper columns
    for c in ('Issued', 'Registration', 'CheckDate', 'CallDate'):
        if c in matches.columns:
            matches[c] = pd.to_datetime(matches[c], errors='coerce')

    def fmt_date(v):
        if pd.isna(v):
            return '(none)'
        try:
            return v.strftime('%m/%d/%Y')
        except Exception:
            return str(v)

    total_p = matches['Principal'].sum() if 'Principal' in matches.columns else 0

    block = [
        f"\n=== DATE LOOKUP: warrants {verb} {target_date.strftime('%m/%d/%Y')} ===",
        f"Found {len(matches)} warrant(s), total principal ${total_p:,.2f}.",
    ]

    # Summary by County / SID
    if 'County' in matches.columns and 'SID' in matches.columns:
        sid_summary = matches.groupby(['County', 'SID']).agg(
            w=('Principal', 'count'),
            p=('Principal', 'sum'),
        ).reset_index().sort_values(['County', 'SID'])
        block.append("  By County / SID:")
        for _, r in sid_summary.iterrows():
            block.append(
                f"    {r['County']} SID {int(r['SID'])}: "
                f"{int(r['w'])} warrants, ${r['p']:,.2f}"
            )

    # Expense breakdown
    if 'Expense' in matches.columns:
        exp_summary = matches['Expense'].fillna('(unspecified)').value_counts()
        block.append("  By Expense category:")
        for exp, cnt in exp_summary.items():
            exp_total = matches[matches['Expense'].fillna('(unspecified)') == exp]['Principal'].sum()
            block.append(f"    {exp}: {cnt} warrants, ${exp_total:,.2f}")

    # Row-level detail — cap to avoid ballooning context for large dates
    if len(matches) <= max_rows_detail:
        # Sort by county, SID, warrant number
        detail = matches.sort_values(['County', 'SID', 'Number']) \
            if 'Number' in matches.columns else matches
        block.append("  Row-level detail:")
        block.append("  Columns: County | SID | WarrantNo | Type | Payee | Principal | "
                     "Issued | Registration | CheckDate | CallDate | Expense")
        for _, r in detail.iterrows():
            wnum = int(r['Number']) if pd.notna(r.get('Number')) else '?'
            wtype = r.get('Type', '?') or '?'
            county = r.get('County', '?') or '?'
            sid = int(r['SID']) if pd.notna(r.get('SID')) else '?'
            payee = r.get('Payee') or '(unknown)'
            principal = r.get('Principal')
            pstr = f"${principal:,.2f}" if pd.notna(principal) else '$?'
            block.append(
                f"    {county} | SID {sid} | #{wnum} | {wtype} | {payee} | {pstr} | "
                f"Issued {fmt_date(r.get('Issued'))} | "
                f"Reg {fmt_date(r.get('Registration'))} | "
                f"Check {fmt_date(r.get('CheckDate'))} | "
                f"Call {fmt_date(r.get('CallDate'))} | "
                f"{r.get('Expense') or '(unspecified)'}"
            )
    else:
        block.append(
            f"  (Row-level detail omitted — {len(matches)} warrants exceeds the "
            f"{max_rows_detail}-row inline limit. Summary totals above are complete.)"
        )

    block.append("=== END DATE LOOKUP ===")
    return '\n'.join(block)


def _find_holder_rows(question: str, max_rows_detail: int = 150) -> str:
    """
    Detect when a question is asking about warrants held by a specific holder
    (e.g. "show me warrants held by David B. Henning", "what does Barker
    Revocable Trust hold", "FIRNBANK warrants outstanding"). Returns a formatted
    block summarizing the holdings with per-SID totals and row-level detail
    when reasonably sized.

    Distinct from the payee helpers — the Holder column identifies who
    currently owns each warrant, versus Payee which records who was paid.
    """
    import re as _re

    q = question.lower()

    # Keyword trigger — "held", "holds", "holding", "holder", "outstanding to <name>",
    # or explicit holder language
    holder_triggers = [
        'held by', 'holding', 'holder', 'holds',
        'warrants held', 'what does', "what's held",
        'outstanding to', 'owned by',
    ]
    trigger_hit = any(trig in q for trig in holder_triggers)
    # Also trigger on direct "<Holder Name> warrants" pattern if we can find the name
    # in the holder list below.

    df = _get_warrants_df_cached()
    if df is None or 'Holder' not in df.columns:
        return ''

    # Find the best-matching holder using shared payee-detection logic.
    # Holder names are a small finite list, so iterate that instead of all payees.
    all_holders = df['Holder'].dropna().unique().tolist()
    holder = _detect_payee_from_question(q, all_holders)
    if not holder:
        return ''

    # If we found a holder but no trigger word was present, only proceed when the
    # holder name is substantial enough that the match is unambiguous
    # (>= 8 chars) — otherwise decline to avoid hijacking unrelated questions.
    if not trigger_hit and len(str(holder)) < 8:
        return ''

    today_ = pd.Timestamp.today().normalize()
    call_col = pd.to_datetime(df['CallDate'], errors='coerce') if 'CallDate' in df.columns else None
    reg_col = pd.to_datetime(df['Registration'], errors='coerce') if 'Registration' in df.columns else None

    # Decide scope: outstanding vs. all (including called). Default to outstanding,
    # since that's what "held" typically means. If the user says "all", "ever",
    # "lifetime", "total history", include called.
    if any(word in q for word in ['all ', 'ever ', 'lifetime', 'history', 'total history', 'called', 'redeemed']):
        scope = 'all'
        rows = df[df['Holder'] == holder].copy()
    else:
        scope = 'outstanding'
        if call_col is None:
            return ''
        mask = (df['Holder'] == holder) & (call_col.isna() | (call_col > today_))
        if reg_col is not None:
            mask = mask & reg_col.notna()
        rows = df[mask].copy()

    if rows.empty:
        return (
            f"\n=== HOLDER LOOKUP: {holder} ({scope}) ===\n"
            f"No {scope} warrants found for {holder}.\n"
            f"=== END HOLDER LOOKUP ==="
        )

    # Normalize date columns
    for c in ('Issued', 'Registration', 'CheckDate', 'CallDate'):
        if c in rows.columns:
            rows[c] = pd.to_datetime(rows[c], errors='coerce')

    def fmt_date(v):
        if pd.isna(v):
            return '(none)'
        try:
            return v.strftime('%m/%d/%Y')
        except Exception:
            return str(v)

    # Combined interest
    int_combined = pd.Series(0.0, index=rows.index)
    if 'InterestAccrued' in rows.columns:
        int_combined = int_combined + rows['InterestAccrued'].fillna(0)
    if 'InterestPaid' in rows.columns:
        int_combined = int_combined + rows['InterestPaid'].fillna(0)
    total_p = rows['Principal'].sum() if 'Principal' in rows.columns else 0
    total_i = int_combined.sum()

    block = [
        f"\n=== HOLDER LOOKUP: {holder} ({scope}) ===",
        f"Found {len(rows)} warrant(s), P=${total_p:,.2f}, I=${total_i:,.2f}, "
        f"P+I=${total_p + total_i:,.2f}.",
    ]

    # Per-SID summary
    if 'County' in rows.columns and 'SID' in rows.columns:
        rows['_int'] = int_combined
        sid_sum = rows.groupby(['County', 'SID']).agg(
            w=('Principal', 'count'),
            p=('Principal', 'sum'),
            i=('_int', 'sum'),
        ).reset_index().sort_values(['County', 'SID'])
        block.append("  By County / SID:")
        for _, r in sid_sum.iterrows():
            block.append(
                f"    {r['County']} SID {int(r['SID'])}: "
                f"{int(r['w'])} warrants, P=${r['p']:,.2f}, I=${r['i']:,.2f}"
            )

    # Row-level detail if reasonably sized
    if len(rows) <= max_rows_detail:
        detail = rows.sort_values(['County', 'SID', 'Number']) \
            if 'Number' in rows.columns else rows
        block.append("  Row-level detail:")
        block.append("  Columns: County | SID | WarrantNo | Type | Principal | Interest | "
                     "Issued | Registration | CallDate | Payee | Expense")
        for _, r in detail.iterrows():
            wnum = int(r['Number']) if pd.notna(r.get('Number')) else '?'
            wtype = r.get('Type', '?') or '?'
            county = r.get('County', '?') or '?'
            sid = int(r['SID']) if pd.notna(r.get('SID')) else '?'
            principal = r.get('Principal')
            pstr = f"${principal:,.2f}" if pd.notna(principal) else '$?'
            interest = r.get('_int', 0)
            istr = f"${interest:,.2f}" if pd.notna(interest) else '$0.00'
            payee = r.get('Payee') or '(unknown)'
            block.append(
                f"    {county} | SID {sid} | #{wnum} | {wtype} | P={pstr} | I={istr} | "
                f"Issued {fmt_date(r.get('Issued'))} | "
                f"Reg {fmt_date(r.get('Registration'))} | "
                f"Call {fmt_date(r.get('CallDate'))} | "
                f"{payee} | {r.get('Expense') or '(unspecified)'}"
            )
    else:
        block.append(
            f"  (Row-level detail omitted — {len(rows)} warrants exceeds the "
            f"{max_rows_detail}-row inline limit. Summary totals above are complete.)"
        )

    block.append("=== END HOLDER LOOKUP ===")
    return '\n'.join(block)


def _find_amount_only_rows(question: str, max_rows_detail: int = 60) -> str:
    """
    Detect when a question asks whether any warrant has a specific dollar amount,
    with NO payee specified (e.g. "are there any warrants for $1,735.75",
    "does any warrant have a principal of $50,000"). Returns either a list of
    matching rows or an explicit "no match" block so the model can answer
    definitively instead of falling back to "I cannot confirm from the summary".

    Only fires when the question pattern suggests an amount-only lookup — i.e.
    asks "any" / "does any" / "is there" about a specific principal value,
    without naming a payee, SID, or holder. This keeps it from stealing queries
    that the payee+amount or date helpers should handle.
    """
    import re as _re
    q = question.lower()

    # Parse a dollar amount — same logic as payee+amount helper
    amt_match = _re.search(
        r'\$?\s*(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+\.\d{1,2})',
        q
    )
    if not amt_match:
        return ''
    try:
        amount = float(amt_match.group(1).replace(',', ''))
    except ValueError:
        return ''
    if amount < 1:
        return ''

    # Only trigger when the question has a clear "any/is there/does any" shape.
    trigger_patterns = [
        r'\bany warrant',
        r'\bany warrants',
        r'\bis there\b',
        r'\bare there\b',
        r'\bdoes any\b',
        r'\bdo any\b',
        r'\bdid any\b',
        r'\bwarrant.{1,30}amount',
        r'\bamount of\s*\$?\d',
        r'\bprincipal.{1,30}(of|=|equal)',
        r'\bever a warrant',
        r'\bwas there a warrant',
    ]
    trigger = any(_re.search(pat, q) for pat in trigger_patterns)
    if not trigger:
        return ''

    # If the question also names a payee/holder/SID, defer to the more-specific
    # helpers. Do a quick heuristic scan.
    if _re.search(r'\bsid\s*#?\s*\d', q):
        return ''
    # Check if any payee or holder name is plausibly mentioned. We skip this
    # helper if so — the payee+amount helper will handle it.
    df = _get_warrants_df_cached()
    if df is None:
        return ''
    # Quick heuristic: if the question contains any 3+ letter word found in the
    # top 100 payees or holders, defer.
    top_payees = df['Payee'].dropna().value_counts().head(200).index.tolist()
    top_holders = df['Holder'].dropna().value_counts().head(50).index.tolist() if 'Holder' in df.columns else []
    entity_words = set()
    for name in top_payees + top_holders:
        for w in str(name).lower().split():
            cleaned = w.strip('.,&()')
            if len(cleaned) >= 4 and cleaned not in (
                'inc', 'llc', 'company', 'corp', 'corporation', 'group', 'ltd',
                'bank', 'trust', 'partners', 'services', 'service', 'the',
                'and', 'for', 'fund', 'district',
            ):
                entity_words.add(cleaned)
    for alias in PAYEE_ALIASES.keys():
        entity_words.add(alias)
    if any(_re.search(r'\b' + _re.escape(w) + r'\b', q) for w in entity_words):
        # A named entity is in the question — defer to the more-specific helper
        return ''

    # Exact-principal matches
    matches = df[df['Principal'].round(2) == round(amount, 2)].copy()

    # Sum-pair detection: groups of warrants that share (Payee, County, SID,
    # Issued date) and whose Principal values SUM to the target amount.
    # The classic shape here is a Bluestem fee split across GF and CF on the
    # same issue day — the full fee is the group sum, individual warrants are
    # fund-specific allocations. Tight grouping (same payee+SID+date) keeps
    # noise low: random pairs from different SIDs or payees are excluded.
    # Uses a cached groupby to avoid paying 600ms-1.7s per query.
    sum_groups_all = _get_warrants_sum_groups_cached()
    if sum_groups_all is not None and not sum_groups_all.empty:
        sum_groups = sum_groups_all[
            sum_groups_all['total'].round(2) == round(amount, 2)
        ]
    else:
        sum_groups = pd.DataFrame()

    # For drilling into group rows later we need a parsed-dates DataFrame
    df_grouped = df.copy()
    if 'Issued' in df_grouped.columns:
        df_grouped['Issued'] = pd.to_datetime(df_grouped['Issued'], errors='coerce')

    # If neither exact matches nor sum-pair groups exist, report clean negative
    if matches.empty and sum_groups.empty:
        return (
            f"\n=== AMOUNT-ONLY LOOKUP: principal = ${amount:,.2f} ===\n"
            f"No warrants found in the dataset with principal exactly ${amount:,.2f}, "
            f"and no groups of warrants sharing the same payee+SID+issue-date were found "
            f"whose principal values sum to ${amount:,.2f}. "
            f"This is an authoritative negative result — the dataset has been scanned in full.\n"
            f"=== END AMOUNT-ONLY LOOKUP ==="
        )

    # Normalize date columns for display on exact-match set
    for c in ('Issued', 'CheckDate', 'Registration', 'CallDate'):
        if c in matches.columns:
            matches[c] = pd.to_datetime(matches[c], errors='coerce')

    def fmt_date(v):
        if pd.isna(v):
            return '(none)'
        try:
            return v.strftime('%m/%d/%Y')
        except Exception:
            return str(v)

    block = [f"\n=== AMOUNT-ONLY LOOKUP: principal = ${amount:,.2f} ==="]

    # Section 1: exact principal matches
    if not matches.empty:
        if 'Issued' in matches.columns:
            matches = matches.sort_values('Issued', ascending=False, na_position='last')
        total_count = len(matches)
        limit = min(max_rows_detail, total_count)
        head = matches.head(limit)
        block.append(
            f"\nExact-principal matches: {total_count} warrant(s) with principal = ${amount:,.2f}."
            + (f" Showing the {limit} most recent by issue date." if total_count > limit else "")
        )
        block.append(
            "Columns: County | SID | WarrantNo | Type | Payee | Issued | CheckDate | Expense"
        )
        for _, r in head.iterrows():
            wnum = int(r['Number']) if pd.notna(r.get('Number')) else '?'
            wtype = r.get('Type', '?') or '?'
            county = r.get('County', '?') or '?'
            sid = int(r['SID']) if pd.notna(r.get('SID')) else '?'
            payee = r.get('Payee') or '(unknown)'
            block.append(
                f"  {county} | SID {sid} | #{wnum} | {wtype} | {payee} | "
                f"Issued {fmt_date(r.get('Issued'))} | "
                f"CheckDate {fmt_date(r.get('CheckDate'))} | "
                f"{r.get('Expense') or '(unspecified)'}"
            )
    else:
        block.append(
            f"\nNo warrants in the dataset have a principal of exactly ${amount:,.2f}."
        )

    # Section 2: sum-pair groups (same payee+SID+date summing to amount)
    if not sum_groups.empty:
        block.append(
            f"\nSum-pair matches: {len(sum_groups)} group(s) of warrants that share "
            f"the same payee, SID, and issue date and whose principal values sum to "
            f"${amount:,.2f}. This pattern is common for fee-split payments (e.g. a "
            f"single fee allocated across a GF and a CF warrant on the same day)."
        )
        for _, g in sum_groups.iterrows():
            payee = g['Payee']
            county = g['County']
            sid = int(g['SID']) if pd.notna(g['SID']) else '?'
            issued_str = fmt_date(g['Issued'])
            block.append(
                f"\n  Group: {payee} / {county} SID {sid} / Issued {issued_str} — "
                f"{int(g['count'])} warrants summing to ${g['total']:,.2f}"
            )
            # Show each warrant in the group
            group_rows = df_grouped[
                (df_grouped['Payee'] == payee) &
                (df_grouped['County'] == county) &
                (df_grouped['SID'] == g['SID']) &
                (df_grouped['Issued'] == g['Issued'])
            ].copy()
            for c in ('CheckDate', 'Registration', 'CallDate'):
                if c in group_rows.columns:
                    group_rows[c] = pd.to_datetime(group_rows[c], errors='coerce')
            for _, r in group_rows.sort_values('Number').iterrows():
                wnum = int(r['Number']) if pd.notna(r.get('Number')) else '?'
                wtype = r.get('Type', '?') or '?'
                principal = r.get('Principal')
                pstr = f"${principal:,.2f}" if pd.notna(principal) else '$?'
                block.append(
                    f"    #{wnum} | {wtype} | {pstr} | "
                    f"CheckDate {fmt_date(r.get('CheckDate'))} | "
                    f"{r.get('Expense') or '(unspecified)'}"
                )

    block.append("=== END AMOUNT-ONLY LOOKUP ===")
    return '\n'.join(block)


def _find_payee_sid_rows(question: str, max_rows: int = 40) -> str:
    """
    Detect when a question is asking about row-level warrant detail for a specific
    payee within a specific SID (e.g. "last payment to OPPD in SID 366", "recent
    warrants to Kutak Rock in Douglas 630", "payments to MUD in 365"). If so,
    query the DataFrame directly and return a formatted context block to append
    for this request only. Returns empty string if no match.

    This is a workaround for questions that need row-level detail — the main
    context aggregates by payee×expense×year and drops individual warrant numbers.
    """
    import re as _re

    q = question.lower()

    # Must mention a SID number
    sid_match = _re.search(r'sid\s*#?\s*(\d{1,4})|\b(?:in|for)\s+(?:sarpy|douglas|cass|dodge|saunders)?\s*(\d{3,4})\b', q)
    if not sid_match:
        sid_match = _re.search(r'\bsid\s+(\d{1,4})\b', q)
    if not sid_match:
        return ''
    sid_num = next((g for g in sid_match.groups() if g), None)
    if not sid_num:
        return ''
    try:
        sid_num = int(sid_num)
    except ValueError:
        return ''

    # Use the cached DataFrame loaded at startup. Re-reading the 12MB Excel on
    # every helper call would add 3+ seconds of latency and contend with the
    # background context-warming thread.
    df = _get_warrants_df_cached()
    if df is None:
        return ''

    # Filter to the SID first
    df_sid = df[df['SID'] == sid_num]
    if df_sid.empty:
        return ''

    # For every payee in this SID, check whether their name appears in the question.
    payees_in_sid = df_sid['Payee'].dropna().unique().tolist()
    best_match = _detect_payee_from_question(q, payees_in_sid)

    if not best_match:
        return ''

    # Pull rows for this payee in this SID, sorted by Issued desc
    rows = df_sid[df_sid['Payee'] == best_match].copy()
    if rows.empty:
        return ''

    if 'Issued' in rows.columns:
        rows['Issued'] = pd.to_datetime(rows['Issued'], errors='coerce')
        rows = rows.sort_values('Issued', ascending=False, na_position='last')
    if 'CheckDate' in rows.columns:
        rows['CheckDate'] = pd.to_datetime(rows['CheckDate'], errors='coerce')

    # Limit to most recent max_rows
    rows = rows.head(max_rows)

    # County for the SID (should be unique)
    counties = df_sid['County'].dropna().unique()
    county = counties[0] if len(counties) > 0 else '?'

    # Format
    block = [
        f"\n=== ROW-LEVEL DETAIL: {best_match} in {county} SID {sid_num} ===",
        f"(Most recent {len(rows)} warrant(s), sorted by Issued date descending.)",
        "Columns: WarrantNo | Type | Issued | Principal | CheckDate | Expense | Registration | CallDate",
    ]
    for _, r in rows.iterrows():
        def fmt_date(v):
            if pd.isna(v):
                return '(none)'
            try:
                return v.strftime('%m/%d/%Y')
            except Exception:
                return str(v)
        wnum = int(r['Number']) if pd.notna(r.get('Number')) else '?'
        wtype = r.get('Type', '?') or '?'
        issued = fmt_date(r.get('Issued'))
        principal = r.get('Principal')
        pstr = f"${principal:,.2f}" if pd.notna(principal) else '$?'
        check = fmt_date(r.get('CheckDate'))
        expense = r.get('Expense') or '(unspecified)'
        reg = fmt_date(r.get('Registration'))
        called = fmt_date(r.get('CallDate'))
        block.append(
            f"  #{wnum} | {wtype} | Issued {issued} | {pstr} | CheckDate {check} | "
            f"{expense} | Reg {reg} | Call {called}"
        )
    block.append("=== END ROW-LEVEL DETAIL ===")
    return '\n'.join(block)


@app.route('/api/warrant-query', methods=['POST'])
@login_required
def warrant_query():
    """Answer a natural language question about the warrant dataset."""
    try:
        data = request.get_json(silent=True)
        question = (data or {}).get('question', '').strip()
        if not question:
            return jsonify({'error': 'No question provided'}), 400

        context = _get_warrant_context()

        # If the question is about a specific payee within a specific SID, look up
        # row-level warrant detail directly from the DataFrame. The main context
        # aggregates by payee×expense×year and drops individual warrant numbers /
        # dates, so row-level questions cannot be answered from it alone. We append
        # this detail to the UNCACHED portion of the user message so the cached
        # context prefix stays stable and cache hits are preserved.
        # Row-level lookups. Several question shapes need direct DataFrame queries:
        #   1. Payee + SID:     "last payment to OPPD in SID 366"
        #   2. Payee + amount:  "has there ever been a payment to Bluestem for $3,002.85"
        #   3. Specific date:   "what warrants were registered 04/20/2026"
        #   4. Holder:          "warrants held by David B. Henning"
        #   5. Amount only:     "are there any warrants for $1,735.75"
        # The main aggregate context cannot answer these — it groups by
        # payee×expense×year and strips individual warrant numbers, SIDs, dates,
        # and holder-level detail. We run all helpers and append whichever(s)
        # find matching data. These blocks go in the UNCACHED portion of the
        # user message so the cached context prefix stays stable.
        row_detail_sid = _find_payee_sid_rows(question)
        row_detail_amt = _find_payee_amount_rows(question)
        row_detail_date = _find_date_rows(question)
        row_detail_holder = _find_holder_rows(question)
        row_detail_amt_only = _find_amount_only_rows(question)
        row_detail = '\n'.join(
            block for block in (
                row_detail_sid, row_detail_amt,
                row_detail_date, row_detail_holder,
                row_detail_amt_only,
            ) if block
        )

        # Build a readable payee-alias list for the system prompt so the model knows
        # to treat short forms as equivalent to canonical payee names. Use the
        # canonical-names map for display (readable full names) rather than the
        # substring-matching map.
        from collections import defaultdict as _dd
        _canon_to_aliases = _dd(list)
        for short, canon in PAYEE_CANONICAL_NAMES.items():
            _canon_to_aliases[canon].append(short)
        alias_lines = []
        for canon, shorts in _canon_to_aliases.items():
            alias_lines.append(f"  {' / '.join(shorts)}  →  '{canon}'")
        alias_block = '\n'.join(alias_lines)

        system_prompt = (
            "You are a financial data analyst assistant for Bluestem Capital Partners, "
            "specializing in Nebraska Sanitary and Improvement District (SID) warrant data. "
            "You have been provided a summary of the warrant history dataset. "
            "Answer questions accurately, concisely, and professionally. "
            "Format dollar amounts with $ and commas. Be direct and specific. "
            "CRITICAL: Users often use short names or abbreviations for common payees. "
            "Treat these as EQUIVALENT to the canonical payee names — when you look up data "
            "in the context for 'OPPD' you must find it under 'Omaha Public Power District'. "
            "Canonical alias list:\n" + alias_block + "\n"
            "Never tell the user 'no data for OPPD' when the context has 'Omaha Public Power "
            "District'. Apply the alias lookup silently before reporting any result. "
            "CRITICAL: If the user's message contains a 'ROW-LEVEL DETAIL' block (appended when the "
            "question asks about a specific payee within a specific SID), use that row-level data "
            "to answer — it contains warrant numbers, issue dates, principal amounts, check dates, "
            "and expense categories. For a 'last payment' question, report the most recent "
            "Issued date's warrants (there may be multiple warrants issued on the same date) with "
            "warrant number, principal, and check date for each. If CheckDate is '(none)', note "
            "that no check date is recorded yet. "
            "CRITICAL: If the user's message contains a 'PAYEE+AMOUNT LOOKUP' block (appended when "
            "the question asks whether a specific payee ever received a specific dollar amount), "
            "use it as the authoritative answer. If the block shows matching warrants, list each "
            "with its County, SID, warrant number, Issued date, principal, CheckDate, and Expense. "
            "If the block states no matching warrants were found, answer directly that the amount "
            "has not been paid to that payee. Never fall back to 'I cannot confirm from the summary' "
            "when a PAYEE+AMOUNT LOOKUP block is present — it contains the definitive answer. "
            "CRITICAL: If the user's message contains a 'DATE LOOKUP' block (appended when the "
            "question asks about warrants associated with a specific date — registered, issued, "
            "check-dated, or called), use it as the authoritative answer. Summarize by County/SID "
            "and by Expense category, then list relevant row-level detail. If the count is small "
            "enough that the block includes full row-level detail, include warrant numbers and "
            "payees in your response. If the block shows no matches, answer directly that no "
            "warrants were recorded for that date. Never fall back to 'row-level data is not "
            "included' when a DATE LOOKUP block is present. "
            "CRITICAL: If the user's message contains a 'HOLDER LOOKUP' block (appended when the "
            "question asks what warrants a specific holder owns), use it as the authoritative "
            "answer. Report the P, I, and P+I totals, the per-SID breakdown, and the row-level "
            "detail when present. Default scope is outstanding (not yet redeemed); the block "
            "header will say '(outstanding)' or '(all)' depending on the question. Never respond "
            "with 'row-level detail is not available' when a HOLDER LOOKUP block is present. "
            "CRITICAL: If the user's message contains an 'AMOUNT-ONLY LOOKUP' block (appended when "
            "the question asks whether any warrant has a specific dollar amount), use it as the "
            "authoritative answer. The block has TWO sections: (1) exact-principal matches — "
            "individual warrants with that exact principal; (2) sum-pair matches — groups of "
            "warrants that share the same payee, SID, and issue date and whose principal values "
            "sum to the target amount (a common pattern for fee-split payments across GF and CF "
            "funds). When sum-pair matches exist, ALWAYS mention them in your answer — even if "
            "exact-principal matches are zero, a sum-pair match is a meaningful finding. For "
            "example, if the question asks about $1,735.75 and there are no exact matches but "
            "one sum-pair group shows two Bluestem warrants (#916 GF for $624.06 + #922 CF for "
            "$1,111.69 on the same date in SID 616) that total $1,735.75, report that as the "
            "answer. If both sections are empty, the block states this is an authoritative "
            "negative result — answer 'No, no warrants have a principal of exactly $X.XX and no "
            "fee-split pairs sum to that amount either.' NEVER fall back to 'I cannot confirm "
            "from the summary' when an AMOUNT-ONLY LOOKUP block is present. "
            "CRITICAL: When a question asks about 'outstanding' warrants by county (or which "
            "county has the most outstanding warrants), use the 'OUTSTANDING warrants by county' "
            "block — NOT the 'Principal by county (ALL warrants, historic + outstanding)' block. "
            "The all-warrants block totals every warrant ever issued including redeemed ones, "
            "while the outstanding block only counts warrants that are still held (Registration "
            "set, CallDate blank or in the future). These are very different numbers. If the "
            "question says 'outstanding', use the outstanding block; if it says 'total' or 'all "
            "time' or asks about history, use the all-warrants block. "
            "CRITICAL: When a question asks about SID activity — which SIDs are currently issuing "
            "warrants vs which ones have outstanding warrants but no recent registrations (e.g. "
            "'are there any SIDs that have outstanding warrants and have not registered warrants "
            "in 2026', 'which SIDs are inactive', 'which SIDs are still issuing') — use the 'SID "
            "activity status' block. It lists every INACTIVE SID with outstanding warrant count, "
            "principal, and last registration date. Report each inactive SID by name (County + "
            "SID number) along with its outstanding totals and last registration date. Do NOT "
            "respond with 'row-level activity data is not available' when the SID activity status "
            "block is present. "
            "CRITICAL: For filter-style questions that combine outstanding totals with current-year "
            "registration activity per SID (e.g. 'which SIDs have over $5M outstanding and "
            "registered warrants in 2026', 'top SIDs by outstanding principal', 'which SIDs "
            "registered the most in the current year') — use the 'Per-SID outstanding + current-"
            "year registration activity' table inside the SID activity status block. Each line "
            "lists County SID, outstanding warrant count and principal, plus current-year "
            "registration count and principal. Filter and sort the listed SIDs by whatever the "
            "question asks (e.g. for 'over $5M outstanding registered in 2026', take the rows "
            "where outstanding > $5M AND current-year registrations > 0). The listing is sorted "
            "by outstanding principal descending. Do NOT respond with 'per-SID breakdown not "
            "available' when this table is present. "
            "CRITICAL: When a question specifies an expense category (e.g. 'Financial Advisory Fees', "
            "'Warrant Structuring Fees', 'Bond Counsel', 'Paving'), you MUST filter on BOTH the "
            "payee AND the expense category — do not sum across expense categories. "
            "A single payee (e.g. Bluestem Capital Partners) often receives warrants in multiple "
            "distinct expense categories; reporting a payee's total across all categories when the "
            "user asked about a specific category is a significant error. Use the "
            "'Warrant issuances by payee, expense category, and year' cross-tab for these questions. "
            "CRITICAL: When a question asks about calls for a holder (past calls, future/scheduled "
            "calls, or calls on a specific date), ALWAYS include BOTH the principal and the interest "
            "amounts, and state the P+I total. Never report principal alone for call questions — "
            "interest is material to the holder and is always part of a call answer. "
            "ALSO always include the SID numbers involved in the call. For future holder calls, the "
            "'Future-called warrants by holder / call date / county / SID' block gives you the exact "
            "SID-level breakdown per holder. List the county and SID numbers (e.g. 'across Sarpy "
            "SIDs 311, 343, 344, 346, 353, 359, 362, 363, 377'). For past calls, SID-level breakdown "
            "is not pre-aggregated in the context — note this and offer totals at the level available. "
            "CRITICAL: Lead with the direct answer. Before writing any response, verify that the "
            "data you're about to report actually exists in the context. Do not write contradictory "
            "prose such as 'there is no call on X date' followed by data showing a call on X date. "
            "If the requested data IS in the context, state the answer confidently and directly. "
            "If it is NOT in the context, say so clearly without also listing conflicting data. "
            "CRITICAL: When asked about interest payment dates for a month (e.g. 'what SIDs pay "
            "interest in May'), use the 'Annual interest payment schedule' block. Annual interest "
            "ONLY applies to CF (Construction Fund) warrants — GF (General Fund) warrants do NOT "
            "pay annual interest. The schedule is already filtered to SIDs with currently "
            "outstanding CF warrants; report the numbers as-is. Include ALL SIDs listed for the "
            "requested month, grouped by county, with the specific payment date (MM/DD), "
            "outstanding CF principal (P), and accrued interest (I) for each SID. "
            "If a user asks about interest for a SID that is NOT in the schedule, explain that "
            "the SID has no outstanding CF warrants (either it's GF-only, or all CF warrants have "
            "been redeemed) and therefore does not currently pay annual interest. Never invent an "
            "interest-payment date for a SID not in the schedule. "
            "CRITICAL: Date formats are EQUIVALENT regardless of zero-padding or separator. "
            "'5/12/26', '5-12-26', '5/12/2026', '05/12/26', '5/12', and '05/12/2026' ALL refer "
            "to the same date. Before claiming a date does not exist in the dataset, normalize "
            "the user's date format to match the context's MM/DD/YYYY format. Never say 'no call "
            "on 5/12/2026' when the context contains '05/12/2026' — those are the same date."
        )

        answer = _anthropic_request({
            'model': 'claude-sonnet-4-6',
            'max_tokens': 8192,
            'system': system_prompt,
            'messages': [
                {
                    'role': 'user',
                    'content': [
                        # Cache the large static context block. First request writes the cache
                        # (~5s overhead), subsequent requests within 5 min read at 10% input cost
                        # and are substantially faster (typically 30-60% latency reduction on
                        # 100K-token prefixes).
                        {
                            'type': 'text',
                            'text': context,
                            'cache_control': {'type': 'ephemeral'},
                        },
                        {
                            'type': 'text',
                            'text': (
                                f"{row_detail}\n\nUser question: {question}"
                                if row_detail
                                else f"\n\nUser question: {question}"
                            ),
                        },
                    ],
                }
            ]
        })

        return jsonify({'answer': answer})

    except Exception as e:
        return jsonify({'error': f'Query failed: {str(e)}'}), 500


@app.route('/api/warrant-date')
@login_required
def warrant_date():
    """Return the as-of date from the warrant history file."""
    _, file_date = _load_warrants_df()
    return jsonify({'date': file_date})


@app.route('/api/health')
@login_required
def health():
    """Diagnostic endpoint — checks all dependencies and template files."""
    status = {}
    # Check imports
    for lib in ['pikepdf', 'pandas', 'scipy', 'dateutil']:
        try:
            __import__(lib)
            status[lib] = 'ok'
        except ImportError as e:
            status[lib] = f'MISSING: {e}'
    # Check template files
    for fname in ['blank_8038G_irs.pdf', 'blank_8038GC_irs.pdf', '8038_Info.xlsx']:
        path = os.path.join(os.path.dirname(__file__), fname)
        status[fname] = f'{os.path.getsize(path)} bytes' if os.path.exists(path) else 'MISSING'
    # Check 8038_Info loaded
    status['8038_info_rows'] = len(INFO_8038)
    return jsonify(status)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
