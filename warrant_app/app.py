import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import date, datetime
from flask import Flask, request, jsonify, send_file, render_template, session, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import openpyxl
from functools import wraps

# Load API key from .env file if present
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
except ImportError:
    pass

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-change-in-production')

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
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode('utf-8'))
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

Find ANY of these issues:

1. AMOUNT TYPOS: Extra or missing digits in any warrant amount (e.g. "$6,193.119" instead of "$6,193.19").

2. MISSING PAYEES: Any warrant with a blank or missing payee name.

3. PAYEE INCONSISTENCIES: Payee name differs between the consolidated narrative section and the itemized warrant section.

4. GROUPING RULE - apply before flagging amount mismatches:
   The consolidated section often shows a SINGLE TOTAL for a group of warrants to the same payee.
   Always SUM all extracted warrants for that payee and compare the SUM to the consolidated figure.
   Never compare a single warrant amount to a group total. Only flag if the SUM does not reconcile.

5. FEE VERIFICATION - this is critical and must always be checked:
   Advisory fees (typically Bluestem Capital Partners) and underwriting/placement fees (typically Northland Securities, Ameritas Investment, SouthState|DuncanWilliams, or Access Bank) are calculated as a percentage of the warrants issued.
   The PDF states the percentage and base explicitly, e.g. "advisory fees (2% of $14,517.11)" or "underwriting fees (2% of $24,894.28)".

   CRITICAL DISTINCTION — Bluestem and other payees sometimes appear on multiple warrants in the same section for DIFFERENT purposes:
   - "Advisory fees on warrants issued at this meeting (X%)" — this IS the percentage-based fee; verify the math
   - "Financial Advisor / Fiscal Agent services for FY 20XX/XX" or any flat fee with an invoice number (e.g. "#3501") — this is a flat retainer fee, NOT a percentage-based fee; do NOT include it in the fee base and do NOT verify it as a percentage calculation
   - Annual paying agent / registrar fees (e.g. SID Services LLC) — also flat fees, exclude from base
   When Bluestem appears on multiple warrants, identify which one is the advisory fee (it will explicitly state a percentage) and only verify that one.

   For each percentage-based fee warrant, you must:
   a) Find the stated percentage and base amount in the PDF text
   b) Calculate: stated_percentage × stated_base = expected_fee
   c) Compare expected_fee to the extracted warrant amount
   d) Also verify the stated base itself is correct by summing the applicable preceding warrants:
      - Exclude from the base: the advisory fee warrant itself, the underwriting fee warrant, any flat retainer/fiscal agent fees, any annual paying agent fees
      - The underwriting fee base = same warrants as advisory base PLUS the advisory fee warrant
   e) Flag ONLY if the discrepancy is greater than $0.05. If the amounts match within $0.05, do NOT flag.

   Example: If Bluestem's fee is stated as "2% of $626,541.68" but the eligible preceding warrants sum to $626,541.68, and 2% = $12,530.83, but the warrant shows $18,364.32 — flag it.
   Also flag if the stated base doesn't match the actual sum of eligible preceding warrants.

IMPORTANT: Never flag a match. Apply this test before adding ANY flag: calculate the expected value yourself, compare it to the extracted value, and only flag if they differ by more than $0.05. If the numbers are equal or within $0.05, do NOT generate a flag — not even to note that they match. Silence means everything is correct. If your flag message would contain phrases like "amounts match", "totals match", "this should be based on", or show the same dollar figure on both sides of a comparison, discard that flag entirely.

Return ONLY a JSON array of issue objects:
{
  "severity": "error" or "warning",
  "warrant": "<warrant number(s) or 'General' or 'Construction' or 'N/A'>",
  "field": "amount" or "payee" or "total" or "fee",
  "message": "<plain-English description including the expected value>"
}

Empty array [] if no issues. No markdown, no explanation."""

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
    return result if isinstance(result, list) else []

# ── Excel: Sarpy County ───────────────────────────────────────────────────────
def build_excel_sarpy(extracted: dict, dt_wt: date, dt_reg: date) -> bytes:
    sid_num  = extracted['sid_number']
    sid_name = SID_NAMES.get(sid_num, '')
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
            ws[f'A{r}'].value = f'=IF(C{r}="","",0.07)'
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
            ws[f'M{r}'].value = 7

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

    sid_num  = extracted.get('sid_number', 'XXX')
    filename = f'SID_{sid_num}_Warrants_to_Register_{dt_reg.strftime("%m-%d-%Y")}.xlsx'

    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
