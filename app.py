from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import smtplib, os, io
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime

app = Flask(__name__)
CORS(app)

SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', 587))
SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASS = os.environ.get('SMTP_PASS', '')
FIRM_EMAIL = os.environ.get('FIRM_EMAIL', 'info@tflaw.co.uk')

def tick(val):
    return 'X' if val else ' '

def v(d, key):
    return str(d.get(key, '') or '')

def make_aa(d):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)

    def hdr(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11)
        p.paragraph_format.space_after = Pt(2)
        # Dark background effect via shading
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1F3864')
        pPr.append(shd)
        run.font.color.rgb = None
        from docx.shared import RGBColor
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def field(label, value, bold_val=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(label + '  ')
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = p.add_run(value)
        r2.bold = bold_val
        r2.font.size = Pt(10)
        from docx.shared import RGBColor
        r2.font.color.rgb = RGBColor(0x00, 0x00, 0x80)

    def field2(l1, v1, l2, v2):
        from docx.oxml.ns import qn
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        table.autofit = False
        widths = [2700, 3600, 2700, 3600]
        from docx.shared import Twips
        for i, cell in enumerate(table.rows[0].cells):
            cell.width = Twips(widths[i])
        cells = table.rows[0].cells
        cells[0].paragraphs[0].add_run(l1).bold = True
        cells[1].paragraphs[0].add_run(v1)
        cells[2].paragraphs[0].add_run(l2).bold = True
        cells[3].paragraphs[0].add_run(v2)
        for cell in cells:
            for para in cell.paragraphs:
                para.paragraph_format.space_before = Pt(2)
                para.paragraph_format.space_after = Pt(2)
        doc.add_paragraph()

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run('Civil Advice and Assistance and Civil ABWOR Legal Aid Online Declaration')
    tr.bold = True
    tr.font.size = Pt(13)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run('AA/LAO/CIV  |  Thompson Family Law Solicitors  |  Subject: ' + v(d,'subject')).font.size = Pt(10)
    doc.add_paragraph()

    hdr('A.  Applicant Details')
    field2('Forename:', v(d,'fname'), 'Surname:', v(d,'lname'))
    field2('Date of Birth:', v(d,'dob'), 'NI Number:', v(d,'ni') or 'Not provided')
    field2('Telephone:', v(d,'tel'), 'Email:', v(d,'email'))
    field2('Contact by Email:', v(d,'cemail'), 'Previous Solicitor:', v(d,'prevsol'))
    if v(d,'prevsol') == 'Yes':
        field('Previous solicitor details:', v(d,'prevsol_det'))
    field('Home Address:', v(d,'haddr') + ', ' + v(d,'hpc'))
    field('Communication Support:', v(d,'comm') or 'None/no support needed')
    doc.add_paragraph()

    hdr('Applicant/Client Equality Information')
    field2('Sex:', v(d,'sex'), 'Care Experience:', v(d,'care'))
    field('Disability:', v(d,'disabilities'))
    field('National Identity:', v(d,'natid'))
    field('Ethnic Group:', v(d,'ethnic'))
    doc.add_paragraph()

    hdr('B.  Applicant Assistance')
    field2('Other assistance with legal costs:', v(d,'otherassist'), '', '')
    if v(d,'otherassist') == 'Yes':
        field('Why cannot be used:', v(d,'oa_why'))
    doc.add_paragraph()

    hdr('C.  Financial Details')
    field2('Spouse/Partner:', v(d,'partner'), 'Contrary Interest:', v(d,'contrary'))
    if v(d,'partner') == 'Yes':
        field2('Partner Name:', v(d,'pfname')+' '+v(d,'plname'), 'Partner DOB/NI:', v(d,'pdob')+' / '+v(d,'pni'))
    field2('Dependants with client:', v(d,'dep1'), 'Dependants not with client:', v(d,'dep2'))
    field2('Passported Benefits - Client:', v(d,'pass_c'), 'Partner:', v(d,'pass_p'))
    doc.add_paragraph()

    hdr('D.  Bank Accounts and Capital')
    banks = [
        (v(d,'ba1n'), v(d,'ba1num'), v(d,'ba1t'), v(d,'ba1b')),
        (v(d,'ba2n'), v(d,'ba2num'), v(d,'ba2t'), v(d,'ba2b')),
        (v(d,'ba3n'), v(d,'ba3num'), v(d,'ba3t'), v(d,'ba3b')),
    ]
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    from docx.shared import Twips, RGBColor
    hdrs = ['Bank', 'Acct No (last 4)', 'Type', 'Balance']
    for i, cell in enumerate(table.rows[0].cells):
        r = cell.paragraphs[0].add_run(hdrs[i])
        r.bold = True
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1F3864')
        tcPr.append(shd)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for name, num, typ, bal in banks:
        if name or num:
            row = table.add_row()
            row.cells[0].paragraphs[0].add_run(name)
            row.cells[1].paragraphs[0].add_run(num)
            row.cells[2].paragraphs[0].add_run(typ)
            row.cells[3].paragraphs[0].add_run(('PS' + bal) if bal else '')
    doc.add_paragraph()

    hdr('E.  Income Details')
    field('Non-passport Benefits:', v(d,'nonpass'))
    field2('Pay/Sick Pay (weekly net):', 'PS'+v(d,'earn_pay') if v(d,'earn_pay') else '--',
           'Self-employed drawings:', 'PS'+v(d,'earn_se') if v(d,'earn_se') else '--')
    field('Documentary Evidence Provided:', v(d,'evidence'))
    doc.add_paragraph()

    hdr('Applicant Declaration and Signature')
    decl_items = [
        'This is a true statement of my personal and financial circumstances.',
        'I understand that if I give false information to the Scottish Legal Aid Board (SLAB), I may be prosecuted.',
        'I understand that SLAB can make any enquiries and get any information it needs to deal with this application.',
        'I agree to SLAB obtaining and/or checking information with others such as my employer, banks, credit reference agencies, the Department for Work and Pensions and HM Revenue and Customs.',
        'I agree to the disclosure of the application, associated documentation and my case file held by my solicitor, to SLAB for audit and/or quality assurance.',
        'SLAB may use the information provided for the prevention and detection of fraud.',
        'I agree that all consents will be effective for a period of not less than five years from the date of signature.',
    ]
    for i, item in enumerate(decl_items, 1):
        p = doc.add_paragraph(style='List Number')
        p.add_run(item).font.size = Pt(9)

    doc.add_paragraph()
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.rows[0].cells[0].paragraphs[0].add_run('Signature of Applicant/Representative:').bold = True
    sig_table.rows[0].cells[1].paragraphs[0].add_run('Date:').bold = True
    sig_table.rows[1].cells[0].paragraphs[0].add_run(v(d,'signame')).font.italic = True
    sig_table.rows[1].cells[1].paragraphs[0].add_run(v(d,'sigdate'))
    doc.add_paragraph()

    hdr('Solicitor Declaration')
    sol_items = [
        'I consent to the disclosure of the application, associated documentation and client case file for quality assurance including audit and peer review, at any stage.',
        'I accept responsibility for any act or omission in relation to the completion and submission of the application on Legal Aid Online (LAOL) by me or on my behalf.',
        'I have satisfied myself that my client qualifies financially for advice and assistance.',
        'I will retain this signed, completed document in paper form or electronically and will send it to SLAB upon request.',
    ]
    for item in sol_items:
        p = doc.add_paragraph(style='List Number')
        p.add_run(item).font.size = Pt(9)

    doc.add_paragraph()
    sol_table = doc.add_table(rows=2, cols=2)
    sol_table.style = 'Table Grid'
    sol_table.rows[0].cells[0].paragraphs[0].add_run('Signature of Solicitor:').bold = True
    sol_table.rows[0].cells[1].paragraphs[0].add_run('Date:').bold = True
    sol_table.rows[1].cells[0].paragraphs[0].add_run('Michael Thompson').font.italic = True
    sol_table.rows[1].cells[1].paragraphs[0].add_run(v(d,'sigdate'))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def make_civ(d):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)

    def hdr(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11)
        p.paragraph_format.space_after = Pt(2)
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1F3864')
        pPr.append(shd)
        from docx.shared import RGBColor
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def field(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(label + '  ')
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = p.add_run(value)
        r2.font.size = Pt(10)
        from docx.shared import RGBColor
        r2.font.color.rgb = RGBColor(0x00, 0x00, 0x80)

    def field2(l1, v1, l2, v2):
        from docx.shared import Twips
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        widths = [2700, 3600, 2700, 3600]
        for i, cell in enumerate(table.rows[0].cells):
            cell.width = Twips(widths[i])
        cells = table.rows[0].cells
        cells[0].paragraphs[0].add_run(l1).bold = True
        cells[1].paragraphs[0].add_run(v1)
        cells[2].paragraphs[0].add_run(l2).bold = True
        cells[3].paragraphs[0].add_run(v2)
        for cell in cells:
            for para in cell.paragraphs:
                para.paragraph_format.space_before = Pt(2)
                para.paragraph_format.space_after = Pt(2)
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run('Civil Legal Aid - Legal Aid Online Declaration')
    tr.bold = True
    tr.font.size = Pt(13)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run('CIV/SOL  |  Thompson Family Law Solicitors  |  Subject: ' + v(d,'subject')).font.size = Pt(10)
    doc.add_paragraph()

    hdr('A.  Applicant Details')
    field2('Forename:', v(d,'fname'), 'Surname:', v(d,'lname'))
    field2('Date of Birth:', v(d,'dob'), 'NI Number:', v(d,'ni') or 'Not provided')
    field2('Telephone:', v(d,'tel'), 'Email:', v(d,'email'))
    field2('Contact by Email:', v(d,'cemail'), '', '')
    field('Home Address:', v(d,'haddr') + ', ' + v(d,'hpc'))
    field('Communication Support:', v(d,'comm') or 'None/no support needed')
    doc.add_paragraph()

    hdr('Equality Information')
    field2('Sex:', v(d,'sex'), 'Care Experience:', v(d,'care'))
    field('Disability:', v(d,'disabilities'))
    field('National Identity:', v(d,'natid'))
    field('Ethnic Group:', v(d,'ethnic'))
    doc.add_paragraph()

    hdr('B.  Other Rights and Resources')
    field2('Other assistance with legal costs:', v(d,'otherassist'), '', '')
    doc.add_paragraph()

    hdr('C.  Financial Details - Passported Benefits')
    field2('Passported Benefits - Client:', v(d,'pass_c'), 'Partner:', v(d,'pass_p'))
    doc.add_paragraph()

    hdr('Spouse or Partner Details')
    field2('Has Spouse/Partner:', v(d,'partner'), 'Contrary Interest:', v(d,'contrary'))
    if v(d,'partner') == 'Yes':
        field2('Partner Name:', v(d,'pfname')+' '+v(d,'plname'), 'Partner DOB/NI:', v(d,'pdob')+' / '+v(d,'pni'))
    doc.add_paragraph()

    hdr('Dependants')
    field2('Dependants with client:', v(d,'dep1'), 'Dependants not with client:', v(d,'dep2'))
    doc.add_paragraph()

    hdr('D.  Capital and Other Assets')
    banks = [
        (v(d,'ba1n'), v(d,'ba1num'), v(d,'ba1t'), v(d,'ba1b')),
        (v(d,'ba2n'), v(d,'ba2num'), v(d,'ba2t'), v(d,'ba2b')),
        (v(d,'ba3n'), v(d,'ba3num'), v(d,'ba3t'), v(d,'ba3b')),
    ]
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    from docx.shared import RGBColor
    hdrs = ['Bank', 'Acct No (last 4)', 'Type', 'Balance']
    for i, cell in enumerate(table.rows[0].cells):
        r = cell.paragraphs[0].add_run(hdrs[i])
        r.bold = True
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1F3864')
        tcPr.append(shd)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for name, num, typ, bal in banks:
        if name or num:
            row = table.add_row()
            row.cells[0].paragraphs[0].add_run(name)
            row.cells[1].paragraphs[0].add_run(num)
            row.cells[2].paragraphs[0].add_run(typ)
            row.cells[3].paragraphs[0].add_run(('PS' + bal) if bal else '')
    doc.add_paragraph()

    hdr('E.  Income Details')
    field('Non-passport Benefits:', v(d,'nonpass'))
    field2('Pay/Sick Pay (weekly net):', 'PS'+v(d,'earn_pay') if v(d,'earn_pay') else '--',
           'Self-employed drawings:', 'PS'+v(d,'earn_se') if v(d,'earn_se') else '--')
    field('Documentary Evidence Provided:', v(d,'evidence'))
    doc.add_paragraph()

    hdr('Applicant Declaration and Signature')
    decl_items = [
        'This is a true statement of my personal and financial circumstances.',
        'I understand that if I give false information to the Scottish Legal Aid Board (SLAB), I may be prosecuted.',
        'I understand that SLAB can make any enquiries and get any information it needs to deal with this application.',
        'I agree to SLAB checking information with my employer, banks, DWP and HMRC.',
        'I must tell my solicitor immediately if there are any changes in my or my partner financial circumstances including a change in benefits.',
        'If my solicitor does special urgency work for me I know that SLAB may require me to pay a contribution.',
        'I agree to the disclosure of the application and case file to SLAB for audit and/or quality assurance.',
        'SLAB may share information for fraud prevention. I consent to disclosure of my personal data.',
        'All consents are effective for not less than five years from the date of signature.',
    ]
    for i, item in enumerate(decl_items, 1):
        p = doc.add_paragraph(style='List Number')
        p.add_run(item).font.size = Pt(9)

    doc.add_paragraph()
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.rows[0].cells[0].paragraphs[0].add_run('Signature of Applicant/Representative:').bold = True
    sig_table.rows[0].cells[1].paragraphs[0].add_run('Date:').bold = True
    sig_table.rows[1].cells[0].paragraphs[0].add_run(v(d,'signame')).font.italic = True
    sig_table.rows[1].cells[1].paragraphs[0].add_run(v(d,'sigdate'))
    doc.add_paragraph()

    hdr('Solicitor Declaration')
    sol_items = [
        'I consent to the disclosure of the application, associated documentation and client case file for quality assurance including audit and peer review, at any stage.',
        'I accept responsibility for any act or omission in relation to the completion and submission of the application on Legal Aid Online (LAOL) by me or on my behalf.',
        'I will retain this signed, completed document in paper form or electronically and will send it to SLAB upon request.',
    ]
    for item in sol_items:
        p = doc.add_paragraph(style='List Number')
        p.add_run(item).font.size = Pt(9)

    doc.add_paragraph()
    sol_table = doc.add_table(rows=2, cols=2)
    sol_table.style = 'Table Grid'
    sol_table.rows[0].cells[0].paragraphs[0].add_run('Signature of Solicitor:').bold = True
    sol_table.rows[0].cells[1].paragraphs[0].add_run('Date:').bold = True
    sol_table.rows[1].cells[0].paragraphs[0].add_run('Michael Thompson').font.italic = True
    sol_table.rows[1].cells[1].paragraphs[0].add_run(v(d,'sigdate'))

    hdr('Partner Declaration')
    p = doc.add_paragraph()
    p.add_run('Partner name:  ').bold = True
    p.add_run(v(d,'pfname') + ' ' + v(d,'plname'))
    doc.add_paragraph('I have seen the financial information given in this declaration. It is a true statement of my personal and financial circumstances.')

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def send_email(d, aa_buf, civ_buf):
    lname = v(d, 'lname')
    fname = v(d, 'fname')
    ref = v(d, 'ref')

    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = FIRM_EMAIL
    msg['Subject'] = f'New Legal Aid Declaration - {fname} {lname} - {ref}'

    body = f"""New SLAB declaration submitted via online form.

Client: {fname} {lname}
DOB: {v(d,'dob')}
NI: {v(d,'ni') or 'Not provided'}
Tel: {v(d,'tel')}
Email: {v(d,'email')}
Address: {v(d,'haddr')}, {v(d,'hpc')}
Subject Matter: {v(d,'subject')}
Passported Benefits: {v(d,'pass_c')}
Evidence Provided: {v(d,'evidence')}
Date Signed: {v(d,'sigdate')}
Reference: {ref}

Both completed declaration documents are attached.
Please countersign and submit to SLAB via Legal Aid Online.
"""
    msg.attach(MIMEText(body, 'plain'))

    for buf, filename in [
        (aa_buf, f'AA_LAO_CIV_{lname}_{fname}.docx'),
        (civ_buf, f'CIV_SOL_{lname}_{fname}.docx'),
    ]:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(buf.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)

@app.route('/submit', methods=['POST'])
def submit():
    try:
        d = request.json
        d['ref'] = 'TFL-' + str(int(datetime.now().timestamp()))[-6:]
        aa_buf = make_aa(d)
        civ_buf = make_civ(d)
        send_email(d, aa_buf, civ_buf)
        return jsonify({'ok': True, 'ref': d['ref']})
    except Exception as e:
        print('Error:', e)
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/')
def home():
    return 'SLAB Forms Server - OK'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
