from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
from docx.shared import Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, io, base64
from datetime import datetime
import sendgrid
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

app = Flask(__name__)
CORS(app)

SENDGRID_API_KEY = os.environ.get('SENDGRID_API_KEY', '')
FIRM_EMAIL = os.environ.get('FIRM_EMAIL', 'info@tflaw.co.uk')
FROM_EMAIL = os.environ.get('FROM_EMAIL', 'noreply@tflaw.co.uk')

def v(d, key):
    return str(d.get(key, '') or '')

def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def shade_para(para, hex_color):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)

def add_hdr(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    shade_para(p, '1F3864')
    run = p.add_run('  ' + text)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

def add_field(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(label + '  ')
    r1.bold = True
    r1.font.size = Pt(9)
    r2 = p.add_run(value or '--')
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(0x00, 0x00, 0x80)

def add_field2(doc, l1, v1, l2, v2):
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    widths = [2500, 3800, 2500, 3800]
    for i, cell in enumerate(table.rows[0].cells):
        cell.width = Twips(widths[i])
        cell.paragraphs[0].paragraph_format.space_before = Pt(2)
        cell.paragraphs[0].paragraph_format.space_after = Pt(2)
    cells = table.rows[0].cells
    r = cells[0].paragraphs[0].add_run(l1)
    r.bold = True
    r.font.size = Pt(9)
    cells[1].paragraphs[0].add_run(v1 or '--').font.size = Pt(9)
    if l2:
        r2 = cells[2].paragraphs[0].add_run(l2)
        r2.bold = True
        r2.font.size = Pt(9)
    cells[3].paragraphs[0].add_run(v2 or '--').font.size = Pt(9)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)

def add_decl_items(doc, items):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph(style='List Number')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.add_run(item).font.size = Pt(8.5)

def add_sig_table(doc, signer_label, sig_name, date):
    doc.add_paragraph()
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'
    for cell in table.rows[0].cells:
        shade_cell(cell, 'E8E8F0')
    cells0 = table.rows[0].cells
    r = cells0[0].paragraphs[0].add_run(signer_label)
    r.bold = True
    r.font.size = Pt(9)
    r2 = cells0[1].paragraphs[0].add_run('Date:')
    r2.bold = True
    r2.font.size = Pt(9)
    cells1 = table.rows[1].cells
    sig_run = cells1[0].paragraphs[0].add_run(sig_name)
    sig_run.font.size = Pt(18)
    sig_run.font.italic = True
    sig_run.font.color.rgb = RGBColor(0x0D, 0x1B, 0x2A)
    cells1[1].paragraphs[0].add_run(date).font.size = Pt(9)
    for cell in table.rows[2].cells:
        cell.paragraphs[0].add_run('').font.size = Pt(6)
    doc.add_paragraph()

def make_aa(d):
    doc = Document()
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run('Civil Advice and Assistance / Civil ABWOR Legal Aid Online Declaration')
    tr.bold = True
    tr.font.size = Pt(13)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run('AA/LAO/CIV  |  Thompson Family Law  |  Subject: ' + v(d,'subject')).font.size = Pt(9)
    add_hdr(doc, 'A.  Applicant Details')
    add_field2(doc, 'Forename:', v(d,'fname'), 'Surname:', v(d,'lname'))
    add_field2(doc, 'Date of Birth:', v(d,'dob'), 'NI Number:', v(d,'ni') or 'Not provided')
    add_field2(doc, 'Telephone:', v(d,'tel'), 'Email:', v(d,'email'))
    add_field2(doc, 'Contact by Email:', v(d,'cemail'), 'Previous Solicitor:', v(d,'prevsol'))
    if v(d,'prevsol') == 'Yes':
        add_field(doc, 'Previous solicitor details:', v(d,'prevsol_det'))
    add_field(doc, 'Home Address:', v(d,'haddr') + ', ' + v(d,'hpc'))
    add_hdr(doc, 'Equality Information')
    add_field2(doc, 'Sex:', v(d,'sex'), 'Care Experience:', v(d,'care'))
    add_field(doc, 'Disability:', v(d,'disabilities'))
    add_field2(doc, 'National Identity:', v(d,'natid'), 'Ethnic Group:', v(d,'ethnic'))
    add_hdr(doc, 'B.  Applicant Assistance')
    add_field2(doc, 'Other assistance with legal costs:', v(d,'otherassist'), '', '')
    if v(d,'otherassist') == 'Yes':
        add_field(doc, 'Why cannot be used:', v(d,'oa_why'))
    add_hdr(doc, 'C.  Financial Details')
    add_field2(doc, 'Spouse/Partner:', v(d,'partner'), 'Contrary Interest:', v(d,'contrary'))
    if v(d,'partner') == 'Yes':
        add_field2(doc, 'Partner Name:', v(d,'pfname')+' '+v(d,'plname'), 'Partner DOB/NI:', v(d,'pdob')+' / '+v(d,'pni'))
    add_field2(doc, 'Dependants with client:', v(d,'dep1'), 'Dependants not with client:', v(d,'dep2'))
    add_field2(doc, 'Passported Benefits - Client:', v(d,'pass_c'), 'Partner:', v(d,'pass_p'))
    add_hdr(doc, 'D.  Bank Accounts and Capital')
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    for cell, hdr_text in zip(table.rows[0].cells, ['Bank', 'Acct No (last 4)', 'Type', 'Balance']):
        shade_cell(cell, '1F3864')
        r = cell.paragraphs[0].add_run(hdr_text)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for name, num, typ, bal in [(v(d,'ba1n'),v(d,'ba1num'),v(d,'ba1t'),v(d,'ba1b')),(v(d,'ba2n'),v(d,'ba2num'),v(d,'ba2t'),v(d,'ba2b')),(v(d,'ba3n'),v(d,'ba3num'),v(d,'ba3t'),v(d,'ba3b'))]:
        if name or num:
            row = table.add_row()
            for cell, val in zip(row.cells, [name, num, typ, ('£'+bal) if bal else '']):
                cell.paragraphs[0].add_run(val).font.size = Pt(9)
    doc.add_paragraph()
    add_hdr(doc, 'E.  Income Details')
    add_field(doc, 'Non-passport Benefits:', v(d,'nonpass'))
    add_field2(doc, 'Pay/Sick Pay (weekly net):', ('£'+v(d,'earn_pay')) if v(d,'earn_pay') else '--', 'Self-employed drawings:', ('£'+v(d,'earn_se')) if v(d,'earn_se') else '--')
    add_field(doc, 'Documentary Evidence Provided:', v(d,'evidence'))
    add_hdr(doc, 'Applicant Declaration')
    add_decl_items(doc, [
        'This is a true statement of my personal and financial circumstances.',
        'I understand that if I give false information to the Scottish Legal Aid Board (SLAB), I may be prosecuted.',
        'I understand that SLAB can make any enquiries and get any information it needs to deal with this application.',
        'I agree to SLAB obtaining and/or checking information with my employer, banks, DWP and HMRC.',
        'I agree to the disclosure of the application and case file held by my solicitor to SLAB for audit and quality assurance.',
        'SLAB may use the information provided for the prevention and detection of fraud.',
        'I agree that all consents will be effective for not less than five years from the date of signature.',
    ])
    add_sig_table(doc, 'Signature of Applicant / Representative:', v(d,'signame'), v(d,'sigdate'))
    add_hdr(doc, 'Solicitor Declaration')
    add_decl_items(doc, [
        'I consent to the disclosure of the application and client case file for quality assurance including audit and peer review.',
        'I accept responsibility for the completion and submission of the application on Legal Aid Online (LAOL).',
        'I have satisfied myself that my client qualifies financially for advice and assistance.',
        'I will retain this signed document in paper form or electronically and will send it to SLAB upon request.',
    ])
    add_sig_table(doc, 'Signature of Solicitor:', 'Michael Thompson', v(d,'sigdate'))
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def make_civ(d):
    doc = Document()
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run('Civil Legal Aid - Legal Aid Online Declaration')
    tr.bold = True
    tr.font.size = Pt(13)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run('CIV/SOL  |  Thompson Family Law  |  Subject: ' + v(d,'subject')).font.size = Pt(9)
    add_hdr(doc, 'A.  Applicant Details')
    add_field2(doc, 'Forename:', v(d,'fname'), 'Surname:', v(d,'lname'))
    add_field2(doc, 'Date of Birth:', v(d,'dob'), 'NI Number:', v(d,'ni') or 'Not provided')
    add_field2(doc, 'Telephone:', v(d,'tel'), 'Email:', v(d,'email'))
    add_field2(doc, 'Contact by Email:', v(d,'cemail'), '', '')
    add_field(doc, 'Home Address:', v(d,'haddr') + ', ' + v(d,'hpc'))
    add_hdr(doc, 'Equality Information')
    add_field2(doc, 'Sex:', v(d,'sex'), 'Care Experience:', v(d,'care'))
    add_field(doc, 'Disability:', v(d,'disabilities'))
    add_field2(doc, 'National Identity:', v(d,'natid'), 'Ethnic Group:', v(d,'ethnic'))
    add_hdr(doc, 'B.  Other Rights and Resources')
    add_field2(doc, 'Other assistance with legal costs:', v(d,'otherassist'), '', '')
    add_hdr(doc, 'C.  Financial Details - Passported Benefits')
    add_field2(doc, 'Passported Benefits - Client:', v(d,'pass_c'), 'Partner:', v(d,'pass_p'))
    add_hdr(doc, 'Spouse or Partner Details')
    add_field2(doc, 'Has Spouse/Partner:', v(d,'partner'), 'Contrary Interest:', v(d,'contrary'))
    if v(d,'partner') == 'Yes':
        add_field2(doc, 'Partner Name:', v(d,'pfname')+' '+v(d,'plname'), 'Partner DOB/NI:', v(d,'pdob')+' / '+v(d,'pni'))
    add_hdr(doc, 'Dependants')
    add_field2(doc, 'Dependants with client:', v(d,'dep1'), 'Dependants not with client:', v(d,'dep2'))
    add_hdr(doc, 'D.  Capital and Other Assets')
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    for cell, hdr_text in zip(table.rows[0].cells, ['Bank', 'Acct No (last 4)', 'Type', 'Balance']):
        shade_cell(cell, '1F3864')
        r = cell.paragraphs[0].add_run(hdr_text)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for name, num, typ, bal in [(v(d,'ba1n'),v(d,'ba1num'),v(d,'ba1t'),v(d,'ba1b')),(v(d,'ba2n'),v(d,'ba2num'),v(d,'ba2t'),v(d,'ba2b'))]:
        if name or num:
            row = table.add_row()
            for cell, val in zip(row.cells, [name, num, typ, ('£'+bal) if bal else '']):
                cell.paragraphs[0].add_run(val).font.size = Pt(9)
    doc.add_paragraph()
    add_hdr(doc, 'E.  Income Details')
    add_field(doc, 'Non-passport Benefits:', v(d,'nonpass'))
    add_field2(doc, 'Pay/Sick Pay (weekly net):', ('£'+v(d,'earn_pay')) if v(d,'earn_pay') else '--', 'Self-employed drawings:', ('£'+v(d,'earn_se')) if v(d,'earn_se') else '--')
    add_field(doc, 'Documentary Evidence Provided:', v(d,'evidence'))
    add_hdr(doc, 'Applicant Declaration')
    add_decl_items(doc, [
        'This is a true statement of my personal and financial circumstances.',
        'I understand that if I give false information to the Scottish Legal Aid Board (SLAB), I may be prosecuted.',
        'I understand that SLAB can make any enquiries and get any information it needs.',
        'I agree to SLAB checking information with my employer, banks, DWP and HMRC.',
        'I must tell my solicitor immediately if there are any changes in my or my partner financial circumstances.',
        'If my solicitor does special urgency work I know SLAB may require me to pay a contribution.',
        'I agree to disclosure of the application and case file to SLAB for audit and quality assurance.',
        'SLAB may share information for fraud prevention. I consent to disclosure of my personal data.',
        'All consents are effective for not less than five years from the date of signature.',
    ])
    add_sig_table(doc, 'Signature of Applicant / Representative:', v(d,'signame'), v(d,'sigdate'))
    add_hdr(doc, 'Solicitor Declaration')
    add_decl_items(doc, [
        'I consent to the disclosure of the application and client case file for quality assurance including audit and peer review.',
        'I accept responsibility for the completion and submission of the application on Legal Aid Online (LAOL).',
        'I will retain this signed document in paper form or electronically and will send it to SLAB upon request.',
    ])
    add_sig_table(doc, 'Signature of Solicitor:', 'Michael Thompson', v(d,'sigdate'))
    add_hdr(doc, 'Partner Declaration')
    add_field(doc, 'Partner name:', v(d,'pfname') + ' ' + v(d,'plname'))
    doc.add_paragraph('I have seen the financial information given in this declaration. It is a true statement of my personal and financial circumstances.').font.size = Pt(9)
    add_sig_table(doc, 'Signature of Partner:', '', v(d,'sigdate'))
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def send_email(d, aa_buf, civ_buf):
    lname = v(d, 'lname')
    fname = v(d, 'fname')
    ref = v(d, 'ref')
    body = f"""New SLAB declaration submitted online.

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
    message = Mail(
        from_email=FROM_EMAIL,
        to_emails=FIRM_EMAIL,
        subject=f'New Legal Aid Declaration - {fname} {lname} - {ref}',
        plain_text_content=body
    )
    for buf, filename in [(aa_buf, f'AA_LAO_CIV_{lname}_{fname}.docx'), (civ_buf, f'CIV_SOL_{lname}_{fname}.docx')]:
        encoded = base64.b64encode(buf.read()).decode()
        attachment = Attachment(
            FileContent(encoded),
            FileName(filename),
            FileType('application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            Disposition('attachment')
        )
        message.attachment = attachment
    sg = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
    sg.send(message)

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
        print('Error:', str(e))
        import traceback
        traceback.print_exc()
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/')
def home():
    return 'SLAB Forms Server - OK'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
