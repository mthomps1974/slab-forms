from flask import Flask, request, jsonify
from flask_cors import CORS
import os, io, base64
from datetime import datetime

app = Flask(__name__)
CORS(app)

SENDGRID_API_KEY = os.environ.get('SENDGRID_API_KEY', '')
FIRM_EMAIL = os.environ.get('FIRM_EMAIL', 'info@tflaw.co.uk')
FROM_EMAIL = os.environ.get('FROM_EMAIL', '')

def v(d, key):
    return str(d.get(key, '') or '')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        d = request.json
        ref = 'TFL-' + str(int(datetime.now().timestamp()))[-6:]

        print('=== SUBMIT RECEIVED ===')
        print('Client:', v(d,'fname'), v(d,'lname'))
        print('SENDGRID_API_KEY set:', bool(SENDGRID_API_KEY))
        print('FROM_EMAIL:', FROM_EMAIL)
        print('FIRM_EMAIL:', FIRM_EMAIL)

        # Step 1: test sendgrid import
        print('Importing sendgrid...')
        import sendgrid
        from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
        print('Sendgrid imported OK')

        # Step 2: test docx import
        print('Importing docx...')
        from docx import Document
        from docx.shared import Pt, RGBColor, Twips
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        print('Docx imported OK')

        # Step 3: build a simple test doc
        print('Building test doc...')
        doc = Document()
        doc.add_paragraph('Test: ' + v(d,'fname') + ' ' + v(d,'lname'))
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        print('Doc built OK, size:', len(buf.getvalue()))

        # Step 4: send email
        print('Sending email...')
        message = Mail(
            from_email=FROM_EMAIL,
            to_emails=FIRM_EMAIL,
            subject='Test - ' + v(d,'fname') + ' ' + v(d,'lname') + ' - ' + ref,
            plain_text_content='Test submission from ' + v(d,'fname') + ' ' + v(d,'lname')
        )
        encoded = base64.b64encode(buf.getvalue()).decode()
        attachment = Attachment(
            FileContent(encoded),
            FileName('test.docx'),
            FileType('application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            Disposition('attachment')
        )
        message.attachment = attachment

        sg = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
        response = sg.send(message)
        print('Email sent! Status:', response.status_code)

        return jsonify({'ok': True, 'ref': ref})

    except Exception as e:
        import traceback
        print('=== ERROR ===')
        print(str(e))
        traceback.print_exc()
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/')
def home():
    return 'SLAB Forms Server - OK'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
