import os
from flask import Flask, request, redirect, url_for, render_template, send_from_directory
from werkzeug.utils import secure_filename
import smtplib
from smtplib import SMTP_SSL as SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import re 
import io
from windwardrestapi.Model import Template, SqlDataSource
from windwardrestapi.Api import WindwardClient as client
import os
import time
import base64
import zipfile


app = Flask(__name__)

app.config['upload_folder'] = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
app.config['generated_docs_folder'] = os.path.dirname(os.path.abspath(__file__)) + '/generated_docs/'

app.config['SMTPserver'] = 'email1131.luxsci.com'
app.config['SMTPport'] = 465
app.config['SSL'] = True #if false, then TLS
app.config['sender_address'] = 'sandbox@babsondx.com'
app.config['sender_password'] = 'B@bs0nDx!1205'

#generating document
windwardClientAddress = "http://localhost:8080/"
datasourceName = "HarvestODBC"
datasourceClassName = "System.Data.Odbc"
datasourceConnectionString = "driver=4D v15 ODBC Driver 64-bit;SERVER=10.0.0.20;PORT=19812;UID=Orchard;pwd=B@bs0nDx!"


allowedFileTypes = {'docx'}
def isAllowedFile(filename):
   return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowedFileTypes

def getAllUploadedTemplates():
	files = []
	for file in os.listdir(app.config['upload_folder']):
		if file.endswith(".docx"):
			files.append(file.rsplit('.', 1)[0])
	return files

def generateDocFromTemplate (filename):
	newClient = client.WindwardClient(windwardClientAddress)
	sqlDS = SqlDataSource.SqlDataSource(name=datasourceName, className=datasourceClassName, connectionString=datasourceConnectionString)
	file_path = os.path.abspath(os.path.join(app.config['upload_folder'], filename))
	templatePDF = Template.Template(data=file_path, outputFormat=Template.outputFormatEnum.PDF, datasources=sqlDS)
	generatedDoc = newClient.postDocument(templatePDF)
	while True:
		document_status = newClient.getDocumentStatus(generatedDoc.guid)
		if document_status != 302:
			time.sleep(1)
		else:
			break
	getDocument = newClient.getDocument(generatedDoc.guid)
	new_doc_name = "generated_docs/" + filename.rsplit('.', 1)[0] + ".pdf"
	with open(new_doc_name, "wb") as fh:
		fh.write(base64.standard_b64decode(getDocument.data))


def email (form):
	recipients = re.findall('[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}', form['recipients'])
	for rec in recipients:
		###creating email
		msg = MIMEMultipart()
		msg['From'] = sender_address
		msg['To'] = rec
		msg['Subject'] = form['subject']
		msg.attach(MIMEText(form['body'], 'plain'))

		#attaching pdf
		with open("generated_docs/" + form['attachment'], "rb") as file:
			docAttachment = MIMEBase("application", "octet-stream")
			docAttachment.set_payload(file.read())
		encoders.encode_base64(docAttachment)

		docAttachment.add_header("Content-Disposition", f"attachment; filename= {form['attachment']}",)
		msg.attach(docAttachment)

		if SSL:
			s = smtplib.SMTP_SSL("{}:{}".format(SMTPserver, SMTPport))
			s.login(sender_address, sender_password)
			s.sendmail(sender_address, rec, msg.as_string())
			s.quit()
		else:
			s = smtplib.SMTP(SMTPserver, SMTPport)
			s.starttls() 
			s.login(sender_address, sender_pass)
			s.sendmail(sender_address, rec, msg.as_string()) 
			s.quit() 
	return recipients


@app.route('/', methods=['GET', 'POST'])
def main_func():
	return redirect(url_for('Documents'))

@app.route('/Documents', methods=['GET', 'POST'])
def Documents():
	if request.method == 'POST':
		if 'uploadNewTemplate' in request.form:
			return render_template('documents.html', display="Upload new template modal", files=getAllUploadedTemplates())
		elif 'uploadedFile' in request.form:
			if 'file' in request.files:
				file = request.files['file']
				if file and file.filename != '' and isAllowedFile(file.filename):
					filename = secure_filename(file.filename)
					file.save(os.path.join(app.config['upload_folder'], filename))
					return render_template('documents.html', display="Upload success", files=getAllUploadedTemplates())
			return render_template('documents.html', display="Upload failure", files=getAllUploadedTemplates())
		elif 'seeMoreOptions' in request.form or 'goBack' in request.form:
			return render_template('documents.html', display = "More options", file_name=request.form['file_name'], files=getAllUploadedTemplates())
		elif 'downloadTemplate' in request.form:
			return send_from_directory(app.config['upload_folder'], request.form['file_name'] + ".docx", as_attachment=True) 
		elif 'downloadPDF' in request.form:
			generated_doc_already_created = os.path.exists(os.path.join(app.config['generated_docs_folder'], request.form['file_name'] + ".pdf"))
			if not generated_doc_already_created:
				generateDocFromTemplate(secure_filename(request.form['file_name'] + ".docx"))
			return send_from_directory(app.config['generated_docs_folder'], request.form['file_name'] + ".pdf", as_attachment=True) 
		elif 'email' in request.form or 'goBackToEmail' in request.form:
			generated_doc_already_created = os.path.exists(os.path.join(app.config['generated_docs_folder'], request.form['file_name'] + ".pdf"))
			if not generated_doc_already_created:
				generateDocFromTemplate(secure_filename(request.form['file_name'] + ".docx"))
			return render_template('documents.html', display="Create email", file_name=request.form['file_name'], files=getAllUploadedTemplates())
		elif 'emailSend' in request.form:
			recipients = email(request.form)
			return render_template('documents.html', display="Email sent", recipients=recipients, file_name=request.form['file_name'], files=getAllUploadedTemplates())
		elif 'delete' in request.form:
			return render_template('documents.html', display="Confirm delete", file_name=request.form['file_name'], files=getAllUploadedTemplates())
		elif 'deleteConfirmYes' in request.form:
			path_to_template = os.path.join(app.config['upload_folder'], request.form['file_name'] + ".docx")
			path_to_generated_doc = os.path.join(app.config['generated_docs_folder'], request.form['file_name'] + ".pdf")
			if (os.path.exists(path_to_template)):
				os.remove(path_to_template)
			if (os.path.exists(path_to_generated_doc)):
				os.remove(path_to_generated_doc)
			return render_template('documents.html', display="Deleted", file_name=request.form['file_name'], files=getAllUploadedTemplates())
	return render_template('documents.html', display="Main screen", files=getAllUploadedTemplates())


@app.route('/SMTP', methods=['GET', 'POST'])
def SMTP():
	edit_mode = False
	if request.method == 'POST':
		if 'edit' in request.form:
			edit_mode = True
		elif 'save' in request.form:
			app.config['SMTPserver'] = request.form['server']
			app.config['SMTPport'] = request.form['port']
			if request.form['SSLorTLS'] == 'SSL':
				app.config['SSL'] = True
			else:
				app.config['SSL'] = False
			app.config['sender_address'] = request.form['email address']
			app.config['sender_password'] = request.form['password']
	return render_template('SMTP.html', SMTPserver=app.config['SMTPserver'], SMTPport=app.config['SMTPport'], SSL=app.config['SSL'], sender_address=app.config['sender_address'], sender_password=app.config['sender_password'], edit_mode=edit_mode)

if __name__ == '__main__':
	app.run(host='127.0.0.1', port=8082, debug=True)


