import os
import csv
import sys
import subprocess
from datetime import datetime

from flask import Flask, request, redirect, url_for, flash, render_template
from werkzeug.utils import secure_filename

BASE_FOLDER = os.path.split(os.path.abspath(__file__))[0]
SPECIAL_KEYS = ("staff_name", "Title Code", "Start", "End", "Payrate","-", "PI")
OUTPUT_FIELDNAMES = ("Name", "PI", "Title Code", "Dist %", "Acct", "Subacct", "Start", "End", "Payrate")
DATE_FORMAT ="%Y-%m-%d_%H-%M-%S"
UPLOAD_FOLDER = os.path.join(BASE_FOLDER, "uploads")
TRANSFORM_FOLDER = os.path.join(BASE_FOLDER, "transforms")
STATIC_FOLDER = os.path.join(BASE_FOLDER, "static")
DOWNLOADS_FOLDER = os.path.join(BASE_FOLDER, "static", "downloads")
#CONVERTER_PATH = os.path.join(sys.exec_prefix, "Scripts", "xlsx2csv")
CONVERTER_PATH = os.path.join(BASE_FOLDER, "xls2csv", "xlsx2csv.py")
ALLOWED_EXTENSIONS = ('.xlsx', '.csv')
SECRET_KEY = b"W\x8c\xb8\xf6I3\\1\xdb\xbdZ'\x90\x08\xb5v\xf1 \xff\xa8\x15v1R"

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = SECRET_KEY

if not os.path.exists(UPLOAD_FOLDER):
	os.mkdir(UPLOAD_FOLDER)
if not os.path.exists(TRANSFORM_FOLDER):
	os.mkdir(TRANSFORM_FOLDER)
if not os.path.exists(STATIC_FOLDER):
	os.mkdir(STATIC_FOLDER)
if not os.path.exists(DOWNLOADS_FOLDER):
	os.mkdir(DOWNLOADS_FOLDER)


class CSV_Error(BaseException):
	pass


@app.route('/', methods=['GET', 'POST'])
def upload_file():
	"""
		From Flask tutorials at http://flask.pocoo.org/docs/0.12/patterns/fileuploads/
	:return:
	"""
	if request.method == 'POST':
		# check if the post request has the file part
		if 'file' not in request.files:
			flash('No file part')
			return redirect(request.url)
		file = request.files['file']
		# if user does not select file, browser also
		# submit a empty part without filename
		if file.filename == '':
			flash('No selected file')
			return redirect(request.url)
		if file and allowed_file(file.filename):
			new_filename = secure_filename(file.filename)
			new_path = os.path.join(app.config['UPLOAD_FOLDER'], new_filename)
			file.save(new_path)

			try:
				converted_file = reformat_file(new_path)
			except CSV_Error:
				flash("Unable to Convert to CSV!")
				return render_template("index.html")

			return render_template("download.html", file_url="/static/downloads/{}".format(converted_file))
		elif file:  # aka, did allowed_file fail?
			flash("File Extension not allowed!")

	return render_template("index.html")


def check_and_convert_xlsx(path):

	outfile = os.path.join(TRANSFORM_FOLDER, "xlsx_converted_{}".format(datetime.strftime(datetime.now(), DATE_FORMAT)))
	if os.path.splitext(path)[1].lower() == ".xlsx":
		return_code = subprocess.call([sys.executable, CONVERTER_PATH, path, outfile])
		if return_code != 0:
			raise CSV_Error("Unable to convert to CSV! Return Code = {}".format(return_code))
		return outfile
	else:
		return path


def modify_account_headers(path, header_row=0, account_row=2, subaccount_row=1):

	new_file_path = os.path.join(TRANSFORM_FOLDER, "corrected_header_{}.csv".format(datetime.strftime(datetime.now(), DATE_FORMAT)))

	with open(path) as csv_file:
		csv_data = csv.reader(csv_file)
		data = []
		for row in csv_data:
			data.append(row)

	number_of_columns = len(data[header_row])
	new_header = [None] * number_of_columns  # make a new, empty header of the correct length

	for col, value in enumerate(data[header_row]):
		if value == "":
			new_header[col] = "{}-{}".format(data[account_row][col], data[subaccount_row][col])  # write the items to the new header
		else:
			new_header[col] = data[header_row][col]  # when the header is defined, keep it

	new_header[0] = "staff_name"  # override default which would read it as "account-subacct"

	with open(new_file_path, 'w', newline="\n", encoding="utf-8") as outfile:
		writer = csv.writer(outfile)
		writer.writerow(new_header)
		for row in data[3:]:  # get all of the rows after the crappy header
			writer.writerow(row)

	return new_file_path


def reformat_file(path, field_order=OUTPUT_FIELDNAMES):

	messages = clean_folder(TRANSFORM_FOLDER)
	messages.append(clean_folder(DOWNLOADS_FOLDER))

	file_path = modify_account_headers(check_and_convert_xlsx(path))  # gives us a CSV, then corrects the headers on that CSV to something normal
	output_file = os.path.join(DOWNLOADS_FOLDER, "transformed_{}.csv".format(datetime.strftime(datetime.now(), DATE_FORMAT)))

	output_rows = []
	with open(file_path) as input_file:
		data = csv.DictReader(input_file)
		for row in data:
			output_rows += convert_row(row)

	with open(output_file, 'w', newline="\n", encoding="utf-8") as output:
		writer = csv.DictWriter(output, fieldnames=field_order)
		writer.writeheader()
		writer.writerows(output_rows)

	messages.append(clean_folder(UPLOAD_FOLDER))

	return os.path.split(output_file)[1]


def clean_folder(path):
	messages = []
	for filename in os.listdir(path):
		try:
			os.remove(os.path.join(path, filename))
		except:
			messages.append("Unable to delete {}".format(os.path.join(path, filename)))

	return messages


def convert_row(row, special_keys=SPECIAL_KEYS):

	output_rows = []
	for key in list(set(row.keys())-set(special_keys)):  # for all the keys that aren't in special_keys, bascially, for all of the accounts:

		if row[key] is None or row[key] == "" or row[key] == " ":  # if this record isn't in use for the individual
			continue

		if "PI" in row:
			pi = row["PI"]
		else:
			pi = None

		acct, sbacct = key.split("-")

		output_rows.append({
							"Name": row["staff_name"],
							"PI": pi,
							"Title Code": row["Title Code"],
							"Dist %": row[key],
							"Acct": acct,
							"Subacct": sbacct,
							"Start": row["Start"],
							"End": row["End"],
							"Payrate": row["Payrate"]
		})
	return output_rows


def allowed_file(filename):
	"""
		From Flask tutorials at http://flask.pocoo.org/docs/0.12/patterns/fileuploads/
	:param filename:
	:return:
	"""
	return '.' in filename and os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


if __name__ == '__main__':
	app.run()
