import numpy
import pandas as pd
from reportlab.platypus import BaseDocTemplate, Frame, Table, TableStyle, PageTemplate, Paragraph, Spacer
from reportlab.lib.pagesizes import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime
import re
import os
import shutil
import warnings
from PyPDF2 import PdfMerger
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

contact = "Anastasiya Slaughter"
phone = "734-444-4624"
email = "anastasiya.slaughter@einsteinmed.edu"

columns = {
	'cage_barcode': 7,
	'mouseline': 5,
	'tags': 4,
	'genotype': 6,
	'disposition': 2
}

out_dir = "tmp_output/"

def createCard(general_data, specific_data, index):
	card_width = 5 * inch
	card_height = 3 * inch

	doc = BaseDocTemplate(
		out_dir + "card" + str(index) + ".pdf",
	    pagesize=(card_width, card_height),
	    leftMargin=0,
	    rightMargin=0,
	    topMargin=0,
	    bottomMargin=0
	)

	frame = Frame(
	    x1=0,
	    y1=0,
	    width=card_width,
	    height=card_height,
	    showBoundary=1
	)

	template = PageTemplate(id='CardTemplate', frames=[frame])
	doc.addPageTemplates([template])

	styles = getSampleStyleSheet()
	title_style = styles["Heading3"]
	title_style.alignment = 0  # 0=left, 1=center align, 2=right
	info_style = styles["Normal"]
	info_style.fontSize = 6

	# title = Paragraph("Flashcard Title", title_style)

	info_data = [
	    [Paragraph("Barcode: " + general_data['barcode'], info_style), Paragraph("Contact: "+ general_data['contact'], info_style)],
	    [Paragraph("Mouseline: " + general_data['mouseline'], info_style), Paragraph("Phone Number: " + general_data['phone'], info_style)],
	    [Paragraph("Set up Date: " + general_data['date'], info_style), Paragraph("Study: " + general_data['study'], info_style)],
	    [Paragraph("Count: " + general_data['count'], info_style), Paragraph("Email: " + general_data['email'], info_style)],
	]

	info_table = Table(
	    info_data,
	    colWidths=[card_width / 3] * 2,
	    hAlign='LEFT',
	    style=[
	        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
	        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
	        ('TOPPADDING', (0, 0), (-1, -1), 0),
	        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
	        ('LEFTPADDING', (0, 0), (-1, -1), 0),
	        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
	    ]
	)

	cell_style = styles['Normal']
	cell_style.fontSize = 6
	cell_style.leading = 8  # line height

	data = [[Paragraph(cell, cell_style) for cell in row] for row in specific_data]

	main_table = Table(data)

	main_table.setStyle(TableStyle([
	    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
	    ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
	    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
	    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
	    ('LEFTPADDING', (0, 0), (-1, -1), 4),
	    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
	    ('TOPPADDING', (0, 0), (-1, -1), 2),
	    ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
	]))

	if general_data['disposition'].lower() == "mating":
		mating_style = ParagraphStyle(
		    name="RedOnBlack",
		    parent=styles["Heading3"],
		    textColor=colors.red,
		    backColor=colors.black,
		    alignment=1  # Center align
		)

		mating = Paragraph("Mating!", mating_style)
	
		elements = [
		    info_table,
		    main_table,
		    mating
		]
	else:
		elements = [
		    info_table,
		    main_table
		]

	doc.build(elements)

def getData(row, index):
	general_data = {}
	general_data['barcode'] = str(int(row.iloc[columns["cage_barcode"]]))
	general_data['mouseline'] = row.iloc[columns["mouseline"]]
	general_data['date'] = datetime.today().strftime("%m-%d-%Y")
	general_data['count'] = ''
	general_data['contact'] = contact
	general_data['phone'] = phone
	general_data['study'] = ''
	general_data['email'] = email
	general_data['disposition'] = row.iloc[columns["disposition"]]

	specific_data = [['Tag', "Alt. ID", "Sex", "DOB", "Genotype"]]

	num_mice = len(row.iloc[columns["genotype"]].splitlines())
	tags = row.iloc[columns["tags"]].splitlines()
	genotypes = row.iloc[columns["genotype"]].splitlines()

	for i in range(0, num_mice):
		new_row = [
		    re.match(r'^(\d+)', tags[i]).group(1) if re.match(r'^(\d+)', tags[i]) else "",
		    '',
		    re.search(r'\[\s*(\w),', tags[i]).group(1) if re.search(r'\[\s*(\w),', tags[i]) else "",
		    re.search(r'(\d{2}-\d{2}-\d{4})', tags[i]).group(1) if re.search(r'(\d{2}-\d{2}-\d{4})', tags[i]) else "",
		    genotypes[i]
		]

		specific_data.append(new_row)

	createCard(general_data, specific_data, index)

def mergeFiles():
	output_path = "cards.pdf"

	merger = PdfMerger()

	pdf_files = sorted(f for f in os.listdir(out_dir) if f.lower().endswith('.pdf'))

	for pdf_file in pdf_files:
	    full_path = os.path.join(out_dir, pdf_file)
	    merger.append(full_path)

	merger.write(output_path)
	merger.close()

	print(f"Merged {len(pdf_files)} files into {output_path}")

def main():
	print("Main function")

	if os.path.exists(out_dir):
		shutil.rmtree(out_dir)

	os.makedirs(out_dir)

	df = pd.read_excel("softmousedb.xlsx", engine='openpyxl')

	# for index, row in df.iterrows():
	# 	getData(row, index)
	# 	if index > 2:
	# 		break

	for index, row in df.iterrows():
		try:
			getData(row, index)
		except Exception as e:
			print(f"Error at row {index}: {e}")

	mergeFiles()

	shutil.rmtree(out_dir)


# Standard boilerplate to run main()
if __name__ == "__main__":
    main()


