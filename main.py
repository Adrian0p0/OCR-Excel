from PIL import Image
from numpy import zeros
import pytesseract
import fitz
import io,os,openpyxl,re
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook


def process_image(iamge_name, lang_code):
	return pytesseract.image_to_string(Image.open(iamge_name), lang=lang_code)


def extract_img_pdf(file):
	pdf_file = fitz.open(file)

	for page_index in range(len(pdf_file)):
		page = pdf_file[page_index]
		
		for image_index, img in enumerate(page.get_images(), start=1):
			xref = img[0]
			base_image = pdf_file.extract_image(xref)
			image_bytes = base_image["image"]
			image_ext = base_image["ext"]
			image = Image.open(io.BytesIO(image_bytes))
			image.save(open(f"image.png", "wb"))
			extract_data()

def excell_check():
	book = Workbook()
	sheet = book.active
	sheet['A1'] = 'NR DOCUMENT'
	sheet['B1'] = 'DATA DOCUMENT'
	sheet['C1'] = 'NUME FURNIZOR'
	sheet['D1'] = 'DENUMIRE ARTICOL'
	sheet['E1'] = 'CANTITATE'
	book.save('avize.xlsx')
	print('Excel Created')

def excell_write(texte):
	book = openpyxl.load_workbook('avize.xlsx')
	sheet = book.active
	data = (texte[0], texte[1], texte[2], texte[3], texte[4], texte[5], texte[6], texte[7], texte[8], texte[9])
	sheet.append(data)
	book.save('avize.xlsx')
	print('Excel added data')

def extract_data():
	data = process_image("image.png", "eng")

	file = open('data.txt', 'w')
	file.write(data)
	file.close()

	vals = [None] * 5

	patern = re.search(r"(Nr\.\saviz|Numar\saviz|Nr\.\s?tichet).*?(\d{1,10})", data, re.IGNORECASE )

	if patern:
		if patern.group(2):
			vals[0] = patern.group(2)
		else:
			vals[0] = 'ERROR'
	else:
		vals[0] = 'ERROR'

	patern = re.search(r"Data.*?(\d{2}\.\d{2}\.\d{4})", data, re.IGNORECASE )

	if patern:
		if patern.group(1):
			vals[1] = patern.group(1)
		else:
			vals[1] = 'ERROR'
	else:
		vals[1] = 'ERROR'

	patern = re.search(r"Furnizor:(.*?)[Nn]r\.|furnizor:.*?Nume:(.*?)Nume:", data, re.DOTALL )

	if patern:
		if patern.group(1):
			vals[2] = patern.group(1).strip()
		elif patern.group(2):
			vals[2] = patern.group(2).strip()
		else:
			vals[2] = 'ERROR'
	else:
		vals[2] = 'ERROR'

		patern = re.search(r"Furnizor:(.*?)[Nn]r\.|furnizor:.*?Nume:(.*?)Nume:", data, re.DOTALL )

	print(patern.group(0))
	print(patern.group(1))

	if patern:
		if patern.group(1):
			vals[3] = patern.group(1).strip()
		else:
			vals[3] = 'ERROR'
	else:
		vals[3] = 'ERROR'

	print(vals)
	#excell_write(["wqe",'asda','asdasd','asda','asdasd','asda','asdasd','asda','asdasd','asdasd'])

if  __name__ == '__main__':
	root = tk.Tk()
	root.withdraw()
	if not os.path.exists('avize.xlsx'):
		excell_check()

	file_path = filedialog.askopenfilename()
	extract_img_pdf(file_path)

	

	
	