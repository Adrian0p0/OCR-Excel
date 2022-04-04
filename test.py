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

	data = pytesseract.image_to_data(Image.open('image.png'))

	print(data[90:99])

	vals = [None] * 5

	print(vals)
	#excell_write(["wqe",'asda','asdasd','asda','asdasd','asda','asdasd','asda','asdasd','asdasd'])

if  __name__ == '__main__':
	root = tk.Tk()
	root.withdraw()
	if not os.path.exists('avize.xlsx'):
		excell_check()

	file_path = filedialog.askopenfilename()
	extract_img_pdf(file_path)

	

	
	