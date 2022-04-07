from PIL import Image
from numpy import zeros
import pytesseract
import io,os,openpyxl,re,cv2,fitz
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

vals = [None] * 5

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
			image.save(open(f"tmp/res_img.png", "wb"))

def excell_check():
	book = Workbook()
	sheet = book.active
	sheet['A1'] = 'NR DOCUMENT'
	sheet['B1'] = 'DATA DOCUMENT'
	sheet['C1'] = 'NUME FURNIZOR'
	sheet['D1'] = 'DENUMIRE ARTICOL'
	sheet['E1'] = 'CANTITATE'
	book.save('avize.xlsx')

def excell_write(texte):
	book = openpyxl.load_workbook('avize.xlsx')
	sheet = book.active
	sheet.append(texte)
	book.save('avize.xlsx')

def image_section(folder,section):

	img_rgb = cv2.imread('tmp/res_img.png')
	img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
	template = cv2.imread(f'templates/{folder}/{section}.png', cv2.IMREAD_GRAYSCALE)

	w, h = template.shape

	res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
	_, _, _, maxLoc=cv2.minMaxLoc(res)

	#cv2.rectangle(img_rgb, maxLoc, (maxLoc[0]+h, maxLoc[1]+w), (0, 255, 255), 2)

	if section == 'product':
		offset = int(w/2)-10
	else:
		offset = 0

	crop_img = img_rgb[maxLoc[1]+offset:maxLoc[1]+w, maxLoc[0]:maxLoc[0]+h, :] 

	cv2.imwrite(f'tmp/res_{section}.png', crop_img)

	extract_data(section)

def extract_data(section):
	global vals

	data = pytesseract.image_to_string(Image.open(f"tmp/res_{section}.png"), lang="eng")

	if section == 'nr_data':

		patern = re.search(r"(Numar\s?.*?\:?|Nr\..*?\:).*?(\d{1,10})", data, re.DOTALL )
		if patern:
			vals[0] = patern.group(2)

		patern = re.search(r"(Data.*?\:).*?([0-9\.]{10})", data, re.DOTALL )
		if patern:
			vals[1] = patern.group(2)

	if section == 'product':
		
		product = data.strip().split('\n')
		product = product[len(product)-1]

		patern1 = re.search(r"(\w{2,}\s?.*?)\s(?:\w{2})\s(\d{1,}.\d{2})", product, re.DOTALL )
		patern2 = re.search(r"(\w{2,}\s?[A-z0-9\-\_]*?)\s{1,}(?:\w{2})\s{1,}(\d{1,}.\d{2})", data.replace("\n", " "), re.DOTALL )

		if patern1:
			print("Tipar 1")
			vals[3] = patern1.group(1).strip()
			vals[4] = patern1.group(2).strip()

		elif patern2:
			print("Tipar 2")
			vals[3] = patern2.group(1).strip()
			vals[4] = patern2.group(2).strip()		

def identify_template():
	global vals

	with open('furnizori.txt') as f:
		contents = f.read().splitlines()

	furnizori = '|'.join(contents)
	
	data = pytesseract.image_to_string(Image.open(f"tmp/res_img.png"), lang="eng")
	patern = re.search(f"(Furnizor\:).*?({furnizori})", data, re.DOTALL|re.IGNORECASE )

	if patern:
		vals[2] = patern.group(2).strip()
		return patern.group(2).strip().split(" ")[0].lower().strip()
	else:
		return 'ERROR'

def main():
	root = tk.Tk()
	root.withdraw()
	if not os.path.exists('avize.xlsx'):
		excell_check()

	file_path = filedialog.askopenfilename(initialdir = "/", title = "Select file", filetypes = (("PDF","*.pdf"),("All files","*.*")))
	print(file_path)
	if file_path:
		extract_img_pdf(file_path)
		furnizor = identify_template()

		if not furnizor == 'ERROR':
			image_section(furnizor,'nr_data')
			image_section(furnizor,'product')
		
			excell_write(vals)

		print(vals)

if  __name__ == '__main__':
	main()