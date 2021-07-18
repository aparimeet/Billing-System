import tkinter as tk
from tkinter.constants import ACTIVE, DISABLED, END
import docx2pdf
from docx2pdf import convert
from docx import Document
from docx.shared import Inches, Cm, Mm, Emu
from docxtpl import DocxTemplate, InlineImage
import random
import datetime


#Import template document
template = DocxTemplate('template.docx')

#To create 4 digit randam number
def generate_invoice_number ():
	file = open("invoice_numbers.txt", "r")
	invoice_number = int(file.read())
	invoice_number += 1
	file.close()
	file = open("invoice_numbers.txt", "w")
	file.write(str(invoice_number) + "\n")
	file.close()
	return invoice_number


itemName = 0
itemPrice = 0
totalPrice = 0
table_contents = []
itemNameList = []
itemPriceList = []

choice = True
while(choice == True):
	itemName = int(input("Enter the model number: "))
	itemPrice = int(input("Enter the price: "))

	#itemNameList.append(itemName)
	itemPriceList.append(itemPrice)
	totalPrice = sum(itemPriceList)


	table_contents.append({
		'Item_Name' : itemName,
		'Price' : itemPrice
		})

	print("Do you want to add any more items?")
	check = input()
	if check in {"Y", "YES", "Yes", "yes", "y"}:
		choice = True
	else:
		choice = False

image2 = InlineImage(template,'makeinindia.png',Cm(5))

#Add values to template
context = {
	'title': 'Purchase Invoice',
	'day': datetime.datetime.now().strftime('%d'),
    'month': datetime.datetime.now().strftime('%b'),
    'year': datetime.datetime.now().strftime('%Y'),
    'table_contents': table_contents,
    #'image1': image
    'image2' : image2,
    'totalPrice': totalPrice
}

#Render automated report
template.render(context)
invoice_number = generate_invoice_number()
filename = 'generated_invoice_' + str(invoice_number) + '.docx'
template.save(filename)

convert(filename)
