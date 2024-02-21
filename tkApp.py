import tkinter as tk
import json
import pandas as pd
import openpyxl
from tkinter import Tk, messagebox, filedialog

root=tk.Tk()
root.title("Json To Excel by dsp")
# setting the windows size
root.geometry("300x100")

json_string = tk.StringVar()

def submit():

	data = json_string.get()
	dict = json.loads(json.loads(data))
	name = []
	secondValue= []
	place = []
	url = []

	x = dict.get('Companies')

	for items in range(len(x)):
		name.append(x[items].get('CompanyName'))
		secondValue.append(x[items].get('Nace2Description'))
		place.append(x[items].get('PrefectureName'))
		url.append("www.findbiz.gr" + x[items].get('FriendlyUrl'))

	df = pd.DataFrame({
		'ΕΠΩΝΥΜΙΑ': name,
		'ΔΡΑΣΤΗΡΙΟΤΗΤΑ': secondValue,
		'ΠΕΡΙΟΧΗ': place,
		'URL' : url,
	}) 
	
	# Write DataFrame to Excel 
	file_excel = filedialog.asksaveasfilename(filetypes=[('excel file','*.xlsx')], defaultextension='.xlsx')
	df.to_excel(file_excel, sheet_name='CR_DUMP')
	messagebox.showinfo("ΤΟ ΑΡΧΕΙΟ ΑΠΟΘΗΚΕΥΤΗΚΕ", "Success")

	
json_label = tk.Label(root, text = 'Json String', font=('calibre',10, 'bold'))
json_entry = tk.Entry(root,textvariable = json_string, font=('calibre',10,'normal'))
json_entry.insert(0, "Enter json string")

sub_btn=tk.Button(root,text = 'Submit', command = submit)

json_label.grid(row=0,column=0, sticky="ew")
json_entry.grid(row=0,column=1, sticky="ew")

sub_btn.grid(row=0,column=2)

root.grid_columnconfigure((0,3), weight=1)
root.grid_rowconfigure((0,1), weight=1)


# performing an infinite loop 
# for the window to display
root.mainloop()
