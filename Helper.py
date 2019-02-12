from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from threading import Thread
from tkinter import messagebox #NEW
import shutil
import codecs
import os

from Lyca import processLyca
from Diapazon import diapazon
from KS import processKS
from Metropoliten import process_metropoliten

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
def accept_button():
	lyca_label_status['text']="Статус: Обработка..."
	new_process = Thread( target=processlyca, args = [folder_path.get() + "/", orderName.get(), jobNumber.get(), itf.get(), articul.get(), zakaz.get(), face.get(), 
															innerCode.get(), outerCode.get(), customer.get(), label30x20.get(), lyca_label_status] )
	lines = []

	with open("settings.ini", "r") as file:
		lines = file.readlines()
		lyca_setting_index = -1
		for entry_number in range(len(lines)):
			if "Лука" in lines[entry_number]:
				lyca_setting_index = entry_number
		if lyca_setting_index != -1:
			for index in range(len(lyca_list_of_data_entries) + 1):
				del lines[lyca_setting_index]
	lines.append("Лукамобайл\n")
	for object in lyca_list_of_data_entries:
		lines.append(str(object.get()) + "\n")
	with open("settings.ini", "w") as file:
		file.writelines(lines)

	new_process.start()
def shutDown():
	shutil.rmtree('__pycache__', ignore_errors=True)
	root.destroy()
def refresh(event):
	if tabControl.index("current") == 0:
		root.geometry("700x311")
	if tabControl.index("current") == 1:
		root.geometry("700x214")
	if tabControl.index("current") == 2:
		root.geometry("700x299")	
	if tabControl.index("current") == 3:
		root.geometry("700x132")	
#Подмена функции Paste в Windows
def paste(obj):
	obj.widget.delete(0, END)
	obj.widget.insert(0,root.clipboard_get())
	#obj.widget.insert(0,root.clipboard_get().encode(encoding='UTF-8',errors='replace'))
	
root = Tk()
root.event_delete('<<Paste>>', '<Control-v>')
#Подмена функции Paste в Windows
root.bind("<Control-v>",paste)
root.title("Helper  v0.004")

root.resizable(0, 0)
tabControl = ttk.Notebook(root)
tabControl.bind("<<NotebookTabChanged>>", refresh)
tab_lyca = ttk.Frame(tabControl)
tab_lyca_Line1 = ttk.Frame(tab_lyca)
tab_lyca_Line2 = ttk.Frame(tab_lyca)
tab_lyca_Line3 = ttk.Frame(tab_lyca)
tab_lyca_Line4 = ttk.Frame(tab_lyca)
tab_lyca_Line5 = ttk.Frame(tab_lyca)
tab_lyca_Line6 = ttk.Frame(tab_lyca)
tab_lyca_Line7 = ttk.Frame(tab_lyca)
tab_lyca_Line8 = ttk.Frame(tab_lyca)
tab_lyca_Line9 = ttk.Frame(tab_lyca)
tab_lyca_Line10 = ttk.Frame(tab_lyca)
tab_lyca_Line11 = ttk.Frame(tab_lyca)
tab_lyca_Line12 = ttk.Frame(tab_lyca)
tab_lyca_Line13 = ttk.Frame(tab_lyca)

tab_Diapazon = ttk.Frame(tabControl)
tab_Diapazon_Line1 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line2 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line3 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line4 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line5 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line6 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line7 = ttk.Frame(tab_Diapazon)
tab_Diapazon_Line8 = ttk.Frame(tab_Diapazon)

tab_KS = ttk.Frame(tabControl)
tab_KS_Line1 = ttk.Frame(tab_KS)
tab_KS_Line2 = ttk.Frame(tab_KS)
tab_KS_Line3 = ttk.Frame(tab_KS)
tab_KS_Line4 = ttk.Frame(tab_KS)
tab_KS_Line5 = ttk.Frame(tab_KS)
tab_KS_Line6 = ttk.Frame(tab_KS)
tab_KS_Line4_1 = ttk.Frame(tab_KS)
tab_KS_Line4_2 = ttk.Frame(tab_KS)
tab_KS_Line4_3 = ttk.Frame(tab_KS)
tab_KS_Line4_4 = ttk.Frame(tab_KS)
tab_KS_Line4_5 = ttk.Frame(tab_KS)
tab_KS_Line4_6 = ttk.Frame(tab_KS)

tab_Metropoliten = ttk.Frame(tabControl)
tab_Metropoliten_Line1 = ttk.Frame(tab_Metropoliten)
tab_Metropoliten_Line2 = ttk.Frame(tab_Metropoliten)
tab_Metropoliten_Line3 = ttk.Frame(tab_Metropoliten)
tab_Metropoliten_Line4 = ttk.Frame(tab_Metropoliten)

tabControl.add(tab_lyca, text="lyca")
tabControl.add(tab_Diapazon, text="Диапазон")
tabControl.add(tab_KS, text="Киевстар")
tabControl.add(tab_Metropoliten, text="Метрополитен")

tabControl.pack(expand=1, fill="both")

#lyca tab
folder_path = StringVar()
orderName = StringVar()
jobNumber = StringVar()
itf = StringVar()
articul = StringVar()
zakaz = StringVar()
face = StringVar()
innerCode = StringVar()
outerCode = StringVar()
customer = StringVar()
label30x20 = IntVar()
status = StringVar()
lyca_list_of_data_entries = [folder_path, orderName, jobNumber, itf, articul, zakaz, face, innerCode, outerCode, customer, label30x20]

status.set("Статус: Готов к работе")

tab_lyca_Line1.pack(side = TOP, fill = X)
label2 = Label(tab_lyca_Line1, text="Выберите папку с БД:", width = 17)
label2.pack(side = LEFT)
entry1 = Entry(master=tab_lyca_Line1,textvariable=folder_path)
entry1.pack(side = LEFT, expand = True, fill = BOTH)
button1 = Button(tab_lyca_Line1, text="Browse", command=browse_button)
button1.pack(side = RIGHT)

tab_lyca_Line2.pack(side = TOP, fill = X)
label3 = Label(tab_lyca_Line2, text="Название продукции:", width = 17)
label3.pack(side = LEFT)
entry2 = Entry(master=tab_lyca_Line2,textvariable=orderName)
entry2.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line3.pack(side = TOP, fill = X)
label4 = Label(tab_lyca_Line3, text="Job Number:", width = 17)
label4.pack(side = LEFT)
entry3 = Entry(master=tab_lyca_Line3,textvariable=jobNumber)
entry3.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line4.pack(side = TOP, fill = X)
label5 = Label(tab_lyca_Line4, text="ITF код:", width = 17)
label5.pack(side = LEFT)
entry4 = Entry(master=tab_lyca_Line4,textvariable=itf)
entry4.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line5.pack(side = TOP, fill = X)
label6 = Label(tab_lyca_Line5, text="Артикул:", width = 17)
label6.pack(side = LEFT)
entry5 = Entry(master=tab_lyca_Line5,textvariable=articul)
entry5.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line6.pack(side = TOP, fill = X)
label7 = Label(tab_lyca_Line6, text="Номер заказа:", width = 17)
label7.pack(side = LEFT)
entry6 = Entry(master=tab_lyca_Line6,textvariable=zakaz)
entry6.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line7.pack(side = TOP, fill = X)
label9 = Label(tab_lyca_Line7, text="Номинал:", width = 17)
label9.pack(side = LEFT)
entry7 = Entry(master=tab_lyca_Line7,textvariable=face)
entry7.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line8.pack(side = TOP, fill = X)
label10 = Label(tab_lyca_Line8, text="Внутренний код:", width = 17)
label10.pack(side = LEFT)
entry8 = Entry(master=tab_lyca_Line8,textvariable=innerCode)
entry8.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line9.pack(side = TOP, fill = X)
label11 = Label(tab_lyca_Line9, text="Внешний код:", width = 17)
label11.pack(side = LEFT)
entry9 = Entry(master=tab_lyca_Line9,textvariable=outerCode)
entry9.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line10.pack(side = TOP, fill = X)
label12 = Label(tab_lyca_Line10, text="Заказчик:", width = 17)
label12.pack(side = LEFT)
entry10 = Entry(master=tab_lyca_Line10,textvariable=customer)
entry10.pack(side = RIGHT, expand = True, fill = X)

tab_lyca_Line11.pack(side = TOP, fill = X)
label13 = Label(tab_lyca_Line11, text="Этикетка-сэндвич:", width = 17)
label13.pack(side = LEFT)
checkBox1 = Checkbutton(tab_lyca_Line11, variable=label30x20)
checkBox1.pack(side = LEFT)

tab_lyca_Line12.pack(side = TOP, fill = X)
button2 = Button(tab_lyca_Line12, text="Старт", command=accept_button)
button2.pack(fill = X)

tab_lyca_Line13.pack(side = TOP, fill = X)
lyca_label_status = Label(tab_lyca_Line12, text=status.get())
lyca_label_status.pack()


#Диапазон
diapazon_folder_path = StringVar()
diapazon_first_number = StringVar()
diapazon_last_number = StringVar()
diapazon_mode = IntVar()
diapazon_mode.set(1)
file_extension = ""

def diapazon_browse_button():
    global diapazon_folder_path
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[('XLSX', ".xlsx"),('TXT', ".txt")])
    diapazon_folder_path.set(filename)
def diapazon_accept_button():
	diapazon_difference = int(diapazon_last_number.get()) - int(diapazon_first_number.get()) + 1
	diapazon_type = "Нет контрольной цифры"
	if diapazon_mode.get() == 2:
		diapazon_type = "CODE 128"
	elif diapazon_mode.get() == 3:
		diapazon_type = "EAN 13"
	result = messagebox.askquestion("Продолжить?", diapazon_type + "\n\nКоличество записей: " + str(diapazon_difference) + "\n\nХотите продолжить?", icon='warning')
	if result == "yes":
		if diapazon_mode.get() == 3:
			error = 0
			if len(diapazon_first_number.get()) != 12:
				diapazon_first_number.set("НЕПРАВИЛЬНОЕ КОЛИЧЕСТВО СИМВОЛОВ")
				error = 1
			if len(diapazon_last_number.get()) != 12:
				diapazon_last_number.set("НЕПРАВИЛЬНОЕ КОЛИЧЕСТВО СИМВОЛОВ")
				error = 1
			if error == 1:
				return
		if diapazon_first_number.get() > diapazon_last_number.get():
			diapazon_last_number.set("ПОСЛЕДНИЙ НОМЕР МЕНЬШЕ ПЕРВОГО")
			return
		filename_temp, file_extension = os.path.splitext(diapazon_folder_path.get())
		#diapazon(diapazon_folder_path.get(), diapazon_first_number.get(), diapazon_last_number.get(), diapazon_mode.get(), file_extension, diapazon_label_status)
		new_process = Thread( target=diapazon, args = [ diapazon_folder_path.get(), diapazon_first_number.get(), diapazon_last_number.get(), diapazon_mode.get(), file_extension,
														diapazon_label_status] )
		new_process.start()

tab_Diapazon_Line1.pack(side = TOP, fill = X)
diapazon_label_selected_directory = Label(tab_Diapazon_Line1, text="Укажите папку:", width = 17)
diapazon_label_selected_directory.pack(side = LEFT)
diapazon_entry_selected_directory = Entry(master=tab_Diapazon_Line1,textvariable=diapazon_folder_path)
diapazon_entry_selected_directory.pack(side = LEFT, fill = BOTH, expand = True)
diapazon_button_selected_directory = Button(tab_Diapazon_Line1, text="Browse", command=diapazon_browse_button)
diapazon_button_selected_directory.pack(side = RIGHT)

tab_Diapazon_Line2.pack(side = TOP, fill = X)
diapazon_label_first_number = Label(tab_Diapazon_Line2, text="Первый номер:", width = 17)
diapazon_label_first_number.pack(side = LEFT)
diapazon_entry_first_number = Entry(master=tab_Diapazon_Line2,textvariable=diapazon_first_number)
diapazon_entry_first_number.pack(side = LEFT, fill = X, expand = True)

tab_Diapazon_Line3.pack(side = TOP, fill = X)
diapazon_label_last_number = Label(tab_Diapazon_Line3, text="Последний номер:", width = 17)
diapazon_label_last_number.pack(side = LEFT)
diapazon_entry_last_number = Entry(master=tab_Diapazon_Line3,textvariable=diapazon_last_number)
diapazon_entry_last_number.pack(side = LEFT, fill = X, expand = True)

tab_Diapazon_Line4.pack(side = TOP, fill = X)
diapazon_radiobutton_none = Radiobutton(master=tab_Diapazon_Line4, text="Без контрольной цифры", variable=diapazon_mode, value=1)
diapazon_radiobutton_none.pack(fill = X)

tab_Diapazon_Line5.pack(side = TOP, fill = X)
diapazon_radiobutton_code128 = Radiobutton(master=tab_Diapazon_Line5, text="Code 128 (контрольная цифра)", variable=diapazon_mode, value=2)
diapazon_radiobutton_code128.pack(fill = X)

tab_Diapazon_Line6.pack(side = TOP, fill = X)
diapazon_radiobutton_ean13 = Radiobutton(master=tab_Diapazon_Line6, text="EAN 13 (контрольная цифра)", variable=diapazon_mode, value=3)
diapazon_radiobutton_ean13.pack(fill = X)

tab_Diapazon_Line7.pack(side = TOP, fill = X)
diapazon_start_button = Button(tab_Diapazon_Line7, text="Старт", command=diapazon_accept_button)
diapazon_start_button.pack(fill = X)

tab_Diapazon_Line8.pack(side = TOP, fill = X)
diapazon_label_status = Label(tab_Diapazon_Line8, text=status.get(), width = 17)
diapazon_label_status.pack(fill = X)

#Тестирование диапазона
diapazon_entry_selected_directory.insert(0,"D:/1.xlsx")
diapazon_entry_first_number.insert(0, "000000315100333")
diapazon_entry_last_number.insert(0, "000000315100352")

#Киевстар
KS_path_to_database = StringVar()
KS_path_to_razbivka = StringVar()
KS_is_dual = IntVar()
KS_to_eps = IntVar()
KS_check_quantity = IntVar()
KS_activation_date = StringVar()
KS_material = StringVar()
KS_order_number = StringVar()
KS_data_otpravki = StringVar()
KS_order_name = StringVar()
KS_date_name = StringVar()
KS_contract_name = StringVar()
KS_list_of_data_entries = [KS_path_to_database, KS_path_to_razbivka, KS_activation_date, KS_material, KS_order_number, KS_data_otpravki,
								KS_order_name, KS_date_name, KS_contract_name, KS_is_dual, KS_to_eps, KS_check_quantity]

def KS_database_browse_button():
    global KS_path_to_database
    filename = filedialog.askopenfilename(title = "Укажите файл с базой данных",filetypes = (("TXT","*.txt"),))
    KS_path_to_database.set(filename)
def KS_razbivka_browse_button():
    global KS_path_to_razbivka
    filename = filedialog.askopenfilename(title = "Укажите файл с разбивкой",filetypes = (("XLSX","*.xlsx"),))
    KS_path_to_razbivka.set(filename)
def KS_accept_button():
	new_process = Thread(	target=processKS, args=[KS_path_to_database.get(), KS_path_to_razbivka.get(),
							KS_label_status, KS_is_dual.get() + 1, KS_to_eps.get(), KS_check_quantity.get(),
							KS_activation_date.get(), KS_material.get(), KS_order_number.get(), 
							KS_data_otpravki.get(), KS_order_name.get(), KS_date_name.get(), KS_contract_name.get()])
	#Загрузка параметров последнего запуска
	lines = []
	with open("settings.ini", "r") as file:
		lines = file.readlines()
		KS_setting_index = -1
		for entry_number in range(len(lines)):
			if "КС" in lines[entry_number]:
				KS_setting_index = entry_number
		if KS_setting_index != -1:
			for index in range(len(KS_list_of_data_entries) + 1):
				del lines[KS_setting_index]
	lines.append("Киевстар\n")
	for object in KS_list_of_data_entries:
		lines.append(str(object.get()) + "\n")
	with open("settings.ini", "w") as file:
		file.writelines(lines)
	new_process.start()

tab_KS_Line1.pack(side = TOP, fill = X)
KS_label_selected_file = Label(tab_KS_Line1, text="Выберите папку с БД:", width = 23)
KS_label_selected_file.pack(side = LEFT)
KS_entry_selected_file = Entry(master=tab_KS_Line1,textvariable=KS_path_to_database)
KS_entry_selected_file.pack(side = LEFT, expand = True, fill = BOTH)
KS_button_selected_file = Button(tab_KS_Line1, text="Browse", command=KS_database_browse_button)
KS_button_selected_file.pack(side = RIGHT)

tab_KS_Line2.pack(side = TOP, fill = X)
KS_label_selected_razbivka = Label(tab_KS_Line2, text="Выберите файл разбивки:", width = 23)
KS_label_selected_razbivka.pack(side = LEFT)
KS_entry_selected_razbivka = Entry(master=tab_KS_Line2,textvariable=KS_path_to_razbivka)
KS_entry_selected_razbivka.pack(side = LEFT, expand = True, fill = BOTH)
KS_button_selected_razbivka = Button(tab_KS_Line2, text="Browse", command=KS_razbivka_browse_button)
KS_button_selected_razbivka.pack(side = RIGHT)

tab_KS_Line3.pack(side = TOP, fill = X)
KS_checkbox_dual = Checkbutton(tab_KS_Line3, variable=KS_is_dual)
KS_checkbox_dual.pack(side = LEFT)
KS_checkbox_dual_label = Label(tab_KS_Line3, text="Дуальные карты", width = 23)
KS_checkbox_dual_label.pack(side = LEFT)
KS_checkbox_eps = Checkbutton(tab_KS_Line3, variable=KS_to_eps)
KS_checkbox_eps.pack(side = LEFT, expand = True, fill = BOTH)
KS_checkbox_eps_label = Label(tab_KS_Line3, text="ЭПС", width = 23)
KS_checkbox_eps_label.pack(side = LEFT, expand = True, fill = BOTH)
KS_checkbox_quantity_check_label = Label(tab_KS_Line3, text="Не сравнивать количество", width = 23)
KS_checkbox_quantity_check_label.pack(side = RIGHT)
KS_checkbox_quantity_check = Checkbutton(tab_KS_Line3, variable=KS_check_quantity)
KS_checkbox_quantity_check.pack(side = RIGHT)

tab_KS_Line4.pack(side = TOP, fill = X)
KS_label_activation_date = Label(tab_KS_Line4, text="Дата активации:", width = 23)
KS_label_activation_date.pack(side = LEFT)
KS_entry_activation_date = Entry(master=tab_KS_Line4,textvariable=KS_activation_date)
KS_entry_activation_date.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line4_1.pack(side = TOP, fill = X)
KS_label_material = Label(tab_KS_Line4_1, text="Тип товара:", width = 23)
KS_label_material.pack(side = LEFT)
KS_entry_material = Entry(master=tab_KS_Line4_1,textvariable=KS_material)
KS_entry_material.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line4_2.pack(side = TOP, fill = X)
KS_label_order_number = Label(tab_KS_Line4_2, text="Номер заказа:", width = 23)
KS_label_order_number.pack(side = LEFT)
KS_entry_order_number = Entry(master=tab_KS_Line4_2,textvariable=KS_order_number)
KS_entry_order_number.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line4_3.pack(side = TOP, fill = X)
KS_label_data_otpravki = Label(tab_KS_Line4_3, text="Дата отгрузки:", width = 23)
KS_label_data_otpravki.pack(side = LEFT)
KS_entry_data_otpravki = Entry(master=tab_KS_Line4_3,textvariable=KS_data_otpravki)
KS_entry_data_otpravki.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line4_4.pack(side = TOP, fill = X)
KS_label_order_name = Label(tab_KS_Line4_4, text="Номер заказа(для Киевстар):", width = 23)
KS_label_order_name.pack(side = LEFT)
KS_entry_order_name = Entry(master=tab_KS_Line4_4,textvariable=KS_order_name)
KS_entry_order_name.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line4_5.pack(side = TOP, fill = X)
KS_label_date_name = Label(tab_KS_Line4_5, text="Дата заказа(для Киевстар):", width = 23)
KS_label_date_name.pack(side = LEFT)
KS_entry_date_name = Entry(master=tab_KS_Line4_5,textvariable=KS_date_name)
KS_entry_date_name.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line4_6.pack(side = TOP, fill = X)
KS_label_contract_name = Label(tab_KS_Line4_6, text="Номер договора:", width = 23)
KS_label_contract_name.pack(side = LEFT)
KS_entry_contract_name = Entry(master=tab_KS_Line4_6,textvariable=KS_contract_name)
KS_entry_contract_name.pack(side = LEFT, expand = True, fill = X)

tab_KS_Line5.pack(side = TOP, fill = X)
KS_start_button = Button(tab_KS_Line5, text="Старт", command=KS_accept_button)
KS_start_button.pack(fill=X)

tab_KS_Line6.pack(side = TOP, fill = X)
KS_label_status = Label(tab_KS_Line6, text=status.get())
KS_label_status.pack()

#Тестирование Киевстара

try:
	with codecs.open("settings.ini", "r", "ansi") as file:
		lines = file.readlines()
		KS_setting_index = -1
		lyca_setting_index = -1
		for entry_number in range(len(lines)):
			if "КC" in lines[entry_number]:
				KS_setting_index = entry_number
			if "Лука" in lines[entry_number]:
				lyca_setting_index = entry_number
		if KS_setting_index != -1:
			index = KS_setting_index
			for object in KS_list_of_data_entries:
				index += 1
				if(KS_setting_index - index > 9):
					object.set(int(lines[index].rstrip()))
				else:
					object.set(lines[index].rstrip())
		if lyca_setting_index != -1:
			index = lyca_setting_index
			for object in lyca_list_of_data_entries:
				index += 1
				if(lyca_setting_index - index > 10):
					object.set(int(lines[index].rstrip()))
				else:
					object.set(lines[index].rstrip())
		
except Exception as e:
	print(e)

#Metropoliten tab
metropoliten_database_path = StringVar()
metropoliten_gromadski = IntVar()
metropoliten_studentski = IntVar()

def metropoliten_browse_button():
    filename = filedialog.askopenfilename()
    metropoliten_database_path.set(filename)
def metropoliten_start_button():
	metropoliten_status_label['text']="Статус: Обработка..."
	new_process = Thread( target=process_metropoliten, args = [metropoliten_database_path.get(), metropoliten_gromadski.get(), metropoliten_studentski.get(), metropoliten_status_label] )
	new_process.start()

tab_Metropoliten_Line1.pack(side = TOP, fill = X)
metropoliten_browse_label = Label(tab_Metropoliten_Line1, text="Выберите папку с БД:", width = 22)
metropoliten_browse_label.pack(side = LEFT)
metropoliten_browse_entry = Entry(master=tab_Metropoliten_Line1,textvariable=metropoliten_database_path)
metropoliten_browse_entry.pack(side = LEFT, expand = True, fill = BOTH)
metropoliten_browse_button = Button(tab_Metropoliten_Line1, text="Browse", command=metropoliten_browse_button)
metropoliten_browse_button.pack(side = RIGHT)

tab_Metropoliten_Line2.pack(side = TOP, fill = X)
metropoliten_gromadski_label = Label(tab_Metropoliten_Line2, width = 25)
metropoliten_gromadski_label.pack(side = LEFT)
metropoliten_gromadski_checkBox = Checkbutton(tab_Metropoliten_Line2, pady=5, text="Гражданские", variable=metropoliten_gromadski)
metropoliten_gromadski_checkBox.pack(side = LEFT)
metropoliten_studentski_label = Label(tab_Metropoliten_Line2, width = 25)
metropoliten_studentski_label.pack(side = RIGHT)
metropoliten_studentski_checkBox = Checkbutton(tab_Metropoliten_Line2, pady=5, text="Студенческие", variable=metropoliten_studentski)
metropoliten_studentski_checkBox.pack(side = RIGHT)

tab_Metropoliten_Line3.pack(side = TOP, fill = X)
metropoliten_start_button = Button(tab_Metropoliten_Line3, text="Старт", command=metropoliten_start_button)
metropoliten_start_button.pack(fill = X)

tab_Metropoliten_Line4.pack(side = TOP, fill = X)
metropoliten_status_label = Label(tab_Metropoliten_Line4)
metropoliten_status_label["text"] = "Готов к работе"
metropoliten_status_label.pack()

root.protocol("WM_DELETE_WINDOW", shutDown)

root.mainloop()