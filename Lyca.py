import pandas as pd
import numpy as np
import os
import xlsxwriter
from tkinter import *

def processLyca (path, orderName, jobNumber, itf, articul, zakaz, faceValue, innerCode, outerCode, customer, label30x20, lyca_label_status):
	dfs = []
	pathDatabase = path + "БД/"
	#Общая обработка в один файл
	lyca_label_status["text"] = "Обработка исходников..."
	print("Обработка исходников...")
	for filename in os.listdir(pathDatabase):
		print(filename)
		data = pd.read_excel(pathDatabase+filename, dtype = str)
		gg = pd.ExcelFile(pathDatabase+filename)
		data['sheet_name'] = gg.sheet_names[0]
		dfs.append(data)
	result = pd.concat(dfs)

	outputFileName = path + path.split("/")[-2] + ".xlsx"
	result.to_excel(outputFileName, index=False, header=False)
	data = pd.read_excel(outputFileName, header=None, dtype = str)
	#Обработка этикеток
	lyca_label_status["text"] = "Обработка этикеток..."
	print("Обработка этикеток...")
	boxDatabase = pd.DataFrame(dtype = str, columns = ["ICCID1", "ICCID2", "BATCH", "NAME", "JN", "ITF", "BOX1", "BOX2", "QTTY", "ORDER"])
	bigBoxDatabase = pd.DataFrame(dtype = str, columns = ["ICCID1", "ICCID2", "BATCH", "NAME", "JN", "ITF", "BOX1", "BOX2", "QTTY", "ORDER"])
	a4Database = pd.DataFrame(dtype = str, columns = ["ICCID1", "ICCID2", "BATCH", "NAME", "JN", "ITF", "ART", "BOX", "QTTY", "ZAKAZ", "ORDER"])
	index1, index2 = 0, 1
	while index1 < data[0].count():
		boxDatabase.at[index2,'ICCID1'] = data.at[index1, 0]
		boxDatabase.at[index2,'ICCID2'] = data.at[index1 + 99, 0]
		boxDatabase.at[index2,'BATCH'] = data.at[index1, 6]
		boxDatabase.at[index2,'NAME'] = orderName
		boxDatabase.at[index2,'JN'] = jobNumber
		boxDatabase.at[index2,'ITF'] = itf
		boxDatabase.at[index2,'BOX1'] = str(index2)
		boxDatabase.at[index2,'BOX2'] = str(data[0].count() // 100)
		boxDatabase.at[index2,'QTTY'] = "100"
		boxDatabase.at[index2,'ORDER'] = str(index2).zfill(5)
		index1 += 100
		index2 += 1
	index1, index2 = 0, 1
	while index1 < data[0].count():
			bigBoxDatabase.at[index2,'ICCID1'] = data.at[index1, 0]
			bigBoxDatabase.at[index2,'ICCID2'] = data.at[index1 + 499, 0]
			bigBoxDatabase.at[index2,'BATCH'] = data.at[index1, 6]
			bigBoxDatabase.at[index2,'NAME'] = orderName
			bigBoxDatabase.at[index2,'JN'] = jobNumber
			bigBoxDatabase.at[index2,'ITF'] = itf
			bigBoxDatabase.at[index2,'BOX1'] = str(index2)
			bigBoxDatabase.at[index2,'BOX2'] = str(data[0].count() // 500)
			bigBoxDatabase.at[index2,'QTTY'] = "500"
			bigBoxDatabase.at[index2,'ORDER'] = str(index2).zfill(5)
			index1 += 500
			index2 += 1
	index1, index2 = 0, 1
	while index1 < data[0].count():
		a4Database.at[index2,'ICCID1'] = data.at[index1, 0]
		a4Database.at[index2,'ICCID2'] = data.at[index1 + 9999, 0]
		a4Database.at[index2,'BATCH'] = data.at[index1, 6]
		a4Database.at[index2,'NAME'] = orderName
		a4Database.at[index2,'JN'] = jobNumber
		a4Database.at[index2,'ITF'] = itf
		a4Database.at[index2,'ART'] = articul
		a4Database.at[index2,'BOX'] = str(index2)
		a4Database.at[index2,'QTTY'] = "10000"
		a4Database.at[index2,'ORDER'] = str(index2).zfill(5)
		a4Database.at[index2,'ZAKAZ'] = zakaz
		index1 += 10000
		index2 += 1
	try: os.makedirs(path + "Этикетка")
	except OSError as e: 
		if e.errno != errno.EEXIST:
			raise
	boxDatabase.to_excel(path + "Этикетка/80x60_Kor.xlsx", index=False, sheet_name='Лист1')
	bigBoxDatabase.to_excel(path + "Этикетка/80x60_Jashch.xlsx", index=False, sheet_name='Лист1')
	a4Database.to_excel(path + "Этикетка/A4.xlsx", index=False, sheet_name='Лист1')
	#База для Спеклера
	lyca_label_status["text"] = "Создание базы для Спеклера..."
	print("Создание базы для Спеклера...")
	speklerDatabase = pd.DataFrame(dtype = str)
	speklerDatabase[0] = data[0]
	index1 = 0
	while index1 < data[0].count():
		speklerDatabase.loc[index1 : index1+99, 1] = data.at[index1 + 99, 0]
		index1 += 100
	index1 = 0
	while index1 < data[0].count():
		speklerDatabase.loc[index1 : index1+499, 2] = data.at[index1 + 499, 0]
		index1 += 500
	index1 = 0
	while index1 < data[0].count():
		speklerDatabase.loc[index1 : index1+9999, 3] = data.at[index1, 6]
		index1 += 10000
	np.savetxt(path + path.split("/")[-2] + ".txt", speklerDatabase.values, fmt="%s", delimiter=';', newline='\r\n')

	#Этикетка сендвич
	if label30x20 == 1:
		lyca_label_status["text"] = "Создание базы этикетки-сэндвич..."
		print("Создание базы этикетки-сэндвич...")
		label30x20Database = pd.DataFrame(dtype = str, columns = ["CODE128", "PUK", "BATCH", "ORDER", "PACK"])
		label30x20Database["CODE128"] = data[0]
		label30x20Database["PUK"] = data[3]
		index1 = 0
		while index1 < data[0].count():
			label30x20Database.at[index1, "ORDER"] = str(index1 + 1).zfill(5)
			index1 += 1
		index1 = 0
		while index1 < data[0].count():
			label30x20Database.loc[index1 : index1+9999, "BATCH"] = data.at[index1, 6]
			index1 += 10000
		index1, index2 = 0, 1
		while index1 < data[0].count():
			label30x20Database.loc[index1 : index1+999, "PACK"] = str(index2).zfill(4)
			index1 += 1000
			index2 += 1
		label30x20Database.to_excel(path + "Этикетка/" + zakaz + "_30x20.xlsx", index=False, sheet_name='Лист1')
	#Упаковочная БД
	lyca_label_status["text"] = "Создание упаковочного листа и БД..."	
	print("Создание упаковочной листа и БД...")
	packingDatabase = pd.DataFrame(dtype = str, columns = ["ICCID", "PIN1", "PUK1", "PIN2", "PUK2", "BATCH", "BOX NO", "BIG BOX NO", "VALUE", "INNER TRIO LABEL GTIN", "OUTER TRIO LABEL GTIN"])
	packingDatabase["ICCID"] = data[0]
	packingDatabase["PIN1"] = data[2]
	packingDatabase["PUK1"] = data[3]
	packingDatabase["PIN2"] = data[4]
	packingDatabase["PUK2"] = data[5]
	packingDatabase["VALUE"] = faceValue
	packingDatabase["INNER TRIO LABEL GTIN"] = innerCode
	packingDatabase["OUTER TRIO LABEL GTIN"] = outerCode
	index1 = 0
	while index1 < data[0].count():
		packingDatabase.loc[index1 : index1+9999, "BATCH"] = data.at[index1, 6]
		index1 += 10000
	index1, index2 = 0, 1
	while index1 < data[0].count():
		packingDatabase.loc[index1 : index1+99, "BOX NO"] = str(index2)
		index1 += 100
		index2 +=1
	index1, index2 = 0, 1
	while index1 < data[0].count():
		packingDatabase.loc[index1 : index1+499, "BIG BOX NO"] = str(index2)
		index1 += 500
		index2 +=1
	print("OK 2")
	#packingDatabase.to_excel(path + "Упаковочный лист/Database_" + jobNumber + "_LM_" + faceValue + "_" + str(data[0].count()) + ".xlsx", index=False)
	#Украшательство
	try: os.makedirs(path + "Упаковочный лист")
	except OSError as e: 
		if e.errno != errno.EEXIST:
			raise
	writer = pd.ExcelWriter(path + "Упаковочный лист/Database_" + jobNumber + "_LM_" + faceValue + "_" + str(data[0].count()) + ".xlsx", engine='xlsxwriter')
	packingDatabase.to_excel(writer, sheet_name='Database', index=False)
	print("OK 3")
	workbook = writer.book
	worksheet = writer.sheets['Database']
	worksheet.set_column(0, 0, 19.57)
	worksheet.set_column(1, 1, 4.43)
	worksheet.set_column(2, 2, 8.29)
	worksheet.set_column(3, 3, 4.43)
	worksheet.set_column(4, 4, 8.29)
	worksheet.set_column(5, 5, 8.86)
	worksheet.set_column(6, 6, 7.29)
	worksheet.set_column(7, 7, 10.86)
	worksheet.set_column(8, 8, 9.86)
	worksheet.set_column(9, 9, 21.29)
	worksheet.set_column(10, 10, 21.71)
	cellFormat = workbook.add_format({"font_name":"Arial", "font_size":"10"})
	workbook.formats[0].set_font_name("Arial")
	workbook.formats[0].set_font_size(10)
	headerFormat = workbook.add_format({"bg_color":"yellow"})
	worksheet.set_row(0, 15.0)
	worksheet.conditional_format(0,0,0,10,{'type':'cell', 'criteria': '!=', 'value':'0', 'format':headerFormat})
	print("OK 4")
	writer.save()
	print("OK 5")

	#Упаковочный лист
	workbook = xlsxwriter.Workbook(path + "Упаковочный лист/Packing list_" + jobNumber + "_LM_" + faceValue + "_" + str(data[0].count()) + ".xlsx")
	worksheetPackingList = workbook.add_worksheet('Packing list')
	worksheetPackingList.set_column(0, 0, 7.29)
	worksheetPackingList.set_column(1, 1, 12.71)
	worksheetPackingList.set_column(2, 3, 19.57)
	worksheetPackingList.set_column(4, 4, 10.29)
	worksheetPackingList.set_column(5, 5, 8.86)
	worksheetPackingList.set_column(6, 6, 22.57)
	worksheetPackingList.set_column(7, 7, 22.43)
	headerFormat = workbook.add_format({"bold":True, "font_name":"Antique Olive", "font_size":"16", "italic":True, "font_color":"red", "align": "center"})
	fieldNameFormat = workbook.add_format({"font_name":"Antique Olive", "font_size":"16", "font_color":"red"})
	fieldValueFormat = workbook.add_format({"font_name":"Antique Olive", "font_size":"16", "font_color":"red", "bold":True})
	fieldNameFormat2 = workbook.add_format({"font_name":"Antique Olive", "font_size":"10", "bold":True, "border":2,'text_wrap':True, "align": "center", "valign": "vcenter"})
	fieldValueFormat2 = workbook.add_format({"font_name":"Antique Olive", "font_size":"10", "bold":True, 'text_wrap':True, "align": "center", "right":1, "bottom":1})
	fieldValueFormat2Red = workbook.add_format({"font_name":"Antique Olive", "font_size":"10", "bold":True, 'text_wrap':True, "align": "center", "font_color":"red", "right":1, "bottom":1})
	fieldValueFormat2RedLeftBorder = workbook.add_format({"font_name":"Antique Olive", "font_size":"10", "bold":True, 'text_wrap':True, "align": "center", "font_color":"red", "left":2, "right":1, "bottom":1})
	fieldValueFormat2RedRightBorder = workbook.add_format({"font_name":"Antique Olive", "font_size":"10", "bold":True, 'text_wrap':True, "align": "center", "font_color":"red", "right":2, "left":1, "bottom":1})
	borderFormat = workbook.add_format({"border":2})
	borderFormat2 = workbook.add_format({"top":2})
	borderFormat3 = workbook.add_format({"border":1})
	verticalAlignFormat = workbook.add_format({"valign": "vcenter"})

	worksheetPackingList.merge_range(0,1,0,2, "Packing List", headerFormat)
	worksheetPackingList.merge_range(1,1,1,2, "Date:", fieldNameFormat)
	worksheetPackingList.merge_range(2,1,2,2, "Customer:", fieldNameFormat)
	worksheetPackingList.write(3,1, "Supplier:", fieldNameFormat)
	worksheetPackingList.merge_range(4,1,4,2, "Product:", fieldNameFormat)
	worksheetPackingList.write(5,1, "PO:", fieldNameFormat)
	worksheetPackingList.write(6,1, "Quantity:", fieldNameFormat)

	worksheetPackingList.merge_range(1,3,1,8, "26.07.2018", fieldValueFormat) # Фиксированная дата
	worksheetPackingList.merge_range(2,3,2,8, customer, fieldValueFormat)
	worksheetPackingList.merge_range(3,3,3,8, "SPEKL Ltd., 35, Svitlitskogo Str., Kyiv 04123, Ukraine", fieldValueFormat)
	worksheetPackingList.merge_range(4,3,4,8, "lyca " + faceValue, fieldValueFormat)
	worksheetPackingList.merge_range(5,3,5,8, jobNumber, fieldValueFormat)
	worksheetPackingList.merge_range(6,3,6,8, str(data[0].count()), fieldValueFormat)

	worksheetPackingList.write(7,0, "Pallet No.", fieldNameFormat2)
	worksheetPackingList.write(7,1, "Batch No.", fieldNameFormat2)
	worksheetPackingList.write(7,2, "ICCID-Start", fieldNameFormat2)
	worksheetPackingList.write(7,3, "ICCID – End", fieldNameFormat2)
	worksheetPackingList.write(7,4, "Face value", fieldNameFormat2)
	worksheetPackingList.write(7,5, "QTY", fieldNameFormat2)

	index1, index2 = 0, 1

	while index1 < data[0].count():
		worksheetPackingList.write(index2+7,0, str(index2), fieldValueFormat2RedLeftBorder)
		worksheetPackingList.write(index2+7,1, data.at[index1, 6], fieldValueFormat2Red)
		worksheetPackingList.write(index2+7,2, data.at[index1, 0], fieldValueFormat2)
		worksheetPackingList.write(index2+7,3, data.at[index1+9999, 0], fieldValueFormat2)
		worksheetPackingList.write(index2+7,4, faceValue, fieldValueFormat2)
		worksheetPackingList.write(index2+7,5, "10000", fieldValueFormat2RedRightBorder)
		worksheetPackingList.set_row(index2+7, 12.75, verticalAlignFormat)
		index2 += 1
		index1 += 10000

	worksheetPackingList.conditional_format(index2+7,0,index2+7,5,{'type':'cell', 'criteria': '=', 'value':'0', 'format':borderFormat2})
	workbook.close()

	#Файл для склада
	name = zakaz + "_lyca_" + faceValue + "_" + str(data[0].count())
	path_to_store_directory = "Y:/Manufacture/Склад/lyca/" + name
	try: 
		os.makedirs(path_to_store_directory)
		print("Обработка файла Excel для складов")
		lyca_label_status["text"] = "Обработка файла Excel для складов..."	
		database_for_warehouse = pd.DataFrame(dtype=str, columns = ["Batch", "ICCID_from", "ICCID_to", "Quantity", "Name", "Job Number", "Status"])
		database_for_warehouse["Batch"] = a4Database["BATCH"]
		database_for_warehouse["ICCID_from"] = a4Database["ICCID1"]
		database_for_warehouse["ICCID_to"] = a4Database["ICCID2"]
		database_for_warehouse["Quantity"] = a4Database["QTTY"]
		database_for_warehouse["Name"] = a4Database["NAME"]
		database_for_warehouse["Job Number"] = a4Database["JN"]
		writer = pd.ExcelWriter(path_to_store_directory + "/" + name + ".xlsx", engine = 'xlsxwriter')
		database_for_warehouse.to_excel(writer, sheet_name='Разбивка', index=False)
	except OSError as e: 
		print("Файл для склада уже существует")
		if e.errno != errno.EEXIST:
			raise
	print("ГОТОВО!")

	lyca_label_status["text"] = "Готово!"	