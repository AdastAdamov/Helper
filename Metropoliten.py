import pandas as pd
import re
import numpy as np

from tkinter import *
from Diapazon import code128

def search (file_zayavka, target, sheet_default = None):
	x, y, found = 0, 0, False
	for sheet in file_zayavka.sheet_names:
		database_zayavka = file_zayavka.parse(sheet)
		if sheet == "диапазон": continue
		if sheet_default != None and sheet != sheet_default: continue			
		x = 0
		for column in list(database_zayavka):
			y = 0
			for item in list(database_zayavka[column]):
				if str(item).find(target) is not -1:
					found = True
					break
				y += 1
			if found: break
			x += 1
		if found: break
	if found == False:
		x, y = 0, 0
	return x, y, found, database_zayavka

def process_metropoliten (path, gromadski, studentski, metropoliten_status_label):
	file_zayavka = pd.ExcelFile(path, dtype = str)
	file_settings = pd.ExcelFile("Metropoliten.xlsx", dtype = str)

	database = pd.DataFrame(dtype = str, columns = ["NUM","CODE128","VID1","VID2","VID3","VID4","VID5","DATE","PRICE2","Этикетки","IDD"])
	x, y, found, database_zayavka = search(file_zayavka, "Строк дії проїзних квитків")
	date = ""
	if found:	date = database_zayavka.iat[y, x+1]
	
	#Громадські
	x, y, found, database_zayavka = search(file_zayavka, "Громадські :")
	if found and gromadski == 1:
		print("Обработка гражданских...")
		metropoliten_status_label["text"] = "Обработка гражданских..."
		format_code = re.compile("\d+\s+\d+\s+\d+\s+\d+\w\s+\W\s\d+\s+\d+\s+\d+\s+\d+\w")
		result = format_code.search(database_zayavka.iat[y, x]).group()
		code_start, code_end = "", ""
		for letter in result:
			if len(code_start) < 15:
				if letter.isnumeric():
					code_start += letter
			else:
				if letter.isnumeric():
					code_end += letter
		counter = 0
		code_start = int(code_start)
		code_end = int(code_end)
		counter_max = code_end - code_start + 1
		x, y, found, database_zayavka = search(file_zayavka, "Дані для персоналізації проїзних квитків ГРОМАДСЬКИЙ:")
		y += 1
		x += 3
		index_start = 0
		while code_start <= code_end:
			y = y + 1
			type = database_zayavka.iat[y, x]
			if str(type) == "nan" or str(database_zayavka.iat[y, x+2]) == "nan": 
				if counter > counter_max:
					print("Ошибка в количестве записей. Обработано больше необходимого")
					metropoliten_status_label["text"] = "Ошибка в количестве записей. Обработано больше необходимого"
					break
				if y == 200:
					print("Обработано 200 строк исходника")
					metropoliten_status_label["text"] = "Обработано 200 строк исходника"
					break #ОСТАНОВКА СКРИПТА
				continue
			type = type.replace(" ","")
			x2, y2, found, database_settings = search(file_settings, type, "ГР")
			if found: 
				vid1 = database_settings.iat[y2, x2+1]
				vid2 = database_settings.iat[y2, x2+2]
				vid3 = database_settings.iat[y2, x2+3]
				vid4 = database_settings.iat[y2, x2+4]
				vid5 = database_settings.iat[y2, x2+5]
				etiketky = database_settings.iat[y2, x2+6]
			else:
				print("Незнакомое обозначение: " + type)
				metropoliten_status_label["text"] = "Незнакомое обозначение: " + type
				return
			index_max = index_start + int(database_zayavka.iat[y, x+2])
			index = index_start
			counter_of_inner_cards = 0
			while index < index_max:
				counter_of_inner_cards += 1
				number = code128(str(code_start)).zfill(16)
				database.at[index,"NUM"] =  number
				database.at[index,"CODE128"] =  number[0:4] + " " + number[4:8] + " " + number[8:12] + " " + number[12:16]
				database.at[index,"IDD"] = counter_of_inner_cards
				code_start += 1
				index += 1
			counter += counter_of_inner_cards
			database.loc[index_start : index_max-1, "VID1"] = vid1
			database.loc[index_start : index_max-1, "VID2"] = vid2
			database.loc[index_start : index_max-1, "VID3"] = vid3
			database.loc[index_start : index_max-1, "VID4"] = vid4
			database.loc[index_start : index_max-1, "VID5"] = vid5
			database.loc[index_start : index_max-1, "Этикетки"] = etiketky			
			database.loc[index_start : index_max-1, "PRICE2"] = str(database_zayavka.iat[y, x+1]) + " грн"
			database["DATE"] = date
			index_start = index_start + index_max
		if counter < counter_max:
			print("Ошибка в количестве записей. Обработано меньше необходимого.")
			metropoliten_status_label["text"] = "Ошибка в количестве записей. Обработано меньше необходимого."
			return
		database.to_excel("GR.xlsx", index=False, sheet_name='Sheet 1')
	#Студентські
	database = pd.DataFrame(dtype = str, columns = ["NUM","CODE128","VID1","VID2","VID3","VID4","VID5","DATE","PRICE2","Этикетки","IDD"])
	x, y, found, database_zayavka = search(file_zayavka, "Студентські :")
	if found and studentski == 1:
		print("Обработка студенческих...")
		metropoliten_status_label["text"] = "Обработка студенческих..."
		format_code = re.compile("\d+\s+\d+\s+\d+\s+\d+\w\s+\W\s\d+\s+\d+\s+\d+\s+\d+\w")
		result = format_code.search(database_zayavka.iat[y, x]).group()
		code_start, code_end = "", ""
		for letter in result:
			if len(code_start) < 15:
				if letter.isnumeric():
					code_start += letter
			else:
				if letter.isnumeric():
					code_end += letter
		counter = 0
		code_start = int(code_start)
		code_end = int(code_end)
		counter_max = code_end - code_start + 1
		x, y, found, database_zayavka = search(file_zayavka, "Дані для персоналізації проїзних квитків СТУДЕНТСЬКІ:")
		y += 1
		x += 3
		index_start = 0
		while code_start < code_end:
			y = y + 1
			type = database_zayavka.iat[y, x]
			if str(type) == "nan" or str(database_zayavka.iat[y, x+2]) == "nan": 
				if counter > counter_max:
					print("Ошибка в количестве записей")
					metropoliten_status_label["text"] = "Ошибка в количестве записей"
					break
				if y == 200:
					print("Обработано 200 строк исходника")
					metropoliten_status_label["text"] = "Обработано 200 строк исходника"
					break #ОСТАНОВКА СКРИПТА
				continue
			type = type.replace(" ","")
			x2, y2, found, database_settings = search(file_settings, type, "СТ")
			if found: 
				vid1 = database_settings.iat[y2, x2+1]
				vid2 = database_settings.iat[y2, x2+2]
				vid3 = database_settings.iat[y2, x2+3]
				vid4 = database_settings.iat[y2, x2+4]
				vid5 = database_settings.iat[y2, x2+5]
				etiketky = database_settings.iat[y2, x2+6]
			else:
				print("Незнакомое обозначение: " + type)
				metropoliten_status_label["text"] = "Незнакомое обозначение: " + type
				return
			index_max = index_start + int(database_zayavka.iat[y, x+2])
			index = index_start
			counter_of_inner_cards = 0
			while index < index_max:
				counter_of_inner_cards += 1
				number = code128(str(code_start)).zfill(16)
				database.at[index,"NUM"] =  number
				database.at[index,"CODE128"] =  number[0:4] + " " + number[4:8] + " " + number[8:12] + " " + number[12:16]
				database.at[index,"IDD"] = counter_of_inner_cards
				code_start += 1
				index += 1
			counter += counter_of_inner_cards
			database.loc[index_start : index_max-1, "VID1"] = vid1
			database.loc[index_start : index_max-1, "VID2"] = vid2
			database.loc[index_start : index_max-1, "VID3"] = vid3
			database.loc[index_start : index_max-1, "VID4"] = vid4
			database.loc[index_start : index_max-1, "VID5"] = vid5
			database.loc[index_start : index_max-1, "Этикетки"] = etiketky
			database.loc[index_start : index_max-1, "PRICE2"] = str(database_zayavka.iat[y, x+1]) + " грн"
			database["DATE"] = date
			index_start = index_start + index_max
		if counter < counter_max:
			print("Ошибка в количестве записей")
			metropoliten_status_label["text"] = "Ошибка в количестве записей"
			return
		database.to_excel("ST.xlsx", index=False, sheet_name='Sheet 1')
	metropoliten_status_label["text"] = "Готово!"

if __name__ == "__main__":	
	path = "D:/! Helper/Метрополітен_Проезд_80_16500(78)+/Исходники/ЗАЯВКА №10.xls"
	process_metropoliten(path, 0, 1)

'''
Механизм поиска
for sheet in file_zayavka.sheet_names:
	database_zayavka = file_zayavka.parse(z)
	x = 0
	for column in list(database_zayavka):
		y = 0
		for item in list(database_zayavka[column]):
			if str(item) == "Де Х – контрольна цифра, що генерується Постачальником":
				print(z, x, y, item)
				found = True
				break
			y += 1
		if found: break
		x += 1
	if found: break
	z += 1
print(database_zayavka.iat[y, x])
'''