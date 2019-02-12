import pandas as pd
import numpy as np
import os
import errno
import math
import xlsxwriter
from Diapazon import code128
from tkinter import *
import datetime

def processKS(path_to_file, path_to_razbivka, status, dual, eps, no_check_batch_quantity, activation_date, material, 
					order_number, data_otpravki, order_name, order_date, contract_name):
	status["text"]="Обработка txt-файла..."
	if material[0] != '"':	material = '"' + material
	if material[-1] != '"':	material += '"'
	material_name = material.replace("/","").replace('"',"")
	material_name2 = material.replace("/","_").replace('"',"")
	database_iccid = []
	database_msisdn = []
	database_excel = pd.DataFrame(dtype = str, columns = ["ICCID","MSISDN","BATCH"])
	file_razbivka = pd.read_excel(path_to_razbivka, dtype = str)
	quantity_of_cards_in_total = 0
	entry_name_quantity_of_cards_per_batch = ""
	entry_name_quantity_of_cards_per_order = ""
	try:
		file_razbivka.at[0, "факт кть"]
		entry_name_quantity_of_cards_per_batch = "факт кть"
	except:
		try:
			file_razbivka.at[0, "кть сім-карток"]
			entry_name_quantity_of_cards_per_batch = "кть сім-карток"
		except:
			print("Нестандартная шапка в файле разбивки")
			status["text"]="Нестандартная шапка в файле разбивки"
			return
	try:
		file_razbivka.at[0, "СПкть"]
		entry_name_quantity_of_cards_per_order = "СПкть"
	except:
		try:
			file_razbivka.at[0, "кть сім-карток"]
			entry_name_quantity_of_cards_per_order = "кть сім-карток"
		except:
			print("Нестандартная шапка в файле разбивки")
			status["text"]="Нестандартная шапка в файле разбивки"
			return
	batch_quantity = {}

	with open(path_to_file) as f:
		index = 1
		temp = []
		for line in f:
			if index < 3:
				index += 1
				continue
			temp = line.split()
			if line[0] == "-":
				break
			database_iccid.append(temp[1])
			database_msisdn.append(temp[2])
			index += 1
			quantity_of_cards_in_total += 1

		database_excel["ICCID"] = database_iccid
		database_excel["MSISDN"] = database_msisdn
		database_excel.sort_values(by = ["ICCID"], inplace = True)
		database_excel.reset_index(drop=True, inplace = True)
		
		batch_name = None
		index, quantity, end_is_found = 0, 0, 0
		status["text"]="Сравнение с файлом разбивки..."
		for entry in range(database_excel["ICCID"].count()):
			if database_excel.at[entry, "ICCID"] == file_razbivka.at[index, "ІСС початок"]:
				print("Найден первый штрих-код:", database_excel.at[entry, "ICCID"])
				batch_name = file_razbivka.at[index, "Батч"]
				if no_check_batch_quantity == 1: batch_quantity.update({batch_name : 0})
				end_is_found = 0
			if batch_name is not None:
				database_excel.at[entry, "BATCH"] = batch_name
				if no_check_batch_quantity == 1: batch_quantity[batch_name] += 1
				database_excel.at[entry, "ICCID"] = code128(database_excel.at[entry, "ICCID"]) + "F"
				
			if database_excel.at[entry, "ICCID"][:-2] == file_razbivka.at[index, "ІСС кінець"]:
				print("Найден второй штрих-код:", database_excel.at[entry, "ICCID"][:-2])
				batch_name = None
				end_is_found = 1
				index += 1			
				if no_check_batch_quantity == 1: continue
				quantity += int(file_razbivka.at[index-1, entry_name_quantity_of_cards_per_batch].replace(" ",""))
				if entry + 1 == quantity:
					continue
				else:
					print("Ошибка! Количество карт не соотвествует указанному в разбивке:", entry + 1, quantity)
					return
			if batch_name is None:
				print("Не найден первый штрих-код:", file_razbivka.at[index, "ІСС початок"])
				status["text"]="Ошибка!"
				return
		if end_is_found == 0:
			print("Не найден второй штрих-код:", file_razbivka.at[index, "ІСС кінець"])
			status["text"]="Ошибка!"
			return
		if quantity_of_cards_in_total != int(file_razbivka.at[0, entry_name_quantity_of_cards_per_order].replace(" ","")) * dual: #0.04
			status["text"]="Количество карт в файле с разбивкой не сходится с количеством в базе!"
			print("Количество карт в файле с разбивкой не сходится с количеством в базе.")
			return
	database_excel_for_client = None
	#Удаление чётных рядов для дуальных карт
	if dual == 2:
		#Удаление нечётных рядов для второй БД...
		'''
		status["text"]="Удаление нечётных рядов для второй БД..."
		database_excel_for_client = database_excel.copy()
		database_excel_for_client.drop(database_excel_for_client.index[0::2], inplace = True)
		database_excel_for_client.reset_index(drop=True, inplace = True)
		'''
		status["text"]="Удаление чётных рядов для дуальных карт..."
		database_excel.drop(database_excel.index[1::2], inplace = True)
		database_excel.reset_index(drop=True, inplace = True)	
		quantity_of_cards_in_total = quantity_of_cards_in_total // 2 #0.04

	#Создание базы для ЭПС
	if eps == 1:
		database_excel["Material"] = material
		database_excel["Date"] = activation_date
		status["text"]="Запись обработанной базы в отдельный файл..."
		database_excel.to_excel(os.path.dirname(path_to_razbivka) + "/" + order_number + "_" + material_name + "_" + str(quantity_of_cards_in_total) + ".xlsx", index=False)
		print("Готово!")
		status["text"]="Готово!"
		return
	#Дополнение коробок картами с конца базы...
	status["text"]="Дополнение коробок картами с конца базы..."
	list_of_numbers = []
	index_start = 1
	index_last = database_excel["ICCID"].count()
	while index_start <= index_last:
		list_of_numbers.append(index_start)
		index_start += 1

	index_start = 1
	quantity = index_last
	shift_number = 0
	shift_size = 0
	box_counter = 1
	while index_start < index_last:
		if shift_number == 0 and database_excel.at[index_start, "BATCH"] != database_excel.at[index_start - 1, "BATCH"]:
			shift_number = 1
			memory = index_start
			shift_size = 100 - box_counter % 100
			index_last -= shift_size
		if shift_number == 1:
			list_of_numbers[index_start] = list_of_numbers[index_start] + quantity
		index_start += 1
		box_counter += 1
		if index_start == index_last and shift_number == 1:
			index_start = memory + 1
			box_counter = 1
			shift_number = 0
	database_excel["ORDER"] = pd.Series(list_of_numbers, dtype=int)
	database_excel.sort_values(by = ["ORDER"], inplace = True)
	database_excel.reset_index(drop=True, inplace = True)
	database_excel.drop("ORDER", axis=1, inplace = True)
	
	#Дополнение коробок картами с конца базы второй БД...
	'''
	if dual == 2:
		status["text"]="Дополнение коробок картами с конца базы второй БД..."
		list_of_numbers = []
		index_start = 1
		index_last = database_excel_for_client["ICCID"].count()
		while index_start <= index_last:
			list_of_numbers.append(index_start)
			index_start += 1

		index_start = 1
		quantity = index_last
		shift_number = 0
		shift_size = 0
		box_counter = 1
		while index_start < index_last:
			if shift_number == 0 and database_excel_for_client.at[index_start, "BATCH"] != database_excel_for_client.at[index_start - 1, "BATCH"]:
				shift_number = 1
				memory = index_start
				shift_size = 100 - box_counter % 100
				index_last -= shift_size
			if shift_number == 1:
				list_of_numbers[index_start] = list_of_numbers[index_start] + quantity
			index_start += 1
			box_counter += 1
			if index_start == index_last and shift_number == 1:
				index_start = memory + 1
				box_counter = 1
				shift_number = 0		
		database_excel_for_client["ORDER"] = pd.Series(list_of_numbers, dtype=int)
		database_excel_for_client.sort_values(by = ["ORDER"], inplace = True)
		database_excel_for_client.reset_index(drop=True, inplace = True)
		database_excel_for_client.drop("ORDER", axis=1, inplace = True)
	'''
	status["text"]="Запись обработанной базы в формате Excel..."
	database_excel.to_excel(os.path.dirname(path_to_razbivka) + "/" + order_number + "_" + material_name + "_" + str(quantity_of_cards_in_total) + ".xlsx", index=False)
	
	#База на этикетки
	status["text"]="Создание этикетки 100х60..."
	database_excel_100x60 = None #1.04
	if dual == 1:
		database_excel_100x60 = pd.DataFrame(dtype=str, columns=[	"ICCID_from","ICCID_to","MSISDN_from","MSISDN_to","Batch","Pallet",
																	"Date", "NUMBOX", "Quant", "Material", "orders"])
	else:
		database_excel_100x60 = pd.DataFrame(dtype=str, columns=[	"ICCID_from","ICCID_to","MSISDN_from","MSISDN_to","Batch","Pallet",
																	"Date", "NUMBOX", "Quant", "SIM_Quantity", "Material", "orders"])		
	index1, index2 = 0, 0
	index_last = database_excel["ICCID"].count()
	while index1 < index_last - 100:
			database_excel_100x60.at[index2, "ICCID_from"] = database_excel.at[index1, "ICCID"]
			database_excel_100x60.at[index2+1, "ICCID_from"] = database_excel.at[index1, "ICCID"]
			database_excel_100x60.at[index2, "ICCID_to"] = database_excel.at[index1+99, "ICCID"]
			database_excel_100x60.at[index2+1, "ICCID_to"] = database_excel.at[index1+99, "ICCID"]
			database_excel_100x60.at[index2, "MSISDN_from"] = database_excel.at[index1, "MSISDN"]
			database_excel_100x60.at[index2+1, "MSISDN_from"] = database_excel.at[index1, "MSISDN"]
			database_excel_100x60.at[index2, "MSISDN_to"] = database_excel.at[index1+99, "MSISDN"]
			database_excel_100x60.at[index2+1, "MSISDN_to"] = database_excel.at[index1+99, "MSISDN"]
			database_excel_100x60.at[index2, "NUMBOX"] = "B" + order_number[-2:] + str(index2 // 2 + 1).zfill(5)
			database_excel_100x60.at[index2+1, "NUMBOX"] = "C" + order_number[-2:] + str(index2 // 2 + 1).zfill(5)
			database_excel_100x60.at[index2, "orders"] = str(index2 + 1).zfill(5)
			database_excel_100x60.at[index2+1, "orders"] = str(index2 + 2).zfill(5)
			index1 += 100
			index2 += 2
	database_excel_100x60.at[index2, "ICCID_from"] = database_excel.at[index1, "ICCID"]
	database_excel_100x60.at[index2+1, "ICCID_from"] = database_excel.at[index1, "ICCID"]
	database_excel_100x60.at[index2, "ICCID_to"] = database_excel.at[index1+(index_last % 100 -1), "ICCID"]
	database_excel_100x60.at[index2+1, "ICCID_to"] = database_excel.at[index1+(index_last % 100 -1), "ICCID"]
	database_excel_100x60.at[index2, "MSISDN_from"] = database_excel.at[index1, "MSISDN"]
	database_excel_100x60.at[index2+1, "MSISDN_from"] = database_excel.at[index1, "MSISDN"]
	database_excel_100x60.at[index2, "MSISDN_to"] = database_excel.at[index1+(index_last % 100 -1), "MSISDN"]
	database_excel_100x60.at[index2+1, "MSISDN_to"] = database_excel.at[index1+(index_last % 100 -1), "MSISDN"]
	database_excel_100x60.loc[0 : index2 - 1, "Quant"] = 100
	database_excel_100x60.loc[index2 : index2 + 1, "Quant"] = index_last % 100
	if dual == 2:
		database_excel_100x60.loc[0 : index2 - 1, "SIM_Quantity"] = "200"
		database_excel_100x60.loc[index2 : index2 + 1, "SIM_Quantity"] = str(index_last % 100 * 2)
	database_excel_100x60.loc[0 : index2 + 1, "Date"] = activation_date
	database_excel_100x60.loc[0 : index2 + 1, "Material"] = material
	database_excel_100x60.at[index2, "NUMBOX"] = "B" + order_number[-2:] + str(index2 // 2 + 1).zfill(5)
	database_excel_100x60.at[index2+1, "NUMBOX"] = "C" + order_number[-2:] + str(index2 // 2 + 1).zfill(5)
	database_excel_100x60.at[index2, "orders"] = str(index2 + 1).zfill(5)
	database_excel_100x60.at[index2+1, "orders"] = str(index2 + 2).zfill(5)

	index_batch, index1, double_label = 0, 0, 0
	#Поле батч в этикетке 100х60
	for index in range( database_excel_100x60["ICCID_from"].count() ):
		batch = file_razbivka.at[index_batch, "Батч"]
		if (index // 2) * 100 + 99 >= index_last:
			database_excel_100x60.at[index, "Batch"] = batch
		elif database_excel.at[(index // 2) * 100, "BATCH"]  == database_excel.at[(index // 2) * 100 + 99, "BATCH"]:
			database_excel_100x60.at[index, "Batch"] = batch
		else:
			database_excel_100x60.at[index, "Batch"] = batch + "," + database_excel.at[(index // 2) * 100 + 99, "BATCH"]
			if double_label == 0:
				double_label = 1
				continue
			else:
				double_label = 0
				index_batch += 1
	#Поле номер коробки в этикетке 100х60
	shift = 0
	current_batch_number = 0
	current_batch_quantity = 0
	if no_check_batch_quantity == 0:
		current_batch_quantity = math.ceil(int(file_razbivka.at[current_batch_number, entry_name_quantity_of_cards_per_batch].replace(" ","")) / 100)
	else:
		current_batch_quantity = math.ceil(int(batch_quantity.get(file_razbivka.at[current_batch_number, "Батч"])) / 100)
	batch_end_found = False
	for index in range( database_excel_100x60["ICCID_from"].count() ):
		database_excel_100x60.at[index, "Pallet"] = str(index // 2 + 1 - shift) + "/" + str(current_batch_quantity // dual)
		if index / 2 + 1 - shift == current_batch_quantity // dual:
			batch_end_found = True
			continue
		if batch_end_found:
			shift += current_batch_quantity // dual
			current_batch_number += 1
			if current_batch_number == file_razbivka["Батч"].count():
				break
			if no_check_batch_quantity == 0:
				current_batch_quantity = math.ceil(int(file_razbivka.at[current_batch_number, entry_name_quantity_of_cards_per_batch]) / 100)
			else:
				current_batch_quantity = math.ceil(int(batch_quantity.get(file_razbivka.at[current_batch_number, "Батч"])) / 100)
			batch_end_found = False

	#Создание базы на этикетку А4
	status["text"]="Создание этикетки А4..."
	database_excel_A4 = pd.DataFrame(dtype=str, columns = [	"Zakaz", "Otpr","Material1","Date" , "Pallet", "ICCID_from", "ICCID_to", 
															"MSISDN_from", "MSISDN_to", "Batch", "Quant", "NUMBOX", "NUMBOX1", "NUMBOX2"])
	if dual == 2:
		database_excel_A4 = pd.DataFrame(dtype=str, columns = [	"Zakaz", "Otpr","Material1","Date" , "Pallet", "ICCID_from", "ICCID_to", 
															"MSISDN_from", "MSISDN_to", "Batch", "Quant", "NUMBOX", "NUMBOX1", "NUMBOX2", "SIM_Quantity"])
	index_label, counter = 0, 1
	index_label_end = database_excel["ICCID"].count() // 10000
	leftover = database_excel["ICCID"].count() % 10000
	while index_label < index_label_end:
		database_excel_A4.at[index_label, "Pallet"] = str(index_label+1) + " of " + str(index_label_end + 1)
		if dual == 2: database_excel_A4.at[index_label, "SIM_Quantity"] = 20000
		database_excel_A4.at[index_label, "Quant"] = 10000
		database_excel_A4.at[index_label, "ICCID_from"] = database_excel.at[index_label * 10000, "ICCID"]
		database_excel_A4.at[index_label, "ICCID_to"] = database_excel.at[index_label * 10000 + 9999, "ICCID"]
		database_excel_A4.at[index_label, "MSISDN_from"] = database_excel.at[index_label * 10000, "MSISDN"]
		database_excel_A4.at[index_label, "MSISDN_to"] = database_excel.at[index_label * 10000 + 9999, "MSISDN"]
		database_excel_A4.at[index_label, "NUMBOX"] = "A" + order_number[-2:] + str(index_label + 1).zfill(5) #1.04
		database_excel_A4.at[index_label, "NUMBOX1"] = "B" + order_number[-2:] + str(index_label * 100 + 1).zfill(5)
		database_excel_A4.at[index_label, "NUMBOX2"] = "B" + order_number[-2:] + str(index_label * 100 + 100).zfill(5)
		index = 0
		while index < 100:
			database_excel_A4.at[index_label, "ICCID" + str(index+1)] = database_excel.at[index_label * 10000 + index * 100, "ICCID"]
			index += 1
		index = 0
		while index < 100:
			database_excel_A4.at[index_label, "NUM" + str(index+1)] = "Box № " + str(counter)
			if database_excel.at[index_label * 10000 + index * 100, "BATCH"] == database_excel.at[index_label * 10000 + (index + 1) * 100, "BATCH"]:
				counter += 1
			else:
				counter = 1
			index += 1
		#Определение батчей
		batch_list = []
		index = index_label
		new_batch_found_counter = 0
		index = index_label * 10000 
		while index < index_label * 10000 + 10000:
			if len(batch_list) == 0:	
				batch_list.append(database_excel.at[index, "BATCH"])
			else:
				for batch in batch_list:
					if batch != database_excel.at[index, "BATCH"]:
						new_batch_found_counter += 1
				if new_batch_found_counter == len(batch_list):
					batch_list.append(database_excel.at[index, "BATCH"])
				new_batch_found_counter = 0
			index += 1
		batch_final = ""
		index = 0
		while index < len(batch_list):
			batch_final += batch_list[index]
			if index + 1 != len(batch_list): batch_final += ","
			index += 1
		database_excel_A4.at[index_label, "Batch"] = batch_final
		index_label += 1
	#Для последней паллеты если она не 10000
	if database_excel["ICCID"].count() % 10000 != 0:
		database_excel_A4.at[index_label, "Pallet"] = str(index_label+1) + " of " + str(index_label_end + 1)
		database_excel_A4.at[index_label, "Quant"] = database_excel["ICCID"].count() % 10000
		if dual == 2:	database_excel_A4.at[index_label, "SIM_Quantity"] = database_excel["ICCID"].count() % 10000 * 2
		database_excel_A4.at[index_label, "ICCID_from"] = database_excel.at[index_label * 10000, "ICCID"]
		database_excel_A4.at[index_label, "ICCID_to"] = database_excel.at[index_label * 10000 + leftover - 1, "ICCID"]		
		database_excel_A4.at[index_label, "MSISDN_from"] = database_excel.at[index_label * 10000, "MSISDN"]
		database_excel_A4.at[index_label, "MSISDN_to"] = database_excel.at[index_label * 10000 + leftover - 1, "MSISDN"]
		database_excel_A4.at[index_label, "NUMBOX"] = "A" + order_number[-2:] + str(index_label + 1).zfill(5) #1.04
		database_excel_A4.at[index_label, "NUMBOX1"] = "B" + order_number[-2:] + str(index_label * 100 + 1).zfill(5)
		database_excel_A4.at[index_label, "NUMBOX2"] = "B" + order_number[-2:] + str(index_label * 100 + math.ceil(leftover / 100)).zfill(5)
		index = 0
		while index < 100:
			if index_label * 10000 + index * 100 >= database_excel["ICCID"].count(): break
			database_excel_A4.at[index_label, "ICCID" + str(index+1)] = database_excel.at[index_label * 10000 + index * 100, "ICCID"]
			index += 1
		index = 0
		while index < 100:
			database_excel_A4.at[index_label, "NUM" + str(index+1)] = "Box № " + str(counter)
			if index_label * 10000 + (index + 1) * 100 >= database_excel["ICCID"].count(): 
				#database_excel_A4.at[index_label, "NUM" + str(index+1)] = "Box № " + str(counter+1)
				break
			if database_excel.at[index_label * 10000 + index * 100, "BATCH"] == database_excel.at[index_label * 10000 + (index + 1) * 100, "BATCH"]:
				counter += 1
			else:
				counter = 1
			index += 1
		batch_list = []
		index = index_label
		new_batch_found_counter = 0
		index = index_label * 10000 
		while index < index_label * 10000 + leftover:
			if len(batch_list) == 0:	
				batch_list.append(database_excel.at[index, "BATCH"])
			else:
				for batch in batch_list:
					if batch != database_excel.at[index, "BATCH"]:
						new_batch_found_counter += 1
				if new_batch_found_counter == len(batch_list):
					batch_list.append(database_excel.at[index, "BATCH"])
				new_batch_found_counter = 0
			index += 1
		batch_final = ""
		index = 0
		while index < len(batch_list):
			batch_final += batch_list[index]
			if index + 1 != len(batch_list): batch_final += ","
			index += 1
		database_excel_A4.at[index_label, "Batch"] = batch_final	
	database_excel_A4["Zakaz"] = int(order_number)
	database_excel_A4["Otpr"] = data_otpravki
	database_excel_A4["Material1"] = material
	database_excel_A4["Date"] = activation_date

		

	try: os.makedirs(os.path.dirname(path_to_razbivka) + "/Этикетка")
	except OSError as e: 
		if e.errno != errno.EEXIST:
			raise
	writer = None
	if dual == 2:
		writer = pd.ExcelWriter(os.path.dirname(path_to_razbivka) + "/Этикетка" + "/100x60_A4_DUAL.xlsx", engine = 'xlsxwriter')
	else:
		writer = pd.ExcelWriter(os.path.dirname(path_to_razbivka) + "/Этикетка" + "/100x60_A4.xlsx", engine = 'xlsxwriter')
	database_excel_100x60.to_excel(writer, sheet_name='Печать_100х60', index=False)
	database_excel_A4.to_excel(writer, sheet_name='Печать_A4', index=False)
	writer.save()
	writer.close()

	#База для Спеклера
	status["text"]="Создание базы для Спеклера..."
	database_spekler = pd.DataFrame(dtype=str, columns=["ICCID","BOX_C","BOX_B","BATCH","BOX_A"])
	database_spekler["ICCID"] = database_excel["ICCID"]
	database_spekler["BATCH"] = database_excel["BATCH"]
	index1, index2 = 0, 0
	while index1 < database_excel_100x60["ICCID_from"].count():
		size = int(database_excel_100x60.at[index1, "Quant"]) - 1
		database_spekler.loc[index2 : index2+size, "BOX_C"] = database_excel_100x60.at[index1+1, "NUMBOX"]
		database_spekler.loc[index2 : index2+size, "BOX_B"] = database_excel_100x60.at[index1, "NUMBOX"]
		index1 += 2
		index2 += 100
	index1, index2 = 0, 0
	while index1 < database_excel_A4["ICCID_from"].count(): #1.04
		size = int(database_excel_A4.at[index1, "Quant"]) - 1
		database_spekler.loc[index2 : index2+size, "BOX_A"] = database_excel_A4.at[index1, "NUMBOX"]
		index1 += 1
		index2 += 10000
	database_spekler_path = os.path.dirname(path_to_razbivka) + "/" + str(order_number) + "_" + material_name + "_" + str(database_excel["ICCID"].count())  + ".txt"
	np.savetxt(database_spekler_path, database_spekler.values, fmt="%s", delimiter=';', newline='\n')

	#Упаковочный лист и БД
	if dual == 2:
		status["text"]="Создание первой упаковочной базы данных..."
	else:
		status["text"]="Создание упаковочной базы данных..."
	try: os.makedirs(os.path.dirname(path_to_razbivka) + "/Упаковочный лист")
	except OSError as e: 
		if e.errno != errno.EEXIST:
			raise
	database_for_packing = pd.DataFrame(dtype=str, columns=["BOX_ICCID","ICCID","MSISDN","QUANTITY"])
	database_for_packing["ICCID"] = database_excel["ICCID"]
	database_for_packing["MSISDN"] = database_excel["MSISDN"]

	index1, index2 = 0, 0
	while index1 < database_excel_100x60["ICCID_from"].count():
		size = (int(database_excel_100x60.at[index1, "Quant"]) - 1)
		database_for_packing.loc[index2 : (index2+size), "BOX_ICCID"] = database_excel_100x60.at[index1+1, "ICCID_from"]	
		database_for_packing.loc[index2 : (index2+size), "QUANTITY"] = database_excel_100x60.at[index1+1, "Quant"]	
		index1 += 2
		index2 += 100

	database_for_packing_path = os.path.dirname(path_to_razbivka) + "/Упаковочный лист/" + order_name.replace("/","_") + "_" + material_name2 + "_" + str(database_excel["ICCID"].count())  + ".txt"
	np.savetxt(database_for_packing_path, database_for_packing.values, fmt="%s", delimiter=' ', newline='\n')
	
	#Вторая упаковочная БД для дуальных (Нет договорённости с заказчиком
	'''
	if dual == 2:
		status["text"]="Создание второй упаковочной базы данных..."
		database_for_packing = pd.DataFrame(dtype=str, columns=["BOX_ICCID","ICCID","MSISDN","QUANTITY"])
		database_for_packing["ICCID"] = database_excel_for_client["ICCID"]
		database_for_packing["MSISDN"] = database_excel_for_client["MSISDN"]

		index1, index2 = 0, 0
		while index1 < database_excel_100x60["ICCID_from"].count():
			size = (int(database_excel_100x60.at[index1, "Quant"]) - 1)
			database_for_packing.loc[index2 : (index2+size), "BOX_ICCID"] = database_excel_for_client.at[index2, "ICCID"]	
			database_for_packing.loc[index2 : (index2+size), "QUANTITY"] = database_excel_100x60.at[index1+1, "Quant"]	
			index1 += 2
			index2 += 100

		database_for_packing_path = os.path.dirname(path_to_razbivka) + "/Упаковочный лист/" + order_name.replace("/","_") + "_" + material_name2 + "_SECOND_" + str(database_excel["ICCID"].count())  + ".txt"
		np.savetxt(database_for_packing_path, database_for_packing.values, fmt="%s", delimiter=' ', newline='\r\n')
	'''
	#Упаковочный лист и оформление
	status["text"]="Создание упаковочного листа..."
	workbook = xlsxwriter.Workbook(os.path.dirname(path_to_razbivka) + "/Упаковочный лист/Упаковочный лист_" + material_name2 + "_" + str(quantity_of_cards_in_total) + ".xlsx")
	worksheetPackingList = workbook.add_worksheet('Упаковочный лист')
	worksheetPackingList.set_column(0, 0, 17)
	worksheetPackingList.set_column(1, 1, 17)
	worksheetPackingList.set_column(2, 2, 9)
	worksheetPackingList.set_column(3, 3, 9)
	worksheetPackingList.set_column(4, 4, 15.38)
	worksheetPackingList.set_column(5, 5, 11.88)
	worksheetPackingList.set_column(6, 6, 10.38)
	worksheetPackingList.set_column(7, 7, 9)
	box_format = workbook.add_format({"bold":True, "font_name":"Arial", "font_size":"10", "align": "center", "valign" : "vcenter", "border":2, 'text_wrap':True})
	box_format2 = workbook.add_format({"bold":True, "font_name":"Arial", "font_size":"10", "align": "center", "valign" : "vcenter", "border":2, 'text_wrap':True, 'bg_color' : '#C0C0C0'})
	box_format3 = workbook.add_format({"bold":True, "font_name":"Arial", "font_size":"10", "align": "center", "valign" : "vcenter", "border":2, 'text_wrap':True, 'bg_color' : 'yellow'})
	y, x, index = 0, 0, 1
	netto_total, brutto_total = 0, 0
	last_index = database_excel_A4["Pallet"].count()
	while index - 1 < last_index:
		worksheetPackingList.write(y, x, "Замовник", box_format)
		worksheetPackingList.merge_range(y,x+1,y,x+7, 'КС', box_format)
		worksheetPackingList.write(y+1, x, "Найменування Товару", box_format)
		worksheetPackingList.merge_range(y+1,x+1,y+1,x+7, 'Стартовий пакет ' + material, box_format)
		worksheetPackingList.set_row(y+1, 26.25)
		worksheetPackingList.write(y+2, x, "Пакувальний лист  №", box_format)
		#Определение текущей даты
		now = datetime.datetime.now()
		current_date = str(now.day).zfill(2) + "." + str(now.month).zfill(2) + "." + str(now.year)
		worksheetPackingList.merge_range(y+2,x+1,y+2,x+7, "1/" + order_name + "/" + contract_name + " від " + current_date, box_format)
		worksheetPackingList.set_row(y+2, 26.25)
		worksheetPackingList.write(y+3, x, "№ піддону", box_format)
		worksheetPackingList.merge_range(y+3,x+1,y+3,x+2, index, box_format)
		worksheetPackingList.merge_range(y+3,x+3,y+3,x+4, "із", box_format)
		worksheetPackingList.merge_range(y+3,x+5,y+3,x+6, last_index, box_format)
		worksheetPackingList.write(y+3, x+7, "", box_format)
		worksheetPackingList.write(y+4, x, "№ Замовлення", box_format)	
		worksheetPackingList.merge_range(y+4,x+1,y+4,x+7, 'Замовлення №' + order_name + " від " + order_date , box_format)
		worksheetPackingList.set_row(y+4, 15)
		worksheetPackingList.write(y+5, x, "Дата поставки", box_format)
		worksheetPackingList.write(y+5, x+1, current_date, box_format)
		worksheetPackingList.write(y+5, x+2, "Батч", box_format)
		worksheetPackingList.write(y+5, x+3, database_excel_A4.at[index-1, "Batch"], box_format)
		worksheetPackingList.merge_range(y+5,x+4,y+5,x+5, "", box_format)
		worksheetPackingList.merge_range(y+5,x+6,y+5,x+7, "", box_format)
		worksheetPackingList.merge_range(y+6, x, y+7, x, "Перший серійний  номер", box_format2)
		worksheetPackingList.merge_range(y+6, x+1, y+7, x+1, "Останній серійний  номер", box_format2)
		worksheetPackingList.merge_range(y+6, x+2, y+7, x+2, "Кількість", box_format2)
		worksheetPackingList.merge_range(y+6, x+3, y+7, x+3, "Коробка №", box_format2)
		worksheetPackingList.merge_range(y+6, x+4, y+7, x+4, "із загальної кількості  коробок", box_format2)
		worksheetPackingList.merge_range(y+6, x+5, y+7, x+5, "Розміри  в см.", box_format2)
		worksheetPackingList.merge_range(y+6, x+6, y+7, x+6, "Вага  брутто", box_format2)
		worksheetPackingList.merge_range(y+6, x+7, y+7, x+7, "Вага нетто", box_format2)
		worksheetPackingList.set_row(y+6, 20)
		worksheetPackingList.set_row(y+7, 20)
		y += 1
		index2 = 0
		while index2 < math.ceil(int(database_excel_A4.at[index-1, "Quant"]) / 100):
			worksheetPackingList.write(y+7+index2, x, database_excel_100x60.at[index2 * 2 + (index - 1) * 200, "ICCID_from"][6:-2], box_format)
			worksheetPackingList.write(y+7+index2, x+1, database_excel_100x60.at[index2 * 2 + (index - 1) * 200, "ICCID_to"][6:-2], box_format)
			worksheetPackingList.write(y+7+index2, x+2, int(database_excel_100x60.at[index2 * 2 + (index - 1) * 200, "Quant"]), box_format)
			box_number, box_quantity = database_excel_100x60.at[index2 * 2 + (index - 1) * 200, "Pallet"].split("/")
			worksheetPackingList.write(y+7+index2, x+3, int(box_number), box_format)
			worksheetPackingList.write(y+7+index2, x+4, int(box_quantity), box_format)
			worksheetPackingList.write(y+7+index2, x+5, "28x15х13,5", box_format)
			worksheetPackingList.write(y+7+index2, x+6, "3,1 кг", box_format)
			worksheetPackingList.write(y+7+index2, x+7, "3,0 кг", box_format)
			index2 += 1
		worksheetPackingList.write(y+7+index2, x, "Розміри піддона: 120х80х15 см", box_format)
		netto = math.ceil(int(database_excel_A4.at[index-1, "Quant"]) / 100) * 3
		netto_total += netto
		brutto = int(round(math.ceil(int(database_excel_A4.at[index-1, "Quant"]) / 100) * 3.1 + 25, 0))
		brutto_total += brutto
		worksheetPackingList.write(y+7+index2, x+1, "Вага нетто: " + str(netto) + ",0 кг Вага брутто: " + str(brutto) + ",0 кг", box_format)
		worksheetPackingList.write(y+7+index2, x+2, "", box_format)
		worksheetPackingList.write(y+7+index2, x+3, "", box_format)
		worksheetPackingList.write(y+7+index2, x+4, "", box_format)
		worksheetPackingList.write(y+7+index2, x+5, "", box_format)
		worksheetPackingList.write(y+7+index2, x+6, "", box_format)
		worksheetPackingList.write(y+7+index2, x+7, "", box_format)
		y = y+8+index2
		index += 1
	worksheetPackingList.merge_range(y,x,y,x+3, "Загальна кількість Товару:", box_format)
	worksheetPackingList.merge_range(y,x+4,y,x+6, database_excel["ICCID"].count() , box_format3)
	worksheetPackingList.merge_range(y+1,x,y+1,x+3, "Загальна кількість коробок:", box_format)
	worksheetPackingList.merge_range(y+1,x+4,y+1,x+6, database_excel_100x60["Pallet"].count() // 2 , box_format3)
	worksheetPackingList.merge_range(y+2,x,y+2, x+3, "Загальна кількість піддонів:", box_format)
	worksheetPackingList.merge_range(y+2,x+4,y+2,x+6, database_excel_A4["Pallet"].count(), box_format3)
	worksheetPackingList.merge_range(y+3,x,y+3,x+3, "Загальна вага нетто, кг:", box_format)
	worksheetPackingList.merge_range(y+3,x+4,y+3,x+6, netto_total, box_format3)
	worksheetPackingList.merge_range(y+4,x,y+4,x+3, "Загальна вага брутто, кг:", box_format)
	worksheetPackingList.merge_range(y+4,x+4,y+4,x+6, brutto_total, box_format3)
	workbook.close()
	
	#Файл для склада
	name = order_number + "_" + material_name + "_" + str(quantity_of_cards_in_total)
	path_to_store_directory = "Y:/Manufacture/Склад/Киевстар/" + name
	try: 
		os.makedirs(path_to_store_directory)
		print("Обработка файла Excel для складов")
		status["text"]="Обработка файла Excel для складов..."
		database_for_warehouse = pd.DataFrame(dtype=str, columns = ["Pallet", "ICCID_from", "ICCID_to", "Quantity", "Name", "Batch", "Status"])
		database_for_warehouse["Pallet"] = database_excel_A4["Pallet"]
		database_for_warehouse["ICCID_from"] = database_excel_A4["ICCID_from"]
		database_for_warehouse["ICCID_to"] = database_excel_A4["ICCID_to"]
		database_for_warehouse["Quantity"] = database_excel_A4["Quant"]
		database_for_warehouse["Name"] = database_excel_A4["Material1"]
		database_for_warehouse["Batch"] = database_excel_A4["Batch"]
		writer = pd.ExcelWriter(path_to_store_directory + "/" + name + ".xlsx", engine = 'xlsxwriter')
		database_for_warehouse.to_excel(writer, sheet_name='Разбивка', index=False)
	except Exception as e: 
		print("Файл для склада уже существует")
		#if e.errno != errno.EEXIST:
		pass
	print("ГОТОВО!")
	status["text"]="Готово!"