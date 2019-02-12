import pandas as pd
import numpy as np
import os
import xlsxwriter
import math
from tkinter import *

def diapazon(path_directory, first_Number, last_Number, mode, file_extension, status):
	status["text"]="Обработка..."
	list = []
	first_Number_int = int(first_Number)
	last_Number_int = int(last_Number)
	first_Number_str = first_Number
	last_Number_str = last_Number
	length = len(first_Number)
	if mode == 1:
		while first_Number_int <= last_Number_int:
			list.append(str(first_Number_int).zfill(length))
			first_Number_int += 1

	if mode == 2:
		while first_Number_int <= last_Number_int:
			sum = 0
			first_Number_String = str(first_Number_int)
			for digit_index in range(len(first_Number_String)):
				if digit_index % 2 != 0:
					sum += int(first_Number_String[-(digit_index + 1)])
				else:
					double_number = str(2 * int(first_Number_String[-(digit_index + 1)]))
					sum += int(double_number[-1])
					if len(double_number) > 1:
						sum += int(double_number[-2])
			nearest_number = int(math.ceil(sum / 10.0)) * 10
			check_digit = nearest_number - sum
			list.append(str(first_Number_int).zfill(length) + str(check_digit))
			first_Number_int += 1
	if mode == 3:
		while int(first_Number_str) <= int(last_Number_str):
			sum1, sum2 = 0, 0
			for digit in first_Number_str[-1::-2]:
				sum1 += int(digit)
			sum1 *= 3
			for digit in first_Number_str[-2::-2]:
				sum2 += int(digit)
			sum = sum1 + sum2
			nearest_number = int(math.ceil(sum / 10.0)) * 10
			check_digit = nearest_number - sum
			list.append(first_Number_str.zfill(12) + str(check_digit))
			first_Number_str = str(int(first_Number_str) + 1)
	database = pd.DataFrame(list)
	if file_extension == ".xlsx":
		database.to_excel(path_directory, index=False, header=False, sheet_name="Лист1")
	elif file_extension == ".txt":
		np.savetxt(path_directory, database.values, fmt="%s", newline='\r\n')
	status["text"]="Готово!"
def code128(number):
	sum = 0
	for digit_index in range(len(number)):
		if digit_index % 2 != 0:
			sum += int(number[-(digit_index + 1)])
		else:
			double_number = str(2 * int(number[-(digit_index + 1)]))
			sum += int(double_number[-1])
			if len(double_number) > 1:
				sum += int(double_number[-2])
		nearest_number = int(math.ceil(sum / 10.0)) * 10
		check_digit = nearest_number - sum
	return number + str(check_digit)