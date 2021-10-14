import shutil
import os
import pandas as pd
import re

def getListFromData(nameData, data):
	isNotNull = pd.notnull(data[nameData])
	size = len(isNotNull)
	items = []
	item = 0
	countRow = 2
	res = []
	for i in range(size):
		if isNotNull[i]:
			item = str(data[nameData][i]).strip().replace(" ", "").replace("(", "").replace(")", "").replace("-", "").replace(".0", "").replace("\'", "")
			item = str(item).strip().replace("+", "")
			if len(item) == 11 and item.isdigit():
				res.append(item)
			else:
				print('Что-то не так с номером. Строка:', countRow, 'Столбец: ', nameData, 'Номер: ', item)

		countRow += 1
	
	##countRow = 2
	#for item in items:
	#	if len(item) == 11 and item.isdigit():
	#		res.append(item)
	#	else:
	#		print(type(item))
	#		print(item.isdigit())
	#		print('Что-то не так с номером.\nСтрока:', countRow, '\nСтолбец: ', nameData, '\nНомер: ', item)
	#	countRow += 1
	return res

def MovePhoneNumbersFromFolder(numbers, nameFolder, files, currentDir):
	print("Компания: ", nameFolder)

	logFileName = "Log File %s.txt" % nameFolder
	try:
	    os.mkdir(nameFolder)
	except OSError:
	    print("Создать директорию %s не удалось" % nameFolder)
	else:
	    print("Успешно создана директория %s " % nameFolder)
	logFile = open(nameFolder + "\\" + logFileName, 'w')
	countNumbersGood = 0
	countNumbersBed = 0
	isFound = False
	
	lenNumber = len(numbers)
	countNumbers = 0
	step = lenNumber // 100
	isShowPercent = True
	percentValue = step
	percent = 0
	percentStep = 0
	percent = 101 # 101 потому что если номеров меньше 100, чтобы не выводить проценты
	if lenNumber >= 100:
		percentStep = int(lenNumber / step)
		percent = 0

	duplicatesFiles = []
	duplicatesFilesReport = []

	for number in numbers:
		number7or8 = 0
		if number.find('8') == 0:
			number7or8 = re.sub('8', '7', number, 1)
		else:
			number7or8 = re.sub('7', '8', number, 1)
		
		for file in files:
			if file.find(number) != -1 or file.find(number7or8) != -1:
				isFound = True
				duplicatesFiles.append(file)
		
		if len(duplicatesFiles) > 1:
			duplicatesFilesReport += duplicatesFiles
			duplicatesFiles.clear()
		elif len(duplicatesFiles) == 1:
			countNumbersGood += 1
			shutil.copy(currentDir + "\\" + duplicatesFiles[0], currentDir + "\\" + nameFolder)
			#print(countNumbersGood, duplicatesFiles[0])
			duplicatesFiles.clear()
		else:
			duplicatesFiles.clear()
		
		if isFound == False:
			countNumbersBed += 1
			logFile.write(number)
			logFile.write('\n')
		else:
			isFound = False

		if countNumbers == percentValue and percent <= 100:
			print('[','#' * (percent),' ' * (100 - percent),']', end='')
			print('\r', end='')
			percent += 1
			percentValue += step
		countNumbers += 1
	print()

	if len(duplicatesFilesReport) > 1:
		try:
			os.mkdir(nameFolder + "\\" + nameFolder + " Дубликаты")
		except OSError:
			print("Создать директорию дубликаты не удалось")
		else:
			print("Успешно создана директория дубликаты ")
		
		for file in duplicatesFilesReport:
			shutil.copy(currentDir + "\\" + file, currentDir + "\\" + nameFolder + "\\" + nameFolder + " Дубликаты")

			#if os.path.exists(currentDir + "\\" + file):
			#	print("Дубликаты: ", file)

	logFile.write("План: " + str(len(numbers)) + '\n')
	logFile.write("Факт: " + str(countNumbersGood) + '\n')
	logFile.write("Не найдено: " + str(countNumbersBed) + '\n')
	logFile.write("Дублей: " + str(len(duplicatesFilesReport)) + '\n')
	logFile.close()

	print("План: " + str(len(numbers)) + '\n')
	print("Факт: " + str(countNumbersGood) + '\n')
	print("Не найдено: " + str(countNumbersBed) + '\n')
	print("Дублей: " + str(len(duplicatesFilesReport)) + '\n')

	print()

def main():
	nameDB = 'БАЗА АУДИО.xlsx'
	print("Имя файла с БД: " + nameDB)
	try:
		data = pd.read_excel('БАЗА АУДИО.xlsx', sheet_name=0)
	except BaseException:
		print(input("Что-то с базой..."))

		return
		print(input())
	currentDir = os.getcwd()
	files = os.listdir(currentDir)
	
	columnNames = data.columns

	availableNames = []
	for name in columnNames:
		if name.find('Unnamed') == -1:
			availableNames.append(name)
	
	for name in availableNames:
		numbers = getListFromData(name, data)
		MovePhoneNumbersFromFolder(numbers, name, files, currentDir)
	
	print(input("Программа закончила свою работу. Нажмите Enter"))

if __name__ == "__main__":
    main()