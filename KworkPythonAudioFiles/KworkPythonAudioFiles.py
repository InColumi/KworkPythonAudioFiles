import shutil
import os
import pandas as pd
import re

data = pd.read_excel('БАЗА АУДИО.xlsx', sheet_name=0)

#names = open("Names.txt", 'r').readlines()
#for name in names:
#	newFile = open(name[0:name.find('.')] + ".mp3", 'w')
#	newFile.close()
def getListFromData(nameData):
	isNotNull = pd.notnull(data[nameData])
	items = data.loc[isNotNull,nameData]
	
	items = [str(item).strip().replace("-", "").replace("(", "").replace(")", "").replace(" ", "") for item in items]
	res = []
	for item in items:
		if item.isdigit():
			res.append(item)
	return res

currentDir = os.getcwd()
files = os.listdir(currentDir)

listWithAllDuplicates = list()
duplicatesAllNumber = []

def MovePhoneNumbersFromFolder(numbers, nameFolder):
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
	
	duplicatesFiles = []
	duplicatesFilesReport = []
	lenNumber = len(numbers)
	countNumbers = 0
	step = lenNumber // 100
	percentStep = int(lenNumber / step)
	percentValue = step
	percent = 0
	for number in numbers:
		number7or8 = 0
		if number.find('8') == 0:
			number7or8 = re.sub('8', '7', number, 1)
		else:
			number7or8 = re.sub('7', '8', number, 1)
		
		for file in files:
			if file.find(number) != -1 or file.find(number7or8) != -1:
				if file.find(number) != -1:
					duplicatesAllNumber.append(number)
				else:
					duplicatesAllNumber.append(number7or8)
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
			print(percent,"%", sep='')
			percent += 1
			percentValue += step
		countNumbers += 1


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
	if len(numbers) - countNumbersGood >= 0:
		logFile.write("Не найдено: " + str(countNumbersBed) + '\n')
	else:
		logFile.write("Не найдено: " + str(0) + '\n')
		logFile.write("Дубликаты: " + str(countNumbersGood - len(numbers)) + '\n')

	logFile.close()
	print("План: ", len(numbers))
	print("Факт: ", str(countNumbersGood))
	if len(numbers) - countNumbersGood >= 0:
		print("Не найдено: ", str(countNumbersBed))
	else:
		print("Не найдено: ", 0)
		print("Дубликаты: ", str(countNumbersBed) - len(numbers))
	print()
	#return duplicatesFilesReport

teletel = "Teletel"
ozon = "OZON"
colCetre = "Call-Cetre"
yandexRec = "ЯндексРекрутеры"
yandexKC = "ЯндексКЦ"

#d = MovePhoneNumbersFromFolder(getListFromData(teletel), "Teletel")
#d += MovePhoneNumbersFromFolder(getListFromData(ozon), "OZON")
#d += MovePhoneNumbersFromFolder(getListFromData(colCetre), "Call-Cetre")
#d += MovePhoneNumbersFromFolder(getListFromData(yandexRec), "ЯндексРекрутеры")
#d += MovePhoneNumbersFromFolder(getListFromData(yandexKC), "ЯндексКЦ")

numbersTeletel = getListFromData(teletel)
numbersOzon = getListFromData(ozon)
numbersColCetre = getListFromData(colCetre)
numbersYandexRec = getListFromData(yandexRec)
numbersYandexKC = getListFromData(yandexKC)

MovePhoneNumbersFromFolder(numbersTeletel, "Teletel")
MovePhoneNumbersFromFolder(numbersOzon, "OZON")
MovePhoneNumbersFromFolder(numbersColCetre, "Call-Cetre")
MovePhoneNumbersFromFolder(numbersYandexRec, "ЯндексРекрутеры")
MovePhoneNumbersFromFolder(numbersYandexKC, "ЯндексКЦ")

print("Exit")
print(input())