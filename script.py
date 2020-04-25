from openpyxl import Workbook, load_workbook

class Handler():
	def __init__(self,fileName):
		### takes in import file
		self.fileName = fileName
		self.loadFile(fileName)

	def loadFile(self,fileName):
		self.wb = load_workbook(fileName)

	def getSheetNames(self):
		sheetNames = self.wb.sheetnames
		print(sheetNames)

	def getRecordsFromSheet(self,sheetName):
		print(self.wb[sheetName]['B3'].value)
		x = self.findEdge(sheetName)
		print(x)

	def findEdge(self,sheetName):
		sheet = self.wb[sheetName]
		counter = 1
		while True:
			if sheet['A'+str(counter)].value == None:
				break
			else:
				counter += 1
		return counter


handler = Handler('4-24-2020.xlsx')
handler.getSheetNames()
handler.getRecordsFromSheet('Abreu, Rosaura')