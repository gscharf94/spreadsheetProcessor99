from openpyxl import Workbook, load_workbook

class Handler():
	def __init__(self,fileName):
		### takes in import file
		self.fileName = fileName
		self.wb = load_workbook(fileName)
		self.sheetNames = self.wb.sheetnames

	def getRecordsFromSheet(self,sheetName):
		sheet = self.wb[sheetName]
		edge = self.findEdge(sheet)

		rows = []
		
		cols = ['A','B','C','D','E','F','G','H']

		for x in range(2,edge):
			row = []
			for y in cols:
				loc = y+str(x)
				row.append(sheet[loc].value)
			rows.append(row)

		return rows

	def findEdge(self,sheet):
		counter = 1
		while True:
			if sheet['A'+str(counter)].value == None:
				break
			else:
				counter += 1
		return counter

	def initMainDict(self):
		canvNames = self.sheetNames[2:]
		self.mainDict = {}
		for name in canvNames:
			rows = self.getRecordsFromSheet(name)
			self.mainDict[name] = rows

	def initExtraDict(self):
		self.extraDict = {}
		sheet = self.wb[self.sheetNames[0]]
		edge = self.findEdge(sheet)

		cols = ['B','D','F']
		rows = []
		for x in range(1,edge):
			row = []
			for y in cols:
				loc = y+str(x)
				row.append(sheet[loc].value)
			rows.append(row)
		print(rows)


handler = Handler('4-24-2020.xlsx')
# handler.getRecordsFromSheet('Abreu, Rosaura')
# handler.initMainDict()
handler.initExtraDict()