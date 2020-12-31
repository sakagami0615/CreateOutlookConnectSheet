import xlrd


class ExcelValue:

	def __init__(self, filename):
		
		def GetSheetValue(book, sheet_name):
			sheet = book.sheet_by_name(sheet_name)
			sheet_value = [[sheet.cell(row, col).value for row in range(sheet.nrows)] for col in range(sheet.ncols)]
			return sheet_value, sheet.nrows, sheet.ncols

		book = xlrd.open_workbook(filename)
		self.sheet_names = tuple([sheet.name for sheet in book.sheets()])
		self.sheet_dict = dict()
		
		for sheet_name in self.sheet_names:
			sheet_value, sheet_row, sheet_col = GetSheetValue(book, sheet_name)
			self.sheet_dict[sheet_name] = dict()
			self.sheet_dict[sheet_name]['value'] = sheet_value
			self.sheet_dict[sheet_name]['nrows'] = sheet_row
			self.sheet_dict[sheet_name]['ncols'] = sheet_col


	def GetValue(self, sheet_name):
		return self.sheet_dict[sheet_name]['value']
	
	def GetRowNum(self, sheet_name):
		return self.sheet_dict[sheet_name]['nrows']
	
	def GetColNum(self, sheet_name):
		return self.sheet_dict[sheet_name]['ncols']