import csv


class CsvValue:

	def __init__(self, filename):
		
		def GetSheetValue(reader):
			sheet_col = 0
			row_value = []
			for (index, row) in enumerate(reader):
				if index == 0: sheet_col = len(row)
				row_value.append(row)
			sheet_row = len(row_value)	

			sheet_value = [[row_value[row_index][col_index] for row_index in range(sheet_row)] for col_index in range(sheet_col)]
			return sheet_value, sheet_row, sheet_col

		with open(filename, encoding='utf-8_sig') as f:
			reader = csv.reader(f)
			sheet_value, sheet_row, sheet_col = GetSheetValue(reader)

		self.sheet_dict = dict()
		self.sheet_dict['value'] = sheet_value
		self.sheet_dict['nrows'] = sheet_row
		self.sheet_dict['ncols'] = sheet_col

	
	def GetValue(self):
		return self.sheet_dict['value']
	
	def GetRowNum(self):
		return self.sheet_dict['nrows']
	
	def GetColNum(self):
		return self.sheet_dict['ncols']
