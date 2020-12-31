import os
import tkinter
import tkinter.ttk
import tkinter.filedialog

import utility
from csv_value import CsvValue
from export_frame import ConnectValue


class SelectCsvFrameWidget:

	def __init__(self, parent):		
		# --------------------------------------------------------------------------------
		# Init Canvas BEGIN
		# --------------------------------------------------------------------------------
		self.canvas = tkinter.Canvas(parent.master, height=parent.height, width=parent.width)
		self.canvas.create_line(5, 70, 290, 70, fill='black')
		self.canvas.grid(row=0, column=0)
		
		self.frame_select = tkinter.Frame(self.canvas)
		self.frame_input = tkinter.Frame(self.canvas)
		self.frame_setting = tkinter.Frame(self.canvas)
		self.frame_treeview = tkinter.Frame(self.canvas)

		self.frame_select_width = 300
		self.frame_input_width = 300
		self.frame_setting_width = 300
		self.frame_treeview_width = width = parent.width - 300

		self.frame_select_height = 70		# 本来であれば80だが、-10して線描画領域を用意している
		self.frame_input_height = 120
		self.frame_setting_height = parent.height - 200
		self.frame_treeview_height = parent.height
		
		self.frame_select_pos = (0, 0)
		self.frame_input_pos = (0, 80)
		self.frame_setting_pos = (0, 200)
		self.frame_treeview_pos = (300, 0)
		
		self.canvas.create_window(self.frame_select_pos, window=self.frame_select, anchor=tkinter.NW, width=self.frame_select_width, height=self.frame_select_height)
		self.canvas.create_window(self.frame_input_pos, window=self.frame_input, anchor=tkinter.NW, width=self.frame_input_width, height=self.frame_input_height)
		self.canvas.create_window(self.frame_setting_pos, window=self.frame_setting, anchor=tkinter.NW, width=self.frame_setting_width, height=self.frame_setting_height)
		self.canvas.create_window(self.frame_treeview_pos, window=self.frame_treeview, anchor=tkinter.NW, width=self.frame_treeview_width, height=self.frame_treeview_height)
		# --------------------------------------------------------------------------------
		# Init Canvas END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Layout frame_select BEGIN
		# --------------------------------------------------------------------------------
		# Create Button 
		self.button_select_file = tkinter.Button(self.frame_select, text='File Select', command=parent.ButtonClick_SelectFile)
		self.button_select_file.grid(row=0, column=0, padx=10, pady=5)
		
		# Create Entry
		self.select_file = tkinter.Entry(self.frame_select, text='')
		self.select_file.configure(state='readonly')
		self.select_file.grid(row=0, column=1, padx=10, pady=5, ipadx=30)
		# --------------------------------------------------------------------------------
		# Layout frame_select END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Layout frame_input BEGIN
		# --------------------------------------------------------------------------------
		# Create Label
		self.label_header_line = tkinter.ttk.Label(self.frame_input, text='Header ColLine')
		self.label_header_line.grid(row=0, column=0, padx=10, pady=5, ipadx=0, ipady=0)

		self.label_lastname_line = tkinter.ttk.Label(self.frame_input)
		self.label_lastname_line.grid(row=1, column=0, padx=10, pady=5, ipadx=0, ipady=0)

		self.label_firstname_line = tkinter.ttk.Label(self.frame_input)
		self.label_firstname_line.grid(row=2, column=0, padx=10, pady=5, ipadx=0, ipady=0)

		self.label_address_line = tkinter.ttk.Label(self.frame_input, text='Address RowLine')
		self.label_address_line.grid(row=3, column=0, padx=10, pady=5, ipadx=0, ipady=0)
		
		# Create Entry
		self.input_header_line = tkinter.Entry(self.frame_input, text='')
		self.input_header_line.grid(row=0, column=1, padx=10, pady=0, ipadx=10, ipady=0)

		self.input_lastname_line = tkinter.Entry(self.frame_input, text='')
		self.input_lastname_line.grid(row=1, column=1, padx=10, pady=0, ipadx=10, ipady=0)

		self.input_firstname_line = tkinter.Entry(self.frame_input, text='')
		self.input_firstname_line.grid(row=2, column=1, padx=10, pady=0, ipadx=10, ipady=0)
		
		self.input_address_line = tkinter.Entry(self.frame_input, text='')
		self.input_address_line.grid(row=3, column=1, padx=10, pady=0, ipadx=10, ipady=0)
		# --------------------------------------------------------------------------------
		# Layout frame_input END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Layout frame_setting BEGIN
		# --------------------------------------------------------------------------------
		# Create Checkbutton
		self.variable_namemode = tkinter.BooleanVar()
		self.variable_namemode.set(False)
		self.check_namemode = tkinter.Checkbutton(self.frame_setting, variable=self.variable_namemode, text='', command=parent.CheckButton_ModeChange)
		self.check_namemode.grid(row=0, column=0, padx=10, pady=0, ipadx=0, ipady=0)

		# Create Label
		self.label_header_line = tkinter.ttk.Label(self.frame_setting, text='Name separate space char')
		self.label_header_line.grid(row=0, column=1, padx=0, pady=0, ipadx=0, ipady=0)

		# Create Button
		self.button_view = tkinter.Button(self.frame_setting, text='View', command=parent.ButtonClick_View)
		self.button_view.grid(row=0, column=2, padx=35, pady=10, ipadx=10, ipady=0)
		# --------------------------------------------------------------------------------
		# Layout frame_setting END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Layout frame_treeview BEGIN
		# --------------------------------------------------------------------------------
		# Create treeview
		self.treeview = tkinter.ttk.Treeview(self.frame_treeview)
		self.treeview.place(x=0, y=0, height=230, width=470)

		# Create Scrollbar
		self.vbar_treeview = tkinter.Scrollbar(self.frame_treeview, orient=tkinter.VERTICAL, command=self.treeview.yview)
		self.treeview.configure(yscrollcommand=self.vbar_treeview.set)
		self.vbar_treeview.place(x=470, y=0, height=230)
		
		self.hbar_treeview = tkinter.Scrollbar(self.frame_treeview, orient=tkinter.HORIZONTAL, command=self.treeview.xview)
		self.treeview.configure(xscrollcommand=self.hbar_treeview.set)
		self.hbar_treeview.place(x=0, y=230, width=470)
		# --------------------------------------------------------------------------------
		# Layout frame_treeview END
		# --------------------------------------------------------------------------------


class SelectCsvFrame(tkinter.ttk.Frame):

	def __init__(self, master, tab_name, tab_index, width, height):
		super().__init__(master)
		
		# --------------------------------------------------------------------------------
		# Init Variable BEGIN
		# --------------------------------------------------------------------------------
		self.master = master
		self.tab_type = 'LOAD'
		self.tab_name = tab_name
		self.tab_index = tab_index
		self.width = width
		self.height = height
		self.filetype = 'CSV'
		self.csv_object = None

		self.VALUE_WIDTH_NORM = 9		# 1文字の描画幅
		self.VALUE_MIN_WIDTH = 5*9		# データ列の最小
		# --------------------------------------------------------------------------------
		# Init Variable END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Init Window BEGIN
		# --------------------------------------------------------------------------------
		self.widget = SelectCsvFrameWidget(self)
		self.CheckButton_ModeChange()
		# --------------------------------------------------------------------------------
		# Init Window END
		# --------------------------------------------------------------------------------

	
	def GetContentList(self):

		if not self.csv_object: return False, None

		sheet_value = self.csv_object.GetValue()
		row_size = self.csv_object.GetRowNum()

		# フルネームを空白文字で名字と名前に分割するパターン
		if self.widget.variable_namemode.get():
			header_line = self.GetHeaderLine()
			fullname_line = self.GetLastnameLine()
			address_line = self.GetAddressLine()

			# 未入力項目があった場合は、Falseを返す
			if (fullname_line == -1) or (address_line == -1): return False, None

			# ヘッダ行以下の値を取得する
			def SplitName(fullname):
				tokens = fullname.split()
				if len(tokens) == 2:
					lastname = tokens[0]
					firstname = tokens[1]
				else:
					lastname = fullname
					firstname = ''
				return lastname, firstname
			
			address_list = [sheet_value[address_line][row_index] for row_index in range(row_size) if row_index > header_line]
			lastname_list = []
			firstname_list = []
			for row_index in range(row_size):
				if row_index > header_line:
					lastname, firstname = SplitName(sheet_value[fullname_line][row_index])
					lastname_list.append(lastname)
					firstname_list.append(firstname)

		# 指定した列から名字と名前を取得するパターン
		else:
			header_line = self.GetHeaderLine()
			lastname_line = self.GetLastnameLine()
			firstname_line = self.GetFirstnameLine()
			address_line = self.GetAddressLine()

			# 未入力項目があった場合は、Falseを返す
			if (lastname_line == -1) or (firstname_line == -1) or (address_line == -1): return False, None

			# ヘッダ行以下の値を取得する
			lastname_list = [sheet_value[lastname_line][row_index] for row_index in range(row_size) if row_index > header_line]
			firstname_list = [sheet_value[firstname_line][row_index] for row_index in range(row_size) if row_index > header_line]
			address_list = [sheet_value[address_line][row_index] for row_index in range(row_size) if row_index > header_line]
		
		# 連絡先レコードリストに変換する
		pre_connect_list = [ConnectValue(lastname, firstname, address) for(lastname, firstname, address) in zip(lastname_list, firstname_list, address_list)]
		# 連絡先のアドレスにドメインがないものを除去する
		# (レコードが一つもない場合は、Falseを返す)
		connect_list = [connect for connect in pre_connect_list if len(connect.address.split('@')) > 1]
		if len(connect_list) == 0: return False, None

		return True, connect_list
	

	def GetHeaderLine(self):

		row_size = self.csv_object.GetRowNum()
		input_header_line = self.widget.input_header_line.get()
		
		# ヘッダ行の入力値を数値に変換する
		# (数値の変換に失敗 or データ行数より大きい値の場合は-1を返却する)
		header_index = -1
		if input_header_line.isdecimal():
			header_line = int(input_header_line)
			if header_line <= row_size:
				header_index = header_line - 1
		
		return header_index
	

	def GetLastnameLine(self):

		col_size = self.csv_object.GetColNum()
		input_lastname_line = self.widget.input_lastname_line.get()

		# 名字列の入力値を数値に変換する
		# (数値の変換に失敗 or データ列数より大きい値の場合は-1を返却する)
		lastname_index = -1
		if input_lastname_line.isdecimal():
			lastname_line = int(input_lastname_line)
			if lastname_line <= col_size:
				lastname_index = lastname_line - 1
		
		return lastname_index
	

	def GetFirstnameLine(self):

		col_size = self.csv_object.GetColNum()
		input_firstname_line = self.widget.input_firstname_line.get()

		# 名前列の入力値を数値に変換する
		# (数値の変換に失敗 or データ列数より大きい値の場合は-1を返却する)
		firstname_index = -1
		if input_firstname_line.isdecimal():
			firstname_line = int(input_firstname_line)
			if firstname_line <= col_size:
				firstname_index = firstname_line - 1
		
		return firstname_index
	

	def GetAddressLine(self):

		col_size = self.csv_object.GetColNum()
		input_address_line = self.widget.input_address_line.get()

		# アドレス列の入力値を数値に変換する
		# (数値の変換に失敗 or データ列数より大きい値の場合は-1を返却する)
		address_index = -1
		if input_address_line.isdecimal():
			address_line = int(input_address_line)
			if address_line <= col_size:
				address_index = address_line - 1
		
		return address_index

	
	def CheckButton_ModeChange(self):

		if self.widget.variable_namemode.get():
			self.widget.label_firstname_line.config(state='disable')
			self.widget.input_firstname_line.config(state='disable')
			self.widget.label_lastname_line.config(text='FullName RowLine')
			self.widget.label_firstname_line.config(text='None')
		else:
			self.widget.label_firstname_line.config(state='normal')
			self.widget.input_firstname_line.config(state='normal')
			self.widget.label_lastname_line.config(text='LastName RowLine')
			self.widget.label_firstname_line.config(text='FirstName RowLine')


	def ButtonClick_SelectFile(self):
	
		file_types = [('', '*.csv;*.CSV')]
		init_dir = os.path.abspath(os.path.dirname(__file__))
		file_name = tkinter.filedialog.askopenfilename(filetypes=file_types, initialdir=init_dir)
		
		if file_name == '': return
		
		# 編集有効にし、ファイル名をセットし、読込のみに変更する
		self.widget.select_file.configure(state='normal')
		self.widget.select_file.insert(tkinter.END, file_name)
		self.widget.select_file.configure(state='readonly')
		
		self.csv_object = CsvValue(file_name)
	
	
	def ButtonClick_View(self):
		
		if not self.csv_object: return
		
		sheet_value = self.csv_object.GetValue()
		row_size = self.csv_object.GetRowNum()
		col_size = self.csv_object.GetColNum()
		header_index = self.GetHeaderLine()
		

		# 現在のテーブルを削除
		for index in self.widget.treeview.get_children(): self.widget.treeview.delete(index)

		# 列インデックス,表スタイルの設定
		self.widget.treeview['columns'] = tuple([col_index + 1 for col_index in range(col_size)])
		self.widget.treeview['show'] = 'headings'
		
		# 各列幅の設定
		for col_index in range(col_size):
			value_lengths = [utility.CalcStringWidth(value)*self.VALUE_WIDTH_NORM for value in sheet_value[col_index]]
			value_width = max(value_lengths)
			self.widget.treeview.column(col_index + 1, width=max(value_width, self.VALUE_MIN_WIDTH))

		# 各列のヘッダー設定
		col_headers = ['row{}'.format(col_index + 1) for col_index in range(col_size)]
		[self.widget.treeview.heading(col_index + 1, text=col_header) for (col_index, col_header) in enumerate(col_headers)]

		# テーブル描画
		for record_index in range(row_size):
			tag = 'header' if record_index == header_index else ''
			record_data = tuple([value[record_index] for value in sheet_value])
			self.widget.treeview.insert('', 'end', values=record_data, tag=tag)
		self.widget.treeview.tag_configure('header', background='#00ee00')
