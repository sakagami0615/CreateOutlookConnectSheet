import csv
import tkinter
import tkinter.ttk
import tkinter.filedialog
import tkinter.messagebox

import usersetting


class ConnectValue:

	def __init__(self, lastname, firstname, address):
		self.lastname = lastname
		self.firstname = firstname
		self.address = address


class ExportSheetFrameWidget:

	def __init__(self, parent):
		# --------------------------------------------------------------------------------
		# Init Canvas & Scrollbar BEGIN
		# --------------------------------------------------------------------------------
		# Creater Canvas
		WIDTH_MARGIN = -30
		HEIGHT_MARGIN = -40
		self.canvas = tkinter.Canvas(parent.master, width=parent.width+WIDTH_MARGIN, height=parent.height+HEIGHT_MARGIN, bg='white')
		self.canvas.grid(row=1, column=0)
		
		# Create Scrollbar
		self.vbar = tkinter.ttk.Scrollbar(parent.master, orient=tkinter.VERTICAL)
		self.vbar.grid(row=1, column=1, sticky='ns')
		
		# Set Scrollcommand
		self.vbar.config(command=self.canvas.yview)
		self.canvas.config(yscrollcommand=self.vbar.set)
		# --------------------------------------------------------------------------------
		# Init Canvas & Scrollbar END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Init Frame BEGIN
		# --------------------------------------------------------------------------------
		self.frame = tkinter.Frame(self.canvas, bg='white')		#背景を白に
		self.canvas.create_window((0,0), window=self.frame, anchor=tkinter.NW, width=self.canvas.cget('width'))		#anchor<=NWで左上に寄せる
		# --------------------------------------------------------------------------------
		# Init Frame END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Init Button BEGIN
		# --------------------------------------------------------------------------------
		self.button_export = tkinter.Button(self.frame, text='Export File', command=parent.ButtonClick_Export)
		self.button_export.grid(row=1, column=0, padx=5, pady=5, ipadx=0, ipady=0)
		# --------------------------------------------------------------------------------
		# Init Button END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Init Label BEGIN
		# --------------------------------------------------------------------------------
		label_font = tkinter.font.Font(self.frame, family='', size=9, weight='bold')

		self.select_width = 5
		self.label_select = tkinter.Label(self.frame, width=self.select_width, text='Select', font=label_font, background='white')
		self.label_select.grid(row=2, column=0, padx=0, pady=0, ipadx=0, ipady=0)
		
		self.lastname_width = 10
		self.label_lastname = tkinter.Label(self.frame, width=self.lastname_width, text='LastName', font=label_font, background='white')
		self.label_lastname.grid(row=2, column=1, padx=0, pady=0, ipadx=0, ipady=0)

		self.firstname_width = 10
		self.label_firstname = tkinter.Label(self.frame, width=self.firstname_width, text='FirstName', font=label_font, background='white')
		self.label_firstname.grid(row=2, column=2, padx=0, pady=0, ipadx=0, ipady=0)
		
		self.displayname_width = 20
		self.label_displayname = tkinter.Label(self.frame, width=self.displayname_width, text='DisplayName', font=label_font, background='white')
		self.label_displayname.grid(row=2, column=3, padx=0, pady=0, ipadx=0, ipady=0)

		self.address_width = 50
		self.label_address = tkinter.Label(self.frame, width=self.address_width, text='Address', font=label_font, background='white')
		self.label_address.grid(row=2, column=4, padx=0, pady=0, ipadx=0, ipady=0)
		# --------------------------------------------------------------------------------
		# Init Label END
		# --------------------------------------------------------------------------------
		
		
class ExportSheetFrame(tkinter.Frame):

	def __init__(self, master, tab_name, tab_index, width, height, multi_connect_list):
		super().__init__(master)

		# --------------------------------------------------------------------------------
		# Init Variable BEGIN
		# --------------------------------------------------------------------------------
		self.master = master
		self.tab_name = tab_name
		self.tab_type = 'EXPORT'
		self.tab_index = tab_index
		self.width = width
		self.height = height

		self.connect_list = self.MergeConnectList(multi_connect_list)
		self.record_num = len(self.connect_list)

		self.select_buttons = []
		self.select_variables = []
		self.lastname_labels = []
		self.firstname_labels = []
		self.displayname_labels = []
		self.address_labels = []
		# --------------------------------------------------------------------------------
		# Init Variable END
		# --------------------------------------------------------------------------------
		
		# --------------------------------------------------------------------------------
		# Init Window BEGIN
		# --------------------------------------------------------------------------------
		self.widget = ExportSheetFrameWidget(self)
		self.CreateTableView()
		# --------------------------------------------------------------------------------
		# Init Window END
		# --------------------------------------------------------------------------------


	def MergeConnectList(self, multi_connect_list):
		
		name_table_dict = dict()

		def Join(connect):
			# 名字と名前が分割できているデータを辞書に登録する
			if connect.firstname != '':
				name_key = '{}{}'.format(connect.lastname, connect.firstname)
				name_table_dict[name_key] = [connect.lastname, connect.firstname]
			join_string = '{}{}</>{}'.format(connect.lastname, connect.firstname, connect.address)
			return join_string
		
		def Separate(join_string):
			tokens = join_string.split('</>')
			name_key = tokens[0]
			# 名前辞書から名字と名前を取得
			if name_table_dict.get(name_key):
				lastname = name_table_dict[name_key][0]
				firstname = name_table_dict[name_key][1]
			else:
				lastname = name_key
				firstname = ''
			connect = ConnectValue(lastname, firstname, tokens[1])
			return connect

		join_string_set = set([Join(connect) for connect_list in multi_connect_list for connect in connect_list])
		connect_list = [Separate(join_string) for join_string in list(join_string_set)]
		connect_list = sorted(connect_list, key=lambda x:(x.lastname, x.firstname))

		return connect_list
	

	def CreateTableView(self):
		
		#スクロール可動域の設定
		LINEWIDTH = 25
		scroll_range = LINEWIDTH*(self.record_num + 1)
		self.widget.canvas.config(scrollregion=(0, 0, self.width, scroll_range))

		# レコード更新
		ROWOFFSET = 3
		for (index, connect) in enumerate(self.connect_list):
			linecolor = '#cdfff7' if index%2 == 0 else 'white'

			lastname = connect.lastname
			firstname = connect.firstname
			address = connect.address
			displayname = usersetting.GetDisplayName(lastname ,firstname, address)

			booleanvar = tkinter.BooleanVar()
			booleanvar.set(True)
			checkbutton = tkinter.Checkbutton(self.widget.frame, variable=booleanvar, width=self.widget.select_width, text='', background='white')
			checkbutton.grid(row=index+ROWOFFSET, column=0, padx=0, pady=0, ipadx=0, ipady=0)
			self.select_buttons.append(checkbutton)
			self.select_variables.append(booleanvar)

			lastname_label = tkinter.Entry(self.widget.frame, width=self.widget.lastname_width, background=linecolor)
			lastname_label.insert(0, lastname)
			lastname_label.grid(row=index+ROWOFFSET, column=1, padx=0, pady=0, ipadx=0, ipady=0)
			self.lastname_labels.append(lastname_label)

			firstname_label = tkinter.Entry(self.widget.frame, width=self.widget.firstname_width, background=linecolor)
			firstname_label.insert(0, firstname)
			firstname_label.grid(row=index+ROWOFFSET, column=2, padx=0, pady=0, ipadx=0, ipady=0)
			self.firstname_labels.append(firstname_label)
		
			displayname_label = tkinter.Entry(self.widget.frame, width=self.widget.displayname_width, background=linecolor)
			displayname_label.insert(0, displayname)
			displayname_label.grid(row=index+ROWOFFSET, column=3, padx=0, pady=0, ipadx=0, ipady=0)
			self.displayname_labels.append(displayname_label)

			address_label = tkinter.Entry(self.widget.frame, width=self.widget.address_width, background=linecolor)
			address_label.insert(0, address)
			address_label.grid(row=index+ROWOFFSET, column=4, padx=0, pady=0, ipadx=0, ipady=0)
			self.address_labels.append(address_label)


	def ButtonClick_Export(self):
		
		export_indexs = [index for index in range(self.record_num) if self.select_variables[index].get()]

		if len(export_indexs) == 0: return

		outlook_connects = []
		for index in export_indexs:
			lastname = self.lastname_labels[index].get()
			firstname = self.firstname_labels[index].get()
			displayname = self.displayname_labels[index].get()
			address = self.address_labels[index].get()
			
			outlook_record = usersetting.ConvertOutlookFormat(lastname, firstname, displayname, address)
			outlook_connects.append([outlook_record[header] for header in usersetting.OUTLOOK_FORMAT_HEADER])
		
		filename = tkinter.filedialog.asksaveasfilename(initialdir='/', title = '', filetypes=[('csv file','*.csv')])
		with open(filename, 'w', newline='') as f:
			writer = csv.writer(f)
			writer.writerow(usersetting.OUTLOOK_FORMAT_HEADER)
			writer.writerows(outlook_connects)
		
		tkinter.messagebox.showinfo('Success', 'Connect Sheet Export Success')
