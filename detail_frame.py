import tkinter
import tkinter.ttk
import tkinter.font


class DetailFrameWidget:

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
		# Init Label BEGIN
		# --------------------------------------------------------------------------------
		label_font = tkinter.font.Font(self.frame, family='', size=9, weight='bold')

		self.select_width = 5
		self.label_select = tkinter.Label(self.frame, width=self.select_width, text='Select', font=label_font, background='white')
		self.label_select.grid(row=1, column=0, padx=0, pady=0, ipadx=0, ipady=0)
		
		self.tabname_width = 15
		self.label_tabname = tkinter.Label(self.frame, width=self.tabname_width, text='TabName', font=label_font, background='white')
		self.label_tabname.grid(row=1, column=1, padx=0, pady=0, ipadx=0, ipady=0)

		self.type_width = 15
		self.label_type = tkinter.Label(self.frame, width=self.type_width, text='Type', font=label_font, background='white')
		self.label_type.grid(row=1, column=2, padx=0, pady=0, ipadx=0, ipady=0)
		
		self.filepath_width = 70
		self.label_filepath = tkinter.Label(self.frame, width=self.filepath_width, text='FilePath', font=label_font, background='white')
		self.label_filepath.grid(row=1, column=3, padx=0, pady=0, ipadx=0, ipady=0)
		# --------------------------------------------------------------------------------
		# Init Label END
		# --------------------------------------------------------------------------------
		
		
class DetailFrame(tkinter.Frame):

	def __init__(self, master, tab_name, tab_index, width, height):
		super().__init__(master)

		# --------------------------------------------------------------------------------
		# Init Variable BEGIN
		# --------------------------------------------------------------------------------
		self.master = master
		self.tab_name = tab_name
		self.tab_type = 'MAIN'
		self.tab_index = tab_index
		self.width = width
		self.height = height
		
		self.select_buttons = []
		self.select_variables = []
		self.tabname_labels = []
		self.type_labels = []
		self.flepath_labels = []
		self.record_num = 0
		# --------------------------------------------------------------------------------
		# Init Variable END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Init Window BEGIN
		# --------------------------------------------------------------------------------
		self.widget = DetailFrameWidget(self)
		self.UpdateTreeview([])
		# --------------------------------------------------------------------------------
		# Init Window END
		# --------------------------------------------------------------------------------


	def UpdateTreeview(self, tab_frames):

		load_tab_frames = [tab_frame for tab_frame in tab_frames if tab_frame.tab_type == 'LOAD']
		load_tab_num = len(load_tab_frames)

		#スクロール可動域の設定
		LINEWIDTH = 25
		scroll_range = LINEWIDTH*(load_tab_num + 1)
		self.widget.canvas.config(scrollregion=(0, 0, self.width, scroll_range))
		
		# 色付け
		for index in range(self.record_num):
			linecolor = '#cdfff7' if index%2 == 0 else 'white'
			self.select_buttons[index].config(background='white')
			self.tabname_labels[index].config(background=linecolor)
			self.type_labels[index].config(background=linecolor)
			self.flepath_labels[index].config(background=linecolor)
		
		# レコード更新
		for (index, load_tab_frame) in enumerate(load_tab_frames):
			tab_name = load_tab_frame.tab_name
			filetype = load_tab_frame.filetype
			select_file = load_tab_frame.widget.select_file.get()
			self.tabname_labels[index].config(text=tab_name)
			self.type_labels[index].config(text=filetype)
			self.flepath_labels[index].config(text=select_file)
	
	
	def AppendRecord(self):
		
		ROWOFFSET = 2
		index = self.record_num
		
		booleanvar = tkinter.BooleanVar()
		booleanvar.set(True)
		checkbutton = tkinter.Checkbutton(self.widget.frame, variable=booleanvar, width=self.widget.select_width, text='')
		checkbutton.grid(row=index+ROWOFFSET, column=0, padx=0, pady=0, ipadx=0, ipady=0)
		self.select_buttons.append(checkbutton)
		self.select_variables.append(booleanvar)
		
		tabname_label = tkinter.Label(self.widget.frame, width=self.widget.tabname_width)
		tabname_label.grid(row=index+ROWOFFSET, column=1, padx=0, pady=0, ipadx=0, ipady=0)
		self.tabname_labels.append(tabname_label)

		type_label = tkinter.Label(self.widget.frame, width=self.widget.type_width)
		type_label.grid(row=index+ROWOFFSET, column=2, padx=0, pady=0, ipadx=0, ipady=0)
		self.type_labels.append(type_label)
	
		flepath_label = tkinter.Label(self.widget.frame, width=self.widget.filepath_width)
		flepath_label.grid(row=index+ROWOFFSET, column=3, padx=0, pady=0, ipadx=0, ipady=0)
		self.flepath_labels.append(flepath_label)

		self.record_num = self.record_num + 1

	
	def DeleteRecord(self, tab_name):

		tab_names = [tabname_label['text'] for tabname_label in self.tabname_labels]
		index = tab_names.index(tab_name)

		self.select_buttons[index].pack_forget()
		self.select_buttons[index].destroy()
		self.select_buttons.pop(index)
		self.select_variables.pop(index)
		self.tabname_labels[index].pack_forget()
		self.tabname_labels[index].destroy()
		self.tabname_labels.pop(index)
		self.type_labels[index].pack_forget()
		self.type_labels[index].destroy()
		self.type_labels.pop(index)
		self.flepath_labels[index].pack_forget()
		self.flepath_labels[index].destroy()
		self.flepath_labels.pop(index)
		self.record_num = self.record_num - 1
