import tkinter
import tkinter.ttk
from functools import partial

from detail_frame import DetailFrame
from excel_frame import SelectExcelFrame
from csv_frame import SelectCsvFrame
from export_frame import ExportSheetFrame


class AppicationWidget:

	def __init__(self, parent):
		# --------------------------------------------------------------------------------
		# Init MenuBar BEGIN
		# --------------------------------------------------------------------------------
		# メニューに子メニューを作成する 
		self.menubar_file = tkinter.Menu(parent)
		self.menubar_file.add_command(label='Export Sheet', command=parent.Menu_ExportSheet)
		self.menubar_file.add_separator()
		self.menubar_file.add_command(label='Exit', command=quit)

		self.menubar_tab = tkinter.Menu(parent)
		self.menubar_tab.add_command(label='Create Load Excel Tab', command=parent.Menu_LoadExcelTab)
		self.menubar_tab.add_command(label='Create Load Csv Tab', command=parent.Menu_LoadCsvTab)
		self.menubar_tab.add_separator()

		# メニューバー作成&画面にセット
		self.menubar = tkinter.Menu(parent)
		self.menubar.add_cascade(label='File', menu=self.menubar_file)
		self.menubar.add_cascade(label='Tab', menu=self.menubar_tab)
		parent.config(menu=self.menubar)
		# --------------------------------------------------------------------------------
		# Init MenuBar END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Init Tab Bar BEGIN
		# --------------------------------------------------------------------------------
		self.notebook = tkinter.ttk.Notebook(width=parent.width, height=parent.height)
		self.notebook.bind('<<NotebookTabChanged>>', parent.Notebook_ChangeTab)
		# --------------------------------------------------------------------------------
		# Init Tab Bar END
		# --------------------------------------------------------------------------------


class Appication(tkinter.Tk):

	def __init__(self):
		super().__init__()

		# --------------------------------------------------------------------------------
		# Init Variable BEGIN
		# --------------------------------------------------------------------------------
		self.width = 800
		self.height = 280

		self.tab_frames = []
		self.export_tab_cnt = 0
		self.load_tab_cnt = 0
		# --------------------------------------------------------------------------------
		# Init Variable END
		# --------------------------------------------------------------------------------
		
		# --------------------------------------------------------------------------------
		# Init Window BEGIN
		# --------------------------------------------------------------------------------
		self.geometry('{}x{}'.format(self.width, self.height))
		self.title('CreateOutlookConnectSheet')
		self.resizable(width=False, height=False)
		self.widget = AppicationWidget(self)
		# --------------------------------------------------------------------------------
		# Init Window END
		# --------------------------------------------------------------------------------

		# --------------------------------------------------------------------------------
		# Add MainTab Begin
		# --------------------------------------------------------------------------------
		tab_name = 'Main'
		tab_notebook = tkinter.Frame(self.widget.notebook)
		tab_frame = DetailFrame(tab_notebook, tab_name, 0, self.width, self.height)
		self.widget.notebook.add(tab_notebook, text=tab_name, padding=3)
		self.widget.notebook.pack(expand=1, fill='both')
		self.tab_frames.append(tab_frame)
		# --------------------------------------------------------------------------------
		# Add MainTab END
		# --------------------------------------------------------------------------------


	def Menu_ExportSheet(self):

		# Loadタブの連絡先リストを取得する
		# (Loadタブがあっても、ファイルおよび氏名/名前/アドレス列指定がない場合は、空リストとなる)
		load_tab_frames = [tab_frame for tab_frame in self.tab_frames if tab_frame.tab_type == 'LOAD']
		multi_connect_info = [tab_frame.GetContentList() for (tab_frame, is_enable) in zip(load_tab_frames, self.tab_frames[0].select_variables) if is_enable.get()]
		multi_connect_list = [connect_info[1] for connect_info in multi_connect_info if connect_info[0]]

		# 連絡先リストがない場合は、何もしない
		if len(multi_connect_list) == 0: return

		export_tab_cnt = self.export_tab_cnt + 1
		tab_index = len(self.tab_frames)

		# Export用のタブを生成する
		tab_name = 'ExportTab{}'.format(export_tab_cnt)
		tab_notebook = tkinter.Frame(self.widget.notebook)
		tab_frame = ExportSheetFrame(tab_notebook, tab_name, tab_index, self.width, self.height, multi_connect_list)
		self.widget.notebook.add(tab_notebook, text=tab_name, padding=3)
		self.widget.notebook.pack(expand=1, fill='both')

		# 生成タブに切り替え
		self.widget.notebook.select(tab_index)

		# メニューバーに削除用ボタンを追加
		nemu_tab_name = 'Delete {}'.format(tab_name)
		self.widget.menubar_tab.add_command(label=nemu_tab_name, command=partial(self.Menu_DeleteExportTab, tab_name))

		self.tab_frames.append(tab_frame)
		self.export_tab_cnt = export_tab_cnt

	
	def Menu_LoadExcelTab(self):
		load_tab_cnt = self.load_tab_cnt + 1
		tab_index = len(self.tab_frames)

		# Load用のタブを生成する
		tab_name = 'LoadTab{}'.format(load_tab_cnt)
		tab_notebook = tkinter.Frame(self.widget.notebook)
		tab_frame = SelectExcelFrame(tab_notebook, tab_name, tab_index, self.width, self.height)
		self.widget.notebook.add(tab_notebook, text=tab_name, padding=3)
		self.widget.notebook.pack(expand=1, fill='both')

		# 生成タブに切り替え
		self.widget.notebook.select(tab_index)

		# メニューバーに削除用ボタンを追加
		nemu_tab_name = 'Delete {}'.format(tab_name)
		self.widget.menubar_tab.add_command(label=nemu_tab_name, command=partial(self.Menu_DeleteLoadTab, tab_name))

		self.tab_frames.append(tab_frame)
		self.load_tab_cnt = load_tab_cnt

		# テーブルを再描画する
		self.tab_frames[0].AppendRecord()
		self.tab_frames[0].UpdateTreeview(self.tab_frames)
	

	def Menu_LoadCsvTab(self):

		load_tab_cnt = self.load_tab_cnt + 1
		tab_index = len(self.tab_frames)

		# Load用のタブを生成する
		tab_name = 'LoadTab{}'.format(load_tab_cnt)
		tab_notebook = tkinter.Frame(self.widget.notebook)
		tab_frame = SelectCsvFrame(tab_notebook, tab_name, tab_index, self.width, self.height)
		self.widget.notebook.add(tab_notebook, text=tab_name, padding=3)
		self.widget.notebook.pack(expand=1, fill='both')

		# 生成タブに切り替え
		self.widget.notebook.select(tab_index)

		# メニューバーに削除用ボタンを追加
		nemu_tab_name = 'Delete {}'.format(tab_name)
		self.widget.menubar_tab.add_command(label=nemu_tab_name, command=partial(self.Menu_DeleteLoadTab, tab_name))

		self.tab_frames.append(tab_frame)
		self.load_tab_cnt = load_tab_cnt

		# テーブルを再描画する
		self.tab_frames[0].AppendRecord()
		self.tab_frames[0].UpdateTreeview(self.tab_frames)

	
	def Menu_DeleteLoadTab(self, tab_name):

		# 削除対象のタブインデックスとメニューバーの名前を取得する
		tab_names = [tab_frame.tab_name for tab_frame in self.tab_frames]
		index = tab_names.index(tab_name)
		nemubar_tab_name = 'Delete {}'.format(tab_name)

		# 指定したLoadタブとメニューバーを削除する
		self.widget.notebook.forget(index)
		self.widget.menubar_tab.delete(nemubar_tab_name)
		
		self.tab_frames.pop(index)
		
		# テーブルを再描画する
		self.tab_frames[0].DeleteRecord(tab_name)
		self.tab_frames[0].UpdateTreeview(self.tab_frames)
	

	def Menu_DeleteExportTab(self, tab_name):

		# 削除対象のタブインデックスとメニューバーの名前を取得する
		tab_names = [tab_frame.tab_name for tab_frame in self.tab_frames]
		index = tab_names.index(tab_name)
		nemubar_tab_name = 'Delete {}'.format(tab_name)

		# 指定したExportタブとメニューバーを削除する
		self.widget.notebook.forget(index)
		self.widget.menubar_tab.delete(nemubar_tab_name)
		
		self.tab_frames.pop(index)
	

	def Notebook_ChangeTab(self, event):
		
		# 移動先のタブ名を取得する
		notebook = event.widget
		change_tab_name = notebook.tab(notebook.select(), 'text')
		
		# メインタブ移動時、テーブルを再描画する
		if change_tab_name == 'Main':
			self.tab_frames[0].UpdateTreeview(self.tab_frames)
	

	def run(self):

		self.mainloop()

