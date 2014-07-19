#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import wx
import wx.grid as gridlib
import codecs

class GridFileDropTarget(wx.FileDropTarget):
	def __init__(self, grid):
		wx.FileDropTarget.__init__(self)
		self.grid = grid

	def OnDropFiles(self, x, y, filenames):
		# ドロップされたファイルを開く
		self.grid.openFile(filenames[0])

class SimpleGrid(gridlib.Grid):
	WIDTH = 25
	HEIGHT = 25
	def __init__(self, parent, frame):
		gridlib.Grid.__init__(self, parent, -1)
		self.parent = parent
		self.frame = frame
		self.moveTo = None

		self.Bind(wx.EVT_IDLE, self.OnIdle)

		self.SetDropTarget(GridFileDropTarget(self))

		# ひとまず25x25固定とする
		self.CreateGrid(self.HEIGHT, self.WIDTH)

		# バインド
		self.Bind(gridlib.EVT_GRID_CELL_LEFT_CLICK, self.OnCellLeftClick)

		self.Bind(gridlib.EVT_GRID_CELL_CHANGE, self.OnCellChange)
		self.Bind(gridlib.EVT_GRID_SELECT_CELL, self.OnSelectCell)

		self.Bind(gridlib.EVT_GRID_EDITOR_SHOWN, self.OnEditorShown)
		self.Bind(gridlib.EVT_GRID_EDITOR_HIDDEN, self.OnEditorHidden)
		self.Bind(gridlib.EVT_GRID_EDITOR_CREATED, self.OnEditorCreated)

		self.filename = None

		self.log = sys.stdout

		self.cells = {}

	def openFile(self, filename):
		self.firsts = {}
		for j in range(self.HEIGHT):
			for i in range(self.WIDTH):
				# セルの色を元に戻す
				self.SetCellBackgroundColour(j, i, wx.WHITE)

		j = 0
		for line in codecs.open(filename, "r", "utf-8"):
			i = 0

			line = line.rstrip()
			data = line.split(",")
			for v in data:
				self.SetCellValue(j, i, v)
				# 初期値を保存しておく
				self.firsts[i+j*self.HEIGHT] = v
				i += 1
			j += 1
		self.filename = filename
		self.SetStatusText("Open file '%s'."%filename)

	def save(self):
		if self.filename == None:
			self.SetStatusText("Error: Please open file.")
			return

		self.firsts = {}
		out = ""
		bEnd = False
		for j in range(self.HEIGHT):
			for i in range(self.WIDTH):
				v = self.GetCellValue(j, i)
				self.firsts[i + j*self.HEIGHT] = v
				# セルの色を元に戻す
				self.SetCellBackgroundColour(j, i, wx.WHITE)
				if v == "":
					if i == 0: bEnd = True
					break
				if i != 0:
					out += ","+v
				else:
					# 行頭文字
					if j == 0:
						# 最初の文字
						out = v
					else:
						out += "\n"+v
			if bEnd: break

		fOut = codecs.open(self.filename, "w", "utf-8")
		fOut.write(out)
		fOut.close
		self.SetStatusText("Save as '%s'."%self.filename)


	def OnCellLeftClick(self, evt):
		self.log.write("OnCellLeftClick: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnCellRightClick(self, evt):
		self.log.write("OnCellRightClick: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnCellLeftDClick(self, evt):
		self.log.write("OnCellLeftDClick: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnCellRightDClick(self, evt):
		self.log.write("OnCellRightDClick: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnCellChange(self, evt):
		self.log.write("OnCellChange: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))

	def OnIdle(self, evt):
		if self.moveTo != None:
			self.SetGridCursor(self.moveTo[0], self.moveTo[1])
			self.moveTo = None
		evt.Skip()

	def OnSelectCell(self, evt):
		if evt.Selecting():
			# 選択済み
			msg = 'Selected'
		else:
			# 非選択状態
			msg = 'Deselected'
		self.checkDiff(evt)

		print("OnSelectCell: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))

		if self.IsCellEditControlEnabled():
			print("IsCellEditControlEnabled")
			self.HideCellEditControl()
			self.DisableCellEditControl()

		evt.Skip()

	def Cells(self, evt):
		# 現在のセルの値を取得する
		return self.GetCellValue(evt.GetRow(), evt.GetCol())

	def Firsts(self, evt):
		# 初期値を取得する
		idx = evt.GetRow() * self.HEIGHT + evt.GetCol()
		if idx in self.firsts:
			return self.firsts[idx]
		else:
			return ""

	def checkDiff(self, evt):
		# 初期値と相違があれがセルの色を変える。そうでなければ白に戻す
		color = wx.WHITE
		if self.Cells(evt) != self.Firsts(evt):
			color = wx.Colour(255, 211, 255)
		self.SetCellBackgroundColour(evt.GetRow(), evt.GetCol(), color)
	def OnEditorShown(self, evt):
		# セルのエディット開始
		self.checkDiff(evt)
		self.log.write("OnEditorShown: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnEditorHidden(self, evt):
		# セルのエディット終了
		self.checkDiff(evt)
		self.log.write("OnEditorHidden: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnEditorCreated(self, evt):
		# セルのエディット（初回のみ）
		self.log.write("OnEditorCreated: (%d,%d) %s\t" %
			(evt.GetRow(), evt.GetCol(), evt.GetControl()))

	def SetStatusText(self, msg):
		self.frame.statusbar.SetStatusText(msg)


class AppFrame(wx.Frame):
	def __init__(self, parent, log):
		wx.Frame.__init__(self, parent, -1, "CSV Editor", size=(640, 480))

		# パネル生成
		panel = wx.Panel(self, -1, style=0)
		panel.SetBackgroundColour("#CFCFCF")

		# ツールバー生成
		toolbar = wx.ToolBar(panel)
		self.toolbar = toolbar
		# 保存アイコン追加
		toolbar.AddLabelTool(wx.ID_SAVE, "Save", wx.Bitmap("./icons/save.gif"))
		toolbar.Realize()

		self.Bind(wx.EVT_TOOL, self.OnSave, id=wx.ID_SAVE)

		# グリッド生成
		grid = SimpleGrid(panel, self)
		self.grid = grid

		# メニューバー生成
		menu_file = wx.Menu()
		menu_save = wx.MenuItem(menu_file, 1, u"&Save\tCtrl+S")
		menu_file.AppendItem(menu_save)
		menu_file.AppendSeparator()
		menu_file.Append(2, "&Quit", "Quit CsvEditor")

		self.Bind(wx.EVT_MENU, self.OnSave, menu_save)

		menu_edit = wx.Menu()
		menu_edit.Append(3, u"Copy")
		menu_edit.Append(4, u"Paste")

		menu_bar = wx.MenuBar()
		menu_bar.Append(menu_file, u"File")
		menu_bar.Append(menu_edit, u"Edit")

		self.SetMenuBar(menu_bar)

		# ステータスバー生成
		statusbar = wx.StatusBar(panel)
		self.statusbar = statusbar
		statusbar.SetStatusText("Please open csv file.")

		bs = wx.BoxSizer(wx.VERTICAL)
		bs.Add(toolbar)
		bs.Add(grid, 1, wx.GROW|wx.ALL, 5)
		bs.Add(statusbar)
		panel.SetSizer(bs)

	def OnExit(self, evt):
		self.Close()

	def OnSave(self, evt):
		self.grid.save()

if __name__ == "__main__":
	import sys
	from wx.lib.mixins.inspection import InspectableApp
	application = InspectableApp(False)

	frame = AppFrame(None, sys.stdout)
	frame.Show()

	# テスト用コード
	frame.grid.openFile("test.csv")

	application.MainLoop()
