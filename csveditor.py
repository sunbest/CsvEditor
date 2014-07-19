#!/usr/bin/env python
# -*- coding: utf-8 -*-

import wx
import wx.grid as gridlib
import codecs

class GridFileDropTarget(wx.FileDropTarget):
	def __init__(self, grid):
		wx.FileDropTarget.__init__(self)
		self.grid = grid

	def OnDropFiles(self, x, y, filenames):
		# 指定のファイルを開く
		self.grid.openFile(filenames[0])

class SimpleGrid(gridlib.Grid):
	def __init__(self, parent, log):
		gridlib.Grid.__init__(self, parent, -1)
		self.log = log
		self.moveTo = None

		self.Bind(wx.EVT_IDLE, self.OnIdle)

		self.SetDropTarget(GridFileDropTarget(self))

		self.CreateGrid(25, 25)
		# self.SetColSize(3, 200)
		# self.SetRowSize(4, 45)

		self.filename = None

	def openFile(self, filename):
		j = 0
		for line in codecs.open(filename, "r", "utf-8"):
			i = 0

			line = line.rstrip()
			data = line.split(",")
			for v in data:
				self.SetCellValue(j, i, v)
				i += 1
			j += 1
		self.filename = filename

	def save(self):
		if self.filename == None:
			print "Error: Please open file."
			return

		out = ""
		bEnd = False
		for j in range(25):
			for i in range(25):
				v = self.GetCellValue(j, i)
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
		print "Save done '%s'"%self.filename


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

		self.log.write("OnSelectCell: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))

		if self.IsCellEditControlEnabled():
			self.HideCellEditControl()
			self.DisableCellEditControl()

		evt.Skip()

class AppFrame(wx.Frame):
	def __init__(self, parent, log):
		wx.Frame.__init__(self, parent, -1, "CSV Editor", size=(640, 480))

		# panel = wx.Panel(self, wx.ID_ANY)
		panel = wx.Panel(self, -1, style=0)
		panel.SetBackgroundColour("#CFCFCF")

		toolbar = wx.ToolBar(panel)
		self.toolbar = toolbar
		toolbar.AddLabelTool(wx.ID_SAVE, "Save", wx.Bitmap("./icons/save.gif"))
		toolbar.Realize()

		self.Bind(wx.EVT_TOOL, self.OnSave, id=wx.ID_SAVE)

		grid = SimpleGrid(panel, log)
		self.grid = grid
		bs = wx.BoxSizer(wx.VERTICAL)
		bs.Add(toolbar)
		bs.Add(grid, 1, wx.GROW|wx.ALL, 5)
		panel.SetSizer(bs)

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

	def OnExit(self, evt):
		self.Close()

	def OnSave(self, evt):
		#self.grid.save()
		self.toolbar.ToggleTool(wx.ID_SAVE, False)
		print "Save done."

if __name__ == "__main__":
	import sys
	from wx.lib.mixins.inspection import InspectableApp
	application = InspectableApp(False)

	frame = AppFrame(None, sys.stdout)
	frame.Show()

	# テスト用コード
	frame.grid.openFile("test.csv")

	application.MainLoop()
