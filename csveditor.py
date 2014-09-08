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

class PopupMenu(wx.Menu):
	def __init__(self, parent, rowCount, colCount):
		wx.Menu.__init__(self)
		self.parent = parent

		insMenu = "Insert"
		delMenu = "Delete"
		if rowCount > 0:
			insMenu += " %d rows"%(rowCount)
			delMenu += " %d rows"%(rowCount)
		if colCount > 0:
			insMenu += " %d cols"%(colCount)
			delMenu += " %d cols"%(colCount)

		item = wx.MenuItem(self, wx.NewId(), insMenu)
		self.AppendItem(item)
		self.Bind(wx.EVT_MENU, self.Insert, id=item.GetId())
		item = wx.MenuItem(self, wx.NewId(), delMenu)
		self.AppendItem(item)
		self.Bind(wx.EVT_MENU, self.Delete, id=item.GetId())

	def Insert(self, evt):
		self.parent.Insert(evt)

	def Delete(self, evt):
		self.parent.Delete(evt)

class RangeSelect:
	def __init__(self):
		self.rowStart = -1
		self.colStart = -1
		self.rowCount = 1
		self.colCount = 1
		self.rowSelect = -1
		self.colSelect = -1
	def insert(self, grid):
		if self.rowSelect >= 0:
			if self.rowCount == 0:
				grid.InsertRows(self.rowSelect, 1)
			else:
				grid.InsertRows(self.rowStart, self.rowCount)
		if self.colSelect >= 0:
			if self.colCount == 0:
				grid.InsertCols(self.colSelect, 1)
			else:
				grid.InsertCols(self.colStart, self.colCount)
	def delete(self, grid):
		if self.rowSelect >= 0:
			if self.rowCount == 0:
				grid.DeleteRows(self.rowSelect, 1)
			else:
				grid.DeleteRows(self.rowStart, self.rowCount)
		if self.colSelect >= 0:
			if self.colCount == 0:
				grid.DeleteCols(self.colSelect, 1)
			else:
				grid.DeleteCols(self.colStart, self.colCount)
	def setRange(self, topRow, bottomRow, leftCol, rightCol):
		if topRow == bottomRow:
			if self.rowStart >= 0:
				if self.rowStart > topRow:
					self.rowCount = self.rowStart - topRow + 1
					self.rowStart = topRow
				else:
					self.rowCount = bottomRow - self.rowStart + 1
			else:
				self.rowStart = topRow
				self.rowCount = 0
		else:
			self.rowStart = -1
			self.rowCount = 0

		if leftCol == rightCol:
			if self.colStart >= 0:
				if self.colStart > leftCol:
					self.colCount = self.colStart - leftCol + 1
					self.colStart = leftCol
				else:
					self.colCount = rightCol - self.colStart + 1
			else:
				self.colStart = leftCol
				self.colCount = 0
		else:
			self.colStart = -1
			self.colCount = 0

class HistoryMgr:
	""" 履歴管理 """
	def __init__(self, parent):
		self.buffers = []
		self.grid = parent
	def push(self, row, col, val):
		item = "%d,%d,%s"%(row, col, val)
		self.buffers.append(item)
	def pop(self):
		items = self.buffers.pop().split(",")
		return (int(items[0]), int(items[1]), items[2])
	def length(self):
		return len(self.buffers)
	def undo(self):
		if self.length() <= 0:
			return

		# 履歴から取り出し
		row, col, val = self.pop()
		# 値を戻す
		self.grid.SetCellValue(row, col, val)
		# 変更したセルにカーソルを移動する
		self.grid.SetGridCursor(row, col)

	def clear(self):
		""" 履歴をすべて消す """
		while self.length() > 0:
			self.pop()

class SimpleGrid(gridlib.Grid):
	WIDTH = 25
	HEIGHT = 100
	def __init__(self, parent, frame):
		gridlib.Grid.__init__(self, parent, -1)
		self.parent = parent
		self.frame = frame
		self.moveTo = None
		self.panel = wx.Panel(parent, -1)

		# 範囲選択管理
		self.rangeSelect = RangeSelect()

		# 履歴管理
		self.histories = HistoryMgr(self)

		self.Bind(wx.EVT_IDLE, self.OnIdle)

		self.SetDropTarget(GridFileDropTarget(self))

		# ひとまず WIDTHxHEIGHT 固定とする
		self.CreateGrid(self.HEIGHT, self.WIDTH)

		# バインド
		self.Bind(wx.EVT_IDLE, self.OnIdle)

		self.Bind(gridlib.EVT_GRID_CELL_LEFT_CLICK, self.OnCellLeftClick)
		self.Bind(gridlib.EVT_GRID_CELL_RIGHT_CLICK, self.OnCellRightClick)
		self.Bind(gridlib.EVT_GRID_LABEL_LEFT_CLICK, self.OnLabelLeftClick)
		self.Bind(gridlib.EVT_GRID_LABEL_RIGHT_CLICK, self.OnLabelRightClick)

		self.Bind(gridlib.EVT_GRID_CELL_CHANGE, self.OnCellChange)
		self.Bind(gridlib.EVT_GRID_SELECT_CELL, self.OnSelectCell)
		self.Bind(gridlib.EVT_GRID_RANGE_SELECT, self.OnRangeSelect)

		self.Bind(gridlib.EVT_GRID_EDITOR_SHOWN, self.OnEditorShown)
		self.Bind(gridlib.EVT_GRID_EDITOR_HIDDEN, self.OnEditorHidden)
		self.Bind(gridlib.EVT_GRID_EDITOR_CREATED, self.OnEditorCreated)

		self.filename = None

		self.log = sys.stdout

		self.firsts = {}

		# 何か変更点があるかどうか
		self.bChange = False

	def isChange(self):
		return self.bChange

	def openFile(self, filename):
		self.bChange = False
		self.firsts = {}
		for j in range(self.HEIGHT):
			for i in range(self.WIDTH):
				# セルの色を元に戻す
				self.SetCellBackgroundColour(j, i, wx.WHITE)
				# 値も消す
				self.SetCellValue(j, i, "")

		j = 0
		for line in codecs.open(filename, "r", "utf-8"):
			i = 0

			line = line.rstrip()
			data = line.split(",")
			for v in data:
				self.SetCellValue(j, i, v)
				# 初期値を保存しておく
				self.firsts[i+j*self.WIDTH] = v
				i += 1
			j += 1
		self.filename = filename
		self.SetStatusText("Open file '%s'."%filename)

		# テキストに合わせて自動リサイズする
		self.AutoSize()

	def saveas(self):
		""" 名前をつけて保存 """
		dirName = ''
		dialog = wx.FileDialog(self, "Input a save file", dirName, "", "*.csv", wx.FD_SAVE)
		result = dialog.ShowModal()
		dialog.Destroy()
		if result == wx.ID_OK:
			fileName = dialog.GetFilename()
			dirName = dialog.GetDirectory()
			self.filename = dirName + "/" + fileName

	def save(self, bSaveAs=False):
		if self.filename == None or bSaveAs:
			self.saveas()
			if self.filename == None:
				self.SetStatusText("Error: Please open file.")
				return


		self.bChange = False
		self.firsts = {}
		out = ""
		bEnd = False
		nColMax = 0
		for j in range(self.HEIGHT):
			for i in range(self.WIDTH):
				v = self.GetCellValue(j, i)
				self.firsts[i + j*self.WIDTH] = v
				# セルの色を元に戻す
				self.SetCellBackgroundColour(j, i, wx.WHITE)
				if v == "":
					if i == 0:
						# 1列目に何もなければ終了
						bEnd = True
					elif i > nColMax:
						# 最大列を保存
						nColMax = i
					else:
						# 足りない列だけカンマを付加
						col = nColMax - i
						out += "," * col
					# この行の処理はおしまい
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
		print("Save as '%s'."%self.filename)

	def Insert(self, evt):
		self.rangeSelect.insert(self)
		# 履歴を消す
		self.histories.clear()

	def Delete(self, evt):
		self.rangeSelect.delete(self)
		# 履歴を消す
		self.histories.clear()

	def Cut(self):
		if wx.TheClipboard.Open():
			val = self.Cells()
			wx.TheClipboard.SetData(wx.TextDataObject(val))
			wx.TheClipboard.Flush()
			wx.TheClipboard.Close()
			self.SetCell("")

	def Copy(self):
		if wx.TheClipboard.Open():
			val = self.Cells()
			wx.TheClipboard.SetData(wx.TextDataObject(val))
			wx.TheClipboard.Flush()
			wx.TheClipboard.Close()

	def Paste(self):
		if wx.TheClipboard.Open():
			do = wx.TextDataObject()
			wx.TheClipboard.GetData(do)
			val = do.GetText()
			wx.TheClipboard.Close()
			if val != "":
				self.SetCell(val)

	def OnCellLeftClick(self, evt):
		self.log.write("OnCellLeftClick: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnCellRightClick(self, evt):
		self.log.write("OnCellRightClick: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnLabelLeftClick(self, evt):
		# ラベル左クリック
		evt.Skip()

	def OnLabelRightClick(self, evt):
		# ラベル右クリック
		self.rangeSelect.rowSelect = evt.GetRow()
		self.rangeSelect.colSelect = evt.GetCol()

		# ボップアップメニュー表示
		self.PopupMenu(PopupMenu(self, self.rangeSelect.rowCount, self.rangeSelect.colCount), evt.GetPosition())

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

		# 変更点チェック
		self.checkDiff(evt)

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

		col = chr(65+evt.GetCol())
		msg = "%s:%d %s"%(col, evt.GetRow(), self.Cells(evt))
		self.SetStatusText(msg)

		evt.Skip()

	def OnRangeSelect(self, evt):
		# 範囲選択
		print("OnRangeSelect: Row %d:%d Col %d:%d\n" %
			(evt.GetTopRow(), evt.GetBottomRow(), evt.GetLeftCol(), evt.GetRightCol()))
		self.rangeSelect.setRange(evt.GetTopRow(), evt.GetBottomRow(), evt.GetLeftCol(), evt.GetRightCol())

		evt.Skip()

	def Cells(self, evt=None):
		# 現在のセルの値を取得する
		if evt == None:
			row = self.GetGridCursorRow()
			col = self.GetGridCursorCol()
			return self.GetCellValue(row, col)
		else:
			return self.GetCellValue(evt.GetRow(), evt.GetCol())

	def GetCell(self, row, col):
		# 指定のセルの値を取得する
		return self.GetCellValue(row, col)

	def SetCell(self, val, row=-1, col=-1):
		# セルに値を設定する
		if row < 0:
			row = self.GetGridCursorRow()
		if col < 0:
			col = self.GetGridCursorCol()

		if val == self.GetCell(row, col):
			# 変更不要
			return

		# 履歴に追加
		self.histories.push(row, col, self.GetCell(row, col))
		return self.SetCellValue(row, col, val)

	def Firsts(self, evt):
		# 初期値を取得する
		idx = evt.GetRow() * self.WIDTH+ evt.GetCol()
		if idx in self.firsts:
			return self.firsts[idx]
		else:
			return ""

	def checkDiff(self, evt):
		# 初期値と相違があればセルの色を変える。そうでなければ白に戻す
		color = wx.WHITE
		if self.Cells(evt) != self.Firsts(evt):
			color = wx.Colour(255, 211, 255)
			self.bChange = True # 何か変更があった

		self.SetCellBackgroundColour(evt.GetRow(), evt.GetCol(), color)
	def OnEditorShown(self, evt):
		# セルのエディット開始
		self.checkDiff(evt)
		self.log.write("OnEditorShown: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnEditorHidden(self, evt):
		# セルのエディット終了
		# 履歴に追加
		self.histories.push(evt.GetRow(), evt.GetCol(), self.Cells(evt))
		self.checkDiff(evt)
		self.log.write("OnEditorHidden: (%d,%d) %s\n" %
			(evt.GetRow(), evt.GetCol(), evt.GetPosition()))
		evt.Skip()

	def OnEditorCreated(self, evt):
		# セルのエディット（初回のみ）
		self.log.write("OnEditorCreated: (%d,%d) %s\t" %
			(evt.GetRow(), evt.GetCol(), evt.GetControl()))

	def SetStatusText(self, msg):
		self.frame.statusbar.SetStatusText("[%s] %s"%(self.filename, msg))

	def checkSave(self):
		# セーブ終了確認
		if self.isChange():
			result = wx.MessageBox("Are you sure you want to save change?", "Confirm Save", wx.OK | wx.CANCEL | wx.ICON_QUESTION)
			if result == wx.OK:
				self.save()

class AppFrame(wx.Frame):
	def __init__(self, parent, log):
		wx.Frame.__init__(self, parent, -1, "CSV Editor", size=(640, 480))

		# パネル生成
		panel = wx.Panel(self, -1, style=0)
		panel.SetBackgroundColour("#CFCFCF")

		# ツールバー生成
		toolbar = wx.ToolBar(panel)
		self.toolbar = toolbar
		toolbar.AddLabelTool(wx.ID_SAVE, "Save", wx.Bitmap("./save.gif")) # 保存アイコン追加
		toolbar.AddLabelTool(wx.ID_OPEN, "Open", wx.Bitmap("./open.gif")) # 開くアイコン
		toolbar.AddLabelTool(wx.ID_UNDO, "Undo", wx.Bitmap("./undo.gif")) # Undoアイコン
		toolbar.Realize()

		self.Bind(wx.EVT_TOOL, self.OnSave, id=wx.ID_SAVE)
		self.Bind(wx.EVT_TOOL, self.OnOpen, id=wx.ID_OPEN)
		self.Bind(wx.EVT_TOOL, self.OnUndo, id=wx.ID_UNDO)

		# グリッド生成
		grid = SimpleGrid(panel, self)
		self.grid = grid

		# メニューバー生成
		menu_file = wx.Menu()
		menu_open = wx.MenuItem(menu_file, 1, u"&Open\tCtrl+O")
		menu_save = wx.MenuItem(menu_file, 2, u"&Save\tCtrl+S")
		menu_saveas = wx.MenuItem(menu_file, 3, u"&Save As...\tShift+Ctrl+S")
		menu_file.AppendItem(menu_open)
		menu_file.AppendItem(menu_save)
		menu_file.AppendItem(menu_saveas)
		menu_file.AppendSeparator()
		menu_exit = wx.MenuItem(menu_file, 4, "&Quit", "Quit CsvEditor")
		menu_file.AppendItem(menu_exit)

		self.Bind(wx.EVT_MENU, self.OnOpen, menu_open)
		self.Bind(wx.EVT_MENU, self.OnSave, menu_save)
		self.Bind(wx.EVT_MENU, self.OnSaveAs, menu_saveas)
		self.Bind(wx.EVT_MENU, self.OnExit, menu_exit)

		menu_edit  = wx.Menu()
		menu_undo  = wx.MenuItem(menu_edit, 5, "&Undo\tCtrl+Z")
		menu_cut   = wx.MenuItem(menu_edit, 6, "&Cut\tCtrl+X")
		menu_copy  = wx.MenuItem(menu_edit, 7, "&Copy\tCtrl+C")
		menu_paste = wx.MenuItem(menu_edit, 8, "&Paste\tCtrl+V")
		menu_edit.AppendItem(menu_undo)
		menu_edit.AppendItem(menu_cut)
		menu_edit.AppendItem(menu_copy)
		menu_edit.AppendItem(menu_paste)

		self.Bind(wx.EVT_MENU, self.OnUndo, menu_undo)
		self.Bind(wx.EVT_MENU, self.OnCut, menu_cut)
		self.Bind(wx.EVT_MENU, self.OnCopy, menu_copy)
		self.Bind(wx.EVT_MENU, self.OnPaste, menu_paste)

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

		self.Bind(wx.EVT_CLOSE, self.OnExit)

	def OnExit(self, evt):
		self.grid.checkSave()
		# self.Close()
		wx.Exit()

	def OnOpen(self, evt):
		self.grid.checkSave()

		dirName = ''
		# ファイル選択ダイアログの表示
		dialog = wx.FileDialog(self, "Choose a file", dirName, "", "*.csv", wx.OPEN)

		# OKボタンが押されるまで表示
		if dialog.ShowModal() == wx.ID_OK:
			fileName = dialog.GetFilename()
			dirName = dialog.GetDirectory()
			self.grid.openFile(dirName + "/" + fileName)
		dialog.Destroy()

	def OnSave(self, evt):
		self.grid.save()

	def OnSaveAs(self, evt):
		self.grid.save(True)

	def OnUndo(self, evt):
		self.grid.histories.undo()

	def OnCut(self, evt):
		self.grid.Cut()

	def OnCopy(self, evt):
		self.grid.Copy()

	def OnPaste(self, evt):
		self.grid.Paste()

if __name__ == "__main__":
	import sys
	from wx.lib.mixins.inspection import InspectableApp
	application = InspectableApp(False)

	frame = AppFrame(None, sys.stdout)
	frame.Show()

	if len(sys.argv) > 1:
		# 起動引数があればファイルを開く
		frame.grid.openFile(sys.argv[1])
	# テスト用コード
	# frame.grid.openFile("test.csv")

	application.MainLoop()
