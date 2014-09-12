
#excelWin32Com.py 
#Author by kiyou yang

from win32com.client import Dispatch
import win32com.client 

class EasyExcel:
	"a class for to solve excel problem"
	def __init__(self, fileName = None):
		self.xlApp = win32com.client.Dispatch('Excel.Application')
		if fileName:
			self.fileName = fileName
			self.xlBook = self.xlApp.Workbooks.Open(fileName)
		else:
			self.xlBook = self.xlApp.Workbooks.Add()
			self.fileName = ''

	def save(self, newFileName=None):
		if newFileName:
			self.fileName = newFileName
			self.xlBook.SaveAs(newFileName)
		else:
			self.xlBook.Save()

	def close(self):
		self.xlBook.Close(SaveChanges=0)
		self.xlApp.Application.Quit()
		del self.xlApp

	def getCell(self, sheet, row, column):
		"get cell value"
		sht = self.xlBook.Worksheets(sheet)
		return sht.Cells(row,column).Value

	def setCell(self, sheet, row, column, value):
		"set cell value"
		sht = self.xlBook.Worksheets(sheet)
		sht.Cells(row, column).Value = value

	def getRange(self, sheet, row1, column1, row2, column2):
		"get a 2d array"
		sht = self.xlBook.Worksheets(sheet)
		return sht.Range(sht.Cells(row1, column1), sht.Cells(row2, column2)).Value

	def copy(self, beforre):
		"copy sheet"
		shts = self.xlBook.Worksheets
		shts(1).Copy(None, shts(1))