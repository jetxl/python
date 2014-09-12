import os
import sys
import subprocess
import PcTest
import AndroidTest
import xlrd

print "Hello, World!";	
def main():
	xlsx = excel.EasyExcel(os.path.join(os.getcwd(),'结果.xlsx'))
	xlsx.setCell('Sheet1', 1, 1,'呵呵')
	xlsx.save()
	xlsx.close()

if __name__ == '__main__':
	#PcTest.main()
	AndroidTest.main()