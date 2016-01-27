#!/usr/bin/python
import xlsxwriter
import random
import urllib2
import json
import time
import datetime


'''
The file stock_list format is follows where "A" column is 
reserved for date and the text file should start with B
B : YHOO : NASDAQ
C : CSCO : NASDAQ

The output will be saved in "output_file"

Every weekend we will take a backup of output file and save
file as output_file_backup
'''
stock_list  = '/Users/Desktop/stock_monitor_list.txt'
output_file = '/Users/Desktop/stock_history.xlsx'
output_file_backup =  '/Users/Desktop/stock_history_backup.xlsx'


'''
class GoogleFinanceAPI 
This will provide the closing stock price
'''
class GoogleFinanceAPI:
	def __init__(self):
		self.prefix = "http://finance.google.com/finance/info?client=ig&q="
	
	def get(self,symbol,exchange):
		url = self.prefix+"%s:%s"%(exchange,symbol)
		u = urllib2.urlopen(url)
		content = u.read()
		
		obj = json.loads(content[3:])
		return obj[0]

'''
class excelWriter
This will write the quotes to the Excel Sheet
'''
class excelWriter:
	def __init__(self,filepath):
		self.workbook = xlsxwriter.Workbook(filepath) #filepath
		self.worksheet = self.workbook.add_worksheet()
	def write(self, column, value):
		try:
			# send the location and string to write in Excel sheet
			self.worksheet.write(column, value) 
		except IndexError as err:
			print 'Out of range' + str(err)
		except Exception as err:
			print str(err)
		finally:
			pass

	def get_average_price(self, start_date, end_date, stock_name):
		#Todo
		pass
	def get_lowest_price(self, start_date, end_date, stock_name):
		#Todo
		pass
	def get_highest_price(self, start_date, end_date, stock_name):
		#Todo
		pass

	def close(self):
		self.workbook.close()

if __name__ == "__main__":
	index = 2;
	company_name_index = 1;
	googleApi = GoogleFinanceAPI()
	
	obj1 = excelWriter(output_file)
	
	while True:
		'''
		Excel sheet where the Stock History will be saved
		'''
		column = "A"+str(company_name_index)
		obj1.write(column, "DATE")
		
		try:
			file_stock_list = open(stock_list, 'r')
			'''
			The file format is follows
			B : FEYE : NASDAQ
			C : CSCO : NASDAQ
			'''
			for line in file_stock_list:
				stock = line.split(": ")
				column = stock[0].rstrip() + str(company_name_index)
				obj1.write(column, stock[1])
		except IOError as e:
			print "I/O error({0}): {1}".format(e.errno, e.strerror)
		finally:
			file_stock_list.close()	
		
		today_date= time.strftime("%d/%m/%Y")
		#print today_date
		
		index = index + 1
		today_date= time.strftime("%d/%m/%Y")
		d = datetime.datetime.now()

		if not d.isoweekday() in range(1, 6):
			#We dont want to check  the stock values on Saturday and Sunday
			time.sleep(60*60*24)
			# Saving a backup of the file every week
			shutil.copy(output_file, output_file_backup)
			continue
		
		column = "A" + str(index)
		obj1.write(column, today_date)

		try:
			file_stock_list = open(stock_list, 'r')
			for line in file_stock_list:
				stock = line.split(": ")

				column = stock[0].rstrip() + str(index)
				#print column, stock[0], stock[1], stock[2]
				try:
					#print stock[1], stock[2]
					quote = googleApi.get(stock[1], "NASDAQ")
					print quote['l_cur']
					obj1.write(column, quote['l_cur'])
					time.sleep(1)
				except:
					continue
		except IOError as e:
			print "I/O error({0}): {1}".format(e.errno, e.strerror)
		finally:
			file_stock_list.close()	
		
		#Sleep for 1 day
		time.sleep(60*60*24)
	obj1.close()
