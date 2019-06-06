from bs4 import BeautifulSoup
from urllib.request import urlopen
import xlwt
import requests
import openpyxl
import xlrd
from xlutils.copy import copy

line_in_list = ['https://www.ebay.co.uk/sch/i.html?_odkw=the+shining+1977&LH_Complete=1&LH_Sold=1&_osacat=0&_from=R40&_trksid=m570.l1313&_nkw=the+shining+doubleday+1977&_sacat=0',
 'https://www.ebay.co.uk/sch/i.html?_odkw=the+shining+doubleday+1977&LH_Complete=1&LH_Sold=1&_osacat=0&_from=R40&_trksid=m570.l1313&_nkw=catcher+in+the+rye+first+edition&_sacat=0'
] 
books_list = ['the shining 1977','Catcher in the rye first edition']

sheet = [None] * len(line_in_list)

crawler = xlwt.Workbook(encoding='utf-8', style_compression = 0)
for count,books in enumerate(line_in_list,0):
	sheet[count] = crawler.add_sheet('Sheet' + str(count), cell_overwrite_ok = True)

for cor,websites in enumerate(line_in_list):
	i = 1
	j = 1
	url = websites	
	response = requests.get(url)
	soup = BeautifulSoup(response.text, 'html.parser')

	for price_box in soup.findAll('span', attrs={'class': 'bold bidsold'}):
		price = price_box.text.strip()
		sheet[cor].write(0,0,'Sold For')
		sheet[cor].write(i,0,price)
		i=i+1
	
	for date_box in soup.findAll('span', attrs={'class': 'tme'}):
		date = date_box.text.strip()
		sheet[cor].write(0,1,'Sold Date')
		sheet[cor].write(j,1,date)
		j=j+1
		
crawler.save("title.xls")


rb = xlrd.open_workbook('title.xls')
wb = copy(rb)

for count,books in enumerate(books_list,0):
	idx = rb.sheet_names().index('Sheet' + str(count))
	wb.get_sheet(idx).name = books
	wb.save('title.xls')

	
