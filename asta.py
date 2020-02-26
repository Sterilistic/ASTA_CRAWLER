from selenium import webdriver
from openpyxl.workbook import Workbook
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re
import pandas as pd
import os, random, time
from openpyxl import Workbook
import sys

url = "https://web.asta.org/iMIS/ASTA/Contacts/People_Search.aspx"

desktop_agent = ['Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_4) AppleWebKit/600.7.12 (KHTML, like Gecko) Version/8.0.7 Safari/600.7.12']
rand_agent = desktop_agent[random.randrange(0,len(desktop_agent))]
profile = webdriver.FirefoxProfile()
profile.set_preference("general.useragent.override", rand_agent)
_driver = webdriver.Firefox(profile)


_driver.implicitly_wait(30)
_driver.get(url)

find_button = _driver.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_sbtnSearch"]')
find_button.click()
_driver.implicitly_wait(30)
page_numbers = _driver.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_gvResults"]/tbody/tr[1]/td/table/tbody/tr')


members_name = []
members_email = []
members_mobile = []
members_address = []
members_state =[]



def get_table_data(page):
	"""
	scrape table data(member name & portfolio link) from source url 
	"""
	while page <=1:
		time.sleep(5)
		table = _driver.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_gvResults"]/tbody')
		time.sleep(5)
		count = 0
		print(page)
		for row in table.find_elements_by_xpath('./tr'):
			_driver.implicitly_wait(30)
			time.sleep(3) 
			try:
				if row.get_attribute("class") == "cssPager":
					count = count + 1

				if count == 2:
					page = page + 1
					link1 = "javascript:__doPostBack('ctl01$TemplateBody$WebPartManager1$gwpciBPDirectorySearch$ciBPDirectorySearch$gvResults'"
					link2 = ",'Page${0}')".format(str(page))
					next_page_link = link1 + link2
					_driver.get(next_page_link)
					time.sleep(5)
					get_table_data(page)

				member_name = row.find_element_by_xpath('./td').text	
				member_link = row.find_element_by_link_text(member_name).get_attribute('href')
				print(member_link)
				get_profile_data(member_name, member_link)
				

			except Exception as e:
				continue


# get_table_data(page=1)



def get_profile_data(member_name, member_link):
	"""
	fetch member's profile details
	"""
	
	members_name.append(member_name)
	_driver = webdriver.Firefox(profile)
	_driver.close()
	_driver = webdriver.Firefox(profile)
	time.sleep(5)
	_driver.get(member_link)
	time.sleep(5)
	address = _driver.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__divAddress"]/div/div[2]')
	address_list = address.find_elements_by_xpath('.//span[@id = "ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__address"]')[0].get_attribute("innerHTML").split('<br>')
	members_address.append(address_list[0])
	members_state.append(address_list[1])

	try:
		contact_number = address.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__phoneNumber"]').text
		members_mobile.append(contact_number)
		
	except Exception as e:
		members_mobile.append('')
		pass
	try:
		email = address.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__email"]').text
		members_email.append(email)
	except Exception as e:
		members_email.append('')
		pass

	_driver.close()
	_driver.get(url)
	time.sleep(5)
	find_button = _driver.find_element_by_xpath('//*[@id="ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_sbtnSearch"]')
	find_button.click()
	_driver.implicitly_wait(30)
 

def prepare_excel():
	"""

	"""
	data = [
		{"Name":members_name},
		{"Email":members_email},
		{"Mobile":members_mobile},
		{"Address":members_address},
		{"State":members_state}
	]
	wb = Workbook()
	ws1 = wb.active
	file_extention = 'asta_data' + '.xlsx'
	file_path = sys.path[0] + "/" + file_extention
	excel_file = wb.save(file_path)
	final_df = pd.DataFrame()
	for id in range(0, len(data)):
		df = pd.DataFrame.from_dict(data[id])
		final_df = pd.concat([final_df, df], axis=1)

	final_df.to_excel(file_extention)


if __name__ == '__main__':
	get_table_data(page=1)
	prepare_excel()
	# get_profile_data('https://web.asta.org/imis/profile?ID=900276578&firstName=Alexandra&lastName=Azpurua&companyName=')	







