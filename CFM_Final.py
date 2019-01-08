from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
#from selenium.webdriver.common.by import By
#import requests 
import urllib.request
#from bs4 import BeautifulSoup
from html.parser import HTMLParser
import re
import xlsxwriter
from openpyxl import load_workbook


#Login info
user = ""
passwd = ""

#Report info
month = "December"
year = "2018"
startDate ="2018/12/01"
endDate = "2019/01/01"

# Browser : Firefox.
driver = webdriver.Firefox()
driver.get("https://tracks.lufthansa.com/login.jsp")
time.sleep(2)

# Lufthensa login.
username=driver.find_element_by_id("login-form-username")
username.send_keys(user)
time.sleep(2)

#time.sleep(2)
password=driver.find_element_by_id("login-form-password")
password.send_keys(passwd)
time.sleep(2)
login=driver.find_element_by_id("login-form-submit")
login.click()

#redirect in search page
driver.get("https://tracks.lufthansa.com/secure/ConfigurePortalPages!default.jspa#view=search")
time.sleep(2)

#Month selection
search_month = driver.find_element_by_name("searchName")
search_month.send_keys(month)
time.sleep(2)
searchButton = driver.find_element_by_name("Search")
searchButton.click()
time.sleep(2)

#selected month report
create_report = month.capitalize() + " " + year
monthly_details = driver.find_element_by_link_text(create_report)
driver.implicitly_wait(2)
monthly_details.click()
time.sleep(2)

#Select report for CFM
ss = driver.find_elements_by_class_name("legend-item-label")
cfm = "CFM"
sss = driver.find_element_by_link_text(cfm)
for ele in ss :	
	d = ele.text
	if d == cfm :		
		driver.implicitly_wait(2)		
		sss.click()
		break
		
### For CFM 
component_list = ['Bank Registration' ,'BI', 'GRM','OpenText','Payment','Payment Inquiry Agent','Payment Inquiry LGBS','Payment Cancelation by Agent','Payment Rejection by LGBS (SAP)','Payment Status for All','Worldshop','Interfaces','Undelete Feedback', 'Cancellation','CFM']
#create new list for valid one
newList = []

#send details for create a report 
count_list = 6
#create excel
# set file path
filepath=(r"G:\Waqar FAS\Reply\SD Reports\4 Oct 2018\Nov 2018 SD Report Final Version\Dec 2018 SD Report/CFM.xlsx")
# load demo.xlsx 
wb=load_workbook(filepath)
# select demo.xlsx
wb.insert_image('B2', 'python.png')


sheet=wb.active



for component_list_individual in component_list :
	
	
	put = driver.find_element_by_id('advanced-search')
	put.clear()
	a = 'J%i' %count_list
	b = 'K%i' %count_list
	c = 'L%i' %count_list
	if(component_list_individual == "Other"):
		#print(component_list_individual)
		count_list +=1
		newList.append(component_list_individual)
		sheet[a] = component_list_individual
		sheet[b] = 66
	
	else:
		send_keys ='project = Cheetah AND issuetype in (Incident, "Service Request") AND summary !~ hyperic AND Created >= ''"%s"'' AND Created <= ''"%s"'' AND component in (''"%s"'') AND component = CFM AND project = CHEETAH' %(startDate , endDate,  component_list_individual)
		put.send_keys(send_keys)
		driver.find_element_by_css_selector('.aui-item.aui-button.aui-button-subtle.search-button').click()
		time.sleep(5)
		result = driver.find_elements_by_xpath('//span[@class="results-count-total results-count-link"][1]')
		l = len(result)
		#print(l)
		if(l > 0):
			count_list +=1
			#print(component_list_individual)
			
			insert = int(result[0].text)
			if(insert > 0) :
				if(component_list_individual == "CFM"): 
						
						sheet[a] = 'Total'
						sheet[b] = insert
						#deduct += intt
						
				else :
					newList.append(component_list_individual)
					sheet[b] =  insert
					sheet[a] = component_list_individual
				        
					
						#worksheet.write_number(c,  intt*100/per)
					#sum += intt
			else:
			#print('No value for ', component_list_individual)
			 continue

		
t = len(newList) + 3





# For Service Request and Incident

count_list = 290
component_list_s = ['Incident', 'Service Request']
for list in component_list_s:
	a = 'J%i' %count_list
	b = 'K%i' %count_list
	#c = 'Y%i' %count_list
	put = driver.find_element_by_id('advanced-search')
	put.clear()
	time.sleep(5)
	send_keys_SrI = 'project = Cheetah AND type in (Incident, "Service Request") AND type in (Incident, "Service Request") AND Created >= ''"%s"'' AND Created <= ''"%s"'' AND component = CFM AND issuetype = ''"%s"''' %(startDate, endDate, list)
	put.send_keys(send_keys_SrI)
	driver.find_element_by_css_selector('.aui-item.aui-button.aui-button-subtle.search-button').click()
	time.sleep(5)
	results = driver.find_elements_by_xpath('//span[@class="results-count-total results-count-link"][1]')
	l = len(results)
	if(l > 0) :
		count_list +=1
		insert = int(results[0].text)
		#worksheet.write(a,  list)
		#worksheet.write(b,  insert)
		sheet[b] =  insert
		sheet[a] =  list
		

	





##### Service Requests Tickets by CFM

count_list = 419
newList1 = []
for component_list_individual in component_list :
	
	
	put = driver.find_element_by_id('advanced-search')
	put.clear()
	a = 'J%i' %count_list
	b = 'K%i' %count_list
	c = 'L%i' %count_list
	if(component_list_individual == "Other"):
		#print(component_list_individual)
		count_list +=1
		newList1.append(component_list_individual)
		#worksheet.write(a,  component_list_individual)
		#worksheet.write(b,  0)
                #sheet[b] =0
		sheet[a] = component_list_individual

	
	else:
		send_keys ='project = Cheetah AND issuetype in (Incident, "Service Request") AND summary !~ hyperic AND Created >= ''"%s"'' AND Created <= ''"%s"'' AND component in (''"%s"'') AND component = CFM AND issuetype = "Service Request"' %(startDate, endDate, component_list_individual)
		put.send_keys(send_keys)
		driver.find_element_by_css_selector('.aui-item.aui-button.aui-button-subtle.search-button').click()
		time.sleep(5)
		results = driver.find_elements_by_xpath('//span[@class="results-count-total results-count-link"][1]')
		l = len(results)
		#print(l)
		if(l > 0):
			count_list +=1
			#print(component_list_individual)
			
			insert = int(results[0].text)
			if(insert > 0) :
				if(component_list_individual == "CFM"): 
						
						#worksheet.write(a,  "Total", bold)
						#worksheet.write(b,  insert)
						sheet[b] =  insert
						sheet[a] = "Total"
						#deduct += intt
						
				else :
					newList1.append(component_list_individual)
					#worksheet.write(a,  component_list_individual)
					#worksheet.write(b,  insert)
					sheet[b] =  insert
					sheet[a] = component_list_individual	
		else:
			
			continue
		
t = len(newList1) + 80

print("Your Report is created")


			



		
# Create Graph for CFM data





####  	Incident Tickets by CFM

count_list = 628
newListIncident = []
for component_list_individual in component_list :
	
	
	put = driver.find_element_by_id('advanced-search')
	put.clear()
	a = 'J%i' %count_list
	b = 'K%i' %count_list
	c = 'L%i' %count_list
	if(component_list_individual == "Other"):
		#print(component_list_individual)
		count_list +=1
		newListIncident.append(component_list_individual)
		#worksheet.write(a,  component_list_individual)
		#worksheet.write(b,  0)

		sheet[b] = 0
		sheet[a] = component_list_individual
	
	else:
		send_keys ='project = Cheetah AND issuetype in (Incident) AND summary !~ hyperic AND Created >= ''"%s"'' AND Created <= ''"%s"'' AND component in (''"%s"'') AND component = CFM AND issuetype = "Incident"' %(startDate, endDate, component_list_individual)
		put.send_keys(send_keys)
		driver.find_element_by_css_selector('.aui-item.aui-button.aui-button-subtle.search-button').click()
		time.sleep(5)
		results = driver.find_elements_by_xpath('//span[@class="results-count-total results-count-link"][1]')
		l = len(results)
		#print(l)
		if(l > 0):
			count_list +=1
			#print(component_list_individual)
			
			insert = int(results[0].text)
			if(insert > 0) :
				if(component_list_individual == "CFM"): 
						
						#worksheet.write(a,  "Total", bold)
						#worksheet.write(b,  insert)
						sheet[b] =  insert
						sheet[a] ="Total"
						#deduct += intt
						
				else :
					newListIncident.append(component_list_individual)
					#worksheet.write(a,  component_list_individual)
					#worksheet.write(b,  insert)
					sheet[b] =  insert
					sheet[a] = component_list_individual
		else:
			
			continue
	
t = len(newListIncident) + 112

#close excel 
wb.save(filepath)

print("Done")