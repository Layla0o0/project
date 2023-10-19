# project
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait ,Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException, NoSuchWindowException
import time
import os
import pandas as pd



# tinker symbol for 9 banks
key_words= ['VCB','VIB','VBB','TPB','TCB','MBB','CTG','ACB','BID']

for key_word in key_words:
	print(f" Scraping {key_word} ")
	
	try:
		driver =  webdriver.Chrome()
		driver.get('https://cafef.vn')
		element = driver.find_element(By.ID, 'CafeF_SearchKeyword_Company')
		element.clear()
		element.send_keys(key_word)
		element.send_keys(Keys.ENTER)
		
		try:
			reports = WebDriverWait(driver,20).until(
				EC.presence_of_element_located((By.LINK_TEXT,'Xem đầy đủ'))
				)
			reports.click()
			
			try:
				year = WebDriverWait(driver,10).until(
					EC.presence_of_element_located((By.ID,'rdo0'))
					)
				year.click()
				

			except (NoSuchElementException,ElementNotInteractableException) as e:
				print('there is an error in step 3', e)
				continue
			else:
				categories = ['//*[@id="aNhom1"]','//*[@id="aNhom2"]','//*[@id="ContentPlaceHolder1_aNhom3"]','//*[@id="aNhom4"]'] 
				sheet_names = ["Balanse sheet", "Income Statement", "Indirect Cash Flow", "Cash Flow"]
				df_list=[]
				for category in categories:
					print(f"Processing report : {category}")

					try:
						tab = driver.find_element(By.XPATH, f'{category}')
						tab.click()

			
					except(NoSuchElementException,ElementNotInteractableException) as e:
						print('there is an error in step 4.1', e)
						continue
					else:
						df = pd.read_html(driver.page_source)
						column_name = list(df[3].iloc[0, 1:5])
						column_name.insert(0, 'Chỉ tiêu')
						report_df = df[4].iloc[:,0:5]                        
						report_df.columns = column_name
						df_list.append(report_df)                        
				for ti in range(3):
					for _ in range(4):
						before = driver.find_element(By.XPATH, '//*[@id="tblGridData"]/tbody/tr/td[1]/div/a[1]')
						before.click()
						i=0                        
					for category in categories:
						try:
							tab = driver.find_element(By.XPATH, f'{category}')
							tab.click()
							print(f"Scraping {category} for {ti+1} times")
							df1 = pd.read_html(driver.page_source)
							column_name1 = list(df1[3].iloc[0, 1:5])
							report_df1 = df1[4].iloc[:,1:5]
							report_df1.columns = column_name1
							for t in range(3,0,-1):
								df_list[i].insert(1,report_df1.columns[t],report_df1.iloc[:,t])
								print(f'insert {category} data for {ti+1}')
							i+=1

						except NoSuchElementException as e:
							print('there is error in step 4.2.0')




		
		except (NoSuchElementException,ElementNotInteractableException,TimeoutException) as e:
			print('there is an error in step 2', e)
			continue
	
	except (NoSuchElementException,ElementNotInteractableException,NoSuchWindowException) as e :
		print('there is an error in step 1', e)

		continue
	else:
		# create a excel file
		xlwriter = pd.ExcelWriter(f"{key_word}"+".xlsx")
		for k in range(4):
			temporary_df = df_list[k].iloc[::3]
			temporary_df.to_excel(xlwriter, sheet_name = sheet_names[k],index = False)
		xlwriter.save()	
		
	finally:
		
		driver.quit()


						
						
						
						
						
						
