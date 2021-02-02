#!/usr/bin/env python3
"""
autoIExam: Automate trivial IExam tasks that shouldn't have been needed.

Copyright (C) 2021 Muhammed Shamil K

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import xlrd
import json
import time
import sys

done = []

CONFIG = {}

def populate_config():
	global CONFIG
	with open('config.json') as f:
		CONFIG = json.loads(f.read())

def add_to_done(no):
	global done
	done.append(no)
	with open('done.json', 'w') as f:
		f.write(json.dumps(done))

def read_done():
	global done
	with open('done.json', 'r') as f:
		content = f.read()
		try:
			done = json.loads(content)
		except:
			done = []


driver = webdriver.Firefox(executable_path="./geckodriver" if sys.platform.startswith('linux') else "./geckodriver.exe")
wait = WebDriverWait(driver, 120)
driver.implicitly_wait(120)

def login():
	global tok, CONFIG
	driver.get('https://sampoorna.kite.kerala.gov.in:446/')

	username = wait.until(EC.element_to_be_clickable((By.ID, 'user_username')))
	password = wait.until(EC.element_to_be_clickable((By.ID, 'user_password')))

	username.send_keys(CONFIG['username'])
	password.send_keys(CONFIG['password'])

	submit = driver.find_element_by_name('commit')
	submit.click()

	ibutton = wait.until(EC.element_to_be_clickable((By.ID, 'iexam_button')))
	ibutton.click()

	driver.execute_script('closeModal();')
	tok = driver.execute_script('return tok.value;')
	driver.get(f'https://sslcexam.kerala.gov.in/candidate_registration.php?tok={tok}')

def add_student(admission_no, mal_name, house_name, street, post_office, pin, district):
	add_new_btn = driver.find_element(By.XPATH, "//button[@title='Click to add new candidate']")
	add_new_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Click to add new candidate']")))
	add_new_btn.click()

	admission_input = driver.find_element(By.ID, 'import_adm_no')
	admission_input = wait.until(EC.element_to_be_clickable((By.ID, 'import_adm_no')))
	admission_input.send_keys(admission_no)
	add_button = wait.until(EC.element_to_be_clickable((By.ID, 'import_ok_btn')))
	add_button.click()
	cont_button = wait.until(EC.element_to_be_clickable((By.ID, 'continue_btn')))
	cont_button.click()

	mal_name_input = driver.find_element(By.ID, 'name_mal')
	mal_name_input = wait.until(EC.element_to_be_clickable((By.ID, 'name_mal')))
	mal_name_input.send_keys(mal_name)

	addr1_input = driver.find_element(By.ID, 'address1')
	addr1_input.clear()
	addr1 = "{house_name}, {street}".format(house_name=house_name, street=street)
	addr1_input.send_keys(addr1)

	addr2_input = driver.find_element(By.ID, 'address2')
	addr2_input.clear()
	addr2 = "{post_office} PO, {pin}, {district}".format(post_office=post_office, pin=pin, district=district)
	addr2_input.send_keys(addr2)


	save_btn = driver.find_element(By.ID, 'save_btn')
	save_btn = wait.until(EC.element_to_be_clickable((By.ID, 'save_btn')))
	save_btn.click()

	submit_btn = driver.find_element(By.ID, 'final_save_btn')
	submit_btn = wait.until(EC.element_to_be_clickable((By.ID, 'final_save_btn')))
	submit_btn.click()
	add_to_done(admission_no)
	cont_n_btn = driver.find_element(By.CSS_SELECTOR, '.sweet-alert .confirm')
	cont_n_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.sweet-alert .confirm')))
	cont_n_btn.click()
	#time.sleep(5)

def process_report():
	global done
	book = xlrd.open_workbook('report.xls')
	for row in book[0]:
		adm = row[0].value
		if not adm.isnumeric():
			continue
		if adm in done:
			continue
		mal_name = row[1].value.strip(' .,')
		house_name = row[2].value.strip(' .,')
		street_name = row[3].value.strip(' .,')
		post_office = row[4].value.strip(' .,')
		pin_code = row[5].value.strip(' .,')
		district = row[6].value.strip(' .,')
		add_student(adm, mal_name, house_name, street_name, post_office, pin_code, district)
		

def main():
	try:
		populate_config()
		read_done()
		login()
		process_report()
	finally:
		driver.quit()

if __name__ == '__main__':
	main()