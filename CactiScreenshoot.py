'''
disini nambahin untuk export ke excel sekaligus
'''
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from PIL import Image
import time
import os
from getpass import getpass
import xlsxwriter

user_login = input("Username: ")
pass_login = getpass()
daftar_site_input = input("Input site list(site1.com,site2.com,site3.com): ")
daftar_site = daftar_site_input.split(",")

print("\n\n")
print("url format: ")
print("no;nama_remote;SID;link_MRTG")
print("\n\n")

while True:
	url_file = input("Input url file: ")
	urls = open(url_file,"r").readlines()
	if len(urls[0].split(";")) == 4:
		break
	else:
		print("Format is wrong, chose another file!")
		continue

start_date = input("Input start date (yyyy-mm-dd hh:mm): ")
end_date = input("Input end date (yyyy-mm-dd hh:mm): ")

while True:
	try:
		folder = input("Input folder name to save the result: ")
		os.mkdir(folder)
		break
	except FileExistsError:
		print("Folder exist! Chose another folder!")
		continue
file_excel = input("Input excel name: ")

workbook = xlsxwriter.Workbook("{}/{}".format(folder,file_excel))
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
worksheet.set_default_row(112)
worksheet.set_column('E:E', 42)


DRIVER = 'chromedriver'
driver = webdriver.Chrome(DRIVER)

for site in daftar_site:

	driver.get('http://{}'.format(site.strip()))
	username = driver.find_element_by_name('login_username')
	password = driver.find_element_by_name('login_password')

	username.send_keys(user_login)
	password.send_keys(pass_login)

	driver.find_element_by_xpath("//input[@value='Login']").click()

	try: #kasih try biar kalau gagal login kan nda ada element date1 dll
		date1 = driver.find_element_by_name('date1')
		date2 = driver.find_element_by_name('date2')
		date1.clear()
		date2.clear()
		date1.send_keys(tanggal_mulai)
		date2.send_keys(tanggal_ahir)
		try: #try ini artinya jika ada Refresh, kalau nda ada berarti refresh
			driver.find_element_by_xpath("//input[@value='Refresh']").click()
		except:
			driver.find_element_by_xpath("//input[@value='refresh']").click()
			continue

	except:
		continue

for url in urls:
	nama_hasil = url.split(";")[2]
	url_nya = url.split(";")[3]
	if "http" in url_nya:
		driver.get(url_nya)
		screenshot = driver.save_screenshot('{}/{}a.png'.format(folder, nama_hasil))
		img = Image.open('{}/{}a.png'.format(folder, nama_hasil))
		img2 = img.crop((303, 206, 903, 453))#305, 208, 905, 455
		os.remove('{}/{}a.png'.format(folder, nama_hasil))
		img2.save('{}/{}.png'.format(folder, nama_hasil))

driver.quit()
print("\n\nInput the picture to excel. Whait a minutes.......")

worksheet.write("A5", "No", bold)
worksheet.write("B5", "Nama Remote", bold)
worksheet.write("C5", "SID", bold)
worksheet.write("D5", "Link MRTG", bold)
worksheet.write("E5", "Capture MRTG", bold)

row = 6
col = 0

for url in urls:
	no = url.split(";")[0]
	nama_remote = url.split(";")[1]
	sid = url.split(";")[2]
	link_mrtg = url.split(";")[3]

	worksheet.write(row, col, no)
	worksheet.write(row, col+1, nama_remote)
	worksheet.write(row, col+2, sid)
	worksheet.write(row, col+3, link_mrtg)
	worksheet.insert_image(row, col+4, "{}/{}.png".format(folder, sid), {'x_scale': 0.5, 'y_scale': 0.6})
	row = row + 1
workbook.close()
print("\n\nProgram finished!\n\n")
