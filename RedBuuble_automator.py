import tkinter as tk
from tkinter import filedialog, Text
from  tkinter import *
import os
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from selenium.common.exceptions import TimeoutException



root = tk.Tk()
apps = []

def AddFile():

	for widget in frame6.winfo_children():
		widget.destroy()

	global filename

	filename = filedialog.askopenfilename(initialdir="/", title= "Select File",
		filetypes=(("excel files","*.xlsx"),("all files", "*.*")))
	apps.append(filename)

	for app in apps:
		label= tk.Label(frame6, text=app, bg="gray")
		label.pack()

def Start():
	
	# set file path
	filepath=filename
	# load excel-example.xlsx
	wb=load_workbook(filepath)
	# activate demo.xlsx
	sheet=wb.active
	# get b1 cell value
	PATH = "/Users/ghost/Python/chromeDriver/chromedriver"
	driver = webdriver.Chrome(PATH)
	driver.get("https://www.redbubble.com/auth/login")

	# Store google search box WebElement
	
	try:
		nb = int(maxupload.get())
	except ValueError:
		print("erreur")

	driver.find_element_by_id("ReduxFormInput1").send_keys(sheet["A2"].value)
	driver.find_element_by_id("ReduxFormInput2").send_keys(sheet["B2"].value)
	time.sleep(90)
	for i in range(2,nb):
		rowC = 'C{0}'.format(i)
		rowD ='D{0}'.format(i)
		rowE = 'E{0}'.format(i)
		rowF = 'F{0}'.format(i)
		driver.get(sheet[rowC].value)
		driver.find_element(By.ID, "work_title_en").clear()
		driver.find_element(By.ID, "work_tag_field_en").clear()
		driver.find_element(By.ID, "work_title_en").send_keys(sheet[rowD].value)
		driver.find_element(By.ID, "work_tag_field_en").send_keys(sheet[rowE].value)
		driver.find_element(By.ID, "rightsDeclaration").click()
		driver.find_element(By.ID, "select-image-base").send_keys(sheet[rowF].value)
		time.sleep(30)
		driver.find_element(By.ID, "submit-work").click()
		time.sleep(30)
		driver.execute_script("window.open('about:blank', 'tab');")
		driver.switch_to.window("tab")
		url= sheet[rowC].value
		driver.get(url)


canvas = tk.Canvas(root, height=500, width=500, bg="#FFE1CF")
canvas.pack()



frame = tk.Frame(root, bg="#2E2B2B")
frame.place(relwidth= 1, relheight= 0.05)
label = Label(frame, text = "Welcome To RedBubble Automator", bg= "#2E2B2B", fg= "white",font="Times 18")
label.pack()

frame2 = tk.Frame(root, bg="#FDC5AB")
frame2.place(relwidth= 0.9, relheight= 0.05, rely= 0.2, relx= 0.04)
label = Label(frame2, text = "Settings", bg= "#FDC5AB", fg= "#2E2B2B")
label.pack()

frame3 = tk.Frame(root, bg="#2E2B2B")
frame3.place(relwidth= 1, relheight=0.05, rely= 0.95)
label = Label(frame3, text = "Welcome to my First version of my redbubble Automator", bg= "#2E2B2B", fg= "white", font="Times 15")
label.pack()

frame4 = tk.Frame(root, bg="#FDC5AB")
frame4.place(relwidth= 0.9, relheight= 0.38, rely= 0.35, relx= 0.04)
label = Label(frame4, text = "Upload xlsx File", bg= "#FDC5AB", fg= "#2E2B2B",width=20,height=4,anchor=W)
label.grid(row=0,column=0)
label2 = Label(frame4, text = "Max Upload", bg= "#FDC5AB", fg= "#2E2B2B", width=20,height=3, anchor=W)
label2.grid(row=1,column=0)

frame5 = tk.Frame(root, bg="#FFE1CF")
frame5.place(relwidth= 0.27, relheight= 0.04, rely= 0.45, relx= 0.5)
b = tk.Button(frame5, text="Choose an xlsx file", command=AddFile, width=15)
b.grid(row=1,column=1)

frame6 = tk.Frame(root, bg="#FFE1CF")
frame6.place(relwidth= 0.4, relheight= 0.04, rely= 0.4, relx= 0.435)

frame7 = tk.Frame(root, bg="#FFE1CF")
frame7.place(relwidth= 0.9, relheight= 0.04, rely= 0.8, relx= 0.04)
b = tk.Button(frame7, text="Start", width=60, command=Start)
b.pack()

frame8 = tk.Frame(root, bg="#FFE1CF")
frame8.place(relwidth= 0.4, relheight= 0.04, rely= 0.52, relx= 0.435)
maxupload= tk.Entry(frame8, width=25)
maxupload.pack()


root.mainloop()  












