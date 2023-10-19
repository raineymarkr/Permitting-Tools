import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk
import pyodbc
import datetime
import subprocess
import sqlite3
import os
from docx.shared import Cm
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from PIL import ImageTk, Image
from pdf2image import convert_from_path
import win32com.client
import requests
import urllib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time

# Main Configuration

database = r'.\database.db'


#This is a list of all of the types of documents one may wish to generate.
document_types = {
    "Fee Letter": None,
    "Inspection Report": None,
    "Public Notice": {
        "IP": None,
        "LOP": None,
        "VAR": None,
        "NRU": None,
        "BSSE": None,
        "FAA": None,
        "OCS": None
    },
    "Permit": {
        "IP": None,
        "LOP": None,
        "VAR": None,
        "NRU": None,
        "401": None,
        "Time Extension": None,
        "No Permit Required": None
    }
}
icon = r'.\free.ico'



var_codes = {
    0:["Choose",""],
    1:["Dredging/Filling","335-8-2-.02"],
    2:["Mitigation","335-8-2-.03"],
    3:["Marinas","335-8-2-.04"],
    4:["Piers, Docks, Boathouses, and Other Pile Supported Structures","335-8-2-.05"],
    5:["Shoreline Stabilization and Erosion Mitigation","335-8-2-.06"],
    6:["Canals, Ditches, Boatslips ","335-8-2-.07"],
    7:["Construction/Other on Dunes","335-8-2-.08"],
    8:["GWE","335-8-2-.09"],
    9:["Siting, Construction and Operation of Energy Facilities","335-8-2-.010"],
    10:["CRD","335-8-2-.11"],
    11:["Discharge to Coastal Waters","335-8-2-.12"]
}

text_padding = 5
main = ttk.Window(themename='yeti')
main.iconbitmap(icon)
main.title("ADEM Coastal Document Genie")
windowcolor = tk.StringVar()
windowcolor.set('yeti')
style = ttk.Style()
countynum = ""
output_path = ""

#UTILITY FUNCTIONS
def toggle_dark_mode():
    conn = sqlite3.connect(database)
    c = conn.cursor()
    c.execute("SELECT Dark FROM settings")
    data = c.fetchall()  
    if data[0][0] == 0:
        style.theme_use('darkly')
        sql = f"UPDATE settings SET Dark = 1"
        c.execute(sql)
    if data[0][0] == 1:
        style.theme_use('yeti')
        sql = f"UPDATE settings SET Dark = 0"
        c.execute(sql)
    conn.commit()
    conn.close()


def open_file(filename):
    subprocess.Popen(["start",'', filename], shell=True)

def delete_previous_word(event):
    widget = event.widget
    current_text = widget.get()
    current_index = widget.index(tk.INSERT)
    
    
    # Find the start index of the previous word
    start_index = current_index - 1
    while start_index >= 0 and not current_text[start_index].isspace():
        start_index -= 1
    
    # Delete the previous word
    widget.delete(start_index + 1, current_index)

def delete_previous_word2(event):
    widget = event.widget
    current_index = widget.index(tk.INSERT)
    line, col = map(int, current_index.split('.'))
    
    # Get the current line text
    current_line = widget.get(f"{line}.0", f"{line}.end")
    
    # Find the start index of the previous word
    start_index = col - 1
    while start_index >= 0 and current_line[start_index].isspace():
        start_index -= 1
    while start_index >= 0 and not current_line[start_index].isspace():
        start_index -= 1
    
    # Delete the previous word
    if start_index >= 0:
        widget.delete(f"{line}.{start_index + 1}", current_index)
    else:
        # Delete from the beginning of the line if start_index < 0
        widget.delete(f"{line}.0", current_index)

def render_document(template, context, acamp, sam="", county="",perm_type="", doc_type=""):
    template.render(context)
    if county.lower() == 'mobile':
        countynum = ' 097'
    elif county.lower() == 'baldwin':
        countynum = ' 002'
    else:
        countynum = ' xxx'
    date = datetime.date.today()
    date.strftime("%m-%d-%y")

    acamp_folder = output_path+fr'\{acamp}'
    if os.path.exists(acamp_folder):
        filename = output_path+fr'\{acamp}\xxx '+ acamp +' '+ countynum +' ' +str(date)+ ' ' + perm_type +' '+ sam +' '+ doc_type +'.docx'
    else:
        filename =r'.\output\xxx ' + acamp +' '+ countynum +' ' +str(date)+ ' ' + perm_type +' '+ sam +' '+ doc_type +'.docx'
    template.save(filename.format(acamp))
    open_file(filename)

def send_email(subject_data, to_data, body_data):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.getNamespace("MAPI")

    drafts_folder = namespace.GetDefaultFolder(16)

    new_mail = outlook.CreateItem(0)

    new_mail.Subject = subject_data
    new_mail.Body = body_data
    new_mail.To = to_data

    try:
        new_mail.Save()
    except Exception as e:
        pass

def find_zip(address, city):
    zip_API = r"XWCGJUCYNPDBDK4EI7FD"
    params = {
        'address': address,
        'city': city,
        'state': 'AL',
        'key' : zip_API
    }
    encoded_params = urllib.parse.urlencode(params)
    url = f'https://api.zip-codes.com/ZipCodesAPI.svc/1.0/ZipCodeOfAddress?{encoded_params}'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        zip = data["Result"]["Address"]["Zip5"]
    else:
        zip = ''
    return zip

def findPID(address, city, county):
    driver = webdriver.Firefox()
    if(county.lower() == "mobile"):
        # Navigate to the website
        driver.get(r"https://cityofmobile.maps.arcgis.com/apps/webappviewer/index.html?id=44b3d1ecf57d4daa919a1e40ecca0c02")
        time.sleep(2)
        # Find form elements and fill them in
        search_box = driver.find_element(By.XPATH, '//*[@id="esri_dijit_Search_0_input"]')  # Search Box

        search_box.send_keys(address)

        # Submit the form
        search_button = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[7]/div[1]/div/div/div[2]")  # Replace 'login_button_id' with the actual ID of the login button
        search_button.click()
    else:
        # Navigate to the website
        driver.get(r"https://isv.kcsgis.com/al.baldwin_revenue/")

        accept_button = driver.find_element(By.XPATH, '/html/body/div[6]/div[2]/div[2]/button')
        accept_button.click()
        
        time.sleep(2)

        map_button = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '/html/body/div[2]/aside/button')))
        map_button.click()

        # Find form elements and fill them in
        search_box = driver.find_element(By.XPATH, '//*[@id="esri_dijit_Search_1_input"]')  # Search Box
        key = address + ", " + city +", Alabama"
        search_box.send_keys(key)

        # Submit the form
        search_button = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div/div/div[2]/div/div/div[2]")  # Replace 'login_button_id' with the actual ID of the login button
        search_button.click()

def openFolder():
    print(output_path)
    os.startfile(os.path.normpath(output_path))

def open_file_dialog():
    folder_path = filedialog.askdirectory()
    global output_path
    if folder_path:
        output_path = folder_path
        # Do something with the selected folder_path
        conn = sqlite3.connect(database)
        c = conn.cursor()
        sql = "UPDATE settings SET Output = ?"
        c.execute(sql, (folder_path,))
        conn.commit()
        conn.close()

    
        


#BEGIN PNOT WINDOW
def open_pnotinput_window():
    global pnot1
    global honorific, first_name, last_name, title, project_address
    global agent_name, agent_address
    global city, state
    global project_name, project_city, project_county
    global fee_amount, projcoords, project_description
    global adem_employee, adem_email, sam, acamp
    global timein, timeout, complaint, parcel_id
    global phone, comments, photos, participants
    pnot1 = ttk.Toplevel()
    pnot1.title("ADEM Coastal Document Genie")
    pnot1.iconbitmap(icon)

    pnot1.bind('<Return>', lambda event: get_pnot_values(acamp.get(), sam.get(), project_name.get(), project_address.get(), project_city.get(), project_county.get(), project_description.get(1.0, ttk.END), var_code.get(), parcel_id.get(), federal_agency.get()))

    left_frame = ttk.Frame(pnot1, )
    left_frame.pack(side=ttk.LEFT, padx=10)

    right_frame = ttk.Frame(pnot1, )
    right_frame.pack(side=ttk.LEFT, padx=10)

    greeting = ttk.Label(left_frame, text="Please provide the following information:", )

    greeting.pack(padx=text_padding, pady=text_padding)

    database_button = ttk.Button(left_frame, text = 'Load from Database', command = show_data)
    database_button.pack()
    
    honorific = ttk.Entry(left_frame)
    acamp_label = ttk.Label(left_frame, text="ACAMP Number:")
    acamp_label.pack(padx=text_padding, pady=text_padding)
    acamp = ttk.Entry(left_frame)
    acamp.bind("<Control-BackSpace>", delete_previous_word)
    acamp.pack(padx=text_padding, pady=text_padding)

    sam = ttk.Entry(left_frame)
    sam.bind("<Control-BackSpace>", delete_previous_word)

    if pnottype != "BSSE" and pnottype != "FAA" and pnottype != "OCS" :
        sam_label = ttk.Label(left_frame, text="SAM Number:")
        sam_label.pack(padx=text_padding, pady=text_padding)
        
        sam.pack(padx=text_padding, pady=text_padding)

    project_name_label = ttk.Label(left_frame, text="Project Name:")
    project_name_label.pack(padx=text_padding, pady=text_padding)
    project_name = ttk.Entry(left_frame)
    project_name.bind("<Control-BackSpace>", delete_previous_word)
    project_name.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Project Address/Location:")
    address_label.pack(padx=text_padding, pady=text_padding)
    project_address = ttk.Entry(left_frame)
    project_address.pack(padx=text_padding, pady=text_padding)
    project_address.bind("<Control-BackSpace>", delete_previous_word)

    project_city_label = ttk.Label(left_frame, text="Project City:")
    project_city_label.pack(padx=text_padding, pady=text_padding)
    project_city = ttk.Entry(left_frame)
    project_city.pack(padx=text_padding, pady=text_padding)
    project_city.bind("<Control-BackSpace>", delete_previous_word)

    project_county_label = ttk.Label(left_frame, text="Project County:")
    project_county_label.pack(padx=text_padding, pady=text_padding)
    

    project_county = ttk.Combobox(left_frame, values=['Mobile','Baldwin','Washington'])
    project_county.pack(padx=text_padding, pady=text_padding)
    variancecodes_label = ttk.Label(left_frame, text='Variance from Code:')
    var_code = ttk.Entry(left_frame)
    var_code.bind("<Control-BackSpace>", delete_previous_word)

    var_list = []
    for i in var_codes:
        var_list.append(var_codes.get(i)[0])


    clicked2 = ttk.StringVar()

    clicked2.set( "Choose Code:" )

    drop2 = ttk.OptionMenu( left_frame, clicked2, *var_list)


    def callback3(*args):
        for var in var_codes:
            for i in var_list:
                if clicked2.get() == var_codes.get(var)[0]:
                    var_code.delete(0,ttk.END)
                    var_code.insert(0, var_codes.get(var)[1])
        

    clicked2.trace("w", callback3)

    parcelid_label = ttk.Label(left_frame, text="Parcel ID:")
    parcel_id = ttk.Entry(left_frame)
    parcel_id.bind("<Control-BackSpace>", delete_previous_word)

    if pnottype == "VAR":        
        parcelid_label.pack(padx=text_padding, pady=text_padding)
        parcel_id.pack(padx=text_padding, pady=text_padding)
        find_pid_button = ttk.Button(left_frame,text='Find PID',command= lambda:findPID(project_address.get(),project_city.get(),project_county.get()))
        find_pid_button.pack()
        
        variancecodes_label.pack(padx=text_padding, pady=text_padding)
        drop2.pack(padx=text_padding, pady=text_padding)
        var_code.pack(padx=text_padding, pady=text_padding)

    federal_agency = ttk.Entry(left_frame)
    federal_agency.bind("<Control-BackSpace>", delete_previous_word)
    
    if pnottype == "FAA":
        fedagency_label = ttk.Label(left_frame, text="Federal Agency:")
        fedagency_label.pack(padx=text_padding, pady=text_padding)
        
        federal_agency.pack(padx=text_padding, pady=text_padding)
    
    project_desc_label = ttk.Label(right_frame, text="Project Description:")
    project_desc_label.pack(padx=text_padding, pady=text_padding)
    project_description = ttk.Text(right_frame)
    project_description.bind("<Control-BackSpace>", delete_previous_word2)
    project_description.pack(padx=text_padding, pady=text_padding)

    submit_button = ttk.Button(right_frame, text="Submit", command=lambda: get_pnot_values(acamp.get(), sam.get(), project_name.get(), project_address.get(), project_city.get(), project_county.get(), project_description.get(1.0, ttk.END), var_code.get(), parcel_id.get(), federal_agency.get()))
    submit_button.pack(padx=text_padding, pady=text_padding)

def get_pnot_values(acamp, sam="", project_name="", project_address="", project_city="", project_county="", project_description="", var_code="", parcel_id="", federal_agency=""):
    if pnottype == "IP":
        pnot_LOP(acamp, sam, project_name, project_address, project_city, project_county, project_description)
    elif pnottype == "LOP":
        pnot_LOP(acamp, sam, project_name, project_address, project_city, project_county, project_description)
    elif pnottype == "VAR":
        pnot_VAR(acamp, sam, project_name, project_address, project_city, project_county, project_description, var_code, parcel_id)
    elif pnottype == "NRU":
        pnot_NRU(acamp, sam, project_name, project_address, project_city, project_county, project_description)
    elif pnottype == "BSSE":
        pnot_BSSE(acamp, project_name, project_address, project_city, project_county, project_description)
    elif pnottype == "FAA":
        pnot_FAA(acamp, project_address, project_city, project_county, federal_agency, project_description)
    elif pnottype == "OCS":
        pnot_OCS(acamp, project_name, project_address, project_description)
    
    pnot.destroy()
    pnot1.destroy()

def pnot_BSSE(acamp, project_name, project_address, project_city, project_county, project_description):
    template = DocxTemplate(r'.\templates\BSEEPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'Project_Name': project_name,
        'Project_Location': project_address,
        'Project_Description': project_description,
        'Project_City': project_city,
        'Project_County': project_county
    }

    

    insert_data(acamp, context)
    render_document(template, context, acamp, sam="", county=project_county ,perm_type="", doc_type="BSSE_PNOT")
    paperlist = ''
    if project_county == 'Baldwin':
        paperlist = 'The Islander\nLagniappe'
    else:
        paperlist = 'Lagniappe'
    body = "For Publication.\n" + paperlist +"\nThank you, Kelly!"
    send_email('COASTAL PROGRAM • PNOT • '+acamp,'KBozeman@adem.alabama.gov',body)


def pnot_VAR(acamp, sam, project_name, project_address, project_city, project_county,project_description,var_code,parcel_id):
    template = DocxTemplate(r'.\templates\VARPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'SAM_Number': sam,
        'Project_Name': project_name,
        'Project_Location': project_address,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'Project_Description': project_description,
        'var_code': var_code,
    }

    insert_data(acamp, context)
    render_document(template, context, acamp, sam, county=project_county ,perm_type="", doc_type="VAR_PNOT")
    paperlist = ''
    if project_county == 'Baldwin':
        paperlist = 'The Islander\nLagniappe'
    else:
        paperlist = 'Lagniappe'
    body = "For Publication.\n" + paperlist +"\nThank you, Kelly!"
    send_email('COASTAL PROGRAM • PNOT • '+acamp,'KBozeman@adem.alabama.gov',body)

def pnot_NRU(acamp, sam, project_name, project_address, project_city, project_county,project_description):
    template = DocxTemplate(r'.\templates\NRUPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'SAM_Number': sam,
        'Project_Name': project_name,
        'Project_Location': project_address,
        'Project_City': project_city,
        'Project_County': project_county,
        'Project_Description': project_description,
    }


    insert_data(acamp, context)

    render_document(template, context, acamp, sam, county=project_county ,perm_type="", doc_type="NRU_PNOT")
    paperlist = ''
    if project_county == 'Baldwin':
        paperlist = 'The Islander\nLagniappe'
    else:
        paperlist = 'Lagniappe'
    body = "For Publication.\n" + paperlist +"\nThank you, Kelly!"
    send_email('COASTAL PROGRAM • PNOT • '+acamp,'KBozeman@adem.alabama.gov',body)

def pnot_FAA(acamp, project_address, project_city, project_county, federal_agency, project_description):
    template = DocxTemplate(r'.\templates\FAAPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'Federal_Agency': federal_agency,
        'Project_Location': project_address,
        'Project_City': project_city,
        'Project_County': project_county,
        'Project_Description': project_description,
    }
    insert_data(acamp, context)
    render_document(template, context, acamp, sam, county=project_county ,perm_type="", doc_type="FAA_PNOT")
    paperlist = ''
    if project_county == 'Baldwin':
        paperlist = 'The Islander\nLagniappe'
    else:
        paperlist = 'Lagniappe'
    body = "For Publication.\n" + paperlist +"\nThank you, Kelly!"
    send_email('COASTAL PROGRAM • PNOT • '+acamp,'KBozeman@adem.alabama.gov',body)

def pnot_LOP(acamp, sam, project_name, project_address, project_city, project_county,project_description):
    template = DocxTemplate(r'.\templates\LOPPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'SAM_Number': sam,
        'Project_Name': project_name,
        'Project_Location': project_address,
        'Project_City': project_city,
        'Project_County': project_county,
        'Project_Description': project_description,
    }
    insert_data(acamp, context)
    render_document(template, context, acamp, sam, county=project_county ,perm_type="", doc_type="LOP_PNOT")
    paperlist = ''
    if project_county == 'Baldwin':
        paperlist = 'The Islander\nLagniappe'
    else:
        paperlist = 'Lagniappe'
    body = "For Publication.\n" + paperlist +"\nThank you, Kelly!"
    send_email('COASTAL PROGRAM • PNOT • '+acamp,'KBozeman@adem.alabama.gov',body)

def pnot_OCS(acamp, project_name, project_address, project_description):
    template = DocxTemplate(r'.\templates\OCSPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'Project_Name': project_name,
        'Project_Location': project_address,
        'Project_Description': project_description,
    }
    insert_data(acamp, context)
    render_document(template, context, acamp, sam, county=project_county ,perm_type="", doc_type="OCS_PNOT")
    paperlist = ''
    if project_county == 'Baldwin':
        paperlist = 'The Islander\nLagniappe'
    else:
        paperlist = 'Lagniappe'
    body = "For Publication.\n" + paperlist +"\nThank you, Kelly!"
    send_email('COASTAL PROGRAM • PNOT • '+acamp,'KBozeman@adem.alabama.gov',body)

def set_pnottype(document_type):
    global pnottype
    pnottype = document_type
    global pnot1
    open_pnotinput_window()

def open_pnot_window():
    global pnot
    
    global pnottype

    # PNOT Choice Window
    pnot = ttk.Toplevel()
    pnot.title("ADEM Coastal Document Genie")
    pnot.iconbitmap(icon)
    chosen_type = list(document_types.keys())[2]
    subtypes = document_types[chosen_type]
    greeting = ttk.Label(pnot, text="What type of Public Notice do you want to generate?")
    greeting.pack(padx=text_padding, pady=text_padding)

    for i, document_type in enumerate(subtypes.keys()):
        document_button = ttk.Button(pnot, text=f"{i+1}. {document_type}")
        document_button.pack(padx=text_padding, pady=text_padding)
        document_button.configure(command=lambda doc_type=document_type: set_pnottype(doc_type))
    
    #END  PNOT WINDOW
#END PNOTS
def print_entry_content(var_name, index, mode):
    print(f"{var_name}: {vars()[var_name].get()}")
#BEGIN perm WINDOW
def open_perminput_window():
    global perm1
    global honorific, first_name, last_name, title, project_address
    global agent_name, agent_address, var_code
    global city, state, zip, project_description
    global project_name, project_city, project_county
    global fee_amount, projcoords, parcel_id
    global adem_employee, adem_email, sam, acamp
    global timein, timeout, complaint, fee_received
    global phone, comments, photos, participants
    global prefile_date, notice_type, pnot_date, jpn_date

    perm1 = ttk.Toplevel()
    perm1.iconbitmap(icon)
    perm1.title("ADEM Coastal Document Genie")
    perm1.bind('<Return>', lambda event: get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), project_address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, ttk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get(), var_code.get()))

    left_frame = ttk.Frame(perm1, )
    left_frame.pack(side=ttk.LEFT)

    middle_frame = ttk.Frame(perm1, )
    middle_frame.pack(side=ttk.LEFT)

    right_frame = ttk.Frame(perm1, )
    right_frame.pack(side=ttk.LEFT)

    greeting = ttk.Label(left_frame, text="Please provide the following information:")
    greeting.pack(padx=text_padding, pady=text_padding)

    database_button = ttk.Button(left_frame, text = 'Load from Database', command = show_data)
    database_button.pack()

    acamp_label = ttk.Label(middle_frame, text="ACAMP Number:")
    acamp_label.pack(padx=text_padding, pady=text_padding)
    acamp = ttk.Entry(middle_frame)
    acamp.bind("<Control-BackSpace>", delete_previous_word)
    acamp.pack(padx=text_padding, pady=text_padding)

    sam_label = ttk.Label(middle_frame, text="SAM Number:")
    sam_label.pack(padx=text_padding, pady=text_padding)
    sam = ttk.Entry(middle_frame)
    sam.bind("<Control-BackSpace>", delete_previous_word)
    sam.pack(padx=text_padding, pady=text_padding)

    honorific_label = ttk.Label(left_frame, text="Honorific:")
    honorific_label.pack(padx=text_padding, pady=text_padding)
    honorific = ttk.Entry(left_frame)
    honorific.bind("<Control-BackSpace>", delete_previous_word)
    honorific.pack(padx=text_padding, pady=text_padding)

    first_name_label = ttk.Label(left_frame, text="First Name:")
    first_name_label.pack(padx=text_padding, pady=text_padding)
    first_name = ttk.Entry(left_frame)
    first_name.bind("<Control-BackSpace>", delete_previous_word)
    first_name.pack(padx=text_padding, pady=text_padding)

    last_name_label = ttk.Label(left_frame, text="Last Name:")
    last_name_label.pack(padx=text_padding, pady=text_padding)
    last_name = ttk.Entry(left_frame)
    last_name.bind("<Control-BackSpace>", delete_previous_word)
    last_name.pack(padx=text_padding, pady=text_padding)

    title_label = ttk.Label(left_frame, text="Title:")
    title_label.pack(padx=text_padding, pady=text_padding)
    title = ttk.Entry(left_frame)
    title.bind("<Control-BackSpace>", delete_previous_word)
    title.pack(padx=text_padding, pady=text_padding)

    agents = get_agents()
    agent_list = []
    for i in agents:
        agent_list.append(i[0])
    # Create Label
    label2 = ttk.Label(left_frame , text = "Choose Agent: " )
    label2.pack(padx=text_padding, pady=text_padding)  

    clicked1 = ttk.StringVar()

    clicked1.set( "Choose Agent:" )

    drop1 = ttk.OptionMenu( left_frame, clicked1, *agent_list)
    drop1.pack(padx=text_padding, pady=text_padding)

    def callback2(*args):
        for agent in agents:
            for i in agent_list:
                if clicked1.get() == agent[0]:
                    agent_name.delete(0,ttk.END)
                    agent_address.delete(0,ttk.END)
                    agent_address.insert(0, agent[1])
                    agent_name.insert(0, agent[0])
                    city.delete(0, ttk.END)
                    city.insert(0, agent[2])
                    state.delete(0, ttk.END)
                    state.insert(0, agent[3])
        

    clicked1.trace("w", callback2)

    agent_name_label = ttk.Label(left_frame, text="Agent Full Name:")
    agent_name_label.pack(padx=text_padding, pady=text_padding)
    agent_name = ttk.Entry(left_frame)
    agent_name.bind("<Control-BackSpace>", delete_previous_word)
    agent_name.pack(padx=text_padding, pady=text_padding)

    agent_address_label = ttk.Label(left_frame, text="Agent Address:")
    agent_address_label.pack(padx=text_padding, pady=text_padding)
    agent_address = ttk.Entry(left_frame)
    agent_address.bind("<Control-BackSpace>", delete_previous_word)
    agent_address.pack(padx=text_padding, pady=text_padding)

    city_label = ttk.Label(left_frame, text="City:")
    city_label.pack(padx=text_padding, pady=text_padding)
    city = ttk.Entry(left_frame)
    city.bind("<Control-BackSpace>", delete_previous_word)
    city.pack(padx=text_padding, pady=text_padding)

    state_label = ttk.Label(left_frame, text="State:")
    state_label.pack(padx=text_padding, pady=text_padding)
    state = ttk.Entry(left_frame)
    state.bind("<Control-BackSpace>", delete_previous_word)
    state.pack(padx=text_padding, pady=text_padding)

    zip_code_label = ttk.Label(left_frame, text="Zip Code:")
    zip_code_label.pack(padx=text_padding, pady=text_padding)
    zip = ttk.Entry(left_frame)
    zip.bind("<Control-BackSpace>", delete_previous_word)
    zip.pack(padx=text_padding, pady=text_padding)

    get_zip = ttk.Button(left_frame, text ='Get Zip', command=lambda:zip.insert(0, find_zip(agent_address.get(), city.get())))
    get_zip.pack()

    project_name_label = ttk.Label(middle_frame, text="Project Name:")
    project_name_label.pack(padx=text_padding, pady=text_padding)
    project_name = ttk.Entry(middle_frame)
    project_name.bind("<Control-BackSpace>", delete_previous_word)
    project_name.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(middle_frame, text="Project Address:")
    address_label.pack(padx=text_padding, pady=text_padding)
    project_address = ttk.Entry(middle_frame)
    project_address.bind("<Control-BackSpace>", delete_previous_word)
    project_address.pack(padx=text_padding, pady=text_padding)

    project_city_label = ttk.Label(middle_frame, text="Project City:")
    project_city_label.pack(padx=text_padding, pady=text_padding)
    project_city = ttk.Entry(middle_frame)
    project_city.bind("<Control-BackSpace>", delete_previous_word)
    project_city.pack(padx=text_padding, pady=text_padding)

    project_county_label = ttk.Label(middle_frame, text="Project County:")
    project_county_label.pack(padx=text_padding, pady=text_padding)
    
    project_county = ttk.Combobox(middle_frame, values=['Mobile','Baldwin','Washington'])
    project_county.pack(padx=text_padding, pady=text_padding)

    parcel_id_label = ttk.Label(middle_frame, text="Parcel ID:")
    parcel_id_label.pack(padx=text_padding, pady=text_padding)
    parcel_id = ttk.Entry(middle_frame)
    parcel_id.bind("<Control-BackSpace>", delete_previous_word)
    parcel_id.pack(padx=text_padding, pady=text_padding)

    find_pid_button = ttk.Button(middle_frame,text='Find PID',command= lambda:findPID(project_address.get(),project_city.get(),project_county.get()))
    find_pid_button.pack()

    prefile_date_label = ttk.Label(middle_frame, text="Prefile Date:")
    prefile_date_label.pack(padx=text_padding, pady=text_padding)
    prefile_date = ttk.Entry(middle_frame)
    prefile_date.bind("<Control-BackSpace>", delete_previous_word)
    prefile_date.pack(padx=text_padding, pady=text_padding)

    notice_type_label = ttk.Label(middle_frame, text="Notice Type:")
    notice_type_label.pack(padx=text_padding, pady=text_padding)
    notice_type = ttk.Entry(middle_frame)
    notice_type.bind("<Control-BackSpace>", delete_previous_word)
    notice_type.pack(padx=text_padding, pady=text_padding)

    jpn_date_label = ttk.Label(middle_frame, text="USACE JPN Date:")
    jpn_date_label.pack(padx=text_padding, pady=text_padding)
    jpn_date = ttk.Entry(middle_frame)
    jpn_date.bind("<Control-BackSpace>", delete_previous_word)
    jpn_date.pack(padx=text_padding, pady=text_padding)

    pnot_date_label = ttk.Label(middle_frame, text="ADEM PNOT Date:")
    pnot_date_label.pack(padx=text_padding, pady=text_padding)
    pnot_date = ttk.Entry(middle_frame)
    pnot_date.bind("<Control-BackSpace>", delete_previous_word)
    pnot_date.pack(padx=text_padding, pady=text_padding)

    exp_date_label = ttk.Label(middle_frame, text="Expiration Date:")
    exp_date = ttk.Entry(middle_frame)
    exp_date.bind("<Control-BackSpace>", delete_previous_word)
    exp_date1_label = ttk.Label(middle_frame, text="New Expiration Date:")
    exp_date1 = ttk.Entry(middle_frame)
    exp_date1.bind("<Control-BackSpace>", delete_previous_word)
    
    if permtype == "Time Extension":
        
        exp_date_label.pack(padx=text_padding, pady=text_padding)
        exp_date.pack(padx=text_padding, pady=text_padding)
        exp_date1_label.pack(padx=text_padding, pady=text_padding)
        exp_date1.pack(padx=text_padding, pady=text_padding)

    var_code = ttk.Entry(right_frame)
    var_code.bind("<Control-BackSpace>", delete_previous_word)

    if permtype == "VAR":

            var_list = []
            for i in var_codes:
                var_list.append(var_codes.get(i)[0])
            # Create Label
            label3 = ttk.Label(right_frame , text = "Variance From Code: " )
            label3.pack(padx=text_padding, pady=text_padding)  

            clicked2 = ttk.StringVar()

            clicked2.set( "Choose Code:" )

            drop2 = ttk.OptionMenu( right_frame, clicked2, *var_list)
            drop2.pack(padx=text_padding, pady=text_padding)

            def callback3(*args):
                for var in var_codes:
                    for i in var_list:
                        if clicked2.get() == var_codes.get(var)[0]:
                            var_code.delete(0,ttk.END)
                            var_code.insert(0, var_codes.get(var)[1])
                

            clicked2.trace("w", callback3)
            var_code.pack(padx=text_padding, pady=text_padding)

    npdes_num_label = ttk.Label(middle_frame, text="NPDES Permit:")
    npdes_num = ttk.Entry(middle_frame)
    npdes_num.bind("<Control-BackSpace>", delete_previous_word)
    npdes_date_label = ttk.Label(middle_frame, text="NPDES Permit Date:")
    npdes_date = ttk.Entry(middle_frame)
    npdes_date.bind("<Control-BackSpace>", delete_previous_word)
    parcel_size_label = ttk.Label(middle_frame,text="Parcel Size (Ac):")
    parcel_size = ttk.Entry(middle_frame)
    parcel_size.bind("<Control-BackSpace>", delete_previous_word)
    
    if permtype == "NRU":
        npdes_num_label.pack(padx=text_padding, pady=text_padding)
        npdes_num.pack(padx=text_padding, pady=text_padding)
        npdes_date_label.pack(padx=text_padding, pady=text_padding)
        npdes_date.pack(padx=text_padding, pady=text_padding)
        parcel_size_label.pack(padx=text_padding, pady=text_padding)
        parcel_size.pack(padx=text_padding, pady=text_padding)

    if permtype == "IP":
        prefile_date_label.pack(padx=text_padding, pady=text_padding)
        prefile_date.pack(padx=text_padding, pady=text_padding)
    
    fee_amount_label = ttk.Label(right_frame, text="Fee Amount:")
    fee_amount_label.pack(padx=text_padding, pady=text_padding)
    fee_amount = ttk.Entry(right_frame)
    fee_amount.bind("<Control-BackSpace>", delete_previous_word)
    fee_amount.pack(padx=text_padding, pady=text_padding)

    fee_received_label = ttk.Label(right_frame, text="Fee Received:")
    fee_received_label.pack(padx=text_padding, pady=text_padding)
    fee_received = ttk.Entry(right_frame)
    fee_received.bind("<Control-BackSpace>", delete_previous_word)
    fee_received.pack(padx=text_padding, pady=text_padding)

    project_description_label = ttk.Label(right_frame, text="Project Description:")
    project_description_label.pack(padx=text_padding, pady=text_padding)
    project_description = ttk.Text(right_frame)
    project_description.bind("<Control-BackSpace>", delete_previous_word2)
    project_description.pack(padx=text_padding, pady=text_padding)

    get_data3()

    adem_employee_label = ttk.Label(right_frame, text="ADEM Employee:")
    adem_employee_label.pack(padx=text_padding, pady=text_padding)
    adem_employee = ttk.Entry(right_frame,textvariable=name_var)
    adem_employee.bind("<Control-BackSpace>", delete_previous_word)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    adem_email_label = ttk.Label(right_frame, text="ADEM Email:")
    adem_email_label.pack(padx=text_padding, pady=text_padding)
    adem_email = ttk.Entry(right_frame, textvariable= email_var)
    adem_email.bind("<Control-BackSpace>", delete_previous_word)
    adem_email.pack(padx=text_padding, pady=text_padding)
    
    adem_pronoun = ""
    def get_pronoun():
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT Pronoun FROM settings WHERE Name = ?", (adem_employee.get(),))
        data = c.fetchall()
        adem_pronoun = data[0][0]
        print(adem_pronoun)
        get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), project_address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, ttk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get(),var_code.get())

    submit_button = ttk.Button(right_frame, text="Submit", command=lambda: get_pronoun())
    submit_button.pack(padx=text_padding, pady=text_padding)

    

def get_perm_values(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1, npdes_date, npdes_num, parcel_size, var_code):
    if permtype == "IP":
        perm_LOP(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email)
    elif permtype == "LOP":
        perm_LOP(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email)
    elif permtype == "VAR":
        perm_VAR(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email,var_code)
    elif permtype == "NRU":
        perm_NRU(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, npdes_date,npdes_num)
    elif permtype == "401":
        perm_401(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email)
    elif permtype == "Time Extension":
        perm_TIMEEXT(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1)    
    elif permtype == "No Permit Required":
        perm_NOREQ(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, adem_employee, adem_email)
    
    perm.destroy()
    perm1.destroy()


def perm_401(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email):
# Import template document
    template = DocxTemplate(r'.\templates\401WQC_Temp.docx')
    template2 = DocxTemplate(r'.\templates\401Rat_Temp.docx')

    # Declare template variables
    context = {
        
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': project_address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip,
        'Project_Name': project_name,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'SAM_Number': sam,
        'ACAMP_Number': acamp,
        'Prefile_Date': prefile_date,
        'Notice_Type': notice_type,
        'JPN_Date': jpn_date,
        'PNOT_Date': pnot_date,
        'Project_Description': project_description,
        'Fee_Amount': fee_amount,
        'Fee_Received': fee_received,
        'ADEM_Employee': adem_employee,
        'ADEM_Email': adem_email,
        'ADEM_Pronoun' : pronoun_var.get()
    }
    
    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 003'
    else:
        countynum = ' xxx'
    agent = {
        'name': agent_name,
        'address': agent_address,
        'city': city,
        'state': state,
        'email': ''
    }

    insert_agent_data(agent_name, agent)
    insert_data(acamp, context)
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""

    masteridBody = f""" Hi Spring, 

I need a master ID for the application below. Please let me know if you have any questions or concerns! 

Permitee Name: {project_name}
Permit Number: ACAMP-{acamp}
Initial Application
Date application received: {prefile_date}
Facility Name: None – Single Family Home
Parcel Number(s): {parcel_id}
Facility Address: {project_address}
Latitude/Longitude: 
Offshore: No
Fee Amount Paid: ${fee_amount}
Master ID: 

Thank you!

"""

    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
    send_email('Master ID: ACAMP-' + acamp + ' // '+ project_name,'STate@adem.alabama.gov', masteridBody)
    # Render automated report
    render_document(template, context, acamp, sam, county=project_county ,perm_type="401WQ", doc_type="401WQ")
    render_document(template2, context, acamp, sam, county=project_county ,perm_type="401WQ", doc_type="RATIONALE")
    
    

def perm_LOP(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email):
    # Import template document
    templatePerm1 = DocxTemplate('templates/LOPW_Temp.docx')
    templatePerm2 = DocxTemplate('templates/LOPC_Temp.docx')
    templateRat = DocxTemplate('templates/LOPRat_Temp.docx')
    
    # Declare template variables
    context = {
        
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': project_address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip,
        'Project_Name': project_name,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'SAM_Number': sam,
        'ACAMP_Number': acamp,
        'Prefile_Date': prefile_date,
        'Notice_Type': notice_type,
        'JPN_Date': jpn_date,
        'PNOT_Date': pnot_date,
        'Project_Description': project_description,
        'Fee_Amount': fee_amount,
        'Fee_Received': fee_received,
        'ADEM_Employee': adem_employee,
        'ADEM_Email': adem_email,
        'ADEM_Pronoun' : pronoun_var.get()
    }

    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 003'
    else:
        countynum = ' xxx'
    
    agent = {
        'name': agent_name,
        'address': agent_address,
        'city': city,
        'state': state,
        'email': ''
    }

    insert_agent_data(agent_name, agent)
    insert_data(acamp, context)
    # Render automated report
    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)

    masteridBody = f""" Hi Spring, 

I need a master ID for the application below. Please let me know if you have any questions or concerns! 

Permitee Name: {project_name}
Permit Number: ACAMP-{acamp}
Initial Application
Date application received: {prefile_date}
Facility Name: None – Single Family Home
Parcel Number(s): {parcel_id}
Facility Address: {project_address}
Latitude/Longitude: 
Offshore: No
Fee Amount Paid: ${fee_amount}
Master ID: 

Thank you!

"""
    send_email('Master ID: ACAMP-' + acamp + ' // '+ project_name,'STate@adem.alabama.gov', masteridBody)
    render_document(templatePerm2, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="CZM")
    render_document(templatePerm1, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="401WQ")
    render_document(templateRat,context,acamp,sam,project_county,"CZCERT","RATIONALE")
    

def perm_VAR(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email,var_code):
    # Import template document
    templatePerm1 = DocxTemplate('templates/LOPW_Temp.docx')
    templatePerm2 = DocxTemplate('templates/VARC_Temp.docx')
    templateRat = DocxTemplate('templates/LOPRat_Temp.docx')
    
    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 002'
    else:
        countynum = ' xxx'

    # Declare template variables
    context = {
        
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': project_address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip,
        'Project_Name': project_name,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'SAM_Number': sam,
        'ACAMP_Number': acamp,
        'Prefile_Date': prefile_date,
        'Notice_Type': notice_type,
        'JPN_Date': jpn_date,
        'PNOT_Date': pnot_date,
        'Project_Description': project_description,
        'Fee_Amount': fee_amount,
        'Fee_Received': fee_received,
        'ADEM_Employee': adem_employee,
        'ADEM_Email': adem_email,
        'var_code': var_code,
        'ADEM_Pronoun' : pronoun_var.get()
    }

    insert_data(acamp, context)
    agent = {
        'name': agent_name,
        'address': agent_address,
        'city': city,
        'state': state,
        'email': ''
    }

    insert_agent_data(agent_name, agent)
    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review (VARIANCE): ' + acamp,'CMcNeill@adem.alabama.gov',body)

    masteridBody = f""" Hi Spring, 

I need a master ID for the application below. Please let me know if you have any questions or concerns! 

Permitee Name: {project_name}
Permit Number: ACAMP-{acamp}
Initial Application
Date application received: {prefile_date}
Facility Name: None – Single Family Home
Parcel Number(s): {parcel_id}
Facility Address: {project_address}
Latitude/Longitude: 
Offshore: No
Fee Amount Paid: ${fee_amount}
Master ID: 

Thank you!

"""
    send_email('Master ID: ACAMP-' + acamp + ' // '+ project_name,'STate@adem.alabama.gov', masteridBody)
    render_document(templatePerm2, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="CZM")
    render_document(templatePerm1, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="401WQ")
    render_document(templateRat,context,acamp,sam,project_county,"CZCERT","RATIONALE")



def perm_NRU(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email,npdes_date,npdes_num):
    # Import template document
    templaten = DocxTemplate('templates/NRU_Temp.docx')
    templatec = DocxTemplate('templates/LOPC_Temp.docx')
    template2 = DocxTemplate('templates/NRURat_Temp.docx')
    


    # Declare template variables
    context = {
        
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': project_address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip,
        'Project_Name': project_name,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'SAM_Number': sam,
        'ACAMP_Number': acamp,
        'Prefile_Date': prefile_date,
        'Notice_Type': notice_type,
        'JPN_Date': jpn_date,
        'PNOT_Date': pnot_date,
        'Project_Description': project_description,
        'Fee_Amount': fee_amount,
        'Fee_Received': fee_received,
        'ADEM_Employee': adem_employee,
        'ADEM_Email': adem_email,
        'NPDES_Date': npdes_date,
        'NPDES_Number': npdes_num,
        'ADEM_Pronoun' : pronoun_var.get()
    }
    insert_data(acamp, context)

    agent = {
        'name': agent_name,
        'address': agent_address,
        'city': city,
        'state': state,
        'email': ''
    }

    insert_agent_data(agent_name, agent)
    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
    masteridBody = f""" Hi Spring, 

I need a master ID for the application below. Please let me know if you have any questions or concerns! 

Permitee Name: {project_name}
Permit Number: ACAMP-{acamp}
Initial Application
Date application received: {prefile_date}
Facility Name: None – Single Family Home
Parcel Number(s): {parcel_id}
Facility Address: {project_address}
Latitude/Longitude: 
Offshore: No
Fee Amount Paid: ${fee_amount}
Master ID: 

Thank you!

"""
    send_email('Master ID: ACAMP-' + acamp + ' // '+ project_name,'STate@adem.alabama.gov', masteridBody)
    render_document(templaten, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="NRU")
    render_document(templatec, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="CZM")
    render_document(template2,context,acamp,sam,project_county,"CZCERT","RATIONALE")


def perm_TIMEEXT(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1):
    # Import template document
    template = DocxTemplate(r'.\templates\401EXT_Temp.docx')
    
    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 002'
    else:
        countynum = ' xxx'
    
    # Declare template variables
    context = {
        
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': project_address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip,
        'Project_Name': project_name,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'SAM_Number': sam,
        'ACAMP_Number': acamp,
        'Prefile_Date': prefile_date,
        'Notice_Type': notice_type,
        'JPN_Date': jpn_date,
        'PNOT_Date': pnot_date,
        'Project_Description': project_description,
        'Fee_Amount': fee_amount,
        'Fee_Received': fee_received,
        'ADEM_Employee': adem_employee,
        'ADEM_Email': adem_email,
        'Expiration_Date': exp_date,
        'New_Expiration': exp_date1,
        'ADEM_Pronoun' : pronoun_var.get()
    }
    insert_data(acamp, context)

    agent = {
        'name': agent_name,
        'address': agent_address,
        'city': city,
        'state': state,
        'email': ''
    }

    insert_agent_data(agent_name, agent)

    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
    # Render automated report
    render_document(template,context,acamp,sam,project_county,"","Time Extension")

def perm_NOREQ(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, adem_employee, adem_email):
    # Import template document
    template = DocxTemplate(r'.\templates\NPR_Temp.docx')
    
    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 002'
    else:
        countynum = ' xxx'
    
    # Declare template variables
    context = {
        
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': project_address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip,
        'Project_Name': project_name,
        'Project_Location': project_address,
        'Project_City': project_city,
        'Project_County': project_county,
        'SAM_Number': sam,
        'ACAMP_Number': acamp,
        'ADEM_Employee': adem_employee,
        'ADEM_Email': adem_email,
        'ADEM_Pronoun' : pronoun_var.get()
    }
    insert_data(acamp, context)

    agent = {
        'name': agent_name,
        'address': agent_address,
        'city': city,
        'state': state,
        'email': ''
    }

    insert_agent_data(agent_name, agent)

    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
    render_document(template,context,acamp,sam,project_county,"","No Permit Required")

def set_permtype(document_type):
    global permtype
    permtype = document_type
    global perm1
    open_perminput_window()

def open_perm_window():
    global perm
    
    global permtype
    

    # perm Choice Window
    perm = ttk.Toplevel()
    perm.title("ADEM Coastal Document Genie")
    perm.iconbitmap(icon)
    chosen_type = list(document_types.keys())[3]
    subtypes = document_types[chosen_type]
    greeting = ttk.Label(perm, text="What type of Permit do you want to generate?")
    greeting.pack(padx=text_padding, pady=text_padding)

    for i, document_type in enumerate(subtypes.keys()):
        document_button = ttk.Button(perm, text=f"{i+1}. {document_type}")
        document_button.pack(padx=text_padding, pady=text_padding)
        document_button.configure(command=lambda doc_type=document_type: set_permtype(doc_type))
    
    #END  perm WINDOW
#END PERMITS

# BEGIN INSPECTION REPORT
def get_inspr_values():
    # Import template document
    template = DocxTemplate(r'.\templates\Insp_Temp.docx')


    context = {
        
        'time_in': timein.get(),
        'time_out': timeout.get(),
        'Applicant_FirstName': first_name.get(),
        'Applicant_LastName': last_name.get(),
        'Applicant_Phone': phone.get(),
        'Applicant_Address': project_address.get(),
        'Proj_Cords': projcoords.get(),
        'Proj_Complaint': complaint.get(),
        'Project_Name': project_name.get(),
        'Project_City': project_city.get(),
        'Project_County': project_county.get(),
        'SAM_Number': sam.get(),
        'ACAMP_Number': acamp.get(),
        'Project_Description': comments.get(1.0, ttk.END),
        'Photos': photos.get(),
        'Other_Names': participants.get(),
        'ADEM_Employee': adem_employee.get(),
        'ADEM_Email': adem_email.get()
    }

    data = {'Applicant_FirstName': first_name.get(),
        'Applicant_LastName': last_name.get(),
        'Applicant_Address': project_address.get(),
        'Project_Name': project_name.get(),
        'Project_City': project_city.get(),
        'Project_County': project_county.get(),
        'SAM_Number': sam.get(),
        'ACAMP_Number': acamp.get(),
        'Project_Description': comments.get(1.0, ttk.END),
        'ADEM_Employee': adem_employee.get(),
        'ADEM_Email': adem_email.get()}
    insert_data(acamp.get(), context)
    # Render automated report
    render_document(template,context,context.get('ACAMP_Number'),context.get('SAM_Number'),"","","Inspection Report")
    perm.destroy()
# Inspection Report Window
def open_inspr_window():
    # Inspection Report Window
    global inspr
    inspr = ttk.Toplevel()
    inspr.title("ADEM Coastal Document Genie")
    inspr.iconbitmap(icon)
    inspr.bind('<Return>', lambda event: get_inspr_values())
    greeting = ttk.Label(inspr, text="Please provide the following information:")
    greeting.pack(padx=text_padding, pady=text_padding)

    database_button = ttk.Button(inspr, text = 'Load from Database', command = show_data)
    database_button.pack()
    
    global honorific, first_name, last_name, title, project_address
    global agent_name, agent_address
    global city, state
    global project_name, project_city, project_county
    global fee_amount, projcoords
    global adem_employee, adem_email, sam, acamp
    global timein, timeout, complaint
    global phone, comments, photos, participants
    
    # Frame for left column
    left_frame = ttk.Frame(inspr, )
    left_frame.pack(side=ttk.LEFT, padx=10)

    # Frame for right column
    right_frame = ttk.Frame(inspr, )
    right_frame.pack(side=ttk.LEFT, padx=10)

    # Entry fields with labels in left column
    acamp_label = ttk.Label(left_frame, text="ACAMP Number:")
    acamp_label.pack(padx=text_padding, pady=text_padding)
    acamp = ttk.Entry(left_frame)
    acamp.bind("<Control-BackSpace>", delete_previous_word)
    acamp.pack(padx=text_padding, pady=text_padding)

    honorific = ttk.Entry(left_frame)

    sam_label = ttk.Label(left_frame, text="SAM Number:")
    sam_label.pack(padx=text_padding, pady=text_padding)
    sam = ttk.Entry(left_frame)
    sam.bind("<Control-BackSpace>", delete_previous_word)
    sam.pack(padx=text_padding, pady=text_padding)

    timein_label = ttk.Label(left_frame, text="Inspection Time Start:")
    timein_label.pack(padx=text_padding, pady=text_padding)
    timein = ttk.Entry(left_frame)
    timein.bind("<Control-BackSpace>", delete_previous_word)
    timein.pack(padx=text_padding, pady=text_padding)

    timeout_label = ttk.Label(left_frame, text="Inspection Time End:")
    timeout_label.pack(padx=text_padding, pady=text_padding)
    timeout = ttk.Entry(left_frame)
    timeout.bind("<Control-BackSpace>", delete_previous_word)
    timeout.pack(padx=text_padding, pady=text_padding)

    firstname_label = ttk.Label(left_frame, text="Applicant First Name:")
    firstname_label.pack(padx=text_padding, pady=text_padding)
    first_name = ttk.Entry(left_frame)
    first_name.bind("<Control-BackSpace>", delete_previous_word)
    first_name.pack(padx=text_padding, pady=text_padding)

    lastname_label = ttk.Label(left_frame, text="Applicant Last Name:")
    lastname_label.pack(padx=text_padding, pady=text_padding)
    last_name = ttk.Entry(left_frame)
    last_name.bind("<Control-BackSpace>", delete_previous_word)
    last_name.pack(padx=text_padding, pady=text_padding)

    complaint_label = ttk.Label(left_frame, text="Complaint #:")
    complaint_label.pack(padx=text_padding, pady=text_padding)
    complaint = ttk.Entry(left_frame)
    complaint.pack(padx=text_padding, pady=text_padding)

    # Entry fields with labels in right column
    phone_label = ttk.Label(left_frame, text="Applicant Phone Number:")
    phone_label.pack(padx=text_padding, pady=text_padding)
    phone = ttk.Entry(left_frame)
    phone.bind("<Control-BackSpace>", delete_previous_word)
    phone.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Applicant Address:")
    address_label.pack(padx=text_padding, pady=text_padding)
    project_address = ttk.Entry(left_frame)
    project_address.bind("<Control-BackSpace>", delete_previous_word)
    project_address.pack(padx=text_padding, pady=text_padding)

    projcoords_label = ttk.Label(right_frame, text="Project Coordinates:")
    projcoords_label.pack(padx=text_padding, pady=text_padding)
    projcoords = ttk.Entry(right_frame)
    projcoords.bind("<Control-BackSpace>", delete_previous_word)
    projcoords.pack(padx=text_padding, pady=text_padding)

    project_name_label = ttk.Label(right_frame, text="Project Name:")
    project_name_label.pack(padx=text_padding, pady=text_padding)
    project_name = ttk.Entry(right_frame)
    project_name.bind("<Control-BackSpace>", delete_previous_word)
    project_name.pack(padx=text_padding, pady=text_padding)

    project_city_label = ttk.Label(right_frame, text="Project City:")
    project_city_label.pack(padx=text_padding, pady=text_padding)
    project_city = ttk.Entry(right_frame)
    project_city.bind("<Control-BackSpace>", delete_previous_word)
    project_city.pack(padx=text_padding, pady=text_padding)

    project_county_label = ttk.Label(right_frame, text="Project County:")
    project_county_label.pack(padx=text_padding, pady=text_padding)
    project_county = ttk.Combobox(right_frame, values=['Mobile','Baldwin','Washington'])
    project_county.pack(padx=text_padding, pady=text_padding)

    photos_label = ttk.Label(right_frame, text="Photos Taken? (Yes/No):")
    photos_label.pack(padx=text_padding, pady=text_padding)
    photos = ttk.Entry(right_frame)
    photos.bind("<Control-BackSpace>", delete_previous_word)
    photos.pack(padx=text_padding, pady=text_padding)

    participants_label = ttk.Label(right_frame, text="Other Participants (Name, Org):")
    participants_label.pack(padx=text_padding, pady=text_padding)
    participants = ttk.Entry(right_frame)
    participants.bind("<Control-BackSpace>", delete_previous_word)
    participants.pack(padx=text_padding, pady=text_padding)

    label1 = ttk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack(padx=text_padding, pady=text_padding)  
    get_data3()
    yourname_label = ttk.Label(right_frame, text="Your Name:")
    yourname_label.pack(pady=text_padding)
    adem_employee = ttk.Entry(right_frame, textvariable= name_var)
    adem_employee.bind("<Control-BackSpace>", delete_previous_word)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    youremail_label = ttk.Label(right_frame, text="Your Email:")
    youremail_label.pack(pady=text_padding)    
    adem_email = ttk.Entry(right_frame, textvariable= email_var)
    adem_email.bind("<Control-BackSpace>", delete_previous_word)
    adem_email.pack(padx=text_padding, pady=text_padding)

    adem_pronoun = ""

    def get_pronoun():
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT Pronoun FROM settings WHERE Name = ?", (adem_employee.get(),))
        data = c.fetchall()
        adem_pronoun = data[0][0]
        print(adem_pronoun)
        get_inspr_values()


    comments_label = ttk.Label(inspr, text="Comments/Site Observations:")
    comments_label.pack(padx=text_padding, pady=text_padding)
    comments = ttk.Text(inspr)
    comments.bind("<Control-BackSpace>", delete_previous_word)
    comments.pack(padx=text_padding, pady=text_padding)

    
    # Button to retrieve input values
    submit_button = ttk.Button(right_frame, text="Submit", command=get_pronoun)
    submit_button.pack(padx=text_padding, pady=text_padding,side=ttk.LEFT)
# END INSPECTION REPORT

#BEGIN FEE SHEET
#Fee Sheet Compiler
def get_feel_values():
    #Import template document
    template = DocxTemplate(r'.\templates\FEEL_Temp.docx')
    Agent_Email = agent_email.get()
    context = {
        
        'Applicant_Honorific': honorific.get(),
        'Applicant_FirstName': first_name.get(),
        'Applicant_LastName': last_name.get(),
        'Applicant_Address': project_address.get(),
        'Applicant_Title': title.get(),
        'Agent_Name': agent_name.get(),
        'Agent_Address': agent_address.get(),
        'ACity': city.get(),
        'AState': state.get(),
        'AZip': zip.get(),
        'Project_Name': project_name.get(),
        'Project_City': project_city.get(),
        'Project_County': project_county.get(),
        'SAM_Number': sam.get(),
        'ACAMP_Number': acamp.get(),
        'FEE_Amount': fee_amount.get(),
        'ADEM_Employee': adem_employee.get(),
        'ADEM_Email': adem_email.get()
    }
    insert_data(acamp.get(), context)

    agent = {
        'name': agent_name.get(),
        'address': agent_address.get(),
        'city': city.get(),
        'state': state.get(),
        'email': ''
    }

    insert_agent_data(agent_name.get(), agent)

    #Render automated report
    body = f"""\
    ACAMP: {acamp.get()}
    SAM: {sam.get()}
    Facility Name: {project_name.get()}"""

    send_email('ADEM Fee Letter: ' + acamp.get(),Agent_Email,body)
    render_document(template,context,context.get('ACAMP_Number'),context.get('SAM_Number'),"","","FEEL")
    

    feel.destroy()


def open_feel_window():
    global feel    
    global honorific, first_name, last_name, title, project_address
    global agent_name, agent_address, agent_email
    global city, state, zip
    global project_name, project_city, project_county
    global fee_amount, projcoords
    global adem_employee, adem_email, sam, acamp
    global timein, timeout, complaint
    global phone, comments, photos, participants
   

    # Fee Letter Window
    feel = ttk.Toplevel()
    feel.title("ADEM Coastal Document Genie")
    feel.iconbitmap(icon)
    feel.bind('<Return>', lambda event: get_feel_values())
    
    left_frame = ttk.Frame(feel, )
    left_frame.pack(side=ttk.LEFT)

    right_frame = ttk.Frame(feel, )
    right_frame.pack(side=ttk.LEFT)
    
    chosen_type = list(document_types.keys())[0]
    subtypes = document_types[chosen_type]
    greeting = ttk.Label(left_frame, text="Please provide the following information: ", )
    greeting.pack(padx=text_padding, pady=text_padding)

    database_button = ttk.Button(left_frame, text = 'Load from Database', command = show_data)
    database_button.pack()


    # Create input fields
    sam_label = ttk.Label(right_frame, text="SAM Number:")
    sam_label.pack(pady=text_padding)
    sam = ttk.Entry(right_frame)
    sam.bind("<Control-BackSpace>", delete_previous_word)
    sam.pack(padx=text_padding, pady=text_padding)

    acamp_label = ttk.Label(right_frame, text="ACAMP Number):")
    acamp_label.pack(pady=text_padding)
    acamp = ttk.Entry(right_frame)
    acamp.bind("<Control-BackSpace>", delete_previous_word)
    acamp.pack(padx=text_padding, pady=text_padding)

    honorific_label = ttk.Label(left_frame, text="Applicant Honorific (Mr./Ms./Dr./etc):")
    honorific_label.pack(pady=text_padding)
    honorific = ttk.Entry(left_frame)
    honorific.bind("<Control-BackSpace>", delete_previous_word)
    honorific.pack(padx=text_padding, pady=text_padding)

    firstname_label = ttk.Label(left_frame, text="Applicant First Name:")
    firstname_label.pack(pady=text_padding)
    first_name = ttk.Entry(left_frame)
    first_name.bind("<Control-BackSpace>", delete_previous_word)
    first_name.pack(padx=text_padding, pady=text_padding)

    lastname_label = ttk.Label(left_frame, text="Applicant Last Name:")
    lastname_label.pack(pady=text_padding)
    last_name = ttk.Entry(left_frame)
    last_name.bind("<Control-BackSpace>", delete_previous_word)
    last_name.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Applicant Address:")
    address_label.pack(pady=text_padding)
    project_address = ttk.Entry(left_frame)
    project_address.bind("<Control-BackSpace>", delete_previous_word)
    project_address.pack(padx=text_padding, pady=text_padding)

    title_label = ttk.Label(left_frame, text="Applicant Title or Company:")
    title_label.pack(pady=text_padding)
    title = ttk.Entry(left_frame)
    title.bind("<Control-BackSpace>", delete_previous_word)
    title.pack(padx=text_padding, pady=text_padding)
    agents = get_agents()
    agent_list = []
    for i in agents:
        agent_list.append(i[0])
    # Create Label
    label2 = ttk.Label(left_frame , text = "Choose Agent: " )
    label2.pack(padx=text_padding, pady=text_padding)  

    clicked1 = ttk.StringVar()

    clicked1.set( "Choose Agent:" )

    drop1 = ttk.OptionMenu( left_frame, clicked1, *agent_list)
    drop1.pack(padx=text_padding, pady=text_padding)

    def callback2(*args):
        for agent in agents:
            for i in agent_list:
                if clicked1.get() == agent[0]:
                    agent_name.delete(0,ttk.END)
                    agent_address.delete(0,ttk.END)
                    agent_address.insert(0, agent[1])
                    agent_name.insert(0, agent[0])
                    city.delete(0, ttk.END)
                    city.insert(0, agent[2])
                    state.delete(0, ttk.END)
                    state.insert(0, agent[3])
                    agent_email.delete(0, ttk.END)
                    agent_email.insert(0, agent[4])
        

    clicked1.trace("w", callback2)

    agentname_label = ttk.Label(left_frame, text="Agent Full Name:")
    agentname_label.pack(pady=text_padding)
    agent_name = ttk.Entry(left_frame)
    agent_name.bind("<Control-BackSpace>", delete_previous_word)
    agent_name.pack(padx=text_padding, pady=text_padding)

    agentaddress_label = ttk.Label(left_frame, text="Agent Address:")
    agentaddress_label.pack(pady=text_padding)
    agent_address = ttk.Entry(left_frame)
    agent_address.bind("<Control-BackSpace>", delete_previous_word)
    agent_address.pack(padx=text_padding, pady=text_padding)

    agentemail_label = ttk.Label(left_frame, text="Agent Email:")
    agentemail_label.pack(pady=text_padding)
    agent_email = ttk.Entry(left_frame)
    agent_email.bind("<Control-BackSpace>", delete_previous_word)
    agent_email.pack(padx=text_padding, pady=text_padding)

    city_label = ttk.Label(left_frame, text="City:")
    city_label.pack(pady=text_padding)
    city = ttk.Entry(left_frame)
    city.bind("<Control-BackSpace>", delete_previous_word)
    city.pack(padx=text_padding, pady=text_padding)

    state_label = ttk.Label(left_frame, text="State:")
    state_label.pack(pady=text_padding)
    state = ttk.Entry(left_frame)
    state.bind("<Control-BackSpace>", delete_previous_word)
    state.pack(padx=text_padding, pady=text_padding)

    zip_label = ttk.Label(left_frame, text="Zip:")
    zip_label.pack(pady=text_padding)
    zip = ttk.Entry(left_frame)
    zip.bind("<Control-BackSpace>", delete_previous_word)
    zip.pack(padx=text_padding, pady=text_padding)

    
    get_zip = ttk.Button(left_frame, text ='Get Zip', command=lambda:zip.insert(0, find_zip(agent_address.get(), city.get())))
    get_zip.pack()

    project_name_label = ttk.Label(right_frame, text="Project Name:")
    project_name_label.pack(pady=text_padding)
    project_name = ttk.Entry(right_frame)
    project_name.bind("<Control-BackSpace>", delete_previous_word)
    project_name.pack(padx=text_padding, pady=text_padding)

    project_city_label = ttk.Label(right_frame, text="Project City:")
    project_city_label.pack(pady=text_padding)
    project_city = ttk.Entry(right_frame)
    project_city.bind("<Control-BackSpace>", delete_previous_word)
    project_city.pack(padx=text_padding, pady=text_padding)

    project_county_label = ttk.Label(right_frame, text="Project County:")
    project_county_label.pack(pady=text_padding)

    project_county = ttk.Combobox(right_frame, values=['Mobile','Baldwin','Washington'])
    project_county.pack(padx=text_padding, pady=text_padding)

    # List of fees
    fees = [
        "Commercial and/or Residential Development",
        "Greater than 5 acres and less than 25 acres",
        "25 acres or greater and less than 100 acres",
        "100 acres or greater",
        "Groundwater extraction from a well having capacity of 50gpm or more",
        "Construction on Beaches and Dunes",
        "Single Family Dwelling or 1 Duplex",
        "Single Family Dwelling or 2 Duplexes",
        "Commercial, multi-unit residential structure >2 units, or any other combination of living units not listed",
        "Hardened erosion control structures (retaining walls, bulkheads, rip-rap, and similar structures)",
        "Beach Nourishment Projects on Gulf Beaches",
        "Filling less than 1,000 square feet of State waterbottoms",
        "Filling 1,000 to 100,000 square feet of State waterbottoms",
        "Filling greater than 100,000 square feet of State waterbottoms",
        "Projects Impacting Wetlands",
        "Dredging or filling of less than 1,000 square feet of wetlands",
        "Dredging or filling of 1,000 square feet or more of wetlands",
        "Pile-supported residential, multifamily, or commercial structures (does not include piers, walkways, gazebos)",
        "Projects Impacting Water Bottoms",
        "Filling of less than 1,000 square feet of water bottom",
        "Filling of 1,000 square feet or more of water bottom",
        "Dredging of less than 10,000 cubic yards from water bottom",
        "Dredging of 10,000 to 100,000 cubic yards from water bottom",
        "Dredging of greater than 100,000 cubic yards from water bottom",
        "Construction of coastal or inland marinas, canals, or creek relocation / Modification",
        "Raised creek crossing",
        "Shoreline Stabilization of Non Gulf-Fronting Properties",
        "Shoreline stabilization less than 200 feet (bulkheads, rip-rap)",
        "Shoreline stabilization greater than 200 feet (bulkheads, rip-rap)",
        "Other",
        "Groins, jetties, and other sediment catching structures",
        "Pile-supported piers, docks, boardwalks, etc.",
        "Siting, construction, and operation of energy facilities",
        "Mitigation Bank projects",
        "State agency permits subject to review, not otherwise specified in Schedule B",
        "Federal licenses or permits not specified in Schedule B",
        "Certification for FERC permit or authorization",
        "All other projects and/or consistency reviews not otherwise specified in Schedule B which are subject to ADEM’s Division 8 regulations",
        "Certification transfer or to change the name of the applicant only",
        "Modifications and/or time extensions not requiring public notice",
        "Modifications and/or time extensions requiring public notice",
        "Variance request (additive)"
    ]

    # Corresponding prices
    prices = [
        "",
        9025,
        19070,
        25020,
        3995,
        "",
        1330,
        1750,
        17765,
        2035,
        "",
        1895,
        3785,
        6985,
        "",
        2125,
        4235,
        3940,
        "",
        2125,
        4235,
        2125,
        4235,
        7855,
        4235,
        800,
        "",
        800,
        1330,
        "",
        1680,
        800,
        24480,
        8730,
        1680,
        1680,
        6550,
        800,
        800,
        800,
        800,
        3275
    ]

    # Create price_map dictionary
    price_map = {fees[i]: prices[i] for i in range(len(fees))}
    feeamount_label = ttk.Label(right_frame, text="Fee Amount Due:")
    feeamount_label.pack(pady=text_padding)
    fee_amount = ttk.Entry(right_frame)
    fee_amount.bind("<Control-BackSpace>", delete_previous_word)
    fee_amount.pack(padx=text_padding, pady=text_padding)
    
        # Function to calculate the fee amount
    def calculate_fee():
        total_fee = 0
        for fee, check_var in check_vars.items():
            if check_var.get():
                total_fee += price_map.get(fee, 0)
        fee_amount.delete(0, tk.END)
        fee_amount.insert(0, str(total_fee))

    

    # Function to create a new window for the combo boxes
    def create_new_window():
        new_window = tk.Toplevel()
        new_window.title("Select Additive Prices")
        new_window.geometry("950x1000")
        canvas = tk.Canvas(new_window)
        scrollbar = tk.Scrollbar(new_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        canvas.config(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        #bind mousewheel
        def on_mousewheel(event):
            canvas.yview_scroll(-1*(event.delta//120), "units")

        # Bind the mouse wheel scrolling to the on_mousewheel function
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        global check_vars
        check_vars = {}
        
        for i, fee in enumerate(fees):
            check_var = tk.IntVar()
            check_vars[fee] = check_var
            label = ttk.Label(scrollable_frame, text=fees[i], font=("Helvetica", 12, "bold"))
            
            if (prices[i] != ""):
                label = ttk.Label(scrollable_frame, text=fees[i], font=("Helvetica", 10))
                check_button = tk.Checkbutton(scrollable_frame, text="", variable=check_var)
                check_button.grid(row=i, column=1,columnspan=2, sticky='w')
            label.grid(row=i, column=0, pady=text_padding)
        
        def close_window():
            calculate_fee()
            new_window.destroy()

        btn_calculate = ttk.Button(scrollable_frame, text="Calculate", command=close_window)
        btn_calculate.grid(row=len(fees), columnspan=2)

        scrollable_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))


    # Button to open new window
    btn_open = ttk.Button(right_frame, text="Calculate Fee", command=create_new_window)
    btn_open.pack(pady=10)
    
    # Create Label
    label1 = ttk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack(padx=text_padding, pady=text_padding)  
    get_data3()
    yourname_label = ttk.Label(right_frame, text="Your Name:")
    yourname_label.pack(pady=text_padding)
    adem_employee = ttk.Entry(right_frame, textvariable= name_var)
    adem_employee.bind("<Control-BackSpace>", delete_previous_word)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    youremail_label = ttk.Label(right_frame, text="Your Email:")
    youremail_label.pack(pady=text_padding)    
    adem_email = ttk.Entry(right_frame, textvariable= email_var)
    adem_email.bind("<Control-BackSpace>", delete_previous_word)
    adem_email.pack(padx=text_padding, pady=text_padding)

    adem_pronoun = ""
    def get_pronoun():
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT Pronoun FROM settings WHERE Name = ?", (adem_employee.get(),))
        data = c.fetchall()
        adem_pronoun = data[0][0]
        print(adem_pronoun)
        get_feel_values()

    
    # Button to retrieve input values
    submit_button = ttk.Button(right_frame, text="Submit", command=get_pronoun)
    submit_button.pack(padx=text_padding, pady=text_padding,side=ttk.LEFT)

#END FEE SHEET

#DATABASE FUNCTIONS
def create_database():
    db_exists = os.path.exists(database)
    #Check if database exists
    if not db_exists:
        # Connect to the database (this will create it)
        conn = sqlite3.connect(database)
        c = conn.cursor()

        # Create the table
        c.execute("""
            CREATE TABLE applicants (
                ACAMP_Number TEXT,
                SAM_Number TEXT,
                Project_Name TEXT,
                Project_Location TEXT,
                Project_City TEXT,
                Project_County TEXT,
                Project_Description TEXT,
                var_code TEXT,
                Parcel_ID TEXT,
                Applicant_Honorific TEXT,
                Applicant_FirstName TEXT,
                Applicant_LastName TEXT,
                Applicant_Address TEXT,
                Applicant_Title TEXT,
                Agent_Name TEXT,
                Agent_Address TEXT,
                ACity TEXT,
                AState TEXT,
                AZip REAL,
                Prefile_Date TEXT,
                Notice_Type TEXT,
                JPN_Date TEXT,
                PNOT_Date TEXT,
                Fee_Amount REAL,
                Fee_Received REAL,
                Expiration_Date TEXT,
                New_Expiration TEXT,
                NPDES_Date TEXT,
                NPDES_Number TEXT,
                ADEM_Employee TEXT,
                ADEM_Email TEXT,
                time_in TEXT,
                time_out TEXT,
                Proj_Cords TEXT,
                Proj_Complaints TEXT,
                Photos TEXT,
                Other_Names TEXT
            )
        """)

        c.execute("""
            CREATE TABLE settings (
                Dark INTEGER,
                Output TEXT,
                Pronoun TEXT,
                First INTEGER,
                Name TEXT,
                Email TEXT
            )
        """)

        c.execute(f"INSERT INTO settings (Dark, Output, First) VALUES (0, '.\output', 0)")

        c.execute("""
            CREATE TABLE agents (
                name TEXT,
                address TEXT,
                city TEXT,
                state TEXT,
                email TEXT
            )
        """)

        c.execute(f"INSERT INTO agents (name) VALUES ('CHOOSE AGENT');")
        c.execute(f'INSERT INTO "agents" ("name", "address", "city", "state", "email") VALUES ("Barry Vittor", "8060 Cottage Hill Road", "Mobile", "Alabama", "bvittor@bvaenviro.com");')
        c.execute(f'INSERT INTO "agents" ("name", "address", "city", "state", "email") VALUES ("Gena Todia", "PO Box 2694", "Daphne", "Alabama", "jaget@zebra.net");')
        c.execute(f'INSERT INTO "agents" ("name", "address", "city", "state", "email") VALUES ("Ecosolutions", "PO Bo 361", "Montrose", "Alabama", "ecosolutionsinc@bellsouth.net");')
        c.execute(f'INSERT INTO "agents" ("name", "address", "city", "state", "email") VALUES ("Cathy Barnette", "25353 Friendship Road", "Daphne", "Alabama", "cbarnette@dewberry.com");')
        # Commit the changes and close the connection
        conn.commit()
        conn.close()


def get_data():
    conn = sqlite3.connect(database)
    c = conn.cursor()
    c.execute("SELECT ACAMP_Number, SAM_Number, Project_Name, Project_Location, Project_City, Project_County FROM applicants")
    data = c.fetchall()
    conn.close()
    return data

def get_data2():
    conn = sqlite3.connect(database)
    c = conn.cursor()
    c.execute("SELECT * FROM applicants")
    data = c.fetchall()
    conn.close()
    return data
global name_var
name_var = ttk.StringVar()
global email_var
email_var = ttk.StringVar()
global pronoun_var
pronoun_var = ttk.StringVar()

def get_data3():
    conn = sqlite3.connect(database)
    c = conn.cursor()
    c.execute("SELECT Name,Email,Pronoun FROM settings")
    data = c.fetchall()
    name_var.set(data[0][0])
    email_var.set(data[0][1])
    pronoun_var.set(data[0][2])
    conn.close()
    return data

def insert_data(acamp_number, data):
    # Connect to the database
    conn = sqlite3.connect(database)
    c = conn.cursor()

    # Check if a row with the matching ACAMP number exists
    c.execute("SELECT * FROM applicants WHERE ACAMP_Number=?", (acamp_number,))
    row = c.fetchone()
    keys = list(data.keys())
    last_key = keys[-1]  # Get the last key

    del data[last_key]  # Remove the last key-value pair
    
    if row is not None:
        # If a match is found, update the row with new data
        columns = ', '.join(f"{column}=?" for column in data.keys())
        sql = f"UPDATE applicants SET {columns} WHERE ACAMP_Number=?"
        c.execute(sql, list(data.values()) + [acamp_number])
    else:
        # If no match is found, insert a new row with the new data
        columns = ', '.join(data.keys())
        placeholders = ', '.join('?' * len(data))
        sql = f"INSERT INTO applicants ({columns}) VALUES ({placeholders})"
        c.execute(sql, list(data.values()))

    # Commit the changes and close the connection
    conn.commit()
    conn.close()

def insert_agent_data(agent_name, data):
    # Connect to the database
    conn = sqlite3.connect(database)
    c = conn.cursor()

    # Check if a row with the matching ACAMP number exists
    c.execute("SELECT * FROM agents WHERE name=?", (agent_name,))
    row = c.fetchone()

    if row is not None:
        # If a match is found, update the row with new data
        columns = ', '.join(f"{column}=?" for column in data.keys())
        sql = f"UPDATE agents SET {columns} WHERE name=?"
        c.execute(sql, list(data.values()) + [agent_name])
    else:
        # If no match is found, insert a new row with the new data
        columns = ', '.join(data.keys())
        placeholders = ', '.join('?' * len(data))
        sql = f"INSERT INTO agents ({columns}) VALUES ({placeholders})"
        c.execute(sql, list(data.values()))

    # Commit the changes and close the connection
    conn.commit()
    conn.close()
    

def show_data():
    data = ttk.Toplevel()
    tree = ttk.Treeview(data)    

    # Define the column names
    columns = ['ACAMP #', 'SAM #', 'Project Name', 'Address', 'City', 'County']
    
    search_label = ttk.Label(data, text="Search")
    search_val = ttk.StringVar()
    search = ttk.Entry(data, textvariable=search_val)
    search_label.pack()
    search.pack()

    # Create the columns
    tree["columns"] = columns
    for column in columns:
        tree.column(column, width=100)
        tree.heading(column, text=column)
        tree.heading(column,anchor=tk.W)
    
    tree.column('#0', width=0)

    def replace_field(widget,text):
        if text != None:
            if isinstance(widget, ttk.Entry):
                widget.delete(0,tk.END)
                widget.insert(0,text)
            else:
                widget.delete(1.0,tk.END)
                widget.insert(1.0,text)
        else:
            if isinstance(widget, ttk.Entry):
                widget.insert(0,"")
            else:
                widget.insert(1.0,"")
    

    def onDoubleClick(event):
        global honorific, first_name, last_name, title, project_address
        global agent_name, agent_address
        global city, state, zip, var_code
        global project_name, project_city, project_county
        global fee_amount, projcoords, fee_received
        global adem_employee, adem_email, sam, acamp
        global timein, timeout, complaint, project_description
        global phone, comments, photos, participants, parcel_id
        item = tree.selection()[0]
        values=tree.item(item,"values")

        for datum in get_data2():
            if datum[0] == values[0]:
                try:    
                    if acamp.winfo_exists:
                        replace_field(acamp, datum[0])
                except NameError:
                    pass
                except AttributeError:
                    pass 
                except Exception:
                    pass  
                try:        
                    if sam.winfo_exists:
                        replace_field(sam, datum[1])
                except NameError:
                    pass
                except AttributeError:
                    pass 
                except Exception:
                    pass   
                try:
                    if project_name.winfo_exists:
                        replace_field(project_name, datum[2])
                except NameError:
                    pass
                except AttributeError:
                    pass 
                except Exception:
                    pass   
                try:    
                    if honorific.winfo_exists:
                        replace_field(honorific, datum[9])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass    
                try:
                    if first_name.winfo_exists:
                        replace_field(first_name, datum[10])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:    
                    if last_name.winfo_exists:
                        replace_field(last_name, datum[11])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:    
                    if project_address.winfo_exists:
                        replace_field(project_address, datum[3])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:    
                    if title.winfo_exists:
                        replace_field(title, datum[13])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:
                    if agent_name.winfo_exists:
                        replace_field(agent_name, datum[14])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:
                    if agent_address.winfo_exists:
                        replace_field(agent_address, datum[15])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:
                    if city.winfo_exists:
                        replace_field(city, datum[16])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:
                    if state.winfo_exists:
                        replace_field(state, datum[17])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass      
                try:
                    if zip.winfo_exists:
                        replace_field(zip, datum[18])
                except NameError:
                    pass
                except AttributeError:
                    pass
                except Exception:
                    pass
                try:
                    if parcel_id.winfo_exists:
                        replace_field(parcel_id, datum[8])
                except NameError as e:
                    print('NE', e)
                except AttributeError:
                    print('AE', e)  
                except Exception:
                    print('E', e)      
                try:
                    if project_city.winfo_exists:
                        replace_field(project_city, datum[4])
                except NameError:
                    pass
                except AttributeError:
                    pass  
                except Exception:
                    pass    
                try:
                    if project_county.winfo_exists:
                        replace_field(project_county, datum[5])
                except NameError:
                    pass
                except AttributeError:
                    pass 
                except Exception:
                    pass  
                try:
                    if fee_amount.winfo_exists:
                        replace_field(fee_amount, datum[23])
                except NameError:
                    pass
                except AttributeError:
                    pass    
                except Exception:
                    pass 
                try:
                    if fee_received.winfo_exists:
                        replace_field(fee_received, datum[24])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass   
                try:
                    if adem_email.winfo_exists:
                        replace_field(adem_email, datum[30])
                except NameError:
                    pass
                except AttributeError:
                    pass   
                except Exception:
                    pass   
                try:
                    if adem_employee.winfo_exists:
                        replace_field(adem_employee, datum[29])
                except NameError:
                    pass
                except AttributeError:
                    pass   
                except Exception:
                    pass    
                try:
                    if project_description.winfo_exists:
                        replace_field(project_description, datum[6])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass  
                try:
                    if prefile_date.winfo_exists:
                        replace_field(prefile_date, datum[19])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass
                try:
                    if notice_type.winfo_exists:
                        replace_field(notice_type, datum[20])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass  
                try:
                    if pnot_date.winfo_exists:
                        replace_field(pnot_date, datum[22])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass 
                try:
                    if jpn_date.winfo_exists:
                        replace_field(jpn_date, datum[21])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass 
                try:
                    if var_code.winfo_exists:
                        replace_field(var_code, datum[7])
                except NameError:
                    pass
                except AttributeError:
                    pass      
                except Exception:
                    pass 

    def onDel(event):
        try:
            item = tree.selection()[0]
            values = tree.item(item, "values")
            
            conn = sqlite3.connect(database)
            c = conn.cursor()

            # Print the SQL query for debugging
            delete_query = "DELETE FROM applicants WHERE ACAMP_Number = ?"
            
            c.execute(delete_query, (values[0],))
            conn.commit()
            conn.close()

            # Delete the selected item from the Treeview
            tree.delete(item)

        except sqlite3.Error as e:
            print("SQLite error:", e)
        except IndexError:
            print("No item selected.")
        except Exception as e:
            print("An error occurred:", e)
        

    def search_data(*args):
        search_term = search_val.get()
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT ACAMP_Number, SAM_Number, Project_Name, Project_Location, Project_City, Project_County FROM applicants WHERE ACAMP_Number LIKE ? OR SAM_Number LIKE ?", ('%'+search_term+'%', '%'+search_term+'%'))
        data = c.fetchall()
        conn.close()
        return data

    def on_search_change(*args):
        if search_val.get():
            data = search_data(args)
            tree.delete(*tree.get_children())
            for row in data:
                tree.insert('', 'end', values=row)
        else:
            data = get_data()
            tree.delete(*tree.get_children())
            for row in data:
                tree.insert('', 'end', values=row)
        tree.pack()
        tree.bind("<Double-1>", onDoubleClick)


    search_val.trace(mode="w", callback = on_search_change)
                        

        
    # Insert the data
    data = get_data()
    for row in data:
        tree.insert('', 'end', values=row)
    
    tree.pack()
    tree.bind("<Double-1>", onDoubleClick)
    tree.bind("<Delete>",onDel)

#Main Screen Contents
greeting = ttk.Label(text="What do you want to generate?")
greeting.pack(padx=text_padding, pady=text_padding)
create_database()

for i, document_type in enumerate(document_types.keys()):
    document_button = ttk.Button(main, text=f"{i+1}. {document_type}")
    document_button.pack(padx=text_padding, pady=text_padding)
    if document_type == "Public Notice":
        document_button.configure(command=open_pnot_window)
    elif document_type == "Permit":
        document_button.configure(command=open_perm_window)
    elif document_type == "Fee Letter":
        document_button.configure(command=open_feel_window)
    elif document_type == "Inspection Report":
        document_button.configure(command=open_inspr_window)
def open_employee_window():
    employee_window = ttk.Toplevel()
    employee_window.title("ADEM Coastal Document Genie")
    employee_window.iconbitmap(icon)
    get_data3()
    greeting = ttk.Label(employee_window,text="Please input your information:").pack(padx=text_padding, pady=text_padding)
    name_label = ttk.Label(employee_window, text="Your Full Name: ").pack(padx=text_padding, pady=text_padding)
    name = ttk.Entry(employee_window, textvariable=name_var)
    name.pack(padx=text_padding, pady=text_padding)
    email_label = ttk.Label(employee_window, text="Your Email: ").pack(padx=text_padding, pady=text_padding)
    email = ttk.Entry(employee_window, textvariable=email_var)
    email.pack(padx=text_padding, pady=text_padding)
    email_label2 = ttk.Label(employee_window, text="@adem.alabama.gov").pack(pady=text_padding)
    pronoun_label = ttk.Label(employee_window, text="Your Preferred Pronoun: ").pack(padx=text_padding, pady=text_padding)
    pronoun = ttk.Entry(employee_window, textvariable=pronoun_var)
    pronoun.pack(padx=text_padding, pady=text_padding)

    def write_employee_info():
        try:
            conn = sqlite3.connect(database)
            c = conn.cursor()
            c.execute("UPDATE settings SET Name=?", (name.get(),))
            c.execute("UPDATE settings SET Pronoun=?", (pronoun.get(),))
            c.execute("UPDATE settings SET Email=?", (email.get(),))
            conn.commit()
        except sqlite3.Error as e:
            print("SQLite error:", e)
        except Exception as e:
            print("An error occurred:", e)
        finally:
            conn.close()
            employee_window.destroy()
    save = ttk.Button(employee_window, text='Save', command=lambda: write_employee_info()).pack(padx=text_padding, pady=text_padding)

def open_options_window():
    options = ttk.Toplevel()
    options.title("ADEM Coastal Document Genie")
    options.iconbitmap(icon)
    greeting = ttk.Label(options, text="Please Choose an Option Below:").pack(padx=text_padding, pady=text_padding)
    database_button = ttk.Button(options, text = 'View Database', command = show_data)
    set_employee_data = ttk.Button(options, text='Input Employee Information', command=open_employee_window).pack(padx=text_padding, pady=text_padding)
    set_output_folder = ttk.Button(options, text = 'Select Output Folder', command=open_file_dialog).pack(padx=text_padding, pady=text_padding)
    output_button = ttk.Button(options, text = 'Open Output Folder', command = openFolder)
    output_button.pack(padx=text_padding, pady=text_padding)
    database_button.pack(padx=text_padding, pady=text_padding)
    darkmode = ttk.Checkbutton(options, text='Dark Mode', variable=windowcolor, onvalue = 'darkly', offvalue='yeti', command=toggle_dark_mode)
    darkmode.pack(padx=text_padding, pady=text_padding)
    try:
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT Dark FROM settings")
        data = c.fetchall()
        if data[0][0] == 1:
            windowcolor.set('darkly')
    except sqlite3.Error as e:
        print("SQLite error:", e)
    except Exception as e:
        print("An error occurred:", e)
    finally:
        conn.close()

options_button = ttk.Button(main,text = 'Options', command = open_options_window, bootstyle="warning")
options_button.pack(padx=text_padding, pady=text_padding)

def get_agents():
    try:
        conn = sqlite3.connect(r'.\database.db')
        c = conn.cursor()
        c.execute("SELECT * FROM agents")
        data = c.fetchall()
        return data
    except sqlite3.Error as e:
        print('SQLite error:', e)
    except Exception as e:
        print('Error:', e)
    finally:
        conn.close()

global agents 
agents = get_agents()
def open_first():
    first = ttk.Toplevel()
    first.title("ADEM Coastal Document Genie")
    first.iconbitmap(icon)
    greeting = ttk.Label(first, text="Enter Your Information. If you'd like, choose a custom Project Path. ").pack(padx=text_padding, pady=text_padding)
    close = ttk.Button(first, text = "Close", command=lambda:first.destroy())
    close.pack(padx=text_padding, pady=text_padding)

def check_settings():
    global output_path
    try:
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT First FROM settings")
        data = c.fetchall()
        print(data[0][0])
        if data[0][0] == 0:
            #Run Settings Menu
            open_options_window()
            open_employee_window()
            open_first()
            c.execute("UPDATE settings SET First=1")
            conn.commit()
    except sqlite3.Error as e:
        print("SQLITE Error", e)
    except Exception as e:
        print("Error", e)
    finally:
        conn.close()
        

    try:
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT Dark FROM settings")
        data = c.fetchall()
        if data[0][0] == 1:
            style.theme_use('darkly')
    except sqlite3.Error as e:
        print("SQLite error:", e)
    except Exception as e:
        print("An error occurred:", e)
    finally:
        conn.close()
    try:
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT Output FROM settings")
        data = c.fetchall()
        output_path = data[0][0]
    except sqlite3.Error as e:
        print(e)
    finally:
        conn.close()

# Call the function to check and apply dark mode if needed
check_settings()

main.mainloop()
