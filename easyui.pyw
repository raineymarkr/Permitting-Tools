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

# Main Configuration

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
icon = r"free.ico"
#list of active permitters
permitters = {
    0:["Choose",''],
    1:["Mark Rainey","mark.rainey"],
    2:["Katie Smith", "katiem.smith"],
    3:["Sarila Mickle", "sarila.mickle"],
    4:["Autumn Nitz", "autumn.nitz"]
}

text_padding = 5
main = ttk.Window(themename='yeti')
main.iconbitmap(icon)
main.title("ADEM Coastal Document Genie")
windowcolor = tk.StringVar()
windowcolor.set('yeti')
style = ttk.Style()
countynum = ""

#UTILITY FUNCTIONS
def toggle_dark_mode():
    conn = sqlite3.connect('database.db')
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
    print('opened /' + filename)  

def display_pdf(pdf_path):
    # Convert PDF to images
    images = convert_from_path(pdf_path, 500, poppler_path=r'.\poppler-0.68.0\bin')
    
    # Select the first image from the list
    image = images[0]

    # Create a Tkinter window
    feesheet = ttk.Toplevel()
    feesheet.title("Fee Sheet")
    feesheet.geometry('800x850')
    feesheet.iconbitmap(icon)

    # Create a Canvas widget to display the PDF pages
    canvas = ttk.Canvas(feesheet, bg="white")
    canvas.pack(fill=ttk.BOTH, expand=True)

    

    # Define the desired display size
    desired_width = 800  # Adjust the width as needed
    desired_height = 950  # Adjust the height as needed
    feesheet.config(width=desired_width, height=desired_height)

    # Resize the image to fit the desired display size
    image = image.resize((desired_width, desired_height))

    # Display the PDF pages as images on the Canvas widget
    img = ImageTk.PhotoImage(image)
    canvas.img = img  # Store the reference to the image object
    canvas.create_image(0, 0, anchor=ttk.NW, image=img)

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
    filename ='output/xxx ' + acamp +' '+ countynum +' ' +str(date)+ ' ' + perm_type +' '+ sam +' '+ doc_type +'.docx'
    template.save(filename.format(acamp))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

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
        print(f"An error occurred: {e}")



#BEGIN PNOT WINDOW
def open_pnotinput_window():
    global pnot1
    global honorific, first_name, last_name, title, project_address
    global agent_name, agent_address
    global city, state
    global project_name, project_city, project_county
    global fee_amount, projcoords
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
    project_county = ttk.Entry(left_frame)
    project_county.pack(padx=text_padding, pady=text_padding)
    project_county.bind("<Control-BackSpace>", delete_previous_word)

    variancecodes_label = ttk.Label(left_frame, text="Variance Codes:")
    var_code = ttk.Entry(left_frame)
    var_code.bind("<Control-BackSpace>", delete_previous_word)

    parcelid_label = ttk.Label(left_frame, text="Parcel ID:")
    parcel_id = ttk.Entry(left_frame)
    parcel_id.bind("<Control-BackSpace>", delete_previous_word)

    if pnottype == "VAR":        
        variancecodes_label.pack(padx=text_padding, pady=text_padding)
        parcel_id.pack(padx=text_padding, pady=text_padding)
        parcelid_label.pack(padx=text_padding, pady=text_padding)
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
    template = DocxTemplate('templates/BSEEPNOT_Temp.docx')
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
    template = DocxTemplate('templates/VARPNOT_Temp.docx')
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
    template = DocxTemplate('templates/NRUPNOT_Temp.docx')
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
    template = DocxTemplate('templates/FAAPNOT_Temp.docx')
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
    template = DocxTemplate('templates/LOPPNOT_Temp.docx')
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
    template = DocxTemplate('templates/OCSPNOT_Temp.docx')
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

#BEGIN perm WINDOW
def open_perminput_window():
    global perm1
    global honorific, first_name, last_name, title, project_address
    global agent_name, agent_address
    global city, state, zip, project_description
    global project_name, project_city, project_county
    global fee_amount, projcoords
    global adem_employee, adem_email, sam, acamp
    global timein, timeout, complaint
    global phone, comments, photos, participants
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

    project_name_label = ttk.Label(middle_frame, text="Project Name:")
    project_name_label.pack(padx=text_padding, pady=text_padding)
    project_name = ttk.Entry(middle_frame)
    project_name.bind("<Control-BackSpace>", delete_previous_word)
    project_name.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Project Address:")
    address_label.pack(padx=text_padding, pady=text_padding)
    project_address = ttk.Entry(left_frame)
    project_address.bind("<Control-BackSpace>", delete_previous_word)
    project_address.pack(padx=text_padding, pady=text_padding)

    project_city_label = ttk.Label(middle_frame, text="Project City:")
    project_city_label.pack(padx=text_padding, pady=text_padding)
    project_city = ttk.Entry(middle_frame)
    project_city.bind("<Control-BackSpace>", delete_previous_word)
    project_city.pack(padx=text_padding, pady=text_padding)

    project_county_label = ttk.Label(middle_frame, text="Project County:")
    project_county_label.pack(padx=text_padding, pady=text_padding)
    project_county = ttk.Entry(middle_frame)
    project_county.bind("<Control-BackSpace>", delete_previous_word)
    project_county.pack(padx=text_padding, pady=text_padding)

    parcel_id_label = ttk.Label(middle_frame, text="Parcel ID:")
    parcel_id_label.pack(padx=text_padding, pady=text_padding)
    parcel_id = ttk.Entry(middle_frame)
    parcel_id.bind("<Control-BackSpace>", delete_previous_word)
    parcel_id.pack(padx=text_padding, pady=text_padding)

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

    var_code_label = ttk.Label(middle_frame, text="Variance from code:")
    var_code = ttk.Entry(middle_frame)
    var_code.bind("<Control-BackSpace>", delete_previous_word)

    if permtype == "VAR":
        var_code_label.pack(padx=text_padding, pady=text_padding)
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

    permitter_list = []
    for i in permitters:
        permitter_list.append(permitters.get(i)[0])
    
    # Create Label
    label1 = ttk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack(padx=text_padding, pady=text_padding, side=ttk.LEFT)  

    clicked = ttk.StringVar()

    clicked.set( "Choose ADEM Permitter:" )

    drop = ttk.OptionMenu( right_frame, clicked, *permitter_list)
    drop.pack(padx=text_padding, pady=text_padding, side=ttk.LEFT)

    def callback(*args):
        for i in permitters:
            if clicked.get() == permitters.get(i)[0]:
                adem_email.delete(0,ttk.END)
                adem_employee.delete(0,ttk.END)
                adem_employee.insert(0, permitters[i][0])
                adem_email.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    adem_employee_label = ttk.Label(right_frame, text="ADEM Employee:")
    adem_employee_label.pack(padx=text_padding, pady=text_padding)
    adem_employee = ttk.Entry(right_frame)
    adem_employee.bind("<Control-BackSpace>", delete_previous_word)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    adem_email_label = ttk.Label(right_frame, text="ADEM Email:")
    adem_email_label.pack(padx=text_padding, pady=text_padding)
    adem_email = ttk.Entry(right_frame)
    adem_email.bind("<Control-BackSpace>", delete_previous_word)
    adem_email.pack(padx=text_padding, pady=text_padding)

    submit_button = ttk.Button(right_frame, text="Submit", command=lambda: get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), project_address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, ttk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get(),var_code.get()))
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
    template = DocxTemplate('templates/401WQC_Temp.docx')
    template2 = DocxTemplate('templates/401Rat_Temp.docx')

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
        'ADEM_Email': adem_email
    }
    
    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 002'
    else:
        countynum = ' xxx'

    insert_data(acamp, context)
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
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
        'ADEM_Email': adem_email
    }

    if project_county.lower() == 'mobile':
        countynum = ' 097'
    elif project_county.lower() == 'baldwin':
        countynum = ' 002'
    else:
        countynum = ' xxx'
    insert_data(acamp, context)
    # Render automated report
    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
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
        'var_code': var_code
    }

    insert_data(acamp, context)
    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review (VARIANCE): ' + acamp,'CMcNeill@adem.alabama.gov',body)
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
        'NPDES_Number': npdes_num
    }
    insert_data(acamp, context)
    # Render automated report
    body = f"""\
    ACAMP: {acamp}
    SAM: {sam}
    Facility Name: {project_name}
    Summary: {project_description}"""
    send_email('For Review: ' + acamp,'CMcNeill@adem.alabama.gov',body)
    render_document(templaten, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="NRU")
    render_document(templatec, context, acamp, sam, county=project_county ,perm_type="CZCERT", doc_type="CZM")
    render_document(template2,context,acamp,sam,project_county,"CZCERT","RATIONALE")


def perm_TIMEEXT(acamp, sam, honorific, first_name, last_name, project_address, title, agent_name, agent_address, city, state, zip, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1):
    # Import template document
    template = DocxTemplate('templates/401EXT_Temp.docx')
    
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
        'New_Expiration': exp_date1
    }
    insert_data(acamp, context)
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
    template = DocxTemplate('templates/NPR_Temp.docx')
    
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
    }
    insert_data(acamp, context)
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
    template = DocxTemplate('templates/Insp_Temp.docx')


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
    project_county = ttk.Entry(right_frame)
    project_county.bind("<Control-BackSpace>", delete_previous_word)
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

    permitter_list = []
    for i in permitters:
        permitter_list.append(permitters.get(i)[0])
    
    # Create Label
    label1 = ttk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack(padx=text_padding, pady=text_padding)  

    clicked = ttk.StringVar()

    clicked.set( "Choose ADEM Permitter:" )

    drop = ttk.OptionMenu( right_frame, clicked, *permitter_list)
    drop.pack(padx=text_padding, pady=text_padding)

    def callback(*args):
        for i in permitters:
            if clicked.get() == permitters.get(i)[0]:
                adem_email.delete(0,ttk.END)
                adem_employee.delete(0,ttk.END)
                adem_employee.insert(0, permitters[i][0])
                adem_email.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    yourname_label = ttk.Label(right_frame, text="Your Name:")
    yourname_label.pack(padx=text_padding, pady=text_padding)
    adem_employee = ttk.Entry(right_frame)
    adem_employee.bind("<Control-BackSpace>", delete_previous_word)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    youremail_label = ttk.Label(right_frame, text="Your Email:")
    youremail_label.pack(padx=text_padding, pady=text_padding)
    adem_email = ttk.Entry(right_frame)
    adem_email.bind("<Control-BackSpace>", delete_previous_word)
    adem_email.pack(padx=text_padding, pady=text_padding)


    comments_label = ttk.Label(inspr, text="Comments/Site Observations:")
    comments_label.pack(padx=text_padding, pady=text_padding)
    comments = ttk.Text(inspr)
    comments.bind("<Control-BackSpace>", delete_previous_word)
    comments.pack(padx=text_padding, pady=text_padding)

    
    # Button to retrieve input values
    submit_button = ttk.Button(inspr, text="Submit", command=get_inspr_values)
    submit_button.pack(padx=text_padding, pady=text_padding)

# END INSPECTION REPORT

#BEGIN FEE SHEET
#Fee Sheet Compiler
def get_feel_values():
    #Import template document
    template = DocxTemplate('templates/FEEL_Temp.docx')
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
    project_county = ttk.Entry(right_frame)
    project_county.bind("<Control-BackSpace>", delete_previous_word)
    project_county.pack(padx=text_padding, pady=text_padding)

    feeamount_label = ttk.Label(right_frame, text="Fee Amount Due:")
    feeamount_label.pack(pady=text_padding)
    fee_amount = ttk.Entry(right_frame)
    fee_amount.bind("<Control-BackSpace>", delete_previous_word)
    fee_amount.pack(padx=text_padding, pady=text_padding)

    # Button to retrieve input values
    file_path=".\Fee_List.pdf"
    feel_button = ttk.Button(right_frame, text="Fee List", command=lambda: display_pdf(file_path))
    #feel_button = ttk.Button(right_frame, text="Fee List", command=lambda: subprocess.Popen(['start', '', file_path], shell=True))
    feel_button.pack(padx=text_padding, pady=text_padding)

    permitter_list = []
    for i in permitters:
        permitter_list.append(permitters.get(i)[0])
    
    # Create Label
    label1 = ttk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack(padx=text_padding, pady=text_padding)  

    clicked = ttk.StringVar()

    clicked.set( "Choose ADEM Permitter:" )

    drop = ttk.OptionMenu( right_frame, clicked, *permitter_list)
    drop.pack(padx=text_padding, pady=text_padding,side=ttk.LEFT)

    def callback(*args):
        for i in permitters:
            if clicked.get() == permitters.get(i)[0]:
                adem_email.delete(0,ttk.END)
                adem_employee.delete(0,ttk.END)
                adem_employee.insert(0, permitters[i][0])
                adem_email.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    yourname_label = ttk.Label(right_frame, text="Your Name:")
    yourname_label.pack(pady=text_padding)
    adem_employee = ttk.Entry(right_frame)
    adem_employee.bind("<Control-BackSpace>", delete_previous_word)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    youremail_label = ttk.Label(right_frame, text="Your Email:")
    youremail_label.pack(pady=text_padding)    
    adem_email = ttk.Entry(right_frame)
    adem_email.bind("<Control-BackSpace>", delete_previous_word)
    adem_email.pack(padx=text_padding, pady=text_padding)
    
    # Button to retrieve input values
    submit_button = ttk.Button(right_frame, text="Submit", command=get_feel_values)
    submit_button.pack(padx=text_padding, pady=text_padding,side=ttk.LEFT)
 
#END FEE SHEET

#DATABASE FUNCTIONS
def create_database():
    db_exists = os.path.exists('database.db')
    if not db_exists:
        # Connect to the database (this will create it)
        conn = sqlite3.connect('database.db')
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
                Dark INTEGER
            )
        """)

        c.execute(f"INSERT INTO settings (Dark) VALUES (0)")

        # Commit the changes and close the connection
        conn.commit()
        conn.close()

def get_data():
    conn = sqlite3.connect('database.db')
    c = conn.cursor()
    c.execute("SELECT ACAMP_Number, SAM_Number, Project_Name, Project_Location, Project_City, Project_County FROM applicants")
    data = c.fetchall()
    conn.close()
    return data

def get_data2():
    conn = sqlite3.connect('database.db')
    c = conn.cursor()
    c.execute("SELECT * FROM applicants")
    data = c.fetchall()
    conn.close()
    return data

def insert_data(acamp_number, data):
    # Connect to the database
    conn = sqlite3.connect('database.db')
    c = conn.cursor()

    # Check if a row with the matching ACAMP number exists
    c.execute("SELECT * FROM applicants WHERE ACAMP_Number=?", (acamp_number,))
    row = c.fetchone()

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
        global city, state, zip
        global project_name, project_city, project_county
        global fee_amount, projcoords
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
                    if parcel_id.winfo_exists:
                        replace_field(parcel_id, datum[8])
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

    def onDel(event):
        try:
            item = tree.selection()[0]
            values = tree.item(item, "values")
            
            conn = sqlite3.connect('database.db')
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
        conn = sqlite3.connect('database.db')
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
        print(search_val.get())
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

def open_options_window():
    options = ttk.Toplevel()
    options.title("ADEM Coastal Document Genie")
    options.iconbitmap(icon)
    greeting = ttk.Label(options, text="Please Choose an Option Below:").pack(padx=text_padding, pady=text_padding)
    database_button = ttk.Button(options, text = 'View Database', command = show_data)
    database_button.pack(padx=text_padding, pady=text_padding)
    def openFolder():
        os.startfile(os.path.normpath('output'))
    output_button = ttk.Button(options, text = 'Open Output Folder', command = openFolder)
    output_button.pack(padx=text_padding, pady=text_padding)

    darkmode = ttk.Checkbutton(options, text='Dark Mode', variable=windowcolor, onvalue = 'darkly', offvalue='yeti', command=toggle_dark_mode)
    darkmode.pack(padx=text_padding, pady=text_padding)
    try:
        conn = sqlite3.connect('database.db')
        c = conn.cursor()
        c.execute("SELECT Dark FROM settings")
        data = c.fetchall()
        print(data[0][0])
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

def check_dark_mode():
    try:
        conn = sqlite3.connect('database.db')
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

# Call the function to check and apply dark mode if needed
check_dark_mode()


main.mainloop()
