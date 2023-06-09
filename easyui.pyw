import tkinter as tk
import ttkbootstrap as ttk
import datetime
import subprocess
from docx.shared import Cm
from docxtpl import DocxTemplate
from PIL import ImageTk, Image
from pdf2image import convert_from_path

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
        "Time Extension": None
    }
}

#list of active permitters
permitters = {
    0:["Choose",''],
    1:["Mark Rainey","mark.rainey"],
    2:["Katie Smith", "katiem.smith"],
    3:["Sarila Mickle", "sarila.mickle"],
    4:["Autumn Nitz", "autumn.nitz"]
}

# Main Configuration

text_padding = 5
main = ttk.Window(themename='yeti')
main.title("ADEM Coastal Document Genie")
windowcolor = tk.StringVar()
windowcolor.set('yeti')
main.iconbitmap("free.ico")

def toggle_dark_mode():
    if windowcolor.get() == 'darkly':
        style.theme_use('darkly')
    else:
        style.theme_use('yeti')

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

#BEGIN PNOT WINDOW
def open_pnotinput_window():
    global pnot1
    pnot1 = ttk.Toplevel()
    pnot1.title("ADEM Coastal Document Genie")
    pnot1.iconbitmap("free.ico")

    pnot1.bind('<Return>', lambda event: get_pnot_values(acamp.get(), sam.get(), project_name.get(), project_location.get(), project_city.get(), project_county.get(), project_desc.get(1.0, ttk.END), var_code.get(), parcel_id.get(), federal_agency.get()))

    left_frame = ttk.Frame(pnot1, )
    left_frame.pack(side=ttk.LEFT, padx=10)

    right_frame = ttk.Frame(pnot1, )
    right_frame.pack(side=ttk.LEFT, padx=10)

    greeting = ttk.Label(left_frame, text="Please provide the following information:", )

    greeting.pack(padx=text_padding, pady=text_padding)
    
    acamp_label = ttk.Label(left_frame, text="ACAMP Number:")
    acamp_label.pack(padx=text_padding, pady=text_padding)
    acamp = ttk.Entry(left_frame)
    acamp.pack(padx=text_padding, pady=text_padding)

    sam = ttk.Entry(left_frame)

    if pnottype != "BSSE" and pnottype != "FAA" and pnottype != "OCS" :
        sam_label = ttk.Label(left_frame, text="SAM Number:")
        sam_label.pack(padx=text_padding, pady=text_padding)
        
        sam.pack(padx=text_padding, pady=text_padding)

    project_name_label = ttk.Label(left_frame, text="Project Name:")
    project_name_label.pack(padx=text_padding, pady=text_padding)
    project_name = ttk.Entry(left_frame)
    project_name.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Project Address/Location:")
    address_label.pack(padx=text_padding, pady=text_padding)
    project_location = ttk.Entry(left_frame)
    project_location.pack(padx=text_padding, pady=text_padding)

    projectcity_label = ttk.Label(left_frame, text="Project City:")
    projectcity_label.pack(padx=text_padding, pady=text_padding)
    project_city = ttk.Entry(left_frame)
    project_city.pack(padx=text_padding, pady=text_padding)

    projectcounty_label = ttk.Label(left_frame, text="Project County:")
    projectcounty_label.pack(padx=text_padding, pady=text_padding)
    project_county = ttk.Entry(left_frame)
    project_county.pack(padx=text_padding, pady=text_padding)

    variancecodes_label = ttk.Label(left_frame, text="Variance Codes:")
    
    var_code = ttk.Entry(left_frame)

    parcelid_label = ttk.Label(left_frame, text="Parcel ID:")
    
    parcel_id = ttk.Entry(left_frame)

    if pnottype == "VAR":        
        variancecodes_label.pack(padx=text_padding, pady=text_padding)
        parcel_id.pack(padx=text_padding, pady=text_padding)
        parcelid_label.pack(padx=text_padding, pady=text_padding)
        var_code.pack(padx=text_padding, pady=text_padding)

    federal_agency = ttk.Entry(left_frame)
    
    if pnottype == "FAA":
        fedagency_label = ttk.Label(left_frame, text="Federal Agency:")
        fedagency_label.pack(padx=text_padding, pady=text_padding)
        
        federal_agency.pack(padx=text_padding, pady=text_padding)
    
    project_desc_label = ttk.Label(right_frame, text="Project Description:")
    project_desc_label.pack(padx=text_padding, pady=text_padding)
    project_desc = ttk.Text(right_frame)
    project_desc.pack(padx=text_padding, pady=text_padding)

    submit_button = ttk.Button(right_frame, text="Submit", command=lambda: get_pnot_values(acamp.get(), sam.get(), project_name.get(), project_location.get(), project_city.get(), project_county.get(), project_desc.get(1.0, ttk.END), var_code.get(), parcel_id.get(), federal_agency.get()))
    submit_button.pack(padx=text_padding, pady=text_padding)

def get_pnot_values(acamp, sam="", project_name="", project_location="", project_city="", project_county="", project_desc="", var_code="", parcel_id="", federal_agency=""):
    if pnottype == "IP":
        pnot_LOP(acamp, sam, project_name, project_location, project_city, project_county, project_desc)
    elif pnottype == "LOP":
        pnot_LOP(acamp, sam, project_name, project_location, project_city, project_county, project_desc)
    elif pnottype == "VAR":
        pnot_VAR(acamp, sam, project_name, project_location, project_city, project_county, project_desc, var_code, parcel_id)
    elif pnottype == "NRU":
        pnot_NRU(acamp, sam, project_name, project_location, project_city, project_county, project_desc)
    elif pnottype == "BSSE":
        pnot_BSSE(acamp, project_name, project_location, project_city, project_county, project_desc)
    elif pnottype == "FAA":
        pnot_FAA(acamp, project_location, project_city, project_county, federal_agency, project_desc)
    elif pnottype == "OCS":
        pnot_OCS(acamp, project_name, project_location, project_desc)
    
    pnot.destroy()
    pnot1.destroy()

def pnot_BSSE(acamp, project_name, project_location, project_city, project_county, project_desc):
    template = DocxTemplate('templates/BSEEPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'Project_Name': project_name,
        'Project_Location': project_location,
        'Project_Description': project_desc,
        'Project_City': project_city,
        'Project_County': project_county
    }
    template.render(context)
    filename ='output/' + acamp+' BSSE_PNOT.docx'
    template.save(filename.format(acamp))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

def pnot_VAR(acamp, sam, project_name, project_location, project_city, project_county,project_desc,var_code,parcel_id):
    template = DocxTemplate('templates/VARPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'SAM_Number': sam,
        'Project_Name': project_name,
        'Project_Location': project_location,
        'Project_City': project_city,
        'Project_County': project_county,
        'Parcel_ID': parcel_id,
        'Project_Description': project_desc,
        'var_code': var_code,
    }
    template.render(context)
    filename ='output/' + acamp+' ' + sam +' VAR_PNOT.docx'
    template.save(filename.format(acamp, sam))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

def pnot_NRU(acamp, sam, project_name, project_location, project_city, project_county,project_desc):
    template = DocxTemplate('templates/NRUPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'SAM_Number': sam,
        'Project_Name': project_name,
        'Project_Location': project_location,
        'Project_City': project_city,
        'Project_County': project_county,
        'Project_Description': project_desc,
    }
    template.render(context)
    filename ='output/' + acamp+' ' + sam +' NRU_PNOT.docx'
    template.save(filename.format(acamp, sam))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

def pnot_FAA(acamp, project_location, project_city, project_county, federal_agency, project_desc):
    template = DocxTemplate('templates/FAAPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'Federal_Agency': federal_agency,
        'Project_Location': project_location,
        'Project_City': project_city,
        'Project_County': project_county,
        'Project_Description': project_desc,
    }
    template.render(context)
    filename = 'output/' + acamp+' ' + sam +' FAA_PNOT.docx'
    template.save(filename.format(acamp, sam))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

def pnot_LOP(acamp, sam, project_name, project_location, project_city, project_county,project_desc):
    template = DocxTemplate('templates/LOPPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'SAM_Number': sam,
        'Project_Name': project_name,
        'Project_Location': project_location,
        'Project_City': project_city,
        'Project_County': project_county,
        'Project_Description': project_desc,
    }
    template.render(context)
    filename ='output/' + acamp+' ' + sam +' LOP_PNOT.docx'
    template.save(filename.format(acamp, sam))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

def pnot_OCS(acamp, project_name, project_location, project_desc):
    template = DocxTemplate('templates/OCSPNOT_Temp.docx')
    context = {
        'ACAMP_Number': acamp,
        'Project_Name': project_name,
        'Project_Location': project_location,
        'Project_Description': project_desc,
    }
    template.render(context)
    filename ='output/' + acamp+' ' +' OCS_PNOT.docx'
    template.save(filename.format(acamp))
    open_file(filename)
    print("Files successfully generated in /output/ folder.")

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
    pnot.iconbitmap("free.ico")
    chosen_type = list(document_types.keys())[2]
    subtypes = document_types[chosen_type]
    greeting = ttk.Label(pnot, text="What type of Public Notice do you want to generate?")
    greeting.pack(padx=text_padding, pady=text_padding)

    for i, document_type in enumerate(subtypes.keys()):
        document_button = ttk.Button(pnot, text=f"{i+1}. {document_type}")
        document_button.pack(padx=text_padding, pady=text_padding)
        document_button.configure(command=lambda doc_type=document_type: set_pnottype(doc_type))
    
    #END  PNOT WINDOW


#BEGIN perm WINDOW
def open_perminput_window():
    global perm1
    perm1 = ttk.Toplevel()
    perm1.title("ADEM Coastal Document Genie")
    perm.iconbitmap("free.ico")
    perm1.bind('<Return>', lambda event: get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip_code.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, ttk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get(), var_code.get()))

    left_frame = ttk.Frame(perm1, )
    left_frame.pack(side=ttk.LEFT)

    middle_frame = ttk.Frame(perm1, )
    middle_frame.pack(side=ttk.LEFT)

    right_frame = ttk.Frame(perm1, )
    right_frame.pack(side=ttk.LEFT)

    greeting = ttk.Label(left_frame, text="Please provide the following information:")
    greeting.pack(padx=text_padding, pady=text_padding)

    acamp_label = ttk.Label(middle_frame, text="ACAMP Number:")
    acamp_label.pack(padx=text_padding, pady=text_padding)
    acamp = ttk.Entry(middle_frame)
    acamp.pack(padx=text_padding, pady=text_padding)

    sam_label = ttk.Label(middle_frame, text="SAM Number:")
    sam_label.pack(padx=text_padding, pady=text_padding)
    sam = ttk.Entry(middle_frame)
    sam.pack(padx=text_padding, pady=text_padding)

    honorific_label = ttk.Label(left_frame, text="Honorific:")
    honorific_label.pack(padx=text_padding, pady=text_padding)
    honorific = ttk.Entry(left_frame)
    honorific.pack(padx=text_padding, pady=text_padding)

    first_name_label = ttk.Label(left_frame, text="First Name:")
    first_name_label.pack(padx=text_padding, pady=text_padding)
    first_name = ttk.Entry(left_frame)
    first_name.pack(padx=text_padding, pady=text_padding)

    last_name_label = ttk.Label(left_frame, text="Last Name:")
    last_name_label.pack(padx=text_padding, pady=text_padding)
    last_name = ttk.Entry(left_frame)
    last_name.pack(padx=text_padding, pady=text_padding)

    title_label = ttk.Label(left_frame, text="Title:")
    title_label.pack(padx=text_padding, pady=text_padding)
    title = ttk.Entry(left_frame)
    title.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Applicant Address:")
    address_label.pack(padx=text_padding, pady=text_padding)
    address = ttk.Entry(left_frame)
    address.pack(padx=text_padding, pady=text_padding)

    agent_name_label = ttk.Label(left_frame, text="Agent Full Name:")
    agent_name_label.pack(padx=text_padding, pady=text_padding)
    agent_name = ttk.Entry(left_frame)
    agent_name.pack(padx=text_padding, pady=text_padding)

    agent_address_label = ttk.Label(left_frame, text="Agent Address:")
    agent_address_label.pack(padx=text_padding, pady=text_padding)
    agent_address = ttk.Entry(left_frame)
    agent_address.pack(padx=text_padding, pady=text_padding)

    city_label = ttk.Label(left_frame, text="City:")
    city_label.pack(padx=text_padding, pady=text_padding)
    city = ttk.Entry(left_frame)
    city.pack(padx=text_padding, pady=text_padding)

    state_label = ttk.Label(left_frame, text="State:")
    state_label.pack(padx=text_padding, pady=text_padding)
    state = ttk.Entry(left_frame)
    state.pack(padx=text_padding, pady=text_padding)

    zip_code_label = ttk.Label(left_frame, text="Zip Code:")
    zip_code_label.pack(padx=text_padding, pady=text_padding)
    zip_code = ttk.Entry(left_frame)
    zip_code.pack(padx=text_padding, pady=text_padding)

    project_name_label = ttk.Label(middle_frame, text="Project Name:")
    project_name_label.pack(padx=text_padding, pady=text_padding)
    project_name = ttk.Entry(middle_frame)
    project_name.pack(padx=text_padding, pady=text_padding)

    project_city_label = ttk.Label(middle_frame, text="Project City:")
    project_city_label.pack(padx=text_padding, pady=text_padding)
    project_city = ttk.Entry(middle_frame)
    project_city.pack(padx=text_padding, pady=text_padding)

    project_county_label = ttk.Label(middle_frame, text="Project County:")
    project_county_label.pack(padx=text_padding, pady=text_padding)
    project_county = ttk.Entry(middle_frame)
    project_county.pack(padx=text_padding, pady=text_padding)

    parcel_id_label = ttk.Label(middle_frame, text="Parcel ID:")
    parcel_id_label.pack(padx=text_padding, pady=text_padding)
    parcel_id = ttk.Entry(middle_frame)
    parcel_id.pack(padx=text_padding, pady=text_padding)

    prefile_date_label = ttk.Label(middle_frame, text="Prefile Date:")
    prefile_date_label.pack(padx=text_padding, pady=text_padding)
    prefile_date = ttk.Entry(middle_frame)
    prefile_date.pack(padx=text_padding, pady=text_padding)

    notice_type_label = ttk.Label(middle_frame, text="Notice Type:")
    notice_type_label.pack(padx=text_padding, pady=text_padding)
    notice_type = ttk.Entry(middle_frame)
    notice_type.pack(padx=text_padding, pady=text_padding)

    jpn_date_label = ttk.Label(middle_frame, text="USACE JPN Date:")
    jpn_date_label.pack(padx=text_padding, pady=text_padding)
    jpn_date = ttk.Entry(middle_frame)
    jpn_date.pack(padx=text_padding, pady=text_padding)

    pnot_date_label = ttk.Label(middle_frame, text="ADEM PNOT Date:")
    pnot_date_label.pack(padx=text_padding, pady=text_padding)
    pnot_date = ttk.Entry(middle_frame)
    pnot_date.pack(padx=text_padding, pady=text_padding)

    exp_date_label = ttk.Label(middle_frame, text="Expiration Date:")
    exp_date = ttk.Entry(middle_frame)
    exp_date1_label = ttk.Label(middle_frame, text="New Expiration Date:")
    exp_date1 = ttk.Entry(middle_frame)
    
    if permtype == "Time Extension":
        exp_date_label.pack(padx=text_padding, pady=text_padding)
        exp_date.pack(padx=text_padding, pady=text_padding)
        exp_date1_label.pack(padx=text_padding, pady=text_padding)
        exp_date1.pack(padx=text_padding, pady=text_padding)

    var_code_label = ttk.Label(middle_frame, text="Variance from code:")
    var_code = ttk.Entry(middle_frame)

    if permtype == "VAR":
        var_code_label.pack(padx=text_padding, pady=text_padding)
        var_code.pack(padx=text_padding, pady=text_padding)

    npdes_num_label = ttk.Label(middle_frame, text="NPDES Permit:")
    npdes_num = ttk.Entry(middle_frame)
    npdes_date_label = ttk.Label(middle_frame, text="NPDES Permit Date:")
    npdes_date = ttk.Entry(middle_frame)
    parcel_size_label = ttk.Label(middle_frame,text="Parcel Size (Ac):")
    parcel_size = ttk.Entry(middle_frame)
    
    if permtype == "NRU":
        npdes_num_label.pack(padx=text_padding, pady=text_padding)
        npdes_num.pack(padx=text_padding, pady=text_padding)
        npdes_date_label.pack(padx=text_padding, pady=text_padding)
        npdes_date.pack(padx=text_padding, pady=text_padding)
        parcel_size_label.pack(padx=text_padding, pady=text_padding)
        parcel_size.pack(padx=text_padding, pady=text_padding)

    prefile_label = ttk.Label(middle_frame, text="Prefile Date:")
    prefile_date = ttk.Entry(middle_frame)

    if permtype == "IP":
        prefile_label.pack(padx=text_padding, pady=text_padding)
        prefile_date.pack(padx=text_padding, pady=text_padding)
    
    fee_amount_label = ttk.Label(right_frame, text="Fee Amount:")
    fee_amount_label.pack(padx=text_padding, pady=text_padding)
    fee_amount = ttk.Entry(right_frame)
    fee_amount.pack(padx=text_padding, pady=text_padding)

    fee_received_label = ttk.Label(right_frame, text="Fee Received:")
    fee_received_label.pack(padx=text_padding, pady=text_padding)
    fee_received = ttk.Entry(right_frame)
    fee_received.pack(padx=text_padding, pady=text_padding)

    project_description_label = ttk.Label(right_frame, text="Project Description:")
    project_description_label.pack(padx=text_padding, pady=text_padding)
    project_description = ttk.Text(right_frame)
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
                print(permitters[i])
                adem_email.delete(0,ttk.END)
                adem_employee.delete(0,ttk.END)
                adem_employee.insert(0, permitters[i][0])
                adem_email.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    adem_employee_label = ttk.Label(right_frame, text="ADEM Employee:")
    adem_employee_label.pack(padx=text_padding, pady=text_padding)
    adem_employee = ttk.Entry(right_frame)
    adem_employee.pack(padx=text_padding, pady=text_padding)

    adem_email_label = ttk.Label(right_frame, text="ADEM Email:")
    adem_email_label.pack(padx=text_padding, pady=text_padding)
    adem_email = ttk.Entry(right_frame)
    adem_email.pack(padx=text_padding, pady=text_padding)

    submit_button = ttk.Button(right_frame, text="Submit", command=lambda: get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip_code.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, ttk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get(),var_code.get()))
    submit_button.pack(padx=text_padding, pady=text_padding)


def get_perm_values(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1, npdes_date, npdes_num, parcel_size, var_code):
    if permtype == "IP":
        perm_LOP(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email)
    elif permtype == "LOP":
        perm_LOP(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email)
    elif permtype == "VAR":
        perm_VAR(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email,var_code)
    elif permtype == "NRU":
        perm_NRU(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, npdes_date,npdes_num)
    elif permtype == "401":
        perm_401(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email)
    elif permtype == "TIMEEXT":
        perm_TIMEEXT(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1)    
    
    perm.destroy()
    perm1.destroy()

def perm_401(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email):
# Import template document
    template = DocxTemplate('templates/401WQC_Temp.docx')
    template2 = DocxTemplate('templates/401Rat_Temp.docx')

    # Declare template variables
    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip_code,
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

    # Render automated report
    template.render(context)
    template.save('output/' + acamp + ' ' + sam + '_401WQC_Docs.docx')
    template2.render(context)
    template2.save('output/' + acamp + ' ' + sam + '401Rat_Docs.docx')

    open_file('output/' + acamp + ' ' + sam +  '_401WQC_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '401Rat_Docs.docx')

    print("Files successfully generated in /output/ folder.")

def perm_LOP(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email):
    # Import template document
    templatePerm1 = DocxTemplate('templates/LOPW_Temp.docx')
    templatePerm2 = DocxTemplate('templates/LOPC_Temp.docx')
    templateRat = DocxTemplate('templates/LOPRat_Temp.docx')
    
    # Declare template variables
    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip_code,
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

    # Render automated report
    templatePerm1.render(context)
    templatePerm1.save('output/' + acamp + ' ' + sam + '_CZM_Docs.docx')
    templatePerm2.render(context)
    templatePerm2.save('output/' + acamp + ' ' + sam + '_401WQ_Docs.docx')
    templateRat.render(context)
    templateRat.save('output/' + acamp + ' ' + sam + '_LOP_Rational.docx')
    open_file('output/' + acamp + ' ' + sam + '_CZM_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '_401WQC_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '_LOP_Rational.docx')
    print("Files successfully generated in the 'output' folder.")

def perm_VAR(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email,var_code):
    # Import template document
    templatePerm1 = DocxTemplate('templates/LOPW_Temp.docx')
    templatePerm2 = DocxTemplate('templates/VARC_Temp.docx')
    templateRat = DocxTemplate('templates/LOPRat_Temp.docx')
    
    # Declare template variables
    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip_code,
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
        'Parcel_ID': parcel_id
    }

    # Render automated report
    #Render automated report
    templatePerm1.render(context)
    templatePerm1.save('output/'+acamp+' ' + sam +'_401WQ_Docs.docx')
    templatePerm2.render(context)
    templatePerm2.save('output/' + acamp + ' ' + sam + '_VAR_Docs.docx')
    templateRat.render(context)
    templateRat.save('output/' + acamp + ' ' + sam + '_LOP_Rational.docx')
    open_file('output/' + acamp + ' ' + sam + '_VAR_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '_401WQC_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '_LOP_Rational.docx')
    print("Files successfully generated in the 'output' folder.")


def perm_NRU(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email,npdes_date,npdes_num):
    # Import template document
    templaten = DocxTemplate('templates/NRU_Temp.docx')
    templatec = DocxTemplate('templates/LOPC_Temp.docx')
    template2 = DocxTemplate('templates/NRURat_Temp.docx')
    
    # Declare template variables
    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip_code,
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

    # Render automated report
    templaten.render(context)
    templaten.save('output/' + acamp + ' ' + sam + '_NRU_Docs.docx')
    templatec.render(context)
    templatec.save('output/' + acamp + ' ' + sam + '_CZM_Docs.docx')
    template2.render(context)
    template2.save('output/' + acamp + ' ' + sam + '_NRU_Rational.docx')

    open_file('output/' + acamp + ' ' + sam + '_CZM_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '_NRU_Docs.docx')
    open_file('output/' + acamp + ' ' + sam + '_NRU_Rational.docx')
    
    print("Files successfully generated in the 'output' folder.")


def perm_TIMEEXT(acamp, sam, honorific, first_name, last_name, address, title, agent_name, agent_address, city, state, zip_code, project_name, project_city, project_county, parcel_id, prefile_date, notice_type, jpn_date, pnot_date, project_description, fee_amount, fee_received, adem_employee, adem_email, exp_date, exp_date1):
    # Import template document
    template = DocxTemplate('templates/401EXT_Temp.docx')
    
    # Declare template variables
    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'Applicant_Honorific': honorific,
        'Applicant_FirstName': first_name,
        'Applicant_LastName': last_name,
        'Applicant_Address': address,
        'Applicant_Title': title,
        'Agent_Name': agent_name,
        'Agent_Address': agent_address,
        'ACity': city,
        'AState': state,
        'AZip': zip_code,
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

    # Render automated report
    template.render(context)
    template.save('output/' + acamp + ' ' + sam + '_TimeExt_Docs.docx')
    
    open_file('output/' + acamp + ' ' + sam + '_TimeExt_Docs.docx')
    
    print("Files successfully generated in the 'output' folder.")


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
    perm.iconbitmap("free.ico")
    chosen_type = list(document_types.keys())[3]
    subtypes = document_types[chosen_type]
    greeting = ttk.Label(perm, text="What type of Permit do you want to generate?")
    greeting.pack(padx=text_padding, pady=text_padding)

    for i, document_type in enumerate(subtypes.keys()):
        document_button = ttk.Button(perm, text=f"{i+1}. {document_type}")
        document_button.pack(padx=text_padding, pady=text_padding)
        document_button.configure(command=lambda doc_type=document_type: set_permtype(doc_type))
    
    #END  perm WINDOW

# BEGIN INSPECTION REPORT
def get_inspr_values():
    # Import template document
    template = DocxTemplate('templates/Insp_Temp.docx')

    # Declare template variables
    
    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'time_in': timein.get(),
        'time_out': timeout.get(),
        'Applicant_FirstName': firstname.get(),
        'Applicant_LastName': lastname.get(),
        'Applicant_Phone': phone.get(),
        'Applicant_Address': address.get(),
        'Proj_Cords': projcoords.get(),
        'Proj_Complaint': complaint.get(),
        'Project_Name': projectname.get(),
        'Project_City': projectcity.get(),
        'Project_County': projectcounty.get(),
        'SAM_Number': sam.get(),
        'ACAMP_Number': acamp.get(),
        'Project_Desc': comments.get(1.0, ttk.END),
        'Photos': photos.get(),
        'Other_Names': participants.get(),
        'ADEM_Employee': yourname.get(),
        'ADEM_Email': youremail.get()
    }

    # Render automated report
    template.render(context)
    template.save('output/xxx ' + " " + context.get('ACAMP_Number') + " " + context.get('SAM_Number') + ' _INSP.docx')
    inspr.destroy()
    print("Files successfully generated in /output/ folder.")

# Inspection Report Window
def open_inspr_window():

    global honorific, firstname, lastname, title, address
    global timein, timeout, complaint, projcoords
    global city, state, zip
    global projectname, projectcity, projectcounty
    global phone, comments, photos, participants
    global yourname, youremail, sam, acamp
    global inspr

    # Inspection Report Window
    inspr = ttk.Toplevel()
    inspr.title("ADEM Coastal Document Genie")
    inspr.iconbitmap("free.ico")
    inspr.bind('<Return>', lambda event: get_inspr_values())
    greeting = ttk.Label(inspr, text="Please provide the following information:")
    greeting.pack(padx=text_padding, pady=text_padding)

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
    acamp.pack(padx=text_padding, pady=text_padding)

    sam_label = ttk.Label(left_frame, text="SAM Number:")
    sam_label.pack(padx=text_padding, pady=text_padding)
    sam = ttk.Entry(left_frame)
    sam.pack(padx=text_padding, pady=text_padding)

    timein_label = ttk.Label(left_frame, text="Inspection Time Start:")
    timein_label.pack(padx=text_padding, pady=text_padding)
    timein = ttk.Entry(left_frame)
    timein.pack(padx=text_padding, pady=text_padding)

    timeout_label = ttk.Label(left_frame, text="Inspection Time End:")
    timeout_label.pack(padx=text_padding, pady=text_padding)
    timeout = ttk.Entry(left_frame)
    timeout.pack(padx=text_padding, pady=text_padding)

    firstname_label = ttk.Label(left_frame, text="Applicant First Name:")
    firstname_label.pack(padx=text_padding, pady=text_padding)
    firstname = ttk.Entry(left_frame)
    firstname.pack(padx=text_padding, pady=text_padding)

    lastname_label = ttk.Label(left_frame, text="Applicant Last Name:")
    lastname_label.pack(padx=text_padding, pady=text_padding)
    lastname = ttk.Entry(left_frame)
    lastname.pack(padx=text_padding, pady=text_padding)

    complaint_label = ttk.Label(left_frame, text="Complaint #:")
    complaint_label.pack(padx=text_padding, pady=text_padding)
    complaint = ttk.Entry(left_frame)
    complaint.pack(padx=text_padding, pady=text_padding)

    # Entry fields with labels in right column
    phone_label = ttk.Label(left_frame, text="Applicant Phone Number:")
    phone_label.pack(padx=text_padding, pady=text_padding)
    phone = ttk.Entry(left_frame)
    phone.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Applicant Address:")
    address_label.pack(padx=text_padding, pady=text_padding)
    address = ttk.Entry(left_frame)
    address.pack(padx=text_padding, pady=text_padding)

    projcoords_label = ttk.Label(right_frame, text="Project Coordinates:")
    projcoords_label.pack(padx=text_padding, pady=text_padding)
    projcoords = ttk.Entry(right_frame)
    projcoords.pack(padx=text_padding, pady=text_padding)

    projectname_label = ttk.Label(right_frame, text="Project Name:")
    projectname_label.pack(padx=text_padding, pady=text_padding)
    projectname = ttk.Entry(right_frame)
    projectname.pack(padx=text_padding, pady=text_padding)

    projectcity_label = ttk.Label(right_frame, text="Project City:")
    projectcity_label.pack(padx=text_padding, pady=text_padding)
    projectcity = ttk.Entry(right_frame)
    projectcity.pack(padx=text_padding, pady=text_padding)

    projectcounty_label = ttk.Label(right_frame, text="Project County:")
    projectcounty_label.pack(padx=text_padding, pady=text_padding)
    projectcounty = ttk.Entry(right_frame)
    projectcounty.pack(padx=text_padding, pady=text_padding)

    photos_label = ttk.Label(right_frame, text="Photos Taken? (Yes/No):")
    photos_label.pack(padx=text_padding, pady=text_padding)
    photos = ttk.Entry(right_frame)
    photos.pack(padx=text_padding, pady=text_padding)

    participants_label = ttk.Label(right_frame, text="Other Participants (Name, Org):")
    participants_label.pack(padx=text_padding, pady=text_padding)
    participants = ttk.Entry(right_frame)
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
                print(permitters[i])
                youremail.delete(0,ttk.END)
                yourname.delete(0,ttk.END)
                yourname.insert(0, permitters[i][0])
                youremail.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    yourname_label = ttk.Label(right_frame, text="Your Name:")
    yourname_label.pack(padx=text_padding, pady=text_padding)
    yourname = ttk.Entry(right_frame)
    yourname.pack(padx=text_padding, pady=text_padding)

    youremail_label = ttk.Label(right_frame, text="Your Email:")
    youremail_label.pack(padx=text_padding, pady=text_padding)
    youremail = ttk.Entry(right_frame)
    youremail.pack(padx=text_padding, pady=text_padding)


    comments_label = ttk.Label(inspr, text="Comments/Site Observations:")
    comments_label.pack(padx=text_padding, pady=text_padding)
    comments = ttk.Text(inspr)
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

    context = {
        'title': 'Automated Report',
        'day': datetime.datetime.now().strftime('%d'),
        'month': datetime.datetime.now().strftime('%b'),
        'year': datetime.datetime.now().strftime('%Y'),
        'Applicant_Honorific': honorific.get(),
        'Applicant_FirstName': firstname.get(),
        'Applicant_LastName': lastname.get(),
        'Applicant_Address': address.get(),
        'Applicant_Title': title.get(),
        'Agent_Name': agentname.get(),
        'Agent_Address': agentaddress.get(),
        'ACity': city.get(),
        'AState': state.get(),
        'AZip': zip.get(),
        'Project_Name': projectname.get(),
        'Project_City': projectcity.get(),
        'Project_County': projectcounty.get(),
        'SAM_Number': sam.get(),
        'ACAMP_Number': acamp.get(),
        'FEE_Amount': feeamount.get(),
        'ADEM_Employee': yourname.get(),
        'ADEM_Email': youremail.get()
    }
    #Render automated report
    template.render(context)
    template.save('output/XXX ' + " " + context.get('ACAMP_Number') + " " + context.get('SAM_Number') + ' _FEEL.docx')
    open_file('output/XXX ' + " " + context.get('ACAMP_Number') + " " + context.get('SAM_Number') + ' _FEEL.docx')

    feel.destroy()
    print("Files successfully generated in /output/ folder.")

def open_feel_window():
    global feel    

    # Fee Letter Window
    feel = ttk.Toplevel()
    feel.title("ADEM Coastal Document Genie")
    feel.iconbitmap("free.ico")
    feel.bind('<Return>', lambda event: get_feel_values())
    
    left_frame = ttk.Frame(feel, )
    left_frame.pack(side=ttk.LEFT)

    right_frame = ttk.Frame(feel, )
    right_frame.pack(side=ttk.LEFT)
    
    chosen_type = list(document_types.keys())[0]
    subtypes = document_types[chosen_type]
    greeting = ttk.Label(left_frame, text="Please provide the following information: ", )
    greeting.pack(padx=text_padding, pady=text_padding)
    global honorific, firstname, lastname, title, address
    global agentname, agentaddress
    global city, state, zip
    global projectname, projectcity, projectcounty
    global feeamount
    global yourname, youremail, sam, acamp

    # Create input fields
    sam_label = ttk.Label(right_frame, text="SAM Number:")
    sam_label.pack(pady=text_padding)
    sam = ttk.Entry(right_frame)
    sam.pack(padx=text_padding, pady=text_padding)

    acamp_label = ttk.Label(right_frame, text="ACAMP Number):")
    acamp_label.pack(pady=text_padding)
    acamp = ttk.Entry(right_frame)
    acamp.pack(padx=text_padding, pady=text_padding)

    honorific_label = ttk.Label(left_frame, text="Applicant Honorific (Mr./Ms./Dr./etc):")
    honorific_label.pack(pady=text_padding)
    honorific = ttk.Entry(left_frame)
    honorific.pack(padx=text_padding, pady=text_padding)

    firstname_label = ttk.Label(left_frame, text="Applicant First Name:")
    firstname_label.pack(pady=text_padding)
    firstname = ttk.Entry(left_frame)
    firstname.pack(padx=text_padding, pady=text_padding)

    lastname_label = ttk.Label(left_frame, text="Applicant Last Name:")
    lastname_label.pack(pady=text_padding)
    lastname = ttk.Entry(left_frame)
    lastname.pack(padx=text_padding, pady=text_padding)

    address_label = ttk.Label(left_frame, text="Applicant Address:")
    address_label.pack(pady=text_padding)
    address = ttk.Entry(left_frame)
    address.pack(padx=text_padding, pady=text_padding)

    title_label = ttk.Label(left_frame, text="Applicant Title or Company:")
    title_label.pack(pady=text_padding)
    title = ttk.Entry(left_frame)
    title.pack(padx=text_padding, pady=text_padding)

    agentname_label = ttk.Label(left_frame, text="Agent Full Name:")
    agentname_label.pack(pady=text_padding)
    agentname = ttk.Entry(left_frame)
    agentname.pack(padx=text_padding, pady=text_padding)

    agentaddress_label = ttk.Label(left_frame, text="Agent Address:")
    agentaddress_label.pack(pady=text_padding)
    agentaddress = ttk.Entry(left_frame)
    agentaddress.pack(padx=text_padding, pady=text_padding)

    city_label = ttk.Label(left_frame, text="City:")
    city_label.pack(pady=text_padding)
    city = ttk.Entry(left_frame)
    city.pack(padx=text_padding, pady=text_padding)

    state_label = ttk.Label(left_frame, text="State:")
    state_label.pack(pady=text_padding)
    state = ttk.Entry(left_frame)
    state.pack(padx=text_padding, pady=text_padding)

    zip_label = ttk.Label(left_frame, text="Zip:")
    zip_label.pack(pady=text_padding)
    zip = ttk.Entry(left_frame)
    zip.pack(padx=text_padding, pady=text_padding)

    projectname_label = ttk.Label(right_frame, text="Project Name:")
    projectname_label.pack(pady=text_padding)
    projectname = ttk.Entry(right_frame)
    projectname.pack(padx=text_padding, pady=text_padding)

    projectcity_label = ttk.Label(right_frame, text="Project City:")
    projectcity_label.pack(pady=text_padding)
    projectcity = ttk.Entry(right_frame)
    projectcity.pack(padx=text_padding, pady=text_padding)

    projectcounty_label = ttk.Label(right_frame, text="Project County:")
    projectcounty_label.pack(pady=text_padding)
    projectcounty = ttk.Entry(right_frame)
    projectcounty.pack(padx=text_padding, pady=text_padding)

    feeamount_label = ttk.Label(right_frame, text="Fee Amount Due:")
    feeamount_label.pack(pady=text_padding)
    feeamount = ttk.Entry(right_frame)
    feeamount.pack(padx=text_padding, pady=text_padding)

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
                print(permitters[i])
                youremail.delete(0,ttk.END)
                yourname.delete(0,ttk.END)
                yourname.insert(0, permitters[i][0])
                youremail.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    yourname_label = ttk.Label(right_frame, text="Your Name:")
    yourname_label.pack(pady=text_padding)
    yourname = ttk.Entry(right_frame)
    yourname.pack(padx=text_padding, pady=text_padding)

    youremail_label = ttk.Label(right_frame, text="Your Email:")
    youremail_label.pack(pady=text_padding)    
    youremail = ttk.Entry(right_frame)
    youremail.pack(padx=text_padding, pady=text_padding)
    
    # Button to retrieve input values
    submit_button = ttk.Button(right_frame, text="Submit", command=get_feel_values)
    submit_button.pack(padx=text_padding, pady=text_padding,side=ttk.LEFT)

    
#END FEE SHEET



#Main Screen Contents
greeting = ttk.Label(text="What do you want to generate?")
greeting.pack(padx=text_padding, pady=text_padding)

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

style = ttk.Style()
darkmode = ttk.Checkbutton(main, text='Dark Mode', variable=windowcolor, onvalue = 'darkly', offvalue='yeti', command=toggle_dark_mode)
darkmode.pack()

main.mainloop()
