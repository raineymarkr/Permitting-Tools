import tkinter as tk
import datetime
from docx.shared import Cm
from docxtpl import DocxTemplate

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

permitters = {
    0:["Mark Rainey","mark.rainey"],
    1:["Katie Smith", "katiem.smith"],
    2:["Sarila Mickle", "sarila.mickle"],
    3:["Autumn Nitz", "autumn.nitz"]
}

# Main Configuration
bgcol = "#c2c2d6"
button_padding = 10
text_padding = 5
main = tk.Tk()
main.configure(background=bgcol, width=1000, height=1000)
main.title("ADEM Coastal Document Genie")
#main.iconbitmap("free.ico")


#BEGIN PNOT WINDOW
def open_pnotinput_window():
    global pnot1
    pnot1 = tk.Toplevel()
    pnot1.bind('<Return>', lambda event: get_pnot_values(acamp.get(), sam.get(), project_name.get(), project_location.get(), project_city.get(), project_county.get(), project_desc.get(1.0, tk.END), var_code.get(), parcel_id.get(), federal_agency.get()))
    pnot1.configure(background=bgcol, width=200, height=500)

    left_frame = tk.Frame(pnot1, bg=bgcol)
    left_frame.pack(side=tk.LEFT, padx=10)

    right_frame = tk.Frame(pnot1, bg=bgcol)
    right_frame.pack(side=tk.LEFT, padx=10)

    greeting = tk.Label(pnot1, text="Please provide the following information:", bg=bgcol, height=5, width=100)
    greeting.pack()
    
    acamp_label = tk.Label(left_frame, text="ACAMP Number:")
    acamp_label.pack()
    acamp = tk.Entry(left_frame)
    acamp.pack(pady=button_padding)

    sam = tk.Entry(left_frame)

    if pnottype != "BSSE" and pnottype != "FAA" and pnottype != "OCS" :
        sam_label = tk.Label(left_frame, text="SAM Number:")
        sam_label.pack()
        
        sam.pack(pady=button_padding)

    project_name_label = tk.Label(left_frame, text="Project Name:")
    project_name_label.pack()
    project_name = tk.Entry(left_frame)
    project_name.pack(pady=button_padding)

    address_label = tk.Label(left_frame, text="Project Address/Location:")
    address_label.pack()
    project_location = tk.Entry(left_frame)
    project_location.pack(pady=button_padding)

    projectcity_label = tk.Label(left_frame, text="Project City:")
    projectcity_label.pack()
    project_city = tk.Entry(left_frame)
    project_city.pack(pady=button_padding)

    projectcounty_label = tk.Label(left_frame, text="Project County:")
    projectcounty_label.pack()
    project_county = tk.Entry(left_frame)
    project_county.pack(pady=button_padding)

    variancecodes_label = tk.Label(right_frame, text="Variance Codes:")
    
    var_code = tk.Entry(right_frame)

    parcelid_label = tk.Label(right_frame, text="Parcel ID:")
    
    parcel_id = tk.Entry(right_frame)

    if pnottype == "VAR":        
        variancecodes_label.pack()
        parcel_id.pack(pady=button_padding)
        parcelid_label.pack()
        var_code.pack(pady=button_padding)

    federal_agency = tk.Entry(right_frame)
    
    if pnottype == "FAA":
        fedagency_label = tk.Label(right_frame, text="Federal Agency:")
        fedagency_label.pack()
        
        federal_agency.pack(pady=button_padding)
    
    project_desc_label = tk.Label(right_frame, text="Project Description:")
    project_desc_label.pack()
    project_desc = tk.Text(right_frame, width=50, height=10)
    project_desc.pack(pady=button_padding)

    submit_button = tk.Button(pnot1, text="Submit", command=lambda: get_pnot_values(acamp.get(), sam.get(), project_name.get(), project_location.get(), project_city.get(), project_county.get(), project_desc.get(1.0, tk.END), var_code.get(), parcel_id.get(), federal_agency.get()))
    submit_button.pack(pady=button_padding)

def get_pnot_values(acamp, sam, project_name, project_location, project_city, project_county, project_desc, var_code="", parcel_id="", federal_agency=""):
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
    template.save('output/xxx_BSSE_PNOT.docx'.format(acamp))
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
    template.save('output/_VAR_PNOT.docx')
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
    template.save('output/_NRU_PNOT.docx')
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
    template.save('output/_FAA_PNOT.docx')
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
    template.save('output/_LOP_PNOT.docx'.format(acamp, sam))
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
    template.save('output/_OCS_PNOT.docx')
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
    pnot = tk.Toplevel()
    pnot.configure(background=bgcol)
    chosen_type = list(document_types.keys())[2]
    subtypes = document_types[chosen_type]
    greeting = tk.Label(pnot, text="What type of Public Notice do you want to generate?", bg=bgcol, height=5, width=100)
    greeting.pack()

    for i, document_type in enumerate(subtypes.keys()):
        document_button = tk.Button(pnot, text=f"{i+1}. {document_type}", padx=button_padding, pady=button_padding)
        document_button.pack(pady=button_padding)
        document_button.configure(command=lambda doc_type=document_type: set_pnottype(doc_type))
    
    #END  PNOT WINDOW


#BEGIN perm WINDOW
def open_perminput_window():
    global perm1
    perm1 = tk.Toplevel()
    perm1.bind('<Return>', lambda event: get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip_code.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, tk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get(), var_code.get()))
    perm1.configure(background=bgcol)

    left_frame = tk.Frame(perm1, bg=bgcol)
    left_frame.pack(side=tk.LEFT)

    middle_frame = tk.Frame(perm1, bg=bgcol)
    middle_frame.pack(side=tk.LEFT)

    right_frame = tk.Frame(perm1, bg=bgcol)
    right_frame.pack(side=tk.LEFT)

    greeting = tk.Label(middle_frame, text="Please provide the following information:", bg=bgcol, height=2, width=100)
    greeting.pack()

    honorific_label = tk.Label(left_frame, text="Honorific:")
    honorific_label.pack()
    honorific = tk.Entry(left_frame)
    honorific.pack(pady=button_padding)

    first_name_label = tk.Label(left_frame, text="First Name:")
    first_name_label.pack()
    first_name = tk.Entry(left_frame)
    first_name.pack(pady=button_padding)

    last_name_label = tk.Label(left_frame, text="Last Name:")
    last_name_label.pack()
    last_name = tk.Entry(left_frame)
    last_name.pack(pady=button_padding)

    title_label = tk.Label(left_frame, text="Title:")
    title_label.pack()
    title = tk.Entry(left_frame)
    title.pack(pady=button_padding)

    address_label = tk.Label(left_frame, text="Address:")
    address_label.pack()
    address = tk.Entry(left_frame)
    address.pack(pady=button_padding)

    agent_name_label = tk.Label(left_frame, text="Agent Full Name:")
    agent_name_label.pack()
    agent_name = tk.Entry(left_frame)
    agent_name.pack(pady=button_padding)

    agent_address_label = tk.Label(left_frame, text="Agent Address:")
    agent_address_label.pack()
    agent_address = tk.Entry(left_frame)
    agent_address.pack(pady=button_padding)

    city_label = tk.Label(left_frame, text="City:")
    city_label.pack()
    city = tk.Entry(left_frame)
    city.pack(pady=button_padding)

    state_label = tk.Label(left_frame, text="State:")
    state_label.pack()
    state = tk.Entry(left_frame)
    state.pack(pady=button_padding)

    zip_code_label = tk.Label(left_frame, text="Zip Code:")
    zip_code_label.pack()
    zip_code = tk.Entry(left_frame)
    zip_code.pack(pady=button_padding)

    project_name_label = tk.Label(middle_frame, text="Project Name:")
    project_name_label.pack()
    project_name = tk.Entry(middle_frame)
    project_name.pack(pady=button_padding)

    project_city_label = tk.Label(middle_frame, text="Project City:")
    project_city_label.pack()
    project_city = tk.Entry(middle_frame)
    project_city.pack(pady=button_padding)

    project_county_label = tk.Label(middle_frame, text="Project County:")
    project_county_label.pack()
    project_county = tk.Entry(middle_frame)
    project_county.pack(pady=button_padding)

    parcel_id_label = tk.Label(middle_frame, text="Parcel ID:")
    parcel_id_label.pack()
    parcel_id = tk.Entry(middle_frame)
    parcel_id.pack(pady=button_padding)

    prefile_date_label = tk.Label(middle_frame, text="Prefile Date:")
    prefile_date_label.pack()
    prefile_date = tk.Entry(middle_frame)
    prefile_date.pack(pady=button_padding)

    notice_type_label = tk.Label(middle_frame, text="Notice Type:")
    notice_type_label.pack()
    notice_type = tk.Entry(middle_frame)
    notice_type.pack(pady=button_padding)

    jpn_date_label = tk.Label(middle_frame, text="JPN Date:")
    jpn_date_label.pack()
    jpn_date = tk.Entry(middle_frame)
    jpn_date.pack(pady=button_padding)

    pnot_date_label = tk.Label(middle_frame, text="PNOT Date:")
    pnot_date_label.pack()
    pnot_date = tk.Entry(middle_frame)
    pnot_date.pack(pady=button_padding)

    exp_date_label = tk.Label(middle_frame, text="Expiration Date:")
    exp_date = tk.Entry(middle_frame)
    exp_date1_label = tk.Label(middle_frame, text="New Expiration Date:")
    exp_date1 = tk.Entry(middle_frame)
    
    if permtype == "Time Extension":
        exp_date_label.pack()
        exp_date.pack(pady=button_padding)
        exp_date1_label.pack()
        exp_date1.pack(pady=button_padding)

    var_code_label = tk.Label(middle_frame, text="Variance from code:")
    var_code = tk.Entry(middle_frame)

    if permtype == "VAR":
        var_code_label.pack(pady=button_padding)
        var_code.pack(pady=button_padding)

    npdes_num_label = tk.Label(middle_frame, text="NPDES Permit:")
    npdes_num = tk.Entry(middle_frame)
    npdes_date_label = tk.Label(middle_frame, text="NPDES Permit Date:")
    npdes_date = tk.Entry(middle_frame)
    parcel_size_label = tk.Label(middle_frame,text="Parcel Size (Ac):")
    parcel_size = tk.Entry(middle_frame)
    
    if permtype == "NRU":
        npdes_num_label.pack()
        npdes_num.pack(pady=button_padding)
        npdes_date_label.pack()
        npdes_date.pack(pady=button_padding)
        parcel_size_label.pack()
        parcel_size.pack(pady=button_padding)

    prefile_label = tk.Label(middle_frame, text="Prefile Date:")
    prefile_date = tk.Entry(middle_frame)

    if permtype == "IP":
        prefile_label.pack(pady=button_padding)
        prefile_date.pack(pady=button_padding)
    

    project_description_label = tk.Label(right_frame, text="Project Description:")
    project_description_label.pack()
    project_description = tk.Text(right_frame, width=50, height=10)
    project_description.pack(pady=button_padding)

    fee_amount_label = tk.Label(right_frame, text="Fee Amount:")
    fee_amount_label.pack()
    fee_amount = tk.Entry(right_frame)
    fee_amount.pack(pady=button_padding)

    fee_received_label = tk.Label(right_frame, text="Fee Received:")
    fee_received_label.pack()
    fee_received = tk.Entry(right_frame)
    fee_received.pack(pady=button_padding)

    acamp_label = tk.Label(right_frame, text="ACAMP Number:")
    acamp_label.pack()
    acamp = tk.Entry(right_frame)
    acamp.pack(pady=button_padding)

    sam_label = tk.Label(right_frame, text="SAM Number:")
    sam_label.pack()
    sam = tk.Entry(right_frame)
    sam.pack(pady=button_padding)

    permitter_list = []
    for i in permitters:
        permitter_list.append(permitters.get(i)[0])
    
    # Create Label
    label1 = tk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack()  

    clicked = tk.StringVar()

    clicked.set( "Choose ADEM Permitter:" )

    drop = tk.OptionMenu( right_frame, clicked, *permitter_list)
    drop.pack(pady=button_padding)

    def callback(*args):
        for i in permitters:
            if clicked.get() == permitters.get(i)[0]:
                print(permitters[i])
                adem_email.delete(0,tk.END)
                adem_employee.delete(0,tk.END)
                adem_employee.insert(0, permitters[i][0])
                adem_email.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    adem_employee_label = tk.Label(right_frame, text="ADEM Employee:")
    adem_employee_label.pack()
    adem_employee = tk.Entry(right_frame)
    adem_employee.pack(pady=button_padding)

    adem_email_label = tk.Label(right_frame, text="ADEM Email:")
    adem_email_label.pack()
    adem_email = tk.Entry(right_frame)
    adem_email.pack(pady=button_padding)

    submit_button = tk.Button(right_frame, text="Submit", command=lambda: get_perm_values(acamp.get(), sam.get(), honorific.get(), first_name.get(), last_name.get(), address.get(), title.get(), agent_name.get(), agent_address.get(), city.get(), state.get(), zip_code.get(), project_name.get(), project_city.get(), project_county.get(), parcel_id.get(), prefile_date.get(), notice_type.get(), jpn_date.get(), pnot_date.get(), project_description.get(1.0, tk.END), fee_amount.get(), fee_received.get(), adem_employee.get(), adem_email.get(),exp_date.get(), exp_date1.get(), npdes_date.get(), npdes_num.get(), parcel_size.get()))
    submit_button.pack(pady=button_padding)


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
    template.save('output/ ' + acamp + ' ' + sam + ' _401WQC_Docs.docx')
    template2.render(context)
    template2.save('output/401Rat_Temp.docx')

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
    templatePerm1.save('output/_CZM_Docs.docx')
    templatePerm2.render(context)
    templatePerm2.save('output/_401WQ_Docs.docx')
    templateRat.render(context)
    templateRat.save('output/_LOP_Rational.docx')
    
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
    templatePerm2.save('output/_VAR_Docs.docx')
    templateRat.render(context)
    templateRat.save('output/_LOP_Rational.docx')
    
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
        'NPDES_Date': NPDES_date,
        'NPDES_Number': npdes_num
    }

    # Render automated report
    templaten.render(context)
    templaten.save('output/' + acamp + ' ' + sam + '_NRU_Docs.docx')
    templatec.render(context)
    templatec.save('output/' + acamp + ' ' + sam + '_CZM_Docs.docx')
    template2.render(context)
    template2.save('output/' + acamp + ' ' + sam + '_NRU_Rational.docx')
    
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
    perm = tk.Toplevel()
    perm.configure(background=bgcol)
    chosen_type = list(document_types.keys())[3]
    subtypes = document_types[chosen_type]
    greeting = tk.Label(perm, text="What type of Permit do you want to generate?", bg=bgcol, height=5, width=100)
    greeting.pack()

    for i, document_type in enumerate(subtypes.keys()):
        document_button = tk.Button(perm, text=f"{i+1}. {document_type}", padx=button_padding, pady=button_padding)
        document_button.pack(pady=button_padding)
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
        'Project_Desc': comments.get(1.0, tk.END),
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
    inspr = tk.Toplevel()
    inspr.bind('<Return>', lambda event: get_inspr_values())
    inspr.configure(background=bgcol, width=500, height=500)
    greeting = tk.Label(inspr, text="Please provide the following information:", bg=bgcol, height=5, width=100)
    greeting.pack()

    # Frame for left column
    left_frame = tk.Frame(inspr, bg=bgcol)
    left_frame.pack(side=tk.LEFT, padx=10)

    # Frame for right column
    right_frame = tk.Frame(inspr, bg=bgcol)
    right_frame.pack(side=tk.LEFT, padx=10)

    # Entry fields with labels in left column
    acamp_label = tk.Label(left_frame, text="ACAMP Number:")
    acamp_label.pack()
    acamp = tk.Entry(left_frame)
    acamp.pack(pady=button_padding)

    sam_label = tk.Label(left_frame, text="SAM Number:")
    sam_label.pack()
    sam = tk.Entry(left_frame)
    sam.pack(pady=button_padding)

    timein_label = tk.Label(left_frame, text="Inspection Time Start:")
    timein_label.pack()
    timein = tk.Entry(left_frame)
    timein.pack(pady=button_padding)

    timeout_label = tk.Label(left_frame, text="Inspection Time End:")
    timeout_label.pack()
    timeout = tk.Entry(left_frame)
    timeout.pack(pady=button_padding)

    firstname_label = tk.Label(left_frame, text="Applicant First Name:")
    firstname_label.pack()
    firstname = tk.Entry(left_frame)
    firstname.pack(pady=button_padding)

    lastname_label = tk.Label(left_frame, text="Applicant Last Name:")
    lastname_label.pack()
    lastname = tk.Entry(left_frame)
    lastname.pack(pady=button_padding)

    complaint_label = tk.Label(left_frame, text="Complaint #:")
    complaint_label.pack()
    complaint = tk.Entry(left_frame)
    complaint.pack(pady=button_padding)

    # Entry fields with labels in right column
    phone_label = tk.Label(left_frame, text="Applicant Phone Number:")
    phone_label.pack()
    phone = tk.Entry(left_frame)
    phone.pack(pady=button_padding)

    address_label = tk.Label(left_frame, text="Applicant Address:")
    address_label.pack()
    address = tk.Entry(left_frame)
    address.pack(pady=button_padding)

    projcoords_label = tk.Label(right_frame, text="Project Coordinates:")
    projcoords_label.pack()
    projcoords = tk.Entry(right_frame)
    projcoords.pack(pady=button_padding)

    projectname_label = tk.Label(right_frame, text="Project Name:")
    projectname_label.pack()
    projectname = tk.Entry(right_frame)
    projectname.pack(pady=button_padding)

    projectcity_label = tk.Label(right_frame, text="Project City:")
    projectcity_label.pack()
    projectcity = tk.Entry(right_frame)
    projectcity.pack(pady=button_padding)

    projectcounty_label = tk.Label(right_frame, text="Project County:")
    projectcounty_label.pack()
    projectcounty = tk.Entry(right_frame)
    projectcounty.pack(pady=button_padding)

    photos_label = tk.Label(right_frame, text="Photos Taken? (Yes/No):")
    photos_label.pack()
    photos = tk.Entry(right_frame)
    photos.pack(pady=button_padding)

    participants_label = tk.Label(right_frame, text="Other Participants (Name, Org):")
    participants_label.pack()
    participants = tk.Entry(right_frame)
    participants.pack(pady=button_padding)

    permitter_list = []
    for i in permitters:
        permitter_list.append(permitters.get(i)[0])
    
    # Create Label
    label1 = tk.Label(right_frame , text = "Choose ADEM Permitter: " )
    label1.pack()  

    clicked = tk.StringVar()

    clicked.set( "Choose ADEM Permitter:" )

    drop = tk.OptionMenu( right_frame, clicked, *permitter_list)
    drop.pack(pady=button_padding)

    def callback(*args):
        for i in permitters:
            if clicked.get() == permitters.get(i)[0]:
                print(permitters[i])
                youremail.delete(0,tk.END)
                yourname.delete(0,tk.END)
                yourname.insert(0, permitters[i][0])
                youremail.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    yourname_label = tk.Label(right_frame, text="Your Name:")
    yourname_label.pack()
    yourname = tk.Entry(right_frame)
    yourname.pack(pady=button_padding)

    youremail_label = tk.Label(right_frame, text="Your Email:")
    youremail_label.pack()
    youremail = tk.Entry(right_frame)
    youremail.pack(pady=button_padding)


    comments_label = tk.Label(inspr, text="Comments/Site Observations:")
    comments_label.pack()
    comments = tk.Text(inspr, width=50, height=10)
    comments.pack(pady=button_padding)

    
    # Button to retrieve input values
    submit_button = tk.Button(inspr, text="Submit", command=get_inspr_values)
    submit_button.pack(pady=button_padding)

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
    feel.destroy()
    print("Files successfully generated in /output/ folder.")

#Fee Sheet Window
def open_feel_window():
    global feel    

    # Fee Letter Window
    feel = tk.Toplevel()
    feel.bind('<Return>', lambda event: get_feel_values())
    feel.configure(background=bgcol, width=500, height=500)
    chosen_type = list(document_types.keys())[0]
    subtypes = document_types[chosen_type]
    greeting = tk.Label(feel, text="Please provide the following information: ", bg=bgcol, height=5, width=100)
    greeting.pack()
    global honorific, firstname, lastname, title, address
    global agentname, agentaddress
    global city, state, zip
    global projectname, projectcity, projectcounty
    global feeamount
    global yourname, youremail, sam, acamp

    # Create input fields
    sam_label = tk.Label(feel, text="SAM Number:")
    sam_label.pack(pady=text_padding)
    sam = tk.Entry(feel)
    sam.pack()

    acamp_label = tk.Label(feel, text="ACAMP Number):")
    acamp_label.pack(pady=text_padding)
    acamp = tk.Entry(feel)
    acamp.pack()

    honorific_label = tk.Label(feel, text="Applicant Honorific (Mr./Ms./Dr./etc):")
    honorific_label.pack(pady=text_padding)
    honorific = tk.Entry(feel)
    honorific.pack()

    firstname_label = tk.Label(feel, text="Applicant First Name:")
    firstname_label.pack(pady=text_padding)
    firstname = tk.Entry(feel)
    firstname.pack()

    lastname_label = tk.Label(feel, text="Applicant Last Name:")
    lastname_label.pack(pady=text_padding)
    lastname = tk.Entry(feel)
    lastname.pack()

    address_label = tk.Label(feel, text="Applicant Address:")
    address_label.pack(pady=text_padding)
    address = tk.Entry(feel)
    address.pack()

    title_label = tk.Label(feel, text="Applicant Title or Company:")
    title_label.pack(pady=text_padding)
    title = tk.Entry(feel)
    title.pack()

    agentname_label = tk.Label(feel, text="Agent Full Name:")
    agentname_label.pack(pady=text_padding)
    agentname = tk.Entry(feel)
    agentname.pack()

    agentaddress_label = tk.Label(feel, text="Agent Address:")
    agentaddress_label.pack(pady=text_padding)
    agentaddress = tk.Entry(feel)
    agentaddress.pack()

    city_label = tk.Label(feel, text="City:")
    city_label.pack(pady=text_padding)
    city = tk.Entry(feel)
    city.pack()

    state_label = tk.Label(feel, text="State:")
    state_label.pack(pady=text_padding)
    state = tk.Entry(feel)
    state.pack()

    zip_label = tk.Label(feel, text="Zip:")
    zip_label.pack(pady=text_padding)
    zip = tk.Entry(feel)
    zip.pack()

    projectname_label = tk.Label(feel, text="Project Name:")
    projectname_label.pack(pady=text_padding)
    projectname = tk.Entry(feel)
    projectname.pack()

    projectcity_label = tk.Label(feel, text="Project City:")
    projectcity_label.pack(pady=text_padding)
    projectcity = tk.Entry(feel)
    projectcity.pack()

    projectcounty_label = tk.Label(feel, text="Project County:")
    projectcounty_label.pack(pady=text_padding)
    projectcounty = tk.Entry(feel)
    projectcounty.pack()

    feeamount_label = tk.Label(feel, text="Fee Amount Due:")
    feeamount_label.pack(pady=text_padding)
    feeamount = tk.Entry(feel)
    feeamount.pack()

    permitter_list = []
    for i in permitters:
        permitter_list.append(permitters.get(i)[0])
    
    # Create Label
    label1 = tk.Label(feel , text = "Choose ADEM Permitter: " )
    label1.pack()  

    clicked = tk.StringVar()

    clicked.set( "Choose ADEM Permitter:" )

    drop = tk.OptionMenu( feel, clicked, *permitter_list)
    drop.pack(pady=button_padding)

    def callback(*args):
        for i in permitters:
            if clicked.get() == permitters.get(i)[0]:
                print(permitters[i])
                youremail.delete(0,tk.END)
                yourname.delete(0,tk.END)
                yourname.insert(0, permitters[i][0])
                youremail.insert(0, permitters[i][1])
        

    clicked.trace("w", callback)

    yourname_label = tk.Label(feel, text="Your Name:")
    yourname_label.pack(pady=text_padding)
    yourname = tk.Entry(feel)
    yourname.pack()

    youremail_label = tk.Label(feel, text="Your Email:")
    youremail_label.pack(pady=text_padding)    
    youremail = tk.Entry(feel)
    youremail.pack()
    
    # Button to retrieve input values
    submit_button = tk.Button(feel, text="Submit", command=get_feel_values)
    submit_button.pack(pady=button_padding)
#END FEE SHEET



#Main Screen Contents
greeting = tk.Label(text="What do you want to generate?", bg=bgcol, height=5, width=100)
greeting.pack()

for i, document_type in enumerate(document_types.keys()):
    document_button = tk.Button(main, text=f"{i+1}. {document_type}", pady=button_padding)
    document_button.pack(pady=button_padding)
    if document_type == "Public Notice":
        document_button.configure(command=open_pnot_window)
    elif document_type == "Permit":
        document_button.configure(command=open_perm_window)
    elif document_type == "Fee Letter":
        document_button.configure(command=open_feel_window)
    elif document_type == "Inspection Report":
        document_button.configure(command=open_inspr_window)


main.mainloop()
