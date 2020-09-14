from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter as tk
from functools import partial
import os
from pathlib import Path
import pandas as pd
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Alignment, Side
import calendar
import logging
from states import *

import re
master = Tk()
master.title("Form Creator")
master.minsize(640,400)

from states import Goa,Karnataka,Chandigarh,Delhi,Maharashtra
Goa=Goa.Goa
Karnataka=Karnataka.Karnataka
Chandigarh=Chandigarh.Chandigarh
Delhi=Delhi.Delhi
Maharashtra=Maharashtra.Maharashtra
#backend code starts here

#from states.global_variables import *
"""
systemdrive = os.getenv('WINDIR')[0:3]
dbfolder = os.path.join(systemdrive,'Forms\DB')
#dbfolder = "D:\Company Projects\Form creator\DB"
State_forms = os.path.join(systemdrive,'Forms\State forms')
#State_forms = "D:\Company Projects\Form creator\State forms"
Statefolder = Path(State_forms)
logfolder = os.path.join(systemdrive,'Forms\logs')
#logfolder = "D:\Company Projects\Form creator\logs"


monthdict= {'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}


log_filename = datetime.datetime.now().strftime(os.path.join(logfolder,'logfile_%d_%m_%Y_%H_%M_%S.log'))
logging.basicConfig(filename=log_filename, level=logging.INFO)
"""

def create_pdf(folderlocation,file_name):
    import win32com.client
    from pywintypes import com_error



    excel_filename = file_name
    pdf_filename = file_name.split('.')[0]+'.pdf'

    


    # Path to original excel file
    WB_PATH=os.path.join(folderlocation,excel_filename)
    # PDF path when saving
    PATH_TO_PDF =os.path.join(folderlocation,pdf_filename)

    logging.info(WB_PATH)
    logging.info(PATH_TO_PDF)


    excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = False

    try:
        logging.info('Start conversion to PDF')

        # Open
        wb = excel.Workbooks.Open(WB_PATH)

        sheetnumbers= len(pd.ExcelFile(WB_PATH).sheet_names)

        # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
        ws_index_list = list(range(1,sheetnumbers+1))
        wb.WorkSheets(ws_index_list).Select()

        # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except com_error as e:
        logging.info('failed.')
    else:
        logging.info('Succeeded.')
    finally:
        wb.Close()
        excel.Quit()


def Tamilnadu(data,contractor_name,contractor_address,filelocation,month,year):
    logging.info('Tamilnadu forms')

def Telangana(data,contractor_name,contractor_address,filelocation,month,year):
    logging.info('Telangana forms')

def Uttar_Pradesh(data,contractor_name,contractor_address,filelocation,month,year):
    logging.info('Uttar Pradesh forms')


Stateslist = ['Chandigarh','Karnataka','Maharashtra','Delhi','Telangana','Uttar Pradesh','Tamilnadu','Goa']

State_Process = {'Goa':Goa,'Chandigarh':Chandigarh,'Karnataka':Karnataka,'Maharashtra':Maharashtra,
                        'Delhi':Delhi,'Telangana':Telangana,'Uttar Pradesh':Uttar_Pradesh,'Tamilnadu':Tamilnadu}

companylist = ['SVR LTD','PRY Wine Ltd','CDE Technology Ltd']



def svrDataProcess(inputfolder,month,year):
    logging.info('scr data process running')

def pryDataProcess(inputfolder,month,year):
    logging.info('pry data process running')

    emp_df_columns = ['Employee Code', 'Employee Name', 'Company Name', 'Grade', 'Branch',
       'Department', 'Designation', 'Division', 'Group', 'Category', 'Unit',
       'Location Code', 'State', 'Date of Birth', 'Date Joined',
       'Date of Confirmation', 'Date Left', 'Title', 'Last Inc. Date',
       'Ticket Number', 'Local Address 1', 'Local Address 2',
       'Local Address 3', 'Local Address 4', 'Local City Name',
       'Local District Name', 'Local PinCode', 'Local State Name',
       'Residence Tel No.', 'Permanent Address 1', 'Permanent Address 2',
       'Permanent Address 3', 'Permanent Address 4', 'UAN Number',
       'Permanent Tel No.', 'Office Tel No.', 'Extension Tel No.',
       'Mobile Tel No.', "Father's Name", 'Gender', 'Age', 'Number of Months',
       'Marital Status', 'PT Number', 'PF Number (Old Version)', 'PF Number',
       'PF Number (WithComPrefix)', 'PAN Number', 'ESIC Number (Old Version)',
       'ESIC Number', 'ESIC Number (CompPrefix)', 'FPF Number', 'PF Flag',
       'ESIC Flag', 'PT Flag', 'Bank A/c Number', 'Bank Name', 'Mode',
       'Account Code', 'E-Mail', 'Remarks', 'PF Remarks', 'ESIC Remarks',
       'ESIC IMP Code', 'ESIC IMP Name', 'Employee Type (For PF)',
       'Freeze Account', 'Freeze Date', 'Freeze Reason', 'Type of House (In)',
       'Comp. adn.', 'Staying In (Metro Type)', 'Children (For CED)',
       'TDS Rate', 'Resignation Date', 'Reason for Leaving', 'Bank A/C No.1',
       'Bank A/C No.2', 'Bank A/C No.3', 'Alt.Email', 'Emp Status',
       'Probation Date', 'Surcharge Flag', 'Gratuity Code',
       'Resign Offer Date', 'Permanent City', 'Permanent District',
       'Permanent Pin Code', 'Permanent State', 'Spouse Name',
       'PF Joining Date', 'PRAN Number', 'Group Joining Date', 'Aadhar Number',
       'Child in Hostel (For CED)', 'Total Exp in Years', 'P', 'L']

    sal_df_columns = ['PAY_DATE', 'EMPCODE', 'EMPNAME', 'DOJ', 'GENDER', 'DOB', 'DOL',
       'OT HOURS', 'STDDAYS', 'LOP DAYS', 'WRKDAYS', 'BASIC+DA',
       'BASIC+DA ARREAR', 'HOUSE RENT ALLOWANCE', 'HRA ARREAR',
       'MEDICAL ALLOWANCE', 'MEDICAL ALLOWANCE ARREARS',
       'CHILD EDUCATION ALLOWANCE', 'CHILD EDUCATION ALL ARREARS',
       'HELPER ALLOWANCE', 'HELPER ALLOWANCE ARREARS', 'OTHER ALLOWANCE',
       'OTHER ALLOWANCE  ARREARS', 'OTHER EARNINGS NON TAXABLE',
       'OTHER EARNINGS', 'SALES COMMISSION', 'PROFITABILITY & SALES',
       'MBO PAYMENT', 'STAT BONUS PAYMENT', 'VEHICLE SUBSIDY', 'EXGRATIA',
       'NOTICE PAY BUYOUTS', 'OVERTIME', 'MARRIAGE GIFT', 'REFERRAL BONUS',
       'FESTIVAL ADVANCE PAY', 'LONG TERM SERVICE AWARD', 'LEAVE ENCASHMENT',
       'COMPANY SALES PAYMENT', 'PROFITABILITY PAYMENT',
       'Ticket Restaurant Meal Card', 'UNIFORM CLEARANCE', 'LTA REIMBURSEMENT',
       'GRAPE QUALITY PAY', 'SHORT PAY EARNING',
       'HOSPITALITY PERFORMANCE BONUS', 'OVERSEAS SPECIAL ALLOWANCE',
       'GROSS_EARN', 'Professional Tax', 'ESI']

    file_list = os.listdir(inputfolder)
    logging.info('input folder is '+str(inputfolder))
    for f in file_list:
        if f[0:6].upper()=='MASTER':
            masterfilename = f
            logging.info('masterfilename is :'+f)
        if f[0:6].upper()=='SALARY':
            salaryfilename = f
            logging.info('salaryfilename is :'+f)
        if f[0:10].upper()=='ATTENDANCE':
            attendancefilename = f
            logging.info('attendancefilename is :'+f)
        if f[0:5].upper()=='LEAVE':
            leavefilename = f
            logging.info('leavefilename is :'+f)
        if f[0:14].upper()=='LEFT EMPLOYEES':
            leftempfilename = f
            logging.info('leftempfilename is :'+f)
        if f[0:9].upper()=='CDE UNITS':
            unitfilename = f
            logging.info('unitfilename is :'+f)
    
    logging.info('file names set')
    
    if 'masterfilename' in locals():
        masterfile = os.path.join(inputfolder,masterfilename)
        employee_data = pd.read_excel(masterfile)
        employee_data.dropna(how='all', inplace=True)
        employee_data.reset_index(drop=True, inplace=True)
        employee_data.columns = employee_data.iloc[0]
        employee_data.drop(0, inplace=True)
        logging.info('employee data loaded')
    else:
        employee_data = pd.DataFrame(columns = emp_df_columns)
        logging.error('employee data not available setting empty dataset')
    if 'salaryfilename' in locals():
        salaryfile = os.path.join(inputfolder,salaryfilename)
        salary_data = pd.read_excel(salaryfile)
    else:
        salary_data = pd.DataFrame(columns= sal_df_columns)
        logging.error('salary data not available setting empty dataset')
    if 'attendancefilename' in locals():
        attendancefile = os.path.join(inputfolder,attendancefilename)
        attendance_data_1 = pd.read_excel(attendancefile, sheet='HO')
        attendance_data_2 = pd.read_excel(attendancefile, sheet='ROI')
        attendance_data = attendance_data_1.concat(attendance_data_2)
    else:
        attendance_data = pd.DataFrame(columns= sal_df_columns)
        logging.error('attendance data not available setting empty dataset')

    
    pry_data = salary_data.merge(attendance_data, how='left', left_on='EMPCODE', right_on='Empl Code')




def cdeDataProcess(inputfolder,month,year):
    global nomatch
    nomatch=''
    logging.info('cde data process running')

    emp_df_columns = ['Employee Code', 'Employee Name', 'Company Name', 'Grade', 'Branch',
       'Department', 'Designation', 'Division', 'Group', 'Category', 'Unit',
       'Location Code', 'State', 'Date of Birth', 'Date Joined',
       'Date of Confirmation', 'Date Left', 'Title', 'Last Inc. Date',
       'Ticket Number', 'Local Address 1', 'Local Address 2',
       'Local Address 3', 'Local Address 4', 'Local City Name',
       'Local District Name', 'Local PinCode', 'Local State Name',
       'Residence Tel No.', 'Permanent Address 1', 'Permanent Address 2',
       'Permanent Address 3', 'Permanent Address 4', 'UAN Number',
       'Permanent Tel No.', 'Office Tel No.', 'Extension Tel No.',
       'Mobile Tel No.', "Father's Name", 'Gender', 'Age', 'Number of Months',
       'Marital Status', 'PT Number', 'PF Number (Old Version)', 'PF Number',
       'PF Number (WithComPrefix)', 'PAN Number', 'ESIC Number (Old Version)',
       'ESIC Number', 'ESIC Number (CompPrefix)', 'FPF Number', 'PF Flag',
       'ESIC Flag', 'PT Flag', 'Bank A/c Number', 'Bank Name', 'Mode',
       'Account Code', 'E-Mail', 'Remarks', 'PF Remarks', 'ESIC Remarks',
       'ESIC IMP Code', 'ESIC IMP Name', 'Employee Type (For PF)',
       'Freeze Account', 'Freeze Date', 'Freeze Reason', 'Type of House (In)',
       'Comp. adn.', 'Staying In (Metro Type)', 'Children (For CED)',
       'TDS Rate', 'Resignation Date', 'Reason for Leaving', 'Bank A/C No.1',
       'Bank A/C No.2', 'Bank A/C No.3', 'Alt.Email', 'Emp Status',
       'Probation Date', 'Surcharge Flag', 'Gratuity Code',
       'Resign Offer Date', 'Permanent City', 'Permanent District',
       'Permanent Pin Code', 'Permanent State', 'Spouse Name',
       'PF Joining Date', 'PRAN Number', 'Group Joining Date', 'Aadhar Number',
       'Child in Hostel (For CED)', 'Total Exp in Years', 'P', 'L']

    salary_df_columns = ['Sr', 'DivisionName', 'Sal Status', 'Emp Code', 'Emp Name', 'DesigName',
       'Date Joined', 'UnitName', 'Branch', 'Days Paid', 'Earned Basic', 'HRA',
       'Conveyance', 'Medical Allowance', 'Telephone Reimb',
       'Tel and Int Reimb', 'Bonus', 'Other Allowance', 'Fuel Reimb',
       'Prof Dev Reimb', 'Corp Attire Reimb', 'Meal Allowance',
       'Special Allowance', 'Personal Allowance', 'CCA', 'Other Reimb',
       'Arrears', 'Other Earning', 'Variable Pay', 'Leave Encashment',
       'Stipend', 'Consultancy Fees', 'Total Earning', 'Insurance', 'CSR',
       'PF', 'ESIC', 'P.Tax', 'LWF EE', 'Salary Advance', 'Loan Deduction',
       'Loan Interest', 'Other Deduction', 'TDS', 'Total Deductions',
       'Net Paid', 'BankName', 'Bank A/c Number', 'Account Code', 'Remarks',
       'PF Number (Old)', 'UAN Number', 'ESIC Number', 'Personal A/c Number',
       'E-Mail', 'Mobile No.', 'FIXED MONTHLY GROSS', 'CHECK CTC Gross']

    atten_df_columns = ['Emp Code', 'Employee Name', 'Branch', 'Designation', 'Sat\r\n01/02',
       'Sun\r\n02/02', 'Mon\r\n03/02', 'Tue\r\n04/02', 'Wed\r\n05/02',
       'Thu\r\n06/02', 'Fri\r\n07/02', 'Sat\r\n08/02', 'Sun\r\n09/02',
       'Mon\r\n10/02', 'Tue\r\n11/02', 'Wed\r\n12/02', 'Thu\r\n13/02',
       'Fri\r\n14/02', 'Sat\r\n15/02', 'Sun\r\n16/02', 'Mon\r\n17/02',
       'Tue\r\n18/02', 'Wed\r\n19/02', 'Thu\r\n20/02', 'Fri\r\n21/02',
       'Sat\r\n22/02', 'Sun\r\n23/02', 'Mon\r\n24/02', 'Tue\r\n25/02',
       'Wed\r\n26/02', 'Thu\r\n27/02', 'Fri\r\n28/02', 'Sat\r\n29/02',
       'Total\r\nDP', 'Total\r\nABS', 'Total\r\nLWP', 'Total\r\nCL',
       'Total\r\nSL', 'Total\r\nPL', 'Total\r\nL1', 'Total\r\nL2',
       'Total\r\nL3', 'Total\r\nL4', 'Total\r\nL5', 'Total\r\nCO-',
       'Total\r\nCO+', 'Total\r\nOL', 'Total\r\nWO', 'Total\r\nPH',
       'Total\r\nEO', 'Total\r\nWOP', 'Total\r\nPHP', 'Total\r\nOT Hrs',
       'Total\r\nLT Hrs']

    leave_df_columns = ['Emp. Code', 'Emp. Name', 'Leave Type', 'Opening', 'Monthly Increment',
       'Used', 'Closing', 'Leave Accrued', 'Encash']

    leftemp_df_columns = ['Employee Name', 'Employee Code', 'Date Joined', 'Date Left',
       'UAN Number']

    unit_df_columns = ['Unit', 'Location_code','Location', 'Address', 'Registration_no', 'PE_or_contract',
       'State_or_Central', 'start_time', 'end_time', 'rest_interval','Contractor_name','Contractor_Address']

    logging.info('column variables set')

    

    
    file_list = os.listdir(inputfolder)
    logging.info('input folder is '+str(inputfolder))
    for f in file_list:
        if f[0:6].upper()=='MASTER':
            masterfilename = f
            logging.info('masterfilename is :'+f)
        if f[0:6].upper()=='SALARY':
            salaryfilename = f
            logging.info('salaryfilename is :'+f)
        if f[0:10].upper()=='ATTENDANCE':
            attendancefilename = f
            logging.info('attendancefilename is :'+f)
        if f[0:5].upper()=='LEAVE':
            leavefilename = f
            logging.info('leavefilename is :'+f)
        if f[0:14].upper()=='LEFT EMPLOYEES':
            leftempfilename = f
            logging.info('leftempfilename is :'+f)
        if f[0:9].upper()=='CDE UNITS':
            unitfilename = f
            logging.info('unitfilename is :'+f)
    
    logging.info('file names set')
    
    if 'masterfilename' in locals():
        masterfile = os.path.join(inputfolder,masterfilename)
        employee_data = pd.read_excel(masterfile)
        employee_data.dropna(how='all', inplace=True)
        employee_data.reset_index(drop=True, inplace=True)
        logging.info('employee data loaded')
    else:
        employee_data = pd.DataFrame(columns = emp_df_columns)
        logging.error('employee data not available setting empty dataset')
    if 'salaryfilename' in locals():
        salaryfile = os.path.join(inputfolder,salaryfilename)
        salary_data = pd.read_excel(salaryfile)
        salary_data.dropna(how='all', inplace=True)
        salary_data.reset_index(drop=True, inplace=True)
        logging.info('salary data loaded')
    else:
        salary_data = pd.DataFrame(columns = salary_df_columns)
        logging.info('salary data not available setting empty dataset')
    if 'attendancefilename' in locals():
        attendancefile = os.path.join(inputfolder,attendancefilename)
        attendance_data = pd.read_excel(attendancefile)
        attendance_data.dropna(how='all', inplace=True)
        attendance_data.reset_index(drop=True, inplace=True)
        logging.info('attendance data loaded')
    else:
        attendance_data = pd.DataFrame(columns = atten_df_columns)
        logging.info('attendance data not available setting empty dataset')
    if 'leavefilename' in locals():
        leavefile = os.path.join(inputfolder,leavefilename)
        leave_data = pd.read_excel(leavefile)
        leave_data.dropna(how='all', inplace=True)
        leave_data.reset_index(drop=True, inplace=True)
        logging.info('leave data loaded')
    else:
        leave_data = pd.DataFrame(columns = leave_df_columns)
        logging.info('leave data not available setting empty dataset')
    if 'leftempfilename' in locals():
        leftempfile = os.path.join(inputfolder,leftempfilename)
        leftemp_data = pd.read_excel(leftempfile)
        leftemp_data.dropna(how='all', inplace=True)
        leftemp_data.reset_index(drop=True, inplace=True)
        logging.info('left employees data loaded')
    else:
        leftemp_data = pd.DataFrame(columns = leftemp_df_columns)
        logging.info('left employees data not available setting empty dataset')
    if 'unitfilename' in locals():
        unitfile = os.path.join(inputfolder,unitfilename)
        unit_data = pd.read_excel(unitfile)
        unit_data.dropna(how='all', inplace=True)
        unit_data.reset_index(drop=True, inplace=True)
        logging.info('unit data loaded')
    else:
        unit_data = pd.DataFrame(columns = unit_df_columns)
        logging.info('unit data not available setting empty dataset')

    employee_data.drop(columns='Date Left', inplace=True)

    employee_data['Location Code'] = employee_data['Location Code'].astype(int)
    employee_data['Employee Code'] = employee_data['Employee Code'].astype(str)

    unit_data['Location Code'] = unit_data['Location Code'].astype(int)

    salary_data.drop(columns=list(employee_data.columns.intersection(salary_data.columns)), inplace=True)

    salary_data['Emp Code'] = salary_data['Emp Code'].astype(str)

    attendance_data.drop(columns=['Employee Name', 'Branch', 'Designation'], inplace=True)
    attendance_data['Emp Code'] = attendance_data['Emp Code'].astype(str)

    leave_data['Emp. Code'] = leave_data['Emp. Code'].astype(str)

    leftemp_data.drop(columns=['Employee Name', 'Date Joined', 'UAN Number'],inplace=True)

    leftemp_data['Employee Code'] = leftemp_data['Employee Code'].astype(str)


    
    
    CDE_Data = employee_data.merge(unit_data,how='left',on='Location Code').merge(
        salary_data,how='left',left_on='Employee Code',right_on='Emp Code').merge(
            attendance_data,how='left',left_on='Employee Code', right_on='Emp Code').merge(
                leave_data, how='left', left_on='Employee Code', right_on='Emp. Code').merge(
                    leftemp_data, how='left', on='Employee Code')
    
    logging.info('merged all data sets')

    rename_list=[]
    renamed=[]
    drop_list=[]
    for x in list(CDE_Data.columns):
        if x[-2:]=='_x':
            rename_list.append(x)
            renamed.append(x[0:-2])
        if x[-2:]=='_y':
            drop_list.append(x)
    
    rename_dict = dict(zip(rename_list,renamed))

    CDE_Data.rename(columns=rename_dict, inplace=True)

    logging.info('columns renamed correctly')

    CDE_Data.drop(columns=drop_list, inplace=True)

    logging.info('dropped duplicate columns')

    monthyear = month+' '+str(year)
    if monthyear.upper() in masterfilename.upper():
        logging.info('month year matches with data')

        #for all state employees(PE+contractor)

        statedata = CDE_Data[CDE_Data['State_or_Central']=='State'].copy()
        statedata.State='Maharashtra'
        CDE_States = list(statedata['State'].unique())
        print("--------------------")
        print(CDE_States)
        for state in CDE_States:
            unit_with_location = list((statedata[statedata.State==state]['Unit']+','+statedata[statedata.State==state]['Location']).unique())
            print(unit_with_location)
            for UL in unit_with_location:
                UL=re.sub(r"(\.*\s+)$","",UL)
                inputdata = statedata[(statedata['State']==state) & (statedata['Unit']==UL.split(',')[0]) & (statedata['Location']==UL.split(',')[1])].copy()
                
                inputdata['Contractor_name'] = inputdata['Contractor_name'].fillna(value='')
                inputdata['Contractor_Address'] = inputdata['Contractor_Address'].fillna(value='')
                inpath = os.path.join(inputfolder,'Registers','States',state,UL)
                if os.path.exists(inpath):
                    logging.info('running state process')
                    contractor_name= inputdata['Contractor_name'].unique()[0]
                    contractor_address= inputdata['Contractor_Address'].unique()[0]
                    State_Process[state](data=inputdata,contractor_name=contractor_name,contractor_address=contractor_address,filelocation=inpath,month=month,year=year)
                else:
                    logging.info('making directory')
                    os.makedirs(inpath)
                    logging.info('directory created')
                    contractor_name= inputdata['Contractor_name'].unique()[0]
                    contractor_address= inputdata['Contractor_Address'].unique()[0]
                    State_Process[state](data=inputdata,contractor_name=contractor_name,contractor_address=contractor_address,filelocation=inpath,month=month,year=year)
                    #h
        #for contractors form
        contractdata = CDE_Data[(CDE_Data['State_or_Central']=='State') & (CDE_Data['PE_or_contract']=='Contract')].copy()
        contractor_units = list((contractdata['Unit']+','+contractdata['Location']).unique())
        for UL in contractor_units:
            inputdata = contractdata[(contractdata['Unit']==UL.split(',')[0]) & (contractdata['Location']==UL.split(',')[1])]
            
            inpath = os.path.join(inputfolder,'Registers','Contractors',UL)
            if os.path.exists(inpath):
                logging.info('running contractor process')
            else:
                logging.info('making directory')
                os.makedirs(inpath)
                logging.info('directory created')

        #for central form
        centraldata = CDE_Data[CDE_Data['State_or_Central']=='Central'].copy()
        central_units = list((centraldata['Unit']+','+centraldata['Location']).unique())
        for UL in central_units:
            inputdata = centraldata[(centraldata['Unit']==UL.split(',')[0]) & (centraldata['Location']==UL.split(',')[1])]
            
            inpath = os.path.join(inputfolder,'Registers','Central',UL)
            if os.path.exists(inpath):
                logging.info('running contractor process')
            else:
                logging.info('making directory')
                os.makedirs(inpath)
                logging.info('directory created')
    
    else:
        nomatch = "Date you mentioned doesn't match with Input data"
        logging.error(nomatch)

    

DataProcess = {'SVR LTD':svrDataProcess,'PRY Wine Ltd':pryDataProcess,'CDE Technology Ltd':cdeDataProcess}


def CompanyDataProcessing(company,inputfolder,month,year):
    inputfolder = Path(inputfolder)
    yr = int(year)
    DataProcess[company](inputfolder,month,yr)

#backend code ends here

companies = ['SVR LTD','PRY Wine Ltd','CDE Technology Ltd']

Months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']

Years = ['2017','2018','2019','2020']

companyname = tk.StringVar()

month = tk.StringVar()
year = tk.StringVar()


folderLabel = ttk.LabelFrame(master, text="Select the Company")
folderLabel.grid(column=0,row=1,padx=20,pady=20)

companynameLabel = Label(folderLabel, text="Company Name")
companynameLabel.grid(column=1, row=0, padx=20,pady=20)

comapnynameEntry = ttk.Combobox(folderLabel,values=companies,textvariable=companyname)
comapnynameEntry.grid(column=2, row=0, padx=20,pady=20)

MonthLabel = Label(folderLabel, text="Month and Year")
MonthLabel.grid(column=1, row=2, padx=20,pady=20)

MonthEntry = ttk.Combobox(folderLabel,values=Months,textvariable=month)
MonthEntry.grid(column=2, row=2, padx=20,pady=20)

YearEntry = ttk.Combobox(folderLabel,values=Years,textvariable=year)
YearEntry.grid(column=3, row=2, padx=20,pady=20)

def disfo():
    foldername = filedialog.askdirectory()
    logging.info(foldername)
    logging.info(type(foldername))
    foldernamelabel.configure(text=foldername)


button = ttk.Button(folderLabel, text = "Select Company Folder", command=disfo)
button.grid(column=1, row=1, columnspan=2,padx=20, pady=20)

foldernamelabel = Label(folderLabel, text="")
foldernamelabel.grid(column=1, row=3, columnspan=2,padx=20,pady=20)




def generateforms(comp,mn,yr):
    company=comp.get()

    month = mn.get()
    year = yr.get()


    getfolder = foldernamelabel.cget("text")



    logging.info(type(company))
    logging.info(company)

    logging.info(type(getfolder))
    logging.info(getfolder)

    if (company =="" and getfolder =="" and (month =="" or year =="")):
        report.configure(text="Please select month year, company folder and company name")
    elif (company=="" and getfolder =="" and not(month =="" or year =="")):
        report.configure(text="Please select company folder and company name")
    elif (company=="" and getfolder !="" and not(month =="" or year =="")):
        report.configure(text="Please select company name")
    elif (company!="" and getfolder =="" and not(month =="" or year =="")):
        report.configure(text="Please select company folder")
    elif (company =="" and getfolder !="" and (month =="" or year=="")):
        report.configure(text="Please select month year and company name")
    elif (company!="" and getfolder=="" and (month =="" or year=="")):
        report.configure(text="Please select month year and company folder")
    elif (company!="" and getfolder!="" and (month =="" or year=="")):
        report.configure(text="Please select month year")
    else:
        logging.info(company, getfolder,  month,  year)
        report.configure(text="Processing")
        try:
            CompanyDataProcessing(company,getfolder,month,year)
        except Exception as e:
            logging.info('Failed')
            report.configure('Failed')
        else:
            if nomatch=='':
                logging.info('Completed Form Creation')
                report.configure(text='Completed Form Creation')
            else:
                logging.info(nomatch)
                report.configure(text=nomatch)
        finally:
            logging.info('done')
        
def convert_forms_to_pdf():

    getfolder = foldernamelabel.cget("text")

    if getfolder=="":
        report.configure(text="Please select company folder")
    else:
        registerfolder = os.path.join(Path(getfolder),'Registers')
        if os.path.exists(registerfolder):
            for root, dirs, files in os.walk(registerfolder):
                for fileis in files:
                    if fileis.endswith(".xlsx"):
                        try:
                            create_pdf(root,fileis)
                        except Exception as e:
                            logging.info('Failed pdf Conversion')
                            report.configure(text="Failed")
                        else:
                            logging.info('Completed pdf Conversion')
                            report.configure(text="Completed")
                        finally:
                            logging.info('done')
        else:
            report.configure(text="Registers not available")
                        



generateforms = partial(generateforms,companyname,month,year)

button = ttk.Button(master, text = "Generate Forms", command=generateforms)
button.grid(column=1, row=1, columnspan=2,padx=20, pady=20)

Detailbox = ttk.LabelFrame(master, text="")
Detailbox.grid(column=0,row=2,padx=20,pady=20)

report = Label(Detailbox, text="                                                            ")
report.grid(column=0, row=0, padx=20,pady=20)



button2 = ttk.Button(master, text = "Convert forms to PDF", command=convert_forms_to_pdf)
button2.grid(column=0, row=3, columnspan=2,padx=20, pady=20)


mainloop()


