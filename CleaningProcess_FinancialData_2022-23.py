# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 16:06:44 2022

@author: londoncm
"""
import pandas as pd
import os
from tkinter import *
import tkinter.messagebox


# Adding the path where is the data located

os.chdir(r'Z:\Python\Test\Financial_Data') 

os.listdir()


# --------------Import Main Data (Excel)-------------------------

df1 = pd.read_excel('FinancialData2021_23_Python_V2.xlsx',
                    sheet_name='Data_SAP_Apr_2022-23', dtype=object) # change the sheet name every month
df1.head()



# Null reeplacement Main Data

df1['Posting Date'].fillna('6-30-2022', inplace=True)
df1['Posting Date'] = pd.to_datetime(df1['Posting Date'],
        format='%m-%d-%Y')

df1['Cost Center'].fillna('0', inplace=True)

df1['Profit Center'].fillna('NA', inplace=True)

df1['WBS'].fillna('0', inplace=True)

df1['Purchasing document'].fillna('NA', inplace=True)

df1['Vendor'].fillna('NA', inplace=True)

df1['Cost Center'] = df1['Cost Center'].values.astype('str')
df1['Account Number'] = df1['Account Number'].values.astype('str')
df1['WBS'] = df1['WBS'].values.astype('str')


# Remove colummns we do not use in Main Data

df1 = df1.drop([
    'Document Type',
    'Ref. Document',
    'Document Date',
    'Curr. Key of PrCtr LC',
    'In transaction currency',
    'Curr. Key Trans. Curr.',
    'Text',
    #'Purchasing document',
    #'Vendor',
    'User Name',
    ], axis=1)


# Create new columns to merge main data with mapping1.  

def center_code(row):

    if row['Profit Center'].startswith('COR') and row['WBS'] == '0':

        row['WBS_Masking'] = (row['Profit Center'])[0:3]
    elif row['Profit Center'].startswith('VPR') and row['WBS'] == '0':

        row['WBS_Masking'] = (row['Profit Center'])[0:3]
    else:

        row['WBS_Masking'] = (row['WBS'])[0:3]

    return row

# Adding the new column (WBS Masking) in the financial data

df1 = df1.apply(center_code, axis=1) 



df1['WBS_Masking'].fillna('0', inplace=True)


def code(row):
    
    if row['Profit Center'].startswith('VPR') and row['WBS_Masking'] \
        == 'GEN':
        row['Code'] = row['WBS']
    elif row['WBS_Masking'] == 'GEN' and row['Profit Center'
            ].startswith('COR'):
        row['Code'] = row['WBS']
    elif row['WBS_Masking'] == '0' or row['WBS_Masking'] == 'AUX' \
        or row['WBS_Masking'] == 'GEN' or row['WBS_Masking'] == 'RBA':
        row['Code'] = (row['Profit Center'])[0:7]
    elif row['WBS_Masking'] == 'REI':
        row['Code'] = row['WBS']
    elif row['WBS_Masking'] == 'COR' or row['WBS_Masking'] == 'VPR':
        row['Code'] = row['Cost Center']
    else:
        row['Code'] = row['WBS_Masking']
    return row


df1 = df1.apply(code, axis=1)

df1['Code'] = df1['Code'].values.astype('str')

# ---------Import Mapping1 data. Mapping 1 contain departments and sectors.---------------------------
df2 = pd.read_excel('FinancialData2021_23_Python_V2.xlsx',
                    sheet_name='Mapping1', dtype=object)

df2.head()

# Null reeplacement Mapping1 Data


df2.fillna('0', inplace=True)
df2['Code'] = df2['Code'].values.astype('str')
df2 = df2.drop_duplicates()

# Cheking that all Codes in Main Data in Mapping1 Data

df1['Check'] = df1.Code.isin(df2.Code)
print(df1['Check'].value_counts())
df1['Check'] = df1['Check'].values.astype('str')

df_falses = df1[(df1['Check']=='False')]
df_falses= df_falses['Code'].drop_duplicates()
print(df_falses.to_markdown())



if 'False' in df1['Check'].tolist():


    window = Tk()
    window.geometry('400x335')
    window.resizable(0, 0)
    window.title('Enter Missing Data')
    
    
    subframe = Frame(window, bg='#24a7b1')
    subframe.grid(column=0, row=0, sticky='nsew')
     
    subframe2 = Frame(window, bg='#24a7b1')
    subframe2.grid(column=0, row=1, sticky='nsew')
    
    
    ### --- code Input
    
    label00 = Label(subframe, text = 'Code', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=0,
             sticky = 'ew', padx = 5, pady = 5)
    code_entry = Entry(subframe,width=20, font=('Arial', 12))
    code_entry.grid(row = 0, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- sector list Input
    label20 = Label(subframe, text = 'Sector', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label20.grid(row = 2, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    sector_list_series = ['Select Sector','Research Center','Research Funding', 'Research Office', \
                    'Core Lab & RI', 'President Initiative']
    
    
    sector_list_var = StringVar(window)
    sector_list_var.set(sector_list_series[0])
    sector_list_entry = OptionMenu(subframe, sector_list_var, *sector_list_series)
    sector_list_entry.config(width=40)
    sector_list_entry.grid(row = 2, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- department list input
    label30 = Label(subframe, text = 'Department', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label30.grid(row = 3, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    department_List_series = [
        'Select Department',
        'RC Advance Membranes',
        'RC ANPERC',
        'RC Catalysis',
        'RC Clean Combustion',
        'RC Computational Bioscience',
        'RC Desert Agriculture',
        'RC Extreme Computing',
        'RC RC3',
        'RC Red Sea',
        'RC Solar',
        'RC VCC',
        'RC Water Desalination',
        'GCR',
        'URF',
        'External Research',
        'Baseline',
        'Discretionary',
        'CCF',
        'Center Partnership',
        'Res Partnership',
        'Research Capital',
        'KRO',
        'Office of A-VPR & COO',
        'Office Of the VPR',
        'Research Funding and Services (RFS)',
        'Research Support and Valorization',
        'Research Translation and Partnerships',
        'VPR Projects',
        'Analytical Chemistry Core Lab',
        'Animal Resources Facility Core Lab',
        'Bioscience Core Lab',
        'Central Workshop Core Lab',
        'CLRI Projects',
        'CLRI Research Park',
        'Coastal & Marine Resources Core Lab',
        'Core Labs Operation & Support',
        'Core Res Park / Grants Projects',
        'Greenhouse Core Lab',
        'Imaging & Characterization Core Lab',
        'LEM',
        'Nanofabrication Core Lab',
        'New Energy Oasis Facility Core Lab',
        'Radiation Labelling Facility Core Lab',
        'Supercomputing Core Lab',
        'Visualization Core Lab',
        'Artifical Interlligence Initiative',
        'Central Node - Researchers',
        'Circular Carbin',
        'Climate and Livability Initiative',
        'Core Lab-Cloud Bursting',
        'G20/S20 Support Office',
        'Impact Acceleration',
        'Metagenomic RS Impact Focused - Carlos/PF',
        'Near Term Grant Challenge',
        'NEOM CoE',
        'Smart Health - Operation',
        'Smart Health Initiative',
        'Translational Grant',
        'VPR Strategic Engagements',
        'Reefscape Restoration Initiative - Shushah funded'
        ]
    
    department_List_var = StringVar(window)
    department_List_var.set(department_List_series[0])
    department_list_entry = OptionMenu(subframe, department_List_var, *department_List_series)
    department_list_entry.config(width=40)
    department_list_entry.grid(row = 3, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    (code1, department1, sector1) = ([], [], [])
    
    def add_data():
        global code1, department1, sector1
    
        code1.append(code_entry.get())
        department1.append(department_List_var.get())
        sector1.append(sector_list_var.get())
    
        code_entry.delete(0, END)
        sector_list_var.set(sector_list_series[0])
        department_List_var.set(department_List_series[0])
    
    def save_data():
        global code1, department1, sector1
    
        data = {'Code': code1, 'Department': department1, 'Sector': sector1}
        name_excel = 'new_mapping1' + '.xlsx'
        df = pd.DataFrame(data, columns=['Code', 'Department', 'Sector'])
        df.to_excel(name_excel, index=False)
        tkinter.messagebox.showinfo('Finish Process Mapping1','The data has been saved. Please close the window')
        
    
    Add = Button(
        subframe,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Add',
        bg='#f08823',
        bd=5,
        command=add_data,
        )
    Add.grid(columnspan=2, row=5, pady=20, padx=10)
    
    file = Label(
        subframe2,
        text='Save data after adding all missing data',
        width=40,
        bg='#24a7b1',
        font=('Arial', 11),
        fg='black',
        )
    file.grid(column=0, row=0, pady=20, padx=10)
    
    
    save = Button(
        subframe2,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Save',
        bg='#bdcf30',
        bd=5,
        command=save_data,
        )
    save.grid(column=0, row=2, pady=20, padx=10)
        
    
    window.mainloop()
    
    # Concat new data in Mapping1


    df2_new_mapping = pd.read_excel('new_mapping1.xlsx', dtype=object)

    df2 = pd.concat([df2, df2_new_mapping])

    df1['Check'] = df1.Code.isin(df2.Code)
    print(df1['Check'].value_counts())

    # Merge Main data with Mapping 1

    df1 = pd.merge(df1, df2, on='Code', how='left')
    
else:
    df1 = pd.merge(df1, df2, on='Code', how='left')
    


#---------------Import Mapping2 data---------------------------

df3 = pd.read_excel('FinancialData2021_23_Python_V2.xlsx',
                    sheet_name='Mapping2', dtype=object)
df3.head()

df3['WBS / CC Description'].fillna('NA', inplace=True)
df3['Status'].fillna('NA', inplace=True)
df3['Fund Program'].fillna('NA', inplace=True)
df3['WBS Owner Name'].fillna('NA', inplace=True)
df3['WBS Owner-KAUST ID'].fillna('NA', inplace=True)
df3['WBS Owner Status'].fillna('NA', inplace=True)
df3['Based on WBS Description Faculty Name'].fillna('NA', inplace=True)

df3.fillna('0', inplace=True)
df3 = df3.drop_duplicates()

df3['WBS / CC'] = df3['WBS / CC'].values.astype('str')

df3['Fund Program'] = df3['Fund Program'].values.astype('str')
df3['Profit Center'] = df3['Profit Center'].values.astype('str')
df3['WBS Owner Name'] = df3['WBS Owner Name'].values.astype('str')
df3['Based on WBS Description Faculty Name'] = df3['Based on WBS Description Faculty Name'].values.astype('str')

def mapping2_code(row):
    
    if row['WBS']== '0':
       row['WBS / CC'] = row['Cost Center']
    else:
       row['WBS / CC'] = row['WBS']
    return row   

df1 = df1.apply(mapping2_code, axis=1)

df1['WBS / CC'] = df1['WBS / CC'].values.astype('str')


df1['Check_df3'] = df1['WBS / CC'].isin(df3['WBS / CC'])
print(df1['Check_df3'].value_counts())
df1['Check_df3'] = df1['Check_df3'].values.astype('str')


df_falses2 = df1[(df1['Check_df3']=='False')]
df_falses2= df_falses2['WBS / CC'].drop_duplicates()
print(df_falses2.to_markdown())

# ----------------Enter missing data in Mapping2-------------------------------


if 'False' in df1['Check_df3'].tolist():
    
    window = Tk()
    window.geometry('625x550')
    window.resizable(0, 0)
    window.title('Enter Missing Data')
    
    
    subframe = Frame(window, bg='#24a7b1')
    subframe.grid(column=0, row=0, sticky='nsew')
     
    subframe2 = Frame(window, bg='#24a7b1')
    subframe2.grid(column=0, row=1, sticky='nsew')
    
    
    ### --- WBS / CC Input
    
    label00 = Label(subframe, text = 'WBS / CC', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=0,
             sticky = 'ew', padx = 5, pady = 5)
    WBS_entry = Entry(subframe,width=20, font=('Arial', 12))
    WBS_entry.grid(row = 0, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- WBS / CC Description Input
    
    label01 = Label(subframe, text = 'WBS / CC Description', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=1,
             sticky = 'ew', padx = 5, pady = 5)
    WBS_CC_entry = Entry(subframe,width=20, font=('Arial', 12))
    WBS_CC_entry.grid(row = 1, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- Fund Program list Input
    
    label02 = Label(subframe, text = 'Fund Program', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label02.grid(row = 2, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    fund_program_list_series = df3['Fund Program'].tolist()
    fund_program_list_series = set(fund_program_list_series)
    fund_program_list_series = sorted(fund_program_list_series)
    fund_program_list_series.insert(0,"Select Fund Program")
    
    fund_program_list_var = StringVar(window)
    fund_program_list_var.set(fund_program_list_series[0])
    fund_program_list_entry = OptionMenu(subframe, fund_program_list_var, *fund_program_list_series)
    fund_program_list_entry.config(width=40)
    fund_program_list_entry.grid(row = 2, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    label03 = Label(subframe, text = 'Profit Center', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=3,
             sticky = 'ew', padx = 5, pady = 5)
    profit_center_entry = Entry(subframe,width=20, font=('Arial', 12))
    profit_center_entry.grid(row = 3, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    ### --- Based on WBS Description Faculty Name list Input
    
    label04 = Label(subframe, text = 'Based on WBS Description Faculty Name', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label04.grid(row = 4, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    faculty_name_list_series = df3['WBS Owner Name'].tolist()
    faculty_name_list_series = set(faculty_name_list_series)
    faculty_name_list_series = sorted(faculty_name_list_series)
    faculty_name_list_series.insert(0,"Select Faculty Name")
    
    faculty_name_list_var = StringVar(window)
    faculty_name_list_var.set(faculty_name_list_series[0])
    faculty_name_list_entry = OptionMenu(subframe, faculty_name_list_var, *faculty_name_list_series)
    faculty_name_list_entry.config(width=40)
    faculty_name_list_entry.grid(row = 4, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    ### --- Status list Input
    
    label05 = Label(subframe, text = 'Status', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label05.grid(row = 5, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    status_list_series = ['Select Status','Active','Departed/Adjunct','Retired inactive', 'Departed', \
                    'On leave', 'Active-PT','On boarding','NA']
    
    
    status_list_var = StringVar(window)
    status_list_var.set(status_list_series[0])
    status_list_entry = OptionMenu(subframe, status_list_var, *status_list_series)
    status_list_entry.config(width=40)
    status_list_entry.grid(row = 5, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- WBS Owner-KAUST ID Input
    
    label06 = Label(subframe, text = 'WBS Owner-KAUST ID', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=6,
             sticky = 'ew', padx = 5, pady = 5)
    wbs_owner_KAUST_ID_entry = Entry(subframe,width=20, font=('Arial', 12))
    wbs_owner_KAUST_ID_entry.grid(row = 6, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    ### --- WBS Owner Status Input
    
    label07 = Label(subframe, text = 'WBS Owner Status', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label07.grid(row = 7, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    owner_status_list_series = ['Select Owner Status','Active','Departed/Adjunct','Retired inactive', 'Departed', \
                    'On leave', 'Active-PT','On boarding','NA']
    
    
    owner_status_list_var = StringVar(window)
    owner_status_list_var.set(owner_status_list_series[0])
    owner_status_list_entry = OptionMenu(subframe, owner_status_list_var, *owner_status_list_series)
    owner_status_list_entry.config(width=40)
    owner_status_list_entry.grid(row = 7, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- WBS Owner Name Input
    
    label08 = Label(subframe, text = 'Based on WBS Description Faculty Name', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label08.grid(row = 8, column = 0, sticky = 'ew', padx = 5, pady = 5)

    name_list_var = StringVar(window)
    name_list_var.set(faculty_name_list_series[0])
    name_list_entry = OptionMenu(subframe, faculty_name_list_var, *faculty_name_list_series)
    name_list_entry.config(width=40)
    name_list_entry.grid(row = 8, column = 1, sticky = 'ew', padx = 5, pady = 5)
    

    (wbs_descrip1, wbs_descrip_cc1, fund_program1, profit_center1, basedWBSdescrip_Faculty_Name1, status1, wbs_owner_KAUST_ID1, wbs_owner_status1, name1) = \
    ([], [], [], [], [], [], [], [], [])
    
    def add_data():
        global wbs_descrip1, wbs_descrip_cc1, fund_program1, profit_center1, basedWBSdescrip_Faculty_Name1, status1, wbs_owner_KAUST_ID1, wbs_owner_status1, name1
    
        wbs_descrip1.append(WBS_entry.get())
        wbs_descrip_cc1.append(WBS_CC_entry.get())
        fund_program1.append(fund_program_list_var.get())
        profit_center1.append(profit_center_entry.get())
        basedWBSdescrip_Faculty_Name1.append(faculty_name_list_var.get())
        status1.append(status_list_var.get())
        wbs_owner_KAUST_ID1.append(wbs_owner_KAUST_ID_entry.get())
        wbs_owner_status1.append(owner_status_list_var.get())
        name1.append(name_list_var.get())
        
        profit_center_list_var
    
        WBS_entry.delete(0, END)
        WBS_CC_entry.delete(0, END)
        fund_program_list_var.set(fund_program_list_series[0])
        profit_center_list_var.set(0, END)
        faculty_name_list_var.set(faculty_name_list_series[0])
        status_list_var.set(status_list_series[0])
        wbs_owner_KAUST_ID_entry.delete(0, END)
        owner_status_list_var.set(owner_status_list_series[0])
        name_list_var.set(faculty_name_list_series[0])
    
    def save_data():
        global wbs_descrip1, wbs_descrip_cc1, fund_program1, profit_center1, \
                basedWBSdescrip_Faculty_Name1, status1, wbs_owner_KAUST_ID1, \
                wbs_owner_status1, name1
    
        data = {
            'WBS / CC': wbs_descrip1,
            'WBS / CC Description': wbs_descrip_cc1,
            'Fund Program': fund_program1,
            'Profit Center': profit_center1,
            'Based on WBS Description Faculty Name': basedWBSdescrip_Faculty_Name1,
            'Status': status1,
            'WBS Owner-KAUST ID': wbs_owner_KAUST_ID1,
            'WBS Owner Status': wbs_owner_status1,
            'WBS Owner Name': name1
            }
        name_excel = 'new_mapping2' + '.xlsx'
        df = pd.DataFrame(data, columns=['WBS / CC',
                  'WBS / CC Description', 'Fund Program', 'Profit Center', 'Based on WBS Description Faculty Name', 'Status',
                  'WBS Owner-KAUST ID', 'WBS Owner Status', 'WBS Owner Name'])
        df.to_excel(name_excel, index=False)
        tkinter.messagebox.showinfo('Finish Process Mapping2','The data has been saved. Please close the window')
        
    
    Add = Button(
        subframe,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Add',
        bg='#f08823',
        bd=5,
        command=add_data,
        )
    Add.grid(columnspan=2, row=10, pady=20, padx=10)
    
    file = Label(
        subframe2,
        text='Save data after adding all missing data',
        width=67,
        bg='#24a7b1',
        font=('Arial', 11),
        fg='black',
        )
    file.grid(column=0, row=0, pady=20, padx=10)
    
    
    save = Button(
        subframe2,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Save',
        bg='#bdcf30',
        bd=5,
        command=save_data,
        )
    save.grid(column=0, row=2, pady=20, padx=10)
        
    
    window.mainloop()
    
    # Concat new data with Mapping2

    df3_new_mapping = pd.read_excel('new_mapping2.xlsx', dtype=object)

    df3 = pd.concat([df3, df3_new_mapping])

    df1['Check_df3'] = df1['WBS / CC'].isin(df3['WBS / CC'])
    print(df1['Check_df3'].value_counts())

    # Merge Main data with Mapping 1

    df1 = pd.merge(df1, df3, on='WBS / CC', how='left')

else:
    df1 = pd.merge(df1, df3, on='WBS / CC', how='left')

df1 = df1.drop([
    'Status',
    'Profit Center_y',
    'Based on WBS Description Faculty Name'
    ], axis=1)

df1 = df1.rename(columns={
    'WBS Owner Status': 'Status',
    'WBS Owner Name': 'Faculty Name',
    'WBS Owner-KAUST ID': 'KAUST ID'
    
    })



# Import B&P Data 


#---------------Import B&P data-------------------------------
df4 = pd.read_excel('FinancialData2021_23_Python_V2.xlsx',
                    sheet_name='B&P_GL_Grp', dtype=object)
df4.head()


# Null reeplacement B&P Data

df4.fillna('0', inplace=True)
df4['Account Number'] = df4['Account Number'].values.astype('str')
df4 = df4.drop(['GL Description', 'Level2', 'Budget Account'], axis=1)
df4 = df4.drop_duplicates()

# Cheking that all Account Number in Main Data is in B&P Data

df1['Check_df4'] = df1['Account Number'].isin(df4['Account Number'])
print(df1['Check_df4'].value_counts())
df1['Check_df4'] = df1['Check_df4'].values.astype('str')

df4 = df4.drop([
    'Level1'
    ], axis=1)

# Change the columns order

df4 = df4[[
    'Account Number',
    'GL Level 1',
    'GL Level 2',
    'S&B / OPEX'
    ]]


df_falses3 = df1[(df1['Check_df4']=='False')]
df_falses3= df_falses3['Account Number'].drop_duplicates()
print(df_falses3.to_markdown())


# ----------------Enter missing data in B&P Data-------------------------------


if 'False' in df1['Check_df4'].tolist():
    
    window = Tk()
    window.geometry('440x520')
    window.resizable(0, 0)
    window.title('Enter Missing Data in Mapping 3')
    
    
    subframe = Frame(window, bg='#24a7b1')
    subframe.grid(column=0, row=0, sticky='nsew')
     
    subframe2 = Frame(window, bg='#24a7b1')
    subframe2.grid(column=0, row=1, sticky='nsew')
    
    
    ### --- Account Number Input
    
    label00 = Label(subframe, text = 'Account Number', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=0,
             sticky = 'ew', padx = 5, pady = 5)
    account_number1_entry = Entry(subframe,width=20, font=('Arial', 12))
    account_number1_entry.grid(row = 0, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- GL Description list Input
    label10 = Label(subframe, text = 'GL Description', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=1,
             sticky = 'ew', padx = 5, pady = 5)
    gl_description_entry = Entry(subframe,width=20, font=('Arial', 12))
    gl_description_entry.grid(row = 1, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- Level 2 list Input
    label20 = Label(subframe, text = 'Level 2', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=2,
             sticky = 'ew', padx = 5, pady = 5)
    level2_entry = Entry(subframe,width=20, font=('Arial', 12))
    level2_entry.grid(row = 2, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- Budget Account list Input
    label30 = Label(subframe, text = 'Budget Account', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=3,
             sticky = 'ew', padx = 5, pady = 5)
    budget_account_entry = Entry(subframe,width=20, font=('Arial', 12))
    budget_account_entry.grid(row = 3, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- GL Level 2 list Input
    label40 = Label(subframe, text = 'GL Level 2', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=4,
             sticky = 'ew', padx = 5, pady = 5)
    gl_level2_entry = Entry(subframe,width=20, font=('Arial', 12))
    gl_level2_entry.grid(row = 4, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- Level 1 list Input
    label50 = Label(subframe, text = 'Level 1', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=5,
             sticky = 'ew', padx = 5, pady = 5)
    level1_entry = Entry(subframe,width=20, font=('Arial', 12))
    level1_entry.grid(row = 5, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    ### --- GL Level 1 list Input
    label60 = Label(subframe, text = 'GL Level 1', font = ('Arial Bold', 12), bg = 'white', fg = 'black').grid(column=0, row=6,
             sticky = 'ew', padx = 5, pady = 5)
    gl_level1_entry = Entry(subframe,width=20, font=('Arial', 12))
    gl_level1_entry.grid(row = 6, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    
    ### --- Staff / Non Staff list input
    label70 = Label(subframe, text = 'S&B / OPEX', font = ('Arial Bold', 12), bg = 'white', fg = 'black')
    label70.grid(row = 7, column = 0, sticky = 'ew', padx = 5, pady = 5)
    
    staff_non_staff_List_series = ['Select S&B / OPEX',
        'S&B',
        'OPEX'
        ]
    
    staff_non_staff_List_var = StringVar(window)
    staff_non_staff_List_var.set(staff_non_staff_List_series[0])
    staff_non_staff_list_entry = OptionMenu(subframe, staff_non_staff_List_var, *staff_non_staff_List_series)
    staff_non_staff_list_entry.config(width=40)
    staff_non_staff_list_entry.grid(row = 7, column = 1, sticky = 'ew', padx = 5, pady = 5)
    
    
    account_number1, gl_description1, level2, budget_account1, gl_level2, level1, gl_level1, staf_non_staf1 = \
    [], [], [], [], [], [], [], []
    
    def add_data():
        global account_number1, gl_description1, level2, budget_account1, gl_level2, level1, gl_level1, staf_non_staf1
    
        account_number1.append(account_number1_entry.get())
        gl_description1.append(gl_description_entry.get())
        level2.append(level2_entry.get())
        budget_account1.append(budget_account_entry.get())
        gl_level2.append(gl_level2_entry.get())
        level1.append(level1_entry.get())
        gl_level1.append(gl_level1_entry.get())
        staf_non_staf1.append(staff_non_staff_List_var.get())
    
        account_number1_entry.delete(0, END)
        gl_description_entry.delete(0, END)
        level2_entry.delete(0, END)
        budget_account_entry.delete(0, END)
        gl_level2_entry.delete(0, END)
        level1.delete(0, END)
        gl_level1.delete(0, END)
        staff_non_staff_List_var.set(staff_non_staff_List_series[0])
    
    
    def save_data():
        global account_number1, gl_description1, level2, budget_account1, gl_level2, level1, gl_level1, staf_non_staf1
    
        data = {'Account Number': account_number1, 'GL Description': gl_description1,\
                'Level 2': level2, 'Budget Account': budget_account1, 'GL Level 2': gl_level2,\
                    'Level 1': level1, 'S&B / OPEX': staf_non_staf1}
        name_excel = 'new_mapping3' + '.xlsx'
        df = pd.DataFrame(data, columns=['Account Number', 'GL Description', \
                                         'Level 2', 'Budget Account', 'GL Level 2', 'Level 1', 'S&B / OPEX'])
        df.to_excel(name_excel, index=False)
        tkinter.messagebox.showinfo('Finish Process Mapping3','The data has been saved. Please close the window')
        
    
    Add = Button(
        subframe,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Add',
        bg='#f08823',
        bd=5,
        command=add_data,
        )
    Add.grid(columnspan=2, row=8, pady=20, padx=10)
    
    file = Label(
        subframe2,
        text='Save data after adding all missing data',
        width=40,
        bg='#24a7b1',
        font=('Arial', 11),
        fg='black',
        )
    file.grid(column=0, row=0, pady=30, padx=35)
    
    
    save = Button(
        subframe2,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Save',
        bg='#bdcf30',
        bd=5,
        command=save_data,
        )
    save.grid(column=0, row=2, pady=20, padx=10)
        
    
    window.mainloop()
    
    # Concat new data with B&P Data

    df4_new_mapping = pd.read_excel('new_mapping3.xlsx', dtype=object)

    df4 = pd.concat([df4, df4_new_mapping])

    df1['Check_df4'] = df1['Account Number'].isin(df4['Account Number'])
    print(df1['Check_df4'].value_counts())

    # Merge Main data with Mapping 1

    df1 = pd.merge(df1, df4, on='Account Number', how='left')

else:
    df1 = pd.merge(df1, df4, on='Account Number', how='left')
    

# Other cleaning process

FinancialData = df1.drop([
    'WBS_Masking',
    'Code',
    'Check',
    'Check_df3', 
    'WBS / CC',
    'Check_df4',
    'Posting period'
    ], axis=1)

# Rename Columns

FinancialData = FinancialData.rename(columns={
    'Amount': 'Amount USD',
    'Profit Center_x': 'Profit Center'
    })

# Change the columns order

FinancialData = FinancialData[[
    'Ledger',
    'Posting Date',
    'Cost Center',
    'Profit Center',
    'Account Number',
    'Acc.Text',
    'WBS',
    'Amount USD',
    'Department',
    'Sector',
    'WBS / CC Description',
    'Fund Program',
    'S&B / OPEX',
    'GL Level 1',
    'GL Level 2',
    'Faculty Name',
    'Status',
    'Purchasing document',
    'Vendor',
    'KAUST ID'
    ]]

FinancialData['Sector'].fillna('0', inplace=True)

# Remove rows when sector is Division & Faculty
FinancialData = FinancialData.loc[FinancialData['Sector'] != 'Division & Faculty', :]
FinancialData = FinancialData.loc[FinancialData['Sector'] != 'Provost', :]
FinancialData = FinancialData.loc[FinancialData['Sector'] != 'VPAA', :]
FinancialData = FinancialData.loc[FinancialData['Sector'] != '0', :]




# Reeplace values

FinancialData['Department'] = FinancialData['Department'
        ].replace(['CLRI Research Park'], 'Research Park')



# Save Data in Excel

FinancialData.to_excel('ResultsFinancialData_April_2022-23.xlsx', index = False) # Change excel file name every month















