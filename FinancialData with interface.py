
# -*- coding: utf-8 -*-
"""
Created on Tue Oct  4 08:19:12 2022

@author: londoncm
"""
# Import libraries
import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, Frame, END


# Change path
current_path = os.getcwd()
current_path

# Adding the path where is the data located

os.chdir(r'C:\Users\LONDONCM\Documents\Financial_Report') 

os.listdir()


# --------------Import Main Data (SAP)-------------------------

df1 = pd.read_excel('FinancialData2021_22_Python_interface.xlsx',
                    sheet_name='Data_SAP_Sept_2022-23', dtype=object)
df1.head()



# Null reeplacement Main Data

df1['Posting Date'].fillna('6-30-2022', inplace=True)
df1['Posting Date'] = pd.to_datetime(df1['Posting Date'],
        format='%m-%d-%Y')

df1['Cost Center'].fillna('0', inplace=True)

df1['Profit Center'].fillna('NA', inplace=True)

df1['WBS'].fillna('0', inplace=True)

df1['Cost Center'] = df1['Cost Center'].values.astype('str')
df1['Account Number'] = df1['Account Number'].values.astype('str')

# Remove colummns we do not use in Main Data

df1 = df1.drop([
    'Document Type',
    'Ref. Document',
    'Document Date',
    'Curr. Key of PrCtr LC',
    'In transaction currency',
    'Curr. Key Trans. Curr.',
    'Text',
    'Purchasing document',
    'Vendor',
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
df2 = pd.read_excel('FinancialData2021_22_Python_interface.xlsx',
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

# ----------------Enter missing data in Mapping1-------------------------------


if 'False' in df1['Check'].tolist():
    
    windows = Tk()
    windows.config(bg='white')
    windows.geometry('650x250')
    windows.resizable(0, 0)
    windows.title('There are missing data in Mapping 1. Please enter the missing data')
    
    (code1, department1, sector1) = ([], [], [])
    
    
    def agregar_datos():
        global code1, department1, sector1
    
        code1.append(enter_code.get())
        department1.append(enter_department.get())
        sector1.append(enter_sector.get())
    
        enter_code.delete(0, END)
        enter_department.delete(0, END)
        enter_sector.delete(0, END)
    
    
    def save_data():
        global code1, department1, sector1
    
        data = {'Code': code1, 'Department': department1, 'Sector': sector1}
        name_excel = str(file_name.get() + '.xlsx')
        df = pd.DataFrame(data, columns=['Code', 'Department', 'Sector'])
        df.to_excel(name_excel, index=False)
        file_name.delete(0, END)
    
    
    frame1 = Frame(windows, bg='gray15')
    frame1.grid(column=0, row=0, sticky='nsew')
    frame2 = Frame(windows, bg='gray16')
    frame2.grid(column=1, row=0, sticky='nsew')
    
    Code = Label(frame1, text='Code', width=16).grid(column=0, row=0,
            pady=20, padx=10)
    enter_code = Entry(frame1, width=20, font=('Arial', 14))
    enter_code.grid(column=1, row=0)
    
    Department = Label(frame1, text='Department', width=16).grid(column=0,
            row=1, pady=20, padx=10)
    enter_department = Entry(frame1, width=20, font=('Arial', 14))
    enter_department.grid(column=1, row=1)
    
    Sector = Label(frame1, text='Sector', width=16).grid(column=0, row=2,
            pady=20, padx=10)
    enter_sector = Entry(frame1, width=20, font=('Arial', 14))
    enter_sector.grid(column=1, row=2)
    
    Add = Button(
        frame1,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Add',
        bg='orange',
        bd=5,
        command=agregar_datos,
        )
    Add.grid(columnspan=2, row=5, pady=20, padx=10)
    
    file = Label(
        frame2,
        text='Enter file name',
        width=25,
        bg='gray16',
        font=('Arial', 12, 'bold'),
        fg='white',
        )
    file.grid(column=0, row=0, pady=20, padx=10)
    
    file_name = Entry(frame2, width=23, font=('Arial', 14),
                      highlightbackground='green', highlightthickness=4)
    file_name.grid(column=0, row=1, pady=1, padx=10)
    
    save = Button(
        frame2,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Save',
        bg='green2',
        bd=5,
        command=save_data,
        )
    save.grid(column=0, row=2, pady=20, padx=10)
    
    windows.mainloop()
    
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

df3 = pd.read_excel('FinancialData2021_22_Python_interface.xlsx',
                    sheet_name='Mapping2', dtype=object)
df3.head()

df3['WBS / CC Description'].fillna('NA', inplace=True)
df3['Status'].fillna('NA', inplace=True)
df3['Fund Program'].fillna('NA', inplace=True)
df3['Faculty Name'].fillna('NA', inplace=True)
df3.fillna('0', inplace=True)
df3 = df3.drop_duplicates()

df3['WBS Description'] = df3['WBS Description'].values.astype('str')

df3['Fund Program'] = df3['Fund Program'].values.astype('str')
df3['Profit Center'] = df3['Profit Center'].values.astype('str')
df3['Faculty Name'] = df3['Faculty Name'].values.astype('str')

def mapping2_code(row):
    
    if row['WBS']== '0':
       row['WBS Description'] = row['Cost Center']
    else:
       row['WBS Description'] = row['WBS']
    return row   

df1 = df1.apply(mapping2_code, axis=1)

df1['WBS Description'] = df1['WBS Description'].values.astype('str')


df1['Check_df3'] = df1['WBS Description'].isin(df3['WBS Description'])
print(df1['Check_df3'].value_counts())
df1['Check_df3'] = df1['Check_df3'].values.astype('str')

df3 = df3.drop([
    'Profit Center', 
    'Based on WBS Description Faculty Name ',
    'Name'
    ], axis=1)

df_falses2 = df1[(df1['Check_df3']=='False')]
df_falses2= df_falses2['WBS Description'].drop_duplicates()
print(df_falses2.to_markdown())


# ----------------Enter missing data in Mapping2-------------------------------


if 'False' in df1['Check_df3'].tolist():
    
    windows = Tk()
    windows.config(bg='white')
    windows.geometry('650x400')
    windows.resizable(0, 0)
    windows.title('There are missing data in Mapping 2. Please enter the missing data')
    
    (wbs_descrip1, wbs_descrip_cc1, fund_program1, status1, faculty_name1) = \
    ([], [], [], [], [])
    
    
    def add_data():
        global wbs_descrip1, wbs_descrip_cc1, fund_program1, status1, faculty_name1
    
        wbs_descrip1.append(enter_wbs_descrip.get())
        wbs_descrip_cc1.append(enter_wbs_descrip_cc.get())
        fund_program1.append(enter_fund_program.get())
        status1.append(enter_status.get())
        faculty_name1.append(enter_faculty_name.get())
    
        enter_wbs_descrip.delete(0, END)
        enter_wbs_descrip_cc.delete(0, END)
        enter_fund_program.delete(0, END)
        enter_status.delete(0, END)
        enter_faculty_name.delete(0, END)
    
    
    def save_data():
        global wbs_descrip1, wbs_descrip_cc1, fund_program1, status1, faculty_name1
    
        data = {
            'WBS Description': wbs_descrip1,
            'WBS / CC Description': wbs_descrip_cc1,
            'Fund Program': fund_program1,
            'Status': status1,
            'Faculty Name': faculty_name1,
            }
        name_excel = str(file_name.get() + '.xlsx')
        df = pd.DataFrame(data, columns=['WBS Description',
                  'WBS / CC Description', 'Fund Program', 'Status',
                  'Faculty Name'])
        df.to_excel(name_excel, index=False)
        file_name.delete(0, END)
    
    
    frame1 = Frame(windows, bg='gray15')
    frame1.grid(column=0, row=0, sticky='nsew')
    frame2 = Frame(windows, bg='gray16')
    frame2.grid(column=1, row=0, sticky='nsew')
    
    WBS_Description = Label(frame1, text='WBS Description', width=16).grid(column=0, row=0,
            pady=20, padx=10)
    enter_wbs_descrip = Entry(frame1, width=20, font=('Arial', 14))
    enter_wbs_descrip.grid(column=1, row=0)
    
    WBS_CC_Description = Label(frame1, text='WBS / CC Description', width=16).grid(column=0,
            row=1, pady=20, padx=10)
    enter_wbs_descrip_cc = Entry(frame1, width=20, font=('Arial', 14))
    enter_wbs_descrip_cc.grid(column=1, row=1)
    
    Fund_Program = Label(frame1, text='Fund Program', width=16).grid(column=0, row=2,
            pady=20, padx=10)
    enter_fund_program = Entry(frame1, width=20, font=('Arial', 14))
    enter_fund_program.grid(column=1, row=2)
    
    Status = Label(frame1, text='Status', width=16).grid(column=0, row=3,
            pady=20, padx=10)
    enter_status = Entry(frame1, width=20, font=('Arial', 14))
    enter_status.grid(column=1, row=3)
    
    Faculty_Name = Label(frame1, text='Faculty Name', width=16).grid(column=0, row=4,
            pady=20, padx=10)
    enter_faculty_name = Entry(frame1, width=20, font=('Arial', 14))
    enter_faculty_name.grid(column=1, row=4)
    
    Add = Button(
        frame1,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Add',
        bg='orange',
        bd=5,
        command=add_data,
        )
    Add.grid(columnspan=2, row=5, pady=20, padx=10)
    
    file = Label(
        frame2,
        text='Enter file name',
        width=25,
        bg='gray16',
        font=('Arial', 12, 'bold'),
        fg='white',
        )
    file.grid(column=0, row=0, pady=20, padx=10)
    
    file_name = Entry(frame2, width=23, font=('Arial', 14),
                      highlightbackground='green', highlightthickness=4)
    file_name.grid(column=0, row=1, pady=1, padx=10)
    
    save = Button(
        frame2,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Save',
        bg='green2',
        bd=5,
        command=save_data,
        )
    save.grid(column=0, row=2, pady=20, padx=10)
    
    windows.mainloop()
    
    # Concat new data with Mapping2

    df3_new_mapping = pd.read_excel('new_mapping2.xlsx', dtype=object)

    df3 = pd.concat([df3, df3_new_mapping])

    df1['Check_df3'] = df1['WBS Description'].isin(df3['WBS Description'])
    print(df1['Check_df3'].value_counts())

    # Merge Main data with Mapping 1

    df1 = pd.merge(df1, df3, on='WBS Description', how='left')

else:
    df1 = pd.merge(df1, df3, on='WBS Description', how='left')



# Import B&P Data 


#---------------Import B&P data-------------------------------
df4 = pd.read_excel('FinancialData2021_22_Python_interface.xlsx',
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
    'Staff / Non Staff'
    ]]


df_falses3 = df1[(df1['Check_df4']=='False')]
df_falses3= df_falses3['Account Number'].drop_duplicates()
print(df_falses3.to_markdown())


# ----------------Enter missing data in B&P Data-------------------------------


if 'False' in df1['Check_df4'].tolist():
    
    windows = Tk()
    windows.config(bg='white')
    windows.geometry('650x350')
    windows.resizable(0, 0)
    windows.title('There are missing data in B&P Data. Please enter the missing data')
    
    (account_number1, gl_level1, gl_level2, staf_non_staf1) = \
    ([], [], [], [])
    
    
    def add_data():
        global account_number1, gl_level1, gl_level1, staf_non_staf1
    
        account_number1.append(enter_account_number.get())
        gl_level1.append(enter_gl_level1.get())
        gl_level2.append(enter_gl_level2.get())
        staf_non_staf1.append(enter_staf_non_staf.get())
    
        enter_account_number.delete(0, END)
        enter_gl_level1.delete(0, END)
        enter_gl_level2.delete(0, END)
        enter_staf_non_staf.delete(0, END)
    
    
    def save_data():
        global account_number1, gl_level1, gl_level2, staf_non_staf1
    
        data = {
            'Account Number': account_number1,
            'GL Level 1': gl_level1,
            'GL Level 2': gl_level2,
            'Staff / Non Staff': staf_non_staf1
            }
        name_excel = str(file_name.get() + '.xlsx')
        df = pd.DataFrame(data, columns=['Account Number',
                  'GL Level 1', 'GL Level 2', 'Staff / Non Staff'])
        df.to_excel(name_excel, index=False)
        file_name.delete(0, END)
    
    
    frame1 = Frame(windows, bg='gray15')
    frame1.grid(column=0, row=0, sticky='nsew')
    frame2 = Frame(windows, bg='gray16')
    frame2.grid(column=1, row=0, sticky='nsew')
    
    Account_Number = Label(frame1, text='Account Number', width=16).grid(column=0, row=0,
            pady=20, padx=10)
    enter_account_number = Entry(frame1, width=20, font=('Arial', 14))
    enter_account_number.grid(column=1, row=0)
    
    GL_Level1 = Label(frame1, text='GL Level 1', width=16).grid(column=0,
            row=1, pady=20, padx=10)
    enter_gl_level1 = Entry(frame1, width=20, font=('Arial', 14))
    enter_gl_level1.grid(column=1, row=1)
    
    GL_Level2 = Label(frame1, text='GL Level 2', width=16).grid(column=0, row=2,
            pady=20, padx=10)
    enter_gl_level2 = Entry(frame1, width=20, font=('Arial', 14))
    enter_gl_level2.grid(column=1, row=2)
    
    Staff_non_Staff = Label(frame1, text='Staff / Non Staff', width=16).grid(column=0, row=3,
            pady=20, padx=10)
    enter_staf_non_staf = Entry(frame1, width=20, font=('Arial', 14))
    enter_staf_non_staf.grid(column=1, row=3)
    

    Add = Button(
        frame1,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Add',
        bg='orange',
        bd=5,
        command=add_data,
        )
    Add.grid(columnspan=2, row=5, pady=20, padx=10)
    
    file = Label(
        frame2,
        text='Enter file name',
        width=25,
        bg='gray16',
        font=('Arial', 12, 'bold'),
        fg='white',
        )
    file.grid(column=0, row=0, pady=20, padx=10)
    
    file_name = Entry(frame2, width=23, font=('Arial', 14),
                      highlightbackground='green', highlightthickness=4)
    file_name.grid(column=0, row=1, pady=1, padx=10)
    
    save = Button(
        frame2,
        width=20,
        font=('Arial', 12, 'bold'),
        text='Save',
        bg='green2',
        bd=5,
        command=save_data,
        )
    save.grid(column=0, row=2, pady=20, padx=10)
    
    windows.mainloop()
    
    # Concat new data with B&P Data

    df4_new_mapping = pd.read_excel('new_mapping3.xlsx', dtype=object)

    df4 = pd.concat([df4, df4_new_mapping])

    df1['Check_df4'] = df1['Account Number'].isin(df4['Acoount Number'])
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
    'WBS Description',
    'Check_df4'
    ], axis=1)

# Rename Columns

FinancialData = FinancialData.rename(columns={
    'Amount': 'Amount USD',
    'WBS / CC Description': 'WBS Description',
    'Staff / Non Staff': 'S&B/OPEX',
    })

# Change the columns order

FinancialData = FinancialData[[
    'Ledger',
    'Posting period',
    'Posting Date',
    'Cost Center',
    'Profit Center',
    'Account Number',
    'Acc.Text',
    'WBS',
    'Amount USD',
    'Department',
    'Sector',
    'WBS Description',
    'Fund Program',
    'S&B/OPEX',
    'GL Level 1',
    'GL Level 2',
    'Faculty Name',
    'Status'
    ]]

# Reeplace values

FinancialData['S&B/OPEX'] = FinancialData['S&B/OPEX'
        ].replace(['Non Staff Cost'], 'OPEX')

FinancialData['S&B/OPEX'] = FinancialData['S&B/OPEX'
        ].replace(['Staff Cost'], 'S&B')
FinancialData['Department'] = FinancialData['Department'].replace(['KRO'
        ], 'Research Operations')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['RC VCC'], 'RC Visual Computing')

FinancialData['Department'] = FinancialData['Department'
        ].replace(['Central Workshop Core Lab'], 'Central Workshop')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Analytical Chemistry Core Lab'],
                  'Analytical Chemistry')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Animal Resources Facility Core Lab'],
                  'Animal Resources Facility')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Bioscience Core Lab'], 'Bioscience')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Coastal & Marine Resources Core Lab'],
                  'Coastal & Marine Resources')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Core Labs Operation & Support'],
                  'Operation & Support')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Greenhouse Core Lab'], 'Greenhouse')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Imaging & Characterization Core Lab'],
                  'Imaging & Characterization')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Nanofabrication Core Lab'], 'Nanofabrication')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['New Energy Oasis Facility Core Lab'],
                  'New Energy Oasis Facility')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Radiation Labelling Facility Core Lab'],
                  'Radiation Labelling Facility')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Supercomputing Core Lab'], 'Supercomputing')
FinancialData['Department'] = FinancialData['Department'
        ].replace(['Visualization Core Lab'], 'Visualization')


FinancialData['Fund Program'] = FinancialData['Fund Program'
        ].replace(['Strategic Partnership'], 'Support to asepc')

 # Export Data in Excel
 
FinancialData.to_excel('FinancialDataPython_Sept_interface_v2.xlsx', index = False)


