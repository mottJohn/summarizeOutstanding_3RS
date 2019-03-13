import pandas as pd 
import datetime
import os

#######
#for fast checking
#######

"""
xls = pd.ExcelFile('ET Inspection Register 3206.xlsx')
xls_code = pd.ExcelFile('codesReference_3206.xlsx')
contract_no = '3206'
"""

contract_no = input('Please enter contract number: ')
xls_file = input('Please enter file path of the excel template: ')
xls_file = os.path.abspath(xls_file)

xls_code = input('Please enter the file path of the code reference: ')
xls_code = os.path.abspath(xls_code)

xls = pd.ExcelFile(xls_file)
xls_code = pd.ExcelFile(xls_code)

code = xls_code.parse('Sheet1')

#function to get the time delta between today and data of inspection
def timedelta(d):
    delta = datetime.datetime.now() - d
    return delta.days

#function to filter outstanding cases
def filterOutstanding(barge_data, code):
    
    outstanding_barge_data = barge_data[(barge_data['Observations (O)/ Reminders (R)'] == 'O') & barge_data['Date of Closed Out'].isnull()] #filter

    outstanding_barge_data_selected = outstanding_barge_data[['ID', 'Unnamed: 1', 'Unnamed: 2', 'Date of Inspection', 'Type of Inspection', 'Description', 'Outstanding 1', 'Outstanding 2', 'Remarks']] #select requiring columns

    outstanding_barge_data_selected['Outstanding Days'] = outstanding_barge_data_selected['Date of Inspection'].apply(lambda x: timedelta(x)) #calculate timedelta

    outstanding_barge_data_selected = pd.merge(code, outstanding_barge_data_selected, how='inner', left_on='Name of Barge (Short)', right_on='ID')
    outstanding_barge_data_selected['ID'] = outstanding_barge_data_selected['ID'].astype(str) + outstanding_barge_data_selected['Unnamed: 1'].astype(str) + outstanding_barge_data_selected['Unnamed: 2'].astype(int).astype(str) #concat to one column
    outstanding_barge_data_selected = outstanding_barge_data_selected.drop(['Unnamed: 1', 'Unnamed: 2'], axis = 1) #drop columns

    #change order of columns
    cols = outstanding_barge_data_selected.columns.tolist()
    cols = cols[:-2] +cols[-1:] + cols[-2:-1]
    outstanding_barge_data_selected = outstanding_barge_data_selected[cols]
    return outstanding_barge_data_selected

summaryTable = pd.DataFrame()

for i in code['Name of Barge (Short)']:
    barge_data = xls.parse(i , header= 4, na_values = '', keep_default_na = False) #default na value is empty string, ignore default NaN
    barge_data['Observations (O)/ Reminders (R)'] = barge_data['Observations (O)/ Reminders (R)'].astype(str)
    print("Working on {}".format(i))
    filtered_data = filterOutstanding(barge_data, code)

    try:
        summaryTable = pd.concat([summaryTable, filtered_data])
    except:
        summaryTable = filtered_data

pivot = summaryTable.groupby(['Name of Barge','Name of Barge (Short)']).count()['Contract']
pivot = pivot.rename(columns = {'Contract': 'Count of Outstanding Cases'})

#################################
# FOR CAV
#################################

#function to filter outstanding cases
def filterOutstanding_CAV(barge_data):
    
    outstanding_barge_data = barge_data[(barge_data['Observations (O)'] == 'O') & barge_data['Date of Closed Out'].isnull()] #filter

    outstanding_barge_data_selected = outstanding_barge_data[['ID', 'Unnamed: 1', 'Unnamed: 2', 'Date', 'Type of deviation', 'Action done by Contractor']] #select requiring columns

    outstanding_barge_data_selected['Date'] = pd.to_datetime(outstanding_barge_data_selected['Date'], format='%Y%m%d')
    outstanding_barge_data_selected['Outstanding Days'] = outstanding_barge_data_selected['Date'].apply(lambda x: timedelta(x)) #calculate timedelta

    outstanding_barge_data_selected['ID'] = outstanding_barge_data_selected['ID'].astype(str) + outstanding_barge_data_selected['Unnamed: 1'].astype(str) + outstanding_barge_data_selected['Unnamed: 2'].astype(int).astype(str) #concat to one column
    outstanding_barge_data_selected = outstanding_barge_data_selected.drop(['Unnamed: 1', 'Unnamed: 2'], axis = 1) #drop columns

    #change order of columns
    #cols = outstanding_barge_data_selected.columns.tolist()
    #cols = cols[:-2] +cols[-1:] + cols[-2:-1]
    #outstanding_barge_data_selected = outstanding_barge_data_selected[cols]
    return outstanding_barge_data_selected

print("Working on {}".format('MTRMPCAV'))

barge_data = xls.parse('MTRMPCAV' , header= 2, na_values = '', keep_default_na = False) #default na value is empty string, ignore default NaN
barge_data['Observations (O)'] = barge_data['Observations (O)'].astype(str)
filtered_data_CAV = filterOutstanding_CAV(barge_data)

#save to excel with two sheets
with pd.ExcelWriter('summaryData_{}.xlsx'.format(contract_no)) as writer:
    summaryTable.to_excel(writer, index = False, sheet_name='Summary Data')
    pivot.to_excel(writer, sheet_name='Summary Table')
    filtered_data_CAV.to_excel(writer, index = False, sheet_name='Summary Data (MTRMPCAV)')

print("DONE!")