import pandas as pd 
import datetime
import os

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

    outstanding_barge_data_selected = outstanding_barge_data[['ID', 'Date of Inspection', 'Outstanding 1', 'Outstanding 2']] #select requiring columns

    outstanding_barge_data_selected['Outstanding Days'] = outstanding_barge_data_selected['Date of Inspection'].apply(lambda x: timedelta(x)) #calculate timedelta

    outstanding_barge_data_selected = pd.merge(code, outstanding_barge_data_selected, how='inner', left_on='Name of Barge (Short)', right_on='ID')
    outstanding_barge_data_selected = outstanding_barge_data_selected.drop(columns = ['ID']) #drop duplicated column

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
print("DONE!")

#save to excel with two sheets
with pd.ExcelWriter('summaryData_{}.xlsx'.format(contract_no)) as writer:
    summaryTable.to_excel(writer, index = False, sheet_name='Summary Data')
    pivot.to_excel(writer, sheet_name='Summary Table')