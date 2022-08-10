from numpy import NaN
import pandas as pd
import openpyxl as op
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
from traitlets import default

file_name = r'Metadata MINIMISE-Death 1.3 09 Aug 2022.xlsm'
file_path = r'C:/Users/rehbgap/Python/Form2Metadata/'
full_path = f"{file_path}{file_name}"
print(full_path)
df_metadata = pd.read_excel(io=full_path, header=0, dtype='str')

choices_cols = ['list_name','label','name','image']
survey_cols = ['type','name','label','appearance','required']

df_settings = pd.DataFrame({'form_title':[str(df_metadata['Description'][0])],'form_id':'','version':'','style':['pages theme-grid']},dtype='str')
df_choices = pd.DataFrame(columns=choices_cols,dtype='str')
df_survey = pd.DataFrame(columns=survey_cols,dtype='str')

new_row = pd.DataFrame(columns=survey_cols,dtype='str')
for row in df_metadata.itertuples(index=False):
    quest_code = str(row.Code).lower().strip()
    quest_type = str(row.Type).lower().strip()

    if quest_type =='category':
        quest_type='select_one'

    #if quest_type.find('group')!=-1 :
    #    quest_code = quest_type
    #    quest_type = ''

    df_survey.loc[len(df_survey.index)] = [
        f"{quest_type if pd.notna(row.Type) else ''}", #Type
        quest_code if pd.notna(row.Code) else '', #Name
        row.Description, #Label
        #Format
        f"{'Length:'+row.Length if pd.notna(row.Length) else ''} {'Format:'+row.Format if pd.notna(row.Format) else ''}",
        'yes'
    ]

wb = op.Workbook()
ws1 = wb.active
ws1.title='Settings'
for r in dataframe_to_rows(df_settings, index=False, header=True):
    ws1.append(r)

ws2=wb.create_sheet("survey")
wb.active=1
for r in dataframe_to_rows(df_survey, index=False, header=True):
    ws2.append(r)

ws3=wb.create_sheet("choices")
wb.active=2
for r in dataframe_to_rows(df_choices, index=False, header=True):
    ws3.append(r)

wb.active=0
wb.save(f"{file_path}OC-DEF.xlsx")


