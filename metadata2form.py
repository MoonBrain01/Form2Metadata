from numpy import NaN
import pandas as pd
import openpyxl as op
from openpyxl.utils.dataframe import dataframe_to_rows
#from openpyxl.styles import Font, PatternFill
import os

import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("md_file", help="Name of metadata File ")
    args = parser.parse_args()
    full_path = args.md_file
    

    df_metadata = pd.read_excel(io=full_path, header=0, keep_default_na=False, dtype='str')

    choices_data = {
        'list_name': ['yn','yn'],
        'name' : ['1','0'],
        'label' :['YES','NO'],
        'image':['',''],
        'order':['','']
    }

    settings_data = {
        'form_title':[str(df_metadata['Description'][0])],
        'form_id':[''],
        'version':[''],
        'style':['pages theme-grid']
    }

    survey_cols = ['type','name','label','appearance','required', 'relevant','calculation']

    df_settings = pd.DataFrame.from_dict(settings_data, dtype='str')
    df_choices = pd.DataFrame.from_dict(choices_data, dtype='str')
    df_survey = pd.DataFrame(columns=survey_cols, dtype='str')

    is_calculation = False

    for row in df_metadata.itertuples(index=False):
        #Skip row if blank
        if (pd.isna(row.Code) or row.Code=='' ) and (pd.isna(row.Type) or row.Type=='' ): 
            continue

        quest_code = str(row.Code).lower().strip()
        quest_type = str(row.Type).lower().strip()
        quest_desc = str(row.Description).strip()
        quest_collect = ''
        quest_length = ''
        quest_format = ''
        quest_calc = ''

        if quest_type.startswith('calculat') and is_calculation == False:
            calc_code = quest_code
            calc_desc = quest_desc
            is_calculation = True
            continue

        if is_calculation:
            quest_type = 'calculate'
            quest_code = calc_code
            quest_calc = quest_desc
            quest_desc = calc_desc
            is_calculation = False

        if quest_type.startswith('collect'):
            quest_collect=quest_desc
            quest_type=''
            quest_desc=''

        if quest_type =='category':
            quest_type='select_one'

        if pd.isna(row.Type): 
            quest_type = ''
        if pd.isna(row.Code): 
            quest_code = ''
        if pd.notna(row.Length) and row.Length!='': 
            quest_length = 'Length:'+row.Length 
            
        if pd.notna(row.Format) and row.Format!='': 
            quest_format = 'Format:'+row.Format 

        df_survey.loc[len(df_survey.index)] = [
            quest_type  #Type
            ,quest_code #Name
            ,quest_desc #Label
            ,f"{quest_length} {quest_format}" #Appearence
            ,'yes'# Required
            ,quest_collect # relevant
            ,quest_calc #Calculation
        ]

    wb = op.Workbook()
    ws1 = wb.active
    ws1.title='settings'
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

    # Save the file
    file_path = os.path.dirname(full_path)
    file_name = os.path.basename(full_path)
    dest = f"{file_path}\OC-{file_name}.xlsx"

    wb.active=0
    wb.save(dest)

if __name__ =='__main__':
    main()
