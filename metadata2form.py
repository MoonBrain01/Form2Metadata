from numpy import NaN
import pandas as pd
import openpyxl as op
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
import os

import argparse

def main(fp):
    parser = argparse.ArgumentParser()
    parser.add_argument("md_file", help="Name of metadata File ")
    args = parser.parse_args()
    full_path = args.md_file


    df_metadata = pd.read_excel(io=full_path, header=0, keep_default_na=False, dtype='str')

    choices_cols = ['list_name','label','name','image']
    survey_cols = ['type','name','label','appearance','required', 'relevant', 'constraint_message']

    df_settings = pd.DataFrame({'form_title':[str(df_metadata['Description'][0])],'form_id':'','version':'','style':['pages theme-grid']},dtype='str')
    df_choices = pd.DataFrame(columns=choices_cols,dtype='str')
    df_survey = pd.DataFrame(columns=survey_cols,dtype='str')

    new_row = pd.DataFrame(columns=survey_cols,dtype='str')
    for row in df_metadata.itertuples(index=False):
        #Skip row if blank
        if (pd.isna(row.Code) or row.Code=='' ) and (pd.isna(row.Type) or row.Type=='' ): 
            continue

        quest_code = str(row.Code).lower().strip()
        quest_type = str(row.Type).lower().strip()
        quest_desc = str(row.Description).strip()
        quest_collect = ''
        quest_warn = ''
        quest_length = ''
        quest_format = ''

        if quest_type.startswith('warn if'):
            quest_warn=quest_desc
            quest_type=''
            quest_desc=''

        if quest_type.startswith('collect if'):
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
            ,quest_warn #constraint message
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
    full_path = r"G:\Documents\Python\Form2Metadata\test\Metadata MINIMISE-Death 1.3 09 Aug 2022.xlsm"
    main(full_path)
