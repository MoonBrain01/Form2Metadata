import pandas as pd
import openpyxl as op
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import time
import unicodedata as ud

import os
import sys
import argparse

# Get paramters passed via command line
parser = argparse.ArgumentParser(
    description='Convert CCTU metadata spreadsheet into a basic Open Clinica form definition file.')
parser.add_argument(
    "metadata_file", help="Full path and name of metadata file)")
args = parser.parse_args()

full_path = args.metadata_file

if not os.path.exists(full_path):
    sys.exit(f"'File not found - {full_path}")

# List of columns expected in the metadata file
md_columns = ('Code', 'Type', 'Description', 'Length', 'Format')

# List of valid metadata question types
valid_types = ('note', 'integer', 'decimal', 'category',
               'text', 'date', 'group', 'table')

# Read in the metadata file
df_excel = pd.read_excel(
    full_path, keep_default_na=False, dtype='str', sheet_name=None)

# Cycle through each worksheet in the Excel file
for ws in df_excel.keys():

    # Skip blank/empty worksheet
    if df_excel[ws].empty:
        continue

    # Retrieve a worksheet
    df_metadata = df_excel[ws]

    # Check the worksheet contains the columns expected in a metadata worksheet.
    # If any column is missing, skip to the next worksheet
    if len(list(md_col in df_metadata.columns for md_col in md_columns)) != len(md_columns):
        continue

    # Initialise dataframes
    df_choices = pd.DataFrame(columns=['list_name', 'label', 'name'])
    df_survey = pd.DataFrame(columns=['type', 'name', 'label', 'bind::oc:itemgroup',
                                      'required', 'appearance'])

    # Populate Settings worksheet
    df_settings = pd.DataFrame(
        columns=['form_title', 'form_id', 'version', 'style', 'namespaces'])

    form_title = df_metadata[df_metadata['Type'].isin(
        ['Form:', 'Form'])]['Description'].to_string(index=False)

    df_settings = df_settings.append({'form_title': form_title, 'form_id': '', 'version': '0', 'style': 'theme-grid',
                                      'namespaces': 'oc="http://openclinica.org/xforms" , OpenClinica="http://openclinica.com/odm"'}, ignore_index=True)

    ques_count = 0

    # Each group is given a unique group ID - a sequential number prefixed with group_
    group_count = 0
    # As groups can contain subgroups, this list is used as a stack of group codes.
    group_code_list = []
    # Used to indicate when a category/select question has started.
    # When True, the assumption is that the rows that follow are the choices for
    # the select question.
    is_select = False

    table_count = 0
    is_table = False

    for row in df_metadata.itertuples():
        ques_type = str(row.Type).strip().lower().strip(':')
        ques_label = str(row.Description).strip()

        # If the previous question was a category/select question,
        # assume the following row are list options as long as the question type is a number
        if is_select:
            if ques_type.isnumeric():
                if is_table:
                    list_name = table_list_code
                else:
                    list_name = ques_code
                # If the list for the first question in the table has been created,
                # do not create the others as they will be duplicates

                df_choices = df_choices.append(
                    {'list_name': list_name, 'label': ques_label, 'name': ques_type}, ignore_index=True)
                continue

            else:
                # If it is not a number, assume it is the end of the list options for the select question
                is_select = False

        # Skip row if it is not a valid question type
        if not (ques_type in valid_types or ques_type.startswith('group') or ques_type.startswith('table')):
            continue

        # Group/Table tags
        if re.search("^group\s*start$", ques_type) or re.search("^table\s*start$", ques_type):
            if re.search("^table\s*start$", ques_type):
                is_table = True
                table_count += 1
                table_list_code = f"table_{table_count:03d}"
                group_code = table_list_code if str(row.Code).strip() == '' or str(
                    row.Code).strip() == None else str(row.Code).strip()
                group_appearance = 'table-list'
            else:
                group_appearance = 'field-list'
                group_count += 1
                group_code = f"group_{group_count:03d}" if str(row.Code).strip() == '' or str(
                    row.Code).strip() == None else str(row.Code).strip()

            group_code_list.append([group_code, ques_label])

            df_survey = df_survey.append(
                {'type': "begin group", 'name': group_code, 'label': ques_label, 'bind::oc:itemgroup': '', 'required': '', 'appearance': group_appearance}, ignore_index=True)
            continue

        if re.search("^group\s*end$", ques_type) or re.search("^table\s*end$", ques_type):
            group_code, group_label = group_code_list.pop()

            # If the end of the table, reset all the flags
            if re.search("^table\s*start$", ques_type):
                is_table = False

            df_survey = df_survey.append(
                {'type': "end group", 'name': group_code, 'label': group_label, 'bind::oc:itemgroup': '', 'required': '', 'appearance': ''}, ignore_index=True)
            continue

        ques_count += 1
        ques_code = f"ques_{ques_count:04d}" if str(row.Code).strip() == '' or str(
            row.Code).strip() == None else str(row.Code).strip()

        if ques_type == 'category':
            if is_table:
                list_type = f"select_one {table_list_code}"
            else:
                list_type = f"select_one {ques_code}"

            df_survey = df_survey.append(
                {'type': list_type, 'name': ques_code, 'label': ques_label, 'bind::oc:itemgroup': 'main', 'required': 'yes', 'appearance': ''}, ignore_index=True)

            is_select = True
            continue

        if ques_type == 'note':
            # Notes cannot have a value in the Required and ItemGroup column
            df_survey = df_survey.append(
                {'type': ques_type, 'name': ques_code, 'label': ques_label, 'bind::oc:itemgroup': '', 'required': '', 'appearance': ''}, ignore_index=True)
        else:
            df_survey = df_survey.append(
                {'type': ques_type, 'name': ques_code, 'label': ques_label, 'bind::oc:itemgroup': 'main', 'required': 'yes', 'appearance': ''}, ignore_index=True)

    # De-duplicate choices dataframe
    df_choices.drop_duplicates(inplace=True, ignore_index=False)

    # Convert the Metadata dataframe into an Excel object
    # https://openpyxl.readthedocs.io/en/stable/pandas.html#:~:text=Working%20with%20Pandas%20Dataframes%20%C2%B6%20The%20openpyxl.utils.dataframe.dataframe_to_rows%20%28%29,wb.active%20for%20r%20in%20dataframe_to_rows%28df%2C%20index%3DTrue%2C%20header%3DTrue%29%3A%20ws.append%28r%29

    def df_to_excel(wb, ws_title, df):
        ws = wb.create_sheet(title=ws_title)
        # Read each row in the dataframe and add it to the worksheet
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

    # Create a workbook to hold form definition
    wb = op.Workbook()
    df_to_excel(wb, 'settings', df_settings)
    df_to_excel(wb, 'choices', df_choices)
    df_to_excel(wb, 'survey', df_survey)
    del wb['Sheet']  # Delete blank worksheet

    # Save the dataframe as an Excel file
    file_path = os.path.dirname(full_path)
    file_name = os.path.basename(full_path).split('.')[0]
    dest = f"{file_path}\OC-{re.sub(r'[^a-zA-Z0-9]','',ws)}-{int(time.time())}.xlsx"
    wb.save(dest)

    # Delete the dataframes
    del df_metadata, df_settings, df_choices, df_survey
