import pandas as pd
import openpyxl as op
from openpyxl.utils.dataframe import dataframe_to_rows
import re

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

# Read in the metadata file
df_metadata = pd.read_excel(full_path, keep_default_na=False, dtype='str')

choices_cols = ('list_name', 'label', 'name')
survey_cols = ('type', 'name', 'label', 'bind::oc:itemgroup',
               'required', 'appearance')
settings_cols = ('form_title', 'form_id', 'version', 'style', 'namespaces')
# Initialise dataframes
df_choices = pd.DataFrame(columns=choices_cols)
df_survey = pd.DataFrame(columns=survey_cols)

# Populate Settings worksheet
df_settings = pd.DataFrame(columns=settings_cols)
form_title = df_metadata[df_metadata['Type'].isin(
    ['Form:', 'Form'])]['Description'].to_string(index=False)
df_settings = df_settings.append(
    dict(zip(settings_cols, (form_title, '', '0', 'theme-grid', 'oc="http://openclinica.org/xforms" , OpenClinica="http://openclinica.com/odm"'))), ignore_index=True)

valid_types = ('note', 'integer', 'decimal',
               'category', 'text', 'date', 'group')
ques_count = 0
group_count = 0
group_code_list = []
is_select = False

for row in df_metadata.itertuples():
    ques_type = str(row.Type).strip().lower().strip(':')
    ques_label = str(row.Description).strip()

    # If the previous question was a category/select question,
    # assume the following row are list options as long as the question type is a number
    if is_select:
        if ques_type.isnumeric():
            df_choices = df_choices.append(
                dict(zip(choices_cols, (ques_code, ques_label, ques_type))), ignore_index=True)
            continue
        else:
            # If it is not a number, assume it is the end of the list options for the select question
            is_select = False

    # Skip row if it is not a valid question type
    if not (ques_type in valid_types or ques_type.startswith('group')):
        continue

    if re.search("^group\s*start$", ques_type):
        group_count += 1
        group_code = f"group_{group_count:03d}"
        group_code_list.append([group_code, ques_label])

        df_survey = df_survey.append(
            dict(zip(survey_cols, ("group start", group_code, ques_label, 'main', '', 'field-list'))), ignore_index=True)
        continue

    if re.search("^group\s*end$", ques_type):
        group_code, group_label = group_code_list.pop()
        df_survey = df_survey.append(
            dict(zip(survey_cols, ("group end", group_code, group_label, 'main', '', ''))), ignore_index=True)
        continue

    ques_count += 1
    ques_code = f"ques_{ques_count:04d}"

    if ques_type == 'category':
        df_survey = df_survey.append(
            dict(zip(survey_cols, (f"select_one {ques_code}", ques_code, ques_label, 'main', 'yes', ''))), ignore_index=True)
        is_select = True
        continue

    if ques_type == 'note':
        # Notes cannot have a value in the Required and ItemGroup column
        df_survey = df_survey.append(dict(
            zip(survey_cols, (ques_type, ques_code, ques_label, '', '', ''))), ignore_index=True)
    else:
        df_survey = df_survey.append(dict(zip(
            survey_cols, (ques_type, ques_code, ques_label, 'main', 'yes', ''))), ignore_index=True)

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
dest = f"{file_path}\OC-{file_name}.xlsx"
wb.save(dest)
