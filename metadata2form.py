import pandas as pd
# import openpyxl as op
# from openpyxl.utils.dataframe import dataframe_to_rows

# from openpyxl.formatting import Rule
# from openpyxl.styles import Font, PatternFill, Alignment

# import re
# import os
# import sys
#import argparse

# Get
# parser = argparse.ArgumentParser()
# parser.add_argument(
#     "metadata_file", help="Name of metadata file)")
# args = parser.parse_args()

# full_path = args.form_def

# if not os.path.exists(full_path):
#     sys.exit(f"'File not found - {full_path}")

full_path = r"G:\Documents\Python\Form2Metadata\Metadata MINIMISE-Death.xlsm"

# Read in the metadata file
df_metadata = pd.read_excel(full_path, keep_default_na=False, dtype='str')

settings_header = ('form_title', 'form_id', 'version')
choices_header = ('list_name', 'label', 'name')
survey_header = ('type', 'name', 'label', 'bind:oc:itemgroup')

# Initialise dataframes
df_choices = pd.DataFrame(columns=choices_header)
df_survey = pd.DataFrame(columns=survey_header)

df_settings = pd.DataFrame(columns=settings_header)
form_title = df_metadata[df_metadata['Type'].isin(['Form:','Form'])]['Description'].to_string(index=False)
df_settings=df_settings.append({'form_title':form_title,'form_id':'','version':''}, ignore_index=True)

valid_types = ('note','integer','decimal','category','text','note:','integer:','decimal:','category:','text:')
ques_count=0
is_select = False

for row in df_metadata.itertuples():
    ques_type=str(row.Type).strip().lower() 
    ques_label=str(row.Description).strip()

    if is_select:
        if ques_type.isnumeric():
            df_choices=df_choices.append({'list_name':ques_code,'label':ques_type,'name':ques_label},ignore_index=True)
            continue
        else:
            is_select = False

    if not (ques_type in valid_types):
        continue

    ques_count +=1
    ques_code = f"ques_{ques_count:03d}"

    if ques_type.startswith('category'):
        df_survey=df_survey.append({'type':f"select_one {ques_code}",'name':ques_code,'label':ques_label,'bind:oc:itemgroup':'main'},ignore_index=True)
        is_select = True
        continue
    df_survey=df_survey.append({'type':ques_type,'name':ques_code,'label':ques_label,'bind:oc:itemgroup':'main'},ignore_index=True)

print(df_choices)
print(df_survey)



# Convert the Metadata dataframe into an Excel object
# https://openpyxl.readthedocs.io/en/stable/pandas.html#:~:text=Working%20with%20Pandas%20Dataframes%20%C2%B6%20The%20openpyxl.utils.dataframe.dataframe_to_rows%20%28%29,wb.active%20for%20r%20in%20dataframe_to_rows%28df%2C%20index%3DTrue%2C%20header%3DTrue%29%3A%20ws.append%28r%29
# wb = op.Workbook()
# ws = wb.active

# # Worksheet title
# ws.title = re.sub(r'\W+', ' ',form_title)

# # Read each row in the dataframe and add it to the worksheet
# for r in dataframe_to_rows(df_metadata, index=False, header=True):
#     ws.append(r)

# # Format the Header Row
# for c in ws['A1:E1'][0]:
#     c.font = Font(color='FFFFFF')  # White
#     c.fill = PatternFill('solid', fgColor='808080')  # Grey

# # Format the Form Name row
# for c in ws['A2:E2'][0]:
#     c.font = Font(color='000000')  # White
#     c.fill = PatternFill('solid', fgColor='90EE90')  # Green

# # Format the colours
# for r in ws:
#     quest_type = str(r[1].value)
#     if quest_type.lower().endswith('group') or quest_type.lower().endswith('repeat'):
#         for c in r:
#             c.fill = PatternFill('solid', fgColor='FFD800')
#         continue

#     if quest_type == 'Note:' or quest_type == 'Note':
#         for c in r:
#             c.font = Font(color='FF0000')
#         continue

#     if quest_type == 'Calculate':
#         for c in r:
#             c.fill = PatternFill('solid', fgColor='FCD5B4')
#             c.alignment = Alignment(vertical='top', wrap_text=True)
#         continue

#     if quest_type == 'CALCULATE:':
#         for c in r:
#             c.alignment = Alignment(vertical='top', wrap_text=True)
#         continue

#     if quest_type in ['Category', 'Text', 'Integer', 'Decimal', 'Date']:
#         for c in r:
#             c.fill = PatternFill('solid', fgColor='AED8E6')
#             r[1].fill = PatternFill('solid', fgColor='AED8E6')
#         continue

# # Save the Excel object as an Excel file
# file_path = os.path.dirname(full_path)
# file_name = os.path.basename(full_path)
# dest = f"{file_path}\OC-{file_name}"
# wb.save(dest)
