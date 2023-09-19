# Add Notes - Add Note objects to an Open Clinica form definition to display the value of calculated questions.
# Garrie Powers 14-JUL-2023

import pandas as pd

# This is used to override deprication warnings
import warnings
warnings.filterwarnings('ignore', category=UserWarning) 

xls=pd.ExcelFile(fnm)

#Read each worksheet into a seprarate dataframe
settings=pd.read_excel(xls,sheet_name='settings',dtype=str)
choices=pd.read_excel(xls,sheet_name='choices',dtype=str)
survey=pd.read_excel(xls,sheet_name='survey',dtype=str).fillna('')

#Find all the calculate questions
mask=survey['type']=='calculate'
cals=survey[mask]['name'] 

#Create a copy of the Survey dataframe, to which the notes objects will be inserted
new_survey = survey.copy()
#Blank row used to create the Notes row to be inserted
blank_row = pd.Series('',new_survey.columns)

for qname  in cals:
    #Get then name of the calculate question
    question_name = qname.strip()
    #Find the index position of the question in the copy of the Survey
    mask = new_survey['name']==qname
    idx = new_survey.index[mask].tolist()[0]

    new_idx = idx+0.5 #Places the Note between the Calculate question and the next question
    new_name = f"x_{question_name}" # give the Note question a prefix of x_ so they are easier to identify & delete.
    new_label= "<span style=\"color:green;\">**["+question_name+"]** = ${"+question_name+"}</span>"
    relevant = "${"+question_name+"}!='NaN'" # Only display the note if the Calculate question has a value (avoids explaining what NaN means!)
    
    new_row = blank_row
    new_row['type','name','label','relevant']=['note', new_name, new_label, relevant]
    new_survey.loc[new_idx]=new_row
    
new_survey=new_survey.sort_index().reset_index(drop=True) # Resort the dataframe and remove the index column.

