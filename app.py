import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, send_file
import os

project_root = os.path.dirname(__file__)
template_path = os.path.join(project_root, './')
app = Flask(__name__, template_folder=template_path)

unique_list_of_accounts_and_themes = []
list_of_accounts_and_themes = []
indexes_of_accounts_and_themes = []
final_accountnumber_set, final_theme_set, final_sus_set_answerId, final_sus_set_answerText, final_overall_set_answerId,final_overall_set_answerText, final_pai_1_set_answerId, final_pai_1_set_answerText, final_pai_2_set_answerId, final_pai_2_set_answerText, final_tax_set_answerId, final_tax_set_answerText = ([] for i in range(12))
total_accounts_sustainable_theme,total_accounts_sustainable_plus_theme,total_accounts_income_theme,total_accounts_income_plus_theme , total_accounts_growth_theme, total_accounts_growth_plus_theme = ([] for i in range(6))

def read_excel_data(uploaded_file):
    #read excel data
    mifid_data = pd.read_excel(uploaded_file)
    mifid_account_data = pd.read_excel(uploaded_file, usecols=["INVESTMENT_ACCOUNT"])
    mifid_theme_data = pd.read_excel(uploaded_file, usecols=["THEME"])
    list_of_accounts = mifid_account_data.values.flatten()
    list_of_themes = mifid_theme_data.values.flatten()
    for i in range(len(list_of_accounts)):
        list_of_accounts_and_themes.append({'accountNumber':list_of_accounts[i], 'theme':list_of_themes[i]})
        indexes_of_accounts_and_themes.append(i)
    create_unique_theme_object(list_of_accounts_and_themes)
    create_answer_ids_texts_array(mifid_data)

def create_unique_theme_object(list_of_accounts_and_themes):
    unique_set = set()
    #create a list of accounts with different themes available
    for item in list_of_accounts_and_themes:
        unique_key = (item['accountNumber'], item['theme'])
        if unique_key not in unique_set:
            unique_set.add(unique_key)
            unique_list_of_accounts_and_themes.append(item)
            
def create_answer_ids_texts_array(mifid_data):
    for item in unique_list_of_accounts_and_themes:
        question_ids = mifid_data.loc[(mifid_data['INVESTMENT_ACCOUNT']==item['accountNumber']) & (mifid_data['THEME']==item['theme']), 'QUESTION_ID']
        answer_ids = mifid_data.loc[(mifid_data['INVESTMENT_ACCOUNT']==item['accountNumber']) & (mifid_data['THEME']==item['theme']), 'ANSWER_ID']
        answer_texts = mifid_data.loc[(mifid_data['INVESTMENT_ACCOUNT']==item['accountNumber']) & (mifid_data['THEME']==item['theme']), 'ANSWER_TEXT']
        questionIds = question_ids.values.flatten()
        answerIds = answer_ids.values.flatten()
        answerTexts = answer_texts.values.flatten()
        for i in range(len(questionIds)):
            item[questionIds[i]] = {'answerId': answerIds[i], 'answerText':answerTexts[i]}

def set_final_data_for_data_frame():
    for item in unique_list_of_accounts_and_themes:
        final_accountnumber_set.append(item['accountNumber'])
        final_theme_set.append(item['theme'])
        final_sus_set_answerId.append(item['SUS']['answerId'] if 'SUS' in item else '')
        final_sus_set_answerText.append(item['SUS']['answerText'] if 'SUS' in item else '')
        final_overall_set_answerId.append(item['OVERALL']['answerId'] if 'OVERALL' in item else '')
        final_overall_set_answerText.append(item['OVERALL']['answerText'] if 'OVERALL' in item else '')
        final_pai_1_set_answerId.append(item['PAI_1']['answerId'] if 'PAI_1' in item else '')
        final_pai_1_set_answerText.append(item['PAI_1']['answerText'] if 'PAI_1' in item else '')
        final_pai_2_set_answerId.append(item['PAI_2']['answerId'] if 'PAI_2' in item else '')
        final_pai_2_set_answerText.append(item['PAI_2']['answerText'] if 'PAI_2' in item else '')
        final_tax_set_answerId.append(item['TAX']['answerId'] if 'TAX' in item else '')
        final_tax_set_answerText.append(item['TAX']['answerText'] if 'TAX' in item else '')
    
def add_accountnumber():
    for item in unique_list_of_accounts_and_themes:
        if item['theme']=='BE_SUSTAINABLE_PLUS':
           total_accounts_sustainable_plus_theme.append(item['accountNumber'])
        if item['theme']=='BE_SUSTAINABLE':
           total_accounts_sustainable_theme.append(item['accountNumber'])
        if item['theme']=='INCOME_PLUS':
           total_accounts_income_plus_theme.append(item['accountNumber'])
        if item['theme']=='INCOME':
           total_accounts_income_theme.append(item['accountNumber'])
        if item['theme']=='GROWTH_PLUS':
           total_accounts_growth_plus_theme.append(item['accountNumber'])
        if item['theme']=='GROWTH':
           total_accounts_growth_theme.append(item['accountNumber'])
  
def draw_pie_chart():
    total_theme_labels =['BE_SUSTAINABLE', 'BE_SUSTAINABLE_PLUS', 'INCOME', 'INCOME_PLUS', 'GROWTH', 'GROWTH_PLUS']
    total_theme_counts = [len(total_accounts_sustainable_theme), len(total_accounts_sustainable_plus_theme), len(total_accounts_income_theme), len(total_accounts_income_plus_theme), len(total_accounts_growth_theme), len(total_accounts_growth_plus_theme)]
    plt.pie(total_theme_counts, labels=total_theme_labels, autopct='%1.1f%%')
    plt.title('Theme usage of accounts')   

def generate_excel_sheet():
    # Create a sample DataFrame
    data_to_update = {
        ('ACCOUNT_NUMBER', ''): final_accountnumber_set,
        ('THEME', ''): final_theme_set,
        ('SUS', 'ANSWER_ID'): final_sus_set_answerId,
        ('SUS', 'ANSWER_TEXT'): final_sus_set_answerText,
        ('OVERALL', 'ANSWER_ID'): final_overall_set_answerId,
        ('OVERALL', 'ANSWER_TEXT'): final_overall_set_answerText,
        ('PAI_1', 'ANSWER_ID'): final_pai_1_set_answerId,
        ('PAI_1', 'ANSWER_TEXT'): final_pai_1_set_answerText,
        ('PAI_2', 'ANSWER_ID'): final_pai_2_set_answerId,
        ('PAI_2', 'ANSWER_TEXT'): final_pai_2_set_answerText,
        ('TAX', 'ANSWER_ID'): final_tax_set_answerId,
        ('TAX', 'ANSWER_TEXT'): final_tax_set_answerText,
        }
    #data_to_update[('Info', 'Name')] = ['Johnyass', 'Alicewonderland', 'Bobby']
    final_data = pd.DataFrame(data_to_update)
    #plot = final_data.plot.pie(y='SUS', figsize=(5, 5))
    output_path = 'output.xlsx'
    # Set the column names as subheaders
    #final_data.columns = pd.MultiIndex.from_tuples(final_data.columns)
    final_data.to_excel(output_path)

 
def generate_data():
    set_final_data_for_data_frame()
    add_accountnumber()
    draw_pie_chart()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Get the uploaded file
        uploaded_file = request.files['file']
        
        print('inside post')
        
        # Read the Excel file
        read_excel_data(uploaded_file)
        
        # Manipulate the input (example: doubling all values)
        generate_data()
        
        # Create a new Excel file
        generate_excel_sheet()
        
        # Return the download link for the output file
        return f'''
            <h2>File Uploaded and Processed Successfully!</h2>
            <a href="/download">Download Processed File</a>
        '''
    return render_template('index.html')

@app.route('/download')
def download_file():
    path = 'output.xlsx'
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run()
    