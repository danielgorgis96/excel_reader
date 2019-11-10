from flask import Flask, render_template, request, send_file
app = Flask(__name__)
import pandas as pd
from werkzeug import secure_filename




@app.route('/Menu', methods=['GET','POST'])
def Menu_func():
    return render_template('Menu.html')


@app.route('/Guide', methods=['GET','POST'])
def index_test():
    return render_template('Guide.html')


@app.route('/', methods=['GET','POST'])
def index():
    return render_template('index.html')

@app.route('/Program', methods=['GET','POST'])
def home():
    if request.method == 'POST':
        
        #request form fields 
        #request form excel files
        f1 = request.files['file1']
        f2 = request.files['file2']
        f3 = request.files['file3']

        #request form column names
        f1_column = request.form['ColumnnameFile1']
        f2_column = request.form['ColumnnameFile2']
        filter_column = request.form['ColumnnameFile3']
        #request form sheet names
        f1_sheet = request.form['Sheetname1']
        f2_sheet = request.form['Sheetname2']
        filter_sheet = request.form['Sheetname3']
        
        #request form with searchable cols
        split_string = request.form['SearchCols']
        List_Of_Cols_To_Search = [x.strip() for x in split_string.split(',')]


        #save files
        f1.save(secure_filename(f1.filename))
        f2.save(secure_filename(f2.filename))
        f3.save(secure_filename(f3.filename))
        
        #use pd pandas to read .xlsx
        df1 = pd.read_excel(f1, sheet_name = f1_sheet)
        df2 = pd.read_excel(f2, sheet_name = f2_sheet)
        df3 = pd.read_excel(f3, sheet_name = filter_sheet)

        #---------------------------#---------------------#--------------------------#--------------------#--------------------#
        #Convert DF vars' to str. And to uppercase letters
        df1[f1_column] = df1[f1_column].str.upper() 
        df2[f2_column] = df2[f2_column].str.upper()
        df3[filter_column] = df3[filter_column].str.upper()

        #Remove all hosts that aren't unique from df1
        filtered_unique_hosts_DF = df1[df1[f1_column].isin(df2[f2_column])]

        #get dataframe column to list 
        filter_list = []

        for i in df3[filter_column]:
            filter_list.append(str(i))
        

        #Loop through all searchable cols
        for i in List_Of_Cols_To_Search:
            Last_Result_DF = filtered_unique_hosts_DF[filtered_unique_hosts_DF[i].isin(filter_list)]




        #Convert variable to excel for download
        Last_Result_DF.to_excel('LastResult.xlsx')


        return send_file('LastResult.xlsx') 
  

    return render_template('Program.html')


@app.route('/upload', methods=['GET','POST'])

def send():
    
    if request.method == 'POST':
        f = request.files['file1']
        f.save(secure_filename(f.filename))

        df = pd.read_excel(f)

        oneList = []
        for i in df['HOST ID']:
            oneList.append(str(i))
        print(oneList[5])

        df.to_excel('SomeFile.xlsx')
        
        
        return send_file('SomeFile.xlsx') 
        

    return render_template('home.html')


if __name__ == "__main__":
    app.run(debug=True)
