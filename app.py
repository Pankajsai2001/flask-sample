from flask import Flask, render_template, request, flash
import openpyxl
import datetime
import pandas as pd

application = Flask(__name__)
app=application
## global variables
path='appra.xlsx'
var={}

## secret key
app.config['SECRET_KEY']="my secret key"



@app.route('/')
def index():
    return render_template('index.html')

## Bussiness Value Analysis

@app.route('/appral/evaluation')
def evaluation():
    global var

    # reading excel sheet as dataframe 
    df = pd.read_excel(path,sheet_name='Business Value Analysis')

    # column headings as a list for ease access in html page
    column_names=df.columns.values

    # taking row data as nested list
    row_data= df.values.tolist()

    # a column name to fill answers
    comp='Answers'

    # keeping dataframe in global variable for future use
    var['df']=df

    return render_template('appra.html',column_names=column_names,row_data=row_data,comp=comp,zip=zip)


@app.route('/appral/client',methods=['POST'])
def client():
    global var

    ## get df from 
    df = var['df']
    ## requesting values from appra.html page form data 
    teleco_name=request.form.get("sheet_name")
    name=request.form.get("name")
    dt=datetime.datetime.now().replace(microsecond=0)
    mail=request.form.get('email')

    ## saving variable and data to gobal variable
    var['teleco_name']=teleco_name
    var['name']=name
    var['dt']=dt
    var['mail']=mail

    ## requesting table answers values from appra.html page
    l=[request.form.get('bussana'+str(i)) for i in range(1,len(df)+1)]

    df2 = df.copy()
    df2['Answers']=l
    var['df2']=df2
    df_html = df2.to_html(index=False,classes="pankaj table",border=0)

    return render_template("changes.html",df_html=df_html,sheet_name=teleco_name,name=name,time=dt,email=mail)

@app.route('/appral/dashboard')
def dashboard():
    global var
    str_dt=str(var['dt']).replace(':','.')
    df3=var['df2']
    wb=openpyxl.Workbook()
    sheet=wb.active
    sheet.title=str_dt
    wb.save(var['teleco_name']+'.xlsx')

    with pd.ExcelWriter(var['teleco_name']+'.xlsx', mode="a", if_sheet_exists="replace",engine="openpyxl") as writer:   
            df3.to_excel(writer, sheet_name=str_dt,index=False)

    flash("Submitted successfully")
    return render_template('index.html')

## Technical Fitment Analysis

@app.route('/appral/techeval')
def techeval():
    global var

    # reading excel sheet as dataframe 
    tech_df = pd.read_excel(path,sheet_name='Technical Fitment Analysis')

    # column headings as a list for ease access in html page
    column_names=tech_df.columns.values

    # taking row data as nested list
    row_data= tech_df.values.tolist()

    # a column name to fill answers
    comp='Answers'

    # keeping dataframe in global variable for future use
    var['tech_df']=tech_df

    return render_template('techfi.html',column_names=column_names,row_data=row_data,comp=comp,zip=zip)

@app.route('/appral/techfi',methods=['POST'])
def techfi():
    global var

    ## get df from 
    tech_df = var['tech_df']
    ## requesting values from appra.html page form data 
    teleco_name=request.form.get("sheet_name")
    name=request.form.get("name")
    dt=datetime.datetime.now().replace(microsecond=0)
    mail=request.form.get('email')

    ## saving variable and data to gobal variable
    var['teleco_name']=teleco_name
    var['name']=name
    var['dt']=dt
    var['mail']=mail

    ## requesting table answers values from appra.html page
    l=[request.form.get('techni'+str(i)) for i in range(1,len(tech_df)+1)]

    df2 = tech_df.copy()
    df2['Answers']=l
    var['df2']=df2
    df_html = df2.to_html(index=False,classes="pankaj table",border=0)

    return render_template("changes.html",df_html=df_html,sheet_name=teleco_name,name=name,time=dt,email=mail)


if __name__=="__main__":
    app.run(port=8800)