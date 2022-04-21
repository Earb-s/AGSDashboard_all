import streamlit as st
import pandas as pd
import base64
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import win32com.client
import time
import plotly.express as px
import plotly.graph_objects as go
from streamlit_folium import folium_static
import folium
import branca
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import gspread_dataframe as gd
from datetime import datetime

st.set_page_config(layout="wide")
col4, col5, col6  = st.columns((5,2,1))
col4.write("""
# AGS Dashboard and monthly update:
:barely_sunny:
Additive and Green Solution/ Research and Innovation Center / SCG .
""")



def main():
    
    data = r'AGS.csv'
    df = pd.read_csv(data,encoding='utf-8')
    ID = df["Name"].unique() 
    month = df["Month"].unique() 
    Report = df["Report status"].unique() 
    scope = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    gc = gspread.authorize(credentials)
    menu = ["AGS Dashboard","Update Project","Meeting","Customer Visit","Revenue","Expense"]
    choice = col5.radio("Please select menu for monthly update",menu)
    st.write('---') 

    if choice == "Update Project":
        stored = pd.DataFrame()
        st.subheader('Project Update')
        col1, col2 ,col3 = st.columns((1,1,1))
        selectmonth = col1.selectbox('Please select month to update', month)
        selectID = col1.selectbox('Who are you', ID)
        sortedproject = df.loc[df['Name'] == selectID]
        project = sortedproject["Project"].unique()
        selectProject = col1.selectbox('Select Project to update', project)
        st.caption(selectProject)
        

        #col1.write('Month: {}'.format(selectmonth))
        #col1.write('Researcher: {}'.format(selectID))
        #col1.write('Project: {}'.format(selectProject))

        Status= df["Status"].unique()
        Projecttype= df["Current type of project"].unique()
        ChooseStatus = col2.radio('Select status', Status)
        ReportStatus = col2.radio('Select report status', Report)
        Chooseprojecttype = col2.radio('Select current project type', Projecttype)
        string = col3.text_area('Project Progress Update', height=200)

        if col3.button('Submit'):
        #st.write(string)
            metadata = {'Month': [selectmonth],
             'Researcher': [selectID],
             'Project': [selectProject],
             'Current Status': [ChooseStatus],
             'Report Status': [ReportStatus],
             'Current Type': [Chooseprojecttype],
             'Project Progress': [string],}
            stored = pd.DataFrame(metadata)
            values = stored.values.tolist()
            sheetName = 'project_update'  
            sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1vPlnq902WzI7qWPCJjcf84bCytEqkWURL6b7RN9axXw/edit?usp=sharing")
            sh.values_append(sheetName, {'valueInputOption': 'USER_ENTERED'}, {'values': values})
        st.subheader('Summary')
        st.write(stored)
        #stored.to_csv('project_update.csv',encoding='utf-8')
        #stored.to_csv('project_update.csv', mode='a', index=True, header=False)


    elif choice == "Meeting":
        st.subheader('Update meeting in this month')
        stored2 = pd.DataFrame()
        col7, col8 ,col9 = st.columns((1,1,2))
        selectmonth = col7.selectbox('Please select month to update', month)
        selectID = col7.selectbox('Who are you', ID)
        sortedproject = df.loc[df['Name'] == selectID]
        project = sortedproject["Project"].unique()
        selectProject = col7.selectbox('Select Project to update', project)
        st.caption(selectProject)
        #col7.write('Month: {}'.format(selectmonth))
        #col7.write('Researcher: {}'.format(selectID))
        #col7.write('Project: {}'.format(selectProject))
        totalmeeting = col8.number_input('input total meeting this month', step = 1)
        if 'totalmeeting' not in st.session_state:
            st.session_state.totalmeeting = totalmeeting   
        meetingtype = df["Type of meeting"].unique()
        Choosemeeting1 = col9.selectbox('meeting type1', meetingtype)
        Choosemeeting2 = col9.selectbox('meeting type2', meetingtype)
        Choosemeeting3 = col9.selectbox('meeting type3', meetingtype)
        Choosemeeting4 = col9.selectbox('meeting type4', meetingtype)
        Choosemeeting5 = col9.text_input('Other meeting type')
        #stored2 = pd.DataFrame()
        submit2 = col9.button('Submit')
        if submit2:
        
            metadata2 = {'Month': [selectmonth],
                         'Researcher': [selectID],
                         'Project': [selectProject],
                         'Total Meeting': [totalmeeting],
                         'Meeting1': [Choosemeeting1],
                         'Meeting2': [Choosemeeting2],
                         'Meeting3': [Choosemeeting3],
                         'Meeting4': [Choosemeeting4],
                         'Meeting5': [Choosemeeting5],}

            stored2 = pd.DataFrame(metadata2)
            values = stored2.values.tolist()
            sheetName = 'meeting_update'  
            sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1vPlnq902WzI7qWPCJjcf84bCytEqkWURL6b7RN9axXw/edit?usp=sharing")
            sh.values_append(sheetName, {'valueInputOption': 'USER_ENTERED'}, {'values': values})
            #stored2.to_csv('meeting_update.csv',encoding='utf-8')
            #stored2.to_csv('meeting_update.csv', mode='a', index=True, header=False)
        st.subheader('Summary')
            #totalmeeting = stored2.iloc[0,0]
        st.write(stored2)
        
        
    elif choice == "Customer Visit":
        st.subheader('Customer Visit in this month')
        stored3 = pd.DataFrame()
        col10,col11 ,col12 = st.columns((1,1,1))
        selectmonth = col10.selectbox('Please select month to update', month)
        selectID = col10.selectbox('Who are you', ID)
        #col10.write('Month: {}'.format(selectmonth))
        #col10.write('Researcher: {}'.format(selectID))
        #col10.write('Project: {}'.format(selectProject))
        #totalvisit = col11.number_input('input total visit this month', step = 1)
        data = r'location.csv'
        dflocation = pd.read_csv(data,encoding='utf-8')
        province = dflocation["province"].unique()
        date = col10.date_input("Visit date")
        datestr = date.strftime("%Y/%m/%d")
        selectprovince = col12.selectbox('Select Province', province)
        sortedprovince = dflocation.loc[dflocation['province'] == selectprovince]
        district = sortedprovince["district"].unique()
        selectdistrict = col12.selectbox('Select district', district)
        sorteddistrict = sortedprovince.loc[sortedprovince['district'] == selectdistrict]
        
        subdistrict = sorteddistrict["subdistrict"].unique()
        selectsubdistrict = col12.selectbox('Select subdistrict', subdistrict)
        sortedsubdistrict = sorteddistrict.loc[sorteddistrict['subdistrict'] == selectsubdistrict]
        customer_name = col11.text_input('Customer name or place')
        customer_details = col11.text_area('Customer detail: name of participants/objective/output', height=200)
        
        if col12.button('Submit'):
        #st.write(string)
            lat = sortedsubdistrict.iloc[0,4]
            long = sortedsubdistrict.iloc[0,5]
            metadata3 = {'Date':[datestr],
                         'ID':[selectID], 
                         'Customer name':[customer_name],
             'Latitude': [lat],
             'Longtitude': [long],
             'Details': [customer_details],
                       }

            stored3 = pd.DataFrame(metadata3)
            lat = stored3.iloc[0,3]
            long = stored3.iloc[0,4]
            m = folium.Map(location=[lat, long], zoom_start=10)
            tooltip = "This place"
            folium.Marker(
            [lat, long], popup="This place", tooltip=tooltip
            ).add_to(m)
            folium_static(m)
            values = stored3.values.tolist()
            sheetName = 'Customer_update'  
            sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1vPlnq902WzI7qWPCJjcf84bCytEqkWURL6b7RN9axXw/edit?usp=sharing")
            sh.values_append(sheetName, {'valueInputOption': 'USER_ENTERED'}, {'values': values})
            #stored3.to_csv('Customer_update.csv',encoding='utf-8')
            #stored3.to_csv('Customer_update.csv', mode='a', index=True, header=False)
        st.subheader('Summary')       
        st.write(stored3)
        
    elif choice == "Revenue":
        st.subheader('Update Revenue this month')
        stored4 = pd.DataFrame()
        col13,col14 ,col15 = st.columns((1,1,1))
        selectmonth = col13.selectbox('Please select month to update', month)
        selectID = col13.selectbox('Who are you', ID)
        Revfrom = col14.text_input('Revenue from')
        amount = col14.number_input('Amount (MB)', step = 0.01)
        details = col15.text_area('Details', height=200)
        if col15.button('Submit'):
            metadata4 = {'Month':[selectmonth],
             'ID': [selectID],
             'Revenue from': [Revfrom],
             'Amount':[amount], 
             'Details':[details],}

            stored4 = pd.DataFrame(metadata4)
            values = stored4.values.tolist()
            sheetName = 'Revenue_update'  
            sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1vPlnq902WzI7qWPCJjcf84bCytEqkWURL6b7RN9axXw/edit?usp=sharing")
            sh.values_append(sheetName, {'valueInputOption': 'USER_ENTERED'}, {'values': values})
            #stored4.to_csv('Revenue_update.csv',encoding='utf-8')
            #stored4.to_csv('Revenue_update.csv', mode='a', index=True, header=False)
        st.subheader('Summary')       
        st.write(stored4)
#if col6.button('Finish and submit all sections'):
             #metadata = {'Month': [selectmonth],
             #'Researcher': [selectID],
             #'Project': [selectProject],
            # 'Current Status': [ChooseStatus],
            # 'Current Type': [Chooseprojecttype],
           #  'Project Progress': [string],}
    #stored = pd.DataFrame(metadata)
    elif choice == "Expense":
        st.subheader('Update Expense')
        concat = pd.DataFrame()
        col16,col17 ,col18,col19 = st.columns((1,2,2,2))
        data = r'FOGA.csv'
        dfFOGA = pd.read_csv(data,encoding='utf-8')
        selectmonth = col16.selectbox('Please select month to update', month)
        E1 = dfFOGA.iloc[0,1]
        E2 = dfFOGA.iloc[1,1]
        E3 = dfFOGA.iloc[2,1]
        E4 = dfFOGA.iloc[3,1]
        E5 = dfFOGA.iloc[4,1]
        E6 = dfFOGA.iloc[5,1]
        E7 = dfFOGA.iloc[6,1]
        E8 = dfFOGA.iloc[7,1]
        E9 = dfFOGA.iloc[8,1]
        
        col17.write('Please fill actual expense (Baht)')
        AE1 = col17.number_input(E1 ,step= 1000)
        AE2 = col17.number_input(E2 ,step= 1000)
        AE3 = col17.number_input(E3 ,step= 1000)
        col18.write('Please fill actual expense (Baht)')
        AE4 = col18.number_input(E4 ,step= 1000)
        AE5 = col18.number_input(E5 ,step= 1000)
        AE6 = col18.number_input(E6 ,step= 1000)
        col19.write('Please fill actual expense (Baht)')
        AE7 = col19.number_input(E7 ,step= 1000)
        AE8 = col19.number_input(E8 ,step= 1000)
        AE9 = col19.number_input(E9 ,step= 1000)
        if col19.button('Submit'):
            metadata5 = {E1:[AE1],
             E2: [AE2],
             E3: [AE3],
             E4: [AE4], 
             E5: [AE5],
             E6:[AE6],
             E7: [AE7],
             E8: [AE8],
             E9: [AE9], 
               }

            stored5 = pd.DataFrame(metadata5)
            Tstored5 = stored5.transpose()
            
            Tstored5d = pd.DataFrame(Tstored5)
            Tstored5d = Tstored5.columns = ['Actual_{}'.format(selectmonth)]
            if  selectmonth == 'Jan':
                planmonth = dfFOGA.iloc[:,2]
            elif selectmonth == 'Feb':
                planmonth = dfFOGA.iloc[:,3]
            elif selectmonth == 'Mar':
                planmonth = dfFOGA.iloc[:,4]
            elif selectmonth == 'Apr':
                planmonth = dfFOGA.iloc[:,5]
            elif selectmonth == 'May':
                planmonth = dfFOGA.iloc[:,6]
            elif selectmonth == 'Jun':
                planmonth = dfFOGA.iloc[:,7]
            elif selectmonth == 'Jul':
                planmonth = dfFOGA.iloc[:,8]
            elif selectmonth == 'Aug':
                planmonth = dfFOGA.iloc[:,9]
            elif selectmonth == 'Sep':
                planmonth = dfFOGA.iloc[:,10]
            elif selectmonth == 'Oct':
                planmonth = dfFOGA.iloc[:,11]
            elif selectmonth == 'Nov':
                planmonth = dfFOGA.iloc[:,12]
            elif selectmonth == 'Dec':
                planmonth = dfFOGA.iloc[:,13]
            planmonth.reset_index(drop=True, inplace=True)
            Tstored5.reset_index(drop=False, inplace=True)
            #Tstored5d = pd.DataFrame (Tstored5d, columns = ['Actual{}'.format(selectmonth)])
            concat = pd.concat([Tstored5, planmonth],axis=1,)
            
            #concat = concat.reset_index()
            concat2 = concat.transpose()
            concat2 =concat2.astype(str)
            concat3 = concat2.iloc[1:3,:]
            concat2.reset_index(drop=False, inplace=True)
            concat3.reset_index(drop=False, inplace=True)
            valuesa = concat3.to_numpy(copy=True)
            values =valuesa.tolist()
            #values = np.insert(values, 0, c, axis=1)
            sheetName = 'Expense_update'  
            sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1vPlnq902WzI7qWPCJjcf84bCytEqkWURL6b7RN9axXw/edit?usp=sharing")
            sh.values_append(sheetName, {'valueInputOption': 'USER_ENTERED'}, {'values': values})
            #concat2.to_csv('Expense_update.csv',encoding='utf-8')
            #concat3.to_csv('Expense_update.csv', mode='a', index=True, header=False)
        
        st.subheader('Summary')       
    
        concat
        valuesa

if __name__ == '__main__':
    main()