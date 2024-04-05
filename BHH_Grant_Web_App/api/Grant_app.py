from flask import Flask,request,send_file,render_template
import pandas as pd
import numpy as np

app = Flask(__name__)

@app.route("/", methods =['GET','POST'])
def index(): 
    return render_template("index.html")



@app.post('/view')
def view():


    Grant_fileName_path = request.files['Grant file']
    EagleSchedule_fileName_path = request.files['Eagle file']

    try:

        Masterlist_fileName_path = request.files['Master file']
    except:
        pass
    try:
        BAFI_Filename_path = request.files['Bafi file']
    except:
        BAFI_Filename_path =''
        pass
    


    Grant_fileName_path.save(Grant_fileName_path.filename)
    EagleSchedule_fileName_path.save(EagleSchedule_fileName_path.filename)


    try:
        Masterlist_fileName_path.save(Masterlist_fileName_path.filename)
    except:
        pass

    try:
     BAFI_Filename_path.save(BAFI_Filename_path.filename)
    except:
        pass


    df = pd.read_excel(Grant_fileName_path)
    

    try:
        # Open the Excel file
        Masterlist_df = pd.read_excel(Masterlist_fileName_path,sheet_name=0)

    # Continue with further processing
    except Exception as e:
    # Handle the case where the file does not exist
         print("Excel spreadsheet does not exist. Please check the file path.")
         Masterlist_df = None

    try:
         # Open the Excel file
            Bafi_df = pd.read_excel(BAFI_Filename_path)

         # Continue with further processing
    except Exception as e:
         # Handle the case where the file does not exist
         print("Excel spreadsheet does not exist. Please check the file path.")
         Bafi_df = None

#---------------------------------------------------------------------------------
    df.rename(columns = {"Application: Application Name": "Application Identifier",
                     "Application Submitted by": "Full Name",}, inplace=True)



    df['Delivery Date (Produce/Dry)'] = ''
    df['Identifier']= ''



    df[['First Name','Last Name']] = df['Full Name'].str.split(' ', n =1, expand = True)
    df = df.reindex(columns=['Identifier','Application Identifier', 'First Name',
                         'Last Name', 'Full Name','Home Phone','Mobile Phone','Street',
                         'Apt.','City',
                         'State','Zip + 4',
                         'Delivery Date (Produce/Dry)','Schedule Cycle'])
    df.reset_index(drop = True)

    df['ID'] = df['Full Name'] +'_'+ df['City']
    # Read Eagle Schedule into a dataframe

    eagle_df_overview = pd.read_excel(EagleSchedule_fileName_path)

    eagle_df_overview['ID'] = eagle_df_overview['Name']+'_'+eagle_df_overview['City']
    eagle_df_overview.rename(columns ={'BHH Only': "Grant Project Code"},inplace = True)

    df = pd.merge(left=eagle_df_overview,right=df,how='right',on='ID')
    df = df.dropna(subset=['Name'])

    df = df[['Identifier','Application Identifier', 'First Name',
                         'Last Name', 'Full Name','Home Phone','Mobile Phone',
                         'Street_y','Apt.','City_y',
                         'Zip + 4_y','County Location',
                        "HH#",'Delivery Date (Produce/Dry)',
                        'Schedule Cycle',"Grant Project Code","ID"]]

    df.rename(columns={"City_y":"City","State_y":"State","Zip + 4_y":"Zip + 4",
                   "HH#":"Number of People in family","Street_y":"Street"}, inplace= True)


#-----------------------------------------------------------------

    eagle_df = pd.ExcelFile(EagleSchedule_fileName_path)
    eagle_ss=eagle_df.sheet_names[1:]
    eagle_items = []
    for sheet_name in eagle_ss:

        schedule_df = pd.read_excel(eagle_df,sheet_name)
        schedule_df.columns = schedule_df.iloc[0]#drop the header
        schedule_df = schedule_df[1:]#make the second row the header

        eagle_items.append(schedule_df)#merge everything togther



    # Merge all the sheets togehter
    eagle_items_df = pd.concat(eagle_items,ignore_index=True)
    eagle_items_df = eagle_items_df[eagle_items_df['Name'].str.startswith('ROUTE')==False]


    eagle_items_df['ID']= eagle_items_df['Name'] +'_'+ eagle_items_df['City']

    #Merge the Grant dataframe and the Schedule Dataframe together
    df = pd.merge(left=eagle_items_df,right=df,how='right',on='ID')
    #df.drop_duplicates(subset='ID',inplace = True)#drop any duplicates
    #drop any na or blanks
    #df = df.dropna(subset=['Name'])

    if "City_y" in df.columns:
     City = "City_y"
    elif "City" in df.columns:
     City = "City"
    else:
     pass


    #create a total column
    df['Total']= 0

    df['Total'] = df['Produce'] + df['Dry']



    df = df.dropna(subset=['Name'])


    # Reorganize the Dataframe to have the desired outputs
    df = df[['Identifier','Application Identifier', 'First Name',
                         'Last Name', 'Full Name','Home Phone','Mobile Phone',
                         'Street','Apt.',City,
                         'Zip + 4','County Location',
                         'Number of People in family',
                         'Delivery Date (Produce/Dry)','Schedule Cycle','Total',
                         'Produce','Dry',"Grant Project Code","ID"]]




#--------------------------------MasterList-----------------------------
    if Masterlist_df is not None:
      if "Identifier" in Masterlist_df:
        identifier_column = "Identifier"
    elif "Indentier" in Masterlist_df:
        identifier_column = "Indentier"
    else:
     pass




    if Masterlist_df is not None:
     Masterlist_df[identifier_column]= Masterlist_df[identifier_column].map(str)
    # Merge the Dataframe and the Masterlist to get the identifier
     Masterlist_df['ID']= Masterlist_df['Application Submitted by']+ '_' + Masterlist_df['City']
    df= pd.merge(left= Masterlist_df,right = df ,on ='ID',how='right')
    if 'County Location_x' in df.columns:
      df.drop(columns=['County Location_x'], inplace=True)
    if 'Identifier_y' in df.columns:
      df.drop(columns=['Identifier_y'], inplace =True)
    if 'County Location_y' in df.columns:
       df.rename(columns={'County Location_y':'County Location'},inplace=True)
    if 'First Name_y' in df.columns:
       df.rename(columns={"First Name_y":'First Name'},inplace=True)
    if 'Last Name_y' in df.columns:
       df.rename(columns={"Last Name_y":'Last Name'},inplace=True)
    if 'Apt._y' in df.columns:
       df.rename(columns={"Apt._y":'Apt.'},inplace=True)
    if 'Identifier_x' in df.columns:
       df.rename(columns={"Identifier_x":'Identifier'},inplace=True)
    else:
      pass



#-------------------------------------------------------------------------

#Date Schedule days of the week Match the Applicant with the day of delivery

    try:
     Day_of_the_week = pd.read_excel(EagleSchedule_fileName_path, sheet_name=[0, 1, 2, 3, 4])
    except ValueError as e:
     if "Worksheet index 4 is invalid" in str(e):
        # Handle the case where the fifth sheet is missing
        Day_of_the_week = pd.read_excel(EagleSchedule_fileName_path, sheet_name=[0, 1, 2, 3])
     else:
        raise e

    try:
        schedule_df = Day_of_the_week[0]
    except IndexError:
     pass

    try:
        tuesday_df = Day_of_the_week[1]
    except IndexError:
        pass

    try:
         wednesday_df = Day_of_the_week[2]
    except IndexError:
        pass

    try:
        thursday_df = Day_of_the_week[3]
    except IndexError:
        pass



    tuesday_df.columns = tuesday_df.iloc[0]
    tuesday_df = tuesday_df[1:]

    wednesday_df.columns = wednesday_df.iloc[0]
    wednesday_df = wednesday_df[1:]

    thursday_df.columns = thursday_df.iloc[0]
    thursday_df = thursday_df[1:]

    for dow_name in df['Full Name']:
        Week_day = tuesday_df.loc[tuesday_df['Name']==dow_name,'Name']
        if Week_day.size > 0:
             df.loc[df['Full Name']==dow_name,'Delivery Date (Produce/Dry)'] = 'Tuesday'

    for dow_name in df['Full Name']:
        Week_day = wednesday_df.loc[wednesday_df['Name']==dow_name,'Name']
        if Week_day.size > 0:
             df.loc[df['Full Name']==dow_name,'Delivery Date (Produce/Dry)'] = 'Wednesday'


        for dow_name in df['Full Name']:
            Week_day = thursday_df.loc[thursday_df['Name']==dow_name,'Name']
            if Week_day.size > 0:
                 df.loc[df['Full Name']==dow_name,'Delivery Date (Produce/Dry)'] = 'Thursday'
#-------------------------------------------------------------------------

    if "Dry" in df.columns:
        Dry_variable = "Dry"
    elif "Dry_y" in df.columns:
        Dry_variable = "Dry_y"
    else:
    # Handle the case where neither column exists
    # (e.g., raise an error or create a new column)
     pass

    if "Produce" in df.columns:
        Produce_variable = "Produce"
    elif "Produce_y" in df.columns:
         Produce_variable = "Produce_y"
    else:
         pass

    if "Total" in df.columns:
        Total_variable = "Total"
    elif "Total_y" in df.columns:
        Total_variable = "Total_y"
    else:
        pass

    if "Number of People in family" in df.columns:
        HouseHold = "Number of People in family"
    elif "Number of People in family_y"in df.columns:
         HouseHold = "Number of People in family_y"
    else:
         pass

    if "Grant Project Code" in df.columns:
        Grant_Code = "Grant Project Code"
    elif "Grant Project Code_y" in df.columns:
        Grant_Code = "Grant Project Code_y"
    else:
        pass

    if "Schedule Cycle" in df.columns:
        Schedule_cycle = "Schedule Cycle"
    elif "Schedule Cycle_y" in df.columns:
        Schedule_cycle = "Schedule Cycle_y"
    else:
        pass

    if "Identifier" in df.columns:
        Identi_vari = "Identifier"
    elif "Indentifier_x" in df.columns:
        Identi_vari = "Identifier_x"
    else:
        pass

    if "Home Phone" in df.columns:
        phone_home = "Home Phone"
    elif "Home Phone_y" in df.columns:
         phone_home = "Home Phone_y"
    else:
     pass


    if "Street" in df.columns:
     Street = "Street"
    elif "Street_y" in df.columns:
     Street = "Street_y"
    else:
     pass

    if "Mobile Phone" in df.columns:
        phone_mobile = "Mobile Phone"
    elif "Mobile Phone_y" in df.columns:
        phone_mobile = "Mobile Phone_y"
    else:
        pass

    if "Apt." in df.columns:
        Apart = "Apt."
    elif "Apt._y" in df.columns:
        Apart = "Apt.__y"
    else:
        pass

    if 'Zip + 4' in df.columns:
        Zip = 'Zip + 4'
    elif 'Zip + 4_y' in df.columns:
        Zip = 'Zip + 4_y'

#-----------------------------------------------------------

    unique_values = df['Full Name'].unique()

  # Create an empty list to store the duplicates
    duplicates = []

    # Iterate through the unique values
    for value in unique_values:
    # Count the number of times the value appears in the column
         count = df['Full Name'].value_counts()[value]

    # If the value appears more than once, add it to the duplicates list
         if count > 1:
            duplicates.append(value)

    print(duplicates)

    duplicates_df = pd.DataFrame(data=duplicates)

#-----------------------------------------------------------------


    df[Dry_variable] = df[Dry_variable].fillna(0)
    df[Dry_variable] = df[Dry_variable].astype(int)

    df[Produce_variable] = df[Produce_variable].fillna(0)
    df[Produce_variable] = df[Produce_variable].astype(int)

    df[Total_variable] = df[Total_variable].fillna(0)
    df[Total_variable] = df[Total_variable].astype(int)

    Total_boxes = df[Total_variable].sum()
    Produce_Boxes_Total = df[Produce_variable].sum()
    Dry_Boxes_Total= df[Dry_variable].sum()

    Produce_Total_Orange = df.loc[df['County Location'].str.contains('^Orange'),Produce_variable].sum()
    dry_total_Orange = df.loc[df['County Location'].str.contains('^Orange'),Dry_variable].sum()

    Produce_Total_Osceola = df.loc[df['County Location'].str.contains('Osceola'),Produce_variable].sum()
    dry_total_Osceola = df.loc[df['County Location'].str.contains('Osceola'),Dry_variable].sum()

    Produce_Total_lake = df.loc[df['County Location'].str.contains('Lake'),Produce_variable].sum()
    dry_total_lake = df.loc[df['County Location'].str.contains('Lake'),Dry_variable].sum()

    Produce_Total_brevard = df.loc[df['County Location'].str.contains( 'Brevard'),Produce_variable].sum()
    dry_total_brevard = df.loc[df['County Location'].str.contains('Brevard'),Dry_variable].sum()

    Produce_Total_Seminole = df.loc[df['County Location'].str.contains('Seminole'),Produce_variable].sum()
    dry_total_Seminole = df.loc[df['County Location'].str.contains('Seminole'),Dry_variable].sum()

    Produce_Total_Volusia = df.loc[df['County Location'].str.contains( 'Volusia'),Produce_variable].sum()
    dry_total_Volusia = df.loc[df['County Location'].str.contains('Volusia'),Dry_variable].sum()

    Produce_Total_Marion = df.loc[df['County Location'].str.contains( 'Marion'),Produce_variable].sum()
    dry_total_Marion = df.loc[df['County Location'].str.contains('Marion'),Dry_variable].sum()

    analyze_Table = {'County':['Total','Orange','Osceola','Volusia','Lake',
                           'Seminole','Brevard','Marion'],
                 'Total':[Total_boxes,None,None,None,None,None,None,None],
                 'Produce Total':[Produce_Boxes_Total,Produce_Total_Orange,
                                  Produce_Total_Osceola,Produce_Total_Volusia,
                                  Produce_Total_lake,Produce_Total_Seminole,
                                  Produce_Total_brevard,Produce_Total_Marion],
                 'Dry Total':[Dry_Boxes_Total,dry_total_Orange,dry_total_Osceola,
                              dry_total_Volusia,dry_total_lake,dry_total_Seminole,
                              dry_total_brevard,dry_total_Marion]}

    analyze_df = pd.DataFrame(data=analyze_Table)

#----------------------------------------------------------------------------
    df = df[[Identi_vari,'Application Identifier', 'First Name',
                         'Last Name', 'Full Name',phone_home,phone_mobile,
                         Street,Apart,City,
                         Zip,'County Location',
                         'Delivery Date (Produce/Dry)',Schedule_cycle,HouseHold,Total_variable,
                         Produce_variable,Dry_variable,Grant_Code]]

    df.rename(columns={"Identifier_x":"Identifier","Home Phone_y":"Home Phone",
                   "Mobile Phone_y":"Mobile Phone","Street_y":"Street",
                   "Apt._y":"Apt.",
                   "Number of People in family_y":"Number of People in family",
                   "Schedule Cycle_y":"Schedule Cycle","Total_y":"Total",
                   "Produce_y":"Produce","Dry_y":"Dry",
                   "Grant Project Code_y":"Grant Project Code"},
          inplace= True)

#------------------------------------------------------------------------------
    with pd.ExcelWriter('Output_file/identifier.xlsx')as writer:
        analyze_df.to_excel(writer,sheet_name='Grant_Analysis',index=False)
        df.to_excel(writer,sheet_name='Grant_Output',index=False)
        duplicates_df.to_excel(writer,sheet_name='Duplicates',index=False)

    return  render_template("Output.html")

@app.route('/export_excel',methods=['POST'])
def export_excel():
      
    return send_file('Output_file/identifier.xlsx',as_attachment=True)

 

if __name__ == "__main__":
        app.run(debug=True)
 

