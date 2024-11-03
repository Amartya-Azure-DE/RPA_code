from GoGlobalRPA import *
from pandas import read_excel, DataFrame
import os
import logging
import psutil
import sys

SuccessCount = 0
FailureCount = 0
StartTime = datetime.now()
logging.basicConfig(filename='syslog.log', filemode='a+',format='%(asctime)s - %(message)s', level=logging.INFO)

#Global values
globalvar = {}
configfname = 'D:\\bot\\zip_code\\RPAtemplate.xlsx' #path of the xlsx file from where we are taking the inputs change path according to your directory.
sheet1 = 'Global'
df1 = read_excel(configfname, sheet_name=sheet1)
for row in df1.iterrows():
    globalvar[row[1][0]] = row[1][1]
    #print("Global variables assigned for "+ globalvar[row[1][0]])
    logging.info("Global variables assigned for "+ globalvar[row[1][0]])

#if program is memory resident
if not (os.path.exists(globalvar['client_exe'])):
    #print("ERROR : NMS Client exe File",globalvar['client_exe'],"not found")a
    logging.info("ERROR : NMS Client exe File"+globalvar['client_exe']+"not found")
    exit()

#login
sheet2 = 'login'
df2= read_excel(configfname, sheet_name=sheet2)
if not "goglobal_ux.exe" in (i.name() for i in psutil.process_iter()):
    os.startfile(globalvar['client_exe'])
    time.sleep(2)
    e= RPAsteps(df2,globalvar)
    if e.errflag ==0 :
        print(e.errrcontext)
        logging.info(e.errcontext)
        sys.exit()
    #print("Login into the NMS client is successful")
    logging.info("Login into the NMS client is successful")
#execute
sheet3 = 'execute'
df3= read_excel(configfname, sheet_name=sheet3)
file_name=globalvar['file_name'] #excel file with data for execution
my_sheet = 'Sheet1'  # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
#
try:
    df= read_excel(file_name, sheet_name=my_sheet)
    for i in df['CIRCLE'].unique():
        for j in df[df["CIRCLE"].str.contains(i).fillna(False)].index:
         if df['Path_Id'][j]!=df['Path_Id'][j]:  # detects NAN;Null
           continue
         if df['Remarks'][j] == df['Remarks'][j]:
           continue
         if"goglobal_ux.exe" in (i.name() for i in psutil.process_iter()):
           globalvar['path_id']=str(df['Path_Id'][j])
           globalvar['new_label']=df['New_User_Label'][j]
           e=RPAsteps(df3,globalvar)
           if (e.errflag == 0 or e.errflag ==2):
               print("ERROR: Label not edited for path id " + globalvar['path_id'] + " " + e.errcontext)
               logging.info("ERROR: Label not edited for path id " + globalvar['path_id'] + " " + e.errcontext)
               df['Remarks'][j]= "Failed"
               df.to_excel(file_name, sheet_name= my_sheet, index = False)
               FailureCount = FailureCount+1
           else:
               print("SUCCESS : Label edited for path id " + globalvar['path_id'] + " ; old Label: " + pyperclip.paste() +" ; new label: " + globalvar['new_label'])
               logging.info("SUCCESS : Label edite for path id " + globalvar['path_id'] + " ; old Label: " + pyperclip.paste() +" ; new label: " + globalvar['new_label'])
               df['Remarks'][j]= "Success"
               df.to_excel(file_name, sheet_name= my_sheet, index= False)
               SuccessCount=SuccessCount+1
         else:
             logging.info("NMS client not active")
         print("Start Time:",StartTime,"Success Count:", SuccessCount, "Failure Count:", FailureCount, "Total:", SuccessCount + FailureCount)

except FileNotFoundError:
    #print("Excel File:"+globalvar['file_name'] + " not found")
    logging.info("Excel File:"+globalvar['file_name'] + " not found")
except Exception as e:
    #print("ERROR", e)
    logging.info("ERROR "+str(e))