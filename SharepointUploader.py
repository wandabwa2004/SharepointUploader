# ---- DO NOT DELETE -- Author - Herman Wandabwa --- wandabwa2004@gmail.com 
# GitHub -- https://github.com/wandabwa2004/DS_Projects


import pandas  as pd 
import glob
import os
import pendulum 
import streamlit  as st #Works with streamlit 0.82.0
import time 

#Sharepoint related packages 
from shareplum import Site #Sharepoint  0.5.1
from shareplum import Office365
from shareplum.site import Version

# SETTING PAGE CONFIG TO WIDE MODE

st.set_page_config(layout="wide")



def upload_files_to_sharepoint(shpt_folder, path, share_point, shrpnt_site, columnname, username_shrpt,password_shrpt,input_folder):
     
   
    authcookie = Office365(str(share_point), username=username_shrpt, password=password_shrpt).GetCookies() #Connection details to Office 365
    site = Site(str(shrpnt_site), version=Version.v365, authcookie=authcookie) #Sharepoint  365 
    folder = site.Folder(str(shpt_folder)+'/'+foldername) #Creates a new folder in the base folder path.

    files = glob.glob(path+"\\*.xlsx") #Can be changed to any file format e.g. .csv or  .txt
    for file in files:
        try:

            with open(file, mode='rb') as rowFile:
                fileContent = rowFile.read()
            folder.upload_file(fileContent, os.path.basename(file)) #Upload all files matchng the extension above.
        
        except  Exception:
            columnname.error("Sorry, there was an issue uploading  the file(s). Please re-run the uploading  process.")
    columnname.info("All done! The files must all be here: -->"+str(shrpnt_site)+str(shpt_folder)+'/'+foldername) #Displays the full path to the final URL 


with st.beta_expander("Upload to Sharepoint:",expanded=True):
    col1,col2,col3 = st.beta_columns(3)
    col1.header("Utilities")
    upload_url = col1.text_input("Please Copy and Paste the folder path to the files to be uploaded")
    if (upload_url == ""):
        col1.warning("Sorry you cannot  leave  this  field empty. Please enter a valid path.")
    else:

        col2.header("File(s) Directory")
        in_folder = upload_url       
        todaysdate = pendulum.now()
        foldername = todaysdate.strftime('%d-%m-%Y') #This format can be changed to fit any other. 
        dir_list = os.listdir(in_folder)
        col2.info("Files location: ---> "+str(in_folder))
        col2.info("Number of files in folder: --->"+str(len(glob.glob(in_folder+"\\*.xlsx")))) #This  can be any other file format e.g. .CSV, .txt etc

        upload_choice = col2.selectbox("Upload to Sharepoint?",("No","Yes"))
        if (upload_choice == "Yes"):
            col3.header("SharePoint Details")
            username_shrpt = col3.text_input("Enter your  Sharepoint username. Should be  your  Sharepoint email address","")    
            if (username_shrpt ==""):
                col3.warning("Sorry the username field is blank. Please enter valid Sharepoint email.")
            else:
                password_shrpt = col3.text_input("Enter your  Sharepoint Password",type="password") #Office 365 password
                if (password_shrpt ==""):
                    col3.warning("Sorry the password field MUST be filled. Please enter a valid Sharepoint password.")
                else:
                    
                    sharepoint_url = col3.text_input("Enter your Sharepoint URL e.g. https://xxxx.sharepoint.com/","")
                    if (sharepoint_url == ""):
                        col3.warning("Sorry, you CANNOT  leave this  field empty.")
                    else:
                        sharepointsite = col3.text_input("Enter your Sharepoint Site  e.g. https://xxxx.sharepoint.com/sites/yyyyy","")
                        if (sharepointsite == ""):
                            col3.warning("Sorry, you CANNOT  leave this  field empty.")
                        else:
                            doc_library = col3.text_input("Enter the folder path in your Sharepoint site where a new folder will be stored and files uploaded to it e.g. Shared Documents/New Folder","")
                            if (doc_library == ""):
                                col3.warning("Sorry, you CANNOT  leave this  field empty.")
                            else:
                              
                                st.spinner("Starting  the upload in 5 seconds ...")
                                time.sleep(5)
                                try:
                                    upload_files_to_sharepoint(doc_library, in_folder, sharepoint_url, sharepointsite, col3, username_shrpt,password_shrpt,in_folder)#Calls the function to upload the files
                                except Exception:
                                    col3.error("Sorry the upload  failed. Please re-run the  process.")
                                    


            



