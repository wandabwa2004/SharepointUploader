# Sharepoint Uploader

## Version 0.1.0

The Sharepoint Uploader is a tool designed using Python and Streamlit to help you upload files to an online Sharepoint location. This works with Sharepoint  365 but can be  modified to fit earlier Sharepoint versions. Current functionality includes:

* Specifying the folder path to the files to be  uploaded (Source URL).
* Summary information of the files to be  uploaded. 
* Specification of Sharepoint login and related upload details. 
* Creation of  a folder based on the todays date format in the base URL that is user specified.
* Upload of the files matching  the  specified extension (currently .xlsx) to the  new folder in the base URL. File format can be changed

## Notes on Usage
* A deployed version of  the app can be found here https://sharepointuploader.herokuapp.com/. The app can also be cloned and run locally using streamlit: `streamlit run SharepointUploader.py`. When doing this, ensure you have the required modules listed in the requirements file.
* Make sure  the account  details for  accessing  Sharepoint on your  domain are valid. Normally, the username  is your domain specific email and password. 

## Bugs, Enhancements and Comments
All comments, bug reports and enhancement requests are welcome. To do so, please submit a new issue and I will work hard on improving the app. 

## Future Functionality
Future functionality will likely include:
* Option to specify the file formsts  to be uploaded in a folder with mixed file types. 
* Email trigger to the  username once  the files are all uploaded. 
