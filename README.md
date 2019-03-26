# SPO-RefreshAndUpload
Overview: Refresh and upload specific files or folders to SharePoint Online, using an excel file to map the file to a specific site, document library, and set metadata.


This is a fairly simple and straightforward PowerShell script that is used to automate the refreshing of MS Excel files and then (optionally) upload them to a SharePoint Online Document library, with optional metadata specified.

Each folder requires a mapfile.xls, which is a basic excel file with five columns that map each file to where you want it, how you want it, etc.

The columns are as follows:<br>
  ```Filename -- The exact name of the file (Case-insentive) to be mapped.```<br>
  ```SiteURL -- Link to the SPO Site this file is to be uploaded to. ```<br>
  ```Folder -- Document Library/Folder to upload this file to.``` <br>
  ```MetaData -- ; separated list of metadata in the format of key=value to apply.``` <br>
  ```RefreshOnly -- If set to TRUE, will only refresh the file (other columns can be blank in this case)<br><br>

  The columns are case-sentive, use the example file included.
 
 You invoke the script using: ```powershell SPO-RefreshAndUpload.ps1 -Source fileOrfolder```<br>
 There is an optional -UploadOnly switch that will only upload the files to SPO, without refeshing them.
 
 A mapfile.xls must exist in the same location as the file you specified, or inside the folder you specified. If a folder is given, it will loop through every file and check for an entry in the mapfile. 
 
 Note that you also need to setup your credentials the first time you use this on any site. If you set the username and password inside the script, it will check for and then create the credentials if needed.
 
  
