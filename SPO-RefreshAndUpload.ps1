<#
 This script is for automated refreshing of excel files and uploading to SharePoint Online
 -Source can be a single file, or a Folder

 If -UploadOnly is specified, then the excel files are not refreshed before uploading.

 This requires a mapfile.xlsx file to be present in the given folder, or in the folder where the specified file is.
 mapfile.xlsx is a simple file with 4 columns: FileName, SiteURL, Folder, MetaData
	Filename is the exact name of the file you are referencing. E.g. Goal Report 2019 FL.xlsx
	SiteURL is the URL of the destination SPO site. E.g. https://southernglazers.sharepoint.com/sites/CPWS/Florida/
	Folder is the List and Folder to upload the file to. E.g. Report Center/
	MetaData is optional, and is a ; separated list of Key=Value metadata pairs. E.g. Visibility=Visible;Category=Goals

 The mapfile must contain at least two entries.

 This script requires two Powershell modules:
 Install-Module SharePointPnPPowerShellOnline
 Install-Module ImportExcel
 
 If you are running an older version of Windows/Powershell, you may need to first install PowerShellGet
    See this link for more information: https://docs.microsoft.com/en-us/powershell/gallery/overview
	
 -Michael Taylor (michaeltaylor@sgws.com)
 - 3/24/2019 - v1.0 -- Initial Commit
 
#>

param (
	[Parameter(Mandatory=$true)][string]$Source,
	[switch]$UploadOnly = $false
)

# Setup variables
$lastURL = $null					# Last used URL. So we don't keep connecting if we don't need to.
$sourceFolder = $null				# Folder we're working.
$sourceFile = $null					# Source filename (if given a single file)
$singleFile = $false				# If $true, we were passed a single file instead of a folder.
$mapfileData = $null				# Placeholder for the map data
$debug = $false						# if $true, log extra detail to (folder)\uploadDetail.log
$username = ""						# Username for setting credentials
$password = ""						# Password for setting credentials
$logFile = $null

Import-Module ImportExcel
Import-Module SharePointPnPPowerShellOnline

# PowerShell reads from top to bottom, so subroutines need to be declared before they are used. I dislike this.

# Given a path\file, a folder, and metadata, upload file. Must be connected first.
Function Upload-File {
	param ([string]$File, [string]$Folder, $MetaData)

	$Values = @{}

	if ($MetaData)
	{
		$metaSplit = $MetaData.Split(';')
		Foreach ($m in $metaSplit) {
			$mm = $m.Split('=')
			$Values.Add($mm[0], $mm[1])
		}
	}

	Try
	{
		Add-PnPFile -Checkout -Path $File -Folder $Folder -Values $Values -ErrorAction Stop
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Log-Error -Method "Add-PnPFile" -Msg "$ErrorMessage / $FailedItem"
		Exit
	}
	Finally
	{
		Log-Detail -Msg "Add-PnpFile - $File to $global:lastURL / $Folder"
	}
}


#Log Errors to c:\triggers\error.log
#We do this to capture errors that can happen before we find a proper source folder.
function Log-Error {
	param([string]$Method, [string]$Msg)

	$Time = Get-Date -UFormat "%m.%d.%Y-%H.%M"
	if ($global.logFile) {
		"[$Time] ($Method) - $Msg" | Out-File $global.logFile -Append
	}
}

# Log details to (sourceFolder)\uploadDetail.log
function Log-Detail {
	param([string]$Msg)
	$Time = Get-Date -UFormat "%m.%d.%Y-%H.%M"
	if ($debug -eq $true) {
		"[$Time] $Msg" | Out-File $sourceFolder + "\uploadDetail.log" -Append
	}
}

#Refresh a given file. If xlObj is passed, use that instead of starting a new one.

function Refresh-File {
	param([string]$File, $xlObj)


	$localxlObj = $null

	if ((Test-Path $File) -ne $true)
	{
		Log-Error -Method "RefreshFile" -Msg "File not found: $rFile"
		return
	}

	# If we were given an XlObj, use it, otherwise, create it.
	if ($xlObj) {
		$localxlObj = $xlObj
	} else {
		$localxlObj = New-Object -ComObject "Excel.Application"
		$localxlObj.Visible = $true
		$localxlObj.DisplayAlerts = $false
	}
	$wbobj = $localxlObj.Workbooks.Open($File)
	$wbobj.RefreshAll()
	$wbobj.SaveAs($File)
	$wbobj.Close()

	# If we were given the excel object, don't close it.
	if ($xlObj) {
		return
	} else {
		$localxlObj.Quit()
		return
	}
}



# Begin main script

# Make sure the $Source is a valid file or folder.
if ((Test-Path $Source) -ne $true) {
	Log-Error -Method "Startup" -Msg "Source is not valid: Test-Path = $false"
	Exit 1
}

# Make sure credentials are setup
if (-NOT (Get-PnPStoredCredential -Name https://southernglazers.sharepoint.com))
{
	Try
	{
		Add-PnPStoredCredential -Name https://southernglazers.sharepoint.com -Username $username -Password (ConvertTo-SecureString -String $password -AsPlainText -Force) -ErrorAction Stop
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Log-Error -Method "Add-PnpStoredCredential" -Msg "$ErrorMessage - $FailedItem"
		Exit 1
	}
}

#if we are given a file, parse the directory from the file. Otherwise, $source IS the directory.
If ($Source -like "*.xls*")
{
	Try
	{
		$sourceFolder = (Get-ChildItem $Source -ErrorAction Stop).Directory.FullName
	}
	Catch {
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Log-Error -Method "File.Directory.FullName" -Msg "$ErrorMessage - $FailedItem"
		Exit 1
	}
	$mapfile = $sourceFolder + "\mapfile.xlsx"
	$sourceFile = (Get-ChildItem $Source).Name
	$singleFile = $true

} else { #we got a folder
	$sourceFolder = $Source
	$mapFile = $sourceFolder + "\mapfile.xlsx"
	$singleFile = $false
}

# Try to load the mapfile. If this fails, stop the script.
Try
{
	$mapfileData = Import-Excel $mapFile -ErrorAction Stop
}
Catch
{
	$ErrorMessage = $_.Exception.Message
	$FailedItem = $_.Exception.ItemName
	Log-Error -Method "LoadMapFile" -Msg "$ErrorMessage - $FailedItem"
	Exit
}

# If we got one file, do the magic on it. Otherwise, assume it's a folder, and do a lot of magic on the whole damn thing.
if ($singleFile -eq $true) {
	$currentData = $null

	# Look for filename in the map file
	Foreach ($row in $mapfileData) {
		if ($row.Filename -eq $sourceFile) {
			$currentData = $row
			break
		}
	}

	if ($currentData) { 	# We found a reference in the mapfile. Try to connect.

		# But only if the file is NOT Set refreshonly
		if ($currentData.RefreshOnly -ne $true) {
			Try
			{
				Connect-PnPOnline -Url $currentData.SiteURL.ToString() -ErrorAction Stop
			}
			Catch
			{
				$ErrorMessage = $_.Exception.Message
				$FailedItem = $_.Exception.ItemName
				Log-Error -Method "Connect-PnPOnline" -Msg "$ErrorMessage - $FailedItem"
				Exit 1
			}
		}

		# Don't refresh if we used -UploadOnly
		If ($UploadOnly -eq $false)
		{
			Refresh-File -File $source -xlObj $null
		}

		# Don't upload if the file is set RefeshOnly.
		if ($currentData.RefreshOnly -ne $true)
		{
			Upload-File -File $source -Folder $currentData.Folder.ToString() -MetaData $currentData.MetaData
		}
		Exit
	} else { 	# File not found in the mapfile. Stop.
		Log-Error -Method "Upload-File" -Msg "File $currentFile not found in $mapFile"
		Exit 1
	}

} else { # Folder. Loop over the entire damn thing, but not subfolders. Because reasons.

	$fileList = (Get-ChildItem $sourceFolder -Filter "*.xls*")
	$totalFiles = $fileList.Count
	$filesUploaded = 0
	$filesFailed = 0
	$xlObj = $null

	if ($fileList.Count -eq 0) {
		Log-Error -Method "Load" -Msg "No files found in $sourceFolder"
		exit 1
	}
	# If we were given -Uploadonly switch, don't bother opening excel.
	if ($UploadOnly -eq $false) {
		$xlObj = New-Object -ComObject "Excel.Application"
		$xlObj.Visible = $true
		$xlObj.DisplayAlerts = $false
	}

	Foreach ($fFile in $fileList)
	{
		$currentData = $null

		If ($fFile.Name -eq "mapfile.xls") {
			Continue
		}

		Foreach ($row in $mapfileData)
		{
			if ($row.Filename -eq $fFile.Name) {
				$currentData = $row
				break
			}
		}

		If ($currentData -AND ($currentData.RefreshOnly -ne $true)) {
			# Only connect if we aren't already connected to the site.
			If ($global:lastURL -ne $currentData.SiteURL) {

				$global:lastURL = $currentData.SiteURL
				Try
				{
					Connect-PnPOnline -Url $currentData.SiteURL -ErrorAction Stop
				}
				Catch
				{
					$ErrorMessage = $_.Exception.Message
					$FailedItem = $_.Exception.ItemName
					Log-Error -Method "Connect-PnPOnline" -Msg "$ErrorMessage - $FailedItem"
					$global:lastURL = "error"
					Continue
				}
				Finally
				{
					Log-Detail -Msg "Connect-PnPOnline - " + $currentData.SiteURL.ToString() + " for $currentData.FileName"
				}
			}
			# Refresh data if we didn't specify -UploadOnly
			if ($UploadOnly -eq $false) {
				Refresh-File -File $fFile.FullName -xlObj $xlObj
			}
			if ($currentData.RefreshOnly -ne $true) {
				Upload-File -File $fFile.FullName.ToString() -Folder $currentData.Folder.ToString() -MetaData $currentData.MetaData
				$filesUploaded++
			}

		} else { # File isn't in map. Note in the log and keep going.
			Log-Error -Method "UploadFile" -Msg "File $fFile.Name is not in mapfile $mapFile"
			$filesFailed++
			Continue
		}
	}
	if ($xlObj) {
		$xlObj.Quit()
	}

	Log-Detail -Msg "Folder complete. Total files: $totalFiles, Files uploaded: $filesUploaded, Files failed: $filesFailed"
	exit
}
