
#https://www.microsoft.com/en-us/download/details.aspx?id=42038

#Load SharePoint CSOM Assemblies

# load the utilites script
$WORKINGDIR = Get-Location
$MYHOME=Get-Content -Path Env:\HOMEPATH
$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Tools-utils.ps1"
. $utils_file


# PACKUP


#Gen-Icon "C:\Users\burtn\Development\icons\darkemblem" "velox.ico"

Create-ToolsZip "//wsl.localhost/Ubuntu/home/burtnolej/sambashare/veloxmon/excelvba" `
    @("DV.xlsm","VBAUtils.xlsm","MV.xlsm","MO.xlsm") `
    "C:\Users\burtn\Downloads" `
    "C:\Users\burtn\Development\icons\darkemblem\velox.ico"

#Write-OneDrive "jon.butler@veloxfintech.com" `
#    "C:\Users\burtn\Downloads\Tools.zip" `
#    "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
#    "Documents"



# UNPACKING / INSTALL

#Get-OneDrive "jon.butler@veloxfintech.com" `
#    "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
#    "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Tools.zip" `
#    "C:\Users\burtn\Downloads\Tools.zip"


Install-ToolsZip "C:\Users\burtn\Downloads\Tools.zip" `
            "C:\Users\burtn\Tools\Deploy"
