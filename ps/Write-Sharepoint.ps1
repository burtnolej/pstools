 param (
    [string]$mode = "packup"
 )


# load the utilites script
$WORKINGDIR = Get-Location
$MYHOME=Get-Content -Path Env:\HOMEPATH
$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Tools-utils.ps1"
. $utils_file


# PACKUP
$output = $null

if ($mode -eq "packup") {

    Log-Output -result ([ref]$output) `
            -status "NOTIFY" `
            -action "Run Mode" `
            -object $mode `
            -message "packing up icon and exe's for deployment"
    Write-Host $output

    Gen-Icon "C:\Users\burtn\Development\icons\darkemblem" "velox.ico"

    Create-ToolsZip "//wsl.localhost/Ubuntu/home/burtnolej/sambashare/veloxmon/excelvba" `
        @("DV.xlsm","VBAUtils.xlsm","MV.xlsm","MO.xlsm") `
        "C:\Users\burtn\Downloads" `
        "C:\Users\burtn\Development\icons\darkemblem\velox.ico"

    Write-OneDrive "jon.butler@veloxfintech.com" `
        "C:\Users\burtn\Downloads\Tools.zip" `
        "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
        "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Deploy" `
        "Tools.zip"
}
elseif ($mode -eq "unpack") {

    # UNPACKING / INSTALL
    Log-Output -result ([ref]$output) `
            -status "NOTIFY" `
            -action "Run Mode" `
            -object $mode `
            -message "Unpacking remote zip to local filesystem"
    Write-Host $output

    Get-OneDrive "jon.butler@veloxfintech.com" `
        "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
        "Tools.zip" `
        "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Deploy" `
        "C:\Users\burtn\Downloads\Tools.zip"

    Install-ToolsZip "C:\Users\burtn\Downloads\Tools.zip" `
                "C:\Users\burtn\Tools\Deploy" `
                "E:\Velox Financial Technology\OneDrive - Velox Financial Technology\Desktop" `
                "C:\Users\burtn\Tools\Deploy\vbautils.xlsm" `
                "C:\Users\burtn\Tools\Deploy\velox.ico"

}
else {
    Log-Output -result ([ref]$output) `
            -status "ERROR" `
            -action "Check Mode" `
            -object $mode `
            -message "mode must be packup or unpack"
    Write-Host $output
}
