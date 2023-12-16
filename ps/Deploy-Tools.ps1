 param (
    [string]$AdminName = "jon.butler@veloxfintech.com",
    [string]$mode = "packup",
    [string]$desktopFolder
 )



# load the utilites script
$WORKINGDIR = Get-Location
$MYHOME=Get-Content -Path Env:\HOMEPATH
$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Tools-utils.ps1"
. $utils_file

$tmp_zipfile=Join-Path -Path $MYHOME -ChildPath "Downloads"
$tmp_zipfile=Join-Path -Path $tmp_zipfile -ChildPath "Tools.zip"
$deploy_folder=Join-Path -Path $MYHOME -ChildPath "Deploy"
$target_app=Join-Path -Path $deploy_folder -ChildPath "vbautils.xlsm"
$icon_file=Join-Path -Path $deploy_folder -ChildPath "velox.ico"

if (([string]::IsNullOrEmpty($desktopFolder))) {
    $desktopFolder=Join-Path -Path $MYHOME -ChildPath "Desktop"
}

$output = $null

Log-Output -result ([ref]$output) `
        -status "NOTIFY" `
        -action "Set Argument" `
        -object $mode `
        -message "mode ="
Write-Host $output

Log-Output -result ([ref]$output) `
        -status "NOTIFY" `
        -action "Set Argument" `
        -object $AdminName `
        -message "AdminName ="
Write-Host $output

Log-Output -result ([ref]$output) `
        -status "NOTIFY" `
        -action "Set Argument" `
        -object $desktopFolder `
        -message "desktopFolder ="
Write-Host $output



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

    Get-OneDrive $AdminName `
        "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
        "Tools.zip" `
        "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Deploy" `
        $tmp_zipfile

    Install-ToolsZip $tmp_zipfile `
                $deploy_folder `
                $desktopFolder `
                $target_app `
                $icon_file

}
else {
    Log-Output -result ([ref]$output) `
            -status "ERROR" `
            -action "Check Mode" `
            -object $mode `
            -message "mode must be packup or unpack"
    Write-Host $output
}
