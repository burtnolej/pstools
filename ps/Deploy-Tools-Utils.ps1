function Tools-Packup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][String]$AdminName
    )

    $output = $null

    $MYHOME=Get-Content -Path Env:\HOMEPATH

    $icon_folder=Join-Path -Path $MYHOME -ChildPath "Development\icons\darkemblem"
    $tmp_zipfile=Join-Path -Path $MYHOME -ChildPath "Downloads"
    

    Log-Output -result ([ref]$output) `
            -status "NOTIFY" `
            -action "Run Mode" `
            -object "Pack Up" `
            -message ""
    Write-Host $output

    Gen-Icon $icon_folder "velox.ico"

    $icon_file=Join-Path -Path $icon_folder -ChildPath "velox.ico"

    Create-ToolsZip "//wsl.localhost/Ubuntu/home/burtnolej/sambashare/veloxmon/excelvba" `
        @("DV.xlsm","VBAUtils.xlsm","MV.xlsm","MO.xlsm") `
        $tmp_zipfile `
        $icon_file

    $tmp_zipfile=Join-Path -Path $tmp_zipfile -ChildPath "Tools.zip"

    Write-OneDrive $AdminName `
        $tmp_zipfile `
        "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
        "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Deploy" `
        "Tools.zip"
}

function Tools-UnPackup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][String]$AdminName,
        [String]$desktopFolder
    )

    $output = $null

    $MYHOME=Get-Content -Path Env:\HOMEPATH

    $tmp_zipfile=Join-Path -Path $MYHOME -ChildPath "Downloads"
    $tmp_zipfile=Join-Path -Path $tmp_zipfile -ChildPath "Tools.zip"
    $deploy_folder=Join-Path -Path $MYHOME -ChildPath "Deploy"
    $target_app=Join-Path -Path $deploy_folder -ChildPath "vbautils.xlsm"
    $icon_file=Join-Path -Path $deploy_folder -ChildPath "velox.ico"

    $output = $null

    if (([string]::IsNullOrEmpty($desktopFolder))) {
        $desktopFolder=Join-Path -Path $MYHOME -ChildPath "Desktop"
    }

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
