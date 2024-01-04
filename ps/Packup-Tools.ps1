﻿ param (
    [string]$working_dir = "C:\Users\burtn\Development\ps",
    [string]$desktop_folder
 )


$utils_file = Join-Path -Path $working_dir -ChildPath "Tools-Utils.ps1"
$deploy_tools_file = Join-Path -Path $working_dir -ChildPath "Deploy-Tools-Utils.ps1"
. $utils_file
. $deploy_tools_file


Set-Location -Path $working_dir
Read-Host "Press any key to continue ........ Starting : $(Get-ScriptName)"

Tools-Packup -AdminName "jon.butler@veloxfintech.com"

Read-Host "Press any key to exit ........ Completed : $(Get-ScriptName)"