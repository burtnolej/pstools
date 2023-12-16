
$working_dir = "C:\Users\burtn\Development\ps"


$utils_file = Join-Path -Path $working_dir -ChildPath "Tools-Utils.ps1"
$deploy_tools_file = Join-Path -Path $working_dir -ChildPath "Deploy-Tools-Utils.ps1"
. $utils_file
. $deploy_tools_file

Deploy-ToolsExe