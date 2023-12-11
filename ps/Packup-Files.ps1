

# load the utilites script
$WORKINGDIR = Get-Location
$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Tools-utils.ps1"
. $utils_file

function Create-ToolsZip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$imagefiles_folder,
        [Parameter(Mandatory)]
        [String]$imagefilenames,
        [Parameter(Mandatory)]
        [String]$target_zip_dir,
        [Parameter(Mandatory)]
        [String]$icon_file  
    )

    $input_files_string=""

    for ($i=0; $i -lt $imagefilenames.Length; $i++) {

        # xlsm lives in a folder of the same name
        $input_file =  $imagefilenames[$i] + "\" + $imagefilenames[$i]
        $input_file = Join-Path -Path $imagefiles_folder -ChildPath  $input_file
        $input_file = '"' + $input_file + '"'

        if ($input_files_string -eq "") {
            $input_files_string = $input_file
        }
        else {
            $input_files_string = $input_files_string +"," + $input_file
        }
    }

    # add the icon file
    $input_files_string = $input_files_string +"," + $icon_file

    $zipfile=Join-Path -Path $target_zip_dir -ChildPath "Tools.zip"

    # remove the target Zip if it exists
    if (Test-Path -Path $zipfile) {
        Remove-Item  $zipfile
        Log-Output "OK" "Check Old Zip" $zipfile "Removed"
    } else {
        Log-Output "ERROR" "Check Zip File" $zipfile "Not Created!"
        exit
    }

    $compress = @{
        #Path = $input_files_string
        #Path = "\\wsl.localhost\Ubuntu\home\burtnolej\sambashare\veloxmon\excelvba\DV.xlsm\DV.xlsm","\\wsl.localhost\Ubuntu\home\burtnolej\sambashare\veloxmon\excelvba\vbautils.xlsm\vbautils.xlsm","\\wsl.localhost\Ubuntu\home\burtnolej\sambashare\veloxmon\excelvba\MV.xlsm\MV.xlsm","\\wsl.localhost\Ubuntu\home\burtnolej\sambashare\veloxmon\excelvba\MO.xlsm\MO.xlsm"
        Path = Invoke-Expression $input_files_string
        CompressionLevel = "Fastest"
        DestinationPath = $zipfile
    }
    Compress-Archive @compress

    # Check that the target folder exists
    if (Test-Path -Path $zipfile) {
        Log-Output "OK" "Check Zip File" $zipfile "Created"
    } else {
        Log-Output "ERROR" "Check Zip File" $zipfile "Not Created!"
        exit
    }
}

