
# load the utilites script
$WORKINGDIR = Get-Location
$MYHOME=Get-Content -Path Env:\HOMEPATH
$RUNTIME=Get-Date -Format "MMddyyyy_HHmmss"

$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Tools-utils.ps1"
. $utils_file


$source_folder=Join-Path -Path $MYHOME -ChildPath "Tools\excelvba"
$target_folder=Join-Path -Path $MYHOME -ChildPath "Tools\Deploy"
$lnk_folder = "E:\Velox Financial Technology\OneDrive - Velox Financial Technology\Desktop\"
$launch_file=Join-Path -Path $target_folder -ChildPath "vbautils.xlsm"
$icon_file=Join-Path -Path $MYHOME -ChildPath "Downloads\veloxemblem_icon\velox.ico"


# &"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe" .\96x96.png .\72x72.png .\64x64.png .\48x48.png .\36x36.png .\32x32.png .\24x24.png .\16x16.png velox.ico
#"E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Digital Assets\Adobe Illustrator\velox icon\wiondows icon copy.ai"
$values = @("DV.xlsm","vbautils.xlsm","MV.xlsm","MO.xlsm")
#$values = @("DV.xlsm")


two scripts pack up and unpack

#$compress = @{
#    Path = "C:\Reference\Draftdoc.docx", "C:\Reference\Images\*.vsd"
#    CompressionLevel = "Fastest"
#   DestinationPath = "C:\Archives\Draft.zip"
#  }
#  Compress-Archive @compress


for ($i=0; $i -lt $values.Length; $i++) {

    #$source_file = "$source_folder\" + $values[$i]

    $source_file = "$source_folder\" + $values[$i] + "\" + $values[$i]
    $target_file = "$target_folder\" + $values[$i]


    # Check that the source XLSM exists
    if (Test-Path -Path $source_file) {
        Log-Output "ERROR" "Check Source File" $source_file "File exists!"
    } else {
        Log-Output "OK" "Check Source File" $source_file "File doesn't exist"
        exit
    }

    # Check that the target folder exists
    if (Test-Path -Path $target_folder) {
        Log-Output "ERROR" "Check target Folder" $target_folder "Folder exists!"
    } else {
        Log-Output "OK" "Check target Folder" $target_folder "Folder doesn't exist"
        exit
    }

    # Check whether the the target file already exists
    if (Test-Path -Path $target_file) {
        Log-Output "ERROR" "Check target File" $target_file "file exists!"
        $backup_file = $target_file +"." + $RUNTIME
        Move-Item $target_file -Destination $backup_file
        Log-Output "OK" "Create Backup" $backup_file "Created"
    } else {
        Log-Output "OK" "Check target File" $target_file "Folder doesn't exist"
    }

    try {
        Copy-Item $source_file  -Destination $target_file
        Log-Output "OK" "Copy File" $target_file "Copied"
    }
    catch {
        Log-Output "ERROR" "Copy File" $source_file "Failed", $_
        exit
    }
}


# Create shortcut file and put on the desktop
$shortcutFile = "$lnk_folder\" + "velox.lnk"

try {
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $shortcut.TargetPath = $launch_file
    $shortcut.IconLocation = $icon_file
    $shortcut.Save()
}
catch {
    Log-Output "ERROR" "Create Shortcut" $ShortcutFile "Failed", $_
    exit
}
Log-Output "OK" "Create Shortcut" $ShortcutFile "Created"


