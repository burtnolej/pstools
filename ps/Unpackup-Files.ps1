# load the utilites script
$WORKINGDIR = Get-Location
$MYHOME=Get-Content -Path Env:\HOMEPATH
$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Tools-utils.ps1"
. $utils_file



function Install-ToolsZip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$ZipFilePath,
        [Parameter(Mandatory)]
        [String]$DestinationPath
    )

    #$ZipFilePath = "C:\Users\burtn\Downloads\Tools.zip"
    #$DestinationPath  = "C:\Users\burtn\Tools\Deploy"
 
    Expand-Archive -LiteralPath $ZipFilePath -DestinationPath $DestinationPath


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
}

Install-ToolsZip "C:\Users\burtn\Downloads\Tools.zip" `
            "C:\Users\burtn\Tools\Deploy"