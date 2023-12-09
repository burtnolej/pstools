
$source_folder="C:\Users\burtn\Downloads"
$target_folder="C:\Users\burtn\Downloads\tmp"
$lnk_folder="C:\Users\burtn\Downloads"
$launch_file="foo2.txt"
$icon_file="C:\Users\burtn\Downloads\veloxemblem_icon\velox.ico"

# &"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe" .\96x96.png .\72x72.png .\64x64.png .\48x48.png .\36x36.png .\32x32.png .\24x24.png .\16x16.png velox.ico
#"E:\new_onedrive\Velox Financial Technology\Velox Shared Drive - Documents\General\Digital Assets\Adobe Illustrator\velox icon\wiondows icon copy.ai"

if (Test-Path -Path $target_folder) {
    Write-Output "Path exists!"
} else {
    Write-Output "Path doesn't exist."
    #New-Item -ItemType Directory -Path $target_folder
}


$values = "foo.txt", "foo2.txt"

for ($i=0; $i -lt $values.Length; $i++) {

    $source_file = "$source_folder\" + $values[$i]
    $target_file = "$target_folder\" + $values[$i]

    if (Test-Path -Path $source_file) {
        Write-Output "$source_file File exists!"
    } else {
        Write-Output "$source_file File doesn't exist."
        exit
    }

    try {
        Copy-Item $source_file  -Destination $target_file
    }
    catch {
        Write-Output "Error"
        Write-Output $_
    }
}


$lnk_file = "$lnk_folder\" + "velox.lnk"
$shortcutFile = "$target_folder\" + $launch_file
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$shortcut.TargetPath = $lnk_file
$shortcut.IconLocation = $icon_file
$shortcut.Save()