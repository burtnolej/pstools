
#https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Tools.zip
#curl -o Tools.zip https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/Tools.zip
#. "c:\scratch\b.ps1"
# &"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe" .\96x96.png .\72x72.png .\64x64.png .\48x48.png .\36x36.png .\32x32.png .\24x24.png .\16x16.png velox.ico

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

function Create-ToolsZip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$imagefiles_folder,
        [Parameter(Mandatory)]
        [Array]$imagefilenames,
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
            $input_files_string = $input_files_string +"," +  $input_file 
        }
    }

    # add the icon file
    $input_files_string = $input_files_string +"," + '"' + $icon_file + '"'

    $zipfile=Join-Path -Path $target_zip_dir -ChildPath "Tools.zip"

    # remove the target Zip if it exists
    if (Test-Path -Path $zipfile) {
        Remove-Item  $zipfile
        Log-Output "OK" "Check Old Zip" $zipfile "Removed"
    }

    $compress = @{
        Path = Invoke-Expression $input_files_string
        CompressionLevel = "Fastest"
        DestinationPath = $zipfile
    }
    Compress-Archive @compress

    # Check that the target folder exists
    if (Test-Path -Path $zipfile) {
        Log-Output "OK" "Check Zip File" $zipfile "Created"
    } 
    else {
        Log-Output "ERROR" "Check Zip File" $zipfile "Not Created!"
        exit
    }
}

function Gen-Icon {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$image_folder,
        [Parameter(Mandatory)]
        [String]$icon_file
    )

    $pngs = @("96x96.png","72x72.png","64x64.png","48x48.png","36x36.png","32x32.png","24x24.png","16x16.png")

    $command = New-Object -TypeName System.Text.StringBuilder
    
    $null = $command.Append('"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"')
    $null = $command.Append(' "')

    for ($i=0; $i -lt $pngs.Length; $i++) {
        $png_file = Join-Path -Path $image_folder -ChildPath  $pngs[$i]
        $null = $command.Append($png_file)
        $null = $command.Append('" "')
    }
   
    $icon_file = Join-Path -Path $image_folder -ChildPath  $icon_file

    $null = $command.Append("$icon_file")
    $null = $command.Append('"')

    Invoke-Expression "& $command.ToString()"
}


function Log-Output {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$status,
        [Parameter(Mandatory)]
        [String]$action,
        [Parameter(Mandatory)]
        [String]$object,
        [Parameter(Mandatory)]
        [String]$message,
        [Parameter()]
        [String]$errormsg
    )

    $LOGTIME=Get-Date -Format "MMddyyyy_HHmmss"
    
    $SB = New-Object -TypeName System.Text.StringBuilder
    
    $null = $SB.Append($LOGTIME.PadRight(18," "))
    $null = $SB.Append($status.PadRight(7," "))
    $null = $SB.Append($action.PadRight(20," "))
    $null = $SB.Append($message.PadRight(20," "))
    $null = $SB.Append($object.PadRight(100," "))
    

    if ($PSBoundParameters.ContainsKey('errormsg')) {
        $null = $SB.Append($errormsg)
    }
    $SB.ToString()
}

function Write-OneDrive {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$AdminName,
        [Parameter(Mandatory)]
        [String]$SourceFile,
        [Parameter(Mandatory)]
        [String]$WebUrl,
        [Parameter(Mandatory)]
        [String]$LibraryName
    )

    $CLIENTDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    $CLIENTRUNTIMEDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    if (Test-Path -Path $CLIENTDLL) {
        Log-Output "OK" "Check Onedrive DLLs"  $CLIENTDLL "Found!"
    }
    else {
        Log-Output "ERROR" "Check Onedrive DLLs"  $CLIENTDLL "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
    }

    Add-Type -Path $CLIENTDLL
    Add-Type -Path $CLIENTRUNTIMEDLL

    #$AdminPassword ="4o5yWohgxOB8"

    try {
        $Credential =Get-Credential -Credential $AdminName
    }
    catch {
        Log-Output "ERROR" "Get Credential" $AdminName "Failed" $_
        exit
    }
    
    #Setup Credentials to connect
    $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password)

    #Set up the context
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
    $Context.Credentials = $Credentials

    #Get the Library
    $Library =  $Context.Web.Lists.GetByTitle($LibraryName)

    #Get the file from disk
    $FileStream = ([System.IO.FileInfo] (Get-Item $SourceFile)).OpenRead()

    #Get File Name from source file path
    $SourceFileName = Split-path $SourceFile -leaf

    #sharepoint online upload file powershell
    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $FileCreationInfo.Overwrite = $true
    $FileCreationInfo.ContentStream = $FileStream
    $FileCreationInfo.URL = $SourceFileName
    $FileUploaded = $Library.RootFolder.Files.Add($FileCreationInfo)
    
    #powershell upload single file to sharepoint online
    $Context.Load($FileUploaded)
    $Context.ExecuteQuery()
    
    #Close file stream
    $FileStream.Close()
    
    write-host "File has been uploaded!"
}

function Get-OneDrive{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$AdminName,
        [Parameter(Mandatory)]
        [String]$SiteUrl,
        [Parameter(Mandatory)]
        [String]$FileUrl,
        [Parameter(Mandatory)]
        [String]$TargetFile
    )

    $CLIENTDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    $CLIENTRUNTIMEDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    if (Test-Path -Path $CLIENTDLL) {
        Log-Output "OK" "Check Onedrive DLLs"  $CLIENTDLL "Found!"
    }
    else {
        Log-Output "ERROR" "Check Onedrive DLLs"  $CLIENTDLL "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
    }

    Add-Type -Path $CLIENTDLL
    Add-Type -Path $CLIENTRUNTIMEDLL

    #$AdminPassword ="4o5yWohgxOB8"

    try {
        $Credential =Get-Credential -Credential $AdminName
    }
    catch {
        Log-Output "ERROR" "Get Credential" $AdminName "Failed" $_
        exit
    }
    
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    try {
        $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password)
        $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$FileUrl)
    }
    catch {
        Log-Output "ERROR" "Sharepoint Logon" $AdminName "Failed" $_
        exit
    }

    $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
    $FileInfo.Stream.CopyTo($WriteStream)
    $WriteStream.Close()

    Log-Output "OK" "Download File" $TargetFile "Downloaded"
}