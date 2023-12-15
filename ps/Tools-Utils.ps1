
function Install-ToolsZip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$ZipFilePath,
        [Parameter(Mandatory)]
        [String]$DestinationPath
    )
 
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

    
    # &"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe" .\96x96.png .\72x72.png .\64x64.png .\48x48.png .\36x36.png .\32x32.png .\24x24.png .\16x16.png velox.ico
}

function Log-Output {
    [CmdletBinding()]
    param (
        [ref]$result,
        [String]$status,
        [String]$action,
        [String]$object,
        [String]$message,
        [String]$errormsg
    )

    $LOGTIME=Get-Date -Format "MMddyyyy_HHmmss"
    
    $sb = New-Object -TypeName System.Text.StringBuilder
    
    $null = $sb.Append($LOGTIME.PadRight(18," "))
    $null = $sb.Append($status.PadRight(7," "))
    $null = $sb.Append($action.PadRight(20," "))
    $null = $sb.Append($message.PadRight(20," "))
    $null = $sb.Append($object.PadRight(100," "))
    

    if ($PSBoundParameters.ContainsKey('errormsg')) {
        $null = $result.Append($errormsg)
    }
    $result.value = $sb.ToString()
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
        [String]$LibraryName,
        [Parameter(Mandatory)]
        [String]$TargetFile
    )

    $output=$null

    $CLIENTDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    $CLIENTRUNTIMEDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    
    if (Test-Path -Path $CLIENTDLL) {
        Log-Output -result ([ref]$output) -status "OK" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Found!"
        write-host  $output
    }
    else {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
        write-host  $output
    }

    Add-Type -Path $CLIENTDLL
    Add-Type -Path $CLIENTRUNTIMEDLL

    $AdminPassword ="4o5yWohgxOB8"

    $SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force

    try {
        if (-not ([string]::IsNullOrEmpty($AdminPassword)))
        {
            $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$SecurePassword)
        }
        else {
            $Credential =Get-Credential -Credential $AdminName
            #Setup Credentials to connect
            $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password) 
        }
    }       
    catch {
            Log-Output -result ([ref]$output) -status "ERROR" -action "Get Credential" -object $AdminName -message "Failed" -errormsg  $_
            write-host $output
            exit
    }

    

    #Set up the context
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
    $Context.Credentials = $Credentials

    $targetFolder = $Context.Web.GetFolderByServerRelativeUrl($LibraryName)

    #Get the file from disk
    $FileStream = ([System.IO.FileInfo] (Get-Item $SourceFile)).OpenRead()

    #sharepoint online upload file powershell
    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $FileCreationInfo.Overwrite = $true
    $FileCreationInfo.ContentStream = $FileStream
    $FileCreationInfo.URL = $TargetFile
    #$FileUploaded = $Library.RootFolder.Files.Add($FileCreationInfo)
    $FileUploaded = $targetFolder.Files.Add($FileCreationInfo)
    
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
        [String]$FileFolder,
        [Parameter(Mandatory)]
        [String]$TargetFile
    )

    $output=$null

    $CLIENTDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    $CLIENTRUNTIMEDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    if (Test-Path -Path $CLIENTDLL) {
        Log-Output -result ([ref]$output) -status "OK" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Found!"
        write-host  $output
    }
    else {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
        write-host  $output
    }

    Add-Type -Path $CLIENTDLL
    Add-Type -Path $CLIENTRUNTIMEDLL

    $AdminPassword ="4o5yWohgxOB8"

    $SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)

    try {
        if (-not ([string]::IsNullOrEmpty($AdminPassword)))
        {
            $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$SecurePassword)
        }
        else {
            $Credential =Get-Credential -Credential $AdminName
            $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password) 
        }
    }
    catch {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Get Credential" -object $AdminName -message "Failed" -errormsg  $_
        Write-host $output
        exit
    }
    
    $FileUrl = $FileFolder + "/" + $FileUrl

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    try {
        #$Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password)
        $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$SecurePassword)
        $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$FileUrl)
    }
    catch {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Sharepoint Logon" -object $AdminName -message "Failed" -errormsg  $_
        write-host $output
        exit
    }

    $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
    $FileInfo.Stream.CopyTo($WriteStream)
    $WriteStream.Close()

    Log-Output -result ([ref]$output) -status "OK" -action "Download File" -object $TargetFile -message "Downloaded"
    write-Output $output
}

function Move-NewJobTitlesCsv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$AdminName,
        [Parameter(Mandatory)]
        [String]$SiteUrl,
        [Parameter(Mandatory)]
        [String]$FolderUrl,
        [Parameter(Mandatory)]
        [String]$TargetFolderUrl  
    )

    $output=$null

    $CLIENTDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    $CLIENTRUNTIMEDLL="C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    if (Test-Path -Path $CLIENTDLL) {
        Log-Output -result ([ref]$output) -status "OK" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Found!"
        Write-host $output
    }
    else {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Check Onedrive DLLs" -object $CLIENTDLL -message "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
        Write-host $output
    }

    Add-Type -Path $CLIENTDLL
    Add-Type -Path $CLIENTRUNTIMEDLL

    $AdminPassword ="4o5yWohgxOB8"
    $SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force

    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)

    try {
        if (-not ([string]::IsNullOrEmpty($AdminPassword)))
        {
            $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$SecurePassword)
        }
        else {
            $Credential =Get-Credential -Credential $AdminName
            $Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminName,$Credential.Password) 
        }
    }
    catch {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Get Credential" -object $AdminName -message "Failed" -errormsg  $_
        Write-host $output
        exit
    }
    
    try {

        #Get the Folder and Files
        $Folder=$Context.Web.GetFolderByServerRelativeUrl($FolderUrl)
        $Context.Load($Folder)
        $Context.Load($Folder.Files)
        $Context.ExecuteQuery()
 
        #Iterate through each File in the folder
        Foreach($File in $Folder.Files)
        {
            #Write-Host $File.Name
            #Write-Host $File.Name.gettype()
            
            $TargetFileUrl = Join-Path -Path $TargetFolderUrl -ChildPath  $File.Name
            $File.MoveTo($TargetFileUrl, [Microsoft.SharePoint.Client.MoveOperations]::Overwrite)
            $Context.ExecuteQuery()
        }
    }
    catch {
        Log-Output -result ([ref]$output) -status "ERROR" -action "Sharepoint Logon" -object $AdminName -message "Failed" -errormsg  $_
        Write-host $output
        exit
    }

    return $File.Name
}

function Run-PythonJobParser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$jobrules_file,
        [Parameter(Mandatory)]
        [String]$person_file,
        [Parameter(Mandatory)]
        [String]$organisation_file,
        [Parameter(Mandatory)]
        [String]$output_file,
        [Parameter(Mandatory)]
        [String]$debugflag,
        [Parameter(Mandatory)]
        [String]$delimiter
    )


    $command = New-Object -TypeName System.Text.StringBuilder
    
    #$null = $command.Append('"')
    $null = $command.Append('C:\Users\burtn\AppData\Local\Microsoft\WindowsApps\python3.11.exe')
    $null = $command.Append(' C:\Users\burtn\Development\py\capsule_parse_jobtitle.py ')
    $null = $command.Append("rulesfile=$jobrules_file")
    $null = $command.Append(' ')
    $null = $command.Append("personsfile=$person_file")
    $null = $command.Append(' ')
    $null = $command.Append("clientsfile=$organisation_file")
    $null = $command.Append(' ')
    $null = $command.Append("outputfile=$output_file")
    $null = $command.Append(' ')
    $null = $command.Append("debug=$debugflag")
    $null = $command.Append(' ')
    $null = $command.Append("delimiter='"+$delimiter+"'")
    #$null = $command.Append('"')


    Write-Output $command.ToString()
    Invoke-Expression "& $command"
    #Invoke-Expression $var -OutVariable | Tee-Object -Variable $out
}


