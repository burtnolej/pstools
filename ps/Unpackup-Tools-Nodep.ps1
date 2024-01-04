

Read-Host "Installing ..... Press RETURN to continue"

#$utils_file = Join-Path -Path $working_dir -ChildPath "Tools-Utils.ps1"
#$deploy_tools_file = Join-Path -Path $working_dir -ChildPath "Deploy-Tools-Utils.ps1"
#. $utils_file
#. $deploy_tools_file

function Install-SharepointDLL {

    $output=$null

    #Parameters
    $DownloadURL = "https://download.microsoft.com/download/B/3/D/B3DA6839-B852-41B3-A9DF-0AFA926242F2/sharepointclientcomponents_16-6906-1200_x64-en-us.msi"
    $Assemblies= @(
            "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll",
            "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
        )
 
    #Check if all assemblies given in the list are found
    $FileExist = $True
    ForEach ($File in $Assemblies)
    {
        #Check if CSOM Assemblies are Found
        If(!(Test-Path $File))
        {
            $FileExist = $False; Break;
        }
    }
 
    #Download and Install CSOM Assemblies
    If(!$FileExist)
    {
        #Download the SharePoint Online Client SDK
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "Check for SP DLL's" `
                -object "" `
                -message "Missing, Downloading ..."
        Write-Host $output

        $InstallerPath = "$Env:TEMP\SharePointOnlineClientComponents16.msi"
        Invoke-WebRequest $DownloadURL -OutFile $InstallerPath

     
        #Start Installation
        Start-Process MSIExec.exe -ArgumentList "/i $InstallerPath /qb" -Wait

        Log-Output -result ([ref]$output) `
                -status "NOTIFY" `
                -action "Installing SP DLL's" `
                -object "" `
                -message "Done ..."
        Write-Host $output

    }
    Else
    {
        Log-Output -result ([ref]$output) `
                -status "NOTIFY" `
                -action "Check for SP DLL's" `
                -object "" `
                -message "Already Installed"
        Write-Host $output
    }

    #Read more: https://www.sharepointdiary.com/2018/12/download-install-sharepoint-online-client-side-sdk-using-powershell.html#ixzz8M9JYE2ci
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
    $null = $sb.Append($action.PadRight(25," "))
    $null = $sb.Append($message.PadRight(50," "))
    $null = $sb.Append($object.PadRight(100," "))
    

    if ($PSBoundParameters.ContainsKey('errormsg')) {
        $null = $sb.Append($errormsg)
    }
    $result.value = $sb.ToString()
}

function Tools-UnPackup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][String]$AdminName,
        [String]$desktopFolder
    )

    Install-SharepointDLL

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

    if ($AdminName -eq "jon.butler@veloxfintech.com") {
        $AdminPassword ="4o5yWohgxOB8"
        $SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
    }

    
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

    try {
        $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$FileUrl)
    }
    catch {
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "Sharepoint Logon" `
                -object $AdminName `
                -message "Failed" `
                -errormsg  $_
        write-host $output
        exit
    }

    $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
    $FileInfo.Stream.CopyTo($WriteStream)
    $WriteStream.Close()

    Log-Output -result ([ref]$output) -status "OK" -action "Download File" -object $TargetFile -message "Downloaded"
    write-Output $output
}

function Install-ToolsZip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$ZipFilePath,
        [Parameter(Mandatory)]
        [String]$DestinationPath,
        [Parameter(Mandatory)]
        [String]$LinkFolder,
        [Parameter(Mandatory)]
        [String]$TargetPath,
        [Parameter(Mandatory)]
        [String]$IconLocation
    )
 
    $output=$null

    if (Test-Path -Path $DestinationPath) {
        Remove-Item -Path $DestinationPath -Recurse -Force
        Log-Output -result ([ref]$output) `
                -status "OK" `
                -action "Check Deploy folder" `
                -object $zipfile `
                -message "Exists/Removed"
        Write-Host $output
        $output=$null
    }

    if (Test-Path -Path $DestinationPath) {
    }
    else {
        $Folder = New-Item -ItemType Directory -Path $DestinationPath
    }

    $output=$null

    Expand-Archive -LiteralPath $ZipFilePath -DestinationPath $DestinationPath
    Log-Output -result ([ref]$output) `
            -status "OK" `
            -action "Archive Expanded" `
            -object $ZipFilePath `
            -message $DestinationPath
    Write-Host $output

    # Create shortcut file and put on the desktop
    $shortcutFile = "$LinkFolder\" + "velox.lnk"

    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $shortcut.TargetPath = $TargetPath
        $shortcut.IconLocation = $IconLocation
        $shortcut.Save()

        Log-Output -result ([ref]$output) `
                -status "OK" `
                -action "Create Shortcut" `
                -object $shortcutFile `
                -message "Created"
        Write-Host $output

    }
    catch {
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "Create Shortcut" `
                -object "foo" `
                -message "Failed" `
                -errormsg $_
        Write-Host $output
        $output=$null
    }
}

Tools-UnPackup -AdminName "jon.butler@veloxfintech.com" $desktop_folder

Read-Host "Completed ..... Press RETURN to continue"