 param (
    [Parameter(Mandatory)][string]$source_file
 )

function Upload-MMReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$source_file
    )
    
    Install-SharepointDLL

    $output = $null

    $MYHOME=Get-Content -Path Env:\HOMEPATH

    $target_file=($source_file).Split("\")[-1]

    #Write-Host $source_file
    #Write-Host $target_file
    Write-OneDrive "jon.butler@veloxfintech.com" `
        $source_file `
        "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
        "/sites/VeloxSharedDrive/Shared%20Documents/General/Sales Team/Meeting Summaries" `
        $target_file

    exit
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
        Log-Output -result ([ref]$output) `
                -status "OK" `
                -action "Check Onedrive DLLs" `
                -object $CLIENTDLL `
                -message "Found!"
        write-host  $output
        $output=$null
    }
    else {
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "Check Onedrive DLLs" `
                -object $CLIENTDLL `
                -message "Install : https://www.microsoft.com/en-us/download/details.aspx?id=42038"
        write-host  $output
        $output=$null
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
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "Get Credential" `
                -object $AdminName `
                -message "Failed" `
                -errormsg  $_
        write-host $output
        $output=$null
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
 
    Log-Output -result ([ref]$output) `
            -status "NOTIFY" `
            -action "Status Update" `
            -object "foo" `
            -message "File has been uploaded"
    write-host  $output
    $output=$null

}

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

Upload-MMReport $source_file
