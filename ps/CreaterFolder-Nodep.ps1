 param ($site_parent,$target_folder)

#$utils_file = Join-Path -Path $working_dir -ChildPath "Tools-Utils.ps1"
#. $utils_file

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

function Create-OneDriveFolder{

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$AdminName,
        [Parameter(Mandatory)]
        [String]$WebUrl,
        [Parameter(Mandatory)]
        [String]$LibraryName,
        [Parameter(Mandatory)]
        [String]$TargetFolder
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

    $parentFolder = $Context.Web.GetFolderByServerRelativeUrl($LibraryName)

    $Folder = $parentFolder.Folders.Add($TargetFolder)

    $parentFolder.Context.ExecuteQuery()

    Log-Output -result ([ref]$output) `
            -status "NOTIFY" `
            -action "Status Update" `
            -object $targetfolder `
            -message "Folder has been created"
    #write-host  $output
}


$folderstring = Create-OneDriveFolder "jon.butler@veloxfintech.com" `
    "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
    $site_parent `
    $target_folder

Write-Host  $folderstring