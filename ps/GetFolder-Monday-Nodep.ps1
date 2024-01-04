 param (
    [string]$working_dir = "C:\Users\burtn\Development\ps",
    [string]$desktop_folder
 )

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

function Get-OneDriveSubFolders {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$AdminName,
        [Parameter(Mandatory)]
        [String]$SiteUrl,
        [Parameter(Mandatory)]
        [String]$FileFolder
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

    Try {

        #Get the Folder and Files
        $Folder=$Context.Web.GetFolderByServerRelativeUrl($FileFolder)
        $Context.Load($Folder)
        #$Context.Load($Folder.Files)
        $Context.Load($Folder.Folders)
        $Context.ExecuteQuery()
 
        $sb = New-Object -TypeName System.Text.StringBuilder

        #Iterate through each File in the folder
        #Foreach($File in $Folder.Files)
        Foreach($Subfolder in $Folder.Folders)
        {
            #$null = $sb.Append($File.Name)
            $null = $sb.Append($Subfolder.Name + "^")
        }
    }
    Catch{
        Log-Output -result ([ref]$output) `
                -status "ERROR" `
                -action "List Folder" `
                -object $FileFolder `
                -message "Failed" `
                -errormsg  $_
        write-host $output
        exit
    }

    Write-Host $sb
}

$folderstring = Get-OneDriveSubFolders "jon.butler@veloxfintech.com" `
    "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
    "/sites/VeloxSharedDrive/Shared%20Documents/General/Monday"

Write-Host  $folderstring