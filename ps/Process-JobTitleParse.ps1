
# load the utilites script
$WORKINGDIR = Get-Location
$MYHOME=Get-Content -Path Env:\HOMEPATH
$utils_file = Join-Path -Path $WORKINGDIR -ChildPath ".\Tools-utils.ps1"
#$utils_file = Join-Path -Path $WORKINGDIR -ChildPath "Development\ps\Tools-utils.ps1"
#. $utils_file 
.  "./Tools-utils.ps1"

$output=$null
$file_to_process

for ($i=1; $i=2; $i++)
{
    $file_to_process = Move-NewJobTitlesCsv "jon.butler@veloxfintech.com" `
            "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
            "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/datafiles" `
            "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/datafiles/processing"

    if ($file_to_process -eq $null) {
        Log-Output -result ([ref]$output) -status "NOTIFY" -action "Check for file" -object $file_to_process -message "Failed"
        write-host $output
        Start-Sleep -Seconds 5
    }
    else {
        $fileparent=($file_to_process).Split(".")

        $LOGTIME=Get-Date -Format "MMddyyyy_HHmmss"
    
        $outputfile = New-Object -TypeName System.Text.StringBuilder
        $null = $outputfile.Append($fileparent[0])
        $null = $outputfile.Append("_")
        $null = $outputfile.Append($LOGTIME)
        $null = $outputfile.Append(".csv")

        write-host $file_to_process

        Get-OneDrive "jon.butler@veloxfintech.com" `
                "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
                $file_to_process `
                "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/datafiles/processing" `
                "C:\Users\burtn\Development\csv\jobtitlerules_new.csv"

        Start-Sleep -Seconds 5

        Run-PythonJobParser "C:\Users\burtn\Development\csv\jobtitlerules_new.csv" `
                \\wsl.localhost\Ubuntu\var\www\html\datafiles\person.csv `
                \\wsl.localhost\Ubuntu\var\www\html\datafiles\organisation.csv `
                output.txt `
                "True" `
                ","

        Start-Sleep -Seconds 3

        write-host $outputfile
        Write-OneDrive "jon.butler@veloxfintech.com" `
               output.txt `
               "https://veloxfintechcom.sharepoint.com/sites/VeloxSharedDrive" `
               "/sites/VeloxSharedDrive/Shared%20Documents/General/Tools/datafiles/results" `
               $outputfile

        Start-Sleep -Seconds 5

        # need to add a move to processed
        # need to remove the interim files 
        # "C:\Users\burtn\Development\csv\jobtitlerules_new.csv" and
        # output.txt
    }
}