param (
    [string]$desktop = 3,
    [Parameter(Mandatory=$true)][string]$filepath
)
 
#Get-DesktopName


# how to run windows store apps from the command line
https://www.auslogics.com/en/articles/how-to-open-microsoft-store-apps-from-command-prompt/

Get-Desktop $desktop | Switch-Desktop
Start-Process -FilePath $filepath


#PS C:\Users\burtn> Get-DesktopName
#Desktop 2
#S C:\Users\burtn> Get-DesktopName | Switch-Desktop
#PS C:\Users\burtn> 3 | Switch-Desktop
#PS C:\Users\burtn> 1 | Switch-Desktop
#PS C:\Users\burtn> 2 | Switch-Desktop