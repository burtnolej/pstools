param (
    [string]$cwd = "C:\Users\burtn"
)

#Start-Process -FilePath "C:\Program Files\Git\git-bash.exe" -ArgumentList "--cd=C:\"
Start-Process -FilePath "C:\Program Files\Git\git-bash.exe" -ArgumentList "--cd=$cwd"