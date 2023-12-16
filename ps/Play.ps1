 # how to pass in param line arguments
 # needs to be placed on the first non comment line
 param (
    [Parameter(Mandatory)][String]$arg1,
    [string]$arg2 = "val2", # sets a default value
    [string]$arg3 = "val1",
    [string]$arg4 # no default value
 )

 # test passed in arguments
 if (([string]::IsNullOrEmpty($arg4))) {
    Write-Host "arg4 is not set"
    $foo = "bar"
}

