## Console output encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Include Functions
. $PSScriptRoot\RealTime-Func.ps1
. $PSScriptRoot\Angel-Query.ps1


$res = get-case("07643728")
write-host $res