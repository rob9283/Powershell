$junk = Import-Csv .\junk.csv
Foreach ($item in $junk) {Get-AppxPackage | where {$_.name -like "*"+$item.name+"*"} | Remove-AppxPackage }
Foreach ($item in $junk) {Get-AppxProvisionedPackage -Online | where {$_.packagename -like "*"+$item.name+"*"} | Remove-AppxProvisionedPackage -Online}
