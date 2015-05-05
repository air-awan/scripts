Param(
[string]$serverName, #ADWS hostname
[string]$ADFilter, #AD entry filter
[string]$accountFilter #Account to look for
)
Get-ADComputer -Filter "*" -Server $serverName | 
Where-Object {$_.DistinguishedName -ilike $ADFilter} |
ForEach-Object {
Write-Host -ForegroundColor Cyan $_.name 
Get-WmiObject -Class win32_service -ComputerName $_.name | 
Where-Object {$_.startname -ilike $accountFilter} | 
Format-Table -Property Name,StartName -AutoSize}