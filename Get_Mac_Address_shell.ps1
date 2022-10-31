Write-Output "LOADING...`n"

$network = Get-WmiObject win32_networkadapterconfiguration | select description, macaddress | where { $_.description -Like "*ether*" }
$network_mac  = $network.macaddress
$network_card = $network.description

Set-Clipboard -Value $network_mac
Write-Host "CARD: $network_card" -ForegroundColor green
Write-Host "MAC address $network_mac Copied to clipboard" -ForegroundColor green
Write-Host "_________________________________________________"-ForegroundColor yellow
Start-Sleep 4