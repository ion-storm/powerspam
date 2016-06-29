<# simply place Outlook .msg Email files in the same directory as this then run to grab the sender ip  
#>
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
Get-ChildItem -Filter *.msg|`
ForEach-Object {
$outlook = New-Object -comobject outlook.application
$msg = $outlook.CreateItemFromTemplate($_.FullName)
$headers = ''
$headers = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x007D001E')
$headers | findstr "From:" | findstr /V "Envelope"
$headers | findstr /r "[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*" | Select-String -Pattern 'ESMTP id'
$headers | findstr /r "[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*" | Select-String -Pattern 'Apparent-Source-IP'
<# May need in the future
$outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($msg) | Out-Null
Remove-Variable outlook | Out-Null
Remove-Variable msg | Out-Null #>
}
