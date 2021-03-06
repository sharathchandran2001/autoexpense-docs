
# to output the unread emails

Function Global:Get-Email {
Param(
[String]$Folder = "InBox",
[String]$Test ="Unread",
[String]$Compare =$True
    )
Process{
$Folder = $Folder
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")
$NameSpace.Folders.Item(1)
$Email = $NameSpace.Folders.Item(1).Folders.Item($Folder).Items
Clear-Host
Write-Host "Trawling through Outlook, please wait …."
$Email | Where-Object {$_.$Test -match $Compare} | Sort-Object -Property `
@{Expression = "Unread";Descending=$true}, `
@{Expression = "Importance";Descending=$true}, `
@{Expression = "SenderEmailAddress";Descending=$false} -Unique `
| Format-Table Subject, " ", SenderEmailAddress -AutoSize | Out-String
    } # End of main section 'Process'
}

Get-Email  #-Compare false
