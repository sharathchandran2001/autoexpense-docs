
# To find email based on subject

Function Global:Get-Email {
Param(
[String]$Folder = "InBox",
[String]$Test ="Read",
#[String]$Test ="Unread",
[String]$Subject ='RE: Feasibility Test on your machine',
[String]$Actual =$True,
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
#$Email | Where-Object {$_.$Test -match $Compare} | Sort-Object -Property `
$Email | Where-Object {$_.Subject -match $Subject} | Sort-Object -Property `
@{Expression = "Unread";Descending=$true}, `
@{Expression = "Importance";Descending=$true}, `
@{Expression = "SenderEmailAddress";Descending=$false} -Unique `
| Format-Table Subject, " ", SenderEmailAddress -AutoSize | Out-String
    } # End of main section 'Process'
}

Get-Email  #- check outpuut
