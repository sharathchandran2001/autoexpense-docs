﻿#https://stackoverflow.com/questions/4037939/powershell-says-execution-of-scripts-is-disabled-on-this-system
#set-executionpolicy unrestricted
#$inbox = Get-OutlookInBox
#$inbox | Group-Object -Property senderName -NoElement | Sort-Object count

Get-OutlookInBox function

Function Get-OutlookInBox

{

 Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null

 $olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]

 $outlook = new-object -comobject outlook.application

 $namespace = $outlook.GetNameSpace(“MAPI”)

 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)

 $folder.items |

 Select-Object -Property Subject, ReceivedTime, Importance, SenderName

} #end function Get-OutlookInbox
