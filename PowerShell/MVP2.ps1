[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
$folder.items.count
$inbox_mails = $folder.items
$FilePath = "C:\Users\Public\AutoExpenseSubmission\PS\TEMP"
$OutFileName = "C:\Users\Public\AutoExpenseSubmission\PS\TEMP\abcd.pptx"

#############    Enter Email Subject and SEnder name below  #####################

$Sub= "RE: CGI Testing services overview - for Aparna"  
$Sender = "Yeri Ranganathan, Lakshmi Narasimhan"

#############    Enter Email Subject and SEnder name above  #####################


foreach ($mail in $inbox_mails){
Write-Host $mail.SenderName                      
Write-Host $mail.Subject
}

foreach ($mail in $inbox_mails){
    #Write-Host $mail.Subject

    if( Compare-Object $Sub $mail.Subject )
    {
        #Write-Host 'not found'
    }
    else
   
    
    {
    Write-Host ================Print Subject ===========
    Write-Host $mail.Subject
        if( Compare-Object $Sender $mail.SenderName ){
        
        }
        
        else{
    
        Write-Host ================Print Count ===========
        Write-Host $mail.Attachments.Count
        #Write-Host $mail.Attachments.ToString() 
        $mail.Attachments | foreach {
       Write-Host ================Print File name ===========
        Write-Host $_.filename
         # $_.SaveAsFile((Join-Path $FilePath "abc.pptx"))
         $_.SaveAsFile($OutFileName)
        }
        }

    }
    
    
    
}
