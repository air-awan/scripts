#Reference: http://www.phy2vir.com/issues-with-dragging-items-in-some-folders-from-outlook-after-migrating-to-office365/
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$Username="admin@domain.com" #email address of the affected user
$Password="Password" #password
$Mailbox="user@domain.com"
 
#$service contains all the details to create a connection to the mailbox
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService 
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
$service.AutodiscoverUrl($Mailbox,{$true})
 
#This section provides all the details of the folder to be worked on within the mailbox
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$Mailbox) #or ::Root,$Mailbox)
$froot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid) 
$fview = new-object Microsoft.Exchange.WebServices.Data.FolderView(1000) #get the first 1000 records
$fview.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
 
#The processing begins
$i = 1
$j = 0
do {
 $folders = $froot.FindFolders($fView) # $folders contains the folders that are present in $view
 ForEach ($f in $folders.Folders) { # loop through each folder in the root folder specified in $froot
 Write-host -NoNewLine $i, $f.displayname #Output the number and the name of the folder that is being processed
 Write-Host
 $i++
 if (-not $f.folderclass) { #if the FolderClass Property is not set,
 Write-Host
 $f
 $r = Read-Host "Update" $f.displayname " (y/n) ?" #a prompt appears asking if this FolderClass Property should be updated
 if ($r -eq "y") { #if the answer is y, the property will be change to IPF.Note
 $f.folderclass = "IPF.Note" 
 $j++
 $f.Update() #apply the change
 }
 } 
 }
 $fview.Offset += 1000 #get the next 1000 records
} while ($folders.moreAvailable) #the process will repeat while there are more folders to process.
Write-Host -NoNewLine $i, "objects." #this line will output the total number of folders processed.
Write-Host
Write-Host -NoNewLine $j, "updated objects." #this line will output the total number of "FolderClass" properties updated.
Write-Host