# Gets the path to the user appdata folder
$AppData = (Get-Item env:appdata).value
# This is the default signature folder for Outlook
$localSignatureFolder = $AppData+'\Microsoft\Signatures'
# This is a shared folder on your network where the signature template should be
$templateFilePath = "\\192.168.1.2\it\Template DIR"

#clear current signatures from Machine
#Remove-Item –path $localSignatureFolder –recurse
#New-Item -Path $localSignatureFolder -Name "$AppData\Microsoft\" -ItemType "Signatures"

# Get the current logged in username
$userName = $env:username

# The following 5 lines will query AD and get an ADUser object with all information
$filter = "(&(objectCategory=User)(samAccountName=$userName))"
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.Filter = $filter
$ADUserPath = $searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()

# Now extract all the necessary information for the signature
$name = $ADUser.DisplayName
$email = $ADUser.mail
$job = $ADUser.title
$department = $ADUser.department
$phone = $ADUser.telephonenumber
$office = $ADUser.physicalDeliveryOfficeName
$landline = $ADUser.HomePhone
$fax = $ADUser.facsimileTelephoneNumber

$namePlaceHolder = "%%DisplayName%%"
$emailPlaceHolder = "%%Email%%"
$jobPlaceHolder = "%%JobTitle%%"
$departmentPlaceHolder = "%%Department%%"
$phonePlaceHolder = "%%Mobile%%"
$landlinePlaceHolder = "%%LandLine%%"
$FaxPlaceHolder = "%%Fax%%"


$rawTemplate = get-content $templateFilePath"\Template.htm"

$signature = $rawTemplate -replace $namePlaceHolder,$name
$rawTemplate = $signature

$signature = $rawTemplate -replace $emailPlaceHolder,$email
$rawTemplate = $signature

$signature = $rawTemplate -replace $phonePlaceHolder,$phone
$rawTemplate = $signature

$signature = $rawTemplate -replace $jobPlaceHolder,$job
$rawTemplate = $signature

$signature = $rawTemplate -replace $landlinePlaceHolder,$landline
$rawTemplate = $signature

$signature = $rawTemplate -replace $FaxPlaceHolder,$Fax
$rawTemplate = $signature

$signature = $rawTemplate -replace $departmentPlaceHolder,$department

# Save it as <username>.htm
$fileName = $localSignatureFolder + "\" + "Template" + ".htm"

# Gets the last update time of the template.
if(test-path $templateFilePath){
    $templateLastModifiedDate = [datetime](Get-ItemProperty -Path $templateFilePath -Name LastWriteTime).lastwritetime
}

# Checks if there is a signature and its last update time
#if(test-path $filename){
    #$signatureLastModifiedDate = [datetime](Get-ItemProperty -Path $filename -Name LastWriteTime).lastwritetime
   # if((get-date $templateLastModifiedDate) -gt (get-date $signatureLastModifiedDate)){
    #    $signature > $fileName
  #  }
#}else{
    $signature > $fileName
#}

Remove-Item -Recurse -Force "$AppData\Microsoft\Signatures\Template_files\"

Copy-Item "$templateFilePath\TSL Beira Template_files\" -Destination "$AppData\Microsoft\Signatures\Template_files\" -recurse

#Convert TO RTF
$wrd = new-object -com word.application 
 
# Make Word Visible 
$wrd.visible = $true 
 
# Open a document  
$doc = $wrd.documents.open("$AppData\Microsoft\Signatures\Template.htm") 

# Save as rtf
$opt = 6
$name = "$AppData\Microsoft\Signatures\Template.rtf"
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)

# Close and go home
$wrd.Quit()

    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value "Template" -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings') { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value "Template" -PropertyType 'String' -Force
    } 