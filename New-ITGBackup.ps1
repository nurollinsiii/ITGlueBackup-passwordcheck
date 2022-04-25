#####################################################################
#The unofficial Kelvin Tegelaar IT-Glue Backup script. Run this script whenever you want to create a backup of the ITGlue database.
#Creates a file called "password.html" with all passwords in Plain-text. please only store in secure location.
#Creates folders per organisation, and copies flexible assets there as HTML table  &amp; CSV file for data portability.
$APIKEy = "xxxxxxxx"
$APIEndpoint = "https://api.itglue.com"
$ExportDir = "C:\ImageNet\ITGBackup"
#####################################################################
if (!(Test-Path $ExportDir)) {
    Write-Host "Creating backup directory" -ForegroundColor Green
    new-item $ExportDir -ItemType Directory 
}
#ITGlue Download starts here
If (Get-Module -ListAvailable -Name "ITGlueAPI") { Import-module ITGlueAPI } Else { install-module ITGlueAPI -Force; import-module ITGlueAPI }
#Settings IT-Glue logon information
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $APIKEy
$i = 0
#grabbing all orgs for later use.
do {
    $orgs += (Get-ITGlueOrganizations -page_size 1000 -page_number $i).data
    $i++
    Write-Host "Retrieved $($orgs.count) Organisations" -ForegroundColor Yellow
}while ($orgs.count % 1000 -eq 0 -and $orgs.count -ne 0)
#Grabbing all passwords.
$i = 0
Write-Host "Getting passwords" -ForegroundColor Green
do {
    $i++
    $PasswordList += (Get-ITGluePasswords -page_size 1000 -page_number $i).data
    Write-Host "Retrieved $($PasswordList.count) Passwords" -ForegroundColor Yellow
}while ($PasswordList.count % 1000 -eq 0 -and $PasswordList.count -ne 0)
Write-Host "Processing Passwords. This might take some time." -ForegroundColor Yellow
$Passwords = foreach ($PasswordItem in $passwordlist) {
    (Get-ITGluePasswords -show_password $true -id $PasswordItem.id).data
}
Write-Host "Processed Passwords. Moving on." -ForegroundColor Yellow
Write-Host "Creating backup directory per organisation." -ForegroundColor Green
foreach ($org in $orgs) {
    if (!(Test-Path "$($ExportDir)\$($org.attributes.name)")) { 
        $org.attributes.name = $($org.attributes.name).Replace('\W', " ")
        new-item "$($ExportDir)\$($org.attributes.name)" -ItemType Directory | out-null 
        Write-Host "Creating password file for $($org.attributes.name)" -ForegroundColor Green
        $Passwords.attributes | where-object {$_.'organization-name' -eq $($org.attributes.name) } |select-object 'organization-name', name, username, password, url, created-at, updated-at | export-excel "$($ExportDir)\$($org.attributes.name)\passwords.xlsx"
    }
}
write-host "Exporting all Password to $($ExportDir)\passwords.xlsx"
$Passwords.attributes | select-object 'organization-name', name, username, password, url | export-excel -Path './passwords.xlsx'
