#####################################################################
#PLEASE RUN THIS SECONDARILY TO THE ITGLUEBACKUP PoSH SCRIPT.
#!!!!REMEMBER TO CHANGE PATH TO YOUR ROOT DIRECTORY 
#####################################################################
$report = [System.Collections.ArrayList]@()
$files = Get-ChildItem *.xlsx -path "c:\PathToYourRootDirectory\*" -Recurse
ForEach ($file in $files) {
    $test = import-excel $file
    $regex = "[^a-zA-Z0-9]"
    $regex1 = "[a-z]"
    $regex2 = "[A-Z]"
    $regex3 = "[0-9]"
    $test | % {
        $objDat = [PSCustomObject]@{
            organization      = $_."organization-name"
            Name              = $_.name
            username          = $_.username
            url               = $_.url
            password1         = $_.password.length
            Password          = $_.password
            SpecialCharacters = $_.password -cmatch $regex
            LowerCase         = $_.password –cmatch $regex1
            UpperCase         = $_.password –cmatch $regex2
            Numbers           = $_.password –cmatch $regex3
        }    
        $report.add($objDat) | Out-Null
    }
}
$report | Export-excel -path "c:\Users\nealzo\Downloads\PasswordSample\Final_exported_list_$(Get-Date -f yyyy_MM_dd).xlsx"
