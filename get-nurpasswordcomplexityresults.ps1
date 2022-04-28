#####################################################################
#PLEASE RUN THIS SECONDARILY TO THE ITGLUEBACKUP PoSH SCRIPT.
#!!!!REMEMBER TO CHANGE PATH TO YOUR ROOT DIRECTORY 
#####################################################################
# Report array
$report = [System.Collections.ArrayList]@()
# Reference list of bad passwords
$badpasswords = Get-Content -path "C:\etc\file.txt"
#Root Directory of Companies with respective xlsx files
$files = Get-ChildItem *.xlsx -path "c:\etc\PasswordSample\*" -Recurse
ForEach ($file in $files) {
    $test = import-excel $file
    $regex = "[^a-zA-Z0-9]"
    $regex1 = "[a-z]"
    $regex2 = "[A-Z]"
    $regex3 = "[0-9]"
    #wcpw = wildcard password. This is going to match things like "*password*". Anything before or after you will get a hit on.
    $wcpw1  = "admin"
    $wcpw2  = '@dmin'
    $wcpw3  = "pas$$word"
    $wcpw4  = "password"
    $test | % {
        $objDat = [PSCustomObject]@{
            organization      = $_."organization-name"
            Name              = $_.name
            username          = $_.username
            url               = $_.url
            PasswordLength    = $_.password.length
            Password          = $_.password
            SpecialCharacters = $_.password -cmatch $regex
            LowerCase         = $_.password –cmatch $regex1
            UpperCase         = $_.password –cmatch $regex2
            Numbers           = $_.password –cmatch $regex3
            BadPassword       = $badpasswords -contains $_.password
            wcpw1             = $_.password -match $wcpw1
            wcpw2             = $_.password -match $wcpw2
            wcpw3             = $_.password -match $wcpw3
            wcpw4             = $_.password -match $wcpw4

        }    
        $report.add($objDat) | Out-Null
    }
}
$report | Export-excel -path "c:\etc\PasswordSample\Final_exported_list_$(Get-Date -f yyyy_MM_dd).xlsx"
