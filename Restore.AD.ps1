<#
A.
Name: Joseph Lepkowski
Student ID: 010128419
#>

# B.1 Check if the 'Finance' OU exists
$OU = Get-ADOrganizationalUnit -Filter {Name -eq 'Finance'} -ErrorAction SilentlyContinue
if ($OU) {
    Write-Host -ForegroundColor red "Finance OU exists, deleting it..."
    Remove-ADOrganizationalUnit -Identity $OU.DistinguishedName -Recursive
    Write-Host -ForegroundColor red "Finance OU deleted."
} else {
    Write-Host -ForegroundColor green "Finance OU does not exist."
}

# B.2 Create the 'Finance' OU
Write-Host "Creating the 'Finance' OU..."
New-ADOrganizationalUnit -Name "Finance" -Path "DC=consultingfirm,DC=com" -DisplayName 'Finance' -ProtectedFromAccidentalDeletion $false
Write-Host -ForegroundColor Green "Finance OU created."


# B.3 Import users from CSV and add them to the Finance OU
$csvPath = Join-Path $PSScriptRoot 'financePersonnel.csv'
$users = Import-csv -Path $csvPath
$Path = "OU=Finance,DC=consultingfirm,DC=com"

foreach ($ADUser in $users) {
    $firstname = $ADUser.First_Name
    $lastname = $ADUser.Last_Name
    $displayname = $firstname + " " + $lastname
    $samAcct = $ADUser.samAccount -replace '[^\w\-]', '' -replace '\.', '' # Remove invalid characters and dots
    $postalcode = $ADUser.PostalCode
    $officephone = $ADUser.OfficePhone 
    $mobilephone = $ADUser.MobilePhone

    if ($displayname.Length -gt 20) {
        $displayname = $displayname.Substring(0, 20)
    } 
    $name = "$firstname $lastname"

    $userParams = @{
        SamAccountName = $samAcct
        GivenName = $firstname
        Surname = $lastname
        DisplayName = $displayname
        Postalcode = $postalcode
        OfficePhone = $officephone
        MobilePhone = $mobilephone
        AccountPassword = (ConvertTo-SecureString 'Passw0rd!' -AsPlainText -Force)
        Enabled = $true 
        Path = $Path
        Name = $name
    }
    try {
        Write-Host -ForegroundColor Cyan "Creating User '$displayname', $firstname', '$lastname', '$postalcode', '$officephone', '$mobilephone'"
        New-ADUser @userParams -ErrorAction Stop
        Write-Host -ForegroundColor Green "User '$displayname' has been created."
    }
    catch {
        Write-Host -ForegroundColor Red "An error occured while creating user '$displayname' : $_"
    }
}

#G.1.III
try {
    $users = Get-ADUser -Filter * -SearchBase 'ou=Finance,dc=consultingfirm,dc=com' -Properties DisplayName, PostalCode, OfficePhone, MobilePhone > .\AdResults.txt
    $adResultsPath = Join-Path $PSScriptRoot 'AdResults.txt'
    
    $users | ForEach-Object {
        $user = $_
        $output = "Display Name: $($user.DisplayName)"
        $output += "`nPostal Code: $($user.PostalCode)"
        $output += "`nOffice Phone: $($user.OfficePhone)"
        $output += "`nMobile Phone: $($user.MobilePhone)"
        $output += "`n"
        
        $output | Out-File -FilePath $adResultsPath -Append -Encoding UTF8
    }

    Write-Host -ForegroundColor Green "The 'AdResults.txt' file has been generated."
}

catch {
    Write-Host -ForegroundColor Red "An error occurred while generating the 'AdResults.txt' file: $_"
}

#B.4.Generate the output file for submission
#$adResultsPath = Join-Path $PSScriptRoot '.\AdResults.txt'



#A.  Create a PowerShell script named “Restore-AD.ps1” within the attached “Requirements2” folder. Create a comment block and include your first and last name along with your student ID.


#B.  Write the PowerShell commands in “Restore-AD.ps1” that perform all the following functions without user interaction:

#1.  Check for the existence of an Active Directory Organizational Unit (OU) named “Finance.” Output a message to the console that indicates if the OU exists or if it does not. If it already exists, delete it and output a message to the console that it was deleted.

#2.  Create an OU named “Finance.” Output a message to the console that it was created.

#3.  Import the financePersonnel.csv file (found in the attached “Requirements2” directory) into your Active Directory domain and directly into the finance OU.
 #Be sure to include the following properties:

#•   First Name

#•   Last Name

#•   Display Name (First Name + Last Name, including a space between)

#•   Postal Code

#•   Office Phone

#•   Mobile Phone

#4.  Include this line at the end of your script to generate an output file for submission:
Write-Host -ForegroundColor Cyan "Generating Output File"
 Get-ADUser -Filter * -SearchBase “ou=Finance,dc=consultingfirm,dc=com” -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt