
# https://www.fakenamegenerator.com/advanced.php

Set-Date -Date (Get-Date).AddDays(-63)
$data = import-csv  .\data.csv | sort {Get-Random} | select -first 100
$log = 0

$segments_list = @"
country;company;upn_suffix;ou;short
FR;Primeo-Energie France SAS;primeo-energie.fr;OU=Primeo-Energie France,OU=FR,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;PE
FR;RCUE;rc-ue.fr;OU=Others,OU=FR,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;RCUE
FR;Aventron;aventron.com;OU=Aventron,OU=FR,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;AVENTRON
CH;Primeo-Energie AG;primeo-energie.ch;OU=Primeo-Energie AG,OU=CH,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;ENE
CH;Wärme;primeo-energie.ch;OU=Wärme,OU=CH,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;WAR
CH;Netz AG;primeo-energie.ch;OU=Netz AG,OU=CH,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;NET
CH;Cosmos;cosmos.ch;OU=Others,OU=CH,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;COS
CH;Energie AG;primeo-energie.ch;OU=Energie AG,OU=CH,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;ENRG
DE;Partner DE 1;par-de.de;OU=Partner DE 1,OU=DE,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;PDE
DE;Brucker AG;brucker.de;OU=Others,OU=DE,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;BRU
DE;Salt Germany;salt.de;OU=Others,OU=DE,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;SALT
ES;Partner ES;par-es.es;OU=Others,OU=ES,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;PES
ES;Espejo Mágico;espejo-magico.es;OU=Others,OU=ES,OU=1_provisioned,OU=10_users,OU=primeo-energie,DC=pe,DC=ch;ESPMA
"@ | convertfrom-csv -Delimiter ";"

$department_weighted = @"
[
    { "DPT": "ADM", "Weight": 3 },
    { "DPT": "MRKT", "Weight": 2 },
    { "DPT": "OPS", "Weight": 5 },
    { "DPT": "IT", "Weight": 1 },
    { "DPT": "ACC", "Weight": 1 },
    { "DPT": "FIN", "Weight": 1 }
]
"@ | ConvertFrom-Json

$office_weighted = @"
[
    { "country": "CH", "city": "Münchenstein", "Weight" : 6 },
    { "country": "CH", "city": "Olten", "Weight" : 3 },
    { "country": "CH", "city": "Winterthur", "Weight" : 1 },
    { "country": "FR", "city": "Saint-Louis", "Weight" : 4 },
    { "country": "FR", "city": "Paris", "Weight" : 1 },
    { "country": "FR", "city": "Strasbourg", "Weight" : 1 },
    { "country": "FR", "city": "Lyon", "Weight" : 1 },
    { "country": "DE", "city": "Karlsruhe", "Weight" : 1 },
    { "country": "ES", "city": "Madrid", "Weight" : 1 }
]
"@ | ConvertFrom-Json

$random_department_array = @()
$department_weighted | ForEach-Object {
    $this_config_line = $_
    1..$this_config_line.Weight | ForEach-Object { $random_department_array += $this_config_line.DPT }
}

$random_office_array = @()
$office_weighted | ForEach-Object {
    $this_config_line = $_
    1..$this_config_line.Weight | ForEach-Object { $random_office_array += $this_config_line }
}

function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [ValidateRange(4, [int]::MaxValue)]
        [int] $length,
        [int] $upper = 1,
        [int] $lower = 1,
        [int] $numeric = 1,
        [int] $special = 1
    )
    if ($upper + $lower + $numeric + $special -gt $length) {
        throw "number of upper/lower/numeric/special char must be lower or equal to length"
    }
    $uCharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $lCharSet = "abcdefghijklmnopqrstuvwxyz"
    $nCharSet = "0123456789"
    $sCharSet = "*-+!?=@"
    $charSet = ""
    if ($upper -gt 0) { $charSet += $uCharSet }
    if ($lower -gt 0) { $charSet += $lCharSet }
    if ($numeric -gt 0) { $charSet += $nCharSet }
    if ($special -gt 0) { $charSet += $sCharSet }
    
    $charSet = $charSet.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
    $rng.GetBytes($bytes)
 
    $result = New-Object char[]($length)
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
    }
    $password = (-join $result)
    $valid = $true
    if ($upper -gt ($password.ToCharArray() | Where-Object { $_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
    if ($lower -gt ($password.ToCharArray() | Where-Object { $_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
    if ($numeric -gt ($password.ToCharArray() | Where-Object { $_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }
    if ($special -gt ($password.ToCharArray() | Where-Object { $_ -cin $sCharSet.ToCharArray() }).Count) { $valid = $false }
 
    if (!$valid) {
        $password = Get-RandomPassword $length $upper $lower $numeric $special
    }
    return $password
}

function Remove-StringDiacritic {
    <#
    .SYNOPSIS
        This function will remove the diacritics (accents) characters from a string.
        
    .DESCRIPTION
        This function will remove the diacritics (accents) characters from a string.
    
    .PARAMETER String
        Specifies the String on which the diacritics need to be removed
    
    .PARAMETER NormalizationForm
        Specifies the normalization form to use
        https://msdn.microsoft.com/en-us/library/system.text.normalizationform(v=vs.110).aspx
    
    .EXAMPLE
        PS C:\> Remove-StringDiacritic "L'été de Raphaël"
        
        L'ete de Raphael
    
    .NOTES
        Francois-Xavier Cat
        @lazywinadmin
        www.lazywinadmin.com
#>
    
    param
    (
        [ValidateNotNullOrEmpty()]
        [Alias('Text')]
        [System.String]$String,
        [System.Text.NormalizationForm]$NormalizationForm = "FormD"
    )
    
    BEGIN {
        $Normalized = $String.Normalize($NormalizationForm)
        $NewString = New-Object -TypeName System.Text.StringBuilder
        
    }
    PROCESS {
        $normalized.ToCharArray() | ForEach-Object -Process {
            if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$NewString.Append($psitem)
            }
        }
    }
    END {
        Write-Output $($NewString -as [string])
    }
}

$random_users = @()


$data | Select-Object * | ForEach-Object {
    $this_user = $_
    #$this_user | ConvertTo-Json -Compress
    $this_active_directory_user = "" | Select-Object employeeid, employeetype, department, "GivenName", "Surname", "DisplayName", "TelephoneNumber", "City", "Country", "Company", "SAMAccountName", "UserPrincipalName", "AccountPassword", "Description", "Path", psw,company_short
    if ( $log -eq 1 ) { write-host ( "#### New row, number {0} #####" -F $this_user.Number ) }
    
    #region TelephoneNumber
    if ( $log -eq 1 ) { write-host ("    - TelephoneNumber:{0}" -F $this_user.TelephoneNumber ) }
    $this_user.TelephoneNumber = $this_user.TelephoneNumber -replace '[^a-zA-Z0-9]', ''
    if ( $this_user.TelephoneNumber.Substring(0, 1) -eq "0") {
        $this_user.TelephoneNumber = $this_user.TelephoneNumber -replace "^0" , ""
    }
    if ( $log -eq 1 ) { write-host ("    - Country:{0}" -F $this_user.Country ) }
    switch ($this_user.Country) {
        "ES" { $prefix = "+34" }
        "FR" { $prefix = "+33" }
        "DE" { $prefix = "+49" }
        "CH" { $prefix = "+41" }
        default { $prefix = "+0" }
    }
    $this_active_directory_user.TelephoneNumber = "{0}{1}" -F $prefix , $this_user.TelephoneNumber
    if ( $log -eq 1 ) { write-host ("  > TelephoneNumber:{0}" -F $this_active_directory_user.TelephoneNumber ) }
    #endregion
    
    #region sn, givenname, samaccountname, employeeid, employeetype
    $this_active_directory_user.Surname = $this_user.Surname
    $this_active_directory_user.GivenName = $this_user.givenname
    if ( $log -eq 1 ) { write-host ("    - Surname:{0}" -F $this_active_directory_user.Surname ) }
    if ( $log -eq 1 ) { write-host ("    - GivenName:{0}" -F $this_active_directory_user.GivenName ) }
    
    $this_active_directory_user.SAMAccountName = "{0}{1}{2}" -F ($this_active_directory_user.GivenName).Substring(0, 1), ($this_active_directory_user.Surname).Substring(0, 2).ToUpper(), '{0:d3}' -f [int]$this_user.Number 
    $this_active_directory_user.SAMAccountName = Remove-StringDiacritic -String $this_active_directory_user.SAMAccountName
    if ( $log -eq 1 ) { write-host ("  > SAMAccountName:{0}" -F $this_active_directory_user.SAMAccountName ) }

    #employeeid
    $this_active_directory_user.employeeid = '{0:d8}' -f [int]$this_user.Number 

    #employeetype
    $this_active_directory_user.employeetype = "employee"

    #displayname 
    $this_active_directory_user.DisplayName = "{1} {0}" -F $this_active_directory_user.GivenName, $this_active_directory_user.Surname
    #endregion

    #region company & upn
    $this_number_curent = [int]$this_user.Number 
    #write-host " input : $this_number_curent"
    $this_companies = $segments_list | Where-Object { $_.country -eq $this_user.Country } 
    while ( $this_number_curent -ge $this_companies.count ) {
        $this_number_curent = $this_number_curent - $this_companies.count 
        #write-host " iteration: $this_number_curent"
    }
    #write-host " ouput : $this_number_curent"

    $this_active_directory_user.company = $this_companies[$this_number_curent - 1].company
    $this_active_directory_user.company_short = $this_companies[$this_number_curent - 1].short
    $this_active_directory_user.UserPrincipalName = "{0}.{1}@{2}" -F $this_active_directory_user.GivenName, $this_active_directory_user.Surname, $this_companies[$this_number_curent - 1].upn_suffix
    $this_active_directory_user.path = $this_companies[$this_number_curent - 1].ou
    
    if ( $log -eq 1 ) { write-host ("  > company:{0}" -F $this_active_directory_user.company ) }
    if ( $log -eq 1 ) { write-host ("  > UserPrincipalName:{0}" -F $this_active_directory_user.UserPrincipalName ) }
    #endregion

    # department
    $this_active_directory_user.department = "{0}-{1}" -F $this_active_directory_user.company_short ,($random_department_array[(Get-Random -Minimum 0 -Maximum ($random_department_array.count - 1))] )
    if ( $log -eq 1 ) { write-host ("  > department:{0}" -F $this_active_directory_user.department) }

    # city & country
    $this_active_directory_user.Country = $this_user.Country
    $localised_random_office_array = $random_office_array | Where-Object { $_.country -eq $this_user.Country }
    #write-host ( " count :{0}" -F $localised_random_office_array.count )
    
    $this_active_directory_user.city = $localised_random_office_array[(Get-Random -Minimum 0 -Maximum (($localised_random_office_array | Measure-Object).count ))].city

    $psw = Get-RandomPassword -length 16
    $this_active_directory_user.AccountPassword = ConvertTo-SecureString -AsPlainText $psw -Force 
    
    $this_active_directory_user.Description = " psw:{0}" -F $psw
    $this_active_directory_user.psw = $psw


    $random_users += $this_active_directory_user
    #$this_active_directory_user | ConvertTo-Json -Compress

}


$random_users | ForEach-Object {
    $employeeid = $_.employeeid
    $employeetype = $_.employeetype
    $department = $_.department
    $GivenName = $_.GivenName
    $Surname = $_.Surname
    $DisplayName = $_.DisplayName
    $TelephoneNumber = $_.TelephoneNumber
    $City = $_.City
    $Country = $_.Country
    $Company = $_.Company
    $SAMAccountName = $_.SAMAccountName
    $UserPrincipalName = $_.UserPrincipalName
    $AccountPassword = $_.AccountPassword
    $Description = $_.Description
    $Path = $_.Path
    $PasswordClearText = $_.psw

    $UserProperties = @{
        "GivenName"             = $GivenName
        "Surname"               = $Surname
        "employeeid"            = $employeeid
        "department"            = $department
        "Name"                  = $SAMAccountName + " (" + $Surname + " " + $GivenName + ")"
        "DisplayName"           = $DisplayName
        "OfficePhone"           = $TelephoneNumber
        "City"                  = $City
        "Country"               = $Country
        "Company"               = $Company
        "SAMAccountName"        = $SAMAccountName
        "UserPrincipalName"     = $UserPrincipalName
        "Enabled"               = $true
        "accountpassword"       = (ConvertTo-SecureString -AsPlainText "replacedImmediately20023" -Force)
        "ChangePasswordAtLogon" = $False
        "Description"           = $Description
        "Path"                  = $Path
    }

    TRY { 
        New-ADUser @UserProperties 
        $PasswordClearText
        set-adaccountpassword  -Identity $SAMAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $PasswordClearText -Force)
        $cmd = "NET USE \\dc1\netlogon /U:pe\{0} {1}" -F $samaccountname, $PasswordClearText
    
        Invoke-Expression $cmd
        Invoke-Expression "net use * /d /y"

        Set-Date -Adjust 8:00:0

    }
    CATCH { 
        

    }



}


w32tm /resync /force