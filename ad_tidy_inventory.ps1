Clear-Host
#region ## script information
$Global:Version = "0.0.1"
# HAR3005, Primeo-Energie, 20240209
#    Initial draft
#endregion


Set-Location $PSScriptRoot
$Global:ConfigFileName = "ad_tidy.config.json"
$Global:RulesConfigFileName = "ad_tidy.rules.config.csv"


#region ## global configuration variables
$Global:Debug = $true
$Global:WhatIf = $false # no actual sql operation happen
$Global:LogLocation = $PSScriptRoot
#$Global:LogLocation = "C:\IT_Staff\Logs"
#endregion

#region ## script specific configuration
$SQLRecordTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$Global:MatchError = 0
#endregion

#Region # PRE-Script
Set-Location $PSScriptRoot
Get-ChildItem $Global:LogLocation | Where-Object { $_.Length -gt 50 * 1024 * 1024 } | ForEach-Object {	Remove-Item $_.FullName -Force }  # clearing any log files bigger than 50 Mb
#endregion

#region ## Includes management
Set-Location .\includes
Get-ChildItem -Filter *.ps1 | ForEach-Object {
    . $_.FullName
    TRY { Global:log -text (" --> Included '{0}'" -f $_.name ) -Hierarchy "Includes" }  CATCH {}
}
Global:log -text ("Done.") -Hierarchy "Includes:Generic"
Set-Location $PSScriptRoot
#endregion

#region ## config definition
Set-Location $PSScriptRoot
$len = (Get-Item $Global:ConfigFileName ).length / 1kb
$mod = (Get-Item $Global:ConfigFileName ).LastWriteTime
$global:Config = Get-Content .\$Global:ConfigFileName -Raw -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | ConvertFrom-Json -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
Global:Log -text ("Json config loaded ({0}kb, modified on {1})" -f $len, $mod  ) -hierarchy ("{0}" -f ($MyInvocation.ScriptName).split("\")[($MyInvocation.ScriptName).split("\").count - 1])


$len = (Get-Item $Global:RulesConfigFileName ).length / 1kb
$mod = (Get-Item $Global:RulesConfigFileName ).LastWriteTime
$global:Rules = Import-Csv .\$Global:RulesConfigFileName -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | Select-Object *, @{name = "sort"; expression = { [int]$_."Processing order" } }
Global:Log -text ("csv rules config loaded ( {0}kb, modified on {1}, {2} row(s) )" -f $len, $mod, ($global:Rules | Measure-Object).Count  ) -hierarchy ("{0}" -f ($MyInvocation.ScriptName).split("\")[($MyInvocation.ScriptName).split("\").count - 1])


#endregion

#region main
Global:log -text ("Start V{0}" -F $Global:Version) -Hierarchy "Main"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

Global:ADTidy_Inventory_Users_sql_table_check

$last_update = Global_ADTidy_Iventory_Users_last_update

if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
    $filter = "*"
}
else {
    
    $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
    $filter = "whenchanged -ge '$filter_date'"
}



#region users
Global:log -text ("retrieving users from AD, filter='{0}'" -F $filter) -Hierarchy "Main:Users"
$users = @()
<# PRD#>
Get-ADUser  -properties $global:config.Configurations.inventory.'Active Directory Attributes' -filter $filter  | ForEach-Object { 
    <# DEV 
Get-ADUser  -properties $global:config.Configurations.inventory.'Active Directory Attributes' -filter "samaccountname -eq 'har3005'"  | ForEach-Object { #> 


    $this_row = $_
    $this_calculated_row = "" | Select-Object ignore

    $this_row | Get-Member | Where-Object { $_.membertype -eq "property" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute = $_
        $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
        Switch ( $this_attribute ) {
            "accountexpires" {
                $AccountExpiresRaw = [string]($this_row."$this_attribute")
                if ( $AccountExpiresRaw -eq "9223372036854775807" ) { 
                    $AccountExpires = $null
                }
                else {
                    $AccountExpires = [datetime]::FromFileTime($AccountExpiresRaw).ToString("yyyy-MM-dd HH:mm:ss") 
                    if ( $AccountExpires -eq '1601-01-01 01:00:00' ) { $AccountExpires = $null }
                }
                if ( $AccountExpires.Length -eq 0 ) { $CalulatedValue = "NULL" }
                ELSE { $CalulatedValue = "$AccountExpires" }
                           
                $this_calculated_row."$this_attribute" = $CalulatedValue

            }
            "lastLogonTimestamp" {
                $AccountExpiresRaw = [string]($this_row."$this_attribute")
                if ( $AccountExpiresRaw -eq "9223372036854775807" ) { 
                    $AccountExpires = $null
                }
                else {
                    $AccountExpires = [datetime]::FromFileTime($AccountExpiresRaw).ToString("yyyy-MM-dd HH:mm:ss") 
                    if ( $AccountExpires -eq '1601-01-01 01:00:00' ) { $AccountExpires = $null }
                }
                if ( $AccountExpires.Length -eq 0 ) { $CalulatedValue = "NULL" }
                ELSE { $CalulatedValue = "$AccountExpires" }
                           
                $this_calculated_row."$this_attribute" = $CalulatedValue

            }
            "pwdLastSet" {
                $pwdLastSetRaw = [string]($this_row."$this_attribute")
                if ( $pwdLastSetRaw -eq "9223372036854775807" ) { 
                    $pwdLastSet = $null
                }
                else {
                    $pwdLastSet = [datetime]::FromFileTime($pwdLastSetRaw).ToString("yyyy-MM-dd HH:mm:ss") 
                    if ( $pwdLastSet -eq '1601-01-01 01:00:00' ) { $pwdLastSet = $null }
                }
                if ( $pwdLastSet.Length -eq 0 ) { $CalulatedValue = "NULL" }
                ELSE { $CalulatedValue = "$pwdLastSet" }
                           

                $this_calculated_row."$this_attribute" = $CalulatedValue
                
            }
            "useraccountcontrol" {
                #write-host "> useraccountcontrol"
                $CalulatedValue = Global:DecodeUserAccountControl ([int][string]($this_row."$this_attribute"))
                #$CalulatedValue = "'$CalulatedValue'"
                $this_calculated_row."$this_attribute" = $CalulatedValue
            }
            "thumbnailPhoto" {
                if ( 0) {
                    #Write-Host "thumbnailPhoto"
                            
                    if ( $ThisFieldValue -ne $null) {
                        $CalulatedValue = "Is set"
                    }
                    else {
                        $CalulatedValue = "NULL"
                    }
                    #Write-Host "value:$ThisFieldValue"
                }
            }
            "whenChanged" {
                TRY {
                    $CalulatedValue = '{0:yyyy-MM-dd HH:mm:ss}' -f $this_row."$this_attribute"
                    $this_calculated_row."$this_attribute" = $CalulatedValue 
                }
                CATCH {
                    $this_calculated_row."$this_attribute" = $null
                }
                #write-host ( "   > {0} - {1}" -F ($this_row."$this_attribute"), $this_attribute)
                
                
            }
            "whenCreated" {
                TRY {
                    $CalulatedValue = '{0:yyyy-MM-dd HH:mm:ss}' -f $this_row."$this_attribute"
                    $this_calculated_row."$this_attribute" = $CalulatedValue 
                }
                CATCH {
                    $this_calculated_row."$this_attribute" = $null
                }
                #write-host ( "   > {0} - {1}" -F ($this_row."$this_attribute"), $this_attribute)
                
            }
            default {
                $this_calculated_row."$this_attribute" = $this_row."$this_attribute" -replace "'", "''" 
            }

        }
    }
    
    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, surname)
    $users += $this_calculated_row

}

Global:log -text ("{0} user objects retrieved" -F $users.Count) -Hierarchy "Main:Users"

$users | Select-Object * -ExcludeProperty name, objectclass, enabled | ForEach-Object {
    $this_user = $_
    Global:ADTidy_Inventory_Users_sql_update -Fields $this_user

    
}

#endregion

#endregion