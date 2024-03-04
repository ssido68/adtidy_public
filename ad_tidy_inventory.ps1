Clear-Host
#region ## script information
$Global:Version = "1.0.0"
# HAR3005, Primeo-Energie, 20240227
#    gather delta users, computers, OU and groups objects from Active Directory based of last whenchanged attribute for newer records
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


#region record management, init and templates
$record_template = "" | Select-Object record_type, rule_name, target_list, result_summary, log_json
$summary_template = "" | Select-Object retrieved, updated, created, deleted
$target_item_template = "" | Select-Object name, action
#endregion


#region OU
Global:ADTidy_Inventory_OU_sql_table_check

#region record init
$ou_record = $record_template | Select-Object *
$ou_record.record_type = "AdTidy.inventory"
$ou_record.rule_name = "OU"
$ou_record.target_list = @()
$ou_record.log_json = @()


$ou_summary = $summary_template | Select-Object *
$ou_summary.retrieved = 0
$ou_summary.updated = 0
$ou_summary.created = 0
$ou_summary.deleted = 0

$ou_record.result_summary = $ou_summary | ConvertTo-Json -Compress

$ou_target_item_array = @()
#endregion

$last_update = Global_ADTidy_Iventory_OU_last_update

if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
    $filter = "*"
}
else {
    
    $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
    $filter = "whenchanged -gt '$filter_date'"
}
Global:log -text ("retrieving users from AD, filter='{0}'" -F $filter) -Hierarchy "Main:Ou"
$Organizational_units = @()
Get-ADOrganizationalUnit -Filter $filter -Properties $global:Config.Configurations.inventory.'OU Active Directory Attributes' | Select-Object name, whenCreated, whenChanged, distinguishedname, objectguid, businessCategory, managedBy | Sort-Object whenChanged | ForEach-Object { 
    $this_row = $_

    $this_calculated_row = "" | Select-Object ignore

    $this_row | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute = $_
        $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
        Switch ( $this_attribute ) {
            "businesscategory" {
                TRY {
                    $businesscategory_string = ""
                    $this_row."$this_attribute" | ForEach-Object {
                        $businesscategory_string = "{0}{1}," -F $businesscategory_string, $_
                    }
                    # remove last ',' from string
                    $businesscategory_string = $businesscategory_string -replace ".$"
                    $this_calculated_row."$this_attribute" = $businesscategory_string
                    if (([string]$this_row."$this_attribute").length -eq 0 ) { $this_calculated_row."$this_attribute" = "NULL" }
                }
                CATCH {

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
                if (([string]$this_row."$this_attribute").length -eq 0 ) { $this_calculated_row."$this_attribute" = "NULL" }
            }

        }
    }
    
    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore)
    $Organizational_units += $this_calculated_row

}

if (  $Organizational_units.Count -eq 0 ) {
    Global:log -text ("No Ou objects changes found" -F $Organizational_units.Count) -Hierarchy "Main:Ou:delta changes" -type warning

}
else {
    Global:log -text ("{0} Ou objects retrieved" -F $Organizational_units.Count) -Hierarchy "Main:Ou:delta changes"


    $Organizational_units | Select-Object * | ForEach-Object {
        $this_ou = $_
        $this_record_item = $target_item_template | Select-Object *
        $this_record_item.name = $this_ou.distinguishedname
        switch ( Global:ADTidy_Inventory_OU_sql_update -Fields $this_ou) {
            "updated" {
                $this_record_item.action = "updated"
                $ou_summary.updated++
            }
            "new" {
                $this_record_item.action = "created"
                $ou_summary.created++

            }
        }
        $ou_target_item_array += $this_record_item
    }
}
#region deleted records detect
$sql_current_records = Global_ADTidy_Iventory_OU_all_current_records 
$ad_current_records = Get-ADOrganizationalUnit -filter * -Properties objectguid | Select-Object -ExpandProperty objectguid
$ou_summary.retrieved = $ad_current_records.Count
Global:log -text ("SQL:{0} current records, AD:{1} current records " -F $sql_current_records.Count, $ad_current_records.Count) -Hierarchy "Main:Ou:deleted detect"
$flag_deleted = 0

$sql_current_records | ForEach-Object {
    $this_sql_record = $_
    if ( $ad_current_records -notcontains $this_sql_record.ad_objectguid) {
        Global:log -text ("Detected a deleted OU:'{0}' " -F ($this_sql_record | Select-Object ad_name, ad_objectguid, ad_dinstinguishedname | ConvertTo-Json -Compress)) -Hierarchy "Main:OU:deleted detect"
        $delete_record = $this_sql_record | Select-Object @{name = 'Objectguid'; expression = { $_.ad_ObjectGUID } }, record_status
        $delete_record.record_status = "Deleted"
        Global:ADTidy_Inventory_OU_sql_update -Fields $delete_record
        $ou_summary.deleted++
        $flag_deleted = 1


        $this_record_item = $target_item_template | Select-Object *
        $this_record_item.name = $this_sql_record.ad_distinguishedname
        $this_record_item.action = "deleted"
        $ou_target_item_array += $this_record_item


    }
}
if ( $flag_deleted -eq 0) {
    Global:log -text ("No deleted OU record found." ) -Hierarchy "Main:OU:deleted detect" -type warning

}

$ou_record.result_summary = $ou_summary | ConvertTo-Json -Compress
$ou_record.target_list = $ou_target_item_array | ConvertTo-Json -Compress
$ou_record
exit
#endregion

    
#endregion

#region users
Global:ADTidy_Inventory_Users_sql_table_check

$last_update = Global_ADTidy_Iventory_Users_last_update

if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
    $filter = "*"
}
else {
    
    $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
    $filter = "whenchanged -gt '$filter_date'"
}
Global:log -text ("retrieving users from AD, filter='{0}'" -F $filter) -Hierarchy "Main:Users"
$users = @()

#region delta changes
<# PRD#>
Get-ADUser  -properties $global:config.Configurations.inventory.'Users Active Directory Attributes' -filter $filter  | ForEach-Object { 
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
                if (([string]$this_row."$this_attribute").length -eq 0 ) { $this_calculated_row."$this_attribute" = "NULL" }
            }

        }
    }
    
    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, surname)
    $users += $this_calculated_row

}

if ( $users.Count -eq 0 ) {
    Global:log -text ("No  user objects changes found" -F $users.Count) -Hierarchy "Main:Ou:delta changes" -type warning
}
else {
    Global:log -text ("{0} user objects retrieved" -F $users.Count) -Hierarchy "Main:Users:delta changes"

    $users | Select-Object * -ExcludeProperty name, objectclass, enabled | ForEach-Object {
        $this_user = $_
        Global:ADTidy_Inventory_Users_sql_update -Fields $this_user
    }

}

#endregion

#region deleted records detect
$sql_current_records = Global_ADTidy_Iventory_Users_all_current_records 
$ad_current_records = Get-ADUser -filter * -Properties objectguid | Select-Object -ExpandProperty objectguid
Global:log -text ("SQL:{0} current records, AD:{1} current records " -F $sql_current_records.Count, $ad_current_records.Count) -Hierarchy "Main:Users:deleted detect"
$flag_deleted = 0

$sql_current_records | ForEach-Object {
    $this_sql_record = $_
    if ( $ad_current_records -notcontains $this_sql_record.ad_objectguid) {
        Global:log -text ("Detected a deleted AD account:'{0}' " -F ($this_sql_record | Select-Object ad_samaccountname, ad_objectguid, ad_sid | ConvertTo-Json -Compress)) -Hierarchy "Main:Users:deleted detect"
        $delete_record = $this_sql_record | Select-Object @{name = 'Objectguid'; expression = { $_.ad_ObjectGUID } }, record_status
        $delete_record.record_status = "Deleted"
        Global:ADTidy_Inventory_Users_sql_update -Fields $delete_record
        $flag_deleted = 1
    }
}
if ( $flag_deleted -eq 0) {
    Global:log -text ("No deleted user record found." ) -Hierarchy "Main:Users:deleted detect" -type warning

}
#endregion

#endregion

#region computers
Global:ADTidy_Inventory_Computers_sql_table_check


$last_update = Global_ADTidy_Iventory_Computers_last_update

if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
    $filter = "*"
}
else {
    
    $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
    $filter = "whenchanged -gt '$filter_date'"
}
Global:log -text ("retrieving computers from AD, filter='{0}'" -F $filter) -Hierarchy "Main:Users"

$computers = @()

#region delta changes
<# PRD#>
Get-ADComputer -properties $global:config.Configurations.inventory.'Computers Active Directory Attributes' -filter $filter  | ForEach-Object { 
    <# DEV 
Get-ADUser  -properties $global:config.Configurations.inventory.'Active Directory Attributes' -filter "samaccountname -eq 'har3005'"  | ForEach-Object { #> 


    $this_row = $_
    $this_calculated_row = "" | Select-Object ignore

    $this_row | Get-Member | Where-Object { $_.membertype -eq "property" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute = $_
        $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
        Switch ( $this_attribute ) {
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
                if (([string]$this_row."$this_attribute").length -eq 0 ) { $this_calculated_row."$this_attribute" = "NULL" }
            }

        }
    }
    
    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, enabled)
    $computers += $this_calculated_row

}

if ( $computers.Count -eq 0 ) {
    Global:log -text ("No  computers objects changes found" -F $computers.Count) -Hierarchy "Main:Computers:delta changes" -type warning
}
else {
    Global:log -text ("{0} computers objects retrieved" -F $computers.Count) -Hierarchy "Main:Computers:delta changes"

    $computers | Select-Object * -ExcludeProperty  objectclass, DNSHostName, UserPrincipalName | ForEach-Object {
        $this_computer = $_
        Global:ADTidy_Inventory_Computers_sql_update -Fields $this_computer
    }

}
#endregion

#region deleted records detect
$sql_current_records = Global_ADTidy_Iventory_Computers_all_current_records 
$ad_current_records = Get-Adcomputer -filter * -Properties objectguid | Select-Object -ExpandProperty objectguid
Global:log -text ("SQL:{0} current records, AD:{1} current records " -F $sql_current_records.Count, $ad_current_records.Count) -Hierarchy "Main:Computers:deleted detect"
$flag_deleted = 0

$sql_current_records | ForEach-Object {
    $this_sql_record = $_
    if ( $ad_current_records -notcontains $this_sql_record.ad_objectguid) {
        Global:log -text ("Detected a deleted AD account:'{0}' " -F ($this_sql_record | Select-Object ad_samaccountname, ad_objectguid, ad_sid | ConvertTo-Json -Compress)) -Hierarchy "Main:Computers:deleted detect"
        $delete_record = $this_sql_record | Select-Object @{name = 'Objectguid'; expression = { $_.ad_ObjectGUID } }, record_status
        $delete_record.record_status = "Deleted"
        Global:ADTidy_Inventory_Computers_sql_update -Fields $delete_record
        $flag_deleted = 1
    }
}
if ( $flag_deleted -eq 0) {
    Global:log -text ("No deleted computer record found." ) -Hierarchy "Main:Computers:deleted detect" -type warning

}
#endregion

#endregion

#region Groups
Global:ADTidy_Inventory_Groups_sql_table_check


$last_update = Global_ADTidy_Iventory_Groups_last_update

if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
    $filter = "*"
}
else {
    
    $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
    $filter = "whenchanged -gt '$filter_date'"
}
Global:log -text ("retrieving users from AD, filter='{0}'" -F $filter) -Hierarchy "Main:Groups"
$groups = @()


#region delta changes
<# PRD#>
Get-ADGroup -properties $global:config.Configurations.inventory.'Groups Active Directory Attributes' -filter $filter  | Sort-Object whenchanged  | ForEach-Object { 
    <# DEV 
Get-ADUser  -properties $global:config.Configurations.inventory.'Active Directory Attributes' -filter "samaccountname -eq 'har3005'"  | ForEach-Object { #> 


    $this_row = $_
    $this_calculated_row = "" | Select-Object ignore

    $this_row | Get-Member | Where-Object { $_.membertype -eq "property" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute = $_
        $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
        Switch ( $this_attribute ) {
            "member" {
                $members_array = @()
                $this_row."$this_attribute" | ForEach-Object {
                    $line = "" | Select-Object distinguishedname
                    $line.distinguishedName = $_
                    $members_array += $line
                }
                $this_calculated_row = $this_calculated_row | Select-Object *, "xml_members"
                $this_calculated_row."xml_members" = Global:ConvertTo-SimplifiedXML -InputObject $members_array -RootNodeName "Members" -NodeName "Member"
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
                if (([string]$this_row."$this_attribute").length -eq 0 ) { $this_calculated_row."$this_attribute" = "NULL" }
            }

        }
    }
    
    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, surname)
    $groups += $this_calculated_row

}

if ( $groups.Count -eq 0 ) {
    Global:log -text ("no group objects changes found" -F $groups.Count) -Hierarchy "Main:Groups:delta changes" -type warning
}
else {
    Global:log -text ("{0} group objects retrieved" -F $groups.Count) -Hierarchy "Main:Groups:delta changes"

    $groups | Select-Object * -ExcludeProperty objectclass | ForEach-Object {
        $this_group = $_
        Global:ADTidy_Inventory_Groups_sql_update -Fields $this_group
    }
}
#endregion


#region deleted records detect
$sql_current_records = Global_ADTidy_Iventory_Groups_all_current_records 
$ad_current_records = Get-ADGroup -filter * -Properties objectguid | Select-Object -ExpandProperty objectguid
Global:log -text ("SQL:{0} current records, AD:{1} current records " -F $sql_current_records.Count, $ad_current_records.Count) -Hierarchy "Main:Groups:deleted detect"
$flag_deleted = 0

$sql_current_records | ForEach-Object {
    $this_sql_record = $_
    if ( $ad_current_records -notcontains $this_sql_record.ad_objectguid) {
        Global:log -text ("Detected a deleted AD group:'{0}' " -F ($this_sql_record | Select-Object ad_samaccountname, ad_objectguid, ad_sid | ConvertTo-Json -Compress)) -Hierarchy "Main:Users:deleted detect"
        $delete_record = $this_sql_record | Select-Object @{name = 'Objectguid'; expression = { $_.ad_ObjectGUID } }, record_status
        $delete_record.record_status = "Deleted"
        Global:ADTidy_Inventory_Groups_sql_update -Fields $delete_record
        $flag_deleted = 1
    }
}
if ( $flag_deleted -eq 0) {
    Global:log -text ("No deleted group record found." ) -Hierarchy "Main:Groups:deleted detect" -type warning

}
#endregion

#endregion
