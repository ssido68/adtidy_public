Clear-Host
#region ## script information
$Global:Version = "1.0.0"
# HAR3005, Primeo-Energie, 20240227
#    gather delta users, computers, OU and groups objects from Active Directory based of last whenchanged attribute for newer records
$Global:Version = "1.0.1"
# HAR3005, Primeo-Energie, 20240301
#    Added group nested membership lookup
#endregion



#region ## global configuration variables
$Global:Debug = $true
$Global:WhatIf = $false # no actual sql operation happen
$Global:LogLocation = $PSScriptRoot
#$Global:LogLocation = "C:\IT_Staff\Logs"
Set-Location $PSScriptRoot
$Global:ConfigFileName = "ad_tidy.config.json"

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
#endregion

#region main
Global:log -text ("Start V{0}" -F $Global:Version) -Hierarchy "Main"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$flag_inventory_ou = $false
$flag_inventory_users = $false
$flag_inventory_computers = $false
$flag_inventory_groups = $true


#region record management, init and templates
$record_template = "" | Select-Object record_type, rule_name, target_list, result_summary
$summary_template = "" | Select-Object database, retrieved, updated, created, deleted
$target_item_template = "" | Select-Object name, action
#endregion


#region OU
if ($flag_inventory_ou -eq $true ) {

    Global:ADTidy_Inventory_OU_sql_table_check

    #region record init
    $ou_record = $record_template | Select-Object *
    $ou_record.record_type = "AdTidy.inventory"
    $ou_record.rule_name = "OU"
    $ou_record.target_list = @()


    $ou_summary = $summary_template | Select-Object *
    $ou_summary.database = 0
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

    #region delta changes
    Get-ADOrganizationalUnit -Filter $filter -Properties $global:Config.Configurations.inventory.'OU Active Directory Attributes' -Server $global:Config.Configurations.'target domain controller' | Sort-Object whenchanged | Select-Object name, whenCreated, whenChanged, distinguishedname, objectguid, businessCategory, managedBy | Sort-Object whenChanged | ForEach-Object { 
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


        $Organizational_units | Sort-Object whenchanged | ForEach-Object {
            $this_ou = $_
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_ou.distinguishedname
            switch ( Global:ADTidy_Inventory_OU_sql_update -Fields $this_ou) {
                "update" {
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
    #endregion

    #region deleted records detect
    $sql_current_records = Global_ADTidy_Iventory_OU_all_current_records 
    $ad_current_records = Get-ADOrganizationalUnit -filter * -Properties objectguid  -Server $global:Config.Configurations.'target domain controller' | Select-Object -ExpandProperty objectguid
    $ou_summary.retrieved = $ad_current_records.Count
    $ou_summary.database = $sql_current_records.Count
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
    #endregion

    #region missing records
    if ( $sql_current_records.Count -lt $ad_current_records.Count ) {
        Global:log -text ("sql_current_records.Count < ad_current_records.Count, {0} missing records in database.... " -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Ou:missing records" -type warning
        $existing_sql_objects_guid = ( $sql_current_records | Select-Object -ExpandProperty ad_objectguid )
        $Organizational_units_missing = @()
        $Organizational_units_missing_count = 0
        $ad_current_records | ForEach-Object {
            $this_ad_record = $_
            #Global:log -text ("checking guid={0}" -F $this_ad_record.guid) -Hierarchy "Main:Ou:missing records" 
            if ( $existing_sql_objects_guid -notcontains $this_ad_record.guid -and $Organizational_units_missing_count -lt $global:Config.Configurations.inventory.'max missing records') {
                $Organizational_units_missing_count++
                Global:log -text ("missing object guid={0}" -F $this_ad_record.guid, $this_ad_record.distinguishedName ) -Hierarchy "Main:Ou:missing records" 
                $filter = "objectguid -eq  '{0}'" -F $this_ad_record.guid
                Get-ADOrganizationalUnit -Filter $filter -Properties $global:Config.Configurations.inventory.'OU Active Directory Attributes' -Server $global:Config.Configurations.'target domain controller' | Sort-Object whenchanged | Select-Object name, whenCreated, whenChanged, distinguishedname, objectguid, businessCategory, managedBy | ForEach-Object { 
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
                    $Organizational_units_missing += $this_calculated_row

                }

            }
        
        }
        if ( $Organizational_units_missing_count -eq $global:Config.Configurations.inventory.'max missing records' ) {
            Global:log -text ("Max missing records count reached ({0})" -F $global:Config.Configurations.inventory.'max missing records' ) -Hierarchy "Main:Ou:missing records" -type warning
        }

    
        $Organizational_units_missing | ForEach-Object {
            $this_ou = $_
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_ou.distinguishedname
            switch ( Global:ADTidy_Inventory_OU_sql_update -Fields $this_ou) {
                "update" {
                    $this_record_item.action = "updated"
                    $ou_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $ou_summary.created++

                }
            }
            $ou_summary.database++

            $ou_target_item_array += $this_record_item
        }

    }
    else {
        Global:log -text ("sql_current_records.Count = ad_current_records.Count, no missing records in database" -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Ou:missing records"
    }
    #endregion

    $ou_record.result_summary = $ou_summary | ConvertTo-Json -Compress
    $ou_record.target_list = $ou_target_item_array | ConvertTo-Json -Compress

    Global:ADTidy_Records_sql_update -Fields $ou_record
}
#endregion

#region users
if ($flag_inventory_users -eq $true ) {

    Global:ADTidy_Inventory_Users_sql_table_check

    #region record init
    $users_record = $record_template | Select-Object *
    $users_record.record_type = "AdTidy.inventory"
    $users_record.rule_name = "users"
    $users_record.target_list = @()


    $users_summary = $summary_template | Select-Object *
    $users_summary.database = 0
    $users_summary.retrieved = 0
    $users_summary.updated = 0
    $users_summary.created = 0
    $users_summary.deleted = 0

    $users_record.result_summary = $users_summary | ConvertTo-Json -Compress

    $users_target_item_array = @()
    #endregion

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
    Get-ADUser  -properties $global:config.Configurations.inventory.'Users Active Directory Attributes' -filter $filter  -Server $global:Config.Configurations.'target domain controller' | ForEach-Object { 
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
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_user.samaccountname
            switch ( Global:ADTidy_Inventory_Users_sql_update -Fields $this_user) {
                "update" {
                    $this_record_item.action = "updated"
                    $users_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $users_summary.created++
                }
            }
            $users_target_item_array += $this_record_item
        }

    }

    #endregion

    #region deleted records detect
    $sql_current_records = Global_ADTidy_Iventory_Users_all_current_records 
    $ad_current_records = Get-ADUser -filter * -Properties objectguid  -Server $global:Config.Configurations.'target domain controller' | Select-Object -ExpandProperty objectguid
    $users_summary.retrieved = $ad_current_records.Count
    $users_summary.database = $sql_current_records.Count

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

            $users_summary.deleted++
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_sql_record.ad_samaccountname
            $this_record_item.action = "deleted"
            $users_target_item_array += $this_record_item


        }
    }
    if ( $flag_deleted -eq 0) {
        Global:log -text ("No deleted user record found." ) -Hierarchy "Main:Users:deleted detect" -type warning

    }
    #endregion


    #region missing records
    if ( $sql_current_records.Count -lt $ad_current_records.Count ) {
        Global:log -text ("sql_current_records.Count < ad_current_records.Count, {0} missing records in database.... " -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Users:missing records" -type warning
        $existing_sql_objects_guid = ( $sql_current_records | Select-Object -ExpandProperty ad_objectguid )
        $users_missing = @()
        $users_missing_count = 0
        $ad_current_records | ForEach-Object {
            $this_ad_record = $_
            #Global:log -text ("checking guid={0}" -F $this_ad_record.guid) -Hierarchy "Main:Ou:missing records" 
            if ( $existing_sql_objects_guid -notcontains $this_ad_record.guid -and $users_missing_count -lt $global:Config.Configurations.inventory.'max missing records') {
                $users_missing_count++
                Global:log -text ("missing object guid={0}" -F $this_ad_record.guid ) -Hierarchy "Main:Ou:missing records" 
                $filter = "objectguid -eq  '{0}'" -F $this_ad_record.guid
                Get-ADUser  -properties $global:config.Configurations.inventory.'Users Active Directory Attributes' -filter $filter  -Server $global:Config.Configurations.'target domain controller' | ForEach-Object { 
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
    
                    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, surname, enabled, objectclass, name)
                    $users_missing += $this_calculated_row

                }

            }
        
        }
        if ( $users_missing_count -eq $global:Config.Configurations.inventory.'max missing records' ) {
            Global:log -text ("Max missing records count reached ({0})" -F $global:Config.Configurations.inventory.'max missing records' ) -Hierarchy "Main:Users:missing records" -type warning
        }

        $users_missing | ForEach-Object {
            $this_user = $_
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_user.SamAccountName
            switch ( Global:ADTidy_Inventory_Users_sql_update -Fields $this_user) {
                "update" {
                    $this_record_item.action = "updated"
                    $users_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $users_summary.created++

                }
            }
        

            $users_target_item_array += $this_record_item
        }
    }
    else {
        Global:log -text ("sql_current_records.Count = ad_current_records.Count, no missing records in database" -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Users:missing records"
    }
    #endregion

    $users_record.result_summary = $users_summary | ConvertTo-Json -Compress
    $users_record.target_list = $users_target_item_array | ConvertTo-Json -Compress

    Global:ADTidy_Records_sql_update -Fields $users_record
}
#endregion

#region computers
if ($flag_inventory_computers -eq $true ) {
    Global:ADTidy_Inventory_Computers_sql_table_check

    #region record init
    $computers_record = $record_template | Select-Object *
    $computers_record.record_type = "AdTidy.inventory"
    $computers_record.rule_name = "computers"
    $computers_record.target_list = @()


    $computers_summary = $summary_template | Select-Object *
    $computers_summary.database = 0
    $computers_summary.retrieved = 0
    $computers_summary.updated = 0
    $computers_summary.created = 0
    $computers_summary.deleted = 0

    $computers_record.result_summary = $computers_summary | ConvertTo-Json -Compress

    $computers_target_item_array = @()
    #endregion

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
    Get-ADComputer -properties $global:config.Configurations.inventory.'Computers Active Directory Attributes' -filter $filter -Server $global:Config.Configurations.'target domain controller' | ForEach-Object { 
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
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_computer.samaccountname
            switch ( Global:ADTidy_Inventory_Computers_sql_update -Fields $this_computer) {
                "update" {
                    $this_record_item.action = "updated"
                    $computers_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $computers_summary.created++
                }
            }
            $computers_target_item_array += $this_record_item
        }

    }
    #endregion

    #region deleted records detect
    $sql_current_records = Global_ADTidy_Iventory_Computers_all_current_records 
    $ad_current_records = Get-Adcomputer -filter * -Properties objectguid  -Server $global:Config.Configurations.'target domain controller' | Select-Object -ExpandProperty objectguid
    $computers_summary.retrieved = $ad_current_records.Count
    $computers_summary.database = $sql_current_records.Count

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
            $computers_summary.deleted++
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_sql_record.ad_samaccountname
            $this_record_item.action = "deleted"
            $computers_target_item_array += $this_record_item


        }
    }
    if ( $flag_deleted -eq 0) {
        Global:log -text ("No deleted computer record found." ) -Hierarchy "Main:Computers:deleted detect" -type warning

    }
    #endregion

    #region missing records
    if ( $sql_current_records.Count -lt $ad_current_records.Count ) {
        Global:log -text ("sql_current_records.Count < ad_current_records.Count, {0} missing records in database.... " -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Computers:missing records" -type warning
        $existing_sql_objects_guid = ( $sql_current_records | Select-Object -ExpandProperty ad_objectguid )
        $computers_missing = @()
        $computers_missing_count = 0
        $ad_current_records | ForEach-Object {
            $this_ad_record = $_
            #Global:log -text ("checking guid={0}" -F $this_ad_record.guid) -Hierarchy "Main:Ou:missing records" 
            if ( $existing_sql_objects_guid -notcontains $this_ad_record.guid -and $computers_missing_count -lt $global:Config.Configurations.inventory.'max missing records') {
                $computers_missing_count++
                Global:log -text ("missing object guid={0}" -F $this_ad_record.guid ) -Hierarchy "Main:Computers:missing records" 
                $filter = "objectguid -eq  '{0}'" -F $this_ad_record.guid
                Get-ADComputer -properties $global:config.Configurations.inventory.'Computers Active Directory Attributes' -filter $filter  -Server $global:Config.Configurations.'target domain controller' | ForEach-Object { 
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
    
                    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, surname, enabled, objectclass, UserPrincipalName, DNSHostName )
                    $computers_missing += $this_calculated_row

                }

            }
        
        }
        if ( $computers_missing_count -eq $global:Config.Configurations.inventory.'max missing records' ) {
            Global:log -text ("Max missing records count reached ({0})" -F $global:Config.Configurations.inventory.'max missing records' ) -Hierarchy "Main:Computers:missing records" -type warning
        }

        $computers_missing | ForEach-Object {
            $this_computer = $_
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_computer.SamAccountName
            switch ( Global:ADTidy_Inventory_Computers_sql_update -Fields $this_computer) {
                "update" {
                    $this_record_item.action = "updated"
                    $computers_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $computers_summary.created++

                }
            }
        

            $computers_target_item_array += $this_record_item
        }
    }
    else {
        Global:log -text ("sql_current_records.Count = ad_current_records.Count, no missing records in database" -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Computers:missing records"
    }
    #endregion
    $computers_record.result_summary = $computers_summary | ConvertTo-Json -Compress
    $computers_record.target_list = $computers_target_item_array | ConvertTo-Json -Compress
    Global:ADTidy_Records_sql_update -Fields $computers_record
}
#endregion

#region Groups
if ($flag_inventory_groups -eq $true ) {
    Global:ADTidy_Inventory_Groups_sql_table_check
    #region record init
    $groups_record = $record_template | Select-Object *
    $groups_record.record_type = "AdTidy.inventory"
    $groups_record.rule_name = "groups"
    $groups_record.target_list = @()



    $groups_summary = $summary_template | Select-Object *
    $groups_summary.database = 0
    $groups_summary.retrieved = 0
    $groups_summary.updated = 0
    $groups_summary.created = 0
    $groups_summary.deleted = 0

    $groups_record.result_summary = $groups_summary | ConvertTo-Json -Compress

    $groups_target_item_array = @()
    #endregion

    $last_update = Global_ADTidy_Iventory_Groups_last_update

    if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
        $filter = "*"
    }
    else {
    
        $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
        $filter = "whenchanged -gt '$filter_date'"
    }
    Global:log -text ("retrieving groups from AD, filter='{0}'" -F $filter) -Hierarchy "Main:Groups"
    $groups = @()


    #region delta changes
    <# PRD#>
    Get-ADGroup -properties $global:config.Configurations.inventory.'Groups Active Directory Attributes' -filter $filter -Server $global:Config.Configurations.'target domain controller' | Sort-Object whenchanged  | ForEach-Object { 
        <# DEV 
Get-ADUser  -properties $global:config.Configurations.inventory.'Active Directory Attributes' -filter "samaccountname -eq 'har3005'"  | ForEach-Object { #> 


        $this_row = $_
        $this_calculated_row = "" | Select-Object ignore

        $this_row | Get-Member | Where-Object { $_.membertype -eq "property" } | Select-Object -ExpandProperty name | ForEach-Object {
        
            $this_attribute = $_
            $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
            Switch ( $this_attribute ) {
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
        #region members and nested members (recursive)
        $members_array = @()
        Get-ADGroupMember -Identity $this_row.name -Server $global:Config.Configurations.'target domain controller' | ForEach-Object {
            $line = "" | Select-Object distinguishedname, membership
            $line.membership = 'direct'
            $line.distinguishedName = $_.distinguishedname
            if (( $members_array | Select-Object -ExpandProperty distinguishedName ) -notcontains $_.distinguishedname ) {
                $members_array += $line
            }
        }
        Get-ADGroupMember -Identity $this_row.name -Recursive -Server $global:Config.Configurations.'target domain controller' | ForEach-Object {
            $line = "" | Select-Object distinguishedname, membership
            $line.membership = 'nested'
            $line.distinguishedName = $_.distinguishedname
            if (( $members_array | Select-Object -ExpandProperty distinguishedName ) -notcontains $_.distinguishedname ) {
                $members_array += $line
            }
        }
        $this_calculated_row = $this_calculated_row | Select-Object *, "xml_members"
        if ( $members_array.count -ne 0) {
            $this_calculated_row."xml_members" = Global:ConvertTo-SimplifiedXML -InputObject $members_array -RootNodeName "Members" -NodeName "Member"
        }
        else {
            $this_calculated_row."xml_members" = $null
        }

        #endregion
    
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
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_group.samaccountname
            switch ( Global:ADTidy_Inventory_Groups_sql_update -Fields $this_group) {
                "update" {
                    $this_record_item.action = "updated"
                    $groups_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $groups_summary.created++
                }
            }
            $groups_target_item_array += $this_record_item

        
        }
    }
    #endregion


    #region deleted records detect
    $sql_current_records = Global_ADTidy_Iventory_Groups_all_current_records 
    $ad_current_records = Get-ADGroup -filter * -Properties objectguid -Server $global:Config.Configurations.'target domain controller' | Select-Object -ExpandProperty objectguid
    $groups_summary.retrieved = $ad_current_records.Count
    $groups_summary.database = $sql_current_records.Count
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

            $groups_summary.deleted++
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_sql_record.ad_name
            $this_record_item.action = "deleted"
            $groups_target_item_array += $this_record_item


        }
    }
    if ( $flag_deleted -eq 0) {
        Global:log -text ("No deleted group record found." ) -Hierarchy "Main:Groups:deleted detect" -type warning

    }
    #endregion


    #region missing records
    if ( $sql_current_records.Count -lt $ad_current_records.Count ) {
        Global:log -text ("sql_current_records.Count < ad_current_records.Count, {0} missing records in database.... " -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Groups:missing records" -type warning
        $existing_sql_objects_guid = ( $sql_current_records | Select-Object -ExpandProperty ad_objectguid )
        $groups_missing = @()
        $groups_missing_count = 0
        $ad_current_records | ForEach-Object {
            $this_ad_record = $_
            #Global:log -text ("checking guid={0}" -F $this_ad_record.guid) -Hierarchy "Main:Ou:missing records" 
            if ( $existing_sql_objects_guid -notcontains $this_ad_record.guid -and $computers_missing_count -lt $global:Config.Configurations.inventory.'max missing records') {
                $groups_missing_count++
                Global:log -text ("missing object guid={0}" -F $this_ad_record.guid ) -Hierarchy "Main:Groups:missing records" 
                $filter = "objectguid -eq  '{0}'" -F $this_ad_record.guid
                Get-ADGroup -properties $global:config.Configurations.inventory.'Groups Active Directory Attributes' -filter $filter -Server $global:Config.Configurations.'target domain controller' | ForEach-Object { 
                    $this_row = $_
                    $this_calculated_row = "" | Select-Object ignore

                    #region members and nested members (recursive)
                    $members_array = @()
                    Get-ADGroupMember -Identity $this_row.name -Server $global:Config.Configurations.'target domain controller' | ForEach-Object {
                        $line = "" | Select-Object distinguishedname, membership
                        $line.membership = 'direct'
                        $line.distinguishedName = $_.distinguishedname
                        if (( $members_array | Select-Object -ExpandProperty distinguishedName ) -notcontains $_.distinguishedname ) {
                            $members_array += $line
                        }
                    }
                    Get-ADGroupMember - -Identity $this_row.name -Recursive -Server $global:Config.Configurations.'target domain controller' | ForEach-Object {
                        $line = "" | Select-Object distinguishedname, membership
                        $line.membership = 'nested'
                        $line.distinguishedName = $_.distinguishedname
                        if (( $members_array | Select-Object -ExpandProperty distinguishedName ) -notcontains $_.distinguishedname ) {
                            $members_array += $line
                        }
                    }
                    $this_calculated_row = $this_calculated_row | Select-Object *, "xml_members"
                    if ( $members_array.count -ne 0) {
                        $this_calculated_row."xml_members" = Global:ConvertTo-SimplifiedXML -InputObject $members_array -RootNodeName "Members" -NodeName "Member"
                    }
                    else {
                        $this_calculated_row."xml_members" = $null
                    }

                    #endregion

                    $this_row | Get-Member | Where-Object { $_.membertype -eq "property" } | Select-Object -ExpandProperty name | ForEach-Object {
                        $this_attribute = $_
                        $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
                        Switch ( $this_attribute ) {
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
    
                    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore, surname, enabled, objectclass, UserPrincipalName, DNSHostName )
                    $groups_missing += $this_calculated_row

                }

            }
        
        }
        if ( $groups_missing_count -eq $global:Config.Configurations.inventory.'max missing records' ) {
            Global:log -text ("Max missing records count reached ({0})" -F $global:Config.Configurations.inventory.'max missing records' ) -Hierarchy "Main:Groups:missing records" -type warning
        }

        $groups_missing | ForEach-Object {
            $this_group = $_
            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_group.SamAccountName
            switch ( Global:ADTidy_Inventory_Groups_sql_update -Fields $this_group) {
                "update" {
                    $this_record_item.action = "updated"
                    $groups_summary.updated++
                }
                "new" {
                    $this_record_item.action = "created"
                    $groups_summary.created++

                }
            }
        

            $groups_target_item_array += $this_record_item
        }
    }
    else {
        Global:log -text ("sql_current_records.Count = ad_current_records.Count, no missing records in database" -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy "Main:Groups:missing records"
    }
    #endregion

    $groups_record.result_summary = $groups_summary | ConvertTo-Json -Compress
    $groups_record.target_list = $groups_target_item_array | ConvertTo-Json -Compress
    Global:ADTidy_Records_sql_update -Fields $groups_record
}
#endregion
