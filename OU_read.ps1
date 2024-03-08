Clear-Host
#region ## script information
$Global:Version = "1.0.0"
# HAR3005, Primeo-Energie, 20240308
#
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

#region record management, init and templates
$record_template = "" | Select-Object record_type, rule_name, target_list, result_summary
$summary_template = "" | Select-Object database, retrieved, updated, created, deleted
$target_item_template = "" | Select-Object name, action
#endregion

$objects_loop_config = @(
    [pscustomobject]@{
        object_type                   = 'OU'
        enabled                       = $True
        function_check_table          = "Global:ADTidy_Inventory_OU_sql_table_check"
        function_last_update          = "Global_ADTidy_Iventory_OU_last_update"
        function_current_sql_records  = "Global_ADTidy_Iventory_OU_all_current_records"
        function_active_directory_get = "Get-ADOrganizationalUnit"
        function_sql_update           = "Global:ADTidy_Inventory_OU_sql_update"
    }
)

#region script specific functions
function Attribute_formatter {
    # formatts input $value based on the attribute $attribute_name
    param(
        [Parameter(Mandatory = $true)] $attribute,
        [Parameter(Mandatory = $true)] $value
    )
    Switch ( $attribute ) {
        "accountexpires" {
            # Users
            $pwdLastSetRaw = [string]($value)
            if ( $pwdLastSetRaw -eq "9223372036854775807" ) { 
                $pwdLastSet = $null
            }
            else {
                $pwdLastSet = [datetime]::FromFileTime($pwdLastSetRaw).ToString("yyyy-MM-dd HH:mm:ss") 
                if ( $pwdLastSet -eq '1601-01-01 01:00:00' ) { $pwdLastSet = $null }
            }
            if ( $pwdLastSet.Length -eq 0 ) { $returned_value = "NULL" }
            ELSE { $returned_value = "$pwdLastSet" }

        }
        "lastLogonTimestamp" {
            # Users / Computers
            $pwdLastSetRaw = [string]($value)
            if ( $pwdLastSetRaw -eq "9223372036854775807" ) { 
                $pwdLastSet = $null
            }
            else {
                $pwdLastSet = [datetime]::FromFileTime($pwdLastSetRaw).ToString("yyyy-MM-dd HH:mm:ss") 
                if ( $pwdLastSet -eq '1601-01-01 01:00:00' ) { $pwdLastSet = $null }
            }
            if ( $pwdLastSet.Length -eq 0 ) { $returned_value = "NULL" }
            ELSE { $returned_value = "$pwdLastSet" }
        }
        "pwdLastSet" {
            # Users
            $pwdLastSetRaw = [string]($value)
            if ( $pwdLastSetRaw -eq "9223372036854775807" ) { 
                $pwdLastSet = $null
            }
            else {
                $pwdLastSet = [datetime]::FromFileTime($pwdLastSetRaw).ToString("yyyy-MM-dd HH:mm:ss") 
                if ( $pwdLastSet -eq '1601-01-01 01:00:00' ) { $pwdLastSet = $null }
            }
            if ( $pwdLastSet.Length -eq 0 ) { $returned_value = "NULL" }
            ELSE { $returned_value = "$pwdLastSet" }
        }
        "useraccountcontrol" {
            # Users
            $returned_value = Global:DecodeUserAccountControl ([int][string]($value))
        }
        "thumbnailPhoto" {
            # Users
            if ( $value -ne $null) {
                $returned_value = "Is set"
            }
            else {
                $returned_value = "NULL"
            }
        }
        "businesscategory" {
            # OU
            TRY {
                $businesscategory_string = ""
                $value | ForEach-Object {
                    $businesscategory_string = "{0}{1}," -F $businesscategory_string, $_
                }
                # remove last ',' from string
                $businesscategory_string = $businesscategory_string -replace ".$"
                $returned_value = $businesscategory_string
                if (([string]$value).length -eq 0 ) { $returned_value = "NULL" }
            }
            CATCH {

            }
        }
        "whenChanged" {
            TRY {
                $CalulatedValue = '{0:yyyy-MM-dd HH:mm:ss}' -f $value
                $returned_value = $CalulatedValue 
            }
            CATCH {
                $returned_value = $null
            }
                    
                
                
        }
        "whenCreated" {
            TRY {
                $CalulatedValue = '{0:yyyy-MM-dd HH:mm:ss}' -f $value
                $returned_value = $CalulatedValue 
            }
            CATCH {
                $returned_value = $null
            }

                
        }
        default {
            $returned_value = $value -replace "'", "''" 
            if (([string]$value).length -eq 0 ) { $returned_value = "NULL" }
        }

    }
    return $returned_value
}
#endregion

$attributes = @(
    "distinguishedname",
    "businessCategory",
    "name"
)

$ou = @(
    "OU=__Stagging,DC=pe,DC=ch",
    "OU=0_Identity_Sytems,DC=pe,DC=ch",
    "OU=1_Application_and_Services,DC=pe,DC=ch",
    "OU=2_Users_and_Devices,DC=pe,DC=ch"
)

$ou_array = @()
$ou | ForEach-Object {
    $this = $_
    $tier = ($this.split("=")[1]).substring(0, 1)
    write-host $tier
    Get-ADOrganizationalUnit -Filter * -Properties $attributes -SearchBase $this | Select-Object $attributes | ForEach-Object {
        write-host "." -NoNewline
        $this_ou = $_ | Select-Object *, level, tier, groupOU, calculated_taggs
        $this_ou.level = ($this_ou.distinguishedName).split(",").count - 2
        $this_ou.tier = $tier

        # Set-ADOrganizationalUnit -Identity $this_ou.distinguishedName -Clear businessCategory

        if ( ($this_ou.distinguishedName) -match "ou=groups,") {
            $this_ou.groupOU = "yes"
        }
        $this_ou.calculated_taggs = @()
        if ( $this_ou.distinguishedName -match "ou=applications," ) { $this_ou.calculated_taggs += "Applications" }
        if ( $this_ou.distinguishedName -match "ou=tier" ) { $this_ou.calculated_taggs += "Tier" }
        if ( $this_ou.distinguishedName -match "ou=organisation") { $this_ou.calculated_taggs += "Organisation" }
        if ( $this_ou.distinguishedName -match "ou=delegation" ) { $this_ou.calculated_taggs += "Delegation" }
        if ( $this_ou.distinguishedName -match "ou=ou" ) { $this_ou.calculated_taggs += "ou" }
        if ( $this_ou.distinguishedName -match "ou=it" ) { $this_ou.calculated_taggs += "it" }
        if ( $this_ou.distinguishedName -match "ou=hr" ) { $this_ou.calculated_taggs += "hr" }
        if ( $this_ou.distinguishedName -match "ou=infra" ) { $this_ou.calculated_taggs += "Infrastructure" }
        if ( $this_ou.distinguishedName -match "ou=gpo" ) { $this_ou.calculated_taggs += "gpo" }
        if ( $this_ou.distinguishedName -match "ou=active-directory" ) { $this_ou.calculated_taggs += "active-directory" }
        if ( $this_ou.distinguishedName -match "OU=pki" ) { $this_ou.calculated_taggs += "pki" }
        if ( $this_ou.distinguishedName -match "OU=network" ) { $this_ou.calculated_taggs += "Network" }
        if ( $this_ou.distinguishedName -match "OU=servers,OU=delegation," ) { $this_ou.calculated_taggs += "Servers" }
        if ( $this_ou.distinguishedName -match "ou=deployment" ) { $this_ou.calculated_taggs += "Deployement" }
        if ( $this_ou.distinguishedName -match "ou=file-server" ) { $this_ou.calculated_taggs += "File-server" }
        if ( $this_ou.distinguishedName -match "ou=group-managed" ) { $this_ou.calculated_taggs += "Group managed service accounts" }
        if ( $this_ou.distinguishedName -match "ou=group-policies" ) { $this_ou.calculated_taggs += "Group policies" }
        if ( $this_ou.distinguishedName -match "ou=collaboration" ) { $this_ou.calculated_taggs += "Collaboration" }
        if ( $this_ou.distinguishedName -match "ou=monitoring" ) { $this_ou.calculated_taggs += "Monitoring" }
        if ( $this_ou.distinguishedName -match "ou=mobile-device-management" ) { $this_ou.calculated_taggs += "mobile-device-management" }
        if ( $this_ou.distinguishedName -match "ou=sql," ) { $this_ou.calculated_taggs += "SQL" }
        if ( $this_ou.distinguishedName -match "ou=license," ) { $this_ou.calculated_taggs += "License" }
        if ( $this_ou.distinguishedName -match "ou=systems," ) { $this_ou.calculated_taggs += "Systems" }
        if ( $this_ou.distinguishedName -match "ou=distribution," ) { $this_ou.calculated_taggs += "Distribution" }

        #region check businesscategories
        $this_ou.calculated_taggs | ForEach-Object {
            $this_tag = $_
            $flag_tag_found = $false
            if ( $this_ou.businessCategory -contains $this_tag ) {
                $flag_tag_found = $true
            }

            if ( $flag_tag_found -eq $false) {
                Set-ADOrganizationalUnit -Identity $this_ou.distinguishedName -Add @{businessCategory = "$this_tag" }
            }
        }
        #endregion

        $ou_array += $this_ou

    }

}
$ou_array | Format-Table -auto

exit
#region objects_loop_config handler
$objects_loop_config | Where-Object { $_.enabled -eq $true } | ForEach-Object {
    $this_iteration_config = $_
    Global:log -text ("Starting because enabled = 'true'" -F $filter) -Hierarchy ("Main:{0}" -F $this_iteration_config.object_type)

    #region database table verification
    Global:log -text (" > checking database table..." ) -Hierarchy ("Main:{0}" -F $this_iteration_config.object_type) -type warning
    & $this_iteration_config.function_check_table
    #endregion

    #region record init
    $record = $record_template | Select-Object *
    $record.record_type = "AdTidy.inventory"
    $record.rule_name = $this_iteration_config.object_type
    $record.target_list = @()


    $summary = $summary_template | Select-Object *
    $summary.database = 0
    $summary.retrieved = 0
    $summary.updated = 0
    $summary.created = 0
    $summary.deleted = 0

    $record.result_summary = $ou_summary | ConvertTo-Json -Compress

    $target_item_array = @()
    #endregion

    #region database max whenchanged attribute
    $last_update = & $this_iteration_config.function_last_update

    if ( ([string]$last_update.maxrecord).Length -eq 0 ) {
        $filter = "*"
        $flag_first_run = $true
    }
    else {
        $flag_first_run = $false
        $filter_date = get-date $last_update.maxrecord  | ForEach-Object touniversaltime | get-date -format yyyyMMddHHmmss.0Z
        $filter = "whenchanged -gt '$filter_date'"
    }
    Global:log -text ("Retrieving records from AD using filter='{0}'" -F $filter) -Hierarchy ("Main:{0}" -F $this_iteration_config.object_type)
    #endregion

    #region delta changes
    #region # evaluate amount of records matching delta filter
    #& $this_iteration_config.function_active_directory_get -Filter $filter -Properties $global:Config.Configurations.inventory.attributes."$($this_iteration_config.object_type)" -Server $global:Config.Configurations.'target domain controller' | Sort-Object whenChanged | ForEach-Object { 
    $objects_matched = & $this_iteration_config.function_active_directory_get -Filter $filter -Properties objectguid, whenChanged -Server $global:Config.Configurations.'target domain controller' | Select-Object objectguid, whenChanged
    if ( $objects_matched.count -gt $global:Config.Configurations.inventory.'max insert limit' ) {
        Global:log -text ("Amount of matching records ({0}) above defined limite ({1}), restricting items processed to limit" -F $objects_matched.count, $global:Config.Configurations.inventory.'max insert limit') -Hierarchy ("Main:{0}:Delta" -F $this_iteration_config.object_type) -type warning
        $objects_to_process = $objects_matched | Sort-Object whenChanged  | Select-Object -first $global:Config.Configurations.inventory.'max insert limit'
        $flag_chunked = $true
    }
    else {
        Global:log -text ("Amount of matching records ({0}) below defined limit ({1}), processing all" -F $objects_matched.count, $global:Config.Configurations.inventory.'max insert limit') -Hierarchy ("Main:{0}:Missing objects" -F $this_iteration_config.object_type) 
        $objects_to_process = $objects_matched
        $flag_chunked = $false
    }
    #endregion

    #region active directory attribute read for selected chunk 
    $AD_objects = @()
    $objects_to_process | ForEach-Object {
        $this_object_to_process = $_
        $filter = "objectguid -eq '{0}'" -F $this_object_to_process.objectguid
        $this_row = & $this_iteration_config.function_active_directory_get -Filter $filter -Properties $global:Config.Configurations.inventory.attributes."$($this_iteration_config.object_type)" -Server $global:Config.Configurations.'target domain controller' | Select-Object $global:Config.Configurations.inventory.attributes."$($this_iteration_config.object_type)"
        $this_calculated_row = "" | Select-Object ignore

        $this_row | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
            $this_attribute = $_
            $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
            if ( $this_row."$this_attribute" -eq $null) {
                $this_calculated_row."$this_attribute" = "NULL"    
            }
            else {
                $this_calculated_row."$this_attribute" = Attribute_formatter -attribute $this_attribute -value $this_row."$this_attribute"
            }
        }
        
        if ( $this_iteration_config.object_type -eq "Groups") {
            #region members and nested members (recursive)
            Global:log -text ("Direct..." -F $filter) -Hierarchy ("Main:{0}:Delta:Members" -F $this_iteration_config.object_type) 
            $members_array = @()
            Get-ADGroupMember -Identity $this_row.name -Server $global:Config.Configurations.'target domain controller' | ForEach-Object {
                $line = "" | Select-Object distinguishedname, membership
                $line.membership = 'direct'
                $line.distinguishedName = $_.distinguishedname
                if (( $members_array | Select-Object -ExpandProperty distinguishedName ) -notcontains $_.distinguishedname ) {
                    $members_array += $line
                }
            }
            Global:log -text ("Nested..." -F $filter) -Hierarchy ("Main:{0}:Delta:Members" -F $this_iteration_config.object_type) 
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
        }

        $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore)

        $this_record_item = $target_item_template | Select-Object *
        switch ($this_iteration_config.object_type) {
            "OU" { $this_record_item.name = $this_calculated_row.distinguishedname }
            default { $this_record_item.name = $this_calculated_row.samaccountname }
        }
        
        switch ( & $this_iteration_config.function_sql_update -Fields $this_calculated_row) {
            "update" {
                $this_record_item.action = "updated"
                $summary.updated++
            }
            "new" {
                $this_record_item.action = "created"
                $summary.created++

            }
        }
        $target_item_array += $this_record_item
    }
    #endregion
    #endregion

    #region deleted records detect
    Global:log -text (" > retrieving current SQL records..." ) -Hierarchy ("Main:{0}:deleted records" -F $this_iteration_config.object_type)

    $sql_current_records = & $this_iteration_config.function_current_sql_records
    $ad_current_records = & $this_iteration_config.function_active_directory_get -filter * -Properties objectguid  -Server $global:Config.Configurations.'target domain controller' | Select-Object -ExpandProperty objectguid
    $summary.retrieved = $ad_current_records.Count
    $summary.database = $sql_current_records.Count
    Global:log -text ("SQL:{0} current records, AD:{1} current records " -F $sql_current_records.Count, $ad_current_records.Count) -Hierarchy ("Main:{0}:deleted records" -F $this_iteration_config.object_type)

    $flag_object_deleted = $false
    $sql_current_records | ForEach-Object {
        $this_sql_record = $_
        if ( $ad_current_records -notcontains $this_sql_record.ad_objectguid) {
            Global:log -text ("Detected a deleted object:'{0}' " -F ($this_sql_record | Select-Object ad_name, ad_objectguid, ad_dinstinguishedname | ConvertTo-Json -Compress)) -Hierarchy ("Main:{0}:deleted records" -F $this_iteration_config.object_type)
            $delete_record = $this_sql_record | Select-Object @{name = 'Objectguid'; expression = { $_.ad_ObjectGUID } }, record_status
            $delete_record.record_status = "Deleted"
            & $this_iteration_config.function_sql_update -Fields $delete_record
            $summary.deleted++
            $flag_object_deleted = $true

            $this_record_item = $target_item_template | Select-Object *
            $this_record_item.name = $this_sql_record.ad_distinguishedname
            $this_record_item.action = "deleted"
            $target_item_array += $this_record_item


        }
    }
    if ( $flag_object_deleted -eq $false) {
        Global:log -text ("No deleted record found." ) -Hierarchy ("Main:{0}:deleted records" -F $this_iteration_config.object_type) -type warning

    }
    #endregion

    #region missing records
    if ( $flag_chunked -eq $false) {
        if ( $sql_current_records.Count -lt $ad_current_records.Count ) {
            Global:log -text ("sql_current_records.Count < ad_current_records.Count, {0} missing records in database.... " -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy ("Main:{0}:Delta" -F $this_iteration_config.object_type) -type warning 
            $existing_sql_objects_guid = ( $sql_current_records | Select-Object -ExpandProperty ad_objectguid )
            $Organizational_units_missing = @()
            $Organizational_units_missing_count = 0
            $ad_current_records | ForEach-Object {
                $this_ad_record = $_
                if ( $existing_sql_objects_guid -notcontains $this_ad_record.guid -and $Organizational_units_missing_count -lt $global:Config.Configurations.inventory.'max missing records') {
                    $Organizational_units_missing_count++
                    Global:log -text ("missing object guid={0}" -F $this_ad_record.guid, $this_ad_record.distinguishedName ) -Hierarchy ("Main:{0}:Delta" -F $this_iteration_config.object_type)
                    $filter = "objectguid -eq  '{0}'" -F $this_ad_record.guid
                    $this_row = & $this_iteration_config.function_active_directory_get -Filter $filter -Properties $global:Config.Configurations.inventory.attributes."$($this_iteration_config.object_type)" -Server $global:Config.Configurations.'target domain controller' | Select-Object $global:Config.Configurations.inventory.attributes."$($this_iteration_config.object_type)"
                    $this_calculated_row = "" | Select-Object ignore

                    $this_row | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
                        $this_attribute = $_
                        $this_calculated_row = $this_calculated_row | Select-Object *, $this_attribute
                        if ( $this_row."$this_attribute" -eq $null) {
                            $this_calculated_row."$this_attribute" = "NULL"    
                        }
                        else {
                            $this_calculated_row."$this_attribute" = Attribute_formatter -attribute $this_attribute -value $this_row."$this_attribute"
                        }
                    }
    
                    if ( $this_iteration_config.object_type -eq "Groups") {
                        #region members and nested members (recursive)
                        Global:log -text ("Direct..." -F $filter) -Hierarchy ("Main:{0}:Delta:Members" -F $this_iteration_config.object_type) 
                        $members_array = @()
                        Get-ADGroupMember -Identity $this_row.name -Server $global:Config.Configurations.'target domain controller' | ForEach-Object {
                            $line = "" | Select-Object distinguishedname, membership
                            $line.membership = 'direct'
                            $line.distinguishedName = $_.distinguishedname
                            if (( $members_array | Select-Object -ExpandProperty distinguishedName ) -notcontains $_.distinguishedname ) {
                                $members_array += $line
                            }
                        }
                        Global:log -text ("Nested..." -F $filter) -Hierarchy ("Main:{0}:Delta:Members" -F $this_iteration_config.object_type) 
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
                    }
                    $this_calculated_row = ($this_calculated_row | Select-Object * -ExcludeProperty ignore)

                    $this_record_item = $target_item_template | Select-Object *
                    switch ($this_iteration_config.object_type) {
                        "OU" { $this_record_item.name = $this_calculated_row.distinguishedname }
                        default { $this_record_item.name = $this_calculated_row.samaccountname }
                    }
        
                    switch ( & $this_iteration_config.function_sql_update -Fields $this_calculated_row) {
                        "update" {
                            $this_record_item.action = "updated"
                            $summary.updated++
                        }
                        "new" {
                            $this_record_item.action = "created"
                            $summary.created++

                        }
                    }
                    $target_item_array += $this_record_item           
                }
        


            }
        }   
        else {
        
            Global:log -text ("sql_current_records.Count = ad_current_records.Count, no missing records in database" -F ($ad_current_records.Count - $sql_current_records.Count) ) -Hierarchy ("Main:{0}:Missing objects" -F $this_iteration_config.object_type) 
        }
    }
    else {
        Global:log -text (" flag_checked -eq 'true', missing objects skipped." ) -Hierarchy ("Main:{0}:Missing objects" -F $this_iteration_config.object_type)         
    }
    #endregion

    $record.result_summary = $summary | ConvertTo-Json -Compress
    $record.target_list = $target_item_array | ConvertTo-Json -Compress

    Global:ADTidy_Records_sql_update -Fields $record
}

#endregion
#endregion