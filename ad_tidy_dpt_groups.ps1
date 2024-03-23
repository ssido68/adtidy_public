Clear-Host
#region ## script information
$Global:Version = "0.0.1"
# HAR3005, Primeo-Energie, 20240209
#    Initial draft
$Global:Version = "1.0.1"
# HAR3005, Primeo-Energie, 20240303
#    pre-HR/IT
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

#region # Main 
Global:log -text ("Start V{0}" -F $Global:Version) -Hierarchy "Main"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
# purge temp storage file
$Global:ConfigPurgeTempFilesFilter = @("*.temp.csv", "*.txt")
$Global:ConfigPurgeTempFilesFilter | ForEach-Object {
    $ThisFilter = $_
    Global:log -text (" # remove temp storage files '{0}'" -F $ThisFilter) -Hierarchy "Main" -type warning
    Remove-Item $ThisFilter
}


#region retrieve configuration: target segment groups OU
$global:segments_ou = $global:Config.Configurations.'company groups'.segments_groups_ou
$global:hr_departments_ou = $global:Config.Configurations.'company groups'.hr_departments_groups_ou
$global:it_departments_ou = $global:Config.Configurations.'company groups'.it_departments_groups_ou
#endregion

#region get target OU attributes, looking for manager information
#region segment OU information loading
$global:segments_ou_adobject = Get-ADOrganizationalUnit -Identity $global:segments_ou -Properties managedBy, businessCategory
$global:segments_ou_manager = $global:segments_ou_adobject.ManagedBy

$segment_ou_hashtags = @()
$global:segments_ou_adobject.businessCategory | ForEach-Object {
    $this_category = $_
    switch ( $this_category ) {
        "Organisation" { $segment_ou_hashtags += "#Orga" }
        default { $segment_ou_hashtags += "#{0}" -F $this_category }
    }


}
#endregion

#region hr group OU information loading
$global:hr_department_ou_adobject = Get-ADOrganizationalUnit -Identity $global:hr_departments_ou -Properties managedBy, businessCategory
$global:hr_department_ou_manager = $global:hr_department_ou_adobject.ManagedBy
$hr_ou_hashtags = @()
$global:hr_department_ou_adobject.businessCategory | ForEach-Object {
    $this_category = $_
    switch ( $this_category ) {
        "Organisation" { $hr_ou_hashtags += "#Orga" }
        default { $hr_ou_hashtags += "#{0}" -F $this_category }
    }
}
#endregion

#region it group OU information loading
$global:it_department_ou_adobject = Get-ADOrganizationalUnit -Identity $global:it_departments_ou -Properties managedBy, businessCategory
$global:it_department_ou_manager = $global:it_department_ou_adobject.ManagedBy
$it_ou_hashtags = @()
$global:it_department_ou_adobject.businessCategory | ForEach-Object {
    $this_category = $_
    switch ( $this_category ) {
        "Organisation" { $it_ou_hashtags += "#Orga" }
        default { $it_ou_hashtags += "#{0}" -F $this_category }
    }
}
#endregion


$global:group_security = 'Security'
$global:group_scope = 'Universal'
#endregion

#region get ad user that have employeeid
$raw_user_data = Get-ADUser -LDAPFilter "(employeeID=*)" -Properties employeeid, department, company, distinguishedname
$raw_data_segments_groupby = $raw_user_data | Group-Object company
#endregion

#region status row array templates
$status_log_row_template = "" | Select-Object group_type, group_name, group_distinguishedname, flag_exists, flag_membership_changed, flag_reporting, action_logs, action_type
$status_log_action_log_row_template = "" | Select-Object timestamp, type, target, result
#endregion

#region pre-loop-logs
Global:log -text ("Employee records count: {0}" -F $raw_user_data.count) -Hierarchy "Main:loop configuration"

Global:log -text ("`$global:segments_ou_manager {0}" -F $global:segments_ou_manager) -Hierarchy "Main:loop configuration"
Global:log -text ("`$segment_ou_hashtags: {0}" -F $segment_ou_hashtags ) -Hierarchy "Main:loop configuration"


Global:log -text ("`$global:it_department_ou_manager {0}" -F $global:it_department_ou_manager) -Hierarchy "Main:loop configuration"
Global:log -text ("`$it_ou_hashtags: {0}" -F $it_ou_hashtags ) -Hierarchy "Main:loop configuration"

Global:log -text ("`$global:hr_department_ou_manager {0}" -F $global:hr_department_ou_manager) -Hierarchy "Main:loop configuration"
Global:log -text ("`$hr_ou_hashtags: {0}" -F $hr_ou_hashtags ) -Hierarchy "Main:loop configuration"
#endregion

#region group loop
Global:log -text ("Start") -Hierarchy "Main:Department Loop" 
$raw_data_departments_groupby = $raw_user_data | Group-Object company, department | Select-Object @{name = "company"; expression = { ($_.Group[0].company).replace(" ", "_") } }, @{name = "department"; expression = { $_.Group[0].department } }, group | Sort-Object company
$hr_department_groups_existing = Get-ADGroup -SearchBase ($global:hr_departments_ou) -filter * -Properties managedby, name, info, distinguishedname, extensionattribute1 | Select-Object name, managedby, info, distinguishedname, extensionattribute1
$it_department_groups_existing = Get-ADGroup -SearchBase ($global:it_departments_ou) -filter * -Properties managedby, name, info, distinguishedname, extensionattribute1 | Select-Object name, managedby, info, distinguishedname, extensionattribute1

Global:log -text ("`$raw_data_departments_groupby count :{0}" -f $raw_data_departments_groupby.count ) -Hierarchy "Main:Department Loop" 
Global:log -text ("`$hr_department_groups_existing count :{0}" -f $hr_department_groups_existing.count ) -Hierarchy "Main:Department Loop" 
Global:log -text ("`$it_department_groups_existing count :{0}" -f $it_department_groups_existing.count ) -Hierarchy "Main:Department Loop" 

$raw_data_departments_groupby | Select-Object -Unique -ExpandProperty company | ForEach-Object { #loop through company (segments) from the groupby array
    $this_segment = $_
    Global:log -text ("Start" ) -Hierarchy ("Main:Department Loop:{0}" -F $this_segment) 

    $target_segment_group_name = ( "{1}{0}" -F ($this_segment -replace " ", "_") , $global:Config.Configurations.'company groups'.segments_group_name_prefix )
    Global:log -text ("`$target_segment_group_name : {0}" -f $target_segment_group_name ) -Hierarchy ("Main:Department Loop:{0}:{1}" -F $this_segment, $this_department_row.department) 
    $segment_HR_groups = @()
    $raw_data_departments_groupby | Where-Object { $_.company -eq $this_segment } | ForEach-Object {
        $this_department_row = $_ 
        Global:log -text ("Start" ) -Hierarchy ("Main:Department Loop:{0}:{1}" -F $this_segment, $this_department_row.department) 

        #region HR  group
        Global:log -text ("`$this_department_hr_group : {0}" -f $this_department_hr_group ) -Hierarchy ("Main:Department Loop:{0}:{1}" -F $this_segment, $this_department_row.department) 
        $this_department_hr_group = ( "{1}{0}" -F $this_department_row.department, $global:Config.Configurations.'company groups'.hr_groups_prefix )
        #region update/creation
        #endregion
        #region memberships
        #endregion
        #endregion


        #region it group
        $this_department_it_group = ( "{1}{0}" -F $this_department_row.department, $global:Config.Configurations.'company groups'.it_groups_prefix )
        Global:log -text ("`$this_department_it_group : {0}" -f $this_department_it_group ) -Hierarchy ("Main:Department Loop:{0}:{1}" -F $this_segment, $this_department_row.department) 
        #region update/creation
        #endregion
        #region memberships
        #endregion

        #endregion


        Global:log -text ("End" ) -Hierarchy ("Main:Department Loop:{0}:{1}" -F $this_segment, $this_department_row.department) 
    }

    #region segment group 
    #region update/creation
    #endregion
    #region memberships
    #endregion
    #endregion


    Global:log -text ("End" -f $it_department_groups_existing.count ) -Hierarchy ("Main:Department Loop:{0}" -F $this_company) 
}


if (0) {
    $raw_data_departments_groupby | Select-Object -Unique -ExpandProperty company | ForEach-Object { #loop through company (segments) from the groupby array
        $this_segment = $_
        Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
        $raw_data_departments_groupby | Where-Object { $_.company -eq $this_segment } | Select-Object -ExpandProperty department | ForEach-Object { # all dpt in this segment

            #region hr group
            $this_department_no_prefix = $_
            $this_department_hr_group = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.hr_groups_prefix )

            Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )
            $status_log_row = $status_log_row_template | Select-Object * # group_type, group_name, group_distinguishedname, flag_exists, flag_membership_changed, flag_reporting, action_logs, action_type
            $status_log_row.flag_reporting = $false
            $status_log_row.flag_exists = $false
            $status_log_row.group_type = "department"
            $status_log_row.group_name = $this_department_hr_group
        

            Global:log -text ("? existing...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) -type warning
            $department_groups_names = $department_groups_existing | Select-Object -ExpandProperty name
        
            if ( $department_groups_names -contains $this_department_hr_group) {
                Global:log -text ("...Yes") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                $status_log_row.flag_exists = $true
                $status_log_row.action_type = "Update"
                $status_log_row.group_distinguishedname = $department_groups_existing | Where-Object { $_.name -eq $this_department } | Select-Object -ExpandProperty distinguishedname
            }
            else {
                Global:log -text ("...No") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning
                $status_log_row.action_type = "Create"
            }

            if ( $status_log_row.flag_exists -eq $false) {
                Global:log -text ("Creating...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
            

                $groupParams = @{
                    "Name" = $this_department_hr_group
                    "Path" = $global:hr_departments_ou
                }
                try {
                    # attempt to create the missing group
           
                    # Create a new Active Directory group using the array of parameters
                    Global:log -text (" > Done") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                    New-ADGroup @groupParams -GroupScope $global:group_scope -GroupCategory $global:group_security
                    Global:log -text (" delay...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )  -type warning
                    #Start-Sleep -Seconds 1
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "create group"
                    $log_record.result = 'success'
                    $log_record.target = $global:hr_departments_ou
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
            
                }
                catch {
                    # Catch and handle the error
                    $error_details = $_.Exception.Message
                    $flag_segment_group_update_success = 0
                    Global:log -text (" > group created/updated failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "create group"
                    $log_record.result = 'failed:{0}' -F $error_details 
                    $log_record.target = $global:hr_departments_ou
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true

                }

                Global:log -text ("Updating managedBy attribute") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                try {
                    # Attempt to update the group manager
                    Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:hr_department_ou_manager }
                    Global:log -text (" > Manager updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 

                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update managedBy attributes"
                    $log_record.result = "success"
                    $log_record.target = @{"managedby" = $global:hr_department_ou_manager; } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true

                }
                catch {
                    # Catch and handle the error
                    $error_details = $_.Exception.Message
                    Global:log -text (" > Manager update failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update managedBy attributes"
                    $log_record.result = 'failed:{0}' -F $error_details 
                    $log_record.target = @{"managedby" = $global:hr_department_ou_manager; } | ConvertTo-Json
                    $status_log_row.action_logs += $log_record
                    $status_log_row.flag_reporting = $true
                }

                Global:log -text ("Updating extensionattribute1 (hastags) attribute") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                try {
                    # Attempt to update the group manager
                    Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"extensionattribute1" = ($hr_ou_hashtags -join " ") }
                    Global:log -text (" > extensionattribute1 updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 

                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update extensionattribute1 attribute"
                    $log_record.result = "success"
                    $log_record.target = @{"extensionattribute1" = ($hr_ou_hashtags -join " "); } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true

                }
                catch {
                    # Catch and handle the error
                    $error_details = $_.Exception.Message
                    Global:log -text (" > extensionattribute1 update failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update extensionattribute1 attribute"
                    $log_record.result = 'failed:{0}' -F $error_details 
                    $log_record.target = @{"extensionattribute1" = ($hr_ou_hashtags -join " "); } | ConvertTo-Json
                    $status_log_row.action_logs += $log_record
                    $status_log_row.flag_reporting = $true
                }

            }

            if ($status_log_row.flag_exists -eq $true) {
                Global:log -text ("Checking for managedBy attribute updates...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 

                $current_department = $department_groups_existing | Where-Object { $_.name -eq $status_log_row.group_name } 
            
                if ( $current_department.managedby -ne $global:hr_department_ou_manager ) {
                    Global:log -text ("...required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )  -type warning
                
                    try {
                        # Attempt to update the group manager
                        Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:hr_department_ou_manager; }
                        Global:log -text (" > Manager updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                        $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                        $log_record.timestamp = get-date
                        $log_record.type = "update attributes"
                        $log_record.result = "success"
                        $log_record.target = @{"managedby" = $global:hr_department_ou_manager; } | ConvertTo-Json
                        $status_log_row.action_logs = @($log_record)
                        $status_log_row.flag_reporting = $true
                    }
                    catch {
                        # Catch and handle the error
                        $department_group_update_error = $_.Exception.Message
                        Global:log -text (" > Manager updated failed:{0}" -F $department_group_update_error)-Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) -type error
                        $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                        $log_record.timestamp = get-date
                        $log_record.type = "update attributes"
                        $log_record.result = 'failed:{0}' -F $department_group_update_error 
                        $log_record.target = @{"managedby" = $global:hr_department_ou_manager; } | ConvertTo-Json
                        $status_log_row.action_logs = @($log_record)
                        $status_log_row.flag_reporting = $true
                    }
                }
                else {
                    Global:log -text ("...not required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update attributes"
                    $log_record.result = "skipped"
                    $log_record.target = @{"managedby" = $global:hr_department_ou_manager; } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $false
                }

                Global:log -text ("Checking for extensionattribute1 attribute updates...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                if ( $current_department.extensionattribute1 -ne ($hr_ou_hashtags -join " ") ) {
                    Global:log -text ("...required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )  -type warning
                
                    try {
                        # Attempt to update the group manager
                        Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"extensionattribute1" = ($hr_ou_hashtags -join " "); }
                        Global:log -text (" > extensionattribute1 updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                        $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                        $log_record.timestamp = get-date
                        $log_record.type = "update extensionattribute1 attribute"
                        $log_record.result = "success"
                        $log_record.target = @{"extensionattribute1" = ($hr_ou_hashtags -join " "); } | ConvertTo-Json
                        $status_log_row.action_logs = @($log_record)
                        $status_log_row.flag_reporting = $true
                    }
                    catch {
                        # Catch and handle the error
                        $error_details = $_.Exception.Message
                        Global:log -text (" > Manager updated failed:{0}" -F $department_group_update_error)-Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) -type error
                        $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                        $log_record.timestamp = get-date
                        $log_record.type = "update extensionattribute1 attribute"
                        $log_record.result = 'failed:{0}' -F $error_details 
                        $log_record.target = @{"extensionattribute1" = ($hr_ou_hashtags -join " "); } | ConvertTo-Json
                        $status_log_row.action_logs = @($log_record)
                        $status_log_row.flag_reporting = $true
                    }
                }
                else {
                    Global:log -text ("...not required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group ) 
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update extensionattribute1 attribute"
                    $log_record.result = "skipped"
                    $log_record.target = @{"extensionattribute1" = ($hr_ou_hashtags -join " "); } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $false
                }

            }

            $status_log_row.flag_membership_changed = $false
            #region check memberships
            Global:log -text ("Checking memberships") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )
            $raw_data_departments_groupby | Where-Object { $_.department -eq $this_department_no_prefix } | ForEach-Object {
                $this_group = $_
                $this_ad_group = Get-ADGroup  $status_log_row.group_name -Properties member, distinguishedname
                $current_members = $this_ad_group | Select-Object -ExpandProperty member
                $this_ad_group_distinguishedname = $this_ad_group.distinguishedname
                Global:log -text (" > current : {0}" -f $current_members.count) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )
                Global:log -text (" > required : {0}" -f ($this_group.Group).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )

                $this_group_members_status = @()
                #region missing members
                $this_group.Group | ForEach-Object {
                    $this_group_members_row = "" | Select-Object distinguishedname, status, report
                    $this_group_members_row.distinguishedname = $_.distinguishedname
                    $this_group_members_row.report = 0

                    if ( $current_members -contains $this_group_members_row.distinguishedname ) {
                        $this_group_members_row.status = "present"
                    }
                    else {
                        $this_group_members_row.status = "missing"
                        $this_group_members_row.report = 1
                    }
                    $this_group_members_status += $this_group_members_row
                }
                #endregion

                #region deleted members
                $current_members | ForEach-Object {
                    $this_current_group_member = $_
                    if ( ($this_group.group | Select-Object -ExpandProperty distinguishedname) -notcontains $this_current_group_member) {
                        $this_group_members_row = "" | Select-Object distinguishedname, status, report
                        $this_group_members_row.distinguishedname = $this_current_group_member
                        $this_group_members_row.status = "delete"
                        $this_group_members_row.report = 1
                        $this_group_members_status += $this_group_members_row
                    }
                }


                if ( ($this_group_members_status | Where-Object { $_.status -eq "delete" }).count -ne 0 ) {
                    Global:log -text (" > members to delete : {0}" -f ($this_group_members_status | Where-Object { $_.status -eq "delete" }).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                }
                #endregion

                #region proceed with detected membership changes
                if (1) {
                    if ( ( $this_group_members_status | Where-Object { $_.report -eq 1 } ).count -ne 0 ) {
                        Global:log -text ("proceeding with {0} membership changes" -F ( $this_group_members_status | Where-Object { $_.report -eq 1 } ).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                        $this_group_members_status | Group-Object status | ForEach-Object {
                            $this_stats_row = $_
                            switch ($this_stats_row.name) {
                                "missing" { 
                                    $txt = "   > adding {0} member(s)..." -F $this_stats_row.count 
                                    Global:log -text ($txt ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  
                                    $this_member_action_log_row = $status_log_action_log_row_template | Select-Object *
                                    $this_member_action_log_row.timestamp = Get-Date
                                    $this_member_action_log_row.type = "adding member"
                                    $this_member_action_log_row.target = $this_stats_row.group | Select-Object -ExpandProperty distinguishedname | ConvertTo-Json -Compress

                                    #region adding missing members
                                    try {
                                        # Attempt to add missing members
                                        Add-ADGroupMember -Identity $this_ad_group_distinguishedname -Members ($this_stats_row.group | Select-Object -ExpandProperty distinguishedname )
                                        Global:log -text ("... Success" ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
                                        $this_member_action_log_row.result = "success"
                                    }
                                    catch {
                                        # Catch and handle the error
                                        $add_members_error = $_.Exception.Message
                                        Global:log -text ("... Failed:{0}" -F $add_members_error) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                                        $this_member_action_log_row.result = 'failed:{0}' -F $add_members_error 
                                    }
                                    #endregion
                            
                                }
                                "delete" { 
                                    $txt = "   > removing {0} member(s)..." -F $this_stats_row.count 
                                    Global:log -text ($txt ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  
                                    $this_member_action_log_row = $status_log_action_log_row_template | Select-Object *
                                    $this_member_action_log_row.timestamp = Get-Date
                                    $this_member_action_log_row.type = "removing member"
                                    $this_member_action_log_row.target = $this_stats_row.group | Select-Object -ExpandProperty distinguishedname | ConvertTo-Json -Compress

                                    #region removing members
                                    try {
                                        # Attempt to remove  members
                                        $this_stats_row.group | Select-Object -ExpandProperty distinguishedname | ForEach-Object {
                                            $this_member_to_remove_distinguishedname = $_
                                            Get-ADGroup -Identity $this_ad_group_distinguishedname | Remove-ADGroupMember -Members $this_member_to_remove_distinguishedname -Confirm:$false
                                        }
                                        Global:log -text ("... Success" ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
                                        $this_member_action_log_row.result = "success"
                                    }
                                    catch {
                                        # Catch and handle the error
                                        $add_members_error = $_.Exception.Message
                                        Global:log -text ("... Failed:{0}" -F $add_members_error) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                                        $this_member_action_log_row.result = 'failed:{0}' -F $add_members_error 
                                    }
                                    #endregion
                                }
                            }
                            $status_log_row.flag_membership_changed = $true
                            $status_log_row.flag_reporting = $True
                            $status_log_row.action_logs += $this_member_action_log_row
                        }
                    }
                }
                #endregion


            
            }
        
            #endregion
        
            #endregion

            Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )

            $whole_status += $status_log_row
        
        }
        Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
    }
}



Global:log -text ("End") -Hierarchy "Main:Department Loop" 
#endrgion

EXIT

#region Department Groups
# group by, splitting 'name' instead of the comma separated value of multiple group-object columns
Global:log -text ("Start") -Hierarchy "DepartmentGroups" 
Global:log -text ("Start") -Hierarchy "DepartmentGroups:HR Groups" 
Global:log -text ("End") -Hierarchy "DepartmentGroups:HR Groups" 

Global:log -text ("Start") -Hierarchy "DepartmentGroups:IT Groups" 
$raw_data_departments_groupby = $raw_user_data | Group-Object company, department | Select-Object @{name = "company"; expression = { ($_.Group[0].company).replace(" ", "_") } }, @{name = "department"; expression = { $_.Group[0].department } }, group | Sort-Object company
$department_groups_existing = Get-ADGroup -SearchBase ($global:it_departments_ou) -filter * -Properties managedby, name, info, distinguishedname, extensionattribute1 | Select-Object name, managedby, info, distinguishedname, extensionattribute1
$raw_data_departments_groupby | Select-Object -Unique -ExpandProperty company | ForEach-Object { #loop through company (segments) from the groupby array
    $this_segment = $_
    Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
    $raw_data_departments_groupby | Where-Object { $_.company -eq $this_segment } | Select-Object -ExpandProperty department | ForEach-Object { # all dpt in this segment

        #region hr group
        $this_department_no_prefix = $_
        $this_department_it_group = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.it_groups_prefix )
        $this_department_hr_group_member = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.hr_groups_prefix )

        Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )
        $status_log_row = $status_log_row_template | Select-Object * # group_type, group_name, group_distinguishedname, flag_exists, flag_membership_changed, flag_reporting, action_logs, action_type
        $status_log_row.flag_reporting = $false
        $status_log_row.flag_exists = $false
        $status_log_row.group_type = "department(IT)"
        $status_log_row.group_name = $this_department_it_group
        

        Global:log -text ("? existing...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) -type warning
        $department_groups_names = $department_groups_existing | Select-Object -ExpandProperty name
        
        if ( $department_groups_names -contains $this_department_it_group) {
            Global:log -text ("...Yes") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            $status_log_row.flag_exists = $true
            $status_log_row.action_type = "Update"
            $status_log_row.group_distinguishedname = $department_groups_existing | Where-Object { $_.name -eq $this_department } | Select-Object -ExpandProperty distinguishedname
        }
        else {
            Global:log -text ("...No") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning
            $status_log_row.action_type = "Create"
        }

        if ( $status_log_row.flag_exists -eq $false) {
            Global:log -text ("Creating...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
            

            $groupParams = @{
                "Name" = $this_department_it_group
                "Path" = $global:it_departments_ou
            }
            try {
                # attempt to create the missing group
           
                # Create a new Active Directory group using the array of parameters
                Global:log -text (" > Done") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
                New-ADGroup @groupParams -GroupScope $global:group_scope -GroupCategory $global:group_security
                Global:log -text (" delay...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group )  -type warning
                #Start-Sleep -Seconds 1
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "create group"
                $log_record.result = 'success'
                $log_record.target = $global:it_departments_ou
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true
            
            }
            catch {
                # Catch and handle the error
                $error_details = $_.Exception.Message
                $flag_segment_group_update_success = 0
                Global:log -text (" > group created/updated failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "create group"
                $log_record.result = 'failed:{0}' -F $error_details 
                $log_record.target = $global:it_departments_ou
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true

            }

            Global:log -text ("Updating managedBy attribute") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:it_department_ou_manager }
                Global:log -text (" > Manager updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 

                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update managedBy attributes"
                $log_record.result = "success"
                $log_record.target = @{"managedby" = $global:it_department_ou_manager; } | ConvertTo-Json
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true

            }
            catch {
                # Catch and handle the error
                $error_details = $_.Exception.Message
                Global:log -text (" > Manager update failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update managedBy attributes"
                $log_record.result = 'failed:{0}' -F $error_details 
                $log_record.target = @{"managedby" = $global:it_department_ou_manager; } | ConvertTo-Json
                $status_log_row.action_logs += $log_record
                $status_log_row.flag_reporting = $true
            }

            Global:log -text ("Updating extensionattribute1 (hastags) attribute") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_itgroup ) 
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"extensionattribute1" = ($it_ou_hashtags -join " ") }
                Global:log -text (" > extensionattribute1 updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 

                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update extensionattribute1 attribute"
                $log_record.result = "success"
                $log_record.target = @{"extensionattribute1" = ($it_ou_hashtags -join " "); } | ConvertTo-Json
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true

            }
            catch {
                # Catch and handle the error
                $error_details = $_.Exception.Message
                Global:log -text (" > extensionattribute1 update failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update extensionattribute1 attribute"
                $log_record.result = 'failed:{0}' -F $error_details 
                $log_record.target = @{"extensionattribute1" = ($it_ou_hashtags -join " "); } | ConvertTo-Json
                $status_log_row.action_logs += $log_record
                $status_log_row.flag_reporting = $true
            }

        }

        if ($status_log_row.flag_exists -eq $true) {
            Global:log -text ("Checking for managedBy attribute updates...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 

            $current_department = $department_groups_existing | Where-Object { $_.name -eq $status_log_row.group_name } 
            
            if ( $current_department.managedby -ne $global:it_department_ou_manager ) {
                Global:log -text ("...required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )  -type warning
                
                try {
                    # Attempt to update the group manager
                    Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:it_department_ou_manager; }
                    Global:log -text (" > Manager updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update attributes"
                    $log_record.result = "success"
                    $log_record.target = @{"managedby" = $global:it_department_ou_manager; } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
                }
                catch {
                    # Catch and handle the error
                    $department_group_update_error = $_.Exception.Message
                    Global:log -text (" > Manager updated failed:{0}" -F $department_group_update_error)-Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update attributes"
                    $log_record.result = 'failed:{0}' -F $department_group_update_error 
                    $log_record.target = @{"managedby" = $global:it_department_ou_manager; } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
                }
            }
            else {
                Global:log -text ("...not required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update attributes"
                $log_record.result = "skipped"
                $log_record.target = @{"managedby" = $global:it_department_ou_manager; } | ConvertTo-Json
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $false
            }

            Global:log -text ("Checking for extensionattribute1 attribute updates...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
            if ( $current_department.extensionattribute1 -ne ($it_ou_hashtags -join " ") ) {
                Global:log -text ("...required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group )  -type warning
                
                try {
                    # Attempt to update the group manager
                    Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"extensionattribute1" = ($it_ou_hashtags -join " "); }
                    Global:log -text (" > extensionattribute1 updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update extensionattribute1 attribute"
                    $log_record.result = "success"
                    $log_record.target = @{"extensionattribute1" = ($it_ou_hashtags -join " "); } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
                }
                catch {
                    # Catch and handle the error
                    $error_details = $_.Exception.Message
                    Global:log -text (" > Manager updated failed:{0}" -F $department_group_update_error)-Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update extensionattribute1 attribute"
                    $log_record.result = 'failed:{0}' -F $error_details 
                    $log_record.target = @{"extensionattribute1" = ($it_ou_hashtags -join " "); } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
                }
            }
            else {
                Global:log -text ("...not required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_it_group ) 
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update extensionattribute1 attribute"
                $log_record.result = "skipped"
                $log_record.target = @{"extensionattribute1" = ($it_ou_hashtags -join " "); } | ConvertTo-Json
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $false
            }

        }

        $status_log_row.flag_membership_changed = $false
        #region check memberships
        Global:log -text ("Checking memberships") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )
        $raw_data_departments_groupby | Where-Object { $_.department -eq $this_department_no_prefix } | ForEach-Object {
            $this_group = $_
            $this_ad_group = Get-ADGroup  $status_log_row.group_name -Properties member, distinguishedname
            $current_members = $this_ad_group | Select-Object -ExpandProperty member
            $this_ad_group_distinguishedname = $this_ad_group.distinguishedname
            Global:log -text (" > current : {0}" -f $current_members.count) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )
            Global:log -text (" > required : {0}" -f ($this_group.Group).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department_hr_group )

            $this_group_members_status = @()
            #region missing members
            $this_group.Group | ForEach-Object {
                $this_group_members_row = "" | Select-Object distinguishedname, status, report
                $this_group_members_row.distinguishedname = $_.distinguishedname
                $this_group_members_row.report = 0

                if ( $current_members -contains $this_group_members_row.distinguishedname ) {
                    $this_group_members_row.status = "present"
                }
                else {
                    $this_group_members_row.status = "missing"
                    $this_group_members_row.report = 1
                }
                $this_group_members_status += $this_group_members_row
            }
            #endregion

            #region deleted members
            $current_members | ForEach-Object {
                $this_current_group_member = $_
                if ( ($this_group.group | Select-Object -ExpandProperty distinguishedname) -notcontains $this_current_group_member) {
                    $this_group_members_row = "" | Select-Object distinguishedname, status, report
                    $this_group_members_row.distinguishedname = $this_current_group_member
                    $this_group_members_row.status = "delete"
                    $this_group_members_row.report = 1
                    $this_group_members_status += $this_group_members_row
                }
            }


            if ( ($this_group_members_status | Where-Object { $_.status -eq "delete" }).count -ne 0 ) {
                Global:log -text (" > members to delete : {0}" -f ($this_group_members_status | Where-Object { $_.status -eq "delete" }).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            }
            #endregion

            #region proceed with detected membership changes
            if (1) {
                if ( ( $this_group_members_status | Where-Object { $_.report -eq 1 } ).count -ne 0 ) {
                    Global:log -text ("proceeding with {0} membership changes" -F ( $this_group_members_status | Where-Object { $_.report -eq 1 } ).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                    $this_group_members_status | Group-Object status | ForEach-Object {
                        $this_stats_row = $_
                        switch ($this_stats_row.name) {
                            "missing" { 
                                $txt = "   > adding {0} member(s)..." -F $this_stats_row.count 
                                Global:log -text ($txt ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  
                                $this_member_action_log_row = $status_log_action_log_row_template | Select-Object *
                                $this_member_action_log_row.timestamp = Get-Date
                                $this_member_action_log_row.type = "adding member"
                                $this_member_action_log_row.target = $this_stats_row.group | Select-Object -ExpandProperty distinguishedname | ConvertTo-Json -Compress

                                #region adding missing members
                                try {
                                    # Attempt to add missing members
                                    Add-ADGroupMember -Identity $this_ad_group_distinguishedname -Members ($this_stats_row.group | Select-Object -ExpandProperty distinguishedname )
                                    Global:log -text ("... Success" ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
                                    $this_member_action_log_row.result = "success"
                                }
                                catch {
                                    # Catch and handle the error
                                    $add_members_error = $_.Exception.Message
                                    Global:log -text ("... Failed:{0}" -F $add_members_error) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                                    $this_member_action_log_row.result = 'failed:{0}' -F $add_members_error 
                                }
                                #endregion
                            
                            }
                            "delete" { 
                                $txt = "   > removing {0} member(s)..." -F $this_stats_row.count 
                                Global:log -text ($txt ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  
                                $this_member_action_log_row = $status_log_action_log_row_template | Select-Object *
                                $this_member_action_log_row.timestamp = Get-Date
                                $this_member_action_log_row.type = "removing member"
                                $this_member_action_log_row.target = $this_stats_row.group | Select-Object -ExpandProperty distinguishedname | ConvertTo-Json -Compress

                                #region removing members
                                try {
                                    # Attempt to remove  members
                                    $this_stats_row.group | Select-Object -ExpandProperty distinguishedname | ForEach-Object {
                                        $this_member_to_remove_distinguishedname = $_
                                        Get-ADGroup -Identity $this_ad_group_distinguishedname | Remove-ADGroupMember -Members $this_member_to_remove_distinguishedname -Confirm:$false
                                    }
                                    Global:log -text ("... Success" ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
                                    $this_member_action_log_row.result = "success"
                                }
                                catch {
                                    # Catch and handle the error
                                    $add_members_error = $_.Exception.Message
                                    Global:log -text ("... Failed:{0}" -F $add_members_error) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                                    $this_member_action_log_row.result = 'failed:{0}' -F $add_members_error 
                                }
                                #endregion
                            }
                        }
                        $status_log_row.flag_membership_changed = $true
                        $status_log_row.flag_reporting = $True
                        $status_log_row.action_logs += $this_member_action_log_row
                    }
                }
            }
            #endregion


            
        }
        
        #endregion
        
        #endregion

        Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )

        $whole_status += $status_log_row
        
    }
    Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
}
Global:log -text ("End") -Hierarchy "DepartmentGroups:IT Groups" 

Global:log -text ("End") -Hierarchy "DepartmentGroups" 


$whole_status | Where-Object { $_.flag_reporting -eq $true }

#
#endregion

#region Segment Groups
Global:log -text ("Start") -Hierarchy "SegmentsGroups" 
$segment_groups_existing = Get-ADGroup -SearchBase ($global:segments_ou) -filter * -Properties managedby, name, distinguishedname, extensionattribute1 | Select-Object name, managedby, distinguishedname, extensionattribute1
$segment_groups_required = @()
$raw_data_segments_groupby | Select-Object -ExpandProperty name | ForEach-Object {
    $this_segment_raw_name = $_
    $segment_groups_required += "{1}" -F $global:Config.Configurations.'company groups'.segments_group_name_prefix, (( Remove-StringDiacritic -string $this_segment_raw_name).replace(" ", "-"))
}


$whole_status = @()
$segment_groups_required | ForEach-Object {
    $this_required_segment = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.segments_group_name_prefix )

    $status_log_row = $status_log_row_template | Select-Object * # group_type, group_name, group_distinguishedname, flag_exists, flag_membership_changed, flag_reporting, action_logs, action_type
    $status_log_row.flag_reporting = $false
    $status_log_row.group_type = "segment"
    $status_log_row.group_name = $this_required_segment.replace(" ", "_")

    $segment_groups_names = $segment_groups_existing | Select-Object -ExpandProperty name
    if ( $segment_groups_names -contains $status_log_row.group_name ) {
        $status_log_row.flag_exists = $true
        $status_log_row.action_type = "-"
        Global:log -text ("Update?") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
        $current_manager = $segment_groups_existing | Where-Object { $_.name -eq $status_log_row.group_name } | Select-Object -ExpandProperty managedby
        $status_log_row.group_distinguishedname = $segment_groups_existing | Where-Object { $_.name -eq $status_log_row.group_name } | Select-Object -ExpandProperty distinguishedname
        if ( $current_manager -eq $global:segments_ou_manager ) {
            Global:log -text (" > not required. manager didn't change") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
            $status_log_row.action_type = "-"
            $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
            $log_record.timestamp = get-date
            $log_record.type = "update manager"
            $log_record.result = "skipped"
            $log_record.target = $current_manager
            $status_log_row.action_logs = @($log_record)

        }
        else {
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }

                Global:log -text (" > Manager updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update manager"
                $log_record.result = "success"
                $log_record.target = $current_manager
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true
            }
            catch {
                # Catch and handle the error
                $segment_group_update_error = $_.Exception.Message
                $flag_segment_group_update_success = 0
                Global:log -text (" > Manager updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update manager"
                $log_record.result = 'failed:{0}' -F $segment_group_update_error 
                $log_record.target = $current_manager
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true
            }
        }

        $extensionattribute1 = $segment_ou_hashtags -join " "
        $current_extensionattribute1 = $segment_groups_existing | Where-Object { $_.name -eq $status_log_row.group_name } | Select-Object -ExpandProperty extensionattribute1        
        if ( $current_extensionattribute1 -eq $extensionattribute1 ) {
            Global:log -text (" > not required. hashtags ( extensionattribute1 ) didn't change") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
            $status_log_row.action_type = "-"
            $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
            $log_record.timestamp = get-date
            $log_record.type = "update extensionattribute1"
            $log_record.result = "skipped"
            $log_record.target = $current_extensionattribute1
            $status_log_row.action_logs = @($log_record)

        }
        else {
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"extensionattribute1" = $extensionattribute1 }

                Global:log -text (" > extensionattribute1 updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update extensionattribute1"
                $log_record.result = "success"
                $log_record.target = $extensionattribute1
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true
            }
            catch {
                # Catch and handle the error
                $extensionattribute1_update_error = $_.Exception.Message
                $flag_segment_group_update_success = 0
                Global:log -text (" > extensionattribute1 updated failed:{0}" -F $extensionattribute1_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update extensionattribute1"
                $log_record.result = 'failed:{0}' -F $extensionattribute1_update_error 
                $log_record.target = $extensionattribute1
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true
            }
        }

        
    }
    else {
        $status_log_row.flag_exists = $false
        $status_log_row.action_type = "Create"
        Global:log -text ("Create") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
        #region creating new segment groups
        # Define an array with parameters
        $groupParams = @{
            "Name" = $status_log_row.group_name
            "Path" = $global:segments_ou
        }
        try {
            # attempt to create the missing group
           
            # Create a new Active Directory group using the array of parameters
            Global:log -text (" > Group created successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
            New-ADGroup @groupParams -GroupScope $global:group_scope -GroupCategory $global:group_security
            Global:log -text (" delay...") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name) -type warning

            $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
            $log_record.timestamp = get-date
            $log_record.type = "create group"
            $log_record.result = 'success'
            $log_record.target = $global:segments_ou
            $status_log_row.action_logs = @($log_record)
            $status_log_row.flag_reporting = $true
     
        }
        catch {
            # Catch and handle the error
            $segment_group_update_error = $_.Exception.Message
            $flag_segment_group_update_success = 0
            Global:log -text (" > group created/updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name) -type error
            $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
            $log_record.timestamp = get-date
            $log_record.type = "create group"
            $log_record.result = 'failed:{0}' -F $segment_group_update_error 
            $log_record.target = $global:segments_ou
            $status_log_row.action_logs = @($log_record)
            $status_log_row.flag_reporting = $true

        }

        try {
            Global:log -text (" > updating manager...") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
            $status_log_row.group_distinguishedname = Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) -Properties distinguishedname | Select-Object -ExpandProperty distinguishedname
            Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }
            $flag_segment_group_update_success = 1
            Global:log -text (" > Manager updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)
            $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
            $log_record.timestamp = get-date
            $log_record.type = "update manager"
            $log_record.result = 'success' 
            $log_record.target = $global:segments_ou_manager
            $status_log_row.action_logs += $log_record
            $status_log_row.flag_reporting = $true


        }
        catch {
            # Catch and handle the error
            $segment_group_update_error = $_.Exception.Message
            $flag_segment_group_update_success = 0
            Global:log -text (" > group created/updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name) -type error
            $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
            $log_record.timestamp = get-date
            $log_record.type = "update manager"
            $log_record.result = 'failed:{0}' -F $segment_group_update_error 
            $log_record.target = $global:segments_ou_manager
            $status_log_row.action_logs += $log_record
            $status_log_row.flag_reporting = $true

        }
        #endregion

    }
    $whole_status += $status_log_row
}

Global:log -text ("End") -Hierarchy "SegmentsGroups" 
$whole_status | Where-Object { $_.flag_reporting -eq $true }
#endregion

Global:log -text ("End") -Hierarchy "Main"
#endregion # Main
