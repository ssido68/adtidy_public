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
$global:departments_ou = $global:Config.Configurations.'company groups'.departments_groups_ou
#endregion

#region get target OU attributes, looking for manager information
$global:segments_ou_adobject = Get-ADOrganizationalUnit -Identity $global:segments_ou
$global:segments_ou_manager = $global:segments_ou_adobject.ManagedBy

$global:department_ou_adobject = Get-ADOrganizationalUnit -Identity $global:departments_ou
$global:department_ou_manager = $global:department_ou_adobject.ManagedBy


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

#region Segment Groups
Global:log -text ("Start") -Hierarchy "SegmentsGroups" 
$segment_groups_existing = Get-ADGroup -SearchBase ($global:segments_ou) -filter * -Properties managedby, name, distinguishedname | Select-Object name, managedby, distinguishedname
$segment_groups_required = $raw_data_segments_groupby | Select-Object -ExpandProperty name

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


#region Department Groups
# group by, splitting 'name' instead of the comma separated value of multiple group-object columns
Global:log -text ("Start") -Hierarchy "DepartmentGroups" 
$raw_data_departments_groupby = $raw_user_data | Group-Object company, department | Select-Object @{name = "company"; expression = { ($_.Group[0].company).replace(" ", "_") } }, @{name = "department"; expression = { $_.Group[0].department } }, group | Sort-Object company
$department_groups_existing = Get-ADGroup -SearchBase ($global:departments_ou) -filter * -Properties managedby, name, info, distinguishedname | Select-Object name, managedby, info, distinguishedname
$raw_data_departments_groupby | Select-Object -Unique -ExpandProperty company | ForEach-Object { #loop through company (segments) from the groupby array
    $this_segment = $_
    Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
    $raw_data_departments_groupby | Where-Object { $_.company -eq $this_segment } | Select-Object -ExpandProperty department | ForEach-Object { # all dpt in this segment
        $this_department = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.departments_group_name_prefix )

        Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
        $status_log_row = $status_log_row_template | Select-Object * # group_type, group_name, group_distinguishedname, flag_exists, flag_membership_changed, flag_reporting, action_logs, action_type
        $status_log_row.flag_reporting = $false
        $status_log_row.flag_exists = $false
        $status_log_row.group_type = "department"
        $status_log_row.group_name = $this_department
        

        Global:log -text ("? existing...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning
        $department_groups_names = $department_groups_existing | Select-Object -ExpandProperty name
        
        if ( $department_groups_names -contains $this_department) {
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
            Global:log -text ("Creating...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            

            $groupParams = @{
                "Name" = $this_department
                "Path" = $global:departments_ou
            }
            try {
                # attempt to create the missing group
           
                # Create a new Active Directory group using the array of parameters
                Global:log -text (" > Done") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                New-ADGroup @groupParams -GroupScope $global:group_scope -GroupCategory $global:group_security
                Global:log -text (" delay...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                #Start-Sleep -Seconds 1
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "create group"
                $log_record.result = 'success'
                $log_record.target = $global:departments_ou
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true
            
            }
            catch {
                # Catch and handle the error
                $error_details = $_.Exception.Message
                $flag_segment_group_update_success = 0
                Global:log -text (" > group created/updated failed:{0}" -F $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "create group"
                $log_record.result = 'failed:{0}' -F $error_details 
                $log_record.target = $global:departments_ou
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true

            }

            Global:log -text ("Updating attributes") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group) | Set-ADGroup -Replace @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value }
                Global:log -text (" > Manager updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name)

                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update attributes"
                $log_record.result = "success"
                $log_record.target = @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value } | ConvertTo-Json
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $true

            }
            catch {
                # Catch and handle the error
                $error_details = $_.Exception.Message
                Global:log -text (" > Manager updated failed:{0}" -F $error_details) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group_name) -type error
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update attributes"
                $log_record.result = 'failed:{0}' -F $error_details 
                $log_record.target = @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value } | ConvertTo-Json
                $status_log_row.action_logs += $log_record
                $status_log_row.flag_reporting = $true
            }
        }

        if ($status_log_row.flag_exists -eq $true) {
            Global:log -text ("Checking for attribute updates...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 

            $current_department = $department_groups_existing | Where-Object { $_.name -eq $status_log_row.group_name } 
            
            if ( $current_department.managedby -ne $global:department_ou_manager -or $current_department.info -ne $global:Config.Configurations.'company groups'.departments_note_attribute_value ) {
                Global:log -text ("...required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                
                try {
                    # Attempt to update the group manager
                    Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group_name) | Set-ADGroup -Replace @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value }
                    Global:log -text (" > Manager updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update attributes"
                    $log_record.result = "success"
                    $log_record.target = @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
                }
                catch {
                    # Catch and handle the error
                    $department_group_update_error = $_.Exception.Message
                    Global:log -text (" > Manager updated failed:{0}" -F $department_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type error
                    $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                    $log_record.timestamp = get-date
                    $log_record.type = "update attributes"
                    $log_record.result = 'failed:{0}' -F $department_group_update_error 
                    $log_record.target = @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value } | ConvertTo-Json
                    $status_log_row.action_logs = @($log_record)
                    $status_log_row.flag_reporting = $true
                }
            }
            else {
                Global:log -text ("...not required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                $log_record = $status_log_action_log_row_template | Select-Object * # timestamp, type, target, result
                $log_record.timestamp = get-date
                $log_record.type = "update attributes"
                $log_record.result = "skipped"
                $log_record.target = @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value } | ConvertTo-Json
                $status_log_row.action_logs = @($log_record)
                $status_log_row.flag_reporting = $false
            }
        }

        #region check memberships
        Global:log -text ("Checking memberships") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
        $raw_data_departments_groupby | Select-Object -first 1 | ForEach-Object {
            $this_group = $_
            $this_ad_group = Get-ADGroup  $status_log_row.group_name -Properties member, distinguishedname
            $current_members = $this_ad_group | Select-Object -ExpandProperty member
            $this_ad_group_distinguishedname = $this_ad_group.distinguishedname
            Global:log -text (" > current : {0}" -f $current_members.count) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
            Global:log -text (" > required : {0}" -f ($this_group.Group).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )

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

            if (0) {
                #region proceed with detected membership changes
                if ( ( $this_group_members_status | Where-Object { $_.report -eq 1 } ).count -ne 0 ) {
                
                    Global:log -text ("proceeding with {0} membership changes" -F ( $this_group_members_status | Where-Object { $_.report -eq 1 } ).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                
                    #region membership change loop 
                    $this_group_members_status | Where-Object { $_.report -eq 1 } | ForEach-Object {
                        $this_membership_change_record = $_
                        $this_result = $result_array_row | Select-Object *
                        $this_result.action = "membership updates"
                        $status_log_row.report = 1

                        switch ($this_membership_change_record.status) {
                            "delete" {
                                try {
                                    # Attempt to remove membership
                                
                                    Get-ADGroup -Identity $this_ad_group.DistinguishedName | Remove-ADGroupMember -Members $this_membership_change_record.distinguishedname -Confirm:$false
                                    Global:log -text (" - Removed membership of '{0}'" -F $this_membership_change_record.distinguishedname) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                                    $this_dn_array = "" | Select-Object distinguishedname, result
                                    $this_dn_array.distinguishedname = $this_membership_change_record.distinguishedname
                                    $this_dn_array.result = "success"
                                    $this_result.result += $this_dn_array
                                }
                                catch {
                                    # Catch and handle the error
                                    $error_details = $_.Exception.Message
                                    Global:log -text (" ! Failed to removed membership of '{0}':{1}" -F $this_membership_change_record.distinguishedname, $error_details) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type error
                                    $this_dn_array = "" | Select-Object distinguishedname, result
                                    $this_dn_array.distinguishedname = $this_membership_change_record.distinguishedname
                                    $this_dn_array.result = "failed"
                                    $this_result.result += $error_details
                                }
                            }
                        }
                        $status_log_row.result += @($this_result)


                    }
                    #endregion
                


                }
                else {
                    Global:log -text ("no membership changes required." ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                }
                #endregion

                if ( ( $this_group_members_status | Where-Object { $_.status -eq "missing" }).count -ne 0 ) {
                    Global:log -text (" > missing member(s)..." ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning

                    #region adding missing members
                    try {
                        # Attempt to add missing members
                        Add-ADGroupMember -Identity $this_ad_group_distinguishedname -Members ($this_group_members_status | Where-Object { $_.status -eq 'missing' } | Select-Object -ExpandProperty distinguishedname)
                        Global:log -text ("... added {0} members" -F ($this_group_members_status | Where-Object { $_.status -eq 'missing' } ).count) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
                        $this_result = $result_array_row | Select-Object *
                        $this_result.action = "update manager"
                        $this_result.result = "success"
                        $status_log_row.result = @($this_result)
                        $status_log_row.report = 1
                        $status_log_row.result = @($this_result)
                    }
                    catch {
                        # Catch and handle the error
                        $add_members_error = $_.Exception.Message
                        Global:log -text ("... failed:{0}" -F $add_members_error) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                        $this_result = $result_array_row | Select-Object *
                        $this_result.action = "update manager"
                        $this_result.result = 'failed:{0}' -F $add_members_error 
                        $status_log_row.result = @($this_result)
                        $status_log_row.report = 1
                        $status_log_row.result = @($this_result)
                    }
                
                
                    #endregion

                }
                else {
                    Global:log -text (" > no missing members" ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                }
            }
        }

        #endregion


        Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )

        $whole_status += $status_log_row
        
    }
    Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
}
Global:log -text ("End") -Hierarchy "DepartmentGroups" 


$whole_status | Where-Object { $_.flag_reporting -eq $true }

#
#endregion


Global:log -text ("End") -Hierarchy "Main"
#endregion # Main
