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
#region status row array template
$status_log_row_template = "" | Select-Object type, group, exists, action, result, report

#endregion

#region Segment Groups
Global:log -text ("Start") -Hierarchy "SegmentsGroups" 
$segment_groups_existing = Get-ADGroup -SearchBase ($global:segments_ou) -filter * -Properties managedby, name | Select-Object name, managedby
$segment_groups_required = $raw_data_segments_groupby | Select-Object -ExpandProperty name

$whole_status = @()
$segment_groups_required | ForEach-Object {
    $this_required_segment = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.segments_group_name_prefix )

    $status_log_row = $status_log_row_template | Select-Object *
    $status_log_row.report = 0
    $status_log_row.type = "segment group"
    $result_array_row = "" | Select-Object action, result
    $status_log_row.group = $this_required_segment.replace(" ", "_")
    $segment_groups_names = $segment_groups_existing | Select-Object -ExpandProperty name
    if ( $segment_groups_names -contains $status_log_row.group ) {
        $status_log_row.exists = $true
        $status_log_row.action = "update"
        Global:log -text ("Update?") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
        $current_manager = $segment_groups_existing | Where-Object { $_.name -eq $status_log_row.group } | Select-Object -ExpandProperty managedby
        if ( $current_manager -eq $global:segments_ou_manager ) {
            Global:log -text (" > not required. manager didn't change") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
            $this_result = $result_array_row | Select-Object *
            $this_result.action = "update manager"
            $this_result.result = "skipped"
            $status_log_row.result = @($this_result)

        }
        else {
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }

                Global:log -text (" > Manager updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "update manager"
                $this_result.result = "success"
                $status_log_row.result = @($this_result)
            }
            catch {
                # Catch and handle the error
                $segment_group_update_error = $_.Exception.Message
                $flag_segment_group_update_success = 0
                Global:log -text (" > Manager updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type error
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "update manager"
                $this_result.result = 'failed:{0}' -F $segment_group_update_error 
                $status_log_row.result = @($this_result)
                $status_log_row.report = 1

            }
        }


        
    }
    else {
        $status_log_row.exists = $false
        $status_log_row.action = "create"
        Global:log -text ("Create") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
        #region creating new segment groups
        # Define an array with parameters
        $groupParams = @{
            "Name" = $status_log_row.group
            "Path" = $global:segments_ou
        }
        try {
            # attempt to create the missing group
           
            # Create a new Active Directory group using the array of parameters
            Global:log -text (" > Group created successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
            New-ADGroup @groupParams -GroupScope $global:group_scope -GroupCategory $global:group_security
            Global:log -text (" delay...") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type warning

            $this_result = $result_array_row | Select-Object *
            $this_result.action = "create group"
            $this_result.result = 'success' 
            $status_log_row.result = @($this_result)
            $status_log_row.report = 1
            
        }
        catch {
            # Catch and handle the error
            $segment_group_update_error = $_.Exception.Message
            $flag_segment_group_update_success = 0
            Global:log -text (" > group created/updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type error
            $this_result = $result_array_row | Select-Object *
            $this_result.action = "create group"
            $this_result.result = 'failed:{0}' -F $segment_group_update_error 
            $status_log_row.result = @($this_result)
            $status_log_row.report = 1
        }

        #Start-Sleep -Seconds 2

        try {
            Global:log -text (" > updating manager...") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
            Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }
            $flag_segment_group_update_success = 1
            Global:log -text (" > Manager updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
            $this_result = $result_array_row | Select-Object *
            $this_result.action = "update manager"
            $this_result.result = 'success' 
            $status_log_row.result += $this_result
            $status_log_row.report = 1

        }
        catch {
            # Catch and handle the error
            $segment_group_update_error = $_.Exception.Message
            $flag_segment_group_update_success = 0
            Global:log -text (" > group created/updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type error
            $this_result = $result_array_row | Select-Object *
            $this_result.action = "update manager"
            $this_result.result = 'failed:{0}' -F $segment_group_update_error 
            $status_log_row.result += $this_result
            $status_log_row.report = 1
        }


        #endregion

    }
    $whole_status += $status_log_row
}
Global:log -text ("End") -Hierarchy "SegmentsGroups" 

#endregion


#region Department Groups
# group by, splitting 'name' instead of the comma separated value of multiple group-object columns
Global:log -text ("Start") -Hierarchy "DepartmentGroups" 
$raw_data_departments_groupby = $raw_user_data | Group-Object company, department | Select-Object @{name = "company"; expression = { ($_.Group[0].company).replace(" ", "_") } }, @{name = "department"; expression = { $_.Group[0].department } }, group | Sort-Object company
$department_groups_existing = Get-ADGroup -SearchBase ($global:departments_ou) -filter * -Properties managedby, name, info | Select-Object name, managedby, info
$raw_data_departments_groupby | Select-Object -Unique -ExpandProperty company | ForEach-Object { #loop through company (segments) from the groupby array
    $this_segment = $_
    Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
    $raw_data_departments_groupby | Where-Object { $_.company -eq $this_segment } | Select-Object -ExpandProperty department | ForEach-Object { # all dpt in this segment
        $this_department = ( "{1}{0}" -F $_, $global:Config.Configurations.'company groups'.departments_group_name_prefix )

        Global:log -text ("Start") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
        $status_log_row = $status_log_row_template | Select-Object * # type, group, exists, action, result, report
        $status_log_row.type = "department"
        $status_log_row.group = $this_department
        $status_log_row.exists = $false
        Global:log -text ("? existing...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning
        $department_groups_names = $department_groups_existing | Select-Object -ExpandProperty name
        
        if ( $department_groups_names -contains $this_department) {
            Global:log -text ("...Yes") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            $status_log_row.exists = $true
            $status_log_row.action = "Update"
        }
        else {
            Global:log -text ("...No") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning
            $status_log_row.action = "Create"
        }

        if ( $status_log_row.exists -eq $false) {
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
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "create group"
                $this_result.result = 'success' 
                $status_log_row.result = @($this_result)
                $status_log_row.report = 1
            
            }
            catch {
                # Catch and handle the error
                $segment_group_update_error = $_.Exception.Message
                $flag_segment_group_update_success = 0
                Global:log -text (" > group created/updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "create group"
                $this_result.result = 'failed:{0}' -F $segment_group_update_error 
                $status_log_row.result = @($this_result)
                $status_log_row.report = 1
            }

            Global:log -text ("Updating attributes") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            try {
                # Attempt to update the group manager
                Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group) | Set-ADGroup -Replace @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value }
                Global:log -text (" > Manager updated successfully") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "update attributes"
                $this_result.result = "success"
                $status_log_row.result += @($this_result)
            }
            catch {
                # Catch and handle the error
                $department_group_update_error = $_.Exception.Message
                Global:log -text (" > Manager updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type error
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "update attributes"
                $this_result.result = 'failed:{0}' -F $department_group_update_error 
                $status_log_row.result += @($this_result)
                $status_log_row.report = 1

            }
        }

        if ($status_log_row.exists -eq $true) {
            Global:log -text ("Checking for attribute updates...") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 

            $current_department = $department_groups_existing | Where-Object { $_.name -eq $status_log_row.group } 
            
            if ( $current_department.managedby -ne $global:department_ou_manager -or $current_department.info -ne $global:Config.Configurations.'company groups'.departments_note_attribute_value ) {
                Global:log -text ("...required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )  -type warning
                $this_result.action = "update attributes"
                
                try {
                    # Attempt to update the group manager
                    Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group) | Set-ADGroup -Replace @{"managedby" = $global:department_ou_manager; "info" = $global:Config.Configurations.'company groups'.departments_note_attribute_value }
                    Global:log -text (" > Manager updated successfully") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type error
                    $this_result = $result_array_row | Select-Object *
                    $this_result.action = "update manager"
                    $this_result.result = "success"
                    $status_log_row.result = @($this_result)
                    $status_log_row.report = 1
                    $status_log_row.result = @($this_result)
                }
                catch {
                    # Catch and handle the error
                    $department_group_update_error = $_.Exception.Message
                    Global:log -text (" > Manager updated failed:{0}" -F $segment_group_update_error) -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group) -type error
                    $this_result = $result_array_row | Select-Object *
                    $this_result.action = "update manager"
                    $this_result.result = 'failed:{0}' -F $department_group_update_error 
                    $status_log_row.result = @($this_result)
                    $status_log_row.report = 1
                    $status_log_row.result = @($this_result)
                }
                

            }
            else {
                Global:log -text ("...not required") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
                $this_result = $result_array_row | Select-Object *
                $this_result.action = "update attributes"
                $this_result.result = "skipped"
                $status_log_row.report = 0
                $status_log_row.result = @($this_result)

            }
            

            
            



        }

        #region check memberships
        Global:log -text ("Checking memberships") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
        $raw_data_departments_groupby | Select-Object -first 1 | ForEach-Object {
            $this_group = $_
            $current_members = Get-ADGroup  $status_log_row.group -Properties member | Select-Object -ExpandProperty member
            Global:log -text (" > current : {0}" -f $current_members.count) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )
            Global:log -text (" > required : {0}" -f ($this_group.Group).count ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )

            $this_group_members_status = @()
            $this_group.Group | ForEach-Object {
                $this_group_members_row = "" | Select-Object distinguishedname, status
                $this_group_members_row.distinguishedname = $_.distinguishedname
                if ( $current_members -contains $this_group_members_row.distinguishedname ) {
                    $this_group_members_row.status = "present"
                }
                else {
                    $this_group_members_row.status = "missing"
                }
                $this_group_members_status += $this_group_members_row
            }

            if ( ( $this_group_members_status | Where-Object { $_.status -eq "missing" }).count -ne 0 ) {
                Global:log -text (" > missing member(s)..." ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) -type warning

                #region adding missing members
                
                #endregion

            }
            else {
                Global:log -text (" > no missing members" ) -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department ) 
            }
        }

        #endregion


        Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}>{1}" -F $this_segment, $this_department )

        $whole_status += $status_log_row
    }
    Global:log -text ("End") -Hierarchy ("DepartmentGroups:{0}" -F $this_segment)
}
Global:log -text ("End") -Hierarchy "DepartmentGroups" 


$whole_status | Where-Object { $_.report -eq 1 }

#
#endregion


Global:log -text ("End") -Hierarchy "Main"
#endregion # Main
