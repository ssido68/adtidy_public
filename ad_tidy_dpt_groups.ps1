﻿Clear-Host
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
#endregion

#region get ad user that have employeeid
$raw_user_data = Get-ADUser -LDAPFilter "(employeeID=*)" -Properties employeeid, department, company, distinguishedname
$raw_data_segments_groupby = $raw_user_data | Group-Object company
#endregion


#region Segment Groups
Global:log -text ("Start") -Hierarchy "SegmentsGroups" 
$segment_groups_existing = Get-ADGroup -SearchBase ($global:segments_ou) -filter * | Select-Object -ExpandProperty name
$segment_groups_required = $raw_data_segments_groupby | Select-Object -ExpandProperty name

$status_status = @()
$segment_groups_required | ForEach-Object {
    $this_required_segment = $_
    $status_log_row = "" | Select-Object type, group, exists, action, result, report
    $status_log_row.report = 0
    $status_log_row.type = "segment group"
    $result_array_row = "" | Select-Object action, result
    $status_log_row.group = $this_required_segment.replace(" ", "_")
    if ( $segment_groups_existing -contains $status_log_row.group ) {
        $status_log_row.exists = $true
        $status_log_row.action = "update"
        Global:log -text ("Update") -Hierarchy ("SegmentsGroups:{0}" -F $status_log_row.group)
        try {
            # Attempt to update the group manager
            Get-ADGroup -Filter ('name -eq "{0}"' -F $status_log_row.group) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }
            $flag_segment_group_update_success = 1
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
            New-ADGroup @groupParams -GroupScope Universal -GroupCategory Security 
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

        Start-Sleep -Seconds 2

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
    $status_status += $status_log_row
}
Global:log -text ("End") -Hierarchy "SegmentsGroups" 

#endregion

$status_status | Where-Object { $_.report -eq 1 }



Global:log -text ("End") -Hierarchy "Main"
#endregion # Main
