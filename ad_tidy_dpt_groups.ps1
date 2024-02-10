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
#endregion

#region get ad user that have employeeid
$raw_user_data = Get-ADUser -LDAPFilter "(employeeID=*)" -Properties employeeid, department, company, distinguishedname
$raw_data_segments_groupby = $raw_user_data | Group-Object company
#endregion


#region Segment Groups
Global:log -text ("Start") -Hierarchy "SegmentsGroups" 
$segment_groups_existing = Get-ADGroup -SearchBase ($global:segments_ou) -filter * | Select-Object -ExpandProperty name
$segment_groups_required = $raw_data_segments_groupby | Select-Object -ExpandProperty name

$segment_groups_status = @()
$segment_groups_required | ForEach-Object {
    $this_required_segment = $_
    $segment_group_status_row = "" | Select-Object group, exists, action, result 
    $segment_group_status_row.group = $this_required_segment.replace(" ", "_")
    if ( $segment_groups_existing -contains $segment_group_status_row.group ) {
        $segment_group_status_row.exists = $true
        $segment_group_status_row.action = "update"
        Global:log -text ("Update") -Hierarchy ("SegmentsGroups:{0}" -F $segment_group_status_row.group)
        Get-ADGroup -Filter ('name -eq "{0}"' -F $segment_group_status_row.group) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }
    }
    else {
        $segment_group_status_row.exists = $false
        $segment_group_status_row.action = "create"
        Global:log -text ("Create") -Hierarchy ("SegmentsGroups:{0}" -F $segment_group_status_row.group)
        #region creating new segment groups
        # Define an array with parameters
        $groupParams = @{
            "Name" = $segment_group_status_row.group
            "Path" = $global:segments_ou
        }

        # Create a new Active Directory group using the array of parameters
        New-ADGroup @groupParams -GroupScope Universal -GroupCategory Security 
        Get-ADGroup -Filter ('name -eq "{0}"' -F $segment_group_status_row.group) | Set-ADGroup -Replace @{"managedby" = $global:segments_ou_manager }
        #endregion

    }
    $segment_groups_status += $segment_group_status_row
}
Global:log -text ("End") -Hierarchy "SegmentsGroups" 

#endregion





Global:log -text ("End") -Hierarchy "Main"
#endregion # Main
