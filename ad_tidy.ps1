Clear-Host
#region ## script information
$Global:Version = "0.0.1"
# HAR3005, Primeo-Energie, 202310xx
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

#region rules loop
Global:log -text ("Start") -Hierarchy "Main:Rules"
$global:Rules | ForEach-Object {
    $this_rule = $_ 
    Global:log -text ("Init") -Hierarchy ("Main:Rules:{0}" -F $this_rule.Description)

    #region object selection
    Global:log -text ("composing filter...") -Hierarchy ("Main:Rules:{0}" -F $this_rule.Description)

    #region per filter loop
    $filters = $this_rule."Filter definition" | ConvertFrom-Json 
    $filters | Sort-Object sequence | ForEach-Object {
        $global:ldap_query = ""
        $this_filter = $_ 
        #Global:log -text ("init") -Hierarchy ("Main:Rules:{0}:{1}" -F $this_rule.Description, $this_filter.name)
        Switch ($this_filter.Type) {
            "PowerShell" {
                TRY { Remove-Variable -Scope global -Name $this_filter.OutputName -ErrorAction SilentlyContinue } CATCH {}
                New-Variable -Scope global -Name $this_filter.OutputName -Value ( Invoke-Expression -Command $this_filter.Filter.script )
                Global:log -text ("[Powershell] variable '{1}' (Global)= '{0}'" -F (Get-Variable -Scope global -Name $this_filter.OutputName -ValueOnly), $this_filter.OutputName ) -Hierarchy ("Main:Rules:{0}:{1}" -F $this_rule.Description, $this_filter.name)
            }
            "LDAP" {
                if ( $this_filter.Filter.Parameters -ne $null ) {
                    Global:log -text ("[LDAP] parameters definition found, applying...") -Hierarchy ("Main:Rules:{0}:{1}" -F $this_rule.Description, $this_filter.name)
                    $this_filter.Filter.Parameters | Get-Member | Where-Object { $_.membertype -eq "noteproperty" } | ForEach-Object {
                        $this_parameter_name = $_.name
                        Global:log -text ("[LDAP] templating, replacing '<{0}>' with value of global variable named '{1}'" -F $this_parameter_name, $this_filter.Filter.Parameters."$this_parameter_name") -Hierarchy ("Main:Rules:{0}:{1}" -F $this_rule.Description, $this_filter.name)
                        $global:ldap_query = ($this_filter.Filter.query).Replace(("<{0}>" -F $this_parameter_name ), (Get-Variable -Scope global -Name ($this_filter.Filter.Parameters."$this_parameter_name") -ValueOnly))
                        Global:log -text ("[LDAP]  resulting ldap query '{0}'" -F $global:ldap_query) -Hierarchy ("Main:Rules:{0}:{1}" -F $this_rule.Description, $this_filter.name)

                    }
                    
                } else {
                    Global:log -text ("[LDAP] no parameters definition found") -Hierarchy ("Main:Rules:{0}:{1}" -F $this_rule.Description, $this_filter.name) -type warning

                }
            }
        }

    }

    #region object type addition
    switch ($this_rule.'target objects') {
        "user" {
            $ldap_append_string = "(objectCategory=person)(objectClass=user)"
        }
    }

    $global:ldap_query = "(&{0}{1})" -F $ldap_append_string, $global:ldap_query
    #endregion
    #endregion

    #region AD query
    #region attributes list calculation
    $query_attributes = @()
    ($this_rule.'reporting fields').split(",") | ForEach-Object {
        $query_attributes += $_
    }
    #endregion

    #region active directory querying
    Global:log -text ("running ldap:" -F $global:ldap_query ) -Hierarchy ("Main:Rules:{0}:query" -F $this_rule.Description) 
    Global:log -text (" - filter: {0}" -F $global:ldap_query ) -Hierarchy ("Main:Rules:{0}:query" -F $this_rule.Description) 
    Global:log -text (" - attributes: {0}" -F [string]($query_attributes -join ",") ) -Hierarchy ("Main:Rules:{0}:query" -F $this_rule.Description) 
    $filtered_objects = Get-ADObject -LDAPFilter $global:ldap_query -Properties $query_attributes
    Global:log -text (" = {0} object(s) returned " -F $filtered_objects.count ) -Hierarchy ("Main:Rules:{0}:query" -F $this_rule.Description) 
    #endregion
    

    #endregion
    #endregion

    #region action selection
    if ( $this_rule.action -ne $null) {
        Global:log -text ("proceeding with action") -Hierarchy ("Main:Rules:{0}:action" -F $this_rule.Description)
        $rule_actions = $this_rule.action | ConvertFrom-Json
        <# {
            "Name": "disable through AttributeAction",
            "Type": "AttributeAction",
            "Sequence": 3,
            "Action": "disable"
        } #>

        $rule_actions | ForEach-Object {
            $_
            $this_action = $_
            Global:log -text ("[{0}]" -F $this_action.type ) -Hierarchy ("Main:Rules:{0}:action:{1}:{2}" -F $this_rule.Description.$this_action.Sequence, $this_action.Name)
            switch ( $this_action.type) {
                #"AttributeAction" { }
                default { Global:log -text (" ! default hit" -F $this_action.type ) -Hierarchy ("Main:Rules:{0}:action:{1}:{2}" -F $this_rule.Description.$this_action.Sequence, $this_action.Name) -type error }
            }
        }





        Global:log -text ("done") -Hierarchy ("Main:Rules:{0}:action" -F $this_rule.Description)
    } else {
        Global:log -text ("no action defined for rule, skipping") -Hierarchy ("Main:Rules:{0}:action" -F $this_rule.Description) -type warning
    }
    #endregion

    #region reporting
    $standard_attributes = @("objectclass", "objectguid") # for refining output array
    if ( $this_rule.reporting -ne $null) {
        Global:log -text ("reporting is turned on for this rule") -Hierarchy ("Main:Rules:{0}:reporting" -F $this_rule.Description)
        $standard_attributes = @("objectclass", "objectguid") # for refining output array
        $empty_line = "" | Select-Object ignore_this_attribute

        # init pre object loop
        $raw_filtered_objects = $filtered_objects | Select-Object *
        $currated_filtered_objects = @()

        #region calculated which properties are supposed to be included in reporting
        $raw_properties = $raw_filtered_objects | Get-Member | Where-Object { $_.membertype -eq "noteproperty" } | Select-Object name | Sort-Object -Property Name -Unique | Select-Object -ExpandProperty name
        $filtered_properties = @()
        $raw_properties | ForEach-Object {
            if ($standard_attributes -contains $_ -or ( ($this_rule.'reporting fields').split(",") -contains $_ ) ) {
                $filtered_properties += $_
            }
        }
        $filtered_properties | ForEach-Object { # populate single row with column for all required attribute for output based of config file
            $empty_line = $empty_line | Select-Object *, $_
        }
        #endregion

        #region currated array preparation
        Global:log -text ("currating output array...") -Hierarchy ("Main:Rules:{0}:reporting" -F $this_rule.Description)
        $raw_filtered_objects <# | Select-Object -First 2 #> | ForEach-Object {
            #init single loop variables
            $this_object_line = $_
            $this_object_new_line = $empty_line | Select-Object * -ExcludeProperty ignore_this_attribute
            #Global:log -text ("this_object_new_line:{0}" -F $this_object_new_line | ConvertTo-Json -Compress) -Hierarchy ("Main:Rules:{0}:DEBUG" -F $ThisRule.Description)
            $filtered_properties | ForEach-Object {
                $this_property = $_
                switch ($this_property) {
                    # recipes for special attribute calculation
                    "useraccountcontrol" { $this_object_new_line."$this_property" = Global:DecodeUserAccountControl -UAC $this_object_line."$this_property" }
                    "lastlogontimestamp" { $this_object_new_line."$this_property" = [datetime]::FromFileTime($this_object_line."$this_property") }
                    "pwdLastSet" { $this_object_new_line."$this_property" = [datetime]::FromFileTime($this_object_line."$this_property") }
                    default { $this_object_new_line."$this_property" = $this_object_line."$this_property" } # default processing, direct value assignment
                }
            }
            $currated_filtered_objects += $this_object_new_line | Select-Object * -ExcludeProperty ignore_this_attribute
        }
        #endregion

    } else {
        Global:log -text ("reporting is turned off for this rule") -Hierarchy ("Main:Rules:{0}:reporting" -F $this_rule.Description) -type warning
    }
    #$currated_filtered_objects | Select-Object -First 10 | Format-Table -AutoSize
    Global:log -text ("Done.") -Hierarchy ("Main:Rules:{0}" -F $ThisRule.Description)
    #endregion
}
Global:log -text ("End") -Hierarchy "Main:Rules"
#endregion
Global:log -text ("End") -Hierarchy "Main"
#endregion # Main
