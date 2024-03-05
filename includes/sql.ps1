#region : SQL Utility DB connection init.

function Global:sql_query {
    param(
        $SqlInstance = $global:Config.Configurations.inventory."sql hostname",
        $Database = $global:Config.Configurations.inventory.'sql database',
        [Parameter(Mandatory = $true)] $query
    )

    $Global:adtools_sql_query_last_error = "Ok"

    Global:log -Hierarchy "function:sql_query" -text ("SqlInstance={0},Database={1},TrustServerCertificate=false,Query={2}" -F $SqlInstance, $Database, $query )
    
    # https://github.com/dataplat/dbatools/discussions/7680
    Set-DbatoolsConfig -FullName 'sql.connection.trustcert' -Value $true -Register
    if ($query -match 'INSERT') {
        $query = "{0};SELECT SCOPE_IDENTITY() as id;" -F $query
    }
    try {
        $temp = Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $query -MessagesToOutput -EnableException 
    } 
    
    catch {
        $Global:adtools_sql_query_last_error = $_
        Global:log -Hierarchy "function:sql_query" -text ("error details:{0}" -F $Global:adtools_sql_query_last_error) -type error
        exit
    }
    Global:log -Hierarchy "function:sql_query" -text ("returned rows={0}" -F ($temp | Measure-Object).count )
    if ($query -match 'INSERT') {
        Global:log -Hierarchy "function:sql_query" -text ("returned id='{0}'" -F $temp.id )
        return $temp.id
    }
    else {
        return $temp
    }
    
    #return $temp.id



}


function Global:SQL_ADImport_All_Employees {
    param(
    )

    $Query = "SELECT * FROM {0} WHERE ad_employeeid is not null " -F "[view_AD_Employees]"
    
    return Global:sql_query -query $Query
}

function Global:SQL_ADimport_Summary_reports {
    param(
    )

    $Query = "SELECT COUNT (record_id) as amount,record_status,entry_type,user_type,execution_status_code,[execution_mode]
        FROM (SELECT [record_id],[record_status],[entry_type],[user_type],[execution_status_code],[execution_mode] FROM [ITTool].[dbo].[ITTOOL_adimport_records] WHERE datediff(d,entry_first_occurrence ,getdate()) <1 OR datediff(d,entry_last_occurrence ,getdate()) <1 ) FILTERED
        GROUP BY record_status,entry_type,user_type,execution_status_code,[execution_mode]"
    
    return Global:sql_query -query $Query

}



function Global:ADTidy_Inventory_Users_sql_table_check {
    param(
        $Table_Name = "ADTidy_Inventory_Users"
    )

    $config = @"
{"Fields": [{"name": "record_source","type": "VARCHAR(100)"},{"name": "record_lastupdate","type": "DATETIME"},{"name": "record_status","type": "VARCHAR(50)"},{"name": "ad_whenCreated","type": "DATETIME"},{"name": "ad_whenChanged","type": "DATETIME"},{"name": "ad_distinguishedname","type": "VARCHAR(MAX)"},{"name": "ad_lastlogontimestamp","type": "DATETIME","nullable": 1},{"name": "ad_pwdLastSet","type": "DATETIME","nullable": 1},{"name": "ad_extensionAttribute2","type": "VARCHAR(50)","nullable": 1},{"name": "ad_samaccountname","type": "VARCHAR(33)"},{"name": "ad_userprincipalname","type": "VARCHAR(100)","nullable": 1},{"name": "ad_objectguid","type": "VARCHAR(50)"},{"name": "ad_sid","type": "VARCHAR(50)"},{"name": "ad_userAccountControl","type": "VARCHAR(100)"},{"name": "ad_accountExpires","type": "DATETIME","nullable": 1},{"name": "ad_extensionAttribute4","type": "VARCHAR(20)","nullable": 1},{"name": "ad_extensionAttribute7","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_givenName","type": "VARCHAR(50)","nullable": 1},{"name": "ad_sn","type": "VARCHAR(100)","nullable": 1},{"name": "ad_initials","type": "VARCHAR(10)","nullable": 1},{"name": "ad_displayname","type": "VARCHAR(50)","nullable": 1},{"name": "ad_division","type": "VARCHAR(20)","nullable": 1},{"name": "ad_description","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_info","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_company","type": "VARCHAR(100)","nullable": 1},{"name": "ad_department","type": "VARCHAR(16)","nullable": 1},{"name": "ad_extensionAttribute5","type": "VARCHAR(20)","nullable": 1},{"name": "ad_departmentnumber","type": "VARCHAR(20)","nullable": 1},{"name": "ad_title","type": "VARCHAR(50)","nullable": 1},{"name": "ad_employeeid","type": "VARCHAR(30)","nullable": 1},{"name": "ad_employeetype","type": "VARCHAR(30)","nullable": 1},{"name": "ad_extensionAttribute1","type": "VARCHAR(70)","nullable": 1},{"name": "ad_manager","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_thumbnailPhoto","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_physicaldeliveryofficename","type": "VARCHAR(100)","nullable": 1},{"name": "ad_streetaddress","type": "VARCHAR(50)","nullable": 1},{"name": "ad_postalcode","type": "VARCHAR(20)","nullable": 1},{"name": "ad_l","type": "VARCHAR(50)","nullable": 1},{"name": "ad_c","type": "VARCHAR(2)","nullable": 1},{"name": "ad_extensionAttribute3","type": "VARCHAR(2)","nullable": 1},{"name": "ad_preferredLanguage","type": "VARCHAR(30)","nullable": 1},{"name": "ad_telephonenumber","type": "VARCHAR(50)","nullable": 1},{"name": "ad_mobile","type": "VARCHAR(50)","nullable": 1},{"name": "ad_MsExchUserCulture","type": "VARCHAR(30)","nullable": 1},{"name": "ad_mail","type": "VARCHAR(100)","nullable": 1},{"name": "ad_homeMdb","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_msExchMailboxGuid","type": "VARCHAR(50)","nullable": 1},{"name": "ad_proxyaddresses","type": "XML","nullable": 1},{"name": "ad_extensionAttribute6","type": "VARCHAR(MAX)","nullable": 1},{"name": "az_MFA","type": "XML","nullable": 1},{"name": "xml_extended_attributes","type": "XML","nullable": 1}],"FieldsAssignement": [{"name": "record_source","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "record_source"}]}]},{"name": "record_lastupdate","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "current_datetime"}]}]},{"name": "record_status","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "status"}]}]},{"name": "ad_whenCreated","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenCreated"}]}]},{"name": "ad_whenChanged","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenChanged"}]}]},{"name": "ad_distinguishedname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "distinguishedname"}]}]},{"name": "ad_lastlogontimestamp","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "lastlogontimestamp"}]}]},{"name": "ad_pwdLastSet","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "pwdLastSet"}]}]},{"name": "ad_extensionAttribute2","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionattribute2"}]}]},{"name": "ad_samaccountname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "samaccountname"}]}]},{"name": "ad_userprincipalname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "userprincipalname"}]}]},{"name": "ad_objectguid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "objectguid"}]}]},{"name": "ad_sid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "SID"}]}]},{"name": "ad_userAccountControl","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "useraccountcontrol"}]}]},{"name": "ad_accountExpires","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "accountexpires"}]}]},{"name": "ad_extensionAttribute4","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionAttribute4"}]}]},{"name": "ad_extensionAttribute7","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionAttribute7"}]}]},{"name": "ad_givenName","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "givenName"}]}]},{"name": "ad_sn","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "sn"}]}]},{"name": "ad_initials","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "initials"}]}]},{"name": "ad_displayname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "displayname"}]}]},{"name": "ad_division","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "division"}]}]},{"name": "ad_description","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "description"}]}]},{"name": "ad_info","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "info"}]}]},{"name": "ad_company","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "company"}]}]},{"name": "ad_department","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "department"}]}]},{"name": "ad_extensionAttribute5","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionattribute5"}]}]},{"name": "ad_departmentnumber","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "departmentnumber"}]}]},{"name": "ad_title","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "title"}]}]},{"name": "ad_employeeid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "employeeid"}]}]},{"name": "ad_employeetype","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "employeetype"}]}]},{"name": "ad_extensionAttribute1","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionattribute1"}]}]},{"name": "ad_manager","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "manager"}]}]},{"name": "ad_thumbnailPhoto","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "thumbnailPhoto"}]}]},{"name": "ad_physicaldeliveryofficename","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "physicaldeliveryofficename"}]}]},{"name": "ad_streetaddress","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "streeatddress"}]}]},{"name": "ad_postalcode","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "postalcode"}]}]},{"name": "ad_l","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "l"}]}]},{"name": "ad_c","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "c"}]}]},{"name": "ad_extensionAttribute3","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionAttribute3"}]}]},{"name": "ad_preferredLanguage","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "preferredLanguage"}]}]},{"name": "ad_telephonenumber","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "telephonenumber"}]}]},{"name": "ad_mobile","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "mobile"}]}]},{"name": "ad_MsExchUserCulture","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "MsExchUserCulture"}]}]},{"name": "ad_mail","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "mail"}]}]},{"name": "ad_homeMdb","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "homeMdb"}]}]},{"name": "ad_msExchMailboxGuid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "msExchMailboxGuid"}]}]},{"name": "ad_proxyaddresses","Recipe":[{"ORDER":"1", "Content": [ {"Source": "","Type": ""}]}]},{"name": "ad_extensionAttribute6","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "extensionattribute6"}]}]},{"name": "az_MFA","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "mfa"}]}]},{"name": "xml_extended_attributes","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "xml_extended_attributes"}]}]}]}										
"@ | ConvertFrom-Json


    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Running check for table:{0}" -F $Table_Name) -type warning
    $Query_Check_Table_Exists = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{0}'" -F $Table_Name
    $Result_Check_Table_Exists = Global:sql_query -query $Query_Check_Table_Exists

    if ( ($Result_Check_Table_Exists | Measure-Object).count -eq 0 ) {
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' does not exist" -F $Table_Name) -type warning
        $Query_Create_Table_main = "CREATE TABLE {0} ({1}{2});"
        $Query_Create_Table_Constraint = "CONSTRAINT {0} UNIQUE ({1})" 
        
        $Query_String_Fields = ""
        $Query_String_Constraints = ""
        $Has_Contraints = 0
        $config.Fields | ForEach-Object {
            if ($_.nullable -ne 1) { $string_nullable = " NOT NULL" } else { $string_nullable = " NULL" }
            if ($_.id -eq 1) { $string_id = " IDENTITY(1,1) NOT NULL"; $string_nullable = $null } else { $string_id = $null }
            $Query_String_Fields += "`n{0} {1}{2}{3}," -F $_.name, $_.type, $string_id, $string_nullable
            if ($_.constraints -eq 1) { $Query_String_Constraints += "{0}," -F $_.Name; $Has_Contraints = 1 }
            
    
        }
        $Query_String_Fields = $Query_String_Fields.Substring(1)# removes first character  from string
        if ( $Has_Contraints -eq 0 ) {
            #write-host "no constraints"
            $Query_String_Fields = $Query_String_Fields -replace ".$" # removes last character from string
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, $Query_String_Fields, $null
        }
        else {
            #write-host "with constraints"
            $Query_String_Constraints = $Query_String_Constraints -replace ".$" # removes last character from string
            $Query_String_Constraints = $Query_Create_Table_Constraint -F ("CONSTRAINT_{0}" -F $Table_Name), $Query_String_Constraints
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, ("`n" + $Query_String_Fields), ("`n" + $Query_String_Constraints)
        }
        IF ( $Global:WhatIf -ne $true ) {
            Global:sql_query -query $Query_Final_Create_Query 
        }
        ELSE {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Query_Final_Create_Query ) -type warning
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
        }


    }
    else {
        $Table_schema_name = "[{0}].[{1}].[{2}]" -F $Result_Check_Table_Exists.TABLE_CATALOG, $Result_Check_Table_Exists.TABLE_SCHEMA, $Result_Check_Table_Exists.TABLE_NAME
        
        $Query_Select_all = "SELECT * FROM {0}" -F $Table_Name
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' exists, {1} rows in it" -F $Table_Name, (Global:sql_query -query $Query_Select_all).count) 
        
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text "Verifiying columns..."
        $Query_List_columns = "SELECT col.name AS column_name,t.name AS data_type,col.max_length AS data_type_detail FROM sys.tables AS tab INNER JOIN sys.columns AS col ON tab.object_id = col.object_id LEFT JOIN sys.types AS t ON col.user_type_id = t.user_type_id WHERE tab.name = '{0}'" -F $Table_Name
        $SqlCurrentColumns = Global:sql_query -query $Query_List_columns

        $Queries_Update_table = ""
        $config.Fields | ForEach-Object {
            $this = $_ | Select-Object *, type_sql, type_detail

            #write-host ( "split of {0}, count:{1}" -F $this.type, ($this.type).split("(").count )
            if ( ($this.type).split("(").count -eq 1) {
                # no detail type such as DATETIME or INT
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = $null
            }
            else {
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = ($this.type).split("(")[1] -replace ".$"
            }
            #write-host ( "json this:{0}" -F ( $this | ConvertTo-Json -Compress ))
            
            $MatchingColumn = "" | Select-Object found, same_type, type_details_match
            $MatchingColumn.found = 0
            $MatchingColumn.same_type = 0
            $MatchingColumn.type_details_match = 0
            $SqlCurrentColumns | Where-Object { $_.column_name -eq $this.name } | ForEach-Object {
                $MatchingColumn.found = 1
                if ( $_.data_type -eq $this.type_sql ) { 
                    $MatchingColumn.same_type = 1
                    if ($_.data_type_detail -eq -1) { $temp_data_type_detail = [string]"MAX" } ELSE { $temp_data_type_detail = [string]$_.data_type_detail }
                    if ( $this.type_detail -ne $null) {
                        if ( $temp_data_type_detail -eq $this.type_detail) {
                            $MatchingColumn.type_details_match = 1
                        }
                        else {
                            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type detail for column ({3}) {2} :({0} <> {1})." -F $this.type_detail, $temp_data_type_detail, $this.name, $this.type) -type warning
                        }
                    }
                    else {
                        $MatchingColumn.type_details_match = 1
                    }

                }
                else {
                    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type for column {2} ({0} <> {1})." -F $this.type_sql, $_.data_type, $this.name) -type warning
                }
                #write-host ( "json MatchingColumn:{0}" -F ( $MatchingColumn | ConvertTo-Json -Compress ))

            }
        

            if ( $MatchingColumn.found -eq 0 ) {  
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "column {0} missing" -F $this.name) -type warning

            }

            $Query_Run = $null
            $Query_Alter_Table_Add_Column = "ALTER TABLE {0} ADD {1} {2};"
            $Query_Alter_Table_Alter_Column = "ALTER TABLE {0} ALTER COLUMN {1} {2};"

            if ( $MatchingColumn.found -eq 0 ) { 
                $Query_Run = $Query_Alter_Table_Add_Column 
            }
            else {
                if ( $MatchingColumn.same_type -eq 0 -or $MatchingColumn.type_details_match -eq 0 ) {
                    $Query_Run = $Query_Alter_Table_Alter_Column 
                }
            }

            if ( $Query_Run -ne $null ) {
                $Query_Run = $Query_Run -F $Table_schema_name, $this.name, $this.type
                $Queries_Update_table += $Query_Run + "`n"
            }

        }

        if ( $Queries_Update_table.Length -ge 1 ) {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "resulting alter table queries:{0}" -F $Queries_Update_table)
            IF ( $Global:WhatIf -ne $true ) {
                Global:sql_query -query $Queries_Update_table 
            }
            ELSE {
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Queries_Update_table ) -type warning
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
            }

        }
        else {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Table {0} is matching the definition of the config.json at run time." -F $Table_Name)
        }

        
    }
}
function Global:ADTidy_Inventory_Users_sql_update {
    param(
        [Parameter(Mandatory = $true)] [array]$Fields,
        $Table_Name = "ADTidy_Inventory_Users"
    )

    #region ADTidy_Inventory_Users_sql_update specific
    $prefixed_fields = "" | Select-Object ignore
    $varchar_field = @()
    $ObjectGUID = $Fields.ObjectGUID
    $Fields | Select-Object -ExcludeProperty ObjectGUID | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute_name = $_
        switch ($this_attribute_name) {
            "record_status" { $prefixed_attribute_name = $this_attribute_name }
            default { $prefixed_attribute_name = "ad_{0}" -F $this_attribute_name }

        }
        
        $varchar_field += $prefixed_attribute_name
        $prefixed_fields = $prefixed_fields | Select-Object *, $prefixed_attribute_name
        $prefixed_fields."$prefixed_attribute_name" = $Fields."$this_attribute_name"

    }
    $Fields = $prefixed_fields
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "prefixed_fields:{0}" -F $prefixed_fields | ConvertTo-Json -Compress    )

    #endregion

    #region function internal definitions
    $table = "ADTidy_Inventory_Users"
    #$sql_varchar_fields = @("entry_type", "entry_details", "user_type", "user_employeeid", "user_samaccountname", "user_guid", "execution_status_details", "execution_mode", "execution_operator_name", "execution_operator_action_timestamp", "execution_logs", "execution_last_update", "record_status")
    #$sql_special_fields = @("entry_repeat", "entry_first_occurrence", "entry_last_occurrence", "execution_status_code")
    #endregion

    #region create INSERT statement field and value pairs, loop throuhg $Fields array matching sql_field definition
    $sql_statement_fields = ""
    $sql_statement_values = ""
    $sql_update_statement = "$sql_update_statement"
    $varchar_field | ForEach-Object {
        $thisField = $_
        if ( $Fields."$thisField" -ne $null ) {
            #Global:Log -text (" + `$Fields contains attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
            $sql_statement_fields = "{0} [{1}]," -F $sql_statement_fields, $thisField
            if ($Fields."$thisField" -eq "NULL") {
                $sql_statement_values = "{0} {1}," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]={2}," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            else {
                $sql_statement_values = "{0} '{1}'," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]='{2}'," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            
            
        }
        else {
            #Global:Log -text (" ! `$Fields misses attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
        }
    }
    #endregion

    #region check if this user ( guid based ) exists 
    $exist_query = "select * FROM {0} WHERE ad_objectguid ='{1}'" -F $Table_Name, $Fields.ad_objectguid
    $current_record = Global:sql_query -query $exist_query
    if ( $current_record.count -ne 0) {
        $action_type = "update"
    }
    else {
        $action_type = "new"
    }
    #endregion

    
    #region new record
    if ( $action_type -eq "new") {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = new"  )


        $sql_statement_fields = "{0} [record_source]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'ADTidy'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_status]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'Current'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_lastupdate]," -F $sql_statement_fields
        $sql_statement_values = "{0} GETDATE()," -F $sql_statement_values
        
        # remove last ',' from both fields and values strings
        $sql_statement_fields = $sql_statement_fields -replace ".$"
        $sql_statement_values = $sql_statement_values -replace ".$"

        $sql_statement_insert = " INSERT INTO {0} ({1}) VALUES ({2})" -F $table, $sql_statement_fields, $sql_statement_values
        
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_insert  )
        $insert_result = Global:sql_query -query $sql_statement_insert


        #Global:Log -text ("Inserted row id: '{0}' " -F $insert_result.rowid) -hierarchy "function:adimport_sql_update:DEBUG"
        return $action_type
    }
    #endregion


    #region update record
    if ( $action_type -eq "update" ) {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = update"  )

        $sql_update_filter = " [ad_ObjectGUID] = '{0}'" -F $ObjectGUID

        $sql_update_statement = "{0} [record_lastupdate] = GETDATE()," -F $sql_update_statement

        # remove last ',' from string
        $sql_update_statement = $sql_update_statement -replace ".$"

        $sql_statement_update = " UPDATE {0} SET {1} WHERE {2}" -F $table, $sql_update_statement, $sql_update_filter
        #Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_update  )
        Global:sql_query -query $sql_statement_update
        return $action_type
    }
    #endregion




}
function Global_ADTidy_Iventory_Users_last_update {
    param(
        $Table_Name = "ADTidy_Inventory_Users"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving last run timestamp..." )

    return Global:sql_query -query ("SELECT max(ad_whenchanged) as maxrecord  FROM [{0}]" -F $Table_Name )
}
function Global_ADTidy_Iventory_Users_all_current_records {
    param(
        $Table_Name = "ADTidy_Inventory_Users"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving all user records..." )

    return Global:sql_query -query ("SELECT [ad_samaccountname],[ad_objectguid],[ad_sid] FROM [ittool].[dbo].[{0}] WHERE [record_status] = 'Current'" -F $Table_Name )
}




function Global:ADTidy_Inventory_OU_sql_table_check {
    param(
        $Table_Name = "ADTidy_Inventory_Organizational_Unit"
    )

    $config = @"
{"Fields": [{"name": "record_source","type": "VARCHAR(100)"},{"name": "record_lastupdate","type": "DATETIME"},{"name": "record_status","type": "VARCHAR(50)"},{"name": "ad_whenCreated","type": "DATETIME"},{"name": "ad_whenChanged","type": "DATETIME"},{"name": "ad_distinguishedname","type": "VARCHAR(MAX)"},{"name": "ad_objectguid","type": "VARCHAR(50)"},{"name": "ad_name","type": "VARCHAR(100)"},{"name": "ad_businessCategory","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_managedBy","type": "VARCHAR(MAX)","nullable": 1}],"FieldsAssignement": [{"name": "record_source","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "record_source"}]}]},{"name": "record_lastupdate","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "current_datetime"}]}]},{"name": "record_status","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "status"}]}]},{"name": "ad_whenCreated","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenCreated"}]}]},{"name": "ad_whenChanged","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenChanged"}]}]},{"name": "ad_distinguishedname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "distinguishedname"}]}]},{"name": "ad_objectguid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "objectguid"}]}]},{"name": "ad_name","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "name"}]}]},{"name": "ad_businessCategory","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "businessCategory"}]}]},{"name": "ad_managedBy","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "ManagedBy"}]}]}]}										
"@ | ConvertFrom-Json


    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Running check for table:{0}" -F $Table_Name) -type warning
    $Query_Check_Table_Exists = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{0}'" -F $Table_Name
    $Result_Check_Table_Exists = Global:sql_query -query $Query_Check_Table_Exists

    if ( ($Result_Check_Table_Exists | Measure-Object).count -eq 0 ) {
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' does not exist" -F $Table_Name) -type warning
        $Query_Create_Table_main = "CREATE TABLE {0} ({1}{2});"
        $Query_Create_Table_Constraint = "CONSTRAINT {0} UNIQUE ({1})" 
        
        $Query_String_Fields = ""
        $Query_String_Constraints = ""
        $Has_Contraints = 0
        $config.Fields | ForEach-Object {
            if ($_.nullable -ne 1) { $string_nullable = " NOT NULL" } else { $string_nullable = " NULL" }
            if ($_.id -eq 1) { $string_id = " IDENTITY(1,1) NOT NULL"; $string_nullable = $null } else { $string_id = $null }
            $Query_String_Fields += "`n{0} {1}{2}{3}," -F $_.name, $_.type, $string_id, $string_nullable
            if ($_.constraints -eq 1) { $Query_String_Constraints += "{0}," -F $_.Name; $Has_Contraints = 1 }
            
    
        }
        $Query_String_Fields = $Query_String_Fields.Substring(1)# removes first character  from string
        if ( $Has_Contraints -eq 0 ) {
            #write-host "no constraints"
            $Query_String_Fields = $Query_String_Fields -replace ".$" # removes last character from string
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, $Query_String_Fields, $null
        }
        else {
            #write-host "with constraints"
            $Query_String_Constraints = $Query_String_Constraints -replace ".$" # removes last character from string
            $Query_String_Constraints = $Query_Create_Table_Constraint -F ("CONSTRAINT_{0}" -F $Table_Name), $Query_String_Constraints
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, ("`n" + $Query_String_Fields), ("`n" + $Query_String_Constraints)
        }
        IF ( $Global:WhatIf -ne $true ) {
            Global:sql_query -query $Query_Final_Create_Query 
        }
        ELSE {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Query_Final_Create_Query ) -type warning
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
        }


    }
    else {
        $Table_schema_name = "[{0}].[{1}].[{2}]" -F $Result_Check_Table_Exists.TABLE_CATALOG, $Result_Check_Table_Exists.TABLE_SCHEMA, $Result_Check_Table_Exists.TABLE_NAME
        
        $Query_Select_all = "SELECT * FROM {0}" -F $Table_Name
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' exists, {1} rows in it" -F $Table_Name, (Global:sql_query -query $Query_Select_all).count) 
        
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text "Verifiying columns..."
        $Query_List_columns = "SELECT col.name AS column_name,t.name AS data_type,col.max_length AS data_type_detail FROM sys.tables AS tab INNER JOIN sys.columns AS col ON tab.object_id = col.object_id LEFT JOIN sys.types AS t ON col.user_type_id = t.user_type_id WHERE tab.name = '{0}'" -F $Table_Name
        $SqlCurrentColumns = Global:sql_query -query $Query_List_columns

        $Queries_Update_table = ""
        $config.Fields | ForEach-Object {
            $this = $_ | Select-Object *, type_sql, type_detail

            #write-host ( "split of {0}, count:{1}" -F $this.type, ($this.type).split("(").count )
            if ( ($this.type).split("(").count -eq 1) {
                # no detail type such as DATETIME or INT
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = $null
            }
            else {
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = ($this.type).split("(")[1] -replace ".$"
            }
            #write-host ( "json this:{0}" -F ( $this | ConvertTo-Json -Compress ))
            
            $MatchingColumn = "" | Select-Object found, same_type, type_details_match
            $MatchingColumn.found = 0
            $MatchingColumn.same_type = 0
            $MatchingColumn.type_details_match = 0
            $SqlCurrentColumns | Where-Object { $_.column_name -eq $this.name } | ForEach-Object {
                $MatchingColumn.found = 1
                if ( $_.data_type -eq $this.type_sql ) { 
                    $MatchingColumn.same_type = 1
                    if ($_.data_type_detail -eq -1) { $temp_data_type_detail = [string]"MAX" } ELSE { $temp_data_type_detail = [string]$_.data_type_detail }
                    if ( $this.type_detail -ne $null) {
                        if ( $temp_data_type_detail -eq $this.type_detail) {
                            $MatchingColumn.type_details_match = 1
                        }
                        else {
                            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type detail for column ({3}) {2} :({0} <> {1})." -F $this.type_detail, $temp_data_type_detail, $this.name, $this.type) -type warning
                        }
                    }
                    else {
                        $MatchingColumn.type_details_match = 1
                    }

                }
                else {
                    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type for column {2} ({0} <> {1})." -F $this.type_sql, $_.data_type, $this.name) -type warning
                }
                #write-host ( "json MatchingColumn:{0}" -F ( $MatchingColumn | ConvertTo-Json -Compress ))

            }
        

            if ( $MatchingColumn.found -eq 0 ) {  
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "column {0} missing" -F $this.name) -type warning

            }

            $Query_Run = $null
            $Query_Alter_Table_Add_Column = "ALTER TABLE {0} ADD {1} {2};"
            $Query_Alter_Table_Alter_Column = "ALTER TABLE {0} ALTER COLUMN {1} {2};"

            if ( $MatchingColumn.found -eq 0 ) { 
                $Query_Run = $Query_Alter_Table_Add_Column 
            }
            else {
                if ( $MatchingColumn.same_type -eq 0 -or $MatchingColumn.type_details_match -eq 0 ) {
                    $Query_Run = $Query_Alter_Table_Alter_Column 
                }
            }

            if ( $Query_Run -ne $null ) {
                $Query_Run = $Query_Run -F $Table_schema_name, $this.name, $this.type
                $Queries_Update_table += $Query_Run + "`n"
            }

        }

        if ( $Queries_Update_table.Length -ge 1 ) {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "resulting alter table queries:{0}" -F $Queries_Update_table)
            IF ( $Global:WhatIf -ne $true ) {
                Global:sql_query -query $Queries_Update_table 
            }
            ELSE {
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Queries_Update_table ) -type warning
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
            }

        }
        else {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Table {0} is matching the definition of the config.json at run time." -F $Table_Name)
        }

        
    }
}
function Global:ADTidy_Inventory_OU_sql_update {
    param(
        [Parameter(Mandatory = $true)] [array]$Fields,
        $Table_Name = "ADTidy_Inventory_Organizational_Unit"
    )

    #region ADTidy_Inventory_Users_sql_update specific
    $prefixed_fields = "" | Select-Object ignore
    $varchar_field = @()
    $ObjectGUID = $Fields.ObjectGUID
    $Fields | Select-Object -ExcludeProperty ObjectGUID | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute_name = $_
        switch ($this_attribute_name) {
            "record_status" { $prefixed_attribute_name = $this_attribute_name }
            default { $prefixed_attribute_name = "ad_{0}" -F $this_attribute_name }

        }
        
        $varchar_field += $prefixed_attribute_name
        $prefixed_fields = $prefixed_fields | Select-Object *, $prefixed_attribute_name
        $prefixed_fields."$prefixed_attribute_name" = $Fields."$this_attribute_name"

    }
    $Fields = $prefixed_fields
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "prefixed_fields:{0}" -F $prefixed_fields | ConvertTo-Json -Compress    )

    #endregion

    #region function internal definitions
    $table = $Table_Name
    #$sql_varchar_fields = @("entry_type", "entry_details", "user_type", "user_employeeid", "user_samaccountname", "user_guid", "execution_status_details", "execution_mode", "execution_operator_name", "execution_operator_action_timestamp", "execution_logs", "execution_last_update", "record_status")
    #$sql_special_fields = @("entry_repeat", "entry_first_occurrence", "entry_last_occurrence", "execution_status_code")
    #endregion

    #region create INSERT statement field and value pairs, loop throuhg $Fields array matching sql_field definition
    $sql_statement_fields = ""
    $sql_statement_values = ""
    $sql_update_statement = "$sql_update_statement"
    $varchar_field | ForEach-Object {
        $thisField = $_
        if ( $Fields."$thisField" -ne $null ) {
            #Global:Log -text (" + `$Fields contains attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
            $sql_statement_fields = "{0} [{1}]," -F $sql_statement_fields, $thisField
            if ($Fields."$thisField" -eq "NULL") {
                $sql_statement_values = "{0} {1}," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]={2}," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            else {
                $sql_statement_values = "{0} '{1}'," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]='{2}'," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            
            
        }
        else {
            #Global:Log -text (" ! `$Fields misses attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
        }
    }
    #endregion

    #region check if this user ( guid based ) exists 
    $exist_query = "select * FROM {0} WHERE ad_objectguid ='{1}'" -F $Table_Name, $Fields.ad_objectguid
    $current_record = Global:sql_query -query $exist_query
    if ( $current_record.count -ne 0) {
        $action_type = "update"
    }
    else {
        $action_type = "new"
    }
    #endregion

    
    #region new record
    if ( $action_type -eq "new") {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = new"  )


        $sql_statement_fields = "{0} [record_source]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'ADTidy'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_status]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'Current'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_lastupdate]," -F $sql_statement_fields
        $sql_statement_values = "{0} GETDATE()," -F $sql_statement_values
        
        # remove last ',' from both fields and values strings
        $sql_statement_fields = $sql_statement_fields -replace ".$"
        $sql_statement_values = $sql_statement_values -replace ".$"

        $sql_statement_insert = " INSERT INTO {0} ({1}) VALUES ({2})" -F $table, $sql_statement_fields, $sql_statement_values
        
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_insert  )
        $insert_result = Global:sql_query -query $sql_statement_insert


        #Global:Log -text ("Inserted row id: '{0}' " -F $insert_result.rowid) -hierarchy "function:adimport_sql_update:DEBUG"
        return $action_type
    }
    #endregion


    #region update record
    if ( $action_type -eq "update" ) {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = update"  )

        $sql_update_filter = " [ad_ObjectGUID] = '{0}'" -F $ObjectGUID

        $sql_update_statement = "{0} [record_lastupdate] = GETDATE()," -F $sql_update_statement

        # remove last ',' from string
        $sql_update_statement = $sql_update_statement -replace ".$"

        $sql_statement_update = " UPDATE {0} SET {1} WHERE {2}" -F $table, $sql_update_statement, $sql_update_filter
        #Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_update  )
        Global:sql_query -query $sql_statement_update
        return $action_type
    }
    #endregion




}
function Global_ADTidy_Iventory_OU_last_update {
    param(
        $Table_Name = "ADTidy_Inventory_Organizational_Unit"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving last run timestamp..." )

    return Global:sql_query -query ("SELECT max(ad_whenchanged) as maxrecord  FROM [{0}]" -F $Table_Name )
}
function Global_ADTidy_Iventory_OU_all_current_records {
    param(
        $Table_Name = "ADTidy_Inventory_Organizational_Unit"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrievingall OU..." )

    return Global:sql_query -query ("SELECT [ad_name],[ad_objectguid],[ad_distinguishedname] FROM [ittool].[dbo].[{0}] WHERE [record_status] = 'Current'" -F $Table_Name )
}



function Global:ADTidy_Inventory_Groups_sql_table_check {
    param(
        $Table_Name = "ADTidy_Inventory_Groups"
    )

    $config = @"
{"Fields": [{"name": "record_source","type": "VARCHAR(100)"},{"name": "record_lastupdate","type": "DATETIME"},{"name": "record_status","type": "VARCHAR(50)"},{"name": "ad_GroupCategory","type": "VARCHAR(50)"},{"name": "ad_GroupScope","type": "VARCHAR(50)"},{"name": "ad_whenCreated","type": "DATETIME"},{"name": "ad_whenChanged","type": "DATETIME"},{"name": "ad_sid","type": "VARCHAR(50)"},{"name": "ad_objectguid","type": "VARCHAR(50)"},{"name": "ad_samaccountname","type": "VARCHAR(50)"},{"name": "ad_name","type": "VARCHAR(100)"},{"name": "ad_distinguishedname","type": "VARCHAR(MAX)"},{"name": "ad_description","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_info","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_managedBy","type": "VARCHAR(MAX)","nullable": 1},{"name": "xml_members","type": "XML","nullable": 1},{"name": "ad_extensionattribute1","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_extensionattribute2","type": "VARCHAR(50)","nullable": 1}],"FieldsAssignement": [{"name": "record_source","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "record_source"}]}]},{"name": "record_lastupdate","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "current_datetime"}]}]},{"name": "record_status","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "status"}]}]},{"name": "ad_GroupCategory","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "groupcategory"}]}]},{"name": "ad_GroupScope","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "groupscope"}]}]},{"name": "ad_whenCreated","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenCreated"}]}]},{"name": "ad_whenChanged","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenChanged"}]}]},{"name": "ad_sid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "sid"}]}]},{"name": "ad_objectguid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "objectguid"}]}]},{"name": "ad_samaccountname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "samaccountname"}]}]},{"name": "ad_name","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "name"}]}]},{"name": "ad_distinguishedname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "distinguishedname"}]}]},{"name": "ad_description","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "description"}]}]},{"name": "ad_info","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "info"}]}]},{"name": "ad_managedBy","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "managedBy"}]}]},{"name": "xml_members","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "members"}]}]},{"name": "ad_extensionattribute1","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "businessCategory"}]}]},{"name": "ad_extensionattribute2","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "ManagedBy"}]}]}]}										
"@ | ConvertFrom-Json


    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Running check for table:{0}" -F $Table_Name) -type warning
    $Query_Check_Table_Exists = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{0}'" -F $Table_Name
    $Result_Check_Table_Exists = Global:sql_query -query $Query_Check_Table_Exists

    if ( ($Result_Check_Table_Exists | Measure-Object).count -eq 0 ) {
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' does not exist" -F $Table_Name) -type warning
        $Query_Create_Table_main = "CREATE TABLE {0} ({1}{2});"
        $Query_Create_Table_Constraint = "CONSTRAINT {0} UNIQUE ({1})" 
        
        $Query_String_Fields = ""
        $Query_String_Constraints = ""
        $Has_Contraints = 0
        $config.Fields | ForEach-Object {
            if ($_.nullable -ne 1) { $string_nullable = " NOT NULL" } else { $string_nullable = " NULL" }
            if ($_.id -eq 1) { $string_id = " IDENTITY(1,1) NOT NULL"; $string_nullable = $null } else { $string_id = $null }
            $Query_String_Fields += "`n{0} {1}{2}{3}," -F $_.name, $_.type, $string_id, $string_nullable
            if ($_.constraints -eq 1) { $Query_String_Constraints += "{0}," -F $_.Name; $Has_Contraints = 1 }
            
    
        }
        $Query_String_Fields = $Query_String_Fields.Substring(1)# removes first character  from string
        if ( $Has_Contraints -eq 0 ) {
            #write-host "no constraints"
            $Query_String_Fields = $Query_String_Fields -replace ".$" # removes last character from string
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, $Query_String_Fields, $null
        }
        else {
            #write-host "with constraints"
            $Query_String_Constraints = $Query_String_Constraints -replace ".$" # removes last character from string
            $Query_String_Constraints = $Query_Create_Table_Constraint -F ("CONSTRAINT_{0}" -F $Table_Name), $Query_String_Constraints
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, ("`n" + $Query_String_Fields), ("`n" + $Query_String_Constraints)
        }
        IF ( $Global:WhatIf -ne $true ) {
            Global:sql_query -query $Query_Final_Create_Query 
        }
        ELSE {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Query_Final_Create_Query ) -type warning
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
        }


    }
    else {
        $Table_schema_name = "[{0}].[{1}].[{2}]" -F $Result_Check_Table_Exists.TABLE_CATALOG, $Result_Check_Table_Exists.TABLE_SCHEMA, $Result_Check_Table_Exists.TABLE_NAME
        
        $Query_Select_all = "SELECT * FROM {0}" -F $Table_Name
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' exists, {1} rows in it" -F $Table_Name, (Global:sql_query -query $Query_Select_all).count) 
        
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text "Verifiying columns..."
        $Query_List_columns = "SELECT col.name AS column_name,t.name AS data_type,col.max_length AS data_type_detail FROM sys.tables AS tab INNER JOIN sys.columns AS col ON tab.object_id = col.object_id LEFT JOIN sys.types AS t ON col.user_type_id = t.user_type_id WHERE tab.name = '{0}'" -F $Table_Name
        $SqlCurrentColumns = Global:sql_query -query $Query_List_columns

        $Queries_Update_table = ""
        $config.Fields | ForEach-Object {
            $this = $_ | Select-Object *, type_sql, type_detail

            #write-host ( "split of {0}, count:{1}" -F $this.type, ($this.type).split("(").count )
            if ( ($this.type).split("(").count -eq 1) {
                # no detail type such as DATETIME or INT
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = $null
            }
            else {
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = ($this.type).split("(")[1] -replace ".$"
            }
            #write-host ( "json this:{0}" -F ( $this | ConvertTo-Json -Compress ))
            
            $MatchingColumn = "" | Select-Object found, same_type, type_details_match
            $MatchingColumn.found = 0
            $MatchingColumn.same_type = 0
            $MatchingColumn.type_details_match = 0
            $SqlCurrentColumns | Where-Object { $_.column_name -eq $this.name } | ForEach-Object {
                $MatchingColumn.found = 1
                if ( $_.data_type -eq $this.type_sql ) { 
                    $MatchingColumn.same_type = 1
                    if ($_.data_type_detail -eq -1) { $temp_data_type_detail = [string]"MAX" } ELSE { $temp_data_type_detail = [string]$_.data_type_detail }
                    if ( $this.type_detail -ne $null) {
                        if ( $temp_data_type_detail -eq $this.type_detail) {
                            $MatchingColumn.type_details_match = 1
                        }
                        else {
                            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type detail for column ({3}) {2} :({0} <> {1})." -F $this.type_detail, $temp_data_type_detail, $this.name, $this.type) -type warning
                        }
                    }
                    else {
                        $MatchingColumn.type_details_match = 1
                    }

                }
                else {
                    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type for column {2} ({0} <> {1})." -F $this.type_sql, $_.data_type, $this.name) -type warning
                }
                #write-host ( "json MatchingColumn:{0}" -F ( $MatchingColumn | ConvertTo-Json -Compress ))

            }
        

            if ( $MatchingColumn.found -eq 0 ) {  
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "column {0} missing" -F $this.name) -type warning

            }

            $Query_Run = $null
            $Query_Alter_Table_Add_Column = "ALTER TABLE {0} ADD {1} {2};"
            $Query_Alter_Table_Alter_Column = "ALTER TABLE {0} ALTER COLUMN {1} {2};"

            if ( $MatchingColumn.found -eq 0 ) { 
                $Query_Run = $Query_Alter_Table_Add_Column 
            }
            else {
                if ( $MatchingColumn.same_type -eq 0 -or $MatchingColumn.type_details_match -eq 0 ) {
                    $Query_Run = $Query_Alter_Table_Alter_Column 
                }
            }

            if ( $Query_Run -ne $null ) {
                $Query_Run = $Query_Run -F $Table_schema_name, $this.name, $this.type
                $Queries_Update_table += $Query_Run + "`n"
            }

        }

        if ( $Queries_Update_table.Length -ge 1 ) {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "resulting alter table queries:{0}" -F $Queries_Update_table)
            IF ( $Global:WhatIf -ne $true ) {
                Global:sql_query -query $Queries_Update_table 
            }
            ELSE {
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Queries_Update_table ) -type warning
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
            }

        }
        else {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Table {0} is matching the definition of the config.json at run time." -F $Table_Name)
        }

        
    }
}
function Global_ADTidy_Iventory_Groups_last_update {
    param(
        $Table_Name = "ADTidy_Inventory_Groups"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving last run timestamp..." )

    return Global:sql_query -query ("SELECT max(ad_whenchanged) as maxrecord  FROM [{0}]" -F $Table_Name )
}
function Global:ADTidy_Inventory_Groups_sql_update {
    param(
        [Parameter(Mandatory = $true)] [array]$Fields,
        $Table_Name = "ADTidy_Inventory_Groups"
    )

    #region ADTidy_Inventory_Users_sql_update specific
    $prefixed_fields = "" | Select-Object ignore
    $varchar_field = @()
    $ObjectGUID = $Fields.ObjectGUID
    $Fields | Select-Object -ExcludeProperty ObjectGUID | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute_name = $_
        switch ($this_attribute_name) {
            "record_status" { $prefixed_attribute_name = $this_attribute_name }
            "xml_members" { $prefixed_attribute_name = $this_attribute_name }
            default { $prefixed_attribute_name = "ad_{0}" -F $this_attribute_name }

        }
        
        $varchar_field += $prefixed_attribute_name
        $prefixed_fields = $prefixed_fields | Select-Object *, $prefixed_attribute_name
        $prefixed_fields."$prefixed_attribute_name" = $Fields."$this_attribute_name"

    }
    $Fields = $prefixed_fields
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "prefixed_fields:{0}" -F $prefixed_fields | ConvertTo-Json -Compress    )

    #endregion

    #region function internal definitions
    $table = $Table_Name
    #$sql_varchar_fields = @("entry_type", "entry_details", "user_type", "user_employeeid", "user_samaccountname", "user_guid", "execution_status_details", "execution_mode", "execution_operator_name", "execution_operator_action_timestamp", "execution_logs", "execution_last_update", "record_status")
    #$sql_special_fields = @("entry_repeat", "entry_first_occurrence", "entry_last_occurrence", "execution_status_code")
    #endregion

    #region create INSERT statement field and value pairs, loop throuhg $Fields array matching sql_field definition
    $sql_statement_fields = ""
    $sql_statement_values = ""
    $sql_update_statement = "$sql_update_statement"
    $varchar_field | ForEach-Object {
        $thisField = $_
        if ( $Fields."$thisField" -ne $null ) {
            #Global:Log -text (" + `$Fields contains attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
            $sql_statement_fields = "{0} [{1}]," -F $sql_statement_fields, $thisField
            if ($Fields."$thisField" -eq "NULL") {
                $sql_statement_values = "{0} {1}," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]={2}," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            else {
                $sql_statement_values = "{0} '{1}'," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]='{2}'," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            
            
        }
        else {
            #Global:Log -text (" ! `$Fields misses attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
        }
    }
    #endregion

    #region check if this user ( guid based ) exists 
    $exist_query = "select * FROM {0} WHERE ad_objectguid ='{1}'" -F $Table_Name, $Fields.ad_objectguid
    $current_record = Global:sql_query -query $exist_query
    if ( $current_record.count -ne 0) {
        $action_type = "update"
    }
    else {
        $action_type = "new"
    }
    #endregion

    
    #region new record
    if ( $action_type -eq "new") {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = new"  )


        $sql_statement_fields = "{0} [record_source]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'ADTidy'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_status]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'Current'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_lastupdate]," -F $sql_statement_fields
        $sql_statement_values = "{0} GETDATE()," -F $sql_statement_values
        
        # remove last ',' from both fields and values strings
        $sql_statement_fields = $sql_statement_fields -replace ".$"
        $sql_statement_values = $sql_statement_values -replace ".$"

        $sql_statement_insert = " INSERT INTO {0} ({1}) VALUES ({2})" -F $table, $sql_statement_fields, $sql_statement_values
        
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_insert  )
        $insert_result = Global:sql_query -query $sql_statement_insert


        #Global:Log -text ("Inserted row id: '{0}' " -F $insert_result.rowid) -hierarchy "function:adimport_sql_update:DEBUG"
        return $action_type

    }
    #endregion


    #region update record
    if ( $action_type -eq "update" ) {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = update"  )

        $sql_update_filter = " [ad_ObjectGUID] = '{0}'" -F $ObjectGUID

        $sql_update_statement = "{0} [record_lastupdate] = GETDATE()," -F $sql_update_statement

        # remove last ',' from string
        $sql_update_statement = $sql_update_statement -replace ".$"

        $sql_statement_update = " UPDATE {0} SET {1} WHERE {2}" -F $table, $sql_update_statement, $sql_update_filter
        #Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_update  )
        Global:sql_query -query $sql_statement_update
        return $action_type
    }
    #endregion




}
function Global_ADTidy_Iventory_Groups_all_current_records {
    param(
        $Table_Name = "ADTidy_Inventory_Groups"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving all group records" )

    return Global:sql_query -query ("SELECT [ad_name],[ad_objectguid],[ad_distinguishedname] FROM [ittool].[dbo].[{0}] WHERE [record_status] = 'Current'" -F $Table_Name )
}


function Global:ADTidy_Inventory_Computers_sql_table_check {
    param(
        $Table_Name = "ADTidy_Inventory_Computers"
    )

    $config = @"
{"Fields": [{"name": "record_source","type": "VARCHAR(100)"},{"name": "record_lastupdate","type": "DATETIME"},{"name": "ad_lastlogontimestamp","type": "DATETIME","nullable": 1},{"name": "record_status","type": "VARCHAR(50)"},{"name": "ad_whenCreated","type": "DATETIME"},{"name": "ad_whenChanged","type": "DATETIME"},{"name": "ad_objectguid","type": "VARCHAR(50)"},{"name": "ad_samaccountname","type": "VARCHAR(50)"},{"name": "ad_name","type": "VARCHAR(100)"},{"name": "ad_distinguishedname","type": "VARCHAR(MAX)"},{"name": "ad_description","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_info","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_managedBy","type": "VARCHAR(MAX)","nullable": 1},{"name": "ad_operatingSystem","type": "VARCHAR(100)","nullable": 1},{"name": "ad_operatingSystemVersion","type": "VARCHAR(50)","nullable": 1},{"name": "ad_extensionattribute1","type": "VARCHAR(50)","nullable": 1},{"name": "ad_extensionattribute2","type": "VARCHAR(50)","nullable": 1},{"name": "ad_extensionattribute3","type": "VARCHAR(50)","nullable": 1},{"name": "ad_extensionattribute4","type": "VARCHAR(50)","nullable": 1},{"name": "ad_sid","type": "VARCHAR(50)"}],"FieldsAssignement": [{"name": "record_source","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "record_source"}]}]},{"name": "record_lastupdate","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Calculation","Type": "current_datetime"}]}]},{"name": "ad_lastlogontimestamp","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "lastlogontimestamp"}]}]},{"name": "record_status","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "status"}]}]},{"name": "ad_whenCreated","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenCreated"}]}]},{"name": "ad_whenChanged","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "whenChanged"}]}]},{"name": "ad_objectguid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "objectguid"}]}]},{"name": "ad_samaccountname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "samaccountname"}]}]},{"name": "ad_name","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "name"}]}]},{"name": "ad_distinguishedname","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "distinguishedname"}]}]},{"name": "ad_description","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "description"}]}]},{"name": "ad_info","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "info"}]}]},{"name": "ad_managedBy","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "managedBy"}]}]},{"name": "ad_operatingSystem","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "os1"}]}]},{"name": "ad_operatingSystemVersion","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "os2"}]}]},{"name": "ad_extensionattribute1","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "ea1"}]}]},{"name": "ad_extensionattribute2","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "ea2"}]}]},{"name": "ad_extensionattribute3","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "ea3"}]}]},{"name": "ad_extensionattribute4","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "ea4"}]}]},{"name": "ad_sid","Recipe":[{"ORDER":"1", "Content": [ {"Source": "Field","Type": "sid"}]}]}]}										
"@ | ConvertFrom-Json


    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Running check for table:{0}" -F $Table_Name) -type warning
    $Query_Check_Table_Exists = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{0}'" -F $Table_Name
    $Result_Check_Table_Exists = Global:sql_query -query $Query_Check_Table_Exists

    if ( ($Result_Check_Table_Exists | Measure-Object).count -eq 0 ) {
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' does not exist" -F $Table_Name) -type warning
        $Query_Create_Table_main = "CREATE TABLE {0} ({1}{2});"
        $Query_Create_Table_Constraint = "CONSTRAINT {0} UNIQUE ({1})" 
        
        $Query_String_Fields = ""
        $Query_String_Constraints = ""
        $Has_Contraints = 0
        $config.Fields | ForEach-Object {
            if ($_.nullable -ne 1) { $string_nullable = " NOT NULL" } else { $string_nullable = " NULL" }
            if ($_.id -eq 1) { $string_id = " IDENTITY(1,1) NOT NULL"; $string_nullable = $null } else { $string_id = $null }
            $Query_String_Fields += "`n{0} {1}{2}{3}," -F $_.name, $_.type, $string_id, $string_nullable
            if ($_.constraints -eq 1) { $Query_String_Constraints += "{0}," -F $_.Name; $Has_Contraints = 1 }
            
    
        }
        $Query_String_Fields = $Query_String_Fields.Substring(1)# removes first character  from string
        if ( $Has_Contraints -eq 0 ) {
            #write-host "no constraints"
            $Query_String_Fields = $Query_String_Fields -replace ".$" # removes last character from string
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, $Query_String_Fields, $null
        }
        else {
            #write-host "with constraints"
            $Query_String_Constraints = $Query_String_Constraints -replace ".$" # removes last character from string
            $Query_String_Constraints = $Query_Create_Table_Constraint -F ("CONSTRAINT_{0}" -F $Table_Name), $Query_String_Constraints
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, ("`n" + $Query_String_Fields), ("`n" + $Query_String_Constraints)
        }
        IF ( $Global:WhatIf -ne $true ) {
            Global:sql_query -query $Query_Final_Create_Query 
        }
        ELSE {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Query_Final_Create_Query ) -type warning
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
        }


    }
    else {
        $Table_schema_name = "[{0}].[{1}].[{2}]" -F $Result_Check_Table_Exists.TABLE_CATALOG, $Result_Check_Table_Exists.TABLE_SCHEMA, $Result_Check_Table_Exists.TABLE_NAME
        
        $Query_Select_all = "SELECT * FROM {0}" -F $Table_Name
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' exists, {1} rows in it" -F $Table_Name, (Global:sql_query -query $Query_Select_all).count) 
        
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text "Verifiying columns..."
        $Query_List_columns = "SELECT col.name AS column_name,t.name AS data_type,col.max_length AS data_type_detail FROM sys.tables AS tab INNER JOIN sys.columns AS col ON tab.object_id = col.object_id LEFT JOIN sys.types AS t ON col.user_type_id = t.user_type_id WHERE tab.name = '{0}'" -F $Table_Name
        $SqlCurrentColumns = Global:sql_query -query $Query_List_columns

        $Queries_Update_table = ""
        $config.Fields | ForEach-Object {
            $this = $_ | Select-Object *, type_sql, type_detail

            #write-host ( "split of {0}, count:{1}" -F $this.type, ($this.type).split("(").count )
            if ( ($this.type).split("(").count -eq 1) {
                # no detail type such as DATETIME or INT
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = $null
            }
            else {
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = ($this.type).split("(")[1] -replace ".$"
            }
            #write-host ( "json this:{0}" -F ( $this | ConvertTo-Json -Compress ))
            
            $MatchingColumn = "" | Select-Object found, same_type, type_details_match
            $MatchingColumn.found = 0
            $MatchingColumn.same_type = 0
            $MatchingColumn.type_details_match = 0
            $SqlCurrentColumns | Where-Object { $_.column_name -eq $this.name } | ForEach-Object {
                $MatchingColumn.found = 1
                if ( $_.data_type -eq $this.type_sql ) { 
                    $MatchingColumn.same_type = 1
                    if ($_.data_type_detail -eq -1) { $temp_data_type_detail = [string]"MAX" } ELSE { $temp_data_type_detail = [string]$_.data_type_detail }
                    if ( $this.type_detail -ne $null) {
                        if ( $temp_data_type_detail -eq $this.type_detail) {
                            $MatchingColumn.type_details_match = 1
                        }
                        else {
                            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type detail for column ({3}) {2} :({0} <> {1})." -F $this.type_detail, $temp_data_type_detail, $this.name, $this.type) -type warning
                        }
                    }
                    else {
                        $MatchingColumn.type_details_match = 1
                    }

                }
                else {
                    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type for column {2} ({0} <> {1})." -F $this.type_sql, $_.data_type, $this.name) -type warning
                }
                #write-host ( "json MatchingColumn:{0}" -F ( $MatchingColumn | ConvertTo-Json -Compress ))

            }
        

            if ( $MatchingColumn.found -eq 0 ) {  
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "column {0} missing" -F $this.name) -type warning

            }

            $Query_Run = $null
            $Query_Alter_Table_Add_Column = "ALTER TABLE {0} ADD {1} {2};"
            $Query_Alter_Table_Alter_Column = "ALTER TABLE {0} ALTER COLUMN {1} {2};"

            if ( $MatchingColumn.found -eq 0 ) { 
                $Query_Run = $Query_Alter_Table_Add_Column 
            }
            else {
                if ( $MatchingColumn.same_type -eq 0 -or $MatchingColumn.type_details_match -eq 0 ) {
                    $Query_Run = $Query_Alter_Table_Alter_Column 
                }
            }

            if ( $Query_Run -ne $null ) {
                $Query_Run = $Query_Run -F $Table_schema_name, $this.name, $this.type
                $Queries_Update_table += $Query_Run + "`n"
            }

        }

        if ( $Queries_Update_table.Length -ge 1 ) {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "resulting alter table queries:{0}" -F $Queries_Update_table)
            IF ( $Global:WhatIf -ne $true ) {
                Global:sql_query -query $Queries_Update_table 
            }
            ELSE {
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Queries_Update_table ) -type warning
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
            }

        }
        else {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Table {0} is matching the definition of the config.json at run time." -F $Table_Name)
        }

        
    }
}
function Global_ADTidy_Iventory_Computers_last_update {
    param(
        $Table_Name = "ADTidy_Inventory_Computers"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving max whenchanged timestamp..." )

    return Global:sql_query -query ("SELECT max(ad_whenchanged) as maxrecord  FROM [{0}]" -F $Table_Name )
}
function Global:ADTidy_Inventory_Computers_sql_update {
    param(
        [Parameter(Mandatory = $true)] [array]$Fields,
        $Table_Name = "ADTidy_Inventory_Computers"
    )

    #region ADTidy_Inventory_Users_sql_update specific
    $prefixed_fields = "" | Select-Object ignore
    $varchar_field = @()
    $ObjectGUID = $Fields.ObjectGUID
    $Fields | Select-Object -ExcludeProperty ObjectGUID | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute_name = $_
        switch ($this_attribute_name) {
            "record_status" { $prefixed_attribute_name = $this_attribute_name }
            "xml_members" { $prefixed_attribute_name = $this_attribute_name }
            default { $prefixed_attribute_name = "ad_{0}" -F $this_attribute_name }

        }
        
        $varchar_field += $prefixed_attribute_name
        $prefixed_fields = $prefixed_fields | Select-Object *, $prefixed_attribute_name
        $prefixed_fields."$prefixed_attribute_name" = $Fields."$this_attribute_name"

    }
    $Fields = $prefixed_fields
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "prefixed_fields:{0}" -F $prefixed_fields | ConvertTo-Json -Compress    )

    #endregion

    #region function internal definitions
    $table = $Table_Name
    #$sql_varchar_fields = @("entry_type", "entry_details", "user_type", "user_employeeid", "user_samaccountname", "user_guid", "execution_status_details", "execution_mode", "execution_operator_name", "execution_operator_action_timestamp", "execution_logs", "execution_last_update", "record_status")
    #$sql_special_fields = @("entry_repeat", "entry_first_occurrence", "entry_last_occurrence", "execution_status_code")
    #endregion

    #region create INSERT statement field and value pairs, loop throuhg $Fields array matching sql_field definition
    $sql_statement_fields = ""
    $sql_statement_values = ""
    $sql_update_statement = "$sql_update_statement"
    $varchar_field | ForEach-Object {
        $thisField = $_
        if ( $Fields."$thisField" -ne $null ) {
            #Global:Log -text (" + `$Fields contains attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
            $sql_statement_fields = "{0} [{1}]," -F $sql_statement_fields, $thisField
            if ($Fields."$thisField" -eq "NULL") {
                $sql_statement_values = "{0} {1}," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]={2}," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            else {
                $sql_statement_values = "{0} '{1}'," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]='{2}'," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            
            
        }
        else {
            #Global:Log -text (" ! `$Fields misses attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
        }
    }
    #endregion

    #region check if this user ( guid based ) exists 
    $exist_query = "select * FROM {0} WHERE ad_objectguid ='{1}'" -F $Table_Name, $Fields.ad_objectguid
    $current_record = Global:sql_query -query $exist_query
    if ( $current_record.count -ne 0) {
        $action_type = "update"
    }
    else {
        $action_type = "new"
    }
    #endregion

    
    #region new record
    if ( $action_type -eq "new") {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = new"  )


        $sql_statement_fields = "{0} [record_source]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'ADTidy'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_status]," -F $sql_statement_fields
        $sql_statement_values = "{0} 'Current'," -F $sql_statement_values
        $sql_statement_fields = "{0} [record_lastupdate]," -F $sql_statement_fields
        $sql_statement_values = "{0} GETDATE()," -F $sql_statement_values
        
        # remove last ',' from both fields and values strings
        $sql_statement_fields = $sql_statement_fields -replace ".$"
        $sql_statement_values = $sql_statement_values -replace ".$"

        $sql_statement_insert = " INSERT INTO {0} ({1}) VALUES ({2})" -F $table, $sql_statement_fields, $sql_statement_values
        
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_insert  )
        $insert_result = Global:sql_query -query $sql_statement_insert


        #Global:Log -text ("Inserted row id: '{0}' " -F $insert_result.rowid) -hierarchy "function:adimport_sql_update:DEBUG"
        return $action_type
    }
    #endregion


    #region update record
    if ( $action_type -eq "update" ) {
        Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = update"  )

        $sql_update_filter = " [ad_ObjectGUID] = '{0}'" -F $ObjectGUID

        $sql_update_statement = "{0} [record_lastupdate] = GETDATE()," -F $sql_update_statement

        # remove last ',' from string
        $sql_update_statement = $sql_update_statement -replace ".$"

        $sql_statement_update = " UPDATE {0} SET {1} WHERE {2}" -F $table, $sql_update_statement, $sql_update_filter
        #Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_update  )
        Global:sql_query -query $sql_statement_update
        return $action_type
    }
    #endregion




}
function Global_ADTidy_Iventory_Computers_all_current_records {
    param(
        $Table_Name = "ADTidy_Inventory_Computers"
    )
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Retrieving all computer records" )

    return Global:sql_query -query ("SELECT [ad_name],[ad_objectguid],[ad_distinguishedname] FROM [ittool].[dbo].[{0}] WHERE [record_status] = 'Current'" -F $Table_Name )
}


function Global:ADTidy_Records_sql_table_check {
    param(
        $Table_Name = "ADTidy_Records"
    )

    $config = @"
{ "Fields": [ { "name": "record_id", "type": "INT", "id": 1 }, { "name": "record_timestamp", "type": "DATETIME" }, { "name": "record_type", "type": "VARCHAR(50)" }, { "name": "rule_name", "type": "VARCHAR(50)", "nullable": 1 }, { "name": "rule_details", "type": "VARCHAR(MAX)", "nullable": 1 }, { "name": "target_list", "type": "VARCHAR(MAX)", "nullable": 1 }, { "name": "result_summary", "type": "VARCHAR(MAX)", "nullable": 1 }, { "name": "log_json", "type": "VARCHAR(MAX)", "nullable": 1 }] }
"@ | ConvertFrom-Json


    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("Running check for table:{0}" -F $Table_Name) -type warning
    $Query_Check_Table_Exists = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{0}'" -F $Table_Name
    $Result_Check_Table_Exists = Global:sql_query -query $Query_Check_Table_Exists

    if ( ($Result_Check_Table_Exists | Measure-Object).count -eq 0 ) {
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' does not exist" -F $Table_Name) -type warning
        $Query_Create_Table_main = "CREATE TABLE {0} ({1}{2});"
        $Query_Create_Table_Constraint = "CONSTRAINT {0} UNIQUE ({1})" 
        
        $Query_String_Fields = ""
        $Query_String_Constraints = ""
        $Has_Contraints = 0
        $config.Fields | ForEach-Object {
            if ($_.nullable -ne 1) { $string_nullable = " NOT NULL" } else { $string_nullable = " NULL" }
            if ($_.id -eq 1) { $string_id = " IDENTITY(1,1) NOT NULL"; $string_nullable = $null } else { $string_id = $null }
            $Query_String_Fields += "`n{0} {1}{2}{3}," -F $_.name, $_.type, $string_id, $string_nullable
            if ($_.constraints -eq 1) { $Query_String_Constraints += "{0}," -F $_.Name; $Has_Contraints = 1 }
            
    
        }
        $Query_String_Fields = $Query_String_Fields.Substring(1)# removes first character  from string
        if ( $Has_Contraints -eq 0 ) {
            #write-host "no constraints"
            $Query_String_Fields = $Query_String_Fields -replace ".$" # removes last character from string
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, $Query_String_Fields, $null
        }
        else {
            #write-host "with constraints"
            $Query_String_Constraints = $Query_String_Constraints -replace ".$" # removes last character from string
            $Query_String_Constraints = $Query_Create_Table_Constraint -F ("CONSTRAINT_{0}" -F $Table_Name), $Query_String_Constraints
            $Query_Final_Create_Query = $Query_Create_Table_main -F $Table_Name, ("`n" + $Query_String_Fields), ("`n" + $Query_String_Constraints)
        }
        IF ( $Global:WhatIf -ne $true ) {
            Global:sql_query -query $Query_Final_Create_Query 
        }
        ELSE {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Query_Final_Create_Query ) -type warning
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
        }


    }
    else {
        $Table_schema_name = "[{0}].[{1}].[{2}]" -F $Result_Check_Table_Exists.TABLE_CATALOG, $Result_Check_Table_Exists.TABLE_SCHEMA, $Result_Check_Table_Exists.TABLE_NAME
        
        $Query_Select_all = "SELECT * FROM {0}" -F $Table_Name
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("table '{0}' exists, {1} rows in it" -F $Table_Name, (Global:sql_query -query $Query_Select_all).count) 
        
        Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text "Verifiying columns..."
        $Query_List_columns = "SELECT col.name AS column_name,t.name AS data_type,col.max_length AS data_type_detail FROM sys.tables AS tab INNER JOIN sys.columns AS col ON tab.object_id = col.object_id LEFT JOIN sys.types AS t ON col.user_type_id = t.user_type_id WHERE tab.name = '{0}'" -F $Table_Name
        $SqlCurrentColumns = Global:sql_query -query $Query_List_columns

        $Queries_Update_table = ""
        $config.Fields | ForEach-Object {
            $this = $_ | Select-Object *, type_sql, type_detail

            #write-host ( "split of {0}, count:{1}" -F $this.type, ($this.type).split("(").count )
            if ( ($this.type).split("(").count -eq 1) {
                # no detail type such as DATETIME or INT
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = $null
            }
            else {
                $this.type_sql = ($this.type).split("(")[0]
                $this.type_detail = ($this.type).split("(")[1] -replace ".$"
            }
            #write-host ( "json this:{0}" -F ( $this | ConvertTo-Json -Compress ))
            
            $MatchingColumn = "" | Select-Object found, same_type, type_details_match
            $MatchingColumn.found = 0
            $MatchingColumn.same_type = 0
            $MatchingColumn.type_details_match = 0
            $SqlCurrentColumns | Where-Object { $_.column_name -eq $this.name } | ForEach-Object {
                $MatchingColumn.found = 1
                if ( $_.data_type -eq $this.type_sql ) { 
                    $MatchingColumn.same_type = 1
                    if ($_.data_type_detail -eq -1) { $temp_data_type_detail = [string]"MAX" } ELSE { $temp_data_type_detail = [string]$_.data_type_detail }
                    if ( $this.type_detail -ne $null) {
                        if ( $temp_data_type_detail -eq $this.type_detail) {
                            $MatchingColumn.type_details_match = 1
                        }
                        else {
                            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type detail for column ({3}) {2} :({0} <> {1})." -F $this.type_detail, $temp_data_type_detail, $this.name, $this.type) -type warning
                        }
                    }
                    else {
                        $MatchingColumn.type_details_match = 1
                    }

                }
                else {
                    Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Missmatching type for column {2} ({0} <> {1})." -F $this.type_sql, $_.data_type, $this.name) -type warning
                }
                #write-host ( "json MatchingColumn:{0}" -F ( $MatchingColumn | ConvertTo-Json -Compress ))

            }
        

            if ( $MatchingColumn.found -eq 0 ) {  
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "column {0} missing" -F $this.name) -type warning

            }

            $Query_Run = $null
            $Query_Alter_Table_Add_Column = "ALTER TABLE {0} ADD {1} {2};"
            $Query_Alter_Table_Alter_Column = "ALTER TABLE {0} ALTER COLUMN {1} {2};"

            if ( $MatchingColumn.found -eq 0 ) { 
                $Query_Run = $Query_Alter_Table_Add_Column 
            }
            else {
                if ( $MatchingColumn.same_type -eq 0 -or $MatchingColumn.type_details_match -eq 0 ) {
                    $Query_Run = $Query_Alter_Table_Alter_Column 
                }
            }

            if ( $Query_Run -ne $null ) {
                $Query_Run = $Query_Run -F $Table_schema_name, $this.name, $this.type
                $Queries_Update_table += $Query_Run + "`n"
            }

        }

        if ( $Queries_Update_table.Length -ge 1 ) {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "resulting alter table queries:{0}" -F $Queries_Update_table)
            IF ( $Global:WhatIf -ne $true ) {
                Global:sql_query -query $Queries_Update_table 
            }
            ELSE {
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("{0}" -F $Queries_Update_table ) -type warning
                Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ("WHATIF=`$true") -type error
            }

        }
        else {
            Global:Log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand ) -text ( "Table {0} is matching the definition of the config.json at run time." -F $Table_Name)
        }

        
    }
}

function Global:ADTidy_Records_sql_update {
    param(
        [Parameter(Mandatory = $true)] [array]$Fields,
        $Table_Name = "ADTidy_Records"
    )

    #region ADTidy_Inventory_Users_sql_update specific
    $prefixed_fields = "" | Select-Object ignore
    $varchar_field = @()
    $ObjectGUID = $Fields.ObjectGUID
    $Fields | Select-Object -ExcludeProperty ObjectGUID | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object -ExpandProperty name | ForEach-Object {
        $this_attribute_name = $_
        <# switch ($this_attribute_name) {
            "record_status" { $prefixed_attribute_name = $this_attribute_name }
            "xml_members" { $prefixed_attribute_name = $this_attribute_name }
            default { $prefixed_attribute_name = "ad_{0}" -F $this_attribute_name }

        }#>
        $varchar_field += $this_attribute_name

    }
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "prefixed_fields:{0}" -F $prefixed_fields | ConvertTo-Json -Compress    )

    #endregion#

    #region function internal definitions
    $table = $Table_Name
    #$sql_varchar_fields = @("entry_type", "entry_details", "user_type", "user_employeeid", "user_samaccountname", "user_guid", "execution_status_details", "execution_mode", "execution_operator_name", "execution_operator_action_timestamp", "execution_logs", "execution_last_update", "record_status")
    #$sql_special_fields = @("entry_repeat", "entry_first_occurrence", "entry_last_occurrence", "execution_status_code")
    #endregion

    #region create INSERT statement field and value pairs, loop throuhg $Fields array matching sql_field definition
    $sql_statement_fields = ""
    $sql_statement_values = ""
    $sql_update_statement = "$sql_update_statement"
    $varchar_field | ForEach-Object {
        $thisField = $_
        if ( $Fields."$thisField" -ne $null ) {
            #Global:Log -text (" + `$Fields contains attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
            $sql_statement_fields = "{0} [{1}]," -F $sql_statement_fields, $thisField
            if ($Fields."$thisField" -eq "NULL") {
                $sql_statement_values = "{0} {1}," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]={2}," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            else {
                $sql_statement_values = "{0} '{1}'," -F $sql_statement_values, $Fields."$thisField"
                $sql_update_statement = "{0} [{1}]='{2}'," -F $sql_update_statement, $thisField, $Fields."$thisField"
            }
            
            
        }
        else {
            #Global:Log -text (" ! `$Fields misses attribute '{0}'" -F $thisField) -hierarchy "function:adimport_sql_update:DEBUG"
        }
    }
    #endregion


    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "action_type = new"  )
    $sql_statement_fields = "{0} [record_timestamp]," -F $sql_statement_fields
    $sql_statement_values = "{0} GETDATE()," -F $sql_statement_values
        
    # remove last ',' from both fields and values strings
    $sql_statement_fields = $sql_statement_fields -replace ".$"
    $sql_statement_values = $sql_statement_values -replace ".$"

    $sql_statement_insert = " INSERT INTO {0} ({1}) VALUES ({2})" -F $table, $sql_statement_fields, $sql_statement_values
        
    Global:log -Hierarchy ("function:{0}" -F $MyInvocation.MyCommand )  -text ( "query:'{0}'" -F $sql_statement_insert  )
    $insert_result = Global:sql_query -query $sql_statement_insert


    Global:Log -text ("Inserted row id: '{0}' " -F ($insert_result | convertto-json -Compress) ) -hierarchy "function:adimport_sql_update:DEBUG"
    return $insert_result.rowid
 


}
										

