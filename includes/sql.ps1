#region : SQL Utility DB connection init.

function Global:sql_query {
    param(
        $SqlInstance = $global:Config."SQL connections".ITtool.Instance,
        $Database = $global:Config."SQL connections".ITtool.Database,
        [Parameter(Mandatory = $true)] $query
    )

    $Global:adtools_sql_query_last_error = "Ok"

    Global:log -Hierarchy "function:sql_query" -text ("SqlInstance={0},Database={1},TrustServerCertificate=false,Query={2}" -F $SqlInstance, $Database, $query )
    
    # https://github.com/dataplat/dbatools/discussions/7680
    Set-DbatoolsConfig -FullName 'sql.connection.trustcert' -Value $true -Register

    try {
        $temp = Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $query -MessagesToOutput -EnableException 
    } 
    
    catch {
        $Global:adtools_sql_query_last_error = $_
        Global:log -Hierarchy "function:sql_query" -text ("error details:{0}" -F $Global:adtools_sql_query_last_error) -type error
    }
    Global:log -Hierarchy "function:sql_query" -text ("returned rows={0}" -F ($temp | Measure-Object).count )
    return $temp

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

