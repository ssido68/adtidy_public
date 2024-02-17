function Global:log {
    param (
        $file = ("{0}\{1}.log.txt" -F $Global:LogLocation, ($MyInvocation.ScriptName).split("\")[($MyInvocation.ScriptName).split("\").count - 1]) ,
        [Parameter(Mandatory = $true)] $text, 
        $ToConsole = $true,
        $ToFile = $true,
        [ValidateSet("info", "warning", "error")] [String[]] $type = "info",
        $Hierarchy = $false,
        $force = $false
    ) 
    $scriptName = ($MyInvocation.ScriptName).split("\")[($MyInvocation.ScriptName).split("\").count - 1]
    switch ( $type ) {
        "info" {
            $color = "darkgreen"
        }
        "warning" {
            $color = "darkyellow" 
        }
        "error" {
            $color = "darkred" 
        }
    }

    # Log to interface
    if ( $Global:Debug -or $force -eq $true) {
        if ($ToFile) {
            if (!$Hierarchy ) { $logTemplate = "{0}:{1}# {2}" } else { $logTemplate = "{0}:{1}#{3}# {2}" }
            $logTemplate -F (Get-Date -Format "yyyyMMdd-HH:mm:ss"), $scriptName, $text, $Hierarchy | Out-File $file -Encoding utf8 -Append
        }

        if ($ToConsole) {
            Write-Host -ForegroundColor DarkGray (Get-Date -Format "yyyyMMdd-HH:mm:ss") -NoNewline
            Write-Host -ForegroundColor White ":" -NoNewline
            Write-Host -ForegroundColor DarkCyan $scriptName -NoNewline
            if ($Hierarchy ) { Write-Host -ForegroundColor DarkMagenta ("#{0}" -F $Hierarchy) -NoNewline }
            Write-Host -ForegroundColor White "# " -NoNewline
            Write-Host -ForegroundColor $color $text
        }

        if ( $MainForm.ishandlecreated ) {
            $RichTextBoxLogs.selectioncolor = [Drawing.Color]::DarkGray 
            $RichTextBoxLogs.AppendText((Get-Date -Format "yyyyMMdd-HH:mm:ss"))
            $RichTextBoxLogs.selectioncolor = [Drawing.Color]::White
            $RichTextBoxLogs.AppendText(":")
            $RichTextBoxLogs.selectioncolor = [Drawing.Color]::Magenta
            $RichTextBoxLogs.AppendText($Hierarchy)
            $RichTextBoxLogs.selectioncolor = [Drawing.Color]::White
            $RichTextBoxLogs.AppendText(":")
            switch ( $type ) {
                "info" {
                    $RichTextBoxLogs.selectioncolor = [Drawing.Color]::DarkGreen
                }
                "warning" {
                    $RichTextBoxLogs.selectioncolor = [Drawing.Color]::DarkOrange
                }
                "error" {
                    $RichTextBoxLogs.selectioncolor = [Drawing.Color]::DarkRed
                }
            }
            $RichTextBoxLogs.AppendText($text)
            $RichTextBoxLogs.AppendText("`n")
            $RichTextBoxLogs.ScrollToCaret()
        } 
    }
} 
function Global:ConvertTo-DataTable {
    <# https://fitch.tech/2014/08/09/convertto-datatable/
 .EXAMPLE
 $DataTable = ConvertTo-DataTable $Source
 .PARAMETER Source
 An array that needs converted to a DataTable object
 #>
    [CmdLetBinding(DefaultParameterSetName = "None")]
    param(
        [Parameter(Position = 0, Mandatory = $true)][System.Array]$Source,
        [Parameter(Position = 1, ParameterSetName = 'Like')][String]$Match = ".+",
        [Parameter(Position = 2, ParameterSetName = 'NotLike')][String]$NotMatch = ".+"
    )
    if ($NotMatch -eq ".+") {
        $Columns = $Source[0] | Select-Object * | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match "($Match)" }
    }
    else {
        $Columns = $Source[0] | Select-Object * | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -notmatch "($NotMatch)" }
    }
    $DataTable = New-Object System.Data.DataTable
    foreach ($Column in $Columns.Name) {
        $DataTable.Columns.Add("$($Column)") | Out-Null
    }
    #For each row (entry) in source, build row and add to DataTable.
    foreach ($Entry in $Source) {
        $Row = $DataTable.NewRow()
        foreach ($Column in $Columns.Name) {
            $Row["$($Column)"] = if ($Entry.$Column -ne $null) { ($Entry | Select-Object -ExpandProperty $Column) -join ', ' }else { $null }
        }
        $DataTable.Rows.Add($Row)
    }
    #Validate source column and row count to DataTable
    if ($Columns.Count -ne $DataTable.Columns.Count) {
        throw "Conversion failed: Number of columns in source does not match data table number of columns"
    }
    else { 
        if ($Source.Count -ne $DataTable.Rows.Count) {
            throw "Conversion failed: Source row count not equal to data table row count"
        }
        #The use of "Return ," ensures the output from function is of the same data type; otherwise it's returned as an array.
        else {
            Return , $DataTable
        }
    }
}

function Global:Compare-mObject {
    <#
    .SYNOPSIS
    Compare differences between two sets of data (Reference vs Difference) based on specified Property
    .DESCRIPTION
    - Indicators within the "Exists" property
    - Data that ONLY exists in Reference: "<="
    - Data that ONLY exists in Difference: "=>"
    - Data that exists in BOTH Reference and Difference: "=="
    - Some data may contain non-ASCII characters, such as Umlauts (https://stackoverflow.com/questions/48947151/import-csv-export-csv-with-german-umlauts-ä-ö-ü)
    Use "-Encoding UTF8" with Import-CSV and Export-CSV to handle UTF-8 (non-ASCII) characters
    .NOTES
    Script Created: 8/14/2019 Michael Yuen (www.yuenx.com)
    Change History
    - 8/20/2020 Michael Yuen: Turned original function into a standalone cmdlet
    .EXAMPLE
    Compare-mObjects -Reference "$REFERENCE_DATA" -Difference "$DIFFERENCE_DATA" -Property "samAccountName"
    Compare object differences between Reference and Difference using the "samAccountName" property as reference
    .EXAMPLE
    $Ref = Get-ADGroupMember "Group1"; $Diff = Get-ADGroupMember "Group2"; Compare-mObjects -Reference $Ref -Difference $Diff -Property "samAccountName"
    Compare AD group membership differences between "Group1" and "Group2" using the "samAccountName" property as reference
    #>
    [CmdletBinding()]
    param (
        # REFERENCE data
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $false)]$Reference,
        # DIFFERENCE data
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline = $false)]$Difference,
        # What Property to compare on (ie. distinguishedName)
        [Parameter(Position = 2, Mandatory = $true, ValueFromPipeline = $false)]$Property
    )
    <# Note: Under SideIndicator column: == item exists in both ReferenceObject and DifferenceObject
    <= item exists only in ReferenceObjet
    => item exists only in DifferenceObject
    Compare-Object: use -IncludeEqual parameter to include any values that exist in both files
    use -ExcludeDifferent parameter to exclude any values that don't exist in both files
    -SyncWindow specifies how far around to look to find the same element.
    Default: 5 (looks for +/- 5 elements around), which is good for up to 11 elements
    Using -SyncWindow 100 would be good for up to 201 elements
    #>
    $Result = Compare-Object -ReferenceObject $Reference -DifferenceObject $Difference -SyncWindow 5000 -IncludeEqual `
        -Property $Property -PassThru | Sort-Object SideIndicator
    # Modify SideIndicator to be readable and include all supplied properties
    $ResultExpanded = $Result | Select-Object `
    @{n = "Exists"; e = { If ($_.SideIndicator -like "<=") { Write-Output "<= In REFERENCE only" } `
                elseif ($_.SideIndicator -like "=>") { Write-Output "=> In DIFFERENCE only" } `
                elseif ($_.SideIndicator -like "==") { Write-Output "== In BOTH (Reference & Difference)" }
            else { Write-Output "N/A" } }
    }, * -ExcludeProperty PropertyNames, AddedProperties, RemovedProperties, ModifiedProperties, PropertyCount
    Return ( $ResultExpanded ) #| ? { $_.SideIndicator -ne "=="} )
}

Function Global:Get-ListValues {
    <# https://www.dowst.dev/format-data-returned-from-get-pnplistitem/
        .SYNOPSIS
        Use to create a PowerShell Object with only the columns you want,
        based on the data returned from the Get-PnPListItem command.
        
        .DESCRIPTION
        Creates a custom PowerShell object you can use in your script.
        It only creates properties for custom properties on the list and
        a few common ones. Filters out a lot of junk you don't need.
        
        .PARAMETER ListItems
        The value returns from a Get-PnPListItem command
        
        .PARAMETER List
        The name of the list in SharePoint. Should be the same value
        passed to the -List parameter on the Get-PnPListItem command
        
        .EXAMPLE
        $listItems = Get-PnPListItem -List $List 
        $ListValues = Get-ListValues -listItems $listItems -List $List
        
        
#>
    param(
        [Parameter(Mandatory = $true)]$ListItems,
        [Parameter(Mandatory = $true)]$List
    )
    # begin by gettings the fields that where created for this list and a few other standard field
    begin {
        $standardFields = 'Title', 'Modified', 'Created', 'Author', 'Editor'
        # get the list from SharePoint
        $listObject = Get-PnPList -Identity $List
        # Get the fields for the list
        $fields = Get-PnPField -List $listObject
        # create variable with only the fields we want to return
        $StandardFields = $fields | Where-Object { $_.FromBaseType -ne $true -or $standardFields -contains $_.InternalName } | 
        Select-Object @{l = 'Title'; e = { $_.Title.Replace(' ', '') } }, InternalName
    }
            
    process {
        # process through each item returned and create a PS object based on the fields we want
        [System.Collections.Generic.List[PSObject]] $ListValues = @()
        foreach ($item in $listItems) {
            # add field with the SharePoint object incase you need to use it in a Set-PnPListItem or Remove-PnPListItem
            $properties = @{SPObject = $item }
            foreach ($field in $StandardFields) {
                $properties.Add($field.Title, $item[$field.InternalName])
            }
            $ListValues.Add([pscustomobject]$properties)
        }
    }
            
    end {
        # return our new object
        $ListValues
    }
}

function Global:ExpandSharepointLookupValue ($Value) {      
    if ($Value -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
        $Value[0].LookupValue
    }
    else {
        $Value
    }
}

function Global:Send-Mail {
    # SAMPLE :
    # Send-Mail  -recipient "rhambalek@primeo-energie.ch"  -MSG_CONTENT "test msg" -MSG_TYPE "test message" -BCC_ADMINS 1
	
	
    #region PARAM
    param (
		
        [parameter(Mandatory = $true, HelpMessage = "email of the recipient")]
        [array]$recipient,
		
        [parameter(Mandatory = $false, HelpMessage = "if message should be BCCed admin set=1")]
        [string]$BCC_ADMINS = 0,
		
        [parameter(Mandatory = $true, HelpMessage = "Email content in HTML")]
        [string]$MSG_CONTENT,

        [parameter(Mandatory = $true, HelpMessage = "Email title, header is fixed by called function")]
        [string]$MSG_TITLE
		
        
    )
	
    #endregion
	
    #region BEGIN
    BEGIN {
        IF ($Global:Debug -EQ $true ) { $BCC_ADMINS = 1 }
        
        $EMAIL_SMTP_GATEWAY = $global:Config.configurations.mail.gateway

        IF ($BCC_ADMINS -EQ 1 ) {
            $ArrayOfStrings = @($global:Config.configurations.mail.admins, $global:Config.configurations.mail."collaboration team channel mail")
            $EMAIL_BCC_ADMINS = $ArrayOfStrings -join ";"
        }
    
        $SMTP = $EMAIL_SMTP_GATEWAY
        $FROM = "Primeo.Powershell.MsTeams." + $env:computername + "@primeo-energie.ch"
        $SUBJECT_ROOT = "Microsoft Teams Archiving - "
        $ADMINS = $EMAIL_BCC_ADMINS
		
       
		
			
    }

    #region PROCESS
    PROCESS {
        $msg = New-Object Net.Mail.MailMessage
        $smtp = New-Object Net.Mail.SmtpClient($SMTP)
        $msg.From = $FROM
        $msg.Subject = "$SUBJECT_ROOT $MSG_TITLE"

  
        IF ( $recipient.split(";").count -NE 1 ) {

            $recipient.split(";") | ForEach-Object {
                IF ( ( $_).length -NE 0 ) {
                    $msg.TO.Add( $_)
                    Global:Log -Hierarchy "Function:Send-Mail" -text ("Adding recipient:{0}" -F $_ )
                }
                
            }
        }
        ELSE {
            $msg.TO.Add( $recipient)
            Global:Log -Hierarchy "Function:Send-Mail" -text ("Adding recipient:{0}" -F $recipient )
        }
       
        if ($BCC_ADMINS -eq 1) {
            #$ADMINS.split(";") | ForEach-Object { $msg.BCC.Add($_) }
            Global:Log -Hierarchy "Function:Send-Mail" -text ("EMAIL_BCC_ADMINS:{0}" -F $ADMINS )
        }
        
        #
        $msg.isbodyhtml = $True
        $msg.Body = $MSG_CONTENT
        Global:Log -Hierarchy "Function:Send-Mail" -text ("MSG_CONTENT (length):{0}" -F $MSG_CONTENT.Length )

        $Success = 1
        TRY { $smtp.Send($msg) } CATCH {
            $Success = 0 
            Global:Log -Hierarchy "Function:Send-Mail" -text ("Sending failed" ) -type error
        }
		
        IF ($Success -EQ 1 ) {
            Global:Log -Hierarchy "Function:Send-Mail" -text ("Sent successfully" )
        }
        return $Success
		
    }
    #endregion
	
    #region END
    END {
	
    }
    #endregion
}
Function Global:DecodeUserAccountControl ([int]$UAC) {
    $UACPropertyFlags = @(
        "SCRIPT",
        "ACCOUNT_DISABLED",
        "RESERVED",
        "HOMEDIR_REQUIRED",
        "LOCKOUT",
        "PASSWD_NOT_REQUIRED",
        "PASSWD_CANNOT_BE_CHANGE",
        "ENCRYPTED_TEXT_PWD_ALLOWED",
        "TEMP_DUPLICATE_ACCOUNT",
        "NORMAL_ACCOUNT",
        "RESERVED",
        "INTERDOMAIN_TRUST_ACCOUNT",
        "WORKSTATION_TRUST_ACCOUNT",
        "SERVER_TRUST_ACCOUNT",
        "RESERVED",
        "RESERVED",
        "PASSWORD_NEVER_EXPIRES",
        "MNS_LOGON_ACCOUNT",
        "SMARTCARD_REQUIRED",
        "TRUSTED_FOR_DELEGATION",
        "NOT_DELEGATED",
        "USE_DES_KEY_ONLY",
        "DONT_REQ_PREAUTH",
        "PASSWORD_EXPIRED",
        "TRUSTED_TO_AUTH_FOR_DELEGATION",
        "RESERVED",
        "PARTIAL_SECRETS_ACCOUNT"
        "RESERVED"
        "RESERVED"
        "RESERVED"
        "RESERVED"
        "RESERVED"
    )
    $Attributes = ""
    1..($UACPropertyFlags.Length) | Where-Object { $UAC -bAnd [math]::Pow(2, $_) } | ForEach-Object { If ($Attributes.Length -EQ 0) { $Attributes = $UACPropertyFlags[$_] } Else { $Attributes = $Attributes + " | " + $UACPropertyFlags[$_] } }
    Return $Attributes
}
function Global:Get-FileEncoding {
    [CmdletBinding()]
    param (
        [Alias("PSPath")]
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [String]$Path
        ,
        [Parameter(Mandatory = $False)]
        [System.Text.Encoding]$DefaultEncoding = [System.Text.Encoding]::ASCII
    )
    
    process {
        [Byte[]]$bom = Get-Content -Encoding Byte -ReadCount 4 -TotalCount 4 -Path $Path
        
        $encoding_found = $false
        
        foreach ($encoding in [System.Text.Encoding]::GetEncodings().GetEncoding()) {
            $preamble = $encoding.GetPreamble()
            if ($preamble) {
                foreach ($i in 0..$preamble.Length) {
                    if ($preamble[$i] -ne $bom[$i]) {
                        break
                    }
                    elseif ($i -eq $preable.Length) {
                        $encoding_found = $encoding
                    }
                }
            }
        }
        
        if (!$encoding_found) {
            $encoding_found = $DefaultEncoding
        }
    
        $encoding_found
    }
}

function Global:API-ADimport_compare_Submit {
    param (
        [Alias("Data Set")]
        [Parameter(Mandatory = $True)]
        [array]$Data
        , [Alias("API call URI")]
        [Parameter(Mandatory = $False)]
        $API_URI = $global:Config.API.URI_API_Compare    
        , [Alias("API Auth Key")]
        [Parameter(Mandatory = $False)]
        $API_AuthKey = $global:Config.API.AuthKey_API_Compare  
    )

    $ApiCall_ScriptBlock = { 
        Param($URI, $Data, $AuthKey)
        
        $Array = $Data | ConvertFrom-Json 
        $JsonData = $Array | ConvertTo-Json
        $ApiCallSuccessful = 1
        $requestHeaders = @{"AuthKey" = $AuthKey }
    
        $returnArray = "" | Select-Object API_call, Status_message, target, submitted_json
        $returnArray.target = $Array.item.Target_object
        if ( $Array.item.Record_Id -eq $null ) {
            $URI = $URI + "new"
        }
        else {
            $URI = $URI + "update"  
        }
        
        $returnArray.API_call = "Successful"
        $returnArray.Status_message = ( "Target: {0} " -F $Array.item.Target_object )
        
        if (1) {
            TRY { $result = Invoke-RestMethod -Method Post -Uri $URI -Body $JsonData -ContentType "application/json; charset=utf-8" -Headers $requestHeaders }
            CATCH {
               
               
                $returnArray.submitted_json = $Array | ConvertTo-Json -Compress
                $ApiCallSuccessful = 0 
                $result = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($result)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = ( $reader.ReadToEnd() | ConvertFrom-Json )
                $arrayOfErrors = $responseBody.errors | ForEach-Object {
                    $_
                }
    
    
                $ErrorCode = $_.Exception.Response.StatusCode.value__ 
                $ErrorDescription = $arrayOfErrors | ConvertTo-Json -Compress
                $returnArray.API_call = "Failed"
                $returnArray.Status_message = "Error code:$ErrorCode : (B) $ErrorDescription"

            }        
        } 
    
        return $returnArray
    }



    $TotalJobStarted = 1
    $JobStarted = 0
    $JobsOutput = @()
    Global:log -text ("Running API calls ( parallel )" ) -Hierarchy "function:API-ADimport_compare_Submit" 

    $MaxParallelJobs = [int]$global:Config.Api."Parallel_Jobs"

    $Data | ForEach-Object {
        $DataItem = $_
        
        if ( $JobStarted -eq $MaxParallelJobs ) {
            $OutputArray = @()
            $JobStarted = 0
            While ( (Get-Job | Where-Object { $_.State -eq 'Running' -and $_.name -ne "dbatools_Timer" } | Measure-Object ).count -ne 0 ) {
                Get-Job | Where-Object { $_.State -eq 'Completed' } | ForEach-Object {
                    $returned = Receive-Job $_
                    $JobsOutput += $returned
                    $output = "Job with name '{0}' has returned the following: '{1}'" -F $_.Name, $returned
                    #Write-Host $output 
                    Remove-Job $_
                }
                Start-Sleep 1
                $msg = "{0} job(s) still running" -F (Get-Job | Where-Object { $_.State -eq 'Running' } | Measure-Object ).count
                #Write-Host $msg 
            }
        }

        Global:log -text ("starting job {0} / {1}" -F $TotalJobStarted, $Data.Count ) -Hierarchy "Main:API-ADimport_compare_Submit" 

        $ApiCallData = @{
            "item" = [Ordered] @{
                "Record_id"             = $DataItem.record_id
                "Record_type"           = $DataItem.record_type
                "Target_object"         = $DataItem.target_object
                "Action_type"           = $DataItem.action_type
                "Action_details"        = $DataItem.action_details
                "Action_remarks"        = $DataItem.action_remarks
                "Action_Status"         = $DataItem.action_status
                "Action_Result"         = $DataItem.action_result
                "Action_Result_Details" = $DataItem.action_result_details
                "Validation_Operator"   = $DataItem.validation_operator
            }
        }  
        
        $JsonApiCallData = $ApiCallData | ConvertTo-Json -Compress
        Write-Host $JsonApiCallData
        
        #Write-Host "start job info : -ArgumentList $API_URI, $JsonApiCallData, $API_AuthKey"

        Start-Job -ScriptBlock $ApiCall_ScriptBlock -ArgumentList $API_URI, $JsonApiCallData, $API_AuthKey | Out-Null

        $TotalJobStarted++
        $JobStarted++
    }

    Get-Job | Where-Object { $_.State -eq 'Completed' -and $_.name -ne "dbatools_Timer" } | ForEach-Object {
        $returned = Receive-Job $_
        $JobsOutput += $returned
        $output = "Job with name '{0}' has returned the following: '{1}'" -F $_.Name, $returned
        #Write-Host $output 
        Remove-Job $_
    }


    if (0) {
        $Data | ForEach-Object {
            $DataItem = $_
            
            if ( $JobStarted -eq $MaxParallelJobs ) {
                $OutputArray = @()
                $JobStarted = 0
                While ( (Get-Job | Where-Object { $_.State -eq 'Running' -and $_.name -ne "dbatools_Timer" } | Measure-Object ).count -ne 0 ) {
                    Get-Job | Where-Object { $_.State -eq 'Completed' -and $_.name -ne "dbatools_Timer" } | ForEach-Object {
                        $returned = Receive-Job $_
                        $JobsOutput += $returned
                        $output = "Job with name '{0}' has returned the following: '{1}'" -F $_.Name, $returned
                        #Write-Host $output 
                        Remove-Job $_
                    }
                    Start-Sleep 1
                    $msg = "{0} job(s) still running" -F (Get-Job | Where-Object { $_.State -eq 'Running' } | Measure-Object ).count
                    #Write-Host $msg 
                }
            }
    
            Global:log -text ("starting job {0} / {1}" -F $TotalJobStarted, $Data.Count ) -Hierarchy "Main:API-ADimport_compare_Submit" 
    
            $ApiCallData = @{
                "item" = [Ordered] @{
                    "Record_id"             = $DataItem.record_id
                    "Record_type"           = $DataItem.record_type
                    "Target_object"         = $DataItem.target_object
                    "Action_type"           = $DataItem.action_type
                    "Action_details"        = $DataItem.action_details
                    "Action_Remarks"        = $DataItem.action_remarks
                    "Action_Status"         = $DataItem.action_status
                    "Action_Result"         = $DataItem.action_result
                    "Action_Result_Details" = $DataItem.action_result_details
                    "Validation_Operator"   = $DataItem.validation_operator
                }
            }  
            
            $JsonApiCallData = $ApiCallData | ConvertTo-Json -Compress
    
            # Write-Host "start job info : -ArgumentList $API_URI, $JsonApiCallData, $API_AuthKey"
    
            Start-Job -ScriptBlock $ApiCall_ScriptBlock -ArgumentList $API_URI, $JsonApiCallData, $API_AuthKey 
    
            $TotalJobStarted++
            $JobStarted++
        }
    
        Get-Job | Where-Object { $_.State -eq 'Completed' -and $_.name -ne "dbatools_Timer" } | ForEach-Object {
            $returned = Receive-Job $_
            $OutputArray += $returned
            $output = "Job with name '{0}' has returned the following: '{1}'" -F $_.Name, $returned
            #Write-Host $output 
            Remove-Job $_
        }
    }

    
    return $JobsOutput



}


function Global:API_Call () {
    param (
    
        [string]$AuthKey = $null,
        [string]$ContentType = "application/json; charset=utf-8",
        [string]$Method = "Post",
        [Parameter(Mandatory = $true)] [string]$URI,
        [Parameter(Mandatory = $true)] [array]$Body
    )

    if ( $AuthKey -ne $null) {
        $Headers = @{"AuthKey" = $AuthKey }
        $LogAuthString = ", Authentication is on"
    }
    else {
        $Headers = $null
        $LogAuthString = $null
    }

    $BodyJson = [string]($Body | ConvertTo-Json -Compress)
    $ReturnedArray = "" | Select-Object ReturnCode, details, "data"
    $ReturnedArray.ReturnCode = $null
    $ReturnedArray.details = $null
    $ApiCallSuccessful = 1
    Global:log -text ("URI:{0}, Body:{1}{2}, ContentType:{3}" -F $URI, $BodyJson , $LogAuthString, $ContentType) -Hierarchy "function:Global:API_Call" 

    TRY { $data = Invoke-WebRequest -Method $Method -Uri $URI -Body $BodyJson -ContentType $ContentType -Headers $Headers }
    CATCH {

        $ApiCallSuccessful = 0 
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $ErrorDescription = $reader.DiscardBufferedData()
        $ErrorDetails = $reader.ReadToEnd() | ConvertFrom-Json
        #write-host ( "ErrorDetails:{0}"-F $ErrorDetails |ConvertTo-Json -Compress )
        $errorsHeaders = $ErrorDetails.errors | Get-Member | Where-Object { $_.membertype -eq "NoteProperty" } | Select-Object name -ExpandProperty name
        #write-host ( "errorsHeaders:{0}"-F $errorsHeaders |ConvertTo-Json -Compress )
        $Errors = @()
        $errorsHeaders | ForEach-Object {
            $Errors += [string]$ErrorDetails.errors."$_"
            
        
            $ErrorCode = $_.Exception.Response.StatusCode.value__ 
        
            $errorItem = "" | Select-Object title, details
            $errorItem.title = $ErrorDetails.title
            $ReturnedArray.ReturnCode = $ErrorDetails.status
            $errorItem.details = $Errors 
            $ReturnedArray.details = $errorItem | ConvertTo-Json -Compress
            Global:log -text (" Call failed: {0}" -F $ReturnedArray.details | ConvertTo-Json -Compress) -Hierarchy "function:Global:API_Call" -type error
        }
    } 

    if ( $ApiCallSuccessful -eq 0) {
        $ReturnedArray.ReturnCode = "Failed"
        Global:log -text (" Call Failed, timed out?") -Hierarchy "function:Global:API_Call"     
    }

    if ( $data.content -ne $null) {
        $ReturnedArray.ReturnCode = $data.StatusCode
        $ReturnedArray."data" = $data.content
        Global:log -text (" Call successfull, returned content is {1} object(s), total string length: {0}" -F ($data.content).length, (($data.content | ConvertFrom-Json) | Measure-Object).count) -Hierarchy "function:Global:API_Call" 
    }

    return $ReturnedArray
}