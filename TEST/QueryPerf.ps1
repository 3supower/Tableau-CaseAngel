$query = "SELECT 
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c, 
    Account.Name,
    Account.CSM_Name__c,
    Account.CSM_Email__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
        WHERE CreatedDate=TODAY and field='Owner'),
    (SELECT CreatedById, body FROM Feeds),
    (SELECT MilestoneTypeId,TargetDate,TimeRemainingInDays,TimeRemainingInHrs,TimeRemainingInMins,IsViolated FROM CaseMilestones)
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "

$query2 = "SELECT 
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c, 
    Account.Name,
    Account.CSM_Name__c,
    Account.CSM_Email__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
        WHERE CreatedDate=TODAY and field='Owner'),
    (SELECT MilestoneTypeId,TargetDate,TimeRemainingInDays,TimeRemainingInHrs,TimeRemainingInMins,IsViolated FROM CaseMilestones)
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "

$query3 = "SELECT 
    RecordTypeId,
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c, 
    Account.Name,
    Account.CSM_Name__c,
    Account.CSM_Email__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
        WHERE CreatedDate=TODAY and field='Owner')
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "

$query4 = "SELECT 
    RecordTypeId,
	Id, 
	CaseNumber,
	Priority, 
	Case_Age__c, 
    Status,
    Description,
    Preferred_Case_Language__c,
    Case_Preferred_Timezone__c,
	Tier__c,
	Entitlement_Type__c,
	Category__c, 
	Product__c, 
	Subject, 
	First_Response_Complete__c, 
	CreatedDate,
	Plan_of_Action_Status__c, 
	Case_Owner_Name__c,
    AccountId,
    IsEscalated, Escalated_Case__c,
    ClosedDate, IsClosed, isClosedText__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById 
        FROM Histories 
        WHERE CreatedDate=TODAY and field='Owner')
FROM Case 
WHERE
	RecordTypeId='012600000000nrwAAA' AND 
    ( (IsClosed=False) OR (IsClosed=True AND ClosedDate=TODAY) ) AND
	Preferred_Support_Region__c ='APAC' AND 
	Preferred_Case_Language__c != 'Japanese' AND 
    Tier__c != 'Admin'
" -replace "`n", " "


function get-all {
    Param($Query)
    Do {
        $ts_start = (Get-Date)
        $ts = $ts_start.ToString("yyyy-MM-dd-HH:mm:ss")
        Write-Host "Query starts at $ts ==" -ForegroundColor Yellow
        
        $json_result = (sfdx force:data:soql:query -q $Query -u vscodeOrg --json)
        
    } While (($null -eq $json_result) -or ($json_result -eq $false))

    $ts_end = (Get-Date)
    $ts = $ts_end.ToString("yyyy-MM-dd-HH:mm:ss")
    Write-Host "Query Finished at $ts" -ForegroundColor Yellow

    $els = $ts_end - $ts_start
    Write-Host "Elapsed Time: $els"
    # $totalRecordNumber = $json_result.

    $raw_obj = ($json_result | ConvertFrom-Json).result.records
    $new_obj = @()

    $raw_obj | ForEach-Object {

        # Add CSM Property in the PSCustomObject = results set
        $_ | Add-Member -MemberType NoteProperty -Name "CSM" -Value $null

        # Handoff
        <#
        if ( ($_.Plan_of_Action_Status__c -ne $null) -and ($_.Plan_of_Action_Status__c -like "*apac*") ) {
            Write-Host $_.CaseNumber -ForegroundColor Magenta -NoNewline
            Write-Host " <== HandOff: " -ForegroundColor Magenta -NoNewline
            Write-Host $_.Plan_of_Action_Status__c -ForegroundColor Magenta
        # Protips
        } elseif ( ($_.Plan_of_Action_Status__c -ne $null) -and ($_.Plan_of_Action_Status__c -like "*protip*") ) {
            Write-Host $_.CaseNumber -ForegroundColor DarkBlue -NoNewline
            Write-Host " <-- " -ForegroundColor DarkBlue -NoNewline
            Write-Host $_.Plan_of_Action_Status__c -ForegroundColor DarkBlue
            # $_.feeds = 'YES'
        }
        #>

        if ($null -ne $_.Account.CSM_Name__c) {
            # Write-Host $_.Account.AnnualRevenue
            # Write-Host $_.Account.frmtAmnt
            # Write-Host $_.Account.cnvAmnt
            # $ARR = ToKMB($_.Account.AnnualRevenue)
            # $_.Account.CSM_Name__c = "YES"
            $_.CSM = "YES"
        }
        
        # Changed owner Today
        Write-Host "My History"
        Write-Host $_.Histories.records
        if ($_.Histories.records -ne $null) {
            Write-Host "I am done today"
            $_.Histories = "YES"
        }

        # Protips
        
        if ($_.feeds.records.body -like "*protip*") {
            # Write-Host $_.CaseNumber -ForegroundColor Yellow -NoNewline
            # Write-Host " <-- protip" -ForegroundColor Yellow
            # Write-Host $_.feeds.records.body -ForegroundColor Yellow
            $_.feeds.records = "YES"
        } else {
            # Write-Host $_.CaseNumber -ForegroundColor DarkCyan
            # Write-Host $_.feeds.records.body -ForegroundColor Cyan
            $_.feeds.records = "" #$null
        }
        

        if ($_.Entitlement_Type__c -eq 'TOL Premium') {
            $_.Entitlement_Type__c = 'Premium'
        }

        $new_obj += $_
    }

    #Filter the new object to keep the cases unassigned or assigned to somebody TODAY only.
    $new_obj = ($new_obj | Where-Object { ($_.Case_Owner_Name__c -eq $null) -or (  ($_.Case_Owner_Name__c -ne $null) -and ($_.Histories -eq "YES") ) })

    return $new_obj
}

function get-small {
    Param($Query)
    Do {
        $ts_start = (Get-Date)
        $ts = $ts_start.ToString("yyyy-MM-dd-HH:mm:ss")
        Write-Host "Query starts at $ts ==" -ForegroundColor Yellow
        
        $json_result = (sfdx force:data:soql:query -q $Query -u vscodeOrg --json)
        
    } While (($null -eq $json_result) -or ($json_result -eq $false))

    $ts_end = (Get-Date)
    $ts = $ts_end.ToString("yyyy-MM-dd-HH:mm:ss")
    Write-Host "Query Finished at $ts" -ForegroundColor Yellow

    $els = $ts_end - $ts_start
    Write-Host "Elapsed Time: $els"

    $raw_obj = ($json_result | ConvertFrom-Json).result.records
    $new_obj = @()
    # $raw_obj | gm
    $raw_obj | ForEach-Object {

        # Add CSM Property in the PSCustomObject = results set
        $_ | Add-Member -MemberType NoteProperty -Name "CSM" -Value $null

        if ($null -ne $_.Account.CSM_Name__c) {
            $_.CSM = "YES"
        }

        
        if (($_.isClosed -eq $true) -and ( ($_.Case_Owner_Name__c -eq $null) -or ($_.Case_Owner_Name__c -eq '') ) ) {
            Write-Host "What is the case ownername?"
            # $_.Case_Owner_Name__c = "By Customer"
            Write-Host $_.Case_Owner_Name__c
        }
        
        # Changed owner Today
        Write-Host "My history:"
        Write-Host $_.Histories.records
        if ($null -ne $_.Histories.records) {
            $_.Histories = "YES"
            Write-Host "Yes, I am done today!"
        }

        if ($_.Entitlement_Type__c -eq 'TOL Premium') {
            $_.Entitlement_Type__c = 'Premium'
        }

        $new_obj += $_
    }

    #Filter the new object to keep the cases unassigned or assigned to somebody TODAY only.
    $new_obj = ($new_obj | Where-Object { ($_.Case_Owner_Name__c -eq $null) -or (  ($_.Case_Owner_Name__c -ne $null) -and ($_.Histories -eq "YES") ) })

    return $new_obj
}


function get-tiny {
    Param($Query)
    Do {
        $ts_start = (Get-Date)
        $ts = $ts_start.ToString("yyyy-MM-dd-HH:mm:ss")
        Write-Host "Query starts at $ts ==" -ForegroundColor Yellow
        
        $json_result = (sfdx force:data:soql:query -q $Query -u vscodeOrg --json)
        
    } While (($null -eq $json_result) -or ($json_result -eq $false))

    $ts_end = (Get-Date)
    $ts = $ts_end.ToString("yyyy-MM-dd-HH:mm:ss")
    Write-Host "Query Finished at $ts" -ForegroundColor Yellow

    $els = $ts_end - $ts_start
    Write-Host "Elapsed Time: $els"

    $raw_obj = ($json_result | ConvertFrom-Json).result.records
    $new_obj = @()
    # $raw_obj | gm
    $raw_obj | ForEach-Object {
        if (($_.isClosed -eq $true) -and ( ($_.Case_Owner_Name__c -eq $null) -or ($_.Case_Owner_Name__c -eq '') ) ) {
            Write-Host "What is the case ownername?"
            # $_.Case_Owner_Name__c = "By Customer"
            Write-Host $_.Case_Owner_Name__c
        }
        
        # Changed owner Today
        # Write-Host "My history:"
        # Write-Host $_.Histories.records
        if ($null -ne $_.Histories.records) {
            $_.Histories = "YES"
            Write-Host "Yes, I am done today!"
        }

        if ($_.Entitlement_Type__c -eq 'TOL Premium') {
            $_.Entitlement_Type__c = 'Premium'
        }

        $new_obj += $_
    }

    #Filter the new object to keep the cases unassigned or assigned to somebody TODAY only.
    $new_obj = ($new_obj | Where-Object { ($_.Case_Owner_Name__c -eq $null) -or (  ($_.Case_Owner_Name__c -ne $null) -and ($_.Histories -eq "YES") ) })

    return $new_obj
}


# get-all($query)
# get-small($query3)

$all = get-tiny($query4)

$all | %{
    if ($_.CaseNumber -eq "07414350") {
        $_.CaseNumber
        $_.CreatedDate
        [datetime]$_.CreatedDate
        ([datetime]$_.CreatedDate).ToUniversalTime()
        [datetime]"8/7/2021 23:59"
        if (([datetime]$_.CreatedDate).ToUniversalTime() -gt [datetime]"8/7/2021 23:59") {
            Write-Host "Yes"
        } else {
            Write-Host "No"
        }
    }
    
    Write-Host ""
}

# $all = $all | ?{ (([datetime]$_.CreatedDate) -gt [datetime]"8/7/2021") }

# $all