## Console output encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

## Variables
$user =           $env:USERNAME
$sp_path =        "C:\Users\$user\Tableau Software Inc\APAC Tech Support - Documents" # SharePoint Local folder path
$file_name =      "APACTechSupRealTimeDashboard.xlsx"
$sfile =          "$sp_path\$file_name" # SharePoint local Path 

## SharePoint (Online) ## 
$sp_root =   "https://tableau.sharepoint.com/sites/APACTechSupport/Shared%20Documents"
$sp_file =   "$sp_root/$file_name"
$web_lnk =   $sp_file + "?web=1"

## Slack ##
#ts-apac-support
# $uri =     "https://hooks.slack.com/services/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"
#ts-apac-sydney-py
$uri =       "https://hooks.slack.com/services/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"

### SalesForce Query ###
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
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById FROM Histories WHERE CreatedDate=TODAY and field='Owner'),
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
#
# (SELECT CreatedById, body FROM Feeds),
# Id IN (SELECT CaseID FROM CaseHistory ) AND Id IN (SELECT ParentId FROM CaseFeed) AND Id IN (SELECT CaseId FROM CaseMilestone)
# Account.AnnualRevenue,
# FORMAT(Account.AnnualRevenue) frmtAmnt,
# convertCurrency(Account.AnnualRevenue) cnvAmnt,
# (Status='New' or Status='Active' or Status='Re-opened') AND
# ORDER BY Case_Owner_Name__c, Tier__c, Priority ASC, Case_Age__c DESC" -replace "`n", " "
# ORDER BY Case_Owner_Name__c, Tier__c DESC, Entitlement_Type__c DESC, Priority,Case_Age__c DESC" -replace "`n", " "
# ORDER BY Case_Owner_Name__c, Tier__c DESC, Account.CSM_Name__c DESC, Priority,Case_Age__c DESC" -replace "`n", " "
# ORDER BY Case_Owner_Name__c, Priority,Case_Age__c DESC" -replace "`n", " "

# sfdx force:auth:web:login -a vscodeOrg
# sfdx force:org:list --all
# sfdx force:data:soql:query -u vscodeOrg -q "select id from user limit 5"

function get-all {
    Param($Query)
    Do {
        $ts = (Get-Date).ToString("yyyy-MM-dd-HH:mm:ss")
        Write-Host "Query starts at $ts ==" -ForegroundColor Yellow
        
        $json_result = (sfdx force:data:soql:query -q $Query -u vscodeOrg --json)
        
    } While (($null -eq $json_result) -or ($json_result -eq $false))

    Write-Host "Query Finished" -ForegroundColor Yellow
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
        if ($null -ne $_.Histories.records) {
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


function filter-oldp3p4 {
    param($In)
    return ( $In | Where-Object { !( !( ($_.Entitlement_Type__c -match "Premium") -or ($_.Entitlement_Type__c -match "Extended") -or ($_.Entitlement_Type__c -match "Elite") ) -and (($_.Priority -eq "P3") -or ($_.Priority -eq "P4")) -and ($_.Case_Age__c -lt 63) ) } )
}

function filter-desktop {
    param($In)
    return ( $In | Where-Object { (
        ($_.Product__c -eq "Tableau Desktop") -or 
        ($_.Product__c -eq "Tableau Public Desktop") -or 
        ($_.Product__c -eq "Tableau Reader") -or 
        ($_.Product__c -eq "Tableau Prep") -or 
        ($_.Product__c -eq "Tableau Public Desktop") ) # -and ($_.Case_Owner_Name__c -eq $null)
    })
}

function filter-server {
    param($In)
    return ( $In | Where-Object { (
        ($_.Product__c -eq "Tableau Server") -or 
        ($_.Product__c -eq "Tableau Public Server") -or 
        ($_.Product__c -eq "Tableau Online") -or 
        ($_.Product__c -eq "Tableau Mobile") -or 
        ($_.Product__c -eq "Tableau Resource Monitoring Tool") -or 
        ($_.Product__c -eq "Tableau Content Migration Tool") ) # -and ($_.Case_Owner_Name__c -eq $null)
    })
}

function filter-premium {
    param($In)
    return ( $In | Where-Object { (
        # ($_.Tier__c -eq "Premium") -or ($_.Entitlement_Type__c -match "Extended") -or ($_.Entitlement_Type__c -match "Elite")) -and
        ($_.Entitlement_Type__c -match "Premium") -or ($_.Entitlement_Type__c -match "Extended") -or ($_.Entitlement_Type__c -match "Elite")) -and
        ($_.Case_Owner_Name__c -eq $null)
    })
}

function Filter-P1P2 {
    param($InputList)
    return ( $InputList | Where-Object {
        ( ($_.Priority -eq "P1") -or ($_.Priority -eq "P2") ) -and
        ($_.Case_Owner_Name__c -eq $null)
    })
}

function Filter-P3P4 {
    param($InputList)
    return ( $InputList | ?{
        ( ($_.Priority -eq "P3") -or ($_.Priority -eq "P4") ) -and ($_.Case_Age__c -gt 99) -and
        ($_.Case_Owner_Name__c -eq $null)
    })
}

function filter-urgent { # Action Now!
    param($In)
    return ( $In | ?{ (
        ( ($_.Priority -eq "P1") -and ($_.Case_Age__c -gt 6) ) -or 
        ( ($_.Priority -eq "P2") -and ($_.Case_Age__c -gt 18) ) -or 
        ( ($_.Priority -eq "P3") -and ($_.Case_Age__c -gt 50) ) -or
        ( ($_.Priority -eq "P4") -and ($_.Case_Age__c -gt 70) ) -or
        ( ($_.Priority -eq "P1") -and ($_.Entitlement_Type__c -match "Extended") ) -or
        ( ($_.Priority -eq "P2") -and ($_.Entitlement_Type__c -match "Extended") ) -or
        ($_.Tier__c -eq "Premium") ) -and
        ($_.Case_Owner_Name__c -eq $null)
    })
}

function Filter-Unassigned {
    param (
        $InputList
    )
    return ($InputList | ? {$_.Case_Owner_Name__c -eq $null} )
}

function Filter-Assigned {
    param (
        $InputList
    )
    return ($InputList | ? {$_.Case_Owner_Name__c -ne $null} )
}

function Filter-Language {
    param (
        $Language, $InputList
    )

    if ($Language -eq "KO") {
        $OutList = $InputList | ? { ($_.Preferred_Case_Language__c -eq "Korean") -and ($_.Case_Owner_Name__c -eq $null)}
    } elseif ($Language -eq "CN") {
        $OutList = $InputList | ? { 
            (($_.Preferred_Case_Language__c -eq "Mandarin") -or 
            ($_.Preferred_Case_Language__c -eq "Chinese Traditional") -or 
            ($_.Preferred_Case_Language__c -eq "Chinese Simplified")) -and
            ($_.Case_Owner_Name__c -eq $null)
        }
    } elseif ($Language -eq "EN") {
        $OutList = $InputList | ? { 
            !(($_.Preferred_Case_Language__c -eq "Mandarin") -and 
            ($_.Preferred_Case_Language__c -eq "Chinese Traditional") -and
            ($_.Preferred_Case_Language__c -eq "Chinese Simplified") -and
            ($_.Preferred_Case_Language__c -eq "Korean")) -and
            ($_.Case_Owner_Name__c -eq $null)
        }
    }

    return $OutList
}

function Filter-Timezone {
    param($Timezone, $InputList)

    # return ($InputList | ? {($_.Case_Preferred_Timezone__c).split(" ")[0] -eq "India"} )
    return ( $InputList | ? { ($_.Case_Preferred_Timezone__c -like "*India*")  -and ($_.Case_Owner_Name__c -eq $null)} )
}

function Get-AESTDate {
    return [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'AUS Eastern Standard Time').ToString("dd/MM/yy hh:mm:ss")
}

function update_sheet {
    [CmdletBinding()]
    param($sheet, $list)

    $sheet.Unprotect()

    $hyperroot = "https://tableau.my.salesforce.com/"
    # $datetime = (Get-Date).ToString("yy-MM-dd HH:mm")
    # $datetime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'AUS Eastern Standard Time').ToString("yyyy-MM-dd hh:mm:ss")
    # $datetime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'AUS Eastern Standard Time').ToString("dd-MM-yy hh:mm:ss")
    $datetime = Get-AESTDate
    Write-Host $sheet.Name "is being updated at $datetime" -ForegroundColor Cyan

    # Clear all contents 
    $sheet.Cells.ClearContents() | Out-Null
    $sheet.Cells.ClearFormats() | Out-Null
    # $sheet.UsedRange.ClearContents()

    $sheet.Cells.Item(1, 13) = "Updated: $datetime"
    $sheet.Cells.Item(1, 3) = "Count:"
    $q_list = ($list | ? { $_.Case_Owner_Name__c -eq $null })
    $a_list = ($list | ? { $_.Case_Owner_Name__c -ne $null })
    
    ## Queue Count
    $sheet.Cells.Item(1, 4) = $q_list.Count
    $sheet.Cells.Item(1, 4).Font.Bold = $true
    ## Assigned Count
    $sheet.Cells.Item(1, 5) = $a_list.Count
    ## Total Count
    $sheet.Cells.Item(1, 6) = $list.Count
    $sheet.Cells.Item(1, 6).Interior.ColorIndex = 6

    ## Create a table header
    $header = @{
        "Case_Owner_Name__c"           = 1
        "Protip"                       = 2
        "CaseNumber"                   = 3
        "Escalated_Case__c"            = 4
        "Account.CSM_Name__c"          = 5
        "Priority"                     = 6
        "Case_Age__c"                  = 7
        "Entitlement_Type__c"          = 8
        "First_Response_Complete__c"   = 9
        "Preferred_Case_Language__c"   = 10
        "Case_Preferred_Timezone__c"   = 11
        "Product__c"                   = 12
        "Category__c"                  = 13
        "Subject"                      = 14
        "Account.Name"                 = 15
        "HandOff"                      = 16
        'Account.ARR'                   = 17
    }
    # $header['language']

    # $sheet.Name
    $sheet.Cells.Item(2, $header['Case_Owner_Name__c']) = "Engineers"
    $sheet.Cells.Item(2, $header['Case_Owner_Name__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Case_Owner_Name__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Protip']) = "Protip"
    $sheet.Cells.Item(2, $header['Protip']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Protip']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['CaseNumber']) = "CaseNumber"
    $sheet.Cells.Item(2, $header['CaseNumber']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['CaseNumber']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Escalated_Case__c']) = "Esclt"
    $sheet.Cells.Item(2, $header['Escalated_Case__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Escalated_Case__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Account.CSM_Name__c']) = "CSM/ARR"
    $sheet.Cells.Item(2, $header['Account.CSM_Name__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Account.CSM_Name__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Priority']) = "P"
    $sheet.Cells.Item(2, $header['Priority']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Priority']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Case_Age__c']) = "Age"
    $sheet.Cells.Item(2, $header['Case_Age__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Case_Age__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Entitlement_Type__c']) = "Entitlement"
    $sheet.Cells.Item(2, $header['Entitlement_Type__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Entitlement_Type__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['First_Response_Complete__c']) = "F/R"
    $sheet.Cells.Item(2, $header['First_Response_Complete__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['First_Response_Complete__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Preferred_Case_Language__c']) = "Language"
    $sheet.Cells.Item(2, $header['Preferred_Case_Language__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Preferred_Case_Language__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Case_Preferred_Timezone__c']) = "TimeZone"
    $sheet.Cells.Item(2, $header['Case_Preferred_Timezone__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Case_Preferred_Timezone__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Product__c']) = "Product"
    $sheet.Cells.Item(2, $header['Product__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Product__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Category__c']) = "Category"
    $sheet.Cells.Item(2, $header['Category__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Category__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Subject']) = "Subject"
    $sheet.Cells.Item(2, $header['Subject']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Subject']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Account.Name']) = "Account"
    $sheet.Cells.Item(2, $header['Account.Name']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Account.Name']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['HandOff']) = "HandOff"
    $sheet.Columns.Item('Q').NumberFormat = "$#,##0"

    # Freeze Panes
    $sheet.Select()
    $sheet.Application.ActiveWindow.SplitRow = 2
    $sheet.Application.ActiveWindow.FreezePanes = $true

    $i = 3
    foreach ($row in $list) {
        # Unhide rows that was hidden previously
        $sheet.Rows($i).Hidden = $false

        # Engineer Names        
        if ($row.Case_Owner_Name__c -ne $null) {
            # $sheet.Cells.Item($i, $header['Case_Owner_Name__c']) = $row.Case_Owner_Name__c.Split(" ")[0]
            $sheet.Cells.Item($i, $header['Case_Owner_Name__c']) = $row.Case_Owner_Name__c

        }

        # Case Number
        # $sheet.Cells.Item($i, $header['CaseNumber']) = $row.CaseNumber
        # Write-Host $row.Description
        $sheet.Hyperlinks.Add(
            $sheet.Cells.Item($i, $header['CaseNumber']),
            $hyperroot+$row.Id,
            "",
            $hyperroot+$row.Id,
            $row.CaseNumber
        ) | Out-Null

        # Created Date
        # $sheet.Cells.Item($i, 3) = ([datetime]$row.CreatedDate).ToString("yyyy-MM-dd HH:mm")
        # $sheet.Cells.Item($i, 2) = $row.CreatedDate

        # Escalated
        if ($row.Escalated_Case__c -eq $true) {
            $sheet.Cells.Item($i, $header['Escalated_Case__c']) = "YES"
            $sheet.Cells.Item($i, $header['Escalated_Case__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['Escalated_Case__c']).Font.ColorIndex = 3
        } else {
            # $sheet.Cells.Item($i, 12) = $row.Escalated_Case__c
        }

        # Priority
        if ($row.Priority -eq "P1") {
            $sheet.Cells.Item($i, $header['Priority']) = $row.Priority
            $sheet.Cells.Item($i, $header['Priority']).Font.ColorIndex = 3
            $sheet.Cells.Item($i, $header['Priority']).Font.Bold = $true
        }
        elseif ($row.Priority -eq "P2") {
            $sheet.Cells.Item($i, $header['Priority']) = $row.Priority
            $sheet.Cells.Item($i, $header['Priority']).Font.ColorIndex = 46
            # $sheet.Cells.Item($i, 4).Font.Bold = $true
        }
        elseif ($row.Priority -eq "P3") {
            $sheet.Cells.Item($i, $header['Priority']) = $row.Priority
            $sheet.Cells.Item($i, $header['Priority']).Font.ColorIndex = 5
        }
        elseif ($row.Priority -eq "P4") {
            $sheet.Cells.Item($i, $header['Priority']) = $row.Priority
            $sheet.Cells.Item($i, $header['Priority']).Font.ColorIndex = 10
        }
        else {
            $sheet.Cells.Item($i, $header['Priority']) = $row.Priority
        }

        # Case Age
        if ($row.Case_Age__c -lt 100) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
        }
        elseif (($row.Case_Age__c -ge 100) -and ($row.Case_Age__c -le 199)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 6
        }
        elseif (($row.Case_Age__c -ge 200) -and ($row.Case_Age__c -le 299)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 45
        }
        elseif (($row.Case_Age__c -ge 300) -and ($row.Case_Age__c -le 399)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 46
        }
        elseif (($row.Case_Age__c -ge 400) -and ($row.Case_Age__c -le 499)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 3
        }
        elseif (($row.Case_Age__c -ge 500) -and ($row.Case_Age__c -le 599)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 53
        } else {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            Write-Host "I am Too Old" $row.Case_Age__c "hours" -ForegroundColor Yellow
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 1
            $sheet.Cells.Item($i, $header['Case_Age__c']).Font.ColorIndex = 2
        }

        # Entitlement
        if ( ($row.Entitlement_Type__c -eq "Premium") -or ($row.Entitlement_Type__c -eq "Elite") -or ($row.Entitlement_Type__c -eq "Extended") ) {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 3
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.ColorIndex = 2
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.Bold = $true
        } else {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
        }

        # First Response
        if ($row.First_Response_Complete__c -eq $true) {
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "YES"
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 10 # Green
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"

        } elseif ($row.First_Response_Complete__c -eq $false) {
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "no"
            # $sheet.Cells.Item($i, 4).Font.ColorIndex = 16 #Grey

            if ( (($row.Priority -eq 'P1') -and ( ($row.Case_Age__c -ge 6) -and ($row.Case_Age__c -le 8) )) -or (($row.Priority -eq 'P2') -and ( ($row.Case_Age__c -ge 20) -and ($row.Case_Age__c -le 24) )) ) {
                $sheet.Range("A$i","O$i").interior.colorindex = 27
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
            }

            if ( ( ($row.Priority -eq 'P1') -and ($row.Case_Age__c -gt 8) ) -or ( ($row.Priority -eq 'P2') -and ($row.Case_Age__c -gt 24) ) ) {
                $sheet.Range("A$i","O$i").interior.colorindex = 38
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"
            }

        } else {
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.First_Response_Complete__c
        }

        # Language
        if ($row.Preferred_Case_Language__c -eq "Korean") {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = ($row.Preferred_Case_Language__c).Split(" ")[0]
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']).Font.ColorIndex = 43 # DarkGreen # 10 # Green
        }
        elseif (($row.Preferred_Case_Language__c -eq "Mandarin") -or ($row.Preferred_Case_Language__c -eq "Chinese Traditional") -or ($row.Preferred_Case_Language__c -eq "Chinese Simplified")) {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = ($row.Preferred_Case_Language__c).Split(" ")[0]
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']).Font.ColorIndex = 45
        }
        else {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = ($row.Preferred_Case_Language__c).Split(" ")[0]
        }

        if ($row.Case_Preferred_Timezone__c) {
            $sheet.Cells.Item($i, $header['Case_Preferred_Timezone__c']) = ($row.Case_Preferred_Timezone__c).split(" ")[0]
        }
        # Product
        $sheet.Cells.Item($i, $header['Product__c']) = ($row.Product__c).Split(" ")[1]
        if (($row.Product__c -eq "Tableau Desktop") -or ($row.Product__c -eq "Tableau Public Desktop") -or ($row.Product__c -eq "Tableau Reader") -or ($row.Product__c -eq "Tableau Prep") -or ($row.Product__c -eq "Tableau Public Desktop") ) {
            $sheet.Cells.Item($i, $header['Product__c']).Interior.ColorIndex = 35
            # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.ThemeColor = xlThemeColorAccent3
            # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.TintAndShade = 0.6
        } else {
            $sheet.Cells.Item($i, $header['Product__c']).Interior.ColorIndex = 37
            # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.ThemeColor = xlThemeColorAccent3
            # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.TintAndShade = 0.6
        }
        
        # Category
        $sheet.Cells.Item($i, $header['Category__c']) = $row.Category__c
        <#
        if ($row.Category__c -eq "Data Connectivity") {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 15
        } elseif ($row.Category__c -eq "Licensing") {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 34
        } elseif (($row.Category__c -eq "Performance")-or ($row.Category__c -eq "Stability")) {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 45
        } elseif ($row.Category__c -eq "Security") {
            $sheet.Cells.Item($i, $header['Category__c']).Font.ColorIndex = 3
        } elseif ($row.Category__c -eq "Authentication") {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 43
        } elseif ($row.Category__c -eq "View Rendering") {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 24
        } elseif ($row.Category__c -eq "Publishing") {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 36
        } elseif ($row.Category__c -like "Installation*") {
            $sheet.Cells.Item($i, $header['Category__c']).Interior.ColorIndex = 4
        }
        #>
        
        # Subject
        $sheet.Cells.Item($i, $header['Subject']) = $row.Subject

        # POA Status
        # $sheet.Cells.Item($i, 13) = $row.Case_Owner_Name__c
        # $sheet.Cells.Item($i, 14) = $row.Histories
        # $sheet.Cells.Item($i, 14) = $row.Person_Owner__c

        # Accounts
        if (($row.Account.Name -eq "Westpac Banking Corporation") -or ($row.Account.Name -eq "GIC Private Limited") -or ($row.Account.Name -eq "Applied Materials, Inc.")) {
            $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name
            $sheet.Cells.Item($i, $header['Account.Name']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['Account.Name']).Interior.ColorIndex = 6
        } elseif($row.Account.Name -eq "Hang Seng Bank Limited") {
            $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name
            $sheet.Cells.Item($i, $header['Account.Name']).Font.Bold = $true
            # $sheet.Cells.Item($i, $header['Account.Name']).Interior.ColorIndex = 3

            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = "Premium"
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 3
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.ColorIndex = 2
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.Bold = $true

        }       
        else {
            $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name
        }

        # Protips Type 1
        $sheet.Cells.Item($i, $header['Protip']) = $row.Feeds.records
        # Protips Type 2
        if ( ($row.Plan_of_Action_Status__c -ne $null) -and ($row.Plan_of_Action_Status__c -like "*protip*") ) {
            $sheet.Cells.Item($i, $header['Feeds']) = $row.Plan_of_Action_Status__c
            $sheet.Cells.Item($i, $header['Feeds']) = "YES"
        }

        # $sheet.Cells.Item($i, $header['Account.CSM_Name__c']) = $row.Account.CSM_Name__c
        <#
        $sheet.Hyperlinks.Add(
            $sheet.Cells.Item($i, $header['Account.CSM_Name__c']),
            "mailto:"+$row.Account.CSM_Email__c, # hyperlink
            "",
            $row.Account.CSM_Email__c, # tooltip text
            $row.Account.CSM_Name__c # Cell display text
        ) | Out-Null
        #>
        <#
        if ($row.Account.CSM_Name__c -ne $null) {
            $sheet.Hyperlinks.Add(
                $sheet.Cells.Item($i, $header['Account.CSM_Name__c']),
                "mailto:"+$row.Account.CSM_Email__c, # hyperlink
                # "",
                "",
                $row.Account.CSM_Name__c+"("+$row.Account.CSM_Email__c+")", # tooltip text
                "YES" # Cell display text
            ) | Out-Null
        }
        #>

        # Write-Host $row.CSM
        $sheet.Cells.Item($i, $header['Account.CSM_Name__c']) = $row.CSM
        # $sheet.Cells.Item($i, $header['Account.CSM_Name__c']) = $row.Account.CSM_Name__c
        # $sheet.Cells.Item($i, $header['Account.CSM_Name__c']).Font.Bold = $true
        $sheet.Cells.Item($i, $header['Account.CSM_Name__c']).HorizontalAlignment = -4152
        if ($row.Account.CSM_Name__c -like "*B") {
            # $sheet.Cells.Item($i, $header['Account.CSM_Name__c']).Font.ColorIndex = 3
        }if ($row.Account.CSM_Name__c -like "*M") {
            # $sheet.Cells.Item($i, $header['Account.CSM_Name__c']).Font.ColorIndex = 10
        }if ($row.Account.CSM_Name__c -like "*K") {
            # $sheet.Cells.Item($i, $header['Account.CSM_Name__c']).Font.ColorIndex = 46
        }
        <#
        Write-Host $row.CaseNumber -NoNewline
        Write-Host $row.Histories -ForegroundColor Red -NoNewline
        if ($row.Feeds.records -eq 'YES') {
            Write-Host $row.Feeds.records -ForegroundColor Green
        } else {
            Write-Host $row.Feeds.records -ForegroundColor DarkCyan
        }
        #>
        # $sheet.Cells.Item($i, $header['Status']) = $row.Status
        if ( ($row.Plan_of_Action_Status__c -ne $null) -and ($row.Plan_of_Action_Status__c -like "*apac*") ) {
            # $sheet.Cells.Item($i, $header['HandOff']) = $row.Plan_of_Action_Status__c
            $sheet.Cells.Item($i, $header['HandOff']) = "To APAC"
        }

        # $sheet.Cells.Item($i, $header['Account.ARR']) = $row.Account.cnvAmnt
        # $sheet.Cells.Item($i, $header['Account.ARR']) = $row.Account.frmtAmnt

        
        # Validating FR
        <#
        # $r = "A"+$i+":C"+$i
        try {
            Write-Host $row.casemilestones.records.TimeRemainingInHrs
            $hr = ($row.casemilestones.records.TimeRemainingInHrs).Split(":")[0]
        } catch {
            $hr = $null
        }

        if ( ($row.Priority -eq 'P1') -and ($row.First_Response_Complete__c -eq $false) ) {
            # Write-Host $r
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.casemilestones.records.TimeRemainingInHrs

            if ( ($hr -lt 2) -and ($hr -ge 0) ) {
                $sheet.Range("A$i","O$i").interior.colorindex = 27
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
            }
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.casemilestones.records.TimeRemainingInHrs
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $hr

        } elseif ( ($row.Priority -eq 'P2') -and ($row.First_Response_Complete__c -eq $false) ) {
            # Write-Host $r
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.casemilestones.records.TimeRemainingInHrs
            if (( $hr -lt 4) -and ($hr -ge 0) ) {
                $sheet.Range("A$i","O$i").interior.colorindex = 27
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
                $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
            }
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.casemilestones.records.TimeRemainingInHrs
        }
        
        if ( ($row.Priority -eq 'P1') -and ($row.First_Response_Complete__c -eq $false) -and ($row.casemilestones.records.IsViolated -eq $true) ) {
            # Write-Host $r
            $sheet.Range("A$i","O$i").interior.colorindex = 38
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.casemilestones.records.TimeRemainingInHrs
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $hr
        } elseif ( ($row.Priority -eq 'P2') -and ($row.First_Response_Complete__c -eq $false) -and ($row.casemilestones.records.IsViolated -eq $true) ) {
            # Write-Host $r
            $sheet.Range("A$i","O$i").interior.colorindex = 38
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.casemilestones.records.TimeRemainingInHrs
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $hr
        }
        #>

        # Entitlement
        if ( ($row.Entitlement_Type__c -match "Premium") -or ($row.Entitlement_Type__c -eq "Elite") -or ($row.Entitlement_Type__c -eq "Extended") ) {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 3
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.ColorIndex = 2
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.Bold = $true
        } else {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
        }

        # Hide Rows if the condition met
        # Write-Host $sheet.Name
        if ( ($sheet.Name -match "Server" -or $sheet.Name -match "All") -and ($row.Entitlement_Type__c -match "Standard") -and ($row.Priority -eq "P3" -or $row.Priority -eq "P4") -and ($row.Case_Age__c -lt 63) -and ($row.Case_Owner_Name__c -eq $null) ) {
            $sheet.Rows($i).Hidden = $true
        }
        
        $i++
    }

    # Column Autofit
    $sheet.columns.AutoFit() | Out-Null
    $sheet.Cells.Item($header['Case_Owner_Name__c']).ColumnWidth = 13
    $sheet.Cells.Item($header['Category__c']).ColumnWidth = 25
    $sheet.Cells.Item($header['Subject']).ColumnWidth = 100
    # $sheet.Cells.Item($header['Account.Name']).ColumnWidth = 35
    Write-Host $sheet.Name " update completed" -ForegroundColor DarkCyan
    # $sheet.Cells.Item($i,1) = "End of Record"
    # $i++
    # $sheet.Cells.Item($i,1).EntireRow.Delete()
    $sheet.Protect('',0,1,0,0,1,0,1,0,0,1,0,1,0,1,1)
}

function Remove-ExcelFile {
    param($FileName)

    # $lockedFile = "APACTechSupRealTimeDashboard.xlsx"
    $lockedFile = $FileName

    $allProcesses = Get-Process
    #$lockedFile is the file path
    foreach ($process in $allProcesses) { 
        $process.Modules | where {$_.FileName -eq $lockedFile} | Stop-Process -Force -ErrorAction SilentlyContinue    
    }

    # $excel = New-Object -ComObject Excel.Application
    # $excel.Application.Visible = $true
    $workbook = $excel.Workbooks.Open($lockedFile)
    $excel.DisplayAlerts = 'False'
    $workbook.SaveAs($lockedFile)
    $workbook.Close
    $excel.DisplayAlerts = 'False'
    $excel.Quit()

    Write-Host "Terminating Excel Process"
    Get-Process | Where-Object { $_.ProcessName -match "Excel" } | Stop-Process -Force
    get-process excel | select MainWindowTitle, Id, StartTime
    Write-Host "Terminating Excel Process X2"
    Get-Process | Where-Object { $_.ProcessName -match "Excel" } | Stop-Process -Force

    Remove-Item $lockedFile
    Write-Host "$lockedFile has been removed" -ForegroundColor Red
}

function Create-ExcelFile {
    param($OutFile)

    Write-Host "Creating an Excel File"
    ### Creating Excel Sheets ###
    $excel = New-Object -ComObject Excel.Application
    $excel.Application.Visible = $true
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    $sheet_in= $workbook.Sheets.Item(1)
    $sheet_in.Name = "India"

    $sheet_cn= $workbook.Sheets.Add()
    $sheet_cn.Name = "Chinese"

    $sheet_ko = $workbook.Sheets.Add()
    $sheet_ko.Name = "Korean"

    $sheet_uss = $workbook.Sheets.Add()
    $sheet_uss.Name = "Unassigned"

    $sheet_ass = $workbook.Sheets.Add()
    $sheet_ass.Name = "Assigned"

    $sheet_dsk = $workbook.Sheets.Add()
    $sheet_dsk.Name = "Desktop"
    $sheet_dsk.Tab.ColorIndex = 10
    
    $sheet_srv = $workbook.Sheets.Add()
    $sheet_srv.Name = "Server"
    $sheet_srv.Tab.ColorIndex = 5

    $sheet_p1p2 = $workbook.Sheets.Add()
    $sheet_p1p2.Name = "P1P2"
    $sheet_p1p2.Tab.ColorIndex = 46

    $sheet_p3p4 = $workbook.Sheets.Add()
    $sheet_p3p4.Name = "Aged P3P4"
    $sheet_p3p4.Tab.ColorIndex = 6
    
    $sheet_pre = $workbook.Sheets.Add()
    $sheet_pre.Name = "Premium"
    $sheet_pre.Tab.ColorIndex = 3
    
    # $sheet_ugt = $workbook.Sheets.Add()
    # $sheet_ugt.Name = "Action Now!"

    $sheet_all = $workbook.Sheets.Add()
    $sheet_all.Name = "All"

    Write-Host $OutFile
    $excel.DisplayAlerts = $false
    $workbook.SaveAs($OutFile)
    # $workbook.SaveAs($OutFile, 51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)
    Write-Host $OutFile
    $excel_pid = (Get-Process excel -ea 0 | Where-Object { $_.MainWindowTitle -like "*$OutFile*" }).Id
    Write-Host "Excel PID: $excel_pid"
    $workbook.Close($false)
    $excel.DisplayAlerts = $false
    $excel.Quit()
    Get-Process excel -ea 0 | Where-Object { $_.id -like $excel_pid } | Stop-Process
}

### Slack ###
function MessageTo-Slack {
    param($ChannelUri, $Message, $Type)

    if ($type -eq "Warning") {
        $text = ":warning:"
    } else {
        $text = ":information_source:"
    }
    
    $body = ConvertTo-Json @{
        text="$text $Message"
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-RestMethod -Method POST -ContentType "application/json" -uri $ChannelUri -Body $body | Out-Null
}

function Slack-Mrkdwn {
    [CmdletBinding()]
    param ($Text)
    return (ConvertTo-Json -Depth 10 @{blocks=@(@{type="section";text=@{type="mrkdwn";text="$Text"}})})
}

function Remove-Cache {
    param (
        $Path
    )
    Remove-Item -Path "$Path\*" -Recurse -Force
}


function Run-MainLoop {
    param($InFile)

    # Timeout for restarting Excel as Excel get buggy after running for long time
    # $timeout = New-TimeSpan -Days 7
    $timeout = New-TimeSpan -Hours 600

    while ($true) {
        # Remove-Cache -Path $CacheLocation
        # Excel.Application.DisplayAlerts=$False
        Write-Host "A new Excel object instaniated!"
        $excel = New-Object -ComObject Excel.Application
        # $excel.Visible = $true
        # Remove-Cache -Path $CacheLocation
        $excel.AutoRecover.Enabled = $False
        $excel.DisplayAlerts = $False
        $excel.Application.AutoRecover.Enabled = $False
        $excel.Application.DisplayAlerts = $False
        $workbook = $excel.Workbooks.Open($InFile)
        # $workbook = $excel.Workbooks.Open($sfile)

        $sheet_in = $workbook.sheets.item("India")
        $sheet_cn = $workbook.sheets.item("Chinese")
        $sheet_ko = $workbook.sheets.item("Korean")
        $sheet_uss = $workbook.sheets.item("Unassigned")
        $sheet_ass = $workbook.sheets.item("Assigned")
        $sheet_dsk = $workbook.sheets.item("Desktop")
        $sheet_srv = $workbook.sheets.item("Server")
        $sheet_p1p2 = $workbook.sheets.item("P1P2")
        $sheet_p3p4 = $workbook.sheets.item("Aged P3P4")
        $sheet_pre = $workbook.sheets.item("Premium")
        # $sheet_ugt = $workbook.sheets.item("Action Now!")
        $sheet_all = $workbook.sheets.item("All")

        $cur_array = @()
        # 
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        
        while ($sw.Elapsed -lt $timeout) {
            # Remove-Cache -Path $CacheLocation
            # Get new list - Get-All has it' own mechanism to infinitive re-try in error conditions. So, $query_success is useless here...
            # $all = get-all -Query $query | Sort-Object Case_Owner_Name__c, @{Expression="Tier__c";Ascending=$true}, Priority,@{Expression="CSM";Descending=$true},@{Expression="Case_Age__c";Descending=$true}
            $all = get-all -Query $query | Sort-Object Case_Owner_Name__c, @{Expression="Entitlement_Type__c";Ascending=$true}, Priority,@{Expression="CSM";Descending=$true},@{Expression="Case_Age__c";Descending=$true}
            # $all = filter-oldp3p4 -In $all

            $new_array = $all | Select-Object -Property CaseNumber, Escalated_Case__c, Priority, First_Response_Complete__c,Product__c, Case_Owner_Name__c, Feeds.records
            
            # Compare the two array - new and current if any changes
            $diff =  Compare-Object -ReferenceObject $cur_array -DifferenceObject $new_array -Property CaseNumber, Priority, First_Response_Complete__c, Product__c, Case_Owner_Name__c, Feeds.records -PassThru -ErrorAction SilentlyContinue
            
            # $diff -eq $null means that there is no difference in $cur_array & $new_array, the changes in terms of number of rows and contents of the columns (CaseNumber, Priority, FR, Owner) value changes.
            if ( ($diff -ne $null) ) {
                Write-Host "New changes detected..." -ForegroundColor Red
                # display what's different
                $diff | %{ Write-Host $_ }

                # Filter the list
                $desktop = filter-desktop -In $all
                # Applying old P3/P4 filter for Server Sheet
                $server = filter-server -In $all
                # $server = filter-oldp3p4 -In $server

                $premium = filter-premium -In $all
                # $urgent = filter-urgent -In $all
                $ass = Filter-Assigned -InputList $all
                $uss = Filter-Unassigned -InputList $all
                $ko = Filter-Language -Language "KO" -InputList $all
                $cn = Filter-Language -Language "CN" -InputList $all
                $in = Filter-Timezone -InputList $all
                $p1p2 = Filter-P1P2 -InputList $all
                $p3p4 = Filter-P3P4 -InputList $all
                # Applying old P3/P4 filter for Server Sheet - !! IMPORTANT to put $all in the last order. 
                # $all = filter-oldp3p4 -In $all

                # Update sheets
                $excel.AutoRecover.Enabled = $False
                $excel.DisplayAlerts=$False
                $excel.Application.DisplayAlerts=$False
                # $workbook.Save()
                # Write-Host "Saved workbook before commit" -ForegroundColor Green

                <#
                update_sheet -sheet $sheet_dsk -list $desktop
                update_sheet -sheet $sheet_srv -list $server -ErrorAction Stop
                update_sheet -sheet $sheet_pre -list $premium -ErrorAction Continue
                update_sheet -sheet $sheet_p1p2 -list $p1p2 -ErrorAction Continue
                update_sheet -sheet $sheet_p3p4 -list $p3p4 -ErrorAction Continue
                update_sheet -sheet $sheet_all -list $all -ErrorAction Continue

                # update_sheet -sheet $sheet_ugt -list $urgent -ErrorAction Continue
                update_sheet -sheet $sheet_ass -list $ass -ErrorAction Continue
                update_sheet -sheet $sheet_uss -list $uss -ErrorAction Continue
                update_sheet -sheet $sheet_ko -list $ko -ErrorAction Continue
                update_sheet -sheet $sheet_cn -list $cn -ErrorAction Continue
                update_sheet -sheet $sheet_in -list $in -ErrorAction Continue
                #>

                
                try {
                    update_sheet -sheet $sheet_dsk -list $desktop
                    update_sheet -sheet $sheet_srv -list $server -ErrorAction Stop
                    update_sheet -sheet $sheet_pre -list $premium -ErrorAction Continue
                    update_sheet -sheet $sheet_p1p2 -list $p1p2 -ErrorAction Continue
                    update_sheet -sheet $sheet_p3p4 -list $p3p4 -ErrorAction Continue
                    update_sheet -sheet $sheet_all -list $all -ErrorAction Continue

                    # update_sheet -sheet $sheet_ugt -list $urgent -ErrorAction Continue
                    update_sheet -sheet $sheet_ass -list $ass -ErrorAction Continue
                    update_sheet -sheet $sheet_uss -list $uss -ErrorAction Continue
                    update_sheet -sheet $sheet_ko -list $ko -ErrorAction Continue
                    update_sheet -sheet $sheet_cn -list $cn -ErrorAction Continue
                    update_sheet -sheet $sheet_in -list $in -ErrorAction Continue
                } catch {
                    Write-Warning $Error[0]
                    $error_name = $Error[0].Exception.GetType().FullName
                    MessageTo-Slack -ChannelUri $uri -Message "Error occurs while updating Sheets with $error_name" -Type "Warning"
                    Copy-Item $sfile -Destination "C:\Downloads" -Recurse -Force
                    Remove-ExcelFile -FileName $sfile
                    $cntdwn = 10
                    do {
                        Start-Sleep -Seconds 1
                        Write-Host "Angel restart count down $cntdwn"
                        $cntdwn--
                    } while ($cntdwn -gt 0)
                    
                    start powershell.exe C:\MyProjects\ps\Angel\Restart-Angel.ps1
                }
                

                # Console Output
                $all | ft @{L = 'Protip'; E = { $_.Case_Age__c } }, CaseNumber, Priority, @{L = 'Age'; E = { $_.Case_Age__c } }, Status, @{L = 'Language'; E = { $_.Preferred_Case_Language__c } }, Tier__c, @{L = 'Category'; E = { $_.Category__c } }, Product__c, Subject -AutoSize | Out-String
                Write-Host "Total Queue Size: " $new_array.Count -ForegroundColor Yellow
                $datetime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Write-Host "Last update: $datetime"

                $cur_array = $new_array
            } else {
                Write-Host "No change detected..." -ForegroundColor Green
            }
            Write-Host "Time Elapsed: " -ForegroundColor Magenta -NoNewline
            Write-Host $sw.Elapsed -ForegroundColor Magenta

            # Writing Ping time on the sheets
            $datetime = Get-AESTDate
            $keepalive = "Check-In: $datetime"
            Write-Host "Ping Time: $keepalive" -ForegroundColor Gray

            foreach ($asheet in $workbook.sheets) {
                if ($asheet.Name -ne "Aged P3 P4") {
                    $asheet.Unprotect()
                    $asheet.Cells.Item(1, 14) = $keepalive
                    $asheet.Cells.Item(1, 14).HorizontalAlignment = -4131
                    $asheet.Protect('',0,1,0,0,1,0,1,0,0,1,0,1,0,1,1)
                } else {
                    $asheet.Unprotect()
                    $asheet.Cells.Item(1, 14) = $keepalive
                    $asheet.Cells.Item(1, 14).HorizontalAlignment = -4131
                }
            }

            $excel.DisplayAlerts=$False
            $excel.Application.DisplayAlerts=$False
            $excel.AutoRecover.Enabled = $False
            # $workbook.Save()
            # $workbook.SaveAs($sfile)
            # Write-Host "Change Saved" -ForegroundColor Green
            # Start-Sleep -Seconds 2
        }
        Write-Host "Times Up! Killing Excel Process" -ForegroundColor Red
        MessageTo-Slack -ChannelUri $uri -Message "Angel Times up"
        # Save current sheets
        $excel.DisplayAlerts = $false
        # $workbook.Save()
        # Write-Host "File Saved before Close" -ForegroundColor DarkGreen
        $workbook.Close($false)
        Write-Host "Workbook Closed" -ForegroundColor DarkGreen
        $excel.Quit()
        Write-Host "Excel Quit" -ForegroundColor DarkGreen
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        Remove-Variable $excel
        [GC]::Collect()
        Get-Process | ? { $_.ProcessName -match "Excel" }
        Get-Process | ? { $_.ProcessName -match "Excel" } | Stop-Process
        # Remove-ExcelFile -FileName $sfile
        start powershell.exe C:\MyProjects\ps\Angel\Restart-Angel.ps1
    }
}



## Main
$Title = "RealTime Angel"
$host.UI.RawUI.WindowTitle = $Title

# Remove-Item -Path "C:\Users\$env:USERNAME\AppData\Local\Microsoft\Office\16.0\OfficeFileCache1\*" -Recurse -Force
# Remove-ExcelFile -FileName $sfile
# Start-Sleep -Seconds 2
# Remove-ExcelFile -FileName $sfile


if (!(Test-Path $sfile -PathType Leaf)) {
    Create-ExcelFile -OutFile $sfile
}
$Text = Slack-Mrkdwn -Text "@here Angel is Starting..."
MessageTo-Slack -ChannelUri $uri -Message "Starting Realtime Angel at $web_lnk"


Run-MainLoop -InFile $sfile