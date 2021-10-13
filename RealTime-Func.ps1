$SMTPServer = 'smarthost.tsi.lan'
$SMTPPort = 25
$From = 'AngelNotification@TableauSoftware.com'


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

        if ($_.Account.CSM_Name__c -ne $null) {
            # Write-Host $_.Account.AnnualRevenue
            # Write-Host $_.Account.frmtAmnt
            # Write-Host $_.Account.cnvAmnt
            # $ARR = ToKMB($_.Account.AnnualRevenue)
            # $_.Account.CSM_Name__c = "YES"
            $_.CSM = "VIP"
        }

        if (($_.isClosed -eq $true) -and ( ($_.Case_Owner_Name__c -eq $null) -or ($_.Case_Owner_Name__c -eq '') ) ) {
            $_.Case_Owner_Name__c = "By Customer"
        }
        
        # Changed owner Today
        if ($_.Histories.records -ne $null) {
            $_.Histories = "YES"
        }

        # Protips
        <#
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
        #>

        if ($_.Entitlement_Type__c -eq 'TOL Premium') {
            $_.Entitlement_Type__c = 'Premium'
        }

        $new_obj += $_
    }

    #Filter the new object to keep the cases unassigned or assigned to somebody TODAY only.
    $new_obj = ($new_obj | Where-Object { ($_.Case_Owner_Name__c -eq $null) -or (  ($_.Case_Owner_Name__c -ne $null) -and ($_.Histories -eq "YES") ) })

    # $new_obj | Export-Csv -Path C:\MyProjects\PS\Angel\all.csv -Force -Encoding UTF8

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
        ($_.Product__c -eq "Tableau Prep Builder") -or
        ($_.Product__c -eq "Tableau Public Desktop") ) # -and ($_.Case_Owner_Name__c -eq $null)
    })
}

function filter-server {
    param($In)
    return ( $In | Where-Object { (
        ($_.Product__c -eq "Tableau Server") -or 
        ($_.Product__c -eq "Tableau Public Server") -or 
        ($_.Product__c -eq "Tableau Online") -or 
        ($_.Product__c -eq "Tableau Bridge") -or 
        ($_.Product__c -eq "Tableau Mobile") -or 
        ($_.Product__c -eq "Connector SDK") -or
        ($_.Product__c -eq "Extract API") -or
        ($_.Product__c -eq "Hyper API") -or
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

function filter-pizza {
    param($InputList)
    return ($InputList | ?{ (([datetime]$_.CreatedDate).ToUniversalTime() -gt [datetime]"8/7/2021 23:59") -or (($_.Entitlement_Type__c -match "Premium") -or ($_.Entitlement_Type__c -match "Extended") -or ($_.Entitlement_Type__c -match "Elite")) } ) # UTC # do not include 8/07/2021
}

function Get-AESTDate {
    return [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'AUS Eastern Standard Time').ToString("dd/MM/yy hh:mm:ss")
}


function Update-Stat {
    param($sheet)

    $sheet.Cells.ClearContents() | Out-Null
    $sheet.Cells.ClearFormats() | Out-Null


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


function Send-AngelNotification {
    [CmdletBinding()]
    param ($Message)

    try {
        Send-MailMessage `
        -SmtpServer $SMTPServer `
        -Port $SMTPPort `
        -From $From `
        -To jchoi@tableau.com `
        -Subject 'Angel Notification' `
        -Body $Message –BodyAsHtml `
        -Priority High `
        -ErrorAction Stop
        
        Write-Host "Email Notification has sent successfully"
    } catch {
        Write-Error "Email Error! Sending Notification failed!"
    }
}