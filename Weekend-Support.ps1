[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

<# Variable Block #>
$query = "SELECT 
    Id, 
    CaseNumber, 
    Priority,
    Description,
    Case_Age__c, 
    Status, 
    Preferred_Case_Language__c, 
    Preferred_Support_Region__c,
    Tier__c, 
    Category__c, 
    Product__c, 
    Subject, 
    First_Response_Complete__c, 
    CreatedDate, 
    Entitlement_Type__c, 
    Plan_of_Action_Status__c, 
    Case_Owner_Name__c,
    AccountId,
    Account.Name,
    IsEscalated, Escalated_Case__c,
    (SELECT CreatedDate, field, OldValue, NewValue, CreatedById FROM Histories WHERE CreatedDate=TODAY and field='Owner')
FROM case 
WHERE 
    Id IN (SELECT CaseID FROM CaseHistory ) AND
    RecordTypeId='012600000000nrwAAA' AND 
    (Status='New' or Status='Active' or Status='Re-opened') AND
    ((Entitlement_Type__c='Premium' AND Priority='P1') OR (Entitlement_Type__c='Premium' AND Priority='P2') OR (Entitlement_Type__c='Elite' AND Priority='P1') OR (Entitlement_Type__c='Extended' AND Priority='P1'))
ORDER BY 
    Preferred_Support_Region__c, Priority, Case_Owner_Name__c, Case_Age__c DESC" -replace "`n", " "

$user =             $env:USERNAME
$excel_file_name =  "Weekend-Support.xlsx"
$sp_local_path =    "C:\Users\$user\Tableau Software Inc\APAC Tech Support - Documents"
$sp_excel_path =    "$sp_local_path\$excel_file_name"
$sp_link =          "https://tableau.sharepoint.com/sites/APACTechSupport/Shared%20Documents/"+$excel_file_name+"?web=1"

$ts_apac_support =  "https://hooks.slack.com/services/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"
$ts_apac_sydney_py ="https://hooks.slack.com/services/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"

$sf_root =          "https://tableau.my.salesforce.com"

<# Function Block #>

function Create-ExcelFile {
    param($OutFile)

    ### Creating Excel Sheets ###
    $excel = New-Object -ComObject Excel.Application
    # $excel.Application.Visible = $true
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    $sheet_all = $workbook.Sheets.Item(1)
    $sheet_all.Name = "Assigned(APAC)"
    $sheet_all.Tab.ColorIndex = 2

    $sheet_uass = $workbook.Sheets.Add()
    $sheet_uass.Name = "Unassigned"
    $sheet_uass.Tab.ColorIndex = 10

    $sheet_ugt = $workbook.Sheets.Add()
    $sheet_ugt.Name = "Action Now!"
    $sheet_ugt.Tab.ColorIndex = 3

    $workbook.SaveAs($OutFile)
    # $excel.DisplayAlerts = $false
    Write-Host $OutFile
    $w_title = $OutFile.Split("\")[-1]
    Write-Host "Weekend Excel File Created at $w_title" -ForegroundColor Yellow
    $excel_pid = (Get-Process excel -ea 0 | Where-Object { $_.MainWindowTitle -like "*$w_title*" }).Id
    Write-Host "Excel PID: $excel_pid"
    $excel.DisplayAlerts = $false
    $workbook.Close($false)
    $excel.Quit()
    Get-Process excel -ea 0 | Where-Object { $_.id -like $excel_pid } | Stop-Process
}

function Write-Sheet {
    param($Sheet, $List)

    $hyperroot = "https://tableau.my.salesforce.com/"
    # $datetime = (Get-Date).ToString("yy-MM-dd HH:mm")
    $datetime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'AUS Eastern Standard Time').ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host $sheet.Name "is being updated at $datetime" -ForegroundColor Cyan

    # Clear all contents 
    $sheet.Cells.ClearContents() | Out-Null
    $sheet.Cells.ClearFormats() | Out-Null
    # $sheet.UsedRange.ClearContents()

    $sheet.Cells.Item(1, 11) = "Refreshed at: $datetime"
    # $sheet.Cells.Item(1, 2) = "$datetime"
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
    $sheet.Cells.Item(1, 6).Interior.ColorIndex = 6

    # $sheet.Shapes.AddTextbox(1,100, 100, 200, 50).TextFrame.Characters.Text = "Test Box"

    ## Create a table header
    $header = @{
        "Case_Owner_Name__c"           = 1
        "CaseNumber"                   = 2
        "Preferred_Support_Region__c"  = 3
        "Escalated_Case__c"            = 4
        "Priority"                     = 5
        "Case_Age__c"                  = 6
        "Entitlement_Type__c"          = 7
        "First_Response_Complete__c"   = 8
        "Preferred_Case_Language__c"   = 9
        "Product__c"                   = 10
        "Category__c"                  = 11
        "Subject"                      = 12
        "Account.Name"                 = 13
    }
    # $header['language']

    # $sheet.Name
    $sheet.Cells.Item(2, $header['Case_Owner_Name__c']) = "Engineers"
    $sheet.Cells.Item(2, $header['Case_Owner_Name__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Case_Owner_Name__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['CaseNumber']) = "CaseNumber"
    $sheet.Cells.Item(2, $header['CaseNumber']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['CaseNumber']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Preferred_Support_Region__c']) = "Region"
    $sheet.Cells.Item(2, $header['Preferred_Support_Region__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Preferred_Support_Region__c']).Interior.ColorIndex = 1
    $sheet.Cells.Item(2, $header['Escalated_Case__c']) = "Esclt"
    $sheet.Cells.Item(2, $header['Escalated_Case__c']).Font.ColorIndex = 2
    $sheet.Cells.Item(2, $header['Escalated_Case__c']).Interior.ColorIndex = 1
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

    $sheet.Select()
    $sheet.Application.ActiveWindow.SplitRow = 2
    $sheet.Application.ActiveWindow.FreezePanes = $true

    $sheet.Cells.Item(3, 2) = "No Action Required! Please Relax~"

    $i = 3
    foreach ($row in $list) {
        # Engineer Names        
        if ($row.Case_Owner_Name__c -ne $null) {
            # $sheet.Cells.Item($i, $header['Case_Owner_Name__c']) = $row.Case_Owner_Name__c.Split(" ")[0]
            $sheet.Cells.Item($i, $header['Case_Owner_Name__c']) = $row.Case_Owner_Name__c

        }

        # Case Number
        # $sheet.Cells.Item($i, $header['CaseNumber']) = $row.CaseNumber
        $sheet.Hyperlinks.Add(
            $sheet.Cells.Item($i, $header['CaseNumber']),
            $hyperroot+$row.Id,
            "",
            $hyperroot+$row.Id,
            $row.CaseNumber
        ) | Out-Null

        # $sheet.Cells.Item($i, $header['CaseNumber']).Hyperlinks(1) 

        # Created Date
        # $sheet.Cells.Item($i, 3) = ([datetime]$row.CreatedDate).ToString("yyyy-MM-dd HH:mm")
        if ($row.Preferred_Support_Region__c -eq "APAC"){
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']) = $row.Preferred_Support_Region__c
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']).Interior.ColorIndex = 10 # Green
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']).Font.ColorIndex = 2 # White
        } elseif ($row.Preferred_Support_Region__c -eq "EMEA") {
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']) = $row.Preferred_Support_Region__c
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']).Interior.ColorIndex = 6 # Yellow
        } else { # USCA
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']) = $row.Preferred_Support_Region__c
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']).Interior.ColorIndex = 5 # Blue
            $sheet.Cells.Item($i, $header['Preferred_Support_Region__c']).Font.ColorIndex = 2 # White
        }
        

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
        elseif (($row.Case_Age__c -gt 100) -and ($row.Case_Age__c -lt 199)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 6
        }
        elseif (($row.Case_Age__c -gt 200) -and ($row.Case_Age__c -lt 299)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 45
        }
        elseif (($row.Case_Age__c -gt 300) -and ($row.Case_Age__c -lt 399)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 46
        }
        elseif (($row.Case_Age__c -gt 400) -and ($row.Case_Age__c -lt 499)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 53
        }
        elseif (($row.Case_Age__c -gt 500) -and ($row.Case_Age__c -lt 599)) {
            $sheet.Cells.Item($i, $header['Case_Age__c']) = $row.Case_Age__c
            $sheet.Cells.Item($i, $header['Case_Age__c']).Interior.ColorIndex = 3
        }

        # Entitlement
        if ($row.Entitlement_Type__c -eq "Premium") {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 3
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.ColorIndex = 2
            # $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.Bold = $true
        } elseif ($row.Entitlement_Type__c -eq "Elite") {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 4
            # $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.ColorIndex = 2
        } elseif ($row.Entitlement_Type__c -eq "Extended") {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 6
        } else {
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = $row.Entitlement_Type__c
        }

        # First Response
        if ($row.First_Response_Complete__c -eq $true) {
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "YES"
            # $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 10 # Green
        } elseif ($row.First_Response_Complete__c -eq $false) {
            # $sheet.Cells.Item($i, 7) = "no"
            # $sheet.Cells.Item($i, 4).Font.ColorIndex = 16 #Grey
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = "NO"
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']).Font.ColorIndex = 3 # Red
        } else {
            $sheet.Cells.Item($i, $header['First_Response_Complete__c']) = $row.First_Response_Complete__c
        }

        # Language
        if ($row.Preferred_Case_Language__c -eq "Korean") {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = $row.Preferred_Case_Language__c
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']).Font.ColorIndex = 43 # DarkGreen # 10 # Green
        }
        elseif (($row.Preferred_Case_Language__c -eq "Mandarin") -or ($row.Preferred_Case_Language__c -eq "Chinese Traditional") -or ($row.Preferred_Case_Language__c -eq "Chinese Simplified")) {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = $row.Preferred_Case_Language__c
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']).Font.ColorIndex = 45
        }
        else {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = $row.Preferred_Case_Language__c
        }

        # Product
        $sheet.Cells.Item($i, $header['Product__c']) = $row.Product__c
        
        # Category
        $sheet.Cells.Item($i, $header['Category__c']) = $row.Category__c

        # Subject
        $sheet.Cells.Item($i, $header['Subject']) = $row.Subject

        # Account Name
        $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name

        $i++
    }
    # Column Autofit
    $sheet.columns.AutoFit() | Out-Null
    $sheet.Cells.Item($header['Product__c']).ColumnWidth = 15
    $sheet.Cells.Item($header['Case_Owner_Name__c']).ColumnWidth = 15
    # $sheet.Cells.Item($header['Category__c']).ColumnWidth = 25
    $sheet.Cells.Item($header['Subject']).ColumnWidth = 100

}

function Get-QueryResult {
    param($Query)
    do {
        $ts = (Get-Date).ToString("yyyy-MM-dd-HH:mm:ss")
        Write-Host "Query starts at $ts ==" -ForegroundColor Yellow
        $json_result = (sfdx force:data:soql:query -q $Query -u vscodeOrg --json)
    } While (($json_result -eq $null) -or ($json_result -eq $false))

    return ($json_result | ConvertFrom-Json).result.records
}

function MessageTo-Slack {
    param($Channel, $Message)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-RestMethod -Method POST -ContentType "application/json" -uri $Channel -Body $Message | Out-Null
}

function Start-WeekendLoop {
    param($InFile)

    $timeout = New-TimeSpan -Hours 12

    while ($true) {
        Write-Host "A new Excel object instaniated!"
        $excel = New-Object -ComObject Excel.Application
        $excel.AutoRecover.Enabled = $false
        $excel.DisplayAlerts = $false
        $excel.Application.AutoRecover.Enabled = $false
        $excel.Application.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($InFile)

        # $sheet_all = $workbook.sheets.item("Weekend All")
        $sheet_ugt = $workbook.sheets.item("Action Now!")
        $sheet_uass = $workbook.sheets.item("Unassigned")
        $sheet_ass = $workbook.sheets.item("Assigned(APAC)")

        $cur_array = @()
        $no_response_old = @()
        
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        
        while ($sw.Elapsed -lt $timeout) {
            # Get new list - Get-All has it' own mechanism to infinitive re-try in error conditions. So, $query_success is useless here...
            $all = Get-QueryResult -Query $query

            $new_array = $all | Select-Object -Property CaseNumber, Escalated_Case__c, Priority, First_Response_Complete__c,Product__c, Case_Owner_Name__c, Preferred_Support_Region__c
            
            # Compare the two array - new and current if any changes
            $diff =  Compare-Object -ReferenceObject $cur_array -DifferenceObject $new_array -Property CaseNumber, Priority, First_Response_Complete__c, Product__c, Case_Owner_Name__c, Preferred_Support_Region__c -PassThru -ErrorAction SilentlyContinue

            # $diff -eq $null means that there is no difference in $cur_array & $new_array, the changes in terms of number of rows and contents of the columns (CaseNumber, Priority, FR, Owner) value changes.
            if ( ($diff -ne $null) ) {
                Write-Host "New changes detected..." -ForegroundColor Magenta
                $diff | %{$_}

                ## Excel Part ##
                # Filter to create lists #
                $action_now = $all | ? {($_.Case_Owner_Name__c -eq $null) -and ($_.First_Response_Complete__c -eq $false)}
                $unassigned = $all | ? {($_.Case_Owner_Name__c -eq $null)}
                $assigned = $all | ? {($_.Preferred_Support_Region__c -eq "APAC") -and ($_.Histories.records -ne $null)}

                # Update sheets
                try {
                    Write-Sheet -Sheet $sheet_ugt -List $action_now
                }
                catch {
                    $body = ConvertTo-Json @{
                        text="Error in Weekend Support Sheet - Action Now"
                    }
                    MessageTo-Slack -Channel $ts_apac_sydney_py -Message $body
                }
                Write-Sheet -Sheet $sheet_ass -List $assigned
                try {
                    Write-Sheet -Sheet $sheet_uass -List $unassigned
                } catch {
                    $body = ConvertTo-Json @{
                        text="Error in Weekend Support Sheet - Unassigned"
                    }
                    MessageTo-Slack -Channel $ts_apac_sydney_py -Message $body
                    # $timeout.Hours = 12
                    # $workbook.Close($false)
                    # $excel.Quit()
                }
                
                
                # Console Output
                # $all | ft CaseNumber, Priority, @{L = 'Age'; E = { $_.Case_Age__c } }, Status, @{L = 'Language'; E = { $_.Preferred_Case_Language__c } }, Tier__c, @{L = 'Category'; E = { $_.Category__c } }, Product__c, Subject -AutoSize | Out-String
                Write-Host "Total Queue Size: " $new_array.Count -ForegroundColor Yellow
                $datetime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Write-Host "Last update: $datetime"

                $cur_array = $new_array
            } else {
                Write-Host "No change detected..." -ForegroundColor Green
            }            
            
            # Update Last Checking Out Time
            $pingtime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'AUS Eastern Standard Time').ToString("yyyy-MM-dd hh:mm:ss")
            foreach ($asheet in $workbook.sheets) {
                $asheet.Cells.Item(1, 12) = "Check-In $pingtime"
            }
            # $sheet_ugt.Cells.Item(1, 12) = "Check-In $pingtime"
            # $sheet_uass.Cells.Item(1, 12) = "Keepalive at $pingtime"

            # $workbook.Save()
            # Write-Host "File Saved..." -ForegroundColor Green

            Write-Host "Time Elapsed" -ForegroundColor DarkMagenta
            Write-Host $sw.Elapsed -ForegroundColor Magenta
            Start-Sleep -Seconds 10
        }
        Write-Host "Times Up! Kill Excel Process" -ForegroundColor Red
        $excel.DisplayAlerts = $false
        $workbook.Save()
        Write-Host "File Saved before Close" -ForegroundColor DarkGreen
        $workbook.Close($false)
        $excel.Quit()
        Get-Process | ? { $_.ProcessName -match "Excel" }
        Get-Process | ? { $_.ProcessName -match "Excel" } | Stop-Process
    }
}


<# Process Block#>

$Title = "Weekend Angel"
$host.UI.RawUI.WindowTitle = $Title

if (!(Test-Path $sp_excel_path -PathType Leaf)) {
    Write-Host "No Excel File exists" -ForegroundColor Red
    Create-ExcelFile -OutFile $sp_excel_path
}

$message = ConvertTo-Json -Depth 10 @{blocks=@(@{type="section";text=@{type="mrkdwn";text=":ariana: Hello ! \n $sp_link"}})}
$message = '{
    "blocks": [
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": ":ariana: Hello Team!\n Click for the Weekend Sheet"
            },
            "accessory": {
                "type": "button",
                "text": {
                    "type": "plain_text",
                    "text": "Weekend Sheet"
                },
                "style":"primary",
                "url": "'+$sp_link+'"
            }
        }
    ]
}'
MessageTo-Slack -Channel $ts_apac_sydney_py -Message $message

Start $sp_link

Start-WeekendLoop -InFile $sp_excel_path