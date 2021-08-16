## Console output encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Include Functions
. $PSScriptRoot\RealTime-Func.ps1
. $PSScriptRoot\Angel-Query.ps1

## Global variables ##
$user =             $env:USERNAME
$sp_path =          "C:\Users\$user\Tableau Software Inc\APAC Tech Support - Documents" # SharePoint Local folder path
$file_name =        "APACTechSupRealTimeDashboard.xlsx"
$sfile =            "$sp_path\$file_name" # SharePoint local Path 
$dwfolder =         "C:\Downloads"
$dwfile =           "C:\Downloads\$file_name"
$kbfolder =         "C:\Users\$user\Documents"
$bkfile =           "C:\Users\$user\Documents\$file_name"
$officeCacheRoot =  "C:\Users\$user\AppData\Local\Microsoft\Office\16.0"
$CacheLocation =    "$officeCacheRoot\OfficeFileCache"
$angel_pid_file =   "C:\MyProjects\PS\Angel\angel.pid"
$excel_pid_file =   "C:\MyProjects\PS\Angel\excel.pid"

## SharePoint (Online) ## 
$sp_root =   "https://tableau.sharepoint.com/sites/APACTechSupport/Shared%20Documents"
$sp_file =   "$sp_root/$file_name"
$web_lnk =   $sp_file + "?web=1"

## Slack ##
#ts-apac-support
# $uri =     "https://hooks.slack.com/services/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"
#ts-apac-sydney-py
$uri =       "https://hooks.slack.com/services/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"

## sfdx commands ##
# sfdx force:auth:web:login -a vscodeOrg
# sfdx force:org:list --all
# sfdx force:data:soql:query -u vscodeOrg -q "select id from user limit 5"

function update_sheet {
    [CmdletBinding()]
    param($sheet, $list)

    $sheet.Unprotect()

    $hyperroot = "https://tableau.my.salesforce.com/"

    $datetime = Get-AESTDate
    Write-Host $sheet.Name "is being updated at $datetime" -ForegroundColor Cyan

    # Clear all contents 
    $sheet.Cells.ClearContents() | Out-Null
    $sheet.Cells.ClearFormats() | Out-Null
    # $sheet.UsedRange.ClearContents()

    $sheet.Cells.Item(1, 13) = "Updated: $datetime"

    $q_list = ($list | ? { $_.Case_Owner_Name__c -eq $null })
    $a_list = ($list | ? { $_.Case_Owner_Name__c -ne $null })
    
    ## Unassigned Count
    $sheet.Cells.Item(1, 3) = "Unassigned: $($q_list.Count)"
    $sheet.Cells.Item(1, 3).Font.Bold = $true
    ## Assigned Count
    $sheet.Cells.Item(1, 5) = "Assigned: $($a_list.Count)"
    ## Total Count
    $sheet.Cells.Item(1, 7) = "Total: $($list.Count)"
    $sheet.Cells.Item(1, 7).Interior.ColorIndex = 6

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
        'CreateDate'                   = 17
    }

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
            $sheet.Cells.Item($i, $header['Case_Owner_Name__c']) = $row.Case_Owner_Name__c
        }

        # Description
        if (($row.Description -eq $null) -or ($row.Description -eq '')) {
            Write-Warning "Description is NULL"
            $row.Description = $row.CaseNumber
        }

        # Case Number
        Write-Host $row.CaseNumber
        $sheet.Hyperlinks.Add(
            $sheet.Cells.Item($i, $header['CaseNumber']),
            $hyperroot+$row.Id,
            "",
            $hyperroot+$row.Id,
            # $row.Description,
            $row.CaseNumber
        ) | Out-Null

        # Escalated
        if ($row.Escalated_Case__c -eq $true) {
            $sheet.Cells.Item($i, $header['Escalated_Case__c']) = "YES"
            $sheet.Cells.Item($i, $header['Escalated_Case__c']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['Escalated_Case__c']).Font.ColorIndex = 3
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
        if ( ($row.Entitlement_Type__c -match "Premium") -or ($row.Entitlement_Type__c -eq "Elite") -or ($row.Entitlement_Type__c -eq "Extended") ) {
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

        } elseif ($row.First_Response_Complete__c -eq $false) {

            # First Response Highlighter
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
        } elseif (($row.Preferred_Case_Language__c -eq "Mandarin") -or ($row.Preferred_Case_Language__c -eq "Chinese Traditional") -or ($row.Preferred_Case_Language__c -eq "Chinese Simplified")) {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = ($row.Preferred_Case_Language__c).Split(" ")[0]
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']).Font.ColorIndex = 45
        } else {
            $sheet.Cells.Item($i, $header['Preferred_Case_Language__c']) = ($row.Preferred_Case_Language__c).Split(" ")[0]
        }

        # Timezone
        if ($row.Case_Preferred_Timezone__c) {
            $sheet.Cells.Item($i, $header['Case_Preferred_Timezone__c']) = ($row.Case_Preferred_Timezone__c).split(" ")[0]
        }

        # Products
        $sheet.Cells.Item($i, $header['Product__c']) = ($row.Product__c).Split(" ")[1]
        if (($row.Product__c -eq "Tableau Desktop") -or ($row.Product__c -eq "Tableau Public Desktop") -or ($row.Product__c -eq "Tableau Reader") -or ($row.Product__c -eq "Tableau Prep") -or ($row.Product__c -eq "Tableau Public Desktop") ) {
            # $sheet.Cells.Item($i, $header['Product__c']).Interior.ColorIndex = 35
            # # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.ThemeColor = xlThemeColorAccent3
            # # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.TintAndShade = 0.6
        } else {
            # $sheet.Cells.Item($i, $header['Product__c']).Interior.ColorIndex = 37
            # # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.ThemeColor = xlThemeColorAccent3
            # # $sheet.Cells.Item($i, $header['Product__c']).Interior.Interior.TintAndShade = 0.6
        }
        
        # Category
        $sheet.Cells.Item($i, $header['Category__c']) = $row.Category__c
        
        # Subject
        $sheet.Cells.Item($i, $header['Subject']) = $row.Subject

        # Accounts
        if (($row.Account.Name -eq "Westpac Banking Corporation") -or ($row.Account.Name -eq "GIC Private Limited") -or ($row.Account.Name -eq "Applied Materials, Inc.")) {
            $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name
            $sheet.Cells.Item($i, $header['Account.Name']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['Account.Name']).Interior.ColorIndex = 6
        } elseif($row.Account.Name -eq "Hang Seng Bank Limited") {
            $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name
            $sheet.Cells.Item($i, $header['Account.Name']).Font.Bold = $true
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']) = "Premium"
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Interior.ColorIndex = 3
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.ColorIndex = 2
            $sheet.Cells.Item($i, $header['Entitlement_Type__c']).Font.Bold = $true
        } else {
            $sheet.Cells.Item($i, $header['Account.Name']) = $row.Account.Name
        }
        
        # Protips Type 1
        # $sheet.Cells.Item($i, $header['Protip']) = $row.Feeds.records
        # Protips Type 2
        if ( ($row.Plan_of_Action_Status__c -ne $null) -and ($row.Plan_of_Action_Status__c -like "*protip*") ) {
            $sheet.Cells.Item($i, $header['Protip']) = "YES"
        }
        
        # CSM
        $sheet.Cells.Item($i, $header['Account.CSM_Name__c']) = $row.CSM
        $sheet.Cells.Item($i, $header['Account.CSM_Name__c']).HorizontalAlignment = -4152

        # Hand Off
        if ( ($row.Plan_of_Action_Status__c -ne $null) -and ($row.Plan_of_Action_Status__c -like "*apac*") ) {
            # $sheet.Cells.Item($i, $header['HandOff']) = $row.Plan_of_Action_Status__c
            $sheet.Cells.Item($i, $header['HandOff']) = "To APAC"
        }
        
        # Hide cases if fresh P3 and P4
        # Write-Host $sheet.Name
        <#
        if ( ($sheet.Name -match "Server" -or $sheet.Name -match "All") -and ($row.Entitlement_Type__c -match "Standard") -and ($row.Priority -eq "P3" -or $row.Priority -eq "P4") -and ($row.Case_Age__c -lt 63) -and ($row.Case_Owner_Name__c -eq $null) ) {
            $sheet.Rows($i).Hidden = $true
        }
        #>

        <# New filter requested by Emma #>
        <## show only p1, p2, premium and escalation only ##>
        <#
        if ( ($sheet.Name -match "Server" -or $sheet.Name -match "All" -or $sheet.Name -match "Aged P3P4" -or $sheet.Name -match "Chinese" -or $sheet.Name -match "Unassigned" -or $sheet.Name -match "India") -and ($row.Entitlement_Type__c -match "Standard") -and (($row.Escalated_Case__c -eq $false) -or ($row.Escalated_Case__c -eq $null)) -and ($row.Priority -eq "P3" -or $row.Priority -eq "P4") ) {
            $sheet.Rows($i).Hidden = $true
        }

        if ($row.Case_Owner_Name__c -ne $null) {
            $sheet.Rows($i).Hidden = $false
        }

        if ($row.Escalated_Case__c -eq $true) {
            $sheet.Rows($i).Hidden = $false
        }
        #>

        $i++
    }

    # Column Autofit
    $sheet.columns.AutoFit() | Out-Null
    $sheet.Cells.Item($header['Case_Owner_Name__c']).ColumnWidth = 13
    $sheet.Cells.Item($header['Category__c']).ColumnWidth = 25
    $sheet.Cells.Item($header['Subject']).ColumnWidth = 100
    $sheet.Protect('',0,1,0,0,1,0,1,0,0,1,0,1,0,1,1)

    Write-Host $sheet.Name " update completed" -ForegroundColor DarkCyan
}

function Run-MainLoop {
    param($InFile)

    $hyperroot = "https://tableau.my.salesforce.com/"

    # Timeout for restarting Excel
    $timeout = New-TimeSpan -Hours 2
    # $timeout = New-TimeSpan -Minutes 1

    while ($true) {
        Write-Host "A new Excel object instaniated!"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Open($InFile)

        Get-Process excel | select MainWindowTitle, Id, StartTime

        $angel_pid = (Get-Process excel –ea 0 | Where-Object { $_.MainWindowTitle –like "*$file_name*" }).Id
        $angel_pid | Set-Content -Path $angel_pid_file
        $angel_pid_file_pid = Get-Content -Path $angel_pid_file
        Write-Host "Angel PID"      $angel_pid
        Write-Host "Angel PID File" $angel_pid_file_pid

        $excel_pid = (Get-Process excel –ea 0).Id
        $excel_pid | Set-Content -Path $excel_pid_file
        $excel_pid_file_pid = Get-Content -Path $excel_pid_file
        Write-Host "Excel PID"      $excel_pid
        Write-Host "Excel PID File" $excel_pid_file_pid

        $excel.Visible = $false
        $excel.AutoRecover.Enabled = $false
        $excel.DisplayAlerts = $false
        

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
        
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        
        while ($sw.Elapsed -lt $timeout) {
            $all = get-all -Query $query | Sort-Object Case_Owner_Name__c, @{Expression="Entitlement_Type__c";Ascending=$true}, Priority,@{Expression="CSM";Descending=$true},@{Expression="Case_Age__c";Descending=$true}
            
            $all = filter-pizza -InputList $all

            $all | Select-Object -Property @{Name="Engineer";Expression="Case_Owner_Name__c"}, `
                                    @{Name="Protip";Expression="Feeds.records"}, `
                                    CaseNumber, `
                                    @{Name="Escl";Expression="Escalated_Case__c"}, `
                                    CSM, `
                                    @{Name="P";Expression="Priority"}, `
                                    @{Name="Age";Expression="Case_Age__c"}, `
                                    @{Name="Entitlement";Expression="Entitlement_Type__c"}, `
                                    @{Name="F/R";E="First_Response_Complete__c"},`
                                    @{N="Language";E="Preferred_Case_Language__c"}, `
                                    @{N="Timezone";E="Case_Preferred_Timezone__c"}, `
                                    @{N="Product";E="Product__c"}, `
                                    @{N="Category";E="Category__c"}, `
                                    Subject, `
                                    @{n="Account";e={$_.Account | Select-Object -expandproperty Name}}, `
                                    @{n="link";e={$hyperroot+$_.Id}} `
                                    | Export-Csv -Path C:\MyProjects\PS\Angel\all.csv -Force -Encoding UTF8 -NoTypeInformation

            $new_array = $all | Select-Object -Property CaseNumber, Escalated_Case__c, Priority, First_Response_Complete__c,Product__c, Case_Owner_Name__c, Feeds.records
            
            # Compare the two array - new and current if any changes
            $diff =  Compare-Object -ReferenceObject $cur_array -DifferenceObject $new_array -Property CaseNumber, Priority, First_Response_Complete__c, Product__c, Case_Owner_Name__c, Feeds.records -PassThru -ErrorAction SilentlyContinue
            
            # $diff -eq $null means that there is no difference in $cur_array & $new_array, the changes in terms of number of rows and contents of the columns (CaseNumber, Priority, FR, Owner) value changes.
            if ( ($diff -ne $null) ) {
                Write-Host "New changes detected..." -ForegroundColor Red
                # display what's different
                # $diff | %{ Write-Host $_ }

                # Filter the list
                $desktop =  filter-desktop -In $all
                $server =   filter-server -In $all
                $premium =  filter-premium -In $all
                # $urgent = filter-urgent -In $all
                $ass =      Filter-Assigned -InputList $all | Sort-Object Priority, Case_Owner_Name__c
                $uss =      Filter-Unassigned -InputList $all
                $ko =       Filter-Language -Language "KO" -InputList $all
                $cn =       Filter-Language -Language "CN" -InputList $all
                $in =       Filter-Timezone -InputList $all
                $p1p2 =     Filter-P1P2 -InputList $all
                $p3p4 =     Filter-P3P4 -InputList $all
                # Applying old P3/P4 filter for Server Sheet - !! IMPORTANT to put $all in the last order. 
                # $all = filter-oldp3p4 -In $all

                # Update sheets                
                try {
                    update_sheet -sheet $sheet_dsk -list $desktop -ErrorAction Stop
                    update_sheet -sheet $sheet_srv -list $server -ErrorAction Continue
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
                    Write-Error "Hey! There is an error while writing sheets"
                    Write-Warning $Error[0]
                    $error_name = $Error[0].Exception.GetType().FullName
                    
                    MessageTo-Slack -ChannelUri $uri -Message "Error occurs while updating Sheets with $error_name" -Type "Warning"
                    Send-AngelNotification -Message "<b style='color:red;'>Hey! There is an error while writing worksheets</b>"

                    Copy-Item $sfile -Destination "C:\Downloads" -Recurse -Force -ErrorAction Continue

                    # $excel.Visible = $true
                    # Remove-ExcelFile -FileName $sfile
                    # $excel.DisplayAlerts = $false
                    # $workbook.SaveAs($dwfile)
                    # $workbook.Close($false)
                    # $excel.Quit()
                    # Write-Host "Stop Process!" -ForegroundColor Cyan
                    Stop-Process -Id $angel_pid
                    
                    $cntdwn = 10
                    do {
                        Start-Sleep -Seconds 1
                        Write-Host "Angel restart count down $cntdwn"
                        $cntdwn--
                    } while ($cntdwn -gt 0)

                    Remove-Item $sfile -Recurse -Force -ErrorAction Continue
                    Remove-Item -Path "C:\Users\jchoi\AppData\Local\Microsoft\Office\16.0\OfficeFileCache*" -Recurse -Force -ErrorAction Continue
                    
                    Write-Warning "Angel sheets on Sharepoint and Office cache are removed!"
                    # Start-Process powershell.exe C:\MyProjects\ps\Angel\Restart-Angel.ps1
                    
                    $cntdwn = 20
                    do {
                        Start-Sleep -Seconds 1
                        Write-Host "Angel restart count down $cntdwn"
                        $cntdwn--
                    } while ($cntdwn -gt 0)

                    Run-MainLoop -InFile $sfile  #spagetti
                }

                # Console Output
                # $all | ft @{L = 'Protip'; E = { $_.Case_Age__c } }, CaseNumber, Priority, @{L = 'Age'; E = { $_.Case_Age__c } }, Status, @{L = 'Language'; E = { $_.Preferred_Case_Language__c } }, Tier__c, @{L = 'Category'; E = { $_.Category__c } }, Product__c, Subject -AutoSize | Out-String
                $datetime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")                
                Write-Host "Total Queue Size: " $new_array.Count -ForegroundColor Yellow
                Write-Host "Last update: $datetime"

                $cur_array = $new_array
            } else {
                Write-Host "No change detected..." -ForegroundColor Green
            }

            Write-Host ">>>> Time Elapsed: " $sw.Elapsed ">>>>" -ForegroundColor Magenta
            
            # Writing Hearbeat on the sheets
            $datetime = Get-AESTDate
            foreach ($asheet in $workbook.sheets) {
                if ( ($asheet.Name -eq "URNext") -or ($asheet.Name -eq "Stat") -or ($asheet.Name -eq "Pizza")) {
                    $asheet.Unprotect()
                } else {
                    $asheet.Unprotect()
                    $asheet.Cells.Item(1, 14) = "Check-In: $datetime"
                    $asheet.Cells.Item(1, 14).HorizontalAlignment = -4131
                    $asheet.Protect('',0,1,0,0,1,0,1,0,0,1,0,1,0,1,1)
                }
            }

            Write-Host "Ping Time: $datetime" -ForegroundColor Gray

            ## The RPC server is unavailable. (Exception from HRESULT: 0x800706BA) occurs sometimes so disable the options below.
            try {
                # $excel.DisplayAlerts=$false
                $workbook.Save()
            } catch {
                
            }
            
            Copy-Item $sfile -Destination $kbfolder -Force
        }

        # Periodic shutdown
        Write-Host "Times Up! Killing Excel Process" -ForegroundColor Red
        MessageTo-Slack -ChannelUri $uri -Message "Angel Times up"
        Send-AngelNotification -Message 'Times Up! Restarting MainLoop'

        # Save current sheets
        # $excel.DisplayAlerts = $false
        $workbook.SaveAs($dwfile)
        $workbook.Close($false)
        $excel.Quit()

        # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        # Remove-Variable $excel
        [GC]::Collect()

        # Get-Process | ? { $_.ProcessName -match "Excel" }
        # Get-Process | ? { $_.ProcessName -match "Excel" } | Stop-Process
        Get-Process | ? {$_.Id -eq $angel_pid} | Stop-Process
        # Start-Process powershell.exe C:\MyProjects\ps\Angel\Restart-Angel.ps1
        Run-MainLoop -InFile $sfile
    }
}

## Main
$Title = "RealTime Angel"
$host.UI.RawUI.WindowTitle = $Title

# Remove-Item -Path $CacheLocation -Recurse -Force
Start $officeCacheRoot
Start $sp_path
Start $PSScriptRoot

if (!(Test-Path $sfile -PathType Leaf)) {
    # Create-ExcelFile -OutFile $sfile
    Copy-Item $dwfile -Destination $sp_path
}

if (!(Test-Path $angel_pid_file -PathType Leaf)) {
    New-Item $angel_pid_file -ItemType file
}

if (!(Test-Path $excel_pid_file -PathType Leaf)) {
    New-Item $excel_pid_file -ItemType file
}

$Text = Slack-Mrkdwn -Text "@here Angel is Starting..."
MessageTo-Slack -ChannelUri $uri -Message "Starting Realtime Angel at $web_lnk"
Send-AngelNotification -Message "<h1>Angel is Starting</h1>"

Run-MainLoop -InFile $sfile