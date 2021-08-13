Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$font_regular = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$font_serif = New-Object System.Drawing.Font("Microsoft Sans Serif", 11,[System.Drawing.FontStyle]::Regular)
$filename = "APACTechSupRealTimeDashboard.xlsx"
$filepath = "C:\Users\jchoi\Tableau Software Inc\APAC Tech Support - Documents\$filename"

function PlayNow {
    Write-Host "Restart Angel"
    start powershell.exe C:\MyProjects\ps\Angel\Restart-Angel.ps1
}

function Stop-RealTimeAngel {
    Write-Host "Stop Angel!"
    Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "RealTime Angel" } | Stop-Process -Force
    Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Angel Monitor" } | Stop-Process -Force
    Get-Process excel -ea 0 | Stop-Process -Force
}

function Remove-RealTimeAngelSheets {
    
    Get-Process excel -ea 0 | Stop-Process -Force
    Remove-Item $filepath
    Write-Host "$filename has been removed"
}

function Remove-Cache {
    $user = $env:USERNAME
    Remove-Item -Path "C:\Users\$user\AppData\Local\Microsoft\Office\16.0\OfficeFileCache11\*" -Recurse -Force
}

function Start-OldCaseAngel {
    Write-Host "Run Old Case Angel"
    start powershell.exe C:\MyProjects\ps\Angel\Get-OldCases2.ps1
}

function Start-WeekendAngel {
    Write-Host "Run Weekend Angel"
    start powershell.exe C:\MyProjects\ps\Angel\Restart-WkdAngel.ps1
}


function Stop-WeekendAngel {
    Write-Host "Stop Weekend Angel!"
    Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Weekend Angel" } | Stop-Process -Force
    Get-Process excel -ea 0 | Stop-Process -Force
}

function Stop-CrashPlan {
    Get-Service | ? { $_.DisplayName -match "Code42 CrashPlan Service" } | Stop-Service -Force
}

function Start-CrashPlan {
    Get-Service | ? { $_.DisplayName -match "Code42 CrashPlan Service" } | Start-Service -Force
}


$MainWindow = New-Object System.Windows.Forms.Form
$MainWindow.ClientSize = "850, 320"
$MainWindow.Text ="RealTime Angel Control Center"
$MainWindow.BackColor = "White"


$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Click to get old list"
$Label.AutoSize = $true
$Label.Width = 25
$Label.Height = 10
$Label.Location = New-Object System.Drawing.Point(20, 20)
$Label.Font = "Arial, 13"


$PlayBtn = New-Object System.Windows.Forms.Button
$PlayBtn.BackColor = "Green"
$PlayBtn.Text = "Restart Realtime Angel"
$PlayBtn.Width = 200
$PlayBtn.Height = 80
$PlayBtn.Location = New-Object System.Drawing.Point(10, 50)
$PlayBtn.Font = "Microsoft Sans Serif, 15"
$PlayBtn.ForeColor = "Red"
$PlayBtn.Add_Click({PlayNow})

$OldBtn = New-Object System.Windows.Forms.Button
$OldBtn.BackColor = "Red"
$OldBtn.ForeColor = "White"
$OldBtn.Text = "Stop Realtime Angel"
$OldBtn.Width = 200
$OldBtn.Height = 80
$OldBtn.Location = New-Object System.Drawing.Point(220, 50)
$OldBtn.Font = $font_regular
$OldBtn.Add_Click({Stop-RealTimeAngel})


$BtnRmAngel = New-Object System.Windows.Forms.Button
$BtnRmAngel.BackColor = "DarkRed"
$BtnRmAngel.ForeColor = "White"
$BtnRmAngel.Text = "Remove Realtime Angel Sheets"
$BtnRmAngel.Width = 200
$BtnRmAngel.Height = 80
$BtnRmAngel.Location = New-Object System.Drawing.Point(430, 50)
$BtnRmAngel.Font = $font_regular
$BtnRmAngel.Add_Click({Remove-RealTimeAngelSheets})


$BtnRmCache = New-Object System.Windows.Forms.Button
$BtnRmCache.BackColor = "Black"
$BtnRmCache.ForeColor = "White"
$BtnRmCache.Text = "Remove Office Cache Files"
$BtnRmCache.Width = 200
$BtnRmCache.Height = 80
$BtnRmCache.Location = New-Object System.Drawing.Point(640, 50)
$BtnRmCache.Font = $font_regular
$BtnRmCache.Add_Click({Remove-Cache})


$BtnRunOld = New-Object System.Windows.Forms.Button
$BtnRunOld.BackColor = "Magenta"
$BtnRunOld.ForeColor = "White"
$BtnRunOld.Text = "Run Old Case Angel"
$BtnRunOld.Width = 200
$BtnRunOld.Height = 80
$BtnRunOld.Location = New-Object System.Drawing.Point(10, 140)
$BtnRunOld.Font = $font_serif
$BtnRunOld.Add_Click({Start-OldCaseAngel})

$BtnRunWkd = New-Object System.Windows.Forms.Button
$BtnRunWkd.BackColor = "Blue"
$BtnRunWkd.ForeColor = "White"
$BtnRunWkd.Text = "Run Weekend Angel"
$BtnRunWkd.Width = 200
$BtnRunWkd.Height = 80
$BtnRunWkd.Location = New-Object System.Drawing.Point(10, 230)
$BtnRunWkd.Font = $font_serif
$BtnRunWkd.Add_Click({Start-WeekendAngel})

$BtnStopWkd = New-Object System.Windows.Forms.Button
$BtnStopWkd.BackColor = "Red"
$BtnStopWkd.ForeColor = "White"
$BtnStopWkd.Text = "Stop Weekend Angel"
$BtnStopWkd.Width = 200
$BtnStopWkd.Height = 80
$BtnStopWkd.Location = New-Object System.Drawing.Point(220, 230)
$BtnStopWkd.Font = $font_serif
$BtnStopWkd.Add_Click({Stop-WeekendAngel})

$MainWindow.Controls.Add($Label)
$MainWindow.Controls.Add($PlayBtn)
$MainWindow.Controls.Add($OldBtn)
$MainWindow.Controls.Add($BtnRunOld)
$MainWindow.Controls.Add($BtnRunWkd)
$MainWindow.Controls.Add($BtnStopWkd)
$MainWindow.Controls.Add($BtnRmAngel)
$MainWindow.Controls.Add($BtnRmCache)

[void]$MainWindow.ShowDialog()