Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "RealTime Angel" } | Stop-Process -Force
Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Angel Monitor" } | Stop-Process -Force
Get-Process excel -ea 0 | Stop-Process -Force