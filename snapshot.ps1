[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
[System.Windows.Forms.Sendkeys]::SendWait("{PRTSC}")
    Start-Sleep -Seconds 2
    C:\Windows\System32\mspaint.exe
    Start-Sleep -seconds 6
    [System.Windows.Forms.Sendkeys]::SendWait("^v")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.Sendkeys]::SendWait("^s")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.Sendkeys]::SendWait("^s")
    Start-Sleep -Seconds 1
    $Request_Report_PaintFile_Name= "whatever the name"
    [System.Windows.Forms.Sendkeys]::SendWait($Request_Report_PaintFile_Name)
    Start-Sleep -Seconds 1  
    [System.Windows.Forms.Sendkeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
   
    [System.Windows.Forms.Sendkeys]::SendWait("j")
    Start-Sleep -Seconds 2
   
    [System.Windows.Forms.Sendkeys]::SendWait("{ENTER}")
    Start-Sleep -Seconds 2
