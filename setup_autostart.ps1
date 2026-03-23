# 創建 Windows 任務計劃程序來實現開機自動啟動 plc2google 容器
# 以管理員身份運行此腳本

$taskName = "Plc2GoogleAutoStart"
$scriptPath = "d:\掌門事業股份有限公司\汐止廠 - 文件\釀造生產部\SourceCode\Other\plc2google\start_plc2google.bat"

# 檢查任務是否已存在
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($task) {
    Write-Host "任務 '$taskName' 已存在，正在更新..."
} else {
    Write-Host "創建新任務 '$taskName'..."
}

# 定義觸發器（開機時啟動）
$trigger = New-ScheduledTaskTrigger -AtStartup

# 定義操作（執行批處理腳本）
$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$scriptPath`""

# 定義設置
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

# 定義用戶帳戶（使用 SYSTEM 帳戶以獲得最高權限）
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# 創建或更新任務
try {
    Register-ScheduledTask `
        -TaskName $taskName `
        -Trigger $trigger `
        -Action $action `
        -Settings $settings `
        -Principal $principal `
        -Description "自動啟動 plc2google Docker 容器" `
        -ErrorAction Stop
    
    Write-Host "任務 '$taskName' 創建/更新成功！"
    Write-Host ""
    Write-Host "任務詳情："
    Get-ScheduledTask -TaskName $taskName | Format-Table -AutoSize
    
    Write-Host ""
    Write-Host "提示：您可以使用以下命令管理此任務："
    Write-Host "  - 查看任務狀態：Get-ScheduledTask -TaskName $taskName"
    Write-Host "  - 手動啟動任務：Start-ScheduledTask -TaskName $taskName"
    Write-Host "  - 刪除任務：Unregister-ScheduledTask -TaskName $taskName -Confirm`$false"
} catch {
    Write-Host "錯誤：$_"
    Write-Host "請確保以管理員身份運行此腳本！"
    exit 1
}
