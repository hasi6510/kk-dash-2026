# 家計ダッシュボード - Windowsタスクスケジューラ 設定スクリプト
# 実行方法: PowerShellを管理者として実行し、このスクリプトを実行してください

$taskName    = "家計ダッシュボード 更新リマインダー"
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$batFile     = Join-Path $scriptDir "run_silent.bat"

# 既存タスクを削除（再設定用）
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "既存タスクを削除しました。" -ForegroundColor Yellow
}

# アクション定義: run_silent.bat を実行
$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$batFile`"" `
    -WorkingDirectory $scriptDir

# トリガー: 毎月10日 08:00
$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 10 -At "08:00"

# 設定: ログオン時のみ実行（パスワード不要）
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable

# タスク登録
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force | Out-Null

Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "  タスクスケジューラの設定が完了しました！" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Write-Host "タスク名: $taskName"
Write-Host "実行日時: 毎月10日 08:00"
Write-Host "実行ファイル: $batFile"
Write-Host ""
Write-Host "毎月の手順:"
Write-Host "  1. MoneyForward ME からCSVをダウンロード"
Write-Host "  2. mf_csv フォルダに保存"
Write-Host "  3. 毎月10日 08:00 に自動でダッシュボードが更新されます"
Write-Host ""
Write-Host "確認方法: タスクスケジューラ を開いて「$taskName」を探してください。"
Write-Host ""
