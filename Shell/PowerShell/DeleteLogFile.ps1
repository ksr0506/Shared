$days = 30
$userProfile = [System.Environment]::GetFolderPath('UserProfile')
$uipathLogPath = "AppData\Local\UiPath\Logs\*"
$targetFolder = Join-Path $userProfile $uipathLogPath

Get-ChildItem -Path $targetFolder -File -Include "*.log","*.txt" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$days) } |
    ForEach-Object {
        Write-Output "Deleting file: $($_.FullName)"
        $_ | Remove-Item -Confirm:$false
    }

    # 사용자가 엔터 키를 누를 때까지 기다림
    Read-Host "Log파일 삭제를 완료했습니다. 창을 닫으려면 엔터 키를 누르세요."
