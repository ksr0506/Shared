# PSWindowsUpdate 모듈이 설치되어 있는지 확인하고, 없다면 설치
if (-not(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Output "PSWindowsUpdate 모듈이 설치되어 있지 않습니다. 설치를 시작합니다."
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber
}

# PSWindowsUpdate 모듈 호출
Import-Module PSWindowsUpdate

# 사용 가능한 업데이트 목록 확인
Write-Output "Windows 업데이트 확인 중..."
$availableUpdates = Get-WindowsUpdate
$availableUpdatesCount = $availableUpdates.Count
Write-Output ("{0}건의 사용 가능한 업데이트 있음" -f $availableUpdatesCount)

# 설치된 업데이트 목록 출력
if ($availableUpdatesCount -gt 0) {
    Write-Output $availableUpdates | Format-Table -Property ComputerName,Status,KB,Size,Title

    # 모든 사용 가능한 업데이트 설치
    Write-Output "모든 업데이트를 설치합니다."
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot
    
    # 사용자가 엔터 키를 누를 때까지 기다림
    Read-Host "업데이트 수행이 완료되었습니다. 창을 닫으려면 엔터 키를 누르세요."
}
else {
    # 사용자가 엔터 키를 누를 때까지 기다림
    Read-Host "설치할 업데이트가 없습니다. 창을 닫으려면 엔터 키를 누르세요."
}