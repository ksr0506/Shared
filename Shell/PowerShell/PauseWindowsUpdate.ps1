# 레지스트리 5가지 항목의  날짜를 오늘 기준 35일 연장합니다.
$pauseFeatureUpdatesStartTime = (Get-Date); $pauseFeatureUpdatesStartTime = $pauseFeatureUpdatesStartTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseFeatureUpdatesStartTime' -Value $pauseFeatureUpdatesStartTime
$pauseFeatureUpdatesEndTime = (Get-Date).AddDays(35); $pauseFeatureUpdatesEndTime = $pauseFeatureUpdatesEndTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseFeatureUpdatesEndTime' -Value $pauseFeatureUpdatesEndTime
$pauseQualityUpdatesStartTime = (Get-Date); $pauseQualityUpdatesStartTime = $pauseQualityUpdatesStartTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseQualityUpdatesStartTime' -Value $pauseQualityUpdatesStartTime
$pauseQualityUpdatesEndTime = (Get-Date).AddDays(35); $pauseQualityUpdatesEndTime = $pauseQualityUpdatesEndTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseQualityUpdatesEndTime' -Value $pauseQualityUpdatesEndTime
$pauseUpdatesExpiryTime = (Get-Date).AddDays(35); $pauseUpdatesExpiryTime = $pauseUpdatesExpiryTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseUpdatesExpiryTime' -Value $pauseUpdatesExpiryTime

# 레지스트리에서 변경 후 PauseUpdatesExpiryTime 값을 조회하고, 해당 값을 직접 추출합니다.
$dateString = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' | Select-Object -ExpandProperty PauseUpdatesExpiryTime
$dateTime = [DateTime]::Parse($dateString)
$updatePauseDate = $dateTime.ToString("yyyy-MM-dd")

# 사용자가 엔터 키를 누를 때까지 기다림
Read-Host  ("{0}에 업데이트가 다시 시작됩니다. 창을 닫으려면 엔터 키를 누르세요." -f $updatePauseDate)