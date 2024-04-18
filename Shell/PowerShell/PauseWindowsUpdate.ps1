# ������Ʈ�� 5���� �׸���  ��¥�� ���� ���� 35�� �����մϴ�.
$pauseFeatureUpdatesStartTime = (Get-Date); $pauseFeatureUpdatesStartTime = $pauseFeatureUpdatesStartTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseFeatureUpdatesStartTime' -Value $pauseFeatureUpdatesStartTime
$pauseFeatureUpdatesEndTime = (Get-Date).AddDays(35); $pauseFeatureUpdatesEndTime = $pauseFeatureUpdatesEndTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseFeatureUpdatesEndTime' -Value $pauseFeatureUpdatesEndTime
$pauseQualityUpdatesStartTime = (Get-Date); $pauseQualityUpdatesStartTime = $pauseQualityUpdatesStartTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseQualityUpdatesStartTime' -Value $pauseQualityUpdatesStartTime
$pauseQualityUpdatesEndTime = (Get-Date).AddDays(35); $pauseQualityUpdatesEndTime = $pauseQualityUpdatesEndTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseQualityUpdatesEndTime' -Value $pauseQualityUpdatesEndTime
$pauseUpdatesExpiryTime = (Get-Date).AddDays(35); $pauseUpdatesExpiryTime = $pauseUpdatesExpiryTime.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ssZ" ); Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' -Name 'PauseUpdatesExpiryTime' -Value $pauseUpdatesExpiryTime

# ������Ʈ������ ���� �� PauseUpdatesExpiryTime ���� ��ȸ�ϰ�, �ش� ���� ���� �����մϴ�.
$dateString = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings' | Select-Object -ExpandProperty PauseUpdatesExpiryTime
$dateTime = [DateTime]::Parse($dateString)
$updatePauseDate = $dateTime.ToString("yyyy-MM-dd")

# ����ڰ� ���� Ű�� ���� ������ ��ٸ�
Read-Host  ("{0}�� ������Ʈ�� �ٽ� ���۵˴ϴ�. â�� �������� ���� Ű�� ��������." -f $updatePauseDate)