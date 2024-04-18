# PSWindowsUpdate ����� ��ġ�Ǿ� �ִ��� Ȯ���ϰ�, ���ٸ� ��ġ
if (-not(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Output "PSWindowsUpdate ����� ��ġ�Ǿ� ���� �ʽ��ϴ�. ��ġ�� �����մϴ�."
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber
}

# PSWindowsUpdate ��� ȣ��
Import-Module PSWindowsUpdate

# ��� ������ ������Ʈ ��� Ȯ��
Write-Output "Windows ������Ʈ Ȯ�� ��..."
$availableUpdates = Get-WindowsUpdate
$availableUpdatesCount = $availableUpdates.Count
Write-Output ("{0}���� ��� ������ ������Ʈ ����" -f $availableUpdatesCount)

# ��ġ�� ������Ʈ ��� ���
if ($availableUpdatesCount -gt 0) {
    Write-Output $availableUpdates | Format-Table -Property ComputerName,Status,KB,Size,Title

    # ��� ��� ������ ������Ʈ ��ġ
    Write-Output "��� ������Ʈ�� ��ġ�մϴ�."
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot
    
    # ����ڰ� ���� Ű�� ���� ������ ��ٸ�
    Read-Host "������Ʈ ������ �Ϸ�Ǿ����ϴ�. â�� �������� ���� Ű�� ��������."
}
else {
    # ����ڰ� ���� Ű�� ���� ������ ��ٸ�
    Read-Host "��ġ�� ������Ʈ�� �����ϴ�. â�� �������� ���� Ű�� ��������."
}