::실행되는 Batch File(*.bat)의 같은 경로에 같은 이름으로 존재하는 PowerShell Script(*.ps1)를 실행시켜줍니다.
::Batch File(*.bat)의 이름이 RunPs1.bat이면, 같은 경로에 존재하는 RunPs1.ps1을 실행 시켜줍니다.
::powershell.exe "~/RunPs1.ps1"
powershell.exe -ExecutionPolicy Remotesigned -File "%~dp0%~n0.ps1"
