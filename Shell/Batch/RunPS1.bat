::실행되는 .bat파일의 같은 경로에 같은 이름으로 존재하는 ps1파일을 실행시켜줍니다.
::.bat파일의 이름이 RunPs1.bat이면, 같은 경로에 존재하는 RunPs1.ps1 파일을 실행 시켜줍니다.
::powershell.exe "~/RunPs1.ps1"
powershell.exe "%~dp0%~n0.ps1"
