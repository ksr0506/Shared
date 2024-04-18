SET NICName=Wi-Fi
SET IP=0.0.0.0
SET SUBNETMASK=0.0.0.0
SET GATEWAY=0.0.0.0
SET DNS1=0.0.0.0
SET DNS2=0.0.0.0

netsh -c int ip set address name=%NICName% source=static addr=%IP% mask=%SUBNETMASK% gateway=%GATEWAY% gwmetric=0
netsh -c int ip set dnsservers name=%NICName% source=static addr=%DNS1% register=primary no
netsh -c int ip add dnsservers name=%NICName% addr=%DNS2% index=2 no
