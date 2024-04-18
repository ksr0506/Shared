SET NICName=Wi-Fi
SET IP=10.225.123.123
SET SUBNETMASK=255.255.255.0
SET GATEWAY=10.225.123.1
SET DNS1=147.6.44.44
SET DNS2=147.6.44.45

netsh -c int ip set address name=%NICName% source=static addr=%IP% mask=%SUBNETMASK% gateway=%GATEWAY% gwmetric=0
netsh -c int ip set dnsservers name=%NICName% source=static addr=%DNS1% register=primary no
netsh -c int ip add dnsservers name=%NICName% addr=%DNS2% index=2 no