(for /f "skip=9 tokens=1,* delims=:" %i in ('netsh wlan show profiles') do @echo SSID:%j & netsh wlan show profile name="%j" key=clear | findstr "Key Content" & echo.) > wifipassword.txt
