@echo off
echo Retrieving saved Wi-Fi profiles and passwords...
for /f "skip=9 tokens=1,2 delims=:" %%i in ('netsh wlan show profiles') do (
    for /f "delims=:" %%j in ("%%i") do (
        echo SSID: %%j >> wifipassword.txt
        netsh wlan show profiles %%j key=clear | findstr "Key Content" >> wifipassword.txt
        echo. >> wifipassword.txt
    )
)
echo Done! Passwords saved in wifipassword.txt
pause
