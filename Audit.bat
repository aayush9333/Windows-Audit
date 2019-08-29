ECHO OFF
CLS
:MENU
ECHO.
ECHO ..............................................
ECHO WINDOWS SERVER AUDIT COMMANDS
ECHO ..............................................
ECHO.
ECHO 1 - START WINDOWS SERVER AUDIT
ECHO 2 - EXIT
ECHO.
SET /P M=Type 1 or 2, then press ENTER: 
IF %M%==1 GOTO server
IF %M%==2 GOTO EOF


:server
cls
ECHO.
echo 			Processing Audit Commands List 
ECHO.
echo 1. Exporting GP Result
gpresult /Z > gpresult.txt
echo ---------------------------------
echo 2. Exporting Local Security Policies
secedit /export /cfg Policy.inf 
echo ---------------------------------
echo 3. Exporting Time Zone Settings
w32tm /tz > Timezone.txt
echo ---------------------------------
echo 4. Exporting Time Settings
time /t > Time.txt
echo ---------------------------------
echo 5. Exporting Account Settings
net accounts > AccountPolicy.txt
echo ---------------------------------
echo 6. Exporting Running Services
net start > services.txt
echo ---------------------------------
echo 7. Performing Port Scan
netstat -ano > ports.txt
echo ---------------------------------
echo 8. Exporting System Logs
wevtutil epl System System.evtx
echo ---------------------------------
echo 9. Exporting Application Logs
wevtutil epl Application Application.evtx
echo ---------------------------------
echo 10. Exporting Security Logs
wevtutil epl Security Security.evtx
echo ---------------------------------
echo 11. Retrieving Maximum Log Size for System, Application and Security Events Log
wevtutil gl System /f:xml > SystemLogMaxSize.txt
wevtutil gl Application /f:xml > ApplicationLogMaxSize.txt
wevtutil gl Security /f:xml > SecurityLogMaxSize.txt
echo ---------------------------------------------
echo 12. Exporting Antivirus signature information
dir "C:\Program Files (x86)\Trend Micro\OfficeScan Client"\lpt$vpn.* | find "lpt$vpn" > trendx86.txt
dir "C:\Program Files\Trend Micro\OfficeScan Client"\lpt$vpn.* | find "lpt$vpn" > trendx64.txt
echo -----------------------------------
echo 13. Exporting Free Disk Information
cscript FreeDiskSpace.vbs
echo ---------------------------------
echo 14. Checking for Installed Updates
cscript GetUpdates.vbs
echo ---------------------------------
echo 15. Exporting Task List
tasklist > TaskList.txt
echo ---------------------------------
echo 16. Exporting System Information
systeminfo > Systeminfo.txt
echo ------------------------------------------------------
echo 17. Running MBSA Command under current user privileges
echo Updating offline catalog file
xcopy "wsusscn2.cab" "C:\Program Files\Microsoft Baseline Security Analyzer 2\" /Y
Cd \Program Files\Microsoft Baseline Security Analyzer 2
MBSACLI /target localhost /qp /catalog "C:\Program Files\Microsoft Baseline Security Analyzer 2\wsusscn2.cab"
echo -------------------------------------------
echo  MBSA scan completed. Copy the report file.
echo --------------------------------
echo "** Audit Commands Completed **
pause
GOTO MENU