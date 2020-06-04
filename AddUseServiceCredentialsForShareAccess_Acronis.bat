@ECHO off

REM Run from command line with “AddUseServiceCredentialsForShareAccess_Acronis > AddUseServiceCredentialsForShareAccess_Acronis_LOG.txt 2>&1” (without quotes) so a log file is generated in the working directory.

SETLOCAL ENABLEDELAYEDEXPANSION

REM initialize counter for nodesUsingTaskkill list
SET n=0

REM initialize counter for nodesForWhichAccessIsDenied list
SET m=0

REM initialize counter for nodesWithRPCErrors list
SET o=0

REM initialize counter for nodesWithoutMMS list
SET p=0

REM initialize counter for nodesStillHanging list
SET q=0

REM initialize counter for nodesRegAddFailed list
SET r=0

FOR %%x in (PASTE COMMA SEPARATED NODE NAMES HERE) do (

	ECHO %%x

	REM add new registry key for Acronis
	REG ADD \\%%x\HKLM\SOFTWARE\Acronis\Global\Configuration /v UseServiceCredentialsForShareAccess /t REG_DWORD /d 1 /f > NUL
	REM save error code
	SET el=!errorlevel!
	IF !el!==0 ECHO "UseServiceCredentialsForShareAccess" registry value added successfully to \\%%x\HKLM\SOFTWARE\Acronis\Global\Configuration.
	IF NOT !el!==0 (
		ECHO Registry value not added. Check node.
		REM add node to list of nodes for which reg add operation failed
		SET nodesRegAddFailed[!r!]=%%x
		SET /a r+=1
	)
	ECHO.	

	REM print service name
	sc \\%%x query MMS | find /I "SERVICE_NAME"

	REM stop Acronis Managed Machine Service
	sc \\%%x stop MMS > NUL
	REM save error code
	SET el=!errorlevel!
	REM save state immediately after call - used to make sure the "STOPPED" state isn't printed twice
	FOR /f "tokens=4" %%y in ('sc \\%%x query MMS ^| find /I "STATE"') do set firstSrvStateAfterStopCall=%%y
	IF NOT "!firstSrvStateAfterStopCall!"=="STOPPED" sc \\%%x query MMS | find /I "STATE"

	REM make sure service has stopped
	CALL :stop_wait %%x !el!
	
	REM save state immediately after :stop_wait - used just to make sure the "STOPPED" state isn't printed twice
	FOR /f "tokens=4" %%y in ('sc \\%%x query MMS ^| find /I "STATE"') do set firstSrvStateAfterStopWait=%%y

	REM if not running, print the STOPPED state and start
	REM (should only be RUNNING or START_PENDING at this point if taskkill was used)
	IF NOT "firstSrvStateAfterStopWait"=="RUNNING" (
		sc \\%%x query MMS | find /I "STATE"
		REM do not need to try to start if state is START_PENDING (only occurs after taskkill)
		IF NOT "firstSrvStateAfterStopWait"=="START_PENDING" sc  \\%%x start MMS > NUL
		REM save error code
		SET el=!errorlevel!
		CALL :start_wait %%x !el!
	)

	REM if after taskkill service bounced immediately, just a double check that it's now RUNNING...
	IF "firstSrvStateAfterStopWait"=="RUNNING" CALL :start_wait %%x 0

	REM print RUNNING state
	sc \\%%x query MMS | find /I "STATE"

	ECHO.
	ECHO -------------------------------------------------
	ECHO.
)

REM print list of nodes with possible hanging MMS service
ECHO NODES THAT DID NOT REPORT RUNNING STATUS FOR MMS SERVICE AFTER RESTART ATTEMPT - CHECK NODES
ECHO --------------------------------------------------------------------------------------------
FOR /L %%a IN (0, 1, !q!) DO (
	IF NOT %%a==!q! ECHO !nodesStillHanging[%%a]!
)
ECHO TOTAL: !q!

ECHO.

REM print list of nodes without MMS service
ECHO NODES WITHOUT ACRONIS MANAGED MACHINE SERVICE ^(MMS^)
ECHO ---------------------------------------------------
FOR /L %%a IN (0, 1, !p!) DO (
	IF NOT %%a==!p! ECHO !nodesWithoutMMS[%%a]!
)
ECHO TOTAL: !p!

ECHO.

REM print list of nodes with possible hanging MMS service
ECHO NODES THAT USED TASKKILL AFTER HANGING DURING CLEAN STOP
ECHO --------------------------------------------------------
FOR /L %%a IN (0, 1, !n!) DO (
	IF NOT %%a==!n! ECHO !nodesUsingTaskkill[%%a]!
)
ECHO TOTAL: !n!

ECHO.

REM print list of nodes for which the subkey add operation failed
ECHO NODES FOR WHICH SUBKEY ADD OPERATION FAILED
ECHO -------------------------------------------
FOR /L %%a IN (0, 1, !r!) DO (
	IF NOT %%a==!r! ECHO !nodesRegAddFailed[%%a]!
)
ECHO TOTAL: !r!

ECHO.

REM print list of unavailable nodes (RPC server unavailable errors)
ECHO NODES FOR WHICH ACCESS PRIVILAGES ARE INSUFFICIENT
ECHO -------------------------------------------------
FOR /L %%a IN (0, 1, !m!) DO (
	IF NOT %%a==!m! ECHO !nodesForWhichAccessIsDenied[%%a]!
)
ECHO TOTAL: !m!

ECHO.

REM print list of nodes for which access permissions are insufficient
ECHO UNAVAILABLE NODES ^(RPC SERVER UNAVAILABLE ERRORS^)
ECHO -------------------------------------------------
FOR /L %%a IN (0, 1, !o!) DO (
	IF NOT %%a==!o! ECHO !nodesWithRPCErrors[%%a]!
)
ECHO TOTAL: !o!

ECHO.
ECHO.

REM comment the "PAUSE" below if outputting to log file via command prompt (see README.txt for more details)
REM PAUSE

EXIT /B

REM Function to ensure service is fully stopped before trying to restart
:stop_wait _nodename _errorcode
REM wait @ 2 second interval for each stop check
timeout /T 2 > NUL
IF NOT "!counter!"=="" SET /a counter+=1
IF "!counter!"=="" SET /a counter=1
REM ECHO COUNTER: !counter! (Debugging Only)
REM break for timeout if service hangs while stopping for 1+ minute or error 1061 (Cannot accept control messages at this time.)
IF !counter!==30 SET forceRestart=1
IF %2==1061 SET forceRestart=1
IF !forceRestart!==1 (
	ECHO.
	ECHO Clean stop failed. Attempting forced restart via taskkill...
	taskkill /s %1 /IM MMS.exe /F"
	ECHO.
	SET counter=""
	REM add node to list of nodes with possible hanging MMS service
	SET nodesUsingTaskkill[!n!]=%1
	SET /a n+=1
	REM reset forceRestart Flag
	SET forceRestart=""
	EXIT /B
)
IF %2==5 (
	ECHO ^(Access Denied. Check user permissions.^)
	REM add node to list of nodes for which access was denied (node's with RPC Server errors)
	SET nodesForWhichAccessIsDenied[!m!]=%1
	SET /a m+=1
	EXIT /B
)
REM IF %2==1061 ECHO ^(Cannot accept control messages at this time. Possibly hung while stopping.^) & EXIT /B
IF %2==1062 ECHO ^(Cannot stop Acronis Managed Machine Service. Already stopped.^) & EXIT /B
IF %2==1060 (
	ECHO ^(Cannot stop Acronis Managed Machine Service because it does not exit on node %1.^)
	REM add node to list of nodes with possible hanging MMS service
	SET nodesWithoutMMS[!p!]=%1
	SET /a p+=1
	EXIT /B
)
IF %2==1722 (
	ECHO ^(RPC Server is unavailable. This probably means the node is offline or unreachable.^)
	REM add node to list of unavailable nodes (node's with RPC Server errors)
	SET nodesWithRPCErrors[!o!]=%1
	SET /a o+=1
	EXIT /B
)
IF %2 GTR 0 ECHO ^(SC ERROR %2. Check Service Control Error.^) & EXIT /B
FOR /f "tokens=4" %%a in ('sc \\%1 query MMS ^| find /I "STATE"') do set _srvState=%%a
REM ECHO ERROR CODE PASSED TO STOP_WAIT: %2 (Debugging Only)
REM IF %2==0 IF NOT "!_srvState!"=="STOPPED" ECHO "NOT STOPPED YET..." & GOTO :stop_wait %1 %2 (Debugging Only)
IF %2==0 IF NOT "!_srvState!"=="STOPPED" GOTO :stop_wait %1 %2
REM reset counter
SET counter=""
EXIT /B

REM Function to ensure service is running before continuing to next node
:start_wait _nodename _errorcode
REM wait @ 1 second interval for each stop check
timeout /T 1 > NUL
IF NOT "!counter!"=="" SET /a counter+=1
IF "!counter!"=="" SET /a counter=1
REM ECHO COUNTER: !counter! (Debugging Only)
REM break for timeout if service hangs while starting for 30+ seconds
IF !counter!==30 (
	ECHO.
	ECHO Start failed. Service still not RUNNING. CHECK NODE.
	SET nodesStillHanging[!q!]=%1
	ECHO.
	SET counter=""
	EXIT /B
)
IF %2==1060 ECHO ^(Cannot start Acronis Managed Machine Service because does not exit on node %1.^) & EXIT /B
IF %2==1056 ECHO ^(An instance of the service is already running. Clean start and taskkill may have failed. Check node.^) & EXIT /B
IF %2==5 EXIT /B
IF %2==1722 EXIT /B
IF %2 GTR 0 ECHO ^(SC ERROR %2. Check Service Control Error.^) & EXIT /B
FOR /f "tokens=4" %%a in ('sc \\%1 query MMS ^| find /I "STATE"') do set _srvState=%%a
REM IF %2==0 IF NOT "!_srvState!"=="RUNNING" ECHO "NOT RUNNING YET..." & GOTO :start_wait %1 %2 (Debugging Only)
IF %2==0 IF NOT "!_srvState!"=="RUNNING" GOTO :start_wait %1 %2
REM reset counter
SET counter=""
EXIT /B

REM Resource 1: https://stackoverflow.com/questions/130193/is-it-possible-to-modify-a-registry-entry-via-a-bat-cmd-script
REM Resource 2: https://stackoverflow.com/questions/2591758/batch-script-loop
REM Resource 3: https://stackoverflow.com/questions/20484151/redirecting-output-from-within-batch-file
REM Resource 4: https://serverfault.com/questions/25081/how-do-i-restart-a-windows-service-from-a-script
REM Resource 5: https://stackoverflow.com/questions/6679907/how-do-setlocal-and-enabledelayedexpansion-work
REM Resource 6: https://stackoverflow.com/questions/8363019/how-to-set-a-variable-equal-to-the-contents-of-another-variable
REM Resource 7: https://stackoverflow.com/questions/3262287/make-an-environment-variable-survive-endlocal
REM Resource 8: https://superuser.com/questions/404737/redirect-pipe-to-a-variable-in-windows-batch-file
REM Resource 9: https://stackoverflow.com/questions/36355490/continue-equivalent-command-in-nested-loop-in-windows-batch
REM Resource 10: https://stackoverflow.com/questions/2143187/logical-operators-and-or-in-dos-batch
REM Resource 11: https://stackoverflow.com/questions/14954271/string-comparison-in-batch-file
REM Resource 12: https://www.computerhope.com/forum/index.php?topic=140911.0
REM Resource 13: https://stackoverflow.com/questions/12976351/escaping-parentheses-within-parentheses-for-batch-file
REM Resource 14: https://stackoverflow.com/questions/37071353/how-to-check-if-a-variable-exists-in-a-batch-file/37073832
REM Resource 15: https://stackoverflow.com/questions/17605767/create-list-or-arrays-in-windows-batch
REM Resource 16: https://stackoverflow.com/questions/10166386/arrays-linked-lists-and-other-data-structures-in-cmd-exe-batch-script/10167990#10167990
REM Resource 17: https://stackoverflow.com/questions/14334850/why-this-code-says-echo-is-off/39141842
REM Resource 18: https://stackoverflow.com/questions/6828751/batch-character-escaping