' -------------------------------------------------------------------------------
'
' Sophos Cloud Install script
' This script has been created to help users to deploy Sophos Cloud Installer through 
' their network using GPO.
'
' sophos_cloud_installer.vbs
' Script created by Rafael Foster (rafaelgfoster at gmail dot com)
'
' -------------------------------------------------------------------------------
'On error resume Next
Set WshShell     = CreateObject("WScript.Shell")
Set objRegEx     = CreateObject("VBScript.RegExp")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE
Const WAIT_BEFORE_INSTALL = 0 'Wait before install Sophos. 0 to install immediately

if CheckSophosInstalled = TRUE Then
	Wscript.Quit
End If

' -------------------------------------------------
' Definição do Regex 
objRegEx.Global     = True
objRegEx.IgnoreCase = True
objRegEx.Pattern    = "\-\d{5,6}$"

' Windows folders variables
strTempFolder   = wshShell.ExpandEnvironmentStrings( "%TEMP%" ) & "\"
strUserName     = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strLogonServer  = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strProgFiles    = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES%" )  & "\"
strProgFilesx86 = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES(x86)%" )  & "\"

For each System in SystemSet 
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
Next

' Uncomment this if you want to skip this to install on servers
'if ( inStr(LCase(SysVersion),"server") <> 0 ) Then
'	Wscript.Quit
'End If

' Sophos Variables
strPreExec               = ""
str_OptionQuiet 		 = TRUE
str_OptionProcutsInstall = "all"         'Choose which products will be installed. Available options are: antivirus, intercept, deviceEncryption or all
strUseProxy              = FALSE         ' Change to TRUE or FALSE if you want to use a Proxy server to cache sophos install
strProxyServer           = "192.168.0.1" ' If strProxyServer have been set to TRUE, change this if user are required to connect to proxy
strProxyUserName         = "proxyuser"   ' If strProxyUserName have been set to TRUE, change this if user are required to connect to proxy
strProxyUserPass         = "proxypass"   ' If strProxyUserPass have been set to TRUE, change this if password are required to connect to proxy

' -----[ Admin Credentials ]---------------------------------------
' Set it to TRUE to use admin credentials to install it on Machine. If you set to False and user does not have 
' privileges on installation, password will be prompted to users
strUseAdminCredentials   = TRUE 
strAdminDomain           = "FOSTER"
strAdminUserAccount      = "admin"
strAdminPassword         = "F0ster@001"

' Change this to determine where this script will get the installer executable and their options.
strSophosInstallLog        = "\\localhost\teste\Log_Install\"
strSophosExecChecker       = "Sophos\Management Communications System\Endpoint\McsClient.exe"
strSophosAuxExec           = "\\localhost\teste\soph_aux_runas.exe"
strlatestSophosInstallFile = "\\localhost\teste\SophosSetup.exe"

' Define log name and path.
strCurrentInstallLog = strSophosInstallLog & "log_installSophosCloud_" & "(" & strComputerName & ").log"

' ------------------------------------------------------------------------------------------------------------------------------------------------
' Creating file for logging
If objFSO.FileExists(strCurrentInstallLog) Then
	Set objLogFile = objFSO.OpenTextFile(strCurrentInstallLog, ForWriting, True)
Else
	Set objLogFile = objFSO.CreateTextFile(strCurrentInstallLog)
End If


' ------------------------------------------------------------------------------------------------------------------------------------------------
objLogFile.WriteLine "-------[ System Information ]---------------------------------------------------"
objLogFile.WriteLine
objLogFile.WriteLine "Usuario               : " & strUserName
objLogFile.WriteLine "Estacao de Trabalho   : " & strComputerName
objLogFile.WriteLine "Servidor de Logon     : " & strLogonServer
objLogFile.WriteLine "Sistema Operacional   : " & SysVersion
objLogFile.WriteLine
objLogFile.WriteLine

objLogFile.WriteLine "-------[ Sophos Install Information ]--------------------------------------------------"
objLogFile.WriteLine 

' ------------------------------------------------------------------------------------------------------------------------------------------------
' Sophos Install Arguments

strSophosInstallArguments=""	
if str_OptionQuiet = TRUE Then
	strSophosInstallArguments =	strSophosInstallArguments & " --quiet"
End if
if Len(str_OptionProcutsInstall) > 0 Then
	strSophosInstallArguments =	strSophosInstallArguments & " --products=" & str_OptionProcutsInstall
End if
if strUseProxy Then
	strSophosInstallArguments =	strSophosInstallArguments & " --proxyaddress=" & strProxyServer
	if Len(strProxyUserName) > 0 Then
		strSophosInstallArguments =	strSophosInstallArguments & " --proxyusername=" & strProxyUserName
	End if
	if Len(strProxyUserPass) > 0 Then
		strSophosInstallArguments =	strSophosInstallArguments & " --proxypassword=" & strProxyUserPass
	End if
	if Len(strUseProxy) > 0 Then
		strSophosInstallArguments =	strSophosInstallArguments & " --proxyaddress=" & strProxyServer
	End if
End if


objLogFile.Writeline "Sophos Install Path: " & strlatestSophosInstallFile 
objLogFile.Writeline "Sophos Install Arguments: " & strSophosInstallArguments 

strSophosInstallFileName = Split(strlatestSophosInstallFile,"\",-1,1)
For Each arrName in strSophosInstallFileName
	strInstallFile = arrName
Next

objFSO.CopyFile strlatestSophosInstallFile, strTempFolder
strlatestSophosInstallFile = strTempFolder & strInstallFile

' --------------------------------------------------------------------------------------------------------------
' Verificar se o diretorio de instalacao padrao Sophos existe

strSophosInstalled = FALSE
If (objFSO.FileExists(strProgFiles & strSophosExecChecker) ) Then
	strSophosInstalled = TRUE
Elseif (objFSO.FileExists(strProgFilesx86 & "\Sophos Inventory Agent\SophosInventory.exe") ) Then
	strSophosInstalled = TRUE
End If

if strSophosInstalled Then
	objLogFile.Writeline "File to Check Sophos Management Communications System (Comunicator with Sophos Central): " & strSophosExecChecker & ": Exists!"  
	objLogFile.Writeline "Sophos Comunication Services exist, but Service not found! Maybe you should reinstall it"
	objLogFile.Writeline "Install finished!"
End if

Wscript.Sleep WAIT_BEFORE_INSTALL

objLogFile.Write "Running Sophos Install exec: " 
objLogFile.Write strlatestSophosInstallFile & " " & strSophosInstallCMDArguments
objLogFile.Writeline ""

if strUseAdminCredentials Then
	strPreExec = strSophosAuxExec & " " & strAdminUserAccount & " " & strAdminPassword & " " & strAdminDomain
	strExecCommand = strPreExec & " " & """" & strlatestSophosInstallFile & " " & strSophosInstallCMDArguments & """" 
	MsgBox(strExecCommand)
	objLogFile.Writeline strExecCommand
	REM MsgBox(strShellOutput.StdOut.ReadLine)
	objLogFile.Writeline "Installing using " & strAdminUserAccount & " account."
else
	WshShell.Run strlatestSophosInstallFile & " " & strSophosInstallCMDArguments, 0, TRUE
End if



objStartFolder = strTempFolder & "\sophos_bootstrap\"

REM if objFSO.FolderExists(objStartFolder) Then

	Set objFolder  = objFSO.GetFolder(objStartFolder)
	Set colFiles   = objFolder.Files

	For Each objFile in colFiles
		if instr(objFile.Name,"Bootstrap_") > 0 Then
			Set objSophosLogFile = objFSO.OpenTextFile(strTempFolder & "\sophos_bootstrap\" & objFile.Name, ForReading, True)
			strFileContent = objSophosLogFile.ReadAll

			objLogFile.WriteLine "-------[ Installation Log Information ]---------------------------------------------------"
			objLogFile.Writeline " " 
			objLogFile.Write strFileContent		
		End if
	Next
REM End if

Function CheckSophosInstalled()
	
	strSophosServiceFound = FALSE
	strComputerName = "."
	 Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2") 

	Set colServices = objWMIService.ExecQuery _ 
		("Select * from Win32_Service") 

	For Each Service in colServices
		if instr(Service.Name,"Sophos") > 0 Then
			strSophosServiceFound = TRUE
		End if
	Next

	CheckSophosInstalled = strSophosServiceFound
	
End Function
