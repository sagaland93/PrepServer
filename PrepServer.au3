#NoTrayIcon
#RequireAdmin
Opt("TrayAutoPause", 0)


#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
	#AutoIt3Wrapper_Icon=..\..\..\Bilder\Icons\Optimize_v2.ico
	#AutoIt3Wrapper_Outfile_x64=.\PrepServer.exe
	#AutoIt3Wrapper_Compression=4
	#AutoIt3Wrapper_UseX64=y
	#AutoIt3Wrapper_Res_Comment=Prep system for Provisioning
	#AutoIt3Wrapper_Res_Description=Prep system for Provisioning
	#AutoIt3Wrapper_Res_Fileversion=2.0.0.76
	#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
	#AutoIt3Wrapper_Res_ProductName=PrepServer
	#AutoIt3Wrapper_Res_ProductVersion=Public release
	#AutoIt3Wrapper_Res_LegalCopyright=© 2020 André Saga Lande
	#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
	#AutoIt3Wrapper_Res_HiDpi=Y
	#AutoIt3Wrapper_Res_File_Add=C:\MyCloud\Bilder\GIF\LoadingBar_141414.gif, RT_RCDATA, GIF_1, 0
	#AutoIt3Wrapper_Run_After=copy "%outx64%" "%scriptdir%\%scriptfile%\%scriptfile%.exe"
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#cs ============================================================================================

	Author:			André Saga Lande

	Description:	PrepServer is a tool used to help maintain and optimize Windows terminal servers.
					Especially useful for Citrix Provisioning and Machine Creation Services.

	+ 	New
	^ 	Fix
	§	Changed
	- 	Removed
	* 	Improved
	! 	Important


	TODO


	RELEASE NOTES


#ce ============================================================================================


; START


;===============================================================================================

;   DECLARATIONS AND INCLUDED FUNCTIONS

#Region		;	DECLARATIONS

	Global $AdminUser = "Administrator"
	Global $AdminGroup = "Administrators"
	Global $Config = @ScriptDir & "\Config.ini"
	Global $LogFolder = @ScriptDir & "\Logs\"
	Global $Log = $LogFolder & "PrepServer_" & @MDAY & "." & @MON & "." & @YEAR & ".log"
	Global $ReportsFolder = @ScriptDir & "\Reports\"
	Global $AppReportsFolder = $ReportsFolder & "Applications\"
	Global $UpdatesReportsFolder = $ReportsFolder & "Windows Updates\"
	Global $ScriptsFolder = @ScriptDir & "\Scripts"
	Global $RegistryFolder = @ScriptDir & "\Registry"
	Global $ServerType = "system"
	Global $HostType = ""
	Global $vDisk = "", $vDisk_CacheType = ""
	Global $vDisk_Type_ReadWrite = "0", $vDisk_Type_Private = "10", $vDisk_Type_Read = "12"
	Global $XenTools, $HyperV, $VMware, $PVS, $MCS
	Global Const $BackgroundColor = 0x141414
	; #1A1A1A
	; #141414
	Global Const $BackgroundColor_Reset = 0xB14D1C
	; #B14D1C
	Global Const $LabelColor = $BackgroundColor
	Global Const $LabelColor_Reset = 0xB14D1C
	Global Const $DefaultFontColor = 0xABABAB
	Global Const $ErrorFontColor = 0xABABAB
	Global Const $SuccessFontColor = 0xABABAB
	; #ABABAB
	Global $GUI

	; Initialize()
	Const $cNotSet = 0
	Global $iMode = $cNotSet
	Global $OverrideShutdownToNo, $OverrideRebootToNo, $RunSilently
	Global $HKLM = "HKEY_LOCAL_MACHINE64"

	; PrepServer()
	Global $hGIF
	Global $PendingReboot

	; _ExecuteFiles()
	Global $ExecuteFiles = False
	Global $BAT_Path, $CurrentBAT, $RunBAT

	; _AutoLogonEnable()
	Global $ScrambledPassword

	; _ReadConfig()
	Global $AutoLogonEnabled, $AutoLogonUserName, $AutoLogonDomain, $AutoLogonPassword, $AutoLogonCount, $PasswordEncrypted
	Global $RunScripts, $EmptyRecycleBin, $ImportRegFiles, $RemoveRegEntries
	Global $Shutdown, $ShutdownPrompt, $ShutdownCancel, $RebootCancel, $Reboot, $RebootPrompt
	Global $DeleteWindowsUpdateCache
	Global $MustBeLicensedWithKMS
	Global $ExitTimer = 8

	; PendingReboot()
	Global $WindowsUpdateHasBeenRun = False

#EndRegion	; > DECLARATIONS
#Region		;	INCLUDES

	#include <WindowsConstants.au3>
	#include <ButtonConstants.au3>
	#include <MsgBoxConstants.au3>
	#include <StringConstants.au3>
	#include <StaticConstants.au3>
	#include <AutoItConstants.au3>
	#include <ColorConstants.au3>
	#include <GUIConstantsEx.au3>
	#include <_GetNewestFile.au3>
	#include <WinAPIShellEx.au3>
	#include <TaskScheduler.au3>
	#include <FontConstants.au3>
	#include <BlockInputEx.au3>
	#include <GUIConstants.au3>
	#include <GIFAnimation.au3>
	#include <Permissions.au3>
	#include <_RunWaitGet.au3>
	#include <WinAPIFiles.au3>
	#include <WinAPIMisc.au3>
	#include <_RegEnumEx.au3>
	#include <_Scrambler.au3>
	#include <WinAPISys.au3>
	#include <WinAPIReg.au3>
	#include <Constants.au3>
	#include <_RegFunc.au3>
	#include <EventLog.au3>
	#include <Services.au3>
	#include <Hotkey.au3>
	#include <Array.au3>
	#include <Crypt.au3>
	#include <File.au3>
	#include <Task.au3>
	#include <DTC.au3>

#EndRegion	; > INCLUDES


#Region		;	HOTKEYS

	HotKeySet("+!q", "_Abort")

#EndRegion	; > HOTKEYS


;===============================================================================================

;   SCRIPT ACTUAL START

#Region		;	INITIALIZE

	Initialize()
	Func Initialize()
		_BlockInputEx(3, "", "{LCTRL}|{LWIN}|{RWIN}")
		If FileExists($LogFolder) Then ; LOG FOLDER
			_FileWriteLog($Log, "[START] Initializing...")
		Else
			_FileWriteLog($Log, "[INFO] Initial setup...")
			_FileWriteLog($Log, "[INFO] Creating Log directory...")
			DirCreate($LogFolder)
			If @error = 0 Then
				_FileWriteLog($Log, "[INFO] Log folder created: " & $LogFolder)
			Else
				_FileWriteLog($Log, "[ERROR] Failed to create Log directory: " & $LogFolder)
			EndIf
		EndIf
		If Not FileExists($ReportsFolder) Then ; REPORT FOLDER
			DirCreate($ReportsFolder)
			If @error = 0 Then _FileWriteLog($Log, "[INFO] Report folder created.")
		EndIf
		If Not FileExists($ScriptsFolder) Then ; SCRIPT FOLDER
			DirCreate($ScriptsFolder)
			If @error = 0 Then _FileWriteLog($Log, "[INFO] Script folder created.")
		EndIf
		If Not FileExists($RegistryFolder) Then ; REGISTRY FOLDER
			DirCreate($RegistryFolder)
			If @error = 0 Then _FileWriteLog($Log, "[INFO] Registry folder created.")
		EndIf
		If Not FileExists($Config) Then
			_CreateConfig()
			_ReadConfig()
		Else
			_ReadConfig()
		EndIf

		; PERMISSION RESOURCES
		_InitiatePermissionResources()

		; CRYPT RESOURCES
		_Crypt_Startup()

		; AUTO LOGON - LOOK FOR UNENCRYPTED PASSWORD
		If Not $AutoLogonPassword = "" And $PasswordEncrypted = "False" Then
			_FileWriteLog($Log, "[INFO] Unencrypted password found... Scrambling password...")
			$ScrambledPassword = _Scramble($AutoLogonPassword)
				If Not @error Then _FileWriteLog($Log, "[INFO] Password successfully encrypted.")
			_UpdateConfig("AutoLogon", "Password", $ScrambledPassword)
				If Not @error Then _FileWriteLog($Log, "[INFO] Password successfully written to 'Config.ini'.")
			_UpdateConfig("AutoLogon", "PasswordEncrypted", "True")
		EndIf

		; CHECK OS BIT-VERSION
		If Not StringInStr(@OSArch, "X64") Then $HKLM = StringReplace($HKLM, "HKEY_LOCAL_MACHINE64", "HKEY_LOCAL_MACHINE")

		; INSTALL FONT
		FileInstall("C:\MyCloud\Operativsystem\Oppsett\Fonts\Octopus_300.otf", "C:\Windows\Temp\Octopus_300.otf")
		_WinAPI_AddFontResourceEx("C:\Windows\Temp\Octopus_300.otf")

		; COLLECT SYSTEM INFORMATION
		_FileWriteLog($Log, "[INFO] Collecting system information...")
		_FileWriteLog($Log, "=======================================")
		_FileWriteLog($Log, "Environment details")
		_FileWriteLog($Log, "Current User: " & @TAB & @UserName)
		_FileWriteLog($Log, "Computer Name: " & @TAB & @ComputerName)
		_FileWriteLog($Log, "Operating System: " & @OSVersion & " " & @OSArch & " " & @OSServicePack)
		_FileWriteLog($Log, "License Status: " & @TAB & LicenseStatus())
		_FileWriteLog($Log, "License Type: " & @TAB & LicenseType())

		$XenTools = RegEnumKey($HKLM & "\SOFTWARE\Citrix\XenTools", 1)
		$HyperV = RegEnumKey($HKLM & "\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters", 1)
		$VMware = RegEnumKey($HKLM & "\SOFTWARE\VMware, Inc.\VMware Tools", 1)
		$MCS = RegEnumKey($HKLM & "\SOFTWARE\Citrix\MachineIdentityServiceAgent", 1)
		$PVS = RegEnumKey($HKLM & "\SOFTWARE\Citrix\ProvisioningServices", 1)
		$vDisk = RegRead($HKLM & "\SYSTEM\CurrentControlSet\Services\bnistack\PvsAgent", "DiskName")
		$vDisk_CacheType = RegRead($HKLM & "\SYSTEM\CurrentControlSet\Services\bnistack\PvsAgent", "WriteCacheType")

		; MACHINE TYPE
		If $PVS Then
			_FileWriteLog($Log, "Platform: " & @TAB & "Citrix Provisioning Target Device vDisk")
			If $vDisk Then
				$ServerType = "vDisk"
				_FileWriteLog($Log, "vDisk: " & @TAB & @TAB & $vDisk)
				If $vDisk_CacheType = $vDisk_Type_ReadWrite Then
					_FileWriteLog($Log, "Type: " & @TAB & @TAB & "Read/Write Version")

				ElseIf $vDisk_CacheType = $vDisk_Type_Private Then
					_FileWriteLog($Log, "Type: " & @TAB & @TAB & "Private Read/Write")
				ElseIf $vDisk_CacheType = $vDisk_Type_Read Then
					_FileWriteLog($Log, "Type: " & @TAB & @TAB & "Read-only")
				EndIf
			EndIf
		ElseIf $MCS Then
			_FileWriteLog($Log, "Platform: " & @TAB & "Machine Creation Services Device")
			$ServerType = "MCS Persistant Disk"
		ElseIf ($VMware or $HyperV Or $XenTools) Then
			_FileWriteLog($Log, "Machine Type: " & @TAB & "Virtual")
			$ServerType = "VM"
		Else
			_FileWriteLog($Log, "Machine Type: " & @TAB & "Physical")
			$HostType = "Physical"
		EndIf

		; HOST TYPE
		If $XenTools Then
			$HostType = "Citrix XenServer"
			_FileWriteLog($Log, "Host Type: " & @TAB & $HostType)
		ElseIf $HyperV Then
			$HostType = "Hyper-V"
			_FileWriteLog($Log, "Host Type: " & @TAB & $HostType)
		ElseIf $VMware Then
			$HostType = "VMware ESXi"
			_FileWriteLog($Log, "Host Type: " & @TAB & $HostType)
		Else
			$HostType = "Unknown"
			_FileWriteLog($Log, "Host Type: " & @TAB & $HostType)
		EndIf
		_FileWriteLog($Log, "=======================================")

		PrepServer()

	EndFunc

#EndRegion	; > INITIALIZE
#Region		;	CONFIG

	Func _CreateConfig()
		Local $hFileOpen
		_FileCreate($Config)
		If Not @error Then
			$hFileOpen = FileOpen($Config, $FO_APPEND)

			FileWrite($Config, "====================================" & @CRLF)
			FileWrite($Config, "====== PrepServer config file ======" & @CRLF)
			FileWrite($Config, "====================================" & @CRLF)
			FileWrite($Config, @CRLF)
			IniWrite($Config, "AutoLogon", "Enabled", "True")
			IniWrite($Config, "AutoLogon", "Username", @UserName)
			IniWrite($Config, "AutoLogon", "Domain", @LogonDomain)
			IniWrite($Config, "AutoLogon", "Password", "")
			IniWrite($Config, "AutoLogon", "PasswordEncrypted", "False")
			IniWrite($Config, "AutoLogon", "LogonCount", "1")
			FileWrite($Config, @CRLF)
			IniWrite($Config, "Tasks", "Adobe Acrobat Update Task", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\Windows\Customer Experience Improvement Program\Consolidator", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\Windows\Customer Experience Improvement Program\UsbCeip", "Disable")
			IniWrite($Config, "Tasks", "GoogleUpdateTaskMachineCore", "Disable")
			IniWrite($Config, "Tasks", "GoogleUpdateTaskMachineUA", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\Windows\WindowsUpdate\Automatic App Update", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\Windows\WindowsUpdate\Scheduled Start", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\Windows\WindowsUpdate\sih", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\Windows\WindowsUpdate\sihboot", "Disable")
			IniWrite($Config, "Tasks", "Microsoft\XblGameSave\XblGameSaveTask", "Disable")
			FileWrite($Config, @CRLF)
			IniWrite($Config, "Services", "AdobeARMservice", "Disable")
			IniWrite($Config, "Services", "AdobeUpdateService", "Disable")
			IniWrite($Config, "Services", "wuauserv", "Disable")
			IniWrite($Config, "Services", "CitrixTelemetryService", "Disable")
			FileWrite($Config, @CRLF)
			IniWrite($Config, "Features", "IPv6", "Disable")
			IniWrite($Config, "Features", "SMB1", "Disable")
			IniWrite($Config, "Features", "Indexing", "Disable")
			IniWrite($Config, "Features", "Offline files", "Disable")
			IniWrite($Config, "Features", "TCP Offloading", "Disable")
			IniWrite($Config, "Features", "Windows Firewall", "Disable")
			IniWrite($Config, "Features", "IE Enhanced Security Configuration", "Disable")
			FileWrite($Config, @CRLF)
			IniWrite($Config, "Functions", "EmptyRecycleBin", "True")
			IniWrite($Config, "Functions", "DeleteWindowsUpdateCache", "True")
			IniWrite($Config, "Functions", "ImportRegFiles", "True")
			IniWrite($Config, "Functions", "RemoveRegEntries", "True")
			IniWrite($Config, "Functions", "RunScripts", "True")
			IniWrite($Config, "Functions", "MustBeLicensedWithKMS", "True")
			IniWrite($Config, "Functions", "Shutdown", "True")
			IniWrite($Config, "Functions", "ShutdownPrompt", "False")
			IniWrite($Config, "Functions", "ShutdownCancel", "False")
			IniWrite($Config, "Functions", "Reboot", "True")
			IniWrite($Config, "Functions", "RebootPrompt", "False")
			IniWrite($Config, "Functions", "RebootCancel", "False")
			IniWrite($Config, "Functions", "ExitTimer", "8")

			FileClose($hFileOpen)
		Else
			_FileWriteLog($Log, "[ERROR] Failed to create config file: " & $Config)
		EndIf
	EndFunc

	Func _ReadConfig()
		Local $Count = 0

		_FileWriteLog($Log, "[INFO] Reading config: " & $Config)

		$AutoLogonEnabled = IniRead($Config, "AutoLogon", "Enabled", "N/A")
			If Not @error Then $Count += 1
		$AutoLogonUserName = IniRead($Config, "AutoLogon", "Username", "N/A")
			If Not @error Then $Count += 1
		$AutoLogonDomain = IniRead($Config, "AutoLogon", "Domain", "N/A")
			If Not @error Then $Count += 1
		$AutoLogonPassword = IniRead($Config, "AutoLogon", "Password", "N/A")
			If Not @error Then $Count += 1
		$PasswordEncrypted = IniRead($Config, "AutoLogon", "PasswordEncrypted", "N/A")
			If Not @error Then $Count += 1
		$AutoLogonCount = IniRead($Config, "AutoLogon", "LogonCount", "N/A")
			If Not @error Then $Count += 1
		$EmptyRecycleBin = IniRead($Config, "Functions", "EmptyRecycleBin", "N/A")
			If Not @error Then $Count += 1
		$DeleteWindowsUpdateCache = IniRead($Config, "Functions", "DeleteWindowsUpdateCache", "N/A")
			If Not @error Then $Count += 1
		$ImportRegFiles = IniRead($Config, "Functions", "ImportRegFiles", "N/A")
			If Not @error Then $Count += 1
		$RemoveRegEntries = IniRead($Config, "Functions", "RemoveRegEntries", "N/A")
			If Not @error Then $Count += 1
		$RunScripts = IniRead($Config, "Functions", "RunScripts", "N/A")
			If Not @error Then $Count += 1
		$MustBeLicensedWithKMS = IniRead($Config, "Functions", "MustBeLicensedWithKMS", "N/A")
			If Not @error Then $Count += 1
		$Shutdown = IniRead($Config, "Functions", "Shutdown", "N/A")
			If Not @error Then $Count += 1
		$ShutdownPrompt = IniRead($Config, "Functions", "ShutdownPrompt", "N/A")
			If Not @error Then $Count += 1
		$ShutdownCancel = IniRead($Config, "Functions", "ShutdownCancel", "N/A")
			If Not @error Then $Count += 1
		$Reboot = IniRead($Config, "Functions", "Reboot", "N/A")
			If Not @error Then $Count += 1
		$RebootPrompt = IniRead($Config, "Functions", "RebootPrompt", "N/A")
			If Not @error Then $Count += 1
		$RebootCancel = IniRead($Config, "Functions", "RebootCancel", "N/A")
			If Not @error Then $Count += 1
		$ExitTimer = IniRead($Config, "Functions", "ExitTimer", "N/A")
			If Not @error Then $Count += 1

		If $Count >= 19 Then _FileWriteLog($Log, "[INFO] Successfully read config.")
	EndFunc

	Func _UpdateConfig($section, $key, $value)
		IniWrite($Config, $section, $key, $value)
		Local $UpdatedValue = IniRead($Config, $section, $key, $value)

		Return $UpdatedValue
	EndFunc

	Func _VerifyConfig()
		;
	EndFunc

#EndRegion	; > CONFIG
#Region		;	PREP SERVER

	Func PrepServer()
		#Region ;	GUI
			Local $Counter = $ExitTimer
			Local $GUI_Width = @DesktopWidth + 50
			Local $GUI_Height = @DesktopHeight + 50
			Local $Label_Width = 500
			Local $Label_Height = 850
			Local $Label_Spacer = 50
			Local $Label_TitleSpacer = 125
			Local $GIF_Height = 20

			$GUI = GUICreate("PrepServer", $GUI_Width, $GUI_Height, -1, -1, BitOR($WS_POPUP, $WS_BORDER), $WS_EX_TOPMOST)
			GUISetBkColor($BackgroundColor)

			; TITLE
			Local $LABEL_Title = GUICtrlCreateLabel("Preparing " & $ServerType & " for Provisioning", 0, $GUI_Height / 2 - $Label_TitleSpacer, $GUI_Width, 100, $SS_CENTER)
			GUICtrlSetFont(-1, 30, 300, "", "Octopus_300", $CLEARTYPE_QUALITY)
			GUICtrlSetColor(-1, $DefaultFontColor)

			; STATUS
			Global $LABEL_StatusText = GUICtrlCreateLabel("Initializing", 0, $GUI_Height / 2 - $Label_Spacer, $GUI_Width, 50, $SS_CENTER)
			GUICtrlSetFont(-1, 20, 300, "", "Octopus_300", $CLEARTYPE_QUALITY)
			GUICtrlSetColor(-1, $DefaultFontColor)
			GUICtrlSetBkColor(-1, $LabelColor)

			Global $LABEL_Counter = GUICtrlCreateLabel("", 0, $GUI_Height / 2 + $Label_Spacer, $GUI_Width, 100, $SS_CENTER)
			GUICtrlSetFont(-1, 80, 300, "", "Octopus_300", $CLEARTYPE_QUALITY)
			GUICtrlSetColor(-1, $DefaultFontColor)
			GUICtrlSetBkColor(-1, $LabelColor)

			; LOADING BAR
			$hGIF = _GUICtrlCreateGIF(@AutoItExe, "10;GIF_1", $GUI_Width / 2 - 100, $GUI_Height / 2 + $Label_Spacer)

			; PERCENT
			Global $LABEL_Percent = GUICtrlCreateLabel("0%", 0, $GUI_Height / 2 + $Label_Spacer * 3, $GUI_Width, 50, $SS_CENTER)
			GUICtrlSetFont(-1, 20, 800, "", "Octopus_300", $CLEARTYPE_QUALITY)
			GUICtrlSetColor(-1, $DefaultFontColor)
			GUICtrlSetBkColor(-1, $LabelColor)

			GUISetState()
		#EndRegion; GUI

		While 1
			Switch GUIGetMsg()
				Case $GUI_EVENT_MINIMIZE
					_Exit()
			EndSwitch

			_WinAPI_RedrawWindow($GUI)
			_WinAPI_RedrawWindow($GUI)

			; PENDING
			__Progress("1", "5")
			GUICtrlSetData($LABEL_StatusText, "Checking for Pending reboot")
			Sleep(500)
			_FileWriteLog($Log, "[INFO] Checking for Pending reboot...")
			$PendingReboot = PendingReboot()
			If @error Then
				_GIF_DeleteGIF($hGIF)
				_Reboot()
			EndIf

			; KMS
			If $MustBeLicensedWithKMS = "True" Then
				GUICtrlSetData($LABEL_StatusText, "Checking Windows License")
				If Not LicenseStatus() And $vDisk_CacheType <> $vDisk_Type_Private And $HostType <> "Physical" Then
					_GIF_DeleteGIF($hGIF)
					_FileWriteLog($Log, "[INFO] Windows is not licensed. Exiting...")
					_FileWriteLog($Log, "[INFO] Activate Windows with a KMS host using a KMS Client key before shutting down and closing this vDisk.")
					GUICtrlSetData($LABEL_Percent, "")
					GUISetBkColor($BackgroundColor_Reset, $GUI)
					GUICtrlSetBkColor($LABEL_Percent, $LabelColor_Reset)
					GUICtrlSetBkColor($LABEL_Counter, $LabelColor_Reset)
					GUICtrlSetBkColor($LABEL_StatusText, $LabelColor_Reset)
					GUICtrlSetData($LABEL_StatusText, "Warning: Windows is not licensed")
					Do
						GUICtrlSetData($LABEL_Counter, $Counter)
						$Counter -= 1
						Sleep(1000)
					Until $Counter = 0
					Sleep(1000)
					_Exit()
				EndIf
			EndIf
			GUICtrlSetData($LABEL_StatusText, "Processing tasks")
			Sleep(500)

			; TASKS
			__Progress("6", "9")
			_ProcessTasks()
			__Progress("10", "19")
			GUICtrlSetData($LABEL_StatusText, "Disabling services")
			Sleep(500)

			; SERVICES
			__Progress("20", "29")
			_DisableServices()
			__Progress("30", "35")

			; REGISTRY
			If $ImportRegFiles = "True" Then
				GUICtrlSetData($LABEL_StatusText, "Importing Registration Entries")
				Sleep(500)
				__Progress("36", "40")
				_ImportRegFiles()
			EndIf
			If $RemoveRegEntries = "True" Then
				GUICtrlSetData($LABEL_StatusText, "Removing Registration Entries")
				Sleep(500)
				__Progress("42", "44")
				_RemoveRegEntries()
			EndIf
			__Progress("45", "55")

			; SCRIPTS
			If $RunScripts = "True" Then
				GUICtrlSetData($LABEL_StatusText, "Running scripts")
				__Progress("56", "70")
				_ExecuteFiles()
			EndIf

			; REPORTS
			GUICtrlSetData($LABEL_StatusText, "Creating reports")
			__Progress("71", "80")
			_ApplicationInstallsReport()
			_WindowsUpdateReport()

			; CLEANUP
			GUICtrlSetData($LABEL_StatusText, "Running cleanup")
			_Cleanup()
			__Progress("81", "84")
			GUICtrlSetData($LABEL_StatusText, "Flushing DNS")
			Sleep(500)

			; FLUSH
			__Progress("85", "93")
			_FlushDNS()
			__Progress("94", "99")

			; PROCESSES
			Processes()
			GUICtrlSetData($LABEL_StatusText, "Finishing up")
			Sleep(500)
			_GIF_DeleteGIF($hGIF)
			GUICtrlSetData($LABEL_Percent, "100%")

			; SHUTDOWN & EXIT
			_Shutdown()
		WEnd

	EndFunc

#EndRegion	; > PREP SERVER


#Region		;	TASKS

	Func _ProcessTasks()
		Local $TasksToDisable, $DisableTask

		_FileWriteLog($Log, "[INFO] Processing Scheduled Tasks...")
		Local $KeyName = $HKLM & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree"
		Local $ListOfTasks = _RegEnumKeyEx($KeyName)
		For $i = 0 to UBound($ListOfTasks,1) - 1
			$ListOfTasks[$i] = StringReplace($ListOfTasks[$i], $KeyName & "\", "")
		Next

		$TasksToDisable = IniReadSection($Config, "Tasks")
		For $i = 0 To UBound($TasksToDisable) - 1
			If _ArraySearch($ListOfTasks, $TasksToDisable[$i][0]) <> -1 Then
				_FileWriteLog($Log, "[INFO] Disabling task: " & $TasksToDisable[$i][0])
				$DisableTask = _TaskChange($TasksToDisable[$i][0], 0)
				If StringInStr($DisableTask, "already been disabled") Then
					_FileWriteLog($Log, "[INFO] Task is already disabled.")
				ElseIf StringInStr($DisableTask, "SUCCESS") Then
					_FileWriteLog($Log, "[INFO] Successfully disabled task.")
				ElseIf StringInStr($DisableTask, "does not exist") Then
					_FileWriteLog($Log, "[ERROR] Failed to disable task because it does not exist.")
				EndIf
			EndIf
		Next
	EndFunc

#EndRegion	; > TASKS
#Region		;	SERVICES

	Func _DisableServices()

		_FileWriteLog($Log, "[INFO] Processing Services...")
		Local $ServicesToDisable = IniReadSection($Config, "Services")
		_ArrayDelete($ServicesToDisable, 0)
		For $i = 0 To UBound($ServicesToDisable) - 1
			_FileWriteLog($Log, "[INFO] Checking service status: " & $ServicesToDisable[$i][0])
			Local $ServiceStatus = _Service_QueryStatus($ServicesToDisable[$i][0])
			If $ServiceStatus[1] = 4 Then
				_FileWriteLog($Log, "[INFO] Service is running.")
				_FileWriteLog($Log, "[INFO] Stopping service...")
				Local $StopService = _Service_Stop($ServicesToDisable[$i][0])
				If @error = 0 Then
					_FileWriteLog($Log, "[INFO] Successfully stopped the service.")
				Else
					_FileWriteLog($Log, "[ERROR] Failed to stop the service.")
				EndIf
			ElseIf $ServiceStatus[1] = 1 Then
				_FileWriteLog($Log, "[INFO] Service is stopped.")
			EndIf


			Local $ServiceConfig = _Service_QueryConfig($ServicesToDisable[$i][0])
			If $ServiceConfig[1] <> 4 Then
				_FileWriteLog($Log, "[INFO] Disabling service...")
				Local $DisableService = _Service_Change($ServicesToDisable[$i][0], $SERVICE_NO_CHANGE, $SERVICE_DISABLED)
				If @error = 0 Then
					_FileWriteLog($Log, "[INFO] Successfully disabled the service.")
				Else
					_FileWriteLog($Log, "[ERROR] Failed to disable the service.")
				EndIf
			ElseIf $ServiceConfig[1] = 4 Then
				_FileWriteLog($Log, "[INFO] Service is disabled.")
			EndIf
		Next

	EndFunc

#EndRegion	; > SERVICES
#Region		;	EXECUTE

	Func _ExecuteFiles()
		Local $BATs
		$ExecuteFiles = True

		_FileWriteLog($Log, "[INFO] Looking for scripts to run...")
		$BAT_Path = @ScriptDir & "\Scripts\"
		$BATs = _FileListToArray($BAT_Path, "*.BAT")
		If IsArray($BATs) Then
			For $i = 1 To $BATs[0]
				_FileWriteLog($Log, "[INFO] Running script: " & $BATs[$i])
				$CurrentBAT = $BAT_Path & $BATs[$i]
				$RunBAT = RunWait(@ComSpec & " /c " & '"' & $CurrentBAT & '"', $BAT_Path, @SW_HIDE)
				If @error Then
					_FileWriteLog($Log, "[ERROR] An error occured when trying to run the script.")
					GUICtrlSetData($LABEL_StatusText, "Error: " & $BAT_Path & $BATs[$i])
					GUICtrlSetColor($LABEL_StatusText, $ErrorFontColor)
				Else
					_FileWriteLog($Log, "[INFO] Finished running script.")
				EndIf
			Next
		Else
			_FileWriteLog($Log, "[INFO] No scripts found.")
		EndIf


	#cs
		Local $EXE_Path = @ScriptDir & "\Executables\"
		Local $EXEs = _FileListToArray($EXE_Path, "*.EXE")
		If (Not IsArray($EXEs)) and (@Error=1) Then
			MsgBox (0,"","No Files\Folders Found.")
		Else
			For $i = 1 to $EXEs[0]
				Local $RunEXE = RunWait($EXE_Path & $EXEs[$i], $EXE_Path, @SW_HIDE)
				If @error Then
					GUICtrlSetData($LABEL_StatusText, "Error: " & $EXEs[$i])
					GUICtrlSetColor($LABEL_StatusText, $ErrorFontColor)
					Sleep(500)
				EndIf
			Next
		EndIf
	#ce
	EndFunc

#EndRegion	; > EXECUTE
#Region		;	REGISTRY

	Func _ImportRegFiles()
		_FileWriteLog($Log, "[INFO] Looking for Registration Entries to import...")
		Local $REG_Path = @ScriptDir & "\Registry\"
		Local $REGs = _FileListToArray($REG_Path, "*.REG")
		If IsArray($REGs) Then
			For $i = 1 to $REGs[0]
				_FileWriteLog($Log, "[INFO] Importing Registry Entry: " & $REGs[$i])
				Local $CurrentREG = $REG_Path & $REGs[$i]
				Local $RunREG = RunWait('regedit /s "' & $CurrentREG & '"')
				If @error Then
					_FileWriteLog($Log, "[ERROR] An error occured when trying to import Registry Entry.")
					GUICtrlSetData($LABEL_StatusText, "Error: " & $REG_Path & $REGs[$i])
					GUICtrlSetColor($LABEL_StatusText, $ErrorFontColor)
					Sleep(500)
				Else
					_FileWriteLog($Log, "[INFO] Successfully imported Registration Entries.")
				EndIf
			Next
		EndIf
	EndFunc

	Func _RemoveRegEntries()
		Local $GracePeriod, $DeleteGraceTimeBombValue, $SetOwnerOnGracePeriod, $GraceTimeBomb
		Local $REG_RDS_RCM = "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\RCM"
		Local $REG_GracePeriod = "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\RCM\GracePeriod"
		Local $REG_FW_POL_Services = $HKLM & "\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\RestrictedServices\Configurable\System"
		Dim $Permissions[1][3] = [[$AdminGroup, 1, $GENERIC_ALL]]

		_FileWriteLog($Log, "[INFO] Removing Registration Entries...")

		; RDS GRACE PERIOD TIMEBOMB
		_FileWriteLog($Log, "[INFO] Looking for Remote Desktop Grace Period data...")
		$GracePeriod = _RegEnumKeyEx($REG_RDS_RCM, 0, "GracePeriod")
		If Not IsArray($GracePeriod) Then
			_FileWriteLog($Log, "[ERROR] Could not find any Remote Desktop Grace Period data")
		Else
			_FileWriteLog($Log, "[INFO] Found Remote Desktop Grace Period data: " & $GracePeriod[1])
			$SetOwnerOnGracePeriod = _EditObjectPermissions($REG_GracePeriod, $Permissions, $SE_REGISTRY_KEY, "Administrators", 0, 1)
			If $SetOwnerOnGracePeriod = 0 Then
				_FileWriteLog($Log, "[ERROR] An error occured when trying to take ownership of object: " & $REG_GracePeriod)
				_FileWriteLog($Log, "[ERROR] _EditObjectPermissions return value: " & $SetOwnerOnGracePeriod)
			Else
				_FileWriteLog($Log, "[INFO] Taken ownership of Registry Key")
				_FileWriteLog($Log, "[INFO] Looking for Grace Period timebomb...")
				$GraceTimeBomb = _RegEnumValEx($REG_GracePeriod)
				If Not IsArray($GraceTimeBomb) Then
					_FileWriteLog($Log, "[ERROR] Could not find Remote Desktop Grace Period Timebomb")
				Else
					_FileWriteLog($Log, "[INFO] Found Remote Desktop Grace Period Timebomb: " & $GraceTimeBomb[1][1])
					_FileWriteLog($Log, "[INFO] Deleting value...")
					$DeleteGraceTimeBombValue = RegDelete($REG_GracePeriod, $GraceTimeBomb[1][1])
					If $DeleteGraceTimeBombValue Then
						_FileWriteLog($Log, "[INFO] Deleted Remote Desktop Grace Period Timebomb data")
					Else
						_FileWriteLog($Log, "[ERROR] An error occured when trying to delete: " & $REG_GracePeriod & "\" & $GraceTimeBomb[1][1])
					EndIf
				EndIf
			EndIf
		EndIf
	EndFunc

	Func _EditReg()
		Local $REG_FW_POL = $HKLM & "\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy"
		;DeleteUserAppContainersOnLogoff

	EndFunc

#EndRegion	; > REGISTRY
#Region		;	SHORTCUTS

	Func _CreateShortcuts()
		; See if they exist on the users desktop first
		; Use a folder that contains all the shortcuts. This way its possible to easily add shortcuts.
		; Create the ones that do now exist (current user, not public desktop)
		; Shortcuts: Change user

		; Check the integrity of the Shortcuts folder. Create new shortcuts if something is missing.
		; Make it possible to disable in the config.ini
	EndFunc

#EndRegion	; > SHORTCUTS
#Region		;	CLEANUP

	Func _Cleanup()
		Local $CleanupWindowsUpdate, $CleanupRecycleBin, $RecycleBinQuery
		Local $WindowsUpdateDataPath = "C:\Windows\SoftwareDistribution"
		Local $RecycleDrive = "C:"
		Local $RecycleItems = 0

		_FileWriteLog($Log, "[INFO] Running cleanup...")
		If $WindowsUpdateHasBeenRun Then
			_FileWriteLog($Log, "[INFO] Windows Update has been run... Not performing cleanup of the cache yet.")
		Else
			If $DeleteWindowsUpdateCache = "True" Then
				_FileWriteLog($Log, "[INFO] Deleting Windows Update data: " & $WindowsUpdateDataPath)
				$CleanupWindowsUpdate = DirRemove($WindowsUpdateDataPath, 1)
				If @error Then
					_FileWriteLog($Log, "[ERROR] An error occured while trying to delete this directory.")
				Else
					_FileWriteLog($Log, "[INFO] Successfully deleted this directory and all subdirectories and files.")
				EndIf
			EndIf
		Endif

		If $EmptyRecycleBin = "True" Then
			_FileWriteLog($Log, "[INFO] Emptying Recycle Bin...")
			$RecycleBinQuery = _WinAPI_ShellQueryRecycleBin($RecycleDrive)
			If IsArray($RecycleBinQuery) Then
				$RecycleItems = $RecycleBinQuery[1]
				_FileWriteLog($Log, "[INFO] Items in Recycle Bin: " & $RecycleBinQuery[1])
				If $RecycleItems > 0 Then
					FileRecycleEmpty($RecycleDrive)
					If @error Then
						_FileWriteLog($Log, "[ERROR] Failed to empty the Recycle Bin.")
					Else
						_FileWriteLog($Log, "[INFO] Successfully emptied the Recycle Bin.")
					EndIf
				Else
					_FileWriteLog($Log, "[INFO] No items in the Recycle Bin... Skipping.")
				EndIf
			Else
				_FileWriteLog($Log, "[INFO] No items in the Recycle Bin... Skipping.")
			EndIf
		EndIf

	EndFunc

#EndRegion	; > CLEANUP
#Region		;	FLUSH

	Func _FlushDNS()
		_FileWriteLog($Log, "[INFO] Flushing DNS and deleting ARP cache...")
		Local $FlushDNS = RunWait(@ComSpec & ' /c ' & 'netsh int ip delete arpcache & ipconfig /flushdns', @ScriptDir, @SW_HIDE)
		If @error Then
			_FileWriteLog($Log, "[ERROR] An error occured when trying execute these commands.")
		Else
			_FileWriteLog($Log, "[INFO] Successfully executed commands.")
		EndIf
	EndFunc

#EndRegion	; > FLUSH
#Region		;	LICENSE

	Func LicenseStatus()
		Local $sCommand = "cscript C:\Windows\System32\slmgr.vbs /DLV"
		Local $IsLicensed, $IsNotification, $RunSLMGR

		$RunSLMGR = _RunWaitGet(@ComSpec & " /c" & $sCommand, 1, "", @SW_HIDE)
		$IsLicensed = StringInStr($RunSLMGR, "License Status: Licensed")
			If $IsLicensed Then Return SetError(1, 1, "Licensed")
		$IsNotification = StringInStr($RunSLMGR, "License Status: Notification")
			If $IsNotification Then Return SetError(0, 1, "Unlicensed")

		Return SetError(0, 1, "N/A")
	EndFunc

	Func LicenseType()
		Local $sCommand = "cscript C:\Windows\System32\slmgr.vbs /DLV"
		Local $IsKMS, $IsMAK, $RunSLMGR

		$RunSLMGR = _RunWaitGet(@ComSpec & " /c" & $sCommand, 1, "", @SW_HIDE)
		$IsKMS = StringInStr($RunSLMGR, "VOLUME_KMSCLIENT")
			If $IsKMS Then Return SetError(1, 1, "KMS")
		$IsMAK = StringInStr($RunSLMGR, "VOLUME_MAK")
			If $IsMAK Then Return SetError(1, 1, "MAK")

		Return SetError(0, 1, "N/A")
	EndFunc

#EndRegion	; > LICENSE
#Region		;	PROCESSES

	Func Processes()
		Local $dotNETopt = "mscorsvw.exe"
		If ProcessExists($dotNETopt) Then
			_FileWriteLog($Log, "[INFO] Found running process: " & $dotNETopt)
			_FileWriteLog($Log, "[INFO] .NET optimization is running. Waiting for it to finish...")
			GUICtrlSetData($LABEL_StatusText, "Waiting for .NET optimization")
			ProcessWaitClose($dotNETopt)
			_FileWriteLog($Log, "[INFO] Optimization is finished.")
		EndIf
	EndFunc

#EndRegion	; > PROCESSES
#Region		;	PENDING

	Func PendingReboot()
		Local $RebootPending, $RebootRequired, $PendingFileRenameOperations
		; 1 = REBOOT PENDING

		$RebootRequired = _RegEnumKeyEx($HKLM & "\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", 0, "RebootRequired")
		$RebootPending = _RegEnumKeyEx($HKLM & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing", 0, "RebootPending")
		$PendingFileRenameOperations = RegRead($HKLM & "\SYSTEM\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations")
		If IsArray($RebootRequired) Then
			$WindowsUpdateHasBeenRun = True
			Return SetError(1, 1, "Windows Update")
		ElseIf IsArray($RebootPending) Then
			Return SetError(1, 2, "Component Based Servicing")
		ElseIf $PendingFileRenameOperations Then
			Return SetError(1, 3, "Pending File Rename Operation")
		EndIf

		Return SetError(0, 0, "No pending reboot")
	EndFunc

#EndRegion	; > PENDING
#Region		;	REPORTS

	Func _SystemHealthReport()

	EndFunc

	Func _ApplicationInstallsReport()
		Local $InstalledAppsPath32 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		Local $InstalledAppsPath64 = "HKLM64\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		Local $AppInfoFilter = "DisplayName|DisplayVersion|DisplayIcon|Publisher|InstallDate|InstallLocation|InstallSource|UninstallString"
		Local $ApplicationsInLatestReport, $ApplicationsReports, $ItemsToDelete, $LatestReport, $oFile, $hFileOpen
		Local $AppList, $AppInfo, $AppList_Full[0][2], $REG_Val, $REG_Data, $ArrayAdd, $NrOfInstalls = 0, $LatestInstallCount = 0
		Local $DisplayNameIndex, $DisplayName, $VersionIndex, $Version, $InstallDateIndex, $InstallDate, $PublishedIndex, $Publisher
		Local $UninstallIndex, $Uninstall, $DisplayIconIndex, $DisplayIcon, $DisplayIconAttrib, $InstallLocationIndex, $InstallLocation


		; LIST APPS FROM THE REGISTRY
		$AppList = _RegEnumKeyEx($InstalledAppsPath64)
		$ItemsToDelete = _RegEnumKeyEx($InstalledAppsPath64, "", "*.KB*")

		If Not IsArray($AppList) Then
			_FileWriteLog($Log, "[ERROR] Failed to list applications from the registry.")
		Else
			If IsArray($ItemsToDelete) Then
				$NrOfInstalls = $AppList[0] - $ItemsToDelete[0] ; GET NUMBER OF APPS MINUS KB FIXES
				For $i = 1 to UBound($ItemsToDelete) - 1 ; CLEAN UP THE APP-LIST ARRAY
					_ArrayDelete($AppList,_ArraySearch($AppList, $ItemsToDelete[$i]))
				Next
			Else
				$NrOfInstalls = $AppList[0]
			EndIf

			; CHECK FOR EXISTING REPORTS. RUN COMPARISON FROM LAST REPORT, IF FOUND.
			$ApplicationsReports = _FileListToArray($AppReportsFolder, "Applications_*", 1)
			If @error = 4 Then
				_FileWriteLog($Log, "[INFO] No Applications Report has been created yet... Creating the first one.")
			Else
				; FIND THE LATEST REPORT SO THAT WE CAN COMPARE AMOUNT OF UPDATES
				$LatestReport = _GetNewestFile($AppReportsFolder, 1)
				$ApplicationsInLatestReport = FileReadLine($LatestReport, 8)
				$LatestInstallCount = StringRegExpReplace($ApplicationsInLatestReport, '[^[:digit:]]', '')
			EndIf

			; CREATE NEW REPORT
			_FileWriteLog($Log, "[INFO] Creating a new Applications Report.")
			$oFile = $AppReportsFolder & "Applications_" & @MDAY & "." & @MON & "." & @YEAR & ".txt"
			If FileExists($oFile) Then
				_FileWriteLog($Log, "[INFO] Looks like one has already been created today... Creating a new report and overwriting the existing one.")
				FileDelete($oFile)
			Else
				_FileCreate($oFile)
			EndIf
			If @error Then
				_FileWriteLog($Log, "[ERROR] Failed to create report-file.")
			Else
				_FileWriteLog($Log, "[INFO] Created file: " & $oFile)
				$hFileOpen = FileOpen($oFile, $FO_APPEND)

				FileWrite($oFile, "====================================" & @CRLF)
				FileWrite($oFile, "======= Applications Summary =======" & @CRLF)
				FileWrite($oFile, "====================================" & @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, "Current Date: " & @TAB & @TAB & @MDAY & "." & @MON & "." & @YEAR & @CRLF)
				FileWrite($oFile, "Computer Name: " & @TAB & @TAB & @ComputerName  & @CRLF)
				FileWrite($oFile, "Apps Installed: " & @TAB & $NrOfInstalls & @CRLF)
				If $NrOfInstalls > $LatestInstallCount And $LatestInstallCount <> 0 Then FileWrite($oFile, "Newly Installed: " & @TAB & $NrOfInstalls - $LatestInstallCount & @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, "Currently Installed Applications")
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, @CRLF)


				; WRITE CONTENT TO REPORT
				For $i = 1 to UBound($AppList) - 1
					$AppInfo = _RegEnumValEx($AppList[$i], 256, $AppInfoFilter)
					_ArraySort($AppInfo, 0, 0, 0, 1)

					; CREATE VARIABLES
					; + DISPLAYNAME
					$DisplayNameIndex = _ArraySearch($AppInfo, "DisplayName")
					If $DisplayNameIndex >= 0 Then
						$DisplayName = $AppInfo[$DisplayNameIndex][3]
						FileWrite($oFile, $DisplayName & @CRLF)
					EndIf

					; + VERSION
					$VersionIndex = _ArraySearch($AppInfo, "DisplayVersion")
					If $VersionIndex >= 0 Then
						$Version = $AppInfo[$VersionIndex][3]
						FileWrite($oFile, "Version: " & $Version & @CRLF)
					EndIf

					; + DISPLAY ICON
					$DisplayIconIndex = _ArraySearch($AppInfo, "DisplayIcon")
					If $DisplayIconIndex >= 0 Then
						$DisplayIcon = $AppInfo[$DisplayIconIndex][3]
						If StringInStr($DisplayIcon, ",0") Then $DisplayIcon = StringTrimRight($DisplayIcon, 2)
						If StringInStr($DisplayIcon, '"') Then $DisplayIcon = StringReplace($DisplayIcon, '"', '')
						If StringInStr($DisplayIcon, "/") Then $DisplayIcon = StringReplace($DisplayIcon, '/', '\')
						If Not FileExists($DisplayIcon) Then $DisplayIcon = "N/A"
					EndIf

					; + INSTALL DATE
					$InstallDateIndex = _ArraySearch($AppInfo, "InstallDate")
					If $DisplayIcon <> "" Then
						$DisplayIconAttrib = FileGetTime($DisplayIcon, 2)
						If IsArray ($DisplayIconAttrib) Then $InstallDate = $DisplayIconAttrib[2] & "." & $DisplayIconAttrib[1] & "." & $DisplayIconAttrib[0]
						FileWrite($oFile, "Install Date: " & $InstallDate & @CRLF)
					ElseIf $InstallDateIndex > 0 Then
						$InstallDate = _Date_Time_Convert($AppInfo[$InstallDateIndex][3], 'yyyyMd', 'd.M.yyyy') ; CONVERT DATE FORMAT
						FileWrite($oFile, "Install Date: " & $InstallDate & @CRLF)
					Else
						$InstallDate = "N/A"
					EndIf
					$DisplayIcon = ""

					; + PUBLISHER
					$PublishedIndex = _ArraySearch($AppInfo, "Publisher")
					If $PublishedIndex >= 0 Then
						$Publisher = $AppInfo[$PublishedIndex][3]
						FileWrite($oFile, "Publisher: " & $Publisher & @CRLF)
					EndIf

					; + INSTALL LOCATION
					$InstallLocationIndex = _ArraySearch($AppInfo, "InstallLocation")
					If $InstallLocationIndex >= 0 Then
						$InstallLocation = $AppInfo[$InstallLocationIndex][3]
						If $InstallLocation <> "" Then
							FileWrite($oFile, "Install Location: " & $InstallLocation & @CRLF)
						EndIf
					EndIf

					; + UNINSTALL STRING
					$UninstallIndex = _ArraySearch($AppInfo, "UninstallString")
					If $UninstallIndex >= 0 Then
						$Uninstall = $AppInfo[$UninstallIndex][3]
						FileWrite($oFile, "Uninstall String: " & $Uninstall & @CRLF)
					EndIf

					If $DisplayNameIndex > 0 Then FileWrite($oFile, @CRLF)
				Next
				_FileWriteLog($Log, "[INFO] Successfully created report.")

			EndIf
		EndIf
	EndFunc

	Func _WindowsUpdateReport()
		Local $strQuery, $objWMIService, $colItems, $oFile, $hFileOpen, $sRegExpPatt
		Local $WindowsUpdateReports, $InstallDate, $HotFixID, $NrOfUpdates = 0, $NrOfNewUpdates = 0, $LatestReport
		Local $ArrayAdd, $UpdatesInLatestReport, $LatestUpdateCount = 0
		Local $DateArray[0][2]


		; WINDOWS UPDATE QUERY
		$strQuery = "winmgmts:" & "{impersonationLevel=impersonate}!\\" & @ComputerName & "\root\cimv2"
		$objWMIService = ObjGet($strQuery)
		$colItems = $objWMIService.ExecQuery("Select * from win32_QuickFixEngineering")

		; CHECK FOR EXISTING REPORTS. RUN COMPARISON FROM LAST REPORT, IF FOUND.
		$WindowsUpdateReports = _FileListToArray($UpdatesReportsFolder, "WindowsUpdates_*", 1)
		If @error = 4 Then
			_FileWriteLog($Log, "[INFO] No Windows Update Report has been created yet... Creating the first one.")
		Else
			; FIND THE LATEST REPORT SO THAT WE CAN COMPARE AMOUNT OF UPDATES
			$LatestReport = _GetNewestFile($UpdatesReportsFolder, 1)
			$UpdatesInLatestReport = FileReadLine($LatestReport, 8)
			$LatestUpdateCount = StringRegExpReplace($UpdatesInLatestReport, '[^[:digit:]]', '')

			_FileWriteLog($Log, "[INFO] Creating a new Windows Update Report.")
		EndIf

		; CREATE NEW REPORT
		$oFile = $UpdatesReportsFolder & "WindowsUpdates_" & @MDAY & "." & @MON & "." & @YEAR & ".txt"
		If FileExists($oFile) Then
			_FileWriteLog($Log, "[INFO] Looks like one has already been created today... Creating a new report and overwriting the existing one.")
			FileDelete($oFile)
		Else
			_FileCreate($oFile)
		EndIf

		If @error Then
			_FileWriteLog($Log, "[ERROR] Failed to create Windows Update Report.")
		Else
			If Not IsObj($colItems) Then
				_FileWriteLog($Log, "[ERROR] Failed to query hotfixes.")
				_FileWriteLog($Log, "[INFO] Deleting report-file.")
				FileDelete($oFile)
				If @error Then _FileWriteLog($Log, "[ERROR] Failed to delete file: " & $oFile)
				If Not @error Then _FileWriteLog($Log, "[INFO] Successfully deleted file: " & $oFile)
			Else
				_FileWriteLog($Log, "[INFO] Created file: " & $oFile)
				$hFileOpen = FileOpen($oFile, $FO_APPEND)

				$NrOfUpdates = $colItems.Count

				FileWrite($oFile, "====================================" & @CRLF)
				FileWrite($oFile, "====== Windows Update Summary ======" & @CRLF)
				FileWrite($oFile, "====================================" & @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, "Current Date: " & @TAB & @TAB & @MDAY & "." & @MON & "." & @YEAR & @CRLF)
				FileWrite($oFile, "Computer Name: " & @TAB & @TAB & @ComputerName  & @CRLF)
				FileWrite($oFile, "Updates Installed: " & @TAB & $NrOfUpdates & @CRLF)
				If $NrOfUpdates > $LatestUpdateCount And $LatestUpdateCount <> 0 Then FileWrite($oFile, "Newly Installed: " & @TAB & $NrOfUpdates - $LatestUpdateCount & @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, "Currently Installed Updates")
				FileWrite($oFile, @CRLF)
				FileWrite($oFile, @CRLF)

				; CREATE ARRAY
				For $QFE in $colItems
					$InstallDate = $QFE.InstalledOn
					$ArrayAdd = $InstallDate & "|" & $QFE.HotfixID
					_ArrayAdd($DateArray, $ArrayAdd)
				Next

				; SORT ARRAY BY DATE
				_ArraySort($DateArray, 1)

				; WRITE CONTENT TO REPORT
				For $i = 0 to UBound($DateArray) - 1
					$HotFixID = $DateArray[$i][1]
					$InstallDate = _Date_Time_Convert($DateArray[$i][0], 'M/d/yyyy', 'd.M.yyyy') ; CONVERT DATE FORMAT

					FileWrite($oFile, $HotFixID & @CRLF)
					FileWrite($oFile, "Install Date: " & $InstallDate & @CRLF)
					FileWrite($oFile, @CRLF)
				Next
				_FileWriteLog($Log, "[INFO] Successfully created report.")

			EndIf
		EndIf
	EndFunc     ; > _WindowsUpdateReport()

#Endregion	; > REPORTS
#Region		;	AUTOPROVISION

	Func _CreateScheduledTask()

	EndFunc

	Func _CreateProvisioningScript()

	EndFunc

#EndRegion	; > AUTOPROVISION
#Region		;	AUTOLOGON

	Func _AutoLogonEnable()
		Local $REG_AutoAdminLogon, $REG_DefaultUserName, $REG_DefaultDomain, $REG_DefaultPassword
		Local $REG_WinlogonPath = $HKLM & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
		Local $REG_RunOncePath = $HKLM & "\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
		Local $DeScrambledPassword

		$ScrambledPassword = IniRead($Config, "AutoLogon", "Password", "N/A")
		$DeScrambledPassword = _DeScramble($ScrambledPassword)

		; AUTO LOGON
		_FileWriteLog($Log, "[INFO] Collecting Auto Logon information...")
		_FileWriteLog($Log, "=======================================")
		_FileWriteLog($Log, "Auto Logon details")
		_FileWriteLog($Log, "AutoLogon Username: " & @TAB & $AutoLogonUserName)
		_FileWriteLog($Log, "AutoLogon Domain: " & @TAB & $AutoLogonDomain)
		_FileWriteLog($Log, "AutoLogon Password: " & @TAB & $ScrambledPassword)
		_FileWriteLog($Log, "AutoLogon Count: " & @TAB & @TAB & $AutoLogonCount)
		_FileWriteLog($Log, "=======================================")

		_FileWriteLog($Log, "[INFO] Writing Auto Logon details to the Registry: " & $REG_WinlogonPath)
		RegWrite($REG_WinlogonPath, "AutoAdminLogon", "REG_SZ", "1")
		If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to write 'AutoAdminLogon' to the Registry.")
		If Not @error Then _FileWriteLog($Log, "[INFO] Successfully configured 'AutoAdminLogon'.")

		RegWrite($REG_WinlogonPath, "DefaultUserName", "REG_SZ", $AutoLogonUserName)
		If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to write 'DefaultUserName' to the Registry.")
		If Not @error Then _FileWriteLog($Log, "[INFO] Successfully configured 'DefaultUserName'.")

		RegWrite($REG_WinlogonPath, "DefaultDomainName", "REG_SZ", $AutoLogonDomain)
		If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to write 'DefaultDomainName' to the Registry.")
		If Not @error Then _FileWriteLog($Log, "[INFO] Successfully configured 'DefaultDomainName'.")

		RegWrite($REG_WinlogonPath, "DefaultPassword", "REG_SZ", $DeScrambledPassword)
		If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to write 'DefaultPassword' to the Registry.")
		If Not @error Then _FileWriteLog($Log, "[INFO] Successfully configured 'DefaultPassword'.")

		RegWrite($REG_WinlogonPath, "AutoLogonCount", "REG_SZ", $AutoLogonCount)
		If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to write 'AutoLogonCount' to the Registry.")
		If Not @error Then _FileWriteLog($Log, "[INFO] Successfully configured 'AutoLogonCount'.")


		; RUN PREPSERVER ONCE AFTER AUTO LOGON
		RegWrite($REG_RunOncePath, 'PrepServer', 'REG_SZ', '"' & @ScriptFullPath & '"')

	EndFunc

	Func _AutoLogonDisable()
		Local $REG_WinlogonPath = $HKLM & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
		Local $REG_IsAutoLogonEnabled, $REG_IsAutoLogonPasswordCached, $REG_AutoLogonCounter

		GUICtrlSetData($LABEL_StatusText, "Checking to see if Auto Logon needs to be disabled")
		_FileWriteLog($Log, "[INFO] Checking to see if Auto Logon needs to be disabled...")
		$REG_IsAutoLogonEnabled = RegRead($REG_WinlogonPath, "AutoAdminLogon")
		$REG_IsAutoLogonPasswordCached = RegRead($REG_WinlogonPath, "DefaultPassword")
		$REG_AutoLogonCounter = RegRead($REG_WinlogonPath, "AutoLogonCount")

		If $REG_IsAutoLogonEnabled = "1" Or $REG_IsAutoLogonPasswordCached > "" Then
			GUICtrlSetData($LABEL_StatusText, "Disabling Auto Logon")
			_FileWriteLog($Log, "[INFO] Disabling Auto Logon in the Registry: " & $REG_WinlogonPath)
			RegWrite($REG_WinlogonPath, "AutoAdminLogon", "REG_SZ", "0")
			If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to edit 'AutoAdminLogon' in the Registry.")
			If Not @error Then _FileWriteLog($Log, "[INFO] Successfully set 'AutoAdminLogon = 0'.")

			RegDelete($REG_WinlogonPath, "DefaultPassword")
			If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to delete 'DefaultPassword' in the Registry.")
			If Not @error Then _FileWriteLog($Log, "[INFO] Successfully deleted 'DefaultPassword' from the Registry.")

			RegWrite($REG_WinlogonPath, "AutoLogonCount", "REG_SZ", "0")
			If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to edit 'AutoLogonCount' in the Registry.")
			If Not @error Then _FileWriteLog($Log, "[INFO] Successfully set 'AutoLogonCount = 0'.")
		Else
			_FileWriteLog($Log, "[INFO] Auto Logon already disabled.")
		EndIf
	EndFunc

#EndRegion	; > AUTOLOGON
#Region		;	CHANGE USER

	Func _ChangeUser($uFlag = 0)


	EndFunc

#EndRegion	; > CHANGE USER
#Region		;	SHUTDOWN & REBOOT

	Func _Shutdown()
		If $ServerType = "system" Then
			_Exit()
		Else
			If $Shutdown = "True" Then
				GUICtrlSetData($LABEL_StatusText, "Shutting down the system")
				_ClosePermissionResources()
				_AutoLogonDisable()
				_Crypt_Shutdown()
				_BlockInputEx(0)
				_Exit("shutdown")
			Else
				_FileWriteLog($Log, "[INFO] Shutdown has been disabled in the config: " & $Config)
				_FileWriteLog($Log, "[INFO] For automatic shutdown, set Shutdown=True")
				_Exit()
			EndIf
		EndIf
	EndFunc

	Func _Reboot()
		Local $Counter = $ExitTimer

		GUICtrlSetData($LABEL_Percent, "")
		GUICtrlSetState($LABEL_Percent, $GUI_HIDE)
		GUISetBkColor($BackgroundColor_Reset, $GUI)
		GUICtrlSetBkColor($LABEL_Percent, $LabelColor_Reset)
		GUICtrlSetBkColor($LABEL_Counter, $LabelColor_Reset)
		GUICtrlSetBkColor($LABEL_StatusText, $LabelColor_Reset)

		If $Reboot = "True" Then
			GUICtrlSetData($LABEL_StatusText, "Pending reboot detected!")
			_FileWriteLog($Log, "[INFO] Pending reboot detected... Reason: " & $PendingReboot)
			If $AutoLogonEnabled = "True" Then
				Sleep(500)
				GUICtrlSetData($LABEL_StatusText, "Pending reboot detected! - Preparing Auto Logon")
				_FileWriteLog($Log, "[INFO] Preparing Auto Logon.")
				_AutoLogonEnable()
			EndIf
			_FileWriteLog($Log, "[INFO] Sleeping for " & $ExitTimer & " secounds...")
			Do
				GUICtrlSetData($LABEL_Counter, $Counter)
				$Counter -= 1
				Sleep(1000)
			Until $Counter = 0
			Sleep(1000)
			_ClosePermissionResources()
			_Crypt_Shutdown()
			_BlockInputEx(0)
			_Exit("reboot")
		Else
			GUICtrlSetData($LABEL_StatusText, "Pending reboot detected! Exiting because reboot has been disabled.")
			_FileWriteLog($Log, "[INFO] Pending reboot detected! Exiting because reboot has been disabled.")
			Do
				GUICtrlSetData($LABEL_Counter, $Counter)
				$Counter -= 1
				Sleep(1000)
			Until $Counter = 0
			Sleep(1000)
			_Exit()
		EndIf
	EndFunc

#EndRegion	; > SHUTDOWN & REBOOT
#Region		;	EXIT

	Func _Abort()
		_FileWriteLog($Log, "[INFO] Hotkey pressed!")
		_Exit()
	EndFunc

	Func _Exit($vParam = "exit")
		Local $CMD = "cmd.exe"

		_FileWriteLog($Log, "[INFO] Exiting application...")
		_FileWriteLog($Log, "[INFO] Looking for running scripts...")
		If ProcessExists($CMD) Then
			_FileWriteLog($Log, "[INFO] Found running process: " & $CMD)
			ProcessClose($CMD)
			If @error Then
				_FileWriteLog($Log, "[ERROR] Failed to close process.")
			Else
				_FileWriteLog($Log, "[INFO] Successfully closed the process.")
			EndIf
		Else
			_FileWriteLog($Log, "[INFO] No running scripts found.")
		EndIf

		_BlockInputEx(0)
		_Crypt_Shutdown()
		_ClosePermissionResources()

		If $vParam = "shutdown" Then
			_AutoLogonDisable()
			_FileWriteLog($Log, "[END] Shutting down.")
			Shutdown(BitOR($SD_SHUTDOWN, $SD_FORCEHUNG))
				If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to perform shutdown.")
			Exit
		ElseIf $vParam = "reboot" Then
			_FileWriteLog($Log, "[END] Rebooting.")
			Shutdown($SD_REBOOT)
				If @error Then _FileWriteLog($Log, "[ERROR] An error occured when trying to perform reboot.")
			Exit
		Else
			_AutoLogonDisable()
			_FileWriteLog($Log, "[END]")
			Exit
		EndIf
	EndFunc

#EndRegion	; > EXIT


;===============================================================================================

;   LOCAL FUNCTIONS

#Region		;	PROGRESS

	Func __Progress($start, $end)
		Local $Percent = Random($start, $end, 1)
		GUICtrlSetData($LABEL_Percent, $Percent & "%")
	EndFunc

#EndRegion	; > PROGRESS


;===============================================================================================


; END