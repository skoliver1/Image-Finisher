#SingleInstance force
#NoEnv

D := SubStr(A_ScriptFullPath, 1, 1)
version := "v2022.02.09"
Title := "Image Finisher " version
SetTitleMatchMode Slow
MyDir := D ":\Files"

if !(A_IsAdmin)
{
   Run *RunAs %A_AhkPath% "%A_ScriptFullPath%"
   ExitApp
}

IfWinActive ahk_exe cmd.exe
	WinMinimize ahk_exe cmd.exe

RegRead InitialRun, HKCU, Software\OSteve Productions, InitialRun
if (InitialRun = "")
{
	gosub GetSystemInfo
	
	If InStr(SystemInfo, "7240")
		Model = 7240
	else If InStr(SystemInfo, "OptiPlex 9020")
		Model = 9020
	else If InStr(SystemInfo, "Latitude E7250")
		Model = 7250
	else If InStr(SystemInfo, "Folio 1040")
		Model = 1040
	else If InStr(SystemInfo, "T1600")
		Model = 1600
	else If InStr(SystemInfo, "T1650")
		Model = 1650
	else If InStr(SystemInfo, "800 G1 DM")
		Model = 800g1
	else If InStr(SystemInfo, "800 G2 DM")
		Model = 800g2
	else If InStr(SystemInfo, "800 G3 DM")
		Model = 800g3
	else If InStr(SystemInfo, "800 G4 DM")
		Model = 800g4
	else If InStr(SystemInfo, "HP EliteBook x360 1030 G2")
		Model = 360G2

	If !InStr(SystemInfo, "Windows 10")
	{
		if (Model = "")
		{
			MsgBox, 16, Image Finisher %version%, This computer does not appear to be compatible and will exit.`n`nGoodbye
			ExitApp
		}
	}
}

RegRead Place, HKCU, Software\OSteve Productions, LastDone
if (Place = "")
{
	MsgBox, 68, Image Finisher %version%, `*`*Supported computers that will receive a BIOS update (if necessary):`nHP 1040 G2`, 1040 G3`, 800 G1`, 800 G2`, 800 G3`, 800 G4`nx360 1030 G2`nDell 9020`, 7240`, 7250`n`nThis program will automate the setup of any computer intended for resale after imaging`, as long as it supports the latest Windows 10 image.  Older OS versions will only work for specific computer models.`nPress CapsLock at any point to kill the script.`n`nOnce you've started`, you may need to finish up a BIOS update and eventually`, shut the computer down.  So once we start`, leave the mouse and keyboard alone unless prompted.`n`n`nReady to start?
	IfMsgBox Yes
	{
		FileCreateShortcut, %D%:\RunMe.bat, %A_Startup%\Shortcut.lnk, %D%:\,,,,,, 7 ; make this run at startup
		FileCreateShortcut, %D%:\RunMe.bat, %A_Desktop%\Shortcut.lnk, %D%:\,,,,,, 7
		RegWrite REG_MULTI_SZ, HKCU, Software\OSteve Productions, SystemInfo, %SystemInfo%
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, InitialRun, 1
		RegWrite REG_DWORD, HKLM, SYSTEM\CurrentControlSet\Control\Session Manager\Power, HiberbootEnabled, 0 ; disable fast startup
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 1
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, Model, %Model%
		gosub 1-Disable_UAC
	}
	else
		ExitApp
}
RegRead Model, HKCU, Software\OSteve Productions, Model

if (SystemInfo = "")
{
	RegRead SystemInfo, HKCU, Software\OSteve Productions, SystemInfo
	if (SystemInfo = "")
		gosub GetSystemInfo
}


SetWorkingDir %MyDir%
if (Place = 1)
	gosub 1-Disable_UAC

if (Place = 2)
	gosub 2-Drivers

if (Place = 3)
	gosub 3-BIOS

if (Place = 4)
	gosub 4-Activate

if (Place = 5)
	gosub Cleanup




GetSystemInfo:
if A_OSVersion in WIN_7,WIN_8,WIN_8.1,WIN_VISTA,WIN_XP ; this is not an expression supported line
{
	Run cmd.exe
	WinWaitActive ahk_exe cmd.exe
	SendInput SystemInfo > %A_Temp%\OSName.txt{Enter}
	Sleep 500
	Process WaitClose, SystemInfo.exe
	WinClose ahk_exe cmd.exe
	FileRead SystemInfo, %A_Temp%\OSName.txt
}
else
{
	Run cmd.exe /c SystemInfo.exe > %A_Temp%\OSName.txt,, hide
	Sleep 500
	Process WaitClose, SystemInfo.exe
	FileRead SystemInfo, %A_Temp%\OSName.txt
}

; required files check
If !InStr(SystemInfo, "Windows 7")
{
	if !FileExist(MyDir . "\ProduKey.exe")
	{
		MsgBox The activation file ProduKey.exe is missing.`nPlease download it from http://nirsoft.net and place inside the "\Files" folder.`nThen run the script again.
		ExitApp
	}
}
return


1-Disable_UAC:
RunWait powershell.exe New-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -PropertyType DWord -Value 0 -Force ; disables UAC
RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 2

If InStr(SystemInfo, "Windows 8")
	RegWrite, REG_DWORD, HKCU, Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage, OpenAtLogon, 0 ; if Windows 8, boot straight to the desktop, instead of Start Screen
Sleep 1000
Run cmd.exe /c shutdown /r /f /t 0,, Hide
ExitApp


2-Drivers:
Sleep 5000
Run caffeine.exe, %MyDir%\bin
If InStr(SystemInfo, "Windows 10")
{
	Loop, Files, %D%:\*, DR ; finding the SXS folder path
	{
		if (A_LoopFileName != "sxs")
			continue
		else 
		{
			Num =
			Loop, Files, %A_LoopFileFullPath%\*, D ; find how many directories in sxs folder
				Num := A_Index
			If (Num = "") ; if there's none
			{
				Run cmd.exe
				Sleep 1000
				Send DISM /Online /Enable-Feature /FeatureName:NetFx3 /LimitAccess /Source:"%A_LoopFileFullPath%"{Enter}
				Sleep 2000
				Process WaitClose, dism.exe, 300000
				WinWaitActive ahk_exe cmd.exe
				Send exit{enter}
			}
		}
	}
	
	Gosub, 3-BIOS
}

If InStr(SystemInfo, "Windows 7")
{
	DriverPath = 
	if (Model = "9020")
		DriverPath = %MyDir%\Dell-9020\9020-7
	if (Model = "1600" or Model = "1650")
		DriverPath = %MyDir%\Dell-1600 1650\7

	if (DriverPath)
	{
		DriverCommand = Get-ChildItem "%DriverPath%" -Recurse -Filter "*.inf" | ForEach-Object { PNPUtil.exe -i -a $_.FullName }
		Run Powershell.exe %DriverCommand%,, Max
		Sleep 3000
		Loop
			{
				Process Exist, powershell.exe
				if (ErrorLevel != 0)
				{
					If WinActive("Windows Security", "anyway")
						Send i
					If WinActive("Windows Security", "should only install")
						Send i
					If WinActive("Windows Update", "10 minutes")
						Send p
					Sleep 1000
					continue
				}
				else
				break
			}
	}
	Run %A_AhkPath% "%MyDir%\bin\Postpone.ahk"
	Loop, Files, %MyDir%\IE11\*.msu, F
	{
		SplashTextOn, 450, 80, %Title%, Installing some updates in the background. Please be patient.`nThis will take about an hour.`nRunning %A_Index% of 9 files.
		RunWait %A_LoopFileFullPath% /quiet /norestart
	}
	SplashTextOn, 450, 50, %Title%, Installing IE11 in the background. Please be patient.
	RunWait %MyDir%\IE11\IE11-Windows6.1-x64-en-us.exe /quiet /norestart
	SplashTextOff
}


RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 3

If InStr(SystemInfo, "Windows 7")
{
	Sleep 30000
	Run cmd.exe /c shutdown /r /f /t 0,, Hide
}
ExitApp



3-BIOS: ; Update BIOS
If InStr(SystemInfo, "Windows 7")
{
	Sleep 30000
	If WinExist("Microsoft Windows", "&Restart Now")
	{
		WinActivate Microsoft Windows, &Restart Now
		Send {Enter}
		ExitApp
	}
}
	

SetTitleMatchMode Fast
if (Model = "7240")
{
	If InStr(SystemInfo, "Dell Inc. A27")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	Run E7240A27.exe /s /r /l="%A_Temp%\BIOS-Update.log", %MyDir%\Dell-7240
	Gosub, myGui
	; WinWait BIOS Update Program, utility will update the system
	; WinActivate BIOS Update Program, utility will update the system
	; Send {Enter}
	; WinWait Confirm BIOS Replacement, Do you wish
	; WinActivate Confirm BIOS Replacement, Do you wish
	; Send {Enter}
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if (Model = "9020")
{
	If InStr(SystemInfo, "Dell Inc. A25")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	Run O9020A25.exe /s /r /l="%A_Temp%\BIOS-Update.log", %MyDir%\Dell-9020
	Gosub, myGui
	; WinWait BIOS Update Program, utility will update the system
	; WinActivate BIOS Update Program, utility will update the system
	; Send {Enter}
	; WinWait Confirm BIOS Replacement, Do you wish
	; WinActivate Confirm BIOS Replacement, Do you wish
	; Send {Enter}
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if (Model = "7250")
{
	If InStr(SystemInfo, "Dell Inc. A21")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	Run E7250A21.exe /s /r /l="%A_Temp%\BIOS-Update.log", %MyDir%\Dell-7240
	Gosub, myGui
	; WinWait BIOS Update Program, utility will update the system
	; WinActivate BIOS Update Program, utility will update the system
	; Send {Enter}
	; WinWait Confirm BIOS Replacement, Do you wish
	; WinActivate Confirm BIOS Replacement, Do you wish
	; Send {Enter}
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if InStr(SystemInfo, "Folio 1040 G2")
{
	If InStr(SystemInfo, "Ver. 01.19")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	; MsgBox, 48, Update BIOS, The BIOS needs to be updated.  Press OK and I'll start it for you and you'll need to finish it.  `n`nAfter the restart`, I'll take care of the rest.
	Gosub, myGui
	Run HPBIOSUPDREC.exe -s -fM76_0119.bin -pPass.bin, %MyDir%\HP-1040\G2\BIOS
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if InStr(SystemInfo, "Folio 1040 G3")
{
	If InStr(SystemInfo, "Ver. 01.52")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	; MsgBox, 48, Update BIOS, The BIOS needs to be updated.  Press OK and I'll start it for you and you'll need to finish it.  `n`nAfter the restart`, I'll take care of the rest.
	Gosub, myGui
	Run HPBIOSUPDREC.exe -s -fN83_0152.bin -pPass.bin, %MyDir%\HP-1040\G3\BIOS
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if (Model = "800g1")
{
	if InStr(SystemInfo, "v02.33")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	Run HpqFlash.exe -fROM.CAB -pPass.bin -s -h, %MyDir%\HP-800\G1\BIOS
	Gosub, myGui
	proc := "HpqFlash.exe"
	Gosub, procWait
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	Run shutdown.exe /r /f /t 00
	ExitApp
}
else if (Model = "800g2")
{
	if InStr(SystemInfo, "Ver. 02.53")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	; MsgBox, 48, Update BIOS, The BIOS needs to be updated.  Press OK and I'll start it for you and you'll need to finish it.  `n`nAfter the restart`, I'll take care of the rest.
	Gosub, myGui
	Run HPBIOSUPDREC.exe -s -fN21_0253.bin -pPass.bin, %MyDir%\HP-800\G2\BIOS
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if (Model = "800g3")
{
	if InStr(SystemInfo, "Ver. 02.38")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	; MsgBox, 48, Update BIOS, The BIOS needs to be updated.  Press OK and I'll start it for you and you'll need to finish it.  `n`nAfter the restart`, I'll take care of the rest.
	Gosub, myGui
	Run HPBIOSUPDREC.exe -s -fP21_0238.bin -pPass.bin, %MyDir%\HP-800\G3\BIOS
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if (Model = "800g4")
{
	if InStr(SystemInfo, "Ver. 02.16.00")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	; MsgBox, 48, Update BIOS, The BIOS needs to be updated.  Press OK and I'll start it for you and you'll need to finish it.  `n`nAfter the restart`, I'll take care of the rest.
	Gosub, myGui
	Run HpFirmwareUpdRec.exe -s -fQ21_021600.bin -pPass.bin, %MyDir%\HP-800\G4\BIOS
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp
}
else if (Model = "360G2")
{
	if InStr(SystemInfo, "P80 Ver. 01.39")
	{
		RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
		gosub 4-Activate
	}
	Gosub, myGui
	Run HPBIOSUPDREC.exe -s -pPass.bin -fP80_0139.bin, %MyDir%\HP-x360\G2\BIOS 1.39
	RegWrite REG_SZ, HKCU, Software\OSteve Productions, LastDone, 4
	WinWaitClose, Notification
	ExitApp

}
gosub 4-Activate


4-Activate: ; activate Windows
Sleep 10000
SetKeyDelay 50
Run %MyDir%\bin\caffeine.exe

gosub pingLoop
If InStr(PingResult, "Reply from")
	gosub EnterKey
else
{
	MsgBox,,%Title%, Please connect the PC to the internet to continue.`nThen Press OK.
	gosub pingLoop
	gosub EnterKey
}


pingLoop:
Loop 20
{
	Sleep 1000
	FileDelete, %A_Temp%\PingResult.txt
	RunWait, cmd /c ping google.com -n 1 >%A_Temp%\PingResult.txt,, hide
	FileRead, PingResult, %A_Temp%\PingResult.txt

	If InStr(PingResult, "Destination host unreachable")
		continue
	If InStr(PingResult, "Reply from")
		break
	Else
		continue
}
return


EnterKey:
If InStr(SystemInfo, "Windows 7")
{
	Run slui 3
	Activated = No
	gosub Cleanup
}
else
{
	SetTitleMatchMode 2
	Run ProduKey.exe
	Sleep 3000
	WinWait ProduKey
	WinActivate ProduKey
	WinWaitActive ProduKey
	Send windows (bios^k
	Sleep 1000
	Key = %Clipboard%
}

Loop
{
	StringUpper Key, Key
	Run cmd.exe /c slmgr.vbs -ipk %Key%,, Hide ; add license
	WinWait Windows Script Host, successfully, 90
	if (ErrorLevel)
	{
		If WinExist("Windows Script Host", "invalid")
		{
			WinActivate Windows Script Host, invalid
			Send {Enter}
			InputBox Key, %Title%, It looks like the License Key didn't take.`nPlease confirm/re-enter the key.,,,,,,,, %Key%
			If (ErrorLevel = 1)
			{
				MsgBox You pressed Cancel.  To re-attempt to activate Windows`, either re-launch the script manually`,restart the computer or manually activate.
				ExitApp
			}
			Sleep 2000
			continue
		}
		else if WinExist("Windows Script Host", "slui.exe 0x2a")
		{
			Send {Enter}
			Run slui.exe
			WinWait Enter a product key
			Sleep 2000
			WinActivate Enter a product key
			Sleep 2000
			Send ^v{Enter}
			WinWait Activate Windows
			WinActivate Activate Windows
			Sleep 3000
			Send {Enter}
			WinWait, Windows is activated,, 90
			if (ErrorLevel)
				{
					MsgBox,262144, %Title%, I didn't detect that Windows accepted the new key.`nYou'll need to activate Windows manually.  Peace out.
					Activated = No
					gosub Cleanup
				}
			WinActivate, Windows is activated
			Send !{F4}
			Sleep 10000
			gosub Expiration
		}
		else
		{
			MsgBox,262144, %Title%, I didn't detect that Windows accepted the new key.`nYou'll need to activate Windows manually.  Peace out.
			Activated = No
		}
		gosub Cleanup
	}
	WinActivate Windows Script Host, successfully
	WinWaitActive Windows Script Host, successfully
	break
}
Sleep 2000
Send {Enter}
Sleep 20000
gosub Activation

Activation:
Run cmd.exe /c slmgr.vbs -ato,, Hide ; activate license
WinWait Windows Script Host, Product activated successfully, 90
if (ErrorLevel)
{
	MsgBox,262144, %Title%, I didn't detect that Windows activated the new key.`nYou'll need to activate Windows manually.  Peace out.
	Activated = No
	gosub Cleanup
}
WinActivate Windows Script Host, Product activated successfully
WinWaitActive Windows Script Host, Product activated successfully
Sleep 2000
Send {Enter}
Sleep 20000
gosub Expiration

Expiration:
Run cmd.exe /c slmgr.vbs -xpr,, Hide ; check the expiration
WinWait Windows Script Host, permanently activated, 90
if (ErrorLevel)
{
	MsgBox,262144, %Title%, I didn't detect that Windows is activated.`nYou'll need to activate Windows manually.  Peace out.
	Activated = No
	gosub Cleanup
}
WinActivate Windows Script Host, permanently activated
WinWaitActive Windows Script Host, permanently activated
Sleep 2000
Send {Enter}
gosub Cleanup


Cleanup:
RunWait powershell.exe New-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -PropertyType DWord -Value 1 -Force ; re-enables UAC
FileDelete %A_Startup%\*.lnk
FileDelete %A_Desktop%\Shortcut.lnk

RegDelete HKCU, Software\OSteve Productions
if (Activated = "No")
	MsgBox,, %Title%, Done.  Go ahead and shut down once you've activated Windows.
else
	MsgBox,, %Title%, Done.  Go ahead and shut down.
ExitApp

CapsLock::
ExitApp

myGui:
Gui, Font, s20
Gui,Add,Text,x19 y14 w350 h150 Center, The BIOS is silently being updated in the background and the computer will restart automatically.
Gui,Show,x1050 y469 w391 h175 , Notification
Return

GuiClose:
Gui, Destroy
Return

procWait:
Loop
{
	Sleep 1000
	Process, Exist, %proc%
	If (ErrorLevel != 0)
		Continue
	Break
}
Return