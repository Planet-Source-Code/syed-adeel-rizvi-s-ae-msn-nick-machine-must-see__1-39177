Attribute VB_Name = "IETOOLS"
' Add Reference to your Application in Internet Explorer's Tools Menu + Icon on Toolbar (VER. 5.0 or Higher)
' Copyright © 2000 Chuck DeLong
'******************************
' VERSION 2.0
'
'Further info on Browser Extensions can be found at...
'http://msdn.microsoft.com/workshop/browser/ext/overview/overview.asp
'
' Registry Functions By:
'*******************************************************************************
' Project:      General Functions
' Program:      Registry Functions
' Author:       V.A. van den Braken
' Version:      1.1
' Date:         30-07-1997, 02-08-1997
' Copyright:    Copyright © 1997 Deltec BV, Naarden
' Description:  Functions to access/modify/write the Windows Registry
'*******************************************************************************
'
' Menustat sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 1999 by Matt Hart
'
' *IMPORTANT*
' Make sure you compile into an exe, then run the exe (Running in design mode will reference IETOOLS.vbp instead of SampleApp.exe and MSIE will not find an *.exe to run!)
' A new instance of MSIE is required for changes to be seen (Add or Delete)
'
Option Explicit
' Shlwapi.dll (MSIE Version Info) (All required...)
Type DllVersionInfo
cbSize As Long
dwMajorVersion As Long '...But the only one we need
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type

Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long

Dim IEMV As DllVersionInfo
Dim CheckReg As String
Dim GetIEMajor As String
Dim Hico As String
Dim Ico As String
Dim Prog As String

Public Function DetectIE()
'See Remarks in Private Sub Form_Load()
CheckReg = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE", "")
IEMV.cbSize = Len(IEMV)
Call DllGetVersion(IEMV)
GetIEMajor = IEMV.dwMajorVersion
If Dir(CheckReg) = "" Or GetIEMajor < 5 Then

Else
CheckReg = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "CLSID")
If CheckReg = "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}" Then

Else

End If
End If
End Function

Public Function mnuAddIE()
' Path of yor App and HotIcon
Hico = "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Icons\Win95\MYCOMP.ICO"
' Path of yor App and Icon
Ico = "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Icons\Win95\NETHOOD.ICO"
' Path of yor App and Apps *.exe name
Prog = "C:\Documents and Settings\AdministratorMaster\Desktop\S.AS Nick Machine.exe"
' Adds your App to MSIE's Tools Menu and adds an Icon on the Toolbar
' {ECC5777A-6E88-BFCE-13CE-81F134789E7B} Any GUID
' Your App (Tools Menu Button Text)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "ButtonText", "S.AS Nick Machine"
' {1FBA04EE-3024-11D2-8F1F-0000F87ABD16} MUST BE THIS GUID
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
' Show Icon if IE Toolbar is  reset
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Default Visible", "Yes"
' Your APP Path and Name (Prog)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Exec", Prog
' Colered icon (Hico)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "HotIcon", Hico
' Default icon (Ico)
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "Icon", Ico
'Statusbar text for your App
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "MenuStatusBar", "YS.AS Nick Machine"
'Tools Menu text for your APP
REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}", "MenuText", "&S.AS Nick Machine"

End Function

Public Function mnuDeleteIE()
' Deletes your App in MSIE's Tools Menu and the Icon on the Toolbar
REGDeleteSetting vHKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Extensions\{ECC5777A-6E88-BFCE-13CE-81F134789E7B}"

End Function
