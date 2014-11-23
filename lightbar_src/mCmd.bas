Attribute VB_Name = "mCmd"
'################################################'
'# Programm:                           LightBar #'
'# Part:               Parser Internal Commands #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             26.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Long, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'original
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long '??????
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function RtlAdjustPrivilege Lib "ntdll" (ByVal a1 As Integer, ByVal a2 As Boolean, ByVal a3 As Boolean, ByRef a4 As Boolean) As Boolean
Private Declare Function ZwShutdownSystem Lib "NTdll.dll" (ByVal eFlag As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long

Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwnd As Long, ByVal sConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long

Public Sub CommandParser(ByVal wCmd1 As String, ByVal wCmd2 As String, ByVal wCmd3 As String)
wCmd1 = Right$(wCmd1, Len(wCmd1) - 5)

If wCmd1 = "shutdown" Then
  If wCmd2 = "shutdown" Then Call Shutdown(1, wCmd3)
  If wCmd2 = "reboot" Then Call Shutdown(2, wCmd3)
  If wCmd2 = "logoff" Then Call Shutdown(3, wCmd3)
End If
If wCmd1 = "cdrom" Then
  If wCmd3 = "open" Then Call OpenCloseCD(1, wCmd2)
  If wCmd3 = "close" Then Call OpenCloseCD(2, wCmd2)
  If wCmd3 = "open/close" Then Call OpenCloseCD(3, wCmd2)
End If
If wCmd1 = "window" Then
  If wCmd2 = "close" Then
    If wCmd3 = "close" Then Call Window_CloseWindow(1)
    If wCmd3 = "quit" Then Call Window_CloseWindow(2)
  End If
  If wCmd2 = "topmost" Then
    If wCmd3 = "top" Then Call Window_TopMost(1)
    If wCmd3 = "notop" Then Call Window_TopMost(2)
  End If
  If wCmd2 = "transparent" Then Call Window_SetTrans(Val(wCmd3))
  If wCmd2 = "transform" Then Call Window_Transform(wCmd3)
End If
If wCmd1 = "sound" Then
  If wCmd2 = "up" Then Call SetVol(1, Val(wCmd3))
  If wCmd2 = "down" Then Call SetVol(2, Val(wCmd3))
End If
If wCmd1 = "winamp" Then
  If wCmd3 = "run" Then Call WinampControl(wCmd2, 1) Else Call WinampControl(wCmd2)
End If
If wCmd1 = "net" Then
  If wCmd2 = "connect" Then Call NetConnect(wCmd3)
End If
If wCmd1 = "other" Then
  If wCmd2 = "datetime" Then Call Other_ShowDateTime(wCmd3)
  If wCmd2 = "datetime paste" Then Call Other_ShowDateTime(wCmd3, 1)
  If wCmd2 = "clear clipboard" Then Clipboard.Clear
  If wCmd2 = "extract usb" Then Call ShellExecute(0&, "open", GetDir("%systemroot%\system32\rundll32.exe"), "shell32.dll,Control_RunDLL hotplug.dll", "", 1)
End If
End Sub

Private Sub Shutdown(ByRef wAct As Byte, ByRef wForce As String)
Dim Flag As Boolean
If wAct = 1 Then
  If wForce <> "force" Then
    Call ShellExecute(0&, "open", "shutdown", "-s -t 00", "", 1)
  Else
    RtlAdjustPrivilege 19, True, False, Flag
    ZwShutdownSystem 2
  End If
End If
If wAct = 2 Then
  If wForce <> "force" Then
    Call ShellExecute(0&, "open", "shutdown", "-r -t 00", "", 1)
  Else
    RtlAdjustPrivilege 19, True, False, Flag
    ZwShutdownSystem 1
  End If
End If
If wAct = 3 Then
  If wForce <> "force" Then
    Call ExitWindowsEx(0, 0&)
  Else
    Call ExitWindowsEx(4, 0&)
  End If
End If
End Sub

Private Sub OpenCloseCD(ByRef wOpenClose As Byte, ByRef wDrive As String)
Dim T As Double
If wOpenClose = 1 Then 'open
  Call mciSendString("Open " & wDrive & ":/: Alias vv" & wDrive & ":/ Type CDAudio ", 0, 0, 0)
  Call mciSendString("Set vv" & wDrive & ":/ Door Open", 0, 0, 0)
End If
If wOpenClose = 2 Then 'close
  Call mciSendString("Open " & wDrive & ":/: Alias vv" & wDrive & ":/ Type CDAudio ", 0, 0, 0)
  Call mciSendString("Set vv" & wDrive & ":/ Door Closed", 0, 0, 0)
End If
If wOpenClose = 3 Then 'open/close
  Call mciSendString("Open " & wDrive & ":/: Alias vv" & wDrive & ":/ Type CDAudio ", 0, 0, 0)
  'snachala probuem otkryt'
  DoEvents: T = Timer: DoEvents
  Call mciSendString("Set vv" & wDrive & ":/ Door Open", 0, 0, 0)
  DoEvents
  If Timer - T < 0.5 Then 'esli operaciya otkrytiya dlilas' men'she polsekundy, znachit eyo _
                                                          voobcshe nebylo, znachit nada zakryvat' cd
    Call mciSendString("Set vv" & wDrive & ":/ Door Closed", 0, 0, 0)
  End If
End If
End Sub

Private Sub Window_CloseWindow(ByRef wMethod As Byte)
If wMethod = 1 Then Call PostMessage(GetLastHWnd, &H10, 0&, 0&) 'close
If wMethod = 2 Then Call PostMessage(GetLastHWnd, &H12, 0&, 0&) 'quit (kill)
End Sub

Private Sub Window_TopMost(ByRef wTop As Byte)
If wTop = 1 Then Call SetWindowPos(GetLastHWnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2) 'top
If wTop = 2 Then Call SetWindowPos(GetLastHWnd, 1, 0, 0, 0, 0, &H10 Or &H1 Or &H2) 'notop
End Sub

Private Sub Window_SetTrans(ByVal wL As Long)
If wL < 0 Then wL = 0
If wL > 255 Then wL = 255
Call SetTransparent(GetLastHWnd, CByte(wL))
End Sub

Private Sub Window_Transform(ByRef wAct As String)
If wAct = "maximize" Then
  If IsZoomed(GetLastHWnd) = 0 Then 'esli NE maksimizirovanno
    Call PostMessage(GetLastHWnd, &H112, &HF030&, 0&) 'max
  Else
    Call PostMessage(GetLastHWnd, &H112, &HF120&, 0&) 'rest
  End If
End If
If wAct = "minimize" Then
  If IsIconic(GetLastHWnd) = 0 Then 'esli NE minimizirovanno
    Call PostMessage(GetLastHWnd, &H112, &HF020&, 0&) 'min
  Else
    Call PostMessage(GetLastHWnd, &H112, &HF120&, 0&) 'rest
  End If
End If
'If wAct = "minimize" Then Call PostMessage(GetLastHWnd, &H112, &HF020&, 0&)
End Sub

Private Sub SetVol(ByRef wAct As Byte, ByRef wPrc As Byte)
Static TV As String  'tekucshij uroven' zvuka
Static TP As Integer 'tekucshij procent zvuka
Dim T1 As String, T2 As String 'temp
Dim wP As Long
If TV <> SetVol_GetVol Then
  TV = SetVol_GetVol
  'uznaem tekucshij procent zvuka
  T1 = "&H" & Mid$(TV, 1, 4)
  T2 = "&H" & Mid$(TV, 5, 4)
  If T1 > T2 Then TP = T1 / 655.35 Else TP = T2 / 655.35
End If
If wAct = 1 Then TP = TP + wPrc
If wAct = 2 Then TP = TP - wPrc
If TP > 100 Then TP = 100
If TP < 0 Then TP = 0
wP = CDbl(65535 / 100) * TP
Dim rV As Double
rV = CDbl(CDbl(wP) * CDbl(65536)) + wP
If rV > 2147483648# Then
  rV = rV - 4294967296#
End If
Call waveOutSetVolume(0, rV)
Call mPrg.GetStatus(2, CStr(TP))
End Sub

Private Function SetVol_GetVol() As String
Dim VV As Long
Call waveOutGetVolume(0, VV)
SetVol_GetVol = Hex(VV)
1000 If Len(SetVol_GetVol) < 8 Then SetVol_GetVol = "0" & SetVol_GetVol: GoTo 1000
End Function

Private Sub WinampControl(ByRef wCmd As String, Optional ByRef wStart As Byte = 0)
Dim waHWnd As Long
Dim waPath As String 'puti do wimampa
Dim sT As Double
Dim SERet As Long

waHWnd = FindWindow("Winamp v1.x", vbNullString)

If waHWnd = 0 Then
  If wStart = 0 Then
    Exit Sub
  Else
    'popytka zapustit' winamp
    waPath = GetRegString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\Winamp", "UninstallString")
    If waPath = "" Then
      waPath = GetRegString(HKEY_CLASSES_ROOT, "Applications\winamp.exe\shell\open\command", "")
      If waPath = "" Then
        waPath = GetRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\Applications\winamp.exe\shell\open\command", "")
        If waPath = "" Then
          Call fMsg.GetMsg(fPrg, 0, MapMsg(48))
          Exit Sub
        End If
      End If
    End If
    waPath = Replace$(waPath, """", "")
    waPath = GetFolder$(waPath)
    SERet = ShellExecute(0&, "open", waPath & "winamp.exe", "", waPath, 1)
    If SERet > -1 And SERet < 33 Then
      waPath = waPath & "winamp.exe"
      Call fMsg.GetMsg(fPrg, 0, MapMsg(49) & " " & waPath)
      Exit Sub
    End If
  End If
End If

'zhdem 20 sekund do poyavleniya winampa
sT = Timer: DoEvents
Do
  waHWnd = FindWindow("Winamp v1.x", vbNullString)
  If waHWnd > 0 Then Exit Do
  If Timer - sT > 20 Then
    waPath = waPath & "winamp.exe"
    Call fMsg.GetMsg(fPrg, 0, MapMsg(49) & " " & waPath)
    Exit Sub
  End If
  DoEvents
Loop

If wCmd = "back" Then
  Call SendMessage(waHWnd, &H111, 40044, vbNull): Call Sleep(111)
  Call mPrg.GetStatus(3, CStr(WinampControl_GetTrackName(waHWnd)))
End If
If wCmd = "play" Then Call SendMessage(waHWnd, &H111, 40045, vbNull)
If wCmd = "pause" Then Call SendMessage(waHWnd, &H111, 40046, vbNull)
If wCmd = "stop" Then Call SendMessage(waHWnd, &H111, 40047, vbNull)
If wCmd = "next" Then
  Call SendMessage(waHWnd, &H111, 40048, vbNull): Call Sleep(111)
  Call mPrg.GetStatus(3, CStr(WinampControl_GetTrackName(waHWnd)))
End If
If wCmd = "shuffle" Then Call SendMessage(waHWnd, &H111, 40023, vbNull)
If wCmd = "close" Then Call SendMessage(waHWnd, &H111, 40001, vbNull)
If wCmd = "volume up" Then Call SendMessage(waHWnd, &H111, 40058, vbNull)
If wCmd = "volume down" Then Call SendMessage(waHWnd, &H111, 40059, vbNull)
If wCmd = "step back" Then
  Call SendMessage(waHWnd, &H111, 40061, vbNull): Call Sleep(111)
  Call mPrg.GetStatus(2, CStr(WinampControl_PlayPosition(waHWnd)))
End If
If wCmd = "step next" Then
  Call SendMessage(waHWnd, &H111, 40060, vbNull): Call Sleep(111)
  Call mPrg.GetStatus(2, CStr(WinampControl_PlayPosition(waHWnd)))
End If
End Sub

Private Function WinampControl_PlayPosition(ByVal wHwnd As Long) As Double
Dim CurTime As Long, AllTime As Long
CurTime = SendMessage(wHwnd, &H400, 0, 105)
AllTime = SendMessage(wHwnd, &H400, 1, 105)
AllTime = AllTime * 1000
WinampControl_PlayPosition = CurTime / (AllTime / 100)
End Function

Private Function WinampControl_GetTrackName(ByVal wHwnd As Long) As String
Dim SS As String, SS1 As String

SS = String(255, " ")
Call GetWindowText(wHwnd, SS, 255)
SS = Trim(SS)
SS = Left$(SS, Len(SS) - 1)
If Right$(SS, 4) = "*** " Then
  SS = Left$(SS, Len(SS) - 4)
End If
SS = Right$(SS, Len(SS) - InStr(SS, "* "))

If InStrRev(SS, " - ") > 0 Then
  SS1 = Left$(SS, InStrRev(SS, " - ") - 2)
End If
WinampControl_GetTrackName = SS1
End Function

Private Sub Other_ShowDateTime(ByRef wFormat As String, Optional ByRef wPaste As Byte = 0)
Dim ShwStr As String
ShwStr = Format(Now, wFormat, vbUseSystemDayOfWeek, vbUseSystem)
If wPaste = 1 Then
  fPrg.Visible = False: DoEvents
  ShwStr = Replace(ShwStr, "(", "")
  ShwStr = Replace(ShwStr, ")", "")
  Call SendKeys(ShwStr, 11)
  fPrg.Visible = True
  ShwStr = MapOth(17) & " " & ShwStr
End If
Call mPrg.GetStatus(3, ShwStr)
End Sub

Private Sub NetConnect(ByVal wDial As String)
Dim Ret As Long
Ret = InternetDial(0&, wDial, &H2, 0, 0&)
If Ret = 0 Then
  Call mPrg.GetStatus(3, MapOth(18) & " (" & wDial & ")")
Else
  Call mPrg.GetStatus(3, MapOth(19) & " " & wDial)
End If
End Sub






























Public Function GetLastHWnd() As Long
If Lck = 0 Then
  fPrg.Visible = False: DoEvents
  GetLastHWnd = GetForegroundWindow
  fPrg.Visible = True
End If
End Function
