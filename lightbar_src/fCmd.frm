VERSION 5.00
Begin VB.Form fCmd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Дополнительные комманды"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fCmd.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDpl 
      Caption         =   "Подключение"
      Height          =   615
      Index           =   6
      Left            =   1875
      TabIndex        =   20
      Top             =   2400
      Width           =   2040
      Begin VB.ComboBox cbDial 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   225
         Width           =   1890
      End
   End
   Begin VB.Frame frDpl 
      Caption         =   "Прозрачность"
      Height          =   615
      Index           =   5
      Left            =   1875
      TabIndex        =   19
      Top             =   1725
      Width           =   2040
      Begin VB.HScrollBar scTrans 
         Height          =   315
         LargeChange     =   50
         Left            =   75
         Max             =   255
         Min             =   1
         TabIndex        =   10
         Top             =   225
         Value           =   204
         Width           =   1440
      End
      Begin VB.Label lTrn 
         Alignment       =   2  'Center
         Caption         =   "80%"
         Height          =   240
         Left            =   1500
         TabIndex        =   25
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame frDpl 
      Caption         =   "Дата / Время"
      Height          =   3540
      Index           =   4
      Left            =   3975
      TabIndex        =   16
      Top             =   150
      Width           =   2040
      Begin VB.TextBox tDtOut 
         BackColor       =   &H8000000F&
         Height          =   1965
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "fCmd.frx":08CA
         Top             =   600
         Width           =   1890
      End
      Begin VB.TextBox tDtIn 
         Height          =   315
         Left            =   75
         TabIndex        =   6
         Text            =   "dd.mm.yyyy hh:nn:ss (mmm, dddd)"
         Top             =   225
         Width           =   1890
      End
      Begin VB.Label infAb2 
         Caption         =   $"fCmd.frx":08F6
         Enabled         =   0   'False
         Height          =   840
         Left            =   1050
         TabIndex        =   18
         Top             =   2625
         Width           =   915
      End
      Begin VB.Label infAb1 
         Caption         =   $"fCmd.frx":0923
         Enabled         =   0   'False
         Height          =   840
         Left            =   75
         TabIndex        =   17
         Top             =   2625
         Width           =   915
      End
   End
   Begin VB.Frame frDpl 
      Caption         =   "Ускорение"
      Height          =   540
      Index           =   3
      Left            =   1875
      TabIndex        =   15
      Top             =   1125
      Width           =   2040
      Begin VB.CheckBox chForce 
         Caption         =   "Ускоренно"
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   225
         Value           =   1  'Checked
         Width           =   1890
      End
   End
   Begin VB.Frame frDpl 
      Caption         =   "Запуск"
      Height          =   915
      Index           =   2
      Left            =   1875
      TabIndex        =   14
      Top             =   150
      Width           =   2040
      Begin VB.CheckBox chRun 
         Caption         =   $"fCmd.frx":0953
         Height          =   615
         Left            =   75
         TabIndex        =   8
         Top             =   225
         Value           =   1  'Checked
         Width           =   1890
      End
   End
   Begin VB.Frame frDpl 
      Caption         =   "Настройки"
      Height          =   1515
      Index           =   1
      Left            =   6150
      TabIndex        =   13
      Top             =   750
      Width           =   2040
      Begin VB.ComboBox cbVolDev 
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   450
         Width           =   1890
      End
      Begin VB.HScrollBar scStp 
         Height          =   315
         LargeChange     =   2
         Left            =   75
         Max             =   10
         Min             =   1
         TabIndex        =   5
         Top             =   1125
         Value           =   2
         Width           =   1890
      End
      Begin VB.Label infStp 
         Caption         =   "Шаг"
         Height          =   240
         Left            =   75
         TabIndex        =   24
         Top             =   900
         Width           =   1290
      End
      Begin VB.Label infDev 
         Caption         =   "Устройство:"
         Height          =   240
         Left            =   75
         TabIndex        =   22
         Top             =   225
         Width           =   1890
      End
      Begin VB.Label lVolStep 
         Alignment       =   1  'Right Justify
         Caption         =   "2 %"
         Height          =   240
         Left            =   1425
         TabIndex        =   21
         Top             =   900
         Width           =   540
      End
   End
   Begin VB.Frame frDpl 
      Caption         =   "Буква CD-ROMа"
      Height          =   615
      Index           =   0
      Left            =   6150
      TabIndex        =   12
      Top             =   75
      Width           =   2040
      Begin VB.ComboBox cbCD 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1890
      End
   End
   Begin VB.ListBox lCmd 
      Height          =   5115
      IntegralHeight  =   0   'False
      ItemData        =   "fCmd.frx":097E
      Left            =   1950
      List            =   "fCmd.frx":0980
      TabIndex        =   2
      Top             =   75
      Width           =   4140
   End
   Begin VB.ListBox lGrp 
      Height          =   5115
      IntegralHeight  =   0   'False
      ItemData        =   "fCmd.frx":0982
      Left            =   75
      List            =   "fCmd.frx":099B
      TabIndex        =   1
      Top             =   75
      Width           =   1815
   End
   Begin VB.CommandButton cCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   465
      Left            =   7200
      TabIndex        =   0
      Top             =   4725
      Width           =   990
   End
   Begin VB.CommandButton cOK 
      Caption         =   "ОК"
      Enabled         =   0   'False
      Height          =   465
      Left            =   6150
      TabIndex        =   23
      Top             =   4725
      Width           =   990
   End
End
Attribute VB_Name = "fCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                    Inner Commands Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             23.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Private Type RAS_ENTRIES
  dwSize As Long
  szEntryname(256) As Byte
End Type
Private Declare Function RasEnumEntriesA Lib "rasapi32.dll" (ByVal Reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long

Private Grp As Integer, Cmd As Integer
Private Const klDpl As Byte = 6

Private Sub cCancel_Click()
MapB(2).wOpr = 0
Me.Hide
End Sub

Private Sub cOK_Click()
Call SaveSettings
Me.Hide
End Sub

Private Sub Form_Activate()
Call mPrg.SetCur(cCancel.hwnd)
End Sub

Private Sub Form_Load()

Call mLng.LoadLang(LangFile, "cmd")

If FormNotTop = 0 Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
Call CmdLoad
End Sub

Private Sub lCmd_Click()
Dim I As Integer
For I = 0 To klDpl Step 1
  frDpl(I).Visible = False
Next I
Cmd = Val(Right$(lCmd.Text, 10))
If Cmd = 0 Then cOK.Enabled = False Else cOK.Enabled = True
Call GetPrp
End Sub

Private Sub lGrp_Click()
Dim I As Integer
For I = 0 To klDpl Step 1
  frDpl(I).Visible = False
Next I
lCmd.Clear
Cmd = 0
Grp = Val(Right$(lGrp.Text, 10))
cOK.Enabled = False
Call GetCmd
End Sub

'################################################'
'### SUBS AND FUNCTIONS #########################'
'################################################'

Private Sub CmdLoad()
Dim I As Integer
Dim ListD() As String
Dim plSize As Long
Dim plEntries As Long
Dim psConName As String
Dim plIndex As Long
Dim RAS(255) As RAS_ENTRIES

'spisok kategorij
lGrp.Clear
For I = 1 To 6 Step 1
  lGrp.AddItem MapCmd(I, 0) & Space(100) & I
Next I
lGrp.AddItem MapCmd(99, 0) & Space(100) & "99"

'rasstavlyaem okna propertiesov
For I = 0 To klDpl Step 1
  frDpl(I).Visible = False
Next I
For I = 0 To klDpl Step 1
  frDpl(I).Left = 410
  frDpl(I).Top = 5
Next I
'zagruzhaem bukvy cdromov
cbCD.Clear
For I = 0 To 25 Step 1
  If GetDriveType(Chr(I + 65) & ":") = 5 Then cbCD.AddItem Chr(I + 65)
Next
If cbCD.ListCount > 0 Then cbCD.Text = cbCD.List(0)
'zagruzhaem spisok audio devajsov

'########

'nahodim vozmozhnye podklyucheniya
ReDim ListD(0)
RAS(0).dwSize = 264
plSize = 256 * RAS(0).dwSize
Call RasEnumEntriesA(vbNullString, vbNullString, RAS(0), plSize, plEntries)
plEntries = plEntries - 1
If plEntries >= 0 Then
  ReDim ListD(plEntries)
  For plIndex = 0 To plEntries
    psConName = StrConv(RAS(plIndex).szEntryname(), vbUnicode)
    ListD(plIndex) = Left$(psConName, InStr(psConName, vbNullChar) - 1)
  Next plIndex
End If
cbDial.Clear
On Error GoTo 0
For I = 0 To UBound(ListD) Step 1
  cbDial.AddItem ListD(I)
Next I
If cbDial.ListCount > 0 Then
  cbDial.ListIndex = 0
End If
End Sub

Private Sub GetCmd()
Dim I As Integer
If Grp = 1 Then 'pitanie
  For I = 1 To 3 Step 1
    lCmd.AddItem MapCmd(1, I) & Space(100) & I
  Next I
End If
If Grp = 2 Then 'cdrom
  For I = 1 To 3 Step 1
    lCmd.AddItem MapCmd(2, I) & Space(100) & I
  Next I
End If
If Grp = 3 Then 'okna
  For I = 1 To 7 Step 1
    lCmd.AddItem MapCmd(3, I) & Space(100) & I
  Next I
End If
If Grp = 4 Then 'zvuk
  For I = 1 To 2 Step 1
    lCmd.AddItem MapCmd(4, I) & Space(100) & I
  Next I
End If
If Grp = 5 Then 'winamp
  For I = 1 To 11 Step 1
    lCmd.AddItem MapCmd(5, I) & Space(100) & I
  Next I
End If
If Grp = 6 Then 'set'
  For I = 1 To 1 Step 1
    lCmd.AddItem MapCmd(6, I) & Space(100) & I
  Next I
End If
If Grp = 99 Then 'raznoe
  For I = 1 To 4 Step 1
    lCmd.AddItem MapCmd(99, I) & Space(100) & I
  Next I
End If
End Sub

Private Sub GetPrp()
If Grp = 1 Then frDpl(3).Visible = True 'pitanie
If Grp = 2 Then frDpl(0).Visible = True 'cdrom
If Grp = 3 Then 'okna
  If Cmd = 5 Then frDpl(5).Visible = True 'prozrachnost'
End If
If Grp = 4 Then 'zvuk
  frDpl(1).Height = 101 '56
  frDpl(1).Visible = True
End If
If Grp = 5 Then frDpl(2).Visible = True 'zapuskat' programmu
If Grp = 6 Then frDpl(6).Visible = True 'dial
If Grp = 99 Then 'raznoe
  If Cmd = 1 Or Cmd = 2 Then frDpl(4).Visible = True: Call tDtIn_Change 'data i vremya
End If
End Sub

Private Sub SaveSettings()
If Grp = 0 Or Cmd = 0 Then Exit Sub
MapB(2).wOpr = 3
MapB(2).wFil = "lbar_"
MapB(2).wPrm = ""
MapB(2).wDir = ""
MapB(2).wShw = 0
MapB(2).wCap = ""
MapB(2).wHtM = 0
MapB(2).wHtK = 0
MapB(2).wHtN = 0
MapB(2).wIFl = ""
MapB(2).wINm = 0
MapB(2).wKmm = ""
If Grp = 1 Then 'pitanie
  MapB(2).wOpr = 4
  MapB(2).wFil = MapB(2).wFil & "shutdown"
  If Cmd = 1 Then MapB(2).wPrm = "shutdown"
  If Cmd = 2 Then MapB(2).wPrm = "reboot"
  If Cmd = 3 Then MapB(2).wPrm = "logoff"
  If chForce.Value = 1 Then
    MapB(2).wDir = "force"
    MapB(2).wCap = " " & MapOth(23)
  End If
  If Cmd = 1 Then MapB(2).wCap = MapCmd(1, 1) & MapB(2).wCap
  If Cmd = 2 Then MapB(2).wCap = MapCmd(1, 2) & MapB(2).wCap
  If Cmd = 3 Then MapB(2).wCap = MapCmd(1, 3) & MapB(2).wCap
  MapB(2).wIFl = "%PROGPATH%\icons.icl"
  If Cmd = 1 Then MapB(2).wINm = 1
  If Cmd = 2 Then MapB(2).wINm = 2
  If Cmd = 3 Then MapB(2).wINm = 14
End If
If Grp = 2 Then 'cdrom
  MapB(2).wFil = MapB(2).wFil & "cdrom"
  MapB(2).wPrm = cbCD.Text
  If Cmd = 1 Then MapB(2).wDir = "open"
  If Cmd = 2 Then MapB(2).wDir = "close"
  If Cmd = 3 Then MapB(2).wDir = "open/close"
  If Cmd = 1 Then MapB(2).wCap = MapCmd(2, 1) & " " & cbCD.Text & ":"
  If Cmd = 2 Then MapB(2).wCap = MapCmd(2, 2) & " " & cbCD.Text & ":"
  If Cmd = 3 Then MapB(2).wCap = MapCmd(2, 3) & " " & cbCD.Text & ":"
  MapB(2).wIFl = "%systemroot%\SYSTEM32\SHELL32.DLL"
  MapB(2).wINm = 27
End If
If Grp = 3 Then 'okna
  MapB(2).wFil = MapB(2).wFil & "window"
  MapB(2).wIFl = "%systemroot%\SYSTEM32\SHELL32.DLL"
  If Cmd = 1 Or Cmd = 2 Then
    MapB(2).wPrm = "close"
    MapB(2).wINm = 132
    If Cmd = 1 Then
      MapB(2).wDir = "close"
      MapB(2).wCap = MapCmd(3, 1)
    End If
    If Cmd = 2 Then
      MapB(2).wDir = "quit"
      MapB(2).wCap = MapCmd(3, 2)
    End If
  End If
  If Cmd = 3 Or Cmd = 4 Then
    MapB(2).wPrm = "topmost"
    MapB(2).wINm = 99
    If Cmd = 3 Then
      MapB(2).wDir = "top"
      MapB(2).wCap = MapCmd(3, 3)
    End If
    If Cmd = 4 Then
      MapB(2).wDir = "notop"
      MapB(2).wCap = MapCmd(3, 4)
    End If
  End If
  If Cmd = 5 Then
    MapB(2).wCap = MapCmd(3, 5)
    MapB(2).wPrm = "transparent"
    MapB(2).wDir = scTrans.Value
    MapB(2).wINm = 3
  End If
  If Cmd = 6 Or Cmd = 7 Then
    MapB(2).wPrm = "transform"
    MapB(2).wINm = 170
    If Cmd = 6 Then
      MapB(2).wDir = "maximize"
      MapB(2).wCap = MapCmd(3, 6)
    End If
    If Cmd = 7 Then
      MapB(2).wDir = "minimize"
      MapB(2).wCap = MapCmd(3, 7)
    End If
  End If
End If
If Grp = 4 Then 'zvuk
  MapB(2).wOpr = 2
  MapB(2).wFil = MapB(2).wFil & "sound"
  If Cmd = 1 Then MapB(2).wPrm = "up"
  If Cmd = 2 Then MapB(2).wPrm = "down"
  MapB(2).wDir = scStp.Value
  If Cmd = 1 Then MapB(2).wCap = MapCmd(4, 1)
  If Cmd = 2 Then MapB(2).wCap = MapCmd(4, 2)
  MapB(2).wIFl = "%systemroot%\system32\Mmsys.cpl"
  MapB(2).wINm = 37
End If
If Grp = 5 Then 'winamp
  If Cmd > 7 Then MapB(2).wOpr = 2 Else MapB(2).wOpr = 1
  MapB(2).wFil = MapB(2).wFil & "winamp"
  If chRun.Value = 0 Then MapB(2).wDir = "" Else MapB(2).wDir = "run"
  MapB(2).wIFl = "%PROGPATH%\icons.icl"
  MapB(2).wINm = Cmd + 2
  If Cmd = 1 Then MapB(2).wPrm = "back": MapB(2).wCap = "WA -> " & MapCmd(5, 1)
  If Cmd = 2 Then MapB(2).wPrm = "play": MapB(2).wCap = "WA -> " & MapCmd(5, 2)
  If Cmd = 3 Then MapB(2).wPrm = "pause": MapB(2).wCap = "WA -> " & MapCmd(5, 3)
  If Cmd = 4 Then MapB(2).wPrm = "stop": MapB(2).wCap = "WA -> " & MapCmd(5, 4)
  If Cmd = 5 Then MapB(2).wPrm = "next": MapB(2).wCap = "WA -> " & MapCmd(5, 5)
  If Cmd = 6 Then MapB(2).wPrm = "shuffle": MapB(2).wCap = "WA -> " & MapCmd(5, 6)
  If Cmd = 7 Then MapB(2).wPrm = "close": MapB(2).wCap = "WA -> " & MapCmd(5, 7)
  If Cmd = 8 Then MapB(2).wPrm = "volume up": MapB(2).wCap = "WA -> " & MapCmd(5, 8)
  If Cmd = 9 Then MapB(2).wPrm = "volume down": MapB(2).wCap = "WA -> " & MapCmd(5, 9)
  If Cmd = 10 Then MapB(2).wPrm = "step back": MapB(2).wCap = "WA -> " & MapCmd(5, 10)
  If Cmd = 11 Then MapB(2).wPrm = "step next": MapB(2).wCap = "WA -> " & MapCmd(5, 11)
End If
If Grp = 6 Then 'set'
  MapB(2).wFil = MapB(2).wFil & "net"
  MapB(2).wPrm = "connect"
  MapB(2).wDir = cbDial.Text
  MapB(2).wCap = MapCmd(6, 1) & " " & cbDial.Text
  MapB(2).wIFl = "%systemroot%\SYSTEM32\SHELL32.DLL"
  MapB(2).wINm = 89
End If
If Grp = 99 Then 'raznoe
  MapB(2).wOpr = 3
  MapB(2).wFil = MapB(2).wFil & "other"
  If Cmd = 1 Or Cmd = 2 Then 'datetime
    If Cmd = 1 Then MapB(2).wPrm = "datetime"
    If Cmd = 2 Then MapB(2).wPrm = "datetime paste"
    MapB(2).wDir = Replace(tDtIn.Text, vbCrLf, " ")
    If Cmd = 1 Then MapB(2).wCap = MapCmd(99, 1)
    If Cmd = 2 Then MapB(2).wCap = MapCmd(99, 2)
    MapB(2).wIFl = "%SYSTEMROOT%\SYSTEM32\Timedate.cpl"
    MapB(2).wINm = 1
  End If
  If Cmd = 3 Then 'clear clipboard
    MapB(2).wPrm = "clear clipboard"
    MapB(2).wCap = MapCmd(99, 3)
    MapB(2).wIFl = "%systemroot%\SYSTEM32\SHELL32.DLL"
    MapB(2).wINm = 55
  End If
  If Cmd = 4 Then 'izvlechenie usb
    MapB(2).wPrm = "extract usb"
    MapB(2).wCap = MapCmd(99, 4)
    MapB(2).wIFl = "%systemroot%\SYSTEM32\SHELL32.DLL"
    MapB(2).wINm = 27
  End If
End If
'wOpr As Byte   ' operaciya
'wFil As String ' put' k failu
'wPrm As String ' parametry
'wDir As String ' katalog po umolchaniyu
'wShw As Byte   ' rezjim otkrytiya
'wCap As String ' caption
'wHtM As Byte   ' hot mod
'wHtK As Byte   ' hot key
'wHtN As Byte   ' blokirovschik goryachih klavish
'wIFl As String ' fajl ikonki
'wINm As Long   ' nomer ikonki
'wKmm As String ' kommentarij
End Sub









































Private Sub scStp_Change()
lVolStep.Caption = scStp.Value & " %"
End Sub
Private Sub scStp_GotFocus()
lCmd.SetFocus
End Sub
Private Sub scStp_Scroll()
Call scStp_Change
End Sub

Private Sub scTrans_Change()
lTrn.Caption = CInt(scTrans.Value / 255 * 100) & " %"
End Sub
Private Sub scTrans_GotFocus()
lCmd.SetFocus
End Sub
Private Sub scTrans_Scroll()
Call scTrans_Change
End Sub

Private Sub tDtIn_Change()
tDtOut.Text = Format(Now, tDtIn.Text, vbUseSystemDayOfWeek, vbUseSystem)
End Sub

Private Sub tDtIn_GotFocus()
tDtIn.SelStart = 0
tDtIn.SelLength = Len(tDtIn.Text)
End Sub
