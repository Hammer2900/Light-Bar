VERSION 5.00
Begin VB.Form fAbt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About MyApp"
   ClientHeight    =   2265
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fAbt.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cHtK 
      Caption         =   "Проверить горячие клавиши"
      Height          =   690
      Left            =   3150
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cPth 
      Caption         =   "Проверить пути"
      Height          =   465
      Left            =   3150
      TabIndex        =   1
      Top             =   75
      Width           =   1215
   End
   Begin VB.Frame frAbout 
      Caption         =   "О программе"
      Height          =   2115
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   3015
      Begin VB.Label lSit 
         Caption         =   "lightbar.narod.ru"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1275
         MouseIcon       =   "fAbt.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1050
         Width           =   1665
      End
      Begin VB.Label lEml 
         Caption         =   "wfs@of.kz"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1275
         MouseIcon       =   "fAbt.frx":1194
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   825
         Width           =   1665
      End
      Begin VB.Label lINFORMATION 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   2115
         Index           =   3
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label infSit 
         Caption         =   "Сайт:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   750
         TabIndex        =   9
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label lVer 
         Caption         =   "Version"
         Height          =   240
         Left            =   750
         TabIndex        =   8
         Top             =   525
         Width           =   2190
      End
      Begin VB.Label lTtl 
         Caption         =   "Application Title"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   750
         TabIndex        =   7
         Top             =   225
         Width           =   2190
      End
      Begin VB.Label infEml 
         Caption         =   "e-Mail:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   750
         TabIndex        =   6
         Top             =   825
         Width           =   540
      End
      Begin VB.Label infInf 
         Alignment       =   2  'Center
         Caption         =   $"fAbt.frx":1A5E
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   75
         TabIndex        =   4
         Top             =   1350
         Width           =   2865
      End
      Begin VB.Image iIcon 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   75
         Top             =   225
         Width           =   540
      End
      Begin VB.Label infWrt 
         Alignment       =   2  'Center
         Caption         =   "Написано на Visual Basic 6.0"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   75
         TabIndex        =   12
         Top             =   1800
         Width           =   2865
      End
   End
   Begin VB.CommandButton cOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   465
      Left            =   3150
      TabIndex        =   0
      Top             =   1725
      Width           =   1215
   End
End
Attribute VB_Name = "fAbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                             About Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             06.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Private Btt As Integer

Private Sub cHtK_Click()
Dim I As Integer, II As Integer
Dim Btt As Integer
Dim SS As String
Dim FF As Long
FF = FreeFile
Open App.Path & "\hotkeys.txt" For Output As #FF
  Print #FF, MapOth(10)
  Print #FF, "(" & App.Path & "\hotkeys.txt)"
  Print #FF, ""
  If HotMod > 0 Or HotKey > 0 Then
    SS = ""
    If HotMod > 0 Then SS = GetTextMod(HotMod)
    If HotKey > 0 Then SS = SS & MapKN(HotKey)
    SS = SS & " = " & MapOth(20)
    Print #FF, SS
    Print #FF, ""
  End If
  
  For I = 0 To bttCol - 1 Step 1
    For II = 0 To bttRow - 1 Step 1
      Btt = (II + 1) * 100 + I + 1
      If MapB(Btt).wOpr > 0 Then
        If MapB(Btt).wHtM > 0 Or MapB(Btt).wHtK > 0 Then
          SS = ""
          If MapB(Btt).wHtM > 0 Then SS = GetTextMod(MapB(Btt).wHtM)
          If MapB(Btt).wHtK > 0 Then SS = SS & MapKN(MapB(Btt).wHtK)
          SS = SS & " = " & MapB(Btt).wCap & " (" & MapB(Btt).wFil & ")"
          Print #FF, SS
        End If
      End If
    Next II
  Next I
Close #FF
Call ShellExecute(0&, "open", App.Path & "\hotkeys.txt", "", "", 1)
Unload Me
End Sub

Private Sub cHtK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fAbt, 2, MapMsg(0))
End Sub

Private Sub cPth_Click()
Dim FF As Long
FF = FreeFile
Open App.Path & "\paths.txt" For Output As #FF
  Print #FF, MapOth(11)
  Print #FF, "(" & App.Path & "\paths.txt)"
  Print #FF, ""
  Print #FF, MapOth(21)
  Print #FF, ""
  Print #FF, "%ALLUSERSPROFILE% = " & GetDir("%ALLUSERSPROFILE%")
  Print #FF, "%APPDATA% = " & GetDir("%APPDATA%")
  Print #FF, "%COMMONPROGRAMFILES% = " & GetDir("%COMMONPROGRAMFILES%")
  Print #FF, "%HOMEDRIVE% = " & GetDir("%HOMEDRIVE%")
  Print #FF, "%HOMEPATH% = " & GetDir("%HOMEPATH%")
  Print #FF, "%PROGRAMFILES% = " & GetDir("%PROGRAMFILES%")
  Print #FF, "%SYSTEMDRIVE% = " & GetDir("%SYSTEMDRIVE%")
  Print #FF, "%SYSTEMROOT% = " & GetDir("%SYSTEMROOT%")
  Print #FF, "%TEMP% = " & GetDir("%TEMP%")
  Print #FF, "%TMP% = " & GetDir("%TMP%")
  Print #FF, "%USERPROFILE% = " & GetDir("%USERPROFILE%")
  Print #FF, "%WINDIR% = " & GetDir("%WINDIR%")
  Print #FF, ""
  Print #FF, MapOth(22)
  Print #FF, ""
  Print #FF, "%PROGPATH% = " & GetDir("%PROGPATH%")
  Print #FF, "%MUSIC% = " & GetDir("%MUSIC%")
  Print #FF, "%DESKTOP% = " & GetDir("%DESKTOP%")
  Print #FF, "%STARTMENUPROG% = " & GetDir("%STARTMENUPROG%")
  Print #FF, "%DOCUMENTS% = " & GetDir("%DOCUMENTS%")
  Print #FF, "%FAVORITES% = " & GetDir("%FAVORITES%")
  Print #FF, "%AUTORUN% = " & GetDir("%AUTORUN%")
  Print #FF, "%RECENT% = " & GetDir("%RECENT%")
  Print #FF, "%SENDTO% = " & GetDir("%SENDTO%")
  Print #FF, "%STARTMENU% = " & GetDir("%STARTMENU%")
  Print #FF, "%VIDEO% = " & GetDir("%VIDEO%")
  Print #FF, "%DESKTOP2% = " & GetDir("%DESKTOP2%")
  Print #FF, "%NETHOOD% = " & GetDir("%NETHOOD%")
  Print #FF, "%FONTS% = " & GetDir("%FONTS%")
  Print #FF, "%TEMPLATES% = " & GetDir("%TEMPLATES%")
  Print #FF, "%AUSTARTMENU% = " & GetDir("%AUSTARTMENU%")
  Print #FF, "%AUSTARTMENUPROG% = " & GetDir("%AUSTARTMENUPROG%")
  Print #FF, "%AUAUTORUN% = " & GetDir("%AUAUTORUN%")
  Print #FF, "%AUDESKTOP% = " & GetDir("%AUDESKTOP%")
  Print #FF, "%APPLICATIONDATA% = " & GetDir("%APPLICATIONDATA%")
  Print #FF, "%PRINTHOOD% = " & GetDir("%PRINTHOOD%")
  Print #FF, "%LOCALSETTAPPDATA% = " & GetDir("%LOCALSETTAPPDATA%")
  Print #FF, "%AUFAVORITES% = " & GetDir("%AUFAVORITES%")
  Print #FF, "%CASHE% = " & GetDir("%CASHE%")
  Print #FF, "%COOKIES% = " & GetDir("%COOKIES%")
  Print #FF, "%HISTORY% = " & GetDir("%HISTORY%")
  Print #FF, "%AUAPPDATA% = " & GetDir("%AUAPPDATA%")
  Print #FF, "%WINDOWS% = " & GetDir("%WINDOWS%")
  Print #FF, "%SYSTEM32% = " & GetDir("%SYSTEM32%")
  Print #FF, "%PROGRAMDIR% = " & GetDir("%PROGRAMDIR%")
  Print #FF, "%PICTURES% = " & GetDir("%PICTURES%")
  Print #FF, "%USERDIR% = " & GetDir("%USERDIR%")
  Print #FF, "%SYSTEM322% = " & GetDir("%SYSTEM322%")
  Print #FF, "%COMMONFILES% = " & GetDir("%COMMONFILES%")
  Print #FF, "%AUTEMPLATES% = " & GetDir("%AUTEMPLATES%")
  Print #FF, "%AUDOCUMENTS% = " & GetDir("%AUDOCUMENTS%")
  Print #FF, "%ADMINISTRATION% = " & GetDir("%ADMINISTRATION%")
  Print #FF, "%AUMUSIC% = " & GetDir("%AUMUSIC%")
  Print #FF, "%AUPICTURES% = " & GetDir("%AUPICTURES%")
  Print #FF, "%AUVIDEO% = " & GetDir("%AUVIDEO%")
  Print #FF, "%RESOURCES% = " & GetDir("%RESOURCES%")
  Print #FF, "%CDBURNING% = " & GetDir("%CDBURNING%")
Close #FF
Call ShellExecute(0&, "open", App.Path & "\paths.txt", "", "", 1)
Unload Me
End Sub

Private Sub cOK_Click()
Call mPrg.SetCur(-3, 0)
Unload Me
End Sub

Private Sub cPth_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fAbt, 2, MapMsg(29))
End Sub

Private Sub Form_Activate()
Call mPrg.SetCur(cOK.hwnd)
End Sub

Private Sub Form_Load()

Call mLng.LoadLang(LangFile, "abt")

iIcon.Picture = fPrg.Icon
If FormNotTop = 0 Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2

Me.Caption = " " & frAbout.Caption & " " & App.Title
lVer.Caption = "Version " & App.Major & "." & App.Minor & " (" & App.Revision & ")"
lTtl.Caption = App.Title
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lEml.FontUnderline = False
lSit.FontUnderline = False
End Sub

Private Sub lblVersion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lEml.FontUnderline = False
lSit.FontUnderline = False
End Sub

Private Sub lEml_Click()
If Btt <> 2 Then
  Call ShellExecute(0&, "open", "mailto:" & lEml.Caption & "?subject=" & App.Title & " v." & App.Major & "." & App.Minor & " (" & App.Revision & ")", "", "", 0)
  Unload Me
End If
End Sub

Private Sub lEml_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Btt = Button
End Sub

Private Sub lEml_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lEml.FontUnderline = False Then lEml.FontUnderline = True
lSit.FontUnderline = False
End Sub

Private Sub lEml_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fAbt, 2, MapMsg(30))
End Sub

Private Sub lINFORMATION_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lEml.FontUnderline = False
lSit.FontUnderline = False
End Sub

Private Sub lSit_Click()
If Btt <> 2 Then
  Call ShellExecute(0&, "open", "http://" & lSit.Caption, "", "", 0)
  Unload Me
End If
End Sub

Private Sub lSit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Btt = Button
End Sub

Private Sub lSit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lEml.FontUnderline = False
If lSit.FontUnderline = False Then lSit.FontUnderline = True
End Sub

Private Sub lSit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fAbt, 2, MapMsg(31))
End Sub
