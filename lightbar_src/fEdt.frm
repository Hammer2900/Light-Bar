VERSION 5.00
Begin VB.Form fEdt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Настройка ярлыка № "
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fEdt.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cReplacePaths 
      Caption         =   "Абсолютные пути"
      Height          =   540
      Index           =   1
      Left            =   5400
      TabIndex        =   62
      Top             =   3600
      Width           =   1365
   End
   Begin VB.CommandButton cReplacePaths 
      Caption         =   "Относительные пути"
      Height          =   540
      Index           =   0
      Left            =   5400
      TabIndex        =   33
      Top             =   3075
      Width           =   1365
   End
   Begin VB.Frame frHotKey 
      Caption         =   "Горячая клавиша"
      Height          =   765
      Left            =   75
      TabIndex        =   56
      Top             =   3750
      Width           =   5265
      Begin VB.CheckBox chHtN 
         Caption         =   "Не блоки- ровать"
         Height          =   465
         Left            =   75
         TabIndex        =   18
         Top             =   225
         Width           =   1140
      End
      Begin VB.TextBox tHtM 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   375
         Width           =   1740
      End
      Begin VB.TextBox tHtK 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton cHtKDel 
         DownPicture     =   "fEdt.frx":08CA
         Height          =   315
         Left            =   1650
         Picture         =   "fEdt.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton cHtK 
         DownPicture     =   "fEdt.frx":0B5E
         Height          =   315
         Left            =   1275
         Picture         =   "fEdt.frx":0CA8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Определить комбинацию клавиш"
         Top             =   375
         Width           =   315
      End
      Begin VB.Label infKey 
         Caption         =   "Клавиша:"
         Height          =   240
         Left            =   3825
         TabIndex        =   58
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label infMod 
         Caption         =   "Модификатор:"
         Height          =   240
         Left            =   2025
         TabIndex        =   57
         Top             =   150
         Width           =   1740
      End
   End
   Begin VB.Frame frChange 
      Caption         =   "Изменение"
      Height          =   1290
      Left            =   5400
      TabIndex        =   55
      Top             =   4275
      Width           =   1365
      Begin VB.CommandButton cPast 
         Caption         =   "Вставить"
         Height          =   465
         Left            =   75
         TabIndex        =   35
         Top             =   750
         Width           =   1215
      End
      Begin VB.CommandButton cCopy 
         Caption         =   "Скопировать"
         Height          =   465
         Left            =   75
         TabIndex        =   34
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame frMove 
      Caption         =   "Перемещение"
      Height          =   1440
      Left            =   5400
      TabIndex        =   53
      Top             =   1575
      Width           =   1365
      Begin VB.CommandButton cMove 
         DownPicture     =   "fEdt.frx":0DF2
         Height          =   315
         Index           =   0
         Left            =   150
         Picture         =   "fEdt.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Alt + <"
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton cMove 
         DownPicture     =   "fEdt.frx":1086
         Height          =   315
         Index           =   2
         Left            =   525
         Picture         =   "fEdt.frx":11D0
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Alt + ^"
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cMove 
         DownPicture     =   "fEdt.frx":131A
         Height          =   315
         Index           =   3
         Left            =   525
         Picture         =   "fEdt.frx":1464
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Alt + v"
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cMove 
         DownPicture     =   "fEdt.frx":15AE
         Height          =   315
         Index           =   1
         Left            =   900
         Picture         =   "fEdt.frx":16F8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Alt + >"
         Top             =   375
         Width           =   315
      End
      Begin VB.Label infMov 
         Alignment       =   2  'Center
         Caption         =   "(все изменения сохранятся)"
         Enabled         =   0   'False
         Height          =   390
         Left            =   75
         TabIndex        =   54
         Top             =   975
         Width           =   1215
      End
   End
   Begin VB.Frame frAction 
      Caption         =   "Обрабатывать по событию"
      Height          =   990
      Left            =   75
      TabIndex        =   52
      Top             =   75
      Width           =   5265
      Begin VB.OptionButton oOpr 
         Caption         =   "При двойном отпускании"
         Height          =   240
         Index           =   4
         Left            =   2700
         TabIndex        =   5
         Top             =   675
         Width           =   2490
      End
      Begin VB.OptionButton oOpr 
         Caption         =   "При отпускании клавиши"
         Height          =   240
         Index           =   3
         Left            =   2700
         TabIndex        =   4
         Top             =   450
         Width           =   2490
      End
      Begin VB.OptionButton oOpr 
         Caption         =   "При задержке нажатия"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   3
         Top             =   675
         Width           =   2490
      End
      Begin VB.OptionButton oOpr 
         Caption         =   "При нажатии клавиши"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   2
         Top             =   450
         Width           =   2490
      End
      Begin VB.OptionButton oOpr 
         Caption         =   "Не обрабатывать"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   225
         Value           =   -1  'True
         Width           =   2490
      End
   End
   Begin VB.CommandButton cClear 
      Caption         =   "Очистить"
      Height          =   465
      Left            =   4125
      TabIndex        =   37
      Top             =   5625
      Width           =   1290
   End
   Begin VB.Frame frTrk 
      Caption         =   "Переход"
      Height          =   1440
      Left            =   5400
      TabIndex        =   50
      Top             =   75
      Width           =   1365
      Begin VB.CommandButton cGo 
         DownPicture     =   "fEdt.frx":1842
         Height          =   315
         Index           =   1
         Left            =   900
         Picture         =   "fEdt.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Ctrl + >"
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton cGo 
         DownPicture     =   "fEdt.frx":1AD6
         Height          =   315
         Index           =   3
         Left            =   525
         Picture         =   "fEdt.frx":1C20
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ctrl + v"
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cGo 
         DownPicture     =   "fEdt.frx":1D6A
         Height          =   315
         Index           =   2
         Left            =   525
         Picture         =   "fEdt.frx":1EB4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ctrl + ^"
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cGo 
         DownPicture     =   "fEdt.frx":1FFE
         Height          =   315
         Index           =   0
         Left            =   150
         Picture         =   "fEdt.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ctrl + <"
         Top             =   375
         Width           =   315
      End
      Begin VB.Label infTrk 
         Alignment       =   2  'Center
         Caption         =   "(все изменения сохранятся)"
         Enabled         =   0   'False
         Height          =   390
         Left            =   75
         TabIndex        =   51
         Top             =   975
         Width           =   1215
      End
   End
   Begin VB.CommandButton cApply 
      Caption         =   "Применить"
      Height          =   465
      Left            =   5475
      TabIndex        =   36
      Top             =   5625
      Width           =   1290
   End
   Begin VB.Frame frOther 
      Caption         =   "Дополнительно"
      Height          =   990
      Left            =   75
      TabIndex        =   47
      Top             =   4575
      Width           =   5265
      Begin VB.ComboBox cbShw 
         Height          =   315
         ItemData        =   "fEdt.frx":2292
         Left            =   1275
         List            =   "fEdt.frx":229F
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   3915
      End
      Begin VB.TextBox tCap 
         Height          =   315
         Left            =   1275
         TabIndex        =   23
         Top             =   225
         Width           =   3915
      End
      Begin VB.Label infStl 
         Caption         =   "Стиль окна:"
         Height          =   240
         Left            =   75
         TabIndex        =   49
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label infInf 
         Caption         =   "Название:"
         Height          =   240
         Left            =   75
         TabIndex        =   48
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame frIco 
      Caption         =   "Значок"
      Height          =   1140
      Left            =   75
      TabIndex        =   43
      Top             =   2550
      Width           =   5265
      Begin VB.CheckBox chClr 
         Caption         =   "Цвет значка"
         Height          =   240
         Left            =   2550
         TabIndex        =   64
         Top             =   825
         Width           =   2265
      End
      Begin VB.CommandButton cClr 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4875
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   750
         Width           =   315
      End
      Begin VB.PictureBox pIco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1650
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   61
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   585
         Width           =   480
      End
      Begin VB.CommandButton cShl 
         DownPicture     =   "fEdt.frx":22E7
         Height          =   315
         Left            =   4875
         Picture         =   "fEdt.frx":2431
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Дополнительные команды"
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cIFl 
         DownPicture     =   "fEdt.frx":257B
         Height          =   315
         Left            =   4500
         Picture         =   "fEdt.frx":26C5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   465
         Index           =   1
         Left            =   2175
         TabIndex        =   17
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   465
         Index           =   0
         Left            =   1275
         TabIndex        =   16
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox tIFl 
         Height          =   315
         Left            =   1275
         TabIndex        =   13
         Top             =   225
         Width           =   3165
      End
      Begin VB.Label lIco 
         Caption         =   "0000 \ 0000"
         Height          =   240
         Left            =   2550
         TabIndex        =   46
         Tag             =   "1"
         Top             =   600
         Width           =   2340
      End
      Begin VB.Label infIcoNum 
         Caption         =   "Номер значка:"
         Height          =   240
         Left            =   75
         TabIndex        =   45
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label infIcoFil 
         Caption         =   "Файл значка:"
         Height          =   240
         Left            =   75
         TabIndex        =   44
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame frCommand 
      Caption         =   "Команда"
      Height          =   1365
      Left            =   75
      TabIndex        =   39
      Top             =   1125
      Width           =   5265
      Begin VB.CommandButton cFilDir 
         DownPicture     =   "fEdt.frx":280F
         Height          =   315
         Left            =   4500
         Picture         =   "fEdt.frx":2959
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cCmd 
         DownPicture     =   "fEdt.frx":2AA3
         Height          =   315
         Left            =   4875
         Picture         =   "fEdt.frx":2BED
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Дополнительные команды"
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cDir 
         DownPicture     =   "fEdt.frx":2D37
         Height          =   315
         Left            =   4875
         Picture         =   "fEdt.frx":2E81
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   975
         Width           =   315
      End
      Begin VB.CommandButton cFil 
         DownPicture     =   "fEdt.frx":2FCB
         Height          =   315
         Left            =   4125
         Picture         =   "fEdt.frx":3115
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox tFil 
         Height          =   315
         Left            =   1275
         TabIndex        =   6
         Top             =   225
         Width           =   2790
      End
      Begin VB.TextBox tPrm 
         Height          =   315
         Left            =   1275
         TabIndex        =   10
         Top             =   600
         Width           =   3915
      End
      Begin VB.TextBox tDir 
         Height          =   315
         Left            =   1275
         TabIndex        =   11
         Top             =   975
         Width           =   3540
      End
      Begin VB.Label lCmm 
         Caption         =   "Команда:"
         Height          =   240
         Left            =   75
         TabIndex        =   42
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lDir 
         Caption         =   "Рабочая папка:"
         Height          =   240
         Left            =   75
         TabIndex        =   40
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lDir2 
         Caption         =   "Параметры 2:"
         Height          =   240
         Left            =   75
         TabIndex        =   60
         Top             =   1050
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lPrm 
         Caption         =   "Параметры:"
         Height          =   240
         Left            =   75
         TabIndex        =   41
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label lPrm2 
         Caption         =   "Параметры 1:"
         Height          =   240
         Left            =   75
         TabIndex        =   59
         Top             =   675
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   465
      Left            =   1425
      TabIndex        =   0
      Top             =   5625
      Width           =   1290
   End
   Begin VB.CommandButton cOK 
      Caption         =   "ОК"
      Height          =   465
      Left            =   75
      TabIndex        =   38
      Top             =   5625
      Width           =   1290
   End
End
Attribute VB_Name = "fEdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                     Button Editor Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             06.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Public wCap As String

Private Sub cApply_Click()
Call SaveSettings
cApply.Enabled = False
End Sub

Private Sub cbShw_Change()
cApply.Enabled = True
End Sub

Private Sub cCancel_Click()
Call DrawSelBtt(0)
Call mPrg.SetCur(EdtB, 0)
Unload Me
End Sub

Private Sub cClear_Click()
oOpr(0).Value = True
tFil.Text = ""
tPrm.Text = ""
tDir.Text = ""
tIFl.Text = ""
tHtM.Tag = 0: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = 0: tHtK.Text = MapKN(Val(tHtK.Tag))

pIco.Cls
lIco.Caption = "(0 \ 0)"
tCap.Text = ""
cbShw.ListIndex = 0
End Sub

Private Sub cClr_Click()
Dim Res As Long
mDlg.hwndOwner = Me.hwnd
Res = mDlg.ShowColor
If Res >= 0 Then cClr.BackColor = Res
cApply.Enabled = True
End Sub

Private Sub cCmd_Click()
Dim OldEdtB
fCmd.Show 1, Me
If MapB(2).wOpr > 0 Then
  OldEdtB = EdtB
  EdtB = 2
  Call Zapolnenie
  EdtB = OldEdtB
  Me.Caption = wCap & " " & EdtB
  cApply.Enabled = True
End If
End Sub

Private Sub cCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(32))
End Sub

Private Sub cCopy_Click()
MapB(1) = MapB(EdtB)
End Sub

Private Sub cCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(33))
End Sub

Private Sub cGo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
  If Index = 0 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(34))
  If Index = 1 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(35))
  If Index = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(36))
  If Index = 3 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(37))
End If
End Sub

Private Sub chClr_Click()
cClr.Enabled = chClr.Value
cApply.Enabled = True
End Sub

Private Sub chHtN_Click()
cApply.Enabled = True
End Sub

Private Sub chHtN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(38))
End Sub

Private Sub cHtK_Click()
fKey.Show 1, Me
tHtM.Tag = RetMod: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = RetKey: tHtK.Text = MapKN(Val(tHtK.Tag))
End Sub

Private Sub cHtK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(39))
End Sub

Private Sub cHtKDel_Click()
tHtM.Tag = 0: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = 0: tHtK.Text = MapKN(Val(tHtK.Tag))
End Sub

Private Sub cHtKDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(40))
End Sub

Private Sub cMove_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
  If Index = 0 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(41))
  If Index = 1 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(42))
  If Index = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(43))
  If Index = 3 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(44))
End If
End Sub

Private Sub cPast_Click()
Dim OldEdtB As Integer
OldEdtB = EdtB
EdtB = 1
Call Zapolnenie
EdtB = OldEdtB
Me.Caption = wCap & " " & EdtB
cApply.Enabled = True
End Sub

Private Sub cDir_Click()
Dim Pth As String
mDlg.FolderDialogTitle = MapOth(12)
mDlg.ShowDirsOnly = True
Pth = mDlg.ShowFolder(tDir.Text)
tDir.SetFocus
If Pth <> "" Then tDir.Text = Pth
End Sub

Private Sub cFil_Click()
Dim Pth As String
Dim SS As String
mDlg.Filter = "All files (*.*)|*.*|"
mDlg.OpenDialogTitle = MapOth(13)
Pth = mDlg.ShowOpen(tFil.Text)
Pth = Replace(Pth, vbNullChar, "")
'tFil.SetFocus
If Pth <> "" Then
  tFil.Text = Pth
  tDir.Text = GetFolder(Pth)
  tIFl.Text = Pth
  'tIFl.Text = GetShortcutPath(tIFl.Text)
  pIco.Tag = 1
  
  tCap.Text = GetName(Pth)
  
  tIFl.SetFocus
  tFil.SetFocus
  
  oOpr(3).Value = True
End If
End Sub

Private Sub cIFl_Click()
Dim Pth As String
mDlg.Filter = MapOth(14) & " (*.bmp;*.dll;*.exe;*.icl;*.ico)|*.bmp;*.dll;*.exe;*.icl;*.ico|All files (*.*)|*.*|"
mDlg.OpenDialogTitle = MapOth(15)
Pth = mDlg.ShowOpen(tIFl.Text)
Pth = Replace(Pth, vbNullChar, "")
tIFl.SetFocus
If Pth <> "" Then
  tIFl.Text = Pth
  
  cIFl.SetFocus
  tIFl.SetFocus
End If
End Sub

Private Sub cFilDir_Click()
Dim FF As Long
Dim SS As String
Dim Pth As String
mDlg.FolderDialogTitle = MapOth(16)
mDlg.ShowDirsOnly = True
Pth = mDlg.ShowFolder(tDir.Text)
tDir.SetFocus
If Pth <> "" Then
  tFil.Text = Pth
  tDir.Text = Pth
  
  tIFl.Text = "%systemroot%\system32\shell32.dll"
  pIco.Tag = 5
  
  If Dir(Pth & "\Desktop.ini", 39) <> "" Then
    FF = FreeFile
    Open Pth & "\Desktop.ini" For Input As #FF
      Do
        If EOF(FF) = True Then Exit Do
        Input #FF, SS
        If Left$(Trim(SS), 8) = "IconFile" Then
          tIFl.Text = Trim(Right$(SS, Len(SS) - InStr(SS, "=")))
        End If
        If Left$(Trim(SS), 9) = "IconIndex" Then
          pIco.Tag = Val(Trim(Right$(SS, Len(SS) - InStr(SS, "=")))) + 1
        End If
      Loop
    Close #FF
  End If
  
  tCap.Text = GetName(Pth)
  tIFl.SetFocus
  tFil.SetFocus
  oOpr(3).Value = True
End If
End Sub

Private Sub cPast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(45))
End Sub

Private Sub cReplacePaths_Click(Index As Integer)
If Index = 0 Then
  tFil.Text = DeGetDir(tFil.Text)
  If Left$(tFil.Text, 5) <> "lbar_" Then tDir.Text = DeGetDir(tDir.Text)
  tIFl.Text = DeGetDir(tIFl.Text)
Else
  tFil.Text = GetDir(tFil.Text)
  If Left$(tFil.Text, 5) <> "lbar_" Then tDir.Text = GetDir(tDir.Text)
  tIFl.Text = GetDir(tIFl.Text)
End If
End Sub

Private Sub cReplacePaths_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(53))
End Sub

Private Sub cShl_Click()
tIFl.SetFocus
tIFl.Text = "%systemroot%\system32\shell32.dll"
cShl.SetFocus
tIFl.SetFocus
End Sub

Private Sub cShl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(46))
End Sub

Private Sub Form_Activate()
Call mPrg.SetCur(cCancel.hwnd)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask Then
  If KeyCode = vbKeyLeft Then Call cGo_Click(0)
  If KeyCode = vbKeyRight Then Call cGo_Click(1)
  If KeyCode = vbKeyUp Then Call cGo_Click(2)
  If KeyCode = vbKeyDown Then Call cGo_Click(3)
End If
If Shift = vbAltMask Then
  If KeyCode = vbKeyLeft Then Call cMove_Click(0)
  If KeyCode = vbKeyRight Then Call cMove_Click(1)
  If KeyCode = vbKeyUp Then Call cMove_Click(2)
  If KeyCode = vbKeyDown Then Call cMove_Click(3)
End If
End Sub

Private Sub oOpr_Click(Index As Integer)
If oOpr(1).Value = True Or oOpr(2).Value = True Then
  cbShw.Enabled = False
  cbShw.ListIndex = 0
Else
  cbShw.Enabled = True
End If
cApply.Enabled = True
End Sub

Private Sub oOpr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(47))
End Sub

Private Sub tHtK_Change()
cApply.Enabled = True
End Sub

Private Sub tHtM_Change()
cApply.Enabled = True
End Sub

Private Sub tIFl_LostFocus()
lIco.Tag = GetIcoCount(tIFl.Text)
If CLng(pIco.Tag) > CLng(lIco.Tag) Then pIco.Tag = CLng(lIco.Tag)
pIco.Cls
If CLng(lIco.Tag) = 0 Then
  pIco.Tag = 0
End If
'fprg.pKntIco.BackColor = rgb(
Call mPrg.DrawIco(tIFl.Text, CInt(pIco.Tag))
pIco.PaintPicture fPrg.pKntIco.Image, 0, 0, pIco.Width / 15, pIco.Height / 15
lIco.Caption = "(" & pIco.Tag & " \ " & lIco.Tag & ")"
End Sub

Private Sub cGo_Click(Index As Integer)
Dim CC As Integer, RR As Integer
Call SaveSettings
Call DrawSelBtt(0)

CC = EdtB Mod 100
RR = (EdtB - CC) / 100

If Index = 0 Then CC = CC - 1
If Index = 1 Then CC = CC + 1
If Index = 2 Then RR = RR - 1
If Index = 3 Then RR = RR + 1
  
  
If CC > bttCol Then CC = 1
If CC < 1 Then CC = bttCol

If RR > bttRow Then RR = 1
If RR < 1 Then RR = bttRow

EdtB = (RR) * 100 + CC

Call Zapolnenie

Call DrawSelBtt(1)

cApply.Enabled = False
End Sub

Private Sub cMove_Click(Index As Integer)
Dim CC As Integer, RR As Integer
Dim SelB As Integer, NewB As Integer

Call SaveSettings

SelB = EdtB
NewB = EdtB

CC = NewB Mod 100
RR = (NewB - CC) / 100
If Index = 0 Then CC = CC - 1
If Index = 1 Then CC = CC + 1
If Index = 2 Then RR = RR - 1
If Index = 3 Then RR = RR + 1
If CC > bttCol Then CC = 1
If CC < 1 Then CC = bttCol
If RR > bttRow Then RR = 1
If RR < 1 Then RR = bttRow
NewB = (RR) * 100 + CC

MapB(0) = MapB(SelB)
MapB(SelB) = MapB(NewB)
MapB(NewB) = MapB(0)

EdtB = SelB
Call DrawSelBtt(0)
If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
If MapB(EdtB).wOpr > 0 Then Call DrawIco(MapB(EdtB).wIFl, MapB(EdtB).wINm): Call DrawHK(MapB(EdtB).wINm)
fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(EdtB).Left + 1, BttCoord(EdtB).Top + 1

EdtB = NewB
Call DrawSelBtt(1)
If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
If MapB(EdtB).wOpr > 0 Then Call DrawIco(MapB(EdtB).wIFl, MapB(EdtB).wINm): Call DrawHK(MapB(EdtB).wINm)
fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(EdtB).Left + 1, BttCoord(EdtB).Top + 1
Call mPrg.GetActivPic
End Sub

Private Sub cOK_Click()
Call SaveSettings
Call DrawSelBtt(0)
Call mPrg.SetCur(EdtB, 0)
Unload Me
End Sub

Private Sub cScr_Click(Index As Integer)
Dim Ic As Long 'dlya iconok
If lIco.Tag > 0 Then
  If Index = 0 Then pIco.Tag = CInt(pIco.Tag) - 1
  If Index = 1 Then pIco.Tag = CInt(pIco.Tag) + 1
  If CInt(pIco.Tag) < 0 Then pIco.Tag = CInt(lIco.Tag)
  If CInt(pIco.Tag) > CInt(lIco.Tag) Then pIco.Tag = 0
  
  lIco.Caption = "(" & pIco.Tag & " \ " & lIco.Tag & ")"
  'If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) else fprg.pKntIco.BackColor = fprg.pKntIco.Tag
  Call mPrg.DrawIco(tIFl.Text, CInt(pIco.Tag))
  pIco.PaintPicture fPrg.pKntIco.Image, 0, 0, pIco.Width / 15, pIco.Height / 15
  
  cApply.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim Ic As Long 'dlya iconok

Call mLng.LoadLang(LangFile, "edt")

If FormNotTop = 0 Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
mDlg.hwndOwner = Me.hwnd

Call Zapolnenie

End Sub

'################################################'
'### SUBS AND FUNCTIONS #########################'
'################################################'

Private Sub Zapolnenie()
cbShw.ListIndex = 0
Me.Caption = " " & wCap & " " & EdtB
tFil.Text = MapB(EdtB).wFil
tPrm.Text = MapB(EdtB).wPrm
tDir.Text = MapB(EdtB).wDir
tIFl.Text = MapB(EdtB).wIFl
tCap.Text = MapB(EdtB).wCap

tHtM.Tag = MapB(EdtB).wHtM: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = MapB(EdtB).wHtK: tHtK.Text = MapKN(Val(tHtK.Tag))
chHtN.Value = MapB(EdtB).wHtN

oOpr(MapB(EdtB).wOpr).Value = True

cbShw.ListIndex = MapB(EdtB).wShw

lIco.Tag = mPrg.GetIcoCount(tIFl.Text)
pIco.Tag = MapB(EdtB).wINm
lIco.Caption = "(" & pIco.Tag & " \ " & lIco.Tag & ")"

chClr.Value = MapB(EdtB).wClr(0)
If chClr.Value = 1 Then
  cClr.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3))
Else
  cClr.BackColor = RGB(MapC(1), MapC(2), MapC(3))
End If

If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
Call mPrg.DrawIco(tIFl.Text, CInt(pIco.Tag))
pIco.PaintPicture fPrg.pKntIco.Image, 0, 0, pIco.Width / 15, pIco.Height / 15

If EdtB > 99 And EdtB < 2200 Then Call DrawSelBtt(1)

cApply.Enabled = False
End Sub

Private Sub SaveSettings()
Dim Color As Long
Dim CC As Integer, RR As Integer
Dim MM As Integer

If oOpr(0).Value = True And tFil.Text <> "" Then
  Call fMsg.GetMsg(fEdt, 1, MapMsg(50), 1)
  If RetMsg = 1 Then oOpr(3).Value = True
End If

If oOpr(0).Value = True Then MapB(EdtB).wOpr = 0
If oOpr(1).Value = True Then MapB(EdtB).wOpr = 1
If oOpr(2).Value = True Then MapB(EdtB).wOpr = 2
If oOpr(3).Value = True Then MapB(EdtB).wOpr = 3
If oOpr(4).Value = True Then MapB(EdtB).wOpr = 4

MapB(EdtB).wFil = tFil.Text
MapB(EdtB).wPrm = tPrm.Text
MapB(EdtB).wDir = tDir.Text
MapB(EdtB).wShw = cbShw.ListIndex
MapB(EdtB).wCap = tCap.Text
MapB(EdtB).wHtN = chHtN.Value
MapB(EdtB).wIFl = tIFl.Text
MapB(EdtB).wINm = CInt(pIco.Tag)
MapB(EdtB).wKmm = ""

If MapB(EdtB).wHtM <> CByte(tHtM.Tag) Or MapB(EdtB).wHtK <> CByte(tHtK.Tag) Then
  MapB(EdtB).wHtM = CByte(tHtM.Tag)
  MapB(EdtB).wHtK = CByte(tHtK.Tag)
End If

MapB(EdtB).wClr(0) = chClr.Value
If chClr.Value = 1 Then
  Color = cClr.BackColor
  MapB(EdtB).wClr(1) = Color Mod 256: Color = Color \ 256
  MapB(EdtB).wClr(2) = Color Mod 256: Color = Color \ 256
  MapB(EdtB).wClr(3) = Color
Else
  MapB(EdtB).wClr(1) = 0
  MapB(EdtB).wClr(2) = 0
  MapB(EdtB).wClr(3) = 0
End If

Call mPrg.SaveStt

If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
If MapB(EdtB).wOpr > 0 Then Call DrawIco(MapB(EdtB).wIFl, MapB(EdtB).wINm): Call DrawHK(MapB(EdtB).wINm)
fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(EdtB).Left + 1, BttCoord(EdtB).Top + 1
Call mPrg.GetActivPic

'Call DrawForm
Call DrawSelBtt(1)
End Sub

Private Sub DrawSelBtt(ByRef wDown As Byte)
Dim CC As Integer, RR As Integer
Dim Btt As RECT

CC = EdtB Mod 100
RR = (EdtB - CC) / 100

CC = CC - 1
RR = RR - 1

Btt = BttCoord(EdtB)
'Btt.Top = RR * (icoH + bttS + 2 + (icoS * 2)) + 3 + bttS
'Btt.Right = icoW + icoS * 2 + 2
'Btt.Bottom = icoH + icoS * 2 + 2

If wDown = 0 Then Call DrawBorder(Btt, 2, MapB(EdtB).wClr(0), MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3))
If wDown = 1 Then Call DrawBorder(Btt, 4, MapB(EdtB).wClr(0), MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3))
End Sub

































Private Sub tCap_Change()
cApply.Enabled = True
End Sub

Private Sub tIFl_Change()
cApply.Enabled = True
End Sub

Private Sub tDir_Change()
cApply.Enabled = True
End Sub

Private Sub tPrm_Change()
cApply.Enabled = True
End Sub

Private Sub tFil_Change()
cApply.Enabled = True
If Left$(tFil.Text, 5) = "lbar_" Then
  lPrm.Visible = False
  lDir.Visible = False
  lPrm2.Visible = True
  lDir2.Visible = True
Else
  lPrm.Visible = True
  lDir.Visible = True
  lPrm2.Visible = False
  lDir2.Visible = False
End If
End Sub

Private Sub tCap_GotFocus()
tCap.SelStart = 0
tCap.SelLength = Len(tCap.Text)
End Sub

Private Sub tDir_GotFocus()
tDir.SelStart = 0
tDir.SelLength = Len(tDir.Text)
End Sub

Private Sub tFil_GotFocus()
tFil.SelStart = 0
tFil.SelLength = Len(tFil.Text)
End Sub

Private Sub tIFl_GotFocus()
tIFl.SelStart = 0
tIFl.SelLength = Len(tIFl.Text)
End Sub

Private Sub tPrm_GotFocus()
tPrm.SelStart = 0
tPrm.SelLength = Len(tPrm.Text)
End Sub
