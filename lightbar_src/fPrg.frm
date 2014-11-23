VERSION 5.00
Begin VB.Form fPrg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Programm v.0.00.000"
   ClientHeight    =   1065
   ClientLeft      =   390
   ClientTop       =   1560
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00644B32&
   Icon            =   "fPrg.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   Tag             =   "1"
   Begin VB.PictureBox pActiv1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00191919&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   3300
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pActiv0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00191919&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   3300
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tFFH 
      Enabled         =   0   'False
      Interval        =   111
      Left            =   3300
      Top             =   75
   End
   Begin VB.PictureBox pKntTime 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   2775
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pMenuIco 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   1050
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Timer tPpp 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2775
      Top             =   75
   End
   Begin VB.Timer tTmr 
      Interval        =   4444
      Left            =   1275
      Top             =   75
   End
   Begin VB.Timer tRun 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2250
      Top             =   525
   End
   Begin VB.Timer tZdr 
      Enabled         =   0   'False
      Left            =   2250
      Top             =   75
   End
   Begin VB.PictureBox pKntIco 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1050
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pKnt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   2775
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tHide 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1725
      Top             =   75
   End
   Begin VB.Timer tShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1725
      Tag             =   "10"
      Top             =   525
   End
   Begin VB.Timer tPrg 
      Interval        =   55
      Left            =   825
      Top             =   75
   End
   Begin VB.CommandButton CCC 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Timer TTT 
      Enabled         =   0   'False
      Interval        =   111
      Left            =   75
      Top             =   525
   End
   Begin VB.Image iTray 
      Height          =   240
      Left            =   525
      Picture         =   "fPrg.frx":08CA
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image iNoIco 
      Height          =   480
      Left            =   525
      Picture         =   "fPrg.frx":170C
      Top             =   525
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "fPrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                              Main Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             06.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Private OldB As Integer    'dlya zapominaniya zydelennoj knopki
Private MDown As Byte      'nazhata li knopka myshy
Private MS As Integer      'dlya raschyota shaga pri raskrytii i skrytii formy
Private BttDown As Integer 'nazhata li knopka klavishej enter ili probel
Private forMove As Integer
Private MM As Integer, mX As Integer, mY As Integer, mB As Integer 'mysh'
Private DragBtt As Integer 'nomer peremecshaemoj knopki

Private Sub CCC_Click()
If NotClearMem = 0 Then fPrg.Hide: Me.WindowState = 1: fPrg.Show: fPrg.Hide: Me.WindowState = 0: fPrg.Show
End Sub

Private Sub Form_Activate()
'If FormNotHide = 0 Then tHide.Enabled = True
Debug.Print vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wX As Integer, wY As Integer
Dim Btt As Integer
If Shift = vbAltMask And KeyCode = vbKeyX Then Call mPrg.ExtPrg
If FormPos = 0 Then
  If fPrg.Top <> 0 Then Exit Sub
Else
  If fPrg.Top <> FB Then Exit Sub
End If
If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
  If OldB = 0 Then
    Btt = 101
  Else
    wX = OldB Mod 100
    wY = (OldB - wX) / 100
    If KeyCode = 37 Then wX = wX - 1 '<
    If KeyCode = 38 Then wY = wY - 1 '^
    If KeyCode = 39 Then wX = wX + 1 '>
    If KeyCode = 40 Then wY = wY + 1 'v
    If wX < 1 Then wX = bttCol
    If wY < 1 Then wY = bttRow
    If wX > bttCol Then wX = 1
    If wY > bttRow Then wY = 1
    Btt = wY * 100 + wX
  End If
  Call DrawBorder(BttCoord(OldB), 2, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
  OldB = Btt
  Call DrawBorder(BttCoord(OldB), 6, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
  If tPpp.Enabled = False Then
    'pechataem v statuse nazvaniya knopok
    If Btt = 0 Then Call DrawText("")
    If Btt > 0 Then DrawText (MapB(Btt).wCap)
  End If
End If
If KeyCode = 13 Or KeyCode = 32 Then
  If OldB > 0 Then
    Call DrawBorder(BttCoord(OldB), 4, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
    BttDown = OldB
    Call Run(OldB, MapKS(KeyCode))
  End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If BttDown > 0 Then
  If KeyCode = 13 Or KeyCode = 32 Then
    If OldB = BttDown Then
      If MapB(OldB).wOpr > 0 Then
        If MapB(OldB).wFil <> "" Then
          Lck = 0
          Call Run(OldB, 3)
        End If
      End If
    End If
  End If
  Call DrawBorder(BttCoord(OldB), 6, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
  BttDown = 0
End If
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
  Call fMsg.GetMsg(fPrg, 0, "LightBar is already running on this system.") 'MapMsg(1)
  End
End If
Randomize Timer
Call mPrg.StrPrg
Lck = 0
If FormPos = 0 Then fPrg.Top = 0 Else fPrg.Top = FB
If FormTop > 0 Then fPrg.Top = FormTop
If FormNotHide = 0 And FormTop = 0 Then Call FullHide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TDrwB As Byte
Lck = 1
If fPrg.ScaleWidth <> frmW Then fPrg.Width = frmW * 15
If fPrg.ScaleHeight <> frmH Then fPrg.Height = frmH * 15

If FormPos = 0 Then
  If fPrg.Top < 0 Then
    If BttToShow = 0 Then
      tShow.Enabled = True
    Else
      If Button = 2 Then
        tShow.Enabled = True
      End If
    End If
    'If Button = BttToShow + 1 Then tShow.Enabled = True
    Exit Sub
  End If
Else
  If fPrg.Top > FB Then
    If BttToShow = 0 Then
      tShow.Enabled = True
    Else
      If Button = 2 Then
        tShow.Enabled = True
      End If
    End If
    Exit Sub
  End If
End If

If MDown = 0 Then
  tPrg.Enabled = True
  OldB = CheckBtt(X, Y)
  TDrwB = 4
  If OldB = -4 Then
    If FormNotHide = 0 Then
      FormNotHide = 1
      TDrwB = 4
    Else
      FormNotHide = 0
      TDrwB = 2
    End If
  End If
  If OldB = -5 Then
    If FormNotTop = 0 Then
      FormNotTop = 1
      TDrwB = 2
      SetWindowPos fPrg.hwnd, -2, 0, 0, 0, 0, &H10 Or &H1 Or &H2
      'Call SetForegroundWindow(fPrg.hwnd)
    Else
      FormNotTop = 0
      TDrwB = 4
      SetWindowPos fPrg.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
  End If
  If OldB = -6 Then
    If FormNotHotKey = 0 Then
      FormNotHotKey = 1
      TDrwB = 2
    Else
      FormNotHotKey = 0
      TDrwB = 4
    End If
  End If
  If OldB > 0 Then
    Call mPrg.DrawBorder(BttCoord(OldB), TDrwB, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
  Else
    Call mPrg.DrawBorder(BttCoord(OldB), TDrwB)
  End If
  MDown = Button
End If
If Button = 1 Then
  MapKS(1) = 1
  If Shift = 0 Then
    Call Run(OldB, 1)
    tRun.Enabled = True
  ElseIf Shift = vbCtrlMask Then
    Call Run(OldB, 200)
  ElseIf Shift = vbShiftMask Then
    Call Run(OldB, 201)
  ElseIf Shift = vbAltMask Then
    DragBtt = OldB
  End If
End If

If CheckBtt(X, Y) = -99 Then forMove = X * 15

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Btt As Integer

'dlya treya
If Lck = 0 Then
  If Button = 0 And Y = 0 And X = 514 Then
    If PoluLck = 0 Then
      If fPrg.ScaleWidth <> frmW Then fPrg.Width = frmW * 15
      If fPrg.ScaleHeight <> frmH Then fPrg.Height = frmH * 15
      fPrg.tShow.Enabled = True
      fPrg.tHide.Enabled = False
      PoluLck = 1
    Else
      fPrg.tHide.Enabled = True
      fPrg.tShow.Enabled = False
      PoluLck = 0
    End If
    Exit Sub
  End If
  If Button = 0 And Y = 0 And X = 517 Then
    fPrg.Tag = 1 - fPrg.Tag
    If NotClearMem = 0 Then fPrg.Hide: Me.WindowState = 1: fPrg.Show: fPrg.Hide:  Me.WindowState = 0: fPrg.Show
    If fPrg.Tag = 1 Then fPrg.Show Else fPrg.Hide
    Exit Sub
  End If
End If

If FormPos = 0 Then
  If fPrg.Top < 0 Then
  '  If tHide.Enabled = True Then tShow.Enabled = True 'esli forma zakryvaetsya, to otkryt' eyo
    If tZdr.Interval > 0 Then 'esli vklyuchena opciya raskrytiya formy pri navedenii myshi
      If X <> 512 Then 'esli vodim mysh'yu ne po treyu
        tZdr.Enabled = True: tPrg.Enabled = True
      End If
    End If
    If tPpp.Enabled = True Then
      tPpp.Enabled = False: tPpp.Enabled = True
    End If
    Exit Sub
  End If
Else
  If fPrg.Top > FB Then
  '  If tHide.Enabled = True Then tShow.Enabled = True 'esli forma zakryvaetsya, to otkryt' eyo
    If tZdr.Interval > 0 Then 'esli vklyuchena opciya raskrytiya formy pri navedenii myshi
      tZdr.Enabled = True: tPrg.Enabled = True
    End If
    If tPpp.Enabled = True Then
      tPpp.Enabled = False: tPpp.Enabled = True
    End If
    Exit Sub
  End If
End If

If BttDown = 0 Then
  mX = X
  mY = Y
  mB = Button
  MM = 1
End If

If Button = 3 Then
  Btt = CheckBtt(X, Y) 'uznaem po kakoj knopke edem
  If Btt = -99 Then 'esli edem po tekstovomu polyu to
    If fPrg.Left + X * 15 - forMove >= 0 Then
      If fPrg.Left + X * 15 - forMove <= Screen.Width - 15 * 15 Then
        fPrg.Left = fPrg.Left + X * 15 - forMove
        FormLeft = fPrg.Left
      Else
        fPrg.Left = Screen.Width - 15 * 15
        FormLeft = fPrg.Left
      End If
    Else
      fPrg.Left = 0
      FormLeft = fPrg.Left
    End If
  End If
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lck = 0
PoluLck = 0
If MDown = Button Then
  tRun.Enabled = False
  MDown = 0
  If OldB > -4 Or OldB < -6 Then Call mPrg.DrawBorder(BttCoord(OldB), 2)
  If OldB = CheckBtt(X, Y) Then 'a eto uzhe znachit klik
    If Button = 1 Then
      If Shift = 0 Then Call Run(OldB, 3)
      If OldB = -1 Then Call mPrg.ExtPrg
      If OldB = -2 Then
        tShow.Enabled = True
        If Lck = 0 Then Lck = 1: fStt.Show 1, Me
        Lck = 0
      End If
      If OldB = -3 Then
        tShow.Enabled = True
        If Lck = 0 Then Lck = 1: fAbt.Show 1, Me
        Lck = 0
      End If
    End If
    If Button = 2 Then
      If Lck = 0 Then
        If OldB > 0 Then
          EdtB = OldB
          If Lck = 0 Then Lck = 1: fEdt.Show 1, Me
          Lck = 0
        End If
      End If
    End If
  Else
    If Shift = vbAltMask Then
      OldB = CheckBtt(X, Y)
      If OldB > 0 Then
        MapB(0) = MapB(OldB)
        MapB(OldB) = MapB(DragBtt)
        MapB(DragBtt) = MapB(0)
        
        EdtB = OldB
        If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
        If MapB(EdtB).wOpr > 0 Then Call DrawIco(MapB(EdtB).wIFl, MapB(EdtB).wINm): Call DrawHK(MapB(EdtB).wINm)
        fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(EdtB).Left + 1, BttCoord(EdtB).Top + 1
        
        EdtB = DragBtt
        If MapB(EdtB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(EdtB).wClr(1), MapB(EdtB).wClr(2), MapB(EdtB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
        If MapB(EdtB).wOpr > 0 Then Call DrawIco(MapB(EdtB).wIFl, MapB(EdtB).wINm): Call DrawHK(MapB(EdtB).wINm)
        fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(EdtB).Left + 1, BttCoord(EdtB).Top + 1
        Call mPrg.GetActivPic
        
        Call SaveStt
      End If
    End If
  End If
  MapKS(1) = 0
End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bB As Integer
Dim Pth As String
Dim SS As String
Dim I As Integer
If Lck = 0 Then
  bB = CheckBtt(X, Y) - 1
1000
  I = I + 1
  bB = bB + 1
  If bB Mod 100 > bttCol Then bB = ((bB \ 100) + 1) * 100
  If bB \ 100 > bttRow Then bB = 101
  
  If Data.GetFormat(vbCFFiles) = True Then
    Pth = GetShortcutPath(Data.Files(I))
  End If
  If Data.GetFormat(vbCFText) = True Then
    Pth = Data.GetData(vbCFText)
  End If
  If Pth <> "" Then
    If bB > 0 Then
      RetMsg = 1
      If MapB(bB).wOpr > 0 Then Call fMsg.GetMsg(fPrg, 1, MapMsg(2) & vbCrLf & "(" & MapB(bB).wFil & ")" & vbCrLf & "(" & Pth & ")", 1)
      Call mPrg.SetCur(bB, 0)
      If RetMsg = 1 Then Call ChangeButton(bB, Pth)
    End If
    If bB = -99 Then
      Clipboard.Clear
      Clipboard.SetText Pth
      Call fMsg.GetMsg(fPrg, 1, MapMsg(3) & vbCrLf & vbCrLf & Pth)
    End If
  End If
  If Data.GetFormat(vbCFFiles) = True Then If Data.Files.Count > I Then GoTo 1000
End If
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Call SetForegroundWindow(fPrg.hwnd)
If FormPos = 0 Then
  If fPrg.Top < 0 Then tShow.Enabled = True
Else
  If fPrg.Top > FB Then tShow.Enabled = True
End If
End Sub

Private Sub tFFH_Timer()
Dim PT As POINTAPI
If Lck = 0 And PoluLck = 0 And FormNotHide = 0 Then
  GetCursorPos PT
  If FormPos = 0 Then If PT.Y < 3 Then tZdr.Enabled = True
  If FormPos = 1 Then If PT.Y > Screen.Height / 15 - 3 Then tZdr.Enabled = True
End If
End Sub

Private Sub tPpp_Timer()
tPpp.Enabled = False
If FormPos = 0 Then
  If fPrg.Top < 0 Then tHide.Enabled = True
Else
  If fPrg.Top > FB Then tHide.Enabled = True
End If
End Sub

Private Sub tRun_Timer()
MapKS(1) = 2
Call Run(OldB, 2)
End Sub

Private Sub tHide_Timer()
tZdr.Enabled = False
tShow.Enabled = False
If FormTop > 0 Then tHide.Enabled = False: Exit Sub
If FormNotHide > 0 Then tShow.Enabled = True: tHide.Enabled = False: Exit Sub
If tPpp.Enabled = True Then
  tPrg.Enabled = False
  tHide.Enabled = False
  Exit Sub
End If
If OldB > 0 Then
  Call DrawBorder(BttCoord(OldB), 2, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
Else
  Call DrawBorder(BttCoord(OldB), 2)
End If
If FormPos = 0 Then
  MS = (fPrg.Top + fPrg.Height) / (CInt(tShow.Tag) + 1)
  If MS < 15 Then MS = 15
  If fPrg.Top - MS < -fPrg.Height + 15 Then
    Call FullHide
  Else
    fPrg.Top = fPrg.Top - MS
  End If
Else
  MS = (Screen.Height - fPrg.Top) / (CInt(tShow.Tag) + 1)
  If MS < 15 Then MS = 15
  If (Screen.Height - fPrg.Top) + MS < 15 Then
    Call FullHide
  Else
    fPrg.Top = fPrg.Top + MS
  End If
End If
End Sub
Private Sub FullHide()
If NotClearMem = 0 Then fPrg.Hide: Me.WindowState = 1: fPrg.Show: fPrg.Hide: Me.WindowState = 0: fPrg.Show
If FormPos = 0 Then
  fPrg.Top = -fPrg.Height + 15
Else
  fPrg.Top = FB + fPrg.Height - 15
End If
Call SetForegroundWindow(GetLastHWnd)
If fPrg.ScaleWidth <> frmW Then fPrg.Width = frmW * 15
If fPrg.ScaleHeight <> frmH Then fPrg.Height = frmH * 15
fPrg.Line (0, 0)-(frmW, 0), ClrFrm
fPrg.Line (0, frmH - 1)-(frmW, frmH - 1), ClrFrm
tHide.Enabled = False
tPrg.Enabled = False
If tFFH.Enabled = True Then Me.Hide
End Sub

Private Sub tShow_Timer()
If FormTop > 0 Then Call SetForegroundWindow(fPrg.hwnd): tShow.Enabled = False: Exit Sub
If Me.Visible = False Then Me.Visible = True
fPrg.Line (0, 0)-(frmW, 0), GenColor(MapC(4), MapC(1), MapC(2), MapC(3))
fPrg.Line (0, frmH - 1)-(frmW, frmH - 1), GenColor(-MapC(4), MapC(1), MapC(2), MapC(3))
tHide.Enabled = False
tPrg.Enabled = True
If OldB > 0 Then
  Call DrawBorder(BttCoord(OldB), 2, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
Else
  Call DrawBorder(BttCoord(OldB), 2)
End If
If FormPos = 0 Then
  MS = -fPrg.Top / (CInt(tShow.Tag) + 1)
  If MS < 15 Then MS = 15
  If fPrg.Top + MS > 0 Then
    Call FullShow
  Else
    fPrg.Top = fPrg.Top + MS
  End If
Else
  MS = (fPrg.Top - FB) / (CInt(tShow.Tag) + 1)
  If MS < 15 Then MS = 15
  If fPrg.Top - FB - MS < 0 Then
    Call FullShow
  Else
    fPrg.Top = fPrg.Top - MS
  End If
End If
End Sub
Private Sub FullShow()
If FormPos = 0 Then
  fPrg.Top = 0
Else
  fPrg.Top = FB
End If
Call SetForegroundWindow(fPrg.hwnd)
If fPrg.ScaleWidth <> frmW Then fPrg.Width = frmW * 15
If fPrg.ScaleHeight <> frmH Then fPrg.Height = frmH * 15
If OldB = 0 Then OldB = 101
If OldB > -4 Or OldB < -6 Then Call DrawBorder(BttCoord(OldB), 6)
tShow.Enabled = False
End Sub

Private Sub tPrg_Timer()
Dim RC As RECT, PT As POINTAPI
Dim Btt As Integer

If Lck = 0 And PoluLck = 0 And FormNotHide = 0 Then
  GetCursorPos PT
  GetWindowRect Me.hwnd, RC
  RC.Left = RC.Left - 20
  RC.Top = RC.Top - 20
  RC.Right = RC.Right + 20
  RC.Bottom = RC.Bottom + 20
  If PT.X < RC.Left Or PT.X > RC.Right Or PT.Y < RC.Top Or PT.Y > RC.Bottom Then
    If FormPos = 0 Then If PT.Y > 3 Then tHide.Enabled = True
    If FormPos = 1 Then If PT.Y < Screen.Height / 15 - 3 Then tHide.Enabled = True
  End If
End If
If TimeNotShow = 0 And pKntTime.Tag <> time Then
  fPrg.pKntTime.Cls
  fPrg.pKntTime.CurrentY = FntTop
  fPrg.pKntTime.Print time
  fPrg.PaintPicture fPrg.pKntTime.Image, fPrg.pKntTime.Left, fPrg.pKntTime.Top
  
  'fPrg.Line (frmW - (txtW + 5 + MBttW), pKnt.Top + FntTop + 2)-(frmW - 10 - MBttW, pKnt.Top + FntTop + 10), MapC(1), BF
  'fPrg.CurrentX = frmW - (txtW + 6 + MBttW): fPrg.CurrentY = pKnt.Top + FntTop
  'fPrg.Print time 'Date & " " & time
  pKntTime.Tag = time
End If

If MM = 1 Then
  Btt = CheckBtt(mX, mY) 'uznaem po kakoj knopke edem
  If mB = 0 Then
    If OldB <> Btt Then
      If OldB > -4 Or OldB < -6 Then
        If OldB > 0 Then
          Call DrawBorder(BttCoord(OldB), 2, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
        Else
          Call DrawBorder(BttCoord(OldB), 2)
        End If
      Else
        If OldB = -4 Then If FormNotHide = 0 Then Call DrawBorder(BttCoord(OldB), 2) Else Call DrawBorder(BttCoord(OldB), 4)
        If OldB = -5 Then If FormNotTop = 0 Then Call DrawBorder(BttCoord(OldB), 4) Else Call DrawBorder(BttCoord(OldB), 2)
        If OldB = -6 Then If FormNotHotKey = 0 Then Call DrawBorder(BttCoord(OldB), 4) Else Call DrawBorder(BttCoord(OldB), 2)
      End If
      OldB = Btt
      If OldB > -4 Or OldB < -6 Then
        If OldB > 0 Then
          Call DrawBorder(BttCoord(OldB), 6, MapB(OldB).wClr(0), MapB(OldB).wClr(1), MapB(OldB).wClr(2), MapB(OldB).wClr(3))
        Else
          Call DrawBorder(BttCoord(OldB), 6)
        End If
      Else
        If OldB = -4 Then If FormNotHide = 0 Then Call DrawBorder(BttCoord(OldB), 6) Else Call DrawBorder(BttCoord(OldB), 7)
        If OldB = -5 Then If FormNotTop = 0 Then Call DrawBorder(BttCoord(OldB), 7) Else Call DrawBorder(BttCoord(OldB), 6)
        If OldB = -6 Then If FormNotHotKey = 0 Then Call DrawBorder(BttCoord(OldB), 7) Else Call DrawBorder(BttCoord(OldB), 6)
      End If
    End If
  End If
  If tPpp.Enabled = False Then
    'pechataem v statuse nazvaniya knopok
    If Btt = 0 Then Call DrawText("")
    If Btt = -1 Then DrawText (MapOth(0))
    If Btt = -2 Then DrawText (MapOth(1))
    If Btt = -3 Then DrawText (MapOth(2))
    If Btt = -4 Then If FormNotHide = 0 Then DrawText (MapOth(3)) Else DrawText (MapOth(4))
    If Btt = -5 Then If FormNotTop = 0 Then DrawText (MapOth(5)) Else DrawText (MapOth(6))
    If Btt = -6 Then If FormNotHotKey = 0 Then DrawText (MapOth(7)) Else DrawText (MapOth(8))
    If Btt > 0 Then DrawText (MapB(Btt).wCap)
  End If
  MM = 0
End If
End Sub

Private Sub tTmr_Timer()
If Lck = 0 Then
  If fPrg.Left <> FormLeft Then fPrg.Left = FormLeft
  If FormPos = 0 Then
    If fPrg.Top <> -fPrg.Height + 15 Then
      If tPrg.Enabled = False Then
        If NotClearMem = 0 Then fPrg.Hide: Me.WindowState = 1: fPrg.Show: fPrg.Hide:  Me.WindowState = 0: fPrg.Show
        tPrg.Enabled = True
      End If
    End If
  Else
    If fPrg.Top <> Screen.Height - 15 Then
      FB = Screen.Height - fPrg.Height
      If tPrg.Enabled = False Then
        If NotClearMem = 0 Then fPrg.Hide: Me.WindowState = 1: fPrg.Show: fPrg.Hide:  Me.WindowState = 0: fPrg.Show
        tPrg.Enabled = True
      End If
    End If
  End If
End If
End Sub

Private Sub TTT_Timer()
'Debug.Print "": Debug.Print "": Debug.Print "": Debug.Print "": Debug.Print "": Debug.Print "":
'Debug.Print "tprg  - " & tPrg.Enabled
'Debug.Print "ttmr  - " & tTmr.Enabled
'Debug.Print "thide - " & tHide.Enabled
'Debug.Print "tshow - " & tShow.Enabled
'Debug.Print "tzdr  - " & tZdr.Enabled
'Debug.Print "trun  - " & tRun.Enabled
'Debug.Print "tppp  - " & tPpp.Enabled
CCC.Caption = ""
CCC.Caption = CCC.Caption & "tprg  - " & tPrg.Enabled & vbCrLf
CCC.Caption = CCC.Caption & "ttmr  - " & tTmr.Enabled & vbCrLf
CCC.Caption = CCC.Caption & "thide - " & tHide.Enabled & vbCrLf
CCC.Caption = CCC.Caption & "tshow - " & tShow.Enabled & vbCrLf
CCC.Caption = CCC.Caption & "tzdr  - " & tZdr.Enabled & vbCrLf
CCC.Caption = CCC.Caption & "trun  - " & tRun.Enabled & vbCrLf
CCC.Caption = CCC.Caption & "tppp  - " & tPpp.Enabled & vbCrLf
End Sub

Private Sub tZdr_Timer()
Dim PT As POINTAPI

If Lck = 0 And PoluLck = 0 And FormNotHide = 0 Then
  GetCursorPos PT
  If FormPos = 0 Then If PT.Y < 3 Then tShow.Enabled = True
  If FormPos = 1 Then If PT.Y > Screen.Height / 15 - 3 Then tShow.Enabled = True
End If

tZdr.Enabled = False
End Sub

'################################################'
'### SUBS AND FUNCTIONS #########################'
'################################################'

Public Sub Run(ByVal wBtt As Integer, ByVal wAct As Byte)
Static chkDblClick As Integer
Dim SERet As Long
Dim twShw As Byte

If wBtt <= 0 Then Exit Sub

If MapB(wBtt).wOpr > 0 Then
  If MapB(wBtt).wFil <> "" Then
    If wAct < 200 Then
      If (MapB(wBtt).wOpr = wAct) Or (wAct = 1 And MapB(wBtt).wOpr = 2) Then
        If Left$(MapB(wBtt).wFil, 5) = "lbar_" Then
          Call mCmd.CommandParser(MapB(wBtt).wFil, MapB(wBtt).wPrm, MapB(wBtt).wDir)
          If wAct = 3 Then tHide.Enabled = True
        Else
          Call SetForegroundWindow(fPrg.hwnd) 'AppActivate fPrg.Caption
          If wAct = 3 Then twShw = MapB(wBtt).wShw + 1 Else twShw = 4
          SERet = ShellExecute(0&, "open", GetDir(MapB(wBtt).wFil), MapB(wBtt).wPrm, GetDir(MapB(wBtt).wDir), twShw)
          If SERet > -1 And SERet < 33 Then Call fMsg.GetMsg(fPrg, 0, mPrg.GetError(0, SERet)): Call mPrg.SetCur(wBtt, 0)
          If SERet > 32 And wAct = 3 Then tHide.Enabled = True
        End If
      End If
    Else
      If Left$(MapB(wBtt).wFil, 5) = "lbar_" Then
        Call fMsg.GetMsg(fPrg, 0, MapMsg(4))
      Else
        Call SetForegroundWindow(fPrg.hwnd) 'AppActivate fPrg.Caption
        If wAct = 200 Then SERet = ShellExecute(0&, "open", GetFolder(GetDir(MapB(wBtt).wFil)), "", "", 1)    'eto znachit nada otkryt' papku
        If wAct = 201 Then SERet = ShellExecute(0&, "explore", GetFolder(GetDir(MapB(wBtt).wFil)), "", "", 1) 'eto znachit nada otkryt' papku s derevom
        If SERet > -1 And SERet < 33 Then Call fMsg.GetMsg(fPrg, 0, mPrg.GetError(0, SERet)): Call mPrg.SetCur(wBtt, 0)
        If SERet > 32 And wAct = 3 Then tHide.Enabled = True
      End If
    End If
  End If
End If

'proverka dvojnogo nazhatiya
If wAct = 3 Then
  If wBtt = chkDblClick Then
    chkDblClick = 0
    Call Run(wBtt, 4)
  Else
    chkDblClick = wBtt
  End If
End If
End Sub

Private Sub DrawText(ByRef wStr As String)
pKnt.Cls
pKnt.CurrentX = 0 '-1
pKnt.CurrentY = FntTop
pKnt.Print wStr
fPrg.PaintPicture pKnt.Image, pKnt.Left, pKnt.Top
End Sub

Private Sub ChangeButton(ByRef wBB As Integer, ByRef wPth As String)
Dim SS As String
Dim FF As Long

MapB(wBB).wOpr = 3
MapB(wBB).wFil = wPth
MapB(wBB).wPrm = ""
MapB(wBB).wDir = GetFolder(wPth)
MapB(wBB).wShw = 0

MapB(wBB).wCap = GetName(wPth)
MapB(wBB).wHtM = 0
MapB(wBB).wHtK = 0
MapB(wBB).wHtN = 0

On Error GoTo 1000
  If GetAttr(wPth) < 16 Or GetAttr(wPth) > 31 Then 'esli put' - ne papka
    MapB(wBB).wIFl = wPth
    If GetIcoCount(MapB(wBB).wIFl) > 0 Then MapB(wBB).wINm = 1 Else MapB(wBB).wINm = 0
  Else
    MapB(wBB).wIFl = "%systemroot%\system32\shell32.dll"
    MapB(wBB).wINm = 5
    
    If Dir(wPth & "\Desktop.ini", 39) <> "" Then
      FF = FreeFile
      Open wPth & "\Desktop.ini" For Input As #FF
        Do
          If EOF(FF) = True Then Exit Do
          Input #FF, SS
          If Left$(Trim(SS), 8) = "IconFile" Then
            MapB(wBB).wIFl = Trim(Right$(SS, Len(SS) - InStr(SS, "=")))
          End If
          If Left$(Trim(SS), 9) = "IconIndex" Then
            MapB(wBB).wINm = Val(Trim(Right$(SS, Len(SS) - InStr(SS, "=")))) + 1
          End If
        Loop
      Close #FF
    End If
  End If
1000

If MapB(wBB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(wBB).wClr(1), MapB(wBB).wClr(2), MapB(wBB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
Call DrawIco(MapB(wBB).wIFl, MapB(wBB).wINm): Call DrawHK(MapB(wBB).wINm)
fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(wBB).Left + 1, BttCoord(wBB).Top + 1
Call mPrg.GetActivPic

MapB(wBB).wKmm = ""
Call mPrg.SaveStt
End Sub


































