VERSION 5.00
Begin VB.Form fMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Информация"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fMsg.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cCancel 
      Height          =   465
      Left            =   75
      Picture         =   "fMsg.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cOK 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   75
      Picture         =   "fMsg.frx":0FB4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2475
      Width           =   465
   End
   Begin VB.TextBox tMsg 
      Height          =   2865
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   75
      Width           =   4590
   End
   Begin VB.Image iIco 
      Height          =   480
      Index           =   2
      Left            =   75
      Picture         =   "fMsg.frx":169E
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image iIco 
      Height          =   480
      Index           =   1
      Left            =   75
      Picture         =   "fMsg.frx":1F68
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image iIco 
      Height          =   480
      Index           =   0
      Left            =   75
      Picture         =   "fMsg.frx":2832
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                          Messages Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             11.05.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Private PrevFormName As String

Private Sub cCancel_Click()
RetMsg = 0
Unload Me
End Sub

Private Sub cOK_Click()
RetMsg = 1
Unload Me
End Sub

Private Sub Form_Activate()

Call mLng.LoadLang(LangFile, "msg")

If cCancel.Visible = True Then Call mPrg.SetCur(cCancel.hwnd) Else Call mPrg.SetCur(cOK.hwnd)
If iIco(0).Visible = True Then Call mPrg.BeeBeep(0)
If iIco(1).Visible = True Then Call mPrg.BeeBeep(1)
If iIco(2).Visible = True Then Call mPrg.BeeBeep(2)
End Sub

Private Sub Form_Load()
iIco(0).Visible = False
iIco(1).Visible = False
iIco(2).Visible = False
If FormNotTop = 0 Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
End Sub

Private Sub Form_Unload(Cancel As Integer)
If PrevFormName = "fPrg" Then Lck = 0
End Sub

'################################################'
'### SUBS AND FUNCTIONS #########################'
'################################################'

Public Sub GetMsg(ByRef wForm As Form, ByRef wIco As Byte, ByRef wText As String, Optional ByRef wCancel As Byte = 0)
tMsg.Text = wText
If wIco >= 0 And wIco <= 2 Then iIco(wIco).Visible = True
If wCancel = 0 Then
  cCancel.Visible = False
Else
  cCancel.Visible = True
  cCancel.Cancel = True
End If
Lck = 1
PrevFormName = wForm.Name
fMsg.Show 1, wForm
End Sub
