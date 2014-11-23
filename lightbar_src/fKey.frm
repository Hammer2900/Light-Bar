VERSION 5.00
Begin VB.Form fKey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Форма отлова клавиш"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fKey.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label infInf 
      Alignment       =   2  'Center
      Caption         =   "Нажмите нужную клавишу или нажмите Esc для отмены"
      Height          =   465
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   2490
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left Shift"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left Ctrl"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   6
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left Alt"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   75
      TabIndex        =   5
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right Shift"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1350
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right Ctrl"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   1350
      TabIndex        =   3
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right Alt"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   1350
      TabIndex        =   2
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left Win"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   75
      TabIndex        =   1
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label lMod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right Win"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   1350
      TabIndex        =   0
      Top             =   1050
      Width           =   1215
   End
End
Attribute VB_Name = "fKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                      Fishing Keys Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             18.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Dim OldFormNotHotKey As Byte

Private Sub Form_Activate()
Call mPrg.SetCur(Me.hwnd)
End Sub

Private Sub Form_Load()
Dim I As Integer

Call mLng.LoadLang(LangFile, "key")

If FormNotTop = 0 Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2

OldFormNotHotKey = FormNotHotKey
FormNotHotKey = 0

wfKey = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
FormNotHotKey = OldFormNotHotKey
End Sub
