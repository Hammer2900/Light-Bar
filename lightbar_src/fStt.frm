VERSION 5.00
Begin VB.Form fStt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Настройки программы"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fStt.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton oPag 
      Caption         =   "Программа"
      Height          =   315
      Index           =   3
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Value           =   -1  'True
      Width           =   1290
   End
   Begin VB.OptionButton oPag 
      Caption         =   "Цвета"
      Height          =   315
      Index           =   2
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   75
      Width           =   1290
   End
   Begin VB.OptionButton oPag 
      Caption         =   "Схема"
      Height          =   315
      Index           =   1
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Width           =   1290
   End
   Begin VB.OptionButton oPag 
      Caption         =   "Окно"
      Height          =   315
      Index           =   0
      Left            =   1425
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton cDefault 
      Caption         =   "По умолчанию"
      Height          =   465
      Left            =   2700
      TabIndex        =   83
      Top             =   4425
      Width           =   1440
   End
   Begin VB.CommandButton cApply 
      Caption         =   "Применить"
      Enabled         =   0   'False
      Height          =   465
      Left            =   4200
      TabIndex        =   82
      Top             =   4425
      Width           =   1215
   End
   Begin VB.CommandButton cCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   465
      Left            =   1350
      TabIndex        =   0
      Top             =   4425
      Width           =   1215
   End
   Begin VB.CommandButton cOK 
      Caption         =   "ОК"
      Height          =   465
      Left            =   75
      TabIndex        =   84
      Top             =   4425
      Width           =   1215
   End
   Begin VB.Frame frPag 
      Caption         =   "Программа"
      Height          =   3915
      Index           =   3
      Left            =   75
      TabIndex        =   108
      Top             =   450
      Width           =   5340
      Begin VB.CheckBox chNotClearMem 
         Caption         =   "Не очищать память"
         Height          =   240
         Left            =   75
         TabIndex        =   8
         Top             =   675
         Width           =   5190
      End
      Begin VB.OptionButton oStt 
         Caption         =   "В папке с программой"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   19
         Top             =   3000
         Width           =   2565
      End
      Begin VB.OptionButton oStt 
         Caption         =   "Абсолютный путь"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   21
         Top             =   3225
         Width           =   2565
      End
      Begin VB.OptionButton oStt 
         Caption         =   "В ""Мои документы"""
         Height          =   240
         Index           =   2
         Left            =   2700
         TabIndex        =   20
         Top             =   3000
         Width           =   2565
      End
      Begin VB.OptionButton oStt 
         Caption         =   "В ""Application Data"""
         Height          =   240
         Index           =   3
         Left            =   2700
         TabIndex        =   22
         Top             =   3225
         Width           =   2565
      End
      Begin VB.TextBox tStt 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3525
         Width           =   4440
      End
      Begin VB.CommandButton cStt 
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   4575
         TabIndex        =   24
         ToolTipText     =   "Создать новый"
         Top             =   3525
         Width           =   315
      End
      Begin VB.CommandButton cStt 
         Caption         =   "="
         Height          =   315
         Index           =   1
         Left            =   4950
         TabIndex        =   25
         ToolTipText     =   "Указать существующий"
         Top             =   3525
         Width           =   315
      End
      Begin VB.TextBox tLngInf 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2400
         Width           =   2790
      End
      Begin VB.TextBox tLngInf 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   840
      End
      Begin VB.TextBox tLngInf 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2175
         Width           =   4290
      End
      Begin VB.ComboBox cbLng 
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1800
         Width           =   4440
      End
      Begin VB.CheckBox chLng 
         Caption         =   "Указать языковой файл"
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   1575
         Width           =   4440
      End
      Begin VB.CommandButton cGenLng 
         Caption         =   "Gen newlng"
         Height          =   465
         Left            =   4575
         TabIndex        =   15
         Top             =   1650
         Width           =   690
      End
      Begin VB.CheckBox chNotAutoFocus 
         Caption         =   "Не фокусировать курсор мыши на элементах"
         Height          =   240
         Left            =   75
         TabIndex        =   7
         Top             =   450
         Width           =   5190
      End
      Begin VB.TextBox tHtK 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   1665
      End
      Begin VB.TextBox tHtM 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2715
      End
      Begin VB.CommandButton cHtKDel 
         DownPicture     =   "fStt.frx":08CA
         Height          =   315
         Left            =   4950
         Picture         =   "fStt.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton cHtK 
         DownPicture     =   "fStt.frx":0B5E
         Height          =   315
         Left            =   4575
         Picture         =   "fStt.frx":0CA8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Определить комбинацию клавиш"
         Top             =   1200
         Width           =   315
      End
      Begin VB.CheckBox chShowInTray 
         Caption         =   "Поместить значок в трей"
         Height          =   240
         Left            =   2850
         TabIndex        =   6
         Top             =   225
         Width           =   2415
      End
      Begin VB.CheckBox chAutoStart 
         Caption         =   "Запускать вместе с Windows"
         Height          =   240
         Left            =   75
         TabIndex        =   5
         Top             =   225
         Width           =   2715
      End
      Begin VB.Label infStt 
         Caption         =   "Хранить файл настроек в:"
         Height          =   240
         Left            =   75
         TabIndex        =   118
         Top             =   2775
         Width           =   5190
      End
      Begin VB.Label lLngInf 
         Caption         =   "Author:"
         Height          =   240
         Index           =   4
         Left            =   1875
         TabIndex        =   113
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label lLngInf 
         Caption         =   "For version:"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   112
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label lLngInf 
         Caption         =   "Language:"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   111
         Top             =   2175
         Width           =   915
      End
      Begin VB.Label infHtK 
         Caption         =   "Горячая клавиша для открытия основного окна:"
         Height          =   240
         Left            =   75
         TabIndex        =   109
         Top             =   975
         Width           =   5190
      End
   End
   Begin VB.Frame frPag 
      Caption         =   "Окно"
      Height          =   3915
      Index           =   0
      Left            =   75
      TabIndex        =   88
      Top             =   450
      Visible         =   0   'False
      Width           =   5340
      Begin VB.TextBox tOts 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   44
         Text            =   "000"
         Top             =   3525
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   23
         Left            =   4950
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3525
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   22
         Left            =   4125
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3525
         Width           =   315
      End
      Begin VB.CheckBox chFormFullHide 
         Caption         =   "Скрывать главное окно полностью"
         Height          =   240
         Left            =   75
         TabIndex        =   36
         Top             =   2550
         Width           =   5190
      End
      Begin VB.TextBox tFnt 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2550
         TabIndex        =   41
         Text            =   "000"
         Top             =   3525
         Width           =   540
      End
      Begin VB.TextBox tAnm 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   37
         Text            =   "000"
         Top             =   2925
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   21
         Left            =   3075
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3525
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   20
         Left            =   2250
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3525
         Width           =   315
      End
      Begin VB.CheckBox chFntI 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3525
         Width           =   315
      End
      Begin VB.CheckBox chFntB 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3450
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3525
         Width           =   315
      End
      Begin VB.ComboBox cbFnt 
         Height          =   315
         Left            =   75
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3525
         Width           =   2115
      End
      Begin VB.CheckBox chScreenBottom 
         Caption         =   "Расположить в низу экрана"
         Height          =   240
         Left            =   75
         TabIndex        =   26
         Top             =   225
         Width           =   2790
      End
      Begin VB.CheckBox chTimeNotShow 
         Caption         =   "Не показывать часы"
         Height          =   240
         Left            =   75
         TabIndex        =   27
         Top             =   450
         Width           =   2790
      End
      Begin VB.CheckBox chSKM 
         Caption         =   "Появляться от СКМ"
         Height          =   240
         Left            =   2925
         TabIndex        =   29
         Top             =   450
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.CheckBox chPKM 
         Caption         =   "Разворачиваться от ПКМ"
         Height          =   240
         Left            =   2925
         TabIndex        =   28
         Top             =   225
         Width           =   2340
      End
      Begin VB.HScrollBar sPolV 
         Height          =   315
         LargeChange     =   50
         Left            =   1275
         Max             =   255
         TabIndex        =   31
         Top             =   1125
         Width           =   3540
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   7
         Left            =   4950
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2925
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   6
         Left            =   4125
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2925
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   5
         Left            =   4950
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2175
         Width           =   315
      End
      Begin VB.TextBox tZdr 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   33
         Text            =   "000"
         Top             =   2175
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   4
         Left            =   4125
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2175
         Width           =   315
      End
      Begin VB.HScrollBar sPol 
         Height          =   315
         LargeChange     =   3276
         Left            =   1275
         TabIndex        =   30
         Top             =   750
         Width           =   3540
      End
      Begin VB.HScrollBar sTrans 
         Height          =   315
         LargeChange     =   50
         Left            =   1275
         Max             =   255
         TabIndex        =   32
         Top             =   1725
         Value           =   200
         Width           =   3540
      End
      Begin VB.Label infOts 
         Caption         =   "Отступ:"
         Height          =   240
         Left            =   4125
         TabIndex        =   114
         Top             =   3300
         Width           =   1140
      End
      Begin VB.Label infFnt 
         Caption         =   "Шрифт в главном окне:"
         Height          =   240
         Left            =   75
         TabIndex        =   110
         Top             =   3300
         Width           =   3990
      End
      Begin VB.Label lTrans 
         Caption         =   "200"
         Height          =   240
         Left            =   4875
         TabIndex        =   107
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label lPol 
         Caption         =   "0"
         Height          =   240
         Left            =   4875
         TabIndex        =   106
         Top             =   825
         Width           =   390
      End
      Begin VB.Label lPolV 
         Caption         =   "0"
         Height          =   240
         Left            =   4875
         TabIndex        =   105
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label infVrI 
         Alignment       =   2  'Center
         Caption         =   "(Если больше 0, то скрытие не будет работать!)"
         Enabled         =   0   'False
         Height          =   240
         Left            =   1200
         TabIndex        =   104
         Top             =   1425
         Width           =   3690
      End
      Begin VB.Label infVrt 
         Caption         =   "Положение по вертикали:"
         Height          =   465
         Left            =   75
         TabIndex        =   103
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label infAnm 
         Caption         =   "Скорость анимации при открытии главного окна (0 - не использовать)(мс.):"
         Height          =   390
         Left            =   75
         TabIndex        =   95
         Top             =   2850
         Width           =   4065
      End
      Begin VB.Label infZdr 
         Caption         =   "Задержка раскрытия основного окна (0 - только по щелчку) (мс.):"
         Height          =   390
         Left            =   75
         TabIndex        =   94
         Top             =   2100
         Width           =   4065
      End
      Begin VB.Label infPol 
         Caption         =   "Положение:"
         Height          =   240
         Left            =   75
         TabIndex        =   90
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label infTrn 
         Caption         =   "Прозрачность:"
         Height          =   240
         Left            =   75
         TabIndex        =   89
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame frPag 
      Caption         =   "Схема"
      Height          =   3915
      Index           =   1
      Left            =   75
      TabIndex        =   85
      Top             =   450
      Visible         =   0   'False
      Width           =   5340
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   15
         Left            =   4950
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   1725
         Width           =   315
      End
      Begin VB.TextBox tIcoS 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   64
         Text            =   "000"
         Top             =   1725
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   14
         Left            =   4125
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1725
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   13
         Left            =   4950
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1425
         Width           =   315
      End
      Begin VB.TextBox tBttS 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   61
         Text            =   "000"
         Top             =   1425
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   12
         Left            =   4125
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1425
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   11
         Left            =   4950
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1125
         Width           =   315
      End
      Begin VB.TextBox tIcoH 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   58
         Text            =   "000"
         Top             =   1125
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   10
         Left            =   4125
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1125
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   9
         Left            =   4950
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   825
         Width           =   315
      End
      Begin VB.TextBox tIcoW 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   55
         Text            =   "000"
         Top             =   825
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   8
         Left            =   4125
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   3
         Left            =   4950
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox tRow 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   52
         Text            =   "000"
         Top             =   525
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   2
         Left            =   4125
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   525
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   19
         Left            =   4950
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2400
         Width           =   315
      End
      Begin VB.TextBox tMBttH 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   70
         Text            =   "000"
         Top             =   2400
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   18
         Left            =   4125
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2400
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   17
         Left            =   4950
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2100
         Width           =   315
      End
      Begin VB.TextBox tMBttW 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   67
         Text            =   "000"
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   16
         Left            =   4125
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chDrawHK 
         Caption         =   "Печатать поверх иконки ""горячую клавишу"""
         Height          =   240
         Left            =   75
         TabIndex        =   73
         Top             =   3600
         Width           =   5190
      End
      Begin VB.TextBox tCol 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4425
         TabIndex        =   49
         Text            =   "000"
         Top             =   225
         Width           =   540
      End
      Begin VB.CommandButton cScr 
         Caption         =   "<"
         Height          =   315
         Index           =   0
         Left            =   4125
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cScr 
         Caption         =   ">"
         Height          =   315
         Index           =   1
         Left            =   4950
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   225
         Width           =   315
      End
      Begin VB.Label infIcS 
         Caption         =   "Расстояние иконки от края кнопки:"
         Height          =   240
         Left            =   75
         TabIndex        =   99
         Top             =   1800
         Width           =   4065
      End
      Begin VB.Label infBtS 
         Caption         =   "Отступ между кнопками:"
         Height          =   240
         Left            =   75
         TabIndex        =   98
         Top             =   1500
         Width           =   4065
      End
      Begin VB.Label infIcH 
         Caption         =   "Высота иконок:"
         Height          =   240
         Left            =   75
         TabIndex        =   97
         Top             =   1200
         Width           =   4065
      End
      Begin VB.Label infIcW 
         Caption         =   "Ширина иконок:"
         Height          =   240
         Left            =   75
         TabIndex        =   96
         Top             =   900
         Width           =   4065
      End
      Begin VB.Label infMBH 
         Caption         =   "Высота кнопок меню:"
         Height          =   240
         Left            =   75
         TabIndex        =   102
         Top             =   2475
         Width           =   4065
      End
      Begin VB.Label infMBW 
         Caption         =   "Ширина кнопок меню:"
         Height          =   240
         Left            =   75
         TabIndex        =   101
         Top             =   2175
         Width           =   4065
      End
      Begin VB.Label infCol 
         Caption         =   "Столбцы:"
         Height          =   240
         Left            =   75
         TabIndex        =   87
         Top             =   300
         Width           =   4065
      End
      Begin VB.Label infRow 
         Caption         =   "Строки:"
         Height          =   240
         Left            =   75
         TabIndex        =   86
         Top             =   600
         Width           =   4065
      End
   End
   Begin VB.Frame frPag 
      Caption         =   "Цвета"
      Height          =   3915
      Index           =   2
      Left            =   75
      TabIndex        =   91
      Top             =   450
      Visible         =   0   'False
      Width           =   5340
      Begin VB.CommandButton cClr 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1725
         Width           =   315
      End
      Begin VB.CheckBox chSlB 
         Caption         =   "Выделять только бордюр"
         Height          =   240
         Left            =   75
         TabIndex        =   78
         Top             =   1800
         Width           =   4815
      End
      Begin VB.PictureBox pActiv 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3150
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   66
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   2550
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cGenColors 
         Caption         =   "Сгенерировать"
         Height          =   465
         Left            =   75
         TabIndex        =   80
         Top             =   3375
         Width           =   1440
      End
      Begin VB.HScrollBar sDown 
         Height          =   315
         LargeChange     =   20
         Left            =   2625
         Max             =   100
         TabIndex        =   77
         Top             =   1350
         Value           =   50
         Width           =   2265
      End
      Begin VB.CommandButton cClr 
         Height          =   315
         Index           =   10
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   975
         Width           =   315
      End
      Begin VB.PictureBox pPrv 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4200
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   66
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   2550
         Width           =   990
      End
      Begin VB.CommandButton cClr 
         Height          =   315
         Index           =   1
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cClr 
         Height          =   315
         Index           =   0
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   225
         Width           =   315
      End
      Begin VB.Label lDown 
         Caption         =   "000"
         Height          =   240
         Left            =   4950
         TabIndex        =   116
         Top             =   1425
         Width           =   315
      End
      Begin VB.Label infDep 
         Caption         =   "Глубина:"
         Height          =   240
         Left            =   75
         TabIndex        =   115
         Top             =   1425
         Width           =   2565
      End
      Begin VB.Label infClH 
         Caption         =   "Цвет свёрнутого окна:"
         Height          =   240
         Left            =   75
         TabIndex        =   100
         Top             =   1050
         Width           =   4815
      End
      Begin VB.Label infCl1 
         Caption         =   "Фон:"
         Height          =   240
         Left            =   75
         TabIndex        =   93
         Top             =   675
         Width           =   4815
      End
      Begin VB.Label infCl0 
         Caption         =   "Шрифт:"
         Height          =   240
         Left            =   75
         TabIndex        =   92
         Top             =   300
         Width           =   4815
      End
   End
End
Attribute VB_Name = "fStt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################'
'# Programm:                           LightBar #'
'# Part:                          Settings Form #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             06.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Private LoadOK As Byte
Private cR As Integer, cG As Integer, cB As Integer
Private cRA As Integer, cGA As Integer, cBA As Integer

Private Sub cApply_Click()
Call SaveSettings
cApply.Enabled = False
End Sub

Private Sub cbFnt_Click()
Call GetPrevievText
End Sub

Private Sub cbFnt_KeyUp(KeyCode As Integer, Shift As Integer)
Call GetPrevievText
End Sub

Private Sub cbLng_Click()
cApply.Enabled = True
Call LoadLngInf
End Sub

Private Sub cbLng_KeyDown(KeyCode As Integer, Shift As Integer)
cApply.Enabled = True
Call LoadLngInf
End Sub

Private Sub cCancel_Click()
fPrg.Left = CInt(sPol.Tag) * 15
TransForm = CInt(sTrans.Tag)
Call SetTransparent(fPrg.hwnd, CByte(TransForm), 1)

fPrg.pKnt.FontName = FntName
fPrg.pKnt.FontSize = FntSize
fPrg.pKnt.FontBold = FntBold
fPrg.pKnt.FontItalic = FntItalic
fPrg.pKntTime.FontName = FntName
fPrg.pKntTime.FontSize = FntSize
fPrg.pKntTime.FontBold = FntBold
fPrg.pKntTime.FontItalic = FntItalic

LoadOK = 0
Call mPrg.SetCur(-2, 0)
Unload Me
End Sub

Private Sub cDefault_Click()
Dim I As Integer

chAutoStart.Value = 0
chScreenBottom.Value = 0
sPol.Value = 50
sPolV.Value = 0
sTrans.Value = 200
tZdr.Text = 0
tAnm.Text = 10
tHtM.Tag = 68: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = 81: tHtK.Text = MapKN(Val(tHtK.Tag))

tCol.Text = 20
tRow.Text = 5
tIcoW.Text = 16
tIcoH.Text = 16
tBttS.Text = 1
tIcoS.Text = 0

tMBttW.Text = 11
tMBttH.Text = 11

cClr(0).BackColor = RGB(0, 25, 50)
cClr(1).BackColor = RGB(150, 175, 200)
cClr(10).BackColor = RGB(200, 0, 0)
cClr(2).BackColor = RGB(200, 200, 0)
Call GetRGB
Call DrawPreviev

FntName = "MS Sans Serif"
FntSize = 8
FntBold = 1
FntItalic = 0
FntTop = -2

fPrg.pKnt.FontName = FntName
fPrg.pKnt.FontSize = FntSize
fPrg.pKnt.FontBold = FntBold
fPrg.pKnt.FontItalic = FntItalic
fPrg.pKntTime.FontName = FntName
fPrg.pKntTime.FontSize = FntSize
fPrg.pKntTime.FontBold = FntBold
fPrg.pKntTime.FontItalic = FntItalic

tFnt.FontName = FntName
tFnt.FontSize = FntSize
tFnt.FontBold = FntBold
tFnt.FontItalic = FntItalic

For I = 0 To cbFnt.ListCount - 1
  If cbFnt.List(I) = FntName Then cbFnt.ListIndex = I: Exit For
Next I
tFnt.Text = FntSize
chFntB.Value = FntBold
chFntI.Value = FntItalic
tOts.Text = FntTop

End Sub

Private Sub cGenColors_Click()
Dim RR As Integer, gG As Integer, bB As Integer
RR = Rnd() * 222
gG = Rnd() * 222
bB = Rnd() * 222

cClr(1).BackColor = RGB(RR, gG, bB)
cClr(2).BackColor = RGB(gG + 50, bB + 50, RR + 50)

If (RR + gG + bB) / 3 < 128 Then
  RR = RR + 100
  gG = gG + 100
  bB = bB + 100
Else
  RR = RR - 100: If RR < 0 Then RR = 0
  gG = gG - 100: If gG < 0 Then gG = 0
  bB = bB - 100: If bB < 0 Then bB = 0
End If
cClr(0).BackColor = RGB(RR, gG, bB)

sDown.Value = Rnd() * 100

Call GetRGB
Call DrawPreviev
cApply.Enabled = True
End Sub

Private Sub cGenColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(7))
End Sub

Private Sub cGenLng_Click()
Call mLng.GenLng
End Sub

Private Sub cGenLng_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(Me, 2, MapMsg(54))
End Sub

Private Sub chAutoStart_Click()
cApply.Enabled = True
End Sub

Private Sub chAutoStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(Me, 2, MapMsg(8))
End Sub

Private Sub chDrawHK_Click()
cApply.Enabled = True
End Sub

Private Sub chDrawHK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(Me, 2, MapMsg(9))
End Sub

Private Sub chFntB_Click()
Call GetPrevievText
End Sub

Private Sub chFntI_Click()
Call GetPrevievText
End Sub

Private Sub chFormFullHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(52))
End Sub

Private Sub chLng_Click()
Dim I As Integer
If chLng.Value = 0 Then
  cbLng.Enabled = False
Else
  If cbLng.ListCount > 0 Then
    cbLng.Enabled = True
    cbLng.ListIndex = 0
    If LangFile <> "" Then
      For I = 0 To cbLng.ListCount - 1 Step 1
        If cbLng.List(I) = LangFile Then cbLng.ListIndex = I
      Next I
    End If
  Else
    Call fMsg.GetMsg(fEdt, 0, MapMsg(10))
  End If
End If
cApply.Enabled = True
End Sub

Private Sub chNotAutoFocus_Click()
cApply.Enabled = True
End Sub

Private Sub chNotClearMem_Click()
cApply.Enabled = True
End Sub

Private Sub chNotClearMem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(56))
End Sub

Private Sub chPKM_Click()
cApply.Enabled = True
End Sub

Private Sub chPKM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(11))
End Sub

Private Sub chScreenBottom_Click()
cApply.Enabled = True
End Sub

Private Sub chScreenBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(Me, 2, MapMsg(12))
End Sub

Private Sub chShowInTray_Click()
cApply.Enabled = True
End Sub

Private Sub chShowInTray_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(13))
End Sub

Private Sub chSKM_Click()
cApply.Enabled = True
End Sub

Private Sub chSKM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(14))
End Sub

Private Sub chSlB_Click()
cClr(2).Enabled = chSlB.Value
Call DrawPreviev
cApply.Enabled = True
End Sub

Private Sub chTimeNotShow_Click()
cApply.Enabled = True
End Sub

Private Sub chTimeNotShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(15))
End Sub

Private Sub cHtK_Click()
fKey.Show 1, Me
tHtM.Tag = RetMod: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = RetKey: tHtK.Text = MapKN(Val(tHtK.Tag))
cApply.Enabled = True
End Sub

Private Sub cHtK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(16))
End Sub

Private Sub cHtKDel_Click()
tHtM.Tag = 0: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = 0: tHtK.Text = MapKN(Val(tHtK.Tag))
cApply.Enabled = True
End Sub

Private Sub cHtKDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(17))
End Sub

Private Sub cOK_Click()
LoadOK = 0
Call SaveSettings
Call mPrg.SetCur(-2, 0)
Unload Me
End Sub

Private Sub cStt_Click(Index As Integer)
Dim Pth As String
Dim SS As String
If Index = 0 Then
  mDlg.Filter = "All files (*.*)|*.*|"
  mDlg.SaveDialogTitle = MapOth(24)
  Pth = mDlg.ShowSave(tStt.Text)
  Pth = Replace(Pth, vbNullChar, "")
  If Pth <> "" Then
    FileCopy SttPath, Pth
    tStt.Text = Pth
  End If
End If
If Index = 1 Then
  mDlg.Filter = "All files (*.*)|*.*|"
  mDlg.OpenDialogTitle = MapOth(25)
  Pth = mDlg.ShowOpen(tStt.Text)
  Pth = Replace(Pth, vbNullChar, "")
  If Pth <> "" Then
    tStt.Text = Pth
  End If
End If
End Sub

Private Sub oPag_Click(Index As Integer)
frPag(0).Visible = False
frPag(1).Visible = False
frPag(2).Visible = False
frPag(3).Visible = False
frPag(Index).Visible = True
End Sub

Private Sub oStt_Click(Index As Integer)
If oStt(1).Value = True Then
  tStt.BackColor = vbWindowBackground
  tStt.Locked = False
Else
  tStt.BackColor = vbButtonFace
  tStt.Locked = True
End If
If oStt(0).Value = True Then tStt.Text = App.Path & "\stt.ini"
If oStt(2).Value = True Then tStt.Text = GetDir("%DOCUMENTS%") & "\LightBar\stt.ini"
If oStt(3).Value = True Then tStt.Text = GetDir("%APPLICATIONDATA%") & "\LightBar\stt.ini"
cApply.Enabled = True
End Sub

Private Sub pPrv_Click()
Call DrawPreviev
End Sub

Private Sub pPrv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call fMsg.GetMsg(fEdt, 2, MapMsg(18))
End Sub

Private Sub sDown_Change()
lDown.Caption = sDown.Value
Call DrawPreviev
cApply.Enabled = True
End Sub

Private Sub sDown_Scroll()
Call sDown_Change
End Sub

Private Sub sPol_Change()
If LoadOK = 1 Then fPrg.Left = sPol.Value * 15
lPol.Caption = sPol.Value
cApply.Enabled = True
End Sub

Private Sub sPol_Scroll()
Call sPol_Change
End Sub

Private Sub sPolV_Change()
If LoadOK = 1 Then fPrg.Top = sPolV.Value * 15
lPolV.Caption = sPolV.Value
If sPolV.Value = 0 And FormPos > 0 Then fPrg.Top = FB
cApply.Enabled = True
End Sub

Private Sub sPolV_Scroll()
Call sPolV_Change
End Sub

Private Sub sTrans_Change()
If LoadOK = 1 Then TransForm = sTrans.Value
Call SetTransparent(fPrg.hwnd, CByte(TransForm), 1)
lTrans.Caption = sTrans.Value
cApply.Enabled = True
End Sub

Private Sub sTrans_Scroll()
Call sTrans_Change
End Sub

Private Sub cScr_Click(Index As Integer)
If Index = 0 Then tCol.Text = Val(tCol.Text) - 1
If Index = 1 Then tCol.Text = Val(tCol.Text) + 1
If Index = 2 Then tRow.Text = Val(tRow.Text) - 1
If Index = 3 Then tRow.Text = Val(tRow.Text) + 1
If Index = 4 Then tZdr.Text = Val(tZdr.Text) - 100
If Index = 5 Then tZdr.Text = Val(tZdr.Text) + 100
If Index = 6 Then tAnm.Text = Val(tAnm.Text) - 10
If Index = 7 Then tAnm.Text = Val(tAnm.Text) + 10
If Index = 8 Then tIcoW.Text = Val(tIcoW.Text) - 1
If Index = 9 Then tIcoW.Text = Val(tIcoW.Text) + 1
If Index = 10 Then tIcoH.Text = Val(tIcoH.Text) - 1
If Index = 11 Then tIcoH.Text = Val(tIcoH.Text) + 1
If Index = 12 Then tBttS.Text = Val(tBttS.Text) - 1
If Index = 13 Then tBttS.Text = Val(tBttS.Text) + 1
If Index = 14 Then tIcoS.Text = Val(tIcoS.Text) - 1
If Index = 15 Then tIcoS.Text = Val(tIcoS.Text) + 1
If Index = 16 Then tMBttW.Text = Val(tMBttW.Text) - 1
If Index = 17 Then tMBttW.Text = Val(tMBttW.Text) + 1
If Index = 18 Then tMBttH.Text = Val(tMBttH.Text) - 1
If Index = 19 Then tMBttH.Text = Val(tMBttH.Text) + 1
If Index = 20 Then tFnt.Text = Val(tFnt.Text) - 1
If Index = 21 Then tFnt.Text = Val(tFnt.Text) + 1
If Index = 22 Then tOts.Text = Val(tOts.Text) - 1
If Index = 23 Then tOts.Text = Val(tOts.Text) + 1

If Val(tCol.Text) < 10 Then tCol.Text = 90
If Val(tCol.Text) > 90 Then tCol.Text = 10
If Val(tRow.Text) < 1 Then tRow.Text = 20
If Val(tRow.Text) > 20 Then tRow.Text = 1
If Val(tZdr.Text) < 0 Then tZdr.Text = tZdr.Text + 5000
If Val(tZdr.Text) > 4999 Then tZdr.Text = tZdr.Text - 5000
If Val(tAnm.Text) < 0 Then tAnm.Text = tAnm.Text + 100
If Val(tAnm.Text) > 99 Then tAnm.Text = tAnm.Text - 100
If Val(tIcoW.Text) < 4 Then tIcoW.Text = 64
If Val(tIcoW.Text) > 64 Then tIcoW.Text = 4
If Val(tIcoH.Text) < 4 Then tIcoH.Text = 64
If Val(tIcoH.Text) > 64 Then tIcoH.Text = 4
If Val(tBttS.Text) < 0 Then tBttS.Text = 50
If Val(tBttS.Text) > 50 Then tBttS.Text = 0
If Val(tIcoS.Text) < 0 Then tIcoS.Text = 20
If Val(tIcoS.Text) > 20 Then tIcoS.Text = 0
If Val(tMBttW.Text) < 9 Then tMBttW.Text = 34
If Val(tMBttW.Text) > 34 Then tMBttW.Text = 9
If Val(tMBttH.Text) < 11 Then tMBttH.Text = 34
If Val(tMBttH.Text) > 34 Then tMBttH.Text = 11
If Val(tFnt.Text) < 2 Then tFnt.Text = 72
If Val(tFnt.Text) > 72 Then tFnt.Text = 2
If Val(tOts.Text) < -50 Then tOts.Text = 50
If Val(tOts.Text) > 50 Then tOts.Text = -50
End Sub

Private Sub Form_Activate()
Call mPrg.SetCur(cCancel.hwnd)
LoadOK = 1
End Sub

Private Sub Form_Load()
Dim I As Integer, II As Integer
Dim SS As String
Dim FF As Long

Call mLng.LoadLang(LangFile, "stt")

If FormNotTop = 0 Then SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
tCol.Text = bttCol
tRow.Text = bttRow
tIcoW.Text = icoW
tIcoH.Text = icoH
tBttS.Text = bttS
tIcoS.Text = icoS
If fPrg.tFFH.Enabled = False Then chFormFullHide.Value = 0 Else chFormFullHide.Value = 1
tZdr.Text = fPrg.tZdr.Interval
tAnm.Text = fPrg.tShow.Tag
tMBttW.Text = MBttW
tMBttH.Text = MBttH
sPol.Max = Screen.Width / 15 - 15: sPol.LargeChange = sPol.Max / 5
sPol.Value = fPrg.Left / 15: sPol.Tag = sPol.Value
sPolV.Max = (Screen.Height - fPrg.Height) / 15: sPolV.LargeChange = sPolV.Max / 5
sPolV.Value = FormTop / 15: sPolV.Tag = sPolV.Value
sTrans.Value = TransForm: sTrans.Tag = sTrans.Value

cClr(0).BackColor = ClrFnt
cClr(1).BackColor = RGB(MapC(1), MapC(2), MapC(3))
cClr(10).BackColor = ClrFrm
sDown.Value = MapC(4)
chSlB.Value = MapC(5)
cClr(2).BackColor = RGB(MapC(6), MapC(7), MapC(8))
Call GetRGB
Call DrawPreviev

'For I = 0 To 10 Step 1
'  cClr(I).BackColor = MapC(I)
'Next I

tHtM.Tag = HotMod: tHtM.Text = GetTextMod(Val(tHtM.Tag))
tHtK.Tag = HotKey: tHtK.Text = MapKN(Val(tHtK.Tag))

If Dir(SpecialFolder(7) + "\lightbar.lnk") = "" Then chAutoStart.Value = 0 Else chAutoStart.Value = 1
If FormPos = 0 Then chScreenBottom.Value = 0 Else chScreenBottom.Value = 1
If TimeNotShow = 0 Then chTimeNotShow.Value = 0 Else chTimeNotShow.Value = 1
If ShowInTray = 0 Then chShowInTray.Value = 0 Else chShowInTray.Value = 1
If NotAutoFocus = 0 Then chNotAutoFocus.Value = 0 Else chNotAutoFocus.Value = 1
If NotClearMem = 0 Then chNotClearMem.Value = 0 Else chNotClearMem.Value = 1

chPKM.Value = BttToShow

'shrifty
cbFnt.Clear
For I = 1 To Screen.FontCount - 1
  cbFnt.AddItem Screen.Fonts(I)
Next I
For I = 0 To cbFnt.ListCount - 1
  If cbFnt.List(I) = FntName Then cbFnt.ListIndex = I: Exit For
Next I
tFnt.Text = FntSize
chFntB.Value = FntBold
chFntI.Value = FntItalic
tOts.Text = FntTop

'yazyki
cbLng.Clear
SS = Dir(App.Path & "\*.lng")
Do
  If SS = "" Then Exit Do
  cbLng.AddItem SS
  SS = Dir
Loop

If LangFile <> "" Then chLng.Value = 1

'fajl nastroek
If Dir(App.Path & "\sttpath.ini") = "" Then
  oStt(0).Value = True
Else
  FF = FreeFile
  Open App.Path & "\sttpath.ini" For Input As #FF
    If EOF(FF) = False Then Input #FF, SS
  Close #FF
  
  If SS <> "" Then
    If SS = "1" Then
      oStt(2).Value = True
      tStt.Text = GetDir("%DOCUMENTS%") & "\lightbar\stt.ini"
    ElseIf SS = "2" Then
      oStt(3).Value = True
      tStt.Text = GetDir("%APPLICATIONDATA%") & "\lightbar\stt.ini"
    Else
      oStt(1).Value = True
      tStt.Text = SS
    End If
  End If
End If
tStt.Tag = tStt.Text

mDlg.hwndOwner = Me.hwnd

'combo1.AddItem
cApply.Enabled = False
End Sub

Private Sub cClr_Click(Index As Integer)
Dim Res As Long
mDlg.hwndOwner = Me.hwnd
Res = mDlg.ShowColor
If Res >= 0 Then cClr(Index).BackColor = Res
Call GetRGB
Call DrawPreviev
cApply.Enabled = True
End Sub

Private Sub tAnm_Change()
cApply.Enabled = True
End Sub

Private Sub tAnm_GotFocus()
tAnm.SelStart = 0
tAnm.SelLength = Len(tAnm.Text)
End Sub

Private Sub tBttS_Change()
cApply.Enabled = True
End Sub

Private Sub tBttS_GotFocus()
tBttS.SelStart = 0
tBttS.SelLength = Len(tBttS.Text)
End Sub

Private Sub tCol_Change()
cApply.Enabled = True
End Sub

Private Sub tFnt_Change()
Call GetPrevievText
End Sub

'################################################'
'### SUBS AND FUNCTIONS #########################'
'################################################'

Private Sub SaveSettings()
Dim I As Integer
Dim SS As String
Dim wDrv As Byte, wDrvNB As Byte 'nado li pererisovyvat' jekran
Dim FF As Long

If Val(tCol.Text) < 10 Or Val(tCol.Text) > 90 Then Call fMsg.GetMsg(fStt, 0, MapMsg(19)): Exit Sub
If Val(tRow.Text) < 1 Or Val(tRow.Text) > 20 Then Call fMsg.GetMsg(fStt, 0, MapMsg(20)): Exit Sub
If Val(tZdr.Text) < 0 Or Val(tZdr.Text) > 5000 Then Call fMsg.GetMsg(fStt, 0, MapMsg(21)): Exit Sub
If Val(tIcoW.Text) < 4 Or Val(tIcoW.Text) > 64 Then Call fMsg.GetMsg(fStt, 0, MapMsg(22)): Exit Sub
If Val(tIcoH.Text) < 4 Or Val(tIcoH.Text) > 64 Then Call fMsg.GetMsg(fStt, 0, MapMsg(23)): Exit Sub
If Val(tBttS.Text) < 0 Or Val(tBttS.Text) > 50 Then Call fMsg.GetMsg(fStt, 0, MapMsg(24)): Exit Sub
If Val(tIcoS.Text) < 0 Or Val(tIcoS.Text) > 20 Then Call fMsg.GetMsg(fStt, 0, MapMsg(25)): Exit Sub
If Val(tMBttW.Text) < 9 Or Val(tMBttW.Text) > 34 Then Call fMsg.GetMsg(fStt, 0, MapMsg(26)): Exit Sub
If Val(tMBttH.Text) < 11 Or Val(tMBttH.Text) > 34 Then Call fMsg.GetMsg(fStt, 0, MapMsg(27)): Exit Sub
If Val(tFnt.Text) < 2 Or Val(tFnt.Text) > 72 Then Call fMsg.GetMsg(fStt, 0, MapMsg(28)): Exit Sub
If Val(tOts.Text) < -50 Or Val(tOts.Text) > 50 Then Call fMsg.GetMsg(fStt, 0, MapMsg(55)): Exit Sub

If bttCol <> Val(tCol.Text) Then bttCol = Val(tCol.Text): wDrv = 1
If bttRow <> Val(tRow.Text) Then bttRow = Val(tRow.Text): wDrv = 1
If icoW <> Val(tIcoW.Text) Then icoW = Val(tIcoW.Text): wDrv = 1
If icoH <> Val(tIcoH.Text) Then icoH = Val(tIcoH.Text): wDrv = 1
If bttS <> Val(tBttS.Text) Then bttS = Val(tBttS.Text): wDrv = 1
If icoS <> Val(tIcoS.Text) Then icoS = Val(tIcoS.Text): wDrv = 1
If MBttW <> Val(tMBttW.Text) Then MBttW = Val(tMBttW.Text): wDrv = 1
If MBttH <> Val(tMBttH.Text) Then MBttH = Val(tMBttH.Text): wDrv = 1

FormLeft = sPol.Value * 15: fPrg.Left = FormLeft
FormTop = sPolV.Value * 15: fPrg.Top = FormTop
TransForm = sTrans.Value
fPrg.tZdr.Interval = Val(tZdr.Text)
If fPrg.tZdr.Interval > 0 Then fPrg.tFFH.Enabled = chFormFullHide.Value Else fPrg.tFFH.Enabled = 0
fPrg.tShow.Tag = Val(tAnm.Text)

MapC(0) = 0

If MapC(1) <> cR Then MapC(1) = cR: wDrv = 1
If MapC(2) <> cG Then MapC(2) = cG: wDrv = 1
If MapC(3) <> cB Then MapC(3) = cB: wDrv = 1
If MapC(4) <> sDown.Value Then MapC(4) = sDown.Value: wDrv = 1

MapC(5) = chSlB.Value
MapC(6) = cRA
MapC(7) = cGA
MapC(8) = cBA

If ClrFnt <> cClr(0).BackColor Then ClrFnt = cClr(0).BackColor: wDrvNB = 1
If ClrFrm <> cClr(10).BackColor Then ClrFrm = cClr(10).BackColor: wDrvNB = 1

If HotMod <> tHtM.Tag Or HotKey <> tHtK.Tag Then
  HotMod = tHtM.Tag
  HotKey = tHtK.Tag
End If

SS = App.Path & "\" & App.EXEName
If Right$(SS, 4) <> ".exe" Then SS = SS & ".exe"
If chAutoStart.Value = 1 Then
  Call AddToAutorun
Else
  If Dir(SpecialFolder(7) + "\lightbar.lnk") <> "" Then Kill SpecialFolder(7) + "\lightbar.lnk"
End If

If chDrawHK.Value <> DrawHotKey Then DrawHotKey = chDrawHK.Value: wDrv = 1

If FormPos <> chScreenBottom.Value Then
  FormPos = chScreenBottom.Value
  If FormTop = 0 Then If FormPos = 0 Then fPrg.Top = 0 Else fPrg.Top = FB
  wDrv = 1
End If

If chTimeNotShow.Value <> TimeNotShow Then TimeNotShow = chTimeNotShow.Value: wDrvNB = 1
If chShowInTray.Value <> ShowInTray Then
  ShowInTray = chShowInTray.Value
  Call TrayMgr(ShowInTray)
End If
If chNotAutoFocus.Value <> NotAutoFocus Then NotAutoFocus = chNotAutoFocus.Value

If chNotClearMem.Value <> NotClearMem Then NotClearMem = chNotClearMem.Value

BttToShow = chPKM.Value

FntName = cbFnt.Text: fPrg.pKnt.FontName = FntName
FntSize = tFnt.Text: fPrg.pKnt.FontSize = FntSize
FntBold = chFntB.Value: fPrg.pKnt.FontBold = FntBold
FntItalic = chFntI.Value: fPrg.pKnt.FontItalic = FntItalic
fPrg.pKntTime.FontName = FntName
fPrg.pKntTime.FontSize = FntSize
fPrg.pKntTime.FontBold = FntBold
fPrg.pKntTime.FontItalic = FntItalic
FntTop = tOts.Text
Call GetPrevievText

If chLng.Value = 0 Then
  LangFile = ""
Else
  LangFile = cbLng.Text
  Call mLng.LoadLang(LangFile)
End If

If tStt.Tag <> tStt.Text Then
  If oStt(0).Value = True Then
    If Dir(App.Path & "\sttpath.ini") <> "" Then Kill App.Path & "\sttpath.ini"
  Else
    FF = FreeFile
    Open App.Path & "\sttpath.ini" For Output As #FF
      If oStt(1).Value = True Then Print #FF, tStt.Text
      If oStt(2).Value = True Then Print #FF, "1"
      If oStt(3).Value = True Then Print #FF, "2"
    Close #FF
  End If
  SttPath = tStt.Text
  If oStt(1).Value = True And Dir(SttPath) = "" Then
    Call fMsg.GetMsg(fPrg, 1, MapMsg(51) & vbCrLf & vbCrLf & "(" & SttPath & ")")
    SttPath = App.Path & "\stt.ini"
  End If
  Call mPrg.LoadStt
  'wDrvNB = 1
End If

If wDrv = 1 Then
  Call DrawForm
ElseIf wDrvNB = 1 Then
  Call DrawForm(0)
End If

If FormPos > 0 Then
  FB = Screen.Height - fPrg.Height
  If FormTop = 0 Then fPrg.Top = FB
End If

Call SetTransparent(fPrg.hwnd, CByte(TransForm), 1)
Call SaveStt

End Sub

Private Sub GetPrevievText()
Dim I As Integer

fPrg.pKnt.FontName = cbFnt.Text
I = Val(tFnt.Text)
If I < 2 Then I = 2
If I > 72 Then I = 72
fPrg.pKnt.FontSize = I
fPrg.pKnt.FontBold = chFntB.Value
fPrg.pKnt.FontItalic = chFntI.Value
fPrg.pKntTime.FontName = FntName
fPrg.pKntTime.FontSize = FntSize
fPrg.pKntTime.FontBold = FntBold
fPrg.pKntTime.FontItalic = FntItalic

fPrg.pKnt.Cls: fPrg.pKnt.CurrentX = 0: fPrg.pKnt.CurrentY = Val(tOts.Text): fPrg.pKnt.Print cbFnt.Text
fPrg.PaintPicture fPrg.pKnt.Image, fPrg.pKnt.Left, fPrg.pKnt.Top

cApply.Enabled = True
End Sub

Private Sub LoadLngInf()
Call mLng.LoadLang(cbLng.Text, "gen")
tLngInf(0).Text = MapGen(0)
tLngInf(1).Text = MapGen(1)
tLngInf(2).Text = MapGen(2)
End Sub

Private Sub GetRGB()
Dim Color As Long
Color = cClr(1).BackColor
cR = Color Mod 256: Color = Color \ 256
cG = Color Mod 256: Color = Color \ 256
cB = Color
Color = cClr(2).BackColor
cRA = Color Mod 256: Color = Color \ 256
cGA = Color Mod 256: Color = Color \ 256
cBA = Color
End Sub

Private Sub DrawPreviev()
Dim I As Integer, II As Integer
Dim Btt As RECT
Dim ab(1) As Byte, dB(1) As Byte

ab(0) = Rnd() * 2: ab(1) = Rnd() * 2
1000
dB(0) = Rnd() * 2: dB(1) = Rnd() * 2
If dB(0) = ab(0) And dB(1) = ab(1) Then GoTo 1000

pPrv.ForeColor = cClr(0).BackColor
pPrv.BackColor = cClr(1).BackColor
pPrv.Cls

Btt.Left = 0: Btt.Top = 0: Btt.Right = 66: Btt.Bottom = 81: Call DrwBrdr(Btt, 0)
Btt.Left = 2: Btt.Top = 2: Btt.Right = 62: Btt.Bottom = 62: Call DrwBrdr(Btt, 1)

Btt.Left = 2: Btt.Top = 65: Btt.Right = 23: Btt.Bottom = 14: Call DrwBrdr(Btt, 1)
Btt.Left = 26: Btt.Top = 65: Btt.Right = 38: Btt.Bottom = 14: Call DrwBrdr(Btt, 1)

For I = 0 To 2 Step 1
  For II = 0 To 2 Step 1
    Btt.Left = I * 19 + 5: Btt.Top = II * 19 + 5: Btt.Right = 18: Btt.Bottom = 18
    If dB(0) = I And dB(1) = II Then
      Call DrwBrdr(Btt, 4)
    Else
      Call DrwBrdr(Btt, 2)
    End If
  Next II
Next I

'knopka nastroek
Btt.Left = 4: Btt.Top = 81 - 13: Btt.Right = 9: Btt.Bottom = 9: Call DrwBrdr(Btt, 2)
pPrv.Line (6, 81 - 10)-(7, 81 - 7), cClr(0).BackColor, B
pPrv.Line (8, 81 - 9)-(11, 81 - 9), cClr(0).BackColor
'knopka o programme
Btt.Left = 14: Btt.Top = 81 - 13: Btt.Right = 9: Btt.Bottom = 9: Call DrwBrdr(Btt, 2)
pPrv.Line (17, 81 - 11)-(20, 81 - 11), cClr(0).BackColor
pPrv.Line (17, 81 - 7)-(20, 81 - 7), cClr(0).BackColor
pPrv.Line (18, 81 - 11)-(18, 81 - 7), cClr(0).BackColor

pPrv.CurrentX = 28: pPrv.CurrentY = 65: pPrv.Print Format(Now, "hh:nn")

If chSlB.Value = 0 Then
  pActiv.BackColor = RGB(200, 200, 200)
  pActiv.PaintPicture pPrv.Image, 0, 0, , , , , , , vbSrcAnd
'Else
'  pActiv.PaintPicture pPrv.Image, 0, 0
End If

Btt.Left = ab(0) * 19 + 5: Btt.Top = ab(1) * 19 + 5: Btt.Right = 18: Btt.Bottom = 18
Call DrwBrdr(Btt, 6)

End Sub

Private Sub DrwBrdr(ByRef wBtt As RECT, ByRef wState As Byte)
Dim C1 As Long, C2 As Long

If wState = 0 Then C1 = GenColor(sDown.Value, cR, cG, cB): C2 = GenColor(-sDown.Value, cR, cG, cB)
If wState = 1 Then C2 = GenColor(sDown.Value, cR, cG, cB): C1 = GenColor(-sDown.Value, cR, cG, cB)
If wState = 2 Then C1 = GenColor(sDown.Value, cR, cG, cB): C2 = GenColor(-sDown.Value, cR, cG, cB)
If wState = 3 Then C2 = GenColor(sDown.Value, cR, cG, cB): C1 = GenColor(-sDown.Value, cR, cG, cB)
If wState = 4 Then C1 = GenColor(-sDown.Value, cR, cG, cB): C2 = GenColor(sDown.Value, cR, cG, cB)
If wState = 5 Then C2 = GenColor(-sDown.Value, cR, cG, cB): C1 = GenColor(sDown.Value, cR, cG, cB)

If chSlB.Value = 1 Then
  If wState = 6 Then C1 = GenColor(sDown.Value, cRA, cGA, cBA): C2 = GenColor(-sDown.Value, cRA, cGA, cBA)
  If wState = 7 Then C2 = GenColor(sDown.Value, cRA, cGA, cBA): C1 = GenColor(-sDown.Value, cRA, cGA, cBA)
End If

wBtt.Right = wBtt.Right - 1
wBtt.Bottom = wBtt.Bottom - 1

1000
If wState = 6 Or wState = 7 Then
  If wBtt.Left > -1 Then
    If chSlB.Value = 0 Then
      pPrv.PaintPicture pActiv.Image, wBtt.Left + 1, wBtt.Top + 1, wBtt.Right - 1, wBtt.Bottom - 1, wBtt.Left + 1, wBtt.Top + 1, wBtt.Right - 1, wBtt.Bottom - 1
    Else
      wState = wState - 4
      GoTo 1000
    End If
  End If
Else
  pPrv.Line (wBtt.Left, wBtt.Top)-(wBtt.Left, wBtt.Top + wBtt.Bottom), C1
  pPrv.Line (wBtt.Left, wBtt.Top)-(wBtt.Left + wBtt.Right + 1, wBtt.Top), C1
  pPrv.Line (wBtt.Left + wBtt.Right, wBtt.Top + wBtt.Bottom)-(wBtt.Left + wBtt.Right, wBtt.Top), C2
  pPrv.Line (wBtt.Left + wBtt.Right, wBtt.Top + wBtt.Bottom)-(wBtt.Left - 1, wBtt.Top + wBtt.Bottom), C2
End If

End Sub




























Private Sub tCol_GotFocus()
tCol.SelStart = 0
tCol.SelLength = Len(tCol.Text)
End Sub

Private Sub tFnt_GotFocus()
tFnt.SelStart = 0
tFnt.SelLength = Len(tFnt.Text)
End Sub

Private Sub tIcoH_Change()
cApply.Enabled = True
End Sub

Private Sub tIcoH_GotFocus()
tIcoH.SelStart = 0
tIcoH.SelLength = Len(tIcoH.Text)
End Sub

Private Sub tIcoS_Change()
cApply.Enabled = True
End Sub

Private Sub tIcoS_GotFocus()
tIcoS.SelStart = 0
tIcoS.SelLength = Len(tIcoS.Text)
End Sub

Private Sub tIcoW_Change()
cApply.Enabled = True
End Sub

Private Sub tIcoW_GotFocus()
tIcoW.SelStart = 0
tIcoW.SelLength = Len(tIcoW.Text)
End Sub

Private Sub tMBttH_Change()
cApply.Enabled = True
End Sub

Private Sub tMBttH_GotFocus()
tMBttH.SelStart = 0
tMBttH.SelLength = Len(tMBttH.Text)
End Sub

Private Sub tMBttW_Change()
cApply.Enabled = True
End Sub

Private Sub tMBttW_GotFocus()
tMBttW.SelStart = 0
tMBttW.SelLength = Len(tMBttW.Text)
End Sub

Private Sub tOts_Change()
Call GetPrevievText
End Sub

Private Sub tOts_GotFocus()
tOts.SelStart = 0
tOts.SelLength = Len(tOts.Text)
End Sub

Private Sub tRow_Change()
cApply.Enabled = True
End Sub

Private Sub tRow_GotFocus()
tRow.SelStart = 0
tRow.SelLength = Len(tRow.Text)
End Sub

Private Sub tStt_Change()
cApply.Enabled = True
End Sub

Private Sub tZdr_Change()
cApply.Enabled = True
If Val(tZdr.Text) > 0 Then
  chFormFullHide.Enabled = True
Else
  chFormFullHide.Enabled = False
  chFormFullHide.Value = 0
End If
End Sub

Private Sub tZdr_GotFocus()
tZdr.SelStart = 0
tZdr.SelLength = Len(tZdr.Text)
End Sub
