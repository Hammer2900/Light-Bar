Attribute VB_Name = "mPrg"
'################################################'
'# Programm:                           LightBar #'
'# Part:                            Main Module #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             06.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Public Type POINTAPI
  X As Long
  Y As Long
End Type
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type SHITEMID
    cB As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
  mkid As SHITEMID
End Type
Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'for upper window
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'for transparent window
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'for tray
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'for icons
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Private Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Type typB
  'dlya shellexecute
  wOpr As Byte   '+ operaciya
  wFil As String '+ put' k failu           \
  wPrm As String '+ parametry               } ili vnutrennyaya komanda i eyo parametry
  wDir As String '+ katalog po umolchaniyu /
  wShw As Byte   '+ rezjim otkrytiya
  '
  wCap As String '+ caption
  wHtM As Byte   '+ hot mod
  wHtK As Byte   '+ hot key
  wHtN As Byte   '+ blokirovschik goryachih klavish
  wIFl As String '+ fajl ikonki
  wINm As Long   '+ nomer ikonki
  wKmm As String '- kommentarij
  
  wClr(4) As Byte '0-tip (0-vypuklyj 1-ploskij) 1-r 2-g 3-b 4-glubina
  
End Type
Public MapB(2200) As typB

Public SttPath As String

Public MapC(8) As Integer '0-tip (0-vypuklyj 1-ploskij) 1-r 2-g 3-b 4-glubina 5-stil' vydeleniya 678-rgb vydeleniya
Public ClrFnt As Long 'font color
Public ClrFrm As Long 'minimize form color
Public bttCol As Integer, bttRow As Integer 'stolbcy,stroki knopok
Public icoW As Integer, icoH As Integer     'shirina,vysota knopok
Public bttS As Integer, icoS As Integer     'otstup mejdu knopkami,otstup iconok ot kraya
Public FormLeft As Integer                  'poziciya formy sleva
Public FormTop As Integer                   'poziciya formy sverhu
Public TransForm As Long                    'prozrachnost' formy
Public FormNotHide As Byte   'skrytie formy
Public FormNotTop As Byte    'poverh vseh okon formy
Public FormNotHotKey As Byte 'lovlya goryachih klavish
Public DrawHotKey As Byte    'otrisovka goryachej klavishy na knopke
Public MBttW As Byte         'razmer knopok menyu po gorizontali
Public MBttH As Byte         'razmer knopok menyu po vertikali
Public TimeNotShow As Byte   'pokaz vremeni v glavnom okne
Public ShowInTray As Byte    'pri zakrytii ne svorachivat' v trey
Public BttToShow As Byte     'ot kakoj knopki myshi otkryvat'sya (0-l 1-r 10-s&l 11-s&r)
Public LangFile As String    'dlya imeni yazykovogo fajla
Public NotAutoFocus As Byte  'avtomaticheskaya fokusirovka kursora na elementah
Public NotClearMem As Byte   'ne ochicshat' pamyat'

Public FntName As String 'dlya nazvaniya shrifta
Public FntSize As Byte   'dlya razmera
Public FntBold As Byte   'dlya zhirnosti
Public FntItalic As Byte 'dlya naklonnosti
Public FntTop As Integer 'dlya otstupa s verhu

Public FormPos As Byte 'poziciya formy na jekrane // 0-sverhu // 1-snizu //
Public FB As Integer   'polojenie fprg.top pri formpos=1

Public Lck As Byte 'zamok na skrytie glavnoj formy
Public PoluLck As Byte 'poluzamok :) jeto kogda glavnaya forma nedolzhna skryvat'sya, NO ona aktivna
                                                     '(takoe byvaet pri otkrytii formy s klaviatury)
Public EdtB As Integer

Public HotMod As Byte, HotKey As Byte
Public HotId As Long

Public ModK As Byte         'polojenie klavish modifikatorov
Public MapKS(255) As Byte   'massiv sostoyaniya klavish
Public MapKN(255) As String 'massiv nazvanij klavish
'Public MapKId(2200) As Long

Public RetMod As Byte, RetKey As Byte
Public RetMsg As Byte

Public txtW As Integer
Public frmW As Long, frmH As Long 'razmery formy (chtob igry eyo ne umen'shali)

Private NID As NOTIFYICONDATA 'dlya ikonki v tree

Private FF As Long

Public Sub StrPrg()
fPrg.Caption = App.Title & " v." & App.Major & "." & App.Minor & " (" & App.Revision & ")"
Dim I As Integer, II As Integer, Btt As Integer

Call GetSttPath

Call LoadColors

If Dir(SttPath) = "" Then Call CreateDefaultIni
Call LoadStt
Call LoadMapKN

Call mLng.LoadLang(LangFile)
Unload fAbt
Unload fCmd
Unload fEdt
Unload fKey
Unload fMsg
Unload fStt

Call SetTransparent(fPrg.hwnd, CByte(TransForm), 1)
If FormNotTop = 0 Then
  SetWindowPos fPrg.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
Else
  SetWindowPos fPrg.hwnd, 1, 0, 0, 0, 0, &H10 Or &H1 Or &H2
End If

fPrg.pKnt.FontName = FntName
fPrg.pKnt.FontSize = FntSize
fPrg.pKnt.FontBold = FntBold
fPrg.pKnt.FontItalic = FntItalic
fPrg.pKntTime.FontName = FntName
fPrg.pKntTime.FontSize = FntSize
fPrg.pKntTime.FontBold = FntBold
fPrg.pKntTime.FontItalic = FntItalic

Call DrawForm

Call TrayMgr(1)

fPrg.Left = FormLeft

If Right$(App.Path, 5) <> "\!src" Then hKbdLL = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)

End Sub

Public Sub ExtPrg()
Dim OldFormNotHide As Byte
Dim OldFormTop As Integer
Dim I As Integer

OldFormNotHide = FormNotHide
OldFormTop = FormTop

Lck = 0
FormNotHide = 0
FormTop = 0
fPrg.tPpp.Enabled = False
fPrg.tHide.Enabled = True
Do
  If fPrg.tHide.Enabled = False Then
    If FormPos = 0 Then fPrg.Top = -fPrg.Height - 15 Else fPrg.Top = Screen.Height - 15
    Exit Do
  End If
  DoEvents
Loop

If hKbdLL <> 0 Then UnhookWindowsHookEx hKbdLL

FormNotHide = OldFormNotHide
FormTop = OldFormTop

Call SaveStt

Unload fAbt
Unload fCmd
Unload fEdt
Unload fKey
Unload fMsg
Unload fStt
Unload fPrg

Call TrayMgr(0)

End
End Sub

Private Sub GetSttPath()
Dim SS As String
SttPath = App.Path & "\stt.ini"
If Dir(App.Path & "\sttpath.ini") = "" Then Exit Sub 'esli fajla redirekta ne sucshestvuet, to uhodim

FF = FreeFile
Open App.Path & "\sttpath.ini" For Input As #FF
  If EOF(FF) = False Then Input #FF, SS
Close #FF

If SS = "1" Then 'moi documenty
  If Dir(GetDir("%DOCUMENTS%"), vbDirectory) <> "" Then
    If Dir(GetDir("%DOCUMENTS%") & "\lightbar", vbDirectory) = "" Then MkDir GetDir("%DOCUMENTS%") & "\LightBar"
    SttPath = GetDir("%DOCUMENTS%") & "\lightbar\stt.ini"
  End If
ElseIf SS = "2" Then 'aplication data
  If Dir(GetDir("%APPLICATIONDATA%"), vbDirectory) <> "" Then
    If Dir(GetDir("%APPLICATIONDATA%") & "\lightbar", vbDirectory) = "" Then MkDir GetDir("%APPLICATIONDATA%") & "\LightBar"
    SttPath = GetDir("%APPLICATIONDATA%") & "\lightbar\stt.ini"
  End If
Else 'polnyj put'
  If SS <> "" Then
    If Dir(SS) = "" Then
      Call fMsg.GetMsg(fPrg, 1, "Not way is found to file with settings." & vbCrLf & vbCrLf & "(" & SS & ")")
    Else
      SttPath = SS
    End If
  End If
End If

End Sub

Private Sub LoadMapKN()
Dim I As Integer
MapKN(0) = "(No)":                MapKN(10) = "Key010":      MapKN(20) = "Caps Lock": MapKN(30) = "Key030":      MapKN(40) = "Arrow Down":   MapKN(50) = "2":      MapKN(60) = "Key060": MapKN(70) = "F": MapKN(80) = "P": MapKN(90) = "Z":             MapKN(100) = "Num 4":      MapKN(110) = "Num .": MapKN(120) = "F9":   MapKN(130) = "?F19":   MapKN(140) = "Key140":      MapKN(150) = "Key150": MapKN(160) = "Left Shift":    MapKN(170) = "Key170":      MapKN(180) = "Mail":   MapKN(190) = ".":      MapKN(200) = "Key200": MapKN(210) = "Key210": MapKN(220) = "\":           MapKN(230) = "Key230": MapKN(240) = "Key240": MapKN(250) = "?PLAY"
MapKN(1) = "Mouse Left button":   MapKN(11) = "Key011":      MapKN(21) = "Key021":    MapKN(31) = "Key031":      MapKN(41) = "?SELECT":      MapKN(51) = "3":      MapKN(61) = "Key061": MapKN(71) = "G": MapKN(81) = "Q": MapKN(91) = "Left Windows":  MapKN(101) = "Num 5":      MapKN(111) = "Num /": MapKN(121) = "F10":  MapKN(131) = "?F20":   MapKN(141) = "Key141":      MapKN(151) = "Key151": MapKN(161) = "Right Shift":   MapKN(171) = "Favorites":   MapKN(181) = "Key181": MapKN(191) = "/":      MapKN(201) = "Key201": MapKN(211) = "Key211": MapKN(221) = "]":           MapKN(231) = "Key231": MapKN(241) = "Key241": MapKN(251) = "?ZOOM"
MapKN(2) = "Mouse Right button":  MapKN(12) = "?CLEAR":      MapKN(22) = "Key022":    MapKN(32) = "Space":       MapKN(42) = "?PRINT":       MapKN(52) = "4":      MapKN(62) = "Key062": MapKN(72) = "H": MapKN(82) = "R": MapKN(92) = "Right Windows": MapKN(102) = "Num 6":      MapKN(112) = "F1":    MapKN(122) = "F11":  MapKN(132) = "?F21":   MapKN(142) = "Key142":      MapKN(152) = "Key152": MapKN(162) = "Left Control":  MapKN(172) = "Browser":     MapKN(182) = "Key182": MapKN(192) = "`":      MapKN(202) = "Key202": MapKN(212) = "Key212": MapKN(222) = "'":           MapKN(232) = "Key232": MapKN(242) = "Key242": MapKN(252) = "?NONAME"
MapKN(3) = "?CANCEL":             MapKN(13) = "Enter":       MapKN(23) = "Key023":    MapKN(33) = "Pade Up":     MapKN(43) = "?EXECUTE":     MapKN(53) = "5":      MapKN(63) = "Key063": MapKN(73) = "I": MapKN(83) = "S": MapKN(93) = "Menu":          MapKN(103) = "Num 7":      MapKN(113) = "F2":    MapKN(123) = "F12":  MapKN(133) = "?F22":   MapKN(143) = "Key143":      MapKN(153) = "Key153": MapKN(163) = "Right Control": MapKN(173) = "Mute":        MapKN(183) = "Key183": MapKN(193) = "Key193": MapKN(203) = "Key203": MapKN(213) = "Key213": MapKN(223) = "Key223":      MapKN(233) = "Key233": MapKN(243) = "Key243": MapKN(253) = "?PA1"
MapKN(4) = "Mouse Middle button": MapKN(14) = "Key014":      MapKN(24) = "Key024":    MapKN(34) = "Page Down":   MapKN(44) = "Print Screen": MapKN(54) = "6":      MapKN(64) = "Key064": MapKN(74) = "J": MapKN(84) = "T": MapKN(94) = "Key094":        MapKN(104) = "Num 8":      MapKN(114) = "F3":    MapKN(124) = "?F13": MapKN(134) = "?F23":   MapKN(144) = "Num Lock":    MapKN(154) = "Key154": MapKN(164) = "Left Alt":      MapKN(174) = "Volume Down": MapKN(184) = "Key184": MapKN(194) = "Key194": MapKN(204) = "Key204": MapKN(214) = "Key214": MapKN(224) = "Key224":      MapKN(234) = "Key234": MapKN(244) = "Key244": MapKN(254) = "?OEM_CLEAR"
MapKN(5) = "Mouse X1 button":     MapKN(15) = "Key015":      MapKN(25) = "Key025":    MapKN(35) = "End":         MapKN(45) = "Insert":       MapKN(55) = "7":      MapKN(65) = "A":      MapKN(75) = "K": MapKN(85) = "U": MapKN(95) = "Key095":        MapKN(105) = "Num 9":      MapKN(115) = "F4":    MapKN(125) = "?F14": MapKN(135) = "?F24":   MapKN(145) = "Scroll Lock": MapKN(155) = "Key155": MapKN(165) = "Right Alt":     MapKN(175) = "Volume Up":   MapKN(185) = "Key185": MapKN(195) = "Key195": MapKN(205) = "Key205": MapKN(215) = "Key215": MapKN(225) = "Key225":      MapKN(235) = "Key235": MapKN(245) = "Key245": MapKN(255) = "Key255"
MapKN(6) = "Mouse X2 button":     MapKN(16) = "Shift":       MapKN(26) = "Key026":    MapKN(36) = "Home":        MapKN(46) = "Delete":       MapKN(56) = "8":      MapKN(66) = "B":      MapKN(76) = "L": MapKN(86) = "V": MapKN(96) = "Num 0":         MapKN(106) = "Num *":      MapKN(116) = "F5":    MapKN(126) = "?F15": MapKN(136) = "Key136": MapKN(146) = "Key146":      MapKN(156) = "Key156": MapKN(166) = "Back":          MapKN(176) = "Key176":      MapKN(186) = ";":      MapKN(196) = "Key196": MapKN(206) = "Key206": MapKN(216) = "Key216": MapKN(226) = "Key226":      MapKN(236) = "Key236": MapKN(246) = "?ATTN"
MapKN(7) = "Key007":              MapKN(17) = "Control":     MapKN(27) = "Escape":    MapKN(37) = "Arrow Left":  MapKN(47) = "?HELP":        MapKN(57) = "9":      MapKN(67) = "C":      MapKN(77) = "M": MapKN(87) = "W": MapKN(97) = "Num 1":         MapKN(107) = "Num +":      MapKN(117) = "F6":    MapKN(127) = "?F16": MapKN(137) = "Key137": MapKN(147) = "Key147":      MapKN(157) = "Key157": MapKN(167) = "Forward":       MapKN(177) = "Key177":      MapKN(187) = "=":      MapKN(197) = "Key197": MapKN(207) = "Key207": MapKN(217) = "Key217": MapKN(227) = "Key227":      MapKN(237) = "Key237": MapKN(247) = "?CRSEL"
MapKN(8) = "Backspace":           MapKN(18) = "Alt":         MapKN(28) = "Key028":    MapKN(38) = "Arrow Up":    MapKN(48) = "0":            MapKN(58) = "Key058": MapKN(68) = "D":      MapKN(78) = "N": MapKN(88) = "X": MapKN(98) = "Num 2":         MapKN(108) = "?SEPARATOR": MapKN(118) = "F7":    MapKN(128) = "?F17": MapKN(138) = "Key138": MapKN(148) = "Key148":      MapKN(158) = "Key158": MapKN(168) = "Key168":        MapKN(178) = "Key178":      MapKN(188) = ",":      MapKN(198) = "Key198": MapKN(208) = "Key208": MapKN(218) = "Key218": MapKN(228) = "Key228":      MapKN(238) = "Key238": MapKN(248) = "?EXSEL"
MapKN(9) = "Tab":                 MapKN(19) = "Pause Break": MapKN(29) = "Key029":    MapKN(39) = "Arrow Right": MapKN(49) = "1":            MapKN(59) = "Key059": MapKN(69) = "E":      MapKN(79) = "O": MapKN(89) = "Y": MapKN(99) = "Num 3":         MapKN(109) = "Num -":      MapKN(119) = "F8":    MapKN(129) = "?F18": MapKN(139) = "Key139": MapKN(149) = "Key149":      MapKN(159) = "Key159": MapKN(169) = "Key169":        MapKN(179) = "Key179":      MapKN(189) = "-":      MapKN(199) = "Key199": MapKN(209) = "Key209": MapKN(219) = "[":      MapKN(229) = "?PROCESSKEY": MapKN(239) = "Key239": MapKN(249) = "?EREOF"
End Sub

Public Sub DrawForm(Optional ByRef wDrwBtt As Byte = 1)
Dim I As Integer, II As Integer
Dim Btt As RECT
Dim InfT As Integer 'polozhenie informacionnyh polej s verhu
Dim tT As Integer

fPrg.Width = (bttCol * (icoW + 2) + (bttS * (bttCol + 1)) + ((icoS * bttCol) * 2) + 6) * 15
fPrg.Height = (bttRow * (icoH + 2) + (bttS * (bttRow + 1)) + ((icoS * bttRow) * 2) + 11 + MBttH) * 15

frmW = fPrg.Width / 15
frmH = fPrg.Height / 15

fPrg.BackColor = RGB(MapC(1), MapC(2), MapC(3))
fPrg.Cls
If wDrwBtt = 1 Then Call GetActivPic

FB = Screen.Height - fPrg.Height
txtW = fPrg.pKntTime.TextWidth("00:00:00")

If FormPos = 0 Then InfT = frmH - 6 - MBttH Else InfT = 2

Btt.Left = 0: Btt.Top = 0: Btt.Right = frmW: Btt.Bottom = frmH: Call DrawBorder(Btt, 0) 'osnovnaya ramka
'ramka knopok
Btt.Left = 2
If FormPos = 0 Then Btt.Top = 2 Else Btt.Top = 7 + MBttH
Btt.Right = frmW - 4: Btt.Bottom = frmH - 9 - MBttH: Call DrawBorder(Btt, 1) 'ramka yarlykov

Btt.Left = 2: Btt.Top = InfT: Btt.Right = 3 + (MBttW + 1) * 5: Btt.Bottom = 4 + MBttH: Call DrawBorder(Btt, 1) 'knopki
If TimeNotShow = 0 Then
  Btt.Left = 6 + (MBttW + 1) * 5: Btt.Top = InfT: Btt.Right = frmW - (txtW + 23 + MBttW * 6): Btt.Bottom = 4 + MBttH: Call DrawBorder(Btt, 1) 'info
  Btt.Left = frmW - (txtW + 11 + MBttW): Btt.Top = InfT: Btt.Right = txtW + 4: Btt.Bottom = 4 + MBttH: Call DrawBorder(Btt, 1) 'time
Else
  Btt.Left = 6 + (MBttW + 1) * 5: Btt.Top = InfT: Btt.Right = frmW - (18 + MBttW * 6): Btt.Bottom = 4 + MBttH
  fPrg.Line (Btt.Left, InfT - 1)-(Btt.Left + Btt.Right, InfT - 3 + MBttH), RGB(MapC(1), MapC(2), MapC(3)), BF
  Call DrawBorder(Btt, 1)  'info
End If
Btt.Left = frmW - 6 - MBttW: Btt.Top = InfT: Btt.Right = 4 + MBttW: Btt.Bottom = 4 + MBttH: Call DrawBorder(Btt, 1) 'zakryt'

If wDrwBtt = 1 Then
  For I = 0 To bttCol - 1 Step 1
    For II = 0 To bttRow - 1 Step 1
      If MapB((II + 1) * 100 + (I + 1)).wClr(0) = 0 Then
        Call DrawBorder(BttCoord((II + 1) * 100 + (I + 1)), 2)
      Else
        Call DrawBorder(BttCoord((II + 1) * 100 + (I + 1)), 2, 1, MapB((II + 1) * 100 + (I + 1)).wClr(1), MapB((II + 1) * 100 + (I + 1)).wClr(2), MapB((II + 1) * 100 + (I + 1)).wClr(3))
      End If
    Next II
  Next I
End If

Call DrawMenuIcons

'knopka vyhoda iz programmy
Call DrawBorder(BttCoord(-1), 2)
fPrg.PaintPicture fPrg.pMenuIco.Image, frmW - 3 - MBttW, InfT + 3, MBttW - 2, MBttH - 2, 0, 0, 7, 7
'knopka nastroek
Call DrawBorder(BttCoord(-2), 2)
fPrg.PaintPicture fPrg.pMenuIco.Image, 5 + (MBttW + 1) * 0, InfT + 3, MBttW - 2, MBttH - 2, 7, 0, 7, 7
'knopka o programme
Call DrawBorder(BttCoord(-3), 2)
fPrg.PaintPicture fPrg.pMenuIco.Image, 5 + (MBttW + 1) * 1, InfT + 3, MBttW - 2, MBttH - 2, 14, 0, 7, 7
'knopka zakrepit'
If FormNotHide = 0 Then Call DrawBorder(BttCoord(-4), 2) Else Call DrawBorder(BttCoord(-4), 4)
fPrg.PaintPicture fPrg.pMenuIco.Image, 5 + (MBttW + 1) * 2, InfT + 3, MBttW - 2, MBttH - 2, 21, 0, 7, 7
'knopka poverh vseh okon
If FormNotTop = 0 Then Call DrawBorder(BttCoord(-5), 4) Else Call DrawBorder(BttCoord(-5), 2)
fPrg.PaintPicture fPrg.pMenuIco.Image, 5 + (MBttW + 1) * 3, InfT + 3, MBttW - 2, MBttH - 2, 28, 0, 7, 7
'knopka otlova goryachih klavish
If FormNotHotKey = 0 Then Call DrawBorder(BttCoord(-6), 4) Else Call DrawBorder(BttCoord(-6), 2)
fPrg.PaintPicture fPrg.pMenuIco.Image, 5 + (MBttW + 1) * 4, InfT + 3, MBttW - 2, MBttH - 2, 35, 0, 7, 7

fPrg.ForeColor = ClrFnt
fPrg.pKntIco.BackColor = RGB(MapC(1), MapC(2), MapC(3)): fPrg.pKntIco.Tag = fPrg.pKntIco.BackColor
fPrg.pKntIco.ForeColor = ClrFnt
fPrg.pKntIco.Width = icoW
fPrg.pKntIco.Height = icoH
fPrg.pKnt.BackColor = RGB(MapC(1), MapC(2), MapC(3))
fPrg.pKnt.ForeColor = ClrFnt
fPrg.pKntTime.BackColor = RGB(MapC(1), MapC(2), MapC(3))
fPrg.pKntTime.ForeColor = ClrFnt
If TimeNotShow = 0 Then
  tT = frmW - (txtW + 27 + MBttW * 6): If tT < 0 Then tT = 0
Else
  tT = frmW - (22 + MBttW * 6): If tT < 0 Then tT = 0
End If
fPrg.pKnt.Width = tT
fPrg.pKnt.Height = MBttH
fPrg.pKnt.Left = 8 + (MBttW + 1) * 5
If FormPos = 0 Then fPrg.pKnt.Top = frmH - MBttH - 4 Else fPrg.pKnt.Top = 4

fPrg.pKntTime.Width = txtW
fPrg.pKntTime.Height = MBttH
fPrg.pKntTime.Left = frmW - (txtW + 9 + MBttW)
If FormPos = 0 Then fPrg.pKntTime.Top = frmH - MBttH - 4 Else fPrg.pKntTime.Top = 4

If wDrwBtt = 1 Then Call DrawLinks
Call GetActivPic
End Sub

Private Sub DrawMenuIcons()
fPrg.pMenuIco.BackColor = RGB(MapC(1), MapC(2), MapC(3))
fPrg.pMenuIco.ForeColor = ClrFnt
'knopka vyhoda iz programmy
fPrg.pMenuIco.Line (1, 1)-(6, 6), ClrFnt
fPrg.pMenuIco.Line (5, 1)-(0, 6), ClrFnt
'knopka nastroek
fPrg.pMenuIco.Line (8, 2)-(9, 5), ClrFnt, B
fPrg.pMenuIco.Line (10, 3)-(13, 3), ClrFnt
'knopka o programme
fPrg.pMenuIco.Line (16, 1)-(19, 1), ClrFnt
fPrg.pMenuIco.Line (17, 1)-(17, 6), ClrFnt
fPrg.pMenuIco.Line (16, 5)-(19, 5), ClrFnt
'knopka zakrepit'
fPrg.pMenuIco.Line (23, 1)-(25, 3), ClrFnt, B
fPrg.pMenuIco.Line (22, 3)-(26, 4), ClrFnt, B
fPrg.pMenuIco.PSet (24, 5), ClrFnt
'knopka poverh vseh okon
fPrg.pMenuIco.Line (29, 1)-(32, 4), ClrFnt, B
fPrg.pMenuIco.Line (30, 5)-(34, 5), ClrFnt
fPrg.pMenuIco.Line (33, 2)-(33, 6), ClrFnt
'knopka otlova goryachih klavish
fPrg.pMenuIco.Line (37, 1)-(39, 2), ClrFnt, BF
fPrg.pMenuIco.Line (36, 3)-(40, 5), ClrFnt, BF
End Sub

Private Sub DrawLinks()
Dim I As Integer, II As Integer
Dim wX As Integer, wY As Integer
Dim pB As Integer
Dim Ic As Long 'dlya iconok
For I = 0 To bttCol - 1 Step 1
  For II = 0 To bttRow - 1 Step 1
    pB = (II + 1) * 100 + I + 1
    If MapB(pB).wOpr > 0 Then
      If MapB(pB).wClr(0) > 0 Then fPrg.pKntIco.BackColor = RGB(MapB(pB).wClr(1), MapB(pB).wClr(2), MapB(pB).wClr(3)) Else fPrg.pKntIco.BackColor = fPrg.pKntIco.Tag
      Call DrawIco(MapB(pB).wIFl, MapB(pB).wINm): Call DrawHK(pB)
      fPrg.PaintPicture fPrg.pKntIco.Image, BttCoord(pB).Left + 1, BttCoord(pB).Top + 1
    End If
  Next II
Next I
End Sub

Public Sub GetActivPic()
'If MapC(5) = 0 Then
  fPrg.pActiv0.Width = fPrg.ScaleWidth: fPrg.pActiv0.Height = fPrg.ScaleHeight
  fPrg.pActiv1.Width = fPrg.ScaleWidth: fPrg.pActiv1.Height = fPrg.ScaleHeight
  
  fPrg.pActiv0.PaintPicture fPrg.Image, 0, 0
  
  fPrg.pActiv1.BackColor = RGB(200, 200, 200)
  fPrg.pActiv1.PaintPicture fPrg.Image, 0, 0, , , , , , , vbSrcAnd
'End If
End Sub

Public Sub DrawBorder(ByRef wBtt As RECT, ByRef wState As Byte, Optional ByRef wFll As Byte = 0, Optional ByRef wRR As Byte = 0, Optional ByRef wGG As Byte = 0, Optional ByRef wBB As Byte = 0)
' // 0-bordyur // 2-knopka // 4-nazhataya knopka // 6-vydelennaya knopka // '
Dim C1 As Long, C2 As Long

wBtt.Right = wBtt.Right - 1
wBtt.Bottom = wBtt.Bottom - 1

If wBtt.Left > -1 Then
  If fPrg.ScaleWidth = fPrg.pActiv0.Width Then
    fPrg.PaintPicture fPrg.pActiv0.Image, wBtt.Left, wBtt.Top, wBtt.Right, wBtt.Bottom, wBtt.Left, wBtt.Top, wBtt.Right, wBtt.Bottom
  End If
End If

If wFll = 0 Then
  wRR = MapC(1)
  wGG = MapC(2)
  wBB = MapC(3)
End If

If wState = 0 Then C1 = GenColor(MapC(4), wRR, wGG, wBB): C2 = GenColor(-MapC(4), wRR, wGG, wBB)
If wState = 1 Then C2 = GenColor(MapC(4), wRR, wGG, wBB): C1 = GenColor(-MapC(4), wRR, wGG, wBB)
If wState = 2 Then C1 = GenColor(MapC(4), wRR, wGG, wBB): C2 = GenColor(-MapC(4), wRR, wGG, wBB)
If wState = 3 Then C2 = GenColor(MapC(4), wRR, wGG, wBB): C1 = GenColor(-MapC(4), wRR, wGG, wBB)
If wState = 4 Then C1 = GenColor(-MapC(4), wRR, wGG, wBB): C2 = GenColor(MapC(4), wRR, wGG, wBB)
If wState = 5 Then C2 = GenColor(-MapC(4), wRR, wGG, wBB): C1 = GenColor(MapC(4), wRR, wGG, wBB)
If MapC(5) = 1 Then
  If wState = 6 Then C1 = GenColor(MapC(4), MapC(6), MapC(7), MapC(8)): C2 = GenColor(-MapC(4), MapC(6), MapC(7), MapC(8))
  If wState = 7 Then C2 = GenColor(MapC(4), MapC(6), MapC(7), MapC(8)): C1 = GenColor(-MapC(4), MapC(6), MapC(7), MapC(8))
End If

1000
If wState = 6 Or wState = 7 Then
  If wBtt.Left > -1 Then
    If MapC(5) = 0 Then
      If fPrg.ScaleWidth = fPrg.pActiv1.Width Then
        fPrg.PaintPicture fPrg.pActiv1.Image, wBtt.Left + 1, wBtt.Top + 1, wBtt.Right - 1, wBtt.Bottom - 1, wBtt.Left + 1, wBtt.Top + 1, wBtt.Right - 1, wBtt.Bottom - 1
      End If
    Else
      wState = wState - 4
      GoTo 1000
    End If
  End If
Else
  fPrg.Line (wBtt.Left, wBtt.Top)-(wBtt.Left, wBtt.Top + wBtt.Bottom), C1
  fPrg.Line (wBtt.Left, wBtt.Top)-(wBtt.Left + wBtt.Right + 1, wBtt.Top), C1
  fPrg.Line (wBtt.Left + wBtt.Right, wBtt.Top + wBtt.Bottom)-(wBtt.Left + wBtt.Right, wBtt.Top), C2
  fPrg.Line (wBtt.Left + wBtt.Right, wBtt.Top + wBtt.Bottom)-(wBtt.Left - 1, wBtt.Top + wBtt.Bottom), C2
End If
End Sub

Public Function GenColor(ByVal wPlus As Integer, ByVal wR As Integer, ByVal wG As Integer, ByVal wB As Integer) As Long
wR = wR + wPlus
wG = wG + wPlus
wB = wB + wPlus
If wR < 0 Then wR = 0
If wG < 0 Then wG = 0
If wB < 0 Then wB = 0
GenColor = RGB(wR, wG, wB)
End Function

Public Sub SetTransparent(hwnd As Long, Layered As Byte, Optional ByRef wTip As Byte = 0)
Dim Ret As Long
Ret = GetWindowLong(hwnd, -20)
If wTip = 1 Then 'glavnoe okno programmy
  Ret = Ret Or &H80000
Else 'ne light bar
  Ret = Ret Xor &H80000
End If
SetWindowLong hwnd, -20, Ret
SetLayeredWindowAttributes hwnd, 0, Layered, &H2
End Sub

Public Sub LoadColors()
ClrFnt = RGB(0, 25, 50)
ClrFrm = RGB(200, 0, 0)
MapC(1) = 150
MapC(2) = 175
MapC(3) = 200
MapC(4) = 50
MapC(5) = 0
MapC(6) = 200
MapC(7) = 200
MapC(8) = 0
End Sub

Public Sub LoadStt()
Dim I As Integer
Dim SttVer As Long
Dim SS As String
Dim RR As Long 'poziciya znaka '='
Dim ZZ As Long 'nomer zapisi
Dim MapS() As String
Dim MapSS() As String
Dim SS2() As String

bttS = 1

'load settings
If Dir(SttPath, vbNormal) <> "" Then
  FF = FreeFile
  Open SttPath For Input As #FF
    Do
      If EOF(FF) = True Then Exit Do
      Line Input #FF, SS
      If SS <> "" Then
        RR = InStr(SS, "=")
        If RR > 0 Then
          MapS = Split(SS, "=")
          MapS(0) = Trim(MapS(0))
          For I = 2 To UBound(MapS) Step 1
            MapS(1) = MapS(1) & "=" & MapS(I)
          Next I
          MapS(1) = Trim(MapS(1))
          
          If MapS(0) = "SttVer" Then SttVer = Val(MapS(1))
          
          If MapS(0) = "BttCol" Then bttCol = Val(MapS(1))
          If MapS(0) = "BttRow" Then bttRow = Val(MapS(1))
          If MapS(0) = "BttSpace" Then bttS = Val(MapS(1))
          If MapS(0) = "IcoWidth" Then icoW = Val(MapS(1))
          If MapS(0) = "IcoHeight" Then icoH = Val(MapS(1))
          If MapS(0) = "IcoSpace" Then icoS = Val(MapS(1))
          If MapS(0) = "FormLeft" Then FormLeft = Val(MapS(1))
          If MapS(0) = "FormTop" Then FormTop = Val(MapS(1))
          If MapS(0) = "FormTrans" Then TransForm = Val(MapS(1))
          If MapS(0) = "FormDelay" Then fPrg.tZdr.Interval = Val(MapS(1))
          If MapS(0) = "FormAnimate" Then fPrg.tShow.Tag = Val(MapS(1))
          If MapS(0) = "HotMod" Then HotMod = Val(MapS(1))
          If MapS(0) = "HotKey" Then HotKey = Val(MapS(1))
          If MapS(0) = "FormNotHide" Then FormNotHide = Val(MapS(1))
          If MapS(0) = "FormNotTop" Then FormNotTop = Val(MapS(1))
          If MapS(0) = "FormNotHotKey" Then FormNotHotKey = Val(MapS(1))
          If MapS(0) = "FormPosition" Then FormPos = Val(MapS(1))
          If MapS(0) = "DrawHotKey" Then DrawHotKey = Val(MapS(1))
          If MapS(0) = "MenuBttWidth" Then MBttW = Val(MapS(1))
          If MapS(0) = "MenuBttHeight" Then MBttH = Val(MapS(1))
          If MapS(0) = "TimeNotShow" Then TimeNotShow = Val(MapS(1))
          If MapS(0) = "ShowInTray" Then ShowInTray = Val(MapS(1))
          If MapS(0) = "BttToShow" Then BttToShow = Val(MapS(1))
          If MapS(0) = "LangFile" Then LangFile = MapS(1)
          If MapS(0) = "MainFont" Then FntName = MapS(1)
          If MapS(0) = "NotAutoFocus" Then NotAutoFocus = Val(MapS(1))
          If MapS(0) = "NotClearMem" Then NotClearMem = Val(MapS(1))
          If MapS(0) = "FormFullHide" Then fPrg.tFFH.Enabled = Val(MapS(1))
          
          If MapS(0) = "ClrFnt" Then ClrFnt = AssembledColor(MapS(1))
          If MapS(0) = "ClrFrm" Then ClrFrm = AssembledColor(MapS(1))
          If MapS(0) = "Clr0" Then MapC(0) = MapS(1)
          If MapS(0) = "Clr1" Then MapC(1) = MapS(1)
          If MapS(0) = "Clr2" Then MapC(2) = MapS(1)
          If MapS(0) = "Clr3" Then MapC(3) = MapS(1)
          If MapS(0) = "Clr4" Then MapC(4) = MapS(1)
          If MapS(0) = "Clr5" Then MapC(5) = MapS(1)
          If MapS(0) = "Clr6" Then MapC(6) = MapS(1)
          If MapS(0) = "Clr7" Then MapC(7) = MapS(1)
          If MapS(0) = "Clr8" Then MapC(8) = MapS(1)
          
          If Val(MapS(0)) > 99 Then
            If Val(MapS(0)) < 2200 Then
              MapSS = Split(MapS(1), ",")
              
              If SttVer < 72 Then
                'do jetoj versii nebylo polya dlya modifikatora dlya hotkeya
                If UBound(MapSS) = 10 Then
                  ReDim Preserve MapSS(11)
                  For I = 11 To 7 Step -1
                    MapSS(I) = MapSS(I - 1)
                  Next I
                  MapSS(6) = "0"
                End If
              End If
              If SttVer < 87 Then
                'do jetogo byli tolk'ko klick i doubleClick, teper' u nih drugie nomera
                If Val(Trim(MapSS(0))) = 1 Then MapSS(0) = "3"
                If Val(Trim(MapSS(0))) = 2 Then MapSS(0) = "4"
              End If
              If SttVer < 201 Then
                'do jetogo avtozagruzka byla v reestre
                If GetRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "LightBar") <> "" Then 'esli byla avtozagruzka iz reestra
                  Call AddToAutorun
                  Call DeleteRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "LightBar", False) 'udalyaem zapis' iz reestra
                  Call fMsg.GetMsg(fPrg, 1, MapMsg(5))
                End If
              End If
              If SttVer < 330479 Then
                'do etoj versii v knopkah nebylo polej s cvetom knopki
                If UBound(MapSS) = 11 Then
                  ReDim Preserve MapSS(16)
                  For I = 12 To 16 Step 1
                    MapSS(I) = "0"
                  Next I
                End If
              End If
              
              If UBound(MapSS) > 15 Then
                For I = 0 To 11 Step 1
                  MapSS(I) = Trim(MapSS(I))
                  MapSS(I) = Replace(MapSS(I), "~'~", ",")
                Next I
                ZZ = CLng(MapS(0))
                MapB(ZZ).wOpr = Val(MapSS(0))
                MapB(ZZ).wFil = MapSS(1)
                MapB(ZZ).wPrm = MapSS(2)
                MapB(ZZ).wDir = MapSS(3)
                MapB(ZZ).wShw = Val(MapSS(4))
                MapB(ZZ).wCap = MapSS(5)
                MapB(ZZ).wHtM = Val(MapSS(6))
                MapB(ZZ).wHtK = Val(MapSS(7))
                MapB(ZZ).wHtN = Val(MapSS(8))
                MapB(ZZ).wIFl = MapSS(9)
                MapB(ZZ).wINm = Val(MapSS(10))
                MapB(ZZ).wKmm = MapSS(11)
                MapB(ZZ).wClr(0) = MapSS(12)
                MapB(ZZ).wClr(1) = MapSS(13)
                MapB(ZZ).wClr(2) = MapSS(14)
                MapB(ZZ).wClr(3) = MapSS(15)
                MapB(ZZ).wClr(4) = MapSS(16)
              End If
            End If
          End If
        End If
      End If
    Loop
  Close #FF
End If

'proverka bagov
If bttCol > 90 Or bttCol < 10 Then bttCol = 20
If bttRow > 20 Or bttRow < 1 Then bttRow = 5
If bttS > 50 Or bttS < 0 Then bttS = 1
If icoW > 64 Or icoW < 4 Then icoW = 16
If icoH > 64 Or icoH < 4 Then icoH = 16
If icoS > 20 Or icoS < 0 Then icoS = 0
If FormLeft > Screen.Width - 225 Or FormLeft < 0 Then FormLeft = 750
If FormTop > Screen.Height - 225 Or FormTop < 0 Then FormTop = 0
If TransForm > 255 Or TransForm < 50 Then TransForm = 200
If MBttW > 32 Or MBttW < 9 Then MBttW = 11
If MBttH > 32 Or MBttH < 11 Then MBttH = 11
If fPrg.tZdr.Interval > 5000 Or fPrg.tZdr.Interval < 0 Then fPrg.tZdr.Interval = 0
If fPrg.tZdr.Interval = 0 Then fPrg.tFFH.Enabled = False
If CInt(fPrg.tShow.Tag) > 100 Or CInt(fPrg.tShow.Tag) < 0 Then fPrg.tShow.Tag = 10

SS2 = Split(FntName, ",")
If UBound(SS2) = 3 Then ReDim Preserve SS2(4): SS2(4) = -2
If UBound(SS2) <> 4 Then
  FntName = "MS Sans Serif"
  FntSize = 8
  FntBold = 1
  FntItalic = 0
  FntTop = -2
Else
  FntName = SS2(0)
  FntSize = Val(SS2(1)): If FntSize < 2 Or FntSize > 72 Then FntSize = 8
  FntBold = Val(SS2(2)): If FntBold <> 0 And FntBold <> 1 Then FntBold = 1
  FntItalic = Val(SS2(3)): If FntItalic <> 0 And FntItalic <> 1 Then FntItalic = 0
  FntTop = Val(SS2(4)): If FntTop < -50 Or FntTop > 50 Then FntTop = -2
End If
fPrg.pKnt.FontName = FntName
fPrg.pKnt.FontSize = FntSize
fPrg.pKnt.FontBold = FntBold
fPrg.pKnt.FontItalic = FntItalic
fPrg.pKntTime.FontName = FntName
fPrg.pKntTime.FontSize = FntSize
fPrg.pKntTime.FontBold = FntBold
fPrg.pKntTime.FontItalic = FntItalic

For I = 99 To 2199 Step 1
  If MapB(I).wOpr > 0 Then
    If MapB(I).wOpr > 4 Then MapB(I).wOpr = 0
    If MapB(I).wShw < 0 Or MapB(I).wShw > 2 Then MapB(I).wShw = 0
    If MapB(I).wHtN < 0 Or MapB(I).wHtN > 9 Then MapB(I).wHtN = 0
  End If
Next I

'peresohranyaem fail
Call SaveStt
End Sub

Public Sub SaveStt()
Dim I As Integer
Dim SS As String
FF = FreeFile
Open SttPath For Output As #FF
  Print #FF, "[Global]"
  
  Print #FF, "SttVer=" & CStr(CLng(App.Major) * 1000000 + CLng(App.Minor) * 10000 + App.Revision)
  
  Print #FF, "[General]"
  
  Print #FF, "BttCol=" & CStr(bttCol)
  Print #FF, "BttRow=" & CStr(bttRow)
  Print #FF, "BttSpace=" & CStr(bttS)
  Print #FF, "IcoWidth=" & CStr(icoW)
  Print #FF, "IcoHeight=" & CStr(icoH)
  Print #FF, "IcoSpace=" & CStr(icoS)
  Print #FF, "FormLeft=" & CStr(FormLeft)
  Print #FF, "FormTop=" & CStr(FormTop)
  Print #FF, "FormTrans=" & CStr(TransForm)
  Print #FF, "FormDelay=" & CStr(fPrg.tZdr.Interval)
  Print #FF, "FormAnimate=" & CStr(fPrg.tShow.Tag)
  Print #FF, "HotMod=" & CStr(HotMod)
  Print #FF, "HotKey=" & CStr(HotKey)
  Print #FF, "FormNotHide=" & CStr(FormNotHide)
  Print #FF, "FormNotTop=" & CStr(FormNotTop)
  Print #FF, "FormNotHotKey=" & CStr(FormNotHotKey)
  Print #FF, "FormPosition=" & CStr(FormPos)
  Print #FF, "DrawHotKey=" & CStr(DrawHotKey)
  Print #FF, "MenuBttWidth=" & CStr(MBttW)
  Print #FF, "MenuBttHeight=" & CStr(MBttH)
  Print #FF, "TimeNotShow=" & CStr(TimeNotShow)
  Print #FF, "ShowInTray=" & CStr(ShowInTray)
  Print #FF, "BttToShow=" & CStr(BttToShow)
  Print #FF, "LangFile=" & LangFile
  Print #FF, "MainFont=" & FntName & "," & CStr(FntSize) & "," & CStr(FntBold) & "," & CStr(FntItalic) & "," & CStr(FntTop)
  Print #FF, "NotAutoFocus=" & CStr(NotAutoFocus)
  Print #FF, "NotClearMem=" & CStr(NotClearMem)
  If fPrg.tFFH.Enabled = True Then Print #FF, "FormFullHide=1"
  If fPrg.tFFH.Enabled = False Then Print #FF, "FormFullHide=0"
  
  Print #FF, "[Colors]"
  Print #FF, "ClrFnt=" & SeparateColor(ClrFnt)
  Print #FF, "ClrFrm=" & SeparateColor(ClrFrm)
  Print #FF, "Clr0=" & MapC(0)
  Print #FF, "Clr1=" & MapC(1)
  Print #FF, "Clr2=" & MapC(2)
  Print #FF, "Clr3=" & MapC(3)
  Print #FF, "Clr4=" & MapC(4)
  Print #FF, "Clr5=" & MapC(5)
  Print #FF, "Clr6=" & MapC(6)
  Print #FF, "Clr7=" & MapC(7)
  Print #FF, "Clr8=" & MapC(8)
  
  Print #FF, "[Buttons]"
  
  For I = 99 To 2199 Step 1
    SS = CStr(MapB(I).wOpr) _
 & "," & Replace(MapB(I).wFil, ",", "~'~") _
 & "," & Replace(MapB(I).wPrm, ",", "~'~") _
 & "," & Replace(MapB(I).wDir, ",", "~'~") _
 & "," & CStr(MapB(I).wShw) _
 & "," & Replace(MapB(I).wCap, ",", "~'~") _
 & "," & CStr(MapB(I).wHtM) _
 & "," & CStr(MapB(I).wHtK) _
 & "," & CStr(MapB(I).wHtN) _
 & "," & Replace(MapB(I).wIFl, ",", "~'~") _
 & "," & CStr(MapB(I).wINm) _
 & "," & Replace(MapB(I).wKmm, ",", "~'~") _
 & "," & MapB(I).wClr(0) _
 & "," & MapB(I).wClr(1) _
 & "," & MapB(I).wClr(2) _
 & "," & MapB(I).wClr(3) _
 & "," & MapB(I).wClr(4)
    If Left$(SS, 20) <> "0,,,,0,,0,0,0,,0,,0," Then Print #FF, I & "=" & SS
  Next I
  
  Print #FF, "[End]"
Close #FF
End Sub

'##################################################################################################'
'######## DLYA ICONOK #############################################################################'
'##################################################################################################'

Public Function GetIcoCount(ByVal wFile As String) As Long
wFile = GetDir(wFile)
GetIcoCount = ExtractIconEx(wFile, -1, 0, 0, 0)
End Function

Public Sub DrawIco(ByVal wFile As String, ByRef wNum As Long)
Static schDraw As Integer
schDraw = schDraw + 1
If schDraw > 8000 Then
  If schDraw = 8001 Then Call fMsg.GetMsg(fPrg, 0, "Error!" & vbCrLf & MapMsg(6))
  GoSub DrawNoIco
End If
Dim Ic As Long
wFile = GetDir(wFile)
fPrg.pKntIco.Cls
If wNum = 0 Then GoSub DrawNoIco
If wNum > GetIcoCount(wFile) Then GoSub DrawNoIco
Call ExtractIconEx(wFile, wNum - 1, Ic, 0, 1)
Call DrawIconEx(fPrg.pKntIco.hdc, 0, 0, Ic, icoW, icoH, 0, 0, 3)
Exit Sub
DrawNoIco:
fPrg.pKntIco.PaintPicture fPrg.iNoIco.Picture, 0, 0, icoW, icoH
fPrg.pKntIco.Refresh
End Sub

Public Sub DrawHK(ByVal wNum As Integer)
If DrawHotKey > 0 Then
  If MapB(wNum).wHtK > 0 Then
    fPrg.pKntIco.CurrentX = fPrg.pKntIco.Width / 2 - fPrg.pKntIco.TextWidth(MapKN(MapB(wNum).wHtK)) / 2
    fPrg.pKntIco.CurrentY = fPrg.pKntIco.Height / 2 - fPrg.pKntIco.TextHeight(MapKN(MapB(wNum).wHtK)) / 2
    fPrg.pKntIco.Print MapKN(MapB(wNum).wHtK)
  End If
End If
End Sub

'##################################################################################################'
'##################################################################################################'
'##################################################################################################'

Public Function GetDir(ByRef wStr As String) As String
GetDir = wStr

GetDir = Replace(GetDir, "%ALLUSERSPROFILE%", Environ("ALLUSERSPROFILE"), , , 1)
GetDir = Replace(GetDir, "%APPDATA%", Environ("APPDATA"), , , vbTextCompare)
GetDir = Replace(GetDir, "%COMMONPROGRAMFILES%", Environ("COMMONPROGRAMFILES"), , , 1)
GetDir = Replace(GetDir, "%HOMEDRIVE%", Environ("HOMEDRIVE"), , , 1)
GetDir = Replace(GetDir, "%HOMEPATH%", Environ("HOMEPATH"), , , 1)
GetDir = Replace(GetDir, "%PROGRAMFILES%", Environ("PROGRAMFILES"), , , 1)
GetDir = Replace(GetDir, "%SYSTEMDRIVE%", Environ("SYSTEMDRIVE"), , , 1)
GetDir = Replace(GetDir, "%SYSTEMROOT%", Environ("SYSTEMROOT"), , , 1)
GetDir = Replace(GetDir, "%TEMP%", Environ("TEMP"), , , 1)
GetDir = Replace(GetDir, "%TMP%", Environ("TMP"), , , 1)
GetDir = Replace(GetDir, "%USERPROFILE%", Environ("USERPROFILE"), , , 1)
GetDir = Replace(GetDir, "%WINDIR%", Environ("WINDIR"), , , 1)

GetDir = Replace(GetDir, "%PROGPATH%", App.Path, , , 1)
GetDir = Replace(GetDir, "%MUSIC%", GetRegString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Music"), , , 1)
GetDir = Replace(GetDir, "%DESKTOP%", SpecialFolder(0), , , 1)
GetDir = Replace(GetDir, "%STARTMENUPROG%", SpecialFolder(2), , , 1)
GetDir = Replace(GetDir, "%DOCUMENTS%", SpecialFolder(5), , , 1)
GetDir = Replace(GetDir, "%FAVORITES%", SpecialFolder(6), , , 1)
GetDir = Replace(GetDir, "%AUTORUN%", SpecialFolder(7), , , 1)
GetDir = Replace(GetDir, "%RECENT%", SpecialFolder(8), , , 1)
GetDir = Replace(GetDir, "%SENDTO%", SpecialFolder(9), , , 1)
GetDir = Replace(GetDir, "%STARTMENU%", SpecialFolder(11), , , 1)
GetDir = Replace(GetDir, "%VIDEO%", SpecialFolder(14), , , 1)
GetDir = Replace(GetDir, "%DESKTOP2%", SpecialFolder(16), , , 1)
GetDir = Replace(GetDir, "%NETHOOD%", SpecialFolder(19), , , 1)
GetDir = Replace(GetDir, "%FONTS%", SpecialFolder(20), , , 1)
GetDir = Replace(GetDir, "%TEMPLATES%", SpecialFolder(21), , , 1)
GetDir = Replace(GetDir, "%AUSTARTMENU%", SpecialFolder(22), , , 1)
GetDir = Replace(GetDir, "%AUSTARTMENUPROG%", SpecialFolder(23), , , 1)
GetDir = Replace(GetDir, "%AUAUTORUN%", SpecialFolder(24), , , 1)
GetDir = Replace(GetDir, "%AUDESKTOP%", SpecialFolder(25), , , 1)
GetDir = Replace(GetDir, "%APPLICATIONDATA%", SpecialFolder(26), , , 1)
GetDir = Replace(GetDir, "%PRINTHOOD%", SpecialFolder(27), , , 1)
GetDir = Replace(GetDir, "%LOCALSETTAPPDATA%", SpecialFolder(28), , , 1)
GetDir = Replace(GetDir, "%AUFAVORITES%", SpecialFolder(31), , , 1)
GetDir = Replace(GetDir, "%CASHE%", SpecialFolder(32), , , 1)
GetDir = Replace(GetDir, "%COOKIES%", SpecialFolder(33), , , 1)
GetDir = Replace(GetDir, "%HISTORY%", SpecialFolder(34), , , 1)
GetDir = Replace(GetDir, "%AUAPPDATA%", SpecialFolder(35), , , 1)
GetDir = Replace(GetDir, "%WINDOWS%", SpecialFolder(36), , , 1)
GetDir = Replace(GetDir, "%SYSTEM32%", SpecialFolder(37), , , 1)
GetDir = Replace(GetDir, "%PROGRAMDIR%", SpecialFolder(38), , , 1)
GetDir = Replace(GetDir, "%PICTURES%", SpecialFolder(39), , , 1)
GetDir = Replace(GetDir, "%USERDIR%", SpecialFolder(40), , , 1)
GetDir = Replace(GetDir, "%SYSTEM322%", SpecialFolder(41), , , 1)
GetDir = Replace(GetDir, "%COMMONFILES%", SpecialFolder(43), , , 1)
GetDir = Replace(GetDir, "%AUTEMPLATES%", SpecialFolder(45), , , 1)
GetDir = Replace(GetDir, "%AUDOCUMENTS%", SpecialFolder(46), , , 1)
GetDir = Replace(GetDir, "%ADMINISTRATION%", SpecialFolder(47), , , 1)
GetDir = Replace(GetDir, "%AUMUSIC%", SpecialFolder(53), , , 1)
GetDir = Replace(GetDir, "%AUPICTURES%", SpecialFolder(54), , , 1)
GetDir = Replace(GetDir, "%AUVIDEO%", SpecialFolder(55), , , 1)
GetDir = Replace(GetDir, "%RESOURCES%", SpecialFolder(56), , , 1)
GetDir = Replace(GetDir, "%CDBURNING%", SpecialFolder(59), , , 1)
End Function

Public Function DeGetDir(ByRef wStr As String) As String
DeGetDir = wStr

DeGetDir = Replace(DeGetDir, Environ("SYSTEMROOT"), "%SYSTEMROOT%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("WINDIR"), "%WINDIR%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("ALLUSERSPROFILE"), "%ALLUSERSPROFILE%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("APPDATA"), "%APPDATA%", , , vbTextCompare)
DeGetDir = Replace(DeGetDir, Environ("COMMONPROGRAMFILES"), "%COMMONPROGRAMFILES%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("HOMEDRIVE"), "%HOMEDRIVE%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("HOMEPATH"), "%HOMEPATH%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("PROGRAMFILES"), "%PROGRAMFILES%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("SYSTEMDRIVE"), "%SYSTEMDRIVE%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("TMP"), "%TMP%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("TEMP"), "%TEMP%", , , 1)
DeGetDir = Replace(DeGetDir, Environ("USERPROFILE"), "%USERPROFILE%", , , 1)

DeGetDir = Replace(DeGetDir, App.Path, "%PROGPATH%", , , 1)
DeGetDir = Replace(DeGetDir, GetRegString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Music"), "%MUSIC%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(0), "%DESKTOP%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(2), "%STARTMENUPROG%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(5), "%DOCUMENTS%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(6), "%FAVORITES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(7), "%AUTORUN%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(8), "%RECENT%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(9), "%SENDTO%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(11), "%STARTMENU%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(14), "%VIDEO%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(16), "%DESKTOP2%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(19), "%NETHOOD%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(20), "%FONTS%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(21), "%TEMPLATES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(22), "%AUSTARTMENU%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(23), "%AUSTARTMENUPROG%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(24), "%AUAUTORUN%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(25), "%AUDESKTOP%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(26), "%APPLICATIONDATA%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(27), "%PRINTHOOD%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(28), "%LOCALSETTAPPDATA%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(31), "%AUFAVORITES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(32), "%CASHE%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(33), "%COOKIES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(34), "%HISTORY%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(35), "%AUAPPDATA%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(36), "%WINDOWS%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(37), "%SYSTEM32%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(38), "%PROGRAMDIR%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(39), "%PICTURES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(40), "%USERDIR%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(41), "%SYSTEM322%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(43), "%COMMONFILES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(45), "%AUTEMPLATES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(46), "%AUDOCUMENTS%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(47), "%ADMINISTRATION%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(53), "%AUMUSIC%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(54), "%AUPICTURES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(55), "%AUVIDEO%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(56), "%RESOURCES%", , , 1)
DeGetDir = Replace(DeGetDir, SpecialFolder(59), "%CDBURNING%", , , 1)
End Function

Public Function GetFolder(ByRef wStr As String) As String
GetFolder = Left$(wStr, InStrRev(wStr, "\"))
End Function

Public Function GetName(ByRef wStr As String) As String
GetName = Right$(wStr, Len(wStr) - (InStrRev(wStr, "\")))
If InStr(GetName, ".") > 0 Then GetName = Left$(GetName, InStrRev(GetName, ".") - 1)
If GetName = LCase(GetName) Then GetName = StrConv(GetName, vbProperCase)
End Function

Private Function SeparateColor(ByVal Color As Long) As String
SeparateColor = Color Mod 256: Color = Color \ 256
SeparateColor = SeparateColor & "," & Color Mod 256: Color = Color \ 256
SeparateColor = SeparateColor & "," & Color
End Function

Private Function AssembledColor(ByVal Color As String) As Long
Dim I As Integer
Dim SS() As String
Dim C(2) As Long
SS = Split(Color, ",")
If UBound(SS) > 1 Then
  For I = 0 To 2 Step 1
    C(I) = Val(Trim(SS(I)))
    If C(I) < 0 Or C(I) > 255 Then C(I) = 0
  Next I
  AssembledColor = RGB(C(0), C(1), C(2))
End If
End Function

Public Function GetNum(ByRef wNum As Byte, ByVal wMas As Byte) As Byte 'vozvrachaet "boolean" znachenie nomera (po schyotu) v "massive" nomerov
Dim MapN(7) As Byte
If wMas < 128 Then MapN(7) = 0 Else MapN(7) = 1: wMas = wMas - 128
If wMas < 64 Then MapN(6) = 0 Else MapN(6) = 1: wMas = wMas - 64
If wMas < 32 Then MapN(5) = 0 Else MapN(5) = 1: wMas = wMas - 32
If wMas < 16 Then MapN(4) = 0 Else MapN(4) = 1: wMas = wMas - 16
If wMas < 8 Then MapN(3) = 0 Else MapN(3) = 1: wMas = wMas - 8
If wMas < 4 Then MapN(2) = 0 Else MapN(2) = 1: wMas = wMas - 4
If wMas < 2 Then MapN(1) = 0 Else MapN(1) = 1: wMas = wMas - 2
If wMas < 1 Then MapN(0) = 0 Else MapN(0) = 1: wMas = wMas - 1
GetNum = MapN(wNum)
End Function

Public Function GetTextMod(ByVal wMod As Byte) As String
GetTextMod = ""
If GetNum(0, wMod) = 1 Then GetTextMod = GetTextMod & "LShift+"
If GetNum(1, wMod) = 1 Then GetTextMod = GetTextMod & "RShift+"
If GetNum(2, wMod) = 1 Then GetTextMod = GetTextMod & "LCtrl+"
If GetNum(3, wMod) = 1 Then GetTextMod = GetTextMod & "RCtrl+"
If GetNum(4, wMod) = 1 Then GetTextMod = GetTextMod & "LWin+"
If GetNum(5, wMod) = 1 Then GetTextMod = GetTextMod & "RWin+"
If GetNum(6, wMod) = 1 Then GetTextMod = GetTextMod & "LAlt+"
If GetNum(7, wMod) = 1 Then GetTextMod = GetTextMod & "RAlt+"
If GetTextMod = "" Then GetTextMod = "(No)"
End Function

Public Function GetError(ByRef wBase As Byte, ByRef wNum As Long)
Dim sErr As String
If wBase = 0 Then                                                             ' -= shellexecute =- '
  sErr = "<<< No description >>>"
  If wNum = 0 Then sErr = "The operating system is out of memory or resources."
  If wNum = 2& Then sErr = "ERROR_FILE_NOT_FOUND" & vbCrLf & "The specified file was not found."
  If wNum = 3& Then sErr = "ERROR_PATH_NOT_FOUND" & vbCrLf & "The specified path was not found."
  If wNum = 11& Then sErr = "ERROR_BAD_FORMAT" & vbCrLf & "The .EXE file is invalid (non-Win32 .EXE or error in .EXE image)."
  If wNum = 5 Then sErr = "SE_ERR_ACCESSDENIED" & vbCrLf & "The operating system denied access to the specified file."
  If wNum = 27 Then sErr = "SE_ERR_ASSOCINCOMPLETE" & vbCrLf & "The filename association is incomplete or invalid."
  If wNum = 30 Then sErr = "SE_ERR_DDEBUSY" & vbCrLf & "The DDE transaction could not be completed because other DDE transactions were being processed."
  If wNum = 29 Then sErr = "SE_ERR_DDEFAIL" & vbCrLf & "The DDE transaction failed."
  If wNum = 28 Then sErr = "SE_ERR_DDETIMEOUT" & vbCrLf & "The DDE transaction could not be completed because the request timed out."
  If wNum = 32 Then sErr = "SE_ERR_DLLNOTFOUND" & vbCrLf & "The specified dynamic-link library was not found. "
  If wNum = 2 Then sErr = "SE_ERR_FNF" & vbCrLf & "The specified file was not found. "
  If wNum = 31 Then sErr = "SE_ERR_NOASSOC" & vbCrLf & "There is no application associated with the given filename extension."
  If wNum = 8 Then sErr = "SE_ERR_OOM" & vbCrLf & "There was not enough memory to complete the operation."
  If wNum = 3 Then sErr = "SE_ERR_PNF" & vbCrLf & "The specified path was not found."
  If wNum = 26 Then sErr = "SE_ERR_SHARE" & vbCrLf & "A sharing violation occurred."
  GetError = "Error " & wNum & vbCrLf & sErr
End If
End Function

Public Function CheckBtt(ByVal wX As Integer, ByVal wY As Integer) As Integer
Dim I As Integer, II As Integer
Dim rX As Integer, rY As Integer
Dim BttTop As Integer
If FormPos > 0 Then BttTop = 5 + MBttH
For I = 0 To bttCol - 1 Step 1
  For II = 0 To bttRow - 1 Step 1
    rX = I * (icoW + bttS + 2 + (icoS * 2)) + 2 + bttS
    rY = II * (icoH + bttS + 2 + (icoS * 2)) + 2 + bttS + BttTop
    If wX > rX And wX < rX + icoW + icoS * 2 + 3 And wY > rY And wY < rY + icoH + icoS * 2 + 3 Then
      CheckBtt = (II + 1) * 100 + I + 1
    End If
  Next II
Next I

'proverka na schelchok po kommandnym batonam
If FormPos = 0 Then BttTop = frmH - 6 - MBttH Else BttTop = 2
If wY > BttTop And wY < BttTop + 3 + MBttH Then
  If wX > frmW - 4 - MBttW And wX < frmW - 4 Then CheckBtt = -1 'vyhod
  If wX > 3 + (MBttW + 1) * 0 And wX < 3 + (MBttW + 1) * 0 + (MBttW + 1) Then CheckBtt = -2 'nastrojki
  If wX > 3 + (MBttW + 1) * 1 And wX < 3 + (MBttW + 1) * 1 + (MBttW + 1) Then CheckBtt = -3 'o programme
  If wX > 3 + (MBttW + 1) * 2 And wX < 3 + (MBttW + 1) * 2 + (MBttW + 1) Then CheckBtt = -4 'zakrepit' okno
  If wX > 3 + (MBttW + 1) * 3 And wX < 3 + (MBttW + 1) * 3 + (MBttW + 1) Then CheckBtt = -5 'poverh vseh okon
  If wX > 3 + (MBttW + 1) * 4 And wX < 3 + (MBttW + 1) * 4 + (MBttW + 1) Then CheckBtt = -6 'otlov goryachih klavish
  If wX > 6 + (MBttW + 1) * 5 And wX < 9 + (MBttW + 1) * 5 + fPrg.pKnt.Width Then CheckBtt = -99  'tekstovoe pole
End If
End Function

Public Function BttCoord(ByVal wNum As Integer) As RECT
Dim I As Integer, II As Integer
Dim BttTop As Integer

If wNum = 0 Or wNum = -99 Then
  BttCoord.Left = -1
  BttCoord.Top = -1
  BttCoord.Right = 1
  BttCoord.Bottom = 1
  Exit Function
End If

If FormPos > 0 Then BttTop = 5 + MBttH
If wNum > 0 Then
  I = wNum Mod 100
  II = (wNum - I) / 100
  I = I - 1
  II = II - 1
  BttCoord.Left = I * (icoW + bttS + 2 + (icoS * 2)) + 3 + bttS
  BttCoord.Top = II * (icoH + bttS + 2 + (icoS * 2)) + 3 + bttS + BttTop
  BttCoord.Right = icoW + icoS * 2 + 2
  BttCoord.Bottom = icoH + icoS * 2 + 2
End If
If wNum < 0 Then 'komandnye knopki
  If FormPos = 0 Then BttCoord.Top = frmH - 4 - MBttH Else BttCoord.Top = 4
  BttCoord.Right = MBttW
  BttCoord.Bottom = MBttH
End If
If wNum = -1 Then BttCoord.Left = frmW - 4 - MBttW 'vyhod
If wNum = -2 Then BttCoord.Left = 4 + ((MBttW + 1) * 0) 'nastrojki
If wNum = -3 Then BttCoord.Left = 4 + ((MBttW + 1) * 1) 'o programme
If wNum = -4 Then BttCoord.Left = 4 + ((MBttW + 1) * 2) 'zakrepit' okno
If wNum = -5 Then BttCoord.Left = 4 + ((MBttW + 1) * 3) 'poverh vseh okon
If wNum = -6 Then BttCoord.Left = 4 + ((MBttW + 1) * 4) 'otlov goryachih klavish
End Function

Public Sub GetStatus(ByRef wTyp As Byte, ByRef wCap As String)
fPrg.pKnt.Cls
fPrg.pKnt.CurrentX = 0 '-1
fPrg.pKnt.CurrentY = FntTop

If wTyp = 1 Then 'zapusk po goryachej klavishe
  fPrg.pKnt.Print MapOth(9) & " " & MapB(Val(wCap)).wCap
End If
If wTyp = 2 Then 'polzunok
  fPrg.pKnt.Line (0, 0)-(fPrg.pKnt.ScaleWidth - 1, fPrg.pKnt.ScaleHeight - 2), ClrFnt, B
  fPrg.pKnt.Line (0, 0)-((fPrg.pKnt.ScaleWidth - 1) / 100 * CDbl(wCap), fPrg.pKnt.ScaleHeight - 2), ClrFnt, BF
End If
If wTyp = 3 Then 'lyuboj tekst
  fPrg.pKnt.Print wCap
End If

fPrg.PaintPicture fPrg.pKnt.Image, fPrg.pKnt.Left, fPrg.pKnt.Top
fPrg.Line (0, 0)-(frmW, 0), GenColor(MapC(4), MapC(1), MapC(2), MapC(3))
fPrg.Line (0, frmH - 1)-(frmW, frmH - 1), GenColor(-MapC(4), MapC(1), MapC(2), MapC(3))
If fPrg.ScaleWidth <> frmW Then fPrg.Width = frmW * 15
If fPrg.ScaleHeight <> frmH Then fPrg.Height = frmH * 15
If FormTop = 0 Then
  If FormPos = 0 Then
    If fPrg.Top < 0 Then fPrg.Top = -fPrg.Height + ((MBttH + 7) * 15)
  Else
    If fPrg.Top > FB Then fPrg.Top = Screen.Height - ((MBttH + 7) * 15)
  End If
End If
fPrg.ZOrder
fPrg.tPpp.Enabled = False
fPrg.tPpp.Enabled = True
End Sub

Public Sub BeeBeep(ByRef wNum As Byte)
DoEvents
If wNum = 0 Then Call APIBeep(200, 60): Call Sleep(60): Call APIBeep(200, 80) 'X
If wNum = 1 Then Call APIBeep(900, 60): Call Sleep(60): Call APIBeep(900, 80) '!
If wNum = 2 Then Call APIBeep(900, 60): Call Sleep(60): Call APIBeep(200, 80) 'i
DoEvents
End Sub

Public Function SpecialFolder(ByVal CSIDL As Long) As String
Dim R As Long
Dim sPath As String
Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260
R = SHGetSpecialFolderLocation(fPrg.hwnd, CSIDL, IDL)
If R = NOERROR Then
  sPath = Space$(MAX_LENGTH)
  R = SHGetPathFromIDList(ByVal IDL.mkid.cB, ByVal sPath)
  If R Then
    SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
  End If
End If
End Function

Public Sub AddToAutorun()
Dim EXpath As String
Dim w As Object, s As String
Dim Link As Object
EXpath = App.Path + "\" + App.EXEName + ".exe"
Set w = CreateObject("WScript.Shell")
s = SpecialFolder(7) + "\lightbar.lnk"
If IsObject(w) Then
  Set Link = w.CreateShortcut(s)
  If IsObject(Link) Then
    Link.Description = "LightBar"
    Link.IconLocation = EXpath
    Link.TargetPath = EXpath
    Link.WindowStyle = 0
    Link.WorkingDirectory = EXpath
    Link.save
  End If
End If
End Sub

Public Function GetShortcutPath(ByVal wPth As String) As String
Dim w As Object, s As Object
GetShortcutPath = wPth
wPth = UCase(wPth)
If Right(wPth, 4) = ".LNK" Or Right(wPth, 4) = ".URL" Then
  Set w = CreateObject("WScript.Shell")
  Set s = w.CreateShortcut(wPth)
  GetShortcutPath = s.TargetPath
  Set s = Nothing
  Set w = Nothing
End If
End Function

Public Sub TrayMgr(ByRef wAct As Byte, Optional ByRef wStr As String = "")
Static Stts As Byte
If wAct = 1 And Stts = 0 And ShowInTray = 1 Then 'sozdanie ikonki v tree
  With NID
    .cbSize = Len(NID)
    .hwnd = fPrg.hwnd
    .uId = vbNull
    .uFlags = &H2 Or &H4 Or &H1
    .uCallBackMessage = &H200
    .hIcon = fPrg.iTray.Picture
    .szTip = "LightBar v." & App.Major & "." & App.Minor & " (" & App.Revision & ")" & Chr(0)
  End With
  Call Shell_NotifyIcon(&H0, NID)
  Stts = 1
End If
If wAct = 0 And Stts = 1 Then 'udalenie ikonki v tree
  Call Shell_NotifyIcon(&H2, NID)
  Stts = 0
End If
End Sub

Public Sub SetCur(ByVal wObj As Long, Optional ByRef wHwnd As Byte = 1)
Dim RC As RECT
Dim lL As Integer, tT As Integer

If NotAutoFocus > 0 Then Exit Sub

If wHwnd = 1 Then
  GetWindowRect wObj, RC
  lL = (RC.Left + RC.Right) / 2
  tT = (RC.Top + RC.Bottom) / 2
Else
  RC = BttCoord(wObj)
  lL = RC.Left + (RC.Right / 2) + (fPrg.Left / 15)
  tT = RC.Top + (RC.Bottom / 2) + (fPrg.Top / 15)
End If
Call SetCursorPos(lL, tT)
End Sub

























