Attribute VB_Name = "mDpl"
'################################################'
'# Programm:                           LightBar #'
'# Part:                     Settings Generator #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             04.05.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

'           1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20         '
'                                                                                                 '
'        +---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+        '
'        |cal|ntp|pai|wrd|wm |mov|snd|vol|cmd|tsk|   |sht|reb|log|   |cd |   |kos|pau|pin|        '
'      1 |c  |d  |nt |pad|plr|mak|rec|   |   |mng|   |dwn|oot|off|   |rom|   |ink|k  |bol|        '
'        +---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+        '
'        |win|win|7  |ner|la |   |   |   |   |   |   |   |   |   |   |dat|   |sap|sol|che|        '
'      2 |amp|rar|zip|o  |   |   |   |   |   |   |   |   |   |   |   |tim|   |er |itr|rvy|        '
'        +---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+        '
'        |   |   |   |   |   |   |moi|c:\|prg|win|tem|   |clo|tra|upp|   |   |   |   |   |        '
'      3 |   |   |   |   |   |   |doc|   |fls|dws|p  |   |se |nsp|wnd|   |   |   |   |   |        '
'        +---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+        '
'        |ie |out|hyp|dia|   |   |   |   |   |   |   |   |   |   |   |   |wa |wa |   |yan|        '
'      4 |   |exp|ter|ler|   |   |   |   |   |   |   |   |   |   |   |   |bck|nxt|   |ex |        '
'        +---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+        '
'        |ope|fir|dm |idm|re |dsk|sky|psi|icq|qip|mir|   |jek|sis|Ust|   |vol|vol|   |goo|        '
'      5 |ra |fox|   |   |get|cal|pe |   |   |   |and|   |ran|tem|UdP|   |dwn|up |   |gle|        '
'        +---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+---+        '

Public Sub CreateDefaultIni()
Dim I As Integer
Dim NN As Integer
Dim SS As String

Call mPrg.LoadColors

bttCol = 20
bttRow = 5
bttS = 1
icoW = 16
icoH = 16
icoS = 0
FormLeft = 50 * 15
FormTop = 0
TransForm = 225
fPrg.tZdr.Interval = 0
fPrg.tShow.Tag = 10
HotMod = 68
HotKey = 81
FormNotHide = 0
FormNotTop = 0
FormNotHotKey = 0
DrawHotKey = 0
MBttW = 11
MBttH = 11
TimeNotShow = 0
ShowInTray = 0
BttToShow = 0
FntName = "MS Sans Serif"
FntSize = 8
FntBold = 1
FntItalic = 0
FntTop = -2

'######## STANDARTNYE PROGRAMMY ###################################################################'

NN = 101
If CreateDefaultIniAdd(NN, "%systemroot%\system32\calc.exe", "Calculator", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%systemroot%\system32\notepad.exe", "Notepad", 2) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%systemroot%\system32\mspaint.exe", "Paint", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Windows NT\Accessories\wordpad.exe", "WordPad", 2) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Windows Media Player\wmplayer.exe", "Windows Media Player", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Movie Maker\moviemk.exe", "Windows Movie Maker", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%systemroot%\system32\sndrec32.exe", "Sound recorder", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%systemroot%\system32\sndvol32.exe", "Volume", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%systemroot%\system32\cmd.exe", "Command prompt", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%systemroot%\system32\taskmgr.exe", "Task manager", 0) = 1 Then NN = NN + 1

'######## NESTANDARTNYE PROGRAMMY #################################################################'

NN = 201
If CreateDefaultIniAdd(NN, "%programfiles%\winamp\winamp.exe", "WinAmp", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\winrar\winrar.exe", "WinRAR", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\7-Zip\7zFM.exe", "7-zip", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Ahead\Nero StartSmart\NeroStartSmart.exe", "NeroStartSmart", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Light Alloy\LA.exe", "Light Alloy", 0) = 1 Then NN = NN + 1

'######## STANDARTNYE SETEVYE PROGRAMMY ###########################################################'

NN = 401
If CreateDefaultIniAdd(NN, "%programfiles%\Internet Explorer\IEXPLORE.EXE", "Internet Explorer", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Outlook Express\msimn.exe", "Outlook Express", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Windows NT\hypertrm.exe", "HyperTerminal", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Windows NT\dialer.exe", "Phone", 0) = 1 Then NN = NN + 1

'######## NESTANDARTNYE SETEVYE PROGRAMMY #########################################################'

NN = 501
If CreateDefaultIniAdd(NN, "%programfiles%\opera\opera.exe", "Opera", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Mozilla Firefox\firefox.exe", "Mozilla Firefox", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Download Master\dmaster.exe", "Download Master", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Internet Download Manager\IDMan.exe", "Internet Download Manager", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\ReGetDx\regetdx.exe", "ReGet Delux", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\DeskCall NG\DeskCallNG.exe", "DeskCall NG", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Skype\Phone\Skype.exe", "Skype", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Psi\psi.exe", "Psi", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\ICQ\Icq.exe", "ICQ", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\QIP\qip.exe", "QIP", 0) = 1 Then NN = NN + 1
If CreateDefaultIniAdd(NN, "%programfiles%\Miranda IM\miranda32.exe", "Miranda IM", 0) = 1 Then NN = NN + 1

'######## PAPKI ###################################################################################'

'307 308 309 310 papki

NN = 307: MapB(NN).wOpr = 3
MapB(NN).wFil = "%documents%": MapB(NN).wCap = "My documents"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 127

NN = 308: MapB(NN).wOpr = 3
MapB(NN).wFil = "C:\": MapB(NN).wCap = "C:"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 9

NN = 309: MapB(NN).wOpr = 3
MapB(NN).wFil = "%programfiles%": MapB(NN).wCap = "Programm Files"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 20

NN = 310: MapB(NN).wOpr = 3
MapB(NN).wFil = "%systemroot%": MapB(NN).wCap = "Windows"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 111

NN = 311: MapB(NN).wOpr = 3
MapB(NN).wFil = "%temp%": MapB(NN).wCap = "Temp"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 33

'######## PANEL' UPRAVLENIYA ######################################################################'

NN = 513: MapB(NN).wOpr = 3
MapB(NN).wFil = "Rundll32.exe": MapB(NN).wPrm = "shell32.dll,Control_RunDLL Desk.cpl"
MapB(NN).wCap = "Control panel -> Screen"
MapB(NN).wIFl = "%systemroot%\system32\Desk.cpl": MapB(NN).wINm = 1

NN = 514: MapB(NN).wOpr = 3
MapB(NN).wFil = "Rundll32.exe": MapB(NN).wPrm = "shell32.dll,Control_RunDLL Sysdm.cpl"
MapB(NN).wCap = "Control panel -> System"
MapB(NN).wIFl = "%systemroot%\system32\Sysdm.cpl": MapB(NN).wINm = 1

NN = 515: MapB(NN).wOpr = 3
MapB(NN).wFil = "Rundll32.exe": MapB(NN).wPrm = "shell32.dll,Control_RunDLL AppWiz.cpl"
MapB(NN).wCap = "Control panel -> Add/remove programms"
MapB(NN).wIFl = "%systemroot%\system32\AppWiz.cpl": MapB(NN).wINm = 1

'######## STANDARTNYE IGRY ########################################################################'

NN = 120
If CreateDefaultIniAdd(NN, "%programfiles%\Windows NT\Pinball\PINBALL.EXE", "Pinball", 0) = 1 Then NN = NN - 1
If CreateDefaultIniAdd(NN, "%SystemRoot%\system32\spider.exe", "Spider", 0) = 1 Then NN = NN - 1
If CreateDefaultIniAdd(NN, "%SystemRoot%\system32\sol.exe", "Sol", 0) = 1 Then NN = NN - 1

NN = 220
If CreateDefaultIniAdd(NN, "%SystemRoot%\system32\mshearts.exe", "Hearts", 0) = 1 Then NN = NN - 1
If CreateDefaultIniAdd(NN, "%SystemRoot%\system32\freecell.exe", "Freecell", 0) = 1 Then NN = NN - 1
If CreateDefaultIniAdd(NN, "%SystemRoot%\system32\winmine.exe", "Winmine", 0) = 1 Then NN = NN - 1

'######## ZAVERSHENIE RABOTY ######################################################################'

NN = 112 'Завершение работы
MapB(NN).wOpr = 4
MapB(NN).wFil = "lbar_shutdown": MapB(NN).wPrm = "shutdown"
MapB(NN).wCap = "Shut down"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 1

NN = 113 'Перезагрузка компьютера
MapB(NN).wOpr = 4
MapB(NN).wFil = "lbar_shutdown": MapB(NN).wPrm = "reboot"
MapB(NN).wCap = "Reboot"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 2

NN = 114 'Выход из системы
MapB(NN).wOpr = 4
MapB(NN).wFil = "lbar_shutdown": MapB(NN).wPrm = "logoff"
MapB(NN).wCap = "Log off"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 14

'######## INTERNET SSYLKI #########################################################################'

NN = 420: MapB(NN).wOpr = 3
MapB(NN).wFil = "http://www.yandex.ru": MapB(NN).wCap = "www.yandex.ru"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 16

NN = 520: MapB(NN).wOpr = 3
MapB(NN).wFil = "http://www.google.com": MapB(NN).wCap = "www.google.com"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 15

'######## DOPOLNITEL'NYE KOMMANDY #################################################################'
'uznayom bukvu pervogo cdroma
SS = "Z"
For I = 25 To 0 Step -1
  If GetDriveType(Chr(I + 65) & ":") = 5 Then
    SS = Chr(I + 65)
    Exit For
  End If
Next
NN = 116: MapB(NN).wOpr = 3
MapB(NN).wFil = "lbar_cdrom": MapB(NN).wPrm = SS: MapB(NN).wDir = "open/close"
MapB(NN).wCap = "Open/Close СD-ROM " & SS & ":"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 27

NN = 216: MapB(NN).wOpr = 1
MapB(NN).wFil = "lbar_other": MapB(NN).wPrm = "datetime": MapB(NN).wDir = "dd.mm.yyyy hh:nn:ss (mmm, dddd)"
MapB(NN).wCap = "Show date/time"
MapB(NN).wIFl = "%SYSTEMROOT%\SYSTEM32\Timedate.cpl": MapB(NN).wINm = 1

NN = 313: MapB(NN).wOpr = 3
MapB(NN).wFil = "lbar_window": MapB(NN).wPrm = "close": MapB(NN).wDir = "close"
MapB(NN).wCap = "Close window"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 132

NN = 314: MapB(NN).wOpr = 3
MapB(NN).wFil = "lbar_window": MapB(NN).wPrm = "transparent": MapB(NN).wDir = "204"
MapB(NN).wCap = "Transparent window"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 3

NN = 315: MapB(NN).wOpr = 3
MapB(NN).wFil = "lbar_window": MapB(NN).wPrm = "topmost": MapB(NN).wDir = "top"
MapB(NN).wCap = "Window topmost"
MapB(NN).wIFl = "%systemroot%\system32\shell32.dll": MapB(NN).wINm = 99

NN = 517: MapB(NN).wOpr = 2
MapB(NN).wFil = "lbar_sound": MapB(NN).wPrm = "down": MapB(NN).wDir = "2"
MapB(NN).wCap = "Volume -"
MapB(NN).wIFl = "%systemroot%\system32\Mmsys.cpl": MapB(NN).wINm = 37

NN = 518: MapB(NN).wOpr = 2
MapB(NN).wFil = "lbar_sound": MapB(NN).wPrm = "up": MapB(NN).wDir = "2"
MapB(NN).wCap = "Volume +"
MapB(NN).wIFl = "%systemroot%\system32\Mmsys.cpl": MapB(NN).wINm = 37

NN = 417: MapB(NN).wOpr = 1
MapB(NN).wFil = "lbar_winamp": MapB(NN).wPrm = "back": MapB(NN).wDir = "run"
MapB(NN).wCap = "WA -> Back"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 3

NN = 418: MapB(NN).wOpr = 1
MapB(NN).wFil = "lbar_winamp": MapB(NN).wPrm = "next": MapB(NN).wDir = "run"
MapB(NN).wCap = "WA -> Next"
MapB(NN).wIFl = "%PROGPATH%\icons.icl": MapB(NN).wINm = 7

Call SaveStt
End Sub

Private Function CreateDefaultIniAdd(ByRef wNN As Integer, ByRef wSS As String, ByRef wSN As String, ByRef wwShw As Byte) As Byte
If Dir(GetDir(wSS)) <> "" Then
  MapB(wNN).wOpr = 3
  MapB(wNN).wFil = wSS: MapB(wNN).wPrm = "": MapB(wNN).wDir = GetFolder(wSS)
  MapB(wNN).wShw = wwShw: MapB(wNN).wCap = wSN: MapB(wNN).wIFl = wSS: MapB(wNN).wINm = 1
  CreateDefaultIniAdd = 1
End If
End Function


































