Attribute VB_Name = "mDlg"
Option Explicit
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'system icon index
Private Const SHGFI_LARGEICON = &H0 'large icon
Private Const SHGFI_SMALLICON = &H1 'small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Const BOLD_FONTTYPE = &H100
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_APPLY = &H200&
Private Const CF_SCREENFONTS = &H1
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_EFFECTS = &H100&
Private Const CF_PALETTE = 9
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type ChooseColor
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Private Type ChooseFont
  lStructSize As Long
  hwndOwner As Long
  hdc As Long
  lpLogFont As Long
  iPointSize As Long
  flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type
Private Type BrowseInfo
  hwndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Private Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private shinfo As SHFILEINFO
Public hwndOwner As Long
Public Filter As String
Public OpenDialogTitle As String
Public SaveDialogTitle As String
Public FolderDialogTitle As String
Public ShowDirsOnly As Boolean
Function ShowFont(fntDefault As StdFont) As StdFont
Dim lFlags As Long, lg As LOGFONT, cf As ChooseFont
Set ShowFont = New StdFont
lFlags = lFlags Or CF_SCREENFONTS
lFlags = (lFlags Or CF_INITTOLOGFONTSTRUCT) And Not (CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE)
'lFlags = lFlags Or CF_EFFECTS
lg.lfHeight = -(fntDefault.Size * ((1440 / 72) / Screen.TwipsPerPixelY))
lg.lfWeight = fntDefault.Weight
lg.lfItalic = fntDefault.Italic
lg.lfUnderline = fntDefault.Underline
lg.lfStrikeOut = fntDefault.Strikethrough
StrToBytes lg.lfFaceName, fntDefault.Name

cf.hInstance = App.hInstance
cf.hwndOwner = hwndOwner
cf.lpLogFont = VarPtr(lg)
cf.iPointSize = fntDefault.Size * 10
cf.flags = lFlags
'cf.rgbColors = nColor
cf.lStructSize = Len(cf)
If ChooseFont(cf) Then
  lFlags = cf.flags
  ShowFont.Bold = cf.nFontType And BOLD_FONTTYPE
  ShowFont.Italic = lg.lfItalic
  ShowFont.Strikethrough = lg.lfStrikeOut
  ShowFont.Underline = lg.lfUnderline
  ShowFont.Weight = lg.lfWeight
  ShowFont.Size = cf.iPointSize / 10
  ShowFont.Name = BytesToStr(lg.lfFaceName)
  'nColor = cf.rgbColors
End If
End Function
Function ShowColor() As Long
Dim cd As ChooseColor
cd.lStructSize = LenB(cd)
cd.hwndOwner = hwndOwner
cd.hInstance = App.hInstance
cd.lpCustColors = String(8 * 16, 0)
If ChooseColor(cd) Then
  ShowColor = cd.rgbResult
Else
  ShowColor = -1
End If
End Function
Function ShowFolder(Optional ByRef wInitDir As String = "") As String
Dim lRes As Long
Dim sTemp As String
Dim iPos As Integer
Dim bi As BrowseInfo
With bi
  .hwndOwner = hwndOwner
  .lpszTitle = lstrcat(FolderDialogTitle, "")
  .ulFlags = Abs(ShowDirsOnly)
End With
lRes = SHBrowseForFolder(bi)
If lRes Then
  sTemp = String(MAX_PATH, vbNullChar)
  SHGetPathFromIDList lRes, sTemp
  CoTaskMemFree lRes
  iPos = InStr(sTemp, vbNullChar)
  If iPos Then sTemp = Left(sTemp, iPos - 1)
End If
ShowFolder = sTemp
End Function
Function ShowOpen(Optional ByRef wInitDir As String = "") As String
Dim OFN As OPENFILENAME
Dim sFilter As String
Dim nRes As Long
sFilter = ConvertFilter(Filter)
OFN.lpstrInitialDir = wInitDir
OFN.hInstance = App.hInstance
OFN.hwndOwner = hwndOwner
OFN.lpstrFile = String(MAX_PATH, vbNullChar)
OFN.lpstrTitle = OpenDialogTitle
OFN.lpstrFilter = sFilter
OFN.nMaxFile = MAX_PATH
OFN.lStructSize = LenB(OFN)
nRes = GetOpenFileName(OFN)
If nRes Then ShowOpen = Trim(OFN.lpstrFile) Else ShowOpen = ""
End Function
Function ShowSave(Optional ByRef wInitDir As String = "") As String
Dim OFN As OPENFILENAME
Dim sFilter As String
Dim nRes As Long
sFilter = ConvertFilter(Filter)
OFN.lpstrInitialDir = wInitDir
OFN.hInstance = App.hInstance
OFN.hwndOwner = hwndOwner
OFN.lpstrFile = String(MAX_PATH, vbNullChar)
OFN.lpstrFilter = sFilter
OFN.lpstrTitle = SaveDialogTitle
OFN.nMaxFile = MAX_PATH
OFN.lStructSize = LenB(OFN)
nRes = GetSaveFileName(OFN)
If nRes Then ShowSave = Trim(OFN.lpstrFile) Else ShowSave = ""
End Function


'Misc Functions
Private Sub StrToBytes(ab() As Byte, s As String)
If GetCount(ab) < 0 Then
  ab = StrConv(s, vbFromUnicode)
Else
  Dim cab As Long
  cab = UBound(ab) - LBound(ab) + 1
  If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
  CopyMem ab(LBound(ab)), s, cab
End If
End Sub
Private Function BytesToStr(ab() As Byte) As String
BytesToStr = StrConv(ab, vbUnicode)
End Function
Private Function GetCount(arr) As Integer
On Error Resume Next
Dim nCount As Integer
nCount = UBound(arr)
If Err Then
  Err.Clear
  GetCount = -1
Else
  GetCount = nCount
End If
End Function
Private Function ConvertFilter(ByVal sFilter) As String
Dim sTemp As String
Dim I As Integer
sTemp = sFilter
For I = 1 To Len(sTemp)
  If Mid(sTemp, I, 1) = "|" Then Mid(sTemp, I, 1) = vbNullChar
Next I
ConvertFilter = sTemp
End Function
Function GetFileTitle(ByVal FileName As String) As String
Dim shinfo As SHFILEINFO
Dim sTemp As String
SHGetFileInfo FileName, 0, shinfo, LenB(shinfo), &H200
sTemp = shinfo.szDisplayName
If InStr(sTemp, vbNullChar) Then sTemp = Left(sTemp, InStr(sTemp, vbNullChar) - 1)
GetFileTitle = sTemp
End Function
Function GetFileIcon(ByVal FileName As String) As Long
Dim shinfo As SHFILEINFO
Dim hIcon As String
hIcon = SHGetFileInfo(FileName, 0&, shinfo, LenB(shinfo), SHGFI_SMALLICON)
GetFileIcon = hIcon
End Function
