Attribute VB_Name = "mScs"
'################################################'
'# Programm:                           LightBar #'
'# Part:                            Hook Module #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             21.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public hKbdLL As Long

Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WH_KEYBOARD_LL = 13

Public Type KBDLLHOOKSTRUCT
  vkCode      As Long
  scanCode    As Long
  flags       As Long
  time        As Long
  dwExtraInfo As Long
End Type
Private P As KBDLLHOOKSTRUCT

Public wfKey As Byte

Private I As Integer, II As Integer
Private Btt As Integer

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If FormNotHotKey > 0 Then
  LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
  Exit Function
End If
If (nCode = HC_ACTION) Then
  If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
    CopyMemory P, ByVal lParam, Len(P)
    
    'get modifikators
    If P.vkCode = 160 Then
      If P.flags < 128 Then
        If GetNum(0, ModK) = 0 Then ModK = ModK + 1
      Else
        If GetNum(0, ModK) = 1 Then ModK = ModK - 1
      End If
    End If
    If P.vkCode = 161 Then
      If P.flags < 128 Then
        If GetNum(1, ModK) = 0 Then ModK = ModK + 2
      Else
        If GetNum(1, ModK) = 1 Then ModK = ModK - 2
      End If
    End If
    
    If P.vkCode = 162 Then
      If P.flags < 128 Then
        If GetNum(2, ModK) = 0 Then ModK = ModK + 4
      Else
        If GetNum(2, ModK) = 1 Then ModK = ModK - 4
      End If
    End If
    If P.vkCode = 163 Then
      If P.flags < 128 Then
        If GetNum(3, ModK) = 0 Then ModK = ModK + 8
      Else
        If GetNum(3, ModK) = 1 Then ModK = ModK - 8
      End If
    End If
    
    If P.vkCode = 91 Then
      If P.flags < 128 Then
        If GetNum(4, ModK) = 0 Then ModK = ModK + 16
      Else
        If GetNum(4, ModK) = 1 Then ModK = ModK - 16
      End If
    End If
    If P.vkCode = 92 Then
      If P.flags < 128 Then
        If GetNum(5, ModK) = 0 Then ModK = ModK + 32
      Else
        If GetNum(5, ModK) = 1 Then ModK = ModK - 32
      End If
    End If
    
    If P.vkCode = 164 Then
      If P.flags < 128 Then
        If GetNum(6, ModK) = 0 Then ModK = ModK + 64
      Else
        If GetNum(6, ModK) = 1 Then ModK = ModK - 64
      End If
    End If
    If P.vkCode = 165 Then
      If P.flags < 128 Then
        If GetNum(7, ModK) = 0 Then ModK = ModK + 128
      Else
        If GetNum(7, ModK) = 1 Then ModK = ModK - 128
      End If
    End If
    
    'esli forma otlova klavish otkryta
    If wfKey = 1 Then
      If (P.vkCode < 160 Or P.vkCode > 165) And _
         (P.vkCode < 91 Or P.vkCode > 92) And _
         P.vkCode <> 27 Then
'        'smotrim, ne naznachena li jeta klavisha na druguyu knopku
'        For I = 0 To bttCol - 1 Step 1
'          For II = 0 To bttRow - 1 Step 1
'            Btt = (II + 1) * 100 + I + 1
'            If MapB(Btt).wOpr > 0 Then
'              If MapB(Btt).wHtM = ModK And MapB(Btt).wHtK = P.vkCode Then 'esli naidena knopka s _
'                              takim je hotkeem to nevypolnyaya nikakih dejstvij vyhodim iz procedury
'                LowLevelKeyboardProc = -1
'                Exit Function
'              End If
'            End If
'          Next II
'        Next I
        RetMod = ModK
        RetKey = P.vkCode
        wfKey = 0: Unload fKey
        LowLevelKeyboardProc = -1
        Exit Function
      Else
        fKey.lMod(0).Enabled = GetNum(0, ModK)
        fKey.lMod(1).Enabled = GetNum(1, ModK)
        fKey.lMod(2).Enabled = GetNum(2, ModK)
        fKey.lMod(3).Enabled = GetNum(3, ModK)
        fKey.lMod(4).Enabled = GetNum(4, ModK)
        fKey.lMod(5).Enabled = GetNum(5, ModK)
        fKey.lMod(6).Enabled = GetNum(6, ModK)
        fKey.lMod(7).Enabled = GetNum(7, ModK)
        If P.vkCode = 27 Then
          wfKey = 0: Unload fKey
          LowLevelKeyboardProc = -1
        End If
      End If
    End If
    
    'podpravlyaem massiv sostoyaniya klavish
    If P.flags < 128 Then 'esli down
      If MapKS(P.vkCode) = 0 Then
        MapKS(P.vkCode) = 1
      Else
        MapKS(P.vkCode) = 2
      End If
    Else 'esli up
      MapKS(P.vkCode) = 3
    End If
    
    'proveryaem sovpadeniya v batonah
    For I = 0 To bttCol - 1 Step 1
      For II = 0 To bttRow - 1 Step 1
        Btt = (II + 1) * 100 + I + 1
        
        If MapB(Btt).wOpr > 0 Then
          If MapB(Btt).wHtM = ModK And MapB(Btt).wHtK = P.vkCode Then
            If MapKS(P.vkCode) = MapB(Btt).wOpr Then Call mPrg.GetStatus(1, CStr(Btt))
            Call fPrg.Run(Btt, MapKS(P.vkCode))
            If MapB(Btt).wHtN = 0 Then
              LowLevelKeyboardProc = -1
              If P.flags >= 128 Then MapKS(P.vkCode) = 0
              Exit Function
            End If
          End If
        End If
        
      Next II
    Next I
    
    'ecshyo nemnogo podpravlyaem massiv sostoyaniya klavish
    If P.flags >= 128 Then MapKS(P.vkCode) = 0
    
    'a ne dlya otkrytiya li formy, nazhataya klavischa?
    If Lck = 0 Then
      If HotMod = ModK And HotKey = P.vkCode Then
        If MapKS(P.vkCode) = 1 Then
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
          LowLevelKeyboardProc = -1
          Exit Function
        End If
      End If
    End If
    
  End If
End If
LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
End Function










































