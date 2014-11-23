Attribute VB_Name = "mVol"
'################################################'
'# Programm:                           LightBar #'
'# Part:                            Hook Module #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             21.04.2007 #'
'# License:                             GNU/GPL #'
'################################################'
'
'Big part is taken from programm "Audio Mixer":
'AUTHOR:
'Name:     Stuart Pennington.
'Location: England.
'Status:   Unemployed.
'PROJECT:
'Name:          Audio Mixer.
'Test Platform: Windows 98 SE.
'Processor:     Pentium II - 300MhZ.

Option Explicit

Type MIXERCAPS
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * 32
     fdwSupport As Long
     cDestinations As Long
End Type

Type MIXERCONTROL
     cbStruct As Long
     dwControlID As Long
     dwControlType As Long
     fdwControl As Long
     cMultipleItems As Long
     szShortName As String * 16
     szName As String * 64
     lMinimum As Long
     lMaximum As Long
     Reserved(10) As Long
End Type

Type MIXERCONTROLDETAILS
     cbStruct As Long
     dwControlID As Long
     cChannels As Long
     item As Long
     cbDetails As Long
     paDetails As Long
End Type

Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type

Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * 64
End Type

Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type

Type Target
     dwType As Long
     dwDeviceID As Long
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * 32
End Type

Type MIXERLINE
     cbStruct As Long
     dwDestination As Long
     dwSource As Long
     dwLineID As Long
     fdwLine As Long
     dwUser As Long
     dwComponentType As Long
     cChannels As Long
     cConnections As Long
     cControls As Long
     szShortName As String * 16
     szName As String * 64
     lpTarget As Target
End Type

Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type

Declare Function mixerClose& Lib "winmm.dll" (ByVal hmx&)
Declare Function mixerGetControlDetails& Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj&, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails&)
Declare Function mixerGetDevCaps& Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId&, pmxcaps As MIXERCAPS, ByVal cbmxcaps&)
Declare Function mixerGetID& Lib "winmm.dll" (ByVal hmxobj&, pumxID&, ByVal fdwId&)
Declare Function mixerGetLineControls& Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj&, pmxlc As MIXERLINECONTROLS, ByVal fdwControls&)
Declare Function mixerGetLineInfo& Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj&, pmxl As MIXERLINE, ByVal fdwInfo&)
Declare Function mixerGetNumDevs& Lib "winmm.dll" ()
Declare Function mixerMessage& Lib "winmm.dll" (ByVal hmx&, ByVal umsg&, ByVal dwParam1&, ByVal dwParam2&)
Declare Function mixerOpen& Lib "winmm.dll" (phmx&, ByVal uMxId&, ByVal dwCallback&, ByVal dwInstance&, ByVal fdwOpen&)
Declare Function mixerSetControlDetails& Lib "winmm.dll" (ByVal hmxobj&, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails&)

Public Const MM_MIXM_LINE_CHANGE = &H3D0
Public Const MM_MIXM_CONTROL_CHANGE = &H3D1

Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&

Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&

Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Public Const MIXER_GETLINEINFOF_LINEID = &H2&
Public Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
Public Const MIXER_GETLINEINFOF_SOURCE = &H1&
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&

Public Const MIXER_OBJECTF_AUX = &H50000000
Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_HMIDIIN = &HC0000000
Public Const MIXER_OBJECTF_HMIDIOUT = &HB0000000
Public Const MIXER_OBJECTF_HMIXER = &H80000000
Public Const MIXER_OBJECTF_HWAVEIN = &HA0000000
Public Const MIXER_OBJECTF_HWAVEOUT = &H90000000
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_MIXER = &H0&
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000

Public Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&

Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&

Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000

Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000

Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000

Public Const MIXERCONTROL_CONTROLTYPE_BASS = &H50030002
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = &H20010000
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = &H10010000
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = &H21010000
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = &H0&
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = &H30040000
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = &H50030004
Public Const MIXERCONTROL_CONTROLTYPE_FADER = &H50030000
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = &H20010004
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = &H60030000
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = &H61030000
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = &H71010001
Public Const MIXERCONTROL_CONTROLTYPE_MONO = &H20010003
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = &H71010000
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = &H20010002
Public Const MIXERCONTROL_CONTROLTYPE_MUX = &H70010001
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = &H20010001
Public Const MIXERCONTROL_CONTROLTYPE_PAN = &H40020001
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = &H10020001
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = &H30050000
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = &H40020002
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = &H30020000
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = &H10020000
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = &H70010000
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = &H40020000
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = &H20010005
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = &H50030003
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = &H30030000
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = &H10030000
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = &H50030001

Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = &H1&
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = &H5&
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST = &H8&
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = &H2&
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = &H3&
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = &H4&
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = &H6&
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = &H0&
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = &H8&
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = &H7&

Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = &H100A&
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = &H1009&
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = &H1005&
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = &H1001&
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST = &H100A&
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = &H1002&
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = &H1003&
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = &H1007&
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = &H1004&
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = &H1006&
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = &H1008&

Public Const MIXERLINE_LINEF_ACTIVE = &H1&
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Public Const MIXERLINE_LINEF_SOURCE = &H80000000

Public Const MIXERLINE_TARGETTYPE_AUX = 5
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1

Public Const MIXERR_BASE = 1024
Public Const MIXERR_INVALCONTROL = 1025
Public Const MIXERR_INVALLINE = 1024
Public Const MIXERR_INVALVALUE = 1026
Public Const MIXERR_LASTERROR = 1026

'########'

Public hMixer&         ' The Handle Of The Mixer.
Public MaxSources&     ' Number Of Output Sources Available.
Public ProductName$    ' Product Name Of The Mixer (Used In The Main Form's Caption).

Private Destinations&  ' Number Of Destination's That The Mixer Support's.

' Used For Aquiring Details About Any Given Mixer Control.
' Fader, Mute, PeakMeter...
Public MCD As MIXERCONTROLDETAILS

Private ML As MIXERLINE


' #########################################################################

' This Is A Type I've Created To Slim Down
' The Coding In The Main Form

Type MIXERSETTINGS
     MxrChannels As Long    ' Indicates Whether A Line Is Mono Or Stereo.
     MxrLeftVol As Long     ' Left Volume Value (Balance).
     MxrRightVol As Long    ' Right Volume Value (Balance).
     MxrVol As Long         ' Fader Volume.
     MxrVolID As Long       ' Fader Control ID.
     MxrMute As Long        ' Mute Status.
     MxrMuteID As Long      ' Mute Control ID.
     MxrPeakID As Long      ' Peak Meter ID.
End Type

' A Dynamic Array Of The Aformentioned Type.

Public MixerState() As MIXERSETTINGS

' #########################################################################


' Addition API Subs And Function's.

Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x1&, ByVal y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Declare Function DrawEdge& Lib "user32" (ByVal ahDc&, lpRect As RECT, ByVal nEdge&, ByVal nFlags&)
Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal x1&, ByVal y1&, ByVal x2&, ByVal y2&)

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr&, ByVal cb&)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr&, struct As Any, ByVal cb&)

Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags&, ByVal dwBytes&)
Declare Function GlobalFree& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)

' Maintainance String, App's Title.
Public Const Ttl = "Volume Control"

Private Sub StartVol()
    ' Need To Check Out The Following Else We Can't Run.
    If Not MixerPresent Then End
    If Not OpenMixer Then End
    If Not GetDeviceCapabilities Then End
    ' If We Got Here, Let's Get Some Mixer Info.
    GetMixerInfo
End Sub
Private Function MixerPresent() As Boolean

    Dim Msg$  ' For Error String.

    ' The "mixerGetNumDevs" API Will Let Us Know If There Is A Mixer Onboard.
    If mixerGetNumDevs() Then
       MixerPresent = True      ' Yes, We Have One.
    Else
       ' No Mixer. This App Is Useless.
       ' Inform The User And Terminate On Return.
       Msg = "Unable to detect a mixer."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Terminating..."
       MsgBox Msg, vbCritical, Ttl & " - Error"
    End If

End Function
Private Function OpenMixer() As Boolean

    Dim Msg$  ' For Error String.

    ' See If We Can Open The Mixer.
    ' If Successful, The Global "hMixer" Variable Will Contain It's Handle.
    If mixerOpen(hMixer, 0, 0, 0, 0) = 0 Then
       OpenMixer = True   ' Yes, We Opened The Mixer.
    Else
       ' Unable To Open The Mixer.
       ' Inform The User And Terminate On Return.
       Msg = "Unable to open mixer."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Terminating..."
       MsgBox Msg, vbCritical, Ttl & " - Error"
    End If

End Function
Private Function GetDeviceCapabilities() As Boolean

    Dim Msg$                   ' For Error String.
    Dim MxrCaps As MIXERCAPS   ' Mixer Capabilities Structure.

    ' Query The Mixer's Capabilitie's.
    If mixerGetDevCaps(0, MxrCaps, Len(MxrCaps)) = 0 Then
       ' Only Interested In The "Destinations" Value And "Product Name".
       ' Destinations Can Be Speakers, Wave In, Voice Recognition Etc...
       Destinations = MxrCaps.cDestinations - 1
       ' Tidy Up The Pruduct Name Ready For Displaying In The Main Form's Caption.
       ProductName = Left(MxrCaps.szPname, InStr(MxrCaps.szPname, vbNullChar) - 1)
       ' Return Success.
       GetDeviceCapabilities = True
    Else
       ' Unable To Aquire Mixer Capabilites.
       ' Inform The User And Terminate On Return.
       Msg = "Unable to aquire mixer capabilities."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Terminating..."
       MsgBox Msg, vbCritical, Ttl & " - Error"
    End If

End Function
Private Sub GetMixerInfo()

    ' Purpose: Scans The Destination's Until The Speaker's Are Found.
    '          Then, Information About All Sources Connected To The Speaker's
    '          Are Saved Into The "MixerState" Array For Use In The Main Form

    Dim Dst&, Src&    ' Destination And Source Counter's.
    Dim ControlID&    ' ID Of A Given Control.

    For Dst = 0 To Destinations
        ' Prep The MIXERLINE Structure.
        ML.cbStruct = Len(ML)
        ML.dwDestination = Dst
        ' Get Destination Line Info.
        mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_DESTINATION

        ' Was The Component Type The Speaker's?
        If ML.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS Then

           ' How Many Item's Are Connected To The Speaker's?
           ' I'm Gonna Set An Upper Limit Of 10 And Set The "MaxSources" Variable.
           If ML.cConnections > 10 Then
              ML.cConnections = 10
              MaxSources = 10
           Else
              MaxSources = ML.cConnections  ' Less Than 10.
           End If

           ' Re-Dimension The "MixerState" Array.
           ' Note: The Array Is Zero Based, Element Zero Is For The Master Voume
           '       The Remaining Elements Are For The Source's.
           ReDim MixerState(MaxSources)

           ' Save The Number Of Channels For The Master Volume.
           MixerState(0).MxrChannels = ML.cChannels
           ' Update The Name Label On The Main Form.
           FrmMxr.LblName(0).Caption = Left(ML.szName, InStr(ML.szName, vbNullChar) - 1)
 
           ' Call The "GetControlID" Function So We Can Get The Control ID
           ' Of The Master Volume.
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
           If ControlID <> 0 Then
              ' Prep The MCD Structure For The Master Volume Fader.
              With MCD
                  .cbDetails = 4  ' Size Of A Long In Byte's.
                  .cbStruct = 24
                  .cChannels = ML.cChannels
                  .dwControlID = ControlID
                  .item = 0
                  .paDetails = VarPtr(MixerState(0).MxrVol)
              End With
              ' Get The Master Volume Setting.
              mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
              ' Track Bar Logic Is The Reverse Of Fader's On A Hardware Mixer
              ' So Reverse The Value.
              MixerState(0).MxrVol = 65535 - MixerState(0).MxrVol
              ' Save The Master Volume Control ID.
              MixerState(0).MxrVolID = MCD.dwControlID
           Else
              ' Couldn't Get It, Disable The Fader.
              FrmMxr.SldrVol(0).Enabled = 0
           End If

           ' Call The "GetControlID" Function So We Can Get The Control ID
           ' Of The Master Mute.
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
           If ControlID <> 0 Then
              ' Prep The MCD Structure For The Master Mute.
              With MCD
                  .cbDetails = 4  ' Size Of A Long In Byte's.
                  .cbStruct = Len(MCD)
                  .cChannels = ML.cChannels
                  .dwControlID = ControlID
                  .item = 0
                  .paDetails = VarPtr(MixerState(0).MxrMute)
              End With
              ' Get The Master Mute Setting.
              mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
              ' Save The Master Mute Control ID.
              MixerState(0).MxrMuteID = MCD.dwControlID
           Else
              ' Couldn't Get It, Disable The Master Mute.
              FrmMxr.ChkMute(0).Enabled = 0
           End If

           ' Does This Control Have A Peak Meter With It?
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_PEAKMETER)
           If ControlID <> 0 Then
              ' It Does, Save It's ID.
              MixerState(0).MxrPeakID = ControlID
           End If

           ' Now That We've Found The Speakers And The Master Volume,
           ' Let's Get The Source's...

           For Src = 0 To ML.cConnections - 1
               ' Prep The "MIXERLINE" Struct For Source's.
               ML.cbStruct = Len(ML)
               ML.dwDestination = Dst
               ML.dwSource = Src
               ' Get The Line Info For The Current Source.
               mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_SOURCE

               ' Save The Channels Of The Source.
               MixerState(Src + 1).MxrChannels = ML.cChannels
               ' Update The Name Label On The Main Form.
               FrmMxr.LblName(Src + 1).Caption = Left(ML.szName, InStr(ML.szName, vbNullChar) - 1)

               ' Call The "GetControlID" Function So We Can Get The Control ID
               ' Of The Current Source Volume.
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
               If ControlID <> 0 Then
                  ' Prep The MCD Structure For The Current Source Volume.
                  With MCD
                      .cbDetails = 4   ' Size Of A Long In Byte's.
                      .cbStruct = Len(MCD)
                      .cChannels = ML.cChannels
                      .dwControlID = ControlID
                      .item = 0
                      .paDetails = VarPtr(MixerState(Src + 1).MxrVol)
                  End With
                  ' Get The Current Source Volume Setting.
                  mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                  ' Save The Volume Setting.
                  MixerState(Src + 1).MxrVol = 65535 - MixerState(Src + 1).MxrVol
                  ' Save The ID
                  MixerState(Src + 1).MxrVolID = MCD.dwControlID
               Else
                  ' Couldn't Get It, So Disable The Control.
                  FrmMxr.SldrVol(Src + 1).Enabled = 0
               End If

               ' Call The "GetControlID" Function So We Can Get The Control ID
               ' Of The Current Source Mute.
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
               If ControlID <> 0 Then
                  ' Prep The MCD Structure For The Current Source Mute.
                  With MCD
                      .cbDetails = 4   ' Size Of A Long In Byte's.
                      .cbStruct = Len(MCD)
                      .cChannels = ML.cChannels
                      .dwControlID = ControlID
                      .item = 0
                      .paDetails = VarPtr(MixerState(Src + 1).MxrMute)
                  End With
                  ' Get The Current Source Mute Setting.
                  mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                  ' Save The Mute Control ID.
                  MixerState(Src + 1).MxrMuteID = MCD.dwControlID
               Else
                  ' Couldn't Get It, So Disable The Control.
                  FrmMxr.ChkMute(Src + 1).Enabled = 0
               End If

               ' Does This Control Have A Peak Meter With It?
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_PEAKMETER)
               If ControlID <> 0 Then
                  ' It Does, Save It's Id.
                  MixerState(Src + 1).MxrPeakID = ControlID
               End If
           Next
           ' We Found The Destination That Is The Speaker's, So Exit The Outer Loop.
           Exit For
        End If
    Next

End Sub
Public Function GetControlID&(ByVal ComponentType&, ByVal ControlType&)

   ' Purpose: Return's The Requested Control ID.

   Dim hMem&
   Dim MC As MIXERCONTROL
   Dim MxrLine As MIXERLINE
   Dim MLC As MIXERLINECONTROLS

   ' Prep The MxrLine Structure.
   MxrLine.cbStruct = Len(MxrLine)
   MxrLine.dwComponentType = ComponentType  ' This Value Sent In.

   ' Get The Line Info.
   If mixerGetLineInfo(hMixer, MxrLine, MIXER_GETLINEINFOF_COMPONENTTYPE) = 0 Then
      ' Prep The MLC Structure.
      MLC.cbStruct = Len(MLC)
      MLC.dwLineID = ML.dwLineID
      MLC.dwControl = ControlType     ' This Value Sent In.
      MLC.cControls = 1
      MLC.cbmxctrl = Len(MC)

      hMem = GlobalAlloc(&H40, Len(MC))
      MLC.pamxctrl = GlobalLock(hMem)

      MC.cbStruct = Len(MC)

      ' Get The Line Control.
      If mixerGetLineControls(hMixer, MLC, MIXER_GETLINECONTROLSF_ONEBYTYPE) = 0 Then
         ' Copy The Data To The MC Structure.
         CopyStructFromPtr MC, MLC.pamxctrl, Len(MC)
         ' Return The Control ID.
         GetControlID = MC.dwControlID
      End If

      GlobalUnlock hMem
      GlobalFree hMem
   End If

End Function

