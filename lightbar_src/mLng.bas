Attribute VB_Name = "mLng"
'################################################'
'# Programm:                           LightBar #'
'# Part:                        Language loader #'
'# Author:                               WFSoft #'
'# Email:                             wfs@of.kz #'
'# Website:                   lightbar.narod.ru #'
'# Date:                             08.07.2007 #'
'# License:                             GNU/GPL #'
'################################################'

Option Explicit
Public MapCmd(99, 99) As String
Public MapMsg(66) As String
Public MapOth(33) As String
Public MapGen(2) As String

Public Sub LoadLang(ByRef wFil As String, Optional ByRef wPrt As String)
If wPrt = "" Then Call GenLang

If Trim(wFil) = "" Then Exit Sub
If Dir(App.Path & "\" & wFil, vbNormal) = "" Then Exit Sub

Dim I As Integer
Dim FF As Long
Dim SS As String

Dim RR As Long 'poziciya znaka '='

Dim MapS() As String
Dim MapSS() As String

Dim wBlc As String

FF = FreeFile
Open App.Path & "\" & wFil For Input As #FF
  Do
    If EOF(FF) = True Then Exit Do
    Line Input #FF, SS
    
    If SS <> "" Then
      If Left(Trim(SS), 1) <> "#" Then
        If Left(Trim(SS), 1) = "[" And Right(Trim(SS), 1) = "]" Then
          wBlc = Trim(SS)
        End If
        
        RR = InStr(SS, "=")
        If RR > 0 Then
          
          MapS = Split(SS, "=")
          MapS(0) = Trim(MapS(0))
          For I = 2 To UBound(MapS) Step 1
            MapS(1) = MapS(1) & "=" & MapS(I)
          Next I
          MapS(1) = Trim(MapS(1))
          
          MapS(1) = Replace(MapS(1), "\n", vbCrLf)
          
          If wBlc = "[gen]" And (wPrt = "" Or wPrt = "gen") Then
            If MapS(0) = "nam" Then MapGen(0) = MapS(1)
            If MapS(0) = "ver" Then MapGen(1) = MapS(1)
            If MapS(0) = "aut" Then MapGen(2) = MapS(1)
          End If
          
          If wBlc = "[abt]" And (wPrt = "" Or wPrt = "abt") Then
            If MapS(0) = "abt" Then fAbt.frAbout.Caption = MapS(1)
            If MapS(0) = "abt" Then fAbt.frAbout.Caption = MapS(1)
            If MapS(0) = "pth" Then fAbt.cPth.Caption = MapS(1)
            If MapS(0) = "htk" Then fAbt.cHtK.Caption = MapS(1)
            If MapS(0) = "bck" Then fAbt.cOK.Caption = MapS(1)
            If MapS(0) = "eml" Then fAbt.infEml.Caption = MapS(1)
            If MapS(0) = "sit" Then fAbt.infSit.Caption = MapS(1)
            If MapS(0) = "inf" Then fAbt.infInf.Caption = MapS(1)
            If MapS(0) = "wrt" Then fAbt.infWrt.Caption = MapS(1) & " Visual Basic 6.0"
          End If
          
          If wBlc = "[cmd]" And (wPrt = "" Or wPrt = "cmd") Then
            If MapS(0) = "cap" Then fCmd.Caption = " " & MapS(1)
            
            If MapS(0) = "k01" Then MapCmd(1, 0) = MapS(1)
            If MapS(0) = "k02" Then MapCmd(2, 0) = MapS(1)
            If MapS(0) = "k03" Then MapCmd(3, 0) = MapS(1)
            If MapS(0) = "k04" Then MapCmd(4, 0) = MapS(1)
            If MapS(0) = "k05" Then MapCmd(5, 0) = MapS(1)
            If MapS(0) = "k06" Then MapCmd(6, 0) = MapS(1)
            If MapS(0) = "k99" Then MapCmd(99, 0) = MapS(1)
            
            If MapS(0) = "e0101" Then MapCmd(1, 1) = MapS(1)
            If MapS(0) = "e0102" Then MapCmd(1, 2) = MapS(1)
            If MapS(0) = "e0103" Then MapCmd(1, 3) = MapS(1)
            If MapS(0) = "e0201" Then MapCmd(2, 1) = MapS(1)
            If MapS(0) = "e0202" Then MapCmd(2, 2) = MapS(1)
            If MapS(0) = "e0203" Then MapCmd(2, 3) = MapS(1)
            If MapS(0) = "e0301" Then MapCmd(3, 1) = MapS(1)
            If MapS(0) = "e0302" Then MapCmd(3, 2) = MapS(1)
            If MapS(0) = "e0303" Then MapCmd(3, 3) = MapS(1)
            If MapS(0) = "e0304" Then MapCmd(3, 4) = MapS(1)
            If MapS(0) = "e0305" Then MapCmd(3, 5) = MapS(1)
            If MapS(0) = "e0306" Then MapCmd(3, 6) = MapS(1)
            If MapS(0) = "e0307" Then MapCmd(3, 7) = MapS(1)
            If MapS(0) = "e0401" Then MapCmd(4, 1) = MapS(1)
            If MapS(0) = "e0402" Then MapCmd(4, 2) = MapS(1)
            If MapS(0) = "e0501" Then MapCmd(5, 1) = MapS(1)
            If MapS(0) = "e0502" Then MapCmd(5, 2) = MapS(1)
            If MapS(0) = "e0503" Then MapCmd(5, 3) = MapS(1)
            If MapS(0) = "e0504" Then MapCmd(5, 4) = MapS(1)
            If MapS(0) = "e0505" Then MapCmd(5, 5) = MapS(1)
            If MapS(0) = "e0506" Then MapCmd(5, 6) = MapS(1)
            If MapS(0) = "e0507" Then MapCmd(5, 7) = MapS(1)
            If MapS(0) = "e0508" Then MapCmd(5, 8) = MapS(1)
            If MapS(0) = "e0509" Then MapCmd(5, 9) = MapS(1)
            If MapS(0) = "e0510" Then MapCmd(5, 10) = MapS(1)
            If MapS(0) = "e0511" Then MapCmd(5, 11) = MapS(1)
            If MapS(0) = "e0601" Then MapCmd(6, 1) = MapS(1)
            If MapS(0) = "e9901" Then MapCmd(99, 1) = MapS(1)
            If MapS(0) = "e9902" Then MapCmd(99, 2) = MapS(1)
            If MapS(0) = "e9903" Then MapCmd(99, 3) = MapS(1)
            If MapS(0) = "e9904" Then MapCmd(99, 4) = MapS(1)
            
            If MapS(0) = "d00cap" Then fCmd.frDpl(0).Caption = MapS(1)
            If MapS(0) = "d01cap" Then fCmd.frDpl(1).Caption = MapS(1)
            If MapS(0) = "d01dev" Then fCmd.infDev.Caption = MapS(1)
            If MapS(0) = "d01stp" Then fCmd.infStp.Caption = MapS(1)
            If MapS(0) = "d02cap" Then fCmd.frDpl(2).Caption = MapS(1)
            If MapS(0) = "d02run" Then fCmd.chRun.Caption = MapS(1)
            If MapS(0) = "d03cap" Then fCmd.frDpl(3).Caption = MapS(1)
            If MapS(0) = "d03frc" Then fCmd.chForce.Caption = MapS(1)
            If MapS(0) = "d04cap" Then fCmd.frDpl(4).Caption = MapS(1)
            If MapS(0) = "d04ab1" Then fCmd.infAb1.Caption = MapS(1)
            If MapS(0) = "d04ab2" Then fCmd.infAb2.Caption = MapS(1)
            If MapS(0) = "d05cap" Then fCmd.frDpl(5).Caption = MapS(1)
            If MapS(0) = "d06cap" Then fCmd.frDpl(6).Caption = MapS(1)
            
            If MapS(0) = "bck" Then fCmd.cOK.Caption = MapS(1)
            If MapS(0) = "cnc" Then fCmd.cCancel.Caption = MapS(1)
          End If
          
          If wBlc = "[edt]" And (wPrt = "" Or wPrt = "edt") Then
            If MapS(0) = "cap" Then fEdt.wCap = " " & MapS(1)
            If MapS(0) = "act" Then fEdt.frAction.Caption = MapS(1)
            If MapS(0) = "ac1" Then fEdt.oOpr(0).Caption = MapS(1)
            If MapS(0) = "ac2" Then fEdt.oOpr(1).Caption = MapS(1)
            If MapS(0) = "ac3" Then fEdt.oOpr(2).Caption = MapS(1)
            If MapS(0) = "ac4" Then fEdt.oOpr(3).Caption = MapS(1)
            If MapS(0) = "ac5" Then fEdt.oOpr(4).Caption = MapS(1)
            If MapS(0) = "cmd" Then fEdt.frCommand.Caption = MapS(1)
            If MapS(0) = "cmm" Then fEdt.lCmm.Caption = MapS(1)
            If MapS(0) = "prm" Then fEdt.lPrm.Caption = MapS(1): fEdt.lPrm2.Caption = MapS(1) & "1": fEdt.lDir2.Caption = MapS(1) & "2"
            If MapS(0) = "dir" Then fEdt.lDir.Caption = MapS(1)
            If MapS(0) = "ico" Then fEdt.frIco.Caption = MapS(1)
            If MapS(0) = "icf" Then fEdt.infIcoFil.Caption = MapS(1)
            If MapS(0) = "icn" Then fEdt.infIcoNum.Caption = MapS(1)
            If MapS(0) = "htk" Then fEdt.frHotKey.Caption = MapS(1)
            If MapS(0) = "blc" Then fEdt.chHtN.Caption = MapS(1)
            If MapS(0) = "mod" Then fEdt.infMod.Caption = MapS(1)
            If MapS(0) = "key" Then fEdt.infKey.Caption = MapS(1)
            If MapS(0) = "oth" Then fEdt.frOther.Caption = MapS(1)
            If MapS(0) = "inf" Then fEdt.infInf.Caption = MapS(1)
            If MapS(0) = "stl" Then fEdt.infStl.Caption = MapS(1)
            If MapS(0) = "trk" Then fEdt.frTrk.Caption = MapS(1)
            If MapS(0) = "tri" Then fEdt.infTrk.Caption = MapS(1)
            If MapS(0) = "mov" Then fEdt.frMove.Caption = MapS(1)
            If MapS(0) = "mvi" Then fEdt.infMov.Caption = MapS(1)
            If MapS(0) = "chn" Then fEdt.frChange.Caption = MapS(1)
            If MapS(0) = "rp1" Then fEdt.cReplacePaths(0).Caption = MapS(1)
            If MapS(0) = "rp2" Then fEdt.cReplacePaths(1).Caption = MapS(1)
            If MapS(0) = "cop" Then fEdt.cCopy.Caption = MapS(1)
            If MapS(0) = "pst" Then fEdt.cPast.Caption = MapS(1)
            If MapS(0) = "cnc" Then fEdt.cCancel.Caption = MapS(1)
            If MapS(0) = "bck" Then fEdt.cOK.Caption = MapS(1)
            If MapS(0) = "clr" Then fEdt.cClear.Caption = MapS(1)
            If MapS(0) = "app" Then fEdt.cApply.Caption = MapS(1)
          End If
          
          If wBlc = "[key]" And (wPrt = "" Or wPrt = "key") Then
            If MapS(0) = "cap" Then fKey.Caption = " " & MapS(1)
            If MapS(0) = "inf" Then fKey.infInf.Caption = MapS(1)
          End If
          
          If wBlc = "[msg]" And (wPrt = "" Or wPrt = "msg") Then
            If MapS(0) = "cap" Then fMsg.Caption = " " & MapS(1)
          End If
          
          If wBlc = "[stt]" And (wPrt = "" Or wPrt = "stt") Then
            If MapS(0) = "cap" Then fStt.Caption = " " & MapS(1)
            If MapS(0) = "pg3" Then fStt.frPag(3).Caption = MapS(1): fStt.oPag(3).Caption = MapS(1)
            If MapS(0) = "pg0" Then fStt.frPag(0).Caption = MapS(1): fStt.oPag(0).Caption = MapS(1)
            If MapS(0) = "pg1" Then fStt.frPag(1).Caption = MapS(1): fStt.oPag(1).Caption = MapS(1)
            If MapS(0) = "pg2" Then fStt.frPag(2).Caption = MapS(1): fStt.oPag(2).Caption = MapS(1)
            
            If MapS(0) = "ast" Then fStt.chAutoStart.Caption = MapS(1)
            If MapS(0) = "sit" Then fStt.chShowInTray.Caption = MapS(1)
            If MapS(0) = "naf" Then fStt.chNotAutoFocus.Caption = MapS(1)
            If MapS(0) = "ncm" Then fStt.chNotClearMem.Caption = MapS(1)
            If MapS(0) = "htk" Then fStt.infHtK.Caption = MapS(1)
            If MapS(0) = "lng" Then fStt.chLng.Caption = MapS(1)
            If MapS(0) = "gnl" Then fStt.cGenLng.Caption = MapS(1)
            If MapS(0) = "frs" Then fStt.infStt.Caption = MapS(1)
            If MapS(0) = "st1" Then fStt.oStt(0).Caption = MapS(1)
            If MapS(0) = "st2" Then fStt.oStt(1).Caption = MapS(1)
            If MapS(0) = "st3" Then fStt.oStt(2).Caption = MapS(1)
            If MapS(0) = "st4" Then fStt.oStt(3).Caption = MapS(1)
            
            If MapS(0) = "btt" Then fStt.chScreenBottom.Caption = MapS(1)
            If MapS(0) = "tim" Then fStt.chTimeNotShow.Caption = MapS(1)
            If MapS(0) = "pkm" Then fStt.chPKM.Caption = MapS(1)
            If MapS(0) = "pol" Then fStt.infPol.Caption = MapS(1)
            If MapS(0) = "vrt" Then fStt.infVrt.Caption = MapS(1)
            If MapS(0) = "vri" Then fStt.infVrI.Caption = MapS(1)
            If MapS(0) = "trn" Then fStt.infTrn.Caption = MapS(1)
            If MapS(0) = "zdr" Then fStt.infZdr.Caption = MapS(1)
            If MapS(0) = "ffh" Then fStt.chFormFullHide.Caption = MapS(1)
            If MapS(0) = "anm" Then fStt.infAnm.Caption = MapS(1)
            If MapS(0) = "fnt" Then fStt.infFnt.Caption = MapS(1)
            If MapS(0) = "ots" Then fStt.infOts.Caption = MapS(1)
            
            If MapS(0) = "col" Then fStt.infCol.Caption = MapS(1)
            If MapS(0) = "row" Then fStt.infRow.Caption = MapS(1)
            If MapS(0) = "icw" Then fStt.infIcW.Caption = MapS(1)
            If MapS(0) = "ich" Then fStt.infIcH.Caption = MapS(1)
            If MapS(0) = "bts" Then fStt.infBtS.Caption = MapS(1)
            If MapS(0) = "ics" Then fStt.infIcS.Caption = MapS(1)
            If MapS(0) = "mbw" Then fStt.infMBW.Caption = MapS(1)
            If MapS(0) = "mbh" Then fStt.infMBH.Caption = MapS(1)
            If MapS(0) = "dhk" Then fStt.chDrawHK.Caption = MapS(1)
            
            If MapS(0) = "cl0" Then fStt.infCl0.Caption = MapS(1)
            If MapS(0) = "cl1" Then fStt.infCl1.Caption = MapS(1)
            If MapS(0) = "clh" Then fStt.infClH.Caption = MapS(1)
            If MapS(0) = "dep" Then fStt.infDep.Caption = MapS(1)
            If MapS(0) = "slb" Then fStt.chSlB.Caption = MapS(1)
            If MapS(0) = "gen" Then fStt.cGenColors.Caption = MapS(1)
            
            If MapS(0) = "cnc" Then fStt.cCancel.Caption = MapS(1)
            If MapS(0) = "bck" Then fStt.cOK.Caption = MapS(1)
            If MapS(0) = "dfl" Then fStt.cDefault.Caption = MapS(1)
            If MapS(0) = "app" Then fStt.cApply.Caption = MapS(1)
          End If
          
          If wBlc = "[messages]" And (wPrt = "" Or wPrt = "messages") Then
            If Val(MapS(0)) >= 0 And Val(MapS(0)) <= 99 Then MapMsg(Val(MapS(0))) = MapS(1)
          End If
          
          If wBlc = "[other]" And (wPrt = "" Or wPrt = "other") Then
            If Val(MapS(0)) >= 0 And Val(MapS(0)) <= 99 Then MapOth(Val(MapS(0))) = MapS(1)
          End If
        End If
      End If
    End If
  Loop
Close #FF
End Sub

Private Sub GenLang()
MapCmd(1, 0) = "�������"
MapCmd(2, 0) = "CD-ROM"
MapCmd(3, 0) = "����"
MapCmd(4, 0) = "����"
MapCmd(5, 0) = "WinAmp"
MapCmd(6, 0) = "����"
MapCmd(99, 0) = "������"

MapCmd(1, 1) = "���������� ������"
MapCmd(1, 2) = "������������"
MapCmd(1, 3) = "����� �� �������"
MapCmd(2, 1) = "������� ������ �D"
MapCmd(2, 2) = "������� ������ �D"
MapCmd(2, 3) = "�������/������� ������ �D"
MapCmd(3, 1) = "������� ���� (Close)"
MapCmd(3, 2) = "������� ���� (Quit)"
MapCmd(3, 3) = "���������� ������ ���� ����"
MapCmd(3, 4) = "���������� ���� ���� ����"
MapCmd(3, 5) = "����������/�������� ���� ������������"
MapCmd(3, 6) = "���������� / ������������"
MapCmd(3, 7) = "�������� / ������������"
MapCmd(4, 1) = "��������� ��������� �����"
MapCmd(4, 2) = "��������� ��������� �����"
MapCmd(5, 1) = "�����"
MapCmd(5, 2) = "����"
MapCmd(5, 3) = "�����"
MapCmd(5, 4) = "����"
MapCmd(5, 5) = "�����"
MapCmd(5, 6) = "��������� �����"
MapCmd(5, 7) = "�������"
MapCmd(5, 8) = "��������� ���������"
MapCmd(5, 9) = "��������� ���������"
MapCmd(5, 10) = "����� �� 5 ������"
MapCmd(5, 11) = "����� �� 5 ������"
MapCmd(6, 1) = "����������� � ..."
MapCmd(99, 1) = "�������� ���� �����"
MapCmd(99, 2) = "�������� ���� �����"
MapCmd(99, 3) = "�������� ClipBoard"
MapCmd(99, 4) = "���������� ���������� USB ����������"

MapMsg(0) = "��� ������� �� ��� ������ ��������� � ���������� ��������� ���� � ������������ � ��������� �������� ���������."
MapMsg(1) = "��� �������� ���� ����� ���������."
MapMsg(2) = "�������� ������������ �����?"
MapMsg(3) = "���� ���������� � ����� ������."
MapMsg(4) = "����������� �� �������."
MapMsg(5) = "������������ ��� ������ LightBar ��� � ������������ � �������. ������ ������ ����� � ����� ""������������"" �������� ������������, � ������ � ������� �������."
MapMsg(6) = "������ �� ���� ������������ ������." & vbCrLf & "����� ���������� ������, ������������� ���������"
MapMsg(7) = "�������������� ��������� �������� �����." & vbCrLf & "�������� ""�������"" ���������� ""����������"" ������."
MapMsg(8) = "�������� ������ LightBar � ����� ""������������"" ��� �������� ������������."
MapMsg(9) = "������������ �� ������� ������ ������ ""������� �������"" ����������� �� ���� ����� (���������� ��� ���������� ����� ������ ������� ��� ������������, ��������: ��� ���������� ���������� Ctrl+Alt+W ������������ ������ ""W"")."
MapMsg(10) = "�� ������� �������������� �������� ������ � ����� � ����������."
MapMsg(11) = "������ ��� ���� ����� ���� ��������� ������������, ���� �������� �� ���� ����� ������� ����, � � ���� ����������, ��� ����� ��������������� ������ ��� ������ ������ ������� ����."
MapMsg(12) = "����������� �������� ���� ��������� ����� ������. ����, ��������, ������ ������ ��������� ������ ����� windows."
MapMsg(13) = "��������� ������ � ��������� ���� (��� ������)." & vbCrLf & "���� ������� �� ������ ����� ������� ����, �� ������� ���� ��������� ����� ��������������� � �������������. � ���� ������ �������, �� ������� ���� ����� ����������/������������ �� ������."
MapMsg(14) = "� ���� ����������, ��� ������ ������� ������� ����, ���� ��������� ����� ���������� �� ����� ������."
MapMsg(15) = "����������/�������� ���� � ������� ���� ���������."
MapMsg(16) = "��������� ""�������"" ������� �� �������� � ������� �������� ���� ���������."
MapMsg(17) = "������� ""�������"" �������."
MapMsg(18) = "���� ���� ��������� �������� �����."
MapMsg(19) = "�������� ����� ���� �� 10 �� 90"
MapMsg(20) = "����� ����� ���� �� 1 �� 20"
MapMsg(21) = "�������� �������� ����� ����� ���� �� 0 �� 5000 ��."
MapMsg(22) = "������ ������ ����� ���� �� 4 �� 64"
MapMsg(23) = "������ ������ ����� ���� �� 4 �� 64"
MapMsg(24) = "������ ����� �������� ����� ���� �� 0 �� 50"
MapMsg(25) = "���������� ������ �� ���� ������ ����� ���� �� 0 �� 20"
MapMsg(26) = "������ ������ ���� ����� ���� �� 9 �� 34"
MapMsg(27) = "������ ������ ���� ����� ���� �� 11 �� 34"
MapMsg(28) = "������ ������ ����� ���� �� 2 �� 72"
MapMsg(29) = "��� ������� �� ��� ������ ��������� � ���������� ��������� ���� � ����������� ����� � �� ����������."
MapMsg(30) = "���� �������, �����������, ���������, ������� ... ? ����������� �� �� e-Mail wfs@of.kz � ����� LightBar (���������� � ������������ ���� �������). ���������� �������� ����."
MapMsg(31) = "�� lightbar.narod.ru ������ ����� ������ ������ ���������."
MapMsg(32) = "�������������� (������ � LightBar) �������. (��� �� �� �������? �������� ��� ������ �� wfs@of.kz, ���-������ ���������.)"
MapMsg(33) = "����������� �������� ������ (�� � ����� ������)."
MapMsg(34) = "���������� ��������� �������� ������, � �������� ��� �������������� ������ ����� �� ��������."
MapMsg(35) = "���������� ��������� �������� ������, � �������� ��� �������������� ������ ������ �� ��������."
MapMsg(36) = "���������� ��������� �������� ������, � �������� ��� �������������� ������ ������ �� ��������."
MapMsg(37) = "���������� ��������� �������� ������, � �������� ��� �������������� ������ ����� �� ��������."
MapMsg(38) = "���� �� ��������� ���� ������� ��� ����������� �����-���� �������� (� ������ ���������) �� ��� �� ��������� ����� ������������� LightBar'��. ���� �� ������ ����� LightBar �������� �������� ������ � ������������� ����� - �������� ���� ������. (�� ������ ������ ������ LightBar'� ��� �� ����������������, ���� �� ���������� ���� � ���� ������� �� 2,3,... ������, �� ��� �������, ��� ��� ���������.)"
MapMsg(39) = "������ ��� ������� ������� ������� ��� ������� ����� ������."
MapMsg(40) = "������ ��� �������� ������� ������� � ������."
MapMsg(41) = "���������� ��������� �������� ������, � ����������� ��� �� ���� ��� � ���� (������ �������� �������)."
MapMsg(42) = "���������� ��������� �������� ������, � ����������� ��� �� ���� ��� � ����� (������ �������� �������)."
MapMsg(43) = "���������� ��������� �������� ������, � ����������� ��� �� ���� ��� � ���� (������ �������� �������)."
MapMsg(44) = "���������� ��������� �������� ������, � ����������� ��� �� ���� ��� � ��� (������ �������� �������)."
MapMsg(45) = "������� �������������� ������ (���� ����� �� ���������� �������, �� ����������� ������ �����)."
MapMsg(46) = "�������� � ���� ���� � ����������� �������."
MapMsg(47) = "���������� ��� ""��� ��������"", ������������ ������ �������� ����� ��� (� ������� �� ������ �������), � ""������� ����������"" �� ������� � ������� ������� (�� ���� ����� 2 ���� ������� ���� � ����� �����, ���������� ������� ����� �������� �� �����)."
MapMsg(48) = "�� ���� ����� ���� � Winamp'�."
MapMsg(49) = "�� ���� ��������� Winamp. ������� �"
MapMsg(50) = "� ���� ""�������"" ���� �����. ��������� ""��������� �������"" - ""��� ���������� �������""?"
MapMsg(51) = "�� ������ ���� � ����� ��������." 'ne ispol'zuetsya
MapMsg(52) = "��� ��������� ����� �������� ������ ��� ������������� �������� ��� �������� �������� ����. ��� ���� ��� �� ������������� � ������ �������, � ���������� � ������ ���������. ���������� ����� ��� ����������� ������� �� ����� ���� (� ����������� �� �������� ������� ��� ������) ������."
MapMsg(53) = "������ ����� � ������ �� ������������� ��� ����������." & vbCrLf & "���������� ���� � ����� ""�������:"", ""������� �����:"" � ""���� ������:"""
MapMsg(54) = "��������� �������� lng ����� (�� ����� ������ � ����� � ���������� � ������ ""newlng.lng"")."
MapMsg(55) = "������ ������ ������ ����� ���� �� -50 �� 50"
MapMsg(56) = "�� ������� ���."

MapOth(0) = "�����"
MapOth(1) = "���������"
MapOth(2) = "� ���������"
MapOth(3) = "��������� ����"
MapOth(4) = "��������� ����"
MapOth(5) = "�� ������ ���� ����"
MapOth(6) = "������ ���� ����"
MapOth(7) = "�� ������ ������� �������"
MapOth(8) = "������ ������� �������"
MapOth(9) = "������"
MapOth(10) = "������ ����������� ������� ������ � ��������� (����� ���������, ������ ������� ���� ����)."
MapOth(11) = "������������� ���������� �� ������ ������������ ��� ���������� ����� (����� ���������, ������ ������� ���� ����)"
MapOth(12) = "����� �����"
MapOth(13) = "����� �����"
MapOth(14) = "����� � ��������"
MapOth(15) = "����� ����� �� ��������"
MapOth(16) = "����� �����"
MapOth(17) = "�������:"
MapOth(18) = "����������� ������� �����������"
MapOth(19) = "������ ��� ����������� �"
MapOth(20) = "�������� �������� ����"
MapOth(21) = "��������� (���������� �����):"
MapOth(22) = "��������������:"
MapOth(23) = "(���������)"
MapOth(24) = "�������� ����� ��������"
MapOth(25) = "����� ����� ��������"

fEdt.wCap = "��������� ������ �"
End Sub

Public Sub GenLng()
Dim FF As Long
Dim I As Integer
FF = FreeFile
Open App.Path & "\newlng.lng" For Output As #FF
  Print #FF, "[gen]"
  Print #FF, "nam=Language"
  Print #FF, "ver=" & App.Major & "." & App.Minor
  Print #FF, "aut=Author"
  Print #FF, ""
  Print #FF, "[abt]"
  Print #FF, "abt" & "=" & Replace(fAbt.frAbout.Caption, vbCrLf, "\n")
  Print #FF, "pth" & "=" & Replace(fAbt.cPth.Caption, vbCrLf, "\n")
  Print #FF, "htk" & "=" & Replace(fAbt.cHtK.Caption, vbCrLf, "\n")
  Print #FF, "bck" & "=" & Replace(fAbt.cOK.Caption, vbCrLf, "\n")
  Print #FF, "eml" & "=" & Replace(fAbt.infEml.Caption, vbCrLf, "\n")
  Print #FF, "sit" & "=" & Replace(fAbt.infSit.Caption, vbCrLf, "\n")
  Print #FF, "inf" & "=" & Replace(fAbt.infInf.Caption, vbCrLf, "\n")
  Print #FF, "wrt" & "=" & Replace(Left$(fAbt.infWrt.Caption, Len(fAbt.infWrt.Caption) - 17), vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "[cmd]"
  Print #FF, "cap" & "=" & Replace(fCmd.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "k01" & "=" & Replace(MapCmd(1, 0), vbCrLf, "\n")
  Print #FF, "k02" & "=" & Replace(MapCmd(2, 0), vbCrLf, "\n")
  Print #FF, "k03" & "=" & Replace(MapCmd(3, 0), vbCrLf, "\n")
  Print #FF, "k04" & "=" & Replace(MapCmd(4, 0), vbCrLf, "\n")
  Print #FF, "k05" & "=" & Replace(MapCmd(5, 0), vbCrLf, "\n")
  Print #FF, "k06" & "=" & Replace(MapCmd(6, 0), vbCrLf, "\n")
  Print #FF, "k99" & "=" & Replace(MapCmd(99, 0), vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "e0101" & "=" & Replace(MapCmd(1, 1), vbCrLf, "\n")
  Print #FF, "e0102" & "=" & Replace(MapCmd(1, 2), vbCrLf, "\n")
  Print #FF, "e0103" & "=" & Replace(MapCmd(1, 3), vbCrLf, "\n")
  Print #FF, "e0201" & "=" & Replace(MapCmd(2, 1), vbCrLf, "\n")
  Print #FF, "e0202" & "=" & Replace(MapCmd(2, 2), vbCrLf, "\n")
  Print #FF, "e0203" & "=" & Replace(MapCmd(2, 3), vbCrLf, "\n")
  Print #FF, "e0301" & "=" & Replace(MapCmd(3, 1), vbCrLf, "\n")
  Print #FF, "e0302" & "=" & Replace(MapCmd(3, 2), vbCrLf, "\n")
  Print #FF, "e0303" & "=" & Replace(MapCmd(3, 3), vbCrLf, "\n")
  Print #FF, "e0304" & "=" & Replace(MapCmd(3, 4), vbCrLf, "\n")
  Print #FF, "e0305" & "=" & Replace(MapCmd(3, 5), vbCrLf, "\n")
  Print #FF, "e0306" & "=" & Replace(MapCmd(3, 6), vbCrLf, "\n")
  Print #FF, "e0307" & "=" & Replace(MapCmd(3, 7), vbCrLf, "\n")
  Print #FF, "e0401" & "=" & Replace(MapCmd(4, 1), vbCrLf, "\n")
  Print #FF, "e0402" & "=" & Replace(MapCmd(4, 2), vbCrLf, "\n")
  Print #FF, "e0501" & "=" & Replace(MapCmd(5, 1), vbCrLf, "\n")
  Print #FF, "e0502" & "=" & Replace(MapCmd(5, 2), vbCrLf, "\n")
  Print #FF, "e0503" & "=" & Replace(MapCmd(5, 3), vbCrLf, "\n")
  Print #FF, "e0504" & "=" & Replace(MapCmd(5, 4), vbCrLf, "\n")
  Print #FF, "e0505" & "=" & Replace(MapCmd(5, 5), vbCrLf, "\n")
  Print #FF, "e0506" & "=" & Replace(MapCmd(5, 6), vbCrLf, "\n")
  Print #FF, "e0507" & "=" & Replace(MapCmd(5, 7), vbCrLf, "\n")
  Print #FF, "e0508" & "=" & Replace(MapCmd(5, 8), vbCrLf, "\n")
  Print #FF, "e0509" & "=" & Replace(MapCmd(5, 9), vbCrLf, "\n")
  Print #FF, "e0510" & "=" & Replace(MapCmd(5, 10), vbCrLf, "\n")
  Print #FF, "e0511" & "=" & Replace(MapCmd(5, 11), vbCrLf, "\n")
  Print #FF, "e0601" & "=" & Replace(MapCmd(6, 1), vbCrLf, "\n")
  Print #FF, "e9901" & "=" & Replace(MapCmd(99, 1), vbCrLf, "\n")
  Print #FF, "e9902" & "=" & Replace(MapCmd(99, 2), vbCrLf, "\n")
  Print #FF, "e9903" & "=" & Replace(MapCmd(99, 3), vbCrLf, "\n")
  Print #FF, "e9904" & "=" & Replace(MapCmd(99, 4), vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "d00cap" & "=" & Replace(fCmd.frDpl(0).Caption, vbCrLf, "\n")
  Print #FF, "d01cap" & "=" & Replace(fCmd.frDpl(1).Caption, vbCrLf, "\n")
  Print #FF, "d01dev" & "=" & Replace(fCmd.infDev.Caption, vbCrLf, "\n")
  Print #FF, "d01stp" & "=" & Replace(fCmd.infStp.Caption, vbCrLf, "\n")
  Print #FF, "d02cap" & "=" & Replace(fCmd.frDpl(2).Caption, vbCrLf, "\n")
  Print #FF, "d02run" & "=" & Replace(fCmd.chRun.Caption, vbCrLf, "\n")
  Print #FF, "d03cap" & "=" & Replace(fCmd.frDpl(3).Caption, vbCrLf, "\n")
  Print #FF, "d03frc" & "=" & Replace(fCmd.chForce.Caption, vbCrLf, "\n")
  Print #FF, "d04cap" & "=" & Replace(fCmd.frDpl(4).Caption, vbCrLf, "\n")
  Print #FF, "d04ab1" & "=" & Replace(fCmd.infAb1.Caption, vbCrLf, "\n")
  Print #FF, "d04ab2" & "=" & Replace(fCmd.infAb2.Caption, vbCrLf, "\n")
  Print #FF, "d05cap" & "=" & Replace(fCmd.frDpl(5).Caption, vbCrLf, "\n")
  Print #FF, "d06cap" & "=" & Replace(fCmd.frDpl(6).Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "bck" & "=" & Replace(fCmd.cOK.Caption, vbCrLf, "\n")
  Print #FF, "cnc" & "=" & Replace(fCmd.cCancel.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "[edt]"
  Print #FF, "cap" & "=" & Replace(fEdt.wCap, vbCrLf, "\n")
  Print #FF, "act" & "=" & Replace(fEdt.frAction.Caption, vbCrLf, "\n")
  Print #FF, "ac1" & "=" & Replace(fEdt.oOpr(0).Caption, vbCrLf, "\n")
  Print #FF, "ac2" & "=" & Replace(fEdt.oOpr(1).Caption, vbCrLf, "\n")
  Print #FF, "ac3" & "=" & Replace(fEdt.oOpr(2).Caption, vbCrLf, "\n")
  Print #FF, "ac4" & "=" & Replace(fEdt.oOpr(3).Caption, vbCrLf, "\n")
  Print #FF, "ac5" & "=" & Replace(fEdt.oOpr(4).Caption, vbCrLf, "\n")
  Print #FF, "cmd" & "=" & Replace(fEdt.frCommand.Caption, vbCrLf, "\n")
  Print #FF, "cmm" & "=" & Replace(fEdt.lCmm.Caption, vbCrLf, "\n")
  Print #FF, "prm" & "=" & Replace(fEdt.lPrm.Caption, vbCrLf, "\n")
  Print #FF, "dir" & "=" & Replace(fEdt.lDir.Caption, vbCrLf, "\n")
  Print #FF, "ico" & "=" & Replace(fEdt.frIco.Caption, vbCrLf, "\n")
  Print #FF, "icf" & "=" & Replace(fEdt.infIcoFil.Caption, vbCrLf, "\n")
  Print #FF, "icn" & "=" & Replace(fEdt.infIcoNum.Caption, vbCrLf, "\n")
  Print #FF, "htk" & "=" & Replace(fEdt.frHotKey.Caption, vbCrLf, "\n")
  Print #FF, "blc" & "=" & Replace(fEdt.chHtN.Caption, vbCrLf, "\n")
  Print #FF, "mod" & "=" & Replace(fEdt.infMod.Caption, vbCrLf, "\n")
  Print #FF, "key" & "=" & Replace(fEdt.infKey.Caption, vbCrLf, "\n")
  Print #FF, "oth" & "=" & Replace(fEdt.frOther.Caption, vbCrLf, "\n")
  Print #FF, "inf" & "=" & Replace(fEdt.infInf.Caption, vbCrLf, "\n")
  Print #FF, "stl" & "=" & Replace(fEdt.infStl.Caption, vbCrLf, "\n")
  Print #FF, "trk" & "=" & Replace(fEdt.frTrk.Caption, vbCrLf, "\n")
  Print #FF, "tri" & "=" & Replace(fEdt.infTrk.Caption, vbCrLf, "\n")
  Print #FF, "mov" & "=" & Replace(fEdt.frMove.Caption, vbCrLf, "\n")
  Print #FF, "mvi" & "=" & Replace(fEdt.infMov.Caption, vbCrLf, "\n")
  Print #FF, "chn" & "=" & Replace(fEdt.frChange.Caption, vbCrLf, "\n")
  Print #FF, "rp1" & "=" & Replace(fEdt.cReplacePaths(0).Caption, vbCrLf, "\n")
  Print #FF, "rp2" & "=" & Replace(fEdt.cReplacePaths(1).Caption, vbCrLf, "\n")
  Print #FF, "cop" & "=" & Replace(fEdt.cCopy.Caption, vbCrLf, "\n")
  Print #FF, "pst" & "=" & Replace(fEdt.cPast.Caption, vbCrLf, "\n")
  Print #FF, "cnc" & "=" & Replace(fEdt.cCancel.Caption, vbCrLf, "\n")
  Print #FF, "bck" & "=" & Replace(fEdt.cOK.Caption, vbCrLf, "\n")
  Print #FF, "clr" & "=" & Replace(fEdt.cClear.Caption, vbCrLf, "\n")
  Print #FF, "app" & "=" & Replace(fEdt.cApply.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "[key]"
  Print #FF, "cap" & "=" & Replace(fKey.Caption, vbCrLf, "\n")
  Print #FF, "inf" & "=" & Replace(fKey.infInf.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "[msg]"
  Print #FF, "cap" & "=" & Replace(fMsg.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "[stt]"
  Print #FF, "cap" & "=" & Replace(fStt.Caption, vbCrLf, "\n")
  Print #FF, "pg3" & "=" & Replace(fStt.frPag(3).Caption, vbCrLf, "\n")
  Print #FF, "pg0" & "=" & Replace(fStt.frPag(0).Caption, vbCrLf, "\n")
  Print #FF, "pg1" & "=" & Replace(fStt.frPag(1).Caption, vbCrLf, "\n")
  Print #FF, "pg2" & "=" & Replace(fStt.frPag(2).Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "ast" & "=" & Replace(fStt.chAutoStart.Caption, vbCrLf, "\n")
  Print #FF, "sit" & "=" & Replace(fStt.chShowInTray.Caption, vbCrLf, "\n")
  Print #FF, "naf" & "=" & Replace(fStt.chNotAutoFocus.Caption, vbCrLf, "\n")
  Print #FF, "ncm" & "=" & Replace(fStt.chNotClearMem.Caption, vbCrLf, "\n")
  Print #FF, "htk" & "=" & Replace(fStt.infHtK.Caption, vbCrLf, "\n")
  Print #FF, "lng" & "=" & Replace(fStt.chLng.Caption, vbCrLf, "\n")
  Print #FF, "gnl" & "=" & Replace(fStt.cGenLng.Caption, vbCrLf, "\n")
  Print #FF, "frs" & "=" & Replace(fStt.infStt.Caption, vbCrLf, "\n")
  Print #FF, "st1" & "=" & Replace(fStt.oStt(0).Caption, vbCrLf, "\n")
  Print #FF, "st2" & "=" & Replace(fStt.oStt(1).Caption, vbCrLf, "\n")
  Print #FF, "st3" & "=" & Replace(fStt.oStt(2).Caption, vbCrLf, "\n")
  Print #FF, "st4" & "=" & Replace(fStt.oStt(3).Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "btt" & "=" & Replace(fStt.chScreenBottom.Caption, vbCrLf, "\n")
  Print #FF, "tim" & "=" & Replace(fStt.chTimeNotShow.Caption, vbCrLf, "\n")
  Print #FF, "pkm" & "=" & Replace(fStt.chPKM.Caption, vbCrLf, "\n")
  Print #FF, "pol" & "=" & Replace(fStt.infPol.Caption, vbCrLf, "\n")
  Print #FF, "vrt" & "=" & Replace(fStt.infVrt.Caption, vbCrLf, "\n")
  Print #FF, "vri" & "=" & Replace(fStt.infVrI.Caption, vbCrLf, "\n")
  Print #FF, "trn" & "=" & Replace(fStt.infTrn.Caption, vbCrLf, "\n")
  Print #FF, "zdr" & "=" & Replace(fStt.infZdr.Caption, vbCrLf, "\n")
  Print #FF, "ffh" & "=" & Replace(fStt.chFormFullHide.Caption, vbCrLf, "\n")
  Print #FF, "anm" & "=" & Replace(fStt.infAnm.Caption, vbCrLf, "\n")
  Print #FF, "fnt" & "=" & Replace(fStt.infFnt.Caption, vbCrLf, "\n")
  Print #FF, "ots" & "=" & Replace(fStt.infOts.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "col" & "=" & Replace(fStt.infCol.Caption, vbCrLf, "\n")
  Print #FF, "row" & "=" & Replace(fStt.infRow.Caption, vbCrLf, "\n")
  Print #FF, "icw" & "=" & Replace(fStt.infIcW.Caption, vbCrLf, "\n")
  Print #FF, "ich" & "=" & Replace(fStt.infIcH.Caption, vbCrLf, "\n")
  Print #FF, "bts" & "=" & Replace(fStt.infBtS.Caption, vbCrLf, "\n")
  Print #FF, "ics" & "=" & Replace(fStt.infIcS.Caption, vbCrLf, "\n")
  Print #FF, "mbw" & "=" & Replace(fStt.infMBW.Caption, vbCrLf, "\n")
  Print #FF, "mbh" & "=" & Replace(fStt.infMBH.Caption, vbCrLf, "\n")
  Print #FF, "dhk" & "=" & Replace(fStt.chDrawHK.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "cl0" & "=" & Replace(fStt.infCl0.Caption, vbCrLf, "\n")
  Print #FF, "cl1" & "=" & Replace(fStt.infCl1.Caption, vbCrLf, "\n")
  Print #FF, "clh" & "=" & Replace(fStt.infClH.Caption, vbCrLf, "\n")
  Print #FF, "dep" & "=" & Replace(fStt.infDep.Caption, vbCrLf, "\n")
  Print #FF, "slb" & "=" & Replace(fStt.chSlB.Caption, vbCrLf, "\n")
  Print #FF, "gen" & "=" & Replace(fStt.cGenColors.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "cnc" & "=" & Replace(fStt.cCancel.Caption, vbCrLf, "\n")
  Print #FF, "bck" & "=" & Replace(fStt.cOK.Caption, vbCrLf, "\n")
  Print #FF, "dfl" & "=" & Replace(fStt.cDefault.Caption, vbCrLf, "\n")
  Print #FF, "app" & "=" & Replace(fStt.cApply.Caption, vbCrLf, "\n")
  Print #FF, ""
  Print #FF, "[messages]"
  For I = 0 To 55 Step 1
    If MapMsg(I) <> "" Then Print #FF, I & "=" & Replace(MapMsg(I), vbCrLf, "\n")
  Next I
  Print #FF, ""
  Print #FF, "[other]"
  For I = 0 To 33 Step 1
    If MapOth(I) <> "" Then Print #FF, I & "=" & Replace(MapOth(I), vbCrLf, "\n")
  Next I
Close #FF
End Sub






































