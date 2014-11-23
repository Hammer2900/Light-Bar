Attribute VB_Name = "mReg"
'*********************************************************************
' REGSITRY.BAS - Contains the code necessary to access the Windows
'                registration datbase.
'*********************************************************************
Option Explicit
'*********************************************************************
' The minimal API calls required to read from and write to the
' registry.
'*********************************************************************
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias _
    "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal _
    dwReserved&, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal Reserved As Long, ByVal dwType As Long, lpData _
    As Any, ByVal cbData As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
    "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
    As Long, ByVal samDesired As Long, lpSecurityAttributes As _
    Long, phkResult As Long, lpdwDisposition As Long) _
    As Long

Private Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey&)
'*********************************************************************
' The constants used in this module for the registry API calls.
'*********************************************************************
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000

Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
        KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY _
        Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
        KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const REG_SZ = 1        ' Unicode null terminated string
Private Const ERROR_SUCCESS = 0
'*********************************************************************
' The numeric constants for the major keys in the registry.  These
' are made public so users can use these numeric constants externally.
'*********************************************************************
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

'*********************************************************************
' This module raises errors using this base value
'*********************************************************************
Public Const ERRBASE As Long = vbObjectError + 6000
Public Const REG_UNSUPPORTED As String = _
                            "<Format Not Supported>"
'*********************************************************************
' Additional API calls for advanced registry features. These features
' are included by default, but may be excluded by setting the
' LEAN_AND_MEAN conditional compilation argument = 1 in your project
' properties dialog.
'*********************************************************************
#If LEAN_AND_MEAN = 0 Then
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias _
    "RegEnumKeyExA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, _
    lpcbName As Long, ByVal lpReserved&, ByVal lpClass$, lpcbClass&, _
    lpftLastWriteTime As Any) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpValueName$ _
    , lpcbValueName&, ByVal lpReserved&, lpType&, lpData As Any, _
    lpcbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey&, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias _
    "RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, _
    ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, _
    lpcbMaxClassLen&, lpcValues&, lpcbMaxValueNameLen&, _
    lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime _
    As Any) As Long

'*********************************************************************
' Additional constants used by these registry API calls.
'*********************************************************************
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const REG_DWORD = 4
#End If

'*********************************************************************
' GetRegString takes three arguments. A HKEY constant (listed above),
' a subkey, and a value in that subkey. This function returns the
' string stored in the strValueName value in the registry.
'*********************************************************************
Public Function GetRegString(hKey As Long, strSubKey As String, _
                                strValueName As String) As String
    Dim strSetting As String
    Dim lngDataLen As Long
    Dim hSubKey As Long
    '*****************************************************************
    ' Open the key. If success, then get the data from the key.
    '*****************************************************************
    If RegOpenKeyEx(hKey, strSubKey, 0, KEY_ALL_ACCESS, hSubKey) = _
        ERROR_SUCCESS Then
        strSetting = Space(255)
        lngDataLen = Len(strSetting)
        '*************************************************************
        ' Query the key for the current setting. If this call
        ' succeeds, then return the string.
        '*************************************************************
        If RegQueryValueEx(hSubKey, strValueName, ByVal 0, _
            REG_SZ, ByVal strSetting, lngDataLen) = _
            ERROR_SUCCESS Then
            If lngDataLen > 1 Then
                GetRegString = Left(strSetting, lngDataLen - 1)
            End If
        Else
        End If
        '*************************************************************
        ' ALWAYS close any keys that you open.
        '*************************************************************
        RegCloseKey hSubKey
    End If
End Function
'*********************************************************************
' SetRegString takes four arguments. A HKEY constant (listed above),
' a subkey, a value in that subkey, and a setting for the key.
'*********************************************************************
Public Sub SetRegString(hKey As Long, strSubKey As String, _
                                strValueName As String, strSetting _
                                As String)
    Dim hNewHandle As Long
    Dim lpdwDisposition As Long
    '*****************************************************************
    ' Create & open the key. If success, then get then write the data
    ' to the key.
    '*****************************************************************
    RegCreateKey hKey, strSubKey, hNewHandle
    RegSetValueEx hNewHandle, strValueName, 0, REG_SZ, ByVal strSetting, Len(strSetting)
    '*****************************************************************
    ' ALWAYS close any keys that you open.
    '*****************************************************************
    RegCloseKey hNewHandle
End Sub
'*********************************************************************
' Extended registry functions begin here
'*********************************************************************
#If LEAN_AND_MEAN = 0 Then
'*********************************************************************
' Returns a DWORD value from a given regsitry key
'*********************************************************************
Public Function GetRegDWord(hKey&, strSubKey$, strValueName$) As Long
    Dim lngDataLen As Long
    Dim hSubKey As Long
    Dim lngRetVal As Long
    '*****************************************************************
    ' Open the key. If success, then get the data from the key.
    '*****************************************************************
    If RegOpenKeyEx(hKey, strSubKey, 0, KEY_ALL_ACCESS, hSubKey) = _
        ERROR_SUCCESS Then
        '*************************************************************
        ' Query the key for the current setting. If this call
        ' succeeds, then return the string.
        '*************************************************************
        lngDataLen = 4 'Bytes
        If RegQueryValueEx(hSubKey, strValueName, ByVal 0, _
            REG_DWORD, lngRetVal, lngDataLen) = ERROR_SUCCESS Then
            GetRegDWord = lngRetVal
        Else
        End If
        '*************************************************************
        ' ALWAYS close any keys that you open.
        '*************************************************************
        RegCloseKey hSubKey
    End If
End Function
'*********************************************************************
' Sets a registry key to a DWORD value
'*********************************************************************
Public Sub SetRegDWord(hKey&, strSubKey$, strValueName$, lngSetting&)
    Dim hNewHandle As Long
    Dim lpdwDisposition As Long
    '*****************************************************************
    ' Create & open the key. If success, then get then write the data
    ' to the key.
    '*****************************************************************
    RegCreateKey hKey, strSubKey, hNewHandle
    RegSetValueEx hNewHandle, strValueName, 0, REG_DWORD, lngSetting, 4
    '*****************************************************************
    ' ALWAYS close any keys that you open.
    '*****************************************************************
    RegCloseKey hNewHandle
End Sub
'*********************************************************************
' Returns an array of all of the registry keys in a given key
'*********************************************************************
Public Function GetRegKeys(hKey&, Optional strSubKey$) As Variant
    Dim hChildKey As Long
    Dim lngSubKeys As Long
    Dim lngMaxKeySize As Long
    Dim lngDataRetBytes As Long
    Dim i As Integer
    '*****************************************************************
    ' Create a string array to hold the return values
    '*****************************************************************
    Dim strRetArray() As String
    '*****************************************************************
    ' If strSubKey was provided, then open it...
    '*****************************************************************
    If Len(strSubKey) Then
        '*************************************************************
        ' Exit if you did not successfully open the child key
        '*************************************************************
        If RegOpenKeyEx(hKey, strSubKey, 0, KEY_ALL_ACCESS, _
            hChildKey) <> ERROR_SUCCESS Then
            'Err.Raise ERRBASE + 4, "GetRegKeys", "RegOpenKeyEx failed!"
            Exit Function
        End If
    '*****************************************************************
    ' Otherwise use the top level hKey handle
    '*****************************************************************
    Else
        hChildKey = hKey
    End If
    '*****************************************************************
    ' Find out the array and value sizes in advance
    '*****************************************************************
    If QueryRegInfoKey(hChildKey, lngSubKeys, lngMaxKeySize) _
        <> ERROR_SUCCESS Or lngSubKeys = 0 Then
        'Err.Raise ERRBASE + 5, "GetRegKeys", "RegQueryInfoKey failed!"
        If Len(strSubKey) Then RegCloseKey hChildKey
        Exit Function
    End If
    '*****************************************************************
    ' Resize the array to fit the return values
    '*****************************************************************
    lngSubKeys = lngSubKeys '- 1 ' Adjust to zero based
    ReDim strRetArray(lngSubKeys) As String
    '*****************************************************************
    ' Get all of the keys
    '*****************************************************************
    For i = 0 To lngSubKeys
        '*************************************************************
        ' Set the buffers to max key size returned from
        ' RegQueryInfoKey
        '*************************************************************
        lngDataRetBytes = lngMaxKeySize
        strRetArray(i) = Space(lngMaxKeySize)
        
        RegEnumKeyEx hChildKey, i, strRetArray(i), _
            lngDataRetBytes, 0&, vbNullString, ByVal 0&, ByVal 0&
        '*************************************************************
        ' Trim off trailing nulls
        '*************************************************************
        strRetArray(i) = Left(strRetArray(i), lngDataRetBytes)
    Next i
    '*****************************************************************
    ' ALWAYS close any key that you open (but NEVER close the top
    ' level keys!!!!)
    '*****************************************************************
    If Len(strSubKey) Then RegCloseKey hChildKey
    '*****************************************************************
    ' Return the string array with the results
    '*****************************************************************
    GetRegKeys = strRetArray
End Function
'*********************************************************************
' Returns a multi dimensional variant array of all the values and
' settings in a given registry subkey.
'*********************************************************************
Public Function GetRegKeyValues(hKey&, strSubKey$) As Variant
    Dim lngNumValues As Long      ' Number values in this key
    
    Dim strValues() As String     ' Value and return array
    Dim lngMaxValSize  As Long    ' Size of longest value
    Dim lngValRetBytes As Long    ' Size of current value
    
    Dim lngMaxSettingSize As Long ' Size of longest REG_SZ in this key
    Dim lngSetRetBytes As Long    ' Size of current REG_SZ
    
    Dim lngSetting As Long        ' Used for DWORD
            
    Dim lngType As Long           ' Type of value returned from
                                  ' RegEnumValue
    
    Dim hChildKey As Long         ' The handle of strSubKey
    Dim i As Integer              ' Loop counter
    '*****************************************************************
    ' Exit if you did not successfully open the child key
    '*****************************************************************
    If RegOpenKeyEx(hKey, strSubKey, 0, KEY_ALL_ACCESS, hChildKey) _
        <> ERROR_SUCCESS Then
        'Err.Raise ERRBASE + 4, "GetRegKeyValues", _
            "RegOpenKeyEx failed!"
        Exit Function
    End If
    '*****************************************************************
    ' Find out the array and value sizes in advance
    '*****************************************************************
    If QueryRegInfoKey(hChildKey, , , lngNumValues, lngMaxValSize, _
        lngMaxSettingSize) <> ERROR_SUCCESS Or lngNumValues = 0 Then
        'Err.Raise ERRBASE + 5, "GetRegKeyValues", _
            "RegQueryInfoKey failed!"
        RegCloseKey hChildKey
        Exit Function
    End If
    '*****************************************************************
    ' Resize the array to fit the return values
    '*****************************************************************
    lngNumValues = lngNumValues - 1 ' Adjust to zero based
    ReDim strValues(0 To lngNumValues, 0 To 1) As String
    '*****************************************************************
    ' Get all of the values and settings for the key
    '*****************************************************************
    For i = 0 To lngNumValues
        '*************************************************************
        ' Make the return buffers large enough to hold the results
        '*************************************************************
        strValues(i, 0) = Space(lngMaxValSize)
        lngValRetBytes = lngMaxValSize
        
        strValues(i, 1) = Space(lngMaxSettingSize)
        lngSetRetBytes = lngMaxSettingSize
        '*************************************************************
        ' Get a single value and setting from the registry
        '*************************************************************
        RegEnumValue hChildKey, i, strValues(i, 0), lngValRetBytes, _
            0, lngType, ByVal strValues(i, 1), lngSetRetBytes
        '*************************************************************
        ' If the return value was a string, then trim trailing nulls
        '*************************************************************
        If lngType = REG_SZ Then
            strValues(i, 1) = Left(strValues(i, 1), lngSetRetBytes - 1)
        '*************************************************************
        ' Else if it was a DWord, call RegEnumValue again to store
        ' the return setting in a long variable
        '*************************************************************
        ElseIf lngType = REG_DWORD Then
            '*********************************************************
            ' We already know the return size of the value because
            ' we got it in the last call to RegEnumValue, so we
            ' can tell RegEnumValue that its buffer size is the
            ' length of the string already returned, plus one (for
            ' the trailing null terminator)
            '*********************************************************
            lngValRetBytes = lngValRetBytes + 1
            '*********************************************************
            ' Make the call again using a long instead of string
            '*********************************************************
            RegEnumValue hChildKey, i, strValues(i, 0), _
                lngValRetBytes, 0, lngType, lngSetting, lngSetRetBytes
            '*********************************************************
            ' Return the long as a string
            '*********************************************************
            strValues(i, 1) = CStr(lngSetting)
        '*************************************************************
        ' Otherwise let the user know that this code doesn't support
        ' the format returned (such as REG_BINARY)
        '*************************************************************
        Else
            strValues(i, 1) = REG_UNSUPPORTED
        End If
        '*************************************************************
        ' Store the return value and setting in a multi dimensional
        ' array with the value in the 0 index and the setting in
        ' the 1 index of the second dimension.
        '*************************************************************
        strValues(i, 0) = RTrim(Left(strValues(i, 0), lngValRetBytes))
        strValues(i, 1) = RTrim(strValues(i, 1))
    Next i
    '*****************************************************************
    ' ALWAYS close any keys you open
    '*****************************************************************
    RegCloseKey hChildKey
    '*****************************************************************
    ' Return the result as an array of strings
    '*****************************************************************
    GetRegKeyValues = strValues
End Function
'*********************************************************************
' Removes a given key from the registry
'*********************************************************************
Public Sub DeleteRegKey(hKey&, strParentKey$, strKeyToDel$, _
    Optional blnConfirm As Boolean)
    '*****************************************************************
    ' Get a handle to the parent key
    '*****************************************************************
    Dim hParentKey As Long
    If RegOpenKeyEx(hKey, strParentKey, 0, KEY_ALL_ACCESS, _
        hParentKey) <> ERROR_SUCCESS Then
        'Err.Raise ERRBASE + 4, "DeleteRegValue", "RegOpenKeyEx failed!"
        Exit Sub
    End If
    '*****************************************************************
    ' If blnConfirm, then make sure the user wants to delete the key
    '*****************************************************************
    If blnConfirm Then
        Call fMsg.GetMsg(fPrg, 1, "Are you sure you want to delete " & strKeyToDel & " ?", 1)
        If RetMsg = 0 Then
            RegCloseKey hParentKey
            Exit Sub
        End If
    End If
    '*****************************************************************
    ' Delete the key then close the parent key
    '*****************************************************************
    RegDeleteKey hParentKey, strKeyToDel
    RegCloseKey hParentKey
End Sub '*********************************************************************
' Removes a given value from a given registry key
'*********************************************************************
Public Sub DeleteRegValue(hKey&, strSubKey$, strValToDel$, _
    Optional blnConfirm As Boolean)
    '*****************************************************************
    ' Get the handle to the subkey
    '*****************************************************************
    Dim hSubKey As Long
    If RegOpenKeyEx(hKey, strSubKey, 0, KEY_ALL_ACCESS, _
        hSubKey) <> ERROR_SUCCESS Then
        'Err.Raise ERRBASE + 4, "DeleteRegValue", "RegOpenKeyEx failed!"
        Exit Sub
    End If
    '*****************************************************************
    ' If blnConfirm, then make sure the user wants to delete the value
    '*****************************************************************
    If blnConfirm Then
        Call fMsg.GetMsg(fPrg, 1, "Are you sure you want to delete " & strValToDel & " ?", 1)
        If RetMsg = 0 Then
            RegCloseKey hSubKey
            Exit Sub
        End If
    End If
    '*****************************************************************
    ' Delete the value then close the subkey
    '*****************************************************************
    RegDeleteValue hSubKey, strValToDel
    RegCloseKey hSubKey
End Sub
'*********************************************************************
' Query the registry to find out information about the values about
' to be returned in subsequent calls.
'*********************************************************************
Private Function QueryRegInfoKey(hKey&, Optional lngSubKeys&, _
    Optional lngMaxKeyLen&, Optional lngValues&, Optional _
    lngMaxValNameLen&, Optional lngMaxValLen&)

    QueryRegInfoKey = RegQueryInfoKey(hKey, vbNullString, _
        ByVal 0&, 0&, lngSubKeys, lngMaxKeyLen, ByVal 0&, lngValues, _
        lngMaxValNameLen, lngMaxValLen, ByVal 0&, ByVal 0&)
    '*****************************************************************
    ' Increase these values to include room for the terminating null
    '*****************************************************************
    lngMaxKeyLen = lngMaxKeyLen + 1
    lngMaxValNameLen = lngMaxValNameLen + 1
    lngMaxValLen = lngMaxValLen + 1
End Function

'Advanced non-Microsoft procedures
' Removes a given key and all subkeys from the registry (useful for Windows NT)
Public Sub DeleteRegKeyEx(hKey&, strParentKey$, strKeyToDel$)
    On Error GoTo Handler:
    Dim SubKeys() As String 'Array for subkeys
    Dim i As Long 'Counter
    'Enum all subkeys
    SubKeys() = GetRegKeys(hKey&, strParentKey$ + "\" + strKeyToDel$)
    'If given key has subkeys then remove them all
    If UBound(SubKeys()) <> 0 Then
        For i = 0 To UBound(SubKeys()) - 1
            DeleteRegKeyEx hKey&, strParentKey$ + "\" + strKeyToDel$, SubKeys(i)
        Next
    End If
    DeleteRegKey hKey&, strParentKey$, strKeyToDel$
    Exit Sub
Handler:
    ReDim SubKeys(0)
    Resume Next
End Sub
#End If


