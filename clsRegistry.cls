VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'--------------------------------------------------
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'--------------------------------------------------
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_BINARY = 3                     ' Free form binary
Const REG_DWORD = 4                      ' 32-bit number
Const ERROR_SUCCESS = 0&

Enum RegHKey
  Classes_Root = 0
  Current_User = 1
  Local_Machine = 2
  Users = 3
  Current_Config = 4
  Dyn_Data = 5
End Enum

Public regKey As RegHKey
Private mKey As Long

Private Function GetKey(hKey As RegHKey)
  
  If hKey = Classes_Root Then
    GetKey = HKEY_CLASSES_ROOT
  ElseIf hKey = Current_Config Then
    GetKey = HKEY_CURRENT_CONFIG
  ElseIf hKey = Current_User Then
    GetKey = HKEY_CURRENT_USER
  ElseIf hKey = Dyn_Data Then
    GetKey = HKEY_DYN_DATA
  ElseIf hKey = Local_Machine Then
    GetKey = HKEY_LOCAL_MACHINE
  ElseIf hKey = Users Then
    GetKey = HKEY_USERS
  End If
  
End Function

Public Sub CreateKey(hKey As RegHKey, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long

hKey = GetKey(hKey)
lRegResult = RegCreateKey(hKey, strPath, hCurKey)

If lRegResult <> ERROR_SUCCESS Then
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Sub DeleteKey(ByVal hKey As RegHKey, ByVal strPath As String)
Dim lRegResult As Long

hKey = GetKey(hKey)
lRegResult = RegDeleteKey(hKey, strPath)

End Sub

Public Sub DeleteValue(ByVal hKey As RegHKey, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

hKey = GetKey(hKey)
lRegResult = RegOpenKey(hKey, strPath, hCurKey)

lRegResult = RegDeleteValue(hCurKey, strValue)

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetSettingString(hKey As RegHKey, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetSettingString = Default
Else
  GetSettingString = ""
End If


hKey = GetKey(hKey)

' Open the key and get length of string
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_SZ Then
    ' initialise string buffer and retrieve string
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
    
    ' format string
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Left(strBuffer, intZeroPos - 1)
    Else
      GetSettingString = strBuffer
    End If

  End If

Else
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingString(hKey As RegHKey, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

hKey = GetKey(hKey)
lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingLong(ByVal hKey As RegHKey, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

hKey = GetKey(hKey)
' Set up default value
If Not IsEmpty(Default) Then
  GetSettingLong = Default
Else
  GetSettingLong = 0
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lDataBufferSize = 4       ' 4 bytes = 32 bits = long

lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_DWORD Then
    GetSettingLong = lBuffer
  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingLong(ByVal hKey As RegHKey, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long

hKey = GetKey(hKey)
lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingByte(ByVal hKey As RegHKey, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long

hKey = GetKey(hKey)

' setup default value
If Not IsEmpty(Default) Then
  If VarType(Default) = vbArray + vbByte Then
    GetSettingByte = Default
  Else
    GetSettingByte = 0
  End If

Else
  GetSettingByte = 0
End If

' Open the key and get number of bytes
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_BINARY Then
  
    ' initialise buffers and retrieve value
    ReDim byBuffer(lDataBufferSize - 1) As Byte
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
    
    GetSettingByte = byBuffer

  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingByte(ByVal hKey As RegHKey, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)
' Make sure that the array starts with element 0 before passing it!
' (otherwise it will not be saved!)

Dim lRegResult As Long
Dim hCurKey As Long

hKey = GetKey(hKey)
lRegResult = RegCreateKey(hKey, strPath, hCurKey)

' Pass the first array element and length of array
lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetAllKeys(hKey As RegHKey, strPath As String) As Variant
' Returns: an array in a variant of strings

Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer

hKey = GetKey(hKey)
lCounter = 0

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do

  'initialise buffers (longest possible length=255)
  lDataBufferSize = 255
  strBuffer = String(lDataBufferSize, " ")
  lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

  If lRegResult = ERROR_SUCCESS Then
  
    'tidy up string and save it
    ReDim Preserve strNames(lCounter) As String
    
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
    Else
      strNames(UBound(strNames)) = strBuffer
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

GetAllKeys = strNames
End Function

Public Function GetAllValues(hKey As RegHKey, strPath As String) As Variant
' Returns: a 2D array.
' (x,0) is value name
' (x,1) is value type (see constants)

Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

hKey = GetKey(hKey)
lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do
  ' Initialise bufffers
  lValueNameSize = 255
  strValueName = String$(lValueNameSize, " ")
  lDataBufferSize = 4000
  
  lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
  
  If lRegResult = ERROR_SUCCESS Then
    
    ' Save the type
    ReDim Preserve strNames(lCounter) As String
    ReDim Preserve lTypes(lCounter) As Long
    lTypes(UBound(lTypes)) = lValueType
    
    'Tidy up string and save it
    intZeroPos = InStr(strValueName, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
    Else
      strNames(UBound(strNames)) = strValueName
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

'Move data into array
Dim Finisheddata() As Variant
ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant

For lCounter = 0 To UBound(strNames)
  Finisheddata(lCounter, 0) = strNames(lCounter)
  Finisheddata(lCounter, 1) = lTypes(lCounter)
Next

GetAllValues = Finisheddata

End Function

