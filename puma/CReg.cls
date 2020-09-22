VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

   Private Const REG_SZ As Long = 1
   Private Const REG_DWORD As Long = 4

   Private Const HKEY_CLASSES_ROOT = &H80000000
   Private Const HKEY_CURRENT_USER = &H80000001
   Private Const HKEY_LOCAL_MACHINE = &H80000002
   Private Const HKEY_USERS = &H80000003

   Private Const ERROR_NONE = 0
   Private Const ERROR_BADDB = 1
   Private Const ERROR_BADKEY = 2
   Private Const ERROR_CANTOPEN = 3
   Private Const ERROR_CANTREAD = 4
   Private Const ERROR_CANTWRITE = 5
   Private Const ERROR_OUTOFMEMORY = 6
   Private Const ERROR_ARENA_TRASHED = 7
   Private Const ERROR_ACCESS_DENIED = 8
   Private Const ERROR_INVALID_PARAMETERS = 87
   Private Const ERROR_NO_MORE_ITEMS = 259

   Private Const KEY_QUERY_VALUE = &H1
   Private Const KEY_SET_VALUE = &H2
   Private Const KEY_ALL_ACCESS = &H3F

   Private Const REG_OPTION_NON_VOLATILE = 0

   Private Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long


Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
       lType As Long, vValue As Variant) As Long

    Dim lValue As Long
    Dim sValue As String

    Select Case lType

        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
               lType, sValue, Len(sValue))

        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
               lType, lValue, 4)

    End Select
        
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
       String, vValue As Variant) As Long

    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)

    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
            ' For strings

        Case REG_SZ:
            sValue = String$(cch, 0)

            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
               sValue, cch)

            If lrc = ERROR_NONE Then

                vValue = Left$(sValue, cch - 1)

            Else

                vValue = Empty

            End If

            ' For DWORDS

        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
               lValue, cch)

            If lrc = ERROR_NONE Then vValue = lValue

        Case Else
            'all other data types not supported
            lrc = -1

    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit

End Function

Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)

    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
       vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
       0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
    
    ' Examples:
    'CreateNewKey "TestKey", HKEY_LOCAL_MACHINE
    'CreateNewKey "TestKey\SubKey1\SubKey2", HKEY_LOCAL_MACHINE

End Sub

Public Function QueryValue(sKeyName As String, sValueName As String) As Variant

    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
       KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue 'retval
    RegCloseKey (hKey)
    
    ' Examples:
    'QueryValue "TestKey\SubKey1", "StringValue"
    
End Function


