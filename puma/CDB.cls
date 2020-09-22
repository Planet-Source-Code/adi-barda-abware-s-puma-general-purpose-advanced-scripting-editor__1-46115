VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_cn As ADODB.Connection

Public Function GetRecordset(ByVal sql As String, Optional ByVal OpenForEdit As Boolean = False) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    
    On Error GoTo err_proc
    
    Set rs = New ADODB.Recordset
     
    rs.CursorLocation = adUseClient
    If OpenForEdit Then
        rs.Open sql, m_cn, adOpenKeyset, adLockOptimistic
    Else
        rs.Open sql, m_cn, adOpenForwardOnly, adLockReadOnly
        rs.ActiveConnection = Nothing
    End If
    
    Set GetRecordset = rs
    
    
    Exit Function
    
err_proc:
    
    
End Function

Public Function Execute(ByVal sql As String) As Boolean

    
    Execute = False
    On Error GoTo err_proc
    
    m_cn.Execute sql
    
    Execute = True
    Exit Function
    
    
err_proc:
    'MsgBox Err.Description
    
End Function
 
Public Function Init(Optional ByVal con_str As String = "") As Boolean

    On Error GoTo err_proc
    
    Dim sConStr As String
    
    Set m_cn = New ADODB.Connection
    
    If con_str <> "" Then
        sConStr = con_str
    Else
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetDBPath("Puma.mdb") & ";Persist Security Info=False"
    End If
    
    m_cn.ConnectionString = sConStr
    m_cn.Open

    Init = True
    Exit Function
    
err_proc:
    Init = False
    
End Function
 
Private Sub Class_Terminate()

    On Error Resume Next
    
    If Not m_cn Is Nothing Then
        m_cn.Close
        Set m_cn = Nothing
    End If
    
End Sub

