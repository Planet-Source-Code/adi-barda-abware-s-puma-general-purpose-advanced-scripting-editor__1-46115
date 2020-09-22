VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function AppPath() As String
    AppPath = App.path
End Function

Public Function StopSequence()
    frmScripts.cmdStop.Value = True
End Function

Public Sub DebugPrint(ByVal str As String)
    DebugWrite str
End Sub

Public Sub SysDoEvents()
    DoEvents
End Sub

Public Sub Wait(ByVal ms As Long)

    Dim i As Long
    
    i = 0
    
    Do
        Sleep 50
        i = i + 50
        DoEvents
        
    Loop Until i >= ms
    
End Sub

Public Sub RunExeFile(ByVal path As String)

    Shell path, vbNormalFocus
    
End Sub

Public Function FileExists(ByVal path As String) As Boolean

    FileExists = CBool(Dir(path) <> "")
    
End Function

Public Sub DeleteFile(ByVal path As String)
    Kill path
End Sub
