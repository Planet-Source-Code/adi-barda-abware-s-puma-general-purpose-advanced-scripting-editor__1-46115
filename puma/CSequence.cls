VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sMsg          As String

Public Function LoadMessageFromFile(ByVal path As String) As String
    LoadMessageFromFile = GetFile(path, "")
End Function


Public Sub StopSequence()
   
    MsgBox "Demo function "
    DebugWrite "Activate StopSequence function"
   
End Sub

Public Sub FireMessage(ByVal path As String, _
                       ByVal queue As String, _
                       ByVal QueueType As Long, _
                       Optional ByVal Transactional As Boolean = True)

    MsgBox "Demo function "
    DebugWrite "Activate FireMessage function"
    
End Sub

Public Sub FireText(ByVal text As String, _
                    ByVal queue As String, _
                    ByVal Queue_type As Long, _
                    Optional ByVal Transactional As Boolean = True)

    MsgBox "Demo function "
    DebugWrite "Activate FireText function"
    

End Sub

Public Function CopyXmlFields(ByRef Fields As Variant, ByVal SrcXML As String, ByVal SrcPath As String, ByVal DestXML As String, ByVal DestPath As String) As String

    MsgBox "Demo function "
    DebugWrite "Activate CopyXmlFields function"
    
End Function

Public Function RunScript(ByVal ScriptKey As String, _
                          ByVal ScriptType As String, _
                          ByVal params As String) As Boolean
                          
    
    On Error GoTo err_proc
    
    Dim i As Long
    Dim oNode As MSComctlLib.Node
    
    Dim objTree As MSComctlLib.TreeView
    
    Set objTree = frmScripts.rtfEditor.SourceTreeView
    With objTree
        For Each oNode In .Nodes
            If oNode.Key = ScriptKey Then
                oNode.Selected = True
                
                frmScripts.rtfEditor.CallNodeSelected oNode, Right$(oNode.Key, Len(oNode.Key) - 1)
                
                Select Case ScriptType
                    Case "SQL"
                        frmScripts.CallSQLScript params
                    Case "VB"
                        frmScripts.CallVBScript
                    Case Else
                    
                End Select
                
                Exit Function
            End If
        Next oNode
    End With
    
err_proc:
    
End Function

