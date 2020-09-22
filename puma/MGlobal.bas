Attribute VB_Name = "MGlobal"
Option Explicit


Public Type General_Params
    LocalAreaNum As String
    OutLineNum As String
    CheckFaxNum As Boolean
End Type

Public g_DB As CDB
Public g_sInputText As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function InputBoxEX(ByVal title As String) As String

    frmInputBox.lblPrompt.Caption = title
    frmInputBox.Show 1
    InputBoxEX = g_sInputText
    
End Function

Public Function GetDBPath(ByVal Filename As String) As String

    Dim s As String
    
    On Error GoTo err_proc
    
    
    s = App.path & "\" & Filename
    GetDBPath = s
    Exit Function
    
err_proc:
    
End Function

Public Function MBX(ByVal Msg As String, _
       Optional MsgStyle As VbMsgBoxStyle = vbOKOnly Or vbInformation Or vbMsgBoxRight Or vbMsgBoxRtlReading) As VbMsgBoxResult

    If MsgStyle <> (vbOKOnly Or vbInformation Or vbMsgBoxRight Or vbMsgBoxRtlReading) Then
        MsgStyle = MsgStyle Or vbMsgBoxRight Or vbMsgBoxRtlReading
    End If
    
    MBX = MsgBox(Msg, MsgStyle, "îòøëú äôöä")
    
End Function

Public Function LSTR(ByVal str As String) As String

    LSTR = Replace$(str, "'", "''")
    
End Function

Public Function GetFileName(ByRef dlg1 As mscomdlg.CommonDialog, ByVal file_ext As String, ByVal desc As String) As String

    'First open the excel sheet
    dlg1.CancelError = True

    On Error GoTo err_proc
    'open dialog box
    dlg1.Filter = desc & " (*." & file_ext & ")|*." & file_ext
    dlg1.DefaultExt = "." & file_ext

    dlg1.ShowOpen
    GetFileName = dlg1.Filename
    Exit Function
    
err_proc:
    
End Function

Public Sub DebugWrite(ByVal Msg As String, Optional ByVal Header As Boolean = False)

    
    If Not Header Then
        Msg = "  " & Msg
    End If
    
    Msg = Now & ": " & Msg
    frmScripts.rtfEditor.WriteDebug Msg
    
End Sub

Public Function GetFile(ByVal path As String, ByRef Feedback As String) As String


    On Error GoTo err_proc

    Dim sMsg            As String
    Dim s               As String
    Dim ff              As Long
    
    
    ff = FreeFile()
    sMsg = ""
    Open path For Input As #ff
    Do Until EOF(ff)
        Line Input #ff, s
        sMsg = sMsg & s & vbNewLine
    Loop
    Close #ff
    
    If Trim(sMsg) = "" Then Exit Function
    
    GetFile = sMsg
    Exit Function
    
err_proc:
    Feedback = Err.Description
    
exit_proc:
    Exit Function


End Function

