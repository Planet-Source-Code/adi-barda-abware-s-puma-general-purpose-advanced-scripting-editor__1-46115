VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmScripts 
   Caption         =   "Test scripts"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmScripts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Puma.abRTFEditor rtfEditor 
      Height          =   7335
      Left            =   0
      TabIndex        =   10
      Top             =   60
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   12938
   End
   Begin VB.CommandButton cmdSaveToFile 
      Caption         =   "Save to file"
      Height          =   525
      Left            =   3870
      Picture         =   "frmScripts.frx":000C
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7410
      Width           =   1005
   End
   Begin VB.CommandButton cmdSqlPlus 
      Caption         =   "Run SQL"
      Height          =   525
      Left            =   6930
      Picture         =   "frmScripts.frx":0596
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7410
      Width           =   915
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit script"
      Height          =   525
      Left            =   5970
      Picture         =   "frmScripts.frx":0998
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7410
      Width           =   945
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   525
      Left            =   9870
      Picture         =   "frmScripts.frx":0DD6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7410
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   525
      Left            =   10800
      Picture         =   "frmScripts.frx":0ED8
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7410
      Width           =   915
   End
   Begin VB.CommandButton cmdTypeNewProc 
      Caption         =   "Script frame"
      Height          =   525
      Left            =   4890
      Picture         =   "frmScripts.frx":1462
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7410
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   9360
      Top             =   7410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New script"
      Height          =   525
      Left            =   0
      Picture         =   "frmScripts.frx":19EC
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7410
      Width           =   975
   End
   Begin VB.CommandButton cmdLoadData 
      Caption         =   "Load script"
      Height          =   525
      Left            =   2850
      Picture         =   "frmScripts.frx":1AEE
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7410
      Width           =   1005
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete script"
      Height          =   525
      Left            =   1740
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmScripts.frx":1BF0
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7410
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   525
      Left            =   990
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmScripts.frx":1CF2
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7410
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Image imgLock 
      Height          =   255
      Index           =   1
      Left            =   330
      Picture         =   "frmScripts.frx":227C
      Top             =   8340
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgLock 
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "frmScripts.frx":26BA
      Top             =   8340
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_objSystem                     As CSystem
Private m_objSequence                   As CSequence
Private m_objReg                        As CReg


Private m_sSQLParams                    As String


Private Sub cmdSaveToFile_Click()

    On Error GoTo err_proc
    
    Dim sPath As String
    
    With Me.dlg1
        .CancelError = True
        .ShowSave
        
        sPath = .Filename
        rtfEditor.SaveToFile sPath
        
    End With
    
    Exit Sub
    
err_proc:
    
End Sub

Private Sub cmdSqlPlus_Click()

    Dim ff As Long
    
    
    With Me.rtfEditor
        If Trim$(.Script) = "" Then
            Exit Sub
        End If
        
        ff = FreeFile
        Open App.path & "\RunSql.sql" For Output As #ff
        Print #ff, .Script
        Close #ff
        
        If Not (.LastNode Is Nothing) Then
            frmSqlPlus.ShowEX Val(Right$(.LastNode.Key, Len(.LastNode.Key) - 1)), Chr$(34) & App.path & "\RunSql.sql" & Chr$(34), m_sSQLParams
        End If
        
        m_sSQLParams = ""
    End With
    
    
End Sub


Private Sub cmdDel_Click()

    If MsgBox("Are you sure you want to delete the current script ?", vbQuestion Or vbYesNo) = vbNo Then
        Exit Sub
    End If
    Me.rtfEditor.DeleteCurrentScript
    
End Sub

Private Sub cmdEdit_Click()
    
    Me.rtfEditor.ScriptLocked = Not Me.rtfEditor.ScriptLocked
    
    If Me.rtfEditor.ScriptLocked Then
        Me.cmdEdit.Caption = "Edit script"
    Else
        Me.cmdEdit.Caption = "Lock script"
    End If
    Me.cmdEdit.Picture = Me.imgLock(Abs(Me.rtfEditor.ScriptLocked))
    
End Sub

Private Sub cmdLoadData_Click()

    Dim sFile As String
    
    On Error GoTo err_proc
    
    sFile = GetFileName(dlg1, "*", "All files")
    If Me.rtfEditor.LoadScriptFromFile(sFile) Then
        If MsgBox("Do you want to link this file to the current node ?", vbQuestion Or vbYesNo) = vbYes Then
            Me.rtfEditor.SetLinkedDoc sFile
        End If
    End If
    
    
    Exit Sub
    
err_proc:
    
End Sub


Private Sub cmdNew_Click()

    Dim s As String
    Dim sName As String
    
    sName = InputBoxEX("Enter Script name please")
    If sName = "" Then
        Exit Sub
    End If
    
    Me.rtfEditor.AddNewScript sName
    Me.cmdTypeNewProc.value = True
    
End Sub

Private Sub cmdRun_Click()
    
    Me.cmdStop.Enabled = True
    Me.cmdRun.Enabled = False
    Me.rtfEditor.Run
    
End Sub

Private Sub cmdSave_Click()

    Me.rtfEditor.SaveScript
    
End Sub

Public Sub CallSQLScript(ByVal params As String)
    m_sSQLParams = params
    cmdSqlPlus_Click
End Sub

Public Sub CallVBScript()
    cmdRun_Click
End Sub


Private Sub cmdStop_Click()

    Me.cmdStop.Enabled = False
    Me.cmdRun.Enabled = True
    
End Sub

Private Sub cmdTypeNewProc_Click()

    Dim s As String
    
    
    If Me.rtfEditor.Script <> "" Then
        If MsgBox("Do you want to overide the existing script ?", vbQuestion Or vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    'Generate loading process script
    s = "'#vbs@@@ - Debugger directive. Don't update this row !"
    s = s & vbNewLine & vbNewLine & "' General process script" & vbNewLine
    s = s & "' Written by: " & vbNewLine
    s = s & "' Date: " & Now & vbNewLine & vbNewLine
    
    Me.rtfEditor.Script = s
    Me.rtfEditor.PaintText
    
End Sub

Private Sub Form_Load()

    'init database connection
    'you can use your connect string or let the object use its own
    rtfEditor.InitDBConnection
    rtfEditor.PopupMenuObject = MDIMain.mnuPopup
    rtfEditor.SourceTreeView.Appearance = ccFlat
    
    SetCodeVBBlocks
    
End Sub

Private Sub SetCodeVBBlocks()
    With rtfEditor
        .AddCodeBlockDefinition "sub ", "end sub", SCMD_IGNORE
        .AddCodeBlockDefinition "function ", "end function", SCMD_IGNORE
        .AddCodeBlockDefinition "for ", "next", SCMD_COMMIT
        .AddCodeBlockDefinition "if ", "end if", SCMD_COMMIT
        .AddCodeBlockDefinition "do ", "loop", SCMD_COMMIT
    End With
        
End Sub

Private Sub Form_Resize()

    With Me.rtfEditor
        .Top = 5 * Screen.TwipsPerPixelY
        .Width = Me.Width - .Left - 10 * Screen.TwipsPerPixelX
    End With
    
End Sub

Private Sub rtfEditor_DebugEnd()

    cmdStop_Click
    
End Sub

Private Sub rtfEditor_NodeSelected(ByVal NodeText As String)
    Me.Caption = NodeText
End Sub

Private Sub rtfEditor_SetDefaultObjects(ScriptObject As MSScriptControl.ScriptControl)

    'event from script control for running new script
    '- and time to set some default objects
    
    With ScriptObject
        Set m_objSequence = New CSequence
        Set m_objSystem = New CSystem
        Set m_objReg = New CReg
        
        .AddObject "system", m_objSystem
        .AddObject "seq", m_objSequence
        .AddObject "reg", m_objReg
    End With
    
End Sub
