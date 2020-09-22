VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl abRTFEditor 
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11745
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   990
      ScaleHeight     =   225
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   2625
   End
   Begin MSComctlLib.ListView lstInteli 
      Height          =   1365
      Left            =   1230
      TabIndex        =   1
      Top             =   3990
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   2408
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1650
      ScaleHeight     =   2145
      ScaleWidth      =   4935
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   4965
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   420
         TabIndex        =   15
         Top             =   300
         Width           =   4485
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   315
         Left            =   3990
         TabIndex        =   14
         Top             =   630
         Width           =   915
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Previous"
         Height          =   315
         Left            =   3030
         TabIndex        =   13
         Top             =   630
         Width           =   915
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Replace text"
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   60
         TabIndex        =   8
         Top             =   990
         Width           =   4845
         Begin VB.TextBox txtReplace 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   870
            TabIndex        =   11
            Top             =   300
            Width           =   3585
         End
         Begin VB.CommandButton cmdReplace 
            Caption         =   "Replace"
            Height          =   315
            Left            =   3360
            TabIndex        =   10
            Top             =   660
            Width           =   1095
         End
         Begin VB.CommandButton cmdReplaceAll 
            Caption         =   "Replace All"
            Height          =   315
            Left            =   2220
            TabIndex        =   9
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Replace:"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   12
            Top             =   300
            Width           =   645
         End
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4680
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Find:"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   16
         Top             =   300
         Width           =   345
      End
      Begin VB.Label lblBar 
         BackColor       =   &H00FF0000&
         Height          =   225
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.TextBox txtDebugCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Debug"
      Top             =   5490
      Width           =   11715
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4920
      Left            =   -90
      ScaleHeight     =   2142.38
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   -150
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.Timer tmrDebug 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4080
      Top             =   4020
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   180
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtDebug 
      Height          =   915
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5760
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1614
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"abRTFEditor.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtScript 
      Height          =   5385
      Left            =   2460
      TabIndex        =   5
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9499
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"abRTFEditor.ctx":00B5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Puma.abTreeView tvScripts 
      Height          =   5415
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   9551
      ID_Field        =   ""
      Father_Field    =   ""
      Name_Field      =   ""
      Table_Name      =   ""
      DataSourceType  =   0
   End
   Begin VB.Image imgSplitter2 
      Height          =   75
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   5400
      Width           =   11700
   End
   Begin VB.Image img 
      Height          =   210
      Index           =   0
      Left            =   3780
      Picture         =   "abRTFEditor.ctx":016A
      Stretch         =   -1  'True
      Top             =   4740
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   4110
      Picture         =   "abRTFEditor.ctx":026E
      Top             =   4740
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgSplitter 
      Height          =   5355
      Left            =   2400
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   60
   End
End
Attribute VB_Name = "abRTFEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'resize constants
Const LOCK_COLOR = &HE0E0E0
Const sglSplitLimit = 500
Const sglSplitLimit2 = 1500
Const ORIGINALSCALE_DEBUG = 11715 / 12000
Const ORIGINALSCALE_TRV = 2385 / 12000
Const ORIGINALSCALE_SCRIPT = 9255 / 12000
Const ORIGINALVSCALE_TRV = 5415 / 6870 '4785 / 7980

'RTF editor objects && flags
Private m_iItem                         As Long
Private m_PaintFlag                     As Boolean
Private m_objLastNode                   As MSComctlLib.Node
Private m_bPaintText                    As Boolean
Private m_objEditor                     As CEditor
Private m_bChangeCode                   As Boolean

Private mbMoving                        As Boolean
Private mbMoving2                       As Boolean

Private Type CodeBlockCondition
    sStartBlock As String
    sEndBlock As String
    ExecutionType As ScriptCommandExecution
End Type
Private m_arrCodeBlocks() As CodeBlockCondition

Public Enum BlockCommandResult
    BCR_CONTINUE = 0
    BCR_STOP = 1
    BCR_ERROR = 2
    BCR_EXECUTED_WITH_ERROR = 3
    BCR_EXECUTED_OK = 4
    
End Enum

Public Enum ScriptCommandExecution
    SCMD_IGNORE = 0
    SCMD_COMMIT = 1
End Enum


'debugger statuses
Private Enum DebugControlerState
    ST_OFF = 0
    ST_RUN = 1
    ST_PAUSE = 2
End Enum

'debug controler - controls the debugger state
Private Type DebugControler
    iState As DebugControlerState
    iCommand As Long
End Type
Private m_objDebugControler As DebugControler

'debug sequence struct
Private Type DebugSequence
    oNode As MSComctlLib.Node 'script node
    sItemCode As String 'script code
    objScript As MSScriptControl.ScriptControl
    iCurrentLine As Long
    iStart As Long
    iEnd As Long
    bFree As Boolean
    iPrevSeq As Long 'the caller sequence if any
    LastCmd As String
End Type
Private m_objDebugSeq(100) As DebugSequence 'up to 100 debug sequence running at once
Private m_iActiveSeq As Long 'current debug sequence

'external object struct
Private Type ScriptObjectDef
    TheObject As Object
    ObjName As String
End Type

Private m_objScriptObjects(100) As ScriptObjectDef 'up to 100 external objects
Private m_bUseDebugger As Boolean 'using debugger?

'DB connection
Private m_cn As ADODB.Connection

'------------------ events -------------------
Public Event SetDefaultObjects(ByRef ScriptObject As MSScriptControl.ScriptControl)
Public Event NodeSelected(ByVal NodeText As String)
Public Event DebugEnd()
Public Event DebugStart()


'Faked caption bar
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION              As Long = 2
Private Const WM_NCLBUTTONDOWN       As Long = &HA1
Private Const WM_LBUTTONDBLCLK       As Long = &H203
Private Const LBL_BACK_COLOR = &HFFC0C0

Public Function AddCodeBlockDefinition(ByVal sStartBlock As String, _
                                       ByVal sEndBlock As String, _
                                       ByVal eExecType As ScriptCommandExecution) As Boolean
                                       
                                       
    Dim iLen As Long
    
    iLen = UBound(m_arrCodeBlocks())
    If m_arrCodeBlocks(iLen).sStartBlock <> "" Then
        iLen = iLen + 1
    End If
    ReDim Preserve m_arrCodeBlocks(iLen)
    
    With m_arrCodeBlocks(iLen)
        .sStartBlock = sStartBlock
        .sEndBlock = sEndBlock
        .ExecutionType = eExecType
    End With
    
End Function

Private Sub lblBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  'Fake to move the window
    
    If Button = vbLeftButton Then
        ReleaseCapture  'release the mouse
        SendMessage picFind.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&  'non-client area button down (in caption)
    End If

End Sub



Private Sub lblClose_Click()
    UserControl.picFind.Visible = False
End Sub

Private Sub picFind_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            UserControl.picFind.Visible = False
    End Select
    
End Sub

Private Sub tvScripts_AfterItemAdd(ItemNode As MSComctlLib.Node, rs As ADODB.Recordset)

    'is it a file or folder?
    If ItemNode.Image = "FILE" Then
        If LenB(rs.Fields("FilePath").Value & "") > 0 Then
            ItemNode.Image = "LINKFILE"
        End If
    End If
    
End Sub

Private Sub UserControl_Initialize()

    ReDim m_arrCodeBlocks(0)
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF5
            m_bUseDebugger = True
            If m_objDebugControler.iState = ST_OFF Then
                Run
                Exit Sub
            End If
        Case vbKeyF12
            UserControl.txtDebug.text = UserControl.txtScript.TextRTF
            
    End Select
    
    If m_objDebugControler.iState = ST_RUN Then
        m_objDebugControler.iCommand = KeyCode
        
    End If

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With imgSplitter2
        picSplitter.Move .Left, .Top, .Width, .Height / 2
    End With
    picSplitter.Visible = True
    mbMoving2 = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > UserControl.Width - sglSplitLimit Then
            picSplitter.Left = UserControl.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sglPos As Single
    
    If mbMoving2 Then
        sglPos = Y + imgSplitter2.Top
        If sglPos < sglSplitLimit2 Then
            picSplitter.Top = sglSplitLimit2
        ElseIf sglPos > UserControl.Height - sglSplitLimit2 Then
            picSplitter.Top = UserControl.Height - sglSplitLimit2
        Else
            picSplitter.Top = sglPos
        End If
    End If
End Sub

Private Sub sizeControls(x As Single)
    On Error Resume Next

    
    'set the width
    If x < 1500 Then x = 1500
    If x > (UserControl.txtDebug.Width - 1500) Then x = UserControl.txtDebug.Width - 1500
    
    With UserControl.tvScripts
        .Left = UserControl.txtDebug.Left
        .Width = x - UserControl.txtDebug.Left
    End With
    
    UserControl.txtScript.Left = x + imgSplitter.Width
    UserControl.txtScript.Width = UserControl.txtDebug.Width - (UserControl.tvScripts.Width + imgSplitter.Width)
    
    imgSplitter.Top = UserControl.tvScripts.Top
    imgSplitter.Left = x
    imgSplitter.Height = UserControl.tvScripts.Height
    
End Sub

Private Sub sizeControlsVertival(Y As Single)
    On Error Resume Next

    
    'set the width
    If Y < 1500 Then Y = 1500
    If Y > (UserControl.Height - 1500) Then Y = UserControl.Height - 1500
    
    With UserControl.tvScripts
        '''.Top = UserControl.txtDebug.Left
        .Height = Y - .Top
        
        UserControl.txtScript.Height = .Height
        UserControl.imgSplitter.Height = .Height
        
        
        imgSplitter2.Top = Y
        imgSplitter2.Width = UserControl.txtDebug.Width

        UserControl.txtDebugCaption.Top = UserControl.tvScripts.Height + UserControl.tvScripts.Top + UserControl.imgSplitter2.Height
        UserControl.txtDebug.Top = UserControl.txtDebugCaption.Top + UserControl.txtDebugCaption.Height
        UserControl.txtDebug.Height = UserControl.Height - UserControl.txtDebug.Top - Screen.TwipsPerPixelY * 5
        
        'UserControl.txtDebug.Top = UserControl.txtScript.Top + UserControl.txtScript.Height + UserControl.imgSplitter2.Height
        'UserControl.txtDebugCaption.Top = UserControl.txtDebug.Top
    End With
    
    
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    sizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub imgSplitter2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    sizeControlsVertival picSplitter.Top
    picSplitter.Visible = False
    mbMoving2 = False
End Sub

'''Public Sub FlashWorkLed(Optional ByVal TurnOff As Boolean = False)
'''
'''    If TurnOff Then
'''        UserControl.picWork.Visible = False
'''        UserControl.cmdStop.Enabled = False
'''        UserControl.cmdRun.Enabled = True
'''    Else
'''        UserControl.picWork.Visible = True
'''        UserControl.lblPool.Visible = Not UserControl.lblPool.Visible
'''    End If
'''
'''End Sub

Public Sub WriteDebug(ByVal Msg As String)


    If Len(UserControl.txtDebug.text) > 4096 Then
        UserControl.txtDebug.text = ""
    End If
    
    UserControl.txtDebug.text = UserControl.txtDebug.text & Msg & vbNewLine
    UserControl.txtDebug.SelStart = Len(UserControl.txtDebug.text)
    
    
End Sub

Private Sub HideDebugCursor()
    m_objEditor.HighLight vbWhite
    UserControl.txtScript.SelLength = 0
End Sub

Private Function TrimEX(ByVal str As String) As String
    
    Dim i As Long
    Dim s As String
    
    str = Trim$(str)
    
    For i = 1 To Len(str)
        s = Mid$(str, i, 1)
        If Asc(s) > 13 Then
            TrimEX = TrimEX & s
        End If
        
    Next i
    
    
    
End Function


Private Sub StartDebug(ByRef ScriptObj As MSScriptControl.ScriptControl)
    
    'get next free sequence stack
    
    m_iActiveSeq = GetFreeSequence() 'get free script sequence from the stack
    With m_objDebugSeq(m_iActiveSeq)
        .LastCmd = ""
        .bFree = False
        .iStart = 1
        Set .objScript = ScriptObj
        
        Set .oNode = UserControl.tvScripts.CurrentNode
        .sItemCode = UserControl.tvScripts.CurrentKey
        
    End With
    
    'set the debug controler
    m_objDebugControler.iState = ST_RUN
    m_objDebugControler.iCommand = vbKeyF8
    
    'start executing
    UserControl.tmrDebug.Enabled = True
    
End Sub

Private Function GetFreeSequence() As Long

    Dim i As Long
    
    GetFreeSequence = -1
    
    For i = 0 To 100
        If m_objDebugSeq(i).bFree Then
            GetFreeSequence = i
            Exit For
        End If
    Next i
    
    
End Function

Private Sub InitDebugSeq()

    Dim i As Long
    
    For i = 0 To 100
        m_objDebugSeq(i).bFree = True
    Next i
    
    m_objDebugControler.iCommand = 0 'key pressed
    m_objDebugControler.iState = ST_OFF 'debuger status
    
End Sub

Private Sub RunScript(ByVal sequence As Long)

    Dim sRow As String
    Dim sCode As String
    Dim bStop As Boolean
    Dim sErr As String
    Dim i As Long
    
    bStop = False
    
    'first, execute the last command
    With m_objDebugSeq(sequence)
        If LenB(.LastCmd) > 0 Then
            .objScript.ExecuteStatement .LastCmd
            .LastCmd = ""
            
            'in case the last command changed the active script
            '   i have to update the sequence number
            sequence = m_iActiveSeq
        End If
    End With
    
    
    Dim iRetval As BlockCommandResult
    
    With m_objDebugSeq(sequence)
    
        Do Until bStop
        
            If .iStart <> 0 Then
                .iEnd = InStr(.iStart, UserControl.txtScript.text, vbNewLine)
                If .iEnd = 0 Then
                    .iStart = 0
                Else
                    sRow = TrimEX(Mid$(UserControl.txtScript.text, .iStart, .iEnd - .iStart))
                    
                    If LenB(sRow) > 0 And InStr(1, sRow, "'") <> 1 Then
                    
                        'Check command blocks
                        
                        For i = 0 To UBound(m_arrCodeBlocks)
                            With m_arrCodeBlocks(i)
                                iRetval = CheckBlock(sequence, sRow, .sStartBlock, .sEndBlock, .ExecutionType, sErr)
                                If iRetval <> BCR_CONTINUE Then
                                    Exit For
                                End If
                            End With
                        Next i
                        
                        
                        'end of block checks
                        bStop = (iRetval = BCR_STOP)
                        
                        
                        If iRetval = BCR_CONTINUE Then
                            UserControl.txtScript.SelStart = .iStart
                            ShowDebugLine UserControl.txtScript.GetLineFromChar(.iStart) + 1, .iStart, sRow
                            .LastCmd = sRow
                            .iStart = .iEnd + 1
                            bStop = True

                        End If
                        
                    
                    Else
                        .iStart = .iEnd + 1
                    End If
                    
                    
                End If
                
            Else
                'end current sequence
                .bFree = True
                If m_iActiveSeq = 0 Then
                    'no previous sequnce, stop process
                    m_bUseDebugger = False
                    m_objDebugControler.iState = ST_OFF
                    UserControl.tmrDebug.Enabled = False
                    HideDebugCursor
                    RaiseEvent DebugEnd
                    bStop = True
                    m_bChangeCode = False 'clear changes flag
                Else
                    'show the previous script
                    SelectActiveScript m_objDebugSeq(sequence - 1).sItemCode
                    CallNodeSelected m_objDebugSeq(sequence - 1).oNode, m_objDebugSeq(sequence - 1).sItemCode
                    
                    'activate it
                    m_iActiveSeq = m_iActiveSeq - 1
                    bStop = True
                    m_objDebugControler.iCommand = vbKeyF8
                    
                    GoTo exit_proc 'skip the iCommand clear in order to -
                                   '-  let it execute the next F8 press
                    
                End If
            
                
            End If 'istart<>0
        
        
        Loop
        
    End With
    
    m_objDebugControler.iCommand = 0
    
    
exit_proc:
    
End Sub

Private Function CheckBlock(ByVal sequence As Long, ByVal sRow As String, _
                            ByVal StartBlock As String, _
                            ByVal EndBlock As String, _
                            ByVal cmd As ScriptCommandExecution, _
                            ByRef sErrMsg As String) As BlockCommandResult
                            
    Dim sCode As String
    Dim iNextPos As Long
    
    With m_objDebugSeq(sequence)
    
        If (InStr(1, LCase$(sRow), StartBlock) = 1) Or (Trim$(LCase$(sRow)) = Trim$(StartBlock)) Then
        
            If cmd = SCMD_IGNORE Then
                
                .iEnd = InStr(.iStart, LCase$(UserControl.txtScript.text), EndBlock)
                If .iEnd > 0 Then
                    .iStart = GetNextCodePosition(.iEnd)
                    CheckBlock = BCR_EXECUTED_OK
                Else
                    sErrMsg = "Error in script: missing " & EndBlock
                    CheckBlock = BCR_ERROR
                End If
                
            Else
            
                'yes, add it to script control
                 .iEnd = InStr(.iStart, LCase$(UserControl.txtScript.text), EndBlock)
                If .iEnd > 0 Then
                    iNextPos = GetNextCodePosition(.iEnd)
                    sCode = Mid$(UserControl.txtScript.text, .iStart, iNextPos - .iStart)
                 
                    UserControl.txtScript.SelStart = .iStart
                    ShowDebugLine UserControl.txtScript.GetLineFromChar(.iStart) + 1, .iStart, sCode
                    .LastCmd = sCode
                 
                    .iStart = iNextPos
                    CheckBlock = BCR_STOP 'bStop = True
                Else
                    sErrMsg = "Error in script: missing " & EndBlock
                    CheckBlock = BCR_ERROR
                End If
                
            End If
        
        
        Else
            CheckBlock = BCR_CONTINUE 'didnt find the code block check the next block
        End If
        
    End With
    
End Function

Private Function GetNextCodePosition(ByVal iEndBlockPos As Long) As Long

    'return next code position
    
    'new row can start after CRLF or :
    With UserControl.txtScript
        GetNextCodePosition = InStr(iEndBlockPos, .text, vbNewLine)
    End With
    
End Function


Private Sub SelectActiveScript(ByVal sItemCode As String)

    Dim oNode As MSComctlLib.Node
    
    With UserControl.tvScripts.SourceTreeView
        For Each oNode In .Nodes
            If oNode.Key = "K" & sItemCode Then
                oNode.Selected = True
            End If
        Next oNode
    End With
    
End Sub

Private Sub ShowDebugLine(ByVal CurrentLine As Long, ByVal iStart As Long, ByVal sRow As String)
    
    
    With UserControl.txtScript
        .SetFocus
        .SelStart = iStart
        .SelLength = Len(Trim$(sRow))
        
        m_objEditor.HighLight vbYellow
    End With
       
   
End Sub


Private Sub AddScriptObjects()

    Dim ff As Long
    Dim s As String
    Dim arrFields As Variant
    Dim i As Long
    
    On Error GoTo err_proc
    
    For i = 0 To 99
        Set m_objScriptObjects(i).TheObject = Nothing
        m_objScriptObjects(i).ObjName = ""
    Next i
    
    i = 0
    ff = FreeFile
    Open App.path & "\ScriptObjects.dat" For Input As #ff
    Do Until EOF(ff)
        Line Input #ff, s
        If s <> "" Then
            'template=Object name;editor name;intlisense file
            'example:Center_XML_Parser.CXMLParser;xml;xml.dat
            
            arrFields = Split(s, ";")
            Set m_objScriptObjects(i).TheObject = CreateObject(arrFields(0))
            m_objScriptObjects(i).ObjName = arrFields(1)
            
            If UBound(arrFields) > 1 Then
                m_objEditor.LoadIntelisence App.path & "\" & arrFields(2)
            End If
            
            i = i + 1
            
        End If
nxt_obj:

    Loop
    Close #ff
    
    'add helper objects intelisense
    m_objEditor.LoadIntelisence App.path & "\inteli.dat"
    Exit Sub
    
err_proc:
   Select Case Err.Number
       Case 429
           'can't create object
           MBX "Object: " & arrFields(0) & " Doesn't exists on this machine."
           Resume nxt_obj
       
       Case Else
           MBX "Error occured: " & Err.Description
           
   End Select
   
End Sub

Private Sub InitEditor()

    Set m_objEditor = New CEditor
    m_objEditor.SetEditorObjects UserControl.txtScript, _
                                 UserControl.lstInteli, _
                                 UserControl.img(0).Picture, _
                                 UserControl.img(1).Picture, _
                                 UserControl.picToolTip, _
                                 UserControl.txtDebug
                                 
    
    AddScriptObjects
    
    'init debuger
    InitDebugSeq
    
    'by default don't use the debugger
    m_bUseDebugger = False
    
End Sub


Public Sub LoadKeywordsFromDB(ByVal sql As String, ByVal color As Long)

    Dim rs As ADODB.Recordset
    
    Set rs = GetRecordset(sql)
    If Not (rs Is Nothing) Then
        Do Until rs.EOF
            m_objEditor.AddEditorWord Mid$(rs.Fields(0).Value, 2, Len(rs.Fields(0).Value) - 2), color
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
    End If
    
End Sub

Private Sub ClearDebug()

    UserControl.txtDebug.text = ""

End Sub


Public Sub CallNodeSelected(ByVal oNode As MSComctlLib.Node, ByVal ItemCode As String)
    tvScripts_NodeSelected oNode, ItemCode
End Sub


Private Sub tmrDebug_Timer()

    
    Select Case m_objDebugControler.iCommand
    
        Case vbKeyF8
            RunScript m_iActiveSeq
            
    End Select
    
        
End Sub

Private Sub tvScripts_NodeSelected(ByVal oNode As MSComctlLib.Node, ByVal ItemCode As String)

    Dim rs As ADODB.Recordset
    
    
    If (m_bChangeCode And m_objDebugControler.iState = ST_OFF) Then
        If MsgBox("Do you want to save changes ?", vbYesNo Or vbQuestion) = vbYes Then
            SaveScript
        End If
    End If
    
    
    m_objEditor.ClearScript
    
    m_bChangeCode = False
    m_iItem = Val(ItemCode)
    
    If m_iItem = 0 Then
        If Not (m_objLastNode Is Nothing) Then
            m_objLastNode.BackColor = vbWhite
        End If
        Exit Sub
    End If
    
    UserControl.txtDebugCaption.text = "Debug: " & oNode.text & " - " & oNode.Key
    
    RaiseEvent NodeSelected(oNode.text)
    '''Me.Caption = oNode.text
    
    If Not (m_objLastNode Is Nothing) Then
        m_objLastNode.BackColor = vbWhite
    End If
    
    Set m_objLastNode = oNode
    m_objLastNode.BackColor = RGB(198, 216, 244)
    
    Set rs = GetRecordset("SELECT Script FROM tbl_Scripts WHERE PK=" & m_iItem)
    
    If Not (rs Is Nothing) Then
        If Not rs.EOF Then
            UserControl.txtScript.text = UserControl.txtScript.text & rs.Fields(0).Value & ""
        End If
        rs.Close
        Set rs = Nothing
        
        m_objEditor.PaintText
    End If
    
exit_proc:
    
    m_bChangeCode = False
    
End Sub


Private Sub txtScript_Change()

    If m_objDebugControler.iState = ST_OFF Then
        m_bChangeCode = True
    End If
    'usercontrol.txtDebug.Text = usercontrol.txtScript.TextRTF
    
End Sub

Public Property Get Editor() As CEditor
    Set Editor = m_objEditor
End Property

Public Function SetLinkedDoc(ByVal path As String) As Boolean

    Dim sSql As String
    
    If LenB(path) > 0 Then
        sSql = "UPDATE tbl_Scripts SET FilePath='" & path & "' WHERE PK=" & m_iItem
        SetLinkedDoc = Execute(sSql)
        If SetLinkedDoc Then
            UserControl.tvScripts.CurrentNode.Image = "LINKFILE"
        End If
    End If

End Function

Public Function ReloadLinkedDoc()

    If m_iItem <> 0 Then
        ReloadDocument m_iItem
        tvScripts_NodeSelected m_objLastNode, m_iItem
    Else
        MsgBox "No script item selected.", vbInformation
    End If
    
    
End Function

Public Sub ReloadDocument(Optional ByVal DocNum As Long = -1)

    Dim sSql As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo err_proc
    
    Screen.MousePointer = vbHourglass
    'Note: if DocNum=-1 then, all the linked documents will be reloaded
    
    sSql = "SELECT FilePath,Script FROM tbl_Scripts WHERE (FilePath IS NOT NULL) AND (FilePath<>'')"
    If DocNum <> -1 Then
        sSql = sSql & " AND PK=" & DocNum
    End If
    
    Set rs = GetRecordset(sSql, True)
    Do Until rs.EOF
        If LenB(rs.Fields(0).Value & "") > 0 Then
            rs.Fields("Script").Value = GetFile(rs.Fields(0).Value & "", "")
        End If
        rs.MoveNext
    Loop
    
    rs.UpdateBatch
    rs.Close
    Set rs = Nothing
    
exit_proc:
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_proc:
    Resume exit_proc
    
End Sub


Public Function LoadScriptFromFile(ByVal path As String) As Boolean

    Dim ff As Long
    Dim s As String
    Dim s2 As String
    
    On Error GoTo err_proc
    
    LoadScriptFromFile = False
    If path <> "" Then
    
        ff = FreeFile
        Open path For Input As ff
        
        Do Until EOF(ff)
            Line Input #ff, s
            If Right$(s, 1) <> vbNewLine Then
                s = s & vbNewLine
            End If
            s2 = s2 & s
        Loop
        
        UserControl.txtScript.text = s2
        m_objEditor.PaintText
        
        LoadScriptFromFile = True
    End If
    
    Exit Function
    
err_proc:
    

End Function

Public Function SaveToFile(ByVal path As String) As Boolean

    On Error GoTo err_proc
    
    Dim ff As Long
    
    
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, txtScript.text
    Close #ff
        
    
    Exit Function
    
err_proc:

End Function

Public Property Get Script() As String
    Script = UserControl.txtScript.text
End Property

Public Property Let Script(ByVal NewVal As String)
    UserControl.txtScript.text = NewVal
End Property

Public Property Get LastNode() As MSComctlLib.Node
    Set LastNode = m_objLastNode
End Property

Public Function DeleteCurrentScript() As Boolean

    UserControl.tvScripts.DeleteCurrentNode
    Set m_objLastNode = Nothing

End Function

Public Property Let ScriptLocked(ByVal NewVal As Boolean)

    UserControl.txtScript.Locked = NewVal
    
    If UserControl.txtScript.Locked Then
        UserControl.txtScript.BackColor = vbWhite 'LOCK_COLOR
    Else
        UserControl.txtScript.BackColor = vbWhite
    End If

End Property

Public Property Get ScriptLocked() As Boolean
    ScriptLocked = UserControl.txtScript.Locked
End Property

Public Function AddNewScript(ByVal script_name As String) As Boolean

    UserControl.tvScripts.Add_Branch script_name
    UserControl.tvScripts.SetFocus

End Function

Public Function Run() As Boolean

    On Error GoTo err_proc
    
    Dim i As Long
    Dim objScript As MSScriptControl.ScriptControl
    
    'stop debugger
    UserControl.tmrDebug.Enabled = False
    
    Set objScript = New MSScriptControl.ScriptControl
    
    UserControl.txtDebug.SetFocus
    
    ClearDebug
    
    
    With objScript
        .Timeout = NoTimeout
        .Language = "VBScript"
        ''.AddCode Me.txtScript.text
        
        'Set helper objects
        RaiseEvent SetDefaultObjects(objScript)
        
        'load dynamic objects
        For i = 0 To 99
            If m_objScriptObjects(i).ObjName <> "" Then
                .AddObject m_objScriptObjects(i).ObjName, m_objScriptObjects(i).TheObject
            End If
        Next i
            
        If m_bUseDebugger Then
        
            'Notify client about debugging process start
            RaiseEvent DebugStart
            AddAllFunctions objScript, UserControl.txtScript.text
            StartDebug objScript
            
        Else
            '.AddCode UserControl.txtScript.text
            .ExecuteStatement UserControl.txtScript.text  '"Process" 'when running without debugger, start at this function
            
        End If
    
    End With
    
    
    Exit Function
    
err_proc:
    MsgBox "Error in VBScript: " & Err.Description
    
End Function

Private Function AddAllFunctions(ByRef objScript As MSScriptControl.ScriptControl, ByVal sScript As String) As Boolean


    Dim iStart As Long
    Dim iEnd As Long
    Dim bStop As Boolean
    Dim sRow As String
    Dim sCode As String
    
    bStop = False
    iStart = 1
    
    Do Until bStop
        
        iEnd = InStr(iStart, sScript, vbNewLine)
        If iEnd = 0 Then
            bStop = True
        Else
            sRow = TrimEX(Mid$(sScript, iStart, iEnd - iStart))
            
            If LenB(sRow) > 0 And InStr(1, sRow, "'") <> 1 Then
                'is it a sub?
                If InStr(1, LCase$(sRow), "sub ") = 1 Then
                    'yes, add it to script control
                    iEnd = InStr(iStart, LCase$(sScript), "end sub")
                    sCode = Mid$(sScript, iStart, iEnd - iStart + 7)
                    objScript.AddCode sCode
                    iStart = iEnd + 7
                ElseIf InStr(1, LCase$(sRow), "function ") = 1 Then ' function?
                   'yes, add it to script control
                    iEnd = InStr(iStart, LCase$(sScript), "end function")
                    sCode = Mid$(sScript, iStart, iEnd - iStart + 12)
                    objScript.AddCode sCode
                    iStart = iEnd + 12
                Else
                    'scan next line
                    iStart = iEnd + 1
                    
                End If
            
            Else
                iStart = iEnd + 1
            End If
            
        End If
    Loop
    
End Function

Public Function SaveScript() As Boolean

    Dim iPos As Long
    
    If m_iItem <> 0 Then
    
        UserControl.txtScript.BackColor = vbGreen
        DoEvents
        Execute "UPDATE tbl_Scripts SET Script='" & LSTR(UserControl.txtScript.text) & "' WHERE PK=" & m_iItem
        Sleep 300
        UserControl.txtScript.BackColor = vbWhite
        
        iPos = UserControl.txtScript.SelStart
        m_objEditor.PaintText
        
        UserControl.txtScript.SelStart = iPos
        UserControl.txtScript.SetFocus
        
        m_bChangeCode = False
        
    End If
    
End Function

Public Function PaintText() As Boolean
    DoEvents
    m_objEditor.PaintText
End Function


Public Function FindText(ByVal txt As String)
    UserControl.txtFind.text = txt
    UserControl.picFind.Visible = True
End Function

Private Sub cmdNext_Click()

    Dim i As Long
    
    With UserControl.txtScript
        .HideSelection = False
        i = .Find(UserControl.txtFind.text, .SelStart + 1)
        If i > .SelStart Then
            .SelStart = i
        End If
    End With
    
End Sub

Private Sub cmdPrev_Click()

    Dim iStart As Long
    
    With UserControl.txtScript
        .HideSelection = False
        iStart = InStrRev(.text, UserControl.txtFind.text, .SelStart - 1, vbTextCompare)
        If iStart > 0 Then
        iStart = .Find(UserControl.txtFind.text, iStart - 1)
            If iStart <> -1 Then
                If iStart < .SelStart Then
                    .SelStart = iStart
                End If
            End If
        End If
        
    End With
    
End Sub

Private Sub cmdReplace_Click()

    With UserControl.txtScript
        .SelText = UserControl.txtReplace.text
    End With
    
End Sub

Private Sub cmdReplaceAll_Click()

    Dim iStart As Long
    Dim bStop As Boolean
    
    bStop = False
    With UserControl.txtScript
    
        .SelStart = 1
        iStart = 1
        
        Do Until bStop
            'find the next pos
            cmdNext_Click
            bStop = CBool((.SelStart = iStart) Or .SelLength = 0)
            iStart = .SelStart
            'replace
            If Not bStop Then
                cmdReplace_Click
            End If
        Loop
        
    End With
    
End Sub

Public Property Get SelText() As String
    SelText = UserControl.txtScript.SelText
End Property

Public Property Get SourceTreeView() As MSComctlLib.TreeView
    Set SourceTreeView = UserControl.tvScripts.SourceTreeView
End Property

Private Sub UserControl_Resize()
    ResizeCTL
End Sub

Private Sub UserControl_Show()


    Dim sConStr As String
    
    
    InitEditor
    
    With UserControl.tvScripts
    
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetDBPath("Puma.mdb") & ";Persist Security Info=False"
        .Connect_To_Database sConStr
        .SetTableStructure "tbl_Scripts", "PK", "Father", "Name"
        .Build_Tree
        
        .SourceTreeView.Nodes(1).Selected = True
        .SourceTreeView.SelectedItem.Expanded = True
        .SourceTreeView.Nodes(1).text = "Scripts"
        .SourceTreeView.Indentation = 1
        
    End With
    
    With UserControl.imgLst
        .ListImages.Add , "Method", UserControl.img(1).Picture
        .ListImages.Add , "Property", UserControl.img(0).Picture
        
    End With
    
    ClearDebug
    
    
    'load all the keywords from my database
    LoadKeywordsFromDB "SELECT  Keyword FROM Keywords1", vbBlue
    LoadKeywordsFromDB "SELECT  Keyword FROM Keywords2", vbRed
    
    m_bChangeCode = False


End Sub


Private Sub ResizeCTL()

    If UserControl.Width > 2500 Then
        'horizontal
        UserControl.imgSplitter.Left = UserControl.tvScripts.Left + UserControl.tvScripts.Width
        UserControl.txtScript.Left = tvScripts.Width + tvScripts.Left + UserControl.imgSplitter.Width
        UserControl.txtScript.Width = UserControl.Width - UserControl.txtScript.Left
        UserControl.txtDebug.Width = UserControl.txtScript.Left + UserControl.txtScript.Width - UserControl.txtDebug.Left
        UserControl.txtDebugCaption.Width = UserControl.txtDebug.Width
        
        
        'vertical
        UserControl.tvScripts.Height = UserControl.Height * ORIGINALVSCALE_TRV
        UserControl.txtScript.Height = UserControl.tvScripts.Height
        UserControl.imgSplitter.Height = UserControl.tvScripts.Height
        UserControl.picSplitter.Height = UserControl.tvScripts.Height
        UserControl.txtDebugCaption.Top = UserControl.tvScripts.Height + UserControl.tvScripts.Top + UserControl.imgSplitter2.Height
        UserControl.txtDebug.Top = UserControl.txtDebugCaption.Top + UserControl.txtDebugCaption.Height
        UserControl.txtDebug.Height = UserControl.Height - UserControl.txtDebug.Top - Screen.TwipsPerPixelY * 5
        UserControl.imgSplitter2.Top = UserControl.tvScripts.Top + UserControl.tvScripts.Height

    End If
    
End Sub

Private Function GetRecordset(ByVal sql As String, Optional ByVal OpenForEdit As Boolean = False) As ADODB.Recordset

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

Private Function Execute(ByVal sql As String) As Boolean

    
    Execute = False
    On Error GoTo err_proc
    
    m_cn.Execute sql
    
    Execute = True
    Exit Function
    
    
err_proc:
    'MsgBox Err.Description
    
End Function
 
Public Function InitDBConnection(Optional ByVal con_str As String = "") As Boolean

    On Error GoTo err_proc
    
    Dim sConStr As String
    
    'set new adodb connection
    Set m_cn = New ADODB.Connection
    
    If con_str <> "" Then
        sConStr = con_str
    Else
        sConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetDBPath("Puma.mdb") & ";Persist Security Info=False"
    End If
    
    m_cn.ConnectionString = sConStr
    m_cn.Open

    InitDBConnection = True
    Exit Function
    
err_proc:
    InitDBConnection = False
    
End Function

Public Property Let PopupMenuObject(ByRef objMenu As VB.Menu)

    tvScripts.PopupMenuObject = objMenu
    
End Property
