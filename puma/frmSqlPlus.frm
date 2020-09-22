VERSION 5.00
Begin VB.Form frmSqlPlus 
   Caption         =   "Run SQL"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   Icon            =   "frmSqlPlus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5790
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3810
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "Command structure"
      ForeColor       =   &H00C00000&
      Height          =   2325
      Left            =   90
      TabIndex        =   3
      Top             =   210
      Width           =   7065
      Begin VB.ComboBox cboConstr 
         Height          =   315
         ItemData        =   "frmSqlPlus.frx":000C
         Left            =   1350
         List            =   "frmSqlPlus.frx":001F
         TabIndex        =   7
         Top             =   300
         Width           =   3855
      End
      Begin VB.TextBox txtParam1 
         Height          =   285
         Left            =   1350
         TabIndex        =   6
         Top             =   870
         Width           =   3855
      End
      Begin VB.TextBox txtParam2 
         Height          =   285
         Left            =   1350
         TabIndex        =   5
         Top             =   1380
         Width           =   3855
      End
      Begin VB.TextBox txtParam3 
         Height          =   285
         Left            =   1350
         TabIndex        =   4
         Top             =   1890
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   1830
         Left            =   5310
         Picture         =   "frmSqlPlus.frx":0089
         Stretch         =   -1  'True
         Top             =   330
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Connect to:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Param #1:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Param #2:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   9
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Param #3:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   8
         Top             =   1890
         Width           =   735
      End
   End
   Begin VB.TextBox txtDebug 
      Appearance      =   0  'Flat
      Height          =   825
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2910
      Width           =   7095
   End
   Begin VB.CommandButton cmdRun 
      Height          =   405
      Left            =   6510
      Picture         =   "frmSqlPlus.frx":0472
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3810
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   2
      Top             =   2700
      Width           =   570
   End
End
Attribute VB_Name = "frmSqlPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sFile As String
Private m_iNodeID As Long
Private m_bShowGUI As Boolean

Public Sub ShowEX(ByVal NodeID As Long, ByVal sFile As String, ByVal params As String)
    
    
    m_sFile = sFile
    Me.cboConstr.ListIndex = 0
    
    m_iNodeID = NodeID
    LoadParams NodeID, params
    EnableParamFields
    
    If params <> "" Then
        m_bShowGUI = False
        cmdRun_Click
    Else
        m_bShowGUI = True
        Me.Show
    End If
    
End Sub

Private Sub LoadParams(ByVal NodeID As Long, ByVal params As String)

    Dim rs As ADODB.Recordset
    Dim arrParams As Variant
    
    If params <> "" Then
        arrParams = Split(params, ",")
        
        On Error GoTo exit_proc
        
        Me.cboConstr.text = arrParams(0)
        Me.txtParam1.text = arrParams(1)
        Me.txtParam2.text = arrParams(2)
        Me.txtParam3.text = arrParams(3)
    Else
        Set rs = g_DB.GetRecordset("SELECT Params FROM tbl_Scripts WHERE PK=" & NodeID)
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                If Not IsNull(rs.Fields(0).Value) Then
                    arrParams = Split(rs.Fields(0).Value, ",")
                    
                    On Error GoTo exit_proc
                    
                    Me.cboConstr.text = arrParams(0)
                    Me.txtParam1.text = arrParams(1)
                    Me.txtParam2.text = arrParams(2)
                    Me.txtParam3.text = arrParams(3)
                    
                End If
                
            End If
        End If
    End If
    
exit_proc:
    On Error Resume Next
    rs.Close
    Set rs = Nothing

End Sub

Private Sub SaveParams(ByVal NodeID As Long)
    
    Dim sSql As String
    
    sSql = "UPDATE tbl_Scripts SET params='" & _
            Replace$(Me.cboConstr.text & "," & Me.txtParam1.text & "," & Me.txtParam2.text & "," & Me.txtParam3.text, "'", "''") & _
            "' WHERE PK=" & NodeID
            
    g_DB.Execute sSql
    
            
End Sub

Private Sub cboConstr_Click()
    Me.txtDebug.text = GetCommand()
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()

    Dim sCmd As String
    
    If m_bShowGUI Then
        If MsgBox("Are you sure you want to run this SQL on: " & Me.cboConstr.text & " ?", vbYesNo Or vbQuestion) = vbYes Then
            sCmd = GetCommand()
            Shell sCmd, vbNormalFocus
            
            SaveParams m_iNodeID
            
        End If
    Else
        sCmd = GetCommand()
        Shell sCmd, vbNormalFocus
    End If
    
    Unload Me
    
End Sub

Private Function GetCommand() As String

    Dim sCmd As String
    
    sCmd = "SQLPLUS " & Me.cboConstr.text & _
           " @" & m_sFile
          
    'add params
    If Trim$(Me.txtParam1.text) <> "" Then
        sCmd = sCmd & " " & Me.txtParam1.text
        
        If Me.txtParam2.Enabled Then
        
            If Trim$(Me.txtParam2.text) <> "" Then
                sCmd = sCmd & " " & Me.txtParam2.text
                
                If Me.txtParam3.Enabled Then
                    If Trim$(Me.txtParam3.text) <> "" Then
                        sCmd = sCmd & " " & Me.txtParam3.text
                    End If 'param 3
                End If
                
            End If 'param 2
            
        End If
        
    End If 'param 1

    GetCommand = sCmd
    
End Function

Private Sub EnableParamFields()

    Me.txtParam2.Enabled = (Trim$(Me.txtParam1.text) <> "")
    Me.txtParam3.Enabled = (Trim$(Me.txtParam2.text) <> "")
    Me.txtDebug.text = GetCommand()
    
End Sub

Private Sub txtParam1_Change()
    EnableParamFields
End Sub

Private Sub txtParam2_Change()
    EnableParamFields
End Sub

Private Sub txtParam3_Change()
    EnableParamFields
End Sub
