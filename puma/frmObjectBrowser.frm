VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmObjectBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object browser"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmObjectBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   9285
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
      Left            =   8280
      TabIndex        =   1
      Top             =   4140
      Width           =   915
   End
   Begin VB.ListBox lstObjects 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   3345
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3315
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   8400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstInteli 
      Height          =   3315
      Left            =   3420
      TabIndex        =   5
      Top             =   390
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5847
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
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      ScaleHeight     =   345
      ScaleWidth      =   9105
      TabIndex        =   4
      Top             =   3720
      Width           =   9135
   End
   Begin VB.Image img 
      Height          =   210
      Index           =   0
      Left            =   3600
      Picture         =   "frmObjectBrowser.frx":058A
      Stretch         =   -1  'True
      Top             =   390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   3930
      Picture         =   "frmObjectBrowser.frx":068E
      Top             =   390
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Interface definition"
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
      Index           =   1
      Left            =   3420
      TabIndex        =   3
      Top             =   180
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Objects"
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
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   660
   End
End
Attribute VB_Name = "frmObjectBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objXML As CXMLParser


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub ShowInterface(ByVal Keywrd As String)

    Dim i As Long
    Dim arrItems As Variant
    Dim objXML As CXMLParser
    Dim iRecCount As Long
    

    Set objXML = New CXMLParser
    
    With Me.lstInteli
    
        .ListItems.Clear
        '.Rows = 0
        Dim objItem As MSComctlLib.ListItem
        
        With objXML
        
            arrItems = m_objXML.FindItems("NEWITEM", "Class='" & Keywrd & "'")
            
            If .SetXMLDoc("<ROOT>" & arrItems(0) & "</ROOT>") Then
                iRecCount = .GetRecordCount("NEWITEM")
                For i = 0 To iRecCount - 1
                    Set objItem = lstInteli.ListItems.Add(, , .GetItem("method", "NEWITEM[" & i & "]"))        ', , SmallIcon
                    objItem.Tag = .GetItem("ToolTip", "NEWITEM[" & i & "]")
                    objItem.SmallIcon = IIf(.GetItem("IsMethod", "NEWITEM[" & i & "]"), "Method", "Property")
                    
                Next i
                
            End If
                   
            
        End With
        
        
        '''ShowToolTip
        
    End With
    
End Sub

Private Sub Form_Load()

    Set m_objXML = New CXMLParser
    m_objXML.SetXMLDoc frmScripts.rtfEditor.Editor.GetXMLInterface()
    
    Dim arrKeywords() As String
    Dim i As Long
    arrKeywords() = frmScripts.rtfEditor.Editor.GetKeywords()
    For i = 0 To UBound(arrKeywords)
        Me.lstObjects.AddItem arrKeywords(i)
    Next i
    
    With Me.lstInteli
        .ColumnHeaders.Add 1, , , .Width - 3 * Screen.TwipsPerPixelX
        .FullRowSelect = True
    End With
    
    
    With Me.imgLst
        .ListImages.Add , "Method", Me.img(1).Picture
        .ListImages.Add , "Property", Me.img(0).Picture
    End With
    
End Sub

Private Sub lstObjects_Click()

    ShowInterface Me.lstObjects.text
    
End Sub
