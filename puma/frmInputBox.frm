VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INPUT"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   660
      Width           =   5325
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "xxxxxxx"
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   735
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    g_sInputText = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()

    g_sInputText = Me.txtText.text
    Unload Me
    
End Sub
