VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puma"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   465
      Left            =   3660
      TabIndex        =   0
      Top             =   2580
      Width           =   885
   End
   Begin VB.Label Label5 
      Caption         =   "(c) 2001-2003 All rights reserved"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2820
      Width           =   2865
   End
   Begin VB.Label Label4 
      Caption         =   "(R)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3210
      TabIndex        =   4
      Top             =   420
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Wrriten by adi barda - malam IT"
      Height          =   255
      Left            =   900
      TabIndex        =   3
      Top             =   1650
      Width           =   2715
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "General purpose run time editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   420
      TabIndex        =   2
      Top             =   1080
      Width           =   4005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "P U M A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   450
      TabIndex        =   1
      Top             =   450
      Width           =   3645
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub
