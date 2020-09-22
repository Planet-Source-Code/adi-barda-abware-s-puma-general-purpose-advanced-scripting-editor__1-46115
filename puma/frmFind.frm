VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Replace text"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   4845
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace All"
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   660
         Width           =   1095
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtReplace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   870
         TabIndex        =   6
         Top             =   300
         Width           =   3585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Replace:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Top             =   420
      Width           =   915
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   315
      Left            =   4110
      TabIndex        =   2
      Top             =   420
      Width           =   915
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   540
      TabIndex        =   1
      Top             =   90
      Width           =   4485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find:"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   345
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

