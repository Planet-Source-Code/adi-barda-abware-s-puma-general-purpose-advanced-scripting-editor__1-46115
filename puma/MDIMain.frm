VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "P-U-M-A"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load script"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuReloadLinkedDocs 
         Caption         =   "Reload all linked documents"
      End
      Begin VB.Menu mnuObjectBrowser 
         Caption         =   "Object browser..."
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpPuma 
         Caption         =   "Puma enterprize"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuReloadDoc 
         Caption         =   "Reload Linked Document..."
      End
      Begin VB.Menu mnuPopupSetLinkedDoc 
         Caption         =   "Set linked document..."
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()

    Me.Caption = "P-U-M-A ( " & App.Major & "." & App.Minor & "." & App.Revision & " )"
    Set g_DB = New CDB
    g_DB.Init
    frmScripts.Show
exit_proc:
    Exit Sub


End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuScriptManager_Click()
    
    frmScripts.Show
exit_proc:
    Exit Sub


End Sub

Private Sub mnuFind_Click()
    
    frmScripts.rtfEditor.FindText frmScripts.rtfEditor.SelText
    
End Sub

Private Sub mnuLoad_Click()
    frmScripts.cmdLoadData.value = True
End Sub

Private Sub mnuObjectBrowser_Click()
    frmObjectBrowser.Show 1
End Sub

Private Sub mnuPopupSetLinkedDoc_Click()

    frmScripts.cmdLoadData.value = True

End Sub

Private Sub mnuReloadDoc_Click()

    'load the current node's linked doc
    frmScripts.rtfEditor.ReloadLinkedDoc
     
End Sub

Private Sub mnuReloadLinkedDocs_Click()

    frmScripts.rtfEditor.ReloadDocument -1
    
End Sub

Private Sub mnuSave_Click()
    frmScripts.cmdSave.value = True
End Sub

Private Sub mnuSaveAs_Click()
    frmScripts.cmdSaveToFile.value = True
End Sub
