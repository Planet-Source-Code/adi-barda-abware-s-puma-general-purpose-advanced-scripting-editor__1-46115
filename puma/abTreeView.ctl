VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl abTreeView 
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "abTreeView.ctx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   4560
   ToolboxBitmap   =   "abTreeView.ctx":000C
   Begin MSComctlLib.TreeView tv1 
      Height          =   3285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5794
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imglst1 
      Left            =   0
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "abTreeView.ctx":031E
            Key             =   "FILE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "abTreeView.ctx":0770
            Key             =   "LINKFILE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "abTreeView.ctx":0BC2
            Key             =   "FOLDER"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "abTreeView.ctx":115C
            Key             =   "ROOT"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "abTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' Link to the wrapped tree view
Public SourceTreeView       As MSComctlLib.TreeView

' Define supported data sources
Public Enum DATA_SOURCE
    MSSQL = 0
    ORACLE = 1
End Enum

Public OracleSequence       As String ' In case of oracle data source - need a sequence
Private cn                  As ADODB.Connection

Private m_DataSourceType    As DATA_SOURCE
Private m_SelfConnected     As Boolean 'Determines whether uses outsource connection or self one
Private m_Table_Name        As String
Private m_ID_Field          As String
Private m_Father_Field      As String
Private m_Name_Field        As String

Private m_FirstHeader       As String
Private m_Hirarchy          As Long
Private Err_Handle_Mode     As Boolean
Private m_SpecialID         As Long
Private m_RootName          As String

Public Event NodeSelected(ByVal oNode As MSComctlLib.Node, ByVal ItemCode As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer, ActiveItem As String)
Public Event BeforeItemAdd(Cancel As Boolean)
Public Event AfterItemAdd(ByRef ItemNode As MSComctlLib.Node, ByRef rs As ADODB.Recordset)

'-Enable pop up menu
Private m_objMenu As VB.Menu

Public Property Let PopupMenuObject(ByVal mnu As VB.Menu)
    Set m_objMenu = mnu
End Property


Public Property Get RootName() As String
Attribute RootName.VB_ProcData.VB_Invoke_Property = "page1"
    RootName = m_RootName
End Property

Public Property Let RootName(ByVal newval As String)
    m_RootName = newval
End Property

Public Property Get Name_Field() As String
Attribute Name_Field.VB_ProcData.VB_Invoke_Property = "page1"
    Name_Field = m_Name_Field
End Property

Public Property Let Name_Field(ByVal str As String)
    m_Name_Field = str
End Property

Public Property Get FirstHeader() As String
Attribute FirstHeader.VB_ProcData.VB_Invoke_Property = "page1"
    FirstHeader = m_FirstHeader
'
End Property

Public Property Let FirstHeader(ByVal sHeader As String)
    m_FirstHeader = sHeader
'
End Property

Public Property Get CurrentKey() As String
    CurrentKey = Right$(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1)
'
End Property

Public Sub SetTableStructure(ByVal TableName As String, _
                             ByVal IDField As String, _
                             ByVal FatherField As String, ByVal NameField As String)
    
    '*Purpose: set the source table definition in order to manipulate the recursive functions
    
    m_Table_Name = TableName
    m_ID_Field = IDField
    m_Father_Field = FatherField
    m_Name_Field = NameField

End Sub

Public Sub Add_Branch(ByVal BranchName As String)

    On Error GoTo err_proc
    
    If Trim(BranchName) = "" Then Exit Sub
    
    ' Check if there is an oracle sequence
    If (DataSourceType = ORACLE) And (Trim(OracleSequence) = "") Then Exit Sub
    
    Dim nodx                        As MSComctlLib.Node
    Dim rcs                         As New ADODB.Recordset
    Dim b                           As Boolean
    Dim s                           As String
    Dim i                           As Long
    Dim sRelative                   As String
    Dim sNewItemKey                 As String
    Dim vHirarchyArray
    Dim iCurrentHirarchy            As Long
    Dim sKeys
    Dim iFather                     As Long
    Dim fCancel                     As Boolean
    
    fCancel = False
    
    sKeys = Array("ROOT", "FILE", "FILE", "FILE") 'picture keys
    iCurrentHirarchy = Val(tv1.SelectedItem.Tag)
    
    RaiseEvent BeforeItemAdd(fCancel)
    If fCancel Then Exit Sub
    
    
    '/// Adds the new item to the database ///
    iFather = Val(Right$(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1))
    
    
    ' Generate SQL script
    s = "INSERT INTO " & m_Table_Name & " ("
    
    ' Oracle doesnt have auto increment so have to use sequence on the ID field
    If DataSourceType = ORACLE Then
        s = s & m_ID_Field & "," ' Manuly add the ID field
    End If
    
    s = s & m_Father_Field & "," & m_Name_Field & ") "
    s = s & " Values ("
    
    ' If its oracle, insert the sequence as well
    If DataSourceType = ORACLE Then
        s = s & OracleSequence & ".nextval,"
    End If
    
    s = s & iFather & ",'" & BranchName & "')"
    cn.Execute s
    
    '/// Adds item to treeview
    If DataSourceType = MSSQL Then
        s = "SELECT TOP 1 * FROM " & m_Table_Name & " ORDER BY 1 DESC"
    Else
        s = "SELECT * FROM (SELECT * FROM " & m_Table_Name & " ORDER BY 1 DESC) WHERE ROWNUM=1"
    End If
    
    rcs.Open s, cn, adOpenForwardOnly, adLockReadOnly
    
    sRelative = tv1.SelectedItem.Key
    sNewItemKey = "K" & rcs(0).Value
    rcs.Close
    Set rcs = Nothing
    
    m_Hirarchy = Val(tv1.SelectedItem.Tag) + 1
    Set nodx = tv1.Nodes.Add(sRelative, tvwChild, sNewItemKey, BranchName)
    nodx.Tag = str$(m_Hirarchy)
    
    If Not nodx.Parent.Expanded Then
        nodx.Parent.Expanded = True
        nodx.Parent.Image = "FOLDER"
        nodx.Parent.Bold = True
    End If
    
    'place a picture
    i = m_Hirarchy - 1 'set pic index
    If UBound(sKeys) < i Then i = UBound(sKeys)
    If i = 0 Then i = 1
    nodx.Image = "FILE"
    nodx.Bold = False
    nodx.Selected = True
    
    'Force click on the new node
    tv1_NodeClick nodx
    
exit_proc:
    Exit Sub


err_proc:
    Err_Handler "abTreeView", "Add_Branch", Err, Err_Handle_Mode
    Resume exit_proc

End Sub

Public Sub DeleteCurrentNode(Optional ByVal NodeKey As String = "")


    On Error GoTo err_proc

    If NodeKey = "" Then
        NodeKey = Right$(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1)
    End If
    
    DeleteNode NodeKey
    
exit_proc:
Exit Sub


err_proc:
    Err_Handler "abTreeView", "Delete_Branch", Err, Err_Handle_Mode
Resume exit_proc


End Sub

Public Property Get CurrentNode() As Node
    
    Set CurrentNode = tv1.SelectedItem
    
End Property

Public Sub DeleteNode(ByVal NodeKey As String)


    On Error GoTo err_proc
    
    
    '//recoursive procedure to delete a specific level in the tree
    Dim rcs         As ADODB.Recordset
    Dim s           As String
    Dim sKey
    
    Set rcs = New ADODB.Recordset
    
    s = "SELECT * FROM " & m_Table_Name & " WHERE " & m_Father_Field & " = " & NodeKey
    rcs.Open s, cn, adOpenKeyset, adLockOptimistic
    If (Not rcs.EOF) And (Not rcs.BOF) Then
        Do Until rcs.EOF
            sKey = rcs(0).Value  'get the next level to delete
            rcs.Delete 'delete current item
            DeleteNode sKey  'delete the next level
            rcs.MoveNext
        Loop
    End If
    'delete the root level itself
    s = "DELETE FROM " & m_Table_Name & " WHERE " & m_ID_Field & " = " & NodeKey
    cn.Execute s
    
    UserControl.tv1.Nodes.Remove ("K" & NodeKey)
    
    rcs.Close
    Set rcs = Nothing
    
exit_proc:
    Exit Sub


err_proc:
    Err_Handler "abTreeView", "DeleteNode", Err, Err_Handle_Mode
Resume exit_proc


End Sub

Private Sub otv_BeforeAddItem(Cancel As Boolean)


    On Error GoTo err_proc

    RaiseEvent BeforeItemAdd(Cancel)
    
exit_proc:
Exit Sub


err_proc:
    Err_Handler "abTreeView", "otv_BeforeAddItem", Err, Err_Handle_Mode
Resume exit_proc


End Sub

Private Sub tv1_AfterLabelEdit(Cancel As Integer, NewString As String)


    On Error GoTo err_proc

    Dim sID          As String
    Dim s            As String
    
    If tv1.SelectedItem.Key = "ROOT" Then Exit Sub
    
    sID = Right$(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1)
    s = "UPDATE " & m_Table_Name & " SET " & m_Name_Field & " = '" & _
        LegalName(NewString) & "' " & _
        "WHERE " & m_ID_Field & " = " & sID
    
    cn.Execute s

exit_proc:
    Exit Sub


err_proc:
    Err_Handler "abTreeView", "tv1_AfterLabelEdit", Err, Err_Handle_Mode
    Resume exit_proc


End Sub


Private Sub tv1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    Dim oNode As MSComctlLib.Node
    
    If Not (m_objMenu Is Nothing) Then
        If Button = 2 Then
            Set oNode = tv1.HitTest(x, Y)
            If Not (oNode Is Nothing) Then
                oNode.Selected = True
                RaiseEvent NodeSelected(oNode, Right$(oNode.Key, Len(oNode.Key) - 1))
                PopupMenu m_objMenu
            End If
        End If
    End If
    
End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim sCode           As String
    
    sCode = Right$(Node.Key, Len(Node.Key) - 1)
    RaiseEvent NodeSelected(Node, sCode)
    
End Sub

Private Sub UserControl_Initialize()
    
    Set SourceTreeView = tv1
    tv1.ImageList = imglst1
    m_Hirarchy = 0
    m_RootName = "ROOT"
    
End Sub

Private Sub UserControl_InitProperties()

    FirstHeader = "GENERAL"
    DataSourceType = MSSQL
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sCode           As String
    
    
    sCode = Right$(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1)
    RaiseEvent KeyDown(KeyCode, Shift, sCode)
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_FirstHeader = PropBag.ReadProperty("FirstHeader", "GENERAL")
    m_ID_Field = PropBag.ReadProperty("ID_Field", "")
    m_Father_Field = PropBag.ReadProperty("Father_Field", "")
    m_Name_Field = PropBag.ReadProperty("Name_Field", "")
    m_Table_Name = PropBag.ReadProperty("Table_Name", "")
    m_DataSourceType = PropBag.ReadProperty("DataSourceType", 0)
    
End Sub

Private Sub UserControl_Resize()

    tv1.Width = Width
    tv1.Height = Height

End Sub

Public Property Get Table_Name() As String
Attribute Table_Name.VB_ProcData.VB_Invoke_Property = "page1"
    Table_Name = m_Table_Name
    
End Property

Public Property Let Table_Name(ByVal vNewValue As String)
    m_Table_Name = vNewValue
    
End Property

Public Sub Connect_To_Database(ByVal ConnectionString As String)


    On Error GoTo err_proc
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = ConnectionString
    cn.Open
    m_SelfConnected = True 'private connection
    
exit_proc:
    Exit Sub


err_proc:
    Err_Handler "abTreeView", "Connect_To_Database", Err, Err_Handle_Mode
    Resume exit_proc

End Sub

Public Sub Use_Connection(Source_Connection As ADODB.Connection)

    Set cn = Source_Connection
    m_SelfConnected = False 'not my connection
    
End Sub

Public Sub Build_Tree()

    On Error GoTo err_proc
    
    If Trim(Table_Name) = "" Then
        MsgBox "No table name has been chosen"
        Exit Sub
    End If
    
    InitTree Table_Name
    
exit_proc:
    Exit Sub


err_proc:
    Err_Handler "abTreeView", "Build_Tree", Err, Err_Handle_Mode
    Resume exit_proc

End Sub

Public Property Get Father_Field() As String
Attribute Father_Field.VB_ProcData.VB_Invoke_Property = "page1"
    Father_Field = m_Father_Field
End Property

Public Property Let Father_Field(ByVal vNewValue As String)
    m_Father_Field = vNewValue
End Property

Public Property Get ID_Field() As String
Attribute ID_Field.VB_ProcData.VB_Invoke_Property = "page1"
    ID_Field = m_ID_Field
End Property

Public Property Let ID_Field(ByVal vNewValue As String)
    m_ID_Field = vNewValue
End Property

Private Sub UserControl_Terminate()
    
    If m_SelfConnected Then 'case its my private connection
        If Not (cn Is Nothing) Then
            cn.Close
            Set cn = Nothing
        End If
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "FirstHeader", FirstHeader, "GENERAL"
    PropBag.WriteProperty "ID_Field", m_ID_Field
    PropBag.WriteProperty "Father_Field", m_Father_Field
    PropBag.WriteProperty "Name_Field", m_Name_Field
    PropBag.WriteProperty "Table_Name", m_Table_Name
    PropBag.WriteProperty "DataSourceType", m_DataSourceType
    
End Sub


Private Sub InitTree(ByVal Table_Name As String, Optional ByVal SpecialID As Long = 1)


    On Error GoTo err_proc
    
    Dim nodx                As MSComctlLib.Node
    Dim sFirstHeader        As String
    Dim iFather             As Long
    
    
    If cn Is Nothing Then
        MsgBox "There is no live connection to database"
        Exit Sub
    End If

    
    With UserControl.tv1
        .Style = tvwTreelinesPlusMinusPictureText
        '.LineStyle = tvwRootLines
        m_SpecialID = SpecialID
        m_Table_Name = Table_Name
    
        sFirstHeader = FirstHeader
        If Trim(sFirstHeader) = "" Then sFirstHeader = "GENERAL"
    
    
        Set nodx = .Nodes.Add(, , m_RootName, sFirstHeader)
        nodx.Tag = "0"
        nodx.Image = "ROOT"
        nodx.Bold = True
        LoadItemsEX iFather, m_RootName, 1
        
    End With
    
exit_proc:
    Exit Sub


err_proc:
    Err_Handler "CTv", "InitTree", Err, Err_Handle_Mode
    Resume exit_proc

End Sub

Private Sub LoadItemsEX(ByVal iFather As Long, _
                        ByVal sFatherKey As String, ByVal iLevelNum As Long)


    On Error GoTo err_proc

    Dim rcs             As New ADODB.Recordset
    Dim s               As String
    Dim i               As Long
    Dim nodx            As MSComctlLib.Node
    Dim sKey            As String
    Dim sKeys 'pictures
    
    sKeys = Array("FILE", "FILE", "FILE", "FILE") 'picture keys
    
    s = "SELECT * FROM " & m_Table_Name & " WHERE " & m_Father_Field & " = " & iFather
    rcs.Open s, cn, adOpenForwardOnly, adLockReadOnly
    
    With UserControl.tv1
        If Not rcs.EOF Then
            Do Until rcs.EOF
                sKey = "K" & rcs(0).Value
                Set nodx = .Nodes.Add(sFatherKey, tvwChild, sKey, rcs(m_Name_Field))
                nodx.Tag = str$(iLevelNum)
                
                If IsParentEX(rcs(m_ID_Field).Value) Then
                    nodx.Image = "FOLDER"
                    nodx.Bold = True
                    RaiseEvent AfterItemAdd(nodx, rcs)
                    LoadItemsEX rcs(m_ID_Field).Value, sKey, iLevelNum + 1
                Else
                    nodx.Image = "FILE"
                    nodx.Bold = False
                    RaiseEvent AfterItemAdd(nodx, rcs)
                End If
                rcs.MoveNext
            Loop
        End If
    End With
    
    rcs.Close
    Set rcs = Nothing
    
exit_proc:
    Exit Sub

err_proc:
    Err_Handler "CTv", "LoadItemsEX", Err, Err_Handle_Mode
    Resume exit_proc

End Sub

Private Sub Err_Handler(ByVal Module As String, ByVal Proc As String, ByVal oErr As ErrObject, Err_Handle_Mode)
    
    'update error log file
    'show error within msgbox only for debug purpose:
    MsgBox "error ocurred in module: " & Module & ", proc: " & Proc & _
            vbNewLine & Err.Description
    
End Sub

Private Function IsParentEX(ByVal iFather As Long) As Boolean


    On Error GoTo err_proc

    IsParentEX = False
    
    Dim rcs             As New ADODB.Recordset
    Dim s               As String
    
    
    s = "SELECT "
    If m_DataSourceType = MSSQL Then
        s = s & " TOP 1 Father "
    ElseIf m_DataSourceType = ORACLE Then
        s = s & " MAX(Father) "
    End If
    
    s = s & " FROM " & m_Table_Name & " WHERE " & m_Father_Field & " = " & iFather
    
    rcs.Open s, cn, adOpenForwardOnly, adLockReadOnly
    IsParentEX = (Not rcs.EOF)
    
    rcs.Close
    Set rcs = Nothing
    
exit_proc:
    Exit Function


err_proc:
    Err_Handler "abTreeView", "IsParentEX", Err, Err_Handle_Mode
    Resume exit_proc

End Function

Private Function LegalName(ByVal str As String) As String

    LegalName = Replace$(str, "'", "''")
    
End Function


Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property


Public Property Get DataSourceType() As DATA_SOURCE
    DataSourceType = m_DataSourceType
End Property

Public Property Let DataSourceType(ByVal vNewValue As DATA_SOURCE)
    m_DataSourceType = vNewValue
End Property
