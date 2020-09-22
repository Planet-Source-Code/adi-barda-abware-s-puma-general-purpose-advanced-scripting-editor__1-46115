VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'
'*CLASS: CXMLParser - Parse PMAPI strings to XML && Vise versa
'*       Written by adi barda - malam projects
'
'*NOTE: All the "path" params are XPath (XML search language) compatible

'Sca standart field seperators
Private Const SOH = &H1 '---------start of header
Private Const STX = &H2 '---------start of text
Private Const FS = &H1C '---------field seperator
Private Const ETX = &H3 '---------end of text
Private Const EOT = &H4 '---------end of transmission

Private m_DomDoc                As DOMDocument
Private m_iOptionBase           As Long ' XML items will be indexed from 0  OR  1
Private m_sError                As String ' Error description
Private Err_Handle_Mode         As Boolean ' Currently not in use
Private m_DefaultParentElement  As String

Option Explicit

Public Function SaveBinaryData(ByVal ItemPath As String, ByVal FilePath As String) As Boolean

    On Error GoTo err_proc
    
    Dim iFile               As Integer
    Dim arrBuffer()         As Byte
    Dim oNode               As IXMLDOMNode
      
      
    SaveBinaryData = False
    
    If Not (m_DomDoc Is Nothing) Then
        
        ' Get the data
        Set oNode = m_DomDoc.documentElement.selectSingleNode(ItemPath)

        ' Use a byte array
        arrBuffer = oNode.nodeTypedValue
            
        ' Write to disk
        iFile = FreeFile()
        Open FilePath For Binary Access Write As iFile
        Put iFile, , arrBuffer
        Close iFile
    
        'Succeeded
        SaveBinaryData = True
        
    End If
    
exit_proc:
    Set oNode = Nothing
    Exit Function
    
err_proc:
    Resume exit_proc

End Function


Public Function AddItem(ByVal Name As String, Optional ByVal Value As String = "", _
                        Optional ByVal ParentPath As String = "", _
                        Optional ByVal SetStructure As Boolean = False) As Boolean

    On Error GoTo err_proc
                        
    '*Purpose: Add new item to the XML string
    'Examples:
    '         1) AddItem ("adi") - adds one 'adi' item to the root element
    '         2) AddItem ("adi","handsome") - adds 'adi' item which contains 'handsome' value
    '         3) AddItem ("name","adi barda","adi[0]") - adds 'name' item which
    '            contains 'adi barda' value, to the FIRST 'adi' item
    '         4) AddItem ("name","yaron zehavi","Officers/member[2]") - adds 'name' item
    '            contains 'yaron zehavi' value, to the 3rd 'member' element which is in
    '            the 'Officers' element
    '         5) AddItem ("name","yaron zehavi","Officers[1]/member[2]") - adds
    '            'name' item, contains 'yaron zehavi' value, to the 3rd 'member' element
    '            which is in the second 'Officers' element
    
  Dim objParent       As IXMLDOMNode
  Dim objNode         As IXMLDOMNode
  Dim objDoc          As DOMDocument
  
    AddItem = False 'function failed by default
    
    'check dom validity
    If (m_DomDoc Is Nothing) Then Exit Function
    
    ' Check if the client wants to add an xml based element
    If SetStructure Then
        ' Create xml based node
        If LenB(Trim$(Name)) <> 0 Then Exit Function ' Must be null string
        ' Create new xml document
        Set objDoc = New DOMDocument
        If Not objDoc.loadXML(Value) Then
            ' Invalid xml - function aborted
            Exit Function
        End If
        
        ' Get the document's element
        Set objNode = objDoc.documentElement
        
    Else
        ' Create regular new node
        If LenB(Trim$(Name)) = 0 Then Exit Function ' Must have a string
        Set objNode = m_DomDoc.createNode(1, Name, "")
        objNode.nodeTypedValue = Value
    End If
    
    'get parent node
    ParentPath = Trim$(ParentPath)
    
    ' Add default parent element
    If Trim$(m_DefaultParentElement) <> "" Then
        If ParentPath <> "" Then
            ParentPath = m_DefaultParentElement & "/" & ParentPath
          Else
            ParentPath = m_DefaultParentElement
        End If
    End If
    
    If LenB(ParentPath) = 0 Then
        Set objParent = m_DomDoc.documentElement 'get the root element
      Else
        If m_iOptionBase = 1 Then ApplyOptionBase1 ParentPath
        Set objParent = m_DomDoc.documentElement.selectSingleNode(ParentPath)
    End If
    
    If (objParent Is Nothing) Then Exit Function 'check parent validity
    
    'add new item to the parent node
    objParent.insertBefore objNode, Null
    
    ' Call release
    Set objDoc = Nothing
    
    AddItem = True
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "AddItem", Err, Err_Handle_Mode
    Resume exit_proc

End Function

Private Function ApplyOptionBase1(ByRef path As String) As Boolean

    On Error GoTo err_proc

    '* Purpose: Convert XMP path to option base 1
    
  Dim i           As Long
  Dim i2          As Long
  Dim ival        As Long
  Dim bScan       As Boolean
    
    i = 1
    ApplyOptionBase1 = False
    bScan = True
    
    Do While bScan
    
        i = InStr(i, path, "[")
        bScan = (i <> 0)
        If bScan Then
            i2 = InStr(i + 1, path, "]")
            bScan = (i2 <> 0)
            If bScan Then
                ' Find current value
                ival = Val(Mid$(path, i + 1, i2 - i - 1))
                ' Decrease current value
                Mid$(path, i + 1, i2 - i - 1) = CStr(ival - 1) & Space$(Len(CStr(ival)) - Len(CStr(ival - 1)))
            End If
        End If
        i = i2 + 1
    Loop
    
    ApplyOptionBase1 = True
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "ApplyOptionBase1", Err, Err_Handle_Mode
    Resume exit_proc

End Function

Private Sub Class_Initialize()

    m_iOptionBase = 0 ' By default option base is 0
    m_sError = "" ' No errors yet
    
End Sub

Public Function CreateNewXML(Optional ByVal RootElement As String = "root") As Boolean

    On Error GoTo err_proc

    '*Purpose: Create new XML string
    
  Dim xmlPI       As IXMLDOMProcessingInstruction
  Dim root        As IXMLDOMNode
  Dim i
  Dim s           As String
    
    Set m_DomDoc = New DOMDocument
    
    'add process instruction:
    Set xmlPI = m_DomDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1255'")
    m_DomDoc.appendChild xmlPI
    
    'add root element:
    Set root = m_DomDoc.createElement(RootElement)
    Set m_DomDoc.documentElement = root
    CreateNewXML = Not (root Is Nothing) 'function's return value
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "CreateNewXML", Err, Err_Handle_Mode
    Resume exit_proc

End Function

Public Property Get DefaultParentElement() As String

    DefaultParentElement = m_DefaultParentElement

End Property

Public Property Let DefaultParentElement(ByVal NewVal As String)

    m_DefaultParentElement = NewVal

End Property

Public Function DeleteItem(ByVal Item As String, Optional ByVal path As String = "") As Boolean

  '* Purpose: Delete an item from the dom document
    
  '* Purpose: Set an item value
    
    On Error GoTo err_proc

    '*Explanations:
    '*              Item - the name of the item you wish to get from the XML doc
    '*              Path - optional - the path where the item is held
    '*Examples:
    '          1) SetItemValue("UNIT_NAME","new name")
    '          2) SetItemValue("UNIT_NAME","new name","MEMBER[2]/UNIT[4]") - gets the unit name from
    '                 - the  5th unit which is in the 3rd member
    
  Dim Element             As IXMLDOMNode
    
    DeleteItem = False
    ' Make XML query
    Item = Trim$(Item)
    If Item = "" Then Exit Function 'Item is not valid
    
    path = Trim$(path)
    
    ' Add default parent element
    If Trim$(m_DefaultParentElement) <> "" Then
        If path <> "" Then
            path = m_DefaultParentElement & "/" & path
          Else
            path = m_DefaultParentElement
        End If
    End If
    
    If path <> "" Then
        path = path & IIf(Right$(path, 1) = "/", "", "/")
        Item = path & Item
        If m_iOptionBase = 1 Then ApplyOptionBase1 Item
    End If
    
    ' Get element from doc tree
    If Not m_DomDoc.documentElement Is Nothing Then 'check dom document validity
        Set Element = m_DomDoc.documentElement.selectSingleNode(Item)
    End If
    
    ' Scan XML tree
    If Not (Element Is Nothing) Then
        Element.parentNode.removeChild Element
        DeleteItem = True
    End If
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "SetItemValue", Err, Err_Handle_Mode
    Resume exit_proc
    
End Function

Private Sub Err_Handler(ByVal Module As String, ByVal Proc As String, Err As ErrObject, Err_Handle_Mode As Boolean)

  ' Centeral error handling procedure
    
  ' Set error property to the current error

    m_sError = "çìä ùâéàä áîåãåì:" & Module & "    áôåð÷öéä:" & Proc & vbNewLine & _
               "úàåø äùâéàä: " & Err.Description
        
End Sub

Public Property Get Error() As String

    Error = m_sError
    m_sError = ""
    
End Property

Public Function FindItems(ByVal Item As String, ByVal Criteria As String, _
                          Optional ByVal path As String = "") As Variant

  '* Purpose: Return an variant array containing all the element numbers
  '           - that answer to the specified criteria
    
  Dim arrIndex()          As Variant
  Dim sXPath              As String
  Dim i                   As Long
  Dim iIndex              As Long
  Dim Element             As IXMLDOMNode
  Dim objList             As IXMLDOMNodeList
    
    ' Validate params
    Item = Trim$(Item): Criteria = Trim$(Criteria): path = Trim$(path)
    If LenB(Item) = 0 Or LenB(Criteria) = 0 Then Exit Function
    
    
    ' Add default parent element
    If LenB(Trim$(m_DefaultParentElement)) <> 0 Then
        If LenB(path) <> 0 Then
            path = m_DefaultParentElement & "/" & path
          Else
            path = m_DefaultParentElement
        End If
    End If
    
    ' Add slash char if needed
    If path <> "" And Right$(path, 1) <> "/" Then path = path & "/"
    
    ' Compose xpath
    sXPath = path & Item & "[" & Criteria & "]"
    
    ' Check if current document is valid
    If (m_DomDoc Is Nothing) Then Exit Function
    
    ' Get result list
    If m_iOptionBase = 1 Then ApplyOptionBase1 sXPath
    Set objList = m_DomDoc.documentElement.selectNodes(sXPath)
    If Not (objList Is Nothing) Then
    
        iIndex = 1
        
        ' Init return array
        ReDim arrIndex(0)
        
        ' Scan all elements && build the return array
        For i = 1 To objList.length
            
            ' Add the new record to the first cell
            arrIndex(0) = arrIndex(0) & objList.Item(i - 1).xml
            
            ' Add the new record as a new cell
            ReDim Preserve arrIndex(iIndex)
            arrIndex(iIndex) = objList.Item(i - 1).xml
            iIndex = iIndex + 1
            
        Next i
        
    End If
    
    ' NOTE:
    ' the returned array is biult in the above format:
    ' 1) then first cell (number 0) contains all the returned elements together
    ' 2) than, each next cell from number 1 to n, contains only one element
    ' So, based on this format, we can make new parser object based either on all
    '     the elements or just on one of them as we wish
    
    ' For Example:
    '
    '   Dim NewParser as New Center_Xml_parser
    '   Dim arr as Variant
    '
    '   arr = objParser.FindItems("child","id=3") ' Get all "child" elements containing id=3
    '   NewParser.SetXMLDoc arr(0) ' Parser based on all the elements
    '   NewParser.SetXMLDoc arr(1) ' Parser based on the first element
    
    ' Retval
    FindItems = arrIndex
    
End Function

Public Function GetItem(ByVal Item As String, _
                        Optional ByVal path As String = "", _
                        Optional ByVal GetStructure As Boolean = False) As String

    On Error GoTo err_proc

    '*Purpose: Retrieve item value
    '*Explanations:
    '*              Item - the name of the item you wish to get from the XML doc
    '*              Path - optional - the path where the item is held
    '*              GetStructure - indicates whether retval will be xml structure or plain value
    
    '*Examples:
    '          1) GetItem("UNIT_NAME")
    '          2) GetItem("UNIT_NAME","MEMBER[2]/UNIT[4]") - gets the unit name from
    '                 - the  5th unit which is in the 3rd member
    
  Dim Element             As IXMLDOMNode
    
    ' Make XML query
    Item = Trim$(Item)
    If LenB(Item) = 0 Then Exit Function 'Item is not valid
    
    path = Trim$(path)
    
    ' Add default parent element
    If LenB(Trim$(m_DefaultParentElement)) <> 0 Then
        If path <> "" Then
            path = m_DefaultParentElement & "/" & path
          Else
            path = m_DefaultParentElement
        End If
    End If
    
    If LenB(path) <> 0 Then
        path = path & IIf(Right$(path, 1) = "/", "", "/")
        Item = path & Item
        If m_iOptionBase = 1 Then ApplyOptionBase1 Item
    End If
    
    ' Get element from doc tree
    If Not m_DomDoc.documentElement Is Nothing Then 'check dom document validity
        Set Element = m_DomDoc.documentElement.selectSingleNode(Item)
    End If
    
    ' Scan XML tree
    If Not (Element Is Nothing) Then
        If GetStructure Then
            GetItem = Element.xml ' Return the first and only item
        Else
            GetItem = Element.nodeTypedValue ' Return the first and only item
        End If
    End If
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "GetItem", Err, Err_Handle_Mode
    Resume exit_proc

End Function


Public Function GetRecordCount(Optional ByVal RecordName As String = "", Optional ByVal ParentPath As String = "") As Long

    On Error GoTo err_proc

    '*Purpose: Retrieve the amount of records within some parent element
    '*Params:
    '       1) ParentPath - standard xPath syntax pointing to the parent path
    
  Dim objList         As IXMLDOMNodeList
  Dim sXP             As String
    
    ParentPath = Trim$(ParentPath)
    
    ' Add default parent element
    If LenB(Trim$(m_DefaultParentElement)) <> 0 Then
        If ParentPath <> "" Then
            ParentPath = m_DefaultParentElement & "/" & ParentPath
          Else
            ParentPath = m_DefaultParentElement
        End If
    End If
    
    ' Build xpath query
    sXP = IIf(LenB(Trim$(ParentPath)) = 0, "", Trim$(ParentPath) & "/") & RecordName
    
    ' Check if current document is valid
    If (m_DomDoc Is Nothing) Then Exit Function
    
    ' Get result list
    If m_iOptionBase = 1 Then ApplyOptionBase1 sXP
    Set objList = m_DomDoc.documentElement.selectNodes(sXP)
    If Not (objList Is Nothing) Then
        GetRecordCount = objList.length
    End If
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "GetRecordCount", Err, Err_Handle_Mode
    Resume exit_proc

End Function

Public Property Get OptionBase() As Long

    OptionBase = m_iOptionBase

End Property

Public Property Let OptionBase(ByVal NewVal As Long)

    m_iOptionBase = IIf(NewVal <> 0, 1, 0)

End Property

Public Function SetItemValue(ByVal Item As String, ByVal Value As String, _
                             Optional ByVal path As String = "") As Boolean

  '* Purpose: Set an item value
    
    On Error GoTo err_proc

    '*Explanations:
    '*              Item - the name of the item you wish to get from the XML doc
    '*              Path - optional - the path where the item is held
    '*Examples:
    '          1) SetItemValue("UNIT_NAME","new name")
    '          2) SetItemValue("UNIT_NAME","new name","MEMBER[2]/UNIT[4]") - gets the unit name from
    '                 - the  5th unit which is in the 3rd member
    
  Dim Element             As IXMLDOMNode
    
    SetItemValue = False
    ' Make XML query
    Item = Trim$(Item)
    If LenB(Item) = 0 Then Exit Function 'Item is not valid
    
    path = Trim$(path)
    
    ' Add default parent element
    If LenB(Trim$(m_DefaultParentElement)) <> 0 Then
        If path <> "" Then
            path = m_DefaultParentElement & "/" & path
          Else
            path = m_DefaultParentElement
        End If
    End If
    
    If LenB(path) <> 0 Then
        path = path & IIf(Right$(path, 1) = "/", "", "/")
        Item = path & Item
        If m_iOptionBase = 1 Then ApplyOptionBase1 Item
    End If
    
    ' Get element from doc tree
    If Not m_DomDoc.documentElement Is Nothing Then 'check dom document validity
        Set Element = m_DomDoc.documentElement.selectSingleNode(Item)
    End If
    
    ' Scan XML tree
    If Not (Element Is Nothing) Then
        Element.nodeTypedValue = Value
        SetItemValue = True
    End If
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "SetItemValue", Err, Err_Handle_Mode
    Resume exit_proc
    
End Function

Public Function SetXMLDoc(ByVal xml As String) As Boolean

    On Error GoTo err_proc

    '*Purpose: load xml string into XML Dom Document
    
    Set m_DomDoc = New DOMDocument
    With m_DomDoc
        SetXMLDoc = .loadXML(xml)
    End With
    
exit_proc:

Exit Function

err_proc:
    Err_Handler "CXMLParser", "SetXMLDoc", Err, Err_Handle_Mode
    Resume exit_proc

End Function

Public Property Get xml() As String

  '*Purpose: Return the current XML string
    
    If Not (m_DomDoc Is Nothing) Then
        xml = m_DomDoc.xml
    End If
    
End Property

':) Ulli's VB Code Formatter V2.6.10 (12/12/2001 16:39:12) 21 + 529 = 550 Lines
