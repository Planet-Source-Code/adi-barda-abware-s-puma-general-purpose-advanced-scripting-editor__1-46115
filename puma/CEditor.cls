VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_USER As Long = &H400
Private Const CFM_BACKCOLOR = &H4000000
Private Const EM_GETCHARFORMAT As Long = (WM_USER + 58)
Private Const EM_SETCHARFORMAT As Long = (WM_USER + 68)
Private Const SCF_SELECTION = &H1&
Private Const LF_FACESIZE As Integer = 32


Private Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60

    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type


Private Const EM_SCROLL = &HB5
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_CHARFROMPOS = &HD7

Private Const EM_GETTEXTRANGE = 1024 + 75
Private Const EM_GETOPTIONS = 1024 + 78

Private Type InterfaceData
    Class As String
    method As String
    IsMethod As Boolean
    ToolTip As String
End Type
Private m_Interface() As InterfaceData

Const abGREEN = 32768

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type EditorWords
    color As Long
    wrd As String
End Type
Private m_arrWords() As EditorWords

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageP Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As POINTAPI) As Long

Private m_bPaintText As Boolean

Private m_iFunctionEndPoint As Long

Private m_bTextChange As Boolean

Private m_sScriptText As String
Private m_iPicHeight As Long
Private m_picToolTip As VB.PictureBox
Private m_Img(1) As stdole.StdPicture
Private WithEvents m_txtScript As RichTextLib.RichTextBox
Attribute m_txtScript.VB_VarHelpID = -1
Private m_DebugTextbox As RichTextLib.RichTextBox
Private WithEvents m_flx As MSComctlLib.ListView
Attribute m_flx.VB_VarHelpID = -1

Private m_iMinPos As Long
Private m_iMaxPos As Long

Private m_arrKeywrds() As String

Private m_iLastHighLightStart As Long
Private m_iLastHighlightLen As Long

Public Sub HighLight(ByVal color As Long)

    Dim iStart As Long
    Dim iLen As Long
    
    
    If m_iLastHighLightStart > 0 Then
    
        iStart = m_txtScript.SelStart
        iLen = m_txtScript.SelLength
        
        m_txtScript.SelStart = m_iLastHighLightStart
        m_txtScript.SelLength = m_iLastHighlightLen
        APIHighlight m_txtScript.BackColor
        
        m_txtScript.SelStart = iStart
        m_txtScript.SelLength = iLen
        
    End If
    
    
    APIHighlight color
    m_txtScript.SelLength = 0
    
End Sub

Private Sub APIHighlight(ByVal color As Long)


  Dim Ret As Long, tmpcolr As Long
  Dim cf As CHARFORMAT2

    m_iLastHighLightStart = m_txtScript.SelStart
    m_iLastHighlightLen = m_txtScript.SelLength
    'if you use a Public Variable to store this you have to reset any change or
    'it will still be set when you make a different call. EG APIHighLight followed by APIBold
    ' both Bolds and Highlights the second instance, If you want to build up a style with one clcik
    'this might be useful but mostly not such a good thing.
    On Error GoTo oops
    With cf
        .cbSize = LenB(cf) 'setup the size of the character format
        .dwMask = CFM_BACKCOLOR 'what to test
        If .crBackColor = 0 Then
            .crBackColor = color
        Else 'NOT .CRBACKCOLOR...
            .crBackColor = m_txtScript.BackColor
        End If
        Ret = SendMessage(m_txtScript.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, cf)
    End With 'CF

Exit Sub

oops:
    MsgBox Err.Description
    

End Sub



Public Function GetKeywords() As String()
    GetKeywords = m_arrKeywrds()
End Function

Public Function GetXMLInterface() As String

    Dim sXML As String
    Dim i As Long
    '
    sXML = "<ROOT>"
    For i = 0 To UBound(m_Interface)
        
        With m_Interface(i)
            sXML = sXML & "<NEWITEM>"
            sXML = sXML & "<Class>" & .Class & "</Class>"
            sXML = sXML & "<IsMethod>" & .IsMethod & "</IsMethod>"
            sXML = sXML & "<method>" & .method & "</method>"
            sXML = sXML & "<ToolTip>" & .ToolTip & "</ToolTip>"
            sXML = sXML & "</NEWITEM>"
        End With
        
    Next i
    sXML = sXML & "</ROOT>"
    GetXMLInterface = sXML
    
End Function


Public Function GetCurrentPosition(mRTF As RichTextBox) As Long

    GetCurrentPosition = SendMessage(mRTF.hWnd, EM_GETTEXTRANGE, ByVal 0&, ByVal 0&)
    
End Function


Public Sub AddEditorWord(ByVal wrd As String, ByVal color As Long)

    Dim i  As Long
    
    i = UBound(m_arrWords())
    
    '''wrd = Trim$(wrd)
    If m_arrWords(0).wrd = "" Then
        i = 0
    Else
        i = i + 1
    End If
    
    ReDim Preserve m_arrWords(i)
    m_arrWords(i).wrd = wrd
    m_arrWords(i).color = color
    
    
End Sub

Public Function LoadIntelisence(ByVal InterfaceFile As String) As Boolean

    Dim ff As Long
    Dim s As String
    Dim arr As Variant
    Dim sToolTip As String
    Dim bMethod As Boolean
    
        
    ff = FreeFile
    Open InterfaceFile For Input As #ff
    Do Until EOF(ff)
        Line Input #ff, s
        arr = Split(s, ";")
            
        If UBound(arr) >= 2 Then
            bMethod = CBool(arr(2) = "1")
        Else
            bMethod = False
        End If
                   
        If UBound(arr) >= 3 Then
            sToolTip = arr(3)
        Else
            sToolTip = ""
        End If
        
        AddInterface arr(0), arr(1), bMethod, sToolTip
        
    Loop
    Close #ff
    
End Function


Private Sub AddInterface(ByVal cls As String, _
                         ByVal method As String, _
                         Optional ByVal IsMethod As Boolean = False, _
                         Optional ByVal ToolTip As String = "")

    
    Dim i As Long
    
    i = UBound(m_Interface)
    If m_Interface(i).Class <> "" Then
        i = i + 1
    End If
    
    ReDim Preserve m_Interface(i)
    
    With m_Interface(i)
    
        .Class = cls
        .method = method
        .IsMethod = IsMethod
        .ToolTip = ToolTip
        
        AddClassKeyword LCase$(cls)
        
    End With
    
End Sub

Private Sub AddClassKeyword(ByVal cls As String)

    Dim i As Long
    Dim bFound As Boolean
    
    bFound = False
    For i = 0 To UBound(m_arrKeywrds)
        bFound = m_arrKeywrds(i) = cls
        If bFound Then
            Exit For
        End If
    Next i
    
    If Not bFound Then
        i = UBound(m_arrKeywrds)
        If m_arrKeywrds(i) <> "" Then
            i = i + 1
        End If
        
        ReDim Preserve m_arrKeywrds(i)
        m_arrKeywrds(i) = cls
    End If
    
End Sub

Public Sub SetEditorObjects(ByRef objRTF As RichTextLib.RichTextBox, _
                           ByRef objGrid As MSComctlLib.ListView, _
                           ByRef img0 As stdole.StdPicture, _
                           ByRef img1 As stdole.StdPicture, _
                           ByRef picToolTip As VB.PictureBox, _
                           ByRef DebugTextbox As RichTextLib.RichTextBox)
                           
    'Init editor objects
    Set m_txtScript = objRTF 'main script text box
    Set m_DebugTextbox = DebugTextbox 'need this for smooth painting
    Set m_flx = objGrid 'intelisence grid
    
    m_flx.ColumnHeaders.Add 1, , , m_flx.Width - 3 * Screen.TwipsPerPixelX
    m_flx.FullRowSelect = True
    
    Set m_Img(0) = img0 'intelisence pictures
    Set m_Img(1) = img1 '"
    Set m_picToolTip = picToolTip 'tool tip object
    
    m_iPicHeight = m_picToolTip.Height
    ClearScript
    
End Sub

Public Sub ShowIntelisence()

    Dim sKeyword As String
    
    sKeyword = GetKeyword()
    If sKeyword <> "" Then
        ShowInterface sKeyword
    End If

End Sub

Private Sub ShowInterface(ByVal Keywrd As String)

    Dim i As Long
    Dim lpPoint As POINTAPI
    
    If m_Interface(0).Class = "" Then
        Exit Sub
    End If
    
    With m_flx
    
        .ListItems.Clear
        '.Rows = 0
        Dim objItem As MSComctlLib.ListItem
        
        For i = 0 To UBound(m_Interface)
            If LCase$(m_Interface(i).Class) = LCase$(Keywrd) Then
            
                Set objItem = .ListItems.Add(, , m_Interface(i).method)        ', , SmallIcon
                objItem.Tag = m_Interface(i).ToolTip
                objItem.SmallIcon = IIf(m_Interface(i).IsMethod, "Method", "Property")
                
            End If
        Next i
                
            
        GetCaretPos lpPoint
        .Left = m_txtScript.Left + (lpPoint.x + 10) * Screen.TwipsPerPixelX
        If .Left + .Width > Screen.Width Then
            .Left = Screen.Width - .Width
        End If
        
        .Top = m_txtScript.Top + (lpPoint.Y + 5) * Screen.TwipsPerPixelY
        
        'Set grid height
         
        If .ListItems.Count >= 6 Then
            .Height = 6 * .ListItems.Item(1).Height
        Else
            .Height = .ListItems.Count * .ListItems.Item(1).Height + 50
        End If
        
        
        .ListItems.Item(1).Selected = True
        
        .Visible = True
        
        .SetFocus
        
        ShowToolTip
        
    End With
    
End Sub

Public Sub ShowToolTip()
    
    Dim sToolTip As String
    Dim iCount As Long
    Dim iPos As Long
    Dim i As Long
    Dim iDelta As Long
    Dim bStop As Boolean
    
    With m_flx
    
        'Show tool tip
        sToolTip = " " & .SelectedItem.Tag
        
        m_picToolTip.Left = .Left
        m_picToolTip.Top = .Top + .Height
        m_picToolTip.Height = m_iPicHeight
        m_picToolTip.Width = m_picToolTip.TextWidth(sToolTip)
        
        'Insert vbNewLine where needed
        iDelta = Screen.Width - (.Left + m_picToolTip.Width)
        If iDelta < 0 Then
            sToolTip = GetFixedToolTip(sToolTip, iDelta, .Left)
        End If
            
        
        If Trim$(sToolTip) <> "" Then
        
            iPos = 0
            iCount = 0
            Do
                iPos = InStr(iPos + 1, sToolTip, vbNewLine)
                If iPos <> 0 Then
                    iCount = iCount + 1
                End If
                
            Loop Until iPos = 0
            
            If iCount > 0 Then
                For i = 1 To iCount
                    m_picToolTip.Height = m_picToolTip.Height + m_iPicHeight
                Next i
            End If
            
            m_picToolTip.Width = m_picToolTip.TextWidth(sToolTip)
            
            If m_picToolTip.Width + .Left > Screen.Width Then
                m_picToolTip.Left = Screen.Width - m_picToolTip.Width
            End If
            
            m_picToolTip.cls
            m_picToolTip.Print sToolTip
            
            m_picToolTip.Visible = True
        Else
            m_picToolTip.Visible = False
        End If
        
    End With
    
End Sub

Private Function GetFixedToolTip(ByVal sToolTip As String, _
                                 ByVal iDelta As Long, ByVal StartPos As Long) As String

    Dim i As Long
    Dim s As String
    Dim bPutChar As Boolean
    Dim bStop As Boolean
    Dim iLastPos As Long
    Dim iLastCut As Long
    Dim sCopy As String
    Dim bHardCut As Boolean
    
    bPutChar = False
    bStop = False
    bHardCut = False
    
    sCopy = sToolTip
    sToolTip = StrReverse(sToolTip)
    iDelta = Abs(iDelta)
    
    
    Do
    
        i = 0
        iLastPos = 0
        iDelta = Abs(iDelta)
        
        Do
            i = InStr(i + 1, sToolTip, ",")
            If i > 0 Then
                If iLastCut = i Then
                    i = iLastPos
                    bHardCut = True
                End If
                
                iLastPos = i
                
                If i = 0 Then
                    bStop = True
                Else
                    s = Left$(sToolTip, i)
                    If ((m_picToolTip.TextWidth(s) >= iDelta) Or (bHardCut)) Then
                        bHardCut = False
                        'OK this is the correct place to put the vbNewline
                        sCopy = Left$(sCopy, Len(sCopy) - i) & vbNewLine & Right$(sCopy, i)
                        iLastCut = i
                        bPutChar = True
                        
                        iDelta = Screen.Width - (m_picToolTip.TextWidth(s) + StartPos)
                        
                        bStop = CBool(iDelta >= 0)
                        i = 0
                    End If
                End If
            End If
            
        Loop Until i = 0
        
        'If we deednt put anything, put it in the last position found
        If (i = 0) And (Not bPutChar) Then
        
            If iLastPos = 0 Then
                bStop = True
            Else
                sCopy = Left$(sCopy, Len(sCopy) - iLastPos) & vbNewLine & Right$(sCopy, iLastPos)
                iDelta = Screen.Width - (m_picToolTip.TextWidth(Left$(sCopy, Len(sCopy) - iLastPos)) + StartPos)
                bStop = CBool(iDelta >= 0)
                
                sToolTip = Right$(sToolTip, Len(sToolTip) - iLastPos)
            End If
            
        End If
    
    Loop Until bStop
    
        
    GetFixedToolTip = sCopy
    
End Function

Private Function GetKeyword() As String

    Dim i As Long
    Dim iPos As Long
    Dim sKeywrd As String
    Dim s As String
    
    With m_txtScript
    
        iPos = .SelStart
        
        For i = iPos To 1 Step -1
            s = Mid$(.text, i, 1)
            
            Select Case Asc(s)
                Case 97 To 122
                Case 65 To 90
                Case 48 To 57
                Case Else
                    Exit For
            End Select
            
            sKeywrd = sKeywrd & s
        Next i
        
    End With
    
    sKeywrd = StrReverse(sKeywrd)
    If Not KeyWordExsist(sKeywrd) Then
        sKeywrd = ""
    End If
    
    'retval
    GetKeyword = sKeywrd
    
End Function

Private Function KeyWordExsist(ByVal Keywrd As String) As Boolean

    Dim arrKeywrds() As Variant
    Dim i As Long
    
    
    KeyWordExsist = False
    
    Keywrd = LCase$(Keywrd)
    
    
    For i = 0 To UBound(m_arrKeywrds)
        KeyWordExsist = CBool(m_arrKeywrds(i) = Keywrd)
        If KeyWordExsist Then
            Exit For
        End If
    Next i
    
    
End Function


Public Sub PaintText()

    Dim iPos As Long
    Dim i As Long
    Dim iCnt As Long
    Dim iBound As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sWorkText As String
    Dim iTextStart As Long
    Dim iTextEnd As Long
    
   
    iCnt = 0
    iStart = 0
    iPos = m_txtScript.SelStart
        
       
    m_sScriptText = m_txtScript.TextRTF
    sWorkText = m_sScriptText
        
        
    iBound = UBound(m_arrWords())
    For i = 0 To iBound
    
        If LenB(m_arrWords(i).wrd) = 0 Then
            Exit For
        End If
        
        
        'Is it the last word to paint ?
        If iBound = i Then
            'Yes it is
            ColorWord m_txtScript, iStart, iCnt, m_arrWords(i).color
        Else
            If m_arrWords(i + 1).color <> m_arrWords(i).color Then
                
                ColorWord m_txtScript, iStart, iCnt, m_arrWords(i).color
                iStart = i + 1
                iCnt = 0
                
            Else
                iCnt = iCnt + 1
            End If
        End If
        
        
        
    Next i
    
    
    'vb script or sql?
    If InStr(1, Trim$(LCase$(m_txtScript.text)), "#vbs@@@") > 0 Then
        
        ColorStrings m_txtScript, Chr$(34), Chr$(34), vbMagenta
        ColorRow m_txtScript, "'", abGREEN
    Else
        ColorStrings m_txtScript, "/*", "*/", abGREEN
        ColorStrings m_txtScript, "'", "'", vbMagenta
        ColorRow m_txtScript, "--", abGREEN
    End If
    
    If iTextStart = 0 Then
        
        m_txtScript.TextRTF = m_sScriptText
    Else
    
        m_txtScript.TextRTF = Replace$(m_txtScript.TextRTF, sWorkText, DelColorTable(m_sScriptText))
    End If
    
    
    If iPos > 0 Then
        iPos = InStr(iPos, m_txtScript.text, Chr$(10))
    End If
    
    m_txtScript.SelStart = iPos
    
    
    m_iMinPos = m_txtScript.GetLineFromChar(iPos)
    m_iMaxPos = m_iMinPos
    
    m_txtScript_SelChange
    
End Sub

Public Sub ColorWord(ByVal objRTF As RichTextLib.RichTextBox, _
                     ByVal StartIndx As Long, _
                     ByVal IndxCount As Long, _
                     ByVal color As Long)

    Dim i As Long
    Dim iStart As Long
    Dim sRTF As String
    Dim iColor As Long
    Dim iVeryStart As Long
    Dim iAsc As Long
    
        
    With objRTF
    
        sRTF = m_sScriptText '.TextRTF
        iVeryStart = InStr(1, sRTF, "}}")
        
        
        For i = StartIndx To StartIndx + IndxCount
            
            
            iStart = InStr(1, sRTF, m_arrWords(i).wrd, vbTextCompare)
            
            
            If iStart > 0 And iStart > iVeryStart Then
            
                'Check && update the color table
                If iColor = 0 Then
                    iColor = SetColorTable(sRTF, color)
                End If
                
                If Right$(m_arrWords(i).wrd, 1) = "." Then
                    sRTF = Replace$(sRTF, " " & m_arrWords(i).wrd, " \cf" & iColor & " " & m_arrWords(i).wrd & "\cf0 ", , , vbTextCompare)
                Else
                    'Color the word
                    sRTF = Replace$(sRTF, " " & m_arrWords(i).wrd & " ", " \cf" & iColor & " " & m_arrWords(i).wrd & " \cf0 ", , , vbTextCompare)
                    
                    'Color the word + vbnewline
                    sRTF = Replace$(sRTF, " " & m_arrWords(i).wrd & vbNewLine, " \cf" & iColor & " " & m_arrWords(i).wrd & " \cf0", , , vbTextCompare)
                
                    sRTF = Replace$(sRTF, vbNewLine & m_arrWords(i).wrd & " ", vbNewLine & " \cf" & iColor & " " & vbNewLine & m_arrWords(i).wrd & " \cf0", , , vbTextCompare)
                End If
                    
            Else
            
                If iStart > 0 Then
                    If Mid$(sRTF, iStart - 1, 1) <> " " Or Mid$(sRTF, iStart + 1, 1) <> " " Then
                        'Check && update the color table
                        If iColor = 0 Then
                            iColor = SetColorTable(sRTF, color)
                        End If
                        
                        'Color the word
                        sRTF = Replace$(sRTF, " " & m_arrWords(i).wrd & " ", " \cf" & iColor & " " & m_arrWords(i).wrd & " \cf0 ", , , vbTextCompare)
                        
                        'Color the word + vbnewline
                        sRTF = Replace$(sRTF, " " & m_arrWords(i).wrd & vbNewLine, " \cf" & iColor & " " & m_arrWords(i).wrd & " \cf0", , , vbTextCompare)
                    
                        sRTF = Replace$(sRTF, vbNewLine & m_arrWords(i).wrd & " ", vbNewLine & " \cf" & iColor & " " & vbNewLine & m_arrWords(i).wrd & " \cf0", , , vbTextCompare)
                    End If
                End If
                
            End If
                
            
        Next i
        
        
        m_sScriptText = sRTF '.TextRTF = sRTF
        
    End With
    
End Sub

Private Function GetColorTable() As String

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long

    'Do we have color table?
    iStart = InStr(1, m_txtScript.TextRTF, "colortbl", vbTextCompare)
    If iStart > 0 Then
        iStart = InStr(1, m_txtScript.TextRTF, "{\colortbl", vbTextCompare)
        iEnd = InStr(iStart, m_txtScript.TextRTF, ";}", vbTextCompare)
        GetColorTable = Mid$(m_txtScript.TextRTF, iStart, iEnd - iStart + 2)
    End If

End Function

Private Function DelColorTable(ByVal ScriptText As String) As String

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long

    'Do we have color table?
    iStart = InStr(1, ScriptText, "colortbl", vbTextCompare)
    If iStart > 0 Then
        iStart = InStr(1, ScriptText, "{\colortbl", vbTextCompare)
        iEnd = InStr(iStart, ScriptText, ";}", vbTextCompare)
        DelColorTable = Replace$(ScriptText, Mid$(ScriptText, iStart, iEnd - iStart + 2), "")
    End If

End Function

Private Function SetColorTable(ByRef rtf As String, _
                               ByVal color As Long) As Long

    Dim iR As Long
    Dim iG As Long
    Dim iB As Long
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sTable As String
    Dim sTmp As String
    Dim sColor As String
    
    GetRGB color, iR, iG, iB
    sColor = "\red" & iR & "\green" & iG & "\blue" & iB & ";"
    
    'Do we have color table?
    iStart = InStr(1, rtf, "colortbl", vbTextCompare)
    If iStart = 0 Then
        'No, Insert color table
        rtf = Replace$(rtf, "}}", "}}{\colortbl ;}", , 1)
        iStart = InStr(1, rtf, "colortbl", vbTextCompare)
    End If
    
    'Do we have the current color?
    If InStr(1, rtf, sColor, vbTextCompare) = 0 Then
        'No, insert color
        i = InStr(1, rtf, "\colortbl")
        i = InStr(i, rtf, "}", vbTextCompare)
        sTmp = Left$(rtf, i - 1)
        rtf = Replace$(rtf, "}", sColor & "}", i, 1, vbTextCompare)
        rtf = sTmp & rtf
    End If
    
    iEnd = InStr(iStart, rtf, ";}")
    sTable = Mid$(rtf, iStart, iEnd - iStart + 1)
    
    i = 0
    iStart = 1
    Do
        iStart = InStr(iStart, sTable, ";")
        If iStart <> 0 Then
            i = i + 1
            iStart = iStart + 1
            If InStr(iStart, sTable, sColor, vbTextCompare) = iStart Then
                SetColorTable = i
                Exit Do
            End If
        End If
    Loop Until iStart = 0
    
End Function

Private Sub GetRGB(ByVal color As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    
    b = Int(color / 65536)
    g = Int((color / 65536 - b) * 65536 / 256)
    r = Int(((color / 65536 - b) * 65536 / 256 - g) * 256)
    
End Sub

Private Sub ColorRow(ByVal objRTF As RichTextLib.RichTextBox, ByVal wrd As String, ByVal color As Long)

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sRTF As String
    Dim sRow As String
    Dim sNewRow As String
    Dim iColor As Long
    Dim iCommaPlace As Long
    
    With objRTF
        
        sRTF = m_sScriptText '.TextRTF
        
        iStart = 1
        Do
            iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
            If iStart > 0 Then
            
                'Check if its hebrew letter
                If Mid$(sRTF, iStart - 1, 1) <> "\" Then
                
                    'Check && update the color table
                    If iColor = 0 Then
                        iColor = SetColorTable(sRTF, color)
                    End If
                    
                    iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
                    iEnd = InStr(iStart, sRTF, "\par")
                    iCommaPlace = InStr(iStart, sRTF, Chr$(34))
                    
                    If iEnd > 0 Then 'And (iCommaPlace = 0 Or iCommaPlace > iEnd)
                        sRow = Mid$(sRTF, iStart, iEnd - iStart)
                        sNewRow = "\cf" & iColor & " " & GetCleanRow(sRow) & "\cf0 "
                        
                        sRTF = Left$(sRTF, iStart - 1) & _
                        Replace$(sRTF, sRow, sNewRow, iStart, 1)
                        
                        iEnd = InStr(iEnd - (Len(sRow) - Len(sNewRow)), sRTF, "\par")
                    End If
                    
                    
                    iStart = iEnd
                    
                Else
                
                    iStart = iStart + 1
                    
                End If 'Not hebrew letter
                
            End If 'iStart>0
            
        Loop Until iStart = 0
        
        m_sScriptText = sRTF '.TextRTF = sRTF
        
    End With
    

End Sub

Private Sub ColorStrings(ByVal objRTF As RichTextLib.RichTextBox, _
                         ByVal wrd As String, _
                         ByVal EndWrd As String, _
                         ByVal color As Long)

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim iBreak As Long
    Dim sRTF As String
    Dim sRow As String
    Dim sNewRow As String
    Dim m_sLastNewRow As String
    Dim iColor As Long
    
    If EndWrd = "" Then
        EndWrd = wrd
    End If
    
    With objRTF
        
        sRTF = m_sScriptText '.TextRTF
        
        iStart = 1
        Do
            iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
            If iStart > 0 Then
            
                'Check && update the color table
                If iColor = 0 Then
                    iColor = SetColorTable(sRTF, color)
                End If
                
                iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
                            
                If Mid$(sRTF, iStart - 1, 1) = "\" Then
                    iStart = iStart + 1
                Else
                
                    iEnd = GetStringEnd(sRTF, iStart + 1, EndWrd) 'InStr(iStart + 1, sRTF, EndWrd, vbTextCompare)
                    iBreak = InStr(iStart + 1, sRTF, vbNewLine)
                    
                    
                    If (iEnd > 0) And ((iBreak > iEnd) Or (iBreak = 0)) Then
                        
                        sRow = Mid$(sRTF, iStart, iEnd + Len(EndWrd) - iStart)
                        sNewRow = GetCleanRow(sRow)
                        m_sLastNewRow = sNewRow
                        
                        sRTF = Left$(sRTF, iStart - 1) & "\cf" & iColor & " " & sNewRow & "\cf0 " & Right$(sRTF, Len(sRTF) - iStart - Len(sRow) + 1)
                    
                        iEnd = InStr(iEnd + Len("\cf" & iColor & " " & sNewRow & "\cf0") - Len(sNewRow), sRTF, wrd)
                        
                    End If
                    
                    iStart = iEnd
                    
                End If
                
                
            End If 'iStart>0
            
        Loop Until iStart = 0
        
        m_sScriptText = sRTF '.TextRTF = sRTF
        
    End With
    

End Sub

Private Function GetStringEnd(ByVal sRTF As String, ByVal iStart As Long, ByVal EndWrd As String) As Long

    Dim bStop As Boolean
    
    'return the string end position?
    
    bStop = False
    Do Until bStop
        GetStringEnd = InStr(iStart, sRTF, EndWrd, vbTextCompare)
        'first, find the string limit character (' or ")
        If GetStringEnd > 0 Then
            'find one, now check for backslash
            bStop = CBool(Mid$(sRTF, GetStringEnd - 1, 1) <> "\")
            'is there a backslash
            If Not bStop Then
                'yes, is there another one before this backslash?
                'if yes, then stop the search. if no, continue
                bStop = CBool(Mid$(sRTF, GetStringEnd - 2, 1) = "\")
            End If
        Else
            bStop = True
        End If
        
        If Not bStop Then
            iStart = GetStringEnd + 1
        End If
        
    Loop
    
End Function

Private Function GetCleanRow(ByVal Row As String) As String

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sTmp As String
    
    GetCleanRow = Row
    iStart = 1
    Do
        iStart = InStr(iStart, Row, "\cf")
        If iStart > 0 Then
            iEnd = InStr(iStart, Row, " ")
            If iStart > 0 Then
                If iEnd = 0 Then
                    iEnd = Len(Row) + 1
                Else
                    iEnd = iEnd + 1
                End If
                sTmp = Mid$(Row, iStart, iEnd - iStart)
                
                'Ignore special RTF chars
                Select Case Mid$(sTmp, 1, 2)
                    Case "\\", "\{", "\}", "\'"
                        iStart = iStart + 2
                    Case Else
                        Row = Replace$(Row, sTmp, "")
                End Select
                
            End If
        End If
        
    Loop Until iStart = 0
    
    GetCleanRow = Row
    
End Function

Public Sub ClearScript()

    m_txtScript.text = ""

End Sub

Public Sub TypeIntelisenceText()

    m_txtScript.SelText = m_flx.SelectedItem.text '   m_flx.TextMatrix(m_flx.Row, 1)
    m_txtScript.SetFocus
    m_flx.Visible = False
    
    'Copy the function position on the screen
    m_iFunctionEndPoint = m_txtScript.SelStart ' + Len(m_flx.TextMatrix(m_flx.Row, 1))
    
    'm_picToolTip.Visible = False
    
End Sub

Private Sub Class_Initialize()
    
    m_bPaintText = False
    
    m_bTextChange = False
    
    ReDim m_arrWords(0)
    ReDim m_Interface(0)
    ReDim m_arrKeywrds(0)
    
End Sub


Private Sub m_flx_DblClick()
    TypeIntelisenceText
End Sub

Private Sub m_flx_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowToolTip
End Sub

Private Sub m_flx_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn, vbKeySpace
            TypeIntelisenceText
            SendKeys (" ")
        Case vbKeyEscape
            m_flx.Visible = False
            m_picToolTip.Visible = False
            m_txtScript.SetFocus
        Case vbKeyBack
            m_flx.Visible = False
            m_picToolTip.Visible = False
            m_txtScript.SetFocus
            SendKeys ("{BS}")
        Case 187 '=
            TypeIntelisenceText
            m_txtScript.SetFocus
            SendKeys ("=")
            
    End Select

End Sub

Private Sub m_flx_LostFocus()
    m_flx.Visible = False
    With m_picToolTip
        .Top = m_flx.Top + .TextHeight("A")
    End With
    'm_picToolTip.Visible = False
End Sub

'''Private Sub m_flx_RowColChange()
'''
'''    If m_flx.Col = 0 Then
'''        m_flx.Col = 1
'''    End If
'''    ShowToolTip
'''
'''End Sub

Private Sub m_txtScript_Change()

    
    If m_bPaintText Then
        m_bPaintText = False
        
        If m_bTextChange Then
            m_DebugTextbox.SetFocus
            PaintText
            m_txtScript.SetFocus
            m_bTextChange = False
        End If
        
    End If
    
End Sub

Private Sub m_txtScript_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn ', vbKeySpace
            m_bPaintText = True
            m_picToolTip.Visible = False
        Case 190 'point (.)
            m_bTextChange = True
            DoEvents
            ShowIntelisence
            'Show intelisence
        Case vbKeyTab
            KeyCode = 0
            m_txtScript.SelText = "    "
            DoEvents
            m_txtScript.SetFocus
        Case vbKeyBack
            m_bTextChange = True
            'Is it time to hide the tooltip ?
            If m_picToolTip.Visible Then
                If CursorOverideFunc Then
                    m_picToolTip.Visible = False
                End If
            End If
        Case 188 'comma (,)
            m_bTextChange = True
        Case Else
            m_bTextChange = True
    End Select

End Sub


Private Function CursorOverideFunc() As Boolean

    CursorOverideFunc = (m_txtScript.SelStart <= m_iFunctionEndPoint)
    
End Function


Private Sub m_txtScript_LostFocus()
    
    If Not (TypeOf frmScripts.ActiveControl Is ListView) Then
        m_picToolTip.Visible = False
    End If
    
End Sub

Private Sub m_txtScript_SelChange()

''''    With m_txtScript
''''        If .GetLineFromChar(.SelStart) < m_iMinPos Then
''''            m_iMinPos = .GetLineFromChar(.SelStart)
''''            Exit Sub
''''        End If
''''
''''        If .GetLineFromChar(.SelStart) > m_iMaxPos Then
''''            m_iMaxPos = .GetLineFromChar(.SelStart)
''''        End If
''''
''''    m_DebugTextbox.Text = GetCurrentPosition(m_txtScript)
''''    m_DebugTextbox.Text = m_DebugTextbox.Text & .SelRTF & vbNewLine & vbNewLine & m_txtScript.TextRTF
''''
''''
'''''''        m_DebugTextbox.Text = "minrow=" & m_iMinPos & vbNewLine & _
'''''''        "maxrow=" & m_iMaxPos
''''
''''    End With
    
    
    
End Sub

