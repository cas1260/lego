VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl CodeHighlight 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "DevDomainCodeHighlight.ctx":0000
   Begin RichTextLib.RichTextBox rich 
      Height          =   2028
      Left            =   216
      TabIndex        =   0
      Top             =   192
      Width           =   2964
      _ExtentX        =   5239
      _ExtentY        =   3572
      _Version        =   393217
      TextRTF         =   $"DevDomainCodeHighlight.ctx":0312
   End
End
Attribute VB_Name = "CodeHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Events
Public Event SelChange()
Public Event Change()
Public Event Click()
Public Event DblClick()

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
Public Event Validate(Cancel As Boolean)
Public TxtCod As RichTextBox
'Public TxtCod As RichTextBox

Public Enum ItemCodeType
    enumKeyword = 1
    enumOperator = 2
    enumFunction = 3
    enumDelimiter = 4
End Enum

Public Enum ProgrammingLanguage
    hlNOHighLight = 0
    hlVisualBasic = 1
    hlJava = 2
    hlhtml = 3
End Enum

Public Enum enumHighlightCode
    hlOnNewLine = 0
    hlAsType = 1
End Enum
' End Enums

' Public variables
Public CompareCase As VbCompareMethod
Public GiveCorrectCase As Boolean

Private bFireSelectionChange As Boolean
Private bListenForChange As Boolean
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private strSeparator(14) As String
Private iSeparatorCount As Integer

Private m_Language As ProgrammingLanguage
Dim HighLightWords() As HightlightedWord
Dim mHighlightCode As enumHighlightCode

Private Type HightlightedWord
    Word As String
    WordType As ItemCodeType
End Type

Private Type CommentTag
    CommentStart As String
    CommentEnd As String
End Type

Private m_Comment() As CommentTag
Private m_CommentCount As Integer

Dim WordCount As Integer

Dim mKeywordColor As OLE_COLOR
Dim mOperatorColor As OLE_COLOR
Dim mDelimiterColor As OLE_COLOR
Dim mForeColor As OLE_COLOR
Dim mFunctionColor As OLE_COLOR

Dim strKeywordColor As String
Dim strOperatorColor As String
Dim strDelimiterColor As String
Dim strForeColor As String
Dim strFunctionColor As String

' API Stuff

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_FMTLINES = &HC8
Private Const EM_LINELENGTH = &HC1
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EC_USEFONTINFO = &HFFFF
Private Const EM_SETMARGINS = &HD3
Private Const EM_GETMARGINS = &HD4
Private Const EM_CANUNDO = &HC6
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETHANDLE = &HBD
Private Const EM_GETMODIFY = &HB8
Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_GETRECT = &HB2
Private Const EM_GETSEL = &HB0
Private Const EM_GETTHUMB = &HBE
Private Const EM_GETWORDBREAKPROC = &HD1
Private Const EM_LIMITTEXT = &HC5
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB

Private Const EM_LINESCROLL = &HB6
Private Const EM_REPLACESEL = &HC2
Private Const EM_SCROLL = &HB5
Private Const EM_SCROLLCARET = &HB7
Private Const EM_SETHANDLE = &HBC
Private Const EM_SETMODIFY = &HB9
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETRECT = &HB3
Private Const EM_SETRECTNP = &HB4
Private Const EM_SETSEL = &HB1
Private Const EM_SETTABSTOPS = &HCB
Private Const EM_SETWORDBREAKPROC = &HD0
Private Const EM_UNDO = &HC7

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Function ColorWord(ByVal sWord As String) As String
Dim iWord As Integer
    For iWord = 0 To WordCount - 1
        If StrComp(sWord, HighLightWords(iWord).Word, CompareCase) = 0 Then
            If GiveCorrectCase Then sWord = HighLightWords(iWord).Word
            ColorWord = "\cf" & HighLightWords(iWord).WordType & " " & sWord & "\cf0 "
            Exit Function
        End If
    Next
    ' The word wo not found
    ColorWord = "\cf0 " & sWord & "\cf0 "
End Function

Private Function GetRTFColor(Color As OLE_COLOR) As String
    Dim lrgb As Long
    lrgb = TranslateColor(Color)
    GetRTFColor = "\red" & (lrgb And &HFF&) & "\green" & (lrgb And &HFF00&) \ &H100 & "\blue" & (lrgb And &HFF0000) \ &H10000 & ";"
End Function

Private Function GetWord(sBlock As String, lngWordStart As Long, lngCharPos As Long, sSep As String) As String
    Dim sWord As String
On Error GoTo en

    sWord = Mid$(sBlock, lngWordStart, lngCharPos - lngWordStart)
        ' Color Word
        If sSep = vbCrLf Then
            sSep = "\par " & vbCrLf
        ElseIf sSep = vbTab Then
                sSep = "\tab "
        ElseIf sSep = "\" Then
                sSep = "\cf2 \\\cf0 "
        ElseIf sSep = "{" Then
                sSep = "\cf2 \{\cf0 "
        ElseIf sSep = "}" Then
                sSep = "\cf2 \}\cf0 "
        ElseIf sSep <> " " And Len(sSep) Then
            sSep = "\cf2 " & sSep & "\cf0 "
        End If
        If lngCharPos - lngWordStart > 0 Then
            GetWord = ColorWord(sWord) & sSep
        Else
            GetWord = sSep
        End If
en:
End Function

Private Function HighlightComment(sComment As String, sEndofComment As String) As String
    sComment = Replace(sComment, "\", "\\")
    sComment = Replace(sComment, "{", "\{")
    sComment = Replace(sComment, "}", "\}")
    sComment = Replace(sComment, vbCrLf, "\par ")
    If sEndofComment = vbCrLf Then
        sComment = sComment & "\par" & vbCrLf
    Else
        If sEndofComment = vbTab Then
            sComment = sComment & "\tab "
        Else
            sComment = sComment & sEndofComment
        End If
    End If
    HighlightComment = "\cf4 " & sComment & "\cf0 "
End Function




' Finds if a string is the start of a comment.
' Returns -1 if it is not or the position in the comment array.
Private Function StartOfComment(sBlock As String, lngCharPos As Long) As Integer
Dim sChar As String
Dim i As Byte
    For i = 0 To m_CommentCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Comment(i).CommentStart))
        If sChar = m_Comment(i).CommentStart Then
            StartOfComment = i
            Exit Function
        End If
    Next
    StartOfComment = -1
End Function

Private Function isSeparator(sBlock As String, lngCharPos As Long) As String
Dim sChar As String
Dim i As Byte
    For i = 0 To iSeparatorCount
        sChar = Mid$(sBlock, lngCharPos, Len(strSeparator(i)))
        If sChar = strSeparator(i) Then
            isSeparator = sChar
            Exit Function
        End If
    Next
End Function
' Finds if a string is the end of a comment.
' Returns -1 if it is not or the position in the comment array.
Private Function EndOfComment(sBlock As String, lngCharPos As Long) As Integer
    Dim sChar As String
    Dim i As Byte
    For i = 0 To m_CommentCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Comment(i).CommentEnd))
        If sChar = m_Comment(i).CommentEnd Then
            EndOfComment = i
            Exit Function
        End If
    Next
    EndOfComment = -1
End Function


Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function
Public Sub AddCommentTag(ByVal CommentTagStart As String, ByVal CommentTagEnd As String)
    ReDim Preserve m_Comment(m_CommentCount)
    With m_Comment(m_CommentCount)
        .CommentStart = CommentTagStart
        .CommentEnd = CommentTagEnd
    End With
    m_CommentCount = m_CommentCount + 1
End Sub

Public Property Let BackColor(newColor As OLE_COLOR)
    rich.BackColor = newColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = rich.BackColor
End Property

Public Property Get Font() As StdFont
    Set Font = rich.Font
End Property
Public Property Set Font(newFont As StdFont)
    Set rich.Font = newFont
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Get FunctionColor() As OLE_COLOR
    FunctionColor = mFunctionColor
End Property


Public Property Let ForeColor(newForeColor As OLE_COLOR)
    mForeColor = newForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Let FunctionColor(newFunctionColor As OLE_COLOR)
    mFunctionColor = newFunctionColor
    strFunctionColor = GetRTFColor(mFunctionColor)
    PropertyChanged "FunctionColor"
End Property

Function HighlightBlock(sBlock As String) As String
    Dim lngCharPos As Long
    Dim lngBlockLength As Long
    Dim sWord As String
    Dim lngCommentStartPos As Long
    Dim byteStartOfComment As Integer
    Dim byteEndOfComment As Integer
    Dim sSep As String
    Dim lngWordStart As Long
    Dim sHighlighted As String
    Dim T As Integer
    Dim bWordFound As Boolean
    Dim bLastStepWasComment As Boolean

    If m_Language = hlNOHighLight Then
        HighlightBlock = sBlock
        Exit Function
    End If
    lngBlockLength = Len(sBlock)
    lngWordStart = 1
    byteStartOfComment = -1
    For lngCharPos = 1 To lngBlockLength
        ' Is this character the start of a comment
        T = StartOfComment(sBlock, lngCharPos)
        If T > -1 And byteStartOfComment = -1 Then
            lngCommentStartPos = lngCharPos
            byteStartOfComment = T
            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, "")
        Else
            'Is this character the end of a comment block
           If byteStartOfComment > -1 Then
                byteEndOfComment = EndOfComment(sBlock, lngCharPos)
                If byteEndOfComment > -1 And byteEndOfComment = byteStartOfComment Then
                    
                    sHighlighted = sHighlighted & HighlightComment(Mid$(sBlock, lngCommentStartPos, (lngCharPos - lngCommentStartPos)), m_Comment(byteEndOfComment).CommentEnd)

                    byteStartOfComment = -1
                    bLastStepWasComment = True
                    lngWordStart = lngCharPos + Len(m_Comment(byteEndOfComment).CommentEnd)
                End If
            Else
                If byteStartOfComment = -1 Then
                    ' Is this character a seperator
                    sSep = isSeparator(sBlock, lngCharPos)
                    Dim SepLength As ItemCodeType
                    SepLength = Len(sSep)
                    If SepLength > 0 Then
                        If lngCharPos <= lngBlockLength Then
                            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, sSep)
                        End If
                        lngWordStart = lngCharPos + SepLength
                            bLastStepWasComment = False
                    End If
                End If
            End If
        End If
    Next
    If byteStartOfComment > -1 Then
        ' Comment Hasn't been closed
        ' Search forward to find the end of the comment
        Dim lngCommentEndPos As Long
        lngCommentEndPos = InStr(lngCharPos, rich.Text, m_Comment(byteStartOfComment).CommentEnd)
        If lngCommentEndPos = 0 Then lngCommentEndPos = Len(rich.Text)
        sHighlighted = sHighlighted & HighlightComment(Mid$(sBlock, lngCommentStartPos, (lngCharPos - lngCommentStartPos)), "")
    Else
        If bLastStepWasComment Then
            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, "")
        Else
            If lngBlockLength - lngWordStart >= 0 Then
                sWord = Mid$(sBlock, lngWordStart, (lngBlockLength - lngWordStart) + 1)
                sHighlighted = sHighlighted & ColorWord(sWord)
            End If
        End If
    End If
    If Len(sHighlighted) = 0 Then Exit Function
    HighlightBlock = "{{\colortbl ;" & strKeywordColor & strOperatorColor & strFunctionColor & strDelimiterColor & "}" & sHighlighted & "}"
End Function
Public Property Get HighlightCode() As enumHighlightCode
    HighlightCode = mHighlightCode
End Property

Public Property Let HighlightCode(newHighlightCode As enumHighlightCode)
    mHighlightCode = newHighlightCode
    PropertyChanged "HighlightCode"
End Property



Public Property Get KeywordColor() As OLE_COLOR
    KeywordColor = mKeywordColor
End Property

Public Property Get DelimiterColor() As OLE_COLOR
    DelimiterColor = mDelimiterColor
End Property


Public Property Let DelimiterColor(newDelimiterColor As OLE_COLOR)
    mDelimiterColor = newDelimiterColor
    strDelimiterColor = GetRTFColor(mDelimiterColor)
    PropertyChanged "DelimiterColor"
End Property
' Returns the text at a given line
Public Property Get Line(lngLine As Long) As String
    ' When sending the EM_GETLINE message
    ' the lParam is a pointer to a buffer that will hold
    ' the returned line of text.
    ' The first word(16 bits) of this buffer contain the
    ' number of characters to return.  So we can set the
    ' number of characters to return a Byte array is used
    Dim bReturnedLineBuffer() As Byte
    Dim LengthOfLine As Long ' length of the line
    Dim LineStart As Long

    LineStart = LineStartPos(LineIndex)
    If LineStart = -1 Then Exit Function
    
    LengthOfLine = LineLength(LineStart)
    If LengthOfLine < 1 Then Exit Function
    
    'Resize the byte array
    ReDim bReturnedLineBuffer(LengthOfLine)

    'Save the length in the first word of the array
    'A Word is two bytes so split the length up
    bReturnedLineBuffer(0) = LengthOfLine And 255
    bReturnedLineBuffer(1) = LengthOfLine \ 256


    SendMessage rich.HWnd, EM_GETLINE, LineIndex, bReturnedLineBuffer(0)

    'Make the byte array a string and return it
    Line = Left$(StrConv(bReturnedLineBuffer, vbUnicode), LengthOfLine)
End Property

Public Property Let LineIndex(lngNewLineIndex As Long)
    rich.SelStart = Abs(LineStartPos(lngNewLineIndex))
End Property

' Returns the length of the line at the given character index
Public Property Get LineLength(CharacterIndex As Long) As Long
    LineLength = SendMessage(rich.HWnd, EM_LINELENGTH, CharacterIndex, 0&)
End Property

Public Property Get LineStartPos(ByVal LineIndex As Long) As Long
    LineStartPos = SendMessage(rich.HWnd, EM_LINEINDEX, LineIndex, 0&)
End Property


' Returns the line number of the current line
Public Property Get LineIndex() As Long
    LineIndex = SendMessage(rich.HWnd, EM_LINEFROMCHAR, ByVal -1, 0&)
End Property


Public Sub LoadFile(strFilename)
    Dim FileNum As Integer
    Dim sData As String
    Dim bListen As Boolean
    bListen = bListenForChange
    bListenForChange = False
    
    FileNum = FreeFile
    Open strFilename For Input As FileNum
        sData = Input(LOF(FileNum), FileNum)
    Close FileNum
    bFireSelectionChange = False
    rich.TextRTF = ""
    rich.SelRTF = HighlightBlock(sData)
    bFireSelectionChange = True
bListenForChange = bListen
End Sub

Public Property Get OperatorColor() As OLE_COLOR
    OperatorColor = mOperatorColor
End Property


Public Property Let KeywordColor(newKeywordColor As OLE_COLOR)
    mKeywordColor = newKeywordColor
    strKeywordColor = GetRTFColor(mKeywordColor)
    PropertyChanged "KeywordColor"
End Property


Public Property Let OperatorColor(newOperatorColor As OLE_COLOR)
    mOperatorColor = newOperatorColor
    strOperatorColor = GetRTFColor(mOperatorColor)
    PropertyChanged "OperatorColor"
End Property



Public Sub SaveFile(strFilename As String)
    rich.SaveFile strFilename, rtfText
End Sub

Public Property Let SelLength(lngNewSelLength As Long)
    rich.SelLength = lngNewSelLength
End Property
Public Property Get SelLength() As Long
    SelLength = rich.SelLength
End Property
Public Property Get SelStart() As Long
    SelStart = rich.SelStart
End Property

Public Property Let SelStart(lngNewSelStart As Long)
    rich.SelStart = lngNewSelStart
End Property

Private Sub SetJava()
    WordCount = 0
    AddWord "abstract"
    AddWord "boolean"
    AddWord "break"
    AddWord "byte"
    AddWord "case"
    AddWord "catch"
    AddWord "char"
    AddWord "class"
    AddWord "const"
    AddWord "continue"
    AddWord "default"
    AddWord "do"
    AddWord "double"
    AddWord "else"
    AddWord "extends"
    AddWord "final"
    AddWord "finally"
    AddWord "float"
    AddWord "for"
    AddWord "goto"
    AddWord "if"
    AddWord "implements"
    AddWord "import"
    AddWord "instanceof"
    AddWord "int"
    AddWord "interface"
    AddWord "long"
    AddWord "native"
    AddWord "new"
    AddWord "package"
    AddWord "private"
    AddWord "protected"
    AddWord "public"
    AddWord "return"
    AddWord "short"
    AddWord "static"
    AddWord "super"
    AddWord "switch"
    AddWord "synchronized"
    AddWord "this"
    AddWord "throw"
    AddWord "throws"
    AddWord "transient"
    AddWord "try"
    AddWord "void"
    AddWord "volatitle"
    AddWord "while"


    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "*", enumOperator
    AddWord "/", enumOperator
    AddWord "%", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord ">=", enumOperator
    AddWord "<=", enumOperator
    AddWord "!=", enumOperator
    AddWord "==", enumOperator
    AddWord "!", enumOperator
    AddWord "&&", enumOperator
    AddWord "||", enumOperator
    AddWord "-", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "^", enumOperator
    AddWord "<<", enumOperator
    AddWord ">>", enumOperator
    AddWord ">>>", enumOperator
    
    AddWord "=", enumOperator
    AddWord "++", enumOperator
    AddWord "--", enumOperator
    AddWord "+=", enumOperator
    AddWord "-=", enumOperator
    AddWord "*=", enumOperator
    AddWord "/=", enumOperator
    AddWord "%=", enumOperator
    AddWord "|=", enumOperator
    AddWord "&=", enumOperator
    AddWord "^=", enumOperator
    AddWord "<<=", enumOperator
    AddWord ">>=", enumOperator
    AddWord ">>>=", enumOperator
    AddWord "new", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    
    AddWord "true", enumOperator
    AddWord "false", enumOperator

    CompareCase = vbBinaryCompare
    ReDim Preserve m_Comment(0)
    m_CommentCount = 0
    AddWord """", enumDelimiter
    AddCommentTag "//", vbCrLf
    AddCommentTag "/*", "*/"
    AddCommentTag "/**", "*/"
    GiveCorrectCase = False
End Sub
Private Sub SetVB()
    WordCount = 0
    Erase m_Comment
    m_CommentCount = 0
    
    'AddWord "#Const"
    'AddWord "#Else"
    'AddWord "#ElseIf"
    'AddWord "#End If"
    'AddWord "#If"
    'AddWord "Alias"
    'AddWord "And"
    'AddWord "As"
    'AddWord "Base"
    'AddWord "Binary"
    'AddWord "Boolean"
    'AddWord "Byte"
    'AddWord "ByVal"
    'AddWord "Call"
    'AddWord "Case"
    'AddWord "CBool"
    'AddWord "CByte"
    'AddWord "CCur"
    'AddWord "CDate"
    'AddWord "CDbl"
    'AddWord "CDec"
    'AddWord "CInt"
    'AddWord "CLng"
    'AddWord "Close"
    'AddWord "Compare"
    'AddWord "Const"
    'AddWord "CSng"
    'AddWord "CStr"
    'AddWord "Currency"
    'AddWord "CVar"
    'AddWord "CVErr"
    'AddWord "Decimal"
    'AddWord "Declare"
    'AddWord "DefBool"
    'AddWord "DefByte"
    'AddWord "DefCur"
    'AddWord "DefDate"
    'AddWord "DefDbl"
    'AddWord "DefDec"
    'AddWord "DefInt"
    'AddWord "DefLng"
    'AddWord "DefObj"
    'AddWord "DefSng"
    'AddWord "DefStr"
    'AddWord "DefVar"
    'AddWord "Dim"
    'AddWord "Do"
    'AddWord "Double"
    'AddWord "Each"
    'AddWord "Else"
    'AddWord "ElseIf"
    'AddWord "End"
    'AddWord "Enum"
    'AddWord "Eqv"
    'AddWord "Erase"
    'AddWord "Error"
    'AddWord "Exit"
    'AddWord "Explicit"
    'AddWord "False"
    'AddWord "For"
    'AddWord "Function"
    'AddWord "Get"
    'AddWord "Global"
    'AddWord "GoSub"
    'AddWord "GoTo"
    'AddWord "If"
    'AddWord "Imp"
    'AddWord "In"

    'AddWord "Integer"
    'AddWord "Is"
    'AddWord "LBound"
    'AddWord "Let"
    'AddWord "Lib"
    'AddWord "Like"
    'AddWord "Line"
    'AddWord "Lock"
    'AddWord "Long"
    'AddWord "Loop"
    'AddWord "LSet"
    'AddWord "Name"
    'AddWord "New"
    'AddWord "Next"
    'AddWord "Not"
    'AddWord "Object"
    'AddWord "On"
    'AddWord "Open"
    'AddWord "Option"
    'AddWord "Optional"
    'AddWord "Or"
    'AddWord "Output"
    'AddWord "Print"
    'AddWord "Private"
    'AddWord "Property"
    'AddWord "Public"
    'AddWord "Put"
    'AddWord "Random"
    'AddWord "Read"
    'AddWord "ReDim"
    'AddWord "Resume"
    'AddWord "Return"
    'AddWord "RSet"
    'AddWord "Seek"
    'AddWord "Select"
    'AddWord "Set"
    'AddWord "Single"
    'AddWord "Spc"
    'AddWord "Static"
    'AddWord "String"
    'AddWord "Stop"
    'AddWord "Sub"
    'AddWord "Tab"
    'AddWord "Then"
    'AddWord "True"
    'AddWord "Type"
    'AddWord "UBound"
    'AddWord "Unlock"
    'AddWord "Variant"
    'AddWord "Wend"
    'AddWord "While"
    'AddWord "With"
    'AddWord "Nothing"
    'AddWord "To"
    
    'AddWord "Input"
    Dim Nome() As String
    ReDim Nome(30) As String
    
    X = 0
    AddWord "Legenda", enumKeyword
    AddWord "Cor de Fundo", enumKeyword
    AddWord "Cor da Letra", enumKeyword
    AddWord "TamX", enumKeyword
    AddWord "TAMY", enumKeyword
    AddWord "PoxX", enumKeyword
    AddWord "PoxY", enumKeyword
    AddWord "Fonte", enumKeyword
    AddWord "Tamanho", enumKeyword
    AddWord "Texto", enumKeyword
    AddWord "Imagem", enumKeyword
    AddWord "Borda", enumKeyword
    AddWord "Comprimir", enumKeyword
    AddWord "Ordem", enumKeyword
    AddWord "3d", enumKeyword
    AddWord "Marca", enumKeyword
    AddWord "Mascara", enumKeyword
    AddWord "Total", enumKeyword
    AddWord "Atual", enumKeyword
    AddWord "Busca", enumKeyword
    
    AddWord "Va Para", enumKeyword
    AddWord "Se", enumKeyword
    AddWord "SeNao", enumKeyword
    AddWord "Msg", enumKeyword
    AddWord "Loop ", enumKeyword
    AddWord "Fim do Programa", enumKeyword
    AddWord "FimSe", enumKeyword
    AddWord "Chama", enumKeyword
    AddWord "Fecha Tela", enumKeyword
    AddWord "Focus", enumKeyword
    AddWord "Adicione", enumKeyword
    AddWord "Selecione", enumKeyword
    AddWord "Limpa", enumKeyword
    AddWord "SimNao", enumFunction
    AddWord "Exe ", enumKeyword
    AddWord "Direita", enumFunction
    AddWord "Esquerda", enumFunction
    AddWord "Procedimento", enumFunction
    AddWord "Db", enumFunction
    AddWord "Tela Cheia()", enumFunction
    AddWord "Caixa", enumFunction
    AddWord "Faça Enquanto ", enumFunction
    AddWord "Visivel", enumFunction
        
    'AddWord "Xor", enumOperator
    AddWord "=", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "<=", enumOperator
    AddWord ">=", enumOperator
    AddWord "=<", enumOperator
    AddWord "=>", enumOperator
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "/", enumOperator
    AddWord "*", enumOperator
    AddWord "<>", enumOperator
    AddWord "&", enumOperator

    AddWord """", enumDelimiter
    CompareCase = vbTextCompare
    AddCommentTag "'", vbCrLf
    AddCommentTag "//", vbCrLf
    AddCommentTag "*", vbCrLf
    AddCommentTag "\\", vbCrLf
    GiveCorrectCase = True
End Sub





Private Sub rich_Change()
    RaiseEvent Change
End Sub

Private Sub rich_Click()
    RaiseEvent Click
End Sub


Private Sub rich_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rich_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = vbKeyTab Then ' Indent
        Dim SelStart As Long
        If rich.SelLength > 0 Then
            Dim strLines() As String
            Dim LineCount As Long, i As Long
            Dim strResult As String
            strLines = Split(rich.SelText, vbCrLf)
            LineCount = UBound(strLines)
            If LineCount > 0 Then
                SelStart = rich.SelStart
                For i = 0 To LineCount - 1
                    strResult = strResult & vbTab & strLines(i) & vbCrLf
                Next
                strResult = strResult & vbTab & strLines(i)
                rich.SelText = strResult
                rich.SelStart = SelStart
                rich.SelLength = Len(strResult)
                KeyCode = 0
            End If
        End If
    End If

End Sub

Private Sub rich_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Dim i As Byte
    If mHighlightCode = hlAsType Then
        For i = 0 To iSeparatorCount
            If KeyAscii = Asc(strSeparator(i)) Then
                    LockWindowUpdate rich.HWnd
                    bFireSelectionChange = False
                    Dim TheStart As Long
                    TheStart = rich.SelStart
                    rich.SelStart = Me.LineStartPos(Me.LineIndex)
                    rich.SelLength = Me.LineLength(rich.SelStart)
                    rich.SelRTF = HighlightBlock(Line(Me.LineIndex))
                    rich.SelStart = TheStart
                    LockWindowUpdate 0
                    bFireSelectionChange = True
                Exit Sub
            End If
        Next
    End If
End Sub


Private Sub rich_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rich_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub rich_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub rich_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub rich_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub


Private Sub rich_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub rich_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub rich_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub rich_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub rich_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub rich_SelChange()
    On Error Resume Next
    Static lngLastLine As Long
    Dim lngNewLine As Long
    Dim TheStart As Long
    
    If bFireSelectionChange Then
        If rich.SelLength = 0 Then
                bFireSelectionChange = False
                lngNewLine = Me.LineIndex
                If lngNewLine <> lngLastLine Then
                    On Error GoTo en:
                    LockWindowUpdate rich.HWnd
                    TheStart = rich.SelStart
                    rich.SelStart = IIf(Me.LineStartPos(lngLastLine) < 0, 0, Me.LineStartPos(lngLastLine))
                    rich.SelLength = Me.LineLength(rich.SelStart)
                    rich.SelRTF = HighlightBlock(Line(lngLastLine))
en:
                    rich.SelStart = TheStart
                    rich.SelLength = SelLength
                    LockWindowUpdate 0
                End If
                lngLastLine = lngNewLine
                bFireSelectionChange = True
        End If
    RaiseEvent SelChange
    End If
End Sub

Private Sub rich_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

Private Sub UserControl_Initialize()
    Set TxtCod = rich
    strSeparator(0) = " "
    strSeparator(1) = vbCrLf
    strSeparator(2) = vbTab
    strSeparator(3) = "("
    strSeparator(4) = ")"
    strSeparator(5) = "="
    strSeparator(6) = "+"
    strSeparator(7) = "-"
    strSeparator(8) = "*"
    strSeparator(9) = ">"
    strSeparator(10) = "<"
    strSeparator(11) = "\"
    strSeparator(12) = "/"
    strSeparator(13) = "{"
    strSeparator(14) = "}"
    iSeparatorCount = 14
    bFireSelectionChange = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    rich.Text = PropBag.ReadProperty("Text", "")
    Language = PropBag.ReadProperty("Language", hlNOHighLight)

    rich.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    KeywordColor = PropBag.ReadProperty("KeywordColor", vbBlue)
    OperatorColor = PropBag.ReadProperty("OperatorColor", vbYellow)
    DelimiterColor = PropBag.ReadProperty("DelimiterColor", vbCyan)
    mForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    FunctionColor = PropBag.ReadProperty("FunctionColor", vbMagenta)
    HighlightCode = PropBag.ReadProperty("HighlightCode", 1)
    
    Set rich.Font = PropBag.ReadProperty("Font", rich.Font)
End Sub

Private Sub UserControl_Resize()
    rich.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Public Sub AddWord(ByVal Word As String, Optional WordType As ItemCodeType = enumKeyword)
    ReDim Preserve HighLightWords(WordCount)
    If WordType = enumDelimiter Then
        AddCommentTag Word, Word
    Else
        With HighLightWords(WordCount)
            .Word = Word
            .WordType = WordType
        End With
        WordCount = WordCount + 1
    End If
End Sub


Public Property Get Text() As String
    Text = rich.Text
End Property

Public Property Get SelText() As String
    SelText = rich.SelText
End Property


Public Property Let SelText(newSelText As String)
bFireSelectionChange = False
    rich.SelRTF = HighlightBlock(newSelText)
    bFireSelectionChange = True
End Property

Public Property Let Text(ByVal vNewValue As String)
    rich.TextRTF = HighlightBlock(vNewValue)
    PropertyChanged "Text"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", rich.Text, ""
    PropBag.WriteProperty "Language", m_Language, hlNOHighLight
    
    PropBag.WriteProperty "BackColor", rich.BackColor, vbWindowBackground
    PropBag.WriteProperty "KeywordColor", mKeywordColor, vbBlue
    PropBag.WriteProperty "OperatorColor", mOperatorColor, vbYellow
    PropBag.WriteProperty "DelimiterColor", mDelimiterColor, vbCyan
    PropBag.WriteProperty "ForeColor", mForeColor, vbWindowText
    PropBag.WriteProperty "FunctionColor", mFunctionColor, vbMagenta
    PropBag.WriteProperty "HighlightCode", mHighlightCode, 1
    
    PropBag.WriteProperty "Font", rich.Font
End Sub




Public Property Get Language() As ProgrammingLanguage
    Language = m_Language
End Property

Public Property Let Language(ByVal vNewValue As ProgrammingLanguage)
Dim sData As String
    If m_Language <> vNewValue Then
        Select Case vNewValue
            Case hlVisualBasic
                SetVB
            Case hlJava
                SetJava
            Case hlhtml
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
                AddCommentTag "<", ">"
            Case hlNOHighLight
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
        End Select
        m_Language = vNewValue

        sData = rich.Text
        rich.TextRTF = ""
        rich.SelRTF = HighlightBlock(sData)
        PropertyChanged "Language"
    End If
End Property
