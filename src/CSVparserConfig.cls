VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVparserConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
' GENERAL INFO:
' This module is part of VBA CSV interface library. Is developed to share configuration
' options through parsing and sniffer methods.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const CHR_COLON As String = ":"
Private Const CHR_COMMA As String = ","
Private Const CHR_SEMICOLON As String = ";"
Private Const CHR_PIPE As String = "|"
Private Const CHR_SPACE As String = " "
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private P_BUFFERSISE As Single
Private P_COMMENTSTOKEN As Long
Private P_DELIMITERSGUESSING As Boolean
Private P_DELIMITERSTOGUESS() As String
Private P_DIALECT As CSVdialect
Private P_DYNAMIC_TYPING As Boolean
Private P_ENDINGRECORD As Long
Private P_HEADERS As Boolean
Private P_HEADEROMISSION As Boolean
Private P_TEMPLATE_DEFINED As Boolean
Private P_TEMPLATE_TARGETS_DEFINED As Boolean
Private P_PATH As String
Private P_SKIP_COMMENTS As Boolean
Private P_SKIP_EMPTYLINES As Boolean
Private P_STARTINGRECORD As Long
Private P_TEMPLATE_TARGETS() As Variant
Private P_TEMPLATE() As Variant
Private P_MULTIENDOFLINECSV As Boolean
Private P_UTF8ENCODED As Boolean
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum TypeConversion
    None = -1
    ToLong = 0
    ToDouble = 1
    ToDate = 2
    ToBoolean = 3
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
Public Property Get bufferSize() As Single
    bufferSize = P_BUFFERSISE
End Property
Public Property Let bufferSize(Size As Single)
    P_BUFFERSISE = Size
End Property
Public Property Get commentsToken() As String
    commentsToken = ChrW(P_COMMENTSTOKEN)
End Property
Public Property Let commentsToken(Token As String)
    Dim ascWTkn As Long
    ascWTkn = AscW(Token)
    Select Case ascWTkn
        Case Is = 33, 35, 36, 37, 38
            P_COMMENTSTOKEN = ascWTkn
        Case Else
            P_COMMENTSTOKEN = 35 '# by default
    End Select
End Property
Public Property Get delimitersGuessing() As Boolean
    delimitersGuessing = P_DELIMITERSGUESSING
End Property
Public Property Let delimitersGuessing(value As Boolean)
    P_DELIMITERSGUESSING = value
End Property
Public Property Get delimitersToGuess() As String()
    delimitersToGuess = P_DELIMITERSTOGUESS
End Property
Public Property Let delimitersToGuess(value() As String)
    P_DELIMITERSTOGUESS = value
End Property
Public Property Get dialect() As CSVdialect
    Set dialect = P_DIALECT
End Property
Public Property Set dialect(dialectObj As CSVdialect)
    If Not dialectObj Is Nothing Then
        Set P_DIALECT = dialectObj
    End If
End Property
Public Property Get dTTemplateDefined() As Boolean
    dTTemplateDefined = P_TEMPLATE_DEFINED
End Property
Public Property Get dTTemplateLinksDefined() As Boolean
    dTTemplateLinksDefined = P_TEMPLATE_TARGETS_DEFINED
End Property
Public Property Get dTypingLinks() As Variant()
    dTypingLinks = P_TEMPLATE_TARGETS
End Property
Public Property Let dTypingLinks(indexes() As Variant)
    If LBound(indexes) = 0 Then
        P_TEMPLATE_TARGETS() = indexes
        Select Case UBound(indexes)
            Case -1
                P_TEMPLATE_TARGETS_DEFINED = False
            Case Else
                P_TEMPLATE_TARGETS_DEFINED = True
        End Select
    End If
End Property
Public Property Get dTypingTemplate() As Variant()
    dTypingTemplate = P_TEMPLATE
End Property
Public Property Let dTypingTemplate(template() As Variant)
    If LBound(template) = 0 Then
        P_TEMPLATE() = template
        Select Case UBound(template)
            Case -1
                P_TEMPLATE_DEFINED = False
            Case Else
                P_TEMPLATE_DEFINED = True
        End Select
    End If
End Property
Public Property Get dynamicTyping() As Boolean
    dynamicTyping = P_DYNAMIC_TYPING
End Property
Public Property Let dynamicTyping(value As Boolean)
    P_DYNAMIC_TYPING = value
End Property
Public Property Get endingRecord() As Long
    endingRecord = P_ENDINGRECORD
End Property
Public Property Let endingRecord(ByVal RecNumber As Long)
    Select Case RecNumber
        Case Is >= 0&
            P_ENDINGRECORD = RecNumber
        Case Else
            P_ENDINGRECORD = 0&
    End Select
End Property
Public Property Get Headers() As Boolean
    Headers = P_HEADERS
End Property
Public Property Let Headers(value As Boolean)
    P_HEADERS = value
End Property
Public Property Get headersOmission() As Boolean
    headersOmission = P_HEADEROMISSION
End Property
Public Property Let headersOmission(value As Boolean)
    P_HEADEROMISSION = value
End Property
Public Property Get skipCommentLines() As Boolean
    skipCommentLines = P_SKIP_COMMENTS
End Property
Public Property Let skipCommentLines(value As Boolean)
    P_SKIP_COMMENTS = value
End Property
Public Property Get skipEmptyLines() As Boolean
    skipEmptyLines = P_SKIP_EMPTYLINES
End Property
Public Property Let skipEmptyLines(value As Boolean)
    P_SKIP_EMPTYLINES = value
End Property
Public Property Get path() As String
    path = P_PATH
End Property
Public Property Let path(FullPath As String)
    P_PATH = FullPath
End Property
Public Property Get startingRecord() As Long
    startingRecord = P_STARTINGRECORD
End Property
Public Property Let startingRecord(ByVal RecNumber As Long)
    Select Case RecNumber
        Case Is >= 1&
            P_STARTINGRECORD = RecNumber
        Case Else
            P_STARTINGRECORD = 1&
    End Select
End Property
Public Property Get multiEndOfLineCSV() As Boolean
    multiEndOfLineCSV = P_MULTIENDOFLINECSV
End Property
Public Property Let multiEndOfLineCSV(ByVal value As Boolean)
    P_MULTIENDOFLINECSV = value
End Property
Public Property Get utf8EncodedFile() As Boolean
    utf8EncodedFile = P_UTF8ENCODED
End Property
Public Property Let utf8EncodedFile(value As Boolean)
    P_UTF8ENCODED = value
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Public Function CopyConfig() As CSVparserConfig
    Dim tmpResult As CSVparserConfig
    
    Set tmpResult = New CSVparserConfig
    tmpResult.bufferSize = P_BUFFERSISE
    tmpResult.skipCommentLines = P_SKIP_COMMENTS
    tmpResult.skipEmptyLines = P_SKIP_EMPTYLINES
    tmpResult.commentsToken = ChrW(P_COMMENTSTOKEN)
    tmpResult.delimitersGuessing = P_DELIMITERSGUESSING
    tmpResult.delimitersToGuess = P_DELIMITERSTOGUESS
    tmpResult.dialect.escapeMode = P_DIALECT.escapeMode
    tmpResult.dialect.quoteToken = P_DIALECT.quoteToken
    tmpResult.dialect.fieldsDelimiter = P_DIALECT.fieldsDelimiter
    tmpResult.dialect.recordsDelimiter = P_DIALECT.recordsDelimiter
    tmpResult.dTypingLinks = P_TEMPLATE_TARGETS
    tmpResult.dTypingTemplate = P_TEMPLATE
    tmpResult.dynamicTyping = P_DYNAMIC_TYPING
    tmpResult.endingRecord = P_ENDINGRECORD

    tmpResult.Headers = P_HEADERS
    tmpResult.headersOmission = P_HEADEROMISSION
    tmpResult.path = P_PATH
    tmpResult.startingRecord = P_STARTINGRECORD
    tmpResult.multiEndOfLineCSV = P_MULTIENDOFLINECSV
    Set CopyConfig = tmpResult
End Function
Public Sub DefineTypingTemplate(ParamArray dType() As Variant)
    P_TEMPLATE() = dType
    Select Case UBound(P_TEMPLATE)
        Case -1
            P_TEMPLATE_DEFINED = False
        Case Else
            P_TEMPLATE_DEFINED = True
    End Select
End Sub
Public Sub DefineTypingTemplateLinks(ParamArray dTargets() As Variant)
    P_TEMPLATE_TARGETS() = dTargets
    Select Case UBound(P_TEMPLATE_TARGETS)
        Case -1
            P_TEMPLATE_TARGETS_DEFINED = False
        Case Else
            P_TEMPLATE_TARGETS_DEFINED = True
    End Select
End Sub
Private Sub Class_Initialize()
    P_BUFFERSISE = 0.5
    P_COMMENTSTOKEN = AscW("#")
    P_DELIMITERSGUESSING = False
    ReDim P_DELIMITERSTOGUESS(0 To 4): P_DELIMITERSTOGUESS(0) = CHR_COMMA: _
                                        P_DELIMITERSTOGUESS(1) = CHR_SEMICOLON: _
                                        P_DELIMITERSTOGUESS(2) = vbTab: _
                                        P_DELIMITERSTOGUESS(3) = CHR_PIPE: _
                                        P_DELIMITERSTOGUESS(4) = CHR_COLON
    Set P_DIALECT = New CSVdialect
    P_DYNAMIC_TYPING = False
    P_ENDINGRECORD = 0&
    P_HEADERS = True
    P_HEADEROMISSION = False
    P_STARTINGRECORD = 1&
    P_SKIP_COMMENTS = True
    P_SKIP_EMPTYLINES = True
    P_MULTIENDOFLINECSV = False
    P_UTF8ENCODED = False
    DefineTypingTemplate
    DefineTypingTemplateLinks
End Sub


