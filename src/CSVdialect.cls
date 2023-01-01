VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVdialect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022-2023 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' Class module developed to share CSV dialects, or group of specific and related
' configuration, which instructs the parser on how to interpret the character set read from a
' CSV file. This container travels through the parsing and sniffer methods.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const CHR_COMMA As String = ","
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private P_ESCAPEMODE As EscapeStyle
Private P_FIELDSDELIMITER As String
Private P_RECORDSDELIMITER As String
Private P_QUOTETOKEN As QuoteTokens
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum QuoteTokens
    Apostrophe = 1
    DoubleQuotes = 2
    Tilde = 3
End Enum
Public Enum EscapeStyle
    rfc4180 = 0
    unix = 1
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#
''' <summary>
''' Gets or sets the escape behavior.
''' 0 --> escaping as per RFC-4180 specs
''' 1 -- > escaping ALSO in unix style
''' </summary>
Public Property Get escapeMode() As EscapeStyle
    escapeMode = P_ESCAPEMODE
End Property
Public Property Let escapeMode(value As EscapeStyle)
    P_ESCAPEMODE = value
End Property
''' <summary>
''' Gets or sets the quote token.
''' </summary>
Public Property Get quoteToken() As QuoteTokens
    quoteToken = P_QUOTETOKEN
End Property
Public Property Let quoteToken(ByVal token As QuoteTokens)
    P_QUOTETOKEN = token
End Property
''' <summary>
''' Gets or sets the fields delimiter.
''' </summary>
Public Property Get fieldsDelimiter() As String
    fieldsDelimiter = P_FIELDSDELIMITER
End Property
Public Property Let fieldsDelimiter(ByVal Delimiter As String)
    Select Case LenB(Delimiter)
        Case 2
            P_FIELDSDELIMITER = Delimiter
        Case Else
            P_FIELDSDELIMITER = CHR_COMMA
    End Select
End Property
''' <summary>
''' Gets or sets the records delimiter.
''' </summary>
Public Property Get recordsDelimiter() As String
    recordsDelimiter = P_RECORDSDELIMITER
End Property
Public Property Let recordsDelimiter(Delimiter As String)
    Select Case Delimiter
        Case vbCrLf, vbCr, vbLf
            P_RECORDSDELIMITER = Delimiter
        Case Else
            P_RECORDSDELIMITER = vbCrLf
    End Select
End Property
Private Sub Class_Initialize()
    P_ESCAPEMODE = rfc4180
    P_QUOTETOKEN = DoubleQuotes
    P_FIELDSDELIMITER = CHR_COMMA
    P_RECORDSDELIMITER = vbCrLf
End Sub
