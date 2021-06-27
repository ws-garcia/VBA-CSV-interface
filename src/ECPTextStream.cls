VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ECPTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2021 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' ECPTextStream is an easy-to-use class module developed to enable I/O operations over "big"
' text files, at high speed, from VBA. The module hasn’t reference to any external API library.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const DualLFchar As String = vbLf & vbLf
Private Const InverseCRLF As String = vbLf & vbCr
Private Const SizeFactor As Long = 524288
Private Const WhiteSpace As String = " "
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private P_ATENDOFSTREAM As Boolean '---------------Indicates if the file's end is reached.
Private P_BUFFERLENGTH As Long '-------------------Number of chars in buffer.
Private P_BUFFERSIZE As Single '-------------------Buffer's size x10 MB (0.5 by default).
Private P_ENDSTREAMONLINEBREAK As Boolean '--------If true, each stream ends on a line break.
Private P_ISOPENSTREAM As Boolean '----------------Indicates if the object is linked to a file
Private P_LINEBREAK As String '--------------------Holds the char used to end a Stream.
Private P_LINEBREAKMATCHINGBEHAVIOR As EndLineMatchingBehavior 'How to find the next line break.
Private P_UNIFIEDLFOUTPUT As Boolean '-------------If true, the buffer string will be returned
'                                                  with the LF char as Line Break.
Private P_STREAMLENGTH As Long '-------------------File len.
Private P_TEXT As String '-------------------------Holds the current stream's text.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
' @Common
Private Last2Chrs As String
Private LastChr As String
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Private Buffer As String
Private BufferDelta As Long
Private BufferEnds As Boolean
Private BufferMark As Long
Private CorrectedPos As Long
Private EndLineMark As EndLineChar
Private FileHandled As Integer
Private InitialPos As Long
Private LCS As Long
Private NullChar As String
Private NullCharPos As Long
Private TmpInitialPos As Long
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum EndLineChar
    CRLF = 0
    CR = 1
    LF = 2
End Enum
Public Enum EndLineMatchingBehavior
    OnlyBackwardSense = 0 'From the end to the beginning of the buffer.
    Bidirectional = 1 'Find the line break in both directions.
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#
' PROPERTIES:
Public Property Get atEndOfStream() As Boolean
Attribute atEndOfStream.VB_Description = "Gets the pointer’s overall status over the text stream."
    atEndOfStream = P_ATENDOFSTREAM
End Property
Public Property Get bufferLength() As Long
Attribute bufferLength.VB_Description = "Gets the number of string characters in the buffer."
    bufferLength = P_BUFFERLENGTH
End Property
Public Property Get bufferSize() As Single
Attribute bufferSize.VB_Description = "Gets or sets the buffer’s size, in MB, for text stream operations."
    bufferSize = P_BUFFERSIZE
End Property
Public Property Let bufferSize(value As Single)
    P_BUFFERSIZE = value
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * SizeFactor)
    Buffer = SPACE$(P_BUFFERLENGTH)
End Property
Public Property Get bufferString() As String
Attribute bufferString.VB_Description = "Gets the text data stored in the buffer."
    If P_UNIFIEDLFOUTPUT Then
        NormalizeLineBreaks
    End If
    bufferString = P_TEXT
End Property
Public Property Get endStreamOnLineBreak() As Boolean
Attribute endStreamOnLineBreak.VB_Description = "Allows to end buffer just after the first, from right to left, line break character."
    endStreamOnLineBreak = P_ENDSTREAMONLINEBREAK
End Property
Public Property Let endStreamOnLineBreak(value As Boolean)
    P_ENDSTREAMONLINEBREAK = value
End Property
Public Property Get isOpenStream() As Boolean
    isOpenStream = P_ISOPENSTREAM
End Property
Public Property Get lineBreak() As String
Attribute lineBreak.VB_Description = "Returns the character used to end the last received stream. The value is vbNullString when the last stream is not forced to end on line break."
    lineBreak = P_LINEBREAK
End Property
Public Property Get linebreakMatchingBehavior() As EndLineMatchingBehavior
    linebreakMatchingBehavior = P_LINEBREAKMATCHINGBEHAVIOR
End Property
Public Property Let linebreakMatchingBehavior(value As EndLineMatchingBehavior)
    P_LINEBREAKMATCHINGBEHAVIOR = value
End Property
Public Property Get pointerPosition() As Long
Attribute pointerPosition.VB_Description = "Gets the overall pointer position over the current text file."
    If P_ISOPENSTREAM Then
        pointerPosition = Seek(FileHandled)
    Else
        pointerPosition = 0
    End If
End Property
Public Property Get streamLength() As Long
Attribute streamLength.VB_Description = "Gets the current opened file’s size, in Bytes."
    streamLength = P_STREAMLENGTH
End Property
Public Property Get unifiedLFOutput() As Boolean
Attribute unifiedLFOutput.VB_Description = "Determines whether the buffer string is returned using only the LF character as a linefeed."
    unifiedLFOutput = P_UNIFIEDLFOUTPUT
End Property
Public Property Let unifiedLFOutput(value As Boolean)
    P_UNIFIEDLFOUTPUT = value
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:
Public Sub CloseStream()
Attribute CloseStream.VB_Description = "Closes the current text file stream."
    If P_ISOPENSTREAM Then
        Close #FileHandled
        P_ISOPENSTREAM = False
    End If
End Sub
Private Sub DoubleBufferSize()
    Dim LCSt As Long
    P_BUFFERSIZE = 2 * P_BUFFERSIZE
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * SizeFactor)
    LCSt = P_STREAMLENGTH - P_BUFFERLENGTH
    Select Case LCSt
        Case Is > 0
            Buffer = SPACE$(P_BUFFERLENGTH)
        Case Else
            Buffer = SPACE$(P_STREAMLENGTH)
    End Select
End Sub
Private Sub FindEOLcharacter()
    Dim CrCharInStream As Boolean
    Dim LfCharInStream  As Boolean
    Dim EOLchr As EndLineChar
    Dim missingEOLchar As Boolean
    Dim EOStream As Boolean
    
    If P_LINEBREAKMATCHINGBEHAVIOR = EndLineMatchingBehavior.Bidirectional Then
        Do
            CrCharInStream = InStrB(1, Buffer, vbCr)
            LfCharInStream = InStrB(1, Buffer, vbLf)
            missingEOLchar = (Not CrCharInStream) And (Not LfCharInStream)
            If missingEOLchar Then
                DoubleBufferSize
                SeekPointer TmpInitialPos
                Get #FileHandled, , Buffer
                InitialPos = Seek(FileHandled)
                BufferMark = LenB(Buffer)
                EOStream = (P_STREAMLENGTH <= InitialPos)
            End If
        Loop While missingEOLchar And Not EOStream
        P_ATENDOFSTREAM = EOStream
    End If
    If Not EOStream Then
        If Not missingEOLchar Then
            Last2Chrs = MidB$(Buffer, BufferMark - 3, 4)
            BufferEnds = (Last2Chrs = vbCrLf)
            Select Case BufferEnds
                Case False
                    LastChr = MidB$(Last2Chrs, 3, 2)
                    BufferEnds = (LastChr = vbCr)
                    Select Case BufferEnds
                        Case False
                            BufferEnds = (LastChr = vbLf)
                            If BufferEnds Then
                                P_LINEBREAK = vbLf
                            Else
                                GoBackToLineBreak
                            End If
                        Case Else
                            P_LINEBREAK = vbCr
                    End Select
                Case Else
                    P_LINEBREAK = vbCrLf
            End Select
        End If
        CorrectedPos = InitialPos - BufferDelta
        BufferDelta = 0
    Else
        NullCharPos = InStrB(Buffer, NullChar)
        If NullCharPos Then
            BufferMark = NullCharPos
        End If
        CorrectedPos = P_STREAMLENGTH + 1
    End If
    Seek #FileHandled, CorrectedPos
End Sub
Private Sub GoBackToLineBreak()
    Do
        BufferMark = BufferMark - 2
        BufferDelta = BufferDelta + 1
        Last2Chrs = MidB$(Buffer, BufferMark - 3, 4)
        BufferEnds = (Last2Chrs = vbCrLf)
        Select Case BufferEnds
            Case False
                LastChr = MidB$(Last2Chrs, 3, 2)
                BufferEnds = (LastChr = vbCr)
            Select Case BufferEnds
                Case False
                    BufferEnds = (LastChr = vbLf)
                    If BufferEnds Then
                        P_LINEBREAK = vbLf
                    End If
                Case Else
                    P_LINEBREAK = vbCr
            End Select
            Case Else
                P_LINEBREAK = vbCrLf
        End Select
    Loop While Not BufferEnds
End Sub
Private Sub NormalizeLineBreaks()
    If InStrB(1, P_TEXT, vbCr, vbBinaryCompare) Then
        P_TEXT = Replace(P_TEXT, vbCr, vbLf, 1)
    End If
    Do While InStrB(1, P_TEXT, DualLFchar, vbBinaryCompare)
        P_TEXT = Replace(P_TEXT, DualLFchar, vbLf, 1)
    Loop
    P_LINEBREAK = vbLf
End Sub
Public Sub OpenStream(FilePath As String)
Attribute OpenStream.VB_Description = "Opens a stream over a text file."
    FileHandled = FreeFile
    Open FilePath For Binary As #FileHandled
    P_ISOPENSTREAM = True
    P_STREAMLENGTH = LOF(FileHandled)
    StartVariables
End Sub
Public Sub ReadText()
Attribute ReadText.VB_Description = "Reads a number of characters from the stream file and saves the result to the current instance."
    If Not P_ATENDOFSTREAM And P_ISOPENSTREAM Then
        If InitialPos = 0 Then InitialPos = 1
        Select Case P_BUFFERLENGTH
            Case Is < LCS
                BufferDelta = 0
                TmpInitialPos = InitialPos
                Get #FileHandled, , Buffer
                InitialPos = Seek(FileHandled)
                BufferMark = LenB(Buffer)
                If P_ENDSTREAMONLINEBREAK Then
                    FindEOLcharacter
                    LCS = P_STREAMLENGTH - CorrectedPos
                    P_TEXT = MidB$(Buffer, 1, BufferMark)
                Else
                    P_LINEBREAK = vbNullString
                    LCS = P_STREAMLENGTH - InitialPos
                    P_ATENDOFSTREAM = False
                    P_TEXT = Buffer
                End If
            Case Else
                Buffer = SPACE$(LCS + 1)
                Get #FileHandled, , Buffer
                NullCharPos = InStrB(Buffer, NullChar)
                If NullCharPos Then
                    Buffer = MidB$(Buffer, 1, NullCharPos)
                End If
                P_ATENDOFSTREAM = True
                P_TEXT = Buffer
        End Select
    End If
End Sub
Public Sub RestartPointer()
Attribute RestartPointer.VB_Description = "Moves the pointer to the text stream’s beginning position."
    If P_ISOPENSTREAM Then
        Seek #FileHandled, 1
        StartVariables
    End If
End Sub
Public Sub SeekPointer(Position As Long)
Attribute SeekPointer.VB_Description = "Moves the pointer, over the target file, to the specified position."
    If Position Then
        If P_ISOPENSTREAM Then
            If P_STREAMLENGTH Then
                Seek #FileHandled, Position
                CorrectedPos = Seek(FileHandled)
                LCS = P_STREAMLENGTH - CorrectedPos
                P_ATENDOFSTREAM = (P_STREAMLENGTH < CorrectedPos)
            End If
        End If
    End If
End Sub
Private Sub StartVariables()
    CorrectedPos = 0
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * SizeFactor)
    Buffer = SPACE$(P_BUFFERLENGTH)
    LCS = P_STREAMLENGTH
    P_ATENDOFSTREAM = False
End Sub
Public Sub WriteBlankLines(Lines As Long, Optional EndLineMark As EndLineChar = 0)
Attribute WriteBlankLines.VB_Description = "Inserts a specified number of blank lines into the current opened text file."
    If P_ISOPENSTREAM Then
        Dim idx As Long
        Select Case EndLineMark
            Case 0
                For idx = 1 To Lines
                    Put #FileHandled, , vbCrLf
                Next idx
            Case 1
                For idx = 1 To Lines
                    Put #FileHandled, , vbCr
                Next idx
            Case 2
                For idx = 1 To Lines
                    Put #FileHandled, , vbLf
                Next idx
        End Select
        P_STREAMLENGTH = LOF(FileHandled)
    End If
End Sub
Public Sub WriteText(ByRef TextData As String)
Attribute WriteText.VB_Description = "Writes the given string to the current opened text file."
    If P_ISOPENSTREAM Then
        Put #FileHandled, , TextData
        P_STREAMLENGTH = LOF(FileHandled)
    End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    P_BUFFERSIZE = 0.5
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * SizeFactor)
    P_ENDSTREAMONLINEBREAK = False
    P_LINEBREAKMATCHINGBEHAVIOR = EndLineMatchingBehavior.Bidirectional
    P_UNIFIEDLFOUTPUT = False
    Buffer = SPACE$(P_BUFFERLENGTH)
    NullChar = ChrW(0)
End Sub
Private Sub Class_Terminate()
    If P_ISOPENSTREAM Then
        CloseStream
    End If
End Sub
