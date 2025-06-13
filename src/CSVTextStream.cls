VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2021-2025 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' CSVTextStream is an easy-to-use class module developed to enable I/O operations over "big"
' text files, at high speed, from VBA. The module hasn’t reference to any external API
' library and has the ability to read and write UTF-8 encoded files.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const DualLFchar As String = vbLf & vbLf
Private Const sizeFactor As Long = 524288
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private P_ATENDOFSTREAM As Boolean '---------------Indicates if the file's end is reached.
Private P_AUTODETECTENCODING As Boolean '----------Indicates when to perform encoding detection _
                                                   over the file.
Private P_BUFFERLENGTH As Long '-------------------Number of chars in buffer.
Private P_BUFFERSIZE As Single '-------------------Buffer's size x10 MB (0.5 by default).
Private P_DETECTEDENCODING As String '-------------String representation of the auto-detected encoding.
Private P_ENDSTREAMONLINEBREAK As Boolean '--------If true, each stream ends on a line break.
Private P_ISOPENSTREAM As Boolean '----------------Indicates if the object is linked to a file
Private P_LINEBREAK As String '--------------------Holds the char used to end a Stream.
Private P_LINEBREAKMATCHINGBEHAVIOR As EndLineMatchingBehavior 'How to find the next line break.
Private P_PATH As String '-------------------------File path for currente stream.
Private P_UNIFIEDLFOUTPUT As Boolean '-------------If true, the buffer string will be returned _
                                                   with the LF char as Line Break.
Private P_STREAMLENGTH As Long '-------------------File len.
Private P_TEXT As String '-------------------------Holds the current stream's text.
Private P_UTF8ENCODED As Boolean '-----------------Indicates when the file is supposed to be _
                                                   UTF8 encoded.
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
Public Property Get autoDetectEncoding() As Boolean
    autoDetectEncoding = P_AUTODETECTENCODING
End Property
Public Property Let autoDetectEncoding(aValue As Boolean)
    P_AUTODETECTENCODING = aValue
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
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * sizeFactor)
    Buffer = Space$(P_BUFFERLENGTH)
End Property
Public Property Get bufferString() As String
Attribute bufferString.VB_Description = "Gets the text data stored in the buffer."
    If P_UNIFIEDLFOUTPUT Then
        NormalizeLineBreaks
    End If
    If Not P_UTF8ENCODED Then
        bufferString = P_TEXT
    Else
        bufferString = UTF8Decode(P_TEXT)
    End If
End Property
Public Property Get detectedEncoding() As String
    detectedEncoding = P_DETECTEDENCODING
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
Public Property Get LineBreak() As String
Attribute LineBreak.VB_Description = "Returns the character used to end the last received stream. The value is vbNullString when the last stream is not forced to end on line break."
    LineBreak = P_LINEBREAK
End Property
Public Property Get linebreakMatchingBehavior() As EndLineMatchingBehavior
    linebreakMatchingBehavior = P_LINEBREAKMATCHINGBEHAVIOR
End Property
Public Property Let linebreakMatchingBehavior(value As EndLineMatchingBehavior)
    P_LINEBREAKMATCHINGBEHAVIOR = value
End Property
Public Property Get path() As String
    path = P_PATH
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
Attribute unifiedLFOutput.VB_Description = "Determines whether the buffer string is returned using only the LF character as a line feed. Similarly, this property instruct to write files without the Unicode Byte Order Mark."
    unifiedLFOutput = P_UNIFIEDLFOUTPUT
End Property
Public Property Let unifiedLFOutput(value As Boolean)
    P_UNIFIEDLFOUTPUT = value
End Property
Public Property Get utf8EncodedFile() As Boolean
Attribute utf8EncodedFile.VB_Description = "Indicates whether the buffer string is returned as a decoded string, assuming the file is UTF8 encoded."
    utf8EncodedFile = P_UTF8ENCODED
End Property
Public Property Let utf8EncodedFile(value As Boolean)
    P_UTF8ENCODED = value
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
        P_PATH = vbNullString
    End If
End Sub
Private Sub CreateSubFolders(FullFilePath As String)
    Dim cPath As String
    Dim SubFolders() As String
    Dim j As Long
    
    SubFolders() = Split(FullFilePath, "\")
    cPath = SubFolders(LBound(SubFolders))
    For j = LBound(SubFolders) + 1 To UBound(SubFolders) - 1
        cPath = cPath & "\" & SubFolders(j)
        If Not FolderExists(cPath) Then
            MkDir cPath           'Create the subdirectory
        End If
    Next j
End Sub
''' <summary>
''' Approach
'''       Read a Sample of the File: We read the first 1024 bytes of the file
'''       (or the entire file if it’s smaller) into a byte array. This sample
'''       size balances efficiency and accuracy for encoding detection.
'''
'''       Check for BOM: We examine the initial bytes for known BOM sequences
'''       that explicitly indicate the encoding:
'''               UTF-8: EF BB BF
'''
'''               UTF-16 LE: FF FE
'''
'''               UTF-16 BE: FE FF
'''
'''       Heuristic Detection if No BOM: If no BOM is found, we apply heuristic
'''       checks to infer the encoding: UTF-16 Detection: Analyze the frequency
'''       of null bytes (0x00) in even and odd positions. A high percentage of nulls
'''       in odd positions suggests UTF-16 LE, while in even positions suggests UTF-16 BE.
'''
'''       UTF-8 Detection: Validate the byte sequence against UTF-8 encoding rules to
'''       ensure it contains no invalid sequences.
'''
'''       ANSI Detection: Default to ANSI (e.g., Windows-1252) if the byte sequence
'''       doesn’t match UTF-16 or UTF-8 patterns.
''' </summary>
''' <param name="filePath">The path to the actual file.</param>
Public Function DetectEncoding(filePath As String) As String
    Dim fileNum As Integer
    Dim byteData() As Byte
    Dim byteCount As Long
    Dim i As Long
    Dim zeroCountLE As Long
    Dim zeroCountBE As Long
    Dim evenBytes As Long
    Dim oddBytes As Long
    
    ' Open the file in binary mode
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    
    If LOF(fileNum) = 0 Then
        DetectEncoding = "ANSI"
        Exit Function
    End If
    ' Read up to the first 1024 bytes
    byteCount = IIf(LOF(fileNum) < 1024, LOF(fileNum), 1024)
    ReDim byteData(0 To byteCount - 1)
    Get #fileNum, , byteData
    Close #fileNum
    
    ' Check for BOM
    If byteCount >= 3 And byteData(0) = &HEF And byteData(1) = &HBB And byteData(2) = &HBF Then
        DetectEncoding = "UTF-8"
    ElseIf byteCount >= 2 And byteData(0) = &HFF And byteData(1) = &HFE Then
        DetectEncoding = "UTF-16 LE"
    ElseIf byteCount >= 2 And byteData(0) = &HFE And byteData(1) = &HFF Then
        DetectEncoding = "UTF-16 BE"
    Else
        ' Heuristic checks when no BOM is present
        ' Count null bytes in even and odd positions for UTF-16 detection
        For i = 0 To byteCount - 1 Step 2
            If i + 1 < byteCount Then
                If byteData(i + 1) = 0 Then zeroCountLE = zeroCountLE + 1
            End If
            If byteData(i) = 0 Then zeroCountBE = zeroCountBE + 1
        Next i
        
        ' Calculate the number of even and odd position bytes
        evenBytes = (byteCount + 1) \ 2  ' Indices 0, 2, 4, ...
        oddBytes = byteCount \ 2         ' Indices 1, 3, 5, ...
        
        ' Check for UTF-16 LE or BE based on null byte frequency
        If oddBytes > 0 And zeroCountLE / oddBytes > 0.5 Then
            DetectEncoding = "UTF-16 LE"
        ElseIf evenBytes > 0 And zeroCountBE / evenBytes > 0.5 Then
            DetectEncoding = "UTF-16 BE"
        ElseIf IsValidUTF8(byteData) Then
            DetectEncoding = "UTF-8"
        Else
            DetectEncoding = "ANSI"
        End If
    End If
End Function
Private Sub DoubleBufferSize()
    Dim LCSt As Long
    P_BUFFERSIZE = 2 * P_BUFFERSIZE
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * sizeFactor)
    LCSt = P_STREAMLENGTH - P_BUFFERLENGTH
    Select Case LCSt
        Case Is > 0
            Buffer = Space$(P_BUFFERLENGTH)
        Case Else
            Buffer = Space$(P_STREAMLENGTH)
    End Select
End Sub
Private Sub FindEOLcharacter()
    Dim CrCharInStream As Boolean
    Dim LfCharInStream  As Boolean
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
Private Function FolderExists(ByVal filePath As String) As Boolean
    FolderExists = CBool(LenB(Dir(filePath, vbDirectory)))
End Function
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
Private Function IsValidUTF8(byteData() As Byte) As Boolean
    Dim i As Long
    Dim ANSIfile As Boolean
    
    i = 0
    ANSIfile = True
    While i <= UBound(byteData)
        If byteData(i) < 128 Then
            ' Single-byte character (ASCII)
            i = i + 1
        Else
            ANSIfile = False
            If byteData(i) >= 192 And byteData(i) <= 223 Then
                ' Two-byte sequence
                If i + 1 > UBound(byteData) Or byteData(i + 1) < 128 Or byteData(i + 1) > 191 Then
                    IsValidUTF8 = False
                    Exit Function
                End If
                i = i + 2
            ElseIf byteData(i) >= 224 And byteData(i) <= 239 Then
                ' Three-byte sequence
                If i + 2 > UBound(byteData) Or byteData(i + 1) < 128 Or byteData(i + 1) > 191 Or _
                   byteData(i + 2) < 128 Or byteData(i + 2) > 191 Then
                    IsValidUTF8 = False
                    Exit Function
                End If
                i = i + 3
            ElseIf byteData(i) >= 240 And byteData(i) <= 247 Then
                ' Four-byte sequence
                If i + 3 > UBound(byteData) Or byteData(i + 1) < 128 Or byteData(i + 1) > 191 Or _
                   byteData(i + 2) < 128 Or byteData(i + 2) > 191 Or byteData(i + 3) < 128 Or byteData(i + 3) > 191 Then
                    IsValidUTF8 = False
                    Exit Function
                End If
                i = i + 4
            Else
                ' Invalid start byte
                IsValidUTF8 = False
                Exit Function
            End If
        End If
    Wend
    IsValidUTF8 = Not ANSIfile
End Function
Private Sub NormalizeLineBreaks()
    Do While InStrB(1, P_TEXT, vbCrLf, vbBinaryCompare)
        P_TEXT = Join$(Split(P_TEXT, vbCrLf), vbLf)
    Loop
    Do While InStrB(1, P_TEXT, vbCr, vbBinaryCompare)
        P_TEXT = Join$(Split(P_TEXT, vbCr), vbLf)
    Loop
    Do While InStrB(1, P_TEXT, DualLFchar, vbBinaryCompare)
        P_TEXT = Join$(Split(P_TEXT, DualLFchar), vbLf)
    Loop
    P_LINEBREAK = vbLf
End Sub
Public Sub OpenStream(filePath As String)
Attribute OpenStream.VB_Description = "Opens a stream over a text file."
    If P_PATH <> filePath Then
        CreateSubFolders filePath
        If P_AUTODETECTENCODING Then
            P_DETECTEDENCODING = DetectEncoding(filePath)
            P_UTF8ENCODED = (P_DETECTEDENCODING = "UTF-8")
        End If
        FileHandled = FreeFile
        Open filePath For Binary As #FileHandled
        P_ISOPENSTREAM = True
        P_STREAMLENGTH = LOF(FileHandled)
        P_PATH = filePath
        StartVariables
    End If
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
                Buffer = Space$(LCS + 1)
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
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * sizeFactor)
    Buffer = Space$(P_BUFFERLENGTH)
    LCS = P_STREAMLENGTH
    P_ATENDOFSTREAM = False
End Sub
Public Function UTF8Decode(ByVal sStr As String) As String
    Dim l As Long, sUTF8 As CSVArrayList, iChar As Long, iChar2 As Long
    
    On Error GoTo UTF8_Decode_error
    Set sUTF8 = New CSVArrayList
    For l = 1 To LenB(sStr) Step 2
        iChar = Asc(MidB$(sStr, l, 2))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
                iChar2 = Asc(MidB$(sStr, l + 2, 2))
                sUTF8.Add ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
                l = l + 2
            Else
                Dim iChar3 As Integer
                iChar2 = Asc(MidB$(sStr, l + 2, 2))
                iChar3 = Asc(MidB$(sStr, l + 4, 2))
                sUTF8.Add ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
                l = l + 4
            End If
        Else
            sUTF8.Add ChrW$(iChar)
        End If
    Next l
    UTF8Decode = Join$(sUTF8.items, vbNullString)
    Set sUTF8 = Nothing
    Exit Function
UTF8_Decode_error:
    Set sUTF8 = Nothing
    UTF8Decode = vbNullString
End Function
Public Sub WriteBlankLines(Lines As Long, Optional EndLineMark As EndLineChar = 0)
Attribute WriteBlankLines.VB_Description = "Inserts a specified number of blank lines into the current opened text file."
    If P_ISOPENSTREAM Then
        Dim idx As Long
        Select Case EndLineMark
            Case 0
                For idx = 1 To Lines
                    WriteText vbCrLf
                Next idx
            Case 1
                For idx = 1 To Lines
                    WriteText vbCr
                Next idx
            Case 2
                For idx = 1 To Lines
                    WriteText vbLf
                Next idx
        End Select
        P_STREAMLENGTH = LOF(FileHandled)
    End If
End Sub
Public Sub WriteText(ByRef TextData As String)
Attribute WriteText.VB_Description = "Writes the given string to the current opened text file."
    If P_ISOPENSTREAM Then
        If Not P_UTF8ENCODED Then
            Put #FileHandled, , TextData
        Else
            Dim BuffferBytes() As Byte
            BuffferBytes = TextData
            Put #FileHandled, , BuffferBytes
            Erase BuffferBytes
        End If
        P_STREAMLENGTH = LOF(FileHandled)
    End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    P_BUFFERSIZE = 0.5
    P_BUFFERLENGTH = CLng(P_BUFFERSIZE * sizeFactor)
    P_ENDSTREAMONLINEBREAK = False
    P_UTF8ENCODED = False
    P_LINEBREAKMATCHINGBEHAVIOR = EndLineMatchingBehavior.Bidirectional
    P_UNIFIEDLFOUTPUT = False
    Buffer = Space$(P_BUFFERLENGTH)
    NullChar = ChrW(0)
    P_AUTODETECTENCODING = True
End Sub
Private Sub Class_Terminate()
    If P_ISOPENSTREAM Then
        CloseStream
    End If
End Sub
