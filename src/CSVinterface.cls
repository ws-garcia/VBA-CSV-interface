VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVinterface"
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
' GENERAL INFO:
' A small RFC-4180 compliant VBA library to manipulate CSV files at the highest speed.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const CHR_APOSTROPHE As String = "'"
Private Const CHR_DOUBLE_QUOTES As String = """"
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private config As parserConfig '-----------------Parser configuration object
Private CSVstream As ECPTextStream '-------------Enables file Stream
Private P_CSV_DATA As ECPArrayList '-------------Holds the CSV data for current instance
Private P_CSV_HEADER As ECPArrayList '-----------Holds the CSV header
Private P_ERROR_DESC As String '-----------------Holds the Error Description
Private P_ERROR_NUMBER As Long '-----------------Holds the Error number
Private P_ERROR_SOURCE As String '---------------Holds the Error source
Private P_SUCCESSFUL_EXPORT As Boolean
Private P_SUCCESSFUL_IMPORT As Boolean
Private P_VARYING_LENGTHS As Boolean '-----------Indicates if the read CSV has varying number _
                                                 of fields per records
Private P_VECTORS_MAX_BOUND As Long '------------Maximum index of vectors on jagged array
Private P_VECTORS_REGULAR_BOUND As Long '--------Regular upper index of vectors on jagged array
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
' @Common
Private FileHandled As Integer '-----------Pointer to the actual CSV file
Private dTTargets() As Variant '-----------Template's links to fields
Private dTTemplate() As Variant '----------Template
Private TemplateDefined As Boolean
Private TargetsDefined As Boolean
Private RequestedFieldsArray() As Long
Private StreamWithQuotes As Boolean
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
Public Property Get count() As Long
    count = P_CSV_DATA.count
End Property
Public Property Get errDescription() As String
Attribute errDescription.VB_Description = "Gets the description for the last occurred error over the current instance."
    errDescription = P_ERROR_DESC
End Property
Public Property Get errNumber() As Long
Attribute errNumber.VB_Description = "Gets the number for the last occurred error over the current instance."
    errNumber = P_ERROR_NUMBER
End Property
Public Property Get errSource() As String
Attribute errSource.VB_Description = "Gets the source for the last occurred error over the current instance."
    errSource = P_ERROR_SOURCE
End Property
Public Property Get exportSuccess() As Boolean
Attribute exportSuccess.VB_Description = "Gets the status for the last export operation."
    exportSuccess = P_SUCCESSFUL_EXPORT
End Property
Public Property Get csvHeader() As ECPArrayList
Attribute csvHeader.VB_Description = "Gets the imported CSV/TSV headers."
    Set csvHeader = P_CSV_HEADER
End Property
Public Property Get importSuccess() As Boolean
Attribute importSuccess.VB_Description = "Gets the status for the last import operation. Returns True if it was success."
    importSuccess = P_SUCCESSFUL_IMPORT
End Property
Public Property Get item(ParamArray Index() As Variant) As Variant
Attribute item.VB_Description = "Gets a field, or an array with an entire record, from the result array on the current instance."
Attribute item.VB_UserMemId = 0
    Select Case UBound(Index)
        Case 0
            item = P_CSV_DATA(CLng(Index(0)))
        Case Is >= 1
            item = P_CSV_DATA(CLng(Index(0)))(CLng(Index(1)))
    End Select
End Property
Public Property Get parseConfig() As parserConfig
Attribute parseConfig.VB_Description = "Gets or sets, the parser configuration object."
    Set parseConfig = config
End Property
Public Property Set parseConfig(configuration As parserConfig)
    Set config = configuration
End Property
Public Property Get rectangularResults() As Boolean
Attribute rectangularResults.VB_Description = "Gets the result array regularity on the current instance. If true, the internal array is not irregular."
    rectangularResults = Not P_VARYING_LENGTHS
End Property
Public Property Get vectorsBound() As Long
Attribute vectorsBound.VB_Description = "Gets the regular bound of the vectors in the result array on the current instance."
    vectorsBound = P_VECTORS_REGULAR_BOUND
End Property
Public Property Get vectorsMaxBound() As Long
Attribute vectorsMaxBound.VB_Description = "Gets the max bound of the vectors in the result array on the current instance."
    vectorsMaxBound = P_VECTORS_MAX_BOUND
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#

Private Sub Class_Initialize()
    ResetToDefault
End Sub

Private Sub Class_Terminate()
    Set P_CSV_DATA = Nothing
    Set P_CSV_HEADER = Nothing
End Sub

Public Sub ClearData()
    Set P_CSV_DATA = New ECPArrayList
    Set P_CSV_HEADER = New ECPArrayList
End Sub

Public Sub CreateJagged(ByRef ArrVar() As Variant, ArraySize As Long, VectorSize As Long)
    Dim Vector() As String
    Dim arrPointer As Long
    
    ReDim ArrVar(0 To ArraySize)
    ReDim Vector(0 To VectorSize)

    For arrPointer = 0 To ArraySize
        ArrVar(arrPointer) = Vector()
    Next arrPointer
    Erase Vector
End Sub

Public Sub DumpToArray(OutPutArray() As Variant)
    If P_SUCCESSFUL_IMPORT Then
        Dim UB1 As Long, ub2 As Long
        Dim Counter1 As Long, Counter2 As Long
        
        UB1 = P_CSV_DATA.count - 1
        ReDim OutPutArray(0 To UB1, 0 To P_VECTORS_MAX_BOUND)
        
        For Counter1 = 0 To UB1
            ub2 = UBound(P_CSV_DATA(Counter1))
            For Counter2 = 0 To ub2
                OutPutArray(Counter1, Counter2) = P_CSV_DATA(Counter1)(Counter2)
            Next Counter2
        Next Counter1
    End If
End Sub

Public Sub DumpToJaggedArray(OutPutArray() As Variant)
    If P_SUCCESSFUL_IMPORT Then
        Dim UB1 As Long
        Dim Counter1 As Long
        
        UB1 = P_CSV_DATA.count - 1
        ReDim OutPutArray(0 To UB1)
        
        For Counter1 = 0 To UB1
            OutPutArray(Counter1) = P_CSV_DATA(Counter1)
        Next Counter1
    End If
End Sub

Public Sub DumpToSheet(Optional WBookName As String, _
                        Optional SheetName As String, _
                        Optional rngName As String = "A1")
    Dim WBook As Workbook
    Dim outputSheet As Worksheet
    Dim OutputRange As Range
    
    On Error Resume Next
    If P_SUCCESSFUL_IMPORT Then
        EnableOptimization
        '@------------------------------------------------------
        'Set the target Workbook
        If WBookName = vbNullString Then
            Set WBook = ThisWorkbook
        ElseIf Not IsWorkbookOpen(WBookName) Then
            Set WBook = Workbooks.Add
        Else
            Set WBook = Workbooks(WBookName)
        End If
        '@------------------------------------------------------
        'Set the target Worksheet
        If IsSheetInWorkbook(SheetName, WBook) Then
            Set outputSheet = WBook.Sheets(SheetName)
        Else
            Set outputSheet = WBook.Sheets.Add
        End If
        '@------------------------------------------------------
        'Dump the data
        Dim tmpOutputArray() As String
        Dim UB1 As Long, ub2 As Long
        Dim Counter1 As Long, Counter2 As Long
        
        UB1 = P_CSV_DATA.count - 1
        ReDim tmpOutputArray(0 To UB1, 0 To P_VECTORS_MAX_BOUND)
        
        For Counter1 = 0 To UB1
            ub2 = UBound(P_CSV_DATA(Counter1))
            For Counter2 = 0 To ub2
                tmpOutputArray(Counter1, Counter2) = P_CSV_DATA(Counter1)(Counter2)
            Next Counter2
        Next Counter1
        '@------------------------------------------------------
        'Set the target Range
        Set OutputRange = outputSheet.Range(rngName) _
                            .Resize _
                            ( _
                            UB1 + 1, _
                            P_VECTORS_MAX_BOUND + 1 _
                            )
        OutputRange.Value2 = tmpOutputArray
        Erase tmpOutputArray
        EnableOptimization False
    End If
End Sub

Private Sub EnableOptimization(Optional Optimize As Boolean = True)
    If Optimize Then
        '@------------------------------------------------------
        'Optimize resource consumption
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        '@------------------------------------------------------
        'Turn Off optimization
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
End Sub

Public Sub ExportToCSV(ByRef csvArray As Variant, _
                        Optional pconfig As parserConfig = Nothing, _
                        Optional PassControlToOS As Boolean = True)
    On Error GoTo ErrHandler_ExportToCSV
    If PassControlToOS Then DoEvents 'Pass the control to the Operative System
    If IsArray(csvArray) Then
        Dim curConfig As parserConfig
        '@----------------------------------------------------------------------------
        'Open a conection to the file and seek its end
        If pconfig Is Nothing Then
            Set curConfig = config
        Else
            Set curConfig = pconfig
        End If
        Dim CSVstream As ECPTextStream
        Dim EOLB As EndLineChar
        Set CSVstream = New ECPTextStream
        If FileExists(curConfig.path) Then
            GuessDelimiters curConfig
            CSVstream.OpenStream curConfig.path
            If CSVstream.streamLength > 0 Then
                Select Case curConfig.recordsDelimiter
                    Case vbCrLf
                        CSVstream.SeekPointer CSVstream.streamLength - 3
                    Case Else
                        CSVstream.SeekPointer CSVstream.streamLength - 1
                End Select
                '@----------------------------------------------------------------------------
                'Check line break at EOF
                CSVstream.ReadText
                If CSVstream.bufferString <> curConfig.recordsDelimiter Then
                    Select Case curConfig.recordsDelimiter
                        Case vbCrLf
                            EOLB = EndLineChar.CRLF
                        Case vbCr
                            EOLB = EndLineChar.CR
                        Case Else
                            EOLB = EndLineChar.LF
                    End Select
                    CSVstream.WriteBlankLines 1, EOLB
                End If
            End If
        Else
            CSVstream.OpenStream curConfig.path
        End If
        '@----------------------------------------------------------------------------
        'Export
        CSVstream.SeekPointer CSVstream.streamLength
        CSVstream.WriteText JoinRecordsFields(csvArray, curConfig)
        Set CSVstream = Nothing
        If P_ERROR_NUMBER = 0& Then
            Close #FileHandled
            P_SUCCESSFUL_EXPORT = True
            P_ERROR_DESC = vbNullString
            P_ERROR_SOURCE = vbNullString
        Else
            GoTo ErrHandler_ExportToCSV
        End If
    Else
        P_SUCCESSFUL_EXPORT = False
        P_ERROR_DESC = "[CSV file Export]: The passed argument isn't an array."
        P_ERROR_NUMBER = vbObjectError + 9009
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Exit Sub
ErrHandler_ExportToCSV:
    Close #FileHandled
    P_SUCCESSFUL_EXPORT = False
    P_ERROR_DESC = "[CSV file Export]: " & P_ERROR_DESC
End Sub

Private Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = CBool(LenB(Dir(FilePath, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)))
End Function

Public Function GetDataFromCSV(csvPathAndFilename As String) As String
    
    On Error GoTo ErrHandler_GetDataFromCSV
    FileHandled = FreeFile
    Open csvPathAndFilename For Binary As #FileHandled
    GetDataFromCSV = SPACE$(LOF(FileHandled))
    Get #FileHandled, , GetDataFromCSV
    Close #FileHandled
ErrHandler_GetDataFromCSV:
    P_ERROR_DESC = "[Get Data From CSV]: " & Err.Description
    P_ERROR_NUMBER = Err.Number
    P_ERROR_SOURCE = Err.Source
End Function

Private Function GetLongFilters(ByRef tmpFilters() As Variant, _
                                ByRef header As ECPArrayList) As Long()
    Dim tmpResult() As Long
    Dim tmpColl As Collection
    Dim trqidx As Long, arridx As Long
    Dim filterCount As Long
    Dim matchFound As Boolean

    If Not header Is Nothing Then
        Set tmpColl = New Collection
        For trqidx = LBound(tmpFilters) To UBound(tmpFilters)
            If IsNumeric(tmpFilters(trqidx)) Then
                On Error Resume Next
                tmpColl.Add CStr(CLng(tmpFilters(trqidx))), CStr(CLng(tmpFilters(trqidx)))
            Else
                arridx = 0
                Do
                    matchFound = (CStr(tmpFilters(trqidx)) = CStr(header(0)(arridx)))
                    If matchFound Then
                        On Error Resume Next
                        tmpColl.Add CStr(arridx + 1), CStr(arridx + 1)
                    End If
                    arridx = arridx + 1
                Loop While Not matchFound And arridx <= header.count - 1
            End If
        Next trqidx
        If tmpColl.count >= 1 Then
            ReDim tmpResult(0 To tmpColl.count - 1)
            For arridx = 0 To tmpColl.count - 1
                tmpResult(arridx) = CLng(tmpColl.item(arridx + 1))
            Next arridx
        End If
        Err.Clear
        GetLongFilters = tmpResult
    End If
End Function

Public Sub GuessDelimiters(ByRef confObj As parserConfig)
    If confObj.path <> vbNullString Then
        '@--------------------------------------------------------------------------------
        'Guess records delimiter
        Dim StreamObj As ECPTextStream, tmpArr() As String
        Dim UBlinesEnds() As Long, LinesEnds() As String
        Dim parsedLines() As Long, delIdx As Long
        Dim tmpDelimiters() As String, guessResults As ECPArrayList
        Dim emptyParam() As Variant, StartRec As Long, EndRec As Long
        
        ReDim UBlinesEnds(0 To 2)
        ReDim LinesEnds(0 To 2)
        ReDim emptyParam(-1 To -1) 'To emule empty ParamArray parameter
        tmpDelimiters() = config.delimitersToGuess
        ReDim parsedLines(LBound(tmpDelimiters) To UBound(tmpDelimiters))
        LinesEnds(0) = vbCrLf
        LinesEnds(1) = vbCr
        LinesEnds(2) = vbLf
        Set StreamObj = New ECPTextStream
        StreamObj.endStreamOnLineBreak = True
        StreamObj.OpenStream confObj.path
        StreamObj.ReadText
        tmpArr = Split(StreamObj.bufferString, LinesEnds(0))
        UBlinesEnds(0) = UBound(tmpArr) 'vbCrLf
        tmpArr = Split(StreamObj.bufferString, LinesEnds(1))
        UBlinesEnds(1) = UBound(tmpArr) 'vbCr
        tmpArr = Split(StreamObj.bufferString, LinesEnds(2))
        UBlinesEnds(2) = UBound(tmpArr) 'vbLf
        Set StreamObj = Nothing
        StartRec = confObj.startingRecord
        EndRec = confObj.endingRecord
        '@--------------------------------------------------------------------------------
        'Guess using 10 samples lines
        confObj.startingRecord = 1
        confObj.endingRecord = 10
        confObj.recordsDelimiter = LinesEnds(MaxIndexVal(UBlinesEnds))
        For delIdx = LBound(tmpDelimiters) To UBound(tmpDelimiters)
            confObj.fieldsDelimiter = tmpDelimiters(delIdx)
            Set guessResults = New ECPArrayList
            StreamParseCSV confObj, guessResults, emptyParam
            If Not guessResults Is Nothing Then
                parsedLines(delIdx) = UBound(guessResults(0))
            Else
                parsedLines(delIdx) = 0
            End If
        Next delIdx
        Set guessResults = Nothing
        confObj.fieldsDelimiter = tmpDelimiters(MaxIndexVal(parsedLines))
        confObj.startingRecord = StartRec
        confObj.endingRecord = EndRec
        Erase UBlinesEnds
        Erase LinesEnds
        Erase tmpDelimiters
        Erase tmpArr
    End If
End Sub
Public Function ImportFromCSV(configObj As parserConfig, _
                            ParamArray FilterColumns() As Variant) As CSVinterface
    DoEvents 'Pass the control to the Operative System
    With configObj
        If FileExists(.path) Then
            If UBound(.dTypingTemplate) >= UBound(.dTypingLinks) Then
                Dim tmpRequested() As Variant
                Dim curConfig As parserConfig
                Dim emptyParam() As Variant
                ReDim emptyParam(-1 To -1) 'Emule empty ParamArray
        
                If .delimitersGuessing Then
                    GuessDelimiters configObj
                End If
                tmpRequested() = FilterColumns
                If .headers Then
                    Dim StartRec As Long, EndRec As Long
                    '@--------------------------------------------------------------------------------
                    'Read the whole header
                    Set P_CSV_HEADER = New ECPArrayList
                    Set curConfig = .CopyConfig
                    curConfig.startingRecord = 1
                    curConfig.endingRecord = 1
                    StreamParseCSV curConfig, P_CSV_HEADER, emptyParam
                    If Not P_CSV_HEADER Is Nothing Then
                        P_CSV_HEADER.ShrinkBuffer
                    End If
                End If
                If P_ERROR_NUMBER <> 0& Then GoTo ErrHandler_ImportFromCSV
                '@--------------------------------------------------------------------------------
                'Read the CSV data
                Set P_CSV_DATA = New ECPArrayList
                StreamParseCSV configObj, P_CSV_DATA, tmpRequested
                If P_ERROR_NUMBER = 0& Then
                    'Save pointers
                    P_SUCCESSFUL_IMPORT = True
                    P_ERROR_DESC = vbNullString
                    P_ERROR_SOURCE = vbNullString
                Else
                    GoTo ErrHandler_ImportFromCSV
                End If
            Else
                P_SUCCESSFUL_IMPORT = False
                P_ERROR_DESC = "[CSV file Import]: the config object is not linked to a CSV file. " _
                                & "Ensure set the path property to valid CSV before import data."
                P_ERROR_NUMBER = vbObjectError + 9007
                P_ERROR_SOURCE = "CSVinterface"
            End If
        End If
    End With
    Set ImportFromCSV = Me
    Exit Function
ErrHandler_ImportFromCSV:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[CSV file Import]: " & P_ERROR_DESC
End Function

Public Function ImportFromCSVString(ByRef CSVstring As String, _
                                    configObj As parserConfig, _
                                    ParamArray FilterColumns() As Variant) As CSVinterface
    DoEvents 'Pass the control to the Operative System
    With configObj
        If UBound(.dTypingTemplate) >= UBound(.dTypingLinks) Then
            If LenB(CSVstring) = 0 Then
                Exit Function
            End If
            Dim tmpRequested() As Variant
            Dim curConfig As parserConfig
            Dim emptyParam() As Variant
            ReDim emptyParam(-1 To -1) 'Emule empty ParamArray
        
            tmpRequested() = FilterColumns
            If .headers Then
                '@--------------------------------------------------------------------------------
                'Read the header
                Set P_CSV_HEADER = New ECPArrayList
                Set curConfig = .CopyConfig
                curConfig.startingRecord = 1
                curConfig.endingRecord = 1
                ParseCSVstring CSVstring, curConfig, P_CSV_HEADER, emptyParam
                If Not P_CSV_HEADER Is Nothing Then
                    P_CSV_HEADER.ShrinkBuffer
                End If
            End If
            If P_ERROR_NUMBER <> 0& Then GoTo ErrHandler_ImportFromString
            '@--------------------------------------------------------------------------------
            'Read the CSV data
            Set P_CSV_DATA = New ECPArrayList
            ParseCSVstring CSVstring, configObj, P_CSV_DATA, tmpRequested
            '@----------------------------------------------------------------------------
            If P_ERROR_NUMBER = 0& Then
                'Save pointers
                P_SUCCESSFUL_IMPORT = True
                P_ERROR_DESC = vbNullString
                P_ERROR_SOURCE = vbNullString
            Else
                GoTo ErrHandler_ImportFromString
            End If
        End If
    End With
    Set ImportFromCSVString = Me
    Exit Function
ErrHandler_ImportFromString:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[CSV file Import]: " & P_ERROR_DESC
End Function

Public Function IsJaggedArray(Arr As Variant) As Boolean
    On Error GoTo IsJaggedArray_Err_Handler
    If IsArray(Arr) Then
        If MultiDimensional(Arr) Then
            IsJaggedArray = False
        Else
            Dim BoundingTest As Variant
            BoundingTest = Arr(LBound(Arr))
            BoundingTest = LBound(BoundingTest)
            IsJaggedArray = True
        End If
    End If
    Exit Function
IsJaggedArray_Err_Handler:
    IsJaggedArray = False
End Function

Private Function IsSheetInWorkbook(SheetName As String, WBook As Workbook) As Boolean
    With WBook
        On Error Resume Next
        IsSheetInWorkbook = (.Sheets(SheetName).Name = SheetName)
        On Error GoTo 0
    End With
End Function

Private Function IsWorkbookOpen(WBookName As String) As Boolean
    Dim WBook As Workbook, BookMatching As Boolean
    
    On Error Resume Next
    For Each WBook In Workbooks
        BookMatching = (WBook.Name = WBookName)
    Next
    IsWorkbookOpen = BookMatching
End Function

Public Sub JaggedToTwoDimArray(ByRef JaggedArray() As Variant, ByRef TwoDimArray() As Variant)
    Dim UBj1 As Long, LBj1 As Long
    Dim UBj2 As Long, LBj2 As Long
    Dim MaxDim1 As Long, MaxDim2 As Long
    Dim jgdCounter1 As Long, jgdCounter2 As Long
    Dim tmpDimension As Long, tdimCounter As Long
    
    On Error GoTo JaggedToTwoDimArray_Err_Handler
    LBj1 = LBound(JaggedArray)
    UBj1 = UBound(JaggedArray)
    MaxDim1 = Abs(UBj1 - LBj1) 'Dimension One in base 0
    MaxDim2 = Abs(UBound(JaggedArray(LBj1)) - LBound(JaggedArray(LBj1))) 'Dimension Two in base 0
    ReDim TwoDimArray(0 To MaxDim1, 0 To MaxDim2)
    For jgdCounter1 = LBj1 To UBj1
        LBj2 = LBound(JaggedArray(jgdCounter1))
        UBj2 = UBound(JaggedArray(jgdCounter1))
        tdimCounter = 0
        tmpDimension = Abs(UBj2 - LBj2)
        If tmpDimension > MaxDim2 Then
            MaxDim2 = tmpDimension
            ReDim Preserve TwoDimArray(0 To MaxDim1, 0 To MaxDim2)
        End If
        For jgdCounter2 = LBj2 To UBj2
            TwoDimArray(jgdCounter1, tdimCounter) = JaggedArray(jgdCounter1)(jgdCounter2)
            tdimCounter = tdimCounter + 1&
        Next jgdCounter2
        Erase JaggedArray(jgdCounter1) 'Free memory
    Next jgdCounter1
JaggedToTwoDimArray_Err_Handler:
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
    P_ERROR_SOURCE = Err.Source
End Sub

Private Function JoinRecordsFields(ByRef RecordsArray As Variant, _
                                    pconfig As parserConfig) As String
    Dim Buffer() As String
    Dim CoerceChr As String
    Dim ConcatenatedArray() As String
    Dim EscapeT As EscapeTokens
    Dim FldDelimiter As String
    Dim iLCounter As Long, jLCounter As Long
    Dim JoinBuffer(0 To 2) As String
    Dim LB1 As Long, UB1 As Long
    Dim LB2 As Long, ub2 As Long
    Dim recDelimiter As String
    
    On Error GoTo JoinRecordsFields_Error
    '@----------------------------------------------------------------------------
    'Expand configuration object
    EscapeT = pconfig.escapeToken
    FldDelimiter = pconfig.fieldsDelimiter
    recDelimiter = pconfig.recordsDelimiter
    '@----------------------------------------------------------------------------
    'Set coerce char
    Select Case EscapeT
        Case 0
            CoerceChr = vbNullString
        Case 1
            CoerceChr = CHR_APOSTROPHE
        Case Else
            CoerceChr = CHR_DOUBLE_QUOTES
    End Select
    '@----------------------------------------------------------------------------
    'Set the char buffer
    JoinBuffer(0) = CoerceChr
    JoinBuffer(1) = vbNullString
    JoinBuffer(2) = CoerceChr
    If MultiDimensional(RecordsArray) Then
        LB1 = LBound(RecordsArray, 1)
        LB2 = LBound(RecordsArray, 2)
        UB1 = UBound(RecordsArray, 1)
        ub2 = UBound(RecordsArray, 2)
        '@----------------------------------------------------------------------------
        'Set array sizes
        ReDim ConcatenatedArray(LB1 To UB1)
        ReDim Buffer(LB2 To ub2)
        '@----------------------------------------------------------------------------
        'Concatenate fields
        For iLCounter = LB1 To UB1
            For jLCounter = LB2 To ub2
                If InStrB(1, RecordsArray(iLCounter, jLCounter), FldDelimiter) Then
                    JoinBuffer(1) = RecordsArray(iLCounter, jLCounter)
                    Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                ElseIf InStrB(1, RecordsArray(iLCounter, jLCounter), CoerceChr) Then
                    JoinBuffer(1) = RecordsArray(iLCounter, jLCounter)
                    Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                ElseIf InStrB(1, RecordsArray(iLCounter, jLCounter), recDelimiter) Then
                    JoinBuffer(1) = RecordsArray(iLCounter, jLCounter)
                    Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                Else
                    Buffer(jLCounter) = RecordsArray(iLCounter, jLCounter)
                End If
                JoinBuffer(1) = vbNullString
            Next jLCounter
            ConcatenatedArray(iLCounter) = Join$(Buffer, FldDelimiter)
        Next iLCounter
        '@----------------------------------------------------------------------------
        'Concatenate records
        JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
        '@----------------------------------------------------------------------------
        'Free the memory
        Erase ConcatenatedArray
        Erase Buffer
    Else
        '@----------------------------------------------------------------------------
        'Check for jagged arrays
        LB1 = LBound(RecordsArray, 1)
        UB1 = UBound(RecordsArray, 1)
        If Not IsJaggedArray(RecordsArray) Then
            '@----------------------------------------------------------------------------
            'Set array sizes
            ReDim ConcatenatedArray(LB1 To UB1)
            ReDim Buffer(0)
            '@----------------------------------------------------------------------------
            'Concatenate fields
            For iLCounter = LB1 To UB1
                If InStrB(1, RecordsArray(iLCounter), FldDelimiter) Then
                    JoinBuffer(1) = RecordsArray(iLCounter)
                    Buffer(0) = Join$(JoinBuffer, vbNullString)
                ElseIf InStrB(1, RecordsArray(iLCounter), CoerceChr) Then
                    JoinBuffer(1) = RecordsArray(iLCounter)
                    Buffer(0) = Join$(JoinBuffer, vbNullString)
                ElseIf InStrB(1, RecordsArray(iLCounter), recDelimiter) Then
                    JoinBuffer(1) = RecordsArray(iLCounter)
                    Buffer(0) = Join$(JoinBuffer, vbNullString)
                Else
                    Buffer(0) = RecordsArray(iLCounter)
                End If
                JoinBuffer(1) = vbNullString
                ConcatenatedArray(iLCounter) = Buffer(0)
            Next iLCounter
            '@----------------------------------------------------------------------------
            'Concatenate records
            JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
            '@----------------------------------------------------------------------------
            'Free the memory
            Erase ConcatenatedArray
            Erase Buffer
        Else
            '@----------------------------------------------------------------------------
            'Set array sizes
            ReDim ConcatenatedArray(LB1 To UB1)
            '@----------------------------------------------------------------------------
            'Concatenate fields
            Dim tmpBuffer As Variant
            For iLCounter = LB1 To UB1
                tmpBuffer = RecordsArray(iLCounter)
                LB2 = LBound(tmpBuffer)
                ub2 = UBound(tmpBuffer)
                ReDim Buffer(LB2 To ub2)
                For jLCounter = LB2 To ub2
                    If InStrB(1, tmpBuffer(jLCounter), FldDelimiter) Then
                        JoinBuffer(1) = tmpBuffer(jLCounter)
                        Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                    ElseIf InStrB(1, tmpBuffer(jLCounter), CoerceChr) Then
                        JoinBuffer(1) = tmpBuffer(jLCounter)
                        Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                    ElseIf InStrB(1, tmpBuffer(jLCounter), recDelimiter) Then
                        JoinBuffer(1) = tmpBuffer(jLCounter)
                        Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                    Else
                        Buffer(jLCounter) = tmpBuffer(jLCounter)
                    End If
                    JoinBuffer(1) = vbNullString
                Next jLCounter
                ConcatenatedArray(iLCounter) = Join$(Buffer, FldDelimiter)
            Next iLCounter
            '@----------------------------------------------------------------------------
            'Concatenate records
            JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
            '@----------------------------------------------------------------------------
            'Free the memory
            Erase ConcatenatedArray
            Erase Buffer
        End If
    End If
    Exit Function
JoinRecordsFields_Error:
    JoinRecordsFields = vbNullString
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
    P_ERROR_SOURCE = Err.Source
End Function

Private Function MaxIndexVal(ByRef Values() As Long) As Long
    Dim tmpVal As Long, idx0 As Long, idx As Long
    
    idx0 = LBound(Values)
    idx = idx0
    tmpVal = Values(idx0)
    For idx0 = idx0 + 1 To UBound(Values)
        If Values(idx0) > tmpVal Then
            idx = idx0
            tmpVal = Values(idx0)
        End If
    Next idx0
    MaxIndexVal = idx
End Function

Private Function MultiDimensional(CheckArray As Variant) As Boolean
    On Error GoTo ErrHandler_MultiDimensional

    If UBound(CheckArray, 2) > 0 Then
        MultiDimensional = True
    End If
    Exit Function
ErrHandler_MultiDimensional:
    MultiDimensional = False '1 dimension
End Function

Private Sub ParseCSVstring(ByRef csvText As String, _
                            configObj As parserConfig, _
                            outputList As ECPArrayList, _
                            ByRef FilterColumns() As Variant)
    Dim ASCIIcharw As Long
    Dim brCounter As Long
    Dim commToken As Long
    Dim dynamicType As Boolean
    Dim EndOF As Boolean
    Dim EndOfStreamBuffer As Boolean
    Dim EndRecord As Long
    Dim EscapeAscW As Long
    Dim EscapeChr As String
    Dim EscapedWFS As Boolean
    Dim FDAscW As Long
    Dim FieldID As Long
    Dim FieldDelimiter As String
    Dim IgnoreLines As Boolean
    Dim ImportDepth As Long
    Dim ImportSwitch As Boolean
    Dim IncludeRecord As Boolean
    Dim Index As Long, MaxIndex As Long
    Dim IsHeader As Boolean
    Dim KnownEnding As Boolean
    Dim LenCurrentIndex As Long
    Dim MaxReqIndex As Long
    Dim OpenedToken As Boolean
    Dim OverStringPointer As Long, hpointer As Long
    Dim RecordEndReached As Boolean
    Dim RecordsCount As Long
    Dim RecordDelimiter As String
    Dim RecordToken As ECPArrayList
    Dim SearchBeginningMark As Long
    Dim SplittedToken As Boolean
    Dim StartRecord As Long
    Dim tmpCSV() As String
    Dim tmpRequested() As Variant
    Dim tmpToken As ECPArrayList
    Dim TokenBeginningPos As Long
    Dim TokenEndingPos As Long
    Dim TokenEndReached As Boolean
    
    On Error GoTo ParseCriticalCSV_Error_Handler
    '@----------------------------------------------------------------------------
    'Start variables
    SearchBeginningMark = 1&
    RecordsCount = 0&
    P_VECTORS_REGULAR_BOUND = 0&
    P_VECTORS_MAX_BOUND = 0&
    Set tmpToken = New ECPArrayList
    TokenEndReached = False
    RecordEndReached = False
    With configObj
        commToken = AscW(.commentsToken)
        dynamicType = .dynamicTyping
        dTTemplate() = .dTypingTemplate
        dTTargets() = .dTypingLinks
        IgnoreLines = .catchMeaninglessRecords
        TemplateDefined = .dTTemplateDefined
        TargetsDefined = .dTTemplateLinksDefined
        FieldDelimiter = .fieldsDelimiter
        RecordDelimiter = .recordsDelimiter
        EndRecord = .endingRecord
        StartRecord = .startingRecord
    End With
    tmpRequested() = FilterColumns
    Select Case UBound(tmpRequested)
        Case -1
            ReDim RequestedFieldsArray(0)
            RequestedFieldsArray(0) = -1
        Case Else
            RequestedFieldsArray() = GetLongFilters(tmpRequested, P_CSV_HEADER)
    End Select
    
    '@----------------------------------------------------------------------------
    'Set variables
    OverStringPointer = SearchBeginningMark
    MaxReqIndex = UBound(RequestedFieldsArray)
    '@----------------------------------------------------------------------------
    'Check if user want Import all records
    KnownEnding = True
    Select Case EndRecord
        Case 0&
            If StartRecord >= 1 Then
                KnownEnding = False
            Else
                ImportDepth = 1&
            End If
        Case Is < StartRecord
            ImportDepth = 1&
        Case Else
            ImportDepth = EndRecord - StartRecord + 1&
    End Select
    If StartRecord = 1& Then
        If configObj.headersOmission And configObj.headers Then
            If ImportDepth > 1& Or EndRecord = 0& Then
                StartRecord = 2&
                ImportDepth = ImportDepth - 1&
            End If
        End If
    End If
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case configObj.escapeToken
        Case 1
            EscapeChr = CHR_APOSTROPHE
        Case Else
            EscapeChr = CHR_DOUBLE_QUOTES
    End Select
    EscapeAscW = AscW(EscapeChr)
    FDAscW = AscW(FieldDelimiter)
    Set RecordToken = New ECPArrayList
    Index = 0
    '@----------------------------------------------------------------------------
    'Populate a temp array
    tmpCSV() = Split(csvText, RecordDelimiter)
    MaxIndex = UBound(tmpCSV)
    brCounter = 0
    Do While LenB(tmpCSV(MaxIndex - brCounter)) = 0
        brCounter = brCounter + 1
    Loop
    MaxIndex = MaxIndex - brCounter
    EndOfStreamBuffer = (Index > MaxIndex)
    '@----------------------------------------------------------------------------
    'Loop through CSV
    Do
        IsHeader = (StartRecord = 1 And RecordsCount = 0)
        If IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            SkipUnwantedLines Index, MaxIndex, tmpCSV, commToken, FDAscW
        End If
        LenCurrentIndex = LenB(tmpCSV(Index))
        '@----------------------------------------------------------------------------
        'Restart pointers
        RecordEndReached = False
        FieldID = 0
        IncludeRecord = (RecordsCount + 1& >= StartRecord)
        Do
            '@----------------------------------------------------------------------------
            'Process the current Token
            TokenEndReached = False
            TokenBeginningPos = OverStringPointer
            Select Case OverStringPointer
                Case Is < LenCurrentIndex
                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                    '@----------------------------------------------------------------------------
                    'Ignore spaces, tabs and remanents LF
                    Select Case FDAscW
                        Case Is <> 9
                            Do While ASCIIcharw = 32 Or ASCIIcharw = 9 Or ASCIIcharw = 10
                                OverStringPointer = OverStringPointer + 2&
                                If OverStringPointer > LenCurrentIndex Then Exit Do
                                ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                            Loop
                        Case Else
                            Do While ASCIIcharw = 32 Or ASCIIcharw = 10
                                OverStringPointer = OverStringPointer + 2&
                                If OverStringPointer > LenCurrentIndex Then Exit Do
                                ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                            Loop
                    End Select
                    OpenedToken = (ASCIIcharw = EscapeAscW)
                Case Else 'Avoid over runs
                    OpenedToken = False
            End Select
            If OpenedToken Then
                '@----------------------------------------------------------------------------
                'Try to escape the Token
                TokenBeginningPos = OverStringPointer + 2&
                SearchBeginningMark = TokenBeginningPos
                hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'EOFld marker
                Do While (Not TokenEndReached And Not EndOF)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not SplittedToken Then SplittedToken = (hpointer = 0)
                    Do While hpointer = 0& 'Escape char may be on subsequent Index
                        If IncludeRecord Then
                            Select Case SearchBeginningMark
                                Case 1 'Store the full Token from the current Index
                                    tmpToken.Add tmpCSV(Index)
                                Case Else 'Store the Right most Token portion from the current Index
                                    tmpToken.Add RightB$(tmpCSV(Index), LenCurrentIndex - SearchBeginningMark + 1)
                            End Select
                        End If
                        tmpCSV(Index) = vbNullString
                        Index = Index + 1&
                        EndOfStreamBuffer = (Index > MaxIndex)
                        If Not EndOfStreamBuffer Then
                            LenCurrentIndex = LenB(tmpCSV(Index))
                            SearchBeginningMark = 1
                            hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                        Else
                            hpointer = 0&
                        End If
                    Loop
                    If hpointer = 0& And EndOfStreamBuffer Then GoTo ParseCriticalCSV_MissingEscapeChar 'Missing escape char
                    OverStringPointer = hpointer + 2&
                    If OverStringPointer < LenCurrentIndex Then
                        ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces, tabs and remanents LF
                        Select Case FDAscW
                            Case Is <> 9
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 9 Or ASCIIcharw = 10
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 10
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                        End Select
                        '@----------------------------------------------------------------------------
                        'Check if we are at the end of the field
                        EscapedWFS = (ASCIIcharw = FDAscW)
                    Else
                        EscapedWFS = False
                    End If
                    If EscapedWFS Then 'Escape char succeeded by fields separator char
                        TokenEndingPos = hpointer - 1&
                        TokenEndReached = True
                        If IncludeRecord Then
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case SplittedToken
                                Case False 'Just save Token
                                    If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                        Select Case dynamicType
                                            Case False
                                                RecordToken.Add MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                TokenEndingPos - TokenBeginningPos + 1)
                                            Case Else
                                                RecordToken.Add TypeData(MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                        TokenEndingPos - TokenBeginningPos + 1), _
                                                                        FieldID, _
                                                                        IsHeader)
                                        End Select
                                    End If
                                Case Else 'The Token is stored on tmpToken object
                                    tmpToken.Add LeftB$(tmpCSV(Index), TokenEndingPos)
                                    If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                        Select Case dynamicType
                                            Case False
                                                RecordToken.Add Join$(tmpToken.items, RecordDelimiter) 'Save whole Token
                                            Case Else
                                                RecordToken.Add TypeData(Join$(tmpToken.items, _
                                                                            RecordDelimiter), _
                                                                        FieldID, _
                                                                        IsHeader)
                                        End Select
                                    End If
                                    tmpToken.Clear 'Clear tmpToken content
                            End Select
                        End If
                        SplittedToken = False 'Restart Token status
                        OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                        RecordEndReached = False
                    Else
                        If OverStringPointer >= LenCurrentIndex Then  'Escape char succeeded
                                                                      'by records separator char
                            TokenEndingPos = hpointer - 1&
                            TokenEndReached = True
                            RecordEndReached = True 'EOR
                            If IncludeRecord Then
                                '@----------------------------------------------------------------------------
                                'Save the Token
                                Select Case SplittedToken
                                    Case False 'Just save Token
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            Select Case dynamicType
                                                Case False
                                                    RecordToken.Add MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                    TokenEndingPos - TokenBeginningPos + 1)
                                                Case Else
                                                    RecordToken.Add TypeData(MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                                TokenEndingPos - TokenBeginningPos + 1), _
                                                                            FieldID, _
                                                                            IsHeader)
                                            End Select
                                        End If
                                    Case Else 'The Token is stored on tmpToken array
                                        tmpToken.Add LeftB$(tmpCSV(Index), OverStringPointer - 3)
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            Select Case dynamicType
                                                Case False
                                                    RecordToken.Add Join$(tmpToken.items, RecordDelimiter) 'Save whole Token
                                                Case Else
                                                    RecordToken.Add TypeData(Join$(tmpToken.items, _
                                                                                RecordDelimiter), _
                                                                                FieldID, _
                                                                                IsHeader)
                                            End Select
                                        End If
                                        tmpToken.Clear 'Clear tmpToken content
                                End Select
                            End If
                            SplittedToken = False 'Restart Token status
                            tmpCSV(Index) = vbNullString
                            Index = Index + 1&
                            EndOfStreamBuffer = (Index > MaxIndex)
                            OverStringPointer = 1&
                        Else
                            If IncludeRecord Then
                                If SplittedToken Then 'Store the Left most Token portion from the current Index
                                    tmpToken.Add LeftB$(tmpCSV(Index), OverStringPointer - 1)
                                End If
                            End If
                            SearchBeginningMark = OverStringPointer
                            hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                            TokenEndReached = False
                        End If
                    End If
                Loop
            Else
                SplittedToken = False
                '@----------------------------------------------------------------------------
                'Try to mark the Starting and Ending position for the field
                hpointer = InStrB(OverStringPointer, tmpCSV(Index), FieldDelimiter)
                If hpointer Then 'Found fields delimiter char
                    TokenEndReached = True
                    TokenEndingPos = hpointer - 1&
                    OverStringPointer = hpointer + 2&
                    RecordEndReached = False
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                            Select Case dynamicType
                                Case False
                                    RecordToken.Add MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                    TokenEndingPos - TokenBeginningPos + 1)
                                Case Else
                                    RecordToken.Add TypeData(MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                TokenEndingPos - TokenBeginningPos + 1), _
                                                                FieldID, _
                                                                IsHeader)
                            End Select
                        End If
                    End If
                Else 'Missing fields delimiter char
                    TokenEndReached = True
                    RecordEndReached = True
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                            Select Case dynamicType
                                Case False
                                    RecordToken.Add RightB$(tmpCSV(Index), LenCurrentIndex _
                                                    - TokenBeginningPos + 1)
                                Case Else
                                    RecordToken.Add TypeData(RightB$(tmpCSV(Index), LenCurrentIndex _
                                                            - TokenBeginningPos + 1), _
                                                            FieldID, _
                                                            IsHeader)
                            End Select
                        End If
                    End If
                    tmpCSV(Index) = vbNullString
                    Index = Index + 1&
                    EndOfStreamBuffer = (Index > MaxIndex)
                    OverStringPointer = 1&
                End If
            End If
            FieldID = FieldID + 1&
        Loop While (Not RecordEndReached And Not EndOfStreamBuffer)
        Select Case P_VECTORS_REGULAR_BOUND
            Case Is > 0
                If RecordToken.count - 1 > P_VECTORS_MAX_BOUND Then
                    P_VECTORS_MAX_BOUND = RecordToken.count - 1
                    RecordToken.ShrinkBuffer
                End If
            Case Else
                Select Case RecordToken.count
                    Case Is > 0
                        P_VECTORS_REGULAR_BOUND = RecordToken.count - 1
                        RecordToken.ShrinkBuffer
                    Case Else
                        P_VECTORS_REGULAR_BOUND = RecordToken.count
                End Select
                P_VECTORS_MAX_BOUND = P_VECTORS_REGULAR_BOUND
        End Select
        RecordsCount = RecordsCount + 1&
        '@----------------------------------------------------------------------------
        'Start saving on given record
        If IncludeRecord Then
            outputList.Add RecordToken.items
            RecordToken.Reinitialize P_VECTORS_REGULAR_BOUND + 1
        End If
        P_VARYING_LENGTHS = (P_VECTORS_REGULAR_BOUND <> P_VECTORS_MAX_BOUND)
        If IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            SkipUnwantedLines Index, MaxIndex, tmpCSV, commToken, FDAscW
        End If
        EndOfStreamBuffer = (Index > MaxIndex)
        If Not EndOfStreamBuffer Then LenCurrentIndex = LenB(tmpCSV(Index))
        Select Case KnownEnding
            Case False
                ImportSwitch = EndOfStreamBuffer
            Case Else
                ImportSwitch = (outputList.count = ImportDepth)
        End Select
    Loop While Not ImportSwitch
    '@----------------------------------------------------------------------------
    'Check if the array need to be shrinked at the output stage
    outputList.ShrinkBuffer
    Erase tmpCSV
    Set RecordToken = Nothing
    Set tmpToken = Nothing
    Exit Sub
ParseCriticalCSV_Error_Handler:
    Set outputList = Nothing
    Set tmpToken = Nothing
    Erase tmpCSV
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description & " [Check the record #" & RecordsCount + 1 _
                                  & ", field #" & RecordToken.count + 1 _
                                   & " on the source CSV file]."
    Set RecordToken = Nothing
    P_ERROR_SOURCE = Err.Source
    Exit Sub
ParseCriticalCSV_StartingIndexOutOfBound:
    Err.Raise Number:=vbObjectError + 9002, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="The index over which the operation starts is out of bounds. Try start the import on a lower index."
    Resume ParseCriticalCSV_Error_Handler
ParseCriticalCSV_MissingEscapeChar:
    Err.Raise Number:=vbObjectError + 9003, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="Missing some escape char. Check the data and try again."
    Resume ParseCriticalCSV_Error_Handler
End Sub

Private Sub Quicksort(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortColumn As Long, _
                        Optional Descending As Boolean = False)
                        
    Dim Ulen As Long
    Dim i As Long, j As Long

    Static RecursionLevel As Long
    RecursionLevel = RecursionLevel + 1
    Ulen = rightt - leftt
    If Ulen < 27 Then 'insertion sort for tiny array
        i = leftt + 1
        Do While i <= rightt
            j = i
            Do While j > leftt
                If P_CSV_DATA(j)(SortColumn) < P_CSV_DATA(j - 1)(SortColumn) Then
                    P_CSV_DATA.Swap j, j - 1
                End If
                j = j - 1
            Loop
            i = i + 1
        Loop
    Else
        Dim Third As Long
        Third = Ulen / div
        '"medians"
        Dim m1 As Long
        Dim m2 As Long
        m1 = leftt + Third
        m2 = rightt - Third
        If m1 <= leftt Then m1 = leftt + 1
        If m2 >= rightt Then m2 = rightt - 1
        If P_CSV_DATA(m1)(SortColumn) < P_CSV_DATA(m2)(SortColumn) Then
            P_CSV_DATA.Swap m1, leftt
            P_CSV_DATA.Swap m2, rightt
        Else
            P_CSV_DATA.Swap m1, rightt
            P_CSV_DATA.Swap m2, leftt
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        pivot1 = P_CSV_DATA(leftt)(SortColumn)
        pivot2 = P_CSV_DATA(rightt)(SortColumn)
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        Do While k <= great
            If P_CSV_DATA(k)(SortColumn) < pivot1 Then
                P_CSV_DATA.Swap k, less
                less = less + 1
            ElseIf P_CSV_DATA(k)(SortColumn) > pivot2 Then
                Do While k < great And P_CSV_DATA(great)(SortColumn) > pivot2
                    great = great - 1
                Loop
                P_CSV_DATA.Swap k, great
                great = great - 1
                If P_CSV_DATA(k)(SortColumn) < pivot1 Then
                    P_CSV_DATA.Swap k, less
                    less = less + 1
                End If
            End If
            k = k + 1
        Loop
        'swaps
        Dim dist As Long
        dist = great - less
        If dist < 13 Then div = div + 1
        P_CSV_DATA.Swap less - 1, leftt
        P_CSV_DATA.Swap great + 1, rightt
        'subarrays
        Quicksort leftt, less - 2, div, SortColumn, Descending
        Quicksort great + 2, rightt, div, SortColumn, Descending
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            k = less
            Do While k <= great
                If P_CSV_DATA(k)(SortColumn) = pivot1 Then
                    P_CSV_DATA.Swap k, less
                    less = less + 1
                ElseIf P_CSV_DATA(k)(SortColumn) = pivot2 Then
                    P_CSV_DATA.Swap k, great
                    great = great - 1
                    If P_CSV_DATA(k)(SortColumn) = pivot1 Then
                        P_CSV_DATA.Swap k, less
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'subarray
        If pivot1 < pivot2 Then
            Quicksort less, great, div, SortColumn, Descending
        End If
    End If
    If Descending = True Then
        If RecursionLevel = 1 Then
            ReverseOrder leftt, rightt
        End If
    End If
    RecursionLevel = RecursionLevel - 1
End Sub

Private Function RequestedField(ByVal FieldIndex As Long, _
                                ByVal mxReq As Long) As Boolean
    Select Case RequestedFieldsArray(0)
        Case -1
            RequestedField = True
        Case Else
            Dim idx As Long
            Dim tmpBool As Boolean
            
            idx = 0
            Do
                tmpBool = (FieldIndex = RequestedFieldsArray(idx))
                idx = idx + 1
            Loop While Not tmpBool And idx <= mxReq
            RequestedField = tmpBool
    End Select
End Function

Public Sub ResetToDefault()
    Set config = New parserConfig
    P_ERROR_DESC = vbNullString
    P_ERROR_NUMBER = 0&
    P_ERROR_SOURCE = vbNullString
    P_SUCCESSFUL_EXPORT = False
    P_SUCCESSFUL_IMPORT = False
    P_VARYING_LENGTHS = False
End Sub

Private Function ReverseOrder(Optional LB As Long = -1, _
                                Optional UB As Long = -1) As Boolean
                                
    Dim i As Long, tmpVar As Variant
    
    On Error Resume Next
    If LB < 0 Then
        LB = 0
    End If
    If UB < 0 Then
        UB = P_CSV_DATA.count - 1
    End If
    If UB > LB Then
        For i = LB To (LB + ((UB - LB - 1) \ 2))
            P_CSV_DATA.Swap i, UB - (i - LB)
        Next i
    End If
    If Err.Number = 0 Then
        ReverseOrder = True
    Else
        ReverseOrder = False
    End If
End Function

Private Sub SkipUnwantedLines(ByRef idx As Long, _
                                ByRef MaxIdx As Long, _
                                ByRef Arr() As String, _
                                ByVal CommentToken As Long, _
                                ByVal FDAscW As Long)
    Dim CurLength As Long
    Dim CharCode As Long
    Dim StrPointer As Long
    Dim UnwantedLine As Boolean

    '@----------------------------------------------------------------------------
    'Skip commented and blank lines if needed
    If idx <= MaxIdx Then
        Do
            CurLength = LenB(Arr(idx))
            UnwantedLine = False
            StrPointer = 1&
            If CurLength = 0 Then 'Empty line found
                UnwantedLine = True
            Else
                CharCode = AscW(Arr(idx))
                If CharCode = CommentToken Then 'Commented line found
                    UnwantedLine = True
                Else
                    '@----------------------------------------------
                    'Ignore non printable chars at the string start
                    Select Case FDAscW
                        Case Is <> 9
                            Do While CharCode >= 0& And CharCode <= 32& And StrPointer < CurLength
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    Exit Do
                                End If
                                CharCode = AscW(MidB$(Arr(idx), StrPointer, 2))
                            Loop
                            If CharCode = CommentToken Then UnwantedLine = True
                        Case Else
                            Do While (CharCode >= 0& And CharCode <= 32&) And (StrPointer < CurLength And CharCode <> 9)
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    Exit Do
                                End If
                                CharCode = AscW(MidB$(Arr(idx), StrPointer, 2))
                            Loop
                            If CharCode = CommentToken Then UnwantedLine = True
                    End Select
                End If
            End If
            If UnwantedLine Then
                idx = idx + 1&
                CurLength = LenB(Arr(idx))
            End If
        Loop While UnwantedLine And idx <= MaxIdx
    End If
End Sub

Public Function Sort(Optional ByVal fromIndex As Long = -1, _
                        Optional ByVal toIndex As Long = -1, _
                        Optional ByVal SortColumn As Long = -1, _
                        Optional ByVal Descending As Boolean = False) As CSVinterface
Attribute Sort.VB_Description = "Sorts the imported CSV data."
    
    On Error GoTo Sort_Error_Handler:
    Dim ParametersWellFormed As Boolean
    
    If P_SUCCESSFUL_IMPORT Then
        ParametersWellFormed = True
        If fromIndex = -1 Then
            If config.headers And Not config.headersOmission Then
                fromIndex = 1
            Else
                fromIndex = 0
            End If
        Else
            If fromIndex <= P_CSV_DATA.count And fromIndex > 0 Then
                fromIndex = fromIndex - 1 'Zero based
            Else
                ParametersWellFormed = False
            End If
        End If
        If toIndex = -1 Then
            toIndex = P_CSV_DATA.count - 1
        Else
            If toIndex <= P_CSV_DATA.count And toIndex > 0 Then
                toIndex = toIndex - 1
            Else
                ParametersWellFormed = False
            End If
        End If
        If fromIndex > toIndex Then
            ParametersWellFormed = False
        End If
        If SortColumn = -1 Then
                SortColumn = 0 'Ensures a column to sort
        ElseIf SortColumn < LBound(P_CSV_DATA(0)) + 1 Or _
                SortColumn > UBound(P_CSV_DATA(0)) + 1 Then
                ParametersWellFormed = False
        Else
            SortColumn = SortColumn - 1
        End If
        If ParametersWellFormed Then
            Quicksort fromIndex, toIndex, 3, SortColumn, Descending
            Set Sort = Me
        End If
    End If
    Exit Function
Sort_Error_Handler:
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = "[Sort] " & Err.Description
    P_ERROR_SOURCE = Err.Source
End Function

Private Sub StreamParseCSV(configObj As parserConfig, _
                            outputList As ECPArrayList, _
                            ByRef FilterColumns() As Variant)
    Dim ASCIIcharw As Long
    Dim commToken As Long
    Dim dynamicType As Boolean
    Dim EndOF As Boolean
    Dim EndOfStreamBuffer As Boolean
    Dim EndRecord As Long
    Dim EscapeAscW As Long
    Dim EscapeChr As String
    Dim EscapedWFS As Boolean
    Dim FDAscW As Long
    Dim FieldID As Long
    Dim FieldDelimiter As String
    Dim IgnoreLines As Boolean
    Dim ImportDepth As Long
    Dim ImportSwitch As Boolean
    Dim IncludeRecord As Boolean
    Dim Index As Long, MaxIndex As Long
    Dim IsHeader As Boolean
    Dim KnownEnding As Boolean
    Dim LenCurrentIndex As Long
    Dim MaxReqIndex As Long
    Dim OpenedToken As Boolean
    Dim OverStringPointer As Long, hpointer As Long
    Dim RecordEndReached As Boolean
    Dim RecordsCount As Long
    Dim RecordDelimiter As String
    Dim RecordToken As ECPArrayList
    Dim SearchBeginningMark As Long
    Dim SplittedToken As Boolean
    Dim StartRecord As Long
    Dim StreamEnd As Boolean
    Dim tmpCSV() As String
    Dim tmpRequested() As Variant
    Dim tmpToken As ECPArrayList
    Dim TokenBeginningPos As Long
    Dim TokenEndingPos As Long
    Dim TokenEndReached As Boolean
    
    On Error GoTo ParseCriticalCSV_Error_Handler
    '@----------------------------------------------------------------------------
    'Start variables
    SearchBeginningMark = 1&
    RecordsCount = 0&
    P_VECTORS_REGULAR_BOUND = 0&
    P_VECTORS_MAX_BOUND = 0&
    Set tmpToken = New ECPArrayList
    Set CSVstream = New ECPTextStream
    With configObj
        CSVstream.bufferSize = .bufferSize
        CSVstream.endStreamOnLineBreak = True
        CSVstream.OpenStream .path
        commToken = AscW(.commentsToken)
        dynamicType = .dynamicTyping
        dTTemplate() = .dTypingTemplate
        dTTargets() = .dTypingLinks
        IgnoreLines = .catchMeaninglessRecords
        TemplateDefined = .dTTemplateDefined
        TargetsDefined = .dTTemplateLinksDefined
        FieldDelimiter = .fieldsDelimiter
        RecordDelimiter = .recordsDelimiter
        EndRecord = .endingRecord
        StartRecord = .startingRecord
    End With
    TokenEndReached = False
    RecordEndReached = False
    tmpRequested() = FilterColumns
    Select Case UBound(tmpRequested)
        Case -1
            ReDim RequestedFieldsArray(0)
            RequestedFieldsArray(0) = -1
        Case Else
            RequestedFieldsArray() = GetLongFilters(tmpRequested, P_CSV_HEADER)
    End Select
    '@----------------------------------------------------------------------------
    'Set variables
    OverStringPointer = SearchBeginningMark
    MaxReqIndex = UBound(RequestedFieldsArray)
    '@----------------------------------------------------------------------------
    'Check if user want Import all records
    KnownEnding = True
    Select Case EndRecord
        Case 0&
            If StartRecord >= 1 Then
                KnownEnding = False
            Else
                ImportDepth = 1&
            End If
        Case Is < StartRecord
            ImportDepth = 1&
        Case Else
            ImportDepth = EndRecord - StartRecord + 1&
    End Select
    If StartRecord = 1& Then
        If configObj.headersOmission And configObj.headers Then
            If ImportDepth > 1& Or EndRecord = 0& Then
                StartRecord = 2&
                ImportDepth = ImportDepth - 1&
            End If
        End If
    End If
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case configObj.escapeToken
        Case 1
            EscapeChr = CHR_APOSTROPHE
        Case Else
            EscapeChr = CHR_DOUBLE_QUOTES
    End Select
    EscapeAscW = AscW(EscapeChr)
    FDAscW = AscW(FieldDelimiter)
    Set RecordToken = New ECPArrayList
    Index = 0
    '@----------------------------------------------------------------------------
    'Get CSV stream and populate a temp array
    CSVstream.ReadText
    StreamWithQuotes = InStrB(1, CSVstream.bufferString, CHR_DOUBLE_QUOTES)
    StreamEnd = CSVstream.atEndOfStream
    tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
    '@----------------------------------------------------------------------------
    'Ensuring the stream is sufficient enough to hold at least one CSV line
    Do While Not StreamEnd And UBound(tmpCSV) = 0
        CSVstream.bufferSize = 2 * configObj.bufferSize
        CSVstream.RestartPointer
        CSVstream.ReadText
        StreamWithQuotes = InStrB(1, CSVstream.bufferString, CHR_DOUBLE_QUOTES)
        StreamEnd = CSVstream.atEndOfStream
        tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
    Loop
    MaxIndex = UBound(tmpCSV)
    If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
        MaxIndex = MaxIndex - 1
    Else
        If tmpCSV(MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
            MaxIndex = MaxIndex - 1
        End If
    End If
    EndOfStreamBuffer = (Index > MaxIndex)
    '@----------------------------------------------------------------------------
    'Loop through CSV using streams
    Do
        IsHeader = (StartRecord = 1 And RecordsCount = 0)
        If IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines Index, CSVstream, tmpCSV, MaxIndex, RecordDelimiter, commToken, FDAscW
        End If
        LenCurrentIndex = LenB(tmpCSV(Index))
        '@----------------------------------------------------------------------------
        'Restart pointers
        RecordEndReached = False
        FieldID = 0
        IncludeRecord = (RecordsCount + 1& >= StartRecord)
        Do
            '@----------------------------------------------------------------------------
            'Process the current Token
            TokenEndReached = False
            TokenBeginningPos = OverStringPointer
            If StreamWithQuotes Then
                Select Case OverStringPointer
                    Case Is < LenCurrentIndex
                        ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces, tabs and remanents LF
                        Select Case FDAscW
                            Case Is <> 9
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 9 Or ASCIIcharw = 10
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 10
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                        End Select
                        OpenedToken = (ASCIIcharw = EscapeAscW)
                    Case Else 'Avoid over runs
                        OpenedToken = False
                End Select
            Else
                OpenedToken = False
            End If
            If OpenedToken Then
                '@----------------------------------------------------------------------------
                'Try to escape the Token
                TokenBeginningPos = OverStringPointer + 2&
                SearchBeginningMark = TokenBeginningPos
                hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'EOFld marker
                Do While (Not TokenEndReached And Not EndOF)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not SplittedToken Then SplittedToken = (hpointer = 0)
                    Do While hpointer = 0& 'Escape char may be on subsequent Index
                        If IncludeRecord Then
                            Select Case SearchBeginningMark
                                Case 1 'Store the full Token from the current Index
                                    tmpToken.Add tmpCSV(Index)
                                Case Else 'Store the Right most Token portion from the current Index
                                    tmpToken.Add RightB$(tmpCSV(Index), LenCurrentIndex - SearchBeginningMark + 1)
                            End Select
                        End If
                        Index = Index + 1&
                        EndOfStreamBuffer = (Index > MaxIndex)
                        If Not EndOfStreamBuffer Then
                            LenCurrentIndex = LenB(tmpCSV(Index))
                            SearchBeginningMark = 1
                            hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                        Else
                            If Not StreamEnd Then
                                '@----------------------------------------------------------------------------
                                'Load new stream
                                CSVstream.ReadText
                                StreamWithQuotes = InStrB(1, CSVstream.bufferString, CHR_DOUBLE_QUOTES)
                                StreamEnd = CSVstream.atEndOfStream
                                tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
                                Index = 0
                                MaxIndex = UBound(tmpCSV)
                                If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                    MaxIndex = MaxIndex - 1
                                Else
                                    If tmpCSV(MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                                        MaxIndex = MaxIndex - 1
                                    End If
                                End If
                                EndOfStreamBuffer = False
                                LenCurrentIndex = LenB(tmpCSV(Index))
                                SearchBeginningMark = 1
                                hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                            Else
                                hpointer = 0&
                            End If
                        End If
                        EndOF = (StreamEnd And EndOfStreamBuffer)
                    Loop
                    If hpointer = 0& And EndOF Then GoTo ParseCriticalCSV_MissingEscapeChar 'Missing escape char
                    OverStringPointer = hpointer + 2&
                    If OverStringPointer < LenCurrentIndex Then
                        ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces, tabs and remanents LF
                        Select Case FDAscW
                            Case Is <> 9
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 9 Or ASCIIcharw = 10
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 10
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                        End Select
                        '@----------------------------------------------------------------------------
                        'Check if we are at the end of the field
                        EscapedWFS = (ASCIIcharw = FDAscW)
                    Else
                        EscapedWFS = False
                    End If
                    If EscapedWFS Then 'Escape char succeeded by fields separator char
                        TokenEndingPos = hpointer - 1&
                        TokenEndReached = True
                        If IncludeRecord Then
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case SplittedToken
                                Case False 'Just save Token
                                    If IncludeRecord Then
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            Select Case dynamicType
                                                Case False
                                                    RecordToken.Add MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                            TokenEndingPos - TokenBeginningPos + 1)
                                                Case Else
                                                    RecordToken.Add TypeData(MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                            TokenEndingPos - TokenBeginningPos + 1), _
                                                                            FieldID, _
                                                                            IsHeader)
                                            End Select
                                        End If
                                    End If
                                Case Else 'The Token is stored on tmpToken object
                                    If IncludeRecord Then
                                        tmpToken.Add LeftB$(tmpCSV(Index), TokenEndingPos)
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            Select Case dynamicType
                                                Case False
                                                    RecordToken.Add Join$(tmpToken.items, RecordDelimiter) 'Save whole Token
                                                Case Else
                                                    RecordToken.Add TypeData(Join$(tmpToken.items, _
                                                                                RecordDelimiter), _
                                                                                FieldID, _
                                                                                IsHeader)
                                            End Select
                                        End If
                                        tmpToken.Clear 'Clear tmpToken content
                                    End If
                            End Select
                        End If
                        SplittedToken = False 'Restart Token status
                        OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                        RecordEndReached = False
                    Else
                        If OverStringPointer >= LenCurrentIndex Then  'Escape char succeeded
                                                                      'by records separator char
                            TokenEndingPos = hpointer - 1&
                            TokenEndReached = True
                            RecordEndReached = True 'EOR
                            If IncludeRecord Then
                                '@----------------------------------------------------------------------------
                                'Save the Token
                                Select Case SplittedToken
                                    Case False 'Just save Token
                                        If IncludeRecord Then
                                            If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                                Select Case dynamicType
                                                    Case False
                                                        RecordToken.Add MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                            TokenEndingPos - TokenBeginningPos + 1)
                                                    Case Else
                                                        RecordToken.Add TypeData(MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                                    TokenEndingPos - TokenBeginningPos + 1), _
                                                                                FieldID, _
                                                                                IsHeader)
                                                End Select
                                            End If
                                        End If
                                    Case Else 'The Token is stored on tmpToken array
                                        If IncludeRecord Then
                                            tmpToken.Add LeftB$(tmpCSV(Index), OverStringPointer - 3)
                                            If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                                Select Case dynamicType
                                                    Case False
                                                        RecordToken.Add Join$(tmpToken.items, RecordDelimiter) 'Save whole Token
                                                    Case Else
                                                        RecordToken.Add TypeData(Join$(tmpToken.items, _
                                                                                RecordDelimiter), _
                                                                                FieldID, _
                                                                                IsHeader)
                                                End Select
                                            End If
                                            tmpToken.Clear 'Clear tmpToken content
                                        End If
                                End Select
                            End If
                            SplittedToken = False 'Restart Token status
                            Index = Index + 1&
                            EndOfStreamBuffer = (Index > MaxIndex)
                            EndOF = (StreamEnd And EndOfStreamBuffer)
                            If EndOfStreamBuffer And Not EndOF Then
                                '@----------------------------------------------
                                'Advance stream
                                CSVstream.ReadText
                                StreamWithQuotes = InStrB(1, CSVstream.bufferString, CHR_DOUBLE_QUOTES)
                                StreamEnd = CSVstream.atEndOfStream
                                tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
                                Index = 0
                                MaxIndex = UBound(tmpCSV)
                                If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                    MaxIndex = MaxIndex - 1
                                Else
                                    If tmpCSV(MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                                        MaxIndex = MaxIndex - 1
                                    End If
                                End If
                                EndOfStreamBuffer = False
                            End If
                            OverStringPointer = 1&
                        Else
                            If IncludeRecord Then
                                If SplittedToken Then 'Store the Left most Token portion from the current Index
                                    tmpToken.Add LeftB$(tmpCSV(Index), OverStringPointer - 1)
                                End If
                            End If
                            SearchBeginningMark = OverStringPointer
                            hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                            TokenEndReached = False
                        End If
                    End If
                Loop
            Else
                SplittedToken = False
                '@----------------------------------------------------------------------------
                'Try to mark the Starting and Ending position for the field
                hpointer = InStrB(OverStringPointer, tmpCSV(Index), FieldDelimiter)
                If hpointer Then 'Found fields delimiter char
                    TokenEndReached = True
                    TokenEndingPos = hpointer - 1&
                    OverStringPointer = hpointer + 2&
                    RecordEndReached = False
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                            Select Case dynamicType
                                Case False
                                    RecordToken.Add MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                    TokenEndingPos - TokenBeginningPos + 1)
                                Case Else
                                    RecordToken.Add TypeData(MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                            TokenEndingPos - TokenBeginningPos + 1), _
                                                            FieldID, _
                                                            IsHeader)
                            End Select
                        End If
                    End If
                Else 'Missing fields delimiter char
                    TokenEndReached = True
                    RecordEndReached = True
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                            Select Case dynamicType
                                Case False
                                    RecordToken.Add RightB$(tmpCSV(Index), LenCurrentIndex _
                                                    - TokenBeginningPos + 1)
                                Case Else
                                    RecordToken.Add TypeData(RightB$(tmpCSV(Index), LenCurrentIndex _
                                                            - TokenBeginningPos + 1), _
                                                                FieldID, _
                                                                IsHeader)
                            End Select
                        End If
                    End If
                    Index = Index + 1&
                    EndOfStreamBuffer = (Index > MaxIndex)
                    EndOF = (StreamEnd And EndOfStreamBuffer)
                    If EndOfStreamBuffer And Not EndOF Then
                        '@----------------------------------------------
                        'Advance stream
                        CSVstream.ReadText
                        StreamWithQuotes = InStrB(1, CSVstream.bufferString, CHR_DOUBLE_QUOTES)
                        StreamEnd = CSVstream.atEndOfStream
                        tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
                        Index = 0
                        MaxIndex = UBound(tmpCSV)
                        If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                            MaxIndex = MaxIndex - 1
                        Else
                            If tmpCSV(MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                                MaxIndex = MaxIndex - 1
                            End If
                        End If
                        EndOfStreamBuffer = False
                    End If
                    OverStringPointer = 1&
                End If
            End If
            FieldID = FieldID + 1&
        Loop While (Not RecordEndReached And Not EndOF)
        Select Case P_VECTORS_REGULAR_BOUND
            Case Is > 0
                If RecordToken.count - 1 > P_VECTORS_MAX_BOUND Then
                    P_VECTORS_MAX_BOUND = RecordToken.count - 1
                    RecordToken.ShrinkBuffer
                End If
            Case Else
                Select Case RecordToken.count
                    Case Is > 0
                        P_VECTORS_REGULAR_BOUND = RecordToken.count - 1
                        RecordToken.ShrinkBuffer
                    Case Else
                        P_VECTORS_REGULAR_BOUND = RecordToken.count
                End Select
                P_VECTORS_MAX_BOUND = P_VECTORS_REGULAR_BOUND
        End Select
        RecordsCount = RecordsCount + 1&
        '@----------------------------------------------------------------------------
        'Start saving on given record
        If IncludeRecord Then
            outputList.Add RecordToken.items
            RecordToken.Reinitialize P_VECTORS_REGULAR_BOUND + 1
        End If
        P_VARYING_LENGTHS = (P_VECTORS_REGULAR_BOUND <> P_VECTORS_MAX_BOUND)
        If IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines Index, CSVstream, tmpCSV, MaxIndex, RecordDelimiter, commToken, FDAscW
        End If
        EndOfStreamBuffer = (Index > MaxIndex)
        EndOF = (StreamEnd And EndOfStreamBuffer)
        If EndOfStreamBuffer And Not EndOF Then
            '@----------------------------------------------
            'Advance stream
            CSVstream.ReadText
            StreamWithQuotes = InStrB(1, CSVstream.bufferString, CHR_DOUBLE_QUOTES)
            StreamEnd = CSVstream.atEndOfStream
            tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
            Index = 0
            MaxIndex = UBound(tmpCSV)
            If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                MaxIndex = MaxIndex - 1
            Else
                If tmpCSV(MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                    MaxIndex = MaxIndex - 1
                End If
            End If
            EndOfStreamBuffer = False
        End If
        If Not EndOfStreamBuffer Then LenCurrentIndex = LenB(tmpCSV(Index))
        Select Case KnownEnding
            Case False
                ImportSwitch = EndOF
            Case Else
                ImportSwitch = (outputList.count = ImportDepth)
        End Select
    Loop While Not ImportSwitch
    '@----------------------------------------------------------------------------
    'Check if the array need to be shrinked at the output stage
    outputList.ShrinkBuffer
    Erase tmpCSV
    Set CSVstream = Nothing
    Set RecordToken = Nothing
    Set tmpToken = Nothing
    Exit Sub
ParseCriticalCSV_Error_Handler:
    Set outputList = Nothing
    Set tmpToken = Nothing
    Set CSVstream = Nothing
    Erase tmpCSV
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description & " [Check the record #" & RecordsCount + 1 _
                                  & ", field #" & RecordToken.count + 1 _
                                   & " on the source CSV file]."
    Set RecordToken = Nothing
    P_ERROR_SOURCE = Err.Source
    Exit Sub
ParseCriticalCSV_StartingIndexOutOfBound:
    Err.Raise Number:=vbObjectError + 9002, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="The index over which the operation starts is out of bounds. Try start the import on a lower index."
    Resume ParseCriticalCSV_Error_Handler
ParseCriticalCSV_MissingEscapeChar:
    Err.Raise Number:=vbObjectError + 9003, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="Missing some escape char. Check the data and try again."
    Resume ParseCriticalCSV_Error_Handler
End Sub

Private Sub StreamSkipUnwantedLines(ByRef idx As Long, _
                                    ByRef Ostream As ECPTextStream, _
                                    ByRef Arr() As String, _
                                    ByRef MaxIdx As Long, _
                                    ByRef recDelimiter As String, _
                                    ByVal CommentToken As Long, _
                                    ByVal FDAscW As Long)
    Dim CurLength As Long
    Dim CharCode As Long
    Dim StrPointer As Long
    Dim UnwantedLine As Boolean

start:
    '@----------------------------------------------------------------------------
    'Skip commented and blank lines if needed
    If idx <= MaxIdx Then
        Do
            CurLength = LenB(Arr(idx))
            UnwantedLine = False
            StrPointer = 1&
            If CurLength = 0 Then 'Empty line found
                UnwantedLine = True
            Else
                CharCode = AscW(Arr(idx))
                If CharCode = CommentToken Then 'Commented line found
                    UnwantedLine = True
                Else
                    '@----------------------------------------------
                    'Ignore non printable chars at the string start
                    Select Case FDAscW
                        Case Is <> 9
                            Do While CharCode >= 0& And CharCode <= 32& And StrPointer < CurLength
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    GoTo AdvanceLine
                                End If
                                CharCode = AscW(MidB$(Arr(idx), StrPointer, 2))
                            Loop
                            UnwantedLine = (CharCode = CommentToken)
                        Case Else
                            Do While (CharCode >= 0& And CharCode <= 32&) And (StrPointer < CurLength And CharCode <> 9)
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    GoTo AdvanceLine
                                End If
                                CharCode = AscW(MidB$(Arr(idx), StrPointer, 2))
                            Loop
                            UnwantedLine = (CharCode = CommentToken)
                    End Select
                End If
            End If
AdvanceLine:
            If UnwantedLine Then
                idx = idx + 1&
                If idx > MaxIdx Then
                    If Not Ostream.atEndOfStream Then
                        '@----------------------------------------------
                        'Advance stream
                        Ostream.ReadText
                        StreamWithQuotes = InStrB(1, Ostream.bufferString, CHR_DOUBLE_QUOTES)
                        Arr() = Split(Ostream.bufferString, recDelimiter)
                        idx = 0
                        MaxIdx = UBound(Arr)
                        If Arr(MaxIdx) = vbNullString Then
                            MaxIdx = MaxIdx - 1
                        End If
                    End If
                End If
            End If
        Loop While UnwantedLine And idx <= MaxIdx
    Else
        If Not Ostream.atEndOfStream Then
            '@----------------------------------------------
            'Advance stream
            Ostream.ReadText
            StreamWithQuotes = InStrB(1, Ostream.bufferString, CHR_DOUBLE_QUOTES)
            Arr() = Split(Ostream.bufferString, recDelimiter)
            idx = 0
            MaxIdx = UBound(Arr)
            If Arr(MaxIdx) = vbNullString Then
                MaxIdx = MaxIdx - 1
            End If
            GoTo start
        End If
    End If
End Sub

Public Sub TwoDimToJaggedArray(ByRef TwoDimArray() As Variant, ByRef JaggedArray() As Variant)
    Dim UBj1 As Long, LBj1 As Long
    Dim UBj2 As Long, LBj2 As Long
    Dim MaxDim1 As Long, MaxDim2 As Long
    Dim jgdCounter1 As Long, jgdCounter2 As Long
    Dim tdimCounter As Long, tdimCounter2 As Long
    
    On Error GoTo TwoDimToJaggedArray_Err_Handler
    If MultiDimensional(TwoDimArray) Then
        LBj1 = LBound(TwoDimArray)
        UBj1 = UBound(TwoDimArray)
        MaxDim1 = Abs(UBj1 - LBj1) 'Dimension One in base 0
        LBj2 = LBound(TwoDimArray, 2)
        UBj2 = UBound(TwoDimArray, 2)
        MaxDim2 = Abs(UBj2 - LBj2) 'Dimension Two in base 0
        '@----------------------------------------------
        'Create the jagged array
        Call CreateJagged(JaggedArray, MaxDim1, MaxDim2)
        '@----------------------------------------------
        'Deconstruct and dump the data
        jgdCounter1 = 0
        For tdimCounter = LBj1 To UBj1
            jgdCounter2 = 0
            For tdimCounter2 = LBj2 To UBj2
                JaggedArray(jgdCounter1)(jgdCounter2) = TwoDimArray(tdimCounter, tdimCounter2)
                TwoDimArray(tdimCounter, tdimCounter2) = vbNullString
                jgdCounter2 = jgdCounter2 + 1&
            Next tdimCounter2
            jgdCounter1 = jgdCounter1 + 1&
        Next tdimCounter
    Else
        LBj1 = LBound(TwoDimArray)
        UBj1 = UBound(TwoDimArray)
        MaxDim1 = Abs(UBj1 - LBj1) 'Dimension One in base 0
        MaxDim2 = 0 'Dimension Two in base 0
        '@----------------------------------------------
        'Create the jagged array
        Call CreateJagged(JaggedArray, MaxDim1, MaxDim2)
        '@----------------------------------------------
        'Deconstruct and dump the data
        jgdCounter1 = 0
        For tdimCounter = LBj1 To UBj1
            JaggedArray(jgdCounter1)(0) = TwoDimArray(tdimCounter)
            TwoDimArray(tdimCounter) = vbNullString
            jgdCounter1 = jgdCounter1 + 1&
        Next tdimCounter
    End If
    Erase TwoDimArray
TwoDimToJaggedArray_Err_Handler:
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
    P_ERROR_SOURCE = Err.Source
End Sub
Private Function TypeData(ByRef data As Variant, _
                            ByVal fieldIdx As Long, _
                            ByVal IsHeader As Boolean) As Variant
    If TemplateDefined Then
        If TargetsDefined Then
            If Not IsHeader Then
                Dim typingIdx As Long, maxTypingIdx As Long
                Dim typingNeed As Boolean
            
                typingIdx = LBound(dTTargets)
                maxTypingIdx = UBound(dTTargets)
                Do
                    typingNeed = (fieldIdx + 1 = dTTargets(typingIdx)) 'fieldIdx is base 0
                    typingIdx = typingIdx + 1
                Loop While typingIdx <= maxTypingIdx And Not typingNeed
                If typingNeed Then
                    Select Case dTTemplate(typingIdx - 1)
                        Case TypeConversion.None 'Strings
                            TypeData = data
                        Case TypeConversion.ToLong
                            TypeData = CLng(data)
                        Case TypeConversion.ToDouble
                            TypeData = CDbl(data)
                        Case TypeConversion.ToDate
                            TypeData = CDate(data)
                        Case TypeConversion.ToBoolean
                            TypeData = CBool(data)
                        Case Else
                            TypeData = data
                    End Select
                Else
                    TypeData = data
                End If
            Else 'Not type headers
                TypeData = data
            End If
        End If
    End If
End Function
