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
'@------------------------------------------------------
' SEQUENTIAL READ VARIABLES
Private static_ASCIIcharw As Long
Private static_commToken As Long
Private static_CSVstream As ECPTextStream
Private static_dTTargets() As Variant
Private static_dTTemplate() As Variant
Private static_dynamicType As Boolean
Private static_EndOF As Boolean
Private static_EndOFStreamBuffer As Boolean
Private static_EscapeAscW As Long
Private static_EscapeChr As String
Private static_EscapedWFS As Boolean
Private static_FDAscW As Long
Private static_FieldDelimiter As String
Private static_FieldID As Long
Private static_Headers As Boolean
Private static_IgnoreLines As Boolean
Private static_ImportDepth As Long
Private static_ImportSwitch As Boolean
Private static_Index As Long, static_MaxIndex As Long
Private static_IsHeader As Boolean
Private static_KnownEnding As Boolean
Private static_LenCurrentIndex As Long
Private static_MaxReqIndex As Long
Private static_OpenedToken As Boolean
Private static_outputList As ECPArrayList
Private static_OverStringPointer As Long, static_hpointer As Long
Private static_RecordDelimiter As String
Private static_RecordEndReached As Boolean
Private static_RecordsCount As Long
Private static_RecordToken As ECPArrayList
Private static_RequestedFieldsArray() As Long
Private static_SearchBeginningMark As Long
Private static_SeqReadStarted As Boolean
Private static_SplittedToken As Boolean
Private static_StreamEnd As Boolean
Private static_StreamWithQuotes As Boolean
Private static_TargetsDefined As Boolean
Private static_TemplateDefined As Boolean
Private static_tmpCSV() As String
Private static_tmpRequested() As Variant
Private static_tmpToken As ECPArrayList
Private static_TokenBeginningPos As Long
Private static_TokenEndingPos As Long
Private static_TokenEndReached As Boolean
Private static_vectorsBound As Long
Private static_vectorsMaxBound As Long
'@------------------------------------------------------
' CSV SUBSET VARIABLES
Private queryFilters As Variant
Private UB As Long
Private LB As Long
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
Public Property Get items() As ECPArrayList
Attribute items.VB_Description = "Gets or sets the result array over the current instance."
    Set items = P_CSV_DATA
End Property
Public Property Set items(values As ECPArrayList)
    Set P_CSV_DATA = values
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

Public Sub CloseSeqReader()
    static_SeqReadStarted = False
    Erase static_tmpCSV
    Set static_CSVstream = Nothing
    Set static_RecordToken = Nothing
    Set static_tmpToken = Nothing
    Set static_outputList = Nothing
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

Public Function CSVsubsetSplit(filePath As String, _
                                Optional subsetColumn As Long = 1, _
                                Optional headers As Boolean = True) As Collection
    Dim CreatedFiles As Collection
    Dim CSVhead As Variant
    Dim CSVreader As CSVinterface
    Dim CSVstream As ECPTextStream
    Dim CSVwriter As CSVinterface
    Dim curItems As ECPArrayList
    Dim ExportSubSet As ECPArrayList
    Dim fExtension As String
    Dim fileCreatedFlag As Boolean
    Dim FileNameAndExtension() As String
    Dim fileOnPathFlag As Boolean
    Dim fName As String
    Dim fRootPath As String
    Dim i As Long
    Dim keyIndex As Long
    Dim outputPath As String
    Dim readerConf As parserConfig
    Dim tmpVar As Variant
    Dim usedFields As Long
    Dim writterConf As parserConfig
    
    Set CreatedFiles = New Collection
    Set CSVreader = New CSVinterface
    Set CSVwriter = New CSVinterface
    Set readerConf = CSVreader.parseConfig
    Set CSVstream = New ECPTextStream
    Set ExportSubSet = New ECPArrayList
    
    If filePath <> vbNullString Then
        On Error Resume Next
        If subsetColumn >= 1 Then
            keyIndex = subsetColumn - 1
        Else
            keyIndex = 0
        End If
        fRootPath = GetFilePath(filePath)
        FileNameAndExtension() = GetFileNameAndExtension(filePath)
        fName = FileNameAndExtension(0)
        fExtension = FileNameAndExtension(1)
        fileCreatedFlag = False
        DoEvents
        '@----------------------------------------------------
        'Sequential reading from file
        With CSVstream
            .endStreamOnLineBreak = True
            .OpenStream filePath
        End With
        '@----------------------------------------------------
        'Parser config
        readerConf.path = filePath
        CSVreader.GuessDelimiters readerConf
        Set writterConf = readerConf.CopyConfig
        '@----------------------------------------------------
        'Parse string
        With readerConf
            .endingRecord = 1
            CSVhead = CSVreader.ImportFromCSV(readerConf)(0)
            .endingRecord = 0
            .headers = headers
        End With
        usedFields = 0
        outputPath = fRootPath & fName & "-subsets\"
        '@----------------------------------------------------
        'Check directory
        If LenB(Dir(outputPath, vbDirectory)) = 0 Then
            MkDir outputPath
        End If
        Do While Not CSVstream.atEndOfStream
            CSVstream.ReadText 'Read next CSV data chunk
            Set curItems = CSVreader.ImportFromCSVString(CSVstream.bufferString, readerConf).items
            If Not fileCreatedFlag Then
                If headers Then
                    i = 1
                Else
                    i = 0
                End If
            Else
                i = 0
            End If
            curItems.Sort i + 1, SortColumn:=keyIndex + 1
            tmpVar = curItems(i)(keyIndex)
            writterConf.path = outputPath & CStr(tmpVar) & "." & fExtension
            If headers Then
                '@----------------------------------------------------
                'Check directory
                fileOnPathFlag = FileExists(writterConf.path)
                If Not fileOnPathFlag Then
                    ExportSubSet.Add CSVhead
                End If
            End If
            For i = i To curItems.count - 1
                If tmpVar = curItems(i)(keyIndex) Then
                    ExportSubSet.Add curItems(i) 'Append data
                Else
                    CreatedFiles.Add writterConf.path, CStr(tmpVar)
                    ExportSubSet.ShrinkBuffer
                    CSVwriter.ExportToCSV ExportSubSet, writterConf, False, False
                    fileCreatedFlag = True
                    ExportSubSet.Reinitialize ExportSubSet.count
                    tmpVar = curItems(i)(keyIndex)
                    writterConf.path = outputPath & CStr(tmpVar) & "." & fExtension
                    If headers Then
                        fileOnPathFlag = FileExists(writterConf.path)
                        If Not fileOnPathFlag Then
                            ExportSubSet.Add CSVhead
                        End If
                    End If
                    ExportSubSet.Add curItems(i) 'Append data
                End If
            Next i
            '@----------------------------------------------------
            'Export remnant data
            CreatedFiles.Add writterConf.path, CStr(tmpVar)
            ExportSubSet.ShrinkBuffer
            CSVwriter.ExportToCSV ExportSubSet, writterConf, False, False
            ExportSubSet.Reinitialize ExportSubSet.count
        Loop
        Set CSVsubsetSplit = CreatedFiles
        Set CreatedFiles = Nothing
        Set CSVreader = Nothing
        Set CSVstream = Nothing
        Set CSVwriter = Nothing
        Set curItems = Nothing
        Set ExportSubSet = Nothing
        Set readerConf = Nothing
        Set writterConf = Nothing
        tmpVar = Null
    End If
End Function

Public Sub DumpToArray(OutPutArray() As Variant, _
                        Optional ByRef DataSource As ECPArrayList = Nothing)
    If P_SUCCESSFUL_IMPORT Then
        Dim UB1 As Long, ub2 As Long
        Dim Counter1 As Long, Counter2 As Long
        
        If DataSource Is Nothing Then
            Set DataSource = P_CSV_DATA
        End If
        UB1 = DataSource.count - 1
        ReDim OutPutArray(0 To UB1, 0 To P_VECTORS_MAX_BOUND)
        
        For Counter1 = 0 To UB1
            ub2 = UBound(DataSource(Counter1))
            For Counter2 = 0 To ub2
                OutPutArray(Counter1, Counter2) = DataSource(Counter1)(Counter2)
            Next Counter2
        Next Counter1
    End If
End Sub

Public Sub DumpToJaggedArray(OutPutArray() As Variant, _
                                Optional ByRef DataSource As ECPArrayList = Nothing)
    If P_SUCCESSFUL_IMPORT Then
        Dim UB1 As Long
        Dim Counter1 As Long
        
        If DataSource Is Nothing Then
            Set DataSource = P_CSV_DATA
        End If
        UB1 = DataSource.count - 1
        ReDim OutPutArray(0 To UB1)
        
        For Counter1 = 0 To UB1
            OutPutArray(Counter1) = DataSource(Counter1)
        Next Counter1
    End If
End Sub

Public Sub DumpToSheet(Optional WBookName As String, _
                        Optional SheetName As String, _
                        Optional rngName As String = "A1", _
                        Optional ByRef DataSource As ECPArrayList = Nothing)
    On Error Resume Next
    If P_SUCCESSFUL_IMPORT Then
        Dim WBook As Workbook
        Dim outputSheet As Worksheet
        Dim OutputRange As Range
        Dim colNumber As Long
        
        If DataSource Is Nothing Then
            Set DataSource = P_CSV_DATA
            colNumber = P_VECTORS_MAX_BOUND
        Else
            colNumber = UBound(DataSource(0))
        End If
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
        
        UB1 = DataSource.count - 1
        ReDim tmpOutputArray(0 To UB1, 0 To colNumber)
        
        For Counter1 = 0 To UB1
            ub2 = UBound(DataSource(Counter1))
            For Counter2 = 0 To ub2
                tmpOutputArray(Counter1, Counter2) = DataSource(Counter1)(Counter2)
            Next Counter2
        Next Counter1
        '@------------------------------------------------------
        'Set the target Range
        Set OutputRange = outputSheet.Range(rngName) _
                            .Resize _
                            ( _
                            UB1 + 1, _
                            colNumber + 1 _
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
                        Optional PassControlToOS As Boolean = True, _
                        Optional enableDelimiterGuessing As Boolean = True)
    On Error GoTo ErrHandler_ExportToCSV
    If PassControlToOS Then DoEvents 'Pass the control to the Operative System
    If IsArray(csvArray) Or TypeName(csvArray) = "ECPArrayList" Then
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
            If enableDelimiterGuessing Then
                GuessDelimiters curConfig
            End If
            CSVstream.OpenStream curConfig.path
            If CSVstream.streamLength > 0 Then
                Select Case curConfig.recordsDelimiter
                    Case vbCrLf
                        CSVstream.SeekPointer CSVstream.streamLength - 1
                    Case Else
                        CSVstream.SeekPointer CSVstream.streamLength
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

Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = CBool(LenB(Dir(filePath, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)))
End Function

Public Function GetCSVsubset(filePath As String, _
                        ByRef filters As Variant, _
                        keyIndex As Long, _
                        Optional configObj As parserConfig = Nothing) As ECPArrayList
                        
    Dim streamReader As ECPTextStream
    Dim curItems As ECPArrayList
    Dim filteredRecords As ECPArrayList
    Dim CSVparser As CSVinterface
    Dim gssCounter As Long
    
    Set CSVparser = New CSVinterface
    Set filteredRecords = New ECPArrayList
    Set streamReader = New ECPTextStream
    
    If filePath <> vbNullString Then
        On Error GoTo err_handler
        queryFilters = filters
        UB = UBound(queryFilters)
        If keyIndex >= 1 Then
            keyIndex = keyIndex - 1 'Zero based
        Else
            keyIndex = 0
        End If
        If UB <> -1 Then
            LB = LBound(queryFilters)
            '@----------------------------------------------------
            ' Sequential reading from file
            With streamReader
                .endStreamOnLineBreak = True
                .OpenStream filePath
                .ReadText
            End With
            '@----------------------------------------------------
            ' Parser config
            If configObj Is Nothing Then
                Set configObj = config
            End If
            configObj.path = filePath
            CSVparser.GuessDelimiters configObj
            '@----------------------------------------------------
            ' Parse string
            Set curItems = CSVparser.ImportFromCSVString(streamReader.bufferString, configObj).items
            If configObj.headers Then
                filteredRecords.Add curItems(0) 'Save the CSV header
                gssCounter = gssCounter + 1
            End If
            For gssCounter = gssCounter To curItems.count - 1
                If MeetsCriterion(curItems(gssCounter)(keyIndex)) Then
                    filteredRecords.Add curItems(gssCounter) 'Append data
                End If
            Next gssCounter
            Do While Not streamReader.atEndOfStream
                streamReader.ReadText 'Read next CSV data chunk
                Set curItems = CSVparser.ImportFromCSVString(streamReader.bufferString, configObj).items
                For gssCounter = 0 To curItems.count - 1
                    If MeetsCriterion(curItems(gssCounter)(keyIndex)) Then
                        filteredRecords.Add curItems(gssCounter) 'Append data
                    End If
                Next gssCounter
            Loop
            Set GetCSVsubset = filteredRecords
            Set filteredRecords = Nothing
            Set CSVparser = Nothing
            Set curItems = Nothing
            Set streamReader = Nothing
        End If
        P_SUCCESSFUL_IMPORT = True
    End If
    Exit Function
err_handler:
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

Private Function GetFileNameAndExtension(path As String) As String()
    Dim tmpResult() As String
    tmpResult() = Split(Mid$(path, InStrRev(path, "\") + 1), ".")
    If UBound(tmpResult) > 1 Then
        Dim i As Single
        Dim tmpName As String
        
        tmpName = tmpResult(0)
        For i = 1 To UBound(tmpResult) - 1
            tmpName = tmpName & "." & tmpResult(i)
        Next i
        tmpResult(0) = tmpName
        tmpResult(1) = tmpResult(i)
        ReDim Preserve tmpResult(0 To 1)
    End If
    GetFileNameAndExtension = tmpResult
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

Private Function GetFilePath(path As String)
    GetFilePath = Mid$(path, 1, InStrRev(path, "\"))
End Function

Public Function GetRecord() As ECPArrayList
    
    On Error GoTo GetRecord_Error_Handler
    If Not static_SeqReadStarted Then Exit Function
    '@----------------------------------------------------------------------------
    'Get CSV record using streams
    If Not static_ImportSwitch Then
        '@----------------------------------------------------------------------------
        'Start variables
        dTTargets() = static_dTTargets
        dTTemplate() = static_dTTemplate
        RequestedFieldsArray() = static_RequestedFieldsArray
        StreamWithQuotes = static_StreamWithQuotes
        TargetsDefined = static_TargetsDefined
        TemplateDefined = static_TemplateDefined
        static_outputList.Reinitialize 1
        static_IsHeader = (static_Headers And static_RecordsCount = 0)
        If static_IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines static_Index, static_CSVstream, static_tmpCSV, static_MaxIndex, static_RecordDelimiter, static_commToken, static_FDAscW
        End If
        static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
        '@----------------------------------------------------------------------------
        'Restart pointers
        static_RecordEndReached = False
        static_FieldID = 0
        Do
            '@----------------------------------------------------------------------------
            'Process the current Token
            static_TokenEndReached = False
            static_TokenBeginningPos = static_OverStringPointer
            If StreamWithQuotes Then
                Select Case static_OverStringPointer
                    Case Is < static_LenCurrentIndex
                        static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces, tabs and remanents LF
                        Select Case static_FDAscW
                            Case Is <> 9
                                Do While static_ASCIIcharw = 32 Or static_ASCIIcharw = 9 Or static_ASCIIcharw = 10
                                    static_OverStringPointer = static_OverStringPointer + 2&
                                    If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                    static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While static_ASCIIcharw = 32 Or static_ASCIIcharw = 10
                                    static_OverStringPointer = static_OverStringPointer + 2&
                                    If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                    static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                Loop
                        End Select
                        static_OpenedToken = (static_ASCIIcharw = static_EscapeAscW)
                    Case Else 'Avoid over runs
                        static_OpenedToken = False
                End Select
            Else
                static_OpenedToken = False
            End If
            If static_OpenedToken Then
                '@----------------------------------------------------------------------------
                'Try to escape the Token
                static_TokenBeginningPos = static_OverStringPointer + 2&
                static_SearchBeginningMark = static_TokenBeginningPos
                static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_EscapeChr) 'EOFld marker
                Do While (Not static_TokenEndReached And Not static_EndOF)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not static_SplittedToken Then static_SplittedToken = (static_hpointer = 0)
                    Do While static_hpointer = 0& 'Escape char may be on subsequent static_Index
                        Select Case static_SearchBeginningMark
                            Case 1 'Store the full Token from the current static_Index
                                static_tmpToken.Add static_tmpCSV(static_Index)
                            Case Else 'Store the Right most Token portion from the current static_Index
                                static_tmpToken.Add RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex - static_SearchBeginningMark + 1)
                        End Select
                        static_Index = static_Index + 1&
                        static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                        If Not static_EndOFStreamBuffer Then
                            static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
                            static_SearchBeginningMark = 1
                            static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_EscapeChr) 'Find escape char
                        Else
                            If Not static_StreamEnd Then
                                '@----------------------------------------------------------------------------
                                'Load new stream
                                static_CSVstream.ReadText
                                StreamWithQuotes = InStrB(1, static_CSVstream.bufferString, static_EscapeChr)
                                static_StreamEnd = static_CSVstream.atEndOfStream
                                static_tmpCSV() = Split(static_CSVstream.bufferString, static_RecordDelimiter)
                                static_Index = 0
                                static_MaxIndex = UBound(static_tmpCSV)
                                If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                    static_MaxIndex = static_MaxIndex - 1
                                Else
                                    If static_tmpCSV(static_MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                                        static_MaxIndex = static_MaxIndex - 1
                                    End If
                                End If
                                static_EndOFStreamBuffer = False
                                static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
                                static_SearchBeginningMark = 1
                                static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_EscapeChr) 'Find escape char
                            Else
                                static_hpointer = 0&
                            End If
                        End If
                        static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                    Loop
                    If static_hpointer = 0& And static_EndOF Then GoTo GetRecord_MissingEscapeChar 'Missing escape char
                    static_OverStringPointer = static_hpointer + 2&
                    If static_OverStringPointer < static_LenCurrentIndex Then
                        static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces, tabs and remanents LF
                        Select Case static_FDAscW
                            Case Is <> 9
                                Do While static_ASCIIcharw = 32 Or static_ASCIIcharw = 9 Or static_ASCIIcharw = 10
                                    static_OverStringPointer = static_OverStringPointer + 2&
                                    If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                    static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While static_ASCIIcharw = 32 Or static_ASCIIcharw = 10
                                    static_OverStringPointer = static_OverStringPointer + 2&
                                    If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                    static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                Loop
                        End Select
                        '@----------------------------------------------------------------------------
                        'Check if we are at the end of the field
                        static_EscapedWFS = (static_ASCIIcharw = static_FDAscW)
                    Else
                        static_EscapedWFS = False
                    End If
                    If static_EscapedWFS Then 'Escape char succeeded by fields separator char
                        static_TokenEndingPos = static_hpointer - 1&
                        static_TokenEndReached = True
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        Select Case static_SplittedToken
                            Case False 'Just save Token
                                If RequestedField(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                    Select Case static_dynamicType
                                        Case False
                                            static_RecordToken.Add MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                            static_TokenEndingPos - static_TokenBeginningPos + 1)
                                        Case Else
                                            static_RecordToken.Add TypeData(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                            static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                            static_FieldID, _
                                                                            static_IsHeader)
                                    End Select
                                End If
                            Case Else 'The Token is stored on static_tmpToken object
                                static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_TokenEndingPos)
                                If RequestedField(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                    Select Case static_dynamicType
                                        Case False
                                            static_RecordToken.Add Join$(static_tmpToken.items, static_RecordDelimiter) 'Save whole Token
                                        Case Else
                                            static_RecordToken.Add TypeData(Join$(static_tmpToken.items, _
                                                                                static_RecordDelimiter), _
                                                                                static_FieldID, _
                                                                                static_IsHeader)
                                    End Select
                                    static_tmpToken.Clear 'Clear static_tmpToken content
                                End If
                        End Select
                        static_SplittedToken = False 'Restart Token status
                        static_OverStringPointer = static_OverStringPointer + 2& 'Advance over the field separator char
                        static_RecordEndReached = False
                    Else
                        If static_OverStringPointer >= static_LenCurrentIndex Then  'Escape char succeeded
                                                                      'by records separator char
                            static_TokenEndingPos = static_hpointer - 1&
                            static_TokenEndReached = True
                            static_RecordEndReached = True 'EOR
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case static_SplittedToken
                                Case False 'Just save Token
                                    If RequestedField(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                        Select Case static_dynamicType
                                            Case False
                                                static_RecordToken.Add MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                            static_TokenEndingPos - static_TokenBeginningPos + 1)
                                            Case Else
                                                static_RecordToken.Add TypeData(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                                    static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                                static_FieldID, _
                                                                                static_IsHeader)
                                        End Select
                                    End If
                                Case Else 'The Token is stored on static_tmpToken array
                                    static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_OverStringPointer - 3)
                                    If RequestedField(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                        Select Case static_dynamicType
                                            Case False
                                                static_RecordToken.Add Join$(static_tmpToken.items, static_RecordDelimiter) 'Save whole Token
                                            Case Else
                                                static_RecordToken.Add TypeData(Join$(static_tmpToken.items, _
                                                                                static_RecordDelimiter), _
                                                                                static_FieldID, _
                                                                                static_IsHeader)
                                        End Select
                                        static_tmpToken.Clear 'Clear static_tmpToken content
                                    End If
                            End Select
                            static_SplittedToken = False 'Restart Token status
                            static_Index = static_Index + 1&
                            static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                            static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                            If static_EndOFStreamBuffer And Not static_EndOF Then
                                '@----------------------------------------------
                                'Advance stream
                                static_CSVstream.ReadText
                                StreamWithQuotes = InStrB(1, static_CSVstream.bufferString, static_EscapeChr)
                                static_StreamEnd = static_CSVstream.atEndOfStream
                                static_tmpCSV() = Split(static_CSVstream.bufferString, static_RecordDelimiter)
                                static_Index = 0
                                static_MaxIndex = UBound(static_tmpCSV)
                                If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                    static_MaxIndex = static_MaxIndex - 1
                                Else
                                    If static_tmpCSV(static_MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                                        static_MaxIndex = static_MaxIndex - 1
                                    End If
                                End If
                                static_EndOFStreamBuffer = False
                            End If
                            static_OverStringPointer = 1&
                        Else
                            If static_SplittedToken Then 'Store the Left most Token portion from the current static_Index
                                static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_OverStringPointer - 1)
                            End If
                            static_SearchBeginningMark = static_OverStringPointer
                            static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_EscapeChr) 'Find escape char
                            static_TokenEndReached = False
                        End If
                    End If
                Loop
            Else
                static_SplittedToken = False
                '@----------------------------------------------------------------------------
                'Try to mark the Starting and Ending position for the field
                static_hpointer = InStrB(static_OverStringPointer, static_tmpCSV(static_Index), static_FieldDelimiter)
                If static_hpointer Then 'Found fields delimiter char
                    static_TokenEndReached = True
                    static_TokenEndingPos = static_hpointer - 1&
                    static_OverStringPointer = static_hpointer + 2&
                    static_RecordEndReached = False
                    '@----------------------------------------------------------------------------
                    'Save the Token
                    If RequestedField(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                        Select Case static_dynamicType
                            Case False
                                static_RecordToken.Add MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                    static_TokenEndingPos - static_TokenBeginningPos + 1)
                            Case Else
                                static_RecordToken.Add TypeData(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                            static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                            static_FieldID, _
                                                            static_IsHeader)
                        End Select
                    End If
                Else 'Missing fields delimiter char
                    static_TokenEndReached = True
                    static_RecordEndReached = True
                    '@----------------------------------------------------------------------------
                    'Save the Token
                    If RequestedField(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                        Select Case static_dynamicType
                            Case False
                                static_RecordToken.Add RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex _
                                                    - static_TokenBeginningPos + 1)
                            Case Else
                                static_RecordToken.Add TypeData(RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex _
                                                            - static_TokenBeginningPos + 1), _
                                                                static_FieldID, _
                                                                static_IsHeader)
                        End Select
                    End If
                    static_Index = static_Index + 1&
                    static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                    static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                    If static_EndOFStreamBuffer And Not static_EndOF Then
                        '@----------------------------------------------
                        'Advance stream
                        static_CSVstream.ReadText
                        StreamWithQuotes = InStrB(1, static_CSVstream.bufferString, static_EscapeChr)
                        static_StreamEnd = static_CSVstream.atEndOfStream
                        static_tmpCSV() = Split(static_CSVstream.bufferString, static_RecordDelimiter)
                        static_Index = 0
                        static_MaxIndex = UBound(static_tmpCSV)
                        If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                            static_MaxIndex = static_MaxIndex - 1
                        Else
                            If static_tmpCSV(static_MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                                static_MaxIndex = static_MaxIndex - 1
                            End If
                        End If
                        static_EndOFStreamBuffer = False
                    End If
                    static_OverStringPointer = 1&
                End If
            End If
            static_FieldID = static_FieldID + 1&
        Loop While (Not static_RecordEndReached And Not static_EndOF)
        Select Case static_vectorsBound
            Case Is > 0
                If static_RecordToken.count - 1 > static_vectorsMaxBound Then
                    static_vectorsMaxBound = static_RecordToken.count - 1
                    static_RecordToken.ShrinkBuffer
                End If
            Case Else
                Select Case static_RecordToken.count
                    Case Is > 0
                        static_vectorsBound = static_RecordToken.count - 1
                        static_RecordToken.ShrinkBuffer
                    Case Else
                        static_vectorsBound = static_RecordToken.count
                End Select
                static_vectorsMaxBound = static_vectorsBound
        End Select
        static_RecordsCount = static_RecordsCount + 1&
        '@----------------------------------------------------------------------------
        'Save record
        static_outputList.Add static_RecordToken.items
        static_RecordToken.Reinitialize static_vectorsBound + 1
        P_VARYING_LENGTHS = (static_vectorsBound <> static_vectorsMaxBound)
        If static_IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines static_Index, static_CSVstream, static_tmpCSV, _
                                    static_MaxIndex, static_RecordDelimiter, static_commToken, static_FDAscW
        End If
        static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
        static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
        If static_EndOFStreamBuffer And Not static_EndOF Then
            '@----------------------------------------------
            'Advance stream
            static_CSVstream.ReadText
            StreamWithQuotes = InStrB(1, static_CSVstream.bufferString, static_EscapeChr)
            static_StreamEnd = static_CSVstream.atEndOfStream
            static_tmpCSV() = Split(static_CSVstream.bufferString, static_RecordDelimiter)
            static_Index = 0
            static_MaxIndex = UBound(static_tmpCSV)
            If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                static_MaxIndex = static_MaxIndex - 1
            Else
                If static_tmpCSV(static_MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
                    static_MaxIndex = static_MaxIndex - 1
                End If
            End If
            static_EndOFStreamBuffer = False
        End If
        If Not static_EndOFStreamBuffer Then static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
        static_ImportSwitch = static_EndOF
        P_VECTORS_REGULAR_BOUND = static_vectorsBound
        P_VECTORS_MAX_BOUND = static_vectorsMaxBound
        '@----------------------------------------------------------------------------
        'Check if the array need to be shrinked at the output stage
        static_outputList.ShrinkBuffer
        Set GetRecord = static_outputList
        P_SUCCESSFUL_IMPORT = True
    Else
        Set GetRecord = Nothing
        CloseSeqReader
    End If
    Exit Function
GetRecord_Error_Handler:
    Set GetRecord = Nothing
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description & " [Check the record #" & static_RecordsCount + 1 _
                                  & ", field #" & static_RecordToken.count + 1 _
                                   & " on the source CSV file]."
    P_ERROR_SOURCE = Err.Source
    Exit Function
GetRecord_StartingIndexOutOfBound:
    Err.Raise Number:=vbObjectError + 9002, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="The static_Index over which the operation starts is out of bounds. Try start the import on a lower static_Index."
    Resume GetRecord_Error_Handler
GetRecord_MissingEscapeChar:
    Err.Raise Number:=vbObjectError + 9003, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="Missing some escape char. Check the data and try again."
    Resume GetRecord_Error_Handler
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
    Dim tmpBuffer As Variant
    If TypeName(RecordsArray) = "ECPArrayList" Then
        '@----------------------------------------------------------------------------
        'Set array sizes
        LB1 = 0
        UB1 = RecordsArray.count - 1
        ReDim ConcatenatedArray(LB1 To UB1)
        '@----------------------------------------------------------------------------
        'Concatenate fields
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
    ElseIf MultiDimensional(RecordsArray) Then
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

Private Function MaxIndexVal(ByRef values() As Long) As Long
    Dim tmpVal As Long, idx0 As Long, Idx As Long
    
    idx0 = LBound(values)
    Idx = idx0
    tmpVal = values(idx0)
    For idx0 = idx0 + 1 To UBound(values)
        If values(idx0) > tmpVal Then
            Idx = idx0
            tmpVal = values(idx0)
        End If
    Next idx0
    MaxIndexVal = Idx
End Function

Private Function MeetsCriterion(value As Variant) As Boolean
    Dim tmpResult As Boolean
    Dim vCounter As Long
    
    vCounter = LB
    Do While vCounter <= UB And tmpResult = False
        tmpResult = (value = queryFilters(vCounter, LBound(queryFilters, 2)))
        vCounter = vCounter + 1
    Loop
    MeetsCriterion = tmpResult
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

Public Sub OpenSeqReader(configObj As parserConfig, _
                            ParamArray FilterColumns() As Variant)
    On Error GoTo OpenSeqReader_Error_Handler
    If configObj.headers Then
        Dim tmpConf As parserConfig
        Set tmpConf = configObj.CopyConfig
        tmpConf.startingRecord = 1
        tmpConf.endingRecord = 1
        ImportFromCSV tmpConf
        Set tmpConf = Nothing
        static_tmpRequested() = FilterColumns
        Select Case UBound(static_tmpRequested)
            Case -1
                ReDim static_RequestedFieldsArray(0)
                static_RequestedFieldsArray(0) = -1
            Case Else
                static_RequestedFieldsArray() = GetLongFilters(static_tmpRequested, P_CSV_HEADER)
        End Select
    Else
        ReDim static_RequestedFieldsArray(0)
        static_RequestedFieldsArray(0) = -1
    End If
    '@----------------------------------------------------------------------------
    'Start variables
    static_SearchBeginningMark = 1&
    static_RecordsCount = 0&
    static_vectorsBound = 0&
    static_vectorsMaxBound = 0&
    Set static_tmpToken = New ECPArrayList
    Set static_CSVstream = New ECPTextStream
    With configObj
        static_CSVstream.bufferSize = .bufferSize
        static_CSVstream.endStreamOnLineBreak = True
        static_CSVstream.OpenStream .path
        static_commToken = AscW(.commentsToken)
        static_dynamicType = .dynamicTyping
        static_dTTemplate() = .dTypingTemplate
        static_dTTargets() = .dTypingLinks
        static_Headers = .headers
        static_IgnoreLines = .catchMeaninglessRecords
        static_TemplateDefined = .dTTemplateDefined
        static_TargetsDefined = .dTTemplateLinksDefined
        static_FieldDelimiter = .fieldsDelimiter
        static_RecordDelimiter = .recordsDelimiter
    End With
    static_TokenEndReached = False
    static_RecordEndReached = False
    '@----------------------------------------------------------------------------
    'Set variables
    static_OverStringPointer = static_SearchBeginningMark
    static_MaxReqIndex = UBound(static_RequestedFieldsArray)
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case configObj.escapeToken
        Case 1
            static_EscapeChr = CHR_APOSTROPHE
        Case Else
            static_EscapeChr = CHR_DOUBLE_QUOTES
    End Select
    static_EscapeAscW = AscW(static_EscapeChr)
    static_FDAscW = AscW(static_FieldDelimiter)
    Set static_RecordToken = New ECPArrayList
    Set static_outputList = New ECPArrayList
    static_Index = 0
    '@----------------------------------------------------------------------------
    'Get CSV stream and populate a temp array
    static_CSVstream.ReadText
    static_StreamWithQuotes = InStrB(1, static_CSVstream.bufferString, static_EscapeChr)
    static_StreamEnd = static_CSVstream.atEndOfStream
    static_tmpCSV() = Split(static_CSVstream.bufferString, static_RecordDelimiter)
    '@----------------------------------------------------------------------------
    'Ensuring the stream is sufficient enough to hold at least one CSV line
    Do While Not static_StreamEnd And UBound(static_tmpCSV) = 0
        static_CSVstream.bufferSize = 2 * configObj.bufferSize
        static_CSVstream.RestartPointer
        static_CSVstream.ReadText
        static_StreamWithQuotes = InStrB(1, static_CSVstream.bufferString, static_EscapeChr)
        static_StreamEnd = static_CSVstream.atEndOfStream
        static_tmpCSV() = Split(static_CSVstream.bufferString, static_RecordDelimiter)
    Loop
    static_MaxIndex = UBound(static_tmpCSV)
    If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
        static_MaxIndex = static_MaxIndex - 1
    Else
        If static_tmpCSV(static_MaxIndex) = vbLf Then 'vbCrLf splited as vbCr
            static_MaxIndex = static_MaxIndex - 1
        End If
    End If
    static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
    static_SeqReadStarted = True
    Exit Sub
OpenSeqReader_Error_Handler:
    static_SeqReadStarted = False
End Sub

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
    
    On Error GoTo ParseCSVstring_Error_Handler
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
    Do While LenB(tmpCSV(MaxIndex - brCounter)) = 0 Or tmpCSV(MaxIndex - brCounter) = vbLf
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
ParseCSVstring_Error_Handler:
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
    Resume ParseCSVstring_Error_Handler
ParseCriticalCSV_MissingEscapeChar:
    Err.Raise Number:=vbObjectError + 9003, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="Missing some escape char. Check the data and try again."
    Resume ParseCSVstring_Error_Handler
End Sub

Private Function RequestedField(ByVal FieldIndex As Long, _
                                ByVal mxReq As Long) As Boolean
    Select Case RequestedFieldsArray(0)
        Case -1
            RequestedField = True
        Case Else
            Dim Idx As Long
            Dim tmpBool As Boolean
            
            Idx = 0
            Do
                tmpBool = (FieldIndex = RequestedFieldsArray(Idx))
                Idx = Idx + 1
            Loop While Not tmpBool And Idx <= mxReq
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

Private Sub SkipUnwantedLines(ByRef Idx As Long, _
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
    If Idx <= MaxIdx Then
        Do
            CurLength = LenB(Arr(Idx))
            UnwantedLine = False
            StrPointer = 1&
            If CurLength = 0 Then 'Empty line found
                UnwantedLine = True
            Else
                CharCode = AscW(Arr(Idx))
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
                                CharCode = AscW(MidB$(Arr(Idx), StrPointer, 2))
                            Loop
                            If CharCode = CommentToken Then UnwantedLine = True
                        Case Else
                            Do While (CharCode >= 0& And CharCode <= 32&) And (StrPointer < CurLength And CharCode <> 9)
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    Exit Do
                                End If
                                CharCode = AscW(MidB$(Arr(Idx), StrPointer, 2))
                            Loop
                            If CharCode = CommentToken Then UnwantedLine = True
                    End Select
                End If
            End If
            If UnwantedLine Then
                Idx = Idx + 1&
                CurLength = LenB(Arr(Idx))
            End If
        Loop While UnwantedLine And Idx <= MaxIdx
    End If
End Sub

Public Function Sort(Optional ByVal fromIndex As Long = -1, _
                        Optional ByVal toIndex As Long = -1, _
                        Optional ByVal SortColumn As Long = -1, _
                        Optional ByVal Descending As Boolean = False) As CSVinterface
Attribute Sort.VB_Description = "Sorts the imported CSV data."
    
    On Error GoTo Sort_Error_Handler:
    
    If P_SUCCESSFUL_IMPORT Then
        If fromIndex = -1 Then
            If config.headers And Not config.headersOmission Then
                fromIndex = 2
            Else
                fromIndex = 1
            End If
        End If
        If toIndex = -1 Then
            toIndex = P_CSV_DATA.count
        End If
        If SortColumn = -1 Then
            SortColumn = 1 'Ensures a column to sort
        End If
        P_CSV_DATA.Sort fromIndex, toIndex, SortColumn, Descending
        Set Sort = Me
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
    
    On Error GoTo StreamParseCSV_Error_Handler
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
    StreamWithQuotes = InStrB(1, CSVstream.bufferString, EscapeChr)
    StreamEnd = CSVstream.atEndOfStream
    tmpCSV() = Split(CSVstream.bufferString, RecordDelimiter)
    '@----------------------------------------------------------------------------
    'Ensuring the stream is sufficient enough to hold at least one CSV line
    Do While Not StreamEnd And UBound(tmpCSV) = 0
        CSVstream.bufferSize = 2 * configObj.bufferSize
        CSVstream.RestartPointer
        CSVstream.ReadText
        StreamWithQuotes = InStrB(1, CSVstream.bufferString, EscapeChr)
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
                                StreamWithQuotes = InStrB(1, CSVstream.bufferString, EscapeChr)
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
                                StreamWithQuotes = InStrB(1, CSVstream.bufferString, EscapeChr)
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
                        StreamWithQuotes = InStrB(1, CSVstream.bufferString, EscapeChr)
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
            StreamWithQuotes = InStrB(1, CSVstream.bufferString, EscapeChr)
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
StreamParseCSV_Error_Handler:
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
    Resume StreamParseCSV_Error_Handler
ParseCriticalCSV_MissingEscapeChar:
    Err.Raise Number:=vbObjectError + 9003, _
                    Source:="CSVinterface Class [ParseCSV_RFC4180]", _
                    Description:="Missing some escape char. Check the data and try again."
    Resume StreamParseCSV_Error_Handler
End Sub

Private Sub StreamSkipUnwantedLines(ByRef Idx As Long, _
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
    If Idx <= MaxIdx Then
        Do
            CurLength = LenB(Arr(Idx))
            UnwantedLine = False
            StrPointer = 1&
            If CurLength = 0 Then 'Empty line found
                UnwantedLine = True
            Else
                CharCode = AscW(Arr(Idx))
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
                                CharCode = AscW(MidB$(Arr(Idx), StrPointer, 2))
                            Loop
                            UnwantedLine = (CharCode = CommentToken)
                        Case Else
                            Do While (CharCode >= 0& And CharCode <= 32&) And (StrPointer < CurLength And CharCode <> 9)
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    GoTo AdvanceLine
                                End If
                                CharCode = AscW(MidB$(Arr(Idx), StrPointer, 2))
                            Loop
                            UnwantedLine = (CharCode = CommentToken)
                    End Select
                End If
            End If
AdvanceLine:
            If UnwantedLine Then
                Idx = Idx + 1&
                If Idx > MaxIdx Then
                    If Not Ostream.atEndOfStream Then
                        '@----------------------------------------------
                        'Advance stream
                        Ostream.ReadText
                        StreamWithQuotes = InStrB(1, Ostream.bufferString, CHR_DOUBLE_QUOTES)
                        Arr() = Split(Ostream.bufferString, recDelimiter)
                        Idx = 0
                        MaxIdx = UBound(Arr)
                        If Arr(MaxIdx) = vbNullString Then
                            MaxIdx = MaxIdx - 1
                        End If
                    End If
                End If
            End If
        Loop While UnwantedLine And Idx <= MaxIdx
    Else
        If Not Ostream.atEndOfStream Then
            '@----------------------------------------------
            'Advance stream
            Ostream.ReadText
            StreamWithQuotes = InStrB(1, Ostream.bufferString, CHR_DOUBLE_QUOTES)
            Arr() = Split(Ostream.bufferString, recDelimiter)
            Idx = 0
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
