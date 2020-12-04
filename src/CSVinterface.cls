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
' Copyright © 2020 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' RFC-4180 compliant class module that allows users exchange data between VBA arrays and
' CSV files at high speed.
'#
' Feel free to visit the cited websites to explore some of the available solutions for
' work with CSV files and VBA Arrays.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
' CREDITS:
' Although this class have original pieces of code and logic, it use code,
' or hints, of many authors. Visit the web sources listed below:
' @1 https://senipah.github.io/VBA-Better-Array
' @2 https://github.com/sdkn104/VBA-CSV
' @3 http://www.vbaexpress.com/forum/showthread.php?25095-Solved-Check-the-file-encode-type
' @4 https://www.codegrepper.com/code-examples/vb/excel+vba+Load+csv+file+into+an+array+rather+than+the+Worksheet
' @5 http://excellerando.blogspot.com/2014/12/writing-excel-range-to-csv-file.html
' @6 https://stackoverflow.com/questions/4191560/create-csv-from-array-in-vba
' @7 https://www.freevbcode.com/ShowCode.asp?ID=3110
' @8 https://stackoverflow.com/questions/1376756/what-is-a-superfast-way-to-read-large-files-line-by-line-in-vba
' @9 https://www.freevbcode.com/ShowCode.asp?ID=7655
' @10 https://www.exceltip.com/excel-macros-and-vba
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const CHR_APOSTROPHE As String = "'"
Private Const CHR_COMMA As String = ","
Private Const CHR_DOUBLE_QUOTES As String = """"
Private Const CHR_SEMICOLON As String = ";"
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private P_ASCW_FIELDS_DELIMITER As Long
Private P_COMMENTSTOKEN As Long
Private P_CONNECTED As Boolean '-----------------Holds the connection state
Private P_CSV_DATA() As Variant '----------------Holds the CSV data for current instance
Private P_ENDING_RECORD As Long '----------------Pointer to the last record to be imported
Private P_ERROR_DESC As String '-----------------Holds the Error Description
Private P_ERROR_NUMBER As Long '-----------------Holds the Error number
Private P_ERROR_SOURCE As String '---------------Holds the Error source
Private P_ESCAPE_TOKEN As EscapeTokens '---------Holds the char used for escape fields
Private P_FIELDS_DELIMITER As String '-----------Holds the current delimiter for CSV fields
Private P_FILEENCODING As String '---------------Holds the encoding for the opened CSV file
Private P_FILENAME As String '-------------------Holds the actual file name
Private P_ITEMS As Long '------------------------Holds the amount of data rows
Private P_QUOTING_MODE As QuotationMode '--------Hols the mode used for fields quotation.
Private P_RECORDS_DELIMITER As String '----------Holds the records delimiter (CR, LF or CRLF)
Private P_STARTING_RECORD As Long '--------------Pointer to the first record to be imported
Private P_SUCCESSFUL_EXPORT As Boolean
Private P_SUCCESSFUL_IMPORT As Boolean
Private P_VARYING_LENGTHS As Boolean '-----------Indicates if the read CSV has varying number _
                                                 of fields per records
Private P_IRREGULAR_VECTORS_INFO As Collection '-Holds the INFO for vectors exceeding the _
                                                 P_VECTORS_MAX_BOUND limit.
Private P_VECTORS_MAX_BOUND As Long '------------Maximum index of vectors on jagged array
Private P_VECTORS_REGULAR_BOUND As Long '--------Regular upper index of vectors on jagged array
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum EscapeTokens
    Apostrophe = 1
    DoubleQuotes = 2
    NullChar = 0
End Enum
Public Enum QuotationMode
    Critical = 0 '--------------Only fields that includes special chars.
    All = 1
End Enum
Public Enum abCharsets
    abError = 0
    abANSI = 1
    abUnicode = 2
    abUnicodeBigEndian = 3
    abUTF8 = 4
    ebUnknown = 5
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
' @Common
Private DataDelimiter As String '----------Concatenation of Fields separator and CoerceType
Private LinesDelimiter As String '---------Concatenation of Records separator and CoerceType
Private FileContent As String '------------Holds the whole CSV file into this string
Private FileHandled As Integer '-----------Pointer to the actual CSV file
Private VectorsBoundData(0 To 1) As Long '-Holds the max index for vectors at the _
                                           P_VECTORS_MAX_BOUND limit
'////////////////////////////////////////////////////////////////////////////////////////////
'#

'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
Public Property Get CommentsToken() As String
Attribute CommentsToken.VB_Description = "Gets or sets the char used for identify comments lines."
    CommentsToken = ChrW(P_COMMENTSTOKEN)
End Property
Public Property Let CommentsToken(Token As String)
    P_COMMENTSTOKEN = AscW(Token)
End Property
Public Property Get Connected() As Boolean
Attribute Connected.VB_Description = "Gets the CSV file connection state. Returns True when the CSV file was opened in memory."
    Connected = P_CONNECTED
End Property
Public Property Get Count() As Long
    Count = P_ITEMS
End Property
Public Property Get EndingRecord() As Long
Attribute EndingRecord.VB_Description = "Gets or sets the record over which the import procedure will ends."
    EndingRecord = P_ENDING_RECORD
End Property
Public Property Let EndingRecord(ByVal RecNumber As Long)
    Select Case RecNumber
        Case Is >= 1&, 0&
            P_ENDING_RECORD = RecNumber
        Case Else
            P_ENDING_RECORD = 1&
    End Select
End Property
Public Property Get ErrDescription() As String
Attribute ErrDescription.VB_Description = "Gets the description for the last occurred error over the current instance."
    ErrDescription = P_ERROR_DESC
End Property
Public Property Get ErrNumber() As Long
Attribute ErrNumber.VB_Description = "Gets the number for the last occurred error over the current instance."
    ErrNumber = P_ERROR_NUMBER
End Property
Public Property Get ErrSource() As String
Attribute ErrSource.VB_Description = "Gets the source for the last occurred error over the current instance."
    ErrSource = P_ERROR_SOURCE
End Property
Public Property Get EscapeToken() As EscapeTokens
Attribute EscapeToken.VB_Description = "Gets or sets the fields escape char for the current instance."
    EscapeToken = P_ESCAPE_TOKEN
End Property
Public Property Let EscapeToken(ByVal Token As EscapeTokens)
    P_ESCAPE_TOKEN = Token
End Property
Public Property Get ExportSuccess() As Boolean
Attribute ExportSuccess.VB_Description = "Gets the status for the last export operation."
    ExportSuccess = P_SUCCESSFUL_EXPORT
End Property
Public Property Get FieldsDelimiter() As String
Attribute FieldsDelimiter.VB_Description = "Gets or sets the char used as delimiter for the CSV file fields in the current instance."
    FieldsDelimiter = P_FIELDS_DELIMITER
End Property
Public Property Let FieldsDelimiter(ByVal Delimiter As String)
    Select Case Delimiter
        Case CHR_COMMA, CHR_SEMICOLON, vbTab
            P_FIELDS_DELIMITER = Delimiter
        Case Else
            P_FIELDS_DELIMITER = CHR_COMMA
    End Select
    P_ASCW_FIELDS_DELIMITER = AscW(P_FIELDS_DELIMITER)
End Property
Public Property Get FileEncoding() As String
Attribute FileEncoding.VB_Description = "Gets the charset used to encode the last opened CSV file."
    FileEncoding = P_FILEENCODING
End Property
Public Property Get fileName() As String
Attribute fileName.VB_Description = "Gets the name of the last opened CSV file."
    fileName = P_FILENAME
End Property
Public Property Get ImportSuccess() As Boolean
Attribute ImportSuccess.VB_Description = "Gets the status for the last import operation. Returns True if it was success."
    ImportSuccess = P_SUCCESSFUL_IMPORT
End Property
Public Property Get IrregularVectors() As Collection
Attribute IrregularVectors.VB_Description = "Gets a collection of arrays with INFO for irregular vectors on the current instance."
    Set IrregularVectors = P_IRREGULAR_VECTORS_INFO
End Property
Public Property Get Item(ParamArray Index() As Variant) As Variant
Attribute Item.VB_Description = "Gets a field, or an array with an entire record, from the result array on the current instance."
Attribute Item.VB_UserMemId = 0
    Select Case UBound(Index)
        Case 0
            Item = P_CSV_DATA(Index(0))
        Case Is >= 1
            Item = P_CSV_DATA(Index(0))(Index(1))
    End Select
End Property
Public Property Get QuotingMode() As QuotationMode
Attribute QuotingMode.VB_Description = "Gets or sets the mode used for fields quoting."
    QuotingMode = P_QUOTING_MODE
End Property
Public Property Let QuotingMode(mode As QuotationMode)
    P_QUOTING_MODE = mode
End Property
Public Property Get RecordsDelimiter() As String
Attribute RecordsDelimiter.VB_Description = "Gets or sets the char used as delimiter for the CSV file records in the current instance."
    RecordsDelimiter = P_RECORDS_DELIMITER
End Property
Public Property Let RecordsDelimiter(Delimiter As String)
    Select Case Delimiter
        Case vbCrLf, vbCr, vbLf
            P_RECORDS_DELIMITER = Delimiter
        Case Else
            P_RECORDS_DELIMITER = vbCrLf
    End Select
End Property
Public Property Get RectangularResults() As Boolean
Attribute RectangularResults.VB_Description = "Gets the result array regularity on the current instance. If true, the internal array is not irregular."
    RectangularResults = Not P_VARYING_LENGTHS
End Property
Public Property Get StartingRecord() As Long
Attribute StartingRecord.VB_Description = "Gets or sets the record over which the import procedure will begin."
    StartingRecord = P_STARTING_RECORD
End Property
Public Property Let StartingRecord(ByVal RecNumber As Long)
    Select Case RecNumber
        Case Is >= 1&
            P_STARTING_RECORD = RecNumber
        Case Else
            P_STARTING_RECORD = 1&
    End Select
End Property
Public Property Get VectorsBound() As Long
Attribute VectorsBound.VB_Description = "Gets the regular bound of the vectors in the result array on the current instance."
    VectorsBound = P_VECTORS_REGULAR_BOUND
End Property
Public Property Get VectorsMaxBound() As Long
Attribute VectorsMaxBound.VB_Description = "Gets the max bound of the vectors in the result array on the current instance."
    VectorsMaxBound = P_VECTORS_MAX_BOUND
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#

Private Function CharsetToString(ByVal Value As abCharsets) As String
Attribute CharsetToString.VB_Description = "Returns a description for the charset based in the Charsets Enumeration."
    Dim strRtnVal As String
    Select Case Value
        Case abCharsets.abANSI
            strRtnVal = "ANSI"
        Case abCharsets.abUTF8
            strRtnVal = "UTF-8"
        Case abCharsets.abUnicode
            strRtnVal = "Unicode"
        Case abCharsets.abUnicodeBigEndian
            strRtnVal = "BigEndian"
        Case Else
            strRtnVal = "Unknown"
    End Select
    CharsetToString = strRtnVal
End Function
Private Sub CloseConnection()
Attribute CloseConnection.VB_Description = "Closes a connection that point to the memory opened CSV file."
    On Error GoTo ErrHandler_CloseConnection
    '@------------------------------------------------------
    'Close the last opened file
    Close #FileHandled
    P_ERROR_DESC = vbNullString
    P_ERROR_NUMBER = 0&
    P_ERROR_SOURCE = vbNullString
    Exit Sub
ErrHandler_CloseConnection:
    P_CONNECTED = False
    P_ERROR_DESC = "[Close connection]: " & Err.Description
    P_ERROR_NUMBER = Err.Number
    P_ERROR_SOURCE = Err.Source
End Sub
Public Sub CreateJagged(ByRef ArrVar() As Variant, ArraySize As Long, VectorSize As Long)
Attribute CreateJagged.VB_Description = "Creates an empty array of vectors, each of which having a fixed custom size."
    Dim Vector() As String
    Dim arrPointer As Long
    
    ReDim ArrVar(0 To ArraySize)
    ReDim Vector(0 To VectorSize)

    For arrPointer = 0 To ArraySize
        ArrVar(arrPointer) = Vector()
    Next arrPointer
    Erase Vector
End Sub
Private Function CSVcolumns(csvArray() As String, _
                                        Optional FieldsDelimiter As String = CHR_COMMA) As Long
Attribute CSVcolumns.VB_Description = "Computes fields amount for a CSV file Only works with the ParseCSV method."
    Dim LB As Long, UB As Long
    Dim lngPos As Long
    Dim MDCounter As Long
    Dim NumOfDelimiter As Long
    
    On Error GoTo ErrHandler_CSVcolumns
    LB = LBound(csvArray)
    UB = UBound(csvArray)
    NumOfDelimiter = 0
    MDCounter = 0
    lngPos = CLng(InStrB(csvArray(LB), FieldsDelimiter))
    Do While lngPos
        MDCounter = MDCounter + 1
        lngPos = CLng(InStrB(lngPos + 2, csvArray(LB), FieldsDelimiter))
    Loop
    NumOfDelimiter = MDCounter
    CSVcolumns = NumOfDelimiter + 1
    Exit Function
ErrHandler_CSVcolumns:
    CSVcolumns = -1
End Function
Public Sub DumpToArray(OutPutArray() As String)
Attribute DumpToArray.VB_Description = "Dumps the CSV data from the current instance to a 2D array."
    If P_SUCCESSFUL_IMPORT Then
        Dim UB1 As Long, UB2 As Long
        Dim Counter1 As Long, Counter2 As Long
        
        UB1 = UBound(P_CSV_DATA)
        ReDim OutPutArray(0 To UB1, 0 To P_VECTORS_MAX_BOUND)
        
        For Counter1 = 0 To UB1
            UB2 = UBound(P_CSV_DATA(Counter1))
            For Counter2 = 0 To UB2
                OutPutArray(Counter1, Counter2) = P_CSV_DATA(Counter1)(Counter2)
            Next Counter2
            Erase P_CSV_DATA(Counter1) 'Release memory
        Next Counter1
        P_SUCCESSFUL_IMPORT = False
        Erase P_CSV_DATA
    End If
End Sub
Public Sub DumpToJaggedArray(OutPutArray() As Variant)
Attribute DumpToJaggedArray.VB_Description = "Dumps the CSV data from the current instance to a jagged array."
    If P_SUCCESSFUL_IMPORT Then
        Dim UB1 As Long
        Dim Counter1 As Long
        
        UB1 = UBound(P_CSV_DATA)
        ReDim OutPutArray(0 To UB1)
        
        For Counter1 = 0 To UB1
            OutPutArray(Counter1) = P_CSV_DATA(Counter1)
            Erase P_CSV_DATA(Counter1) 'Release memory
        Next Counter1
        P_SUCCESSFUL_IMPORT = False
        Erase P_CSV_DATA
    End If
End Sub
Public Sub DumpToSheet(Optional WBookName As String, _
                        Optional SheetName As String, _
                        Optional rngName As String = "A1")
Attribute DumpToSheet.VB_Description = "Dumps the CSV data from the current instance to an Excel Worksheet."
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
        Dim UB1 As Long, UB2 As Long
        Dim Counter1 As Long, Counter2 As Long
        
        UB1 = UBound(P_CSV_DATA)
        ReDim tmpOutputArray(0 To UB1, 0 To P_VECTORS_MAX_BOUND)
        
        For Counter1 = 0 To UB1
            UB2 = UBound(P_CSV_DATA(Counter1))
            For Counter2 = 0 To UB2
                tmpOutputArray(Counter1, Counter2) = P_CSV_DATA(Counter1)(Counter2)
            Next Counter2
            Erase P_CSV_DATA(Counter1) 'Release memory
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
        P_SUCCESSFUL_IMPORT = False
        Erase P_CSV_DATA
    End If
End Sub
Private Sub EnableOptimization(Optional Optimize As Boolean = True)
Attribute EnableOptimization.VB_Description = "Turns On and Off data-dump operation optimization over an Excel Worksheet."
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
Public Sub ExportToCSV(ByRef csvArray As Variant, Optional PassControlToOS As Boolean = True)
Attribute ExportToCSV.VB_Description = "Exports a 2D or jagged array to a CSV file located in the path specified with the OpenConnection method."
    On Error GoTo ErrHandler_ExportToCSV
    If PassControlToOS Then DoEvents 'Pass the control to the Operative System
    If P_CONNECTED And IsArray(csvArray) Then
        '@----------------------------------------------------------------------------
        'Join all the records and write the data on the current opened file
        Put #FileHandled, , JoinRecordsFields(csvArray)
        '@----------------------------------------------------------------------------
        If P_ERROR_NUMBER = 0& Then
            Call CloseConnection
            P_SUCCESSFUL_EXPORT = True
            P_ERROR_DESC = vbNullString
            P_ERROR_SOURCE = vbNullString
        Else
            GoTo ErrHandler_ExportToCSV
        End If
    Else
        P_SUCCESSFUL_EXPORT = False
        P_ERROR_DESC = "[CSV file Export]: the object is not linked to a CSV file or " _
                        & "the passed argument isn't an array. " _
                        & "Ensure use the OpenConnection method before export data."
        P_ERROR_NUMBER = vbObjectError + 9009
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Exit Sub
ErrHandler_ExportToCSV:
    Call CloseConnection
    P_SUCCESSFUL_EXPORT = False
    P_ERROR_DESC = "[CSV file Export]: " & P_ERROR_DESC
End Sub
Private Function FileExists(ByVal filePath As String) As Boolean
Attribute FileExists.VB_Description = "Determines if the file exist in the specified path."
    FileExists = CBool(LenB(Dir(filePath, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)))
End Function
Private Sub GetCSVtext(ByRef OutPutVar As String)
Attribute GetCSVtext.VB_Description = "Gets the CSV content from a file located in the path specified in the OpenConnection method."
    On Error Resume Next
    OutPutVar = SPACE$(LOF(FileHandled))
    Get #FileHandled, , OutPutVar
    Call CloseConnection
End Sub
Public Function GetDataFromCSV(csvPathAndFilename As String) As String
Attribute GetDataFromCSV.VB_Description = "Gets the CSV content from a file located in the specified path."
    
    On Error GoTo ErrHandler_GetDataFromCSV
    Call OpenConnection(csvPathAndFilename)
    GetDataFromCSV = SPACE$(LOF(FileHandled))
    Get #FileHandled, , GetDataFromCSV
    Call CloseConnection
ErrHandler_GetDataFromCSV:
    P_CONNECTED = False
    P_FILENAME = vbNullString
    P_ERROR_DESC = "[Get Data From CSV]: " & Err.Description
    P_ERROR_NUMBER = Err.Number
    P_ERROR_SOURCE = Err.Source
End Function
Private Function GetFileBytes(Optional ByVal truncateToByte As Long = -1) As Byte()
Attribute GetFileBytes.VB_Description = "Reads the CSV file and returns its content in a Byte array."
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    If truncateToByte < 0 Then
        truncateToByte = LOF(FileHandled) - 1
    End If
    lngFileNum = FileHandled
    ReDim bytRtnVal(truncateToByte) As Byte
    Get lngFileNum, , bytRtnVal
    Close lngFileNum
    Open P_FILENAME For Binary As #lngFileNum
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function
Private Function GetFileEncoding() As String
Attribute GetFileEncoding.VB_Description = "Reads the last opened CSV file and returns a string representing its encoding."
    GetFileEncoding = CharsetToString(ReturnCharset)
End Function
Private Function GetRecordIndex(ByVal RecordIndex As Long, _
                                ByRef CSVdataArray() As String, _
                                ByRef MaxIndex As Long) As Long
Attribute GetRecordIndex.VB_Description = "Returns the Index in wich the specified record will start."
    Dim ASCIIcharw As Long
    Dim EscapedWFS As Boolean
    Dim EscapeAscW As Long
    Dim EscapeChr As String
    Dim FDAscW As Long
    Dim Index As Long
    Dim LenCurrentIndex As Long
    Dim OpenedToken As Boolean
    Dim OverStringPointer As Long, hpointer As Long
    Dim RecordEndReached As Boolean
    Dim RecordsCount As Long
    Dim SearchBeginningMark As Long
    Dim SplittedToken As Boolean
    Dim TokenEndReached As Boolean
    
    '@----------------------------------------------------------------------------
    'Start variables
    SearchBeginningMark = 1&
    RecordsCount = 0&
    TokenEndReached = False
    '@----------------------------------------------------------------------------
    'Set variables
    OverStringPointer = SearchBeginningMark
    '@----------------------------------------------------------------------------
    'Set the pointers to the CSVdataArray array
    Index = LBound(CSVdataArray)
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case P_ESCAPE_TOKEN
        Case 1
            EscapeChr = CHR_APOSTROPHE
        Case Else
            EscapeChr = CHR_DOUBLE_QUOTES
    End Select
    EscapeAscW = AscW(EscapeChr)
    FDAscW = AscW(P_FIELDS_DELIMITER)
    '@----------------------------------------------------------------------------
    'Skip lines if needed
    SkipUnwantedLines Index, MaxIndex, CSVdataArray
    '@----------------------------------------------------------------------------
    'Loop through CSVdataArray array
    If RecordIndex >= 1 Then
        RecordIndex = RecordIndex - 1&
        LenCurrentIndex = LenB(CSVdataArray(Index))
        Do While Index <= MaxIndex And RecordsCount < RecordIndex
            RecordEndReached = False 'Restart Record status
            Do While (Not RecordEndReached And Index <= MaxIndex)
                '@----------------------------------------------------------------------------
                'Process the current Token
                TokenEndReached = False
                Select Case OverStringPointer
                    Case Is < LenCurrentIndex
                        ASCIIcharw = AscW(MidB$(CSVdataArray(Index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces and tabs
                        Select Case P_ASCW_FIELDS_DELIMITER
                            Case Is <> 9
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 9
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(CSVdataArray(Index), OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While ASCIIcharw = 32
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(CSVdataArray(Index), OverStringPointer, 2))
                                Loop
                        End Select
                        OpenedToken = (ASCIIcharw = EscapeAscW)
                    Case Else 'Avoid over runs
                        OpenedToken = False
                End Select
                If OpenedToken Then
                    '@----------------------------------------------------------------------------
                    'Try to escape the Token
                    SearchBeginningMark = OverStringPointer + 2&
                    hpointer = InStrB(SearchBeginningMark, CSVdataArray(Index), EscapeChr) 'EOFld marker
                    Do While (Not TokenEndReached And Index <= MaxIndex)
                        '@----------------------------------------------------------------------------
                        'Buffering if needed
                        If Not SplittedToken Then SplittedToken = (hpointer = 0)
                        Do While hpointer = 0& 'Escape char may be on subsequent Index
                            CSVdataArray(Index) = vbNullString
                            Index = Index + 1&
                            If Index <= MaxIndex Then LenCurrentIndex = LenB(CSVdataArray(Index))
                            SearchBeginningMark = 1
                            hpointer = InStrB(SearchBeginningMark, CSVdataArray(Index), EscapeChr) 'Find escape char
                        Loop
                        If hpointer = 0& And Index > MaxIndex Then GoTo GetRecordIndex_Error_Handler 'Missing escape char
                        OverStringPointer = hpointer + 2&
                        LenCurrentIndex = LenB(CSVdataArray(Index))
                        If OverStringPointer < LenCurrentIndex Then
                            ASCIIcharw = AscW(MidB$(CSVdataArray(Index), OverStringPointer, 2))
                            '@----------------------------------------------------------------------------
                            'Ignore spaces and tabs
                            Select Case P_ASCW_FIELDS_DELIMITER
                                Case Is <> 9
                                    Do While ASCIIcharw = 32 Or ASCIIcharw = 9
                                        OverStringPointer = OverStringPointer + 2&
                                        If OverStringPointer > LenCurrentIndex Then Exit Do
                                        ASCIIcharw = AscW(MidB$(CSVdataArray(Index), OverStringPointer, 2))
                                    Loop
                                Case Else
                                    Do While ASCIIcharw = 32
                                        OverStringPointer = OverStringPointer + 2&
                                        If OverStringPointer > LenCurrentIndex Then Exit Do
                                        ASCIIcharw = AscW(MidB$(CSVdataArray(Index), OverStringPointer, 2))
                                    Loop
                            End Select
                            '@----------------------------------------------------------------------------
                            'Check if we are at the end of the field
                            EscapedWFS = (ASCIIcharw = FDAscW)
                        Else
                            EscapedWFS = False
                        End If
                        If EscapedWFS Then 'Escape char succeeded by fields separator char
                            TokenEndReached = True
                            RecordEndReached = False
                            SplittedToken = False 'Restart Token status
                            OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                        Else
                            If OverStringPointer >= LenCurrentIndex Then  'Escape char succeeded
                                                                                        'by records separator char
                                TokenEndReached = True
                                RecordEndReached = True 'EOR
                                SplittedToken = False 'Restart Token status
                                CSVdataArray(Index) = vbNullString
                                Index = Index + 1&
                                If Index <= MaxIndex Then LenCurrentIndex = LenB(CSVdataArray(Index))
                                OverStringPointer = 1&
                            Else
                                SearchBeginningMark = OverStringPointer
                                hpointer = InStrB(SearchBeginningMark, CSVdataArray(Index), EscapeChr) 'Find escape char
                                TokenEndReached = False
                            End If
                        End If
                    Loop
                Else
                    SplittedToken = False
                    '@----------------------------------------------------------------------------
                    'Try to mark the Starting and Ending position for the field
                    hpointer = InStrB(OverStringPointer, CSVdataArray(Index), P_FIELDS_DELIMITER)
                    If hpointer Then 'Found fields delimiter char
                        TokenEndReached = True
                        RecordEndReached = False
                        OverStringPointer = hpointer + 2&
                    Else 'Missing fields delimiter char
                        TokenEndReached = True
                        RecordEndReached = True
                        CSVdataArray(Index) = vbNullString
                        Index = Index + 1&
                        If Index <= MaxIndex Then LenCurrentIndex = LenB(CSVdataArray(Index))
                        OverStringPointer = 1&
                    End If
                End If
            Loop
            RecordsCount = RecordsCount + 1&
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            SkipUnwantedLines Index, MaxIndex, CSVdataArray
            If Index <= MaxIndex Then LenCurrentIndex = LenB(CSVdataArray(Index))
        Loop
        If Index <= MaxIndex Then
            GetRecordIndex = Index
        Else
            GetRecordIndex = -1& 'Not found
        End If
    Else
        GetRecordIndex = -2& 'Invalid Input
    End If
    Exit Function
GetRecordIndex_Error_Handler:
    GetRecordIndex = -3& 'An error was occurs
End Function
Public Sub ImportFromCSV(Optional HeadersOmission As Boolean = False, _
                            Optional PassControlToOS As Boolean = True)
Attribute ImportFromCSV.VB_Description = "Imports the content of a CSV file, located in the path specified with the OpenConnection method, to the current instance."
    If PassControlToOS Then DoEvents 'Pass the control to the Operative System
    If P_CONNECTED Then
        '@--------------------------------------------------------------------------------
        'Read whole file content
        GetCSVtext FileContent
        If LenB(FileContent) = 0 Then
            Call CreateJagged(P_CSV_DATA, 0, 0)
            P_CSV_DATA(0) = Split(FileContent, vbCrLf)
            Exit Sub
        End If
        '@--------------------------------------------------------------------------------
        'Try to load the data
        If Err.Number = 0 Then
            Select Case P_QUOTING_MODE
                Case QuotationMode.Critical
                    ParseCriticalCSV FileContent, HeadersOmission
                Case Else
                    ParseCSV FileContent, HeadersOmission
            End Select
            '@----------------------------------------------------------------------------
            'Free the memory
            FileContent = vbNullString
            '@----------------------------------------------------------------------------
            If P_ERROR_NUMBER = 0& Then
                'Save pointers
                P_SUCCESSFUL_IMPORT = True
                P_ERROR_DESC = vbNullString
                P_ERROR_SOURCE = vbNullString
            Else
                GoTo ErrHandler_ImportFromCSV
            End If
        Else
            GoTo ErrHandler_ImportFromCSV
        End If
    Else
        P_SUCCESSFUL_IMPORT = False
        P_ERROR_DESC = "[CSV file Import]: the object is not linked to a CSV file. " _
                        & "Ensure use the OpenConnection method before import data."
        P_ERROR_NUMBER = vbObjectError + 9007
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Exit Sub
ErrHandler_ImportFromCSV:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[CSV file Import]: " & P_ERROR_DESC
End Sub
Public Sub ImportFromCSVString(ByRef CSVString As String, _
                                Optional HeadersOmission As Boolean = False, _
                                Optional PassControlToOS As Boolean = False)
Attribute ImportFromCSVString.VB_Description = "Imports the content of a CSV file, holded in a string variable, to the current instance."
    If PassControlToOS Then DoEvents 'Pass the control to the Operative System
    If LenB(CSVString) = 0 Then
        Call CreateJagged(P_CSV_DATA, 0, 0)
        P_CSV_DATA(0) = Split(CSVString, vbCrLf)
        Exit Sub
    End If
    '@--------------------------------------------------------------------------------
    'Try to load the data
    Select Case P_QUOTING_MODE
        Case QuotationMode.Critical
            ParseCriticalCSV CSVString, HeadersOmission
        Case Else
            ParseCSV CSVString, HeadersOmission
    End Select
    '@----------------------------------------------------------------------------
    If P_ERROR_NUMBER = 0& Then
        'Save pointers
        P_SUCCESSFUL_IMPORT = True
        P_ERROR_DESC = vbNullString
        P_ERROR_SOURCE = vbNullString
    Else
        GoTo ErrHandler_ImportFromString
    End If
    Exit Sub
ErrHandler_ImportFromString:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[CSV file Import]: " & P_ERROR_DESC
End Sub
Private Function IsANSI() As Boolean
Attribute IsANSI.VB_Description = "Verifies the charset for ANSI encoding."
    Const lngKeyCodeNullChar_c As Long = 0
    Dim bytFile() As Byte
    Dim lngIndx As Long
    Dim lngUprBnd As Long
    bytFile = GetFileBytes
    lngUprBnd = UBound(bytFile)
    For lngIndx = 0 To lngUprBnd
        If bytFile(lngIndx) = lngKeyCodeNullChar_c Then
            Exit For
        End If
    Next
    Erase bytFile
    IsANSI = (lngIndx > lngUprBnd)
End Function
Public Function IsJaggedArray(Arr As Variant) As Boolean
Attribute IsJaggedArray.VB_Description = "Checks if the given array is an array of arrays."
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
Attribute IsSheetInWorkbook.VB_Description = "Checks if the given sheet is part of the given Workbook."
    With WBook
        On Error Resume Next
        IsSheetInWorkbook = (.Sheets(SheetName).Name = SheetName)
        On Error GoTo 0
    End With
End Function
Private Function IsWorkbookOpen(WBookName As String) As Boolean
Attribute IsWorkbookOpen.VB_Description = "Checks if the given Workbook is opened yet."
    Dim WBook As Workbook, BookMatching As Boolean
    
    On Error Resume Next
    For Each WBook In Workbooks
        BookMatching = (WBook.Name = WBookName)
    Next
    IsWorkbookOpen = BookMatching
End Function
Public Sub JaggedToTwoDimArray(ByRef JaggedArray() As Variant, ByRef TwoDimArray() As String)
Attribute JaggedToTwoDimArray.VB_Description = "Deconstructs a jagged array and puts its content into a 2D string array."
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
Private Function JoinRecordsFields(ByRef RecordsArray As Variant) As String
Attribute JoinRecordsFields.VB_Description = "Joins all fields for each record in the specified array."
    Dim Buffer() As String
    Dim CoerceChr As String
    Dim ConcatenatedArray() As String
    Dim iLCounter As Long, jLCounter As Long
    Dim IsJagged As Boolean
    Dim JoinBuffer(0 To 2) As String
    Dim LB1 As Long, UB1 As Long
    Dim LB2 As Long, UB2 As Long
    
    On Error GoTo JoinRecordsFields_Error
    '@----------------------------------------------------------------------------
    'Set coerce char
    Select Case P_ESCAPE_TOKEN
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
        UB2 = UBound(RecordsArray, 2)
        Select Case P_QUOTING_MODE
            Case QuotationMode.Critical
                '@----------------------------------------------------------------------------
                'Set array sizes
                ReDim ConcatenatedArray(LB1 To UB1)
                ReDim Buffer(LB2 To UB2)
                '@----------------------------------------------------------------------------
                'Concatenate fields
                For iLCounter = LB1 To UB1
                    For jLCounter = LB2 To UB2
                        If InStrB(1, RecordsArray(iLCounter, jLCounter), P_FIELDS_DELIMITER) Then
                            JoinBuffer(1) = RecordsArray(iLCounter, jLCounter)
                            Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                        ElseIf InStrB(1, RecordsArray(iLCounter, jLCounter), CoerceChr) Then
                            JoinBuffer(1) = RecordsArray(iLCounter, jLCounter)
                            Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                        ElseIf InStrB(1, RecordsArray(iLCounter, jLCounter), P_RECORDS_DELIMITER) Then
                            JoinBuffer(1) = RecordsArray(iLCounter, jLCounter)
                            Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                        Else
                            Buffer(jLCounter) = RecordsArray(iLCounter, jLCounter)
                        End If
                        JoinBuffer(1) = vbNullString
                    Next jLCounter
                    ConcatenatedArray(iLCounter) = Join$(Buffer, P_FIELDS_DELIMITER)
                Next iLCounter
                '@----------------------------------------------------------------------------
                'Concatenate records
                JoinRecordsFields = Join$(ConcatenatedArray, P_RECORDS_DELIMITER)
                '@----------------------------------------------------------------------------
                'Free the memory
                Erase ConcatenatedArray
                Erase Buffer
            Case Else
                '@----------------------------------------------------------------------------
                'Set array sizes
                ReDim ConcatenatedArray(LB1 To UB1)
                ReDim Buffer(LB2 To UB2)
                '@----------------------------------------------------------------------------
                'Concatenate fields
                For iLCounter = LB1 To UB1
                    For jLCounter = LB2 To UB2
                        Buffer(jLCounter) = RecordsArray(iLCounter, jLCounter)
                    Next jLCounter
                    JoinBuffer(1) = Join$(Buffer, DataDelimiter)
                    ConcatenatedArray(iLCounter) = Join$(JoinBuffer, vbNullString)
                Next iLCounter
                '@----------------------------------------------------------------------------
                'Concatenate records
                JoinRecordsFields = Join$(ConcatenatedArray, P_RECORDS_DELIMITER)
                '@----------------------------------------------------------------------------
                'Free the memory
                Erase ConcatenatedArray
                Erase Buffer
        End Select
    Else
        '@----------------------------------------------------------------------------
        'Check for jagged arrays
        IsJagged = IsJaggedArray(RecordsArray)
        LB1 = LBound(RecordsArray, 1)
        UB1 = UBound(RecordsArray, 1)
        Select Case P_QUOTING_MODE
            Case QuotationMode.Critical
                If Not IsJagged Then
                    '@----------------------------------------------------------------------------
                    'Set array sizes
                    ReDim ConcatenatedArray(LB1 To UB1)
                    ReDim Buffer(0)
                    '@----------------------------------------------------------------------------
                    'Concatenate fields
                    For iLCounter = LB1 To UB1
                        If InStrB(1, RecordsArray(iLCounter), P_FIELDS_DELIMITER) Then
                            JoinBuffer(1) = RecordsArray(iLCounter)
                            Buffer(0) = Join$(JoinBuffer, vbNullString)
                        ElseIf InStrB(1, RecordsArray(iLCounter), CoerceChr) Then
                            JoinBuffer(1) = RecordsArray(iLCounter)
                            Buffer(0) = Join$(JoinBuffer, vbNullString)
                        ElseIf InStrB(1, RecordsArray(iLCounter), P_RECORDS_DELIMITER) Then
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
                    JoinRecordsFields = Join$(ConcatenatedArray, P_RECORDS_DELIMITER)
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
                        UB2 = UBound(tmpBuffer)
                        ReDim Buffer(LB2 To UB2)
                        For jLCounter = LB2 To UB2
                            If InStrB(1, tmpBuffer(jLCounter), P_FIELDS_DELIMITER) Then
                                JoinBuffer(1) = tmpBuffer(jLCounter)
                                Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                            ElseIf InStrB(1, tmpBuffer(jLCounter), CoerceChr) Then
                                JoinBuffer(1) = tmpBuffer(jLCounter)
                                Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                            ElseIf InStrB(1, tmpBuffer(jLCounter), P_RECORDS_DELIMITER) Then
                                JoinBuffer(1) = tmpBuffer(jLCounter)
                                Buffer(jLCounter) = Join$(JoinBuffer, vbNullString)
                            Else
                                Buffer(jLCounter) = tmpBuffer(jLCounter)
                            End If
                            JoinBuffer(1) = vbNullString
                        Next jLCounter
                        ConcatenatedArray(iLCounter) = Join$(Buffer, P_FIELDS_DELIMITER)
                    Next iLCounter
                    '@----------------------------------------------------------------------------
                    'Concatenate records
                    JoinRecordsFields = Join$(ConcatenatedArray, P_RECORDS_DELIMITER)
                    '@----------------------------------------------------------------------------
                    'Free the memory
                    Erase ConcatenatedArray
                    Erase Buffer
                End If
            Case Else
                If Not IsJagged Then
                    JoinRecordsFields = Join$(RecordsArray, P_RECORDS_DELIMITER)
                Else
                    '@----------------------------------------------------------------------------
                    'Set array sizes
                    ReDim ConcatenatedArray(LB1 To UB1)
                    '@----------------------------------------------------------------------------
                    'Concatenate fields
                    For iLCounter = LB1 To UB1
                        ConcatenatedArray(iLCounter) = Join$(RecordsArray(iLCounter), P_FIELDS_DELIMITER)
                    Next iLCounter
                    '@----------------------------------------------------------------------------
                    'Concatenate records
                    JoinRecordsFields = Join$(ConcatenatedArray, P_RECORDS_DELIMITER)
                    '@----------------------------------------------------------------------------
                    'Free the memory
                    Erase ConcatenatedArray
                    Erase Buffer
                End If
        End Select
    End If
    Exit Function
JoinRecordsFields_Error:
    JoinRecordsFields = vbNullString
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
    P_ERROR_SOURCE = Err.Source
End Function
Private Function MultiDimensional(CheckArray As Variant) As Boolean
Attribute MultiDimensional.VB_Description = "Checks if the specified array is multidimensional one."
    On Error GoTo ErrHandler_MultiDimensional

    If UBound(CheckArray, 2) > 0 Then
        MultiDimensional = True
    End If
    Exit Function
ErrHandler_MultiDimensional:
    MultiDimensional = False '1 dimension
End Function
Public Sub OpenConnection(csvPathAndFilename As String, Optional DeleExistingFile As Boolean = False)
Attribute OpenConnection.VB_Description = "Opens a CSV file on memory for binary access."
    On Error GoTo ErrHandler_OpenConnection
    If csvPathAndFilename <> vbNullString Then
        If DeleExistingFile Then 'Delete existing file
            If FileExists(csvPathAndFilename) Then Call Kill(csvPathAndFilename)
        End If
        '@----------------------------------------------------------------------------
        'open the file for binary access
        FileHandled = FreeFile
        Open csvPathAndFilename For Binary As #FileHandled
        '@----------------------------------------------------------------------------
        'Save values
        P_CONNECTED = True
        P_FILENAME = csvPathAndFilename
        P_FILEENCODING = GetFileEncoding
    Else
        P_CONNECTED = False
        P_ERROR_DESC = "[Open connection]: Isn't possible to link the object with the CSV file."
        P_ERROR_NUMBER = vbObjectError + 579
    End If
    Exit Sub
ErrHandler_OpenConnection:
    P_CONNECTED = False
    P_FILENAME = vbNullString
    P_ERROR_DESC = "[Open connection]: " & Err.Description
    P_ERROR_NUMBER = Err.Number
    P_ERROR_SOURCE = Err.Source
End Sub
Private Sub ParseCriticalCSV(ByRef csvText As String, _
                                    Optional HeadersOmission As Boolean = False)
Attribute ParseCriticalCSV.VB_Description = "Parses text strings CSV’s checking the RFC-4180 specs."
    Dim ASCIIcharw As Long
    Dim Buffer() As String
    Dim CurrentBufferIndex As Long
    Dim CurrentRecordSize As Long
    Dim CurrenttmpTokenIndex As Long
    Dim EndRecord As Long
    Dim EscapedWFS As Boolean
    Dim EscapeAscW As Long
    Dim EscapeChr As String
    Dim FDAscW As Long
    Dim Index As Long, MaxIndex As Long
    Dim ImportDepth As Long
    Dim LenCurrentIndex As Long
    Dim MaxBufferIndex As Long
    Dim MaxtmpTokenIndex As Long
    Dim OpenedToken As Boolean
    Dim OutputColumn As Long
    Dim OverStringPointer As Long, hpointer As Long
    Dim RecordEndReached As Boolean
    Dim RecordsCount As Long
    Dim SearchBeginningMark As Long
    Dim SplittedToken As Boolean
    Dim StartRecord As Long
    Dim tmpCSV() As String
    Dim tmpToken() As String
    Dim TokenBeginningPos As Long
    Dim TokenEndingPos As Long
    Dim TokenEndReached As Boolean
    
    On Error GoTo ParseCriticalCSV_Error_Handler
    '@----------------------------------------------------------------------------
    'Start variables
    SearchBeginningMark = 1&
    RecordsCount = 0&
    P_VECTORS_MAX_BOUND = 0&
    MaxBufferIndex = 127
    MaxtmpTokenIndex = 15
    ReDim Buffer(0 To MaxBufferIndex) 'Reserve space to dimension the OutputData array
    ReDim tmpToken(0 To MaxtmpTokenIndex) 'Reserve space to store splits fields
    CurrentBufferIndex = -1&
    CurrenttmpTokenIndex = -1&
    TokenEndReached = False
    RecordEndReached = False
    '@----------------------------------------------------------------------------
    'Set variables
    OverStringPointer = SearchBeginningMark
    ASCIIcharw = AscW(csvText)
    '@----------------------------------------------------------------------------
    'Split the content using the records delimiter char
    tmpCSV() = Split(csvText, P_RECORDS_DELIMITER)
    '@----------------------------------------------------------------------------
    'Set the pointers to the tmpCSV array
    Index = LBound(tmpCSV)
    MaxIndex = UBound(tmpCSV)
    '@----------------------------------------------------------------------------
    'Avoids to read empty records at the end of the file
    Do While LenB(tmpCSV(MaxIndex)) = 0
        MaxIndex = MaxIndex - 1
    Loop
    '@----------------------------------------------------------------------------
    'Check if user want Import all records
    EndRecord = P_ENDING_RECORD - 1&
    StartRecord = P_STARTING_RECORD - 1&
    If EndRecord > MaxIndex Then EndRecord = MaxIndex 'Fit to array size
    Select Case EndRecord
        Case 0&
            If StartRecord >= 0 Then
                ImportDepth = MaxIndex - StartRecord + 1&
            Else
                ImportDepth = 1&
            End If
        Case Is < StartRecord
            ImportDepth = 1&
        Case Else
            ImportDepth = EndRecord - StartRecord + 1&
    End Select
    If StartRecord > 0& Then
        Index = GetRecordIndex(StartRecord + 1&, tmpCSV, MaxIndex)
    ElseIf StartRecord = 0& Then
        If HeadersOmission And ImportDepth > 1& Then
            Index = GetRecordIndex(2&, tmpCSV, MaxIndex)
            ImportDepth = ImportDepth - 1&
        Else
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            SkipUnwantedLines Index, MaxIndex, tmpCSV
        End If
    End If
    If Index < 0& Then GoTo ParseCriticalCSV_StartingIndexOutOfBound
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case P_ESCAPE_TOKEN
        Case 1
            EscapeChr = CHR_APOSTROPHE
        Case Else
            EscapeChr = CHR_DOUBLE_QUOTES
    End Select
    EscapeAscW = AscW(EscapeChr)
    FDAscW = AscW(P_FIELDS_DELIMITER)
    '@----------------------------------------------------------------------------
    'Process the first record with Buffering
    Do While (Not RecordEndReached And Index <= MaxIndex)
        '@----------------------------------------------------------------------------
        'Process the current Token
        LenCurrentIndex = LenB(tmpCSV(Index))
        TokenEndReached = False
        TokenBeginningPos = OverStringPointer
        Select Case OverStringPointer
            Case Is < LenCurrentIndex
                ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                '@----------------------------------------------------------------------------
                'Ignore spaces and tabs
                Select Case P_ASCW_FIELDS_DELIMITER
                    Case Is <> 9
                        Do While ASCIIcharw = 32 Or ASCIIcharw = 9
                            OverStringPointer = OverStringPointer + 2&
                            If OverStringPointer > LenCurrentIndex Then Exit Do
                            ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                        Loop
                    Case Else
                        Do While ASCIIcharw = 32
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
            Do While (Not TokenEndReached And Index <= MaxIndex)
                '@----------------------------------------------------------------------------
                'Buffering if needed
                If Not SplittedToken Then SplittedToken = (hpointer = 0)
                Do While hpointer = 0& 'Escape char may be on subsequent Index
                    If CurrenttmpTokenIndex + 2& > MaxtmpTokenIndex Then
                        MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                        ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                    End If
                    Select Case SearchBeginningMark
                        Case 1 'Store the full Token from the current Index
                            CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                            tmpToken(CurrenttmpTokenIndex) = tmpCSV(Index)
                        Case Else 'Store the Right most Token portion from the current Index
                            CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                            tmpToken(CurrenttmpTokenIndex) = RightB$(tmpCSV(Index), LenCurrentIndex _
                                                            - SearchBeginningMark + 1)
                    End Select
                    tmpCSV(Index) = vbNullString
                    Index = Index + 1&
                    If Index <= MaxIndex Then LenCurrentIndex = LenB(tmpCSV(Index))
                    SearchBeginningMark = 1
                    hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                Loop
                If hpointer = 0& And Index > MaxIndex Then GoTo ParseCriticalCSV_MissingEscapeChar 'Missing escape char
                OverStringPointer = hpointer + 2&
                If OverStringPointer < LenCurrentIndex Then
                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                    '@----------------------------------------------------------------------------
                    'Ignore spaces and tabs
                    Select Case P_ASCW_FIELDS_DELIMITER
                        Case Is <> 9
                            Do While ASCIIcharw = 32 Or ASCIIcharw = 9
                                OverStringPointer = OverStringPointer + 2&
                                If OverStringPointer > LenCurrentIndex Then Exit Do
                                ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                            Loop
                        Case Else
                            Do While ASCIIcharw = 32
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
                    '@----------------------------------------------------------------------------
                    'Save the Token in Buffer
                    CurrentBufferIndex = CurrentBufferIndex + 1&
                    If CurrentBufferIndex > MaxBufferIndex Then
                        MaxBufferIndex = 2& * MaxBufferIndex + 1&
                    End If
                    Select Case SplittedToken
                        Case False 'Just save Token
                            Buffer(CurrentBufferIndex) = MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                TokenEndingPos - TokenBeginningPos + 1)
                        Case Else 'The Token is stored on tmpToken array
                            CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                            If CurrenttmpTokenIndex > MaxtmpTokenIndex Then
                                MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                                ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                            End If
                            tmpToken(CurrenttmpTokenIndex) = LeftB$(tmpCSV(Index), TokenEndingPos)
                            ReDim Preserve tmpToken(0 To CurrenttmpTokenIndex)
                            Buffer(CurrentBufferIndex) = Join$(tmpToken, P_RECORDS_DELIMITER) 'Save whole Token
                            CurrenttmpTokenIndex = -1&
                            ReDim tmpToken(0 To MaxtmpTokenIndex) 'Clear tmpToken array content
                    End Select
                    SplittedToken = False 'Restart Token status
                    OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                    RecordEndReached = False
                Else
                    If OverStringPointer >= LenCurrentIndex Then  'Escape char succeeded
                                                                    'by records separator char
                        TokenEndingPos = hpointer - 1&
                        TokenEndReached = True
                        RecordEndReached = True 'EOR
                        '@----------------------------------------------------------------------------
                        'Save the Token in Buffer
                        CurrentBufferIndex = CurrentBufferIndex + 1&
                        If CurrentBufferIndex > MaxBufferIndex Then
                            MaxBufferIndex = 2& * MaxBufferIndex + 1&
                        End If
                        Select Case SplittedToken
                            Case False 'Just save Token
                                Buffer(CurrentBufferIndex) = MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                TokenEndingPos - TokenBeginningPos + 1)
                            Case Else 'The Token is stored on tmpToken array
                                CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                                If CurrenttmpTokenIndex > MaxtmpTokenIndex Then
                                    MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                                    ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                                End If
                                tmpToken(CurrenttmpTokenIndex) = LeftB$(tmpCSV(Index), OverStringPointer - 3)
                                ReDim Preserve tmpToken(0 To CurrenttmpTokenIndex)
                                Buffer(CurrentBufferIndex) = Join$(tmpToken, P_RECORDS_DELIMITER) 'Save whole Token
                                CurrenttmpTokenIndex = -1&
                                ReDim tmpToken(0 To MaxtmpTokenIndex) 'Clear tmpToken array content
                        End Select
                        SplittedToken = False 'Restart Token status
                        tmpCSV(Index) = vbNullString
                        Index = Index + 1&
                        If Index <= MaxIndex Then LenCurrentIndex = LenB(tmpCSV(Index))
                        OverStringPointer = 1&
                    Else
                        If SplittedToken Then 'Store the Left most Token portion from the current Index
                            CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                            If CurrenttmpTokenIndex > MaxtmpTokenIndex Then
                                MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                                ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                            End If
                            tmpToken(CurrenttmpTokenIndex) = LeftB$(tmpCSV(Index), OverStringPointer - 1)
                        End If
                        SearchBeginningMark = OverStringPointer
                        hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                        TokenEndReached = False
                    End If
                End If
            Loop
        Else
            SplittedToken = False
            CurrentBufferIndex = CurrentBufferIndex + 1&
            If CurrentBufferIndex > MaxBufferIndex Then
                MaxBufferIndex = 2& * MaxBufferIndex + 1&
            End If
            '@----------------------------------------------------------------------------
            'Try to mark the Starting and Ending position for the field
            hpointer = InStrB(OverStringPointer, tmpCSV(Index), P_FIELDS_DELIMITER)
            If hpointer Then 'Found fields delimiter char
                TokenEndReached = True
                TokenEndingPos = hpointer - 1&
                OverStringPointer = hpointer + 2&
                RecordEndReached = False
                '@----------------------------------------------------------------------------
                'Save the Token in Buffer
                Buffer(CurrentBufferIndex) = MidB$(tmpCSV(Index), TokenBeginningPos, _
                                            TokenEndingPos - TokenBeginningPos + 1)
            Else 'Missing fields delimiter char
                TokenEndReached = True
                RecordEndReached = True
                '@----------------------------------------------------------------------------
                'Save the Token in Buffer
                Buffer(CurrentBufferIndex) = RightB$(tmpCSV(Index), LenCurrentIndex _
                                                - TokenBeginningPos + 1)
                tmpCSV(Index) = vbNullString
                Index = Index + 1&
                If Index <= MaxIndex Then LenCurrentIndex = LenB(tmpCSV(Index))
                OverStringPointer = 1&
            End If
        End If
    Loop
    '@----------------------------------------------------------------------------
    'Save the the first record to output array
    P_VECTORS_REGULAR_BOUND = CurrentBufferIndex
    ReDim Preserve Buffer(0 To P_VECTORS_REGULAR_BOUND)
    P_VECTORS_MAX_BOUND = P_VECTORS_REGULAR_BOUND
    CurrentRecordSize = P_VECTORS_MAX_BOUND
    Call CreateJagged(P_CSV_DATA, ImportDepth - 1&, P_VECTORS_REGULAR_BOUND)
    P_CSV_DATA(0) = Buffer()
    Erase Buffer
    RecordsCount = RecordsCount + 1&
    '@----------------------------------------------------------------------------
    'Skip lines if needed
     SkipUnwantedLines Index, MaxIndex, tmpCSV
     LenCurrentIndex = LenB(tmpCSV(Index))
    '@----------------------------------------------------------------------------
    'Loop through tmpCSV array
    Do While (Index <= MaxIndex And RecordsCount < ImportDepth)
        RecordEndReached = False 'Restart Record status
        OutputColumn = 0 'Restart Fields Pointer
        Do While (Not RecordEndReached And Index <= MaxIndex)
            '@----------------------------------------------------------------------------
            'Process the current Token
            TokenEndReached = False
            TokenBeginningPos = OverStringPointer
            Select Case OverStringPointer
                Case Is < LenCurrentIndex
                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                    '@----------------------------------------------------------------------------
                    'Ignore spaces and tabs
                    Select Case P_ASCW_FIELDS_DELIMITER
                        Case Is <> 9
                            Do While ASCIIcharw = 32 Or ASCIIcharw = 9
                                OverStringPointer = OverStringPointer + 2&
                                If OverStringPointer > LenCurrentIndex Then Exit Do
                                ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                            Loop
                        Case Else
                            Do While ASCIIcharw = 32
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
                Do While (Not TokenEndReached And Index <= MaxIndex)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not SplittedToken Then SplittedToken = (hpointer = 0)
                    Do While hpointer = 0& 'Escape char may be on subsequent Index
                        If CurrenttmpTokenIndex + 2& > MaxtmpTokenIndex Then
                            MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                            ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                        End If
                        Select Case SearchBeginningMark
                            Case 1 'Store the full Token from the current Index
                                CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                                tmpToken(CurrenttmpTokenIndex) = tmpCSV(Index)
                            Case Else 'Store the Right most Token portion from the current Index
                                CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                                tmpToken(CurrenttmpTokenIndex) = RightB$(tmpCSV(Index), _
                                                                LenCurrentIndex - SearchBeginningMark + 1)
                        End Select
                        tmpCSV(Index) = vbNullString
                        Index = Index + 1&
                        If Index <= MaxIndex Then LenCurrentIndex = LenB(tmpCSV(Index))
                        SearchBeginningMark = 1
                        hpointer = InStrB(SearchBeginningMark, tmpCSV(Index), EscapeChr) 'Find escape char
                    Loop
                    If hpointer = 0& And Index > MaxIndex Then GoTo ParseCriticalCSV_MissingEscapeChar 'Missing escape char
                    OverStringPointer = hpointer + 2&
                    If OverStringPointer < LenCurrentIndex Then
                        ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces and tabs
                        Select Case P_ASCW_FIELDS_DELIMITER
                            Case Is <> 9
                                Do While ASCIIcharw = 32 Or ASCIIcharw = 9
                                    OverStringPointer = OverStringPointer + 2&
                                    If OverStringPointer > LenCurrentIndex Then Exit Do
                                    ASCIIcharw = AscW(MidB$(tmpCSV(Index), OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While ASCIIcharw = 32
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
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        Select Case SplittedToken
                            Case False 'Just save Token
                                P_CSV_DATA(RecordsCount)(OutputColumn) = MidB$(tmpCSV(Index), _
                                                                                TokenBeginningPos, _
                                                                                TokenEndingPos - TokenBeginningPos + 1)
                            Case Else 'The Token is stored on tmpToken array
                                CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                                If CurrenttmpTokenIndex > MaxtmpTokenIndex Then
                                    MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                                    ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                                End If
                                tmpToken(CurrenttmpTokenIndex) = LeftB$(tmpCSV(Index), TokenEndingPos)
                                ReDim Preserve tmpToken(0 To CurrenttmpTokenIndex)
                                P_CSV_DATA(RecordsCount)(OutputColumn) = Join$(tmpToken, P_RECORDS_DELIMITER) 'Save whole Token
                                CurrenttmpTokenIndex = -1&
                                ReDim tmpToken(0 To MaxtmpTokenIndex) 'Clear tmpToken array content
                        End Select
                        SplittedToken = False 'Restart Token status
                        OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                        RecordEndReached = False
                    Else
                        If OverStringPointer >= LenCurrentIndex Then  'Escape char succeeded
                                                                      'by records separator char
                            TokenEndingPos = hpointer - 1&
                            TokenEndReached = True
                            RecordEndReached = True 'EOR
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case SplittedToken
                                Case False 'Just save Token
                                    P_CSV_DATA(RecordsCount)(OutputColumn) = MidB$(tmpCSV(Index), _
                                                                                TokenBeginningPos, _
                                                                                TokenEndingPos - TokenBeginningPos + 1)
                                Case Else 'The Token is stored on tmpToken array
                                    CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                                    If CurrenttmpTokenIndex > MaxtmpTokenIndex Then
                                        MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                                        ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                                    End If
                                    tmpToken(CurrenttmpTokenIndex) = LeftB$(tmpCSV(Index), OverStringPointer - 3)
                                    ReDim Preserve tmpToken(0 To CurrenttmpTokenIndex)
                                    P_CSV_DATA(RecordsCount)(OutputColumn) = Join$(tmpToken, P_RECORDS_DELIMITER) 'Save whole Token
                                    CurrenttmpTokenIndex = -1&
                                    ReDim tmpToken(0 To MaxtmpTokenIndex) 'Clear tmpToken array content
                            End Select
                            SplittedToken = False 'Restart Token status
                            tmpCSV(Index) = vbNullString
                            Index = Index + 1&
                            OverStringPointer = 1&
                        Else
                            If SplittedToken Then 'Store the Left most Token portion from the current Index
                                CurrenttmpTokenIndex = CurrenttmpTokenIndex + 1&
                                If CurrenttmpTokenIndex > MaxtmpTokenIndex Then
                                    MaxtmpTokenIndex = 2& * MaxtmpTokenIndex + 1&
                                    ReDim Preserve tmpToken(0 To MaxtmpTokenIndex)
                                End If
                                tmpToken(CurrenttmpTokenIndex) = LeftB$(tmpCSV(Index), OverStringPointer - 1)
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
                hpointer = InStrB(OverStringPointer, tmpCSV(Index), P_FIELDS_DELIMITER)
                If hpointer Then 'Found fields delimiter char
                    TokenEndReached = True
                    TokenEndingPos = hpointer - 1&
                    OverStringPointer = hpointer + 2&
                    RecordEndReached = False
                    '@----------------------------------------------------------------------------
                    'Save the Token
                    P_CSV_DATA(RecordsCount)(OutputColumn) = MidB$(tmpCSV(Index), TokenBeginningPos, _
                                                                    TokenEndingPos - TokenBeginningPos + 1)
                Else 'Missing fields delimiter char
                    TokenEndReached = True
                    RecordEndReached = True
                    '@----------------------------------------------------------------------------
                    'Save the Token
                    P_CSV_DATA(RecordsCount)(OutputColumn) = RightB$(tmpCSV(Index), LenCurrentIndex _
                                                                - TokenBeginningPos + 1)
                    tmpCSV(Index) = vbNullString
                    Index = Index + 1&
                    OverStringPointer = 1&
                End If
            End If
            OutputColumn = OutputColumn + 1&
            If OutputColumn > CurrentRecordSize Then
                If Not RecordEndReached Then
                    '@----------------------------------------------------------------------------
                    'Reserve space for store at least another field
                    P_VARYING_LENGTHS = True
                    CurrentRecordSize = CurrentRecordSize + 1&
                    If CurrentRecordSize > P_VECTORS_MAX_BOUND Then
                        P_VECTORS_MAX_BOUND = CurrentRecordSize
                    End If
                    Buffer() = P_CSV_DATA(RecordsCount)
                    ReDim Preserve Buffer(0 To CurrentRecordSize)
                    P_CSV_DATA(RecordsCount) = Buffer
                    Erase Buffer
                    VectorsBoundData(0) = RecordsCount
                    VectorsBoundData(1) = CurrentRecordSize
                    Select Case P_IRREGULAR_VECTORS_INFO.Count
                        Case 0 'First Item
                            P_IRREGULAR_VECTORS_INFO.Add VectorsBoundData
                        Case Else 'Check the last worked Item
                            If RecordsCount = P_IRREGULAR_VECTORS_INFO(P_IRREGULAR_VECTORS_INFO.Count)(0) Then 'Update INFO
                                P_IRREGULAR_VECTORS_INFO.Remove P_IRREGULAR_VECTORS_INFO.Count
                                P_IRREGULAR_VECTORS_INFO.Add VectorsBoundData
                            Else 'Add Item
                                P_IRREGULAR_VECTORS_INFO.Add VectorsBoundData
                            End If
                    End Select
                End If
            End If
        Loop
        CurrentRecordSize = P_VECTORS_REGULAR_BOUND 'Restart jagged size
        RecordsCount = RecordsCount + 1&
        '@----------------------------------------------------------------------------
        'Skip lines if needed
        SkipUnwantedLines Index, MaxIndex, tmpCSV
        If Index <= MaxIndex Then LenCurrentIndex = LenB(tmpCSV(Index))
    Loop
    P_ITEMS = RecordsCount
    '@----------------------------------------------------------------------------
    'Check if the array need to be shrinked at the output stage
    If ImportDepth <> RecordsCount Then
        ReDim Preserve P_CSV_DATA(0 To P_ITEMS - 1&)
    End If
    Erase tmpCSV
    Exit Sub
ParseCriticalCSV_Error_Handler:
    Erase P_CSV_DATA
    Erase tmpCSV
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
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
                    Description:="Missing some escape char at line " & Index & ". Check the data and try again."
    Resume ParseCriticalCSV_Error_Handler
End Sub
Private Sub ParseCSV(ByRef csvText As String, _
                        Optional HeadersOmission As Boolean = False)
Attribute ParseCSV.VB_Description = "Parses text strings CSV’s without check RFC-4180 specs."
    Dim ArrayLB As Long
    Dim ArrayUB As Long
    Dim csvArray() As String
    Dim hPos As Long
    Dim brCounter As Long
    Dim lColumns As Long
    Dim lCounter As Long
    Dim lngLB As Long
    Dim lngUB As Long
    Dim lRows As Long
    Dim mCounter As Long
    Dim NextPos As Long
    Dim PrevPos As Long
    Dim Pointer As Long
    
    On Error GoTo ErrHandler_ParseCS
    '@----------------------------------------------------------------------------
    'Set Coerce Type to records and fields
    Select Case P_ESCAPE_TOKEN
        Case 0
            DataDelimiter = P_FIELDS_DELIMITER
            LinesDelimiter = P_RECORDS_DELIMITER
        Case 1
            DataDelimiter = CHR_APOSTROPHE & P_FIELDS_DELIMITER & CHR_APOSTROPHE
            LinesDelimiter = CHR_APOSTROPHE & P_RECORDS_DELIMITER & CHR_APOSTROPHE
        Case Else
            DataDelimiter = CHR_DOUBLE_QUOTES & P_FIELDS_DELIMITER & CHR_DOUBLE_QUOTES
            LinesDelimiter = CHR_DOUBLE_QUOTES & P_RECORDS_DELIMITER & CHR_DOUBLE_QUOTES
    End Select
    '@----------------------------------------------------------------------------
    'Split each line into an array entry
    csvArray() = Split(csvText, LinesDelimiter)
    '@----------------------------------------------------------------------------
    'Avoid to read empty records
    ArrayLB = CLng(LBound(csvArray))
    ArrayUB = CLng(UBound(csvArray))
    brCounter = 0
    Do While LenB(csvArray(ArrayUB - brCounter)) = 0
        brCounter = brCounter + 1
    Loop
    ArrayUB = ArrayUB - brCounter
    '@----------------------------------------------------------------------------
    'Ommit the coerce char from the upper and lower elements
    If P_ESCAPE_TOKEN <> 0 Then
        csvArray(ArrayLB) = MidB$(csvArray(ArrayLB), 3)
        If RightB$(csvArray(ArrayUB), 4) = vbCrLf Then
            csvArray(ArrayUB) = LeftB$(csvArray(ArrayUB), _
                                LenB(csvArray(ArrayUB)) - 6)
        Else
            csvArray(ArrayUB) = LeftB$(csvArray(ArrayUB), _
                                LenB(csvArray(ArrayUB)) - 2)
        End If
    End If
    '@----------------------------------------------------------------------------
    'Check if user want Import all records
    Select Case P_STARTING_RECORD - 1&
        Case 0&
            lngLB = ArrayLB
        Case Else
            lngLB = P_STARTING_RECORD - 1&
    End Select
    Select Case P_ENDING_RECORD - 1&
        Case 0&
            lngUB = ArrayUB
        Case Is <= lngLB
            lngUB = lngLB
        Case Else
            lngUB = P_ENDING_RECORD - 1&
    End Select
    '@----------------------------------------------------------------------------
    'Take care of headers
    If lngLB = ArrayLB Then
        If lngUB - lngLB > 0 Then
            If Not HeadersOmission Then
                hPos = 0
            Else
                hPos = 1
            End If
        End If
    End If
    lRows = lngUB - lngLB
    P_ITEMS = lRows - hPos + 1&
    lColumns = CSVcolumns(csvArray, DataDelimiter) - 1
    lCounter = lngLB + hPos
    Call CreateJagged(P_CSV_DATA, P_ITEMS - 1&, lColumns)
    P_VECTORS_REGULAR_BOUND = lColumns
    P_VECTORS_MAX_BOUND = P_VECTORS_REGULAR_BOUND
    '@----------------------------------------------------------------------------
    'Process the data
    Pointer = hPos
    Do While lCounter <= lngUB
        mCounter = 0
        PrevPos = 1
        NextPos = CLng(InStrB(PrevPos, csvArray(lCounter), _
                            DataDelimiter))
        Do While NextPos
            P_CSV_DATA(Pointer - hPos)(mCounter) = MidB$(csvArray(lCounter), _
                                                        PrevPos, NextPos - PrevPos)
            PrevPos = NextPos + CLng(LenB(DataDelimiter))
            mCounter = mCounter + 1
            NextPos = CLng(InStrB(PrevPos, csvArray(lCounter), DataDelimiter))
        Loop
        P_CSV_DATA(Pointer - hPos)(mCounter) = MidB$(csvArray(lCounter), PrevPos)
        lCounter = lCounter + 1
        Pointer = Pointer + 1
    Loop
    '@----------------------------------------------------------------------------
    'Free the memory
    Erase csvArray
    Exit Sub
ErrHandler_ParseCS:
    Erase P_CSV_DATA
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
    P_ERROR_SOURCE = Err.Source
End Sub
Private Function ReturnCharset(Optional verifyANSI As Boolean = True) As abCharsets
Attribute ReturnCharset.VB_Description = "Determines the charset for the CSV file."
    Const bytByte0Unicode_c As Byte = 255
    Const bytByte1Unicode_c As Byte = 254
    Const bytByte0UnicodeBigEndian_c As Byte = 254
    Const bytByte1UnicodeBigEndian_c As Byte = 255
    Const bytByte0UTF8_c As Byte = 239
    Const bytByte1UTF8_c As Byte = 187
    Const bytByte2UTF8_c As Byte = 191
    Const lngByte0 As Long = 0
    Const lngByte1 As Long = 1
    Const lngByte2 As Long = 2
    
    Dim bytHeader() As Byte
    Dim eRtnVal As abCharsets
    On Error GoTo Err_Hnd
    bytHeader() = GetFileBytes(lngByte2)
    Select Case bytHeader(lngByte0)
        Case bytByte0Unicode_c
            If bytHeader(lngByte1) = bytByte1Unicode_c Then
                eRtnVal = abCharsets.abUnicode
            End If
        Case bytByte0UnicodeBigEndian_c
            If bytHeader(lngByte1) = bytByte1UnicodeBigEndian_c Then
                eRtnVal = abCharsets.abUnicodeBigEndian
            End If
        Case bytByte0UTF8_c
            If bytHeader(lngByte1) = bytByte1UTF8_c Then
                If bytHeader(lngByte2) = bytByte2UTF8_c Then
                    eRtnVal = abCharsets.abUTF8
                End If
            End If
    End Select
    If Not CBool(eRtnVal) Then
        If verifyANSI Then
            If IsANSI Then
                eRtnVal = abCharsets.abANSI
            Else
                eRtnVal = abCharsets.ebUnknown
            End If
        Else
            eRtnVal = abCharsets.abANSI
        End If
    End If
Exit_Proc:
    On Error Resume Next
    Erase bytHeader
    ReturnCharset = eRtnVal
    Exit Function
Err_Hnd:
    eRtnVal = abCharsets.abError
    P_ERROR_NUMBER = Err.Number
    P_ERROR_DESC = Err.Description
    P_ERROR_SOURCE = Err.Source
    Resume Exit_Proc
End Function
Public Sub ResetToDefault()
Attribute ResetToDefault.VB_Description = "Resets all the options to its default value."
    P_COMMENTSTOKEN = AscW("#")
    P_CONNECTED = False
    P_ENDING_RECORD = 1&
    P_ERROR_DESC = vbNullString
    P_ERROR_NUMBER = 0&
    P_ERROR_SOURCE = vbNullString
    P_ESCAPE_TOKEN = EscapeTokens.DoubleQuotes
    P_FIELDS_DELIMITER = CHR_COMMA
    P_ASCW_FIELDS_DELIMITER = AscW(P_FIELDS_DELIMITER)
    P_FILENAME = vbNullString
    P_QUOTING_MODE = QuotationMode.Critical
    P_RECORDS_DELIMITER = vbCrLf
    P_STARTING_RECORD = 1&
    P_SUCCESSFUL_EXPORT = False
    P_SUCCESSFUL_IMPORT = False
    P_VARYING_LENGTHS = False
    Set P_IRREGULAR_VECTORS_INFO = New Collection
End Sub
Private Sub SkipUnwantedLines(ByRef idx As Long, ByRef MaxIdx As Long, _
                                ByRef Arr() As String)
Attribute SkipUnwantedLines.VB_Description = "Ignores empty, blanks and commented lines."
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
                If CharCode = P_COMMENTSTOKEN Then 'Commented line found
                    UnwantedLine = True
                Else
                    '@----------------------------------------------
                    'Ignore non printable chars at the string start
                    Select Case P_ASCW_FIELDS_DELIMITER
                        Case Is <> 9
                            Do While CharCode >= 0& And CharCode <= 32& And StrPointer < CurLength
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    Exit Do
                                End If
                                CharCode = AscW(MidB$(Arr(idx), StrPointer, 2))
                            Loop
                            If CharCode = P_COMMENTSTOKEN Then UnwantedLine = True
                        Case Else
                            Do While (CharCode >= 0& And CharCode <= 32&) And (StrPointer < CurLength And CharCode <> 9)
                                StrPointer = StrPointer + 2&
                                If StrPointer > CurLength Then 'Non printable line
                                    UnwantedLine = True
                                    Exit Do
                                End If
                                CharCode = AscW(MidB$(Arr(idx), StrPointer, 2))
                            Loop
                            If CharCode = P_COMMENTSTOKEN Then UnwantedLine = True
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
Public Sub TwoDimToJaggedArray(ByRef TwoDimArray() As String, ByRef JaggedArray() As Variant)
Attribute TwoDimToJaggedArray.VB_Description = "Deconstructs a 2D string array and puts its content into a jagged array."
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
Private Sub Class_Initialize()
    ResetToDefault
End Sub

Private Sub Class_Terminate()
    Erase P_CSV_DATA
End Sub
