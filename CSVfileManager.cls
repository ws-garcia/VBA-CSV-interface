VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVfileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2020 W. García
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' This class is inspired in a work available in
' https://www.freevbcode.com/ShowCode.asp?ID=3110
' The class module was designed and tested using
' Windows 7® and is supose to work as well over
' more recent Microsoft™ operative system.
'#
' Use CSVfileManager class to simplify the work
' with comma separated value (CSV) files. It allow
' you to exchange information betwen an VBA array and
' an external CSV file without using Excel WorkSheets,
' neigter any external reference such as [MS Scripting
' Runtime]. The class develop is focusing in the code
' execution performance. Keep in mind that the arrays
' returned are zero(0) based.
'#
' You are free to visit the cited web pages to explore
' some of the available solutions for work with CSV file and
' an VBA Array.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' USAGE:
' #1- Create an instance of the class (Dim CSVmanager As CSVfileManager)
' #2- Connect to CSV [CSVmanager.OpenConnection] using the full path to the target file.
' #3- Call the desired method
'                               CSVmanager.ImportFromCSV
'                               CSVmanager.ExportToCSV
' #4- Read the VBA array using [CSVmanager.CSVdata] property
' NOTE: you can modify the default value for the properties [FieldsDelimiter], [CoercType],
' [StartingRecord] and [EndingRecord] to alter the way the class works.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
' CREDITS:
' Although this class have original pieces of code and logic, it use code,
' or hints, of many authors. Visit the web sources listed below:
' @1 https://www.codegrepper.com/code-examples/vb/excel+vba+Load+csv+file+into+an+array+rather+than+the+worksheet
' @2 http://excellerando.blogspot.com/2014/12/writing-excel-range-to-csv-file.html
' @3 https://stackoverflow.com/questions/4191560/create-csv-from-array-in-vba
' @4 https://senipah.github.io/VBA-Better-Array
' @5 https://www.freevbcode.com/ShowCode.asp?ID=3110
' @6 https://stackoverflow.com/questions/1376756/what-is-a-superfast-way-to-read-large-files-line-by-line-in-vba
' @7 https://github.com/sdkn104/VBA-CSV
' @8 https://www.freevbcode.com/ShowCode.asp?ID=7655
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
Private Const CHR_SPACE As String = " "
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private P_COERCE_TYPE As CoerceTypes '-----------Holds the char used for coerce the fields
Private P_CONNECTED As Boolean '-----------------Holds the connection state
Private P_CSV_DATA As Variant '------------------Holds the CSV data for current instance
Private P_ENDING_RECORD As Long '----------------Pointer to the last record to be imported
Private P_ERROR_DESC As String '-----------------Holds the Error Description
Private P_ERROR_NUMBER As Long '-----------------Holds the Error number
Private P_ERROR_SOURCE As String '---------------Holds the Error source
Private P_FIELDS_DELIMITER As String '-----------Holds the current delimiter for CSV fields
Private P_FILENAME As String '-------------------Holds the actual file name
Private P_RECORDS_DELIMITER As String '----------Holds the records delimiter (CR, LF or CRLF)
Private P_STARTING_RECORD As Long '--------------Pointer to the first record to be imported
Private P_SUCCESSFUL_EXPORT As Boolean
Private P_SUCCESSFUL_IMPORT As Boolean
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum CoerceTypes
    DoubleQuotes = 2
    NotCoerce = 0
    Apostrophe = 1
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
' @Common
Private DataDelimiter As String '------Concatenation of Fields separator and CoerceType
Private LinesDelimiter As String '-----Concatenation of Records separator and CoerceType
Private FileContent As String '--------Holds the whole CSV file into this string
Private FileHandled As Integer '-------Pointer to the actual CSV file
'////////////////////////////////////////////////////////////////////////////////////////////
'#

'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
Public Property Get CoerceType() As CoerceTypes
    CoerceType = P_COERCE_TYPE
End Property
Public Property Let CoerceType(ByVal Coerce As CoerceTypes)
    P_COERCE_TYPE = Coerce
End Property
Public Property Get Connected() As Boolean
    Connected = P_CONNECTED
End Property
Public Property Get CSVdata() As Variant
    CSVdata = P_CSV_DATA
End Property
Public Property Get EndingRecord() As Long
    EndingRecord = P_ENDING_RECORD
End Property
Public Property Let EndingRecord(ByVal RecNumber As Long)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' Allow user decide the record over which the import process will ends. If the P_ENDING_RECORD
' property is set to zero(0), all the records, starting on P_STARTING_RECORD, will be imported.
'#
' ACCES TYPE: R/W
'#
' CONDITIONS: RecNumber >= 1
'////////////////////////////////////////////////////////////////////////////////////////////
    Select Case RecNumber
        Case Is >= 1&
            P_ENDING_RECORD = RecNumber - 1
        Case Else
        P_ENDING_RECORD = 0& 'Set the pointer to the last available record in the CSV file
    End Select
End Property
Public Property Get ErrDescription() As String
    ErrDescription = P_ERROR_DESC
End Property
Public Property Get ErrNumber() As Long
    ErrNumber = P_ERROR_NUMBER
End Property
Public Property Get ErrSource() As String
    ErrSource = P_ERROR_SOURCE
End Property
Public Property Get ExportSuccess() As Boolean
    ExportSuccess = P_SUCCESSFUL_EXPORT
End Property
Public Property Get FieldsDelimiter() As String
    FieldsDelimiter = P_FIELDS_DELIMITER
End Property
Public Property Let FieldsDelimiter(ByVal Delimiter As String)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' Allow user decide the char that will be used for delimit fields in the read and write CSV
' file operations from current instance.
'#
' ACCES TYPE: R/W
'#
' CONDITIONS: Delimiter = [","][";"][vbTab][space]
'////////////////////////////////////////////////////////////////////////////////////////////
    Select Case Delimiter
        Case CHR_COMMA, CHR_SEMICOLON, vbTab, CHR_SPACE
            P_FIELDS_DELIMITER = Delimiter
        Case Else
            P_FIELDS_DELIMITER = CHR_COMMA
    End Select
End Property
Public Property Get fileName() As String
    fileName = P_FILENAME
End Property
Public Property Get ImportSuccess() As Boolean
    ImportSuccess = P_SUCCESSFUL_IMPORT
End Property
Public Property Let RecordsDelimiter(Delimiter As String)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' Allow user decide the char that will be used for delimit records in the read and write CSV
' file operations from current instance.
'#
' ACCES TYPE: R/W
'#
' CONDITIONS: Delimiter = [vbCrLf][vbCr][vbLf]
'////////////////////////////////////////////////////////////////////////////////////////////
    Select Case Delimiter
        Case vbCrLf, vbCr, vbLf
            P_RECORDS_DELIMITER = Delimiter
        Case Else
            P_RECORDS_DELIMITER = vbCrLf
    End Select
End Property
Public Property Get StartingRecord() As Long
    StartingRecord = P_STARTING_RECORD
End Property
Public Property Let StartingRecord(ByVal RecNumber As Long)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' Allow user decide the record over which the import process will starts.
'#
' ACCES TYPE: R/W
'#
' CONDITIONS: RecNumber >= 1
'////////////////////////////////////////////////////////////////////////////////////////////
    Select Case RecNumber
        Case Is >= 1&
            P_STARTING_RECORD = RecNumber - 1
        Case Else
            P_STARTING_RECORD = 0& 'Set the pointer to the first available record in the CSV file
    End Select
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#

Private Sub CloseConnection()
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method close the current opened CSV file. The Private scope helps to avoid error.
'////////////////////////////////////////////////////////////////////////////////////////////
    On Error GoTo ErrHandler_CloseConnection
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

Private Function CSVcolumns(csvArray() As String, _
                            Optional FieldsDelimiter As String = CHR_COMMA) As Long
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method returns a long that indicates the max number (BASE 1) of columns for given
' CSVArray. The function requires an one dimentional array with all the CSV records on it.
' Its returns -1 when an error occurs.
'////////////////////////////////////////////////////////////////////////////////////////////
    ' @Compute columns procedure
    Dim iCounter As Long
    Dim LB As Long, UB As Long '-------Arraybounds
    Dim lngPos As Long
    Dim MDCounter As Long '------------Store csvDelimiter matches
    Dim NumOfDelimiter As Long
    
    On Error GoTo ErrHandler_CSVcolumns
    LB = LBound(csvArray)
    UB = UBound(csvArray)
    NumOfDelimiter = 0
    MDCounter = 0
    'lngPos = CLng(InStr(1, csvArray(LB), FieldsDelimiter))
    lngPos = CLng(InStrB(csvArray(LB), FieldsDelimiter))
    Do While lngPos
        MDCounter = MDCounter + 1
        'lngPos = CLng(InStr(lngPos + 1, csvArray(LB), FieldsDelimiter))
        lngPos = CLng(InStrB(lngPos + 2, csvArray(LB), FieldsDelimiter))
    Loop
    NumOfDelimiter = MDCounter
    CSVcolumns = NumOfDelimiter + 1
    Exit Function
ErrHandler_CSVcolumns:
    CSVcolumns = -1
End Function

Public Sub ExportToCSV(csvArray As Variant)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method export a given VBA array to the location and with the name used in the
' OpenConnection method call. In addition, it uses the CHR_COMMA, and the
' P_COERCE_TYPE that the user's previously defined.
'////////////////////////////////////////////////////////////////////////////////////////////
    On Error GoTo ErrHandler_ExportToCSV
    DoEvents
    If P_CONNECTED Then
        '@----------------------------------------------------------------------------
        'Join all the records
        FileContent = JoinRecordsFields(csvArray)
        '@----------------------------------------------------------------------------
        'open the file for binary access
        'Write the data on the current opened file
        Put #FileHandled, , FileContent
        Call CloseConnection
        '@----------------------------------------------------------------------------
        P_SUCCESSFUL_EXPORT = True
        P_CSV_DATA = csvArray
        P_ERROR_DESC = vbNullString
        P_ERROR_NUMBER = 0&
        P_ERROR_SOURCE = vbNullString
    Else
        P_SUCCESSFUL_EXPORT = False
        P_ERROR_DESC = "[CSV file Export]: the object is not linked to a CSV file. " _
                        & "Ensure use the OpenConnection method before export data."
        P_ERROR_NUMBER = vbObjectError + 9009
        P_ERROR_SOURCE = "CSVfileManager"
    End If
    Exit Sub
ErrHandler_ExportToCSV:
    Call CloseConnection
    P_SUCCESSFUL_EXPORT = False
    P_ERROR_DESC = "[CSV file Export]: " & Err.Description
    P_ERROR_NUMBER = Err.Number
    P_ERROR_SOURCE = Err.Source
End Sub
Public Sub GetCSVtext(ByRef OutPutVar As String)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method read CSV content and pass it to OutPutVar variable. This is not elegant but,
' the purpose is the avoid execution speed losses. To use this method, first, you need to
' open a connection ussing the OpenConnection method, to point to the CSV file.
'////////////////////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
    OutPutVar = Space$(LOF(FileHandled))
    Get #FileHandled, , OutPutVar
    Call CloseConnection
End Sub
Public Function ImportFromCSV(Optional HeadersOmission As Boolean = False) As CSVfileManager
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method import CSV content to a VBA array. The file to be imported has the name used
' in the OpenConnection method call. In addition, it uses the CHR_COMMA, and the
' P_COERCE_TYPE that the user's previously defined. The optional parameter HeadersOmission
' allows user avoid importation of fields titles (1st line).
' The method returns Null when an error occurs.
'////////////////////////////////////////////////////////////////////////////////////////////
    ' @Import procedure
    
    On Error GoTo ErrHandler_ImportFromCSV
    DoEvents
    If P_CONNECTED Then
        '@--------------------------------------------------------------------------------
        'Read whole file content
        GetCSVtext FileContent
        '@--------------------------------------------------------------------------------
        'Try to load the data
        P_CSV_DATA = ParseCSV(FileContent, HeadersOmission)
        '@----------------------------------------------------------------------------
        'Free the memory
        FileContent = vbNullString
        '@----------------------------------------------------------------------------
        'Save results
        Set ImportFromCSV = Me
        P_SUCCESSFUL_IMPORT = True
        P_ERROR_DESC = vbNullString
        P_ERROR_NUMBER = 0&
        P_ERROR_SOURCE = vbNullString
    Else
        P_SUCCESSFUL_IMPORT = False
        P_ERROR_DESC = "[CSV file Export]: the object is not linked to a CSV file. " _
                        & "Ensure use the OpenConnection method before import data."
        P_ERROR_NUMBER = vbObjectError + 9007
        P_ERROR_SOURCE = "CSVfileManager"
    End If
    Exit Function
ErrHandler_ImportFromCSV:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[CSV file Import]: " & Err.Description
    P_ERROR_NUMBER = Err.Number
    P_ERROR_SOURCE = Err.Source
End Function

Private Function JoinRecordsFields(RecordsArray As Variant) As String
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method returns a string that contains the concatenation of each records on the
' RecordsArray variable. All the fields are joined using the especified CHR_COMMA
' and P_COERCE_TYPE variables.
'#
' INPUT:
'           RecordsArray {A = [|a| |b| |c| : |d| |e| |f|]}
'#
' OUTPUTS:
' @comma as Deimiter:
'       [a[P_COERCE_TYPE],[P_COERCE_TYPE]b[P_COERCE_TYPE],[P_COERCE_TYPE]c [P_RECORDS_DELIMITER]
'       d[P_COERCE_TYPE],[P_COERCE_TYPE]e[P_COERCE_TYPE],[P_COERCE_TYPE]f]
' @SemiColom as Deimiter:
'       [a[P_COERCE_TYPE];[P_COERCE_TYPE]b[P_COERCE_TYPE];[P_COERCE_TYPE]c [P_RECORDS_DELIMITER]
'       d[P_COERCE_TYPE];[P_COERCE_TYPE]e[P_COERCE_TYPE];[P_COERCE_TYPE]f]
'////////////////////////////////////////////////////////////////////////////////////////////
    ' @Join records fields proceure
    Dim Buffer() As String
    Dim CoerceChr As String
    Dim ConcatenatedArray() As String
    Dim iLCounter As Long, jLCounter As Long
    Dim JoinBuffer(0 To 2) As String
    Dim LB1 As Long, UB1 As Long
    Dim LB2 As Long, ub2 As Long

    On Error GoTo JoinRecordsFields_Error
    If MultiDimensional(RecordsArray) Then
        LB1 = LBound(RecordsArray, 1)
        LB2 = LBound(RecordsArray, 2)
        UB1 = UBound(RecordsArray, 1)
        ub2 = UBound(RecordsArray, 2)
        '@----------------------------------------------------------------------------
        'Set Data and coerce Delimiter
        Select Case P_COERCE_TYPE
            Case 0 'Fields delimited with P_FIELDS_DELIMITER
                DataDelimiter = P_FIELDS_DELIMITER
                CoerceChr = vbNullString
            Case 1 'Fields delimited with Apostrophe
                DataDelimiter = CHR_APOSTROPHE & P_FIELDS_DELIMITER & CHR_APOSTROPHE
                CoerceChr = CHR_APOSTROPHE
            Case Else  'Fields delimited with Double Quotes
                DataDelimiter = CHR_DOUBLE_QUOTES & P_FIELDS_DELIMITER & CHR_DOUBLE_QUOTES
                CoerceChr = CHR_DOUBLE_QUOTES
        End Select
        '@----------------------------------------------------------------------------
        'Set the char buffer
        JoinBuffer(0) = CoerceChr
        JoinBuffer(1) = vbNullString 'Store the value to be coerced here
        JoinBuffer(2) = CoerceChr
        '@----------------------------------------------------------------------------
        'Set array sizes
        ReDim ConcatenatedArray(LB1 To UB1) 'Resize the output array
        ReDim Buffer(LB2 To ub2)
        '@----------------------------------------------------------------------------
        'Concatenate fields
        For iLCounter = LB1 To UB1 'Loop trhrought all rows
            For jLCounter = LB2 To ub2 'Loop trhrought all columns
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
    Else
        JoinRecordsFields = Join$(RecordsArray, P_RECORDS_DELIMITER)
    End If
    Exit Function
JoinRecordsFields_Error:
    JoinRecordsFields = vbNullString
End Function

Private Function MultiDimensional(CheckArray As Variant) As Boolean
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This function checks if the given array is a multidimensional one.
'////////////////////////////////////////////////////////////////////////////////////////////
    On Error GoTo ErrHandler_MultiDimensional

    If UBound(CheckArray, 2) > 0 Then
        MultiDimensional = True 'more than 1 dimension
    End If
    Exit Function
ErrHandler_MultiDimensional:
    MultiDimensional = False '1 dimension
End Function

Public Sub OpenConnection(csvPathAndFilename As String)
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This method open a CSV file in binary mode, for data Input and Output. If the given path
' exists then the file will be created on that path, otherwise an error occur.
'////////////////////////////////////////////////////////////////////////////////////////////
    On Error GoTo ErrHandler_OpenConnection
    '@----------------------------------------------------------------------------
    'Validate file name and extension
    If csvPathAndFilename <> vbNullString Then
        If CLng(InStrB(csvPathAndFilename, ".csv")) = 0 Then
            csvPathAndFilename = csvPathAndFilename & ".csv"
        Else
            'While (Len(csvPathAndFilename) - _
                    CLng(InStr(csvPathAndFilename, ".csv"))) > 3
            While (LenB(csvPathAndFilename) - _
                    CLng(InStrB(csvPathAndFilename, ".csv"))) > 7
                'csvPathAndFilename = Left$(csvPathAndFilename, Len(csvPathAndFilename) - 1)
                csvPathAndFilename = LeftB$(csvPathAndFilename, LenB(csvPathAndFilename) - 2)
            Wend
        End If
        '@----------------------------------------------------------------------------
        'open the file for binary access
        FileHandled = FreeFile
        Open csvPathAndFilename For Binary As #FileHandled
        '@----------------------------------------------------------------------------
        'Save values
        P_CONNECTED = True
        P_FILENAME = csvPathAndFilename
        P_ERROR_DESC = vbNullString
        P_ERROR_NUMBER = 0&
        P_ERROR_SOURCE = vbNullString
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

Public Function ParseCSV(ByRef csvText As String, _
                        Optional HeadersOmission As Boolean = False) As Variant
'////////////////////////////////////////////////////////////////////////////////////////////
' DESCRIPTION:
' This process a CSV text content and pass it to a VBA array. The optional parameter,
' HeadersOmission, allows user avoid importation of fields titles (1st line).
' The method returns Null when an error occurs.
'////////////////////////////////////////////////////////////////////////////////////////////
    Dim ArrayLB As Long
    Dim ArrayUB As Long
    Dim csvArray() As String
    Dim hPos As Long '-----------------Make the function able to skip the header
    Dim brCounter As Long
    Dim lColumns As Long
    Dim lCounter As Long
    Dim lngLB As Long
    Dim lngUB As Long
    Dim lRows As Long
    Dim mCounter As Long
    Dim NextPos As Long
    Dim OutPutArray() As String
    Dim ParseOption As CoerceTypes '---For distinguish between quotes used on records fields
    Dim PrevPos As Long
    Dim Pointer As Long '--------------Point to the current OutPutArray() element
    Dim sSplit() As String
    Dim tmpDelimiter As String
    
    On Error GoTo ErrHandler_ParseCS
    ParseCSV = Null 'Returns for errors
    If LenB(csvText$) Then
        '@----------------------------------------------------------------------------
        'Set Coerce Type to records and fields
        Select Case P_COERCE_TYPE
            Case 0 'Not corce
                DataDelimiter = P_FIELDS_DELIMITER
                LinesDelimiter = P_RECORDS_DELIMITER
                ParseOption = 0
            Case 1 'Fields delimited with Apostrophe
                DataDelimiter = CHR_APOSTROPHE & P_FIELDS_DELIMITER & CHR_APOSTROPHE
                LinesDelimiter = CHR_APOSTROPHE & P_RECORDS_DELIMITER & CHR_APOSTROPHE
                ParseOption = 1
            Case Else  'Fields delimited with Double Quotes
                DataDelimiter = CHR_DOUBLE_QUOTES & P_FIELDS_DELIMITER & CHR_DOUBLE_QUOTES
                LinesDelimiter = CHR_DOUBLE_QUOTES & P_RECORDS_DELIMITER & CHR_DOUBLE_QUOTES
                ParseOption = 2
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
        If ParseOption <> 0 Then
            'First char omission
            csvArray(ArrayLB) = MidB$(csvArray(ArrayLB), 3)
            'Omit the last last char and VbCrLf if needed
            If RightB$(csvArray(ArrayUB), 4) = vbCrLf Then
                csvArray(ArrayUB) = MidB$(csvArray(ArrayUB), 1, LenB(csvArray(ArrayUB)) - 6)
            Else
                csvArray(ArrayUB) = MidB$(csvArray(ArrayUB), 1, LenB(csvArray(ArrayUB)) - 2)
            End If
        End If
        '@----------------------------------------------------------------------------
        'Check if user want Import all records
        Select Case P_STARTING_RECORD
            Case 0& 'The import will starts in the first record
                lngLB = ArrayLB
            Case Is > ArrayUB 'Out of bound
                lngLB = ArrayUB
            Case Else 'Start imports on user's defined record
                lngLB = P_STARTING_RECORD
        End Select
        Select Case P_ENDING_RECORD
            Case 0& 'Fit Upper bound
                lngUB = ArrayUB
            Case Is < lngLB 'Fit Upper bound
                lngUB = ArrayUB
            Case Is = lngLB 'lngUB nedd to be at least lngLB
                lngUB = lngLB
            Case Is > ArrayUB 'Fit Upper bound
                lngUB = ArrayUB
            Case Else
                lngUB = P_ENDING_RECORD
        End Select
        '@----------------------------------------------------------------------------
        'Take care of headers
        If lngLB = ArrayLB Then  'Only the first record can be omitted
            If lngUB - lngLB > 0 Then 'Two or more records required
                If Not HeadersOmission Then
                    hPos = 0
                Else
                    hPos = 1
                End If
            End If
        End If
        lRows = lngUB - lngLB 'Rows required BASE 0
        lColumns = CSVcolumns(csvArray, DataDelimiter) - 1 'Columns required BASE 0
        lCounter = lngLB + hPos
        ReDim OutPutArray(0 To lRows - hPos, 0 To lColumns) 'Resize the output
        '@----------------------------------------------------------------------------
        'Process the data
        Pointer = hPos 'Point to the position for save the first record
        Do While lCounter <= lngUB
            mCounter = 0 'Column counter
            PrevPos = 1
            'Next position of the pointer over the record
            NextPos = CLng(InStrB(PrevPos, csvArray(lCounter), _
                            DataDelimiter))
            Do While NextPos
                OutPutArray(Pointer - hPos, mCounter) = MidB$(csvArray(lCounter), _
                                                        PrevPos, NextPos - PrevPos)
                PrevPos = NextPos + CLng(LenB(DataDelimiter))
                mCounter = mCounter + 1
                NextPos = CLng(InStrB(PrevPos, csvArray(lCounter), DataDelimiter))
            Loop
            'Handle the last field for the record
            OutPutArray(Pointer - hPos, mCounter) = MidB$(csvArray(lCounter), PrevPos)
            lCounter = lCounter + 1
            Pointer = Pointer + 1
        Loop
        '@----------------------------------------------------------------------------
        'Save results and free the memory
        ParseCSV = OutPutArray
        Erase csvArray
        Erase OutPutArray
        Erase sSplit
    End If
    Exit Function
ErrHandler_ParseCS:
End Function
Public Sub ResetToDefault()
    'Reset the defaults values
    P_COERCE_TYPE = 0
    P_CONNECTED = False
    P_ENDING_RECORD = 0&
    P_ERROR_DESC = vbNullString
    P_ERROR_NUMBER = 0&
    P_FIELDS_DELIMITER = CHR_COMMA
    P_FILENAME = vbNullString
    P_RECORDS_DELIMITER = vbCrLf
    P_STARTING_RECORD = 0&
    P_SUCCESSFUL_EXPORT = True
    P_SUCCESSFUL_IMPORT = True
End Sub
Private Sub Class_Initialize()
    ResetToDefault
End Sub
