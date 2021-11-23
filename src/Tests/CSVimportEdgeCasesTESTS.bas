Attribute VB_Name = "CSVimportEdgeCasesTESTS"
Option Explicit
Private Const CHR_D_QUOTES As String = """"
Private DquotesAsEscapeToken As Boolean
Private ActualResult As CSVArrayList
Private ExpectedResult As CSVArrayList
Private confObj As CSVparserConfig
Private EscapedFieldDelimiterReplacement As String
Public Enum ImportMode
    iStream = 0
    iString = 1
    iSequential = 2
End Enum
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'RUN TEST
Public Sub RunALLtest()
    Dim filePath As String
    
    filePath = ThisWorkbook.path & Application.PathSeparator & "results" & Application.PathSeparator & _
                        "CSV import test - " & Format(Now, "dd-mmm-yyyy h-mm-ss") & ".txt"
                        
    RunStreamImportTest filePath
    RunStringImportTest filePath
    RunSequentialImportTest filePath
    ClearObjects
End Sub
Public Sub RunStreamImportTest(filePath As String)
    ImportTests filePath, ImportMode.iStream
End Sub
Public Sub RunStringImportTest(filePath As String)
    ImportTests filePath, ImportMode.iString
End Sub
Public Sub RunSequentialImportTest(filePath As String)
    ImportTests filePath, ImportMode.iSequential
End Sub
Private Sub ClearObjects()
    Set ActualResult = Nothing
    Set ExpectedResult = Nothing
    Set confObj = Nothing
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'CSVArrayList Generator
Public Function CreateExpectedRecord(fields() As String) As Variant
    Dim elemID As Long
    Dim tmpResult As CSVArrayList
    
    Set tmpResult = New CSVArrayList
    For elemID = LBound(fields) To UBound(fields)
        tmpResult.Add fields(elemID)
    Next
    tmpResult.ShrinkBuffer
    CreateExpectedRecord = tmpResult.items
End Function
Public Function CreateExpectedCSVresult(commaAndpipeDelimitedCSVstring As String) As CSVArrayList
    If commaAndpipeDelimitedCSVstring <> vbNullString Then
        Dim csvRecords() As String
        Dim filedsArray() As String
        Dim cleanedString As String
        Dim elemID As Long
        Dim iCounter As Long
        
        '@--------------------------------------------------------------------------------
        'Replace \r and \n
        cleanedString = Replace(Replace(commaAndpipeDelimitedCSVstring, "\r", vbCr, 1), "\n", vbLf)
        '@--------------------------------------------------------------------------------
        'Replace single quotes if needed
        If DquotesAsEscapeToken Then
            If InStrB(1, cleanedString, "'") Then
                cleanedString = Replace(cleanedString, "'", CHR_D_QUOTES, 1)
            End If
        End If
        csvRecords() = Split(cleanedString, "|")
        Set CreateExpectedCSVresult = New CSVArrayList
        For elemID = LBound(csvRecords) To UBound(csvRecords)
            If csvRecords(elemID) <> vbNullString Then
                filedsArray() = Split(csvRecords(elemID), ",")
            Else
                ReDim filedsArray(0)
                filedsArray(0) = csvRecords(elemID)
            End If
            '@--------------------------------------------------------------------------------
            'Replace existing interior [?] chars
            For iCounter = LBound(filedsArray) To UBound(filedsArray)
                If InStrB(1, filedsArray(iCounter), "?") Then
                    filedsArray(iCounter) = Replace(filedsArray(iCounter), "?", EscapedFieldDelimiterReplacement, 1)
                End If
            Next iCounter
            CreateExpectedCSVresult.Add CreateExpectedRecord(filedsArray)
        Next
        CreateExpectedCSVresult.ShrinkBuffer
    Else
        Set CreateExpectedCSVresult = Nothing
    End If
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Unit testing
Function ImportTests(FullFileName As String, _
                        Optional ReadMode As ImportMode = ImportMode.iStream) As TestSuite
    Set ImportTests = New TestSuite
    Select Case ReadMode
        Case ImportMode.iStream
            ImportTests.Description = "StreamCSVimport"
        Case ImportMode.iString
            ImportTests.Description = "StringCSVimport"
        Case Else
            ImportTests.Description = "SequentialCSVimport"
    End Select

  ' Report results to a text file
    Dim Suite As New TestSuite
    Dim Reporter As New FileReporter
    
    Reporter.WriteTo FullFileName
                        
    Reporter.ListenTo ImportTests
    
    On Error Resume Next
    '@--------------------------------------------------------------------------------
    'Bad comments value specified
    With ImportTests.Test("Bad comments value specified")
        BadCommentsValueSpecified ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 3 records"
    End With
    '@--------------------------------------------------------------------------------
    'Comment with non-default character
    With ImportTests.Test("Comment with non-default character")
        CommentWithNonDefaultCharacter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Commented line at beginning
    With ImportTests.Test("Commented line at beginning")
        CommentedLineAtBeginning ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Commented line at end
    With ImportTests.Test("Commented line at end")
        CommentedLineAtEnd ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Commented line in middle
    With ImportTests.Test("Commented line in middle")
        CommentedLineInMiddle ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Entire file is comment lines
    With ImportTests.Test("Entire file is comment lines")
        EntireFileIsCommentLines ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Input is just a string (a single field)
    With ImportTests.Test("Input is just a string (a single field)")
        InputIsJustAString_ASingleField_ ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input is just empty fields
    With ImportTests.Test("Input is just empty fields")
        InputIsJustEmptyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 and 4 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input is just the delimiter (2 empty fields)
    With ImportTests.Test("Input is just the delimiter (2 empty fields)")
        InputIsJustTheDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 2 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input with only a commented line and blank line after
    With ImportTests.Test("Input with only a commented line and blank line after")
        InputWithOnlyACommentedLineAndBlankLineAfter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Input with only a commented line, without comments enabled
    With ImportTests.Test("Input with only a commented line, without comments enabled")
        InputWithOnlyACommentedLineWithoutCommentsEnabled ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 1 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input without comments with line starting with whitespace
    With ImportTests.Test("Input without comments with line starting with whitespace")
        InputWithoutCommentsWithLineStartingWithWhitespace ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 3 records with 1 field (preserving whitespace)"
    End With
    '@--------------------------------------------------------------------------------
    'Line ends with quoted field
    With ImportTests.Test("Line ends with quoted field")
        LineEndsWithQuotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 4 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Line starts with quoted field
    With ImportTests.Test("Line starts with quoted field")
        LineStartsWithQuotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Misplaced quotes in data, not as opening quotes
    With ImportTests.Test("Misplaced quotes in data, not as opening quotes")
        MisplacedQuotesInDataNotAsOpeningQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Multiple consecutive empty fields
    With ImportTests.Test("Multiple consecutive empty fields")
        MultipleConsecutiveEmptyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 6 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Multiple rows, one column (no delimiter found)
    With ImportTests.Test("Multiple rows, one column (no delimiter found)")
        MultipleRowsOneColumn ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 5 records with 1 fields"
    End With
    '@--------------------------------------------------------------------------------
    'One column input with empty fields
    With ImportTests.Test("One column input with empty fields")
        OneColumnInputWithEmptyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 7 records with 1 fields"
    End With
    '@--------------------------------------------------------------------------------
    'One Row
    With ImportTests.Test("One Row")
        OneRow ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Pipe delimiter
    With ImportTests.Test("Pipe delimiter")
        PipeDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field at end of row (but not at EOF) has quotes
    With ImportTests.Test("Quoted field at end of row (but not at EOF) has quotes")
        QuotedFieldAtEndOfRowButNotAtEOFhasQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field has no closing quote
    With ImportTests.Test("Quoted field has no closing quote")
        QuotedFieldHasNoClosingQuot ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with 5 quotes in a row and a delimiter
    With ImportTests.Test("Quoted field with 5 quotes in a row and a delimiter")
        QuotedFieldWith5QuotesInARowAndADelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with delimiter
    With ImportTests.Test("Quoted field with delimiter")
        QuotedFieldWithDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with escaped quotes at boundaries
    With ImportTests.Test("Quoted field with escaped quotes at boundaries")
        QuotedFieldWithEscapedQuotesAtBoundaries ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with escaped quotes
    With ImportTests.Test("Quoted field with escaped quotes")
        QuotedFieldWithEscapedQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with extra whitespace on edges
    With ImportTests.Test("Quoted field with extra whitespace on edges")
        QuotedFieldWithExtraWhitespaceOnEdges ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with line break
    With ImportTests.Test("Quoted field with line break")
        QuotedFieldWithLineBreak ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with quotes around delimiter
    With ImportTests.Test("Quoted field with quotes around delimiter")
        QuotedFieldWithQuotesAroundDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with quotes on left side of delimiter
    With ImportTests.Test("Quoted field with quotes on left side of delimiter")
        QuotedFieldWithQuotesOnLeftSideOfDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with quotes on right side of delimiter
    With ImportTests.Test("Quoted field with quotes on right side of delimiter")
        QuotedFieldWithQuotesOnRightSideOfDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with Unix escaped quotes at boundaries
    With ImportTests.Test("Quoted field with Unix escaped quotes at boundaries")
        QuotedFieldWithUnixEscapedQuotesAtBoundaries ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with whitespace around quotes
    With ImportTests.Test("Quoted field with whitespace around quotes")
        QuotedFieldWithWhitespaceAroundQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field
    With ImportTests.Test("Quoted field")
        QuotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted fields at end of row with delimiter and line break
    With ImportTests.Test("Quoted fields at end of row with delimiter and line break")
        QuotedFieldsAtEndOfRowWithDelimiterAndLinBreak ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted fields with line breaks
    With ImportTests.Test("Quoted fields with line breaks")
        QuotedFieldsWithLineBreaks
        .IsEqual ActualResult, ExpectedResult, "Expected 3 fields and 1 record"
    End With
    '@--------------------------------------------------------------------------------
    'Row with enough fields but blank field at end
    With ImportTests.Test("Row with enough fields but blank field at end")
        RowWithEnoughFieldsButBlankFieldAtEnd ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Row with too few fields
    With ImportTests.Test("Row with too few fields")
        RowWithTooFewFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 and 2 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Row with too many fields
    With ImportTests.Test("Row with too many fields")
        RowWithTooManyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 3 record with 3 and 5 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Skip empty lines, with empty input
    With ImportTests.Test("Skip empty lines, with empty input")
        SkipEmptyLinesWithEmptyInput ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Skip empty lines, with first line only whitespace
    With ImportTests.Test("Skip empty lines, with first line only whitespace")
        SkipEmptyLinesWithFirstLineOnlyWhitespace ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 1 and 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Skip empty lines, with newline at end of input
    With ImportTests.Test("Skip empty lines, with newline at end of input")
        SkipEmptyLinesWithNewlineAtEndOfInput ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Tab delimiter
    With ImportTests.Test("Tab delimiter")
        TabDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Three comment lines consecutively at beginning of file
    With ImportTests.Test("Three comment lines consecutively at beginning of file")
        ThreeCommentLinesConsecutivelyAtBeginningOfFile ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Two comment lines consecutively at end of file
    With ImportTests.Test("Two comment lines consecutively at end of file")
        TwoCommentLinesConsecutivelyAtEndOfFile ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Two comment lines consecutively
    With ImportTests.Test("Two comment lines consecutively")
        TwoCommentLinesConsecutively ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Two rows
    With ImportTests.Test("Two rows")
        TwoRows ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Unquoted field with quotes at end of field
    With ImportTests.Test("Unquoted field with quotes at end of field")
        UnquotedFieldWithQuotesAtEndOfField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Whitespace at edges of unquoted field
    With ImportTests.Test("Whitespace at edges of unquoted field")
        WhitespaceAtEdgesOfUnquotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Complex CSV syntax
    With ImportTests.Test("Complex CSV syntax")
        ComplexCSVsyntax ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 4 fields"
    End With
    Set ImportTests = Nothing
End Function
Sub GetActualAndExpectedResults(FileName As String, _
                                ExpectedResultString As String, _
                                Optional ReadMode As ImportMode = ImportMode.iStream)
    Dim csv As CSVinterface
    
    Set csv = New CSVinterface
    With confObj
        .path = ThisWorkbook.path & Application.PathSeparator & FileName
        .delimitersGuessing = True
    End With
    Select Case ReadMode
        Case ImportMode.iStream 'Import entire CSV file contents using streams
            csv.ImportFromCSV confObj
            If Not csv Is Nothing Then
                Set ActualResult = csv.items
            Else
                Set ActualResult = Nothing
            End If
        Case ImportMode.iString 'Parse string holding the CSV file contents
            Dim CSVtext As CSVTextStream
            
            Set CSVtext = New CSVTextStream
            CSVtext.OpenStream confObj.path
            CSVtext.ReadText
            csv.ImportFromCSVString CSVtext.bufferString, confObj
            If Not csv Is Nothing Then
                Set ActualResult = csv.items
            Else
                Set ActualResult = Nothing
            End If
            Set CSVtext = Nothing
        Case Else 'Import entire CSV file sequential
            Dim csvRecord As CSVArrayList
            Dim tmpResult As CSVArrayList
            
            Set tmpResult = New CSVArrayList
            csv.OpenSeqReader confObj
            Set csvRecord = csv.GetRecord
            Do While Not csvRecord Is Nothing
                tmpResult.Add csvRecord.item(0)
                Set csvRecord = csv.GetRecord
            Loop
            If tmpResult.count Then
                Set ActualResult = tmpResult
            Else
                Set ActualResult = Nothing
            End If
    End Select
    Set ExpectedResult = CreateExpectedCSVresult(ExpectedResultString)
    DquotesAsEscapeToken = (confObj.dialect.quoteToken = QuoteTokens.DoubleQuotes)
    EscapedFieldDelimiterReplacement = confObj.dialect.fieldsDelimiter
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Testing Functions
Sub QuotedFieldsWithLineBreaks(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Quoted fields with line breaks.csv", "A,B\r\nB,C\r\nC\r\nC", ReadMode
End Sub
Sub BadCommentsValueSpecified(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    confObj.commentsToken = 5 'valid comments token [#](by default) , [!], [$], [%], and [&].
    
    GetActualAndExpectedResults "Bad comments value specified.csv", "a,b,c|5comment|d,e,f", ReadMode
End Sub
Sub CommentWithNonDefaultCharacter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    confObj.commentsToken = "!"
    
    GetActualAndExpectedResults "Comment with non-default character.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub CommentedLineAtBeginning(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Commented line at beginning.csv", "a,b,c", ReadMode
End Sub
Sub CommentedLineAtEnd(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Commented line at end.csv", "a,true,false", ReadMode
End Sub
Sub CommentedLineInMiddle(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Commented line in middle.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub EntireFileIsCommentLines(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Entire file is comment lines.csv", vbNullString, ReadMode
End Sub
Sub InputIsJustAString_ASingleField_(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Input is just a string (a single field).csv", "Abc def", ReadMode
End Sub
Sub InputIsJustEmptyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Input is just empty fields.csv", ",,|,,,", ReadMode
End Sub
Sub InputIsJustTheDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
     
    GetActualAndExpectedResults "Input is just the delimiter (2 empty fields).csv", ",", ReadMode
End Sub
Sub InputWithOnlyACommentedLineAndBlankLineAfter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Input with only a commented line and blank line after.csv", vbNullString, ReadMode
End Sub
Sub InputWithOnlyACommentedLineWithoutCommentsEnabled(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    confObj.skipCommentLines = False 'Disable comment lines
    
    GetActualAndExpectedResults "Input with only a commented line, without comments enabled.csv", "#commented line", ReadMode
End Sub
Sub InputWithoutCommentsWithLineStartingWithWhitespace(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Input without comments with line starting with whitespace.csv", "a| b|c", ReadMode
End Sub
Sub LineEndsWithQuotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Line ends with quoted field.csv", "a,b,c|d,e,f|g,h,i|j,k,l", ReadMode
End Sub
Sub LineStartsWithQuotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Line starts with quoted field.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub MisplacedQuotesInDataNotAsOpeningQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Misplaced quotes in data, not as opening quotes.csv", "A,B 'B',C", ReadMode
End Sub
Sub MultipleConsecutiveEmptyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Multiple consecutive empty fields.csv", "a,b,,,c,d|,,e,,,f", ReadMode
End Sub
Sub MultipleRowsOneColumn(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Multiple rows, one column (no delimiter found).csv", "a|b|c|d|e", ReadMode
End Sub
Sub OneColumnInputWithEmptyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    confObj.skipEmptyLines = False 'Disable skip empty lines

    GetActualAndExpectedResults "One column input with empty fields.csv", "a|b|||c|d|e|", ReadMode
End Sub
Sub OneRow(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "One row.csv", "A,b,c", ReadMode
End Sub
Sub PipeDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Pipe delimiter.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub QuotedFieldAtEndOfRowButNotAtEOFhasQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field at end of row (but not at EOF) has quotes.csv", "a,b,c'c'|d,e,f", ReadMode
End Sub
Sub QuotedFieldHasNoClosingQuot(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field has no closing quote.csv", vbNullString, ReadMode
End Sub
Sub QuotedFieldWithDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with delimiter.csv", "A,B?B,C", ReadMode
End Sub
Sub QuotedFieldWithEscapedQuotesAtBoundaries(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with escaped quotes at boundaries.csv", "A,'B',C", ReadMode
End Sub
Sub QuotedFieldWithEscapedQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with escaped quotes.csv", "A,B'B'B,C", ReadMode
End Sub
Sub QuotedFieldWithExtraWhitespaceOnEdges(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with extra whitespace on edges.csv", "A, B  ,C", ReadMode
End Sub
Sub QuotedFieldWithLineBreak(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with line break.csv", "A,B\nB,C", ReadMode
End Sub
Sub QuotedFieldWithQuotesAroundDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with quotes around delimiter.csv", "A,'?',C", ReadMode
End Sub
Sub QuotedFieldWithQuotesOnLeftSideOfDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with quotes on left side of delimiter.csv", "A,'?,C", ReadMode
End Sub
Sub QuotedFieldWithQuotesOnRightSideOfDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with quotes on right side of delimiter.csv", "A,?',C", ReadMode
End Sub
Sub QuotedFieldWithWhitespaceAroundQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with whitespace around quotes.csv", "A,B,C", ReadMode
End Sub
Sub QuotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field.csv", "A,B,C", ReadMode
End Sub
Sub QuotedFieldsAtEndOfRowWithDelimiterAndLinBreak(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted fields at end of row with delimiter and line break.csv", "a,b,c?c\nc|d,e,f", ReadMode
End Sub
Sub RowWithEnoughFieldsButBlankFieldAtEnd(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Row with enough fields but blank field at end.csv", "A,B,C|a,b,", ReadMode
End Sub
Sub RowWithTooFewFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Row with too few fields.csv", "A,B,C|a,b", ReadMode
End Sub
Sub RowWithTooManyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Row with too many fields.csv", "A,B,C|a,b,c,d,e|f,g,h", ReadMode
End Sub
Sub SkipEmptyLinesWithEmptyInput(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Skip empty lines, with empty input.csv", vbNullString, ReadMode
End Sub
Sub SkipEmptyLinesWithFirstLineOnlyWhitespace(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Skip empty lines, with first line only whitespace.csv", " |a,b,c", ReadMode
End Sub
Sub SkipEmptyLinesWithNewlineAtEndOfInput(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Skip empty lines, with newline at end of input.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub TabDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Tab delimiter.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub ThreeCommentLinesConsecutivelyAtBeginningOfFile(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Three comment lines consecutively at beginning of file.csv", "a,b,c", ReadMode
End Sub
Sub TwoCommentLinesConsecutivelyAtEndOfFile(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Two comment lines consecutively at end of file.csv", "a,b,c", ReadMode
End Sub
Sub TwoCommentLinesConsecutively(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Two comment lines consecutively.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub TwoRows(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Two rows.csv", "A,b,c|d,E,f", ReadMode
End Sub
Sub UnquotedFieldWithQuotesAtEndOfField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Unquoted field with quotes at end of field.csv", "A,B',C", ReadMode
End Sub
Sub WhitespaceAtEdgesOfUnquotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Whitespace at edges of unquoted field.csv", "a,  b ,c", ReadMode
End Sub
Sub QuotedFieldWith5QuotesInARowAndADelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig

    GetActualAndExpectedResults "Quoted field with 5 quotes in a row and a delimiter.csv", "1,cnonce=''?nc='',2", ReadMode
End Sub
Sub QuotedFieldWithUnixEscapedQuotesAtBoundaries(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    confObj.dialect.escapeMode = EscapeStyle.unix 'Enable Unix escape mechanism
    GetActualAndExpectedResults "Quoted field with unix escaped quotes at boundaries.csv", "A,'B',C", ReadMode
End Sub
Sub ComplexCSVsyntax(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Complex CSV syntax.csv", "rec1? fld1,,rec1'?'fld3.1\r\n'?\r\nfld3.2,rec1\r\nfld4" _
                                                        & "|rec2? fld1.1\r\n\r\nfld1.2,rec2 fld2.1'fld2.2'fld2.3,,rec2 fld4", ReadMode
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#

