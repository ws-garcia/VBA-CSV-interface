VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' # FileReporter
'
' Append test results to the given file
'
' ```vba
' Dim Suite As New TestSuite
' ...
'
' Dim Reporter As New FileReporter
' Reporter.WriteTo "path/to/file"
' Reporter.ListenTo Suite
' ```
'
' @class FileReporter
' @author Tim Hall <tim.hall.engr@gmail.com>
' @repository https://github.com/vba-tools/vba-test
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private WithEvents pSuite As TestSuite
Attribute pSuite.VB_VarHelpID = -1
Private FilePath As String
Private Finished As Boolean

''
' Report test results to the given file path
''
Public Sub WriteTo(path As Variant)
    FilePath = path
End Sub

''
' Report test results from the given TestSuite
''
Public Sub ListenTo(Suite As TestSuite)
    If Not pSuite Is Nothing Then
        PrintSummary
    End If
    
    Finished = False
    Set pSuite = Suite
    PrintHeader Suite
End Sub

' ============================================= '

Private Sub PrintHeader(Suite As TestSuite)
    AppendToFile "===" & VBA.IIf(Suite.Description <> "", " " & Suite.Description & " ===", "")
End Sub

Private Sub PrintResult(test As TestCase)
    If test.Result = TestResultType.Skipped Then
        Exit Sub
    End If

    AppendToFile ResultTypeToString(test.Result) & " " & test.Description
    
    If test.Result = TestResultType.Fail Then
        Dim Failure As Variant
        For Each Failure In test.Failures
            AppendToFile "  " & Failure
        Next Failure
    End If
End Sub

Private Sub PrintSummary()
    Dim Total As Long
    Dim Passed As Long
    Dim Failed As Long
    Dim Pending As Long
    Dim Skipped As Long
    
    Total = pSuite.Tests.count
    Passed = pSuite.PassedTests.count
    Failed = pSuite.FailedTests.count
    Pending = pSuite.PendingTests.count
    Skipped = pSuite.SkippedTests.count
    
    Dim Summary As String
    If Failed > 0 Then
        Summary = "FAIL (" & Failed & " of " & Total & " failed"
    Else
        Summary = "PASS (" & Passed & " of " & Total & " passed"
    End If
    If Pending > 0 Then
        Summary = Summary & ", " & Pending & " pending"
    End If
    If Skipped > 0 Then
        Summary = Summary & ", " & Skipped & " skipped)"
    Else
        Summary = Summary & ")"
    End If

    AppendToFile "= " & Summary & " = " & Now & " =" & vbNewLine
End Sub

Private Function ResultTypeToString(ResultType As TestResultType) As String
    Select Case ResultType
    Case TestResultType.Pass
        ResultTypeToString = "+"
    Case TestResultType.Fail
        ResultTypeToString = "X"
    Case TestResultType.Pending
        ResultTypeToString = "."
    End Select
End Function

Private Sub AppendToFile(Message As String)
    If FilePath = "" Then Exit Sub
    
    Dim File As Integer
    File = FreeFile
    
    On Error GoTo Cleanup
    
    Open FilePath For Append As #File
    Print #File, Message
    
Cleanup:
    Close #File
End Sub

Private Sub pSuite_Result(test As TestCase)
    PrintResult test
End Sub

Private Sub pSuite_Group(Suite As TestSuite)
    PrintHeader Suite
End Sub

Private Sub Class_Terminate()
    If Not Finished Then
        PrintSummary
    End If
End Sub

