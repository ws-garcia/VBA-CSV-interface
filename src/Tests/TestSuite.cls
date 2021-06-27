VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' # TestSuite
'
' A collection of tests, with events and results
'
' ```vba
' Dim Suite As New TestSuite
' Suite.Description = "Name"
'
' With Suite.Test("A")
'   .IsEqual 2 + 2, 4
' End With
' ```
'
' @class TestSuite
' @author tim.hall.engr@gmail.com
' @repository https://github.com/vba-tools/vba-test
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Public Enum TestResultType
    Pass
    Fail
    Pending
    Skipped
End Enum

Public Event BeforeEach(test As TestCase)
Public Event Result(test As TestCase)
Public Event AfterEach(test As TestCase)
Public Event Group(Suite As TestSuite)

''
' (Optional) description of suite for display in runners
''
Public Description As String

''
' @type Collection<TestCase>
''
Public Tests As VBA.Collection

''
' @internal
''
Public Parent As TestSuite

''
' Compute suite result from tests
''
Public Property Get Result() As TestResultType
    Result = TestResultType.Pending
    
    Dim test As TestCase
    For Each test In Me.Tests
        If test.Result = TestResultType.Pass Then
            Result = TestResultType.Pass
        ElseIf test.Result = TestResultType.Fail Then
            Result = TestResultType.Fail
            Exit For
        End If
    Next test
End Property

''
' @type Collection<TestCase>
''
Public Property Get PassedTests() As VBA.Collection
    Set PassedTests = GetTestsByType(TestResultType.Pass)
End Property

''
' @type Collection<TestCase>
''
Public Property Get FailedTests() As VBA.Collection
    Set FailedTests = GetTestsByType(TestResultType.Fail)
End Property

''
' @type Collection<TestCase>
''
Public Property Get PendingTests() As VBA.Collection
    Set PendingTests = GetTestsByType(TestResultType.Pending)
End Property

''
' @type Collection<TestCase>
''
Public Property Get SkippedTests() As VBA.Collection
    Set SkippedTests = GetTestsByType(TestResultType.Skipped)
End Property

Public Property Get Self() As TestSuite
    Set Self = Me
End Property

''
' Create a new test case with description
''
Public Function test(Description As String) As TestCase
    Dim Instance As New TestCase
    
    Instance.Description = Description
    Set Instance.Suite = Me
    
    OnTestBefore Instance
    
    Set test = Instance
End Function

''
' Create a new group of tests with description
''
Public Function Group(Description As String) As TestSuite
    Dim Instance As New TestSuite
    
    Instance.Description = Description
    Set Instance.Parent = Me
    
    RaiseEvent Group(Instance)
    
    Set Group = Instance
End Function

''
' @internal
''
Public Sub TestComplete(test As TestCase)
    OnTestResult test
    OnTestAfter test
End Sub

''
' @internal
''
Public Sub OnTestBefore(test As TestCase)
    If Not Me.Parent Is Nothing Then
        Me.Parent.OnTestBefore test
    End If
    
    RaiseEvent BeforeEach(test)
End Sub

''
' @internal
''
Public Sub OnTestResult(test As TestCase)
    Tests.Add test
    RaiseEvent Result(test)
    
    If Not Me.Parent Is Nothing Then
        Me.Parent.OnTestResult test
    End If
End Sub

''
' @internal
''
Public Sub OnTestAfter(test As TestCase)
    RaiseEvent AfterEach(test)
    
    If Not Me.Parent Is Nothing Then
        Me.Parent.OnTestAfter test
    End If
End Sub

' ============================================= '

Private Function GetTestsByType(ResultType As TestResultType) As VBA.Collection
    Dim test As TestCase
    Dim Filtered As New VBA.Collection
    For Each test In Me.Tests
        If test.Result = ResultType Then
            Filtered.Add test
        End If
    Next test

    Set GetTestsByType = Filtered
End Function


Private Sub Class_Initialize()
    Set Tests = New VBA.Collection
End Sub
