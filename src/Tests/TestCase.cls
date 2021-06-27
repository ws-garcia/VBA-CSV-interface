VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' # TestCase
'
' Verify a single test case with assertions
'
' ```vba
' Dim Test As New TestCase
'
' Test.Description = "should add many numbers
' With Test
'   .IsEqual Add(2, 2), 4
'   .IsEqual Add(1, 2, 3), 6
' End With
' ```
'
' @class TestCase
' @author tim.hall.engr@gmail.com
' @repository https://github.com/vba-tools/vba-test
' @license MIT
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private pFailures As VBA.Collection

''
' Set test description as displayed in reporter
''
Public Description As String

''
' Get/set values for test context
' (useful for passing values to BeforeEach/AfterEach)
''
Public Context As Dictionary

''
' @internal
''
Public Planned As Long

''
' @internal
''
Public Successes As Long

''
' @internal
''
Public Skipped As Boolean

''
' @internal
''
Public Suite As TestSuite

Public Property Get Result() As TestResultType
    If Me.Skipped Then
        Result = TestResultType.Skipped
    ElseIf Me.Successes = 0 And Me.Failures.count = 0 Then
        Result = TestResultType.Pending
    ElseIf Me.Failures.count > 0 Then
        Result = TestResultType.Fail
    Else
        Result = TestResultType.Pass
    End If
End Property

Public Property Get Failures() As VBA.Collection
    Dim Total As Long
    Total = Me.Successes + pFailures.count
    
    If Me.Planned > 0 And Me.Planned <> Total Then
        Dim Message As String
        Dim Failure As Variant
        
        Set Failures = New VBA.Collection
        For Each Failure In pFailures
            Failures.Add Failure
        Next Failure
        
        Message = "Total assertions, {0}, does not equal planned, {1}"
        Failures.Add FormatMessage(Message, Total, Me.Planned)
    Else
        Set Failures = pFailures
    End If
End Property

''
' Access test instance (e.g. for passing to custom matchers)
''
Public Property Get Self() As TestCase
    Self = Me
End Property

''
' Check if two values are deep equal (including Array, Collection, and Dictionary)
''
Public Sub IsEqual(a As Variant, B As Variant, Optional Message As String = _
    "Expected {0} to equal {1}")

    Check IsDeepEqual(a, B), Message, a, B
End Sub

''
' Check if two values are not deep equal (including Array, Collection, and Dictionary)
''
Public Sub NotEqual(a As Variant, B As Variant, Optional Message As String = _
    "Expected {0} to not equal {1}")

    Check Not IsDeepEqual(a, B), Message, a, B
End Sub

''
' Check if a value is "truthy"
'
' From https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement
'
' Must evaluate to True or False, or to a data type that is implicitly convertible to Boolean.
' If the expression is a Nullable Boolean variable that evaluates to Nothing, the condition is treated as if the expression is False.
''
Public Sub IsOk(value As Variant, Optional Message As String = _
    "Expected {0} to be ok")

    Check value, Message, value
End Sub

''
' Check if a value is not "truthy" (See .IsOk)
''
Public Sub NotOk(value As Variant, Optional Message As String = _
    "Expected {0} to not be ok")

    Check Not CBool(value), Message, value
End Sub

''
' Check if a value is "undefined": Nothing, Empty, Null, or Missing
''
Public Sub IsUndefined(Optional value As Variant, Optional Message As String = _
    "Expected {0} to be undefined")

    Check IsNothing(value) Or VBA.IsEmpty(value) Or VBA.IsNull(value) Or VBA.IsMissing(value), Message, value
End Sub

''
' Check if a value is not "undefined": Nothing, Empty, Null, or Missing
''
Public Sub NotUndefined(value As Variant, Optional Message As String = _
    "Expected {0} to not be undefined")

    Check Not IsNothing(value) And Not VBA.IsEmpty(value) And Not VBA.IsNull(value) And Not VBA.IsMissing(value), Message, value
End Sub

''
' Check if the current Err value contains an error with values (if given)
''
Public Sub IsError(Optional Number As Long, Optional Source As String, Optional Description As String, Optional Message As String = _
    "Expected {0} to be an error (with Number = {1}, Source = {2}, Description = {3}")
    
    If Err.Number = 0 Then
        pFailures.Add FormatMessage(Message, "[Error Number=0]", Number, Source, Description)
        Exit Sub
    End If
    
    Check (Number = 0 Or Err.Number = Number) _
        And (Source = "" Or Err.Source = Source) _
        And (Description = "" Or Err.Description = Description), Message, FormattedErr, Number, Source, Description
End Sub

''
' Check if the current Err value does not contain an error
''
Public Sub NotError(Optional Message As String = "Expected {0} to not be an error")
    Check Err.Number = 0, Message, FormattedErr
End Sub

''
' Check if a value is included in an arbitrarily nested Array or Collection
''
Public Sub Includes(values As Variant, value As Variant, Optional Message As String = _
    "Expected {1} to be included in {0}")

    If IsCollection(values) Then
        Check CollectionIncludes(values, value), Message, values, value
    ElseIf IsArray(values) Then
        Check ArrayIncludes(values, value), Message, values, value
    Else
        pFailures.Add FormatMessage(Message, values, value) & " (Incompatible type for Values)"
    End If
End Sub

''
' Check if a value is not included in an arbitrarily nested Array or Collection
''
Public Sub NotIncludes(values As Variant, value As Variant, Optional Message As String = _
    "Expected {1} not to be included in {0}")
    
    If IsCollection(values) Then
        Check Not CollectionIncludes(values, value), Message, values, value
    ElseIf IsArray(values) Then
        Check Not ArrayIncludes(values, value), Message, values, value
    Else
        pFailures.Add FormatMessage(Message, values, value) & " (Incompatible type for Values)"
    End If
End Sub

''
' Check if two values are approximately equal, up to the given amount of significant figures
'
' ```vba
' .IsApproximate 1.001, 1.002, 3
'
' ' Equivalent to .IsEqual 1.00e+0, 1.00e+0
' ```
''
Public Sub IsApproximate(a As Variant, B As Variant, SignificantFigures As Integer, Optional Message As String = _
    "Expected {0} to be approximately equal to {1} (with {2} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        pFailures.Add "IsApproximate can only compare from 1 to 15 significant figures"
    Else
        Check IsApproximatelyEqual(a, B, SignificantFigures), Message, a, B, SignificantFigures
    End If
End Sub

''
' Check if two values are approximately equal, up to the given amount of significant figures
'
' ```vba
' .NotApproximate 1.001, 1.009, 3
'
' ' Equivalent to .IsEqual 1.00e+0, 1.01e+0
' ```
''
Public Sub NotApproximate(a As Variant, B As Variant, SignificantFigures As Integer, Optional Message As String = _
    "Expected {0} to not be approximately equal to {1} (with {2} significant figures of precision)")

    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        pFailures.Add "NotApproximate can only compare from 1 to 15 significant figures"
    Else
        Check Not IsApproximatelyEqual(a, B, SignificantFigures), Message, a, B, SignificantFigures
    End If
End Sub

''
' Mark the test as passing
''
Public Sub Pass()
    Me.Successes = 1
    Set pFailures = New VBA.Collection
End Sub

''
' Mark the test as failing
''
Public Sub Fail(Optional Message As String = _
    "Test failed unexpectedly")
    
    pFailures.Add Message
End Sub

''
' Set the planned number of assertions for the test
''
Public Sub Plan(count As Long)
    Planned = count
End Sub

''
' Mark the test as skipped
''
Public Sub Skip()
    Me.Skipped = True
End Sub

' ============================================= '

Private Sub Check(Assertion As Variant, Message As String, ParamArray values() As Variant)
    If Assertion Then
        Me.Successes = Me.Successes + 1
    Else
        pFailures.Add FormatMessage(Message, values)
    End If
End Sub

Private Function IsDeepEqual(a As Variant, B As Variant) As Boolean
    Dim AType As VbVarType
    Dim BType As VbVarType
    
    AType = VBA.VarType(a)
    BType = VBA.VarType(B)

    If VBA.IsError(a) Or VBA.IsError(B) Then
        IsDeepEqual = False
        
    ElseIf VBA.IsArray(a) And VBA.IsArray(B) Then
        IsDeepEqual = IsArrayEqual(a, B)
    
    ElseIf AType = VBA.vbObject Or BType = VBA.vbObject Then
        If AType <> BType Or VBA.TypeName(a) <> VBA.TypeName(B) Then
            IsDeepEqual = False
        ElseIf VBA.TypeName(a) = "Collection" Then
            IsDeepEqual = IsCollectionEqual(a, B)
        ElseIf VBA.TypeName(a) = "Dictionary" Then
            IsDeepEqual = IsDictionaryEqual(a, B)
        ElseIf VBA.TypeName(a) = "ECPArrayList" Then
            IsDeepEqual = IsECPArrayListEqual(a, B)
        Else
            IsDeepEqual = a Is B
        End If
    
    ElseIf VBA.VarType(a) = VBA.vbDouble Or VBA.VarType(B) = VBA.vbDouble Then
        ' It is inherently difficult/almost impossible to check equality of Double
        ' http://support.microsoft.com/kb/78113
        '
        ' -> Compare up to 15 significant figures
        IsDeepEqual = IsApproximatelyEqual(a, B, 15)
    
    Else
        IsDeepEqual = a = B
    End If
End Function

Private Function IsArrayEqual(a As Variant, B As Variant) As Boolean
    If UBound(a) <> UBound(B) Then
        IsArrayEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(a) To UBound(a)
        If Not IsDeepEqual(a(i), B(i)) Then
            IsArrayEqual = False
            Exit Function
        End If
    Next i
    
    IsArrayEqual = True
End Function

Private Function IsCollectionEqual(a As Variant, B As Variant) As Boolean
    If a.count <> B.count Then
        IsCollectionEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To a.count
        If Not IsDeepEqual(a(i), B(i)) Then
            IsCollectionEqual = False
            Exit Function
        End If
    Next i

    IsCollectionEqual = True
End Function
Private Function IsECPArrayListEqual(a As Variant, B As Variant) As Boolean
    If a.count <> B.count Then
        IsECPArrayListEqual = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To a.count - 1
        If Not IsDeepEqual(a(i), B(i)) Then
            IsECPArrayListEqual = False
            Exit Function
        End If
    Next i

    IsECPArrayListEqual = True
End Function
Private Function IsDictionaryEqual(a As Variant, B As Variant) As Boolean
    If UBound(a.Keys) <> UBound(B.Keys) Then
        IsDictionaryEqual = False
        Exit Function
    End If
    
    Dim AKeys As Variant
    Dim BKeys As Variant
    Dim i As Long
    
    AKeys = a.Keys
    BKeys = B.Keys
    
    For i = LBound(AKeys) To UBound(AKeys)
        If AKeys(i) <> BKeys(i) Or a.item(AKeys(i)) <> B.item(BKeys(i)) Then
            IsDictionaryEqual = False
            Exit Function
        End If
    Next i

    IsDictionaryEqual = True
End Function

Private Function IsCollection(value As Variant) As Boolean
    IsCollection = VBA.VarType(value) = VBA.vbObject And VBA.TypeName(value) = "Collection"
End Function

Private Function IsNothing(value As Variant) As Boolean
    If VBA.IsObject(value) Then
        IsNothing = value Is Nothing
    Else
        IsNothing = False
    End If
End Function

Private Function ArrayIncludes(values As Variant, value As Variant) As Boolean
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If VBA.IsArray(values(i)) Then
            If ArrayIncludes(values(i), value) Then
                ArrayIncludes = True
                Exit Function
            End If
        ElseIf IsCollection(values(i)) Then
            If CollectionIncludes(values(i), value) Then
                ArrayIncludes = True
                Exit Function
            End If
        ElseIf IsDeepEqual(values(i), value) Then
            ArrayIncludes = True
            Exit Function
        End If
    Next i

    ArrayIncludes = False
End Function

Private Function CollectionIncludes(values As Variant, value As Variant) As Boolean
    Dim item As Variant
    For Each item In values
        If VBA.IsArray(item) Then
            If ArrayIncludes(item, value) Then
                CollectionIncludes = True
                Exit Function
            End If
        ElseIf IsCollection(item) Then
            If CollectionIncludes(item, value) Then
                CollectionIncludes = True
                Exit Function
            End If
        ElseIf IsDeepEqual(item, value) Then
            CollectionIncludes = True
            Exit Function
        End If
    Next item
    
    CollectionIncludes = False
End Function

Private Function IsApproximatelyEqual(a As Variant, B As Variant, SignificantFigures As Integer) As Boolean
    If SignificantFigures < 1 Or SignificantFigures > 15 Or VBA.IsError(a) Or VBA.IsError(B) Then
        IsApproximatelyEqual = False
        Exit Function
    End If
    
    Dim AValue As String
    Dim BValue As String
    
    AValue = VBA.Format$(a, VBA.Left$("0.00000000000000", SignificantFigures + 1) & IIf(a > 1, "e+0", "e-0"))
    BValue = VBA.Format$(B, VBA.Left$("0.00000000000000", SignificantFigures + 1) & IIf(B > 1, "e+0", "e-0"))
    
    IsApproximatelyEqual = AValue = BValue
End Function

Private Function FormatMessage(Message As String, ParamArray values() As Variant) As String
    Dim value As Variant
    Dim Index As Long
    
    FormatMessage = Message
    For Each value In IIf(VBA.IsArray(values(0)), values(0), values)
        FormatMessage = VBA.Replace(FormatMessage, "{" & Index & "}", PrettyPrint(value))
        Index = Index + 1
    Next value
End Function

Private Function PrettyPrint(value As Variant, Optional Indentation As Long = 0) As String
    If VBA.IsMissing(value) Then
        PrettyPrint = "[Missing]"
        Exit Function
    End If
    
    Dim i As Long
    Dim Indented As String
    Indented = VBA.String$(Indentation + 1, "  ")
    
    Select Case VBA.VarType(value)
    Case VBA.vbObject
        ' Nothing
        If value Is Nothing Then
            PrettyPrint = "[Nothing]"
        
        ' Collection
        ElseIf VBA.TypeName(value) = "Collection" Then
            PrettyPrint = "[Collection [" & vbNewLine
            
            For i = 1 To value.count
                PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                    PrettyPrint(value(i), Indentation + 1) & _
                    IIf(i <> value.count, ",", "") & vbNewLine
            Next i
            
            PrettyPrint = PrettyPrint & Indent(Indentation) & "]"
        
        ' Dictionary
        ElseIf VBA.TypeName(value) = "Dictionary" Then
            PrettyPrint = "[Dictionary {" & vbNewLine
            
            For i = LBound(value.Keys) To UBound(value.Keys)
                PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                    value.Keys(i) & ": " & _
                    PrettyPrint(value.item(value.Keys(i)), Indentation + 1) & _
                    IIf(i <> value.count, ",", "") & vbNewLine
            Next i
            
            PrettyPrint = PrettyPrint & Indent(Indentation) & "}]"
        
        ' Object
        Else
            PrettyPrint = "[" & VBA.TypeName(value) & "]"
        End If
        
    ' Array
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        PrettyPrint = "[" & vbNewLine
        
        For i = LBound(value) To UBound(value)
            PrettyPrint = PrettyPrint & Indent(Indentation + 1) & _
                PrettyPrint(value(i), Indentation + 1) & _
                IIf(i <> UBound(value), ",", "") & vbNewLine
        Next i
        
        PrettyPrint = PrettyPrint & Indent(Indentation) & "]"
    
    ' Empty
    Case VBA.vbEmpty
        PrettyPrint = "[Empty]"
    
    ' Null
    Case VBA.vbNull
        PrettyPrint = "[Null]"
    
    ' String
    Case VBA.vbString
        PrettyPrint = """" & value & """"
    
    ' Everything else
    Case Else
        PrettyPrint = CStr(value)
    End Select
End Function

Private Function FormattedErr() As String
    Dim ErrNumberDetails As String
    
    ErrNumberDetails = IIf(Err.Number < 0, " (" & (Err.Number - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "")
    FormattedErr = "[Error Number=" & Err.Number & ErrNumberDetails & ", Source=" & Err.Source & ", Description=" & Err.Description & "]"
End Function

Private Function Indent(Optional Indentation As Long)
    Indent = VBA.String$(Indentation, "  ")
End Function

Private Sub Class_Initialize()
    Set Me.Context = New Dictionary
    Set pFailures = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Me.Suite.TestComplete Me
    Set Me.Context = Nothing
End Sub
