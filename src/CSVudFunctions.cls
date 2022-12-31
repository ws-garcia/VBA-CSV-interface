VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVudFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'////////////////////////////////////////////////////////////////////////////////////////////
'#
' GENERAL INFO:
' Class module developed to provide samples of user defined functions (UDF).
Private Const Apostrophe As String = "'"

Public Function GCD(ByRef aValues As Variant) As Long 'Expected an array
    Dim t As Long
    Dim u As Long
    Dim v As Long
    Static RecursionLevel As Long
    
    RecursionLevel = RecursionLevel + 1
    u = CLng(aValues(LBound(aValues)))
    v = CLng(aValues(LBound(aValues) + RecursionLevel))
    If u < v Then 'change values
        t = u
        u = v
        v = t
    End If
    Do While v
        t = u
        u = v
        v = t Mod v
    Loop
    If RecursionLevel < UBound(aValues) Then
        aValues(LBound(aValues)) = u
        GCD = GCD(aValues)
    Else
        GCD = u
        RecursionLevel = RecursionLevel - 1
    End If
End Function
Private Function minor(a() As Double, x As Integer, y As Integer) As Double()
    Dim l As Long: l = UBound(a) - 1
    Dim i As Long, j As Long
    Dim Result() As Double
    
    If l > 0 Then ReDim Result(1 To l, 1 To l)
    For i = 1 To l
        For j = 1 To l
            Result(i, j) = a(i - (i >= x), j - (j >= y))
        Next j
    Next i
    minor = Result
End Function

'Adapted from: https://rosettacode.org/wiki/Determinant_and_permanent#VBA
Public Function DET(strArray As Variant) As Double 'Expected an one element string array
    Dim a() As Double
    Dim ArrayHelper As New CSVexpressions
    
    a() = StringTodblArray(ArrayHelper.ArrayFromString(CStr(strArray(LBound(strArray)))))
    DET = DET_(a)
End Function

Private Function DET_(a() As Double) As Double
    If UBound(a) = 1 Then
        On Error GoTo err
        DET_ = a(1, 1)
        Exit Function
    End If
    Dim sgn_ As Integer: sgn_ = 1
    Dim Res As Integer: Res = 0
    Dim i As Integer
    For i = 1 To UBound(a)
        Res = Res + sgn_ * a(1, i) * DET_(minor(a, 1, i))
        sgn_ = sgn_ * -1
    Next i
    DET_ = Res
    Exit Function
err:
    DET_ = a(1)
End Function
'Gets an array from a string like "{{1;2;3};{4;5;6};{7;8;9}}"
Private Function StringTodblArray(ByRef StringArray() As String) As Double()
    Dim i As Long, LB As Long, UB As Long
    Dim j As Long, LB2 As Long, UB2 As Long
    Dim tmpResult() As Double
    
    LB = LBound(StringArray)
    UB = UBound(StringArray)
    LB2 = LBound(StringArray, 2)
    UB2 = UBound(StringArray, 2)
    ReDim tmpResult(LB To UB, LB2 To UB2)
    For i = LB To UB
        For j = LB2 To UB2
            tmpResult(i, j) = CDbl(StringArray(i, j))
        Next j
    Next i
    StringTodblArray = tmpResult
End Function
''' <summary>
''' List is expected to be an array. The last element will be used as
''' the concatenation string.
''' </summary>
Public Function Concat(List As Variant) As String
    Dim startIdx As Long
    Dim endIdx As Long
    Dim i As Long
    Dim tmpResult As String
    Dim joinString As String
    
    startIdx = LBound(List)
    endIdx = UBound(List)
    joinString = MidB$(List(endIdx), 3, LenB(List(endIdx)) - 4)
    tmpResult = MidB$(List(startIdx), 3, LenB(List(startIdx)) - 4)
    For i = startIdx + 1 To endIdx - 1
        If AscW(List(i)) = 39 Then  ' [']
            tmpResult = tmpResult & _
            joinString & _
            MidB$(List(i), 3, LenB(List(i)) - 4)
        Else
            tmpResult = tmpResult & _
            joinString & _
            List(i)
        End If
    Next i
    Concat = Apostrophe & tmpResult & Apostrophe
End Function
