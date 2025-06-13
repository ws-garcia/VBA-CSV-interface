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
' Copyright © 2022-2025 W. García
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

''' <summary>
''' sDate is expected to be an array
''' https://www.codeguru.com/ShowCode.asp?ID=7369
''' </summary>
Public Function WEEKNUM(sDate As Variant) As String
    Dim tmpDate As Date
    
    If sDate(0) Like "'*'" Then
        tmpDate = CDate(Mid(sDate(0), 2, Len(sDate(0)) - 2))
    Else
        tmpDate = CDate(sDate(0))
    End If
   'Monday is set as first day of week
   Dim lngDate As Long
   Dim intWeek As Integer

   'If january 1. is later then thursday, january 1. is not in week 1
   If Not Weekday("01/01/" & Year(tmpDate), vbMonday) > 4 Then
      intWeek = 1
   Else
      intWeek = 0
   End If
   'Sets long-value for january 1.
   lngDate = CLng(CDate("01/01/" & Year(tmpDate)))

   'Finds the first monday of year
   lngDate = lngDate + (8 - Weekday("01/01/" & Year(tmpDate), vbMonday))
   'Increases week by week until set date is passed
   While Not lngDate > CLng(tmpDate)
      intWeek = intWeek + 1
      lngDate = lngDate + 7
   Wend
   'If the date set is not in week 1, this finds latest week previous year
   If intWeek = 0 Then
      intWeek = WEEKNUM("31/12/" & Year(tmpDate) - 1)
   End If
   WEEKNUM = intWeek
End Function


