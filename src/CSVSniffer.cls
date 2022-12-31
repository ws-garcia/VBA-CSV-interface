VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVSniffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
' GENERAL INFO:
' This module is part of the VBA CSV interface library. It is developed as an attempt to
' sniff/guess CSV dialects without user intervention.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
Public Enum FieldDataType
    Known = 1
    UnKnown = 0
End Enum
Private Function LikeCompare(ByRef value As Variant, _
                                        ParamArray Pattern() As Variant) As Boolean
    Dim iCounter As Long
    Dim ParamLB As Long
    Dim ParamUB As Long
    Dim tmpBool As Boolean
    
    ParamLB = LBound(Pattern)
    ParamUB = UBound(Pattern)
    iCounter = ParamLB
    Do
        tmpBool = value Like Pattern(iCounter)
        iCounter = iCounter + 1
    Loop While iCounter <= ParamUB And Not tmpBool
    LikeCompare = tmpBool
End Function
''' <summary>
''' Attempts to detect the data type of a CSV field.
''' </summary>
''' <param name="value">CSV field content.</param>
Public Function DetectDataType(ByRef value As Variant) As FieldDataType
    Dim tmpDataType As FieldDataType
    
    tmpDataType = UnKnown
    If IsNumericData(value) Then
        tmpDataType = FieldDataType.Known
    Else
        If IsDateOrSpecialData(value) Then
            tmpDataType = FieldDataType.Known
        Else
            If IsStructuredOrURI(value) Then
                tmpDataType = FieldDataType.Known
            Else
                If IsFileSystemPath(value) Then
                    tmpDataType = FieldDataType.Known
                End If
            End If
        End If
    End If
    DetectDataType = tmpDataType
End Function
Private Function dmyyyyhhmmDateTime(value As Variant) As Boolean
    'Match DD/MM/YYYY[YYYY/DD/MM] and MM/DD/YYYY[YYYY/MM/DD] HH:MM
    dmyyyyhhmmDateTime = LikeCompare(value, _
                                    "##[-/.]##[-/.]####[T]##:##", _
                                    "####[-/.]##[-/.]##[T]##:##", _
                                    "#[-/.]##[-/.]####[T]##:##", _
                                    "####[-/.]#[-/.]##[T]##:##", _
                                    "#[-/.]#[-/.]####[T]##:##", _
                                    "####[-/.]#[-/.]#[T]##:##", _
                                    "##[-/.]#[-/.]####[T]##:##", _
                                    "####[-/.]##[-/.]#[T]##:##")
End Function
Private Function dmyyyyhhmmssDateTime(value As Variant) As Boolean
    'Match DD/MM/YYYY[YYYY/DD/MM] and MM/DD/YYYY[YYYY/MM/DD] HH:MM:SS
    dmyyyyhhmmssDateTime = LikeCompare(value, _
                                    "##[-/.]##[-/.]####[T]##:##:##", _
                                    "####[-/.]##[-/.]##[T]##:##:##", _
                                    "#[-/.]##[-/.]####[T]##:##:##", _
                                    "####[-/.]#[-/.]##[T]##:##:##", _
                                    "#[-/.]#[-/.]####[T]##:##:##", _
                                    "####[-/.]#[-/.]#[T]##:##:##", _
                                    "##[-/.]#[-/.]####[T]##:##:##", _
                                    "####[-/.]##[-/.]#[T]##:##:##")
End Function
Private Function dmyyyyhhmmssTStampedDateTime(value As Variant) As Boolean
    'Match DD/MM/YYYY[YYYY/DD/MM] and MM/DD/YYYY[YYYY/MM/DD] HH:MM:SS +/- HH:MM
    dmyyyyhhmmssTStampedDateTime = LikeCompare(value, _
                                            "##[-/.]##[-/.]####[T]##:##:##[+-]##:##", _
                                            "####[-/.]##[-/.]##[T]##:##:##[+-]##:##", _
                                            "#[-/.]##[-/.]####[T]##:##:##[+-]##:##", _
                                            "####[-/.]#[-/.]##[T]##:##:##[+-]##:##", _
                                            "#[-/.]#[-/.]####[T]##:##:##[+-]##:##", _
                                            "####[-/.]#[-/.]#[T]##:##:##[+-]##:##", _
                                            "##[-/.]#[-/.]####[T]##:##:##[+-]##:##", _
                                            "####[-/.]##[-/.]#[T]##:##:##[+-]##:##")
End Function
Private Function dmyyyyhhmmTStampedDateTime(value As Variant) As Boolean
    'Match DD/MM/YYYY[YYYY/DD/MM] and MM/DD/YYYY[YYYY/MM/DD] HH:MM +/- HH:MM
    dmyyyyhhmmTStampedDateTime = LikeCompare(value, _
                                            "##[-/.]##[-/.]####[T]##:##[+-]##:##", _
                                            "####[-/.]##[-/.]##[T]##:##[+-]##:##", _
                                            "#[-/.]##[-/.]####[T]##:##[+-]##:##", _
                                            "####[-/.]#[-/.]##[T]##:##[+-]##:##", _
                                            "#[-/.]#[-/.]####[T]##:##[+-]##:##", _
                                            "####[-/.]#[-/.]#[T]##:##[+-]##:##", _
                                            "##[-/.]#[-/.]####[T]##:##[+-]##:##", _
                                            "####[-/.]##[-/.]#[T]##:##[+-]##:##")
End Function
Private Function hhmmssTStampedDateTime(value As Variant) As Boolean
    'Match HH:MM:SS and HH:MM +/- 00:00
    hhmmssTStampedDateTime = LikeCompare(value, _
                                        "##:##:##[+-]##:##", _
                                        "##:##[+-]##:##")
End Function
Private Function IsAlphaNumeric(value As Variant) As Boolean
'Match ABCZ10, nullString and ABCZ_10
    Dim StrLen As Long
    Dim iCounter As Long
    Dim tmpBool As Boolean
    
    StrLen = LenB(value)
    iCounter = 1
    Do
        Select Case AscW(MidB(value, iCounter, 2))
            Case 48 To 57, 65 To 90, 95, 97 To 122
                tmpBool = True
            Case Else
                tmpBool = False
        End Select
        iCounter = iCounter + 2
    Loop While iCounter <= StrLen And tmpBool
    IsAlphaNumeric = tmpBool
End Function
Private Function IsCurrency(value As Variant) As Boolean
    If LikeCompare(value, "[$€£¥]#*[.,]##", "[$€£¥][ ]#*[.,]##") Then
        IsCurrency = IsNumeric(Format(MidB(value, 3), "#,#0.00"))
    Else
        If LikeCompare(value, "[ ][$€£¥]#*[.,]##", "[ ][$€£¥][ ]#*[.,]##") Then
            IsCurrency = IsNumeric(Format(MidB(value, 5), "#,#0.00"))
        End If
    End If
End Function
Private Function IsDateOrSpecialData(value As Variant) As Boolean
    Dim tmpBool As Boolean
    
    tmpBool = IsSpecialData(value)
    If Not tmpBool Then
        If IsDate(value) Then
            tmpBool = True
        Else
            If IsDateTime(value) Then
                tmpBool = True
            End If
        End If
    End If
    IsDateOrSpecialData = tmpBool
End Function
Private Function IsDateTime(value As Variant) As Boolean
    Dim tmpBool As Boolean
    If InStrB(1, value, ":") Then
        tmpBool = hhmmssTStampedDateTime(value)
        If Not tmpBool Then
            If InStrB(1, value, "T") Then
                tmpBool = IsLongOrStampedDateTime(value)
                If Not tmpBool Then
                    tmpBool = IsOtherDateTime(value)
                End If
            End If
        End If
    Else
        'Match [Mon, Jan 1 | Mon., Jan. 15 | Monday, January 31 | ...]
        'and [Lun, Ene 1 | Lun., Ene. 15 | Lunes, Enero 31 | ...]
        tmpBool = LikeCompare(value, _
                            "[SMTWF][uoehra][neduit]*[,][ ][JFMASOND][aepuco][nbrylgptvc]*[ ]#", _
                            "[SMTWF][uoehra][neduit]*[,][ ][JFMASOND][aepuco][nbrylgptvc]*[ ]##", _
                            "[DLMJVS][ouai][mnreb]*[,][ ][EFMAJSOND][neabugcoi][ebrynloptvc]*[ ]#", _
                            "[DLMJVS][ouai][mnreb]*[,][ ][EFMAJSOND][neabugcoi][ebrynloptvc]*[ ]##")
    End If
    IsDateTime = tmpBool
End Function
Private Function IsDotDate(value As Variant) As Boolean
    IsDotDate = LikeCompare(value, _
                            "####[.]##[.]##", _
                            "##[.]##[.]####", _
                            "####[.]##[.]#", _
                            "#[.]##[.]####", _
                            "####[.]#[.]#", _
                            "#[.]#[.]####", _
                            "####[.]#[.]##", _
                            "##[.]#[.]####")
End Function
Private Function IsEmail(value As Variant) As Boolean
    If InStrB(1, value, "@") Then
        If value Like "*[@]*[.]?*?" Then
            Dim StrLen As Long
            Dim iCounter As Long
            Dim tmpBool As Boolean
    
            StrLen = LenB(value)
            iCounter = 1
            Do
                Select Case AscW(MidB(value, iCounter, 2))
                    Case 33 To 39, 42 To 43, 45 To 57, 61, 63 To 90, 94 To 126
                        tmpBool = True
                    Case Else
                        tmpBool = False
                End Select
                iCounter = iCounter + 2
            Loop While iCounter <= StrLen And tmpBool
        End If
    End If
    IsEmail = tmpBool
End Function
Private Function IsFileSystemPath(value As Variant) As Boolean
    Dim tmpBool As Boolean
    
    If IsWindowsAbsolutePath(value) Then
        tmpBool = True
    Else
        If IsUnixAbsolutePath(value) Then
            tmpBool = True
        End If
    End If
    IsFileSystemPath = tmpBool
End Function
Private Function IsIPv4(value As Variant) As Boolean
    If value Like "*.*.*.*" Then
        IsIPv4 = IsValidIPv4(value)
    End If
End Function
Private Function IsISOdate(value As Variant) As Boolean
    'Match YYYY/MM/DDTHH:MM:SSZ and YYYY/MM/DDTHH:MM:SS[+/-]HH:MM
    IsISOdate = LikeCompare(value, _
                            "####[-/.]##[-/.]##T##:##:##Z", _
                            "####[-/.]##[-/.]##T##:##:##[+-]##:##")
End Function
Private Function IsJSfullTextDateTime(value As Variant) As Boolean
    'Match JavaScript full text date and time
    IsJSfullTextDateTime = LikeCompare(value, _
                                    "??? ??? ## #### ##:##:## *-* (*)")
End Function
Private Function IsLongOrStampedDateTime(value As Variant) As Boolean
    Dim tmpBool As Boolean
    tmpBool = IsISOdate(value)
    If Not tmpBool Then
        tmpBool = dmyyyyhhmmssDateTime(value)
        If Not tmpBool Then
            tmpBool = dmyyyyhhmmDateTime(value)
            If Not tmpBool Then
                tmpBool = IsStampedDateTime(value)
            End If
        End If
    End If
    IsLongOrStampedDateTime = tmpBool
End Function
Private Function IsNumericData(value As Variant) As Boolean
    Dim tmpBool As Boolean
    
    If IsNumeric(value) Then
        tmpBool = True
    Else
        If IsCurrency(value) Then
            tmpBool = True
        Else
            If IsPercentage(value) Then
                tmpBool = True
            End If
        End If
    End If
    IsNumericData = tmpBool
End Function
Private Function IsOtherDateTime(value As Variant) As Boolean
    Dim tmpBool As Boolean
    
    'Match YYYY/MM/DD[ ][T]HH:MM:SS.ss and MM/DD/YYYY[ ][T]HH:MM:SS.ss
    If LikeCompare(value, _
                    "####[-/.]##[-/.]##[ T]##:##:##.#*#", _
                    "##[-/.]##[-/.]####[ T]##:##:##.#*#") Then
        Dim tmpMilliSec As String
        
        tmpMilliSec = Split(value, ".")(1)
        tmpBool = IsNumeric(tmpMilliSec)
    Else
        'Match [Mon Dec 31 01:41:00 2001 | Mon Dec 1 01:41:00 2001]
        'and [Mon Dec 31 01:41:00 21 | Mon Dec 1 01:41:00 21]
        tmpBool = LikeCompare(value, _
                            "[SMTWF][uoehra][neduit][ ][JFMASOND][aepuco][nbrylgptvc][ ]##[ ]##:##:##[ ]####", _
                            "[SMTWF][uoehra][neduit][ ][JFMASOND][aepuco][nbrylgptvc][ ]#[ ]##:##:##[ ]####", _
                            "[SMTWF][uoehra][neduit][ ][JFMASOND][aepuco][nbrylgptvc][ ]##[ ]##:##:##[ ]##", _
                            "[SMTWF][uoehra][neduit][ ][JFMASOND][aepuco][nbrylgptvc][ ]#[ ]##:##:##[ ]##")
        If Not tmpBool Then
            tmpBool = IsSpanishDate(value)
            'Match [11.01.2001 | 2001.01.11]
            If Not tmpBool Then
                tmpBool = IsDotDate(value)
            End If
        End If
    End If
    IsOtherDateTime = tmpBool
End Function
Private Function IsPercentage(value As Variant) As Boolean
    If LikeCompare(value, "*#[%]") Then
        IsPercentage = IsNumeric(Format(MidB(value, 1, LenB(value) - 2), "#,#0.00"))
    End If
End Function
Private Function IsSpanishDate(value As Variant) As Boolean
    'Match [Lun Dic 31 01:41:00 2001 | Lun Dic 1 01:41:00 2001]
    'and [Lun Dic 31 01:41:00 21 | Lun Dic 1 01:41:00 21]
    IsSpanishDate = LikeCompare(value, _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]##[ ]##:##:##[ ]####", _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]#[ ]##:##:##[ ]####", _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]##[ ]##:##:##[ ]##", _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]#[ ]##:##:##[ ]##")
End Function
Private Function IsSpecialData(value As Variant) As Boolean
    Dim tmpBool As Boolean
    
    If LenB(value) = 0 Then
        tmpBool = True
    Else
        If IsAlphaNumeric(value) Then
            tmpBool = True
        Else
            If LCase(value) = "n/a" Then
                tmpBool = True
            End If
        End If
    End If
    IsSpecialData = tmpBool
End Function
Private Function IsStampedDateTime(value As Variant) As Boolean
    Dim tmpBool As Boolean
    tmpBool = dmyyyyhhmmssTStampedDateTime(value)
    If Not tmpBool Then
        tmpBool = dmyyyyhhmmTStampedDateTime(value)
        If Not tmpBool Then
            tmpBool = IsJSfullTextDateTime(value)
        End If
    End If
    IsStampedDateTime = tmpBool
End Function
Private Function IsStructuredData(value As Variant) As Boolean
    Dim tmpBool As Boolean
    If InStrB(1, value, "[") Then
        If LikeCompare(value, "[[]*]") Then
            tmpBool = True
        End If
    Else
        If InStrB(1, value, "(") Then
            If LikeCompare(value, "(*)") Then
                tmpBool = True
            ElseIf LikeCompare(value, "bytearray(*)") Then
                tmpBool = True
            End If
            If InStrB(1, value, "{") Then
                If LikeCompare(value, "frozenset({*})") Then
                    tmpBool = True
                End If
            End If
        Else
            If InStrB(1, value, "{") Then
                If LikeCompare(value, "{*}") Then
                    tmpBool = True
                End If
            End If
        End If
    End If
    IsStructuredData = tmpBool
End Function
Private Function IsStructuredOrURI(value As Variant) As Boolean
    Dim tmpBool As Boolean
    
    If IsStructuredData(value) Then
        tmpBool = True
    Else
        If IsEmail(value) Then
            tmpBool = True
        Else
            If IsURL(value) Then
                tmpBool = True
            Else
                If IsIPv4(value) Then
                    tmpBool = True
                End If
            End If
        End If
    End If
    IsStructuredOrURI = tmpBool
End Function
Private Function IsUnixAbsolutePath(value As Variant) As Boolean
    IsUnixAbsolutePath = LikeCompare(value, "/*")
End Function
Private Function IsURL(value As Variant) As Boolean
    If InStrB(1, value, "://") Then
        If value Like "[a-z][a-z]*[a-z]://*" Then
            If value Like "http://*" Or value Like "https://*" _
            Or value Like "ftp://*" Or value Like "ftps://*" Then
                Dim StrLen As Long
                Dim iCounter As Long
                Dim tmpBool As Boolean
    
                StrLen = LenB(value)
                iCounter = 1
                Do
                    Select Case AscW(MidB(value, iCounter, 2))
                        Case 33, 35 To 59, 61, 63 To 91, 93, 95, 97 To 122, 126
                            tmpBool = True
                        Case Else
                            tmpBool = False
                    End Select
                    iCounter = iCounter + 2
                Loop While iCounter <= StrLen And tmpBool
            End If
        End If
    End If
    IsURL = tmpBool
End Function
Private Function IsValidIPv4(value As Variant) As Boolean
    Dim tmpData() As String
    tmpData() = Split(value, ".")
    If UBound(tmpData) - LBound(tmpData) + 1 = 4 Then
        Dim StrLen As Long
        Dim iCounter As Long
        Dim tmpBool As Boolean
    
        StrLen = LenB(value)
        iCounter = 1
        Do
            Select Case AscW(MidB(value, iCounter, 2))
                Case 46, 48 To 57
                    tmpBool = True
                Case Else
                    tmpBool = False
            End Select
            iCounter = iCounter + 2
        Loop While iCounter <= StrLen And tmpBool
        If tmpBool Then
            IsValidIPv4 = IsValidIPv4Range(tmpData)
        Else
            IsValidIPv4 = False
        End If
    End If
End Function
Private Function IsValidIPv4Range(valuesArray As Variant) As Boolean
    Dim iCounter As Long
    Dim tmpBool As Boolean
    
    iCounter = LBound(valuesArray)
    Do
        If CInt(valuesArray(iCounter)) >= 0 And CInt(valuesArray(iCounter)) <= 255 Then
            tmpBool = True
        Else
            tmpBool = False
        End If
        iCounter = iCounter + 1
    Loop While iCounter <= UBound(valuesArray) And tmpBool
    IsValidIPv4Range = tmpBool
End Function
Private Function IsWindowsAbsolutePath(value As Variant) As Boolean
    IsWindowsAbsolutePath = LikeCompare(value, "[A-Za-z]:\*")
End Function
Private Function RecordsAvgFields(ArrayList As CSVArrayList) As Double
    Dim CumulativeFields As Double
    Dim L0 As Long
    
    CumulativeFields = 0
    For L0 = 0 To ArrayList.count - 1
        CumulativeFields = CumulativeFields + UBound(ArrayList(L0)) - _
                            LBound(ArrayList(L0)) + 1
    Next L0
    RecordsAvgFields = CumulativeFields / ArrayList.count
End Function
''' <summary>
''' Calculates the score for a record array as an average of the
''' fields scores. The score for single Known data type field
''' is within the range 0 < x <= 100, and within the range
''' 0 < x <= 20 for single UnKnown data type field.
''' </summary>
''' <param name="strArray">One-dimensional record array.</param>
Private Function RecordScore(ByRef strArray As Variant) As Double
    Dim L0 As Long
    Dim tmpSUM As Double
    Dim FielCount As Long
    
    FielCount = 1 + UBound(strArray) - LBound(strArray)
    tmpSUM = 0
    For L0 = LBound(strArray) To UBound(strArray)
        Select Case DetectDataType(strArray(L0))
            Case FieldDataType.Known
                tmpSUM = tmpSUM + 100
            Case Else
                tmpSUM = tmpSUM + 20
        End Select
    Next L0
    RecordScore = tmpSUM / FielCount
End Function
''' <summary>
''' Calculates a factor for table scoring based in the standard
''' deviation of the number of fields contained in the specified
''' array list.
''' </summary>
''' <param name="ArrayList">CSV array list.</param>
Private Function RecordsConsistencyFactor(ArrayList As CSVArrayList) As Double
    Dim AvgFields As Double
    Dim CumulativeDiff As Double
    Dim L0 As Long
    Dim tmpResult As Double
    
    AvgFields = RecordsAvgFields(ArrayList)
    For L0 = 0 To ArrayList.count - 1
        CumulativeDiff = CumulativeDiff + ((UBound(ArrayList(L0)) + 1) - _
                        (LBound(ArrayList(L0)) + AvgFields)) ^ 2
    Next L0
    If ArrayList.count - 1 Then
        tmpResult = (CumulativeDiff / (ArrayList.count - 1))
    Else
        tmpResult = (CumulativeDiff / ArrayList.count)
    End If
    RecordsConsistencyFactor = 1 / (1 + tmpResult ^ 0.5)
End Function
''' <summary>
''' Calculates a score for the imported data based on the congruence
''' of the detected data and the uniformity of the fields contained
''' in each record. The score is in the range 0 < x <= 100.
''' </summary>
''' <param name="ArrayList">CSV imported data.</param>
Public Function TableScore(ByRef ArrayList As CSVArrayList) As Double
    If Not ArrayList Is Nothing Then
        If ArrayList.count > 0 Then
            Dim L0 As Long
            Dim SumRecScores As Double
            
            SumRecScores = 0
            For L0 = 0 To ArrayList.count - 1
                SumRecScores = SumRecScores + RecordScore(ArrayList(L0))
            Next L0
            TableScore = RecordsConsistencyFactor(ArrayList) * SumRecScores / ArrayList.count
        End If
    End If
End Function
