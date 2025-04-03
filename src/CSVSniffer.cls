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
' Copyright © 2022-2025 W. García
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
Private Function LikeCompare(ByRef value As String, _
                                        ParamArray Pattern() As Variant) As Boolean
    Dim iCounter As Long
    Dim ParamLB As Long
    Dim ParamUB As Long
    Dim tmpBool As Boolean
    
    ParamLB = LBound(Pattern)
    ParamUB = UBound(Pattern)
    iCounter = ParamLB
    Do
        tmpBool = value Like CStr(Pattern(iCounter))
        iCounter = iCounter + 1
    Loop While iCounter <= ParamUB And Not tmpBool
    LikeCompare = tmpBool
End Function
''' <summary>
''' Attempts to detect the data type of a CSV field.
''' </summary>
''' <param name="value">CSV field content.</param>
Public Function DetectDataType(ByRef value As String) As FieldDataType
    Dim tmpDataType As FieldDataType
    Dim tmpValue As String
    
    tmpValue = Trim(value)
    tmpDataType = UnKnown
    If IsNumericData(tmpValue) Then
        tmpDataType = FieldDataType.Known
    Else
        If IsDateOrSpecialData(tmpValue) Then
            tmpDataType = FieldDataType.Known
        Else
            If IsStructuredOrURI(tmpValue) Then
                tmpDataType = FieldDataType.Known
            Else
                If IsFileSystemPath(tmpValue) Then
                    tmpDataType = FieldDataType.Known
                End If
            End If
        End If
    End If
    DetectDataType = tmpDataType
End Function
Private Function dmyyyyhhmmDateTime(value As String) As Boolean
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
Private Function dmyyyyhhmmssDateTime(value As String) As Boolean
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
Private Function dmyyyyhhmmssTStampedDateTime(value As String) As Boolean
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
Private Function dmyyyyhhmmTStampedDateTime(value As String) As Boolean
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
Private Function hhmmssTStampedDateTime(value As String) As Boolean
    'Match HH:MM:SS and HH:MM +/- 00:00
    hhmmssTStampedDateTime = LikeCompare(value, _
                                        "##:##:##[+-]##:##", _
                                        "##:##[+-]##:##")
End Function
Private Function IsAlphaNumeric(value As String) As Boolean
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
Private Function IsSpacedField(value As String) As Boolean
'Match ABCZ 10, nullString, ABCZ_10, [SPACE][,][.][;][:][!][?][']
    Dim StrLen As Long
    Dim iCounter As Long
    Dim tmpBool As Boolean
    
    StrLen = LenB(value)
    iCounter = 1
    Do
        Select Case AscW(MidB(value, iCounter, 2))
            Case 32, 48 To 57, 65 To 90, 95, 97 To 122
                tmpBool = True
            Case Else
                tmpBool = False
        End Select
        iCounter = iCounter + 2
    Loop While iCounter <= StrLen And tmpBool
    IsSpacedField = tmpBool
End Function
Private Function IsCurrency(value As String) As Boolean
    If LikeCompare(value, "[$€£¥]#*[.,]##", "[$€£¥][ ]#*[.,]##") Then
        IsCurrency = IsNumeric(Format(MidB(value, 3), "#,#0.00"))
    Else
        If LikeCompare(value, "[ ][$€£¥]#*[.,]##", "[ ][$€£¥][ ]#*[.,]##") Then
            IsCurrency = IsNumeric(Format(MidB(value, 5), "#,#0.00"))
        End If
    End If
End Function
Private Function IsDateOrSpecialData(value As String) As Boolean
    Dim tmpBool As Boolean
    
    tmpBool = IsSpecialData(value)
    If Not tmpBool Then
        tmpBool = IsDate(value)
        If Not tmpBool Then
            tmpBool = IsDateTime(value)
            If Not tmpBool Then
                tmpBool = IsSpacedField(value)
            End If
        End If
    End If
    IsDateOrSpecialData = tmpBool
End Function
Private Function IsDateTime(value As String) As Boolean
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
Private Function IsDotDate(value As String) As Boolean
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
Private Function IsEmail(value As String) As Boolean
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
Private Function IsFileSystemPath(value As String) As Boolean
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
Private Function IsIPv4(value As String) As Boolean
    If value Like "*.*.*.*" Then
        IsIPv4 = IsValidIPv4(value)
    End If
End Function
Private Function IsISOdate(value As String) As Boolean
    'Match YYYY/MM/DDTHH:MM:SSZ and YYYY/MM/DDTHH:MM:SS[+/-]HH:MM
    IsISOdate = LikeCompare(value, _
                            "####[-/.]##[-/.]##T##:##:##Z", _
                            "####[-/.]##[-/.]##T##:##:##[+-]##:##")
End Function
Private Function IsJSfullTextDateTime(value As String) As Boolean
    'Match JavaScript full text date and time
    IsJSfullTextDateTime = LikeCompare(value, _
                                    "??? ??? ## #### ##:##:## *-* (*)")
End Function
Private Function IsLongOrStampedDateTime(value As String) As Boolean
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
Private Function IsNumericData(value As String) As Boolean
    Dim tmpBool As Boolean
    
    tmpBool = IsNumeric(value)
    If Not tmpBool Then
        tmpBool = IsCurrency(value)
        If Not tmpBool Then
            tmpBool = IsPercentage(value)
            If Not tmpBool Then
                tmpBool = IsNumericList(value)
            End If
        End If
    End If
    IsNumericData = tmpBool
End Function
Private Function IsNumericList(value As String) As Boolean
    Dim tmpSep As String
    Dim i As Long
    Dim n As Long
    Dim tmpChr As String
    Dim tmpResult As Boolean
    
    n = Len(value)
    'Find the first character that is not a number
    Do
        i = i + 1
        tmpChr = Mid$(value, i, 1)
    Loop While Not tmpChr Like "[!0-9.]" And i < n
    tmpSep = tmpChr
    'Check if all not numeric characters are the same
    Do
        Do
            i = i + 1
            tmpChr = Mid$(value, i, 1)
        Loop While tmpChr Like "[0-9.]" And i < n
        If i < n Then
            tmpResult = (tmpChr = tmpSep)
        End If
    Loop While tmpResult And i < n
    IsNumericList = tmpResult
End Function
Private Function IsOtherDateTime(value As String) As Boolean
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
Private Function IsPercentage(value As String) As Boolean
    If LikeCompare(value, "*#[%]") Then
        IsPercentage = IsNumeric(Format(MidB(value, 1, LenB(value) - 2), "#,#0.00"))
    End If
End Function
Private Function IsSpanishDate(value As String) As Boolean
    'Match [Lun Dic 31 01:41:00 2001 | Lun Dic 1 01:41:00 2001]
    'and [Lun Dic 31 01:41:00 21 | Lun Dic 1 01:41:00 21]
    IsSpanishDate = LikeCompare(value, _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]##[ ]##:##:##[ ]####", _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]#[ ]##:##:##[ ]####", _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]##[ ]##:##:##[ ]##", _
                                "[DLMJVS][ouai][mnreb][ ][EFMAJSOND][neabugcoi][ebrynloptvc][ ]#[ ]##:##:##[ ]##")
End Function
Private Function IsSpecialData(value As String) As Boolean
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
Private Function IsStampedDateTime(value As String) As Boolean
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
Private Function IsStructuredData(value As String) As Boolean
    Dim tmpBool As Boolean
    Dim tmpDataPos As Long
    
    If InStrB(1, value, "[") Then
        If LikeCompare(value, "[[]*]") Then
            tmpBool = True
        End If
    Else
        tmpDataPos = InStrB(1, value, "(")
        If tmpDataPos Then
            If LikeCompare(value, "(*)") Then
                tmpBool = True
            ElseIf LikeCompare(value, "[A-Za-z]*[A-Za-z](*)") Then
                tmpBool = IsAlphaNumeric(MidB$(value, 1, tmpDataPos - 1))
            End If
            tmpDataPos = InStrB(1, value, "{")
            If tmpDataPos Then
                If LikeCompare(value, "[A-Za-z]*[A-Za-z]({*})") Then
                    tmpBool = IsAlphaNumeric(MidB$(value, 1, tmpDataPos - 1))
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
Private Function IsStructuredOrURI(value As String) As Boolean
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
Private Function IsUnixAbsolutePath(value As String) As Boolean
    IsUnixAbsolutePath = LikeCompare(value, "/[A-Za-z]*")
End Function
Private Function IsURL(value As String) As Boolean
    Dim tmpValue As String
    
    tmpValue = LCase(value)
    If InStrB(1, tmpValue, "://") Then
        If tmpValue Like "[a-z][a-z]*[a-z]://*" Then
            If tmpValue Like "http://*" Or tmpValue Like "https://*" _
            Or tmpValue Like "ftp://*" Or tmpValue Like "ftps://*" Then
                Dim StrLen As Long
                Dim iCounter As Long
                Dim tmpBool As Boolean
    
                StrLen = LenB(tmpValue)
                iCounter = 1
                Do
                    Select Case AscW(MidB(tmpValue, iCounter, 2))
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
Private Function IsValidIPv4(value As String) As Boolean
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
Private Function IsValidIPv4Range(valuesArray() As String) As Boolean
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
Private Function IsWindowsAbsolutePath(value As String) As Boolean
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
    Dim FieldsCount As Long
    
    FieldsCount = 1 + UBound(strArray) - LBound(strArray)
    tmpSUM = 0
    For L0 = LBound(strArray) To UBound(strArray)
        Select Case DetectDataType(CStr(strArray(L0)))
            Case FieldDataType.Known
                tmpSUM = tmpSUM + 100
            Case Else
                tmpSUM = tmpSUM + 0.1
        End Select
    Next L0
    RecordScore = (tmpSUM ^ 2) / (100 * FieldsCount ^ 2)
End Function
''' <summary>
''' Calculates a factor for table scoring based in the standard
''' deviation of the number of fields contained in the specified
''' array list.
''' </summary>
''' <param name="ArrayList">CSV array list.</param>
Private Function TUniformity(ArrayList As CSVArrayList) As Double()
    Dim AvgFields As Double
    Dim CumulativeDiff As Double
    Dim L0 As Long, L1 As Long
    Dim SIGMAsqr As Double
    Dim MaxFldCount As Long
    Dim MinFldCount As Long
    Dim LB As Long
    Dim UB As Long
    Dim FldCount As Long
    Dim tmpDblArr() As Double
    Dim FldStructChanges As Long
    Dim prevRecFldCount As Long
    Dim modeCounter As Long
    Dim modeCount As Long
    Dim mF As Double
    Dim fRange As Long
    
    ReDim tmpDblArr(0 To 1)
    AvgFields = RecordsAvgFields(ArrayList)
    L1 = ArrayList.count - 1
    For L0 = 0 To L1
        UB = UBound(ArrayList(L0))
        LB = LBound(ArrayList(L0))
        CumulativeDiff = CumulativeDiff + ((UB + 1) - _
                        (LB + AvgFields)) ^ 2
        FldCount = UB - LB + 1
        If L0 = 0 Then
            prevRecFldCount = FldCount
            modeCounter = modeCounter + 1
        Else
            If prevRecFldCount <> FldCount Then
                prevRecFldCount = FldCount
                FldStructChanges = FldStructChanges + 1
                If modeCounter > modeCount Then
                    modeCount = modeCounter
                End If
                modeCounter = 0
            Else
                modeCounter = modeCounter + 1
                If L0 = L1 Then
                    If modeCounter > modeCount Then
                        modeCount = modeCounter
                    End If
                End If
            End If
        End If
        If MinFldCount = 0 Then
            MinFldCount = FldCount
            MaxFldCount = FldCount
        Else
            If FldCount < MinFldCount Then
                MinFldCount = FldCount
            End If
            If FldCount > MaxFldCount Then
                MaxFldCount = FldCount
            End If
        End If
    Next L0
    fRange = MaxFldCount - MinFldCount
    If ArrayList.count - 1 Then
        SIGMAsqr = (CumulativeDiff / (L1))
    Else
        SIGMAsqr = (CumulativeDiff / (L1 + 1))
    End If
    tmpDblArr(0) = (1 / (1 + 2 * SIGMAsqr ^ 0.5))
    If FldStructChanges Then  'Non uniform table
        mF = modeCount / (L1 + 1)
    End If
    tmpDblArr(1) = 2 * (FldStructChanges ^ 2 + 1) * fRange * (1 - mF) / modeCount
    TUniformity = tmpDblArr
End Function
''' <summary>
''' Calculates a score for the imported data based on the congruence
''' of the detected data and the uniformity of the fields contained
''' in each record. The score is in the range 0 < x <= 200.
''' </summary>
''' <param name="ArrayList">CSV imported data.</param>
Public Function TableScore(ByRef ArrayList As CSVArrayList, Threshold As Long) As Double
    If Not ArrayList Is Nothing Then
        If ArrayList.count > 0 Then
            Dim L0 As Long
            Dim SumRecScores As Double
            Dim R() As Double
            Dim UB As Long, LB As Long
            Dim R1 As Double
            Dim R0 As Double
            
            SumRecScores = 0
            For L0 = 0 To ArrayList.count - 1
                SumRecScores = SumRecScores + RecordScore(ArrayList(L0))
            Next L0
            If ArrayList.count > 1 Then
                R = TUniformity(ArrayList)
                TableScore = SumRecScores * ((R(0) / Threshold) + (1 / (R(1) + ArrayList.count)))
            Else
                UB = UBound(ArrayList(0))
                LB = LBound(ArrayList(0))
                R0 = (UB - LB + 1)
                R1 = (SumRecScores ^ 0.5) / 10
                'Select multi-record table over single-record one
                TableScore = SumRecScores * (R1 + (1 / (R0))) / (R0 - Fix(R1 * R0) + 1)
            End If
        End If
    End If
End Function
