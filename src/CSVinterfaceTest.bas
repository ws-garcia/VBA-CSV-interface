Attribute VB_Name = "CSVinterfaceTest"
Option Explicit
Sub Test_Import()
    Dim CSVix As CSVinterface
    Dim s As Double, e As Double, H As Double
    Dim fileName As String, i As Single

    Set CSVix = New CSVinterface
    
    fileName = ThisWorkbook.path & "\Demo_100000records.csv"

    H = 0#
    For i = 1 To 10
        s = Timer
        Call CSVix.OpenConnection(fileName)
        Call CSVix.ImportFromCSV
        e = Timer
        H = H + (e - s)
    Next i
    Debug.Print "CSVinterface [ImportFromCSV]:"; Round(H / 10, 4)
    Debug.Print "*********************************************************"
    Debug.Print Err.Number

    Set CSVix = Nothing
End Sub

