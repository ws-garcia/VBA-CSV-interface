Attribute VB_Name = "CSVfileManagerTest"
Option Explicit
Sub Test_Import()
    Dim CSVmanager As CSVfileManager
    Dim s As Double, e As Double, H As Double
    Dim fileName As String, i As Single

    Set CSVmanager = New CSVfileManager
    
    fileName = ThisWorkbook.path & "\Demo_100000records.csv"

    H = 0#
    For i = 1 To 10
        s = Timer
        Call CSVmanager.OpenConnection(fileName)
        Call CSVmanager.ImportFromCSV
        e = Timer
        H = H + (e - s)
    Next i
    Debug.Print "CSVfileManager [ImportFromCSV]:"; Round(H / 10, 4)
    Debug.Print "*********************************************************"
    Debug.Print Err.Number

    Set CSVmanager = Nothing
End Sub
