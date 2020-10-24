Attribute VB_Name = "CSVinterfaceTest"
Option Explicit
Sub Test_Import()
    Dim CSVix As CSVinterface
    Dim s As Double, e As Double, H As Double
    Dim fileName As String, i As Single
    Dim tmpCSV As String
	 
    Set CSVix = New CSVinterface
    
    fileName = "C:\Demo_400k_records.csv"
	 tmpCSV = CSVix.GetDataFromCSV(fileName)

    H = 0#
    For i = 1 To 10
        s = Timer
        Call CSVix.ImportFromCSVString(tmpCSV)
        e = Timer
        H = H + (e - s)
    Next i
    Debug.Print "CSVinterface [ImportFromCSV]:"; Round(H / 10, 4)
    Debug.Print "*********************************************************"
    Debug.Print Err.Number

    Set CSVix = Nothing
End Sub

