'' DataTemplateInsert Macro

Sub DataTemplateInsert()
'
'  Macro
'

'' workbook1

    Dim xRet As Boolean
    xRet = IsWorkBookOpen("workbook1.csv")
    If xRet Then
        
    'Run code 

    Else
    
    'Run other code

    End If
 

'' workbook2 monday

    Dim xRet1 As Boolean
    xRet1 = IsWorkBookOpen("workbook2.csv")
    If xRet1 Then
        
    'Run code 2
    

    Else

    'run other code 2'

    End If


End Sub

Sub code()

' Find 'text' and Delete 'text' row

'    Rows(6).Delete

    Last = Cells(Rows.Count, "D").End(xlUp).Row
    For i = Last To 1 Step -1
        If (Cells(i, "A").Value) = "text" Then
    'Cells(i, "A").EntireRow.ClearContents ' USE THIS TO CLEAR CONTENTS BUT NOT DELETE ROW
            Cells(i, "A").EntireRow.Delete
        End If
    Next i


        
    Range("B7").FormulaR1C1 = _
        "=SUM(SUMIFS(C,C1,{""Row1"",""Row2"",""Row3""}))"
    Range("C7").FormulaR1C1 = _
        "=AVERAGE(INDEX(C,MATCH(""Row1"",C1, 0)),INDEX(C,MATCH(""Row2"",C1, 0)), INDEX(C,MATCH(""Row3"",C1, 0)))"
    Range("D7").FormulaR1C1 = _
        "=MAX(INDEX(C,MATCH(""Row1"",C1, 0)),INDEX(C,MATCH(""Row2"",C1, 0)), INDEX(C,MATCH(""Row3"",C1, 0)))"

End Sub


Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function
 

Sub Sample()
    Dim xRet As Boolean
    xRet = IsWorkBookOpen("workbook.csv")
    If xRet Then
        MsgBox "The file is open"
    Else
        MsgBox "The file is not open"
    End If
End Sub
