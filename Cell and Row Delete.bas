Attribute VB_Name = "Module6"
Sub KeepRowsWithInfoInColumnC111()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
       
    Set ws = Workbooks("Copy of Master_Payroll Slip Automation CFCF.xlsm").Sheets("OT Report 3")
   
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
   
    For i = lastRow To 3 Step -1 ' Start from the last row and loop upwards
        If IsEmpty(ws.Cells(i, "C").Value) Then ' Check if cell in column C is empty
            ws.Rows(i).Delete ' Delete the entire row
        End If
    Next i
   
   
    ws.Rows("1:2").Delete
   
   
    ws.Columns("A:B").Delete
ws.Columns("I:M").Delete
   
    ' Delete the last row
    ws.Rows(lastRow).Delete
   
   

   
    ' Find the last row with data in Column A
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
   
    ' Delete the row after the last row with data in Column A
    If lastRowA < ws.Cells(ws.Rows.Count, "C").End(xlUp).Row Then
        ws.Rows(lastRowA + 1).Delete
    End If
   
End Sub
