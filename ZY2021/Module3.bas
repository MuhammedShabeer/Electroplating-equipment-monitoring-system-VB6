Attribute VB_Name = "Module3"
Public A_P As String
Public Sub lao_xin()
Dim lao_r As Integer
Dim lao_c As Integer
 min.Adodc2.Refresh
 min.MSFlexGrid2.Rows = min.Adodc2.Recordset.RecordCount + 1
 min.MSFlexGrid2.Cols = min.Adodc2.Recordset.Fields.Count
 min.MSFlexGrid2.ColWidth(0) = 1600
 min.MSFlexGrid2.ColWidth(11) = 400
 min.MSFlexGrid2.ColAlignment(0) = 1
 min.MSFlexGrid2.TextMatrix(0, 0) = min.List2.List(0)
    For i = 1 To 11
     min.MSFlexGrid2.ColWidth(i) = 600
     min.MSFlexGrid2.ColAlignment(i) = 1
     min.MSFlexGrid2.TextMatrix(0, i) = min.List2.List(i)
    Next
 lao_r = min.MSFlexGrid2.Rows - 1
 lao_c = min.MSFlexGrid2.Cols - 1
 If min.Adodc2.Recordset.EOF = False Then
 min.Adodc2.Recordset.MoveFirst
 If (min.Adodc2.Recordset.EOF = False) Then
  For i = 1 To lao_r
   For j = 0 To lao_c
    min.MSFlexGrid2.TextMatrix(i, j) = min.Adodc2.Recordset.Fields(j)
    Next j
    min.Adodc2.Recordset.MoveNext
   Next i
   End If
  Else
 Exit Sub
 End If
End Sub
Public Sub lao_FX()

End Sub
Public Sub PASS()
   A_P = InputBox("Input PWD:", "PWD")
   If A_P <> userpass Then MsgBox "PWD erroe", , "PWD": Exit Sub
End Sub
 
