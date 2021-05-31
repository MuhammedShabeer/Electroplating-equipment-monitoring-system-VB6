Attribute VB_Name = "Module2"
Public Sub input_o1()
Dim dr01 As String
Dim DR02 As Integer
For j = 1 To 10
  If Len(D_NAO(j)) > 4 Then
    For i = 1 To 5
   If Mid(D_NAO(j), i, 1) = "." Then Mid(D_NAO(j), i, 1) = Chr(9)
    Next i
  End If
  a(j) = Format(Val(D_NAO(j)), "0000")
 Next j
 Input_a = a(1) & a(2) & a(3) & a(4) & a(5) & a(6) & a(7) & a(8) & a(9) & a(10)
   For i = 1 To min.MSFlexGrid2.Rows
       min.MSFlexGrid2.Row = i
       min.MSFlexGrid2.Col = 0
   If min.MSFlexGrid2.Text = D_NAO(0) Then
       DR02 = Val(min.MSFlexGrid2.Row)
       dr = Trim(min.MSFlexGrid2.TextMatrix(DR02, 11))
       GoTo INPUT_out01
       Exit For
   End If
   Next i
INPUT_out01:
   If Len(dr) = 1 Then dr01 = 0 & dr Else dr01 = dr
   Input_dr = Trim(Str(3500 + 10 * dr))
   Input_b = Trim(Add_liao.Text2.Text)
    o(3) = "@00WD" & Input_dr & Input_a
    o(4) = "@00WD" & Trim(pindex) & Input_b & Trim(dr01)
    o(5) = ""
    a12 = 4
End Sub
Public Sub REGEDIT()
 intport = Val(GetSetting("zxdms", "settings", "port", "1"))
 setting = GetSetting("zxdms", "settings", "setting", "9600,e,7,2")
 serialno = GetSetting("zxdms", "settings", "SN", "5091")
 SaveSetting "zxdms", "settings", "port", intport
 SaveSetting "zxdms", "settings", "setting", setting
 SaveSetting "zxdms", "settings", "SN", serialno
End Sub
 Public Sub input_index()
 If Not TIN_01 Then   ' = False
    pindex = 4050 + MA
 Else
   'If TIN_01 = True Then
      If MA > 8 And MA < 13 Then
        pindex = 4050 + MA
      ElseIf MA > 22 Then
        pindex = 4050 + MA
      Else
        pindex = 4100 + MA
      End If
    End If
  'End If
 End Sub
Public Function hex_bin()
Select Case hn
    Case "F"
       hnn = "1111"
    Case "E"
       hnn = "1110"
    Case "D"
       hnn = "1101"
    Case "C"
       hnn = "1100"
    Case "B"
       hnn = "1011"
    Case "A"
       hnn = "1010"
    Case "9"
       hnn = "1001"
    Case "8"
       hnn = "1000"
    Case "7"
       hnn = "0111"
    Case "6"
       hnn = "0110"
    Case "5"
       hnn = "0101"
    Case "4"
       hnn = "0100"
    Case "3"
       hnn = "0011"
    Case "2"
       hnn = "0010"
    Case "1"
       hnn = "0001"
    Case "0"
      hnn = "0000"
End Select
Dim sf
For i = 0 To 3
  sf = Mid(hnn, (i + 1), 1)
  If sf = 1 Then bn(i) = True Else bn(i) = False
Next i
End Function

