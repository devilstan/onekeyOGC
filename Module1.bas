Sub labelsort()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("A3:B" & lastrow).sort key1:=Range("A3:A" & lastrow), order1:=xlAscending, Header:=xlNo
End Sub
