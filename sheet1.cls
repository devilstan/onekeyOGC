#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Private Sub onekeyOGC_Click()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim i As Integer
    Dim lastrow As Long
    Dim lastrow2 As Long
    Dim str_tmp As String
    Dim label_rng As Range
    Dim find_result As Variant
    
    'clear first
    Dim rowcnt As Integer
    rowcnt = UsedRange.Rows.Count
    Range(Cells(3, 1), Cells(rowcnt, 4)).Select
    Selection.Clear
    
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(Range("B1").Cells.Value)
    With Worksheets("label")
        lastrow = .Cells(Rows.Count, 2).End(xlUp).Row
        Set label_rng = .Range("A2:A" & lastrow)
    End With
    i = 2
    'loops through each file in the directory and prints their names and path
    For Each objSubFolder In objFolder.subfolders
        DoEvents
        'print folder name
        Cells(i + 1, 1) = objSubFolder.Name
        Set find_result = label_rng.Find(What:=Split(Cells(i + 1, 1), "-")(0), After:=label_rng(1), LookIn:=xlFormulas, _
                                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=True, SearchFormat:=False)
        If find_result Is Nothing Then
            Cells(i + 1, 2) = namechk(objSubFolder.Name)
            'print folder path
            Cells(i + 1, 3) = objSubFolder.Path
            If Cells(i + 1, 2) <> 0 Then
                If InStr(1, Cells(i + 1, 1), Cells(i + 1, 2)) <> 1 Then
                    str_tmp = Range("B1").Cells.Value & "\" & Cells(i + 1, 2)
                    On Error Resume Next
retry_rename:
                    Name Cells(i + 1, 3) As str_tmp
                    If Err.Number = 5 Then
                        str_tmp = str_tmp & "(1)"
                        Err.Clear
                        GoTo retry_rename
                    End If
                    On Error GoTo 0
                End If
            Else
                'Cells(i + 1, 2) = Cells(i + 1, 1)
            End If
        Else
            Cells(i + 1, 2) = Cells(i + 1, 1)
            Cells(i + 1, 3) = objSubFolder.Path
        End If
        i = i + 1
        Sleep 30
    Next objSubFolder
    
    
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("A3:C" & lastrow).sort key1:=Range("B3:B" & lastrow), order1:=xlAscending, Header:=xlNo
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row + 1
    lastrow2 = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & lastrow & ":C" & lastrow2).sort key1:=Range("A" & lastrow & ":A" & lastrow2), order1:=xlAscending, Header:=xlNo

End Sub

Private Function namechk(input_str As String) As String

    Dim rng As Range
    Dim rowcnt As Integer
    Dim lpos As Integer, npos As Integer
    Dim thelabel As String, theNO As String, theNOdigits As Integer
    rowcnt = Worksheets("label").UsedRange.Rows.Count
    
    
    With Worksheets("label")
        Set rng = .Range(.Cells(2, 1), .Cells(rowcnt, 1))
    End With

    For Each lbl In rng
        If lbl.Cells.Value <> "" Then
            thelabel = lbl.Cells.Value
            theNOdigits = Val(lbl.Offset(0, 1).Cells.Value)
            
            lpos = InStr(UCase(input_str), UCase(thelabel))
            If lpos > 0 Then
                '§ä¨ì¼tµP¦r¦ê
                Dim ccnt As Integer
                ccnt = 0
                For c = lpos + Len(lbl.Cells.Value) To Len(input_str)
                    If Mid(input_str, c, 1) Like "[0123456789]" Then
                        ccnt = ccnt + 1
                        If ccnt = theNOdigits Then
                            theNO = Mid(input_str, c - theNOdigits + 1, theNOdigits)
                            If UCase(Mid(input_str, c + 1, 1)) = "R" Then
                                theNO = theNO + "R"
                            ElseIf UCase(Mid(input_str, c + 1, 2)) = "_C" Then
                                theNO = theNO + "_C"
                            End If
                            namechk = UCase(thelabel) & "-" & theNO
                            Exit Function
                        End If
                    Else
                        ccnt = 0
                    End If
                Next
            End If
        End If
    Next lbl
    namechk = ""

End Function


