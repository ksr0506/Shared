Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String
    On Error GoTo Exception

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        mask = Replace(mask, "{" & i & "}", tokens(i))
    Next
    StringFormat = mask

Exit Function
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : StringFormat][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Function
End Function

Public Function ConvertToDate(ByRef src As String, ByRef dateFormat As String) As Variant
    On Error GoTo Exception

    Dim y As Long, m As Long, d As Long, h As Long, min As Long, s As Long
    Dim am As Boolean, pm As Boolean
    Dim pos As Long

    If Len(src) <> Len(dateFormat) Then
        ConvertToDate = CVErr(xlErrNA)  ' #N/A error
        Exit Function
    End If

    pos = InStr(1, dateFormat, "yyyy", vbTextCompare)
    If pos > 0 Then
        y = Val(Mid(src, pos, 4))
    Else: pos = InStr(1, dateFormat, "yy", vbTextCompare)
        If pos > 0 Then
            y = Val(Mid(src, pos, 2))
            If y < 80 Then y = y + 2000 Else y = y + 1900
        End If
    End If

    pos = InStr(1, dateFormat, "mmm", vbTextCompare)
    If pos > 0 Then
        m = Month(DateValue("01 " & (Mid(src, pos, 3)) & " 2000"))
    Else: pos = InStr(1, dateFormat, "MM", vbBinaryCompare)
        If pos > 0 Then m = Val(Mid(src, pos, 2))
    End If

    pos = InStr(1, dateFormat, "dd", vbTextCompare)
    If pos > 0 Then d = Val(Mid(src, pos, 2))

    pos = InStr(1, dateFormat, "hh", vbTextCompare)
    If pos > 0 Then h = Val(Mid(src, pos, 2))
    If InStr(1, src, "am", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "a.m.", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "a. m.", vbTextCompare) > 0 Then am = True
    If InStr(1, src, "pm", vbTextCompare) > 0 Then pm = True
    If InStr(1, src, "p.m.", vbTextCompare) > 0 Then pm = True
    If InStr(1, src, "p. m.", vbTextCompare) > 0 Then pm = True
    If am And h = 12 Then h = 0
    If pm And h <> 12 Then h = h + 12

    pos = InStr(1, dateFormat, "mm", vbBinaryCompare)
    If pos > 0 Then min = Val(Mid(src, pos, 2))

    pos = InStr(1, dateFormat, "ss", vbTextCompare)
    If pos > 0 Then s = Val(Mid(src, pos, 2))

    ConvertToDate = DateSerial(y, m, d) + TimeSerial(h, min, s)

Exit Function
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : ConvertToDate][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Function
End Function

Public Function SortCollection(colInput As Collection) As Collection
    On Error GoTo Exception
    
    Dim iCounter As Integer
    Dim iCounter2 As Integer
    Dim temp As Variant
    
    Set SortCollection = New Collection
    For iCounter = 1 To colInput.Count - 1
        For iCounter2 = iCounter + 1 To colInput.Count
            If colInput(iCounter) > colInput(iCounter2) Then
                temp = colInput(iCounter2)
                colInput.Remove iCounter2
                colInput.Add temp, temp, iCounter
            End If
        Next iCounter2
    Next iCounter
    Set SortCollection = colInput
    
Exit Function
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : SortCollection][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Function
End Function

Public Sub DeleteColumn(ByVal excelWorkBook As Workbook)
    On Error GoTo Exception

    Dim column As Integer

    column = 1

    excelWorkBook.Activate

    Do While ActiveCell.value <> ""

        Cells(1, column).Select

        Select Case ActiveCell.value

        Case "Time", "Time (absolute)", "시간", "시간 (절대)"
            column = column + 1
        Case "Level", "레벨"
            column = column + 1
        Case "Message", "메시지"
            column = column + 1
        Case Else
            Columns(column).Select
            Selection.Delete Shift:=xlLeft
        End Select
    Loop

    Range("A1").Select

Exit Sub
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : DeleteColumn][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Sub
End Sub

Public Sub LogEdit(ByVal excelWorkBook As Workbook)
    On Error GoTo Exception

    Dim num As Long
    Dim cellValue, checkSCPath, scPath, excelFileName, editPass, sSheetName As String

    excelWorkBook.Activate

    For Each SheetName In excelWorkBook.Sheets

        '●해당 시트 선택●
        SheetName.Activate

        '●해당 시트명이 'Sheet1'이면 제거●
        If SheetName.Name = "Sheet1" Then
            Application.DisplayAlerts = False
            SheetName.Delete
            Application.DisplayAlerts = True

        ElseIf InStr(1, SheetName.Name, "AtdBot") <> 0 Then
            Sheets(SheetName.Name).Select
            Cells.Select
            Selection.Columns.AutoFit
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With

            Range("A:F").Select
            Selection.AutoFilter

        '●Job List 시트 작업●
        ElseIf SheetName.Name = "JobList" Then
            num = 2
            Do While SheetName.Range("A" & num) <> ""
                Sheets(SheetName.Name).Select
                cellValue = SheetName.Cells(num, 6).value
                Select Case cellValue

                    Case "Successful"
                    Range("A" & num & ":F" & num).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent5
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    sSheetName = SheetName.Cells(num, 2).value
                    For Each sheetName2 In excelWorkBook.Sheets
                        If InStr(1, sSheetName, sheetName2.Name) <> 0 Then
                            Sheets(sheetName2.Name).Select
                            With ActiveWorkbook.Sheets(sheetName2.Name).Tab
                                .ThemeColor = xlThemeColorAccent5
                                .TintAndShade = 0.399975585192419
                            End With
                            Sheets(SheetName.Name).Select
                            Range("A" & num).Select
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" + sheetName2.Name + "'!A1", TextToDisplay:=SheetName.Cells(num, 1).value
                        End If
                    Next

                    Case "Faulted"
                    Range("A" & num & ":F" & num).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8420607
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    sSheetName = SheetName.Cells(num, 2).value
                    For Each sheetName2 In Application.ActiveWorkbook.Sheets
                        If InStr(1, sSheetName, sheetName2.Name) <> 0 Then
                            Sheets(sheetName2.Name).Select
                            With ActiveWorkbook.Sheets(sheetName2.Name).Tab
                                .Color = 8420607
                                .TintAndShade = 0
                            End With
                            Sheets(SheetName.Name).Select
                            Range("A" & num).Select
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" + sheetName2.Name + "'!A1", TextToDisplay:=SheetName.Cells(num, 1).value
                        End If
                    Next

                    Case "Running"
                    Range("A" & num & ":F" & num).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent4
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    sSheetName = SheetName.Cells(num, 2).value
                    For Each sheetName2 In Application.ActiveWorkbook.Sheets
                        If InStr(1, sSheetName, sheetName2.Name) <> 0 Then
                            Sheets(sheetName2.Name).Select
                            With ActiveWorkbook.Sheets(sheetName2.Name).Tab
                                .ThemeColor = xlThemeColorAccent4
                                .TintAndShade = 0.399975585192419
                            End With
                            Sheets(SheetName.Name).Select
                            Range("A" & num).Select
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" + sheetName2.Name + "'!A1", TextToDisplay:=SheetName.Cells(num, 1).value
                        End If
                    Next

                    Case "Stopped"
                    Range("A" & num & ":F" & num).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent4
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    sSheetName = SheetName.Cells(num, 2).value
                    For Each sheetName2 In Application.ActiveWorkbook.Sheets
                        If InStr(1, sSheetName, sheetName2.Name) <> 0 Then
                            Sheets(sheetName2.Name).Select
                            With ActiveWorkbook.Sheets(sheetName2.Name).Tab
                                .ThemeColor = xlThemeColorAccent4
                                .TintAndShade = 0.399975585192419
                            End With
                            Sheets(SheetName.Name).Select
                            Range("A" & num).Select
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" + sheetName2.Name + "'!A1", TextToDisplay:=SheetName.Cells(num, 1).value
                        End If
                    Next

                    Case Else

                End Select

                num = num + 1

            Loop

            Sheets(SheetName.Name).Select
            Cells.Select
            Selection.Columns.AutoFit

        Else

        '●해당 시트가 작업이 끝난 시트인지 확인●
        num = 2
        editPass = "NoPass"
        Do While SheetName.Range("A" & num) <> ""
            cellValue = SheetName.Cells(num, 2).value
            If cellValue = "Screenshot" Then
                editPass = "Pass"
                Exit Do
            End If
            num = num + 1

        Loop

        '●해당 시트가 작업 대상인지 아닌지 확인●
        If editPass = "NoPass" Then

            '●전체 영역 선택●
            Cells.Select

            '●셀 채우기 없음●
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

            '●셀 가운데 줄 맞춤●
            With Selection
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlCenter
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With

            num = 2

            Do While SheetName.Range("A" & num) <> ""
            
                '● 날짜 포맷 변경●
                
                SheetName.Cells(num, 1) = Format(CDate(SheetName.Cells(num, 1)), "yyyy-mm-dd hh:mm:ss")
                
                '●'Level' Column 확인●
                cellValue = SheetName.Cells(num, 2).value

                Select Case cellValue

                    '●Error●
                    Case "Error"
                    Range("A" & num & ":C" & num).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8420607
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With

                    '●Warning●
                    Case "Warning"
                    Range("A" & num & ":C" & num).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent4
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With

                    Case Else

                End Select

                '●Screenshot Level 추가●
                checkSCPath = SheetName.Cells(num, 3).value
                    If InStr(checkSCPath, "D:\_RPA\") <> 0 And (InStr(checkSCPath, ".jpeg") <> 0 Or InStr(checkSCPath, "[처리로그] - 스크린샷") <> 0) Then
                        scPath = "D:\_수행EXCEPTION\" & Split(Split(checkSCPath, "D:\_RPA\")(1), "\")(UBound(Split(Split(checkSCPath, "D:\_RPA\")(1), "\")))
                        ActiveSheet.Hyperlinks.Add Anchor:=SheetName.Cells(num, 3), Address:=scPath, TextToDisplay:=checkSCPath
                        Range("B" & num).value = "Screenshot"
                        Range("A" & num & ":C" & num).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent5
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                    End If

                num = num + 1

            Loop

            '●필터걸기●
            Range("A:C").Select
            Selection.AutoFilter

            '●열 너비 자동 맞춤●
            Cells.Select
            Selection.Font.Size = 9
            Selection.Columns.AutoFit

            '●첫 행 고정●
            Range("A1").Select
            With ActiveWindow
                .SplitColumn = 0
                .SplitRow = 1
            End With
            ActiveWindow.FreezePanes = True

            Range("A1").Select
            End If

        End If

    Next
    
    '●필터걸기●
    Sheets("JobList").Select
    Range("A:F").Select
    Selection.AutoFilter
    
    '●열 너비 자동 맞춤●
    Cells.Select
    Selection.Font.Size = 9
    Selection.Columns.AutoFit
    
    '●첫 행 고정●
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
    
    '●JobList A1 선택하기●
    For Each SheetName In excelWorkBook.Sheets
        If SheetName.Name <> "JobList" Then SheetName.Visible = xlSheetVeryHidden
    Next
    
    Call CopyCodeModule

    Application.ScreenUpdating = True
    
    excelWorkBook.Save

Exit Sub
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : LogEdit][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Sub
End Sub

Public Sub CopyCodeModule()
    On Error GoTo Exception
    
    Dim destination As Object
    
1   Set destination = Workbooks(StringFormat("{0}.xlsm", Format(Now, "yymmdd"))).VBProject.VBComponents("Sheet1").CodeModule

2   destination.AddFromString _
    "Private Sub Worksheet_Activate()" & vbNewLine & _
    "    For Each ws In ActiveWorkbook.Worksheets" & vbNewLine & _
    "        If ws.Name <> " & Chr(34) & "JobList" & Chr(34) & _
    "Then ws.Visible = xlSheetVeryHidden" & vbNewLine & _
    "    Next" & vbNewLine & "End Sub"
    
3   destination.AddFromString _
    "Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)" & vbNewLine & _
    "    Sheets(CStr(Range(" & Chr(34) & "B" & Chr(34) & "& ActiveCell.Row).value)).Visible = xlSheetVisible" & vbNewLine & _
    "    Sheets(CStr(Range(" & Chr(34) & "B" & Chr(34) & "& ActiveCell.Row).value)).Select" & vbNewLine & _
    "End Sub"
    
Exit Sub
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : CopyCodeModule][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Sub
End Sub

Public Sub GetFolderList(ByVal folderPath As String)
    On Error GoTo Exception
    
    Dim index As Integer
    Dim FSO, folderInfo, subFolderInfo As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set folderInfo = FSO.GetFolder(folderPath)
    Set subFolderInfo = folderInfo.SubFolders
    
    index = 1
    
    Sheets("FolderSize").Cells.Clear
    Columns("C:C").Select
    Selection.NumberFormatLocal = "#,##0_ "
    Range("A1").Select
    Sheets("FolderSize").Cells(1, 1) = "FolderName"
    Sheets("FolderSize").Cells(1, 2) = "Size"
    Sheets("FolderSize").Cells(1, 3) = "Byte"
    Sheets("FolderSize").Cells(2, 4) = "시작시간"
    Sheets("FolderSize").Cells(3, 4) = "종료시간"
    Sheets("FolderSize").Cells(4, 4) = "소요시간"
    Sheets("FolderSize").Cells(2, 5) = Now
    Sheets("FolderSize").Range("A1:B1").Select
    Selection.AutoFilter
    ThisWorkbook.Save
    
    For Each subFolder In subFolderInfo
        index = index + 1
        Call GetFolderSize(subFolder.Path, index)
        Sheets("FolderSize").Cells(3, 5) = Now
    Next
    
    ThisWorkbook.Save
    
Exit Sub
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : GetFolderList][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Sub
End Sub

Public Sub GetFolderSize(ByVal folderPath As String, ByVal rowNum As Integer)
    On Error GoTo Exception
    
    Dim FSO, folderInfo As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set folderInfo = FSO.GetFolder(folderPath)
    
    Sheets("FolderSize").Cells(rowNum, 1) = folderInfo.Name
    Sheets("FolderSize").Cells(rowNum, 2) = SetBytes(folderInfo.Size)
    Sheets("FolderSize").Cells(rowNum, 3) = folderInfo.Size
    
    Set FSO = Nothing
    Set folderInfo = Nothing
        
Exit Sub
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : GetFolderSize][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Sub
End Sub

Public Function SetBytes(bytes) As String
    On Error GoTo Exception

    If bytes >= 1073741824 Then
        SetBytes = Format(bytes / 1024 / 1024 / 1024, "#0.00") & " GB"
    ElseIf bytes >= 1048576 Then
        SetBytes = Format(bytes / 1024 / 1024, "#0.00") & " MB"
    ElseIf bytes >= 1024 Then
        SetBytes = Format(bytes / 1024, "#0.00") & " KB"
    ElseIf bytes < 1024 Then
        SetBytes = Fix(bytes) & " Bytes"
    End If

Exit Function
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : SetBytes][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Function
End Function

Function URLEncode(varText As Variant, Optional blnEncode = True)
    Static objHtmlfile As Object
    
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        With objHtmlfile.parentWindow
            .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        End With
    End If
    
    If blnEncode Then
        URLEncode = objHtmlfile.parentWindow.encode(varText)
    End If
End Function

Function URLDecode(varText As Variant, Optional blnEncode = True)
    Static objHtmlfile As Object

    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        With objHtmlfile.parentWindow
            .execScript "function decode(s) {return decodeURIComponent(s)}", "jscript"
        End With
    End If
    
    If blnEncode Then
        URLDecode = objHtmlfile.parentWindow.decode(varText)
    End If
End Function

Public Function DownloadFile(ByVal downloadURL As String, ByVal savePath As String)

    Dim objHTTP, objStream As Object
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    objHTTP.SetOption 2, objHTTP.GetOption(2)
    objHTTP.Open "GET", downloadURL, False
    objHTTP.send
        
    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objHTTP.responseBody
        objStream.SaveToFile savePath, 1 ' 1 = no overwrite, 2 = overwrite
        objStream.Close
    End If

Exit Function
Exception:
    MsgBox "[Error Num : " & Err.Number & "][Error Source : DownloadFile][Error Msg : " & Err.Description & "]", vbExclamation, "Error"
Exit Function
End Function

