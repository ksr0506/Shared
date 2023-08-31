Public Sub GetjobLog(ByVal JobResult As Boolean)
    On Error GoTo Exception

1   Dim rowNum, retryNum, cnt As Integer
2   Dim allWorkbook As Workbook
3   Dim fileSystemObject, csvFileObject As Object
4   Dim jobID As Range
5   Dim csv, j, jsonData As Variant
6   Dim i As Long
7   Dim saveFolderPath, allWorkbookPath, csvWorkbookPath, tempSheetName As String
    Dim headerNameArray
        
    headerNameArray = Array("ProcessName", "Id", "MachineName", "StartTime", "EndTime", "State")

8   Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
        
9   saveFolderPath = fileSystemObject.BuildPath(ThisWorkbook.Sheets("UserInfo").Range("F3"), Format(Now, "yymmdd"))
10  allWorkbookPath = fileSystemObject.BuildPath(saveFolderPath, Format(Now, "yymmdd") + ".xlsm")
    
    If fileSystemObject.FolderExists(saveFolderPath) = False Then fileSystemObject.CreateFolder saveFolderPath
    
    If fileSystemObject.FolderExists("D:\_수행EXCEPTION") = False Then fileSystemObject.CreateFolder "D:\_수행EXCEPTION"
    
    If fileSystemObject.FileExists(allWorkbookPath) Then fileSystemObject.DeleteFile allWorkbookPath
        
12  Set allWorkbook = Workbooks.Add

13  allWorkbook.SaveAs Filename:=allWorkbookPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
14  allWorkbook.Activate
15  ActiveSheet.Name = "JobList"

    For cnt = 0 To UBound(headerNameArray)
        Sheets("JobList").Cells(1, cnt + 1) = headerNameArray(cnt)
    Next cnt

22  allWorkbook.Save

    For rowNum = 2 To ThisWorkbook.Sheets("JobInfo").Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
        If ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 1) <> "" Then
            
            csvWorkbookPath = fileSystemObject.BuildPath(saveFolderPath, StringFormat("{0}_{1}.csv", ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 1), ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 2)))
            
            If fileSystemObject.FileExists(csvWorkbookPath) = False Then
                            
                For retryNum = 1 To 30
                
                    jsonData = LogDto(ThisWorkbook.Sheets("UserInfo").Range("B3"), organizationUnitId, ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 7))
                    csv = ParseCSVToArray(CStr(jsonData))
                    
                    If UBound(csv, 1) > 1 Then retryNum = 30
                
                Next retryNum
                
            If UBound(csv, 1) > 1 Then
                
                LogResult = True
                
                Set csvFileObject = fileSystemObject.CreateTextFile(csvWorkbookPath, True)
                
                csvFileObject.Write jsonData
                csvFileObject.Close
                
            End If
            
            Else
            
                LogResult = True
            
                csv = ParseCSVToArray(fileSystemObject.OpenTextFile(csvWorkbookPath).ReadAll)
            
            End If
            
30          allWorkbook.Sheets("JobList").Activate

31          Sheets("JobList").Cells(rowNum, 1) = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 1)
32          Sheets("JobList").Cells(rowNum, 2) = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 2)
33          Sheets("JobList").Cells(rowNum, 3) = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 8)
34          Sheets("JobList").Cells(rowNum, 4) = Format(CDate(ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 4)), "yyyy-mm-dd HH:mm:ss")
35          Sheets("JobList").Cells(rowNum, 5) = Format(CDate(ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 5)), "yyyy-mm-dd HH:mm:ss")

            If UBound(csv, 1) > 1 Then

37              Sheets("JobList").Cells(rowNum, 6) = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 3)

38              tempSheetName = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 2)

39              Sheets.Add(After:=Sheets(Sheets.Count)).Name = tempSheetName

                For i = LBound(csv, 1) To UBound(csv, 1)
                    For j = LBound(csv, 2) To UBound(csv, 2)
40                      Sheets(tempSheetName).Cells(i, j) = csv(i, j)
                    Next
                Next
                
41              Call DeleteColumn(allWorkbook)
                
            Else
            
                If ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 6) = "" Then
                
                    Sheets("JobList").Cells(rowNum, 6) = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 3)
                
                Else
                
42                  Sheets("JobList").Cells(rowNum, 6) = ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 6)

                End If

            End If
            
43          allWorkbook.Save
            
        End If

    Next rowNum

44  If LogResult Then Call LogEdit(allWorkbook)

Exit Sub
Exception:
    If Err.Number = 70 Then
45      MsgBox StringFormat("[Error Num : {0}][Error Source : GetjobLog][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, "?씠?쟾?뿉 ?닔?뻾?맂 ?뤃?뜑 ?샊??? ?뙆?씪 ?궘?젣?뿉 ?떎?뙣?뻽?뒿?땲?떎. ????옣 寃쎈줈??? 愿??젴?맂 ?뙆?씪?씠?굹 ?뤃?뜑瑜? ?떕??? ?썑 ?옱?떆?룄?빐二쇱꽭?슂."), vbExclamation, "Error"
    Else
46      MsgBox StringFormat("[Error Num : {0}][Error Source : GetjobLog][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
    End If
Exit Sub
End Sub
