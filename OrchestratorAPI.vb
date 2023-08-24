'
'   UiPath Orchestrator API관련 모듈
'

Option Explicit

Public loginFinish, JobResult, LogResult As Boolean
Public Token, orchestratorURL As String
Public organizationUnitId As Long
Public userCollection As Collection
Public idDictionary As New Dictionary
Public key As Variant

'
'   'Processes' worksheet에서 Job과 Log를 조회하는 작업을 시작하기위한 Start Btn procedure
'

Sub StartButton_Click()
    On Error GoTo Exception

1   ThisWorkbook.Save

    If loginFinish = False Then
2       frmLogin.Show
    Else
3       jobSearchForm.Show
    End If

Exit Sub
Exception:
4   MsgBox StringFormat("[Error Num : {0}][Error Source : StartButton_Click][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'
'   'AuditLogs' worksheet에서 Audit Log를 조회하는 작업을 시작하기위한 Start Btn procedure
'

Sub AuditStartButton_Click()
    On Error GoTo Exception

    If loginFinish = False Then
1       frmLogin.Show
    Else
        Call Authenticate_Cloud("lx_pantos_rpa", "8DEv1AMNXczW3y4U15LL3jYf62jK93n5", "6IWxWuJFHzZBG0jS5u-v9B8cjKm30LhvtzdaUqqru9mNx")
        Call GetUsers(Sheets("UserInfo").Range("B3"))
2       AuditSearchForm.Show
    End If

Exit Sub
Exception:
3   MsgBox StringFormat("[Error Num : {0}][Error Source : AuditStartButton_Click][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'
'   'FolderSize' worksheet에서 폴더별 용량을 확인하는 작업을 시작하기위한 Start Btn procedure
'

Sub folderSizeStartButton_Click()
    On Error GoTo Exception

1       folderPathForm.Show

Exit Sub
Exception:
2   MsgBox StringFormat("[Error Num : {0}][Error Source : folderSizeStartButton_Click][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'
'   Onpremise타입의 Orchestrator에서 API를 호출하기 위한 Authenticate
'
'   #Account#
'
'   [Post]/api/Account/Authenticate
'
'   [Parameters]
'   orchUrl (String)                : 접속할 오케스트레이터 URL
'   tenancyName (String)            : 테넌트 이름,테넌트 명이 없으면 빈값으로 입력
'   usernameOrEmailAddress (String) : 로그인에 사용할 계정 ID 혹은 Email
'   Password (String)               : 계정 PW

Public Sub Authenticate_Onpremise(ByVal orchUrl As String, ByVal tenancyName As String, ByVal usernameOrEmailAddress As String, ByVal Password As String)
    On Error GoTo Exception

1   Dim strJson, requestUrl As String
2   Dim objHTTP, objJson As Object

3   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
4   requestUrl = orchUrl & "/api/account/authenticate"
5   strJson = StringFormat("{""tenancyName"":""{0}"",""usernameOrEmailAddress"":""{1}"",""password"":""{2}""}", tenancyName, usernameOrEmailAddress, Password)

6   objHTTP.SetOption 2, objHTTP.GetOption(2)
7   objHTTP.Open "POST", requestUrl, False
8   objHTTP.setRequestHeader "Content-type", "application/json"
9   objHTTP.send strJson

    If objHTTP.Status = 200 Then
10      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)
11      Token = objJson("result")
12      loginFinish = True
    Else
13      loginFinish = False
    End If

Exit Sub
Exception:
14  MsgBox StringFormat("[Error Num : {0}][Error Source : Authenticate][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'
'   Cloud타입의 Orchestrator에서 API를 호출하기 위한 Authenticate
'

Public Sub Authenticate_Cloud(ByVal tenancyName As String, ByVal clientId As String, ByVal userKey As String)
    On Error GoTo Exception

1   Dim strJson, requestUrl As String
2   Dim objHTTP, objJson As Object

3   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
4   requestUrl = "https://account.uipath.com/oauth/token"
5   strJson = StringFormat("{""grant_type"":""refresh_token"",""client_id"":""{0}"",""refresh_token"":""{1}""}", clientId, userKey)

6   objHTTP.SetOption 2, objHTTP.GetOption(2)
7   objHTTP.Open "POST", requestUrl, False
8   objHTTP.setRequestHeader "Content-type", "application/json"
9   objHTTP.setRequestHeader "X-UIPATH-TenantName", tenancyName
10  objHTTP.send strJson

    If objHTTP.Status = 200 Then
11      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)
12      Token = objJson("access_token")
13      loginFinish = True
    Else
14      loginFinish = False
    End If

Exit Sub
Exception:
15  MsgBox StringFormat("[Error Num : {0}][Error Source : Authenticate][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

Public Sub JobDto(ByVal orchUrl As String, ByVal uipathOrganizationUnitId As Long, ByVal jobState As String, ByVal fromDate As String, ByVal toDate As String)
    On Error GoTo Exception

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary
4   Dim index, searchCount, i As Integer
5   Dim processName As Range
    Dim headerNameArray
    Dim jsonKeyArray
    
    headerNameArray = Array("ProcessName", "Id", "State", "StartTime", "EndTime", "Info", "Key", "MachineName")
    jsonKeyArray = Array("ReleaseName", "Id", "State", "StartTime", "EndTime", "Info", "Key", "HostMachineName")
    
6   index = 1

    searchCount = 0

7   For Each processName In Range("A2:A" & Sheets("Processes").Range("A1").SpecialCells(xlCellTypeLastCell).Row)
        If processName.value <> "" And processName.value <> "Process Name" And processName.EntireRow.Hidden = False Then searchCount = searchCount + 1
    Next
    
8   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    If jobState = "All" Then
9       requestUrl = orchUrl & StringFormat("/odata/Jobs?$filter=EndTime ge {0} and EndTime le {1} and (State eq 'Successful' or State eq 'Faulted' or State eq 'Stopped')", fromDate, toDate)
    ElseIf jobState = "Faulted, Stopped" Then
10      requestUrl = orchUrl & StringFormat("/odata/Jobs?$filter=EndTime ge {0} and EndTime le {1} and (State eq 'Faulted' or State eq 'Stopped')", fromDate, toDate)
    Else
11      requestUrl = orchUrl & StringFormat("/odata/Jobs?$filter=EndTime ge {0} and EndTime le {1} and State eq '{2}'", fromDate, toDate, jobState)
    End If
    
12  objHTTP.SetOption 2, objHTTP.GetOption(2)
13  objHTTP.Open "GET", requestUrl, False
14  objHTTP.setRequestHeader "Content-type", "application/json"
15  objHTTP.setRequestHeader "X-UIPATH-OrganizationUnitId", uipathOrganizationUnitId
16  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
17  objHTTP.send

    If objHTTP.Status = 200 Then
        
18      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)
        If objJson("value").Count > 0 Then

19          Sheets("JobInfo").Cells.Clear
            For i = 0 To UBound(headerNameArray)
                Sheets("JobInfo").Cells(1, i + 1) = headerNameArray(i)
            Next i

28          ThisWorkbook.Save

            For Each value In objJson("value")
            
                If searchCount > 0 Then '특정 Process로 Job 조회할 경우
                
                    For Each processName In Range("A2:A" & Sheets("Processes").Range("A1").SpecialCells(xlCellTypeLastCell).Row)
                    
                        If processName.value <> "" And processName.EntireRow.Hidden = False And InStr(UCase(value("ReleaseName")), UCase(processName.value)) > 0 Then

29                          index = index + 1

                            For i = 0 To UBound(jsonKeyArray)
                            
                                If jsonKeyArray(i) <> "StartTime" And jsonKeyArray(i) <> "EndTime" Then
                                
                                    Sheets("JobInfo").Cells(index, i + 1) = value(jsonKeyArray(i))
                                
                                Else
                                
                                    If jsonKeyArray(i) = "StartTime" Then
                                        If InStr(value("StartTime"), ".") > 0 Then
33                                          Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                        Else
34                                          Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                        End If
                                    End If
                                    
                                    If jsonKeyArray(i) = "EndTime" Then
                                        If InStr(value("EndTime"), ".") > 0 Then
35                                          Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                        Else
36                                          Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                        End If
                                    End If
                                                              
                                End If
                                
                            Next i
    
                        End If
    
                    Next processName
                    
                Else '특정 Process로 검색하지 않고 기간으로 모든 Job 조회할 경우
                
40                  index = index + 1
                
                    For i = 0 To UBound(jsonKeyArray)
                    
                        If jsonKeyArray(i) <> "StartTime" And jsonKeyArray(i) <> "EndTime" Then
                        
                            Sheets("JobInfo").Cells(index, i + 1) = value(jsonKeyArray(i))
                        
                        Else
                        
                            If jsonKeyArray(i) = "StartTime" Then
                                If InStr(value("StartTime"), ".") > 0 Then
                                    Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                Else
                                    Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                End If
                            End If
                            
                            If jsonKeyArray(i) = "EndTime" Then
                                If InStr(value("EndTime"), ".") > 0 Then
                                    Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                Else
                                    Sheets("JobInfo").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(jsonKeyArray(i)), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                                End If
                            End If
                                                      
                        End If
                        
                    Next i
                
                End If
                
            Next value

            If index = 1 Then
51              MsgBox "조회 결과가 없습니다.", vbInformation, "Result"
52              JobResult = False
            Else
53              Sheets("JobInfo").Activate
54              Sheets("JobInfo").Range("A2:H" & Sheets("JobInfo").Range("A1").SpecialCells(xlCellTypeLastCell).Row).sort Key1:=Range("B2"), Order1:=xlAscending
55              Sheets("JobInfo").Range("A2:H" & Sheets("JobInfo").Range("A1").SpecialCells(xlCellTypeLastCell).Row).RemoveDuplicates Columns:=2, Header:=xlYes
56              Sheets("Processes").Activate
57              JobResult = True
            End If

        Else
58          MsgBox "조회 결과가 없습니다.", vbInformation, "Result"
59          JobResult = False
        End If

    End If

Exit Sub
Exception:
60  MsgBox StringFormat("[Error Num : {0}][Error Source : JobDto][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'   로그데이터를 다운로드 받는 module
'
'   RobotLogs
'   [GET] /odata/RobotLogs/UiPath.Server.Configuration.OData.Reports | Bucket Id(Key) 얻기
'   [Parameters]
'   X-UIPATH-OrganizationUnitId : 459468   (LXPantos)

Public Function LogDto(ByVal orchUrl As String, ByVal uipathOrganizationUnitId As Long, ByVal jobKey As String) As Variant
    On Error GoTo Exception

1   Dim requestUrl As String
2   Dim objHTTP As Object
    
3   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
4   requestUrl = orchUrl & StringFormat("/odata/RobotLogs/UiPath.Server.Configuration.OData.Reports?$filter=(Jobkey eq {0} )&amp;$orderby=TimeStamp desc&amp;$top=10&amp;", jobKey)
    
5   objHTTP.SetOption 2, objHTTP.GetOption(2)
6   objHTTP.Open "GET", requestUrl, False
7   objHTTP.setRequestHeader "Content-type", "application/json"
8   objHTTP.setRequestHeader "X-UIPATH-OrganizationUnitId", uipathOrganizationUnitId
9   objHTTP.setRequestHeader "Authorization", "Bearer " & Token
10  objHTTP.send

11  LogDto = objHTTP.responseText

Exit Function
Exception:
12  MsgBox StringFormat("[Error Num : {0}][Error Source : LogDto][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Function
End Function

Public Sub GetjobLog(ByVal JobResult As Boolean)
    On Error GoTo Exception

1   Dim rowNum, retryNum, cnt As Integer
2   Dim allWorkbook As Workbook
3   Dim fileSystemObject, csvFileObject As Object
4   Dim jobID As Range
5   Dim csv, j As Variant
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
                
                    csv = ParseCSVToArray(LogDto(ThisWorkbook.Sheets("UserInfo").Range("B3"), organizationUnitId, ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 7)))
                    
                    If UBound(csv, 1) > 1 Then retryNum = 30
                
                Next retryNum
                
            If UBound(csv, 1) > 1 Then
                
                LogResult = True
                
                Set csvFileObject = fileSystemObject.CreateTextFile(csvWorkbookPath, True)
                csvFileObject.Write LogDto(ThisWorkbook.Sheets("UserInfo").Range("B3"), organizationUnitId, ThisWorkbook.Sheets("JobInfo").Cells(rowNum, 7))
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

Public Sub GetProcesses(ByVal orchUrl As String)
    On Error GoTo Exception

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary
4   Dim index, i As Integer
    Dim headerNameArray
    
    headerNameArray = Array("IsActive", "SupportsMultipleEntryPoints", "RequiresUserInteraction", "Title", "Version", "Key", _
                            "Description", "Published", "IsLatestVersion", "OldVersion", "ReleaseNotes", "Authors", "ProjectType", "Id")
    
7   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
8   requestUrl = orchUrl & "/odata/Processes"
    
10  objHTTP.SetOption 2, objHTTP.GetOption(2)
11  objHTTP.Open "GET", requestUrl, False
12  objHTTP.setRequestHeader "Content-type", "application/json"
13  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
14  objHTTP.send
    
    If objHTTP.Status = 200 Then
    
15      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

        If objJson("value").Count > 0 Then


16          Sheets("AllProcesses").Cells.Clear
            For i = 0 To UBound(headerNameArray)
                Sheets("AllProcesses").Cells(1, i + 1) = headerNameArray(i)
            Next i
            
31          ThisWorkbook.Save

32          For Each value In objJson("value")

33              index = index + 1

                For i = 0 To UBound(headerNameArray)
                
                    If headerNameArray(i) <> "Published" Then
                    
                        Sheets("AllProcesses").Cells(index, i + 1) = value(headerNameArray(i))
                    
                    Else
                    
                            If InStr(value("Published"), ".") > 0 Then
41                              Sheets("AllProcesses").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(headerNameArray(i)), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                            Else
42                              Sheets("AllProcesses").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(headerNameArray(i)), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                            End If
                        
                    End If
                    
                Next i

            Next value

49          ThisWorkbook.Save
        End If
    End If
Exit Sub
Exception:
50  MsgBox StringFormat("[Error Num : {0}][Error Source : GetProcesses][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

Public Sub GetProcessSchedules(ByVal orchUrl As String, ByVal uipathOrganizationUnitId As Long)
    On Error GoTo Exception

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary
4   Dim index, i As Integer
    Dim headerNameArray
    
    headerNameArray = Array("Enabled", "Name", "ReleaseId", "ReleaseKey", "ReleaseName", "PackageName", "EnvironmentName", "EnvironmentId", "JobPriority", "RuntimeType", _
                            "StartProcessCron", "StartProcessCronDetails", "StartProcessCronSummary", "StartProcessNextOccurrence", "StartStrategy", "StopProcessExpression", _
                            "StopStrategy", "ExternalJobKey", "TimeZoneId", "TimeZoneIana", "UseCalendar", "CalendarId", "CalendarName", "StopProcessDate", "InputArguments", _
                            "QueueDefinitionId", "QueueDefinitionName", "ItemsActivationThreshold", "ItemsPerJobActivationTarget", "MaxJobsForActivation", "Id")
                            
6   index = 1

7   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
8   requestUrl = orchUrl & "/odata/ProcessSchedules"
    
10  objHTTP.SetOption 2, objHTTP.GetOption(2)
11  objHTTP.Open "GET", requestUrl, False
12  objHTTP.setRequestHeader "Content-type", "application/json"
13  objHTTP.setRequestHeader "X-UIPATH-OrganizationUnitId", uipathOrganizationUnitId
14  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
15  objHTTP.send
    
    If objHTTP.Status = 200 Then
    
16      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

        If objJson("value").Count > 0 Then


17          Sheets("AllTriger").Cells.Clear
            For i = 0 To UBound(headerNameArray)
                Sheets("AllTriger").Cells(1, i + 1) = headerNameArray(i)
            Next i

48          ThisWorkbook.Save

49          For Each value In objJson("value")

50              index = index + 1

                For i = 0 To UBound(headerNameArray)
                
                    If headerNameArray(i) <> "StartProcessNextOccurrence" Then
                    
                        Sheets("AllTriger").Cells(index, i + 1) = value(headerNameArray(i))
                    
                    Else
                    
                            If InStr(value("StartProcessNextOccurrence"), ".") > 0 Then
73                              Sheets("AllTriger").Cells(index, 14) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value("StartProcessNextOccurrence"), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                            ElseIf IsNull(value("StartProcessNextOccurrence")) Then
                        
                            Else
74                              Sheets("AllTriger").Cells(index, 14) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value("StartProcessNextOccurrence"), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                            End If
                        
                    End If
                    
                Next i
                
            Next value

92          ThisWorkbook.Save
        End If
    End If
Exit Sub
Exception:
93  MsgBox StringFormat("[Error Num : {0}][Error Source : GetProcessSchedules][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

Public Sub GetAuditLogs(ByVal orchUrl As String, ByVal serchKeyword As String, ByVal fromDate As String, ByVal toDate As String)
    On Error GoTo Exception

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary
4   Dim index, i As Integer
5   Dim headerNameArray
    
    headerNameArray = Array("Component", "UserName", "Action", "MethodName", "DisplayName", "OperationText", "ExecutionTime")
    
    index = 1

6   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
7   requestUrl = orchUrl & StringFormat("/odata/AuditLogs?$filter=ExecutionTime ge {0} and ExecutionTime le {1}{2}", fromDate, toDate, serchKeyword)
    
8   objHTTP.SetOption 2, objHTTP.GetOption(2)
9   objHTTP.Open "GET", requestUrl, False
10  objHTTP.setRequestHeader "Content-type", "application/json"
11  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
12  objHTTP.send
    
    If objHTTP.Status = 200 Then
    
13      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

        If objJson("value").Count > 0 Then
        
16          Sheets("AuditLogs").Cells.Clear
            For i = 0 To UBound(headerNameArray)
                Sheets("AuditLogs").Cells(1, i + 1) = headerNameArray(i)
            Next i
22          Sheets("AuditLogs").Range("A1:G1").Select
23          Selection.AutoFilter

24          ThisWorkbook.Save

25          For Each value In objJson("value")

26              index = index + 1
                
                For i = 0 To UBound(headerNameArray)
                
                    If headerNameArray(i) <> "ExecutionTime" Then
                    
                        Sheets("AuditLogs").Cells(index, i + 1) = value(headerNameArray(i))
                    
                    Else
                    
                        If InStr(value("ExecutionTime"), ".") > 0 Then
33                          Sheets("AuditLogs").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(headerNameArray(i)), ".")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                        Else
34                          Sheets("AuditLogs").Cells(index, i + 1) = Format(DateAdd("h", 9, ConvertToDate(Replace(Split(value(headerNameArray(i)), "Z")(0), "T", ""), "YYYY-MM-DDhh:mm:ss")), "yyyy-mm-dd HH:mm:ss")
                        End If
                        
                    End If
                    
                Next i
                
            Next value
            
            Cells.Select
            Selection.Font.Size = 9
35          ThisWorkbook.Save

        End If
    End If
Exit Sub
Exception:
36  MsgBox StringFormat("[Error Num : {0}][Error Source : GetAuditLogs][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

Public Sub GetUsers(ByVal orchUrl As String)
    On Error GoTo Exception

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary
    
4   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
5   Set userCollection = New Collection
    
6   requestUrl = orchUrl & "/odata/Users"
    
7   objHTTP.SetOption 2, objHTTP.GetOption(2)
8   objHTTP.Open "GET", requestUrl, False
9   objHTTP.setRequestHeader "Content-type", "application/json"
10  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
11  objHTTP.send
    
    If objHTTP.Status = 200 Then
    
12      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

        If objJson("value").Count > 0 Then

            For Each value In objJson("value")
            
13              userCollection.Add (value("UserName"))

            Next value
            
14          Call SortCollection(userCollection)
            
        End If
        
    End If
    
Exit Sub
Exception:
15  MsgBox StringFormat("[Error Num : {0}][Error Source : GetAuditLogs][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

Public Sub GetFolders(ByVal orchUrl As String)

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary
    
4   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
5   Set idDictionary = New Dictionary
    
6   requestUrl = orchUrl & "/odata/Folders"
    
7   objHTTP.SetOption 2, objHTTP.GetOption(2)
8   objHTTP.Open "GET", requestUrl, False
9   objHTTP.setRequestHeader "Content-type", "application/json"
10  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
11  objHTTP.send

    If objHTTP.Status = 200 Then
    
12      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

        If objJson("value").Count > 0 Then

            For Each value In objJson("value")
            
13              idDictionary.Add value("DisplayName"), value("Id")

            Next value
            
        End If
        
    End If
    
Exit Sub
Exception:
    MsgBox StringFormat("[Error Num : {0}][Error Source : GetAuditLogs][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'
'   'Processes' worksheet에서 Job과 Log를 조회하는 작업을 시작하기위한 Start Btn procedure
'

Public Sub GetBucketFiles(ByVal orchUrl As String, ByVal bucketId As Long, ByVal directoryPath As String, ByVal fileNameGlob As String, ByVal uipathOrganizationUnitId As Long)

1   Dim requestUrl As String
2   Dim objHTTP, objJson As Object
3   Dim value As Dictionary

4   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
6   requestUrl = orchUrl & StringFormat("/odata/Buckets({0})/UiPath.Server.Configuration.OData.GetFiles?directory={1}&recursive=false&fileNameGlob={2}", bucketId, URLEncode(directoryPath), fileNameGlob)
        
7   objHTTP.SetOption 2, objHTTP.GetOption(2)
8   objHTTP.Open "GET", requestUrl, False
9   objHTTP.setRequestHeader "Content-type", "application/json"
    objHTTP.setRequestHeader "X-UIPATH-OrganizationUnitId", uipathOrganizationUnitId
10  objHTTP.setRequestHeader "Authorization", "Bearer " & Token
11  objHTTP.send

    If objHTTP.Status = 200 Then
    
12      Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

        If objJson("value").Count > 0 Then

            For Each value In objJson("value")
            
13              Call GetReadUri(orchestratorURL, 2829, value("FullPath"), 5, organizationUnitId)

            Next value
            
        End If
        
    End If
    
Exit Sub
Exception:
    MsgBox StringFormat("[Error Num : {0}][Error Source : GetAuditLogs][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'
'   'Processes' worksheet에서 Job과 Log를 조회하는 작업을 시작하기위한 Start Btn procedure
'

Public Sub GetReadUri(ByVal orchUrl As String, ByVal bucketId As Long, ByVal blobFilePath As String, ByVal expiryInMinutes As Long, ByVal uipathOrganizationUnitId As Long)

1   Dim requestUrl, savePath As String
2   Dim objHTTP, objJson, fileSystemObject As Object

4   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
        
    savePath = fileSystemObject.BuildPath("D:\_수행EXCEPTION", blobFilePath)
        
    If fileSystemObject.FileExists(savePath) = False Then 'savePath가 존재하면 작업 수행 제외

6       requestUrl = orchUrl & StringFormat("/odata/Buckets({0})/UiPath.Server.Configuration.OData.GetReadUri?path={1}&expiryInMinutes={2}", bucketId, URLEncode(blobFilePath), expiryInMinutes)
        
7       objHTTP.SetOption 2, objHTTP.GetOption(2)
8       objHTTP.Open "GET", requestUrl, False
9       objHTTP.setRequestHeader "Content-type", "application/json"
        objHTTP.setRequestHeader "X-UIPATH-OrganizationUnitId", uipathOrganizationUnitId
10      objHTTP.setRequestHeader "Authorization", "Bearer " & Token
11      objHTTP.send

        If objHTTP.Status = 200 Then
        
12          Set objJson = JsonConverter.ParseJson(objHTTP.responseText)

            If InStr(objHTTP.responseText, "Uri") <> 0 Then Call DownloadFile(objJson("Uri"), savePath)
        
        End If
    
    End If
Exit Sub
Exception:
    MsgBox StringFormat("[Error Num : {0}][Error Source : GetAuditLogs][Error Line : {1}][Error Msg : {2}]", Err.Number, Erl, Err.Description), vbExclamation, "Error"
Exit Sub
End Sub

'Test 용
Sub GetProcessSchedulesRun()

Dim orchUrl As String
Dim uipathOrganizationUnitId As Long

orchUrl = Sheets("UserInfo").Range("B3")
uipathOrganizationUnitId = 79

Call GetProcessSchedules(orchUrl, uipathOrganizationUnitId)

End Sub

Sub GetBucketFilesTest()

Call GetBucketFiles(Sheets("UserInfo").Range("B3"), 2829, "\", "*230213*", 459468)

End Sub
