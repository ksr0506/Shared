/* 
※ exception 내에서 선언 할 것 ※
# 해설
1.processName         : Project 폴더 내부의 'project.json'을 읽어, json내의 Project 이름을 읽어옴
> Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText("project.json"))("name")

2.workflowFileName    : 에러가 발생한 Workflow파일의 이름
> objFaultedDetails.GetType().GetProperty("WorkflowFile").GetValue(objFaultedDetails,Nothing).ToString

3.activityId          : 해당 Activity Id로 검색하게 되면 오류가 발생한 Activity가 검색 가능(Workflow 별로 번호가 각각 생성되어, 오류가 발생한 Workflow 기준에서 검색해야함.)
> String.Join("0", objFaultedDetails.GetType().GetProperty("ActivityId").GetValue(oFaultedDetails,Nothing).ToString.Split("."c).[Select](Function(s) Integer.Parse(s).ToString("X2")))

4.activityDisplayName : 에러가 발생한 Activity의 DisplayName
> objFaultedDetails.GetType().GetProperty("DisplayName").GetValue(objFaultedDetails,Nothing).ToString

5.errorMsg            : 에러 메시지
> exception.Message

# Log Message 샘플
[1.processName : Test]
[2.workflowFileName : C:\UiPath\Test\Process.xaml]
[3.activityId : 01002]
[4.activityDisplayName : Click 'uipath.studio.project.e...']
[5.errorMsg : 이 Selector에 해당하는 UI 엘리먼트를 찾을 수 없습니다. <wnd app='explo1rer.exe' cls='Shell_TrayWnd' idx='*'/>
<uia cls='Start2' name='시작'/>
]
*/


//Assign
objFaultedDetails(of Object) = exception.Data.item("FaultedDetails")

//Log Message
String.Format("[1.processName : {0}]{5}[2.workflowFileName : {1}]{5}[3.activityId : {2}]{5}[4.activityDisplayName : {3}]{5}[5.errorMsg : {4}]",
Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText("project.json"))("name"),
objFaultedDetails.GetType().GetProperty("WorkflowFile").GetValue(objFaultedDetails,Nothing).ToString,
String.Join("0", objFaultedDetails.GetType().GetProperty("ActivityId").GetValue(oFaultedDetails,Nothing).ToString.Split("."c).[Select](Function(s) Integer.Parse(s).ToString("X2"))),
objFaultedDetails.GetType().GetProperty("DisplayName").GetValue(objFaultedDetails,Nothing).ToString,
exception.Message,
Environment.NewLine)
