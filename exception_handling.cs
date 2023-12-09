/* 
※ exception 내에서 선언 할 것 ※
# 해설
1.processName         : Project 폴더 내부의 'project.json'을 읽어, json내의 Project 이름을 읽어옴
> Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText("project.json"))("name")

2.workflowFileName    : 에러가 발생한 Workflow파일의 이름
> obj_FaultedDetails.GetType().GetProperty("WorkflowFile").GetValue(obj_FaultedDetails,Nothing).ToString

3.activityId          : 해당 Activity Id로 검색하게 되면 오류가 발생한 Activity가 검색 가능(Workflow 별로 번호가 각각 생성되어, 오류가 발생한 Workflow 기준에서 검색해야함.)
> String.Join("0", obj_FaultedDetails.GetType().GetProperty("ActivityId").GetValue(obj_FaultedDetails,Nothing).ToString.Split("."c).[Select](Function(s) Integer.Parse(s).ToString("X2")))
# 작업 순서
  (1) obj_FaultedDetails.GetType().GetProperty("ActivityId").GetValue(obj_FaultedDetails,Nothing).ToString = "1.15"
  (2) ~).ToString.Split("."c)
  (3) ~ Integer.Parse(s).ToString("X2") (https://learn.microsoft.com/ko-kr/dotnet/standard/base-types/standard-numeric-format-strings#hexadecimal-format-specifier-x)
  (4) String.Join("0", ~
# 관련 풀이 : "1.15" -> ("1")("15") -> ("01")("0F") -> "01"+"0"+"0F" -> "0100F"

4.activityDisplayName : 에러가 발생한 Activity의 DisplayName
> obj_FaultedDetails.GetType().GetProperty("DisplayName").GetValue(obj_FaultedDetails,Nothing).ToString

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

# 2023.06.22 추가내용
Try 부분의 Activity가 Invoke Workflow가 아닌경우, Exception.Source에서 정상적으로 DisplayName을 표기 못하듯이, Object 객체에 올바른 데이터가 반환되어 있지 않을 수 있으므로, 예외처리 추가로 필요함.
예시의 Log Message를 그대로 입력할 경우 Error 발생

# 2023.12.09 추가내용
==> Log Message 예외처리 반영하여 개선

*/

//Assign
obj_FaultedDetails(of Object) = exception.Data.item("FaultedDetails")

//Log Message
If(obj_FaultedDetails Is Nothing,
String.Format("exception Source : {0} / exception Message : {1}", systemException.Source, systemException.Message),
String.Format("[1.processName : {0}]{5}[2.workflowFileName : {1}]{5}[3.activityId : {2}]{5}[4.activityDisplayName : {3}]{5}[5.errorMsg : {4}]",
Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText("project.json"))("name"),
obj_FaultedDetails.GetType().GetProperty("WorkflowFile").GetValue(obj_FaultedDetails,Nothing).ToString,
String.Join("0", obj_FaultedDetails.GetType().GetProperty("ActivityId").GetValue(obj_FaultedDetails,Nothing).ToString.Split("."c).[Select](Function(s) Integer.Parse(s).ToString("X2"))),
obj_FaultedDetails.GetType().GetProperty("DisplayName").GetValue(obj_FaultedDetails,Nothing).ToString,
systemException.Message,
Environment.NewLine))
