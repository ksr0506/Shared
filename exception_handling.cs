/* 
# Exception Log Message 예시
[1.Process Name : Test]
[2.WorkflowFile Name : C:\UiPath\Test\Process.xaml]
[3.Activity Name : Click]
[4.Activity Id : 01002]
[5.Activity DisplayName : Click 'uipath.studio.project.e...']
[6.Exception Type : UiPath.Core.SelectorNotFoundException]
[7.Exception Message : 이 Selector에 해당하는 UI 엘리먼트를 찾을 수 없습니다. <wnd app='explo1rer.exe' cls='Shell_TrayWnd' idx='*'/>
<uia cls='Start2' name='시작'/>
]

# Exception Log Message 해설
1.Process Name : Project 폴더 내부의 'project.json'을 읽어, json내의 Project 이름을 읽어옴
> Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText("project.json"))("name")

2.WorkflowFile Name : 에러가 발생한 Workflow 파일의 이름
> exception.Data.item("FaultedDetails").GetType().GetProperty("WorkflowFile").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString

3.Activity Name : 에러가 발생한 Acitivity의 전체 이름 중 실제 검색할 때 이용하는 부분만 추출 (UiPath.Core.Activities.Click -> Click)
> exception.Data.item("FaultedDetails").GetType().GetProperty("ActivityFullName").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString.Split("."c).Last()

4.Activity Id : 해당 Activity Id로 검색하게 되면 오류가 발생한 Activity가 검색 가능(Workflow 별로 번호가 각각 생성되어, 오류가 발생한 Workflow 기준에서 검색해야함.)
> String.Join("0", exception.Data.item("FaultedDetails").GetType().GetProperty("ActivityId").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString.Split("."c).[Select](Function(s) Integer.Parse(s).ToString("X2")))
# 작업 순서
  (1) exception.Data.item("FaultedDetails").GetType().GetProperty("ActivityId").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString = "1.15"
  (2) ~).ToString.Split("."c)
  (3) ~ Integer.Parse(s).ToString("X2") (https://learn.microsoft.com/ko-kr/dotnet/standard/base-types/standard-numeric-format-strings#hexadecimal-format-specifier-x)
  (4) String.Join("0", ~
# 관련 풀이 : "1.15" -> ("1")("15") -> ("01")("0F") -> "01"+"0"+"0F" -> "0100F"

5.Activity DisplayName : 에러가 발생한 Activity의 DisplayName
> exception.Data.item("FaultedDetails").GetType().GetProperty("DisplayName").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString

6.Exception Type : 예외 종류
> exception.GetType

7.Exception Message : 예외 메시지
> exception.Message
*/

//Log Message
If(exception.Data.item("FaultedDetails") Is Nothing,
String.Format("[1.Exception Source : {1}]{0}[2.Exception Type : {2}]{0}[3.Exception Message : {3}]",
Environment.NewLine,
exception.Source,
exception.GetType,
exception.Message),
String.Format("[1.Process Name : {1}]{0}[2.WorkflowFile Name : {2}]{0}[3.Activity Name : {3}]{0}[4.Activity Id : {4}]{0}[5.Activity DisplayName : {5}]{0}[6.Exception Type : {6}]{0}[7.Exception Message : {7}]",
Environment.NewLine,
Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText("project.json"))("name"),
exception.Data.item("FaultedDetails").GetType().GetProperty("WorkflowFile").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString,
exception.Data.item("FaultedDetails").GetType().GetProperty("ActivityFullName").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString.Split("."c).Last(),
String.Join("0", exception.Data.item("FaultedDetails").GetType().GetProperty("ActivityId").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString.Split("."c).[Select](Function(s) Integer.Parse(s).ToString("X2"))),
exception.Data.item("FaultedDetails").GetType().GetProperty("DisplayName").GetValue(exception.Data.item("FaultedDetails"),Nothing).ToString,
exception.GetType,
exception.Message))
