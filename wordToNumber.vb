'################################################################
'[Argument]
'Name : in_sWord / Direction : In / Type : String / 변환할 문자열
'Name : out_lNumber / Direction : In / Type : System.Int64 / 변환된 숫자
'Name : out_bResult / Direction : In / Type : Boolean / 변환 결과 (성공 : True, 실패 : False)
'################################################################

Dim regex As System.Text.RegularExpressions.Regex
Dim lTemp, lTotal As Long
Dim listNumber As List(Of Long)
Dim dicWordTable As Dictionary(Of String, Long)
listNumber = New List(Of Long)
dicWordTable = New Dictionary(Of String, Long) From {
        {"zero",0},{"one",1},{"two",2},{"three",3},{"four",4},{"five",5},{"six",6},  
        {"seven",7},{"eight",8},{"nine",9},{"ten",10},{"eleven",11},{"twelve",12},  
        {"thirteen",13},{"fourteen",14},{"fifteen",15},{"sixteen",16},{"seventeen",17},  
        {"eighteen",18},{"nineteen",19},{"twenty",20},{"thirty",30},{"forty",40},  
        {"fifty",50},{"sixty",60},{"seventy",70},{"eighty",80},{"ninety",90},  
        {"hundred",100},{"thousand",1000},{"hundred thousand",100000},{"million",1000000},  
        {"billion",1000000000},{"trillion",1000000000000},{"quadrillion",1000000000000000},  
        {"quintillion",1000000000000000000}
		}

'괄호 및 괄호안 문자 제거
in_sWord = system.Text.RegularExpressions.Regex.Replace(in_sWord,"\([^)]*\)",String.Empty)

lTemp = 0
lTotal = 0L
		
Try
	For Each  wordItem In regex.Matches(in_sWord,"([A-Z])\w+").Cast(Of Match).Select(Function(x As Match) x.Value.ToLowerInvariant)
		If dicWordTable.ContainsKey(wordItem.ToString) Then
			listNumber.Add(dicWordTable(wordItem.ToString)) 'Dictionary에 Key값이 포함되어 있으면, List에 wordItem 추가
		Else
			Throw New System.Exception("상이한 값 발견") '1개의 wordItem이라도 인식이 불가능하면, 값이 상이할거를 고려하여 예외처리
		End If		
	Next
	
	For Each lNumber In listNumber	
		If lNumber >= 1000 Then 
			 lTotal += lTemp * lNumber
		Else If lNumber >= 100 Then
			  lTemp *= lNumber
		Else 
			  lTemp += lNumber
		End If		
	Next
		
	'console.WriteLine(String.Format("Input Word : {0}",in_sWord))
	'console.WriteLine(String.Format("Total : {0}",lTotal))
	'console.WriteLine(String.Format("Temp : {0}",lTemp))
	out_lNumber = (lTotal + lTemp) * If(in_sWord.StartsWith("minus", StringComparison.InvariantCultureIgnoreCase),-1,1) '음수 계산,  위에 예외처리 때문에 의미없음
	out_bResult = If(out_lNumber <> 0,True,False)
	'console.WriteLine(String.Format("Result Number : {0}",out_lNumber))			
Catch
	out_bResult = False
End Try
