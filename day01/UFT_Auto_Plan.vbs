Dim uftApp
Dim uftResultsOpt
Dim objExcel

'预关闭excel
Dim Wsh
Set Wsh = WScript.CreateObject("WScript.Shell")
'下行是设置延时清除时间 5000等于5秒
WScript.Sleep(5000)
'下行清除进程
Wsh.Run "taskkill /f /im excel.exe",0
Set Wsh=NoThing
WScript.quit



'On Error Resume Next
' 创建 Application 对象
Set uftApp = CreateObject("QuickTest.Application")
' 创建 Run Results Options 对象
Set uftResultsOpt = CreateObject("QuickTest.RunResultsOptions") 
' 创建 objExcel 对象
Set objExcel = CreateObject("Excel.Application")


' uft启动
uftApp.Launch 
' 设置应用可见
uftApp.Visible = True
objExcel.Visible = True

' 设置 uftApp 运行选项
uftApp.Options.Run.RunMode = "Fast"
'uftApp.Options.Run.ViewResults = False


' 设置测试用例批量文件完整路径
TestCaseFilePath="C:\Users\A9MPSZZ\Desktop\BatchJob.xls" 
' 以只读模式打开Excel则设为True
objExcel.Workbooks.Open TestCaseFilePath
' 设置当前的工作页
Set oSheet = objExcel.Sheets.Item(1)
' 获取总行数
maxRowsCount = oSheet.UsedRange.Rows.Count



''''''''''''''''''''''''主逻辑执行测试'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'定义一个变量，循环起始点
Dim Count:Count = 2
'定义项目名，循环获取
Dim colunmnProject
'定义XML对象
Dim objXML 
'定义XML节点
Dim objNode 
'定义错误数量
Dim errorcount
'保存路径的日期Format
Dim today:today = Replace(Date,"/","_")
Dim thistime:thistime = Replace(Time,":","_")

'创建Log文件夹
set fso=createobject("scripting.filesystemobject")
if not fso.folderExists("C:\Users\A9MPSZZ\Downloads\PSD\Result\" + today) Then
	fso.createfolder("C:\Users\A9MPSZZ\Downloads\PSD\Result\" + today )	
End if
set folderName=fso.createfolder("C:\Users\A9MPSZZ\Downloads\PSD\Result\" + today + "\" + thistime)


'循环设定表里的每一个项目
Do 	
	'获取excel里的项目名
	colunmnProject	= oSheet.UsedRange.Cells(Count,1).Value
	if (colunmnProject <> "")	Then
	
	''''''''''''''''''''''运行测试'''''''''''''''''''''''''''''''''
		'设置结果路径
		uftResultsOpt.ResultsLocation = folderName +"\"+ colunmnProject
		' 设置路径
		path="C:\Users\A9MPSZZ\Downloads\PSD\"+colunmnProject
		' 以只读模式打开测试
		uftApp.Open path,True 
		Set uftTest = uftApp.Test
		' 运行测试
		uftTest.Run uftResultsOpt 
		' 关闭测试'
	''''''''''''''''''''''运行测试'''''''''''''''''''''''''''''''''
	''''''''''''''''''''''输出测试结果log''''''''''''''''''''''''''	

	' 创建 xml object 
	Set objXML = CreateObject("Msxml2.DOMDocument.6.0") 
	' 加载xml文件
	objXML.load(folderName + "\" + colunmnProject + "\Report\run_results.xml") 
	' 查找Status子节点 
	Set objNode = objXML.selectSingleNode("Results/ReportNode/Data/Result")
	oSheet.UsedRange.Cells(Count,3).Value = objNode.Text
	
	''''''''''''''''''''''输出测试结果log''''''''''''''''''''''''''	
	Else 
		oSheet.UsedRange.Cells(Count,2).Value = "Such program is no exist"
		oSheet.UsedRange.Cells(Count,3).Value = "-"
	End	If
	
	'MsgBox colunmnProject
	Count = Count + 1
Loop Until Count > maxRowsCount'当Excel配置文件里的项目全部执行完，结束循环


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''主逻辑执行测试''''''''''''''''''''''''''''''''''''''''''''''''''

'保存
objExcel.activeWorkBook.saveAs  folderName +"\result.xls", 56
objExcel.ActiveWorkBook.Saved = True
' 关闭工作簿
objExcel.Workbooks.Close
' 关闭Excel
objExcel.Quit
' 释放 Run Results Options 对象
Set uftResultsOpt = Nothing 
' 释放 uftTest 对象
Set uftTest = Nothing
' 延迟1S
wscript.sleep 1000
' 关闭uft程序
uftApp.Quit
' 释放 Application 对象 
Set uftApp = Nothing 
MsgBox "over"