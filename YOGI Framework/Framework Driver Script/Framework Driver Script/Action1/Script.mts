'######################################################################### Declaration of Public Variables #########################################################################################
Public ConfigFilePath
Public TestExcelPath 'strMainTestDataSheetPath
Public FunctionName 'strFunction
Public DataTablePath 'strTestDataSheetPath
Public ResultsPath 'strResultsFolderPath
Public DLLPath 'strDLLPath
Public HTMLReportPath 'htmlResultFilePath
Public htmlReporter
Public ScreenshotsFolderPath
Public strBrowserSelection

'#############################################################################################################################################################################################

'Prerequisite: Create a user-defined environment variable as "ConfigFilePath" and store the path for Configuration file

'Getting ConfigFilePath from environment variable
ConfigFilePath = Environment.Value("ConfigFilePath")

'Getting values from Config file
ReportTitle = ReadXMLValue(ConfigFilePath, "ReportTitle")
ReportName = ReadXMLValue(ConfigFilePath, "ReportName")
Resources = ReadXMLValue(ConfigFilePath, "Resources")
TestExcelPath = ReadXMLValue(ConfigFilePath, "TestExcelPath")
TestExcelName = ReadXMLValue(ConfigFilePath, "TestExcelName")
DataTablesPath = ReadXMLValue(ConfigFilePath, "DataTablesPath")
LibrariesPath = ReadXMLValue(ConfigFilePath, "LibrariesPath")
RepositoriesPath = ReadXMLValue(ConfigFilePath, "RepositoriesPath")
DLLFilePath = ReadXMLValue(ConfigFilePath, "DLLFilePath")
ResultsPath = ReadXMLValue(ConfigFilePath, "ResultsPath")

'Create a results folder during test run to store the results  
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDate = fnFormatDate(Date,"YYYY/MM/DD","")
strTime = Replace(Trim(FormatDateTime(Now, vbShortTime)),":","")
LatestResultsFolderName = ResultsPath&"\TestAutomationResults_"&strDate&"_"&strTime
If Not objFSO.FolderExists(LatestResultsFolderName) Then
	objFSO.CreateFolder(LatestResultsFolderName)
	ScreenshotsFolderPath = objFSO.CreateFolder(LatestResultsFolderName&"\Screenshots")
	'Store ScreenshotsFolderPath in the env variable
	HTMLReportPath = LatestResultsFolderName&"\Report.html"
End If

'Store screenshot path in env variable
Environment.Value("ScreenshotFolderPath") = ScreenshotsFolderPath

' Create an instance of the htmlReporter class to use methods inside the DLL
Set htmlReporter = DotNetFactory.CreateInstance("UFT_Extent_Reports.HTMLReporter", DLLFilePath & "\ExtentReports_UFT.dll")

'Initialize the report 
htmlReporter.InitializeReport(HTMLReportPath)

'Add Report Name
htmlReporter.AddReportName(ReportName)

'Add sheet at Run Time
DataTable.AddSheet "Test Cases"

'Import the sheets in the corresponding sheet added above
DataTable.ImportSheet TestExcelPath, "Test Cases", "Test Cases"

'Count the total number of sheets in Test Cases sheet
iTestCases = DataTable.GetSheet("Test Cases").GetRowCount

For i = 1 To iTestCases Step 1
	DataTable.GetSheet("Test Cases").SetCurrentRow(i)
	If UCase(DataTable.Value("RunFlag", "Test Cases")) = "TRUE" Then
		sTestCaseID = DataTable.GetSheet("Test Cases").GetParameter("TestCaseID")
		strFunction = DataTable.GetSheet("Test Cases").GetParameter("Function").Value
		FunctionName = strFunction
		strTCName = DataTable.GetSheet("Test Cases").GetParameter("TestCaseName").Value
		reportTheme = DataTable.GetSheet("Test Cases").GetParameter("ReportTheme").Value
		strTestDataSheetName = DataTable.GetSheet("Test Cases").GetParameter("TestDataSheetName").Value
		strTestCategory = DataTable.GetSheet("Test Cases").GetParameter("TestCategory").Value
		strBrowserSelection= DataTable.GetSheet("Test Cases").GetParameter("Browser").Value
		Environment.Value("Browser") = strBrowserSelection
		'Add sheet at Run Time
		DataTable.AddSheet "Test Data"
		strTestDataSheetPath = DataTablePath &"\"&strTestDataSheetName&".xlsx"
		
		'Import the sheet in the corresponding sheet added above
		DataTable.ImportSheet DataTablesPath&strTestDataSheetPath, 1, "Test Data"
		
		'Count the total number of sheets in Test Cases sheet
		rowCnt_TestDataSheet= DataTable.GetSheet("Test Data").GetRowCount
		
		'Add Document Title
		If i = 1 Then
			htmlReporter.AddDocumentTitle(ReportTitle)
		End  If
		
		'Test Name In Report
		htmlReporter.CreateTest(strTCName)
		
		'Assign Autor 
		call htmlReporter.AssignAuthorToTest("Yogiraj")
		
		'Assign Category
		call htmlReporter.AssignCategoryToTest(strTestCategory)
		
		For j = 1 To rowCnt_TestDataSheet
			DataTable.GetSheet("Test Data").SetCurrentRow(j)
			'Execute the testcases for which RunFlag is TRUE in testdatasheet
			If i = 1 Then
				Print "Getting Test Details From: "&TestExcelPath
			End If
			
			Wait 1
			Print "Executing Test Case: "&sTestCaseID
			Wait 1
			Print "Test Case Description: "&strTCName
			Wait 1
			Execute FunctionName & "(htmlReporter)"
			Wait 1
		Next
		
		If Trim(UCase(reportTheme)) = "DARK" Then
			'Change Theme to dark
			htmlReporter.ChangeToDarkTheme
		ElseIf Trim(UCase(reportTheme)) = "NORMAL" Then
			'Generate normal report automatically
		Else
			'Generate normal report automatically
		End If
		
		''Finally generate the html report
		htmlReporter.GenerateReport()	
	End If
Next
