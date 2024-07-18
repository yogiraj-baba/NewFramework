testPath = "C:\NewFramework\YOGI Framework\Framework Driver Script\Framework Driver Script\Test.tsp"
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
DoesFileExist = objFSO.FileExists(testPath)
Set objFSO = Nothing
If DoesFileExist Then
Dim qtApp
Dim qtTest
Dim qtResultsOpt
Set qtApp = CreateObject("QuickTest.Application")
qtApp.Launch
qtApp.Visible = True
qtApp.Open testPath, True
Set qtTest = qtApp.Test
Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtResultsOpt.ResultsLocation = "C:\YOGI Framework\YOGI Framework\Framework Driver Script\Framework Driver Script\Result"
qtTest.Run qtResultsOpt,True
qtTest.Run
qtTest.Close
qtApp.Quit
Else
Wscript.Echo "Test is not found at"&testPath 
End If
