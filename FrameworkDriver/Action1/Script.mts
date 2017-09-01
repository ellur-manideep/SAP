'#******************************************************************************************************************************************
'''Functionality: '''This is Framework Driver Action file..Please put your scenario into Function call and add in the case statement 
'Name 		  :	  	  FrameWorkDriver()
'Input		 	: 			
'limitation    :	 
'Created By	  : 	  Shyju Kumar
'Created On	  :	      july 17, 2017
'Description	: This Driver automatically add all library files, Testdata, object repository and made available to all functions during runtime
'Revision history	:
'#******************************************************************************************************************************************

Public gDictObj
Public strFilePath
call fnAddallresourses()
fnCloseExcelProcess
strExecSummarySheet=environment("strExecSummarySheet")
strTestModules=Environment("strTestModules")
arrWorksheets= getEnvTestSummaryRowCol(strFilePath,strExecSummarySheet,strTestModules)
arrSheetRows= getTestSheetRowCnt(strFilePath,arrWorksheets)

For x = 0 To ubound(arrWorksheets)-1
exeRows=arrSheetRows(x)
	For y = 1 To exeRows-1
		iteration=y
		fnReadExcel strFilePath,arrWorksheets(x),iteration
		set gDictObj=Environment("gDictObj")
		
		gDictObj("XLsheet")=arrWorksheets(x)
		gDictObj("exerow")=iteration
		Environment.Value("XLsheet")=arrWorksheets(x)
		Environment.value("exerow")=iteration

		If lcase(gDictObj("Execute"))="yes"  Then
			
		''''Enter Switch case based on Scenario
		   Select Case gDictObj("Scenario")
			
			    Case "E2E_24_PS_ES_RE_1"
			    	Call fnE2E_24_PS_ES_RE_1()
				Case "SFDC_CreateOpportunity"
					Call fnSFDC_SFDC_CreateOpportunity()
				Case "Create_Quote"
				    Call fnCreateQuote()
				Case "Create_Order"
					Call fnCreateOrder()
			
		      End Select

		End If
		
	Next
	
Next


''''''Please  Modify the Folder Name if Any changes required for the folder names
 ''''*****************************************************************************************************************************************
'Script  for Selecting the Libary files Under a folder
'Name 		  :	 fnAddallresourses
'Input		  : 	   			
'Created By	  : 	 Shyju Kumar
'Created On	  :	 July 27,2017
'''Description: Add All the libary files Placed under a fiolder
'Review: 
'''''******************************************************************************************************************************************
Public Function fnAddallresourses()
   '''Add EnvironmentVariable xml on runtime
  strPathEnv=Split( Environment("TestDir"),Environment("TestName"))
  strFrameWorkPathBasetemp= strPathEnv(0)
  strFrameWorkPath=Split(strFrameWorkPathBasetemp,"QTP_AppDrivers")
  strFrameWorkPathBase=strFrameWorkPath(0)
   Environment.LoadFromFile(strFrameWorkPathBase&"EnvironmentVariables\EnvironmentVariables.xml")
   '''get the base path of Folders
   ''''get all folders path
    strOR=strFrameWorkPathBase&Environment.Value("ORFolder")
	strFP = strFrameWorkPathBase&Environment.Value("FRFolder")
	strDPV =  strFrameWorkPathBase&Environment.Value("DAPFolder")
	strAF = strFrameWorkPathBase&Environment.Value("APFFolder")
	strRF = strFrameWorkPathBase&Environment.Value("RFFolder")
	strRP = strFrameWorkPathBase&Environment.Value("TRFolder")
	Environment.Value("TResFolder")=strRP
	strTDP = strFrameWorkPathBase&Environment.Value("TDFolder")
	strFilePath = strTDP&Environment("SummaryExcelName")
	Environment.value("strFilePath")=strFilePath
	Environment.value("ScreenshotFolder")=strFrameWorkPathBase&Environment.Value("TRScreen")
	''Execute all Library Files
	 call fnExecuteLibraryFiles(strFP)
	 call fnExecuteLibraryFiles(strDPV)
	 call fnExecuteLibraryFiles(strAF)
	 call fnExecuteLibraryFiles(strRF)

	 	'''Add object repository on run time
   call fnAddObjectRepository(strOR)
End Function


''''*****************************************************************************************************************************************
'Script  for Executing the Library files on Run time
'Name 		  :	 fnExecuteLibraryFiles
'Input		  : 	   			strPath
'Created By	  : 	 Shyju Kumar
'Created On	  :	 July 20,2017
'''Description: to Execute the Libary files under a folder
'Reveiw:
'''''******************************************************************************************************************************************
Public Function fnExecuteLibraryFiles(strPath)
Dim objFSO,objFile, objFilesCnt, strFileName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFolder(strPath)
Set objFilesCnt = objFile.Files
valFileCnt=objFile.Files.Count
If valFileCnt=0 Then
	Reporter.ReportEvent micFail,"Library Files","libary Files not exists in the path"&strPath
	Exit Function
End If
For Each FileName in objFilesCnt
		strFileName=FileName.name 
		strFileNameFinal=Split(strFileName,".") 
		If lcase(strFileNameFinal(1))="qfl" and ubound(strFileNameFinal)=1 Then
		LoadFunctionLibrary strPath&strFileName
		End If
Next

End Function
