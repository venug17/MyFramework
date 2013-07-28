'Author - Venugopal Adhikesevan
' Declaration 
Option Explicit 
'For Driver script 
Dim  vPage, vTSCnt, vTc_Id, vObjname, vAction, vValue, vStep, vComment, vExpres, vReq, vBlnstatus, vPath, dtblTI, Itr, vExpected, vScname, vScenario_Name, vHgt
'For Recovery Manager 
Dim vSShot, blnErrTrigger, vRecid, vMaskwidth, vblnReccall, vAppstatus, objPage, vErrno, dicRecman, blnErrMsg, dicLocEnv, blnPage, aItr, vstrExtractor, vstrConcat
'For Result Editor
Dim  blnWriteresult, vResultfile, objFSO, objResultFile, vRescnt, dicResultEditor, arrdict, arrItr, vResult, vSccnt, vPasscnt, vFailcnt, vStime, vEtime, rItr, blnTeststatus, vScstime, vScetime, vSctime, vSkiprowcnt
vPath = Environment.Value("TestDir")
'For Data Parser
Dim regData, vDatResult, regDataresult, dtblTD, vTDCnt, DItr, vRepdata, vstrConresult
Dim vFileName, vSheet, vExe_Id, vParse_data
Dim objRecordSet, objConnection
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = "&H0001"
'Set the default Screenshot path
Environment("vScreenpath") = Environment.Value("TestDir")&"\Images\"
Set dicRecman = CreateObject("Scripting.Dictionary")
Set dicLocEnv = CreateObject("Scripting.Dictionary")
Set dicResultEditor = CreateObject("Scripting.Dictionary")
vTc_Id = ""
vRescnt = 0
vSccnt = 0
vPasscnt = 0
vFailcnt = 0
vSkiprowcnt = 0
vStime = Time
blnTeststatus = True
'Create Result file 
Set objFSO = CreateObject("Scripting.FileSystemObject")
vResultfile = Replace(Replace(Now,":",""),"/","_")
objFSO.CopyFile vPath&"\Result_Header_Pattern.txt",vPath&"\Results\"&vResultfile&".html",True
Set objResultFile = objFSO.OpenTextFile(vPath&"\Results\"&vResultfile&".html", 8)
'Maintaining Datasheet
Datatable.AddSheet("Test_Input")
Datatable.AddSheet("Rcvry_Manager")
Datatable.AddSheet("Test_Data")
Datatable.ImportSheet vPath&"\Input.xls", "Test Execution", "Test_Input"
Datatable.ImportSheet vPath&"\Input.xls", "Test Data", "Test_Data"
Datatable.ImportSheet vPath&"\Input.xls", "RM", "Rcvry_Manager"
Set dtblTI = Datatable.GetSheet("Test_Input")
vTSCnt = dtblTI.GetRowCount
blnErrTrigger = "False"
For Itr = 1 to vTSCnt
	dtblTI.SetCurrentRow(Itr)
	vReq = datatable.Value("Auto","Test_Input")
	If Instr(1,vReq,"y",1)	 Then
		Do Until dicRecman.count = 0
		   dicRecman.RemoveAll
		Loop
		dicRecman.Add "blnErrTrigger", blnErrTrigger
		dicRecman.Add "vAppstatus","True"
		dicRecman.Add "vBlnstatus","True"
		dicRecman.Add "vSshot",""
		dicRecman.Add "vErrno",0
' ------------------------------Scenario Initializer -------------------------------------- 
			If datatable.Value("Tc_Id","Test_Input") <> ""  Then
				If vTc_Id = "" Then
					blnWriteresult = False
					Else
					blnWriteresult = True
				End If
	' ------------------------------Result File Handling-------------------------------------- 
						If blnWriteresult Then
							vScetime = Timer
							If blnTeststatus Then
								blnTeststatus = "Pass"
								vPasscnt = vPasscnt + 1
								Else
								blnTeststatus = "Fail"
								vFailcnt = vFailcnt + 1
							End If			
							vSctime = vScetime - vScstime  
							call fnWriteresult(objResultFile, dicResultEditor, vRescnt, vScenario_Name, blnTeststatus, vSctime, vSkiprowcnt)
							blnWriteresult = False
							vRescnt = 0
							blnTeststatus = True
							vSkiprowcnt = 0
							Do Until dicResultEditor.Count = 0
								dicResultEditor.RemoveAll
							Loop
						End If
	' ------------------------------Result File Handling Ends-------------------------------------- 
				vSccnt = vSccnt + 1
				vScstime = Timer
				vTc_Id = datatable.Value("Tc_Id","Test_Input")
				vScenario_Name = datatable.Value("Scenario","Test_Input")
				dicRecman.Remove "blnErrTrigger"
				dicRecman.Add "blnErrTrigger", "False"
			End If
' ------------------------------Scenario Initializer  Ends-------------------------------------- 
		vRescnt = vRescnt + 1
		vblnReccall = True
		vRecid = datatable.Value("Rec_Id","Test_Input")
		dicRecman.Add "vRecid",vRecid
		objPage = datatable.Value("Page","Test_Input")
		call fnParentBuilder(objPage)
		Set vPage =  Eval(left(objPage,len(objPage)-1))
		vObjname = datatable.Value("Object","Test_Input")
		vAction = datatable.Value("Action","Test_Input")
		vValue = datatable.Value("Value","Test_Input")
		vStep = datatable.Value("Scenario","Test_Input")
		vExpres = datatable.Value("Expected","Test_Input")
'-------------------------------Test Data Parser----------------------------------------
		vFileName = vPath&"\Input.xls"
		vSheet = "Test Data"
		Do Until Not(Instr(1,vValue,"<") > 0)
			Set regData = New Regexp
			regData.Pattern = "<\w*>"
			Set regDataresult = regData.Execute(vValue)
			vDatResult = Replace(regDataresult(0),"<","")
			vDatResult =  Replace(vDatResult,">","")
			On Error Resume Next
            Set objConnection = CreateObject("ADODB.Connection")
			objConnection.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq="&vFileName&";Readonly=False;DefaultDir="&vPath&";"
			Set objRecordSet = CreateObject("ADODB.Recordset")
			objRecordSet.LockType = 2
			objRecordset.Open "Select TestValue FROM [" & vSheet & "$] where TestVariable =  '"&vDatResult&"'", objConnection
			
			vParse_data = objRecordset.fields("TestValue")
			If vParse_data = "" Then
				vParse_data = Replace(Replace(regDataresult(0),"<",""), ">", "")
			End If
			vValue = Replace(vValue, regDataresult(0), vParse_data)
			Set objRecordSet = Nothing 
			Set objConnection = Nothing
		Loop
' ------------------------------Recovery Manager Checkpoint Starts------------------------------------- 
		If dicRecman("blnErrTrigger") or (Not(vPage.Exist(0)) and Lcase(vAction) <> "launchapp") Then
			Call fnRcvrymanager(dicRecman)
			blnErrTrigger = dicRecman("blnErrTrigger")
			vblnReccall = dicRecman("vblnReccall")
			vBlnstatus = vblnReccall
			vComment = vComment&" Error handled by Recovery"
			Do Until Not(dicRecman.Exists("vBlnstatus"))
				dicRecman.Remove("vBlnstatus")
			Loop
			dicRecman("vBlnstatus") = vBlnstatus
		End If
' ------------------------------Recovery Manager Checkpoint Ends------------------------------------- 
		Do while Instr(1,vValue,"|")
			vstrExtractor = ""
			vstrConcat = Split(vValue,"|")
			vstrConresult = dicLocEnv(vstrConcat(1))
			vValue = Replace(vValue, "|"&vstrConcat(1)&"|", vstrConresult)
		Loop
		vPage.Sync
		If vblnReccall Then
				Select Case(Lcase(vAction))
				Case "launchapp"
					Call fnlaunchapp(vPage,vValue,vComment, dicRecman)	
				Case "enterdata"
					Call fnenterdata(vPage,vObjname,vValue,vComment, dicRecman)
				Case "click"
					Call fnclick(vPage,vObjname,vValue,vComment, dicRecman)
				Case "select"
					Call fnselect(vPage,vObjname,vValue,vComment, dicRecman)
				Case "verifyproperty"
					Call fnverify(vPage,vObjname,vValue,vComment, dicRecman)
				Case "verifytablerow","verifytabledata"
					Call fnverifytable(vPage,vAction,vObjname,vValue,vComment, dicRecman)
				Case "printscreen"
					vSshot = vTc_Id
					Call fnCapturescreen(vSshot)
					vComment = "Screenshot"
					dicRecman.Remove("vSshot")
					dicRecman.Add "vSshot", vSshot
				End Select
		End If 
		vPage.Sync
		blnErrTrigger = dicRecman("blnErrTrigger")
		vResult = dicRecman("vBlnstatus")
		datatable.Value("Actual","Test_Input") = vComment	
		datatable.Value("Result","Test_Input") = vResult
		If dicRecman("vSshot") <> ""  and dicRecman("vSshot") <> "Rec_Manager" Then
			vSshot =  "'=HYPERLINK("""&dicRecman("vSshot")&"""|""Screenshot_"&vTc_Id&""")"
			datatable.Value("Screenshot","Test_Input") = vSshot
		End If
		If Not(dicRecman("vAppstatus")) Then
			vComment = "Application failed to launch/Network Error"
			datatable.Value("Actual","Test_Input") = vComment
			blnTeststatus = False
		End If
' ----------------------------Result  Editor Handling--------------------------------------------------------------------------------------- 
		vExpected = datatable.Value("Expected","Test_Input")
		vScname = datatable.Value("Step","Test_Input")
		If vScname = "" Then
			vSkiprowcnt = vSkiprowcnt + 1
		End If
		arrdict = Array("vTc_Id","vScname",  "vExpected", "vComment", "vResult", "vSshot")
		For arrItr = Lbound(arrdict) to Ubound(arrdict)
		   dicResultEditor.Add arrdict(arrItr)&vRescnt, Eval(arrdict(arrItr))
		Next
		blnTeststatus = blnTeststatus and vResult
		If Not(dicRecman("vAppstatus")) Then
				Exit For
		End If
' ----------------------------Result  Editor Handling Ends---------------------------------------------------------------------------------------
		vSshot = ""
		vComment = ""	
	End If
Next
' vSccnt, vPasscnt, vFailcnt, vStime, vEtime
vScetime = Timer
DataTable.ExportSheet vPath&"\Output.xls", "Test_Input"
'Result writing for final scenario
If blnTeststatus Then
	blnTeststatus = "Pass"
	vPasscnt = vPasscnt + 1
	Else
	blnTeststatus = "Fail"
	vFailcnt = vFailcnt + 1
End If
vSctime = vScetime - vScstime  
call fnWriteresult(objResultFile, dicResultEditor, vRescnt, vScenario_Name, blnTeststatus, vSctime, vSkiprowcnt)
vEtime = Time
Call fnWriteresultfooter(objResultFile, vSccnt, vPasscnt, vFailcnt, vStime, vEtime)
objResultFile.Close
Set objFSO = Nothing
Set dicResultEditor = Nothing
Set objResultFile = Nothing
Set dicRecman = Nothing
Set dicLocEnv = Nothing
Msgbox "Your Automation task is completed!!!!!!!"