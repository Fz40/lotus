%REM
	Agent (FormCEOApprove)
	Created Jun 13, 2019 by Sarawut Somton/PTG
	Description: Comments for Agent
%END REM
Option Public
Option Declare
Use "CommonScript"

Sub Initialize
	
	Dim workspace As New NotesUIWorkspace
	Dim curdoc As NotesDocument	
	Dim uidoc As NotesUIDocument
	Set uidoc = workspace.CurrentDocument
	Set curdoc = uidoc.Document
	
	Dim sess As New NotesSession
	Dim view As NotesView
	Dim doc As NotesDocument
	Dim embView As NotesView
	Dim colHead As NotesDocumentCollection	
	Dim docEmp As NotesDocument
	
	Dim rtitem As NotesRichTextItem
	Dim rtitemDetail As NotesRichTextItem
	Dim rtitemGetPrivileges1 As NotesRichTextItem
	Dim rtitemCondition1 As NotesRichTextItem

	
	Dim reView As NotesView
	Dim reDoc As NotesDocument
	Dim emb As NotesEmbeddedObject 
	
	Dim tmpDay As Variant
	Dim AKeys As Variant
	Dim filename As Variant
	Dim keywordfield As String
	Dim comment As Variant
	
	Dim s As New NotesSession
	Dim db As NotesDatabase


	Set db = s.CurrentDatabase
	Set reView = db.GetView("ReportTemplate")
	
	Set embView = db.GetView("EmbReviewerHeaderByMainDocID")
	Set colHead = embView.GetAllDocumentsByKey(curdoc.MainDocID(0),True)
	
	If colHead.Count > 7 Then
		Set reDoc = reView.Getdocumentbykey("ReportCEOApprove_15", True)
	Else
		Set reDoc = reView.Getdocumentbykey("ReportCEOApprove_7", True)
	End If

	Set emb = redoc.GetAttachment(redoc.FileNameReport(0))					
	If emb Is Nothing Then															
		MsgBox "ไม่สามารถอ่านข้อมูล File Template ของตัวรายงานจากฐานข้อมูลได้ กรุณาติดต่อผู้ดูแลระบบ"
		Exit Sub
	End If	
	'------
	
	Const tempPath = "C:\TempReport"		'Path
	'On Error Resume Next
	On Error GoTo ErrorHandler	
	
	Dim WordObj 
	Dim WordDoc 

	
	Set wordObj = CreateObject("Word.Application")

	'Check tempPath are exist -------------------------------------------------------------
	Dim fs As Variant, retF As Variant
	Set fs = CreateObject("Scripting.FileSystemObject")
	retF = fs.FolderExists(TempPath)
	If Not retF Then
		Call fs.CreateFolder(TempPath)
	End If
	'---------------------------------------------------------------------------------------------
	tmpDay = Evaluate({@ReplaceSubString(@Text(@Today);"/";"-")})
	
	AKeys = "เอกสารลงนาม_Memo_CEO_" & CStr(tmpDay(0))
	filename = tempPath + "\" +  AKeys + ".docx" 		' FileName
	Call emb.ExtractFile(filename)
	
NextProcess:		

	Call wordObj.Documents.Add(filename)
	Set wordDoc = wordObj.ActiveDocument
	
	wordDoc.FormFields("DocNo").Result = uidoc.FieldGetText("DocNo")
	wordDoc.FormFields("CreateDate").Result = Format(uidoc.FieldGetText("CreateDate"),"dd/mm/yyyy")
	
	wordDoc.FormFields("Subject").Range = curdoc.Subject(0)	
	wordDoc.FormFields("FromRequester").Result = uidoc.FieldGetText("FromRequester")
	wordDoc.FormFields("Department").Result = uidoc.FieldGetText("Department")
	wordDoc.FormFields("CampaignType").Range = curdoc.CampaignType(0)
	wordDoc.FormFields("CampaignName").Range = curdoc.CampaignName(0)
	
	If uidoc.FieldGetText("CampaignStartTime") ="" Then
		wordDoc.FormFields("CampaignStartDate").Result = uidoc.FieldGetText("CampaignStartDate")
	Else
		wordDoc.FormFields("CampaignStartDate").Result = uidoc.FieldGetText("CampaignStartDate")+" "+uidoc.FieldGetText("CampaignStartTime")
	End If
	If uidoc.FieldGetText("CampaignEndTime") ="" Then
		wordDoc.FormFields("CampaignEndDate").Result = uidoc.FieldGetText("CampaignEndDate")
	Else
		wordDoc.FormFields("CampaignEndDate").Result = uidoc.FieldGetText("CampaignEndDate")+" "+uidoc.FieldGetText("CampaignEndTime")
	End If
	
	'--- ตัวอย่าง Check Box ใน Word
	Dim temp As String
	Dim tempsumtext As String
	Dim count As Integer
	
	temp = uidoc.FieldGetText("PrivilegesChannel")
	If InStr(temp,"Mobile Application") > 0 Then
		wordDoc.FormFields("MobileApplication").CheckBox.Value = True
		wordDoc.FormFields("OtherChannel").Result = "-"
	End If
	If InStr(temp,"Ussd") > 0 Then
		wordDoc.FormFields("Ussd").CheckBox.Value = True
		wordDoc.FormFields("OtherChannel").Result = "-"
	End If
	If InStr(temp,"EDC/Post") > 0 Then
		wordDoc.FormFields("EDC_Post").CheckBox.Value = True
		wordDoc.FormFields("OtherChannel").Result = "-"
	End If
	If InStr(temp,"WebSite") > 0 Then
		wordDoc.FormFields("WebSite").CheckBox.Value = True
		wordDoc.FormFields("OtherChannel").Result = "-"
	End If
	If InStr(temp,"Social") > 0 Then
		wordDoc.FormFields("Social").CheckBox.Value = True
		wordDoc.FormFields("OtherChannel").Result = "-"
	End If
	If InStr(temp,"อื่น ๆ") > 0 Then  
		wordDoc.FormFields("อื่นๆ").CheckBox.Value = True	
		wordDoc.FormFields("OtherChannel").Range = uidoc.FieldGetText("OtherChannel")
	End If
	
	Set rtitem = curdoc.Getfirstitem("Detail")
	wordDoc.FormFields("Detail1").Range = rtitem.getUnformattedText()
	
	If uidoc.FieldGetText("GetPrivileges") ="-" Then
		wordDoc.FormFields("GetPrivileges1").Delete
	Else
		Set rtitem = curdoc.Getfirstitem("GetPrivileges")
		wordDoc.FormFields("GetPrivileges1").Range = rtitem.getUnformattedText()
		
		'wordDoc.FormFields("GetPrivileges1").Range = rtitemGetPrivileges1.getUnformattedText()
	End If
	
	If uidoc.FieldGetText("Condition") ="-" Then
		wordDoc.FormFields("Condition1").Delete
	Else
		Set rtitem = curdoc.Getfirstitem("Condition")
		wordDoc.FormFields("Condition1").Range = rtitem.getUnformattedText()
		
		'wordDoc.FormFields("Condition1").Range = rtitemCondition1.getUnformattedText()
		
		'wordDoc.FormFields("Condition").Result = uidoc.FieldGetText("Condition")
		'For count = 1 To 30
		'If count = 1 Then
		'temp = uidoc.FieldGetText("Condition")
		'tempsumtext = Left(temp,255)		
		'wordDoc.FormFields("Condition1").Result = tempsumtext
		'Else
		'temp = Replace(temp,tempsumtext,"")
		'tempsumtext = Left(temp,255)
		'wordDoc.FormFields("Condition"+CStr(count)).Result = tempsumtext
		'End If	
		'Next
	End If

	Dim requesterUserName As Variant
	requesterUserName = curdoc.GetItemValue("StampStageBy_1")
	If connectEmpDetail(CStr(requesterUserName(0))) = "Success" Then
		wordDoc.FormFields("Requester").Result = getThaiNameByConnectEmpDetail(CStr(requesterUserName(0)))
	Else
		wordDoc.FormFields("Requester").Result = CStr(requesterUserName(0))
	End If

	wordDoc.FormFields("CommentRequester").Result =  uidoc.FieldGetText("CommentStage_1")
	wordDoc.FormFields("DateRequester").Result = Format(uidoc.FieldGetText("StatusStageDate_1"),"dd/mm/yyyy")
	
	Dim Emp_Loop As Integer
	Emp_Loop = 1
	If colHead.Count > 0 Then
		Set docEmp = colHead.GetFirstDocument	
		
		While Not docEmp Is Nothing
			
			If docEmp.AuthorizeView(0) = "Support" Then
				keywordfield = "Supporter"
			ElseIf docEmp.AuthorizeView(0) = "Accessory" Then
				keywordfield = "Accessory"
			ElseIf docEmp.AuthorizeView(0) = "HeadSupport" Then
				keywordfield = "HeadSupporter"
			ElseIf docEmp.AuthorizeView(0) = "HeadAccessory" Then
				keywordfield = "HeadAccessory"
			Else
				Exit Sub
			End If
			
			Dim commonUserName As Variant
			Dim StageDate As Variant
			Dim comments As Variant
			
			commonUserName = docEmp.GetItemValue(keywordfield+"_StampStageBy")
			StageDate  = docEmp.GetItemValue(keywordfield+"_StatusStageDate")
			comments = docEmp.GetItemValue(keywordfield+"_CommentStage")
			docEmp.GetItemValue(keywordfield+"_CommentStage")

			If connectEmpDetail(CStr(commonUserName(0))) = "Success" Then
				wordDoc.FormFields("Head"+CStr(Emp_Loop)).Result = getThaiNameByConnectEmpDetail(CStr(commonUserName(0)))
				wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Result = getJobFunctionByConnectEmpDetail(CStr(commonUserName(0)))
				wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Result = Format(StageDate(0),"dd/mm/yyyy")
				wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Result = "เห็นชอบในบทบาท "+docEmp.BUWorkerName(0)+" ความเห็นเพิ่มเติม "+CStr(comments(0))
				Emp_Loop = Emp_Loop+1
			Else
				If CStr(commonUserName(0)) <> "Auto Approve" Then
					wordDoc.FormFields("Head"+CStr(Emp_Loop)).Result = getThaiNameByConnectEmpDetail(CStr(commonUserName(0)))
					wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Result = ""
					wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Result = Format(StageDate(0),"dd/mm/yyyy")
					wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Result = "เห็นชอบในบทบาท "+docEmp.BUWorkerName(0)+" ความเห็นเพิ่มเติม "+CStr(comments(0))
					Emp_Loop = Emp_Loop+1
				End If
			End If
			
	%REM
			Dim StageDate,DepartmentName As Variant
			DepartmentName = getDepartment(docEmp.WorkerName(0))
			comment = docEmp.GetItemValue(keywordfield+"_CommentStage")
			StageDate  =docEmp.GetItemValue(keywordfield+"_StatusStageDate")
			wordDoc.FormFields("Head"+CStr(Emp_Loop)).Result = docEmp.WorkerName(0)+ chr(13) + Format(StageDate(0),"dd/mm/yyyy")+Chr(13)+CStr(DepartmentName(0))
			'wordDoc.FormFields("DateHead"+CStr(Emp_Loop)).Result = Format(StageDate(0),"dd/mm/yyyy")
			'wordDoc.FormFields("CommentHead"+CStr(Emp_Loop)).Result = CStr(comment(0))
				
	%END REM
			Set docEmp = colHead.GetNextDocument(docEmp)				
		Wend		
		
	End If
	
	Dim NameCEO,PositionCEO As Variant
	
	NameCEO = getKeywordResult("CEONameInReport")
	PositionCEO = getKeywordResult("NameOfPositionCEO")
	wordDoc.FormFields("CEOName").Result = CStr(NameCEO(0))
	wordDoc.FormFields("PositisionNameTH2").Result = CStr(PositionCEO(0))
	wordDoc.FormFields("PositisionNameTH1").Result = CStr(PositionCEO(0))
	
	If colHead.Count < 7 Then
		For Emp_Loop = 1 To 7
			If wordDoc.FormFields("Head"+CStr(Emp_Loop)).Result ="" Then
				wordDoc.FormFields("Head"+CStr(Emp_Loop)).Delete
				wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Delete
				wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Delete
				wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Delete			
			End If
		Next
	Else
		For Emp_Loop = 1 To 15
			If wordDoc.FormFields("Head"+CStr(Emp_Loop)).Result ="" Then
				wordDoc.FormFields("Head"+CStr(Emp_Loop)).Delete
				wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Delete
				wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Delete
				wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Delete			
			End If
		Next
	End If

	MessageBox "Print form CEO Approve Complete. " , 64 , "Print Complete .." 	
	wordObj.visible = True
	Exit Sub
	
ErrorHandler:
	
	MsgBox "Something is error please contact Administrator",0+64,"Error ..."
	wordObj.visible = True
	
End Sub