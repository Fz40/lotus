%REM
	Agent (FormCEOApprove)
	Created Jun 13, 2019 by Sarawut Somton/PTG
	Description: Comments for Agent
%END REM
Option Public
Option Declare
Use "CommonScript"
Dim WordObj 
Dim WordDoc 
	

Sub Initialize
	
	Dim workspace As New NotesUIWorkspace
	Dim uidoc As NotesUIDocument
	Dim curdoc As NotesDocument	
	
	Set uidoc = workspace.CurrentDocument
	Set curdoc = uidoc.Document

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

	Dim Emp_Loop,CountStampedEmpty As Integer
	Dim commonUserName As Variant
	Dim StageDate As Variant
	Dim comments As Variant
	Emp_Loop = 1

	Set db = s.CurrentDatabase
	Set reView = db.GetView("ReportTemplate")
	
	Set embView = db.GetView("EmbReviewerHeaderByMainDocID")
	Set colHead = embView.GetAllDocumentsByKey(curdoc.MainDocID(0),True)

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

		commonUserName = docEmp.GetItemValue(keywordfield+"_StampStageBy")
		If CStr(commonUserName(0)) <> "Auto Approve" Then
			CountStampedEmpty = CountStampedEmpty+1
		End If
		Set docEmp = colHead.GetNextDocument(docEmp)				
	Wend
	If CountStampedEmpty > 7 Then
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
	
	wordDoc.FormFields("DocNo").Range = uidoc.FieldGetText("DocNo")
	wordDoc.FormFields("CreateDate").Range = Format(uidoc.FieldGetText("CreateDate"),"dd/mm/yyyy")
	
	wordDoc.FormFields("Subject").Range = curdoc.Subject(0)	
	'wordDoc.FormFields("FromRequester").Result = uidoc.FieldGetText("FromRequester")
	wordDoc.FormFields("FromRequester").Range = curdoc.FromRequester(0)
	
	wordDoc.FormFields("Department").Range = uidoc.FieldGetText("Department")
	wordDoc.FormFields("CampaignType").Range = curdoc.CampaignType(0)
	wordDoc.FormFields("CampaignName").Range = curdoc.CampaignName(0)
	
	If uidoc.FieldGetText("CampaignStartTime") ="" Then
		wordDoc.FormFields("CampaignStartDate").Range = uidoc.FieldGetText("CampaignStartDate")
	Else
		wordDoc.FormFields("CampaignStartDate").Range = uidoc.FieldGetText("CampaignStartDate")+" "+uidoc.FieldGetText("CampaignStartTime")
	End If
	If uidoc.FieldGetText("CampaignEndTime") ="" Then
		wordDoc.FormFields("CampaignEndDate").Range = uidoc.FieldGetText("CampaignEndDate")
	Else
		wordDoc.FormFields("CampaignEndDate").Range = uidoc.FieldGetText("CampaignEndDate")+" "+uidoc.FieldGetText("CampaignEndTime")
	End If
	
	'--- ตัวอย่าง Check Box ใน Word
	Dim temp As String
	Dim tempsumtext As String
	Dim count As Integer
	
	temp = uidoc.FieldGetText("PrivilegesChannel")
	If InStr(temp,"Mobile Application") > 0 Then
		wordDoc.FormFields("MobileApplication").CheckBox.Value = True
	End If
	If InStr(temp,"USSD") > 0 Then
		wordDoc.FormFields("Ussd").CheckBox.Value = True
	End If
	If InStr(temp,"EDC/Post") > 0 Then
		wordDoc.FormFields("EDC_Post").CheckBox.Value = True
	End If
	If InStr(temp,"WebSite") > 0 Then
		wordDoc.FormFields("WebSite").CheckBox.Value = True
	End If
	If InStr(temp,"Social") > 0 Then
		wordDoc.FormFields("Social").CheckBox.Value = True
	End If
	If InStr(temp,"อื่น ๆ") > 0 Then  
		wordDoc.FormFields("อื่นๆ").CheckBox.Value = True	
		wordDoc.FormFields("OtherChannel").Range = uidoc.FieldGetText("OtherChannel")
	Else
		wordDoc.FormFields("OtherChannel").Range = "-"
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
		wordDoc.FormFields("Requester").Range = getThaiNameByConnectEmpDetail(CStr(requesterUserName(0)))
	Else
		wordDoc.FormFields("Requester").Range = CStr(requesterUserName(0))
	End If

	wordDoc.FormFields("CommentRequester").Range =  uidoc.FieldGetText("CommentStage_1")
	wordDoc.FormFields("DateRequester").Range = Format(uidoc.FieldGetText("StatusStageDate_1"),"dd/mm/yyyy")
	

	
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
			
			commonUserName = docEmp.GetItemValue(keywordfield+"_StampStageBy")
			StageDate  = docEmp.GetItemValue(keywordfield+"_StatusStageDate")
			comments = docEmp.GetItemValue(keywordfield+"_CommentStage")
			docEmp.GetItemValue(keywordfield+"_CommentStage")
			If connectEmpDetail(CStr(commonUserName(0))) = "Success" Then
				wordDoc.FormFields("Head"+CStr(Emp_Loop)).Range = getThaiNameByConnectEmpDetail(CStr(commonUserName(0)))
				wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Range = getJobFunctionByConnectEmpDetail(CStr(commonUserName(0)))
				wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Range = Format(StageDate(0),"dd/mm/yyyy")
				wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Range = "เห็นชอบในบทบาท "+docEmp.BUWorkerName(0)+" ความเห็นเพิ่มเติม "+CStr(comments(0))
                Emp_Loop = Emp_Loop+1
			ElseIf CStr(commonUserName(0)) <> "Auto Approve" Then
                wordDoc.FormFields("Head"+CStr(Emp_Loop)).Range = getThaiNameByConnectEmpDetail(CStr(commonUserName(0)))
                wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Range = ""
                wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Range = Format(StageDate(0),"dd/mm/yyyy")
                wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Range = "เห็นชอบในบทบาท "+docEmp.BUWorkerName(0)+" ความเห็นเพิ่มเติม "+CStr(comments(0))
                Emp_Loop = Emp_Loop+1
			
			End If
	%REM
			Dim StageDate,DepartmentName As Variant
			DepartmentName = getDepartment(docEmp.WorkerName(0))
			comment = docEmp.GetItemValue(keywordfield+"_CommentStage")
			StageDate = docEmp.GetItemValue(keywordfield+"_StatusStageDate")
			wordDoc.FormFields("Head"+CStr(Emp_Loop)).Result = docEmp.WorkerName(0)+ chr(13) + Format(StageDate(0),"dd/mm/yyyy")+Chr(13)+CStr(DepartmentName(0))
			'wordDoc.FormFields("DateHead"+CStr(Emp_Loop)).Result = Format(StageDate(0),"dd/mm/yyyy")
			'wordDoc.FormFields("CommentHead"+CStr(Emp_Loop)).Result = CStr(comment(0))
				
	%END REM
			Set docEmp = colHead.GetNextDocument(docEmp)				
		Wend	
		If CountStampedEmpty < 7 Then
			For Emp_Loop = CountStampedEmpty+1 To 7
				wordDoc.FormFields("Head"+CStr(Emp_Loop)).Range = ""
				wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Range = ""
				wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Range = ""
				wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Range = ""	
			Next
		ElseIf (CountStampedEmpty > 7) And (CountStampedEmpty < 15) Then
			For Emp_Loop = CountStampedEmpty+1 To 15
				wordDoc.FormFields("Head"+CStr(Emp_Loop)).Range = ""
				wordDoc.FormFields("HeadPosition"+CStr(Emp_Loop)).Range = ""
				wordDoc.FormFields("HeadDate"+CStr(Emp_Loop)).Range = ""
				wordDoc.FormFields("HeadComment"+CStr(Emp_Loop)).Range = ""	
			Next

		End If

	End If
	
	Dim NameCEO,PositionCEO As Variant

	NameCEO = getKeywordResult("CEONameInReport")
	PositionCEO = getKeywordResult("NameOfPositionCEO")
	wordDoc.FormFields("CEOName").Range = CStr(NameCEO(0))
	wordDoc.FormFields("PositisionNameTH2").Range = CStr(PositionCEO(0))
	wordDoc.FormFields("PositisionNameTH1").Range = CStr(PositionCEO(0))
	
	MessageBox "Print form CEO Approve Complete. " , 64 , "Print Complete .." 	
	wordObj.visible = True
	Exit Sub
	
ErrorHandler:
	
	MsgBox "Something is error please contact Administrator",0+64,"Error ..."
	wordObj.visible = True
	
End Sub

Function FormFieldExists(FormFieldName As String) As Boolean
	FormFieldExists=False
	ForAll aField In WordDoc.FormFields
		If aField.Name = FormFieldName Then
			FormFieldExists=True
			Exit Function
		End If
	End ForAll
End Function
Function CustomPropertyExists(CustomPropertyName As String) As Boolean
CustomPropertyExists=False
	ForAll propCustom In WordDoc.CustomDocumentProperties
		If propCustom.Name = CustomPropertyName Then
			CustomPropertyExists=True
			Exit Function
		End If
	End ForAll
End Function