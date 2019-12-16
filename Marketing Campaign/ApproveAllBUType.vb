Sub Click(Source As Button)
	Dim ws As New NotesUiWorkspace	
	Dim curdoc As NotesDocument	
	Dim uidoc As NotesUiDocument
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	Dim db As NotesDatabase
	Dim s As New NotesSession
	
	Set db = s.CurrentDatabase
	
	validateflag = validateFieldMainFormReqPartnerCampaign(curdoc,curdoc.step(0),uidoc,"Approve")
	
	If validateflag = False Then 'validate ของ Field ที่อยู่ในหน้าจอ
		Exit Sub
	End If
	
	Dim viewWorker As NotesView
	Dim colWorker As NotesDocumentCollection	
	Dim docWorker As NotesDocument
	
	Set viewWorker = db.GetView("EmbReviewerWorkerStepOnProcessByMainDocID")
	Set colWorker = viewWorker.GetAllDocumentsByKey(curdoc.MainDocID(0),True)
	
	If colWorker.Count = 0 Then
		Msgbox "ไม่สามารถดำเนินการต่อได้ เนื่องจากไม่มีผู้ดำเนินการต่อ กรุณาติดต่อผู้ดูแลระบบ" ,0+48, "Alert !! "
		Exit Sub
	End If
	
	Set docWorker = colWorker.GetFirstDocument
	
	Dim eval As Variant
	Dim AuthorizeView() As String
	Dim i As Integer
	i=0
	While Not docWorker Is Nothing
		eval = Evaluate({@Contains(@Name([cn];WorkerName);@Name([cn];@UserName))}, docWorker) 'Company code <> PTG
		If eval(0) = 1 Then
			Redim Preserve AuthorizeView(i)
			AuthorizeView(i) = docWorker.BUWorkerType(0)+"^"+docWorker.BUWorkerName(0)
			i = i +1
		End If		
		Set docWorker = colWorker.GetNextDocument(docWorker)		
	Wend
	
	Dim AuthorizeViewSelected As Variant
	
	AuthorizeViewSelected = ws.Prompt(PROMPT_OKCANCELLISTMULT, _
	"เลือกรายการ ที่ต้องการ Approve", _
	"เลือกรายการ ที่ต้องการ 1 หรือ มากกว่านั้นจากตัวเลือกด้านล่าง", _
	AuthorizeView, AuthorizeView)
	
	If Isempty(AuthorizeViewSelected) Then
		Exit Sub
	End If
	
	Dim DiaDB As notesdatabase
	Dim DiaConfirm As NotesDocument
	Set DiaDB = S.CurrentDatabase
	Set DiaConfirm = DiaDB.CreateDocument
	
	'++++++++   start คำนวนหาชื่อ Form Dialog กับ คำนวนหาชื่อ Field ใน Step ถัดไป
	
	Dim CalculateDialogAndFieldNextStep As String
	Dim DialogFormName As String
	Dim FieldNextStep As String
	CalculateDialogAndFieldNextStep = CalculateDialogFormName(curdoc,"Approve")
	DialogFormName = Strtoken(CalculateDialogAndFieldNextStep, "^",1)
	FieldNextStep = Strtoken(CalculateDialogAndFieldNextStep, "^",2)
	
	DiaConfirm.FlowName = curdoc.FlowName(0)
	DiaConfirm.BUTypeRequester = curdoc.SaleOrganizationCode(0)
	DiaConfirm.ChoiceStep = curdoc.GetItemValue(FieldNextStep)
	DiaConfirm.SeletctStep = DiaConfirm.ChoiceStep(0)
	
'++++++++   end คำนวนหาชื่อ Form Dialog กับ คำนวนหาชื่อ Field ใน Step ถัดไป
	
	DiaConfirm.MainDocumentNo = curdoc.DocNo(0)
	DiaConfirm.App_Comment = ""
	
	Temp = ws.DialogBox(DialogFormName, True, True,False,True, False, False, "คุณต้องการส่งงาน (Approve) เอกสารเลขที่ :  " + curdoc.DocNo(0)+ " ใช่หรือไม่?", DiaConfirm,True,False)
	If Temp = True Then
		Call s.SetEnvironmentVar("AllowEdit","Yes")
		uidoc.EditMode = True
		Dim dateTime As New NotesDateTime(GetServerDateTime )
		
		Forall tmpAuthorizeViewSelected In AuthorizeViewSelected
			Dim viewEmpWorkerUpdateStatus As NotesView
			Dim docEmpWorker As NotesDocument
			Set viewEmpWorkerUpdateStatus = db.GetView("EmbReviewerWorkerByMainDocIDAndBUWorkerTypeAndBUWorkerName")
			Set docEmpWorker =  viewEmpWorkerUpdateStatus.GetDocumentByKey(curdoc.MainDocID(0)+"^"+tmpAuthorizeViewSelected,True)
			
			If Not docEmpWorker Is Nothing Then
				Call s.SetEnvironmentVar("AllowEdit","Yes")
				uidoc.EditMode = True 
				
				Dim keywordfield As String
				If docEmpWorker.AuthorizeView(0) = "Support" Then
					keywordfield = "Supporter"
				Elseif docEmpWorker.AuthorizeView(0) = "Accessory" Then
					keywordfield = "Accessory"
				Elseif docEmpWorker.AuthorizeView(0) = "HeadSupport" Then
					keywordfield = "HeadSupporter"
				Elseif docEmpWorker.AuthorizeView(0) = "HeadAccessory" Then
					keywordfield = "HeadAccessory"
				Else
					Msgbox "ไม่สามารถดำเนินการต่อได้ เนื่องจากไม่มีผู้ดำเนินการต่อ กรุณาติดต่อผู้ดูแลระบบ" ,0+48, "Alert !! "
					Exit Sub
				End If
				
				docEmpWorker.ReplaceItemValue "MainDocID",  curdoc.MainDocID(0)
				docEmpWorker.OpenDialogBox = "Y"
				
				Call ClearField(docEmpWorker,docEmpWorker.Step(0),curdoc) 'Clear ค่า Field ใน Flow
				
				docEmpWorker.Step = "3"
				docEmpWorker.JobStatusName = "Complete"
				docEmpWorker.OpenDialogBox = ""
				
				docEmpWorker.ReplaceItemValue keywordfield+"_CommentStage",  DiaConfirm.App_Comment(0)
				docEmpWorker.ReplaceItemValue keywordfield+"_StatusStage",  "Approve"
				docEmpWorker.ReplaceItemValue keywordfield+"_StampStageBy", s.CommonUserName
				docEmpWorker.ReplaceItemValue keywordfield+"_StatusStageDate", dateTime.LSLocalTime
				Call docEmpWorker.ComputeWithForm(True,True)
				Call docEmpWorker.Save(True,True)
				
			End If		
		End Forall	
	Else
		Exit Sub
	End If
	
	
	
	
	
	Call GetWorker(curdoc) 'ดึงคนทำงานมาใส่ใน Form
	Call ClearField(curdoc,curdoc.Step(0),curdoc) 'Clear ค่า Field ใน Flow
	
	If checkAllCompleteEmpReviewer(curdoc,"Worker") = True Then
		
		Call ControlFlowStep(curdoc, "Approve",dateTime)
		'Call UpdateSLAStatus(curdoc,"3","Accept")
		Call calculateAutoApprove(curdoc,"WORKER")
		Call setIntializeDocumentEmbReviewer(curdoc,db,"Header") 'update step ของ form EmbReviewer
		Call calculateAutoApprove(curdoc,"Header")
	Else
		Call setFieldWorker(curdoc)	
		'Call SendMail(curdoc)	
	End If	
	Call UpdatembReviewerAndEmbRequestDetail(curdoc,db,curdoc)
	
	Call curdoc.Save(True,True)
	Call uidoc.Refresh
	Call uidoc.Save
	Call uidoc.Close	
	Call ws.ViewRefresh		
End Sub