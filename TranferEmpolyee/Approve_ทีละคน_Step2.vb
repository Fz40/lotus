Sub Click(Source As Button)
	Dim ws As New NotesUiWorkspace	
	Dim curdoc As NotesDocument	
	Dim uidoc As NotesUiDocument
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	Dim db As NotesDatabase
	Dim s As New NotesSession
	Dim docEmployee As NotesDocument	
	
	Set db = s.CurrentDatabase
	
	'validateflag = validateFieldMainFormReqPartnerCampaign(curdoc,curdoc.step(0),uidoc,"Approve")
	
	'If validateflag = False Then 'validate ของ Field ที่อยู่ในหน้าจอ
		'Exit Sub
	'End If
	
	
	Set uiview = ws.CurrentView

	If Not curdoc Is Nothing Then        
		
		Dim dateTime As New NotesDateTime(GetServerDateTime )
		Set viewEmployeeUpdate = db.GetView("EmpEmployeeEditDetail")
		Set docEmployee =  viewEmployeeUpdate.GetDocumentByKey(curdoc.MainDocID(0)+"^"+curdoc.FindEmployeeID(0),True)
		If Not docEmployee Is Nothing Then
			docEmployee.EmbCBDeductedPayroll = curdoc.EmbCBDeductedPayroll(0)
			docEmployee.EmbCBChoice = curdoc.EmbCBChoice(0)
			docEmployee.EmbCBPayin = curdoc.EmbCBPayin(0)
			docEmployee.EmbCBDeducted = curdoc.EmbCBDeducted(0)
			docEmployee.EmbNewSaraly = curdoc.EmbNewSaraly(0)
			docEmployee.EmbOtherSaraly = curdoc.EmbOtherSaraly(0)
			docEmployee.EmbSaralyRemark = curdoc.EmbSaralyRemark(0)
			
			docEmployee.EmbCBStampStatus = "Approve"
			docEmployee.EmbCBStampName = s.CommonUserName
			docEmployee.EmbCBStampDate = dateTime.LSLocalTime
			
			
		End If
	End If
	
	'Call GetWorker(curdoc) 'ดึงคนทำงานมาใส่ใน Form
	'Call ClearField(curdoc,curdoc.Step(0),curdoc) 'Clear ค่า Field ใน Flow
	
	If checkAllCompleteEmpGroupEmployee(curdoc,"CB") = True Then
		'Call ControlFlowStep(curdoc, "Approve",dateTime)
		'Call UpdateSLAStatus(curdoc,"3","Accept")
		'Call calculateAutoApprove(curdoc,"WORKER")
		'Call setIntializeDocumentEmbReviewer(curdoc,db,"Header") 'update step ของ form EmbReviewer
		'Call calculateAutoApprove(curdoc,"Header")
	End If	
	
	Call ClearFieldMainFrom(curdoc, "CB")
	'Call UpdatembReviewerAndEmbRequestDetail(curdoc,db,curdoc)
	
	'Call curdoc.Save(True,True)
	Call docEmployee.Save(True,True)
	'Call doc.Save(True,True)
	Call uidoc.Refresh
	Call uidoc.Save
	'Call uidoc.Close	
	Call ws.ViewRefresh	
	
End Sub