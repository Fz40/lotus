Sub Click(Source As Button)
	Dim ws As New NotesUiWorkspace	
	Dim curdoc As NotesDocument	
	Dim uidoc As NotesUiDocument
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	Dim db As NotesDatabase
	Dim s As New NotesSession
	
	Set db = s.CurrentDatabase
	
	Dim submitflag As Integer
	
	Dim ChkWkNameBuType  As Integer
	Dim DiaDB As notesdatabase
	Dim DiaConfirm As NotesDocument
	Set DiaDB = S.CurrentDatabase
	Set DiaConfirm = DiaDB.CreateDocument
	
	
	DiaConfirm.MainDocumentNo = curdoc.DocNo(0)
	DiaConfirm.MainDocID = curdoc.MainDocID(0)
	
	If curdoc.CommentStage_1(0) <> "" Then
		DiaConfirm.App_Comment = curdoc.StatusStage_1(0)+" "+Cstr(curdoc.StatusStageDate_1(0))+" "+curdoc.CommentStage_1(0)+Chr(13)
	Else
		DiaConfirm.App_Comment = ""
	End If
	
	
	Call s.SetEnvironmentVar("AllowEdit","Yes")
	uidoc.EditMode = True
	
	If uidoc.EditMode = False Then
		Exit Sub
	End If
	
	validateflag = validateFieldMainForm(curdoc,curdoc.step(0),uidoc,"Submit")
	If validateflag = False Then 'validate ของ Field ที่อยู่ในหน้าจอ
		Exit Sub
	End If
	
	If getCountEmployee(curdoc) = 0 Then
		Msgbox "ไม่สามารถดำเนินการต่อได้ เนื่องจากไม่มี ผู้เกี่ยวข้อง หรือ ผู้สนับสนุน กรุณาติดต่อผู้ดูแลระบบ" ,0+48, "Alert !! "
		Exit Sub
	End If
	
	
	Temp = ws.DialogBox("DiaConfirmApprover", True, True,False,True, False, False, "คุณต้องการส่งงานไปให้ผู้เกี่ยวข้องดำเนินการต่อ เอกสารเลขที่ :  " + curdoc.DocNo(0)+ " ใช่หรือไม่?", DiaConfirm,True,False)
	If Temp = True Then
		'Call s.SetEnvironmentVar("AllowEdit","Yes")
		
		Dim dateTime As New NotesDateTime(GetServerDateTime )
		
		
		Call ClearField(curdoc,"clearall",curdoc) 'Clear ค่า Field ใน Flow
		Call clearFieldStampDocumentEmbReviewer(curdoc,DiaDB)
		Call GetWorker(curdoc) 'ดึงคนทำงานมาใส่ใน Form
		
		'Call setIntializeDocumentEmbReviewer(curdoc,db,"Worker") 'update step ของ form EmbReviewer
		'Call setIntializeDocumentEmbReviewer(curdoc,db,"Header") 'update step ของ form EmbReviewer
		
		curdoc.CommentStage_1 = DiaConfirm.App_Comment(0)
		curdoc.StatusStage_1 = "Submit"
		curdoc.StampStageBy_1 = s.CommonUserName
		curdoc.StatusStageDate_1 = dateTime.LSLocalTime
		'curdoc.TimeStampDate_1 = Cdat(dateTime.DateOnly) ยังไม่คำนวน SLA ตอนนี้
		
		If curdoc.CheckRequestCEO(0) = "Y" Then
			curdoc.DisplayFlow = "DisplayFlow2"
		Else
			curdoc.DisplayFlow = "DisplayFlow3"
		End If
		
		If curdoc.DocNo(0) <> "" Then
			Call ControlFlowStep(curdoc, "Submit",dateTime)
			Call calculateAutoApprove(curdoc,"WORKER")
			Call uidoc.Refresh
			
'			Call CalculateSLA(curdoc,curdoc.step(0)) ยังไม่คำนวน SLA ตอนนี้
		End If
		
		
		Call curdoc.Save(True,True)
		Call uidoc.Refresh
		Call uidoc.Save
		Call uidoc.Close	
		
		Call ws.ViewRefresh
		
		If curdoc.DocNo(0) = "" Then
			Dim agent As NotesAgent
			Set agent = _
			db.GetAgent("GenerateDocNo")
			If agent.RunOnServer(curdoc.NoteID) = 0 Then
				Print "Agent ran",, "Success"
			Else
				Print "Agent did not run",, "Failure"
			End If
		End If	
		
	End If
End Sub