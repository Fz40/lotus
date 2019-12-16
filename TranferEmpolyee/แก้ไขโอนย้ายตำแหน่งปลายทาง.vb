Sub Click(Source As Button)
	
	Dim session As New NotesSession
	Dim ws As New NotesUIWorkspace
	Dim db As notesdatabase
	Dim doc As notesdocument
	Dim viewEmployeeUpdate As NotesView
	Dim uidoc As NotesUIDocument
	Dim uiview As NotesUIView
	Dim curdoc As NotesDocument
	Dim allrows As Integer
	
    'Set uidoc = ws.CurrentDocument
	'Set curdoc = uidoc.Document
	Set session = New NotesSession
	Set db = session.currentdatabase
	Set dc = db.unprocesseddocuments
	Set doc = dc.getfirstdocument
	allrows = dc.count'vc.count
	If allrows = 0 Then
		Msgbox "Please select the documents for export",0+64,"Select the Documents Required for Export"
		Exit Sub
	End If
	
	Dim DiaConfirm As NotesDocument
	Dim docEmployee As NotesDocument
	Set DiaConfirm = db.CreateDocument
	Call doc.CopyAllItems( DiaConfirm, True )
	'DiaConfirm.AuthorizeView1 = "SubStaffDetail2"
	DiaConfirm.AuthorizeView1 = "SubStaffTransferDesForDia"
	'DiaConfirm.EmbEmployeeStaffCard =""
	
	Set uiview = ws.CurrentView
	Call uiview.DeselectAll
	
	Temp = ws.DialogBox("DiaInputEmbGroupEmployee", True, True,False,True, False, False, "ขณะนี้ท่านกำลังกรอกข้อมูลในส่วน Requester", DiaConfirm,True,False)
	If Temp = True Then		
		
		Set doc = dc.getfirstdocument
		DiaConfirm.form = "EmbGroupEmployee"
		DiaConfirm.AuthorizeView = ""
		While Not (doc Is Nothing)
			
			Set viewEmployeeUpdate = db.GetView("EmpEmployeeEditDetail")
			Set docEmployee =  viewEmployeeUpdate.GetDocumentByKey(doc.MainDocID(0)+"^"+doc.EmpCode(0),True)
			If Not docEmployee Is Nothing Then
				docEmployee.EmpDestCompanyID = DiaConfirm.EmpDestCompanyID
				docEmployee.EmpDestCompanyID_1 = DiaConfirm.EmpDestCompanyID_1
				docEmployee.EmpDestCompany = DiaConfirm.EmpDestCompany
				docEmployee.EmpDestGroupID = DiaConfirm.EmpDestGroupID
				docEmployee.EmpDestGroupID_1 = DiaConfirm.EmpDestGroupID_1
				docEmployee.EmpDestGroup = DiaConfirm.EmpDestGroup
				docEmployee.EmpDestFieldID = DiaConfirm.EmpDestFieldID
				docEmployee.EmpDestFieldID_1 = DiaConfirm.EmpDestFieldID_1
				docEmployee.EmpDestField = DiaConfirm.EmpDestField
				docEmployee.EmpDestDepartmenID = DiaConfirm.EmpDestDepartmenID
				docEmployee.EmpDestDepartmenID_1 = DiaConfirm.EmpDestDepartmenID_1
				docEmployee.EmpDestDepartmen = DiaConfirm.EmpDestDepartmen
				docEmployee.EmpDestSectionsID = DiaConfirm.EmpDestSectionsID
				docEmployee.EmpDestSectionsID_1 = DiaConfirm.EmpDestSectionsID_1
				docEmployee.EmpDestSections = DiaConfirm.EmpDestSections
				docEmployee.EmpDestDivisionID = DiaConfirm.EmpDestDivisionID
				docEmployee.EmpDestDivisionID_1 = DiaConfirm.EmpDestDivisionID_1
				docEmployee.EmpDestDivision = DiaConfirm.EmpDestDivision
				docEmployee.EmpDestBranchID = DiaConfirm.EmpDestBranchID
				docEmployee.EmpDestBranchID_1 = DiaConfirm.EmpDestBranchID_1
				docEmployee.EmpDestBranch = DiaConfirm.EmpDestBranch
				docEmployee.EmpDestLevelsID = DiaConfirm.EmpDestLevelsID
				docEmployee.EmpDestLevelsID_1 = DiaConfirm.EmpDestLevelsID_1
				docEmployee.EmpDestLevels = DiaConfirm.EmpDestLevels
				docEmployee.EmpDestPositionID = DiaConfirm.EmpDestPositionID
				docEmployee.EmpDestPositionID_1 = DiaConfirm.EmpDestPositionID_1
				docEmployee.EmpDestPosition = DiaConfirm.EmpDestPosition
				docEmployee.EmpDestFrontYard = DiaConfirm.EmpDestFrontYard
				docEmployee.EmpDestStartDate = DiaConfirm.EmpDestStartDate
				docEmployee.EmpDestToDate = DiaConfirm.EmpDestToDate
				docEmployee.EmpDestType = DiaConfirm.EmpDestType
				docEmployee.EmpDestDetail = DiaConfirm.EmpDestDetail
				Call docEmployee.ComputeWithForm(True,True)
				Call docEmployee.Save(True,True)
			End If
			
			
			Set doc = dc.GetNextDocument(doc)	
		Wend
		Msgbox "บันทึกเสร็จ",0+64,"บันทึก"
	End If
	
	'Call uidoc.Refresh()
	
	
End Sub