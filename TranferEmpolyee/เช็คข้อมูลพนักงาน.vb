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
	
	Set uiview = ws.CurrentView
	Call uiview.DeselectAll

    Set doc = dc.getfirstdocument
    DiaConfirm.form = "EmbGroupEmployee"
    DiaConfirm.AuthorizeView = ""
    While Not (doc Is Nothing)
        
        Set viewEmployeeUpdate = db.GetView("EmpEmployeeEditDetail")
        Set docEmployee =  viewEmployeeUpdate.GetDocumentByKey(doc.MainDocID(0)+"^"+doc.EmpCode(0),True)
        If Not docEmployee Is Nothing Then

             DiaConfirm.EmpName = docEmployee.EmpName(0)
             DiaConfirm.EmpCode = docEmployee.EmpCode(0)
             DiaConfirm.company = docEmployee.Company(0)
             DiaConfirm.CompanyGroup = docEmployee.CompanyGroup(0)
             DiaConfirm.CompanyLines = docEmployee.CompanyLines(0)
             DiaConfirm.Department = docEmployee.Department(0)
             DiaConfirm.Sections = docEmployee.Sections(0)
             DiaConfirm.Division = docEmployee.Division(0)
             DiaConfirm.Location = docEmployee.Location(0)
             DiaConfirm.Position = docEmployee.Position(0)
             DiaConfirm.Commander_1 = docEmployee.Commander_1(0)
             DiaConfirm.Commander_2 = docEmployee.Commander_2(0)



             DiaConfirm.EmpDestCompanyID = docEmployee.EmpDestCompanyID(0)
             DiaConfirm.EmpDestCompany = docEmployee.EmpDestCompany(0)
             DiaConfirm.EmpDestGroup = docEmployee.EmpDestGroup(0)       
             DiaConfirm.EmpDestField = docEmployee.EmpDestField(0)          
             DiaConfirm.EmpDestDepartmen = docEmployee.EmpDestDepartmen(0)          
             DiaConfirm.EmpDestSections = docEmployee.EmpDestSections(0)           
             DiaConfirm.EmpDestDivision = docEmployee.EmpDestDivision(0)           
             DiaConfirm.EmpDestBranch = docEmployee.EmpDestBranch(0)           
             DiaConfirm.EmpDestLevels = docEmployee.EmpDestLevels(0)          
             DiaConfirm.EmpDestPosition = docEmployee.EmpDestPosition(0)
             DiaConfirm.EmpDestType = docEmployee.EmpDestType(0)
             DiaConfirm.EmbDestSup_1 = docEmployee.EmbDestSup_1(0)
             DiaConfirm.EmbDestSup_2 = docEmployee.EmbDestSup_2(0)


        End If
        
        
        Set doc = dc.GetNextDocument(doc)	
    Wend
	
	Temp = ws.DialogBox("EmpCheckEmployee", True, True,False,True, False, False, "ขณะนี้ท่านกำลังตรวจสอบข้อมูลพนักงาน", DiaConfirm,True,False)
	If Temp = True Then		
		
		
		'Msgbox "บันทึกเสร็จ",0+64,"บันทึก"
	End If
	
	'Call uidoc.Refresh()
	
	
End Sub