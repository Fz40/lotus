Sub Click(Source As Button)
	
	Dim session As New NotesSession
	Dim ws As New NotesUIWorkspace
	Dim db As notesdatabase
	Dim doc As notesdocument
	Dim viewEmployeeUpdate As NotesView
	Dim uidoc As NotesUIDocument
	Dim uiview As NotesUIView
	Dim curdoc As NotesDocument
	
	Dim docEmployee As NotesDocument
	
	
	Set uiview = ws.CurrentView
	Set uidoc = ws.CurrentDocument 
	Set doc = uidoc.Document 
	If Not doc Is Nothing Then

		Set db = session.CurrentDatabase
		Set viewEmployeeUpdate = db.GetView("EmpEmployeeEditDetail")
		Set docEmployee =  viewEmployeeUpdate.GetDocumentByKey(doc.MainDocID(0)+"^"+doc.FindEmployeeID(0),True)
		If Not docEmployee Is Nothing Then
			
			doc.EmpName = docEmployee.EmpName(0)
			doc.EmpCode = docEmployee.EmpCode(0)
			doc.company = docEmployee.Company(0)
			doc.CompanyGroup = docEmployee.CompanyGroup(0)
			doc.CompanyLines = docEmployee.CompanyLines(0)
			doc.Department = docEmployee.Department(0)
			doc.Sections = docEmployee.Sections(0)
			doc.Division = docEmployee.Division(0)
			doc.Location = docEmployee.Location(0)
			doc.Position = docEmployee.Position(0)
			doc.Commander_1 = docEmployee.Commander_1(0)
			doc.Commander_2 = docEmployee.Commander_2(0)
			
			
			
			doc.EmpDestCompanyID = docEmployee.EmpDestCompanyID(0)
			doc.EmpDestCompany = docEmployee.EmpDestCompany(0)
			doc.EmpDestGroup = docEmployee.EmpDestGroup(0)       
			doc.EmpDestField = docEmployee.EmpDestField(0)          
			doc.EmpDestDepartmen = docEmployee.EmpDestDepartmen(0)          
			doc.EmpDestSections = docEmployee.EmpDestSections(0)           
			doc.EmpDestDivision = docEmployee.EmpDestDivision(0)           
			doc.EmpDestBranch = docEmployee.EmpDestBranch(0)           
			doc.EmpDestLevels = docEmployee.EmpDestLevels(0)          
			doc.EmpDestPosition = docEmployee.EmpDestPosition(0)
			doc.EmpDestType = docEmployee.EmpDestType(0)
			doc.EmbDestSup_1 = docEmployee.EmbDestSup_1(0)
			doc.EmbDestSup_2 = docEmployee.EmbDestSup_2(0)


			doc.EmbCBDeductedPayroll = docEmployee.EmbCBDeductedPayroll(0)
    		doc.EmbCBChoice = docEmployee.EmbCBChoice(0)
    		doc.EmbCBPayin = docEmployee.EmbCBPayin(0)
    		doc.EmbCBDeducted = docEmployee.EmbCBDeducted(0)
    		doc.EmbNewSaraly = docEmployee.EmbNewSaraly(0)
    		doc.EmbOtherSaraly = docEmployee.EmbOtherSaraly(0)
    		doc.EmbSaralyRemark = docEmployee.EmbSaralyRemark(0)
			
		End If
    End If  
	
End Sub