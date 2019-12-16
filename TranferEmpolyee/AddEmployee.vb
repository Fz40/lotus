Sub Click(Source As Button)
	Dim session As New NotesSession
	Dim ws As New NotesUIWorkspace
	Dim uidoc As NotesUIDocument
	Dim curdoc As NotesDocument	
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	Dim db As NotesDatabase
	Dim collection As NotesDocumentCollection
	Dim doc As NotesDocument
	Dim item As NotesItem
	Dim FindValue As String
	
	'validae field rerquestType ต้องเลือก FIeld rerquestType ก่อนถึงทำได้
	
	If curdoc.rerquestType(0) = "P" Then
		
		
		Set db = session.CurrentDatabase
		Set collection = ws.PickListCollection( _
		PICKLIST_CUSTOM, _
		False, _
		curdoc.ServerEmployee(0), _
		curdoc.DBEmployee(0), _
		"AllEmpActive", _
		"รายชื่อข้อมูลพนักงาน", _
		"กรุณาเลือกรายการ จากข้อมูลพนักงาน")
		
		If collection.Count = 0 Then
		Else
			Set doc = collection.GetFirstDocument
			
			Dim DiaConfirm As NotesDocument
			Set DiaConfirm = db.CreateDocument
			
			DiaConfirm.MainDocID = curdoc.MainDocID(0)
			DiaConfirm.AuthorizeView = "SubStaffDetail"
			DiaConfirm.EmpName = doc.titleTH(0) +" "+ doc.empnameTH(0) +" "+ doc.emplnameTH(0)
			DiaConfirm.EmpCode = doc.empid(0)
			DiaConfirm.rerquestType = curdoc.rerquestType(0)
			DiaConfirm.Company = doc.company(0)
            DiaConfirm.CompanyID = doc.companyCode(0)
			DiaConfirm.CompanyGroup = doc.BusinessGroupName(0)
			DiaConfirm.CompanyGroupID = doc.BusinessGroupCode(0)
			DiaConfirm.CompanyLines = doc.BusinessLineName(0)
			DiaConfirm.CompanyLinesID = doc.BusinessLineCode(0)
			DiaConfirm.Department = doc.EmpDept(0)
            DiaConfirm.DepartmentID = doc.Deptcode(0)
            DiaConfirm.Sections = doc.SessionName(0)
            DiaConfirm.SectionsID = doc.SessionCode(0)
            DiaConfirm.Division = doc.DivisionName(0)
            DiaConfirm.DivisionID = doc.DivisionCode(0)
          

            DiaConfirm.Position = doc.JobFunction(0)
            DiaConfirm.PositionID = doc.CodeJobFunction(0)

			DiaConfirm.Levels = Mid(doc.CodeJobFunction(0),1,3)
            DiaConfirm.LevelsID = doc.CodeJobFunction(0)

            DiaConfirm.CommanderID_1 = doc.ReportIDTo(0)
            DiaConfirm.Commander_1 = CStr(doc.FirstNameReportToTH(0))+" "+ CStr(doc.LastNameReportToTH(0))


            DiaConfirm.CommanderID_2 = doc.ReportIDTo_1(0)
            DiaConfirm.Commander_2 = CStr(doc.FirstNameReportToTH_1(0)) +" "+ CStr(doc.LastNameReportToTH_1(0))

            'DiaConfirm. = doc.LocationName(0)
            'DiaConfirm. = doc.LocationCode(0)

			Temp = ws.DialogBox("DiaInputEmbGroupEmployee", True, True,False,True, False, False, "ขณะนี้ท่านกำลังกรอกข้อมูลในส่วน Requester", DiaConfirm,True,False)
			If Temp = True Then				
				DiaConfirm.form = "EmbGroupEmployee"
				DiaConfirm.AuthorizeView = ""
				Call DiaConfirm.ComputeWithForm(True,True)
				Call DiaConfirm.Save(True,True)
				
				
			End If			
		End If
	Elseif curdoc.rerquestType(0) = "G" Then	
		
		IF curdoc.FindCompanyID(0) <> "" Then
			FindValue = curdoc.FindCompanyID(0)
			FindValue = Mid(FindValue,1,3)
		End IF
		IF curdoc.FindGroupID(0) <> "" Then
			FindValue = curdoc.FindGroupID(0)
		End IF
		IF curdoc.FindFieldID(0) <> "" Then
			FindValue = curdoc.FindFieldID(0)
		End IF
		IF curdoc.FindDepartmentID(0) <>"" Then
			FindValue = curdoc.FindDepartmentID(0)
		End IF
		IF curdoc.FindSectionsID(0) <> "" Then
			FindValue = curdoc.FindSectionsID(0)
		End IF
		IF curdoc.FindDivisionID(0) <> "" Then
			FindValue = curdoc.FindDivisionID(0)
		End IF
		IF curdoc.FindBranchID(0) <> "" Then
			FindValue = curdoc.FindBranchID(0)
		End IF
		IF curdoc.FindLevelsID(0) <> "" Then
			FindValue = curdoc.FindLevelsID(0)
		End IF
		IF curdoc.FindPositionID(0) <> "" Then
			FindValue = curdoc.FindPositionID(0)
		End IF

		Set db = session.CurrentDatabase
		Set collection = ws.PickListCollection( _
		PICKLIST_CUSTOM, _
		True, _
		curdoc.ServerEmployee(0), _
		curdoc.DBEmployee(0), _
		"AllEmpActiveByLevel", _
		"รายชื่อข้อมูลพนักงาน", _
		"กรุณาเลือกรายการ จากข้อมูลพนักงาน",FindValue)
		If collection.Count = 0 Then
		Else
			Set doc = collection.GetFirstDocument
			While Not doc Is Nothing
				'Dim DiaConfirm As NotesDocument
				Set DiaConfirm = db.CreateDocument
				DiaConfirm.MainDocID = curdoc.MainDocID(0)
				DiaConfirm.AuthorizeView = "SubStaffDetail"
				DiaConfirm.EmpName = doc.titleTH(0) +" "+ doc.empnameTH(0) +" "+ doc.emplnameTH(0)
				DiaConfirm.EmpCode = doc.empid(0)
				DiaConfirm.rerquestType = curdoc.rerquestType(0)
				DiaConfirm.Company = doc.company(0)
				DiaConfirm.CompanyID = doc.companyCode(0)
				DiaConfirm.CompanyGroup = doc.BusinessGroupName(0)
				DiaConfirm.CompanyGroupID = doc.BusinessGroupCode(0)
				DiaConfirm.CompanyLines = doc.BusinessLineName(0)
				DiaConfirm.CompanyLinesID = doc.BusinessLineCode(0)
				DiaConfirm.Department = doc.EmpDept(0)
				DiaConfirm.DepartmentID = doc.Deptcode(0)
				DiaConfirm.Sections = doc.SessionName(0)
				DiaConfirm.SectionsID = doc.SessionCode(0)
				DiaConfirm.Division = doc.DivisionName(0)
				DiaConfirm.DivisionID = doc.DivisionCode(0)
			

				DiaConfirm.Position = doc.JobFunction(0)
				DiaConfirm.PositionID = doc.CodeJobFunction(0)

				DiaConfirm.Levels = Mid(doc.CodeJobFunction(0),1,3)
            	DiaConfirm.LevelsID = doc.CodeJobFunction(0)

				DiaConfirm.CommanderID_1 = doc.ReportIDTo(0)
				DiaConfirm.Commander_1 = CStr(doc.FirstNameReportToTH(0))+" "+ CStr(doc.LastNameReportToTH(0))


				DiaConfirm.CommanderID_2 = doc.ReportIDTo_1(0)
				DiaConfirm.Commander_2 = CStr(doc.FirstNameReportToTH_1(0)) +" "+ CStr(doc.LastNameReportToTH_1(0))

				'DiaConfirm. = doc.LocationName(0)
				'DiaConfirm. = doc.LocationCode(0)

				Temp = ws.DialogBox("DiaInputEmbGroupEmployee", True, True,False,True, False, False, "ขณะนี้ท่านกำลังกรอกข้อมูลในส่วน Requester", DiaConfirm,True,False)
				If Temp = True Then				
					DiaConfirm.form = "EmbGroupEmployee"
					DiaConfirm.AuthorizeView = ""
					Call DiaConfirm.ComputeWithForm(True,True)
					Call DiaConfirm.Save(True,True)
					
					
				End If


				Set doc = collection.GetNextDocument(doc)				
			Wend
		End IF
	End If
	
End Sub