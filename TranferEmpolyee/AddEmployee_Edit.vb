Sub Click(Source As Button)
	
	Dim session As New NotesSession
	Dim ws As New NotesUIWorkspace
	Dim db As NotesDatabase
	Dim uidoc As NotesUIDocument
	Dim curdoc As NotesDocument	
	Dim doc As NotesDocument
	Dim docInList As NotesDocument
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	
	Dim viewEmp As NotesView
	Dim RowEmp As NotesDocumentCollection	
	Dim collection As NotesDocumentCollection
	Dim EmpInList As NotesDocumentCollection
	Dim item As NotesItem
	Dim FindValue As String
	Dim ViewName As String 
	
	ViewName = "AllEmpActiveByLevel"
	'validae field rerquestType ต้องเลือก FIeld rerquestType ก่อนถึงทำได้
	
	
	If curdoc.rerquestType(0) = "P" Then
		
		If curdoc.EmpTypeTransfer(0) = "" Then
			Msgbox "กรุณาระบุประเภทการโอนย้าย "+appendMsg,0+48, "เพิ่มการโอนย้ายแบบบุคคล"
			Exit Sub
		End If
		If curdoc.requestGroup(0) = "" Then
			Msgbox "กรุณาระบุเป็น โอนย้าย หรือ มอบหมายงาน "+appendMsg,0+48, "เพิ่มการโอนย้ายแบบบุคคล"
			Exit Sub
		End If
		If curdoc.FindEmpID(0) <> "" Or curdoc.FindEmpFirstName(0) <> "" Or curdoc.FindEmpLastName(0) <> "" Then
			
			Set db = session.CurrentDatabase
			Set viewEmp = db.GetView("GroupEmployeeStep1")
			Set EmpInList = viewEmp.GetAllDocumentsByKey(curdoc.MainDocID(0),True)
			If EmpInList.Count >= 1 Then
				Msgbox "กรณีแบบบุคคลสามารถเลือกได้แค่ 1 คน",0+64,"เพิ่มการโอนย้ายแบบบุคคล"
				Exit Sub
			End If
			
			If curdoc.FindEmpID(0) <> "" Then
				FindValue = curdoc.FindEmpID(0)
				ViewName = "AllEmpActiveByName"
			End If
			If curdoc.FindEmpFirstName(0) <> "" Then
				FindValue = curdoc.FindEmpFirstName(0)
				ViewName = "AllEmpActiveByName"
			End If
			If curdoc.FindEmpLastName(0) <> "" Then
				FindValue = curdoc.FindEmpLastName(0)
				ViewName = "AllEmpActiveByName"
			End If
			If curdoc.FindEmpLastName(0) <> "" And curdoc.FindEmpFirstName(0) <> "" Then
				FindValue = curdoc.FindEmpFirstName(0)+" "+curdoc.FindEmpLastName(0)
				ViewName = "AllEmpActiveByName"
			End If
			If curdoc.FindEmpID(0) <> "" And curdoc.FindEmpLastName(0) <> "" And curdoc.FindEmpFirstName(0) <> "" Then
				FindValue = curdoc.FindEmpID(0)+" "+curdoc.FindEmpFirstName(0)+" "+curdoc.FindEmpLastName(0)
				ViewName = "AllEmpActiveByName"
			End If
			
			Set collection = ws.PickListCollection( _
			PICKLIST_CUSTOM, _
			True, _
			curdoc.ServerEmployee(0), _
			curdoc.DBEmployee(0), _
			ViewName, _
			"รายชื่อข้อมูลพนักงาน", _
			"กรุณาเลือกรายการ จากข้อมูลพนักงาน",FindValue)
			
			If collection.Count = 0 Then
				
				Exit Sub
			Elseif collection.Count > 1 Then
				Msgbox "กรณีแบบบุคคลสามารถเลือกได้แค่ 1 คน",0+64,"เพิ่มการโอนย้ายแบบบุคคล"
				Exit Sub
			End If
			
		Else			
			
			Msgbox "กรุณาใส่ข้อมูลที่ต้องการค้นหา",0+64,"เพิ่มการโอนย้ายแบบบุคคล"	
			Exit Sub
			
		End If	
		Call ClearFieldMainFrom(curdoc, "EmployeeDetail")
		Call curdoc.ComputeWithForm(False,False)
		Call uidoc.Save
		Call curdoc.Save(True,True)
		
		Set doc = collection.GetFirstDocument
		
		Dim DiaConfirm As NotesDocument
		Set DiaConfirm = db.CreateDocument
		
		DiaConfirm.MainDocID = curdoc.MainDocID(0)
		DiaConfirm.WorkerName = curdoc.WorkerName(0)
		DiaConfirm.CreateDateFromMainReq = curdoc.CreateDate(0)
		DiaConfirm.AuthorizeView = "SubStaffDetail1"
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
		DiaConfirm.Location = doc.LocationName(0)
		DiaConfirm.LocationID = doc.LocationCode(0)
		
		DiaConfirm.Position = doc.JobFunction(0)
		DiaConfirm.PositionID = doc.CodeJobFunction(0)
		
		DiaConfirm.CommanderID_1 = doc.ReportIDTo(0)
		DiaConfirm.Commander_1 = Cstr(doc.FirstNameReportToTH(0))+" "+ Cstr(doc.LastNameReportToTH(0))
		
		
		DiaConfirm.CommanderID_2 = doc.ReportIDTo_1(0)
		DiaConfirm.Commander_2 = Cstr(doc.FirstNameReportToTH_1(0)) +" "+ Cstr(doc.LastNameReportToTH_1(0))
		
		DiaConfirm.Replace = ""
		DiaConfirm.BPRecruitment = ""
		DiaConfirm.BPLevels = ""
		DiaConfirm.BPSpecify = ""
		DiaConfirm.BPNewPosition = ""
		
		DiaConfirm.EmbEmployeeStaffCard = ""
		DiaConfirm.EmbEmployeeSocialSecurityCard = ""
		DiaConfirm.EmbSocialSecurityNameSigned = ""
		DiaConfirm.EmbSocialSecurityDate = ""
		DiaConfirm.EmbEmployeeMobile = ""
		DiaConfirm.EmpSimNumber = ""
		DiaConfirm.EmbEmployeeCar = ""
		DiaConfirm.EmbCarType = ""
		DiaConfirm.EmbCarlicense = ""
		DiaConfirm.EmbCarNameSigned = ""
		DiaConfirm.EmbCarDate = ""
		DiaConfirm.EmbEmployeeComputer = ""
		DiaConfirm.EmbEmployeeTablet = ""
		DiaConfirm.EmbEmployeeOther = ""
		DiaConfirm.EmbOtherSpecify = ""
		DiaConfirm.EmbOtherNameSigned = ""
		DiaConfirm.EmbOtherDate = ""
		DiaConfirm.EmbEmployeeAdvances = ""
		DiaConfirm.EmbAdvancesMoney = ""
		DiaConfirm.EmbAttachPayin = ""
		DiaConfirm.EmbConditionAdvances = ""
		DiaConfirm.EmbSalaryDeduct = ""
		
		DiaConfirm.EmpTypeTransfer = curdoc.EmpTypeTransfer(0)
		DiaConfirm.requestGroup = curdoc.requestGroup(0)
		DiaConfirm.EmpDestCompanyID = doc.companyCode(0)+"000000000000000000"
		DiaConfirm.EmpDestCompanyID_1 = doc.companyCode(0)+"000000000000000000"
		DiaConfirm.EmpDestCompany = doc.company(0)
		DiaConfirm.EmpDestGroupID = doc.BusinessGroupCode(0)
		DiaConfirm.EmpDestGroupID_1 = doc.BusinessGroupCode(0)
		DiaConfirm.EmpDestGroup = doc.BusinessGroupName(0)
		DiaConfirm.EmpDestFieldID = doc.BusinessLineCode(0)
		DiaConfirm.EmpDestFieldID_1 = doc.BusinessLineCode(0)
		DiaConfirm.EmpDestField = doc.BusinessLineName(0)
		DiaConfirm.EmpDestDepartmenID = doc.Deptcode(0)
		DiaConfirm.EmpDestDepartmenID_1 = doc.Deptcode(0)
		DiaConfirm.EmpDestDepartmen = doc.EmpDept(0)
		DiaConfirm.EmpDestSectionsID = doc.SessionCode(0)
		DiaConfirm.EmpDestSectionsID_1 = doc.SessionCode(0)
		DiaConfirm.EmpDestSections = doc.SessionName(0)
		DiaConfirm.EmpDestDivisionID = doc.DivisionCode(0)
		DiaConfirm.EmpDestDivisionID_1 = doc.DivisionCode(0)
		DiaConfirm.EmpDestDivision = doc.DivisionName(0)
		DiaConfirm.EmpDestBranchID = doc.LocationCode(0)
		DiaConfirm.EmpDestBranchID_1 = doc.LocationCode(0)
		DiaConfirm.EmpDestBranch = doc.LocationName(0)
		DiaConfirm.EmpDestFrontYard = ""
		
		DiaConfirm.EmpDestPosition = ""
		DiaConfirm.EmpDestPositionID = ""
		DiaConfirm.EmpDestLevelsID = ""
		DiaConfirm.EmpDestLevelsID_1 = ""
		DiaConfirm.EmpDestLevels = ""
		DiaConfirm.EmpDestType = ""
		DiaConfirm.EmpDestStartDate = ""
		DiaConfirm.EmpDestToDate = ""
		DiaConfirm.EmpDestDetail = ""
		
		DiaConfirm.EmbDestSup_1 = ""
		DiaConfirm.EmbDestSupID_1 = ""
		DiaConfirm.EmbDestSup_2 = ""
		DiaConfirm.EmbDestSupID_2 = ""
		
		'DiaConfirm.EmpDestPosition = doc.JobFunction(0)
		'DiaConfirm.EmpDestPositionID = doc.CodeJobFunction(0)
		'DiaConfirm.EmpDestLevelsID = Mid(doc.CodeJobFunction(0),1,2)
		
		'DiaConfirm.EmbDestSup_1 = doc.ReportIDTo(0)
		'DiaConfirm.EmbDestSup_2 = doc.ReportIDTo_1(0)
		
		
		Temp = ws.DialogBox("DiaInputEmbSingleEmployee", True, True,False,False, False, False, "ขณะนี้ท่านกำลังกรอกข้อมูลในส่วน Requester", DiaConfirm,True,False)
		If Temp = True Then				
			DiaConfirm.form = "EmbGroupEmployee"
			DiaConfirm.AuthorizeView = ""
			curdoc.ChkAddEmployee = "Yes"
			DiaConfirm.ChkAddEmployee = curdoc.ChkAddEmployee(0)
			Call ClearFieldMainFrom(curdoc, "EmployeeDetail_Full")
			Call DiaConfirm.ComputeWithForm(True,True)
			Call DiaConfirm.Save(True,True)
			Call uidoc.Save()
			Call uidoc.Refresh()
		End If			
		
	Elseif curdoc.rerquestType(0) = "G" Then	
		
		Dim CompanyIDInList As String 
		Dim GroupIDInList As String 
		Dim FieldIDInList As String 
		Dim DepartmentIDInList As String 
		Dim CompanyInList As String 
		Dim GroupInList As String 
		Dim FieldInList As String 
		Dim DepartmentInList As String 
		
		Set db = session.CurrentDatabase
		Set viewEmp = db.GetView("GroupEmployeeStep1")
		Set EmpInList = viewEmp.GetAllDocumentsByKey(curdoc.MainDocID(0),True)
		
		'=== ดึงข้อมูลจากล่าสุด เพื่อเอาเปรียบเทียบว่าฝ่ายเดียวกันไหม =====
		
		If EmpInList.Count > 0 Then
			Set docInList = EmpInList.GetLastDocument
			CompanyIDInList = docInList.CompanyID(0)
			GroupIDInList = docInList.CompanyGroupID(0)
			FieldIDInList = docInList.CompanyLinesID(0)
			DepartmentIDInList = docInList.DepartmentID(0)
			CompanyInList = docInList.Company(0)
			GroupInList = docInList.CompanyGroup(0)
			FieldInList = docInList.CompanyLines(0)
			DepartmentInList = docInList.Department(0)
			
		End If
		
		
		'==== เช็คค่าตัวเลือกในการค้นหา ========
		
		If curdoc.EmpTypeTransfer(0) = "" Then
			Msgbox "กรุณาระบุประเภทการโอนย้าย "+appendMsg,0+48, "เพิ่มการโอนย้ายแบบกลุ่ม"
			Exit Sub
		End If
		If curdoc.requestGroup(0) = "" Then
			Msgbox "กรุณาระบุเป็น โอนย้าย หรือ มอบหมายงาน "+appendMsg,0+48, "เพิ่มการโอนย้ายแบบกลุ่ม"
			Exit Sub
		End If
		If curdoc.FindCompanyID(0) <> "" Then
			FindValue = curdoc.FindCompanyID(0)
			FindValue = Mid(FindValue,1,3)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindGroupID(0) <> "" Then
			FindValue = curdoc.FindGroupID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindFieldID(0) <> "" Then
			FindValue = curdoc.FindFieldID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindDepartmentID(0) <>"" Then
			FindValue = curdoc.FindDepartmentID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindSectionsID(0) <> "" Then
			FindValue = curdoc.FindSectionsID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindDivisionID(0) <> "" Then
			FindValue = curdoc.FindDivisionID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindBranchID(0) <> "" Then
			FindValue = curdoc.FindBranchID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindLevelsID(0) <> "" Then
			'FindValue = curdoc.FindLevelsID(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindPosition(0) <> "" Then
			FindValue = curdoc.FindPosition(0)
			ViewName = "AllEmpActiveByLevel"
			
		End If
		If curdoc.FindEmpID(0) <> "" Then
			FindValue = curdoc.FindEmpID(0)
			ViewName = "AllEmpActiveByName"
			
		End If
		If curdoc.FindEmpFirstName(0) <> "" Then
			FindValue = curdoc.FindEmpFirstName(0)
			ViewName = "AllEmpActiveByName"
			
		End If
		If curdoc.FindEmpLastName(0) <> "" Then
			FindValue = curdoc.FindEmpLastName(0)
			ViewName = "AllEmpActiveByName"
			
		End If
		If curdoc.FindEmpLastName(0) <> "" And curdoc.FindEmpFirstName(0) <> "" Then
			FindValue = curdoc.FindEmpFirstName(0)+" "+curdoc.FindEmpLastName(0)
			ViewName = "AllEmpActiveByName"
			
		End If
		If curdoc.FindEmpID(0) <> "" And curdoc.FindEmpLastName(0) <> "" And curdoc.FindEmpFirstName(0) <> "" Then
			FindValue = curdoc.FindEmpID(0)+" "+curdoc.FindEmpFirstName(0)+" "+curdoc.FindEmpLastName(0)
			ViewName = "AllEmpActiveByName"
		End If
		
		Set db = session.CurrentDatabase
		Set collection = ws.PickListCollection( _
		PICKLIST_CUSTOM, _
		True, _
		curdoc.ServerEmployee(0), _
		curdoc.DBEmployee(0), _
		ViewName, _
		"รายชื่อข้อมูลพนักงาน", _
		"กรุณาเลือกรายการ จากข้อมูลพนักงาน",FindValue)
		If collection.Count = 0 Then
		Else
			If checkAllEmployeeDup(collection,curdoc.MainDocID(0)) Then
					'Msgbox "มีชื่อพนักงานบางคนถูกเพิ่มเข้ามาแล้ว",0+64,"เพิ่มการโอนย้ายแบบกลุ่ม"
				Exit Sub
			End If
			
			Call curdoc.ComputeWithForm(False,False)
			Call curdoc.Save(True,True)
			
			Set doc = collection.GetFirstDocument
			While Not doc Is Nothing
				
				If EmpInList.Count > 0 Then
					
					curdoc.FindCompany = CompanyInList
					curdoc.FindCompanyID = CompanyIDInList
					curdoc.FindGroup = GroupInList
					curdoc.FindGroupID = GroupIDInList
					curdoc.FindField = FieldInList
					curdoc.FindFieldID = FieldIDInList
					curdoc.FindDepartment = DepartmentInList
					curdoc.FindDepartmentID = DepartmentIDInList
					
					If CompanyIDInList <> doc.companyCode(0) Then
						Msgbox "กรุณาเลือพนักงานบริษัทเดียวกัน",0+64,"เพิ่มการโอนย้ายแบบกลุ่ม"
						Call uidoc.Refresh()
						Exit Sub
					Elseif GroupIDInList <> doc.BusinessGroupCode(0) Then
						Msgbox "กรุณาเลือพนักงานกลุามธุรกิจเดียวกัน",0+64,"เพิ่มการโอนย้ายแบบกลุ่ม"
						Call uidoc.Refresh()
						Exit Sub
					Elseif FieldIDInList <> doc.BusinessLineCode(0) Then
						Msgbox "กรุณาเลือพนักงานสายงานเดียวกัน",0+64,"เพิ่มการโอนย้ายแบบกลุ่ม"
						Call uidoc.Refresh()
						Exit Sub
					Elseif DepartmentIDInList <> doc.Deptcode(0) Then
						Msgbox "กรุณาเลือพนักงานฝ่ายเดียวกัน",0+64,"เพิ่มการโอนย้ายแบบกลุ่ม"
						Call uidoc.Refresh()
						Exit Sub
					End If
					
				End If
				
				
				
				'Dim DiaConfirm As NotesDocument
				Set DiaConfirm = db.CreateDocument
				DiaConfirm.MainDocID = curdoc.MainDocID(0)
				DiaConfirm.CreateDateFromMainReq = curdoc.CreateDate(0)
				DiaConfirm.WorkerName = curdoc.WorkerName(0)
				'DiaConfirm.AuthorizeView = "SubStaffDetail1"
				
				'------- ต้นทาง ---------
				
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
				
				DiaConfirm.Location = doc.LocationName(0)
				DiaConfirm.LocationID = doc.LocationCode(0)
				
				DiaConfirm.Position = doc.JobFunction(0)
				DiaConfirm.PositionID = doc.CodeJobFunction(0)
				
				DiaConfirm.CommanderID_1 = doc.ReportIDTo(0)
				DiaConfirm.Commander_1 = Cstr(doc.FirstNameReportToTH(0))+" "+ Cstr(doc.LastNameReportToTH(0))
				
				DiaConfirm.CommanderID_2 = doc.ReportIDTo_1(0)
				DiaConfirm.Commander_2 = Cstr(doc.FirstNameReportToTH_1(0)) +" "+ Cstr(doc.LastNameReportToTH_1(0))
				
				'------- ปลายทาง ---------
				DiaConfirm.EmpTypeTransfer = curdoc.EmpTypeTransfer(0)
				DiaConfirm.requestGroup = curdoc.requestGroup(0)
				DiaConfirm.EmpDestCompanyID = doc.companyCode(0)+"000000000000000000"
				DiaConfirm.EmpDestCompany = doc.company(0)
				DiaConfirm.EmpDestGroup = doc.BusinessGroupName(0)
				DiaConfirm.EmpDestGroupID = doc.BusinessGroupCode(0)
				DiaConfirm.EmpDestField = doc.BusinessLineName(0)
				DiaConfirm.EmpDestFieldID = doc.BusinessLineCode(0)
				DiaConfirm.EmpDestDepartmen = doc.EmpDept(0)
				DiaConfirm.EmpDestDepartmenID = doc.Deptcode(0)
				DiaConfirm.EmpDestSections = doc.SessionName(0)
				DiaConfirm.EmpDestSectionsID = doc.SessionCode(0)
				DiaConfirm.EmpDestDivision = doc.DivisionName(0)
				DiaConfirm.EmpDestDivisionID = doc.DivisionCode(0)
				
				DiaConfirm.EmpDestBranch = doc.LocationName(0)
				DiaConfirm.EmpDestBranchID = doc.LocationCode(0)
				
				'DiaConfirm.EmpDestPosition = doc.JobFunction(0)
				'DiaConfirm.EmpDestPositionID = doc.CodeJobFunction(0)
				'DiaConfirm.EmpDestLevelsID = Mid(doc.CodeJobFunction(0),1,2)
				
				'DiaConfirm.EmbDestSupID_1 = doc.ReportIDTo(0)
				'DiaConfirm.EmbDestSup_1 = Cstr(doc.FirstNameReportToTH(0))+" "+ Cstr(doc.LastNameReportToTH(0))
				
				
				'DiaConfirm.EmbDestSupID_2 = doc.ReportIDTo_1(0)
				'DiaConfirm.EmbDestSup_2 = Cstr(doc.FirstNameReportToTH_1(0)) +" "+ Cstr(doc.LastNameReportToTH_1(0))
				
				DiaConfirm.form = "EmbGroupEmployee"
				DiaConfirm.AuthorizeView = ""
				DiaConfirm.ChkAddEmployee = curdoc.ChkAddEmployee(0)
				
				Call DiaConfirm.ComputeWithForm(True,True)
				Call DiaConfirm.Save(True,True)
				
				Set doc = collection.GetNextDocument(doc)				
			Wend
			curdoc.ChkAddEmployee = "Yes"
			Call ClearFieldMainFrom(curdoc, "EmployeeDetail")
			Call curdoc.ComputeWithForm(False,False)
			Call uidoc.Save()
			Call curdoc.Save(True,True)
			Call uidoc.Refresh()
		End If
	End If
	
End Sub