Function checkAllEmployeeDup(collection As NotesDocumentCollection , MainDocID As String) As Boolean

    checkAllEmployeeDup = True
	
	Dim s As New NotesSession
	Dim db As NotesDatabase
	Dim viewEmpGroupEmployee As NotesView
	Dim colEmpGroupEmployee As NotesDocumentCollection
	
	Set db = s.Currentdatabase
	Set viewEmpGroupEmployee = db.Getview("AllEmbGroup")
	Set colEmpGroupEmployee = viewEmpGroupEmployee.Getalldocumentsbykey(MainDocID, true)
	
	If colEmpGroupEmployee.Count > 0 Then 'เช็คว่ามี Doc EmpGroupEmployee หรือยัง
		
		Dim docSelectEmployee As NotesDocument
		Dim docMatch As NotesDocument
		Dim keys As String
		Dim msg As String
		msg = ""
		Set viewEmpGroupEmployee = db.Getview("GroupEmployeeByMainDocIDAndEmployeeID")
		Set docSelectEmployee = collection.Getfirstdocument()
		While Not docSelectEmployee Is Nothing
			keys = MainDocID+"^"+docSelectEmployee.empid(0)
			Set docMatch = viewEmpGroupEmployee.Getdocumentbykey(keys, True)
			If Not docMatch Is Nothing Then
				If msg = "" Then
					msg = "มี พนักงาน ชื่อ "+docSelectEmployee.titleTH(0) +" "+ docSelectEmployee.empnameTH(0) _
					+" "+ docSelectEmployee.emplnameTH(0)+" รหัส "+docSelectEmployee.empid(0)+" ได้ทำการเลือกแล้ว"+chr(13)
				Else
					msg = msg + "มี พนักงาน ชื่อ "+docSelectEmployee.titleTH(0) +" "+ docSelectEmployee.empnameTH(0) _
					+" "+ docSelectEmployee.emplnameTH(0)+" รหัส "+docSelectEmployee.empid(0)+" ได้ทำการเลือกแล้ว"+chr(13)
				End If
			End If			
			Set docSelectEmployee = collection.Getnextdocument(docSelectEmployee) 
		Wend
		If msg = "" Then
			checkAllEmployeeDup = False
		Else
			MsgBox msg,0+64,"ผิดพลาด รายการที่เลือก"
			checkAllEmployeeDup = True
		End If
		
		
	Else
		checkAllEmployeeDup = False
	End If




End Function