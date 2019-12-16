Sub Click(Source As Button)
	
	Dim session As New NotesSession
	Dim ws As New NotesUIWorkspace
	Dim db As notesdatabase
	Dim doc As notesdocument
	Dim viewEmployeeUpdate As notesview
	Dim uidoc As NotesUIDocument
    Dim uiview As NotesUIView
	Dim curdoc As NotesDocument
	Dim allrows As Integer
	Dim dc As NotesDocumentCollection
	
    'Set uidoc = ws.CurrentDocument
	'Set curdoc = uidoc.Document
	Set session = New NotesSession
	Set db = session.currentdatabase
	Set dc = db.unprocesseddocuments
	
	allrows = dc.count'vc.count
	If allrows = 0 Then
		Msgbox "กรุณาเลือกรายการ",0+64,"คุณไม่ได้ทำการเลือกรายการที่ต้องการ"
		Exit Sub
	End If
	
'	Dim ws As New NotesUIWorkspace
	Dim response As Variant
	Dim values(1) As Variant
	values(0) = "ผู้อนุมัติอันดับ 1"
	values(1) = "ผู้อนุมัติอันดับ 2"
	response = ws.Prompt (PROMPT_OKCANCELLIST, _
	"Please Select Export Excel Choice", _
	"Select Choice to Export.", _
	values(0), values)
	If Isempty (response) Then
		Exit Sub
	Else
		Set pdoc = db.GetProfileDocument("EmployeeProfile")
		
		Set doc = dc.getfirstdocument
		
		'Dim collection As NotesDocumentCollection
		Set collection = ws.PickListCollection( _
		PICKLIST_CUSTOM, _
		False, _
		pdoc.ServerEmployee(0), _
		pdoc.DBEmployee(0), _
		"AllEmpActive", _
		"รายชื่อข้อมูลพนักงาน", _
		"กรุณาเลือกรายการ จากข้อมูลพนักงาน")
		
		If collection.Count = 0 Then
			Exit Sub
		Else
			Dim docEmp As NotesDocument
			Set docEmp = collection.GetFirstDocument
			
			While Not doc Is Nothing
                If response =  "ผู้อนุมัติอันดับ 1" Then
                    If docEmp.NotesCNname(0) ="" Then
                        Msgbox "ชื่อที่ท่านเลือกไม่มี Notes User Common Name",0+64,"เลือกผู้อนุมัติอันดับ 1"
                        doc.Commander_1 = docEmp.NotesCNname(0)
					    doc.CommanderID_1 = docEmp.empid(0)
                    else
                        doc.Commander_1 = docEmp.NotesCNname(0)
					    doc.CommanderID_1 = docEmp.empid(0)
                    End IF
					
				Elseif  response =  "ผู้อนุมัติอันดับ 2" Then
                    If docEmp.NotesCNname(0) ="" Then
                        Msgbox "ชื่อที่ท่านเลือกไม่มี Notes User Common Name",0+64,"เลือกผู้อนุมัติอันดับ 2"
                        doc.Commander_2 = docEmp.NotesCNname(0)
					    doc.CommanderID_2 = docEmp.empid(0)
                    else
                        doc.Commander_2 = docEmp.NotesCNname(0)
					    doc.CommanderID_2 = docEmp.empid(0)
                    End IF
				End If
				Call doc.Save(True,True)
				Set doc = dc.GetNextDocument(doc)
			Wend	
            Msgbox "บันทึกเสร็จ",0+64,"บันทึก"
			Set uiview = ws.CurrentView
            Call uiview.DeselectAll
		End If
	End If
End Sub