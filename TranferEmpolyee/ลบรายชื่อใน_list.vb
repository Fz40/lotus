Sub Click(Source As Button)
	Dim session As New NotesSession
	Dim ws As New NotesUIWorkspace
	Dim uidoc As NotesUIDocument
	Dim curdoc As NotesDocument	
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	Dim db As NotesDatabase
	Dim doc As NotesDocument
	Dim viewEmp As NotesView
	Dim RowEmp As NotesDocumentCollection	
	Dim item As NotesItem
	Dim FindValue As String
	
	'validae field rerquestType ต้องเลือก FIeld rerquestType ก่อนถึงทำได้
	
	If curdoc.rerquestType(0) = "P" Then
		Set db = session.CurrentDatabase
		Set viewEmp = db.GetView("GroupEmployeeStep1")
		Set RowEmp = viewEmp.GetAllDocumentsByKey(curdoc.MainDocID(0),True)
		Set doc = RowEmp.GetFirstDocument
		While Not doc Is Nothing
			doc.CancelDoc="Y"
			Call doc.ComputeWithForm(True,True)
			Call doc.Save(True,True)
			Set doc = RowEmp.GetNextDocument(doc)
		Wend
		Call uidoc.Refresh()
	Elseif curdoc.rerquestType(0) = "G" Then	
		Set db = session.CurrentDatabase
		Set RowEmp = ws.PickListCollection( _
		PICKLIST_CUSTOM, _
		True, _
		db.Server, _
		db.FileName, _
		"plGroupEmployeeByMainDocID", _
		"เลือก รายชื่อพนักงาน ที่ต้องการลบ", _
		"สามารถเลือก รายชื่อพนักงาน ที่ต้องการลบ ได้มากกว่า 1 ",_
		curdoc.MainDocID(0) )
		
		If RowEmp.Count = 0 Then
			Messagebox "User canceled" ,, _
			"Subject item on the document(s)" 
		Else
			Set doc = RowEmp.GetFirstDocument
			
			Dim viewEmb As NotesView
			Dim colEmb As NotesDocumentCollection
			Dim docEmb As NotesDocument
			Dim keyword As String
			
			Set viewEmb= db.GetView("GroupEmployeeByMainDocIDAndEmployeeID")
			
			
			While Not doc Is Nothing
				keyword = curdoc.MaindocID(0)+"^"+doc.EmpCode(0)
				Set colEmb = viewEmb.GetAllDocumentsByKey(keyword,True)
				If colEmb.Count > 0 Then
					Call colEmb.Stampall("CancelDoc", "Y")
				End If			
				Set doc = RowEmp.GetNextDocument(doc)
			Wend

		End If

        Set viewEmp = db.GetView("GroupEmployeeStep1")
        Set RowEmp = viewEmp.GetAllDocumentsByKey(curdoc.MainDocID(0),True)
        If RowEmp.Count < 1 Then
            curdoc.ChkAddEmployee = "No"
        End If

        Call uidoc.Refresh()
		
	End If
	
End Sub