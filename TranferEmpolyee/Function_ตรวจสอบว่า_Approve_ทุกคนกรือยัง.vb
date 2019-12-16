Function checkAllCompleteEmbGroupEmployee(curdoc As NotesDocument , BU As String) As Boolean
	checkAllCompleteEmpGroupEmployee = True
	
	Dim s As New NotesSession
	Dim db As NotesDatabase
	Dim view As NotesView
	Dim colEmbGroupEmployee As NotesDocumentCollection
	Dim docEmbGroupEmployee As NotesDocument
	
	Dim keyword As String
	Set db = s.Currentdatabase
	
	Set view = db.Getview("AllEmbGroup")
	Set colEmbGroupEmployee = view.Getalldocumentsbykey(curdoc.MainDocID(0), False)

	
	If colEmbGroupEmployee.Count > 0 Then
			Set docEmbGroupEmployee = colEmbGroupEmployee.Getfirstdocument()
			While Not docEmbGroupEmployee Is Nothing
                If Bu = "CB" Then
                    If docEmbGroupEmployee.EmbCBStampStatus(0) <> "Approve" Then
						checkAllCompleteEmpGroupEmployee = False
						Exit Function
				    End If
                End IF
                
				Set docEmbGroupEmployee = colEmbGroupEmployee.Getnextdocument(docEmbGroupEmployee)
			Wend
	End If
	
	
End Function