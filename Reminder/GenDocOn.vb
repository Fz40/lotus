Sub Initialize
	Dim s As New NotesSession
	Dim agent As NotesAgent
	Set agent = s.CurrentAgent
	Dim db As NotesDatabase
	Dim doc As NotesDocument
	Dim docMainformRequestReminder As NotesDocument
	'Dim currentRouteID As String
	'Dim newRouteID As Integer


	Set db = s.Currentdatabase
	Set docMainformRequestReminder = db.GetDocumentByID(agent.ParameterDocID)'ดึงรายละเอียดจาก Form MainformRouteDetermination
	'Set docMainformRequestPartnerCampaign = db.GetDocumentByID("235A")'ดึงรายละเอียดจาก Form MainformRouteDetermination
	
	Dim viewDocRunningNo As NotesView
	Dim docDocRunningNo As NotesDocument
	Dim key As String
	
	key = docMainformRequestReminder.FlowNameDocRuningNo(0)+"^"+CStr(Year(today))
	
	Set viewDocRunningNo = db.Getview("DocRunningNo")
	Set docDocRunningNo = viewDocRunningNo.Getdocumentbykey(key,true)
	
	
	Dim mon As String
	
	If Month(Today) > 9 Then
		mon = CStr(Month(Today))
	Else
		mon = "0"+CStr(Month(Today))
	End If
	
	Dim days As String
	
	If Day(Today) > 9 Then
		days = CStr(Day(Today))
	Else
		days = "0"+CStr(Day(Today))
	End If
	
	Dim dateTime As New NotesDateTime(GetServerDateTime )
	
	If docDocRunningNo Is Nothing Then
		Set doc = db.CreateDocument 			
		doc.form = "MasterDataRunningNumber"
		doc.FlowNameDocRuningNo = docMainformRequestReminder.FlowNameDocRuningNo(0)
		doc.Year = Year(Today)
		doc.RuningNumber = 1
		Call doc.Computewithform(true, true)
		Call doc.save(True,true)
		
		docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+"000001"

		'Call CalculateSLA(docMainformRequestPartnerCampaign,docMainformRequestPartnerCampaign.step(0))
		
		Call docMainformRequestReminder.save(True,true)
	'	Call setDocNoEmbMaterialDetail(docMainformRequestPartnerCampaign,docMainformRequestPartnerCampaign.DocNo(0))
	Else
		Dim runningNo As String
		runningNo = CStr(docDocRunningNo.RuningNumber(0) + 1)
		If len(runningNo) = 1 Then
			docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+"00000"+runningNo
		ElseIf Len(runningNo) = 2 Then
			docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+"0000"+runningNo
		Elseif Len(runningNo) = 3 Then
			docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+"000"+runningNo
		elseIf Len(runningNo) = 4 Then
			docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+"00"+runningNo
		ElseIf Len(runningNo) = 5 Then
			docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+"0"+runningNo
		Else
			docMainformRequestReminder.DocNo = CStr(Year(Today))+mon+days+"-"+runningNo
		End If 

		Call docMainformRequestReminder.save(True,True)
	'	Call setDocNoEmbMaterialDetail(docMainformRequestPartnerCampaign,docMainformRequestPartnerCampaign.DocNo(0))
	'	Call checkMaterialDetailActiveByField(docMainformRequestPartnerCampaign,db)
		docDocRunningNo.RuningNumber = docDocRunningNo.RuningNumber(0) + 1
		Call docDocRunningNo.save(True,true)
	End If
	
End Sub