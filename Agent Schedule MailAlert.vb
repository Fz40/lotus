%REM
	Agent Schedule MailAlert
	Created May 2, 2019 by Natthawat Srisombat/PTG
	Description: Comments for Agent
%END REM
Option Public
Option Declare

Use "CommonScript"

Sub Initialize
	Dim s As New NotesSession
	Dim db As NotesDatabase
	Dim viewAlert As NotesView
	Dim dbname As NotesName
	
	Dim viewEmp As NotesView
	Dim colEmp As NotesDocumentCollection	
	Dim docEmp As NotesDocument
	Dim BU As String
	Dim BUName As String
	Dim Position As String
	Dim BUStatus As String
	Dim keywordfield As String
	Dim comment As Variant
	Dim DBpath As String
	Dim eval As Variant
	
	Set db = s.Currentdatabase
	Set dbname = New NotesName(db.Server)
	
	Dim tmpDoc As NotesDocument
	Set tmpDoc = db.Createdocument()
	Dim tmpDate As New NotesDateTime(tmpDoc.Created)
	
	If ChkHoliday(tmpDate) = True Then
		Exit sub
	End If
	
	eval = Evaluate({ "/" + @ReplaceSubstring(@subset(@DBName;-1);"\\";"/") + "/0/"})
	DBpath = "notes://" & dbname.Common &  eval(0)
	
	Set viewAlert = db.Getview("AlertMailStatusOnProcess")
	Dim colWorker As NotesDocumentCollection
	Dim docWorkerAlert As NotesDocument

	Dim allWorkerAlert As Variant
	allWorkerAlert = getAllWorkerOnProcess()
	
	If IsEmpty(allWorkerAlert) Then
		Exit sub
	End If

	Dim vc As NotesViewEntryCollection
	Dim entry As NotesViewEntry
	
	ForAll workerAlert In allWorkerAlert
		'Set colWorker = viewAlert.Getalldocumentsbykey(workerAlert, True)
		
		Set vc = viewAlert.Getallentriesbykey(workerAlert, False)
		
		If vc.Count > 0 Then
			
			Dim docM As NotesDocument
			Dim richStyle As NotesRichTextStyle
			Dim rtBody As NotesRichTextItem
			Dim object As NotesEmbeddedObject	
			Set docM=New NotesDocument(db)

			
			Set richStyle = s.CreateRichTextStyle
			Set rtBody=New NotesRichTextItem(docM,"Body")
			
			Dim toname As Variant
			Dim ccname As Variant
			Dim bccname As Variant
			Dim title As String
			Dim Count_L As Integer
			
			Const Principal = "No-ReplyFromLotusNotesWorkflow"
			Count_L = 1
			docM.Form="Memo"
			docM.SendTo = workerAlert
			docM.CopyTo = ""
			docM.BlindCopyTo = ""
			docM.Principal = Principal
			docM.From = Principal
			
			rtBody.AppendText("เรียน  คุณ " + workerAlert)
			rtBody.AddNewLine(2)
			rtBody.AddTab( 1 )
			rtBody.AppendText("แจ้งตามงานคงค้าง  Partner Campaign")
			
			
			'Set docWorkerAlert = colWorker.Getfirstdocument()
			Set entry = vc.GetFirstEntry()
			Set docWorkerAlert = entry.Document
			
	
			docM.Subject= "Alert Mail "+db.Title +"(งานที่ต้องพิจารณา/งานที่ต้องดำเนินการ จำนวน :  "+ CStr(vc.Count)+") "+docWorkerAlert.txtMailSubject(0)	
	
			
			While Not docWorkerAlert Is Nothing
				rtBody.AddNewLine(2)
				rtBody.AppendText("( "+CStr(Count_L) +" ) "+docWorkerAlert.FormTitle(0))
				rtBody.AddNewLine(1)
				
				If docWorkerAlert.form(0) = "MainFormRequestPartnerCampaign" Then
					rtBody.AppendText("เลขที่อกสาร : "+docWorkerAlert.Docno(0))
				Else
					rtBody.AppendText("เลขที่อกสาร : "+docWorkerAlert.QDocno(0))
				End If
				
				rtBody.AddNewLine(1)
				
				If docWorkerAlert.form(0) = "MainFormRequestPartnerCampaign" Then
					rtBody.AppendText("สถานะเอกสาร : "+docWorkerAlert.JobStatusName(0))
				Else
					rtBody.AppendText("สถานะเอกสาร : "+docWorkerAlert.JobStatus(0))
					rtBody.AddNewLine(1)
					rtBody.AppendText("เรื่องที่ขอความเห็น : "+docWorkerAlert.ReqDetail(0))
				End If
				
				rtBody.AddNewLine(1)
				rtBody.AppendText("หัวข้อเรื่อง : "+docWorkerAlert.Subject(0))
				rtBody.AddNewLine(1)
				rtBody.AppendText("ประเภท Campaign : "+docWorkerAlert.CampaignType(0))
				rtBody.AddNewLine(1)
				rtBody.AppendText("ชื่อ Campaign :"+docWorkerAlert.CampaignName(0))
				rtBody.AddNewLine(1)
				If docWorkerAlert.CampaignStartTime(0) ="" Then
					rtBody.AppendText("วันที่เริ่มต้น Campaign :"+Cstr(docWorkerAlert.CampaignStartDate(0))+"  เวลา : -")
				Else
					rtBody.AppendText("วันที่เริ่มต้น Campaign :"+CStr(docWorkerAlert.CampaignStartDate(0))+"  เวลา :  "+CStr(docWorkerAlert.CampaignStartTime(0)))
				End If
				rtBody.AddNewLine(1)
				If docWorkerAlert.CampaignEndTime(0) ="" Then
					rtBody.AppendText("วันที่สิ้นสุด Campaign :"+CStr(docWorkerAlert.CampaignEndDate(0))+"  เวลา : -")
				Else
					rtBody.AppendText("วันที่สิ้นสุด Campaign :"+CStr(docWorkerAlert.CampaignEndDate(0))+"  เวลา :  "+CStr(docWorkerAlert.CampaignEndTime(0)))
				End If
				
				rtBody.AddNewLine(1)
				rtBody.AppendText("ท่านสามารถเปิดดูเอกสาร โดยคลิ๊กที่ลิงค์ด้านล่างนี้")
				rtBody.AddNewLine(2)
				
				
				richStyle.PassThruHTML = False
				
				Call rtBody.AppendStyle(richStyle)

				'Link For Lotus Notes Mail
				Call rtBody.AppendDocLink(docWorkerAlert, db.Title, "Click here For Lotus Notes Mail >> "+" "+"เลขที่เอกสาร"+" "+":"+" "+docWorkerAlert.MainDocID(0))
				rtBody.AddNewLine(2)
				
				'Link For Webbrower เปิด Document จาก NotesClient
				richStyle.PassThruHTML = True
				Call rtBody.AppendStyle(richStyle)
				Call rtBody.AppendText("<a href = '"+DBpath & CStr(docWorkerAlert.UniversalID) & "?OpenDocument"+"'>" & " Click Here For Gmail "+"</a>")
				richStyle.PassThruHTML = False
				Call rtBody.AppendStyle(richStyle)
				
				If docWorkerAlert.form(0) = "MainFormRequestPartnerCampaign" Then
					
					
					rtBody.AddNewLine(2)
					rtBody.AppendText("     รายละเอียดความเห็นเพิ่มเติม      ")
					
					If docWorkerAlert.CommentStage_1(0) <> "" Then
						rtBody.AddNewLine(1)
						rtBody.AppendText("จาก : "+docWorkerAlert.FromRequester(0)+"    แผนก : "+docWorkerAlert.Department(0))
						rtBody.AddNewLine(1)
						rtBody.AppendText("ความเห็น : "+docWorkerAlert.CommentStage_1(0))
					End If
					
					Set viewEmp = db.GetView("EmbReviewerWorkerByMainDocID")
					Set colEmp = viewEmp.GetAllDocumentsByKey(docWorkerAlert.MainDocID(0),True)
					If colEmp.Count > 0 Then
						Set docEmp = colEmp.GetFirstDocument	
						While Not docEmp Is Nothing
							BU = docEmp.BUWorkerType(0)
							Position = docEmp.BUWorkerPosition(0)
							BUStatus = docEmp.JobStatusName(0)
							BUName = docEmp.BUWorkerName(0)
							If docEmp.AuthorizeView(0) = "Support" Then
								keywordfield = "Supporter"
							ElseIf docEmp.AuthorizeView(0) = "Accessory" Then
								keywordfield = "Accessory"
							ElseIf docEmp.AuthorizeView(0) = "HeadSupport" Then
								keywordfield = "HeadSupporter"
							ElseIf docEmp.AuthorizeView(0) = "HeadAccessory" Then
								keywordfield = "HeadAccessory"
							Else
								Exit Sub
							End If
							comment = docEmp.GetItemValue(keywordfield+"_CommentStage")
							
							If CStr(comment(0)) <> "" Then
								rtBody.AddNewLine(1)
								rtBody.AppendText("จาก  : "+BUName+"  ประเภท  : "+BU+"    ระดับ : "+Position+"   สถานะ  : "+BUStatus)
								rtBody.AddNewLine(1)
								rtBody.AppendText("ความเห็นเพิ่มเติม : "+CStr(comment(0)))
							End If

							Set docEmp = colEmp.GetNextDocument(docEmp)		
						Wend
					End If
					
					Set viewEmp = db.GetView("EmbReviewerHeaderByMainDocID")
					Set colEmp = viewEmp.GetAllDocumentsByKey(docWorkerAlert.MainDocID(0),True)
					If colEmp.Count > 0 Then
						Set docEmp = colEmp.GetFirstDocument	
						While Not docEmp Is Nothing
							BU = docEmp.BUWorkerType(0)
							Position = docEmp.BUWorkerPosition(0)
							BUStatus = docEmp.JobStatusName(0)
							BUName = docEmp.BUWorkerName(0)
							If docEmp.AuthorizeView(0) = "Support" Then
								keywordfield = "Supporter"
							ElseIf docEmp.AuthorizeView(0) = "Accessory" Then
								keywordfield = "Accessory"
							ElseIf docEmp.AuthorizeView(0) = "HeadSupport" Then
								keywordfield = "HeadSupporter"
							ElseIf docEmp.AuthorizeView(0) = "HeadAccessory" Then
								keywordfield = "HeadAccessory"
							Else
								Exit Sub
							End If
							comment = docEmp.GetItemValue(keywordfield+"_CommentStage")
							
							If CStr(comment(0)) <> "" Then
								rtBody.AddNewLine(1)
								'rtBody.AppendStyle(richStyle)
								rtBody.AppendText("จาก  : "+BUName+"  ประเภท  : "+BU+"    ระดับ : "+Position+"   สถานะ  : "+BUStatus)
								rtBody.AddNewLine(1)
								rtBody.AppendText("ความเห็นเพิ่มเติม : "+CStr(comment(0)))
							End If

							Set docEmp = colEmp.GetNextDocument(docEmp)		
						Wend
					End If
					
					'=============  CEO =================
					If docWorkerAlert.CommentStage_2(0) <> "" Then
						rtBody.AddNewLine(1)
						'rtBody.AppendStyle(richStyle)
						rtBody.AppendText("จาก  ตัวแทน CEO : "+"   สถานะ  : "+CStr(docWorkerAlert.StatusStage_2(0)))
						rtBody.AddNewLine(1)
						rtBody.AppendText("ความเห็นเพิ่มเติม : "+CStr(docWorkerAlert.CommentStage_2(0)))
					End If
					
					'============= BU / MKT =================
					If docWorkerAlert.CommentStage_3(0) <> "" Then
						rtBody.AddNewLine(1)
						'rtBody.AppendStyle(richStyle)
						rtBody.AppendText("จาก  BU / MKT : "+"   สถานะ  : "+CStr(docWorkerAlert.StatusStage_3(0)))
						rtBody.AddNewLine(1)
						rtBody.AppendText("ความเห็นเพิ่มเติม : "+CStr(docWorkerAlert.CommentStage_3(0)))
					End If
					
					'=============  IT Tester  =================
					If docWorkerAlert.CommentStage_4(0) <> "" Then
						rtBody.AddNewLine(1)
						'rtBody.AppendStyle(richStyle)
						rtBody.AppendText("จาก  IT Tester : "+"   สถานะ  : "+CStr(docWorkerAlert.StatusStage_4(0)))
						rtBody.AddNewLine(1)
						rtBody.AppendText("ความเห็นเพิ่มเติม : "+CStr(docWorkerAlert.CommentStage_4(0)))
					End If
					
					'============= Academy =================
					If docWorkerAlert.CommentStage_5(0) <> "" Then
						rtBody.AddNewLine(1)
						'rtBody.AppendStyle(richStyle)
						rtBody.AppendText("จาก  Academy : "+"   สถานะ  : "+CStr(docWorkerAlert.StatusStage_5(0)))
						rtBody.AddNewLine(1)
						rtBody.AppendText("ความเห็นเพิ่มเติม : "+CStr(docWorkerAlert.CommentStage_5(0)))
					End If
					
					'=============  BU / MKT ตรวจรับงาน  =================
					If docWorkerAlert.CommentStage_7(0) <> "" Then
						rtBody.AddNewLine(1)
						'rtBody.AppendStyle(richStyle)
						rtBody.AppendText("จาก  BU / MKT ตรวจรับงาน : "+"   สถานะ  : "+CStr(docWorkerAlert.StatusStage_7(0)))
						rtBody.AddNewLine(1)
						rtBody.AppendText("ความเห็นเพิ่มเติม : "+ CStr(docWorkerAlert.CommentStage_7(0)))
					End If
				Else

				End If
				
				Count_L = Count_L +1
				'Set docWorkerAlert = colWorker.Getnextdocument(docWorkerAlert)
				Set entry  = vc.GetNextEntry(entry)
				If Not entry Is Nothing Then
					Set docWorkerAlert = entry.Document
				Else
					GoTo NextWorker
				End If	
			Wend
NextWorker:
			'senmail
			Call docM.send(False)
		End If
	End ForAll

	
	
	
End Sub
Function getAllWorkerOnProcess() As Variant
	Dim s As New NotesSession
	Dim db As NotesDatabase
	Dim viewAlert As NotesView
	Dim doc As NotesDocument
	
	Set db = s.Currentdatabase
	Set viewAlert = db.Getview("AlertMailStatusOnProcess")
	Set doc = viewAlert.Getfirstdocument()
	
	Dim allWorker As Variant
	
	If doc Is Nothing Then
		Exit Function
	End If
	
	While Not doc Is Nothing 
		If IsEmpty(allWorker) Then
			allWorker = doc.WorkerName
		Else
			allWorker = FullTrim(ArrayUnique(ArrayAppend(allWorker,doc.WorkerName)))
		End If
		Set doc = viewAlert.Getnextdocument(doc)
	Wend
	getAllWorkerOnProcess = allWorker
End Function