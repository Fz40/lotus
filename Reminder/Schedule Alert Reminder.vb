%REM
	Agent Schedule Alert Reminder
	Created Aug 14, 2019 by Sarawut Somton/PTG
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
	Dim doc As NotesDocument
	Dim dc As NotesDocumentCollection
	
	Dim DBpath As String
	Dim eval As Variant
	Dim KeyInReminder (8) As String
	Dim BeforeReminder (8) As String
	Dim AfterReminder (7) As String
	
	Set db = s.Currentdatabase
	Set dbname = New NotesName(db.Server)
	
	Dim tmpDoc As NotesDocument
	Set tmpDoc = db.Createdocument()
	Dim tmpDate As New NotesDateTime(tmpDoc.Created)
	Dim tmpBeforeDay As New NotesDateTime(tmpDoc.Created)
	Dim tmpAfterDay As New NotesDateTime(tmpDoc.Created)
	Call tmpBeforeDay.AdjustDay(1)
	Call tmpAfterDay.AdjustDay(-1)
	
	If ChkHoliday(tmpDate) = True Then
		Exit Sub
	End If
	
	eval = Evaluate({ "/" + @ReplaceSubstring(@subset(@DBName;-1);"\\";"/") + "/0/"})
	DBpath = "notes://" & dbname.Common &  eval(0)
	
	KeyInReminder(0) = "Daily"+"^"+ tmpDate.Dateonly
	KeyInReminder(1) = "Daily"+"^"+ "01/01/1990"
	KeyInReminder(2) = "Weekly"+"^"+ tmpDate.Dateonly
	KeyInReminder(3) = "Weekly"+"^0"+ Cstr(Weekday(tmpDate.Dateonly))+"/01/1991"
	KeyInReminder(4) = "Monthy"+"^"+ tmpDate.Dateonly
	If Day(tmpDate.Dateonly) < 10 Then
		KeyInReminder(5) = "Monthy"+"^0"+ Day(tmpDate.Dateonly)+"/01/1992"
	Else
		KeyInReminder(5) = "Monthy"+"^"+ Day(tmpDate.Dateonly)+"/01/1992"
	End If
	KeyInReminder(6) = "Yearly"+"^"+ tmpDate.Dateonly
	KeyInReminder(7) = "Custom"+"^"+ tmpDate.Dateonly
	KeyInReminder(8) = "OneDay"+"^"+ tmpDate.Dateonly
	
	'====== ก่อนการ reminder 1 วัน  ==========
	
	BeforeReminder(0) = "Daily"+"^"+ tmpBeforeDay.Dateonly
	BeforeReminder(1) = "Daily"+"^"+ "01/01/1990"
	BeforeReminder(2) = "Weekly"+"^"+ tmpBeforeDay.Dateonly
	BeforeReminder(3) = "Weekly"+"^0"+ Cstr(Weekday(tmpBeforeDay.Dateonly))+"/01/1991"
	BeforeReminder(4) = "Monthy"+"^"+ tmpBeforeDay.Dateonly
	If Day(tmpBeforeDay.Dateonly) < 10 Then
		BeforeReminder(5) = "Monthy"+"^0"+ Day(tmpBeforeDay.Dateonly)+"/01/1992"
	Else
		BeforeReminder(5) = "Monthy"+"^"+ Day(tmpBeforeDay.Dateonly)+"/01/1992"
	End If
	BeforeReminder(6) = "Yearly"+"^"+ tmpBeforeDay.Dateonly
	BeforeReminder(7) = "Custom"+"^"+ tmpBeforeDay.Dateonly
	BeforeReminder(8) = "OneDay"+"^"+ tmpBeforeDay.Dateonly
	
	'====== หลัง  reminder 1 วัน  ==========
	
	AfterReminder(0) = "Daily"+"^"+ tmpAfterDay.Dateonly
	AfterReminder(1) = "Weekly"+"^"+ tmpAfterDay.Dateonly
	AfterReminder(2) = "Weekly"+"^0"+ Cstr(Weekday(tmpAfterDay.Dateonly))+"/01/1991"
	AfterReminder(3) = "Monthy"+"^"+ tmpAfterDay.Dateonly
	If Day(tmpAfterDay.Dateonly) < 10 Then
		AfterReminder(4) = "Monthy"+"^0"+ Day(tmpAfterDay.Dateonly)+"/01/1992"
	Else
		AfterReminder(4) = "Monthy"+"^"+ Day(tmpAfterDay.Dateonly)+"/01/1992"
	End If
	AfterReminder(5) = "Yearly"+"^"+ tmpAfterDay.Dateonly
	AfterReminder(6) = "Custom"+"^"+ tmpAfterDay.Dateonly
	AfterReminder(7) = "OneDay"+"^"+ tmpAfterDay.Dateonly
	
	Set viewAlert = db.Getview("KeyReminder")
	
	Forall Keys In AfterReminder
		Set dc = viewAlert.Getalldocumentsbykey(Keys , True)
		Set doc = dc.Getfirstdocument()
		
		While Not doc Is Nothing 
			doc.Alert_Reminder = "N"
			Call doc.Save(True,True)
			Set doc = dc.Getnextdocument(doc)
		Wend
	End Forall
	
	Forall Keys In KeyInReminder
		
		Set dc = viewAlert.Getalldocumentsbykey(Keys, True)
		Set doc = dc.Getfirstdocument()
		
		While Not doc Is Nothing 
			
			Dim docM As NotesDocument
			Dim richStyle As NotesRichTextStyle
			Dim rtBody As NotesRichTextItem
			Dim toname As Variant
			Dim ccname As Variant
			Dim bccname As Variant
			Dim title As String
			Dim Count_L As Integer
			
			Set docM=New NotesDocument(db)
			Set richStyle = s.CreateRichTextStyle
			Set rtBody = New NotesRichTextItem(docM,"Body")
			
			Const Principal = "No-ReplyFromLotusNotesWorkflow"
			Count_L = 1
			docM.Form="Memo"
			docM.SendTo = doc.SendMailTo(0)
			docM.CopyTo = ""
			docM.BlindCopyTo = ""
			docM.Principal = Principal
			docM.From = Principal
			
			rtBody.AppendText("เรียน  คุณ " + Cstr(doc.SendMailTo(0)))
			rtBody.AddNewLine(2)
			rtBody.AddTab( 1 )
			rtBody.AppendText("แจ้งเตือน Reminder")
			
			docM.Subject= "Alert Reminder Mail "+ Cstr(doc.Subject(0)) 
			
			rtBody.AddNewLine(1)
			'rtBody.AppendStyle(richStyle)
			rtBody.AppendText("Subject :"+Cstr(doc.Subject(0)))
			rtBody.AddNewLine(1)
			rtBody.AppendText("Repeat Type : "+Cstr(doc.Repeat(0)))
			rtBody.AddNewLine(1)
			rtBody.AppendText("BU Type : "+Cstr(doc.BUType(0)))
			rtBody.AddNewLine(1)
			rtBody.AppendText("Category : "+Cstr(doc.Category(0)))
			rtBody.AddNewLine(1)
			rtBody.AppendText("Location : "+Cstr(doc.Location(0)))
			
			rtBody.AddNewLine(1)
			rtBody.AppendText("ท่านสามารถเปิดดูเอกสาร โดยคลิ๊กที่ลิงค์ด้านล่างนี้")
			rtBody.AddNewLine(2)
			
			
			richStyle.PassThruHTML = False
			
			Call rtBody.AppendStyle(richStyle)
			
				'Link For Lotus Notes Mail
			Call rtBody.AppendDocLink(doc, db.Title, "Click here For Lotus Notes Mail >> "+" "+"เลขที่เอกสาร"+" "+":"+" "+doc.MainDocID(0))
			rtBody.AddNewLine(2)
			
				'Link For Webbrower เปิด Document จาก NotesClient
			richStyle.PassThruHTML = True
			Call rtBody.AppendStyle(richStyle)
			Call rtBody.AppendText("<a href = '"+DBpath & Cstr(doc.UniversalID) & "?OpenDocument"+"'>" & " Click Here For Gmail "+"</a>")
			richStyle.PassThruHTML = False
			Call rtBody.AppendStyle(richStyle)
			
			Call docM.send(False)
			
			doc.Alert_Reminder = "Y"
			Call doc.Save(True,True)
			Set doc = dc.Getnextdocument(doc)
		Wend
	End Forall
	
	Forall Keys In BeforeReminder
		Set dc = viewAlert.Getalldocumentsbykey(Keys , True)
		Set doc = dc.Getfirstdocument()
		
		While Not doc Is Nothing 
			doc.Alert_Reminder = "B"
			Call doc.Save(True,True)
			Set doc = dc.Getnextdocument(doc)
		Wend
	End Forall
	
End Sub
