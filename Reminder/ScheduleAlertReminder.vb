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
	Dim Reminder As String
	Dim BeforeReminder As String
	Dim AfterReminder As String
	Dim xx As String
	
	Set db = s.Currentdatabase
	Set dbname = New NotesName(db.Server)
	
	Dim tmpDoc As NotesDocument
	Set tmpDoc = db.Createdocument()
	Dim tmpDate As New NotesDateTime(tmpDoc.Created)
	Dim tmpBeforeDay As New NotesDateTime(tmpDoc.Created)
	Dim tmpAfterDay As New NotesDateTime(tmpDoc.Created)
	Call tmpBeforeDay.AdjustDay(1)
	Call tmpAfterDay.AdjustDay(-1)
	Reminder = tmpDate.Dateonly
	BeforeReminder = tmpBeforeDay.Dateonly
	AfterReminder = tmpAfterDay.Dateonly
	
	
	eval = Evaluate({ "/" + @ReplaceSubstring(@subset(@DBName;-1);"\\";"/") + "/0/"})
	DBpath = "notes://" & dbname.Common &  eval(0)
	
	Dim strQuery As String
	
	strQuery = {
	SELECT (
	(Form = "MainReminder") 
	& (StatusMarkComplete !="Y") 
	& (Chk_SendEmail !="")
	)
	}
	
	Set dc = db.Search(strQuery, Nothing,0)
	Set doc = dc.Getfirstdocument()
	
	While Not doc Is Nothing
		xx = doc.Subject(0)
		doc.Chk_SendEmail = Null
		Call doc.Save(True, True)
		Set doc = dc.Getnextdocument(doc)	
	Wend
	
	strQuery = ""
	strQuery = {SELECT (
	(Form = "MainReminder")
	& (StatusMarkComplete !="Y")
	&(@IsMember(@Text(@Today);@Text(NotificationDateTotal))))}
	Set dc =db.Search(strQuery, Nothing,0)
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
		doc.Chk_SendEmail = today
		Call doc.Save(True, True)
		
		Set doc = dc.Getnextdocument(doc)
	Wend
	
	strQuery = ""
	strQuery = {SELECT (
	(Form = "MainReminder")
	&(StatusMarkComplete !="Y")
    &(Chk_SendEmail = "")
	&(@IsMember(@Text(@Today);@Text(NotificationAfterDateTotal))))}
	
	Set dc =db.Search(strQuery, Nothing,0)
	Set doc = dc.Getfirstdocument()
	
	While Not doc Is Nothing 
		xx = doc.Subject(0)
		
		Set docM=New NotesDocument(db)
		Set richStyle = s.CreateRichTextStyle
		Set rtBody = New NotesRichTextItem(docM,"Body")
		
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
		rtBody.AppendText("แจ้งเตือนเกินกำหนด Reminder")
		
		docM.Subject= "Overdue Reminder Mail "+ Cstr(doc.Subject(0)) 
		
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
		
		Set doc = dc.Getnextdocument(doc)
		
	Wend
	
End Sub