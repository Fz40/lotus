Sub Click(Source As Button)
	Dim strFileName As String
	strFileName = "Export Excel Reminder "  + Format(Now, "ddmmyyyy") 
	Dim strDIRName As String
	strDIRName = "Reminder"
	Dim dirname As String
	Dim filename As String
	dirname = strDIRName
	filename = strFileName
	Dim ViewName As String
	ViewName = "SelectExportToExcel"
	Dim ExcelHeader As String
	ExcelHeader = "ข้อมูลในระบบ Reminder ( " + Cstr(Now) + ")"
	Dim sess As New notessession
	Dim db As notesdatabase
	Dim doc As notesdocument
	Dim vw As notesview 
	Dim vc As NotesViewEntryCollection
	Dim dc As NotesDocumentCollection
	Dim rows As Integer
	Dim cols As Integer
	Dim strFormula As String
	Dim vVal As Variant
	Dim strName As String
	Dim strUser As String
	Dim xlApp As Variant
	Dim xlsheet As Variant
	Dim allrows As Integer
	Dim cols_ As Integer
	Dim viewEmp As NotesView
	Dim colEmp As NotesDocumentCollection	
	Dim docEmp As NotesDocument
	Dim BU As String
	Dim BUName As String
	Dim Position As String
	Dim BUStatus As String
	Dim keywordfield As String
	Dim comment As Variant
	Dim IndexColor As Integer
	Dim FinalColBU As Integer
	Dim colEmp_Loop As Integer
	Dim colEmp_Header As Integer
	
	Set sess = New NotesSession
	Set db=sess.currentdatabase
	
	Set vw = db.GetView(ViewName) 
	'Set vc = vw.AllEntries
	
	Set dc=db.unprocesseddocuments
	allrows = dc.count'vc.count
	If allrows = 0 Then
		Msgbox "Please select the documents for export",0+64,"Select the Documents Required for Export"
		Exit Sub
	Elseif allrows = 1 Then
		Dim boxType As Long, answer As Integer
		boxType = 4 + 32 + 0 'YesNo + IconQuestionMark + Default=Yes button
		answer% = Messagebox("Only one document has been selected. Do you want to continue?", boxType, "Export Documents - Continue?")
		If answer = 6 Then 'yes
'continue
		Else 'answer = 7 - no
			Msgbox "Please select the documents for export ",0+64,"Select the Documents Required for Export"
			Exit Sub
		End If
	End If
	
	
	'Set doc=vw.getfirstdocument
	Set doc=dc.getfirstdocument
	Print "Total documents selected for export = " + Cstr(allrows) 
	
	Set xlapp = CreateObject( "Excel.Application")
          ' do not display the Excel window 
	xlapp.Visible = False
	xlapp.Workbooks.add 
	Set xlSheet = xlapp.workbooks(1).worksheets(1)
	
	
	
    'add column titles to spreadsheet first of all 
	
	cols=1 
	rows=2
	Print "Adding Column titles to spreadsheet" 
	
	
	xlsheet.Rows(rows).RowHeight = 21
	xlsheet.Rows(rows).Interior.colorindex = 49
	xlsheet.Rows(rows).HorizontalAlignment = -4108
	xlsheet.Rows(rows).VerticalAlignment = -4108
	
	Forall cTitles In vw.Columns 
		
			xlsheet.Cells(rows,cols ).Value= cTitles.Title
			xlsheet.Cells(rows,cols ).font.bold = True
			xlsheet.Cells(rows,cols ).font.Size = 12
			xlsheet.Cells(rows,cols ).Font.ColorIndex = 19
			cols=cols+1

	End Forall
	
	'Add Header Excel
	cols_ = cols /2
	rows=1	
	
	xlsheet.Cells(rows,cols_ ).font.bold = True
	xlsheet.Cells(rows,cols_ ).font.Size = 14
	xlsheet.Cells(rows,cols_ ).Font.ColorIndex = 19
	'xlsheet.Cells(rows,cols_ ).Interior.colorindex = 26
	xlsheet.Cells(rows,cols_ ).Value= ExcelHeader
	xlsheet.Rows(rows).RowHeight = 25
	xlsheet.Rows(rows).Interior.colorindex = 14
	rows=3
	
	Print "Now adding column data" 
	
	While Not (doc Is Nothing)
		If doc.hasitem("$Ref") Then
			Goto nextdocument
		End If
		Print "Processing " + Cstr(rows-1) + " of "+ Cstr(allrows) + " documents" 
               'Add data to cells of the first worksheet in the new workbook
		
		Dim count As Integer
		
		count = 1
		Dim custname(0) As String
		Dim custaddress(0) As String
		cols=1
		Forall i In vw.columns
            strFormula=i.formula
            If strFormula="" Then
                strName=i.itemname
                If strName = "WeeklyType" Then
                    Dim valWeeklyType As Variant
                    valWeeklyType = doc.getitemvalue(strName) 
                    valWeeklyType = Implode(valWeeklyType)
					valWeeklyType = Replace(valWeeklyType,","," ")
					valWeeklyType = Replace(Replace(Replace(Replace(Replace(Replace(Replace(valWeeklyType,"1","อาทิตย์"),"2","จันทร์"),"3","อังคาร"),"4","พุธ"),"5","พฤหัสบดี"),"6","ศุกร์"),"7","เสาร์")
                    vVal(0) = valWeeklyType
                Elseif strName = "MonthlyType" Then
					Dim valMonthlyType As Variant
                    valMonthlyType = doc.getitemvalue(strName) 
                    valMonthlyType = Implode(valMonthlyType)
					valMonthlyType = Replace(valMonthlyType,","," ")
					vVal(0) = valMonthlyType
				Else
                    vVal = doc.getitemvalue(strName) 
                End If
            Else
                vVal=Evaluate(strFormula,doc)
            End If
            xlsheet.Cells(rows,cols ).Value=vVal(0)
            cols=cols+1
		End Forall
		rows=rows+1
		'End If
nextdocument:
		Set doc=dc.getnextdocument(doc)
		'Set doc=vw.getnextdocument(doc)
	Wend
	
	Print "Processing complete - closing spreadsheet" 
	
	
	
	xlapp.Cells.Select
	xlapp.Selection.columns.AutoFit
	xlapp.Rows("1:1").select
	xlapp.selection.font.bold=True
	
	With xlApp.Worksheets(1)
		.PageSetup.PrintTitleRows="$1:$1"
		.PageSetup.centerheader=ExcelHeader
		.PageSetup.RightFooter="Page &P" & Chr$(13) & "Date: &D"
		.PageSetup.CenterFooter=""
		.PageSetup.Orientation = 2
		.PageSetup.Zoom = False
		.PageSetup.FitToPagesWide = 1
		.PageSetup.FitToPagesTall = False
	End With
	
	On Error Resume Next
	Mkdir "c:\"+ dirname  'if exists then it doesnt do anything, otherwise it create sthe directory
			'if the directory does exist kill any previous instance of the file
	
	
	Dim workspace As New NotesUIWorkspace
	Dim response As Variant
	Dim values(1) As Variant
	values(0) = "Download File"
	values(1) = "Export to you Mail Box"
	response = workspace.Prompt (PROMPT_OKCANCELLIST, _
	"Please Select Export Excel Choice", _
	"Select Choice to Export.", _
	values(0), values)
	If Isempty (response) Then
		Messagebox "User canceled", , "Do you Export Excel next time"
		'xlapp.activeworkbook.close
		xlapp.quit
		Set xlapp=Nothing
	Else
		If response =  "Download File" Then
			'xlapp.activeworkbook.saveas "c:\"+ dirname+"\" + filename+" "+Cstr(Day(Today))+Cstr(Month(Today))+Cstr(Year(Today)) + ".xlsx"
			xlapp.activeworkbook.saveas "c:\"+ dirname+"\" + filename + ".xlsx"
			
			xlapp.activeworkbook.close
			xlapp.quit
			Set xlapp=Nothing
			'Messagebox "Download File Complete "+"c:\"+ dirname+"\" + filename +" "+Cstr(Day(Today))+Cstr(Month(Today))+Cstr(Year(Today))+ ".xlsx" , 64 , "Export Complete .." 	
			'Messagebox "Download File Complete" , 64 , "Export Complete .." 	
			Messagebox "Download File Complete "+"c:\"+ dirname+"\" + filename + ".xlsx" , 64 , "Export Complete .." 	
			
		Else
			Kill "c:\"+ dirname+"\" + filename + ".xlsx"
			xlapp.activeworkbook.saveas "c:\"+ dirname+"\" + filename + ".xlsx"
			xlapp.activeworkbook.close
			xlapp.quit
			Set xlapp=Nothing
	'Export End
			
     ' Put attachment in memo
			
			Dim docM As notesdocument
			Dim rtBody As notesrichtextitem
			Dim object As notesembeddedobject	
			Set docM=New notesdocument(db)
			Set rtBody=New notesrichtextitem(docM,"Body")
			docM.Form="Memo"
			strUser=sess.UserName
			docM.SendTo=strUser
			Set object=rtBody.EmbedObject(EMBED_ATTACHMENT,"","C:\"+ strDIRName + "\" + strFileName + ".xlsx")
			docM.Subject=ExcelHeader
			Call docM.send(False,strUser)
'clear up afterwards
			
			Kill "C:\" + strDIRName +"\" + strfilename + ".xlsx"
			Rmdir "C:\" + strDIRName 
			
			Print "Export Complete" 
			Messagebox "This file has been exported to your mail " , 64 , "Export Complete .." 	
			
		End If
		
	End If
	
End Sub