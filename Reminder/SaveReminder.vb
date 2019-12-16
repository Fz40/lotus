Sub Click(Source As Button)
	
	Dim chkRepeat As String
	Dim CountEndDate As  Integer
	Dim CountTotalNoti As Integer
	Dim Start_Date As Variant
	Dim End_Date As Variant 
	Dim Cur_Date As Variant 
	Dim dateAdjust () As Variant
	Dim AfterdateAdjust () As Variant
	Dim dateNoti () As Variant
	Dim item As notesItem
	Dim KeyNameField As String
	Dim KeyNameOptField As String
	Dim valueField As Variant
	Dim valueOptField As Variant
	Dim l_dateAdjust As Integer
	Dim l_AfterdateAdjust As Integer
	
	Set ws = New NotesUiWorkspace
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	
	If ValidateField(curdoc,uidoc) = False Then
		Exit Sub
	End If
	
	Dim dateTime As New NotesDateTime(curdoc.StartsDate(0))
	
	chkRepeat = Cstr(curdoc.Repeat(0))
	Start_Date = curdoc.StartsDate(0)
	
	'If (uidoc.IsNewDoc) Or ( Start_Date >= Today()) Then
	
	l_dateAdjust = 1
	l_AfterdateAdjust = 1
	
	Redim dateAdjust(0)
	Redim AfterdateAdjust (0)
	
	Select Case chkRepeat
		
	Case "OneDay"
		If curdoc.TotalNoti(0) > 0 Then
			Set  dateTime = New NotesDateTime(curdoc.StartsDate(0))
			curdoc.ReminderDateTotal = CDat(dateTime.DateOnly)
			CountTotalNoti = Cint(curdoc.TotalNoti(0))
			
			For i=1 To CountTotalNoti
				KeyNameField = "Noti_"
				KeyNameField = KeyNameField + Cstr(i)
				KeyNameOptField = "Opt_"
				KeyNameOptField = KeyNameOptField + Cstr(i)
				valueField = curdoc.GetItemValue(KeyNameField)
				valueOptField = curdoc.GetItemValue(KeyNameOptField)
				Forall ItemValues In curdoc.ReminderDateTotal
					Set  dateTime = New NotesDateTime(ItemValues)
					If valueOptField(0) = "B" Then
						Call dateTime.AdjustDay(Cint(valueField(0))*-1)
						Redim Preserve dateAdjust(l_dateAdjust-1)
						dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
						'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
						l_dateAdjust = l_dateAdjust +1
					Elseif valueOptField(0) = "A" Then
						Call dateTime.AdjustDay(Cint(valueField(0)))
						Redim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
						AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
						'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
						l_AfterdateAdjust = l_AfterdateAdjust +1
					End If
					
				End Forall
				
			Next
			'curdoc.NotificationDateTotal = dateAdjust
			'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
			curdoc.NotificationDateTotal =  Arrayunique (dateAdjust)
			
			'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
			curdoc.NotificationAfterDateTotal =  Arrayunique (AfterdateAdjust)
			'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
		Else
			Set  dateTime = New NotesDateTime(curdoc.StartsDate(0))
			curdoc.NotificationDateTotal = CDat(dateTime.DateOnly)
			curdoc.ReminderDateTotal = Null
		End If
		
		
	Case "Daily"
		If curdoc.CheckNeverEnd(0) <>"" Then
			
			Call dateTime.AdjustYear(2)
			Start_Date =curdoc.StartsDate(0)
			End_Date = dateTime.LSLocalTime
			CountDate = End_Date - Start_Date
			Set  dateTime = New NotesDateTime(curdoc.StartsDate(0))
			For i = 0 To CountDate
				If i=0 Then
					Call dateTime.AdjustDay(0)
				Else
					Call dateTime.AdjustDay(1)
				End If
				Redim Preserve dateAdjust(i)
				dateAdjust(i) = CDat(dateTime.DateOnly)
			Next
			curdoc.ReminderDateTotal = Arrayunique (dateAdjust)
				'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
			
			
		Else
			
			Start_Date =curdoc.StartsDate(0)
			End_Date = curdoc.EndDate(0)
			CountDate = End_Date - Start_Date
			For i = 0 To CountDate
				If i=0 Then
					Call dateTime.AdjustDay(0)
				Else
					Call dateTime.AdjustDay(1)
				End If
				Redim Preserve dateAdjust(i)
				dateAdjust(i) = CDat(dateTime.DateOnly)
			Next
			curdoc.ReminderDateTotal = Arrayunique (dateAdjust)
				'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
			
		End If
		
		If curdoc.TotalNoti(0) > 0 Then
			CountTotalNoti = Cint(curdoc.TotalNoti(0))
			
			For i=1 To CountTotalNoti
				KeyNameField = "Noti_"
				KeyNameField = KeyNameField + Cstr(i)
				KeyNameOptField = "Opt_"
				KeyNameOptField = KeyNameOptField + Cstr(i)
				valueField = curdoc.GetItemValue(KeyNameField)
				valueOptField = curdoc.GetItemValue(KeyNameOptField)
				Forall ItemValues In curdoc.ReminderDateTotal
					Set  dateTime = New NotesDateTime(ItemValues)
					If valueOptField(0) = "B" Then
						Call dateTime.AdjustDay(Cint(valueField(0))*-1)
						Redim Preserve dateAdjust(l_dateAdjust-1)
						dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
						'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
						l_dateAdjust = l_dateAdjust +1
					Elseif valueOptField(0) = "A" Then
						Call dateTime.AdjustDay(Cint(valueField(0)))
						Redim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
						AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
						'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
						l_AfterdateAdjust = l_AfterdateAdjust +1
					End If
					
				End Forall
				
			Next
			'curdoc.NotificationDateTotal = dateAdjust
			'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
			curdoc.NotificationDateTotal =  Arrayunique (dateAdjust)
			
			'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
			curdoc.NotificationAfterDateTotal =  Arrayunique (AfterdateAdjust)
			'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
			
		Else
			curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
			curdoc.ReminderDateTotal = Null
		End If
		
		
	Case"Weekly"
		If curdoc.WeeklyType(0) <>"" Then
			Dim valueWeekDay As Integer
			Dim valueItem As Integer
			
			valueWeekDay = Weekday(Start_Date)
			Call dateTime.AdjustYear(2)
			Start_Date =curdoc.StartsDate(0)
			End_Date = dateTime.LSLocalTime
			CountDate = End_Date - Start_Date
			j=0
			Forall ItemValues In curdoc.WeeklyType
				Set  dateTime = New NotesDateTime(curdoc.StartsDate(0))
				valueItem = Cint(ItemValues) - valueWeekDay
				Call dateTime.AdjustDay(valueItem)
				For i = 0 To CountDate
					Cur_Date = dateTime.LSLocalTime
					If Cur_Date <= End_Date Then
						Call dateTime.AdjustDay(7)
						Cur_Date = dateTime.LSLocalTime
						If Start_Date < Cur_Date Then
							Redim Preserve dateAdjust(j)
							dateAdjust(j) = CDat(dateTime.DateOnly)
							j=j+1
						End If
					Else
						Exit For
					End If
				Next
			End Forall
			curdoc.ReminderDateTotal = Arrayunique (dateAdjust)
				'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
			
			If curdoc.TotalNoti(0) > 0 Then
				CountTotalNoti = Cint(curdoc.TotalNoti(0))
				
				For i=1 To CountTotalNoti
					KeyNameField = "Noti_"
					KeyNameField = KeyNameField + Cstr(i)
					KeyNameOptField = "Opt_"
					KeyNameOptField = KeyNameOptField + Cstr(i)
					valueField = curdoc.GetItemValue(KeyNameField)
					valueOptField = curdoc.GetItemValue(KeyNameOptField)
					Forall ItemValues In curdoc.ReminderDateTotal
						Set  dateTime = New NotesDateTime(ItemValues)
						If valueOptField(0) = "B" Then
							Call dateTime.AdjustDay(Cint(valueField(0))*-1)
							Redim Preserve dateAdjust(l_dateAdjust-1)
							dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
						'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
							l_dateAdjust = l_dateAdjust +1
						Elseif valueOptField(0) = "A" Then
							Call dateTime.AdjustDay(Cint(valueField(0)))
							Redim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
							AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
						'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
							l_AfterdateAdjust = l_AfterdateAdjust +1
						End If
						
					End Forall
					
				Next
			'curdoc.NotificationDateTotal = dateAdjust
			'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
				curdoc.NotificationDateTotal =  Arrayunique (dateAdjust)
				
			'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
				curdoc.NotificationAfterDateTotal =  Arrayunique (AfterdateAdjust)
			'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
				
			Else
				curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
				curdoc.ReminderDateTotal = Null
			End If
			
		End If
		
	Case"Monthy"
		If curdoc.MonthlyType(0) <>"" Then
			
			Dim Cur_Month As Integer 
			Dim Cur_Year As Integer 
			
			Call dateTime.AdjustYear(2)
			Start_Date =curdoc.StartsDate(0)
			End_Date = dateTime.LSLocalTime
			CountDate = End_Date - Start_Date
			Cur_Month = Month(Start_Date)
			Cur_Year = Year(Start_Date)
			j=0
			Forall ItemValues In curdoc.MonthlyType
				Cur_Date = Datenumber(Cur_Year,Cur_Month, ItemValues)
				Set  dateTime = New NotesDateTime(Cur_Date)
				For i = 0 To CountDate
					Cur_Date = dateTime.LSLocalTime
					If Cur_Date <= End_Date Then
						Call dateTime.AdjustMonth(1)
						Cur_Date = dateTime.LSLocalTime
						If Start_Date < Cur_Date Then
							Redim Preserve dateAdjust(j)
							dateAdjust(j) = CDat(dateTime.DateOnly)
							j=j+1
						End If
					Else
						Exit For
					End If
				Next
			End Forall
			curdoc.ReminderDateTotal = Arrayunique (dateAdjust)
				'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
			
			If curdoc.TotalNoti(0) > 0 Then
				CountTotalNoti = Cint(curdoc.TotalNoti(0))
				
				For i=1 To CountTotalNoti
					KeyNameField = "Noti_"
					KeyNameField = KeyNameField + Cstr(i)
					KeyNameOptField = "Opt_"
					KeyNameOptField = KeyNameOptField + Cstr(i)
					valueField = curdoc.GetItemValue(KeyNameField)
					valueOptField = curdoc.GetItemValue(KeyNameOptField)
					Forall ItemValues In curdoc.ReminderDateTotal
						Set  dateTime = New NotesDateTime(ItemValues)
						If valueOptField(0) = "B" Then
							Call dateTime.AdjustDay(Cint(valueField(0))*-1)
							Redim Preserve dateAdjust(l_dateAdjust-1)
							dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
						'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
							l_dateAdjust = l_dateAdjust +1
						Elseif valueOptField(0) = "A" Then
							Call dateTime.AdjustDay(Cint(valueField(0)))
							Redim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
							AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
						'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
							l_AfterdateAdjust = l_AfterdateAdjust +1
						End If
						
					End Forall
					
				Next
			'curdoc.NotificationDateTotal = dateAdjust
			'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
				curdoc.NotificationDateTotal =  Arrayunique (dateAdjust)
				
			'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
				curdoc.NotificationAfterDateTotal =  Arrayunique (AfterdateAdjust)
			'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
				
			Else
				curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
				curdoc.ReminderDateTotal = Null
			End If
			
		End If
		
	Case"Yearly"
		CountDate = curdoc.YearlyType(0)
		For i = 0 To CountDate -1
			Call dateTime.AdjustYear(1)
			Redim Preserve dateAdjust(i)
			dateAdjust(i) = CDat(dateTime.DateOnly)
		Next
		curdoc.ReminderDateTotal = Arrayunique (dateAdjust)
			'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
		
		If curdoc.TotalNoti(0) > 0 Then
			CountTotalNoti = Cint(curdoc.TotalNoti(0))
			
			For i=1 To CountTotalNoti
				KeyNameField = "Noti_"
				KeyNameField = KeyNameField + Cstr(i)
				KeyNameOptField = "Opt_"
				KeyNameOptField = KeyNameOptField + Cstr(i)
				valueField = curdoc.GetItemValue(KeyNameField)
				valueOptField = curdoc.GetItemValue(KeyNameOptField)
				Forall ItemValues In curdoc.ReminderDateTotal
					Set  dateTime = New NotesDateTime(ItemValues)
					If valueOptField(0) = "B" Then
						Call dateTime.AdjustDay(Cint(valueField(0))*-1)
						Redim Preserve dateAdjust(l_dateAdjust-1)
						dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
						'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
						l_dateAdjust = l_dateAdjust +1
					Elseif valueOptField(0) = "A" Then
						Call dateTime.AdjustDay(Cint(valueField(0)))
						Redim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
						AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
						'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
						l_AfterdateAdjust = l_AfterdateAdjust +1
					End If
					
				End Forall
				
			Next
			'curdoc.NotificationDateTotal = dateAdjust
			'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
			curdoc.NotificationDateTotal =  Arrayunique (dateAdjust)
			
			'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
			curdoc.NotificationAfterDateTotal =  Arrayunique (AfterdateAdjust)
			'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
			
		Else
			curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
			curdoc.ReminderDateTotal = Null
		End If
		
	Case"Custom"
		
		curdoc.ReminderDateTotal = curdoc.CustomDate
		
		If curdoc.TotalNoti(0) > 0 Then
			CountTotalNoti = Cint(curdoc.TotalNoti(0))
			
			For i=1 To CountTotalNoti
				KeyNameField = "Noti_"
				KeyNameField = KeyNameField + Cstr(i)
				KeyNameOptField = "Opt_"
				KeyNameOptField = KeyNameOptField + Cstr(i)
				valueField = curdoc.GetItemValue(KeyNameField)
				valueOptField = curdoc.GetItemValue(KeyNameOptField)
				Forall ItemValues In curdoc.ReminderDateTotal
					Set  dateTime = New NotesDateTime(ItemValues)
					If valueOptField(0) = "B" Then
						Call dateTime.AdjustDay(Cint(valueField(0))*-1)
						Redim Preserve dateAdjust(l_dateAdjust-1)
						dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
						'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
						l_dateAdjust = l_dateAdjust +1
					Elseif valueOptField(0) = "A" Then
						Call dateTime.AdjustDay(Cint(valueField(0)))
						Redim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
						AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
						'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
						l_AfterdateAdjust = l_AfterdateAdjust +1
					End If
					
				End Forall
				
			Next
			'curdoc.NotificationDateTotal = dateAdjust
			'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
			curdoc.NotificationDateTotal =  Arrayunique (dateAdjust)
			
			'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
			curdoc.NotificationAfterDateTotal =  Arrayunique (AfterdateAdjust)
			'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
			
		Else
			curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
			curdoc.ReminderDateTotal = Null
		End If
		
	End Select
	
		'Call WriteHistory(curdoc,actionDate)	
	
	If  Cstr(curdoc.StartsDate(0)) = Cstr(Today()) Then
		Call SendMail(curdoc)
	End If
	'Else
		'Msgbox "กรณีที่ไม่ได้เป็น Reminder ใหม่ หรือ วันที่เริ่ม ไม่น้อยกว่าวันที่ปัจจุบัณ. ระบบจะทำการบันทึกเฉพาะส่วนที่นอกเหนื่อจากหัวข้อ When",0+48, "โปรดตรวจสอบ"
	'End If
	
	Call curdoc.Save(True,True)
	Call uidoc.Refresh
	Call uidoc.Save
	Call uidoc.Close	
	'Call ws.ViewRefresh		
	
	If curdoc.DocNo(0) = "" Then
		Dim sess As New NotesSession
		Dim db As NotesDatabase
		Set db = sess.CurrentDatabase
		Dim agent As NotesAgent
		Set agent = _
		db.GetAgent("GenerateDocNo")
		If agent.RunOnServer(curdoc.NoteID) = 0 Then
			Print "Agent ran",, "Success"
		Else
			Print "Agent did not run",, "Failure"
		End If
	End If	
	
End Sub