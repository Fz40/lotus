Sub Click(Source As Button)
	
	Dim chkRepeat As String
	Dim CountEndDate As  Integer
	Dim CountTotalNoti As Integer
	Dim Start_Date As Variant
	Dim End_Date As Variant 
	Dim Cur_Date As Variant 
	Dim dateAdjust () As Variant
	Dim dateNoti () As Variant
	Dim item As notesItem
	Dim KeyNameField As String
	Dim valueField As Variant
	
	Set ws = New NotesUiWorkspace
	Set uidoc = ws.CurrentDocument
	Set curdoc = uidoc.Document
	
	If ValidateField(curdoc,uidoc) = False Then
		Exit Sub
	End If
	
	Dim dateTime As New NotesDateTime(curdoc.StartsDate(0))
	
	chkRepeat = Cstr(curdoc.Repeat(0))
	Start_Date = curdoc.StartsDate(0)
	
	If (uidoc.IsNewDoc) Or ( Start_Date >= Today()) Then
		
		Select Case chkRepeat
			
		Case "OneDay"
			If curdoc.TotalNoti(0) > 0 Then
				CountTotalNoti = Cint(curdoc.TotalNoti(0))
				For i=1 To CountTotalNoti
					Set  dateTime = New NotesDateTime(curdoc.StartsDate(0))
					KeyNameField = "Noti_"
					KeyNameField = KeyNameField + Cstr(i)
					valueField = curdoc.GetItemValue(KeyNameField)
					Call dateTime.AdjustDay(Cint(valueField(0))*-1)
					Redim Preserve dateAdjust(i-1)
					dateAdjust(i-1) = Cstr(dateTime.DateOnly)
				Next
				curdoc.NotificationDateTotal = dateAdjust
				curdoc.ReminderDateTotal = curdoc.StartsDate(0)
				curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
			Else
				curdoc.ReminderDateTotal = curdoc.StartsDate(0)
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
					dateAdjust(i) = Cstr(dateTime.DateOnly)
				Next
				curdoc.ReminderDateTotal = dateAdjust
				
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
					dateAdjust(i) = Cstr(dateTime.DateOnly)
				Next
				curdoc.ReminderDateTotal = dateAdjust
				
			End If
			
			If curdoc.TotalNoti(0) > 0 Then
				CountTotalNoti = Cint(curdoc.TotalNoti(0))
				j=1
				For i=1 To CountTotalNoti
					Forall ItemValues In curdoc.ReminderDateTotal
						Set  dateTime = New NotesDateTime(ItemValues)
						KeyNameField = "Noti_"
						KeyNameField = KeyNameField + Cstr(i)
						valueField = curdoc.GetItemValue(KeyNameField)
						Call dateTime.AdjustDay(Cint(valueField(0))*-1)
						Redim Preserve dateAdjust(j-1)
						dateAdjust(j-1) = Cstr(dateTime.DateOnly)
						j=j+1
					End Forall
					
				Next
				'curdoc.NotificationDateTotal = dateAdjust
				curdoc.NotificationDateTotal = Arrayunique (dateAdjust)
				'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
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
								dateAdjust(j) = Cstr(dateTime.DateOnly)
								j=j+1
							End If
						Else
							Exit For
						End If
					Next
				End Forall
				curdoc.ReminderDateTotal = dateAdjust
				
				If curdoc.TotalNoti(0) > 0 Then
					CountTotalNoti = Cint(curdoc.TotalNoti(0))
					j=1
					For i=1 To CountTotalNoti
						Forall ItemValues In curdoc.ReminderDateTotal
							Set  dateTime = New NotesDateTime(ItemValues)
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + Cstr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							Call dateTime.AdjustDay(Cint(valueField(0))*-1)
							Redim Preserve dateAdjust(j-1)
							dateAdjust(j-1) = Cstr(dateTime.DateOnly)
							j=j+1
						End Forall
						
					Next
					curdoc.NotificationDateTotal = dateAdjust
					'curdoc.NotificationDateTotal = Arrayunique (dateAdjust)
					'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
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
								dateAdjust(j) = Cstr(dateTime.DateOnly)
								j=j+1
							End If
						Else
							Exit For
						End If
					Next
				End Forall
				curdoc.ReminderDateTotal = dateAdjust
				
				If curdoc.TotalNoti(0) > 0 Then
					CountTotalNoti = Cint(curdoc.TotalNoti(0))
					j=1
					For i=1 To CountTotalNoti
						Forall ItemValues In curdoc.ReminderDateTotal
							Set  dateTime = New NotesDateTime(ItemValues)
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + Cstr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							Call dateTime.AdjustDay(Cint(valueField(0))*-1)
							Redim Preserve dateAdjust(j-1)
							dateAdjust(j-1) = Cstr(dateTime.DateOnly)
							j=j+1
						End Forall
						
					Next
					curdoc.NotificationDateTotal = dateAdjust
					'curdoc.NotificationDateTotal = Arrayunique (dateAdjust)
					'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
				End If

			End If
			
		Case"Yearly"
			CountDate = curdoc.YearlyType(0)
			For i = 0 To CountDate 
				If i=0 Then
					'Call dateTime.AdjustYear(0)
				Else
					Call dateTime.AdjustYear(1)
				End If
				Redim Preserve dateAdjust(i)
				dateAdjust(i) = Cstr(dateTime.DateOnly)
			Next
			curdoc.ReminderDateTotal = dateAdjust
			
			If curdoc.TotalNoti(0) > 0 Then
					CountTotalNoti = Cint(curdoc.TotalNoti(0))
					j=1
					For i=1 To CountTotalNoti
						Forall ItemValues In curdoc.ReminderDateTotal
							Set  dateTime = New NotesDateTime(ItemValues)
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + Cstr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							Call dateTime.AdjustDay(Cint(valueField(0))*-1)
							Redim Preserve dateAdjust(j-1)
							dateAdjust(j-1) = Cstr(dateTime.DateOnly)
							j=j+1
						End Forall
						
					Next
					curdoc.NotificationDateTotal = dateAdjust
					'curdoc.NotificationDateTotal = Arrayunique (dateAdjust)
					'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
				End If
			
		Case"Custom"
			
			curdoc.ReminderDateTotal = curdoc.CustomDate
			
			If curdoc.TotalNoti(0) > 0 Then
					CountTotalNoti = Cint(curdoc.TotalNoti(0))
					j=1
					For i=1 To CountTotalNoti
						Forall ItemValues In curdoc.ReminderDateTotal
							Set  dateTime = New NotesDateTime(ItemValues)
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + Cstr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							Call dateTime.AdjustDay(Cint(valueField(0))*-1)
							Redim Preserve dateAdjust(j-1)
							dateAdjust(j-1) = Cstr(dateTime.DateOnly)
							j=j+1
						End Forall
						
					Next
					curdoc.NotificationDateTotal = dateAdjust
					'curdoc.NotificationDateTotal = Arrayunique (dateAdjust)
					'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
				End If
			
		End Select
		
		'Call WriteHistory(curdoc,actionDate)	
		
		If  Cstr(curdoc.StartsDate(0)) = Cstr(Today()) Then
			Call SendMail(curdoc)
		End If
	Else
		Msgbox "กรุณาระบุ วันที่เริ่ม ไม่น้อยกว่าวันที่ปัจจุบัณ ",0+48, "โปรดตรวจสอบ"
	End If
	
	Call curdoc.Save(True,True)
	Call uidoc.Refresh
	Call uidoc.Save
	Call uidoc.Close	
	Call ws.ViewRefresh		
	
	
	
End Sub