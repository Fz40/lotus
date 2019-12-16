%REM
	Agent Schedule Renew Year Reminder
	Created Oct 2, 2019 by Sarawut Somton/PTG
	Description: Comments for Agent
%END REM
Option Public
Option Declare
Use "CommonScript"
Sub Initialize
	Dim s As New NotesSession
	Dim db As NotesDatabase
	Dim curdoc As NotesDocument
	Dim dc As NotesDocumentCollection

	Dim chkRepeat As String
	Dim CountTotalNoti As Integer
	Dim Start_Date As Variant
	Dim End_Date As Variant 
	Dim Cur_Date As Variant 
	Dim dateAdjust () As Variant
	Dim AfterdateAdjust () As Variant
	Dim dateNoti () As Variant
	Dim item As NotesItem
	Dim KeyNameField As String
	Dim KeyNameOptField As String
	Dim valueField As Variant
	Dim valueOptField As Variant
	Dim l_dateAdjust As Integer
	Dim l_AfterdateAdjust As Integer
	Dim i,j,CountDate As Integer
	
	Dim strQuery As String
	If Month(Today) = 10 Then
		
		
		strQuery = {
		SELECT (
		(Form = "MainReminder")
		&(StatusMarkComplete !="Y")
	    &(Repeat= "Daily":"Weekly":"Monthy")
		&(@Today >= StartsDate))
		}

		Set db = s.Currentdatabase
		Set dc = db.Search(strQuery, Nothing,0)
		Set curdoc = dc.Getfirstdocument()

		
		
		While Not curdoc Is Nothing 

			Dim dateTime As New NotesDateTime(Today)
			chkRepeat = CStr(curdoc.Repeat(0))
			Start_Date = Today

			l_dateAdjust = 1
			l_AfterdateAdjust = 1
			
			ReDim dateAdjust(0)
			ReDim AfterdateAdjust (0)
			
			Select Case chkRepeat

			Case "Daily"
				If curdoc.CheckNeverEnd(0) <>"" Then
					
					Call dateTime.AdjustYear(2)
					Start_Date = Today
					End_Date = dateTime.LSLocalTime
					CountDate = End_Date - Start_Date
					Set  dateTime = New NotesDateTime(Today)
					For i = 0 To CountDate
						If i=0 Then
							Call dateTime.AdjustDay(0)
						Else
							Call dateTime.AdjustDay(1)
						End If
						ReDim Preserve dateAdjust(i)
						dateAdjust(i) = CDat(dateTime.DateOnly)
					Next
					curdoc.ReminderDateTotal = ArrayUnique (dateAdjust)
					'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
					
					
				Else
					
					Start_Date = Today
					End_Date = curdoc.EndDate(0)
					CountDate = End_Date - Start_Date
					If CountDate > 0 Then
						For i = 0 To CountDate
							If i=0 Then
								Call dateTime.AdjustDay(0)
							Else
								Call dateTime.AdjustDay(1)
							End If
							ReDim Preserve dateAdjust(i)
							dateAdjust(i) = CDat(dateTime.DateOnly)
						Next
						curdoc.ReminderDateTotal = ArrayUnique (dateAdjust)
						'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
					End If
					
					
				End If
				
				If CountDate > 0 Then
					If curdoc.TotalNoti(0) > 0 Then
						CountTotalNoti = CInt(curdoc.TotalNoti(0))
						
						For i=1 To CountTotalNoti
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + CStr(i)
							KeyNameOptField = "Opt_"
							KeyNameOptField = KeyNameOptField + CStr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							valueOptField = curdoc.GetItemValue(KeyNameOptField)
							ForAll ItemValues In curdoc.ReminderDateTotal
								Set  dateTime = New NotesDateTime(ItemValues)
								If valueOptField(0) = "B" Then
									Call dateTime.AdjustDay(CInt(valueField(0))*-1)
									ReDim Preserve dateAdjust(l_dateAdjust-1)
									dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
									'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
									l_dateAdjust = l_dateAdjust +1
								ElseIf valueOptField(0) = "A" Then
									Call dateTime.AdjustDay(CInt(valueField(0)))
									ReDim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
									AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
									'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
									l_AfterdateAdjust = l_AfterdateAdjust +1
								End If
								
							End ForAll
							
						Next
						'curdoc.NotificationDateTotal = dateAdjust
						'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
						curdoc.NotificationDateTotal =  ArrayUnique (dateAdjust)
						
						'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
						curdoc.NotificationAfterDateTotal =  ArrayUnique (AfterdateAdjust)
						'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
						
						
					Else
						curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
						curdoc.ReminderDateTotal = Null
					End If
				End If
				
				
				Case"Weekly"
				If curdoc.WeeklyType(0) <>"" Then
					Dim valueWeekDay As Integer
					Dim valueItem As Integer
					
					valueWeekDay = Weekday(Start_Date)
					Call dateTime.AdjustYear(2)
					'Start_Date = curdoc.StartsDate(0)
					End_Date = dateTime.LSLocalTime
					CountDate = End_Date - Start_Date
					j=0
					ForAll ItemValues In curdoc.WeeklyType
						Set  dateTime = New NotesDateTime(Today)
						valueItem = CInt(ItemValues) - valueWeekDay
						Call dateTime.AdjustDay(valueItem)
						For i = 0 To CountDate
							Cur_Date = dateTime.LSLocalTime
							If Cur_Date <= End_Date Then
								Call dateTime.AdjustDay(7)
								Cur_Date = dateTime.LSLocalTime
								If Start_Date < Cur_Date Then
									ReDim Preserve dateAdjust(j)
									
									dateAdjust(j) = CDat(dateTime.DateOnly)
									j=j+1
								End If
							Else
								Exit For
							End If
						Next
					End ForAll
					curdoc.ReminderDateTotal = ArrayUnique (dateAdjust)
					'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
					
					If curdoc.TotalNoti(0) > 0 Then
						CountTotalNoti = CInt(curdoc.TotalNoti(0))
						
						For i=1 To CountTotalNoti
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + CStr(i)
							KeyNameOptField = "Opt_"
							KeyNameOptField = KeyNameOptField + CStr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							valueOptField = curdoc.GetItemValue(KeyNameOptField)
							ForAll ItemValues In curdoc.ReminderDateTotal
								Set  dateTime = New NotesDateTime(ItemValues)
								If valueOptField(0) = "B" Then
									Call dateTime.AdjustDay(CInt(valueField(0))*-1)
									ReDim Preserve dateAdjust(l_dateAdjust-1)
									dateAdjust(l_dateAdjust-1) = dateAdjust(j) = CDat(dateTime.DateOnly)
									'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
									l_dateAdjust = l_dateAdjust +1
								ElseIf valueOptField(0) = "A" Then
									Call dateTime.AdjustDay(CInt(valueField(0)))
									ReDim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
									AfterdateAdjust(l_AfterdateAdjust-1) = dateAdjust(j) = CDat(dateTime.DateOnly)
									'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
									l_AfterdateAdjust = l_AfterdateAdjust +1
								End If
								
							End ForAll
							
						Next
						'curdoc.NotificationDateTotal = dateAdjust
						'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
						curdoc.NotificationDateTotal =  ArrayUnique (dateAdjust)
						
						'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
						curdoc.NotificationAfterDateTotal =  ArrayUnique (AfterdateAdjust)
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
					Start_Date =Today
					End_Date = dateTime.LSLocalTime
					CountDate = End_Date - Start_Date
					Cur_Month = Month(Start_Date)
					Cur_Year = Year(Start_Date)
					j=0
					ForAll ItemValues In curdoc.MonthlyType
						Cur_Date = DateNumber(Cur_Year,Cur_Month, ItemValues)
						Set  dateTime = New NotesDateTime(Cur_Date)
						For i = 0 To CountDate
							Cur_Date = dateTime.LSLocalTime
							If Cur_Date <= End_Date Then
								Call dateTime.AdjustMonth(1)
								Cur_Date = dateTime.LSLocalTime
								If Start_Date < Cur_Date Then
									ReDim Preserve dateAdjust(j)
									dateAdjust(j) = CDat(dateTime.DateOnly)
									j=j+1
								End If
							Else
								Exit For
							End If
						Next
					End ForAll
					curdoc.ReminderDateTotal = ArrayUnique (dateAdjust)
					'curdoc.ReminderDateTotal = sortDataFieldCustomDate(curdoc,"Reminder")
					
					If curdoc.TotalNoti(0) > 0 Then
						CountTotalNoti = CInt(curdoc.TotalNoti(0))
						
						For i=1 To CountTotalNoti
							KeyNameField = "Noti_"
							KeyNameField = KeyNameField + CStr(i)
							KeyNameOptField = "Opt_"
							KeyNameOptField = KeyNameOptField + CStr(i)
							valueField = curdoc.GetItemValue(KeyNameField)
							valueOptField = curdoc.GetItemValue(KeyNameOptField)
							ForAll ItemValues In curdoc.ReminderDateTotal
								Set  dateTime = New NotesDateTime(ItemValues)
								If valueOptField(0) = "B" Then
									Call dateTime.AdjustDay(CInt(valueField(0))*-1)
									ReDim Preserve dateAdjust(l_dateAdjust-1)
									dateAdjust(l_dateAdjust-1) = CDat(dateTime.DateOnly)
									'dateAdjust(l_dateAdjust-2) = Cstr(ItemValues)
									l_dateAdjust = l_dateAdjust +1
								ElseIf valueOptField(0) = "A" Then
									Call dateTime.AdjustDay(CInt(valueField(0)))
									ReDim Preserve AfterdateAdjust(l_AfterdateAdjust-1)
									AfterdateAdjust(l_AfterdateAdjust-1) = CDat(dateTime.DateOnly)
									'AfterdateAdjust(l_AfterdateAdjust-2) = Cstr(ItemValues)
									l_AfterdateAdjust = l_AfterdateAdjust +1
								End If
								
							End ForAll
							
						Next
						'curdoc.NotificationDateTotal = dateAdjust
						'curdoc.ReminderDateTotal = curdoc.StartsDate(0)
						curdoc.NotificationDateTotal =  ArrayUnique (dateAdjust)
						
						'curdoc.NotificationDateTotal = sortDataFieldCustomDate(curdoc,"Noti")
						curdoc.NotificationAfterDateTotal =  ArrayUnique (AfterdateAdjust)
						'curdoc.NotificationAfterDateTotal = sortDataFieldCustomDate(curdoc,"AfterNoti")
						
					Else
						curdoc.NotificationDateTotal = curdoc.ReminderDateTotal
						curdoc.ReminderDateTotal = Null
					End If
					
				End If

		End Select
			
			Call curdoc.Save(True,True)	
			Set curdoc = dc.Getnextdocument(curdoc)
		Wend	
	End If
End Sub