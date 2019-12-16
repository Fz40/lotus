Function validateFieldMainForm(curdoc As NotesDocument) As Boolean
	Dim eval As Variant
	Dim eval1 As Variant
	Dim eval2 As Variant
	Dim appendMsg As String
	Dim Start_Date As Variant
	Dim End_Date As Variant
	Dim Start_Date_Bible As Variant
	Dim End_Date_Bible As Variant
	Dim Start_Date_Media As Variant
	Dim End_Date_Media As Variant
	Dim Start_Test_Date As Variant 
	Dim End_Test_Date As Variant
	
	validateFieldMainForm = False
    If curdoc.EmpTypeTransfer(0) = "" Then
        MsgBox "กรุณาระบุประเภทการโอนย้าย "+appendMsg,0+48, "โปรดตรวจสอบ"
        Call uidoc.Gotofield("EmpTypeTransfer")
        Exit Function
    End If
    If curdoc.requestGroup(0) = "" Then
        MsgBox "กรุณาระบุ  Subject "+appendMsg,0+48, "โปรดตรวจสอบ"
        Call uidoc.Gotofield("requestGroup")
        Exit Function
    End If
    


    validateFieldMainForm = True
End Function