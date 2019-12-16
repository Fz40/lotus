@Command( [EditDocument]; "1" );

Choice :=
@If(FindDepartmentID = "" ;
	@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "Level5" ; "รายชื่อ ส่วนงาน ";"กรุณาเลือก ส่วนงาน" ; 11)
	;@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "MasterSection" ; "รายชื่อ ส่วนงาน ";"กรุณาเลือก ส่วนงาน" ; 8;FindDepartmentID));

@If(Choice ="";@Return("");"");

@Command([ViewRefreshFields]);


@If(FindDepartmentID = "" ;
		@Do(
		@SetField("FindCompanyID";@Word(Choice;"^";1));
		@SetField("FindCompany";@Word(Choice;"^";2));
		@SetField("FindGroupID";@Word(Choice;"^";3));
		@SetField("FindGroup";@Word(Choice;"^";4));
		
		@SetField("FindFieldID";@Word(Choice;"^";5));
		@SetField("FindField";@Word(Choice;"^";6));
		@SetField("FindDepartmentID";@Word(Choice;"^";7));
		@SetField("FindDepartment";@Word(Choice;"^";8));
		@SetField("FindSectionsID";@Word(Choice;"^";9));
		@SetField("FindSections";@Word(Choice;"^";10));
		@SetField("FindDivisionID";"");
		@SetField("FindDivision";"");
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		)
;
		@Do(	
		@SetField("FindSectionsID";@Word(Choice;"^";9));
		@SetField("FindSections";@Word(Choice;"^";10));
		@SetField("FindDivisionID";"");
		@SetField("FindDivision";"");
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		)
);

@Command([ViewRefreshFields])