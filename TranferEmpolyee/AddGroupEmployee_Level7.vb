@Command( [EditDocument]; "1" );

Choice :=
@If(FindDivisionID = "" ;
	@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "Level7" ; "รายชื่อ สาขา / สถานี ";"กรุณาเลือก สาขา / สถานี" ; 14)
	;@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "MasterLocation" ; "รายชื่อ สาขา / สถานี ";"กรุณาเลือก สาขา / สถานี" ; 10;FindDivisionID));

@If(Choice ="";@Return("");"");

@Command([ViewRefreshFields]);

@If(FindDivisionID = "" ;
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
		@SetField("FindDivisionID";@Word(Choice;"^";11));
		@SetField("FindDivision";@Word(Choice;"^";12));
		@SetField("FindBranchID";@Word(Choice;"^";13));
		@SetField("FindBranch";@Word(Choice;"^";14))
		)
;
		@Do(	
		@SetField("FindBranchID";@Word(Choice;"^";13));
		@SetField("FindBranch";@Word(Choice;"^";14))
		)
);

@Command([ViewRefreshFields])