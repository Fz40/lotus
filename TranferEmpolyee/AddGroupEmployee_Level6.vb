@Command( [EditDocument]; "1" );

Choice :=
@If(FindSectionsID = "" ;
	@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "Level6" ; "รายชื่อ แผนก ";"กรุณาเลือก แผนก" ; 13)
	;@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "MasterDivision" ; "รายชื่อ แผนก ";"กรุณาเลือก แผนก" ; 9;FindSectionsID));

@If(Choice ="";@Return("00");"");

@Command([ViewRefreshFields]);



@If(FindSectionsID = "" ;
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
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		)
;
		@Do(	

		@SetField("FindDivisionID";@Word(Choice;"^";11));
		@SetField("FindDivision";@Word(Choice;"^";12));
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		)
);

@Command([ViewRefreshFields])