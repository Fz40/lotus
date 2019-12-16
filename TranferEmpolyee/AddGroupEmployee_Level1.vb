@Command( [EditDocument]; "1" );
Choice :=
@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "MasterCompany" ; "รายชื่อ บริษัท ";"กรุณาเลือก บริษัท" ; 3);
@If(Choice ="";@Return("");"");

REM {
1.CompanyCode
2.company
};

@Command([ViewRefreshFields]);

@Do(
		@SetField("FindCompanyID";@Word(Choice;"^";1));
		@SetField("FindCompany";@Word(Choice;"^";2));
		@SetField("FindGroupID";"");
		@SetField("FindGroup";"");
		
		@SetField("FindFieldID";"");
		@SetField("FindField";"");
		@SetField("FindDepartmentID";"");
		@SetField("FindDepartment";"");
		@SetField("FindSectionsID";"");
		@SetField("FindSections";"");
		@SetField("FindDivisionID";"");
		@SetField("FindDivision";"");
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		);
@Command([ViewRefreshFields])