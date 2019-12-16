@Command( [EditDocument]; "1" );

Choice :=
@If(FindCompanyID = "" ;
	@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "Level2" ; "รายชื่อ กลุ่มธุรกิจ ";"กรุณาเลือก กลุ่มธุรกิจ" ; 5)
	;@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "MasterGroupBussiness" ; "รายชื่อ กลุ่มธุรกิจ ";"กรุณาเลือก กลุ่มธุรกิจ" ; 4;FindCompanyID));

@If(Choice ="";@Return("");"");

@Command([ViewRefreshFields]);

REM {
View Level2 Column 5
1.CompanyCode
2.Company
3.BusinessGroupCode
4.BusinessGroupName

View MasterGroupBussiness Column 4
1.CompanyCode
2.Company
3.BusinessGroupCode
4.BusinessGroupName
};

@If(FindCompanyID = "" ;
		@Do(
		@SetField("FindCompanyID";@Word(Choice;"^";1));
		@SetField("FindCompany";@Word(Choice;"^";2));
		@SetField("FindGroupID";@Word(Choice;"^";3));
		@SetField("FindGroup";@Word(Choice;"^";4));
		
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
		)
;
		@Do(
		@SetField("FindGroupID";@Word(Choice;"^";3));
		@SetField("FindGroup";@Word(Choice;"^";4));
		
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
		)
);

@Command([ViewRefreshFields])