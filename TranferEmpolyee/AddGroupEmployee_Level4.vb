@Command( [EditDocument]; "1" );

Choice :=
@If(FindFieldID = "" ;
	@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "Level4" ; "รายชื่อ ฝ่าย ";"กรุณาเลือก ฝ่าย" ; 9)
	;@PickList([Custom] : [Single] ;ServerEmployee:DBEmployee ; "MasterDept" ; "รายชื่อ ฝ่าย ";"กรุณาเลือก ฝ่าย" ; 7;FindFieldID));

@If(Choice ="";@Return("");"");

@Command([ViewRefreshFields]);

REM {
View Level2 Column 7
1.CompanyCode
2.Company
3.BusinessGroupCode
4.BusinessGroupName
5.BusinessLineCode
6.BusinessLineName

View MasterGroupBussiness Column 7
1.BusinessGroupCode
2.BusinessLineCode
3.BusinessLineName
4.CompanyCode
5.company
6.BusinessGroupName
};

@If(FindFieldID = "" ;
		@Do(
		@SetField("FindCompanyID";@Word(Choice;"^";1));
		@SetField("FindCompany";@Word(Choice;"^";2));
		@SetField("FindGroupID";@Word(Choice;"^";3));
		@SetField("FindGroup";@Word(Choice;"^";4));
		
		@SetField("FindFieldID";@Word(Choice;"^";5));
		@SetField("FindField";@Word(Choice;"^";6));
		@SetField("FindDepartmentID";@Word(Choice;"^";7));
		@SetField("FindDepartment";@Word(Choice;"^";8));
		@SetField("FindSectionsID";"");
		@SetField("FindSections";"");
		@SetField("FindDivisionID";"");
		@SetField("FindDivision";"");
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		)
;
		@Do(	
		@SetField("FindDepartmentID";@Word(Choice;"^";7));
		@SetField("FindDepartment";@Word(Choice;"^";8));
		@SetField("FindSectionsID";"");
		@SetField("FindSections";"");
		@SetField("FindDivisionID";"");
		@SetField("FindDivision";"");
		@SetField("FindBranchID";"");
		@SetField("FindBranch";"")
		)
);

@Command([ViewRefreshFields])