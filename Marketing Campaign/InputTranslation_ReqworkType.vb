key := "IT(Tester)" +"^"+ PrivilegesChannel ;

tmpIT := @DbLookup("Notes":"Nocache";Server:ThisDBName;"MasterDataUnitByTypeUnitAndAndPrivilegesChannel";key;"NameOfWorker";[FailSilent]);
tmpIT := @If( tmpIT = "" | @IsError(tmpIT);@DbLookup("Notes":"NoCache";Server:ThisDBName;"MasterDataUnit";"IT(Tester)";"NameOfWorker";[FailSilent]);
tmpIT);

key := "Academy" +"^"+ PrivilegesChannel ;

tmpBible := @DbLookup("Notes":"Nocache";Server:ThisDBName;"MasterDataUnitByTypeUnitAndAndPrivilegesChannel";key;"NameOfWorker";[FailSilent]);
tmpBible := @If( tmpBible = "" | @IsError(tmpBible);@DbLookup("Notes":"NoCache";Server:ThisDBName;"MasterDataUnit";"Academy";"NameOfWorker";[FailSilent]);
tmpBible);

key := "Media" +"^"+ PrivilegesChannel ;

tmpMedia := @DbLookup("Notes":"Nocache";Server:ThisDBName;"MasterDataUnitByTypeUnitAndAndPrivilegesChannel";key;"NameOfWorker";[FailSilent]);
tmpMedia := @If( tmpMedia = "" | @IsError(tmpMedia);@DbLookup("Notes":"NoCache";Server:ThisDBName;"MasterDataUnit";"Media";"NameOfWorker";[FailSilent]);
tmpMedia);


@If(RequestWorkerType != RequestWorkerType_1 ;

FIELD ITTesterName := @Unique(@Name([CN];tmpIT));
FIELD AcademyName := @Unique(@Name([CN];tmpBible));
FIELD MediaName := @Unique(@Name([CN];tmpMedia));

"");

RequestWorkerType