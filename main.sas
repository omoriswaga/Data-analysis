proc import datafile="/home/u59534451/Research/DocOld.xlsx" dbms=xlsx out=record1 replace;
run;

proc import datafile="/home/u59534451/Research/DocNew.xlsx" dbms=xlsx out=record2 replace;
run;


proc sql;
create table record2_001 as
select *
from record2
order by SUBJID;
quit;

proc sql;
create table record1_001 as
select *
from record1
order by SUBJID;
quit;

proc compare data=record2_001 comp=record1_001
 out=outcomp outnoequal outbase outcomp outdif noprint;
 id SUBJID;
 run;
 
data add_del;
 set outcomp;
 by SUBJID _type_;
 length change $10;
 if first.SUBJID and last.SUBJID then
 do;
 if _type_="BASE" then
 do; change='ADDED'; end;
 else if _type_="COMPARE" then
 do; change='DELETED'; end;
 end;
 run;
 
 proc transpose data= add_del out= t_add_del;
 by  SUBJID change;
 id _type_;
 var SUBJID VISITNUM VISIT UNSCHED SVSTDTC DOMAIN "LAB ID"N "LAB NAME"N REPEATNUMBER SREPEATID SREPEATNUMBER "ENTERED BY"N "ENTERED DATE"N "LAST CHANGED BY"N "LAST CHANGED DATE"N DELETED AEYN AEYN_R AEYN_F AEYN_D;
 quit; 
 
 data chg (where=(change ^= ' '));
 set t_add_del;
 by SUBJID;
length chgtxt chgval $1000;
diff_num = indexc(dif,"0123456789");
 diff_char= index(dif,"X");
 if diff_char>0 or diff_num>0 then
 do;
 change='UPDATED'; chgtxt = strip(_name_)||">OLD:"||strip(compare);
 chgval=strip(chgval)||'#'||strip(chgtxt);
 end;
 
data recordupdate;
set record2_001;
AEYN_char = put(AEYN,6.);
AEYN_D_char = put(AEYN_D,6.);
subject_id = SUBJID;
 
proc sql;
   create table found as
   select *
   from chg full outer join recordupdate
   on ( upcase(chg.BASE) contains upcase(recordupdate.VISITNUM)
   OR upcase(chg.BASE) contains upcase(recordupdate.VISIT)
   OR upcase(chg.BASE) contains upcase(recordupdate.UNSCHED)
   OR upcase(chg.BASE) contains upcase(recordupdate.SVSTDTC)
   OR upcase(chg.BASE) contains upcase(recordupdate.DOMAIN)
   OR upcase(chg.BASE) contains upcase(recordupdate."LAB ID"N )
   OR upcase(chg.BASE) contains upcase(recordupdate."LAB NAME"N)
   OR upcase(chg.BASE) contains upcase(recordupdate.REPEATNUMBER)
   OR upcase(chg.BASE) contains upcase(recordupdate.SREPEATID)
   OR upcase(chg.BASE) contains upcase(recordupdate.SREPEATNUMBER)
   OR upcase(chg.BASE) contains upcase(recordupdate."ENTERED BY"N)
   OR upcase(chg.BASE) contains upcase(recordupdate."ENTERED DATE"N)
   OR upcase(chg.BASE) contains upcase(recordupdate."LAST CHANGED BY"N)
   OR upcase(chg.BASE) contains upcase(recordupdate."LAST CHANGED DATE"N)
   OR upcase(chg.BASE) contains upcase(recordupdate.DELETED)
   OR upcase(chg.BASE) contains upcase(recordupdate.AEYN_char)
   OR upcase(chg.BASE) contains upcase(recordupdate.AEYN_R)
   OR upcase(chg.BASE) contains upcase(recordupdate.AEYN_F)
   OR upcase(chg.BASE) contains upcase(recordupdate.AEYN_D_char) ) and chg.SUBJID = recordupdate.SUBJID;
quit;

/*
 data want;
 merge chg record2;
 by SUBJID;
 run;
*/
proc sql;
	create table NoDuplicates as
	select distinct CHANGE, _NAME_, COMPARE, BASE, subject_id, DOMAIN, VISITNUM, VISIT, UNSCHED, SVSTDTC, "LAB ID"N, "LAB NAME"N, REPEATNUMBER, SREPEATID,  SREPEATNUMBER, "ENTERED BY"N, "ENTERED DATE"N, "LAST CHANGED BY"N, "LAST CHANGED DATE"N, DELETED, AEYN, AEYN_R, AEYN_F,AEYN_D
	from found
	where subject_id <> "";
quit;

data logic;
	set NoDuplicates;
	if (compare NE "" or Base NE "") and (compare=Base) then check = 1;

data NoDuplicates;
	set logic;
	where check NE 1;
	drop check;

ods excel file = "/home/u59534451/Research/output1.xlsx"
 options (sheet_name = "sheetname");
proc report data= NoDuplicates;
compute change;
if change = "UPDATED" then do;
call define(_row_,"style","style=[font_weight=bold background=lightgreen]");
end;
if change = "ADDED" then do;
call define(_row_,"style","style=[font_weight=bold background=lightred]");
end;
if change = "DELETED" then do;
call define(_row_,"style","style=[font_weight=bold background=lightgray]");
end;
endcomp;





%macro diff_report;
ods excel file = "/home/u59534451/Research/output1.xlsx"
 options (sheet_name = "sheetname");
proc report data= want;
%let dsid = %sysfunc(open(want,i));
%do i=1 %to %sysfunc(attrn(&dsid,nvars));
%let varn=%sysfunc(varname(&dsid,&i));
compute &varn./char length=20; 
%if &varn = "change" %then %do;
call define(_col_,"style","style=[font_weight=bold background=lightgreen]");
%end;
%else %if &varn in ("ADDED") %then %do;
call define(_col_,"style","style=[font_weight=bold background=lightred]");
%end;
%else %if &varn in ("Deleted") %then %do;
call define(_col_,"style","style=[font_weight=bold background=lightgray flyover=' "||strip(chgtxt)||" ']");
%end;
endcomp ;
%end;
run;
%mend diff_report;
 
 
 
 
 
 
 
 
 
 
 
 

