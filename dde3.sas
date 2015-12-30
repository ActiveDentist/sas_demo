/*****************************************************************
    定义宏::使用DDE打开Excel的程序（使用绝对路径）
	doublet-style DDE fileref
*****************************************************************/
%macro startxl;

/**使用filename 来生成fileref**/
filename sas2xl dde 'excel|system';
data _null_;
	file sas2xl;
	put '[new(1)]'; /*init one spreadsheet*/
run;
options noxwait noxsync;
%if &syserr ne 0 %then %do;
	x '"C:\Program Files\Microsoft Office\OFFICE14\excel.exe"';
data _null_;
	x=sleep(10);
run;
%end;

%mend startxl;


%startxl;
