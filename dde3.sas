/*****************************************************************
    �����::ʹ��DDE��Excel�ĳ���ʹ�þ���·����
	doublet-style DDE fileref
*****************************************************************/
%macro startxl;

/**ʹ��filename ������fileref**/
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
