/*****************************************************************
    �����::ʹ��DDE��Excel�ĳ��򣨲�ʹ�þ���·����
	doublet-style DDE fileref
*****************************************************************/

%macro startxl;

options noxsync noxwait;
filename sas2xl dde 'excel|system';
data _null_;
 length fid rc start stop time 8;
 fid=fopen('sas2xl','s');
 if (fid le 0) then do;
 rc=system('start excel');
 start=datetime();
 stop=start+10;
 do while (fid le 0);
 fid=fopen('sas2xl','s');
 time=datetime();
 if (time ge stop) then fid=1;
 end;
 end;
 rc=fclose(fid);
run;


%mend startxl;

%startxl;


