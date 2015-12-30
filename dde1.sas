/*****************************************************************
 *** Clear SAS log and output.
 *****************************************************************/
dm 'clear log; clear output';
options mrecall mprint mlogic symbolgen;
%macro createExcelGraph;
%local lib dsn xlsCmdPath outputPath xlsTemplatePath;
/*****************************************************************
 *** Excel application path.
 *****************************************************************/
%let xlsCmdPath = C:\Program Files\Microsoft Office\OFFICE14\excel.exe;
/*****************************************************************
 *** Data library and data set name. In this example, we will
 *** use "AIR" data set from "SASHELP" library.
 *****************************************************************/
%let lib = sashelp;
%let dsn = air;
/*****************************************************************
 *** Excel template file path. This file should contain the Excel
 *** macro called "Macro1".
 *****************************************************************/
%let xlsTemplatePath = D:\template.xlsm;
/*****************************************************************
 *** Excel output file that will have the generated Excel graph.
 *****************************************************************/
%let outputPath = D:\air.xls;
proc sql noprint;
 /*************************************************************
 *** Get number of rows in data set and store the count in
 *** macro variable "total".
 ***
 *** The dictionary.tables are special read-only tables that
 *** contains plenty information about SAS data sets.
 *************************************************************/
select nlobs into :total
 from dictionary.tables
where libname = upcase("&lib.") and memname = upcase("&dsn.");
 /*************************************************************
 *** Add a "total" variable and store the computed total in it.
 *** Since the Excel macro will be reading the third column
 *** (column C) to retrieve the total row count, it is
 *** important to make sure that the "total" variable is stored
 *** at the right location.
 ***
 *** The reason we choose the data set name to be "sheet1" is
 *** because when we do a PROC EXPORT, the Excel spreadsheet 
 *** is automatically named after the SAS data set name.
 *** To simplify the process, we choose "sheet1", which is the
 *** default name for the first Excel spreadsheet.
 *************************************************************/
create table sheet1 as
select date, air, &total. as total
 from &lib..&dsn.;
 quit;
/*****************************************************************
 *** Export the data set to excel file.
 *****************************************************************/
proc export data=sheet1 outfile="&outputPath." dbms=excel replace;
 run;
options noxsync noxwait;
/******************************************************************/
/* *** Open Excel application.*/
/* *****************************************************************/
x "'&xlsCmdPath.'";
/******************************************************************/
/* *** Suspend SAS for 5 seconds to allow Excel to be fully started.*/
/* *****************************************************************/
data _null_;
rc = sleep(5);
run;
/******************************************************************/
/* *** Create a file reference to the excel sheet.*/
/* *****************************************************************/
filename sas2xl dde 'excel|system';
/******************************************************************/
/* *** Use X4ML commands to communicate with Excel.*/
/* *** Although you can use X4ML commands to format the Excel*/
/* *** spreadsheet's look and feel (colors, borders etc...), it is*/
/* *** preferable to keep it to minimum use, and have the Excel*/
/* *** macro to do those formatting. This makes your SAS code look*/
/* *** clean and readable.*/
/* ****/
/* *** So, in the below steps, we will open the Excel template file*/
/* *** and execute "Macro1" Excel macro.*/
/* ******************************************************************/
 data _null_;
 file sas2xl;
/**************************************************************/
/* *** Open Excel template file in read-only mode. This file*/
/* *** should have the Excel macro that you have created.*/
/* *************************************************************/
put "[open(""&xlsTemplatePath."", 0 , true)]";
/**************************************************************/
/* *** Execute the Excel macro. By default, the macro name is*/
/* *** called "Macro1".*/
/* *************************************************************/
put "[run(""Macro1"")]";
/**************************************************************/
/* *** Close the Excel template file.*/
/* *************************************************************/
put '[file.close(false)]'; 

/**************************************************************/
/* *** Close Excel application.*/
/* *************************************************************/
put '[quit()]';
run;
/******************************************************************/
/* *** Suspend SAS for 5 seconds to allow Excel to be fully closed.*/
/* *****************************************************************/
/*data _null_;*/
/*rc = sleep(5);*/
/*run;*/
/******************************************************************/
/* *** Delete sheet1 data set.*/
/* *****************************************************************/
proc datasets library=work nodetails nolist;
	delete sheet1;
quit;
%mend createExcelGraph;
%createExcelGraph; 
