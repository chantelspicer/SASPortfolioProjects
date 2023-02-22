*--------------------------------Accessing Data-------------------------------------;

* Creating a macrovariable to replace the filepath;
%let result=/home/u61481243/EPG1V2/data;

*Ensuring the excel workbook adheres to SAS naming conventions;
options validvarname=v7;

* Turning on graphics;
ods graphics on;

*Creating a library called PROJECT to connect to the excel workbook;
libname PROJECT xlsx "&result/cardiovascular.xlsx";

* Importing the excel worksheet into the PROJECT library;
proc import datafile="&result/cardiovascular.xlsx" out=project.cardiovascular 
		dbms=xlsx replace;
	sheet=cardiovascular;
run;

*---------------------------Exploring and Validating Data---------------------------;

* Listing important indicators of heart disease with maximum heart rate in descending order;
%let heartRate=160;
%let bloodPressure=180;
%let cholesterol=200;

proc sort data=project.cardiovascular out=myocardial;
	by DESCENDING thalachh;
run;

proc print data=myocardial label;
	var age sex thalachh trtbps chol fbs cp;
	where thalachh >= &heartRate or trtbps >= &bloodPressure or chol >= &cholesterol;
	label sex='Sex: Male = 1, Female = 0' 
		thalachh='Maximum Heart Rate'
		trtbps='Resting Systolic BP (on admission)' 
 		chol='Total Cholesterol'
 		fbs='Fasting Blood Sugar > 6.7mmol/L (1 = true, 0 = false)' 
		cp='Angina Chest Pain (1 = typical; 2 = Atypical; 3= Non-anginal; 0 = Asymptomatic)';
run;

* Generating a report that lists the column names, data types (numeric/char), and length;
proc contents data=project.cardiovascular;
run;

* Calculating summary statistics;
proc means data=project.cardiovascular;
	var age trtbps chol thalachh;
run;

* Examines extreme values;
proc univariate data=project.cardiovascular;
	var trtbps chol thalachh;
run;

* Lists unique values and frequencies for 10 rows;
proc freq data=project.cardiovascular (obs=10);
 	tables age sex trtbps chol fbs thalachh;
run;

proc sort data=project.cardiovascular out=work.hearthealth;
	where cp=3 or trtbps >=180;
	by descending cp;
run;

*-------------------------------Analyzing and Reporting Data-------------------------------;

* Frequency reports and graphs for the selected data;
ods noproctitle;
title "Cardiovascular Risk Assessment";
proc freq data=project.cardiovascular order=freq nlevels;
    tables trtbps chol fbs thalachh / 
           nocum plots=freqplot(orient=horizontal scale=percent);
    label trtbps="Blood Pressure"
          chol="Cholesterol"
          fbs="Fasting Blood Sugar"
          thalachh="Resting Heart Rate";
run;

* Preparing statistical bar chart graphics and customizing the appearance of the graph;
title "Cardiovascular Risk Assessment - Blood Pressure";
proc sgplot data=project.cardiovascular;
	where trtbps >= 160;
    hbar trtbps / group=trtbps seglabel
                  fillattrs=(transparency=0.5) dataskin=crisp;
    keylegend / opaque across=1 position=bottomright
                location=inside;
    xaxis grid;
    label trtbps="Systolic BP";
run;
	
*-------------------------------Exporting Results-------------------------------;

* Exporting correlation results to an Excel file and renaming the sheet;
ods excel file="&result/cardiovascularRisk.xlsx" 
		  style=analysis
		  options(sheet_name='BP_Chol_Correlation');

title "Correlation of Blood Pressure and Cholesterol";
proc sgscatter data=project.cardiovascular;
	plot trtbps*chol;
run;

ods excel close;
	
* Removing an active connection to the file;
libname project clear;

ods proctitle;
title;