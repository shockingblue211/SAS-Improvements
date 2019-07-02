/**********************************************************************************************
***********************************************************************************************
Macro Name: DICTIONARY
Created Date/Author/Contact: Oct. 2015/Yuan Liu, Yaqi Jia/YLIU31@emory.edu yjia24@emory.edu
Current Version: V1
Working Environment: SAS 9.4 English version

Purpose:  To produce a data dictionary based on a descriptive statistics summary table for each variable in the dataset. 
The frequency, including # of missing value will be generated for categorical variables; and 
summary statistics (n, mean, median, min, max, standard deviation, # of missing) for numerical 
variables.  

Notes: The data dictionary contains variable name,label, formatted and unformatted levels for categorical variables, and descriptive statistics for all variables. 
The order of variables is the original order in the database. When using this macro to add hyperlink on variable name, set LINK=T.
     

Parameters: 

DATASET        The name of the data set to be analyzed.
CLIST          List of categorical variables, separated by empty space.
NLIST          List of numerical variables, separated by empty space.
CSV            Set to T if you want a csv data dictionary and don't need to add link on variable name. If you set
               both csv and link to T, it will produce a html format dictionary. The default value is F.
OUTPATH        Path for output table to be stored.
FNAME          File name for output table.
LINK           Set to T if you want to add link variable on dictionary name variable. The default value is F.
LINKDATA       The name of the data containing the merge variable and link variable.
MERGEVAR       The name of the merge variable in linkdata by which SAS match dataset and linkdata.
LINKVAR        the name of link variable containing website for each mergevar variable.
DEBUG          Set to T if running in debug mode (optional).  Work datasets will not be deleted 
               in debug mode.  This is useful if you are editing the code or want to further 
               manipulate the resulting data sets.  The default value is F.

***********************************************************************************************
For more details, please see the related documentation
**********************************************************************************************/


%MACRO DICTIONARY(DATASET=, CLIST=, NLIST=,CSV=F,OUTPATH=, FNAME=, link=F,linkdata=,mergevar=,linkvar=,DEBUG=F); 
   
   %let link = %UPCASE(&link);
   %let csv = %UPCASE(&csv);
   %let debug = %UPCASE(&debug);
   

   /* Get list of data sets in work library to avoid deletion later */
   ODS EXCLUDE members Directory;
   ODS OUTPUT Members(nowarn)=_DataSetList;
   PROC DATASETS lib=work memtype=data;
   QUIT;

   /* If there are data sets in the work library */
   %if %sysfunc(exist(_DataSetList)) %then %do;
      PROC SQL noprint;
         select Name
         into :work_sets separated by ' '
         from _DataSetList;
      quit;
   %end;
   %else %do;
      %let work_sets =;
   %end;

   /* Save current options */
   PROC OPTSAVE out=_options;
   RUN;

   /* Format missing values consistently */
   OPTIONS MISSING = " ";

   * Character Variables;

   %IF &CLIST NE  %THEN %DO; 

      %LET N = 1; 

      %DO %UNTIL (%SCAN(&CLIST, &N) =   ); 
         %LET CVAR = %SCAN(&CLIST, &N); 

         ODS EXCLUDE "One-Way Frequencies";
         ods output "One-Way Frequencies" = _temp;
         proc freq data= &DATASET;
            table &CVAR/missprint PLOTS=NONE;
         run;

         data _temp2; 
            length var_name $256. levels $200. fvalue $256.;
            set _temp;
			   var_name = vname(&CVAR);
               levels = strip(vvalue(&CVAR));
			   fvalue = &CVAR;
               keep var_name fvalue levels Frequency Percent table;
           run;

           data _temp3 _temp4;
            set _temp2;
            if levels in  ("." " ") then output _temp3;
            if levels not in  ("." " ") then output _temp4 ;
         run;

         data _freq&N;
            set _temp4 _temp3; 
            if levels in  ("." " ") then levels = "Missing";
               DROP table;
         run;

         %LET N = %EVAL(&N+1);

      %END; 

      %LET N = %EVAL(&N-1);
      DATA _freq_all; 
         SET _freq1-_freq&N; 
      RUN; 
      
   %END;

/*merge the last two columns : freqency and percentage into N(%)   */

   DATA _freq_all; set _freq_all; 
   if Percent ~= . then measurement = TRIM(LEFT(PUT(Frequency, BEST8.2))) || " (" || 
         TRIM(LEFT(PUT(Percent,8.1))) || ")";
   else measurement = TRIM(LEFT(PUT(Frequency, BEST8.2))) ;
   keep var_name fvalue levels measurement ; 
   RUN;



   *NUMERIC VARIABLES ;
   %IF &NLIST NE  %THEN %DO; 

      %LET N = 1; 

      %DO %UNTIL (%SCAN(&NLIST, &N) =   ); 
         %LET NVAR = %SCAN(&NLIST, &N); 

         PROC MEANS DATA=&DATASET noprint; 
            var &NVAR; 
            output out=_summary (drop=_TYPE_ _FREQ_) mean=Mean median=Median min=Minimum 
                    max=Maximum std=Std nmiss=Nmiss/autoname;
         RUN;
          
         DATA _summary;
            set _summary;
            LABEL mean = 'Mean'
               median = 'Median'
               Minimum = 'Minimum'
               Maximum = 'Maximum'
               std = 'Std Dev'
               nmiss = 'Missing';
         RUN;
     
         proc transpose data=_summary out=_summaryt;
         run;

         data _NULL_; 
            set &dataset; 
            call symput('vv', put(vname(&NVAR),$256.));
         run;

         data _summary&N;
            length var_name $256. levels $50. ;
            set _summaryt;
			var_name = "&vv";
            levels = _LABEL_;
            /* This is named frequency just to be consistent with the categorical data */
            /* It is NOT a frequency */
            Frequency =  COL1;
			measurement = TRIM(LEFT(PUT(Frequency,8.2))) ;
            if levels = "N Miss" then levels = "Missing";
            DROP _NAME_ _LABEL_  Frequency COL1;
         run;

         %LET N = %EVAL(&N+1);

      %END; 

      %LET N = %EVAL(&N-1);
      DATA _summary_all; 
         SET _summary1 - _summary&N; 
      RUN; 

   %END;


   /* Combine categorical and numerical results */
   DATA _report;
      set 
      %if &clist ~= %STR() %then %do;
          _freq_all
      %end;
      %if &nlist ~= %STR() %then %do;
          _summary_all
      %end;;
    RUN;
          
   /*Extract total number of observations in the dataset*/
   proc sql noprint; 
      select count(*) into :totalN from &dataset;
   quit;

  /*Prepare the descriptive table for purpose of  data dictionary*/

   		proc contents data = &DATASET 
		out = _cont (keep = name LABEL varnum) noprint;
		run;
		
		data _cont;set _cont; if label = "" then label = name;
   		proc sort data=_cont;by name;run;
		proc sort data=_report; by var_name;run;

		data report_final;
		merge _report (in=a) _cont(in=b rename=(name=var_name));
		by var_name;
		if a;
		run;
  /* make the order of variable in dictionary same as original order*/
		proc sort data=report_final;by varnum;run;
		

	    %if &link=T %then %do;

		   /*delete duplicated variable name in order to merge link data*/ 
		   proc sort data=report_final out=mergedata nodupkey;by var_name;run;
		   proc sort data=&linkdata; by &mergevar;run;
        
           /*merge data dictionary with link data*/
		   data linkdata;
		   merge mergedata (in=a) &linkdata (in=b rename=(&mergevar=var_name &linkvar=linkvar));
		   by var_name;
		   if a;
		   keep varnum var_name linkvar;
		   run;

           proc sort data=linkdata;by varnum;run; 

		   proc sql noprint;
		   select var_name into: name separated by " " from linkdata;
		   quit;

		   /*creat macro variables for each variable name value whoses resolved reference is conresponding linkvar value*/
		   data _null_;
		   set linkdata;
		   call symput(var_name,linkvar);
		   run;
		 %end;


   *---- table template -----;  

   ODS PATH WORK.TEMPLAT(UPDATE)
   SASUSR.TEMPLAT(UPDATE) SASHELP.TMPLMST(READ);

   PROC TEMPLATE;
   DEFINE STYLE STYLES.TABLES;
   NOTES "MY TABLE STYLE"; 
   PARENT=STYLES.MINIMAL;

     STYLE SYSTEMTITLE /FONT_SIZE = 12pt     FONT_FACE = "TIMES NEW ROMAN";

     STYLE HEADER /
           FONT_FACE = "TIMES NEW ROMAN"
            CELLPADDING=8
            JUST=C
            VJUST=C
            FONT_SIZE = 10pt
           FONT_WEIGHT = BOLD; 

     STYLE TABLE /
            FRAME=HSIDES            /* outside borders: void, box, above/below, vsides/hsides, lhs/rhs */
           RULES=GROUP              /* internal borders: none, all, cols, rows, groups */
           CELLPADDING=6            /* the space between table cell contents and the cell border */
            CELLSPACING=6           /* the space between table cells, allows background to show */
            JUST=C
            FONT_SIZE = 10pt
           BORDERWIDTH = 0.5pt;  /* the width of the borders and rules */

     STYLE DATAEMPHASIS /
           FONT_FACE = "TIMES NEW ROMAN"
           FONT_SIZE = 10pt
           FONT_WEIGHT = BOLD;

     STYLE DATA /
           FONT_FACE = "TIMES NEW ROMAN" 
           FONT_SIZE = 10pt;

     STYLE SYSTEMFOOTER /FONT_SIZE = 9pt FONT_FACE = "TIMES NEW ROMAN" JUST=C;
   END;

   RUN; 

   *------- build the table -----;

   OPTIONS ORIENTATION=PORTRAIT MISSING = "-" NODATE;

   %if &csv=T and &link=F %then %do;
      ODS csv FILE= "&OUTPATH.&FNAME &SYSDATE..csv"; 
   %end;
   %if &link=T %then %do;
      ODS html FILE="&OUTPATH.&FNAME &SYSDATE..html";
   %end;

PROC REPORT DATA=report_final HEADLINE HEADSKIP CENTER STYLE(REPORT)={JUST=CENTER} SPLIT='~' nowd 
          SPANROWS LS=256;

      Title "&fname";

      COLUMNS  varnum var_name label fvalue levels  Measurement; 
      
      DEFINE levels/ DISPLAY   "Level"   STYLE(COLUMN) = {JUST = L CellWidth=20%};
      DEFINE Measurement/DISPLAY "N (%) = %trim(&totalN)" STYLE(COLUMN) = {JUST = R CellWidth=15%} ;
	  DEFINE varnum/order order=internal noprint;
	  DEFINE fvalue/DISPLAY "" STYLE(COLUMN) = {JUST = R CellWidth=5%};
	  DEFINE var_name/ Order order=data  "Variable Name"  STYLE(COLUMN) = {JUST = L CellWidth=20%};
	  DEFINE label/ Order order=data  "Variable Label"  STYLE(COLUMN) = {JUST = L CellWidth=20%};

       %if &link = T %then %do;
	   compute var_name ;
       %let i=1;
       %do %until (%scan(&name,&i)= );
       if var_name="%scan(&name, &i)" then do;
	   %let v=%scan(&name, &i);
   	   href="%superq(&v)";
       call define(_col_, "URLP", href);
	   end;
	   %let i=%eval(&i+1);
	   %end;
       endcomp;
	   %end;
       
	  compute after LABEL; line ''; endcomp; 
       
   RUN; 

   %if &csv = T and &link=F %then %do;
      ODS csv CLOSE; 
   %end;

    %if &link=T %then %do;
      ODS html CLOSE;
   %end;

   /* Reload original options that were in use before running the macro */
   PROC OPTLOAD data=_options;
   RUN;

   /* Only delete files if not in debug mode */
   %if &debug ~= T %then %do;

      /* If there are work data sets that should not be deleted */
      %if %sysevalf(%superq(work_sets)~=,boolean) %then %do;
         /* DELETE ALL TEMPORARY DATASETS that were created */
         proc datasets lib=work memtype=data noprint;  
            save &work_sets;
         quit;  
      %end;
      %else %do;
         proc datasets lib=work kill memtype=data noprint;  
         quit; 
      %end;
   %end;

%mend; 


