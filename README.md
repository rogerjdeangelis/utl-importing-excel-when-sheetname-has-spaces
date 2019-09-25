# utl-importing-excel-when-sheetname-has-spaces
Importing excel when sheetname has spaces     

    SAS Forum: Importing excel when sheetname has spaces                                                                                
                                                                                                                                        
        Five Solutions (non involve 'proc import')                                                                                      
                                                                                                                                        
           a. libname engine                                                                                                            
           b. sas passthru                                                                                                              
           c. R xlconnect                                                                                                               
           d. R xlsx                                                                                                                    
           e. R mysql                                                                                                                   
                                                                                                                                        
    github                                                                                                                              
    https://tinyurl.com/y6cepdal                                                                                                        
    https://github.com/rogerjdeangelis/utl-importing-excel-when-sheetname-has-spaces                                                    
                                                                                                                                        
    SAS Forum                                                                                                                           
    https://tinyurl.com/y6cdt7ek                                                                                                        
    https://communities.sas.com/t5/SAS-Programming/Proc-import-with-space-in-sheet-name-how-to-mention-range/m-p/591616                 
                                                                                                                                        
    Hopefully SAS will eventually deprecate 'proc import/export'                                                                        
    *_                   _                                                                                                              
    (_)_ __  _ __  _   _| |_                                                                                                            
    | | '_ \| '_ \| | | | __|                                                                                                           
    | | | | | |_) | |_| | |_                                                                                                            
    |_|_| |_| .__/ \__,_|\__|                                                                                                           
            |_|                                                                                                                         
    ;                                                                                                                                   
                                                                                                                                        
    * make data;                                                                                                                        
                                                                                                                                        
    libname xel "d:/xls/tabSpc.xlsx";                                                                                                   
                                                                                                                                        
    options validvarname=any;                                                                                                           
                                                                                                                                        
    data xel."date today"n;                                                                                                             
      set sashelp.class;                                                                                                                
    run;quit;                                                                                                                           
                                                                                                                                        
    libname xel clear;                                                                                                                  
                                                                                                                                        
                                                                                                                                        
     d:/xls/tabSpce.xlsx                                                                                                                
                                                                                                                                        
      +---------------------------------------+                                                                                         
       |  A    |  B    |  C    |  D    |  E   |                                                                                         
       +--------------------------------------+                                                                                         
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                                                                         
       +-------+-------+-------+-------+------+                                                                                         
     2 |Barbara|14     |M      |63.5   |88    |                                                                                         
       ---------------------------------------+                                                                                         
     3 |James  |13     |M      |73.5   |78    |                                                                                         
       ---------------------------------------+                                                                                         
     ...                                                                                                                                
                                                                                                                                        
     [date today]  ==> Note Space                                                                                                       
                                                                                                                                        
                                                                                                                                        
    *            _               _                                                                                                      
      ___  _   _| |_ _ __  _   _| |_                                                                                                    
     / _ \| | | | __| '_ \| | | | __|                                                                                                   
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                    
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                   
                    |_|                                                                                                                 
    ;                                                                                                                                   
                                                                                                                                        
    WORK.WANT_LIBNAME total obs=19                                                                                                      
                                                                                                                                        
      Name       Sex    Age    Height    Weight                                                                                         
                                                                                                                                        
      Alfred      M      14     69.0      112.5                                                                                         
      Alice       F      13     56.5       84.0                                                                                         
      Barbara     F      13     65.3       98.0                                                                                         
      Carol       F      14     62.8      102.5                                                                                         
      Henry       M      14     63.5      102.5                                                                                         
      James       M      12     57.3       83.0                                                                                         
                                                                                                                                        
    ...                                                                                                                                 
                                                                                                                                        
    *           _ _ _                                                                                                                   
      __ _     | (_) |__  _ __   __ _ _ __ ___   ___                                                                                    
     / _` |    | | | '_ \| '_ \ / _` | '_ ` _ \ / _ \                                                                                   
    | (_| |_   | | | |_) | | | | (_| | | | | | |  __/                                                                                   
     \__,_(_)  |_|_|_.__/|_| |_|\__,_|_| |_| |_|\___|                                                                                   
                                                                                                                                        
    ;                                                                                                                                   
                                                                                                                                        
                                                                                                                                        
    libname xel "d:/xls/tabSpc.xlsx";                                                                                                   
                                                                                                                                        
    options validvarname=any;                                                                                                           
                                                                                                                                        
    data xel."date today"n;                                                                                                             
      set sashelp.class;                                                                                                                
    run;quit;                                                                                                                           
                                                                                                                                        
    libname xel clear;                                                                                                                  
                                                                                                                                        
    *_                             _   _                                                                                                
    | |__      _ __   __ _ ___ ___| |_| |__  _ __ _   _                                                                                 
    | '_ \    | '_ \ / _` / __/ __| __| '_ \| '__| | | |                                                                                
    | |_) |   | |_) | (_| \__ \__ \ |_| | | | |  | |_| |                                                                                
    |_.__(_)  | .__/ \__,_|___/___/\__|_| |_|_|   \__,_|                                                                                
              |_|                                                                                                                       
    ;                                                                                                                                   
                                                                                                                                        
    proc sql dquote=ansi;                                                                                                               
     connect to excel                                                                                                                   
        (Path="d:/xls/tabSpc.xlsx" );                                                                                                   
        create                                                                                                                          
            table want_pass as                                                                                                          
        select                                                                                                                          
            *                                                                                                                           
            from connection to Excel                                                                                                    
            (                                                                                                                           
             Select                                                                                                                     
                *                                                                                                                       
             from                                                                                                                       
               [date_today$]                                                                                                            
            );                                                                                                                          
        disconnect from Excel;                                                                                                          
    Quit;                                                                                                                               
                                                                                                                                        
    *                      _                                 _                                                                          
      ___      _ __  __  _| | ___ ___  _ __  _ __   ___  ___| |_                                                                        
     / __|    | '__| \ \/ / |/ __/ _ \| '_ \| '_ \ / _ \/ __| __|                                                                       
    | (__ _   | |     >  <| | (_| (_) | | | | | | |  __/ (__| |_                                                                        
     \___(_)  |_|    /_/\_\_|\___\___/|_| |_|_| |_|\___|\___|\__|                                                                       
                                                                                                                                        
    ;                                                                                                                                   
                                                                                                                                        
    %utlfkil("d:/xpt/want.xpt");                                                                                                        
                                                                                                                                        
    %utl_submit_r64('                                                                                                                   
    library(XLConnect);                                                                                                                 
    library(SASxport);                                                                                                                  
    wb <- loadWorkbook("d:/xls/tabSpc.xlsx");                                                                                           
    want <- readWorksheet(wb, sheet = 1, colTypes=c("character"));                                                                      
    write.xport(want,file="d:/xpt/want.xpt");                                                                                           
    ');                                                                                                                                 
                                                                                                                                        
    libname xpt xport "d:/xpt/want.xpt";                                                                                                
    data want_xlcon;                                                                                                                    
      set xpt.want;                                                                                                                     
    run;quit;                                                                                                                           
    libname xpt clear;                                                                                                                  
                                                                                                                                        
    *    _                 _                                                                                                            
      __| |    _ __  __  _| |_____  __                                                                                                  
     / _` |   | '__| \ \/ / / __\ \/ /                                                                                                  
    | (_| |_  | |     >  <| \__ \>  <                                                                                                   
     \__,_(_) |_|    /_/\_\_|___/_/\_\                                                                                                  
                                                                                                                                        
    ;                                                                                                                                   
                                                                                                                                        
    %utlfkil("d:/xpt/want.xpt");                                                                                                        
                                                                                                                                        
    %utl_submit_r64('                                                                                                                   
     library(xlsx);                                                                                                                     
     library(Hmisc);                                                                                                                    
     library(SASxport);                                                                                                                 
     want<-read.xlsx("d:/xls/tabSpc.xlsx",1,colClasses=rep("character",16),stringsAsFactors=FALSE);                                     
     write.xport(want,file="d:/xpt/want.xpt");                                                                                          
    ');                                                                                                                                 
                                                                                                                                        
                                                                                                                                        
    libname xpt xport "d:/xpt/want.xpt";                                                                                                
    data want_xlsx;                                                                                                                     
      set xpt.want;                                                                                                                     
    run;quit;                                                                                                                           
    libname xpt clear;                                                                                                                  
                                                                                                                                        
    *                                          _                                                                                        
      ___     _ __   _ __ ___  _   _ ___  __ _| |                                                                                       
     / _ \   | '__| | '_ ` _ \| | | / __|/ _` | |                                                                                       
    |  __/_  | |    | | | | | | |_| \__ \ (_| | |                                                                                       
     \___(_) |_|    |_| |_| |_|\__, |___/\__, |_|                                                                                       
                               |___/        |_|                                                                                         
    ;                                                                                                                                   
                                                                                                                                        
    /*                                                                                                                                  
    if you get this error                                                                                                               
     could not run statement: The used command is not allowed with this MySQL version                                                   
    open mysql client command line and enter the text below                                                                             
                                                                                                                                        
    set global local_infile=true;                                                                                                       
    */                                                                                                                                  
                                                                                                                                        
    %utl_submit_r64('                                                                                                                   
    library(RMySQL);                                                                                                                    
    library(DBI);                                                                                                                       
    library(XLConnect);                                                                                                                 
    db <- dbConnect(MySQL(), dbname = "sakila"                                                                                          
           ,host = "localhost"                                                                                                          
           ,port = 3306                                                                                                                 
           ,user = "root"                                                                                                               
           ,password = "xxxxxxxx");                                                                                                     
    wb <- loadWorkbook("d:/xls/tabSpc.xlsx");                                                                                           
    want_mysql <- readWorksheet(wb, sheet = 1, colTypes=c("character"));                                                                
    if (dbExistsTable(db, "want_mysql")) dbRemoveTable(db, "want_mysql");                                                               
    dbWriteTable(db, name = "want_mysql", value = want_mysql, row.names = FALSE, overwrite=TRUE);                                       
    dbDisconnect(db)                                                                                                                    
    ');                                                                                                                                 
                                                                                                                                        
    libname mysqllib mysql user=root password="xxxxxxxx" database=sakila port=3306;                                                     
                                                                                                                                        
    proc contents data=mysqllib.want_mysql;                                                                                             
    run;quit;                                                                                                                           
                                                                                                                                        
             Alphabetic List of Vari                                                                                                    
                                                                                                                                        
    #    Variable    Type     Len                                                                                                       
                                                                                                                                        
    3    Age         Char    1024                                                                                                       
    4    Height      Char    1024                                                                                                       
    1    Name        Char    1024                                                                                                       
    2    Sex         Char    1024                                                                                                       
    5    Weight      Char    1024                                                                                                       
                                                                                                                                        
                                                                                                                                        
    data want_mysql;                                                                                                                    
       set mysqllib.want_mysql;                                                                                                         
    run;quit;                                                                                                                           
                                                                                                                                        
    libname mysqllib clear;                                                                                                             
                                                                                                                                        
    %utl_optlen(inp=want_mysql,out=want_mysql);                                                                                         
                                                                                                                                        
                                                                                                                                        
                    Variables in                                                                                                        
                                                                                                                                        
        Variable    Type    Len                                                                                                         
                                                                                                                                        
        Name        Char      7                                                                                                         
        Sex         Char      1                                                                                                         
        Age         Char      2                                                                                                         
        Height      Char      4                                                                                                         
        Weight      Char      5                                                                                                         
                                                                                                                                        
    *_                                                                                                                                  
    | | ___   __ _                                                                                                                      
    | |/ _ \ / _` |                                                                                                                     
    | | (_) | (_| |                                                                                                                     
    |_|\___/ \__, |                                                                                                                     
             |___/                                                                                                                      
    ;                                                                                                                                   
                                                                                                                                        
    > library(RMySQL);library(DBI);library(XLConnect);db <- dbConnect(MySQL(),                                                          
    dbname = "sakila"       ,host = "localhost"       ,                                                                                 
    port = 3306       ,user = "root"       ,passw                                                                                       
    ord = "xxxxxxxx");wb <- loadWorkbook("d:/xls/tabSpc.xlsx");                                                                         
    want_mysql <- readWorksheet(wb, sheet = 1, colTypes=c("character"));                                                                
    if (dbExistsTable(db, "want_mysql")) dbRemov                                                                                        
    eTable(db, "want_mysql");dbWriteTable(db, name = "want_mysql",                                                                      
    value = want_mysql, row.names = FALSE, overwrite=TRUE);dbDisconnect(db)                                                             
                                                                                                                                        
    [1] TRUE                                                                                                                            
    [1] TRUE                                                                                                                            
    [1] TRUE                                                                                                                            
    >                                                                                                                                   
    NOTE: 7 lines were written to file PRINT.                                                                                           
    Stderr output:                                                                                                                      
    Loading required package: DBI                                                                                                       
    1: package 'RMySQL' was built under R version 3.5.3                                                                                 
    2: package 'DBI' was built under R version 3.5.3                                                                                    
    Loading required package: XLConnectJars                                                                                             
    XLConnect 0.2-15 by Mirai Solutions GmbH [aut],                                                                                     
      Martin Studer [cre],                                                                                                              
      The Apache Software Foundation [ctb, cph] (Apache POI),                                                                           
      Graph Builder [ctb, cph] (Curvesapi Java library)                                                                                 
    http://www.mirai-solutions.com                                                                                                      
    https://github.com/miraisolutions/xlconnect                                                                                         
    Warning messages:                                                                                                                   
    1: package 'XLConnect' was built under R version 3.5.3                                                                              
    2: package 'XLConnectJars' was built under R version 3.5.3                                                                          
    Picked up _JAVA_OPTIONS: -Xmx51200000                                                                                               
                                                                                                                                        
                                                                                                                                        
                                                                                                                                        
