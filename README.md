# utl_excel_add_formulas
Using SAS to add a formula to an excel sheet

    ```  Three techniques to add formulas to excel(some work with existing workbook)  ```
    ```    ```
    ```    ```
    ```  *          _                         _  ```
    ```    ___   __| |___    _____  _____ ___| |  ```
    ```   / _ \ / _` / __|  / _ \ \/ / __/ _ \ |  ```
    ```  | (_) | (_| \__ \ |  __/>  < (_|  __/ |  ```
    ```   \___/ \__,_|___/  \___/_/\_\___\___|_|  ```
    ```  ;  ```
    ```    ```
    ```    ```
    ```  %utlfkil(d:\temp\formulas.xlsx);  ```
    ```    ```
    ```  ods excel file='d:\temp\formulas.xlsx';  ```
    ```  proc report data=sashelp.class split='@';  ```
    ```  column name sex age height weight bmi;  ```
    ```  define bmi / "BMI"  ```
    ```  computed format=3. style={tagattr="formula:(RC[-1]/(2.2*(RC[-2]/39.37)^2))"};  ```
    ```  compute bmi;  ```
    ```     bmi=0;  ```
    ```  endcomp;  ```
    ```  run;  ```
    ```    ```
    ```  ods excel close;  ```
    ```    ```
    ```  *_ _ _  ```
    ```  | (_) |__  _ __   __ _ _ __ ___   ___  ```
    ```  | | | '_ \| '_ \ / _` | '_ ` _ \ / _ \  ```
    ```  | | | |_) | | | | (_| | | | | | |  __/  ```
    ```  |_|_|_.__/|_| |_|\__,_|_| |_| |_|\___|  ```
    ```  ;  ```
    ```    ```
    ```  %utlfkil(d:/xls/formulas.xlsx);  ```
    ```  libname xel "d:/xls/formulas.xlsx" scan_text=no;  ```
    ```  data xel.have;  ```
    ```     set sashelp.class(obs=3);  ```
    ```     bmi=cats('=E',put(_n_+1,3.),'/(2.2*(D',put(_n_+1,3.),'/39.37)^2)');  ```
    ```     putlog bmi=;  ```
    ```  run;quit;  ```
    ```  libname xel clear;  ```
    ```    ```
    ```  * to activate the calculation of bmi,  ```
    ```   1. open the excel file highlite the bmi column  ```
    ```   2, find and replace '=' with '='  ```
    ```   3. save workbook  ```
    ```    ```
    ```  *____  ```
    ```  |  _ \  ```
    ```  | |_) |  ```
    ```  |  _ <  ```
    ```  |_| \_\  ```
    ```    ```
    ```  ;  ```
    ```    ```
    ```  * create a sheet;  ```
    ```  %utlfkil(d:\xls\xlass.xlsx);  ```
    ```  libname xel "d:/xls/xlass.xlsx";  ```
    ```  data xel.class;  ```
    ```    set sashelp.class;  ```
    ```  run;quit;  ```
    ```  libname xel clear;  ```
    ```    ```
    ```  %utl_submit_r64('  ```
    ```  source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);  ```
    ```  library(XLConnect);  ```
    ```  wb <- loadWorkbook("d:/xls/xlass.xlsx",create = FALSE);  ```
    ```  box<-getBoundingBox(wb, sheet="class");  ```
    ```  colname <- readWorksheet(wb, sheet="class", startRow = box[1,], startCol = box[1,],  ```
    ```  endRow = box[1,], endCol = box[4,]);  ```
    ```  colIndex <- which(names(colname) == "WEIGHT");  ```
    ```  rowIndex <- box[3,1];  ```
    ```  input <- paste(idx2cref(c(box[1,]+1, colIndex, rowIndex, colIndex)), collapse=":");  ```
    ```  formula <- paste("SUM(", input, ")", sep="");  ```
    ```  setCellFormula(wb, "class", rowIndex + 1, colIndex, formula);  ```
    ```  rows<-c((box[1,]+1):rowIndex);  ```
    ```  rowchr<-as.character(rows);  ```
    ```  s<-idx2aref(c((box[1,]+1), colIndex, (box[1,]+1), colIndex-1));  ```
    ```  fml<-paste(substr(s, 1, 1),rowchr,"*703/",substr(s, 4, 4),rowchr,"^2",sep="");  ```
    ```  cols<-rep(colIndex+1,rowIndex-1);  ```
    ```  writeWorksheet(wb, "BMI", sheet = "class", header=F, startRow = 1,startCol=colIndex+1);  ```
    ```  setCellFormula(wb, "class", rows, cols, fml);  ```
    ```  saveWorkbook(wb);  ```
    ```  endsubmit;  ```
    ```  ');  ```
    ```    ```
    ```  Ouptut workbook from R  ```
    ```    ```
    ```    ```
    ```                                               COMPUTED  ```
    ```                A          B          C           D  ```
    ```         -----------------------------------------------  ```
    ```               NAME      HEIGHT    WEIGHT       BMI  ```
    ```    ```
    ```         1 |   Alfred     69.0      112.5    16.61153119   =C2*703/B2^2  ```
    ```         2 |   Alice      56.5       84.0    18.49855118   ...  ```
    ```         3 |   Barbara    65.3       98.0    16.15678844  ```
    ```         4 |   Carol      62.8      102.5    18.27089841  ```
    ```         5 |   Henry      63.5      102.5    17.87029574  ```
    ```         6 |   James      57.3       83.0    17.77150358  ```
    ```         7 |   Jane       59.8       84.5    16.61153119  ```
    ```         8 |   Janet      62.5      112.5    20.2464  ```
    ```         9 |   Jeffrey    62.5       84.0    15.117312  ```
    ```        10 |   John       59.0       99.5    20.09436943  ```
    ```        11 |   Joyce      51.3       50.5    13.49000072  ```
    ```        12 |   Judy       64.3       90.0    15.3029757  ```
    ```        13 |   Louise     56.3       77.0    17.0776953  ```
    ```        14 |   Mary       66.5      112.0    17.80451128  ```
    ```        15 |   Philip     72.0      150.0    20.34143519  ```
    ```        16 |   Robert     64.8      128.0    21.42966011  ```
    ```        17 |   Ronald     67.0      133.0    20.82846959  ```
    ```        18 |   Thomas     57.5       85.0    18.07334594  ```
    ```        19 |   William    66.5      112.0    17.80451128   =C20*703/B20^2  ```
    ```    ```
    ```                                   1900.5 =SUM($C$2:$C$20)  ```
    ```    ```
    ```    ```
