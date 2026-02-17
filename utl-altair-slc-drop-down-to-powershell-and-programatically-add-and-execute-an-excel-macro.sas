%let pgm=utl-altair-slc-drop-down-to-powershell-and-programatically-add-and-execute-an-excel-macro;

%stop_submission;

Altair slc drop down to powershell and programatically add and execute an excel macro

Problem: Add macro to existing excell workbook and execute macro to sum column E(weight)

Too long to post to a list, see github
https://github.com/rogerjdeangelis/utl-altair-slc-drop-down-to-powershell-and-programatically-add-and-execute-an-excel-macro

Process
    Add this macro

    Sub sum_weight()
        Range("E21").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
        Range("E22").Select
    End Sub


   d:/xls/class_final_out.xlsm

      +---------------------------------------+
      |     A   |  B  |  C  |   D    |   E    |
      +---------------------------------------+
   1  | NAME    | SEX | AGE | HEIGHT | WEIGHT |
      +---------+-----+-----+--------+--------+
   2  | ALFRED  |  M  | 14  |   69   | 112.5  |
      +---------+-----+-----+--------+--------+
       ...
      +---------+-----+-----+--------+--------+
   20 | WILLIAM |  M  | 15  |  66.5  | 112    |
      +---------+-----+-----+--------+--------+
   21 |         |     |     |        | 1900.9 | add and execute
      +---------+-----+-----+--------+--------+ vba macro sum_weight
      [CLASS]

PREP

Please enable VBA object model access:

I ran in admin mod, may not be needed?

1. Open Excel manually
2. Go to File ? Options ? Trust Center ? Trust Center Settings
3. Select Macro Settings
4. Check 'Trust access to the VBA project object model'
5  Check 'Enable macros'
6. Click OK and restart Excel

related repo
https://github.com/rogerjdeangelis/utl_programatically_execute_excel_macro_using_wps_proc_python

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

%utlfkil(d:/xls/class_final.xlsx);

libname xls excel "d:/xls/class_final.xlsx";

data xls.class;
informat
  NAME $8.
  SEX $1.
  AGE 8.
  HEIGHT 8.
  WEIGHT 8.
;
input
  NAME SEX AGE HEIGHT WEIGHT;
cards4;
Alfred M 14 69 112.5
Alice F 13 56.5 84
Barbara F 13 65.3 98
Carol F 14 62.8 102.5
Henry M 14 63.5 102.5
James M 12 57.3 83
Jane F 12 59.8 84.5
Janet F 15 62.5 112.5
Jeffrey M 13 62.5 84
John M 12 59 99.5
Joyce F 11 51.3 50.5
Judy F 14 64.3 90
Louise F 12 56.3 77
Mary F 15 66.5 112
Philip M 16 72 150
Robert M 12 64.8 128
Ronald M 15 67 133
Thomas M 11 57.5 85
William M 15 66.5 112
;;;;
run;quit;

libname xls clear;

/***************************************************************************************************************************/
/*  d:/xls/class_final.xlsx                      |                                                                         */
/*    +---------------------------------------+  |  Sub sum_weight()                                                       */
/*    |     A   |  B  |  C  |   D    |   E    |  |      Range("E21").Select                                                */
/*    +---------------------------------------+  |      ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"                    */
/* 1  | NAME    | SEX | AGE | HEIGHT | WEIGHT |  |      Range("E22").Select                                                */
/*    +---------+-----+-----+--------+--------+  |  End Sub                                                                */
/* 2  | ALFRED  |  M  | 14  |   69   | 112.5  |  |                                                                         */
/*    +---------+-----+-----+--------+--------+  |                                                                         */
/*     ...                                       |                                                                         */
/*    +---------+-----+-----+--------+--------+  |                                                                         */
/* 20 | WILLIAM |  M  | 15  |  66.5  | 112    |  |                                                                         */
/*    +---------+-----+-----+--------+--------+  |                                                                         */
/*    [CLASS]                                    |                                                                         */
/***************************************************************************************************************************/
/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

%utlfkil(D:\xls\class_final_out.xlsm);

%slc_psbegin;
cards4;
# Configuration
$inputPath = "D:\xls\class_final.xlsx"
$outputPath = "D:\xls\class_final_out.xlsm"

$vbaCode = @"
Sub sum_weight()
    Range("E21").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
    Range("E22").Select
End Sub
"@

try {
    Write-Host "Starting Excel automation..." -ForegroundColor Yellow

    # Create Excel application object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1  # Enable macros

    # Open the workbook
    Write-Host "Opening workbook: $inputPath" -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($inputPath)

    # Check if VBProject is accessible
    try {
        $vbProject = $workbook.VBProject
        if (-not $vbProject) {
            throw "VBProject is null. VBA access not enabled."
        }

        # Add macro
        Write-Host "Adding macro to workbook..." -ForegroundColor Yellow
        $module = $vbProject.VBComponents.Add(1)
        $module.Name = "WeightModule"
        $module.CodeModule.AddFromString($vbaCode)

        # Save as macro-enabled
        Write-Host "Saving as macro-enabled workbook..." -ForegroundColor Yellow
        $workbook.SaveAs($outputPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled

        # Run the macro
        Write-Host "Running sum_weight macro..." -ForegroundColor Yellow
        $excel.Run("sum_weight")

        # Get the result
        $result = $workbook.Worksheets(1).Range("E21").Value
        Write-Host "Sum in E21: $result" -ForegroundColor Green

        # Save and close
        $workbook.Save()
        $workbook.Close()

        Write-Host "? Macro added and executed successfully!" -ForegroundColor Green
        Write-Host "File saved as: $outputPath" -ForegroundColor Green
    }
    catch {
        Write-Host "? Error accessing VBA project: $_" -ForegroundColor Red
        Write-Host "`nPlease enable VBA object model access:" -ForegroundColor Yellow
        Write-Host "1. Open Excel manually" -ForegroundColor Yellow
        Write-Host "2. Go to File ? Options ? Trust Center ? Trust Center Settings" -ForegroundColor Yellow
        Write-Host "3. Select Macro Settings" -ForegroundColor Yellow
        Write-Host "4. Check 'Trust access to the VBA project object model'" -ForegroundColor Yellow
        Write-Host "5. Click OK and restart Excel" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "? Error: $_" -ForegroundColor Red
}
finally {
    # Cleanup
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
;;;;
%slc_psend;

/**************************************************************************************************************************/
/*  d:/xls/class_final_out.xlsm                                                                                           */
/*    +---------------------------------------+                                                                           */
/*    |     A   |  B  |  C  |   D    |   E    |                                                                           */
/*    +---------------------------------------+                                                                           */
/* 1  | NAME    | SEX | AGE | HEIGHT | WEIGHT |                                                                           */
/*    +---------+-----+-----+--------+--------+                                                                           */
/* 2  | ALFRED  |  M  | 14  |   69   | 112.5  |                                                                           */
/*    +---------+-----+-----+--------+--------+                                                                           */
/*     ...                                                                                                                */
/*    +---------+-----+-----+--------+--------+                                                                           */
/* 20 | WILLIAM |  M  | 15  |  66.5  | 112    |                                                                           */
/*    +---------+-----+-----+--------+--------+                                                                           */
/* 21 |         |     |     |        | 1900.9 | add and execute                                                           */
/*    +---------+-----+-----+--------+--------+ vba macro sum_weight                                                      */
/*    [CLASS]                                                                                                             */
/**************************************************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

Key section of log

Starting Excel automation...
Opening workbook: D:\xls\class_final.xlsx
Adding macro to workbook...
Saving as macro-enabled workbook...
Running sum_weight macro...
Sum in E21: Variant Value (Variant) {get} {set}
? Macro added and executed successfully!
File saved as: D:\xls\class_final_out.xlsm


1                                          Altair SLC      11:02 Tuesday, February 17, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  11:02:22
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.047
      cpu time  : 0.015


NOTE: AUTOEXEC processing completed

1         %slc_psbegin;
2         cards4;

NOTE: The file 'c:\temp\ps_pgm.ps1' is:
      Filename='c:\temp\ps_pgm.ps1',
      Owner Name=BUILTIN\Administrators,
      File size (bytes)=0,
      Create Time=13:37:05 Jul 16 2025,
      Last Accessed=11:02:22 Feb 17 2026,
      Last Modified=11:02:22 Feb 17 2026,
      Lrecl=32767, Recfm=V

NOTE: 79 records were written to file 'c:\temp\ps_pgm.ps1'
      The minimum record length was 80
      The maximum record length was 107
NOTE: The data step took :
      real time : 0.000
      cpu time  : 0.000


3         # Configuration
4         $inputPath = "D:\xls\class_final.xlsx"
5         $outputPath = "D:\xls\class_final_out.xlsm"
6
7         $vbaCode = @"
8         Sub sum_weight()
9             Range("E21").Select
10            ActiveCell.FormulaR1C1 = "=SUM(R[-19]C:R[-1]C)"
11            Range("E22").Select
12        End Sub
13        "@
14
15        try {
16            Write-Host "Starting Excel automation..." -ForegroundColor Yellow
17
18            # Create Excel application object
19            $excel = New-Object -ComObject Excel.Application
20            $excel.Visible = $false
21            $excel.DisplayAlerts = $false
22            $excel.AutomationSecurity = 1  # Enable macros
23
24            # Open the workbook
25            Write-Host "Opening workbook: $inputPath" -ForegroundColor Yellow
26            $workbook = $excel.Workbooks.Open($inputPath)
27
28            # Check if VBProject is accessible
29            try {
30                $vbProject = $workbook.VBProject
31                if (-not $vbProject) {
32                    throw "VBProject is null. VBA access not enabled."
33                }
34
35                # Add macro
36                Write-Host "Adding macro to workbook..." -ForegroundColor Yellow
37                $module = $vbProject.VBComponents.Add(1)
38                $module.Name = "WeightModule"
39                $module.CodeModule.AddFromString($vbaCode)
40
41                # Save as macro-enabled
42                Write-Host "Saving as macro-enabled workbook..." -ForegroundColor Yellow
43                $workbook.SaveAs($outputPath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
44
45                # Run the macro
46                Write-Host "Running sum_weight macro..." -ForegroundColor Yellow
47                $excel.Run("sum_weight")
48
49                # Get the result
50                $result = $workbook.Worksheets(1).Range("E21").Value
51                Write-Host "Sum in E21: $result" -ForegroundColor Green
52
53                # Save and close
54                $workbook.Save()
55                $workbook.Close()
56
57                Write-Host "? Macro added and executed successfully!" -ForegroundColor Green
58                Write-Host "File saved as: $outputPath" -ForegroundColor Green
59            }
60            catch {
61                Write-Host "? Error accessing VBA project: $_" -ForegroundColor Red
62                Write-Host "`nPlease enable VBA object model access:" -ForegroundColor Yellow
63                Write-Host "1. Open Excel manually" -ForegroundColor Yellow
64                Write-Host "2. Go to File ? Options ? Trust Center ? Trust Center Settings" -ForegroundColor Yellow
65                Write-Host "3. Select Macro Settings" -ForegroundColor Yellow
66                Write-Host "4. Check 'Trust access to the VBA project object model'" -ForegroundColor Yellow
67                Write-Host "5. Click OK and restart Excel" -ForegroundColor Yellow
68            }
69        }
70        catch {
71            Write-Host "? Error: $_" -ForegroundColor Red
72        }
73        finally {
74            # Cleanup
75            if ($excel) {
76                $excel.Quit()
77                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
78            }
79            [System.GC]::Collect()
80            [System.GC]::WaitForPendingFinalizers()
81        }
82        ;;;;
83        %slc_psend;

NOTE: The infile rut is:
      Unnamed Pipe Access Device,
      Process=powershell.exe -executionpolicy bypass -file c:/temp/ps_pgm.ps1 >  c:/temp/ps_pgm.log,
      Lrecl=32756, Recfm=V

NOTE: No records were written to file PRINT

NOTE: No records were read from file rut
NOTE: The data step took :
      real time : 7.880
      cpu time  : 0.015



NOTE: The infile rut is:
      Unnamed Pipe Access Device,
      Process=powershell.exe -executionpolicy bypass -file c:/temp/ps_pgm.ps1 >  c:/temp/ps_pgm.log,
      Lrecl=32767, Recfm=V

NOTE: No records were written to file PRINT

NOTE: No records were read from file rut
NOTE: The data step took :
      real time : 7.073
      cpu time  : 0.000



NOTE: The infile 'c:\temp\ps_pgm.log' is:
      Filename='c:\temp\ps_pgm.log',
      Owner Name=BUILTIN\Administrators,
      File size (bytes)=296,
      Create Time=09:57:45 Feb 16 2026,
      Last Accessed=11:02:36 Feb 17 2026,
      Last Modified=11:02:36 Feb 17 2026,
      Lrecl=32767, Recfm=V

Starting Excel automation...
Opening workbook: D:\xls\class_final.xlsx
Adding macro to workbook...
Saving as macro-enabled workbook...
Running sum_weight macro...
Sum in E21: Variant Value (Variant) {get} {set}
? Macro added and executed successfully!
File saved as: D:\xls\class_final_out.xlsm

NOTE: 8 records were read from file 'c:\temp\ps_pgm.log'
      The minimum record length was 27
      The maximum record length was 48
NOTE: The data step took :
      real time : 0.000
      cpu time  : 0.000


84
85
ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 15.128
      cpu time  : 0.078
/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
