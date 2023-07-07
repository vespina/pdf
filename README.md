# pdf
VFP Helper class for report-2-pdf generation


#### OBJECTIVE
This library allows to take a normal VFP report and
generate a PDF representation of the report.

The library can work converting Postscript or XPS 
representations of the report. If autoSetup() is
called, library will try to auto configure both
converters and will auto select the first one
available, priorizing XPS mode.


#### TO INSTALL IN YOUR DEV MACHINE
* Download PDF.PRG and PDF.BIN from DIST folder
* Open VFP and run:
```
    COMPILE pdf
    DO pdf
    pdf.autoSetup()   // THIS WILL GENERATE GXPS.EXE UTILITY
```

#### FILES TO DISTRIBUTE 
* GXPS.EXE (only if you plan to use XPS mode)


##### USAGE
```
    * LOAD LIBRARY INTO MEMORY (PUT THIS IN YOUR MAIN PRG)
    DO pdf
    
    * CONFIGURE MODE (OR USER autoSetup METHOD)
    pdf.cMode = "XPS"   && OR "GS" FOR Ghostscript 
		
    * GENERATE PDF FILE FROM A REPORT FORMAT
    use datafile
    pdf.create("target.pdf", "report.frx", "datafile")
		
    * YOU CAN ALSO PASS REPORT FORM OPTIONS
    pdf.create("target.pdf", "report.frx", "datafile","optional-report-form-clauses")
```

##### OPTIONS
|Option|Description|
|------|-----------|
pdf.cMode|"GS" for Ghostscript, "XPS" for XPS
pdf.oGS.cPrinter|Name of a Postscript printer configured to FILE port.
pdf.oGS.cGSFolder|Location of GSDLL32.DLL [^1]
pdf.oXPS.cPrinter|XPS Printer name (default is "Microsoft XPS Document Writer")
pdf.oXPS.cGXPS|Location of GXPS.EXE (default is current folder) [^2]
pdf.cReportMode|Report engine mode [^3]

[^1]: This library has been tested using Ghostscript version 9.05.  Newer versions may not work.
[^2]: GXPS.EXE is a command line utility that converts XPS files into PDF files.  File will be automatically generated the first time you call autoSetup method.
[^3]: For VFP9 this defines if the PDF is generated using a LISTENER ("LISTENER") or LEGACY ("LEGACY") mode. Default value is "LEGACY".

