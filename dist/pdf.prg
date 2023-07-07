* PDF.PRG
* Library for PDF printing in VFP
*
* Author: Victor Espina (mailto: vespinas@gmail.com)
*
*
* This library allows to take a normal VFP report and
* generate a PDF representation of the report.
*
* The library can work converting Postscript or XPS 
* representations of the report. If autoSetup() is
* called, library will try to auto configure both
* converters and will auto select the first one
* available, priorizing XPS mode.
*
*
* USAGE:
* DO pdf
* IF !pdf.autoSetup()
*  * Manual configuration is required
* ENDIF
* pdf.create("target.pdf", "report.frx", "alias" [,"optional-report-form-clauses"])
*
*
* OPTIONS:
* pdf.cMode:         "GS" for Ghostscript, "XPS" for XPS
* pdf.oGS.cPrinter:  PS Printer
* pdf.oGS.cGSFolder: Location of GSDLL32.DLL
* pdf.oXPS.cPrinter: XPS Printer
* pdf.oXPS.cGXPS:    Location of GXPS.EXE
*
*
* VERSION HISTORY
* Mar 23, 2022  1.10    VES     Improvements when reporting that selected mode is not available
*
* Aug 18, 2020  1.9     VES     Improvements on error handling.
*
* Mar 11, 2020  1.8     VES     Added a 2s idle time after calling Generate() to allow the intermediate file to be fully created, thus 
*                               avoiding innecesary "file not found" errors on Convert() method calls.
*
* Mar 9, 2016   1.7     VES		General changes to add support for VFP versions prior to 9
*
* Jan 1, 2016   1.6		VES		General changes to make the library compatible under ActiveVFP
*
* May 29, 2015  1.5     VES     General changes to include support for password-protected PDFs
*                               New property oOptions for futher PDF customizing.
*                               Starting with this version, programmer can add GS options using oOptions.oGSOptions object
* May 15, 2015  1.4		VES		Changes in Create() method of pdf class and Generate() method of pdfConverter to
*                               allow optional REPORT FORM options
*
* May 5, 2015	1.3		VES		Changes in autoSetup() method of pdfGSConverter to speed the GSDLL32.DLL discovery
*								New lCheckFRX in pdf class
*								Optional FRX checking
*
* May 4, 2015   1.2		VES		Changes un Generate() method of pdfConverter class to support long name output files
*
* May 2, 2015	1.2		VES		Changes in autoSetup() methods to use cError property to store
*								a description of why an specific converter is not available. New
*                               setter for lDebug property on pdf class.
*
* 								New cCapabilities property on pdf class
*								New isModeAvailable method on pdf class
*								
* May 1, 2015	1.1		VES		Bug fixes
*
* May 1, 2015	1.0		VES		Initial version
*
#DEFINE true 	.T.
#DEFINE false	.F.


#IF VERSION(5) >= 800
	#DEFINE _TRY		TRY
	#DEFINE _CATCH	    CATCH TO ex
	#DEFINE _ENDTRY		ENDTRY
	#DEFINE _EMPTY		"EMPTY"
#ELSE
	#DEFINE _TRY		ex=TRY()
	#DEFINE _CATCH	    IF CATCH(@ex)
	#DEFINE _ENDTRY		ENDIF
	#DEFINE _EMPTY		"EMPTYOBJECT"
#ENDIF		




* Loader
SET PROCEDURE TO pdf ADDITIVE
PUBLIC pdf
pdf = CREATEOBJECT("pdf")

RETURN


* pdf (Class)
* Main helper class
*
DEFINE CLASS pdf AS Custom

 * Public properties
 cMode = "XPS"       && Operation mode (XPS, GS)
 oXPS = NULL         && XPS helper
 oGS = NULL          && Ghostscript Helper 
 lError = false      && Error flag
 cError = ""         && Last error description
 lDebug = false      && Debug mode
 nVersion = 1.10     && Version
 cCapabilities = ""  && Available capabilities
 lCheckFRX = true    && Check FRX for potentials incompatibilities with PDF printing
 oOptions = NULL     && Opciones
 oFSO = NULL         && Instance of Scripting.FileSystemObject
 cReportMode = ""    && Report mode used (LISTENER or LEGACY)
 nDSID = 1           && Current datasession (set/get)
 
 * R/O Properties setters
 PROCEDURE nVersion_Assign(vNewVal) 
 
 * Other properties setters
 PROCEDURE lDebug_Assign(vNewVal)
  THIS.oGS.lDebug = vNewVal
  THIS.oXPS.lDebug = vNewVal
 ENDPROC
 
 PROCEDURE nDSID_Assign(vNewVal)
  SET DATASESSION TO (m.vNewVal)
 ENDPROC
 PROCEDURE nDSID_Access
  RETURN SET("DATASESSION")
 ENDPROC
 
 PROCEDURE cCapabilities_Assign(vNewVal)
 PROCEDURE cCapabilities_Access
  LOCAL cList
  cList = ""
  IF THIS.oXPS.isAvailable()
   cList = " XPS"
  ENDIF
  IF THIS.oGS.isAvailable()
   cList = cList  + ",GS"
  ENDIF
  cList = EVL(SUBSTR(cList,2),"none")
  RETURN cList
 ENDPROC
  
 
 * Constructor
 PROCEDURE Init(plAutoSetup)
  THIS.apiDeclare()
  THIS.oOptions = CREATEOBJECT("pdfOptions")
  THIS.oXPS = CREATEOBJECT("pdfXPSHelper")
  THIS.oGS = CREATEOBJECT("pdfGSHelper")
  THIS.oFSO = CREATEOBJECT("Scripting.FileSystemObject")
  IF plAutoSetup
   THIS.autoSetup()
  ENDIF
 ENDPROC


 * apiDeclare (Method)
 * Declare any common API functions
 *
 PROCEDURE apiDeclare
  DECLARE SHGetSpecialFolderPath IN SHELL32.DLL ;
		LONG hwndOwner, ;
		STRING @cSpecialFolderPath, ;
		LONG nWhichFolder 
		
  DECLARE Sleep IN kernel32 ;
        INTEGER dwMilliseconds			
 ENDPROC 


 * autoSetup (Method)
 * Try to auto-configure available converters
 *
 PROCEDURE autoSetup
  THIS.clearError()
  THIS.oGS.autoSetup()
  
  THIS.clearError()
  THIS.oXPS.autoSetup()  
  
  IF THIS.cMode = "XPS" AND !THIS.oXPS.isAvailable() AND THIS.oGS.isAvailable()
   THIS.cMode = "GS"
  ELSE
   IF THIS.cMode = "GS" AND THIS.oXPS.isAvailable() AND !THIS.oGS.isAvailable()
    THIS.cMode = "XPS"
   ENDIF
  ENDIF
  
  RETURN THIS.isModeAvailable("GS") OR THIS.isModeAvailable("XPS")
 ENDPROC
 
 
 
 * isModeAvailable (Method)
 * Check if a given mode is available. If no mode is passed, the current mode will be used
 *
 PROCEDURE isModeAvailable(pcMode)  
  THIS.clearError()
  pcMode = UPPER(EVL(pcMode, THIS.cMode))
  DO CASE
     CASE pcMode == "GS"
          THIS.lError = !THIS.oGS.isAvailable()
          THIS.cError = THIS.oGS.cError
          
     CASE pcMode == "XPS"
          THIS.lError = !THIS.oXPS.isAvailable()
          THIS.cError = THIS.oXPS.cError
  ENDCASE
  RETURN !THIS.lError
 ENDPROC
 
 

 * Create (Method)
 * Creates a PDF file base on the given report format
 *
 PROCEDURE Create(pcOUtput, pcFRX, pcAliasOrDBF, pcReportOptions)
  THIS.clearError()
  IF THIS.lCheckFRX
   IF NOT THIS.CheckFRX(pcFRX)
    RETURN false
   ENDIF
  ENDIF
  IF THIS.lError
   RETURN False
  ENDIF

  _TRY 	 
	  LOCAL lAutoClose, nWkArea
	  lAutoClose = false
	  IF !EMPTY(pcAliasOrDBF) 
	   DO CASE
	      CASE !FILE(pcAliasOrDBF) AND USED(pcAliasOrDBF)
	           SELECT (pcAliasOrDBF)
	           
	      CASE FILE(pcAliasOrDBF)
	           SELECT 0
	           USE (pcAliasOrDBF)
	           lAutoClose = true
	           nWkArea = SELECT()
	   ENDCASE
	  ENDIF
	  
	  IF NOEX() AND EMPTY(ALIAS())
	   THIS.cError = "No datasource is available"
	   THIS.lError = true
	  ENDIF

  _CATCH
     THIS.cError = ex.Message + " at " + ex.Procedure + " (" + ALLTRIM(STR(ex.lineNo)) +  ")"
     THIS.lError = .T.	
  _ENDTRY
  IF THIS.lError
   #IF VERSION(5) < 700
   IF !EMPTY(gcTRYOnError)
   	 ON ERROR &gcTRYOnError
   ENDIF
   #ENDIF
   RETURN .F.
  ENDIF	  

	  
  * Get the appropiate helper
  LOCAL cInput, oHelper
  IF EMPTY(JUSTPATH(pcOutput))
   pcOutput = FULLPATH(pcOutput)
  ENDIF
  DO CASE
     CASE THIS.cMode == "XPS"
          cInput = FORCEEXT(pcOutput, "XPS")
          oHelper = THIS.oXPS
     
     CASE THIS.cMode == "GS"
          cInput = FORCEEXT(pcOutput, "PS")
          oHelper = THIS.oGS
     
     OTHERWISE
          THIS.cError = "Invalid mode (XPS / GS)"
          THIS.lError = true
  ENDCASE
  IF THIS.lError
   RETURN False
  ENDIF
  oHelper.oOptions = THIS.oOptions
  oHelper.nDSID = THIS.nDSID
  	  
  * Check if the selected generator is available
  IF !THIS.lError AND !oHelper.isAvailable()
   THIS.cError = "Selected mode is not available at this time. Make sure your configurations is valid (" + oHelper.cError + ")"
   THIS.lERror = true
   #IF VERSION(5) < 700
   IF !EMPTY(gcTRYOnError)
   	 ON ERROR &gcTRYOnError
   ENDIF
   #ENDIF   
   RETURN .F.
  ENDIF
  
  * Run the report and generate the input file  
  IF !THIS.lError
   LOCAL lGenerated
   lGenerated = oHelper.Generate(pcFRX, cInput, pcReportOptions)
   IF NOT lGenerated
    THIS.cError = "[generate] " + oHelper.cError
    THIS.lError = true
    RETURN .F.
   ENDIF
   THIS.clearError()
   Sleep(2000)  && Gives some time to allow intermediate file to be fully generated
  ENDIF
  THIS.cReportMode = oHelper.cReportMode
	  
	  
  * Convert the file to PDF
  IF NOT ["'] $ cInput
   cInput = ["] + cInput + ["]
  ENDIF
  IF NOT ["'] $ pcOutput
   pcOutput = ["] + pcOutput + ["]
  ENDIF
  _TRY
	  IF !oHelper.Convert(cInput, pcOutput)
	   IF NOEX()
 	    THIS.cError = "[convert] " + oHelper.cError
 	    THIS.lError = true
 	   ENDIF
	  ENDIF

	  IF NOEX() AND !THIS.lError AND oHelper.isFile(cInput)
	   _TRY
	       THIS.oFSO.deleteFile( CHRT(cInput,["],"") )
	   _CATCH
	   _ENDTRY
	  ENDIF
	  
	  IF NOEX() AND lAutoClose
	   USE IN (nWkArea)
	  ENDIF

  _CATCH
     THIS.cError = ex.Message + " at " + ex.Procedure + " (" + ALLTRIM(STR(ex.lineNo)) +  ")"
     THIS.lError = .T.	
  _ENDTRY
  	  
   #IF VERSION(5) < 700
   IF !EMPTY(gcTRYOnError)
   	 ON ERROR &gcTRYOnError
   ENDIF
   #ENDIF
     	  
  RETURN !THIS.lError
 ENDPROC
 


 #IF VERSION(5) >= 900
 PROCEDURE Error(nError, cMethod, nLine)
  THIS.lError = .T.
  THIS.cError = MESSAGE() + " at " + cMethod + "(" + ALLTRIM(STR(nLine)) + ")"
 ENDPROC
 #ENDIF
  

 PROCEDURE clearError
  THIS.lError = False
  THIS.cError = ""
 ENDPROC

 
 * checkFRX (Method)
 * Check FRX file for potential problems
 *
 PROCEDURE checkFRX(pcFRX)
  THIS.clearError()
  LOCAL lResult,nWkArea
  nWkArea = SELECT()
  lResult = True
  _TRY
   SELECT 0
   USE (pcFRX) SHARED AGAIN
   IF NOEX()
    DO CASE
       CASE ATC("DEVICE=", expr) <> 0
            THIS.lError = true
            THIS.cError = "Report contains printer environment. Please remove it and try again"
    ENDCASE
    USE
   ENDIF
  _CATCH
  _ENDTRY
  SELECT (nWkArea)
  RETURN lResult
 ENDPROC


ENDDEFINE
 



 
* pdfConverter (Abstract class)
* Abstract class for convert helpers
*
DEFINE CLASS pdfConverter AS Custom
 lError = false
 cError = ""
 cPrinter = ""
 lDebug = false
 cResolution = ""    && Se mantiene por compatibilidad hacia atras. Usar THIS.oOptions.cResolution
 oOptions = NULL
 cReportMode = ""
 nDSID = 1           && Current datasession (get/set)

 PROCEDURE cResolution_Assign(vNewVal)
   THIS.oOptions.cResolution = vNewVal
 ENDPROC
 PROCEDURE cResolution_Access
   RETURN THIS.oOptions.cResolution
 ENDPROC
 PROCEDURE nDSID_Assign(vNewVal)
  SET DATASESSION TO (m.vNewVal)
 ENDPROC
 PROCEDURE nDSID_Access
  RETURN SET("DATASESSION")
 ENDPROC
 
 PROCEDURE isAvailable()
 PROCEDURE autoSetup()
 PROCEDURE Generate(pcFRX, pcOutput, pcReportOptions)
 PROCEDURE Convert(pcInput, pcOutput)
 
 PROCEDURE Init 
  THIS.oOptions = CREATEOBJECT("pdfOptions")
 ENDPROC

 * Generate (Method)
 * Takes a VFP report format and generates a file output
 *
 PROCEDURE Generate(pcFRX, pcOutput, pcReportOptions)
  LOCAL lReturn 
  lReturn =false
  
  IF THIS.IsFile(pcOutput)
   _TRY
    THIS.deleteFile(pcOutput)
   _CATCH
   _ENDTRY
  ENDIF
 
  IF !THIS.isFile(pcFRX)
   THIS.setError("[generate] File " + pcFRX + " does not exists")
   RETURN false
  ENDIF
 
  pcOutput = THIS.fixFileName(pcOutput)
  pcFRX = THIS.fixFileNAme(pcFRX)
    
  IF EMPTY(ALIAS())
   THIS.setError("[generate] No datasource is available")
   RETURN false
  ENDIF
  
  IF VARTYPE(_REPORTOUTPUT)="C" AND THIS.IsFile(_REPORTOUTPUT)
   lReturn = THIS.listenerMode(pcFRX, pcOutput, pcReportOptions)  
   THIS.cReportMode = "LISTENER"
  ELSE
   lReturn = THIS.legacyMode(pcFRX, pcOutput, pcReportOptions)
   THIS.cReportMode = "LEGACY"
  ENDIF
  
  RETURN lReturn
 ENDPROC
 
 PROTECTED PROCEDURE listenerMode(pcFRX, pcOutput, pcReportOptions)
	IF VARTYPE(pcReportOptions)<>"C"
	 pcReportOptions = ""
	ENDIF 
	PRIVATE oListener AS ReportListener		 	
 	oListener = CREATEOBJECT('ReportListener')
	oListener.AllowModalMessages = false
  	oListener.DynamicLineHeight = true
	oListener.ListenerType = 0
	oListener.QuietMode = true
	LOCAL cCode	
	cCode = [REPORT FORM ] + pcFRX + [ ] + pcReportOptions + " OBJECT oListener TO FILE " + pcOutput				
	_TRY
 	 #IF VERSION(5) >= 900
 	  SET ENGINEBEHAVIOR 90	
 	 #ENDIF
 	 SET PRINTER TO NAME (THIS.cPrinter) 
 	 #IF VERSION(5) >= 900
	  EXECSCRIPT(cCode)
	 #ELSE
	  &cCode
	 #ENDIF
	 SET PRINTER TO
	 IF NOEX() AND !FILE(pcOutput)   &&!pdf.oFSO.fileExists(pcOutput)
	  THIS.setError("The XPS file could not be generated by printer " + THIS.cPrinter + ". No error is available at this time (_REPORTOUTPUT=" + _REPORTOUTPUT + ")")
	 ENDIF
	_CATCH
	 THIS.setError("Report error: " + ex.Message)
	_ENDTRY 	
	RETURN !THIS.lError
 ENDPROC
 
 PROTECTED PROCEDURE legacyMode(pcFRX, pcOutput, pcReportOptions)
  IF VARTYPE(pcReportOptions)<>"C"
   pcReportOptions = "FOR 1=1"
  ENDIF  
  _TRY
   SET PRINTER TO NAME (THIS.cPrinter) 
   REPORT FORM (pcFRX) &pcReportOptions NOCONSOLE TO FILE (pcOutput)
   SET PRINTER TO
  _CATCH
   IF !INLIST(ex.errorNo, 1958)
    THIS.setError("Report error: " + ex.Message)
   ENDIF
  _ENDTRY
  RETURN !THIS.lError
 ENDPROC

 * unPack (Method)
 * Unpack a distributed file
 *
 PROCEDURE unPack(pcFile, pcTarget)
  LOCAL oHelper
  oHelper = CREATEOBJECT("base64Helper")
  oHelper.decodeFile(FILETOSTR(pcFile), pcTarget)
 ENDPROC
 
 
 * isFile (Method)
 * Check if a file exists.  We use FileSystemObject because
 * FILE() fails in hosted environments.
 *
 PROCEDURE isFile(pcFile)
  LOCAL oFSO,cFile
  oFSO = CREATEOBJECT("Scripting.FileSystemObject")
  cFile = CHRTRAN(pcFile,["'],"")
  RETURN FILE(cFile) OR oFSO.FileExists(pcFile)
 ENDPROC
 

 * isFolder (Method)
 * Check if a folder exists.  We use FileSystemObject because
 * DIRECTORY() fails in hosted environments.
 *
 PROCEDURE isFolder(pcFolder)
  LOCAL oFSO
  oFSO = CREATEOBJECT("Scripting.FileSystemObject")
  RETURN oFSO.FolderExists(ADDBS(pcFolder))
 ENDPROC
 
 
 * deleteFile(Method)
 * Deletes a file. We use FileSystemObject because
 * REMOVE fails with quoted file names.
 *
 PROCEDURE deleteFile(pcFile)
   LOCAL oFSO
   oFSO = CREATEOBJECT("Scripting.FileSystemObject")
   oFSO.DeleteFile(pcFile)
 ENDPROC 
 
 
 * fixFileName (MEthod)
 * Fix special file names
 *
 PROCEDURE fixFileName(pcFile)
  IF AT([ ], pcFile) > 0 AND !(["] $ pcFile OR ['] $ pcFile)
   pcFile = ["] + pcFile + ["]
  ENDIF
  RETURN pcFile
 ENDPROC
 
 
 * setError (Method)
 * Activate error flag with an optional description
 *
 PROCEDURE setError(pcError)
  THIS.lError = true
  THIS.cError = pcError
 ENDPROC


 * clearError (Method)
 * Clear error flag and description
 *
 PROCEDURE clearError
  THIS.lError = False
  THIS.cError = ""
 ENDPROC
  
ENDDEFINE






* pdfGSHelper (Class)
* Ghostscript helper
*
#DEFINE CSIDL_PROGRAM_FILES 0x0026
DEFINE CLASS pdfGSHelper AS pdfConverter
 cGSFolder = ""
 oGSUtils = NULL
 
 PROCEDURE Init
  DODEFAULT()
  THIS.oGSUtils = CREATEOBJECT("pdfGSUtils")
 ENDPROC
  
 
 * autoSetup (Method)
 * Check environment to try to find GS components
 *
 PROCEDURE autoSetup
  IF EMPTY(THIS.cGSFolder)
   LOCAL cPrograms,nFound
   LOCAL ARRAY aFound[1]
   cPrograms = SPACE(255)
   SHGetSpecialFolderPath(0, @cPrograms, CSIDL_PROGRAM_FILES)
   IF THIS.lDebug
    WAIT "Finding gsdll32.dll..." WINDOW NOWAIT
   ENDIF
   nFound = findFolders(@aFound, cPrograms, "gs*", 0)
   IF nFound = 0 
    nFound = findFolders(@aFound, cPrograms, "pdfcreator", 0)
   ENDIF
   IF nFound > 0
    cPrograms = aFound[1]
    DIMENSION aFound[1]
    aFound[1] = false
    nFound = 0 
   ENDIF
   nFound = FindFiles(@aFound, cPrograms, "gsdll32.dll")
   IF nFound > 0
    THIS.cGSFolder = JUSTPATH(aFound[1])
   ENDIF
   IF THIS.lDebug
    WAIT "Found at: " + THIS.cGSFolder WINDOW 
   ENDIF
  ENDIF
  
  IF EMPTY(THIS.cPrinter) 
   THIS.cPrinter = findPrinter("driver","PDFCreator")
   IF EMPTY(THIS.cPrinter)
    THIS.cPrinter = findPrinter("name","*PDF*")
   ENDIF
   IF EMPTY(THIS.cPrinter)
    THIS.cPrinter = findPrinter("name","* PDF *")
   ENDIF
   IF EMPTY(THIS.cPrinter)
    THIS.cPrinter = findPrinter("driver","* PS")
   ENDIF  
  ENDIF
 ENDPROC
 
 
 * isAvailable (Method)
 * Check if GS converting is available
 *
 PROCEDURE isAvailable
  DO CASE
     CASE EMPTY(THIS.cGSFolder)
          THIS.setError("Property oGS.cGSFolder is empty")
          
     CASE !THIS.isFolder(THIS.cGSFolder)
          THIS.setError("Folder " + THIS.cGSFolder + " does not exists")
          
     CASE !THIS.isFile(ADDBS(THIS.cGSFolder) + "gsdll32.dll")
          THIS.setError("File gsdll32.dll not found at " + THIS.cGSFolder)
          
     CASE EMPTY(THIS.cPrinter)
          THIS.setError("Property oGS.cPrinter is empty")
  ENDCASE        
  RETURN !THIS.lError
 ENDPROC
 
 
 * Convert (Method)
 * Ghostscript convertion. This code was adapted from Print2PDF library by
 * Paul James (Life-Cycle Technologies, Inc.) mailto:paulj@lifecyc.com
 *
 PROCEDURE Convert(pcInput, pcOutput)
	local lnGSInstanceHandle, lnCallerHandle, loHeap, lnElementCount, lcPtrArgs, lnCounter, lnReturn

	store 0 to lnGSInstanceHandle, lnCallerHandle, lnElementCount, lnCounter, lnReturn
	store null to loHeap
	store "" to lcPtrArgs
	
	* If GSFolder is not in PATH, add it
    IF ATC(THIS.cGSFolder, SET("PATH")) = 0
     LOCAL cPath
     cPath = ALLTRIM(SET("PATH"))
     SET PATH TO cPath + ";" + THIS.cGSFolder
    ENDIF

	set safety off
	loHeap = createobject('Heap')

	**Declare Ghostscript DLLs
	declare long gsapi_new_instance in gsdll32.dll ;
		long @lngGSInstance, long lngCallerHandle
	declare long gsapi_delete_instance in gsdll32.dll ;
		long lngGSInstance
	declare long gsapi_init_with_args in gsdll32.dll ;
		long lngGSInstance, long lngArgumentCount, ;
		long lngArguments
	declare long gsapi_exit in gsdll32.dll ;
		long lngGSInstance

    THIS.oGSUtils.oOptions = THIS.oOptions
    
    pcInput = THIS.fixFileName(pcInput)
    pcOutput = THIS.fixFileName(pcOutput)

    LOCAL oArgs
    oArgs = CREATEOBJECT("Collection")
    oArgs.Add("dummy")
    oArgs = THIS.oGSUtils.getConvertArgs(oArgs)
	oArgs.Add("-sOutputFile=" + pcOutput)	&&Name of the output file
	oArgs.Add("-c")							&&Interprets arguments as PostScript code up to the next argument that begins with "-" followed by a non-digit, or with "@". For example, if the file quit.ps contains just the word "quit", then -c quit on the command line is equivalent to quit.ps there. Each argument must be exactly one token, as defined by the token operator
	oArgs.Add(".setpdfwrite")				&&If this file exists, it uses it as command-line input?
	oArgs.Add("-f")							&&(ends the -c argument started in laArgs[8])	
	oArgs.Add(pcInput)						&&Input File name (.ps file)
        
	dimension  laArgs[oArgs.Count]
	cArgs = ""
	FOR i = 1 TO oArgs.Count
	 laArgs[i] = oArgs.Item[i]
	 cArgs = cArgs + " " + oArgs.Item[i]
	ENDFOR
	*MESSAGEBOX(cArgs)
	
	
	* Load Ghostscript and get the instance handle
	lnReturn = gsapi_new_instance(@lnGSInstanceHandle, @lnCallerHandle)
	if (lnReturn < 0)
		loHeap = null
		RELEASE loHeap
		this.lError = true
		this.cError = "Could not start Ghostscript. (" + ALLTRIM(TRANSFORM(lnReturn,"")) + ")"
		return false
	endif

	* Convert the strings to null terminated ANSI byte arrays
	* then get pointers to the byte arrays.
	lnElementCount = alen(laArgs)
	lcPtrArgs = ""
	for lnCounter = 1 to lnElementCount
		lcPtrArgs = lcPtrArgs + NumToLONG(loHeap.AllocString(laArgs[lnCounter]))
	endfor
	lnPtr = loHeap.AllocBlob(lcPtrArgs)

	lnReturn = gsapi_init_with_args(lnGSInstanceHandle, lnElementCount, lnPtr)
	if (lnReturn < 0)
   	    gsapi_exit(lnGSInstanceHandle)
		loHeap = null
		RELEASE loHeap
		this.lError = true
		this.cError = "Could not Initilize Ghostscript. (" + ALLTRIM(TRANSFORM(lnReturn,"")) + ")"
		return false
	endif

	* Stop the Ghostscript interpreter
	lnReturn=gsapi_exit(lnGSInstanceHandle)
	if (lnReturn < 0)
		loHeap = null
		RELEASE loHeap
		this.lError = true
		this.cError = "Could not Exit Ghostscript."
		return false
	endif


	* release the Ghostscript instance handle'
	=gsapi_delete_instance(lnGSInstanceHandle)

	loHeap = null
	RELEASE loHeap

	if !THIS.IsFile(pcOutput)
		this.lError = true
		this.cError = "Ghostscript could not create the PDF."
		return false
	endif
 ENDPROC
 


ENDDEFINE






* pdfXPSHelper (Class)
* XPS helper
*
DEFINE CLASS pdfXPSHelper AS pdfConverter
 cGXPS = ""
 oGSUtils = NULL
 
 
 PROCEDURE Init
  DODEFAULT()
  THIS.oGSUtils = CREATEOBJECT("pdfGSUtils")
  THIS.cGXPS = FULLPATH("gxps.exe")
  THIS.cPRinter = "Microsoft XPS Document Writer"
 ENDPROC
 
 
 * autoSetup (Method)
 * Check for required dependencies
 PROCEDURE autoSetup
  IF EMPTY(THIS.cGXPS) OR !THIS.IsFile(THIS.cGXPS)
   LOCAL ARRAY aFiles[1]
   LOCAL nFound
   IF THIS.lDebug
    WAIT "Finding gxps.exe..." WINDOW NOWAIT
   ENDIF
   nFound = findFiles(@aFiles, CURDIR(), "gxps.exe", 1)

   IF nFound = 0 
    LOCAL cPrograms
    cPrograms = SPACE(255)
    SHGetSpecialFolderPath(0, @cPrograms, CSIDL_PROGRAM_FILES)   
    nFound = findFiles(@afiles, cPrograms, "gxps.exe", 1)
   ENDIF

   IF nFound = 0 AND THIS.isFile("pdf.bin")
    IF THIS.lDebug
     WAIT "Not found...unpacking...." WINDOW NOWAIT
    ENDIF
    THIS.unPack("pdf.bin", "gxps.exe")
    nFound = findFiles(@aFiles, CURDIR(), "gxps.exe", 1)
   ENDIF  
   IF THIS.lDebug
    WAIT CLEAR
   ENDIF

   IF nFound > 0
    THIS.cGXPS = aFiles[1]
   ENDIF
   IF THIS.lDebug
    WAIT "Found at " + THIS.cGXPS WINDOW
   ENDIF
  ENDIF
  
  IF EMPTY(THIS.cPrinter)
   THIS.cPrinter = findPrinter("port","XPSPort:")
   IF EMPTY(THIS.cPrinter)
    THIS.cPrinter = findPrinter("driver", "Microsoft XPS Document Writer")
   ENDIF
  ENDIF
 ENDPROC
 
 
 * isAvailable (Method)
 * Check if XPS-PDF convertion is possible
 *
 PROCEDURE isAvailable
  DO CASE
     CASE EMPTY(THIS.cGXPS)
          THIS.setError("Property oXPS.cGXPS is empty")
          
     CASE !THIS.IsFile(THIS.cGXPS)
          THIS.setError("File " + THIS.cGXPS + " is not found")
          
     CASE LOWER(JUSTEXT(THIS.cGXPS)) <> "exe"
          THIS.setError("Invalid file " + THIS.cGXPS)
          
     CASE EMPTY(THIS.cPrinter)
          THIS.setError("The Microsoft XPS Document Writer is not available at this time")
  ENDCASE
  RETURN !THIS.lError
 ENDPROC
 
 
 
 * Convert (Method)
 * Convert from XPS to PDF
 *
 PROCEDURE Convert(pcInput, pcOutput)
    IF !THIS.IsFile(pcInput)
     THIS.cError = "File " + pcInput + " could not be found"
     THIS.lERror = false
     RETURN false
    ENDIF
    
	IF THIS.IsFile(pcOutput)
	 _TRY
	  THIS.deleteFile(pcOutput)
	 _CATCH
	 _ENDTRY
	ENDIF

	pcInput = THIS.fixFileNAme(pcInput)
	pcOutput = THIS.fixFIleName(pcOutput)

	
    LOCAL oArgs
    THIS.oGSUtils.oOptions = THIS.oOptions
    oArgs = THIS.oGSUtils.getConvertArgs()
	oArgs.Add("-sOutputFile=" + pcOutput)	&&Name of the output file
 	
	LOCAL cArgs,cArg,nArg
	cArgs = ""
	FOR nArg = 1 TO oArgs.Count
	 cArg = oArgs.Item[nArg]
	 cArgs = cArgs + " " + cArg
	ENDFOR

	LOCAL cBAT, cCMD
	cBAT = THIS.fixFIleName(FORCEEXT(CHRTRAN(pcInput,["'],""), "BAT"))
	cCMD = ["] + THIS.cGXPS + ["] + cArgs + [ ] + pcInput
	*MESSAGEBOX(cCmd)
	
	STRTOFILE(cCmd, cBAT)


	LOCAL oWSH
	oWSH = CREATEOBJECT("wscript.shell")
	_TRY
	 oWSH.Run(cBAT, 0, 1)
	 IF NOEX() AND !THIS.isFile(pcOutput)
	  THIS.setError("GXPS could not create the PDF")
	 ENDIF

	_CATCH
	 THIS.lError = true
	 THIS.cError = ex.Message

    _ENDTRY

	IF !THIS.lDebug AND !THIS.lError
	 _TRY
	  THIS.deleteFile(cBAT)
	 _CATCH
	 _ENDTRY
	ENDIF
	RELEASE oWSH
	 	
	RETURN !THIS.lError
 ENDPROC



ENDDEFINE



* pdfOptions (Class)
* Opciones para la generacion de PDF
*
DEFINE CLASS pdfOptions AS Custom
 cResolution = "600"
 cOwnerPwd = ""
 cUserPwd = ""
 cKeyLength = "128"
 cEncryptionR = "3"
 cPermissions = "-3904"
 oGSOptions = NULL
 PROCEDURE Init
  THIs.oGSOptions = CREATEOBJECT("Collection")
 ENDPROC
ENDDEFINE


* pdfGSUtils (Class)
* Utilidades para Ghostscript
*
DEFINE CLASS pdfGSUtils AS Custom
 oOptions = NULL
 PROCEDURE Init
  THIS.oOptions = CREATEOBJECT("pdfOptions")
 ENDPROC




 PROCEDURE getConvertArgs(poArgs)
    LOCAL oArgs
    oArgs = IIF(PCOUNT() = 0, CREATEOBJECT("Collection"), poArgs)
	oArgs.Add("-dNOPAUSE")					&&Disables Prompt and Pause after each page
	oArgs.Add("-dBATCH")					&&Causes GS to exit after processing file(s)
	oArgs.Add("-dSAFER")					&&Disables the ability to deletefile and renamefile externally
	oArgs.Add("-r"+THIS.oOptions.cResolution)		&&Printer Resolution (300x300, 360x180, 600x600, etc.)
	oArgs.Add("-sDEVICE=pdfwrite")			&&Specifies which "Device" (output type) to use.  "pdfwrite" means PDF file.
	IF !EMPTY(THIS.oOptions.cOwnerPwd)
 	 oArgs.Add("-sOwnerPassword=" + THIS.oOptions.cOwnerPwd) 
	 oArgs.Add("-sUserPassword=" + EVL(THIS.oOptions.cUserPwd, THIS.oOptions.cOwnerPwd))
	 oArgs.Add("-dKeyLength=" + THIS.oOptions.cKeyLength)
	 oArgs.Add("-dEncryptionR=" + THIS.oOptions.cEncryptionR)
	 oArgs.Add("-dPermissions=" + THIS.oOptions.cPermissions)
	ENDIF
	LOCAL cArg,nArg
	FOR nArg = 1 TO THIS.oOptions.oGSOptions.Count
	 cArg = THIS.oOPtions.oGSOptions.Item[nArg]
	 oArgs.Add(cArg)
	ENDFOR
	RETURN oArgs 
 ENDPROC


 PROCEDURE getProtectArgs(poArgs)
    LOCAL oArgs
    oArgs = IIF(PCOUNT() = 0, CREATEOBJECT("Collection"), poArgs)
    oArgs = THIS.getConvertArgs(oArgs)
	oArgs.Add("-dPDFSettings=/prepress")
	oArgs.Add("-dPassThroughJPEGImages=true")
    RETURN oArgs 
 ENDPROC
 
ENDDEFINE


* findFiles (Function)
* Permite ubicar archivos dentro de una ruta, mediante un patròn de búsqueda. La
* función devuelve la cantidad de archivos encontrados y en el parámetro paFiles
* (que debe ser pasado por referencia) devuelve la lista de archivos.
* 
* Autor: Victor Espina 
* Fecha: Oct 2014
*
* Ejemplo:
* LOCAL ARRAY paFiles[1]
* nCount = findFiles(@paFiles, "c:\", "*.prg")
*
* El parametro pnNestLevel permite indicar un limite de recursion, para limitar
* la cantidad de subdirectorios a considerar en la busqueda. Por ejemplo, para
* buscar archivos DLL en una ruta o en sus subdirectorios directos, hariamos:
*
* nCount = findFiles(@paFiles, "c:\windows", "*.dll", 1)
*
FUNCTION findFiles(paFiles, pcFolder, pcWildcard, pnNestlevel, pnCurrentLevel)
 
 * Si asignan valores por defecto a los parametros opcionales
 pcWildcard = IIF(EMPTY(pcWildcard),"*.*",pcWildcard)
 pnNestLevel = IIF(EMPTY(pnNestLevel),-1,pnNestLevel)
 pnCurrentLevel = IIF(EMPTY(pnCurrentLevel),0,pnCurrentLevel)
 
 * Se inicia la busqueda
 LOCAL ARRAY aMatchs[1]
 LOCAL nCount,i,cFile,cSubFolder,nSize
 nSize = ALEN(paFiles,1)
 IF EMPTY(paFiles[1])   && El array esta vacio
  nSize = 0 
 ENDIF
 
 * Buscamos dentro de la ruta indicada
 pcFolder = ADDBS(pcFolder)
 nCount = ADIR(aMatchs, pcFolder + pcWildcard)
 FOR i = 1 TO nCount
  nSize = nSize + 1
  DIMEN paFiles[nSize]
  paFiles[nSize] = pcFolder + aMatchs[i,1]
 ENDFOR
 
 * Si se llego al maximo nivel de anidamiento, se finaliza aqui
 IF pnCurrentLevel = pnNestLevel
  RETURN nSize
 ENDIF
 
 * Buscamos ahora dentro de las carpetas
 nCount = ADIR(aMatchs, pcFolder + "*.*", "D")
 FOR i = 1 TO nCount
  cSubFolder = aMatchs[i,1]
  IF "D" $ aMatchs[i,5] AND !INLIST(cSubFolder,".","..")
   cSubFolder = pcFolder + cSubFolder
   nSize = findFiles(@paFiles, cSubFolder, pcWildcard, pnNestLevel, pnCurrentLevel + 1)
  ENDIF
 ENDFOR
 
 RETURN nSize
ENDFUNC


* findFolders (Function)
* Permite ubicar carpetas dentro de una ruta, mediante un patròn de búsqueda. La
* función devuelve la cantidad de carpetas encontradas y en el parámetro paFolders
* (que debe ser pasado por referencia) devuelve la lista de carpetas.
* 
* Autor: Victor Espina 
* Fecha: May 2015
*
* Ejemplo:
* LOCAL ARRAY paFolders[1]
* nCount = findFolders(@paFolders, "c:\", "Program*")
*
* El parametro pnNestLevel permite indicar un limite de recursion, para limitar
* la cantidad de subdirectorios a considerar en la busqueda. Por ejemplo, para
* buscar carpetas BIN en una ruta o en sus subdirectorios directos, hariamos:
*
* nCount = findFolders(@paFolders, "c:\windows", "bin", 1)
*
FUNCTION findFolders(paFolders, pcFolder, pcWildcard, pnNestlevel, pnCurrentLevel)
 
 * Si asignan valores por defecto a los parametros opcionales
 pcWildcard = IIF(EMPTY(pcWildcard),"*.*",pcWildcard)
 pnNestLevel = IIF(EMPTY(pnNestLevel),-1,pnNestLevel)
 pnCurrentLevel = IIF(EMPTY(pnCurrentLevel),0,pnCurrentLevel)
 
 * Se inicia la busqueda
 LOCAL ARRAY aMatchs[1]
 LOCAL nCount,i,cFolder,cSubFolder,nSize
 nSize = ALEN(paFolders,1)
 IF EMPTY(paFolders[1])   && El array esta vacio
  nSize = 0 
 ENDIF
 
 * Buscamos dentro de la ruta indicada
 pcFolder = ADDBS(pcFolder)
 nCount = ADIR(aMatchs, pcFolder + pcWildcard, "D")
 FOR i = 1 TO nCount
  nSize = nSize + 1
  DIMEN paFolders[nSize]
  paFolders[nSize] = pcFolder + aMatchs[i,1]
 ENDFOR
 
 * Si se llego al maximo nivel de anidamiento, se finaliza aqui
 IF pnCurrentLevel = pnNestLevel
  RETURN nSize
 ENDIF
 
 * Buscamos ahora dentro de las carpetas
 nCount = ADIR(aMatchs, pcFolder + "*.*", "D")
 FOR i = 1 TO nCount
  cSubFolder = aMatchs[i,1]
  IF "D" $ aMatchs[i,5] AND !INLIST(cSubFolder,".","..")
   cSubFolder = pcFolder + cSubFolder
   nSize = findFolders(@paFolders, cSubFolder, pcWildcard, pnNestLevel, pnCurrentLevel + 1)
  ENDIF
 ENDFOR
 
 RETURN nSize
ENDFUNC




* findPrinter (Function)
* Finds a printer by its name, port or driver
*
FUNCTION findPrinter(pcSearchIn, pcSearchFor)
 LOCAL ARRAY aList[1]
 LOCAL nCount, i, cPrinter
 #IF VERSION(5) >= 900
  nCount = APRINTERS(aList, 1)
 #ELSE
  nCount = APRINTERS(aList)
 #ENDIF
 cPrinter = ""
 pcSearchIn = LOWER(pcSearchIn)
 FOR i = 1 TO nCount
  IF (pcSearchIn == "name" AND LIKE(pcSearchFor, aList[i,1])) OR ;
     (pcSearchIn == "port" AND aList[i,2] == pcSearchFor) OR ;
     (VERSION(5) >= 900 AND (pcSearchIn == "driver" AND LIKE(pcSearchFor, aList[i,3]))) 
   cPrinter = aList[i,1]
   EXIT
  ENDIF
 ENDFOR
 RETURN cPrinter
ENDFUNC




**************************************************
*-- Class:        heap 
*-- ParentClass:  custom
*-- BaseClass:    custom
*
*  Another in the family of relatively undocumented sample classes I've inflicted on others
*  Warning - there's no error handling in here, so be careful to check for null returns and
*  invalid pointers.  Unless you get frisky, or you're resource-tight, it should work well.
*
*	Please read the code and comments carefully.  I've tried not to assume much knowledge about
*	just how pointers work, or how memory allocation works, and have tried to explain some of the
*	basic concepts behing memory allocation in the Win32 environment, without having gone into
*	any real details on x86 memory management or the Win32 memory model.  If you want to explore
*	these things (and you should), start by reading Jeff Richter's _Advanced Windows_, especially
*	Chapters 4-6, which deal with the Win32 memory model and virtual memory -really- well.
*
*	Another good source iss Walter Oney's _Systems Programming for Windows 95_.  Be warned that 
*	both of these books are targeted at the C programmer;  to someone who has only worked with
*	languages like VFP or VB, it's tough going the first couple of dozen reads.
*
*	Online resources - http://www.x86.org is the Intel Secrets Homepage.  Lots of deep, dark
*	stuff about the x86 architecture.  Not for the faint of heart.  Lots of pointers to articles
*	from DDJ (Doctor Dobbs Journal, one of the oldest and best magazines on microcomputing.)
*
*   You also might want to take a look at the transcripts from my "Pointers on Pointers" chat
*   sessions, which are available in the WednesdayNightLectureSeries topic on the Fox Wiki,
*   http://fox.wikis.com - the Wiki is a great Web site;  it provides a vast store of information
*   on VFP and related topics, and is probably the best tool available now to develop topics in
*   a collaborative environment.  Well worth checking out - it's a very different mechanism for
*   ongoing discussion of a subject.  It's an on-line message base or chat;  I find
*   myself hitting it when I have a question to see if an answer already exists.  It's
*   much like using a FAQ, except that most things on the Wiki are editable...
*
*	Post-DevCon 2000 revisions:
*
*	After some bizarre errors at DevCon, I reworked some of the methods to
*	consistently return a NULL whenever a bad pointer/inactive pointer in the
*	iaAllocs member array was encountered.  I also implemented NumToLong
*	using RtlMoveMemory(), relying on a BITOR() to recast what would otherwise
*	be a value with the high-order bit set.  The result is it's faster, and
*  an anomaly reported with values between 0xFFFFFFF1-0xFFFFFFFF goes away,
*	at the expense of representing these as negative numbers.  Pointer math
*	still works.
*
*****
*	How HEAP works:
*
*	Overwhelming guilt hit early this morning;  maybe I should explain the 
*	concept of the Heap class	and give an example of how to use it, in 
*	conjunction with the add-on functions that follow in this proc library.
*
*	Windows allocates memory from several places;  it also provides a 
*	way to define your own small corner of the universe where you can 
*	allocate and deallocate blocks of memory for your own purposes.  These
*	public or private memory areas are referred to commonly as heaps.
*
*	VFP is great in most cases;  it provides flexible allocation and 
*	alteration of variables on the fly in a program.  You don't need to 
*	know much about how things are represented internally. This makes 
*	most programming tasks easy.  However, in exchange for VFP's flexibility 
*	in memory variable allocation, we give up several things, the most 
*	annoying of which are not knowing the exact location of a VFP 
*	variable in memory, and not knowing exactly how things are constructed 
*	inside a variable, both of which make it hard to define new kinds of 
*	memory structures within VFP to manipulate as a C-style structure.
*
*	Enter Heap.  Heap creates a growable, private heap, from which you 
*	can allocate blocks of memory that have a known location and size 
*	in your memory address space.  It also provides a way of transferring
*	data to and from these allocated blocks.  You build structures in VFP 
*	strings, and parse the content of what is returned in those blocks by 
*	extracting substrings from VFP strings.
*
*	Heap does its work using a number of Win32 API functions;  HeapCreate(), 
*	which sets up a private heap and assigns it a handle, is invoked in 
*	the Init method.  This sets up the 'heap', where block allocations
*	for the object will be constructed.  I set up the heap to use a base 
*	allocation size of twice the size of a swap file 'page' in the x86 
*	world (8K), and made the heap able to grow;  it adds 8K chunks of memory
*	to itself as it grows.  There's no fixed limit (other than available 
*	-virtual- memory) on the size of the heap constructed;  just realize 
*	that huge allocations are likely to bump heads with VFP's own desire
*	for mondo RAM.
*
*	Once the Heap is established, we can allocate blocks of any size we 
*	want in Heap, outside of VFP's memory, but within the virtual 
*	address space owned by VFP.  Blocks are allocated by HeapAlloc(), and a
*	pointer to the block is returned as an integer.  
*
*	KEEP THE POINTER RETURNED BY ANY Alloc method, it's the key to 
*	doing things with the block in the future.  In addition to being a
*	valid pinter, it's the key to finding allocations tracked in iaAllocs[]
*
*	Periodically, we need to load things into the block we've created.  
*	Thanks to work done by Christof Lange, George Tasker and others, 
*	we found a Win32API call that will do transfers between memory 
*	locations, called RtlMoveMemory().  RtlMoveMemory() acts like the 
*	Win32API MoveMemory() call;  it takes two pointers (destination 
*	and source) and a length.  In order to make life easy, at times 
*	we DECLARE the pointers as INTEGER (we pass a number, which is 
*	treated as a DWORD (32 bit unsigned integer) whose content is the
*	address to use), and at other times as STRING @, which passes the 
*	physical address of a VFP string variable's contents, allowing 
*	RtlMoveMemory() to read and write VFP strings without knowing how 
*	to manipulate VFP's internal variable structures.  RtlMoveMemory() 
*	is used by both the CopyFrom and CopyTo methods, and the enhanced
*	Alloc methods.
*
*	At some point, we're finished with a block of memory.  We can free up 
*	that memory via HeapFree(), which releases a previously-allocated 
*	block on the heap.  It does not compact or rearrange the heap allocations
*	but simply makes the memory allocated no longer valid, and the 
*	address could be reused by another Alloc operation.  We track the 
*	active state of allocations in a member array iaAllocs[] which has 
*	3 members per row;  the pointer, which is used as a key, the actual 
*	size of the allocation (sometimes HeapAlloc() gives you a larger block 
*	than requested;  we can see it here.  This is the property returned 
*	by the SizeOfBlock method) and whether or not it's active and available.
*
*	When we're done with a Heap, we need to release the allocations and 
*	the heap itself.  HeapDestroy() releases the entire heap back to the 
*	Windows memory pool.  This is invoked in the Destroy method of the 
*	class to ensure that it gets explcitly released, since it remains alive 
*	until it is explicitly released or the owning process is released.  I 
*	put this in the Destroy method to ensure that the heap went away when 
*	the Heap object went out of scope.
*
*	The original class methods are:
*
*		Init					Creates the heap for use
*		Alloc(nSize)		Allocates a block of nSize bytes, returns an nPtr 
*								to it.  nPtr is NULL if fail
*		DeAlloc(nPtr)		Releases the block whose base address is nPtr.  
*								Returns true/false
*		CopyTo(nPtr,cSrc)	Copies the Content of cSrc to the buffer at nPtr, 
*								up to the smaller of LEN(cSrc) or the length of 
*								the block (we look in the iaAllocs[] array).  
*								Returns true/false
*		CopyFrom(nPtr)		Copies the content of the block at nPtr (size is 
*								from iaAllocs[]) and returns it as a VFP string.  
*								Returns a string, or NULL if fail
*		SizeOfBlock(nPtr)	Returns the actual allocated size of the block 
*								pointed to by nPtr.  Returns NULL if fail 
*		Destroy()			DeAllocs anything still active, and then frees 
*								the heap.
*****
*  New methods added 2/12/99 EMR -	Attack of the Creeping Feature Creature, 
*												part I
*
*	There are too many times when you know what you want to put in 
*	a buffer when you allocate it, so why not pass what you want in 
*	the buffer when you allocate it?  And we may as well add an option to
*	init the memory to a known value easily, too:
*
*		AllocBLOB(cSrc)	Allocate a block of SizeOf(cSrc) bytes and 
*								copy cSrc content to it
*		AllocString(cSrc)	Allocate a block of SizeOf(cSrc) + 1 bytes and 
*								copy cSrc content to it, adding a null (CHR(0)) 
*								to the end to make it a standard C-style string
*		AllocInitAs(nSize,nVal)
*								Allocate a block of nSize bytes, preinitialized 
*								with CHR(nVal).  If no nVal is passed, or nVal 
*								is illegal (not a number 0-255), init with nulls
*
*****
*	Property changes 9/29/2000
*
*	iaAllocs[] is now protected
*
*****
*	Method modifications 9/29/2000:
*
*	All lookups in iaAllocs[] are now done using the new FindAllocID()
*	method, which returns a NULL for the ID if not found active in the
*	iaAllocs[] entries.  Result is less code and more consistent error
*	handling, based on checking ISNULL() for pointers.
*
*****
*	The ancillary goodies in the procedure library are there to make life 
*	easier for people working with structures; they are not optimal 
*	and infinitely complete, but they do the things that are commonly 
*	needed when dealing with stuff in structures.  The functions are of 
*	two types;  converters, which convert standard C structures to an
*	equivalent VFP numeric, or make a string whose value is equivalent 
*	to a C data type from a number, so that you can embed integers, 
*	pointers, etc. in the strings used to assemble a structure which you 
*	load up with CopyTo, or pull out pointers and integers that come back 
*	embedded in a structure you've grabbed with CopyFrom.
*
*	The second type of functions provided are memory copiers.  The 
*	CopyFrom and CopyTo methods are set up to work with our heap, 
*	and nPtrs must take on the values of block addresses grabbed 
*	from our heap.  There will be lots of times where you need to 
*	get the content of memory not necessarily on our heap, so 
*	SetMem, GetMem and GetMemString go to work for us here.  SetMem 
*	copies the content of a string into the absolute memory block
*	at nPtr, for the length of the string, using RtlMoveMemory(). 
*	BE AWARE THAT MISUSE CAN (and most likely will) RESULT IN 
*	0xC0000005 ERRORS, memory access violations, or similar OPERATING 
*	SYSTEM level errors that will smash VFP like an empty beer can in 
*	a trash compactor.
*
*	There are two functions to copy things from a known address back 
*	to the VFP world.  If you know the size of the block to grab, 
*	GetMem(nPtr,nSize) will copy nSize bytes from the address nPtr 
*	and return it as a VFP string.  See the caveat above.  
*	GetMemString(nPtr) uses a different API call, lstrcpyn(), to 
*	copy a null terminated string from the address specified by nPtr. 
*	You can hurt yourself with this one, too.
*
*	Functions in the procedure library not a part of the class:
*
*	GetMem(nPtr,nSize)	Copy nSize bytes at address nPtr into a VFP string
*	SetMem(nPtr,cSource)	Copy the string in cSource to the block beginning 
*								at nPtr
*	GetMemString(nPtr)	Get the null-terminated string (up to 512 bytes) 
*								from the address at nPtr
*
*	DWORDToNum(cString)	Convert the first 4 bytes of cString as a DWORD 
*								to a VFP numeric (0 to 2^32)
*	SHORTToNum(cString)	Convert the first 2 bytes of cString as a SHORT 
*								to a VFP numeric (-32768 to 32767)
*	WORDToNum(cString)	Convert the first 2 bytes of cString as a WORD 
*								to a VFP numeric  (0 to 65535)
*	NumToDWORD(nInteger)	Converts nInteger into a string equivalent to a 
*								C DWORD (4 byte unsigned)
*	NumToWORD(nInteger)	Converts nInteger into a string equivalent to a 
*								C WORD (2 byte unsigned)
*	NumToSHORT(nInteger)	Converts nInteger into a string equivalent to a 
*								C SHORT ( 2 byte signed)
*
******
*	New external functions added 2/13/99
*
*	I see a need to handle NetAPIBuffers, which are used to transfer 
*	structures for some of the Net family of API calls;  their memory 
*	isn't on a user-controlled heap, but is mapped into the current 
*	application address space in a special way.  I've added two 
*	functions to manage them, but you're responsible for releasing 
*	them yourself.  I could implement a class, but in many cases, a 
*	call to the API actually performs the allocation for you.  The 
*	two new calls are:
*
*	AllocNetAPIBuffer(nSize)	Allocates a NetAPIBuffer of at least 
*										nBytes, and returns a pointer
*										to it as an integer.  A NULL is returned 
*										if allocation fails.
*	DeAllocNetAPIBuffer(nPtr)	Frees the NetAPIBuffer allocated at the 
*										address specified by nPtr.  It returns 
*										true/false for success and failure
*
*	These functions are only available under NT, and will return 
*	NULL or false under Win9x
*
*****
*	Function changes 9/29/2000
*
*	NumToDWORD(tnNum)		Redirected to NumToLONG()
*	NumToLONG(tnNum)		Generates a 32 bit LONG from the VFP number, recast
*								using BITOR() as needed
*	LONGToNum(tcLong)		Extracts a signed VFP INTEGER from a 4 byte string
*
*****
*	That's it for the docs to date;  more stuff to come.  The code below 
*	is copyright Ed Rauh, 1999;  you may use it without royalties in 
*	your own code as you see fit, as long as the code is attributed to me.
*
*	This is provided as-is, with no implied warranty.  Be aware that you 
*	can hurt yourself with this code, most *	easily when using the 
*	SetMem(), GetMem() and GetMemString() functions.  I will continue to 
*	add features and functions to this periodically.  If you find a bug, 
*	please notify me.  It does no good to tell me that "It doesn't work 
*	the way I think it should..WAAAAH!"  I need to know exactly how things 
*	fail to work with the code I supplied.  A small code snippet that can 
*	be used to test the failure would be most helpful in trying
*	to track down miscues.  I'm not going to run through hundreds or 
*	thousands of lines of code to try to track down where exactly 
*	something broke.  
*
*	Please post questions regarding this code on Universal Thread;  I go out 
*	there regularly and will generally respond to questions posed in the
*	message base promptly (not the Chat).  http://www.universalthread.com
*	In addition to me, there are other API experts who frequent UT, and 
*	they may well be able to help, in many cases better than I could.  
*	Posting questions on UT helps not only with getting support
*	from the VFP community at large, it also makes the information about 
*	the problem and its solution available to others who might have the 
*	same or similar problems.
*
*	Other than by UT, especially if you have to send files to help 
*	diagnose the problem, send them to me at edrauh@earthlink.net or 
*	erauh@snet.net, preferably the earthlink.net account.
*
*	If you have questions about this code, you can ask.  If you have 
*	questions about using it with API calls and the like, you can ask.  
*	If you have enhancements that you'd like to see added to the code, 
*	you can ask, but you have the source, and ought to add them yourself.
*	Flames will be ignored.  I'll try to answer promptly, but realize 
*	that support and enhancements for this are done in my own spare time.  
*	If you need specific support that goes beyond what I feel is 
*	reasonable, I'll tell you.
*
*	Do not call me at home or work for support.  Period. 
*	<Mumble><something about ripping out internal organs><Grr>
*
*	Feel free to modify this code to fit your specific needs.  Since 
*	I'm not providing any warranty with this in any case, if you change 
*	it and it breaks, you own both pieces.
*
DEFINE CLASS heap AS custom


	PROTECTED inHandle, inNumAllocsActive,iaAllocs[1,3]
	inHandle = NULL
	inNumAllocsActive = 0
	iaAllocs = NULL
	Name = "heap"

	PROCEDURE Alloc
		*  Allocate a block, returning a pointer to it
		LPARAMETER nSize
		DECLARE INTEGER HeapAlloc IN WIN32API AS HAlloc;
			INTEGER hHeap, ;
			INTEGER dwFlags, ;
			INTEGER dwBytes
		DECLARE INTEGER HeapSize IN WIN32API AS HSize ;
			INTEGER hHeap, ;
			INTEGER dwFlags, ;
			INTEGER lpcMem
		LOCAL nPtr
		WITH this
			nPtr = HAlloc(.inHandle, 0, @nSize)
			IF nPtr # 0
				*  Bump the allocation array
				.inNumAllocsActive = .inNumAllocsActive + 1
				DIMENSION .iaAllocs[.inNumAllocsActive,3]
				*  Pointer
				.iaAllocs[.inNumAllocsActive,1] = nPtr
				*  Size actually allocated - get with HeapSize()
				.iaAllocs[.inNumAllocsActive,2] = HSize(.inHandle, 0, nPtr)
				*  It's alive...alive I tell you!
				.iaAllocs[.inNumAllocsActive,3] = true
			ELSE
				*  HeapAlloc() failed - return a NULL for the pointer
				nPtr = NULL
			ENDIF
		ENDWITH
		RETURN nPtr
	ENDPROC

*	new methods added 2/11-2/12;  pretty simple, actually, but they make 
*	coding using the heap object much cleaner.  In case it isn't clear, 
*	what I refer to as a BString is just the normal view of a VFP string 
*	variable;  it's any array of char with an explicit length, as opposed 
*	to the normal CString view of the world, which has an explicit
*	terminator (the null char at the end.)

	FUNCTION AllocBLOB
		*	Allocate a block of memory the size of the BString passed.  The 
		*	allocation will be at least LEN(cBStringToCopy) off the heap.
		LPARAMETER cBStringToCopy
		LOCAL nAllocPtr
		WITH this
			nAllocPtr = .Alloc(LEN(cBStringToCopy))
			IF ! ISNULL(nAllocPtr)
				.CopyTo(nAllocPtr,cBStringToCopy)
			ENDIF
		ENDWITH
		RETURN nAllocPtr
	ENDFUNC
	
	FUNCTION AllocString
		*	Allocate a block of memory to fill with a null-terminated string
		*	make a null-terminated string by appending CHR(0) to the end
		*	Note - I don't check if a null character precedes the end of the
		*	inbound string, so if there's an embedded null and whatever is
		*	using the block works with CStrings, it might bite you.
		LPARAMETER cString
		RETURN this.AllocBLOB(cString + CHR(0))
	ENDFUNC
	
	FUNCTION AllocInitAs
		*  Allocate a block of memory preinitialized to CHR(nByteValue)
		LPARAMETER nSizeOfBuffer, nByteValue
		IF TYPE('nByteValue') # 'N' OR ! BETWEEN(nByteValue,0,255)
			*	Default to initialize with nulls
			nByteValue = 0
		ENDIF
		RETURN this.AllocBLOB(REPLICATE(CHR(nByteValue),nSizeOfBuffer))
	ENDFUNC

*	This is the end of the new methods added 2/12/99

	PROCEDURE DeAlloc
		*  Discard a previous Allocated block
		LPARAMETER nPtr
		DECLARE INTEGER HeapFree IN WIN32API AS HFree ;
			INTEGER hHeap, ;
			INTEGER dwFlags, ;
			INTEGER lpMem
		LOCAL nCtr
		* Change to use .FindAllocID() and return !ISNULL() 9/29/2000 EMR
		nCtr = NULL
		WITH this
			nCtr = .FindAllocID(nPtr)
			IF ! ISNULL(nCtr)
				=HFree(.inHandle, 0, nPtr)
				.iaAllocs[nCtr,3] = false
			ENDIF
		ENDWITH
		RETURN ! ISNULL(nCtr)
	ENDPROC


	PROCEDURE CopyTo
		*  Copy a VFP string into a block
		LPARAMETER nPtr, cSource
		*  ReDECLARE RtlMoveMemory to make copy parameters easy
		DECLARE RtlMoveMemory IN WIN32API AS RtlCopy ;
			INTEGER nDestBuffer, ;
			STRING @pVoidSource, ;
			INTEGER nLength
		LOCAL nCtr
		nCtr = NULL
		* Change to use .FindAllocID() and return ! ISNULL() 9/29/2000 EMR
		IF TYPE('nPtr') = 'N' AND TYPE('cSource') $ 'CM' ;
		   AND ! (ISNULL(nPtr) OR ISNULL(cSource))
			WITH this
				*  Find the Allocation pointed to by nPtr
				nCtr = .FindAllocID(nPtr)
				IF ! ISNULL(nCtr)
					*  Copy the smaller of the buffer size or the source string
					=RtlCopy((.iaAllocs[nCtr,1]), ;
							cSource, ;
							MIN(LEN(cSource),.iaAllocs[nCtr,2]))
				ENDIF
			ENDWITH
		ENDIF
		RETURN ! ISNULL(nCtr)
	ENDPROC


	PROCEDURE CopyFrom
		*  Copy the content of a buffer back to the VFP world
		LPARAMETER nPtr
		*  Note that we reDECLARE RtlMoveMemory to make passing things easier
		DECLARE RtlMoveMemory IN WIN32API AS RtlCopy ;
			STRING @DestBuffer, ;
			INTEGER pVoidSource, ;
			INTEGER nLength
		LOCAL nCtr, uBuffer
		uBuffer = NULL
		nCtr = NULL
		* Change to use .FindAllocID() and return NULL 9/29/2000 EMR
		IF TYPE('nPtr') = 'N' AND ! ISNULL(nPtr)
			WITH this
				*  Find the allocation whose address is nPtr
				nCtr = .FindAllocID(nPtr)
				IF ! ISNULL(nCtr)
					* Allocate a buffer in VFP big enough to receive the block
					uBuffer = REPL(CHR(0),.iaAllocs[nCtr,2])
					=RtlCopy(@uBuffer, ;
							(.iaAllocs[nCtr,1]), ;
							(.iaAllocs[nCtr,2]))
				ENDIF
			ENDWITH
		ENDIF
		RETURN uBuffer
	ENDPROC
	
	PROTECTED FUNCTION FindAllocID
	 	*   Search for iaAllocs entry matching the pointer
	 	*   passed to the function.  If found, it returns the 
	 	*   element ID;  returns NULL if not found
	 	LPARAMETER nPtr
	 	LOCAL nCtr
	 	WITH this
			FOR nCtr = 1 TO .inNumAllocsActive
				IF .iaAllocs[nCtr,1] = nPtr AND .iaAllocs[nCtr,3]
					EXIT
				ENDIF
			ENDFOR
			RETURN IIF(nCtr <= .inNumAllocsActive,nCtr,NULL)
		ENDWITH
	ENDPROC

	PROCEDURE SizeOfBlock
		*  Retrieve the actual memory size of an allocated block
		LPARAMETERS nPtr
		LOCAL nCtr, nSizeOfBlock
		nSizeOfBlock = NULL
		* Change to use .FindAllocID() and return NULL 9/29/2000 EMR
		WITH this
			*  Find the allocation whose address is nPtr
			nCtr = .FindAllocID(nPtr)
			RETURN IIF(ISNULL(nCtr),NULL,.iaAllocs[nCtr,2])
		ENDWITH
	ENDPROC

	PROCEDURE Destroy
		DECLARE HeapDestroy IN WIN32API AS HDestroy ;
		  INTEGER hHeap

		LOCAL nCtr
		WITH this
			FOR nCtr = 1 TO .inNumAllocsActive
				IF .iaAllocs[nCtr,3]
					.Dealloc(.iaAllocs[nCtr,1])
				ENDIF
			ENDFOR
			HDestroy[.inHandle]
		ENDWITH
		DODEFAULT()
	ENDPROC


	PROCEDURE Init
		DECLARE INTEGER HeapCreate IN WIN32API AS HCreate ;
			INTEGER dwOptions, ;
			INTEGER dwInitialSize, ;
			INTEGER dwMaxSize
		#DEFINE SwapFilePageSize  4096
		#DEFINE BlockAllocSize    2 * SwapFilePageSize
		WITH this
			.inHandle = HCreate(0, BlockAllocSize, 0)
			DIMENSION .iaAllocs[1,3]
			.iaAllocs[1,1] = 0
			.iaAllocs[1,2] = 0
			.iaAllocs[1,3] = false
			.inNumAllocsActive = 0
		ENDWITH
		RETURN (this.inHandle # 0)
	ENDPROC


ENDDEFINE
*
*-- EndDefine: heap
**************************************************
*
*  Additional functions for working with structures and pointers and stuff
*
FUNCTION SetMem
LPARAMETERS nPtr, cSource
*  Copy cSource to the memory location specified by nPtr
*  ReDECLARE RtlMoveMemory to make copy parameters easy
*  nPtr is not validated against legal allocations on the heap
DECLARE RtlMoveMemory IN WIN32API AS RtlCopy ;
	INTEGER nDestBuffer, ;
	STRING @pVoidSource, ;
	INTEGER nLength

RtlCopy(nPtr, ;
		cSource, ;
		LEN(cSource))
RETURN true

FUNCTION GetMem
LPARAMETERS nPtr, nLen
*  Copy the content of a memory block at nPtr for nLen bytes back to a VFP string
*  Note that we ReDECLARE RtlMoveMemory to make passing things easier
*  nPtr is not validated against legal allocations on the heap
DECLARE RtlMoveMemory IN WIN32API AS RtlCopy ;
	STRING @DestBuffer, ;
	INTEGER pVoidSource, ;
	INTEGER nLength
LOCAL uBuffer
* Allocate a buffer in VFP big enough to receive the block
uBuffer = REPL(CHR(0),nLen)
=RtlCopy(@uBuffer, ;
		 nPtr, ;
		 nLen)
RETURN uBuffer

FUNCTION GetMemString
LPARAMETERS nPtr, nSize
*  Copy the string at location nPtr into a VFP string
*  We're going to use lstrcpyn rather than RtlMoveMemory to copy up to a terminating null
*  nPtr is not validated against legal allocations on the heap
*
*	Change 9/29/2000 - second optional parameter nSize added to allow an override
*	of the string length;  no major expense, but probably an open invitation
*	to cliff-diving, since variant CStrings longer than 511 bytes, or less
*	often, 254 bytes, will generally fall down go Boom!
*
DECLARE INTEGER lstrcpyn IN WIN32API AS StrCpyN ;
	STRING @ lpDestString, ;
	INTEGER lpSource, ;
	INTEGER nMaxLength
LOCAL uBuffer
IF TYPE('nSize') # 'N' OR ISNULL(nSize)
	nSize = 512
ENDIF
*  Allocate a buffer big enough to receive the data
uBuffer = REPL(CHR(0), nSize)
IF StrCpyN(@uBuffer, nPtr, nSize-1) # 0
	uBuffer = LEFT(uBuffer, MAX(0,AT(CHR(0),uBuffer) - 1))
ELSE
	uBuffer = NULL
ENDIF
RETURN uBuffer

FUNCTION SHORTToNum
	* Converts a 16 bit signed integer in a structure to a VFP Numeric
 	LPARAMETER tcInt
	LOCAL b0,b1,nRetVal
	b0=asc(tcInt)
	b1=asc(subs(tcInt,2,1))
	if b1<128
		*
		*  positive - do a straight conversion
		*
		nRetVal=b1 * 256 + b0
	else
		*
		*  negative value - take twos complement and negate
		*
		b1=255-b1
		b0=256-b0
		nRetVal= -( (b1 * 256) + b0)
	endif
	return nRetVal

FUNCTION NumToSHORT
*
*  Creates a C SHORT as a string from a number
*
*  Parameters:
*
*	tnNum			(R)  Number to convert
*
	LPARAMETER tnNum
	*
	*  b0, b1, x hold small ints
	*
	LOCAL b0,b1,x
	IF tnNum>=0
		x=INT(tnNum)
		b1=INT(x/256)
		b0=MOD(x,256)
	ELSE
		x=INT(-tnNum)
		b1=255-INT(x/256)
		b0=256-MOD(x,256)
		IF b0=256
			b0=0
			b1=b1+1
		ENDIF
	ENDIF
	RETURN CHR(b0)+CHR(b1)

FUNCTION DWORDToNum
	* Take a binary DWORD and convert it to a VFP Numeric
	* use this to extract an embedded pointer in a structure in a string to an nPtr
	LPARAMETER tcDWORD
	LOCAL b0,b1,b2,b3
	b0=asc(tcDWORD)
	b1=asc(subs(tcDWORD,2,1))
	b2=asc(subs(tcDWORD,3,1))
	b3=asc(subs(tcDWORD,4,1))
	RETURN ( ( (b3 * 256 + b2) * 256 + b1) * 256 + b0)

*!*	FUNCTION NumToDWORD
*!*	*
*!*	*  Creates a 4 byte binary string equivalent to a C DWORD from a number
*!*	*  use to embed a pointer or other DWORD in a structure
*!*	*  Parameters:
*!*	*
*!*	*	tnNum			(R)  Number to convert
*!*	*
*!*	 	LPARAMETER tnNum
*!*	 	*
*!*	 	*  x,n,i,b[] will hold small ints
*!*	 	*
*!*	 	LOCAL x,n,i,b[4]
*!*		x=INT(tnNum)
*!*		FOR i=3 TO 0 STEP -1
*!*			b[i+1]=INT(x/(256^i))
*!*			x=MOD(x,(256^i))
*!*		ENDFOR
*!*		RETURN CHR(b[1])+CHR(b[2])+CHR(b[3])+CHR(b[4])
*			Redirected to NumToLong() using recasting;  comment out
*			the redirection and uncomment NumToDWORD() if original is needed
FUNCTION NumToDWORD
	LPARAMETER tnNum
	RETURN NumToLong(tnNum)
*			End redirection

FUNCTION WORDToNum
	*	Take a binary WORD (16 bit USHORT) and convert it to a VFP Numeric
	LPARAMETER tcWORD
	RETURN (256 *  ASC(SUBST(tcWORD,2,1)) ) + ASC(tcWORD)

FUNCTION NumToWORD
*
*  Creates a C USHORT (WORD) from a number
*
*  Parameters:
*
*	tnNum			(R)  Number to convert
*
	LPARAMETER tnNum
	*
	*  x holds an int
	*
	LOCAL x
	x=INT(tnNum)
	RETURN CHR(MOD(x,256))+CHR(INT(x/256))
	
FUNCTION NumToLong
*
*  Creates a C LONG (signed 32-bit) 4 byte string from a number
*  NB:  this works faster than the original NumToDWORD(), which could have
*	problems with trunaction of values > 2^31 under some versions of VFP with
*	#DEFINEd or converted constant values in excess of 2^31-1 (0x7FFFFFFF).
*	I've redirected NumToDWORD() and commented it out; NumToLong()
*	expects to work with signed values and uses BITOR() to recast values
*  in the range of -(2^31) to (2^31-1), 0xFFFFFFFF is not the same
*  as -1 when represented in an N-type field.  If you don't need to
*  use constants with the high-order bit set, or are willing to let
*  the UDF cast the value consistently, especially using pointer math 
*	on the system's part of the address space, this and its counterpart 
*	LONGToNum() are the better choice for speed, or to save to an I-field.
*
*  To properly cast a constant/value with the high-order bit set, you
*  can BITOR(nVal,0);  0xFFFFFFFF # -1 but BITOR(0xFFFFFFFF,0) = BITOR(-1,0)
*  is true, and converts the N-type in the range 2^31 - (2^32-1) to a
*  twos-complement negative integer value.  You can disable BITOR() casting
*  in this function by commenting the proper line in this UDF();  this 
*	results in a slight speed increase.
*
*  Parameters:
*
*  tnNum			(R)	Number to convert
*
	LPARAMETER tnNum
	DECLARE RtlMoveMemory IN WIN32API AS RtlCopyLong ;
		STRING @pDestString, ;
		INTEGER @pVoidSource, ;
		INTEGER nLength
	LOCAL cString
	cString = SPACE(4)
*	Function call not using BITOR()
*	=RtlCopyLong(@cString, tnNum, 4)
*  Function call using BITOR() to cast numerics
   =RtlCopyLong(@cString, BITOR(tnNum,0), 4)
	RETURN cString
	
FUNCTION LongToNum
*
*	Converts a 32 bit LONG to a VFP numeric;  it treats the result as a
*	signed value, with a range -2^31 - (2^31-1).  This is faster than
*	DWORDToNum().  There is no one-function call that causes negative
*	values to recast as positive values from 2^31 - (2^32-1) that I've
*	found that doesn't take a speed hit.
*
*  Parameters:
*
*  tcLong			(R)	4 byte string containing the LONG
*
	LPARAMETER tcLong
	DECLARE RtlMoveMemory IN WIN32API AS RtlCopyLong ;
		INTEGER @ DestNum, ;
		STRING @ pVoidSource, ;
		INTEGER nLength
	LOCAL nNum
	nNum = 0
	=RtlCopyLong(@nNum, tcLong, 4)
	RETURN nNum
	
FUNCTION AllocNetAPIBuffer
*
*	Allocates a NetAPIBuffer at least nBtes in Size, and returns a pointer
*	to it as an integer.  A NULL is returned if allocation fails.
*	The API call is not supported under Win9x
*
*	Parameters:
*
*		nSize			(R)	Number of bytes to allocate
*
LPARAMETER nSize
IF TYPE('nSize') # 'N' OR nSize <= 0
	*	Invalid argument passed, so return a null
	RETURN NULL
ENDIF
IF ! 'NT' $ OS()
	*	API call only supported under NT, so return failure
	RETURN NULL
ENDIF
DECLARE INTEGER NetApiBufferAllocate IN NETAPI32.DLL ;
	INTEGER dwByteCount, ;
	INTEGER lpBuffer
LOCAL  nBufferPointer
nBufferPointer = 0
IF NetApiBufferAllocate(INT(nSize), @nBufferPointer) # 0
	*  The call failed, so return a NULL value
	nBufferPointer = NULL
ENDIF
RETURN nBufferPointer

FUNCTION DeAllocNetAPIBuffer
*
*	Frees the NetAPIBuffer allocated at the address specified by nPtr.
*	The API call is not supported under Win9x
*
*	Parameters:
*
*		nPtr			(R) Address of buffer to free
*
*	Returns:			true/false
*
LPARAMETER nPtr
IF TYPE('nPtr') # 'N'
	*	Invalid argument passed, so return failure
	RETURN false
ENDIF
IF ! 'NT' $ OS()
	*	API call only supported under NT, so return failure
	RETURN false
ENDIF
DECLARE INTEGER NetApiBufferFree IN NETAPI32.DLL ;
	INTEGER lpBuffer
RETURN (NetApiBufferFree(INT(nPtr)) = 0)

Function CopyDoubleToString
LPARAMETER nDoubleToCopy
*  ReDECLARE RtlMoveMemory to make copy parameters easy
DECLARE RtlMoveMemory IN WIN32API AS RtlCopyDbl ;
	STRING @DestString, ;
	DOUBLE @pVoidSource, ;
	INTEGER nLength
LOCAL cString
cString = SPACE(8)
=RtlCopyDbl(@cString, nDoubleToCopy, 8)
RETURN cString

FUNCTION DoubleToNum
LPARAMETER cDoubleInString
DECLARE RtlMoveMemory IN WIN32API AS RtlCopyDbl ;
	DOUBLE @DestNumeric, ;
	STRING @pVoidSource, ;
	INTEGER nLength
LOCAL nNum
*	Christof Lange pointed out that there's a feature of VFP that results
*	in the entry in the NTI retaining its precision after updating the value
*	directly;  force the resulting precision to a large value before moving
*	data into the temp variable
nNum = 0.000000000000000000
=RtlCopyDbl(@nNum, cDoubleInString, 8)
RETURN nNum


*** End of CLSHEAP ***


#IF VERSION(5) < 800
* Collection (Class)
* Implementacion aproximada de la clase Collection de VFP8+
*
* Autor: Victor Espina
* Fecha: Octubre 2012
*
DEFINE CLASS Collection AS Custom

 DIMEN Keys[1]
 DIMEN Items[1]
 DIMEN Item[1]
 Count = 0
 
 PROCEDURE Init(pnCapacity)
  IF PCOUNT() = 0
   pnCapacity = 0
  ENDIF
  DIMEN THIS.Items[MAX(1,pnCapacity)]
  DIMEN THIS.Keys[MAX(1,pnCapacity)]
  THIS.Count = pnCapacity
 ENDPROC
  
 PROCEDURE Items_Access(nIndex1,nIndex2)
  IF VARTYPE(nIndex1) = "N"
   RETURN THIS.Items[nIndex1]
  ENDIF
  LOCAL i
  FOR i = 1 TO THIS.Count
   IF THIS.Keys[i] == nIndex1
    RETURN THIS.Items[i]
   ENDIF
  ENDFOR
 ENDPROC

 PROCEDURE Items_Assign(cNewVal,nIndex1,nIndex2)
  IF VARTYPE(nIndex1) = "N"
   THIS.Items[nIndex1] = m.cNewVal
  ELSE
   LOCAL i
   FOR i = 1 TO THIS.Count
    IF THIS.Keys[i] == nIndex1
     THIS.Items[i] = m.cNewVal
     EXIT
    ENDIF
   ENDFOR
  ENDIF 
 ENDPROC
 
 PROCEDURE Item_Access(nIndex1, nIndex2)
  RETURN THIS.Items[nIndex1]
 ENDPROC
 
 PROCEDURE Item_Assign(cNewVal, nIndex1, nIndex2)
  THIS.Items[nIndex1] = cNewVal
 ENDPROC


 PROCEDURE Clear
  DIMEN THIS.Items[1]
  DIMEN THIS.Keys[1]
  THIS.Count = 0
 ENDPROC
 
 PROCEDURE Add(puValue, pcKey)
  IF !EMPTY(pcKey) AND THIS.getKey(pcKey) > 0
   RETURN .F.
  ENDIF
  THIS.Count = THIS.Count + 1
  IF ALEN(THIS.Items,1) < THIS.Count
   DIMEN THIS.Items[THIS.Count]
   DIMEN THIS.Keys[THIS.Count]
  ENDIF
  THIS.Items[THIS.Count] = puValue
  THIS.Keys[THIS.Count] = IIF(EMPTY(pcKey),"",pcKey)
 ENDPROC
 
 PROCEDURE Remove(puKeyOrIndex)
  IF VARTYPE(puKeyOrIndex)="C"
   puKeyOrIndex = THIS.getKey(puKeyOrIndex)
  ENDIF
  LOCAL i
  FOR i = puKeyOrIndex TO THIS.Count - 1
   THIS.Items[i] = THIS.Items[i + 1]
   THIS.Keys[i] = THIS.Keys[i + 1]
  ENDFOR
  THIS.Items[THIS.Count] = NULL
  THIS.Keys[THIS.Count] = NULL
  THIS.Count = THIS.Count - 1
 ENDPROC

 PROCEDURE getKey(puKeyOrIndex)
  LOCAL i,uResult
  IF VARTYPE(puKeyOrIndex)="N"
   uResult = THIS.Keys[puKeyOrIndex]
  ELSE
   uResult = 0
   FOR i = 1 TO THIS.Count
    IF THIS.Keys[i] == puKeyOrIndex
     uResult = i
     EXIT
    ENDIF
   ENDFOR
  ENDIF
  RETURN uResult  
 ENDPROC

ENDDEFINE
#ENDIF


* base64Helper
* Clase helper para codificar y decodificar en formato Base64
*
* Autor: Victor Espina
*
* Adaptado a partir del codigo publicado por Anatoliy Mogylevets
* en FoxWikis: http://fox.wikis.com/wc.dll?Wiki~VfpBase64
* 
DEFINE CLASS base64Helper AS Custom
 *
 
 *-- COnstructor
 PROCEDURE Init
  *
  DECLARE INTEGER CryptBinaryToString IN Crypt32;
	STRING @pbBinary, LONG cbBinary, LONG dwFlags,;
	STRING @pszString, LONG @pcchString

  DECLARE INTEGER CryptStringToBinary IN crypt32;
	STRING @pszString, LONG cchString, LONG dwFlags,;
	STRING @pbBinary, LONG @pcbBinary,;
	LONG pdwSkip, LONG pdwFlags
  * 
 ENDPROC


 * encodeString
 * Toma un string y lo convierte en base64
 *
 PROCEDURE encodeString(pcString)
  LOCAL nFlags, nBufsize, cDst
  nFlags=1  && base64
  nBufsize=0
  CryptBinaryToString(@pcString, LEN(pcString),m.nFlags, NULL, @nBufsize)
  cDst = REPLICATE(CHR(0), m.nBufsize)
  IF CryptBinaryToString(@pcString, LEN(pcString), m.nFlags,@cDst, @nBufsize) = 0
   RETURN ""
  ENDIF
  RETURN cDst
 ENDPROC
 
 
 * decodeString
 * Toma una cadena en BAse64 y devuelve la cadena original
 *
 FUNCTION decodeString(pcB64)
  LOCAL nFlags, nBufsize, cDst
  nFlags=1  && base64
  nBufsize=0
  CryptStringToBinary(@pcB64, LEN(m.pcB64),nFlags, NULL, @nBufsize, 0,0)
  cDst = REPLICATE(CHR(0), m.nBufsize)
  IF CryptStringToBinary(@pcB64, LEN(m.pcB64),nFlags, @cDst, @nBufsize, 0,0) = 0
   RETURN ""
  ENDIF
  RETURN m.cDst
 ENDPROC 
 
 
 * encodeFile
 * Toma un archivo y lo codifica en base64
 *
 PROCEDURE encodeFile(pcFile)
  IF NOT THIS.IsFile(pcFile)
   RETURN ""
  ENDIF
  RETURN THIS.encodeString(FILETOSTR(pcFile))
 ENDPROC 
 
 
 * decodeFile
 * Toma una cadena base64, la decodifica y crea un archivo con el contenido
 *
 PROCEDURE decodeFile(pcB64, pcFile)
  LOCAL cBuff
  cBuff = THIS.decodeString(pcB64)
  STRTOFILE(cBuff, pcFile)
 ENDPROC
 *
ENDDEFINE



******************************************************
**
**               VFP 6 SUPPORT
**
******************************************************
#IF VERSION(5) < 700

* EMPTY
* Empty class
*
DEFINE CLASS EmptyObject AS Line
ENDDEFINE


* TRYCATCH.PRG
* Funciones para la implementacion de bloques TRY-CATCH en versiones
* de VFP anteriores a 8.00
*
* Autor: Victor Espina
* Fecha: May 2014
*
*
* Uso:
*
* LOCAL ex
* TRY()
*   un comando
*   IF NOEX()
*    otro comando
*   ENDIF
*   IF NOEX()
*    otro comando
*   ENDIF
*
* IF CATCH(@ex)
*   manejo de error
* ENDIF
*
* ENDTRY()
*
*
* Ejemplo:
*
* lOk = .F.
* TRY()
*   Iniciar()
*   IF NOEX()
*    Terminar()
*   ENDIF
*   lOk = NOEX()
*
* IF CATCH(@ex)
*    MESSAGEBOX(ex.Message)
* ENDIF
* ENTRY()
*
* IF lok
*  ...
* ENDIF
*

PROCEDURE TRY
 IF VARTYPE(gcTRYOnError)="U"
  PUBLIC gcTRYOnError,goTRYEx,gnTRYNestingLevel
  gnTRYNestingLevel = 0
 ENDIF
 goTRYEx = NULL
 gnTRYNestingLevel = gnTRYNestingLevel + 1
 IF gnTRYNestingLevel = 1
  gcTRYOnError = ON("ERROR")
  ON ERROR tryCatch(ERROR(), MESSAGE(), MESSAGE(1), PROGRAM(), LINENO())
 ENDIF
ENDPROC


PROCEDURE CATCH(poEx)
 IF PCOUNT() = 1 AND !ISNULL(goTRYEx)
  poEx = goTRYEx.Clone()
 ENDIF
 LOCAL lEx
 lEx = !ISNULL(goTRYEx)
 ENDTRY()
 RETURN lEx
ENDPROC

PROCEDURE ENDTRY
 IF gnTRYNestingLevel > 0
   gnTRYNestingLevel = gnTRYNestingLevel - 1
 ENDIF
 goTRYEx = NULL
 IF gnTRYNestingLevel = 0 
  IF !EMPTY(gcTRYOnError)
   ON ERROR &gcTRYOnError
  ELSE
   ON ERROR
  ENDIF
 ENDIF
ENDPROC



FUNCTION NOEX()
 RETURN ISNULL(goTRYEx)
ENDFUNC

FUNCTION THROW(pcError)
 ERROR (pcError)
ENDFUNC

PROCEDURE tryCatch(pnErrorNo, pcMessage, pcSource, pcProcedure, pnLineNo)
 goTRYEx = CREATE("_Exception")
 WITH goTRYEx
  .errorNo = pnErrorNo
  .Message = pcMessage
  .Source = pcSource
  .Procedure = pcProcedure
  .lineNo = pnLineNo
  .lineContents = pcSource
 ENDWITH
ENDPROC

DEFINE CLASS _Exception AS Custom
 errorNo = 0
 Message = ""
 Source = ""
 Procedure = ""
 lineNo = 0 
 Details = ""
 userValue = ""
 stackLevel = 0
 lineContents = ""
 

 PROCEDURE Clone
  LOCAL oEx 
  oEx = CREATEOBJECT(THIS.Class)
  oEx.errorNo = THIS.errorNo
  oEx.MEssage = THIS.Message
  oEx.Source = THIS.Source
  oEx.Procedure = THIS.Procedure
  oEx.lineNo = THIS.lineNo
  oEx.Details = THIS.Details
  oEx.stackLevel = THIS.stackLevel
  oEx.userValue = THIS.userValue
  oEx.lineContents = THIS.lineContents
  RETURN oEx
 ENDPROC
ENDDEFINE


* Collection (Class)
* Implementacion aproximada de la clase Collection de VFP8+
*
* Autor: Victor Espina
* Fecha: Octubre 2012
*
DEFINE CLASS Collection AS Custom

 DIMEN Keys[1]
 DIMEN Items[1]
 DIMEN Item[1]
 Count = 0
 
 PROCEDURE Init(pnCapacity)
  IF PCOUNT() = 0
   pnCapacity = 0
  ENDIF
  DIMEN THIS.Items[MAX(1,pnCapacity)]
  DIMEN THIS.Keys[MAX(1,pnCapacity)]
  THIS.Count = pnCapacity
 ENDPROC
  
 PROCEDURE Items_Access(nIndex1,nIndex2)
  IF VARTYPE(nIndex1) = "N"
   RETURN THIS.Items[nIndex1]
  ENDIF
  LOCAL i
  FOR i = 1 TO THIS.Count
   IF THIS.Keys[i] == nIndex1
    RETURN THIS.Items[i]
   ENDIF
  ENDFOR
 ENDPROC

 PROCEDURE Items_Assign(cNewVal,nIndex1,nIndex2)
  IF VARTYPE(nIndex1) = "N"
   THIS.Items[nIndex1] = m.cNewVal
  ELSE
   LOCAL i
   FOR i = 1 TO THIS.Count
    IF THIS.Keys[i] == nIndex1
     THIS.Items[i] = m.cNewVal
     EXIT
    ENDIF
   ENDFOR
  ENDIF 
 ENDPROC
 
 PROCEDURE Item_Access(nIndex1, nIndex2)
  RETURN THIS.Items[nIndex1]
 ENDPROC
 
 PROCEDURE Item_Assign(cNewVal, nIndex1, nIndex2)
  THIS.Items[nIndex1] = cNewVal
 ENDPROC


 PROCEDURE Clear
  DIMEN THIS.Items[1]
  DIMEN THIS.Keys[1]
  THIS.Count = 0
 ENDPROC
 
 PROCEDURE Add(puValue, pcKey)
  IF !EMPTY(pcKey) AND THIS.getKey(pcKey) > 0
   RETURN .F.
  ENDIF
  THIS.Count = THIS.Count + 1
  IF ALEN(THIS.Items,1) < THIS.Count
   DIMEN THIS.Items[THIS.Count]
   DIMEN THIS.Keys[THIS.Count]
  ENDIF
  THIS.Items[THIS.Count] = puValue
  THIS.Keys[THIS.Count] = IIF(EMPTY(pcKey),"",pcKey)
 ENDPROC
 
 PROCEDURE Remove(puKeyOrIndex)
  IF VARTYPE(puKeyOrIndex)="C"
   puKeyOrIndex = THIS.getKey(puKeyOrIndex)
  ENDIF
  LOCAL i
  FOR i = puKeyOrIndex TO THIS.Count - 1
   THIS.Items[i] = THIS.Items[i + 1]
   THIS.Keys[i] = THIS.Keys[i + 1]
  ENDFOR
  THIS.Items[THIS.Count] = NULL
  THIS.Keys[THIS.Count] = NULL
  THIS.Count = THIS.Count - 1
 ENDPROC

 PROCEDURE getKey(puKeyOrIndex)
  LOCAL i,uResult
  IF VARTYPE(puKeyOrIndex)="N"
   uResult = THIS.Keys[puKeyOrIndex]
  ELSE
   uResult = 0
   FOR i = 1 TO THIS.Count
    IF THIS.Keys[i] == puKeyOrIndex
     uResult = i
     EXIT
    ENDIF
   ENDFOR
  ENDIF
  RETURN uResult  
 ENDPROC

ENDDEFINE


* ADDPROPERTY
* Simula la funcion ADDPROPERTY existente en VFP9
*
PROCEDURE AddProperty(poObject, pcProperty, puValue)
 poObject.addProperty(pcProperty, puValue)
ENDPROC

* EVL
* Simula la funcion EVL de VFP9
*
FUNCTION EVL(puValue, puDefault)
 RETURN IIF(EMPTY(puValue), puDefault, puValue)
ENDFUNC

#ENDIF

#IF VERSION(5) > 600
FUNCTION NOEX
 RETURN .T.
ENDFUNC
#ENDIF
