*  TEST_MANUAL.PRG
*
*  SHOW HOW TO GENERATE A PDF USING BOTH
*  GHOSTSCRIPT MODE AND XPS MODE WITH
*  MANUAL CONFIGURATION
*
CLOSE ALL
CLEAR ALL
CLEAR

ON ERROR

DO pdf
PDF.cReportMode = "LEGACY"


?"PDF v" + ALLTRIM(STR(pdf.nVersion,6,2))
?"Test Module"
?
?"Configuring GS mode..."
PDF.oGS.cPRinter = "PSPrinter"
PDF.oGS.cGSFolder = "c:\progra~2\gs\gs9.05\bin\"
??"DONE"

?"Configuring XPS mode..."
PDF.oGS.cPRinter = "PSPrinter"
PDF.oGS.cGSFolder = "c:\progra~2\gs\gs9.05\bin\"
??"DONE"

?"Checking availability..."
?"- GS mode: " + IIF(pdf.oGS.isAvailable(),"Yes","No - " + pdf.oGS.cError)
?"- XPS mode: " + IIF(pdf.oXPS.isAvailable(),"Yes","No - " + pdf.oXPS.cError)


SELECT 0
USE test ALIAS qdata

LOCAL nStart
IF pdf.oGS.isAvailable()
 pdf.cMode = "GS"
 ?"Creating PDF (" + pdf.cMode + " mode)..."
 nStart = SECONDS()
 IF NOT pdf.create("test_gs.pdf", "test.frx", "qdata")
  ??pdf.cError
 ELSE
  ??"Done!",SECONDS() - nStart
 ENDIF
ENDIF
USE IN qdata


IF pdf.oXPS.isAvailable() 
 pdf.cMode = "XPS" 
 ?"Creating PDF (" + pdf.cMode + " mode)..."
 nStart = SECONDS()
 IF NOT pdf.create("test_xps.pdf", "test.frx", "test.dbf")
  ??pdf.cError
 ELSE
  ??"Done!",SECONDS() - nStart
 ENDIF
ENDIF

?
?"Test completed"

CLOSE ALL
