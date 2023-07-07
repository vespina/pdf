*  TEST.PRG
*
*  SHOW HOW TO GENERATE A PDF USING BOTH
*  GHOSTSCRIPT MODE AND XPS MODE AND
*  THE AUTOSETUP FEATURE
*
CLOSE ALL
CLEAR ALL
CLEAR

ON ERROR

DO pdf

?"PDF v" + ALLTRIM(STR(pdf.nVersion,6,2))
?"Test Module"
?
?"Atempting auto-setup..."
IF !pdf.autoSetup()
 ??"FAILED!"
 ?"GS:",pdf.oGS.cERror
 ?"XPS:",pdf.oXPS.cError
 RETURN
ENDIF
??"DONE"

?"GS mode: " + IIF(pdf.oGS.isAvailable(),"Yes","No - " + pdf.oGS.cError)
?"XPS mode: " + IIF(pdf.oXPS.isAvailable(),"Yes","No - " + pdf.oXPS.cError)
?


?"GS Configuration"
?"","Printer:", pdf.oGS.cPrinter
?"","GSFolder:", pdf.oGS.cGSFolder
?"","Resolution:", pdf.oGS.cResolution
?
?"XPS Configuration"
?"","Printer:", pdf.oXPS.cPrinter
?"","gxps:", pdf.oXPS.cGXPS
?"","Resolution:", pdf.oXPS.cResolution
?


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
