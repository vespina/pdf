*  TEST.PRG
*
*  SHOW HOW TO GENERATE A PASSWORD-PROTECTED PDF
*
CLOSE ALL
CLEAR ALL
CLEAR

ON ERROR

DO pdf


* SET PASSWORD
pdf.oOptions.cOwnerPwd = "your-password-goes-here"



?"PDF v" + ALLTRIM(STR(pdf.nVersion,6,2))
?"Password-protected PDF test"
?
?"Checking available modes..."
PDF.oGS.cPRinter = "PSPrinter"
PDF.oGS.cGSFolder = "c:\progra~2\gs\gs9.05\bin\"
PDF.cReportMode = "LEGACY"
?" - GS mode: " + IIF(pdf.oGS.isAvailable(),"Yes","No - " + pdf.oGS.cError)
?" - XPS mode: " + IIF(pdf.oXPS.isAvailable(),"Yes","No - " + pdf.oXPS.cError)
IF !pdf.oGS.isAvailable() AND !pdf.oXPS.isAvailable()
   ?"No available methods found at this time"
   RETURN
ELSE
   PDF.cMode = IIF(pdf.oXPS.isAvailable(),"XPS","GS")
   ?"Using " + PDF.cMode + " mode"
ENDIF


SELECT 0
USE test ALIAS qdata

LOCAL nStart
?"Creating protected PDF..."
nStart = SECONDS()
IF NOT pdf.create("test_protected.pdf", "test.frx", "qdata")
  ??pdf.cError
ELSE
  ??"Done! (PWD: "+pdf.oOptions.cOwnerPwd + ")",SECONDS() - nStart
ENDIF
USE IN qdata


?
?"Test completed"

CLOSE ALL
