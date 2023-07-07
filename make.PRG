* MAKE.PRG
*
* UPDATES DIST FOLDER
*
CLOSE ALL
CLEAR ALL
CLEAR

ERASE DIST\*.*
COPY FILE pdf.prg TO dist\*.*
COPY FILE pdf.bin TO dist\*.*
COPY FILE test*.* TO dist\*.*
ERASE dist\*.bak
ERASE dist\test*.xps
ERASE dist\test*.ps
ERASE dist\*.pdf
