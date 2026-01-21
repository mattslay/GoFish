LOCAL;
 lcProg AS CHARACTER,;
 lcBase AS CHARACTER

*Bestimme Pfad
lcProg = SYS(16)
lcBase = JUSTPATH(m.lcProg)

DO WHILE !EMPTY(m.lcBase)
 IF DIRECTORY(m.lcBase+'\Tools') THEN
*gefunden
  EXIT
 ENDIF &&DIRECTORY(m.lcBase+'\Tools')

 lcBase = JUSTPATH(m.lcBase)
ENDDO &&WHILE !EMPTY(m.lcBase)

DO CASE
 CASE !EMPTY(m.lcBase)
*done
 CASE DIRECTORY('E:\SE\Tools')
  lcBase = 'E:\SE'
 CASE DIRECTORY('E:\Tools')
  lcBase = 'E:'
 CASE DIRECTORY(JUSTDRIVE(m.lcProg)+'\SE\Tools')
  lcBase = JUSTDRIVE(m.lcProg)+'\SE'
 CASE DIRECTORY(JUSTDRIVE(m.lcProg)+'\Tools')
  lcBase = JUSTDRIVE(m.lcProg)
ENDCASE
  
*Menü nach Path laden (.mpr ist im Pfad)
IF !EMPTY(m.lcBase) THEN
 DO lcBase+'\TOOLS\Start_Project.prg' WITH;
  lcProg,;
  '',;
  RGB(236,94,236),;
  RGB(000,000,000),;
  'Source\Images\Fish13.ico',;
  ' - Go! Fish!'
ENDIF &&!EMPTY(m.lcBase)
