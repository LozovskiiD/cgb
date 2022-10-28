SET DELETED ON
SET SYSMENU TO 
SET EXCLUSIVE OFF
SET SYSMENU AUTOMATIC
SET DATE TO GERMAN
SET CENTURY ON
SET TALK OFF
SET SAFETY OFF
PUBLIC pathtarif,pathmain,pathsupl
STORE '' TO pathtarif,pathmain,pathsupl
USE setpath
pathmain=ALLTRIM(mpath)
pathsupl=ALLTRIM(pathprob)
_SCREEN.Visible=.F.
pathtarif=pathmain+';'+pathsupl
SET PATH TO &pathtarif 
USE 
DO mylib WITH 'procmain','money.ico'
