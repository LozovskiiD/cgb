SET DELETED ON
CLOSE TABLES
USE setpath
pathmain=ALLTRIM(mpath)
USE datshtat IN 0

ON ERROR DO erSup
SELECT datshtat
SCAN ALL
     pathcx=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\sprtime.dbf' 
     IF FILE(pathcx)
        USE &pathcx EXCL IN 0
        SELECT sprtime
        ALTER TABLE sprtime ADD COLUMN ntot N(10,2)
        REPLACE ntot WITH T1+T2+T3+T4+T5+T6+T7+T8+T9+T10+T12 ALL       
        USE         
     ENDIF 
     SELECT datshtat
ENDSCAN
*ON ERROR 
*******************************
PROCEDURE ersup