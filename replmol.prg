CLOSE TABLES 
SET DELETED ON 
USE datshtat
cpath='e:\tar_new\cgbnew\'
SCAN ALL
     cdatpath=cpath+ALLTRIM(datshtat.pathtarif)+'\datjob.dbf'
     IF FILE(cdatpath)
        USE &cdatpath IN 0
        SELECT datjob
        ON ERROR DO erSup
        ALTER TABLE  datjob ADD COLUMN lmol L
        ON ERROR 
        USE 
     ENDIF
     SELECT datshtat
ENDSCAN 
***********************************
PROCEDURE erSup