CLOSE TABLES
USE setpath
pathmain=ALLTRIM(mpath)
USE datshtat IN 0
ON ERROR DO erSup
SELECT datshtat
SCAN ALL
     pathtarif=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\sprdolj.dbf' 
     USE &pathtarif EXCL IN 0
     ALTER TABLE sprdolj ADD COLUMN logsex L
     ALTER TABLE sprdolj ADD COLUMN namem C(150)
     ALTER TABLE sprdolj ADD COLUMN namerm C(150)
     ALTER TABLE sprdolj ADD COLUMN namedm C(150)
     ALTER TABLE sprdolj ADD COLUMN nametm C(150)
     SELECT sprdolj
     REPLACE namework WITH name FOR EMPTY(namework) 
     USE
     pathtarif=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\datjobdbf' 
     USE &pathtarif EXCL IN 0
     SELECT datshtat
ENDSCAN
ON ERROR 
*******************************
PROCEDURE ersup