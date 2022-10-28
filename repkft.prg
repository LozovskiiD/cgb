CLOSE ALL
USE sprdolj share
USE \\Mapsoft\TARIFW_new\real\datjob IN 0 share
SELECT sprdolj
SCAN ALL
     SELECT datjob
     LOCATE FOR kd=sprdolj.kod.AND.kv=0
     IF FOUND()
        SELECT sprdolj
        REPLACE kf WITH datjob.kf,namekf WITH datjob.namekf
     ENDIF
     
     SELECT datjob
     LOCATE FOR kd=sprdolj.kod.AND.kv=1
     IF FOUND()
        SELECT sprdolj
        REPLACE kf3 WITH datjob.kf,namekf3 WITH datjob.namekf
     ENDIF
     
     SELECT datjob
     LOCATE FOR kd=sprdolj.kod.AND.kv=2
     IF FOUND()
        SELECT sprdolj
        REPLACE kf2 WITH datjob.kf,namekf2 WITH datjob.namekf
     ENDIF
     
     SELECT datjob
     LOCATE FOR kd=sprdolj.kod.AND.kv=3
     IF FOUND()
        SELECT sprdolj
        REPLACE kf1 WITH datjob.kf,namekf1 WITH datjob.namekf
     ENDIF
     
     SELECT sprdolj     
ENDSCAN