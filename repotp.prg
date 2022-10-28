CLOSE TABLES
USE e:\tar_new\cgbnew\real\rasp SHARED
USE e:\tar_new\cgbnew\tar01.01.22\rasp SHARED IN 0 ALIAS rasp01
USE e:\tar_new\cgbnew\tar01.03.22\rasp SHARED IN 0 ALIAS rasp03
SELECT rasp
REPLACE ndotp WITH 24 ALL
SELECT rasp03
REPLACE ndotp WITH 24 ALL
SELECT rasp01
SCAN ALL
     IF SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
        SELECT rasp
        REPLACE ndatt WITH rasp01.ndatt,ndkont WITH rasp01.ndkont
        FOR i=1 TO 12
            repfete1='fete'+LTRIM(STR(i))
            repfete2='rasp01.fete'+LTRIM(STR(i))
            
            repnight1='night'+LTRIM(STR(i))
            repnight2='rasp01.night'+LTRIM(STR(i))
            
            REPLACE &repfete1 WITH &repfete2, &repnight1 WITH &repnight2
        ENDFOR 
        
     ENDIF 
     
     IF SEEK(STR(kp,3)+STR(kd,3),'rasp03',2)
        SELECT rasp03
        REPLACE ndatt WITH rasp01.ndatt,ndkont WITH rasp01.ndkont
        FOR i=1 TO 12
            repfete1='fete'+LTRIM(STR(i))
            repfete2='rasp01.fete'+LTRIM(STR(i))
            
            repnight1='night'+LTRIM(STR(i))
            repnight2='rasp01.night'+LTRIM(STR(i))
            
            REPLACE &repfete1 WITH &repfete2, &repnight1 WITH &repnight2
        ENDFOR 
        
     ENDIF 
     SELECT rasp01 
ENDSCAN 
CLOSE TABLES 
USE e:\tar_new\cgbnew\real\people SHARED
USE e:\tar_new\cgbnew\real\datjob SHARED IN 0
SELECT datjob
SCAN ALL
     IF SEEK(nidpeop,'people',4)
        REPLACE ndotp WITH people.dayotp
        IF tr=1
           REPLACE ndkont WITH people.daykont
        ENDIF 
     ENDIF 
     SELECT datjob 
ENDSCAN 
CLOSE TABLES 

USE e:\tar_new\cgbnew\tar01.03.22\people SHARED
USE e:\tar_new\cgbnew\tar01.03.22\datjob SHARED IN 0
SELECT datjob
SCAN ALL
     IF SEEK(nidpeop,'people',4)
        REPLACE ndotp WITH people.dayotp
        IF tr=1
           REPLACE ndkont WITH people.daykont
        ENDIF 
     ENDIF 
     SELECT datjob 
ENDSCAN 
*USE e:\tar_new\cgbnew\tar01.03.22\rasp SHARED IN 0 ALIAS rasp03