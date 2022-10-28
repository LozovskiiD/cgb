fTest=CREATEOBJECT('FORMSUPL')
logCheck=.F.
dateCheck=varDtar
WITH fTest   
     .Caption='Проверка на ошибки'
     .Width=400       
     
     DO adLabMy WITH 'fTest',1,'дата проверки - ',20,20,100,2 
     DO adtBoxNew WITH 'fTest','tBoxDate',20,20,RetTxtWidth('99/99/99999'),dHeight,'dateCheck',.F.,.T.,.F.,.F.   
     .lab1.Left=(.Width-.lab1.Width-.tBoxDate.Width-10)/2
     .tBoxDate.Left=.lab1.Left+.lab1.Width+10
     DO adCheckBox WITH 'fTest','check1','исправлять автоматически',.tBoxDate.Top+.tBoxDate.Height+20,10,150,dHeight,'logCheck',0,.T.
     .check1.Left=(.Width-.check1.Width)/2 
     .lab1.Top=.tBoxDate.Top+.tBoxDate.Height-.lab1.Height+5
       
     DO addcontlabel WITH 'fTest','cont1',(.Width-RetTxtWidth('wприступитьw')*2-20)/2,.check1.Top+.check1.Height+20,RetTxtWidth('wприступитьw'),dHeight+3,'приступить','DO starttest'
     DO addcontlabel WITH 'fTest','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fTest.Release' 
     
     DO addShape WITH 'fTest',2,5,.cont1.Top,.cont1.Height,.Width-10,8
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
     DO addShape WITH 'fTest',3,.Shape2.Left,.Shape2.Top,.Shape2.Height,50,8
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.               
      DO adLabMy WITH 'fTest',25,'100%',.Shape2.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab25.Visible=.F.  
     
         
     .Height=.tBoxDate.Height+.check1.Height+.cont1.Height+80    
ENDWITH
DO pasteImage WITH 'fTest'
fTest.Show
*************************************************************************************************************************************
PROCEDURE starttest
STORE 0 TO max_rec,one_pers,pers_ch

WITH fTest
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .Shape2.Visible=.T.
     .Shape3.Visible=.T.
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape3.Width=1
ENDWITH
IF USED('curPeop')
   SELECT curPeop
   USE
ENDIF
IF USED('curTest')
   SELECT curTest
   USE
ENDIF
*IF logCheck
SELECT people
oldOrdPeop=SYS(21)
SET ORDER TO 1
SELECT datJob
oldOrd=SYS(21)
SET FILTER TO 
SET ORDER TO 7

*ENDIF
CREATE CURSOR curTest (kodpeop N(5),fio C(70),kp N(3),kd N(3),npodr C(100),ndolj C(100),textf M)
SELECT curtest
INDEX ON kodpeop TAG T5
SELECT * FROM datjob INTO CURSOR curPeop READWRITE
DELETE FOR tr=4
DELETE FOR !EMPTY(dateOut).AND.dateOut<dateCheck
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
COUNT TO max_rec
SCAN ALL
     SELECT people
     SEEK curPeop.kodpeop
     SELECT datjob
     SEEK curPeop.nid
     SELECT curtest
     APPEND BLANK 
     REPLACE kodpeop WITH curPeop.kodpeop,fio WITH curPeop.fio,kp WITH curPeop.kp,kd WITH curPeop.kd
     IF curpeop.kp=0      && не указано подразделение
        REPLACE textf WITH 'не указано подразделение'
     ENDIF 
     IF curpeop.kd=0       && не указана должность
        REPLACE textf WITH ALLTRIM(textf)+' не указана должность'
     ENDIF
     IF curpeop.namekf=0       && не указан тарифный кфт
        REPLACE textf WITH ALLTRIM(textf)+' не указан тарифный коэффициент'
     ENDIF  
     IF curpeop.tr=0       && не указаy тип работы
        REPLACE textf WITH ALLTRIM(textf)+' не указан тип работы'
        IF curpeop.kse>0.5.AND.logCheck
           REPLACE datjob.tr WITH 1                   
           REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
        ENDIF
     ENDIF 
     IF curpeop.kse=0       && не указаy объем работы
        REPLACE textf WITH ALLTRIM(textf)+' не указан объем работы'
     ENDIF 
     SELECT sprdolj
     LOCATE FOR kod=curpeop.kd
     DO CASE      
        CASE curpeop.kv=1.AND.curpeop.lkv && высшая категория
             IF curpeop.kf#kf3.OR.curpeop.namekf#namekf3
                SELECT curtest                
                REPLACE textf WITH ALLTRIM(textf)+' тарифный разряд( кфт.) '+LTRIM(STR(curpeop.kf))+'-'+LTRIM(STR(curpeop.namekf,5,2))+' не соответствует разряду(кфт.) в спр-ке должностей '+LTRIM(STR(sprdolj.kf3))+'-'+LTRIM(STR(sprdolj.namekf3,5,2))
                IF logCheck
                   REPLACE datJob.kf WITH sprdolj.kf3,datjob.namekf WITH sprdolj.namekf3
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
             ENDIF  
        CASE curpeop.kv=2.AND.curpeop.lkv && первая категория
             IF curpeop.kf#kf2.OR.curpeop.namekf#namekf2
                SELECT curtest
                REPLACE textf WITH ALLTRIM(textf)+' тарифный разряд( кфт.) '+LTRIM(STR(curpeop.kf))+'-'+LTRIM(STR(curpeop.namekf,5,2))+' не соответствует разряду(кфт.) в спр-ке должностей '+LTRIM(STR(sprdolj.kf2))+'-'+LTRIM(STR(sprdolj.namekf2,5,2))
                IF logCheck
                   REPLACE datJob.kf WITH sprdolj.kf2,datjob.namekf WITH sprdolj.namekf2
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
             ENDIF
        CASE curpeop.kv=3.AND.curpeop.lkv && вторая категория
             IF curpeop.kf#kf1.OR.curpeop.namekf#namekf1
                SELECT curtest
                REPLACE textf WITH ALLTRIM(textf)+' тарифный разряд( кфт.) '+LTRIM(STR(curpeop.kf))+'-'+LTRIM(STR(curpeop.namekf,5,2))+' не соответствует разряду(кфт.) в спр-ке должностей '+LTRIM(STR(sprdolj.kf1))+'-'+LTRIM(STR(sprdolj.namekf1,5,2))
                IF logCheck
                   REPLACE datJob.kf WITH sprdolj.kf1,datjob.namekf WITH sprdolj.namekf1
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
             ENDIF
        OTHERWISE         && без категории
             IF curpeop.kf#kf.OR.curpeop.namekf#namekf             
                SELECT curtest
                REPLACE textf WITH ALLTRIM(textf)+' тарифный разряд( кфт.) '+LTRIM(STR(curpeop.kf))+'-'+LTRIM(STR(curpeop.namekf,5,2))+' не соответствует разряду(кфт.) в спр-ке должностей '+LTRIM(STR(sprdolj.kf))+'-'+LTRIM(STR(sprdolj.namekf,5,2))
                IF logCheck
                   REPLACE datJob.kf WITH sprdolj.kf,datjob.namekf WITH sprdolj.namekf
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
             ENDIF  
     ENDCASE
     IF INLIST(curpeop.kat,1,2,5,7)
        SELECT curtest        
        DO CASE 
           CASE !curpeop.lkv.AND.curpeop.pkat#5
                REPLACE textf WITH ALLTRIM(textf)+' надбавка за категорию '+LTRIM(STR(curpeop.pkat))+'%  нужно 5%'
                IF logCheck
                   REPLACE datJob.pkat WITH 5
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
           CASE curpeop.kv=1.AND.curpeop.lkv.AND.curpeop.pkat#30
                REPLACE textf WITH ALLTRIM(textf)+' надбавка за категорию '+LTRIM(STR(curpeop.pkat))+'%  нужно 30%'
                IF logCheck
                   REPLACE datJob.pkat WITH 30
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
           CASE curpeop.kv=2.AND.curpeop.lkv.AND.curpeop.pkat#20
                REPLACE textf WITH ALLTRIM(textf)+' надбавка за категорию '+LTRIM(STR(curpeop.pkat))+'%  нужно 20%'
                IF logCheck
                   REPLACE datJob.pkat WITH 20
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
           CASE curpeop.kv=3.AND.curpeop.lkv.AND.curpeop.pkat#15
                REPLACE textf WITH ALLTRIM(textf)+' надбавка за категорию '+LTRIM(STR(curpeop.pkat))+'%  нужно 15%'     
                IF logCheck
                   REPLACE datJob.pkat WITH 15
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
           CASE curpeop.kv=0.AND.curpeop.pkat#5
                REPLACE textf WITH ALLTRIM(textf)+' надбавка за категорию '+LTRIM(STR(curpeop.pkat))+'%  нужно 5%'
                IF logCheck
                   REPLACE datJob.pkat WITH 5
                   REPLACE textf WITH ALLTRIM(textf)+' - исправлено'
                ENDIF
               
        ENDCASE         
     ENDIF
     IF curpeop.tr=1
        IF people.pkont#curpeop.pkont
           SELECT curtest
           REPLACE textf WITH ALLTRIM(textf)+' % за контракт '+LTRIM(STR(curpeop.pkont))+'%  нужно '+LTRIM(STR(people.pkont))+'%'
           IF logCheck
              REPLACE datJob.pkont WITH people.pkont             
              REPLACE textf WITH ALLTRIM(textf)+' - исправлено'              
           ENDIF
        ENDIF 
     ENDIF 
     SELECT curpeop   
     one_pers=one_pers+1
     pers_ch=one_pers/max_rec*100
     fTest.shape3.Visible=.T.
     fTest.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
     fTest.Shape3.Width=fTest.shape2.Width/100*pers_ch  
ENDSCAN
SELECT people
SET ORDER TO &oldOrdPeop
GO peoprec
SELECT datjob
SET ORDER TO &oldOrd
SELECT curtest
DELETE FOR EMPTY(textf)
COUNT TO maxf
fTest.Visible=.F.
fTest.Release
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Icon='money.ico'
     .Caption='Проверка на ошибки'
     .Width=400 
     IF maxF=0
        DO adLabMy WITH 'fSupl',1,'ошибок не обнаружено',20,20,360,2 
       DO addcontlabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('wприступитьw'))/2,.lab1.Top+.lab1.Height+10,RetTxtWidth('wприступитьw'),dHeight+3,'возврат','fSupl.Release'
       .Height=.lab1.Height+.cont1.Height+60    
     ELSE
        DO adSetupPrnToForm WITH 10,10,380,.F.,.F.   
       *---------------------------------Кнопка печать-------------------------------------------------------------------------
        DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnmistake WITH .T.' 
        *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
        DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
           .Cont1.Width,dHeight+5,'Просмотр','DO prnmistake WITH .F.'
        *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
        DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','fSupl.Release','Выход из печати'
        .Height=.Shape91.Height+.cont1.Height+60    
     ENDIF
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************************
PROCEDURE prnmistake
PARAMETERS par1
term_ch=par1
SELECT curTest
GO TOP
DO procForPrintAndPreview WITH 'repmistake','ошибки тарификации',term_ch
