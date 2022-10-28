IF USED('datJobItog')
   SELECT datJobItog
   USE
ENDIF
IF USED('datKonttog')
   SELECT datKontItog
   USE
ENDIF
IF USED('curItog')
   SELECT curItog
   USE
ENDIF
IF USED('curPrnTarFond')
   SELECT curPrnTarFond
   USE
ENDIF
IF USED('curKatItog')
   SELECT curKatItog
   USE
ENDIF
IF USED('curKontr')
   SELECT curKontr
   USE
ENDIF
*---Курсор для контрактов по разрядам и %
SELECT * FROM sprkoef INTO CURSOR curKontr READWRITE
ALTER TABLE curKontr ADD COLUMN kvotr N(4)      && к-во физ лиц с контрактами по разрядам
ALTER TABLE curKontr ADD COLUMN kvotrst N(10,2)  && к-во ставок в контрактами по разрядам
ALTER TABLE curKontr ADD COLUMN kvotrvac N(10,2) && к-во вакантных ставок в контрактами по разрядам
ALTER TABLE curKontr ADD COLUMN krst N(4,2)     && кф-т кратный базовой ставке
ALTER TABLE curKontr ADD COLUMN kvokr N(3)      && к-во физ лиц с контрактами по кф-там кратным базовой ставке
ALTER TABLE curKontr ADD COLUMN kvokrst N(10,2)  && к-во ставок в контрактами по кратным базовой ставке
ALTER TABLE curKontr ADD COLUMN kvokrvac N(10,2) && к-во вакантных ставок в контрактами по кратным базовой ставке
ALTER TABLE curKontr ADD COLUMN npers N(3)      && % по контракту
ALTER TABLE curKontr ADD COLUMN nperskvo N(4)   && к-во физ лиц с % по контрактам
ALTER TABLE curKontr ADD COLUMN npersst N(10,2)  && к-во ставок по контрактам


SELECT * FROM sprkat INTO CURSOR curKatItog READWRITE
ALTER TABLE curKatItog ADD COLUMN kse N(10,2)
ALTER TABLE curKatItog ADD COLUMN sumOkl N(10,2)
ALTER TABLE curKatItog ADD COLUMN avOkl N(10,2)
ALTER TABLE curKatItog ADD COLUMN kseKont N(10,2)
ALTER TABLE curKatItog ADD COLUMN sumOklKont N(10,2)
ALTER TABLE curKatItog ADD COLUMN sumKont N(10,2)
*ALTER TABLE curKatItog ADD COLUMN kse N(7,2)
SELECT curKatItog
APPEND BLANK
REPLACE kod WITH 99, name WITH 'всего'

SELECT * FROM tarfond WHERE tarfond.vac.AND.!EMPTY(tarfond.persved) INTO CURSOR curPrnTarFond READWRITE 
SELECT curPrnTarFond
INDEX ON num TAG T1
GO TOP
num_cx=0
DO WHILE !EOF()
   num_cx=num_cx+1
   REPLACE num WITH num_cx
   SKIP   
ENDDO
log_vac=.T.
log_vackont=.T.
log_kont=.F.
SELECT * FROM datjob WHERE SEEK(datjob.kodpeop,'people',1).AND.SEEK(STR(kp,3)+STR(kd,3),'rasp',2) INTO CURSOR datKontItog  READWRITE
SELECT datKontItog
ZAP
SELECT * FROM datjob WHERE SEEK(datjob.kodpeop,'people',1).AND.SEEK(STR(kp,3)+STR(kd,3),'rasp',2) INTO CURSOR datJobItog  READWRITE
IF !EMPTY(fltJob)
   SELECT datJobItog
   SET FILTER TO &fltJob
ENDIF
SELECT num,rec,sum_f,sum_fm FROM tarfond WHERE !EMPTY(sum_fm) INTO CURSOR curItog READWRITE
ALTER TABLE curItog ADD COLUMN sumStav N(12,2)
ALTER TABLE curItog ADD COLUMN sumStavKse N(12,2)
ALTER TABLE curItog ADD COLUMN sumVac N(12,2)
ALTER TABLE curItog ADD COLUMN sumVacKse N(12,2)
INDEX ON num TAG T1
SELECT curItog
APPEND BLANK
REPLACE rec WITH 'оъём',sum_fm WITH 'kse',num WITH 98
APPEND BLANK
REPLACE rec WITH 'всего',sum_fm WITH 'msf',num WITH 99

SELECT datJobItog
DELETE FOR date_in>varDTar
DELETE FOR tr=4
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SET ORDER TO 1
SELECT rasp
SCAN ALL
     rKse=rasp.kse
     SELECT datJobItog
     SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
     DO  WHILE rasp.kp=datjobItog.kp.AND.rasp.kd=datJobItog.kd.AND.!EOF()         
        * IF !vac
            rKse=rKse-datJobItog.kse
        * ENDIF    
         SKIP
     ENDDO
     IF rKse#0        
        DO CASE
           CASE rKse<=1
                kvovac=1
           CASE MOD(rKse,1)=0     
                kvovac=INT(rKse)
           CASE MOD(rKse,1)>0     
                kvovac=INT(rKse)+1    
        ENDCASE               
        kvokse=rKse  
        ksevac=0
        FOR i=1 TO kvovac
            ksevac=IIF(kvokse<=1,kvokse,1)
            SELECT datJobItog
            APPEND BLANK       
            REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kse WITH kseVac,vac WITH .T.,kat WITH rasp.kat,;
                    namekf WITH rasp.nkfvac,kf WITH rasp.kfvac     
             SELECT curPrnTarFond
             GO TOP
             DO WHILE !EOF()
                rep_r=ALLTRIM(persved)
                rep_r1='rasp.'+ALLTRIM(persved)
                SELECT datJobItog 
                REPLACE &rep_r WITH &rep_r1         
                SELECT curPrnTarFond
                SKIP
             ENDDO                                    
             SELECT datJobItog                                
             tar_ok=0
             tar_ok=varBaseSt*datJobItog.namekf      
             REPLACE tokl WITH tar_ok,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2)
             totsumf=tokl
             totsumfm=mtokl
             SELECT tarfond
             SET FILTER TO !EMPTY(countvac)
             GO TOP
             DO WHILE !EOF()
                new_sum=sum_f
                new_msum=sum_fm
                SELECT datJobItog
                r_sum=EVAL(tarfond.countvac)     
                IF !EMPTY(tarfond.sum_f) 
                   REPLACE &new_sum WITH r_sum  
                   *REPLACE &new_msum WITH IIF(tarfond.logkse,&new_sum*kse,&new_sum)     
                   REPLACE &new_msum WITH IIF(tarfond.logkse,r_sum*kse,&new_sum)     
                   totsumf=IIF(!EMPTY(tarfond.sum_f),totsumf+EVALUATE(ALLTRIM(tarfond.sum_f)),totsumf)
                   totsumfm=IIF(!EMPTY(tarfond.sum_fm),totsumfm+EVALUATE(ALLTRIM(tarfond.sum_fm)),totsumfm)
                ENDIF     
                SELECT tarfond
                SKIP
             ENDDO
             SET FILTER TO
             SELECT datJobItog
             REPLACE total WITH totsumf,msf WITH totsumfm        
             kvokse=kvokse-1
        ENDFOR  
     ENDIF
     SELECT rasp
ENDSCAN
SELECT datJobItog
DELETE FOR tokl=0
DELETE FOR date_in>varDTar
DELETE FOR tr=4
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
GO TOP
SCAN ALL
     SCATTER TO dim_itog
     SELECT datKontItog
     APPEND BLANK
     GATHER FROM dim_itog
     SELECT datJobItog
ENDSCAN
DO countKontr
DO countItog
DO countAvOklad
SELECT curItog
GO TOP
fItog=CREATEOBJECT('FORMSUPL')
WITH fItog
     .Caption='Итоги'    
     .Width=800
     .Height=600
      DO addPageFrame WITH 'fItog','pageItog',3,0,0,.Width,.Height,.T.
      WITH .pageItog
           .AddObject('mpage1','myPage') 
           .AddObject('mpage2','myPage')      
           .AddObject('mpage3','myPage')              
           WITH .mpage1
                nParent=.Parent
                .BackColor=RGB(255,255,255)
                opage1=.Parent.mPage1
                .Caption='доплаты надбавки'  
                DO procNadDopl
           ENDWITH  
           WITH .mpage2
                nParent=.Parent
                .BackColor=RGB(255,255,255)
                opage2=.Parent.mPage2
                .Caption='средние окл.+сумма окл.с контр'  
                DO procOkladAverage
           ENDWITH  
           WITH .mpage3
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opage3=.Parent.mPage3
               .Caption='распределение по тар. разрядам'  
               DO procNadKontrakt
          ENDWITH  
      ENDWITH
     .Autocenter=.T.
ENDWITH 
fItog.Show
***************************************************************************************
PROCEDURE procNadDopl
WITH opage1
    .AddObject('grdItog','GridMyNew')
     WITH .grdItog        
          .Top=0
          .Left=0
          .Width=fItog.pageItog.Width        
          .Height=.rowHeight*15
          .scrollBars=2          
          .RecordSourceType=1
          .RecordSource='curItog'
          DO addColumnToGrid WITH 'fItog.pageItog.mpage1.grdItog',6
          .Column1.ControlSource='curItog.rec'
          .Column2.ControlSource='curItog.sumStav'
          .Column3.ControlSource='curItog.sumStavKse'
          .Column4.ControlSource='curItog.sumVac'
          .Column5.ControlSource='curItog.sumVacKse'
          .Column1.Header1.Caption='доплата/надбавка'                    
          .Column2.Header1.Caption='на ставку'                    
          .Column3.Header1.Caption='на месяц'                    
          .Column4.Header1.Caption='на ставку'                    
          .Column5.Header1.Caption='на месяц'                    
          .Column2.Width=RetTxtWidth('999999999999')                    
          .Column3.Width=.Column2.Width   
          .Column4.Width=.Column2.Width   
          .Column5.Width=.Column2.Width   
          .Column2.Format='Z'
          .Column3.Format='Z'                 
          .Column4.Format='Z'
          .Column5.Format='Z'
          .SetAll('Alignment',1,'ColumnMy')
          .Column1.Alignment=0
          .Column6.Width=0
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount            
          .SetAll('BackColor',RGB(255,255,255),'ColumnMy')          
          .setAll('DynamicBackColor',"IIF(RECNO(fItog.pageItog.mPage1.grdItog.RecordSource)#fItog.pageItog.mPage1.grdItog.curRec,fItog.pageItog.mPage1.BackColor,dynBackColor)",'columnMy')          
     ENDWITH  
              
      DO adCheckBox WITH 'opage1','checkVac','учитывать вакантные',.grdItog.Top+.grdItog.Height+10,10,150,dHeight,'log_Vac',0,.T.,'DO countItog'
     .checkVac.Left=(fItog.pageItog.Width-.checkVac.Width)/2
     
     *----------------Кнопка печать-------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage1','cont1',(fItog.pageItog.Width-RetTxtWidth('WвозвратW')*2-20)/2,.checkVac.Top+.checkVac.Height+10,RetTxtWidth('WвовзратW'),dHeight+5,'печать',"DO printreport WITH 'repitog','итоги','curItog'",'печать'
     *---------------Кнопка для отказа------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage1','cont2',.cont1.Left+.cont1.Width+20,.cont1.Top,.cont1.Width,.cont1.Height,'возврат','DO exitProcitog','возврат'
     .cont1.Top=fItog.Height-.cont1.Height*2-20
     .cont2.Top=.cont1.Top
     .checkVac.Top=.cont1.Top-.checkVac.Height-10
     .grdItog.Height=.checkVac.Top-15
     
     DO gridSizeNew WITH 'fItog.pageItog.mpage1','grdItog','shapeingrid'     
  
ENDWITH
***************************************************************************************
PROCEDURE procOkladAverage
WITH opage2
     DO adtBoxAsCont WITH 'opage2','cont11',0,0,fItog.pageItog.Width/2,dHeight,'сумма тарифных окладов',2,1
     DO adtBoxAsCont WITH 'opage2','cont12',0,.cont11.Top,.cont11.Width,dHeight,'надбавка за контракт',2,1
     .AddObject('grdOklad','GridMyNew')
     WITH .grdOklad
          .Top=opage2.cont11.Top+opage2.cont11.Height-1
          .Left=0
          .Width=fItog.pageItog.Width        
          .Height=.rowHeight*(RECCOUNT('curkatitog')+1)
          .scrollBars=2          
          .RecordSourceType=1
          .RecordSource='curKatItog'       
          DO addColumnToGrid WITH 'fItog.pageItog.mpage2.grdOklad',8        
          .Column1.ControlSource='curKatItog.name'
          .Column2.ControlSource='curKatItog.kse'
          .Column3.ControlSource='curKatItog.sumOkl'
          .Column4.ControlSource='curKatItog.avOkl'
          .Column5.ControlSource='curKatItog.ksekont'
          .Column6.ControlSource='curKatItog.sumOklkont'
          .Column7.ControlSource='curKatItog.sumKont'
          
          .Column1.Header1.Caption='персонал'                    
          .Column2.Header1.Caption='к-во'                    
          .Column3.Header1.Caption='сумма'                    
          .Column4.Header1.Caption='средн.'           
          .Column5.Header1.Caption='к-во'                    
          .Column6.Header1.Caption='сумма окл.'                    
          .Column7.Header1.Caption='за контр.' 
          
                             
          .Column2.Width=RetTxtWidth('9999999999')                    
          .Column3.Width=RetTxtWidth('99999999999')   
          .Column4.Width=.Column3.Width   
          .Column5.Width=.Column2.Width   
          .Column7.Width=.Column3.Width   
          .Column8.Width=.Column3.Width   
          
          .Columns(.ColumnCount).Width=0
          .Column2.Format='Z'
          .Column3.Format='Z'                 
          .Column4.Format='Z'
          .Column5.Format='Z'
          .Column6.Format='Z'                 
          .Column7.Format='Z'
          .SetAll('Alignment',1,'ColumnMy')
          .Column1.Alignment=0 
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-SYSMETRIC(5)-13-.ColumnCount    
          .SetAll('BackColor',RGB(255,255,255),'ColumnMy')          
          .setAll('DynamicBackColor',"IIF(RECNO(fItog.pageItog.mPage2.grdOklad.RecordSource)#fItog.pageItog.mPage2.grdOklad.curRec,fItog.pageItog.mPage2.BackColor,dynBackColor)",'columnMy')                
     ENDWITH 
     .cont11.Width=.grdOklad.Column1.Width+.grdOklad.Column2.Width+.grdOklad.Column3.Width+.grdOklad.Column4.Width+15
     .cont12.Left=.cont11.Left+.cont11.width-1
     .cont12.Width=.grdOklad.Width-.cont11.Width-1
     
     DO adCheckBox WITH 'opage2','checkVac','учитывать вакантные',.grdOklad.Top+.grdOklad.Height+10,10,150,dHeight,'log_Vackont',0,.T.,'DO countAvOklad'
     .checkVac.Left=(fItog.pageItog.Width-.checkVac.Width)/2
     *----------------Кнопка печать-------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage2','cont1',(fItog.pageItog.Width-RetTxtWidth('WвозвратW')*2-20)/2,.checkVac.Top+.checkVac.Height+10,RetTxtWidth('WвовзратW'),dHeight+5,'печать',"DO printreport WITH 'repitogkat','итоги','curKatItog'",'печать'
     *---------------Кнопка для отказа------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage2','cont2',.cont1.Left+.cont1.Width+20,.cont1.Top,.cont1.Width,.cont1.Height,'возврат','DO exitProcitog','возврат'
     
     .cont1.Top=fItog.Height-.cont1.Height*2-20
     .cont2.Top=.cont1.Top
     .checkVac.Top=.cont1.Top-.checkVac.Height-10
     .grdOklad.Height=.checkVac.Top-.cont11.Height-15
      DO gridSizeNew WITH 'fItog.pageItog.mpage2','grdOklad','shapeingrid1'        
     
ENDWITH          
***************************************************************************************
PROCEDURE procNadKontrakt
WITH opage3
     .AddObject('grdOklad','GridMyNew')
     WITH .grdOklad
          .Top=0
          .Left=0
          .Width=fItog.pageItog.Width        
          .Height=.rowHeight*(RECCOUNT('curKontr')+1)
          .scrollBars=2          
          .RecordSourceType=1
          .RecordSource='curKontr'       
          DO addColumnToGrid WITH 'fItog.pageItog.mpage3.grdOklad',12        
          .Column1.ControlSource='curKontr.kod'
          .Column2.ControlSource='curKontr.kvotr'
          .Column3.ControlSource='curKontr.kvotrst'
          .Column4.ControlSource='curKontr.kvotrVac'
          .Column5.ControlSource='curKontr.krst'
          .Column6.ControlSource='curKontr.kvokr'
          .Column7.ControlSource='curKontr.kvokrst'
          .Column8.ControlSource='curKontr.kvokrvac'
          .Column9.ControlSource='curKontr.npers'
          .Column10.ControlSource='curKontr.nperskvo'
          .Column11.ControlSource='curKontr.npersst'
          
          .Column1.Header1.Caption='т.р.'                    
          .Column2.Header1.Caption='ф.лиц'                    
          .Column3.Header1.Caption='шт.'      
          .Column4.Header1.Caption='вак.'      
          .Column5.Header1.Caption='кфт/б.с'                  
          .Column6.Header1.Caption='ф.лиц'                    
          .Column7.Header1.Caption='шт.'                      
          .Column8.Header1.Caption='вак.'                  
          .Column9.Header1.Caption='% конт.'                  
          .Column10.Header1.Caption='ф.лиц'                    
          .Column11.Header1.Caption='шт.'         
          .Column1.Width=(.Width-SYSMETRIC(5))/11                          
          .SetAll('Width',.column1.Width,'columnMy')
          .Columns(.ColumnCount).Width=0      
          .Column11.Width=.Width-.column1.Width*10-SYSMETRIC(5)-13-.ColumnCount
          
          .SetAll('Format','Z','columnMy')
          .SetAll('Alignment',2,'ColumnMy')   
          .SetAll('BackColor',RGB(255,255,255),'ColumnMy')          
          .setAll('DynamicBackColor',"IIF(RECNO(fItog.pageItog.mPage3.grdOklad.RecordSource)#fItog.pageItog.mPage3.grdOklad.curRec,fItog.pageItog.mPage3.BackColor,dynBackColor)",'columnMy')         
     ENDWITH  
     
     DO adCheckBox WITH 'opage3','checkKont','только с контрактом',.grdOklad.Top+.grdOklad.Height+10,10,150,dHeight,'log_kont',0,.T.,'DO countKontr WITH .T.'
     .checkKont.Left=(fItog.pageItog.Width-.checkKont.Width)/2
     *----------------Кнопка печать-------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage3','cont1',(fItog.pageItog.Width-RetTxtWidth('WвозвратW')*2-20)/2,10,RetTxtWidth('WвовзратW'),dHeight+5,'печать',"DO printreport WITH 'repitogkont','итоги','curKontr'",'печать'
     *---------------Кнопка для отказа------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage3','cont2',.cont1.Left+.cont1.Width+20,.cont1.Top,.cont1.Width,.cont1.Height,'возврат','DO exitProcitog','возврат'          
     .cont1.Top=fItog.Height-.cont1.Height*2-20
     .cont2.Top=.cont1.Top
     .checkKont.Top=.cont1.Top-.checkKont.Height-10
     .grdOklad.Height=.checkKont.Top-15         
      DO gridSizeNew WITH 'fItog.pageItog.mpage3','grdOklad','shapeingrid1'     
     
ENDWITH       
**************************************************************************************************************************
PROCEDURE countKontr 
PARAMETERS par1
SELECT curKontr
REPLACE kvoTrSt WITH 0,kvoTr WITH 0,krst WITH 0,kvoKr WITH 0,kvoKrSt WITH 0,npers WITH 0,nperskvo WITH 0,npersst WITH 0 ALL
DELETE FOR kod=0
*----Выборка по тарифным разрядам------------------------------
*SELECT * FROM datjob WHERE SEEK(datjob.kodpeop,'people',1).AND.SEEK(STR(kp,3)+STR(kd,3),'rasp',2) INTO CURSOR datKontItog  READWRITE
SELECT datKontItog
IF !EMPTY(fltJob)
   SELECT datKontItog
   SET FILTER TO &fltJob
ENDIF
SELECT datKontItog
IF log_kont
   DELETE FOR pkont=0.AND.mkonts=0
ENDIF    
SELECT curKontr
SCAN ALL
     IF kod<20
        SELECT datKontItog
        SUM kse TO kseKontcx FOR kf=curKontr.kod.AND.!vac
        COUNT TO peopKontCx  FOR kf=curKontr.kod.AND.!vac
        SUM kse TO kseKontVac FOR kf=curKontr.kod.AND.vac
        SELECT curKontr
        REPLACE kvoTrSt WITH kseKontCx,kvoTr WITH peopKontCx,kvoTrvac WITH kseKontVac   
     ENDIF    
ENDSCAN
*----Выборка по кф-ту кратному базовой ставке------------------------------
SELECT namekf FROM datjob WHERE namekf>0.AND.!INLIST(kf,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19) DISTINCT INTO CURSOR datPersKf READWRITE
SELECT datPersKf
SCAN ALL
     SELECT datKontItog
     IF log_Kont
        SUM kse TO kseKontcx FOR namekf=datPersKf.namekf.AND.!INLIST(kf,1,2,3,4,5,6,7,8,10,11,12,13,14,15,16,17,18,19).AND.pkont>0
        COUNT TO peopKontCx FOR namekf=datPersKf.namekf.AND.!INLIST(kf,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19).AND.pkont>0
     ELSE 
        SUM kse TO kseKontcx FOR namekf=datPersKf.namekf.AND.!INLIST(kf,1,2,3,4,5,6,1,7,8,9,10,11,12,13,14,15,16,17,18,19).AND.!vac
        COUNT TO peopKontCx FOR namekf=datPersKf.namekf.AND.!INLIST(kf,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19).AND.!vac
        SUM kse TO kseKontVac FOR namekf=datPersKf.namekf.AND.!INLIST(kf,1,2,3,4,5,6,1,7,8,9,10,11,12,13,14,15,16,17,18,19).AND.vac
     ENDIF 
     SELECT curKontr
     LOCATE FOR krSt=0
     IF !FOUND()
        APPEND BLANK 
     ENDIF 
     REPLACE krst WITH datPersKf.namekf,kvoKrSt WITH kseKontCx,kvoKr WITH peopKontCx,kvoKrVac WITH kseKontVac
     SELECT datPersKf 
ENDSCAN
*----Выборка по процентам за контракт------------------------------
SELECT pkont FROM datjob DISTINCT INTO CURSOR datPersKont READWRITE

SELECT datPersKont
SCAN ALL
     SELECT datKontItog
     SUM kse TO kseKontcx FOR pkont=datPersKont.pkont
     COUNT TO peopKontCx  FOR pkont=datPersKont.pkont
     SELECT curKontr
     LOCATE FOR npers=0
     IF !FOUND()
        APPEND BLANK 
     ENDIF 
     REPLACE npers WITH datPersKont.pkont,npersSt WITH kseKontCx,nperskvo WITH peopKontCx
     SELECT datPersKont 
ENDSCAN
SELECT curKontr
SUM kvotr,kvotrst,kvokr,kvokrst,nperskvo,npersst,kvoTrVac,kvoKrVac TO kvotrch,kvotrstch,kvokrch,kvokrstch,nperskvoch,npersstch,kvotrvacch,kvokrvacch
APPEND BLANK
REPLACE kvotr WITH kvotrch,kvotrst WITH kvotrstch,kvokr WITH kvokrch,kvokrst WITH kvokrstch,nperskvo WITH nperskvoch,npersst WITH npersstch,kvotrvac WITH kvotrvacch,kvokrvac WITH kvokrvacch
GO TOP
IF par1
   fItog.pageItog.mpage3.Refresh
ENDIF

**************************************************************************************************************************
PROCEDURE countItog
SELECT curItog
REPLACE sumStav WITH 0,sumStavKse WITH 0,sumVac WITH 0,sumVacKse WITH 0 ALL
SCAN ALL
     sumForStav=sum_f
     sumForStavKse=sum_fm
     IF !EMPTY(sumForStav)
        SELECT datJobItog
        IF log_vac
           SUM &sumForStav TO sum1  
        ELSE 
           SUM &sumForStav TO sum1 FOR !vac
        ENDIF 
        SUM &sumForStav TO sum3 FOR vac
        SELECT curItog
        REPLACE sumStav WITH sum1,sumVac WITH IIF(log_vac,sum3,0)
     ENDIF 
     IF !EMPTY(sumForStavKse)
        SELECT datJobItog
        IF log_vac
           SUM &sumForStavKse TO sum2
        ELSE
           SUM &sumForStavKse TO sum2 FOR !vac
        ENDIF        
        SUM &sumForStavKse TO sum4 FOR vac
        SELECT curItog
        REPLACE sumStavKse WITH sum2,sumVacKse WITH IIF(log_vac,sum4,0)
     ENDIF      
     SELECT curItog
ENDSCAN
SELECT datJobItog
IF log_vac
   SUM total,msf TO sum1,sum11 
   SUM total,msf TO sum2,sum22 FOR vac
   
ELSE
   SUM total,msf TO sum1,sum11 FOR !vac
ENDIF
SELECT curItog
GO BOTTOM
REPLACE sumStav WITH sum1,sumVac WITH IIF(log_vac,sum2,0)
REPLACE sumStavKse WITH sum11,sumVacKse WITH IIF(log_vac,sum22,0)
GO TOP 

**************************************************************************************************************************
PROCEDURE countAvOklad
SELECT curkatItog
SCAN ALL
     SELECT datJobItog
     SUM mtOkl,kse TO mToklch,ksekat FOR IIF(log_vackont,kat=curKatItog.kod,kat=curKatItog.kod.AND.!vac)
     
     SUM mtOkl,kse,mkonts TO oklch,ksech,kontch FOR IIF(log_vackont,kat=curKatItog.kod.AND.mkonts#0,kat=curKatItog.kod.AND.!vac.AND.mkonts#0)
         
     SELECT curKatItog
     REPLACE  sumOkl WITH mToklch,kse WITH kseKat,avOkl WITH IIF(ksekat#0,sumOkl/ksekat,0),kseKont WITH ksech,sumOklKont WITH oklch,sumKont WITH kontch
ENDSCAN
SUM kse, sumOkl,kseKont,sumOklKont,sumKont TO ksech,oklCh,kseKontch,sumOklKontch,sumKontch
GO BOTTOM
REPLACE kod WITH 99, name WITH 'всего',kse WITH ksech,sumOkl WITH oklCh,avOkl WITH IIF(kse#0,sumOkl/kse,0),;
        kseKont WITH kseKontch,sumOklKont WITH sumOklKontch,sumKont WITH sumKontch
GO TOP 
**************************************************************************************************************************
PROCEDURE exitProcitog
ON ERROR DO erSup
fItog.Visible=.F.
fItog.Release
ON ERROR
