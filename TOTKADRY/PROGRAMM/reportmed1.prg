USE fondmed1 IN 0
fSupl=CREATEOBJECT('FORMSUPL')
daterep=DATE()
DIMENSION dim_opt(5),dimOption(3)
STORE 0 TO dim_opt
dim_opt(1)=1
STORE .F. TO dimOption
*dimOption(1) - организация
*dimOption(2) - совокупность
*dimOption(3) - подразделение
dimOption(1)=.T.
STORE '' TO fltch

IF !USED('datagrup')
   USE datagrup IN 0
ENDIF 

SELECT * FROM datagrup INTO CURSOR dopGroup READWRITE
SELECT datagrup 
USE
SELECT dopGroup
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

WITH fSupl
     .Icon='kone.ico'  
     .Caption='Отчеты'
     .ProcExit='DO exitmed1'
     .Width=400
     .Height=400     
     DO addshape WITH 'fSupl',1,10,10,150,500,8  
     DO adCheckBox WITH 'fSupl','checkTot','организация',.Shape1.Top+10,.Shape1.Left+5,150,dHeight,'dimOption(1)',0,.T.,'DO validCheckTot'          
     DO adCheckBox WITH 'fSupl','checkSov','совокупность',.checkTot.Top,.Shape1.Left,150,dHeight,'dimOption(2)',0,.T.,'DO validCheckgroup'
     DO adCheckBox WITH 'fSupl','checkPodr','подразделение',.checkTot.Top,.Shape1.Left,150,dHeight,'dimOption(3)',0,.T.,'DO validCheckpodr'
     .checkTot.Left=.Shape1.Left+(.Shape1.Width-.checkTot.Width-.checkPodr.Width-.checkSov.Width-40)/2
     .checkSov.Left=.checkTot.Left+.checkTot.Width+20
     .checkPodr.Left=.checkSov.Left+.checkSov.Width+20
     .Shape1.Height=.checkTot.Height+20
     
     DO addShape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,100,.Shape1.Width,8
     DO adtBoxAsCont WITH 'fSupl','contYear',.Shape2.Left+30,.Shape2.Top+30,RetTxtWidth('WWWсформировать отчёт заWW'),dHeight,'сформировать отчёт на',2,1     
     DO adTboxNew WITH 'fSupl','tBoxYear',.contYear.Top,.contYear.Left+.contYear.Width-1,RetTxtWidth('99/99/99999'),dHeight,'daterep',.F.,.T.,0 
     .contYear.Left=.Shape2.Left+(.Shape2.Width-.contYear.Width-.tBoxYear.Width+1)/2
     .tBoxYear.Left=.contYear.Left+.contYear.Width-1
     .Shape2.Height=.contYear.Height+60
               
     DO addShape WITH 'fSupl',3,.Shape2.Left,.Shape2.Top+.Shape2.Height+10,100,.Shape1.Width,8
     DO addOptionButton WITH 'fSupl',1,'1-медкадры',.Shape3.Top+10,.Shape3.Left+20,'dim_opt(1)',0,"DO procValOption WITH 'fSupl','dim_opt',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'отчёт-T1',.Option1.Top,.Option1.Left+.Option1.Width+10,'dim_opt(2)',0,"DO procValOption WITH 'fSupl','dim_opt',2",.T. 
    
     .Option1.Left=.Shape3.Left+(.Shape3.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20 
     
     DO addOptionButton WITH 'fSupl',3,'отчёт-T6',.Option1.Top+.Option1.Height+10,.Option1.Left+.Option1.Width+10,'dim_opt(3)',0,"DO procValOption WITH 'fSupl','dim_opt',3",.T. 
     DO addOptionButton WITH 'fSupl',4,'укомплектованность',.Option3.Top,.Option1.Left+.Option1.Width+10,'dim_opt(4)',0,"DO procValOption WITH 'fSupl','dim_opt',4",.T. 
     .Option3.Left=.Shape3.Left+(.Shape3.Width-.Option3.Width-.Option4.Width-20)/2
     .Option4.Left=.Option3.Left+.Option3.Width+20 
     
     DO addOptionButton WITH 'fSupl',5,'сведения по образованию',.Option3.Top+.Option3.Height+10,.Option1.Left+.Option1.Width+10,'dim_opt(5)',0,"DO procValOption WITH 'fSupl','dim_opt',5",.T. 
     .Option5.Left=.Shape3.Left+(.Shape3.Width-.Option5.Width)/2
     
     .Shape3.Height=.Option1.Height*3+40
     
     DO addButtonOne WITH 'fSupl','butPrn',.Shape1.Left+(.Shape1.Width-RetTxtWidth('настройкаw')*3-20)/2,.Shape3.Top+.Shape3.Height+20,'приступить','','DO createreport',39,RetTxtWidth('wнастройкаw'),'формирование отчета' 
     DO addButtonOne WITH 'fSupl','butSetup',.butPrn.Left+.butPrn.Width+10,.butPrn.Top,'настройка','','DO reportsetup',39,.butPrn.Width,'настройка' 
     DO addButtonOne WITH 'fSupl','butRet',.butSetup.Left+.butSetup.Width+10,.butPrn.Top,'возврат','','DO exitmed1',39,.butPrn.Width,'возврат' 
     DO addShape WITH 'fSupl',11,.Shape1.Left,.butPrn.Top,.butPrn.Height,.Shape1.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape1.Left,.Shape1.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.  
          
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape3.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH       
     *-----------------------------Кнопка принять---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont11',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wпринять')*2)-15)/2,.butPrn.Top,RetTxtWidth('wпринятьw'),.butPrn.Height,'принять','DO returnToStaff WITH .T.'
     .cont11.Visible=.F.
     *---------------------------------Кнопка сброс-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont12',.cont11.Left+.cont11.Width+15,.Cont11.Top,.Cont11.Width,.cont11.Height,'сброс','DO returnToStaff WITH .F.'
     .cont12.Visible=.F.
               
     .Height=.Shape1.Height+.Shape2.Height+.Shape3.Height+.butPrn.Height+70
     .Width=.Shape1.Width+20
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************
PROCEDURE exitmed1
fSupl.Release
SELECT fondmed1
USE
SELECT people
*************************************************************************************************************************
PROCEDURE validCheckTot
dimOption(1)=.T.
dimOption(2)=.F.
dimOption(3)=.F.
fSupl.checkPodr.Caption='подразделение'
fSupl.checkSov.Caption='совокупность'
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE validCheckGroup 
dimOption(1)=.F.
dimOption(2)=.T. 
dimOption(3)=.F.
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')  
     .SetAll('Visible',.F.,'MyCommandButton')    
     .cont11.Visible=.T.
     .cont12.Visible=.T.
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopGroup.otm,name'  
     .listBox1.procForClick='DO clickListGroup'
     .listBox1.procForKeyPress='DO KeyPressListGroup' 
     .cont11.procForClick='DO returnToPrnGroup WITH .T.'
     .cont12.procForClick='DO returnToPrnGroup WITH .F.'
     .checkPodr.Caption='подразделение'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListGroup
SELECT dopGroup
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*fSupl.listBox1.Refresh
*************************************************************************************************************************
PROCEDURE keyPressListGroup
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListGroip 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnGroup
PARAMETERS parRet
kvoGroup=0
IF parRet
   SELECT dopGroup
   LOCATE FOR fl
   IF FOUND()
      dimOption(2)=.T.
      dimOption(1)=.F.
      dimOption(3)=.F.
      fltCh=ALLTRIM(sostav)
   ELSE 
      fltCh=''
      dimOption(1)=.T.
      dimOption(2)=.F.
      dimOption(3)=.F.   
   ENDIF  
ELSE 
   fltCh=''  
   SELECT dopGroup
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
   dimOption(1)=.T.
   dimOption(2)=.F.
   dimOption(3)=.F.   
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel') 
     .SetAll('Visible',.T.,'MyCommandButton')    
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .listBox1.Visible=.F.  
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE validCheckPodr
dimOption(1)=.F. 
dimOption(2)=.F.
*dimOption(3)=.F.
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopPodr.otm,name'  
     .listBox1.procForClick='DO clickListPodr'
     .listBox1.procForKeyPress='DO KeyPressListPodr' 
     .cont11.procForClick='DO returnToPrnPodr WITH .T.'
     .cont12.procForClick='DO returnToPrnPodr WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListPodr
SELECT dopPodr
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListPodr
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListPodr 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnPodr
PARAMETERS parRet
kvoPodr=0
IF parRet
   SELECT dopPodr
   fltCh=''
   onlyPodr=.F.
   SCAN ALL
        IF fl 
           fltCh=fltCh+','+LTRIM(STR(kod))+','
           onlyPodr=.T.
           kvoPodr=kvoPodr+1
        ENDIF 
   ENDSCAN  
ELSE 
   fltCh='' 
   onlyPodr=.F.
   SELECT dopPodr
   REPLACE otm WITH '',fl WITH .F. ALL
   dimOption(1)=.T.
   dimOption(2)=.F.
   dimOption(3)=.F.
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .listBox1.Visible=.F.  
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
      dimOption(3)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='подразделение'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH 
*******************************************************
PROCEDURE createreport
DO CASE
   CASE dim_opt(1)=1
        DO createmed1
   CASE dim_opt(2)=1 
        DO createt1
   CASE dim_opt(3)=1 
        DO createt6 
   CASE dim_opt(4)=1 
        DO createcomp     
   CASE dim_opt(5)=1 
        DO createeduc
        
ENDCASE
*******************************************************
PROCEDURE createmed1
IF USED('peopmed')
   SELECT peopmed
   USE
ENDIF
IF USED('jobmed')
   SELECT jobmed
   USE
ENDIF

IF USED('curfondmed1')
   SELECT curfondmed1
   USE
ENDIF
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='Подготовка к формированию отчёта' 
ENDWITH
SELECT * FROM fondmed1 INTO CURSOR curfondmed1 READWRITE
SELECT curfondmed1
INDEX ON nkstr TAG T1
INDEX ON nkstr2 TAG T2
SET ORDER TO 1

SELECT * FROM people INTO CURSOR peopmed READWRITE
SELECT peopmed
APPEND FROM peopout
DELETE FOR lvn  && удаляем внешних совместителей
DELETE FOR !EMPTY(date_out).AND.date_out<=daterep && удаляем уволенных

ALTER TABLE peopmed ADD COLUMN kd N(3)
ALTER TABLE peopmed ADD COLUMN kp N(3)
ALTER TABLE peopmed ADD COLUMN kstr N(3)
ALTER TABLE peopmed ADD COLUMN kat N(2)

SELECT * FROM datjob INTO CURSOR jobmed READWRITE
SELECT jobmed
APPEND FROM datjobout
DELETE FOR tr#1   && оставляем только по основной работе
DELETE FOR dateBeg>daterep   && удаляем принятых после даты отчета
DELETE FOR !EMPTY(dateout).AND.dateOut<daterep   && удаляем уволенных до  даты отчета
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
INDEX ON nidpeop TAG T1

STORE 0 TO kvototd,kvospecd,kvostacd,kvovd,kvo1d,kvo2d,kvooldd   &&женцины врачи
STORE 0 TO kvototn,kvospecn,kvostacn,kvovn,kvo1n,kvo2n,kvooldn   &&женцины средний
STORE 0 TO kvototp,kvospecp,kvostacp,kvovp,kvo1p,kvo2p,kvooldp   &&женцины провизоры
STORE 0 TO kvototf,kvospecf,kvostacf,kvovf,kvo1f,kvo2f,kvooldf   &&женцины фармацевты

STORE 0 TO mold,dekd   && мол. спец. и декр. врачи
STORE 0 TO moln,dekn   && мол. спец. и декр. средний
STORE 0 TO molp,dekp   && мол. спец. и декр. провизоры
STORE 0 TO molf,dekf   && мол. спец. и декр. фармацевты
**Расставляем коды должностей
SELECT peopmed

SCAN ALL
     SELECT jobmed
     SEEK peopmed.nid
     IF FOUND()
        DO CASE
           CASE kse>0.5
                SELECT peopmed
                REPLACE kd WITH jobmed.kd,kat WITH jobmed.kat  
                
           CASE kse<=0.5
                kse_cx=kse
                kd_cx=kd
                SKIP
                DO CASE
                   CASE nidpeop#peopmed.nid.OR.EOF()
                        SKIP-1
                        SELECT peopmed
                        REPLACE kd WITH jobmed.kd,kat WITH jobmed.kat    
                   CASE nidpeop=peopmed.nid.AND.kse<=kse_cx
                        SKIP-1
                        SELECT peopmed
                        REPLACE kd WITH jobmed.kd,kat WITH jobmed.kat       
                   CASE nidpeop=peopmed.nid.AND.kse>kse_cx
                        SELECT peopmed
                        REPLACE kd WITH jobmed.kd,kat WITH jobmed.kat           
                             
                ENDCASE    
        ENDCASE   
     ENDIF 
     SELECT curfondmed1
     LOCATE FOR ','+LTRIM(STR(peopmed.kd))+','$msost
     IF FOUND()
        REPLACE nkvotot WITH nkvotot+1,nkvospec WITH nkvospec+1,nkvostac WITH nkvostac+1,;
                nkvov WITH nkvov+IIF(peopmed.kval=1,1,0),nkvo1 WITH nkvo1+IIF(peopmed.kval=2,1,0),nkvo2 WITH nkvo2+IIF(peopmed.kval=3,1,0),;
                nkvoold WITH nkvoold+IIF(peopmed.pens,1,0)
     ENDIF
     SELECT peopmed
     REPLACE kstr WITH curfondmed1.nkstr 
     SELECT curfondmed1
     LOCATE FOR peopmed.kstr>=kbeg.AND.peopmed.kstr<=kend
     IF FOUND()
        REPLACE kvow WITH kvow+IIF(peopmed.sex=2,1,0)
        IF !EMPTY(peopmed.age)
           yearrepmed1=0
           DO CASE
              CASE MONTH(peopmed.age)<MONTH(daterep)
                   yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)
              CASE MONTH(peopmed.age)=MONTH(daterep)     
                   yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-IIF(DAY(peopmed.age)>DAY(daterep),1,0)
              CASE MONTH(peopmed.age)>MONTH(daterep)
                   yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-1
           ENDCASE
           DO CASE
              CASE yearrepmed1<35
                   REPLACE kvo35 WITH kvo35+1,kvo35w WITH kvo35w+IIF(peopmed.sex=2,1,0) 
              CASE yearrepmed1<45
                   REPLACE kvo45 WITH kvo45+1,kvo45w WITH kvo45w+IIF(peopmed.sex=2,1,0) 
              CASE yearrepmed1<55
                   REPLACE kvo55 WITH kvo55+1,kvo55w WITH kvo55w+IIF(peopmed.sex=2,1,0) 
              CASE yearrepmed1<61
                   REPLACE kvo60 WITH kvo60+1,kvo60w WITH kvo60w+IIF(peopmed.sex=2,1,0) 
              CASE yearrepmed1>=61
                   REPLACE kvo61 WITH kvo61+1,kvo61w WITH kvo61w+IIF(peopmed.sex=2,1,0) 
               
           ENDCASE
        ENDIF
     ENDIF
     
     IF BETWEEN(peopmed.kstr,208,216)
        LOCATE FOR peopmed.kstr=nkstr        
        IF FOUND()
           REPLACE kvow WITH kvow+IIF(peopmed.sex=2,1,0)
           IF !EMPTY(peopmed.age)
              yearrepmed1=0
              DO CASE
                 CASE MONTH(peopmed.age)<MONTH(daterep)
                      yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)
                 CASE MONTH(peopmed.age)=MONTH(daterep)     
                      yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-IIF(DAY(peopmed.age)>DAY(daterep),1,0)
                 CASE MONTH(peopmed.age)>MONTH(daterep)
                      yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-1
              ENDCASE
              DO CASE
                 CASE yearrepmed1<35
                      REPLACE kvo35 WITH kvo35+1,kvo35w WITH kvo35w+IIF(peopmed.sex=2,1,0) 
                 CASE yearrepmed1<45
                      REPLACE kvo45 WITH kvo45+1,kvo45w WITH kvo45w+IIF(peopmed.sex=2,1,0) 
                 CASE yearrepmed1<55
                      REPLACE kvo55 WITH kvo55+1,kvo55w WITH kvo55w+IIF(peopmed.sex=2,1,0) 
                 CASE yearrepmed1<61
                      REPLACE kvo60 WITH kvo60+1,kvo60w WITH kvo60w+IIF(peopmed.sex=2,1,0) 
                 CASE yearrepmed1>=61
                      REPLACE kvo61 WITH kvo61+1,kvo61w WITH kvo61w+IIF(peopmed.sex=2,1,0) 
               
              ENDCASE
           ENDIF
        ENDIF        
     ENDIF
     
     SELECT peopmed
     DO CASE
        CASE kstr>=3.AND.kstr<=104
             IF sex=2
                kvototd=kvototd+1
                kvospecd=kvospecd+1
                kvostacd=kvostacd+1
                kvovd=kvovd+IIF(kval=1,1,0)
                kvo1d=kvo1d+IIF(kval=2,1,0)
                kvo2d=kvo2d+IIF(kval=3,1,0)
                kvooldd=kvooldd+IIF(pens,1,0)
             ENDIF    
        CASE kstr>=202.AND.kstr<=229
             IF sex=2
                kvototn=kvototn+1
                kvospecn=kvospecn+1
                kvostacn=kvostacn+1
                kvovn=kvovn+IIF(kval=1,1,0)
                kvo1n=kvo1n+IIF(kval=2,1,0)
                kvo2n=kvo2n+IIF(kval=3,1,0)
                kvooldn=kvooldn+IIF(pens,1,0)
             ENDIF    
        CASE kstr>=302.AND.kstr<=310
             IF sex=2
                kvototp=kvototp+1
                kvospecp=kvospecp+1
                kvostacp=kvostacp+1
                kvovp=kvovp+IIF(kval=1,1,0)
                kvo1p=kvo1p+IIF(kval=2,1,0)
                kvo2p=kvo2p+IIF(kval=3,1,0)
                kvooldp=kvooldp+IIF(pens,1,0)
             ENDIF    
        CASE kstr>=402.AND.kstr<=406
             IF sex=2
                kvototf=kvototf+1
                kvospecf=kvospecf+1
                kvostacf=kvostacf+1
                kvovf=kvovf+IIF(kval=1,1,0)
                kvo1f=kvo1f+IIF(kval=2,1,0)
                kvo2f=kvo2f+IIF(kval=3,1,0)
                kvooldf=kvooldf+IIF(pens,1,0)
             ENDIF   
     ENDCASE  
     IF mols
        mold=IIF(kat=1,mold+1,mold)
        moln=IIF(kat=2,moln+1,moln)
        molp=IIF(kat=5,molp+1,molp)
        molf=IIF(kat=7,molf+1,molf)
     ENDIF 
     IF dekotp.AND.bdekOtp<dateRep
        dekd=IIF(kat=1,dekd+1,dekd)
        dekn=IIF(kat=2,dekn+1,dekn)
        dekp=IIF(kat=5,dekp+1,dekp)
        dekf=IIF(kat=7,dekf+1,dekf)
     ENDIF 
     SELECT peopmed
ENDSCAN
SELECT curfondmed1
SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=1.AND.nkstr<=104
SEEK 1
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=4.AND.nkstr<=6
SEEK 3
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=8.AND.nkstr<=10
SEEK 7
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=12.AND.nkstr<=16
SEEK 11
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=18.AND.nkstr<=55
SEEK 17
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=57.AND.nkstr<=79
SEEK 56
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=81.AND.nkstr<=89
SEEK 80
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx


SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=91.AND.nkstr<=94
SEEK 90
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=96.AND.nkstr<=102
SEEK 95
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx
SEEK 2
REPLACE nkvotot WITH kvototd,nkvospec WITH kvospecd,nkvostac WITH kvostacd,nkvov WITH kvovd,nkvo1 WITH kvo1d,nkvo2 WITH kvo2d,nkvoold WITH kvooldd

*************Средний медперсонал
SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=202.AND.nkstr<=229
SEEK 200
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx

SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=207.AND.nkstr<=216
SEEK 207
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx
SEEK 201
REPLACE nkvotot WITH kvototn,nkvospec WITH kvospecn,nkvostac WITH kvostacn,nkvov WITH kvovn,nkvo1 WITH kvo1n,nkvo2 WITH kvo2n,nkvoold WITH kvooldn

*************Провизоры
SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=302.AND.nkstr<=310
SEEK 300
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx
SEEK 301
REPLACE nkvotot WITH kvototp,nkvospec WITH kvospecp,nkvostac WITH kvostacp,nkvov WITH kvovp,nkvo1 WITH kvo1p,nkvo2 WITH kvo2p,nkvoold WITH kvooldp
*************Фармацевты
SUM nkvotot,nkvospec,nkvostac,nkvov,nkvo1,nkvo2,nkvoold TO nkvotot_cx,nkvospec_cx,nkvostac_cx,nkvov_cx,nkvo1_cx,nkvo2_cx,nkvoold_cx FOR nkstr>=402.AND.nkstr<=406
SEEK 400
REPLACE nkvotot WITH nkvotot_cx,nkvospec WITH nkvospec_cx,nkvostac WITH nkvostac_cx,nkvov WITH nkvov_cx,nkvo1 WITH nkvo1_cx,nkvo2 WITH nkvo2_cx,nkvoold WITH nkvoold_cx
SEEK 401
REPLACE nkvotot WITH kvototf,nkvospec WITH kvospecf,nkvostac WITH kvostacf,nkvov WITH kvovf,nkvo1 WITH kvo1f,nkvo2 WITH kvo2f,nkvoold WITH kvooldf

SET ORDER TO 2
SUM nkvotot,kvow,kvo35,kvo35w,kvo45,kvo45w,kvo55,kvo55w,kvo60,kvo60w,kvo61,kvo61w TO nkvotot_cx,kvow_cx,kvo35_cx,kvo35w_cx,kvo45_cx,kvo45w_cx,kvo55_cx,kvo55w_cx,kvo60_cx,kvo60w_cx,kvo61_cx,kvo61w_cx FOR nkstr2>=123.AND.nkstr2<=134
SEEK 121
REPLACE nkvotot WITH nkvotot_cx,kvow WITH kvow_cx,kvo35 WITH kvo35_cx,kvo35w WITH kvo35w_cx,kvo45 WITH kvo45_cx,kvo45w WITH kvo45w_cx,kvo55 WITH kvo55_cx,kvo55w WITH kvo55w_cx,kvo60 WITH kvo60_cx,kvo60w WITH kvo60w_cx,kvo61 WITH kvo61_cx,kvo61w WITH kvo61w_cx

SUM kvow,kvo35,kvo35w,kvo45,kvo45w,kvo55,kvo55w,kvo60,kvo60w,kvo61,kvo61w TO kvow_cx,kvo35_cx,kvo35w_cx,kvo45_cx,kvo45w_cx,kvo55_cx,kvo55w_cx,kvo60_cx,kvo60w_cx,kvo61_cx,kvo61w_cx FOR BETWEEN(nkstr2,244,249).OR.BETWEEN(nkstr2,259,271) 
SEEK 243
REPLACE kvow WITH kvow_cx,kvo35 WITH kvo35_cx,kvo35w WITH kvo35w_cx,kvo45 WITH kvo45_cx,kvo45w WITH kvo45w_cx,kvo55 WITH kvo55_cx,kvo55w WITH kvo55w_cx,kvo60 WITH kvo60_cx,kvo60w WITH kvo60w_cx,kvo61 WITH kvo61_cx,kvo61w WITH kvo61w_cx

SET ORDER TO 1
#DEFINE wdWindowStateMaximize 1 
*pathExcel=ALLTRIM(datset.pathword)+'med.xls'
*objExcel=CREATEOBJECT('EXCEL.APPLICATION') 
*excelBook=objExcel.workBooks.Add(pathExcel) 
pathWrd=ALLTRIM(datset.pathword)+'med.doc'
objWord=CREATEOBJECT('WORD.APPLICATION')
nameDoc=objWord.Documents.Open(pathWrd) 
   

docRef=GETOBJECT('','word.basic')
STORE 0 TO max_rec,one_pers,pers_ch
max_rec=103
WITH docRef
     ntable=nameDoc.Tables(6) &&Таблица №1
     nrow=6
     SELECT curfondmed1
     SET ORDER TO 1  
     WITH ntable
          .cell(nrow,2).Range.Select 
          strNum=0      
          DO WHILE .T.
             IF !EMPTY(.cell(nrow,2).Range.Text)
                strNum=VAL(.cell(nrow,2).Range.Text)               
                SEEK strNum
                .cell(nrow,3).Range.Text=IIF(nkvotot#0,LTRIM(STR(nkvotot)),'') 
                .cell(nrow,4).Range.Text=IIF(nkvospec#0,LTRIM(STR(nkvospec)),'')    
                .cell(nrow,5).Range.Text=IIF(nkvostac#0,LTRIM(STR(nkvostac)),'') 
                .cell(nrow,8).Range.Text=IIF(nkvov#0,LTRIM(STR(nkvov)),'') 
                .cell(nrow,9).Range.Text=IIF(nkvo1#0,LTRIM(STR(nkvo1)),'')  
                .cell(nrow,10).Range.Text=IIF(nkvo2#0,LTRIM(STR(nkvo2)),'') 
                .cell(nrow,11).Range.Text=IIF(nkvoold#0,LTRIM(STR(nkvoold)),'') 
             ENDIF 
             nrow=nrow+1
             one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption='Формирование таблицы 1 - '+LTRIM(STR(pers_ch))+'%'       
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
             IF strNum=104
                EXIT
             ENDIF
          ENDDO   
   	 ENDWITH 
   	 fSupl.lab25.Caption='' 
   	 fSupl.Shape12.Width=1
   	 ntable=nameDoc.Tables(7) &&Таблица №2
   	 WITH nTable
   	      .cell(10,3).Range.Text=IIF(dekd#0,LTRIM(STR(dekd)),'')
  	      .cell(11,3).Range.Text=IIF(mold#0,LTRIM(STR(mold)),'')
   	 ENDWITH
   	    	 
     STORE 0 TO max_rec,one_pers,pers_ch
     max_rec=13
   	 ntable=nameDoc.Tables(8) &&Таблица №3
     nrow=5
     SELECT curfondmed1
     SET ORDER TO 2 
     WITH ntable
          .cell(nrow,2).Range.Select 
          strNum=0      
          DO WHILE .T.
             IF !EMPTY(.cell(nrow,2).Range.Text)
                strNum=VAL(.cell(nrow,2).Range.Text)                
                SEEK strNum              
                .cell(nrow,3).Range.Text=IIF(nkvotot#0,LTRIM(STR(nkvotot)),'')   
                .cell(nrow,4).Range.Text=IIF(kvow#0,LTRIM(STR(kvow)),'')    
                .cell(nrow,5).Range.Text=IIF(kvo35#0,LTRIM(STR(kvo35)),'') 
                .cell(nrow,6).Range.Text=IIF(kvo35w#0,LTRIM(STR(kvo35w)),'') 
                .cell(nrow,7).Range.Text=IIF(kvo45#0,LTRIM(STR(kvo45)),'') 
                .cell(nrow,8).Range.Text=IIF(kvo45w#0,LTRIM(STR(kvo45w)),'')
                .cell(nrow,9).Range.Text=IIF(kvo55#0,LTRIM(STR(kvo55)),'') 
                .cell(nrow,10).Range.Text=IIF(kvo55w#0,LTRIM(STR(kvo55w)),'')
                .cell(nrow,11).Range.Text=IIF(kvo60#0,LTRIM(STR(kvo60)),'') 
                .cell(nrow,12).Range.Text=IIF(kvo60w#0,LTRIM(STR(kvo60w)),'')
                .cell(nrow,13).Range.Text=IIF(kvo61#0,LTRIM(STR(kvo61)),'') 
                .cell(nrow,14).Range.Text=IIF(kvo61w#0,LTRIM(STR(kvo61w)),'')
             ENDIF 
             nrow=nrow+1
             one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption='Формирование таблицы 3 - '+LTRIM(STR(pers_ch))+'%'       
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
             IF strNum=134
                EXIT
             ENDIF
          ENDDO   
   	 ENDWITH
   	 fSupl.lab25.Caption='' 
   	 fSupl.Shape12.Width=1
   	    	 
   	 
   	 **средний медперсонал
   	 STORE 0 TO max_rec,one_pers,pers_ch
     max_rec=29
   	 ntable=nameDoc.Tables(9) &&Таблица №4
     nrow=6
     SELECT curfondmed1
     SET ORDER TO 1   
     WITH ntable
          .cell(nrow,2).Range.Select 
          strNum=0      
          DO WHILE .T.
             IF !EMPTY(.cell(nrow,2).Range.Text)
                strNum=VAL(.cell(nrow,2).Range.Text) 
              
                SEEK strNum
                .cell(nrow,3).Range.Text=IIF(nkvotot#0,LTRIM(STR(nkvotot)),'')   
                .cell(nrow,4).Range.Text=IIF(nkvospec#0,LTRIM(STR(nkvospec)),'')    
                .cell(nrow,5).Range.Text=IIF(nkvostac#0,LTRIM(STR(nkvostac)),'') 
                .cell(nrow,8).Range.Text=IIF(nkvov#0,LTRIM(STR(nkvov)),'') 
                .cell(nrow,9).Range.Text=IIF(nkvo1#0,LTRIM(STR(nkvo1)),'')  
                .cell(nrow,10).Range.Text=IIF(nkvo2#0,LTRIM(STR(nkvo2)),'') 
                .cell(nrow,11).Range.Text=IIF(nkvoold#0,LTRIM(STR(nkvoold)),'') 
             ENDIF 
             one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption='Формирование таблицы 4 - '+LTRIM(STR(pers_ch))+'%'   
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
             nrow=nrow+1
             IF strNum=229
                EXIT
             ENDIF
          ENDDO   
   	 ENDWITH
   	 fSupl.lab25.Caption='' 
   	 fSupl.Shape12.Width=1
   	 
   	 ntable=nameDoc.Tables(10) &&Таблица №5
   	 WITH nTable
     	  .cell(7,3).Range.Text=IIF(moln#0,LTRIM(STR(moln)),'')
   	      .cell(8,3).Range.Text=IIF(dekn#0,LTRIM(STR(dekn)),'')
  	     
   	 ENDWITH
   	 
   	 STORE 0 TO max_rec,one_pers,pers_ch
     max_rec=28
   	 ntable=nameDoc.Tables(11) &&Таблица №6
     nrow=5
     SELECT curfondmed1
     SET ORDER TO 2 
     WITH ntable
          .cell(nrow,2).Range.Select 
          strNum=0      
          DO WHILE .T.
             IF !EMPTY(.cell(nrow,2).Range.Text)
                strNum=VAL(.cell(nrow,2).Range.Text)                
                SEEK strNum              
                .cell(nrow,3).Range.Text=IIF(nkvotot#0,LTRIM(STR(nkvotot)),'')   
                .cell(nrow,4).Range.Text=IIF(kvow#0,LTRIM(STR(kvow)),'')    
                .cell(nrow,5).Range.Text=IIF(kvo35#0,LTRIM(STR(kvo35)),'') 
                .cell(nrow,6).Range.Text=IIF(kvo35w#0,LTRIM(STR(kvo35w)),'') 
                .cell(nrow,7).Range.Text=IIF(kvo45#0,LTRIM(STR(kvo45)),'') 
                .cell(nrow,8).Range.Text=IIF(kvo45w#0,LTRIM(STR(kvo45w)),'')
                .cell(nrow,9).Range.Text=IIF(kvo55#0,LTRIM(STR(kvo55)),'') 
                .cell(nrow,10).Range.Text=IIF(kvo55w#0,LTRIM(STR(kvo55w)),'')
                .cell(nrow,11).Range.Text=IIF(kvo60#0,LTRIM(STR(kvo60)),'') 
                .cell(nrow,12).Range.Text=IIF(kvo60w#0,LTRIM(STR(kvo60w)),'')
                .cell(nrow,13).Range.Text=IIF(kvo61#0,LTRIM(STR(kvo61)),'') 
                .cell(nrow,14).Range.Text=IIF(kvo61w#0,LTRIM(STR(kvo61w)),'')
             ENDIF 
             nrow=nrow+1
             one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption='Формирование таблицы 6 - '+LTRIM(STR(pers_ch))+'%'       
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
             IF strNum=271
                EXIT
             ENDIF
          ENDDO   
   	 ENDWITH
   	 fSupl.lab25.Caption='' 
   	 fSupl.Shape12.Width=1
   	 
   	 **провизоры  (таблица 7)
   	 STORE 0 TO max_rec,one_pers,pers_ch
     max_rec=11
   	 ntable=nameDoc.Tables(12)
     nrow=6
     SELECT curfondmed1
     SET ORDER TO 1   
     WITH ntable
          .cell(nrow,2).Range.Select 
          strNum=0      
          DO WHILE .T.
             IF !EMPTY(.cell(nrow,2).Range.Text)
                strNum=VAL(.cell(nrow,2).Range.Text)                
                SEEK strNum
                .cell(nrow,3).Range.Text=IIF(nkvotot#0,LTRIM(STR(nkvotot)),'')   
                .cell(nrow,4).Range.Text=IIF(nkvospec#0,LTRIM(STR(nkvospec)),'')    
                .cell(nrow,5).Range.Text=IIF(nkvostac#0,LTRIM(STR(nkvostac)),'') 
                .cell(nrow,8).Range.Text=IIF(nkvov#0,LTRIM(STR(nkvov)),'') 
                .cell(nrow,9).Range.Text=IIF(nkvo1#0,LTRIM(STR(nkvo1)),'')  
                .cell(nrow,10).Range.Text=IIF(nkvo2#0,LTRIM(STR(nkvo2)),'') 
                .cell(nrow,11).Range.Text=IIF(nkvoold#0,LTRIM(STR(nkvoold)),'') 
             ENDIF 
             one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption='Формирование таблицы 7 - '+LTRIM(STR(pers_ch))+'%'   
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
             nrow=nrow+1
             IF strNum=310
                EXIT
             ENDIF
          ENDDO   
   	 ENDWITH
   	 fSupl.lab25.Caption='' 
   	 fSupl.Shape12.Width=1
   	 
   	 ntable=nameDoc.Tables(13) &&Таблица №8
   	 WITH nTable
       	  .cell(5,3).Range.Text=IIF(molp#0,LTRIM(STR(molp)),'')
   	      .cell(6,3).Range.Text=IIF(dekp#0,LTRIM(STR(dekp)),'')  	      
   	 ENDWITH
   	 
   	 **провизоры  (таблица 9)
   	 STORE 0 TO max_rec,one_pers,pers_ch
     max_rec=7
   	 ntable=nameDoc.Tables(14)
     nrow=5
     SELECT curfondmed1
     SET ORDER TO 1   
     WITH ntable
          .cell(nrow,2).Range.Select 
          strNum=0      
          DO WHILE .T.
             IF !EMPTY(.cell(nrow,2).Range.Text)
                strNum=VAL(.cell(nrow,2).Range.Text)             
                SEEK strNum
                .cell(nrow,3).Range.Text=IIF(nkvotot#0,LTRIM(STR(nkvotot)),'')   
                .cell(nrow,4).Range.Text=IIF(nkvospec#0,LTRIM(STR(nkvospec)),'')    
                .cell(nrow,5).Range.Text=IIF(nkvostac#0,LTRIM(STR(nkvostac)),'') 
                .cell(nrow,8).Range.Text=IIF(nkvov#0,LTRIM(STR(nkvov)),'') 
                .cell(nrow,9).Range.Text=IIF(nkvo1#0,LTRIM(STR(nkvo1)),'')  
                .cell(nrow,10).Range.Text=IIF(nkvo2#0,LTRIM(STR(nkvo2)),'') 
                .cell(nrow,11).Range.Text=IIF(nkvoold#0,LTRIM(STR(nkvoold)),'') 
             ENDIF
              one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption='Формирование таблицы 9 - '+LTRIM(STR(pers_ch))+'%' 
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
             nrow=nrow+1
             IF strNum=406
                EXIT
             ENDIF 
          ENDDO   
   	 ENDWITH
   	 fSupl.lab25.Caption='' 
   	 fSupl.Shape12.Width=1
   	 
     ntable=nameDoc.Tables(15) &&Таблица №10
   	 WITH nTable
       	  .cell(3,3).Range.Text=IIF(molf#0,LTRIM(STR(molf)),'')
   	      .cell(4,3).Range.Text=IIF(dekf#0,LTRIM(STR(dekf)),'')  	      
   	 ENDWITH
   	
ENDWITH  
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH
objWord.WindowState=wdWindowStateMaximize
objWord.Visible=.T. 

**********************************************************************
*                      Отчёт Т1
**********************************************************************
PROCEDURE createt1
IF USED('peopmed')
   SELECT peopmed
   USE
ENDIF
IF USED('jobmed')
   SELECT jobmed
   USE
ENDIF
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     *.lab25.Caption='Подготовка к формированию отчёта' 
ENDWITH
SELECT * FROM people INTO CURSOR peopmed READWRITE
SELECT peopmed
APPEND FROM peopout
DELETE FOR lvn  && удаляем внешних совместителей
DELETE FOR !EMPTY(date_out).AND.date_out<=daterep && удаляем уволенных

ALTER TABLE peopmed ADD COLUMN kd N(3)
ALTER TABLE peopmed ADD COLUMN kp N(3)
ALTER TABLE peopmed ADD COLUMN kat N(2)
ALTER TABLE peopmed ADD COLUMN nksup N(1)
ALTER TABLE peopmed ADD COLUMN kse N(4,2)
INDEX ON nid TAG T1

SELECT * FROM datjob INTO CURSOR jobmed READWRITE
SELECT jobmed
APPEND FROM datjobout
*DELETE FOR tr#1   && оставляем только по основной работе
DELETE FOR dateBeg>daterep   && удаляем принятых после даты отчета
DELETE FOR !EMPTY(dateout).AND.dateOut<daterep   && удаляем уволенных до  даты отчета
INDEX ON nidpeop TAG T1

SCAN ALL 
     IF tr=1    
        SELECT peopmed     
        SEEK jobmed.nidpeop
        IF kse<jobmed.kse
           REPLACE kp WITH jobmed.kp,kd WITH jobmed.kd,kse WITH jobmed.kse,kat WITH jobmed.kat,lvn WITH .F.
        ENDIF 
     ENDIF    
     SELECT jobmed
ENDSCAN 
SELECT peopmed
DELETE FOR lvn
DELETE FOR kp=0
DELETE FOR kd=0
REPLACE nksup WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.nksup,nksup) ALL 

DIMENSION educ_rep(16)
STORE 0 TO educ_rep
SELECT peopmed
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
SCAN ALL
     educ_rep(1)=educ_rep(1)+1  
     educ_rep(2)=educ_rep(2)+IIF(INLIST(nksup,1,2,3),1,0)
     educ_rep(3)=educ_rep(3)+IIF(nkSup=1,1,0)
     educ_rep(4)=educ_rep(4)+IIF(nkSup=2,1,0)
     educ_rep(5)=educ_rep(5)+IIF(nkSup=3,1,0)
     educ_rep(6)=educ_rep(6)+IIF(nkSup=4,1,0)
     educ_rep(7)=educ_rep(7)+IIF(sex=2,1,0)
     educ_rep(8)=educ_rep(8)+IIF(dekotp,1,0)
     educ_rep(9)=educ_rep(9)+IIF(dekotp.AND.sex=1,1,0)
          
        
     IF !EMPTY(peopmed.age)
        yearrepmed1=0
        monthrepmed1=0
        dayrepmed1=0
        DO CASE
           CASE MONTH(peopmed.age)<MONTH(daterep)
                yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)
                monthrepmed1=IIF(DAY(daterep)>=DAY(peopmed.age),MONTH(age)-MONTH(daterep),MONTH(age)-MONTH(daterep)-1)
           CASE MONTH(peopmed.age)=MONTH(daterep)     
                yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-IIF(DAY(peopmed.age)>DAY(daterep),1,0)
                
*                monthrepmed1=IIF(DAY(daterep)>=DAY(peopmed.age),MONTH(age)-MONTH(daterep),MONTH(age)-MONTH(daterep)-1)
           CASE MONTH(peopmed.age)>MONTH(daterep)
                yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-1
                monthrepmed1=12-MONTH(peopmed.age)+MONTH(daterep)-IIF(DAY(daterep)>=DAY(peopmed.age),0,1)
                
               * monthrepmed1=IIF(DAY(daterep)>=DAY(peopmed.age),MONTH(age)-MONTH(daterep),MONTH(age)-MONTH(daterep)-1)
        ENDCASE
        educ_rep(12)=educ_rep(12)+IIF(yearrepmed1<31,1,0)
        educ_rep(13)=educ_rep(13)+IIF(yearrepmed1<16,1,0)
        educ_rep(14)=educ_rep(14)+IIF(BETWEEN(yearrepmed1,16,17),1,0)
        
        
        DO CASE
              CASE yearrepmed1+(monthrepmed1/10)>=57.6.AND.sex=2
                   educ_rep(15)=educ_rep(15)+1         
              CASE yearrepmed1+(monthrepmed1/10)>=62.6.AND.sex=1 
                   educ_rep(16)=educ_rep(16)+1

           ENDCASE
        ENDIF
ENDSCAN

#DEFINE wdWindowStateMaximize 1 
#DEFINE xlCenter -4108            
*#DEFINE xlLeft -4131  
#DEFINE xlRight -4152  
#DEFINE xlThin 2                  
#DEFINE xlMedium -4138            
#DEFINE xlDiagonalDown 5          
#DEFINE xlDiagonalUp 6                 
#DEFINE xlEdgeLeft 7              
#DEFINE xlEdgeTop 8               
#DEFINE xlEdgeBottom 9            
#DEFINE xlEdgeRight 10            
#DEFINE xlInsideVertical 11         
#DEFINE xlInsideHorizontal 12   

pathExcel=ALLTRIM(datset.pathword)+'t1new.xlsx'
objExcel=CREATEOBJECT('EXCEL.APPLICATION') 
excelBook=objExcel.workBooks.Add(pathExcel) 
SELECT peopmed
WITH excelBook.Sheets(1) 
     nrow_cx=6
     FOR i=1 TO 16
        .cells(nrow_cx,3).Value=IIF(educ_rep(i)#0,educ_rep(i),'')       
         nrow_cx=nrow_cx+1
     ENDFOR    
     .cells(1,1).Select
ENDWITH
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH
objExcel.WindowState=wdWindowStateMaximize
objExcel.Visible=.T. 
**********************************************************************
*                      Отчёт Т1 (старый)
**********************************************************************
PROCEDURE createt1old
IF USED('peopmed')
   SELECT peopmed
   USE
ENDIF
IF USED('jobmed')
   SELECT jobmed
   USE
ENDIF
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     *.lab25.Caption='Подготовка к формированию отчёта' 
ENDWITH
SELECT * FROM people INTO CURSOR peopmed READWRITE
SELECT peopmed
APPEND FROM peopout
DELETE FOR lvn  && удаляем внешних совместителей
DELETE FOR !EMPTY(date_out).AND.date_out<=daterep && удаляем уволенных

ALTER TABLE peopmed ADD COLUMN kd N(3)
ALTER TABLE peopmed ADD COLUMN kp N(3)
ALTER TABLE peopmed ADD COLUMN kat N(2)
ALTER TABLE peopmed ADD COLUMN nksup N(1)
ALTER TABLE peopmed ADD COLUMN kse N(4,2)
INDEX ON nid TAG T1

SELECT * FROM datjob INTO CURSOR jobmed READWRITE
SELECT jobmed
APPEND FROM datjobout
*DELETE FOR tr#1   && оставляем только по основной работе
DELETE FOR dateBeg>daterep   && удаляем принятых после даты отчета
DELETE FOR !EMPTY(dateout).AND.dateOut<daterep   && удаляем уволенных до  даты отчета
INDEX ON nidpeop TAG T1

SCAN ALL 
     IF tr=1    
        SELECT peopmed     
        SEEK jobmed.nidpeop
        IF kse<jobmed.kse
           REPLACE kp WITH jobmed.kp,kd WITH jobmed.kd,kse WITH jobmed.kse,kat WITH jobmed.kat,lvn WITH .F.
        ENDIF 
     ENDIF    
     SELECT jobmed
ENDSCAN 
SELECT peopmed
DELETE FOR lvn
DELETE FOR kp=0
DELETE FOR kd=0
REPLACE nksup WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.nksup,nksup) ALL 

DIMENSION educ_rep(19,7)
STORE 0 TO educ_rep
SELECT peopmed
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
SCAN ALL
     educ_rep(1,1)=educ_rep(1,1)+1
     educ_rep(1,2)=educ_rep(1,2)+IIF(INLIST(nksup,1,2,3),1,0)
     educ_rep(1,3)=educ_rep(1,3)+IIF(nkSup=1,1,0)
     educ_rep(1,4)=educ_rep(1,4)+IIF(nkSup=2,1,0)
     educ_rep(1,5)=educ_rep(1,5)+IIF(nkSup=3,1,0)
     educ_rep(1,6)=educ_rep(1,6)+IIF(nkSup=4,1,0)
     educ_rep(1,7)=educ_rep(1,7)+IIF(sex=2,1,0)
     
     educ_rep(18,1)=educ_rep(18,1)+IIF(sex=2,1,0)
     educ_rep(18,2)=educ_rep(18,2)+IIF(sex=2.AND.INLIST(nksup,1,2,3),1,0)
     educ_rep(18,3)=educ_rep(18,3)+IIF(sex=2.AND.nkSup=1,1,0)
     educ_rep(18,4)=educ_rep(18,4)+IIF(sex=2.AND.nkSup=2,1,0)
     educ_rep(18,5)=educ_rep(18,5)+IIF(sex=2.AND.nkSup=3,1,0)
     educ_rep(18,6)=educ_rep(18,6)+IIF(sex=2.AND.nkSup=4,1,0)
   
     
     DO CASE        
        CASE educ=5  && высшее
             educ_rep(2,1)=educ_rep(2,1)+1
             educ_rep(2,7)=IIF(sex=2,educ_rep(2,7)+1,educ_rep(2,7))
             
             educ_rep(2,2)=educ_rep(2,2)+IIF(INLIST(nksup,1,2,3),1,0)
             educ_rep(2,3)=educ_rep(2,3)+IIF(nkSup=1,1,0)
             educ_rep(2,4)=educ_rep(2,4)+IIF(nkSup=2,1,0)
             educ_rep(2,5)=educ_rep(2,5)+IIF(nkSup=3,1,0)
             educ_rep(2,6)=educ_rep(2,6)+IIF(nkSup=4,1,0)
             
        CASE educ=3  && средне-спец.
             educ_rep(3,1)=educ_rep(3,1)+1
             educ_rep(3,7)=IIF(sex=2,educ_rep(3,7)+1,educ_rep(3,7))
             
             educ_rep(3,2)=educ_rep(3,2)+IIF(INLIST(nksup,1,2,3),1,0)
             educ_rep(3,3)=educ_rep(3,3)+IIF(nkSup=1,1,0)
             educ_rep(3,4)=educ_rep(3,4)+IIF(nkSup=2,1,0)
             educ_rep(3,5)=educ_rep(3,5)+IIF(nkSup=3,1,0)
             educ_rep(3,6)=educ_rep(3,6)+IIF(nkSup=4,1,0)             
        CASE educ=6  && проф.-тех.
             educ_rep(4,1)=educ_rep(4,1)+1
             educ_rep(4,7)=IIF(sex=2,educ_rep(4,7)+1,educ_rep(4,7))             
             
             educ_rep(4,2)=educ_rep(4,2)+IIF(INLIST(nksup,1,2,3),1,0)
             educ_rep(4,3)=educ_rep(4,3)+IIF(nkSup=1,1,0)
             educ_rep(4,4)=educ_rep(4,4)+IIF(nkSup=2,1,0)
             educ_rep(4,5)=educ_rep(4,5)+IIF(nkSup=3,1,0)
             educ_rep(4,6)=educ_rep(4,6)+IIF(nkSup=4,1,0)
        CASE INLIST(educ,0,2,4)   && среднее  включены также лица с 0-м кодо и 4-неоконченное высшее
             educ_rep(5,1)=educ_rep(5,1)+1
             educ_rep(5,7)=IIF(sex=2,educ_rep(5,7)+1,educ_rep(5,7))             
             
             educ_rep(5,2)=educ_rep(5,2)+IIF(INLIST(nksup,1,2,3),1,0)
             educ_rep(5,3)=educ_rep(5,3)+IIF(nkSup=1,1,0)
             educ_rep(5,4)=educ_rep(5,4)+IIF(nkSup=2,1,0)
             educ_rep(5,5)=educ_rep(5,5)+IIF(nkSup=3,1,0)
             educ_rep(5,6)=educ_rep(5,6)+IIF(nkSup=4,1,0)
        CASE educ=1  && базовое
             educ_rep(6,1)=educ_rep(6,1)+1
             educ_rep(6,7)=IIF(sex=2,educ_rep(6,7)+1,educ_rep(6,7))             
             
             educ_rep(6,2)=educ_rep(6,2)+IIF(INLIST(nksup,1,2,3),1,0)
             educ_rep(6,3)=educ_rep(6,3)+IIF(nkSup=1,1,0)
             educ_rep(6,4)=educ_rep(6,4)+IIF(nkSup=2,1,0)
             educ_rep(6,5)=educ_rep(6,5)+IIF(nkSup=3,1,0)
             educ_rep(6,6)=educ_rep(6,6)+IIF(nkSup=4,1,0)
     ENDCASE
     IF !EMPTY(peopmed.age)
        yearrepmed1=0
        DO CASE
           CASE MONTH(peopmed.age)<MONTH(daterep)
                yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)
           CASE MONTH(peopmed.age)=MONTH(daterep)     
                yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-IIF(DAY(peopmed.age)>DAY(daterep),1,0)
           CASE MONTH(peopmed.age)>MONTH(daterep)
                yearrepmed1=YEAR(daterep)-YEAR(peopmed.age)-1
        ENDCASE
        DO CASE
              CASE yearrepmed1<16
                   educ_rep(7,1)=educ_rep(7,1)+1
                   educ_rep(7,7)=IIF(sex=2,educ_rep(7,7)+1,educ_rep(7,7))
             
                   educ_rep(7,2)=educ_rep(7,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(7,3)=educ_rep(7,3)+IIF(nkSup=1,1,0)
                   educ_rep(7,4)=educ_rep(7,4)+IIF(nkSup=2,1,0)
                   educ_rep(7,5)=educ_rep(7,5)+IIF(nkSup=3,1,0)
                   educ_rep(7,6)=educ_rep(7,6)+IIF(nkSup=4,1,0)                   
              CASE BETWEEN(yearrepmed1,16,17)
                   educ_rep(8,1)=educ_rep(8,1)+1
                   educ_rep(8,7)=IIF(sex=2,educ_rep(8,7)+1,educ_rep(8,7))
             
                   educ_rep(8,2)=educ_rep(8,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(8,3)=educ_rep(8,3)+IIF(nkSup=1,1,0)
                   educ_rep(8,4)=educ_rep(8,4)+IIF(nkSup=2,1,0)
                   educ_rep(8,5)=educ_rep(8,5)+IIF(nkSup=3,1,0)
                   educ_rep(8,6)=educ_rep(8,6)+IIF(nkSup=4,1,0)                   
              CASE BETWEEN(yearrepmed1,18,24)
                   educ_rep(9,1)=educ_rep(9,1)+1
                   educ_rep(9,7)=educ_rep(9,7)+IIF(sex=2,1,0)
             
                   educ_rep(9,2)=educ_rep(9,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(9,3)=educ_rep(9,3)+IIF(nkSup=1,1,0)
                   educ_rep(9,4)=educ_rep(9,4)+IIF(nkSup=2,1,0)
                   educ_rep(9,5)=educ_rep(9,5)+IIF(nkSup=3,1,0)
                   educ_rep(9,6)=educ_rep(9,6)+IIF(nkSup=4,1,0)                   
              CASE BETWEEN(yearrepmed1,25,29)
                   educ_rep(10,1)=educ_rep(10,1)+1
                   educ_rep(10,7)=educ_rep(10,7)+IIF(sex=2,1,0)
             
                   educ_rep(10,2)=educ_rep(10,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(10,3)=educ_rep(10,3)+IIF(nkSup=1,1,0)
                   educ_rep(10,4)=educ_rep(10,4)+IIF(nkSup=2,1,0)
                   educ_rep(10,5)=educ_rep(10,5)+IIF(nkSup=3,1,0)
                   educ_rep(10,6)=educ_rep(10,6)+IIF(nkSup=4,1,0)                   
              CASE yearrepmed1=30
                   educ_rep(11,1)=educ_rep(11,1)+1
                   educ_rep(11,7)=educ_rep(11,7)+IIF(sex=2,1,0)
             
                   educ_rep(11,2)=educ_rep(11,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(11,3)=educ_rep(11,3)+IIF(nkSup=1,1,0)
                   educ_rep(11,4)=educ_rep(11,4)+IIF(nkSup=2,1,0)
                   educ_rep(11,5)=educ_rep(11,5)+IIF(nkSup=3,1,0)
                   educ_rep(11,6)=educ_rep(11,6)+IIF(nkSup=4,1,0)                                   
              CASE yearrepmed1=31
                   educ_rep(12,1)=educ_rep(12,1)+1
                   educ_rep(12,7)=educ_rep(12,7)+IIF(sex=2,1,0)
             
                   educ_rep(12,2)=educ_rep(12,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(12,3)=educ_rep(12,3)+IIF(nkSup=1,1,0)
                   educ_rep(12,4)=educ_rep(12,4)+IIF(nkSup=2,1,0)
                   educ_rep(12,5)=educ_rep(12,5)+IIF(nkSup=3,1,0)
                   educ_rep(12,6)=educ_rep(12,6)+IIF(nkSup=4,1,0)  
              CASE BETWEEN(yearrepmed1,32,39)
                   educ_rep(13,1)=educ_rep(13,1)+1
                   educ_rep(13,7)=educ_rep(13,7)+IIF(sex=2,1,0)
             
                   educ_rep(13,2)=educ_rep(13,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(13,3)=educ_rep(13,3)+IIF(nkSup=1,1,0)
                   educ_rep(13,4)=educ_rep(13,4)+IIF(nkSup=2,1,0)
                   educ_rep(13,5)=educ_rep(13,5)+IIF(nkSup=3,1,0)
                   educ_rep(13,6)=educ_rep(13,6)+IIF(nkSup=4,1,0)                      
              CASE BETWEEN(yearrepmed1,40,49)
                   educ_rep(14,1)=educ_rep(14,1)+1
                   educ_rep(14,7)=educ_rep(14,7)+IIF(sex=2,1,0)
             
                   educ_rep(14,2)=educ_rep(14,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(14,3)=educ_rep(14,3)+IIF(nkSup=1,1,0)
                   educ_rep(14,4)=educ_rep(14,4)+IIF(nkSup=2,1,0)
                   educ_rep(14,5)=educ_rep(14,5)+IIF(nkSup=3,1,0)
                   educ_rep(14,6)=educ_rep(14,6)+IIF(nkSup=4,1,0)                                                                 
              CASE BETWEEN(yearrepmed1,50,54)
                   educ_rep(15,1)=educ_rep(15,1)+1
                   educ_rep(15,7)=educ_rep(15,7)+IIF(sex=2,1,0)
             
                   educ_rep(15,2)=educ_rep(15,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(15,3)=educ_rep(15,3)+IIF(nkSup=1,1,0)
                   educ_rep(15,4)=educ_rep(15,4)+IIF(nkSup=2,1,0)
                   educ_rep(15,5)=educ_rep(15,5)+IIF(nkSup=3,1,0)
                   educ_rep(15,6)=educ_rep(15,6)+IIF(nkSup=4,1,0)  
              CASE BETWEEN(yearrepmed1,55,59)
                   educ_rep(16,1)=educ_rep(16,1)+1
                   educ_rep(16,7)=educ_rep(16,7)+IIF(sex=2,1,0)
             
                   educ_rep(16,2)=educ_rep(16,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(16,3)=educ_rep(16,3)+IIF(nkSup=1,1,0)
                   educ_rep(16,4)=educ_rep(16,4)+IIF(nkSup=2,1,0)
                   educ_rep(16,5)=educ_rep(16,5)+IIF(nkSup=3,1,0)
                   educ_rep(16,6)=educ_rep(16,6)+IIF(nkSup=4,1,0)                                                                           
              CASE yearrepmed1>=60 
                   educ_rep(17,1)=educ_rep(17,1)+1
                   educ_rep(17,7)=educ_rep(17,7)+IIF(sex=2,1,0)
             
                   educ_rep(17,2)=educ_rep(17,2)+IIF(INLIST(nksup,1,2,3),1,0)
                   educ_rep(17,3)=educ_rep(17,3)+IIF(nkSup=1,1,0)
                   educ_rep(17,4)=educ_rep(17,4)+IIF(nkSup=2,1,0)
                   educ_rep(17,5)=educ_rep(17,5)+IIF(nkSup=3,1,0)
                   educ_rep(17,6)=educ_rep(17,6)+IIF(nkSup=4,1,0)         
           ENDCASE
        ENDIF
ENDSCAN

#DEFINE wdWindowStateMaximize 1 
#DEFINE xlCenter -4108            
*#DEFINE xlLeft -4131  
#DEFINE xlRight -4152  
#DEFINE xlThin 2                  
#DEFINE xlMedium -4138            
#DEFINE xlDiagonalDown 5          
#DEFINE xlDiagonalUp 6                 
#DEFINE xlEdgeLeft 7              
#DEFINE xlEdgeTop 8               
#DEFINE xlEdgeBottom 9            
#DEFINE xlEdgeRight 10            
#DEFINE xlInsideVertical 11         
#DEFINE xlInsideHorizontal 12   

pathExcel=ALLTRIM(datset.pathword)+'t1.xls'
objExcel=CREATEOBJECT('EXCEL.APPLICATION') 
excelBook=objExcel.workBooks.Add(pathExcel) 
SELECT peopmed
COUNT TO sumEmp
COUNT TO sumEmpW

WITH excelBook.Sheets(1)        
     .cells(11,3).Value=IIF(educ_rep(1,1)#0,educ_rep(1,1),'')
     .cells(11,4).Value=IIF(educ_rep(1,2)#0,educ_rep(1,2),'')
     .cells(11,5).Value=IIF(educ_rep(1,3)#0,educ_rep(1,3),'')
     .cells(11,6).Value=IIF(educ_rep(1,4)#0,educ_rep(1,4),'')
     .cells(11,7).Value=IIF(educ_rep(1,5)#0,educ_rep(1,5),'')
     .cells(11,8).Value=IIF(educ_rep(1,6)#0,educ_rep(1,6),'')
     .cells(11,9).Value=IIF(educ_rep(1,7)#0,educ_rep(1,7),'')
     
     nrow_cx=13
     FOR i=2 TO 6
        .cells(nrow_cx,3).Value=IIF(educ_rep(i,1)#0,educ_rep(i,1),'')
        .cells(nrow_cx,4).Value=IIF(educ_rep(i,2)#0,educ_rep(i,2),'')
        .cells(nrow_cx,5).Value=IIF(educ_rep(i,3)#0,educ_rep(i,3),'')
        .cells(nrow_cx,6).Value=IIF(educ_rep(i,4)#0,educ_rep(i,4),'')
        .cells(nrow_cx,7).Value=IIF(educ_rep(i,5)#0,educ_rep(i,5),'')
        .cells(nrow_cx,8).Value=IIF(educ_rep(i,6)#0,educ_rep(i,6),'')
        .cells(nrow_cx,9).Value=IIF(educ_rep(i,7)#0,educ_rep(i,7),'')
         nrow_cx=nrow_cx+1
     ENDFOR 
     nrow_cx=nrow_cx+1
     FOR i=7 TO 19
        .cells(nrow_cx,3).Value=IIF(educ_rep(i,1)#0,educ_rep(i,1),'')
        .cells(nrow_cx,4).Value=IIF(educ_rep(i,2)#0,educ_rep(i,2),'')
        .cells(nrow_cx,5).Value=IIF(educ_rep(i,3)#0,educ_rep(i,3),'')
        .cells(nrow_cx,6).Value=IIF(educ_rep(i,4)#0,educ_rep(i,4),'')
        .cells(nrow_cx,7).Value=IIF(educ_rep(i,5)#0,educ_rep(i,5),'')
        .cells(nrow_cx,8).Value=IIF(educ_rep(i,6)#0,educ_rep(i,6),'')
        .cells(nrow_cx,9).Value=IIF(educ_rep(i,7)#0,educ_rep(i,7),'')
         nrow_cx=nrow_cx+1
     ENDFOR   
     .cells(1,1).Select
ENDWITH
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH
objExcel.WindowState=wdWindowStateMaximize
objExcel.Visible=.T. 

**********************************************************************
PROCEDURE createcomp
IF USED('peopmed')
   SELECT peopmed
   USE
ENDIF
IF USED('jobmed')
   SELECT jobmed
   USE
ENDIF

IF USED('curfondmed1')
   SELECT curfondmed1
   USE
ENDIF
IF USED('curreprasp')
   SELECT curreprasp
   USE
ENDIF 
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='Подготовка к формированию отчёта' 
ENDWITH
SELECT * FROM rasp INTO CURSOR curreprasp READWRITE
SELECT curreprasp
REPLACE kat WITH IIF(SEEK(kd,'sprdolj',1).AND.sprdolj.lspis,1,kat) ALL 
DELETE FOR !INLIST(kat,1,2,5,7)
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
INDEX ON kd TAG T1
SELECT * FROM fondmed1 WHERE lcomp INTO CURSOR curfondmed1 READWRITE
SELECT curfondmed1
APPEND BLANK
REPLACE cname WITH 'средний медперсонал',nkstr WITH 600
APPEND BLANK
REPLACE cname WITH 'провизоры',nkstr WITH 700
APPEND BLANK
REPLACE cname WITH 'фармацевты',nkstr WITH 800
INDEX ON nkstr TAG T1
SET ORDER TO 

SELECT * FROM people INTO CURSOR peopmed READWRITE
SELECT peopmed
APPEND FROM peopout
*DELETE FOR lvn  && удаляем внешних совместителей
DELETE FOR !EMPTY(date_out).AND.date_out<=daterep && удаляем уволенных

ALTER TABLE peopmed ADD COLUMN kd N(3)
ALTER TABLE peopmed ADD COLUMN kp N(3)
ALTER TABLE peopmed ADD COLUMN kstr N(3)
ALTER TABLE peopmed ADD COLUMN kat N(2)
ALTER TABLE peopmed ADD COLUMN kse N(4,2)
ALTER TABLE peopmed ADD COLUMN tr N(1)
INDEX ON nid TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 1
SELECT * FROM datjob INTO CURSOR jobmed READWRITE
SELECT jobmed
APPEND FROM datjobout
DELETE FOR tr=4   && удаляем совмещение
DELETE FOR dateBeg>daterep   && удаляем принятых после даты отчета
DELETE FOR !EMPTY(dateout).AND.dateOut<daterep   && удаляем уволенных до  даты отчета
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) ALL
REPLACE dekotp WITH IIF(SEEK(nidpeop,'peopmed',1),peopmed.dekotp,dekotp) ALL 
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
INDEX ON nidpeop TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 1
SCAN ALL 
     IF tr=1    
        SELECT peopmed     
        SEEK jobmed.nidpeop
        IF kse<jobmed.kse
           REPLACE kp WITH jobmed.kp,kd WITH jobmed.kd,kse WITH jobmed.kse,kat WITH jobmed.kat,tr WITH 1
        ENDIF 
     ENDIF    
     SELECT jobmed
ENDSCAN 
SELECT jobmed
SET ORDER TO 2

SELECT peopmed 
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
SET ORDER TO 2

SELECT curreprasp
SCAN ALL
     SELECT jobmed
     SEEK STR(curreprasp.kp,3)+STR(curreprasp.kd,3)
     kvo_cx=0
     SCAN WHILE kp=curreprasp.kp.AND.kd=curreprasp.kd
          kvo_cx=kvo_cx+IIF(!dekOtp,kse,0)   
     ENDSCAN 
     SELECT curfondmed1
     DO CASE
     *   CASE curreprasp.kat=1 
     *        LOCATE FOR ','+LTRIM(STR(curreprasp.kd))+','$msost
        CASE curreprasp.kat=2
             LOCATE FOR nkstr=600
        CASE curreprasp.kat=5
             LOCATE FOR nkstr=700
        CASE curreprasp.kat=7
             LOCATE FOR nkstr=800  
        OTHERWISE                   
             LOCATE FOR ','+LTRIM(STR(curreprasp.kd))+','$msost
     ENDCASE         
     REPLACE nksesh WITH nksesh+curreprasp.kse,nksez WITH nksez+kvo_cx
     SELECT peopmed
     SEEK STR(curreprasp.kp,3)+STR(curreprasp.kd,3)
     kvo_cx=0
     SCAN WHILE kp=curreprasp.kp.AND.kd=curreprasp.kd
          IF tr=1
             REPLACE curfondmed1.nftot WITH curfondmed1.nftot+1
             IF pens 
                REPLACE curfondmed1.nfpens WITH curfondmed1.nfpens+1
             ENDIF
             IF dekotp.AND.bdekOtp<dateRep
                REPLACE curfondmed1.nfdo WITH curfondmed1.nfdo+1
             ENDIF
             IF kval=1
                REPLACE curfondmed1.nfv WITH curfondmed1.nfv+1
             ENDIF
             IF kval=2
                REPLACE curfondmed1.nf1 WITH curfondmed1.nf1+1
             ENDIF
             IF kval=3
                REPLACE curfondmed1.nf2 WITH curfondmed1.nf2+1
             ENDIF
          ENDIF 
     ENDSCAN 
     
     SELECT curreprasp         
ENDSCAN
SELECT curfondmed1
SET ORDER TO 1
SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,1,104)
SEEK 1
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,4,6)
SEEK 3
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,8,10)
SEEK 7
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,12,16)
SEEK 11
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,18,55)
SEEK 17
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,57,79)
SEEK 56
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,81,89)
SEEK 80
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,91,94)
SEEK 90
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SUM nksesh,nksez,nftot,nfpens,nfdo,nf2,nf1,nfv TO nksesh_cx,nksez_cx,nftot_cx,nfpens_cx,nfdo_cx,nf2_cx,nf1_cx,nfv_cx FOR BETWEEN(nkstr,96,102)
SEEK 95
REPLACE nksesh WITH nksesh_cx,nksez WITH nksez_cx,nftot WITH nftot_cx,nfpens WITH nfpens_cx,nfdo WITH nfdo_cx,nf2 WITH nf2_cx,nf1 WITH nf1_cx,nfv WITH nfv_cx

SEEK 600
*SELECT peopmed
*COUNT TO fdo_cx FOR pens.AND.kat=2
*SELECT 
SEEK 700
SEEK 800

SET ORDER TO 
#DEFINE xlCenter -4108            
#DEFINE xlLeft -4131  
#DEFINE xlRight -4152  
#DEFINE xlThin 2                  
#DEFINE xlMedium -4138            
#DEFINE xlDiagonalDown 5          
#DEFINE xlDiagonalUp 6                 
#DEFINE xlEdgeLeft 7              
#DEFINE xlEdgeTop 8               
#DEFINE xlEdgeBottom 9            
#DEFINE xlEdgeRight 10            
#DEFINE xlInsideVertical 11         
*#DEFINE xlInsideHorizontal 12        
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=40
     .Columns(2).ColumnWidth=10
     .Columns(3).ColumnWidth=10
     .Columns(4).ColumnWidth=10
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=10
     .Columns(7).ColumnWidth=10
     .Columns(8).ColumnWidth=10
     .Columns(9).ColumnWidth=10 
     
     .Range(.Cells(3,1),.Cells(6,1)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Наименование врачебных должностей по специальностям в соответствии с номенклатурой'   
     ENDWITH        
     .Range(.Cells(3,2),.Cells(6,2)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='штатные должности'              
     ENDWITH        
     
     .Range(.Cells(3,3),.Cells(6,3)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='занятые должности'
     ENDWITH        
     
     .Range(.Cells(3,4),.Cells(4,6)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Укомплектованность физическими лицами (по трудовым книжкам)'
     ENDWITH        
     
     .Range(.Cells(3,7),.Cells(3,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Категорированность специалистов'
     ENDWITH       
        
     .Range(.Cells(5,4),.Cells(6,4)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='всего (без учета совместителей)'
     ENDWITH
     
     .Range(.Cells(5,5),.Cells(5,6)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='з них'
     ENDWITH
     
     .Range(.Cells(4,7),.Cells(6,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='II'
     ENDWITH
     
     .Range(.Cells(4,8),.Cells(6,8)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='I'
     ENDWITH
     
     .Range(.Cells(4,9),.Cells(6,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='высшая'
     ENDWITH
     .cells(6,5).Value='пенс.'
     .cells(6,6).Value='д/о'                         
     numberRow=7 
     SELECT curfondmed1
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     SCAN ALL             
          .cells(numberRow,1).Value=cname
          .cells(numberRow,1).WrapText=.T.
          .cells(numberRow,2).Value=IIF(nksesh#0,nksesh,'')
          .Cells(numberRow,2).NumberFormat='0.00'  
          .cells(numberRow,3).Value=IIF(nksez#0,nksez,'')
          .Cells(numberRow,3).NumberFormat='0.00'  
          .cells(numberRow,4).Value=IIF(nftot#0,nftot,'')        
          .cells(numberRow,5).Value=IIF(nfpens#0,nfpens,'')        
          .cells(numberRow,6).Value=IIF(nfdo#0,nfdo,'')         
          .cells(numberRow,7).Value=IIF(nf2#0,nf2,'')     
          .cells(numberRow,8).Value=IIF(nf1#0,nf1,'')             
          .cells(numberRow,9).Value=IIF(nfv#0,nfv,'')       
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch    
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(3,1),.Cells(numberRow-1,9)).Select
    WITH objExcel.Selection
*  *       .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH      
    .Cells(2,1).Select
ENDWITH 
#UNDEFINE xlCenter     
#UNDEFINE xlLeft   
#UNDEFINE xlRight
#UNDEFINE xlThin                 
#UNDEFINE xlMedium             
#UNDEFINE xlDiagonalDown          
#UNDEFINE xlDiagonalUp                
#UNDEFINE xlEdgeLeft            
#UNDEFINE xlEdgeTop             
#UNDEFINE xlEdgeBottom          
#UNDEFINE xlEdgeRight           
#UNDEFINE xlInsideVertical      
#UNDEFINE xlInsideHorizontal   
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH          
*ON ERROR
objExcel.Visible=.T.
**********************************************************************
PROCEDURE createT6
IF USED('peopmed')
   SELECT peopmed
   USE
ENDIF
IF USED('jobmed')
   SELECT jobmed
   USE
ENDIF
IF !USED('datarmy')
    USE datarmy IN 0
ENDIF 
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     *.lab25.Caption='Подготовка к формированию отчёта' 
ENDWITH
SELECT * FROM people INTO CURSOR peopmed READWRITE
SELECT peopmed
APPEND FROM peopout
DELETE FOR lvn  && удаляем внешних совместителей
DELETE FOR !EMPTY(date_out).AND.date_out<=daterep && удаляем уволенных

ALTER TABLE peopmed ADD COLUMN kd N(3)
ALTER TABLE peopmed ADD COLUMN kp N(3)
ALTER TABLE peopmed ADD COLUMN kat N(2)
ALTER TABLE peopmed ADD COLUMN nksup N(1)
ALTER TABLE peopmed ADD COLUMN kse N(4,2)
ALTER TABLE peopmed ADD COLUMN nKzv N(3)
ALTER TABLE peopmed ADD COLUMN lVob L
ALTER TABLE peopmed ADD COLUMN nKf N(2)
INDEX ON nid TAG T1
SCAN ALL
     IF SEEK(nid,'datarmy',2)
        IF EMPTY(datarmy.dateSn).OR.datarmy.datesn>=daterep
           REPLACE lVob WITH .T.,nKzv WITH datarmy.kzv   
        ENDIF 
        
     ENDIF 
ENDSCAN
SELECT * FROM datjob INTO CURSOR jobmed READWRITE
SELECT jobmed
APPEND FROM datjobout
DELETE FOR dateBeg>daterep   && удаляем принятых после даты отчета
DELETE FOR !EMPTY(dateout).AND.dateOut<daterep   && удаляем уволенных до  даты отчета
INDEX ON nidpeop TAG T1
SCAN ALL 
     IF tr=1    
        SELECT peopmed     
        SEEK jobmed.nidpeop
        IF kse<jobmed.kse
           REPLACE kp WITH jobmed.kp,kd WITH jobmed.kd,kse WITH jobmed.kse,kat WITH jobmed.kat,nKf WITH jobmed.kf
        ENDIF 
     ENDIF    
     SELECT jobmed
ENDSCAN 
SELECT peopmed
REPLACE nksup WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.nksup6,nksup) ALL 

DIMENSION army_rep(24,16)
STORE 0 TO army_rep
SELECT peopmed
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
SCAN ALL     
     army_rep(1,1)=army_rep(1,1)+IIF(nksup=1,1,0)
     army_rep(1,2)=army_rep(1,2)+IIF(nksup=1.AND.lVob,1,0)
     army_rep(1,3)=army_rep(1,3)+IIF(nksup=1.AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15),1,0)
     army_rep(1,4)=army_rep(1,4)+IIF(nksup=1.AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8),1,0)
     army_rep(1,5)=army_rep(1,5)+IIF(nksup=1.AND.lVob.AND.INLIST(nkzv,0,1,2),1,0)

     army_rep(2,1)=army_rep(2,1)+IIF(INLIST(kat,1,2,5,7).AND.nksup=2,1,0)
     army_rep(2,2)=army_rep(2,2)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.nksup=2,1,0)
     army_rep(2,3)=army_rep(2,3)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15).AND.nksup=2,1,0)
     army_rep(2,4)=army_rep(2,4)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8).AND.nksup=2,1,0)
     army_rep(2,5)=army_rep(2,5)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.INLIST(nkzv,0,1,2).AND.nksup=2,1,0)
          
     army_rep(8,1)=army_rep(8,1)+IIF(INLIST(kat,1,2,5,7).AND.nksup=2,1,0)
     army_rep(8,2)=army_rep(8,2)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.nksup=2,1,0)
     army_rep(8,3)=army_rep(8,3)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15).AND.nksup=2,1,0)
     army_rep(8,4)=army_rep(8,4)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8).AND.nksup=2,1,0)
     army_rep(8,5)=army_rep(8,5)+IIF(INLIST(kat,1,2,5,7).AND.lVob.AND.INLIST(nkzv,0,1,2).AND.nksup=2,1,0)
     
     army_rep(9,1)=army_rep(9,1)+IIF(INLIST(kat,1,5).AND.nksup=2,1,0)
     army_rep(9,2)=army_rep(9,2)+IIF(INLIST(kat,1,5).AND.lVob.AND.nksup=2,1,0)
     army_rep(9,3)=army_rep(9,3)+IIF(INLIST(kat,1,5).AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15).AND.nksup=2,1,0)
     army_rep(9,4)=army_rep(9,4)+IIF(INLIST(kat,1,5).AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8).AND.nksup=2,1,0)
     army_rep(9,5)=army_rep(9,5)+IIF(INLIST(kat,1,5).AND.lVob.AND.INLIST(nkzv,0,1,2).AND.nksup=2,1,0)
          
     army_rep(10,1)=army_rep(10,1)+IIF(INLIST(kat,2,7).AND.nksup=2,1,0)
     army_rep(10,2)=army_rep(10,2)+IIF(INLIST(kat,2,7).AND.lVob.AND.nksup=2,1,0)
     army_rep(10,3)=army_rep(10,3)+IIF(INLIST(kat,2,7).AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15).AND.nksup=2,1,0)
     army_rep(10,4)=army_rep(10,4)+IIF(INLIST(kat,2,7).AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8).AND.nksup=2,1,0)
     army_rep(10,5)=army_rep(10,5)+IIF(INLIST(kat,2,7).AND.lVob.AND.INLIST(nkzv,0,1,2).AND.nksup=2,1,0)
  
     army_rep(11,1)=army_rep(11,1)+IIF(nksup=3,1,0)
     army_rep(11,2)=army_rep(11,2)+IIF(nksup=3.AND.lVob,1,0)
     army_rep(11,3)=army_rep(11,3)+IIF(nksup=3.AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15),1,0)
     army_rep(11,4)=army_rep(11,4)+IIF(nksup=3.AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8),1,0)
     army_rep(11,5)=army_rep(11,5)+IIF(nksup=3.AND.lVob.AND.INLIST(nkzv,0,1,2),1,0)
     
     army_rep(12,1)=army_rep(12,1)+IIF(nksup=4,1,0)
     army_rep(12,2)=army_rep(12,2)+IIF(nksup=4.AND.lVob,1,0)
     army_rep(12,3)=army_rep(12,3)+IIF(nksup=4.AND.lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15),1,0)
     army_rep(12,4)=army_rep(12,4)+IIF(nksup=4.AND.lVob.AND.INLIST(nkzv,3,4,5,6,7,8),1,0)
     army_rep(12,5)=army_rep(12,5)+IIF(nksup=4.AND.lVob.AND.INLIST(nkzv,0,1,2),1,0)
               
     army_rep(13,1)=army_rep(13,1)+IIF(nksup=4.and.INLIST(nKf,1,2,19,20),1,0)
     army_rep(14,1)=army_rep(14,1)+IIF(nksup=4.and.INLIST(nKf,3,4,21,22),1,0)
     army_rep(15,1)=army_rep(15,1)+IIF(nksup=4.and.!INLIST(nkf,0,1,2,3,4,19,20,21,22),1,0)
      
     army_rep(24,1)=army_rep(24,1)+1
     army_rep(24,2)=army_rep(24,2)+IIF(lVob,1,0)
     army_rep(24,3)=army_rep(24,3)+IIF(lVob.AND.INLIST(nkzv,9,10,11,12,13,14,15),1,0)
     army_rep(24,4)=army_rep(24,4)+IIF(lVob.AND.INLIST(nkzv,3,4,5,6,7,8),1,0)
     army_rep(24,5)=army_rep(24,5)+IIF(lVob.AND.INLIST(nkzv,0,1,2),1,0)
   
ENDSCAN

#DEFINE wdWindowStateMaximize 1 
#DEFINE xlCenter -4108            
*#DEFINE xlLeft -4131  
#DEFINE xlRight -4152  
#DEFINE xlThin 2                  
#DEFINE xlMedium -4138            
#DEFINE xlDiagonalDown 5          
#DEFINE xlDiagonalUp 6                 
#DEFINE xlEdgeLeft 7              
#DEFINE xlEdgeTop 8               
#DEFINE xlEdgeBottom 9            
#DEFINE xlEdgeRight 10            
#DEFINE xlInsideVertical 11         
#DEFINE xlInsideHorizontal 12   
pathExcel=ALLTRIM(datset.pathword)+'t6.xls'
objExcel=CREATEOBJECT('EXCEL.APPLICATION') 
excelBook=objExcel.workBooks.Add(pathExcel) 
SELECT peopmed
COUNT TO sumEmp
COUNT TO sumEmpW
WITH excelBook.Sheets(1)           
     nrow_cx=11
     FOR i=1 TO 24
         .cells(nrow_cx,3).Value=IIF(army_rep(i,1)#0,army_rep(i,1),'')
         .cells(nrow_cx,4).Value=IIF(army_rep(i,2)#0,army_rep(i,2),'')
         .cells(nrow_cx,5).Value=IIF(army_rep(i,3)#0,army_rep(i,3),'')
         .cells(nrow_cx,6).Value=IIF(army_rep(i,4)#0,army_rep(i,4),'')
         .cells(nrow_cx,7).Value=IIF(army_rep(i,5)#0,army_rep(i,5),'')    
         nrow_cx=nrow_cx+1
         IF nrow_cx=31
            nrow_cx=nrow_cx+1
         ENDIF
     ENDFOR    
     .Cells(2,1).Select
ENDWITH
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH
*objExcel.WindowState=wdWindowStateMaximize
objExcel.Visible=.T. 
**********************************************************************
PROCEDURE createeduc
IF USED('peopmed')
   SELECT peopmed
   USE
ENDIF
IF USED('jobmed')
   SELECT jobmed
   USE
ENDIF
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     *.lab25.Caption='Подготовка к формированию отчёта' 
ENDWITH
SELECT * FROM people INTO CURSOR peopmed READWRITE
SELECT peopmed
APPEND FROM peopout
DELETE FOR lvn  && удаляем внешних совместителей
DELETE FOR !EMPTY(date_out).AND.date_out<=daterep && удаляем уволенных

ALTER TABLE peopmed ADD COLUMN kd N(3)
ALTER TABLE peopmed ADD COLUMN kp N(3)
ALTER TABLE peopmed ADD COLUMN kat N(2)
ALTER TABLE peopmed ADD COLUMN nksup N(1)
ALTER TABLE peopmed ADD COLUMN kse N(4,2)
ALTER TABLE peopmed ADD COLUMN lspis L
ALTER TABLE peopmed ADD COLUMN lms L
INDEX ON nid TAG T1

SELECT * FROM datjob INTO CURSOR jobmed READWRITE
SELECT jobmed
APPEND FROM datjobout
*DELETE FOR tr#1   && оставляем только по основной работе
DELETE FOR dateBeg>daterep   && удаляем принятых после даты отчета
DELETE FOR !EMPTY(dateout).AND.dateOut<daterep   && удаляем уволенных до  даты отчета
INDEX ON nidpeop TAG T1

SCAN ALL 
     IF tr=1    
        SELECT peopmed     
        SEEK jobmed.nidpeop
        IF kse<jobmed.kse
           REPLACE kp WITH jobmed.kp,kd WITH jobmed.kd,kse WITH jobmed.kse,kat WITH jobmed.kat,lvn WITH .F.
        ENDIF 
     ENDIF    
     SELECT jobmed
ENDSCAN 
SELECT peopmed
DELETE FOR lvn
DELETE FOR kp=0
DELETE FOR kd=0
REPLACE nksup WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.nksup,nksup),lspis WITH sprdolj.lspis,kat WITH IIF(lspis,1,kat),lms WITH IIF(kat=2.AND.'сестра'$LOWER(sprdolj.name),.T.,.F.) ALL 
*REPLACE lspis WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.lspis,lspis),kat WITH IIF(lspis,1,kat) ALL 
DELETE FOR !INLIST(kat,1,2,4,7)

DIMENSION educ_rep(6,14),dim_str(6)
STORE 0 TO educ_rep
dim_str(1)='высшее медицинское образование'
dim_str(2)='из них руководители'
dim_str(3)='главные медицинские сестры'
dim_str(4)='среднее специальное медицинское образование'
dim_str(5)='из них заведующие (начальники)'
dim_str(6)='медицинские сестры (братья)'
SELECT peopmed
IF dimOption(2).OR.dimOption(3)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltCh)
ENDIF 
SCAN ALL
     yearch=0
     IF !EMPTY(peopmed.age)
        yearch=0
        DO CASE
           CASE MONTH(peopmed.age)<MONTH(daterep)
                yearch=YEAR(daterep)-YEAR(peopmed.age)
           CASE MONTH(peopmed.age)=MONTH(daterep)     
                yearch=YEAR(daterep)-YEAR(peopmed.age)-IIF(DAY(peopmed.age)>DAY(daterep),1,0)
           CASE MONTH(peopmed.age)>MONTH(daterep)
                yearch=YEAR(daterep)-YEAR(peopmed.age)-1
        ENDCASE
     ENDIF
            
     DO CASE
        CASE INLIST(kat,1,5) &&высшее
             educ_rep(1,1)=educ_rep(1,1)+1
             educ_rep(1,2)=educ_rep(1,2)+IIF(sex=2,1,0)
             DO CASE
                CASE yearch<31
                     educ_rep(1,3)=educ_rep(1,3)+1
                     educ_rep(1,4)=educ_rep(1,4)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,31,34)
                     educ_rep(1,5)=educ_rep(1,5)+1
                     educ_rep(1,6)=educ_rep(1,6)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,35,44)
                     educ_rep(1,7)=educ_rep(1,7)+1
                     educ_rep(1,8)=educ_rep(1,8)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,45,55)
                     educ_rep(1,9)=educ_rep(1,9)+1
                     educ_rep(1,10)=educ_rep(1,10)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,56,60)
                     educ_rep(1,11)=educ_rep(1,11)+1
                     educ_rep(1,12)=educ_rep(1,12)+IIF(sex=2,1,0)
                CASE yearch>=61
                     educ_rep(1,13)=educ_rep(1,13)+1
                     educ_rep(1,14)=educ_rep(1,14)+IIF(sex=2,1,0)
             ENDCASE  
             ***руководители
             IF lspis
                educ_rep(2,1)=educ_rep(2,1)+1
                educ_rep(2,2)=educ_rep(2,2)+IIF(sex=2,1,0)
                DO CASE
                   CASE yearch<31
                        educ_rep(2,3)=educ_rep(2,3)+1
                        educ_rep(2,4)=educ_rep(2,4)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,31,34)
                        educ_rep(2,5)=educ_rep(2,5)+1
                        educ_rep(2,6)=educ_rep(2,6)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,35,44)
                        educ_rep(2,7)=educ_rep(2,7)+1
                        educ_rep(2,8)=educ_rep(2,8)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,45,55)
                        educ_rep(2,9)=educ_rep(2,9)+1
                        educ_rep(2,10)=educ_rep(2,10)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,56,60)
                        educ_rep(2,11)=educ_rep(2,11)+1
                        educ_rep(2,12)=educ_rep(2,12)+IIF(sex=2,1,0)
                   CASE yearch>=61
                        educ_rep(2,13)=educ_rep(2,13)+1
                        educ_rep(2,14)=educ_rep(2,14)+IIF(sex=2,1,0)
                ENDCASE  
             ENDIF
             
                        
        CASE INLIST(kat,2,7) &&среднее специальное
             educ_rep(4,1)=educ_rep(4,1)+1
             educ_rep(4,2)=educ_rep(4,2)+IIF(sex=2,1,0)
             DO CASE
                CASE yearch<31
                     educ_rep(4,3)=educ_rep(4,3)+1
                     educ_rep(4,4)=educ_rep(4,4)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,31,34)
                     educ_rep(4,5)=educ_rep(4,5)+1
                     educ_rep(4,6)=educ_rep(4,6)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,35,44)
                     educ_rep(4,7)=educ_rep(4,7)+1
                     educ_rep(4,8)=educ_rep(4,8)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,45,55)
                     educ_rep(4,9)=educ_rep(4,9)+1
                     educ_rep(4,10)=educ_rep(4,10)+IIF(sex=2,1,0)
                CASE BETWEEN(yearch,56,60)
                     educ_rep(4,11)=educ_rep(4,11)+1
                     educ_rep(4,12)=educ_rep(4,12)+IIF(sex=2,1,0)
                CASE yearch>=61
                     educ_rep(4,13)=educ_rep(4,13)+1
                     educ_rep(4,14)=educ_rep(4,14)+IIF(sex=2,1,0)
             ENDCASE
             ***медсестры и братья
             IF lms
                educ_rep(6,1)=educ_rep(6,1)+1
                educ_rep(6,2)=educ_rep(6,2)+IIF(sex=2,1,0)
                DO CASE
                   CASE yearch<31
                        educ_rep(6,3)=educ_rep(6,3)+1
                        educ_rep(6,4)=educ_rep(6,4)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,31,34)
                        educ_rep(6,5)=educ_rep(6,5)+1
                        educ_rep(6,6)=educ_rep(6,6)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,35,44)
                        educ_rep(6,7)=educ_rep(6,7)+1
                        educ_rep(6,8)=educ_rep(6,8)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,45,55)
                        educ_rep(6,9)=educ_rep(6,9)+1
                        educ_rep(6,10)=educ_rep(6,10)+IIF(sex=2,1,0)
                   CASE BETWEEN(yearch,56,60)
                        educ_rep(6,11)=educ_rep(6,11)+1
                        educ_rep(6,12)=educ_rep(6,12)+IIF(sex=2,1,0)
                   CASE yearch>=61
                        educ_rep(6,13)=educ_rep(6,13)+1
                        educ_rep(6,14)=educ_rep(6,14)+IIF(sex=2,1,0)
                ENDCASE  
             ENDIF
     ENDCASE

ENDSCAN

#DEFINE wdWindowStateMaximize 1 
#DEFINE xlCenter -4108            
*#DEFINE xlLeft -4131  
#DEFINE xlRight -4152  
#DEFINE xlThin 2                  
#DEFINE xlMedium -4138            
#DEFINE xlDiagonalDown 5          
#DEFINE xlDiagonalUp 6                 
#DEFINE xlEdgeLeft 7              
#DEFINE xlEdgeTop 8               
#DEFINE xlEdgeBottom 9            
#DEFINE xlEdgeRight 10            
#DEFINE xlInsideVertical 11         
#DEFINE xlInsideHorizontal 12   

objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
SELECT peopmed

WITH excelBook.Sheets(1) 
     .Columns(1).ColumnWidth=25
     .Columns(2).ColumnWidth=8
     .Columns(3).ColumnWidth=8
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8
     .Columns(8).ColumnWidth=8
     .Columns(9).ColumnWidth=8 
     .Columns(10).ColumnWidth=8
     .Columns(11).ColumnWidth=8
     .Columns(12).ColumnWidth=8
     .Columns(13).ColumnWidth=8
     .Columns(14).ColumnWidth=8
     .Columns(15).ColumnWidth=8
     .Columns(16).ColumnWidth=8
     
     .Range(.Cells(2,1),.Cells(3,1)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value=''   
     ENDWITH               
     
     .Range(.Cells(2,2),.Cells(3,2)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='строка'   
     ENDWITH            
     
     .Range(.Cells(2,3),.Cells(3,3)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='всего'   
     ENDWITH           
     
    .Range(.Cells(2,4),.Cells(3,4)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='женщин'   
     ENDWITH                          
     
     .Range(.Cells(2,5),.Cells(2,6)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='до 31'   
     ENDWITH               
     .Range(.Cells(2,7),.Cells(2,8)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='31-34'   
     ENDWITH               
     .Range(.Cells(2,9),.Cells(2,10)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='35-44'   
     ENDWITH                    
    .Range(.Cells(2,11),.Cells(2,12)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='45-55'   
     ENDWITH               
    .Range(.Cells(2,13),.Cells(2,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='56-60'   
     ENDWITH           
    .Range(.Cells(2,15),.Cells(2,16)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='61 и старше'   
     ENDWITH                     
     
     .cells(3,5).Value='всего'
     .cells(3,6).Value='женщин'
     .cells(3,7).Value='всего'
     .cells(3,8).Value='женщин'
     .cells(3,9).Value='всего'
     .cells(3,10).Value='женщин'
     .cells(3,11).Value='всего'
     .cells(3,12).Value='женщин'
     .cells(3,13).Value='всего'
     .cells(3,14).Value='женщин'
     .cells(3,15).Value='всего'
     .cells(3,16).Value='женщин'          
     nrow_cx=4     
     ncolcx=3
     FOR ih=1 TO 6
         ncolcx=3
        .cells(nrow_cx,1).Value=dim_str(ih)
         FOR i=1 TO 14
             .cells(nrow_cx,ncolcx).Value=IIF(educ_rep(ih,i)#0,educ_rep(ih,i),'')
             ncolcx=ncolcx+1
         ENDFOR 
         nrow_cx=nrow_cx+1
     ENDFOR
    .Range(.Cells(2,1),.Cells(nrow_cx-1,16)).Select
    WITH objExcel.Selection
*  *       .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .WrapText=.T.
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH         
    .cells(1,1).Select
ENDWITH
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH
objExcel.WindowState=wdWindowStateMaximize
objExcel.Visible=.T. 
**********************************************************************
PROCEDURE reportsetup
DO CASE
   CASE dim_opt(1)=1
        DO setupmed1
   CASE dim_opt(2)=1
        DO setupt1 
   CASE dim_opt(3)=1
        DO setupt6              
ENDCASE
**********************************************************************
PROCEDURE setupmed1
IF USED('curdolmed')
   SELECT curdolmed
   USE   
ENDIF
SELECT * FROM sprdolj INTO CURSOR curdolmed READWRITE
ALTER TABLE curdolmed ADD COLUMN kstr N(3)
ALTER TABLE curdolmed ADD COLUMN kstr2 N(3)
SELECT curdolmed
INDEX ON namework TAG T1
SCAN ALL
     SELECT fondmed1
     LOCATE FOR ','+LTRIM(STR(curdolmed.kod))+','$msost
     IF FOUND()
        SELECT curdolmed
        REPLACE kstr WITH fondmed1.nkstr
     ENDIF 
     SELECT curdolmed 
ENDSCAN 
GO TOP 
fSupl.Visible=.F.
fSetup=CREATEOBJECT('FORMSUPL')
WITH fSetup
     .Caption='Отчет 1-медкадры'
     .Icon='kone.ico'
     .Width=800
     .Height=600
     .procExit='DO exitsetupmed1'
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=.Parent.Width
          .Height=.Parent.Height
          .ScrollBars=2          
          .ColumnCount=4
          .RecordSourceType=1     
          .RecordSource='curdolmed'
          .Column1.ControlSource='curdolmed.kod'
          .Column2.ControlSource='curdolmed.namework'
          .Column3.ControlSource='curdolmed.kstr'
                  
          .Column1.Width=RetTxtWidth('wкодw')
          .Column3.Width=RetTxtWidth('wтбл 1,4,7,9')         
          .Column2.Width=.Width-.column1.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount       
           .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='код'
          .Column2.Header1.Caption='наименование должности'
          .Column3.Header1.Caption='тбл 1,4,7,9'         
          .Column1.Movable=.F. 
          .Column1.Alignment=1
          .Column2.Alignment=0           
          .Column3.Alignment=1         
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'fSetup','fGrid','shapeingrid'
     DO myColumnTxtBox WITH 'fSetup.fGrid.column3','txtbox3','curdolmed.kstr',.F.,.F.,.F.,'DO validKdstr' 
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fSetup.fGrid.RecordSource)#fSetup.fGrid.curRec,fSetup.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fSetup.fGrid.RecordSource)#fSetup.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR  
     DO addButtonOne WITH 'fSetup','butRead',(.Width-RetTxtWidth('созранитьw')*2-20)/2,.fGrid.Top+.fGrid.Height+20,'редакция','','DO readsetupmed1',39,RetTxtWidth('wнастройкаw'),'формирование отчета' 
     DO addButtonOne WITH 'fSetup','butRet',.butRead.Left+.butRead.Width+10,.butRead.Top,'возврат','','DO exitsetupmed1',39,.butRead.Width,'возврат' 
     
     DO addButtonOne WITH 'fSetup','butRetRead',(.Width-.butRet.Width)/2,.butRet.Top,'возврат','','DO exitReadMed',39,.butRet.Width,'возврат' 
     .butRetRead.Visible=.F.
     .Height=.fGrid.Height+.butRead.Height+40
     .Autocenter=.T.
ENDWITH
fSetup.Show
**********************************************************************
PROCEDURE readsetupmed1
WITH fSetup
     .butRead.Visible=.F.
     .butRet.Visible=.F.
     .butRetRead.Visible=.T.
     .fGrid.Column4.Enabled=.F.
     .fGrid.Column3.Enabled=.T.
ENDWITH
**********************************************************************
PROCEDURE exitReadMed
WITH fSetup
     .butRead.Visible=.T.
     .butRet.Visible=.T.
     .butRetRead.Visible=.F.
     .fGrid.Column3.Enabled=.F.
     .fGrid.Column4.Enabled=.T.
ENDWITH
SELECT fondmed1
REPLACE msost WITH '' ALL
SELECT curdolmed
oldrec=RECNO()
SCAN ALL
     IF kstr#0
        SELECT fondmed1
        LOCATE FOR curdolmed.kstr=nkstr
        REPLACE msost WITH ','+ALLTRIM(msost)+LTRIM(STR(curdolmed.kod))+','
     ENDIF
     SELECT curdolmed
ENDSCAN
GO oldrec
**********************************************************************
PROCEDURE validKdstr

**********************************************************************
PROCEDURE exitSetupmed1
fSetup.Release
fSupl.Visible=.T.
**********************************************************************
PROCEDURE setupt1
IF USED('curdolmed')
   SELECT curdolmed
   USE   
ENDIF
SELECT * FROM sprdolj INTO CURSOR curdolmed READWRITE
SELECT curdolmed
INDEX ON namework TAG T1
INDEX ON kod TAG T2
SET ORDER TO 1
GO TOP 
fSupl.Visible=.F.
fSetup=CREATEOBJECT('FORMSUPL')
WITH fSetup
     .Caption='Отчет Т-1'
     .Icon='kone.ico'
     .Width=800
     .Height=600
     .procExit='DO exitsetupmed1'
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=.Parent.Width
          .Height=.Parent.Height
          .ScrollBars=2          
          .ColumnCount=4
          .RecordSourceType=1     
          .RecordSource='curdolmed'
          .Column1.ControlSource='curdolmed.kod'
          .Column2.ControlSource='curdolmed.namework'
          .Column3.ControlSource='curdolmed.nksup'
                  
          .Column1.Width=RetTxtWidth('wкодw')
          .Column3.Width=RetTxtWidth('wперсонал')         
          .Column2.Width=.Width-.column1.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount       
           .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='код'
          .Column2.Header1.Caption='наименование должности'
          .Column3.Header1.Caption='персонал'         
          .Column1.Movable=.F. 
          .Column1.Alignment=1
          .Column2.Alignment=0           
          .Column3.Alignment=1         
          .Column3.Format='Z'
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'fSetup','fGrid','shapeingrid'
     DO myColumnTxtBox WITH 'fSetup.fGrid.column3','txtbox3','curdolmed.nksup',.F.,.F.,.F.,'DO validNksup' 
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fSetup.fGrid.RecordSource)#fSetup.fGrid.curRec,fSetup.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fSetup.fGrid.RecordSource)#fSetup.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR 
     
     DO adLabMy WITH 'fSetup',1,'1 - руководители  2 - специалисты  3 - др. специалисты  4 - рабочие',.fGrid.Top+.fGrid.Height+10,0,.Width,2,.F.,0
      
     DO addButtonOne WITH 'fSetup','butRead',(.Width-RetTxtWidth('созранитьw')*2-20)/2,.lab1.Top+.lab1.Height,'редакция','','DO readsetupt1',39,RetTxtWidth('wнастройкаw'),'редактирование' 
     DO addButtonOne WITH 'fSetup','butRet',.butRead.Left+.butRead.Width+10,.butRead.Top,'возврат','','DO exitsetupt1',39,.butRead.Width,'возврат' 
     
     DO addButtonOne WITH 'fSetup','butRetRead',(.Width-.butRet.Width)/2,.butRet.Top,'возврат','','DO exitReadt1',39,.butRet.Width,'возврат' 
     .butRetRead.Visible=.F.
     .Height=.fGrid.Height+.butRead.Height+40
     .Autocenter=.T.
ENDWITH
fSetup.Show
**********************************************************************
PROCEDURE validnksup
*IF nksup>5
*   REPLACE nksup WITH 0
*ENDIF 
**********************************************************************
PROCEDURE readsetupt1
WITH fSetup
     .butRead.Visible=.F.
     .butRet.Visible=.F.
     .butRetRead.Visible=.T.
     .fGrid.Column4.Enabled=.F.
     .fGrid.Column3.Enabled=.T.
ENDWITH
**********************************************************************
PROCEDURE exitReadT1
WITH fSetup
     .butRead.Visible=.T.
     .butRet.Visible=.T.
     .butRetRead.Visible=.F.
     .fGrid.Column3.Enabled=.F.
     .fGrid.Column4.Enabled=.T.
ENDWITH
SELECT curdolmed
oldrec=RECNO()
SELECT sprdolj
REPLACE nksup WITH IIF(SEEK(kod,'curdolmed',2),curdolmed.nksup,nksup) ALL
GO oldrec
**********************************************************************
PROCEDURE exitSetupt1
fSetup.Release
fSupl.Visible=.T.
**********************************************************************
PROCEDURE setupt6
IF USED('curdolmed')
   SELECT curdolmed
   USE   
ENDIF
SELECT * FROM sprdolj INTO CURSOR curdolmed READWRITE
SELECT curdolmed
INDEX ON namework TAG T1
INDEX ON kod TAG T2
SET ORDER TO 1
GO TOP 
fSupl.Visible=.F.
fSetup=CREATEOBJECT('FORMSUPL')
WITH fSetup
     .Caption='Отчет Т-1'
     .Icon='kone.ico'
     .Width=800
     .Height=600
     .procExit='DO exitsetupmed1'
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=.Parent.Width
          .Height=.Parent.Height
          .ScrollBars=2          
          .ColumnCount=4
          .RecordSourceType=1     
          .RecordSource='curdolmed'
          .Column1.ControlSource='curdolmed.kod'
          .Column2.ControlSource='curdolmed.namework'
          .Column3.ControlSource='curdolmed.nksup6'
                  
          .Column1.Width=RetTxtWidth('wкодw')
          .Column3.Width=RetTxtWidth('wперсонал')         
          .Column2.Width=.Width-.column1.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount       
           .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='код'
          .Column2.Header1.Caption='наименование должности'
          .Column3.Header1.Caption='персонал'         
          .Column1.Movable=.F. 
          .Column1.Alignment=1
          .Column2.Alignment=0           
          .Column3.Alignment=1         
          .Column3.Format='Z'
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'fSetup','fGrid','shapeingrid'
     DO myColumnTxtBox WITH 'fSetup.fGrid.column3','txtbox3','curdolmed.nksup6',.F.,.F.,.F.,'DO validNksup6' 
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fSetup.fGrid.RecordSource)#fSetup.fGrid.curRec,fSetup.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fSetup.fGrid.RecordSource)#fSetup.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR 
     
     DO adLabMy WITH 'fSetup',1,'1 - руководители  2 - специалисты  3 - др. cлужащие  4 - рабочие',.fGrid.Top+.fGrid.Height+10,0,.Width,2,.F.,0
      
     DO addButtonOne WITH 'fSetup','butRead',(.Width-RetTxtWidth('созранитьw')*2-20)/2,.lab1.Top+.lab1.Height,'редакция','','DO readsetupt6',39,RetTxtWidth('wнастройкаw'),'редактирование' 
     DO addButtonOne WITH 'fSetup','butRet',.butRead.Left+.butRead.Width+10,.butRead.Top,'возврат','','DO exitsetupt1',39,.butRead.Width,'возврат' 
     
     DO addButtonOne WITH 'fSetup','butRetRead',(.Width-.butRet.Width)/2,.butRet.Top,'возврат','','DO exitReadt6',39,.butRet.Width,'возврат' 
     .butRetRead.Visible=.F.
     .Height=.fGrid.Height+.butRead.Height+40
     .Autocenter=.T.
ENDWITH
fSetup.Show
**********************************************************************
PROCEDURE validnksup6
*IF nksup>5
*   REPLACE nksup WITH 0
*ENDIF 
**********************************************************************
PROCEDURE readsetupt6
WITH fSetup
     .butRead.Visible=.F.
     .butRet.Visible=.F.
     .butRetRead.Visible=.T.
     .fGrid.Column4.Enabled=.F.
     .fGrid.Column3.Enabled=.T.
ENDWITH
**********************************************************************
PROCEDURE exitReadT6
WITH fSetup
     .butRead.Visible=.T.
     .butRet.Visible=.T.
     .butRetRead.Visible=.F.
     .fGrid.Column3.Enabled=.F.
     .fGrid.Column4.Enabled=.T.
ENDWITH
SELECT curdolmed
oldrec=RECNO()
SELECT sprdolj
REPLACE nksup6 WITH IIF(SEEK(kod,'curdolmed',2),curdolmed.nksup6,nksup6) ALL
GO oldrec
