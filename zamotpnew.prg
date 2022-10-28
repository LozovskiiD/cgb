IF USED('jobOtp')
   SELECT jobOtp
   USE
ENDIF
=AFIELDS(arJob,'datjob')
CREATE CURSOR jobOtp FROM ARRAY arJob
ALTER TABLE jobOtp ADD COLUMN avtVac L
ALTER TABLE jobOtp ADD COLUMN nVac N(1)
SELECT jobOtp
*INDEX ON STR(np,3)+STR(nd,3)+STR(tr,1) TAG T1
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 1

SELECT datShtat
LOCATE FOR ALLTRIM(pathTarif)=pathTarSupl
STORE 0 TO dim_tot,dim_day
formulaotp='IIF(!lokl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
formulaotp2='IIF(!logOkl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
RESTORE FROM rashotp ADDITIVE
SELECT datJob
SET FILTER TO 
REPLACE nrotp WITH RECNO() ALL 
SET ORDER TO 2
countDate=varDtar
SELECT rasp
SET FILTER TO 
ordOldRasp=SYS(21)
SET ORDER TO 2
SELECT sprpodr 
oldOrdPodr=SYS(21)
SET ORDER TO 2
GO TOP
kpSuplotp=kod
curnamepodr=name
DO selectOtpJob
var_path=FULLPATH('rashotp.mem')
fPodr=CREATEOBJECT('FORMSPR')
logOkl=.F.
WITH fPodr
     .Caption='Расчет планируемых отпусков на оплату труда, для лиц замещающих уходящих в отпуск работников'   
     DO addButtonOne WITH 'fPodr','menuCont1',10,5,'редакция','pencil.ico','Do readzam',39,RetTxtWidth('календарь')+44,'редакция'    
     DO addButtonOne WITH 'fPodr','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'удаление','pencild.ico','Do deletefromzam',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fPodr','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'расчёт','calculate.ico','DO procCountrash',39,.menucont1.Width,'расчёт'       
     DO addButtonOne WITH 'fPodr','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico','DO printRashod',39,.menucont1.Width,'печать' 
     DO addButtonOne WITH 'fPodr','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'настройки','setup.ico','DO setupZam',39,.menucont1.Width,'настройки'  
     DO addButtonOne WITH 'fPodr','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'возврат','undo.ico','DO exitFromProcOtp',39,.menucont1.Width,'возврат'       
     DO addButtonOne WITH 'fPodr','menuexit1',10,5,'возврат','undo.ico','DO exitReadPers',39,RetTxtWidth('возврат')+44,'вовзрат' 
     .menuexit1.Visible=.F.        
     DO addmenureadspr WITH 'fpodr','DO writezam WITH .F.','DO writezam WITH .T.'  
          
     DO addComboMy WITH 'fPodr',1,(.Width-500)/2,.menucont1.Top+.menucont1.Height+5,dheight,500,.T.,'curnamepodr','sprpodr.name',6,.F.,'DO validPodrRash',.F.,.T.      
     .comboBox1.DisplayCount=25
     WITH .fGrid    
          .Top=fpodr.ComboBox1.Top+fpodr.ComboBox1.Height+5 
          .Height=.Parent.Height-.Top
          .Width=.Parent.Width
          .RecordSource='jobOtp'
          DO addColumnToGrid WITH 'fPodr.fGrid',11
          .RecordSourceType=1     
          .Column1.ControlSource='jobOtp.kodpeop'
          .Column2.ControlSource='jobOtp.fio'          
          .Column3.ControlSource="IIF(SEEK(jobOtp.kd,'sprdolj',1),sprdolj.namework,'')"
          .Column4.ControlSource='jobOtp.lOkl'         
          .Column5.ControlSource='jobOtp.kse'                  
          .Column6.ControlSource='jobOtp.dotp'
          .Column7.ControlSource='jobOtp.dzam'
          .Column8.ControlSource='jobOtp.srzp'
          .Column9.ControlSource='jobOtp.zpday'     
          .Column10.ControlSource='jobOtp.zptot'  
          .Column1.Width=RettxtWidth('99999')    
          .Column3.Width=RettxtWidth('99999')
          .Column4.Width=RettxtWidth('9999')
          .Column5.Width=RettxtWidth('99999')
          .Column6.Width=RetTxtWidth('999999.99')     
          .Column7.Width=RetTxtWidth('999999.99')     
          .Column8.Width=RettxtWidth('99999999.99')
          .Column9.Width=RetTxtWidth('999999.99')     
          .Column10.Width=RetTxtWidth('999999.99')     
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=(.Width-.Column1.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width)/2
          .Column3.Width=.Width-.Column1.Width-.Column2.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='ФИО'      
          .Column3.Header1.Caption='Должность'
          .Column4.Header1.Caption='То'
          .Column5.Header1.Caption='Ш.ед.'
          .Column6.Header1.Caption='Дн.отп.'
          .Column7.Header1.Caption='Дн.зам.'
          .Column8.Header1.Caption='Ср.зп.'     
          .Column9.Header1.Caption='За 1 день.'  
          .Column10.Header1.Caption='всего'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'     
          .Column9.Format='Z' 
          .Column10.Format='Z' 
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column3.Alignment=0         
          .Column5.Alignment=1         
          .Column6.Alignment=1         
          .Column7.Alignment=1  
          .Column8.Alignment=1   
          .Column9.Alignment=1   
          .Column10.Alignment=1  
          .Column4.AddObject('checkColumn4','checkContainer')
          .Column4.checkColumn4.AddObject('checkMy','checkBox')
          .Column4.CheckColumn4.checkMy.Visible=.T.
          .Column4.CheckColumn4.checkMy.Caption=''
          .Column4.CheckColumn4.checkMy.Left=10
          .Column4.CheckColumn4.checkMy.BackStyle=0
          .Column4.CheckColumn4.checkMy.ControlSource='jobOtp.lokl'                                                                                                  
          .column4.CurrentControl='checkColumn4'
          .Column4.Sparse=.F.    
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .colNesInf=2              
     ENDWITH
     .AddObject('checkOkl','checkContainer')
     WITH .checkOkl
          .Width=.Parent.fGrid.Column4.Width+2    
          .AddObject('checkMy','MycheckBox')
          WITH .checkMy          
               .Caption=''
               .Left=10
               .BackStyle=0
               .ControlSource='jobOtp.lokl' 
               .Height=dHeight  
               .Visible=.T.                                                                                               
               .procForValid='DO procCheckOkl'
          ENDWITH     
          .Visible=.F.
          .BorderWidth=1
     ENDWITH      
     DO gridSizeNew WITH 'fpodr','fGrid','shapeingrid'
     
ENDWITH
fPodr.Show
********************************************************************************************************************************************************
PROCEDURE exitFromProcOtp
IF USED('datprn')
   SELECT datprn 
   USE 
ENDIF
IF USED('jobOtp')
   SELECT jobOtp 
   USE 
ENDIF
SELECT people
SET FILTER TO 
SELECT sprpodr 
SET ORDER TO &oldOrdPodr
SELECT rasp
SET ORDER TO &ordOldRasp
SET FILTER TO 
SELECT datJob
SET FILTER TO  
SET ORDER TO 4
fPodr.Release
********************************************************************************************************************************************************
PROCEDURE validPodrRash
SELECT sprpodr
kpsuplotp=sprpodr.kod
curnamepodr=fpodr.ComboBox1.Value
DO selectOtpJob
fpodr.Refresh
********************************************************************************************************************************************************
PROCEDURE selectOtpJob
SELECT jobOtp
SET ORDER TO 1
DELETE ALL
APPEND FROM datJob FOR kp=kpsuplotp
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
SELECT rasp
SET ORDER TO 2
SET FILTER TO kp=kpsuplotp
GO TOP
DO WHILE !EOF()
   IF rasp.kse#0
      SELECT datJob
      SET ORDER TO 2
      SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
      kse_cx=rasp.kse
      DO WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()   
          IF date_in>varDtar          
          ELSE 
             kse_cx=kse_cx-datjob.kse
          ENDIF 
          SKIP    
      ENDDO
      IF kse_cx>0  
         SELECT jobOtp
         
         APPEND BLANK
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH kse_cx,;
                    np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,avtVac WITH .T.,nvac WITH 1
         REPLACE pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2           
                    
         tar_ok=0
         tar_ok=varBaseSt*jobOtp.namekf*IIF(pkf#0,pkf,1)                      
         REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2)          
           
         REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                 mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse   
                  
         REPLACE srzp WITH rasp.srzp,zpday WITH rasp.zpday,dzam WITH rasp.dzam,dotp WITH rasp.dotp,zptot WITH rasp.totzp,lokl WITH rasp.lokl,dsttot WITH rasp.dsttot,omotp WITH rasp.omotp
         FOR i=1 TO 12
             repd='rasp.m'+LTRIM(STR(i))
             repd1='d'+LTRIM(STR(i)) 
             
             repdst='rasp.dst'+LTRIM(STR(i))
             repdst1='dst'+LTRIM(STR(i))              
             
             repz='rasp.z'+LTRIM(STR(i))
             repz1='zp'+LTRIM(STR(i))              
             REPLACE &repd1 WITH &repd,&repz1 WITH &repz,&repdst1 WITH &repdst             
         ENDFOR                          
        
      ENDIF       
   ENDIF
   SELECT rasp
   SKIP
ENDDO
SELECT jobOtp
GO TOP
********************************************************************************************************************************************************
PROCEDURE readzam
srZp_cx=0
kse_cx=0
IF rashotp(4)
   SELECT jobOtp
   oldRec=RECNO()
   logOkl=lOkl
   kpJob=kp
   kdJob=kd
   SET ORDER TO 2
   SEEK STR(kpJob,3)+STR(kdJob,3)
   
   SCAN WHILE kp=kpJob.AND.kd=kdJob
        srZp_cx=srZp_cx+&formulaotp2
        kse_cx=kse_cx+kse
   ENDSCAN
   SET ORDER TO 1
   GO oldRec   
   srZp_cx=IIF(kse_cx=0,0,IIF(kse_cx<1,srZp_cx,srZp_cx/kse_cx))
ENDIF

IF jobOtp.avtVac
   SELECT rasp
   SEEK STR(jobOtp.kp,3)+STR(jobOtp.kd,3)
ELSE 
   SELECT datJob
   SET ORDER TO 6
   SEEK STR(jobOtp.kodpeop,4)+STR(jobOtp.kp,3)+STR(jobOtp.kd,3)+STR(jobOtp.tr,1)
   *------- Страховка от идентичных записей (kp+kd+kse+tr)
   IF nrOtp#jobOtp.nrOtp
      SEEK STR(jobOtp.kodpeop,4)
      SCAN WHILE kodPeop=jobOtp.kodpeop
           IF nrotp=jobotp.nrotp
              EXIT 
           ENDIF 
      ENDSCAN      
   ENDIF
ENDIF 
SELECT jobOtp
IF !rashotp(4)
   srZp_cx=&formulaOtp
ENDIF
IF jobOtp.omotp
ELSE 
   *REPLACE srzp WITH &formulaOtp,datjob.srzp WITH jobOtp.srzp
   REPLACE srzp WITH srZp_cx,datjob.srzp WITH jobOtp.srzp   
ENDIF    
doljname=IIF(SEEK(jobOtp.kd,'sprdolj',1),ALLTRIM(sprdolj.namework),'')+' '+LTRIM(STR(datjob.kse,5,2))
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Редактирование - '+ALLTRIM(jobOtp.fio)
     .procExit='DO exitReadZam'
     .Width=800
     DO adTboxAsCont WITH 'fSupl','txtPodr',0,0,.Width,dHeight,doljName,2,1,.T.
     DO adTboxAsCont WITH 'fSupl','txtOkl',0,.txtPodr.Top+.txtPodr.Height-1,RetTxtWidth('wокладw'),dHeight,'оклад',2,1,.T.      
     
     .AddObject('checkOkl','checkContainer')
     WITH .checkOkl          
          .Left=0
          .Top=.Parent.txtOkl.Top+.Parent.txtOkl.Height-1
          .Height=dHeight
          .Width=.Parent.txtOkl.Width
          .BorderWidth=1
          .Visible=.T.  
          .AddObject('checkMy','MycheckBox')
          WITH .checkMy          
               .Caption=''              
               .BackStyle=0
               .ControlSource='jobOtp.lokl'                 
               .Visible=.T.   
               .Top=(.Parent.Height-.Height)/2
               .Left=(.Parent.Width-.Width)/2                                                                                           
               .procForValid='DO procCheckOkl'
          ENDWITH   
          .BorderWidth=1
     ENDWITH 
    
     DO adTboxAsCont WITH 'fSupl','txtDotp',.txtOkl.Left+.txtOkl.Width-1,.txtOkl.Top,RetTxtWidth('wwдни отпускаw'),dHeight,'дни отпуска',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxDotp',.txtDOtp.Top+.txtDotp.Height-1,.txtDOtp.Left,.txtDotp.Width,dHeight,'jobOtp.dotp','Z',.T.,2,.F.,'DO validzamdotp'
     
     DO adTboxAsCont WITH 'fSupl','txtDzam',.txtDotp.Left+.txtDotp.Width-1,.txtDotp.Top,.txtDotp.Width,dHeight,'дни замены',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxDzam',.boxDotp.Top,.txtDzam.Left,.txtDzam.Width,dHeight,'jobOtp.dzam','Z',.T.,2
     
               
          
     DO adTboxAsCont WITH 'fSupl','txtZpl',.txtDZam.Left+.txtDZam.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'зарплата',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxZpl',.boxDotp.Top,.txtZpl.Left,.txtZpl.Width,dHeight,'jobOtp.srzp','Z',.T.,2,.F.,'DO validZamDstOtp'
     .boxZpl.Enabled=IIF(jobOtp.omotp,.T.,.F.)
     
     DO adTboxAsCont WITH 'fSupl','txtZpDay',.txtZpl.Left+.txtZpl.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'за 1 день',2,1,.T.  
     DO adTboxNew WITH 'fSupl','boxZpDay',.boxDotp.Top,.txtZpDay.Left,.txtZpDay.Width,dHeight,'jobOtp.zpDay','Z',.T.,2
     
     DO adTboxAsCont WITH 'fSupl','txtZpTot',.txtZpDay.Left+.txtZpDay.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'всего',2,1,.T.  
     DO adTboxNew WITH 'fSupl','boxZpTot',.boxDotp.Top,.txtZpTot.Left,.txtZpTot.Width,dHeight,'jobOtp.zptot','Z',.F.,2
     
     
     DO adTboxAsCont WITH 'fSupl','txtmZpl',.txtZpTot.Left+.txtZpTot.Width-1,.txtdOtp.Top,RetTxtWidth('wо/рw'),dHeight,'о/р',2,1,.T.
     .AddObject('checkmOkl','checkContainer')
     WITH .checkMOkl          
          .Left=.Parent.txtmZpl.Left
          .Top=.Parent.boxDotp.Top
          .Height=dHeight
          .Width=.Parent.txtmZpl.Width
          .BorderWidth=1
          .Visible=.T.  
          .AddObject('checkMy','MycheckBox')
          WITH .checkMy          
               .Caption=''              
               .BackStyle=0
               .ControlSource='jobOtp.omotp'                 
               .Visible=.T.   
               .Top=(.Parent.Height-.Height)/2
               .Left=(.Parent.Width-.Width)/2                                                                                           
               .procForValid='DO procCheckMOkl'
          ENDWITH   
          .BorderWidth=1
     ENDWITH   
     
     
     
     DO adTboxAsCont WITH 'fSupl','txtMonth',0,.boxDotp.Top+.boxdOtp.Height-1,.txtOkl.Width+.txtDotp.Width+.txtDzam.Width+.txtMzpl.Width+.txtZpl.Width+.txtZpDay.Width+.txtZpTot.Width-6,dHeight,'по месяцам',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','txtMonthName',.txtMonth.Left,.txtMonth.Top+.txtMonth.Height-1,RetTxtWidth('wоктябрьw'),dHeight,'месяц',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','txtMonthDay',.txtMonthName.Left+.txtMonthName.Width-1,.txtMonthName.Top,RetTxtWidth('дни на став.w'),dHeight,'дни.отп.',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','txtMonthSt',.txtMonthDay.Left+.txtMonthDay.Width-1,.txtMonthName.Top,.txtMonthDay.Width,dHeight,'дни на ст.',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','txtMonthSum',.txtMonthSt.Left+.txtMonthSt.Width-1,.txtMonthName.Top,.txtMonthDay.Width,dHeight,'сумма',2,1,.T.  
     
     .txtMonthName.Width=.txtMonth.Width-.txtMonthDay.Width-.txtMonthSt.Width-.txtMonthSum.Width+3
     .txtMonthDay.Left=.txtMonthName.Left+.txtMonthName.Width-1
     .txtMonthSt.Left=.txtMonthDay.Left+.txtMonthDay.Width-1
     .txtMonthSum.Left=.txtMonthSt.Left+.txtMonthSt.Width-1     
     
     DO adTboxAsCont WITH 'fSupl','txtMonth1',0,.txtMonthName.Top+.txtMonthName.Height-1,.txtMonthName.Width,dHeight,'январь',0,1 
     DO adTboxNew WITH 'fSupl','box1',.txtMonth1.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d1','Z',.T.,2,.F.,'DO validMonthOtp WITH 1' 
     DO adTboxNew WITH 'fSupl','box11',.txtMonth1.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst1','Z',.T.,2,.F.,'DO validMonthOtp WITH 1'  
     DO adTboxNew WITH 'fSupl','box12',.txtMonth1.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp1','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth2',0,.txtMonth1.Top+.txtMonth1.Height-1,.txtMonthName.Width,dHeight,'февраль',0,1
     DO adTboxNew WITH 'fSupl','box2',.txtMonth2.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d2','Z',.T.,2,.F.,'DO validMonthOtp WITH 2'  
     DO adTboxNew WITH 'fSupl','box21',.txtMonth2.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst2','Z',.T.,2,.F.,'DO validMonthOtp WITH 2'  
     DO adTboxNew WITH 'fSupl','box22',.txtMonth2.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp2','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth3',0,.txtMonth2.Top+.txtMonth2.Height-1,.txtMonthName.Width,dHeight,'март',0,1
     DO adTboxNew WITH 'fSupl','box3',.txtMonth3.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d3','Z',.T.,2,.F.,'DO validMonthOtp WITH 3'  
     DO adTboxNew WITH 'fSupl','box31',.txtMonth3.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst3','Z',.T.,2,.F.,'DO validMonthOtp WITH 3' 
     DO adTboxNew WITH 'fSupl','box32',.txtMonth3.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp3','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth4',0,.txtMonth3.Top+.txtMonth3.Height-1,.txtMonthName.Width,dHeight,'апрель',0,1
     DO adTboxNew WITH 'fSupl','box4',.txtMonth4.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d4','Z',.T.,2,.F.,'DO validMonthOtp WITH 4'  
     DO adTboxNew WITH 'fSupl','box41',.txtMonth4.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst4','Z',.T.,2,.F.,'DO validMonthOtp WITH 4' 
     DO adTboxNew WITH 'fSupl','box42',.txtMonth4.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp4','Z',.F.,2          
          
     DO adTboxAsCont WITH 'fSupl','txtMonth5',0,.txtMonth4.Top+.txtMonth4.Height-1,.txtMonthName.Width,dHeight,'май',0,1
     DO adTboxNew WITH 'fSupl','box5',.txtMonth5.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d5','Z',.T.,2,.F.,'DO validMonthOtp WITH 5'  
     DO adTboxNew WITH 'fSupl','box51',.txtMonth5.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst5','Z',.T.,2,.F.,'DO validMonthOtp WITH 5' 
     DO adTboxNew WITH 'fSupl','box52',.txtMonth5.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp5','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth6',0,.txtMonth5.Top+.txtMonth5.Height-1,.txtMonthName.Width,dHeight,'июнь',0,1
     DO adTboxNew WITH 'fSupl','box6',.txtMonth6.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d6','Z',.T.,2,.F.,'DO validMonthOtp WITH 6'  
     DO adTboxNew WITH 'fSupl','box61',.txtMonth6.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst6','Z',.T.,2,.F.,'DO validMonthOtp WITH 6' 
     DO adTboxNew WITH 'fSupl','box62',.txtMonth6.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp6','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth7',0,.txtMonth6.Top+.txtMonth6.Height-1,.txtMonthName.Width,dHeight,'июль',0,1
     DO adTboxNew WITH 'fSupl','box7',.txtMonth7.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d7','Z',.T.,2,.F.,'DO validMonthOtp WITH 7'  
     DO adTboxNew WITH 'fSupl','box71',.txtMonth7.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst7','Z',.T.,2,.F.,'DO validMonthOtp WITH 7' 
     DO adTboxNew WITH 'fSupl','box72',.txtMonth7.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp7','Z',.F.,2 
     
     
     DO adTboxAsCont WITH 'fSupl','txtMonth8',0,.txtMonth7.Top+.txtMonth7.Height-1,.txtMonthName.Width,dHeight,'август',0,1
     DO adTboxNew WITH 'fSupl','box8',.txtMonth8.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d8','Z',.T.,2,.F.,'DO validMonthOtp WITH 8'  
     DO adTboxNew WITH 'fSupl','box81',.txtMonth8.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst8','Z',.T.,2,.F.,'DO validMonthOtp WITH 8' 
     DO adTboxNew WITH 'fSupl','box82',.txtMonth8.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp8','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth9',0,.txtMonth8.Top+.txtMonth8.Height-1,.txtMonthName.Width,dHeight,'сентябрь',0,1
     DO adTboxNew WITH 'fSupl','box9',.txtMonth9.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d9','Z',.T.,2,.F.,'DO validMonthOtp WITH 9'  
     DO adTboxNew WITH 'fSupl','box91',.txtMonth9.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst9','Z',.T.,2,.F.,'DO validMonthOtp WITH 9' 
     DO adTboxNew WITH 'fSupl','box92',.txtMonth9.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp9','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth10',0,.txtMonth9.Top+.txtMonth9.Height-1,.txtMonthName.Width,dHeight,'октябрь',0,1
     DO adTboxNew WITH 'fSupl','box10',.txtMonth10.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d10','Z',.T.,2,.F.,'DO validMonthOtp WITH 10'  
     DO adTboxNew WITH 'fSupl','box101',.txtMonth10.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst10','Z',.T.,2,.F.,'DO validMonthOtp WITH 10' 
     DO adTboxNew WITH 'fSupl','box102',.txtMonth10.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp10','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth11',0,.txtMonth10.Top+.txtMonth10.Height-1,.txtMonthName.Width,dHeight,'ноябрь',0,1
     DO adTboxNew WITH 'fSupl','box110',.txtMonth11.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d11','Z',.T.,2,.F.,'DO validMonthOtp WITH 11'  
     DO adTboxNew WITH 'fSupl','box1101',.txtMonth11.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst11','Z',.T.,2,.F.,'DO validMonthOtp WITH 11' 
     DO adTboxNew WITH 'fSupl','box1102',.txtMonth11.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp11','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth12',0,.txtMonth11.Top+.txtMonth11.Height-1,.txtMonthName.Width,dHeight,'декабрь',0,1
     DO adTboxNew WITH 'fSupl','box120',.txtMonth12.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'jobOtp.d12','Z',.T.,2,.F.,'DO validMonthOtp WITH 12' 
     DO adTboxNew WITH 'fSupl','box1201',.txtMonth12.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'jobOtp.dst12','Z',.T.,2,.F.,'DO validMonthOtp WITH 12' 
     DO adTboxNew WITH 'fSupl','box1202',.txtMonth12.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'jobOtp.zp12','Z',.F.,2 
     
     .txtPodr.Width=.txtMonth.Width 
     .Width=.txtPodr.Width
     .Height=dHeight*17
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE procCheckOkl
SELECT jobOtp
*IF !lOkl
   
*ELSE 
*   nms=jobOtp.mtokl
*ENDIF 


IF rashotp(4)
   SELECT jobOtp
   logOkl=lOkl
   oldRec=RECNO()
   kpJob=kp
   kdJob=kd
   nms=0
   kse_cx=0
   SET ORDER TO 2
   SEEK STR(kpJob,3)+STR(kdJob,3)
   
   SCAN WHILE kp=kpJob.AND.kd=kdJob
        nms=nms+&formulaotp2
        kse_cx=kse_cx+kse
   ENDSCAN
   SET ORDER TO 1
   GO oldRec  
   nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx))
ELSE
    nms=&formulaOtp    
ENDIF



REPLACE srzp WITH IIF(kse<=1,nms,nms),zpday WITH IIF(rashotp(1),srzp/rashotp(2),IIF(dotp#0,srzp/dotp,0))
IF !jobOtp.avtVac
   REPLACE datJob.srzp WITH jobOtp.srZp,datJob.zpday WITH jobOtp.zpDay,datJob.lOkl WITH jobOtp.lOkl,omotp WITH jobotp.omotp
ELSE
   REPLACE rasp.srzp WITH jobOtp.srZp,rasp.zpday WITH jobOtp.zpDay,rasp.lOkl WITH jobOtp.lOkl,rasp.ksezotp WITH jobOtp.kse,rasp.omotp WITH jobotp.omotp  
ENDIF   
KEYBOARD '{TAB}'
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE procCheckMOkl
fSupl.boxZpl.Enabled=IIF(jobOtp.omotp,.T.,.F.)
********************************************************************************************************************************************************
PROCEDURE validZamDotp
REPLACE dzam WITH ROUND(kse*dotp,0)
REPLACE zpday WITH IIF(rashotp(1),srzp/rashotp(2)*IIF(rashotp(4),kse,1),IIF(dotp#0,srzp/dotp*IIF(rashotp(4),kse,1),0)) 
IF !jobOtp.avtVac
   REPLACE datJob.dOtp WITH jobOtp.dOtp,datJob.dzam WITH jobOtp.dzam,datJob.zpday WITH jobOtp.zpday,omotp WITH jobotp.omotp
ELSE 
   REPLACE rasp.dOtp WITH jobOtp.dOtp,rasp.dzam WITH jobOtp.dzam,rasp.zpday WITH jobOtp.zpday,rasp.ksezotp WITH jobOtp.kse,rasp.omotp WITH jobotp.omotp
  ENDIF 
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE validZamDstOtp
REPLACE zpday WITH IIF(rashotp(1),srzp/rashotp(2)*IIF(rashotp(4),kse,1),IIF(dotp#0,srzp/dotp*IIF(rashotp(4),kse,1),0))  
IF !jobOtp.avtVac
   REPLACE datJob.zpday WITH jobOtp.zpday
ELSE 
   REPLACE rasp.zpday WITH jobOtp.zpday,rasp.ksezotp WITH jobOtp.kse  
ENDIF 
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE validMonthOtp
PARAMETERS par1
repd='jobOtp.d'+LTRIM(STR(par1))
repdst='jobOtp.dst'+LTRIM(STR(par1))
repzp='jobotp.zp'+LTRIM(STR(par1))

repdm='datJob.d'+LTRIM(STR(par1))
repdstm='datJob.dst'+LTRIM(STR(par1))
repzpm='datJob.zp'+LTRIM(STR(par1))

repdv='rasp.m'+LTRIM(STR(par1))
repdstv='rasp.dst'+LTRIM(STR(par1))
repzpv='rasp.z'+LTRIM(STR(par1))

SELECT jobOtp
*REPLACE &repdst WITH &repd*kse,&repzp WITH &repdst*zpday
REPLACE &repdst WITH &repd*kse,&repzp WITH &repd*zpday
REPLACE zptot WITH zp1+zp2+zp3+zp4+zp5+zp6+zp7+zp8+zp9+zp10+zp11+zp12
REPLACE dsttot WITH dst1+dst2+dst3+dst4+dst5+dst6+dst7+dst8+dst9+dst10+dst11+dst12    
*-----------------замена в datJob или rasp(для вакантных)
IF !jobOtp.avtVac
   REPLACE &repdm WITH &repd,&repdstm WITH &repdst,&repzpm WITH &repzp,datJob.zptot WITH jobOtp.zptot,datJob.dsttot WITH jobOtp.dsttot
ENDIF    
IF jobOtp.avtVac
   REPLACE &repdv WITH &repd,&repdstv WITH &repdst,&repzpv WITH &repzp,rasp.totzp WITH jobOtp.zptot,rasp.dsttot WITH jobOtp.dsttot
ENDIF    
SELECT jobOtp
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE exitReadZam
IF jobOtp.dzam=0
   REPLACE srzp WITH 0,zpday WITH 0 
   IF !jobOtp.avtVac
      REPLACE datJob.zpday WITH 0,datjob.srzp WITH 0
   ELSE 
      REPLACE rasp.zpday WITH 0,rasp.srzp WITH 0,rasp.ksezotp WITH 0
   ENDIF 
ENDIF
SELECT datJob
SET ORDER TO 2
SELECT jobOtp
***************************************************************************************************************************************************
*                   Процедура для настрек по работе с заменой отпусков
***************************************************************************************************************************************************
PROCEDURE setupzam
fsetup=CREATEOBJECT('FORMMY')
WITH fsetup
     .BackColor=RGB(255,255,255)
      DO addShape WITH 'fSetup',1,10,10,dHeight,0,8      
     .procexit='DO exitfsetup'  
     
     DO adCheckBox WITH 'fsetup','checkDay','для расчёта использовать среднее значение',.Shape1.Top+20,.Shape1.Left+20,150,dHeight,'rashotp(1)',0,.F.,'SAVE TO &var_path ALL LIKE rashotp' 
    .Shape1.Width=.checkDay.Width+40    

     DO adLabMy WITH 'fsetup',1,'Среднее кол-во дней',.checkday.Top+.checkDay.Height+10,.Shape1.Left+10,150,0,.T.
     DO adtbox WITH 'fsetup',1,.Lab1.Left+.lab1.Width+5,.lab1.Top,RetTxtWidth('9999999'),dHeight,'rashotp(2)','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashotp'
     
     .lab1.Top=.txtbox1.Top+(.txtbox1.Height-.lab1.Height)
     DO adLabMy WITH 'fsetup',2,'Дата отсчёта',.txtBox1.Top+.txtBox1.Height+10,.lab1.Left,150,0,.T.
     DO adtbox WITH 'fsetup',2,fsetup.lab2.Left+fSetup.lab2.Width+5,.txtbox1.Top+.txtBox1.Height+10,RetTxtWidth('99/99/999999'),dHeight,'countDate','Z',.T.,1
     
     DO adLabMy WITH 'fsetup',3,'Коэффициент уравнивания',.txtBox2.Top+.txtBox2.Height+10,.lab1.Left,150,0,.T.
     DO adtbox WITH 'fsetup',3,fsetup.lab3.Left+fSetup.lab3.Width+5,.txtbox2.Top+.txtBox2.Height+10,RetTxtWidth('99999999'),dHeight,'rashotp(3)','Z',.T.,1
     .txtBox3.InputMask='9.9999'
     
    DO adCheckBox WITH 'fsetup','checkOklad','использовать средний оклад',.txtBox3.Top+.txtBox3.Height+10,.Shape1.Left+20,150,dHeight,'rashotp(4)',0,.F.,'SAVE TO &var_path ALL LIKE rashotp' 
    .Shape1.Width=.checkDay.Width+40    
     
     
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.txtBox1.Width-10)/2
     .txtBox1.Left=.lab1.Left+.lab1.Width+10
     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+5)
     .lab2.Left=.Shape1.Left+(.Shape1.Width-.lab2.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.lab2.Left+.lab2.Width+10
     .lab2.Top=.txtBox2.Top+(.txtBox2.Height-.lab2.Height+5)
     
     .lab3.Left=.Shape1.Left+(.Shape1.Width-.lab3.Width-.txtBox3.Width-10)/2
     .txtBox3.Left=.lab3.Left+.lab3.Width+10
     .lab3.Top=.txtBox3.Top+(.txtBox3.Height-.lab3.Height+5)
     .checkOklad.Left=.Shape1.Left+(.Shape1.Width-.checkOklad.Width)/2     
     .Shape1.Height=.checkDay.Height*2+.txtBox1.Height*3+80 

     .Caption='настройки'   
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20    
ENDWITH
DO pasteImage WITH 'fsetup'
fsetup.Show
**************************************************************************************************************************
PROCEDURE exitfsetup
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
SAVE TO &var_path ALL LIKE rashotp
fsetup.Release
*****************************************************************************************************************************************************
*                         Форма для общего расчёта сведений по замене
*****************************************************************************************************************************************************
PROCEDURE proccountrash
fdel=CREATEOBJECT('FORMMY')
log_del=.F.
DIMENSION dim_del(4)
STORE 0 TO dim_del
dim_del(1)=1
dim_del(2)=0
dim_del(3)=0
log_srzp=.F.
WITH fdel
     .Caption='Расчёт расходов'
     .BackColor=RGB(255,255,255)   
     DO addShape WITH 'fdel',1,10,10,dHeight,50,8     
     DO adLabMy WITH 'fdel',1,'Дата отсчёта',.Shape1.Top+10,.Shape1.Left+15,150,0,.T.
     DO adtbox WITH 'fdel',1,.lab1.Left+.lab1.Width+10,.Shape1.Top+10,RetTxtWidth('99/99/99999'),dHeight,'countDate','Z',.T.,1
     .lab1.Top=.txtbox1.Top+(.txtbox1.Height-.lab1.Height)
     DO addOptionButton WITH 'fdel',1,'расчет по выбранной должности',.txtbox1.Top+.txtbox1.Height+10,.Shape1.Left+15,'dim_del(1)',0,"DO storedimdel WITH 1",.T.
     DO addOptionButton WITH 'fdel',2,'расчёт по подразделению',.Option1.Top+.Option1.Height+10,.Option1.Left,'dim_del(2)',0,"DO storedimdel WITH 2",.T.
     DO addOptionButton WITH 'fdel',3,'расчёт по организации',.Option2.Top+.Option2.Height+10,.Option1.Left,'dim_del(3)',0,"DO storedimdel WITH 3",.T.
     .Shape1.Height=.Option1.height*4+60    
     
     DO addShape WITH 'fdel',4,10,.Shape1.Top+.Shape1.Height+10,dHeight,.Shape1.Width,8 
     DO adCheckBox WITH 'fdel','check1','пересчитать среднюю зарплату',.Shape4.Top+10,.Option1.Left,150,dHeight,'log_srzp',0
     .Shape4.Height=.check1.Height+20
     .Shape4.Width=.check1.Width+30
     .Shape1.Width=.Shape4.Width   
     DO adCheckBox WITH 'fdel','check2','подтверждение выполнения',.Shape4.Top+.Shape4.Height+10,.Shape1.Left,150,dHeight,'log_del',0    
     .check2.Left=.shape4.Left+(.shape4.Width-.check2.Width)/2
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WВыполнениеW')*2-20)/2,.check2.Top+.check2.Height+15,RetTxtWidth('WВыполнениеW'),dHeight+3,'Выполнение','DO countrash'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fdel.Release'         
        
     DO adLabMy WITH 'fdel',4,'Ход выполнения',.check2.Top+.check2.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     .lab4.Visible=.F.        
     DO addShape WITH 'fdel',2,.Shape1.Left,.lab4.Top+.lab4.Height+5,dHeight,.Shape1.Width
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
          
     DO addShape WITH 'fdel',3,.Shape2.Left,.Shape2.Top,dHeight,0
     .Shape3.BackStyle=1
     .Shape3.Visible=.F. 
   
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape4.Height+.cont1.Height+.check1.Height+80
     .lab1.left=.Shape1.Left+(.shape1.Width-.lab1.Width-.txtbox1.Width-10)/2
     .txtbox1.Left=.lab1.Left+.lab1.Width+10
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
*****************************************************************************************************************************************************
*                                             Непосредственно  общий расчёт расходов по замене
*****************************************************************************************************************************************************
PROCEDURE countrash
IF !log_del
   fDel.Release
   RETURN
ENDIF
STORE 0 TO max_rec,one_pers,pers_ch
fdel.cont1.Visible=.F.
fdel.cont2.Visible=.F.
fdel.lab4.Visible=.T. 
fdel.Shape2.Visible=.T.
fdel.Shape3.Visible=.T.
=SYS(2002)
DO CASE
   CASE dim_del(1)=1
        fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus
        SELECT datJob
        SET ORDER TO 6
        SEEK STR(jobOtp.kodpeop,4)+STR(jobOtp.kp,3)+STR(jobOtp.kd,3)+STR(jobOtp.tr,1)
        max_rec=1 
        DO countone
        one_pers=one_pers+1
        pers_ch=one_pers/max_rec*100
        fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch 
   CASE dim_del(2)=1
        SELECT datJob
        SET FILTER TO kp=kpsuplotp
        STORE 0 TO max_rec,one_pers,pers_ch
        COUNT TO max_rec 
        GO TOP        
        DO WHILE !EOF()
           DO countone
           SELECT datJob
           SKIP 
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch               
        ENDDO   
        SELECT rasp
        SET FILTER TO ksezotp#0.AND.kp=kpsuplotp
        GO TOP
        DO WHILE !EOF()
           DO countonevac
           SELECT rasp
           SKIP 
        ENDDO 
        SET FILTER TO      
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO 
        
        SELECT datJob
        SET FILTER TO
        STORE 0 TO max_rec,one_pers,pers_ch
        COUNT TO max_rec 
        GO TOP        
        DO WHILE !EOF()
           DO countone
           SELECT datJob
           SKIP
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch  
        ENDDO
        SELECT rasp
        SET FILTER TO ksezotp#0
        GO TOP
        DO WHILE !EOF()
           DO countonevac
           SELECT rasp
           SKIP 
        ENDDO 
        SET FILTER TO        
ENDCASE
SELECT datJob
SET FILTER TO
SET ORDER TO 2
GO TOP
=INKEY(1)
fDel.Visible=.F.
fdel.Release
*=SYS(2002,1) 
DO selectOtpJob
fpodr.Refresh 
********************************************************************************************************************************************************
*                     Процедура расчёта расходов на замену по одной должности
********************************************************************************************************************************************************
PROCEDURE countone
SELECT datJob
oldJobOrd=SYS(21)
fjobrec=RECNO()
kpOld=kp
kdOld=kd
IF dzam#0   
   nms=0
   kse_cx=0
   **------зарплата и зарплата в день (если указано "пересчитывать спеднюю зарплату")
   IF log_srzp
      STORE 0 TO nms,srms  
      nms=0
      kse_cx=0
      IF rashotp(4)        
         SELECT datjob
         kpOld=kp
         kdOld=kd
         logOkl=lOkl
         SET ORDER TO 2
         SEEK STR(kpOld,3)+STR(kdOld,3)
         SCAN WHILE kp=kpOld.AND.kd=kdOld
              nms=nms+&formulaotp2
              kse_cx=kse_cx+kse
         ENDSCAN          
         SELECT rasp
         SEEK STR(kpOld,3)+STR(kdOld,3)                   
         IF rasp.kse>kse_cx           
            tar_ok=0
            ksevac=rasp.kse-kse_cx
            tar_ok=ROUND(varBaseSt*rasp.nkfvac*IIF(pkf#0,pkf,1)*ksevac,2)           
            IF !logOkl        
               nms=nms+tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksevac,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksevac,2)+;
                    ROUND(varBaseSt/100*pmain*ksevac,2)+ROUND(varBaseSt/100*pmain2*ksevac,2)+IIF(rashotp(3)#0,tar_ok*rashotp(3),0)                     
            ELSE
               nms=nms+tar_ok
            ENDIF         
            kse_cx=kse_cx+ksevac  
         ELSE
            REPLACE ksezotp WITH 0,totzp WITH 0  
         ENDIF        
         nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx) )
        
         SELECT datjob
         SET ORDER TO &oldJobOrd
         GO fjobrec
      ELSE    
         nms=&formulaOtp
      ENDIF
           
      REPLACE srzp WITH nms,zpday WITH IIF(rashotp(1),srzp/rashotp(2)*IIF(rashotp(4),kse,1),IIF(dotp#0,srzp/dotp*IIF(rashotp(4),kse,1),0))       
   ENDIF 
   *-------перерасчёт помесячно     
   FOR h=1 TO 12    
       repd='datJob.d'+LTRIM(STR(h))
       repdst='datJob.dst'+LTRIM(STR(h))
       repzp='datJob.zp'+LTRIM(STR(h))
       IF h>=MONTH(countDate)           
          repd='datJob.d'+LTRIM(STR(h))
          repdst='datJob.dst'+LTRIM(STR(h))
          repzp='datJob.zp'+LTRIM(STR(h))
          *REPLACE &repdst WITH &repd*kse,&repzp WITH &repdst*zpday
          REPLACE &repdst WITH &repd*kse,&repzp WITH &repd*zpday
       ELSE  
          REPLACE &repdst WITH 0,&repzp WITH 0                
       ENDIF       
   ENDFOR     
   REPLACE zptot WITH zp1+zp2+zp3+zp4+zp5+zp6+zp7+zp8+zp9+zp10+zp11+zp12    
   REPLACE dsttot WITH dst1+dst2+dst3+dst4+dst5+dst6+dst7+dst8+dst9+dst10+dst11+dst12     
ELSE
   REPLACE zptot WITH 0,srzp WITH 0,zpday WITH 0,dsttot WITH 0
   REPLACE dst1 WITH 0,dst2 WITH 0,dst3 WITH 0,dst4 WITH 0,dst5 WITH 0,dst6 WITH 0,dst7 WITH 0,dst8 WITH 0,dst9 WITH 0,dst10 WITH 0,dst11 WITH 0,dst12 WITH 0        
   REPLACE zp1 WITH 0,zp2 WITH 0,zp3 WITH 0,zp4 WITH 0,zp5 WITH 0,zp6 WITH 0,zp7 WITH 0,zp8 WITH 0,zp9 WITH 0,zp10 WITH 0,zp11 WITH 0,zp12 WITH 0               
ENDIF
********************************************************************************************************************************************************
*                     Процедура расчёта расходов на замену по одной должности
********************************************************************************************************************************************************
PROCEDURE countonevac
SELECT rasp
IF ksezotp#0   
   **------зарплата и зарплата в день (если указано "пересчитывать спеднюю зарплату")
   DO CASE 
      CASE rashotp(4)
           logOkl=lOkl
           SELECT datjob        
           SET ORDER TO 2
           SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
           nms=00
           kse_cx=0
           SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
                nms=nms+&formulaotp2
                kse_cx=kse_cx+kse
            ENDSCAN           
        
            SELECT rasp
            IF rasp.kse>kse_cx
               tar_ok=0
               ksevac=rasp.kse-kse_cx
               tar_ok=ROUND(varBaseSt*rasp.nkfvac*IIF(pkf#0,pkf,1)*ksevac,2) 
          
               IF !rasp.lokl        
                   nms=nms+tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksevac,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksevac,2)+;
                       ROUND(varBaseSt/100*pmain*ksezotp,2)+ROUND(varBaseSt/100*pmain2*ksevac,2)+IIF(rashotp(3)#0,tar_ok*rashotp(3),0)                                                                                                 
               ELSE
                  nms=nms+tar_ok
               ENDIF         
               kse_cx=kse_cx+ksevac  
            ENDIF          
            nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx) )   
            REPLACE srzp WITH nms,zpday WITH IIF(rashotp(1),srzp/rashotp(2)*IIF(rashotp(4),ksezotp,1),IIF(dotp#0,srzp/dotp,0)*IIF(rashotp(4),ksezotp,1))
      CASE !rashotp(4)
           IF log_srzp
              STORE 0 TO nms,srms 
              tar_ok=0
              tar_ok=ROUND(varBaseSt*rasp.nkfvac*IIF(pkf#0,pkf,1)*ksezotp,2)                      
              nms=tar_ok
              IF !rasp.lOkl
                 nms=tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksezotp,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksezotp,2)+;
                 ROUND(varBaseSt/100*pmain*ksezotp,2)+ROUND(varBaseSt/100*pmain2*ksezotp,2)+IIF(rashotp(3)#0,tar_ok*rashotp(3),0)
              ELSE 
                 nms=tar_ok
              ENDIF    
           ENDIF  
           REPLACE srzp WITH nms,zpday WITH IIF(rashotp(1),srzp/rashotp(2),IIF(dotp#0,srzp/dotp,0))
   ENDCASE 
   *-------перерасчёт помесячно     
   FOR h=1 TO 12    
       repd='rasp.m'+LTRIM(STR(h))
       repdst='rasp.dst'+LTRIM(STR(h))
       repzp='rasp.z'+LTRIM(STR(h))
       IF h>=MONTH(countDate)           
          repd='rasp.m'+LTRIM(STR(h))
          repdst='rasp.dst'+LTRIM(STR(h))
          repzp='rasp.z'+LTRIM(STR(h))
          *REPLACE &repdst WITH &repd*kse,&repzp WITH &repdst*zpday
          REPLACE &repdst WITH &repd*kse,&repzp WITH &repd*zpday
       ELSE  
          REPLACE &repdst WITH 0,&repzp WITH 0                
       ENDIF       
   ENDFOR     
   REPLACE totzp WITH z1+z2+z3+z4+z5+z6+z7+z8+z9+z10+z11+z12    
   REPLACE dsttot WITH dst1+dst2+dst3+dst4+dst5+dst6+dst7+dst8+dst9+dst10+dst11+dst12     
ELSE
   REPLACE totzp WITH 0,srzp WITH 0,zpday WITH 0,dsttot WITH 0
   REPLACE dst1 WITH 0,dst2 WITH 0,dst3 WITH 0,dst4 WITH 0,dst5 WITH 0,dst6 WITH 0,dst7 WITH 0,dst8 WITH 0,dst9 WITH 0,dst10 WITH 0,dst11 WITH 0,dst12 WITH 0        
   REPLACE z1 WITH 0,z2 WITH 0,z3 WITH 0,z4 WITH 0,z5 WITH 0,z6 WITH 0,z7 WITH 0,z8 WITH 0,z9 WITH 0,z10 WITH 0,z11 WITH 0,z12 WITH 0               
ENDIF
***********************************************************************************************************************************************
PROCEDURE storedimdel
PARAMETERS par1
FOR i=1 TO 4
    dim_del(i)=IIF(i=par1,1,0)
ENDFOR
fdel.Refresh
***********************************************************************************************************************************************
PROCEDURE printrashod
DIMENSION dimOpt(2)
dimOpt(1)=1
dimOpt(2)=0
fSupl=CREATEOBJECT('FORMSUPL')
term_ch=.T.
WITH fSupl
     .Caption='Ведомости'
     .procexit='DO exitPrintRepZam'
     DO addShape WITH 'fSupl',1,10,10,dHeight,400,8 
     DO addOptionButton WITH 'fSupl',1,'по должностям',.Shape1.Top+20,.Shape1.Left+15,'dimOpt(1)',0,"DO procValOption WITH 'fSupl','dimOpt',1",.T.
     DO addOptionButton WITH 'fSupl',2,'по сотрудникам',.Option1.Top,.Option1.Left,'dimOpt(2)',0,"DO procValOption WITH 'fSupl','dimOpt',2 ",.T.
     .Option1.Left=.Shape1.Left+(.Shape1.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20
     .Shape1.Height=.Option1.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.F.,.T.
   
     *-----------------------------Кнопка печать---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape91.Left+(.Shape91.Width-(RetTxtWidth('wпросмотрw')*3)-30)/2,;
       .Shape91.Top+.Shape91.Height+20,RetTxtWidth('wпросмотрw'),dHeight+5,'Печать','DO printRepZam WITH 1'
    *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
    DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Просмотр','DO printRepZam WITH 2'
    .SetAll('ForeColor',RGB(0,0,128),'CheckBox')  
    *---------------------------------Кнопка отмена --------------------------------------------------------------------------
    DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.cont2.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','DO exitPrintRepZam','Возврат'
    
     DO addShape WITH 'fSupl',11,.Shape91.Left,.cont1.Top,dHeight,.shape91.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,dHeight,0,8
     .Shape12.BackStyle=1
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+2,.Shape11.Left,.Shape11.Width,2,.F.,0
     .lab25.Visible=.F.             
    
    .Width=.Shape91.Width+40
    .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+70
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*********************************************************************************************
PROCEDURE exitPrintRepZam
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
fSupl.Release
*********************************************************************************************
PROCEDURE printRepZam
PARAMETERS par1
parTerm=par1
IF USED('curprn')
      SELECT curprn
      USE   
   ENDIF
IF dimOpt(1)=1
   SELECT * FROM datjob INTO CURSOR curJobPodr READWRITE
   SELECT curJobPodr
   INDEX ON STR(kp,3)+STR(kd,3) TAG T1
   SET ORDER TO 1
   SELECT * FROM rasp INTO CURSOR curprn READWRITE
   FOR i=1 TO 12
       calnd='nd'+LTRIM(STR(i))
       calzp='nzp'+LTRIM(STR(i))
       ALTER TABLE curprn ADD COLUMN &calnd N(4)
       ALTER TABLE curprn ADD COLUMN &calzp N(12,2)
   ENDFOR 
   ALTER TABLE curprn ADD COLUMN ndotp N(3)
   ALTER TABLE curprn ADD COLUMN nzptot N(12,2)
   ALTER TABLE curprn ADD COLUMN ksezam N(7,2)
   ALTER TABLE curprn ADD COLUMN npp N(3)
   ALTER TABLE curprn ADD COLUMN nIt N(1)
   
   REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
   REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,0) ALL
   INDEX ON STR(np,3)+STR(nd,3) TAG T1
   INDEX ON STR(kp,3)+STR(kd,3) TAG T2
   SET ORDER TO 1
   SELECT curJobPodr
   SELECT curprn
   SCAN ALL
        lKse=.F.
        sKse=0
        SELECT curJobPodr
        SEEK STR(curprn.kp,3)+STR(curprn.kd,3)
        DO WHILE kp=curprn.kp.AND.kd=curprn.kd
           lKse=.F.
           FOR i=MONTH(countDate) TO 12
               nrepzp='zp'+LTRIM(STR(i))
               nrepzp2='curprn.nzp'+LTRIM(STR(i))
               lKse=IIF(&nrepzp#0,.T.,lKse)
               REPLACE &nrepzp2 WITH &nrepzp2+&nrepzp,curprn.nzptot WITH curprn.nzptot+&nrepzp
           ENDFOR
           IF lKse
              REPLACE curprn.kseZam WITH curprn.kseZam+curJobPodr.kse
           ENDIF
           SKIP
        ENDDO
        SELECT curprn
   ENDSCAN   
   
   SELECT rasp
   SET FILTER TO ksezotp#0.AND.totzp#0
   GO TOP
   DO WHILE !EOF()
      IF SEEK(STR(rasp.kp,3)+STR(rasp.kd,3),'curprn',2)
         SELECT curprn  
         REPLACE nzptot WITH nzptot+rasp.totzp,nzp1 WITH nzp1+rasp.z1,nzp2 WITH nzp2+rasp.z2,nzp3 WITH nzp3+rasp.z3,nzp4 WITH nzp4+rasp.z4,;
                nzp5 WITH nzp5+rasp.z5,nzp6 WITH nzp6+rasp.z6,nzp7 WITH nzp7+rasp.z7,nzp8 WITH nzp8+rasp.z8,nzp9 WITH nzp9+rasp.z9,;
                nzp10 WITH nzp10+rasp.z10,nzp11 WITH nzp11+rasp.z11,nzp12 WITH nzp12+rasp.z12,ksezam WITH ksezam+rasp.ksezotp
      ENDIF 
      SELECT rasp
      SKIP
   ENDDO
   SET FILTER TO 
      
   SELECT sprpodr
   SCAN ALL
        SELECT curprn
        SUM kseZam,nzptot,nzp1,nzp2,nzp3,nzp4,nzp5,nzp6,nzp7,nzp8,nzp9,nzp10,nzp11,nzp12 TO kseZam_cx,nzptot_cx,;
            nzp1_cx,nzp2_cx,nzp3_cx,nzp4_cx,nzp5_cx,nzp6_cx,nzp7_cx,nzp8_cx,nzp9_cx,nzp10_cx,nzp11_cx,nzp12_cx FOR kp=sprpodr.kod
        APPEND BLANK 
        REPLACE kp WITH sprpodr.kod,nd WITH 999,np WITH sprpodr.np,nzptot WITH nzptot_cx,kseZam WITH kseZam_cx,nIt WITH 1,;
                nzp1 WITH nzp1_cx,nzp2 WITH nzp2_cx,nzp3 WITH nzp3_cx,nzp4 WITH nzp4_cx,nzp5 WITH nzp5_cx,nzp6 WITH nzp6_cx,;
                nzp7 WITH nzp7_cx,nzp8 WITH nzp8_cx,nzp9 WITH nzp9_cx,nzp10 WITH nzp10_cx,nzp11 WITH nzp11_cx,nzp12 WITH nzp12_cx,named WITH 'по отделению'            
        SELECT sprpodr
   ENDSCAN
   SELECT curprn
   DELETE FOR nzptot=0
   SUM kseZam,nzptot,nzp1,nzp2,nzp3,nzp4,nzp5,nzp6,nzp7,nzp8,nzp9,nzp10,nzp11,nzp12 TO kseZam_cx,nzptot_cx,;
       nzp1_cx,nzp2_cx,nzp3_cx,nzp4_cx,nzp5_cx,nzp6_cx,nzp7_cx,nzp8_cx,nzp9_cx,nzp10_cx,nzp11_cx,nzp12_cx FOR kd#0
   APPEND BLANK 
   REPLACE kp WITH 999,nd WITH 999,np WITH 999,nzptot WITH nzptot_cx,kseZam WITH kseZam_cx,nIt WITH 3,;
           nzp1 WITH nzp1_cx,nzp2 WITH nzp2_cx,nzp3 WITH nzp3_cx,nzp4 WITH nzp4_cx,nzp5 WITH nzp5_cx,nzp6 WITH nzp6_cx,;
           nzp7 WITH nzp7_cx,nzp8 WITH nzp8_cx,nzp9 WITH nzp9_cx,nzp10 WITH nzp10_cx,nzp11 WITH nzp11_cx,nzp12 WITH nzp12_cx,named WITH 'Итого'            
   GO TOP
   kpcx=0
   npcx=0
   DO WHILE !EOF()
      IF kp#kpcx
         npcx=1
         kpcx=kp  
      ENDIF
      REPLACE npp WITH npcx
      SKIP
      npcx=npcx+1
   ENDDO  
   GO TOP
   IF parTerm=1
      DO procForPrintAndPreview WITH 'repzamotppodr','',.T.,'zamOtpPodrToExcel'
   ELSE 
      DO procForPrintAndPreview WITH 'repzamotppodr','',.F. 
   ENDIF 
ELSE   
   SELECT rasp
   SET FILTER TO 
   SELECT * FROM datJob INTO CURSOR curPrn READWRITE
   ALTER TABLE curPrn ADD COLUMN nIt N(1)
   ALTER TABLE curPrn ADD COLUMN nVac N(1)
   ALTER TABLE curPrn ADD COLUMN npp N(3)
   ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
   ALTER TABLE curPrn ALTER COLUMN dsttot N(8)
   ALTER TABLE curPrn ALTER COLUMN zptot N(11,2)

   SELECT curprn
   DELETE FOR zptot=0
   REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
   REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL 
   REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL 
   INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
   INDEX ON STR(kp,3)+STR(kd,3) TAG T2

   SELECT rasp
   SET FILTER TO ksezotp#0.AND.totzp#0
   GO TOP
   DO WHILE !EOF()
      SELECT curprn
      APPEND BLANK
      REPLACE kp WITH rasp.kp,kd WITH rasp.kd,np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,fio WITH 'Вакантная',dotp WITH rasp.dotp,zpday WITH rasp.zpday,srzp WITH rasp.srzp,;
              zptot WITH rasp.totzp,zp1 WITH rasp.z1,zp2 WITH rasp.z2,zp3 WITH rasp.z3,zp4 WITH rasp.z4,zp5 WITH rasp.z5,zp6 WITH rasp.z6,;
              zp7 WITH rasp.z7,zp8 WITH rasp.z8,zp9 WITH rasp.z9,zp10 WITH rasp.z10,zp11 WITH rasp.z11,zp12 WITH rasp.z12
      REPLACE kse WITH rasp.ksezotp,nVac WITH 1        
      SELECT rasp
      SKIP
   ENDDO
   SET FILTER TO 

   SELECT sprpodr
   SCAN ALL
        SELECT curprn
        SUM kse,dsttot,zptot,zp1,zp2,zp3,zp4,zp5,zp6,zp7,zp8,zp9,zp10,zp11,zp12 TO ksecx,dsttotcx,zptotcx,zp1cx,zp2cx,zp3cx,zp4cx,zp5cx,zp6cx,zp7cx,zp8cx,zp9cx,zp10cx,zp11cx,zp12cx FOR kp=sprpodr.kod
        IF dsttotcx#0
           APPEND BLANK
           REPLACE kp WITH sprpodr.kod,np WITH sprpodr.np,nd WITH 98,dsttot WITH dsttotcx,fio WITH 'по отделению',nIt WITH 1,zptot WITH zptotcx,kse WITH ksecx
           REPLACE zp1 WITH zp1cx,zp2 WITH zp2cx,zp3 WITH zp3cx,zp4 WITH zp4cx,zp5 WITH zp5cx,zp6 WITH zp6cx,zp7 WITH zp7cx,zp8 WITH zp8cx,zp9 WITH zp9cx,zp10 WITH zp10cx,zp11 WITH zp11cx,zp12 WITH zp12cx
        ENDIF
        SELECT sprpodr     
   ENDSCAN
   SELECT curprn  
   SUM kse,dsttot,zptot,zp1,zp2,zp3,zp4,zp5,zp6,zp7,zp8,zp9,zp10,zp11,zp12 TO ksecx,dsttotcx,zptotcx,zp1cx,zp2cx,zp3cx,zp4cx,zp5cx,zp6cx,zp7cx,zp8cx,zp9cx,zp10cx,zp11cx,zp12cx FOR nIt=0
   APPEND BLANK
   REPLACE np WITH 999,nd WITH 98,dsttot WITH dsttotcx,fio WITH 'по организации',nIt WITH 9,zptot WITH zptotcx,kse WITH ksecx
   REPLACE zp1 WITH zp1cx,zp2 WITH zp2cx,zp3 WITH zp3cx,zp4 WITH zp4cx,zp5 WITH zp5cx,zp6 WITH zp6cx,zp7 WITH zp7cx,zp8 WITH zp8cx,zp9 WITH zp9cx,zp10 WITH zp10cx,zp11 WITH zp11cx,zp12 WITH zp12cx
   SET ORDER TO 1
   GO TOP
   kpcx=0
   npcx=0
   DO WHILE !EOF()
      IF kp#kpcx
         npcx=1
         kpcx=kp  
      ENDIF
      REPLACE npp WITH npcx
      SKIP
      npcx=npcx+1
   ENDDO  
   GO TOP
   IF parTerm=1
      DO procForPrintAndPreview WITH 'repzamotpnew','',.T.,'zamOtpToExcel'
   ELSE 
      DO procForPrintAndPreview WITH 'repzamotpnew','',.F. 
   ENDIF    
ENDIF
*************************************************************************************************************
PROCEDURE zamOtpToExcel
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
#DEFINE xlInsideHorizontal 12        
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F. 
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH  
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=25
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
     .Columns(17).ColumnWidth=8
     .Columns(18).ColumnWidth=8
     .Columns(19).ColumnWidth=8
     .Columns(20).ColumnWidth=8
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,20)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1    
     .Range(.Cells(rowcx,1),.Cells(rowcx,20)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Расчёт планируемых расходов на оплату лиц, заменяющих уходящих в отпуск работников'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH        
     rowcx=rowcx+1  
                         
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='№'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='наименование подразделения, ФИО работника'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH 
     .cells(rowcx,3).Value='наименование должности'                                                         
     .cells(rowcx,4).Value='дни замены'                                                         
     .cells(rowcx,5).Value='ди отпуска'                                                         
     .cells(rowcx,6).Value='оклад'                                                         
     .cells(rowcx,7).Value='в день'                                                        
             
     .cells(rowcx,8).Value='1'
     .cells(rowcx,9).Value='2'
     .cells(rowcx,10).Value='3'
     .cells(rowcx,11).Value='4'
     .cells(rowcx,12).Value='5'
     .cells(rowcx,13).Value='6'
     .cells(rowcx,14).Value='7'
     .cells(rowcx,15).Value='8'
     .cells(rowcx,16).Value='9'
     .cells(rowcx,17).Value='10'
     .cells(rowcx,18).Value='11'
     .cells(rowcx,19).Value='12'  
     .cells(rowcx,20).Value='всего'                                                           
  
     .Range(.Cells(rowcx,1),.Cells(rowcx,20)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     numberRow=rowcx+1  
     rowtop=numberRow         
     SELECT curPrn
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     GO TOP
     kpold=0
     SCAN ALL
          IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,20)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
          .Cells(numberRow,1).Value=npp                                  
          .Cells(numberRow,2).Value=fio                                       
          .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,'')                                                                                             
          .Cells(numberRow,4).Value=IIF(kse#0,kse,'')                                       
          .Cells(numberRow,4).NumberFormat='0.00'                                                    
          
          .Cells(numberRow,5).Value=IIF(dotp#0,dotp,'')                                       
          .Cells(numberRow,5).NumberFormat='0.00'                                                   
          
          .Cells(numberRow,6).Value=IIF(srzp#0,srzp,'')                                       
          .Cells(numberRow,6).NumberFormat='0.00'                                                              
          .Cells(numberRow,7).Value=IIF(zpday#0,zpday,'')                                       
          .Cells(numberRow,7).NumberFormat='0.00'    
                                                                          
          .Cells(numberRow,8).Value=IIF(zp1#0,zp1,'')                                       
          .Cells(numberRow,8).NumberFormat='0.00'                                                  
          
          .Cells(numberRow,9).Value=IIF(zp2#0,zp2,'')                                       
          .Cells(numberRow,9).NumberFormat='0.00'                                        
          
          .Cells(numberRow,10).Value=IIF(zp3#0,zp3,'')                                       
          .Cells(numberRow,10).NumberFormat='0.00'
          
          .Cells(numberRow,11).Value=IIF(zp4#0,zp4,'')                                       
          .Cells(numberRow,11).NumberFormat='0.00' 
          
          .Cells(numberRow,12).Value=IIF(zp5#0,zp5,'')                                       
          .Cells(numberRow,12).NumberFormat='0.00'
          
          .Cells(numberRow,13).Value=IIF(zp6#0,zp6,'')                                       
          .Cells(numberRow,13).NumberFormat='0.00' 
          
          .Cells(numberRow,14).Value=IIF(zp7#0,zp7,'')
          .Cells(numberRow,14).NumberFormat='0.00'                                          
          
          .Cells(numberRow,15).Value=IIF(zp8#0,zp8,'')
          .Cells(numberRow,15).NumberFormat='0.00'    
          
          .Cells(numberRow,16).Value=IIF(zp9#0,zp9,'')
          .Cells(numberRow,16).NumberFormat='0.00'    
          
          .Cells(numberRow,17).Value=IIF(zp10#0,zp10,'')
          .Cells(numberRow,17).NumberFormat='0.00'    
          
          .Cells(numberRow,18).Value=IIF(zp11#0,zp11,'')
          .Cells(numberRow,18).NumberFormat='0.00' 
             
          .Cells(numberRow,19).Value=IIF(zp12#0,zp12,'')
          .Cells(numberRow,19).NumberFormat='0.00'    
          
          .Cells(numberRow,20).Value=IIF(zptot#0,zptot,'')
          .Cells(numberRow,20).NumberFormat='0.00' 

          numberRow=numberRow+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,20)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1      
      
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,20)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
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
=SYS(2002)
=INKEY(2)
WITH fSupl
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T. 
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH        
objExcel.Visible=.T.

*************************************************************************************************************
PROCEDURE zamOtpPodrToExcel
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
#DEFINE xlInsideHorizontal 12        
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F. 
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH  
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=25
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
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,16)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1    
     .Range(.Cells(rowcx,1),.Cells(rowcx,16)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Расчёт планируемых расходов на оплату лиц, заменяющих уходящих в отпуск работников'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH        
     rowcx=rowcx+1  
                         
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='№'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='наименование должности'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH 
     .cells(rowcx,3).Value='долж.подл. замене'                                                         
             
     .cells(rowcx,4).Value='1'
     .cells(rowcx,5).Value='2'
     .cells(rowcx,6).Value='3'
     .cells(rowcx,7).Value='4'
     .cells(rowcx,8).Value='5'
     .cells(rowcx,9).Value='6'
     .cells(rowcx,10).Value='7'
     .cells(rowcx,11).Value='8'
     .cells(rowcx,12).Value='9'
     .cells(rowcx,13).Value='10'
     .cells(rowcx,14).Value='11'
     .cells(rowcx,15).Value='12'  
     .cells(rowcx,16).Value='всего'                                                           
  
     .Range(.Cells(rowcx,1),.Cells(rowcx,16)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     numberRow=rowcx+1  
     rowtop=numberRow         
     SELECT curPrn
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     GO TOP
     kpold=0
     SCAN ALL
          IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,16)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')        
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
          .Cells(numberRow,1).Value=IIF(nIt=0,npp,'')        
          .Cells(numberRow,2).Value=named
          .Cells(numberRow,3).Value=IIF(ksezam#0,ksezam,'')                                       
          .Cells(numberRow,3).NumberFormat='0.00'                                                    
                                                                                   
          .Cells(numberRow,4).Value=IIF(nzp1#0,nzp1,'')                                       
          .Cells(numberRow,4).NumberFormat='0.00'                                                  
          
          .Cells(numberRow,5).Value=IIF(nzp2#0,nzp2,'')                                       
          .Cells(numberRow,5).NumberFormat='0.00'                                        
          
          .Cells(numberRow,6).Value=IIF(nzp3#0,nzp3,'')                                       
          .Cells(numberRow,6).NumberFormat='0.00'
          
          .Cells(numberRow,7).Value=IIF(nzp4#0,nzp4,'')                                       
          .Cells(numberRow,7).NumberFormat='0.00' 
          
          .Cells(numberRow,8).Value=IIF(nzp5#0,nzp5,'')                                       
          .Cells(numberRow,8).NumberFormat='0.00'
          
          .Cells(numberRow,9).Value=IIF(nzp6#0,nzp6,'')                                       
          .Cells(numberRow,9).NumberFormat='0.00' 
          
          .Cells(numberRow,10).Value=IIF(nzp7#0,nzp7,'')
          .Cells(numberRow,10).NumberFormat='0.00'                                          
          
          .Cells(numberRow,11).Value=IIF(nzp8#0,nzp8,'')
          .Cells(numberRow,11).NumberFormat='0.00'    
          
          .Cells(numberRow,12).Value=IIF(nzp9#0,nzp9,'')
          .Cells(numberRow,12).NumberFormat='0.00'    
          
          .Cells(numberRow,13).Value=IIF(nzp10#0,nzp10,'')
          .Cells(numberRow,13).NumberFormat='0.00'    
          
          .Cells(numberRow,14).Value=IIF(nzp11#0,nzp11,'')
          .Cells(numberRow,14).NumberFormat='0.00' 
             
          .Cells(numberRow,15).Value=IIF(nzp12#0,nzp12,'')
          .Cells(numberRow,15).NumberFormat='0.00'    
          
          .Cells(numberRow,16).Value=IIF(nzptot#0,nzptot,'')
          .Cells(numberRow,16).NumberFormat='0.00' 

          numberRow=numberRow+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,16)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1      
      
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,17)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
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
=SYS(2002)
=INKEY(2)
WITH fSupl
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T. 
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH        
objExcel.Visible=.T.

*************************************************************************************************************
PROCEDURE deletefromzam
fdel=CREATEOBJECT('FORMMY')
log_del=.F.
DIMENSION dim_del(4)
STORE 0 TO dim_del
dim_del(1)=1
WITH fdel
     .Caption='Удаление'
     .BackColor=RGB(255,255,255)
     DO addShape WITH 'fdel',1,10,10,dHeight,8

     DO addOptionButton WITH 'fdel',1,'очистить выбранную запись',.Shape1.Top+10,.Shape1.Left+15,'dim_del(1)',0,"DO storedimdel WITH 1",.T.
     DO addOptionButton WITH 'fdel',2,'удалить подразделение',.Option1.Top+.Option1.Height+10,.Option1.Left,'dim_del(2)',0,"DO storedimdel WITH 2",.T.
     DO addOptionButton WITH 'fdel',3,'удалить все',.Option2.Top+.Option2.Height+10,.Option1.Left,'dim_del(3)',0,"DO storedimdel WITH 3",.T.
     .Shape1.Height=.Option1.height*3+40  
     .Shape1.Width=.Option1.Width+30
     DO adCheckBox WITH 'fdel','check1','подтверждение удаления',.Shape1.Top+.Shape1.Height+10,.Shape1.Left,150,dHeight,'log_del',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+5,.check1.Top+.check1.Height+15,(.shape1.Width-20)/2,dHeight+3,'Выполнение','DO delreczam'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fdel.Release' 
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+.check1.Height+50    
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
*****************************************************************************************************************************************************
*                         Непосредственно удаление информации по замене
*****************************************************************************************************************************************************
PROCEDURE delreczam
IF !log_del 
   RETURN 
ENDIF
DO CASE
   CASE dim_del(1)=1
        SELECT datJob
        SET ORDER TO 6
        SEEK STR(jobOtp.kodpeop,4)+STR(jobOtp.kp,3)+STR(jobOtp.kd,3)+STR(jobOtp.tr,1)
        FOR i=1 TO 12
            repd='d'+LTRIM(STR(i))
            repdst='dst'+LTRIM(STR(i))
            repzp='zp'+LTRIM(STR(i))
            REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
        ENDFOR                      
        REPLACE zptot WITH 0,dtot WITH 0,dsttot WITH 0,zpday WITH 0,srzp WITH 0,dotp WITH 0,dzam WITH 0
       
   CASE dim_del(2)=1
        SELECT datJob
        SET FILTER TO kp=kpsuplotp
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12
               repd='d'+LTRIM(STR(i))
               repdst='dst'+LTRIM(STR(i))
               repzp='zp'+LTRIM(STR(i))
               REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
           ENDFOR                      
           REPLACE zptot WITH 0,dtot WITH 0,dsttot WITH 0,zpday WITH 0,srzp WITH 0,dotp WITH 0,dzam WITH 0
           SKIP
        ENDDO         
              
   CASE dim_del(3)=1
        SELECT datJob 
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12
               repd='d'+LTRIM(STR(i))
               repdst='dst'+LTRIM(STR(i))
               repzp='zp'+LTRIM(STR(i))
               REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
           ENDFOR                      
           REPLACE zptot WITH 0,dtot WITH 0,dsttot WITH 0,zpday WITH 0,srzp WITH 0,dotp WITH 0,dzam WITH 0
           SKIP
        ENDDO           
                  
ENDCASE
fdel.Release
SELECT jobOtp
DO selectOtpJob
fpodr.Refresh
