IF USED('jobOtp')
   SELECT jobOtp
   USE
ENDIF
SELECT datjob
SET FILTER TO 
SET ORDER TO 2
SELECT * FROM datjob INTO CURSOR jobotp READWRITE
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SET ORDER TO 1
**добавляем вакансии
SELECT rasp
SET FILTER TO 
ordOldRasp=SYS(21)
SET ORDER TO 1
SCAN ALL
     IF rasp.kse#0
      SELECT datJob     
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
                    np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,vac WITH .T.
         REPLACE pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2           
                    
         tar_ok=0
         tar_ok=varBaseSt*jobOtp.namekf*IIF(pkf#0,pkf,1)                      
         REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2)          
           
         REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                 mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse   
                  
         *REPLACE srzp WITH rasp.srzp,zpday WITH rasp.zpday,dzam WITH rasp.dzam,dotp WITH rasp.dotp,zptot WITH rasp.totzp,lokl WITH rasp.lokl,dsttot WITH rasp.dsttot,omotp WITH rasp.omotp
         *FOR i=1 TO 12
         *    repd='rasp.m'+LTRIM(STR(i))
         *    repd1='d'+LTRIM(STR(i)) 
         *    
         *    repdst='rasp.dst'+LTRIM(STR(i))
         *    repdst1='dst'+LTRIM(STR(i))              
         *    
         *    repz='rasp.z'+LTRIM(STR(i))
         *    repz1='zp'+LTRIM(STR(i))              
         *    REPLACE &repd1 WITH &repd,&repz1 WITH &repz,&repdst1 WITH &repdst             
         *ENDFOR                          
        
      ENDIF       
   ENDIF
   SELECT rasp  
ENDSCAN

*SELECT datShtat
*LOCATE FOR ALLTRIM(pathTarif)=pathTarSupl
*STORE 0 TO dim_tot,dim_day
formulaotp='IIF(!lokl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
formulaotp2='IIF(!logOkl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
RESTORE FROM rashotp ADDITIVE
SELECT datJob
*REPLACE nrotp WITH RECNO() ALL 
countDate=varDtar
SELECT rasp
GO TOP
SELECT sprpodr 
oldOrdPodr=SYS(21)
SET ORDER TO 2
GO TOP
kpSuplotp=kod
curnamepodr=name
DO selectOtpJob
*SELECT rasp
*SET FILTER TO kp=kpSuplOtp
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
          .RecordSource='rasp'
          DO addColumnToGrid WITH 'fPodr.fGrid',11
          .RecordSourceType=1     
          .Column1.ControlSource='nd'
          .Column2.ControlSource="IIF(SEEK(rasp.kd,'sprdolj',1),sprdolj.namework,'')"
          .Column3.ControlSource='lOkl'         
          .Column4.ControlSource='kse'            
          .Column5.ControlSource='ksezotp'            
          .Column6.ControlSource='dotp'
          .Column7.ControlSource='dzam'
          .Column8.ControlSource='srzp'
          .Column9.ControlSource='zpday'     
          .Column10.ControlSource='zptot'  
          .Column1.Width=RettxtWidth('99999')    
          .Column3.Width=RettxtWidth('99999')
          .Column4.Width=RettxtWidth('999999.99')
          .Column5.Width=RettxtWidth('999999.99')
          .Column6.Width=RetTxtWidth('999999.99')     
          .Column7.Width=RetTxtWidth('999999.99')     
          .Column8.Width=RettxtWidth('99999999.99')
          .Column9.Width=RetTxtWidth('999999.99')     
          .Column10.Width=RetTxtWidth('999999.99')     
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Должность'
          .Column3.Header1.Caption='То'
          .Column4.Header1.Caption='Ш.ед.'
          .Column5.Header1.Caption='Вак.'                    
          .Column6.Header1.Caption='Дн.отп.'
          .Column7.Header1.Caption='Дн.зам.'
          .Column8.Header1.Caption='Ср.зп.'     
          .Column9.Header1.Caption='За 1 день.'  
          .Column10.Header1.Caption='всего'
          .Column4.Format='Z'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'     
          .Column8.Format='Z' 
          .Column9.Format='Z' 
          .Column10.Format='Z' 
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column4.Alignment=1         
          .Column5.Alignment=1         
          .Column6.Alignment=1         
          .Column7.Alignment=1  
          .Column8.Alignment=1   
          .Column9.Alignment=1   
          .Column10.Alignment=1   
          .Column3.AddObject('checkColumn3','checkContainer')
          .Column3.checkColumn3.AddObject('checkMy','checkBox')
          .Column3.CheckColumn3.checkMy.Visible=.T.
          .Column3.CheckColumn3.checkMy.Caption=''
          .Column3.CheckColumn3.checkMy.Left=10
          .Column3.CheckColumn3.checkMy.BackStyle=0
          .Column3.CheckColumn3.checkMy.ControlSource='lokl'                                                                                                  
          .column3.CurrentControl='checkColumn3'
          .Column3.Sparse=.F.    
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .colNesInf=2              
     ENDWITH
     .AddObject('checkOkl','checkContainer')
     WITH .checkOkl
          .Width=.Parent.fGrid.Column3.Width+2    
          .AddObject('checkMy','MycheckBox')
          WITH .checkMy          
               .Caption=''
               .Left=10
               .BackStyle=0
               .ControlSource='lokl' 
               .Height=dHeight  
               .Visible=.T.                                                                                               
               *.procForValid='DO procCheckOkl'
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
PROCEDURE selectOtpJob
SELECT rasp
SET FILTER TO kp=kpsuplotp
GO TOP
DO WHILE !EOF()
   kse_cx=0
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
      SELECT rasp      
   ENDIF
   SELECT rasp
   REPLACE ksezotp WITH kse_cx      
   SKIP
ENDDO
GO TOP
********************************************************************************************************************************************************
PROCEDURE validPodrRash
SELECT sprpodr
kpsuplotp=sprpodr.kod
curnamepodr=fpodr.ComboBox1.Value
DO selectOtpJob
fpodr.Refresh
********************************************************************************************************************************************************
PROCEDURE readzam
srZp_cx=0
kse_cx=0
kpJob=rasp.kp
kdJob=rasp.kd
SELECT jobotp
SEEK STR(kpJob,3)+STR(kdJob,3)  
SCAN WHILE kp=kpJob.AND.kd=kdJob
     srZp_cx=srZp_cx+&formulaotp2
     kse_cx=kse_cx+kse
ENDSCAN  
SELECT rasp 
srZp_cx=IIF(kse_cx=0,0,IIF(kse_cx<1,srZp_cx,srZp_cx/kse_cx))
REPLACE srzp WITH srZp_cx
doljname=IIF(SEEK(rasp.kd,'sprdolj',1),ALLTRIM(sprdolj.namework),'')+' '+LTRIM(STR(datjob.kse,5,2))
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl  
     .Caption='Редактирование'
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
               .ControlSource='rasp.lokl'                 
               .Visible=.T.   
               .Top=(.Parent.Height-.Height)/2
               .Left=(.Parent.Width-.Width)/2                                                                                           
               .procForValid='DO procCheckOkl'
          ENDWITH   
          .BorderWidth=1
     ENDWITH 
    
     DO adTboxAsCont WITH 'fSupl','txtDotp',.txtOkl.Left+.txtOkl.Width-1,.txtOkl.Top,RetTxtWidth('wwдни отпускаw'),dHeight,'дни отпуска',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxDotp',.txtDOtp.Top+.txtDotp.Height-1,.txtDOtp.Left,.txtDotp.Width,dHeight,'rasp.dotp','Z',.T.,2,.F.,'DO validzamdotp'
     
     DO adTboxAsCont WITH 'fSupl','txtDzam',.txtDotp.Left+.txtDotp.Width-1,.txtDotp.Top,.txtDotp.Width,dHeight,'дни замены',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxDzam',.boxDotp.Top,.txtDzam.Left,.txtDzam.Width,dHeight,'rasp.dzam','Z',.T.,2           
          
     DO adTboxAsCont WITH 'fSupl','txtZpl',.txtDZam.Left+.txtDZam.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'зарплата',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxZpl',.boxDotp.Top,.txtZpl.Left,.txtZpl.Width,dHeight,'rasp.srzp','Z',.T.,2,.F.,'DO validZamDstOtp'
     .boxZpl.Enabled=IIF(rasp.omotp,.T.,.F.)
     
     DO adTboxAsCont WITH 'fSupl','txtZpDay',.txtZpl.Left+.txtZpl.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'за 1 день',2,1,.T.  
     DO adTboxNew WITH 'fSupl','boxZpDay',.boxDotp.Top,.txtZpDay.Left,.txtZpDay.Width,dHeight,'rasp.zpDay','Z',.T.,2
     
     DO adTboxAsCont WITH 'fSupl','txtZpTot',.txtZpDay.Left+.txtZpDay.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'всего',2,1,.T.  
     DO adTboxNew WITH 'fSupl','boxZpTot',.boxDotp.Top,.txtZpTot.Left,.txtZpTot.Width,dHeight,'rasp.totzp','Z',.F.,2     
     
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
               .ControlSource='rasp.omotp'                 
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
     DO adTboxNew WITH 'fSupl','box1',.txtMonth1.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m1','Z',.T.,2,.F.,'DO validMonthOtp WITH 1' 
     DO adTboxNew WITH 'fSupl','box11',.txtMonth1.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst1','Z',.T.,2,.F.,'DO validMonthOtp WITH 1'  
     DO adTboxNew WITH 'fSupl','box12',.txtMonth1.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp1','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth2',0,.txtMonth1.Top+.txtMonth1.Height-1,.txtMonthName.Width,dHeight,'февраль',0,1
     DO adTboxNew WITH 'fSupl','box2',.txtMonth2.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m2','Z',.T.,2,.F.,'DO validMonthOtp WITH 2'  
     DO adTboxNew WITH 'fSupl','box21',.txtMonth2.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst2','Z',.T.,2,.F.,'DO validMonthOtp WITH 2'  
     DO adTboxNew WITH 'fSupl','box22',.txtMonth2.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp2','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth3',0,.txtMonth2.Top+.txtMonth2.Height-1,.txtMonthName.Width,dHeight,'март',0,1
     DO adTboxNew WITH 'fSupl','box3',.txtMonth3.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m3','Z',.T.,2,.F.,'DO validMonthOtp WITH 3'  
     DO adTboxNew WITH 'fSupl','box31',.txtMonth3.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst3','Z',.T.,2,.F.,'DO validMonthOtp WITH 3' 
     DO adTboxNew WITH 'fSupl','box32',.txtMonth3.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp3','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth4',0,.txtMonth3.Top+.txtMonth3.Height-1,.txtMonthName.Width,dHeight,'апрель',0,1
     DO adTboxNew WITH 'fSupl','box4',.txtMonth4.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m4','Z',.T.,2,.F.,'DO validMonthOtp WITH 4'  
     DO adTboxNew WITH 'fSupl','box41',.txtMonth4.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst4','Z',.T.,2,.F.,'DO validMonthOtp WITH 4' 
     DO adTboxNew WITH 'fSupl','box42',.txtMonth4.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp4','Z',.F.,2          
          
     DO adTboxAsCont WITH 'fSupl','txtMonth5',0,.txtMonth4.Top+.txtMonth4.Height-1,.txtMonthName.Width,dHeight,'май',0,1
     DO adTboxNew WITH 'fSupl','box5',.txtMonth5.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m5','Z',.T.,2,.F.,'DO validMonthOtp WITH 5'  
     DO adTboxNew WITH 'fSupl','box51',.txtMonth5.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst5','Z',.T.,2,.F.,'DO validMonthOtp WITH 5' 
     DO adTboxNew WITH 'fSupl','box52',.txtMonth5.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp5','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth6',0,.txtMonth5.Top+.txtMonth5.Height-1,.txtMonthName.Width,dHeight,'июнь',0,1
     DO adTboxNew WITH 'fSupl','box6',.txtMonth6.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m6','Z',.T.,2,.F.,'DO validMonthOtp WITH 6'  
     DO adTboxNew WITH 'fSupl','box61',.txtMonth6.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst6','Z',.T.,2,.F.,'DO validMonthOtp WITH 6' 
     DO adTboxNew WITH 'fSupl','box62',.txtMonth6.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp6','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth7',0,.txtMonth6.Top+.txtMonth6.Height-1,.txtMonthName.Width,dHeight,'июль',0,1
     DO adTboxNew WITH 'fSupl','box7',.txtMonth7.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m7','Z',.T.,2,.F.,'DO validMonthOtp WITH 7'  
     DO adTboxNew WITH 'fSupl','box71',.txtMonth7.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst7','Z',.T.,2,.F.,'DO validMonthOtp WITH 7' 
     DO adTboxNew WITH 'fSupl','box72',.txtMonth7.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp7','Z',.F.,2 
     
     
     DO adTboxAsCont WITH 'fSupl','txtMonth8',0,.txtMonth7.Top+.txtMonth7.Height-1,.txtMonthName.Width,dHeight,'август',0,1
     DO adTboxNew WITH 'fSupl','box8',.txtMonth8.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m8','Z',.T.,2,.F.,'DO validMonthOtp WITH 8'  
     DO adTboxNew WITH 'fSupl','box81',.txtMonth8.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst8','Z',.T.,2,.F.,'DO validMonthOtp WITH 8' 
     DO adTboxNew WITH 'fSupl','box82',.txtMonth8.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp8','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth9',0,.txtMonth8.Top+.txtMonth8.Height-1,.txtMonthName.Width,dHeight,'сентябрь',0,1
     DO adTboxNew WITH 'fSupl','box9',.txtMonth9.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m9','Z',.T.,2,.F.,'DO validMonthOtp WITH 9'  
     DO adTboxNew WITH 'fSupl','box91',.txtMonth9.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst9','Z',.T.,2,.F.,'DO validMonthOtp WITH 9' 
     DO adTboxNew WITH 'fSupl','box92',.txtMonth9.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'zp9','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth10',0,.txtMonth9.Top+.txtMonth9.Height-1,.txtMonthName.Width,dHeight,'октябрь',0,1
     DO adTboxNew WITH 'fSupl','box10',.txtMonth10.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m10','Z',.T.,2,.F.,'DO validMonthOtp WITH 10'  
     DO adTboxNew WITH 'fSupl','box101',.txtMonth10.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst10','Z',.T.,2,.F.,'DO validMonthOtp WITH 10' 
     DO adTboxNew WITH 'fSupl','box102',.txtMonth10.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'z10','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth11',0,.txtMonth10.Top+.txtMonth10.Height-1,.txtMonthName.Width,dHeight,'ноябрь',0,1
     DO adTboxNew WITH 'fSupl','box110',.txtMonth11.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m11','Z',.T.,2,.F.,'DO validMonthOtp WITH 11'  
     DO adTboxNew WITH 'fSupl','box1101',.txtMonth11.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst11','Z',.T.,2,.F.,'DO validMonthOtp WITH 11' 
     DO adTboxNew WITH 'fSupl','box1102',.txtMonth11.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'z11','Z',.F.,2 
     
     DO adTboxAsCont WITH 'fSupl','txtMonth12',0,.txtMonth11.Top+.txtMonth11.Height-1,.txtMonthName.Width,dHeight,'декабрь',0,1
     DO adTboxNew WITH 'fSupl','box120',.txtMonth12.Top,.txtMonthDay.Left,.txtMonthDay.Width,dHeight,'m12','Z',.T.,2,.F.,'DO validMonthOtp WITH 12' 
     DO adTboxNew WITH 'fSupl','box1201',.txtMonth12.Top,.txtMonthSt.Left,.txtMonthSt.Width,dHeight,'dst12','Z',.T.,2,.F.,'DO validMonthOtp WITH 12' 
     DO adTboxNew WITH 'fSupl','box1202',.txtMonth12.Top,.txtMonthSum.Left,.txtMonthSum.Width,dHeight,'z12','Z',.F.,2 
     
     .txtPodr.Width=.txtMonth.Width 
     .Width=.txtPodr.Width
     .Height=dHeight*17
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE procCheckOkl

SELECT jobOtp
logOkl=rasp.lOkl
oldRec=RECNO()
kpJob=kp
kdJob=kd
nms=0
kse_cx=0
SEEK STR(kpJob,3)+STR(kdJob,3) 
SCAN WHILE kp=kpJob.AND.kd=kdJob
     nms=nms+&formulaotp2
     kse_cx=kse_cx+kse
ENDSCAN
SET ORDER TO 1
GO oldRec  
nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx))

REPLACE srzp WITH IIF(kse<=1,nms,nms),zpday WITH IIF(rashotp(1),srzp/rashotp(2),IIF(dotp#0,srzp/dotp,0))
KEYBOARD '{TAB}'
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE procCheckMOkl
fSupl.boxZpl.Enabled=IIF(rasp.omotp,.T.,.F.)
********************************************************************************************************************************************************
PROCEDURE validZamDotp
SELECT rasp
REPLACE dzam WITH ROUND(kse*dotp,0)
IF rashotp(1)
   REPLACE zpday WITH IIF(rashotp(2)#0,srzp/rashotp(2),0)
ELSE
   REPLACE zpday WITH IIF(dotp#0,srzp/dotp,0)
ENDIF
 
*REPLACE zpday WITH IIF(rashotp(1),srzp/rashotp(2)*IIF(rashotp(4),kse,1),IIF(dotp#0,srzp/dotp*IIF(rashotp(4),kse,1),0)) 
*IF !jobOtp.avtVac
*   REPLACE datJob.dOtp WITH jobOtp.dOtp,datJob.dzam WITH jobOtp.dzam,datJob.zpday WITH jobOtp.zpday,omotp WITH jobotp.omotp
*ELSE 
*   REPLACE rasp.dOtp WITH jobOtp.dOtp,rasp.dzam WITH jobOtp.dzam,rasp.zpday WITH jobOtp.zpday,rasp.ksezotp WITH jobOtp.kse,rasp.omotp WITH jobotp.omotp
*  ENDIF 
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE validZamDstOtp
SELECT rasp
IF rashotp(1)
   REPLACE zpday WITH IIF(rashotp(2)#0,srzp/rashotp(2),0)
ELSE
   REPLACE zpday WITH IIF(dotp#0,srzp/dotp,0)
ENDIF
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE validMonthOtp
PARAMETERS par1
*repdv='rasp.m'+LTRIM(STR(par1))
*repdstv='rasp.dst'+LTRIM(STR(par1))
*repzpv='rasp.z'+LTRIM(STR(par1))
*REPLACE &repdv WITH &repd,&repdstv WITH &repdst,&repzpv WITH &repzp
REPLACE totzp WITH z1+z2+z3+z4+z5+z6+z7+z8+z9+z10+z11+z12
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE exitReadZam
SELECT rasp
IF dzam=0
   REPLACE srzp WITH 0,zpday WITH 0  
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