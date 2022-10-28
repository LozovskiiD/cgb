IF USED('jobOtp')
   SELECT jobOtp
   USE
ENDIF
=AFIELDS(arJob,'datjob')
CREATE CURSOR jobOtp FROM ARRAY arJob
ALTER TABLE jobOtp ADD COLUMN avtVac L
ALTER TABLE jobOtp ADD COLUMN nVac N(1)
ON ERROR DO erSup
   ALTER TABLE jobOtp ADD COLUMN pzdrav N(3)
   ALTER TABLE jobOtp ADD COLUMN mzdrav N(10,2)
ON ERROR 
SELECT jobOtp
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 1

SELECT datShtat
LOCATE FOR ALLTRIM(pathTarif)=pathTarSupl
STORE 0 TO dim_tot,dim_day

formulaotp='IIF(!lokl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+mzdrav+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
formulaotp2='IIF(!logOkl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+mzdrav+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'

RESTORE FROM rashotp ADDITIVE
RESTORE FROM kfotp ADDITIVE
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
     DO addButtonOne WITH 'fPodr','butExit',10,5,'возврат','undo.ico','DO exitRead',39,RetTxtWidth('возврат')+44,'вовзрат' 
     .butExit.Visible=.F.        
              
     DO addComboMy WITH 'fPodr',1,(.Width-500)/2,.menucont1.Top+.menucont1.Height+5,dheight,500,.T.,'curnamepodr','sprpodr.name',6,.F.,'DO validPodrRash',.F.,.T.      
     .comboBox1.DisplayCount=25
     WITH .fGrid    
          .Top=fpodr.ComboBox1.Top+fpodr.ComboBox1.Height+5 
          .Height=.Parent.Height-.Top
          .Width=.Parent.Width
          .RecordSource='jobOtp'
          DO addColumnToGrid WITH 'fPodr.fGrid',14
          .RecordSourceType=1     
          .Column1.ControlSource='jobOtp.kodpeop'
          .Column2.ControlSource='jobOtp.fio'          
          .Column3.ControlSource="IIF(SEEK(jobOtp.kd,'sprdolj',1),sprdolj.namework,'')"
          .Column4.ControlSource='jobOtp.lOkl'         
          .Column5.ControlSource='jobOtp.kse'                  
          .Column6.ControlSource='jobOtp.ndotp'
          .Column7.ControlSource='jobOtp.ndatt'
          .Column8.ControlSource='jobOtp.ndkont'
          .Column9.ControlSource='jobOtp.ndst'
          .Column10.ControlSource='jobOtp.nsrzp'
          .Column11.ControlSource='jobOtp.nzpmonth'
          .Column12.ControlSource='jobOtp.nzpday'     
          .Column13.ControlSource='jobOtp.nzptot'  
          
          .Column1.Width=RettxtWidth('9999')    
          .Column4.Width=RetTxtWidth('99999')
          .Column5.Width=RettxtWidth('шт.ед')
          .Column6.Width=RettxtWidth('осн.отп.w')
          .Column7.Width=.Column6.Width
          .Column8.Width=.Column6.Width
          .Column9.Width=.Column6.Width
          .Column10.Width=RetTxtWidth('999999.99')     
          .Column11.Width=RetTxtWidth('999999.99')
          .Column12.Width=RetTxtWidth('999999.99')
          .Column13.Width=RetTxtWidth('999999.99')
                    
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=(.Width-.Column1.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-.Column12.Width-.Column13.Width)/2
          .Column3.Width=.Width-.Column1.Width-.Column2.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-.Column12.Width-.Column13.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='ФИО'      
          .Column3.Header1.Caption='должность'
          .Column4.Header1.Caption='то'
          .Column5.Header1.Caption='ш.ед.'
          .Column6.Header1.Caption='осн.от.'
          .Column7.Header1.Caption='атт.от.'
          .Column8.Header1.Caption='кон.от.'     
          .Column9.Header1.Caption='дн.зам.'  
          .Column10.Header1.Caption='ср.зп.'
          .Column11.Header1.Caption='в месяц'
          .Column12.Header1.Caption='в день'
          .Column13.Header1.Caption='всего'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'     
          .Column9.Format='Z' 
          .Column10.Format='Z'
          .Column11.Format='Z' 
          .Column12.Format='Z' 
          .Column13.Format='Z' 
           
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column3.Alignment=0         
          .Column5.Alignment=1         
          .Column6.Alignment=1         
          .Column7.Alignment=1  
          .Column8.Alignment=1   
          .Column9.Alignment=1   
          .Column10.Alignment=1  
          .Column11.Alignment=1   
          .Column12.Alignment=1  
          .Column13.Alignment=1  
          
          .Column4.AddObject('checkColumn4','checkContainer')
          .Column4.checkColumn4.AddObject('checkMy','checkMy')
          .Column4.CheckColumn4.checkMy.Visible=.T.
          .Column4.CheckColumn4.checkMy.Caption=''
          .Column4.CheckColumn4.checkMy.Left=10
          .Column4.CheckColumn4.checkMy.BackStyle=0
          .Column4.CheckColumn4.checkMy.ControlSource='jobOtp.lokl'    
          .Column4.CheckColumn4.checkMy.procValid='DO validdayotp'                                                                                              
          .column4.CurrentControl='checkColumn4'
          .Column4.Sparse=.F.    
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .colNesInf=2              
     ENDWITH      
     DO MyColumntxtBox WITH 'fpodr.fGrid.Column6','tbox6','jobOtp.ndotp',.F.,.F.,'','DO validdayotp'
     DO MyColumntxtBox WITH 'fpodr.fGrid.Column7','tbox7','jobOtp.ndatt',.F.,.F.,'','DO validdayotp'
     DO MyColumntxtBox WITH 'fpodr.fGrid.Column8','tbox8','jobOtp.ndkont',.F.,.F.,'','DO validdayotp'          
     DO gridSizeNew WITH 'fpodr','fGrid','shapeingrid'
     
ENDWITH
fPodr.Show
********************************************************************************************************************************************************
PROCEDURE validdayotp
SELECT jobOtp
REPLACE ndst WITH (ndotp+ndatt+ndkont)*kse
srzp_cx=0
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
SELECT jobOtp
IF !rashotp(4)
   srZp_cx=&formulaOtp
ENDIF
REPLACE nsrzp WITH srZp_cx,nzpmonth WITH nsrzp*kse,nzpday WITH nsrzp/rashotp(2),nzptot WITH ndst*nzpday
IF (ndotp+ndatt+ndkont)=0
   REPLACE nsrzp WITH 0,nzpmonth WITH 0,nzpday WITH 0,nzptot WITH 0
ENDIF
IF jobOtp.avtVac
   SELECT rasp
   SEEK STR(jobOtp.kp,3)+STR(jobOtp.kd,3)     
   REPLACE ndotp WITH jobOtp.ndotp,ndatt WITH jobOtp.ndatt,ndkont WITH jobOtp.ndkont,ndst WITH jobOtp.ndst,lOkl WITH jobOtp.lokl,ksezotp WITH jobOtp.kse,omotp WITH jobotp.omotp 
   REPLACE nsrzp WITH jobOtp.nsrzp,nzpmonth WITH jobOtp.nzpmonth,nzpday WITH jobOtp.nzpday,nzptot WITH jobOtp.nzptot     
ELSE 
   SELECT datJob
   SET ORDER TO 7
   SEEK jobOtp.nid   
   REPLACE ndotp WITH jobOtp.ndotp,ndatt WITH jobOtp.ndatt,ndkont WITH jobOtp.ndkont,ndst WITH jobOtp.ndst,lOkl WITH jobOtp.lokl
   REPLACE nsrzp WITH jobOtp.nsrzp,nzpmonth WITH jobOtp.nzpmonth,nzpday WITH jobOtp.nzpday,nzptot WITH jobOtp.nzptot
ENDIF 
SELECT jobOtp

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
         ON ERROR DO erSup        
            REPLACE pzdrav WITH rasp.pzdrav, mzdrav WITH mtokl/100*pzdrav          
         ON ERROR 
                  
         REPLACE nsrzp WITH rasp.nsrzp,nzpday WITH rasp.nzpday,ndzam WITH rasp.ndzam,ndotp WITH rasp.ndotp,ndatt WITH rasp.ndatt,ndkont WITH rasp.ndkont,nzptot WITH rasp.nzptot,lokl WITH rasp.lokl,;
                 ndst WITH rasp.ndst,omotp WITH rasp.omotp,nzpmonth WITH rasp.nzpmonth                                
        
      ENDIF       
   ENDIF
   SELECT rasp
   SKIP
ENDDO
SELECT jobOtp
GO TOP
********************************************************************************************************************************************************
PROCEDURE readzam
WITH fPodr
     .SetAll('Visible',.F.,'MyCommandButton')
     .butExit.Visible=.T.
     .comboBox1.Enabled=.F.
     .fGrid.Column4.Enabled=.T.
     .fGrid.Column6.Enabled=.T.
     .fGrid.Column7.Enabled=.T.
     .fGrid.Column8.Enabled=.T.
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.F.
ENDWITH
********************************************************************************************************************************************************
PROCEDURE exitRead
WITH fPodr
     .SetAll('Visible',.T.,'MyCommandButton')
     .butExit.Visible=.F.
     .comboBox1.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.     
ENDWITH
*************************************************************************************************************************
PROCEDURE procCheckOkl
SELECT jobOtp
REPLACE ndst WITH (ndotp+ndatt+ndkont)*kse
IF rashotp(4)
   SELECT jobOtp
   logOkl=lOkl
   oldRec=RECNO()
   kpJob=kp
   kdJob=kd
   srZp_cx=0
   kse_cx=0
   SET ORDER TO 2
   SEEK STR(kpJob,3)+STR(kdJob,3)
   
   SCAN WHILE kp=kpJob.AND.kd=kdJob
        srZp_cx=srZp_cx+&formulaotp2
        kse_cx=kse_cx+kse
   ENDSCAN
   SET ORDER TO 1
   GO oldRec  
   srZp_cx=IIF(kse_cx=0,0,IIF(kse_cx<1,srZp_cx,srZp_cx/kse_cx))
ELSE
    srZp_cx=&formulaOtp    
ENDIF
REPLACE nsrzp WITH srZp_cx,nzpmonth WITH nsrzp*kse,nzpday WITH nsrzp/rashotp(2),nzptot WITH ndst*nzpday


*REPLACE srzp WITH IIF(kse<=1,nms,nms),zpday WITH IIF(rashotp(1),srzp/rashotp(2),IIF(dotp#0,srzp/dotp,0))
IF !jobOtp.avtVac
   REPLACE datJob.srzp WITH jobOtp.srZp,datJob.zpday WITH jobOtp.zpDay,datJob.lOkl WITH jobOtp.lOkl,omotp WITH jobotp.omotp
ELSE
   REPLACE rasp.srzp WITH jobOtp.srZp,rasp.zpday WITH jobOtp.zpDay,rasp.lOkl WITH jobOtp.lOkl,rasp.ksezotp WITH jobOtp.kse,rasp.omotp WITH jobotp.omotp  
ENDIF   
KEYBOARD '{TAB}'
fSupl.Refresh
SELECT jobOtp
REPLACE ndst WITH (ndotp+ndatt+ndkont)*kse
srzp_cx=0
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
SELECT jobOtp
IF !rashotp(4)
   srZp_cx=&formulaOtp
ENDIF
REPLACE nsrzp WITH srZp_cx,nzpmonth WITH nsrzp*kse,nzpday WITH nsrzp/rashotp(2),nzptot WITH ndst*nzpday
IF (ndotp+ndatt+ndkont)=0
   REPLACE nsrzp WITH 0,nzpmonth WITH 0,nzpday WITH 0,nzptot WITH 0
ENDIF
IF jobOtp.avtVac
   SELECT rasp
   SEEK STR(jobOtp.kp,3)+STR(jobOtp.kd,3)     
   REPLACE ndotp WITH jobOtp.ndotp,ndatt WITH jobOtp.ndatt,ndkont WITH jobOtp.ndkont,ndst WITH jobOtp.ndst,lOkl WITH jobOtp.lokl,ksezotp WITH jobOtp.kse,omotp WITH jobotp.omotp 
   REPLACE nsrzp WITH jobOtp.nsrzp,nzpmonth WITH jobOtp.nzpmonth,nzpday WITH jobOtp.nzpday,nzptot WITH jobOtp.nzptot  
ELSE 
   SELECT datJob
   SET ORDER TO 7
   SEEK jobOtp.nid   
   REPLACE ndotp WITH jobOtp.ndotp,ndatt WITH jobOtp.ndatt,ndkont WITH jobOtp.ndkont,ndst WITH jobOtp.ndst,lOkl WITH jobOtp.lokl
   REPLACE nsrzp WITH jobOtp.nsrzp,nzpmonth WITH jobOtp.nzpmonth,nzpday WITH jobOtp.nzpday,nzptot WITH jobOtp.nzptot
ENDIF 
SELECT jobOtp

*************************************************************************************************************************
PROCEDURE procCheckMOkl
fSupl.boxZpl.Enabled=IIF(jobOtp.omotp,.T.,.F.)
***************************************************************************************************************************************************
*                   Процедура для настрек по работе с заменой отпусков
***************************************************************************************************************************************************
PROCEDURE setupzam
fsetup=CREATEOBJECT('FORMMY')
varPath1=FULLPATH('kfotp.mem')
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
     
     DO addShape WITH 'fSetup',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,dHeight,.Shape1.Width,8      
     DO adtBoxAsCont WITH 'fsetup','cnt1',.Shape2.Left+20,.Shape2.Top+10,RetTxtWidth('WWсентябрьW'),dHeight,'январь',1,1     
     DO adtbox WITH 'fsetup',111,.cnt1.Left+.cnt1.Width-1,.cnt1.Top,.Shape2.Width-.cnt1.Width-41,dHeight,'kfotp(1)','Z',.T.,0
     .txtBox111.InputMask='999.99'
     .txtBox111.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt2',.cnt1.Left,.cnt1.Top+.cnt1.Height-1,.cnt1.Width,dHeight,'февраль',1,1     
     DO adtbox WITH 'fsetup',112,.txtbox111.Left,.cnt2.Top,.txtbox111.Width,dHeight,'kfotp(2)','Z',.T.,1
     .txtBox112.InputMask='999.99'
     .txtBox112.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt3',.cnt1.Left,.cnt2.Top+.cnt1.Height-1,.cnt1.Width,dHeight,'март',1,1     
     DO adtbox WITH 'fsetup',113,.txtbox111.Left,.cnt3.Top,.txtbox111.Width,dHeight,'kfotp(3)','Z',.T.,1
     .txtBox113.InputMask='999.99'
     .txtBox113.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt4',.cnt1.Left,.cnt3.Top+.cnt3.Height-1,.cnt1.Width,dHeight,'апрель',1,1     
     DO adtbox WITH 'fsetup',114,.txtbox111.Left,.cnt4.Top,.txtbox111.Width,dHeight,'kfotp(4)','Z',.T.,1
     .txtBox114.InputMask='999.99'
     .txtBox114.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt5',.cnt1.Left,.cnt4.Top+.cnt4.Height-1,.cnt1.Width,dHeight,'май',1,1     
     DO adtbox WITH 'fsetup',115,.txtbox111.Left,.cnt5.Top,.txtbox111.Width,dHeight,'kfotp(5)','Z',.T.,1
     .txtBox115.InputMask='999.99'
     .txtBox115.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt6',.cnt1.Left,.cnt5.Top+.cnt5.Height-1,.cnt1.Width,dHeight,'июнь',1,1     
     DO adtbox WITH 'fsetup',116,.txtbox111.Left,.cnt6.Top,.txtbox111.Width,dHeight,'kfotp(6)','Z',.T.,1
     .txtBox116.InputMask='999.99'
     .txtBox116.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt7',.cnt1.Left,.cnt6.Top+.cnt6.Height-1,.cnt1.Width,dHeight,'июль',1,1     
     DO adtbox WITH 'fsetup',117,.txtbox111.Left,.cnt7.Top,.txtbox111.Width,dHeight,'kfotp(7)','Z',.T.,1
     .txtBox117.InputMask='999.99'
     .txtBox117.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt8',.cnt1.Left,.cnt7.Top+.cnt7.Height-1,.cnt1.Width,dHeight,'август',1,1     
     DO adtbox WITH 'fsetup',118,.txtbox111.Left,.cnt8.Top,.txtbox111.Width,dHeight,'kfotp(8)','Z',.T.,1
     .txtBox118.InputMask='999.99'
     .txtBox118.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt9',.cnt1.Left,.cnt8.Top+.cnt8.Height-1,.cnt1.Width,dHeight,'сентябрь',1,1     
     DO adtbox WITH 'fsetup',119,.txtbox111.Left,.cnt9.Top,.txtbox111.Width,dHeight,'kfotp(9)','Z',.T.,1
     .txtBox119.InputMask='999.99'
     .txtBox119.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt10',.cnt1.Left,.cnt9.Top+.cnt9.Height-1,.cnt1.Width,dHeight,'октябрь',1,1     
     DO adtbox WITH 'fsetup',120,.txtbox111.Left,.cnt10.Top,.txtbox111.Width,dHeight,'kfotp(10)','Z',.T.,1
     .txtBox120.InputMask='999.99'
     .txtBox120.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt11',.cnt1.Left,.cnt10.Top+.cnt10.Height-1,.cnt1.Width,dHeight,'ноябрь',1,1     
     DO adtbox WITH 'fsetup',121,.txtbox111.Left,.cnt11.Top,.txtbox111.Width,dHeight,'kfotp(11)','Z',.T.,1
     .txtBox121.InputMask='999.99'
     .txtBox121.Alignment=0     
     DO adtBoxAsCont WITH 'fsetup','cnt12',.cnt1.Left,.cnt11.Top+.cnt11.Height-1,.cnt1.Width,dHeight,'декабрь',1,1     
     DO adtbox WITH 'fsetup',122,.txtbox111.Left,.cnt12.Top,.txtbox111.Width,dHeight,'kfotp(12)','Z',.T.,1
     .txtBox122.InputMask='999.99'
     .txtBox122.Alignment=0     
    
     .Shape2.Height=.cnt1.Height*12+20-11
     
     .Caption='настройки'   
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape2.Height+30    
ENDWITH
DO pasteImage WITH 'fsetup'
fsetup.Show

**************************************************************************************************************************
PROCEDURE exitfsetup
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
SAVE TO &var_path ALL LIKE rashotp
SAVE TO &varpath1 ALL LIKE kfotp
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
        SET ORDER TO 7
        SEEK jobOtp.nid
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
IF ndotp+ndatt+ndkont#0
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
            REPLACE ksezotp WITH 0,nzptot WITH 0  
         ENDIF        
         nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx))        
         SELECT datjob
         SET ORDER TO &oldJobOrd
         GO fjobrec
      ELSE    
         nms=&formulaOtp
      ENDIF         
      REPLACE nsrzp WITH nms
   ENDIF      
   REPLACE ndst WITH (ndotp+ndatt+ndkont)*kse      
   REPLACE nzpmonth WITH nsrzp*kse,nzpday WITH nsrzp/rashotp(2),nzptot WITH ndst*nzpday      
ELSE
   REPLACE nzptot WITH 0,nsrzp WITH 0,nzpday WITH 0,ndst WITH 0,nzpmonth WITH 0   
ENDIF
********************************************************************************************************************************************************
*                     Процедура расчёта расходов на замену по одной должности
********************************************************************************************************************************************************
PROCEDURE countonevac
SELECT rasp
IF ndotp+ndatt+ndkont#0
   **------зарплата и зарплата в день (если указано "пересчитывать спеднюю зарплату")
   DO CASE 
      CASE rashotp(4)
           logOkl=lOkl
           SELECT datjob        
           SET ORDER TO 2
           SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
           nms=0
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
               REPLACE ksezotp WITH ksevac 
               IF !rasp.lokl        
                   nms=nms+tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksevac,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksevac,2)+;
                       ROUND(varBaseSt/100*pmain*ksezotp,2)+ROUND(varBaseSt/100*pmain2*ksevac,2)+IIF(rashotp(3)#0,tar_ok*rashotp(3),0)                                                                                                 
               ELSE
                  nms=nms+tar_ok
               ENDIF         
               kse_cx=kse_cx+ksevac  
            ENDIF          
            nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx) )   
            REPLACE nsrzp WITH nms
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
           REPLACE nsrzp WITH nms
   ENDCASE  
   REPLACE ndst WITH (ndotp+ndatt+ndkont)*ksezotp      
   REPLACE nzpmonth WITH nsrzp*ksezotp,nzpday WITH nsrzp/rashotp(2),nzptot WITH ndst*nzpday      
ELSE
   REPLACE nzptot WITH 0,nsrzp WITH 0,nzpday WITH 0,ndst WITH 0,nzpmonth WITH 0    
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
DIMENSION dimOpt(3)
dimOpt(1)=1
dimOpt(2)=0
dimOpt(3)=0
fSupl=CREATEOBJECT('FORMSUPL')
term_ch=.T.
WITH fSupl
     .Caption='Ведомости'
     .procexit='DO exitPrintRepZam'
     DO addShape WITH 'fSupl',1,10,10,dHeight,400,8 
     DO addOptionButton WITH 'fSupl',1,'по сотрудникам',.Shape1.Top+20,.Shape1.Left+15,'dimOpt(1)',0,"DO procValOption WITH 'fSupl','dimOpt',1",.T.
     DO addOptionButton WITH 'fSupl',2,'по должностям',.Option1.Top,.Option1.Left,'dimOpt(2)',0,"DO procValOption WITH 'fSupl','dimOpt',2 ",.T.
     DO addOptionButton WITH 'fSupl',3,'сводная',.Option1.Top+.Option1.Height+10,.Option1.Left,'dimOpt(3)',0,"DO procValOption WITH 'fSupl','dimOpt',3 ",.T.
     .Option1.Left=.Shape1.Left+(.Shape1.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20
     .Option3.Left=.Shape1.Left+(.Shape1.Width-.Option3.Width)/2
     .Shape1.Height=.Option1.Height*2+50
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.F.,.T.
     DO adButtonPrnToForm WITH 'DO printRepZam WITH 1','DO printRepZam WITH 2','DO exitPrintRepZam',.F.,'fSupl'
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width
     .Width=.Shape91.Width+20
     .Height=.butPrn.Top+.butPrn.Height+10
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

IF USED('curJobSupl')
   SELECT curJobSupl
   USE    
ENDIF

DO CASE   
   CASE dimOpt(1)=1
        SELECT rasp
        SET FILTER TO 
        SELECT * FROM datJob INTO CURSOR curPrn READWRITE
        ALTER TABLE curPrn ADD COLUMN nIt N(1)
        ALTER TABLE curPrn ADD COLUMN nVac N(1)
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
        ALTER TABLE curPrn ALTER COLUMN ndst N(8)
        ALTER TABLE curPrn ALTER COLUMN nzpday N(11,2)
        ALTER TABLE curPrn ALTER COLUMN nzpmonth N(11,2)
        ALTER TABLE curPrn ALTER COLUMN nzptot N(11,2)
        ON ERROR DO erSup
           ALTER TABLE curPrn ADD COLUMN pzdrav N(3)
           ALTER TABLE curPrn ADD COLUMN mzdrav N(10,2)
        ON ERROR 

        SELECT curprn
        DELETE FOR nzptot=0
        REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL 
        REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL 
        INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
        INDEX ON STR(kp,3)+STR(kd,3) TAG T2

        SELECT rasp
        SET FILTER TO ksezotp#0.AND.nzptot#0
        GO TOP
        DO WHILE !EOF()
           SELECT curprn
           APPEND BLANK
           REPLACE kp WITH rasp.kp,kd WITH rasp.kd,np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,fio WITH 'Вакантная',ndotp WITH rasp.ndotp,;
                   ndatt WITH rasp.ndatt,ndst WITH rasp.ndst,ndkont WITH rasp.ndkont,nzpday WITH rasp.nzpday,nsrzp WITH rasp.nsrzp,nzpmonth WITH rasp.nzpmonth,nzptot WITH rasp.nzptot
           REPLACE kse WITH rasp.ksezotp,nVac WITH 1        
           SELECT rasp
           SKIP
        ENDDO
        SET FILTER TO 

        SELECT sprpodr
        SCAN ALL
             SELECT curprn
             SUM kse,ndst,nzpmonth,nzpday,nzptot TO kse_cx,ndst_cx,nzpmonth_cx,nzpday_cx,nzptot_cx FOR kp=sprpodr.kod
             IF ndst_cx#0
                APPEND BLANK
                REPLACE kp WITH sprpodr.kod,np WITH sprpodr.np,nd WITH 98,ndst WITH ndst_cx,fio WITH 'по отделению',nIt WITH 1,nzptot WITH nzptot_cx,kse WITH kse_cx,;
                        nzpmonth WITH nzpmonth_cx,nzpday WITH nzpday_cx

             ENDIF
             SELECT sprpodr     
        ENDSCAN
        SELECT curprn  
        SUM kse,ndst,nzpmonth,nzpday,nzptot TO kse_cx,ndst_cx,nzpmonth_cx,nzpday_cx,nzptot_cx FOR nIt=0
        APPEND BLANK
        REPLACE np WITH 999,nd WITH 98,fio WITH 'по организации',nIt WITH 9,nzptot WITH nzptot_cx,kse WITH kse_cx,;
                ndst WITH ndst_cx,nzpmonth WITH nzpmonth_cx,nzpday WITH nzpday_cx   
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
           DO procForPrintAndPreview WITH 'zamotpusk','',.T.,'zamOtpuskToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'zamotpusk','',.F. 
        ENDIF    
  
   CASE dimOpt(2)=1                    
        SELECT rasp
        SET FILTER TO 
        SELECT * FROM rasp INTO CURSOR curprn READWRITE
        ALTER TABLE curPrn ADD COLUMN nIt N(1)
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ALTER COLUMN ksezotp N(7,2)
        ALTER TABLE curPrn ALTER COLUMN ndst N(8)
        ALTER TABLE curPrn ALTER COLUMN nzpday N(11,2)
        ALTER TABLE curPrn ALTER COLUMN nzptot N(11,2)
        
        ON ERROR DO erSup
           ALTER TABLE curPrn ADD COLUMN pzdrav N(3)
           ALTER TABLE curPrn ADD COLUMN mzdrav N(10,2)
        ON ERROR 
        
        DELETE ALL    
        SELECT * FROM rasp INTO CURSOR curprn1 READWRITE  
        SELECT curprn1
        DELETE FOR nzptot=0
        REPLACE kse WITH ksezotp ALL
                
        SELECT * FROM datjob WHERE ndst>0.AND.nzptot#0 INTO CURSOR curJobPodr READWRITE
        SELECT curJobPodr
        APPEND FROM DBF('curprn1')
        INDEX ON STR(kp,3)+STR(kd,3)+STR(nzpday,8,2) TAG T1
        SET ORDER TO 1
        SELECT rasp
        SCAN ALL
             SELECT curjobpodr
             SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
             IF FOUND()
                SELECT curprn
                APPEND BLANK
                REPLACE kp WITH rasp.kp,kd WITH rasp.kd,nd WITH rasp.nd
                SELECT curjobpodr
                DO WHILE kp=curprn.kp.AND.kd=curprn.kd                
                   SELECT curprn
                   REPLACE ndst WITH ndst+curjobpodr.ndst,nzpday WITH curjobpodr.nzpday,ksezotp WITH ksezotp+curjobpodr.kse                
                   SELECT curjobpodr        
                   SKIP
                   IF kp=rasp.kp.AND.kd=rasp.kd.AND.nzpday#curprn.nzpday
                      SELECT curprn
                      APPEND BLANK
                      REPLACE kp WITH rasp.kp,kd WITH rasp.kd,nd WITH rasp.nd
                      SELECT curjobpodr
                   ENDIF
                ENDDO
             ENDIF
             SELECT rasp             
        ENDSCAN
        SELECT curprn
        DELETE FOR ndst=0
        REPLACE nzptot WITH nzpday*ndst ALL
         
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
        REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,0) ALL
        INDEX ON STR(np,3)+STR(nd,3) TAG T1
        INDEX ON STR(kp,3)+STR(kd,3) TAG T2
        SET ORDER TO 1
        
        SELECT sprpodr
        SCAN ALL
             SELECT curprn
             SUM ksezotp,nzptot,ndst TO ksezotp_cx,nzptot_cx,ndst_cx FOR kp=sprpodr.kod
             APPEND BLANK 
             REPLACE kp WITH sprpodr.kod,nd WITH 999,np WITH sprpodr.np,nzptot WITH nzptot_cx,ndst WITH ndst_cx,nIt WITH 1,;
                     ksezotp WITH ksezotp_cx,named WITH 'по отделению'            
             SELECT sprpodr
        ENDSCAN
        SELECT curprn
        SUM ksezotp,nzptot,ndst TO ksezotp_cx,nzptot_cx,ndst_cx FOR kd=0

        APPEND BLANK 
        REPLACE kp WITH sprpodr.kod,nd WITH 999,np WITH 999,nzptot WITH nzptot_cx,ndst WITH ndst_cx,nIt WITH 3,;
                ksezotp WITH ksezotp_cx,named WITH 'Итого'            
        DELETE FOR nzptot=0
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
           DO procForPrintAndPreview WITH 'zamotpuskdol','',.T.,'zamOtpDolToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'zamotpuskdol','',.F. 
        ENDIF 
   CASE dimOpt(3)=1    
        CREATE CURSOR curotp (kstr N(2),nstr C(15),kse N(7,2),mtokl N(10,2),mtokl1 N(10,2),mtokl2 N(10,2),mstsum N(10,2),mchir N(10,2),mkat N(10,2),mvto N(10,2),mcharw N(10,2),mmain N(10,2),mmain2 N(10,2),mzdrav N(10,2),msupl N(10,2),mtot N(10,2),mRound N(10,2),npers N(6,2),mtotsupl N(10,2))      
        SELECT datjob
        SET FILTER TO
        SELECT * FROM datjob WHERE nzptot>0 INTO CURSOR curPrn READWRITE
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ADD COLUMN nVac N(1)
        ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
        
        ON ERROR DO erSup
           ALTER TABLE curPrn ADD COLUMN pzdrav N(3)
           ALTER TABLE curPrn ADD COLUMN mzdrav N(10,2)
        ON ERROR 
        
        SELECT curPrn
        DELETE FOR tokl=0     
        DELETE FOR date_in>varDTar    
        SELECT rasp
        SET FILTER TO ksezotp#0.AND.nzptot#0
        GO TOP
        DO WHILE !EOF()
           SELECT curprn
           APPEND BLANK
           REPLACE kp WITH rasp.kp,kd WITH rasp.kd,np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,fio WITH 'Вакантная',ndotp WITH rasp.ndotp,lOkl WITH rasp.lOkl;
                   ndatt WITH rasp.ndatt,ndst WITH rasp.ndst,ndkont WITH rasp.ndkont,nzpday WITH rasp.nzpday,nsrzp WITH rasp.nsrzp,nzpmonth WITH rasp.nzpmonth,nzptot WITH rasp.nzptot                                 
           REPLACE kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH rasp.ksezotp,vac WITH .T.,tr WITH 1                 
           tar_ok=0
           tar_ok=varBaseSt*namekf*IIF(pkf#0,pkf,1)                      
           REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2),;
                   pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2   
                                
           REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                   mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse                   
           ON ERROR DO erSup        
              REPLACE pzdrav WITH rasp.pzdrav, mzdrav WITH mtokl/100*pzdrav          
           ON ERROR        
           SELECT rasp           
           SKIP
        ENDDO
        SET FILTER TO 
        
        SELECT * FROM curprn INTO CURSOR curJobSupl READWRITE
        SELECT curJobSupl
        INDEX ON STR(kp,3)+STR(kd,3) TAG T1
        SELECT curotp
        APPEND BLANK 
        REPLACE kstr WITH 13
        SELECT curprn     
        
        
        STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvZdrav,nAvSupl,nAvOkl1,nAvOkl2   
        SCAN ALL
             IF (ndOtp+ndKont+ndatt)#0
                SELECT curJobsupl
                SEEK STR(curprn.kp,3)+STR(curprn.kd,3)
                STORE 0 TO nkse,nAvOkl,nAvOkl1,nAvOkl2,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvZdrav
                SCAN WHILE kp=curprn.kp.AND.kd=curprn.kd
                     nkse=nkse+kse                     
                     IF !curprn.lOkl
                        nAvOkl=nAvOkl+mtokl   
                        nAvSt=nAvSt+mstsum
                        nAvChir=nAvChir+mchir
                        nAvKat=nAvKat+mkat
                        nAvVto=nAvVto+mvto
                        nAvCharw=nAvCharw+mcharw
                        nAvMain=nAvMain+mmain
                        nAvMain2=nAvMain2+mmain2
                        nAvZdrav=nAvZdrav+mzdrav
                        nAvSupl=nAvSupl+mtokl*rashotp(3)
                     ELSE  
                        nAvOkl1=nAvOkl1+mtokl   
                     ENDIF   
                ENDSCAN           
                nAvOkl=IIF(nkse=0,0,IIF(nkse<1,nAvOkl,nAvOkl/nkse))
                nAvOkl1=IIF(nkse=0,0,IIF(nkse<1,nAvOkl1,nAvOkl1/nkse))
                             
                nAvSt=IIF(nkse=0,0,IIF(nkse<1,nAvSt,nAvSt/nkse))
                nAvChir=IIF(nkse=0,0,IIF(nkse<1,nAvChir,nAvChir/nkse))
                nAvKat=IIF(nkse=0,0,IIF(nkse<1,nAvKat,nAvKat/nkse))
                nAvVto=IIF(nkse=0,0,IIF(nkse<1,nAvVto,nAvVto/nkse))
                nAvCharw=IIF(nkse=0,0,IIF(nkse<1,nAvCharw,nAvCharw/nkse))
                nAvMain=IIF(nkse=0,0,IIF(nkse<1,nAvMain,nAvMain/nkse))
                nAvMain2=IIF(nkse=0,0,IIF(nkse<1,nAvMain2,nAvMain2/nkse))
                nAvZdrav=IIF(nkse=0,0,IIF(nkse<1,nAvZdrav,nAvZdrav/nkse))
                nAvSupl=IIF(nkse=0,0,IIF(nkse<1,nAvSupl,nAvSupl/nkse))                    
                SELECT curotp
                REPLACE mtokl WITH mtokl+ROUND(nAvOkl/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mstSum WITH mstsum+ROUND(nAvSt/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mChir WITH mChir+ROUND(nAvChir/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mKat WITH mKat+ROUND(nAvKat/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mVto WITH mVto+ROUND(nAvVto/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mCharw WITH mCharw+ROUND(nAvCharw/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mMain WITH mMain+ROUND(nAvMain/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mMain2 WITH mmain2+ROUND(nAvmain2/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;
                        mSupl WITH mSupl+ROUND(nAvSupl/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt),;                        
                        mtokl1 WITH mtokl1+ROUND(nAvOkl1/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt)     
                        ON ERROR DO erSup
                           REPLACE mZdrav WITH mZdrav+ROUND(nAvZdrav/rashotp(2),2)*curprn.kse*(curprn.ndotp+curprn.ndkont+curprn.ndatt)        
                        ON ERROR 
                REPLACE kse WITH kse+curprn.kse,mtot WITH mtot+curprn.nzptot
                SELECT curprn    
            ENDIF
       ENDSCAN
    
       SELECT curotp  
       REPLACE mtokl2 WITH mtokl+mstsum+mkat+mvto+mchir+msupl+mmain+mmain2+mcharw+mzdrav
*       REPLACE mtotsupl WITH mtokl1+mtokl2
       REPLACE mtotsupl WITH mtokl1+mtokl+mstsum+mkat+mvto+mchir+msupl+mmain+mmain2+mcharw+mzdrav,msupl WITH msupl+(mtot-mtotsupl),mtokl2 WITH mtokl+mstsum+mkat+mvto+mchir+msupl+mmain+mmain2+mcharw+mzdrav        
       mtotcx=mtot
       mtoklcx=mtokl
       mtokl1cx=mtokl1
       mtokl2cx=mtokl2
       mstsumcx=mstsum
       mkatcx=mkat
       mvtocx=mvto
       mcharwcx=mcharw
       mmaincx=mmain
       mmain2cx=mmain2
       mchircx=mchir
       msuplcx=msupl              
       mzdravcx=mzdrav
        
       FOR i=1 TO 12
           APPEND BLANK            
           REPLACE kstr WITH i,nstr WITH dim_month(i),mtot WITH mtotcx/100*kfotp(i),mtokl WITH mtoklcx/100*kfotp(i),mtokl1 WITH mtokl1cx/100*kfotp(i),mtokl2 WITH mtokl2cx/100*kfotp(i),;
                   mstsum WITH mstsumcx/100*kfotp(i),mkat WITH mkatcx/100*kfotp(i),mvto WITH mvtocx/100*kfotp(i),mcharw WITH mcharwcx/100*kfotp(i),mzdrav WITH mzdravcx/100*kfotp(i),;
                   mmain WITH mmaincx/100*kfotp(i),mmain2 WITH mmain2cx/100*kfotp(i),mchir WITH mchircx/100*kfotp(i),msupl WITH msuplcx/100*kfotp(i),npers WITH kfotp(i)
       ENDFOR
        
       SELECT curotp 
       INDEX ON kstr TAG T1 
       GO TOP
       IF parTerm=1
          DO procForPrintAndPreview WITH 'repotpsvod','',.T.,'otpSvodToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'repotpsvod','',.F. 
       ENDIF                 
ENDCASE    
***********************************************************
PROCEDURE otpSvodToExcel
DO startPrnToExcel WITH 'fSupl'
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
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=20
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
     .Columns(15).ColumnWidth=5
     rowcx=2
     rowtop=rowcx     
     .Range(.Cells(rowcx,1),.Cells(rowcx+2,1)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Месяц'          
     ENDWITH  
     .Range(.Cells(rowcx,2),.Cells(rowcx+2,2)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Всего замена отпуска'        
     ENDWITH  
     
     .Range(.Cells(rowcx,3),.Cells(rowcx+2,3)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Оклад без отработки'          
     ENDWITH  
     .Range(.Cells(rowcx,4),.Cells(rowcx+2,4)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Оклад с отработкой' 
     ENDWITH 
     
     .Range(.Cells(rowcx,5),.Cells(rowcx,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='в том числе' 
     ENDWITH           
           
     .Range(.Cells(rowcx+1,5),.Cells(rowcx+2,5)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='оклад'      
          .Font.Size=9
     ENDWITH     
       
     .Range(.Cells(rowcx+1,6),.Cells(rowcx+1,11)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Стимулирующие выплаты'         
     ENDWITH           
                 
      .Range(.Cells(rowcx+1,12),.Cells(rowcx+1,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Компенсирующие выплаты'         
     ENDWITH                        
                
     .Range(.Cells(rowcx,15),.Cells(rowcx+2,15)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='%'       
     ENDWITH      
     .cells(rowcx+2,6).Value='Указ №27 п.2'                                                         
     .cells(rowcx+2,7).Value='№52 п.4.1'                                                         
     .cells(rowcx+2,8).Value='№52 п.3'                                                         
     .cells(rowcx+2,9).Value='№52 п.4.5'
     .cells(rowcx+2,10).Value='№52 п.5'                                                                                                                 
     .cells(rowcx+2,11).Value='№53 п.5 Указ №27 п.3'                                                        
     .cells(rowcx+2,12).Value='№52 п.6.1'                                                         
     .cells(rowcx+2,13).Value='№52 п.6.6'                                                         
     .cells(rowcx+2,14).Value='№53 п.9'                                                              
     .Range(.Cells(rowcx,1),.Cells(rowcx+2,14)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     objExcel.Selection.VerticalAlignment=1
     rowcx=rowcx+3
     SELECT curOtp
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     GO TOP
     kpold=0
     SCAN ALL
          .Cells(rowcx,1).Value=nstr
          .Cells(rowcx,2).Value=IIF(mtot#0,mtot,'')
          .Cells(rowcx,2).NumberFormat='0.00'                                                  
          .Cells(rowcx,3).Value=IIF(mtokl1#0,mtokl1,'')                                                                                             
          .Cells(rowcx,3).NumberFormat='0.00'                                                  
          .Cells(rowcx,4).Value=IIF(mtokl2#0,mtokl2,'')                                       
          .Cells(rowcx,4).NumberFormat='0.00'                                                  
          .Cells(rowcx,5).Value=IIF(mtokl#0,mtokl,'')                                                 
          .Cells(rowcx,5).NumberFormat='0.00'                                                  
          .Cells(rowcx,6).Value=IIF(mstsum#0,mstsum,'')                                       
          .Cells(rowcx,6).NumberFormat='0.00'                                                  
          .Cells(rowcx,7).Value=IIF(mkat#0,mkat,'')                                                                        
          .Cells(rowcx,7).NumberFormat='0.00'                                                            
          .Cells(rowcx,8).Value=IIF(mvto#0,mvto,'')          
          .Cells(rowcx,8).NumberFormat='0.00'                                                            
          .Cells(rowcx,9).Value=IIF(mchir#0,mchir,'')
          .Cells(rowcx,9).NumberFormat='0.00'        
          .Cells(rowcx,10).Value=IIF(mzdrav#0,mzdrav,'')
          .Cells(rowcx,10).NumberFormat='0.00'        
          .Cells(rowcx,11).Value=IIF(msupl#0,msupl,'')                                       
          .Cells(rowcx,11).NumberFormat='0.00'          
          .Cells(rowcx,12).Value=IIF(mmain#0,mmain,'')                                       
          .Cells(rowcx,12).NumberFormat='0.00'          
          .Cells(rowcx,13).Value=IIF(mmain2#0,mmain2,'')                                       
          .Cells(rowcx,13).NumberFormat='0.00'          
          .Cells(rowcx,14).Value=IIF(mcharw#0,mcharw,'')                                       
          .Cells(rowcx,14).NumberFormat='0.00'          
          .Cells(rowcx,15).Value=IIF(npers#0,npers,'')                                       
          .Cells(rowcx,15).NumberFormat='0.00'          
          rowcx=rowcx+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(rowtop,1),.Cells(rowcx-1,15)).Select
      WITH objExcel.Selection
           .Borders(xlEdgeLeft).Weight=xlThin
           .Borders(xlEdgeTop).Weight=xlThin            
           .Borders(xlEdgeBottom).Weight=xlThin
           .Borders(xlEdgeRight).Weight=xlThin
           .Borders(xlInsideVertical).Weight=xlThin
           .Borders(xlInsideHorizontal).Weight=xlThin
           *objExcel.Selection.VerticalAlignment=1      
           .Font.Name='Times New Roman' 
           .Font.Size=9      
           .WrapText=.T.  
      ENDWITH 
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
DO endPrnToExcel WITH 'fSupl'    
objExcel.Visible=.T.        
*************************************************************************************************************
PROCEDURE zamOtpuskToExcel
DO startPrnToExcel WITH 'fSupl'
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
     rowcx=1     
     .Range(.Cells(rowcx,1),.Cells(rowcx,12)).Select  
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
     .Range(.Cells(rowcx,1),.Cells(rowcx,12)).Select  
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
     .cells(rowcx,4).Value='штатные единицы'                                                         
     .cells(rowcx,5).Value='основной отпуск'                                                         
     .cells(rowcx,6).Value='отпуск по результатам аттестации'                                                         
     .cells(rowcx,7).Value='отпуск контракт'                                                        
     .cells(rowcx,8).Value='кол-во дней замены с учетом объема работы'                                                         
     .cells(rowcx,9).Value='средний размер оплаты в месяц'                                                         
     .cells(rowcx,10).Value='средний размер оплаты в месяц с учетом объема работы'                                                         
     .cells(rowcx,11).Value='средний размер оплаты в день'                                                         
     .cells(rowcx,12).Value='сумма'               
                                                                
  
     .Range(.Cells(rowcx,1),.Cells(rowcx,12)).Select
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
              .Range(.Cells(numberRow,1),.Cells(numberRow,12)).Select
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
          .Cells(numberRow,2).Value=fio                                       
          .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,'')                                                                                             
          .Cells(numberRow,4).Value=IIF(kse#0,kse,'')                                       
          .Cells(numberRow,4).NumberFormat='0.00'                                                  
          .Cells(numberRow,5).Value=IIF(ndotp#0,ndotp,'')                                                 
          .Cells(numberRow,6).Value=IIF(ndatt#0,ndatt,'')                                       
          .Cells(numberRow,7).Value=IIF(ndkont#0,ndkont,'')                                                                        
          .Cells(numberRow,8).Value=IIF(ndst#0,ndst,'')          
          .Cells(numberRow,9).Value=IIF(nsrzp#0,nsrzp,'')
          .Cells(numberRow,9).NumberFormat='0.00'        
          .Cells(numberRow,10).Value=IIF(nzpmonth#0,nzpmonth,'')                                       
          .Cells(numberRow,10).NumberFormat='0.00'          
          .Cells(numberRow,11).Value=IIF(nzpday#0,nzpday,'')                                       
          .Cells(numberRow,11).NumberFormat='0.00'          
          .Cells(numberRow,12).Value=IIF(nzptot#0,nzptot,'')                                       
          .Cells(numberRow,12).NumberFormat='0.00'          
          numberRow=numberRow+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,12)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1      
      
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,12)).Select
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
DO endPrnToExcel WITH 'fSupl'     
objExcel.Visible=.T.

*************************************************************************************************************
PROCEDURE zamOtpDolToExcel
DO startPrnToExcel WITH 'fSupl'
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
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=8
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8    
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,6)).Select  
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
     .Range(.Cells(rowcx,1),.Cells(rowcx,6)).Select  
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
     .cells(rowcx,1).Value='№'                                                         
     .cells(rowcx,2).Value='наименование должности'                                                         
     .cells(rowcx,3).Value='штатные единицы'                                                         
     .cells(rowcx,4).Value='кол-во дней замены с учётом объёма работы'                                                         
     .cells(rowcx,5).Value='средний размер оплаты в день'                                                         
     .cells(rowcx,6).Value='сумма'                                                         
                                             
     .Range(.Cells(rowcx,1),.Cells(rowcx,6)).Select
     objExcel.Selection.WrapText=.T.
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
              .Range(.Cells(numberRow,1),.Cells(numberRow,6)).Select
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
          .Cells(numberRow,3).Value=IIF(ksezotp#0,ksezotp,'')                                       
          .Cells(numberRow,3).NumberFormat='0.00'                                                    
                                                                                   
          .Cells(numberRow,4).Value=IIF(ndst#0,ndst,'')    
          
          .Cells(numberRow,5).Value=IIF(nzpday#0,nzpday,'')                                       
          .Cells(numberRow,5).NumberFormat='0.00'                                        
          
          .Cells(numberRow,6).Value=IIF(nzptot#0,nzptot,'')                                       
          .Cells(numberRow,6).NumberFormat='0.00'

          numberRow=numberRow+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,6)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1            
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,6)).Select
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
DO endPrnToExcel WITH 'fSupl'
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
                          
        REPLACE nzptot WITH 0,dtot WITH 0,dsttot WITH 0,zpday WITH 0,srzp WITH 0,dotp WITH 0,dzam WITH 0
       
   CASE dim_del(2)=1
        SELECT datJob
        SET FILTER TO kp=kpsuplotp
        GO TOP
        DO WHILE !EOF()                               
           REPLACE nzptot WITH 0,ndtst WITH 0,nzpday WITH 0,nsrzp WITH 0,nzpmonth WITH 0
           SKIP
        ENDDO         
              
   CASE dim_del(3)=1
        SELECT datJob 
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()                             
           REPLACE zptot WITH 0,dtot WITH 0,dsttot WITH 0,zpday WITH 0,srzp WITH 0,dotp WITH 0,dzam WITH 0
           SKIP
        ENDDO           
                  
ENDCASE
fdel.Release
SELECT jobOtp
DO selectOtpJob
fpodr.Refresh
***************************************************************************************************************
