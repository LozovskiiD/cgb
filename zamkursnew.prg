IF USED('jobKurs')
   SELECT jobKurs
   USE
ENDIF
=AFIELDS(arJob,'datjob')
CREATE CURSOR jobKurs FROM ARRAY arJob
ALTER TABLE jobKurs ADD COLUMN avtVac L
ALTER TABLE jobKurs ADD COLUMN nVac N(1)
SELECT jobKurs
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 1

SELECT datShtat
LOCATE FOR ALLTRIM(pathTarif)=pathTarSupl
STORE 0 TO dim_tot,dim_day
formulaotp='IIF(!lkokl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+mzdrav+IIF(rashkurs(3)#0,mtokl*rashkurs(3),0),mtokl)'
formulaotp2='IIF(!logOkl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+mzdrav+IIF(rashkurs(3)#0,mtokl*rashkurs(3),0),mtokl)'

*formulaotp='IIF(!lokl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+mzdrav+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
*formulaotp2='IIF(!logOkl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+mzdrav+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'

RESTORE FROM rashkurs ADDITIVE
SELECT datJob
SET FILTER TO 
REPLACE nrkurs WITH RECNO() ALL 
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
kpsuplkurs=kod
curnamepodr=name
DO selectKursJob
var_path=FULLPATH('rashkurs.mem')
fPodr=CREATEOBJECT('FORMSPR')

WITH fPodr
     .Caption='Расчет планируемых отпусков на оплату труда, для лиц замещающих уходящих на курсы работников'   
     DO addButtonOne WITH 'fPodr','menuCont1',10,5,'редакция','pencil.ico','Do readzamkurs',39,RetTxtWidth('календарь')+44,'редакция'    
     DO addButtonOne WITH 'fPodr','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'удаление','pencild.ico','Do deletefromkurs',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fPodr','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'расчёт','calculate.ico','DO formCountRash',39,.menucont1.Width,'расчёт'       
     DO addButtonOne WITH 'fPodr','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico','DO printKurs',39,.menucont1.Width,'печать' 
     DO addButtonOne WITH 'fPodr','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'настройки','setup.ico','DO setupKurs',39,.menucont1.Width,'настройки'  
     DO addButtonOne WITH 'fPodr','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'возврат','undo.ico','DO exitFromProcOtp',39,.menucont1.Width,'возврат'       
     DO addButtonOne WITH 'fPodr','menuexit1',10,5,'возврат','undo.ico','DO exitReadPers',39,RetTxtWidth('возврат')+44,'вовзрат' 
     .menuexit1.Visible=.F.        
     DO addmenureadspr WITH 'fpodr','DO writezam WITH .F.','DO writezam WITH .T.'  
          
     DO addComboMy WITH 'fPodr',1,(.Width-500)/2,.menucont1.Top+.menucont1.Height+5,dheight,500,.T.,'curnamepodr','sprpodr.name',6,.F.,'DO validPodrKurs',.F.,.T.      
     .comboBox1.DisplayCount=25
     WITH .fGrid    
          .Top=fpodr.ComboBox1.Top+fpodr.ComboBox1.Height+5 
          .Height=.Parent.Height-.Top
          .Width=.Parent.Width
          .RecordSource='jobKurs'
          DO addColumnToGrid WITH 'fPodr.fGrid',12
          .RecordSourceType=1     
          .Column1.ControlSource='jobKurs.kodpeop'
          .Column2.ControlSource='jobKurs.fio'          
          .Column3.ControlSource="IIF(SEEK(jobKurs.kd,'sprdolj',1),sprdolj.namework,'')"
          .Column4.ControlSource='jobKurs.lkOkl'         
          .Column5.ControlSource='jobKurs.kse'                  
          .Column6.ControlSource='jobKurs.dkurs'  
          .Column7.ControlSource='jobKurs.pol1'  
          .Column8.ControlSource='jobKurs.pol2'  
          .Column9.ControlSource='jobKurs.srzpk'
          .Column10.ControlSource='jobKurs.zpdayk'     
          .Column11.ControlSource='jobKurs.zptotk'  
          .Column1.Width=RettxtWidth('99999')    
          .Column3.Width=RettxtWidth('99999')
          .Column4.Width=RettxtWidth('9999')
          .Column5.Width=RettxtWidth('99999')
          .Column6.Width=RetTxtWidth('999999.99')     
          .Column7.Width=RetTxtWidth('999999')
          .Column8.Width=RetTxtWidth('999999')
          .Column9.Width=RettxtWidth('99999999.99')
          .Column10.Width=RetTxtWidth('999999.99')     
          .Column11.Width=RetTxtWidth('999999.99')     
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=(.Width-.Column1.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width)/2
          .Column3.Width=.Width-.Column1.Width-.Column2.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='ФИО'      
          .Column3.Header1.Caption='Должность'
          .Column4.Header1.Caption='То'
          .Column5.Header1.Caption='Ш.ед.'
          .Column6.Header1.Caption='Дн.отп.'
          .Column7.Header1.Caption='1 пол'
          .Column8.Header1.Caption='2 пол'
          .Column9.Header1.Caption='Ср.зп.'     
          .Column10.Header1.Caption='За 1 день.'  
          .Column11.Header1.Caption='всего'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'
          .Column9.Format='Z'     
          .Column10.Format='Z' 
          .Column11.Format='Z' 
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
          .Column4.AddObject('checkColumn4','checkContainer')
          .Column4.checkColumn4.AddObject('checkMy','checkBox')
          .Column4.CheckColumn4.checkMy.Visible=.T.
          .Column4.CheckColumn4.checkMy.Caption=''
          .Column4.CheckColumn4.checkMy.Left=10
          .Column4.CheckColumn4.checkMy.BackStyle=0
          .Column4.CheckColumn4.checkMy.ControlSource='jobKurs.lkokl'                                                                                                  
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
               .ControlSource='jobKurs.lkokl' 
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
IF USED('jobKurs')
   SELECT jobKurs 
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
PROCEDURE validPodrKurs
SELECT sprpodr
kpsuplkurs=sprpodr.kod
curnamepodr=fpodr.ComboBox1.Value
DO selectKursJob
fpodr.Refresh
********************************************************************************************************************************************************
PROCEDURE selectKursJob
SELECT jobKurs
SET ORDER TO 1
DELETE ALL
APPEND FROM datJob FOR kp=kpsuplkurs
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
SELECT rasp
SET ORDER TO 2
SET FILTER TO kp=kpsuplkurs
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
         SELECT jobKurs
         APPEND BLANK
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH kse_cx,;
                    np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,avtVac WITH .T.,nvac WITH 1
         REPLACE pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2           
         tar_ok=0
         tar_ok=varBaseSt*jobKurs.namekf*IIF(pkf#0,pkf,1)                      
         
         REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2) 
         REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                 mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse  
                 
         ON ERROR DO erSup        
            REPLACE pzdrav WITH rasp.pzdrav, mzdrav WITH mtokl/100*pzdrav          
         ON ERROR         
        
         REPLACE srzpk WITH rasp.srzpk,zpdayk WITH rasp.zpdayk,dkurs WITH rasp.dkurs,zptotk WITH rasp.zptotk,lkokl WITH rasp.lkokl,pol1 WITH rasp.pol1,pol2 WITH rasp.pol2
         FOR i=1 TO 12                        
             repz='rasp.zk'+LTRIM(STR(i))
             repz1='zpk'+LTRIM(STR(i))              
             REPLACE &repz1 WITH &repz
         ENDFOR                                  
      ENDIF       
   ENDIF
   SELECT rasp
   SKIP
ENDDO
SELECT jobKurs
GO TOP
********************************************************************************************************************************************************
PROCEDURE readzamkurs
srZp_cx=0
kse_cx=0
IF rashkurs(4)
   SELECT jobKurs
   oldRec=RECNO()
   logOkl=lkOkl
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

IF jobKurs.avtVac
   SELECT rasp
   SEEK STR(jobKurs.kp,3)+STR(jobKurs.kd,3)    
ELSE 
   SELECT datJob
   SET ORDER TO 6
   SEEK STR(jobKurs.kodpeop,4)+STR(jobKurs.kp,3)+STR(jobKurs.kd,3)+STR(jobKurs.tr,1)
   *------- Страховка от идентичных записей (kp+kd+kse+tr)
   IF nrkurs#jobKurs.nrkurs
      SEEK STR(jobKurs.kodpeop,4)
      SCAN WHILE kodPeop=jobKurs.kodpeop
           IF nrkurs=jobKurs.nrkurs
              EXIT 
           ENDIF 
      ENDSCAN      
   ENDIF
ENDIF
SELECT jobKurs
IF !rashkurs(4)
   srZp_cx=&formulaOtp
ENDIF
REPLACE srzpk WITH srZp_cx,datjob.srzpk WITH jobKurs.srzpk
doljname=IIF(SEEK(jobKurs.kd,'sprdolj',1),ALLTRIM(sprdolj.namework),'')+' '+LTRIM(STR(datjob.kse,5,2))
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Редактирование - '+ALLTRIM(jobKurs.fio)
     .procExit='DO exitReadZam'
     .Width=840
     DO addShape WITH 'fSupl',1,10,10,0,820,8
     DO adTboxAsCont WITH 'fSupl','txtPodr',.Shape1.Left+10,.Shape1.Top+10,800,dHeight,doljName,2,1,.T.
     DO adTboxAsCont WITH 'fSupl','txtOkl',.txtPodr.Left,.txtPodr.Top+.txtPodr.Height-1,RetTxtWidth('wокладw'),dHeight,'оклад',2,1,.T.      
     
     .AddObject('checkOkl','checkContainer')
     WITH .checkOkl          
          .Left=.Parent.txtPodr.Left
          .Top=.Parent.txtOkl.Top+.Parent.txtOkl.Height-1
          .Height=dHeight
          .Width=.Parent.txtOkl.Width
          .BorderWidth=1
          .Visible=.T.  
          .AddObject('checkMy','MycheckBox')
          WITH .checkMy          
               .Caption=''              
               .BackStyle=0
               .ControlSource='jobKurs.lkokl'                 
               .Visible=.T.   
                Top=(.Parent.Height-.Height)/2
               .Left=(.Parent.Width-.Width)/2                                                                                           
               .procForValid='DO procCheckOkl'
          ENDWITH   
          .BorderWidth=1
     ENDWITH   
    
     DO adTboxAsCont WITH 'fSupl','txtDotp',.txtOkl.Left+.txtOkl.Width-1,.txtOkl.Top,(.txtPodr.Width-.txtOkl.Width)/6,dHeight,'дни курсов',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxDotp',.txtDOtp.Top+.txtDotp.Height-1,.txtDOtp.Left,.txtDotp.Width,dHeight,'jobKurs.dkurs','Z',.T.,2,.F.,'DO validzamdkurs'
     
     DO adTboxAsCont WITH 'fSupl','txt1pol',.txtDotp.Left+.txtDotp.Width-1,.txtDotp.Top,.txtDotp.Width,dHeight,'1 пол',2,1,.T.
     DO adTboxNew WITH 'fSupl','box1pol',.boxDotp.Top,.txt1pol.Left,.txt1pol.Width,dHeight,'jobKurs.pol1','Z',.T.,2,.F.,'DO validdaypol'
     
     DO adTboxAsCont WITH 'fSupl','txt2pol',.txt1pol.Left+.txt1pol.Width-1,.txtDotp.Top,.txtDotp.Width,dHeight,'2 пол',2,1,.T.
     DO adTboxNew WITH 'fSupl','box2pol',.boxDotp.Top,.txt2pol.Left,.txt2pol.Width,dHeight,'jobKurs.pol2','Z',.T.,2,.F.,'DO validdaypol'
     
     DO adTboxAsCont WITH 'fSupl','txtZpl',.txt2pol.Left+.txt2pol.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'зарплата',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxZpl',.boxDotp.Top,.txtZpl.Left,.txtZpl.Width,dHeight,'jobKurs.srzpk','Z',.T.,2
     
     DO adTboxAsCont WITH 'fSupl','txtZpDay',.txtZpl.Left+.txtZpl.Width-1,.txtdOtp.Top,.txtDotp.Width,dHeight,'за 1 день',2,1,.T.  
     DO adTboxNew WITH 'fSupl','boxZpDay',.boxDotp.Top,.txtZpDay.Left,.txtZpDay.Width,dHeight,'jobKurs.zpDayk','Z',.T.,2
     
     DO adTboxAsCont WITH 'fSupl','txtZpTot',.txtZpDay.Left+.txtZpDay.Width-1,.txtdOtp.Top,.txtPodr.Width-.txtOkl.Width-.txtDotp.Width*5+6,dHeight,'всего',2,1,.T.  
     DO adTboxNew WITH 'fSupl','boxZpTot',.boxDotp.Top,.txtZpTot.Left,.txtZpTot.Width,dHeight,'jobKurs.zptotk','Z',.F.,2
     
     DO adTboxAsCont WITH 'fSupl','txtMonth',.txtPodr.Left,.boxDotp.Top+.boxdOtp.Height-1,.txtOkl.Width+.txtDotp.Width+.txt1pol.Width+.txt2pol.Width+.txtZpl.Width+.txtZpDay.Width+.txtZpTot.Width-6,dHeight,'по месяцам',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month1',.txtMonth.Left,.txtMonth.Top+.txtMonth.Height-1,(.txtMonth.Width)/12+1,dHeight,'1',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month2',.Month1.Left+.month1.Width-1,.Month1.Top,.Month1.Width,dHeight,'2',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month3',.Month2.Left+.month2.Width-1,.Month1.Top,.Month1.Width,dHeight,'3',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month4',.Month3.Left+.month3.Width-1,.Month1.Top,.Month1.Width,dHeight,'4',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month5',.Month4.Left+.month4.Width-1,.Month1.Top,.Month1.Width,dHeight,'5',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month6',.Month5.Left+.month5.Width-1,.Month1.Top,.Month1.Width,dHeight,'6',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month7',.Month6.Left+.month6.Width-1,.Month1.Top,.Month1.Width,dHeight,'7',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month8',.Month7.Left+.month7.Width-1,.Month1.Top,.Month1.Width,dHeight,'8',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month9',.Month8.Left+.month8.Width-1,.Month1.Top,.Month1.Width,dHeight,'9',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month10',.Month9.Left+.month9.Width-1,.Month1.Top,.Month1.Width,dHeight,'10',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month11',.Month10.Left+.month10.Width-1,.Month1.Top,.Month1.Width,dHeight,'11',2,1,.T.  
     DO adTboxAsCont WITH 'fSupl','Month12',.Month11.Left+.month11.Width-1,.Month1.Top,.txtMonth.Width-.Month1.Width*11+11,dHeight,'12',2,1,.T.  
     
     DO adTboxNew WITH 'fSupl','m1',.Month1.Top+.month1.Height-1,.Month1.Left,.Month1.Width,dHeight,'jobKurs.zpk1','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m2',.m1.Top,.Month2.Left,.Month2.Width,dHeight,'jobKurs.zpk2','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m3',.m1.Top,.Month3.Left,.Month3.Width,dHeight,'jobKurs.zpk3','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m4',.m1.Top,.Month4.Left,.Month4.Width,dHeight,'jobKurs.zpk4','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m5',.m1.Top,.Month5.Left,.Month5.Width,dHeight,'jobKurs.zpk5','Z',.T.,2  
     DO adTboxNew WITH 'fSupl','m6',.m1.Top,.Month6.Left,.Month6.Width,dHeight,'jobKurs.zpk6','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m7',.m1.Top,.Month7.Left,.Month7.Width,dHeight,'jobKurs.zpk7','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m8',.m1.Top,.Month8.Left,.Month8.Width,dHeight,'jobKurs.zpk8','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m9',.m1.Top,.Month9.Left,.Month9.Width,dHeight,'jobKurs.zpk9','Z',.T.,2  
     DO adTboxNew WITH 'fSupl','m10',.m1.Top,.Month10.Left,.Month10.Width,dHeight,'jobKurs.zpk10','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m11',.m1.Top,.Month11.Left,.Month11.Width,dHeight,'jobKurs.zpk11','Z',.T.,2      
     DO adTboxNew WITH 'fSupl','m12',.m1.Top,.Month12.Left,.Month12.Width,dHeight,'jobKurs.zpk12','Z',.T.,2          
     
     .txtPodr.Width=.txtMonth.Width 
     .Shape1.Width=.txtPodr.Width+20
     .Shape1.Height=dHeight*6+20
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE procCheckOkl
SELECT jobKurs

IF rashkurs(4)
   SELECT jobKurs
   logOkl=lkOkl
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
*nms=&formulaOtp     
REPLACE srzpk WITH IIF(kse<=1,nms,nms),zpdayk WITH srzpk/rashkurs(1)*IIF(rashkurs(4),kse,1)   

*REPLACE srzp WITH IIF(kse<=1,nms,nms),zpday WITH IIF(rashotp(1),srzp/rashotp(2),IIF(dotp#0,srzp/dotp,0))

IF !jobKurs.avtVac
   REPLACE datJob.srzpk WITH jobKurs.srZpk,datJob.zpdayk WITH jobKurs.zpDayk,datJob.lkOkl WITH jobKurs.lkOkl
ELSE
   REPLACE rasp.srzpk WITH jobKurs.srZpk,rasp.zpdayk WITH jobKurs.zpDayk,rasp.lkOkl WITH jobKurs.lkOkl,rasp.ksekurs WITH jobKurs.kse    
ENDIF   
KEYBOARD '{TAB}'
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE validZamDkurs
REPLACE zpdayk WITH srzpk/rashkurs(1)*IIF(rashkurs(4),kse,1)
IF !jobKurs.avtVac
   REPLACE datJob.dkurs WITH jobKurs.dkurs,datJob.zpdayk WITH jobKurs.zpdayk  
ELSE 
   REPLACE rasp.dkurs WITH jobKurs.dkurs,rasp.zpdayk WITH jobKurs.zpdayk,rasp.ksekurs WITH jobKurs.kse 
ENDIF 
IF datjob.pol1=0.AND.datjob.pol2=0
    summonth=jobKurs.zpdayk*jobKurs.dKurs/12
    REPLACE zpk1 WITH summonth,zpk2 WITH summonth,zpk3 WITH summonth,zpk4 WITH summonth,zpk5 WITH summonth,zpk6 WITH summonth
    REPLACE zpk7 WITH summonth,zpk8 WITH summonth,zpk9 WITH summonth,zpk10 WITH summonth,zpk11 WITH summonth,zpk12 WITH summonth
    REPLACE zptotk WITH zpk1+zpk2+zpk3+zpk4+zpk5+zpk6+zpk7+zpk8+zpk9+zpk10+zpk11+zpk12
ENDIF
fSupl.Refresh
****************************************************************************************************************************************************
*                      Расчёт сумм по месяцам
****************************************************************************************************************************************************
PROCEDURE validdaypol
SELECT jobKurs
IF pol1>0.OR.pol2>0
   sum1pol=zpdayk*pol1/6
   sum2pol=zpdayk*pol2/6
   REPLACE zpk1 WITH sum1pol,zpk2 WITH sum1pol,zpk3 WITH sum1pol,zpk4 WITH sum1pol,zpk5 WITH sum1pol,zpk6 WITH sum1pol
   REPLACE zpk7 WITH sum2pol,zpk8 WITH sum2pol,zpk9 WITH sum2pol,zpk10 WITH sum2pol,zpk11 WITH sum2pol,zpk12 WITH sum2pol
   REPLACE zptotk WITH zpk1+zpk2+zpk3+zpk4+zpk5+zpk6+zpk7+zpk8+zpk9+zpk10+zpk11+zpk12
ELSE 
   summonth=jobKurs.zpdayk*jobKurs.dKurs/12
   REPLACE zpk1 WITH summonth,zpk2 WITH summonth,zpk3 WITH summonth,zpk4 WITH summonth,zpk5 WITH summonth,zpk6 WITH summonth
   REPLACE zpk7 WITH summonth,zpk8 WITH summonth,zpk9 WITH summonth,zpk10 WITH summonth,zpk11 WITH summonth,zpk12 WITH summonth
   REPLACE zptotk WITH zpk1+zpk2+zpk3+zpk4+zpk5+zpk6+zpk7+zpk8+zpk9+zpk10+zpk11+zpk12
ENDIF    

tot_cx=0
FOR  h=1 TO 12
      rep_cx='zpk'+LTRIM(STR(h))
*   
 *     IF h<MONTH(rashset(7))
    *     REPLACE &rep_cx WITH 0           
    *  ELSE
         tot_cx=tot_cx+ EVALUATE('zpk'+LTRIM(STR(h)))  
   *   ENDIF        
ENDFOR 
REPLACE zptotk WITH tot_cx
fSupl.Refresh
********************************************************************************************************************************************************
PROCEDURE exitReadZam
IF jobKurs.dkurs=0
   REPLACE srzpk WITH 0,zpdayk WITH 0 
   IF !jobKurs.avtVac
      REPLACE datJob.zpdayk WITH 0,datjob.srzpk WITH 0      
   ELSE 
      REPLACE rasp.zpdayk WITH 0,rasp.srzpk WITH 0,rasp.ksekurs WITH 0      
   ENDIF 
ELSE
   IF !jobKurs.avtVac
      REPLACE datJob.zpk1 WITH zpk1,datJob.zpk2 WITH zpk2,datJob.zpk3 WITH zpk3,datJob.zpk4 WITH zpk4,;
              datJob.zpk5 WITH zpk5,datJob.zpk6 WITH zpk6,datJob.zpk7 WITH zpk7,datJob.zpk8 WITH zpk8,;              
              datJob.zpk9 WITH zpk9,datJob.zpk10 WITH zpk10,datJob.zpk11 WITH zpk11,datJob.zpk12 WITH zpk12
      REPLACE datjob.pol1 WITH pol1,datjob.pol2 WITH pol2,datjob.zptotk WITH zptotk        
   ELSE
      REPLACE rasp.zk1 WITH zpk1,rasp.zk2 WITH zpk2,rasp.zk3 WITH zpk3,rasp.zk4 WITH zpk4,;
              rasp.zk5 WITH zpk5,rasp.zk6 WITH zpk6,rasp.zk7 WITH zpk7,rasp.zk8 WITH zpk8,;
              rasp.zk9 WITH zpk9,rasp.zk10 WITH zpk10,rasp.zk11 WITH zpk11,rasp.zk12 WITH zpk12
      REPLACE rasp.pol1 WITH jobKurs.pol1,rasp.pol2 WITH jobKurs.pol2,rasp.zptotk WITH jobKurs.zptotk        
      
   ENDIF   
ENDIF
SELECT datJob
SET ORDER TO 2
SELECT jobKurs
*************************************************************************************************************
PROCEDURE deletefromkurs
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
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+5,.check1.Top+.check1.Height+15,(.shape1.Width-20)/2,dHeight+3,'Выполнение','DO delreckurs'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fdel.Release' 
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+.check1.Height+50    
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
*****************************************************************************************************************************************************
*                         Непосредственно удаление информации по замене
*****************************************************************************************************************************************************
PROCEDURE delreckurs
IF !log_del 
   RETURN 
ENDIF
DO CASE
   CASE dim_del(1)=1
        SELECT datJob
        SET ORDER TO 6
        SEEK STR(jobKurs.kodpeop,4)+STR(jobKurs.kp,3)+STR(jobKurs.kd,3)+STR(jobKurs.tr,1)
        FOR i=1 TO 12
            repzp='zpk'+LTRIM(STR(i))
            REPLACE &repzp WITH 0
        ENDFOR                      
        REPLACE zptotk WITH 0,zpdayk WITH 0,srzpk WITH 0,dkurs WITH 0,pol1 WITH 0,pol2 WITH 0
       
   CASE dim_del(2)=1
        SELECT rasp
        SET FILTER TO 
        REPLACE zptotk WITH 0,zpdayk WITH 0,srzpk WITH 0,dkurs WITH 0,pol1 WITH 0,pol2 WITH 0 FOR kp=kpsuplkurs
        REPLACE zk1 WITH 0,zk2 WITH 0,zk3 WITH 0,zk4 WITH 0,zk5 WITH 0,zk6 WITH 0,zk7 WITH 0,zk8 WITH 0,zk9 WITH 0,zk10 WITH 0,zk11 WITH 0,zk12 WITH 0 FOR kp=kpsuplkurs
        SELECT datJob
        SET FILTER TO kp=kpsuplkurs
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12             
               repzp='zpk'+LTRIM(STR(i))
               REPLACE &repzp WITH 0
           ENDFOR                      
           REPLACE zptotk WITH 0,zpdayk WITH 0,srzpk WITH 0,dkurs WITH 0,pol1 WITH 0,pol2 WITH 0
           SKIP
        ENDDO         
              
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO 
        REPLACE zptotk WITH 0,zpdayk WITH 0,srzpk WITH 0,dkurs WITH 0,pol1 WITH 0,pol2 WITH 0 ALL 
        REPLACE zk1 WITH 0,zk2 WITH 0,zk3 WITH 0,zk4 WITH 0,zk5 WITH 0,zk6 WITH 0,zk7 WITH 0,zk8 WITH 0,zk9 WITH 0,zk10 WITH 0,zk11 WITH 0,zk12 WITH 0 ALL 
        SELECT datJob 
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12            
               repzp='zpk'+LTRIM(STR(i))
               REPLACE &repzp WITH 0
           ENDFOR                      
           REPLACE zptotk WITH 0,zpdayk WITH 0,srzpk WITH 0,dkurs WITH 0,pol1 WITH 0,pol2 WITH 0
           SKIP
        ENDDO           
                  
ENDCASE
fdel.Release
SELECT jobKurs
DO selectKursJob
fpodr.Refresh
***********************************************************************************************************************************************
PROCEDURE storedimdel
PARAMETERS par1
FOR i=1 TO 4
    dim_del(i)=IIF(i=par1,1,0)
ENDFOR
fdel.Refresh
*****************************************************************************************************************************************************
*                         Форма для общего расчёта сведений по замене
*****************************************************************************************************************************************************
PROCEDURE formCountRash
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
     DO adtbox WITH 'fdel',1,.lab1.Left+.lab1.Width+10,.Shape1.Top+10,RetTxtWidth('99/99/999999'),dHeight,'countDate','Z',.T.,1
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
        SEEK STR(jobKurs.kodpeop,4)+STR(jobKurs.kp,3)+STR(jobKurs.kd,3)+STR(jobKurs.tr,1)
        max_rec=1 
        DO countone
        one_pers=one_pers+1
        pers_ch=one_pers/max_rec*100
        fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch 
   CASE dim_del(2)=1
        SELECT datJob
        SET FILTER TO kp=kpsuplkurs
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
        SET FILTER TO ksekurs#0.AND.kp=kpsuplkurs
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
        SET FILTER TO ksekurs#0
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
DO selectKursJob
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
IF dkurs#0   
   **------зарплата и зарплата в день (если указано "пересчитывать спеднюю зарплату")
   IF log_srzp
      STORE 0 TO nms,srms
      nms=0
      kse_cx=0 
      IF rashkurs(4)        
         SELECT datjob
         kpOld=kp
         kdOld=kd
         logOkl=lkOkl
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
                    ROUND(varBaseSt/100*pmain*ksevac,2)+ROUND(varBaseSt/100*pmain2*ksevac,2)+IIF(rashkurs(3)#0,tar_ok*rashkurs(3),0)                     
            ELSE
               nms=nms+tar_ok
            ENDIF         
            kse_cx=kse_cx+ksevac  
         ELSE  
             REPLACE ksekurs WITH 0,zptotk WITH 0
         ENDIF        
         nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx))
        
         SELECT datjob
         SET ORDER TO &oldJobOrd
         GO fjobrec
      ELSE    
         nms=&formulaOtp
      ENDIF      
     * nms=&formulaOtp         
      REPLACE srzpk WITH nms,zpdayk WITH srzpk/rashkurs(1)*IIF(rashkurs(4),kse,1)
   ENDIF 
   
   *-------перерасчёт помесячно 
   IF pol1>0.OR.pol2>0 
      sum1pol=zpdayk*pol1/6
      sum2pol=zpdayk*pol2/6   
      REPLACE zpk1 WITH sum1pol,zpk2 WITH sum1pol,zpk3 WITH sum1pol,zpk4 WITH sum1pol,zpk5 WITH sum1pol,zpk6 WITH sum1pol
      REPLACE zpk7 WITH sum2pol,zpk8 WITH sum2pol,zpk9 WITH sum2pol,zpk10 WITH sum2pol,zpk11 WITH sum2pol,zpk12 WITH sum2pol
   ELSE 
      summonth=zpdayk*dKurs/12
      REPLACE zpk1 WITH summonth,zpk2 WITH summonth,zpk3 WITH summonth,zpk4 WITH summonth,zpk5 WITH summonth,zpk6 WITH summonth
      REPLACE zpk7 WITH summonth,zpk8 WITH summonth,zpk9 WITH summonth,zpk10 WITH summonth,zpk11 WITH summonth,zpk12 WITH summonth
   ENDIF    
   FOR h=1 TO 12
       rep_cx='zpk'+LTRIM(STR(h))
       IF h<MONTH(countDate)
          REPLACE &rep_cx WITH 0     
       ENDIF              
   ENDFOR 
   REPLACE zptotk WITH zpk1+zpk2+zpk3+zpk4+zpk5+zpk6+zpk7+zpk8+zpk9+zpk10+zpk11+zpk12   
ELSE
   REPLACE zptotk WITH 0,srzpk WITH 0,zpdayk WITH 0
   REPLACE zpk1 WITH 0,zpk2 WITH 0,zpk3 WITH 0,zpk4 WITH 0,zpk5 WITH 0,zpk6 WITH 0,zpk7 WITH 0,zpk8 WITH 0,zpk9 WITH 0,zpk10 WITH 0,zpk11 WITH 0,zpk12 WITH 0               
ENDIF
********************************************************************************************************************************************************
*                     Процедура расчёта расходов на замену по одной должности
********************************************************************************************************************************************************
PROCEDURE countonevac
SELECT rasp
IF ksekurs#0   
   **------зарплата и зарплата в день (если указано "пересчитывать спеднюю зарплату")
   
    DO CASE 
      CASE raskurs(4)
           logOkl=lkOkl
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
              IF !rasp.lkokl        
                  nms=nms+tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksevac,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksevac,2)+;
                      ROUND(varBaseSt/100*pmain*ksezotp,2)+ROUND(varBaseSt/100*pmain2*ksevac,2)+IIF(rashkurs(3)#0,tar_ok*rashkurs(3),0)                                                                                                 
              ELSE
                 nms=nms+tar_ok
              ENDIF         
              kse_cx=kse_cx+ksevac  
           ENDIF          
           nms=IIF(kse_cx=0,0,IIF(kse_cx<1,nms,nms/kse_cx) )   
           REPLACE srzpk WITH nms,zpdayk WITH srzpk/rashkurs(1)*IIF(rashkurs(4),kse,1)    
      CASE !rashkurs(4)
           STORE 0 TO nms,srms 
           tar_ok=0
           tar_ok=ROUND(varBaseSt*rasp.nkfvac*IIF(pkf#0,pkf,1)*ksekurs,2)                      
           nms=tar_ok
           IF !rasp.lkOkl
              nms=tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksekurs,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksekurs,2)+;
                  ROUND(varBaseSt/100*pmain*ksekurs,2)+ROUND(varBaseSt/100*pmain2*ksekurs,2)+IIF(rashkurs(3)#0,tar_ok*rashkurs(3),0)
             
           ELSE 
              nms=tar_ok
           ENDIF         
           REPLACE srzpk WITH nms,zpdayk WITH srzpk/rashkurs(1)     
   ENDCASE 

   *-------перерасчёт помесячно   
   IF pol1>0.OR.pol2>0   
      sum1pol=zpdayk*pol1/6
      sum2pol=zpdayk*pol2/6   
      REPLACE zk1 WITH sum1pol,zk2 WITH sum1pol,zk3 WITH sum1pol,zk4 WITH sum1pol,zk5 WITH sum1pol,zk6 WITH sum1pol
      REPLACE zk7 WITH sum2pol,zk8 WITH sum2pol,zk9 WITH sum2pol,zk10 WITH sum2pol,zk11 WITH sum2pol,zk12 WITH sum2pol
   ELSE 
      summonth=zpdayk*dKurs/12
      REPLACE zk1 WITH summonth,zk2 WITH summonth,zk3 WITH summonth,zk4 WITH summonth,zk5 WITH summonth,zk6 WITH summonth
      REPLACE zk7 WITH summonth,zk8 WITH summonth,zk9 WITH summonth,zk10 WITH summonth,zk11 WITH summonth,zk12 WITH summonth
   ENDIF     
   FOR h=1 TO 12
       rep_cx='zk'+LTRIM(STR(h))
       IF h<MONTH(countDate)
          REPLACE &rep_cx WITH 0     
       ENDIF              
   ENDFOR 
   REPLACE zptotk WITH zk1+zk2+zk3+zk4+zk5+zk6+zk7+zk8+zk9+zk10+zk11+zk12  
ELSE
   REPLACE zptotk WITH 0,srzpk WITH 0,zpdayk WITH 0
   REPLACE zk1 WITH 0,zk2 WITH 0,zk3 WITH 0,zk4 WITH 0,zk5 WITH 0,zk6 WITH 0,zk7 WITH 0,zk8 WITH 0,zk9 WITH 0,zk10 WITH 0,zk11 WITH 0,zk12 WITH 0               
ENDIF

***********************************************************************************************************************************************
PROCEDURE printkurs
fSupl=CREATEOBJECT('FORMSUPL')
DIMENSION dimOpt(3)
dimOpt(1)=1
dimOpt(2)=0
dimOpt(3)=0

logWord=.F.
kvo_page=1
page_beg=1
page_end=999
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
DO CASE 
   CASE dimOpt(1)=1
        SELECT rasp
        SET FILTER TO 
        SELECT * FROM datJob INTO CURSOR curPrn READWRITE
        ALTER TABLE curPrn ADD COLUMN nIt N(1)
        ALTER TABLE curPrn ADD COLUMN nVac N(1)
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
        ALTER TABLE curPrn ALTER COLUMN zptotk N(11,2)
        ON ERROR DO erSup
           ALTER TABLE curPrn ADD COLUMN pzdrav N(3)
           ALTER TABLE curPrn ADD COLUMN mzdrav N(10,2)
        ON ERROR 
		      
        SELECT curprn
        DELETE FOR zptotk=0
        REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL 
        REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL 
        INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
        INDEX ON STR(kp,3)+STR(kd,3) TAG T2

        SELECT rasp
        SET FILTER TO ksekurs#0.AND.zptotk#0
        GO TOP
        DO WHILE !EOF()
           SELECT curprn
           APPEND BLANK
           REPLACE kp WITH rasp.kp,kd WITH rasp.kd,np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,fio WITH 'Вакантная',dkurs WITH rasp.dkurs,zpdayk WITH rasp.zpdayk,srzpk WITH rasp.srzpk,;
                   zptotk WITH rasp.zptotk,zpk1 WITH rasp.zk1,zpk2 WITH rasp.zk2,zpk3 WITH rasp.zk3,zpk4 WITH rasp.zk4,zpk5 WITH rasp.zk5,zpk6 WITH rasp.zk6,;
                   zpk7 WITH rasp.zk7,zpk8 WITH rasp.zk8,zpk9 WITH rasp.zk9,zpk10 WITH rasp.zk10,zpk11 WITH rasp.zk11,zpk12 WITH rasp.zk12
           REPLACE kse WITH rasp.ksekurs,nVac WITH 1        
           SELECT rasp
           SKIP
        ENDDO
        SET FILTER TO 
        SELECT sprpodr
        SCAN ALL
             SELECT curprn
             SUM kse,zptotk,zpk1,zpk2,zpk3,zpk4,zpk5,zpk6,zpk7,zpk8,zpk9,zpk10,zpk11,zpk12 TO ksecx,zptotcx,zp1cx,zp2cx,zp3cx,zp4cx,zp5cx,zp6cx,zp7cx,zp8cx,zp9cx,zp10cx,zp11cx,zp12cx FOR kp=sprpodr.kod    
             IF zptotcx#0
                APPEND BLANK
                REPLACE kp WITH sprpodr.kod,np WITH sprpodr.np,nd WITH 98,fio WITH 'по отделению',nIt WITH 1,zptotk WITH zptotcx,kse WITH ksecx
                REPLACE zpk1 WITH zp1cx,zpk2 WITH zp2cx,zpk3 WITH zp3cx,zpk4 WITH zp4cx,zpk5 WITH zp5cx,zpk6 WITH zp6cx,zpk7 WITH zp7cx,zpk8 WITH zp8cx,zpk9 WITH zp9cx,zpk10 WITH zp10cx,zpk11 WITH zp11cx,zpk12 WITH zp12cx
             ENDIF
             SELECT sprpodr     
        ENDSCAN
        SELECT curprn  
        SUM kse,zptotk,zpk1,zpk2,zpk3,zpk4,zpk5,zpk6,zpk7,zpk8,zpk9,zpk10,zpk11,zpk12 TO ksecx,zptotcx,zp1cx,zp2cx,zp3cx,zp4cx,zp5cx,zp6cx,zp7cx,zp8cx,zp9cx,zp10cx,zp11cx,zp12cx FOR nIt=0
        APPEND BLANK
        REPLACE np WITH 999,nd WITH 98,fio WITH 'по организации',nIt WITH 9,zptotk WITH zptotcx,kse WITH ksecx
        REPLACE zpk1 WITH zp1cx,zpk2 WITH zp2cx,zpk3 WITH zp3cx,zpk4 WITH zp4cx,zpk5 WITH zp5cx,zpk6 WITH zp6cx,zpk7 WITH zp7cx,zpk8 WITH zp8cx,zpk9 WITH zp9cx,zpk10 WITH zp10cx,zpk11 WITH zp11cx,zpk12 WITH zp12cx
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
          DO procForPrintAndPreview WITH 'repzamkurs','',.T.,'kursToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'repzamkurs','',.F. 
        ENDIF   
   CASE dimOpt(2)=1  
        SELECT rasp
        SET FILTER TO 
        SELECT * FROM rasp INTO CURSOR curprn READWRITE
        ALTER TABLE curPrn ADD COLUMN nIt N(1)
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ALTER COLUMN ksezotp N(7,2)
        ALTER TABLE curPrn ALTER COLUMN dkurs N(8)
        ALTER TABLE curPrn ALTER COLUMN pol1 N(8)
        ALTER TABLE curPrn ALTER COLUMN pol2 N(8)
        ALTER TABLE curPrn ALTER COLUMN zpdayk N(11,2)
        ALTER TABLE curPrn ALTER COLUMN zptotk N(11,2)
        ON ERROR DO erSup
           ALTER TABLE curPrn ADD COLUMN pzdrav N(3)
           ALTER TABLE curPrn ADD COLUMN mzdrav N(10,2)
        ON ERROR 
        
        DELETE ALL    
        SELECT * FROM rasp INTO CURSOR curprn1 READWRITE  
        SELECT curprn1
        DELETE FOR zptotk=0
        REPLACE kse WITH ksezotp ALL
               
        SELECT * FROM datjob WHERE dkurs>0.AND.zptotk#0.AND.dkurs>0 INTO CURSOR curJobPodr READWRITE
        SELECT curJobPodr
        APPEND FROM DBF('curprn1')
        INDEX ON STR(kp,3)+STR(kd,3)+STR(zpdayk,8,2) TAG T1
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
                   REPLACE dkurs WITH dkurs+curjobpodr.dkurs,zpdayk WITH curjobpodr.zpdayk,ksezotp WITH ksezotp+curjobpodr.kse,zptotk WITH zptotk+curjobpodr.zptotk                
                   SELECT curjobpodr        
                   SKIP
                   IF kp=rasp.kp.AND.kd=rasp.kd.AND.zpdayk#curprn.zpdayk
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
        DELETE FOR dkurs=0
       * REPLACE zptotk WITH zpdayk*dkurs ALL
         
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
        REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,0) ALL
        INDEX ON STR(np,3)+STR(nd,3) TAG T1
        INDEX ON STR(kp,3)+STR(kd,3) TAG T2
        SET ORDER TO 1
        
        SELECT sprpodr
        SCAN ALL
             SELECT curprn
             SUM ksezotp,zptotk,dkurs TO ksezotp_cx,zptotk_cx,dkurs_cx FOR kp=sprpodr.kod
             APPEND BLANK 
             REPLACE kp WITH sprpodr.kod,nd WITH 999,np WITH sprpodr.np,zptotk WITH zptotk_cx,dkurs WITH dkurs_cx,nIt WITH 1,;
                     ksezotp WITH ksezotp_cx,named WITH 'по отделению'            
             SELECT sprpodr
        ENDSCAN
        SELECT curprn
        SUM ksezotp,zptotk,dkurs TO ksezotp_cx,zptotk_cx,dkurs_cx FOR nIt=1
        APPEND BLANK 
        REPLACE kp WITH sprpodr.kod,nd WITH 999,np WITH 999,zptotk WITH zptotk_cx,dkurs WITH dkurs_cx,nIt WITH 3,;
                ksezotp WITH ksezotp_cx,named WITH 'Итого'           
                
        DELETE FOR zptotk=0     
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
           DO procForPrintAndPreview WITH 'zamkursdol','',.T.,'zamKursDolToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'zamkursdol','',.F. 
        ENDIF    
   CASE dimOpt(3)=1  
        repzp1='rasp.zk'+LTRIM(STR(MONTH(varDtar)))
        repzp='curprn.zpk'+LTRIM(STR(MONTH(varDtar)))
        repd='curprn.d'+LTRIM(STR(MONTH(varDtar)))
        repd1='rasp.m'+LTRIM(STR(MONTH(varDtar)))
        IF USED('curprn')
           SELECT curPrn
           USE   
        ENDIF
        IF USED('curJobSup')
           SELECT curJobsup
           USE   
        ENDIF

        IF USED('curOtp')
           SELECT curOtp
           USE
        ENDIF

        CREATE CURSOR curotp (nstr C(15),kse N(7,2),mtokl N(10,2),mtokl1 N(10,2),mtokl2 N(10,2),mstsum N(10,2),mchir N(10,2),mkat N(10,2),mvto N(10,2),mcharw N(10,2),mmain N(10,2),mmain2 N(10,2),mzdrav N(10,2),msupl N(10,2),mtot N(10,2),mRound N(10,2))
        SELECT rasp
        SET FILTER TO 
        SELECT datjob
        SET FILTER TO
        SELECT * FROM datjob INTO CURSOR curPrn READWRITE
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ADD COLUMN nIt N(1)
        ALTER TABLE curPrn ADD COLUMN kHours N(7,2)
        ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
        ON ERROR DO erSup
           ALTER TABLE curPrn ADD COLUMN pzdrav N(3)
           ALTER TABLE curPrn ADD COLUMN mzdrav N(10,2)
        ON ERROR 
        
        SELECT curPrn     
        DELETE FOR tokl=0
        DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
        DELETE FOR date_in>varDtar
        
        REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
        REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
        REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL
*----------------------------------- Автоматическое добавление вакансий---------------------------------------------
*IF formbase.avt_vac
        SELECT rasp
        SET FILTER TO ksezotp>0
        GO TOP
        DO WHILE !EOF()
           IF rasp.ksezotp#0        
              SELECT curPrn
              APPEND BLANK
              REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH rasp.ksezotp,vac WITH .T.,;
                      np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,tr WITH 1,lkokl WITH rasp.lkokl
                 
              tar_ok=0
              tar_ok=varBaseSt*namekf*IIF(pkf#0,pkf,1)                      
              REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2),;
                      pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2             
          
              REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                      mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse        
              REPLACE srzpk WITH rasp.srzpk,zpdayk WITH rasp.zpdayk,dkurs WITH rasp.dkurs,zptotk WITH rasp.zptotk,lkokl WITH rasp.lkokl,pol1 WITH rasp.pol1,pol2 WITH rasp.pol2
              ON ERROR DO erSup        
                 REPLACE pzdrav WITH rasp.pzdrav, mzdrav WITH mtokl/100*pzdrav          
              ON ERROR 
              FOR i=1 TO 12                        
                  repz='rasp.zk'+LTRIM(STR(i))
                  repz1='zpk'+LTRIM(STR(i))              
                  REPLACE &repz1 WITH &repz
              ENDFOR                        
           ENDIF
           SELECT rasp
           SKIP
        ENDDO
*ENDIF  
        SELECT rasp
        SET FILTER TO 
        SELECT * FROM curprn INTO CURSOR curJobSupl READWRITE
        SELECT curJobSupl
        INDEX ON STR(kp,3)+STR(kd,3) TAG T1
        SELECT curprn 
        IF !rashkurs(4)
           SELECT curotp
           FOR i=1 TO 12
               APPEND BLANK
               REPLACE nstr WITH dim_month(i)
           ENDFOR
           SELECT curprn           
           SCAN ALL
                SELECT curprn
                polday=IIF(curprn.pol1>0.OR.curprn.pol2>0,6,12)
                polbeg=IIF(polday=12,1,IIF(pol1>0,1,7))
                IF zptotk>0           
                   FOR i=polbeg TO 12
                       STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvOkl1,nAvOkl2,nAvZdrav
                       repzp='curprn.zpk'+LTRIM(STR(i))                             
                       DO CASE 
                          CASE curprn.lkokl.AND.&repzp#0
                              *nAvOkl=&repzp
                               nAvOkl1=&repzp
                          CASE !curprn.lkokl.AND.&repzp#0
                               nAvOkl=mtokl/rashkurs(1)*curprn.dkurs/polday
                               nAvOkl2=nAvOkl
                               nAvSt=mstsum/rashkurs(1)*curprn.dkurs/polday
                               nAvChir=mchir/rashkurs(1)*curprn.dkurs/polday
                               nAvKat=mkat/rashkurs(1)*curprn.dkurs/polday
                               nAvVto=mvto/rashkurs(1)*curprn.dkurs/polday
                               nAvCharw=mcharw/rashkurs(1)*curprn.dkurs/polday
                               nAvMain=mmain/rashkurs(1)*curprn.dkurs/polday
                               nAvMain2=mmain2/rashkurs(1)*curprn.dkurs/polday
                               nAvZdrav=mzdrav/rashkurs(1)*curprn.dkurs/polday
                               nAvSupl=mtokl*rashkurs(3)/rashkurs(1)*curprn.dkurs/polday
                       ENDCASE 
                       SELECT curotp
                       GO i
                       REPLACE mtokl WITH mtokl+nAvOkl,mtokl1 WITH mtokl1+nAvOkl1,mtokl2 WITH mtokl2+nAvOkl2,mstsum WITH mstSum+nAvSt,mchir WITH mchir+nAvChir,mzdrav WITH mzdrav+nAvZdrav;
                               mkat WITH mkat+nAvKat,mvto WITH mvto+nAvVto,mcharw WITH mcharw+nAvCharw,mmain WITH mmain+nAvMain,mmain2 WITH mmain2+nAvMain2,msupl WITH msupl+nAvSupl,mtot WITH mtot+&repzp                                                                                    
                       SELECT curprn
                   ENDFOR 
                ENDIF 
           ENDSCAN               
        ELSE 
           FOR i=1 TO 12    
               repzp='curprn.zpk'+LTRIM(STR(i))
               SELECT curotp
               APPEND BLANK
               REPLACE nstr WITH dim_month(i)
               SELECT curprn
               STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvOkl1,nAvOkl2,nAvZdrav
               SCAN ALL
                    IF &repZp#0
                       SELECT curJobsupl
                       SEEK STR(curprn.kp,3)+STR(curprn.kd,3)
                       STORE 0 TO nkse,nAvOkl,nAvOkl1,nAvOkl2,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvZdrav
                       SCAN WHILE kp=curprn.kp.AND.kd=curprn.kd
                            nkse=nkse+kse
                            *nAvOkl=nAvOkl+mtokl                            
                            IF !curprn.lkOkl
                               nAvOkl=nAvOkl+mtokl
                               nAvOkl2=nAvOkl2+mtokl   
                               nAvSt=nAvSt+mstsum
                               nAvChir=nAvChir+mchir
                               nAvKat=nAvKat+mkat
                               nAvVto=nAvVto+mvto
                               nAvCharw=nAvCharw+mcharw
                               nAvMain=nAvMain+mmain
                               nAvMain2=nAvMain2+mmain2
                               nAvZdrav=nAvZdrav+mzdrav
                               nAvSupl=nAvSupl+mtokl*rashkurs(3)
                            ELSE  
                               nAvOkl1=nAvOkl1+mtokl   
                            ENDIF   
                       ENDSCAN           
                       nAvOkl=IIF(nkse=0,0,IIF(nkse<1,nAvOkl,nAvOkl/nkse))
                       nAvOkl1=IIF(nkse=0,0,IIF(nkse<1,nAvOkl1,nAvOkl1/nkse))
                       nAvOkl2=IIF(nkse=0,0,IIF(nkse<1,nAvOkl2,nAvOkl2/nkse))
                             
                       nAvSt=IIF(nkse=0,0,IIF(nkse<1,nAvSt,nAvSt/nkse))
                       nAvChir=IIF(nkse=0,0,IIF(nkse<1,nAvChir,nAvChir/nkse))
                       nAvKat=IIF(nkse=0,0,IIF(nkse<1,nAvKat,nAvKat/nkse))
                       nAvVto=IIF(nkse=0,0,IIF(nkse<1,nAvVto,nAvVto/nkse))
                       nAvCharw=IIF(nkse=0,0,IIF(nkse<1,nAvCharw,nAvCharw/nkse))
                       nAvMain=IIF(nkse=0,0,IIF(nkse<1,nAvMain,nAvMain/nkse))
                       nAvMain2=IIF(nkse=0,0,IIF(nkse<1,nAvMain2,nAvMain2/nkse))
                       nAvZdrav=IIF(nkse=0,0,IIF(nkse<1,nAvZdrav,nAvZdrav/nkse))
                       nAvSupl=IIF(nkse=0,0,IIF(nkse<1,nAvSupl,nAvSupl/nkse))                    
                       polday=IIF(curprn.pol1>0.OR.curprn.pol2>0,6,12)
                       SELECT curotp
                       GO i
                       REPLACE mtokl WITH mtokl+(ROUND(nAvOkl/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mstSum WITH mstsum+(ROUND(nAvSt/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mChir WITH mChir+(ROUND(nAvChir/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,;
                               mKat WITH mKat+(ROUND(nAvKat/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mVto WITH mVto+(ROUND(nAvVto/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mCharw WITH mCharw+(ROUND(nAvCharw/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,;
                               mMain WITH mMain+(ROUND(nAvMain/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mMain2 WITH mmain2+(ROUND(nAvmain2/rashkurs(1),2)*curprn.kse*curprn.dkurs)/12,;
                               mzdrav WITH mzdrav+(ROUND(nAvZdrav/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mSupl WITH mSupl+(ROUND(nAvSupl/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,;
                               mtokl1 WITH mtokl1+(ROUND(nAvOkl1/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mtokl2 WITH mtokl2+(ROUND(nAvOkl2/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday                   
           
                       REPLACE kse WITH kse+curprn.kse,mtot WITH mtot+&repzp 
                       SELECT curprn    
                    ENDIF
               ENDSCAN
           ENDFOR
        ENDIF 
        SELECT curotp
        REPLACE mRound WITH mTot-mtokl1-mTokl-mStsum-mChir-mKat-mVto-mCharw-mMain-mMain2-mzdrav-msupl ALL
        REPLACE mSupl WITH mSupl+mRound ALL 
        REPLACE mtokl2 WITH mtokl+mstsum+mchir+mkat+mvto+mcharw+mmain+mmain2+mzdrav+msupl ALL 
        SUM mtokl,mstsum,mChir,mkat,mVto,mCharw,mMain,mMain2,mSupl,mtot,mRound,kse,mtokl1,mtokl2,mzdrav TO mtokl_cx,mstsum_cx,mChir_cx,mkat_cx,mVto_cx,mCharw_cx,mMain_cx,mMain2_cx,mSupl_cx,mtot_cx,mRound_cx,kse_cx,mtokl1_cx,mtokl2_cx,mzdrav_cx
        APPEND BLANK
        REPLACE nstr WITH 'всего',mtokl WITH mtokl_cx,mStsum WITH mstsum_cx,mChir WITH mChir_cx,mKat WITH mkat_cx,mVto WITH mvto_cx,mCharw WITH mCharw_cx,mMain WITH mMain_cx,mMain2 WITH mMain2_cx,;
                mSupl WITH msupl_cx,mTot WITH mtot_cx,mRound WITH mround_cx,kse WITH kse_cx,mtokl1 WITH mtokl1_cx,mtokl2 WITH mtokl2_cx,mzdrav WITH mzdrav_cx
        GO TOP       
        IF parTerm=1
          DO procForPrintAndPreview WITH 'kursSvod','',.T.,'kursSvodToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'kursSvod','',.F. 
        ENDIF       
        *DO procForPrintAndPreview WITH 'reptototp','итоговая по замене отпусков'
        SELECT curotp
        USE
        SELECT curPrn
        USE
        SELECT people   
ENDCASE       
*************************************************************************************************************
PROCEDURE kursSvodToExcel
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
          .Value='Всего замена курсов'        
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
     
     .Range(.Cells(rowcx,5),.Cells(rowcx,13)).Select  
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
          rowcx=rowcx+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(rowtop,1),.Cells(rowcx-1,14)).Select
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
PROCEDURE kursToExcel
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
     .COLUMNS(10).COLUMNWIDTH=8
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
     .Columns(21).ColumnWidth=8
     .Columns(22).ColumnWidth=8
     
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,22)).Select  
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
     .Range(.Cells(rowcx,1),.Cells(rowcx,22)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Расчёт планируемых расходов на оплату труда работниколв, заменяющих уходящих на курсы повышения квалификации'
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
     .cells(rowcx,5).Value='дни курсов'                                                         
     
     .cells(rowcx,6).Value='1-е пол.'                                                         
     .cells(rowcx,7).Value='2-е пол.'                                                        
     
     .cells(rowcx,8).Value='оклад'                                                         
     .cells(rowcx,9).Value='в день'                                                        
             
     .cells(rowcx,10).Value='1'
     .cells(rowcx,11).Value='2'
     .cells(rowcx,12).Value='3'
     .cells(rowcx,13).Value='4'
     .cells(rowcx,14).Value='5'
     .cells(rowcx,15).Value='6'
     .cells(rowcx,16).Value='7'
     .cells(rowcx,17).Value='8'
     .cells(rowcx,18).Value='9'
     .cells(rowcx,19).Value='10'
     .cells(rowcx,20).Value='11'
     .cells(rowcx,21).Value='12'  
     .cells(rowcx,22).Value='всего'                                                           
  
     .Range(.Cells(rowcx,1),.Cells(rowcx,22)).Select
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
              .Range(.Cells(numberRow,1),.Cells(numberRow,22)).Select
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
          
          .Cells(numberRow,5).Value=IIF(dkurs#0,dkurs,'')                                       
          .Cells(numberRow,5).NumberFormat='0.00'                                                   
          
          .Cells(numberRow,6).Value=IIF(pol1#0,pol1,'')                                       
          .Cells(numberRow,6).NumberFormat='0.00'                                               
          
          .Cells(numberRow,7).Value=IIF(pol2#0,pol2,'')                                       
          .Cells(numberRow,7).NumberFormat='0.00'                                             
                    
          .Cells(numberRow,8).Value=IIF(srzpk#0,srzpk,'')                                       
          .Cells(numberRow,8).NumberFormat='0.00'                                                              
          
          .Cells(numberRow,9).Value=IIF(zpdayk#0,zpdayk,'')                                       
          .Cells(numberRow,9).NumberFormat='0.00'    
                                                                          
          .Cells(numberRow,10).Value=IIF(zpk1#0,zpk1,'')                                       
          .Cells(numberRow,10).NumberFormat='0.00'                                                  
          
          .Cells(numberRow,11).Value=IIF(zpk2#0,zpk2,'')                                       
          .Cells(numberRow,11).NumberFormat='0.00'                                        
          
          .Cells(numberRow,12).Value=IIF(zpk3#0,zpk3,'')                                       
          .Cells(numberRow,12).NumberFormat='0.00'
          
          .Cells(numberRow,13).Value=IIF(zpk4#0,zpk4,'')                                       
          .Cells(numberRow,13).NumberFormat='0.00' 
          
          .Cells(numberRow,14).Value=IIF(zpk5#0,zpk5,'')                                       
          .Cells(numberRow,14).NumberFormat='0.00'
          
          .Cells(numberRow,15).Value=IIF(zpk6#0,zpk6,'')                                       
          .Cells(numberRow,15).NumberFormat='0.00' 
          
          .Cells(numberRow,16).Value=IIF(zpk7#0,zpk7,'')
          .Cells(numberRow,16).NumberFormat='0.00'                                          
          
          .Cells(numberRow,17).Value=IIF(zpk8#0,zpk8,'')
          .Cells(numberRow,17).NumberFormat='0.00'    
          
          .Cells(numberRow,18).Value=IIF(zpk9#0,zpk9,'')
          .Cells(numberRow,18).NumberFormat='0.00'    
          
          .Cells(numberRow,19).Value=IIF(zpk10#0,zpk10,'')
          .Cells(numberRow,19).NumberFormat='0.00'    
          
          .Cells(numberRow,20).Value=IIF(zpk11#0,zpk11,'')
          .Cells(numberRow,20).NumberFormat='0.00' 
             
          .Cells(numberRow,21).Value=IIF(zpk12#0,zpk12,'')
          .Cells(numberRow,21).NumberFormat='0.00'    
          
          .Cells(numberRow,22).Value=IIF(zptotk#0,zptotk,'')
          .Cells(numberRow,22).NumberFormat='0.00' 

          numberRow=numberRow+1
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch                    
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,22)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1      
      
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,22)).Select
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
PROCEDURE zamKursDolToExcel
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
          .Value='Расчёт планируемых расходов на оплату лиц, заменяющих уходящих на курсы работников'
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
                                                                                   
          .Cells(numberRow,4).Value=IIF(dkurs#0,dkurs,'')    
          
          .Cells(numberRow,5).Value=IIF(zpdayk#0,zpdayk,'')                                       
          .Cells(numberRow,5).NumberFormat='0.00'                                        
          
          .Cells(numberRow,6).Value=IIF(zptotk#0,zptotk,'')                                       
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
***************************************************************************************************************************************************
*                   Процедура для настрек по работе с заменой курсов
***************************************************************************************************************************************************
PROCEDURE setupkurs
fsetup=CREATEOBJECT('FORMMY')
WITH fsetup
     .BackColor=RGB(255,255,255)
      DO addShape WITH 'fSetup',1,10,10,dHeight,350,8      
     .procexit='DO exitsetupkurs'       
    
     DO adLabMy WITH 'fsetup',1,'Среднее кол-во дней',.Shape1.Top+20,.Shape1.Left+20,150,0,.T.
     DO adtbox WITH 'fsetup',1,.Lab1.Left+.lab1.Width+5,.lab1.Top,RetTxtWidth('9999999'),dHeight,'rashkurs(1)','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashkurs'
     
     .lab1.Top=.txtbox1.Top+(.txtbox1.Height-.lab1.Height)
     DO adLabMy WITH 'fsetup',2,'Дата отсчёта',.txtBox1.Top+.txtBox1.Height+10,.lab1.Left,150,0,.T.
     DO adtbox WITH 'fsetup',2,fsetup.lab2.Left+fSetup.lab2.Width+5,.txtbox1.Top+.txtBox1.Height+10,RetTxtWidth('99/99/999999'),dHeight,'countDate','Z',.T.,1
     
     DO adLabMy WITH 'fsetup',3,'Коэффициент уравнивания',.txtBox2.Top+.txtBox2.Height+10,.lab1.Left,150,0,.T.
     DO adtbox WITH 'fsetup',3,fsetup.lab3.Left+fSetup.lab3.Width+5,.txtbox2.Top+.txtBox2.Height+10,RetTxtWidth('99999999'),dHeight,'rashkurs(3)','Z',.T.,1
     
     DO adCheckBox WITH 'fsetup','checkOklad','использовать средний оклад',.txtBox3.Top+.txtBox3.Height+10,.Shape1.Left+20,150,dHeight,'rashkurs(4)',0,.F.,'SAVE TO &var_path ALL LIKE rashkurs' 
   *  .Shape1.Width=.checkOklad.Width+40    
    
     .txtBox3.InputMask='9.9999'
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
     .Shape1.Height=.checkOklad.Height+.txtBox1.Height*3+80 
   

     .Caption='настройки'   
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20    
ENDWITH
DO pasteImage WITH 'fsetup'
fsetup.Show
**************************************************************************************************************************
PROCEDURE exitsetupkurs
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
SAVE TO &var_path ALL LIKE rashkurs
fsetup.Release