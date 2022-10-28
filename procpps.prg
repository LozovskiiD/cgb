IF USED('jobOtp')
   SELECT jobOtp
   USE
ENDIF
=AFIELDS(arJob,'datjob')
CREATE CURSOR jobOtp FROM ARRAY arJob
ALTER TABLE jobOtp ADD COLUMN avtVac L
ALTER TABLE jobOtp ADD COLUMN nVac N(1)
SELECT jobOtp
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 1

SELECT datShtat
LOCATE FOR ALLTRIM(pathTarif)=pathTarSupl
STORE 0 TO dim_tot,dim_day
formulapps='mtokl+mkonts+mstsum+mvto+mkat+mint+mmols+mosob+mboss+mchir+mcharw+mmain+mmain2+mglav+mprem+mtokl*rashpps'
*formulaotp2='IIF(!logOkl,mtokl+mstsum+mvto+mkat+mchir+mcharw+mmain+mmain2+IIF(rashotp(3)#0,mtokl*rashotp(3),0),mtokl)'
RESTORE FROM rashpps ADDITIVE
SELECT datJob
SET FILTER TO 
*REPLACE nrotp WITH RECNO() ALL 
SET ORDER TO 2
SELECT rasp
SET FILTER TO 
ordOldRasp=SYS(21)
SET ORDER TO 2
SELECT sprpodr 
oldOrdPodr=SYS(21)
SET ORDER TO 2
GO TOP
kppps=kod
curnamepodr=name
DO selectPps
var_path=FULLPATH('rashpps.mem')
fPodr=CREATEOBJECT('FORMSPR')
WITH fPodr
     .Caption='Расчет планируемых отпусков на выплату расходов по ППС'   
     DO addButtonOne WITH 'fPodr','menuCont1',10,5,'редакция','pencil.ico','Do readpps',39,RetTxtWidth('календарь')+44,'редакция'    
 *    DO addButtonOne WITH 'fPodr','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'удаление','pencild.ico','Do deletefrompps',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fPodr','menuCont3',.menucont1.Left+.menucont1.Width+3,5,'расчёт','calculate.ico','DO formcountpps',39,.menucont1.Width,'расчёт'       
     DO addButtonOne WITH 'fPodr','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico','DO printPps',39,.menucont1.Width,'печать' 
     DO addButtonOne WITH 'fPodr','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'настройки','setup.ico','DO setupPPs',39,.menucont1.Width,'настройки'  
     DO addButtonOne WITH 'fPodr','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'возврат','undo.ico','DO exitFromPps',39,.menucont1.Width,'возврат'       
            
     DO addComboMy WITH 'fPodr',1,(.Width-500)/2,.menucont1.Top+.menucont1.Height+5,dheight,500,.T.,'curnamepodr','sprpodr.name',6,.F.,'DO validPodrPps',.F.,.T.      
     .comboBox1.DisplayCount=25
     WITH .fGrid    
          .Top=fpodr.ComboBox1.Top+fpodr.ComboBox1.Height+5 
          .Height=.Parent.Height-.Top
          .Width=.Parent.Width
          .RecordSource='jobOtp'
          DO addColumnToGrid WITH 'fPodr.fGrid',8
          .RecordSourceType=1     
          .Column1.ControlSource='jobOtp.kodpeop'
          .Column2.ControlSource='jobOtp.fio'          
          .Column3.ControlSource="IIF(SEEK(jobOtp.kd,'sprdolj',1),sprdolj.name,'')"
          .Column4.ControlSource='jobOtp.kse'                  
          .Column5.ControlSource='jobOtp.nopps'
          .Column6.ControlSource='jobOtp.npps'
          .Column7.ControlSource='jobOtp.nspps'
          .Column1.Width=RettxtWidth('99999')    
          .Column4.Width=RettxtWidth('999999')
          .Column5.Width=RetTxtWidth('9999999.99')     
          .Column6.Width=RetTxtWidth('9999.99')     
          .Column7.Width=RettxtWidth('9999999.99')
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=(.Width-.Column1.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width)/2
          .Column3.Width=.Width-.Column1.Width-.Column2.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='ФИО'      
          .Column3.Header1.Caption='Должность'
          .Column4.Header1.Caption='Ш.ед.'
          .Column5.Header1.Caption='База'
          .Column6.Header1.Caption='%'
          .Column7.Header1.Caption='Сумма'     
          .Column4.Format='Z'     
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'                     
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column3.Alignment=0         
          .Column4.Alignment=1         
          .Column5.Alignment=1         
          .Column6.Alignment=1         
          .Column7.Alignment=1  
           .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .colNesInf=2              
     ENDWITH
     DO gridSizeNew WITH 'fpodr','fGrid','shapeingrid'
     
ENDWITH
fPodr.Show
********************************************************************************************************************************************************
PROCEDURE validPodrPps
SELECT sprpodr
kppps=sprpodr.kod
curnamepodr=fpodr.ComboBox1.Value
DO selectPps
fpodr.Refresh
********************************************************************************************************************************************************
PROCEDURE selectPps
SELECT jobOtp
SET ORDER TO 1
DELETE ALL
APPEND FROM datJob FOR kp=kppps
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
SELECT rasp
SET ORDER TO 2
SET FILTER TO kp=kppps
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
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,kse WITH kse_cx,;
                    np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0),nd WITH rasp.nd,avtVac WITH .T.,nvac WITH 1
         REPLACE pkont WITH rasp.pkont,pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2,posob WITH rasp.posob,;
                 pglav WITH rasp.pglav,pprem WITH rasp.pprem,pboss WITH rasp.pboss 
         
         *formulapps='mtokl+mkonts+mstsum+mvto+mkat+mint+mmols+mosob+mboss+mchir+mcharw+mmain+mmain2+mglav+mprem+mtokl*rashpps'
                  
                    
         tar_ok=0
         tar_ok=varBaseSt*jobOtp.namekf                      
         REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2)          
           
         REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                 mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse,nopps WITH rasp.nopps,npps WITH rasp.npps,nspps WITH rasp.nspps                    
                                           
      ENDIF       
   ENDIF
   SELECT rasp
   SKIP
ENDDO
SELECT jobOtp
GO TOP

********************************************************************************************************************************************************
PROCEDURE readpps
nBaseZp=0
SELECT jobOtp
nBaseZp=&formulaPps
REPLACE nopps WITH nBaseZp,datjob.nopps WITH jobOtp.nopps
*doljname=IIF(SEEK(jobOtp.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' '+LTRIM(STR(datjob.kse,5,2))
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Редактирование - '+ALLTRIM(jobOtp.fio)
     .procExit='DO exitReadPps'
     .Width=840
     DO addShape WITH 'fSupl',1,10,10,0,820,8
     DO adTboxAsCont WITH 'fSupl','txtOkl',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('ww9999999999ww'),dHeight,'база',2,1,.T.                
     DO adTboxNew WITH 'fSupl','boxOkl',.txtOkl.Top+.txtOkl.Height-1,.txtOkl.Left,.txtOkl.Width,dHeight,'jobOtp.nopps','Z',.F.,2,.F.
    
     DO adTboxAsCont WITH 'fSupl','txtPers',.txtOkl.Left+.txtOkl.Width-1,.txtOkl.Top,RetTxtWidth('w%ППМw'),dHeight,'% ППС',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxPers',.boxOkl.Top,.txtPers.Left,.txtPers.Width,dHeight,'jobOtp.npps','Z',.T.,2,.F.,'DO validpps'
     
     DO adTboxAsCont WITH 'fSupl','txtSum',.txtPers.Left+.txtPers.Width-1,.txtOkl.Top,.txtOkl.Width,dHeight,'сумма',2,1,.T.
     DO adTboxNew WITH 'fSupl','boxSum',.boxOkl.Top,.txtSum.Left,.txtSum.Width,dHeight,'jobOtp.nspps','Z',.F.,2,.F.

     .Shape1.Width=.txtOkl.Width+.txtPers.Width+.txtSum.Width+20 
     .Shape1.Height=dHeight*2+20
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************************************************************************************
PROCEDURE validPps
REPLACE nspps WITH nopps/100*npps
fSupl.boxSum.Refresh
*******************************************************************************************************************************************************
PROCEDURE exitReadPps
IF jobOtp.avtVac
   SELECT rasp
   SEEK STR(jobOtp.kp,3)+STR(jobOtp.kd,3) 
   REPLACE npps WITH jobOtp.npps,nspps WITH jobOtp.nspps,nopps WITH IIF(npps#0,jobotp.nopps,0),ksepps WITH IIF(npps=0,0,jobotp.kse)
 ELSE 
   SELECT datjob
   SET ORDER TO 7
   SEEK jobotp.nid
   REPLACE npps WITH jobOtp.npps,nspps WITH jobOtp.nspps,nopps WITH IIF(npps#0,jobotp.nopps,0)
ENDIF 
***************************************************************************************************************************************************
*                   Процедура для настрек по работе с заменой отпусков
***************************************************************************************************************************************************
PROCEDURE setuppps
fsetup=CREATEOBJECT('FORMMY')
WITH fsetup
     .BackColor=RGB(255,255,255)
      DO addShape WITH 'fSetup',1,10,10,dHeight,0,8      
     .procexit='DO exitfsetuppps'       
         
     DO adLabMy WITH 'fsetup',3,'Для расчёта использовать коэффициент -',.Shape1.Top+20,.Shape1.Left+20,150,0,.T.
     DO adtbox WITH 'fsetup',3,fsetup.lab3.Left+fSetup.lab3.Width+5,.Shape1.Top+20,RetTxtWidth('99999999'),dHeight,'rashpps','Z',.T.,1
     .txtBox3.InputMask='9.9999'
     .lab3.Top=.txtbox3.Top+(.txtbox3.Height-.lab3.Height)  
     .Shape1.Width=.lab3.Width+.txtBox3.Width+45         
     .Shape1.Height=.txtBox3.Height+40
     .Caption='настройки'   
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20    
ENDWITH
DO pasteImage WITH 'fsetup'
fsetup.Show
**************************************************************************************************************************
PROCEDURE exitfsetuppps
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
SAVE TO &var_path ALL LIKE rashpps
fsetup.Release
********************************************************************************************************************************************************
PROCEDURE formcountpps
fdel=CREATEOBJECT('FORMSUPL')
WITH fdel
     .Caption='Общий расчёт'
     DO addShape WITH 'fdel',1,20,20,300,300,8

     DO adLabMy WITH 'fdel',1,'выполнить расчёт?',.Shape1.Top+20,.Shape1.Left+5,.Shape1.Width-10,2,.F.
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+(.shape1.Width-RetTxtWidth('wвозвратw')*2-20)/2,.lab1.Top+.lab1.Height,RetTxtWidth('wвозвратw'),dHeight+3,'расчёт','DO countpps'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'вовзврат','fdel.Release' 
     .Shape1.height=.cont1.Height+.lab1.Height+40
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+40    
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
********************************************************************************************************************************************************
PROCEDURE countpps
fdel.Release
SELECT datjob
SET FILTER TO 
SCAN ALL
     nBaseZp=&formulaPps
     REPLACE nopps WITH IIF(npps#0,nbasezp,0),nspps WITH nopps/100*npps
ENDSCAN
SELECT rasp
SET FILTER TO 
SCAN ALL
     IF ksepps#0.AND.npps#0
        STORE 0 TO nms,srms 
        tar_ok=0
        tar_ok=ROUND(varBaseSt*rasp.nkfvac*ksepps,2)                      
        nms=tar_ok      
        nms=tar_ok+ROUND(varBaseSt/100*dimConstVac(2,2)*ksepps,2)+ROUND(tar_ok/100*pkat,2)+ROUND(tar_ok/100*pvto,2)+ROUND(tar_ok/100*pchir,2)+ROUND(varBaseSt/100*pcharw*ksepps,2)+;
        ROUND(varBaseSt/100*pmain*ksepps,2)+ROUND(varBaseSt/100*pmain2*ksepps,2)+ROUND(tar_ok*pprem/100,2)+ROUND(tar_ok*pkont/100,2)+ROUND(tar_ok*posob/100,2)+ROUND(tar_ok*rashpps,2)
          
        REPLACE nopps WITH nms,nspps WITH nopps/100*npps
     ELSE
        REPLACE nopps WITH 0,nspps WITH 0
     
     
     ENDIF
ENDSCAN
DO selectpps
fPodr.Refresh
********************************************************************************************************************************************************
PROCEDURE printpps
********************************************************************************************************************************************************
PROCEDURE exitFromPps
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