IF USED('salpeop')
   SELECT salpeop
   USE
ENDIF
monthzpl=''
labMonth=''
salPath=''
logStaj=.F.
logTotOkl=.F.
IF FILE('pathpeop.mem')
   RESTORE FROM pathpeop ADDITIVE
   patharcmonth=LEFT(ALLTRIM(pathpeop),LEN(ALLTRIM(pathpeop))-10)+'arcmonth.mem'
   IF FILE(patharcmonth)
      USE &pathpeop ORDER 1 ALIAS salpeop
      RESTORE FROM &patharcmonth ADDITIVE
      monthzpl=arcmonth(2)
      labMonth=monthzpl
      salPath=LEFT(ALLTRIM(pathpeop),LEN(ALLTRIM(pathpeop))-10)
      SELECT salpeop
      SET FILTER TO LEFT(tabfio,7)=monthzpl
   ENDIF    
ELSE   
   patharcmonth=''
ENDIF    
IF USED('jobZpl')
   SELECT jobZpl
   USE   
ENDIF
IF !USED('datagrup')
   USE datagrup IN 0
ENDIF
IF !USED('tarifzp')
   USE tarifzp IN 0
ENDIF
SELECT * FROM datJob INTO CURSOR jobZpl READWRITE
SELECT jobZpl
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,fio) ALL 
INDEX ON fio TAG T1
INDEX ON tabn TAG T2
SET ORDER TO 2
DIMENSION dimFault(3)
STORE .T. TO dimFault
DIMENSION dimstru(2)
dimstru(1)=1
dimstru(2)=0
DIMENSION dimbg(3)
STORE 0 TO dimbg
dimbg(1)=1

CREATE CURSOR proterror (kodpeop N(5),kod N(1), ntab N(5), fio C(70), strerror C(254),kodEr N(1),tr N(1),kse N(5,2),kp N(3),kd N(3)) 
SELECT proterror
GO TOP
fTrans=CREATEOBJECT('FORMSPR')
WITH fTrans   
     .Caption='Перенос в ПО "Учёт труда и заработной "'
     .procExit='DO exitFromProcToZpl'
     DO addButtonOne WITH 'fTrans','menuCont1',10,5,'контроль','check.ico','DO checkTarifToZpl',39,RetTxtWidth('wтаб.номера')+44,'контроль'  
     DO addButtonOne WITH 'fTrans','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'перенос','cash.ico','DO formNewTarif',39,.menucont1.Width,'перенос'             
     DO addButtonOne WITH 'fTrans','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'настройки','setup.ico','DO setupPathZpl',39,.menucont1.Width,'настройки'             
     DO addButtonOne WITH 'fTrans','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'таб.номера','numbers.ico','DO tabzptotrf',39,.menucont1.Width,'присвоение табельных номеров'             
     DO addButtonOne WITH 'fTrans','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'печать','print1.ico',"DO printreport WITH 'repproterror','протокол возможных ошибок','proterror'",39,.menucont1.Width,'печать' 
     DO addButtonOne WITH 'fTrans','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'редакция','pencil.ico','DO readTrfZpl',39,.menucont1.Width,'возврат' 
     DO addButtonOne WITH 'fTrans','menuCont7',.menucont6.Left+.menucont6.Width+3,5,'возврат','undo.ico','DO exitFromProcToZpl',39,.menucont1.Width,'возврат' 
     
     DO addButtonOne WITH 'fTrans','butRet',.menucont1.Left,5,'возврат','undo.ico','DO exitFromReatTrfZpl',39,RetTxtWidth('wвозврат')+44,'возврат' 
     .butRet.Visible=.F.
     
     DO addShape WITH 'fTrans',1,5,.menucont1.Height+20,100,100,8
     DO adCheckBox WITH 'fTrans','check1','ошибки тарификации',.Shape1.Top+10,.Shape1.Left+20,200,dHeight,'dimFault(1)',0,.T.
     DO adCheckBox WITH 'fTrans','check2','тарификация-зарплата',.check1.Top+.check1.Height+5,.check1.Left,.check1.Width,dHeight,'dimFault(2)',0,.T.
     DO adCheckBox WITH 'fTrans','check3','зарплата-тарификация',.check2.Top+.check2.Height+5,.check1.Left,.check1.Width,dHeight,'dimFault(3)',0,.T.
     .Shape1.Height=.check1.Height*3+25
     .Shape1.Width=.check3.Width+40   
     DO adLabMy WITH 'fTrans',2,' контроль ошибок ',.Shape1.Top-10,.Shape1.Left+5,150,0,.T.,1   
     
     DO addShape WITH 'fTrans',3,.Shape1.Left+.Shape1.Width+10,.Shape1.Top,100,100,8   
     
     
     DO addOptionButton WITH 'fTrans',4,'Организация',.check1.Top,.Shape3.Left+20,'dimstru(1)',0,'DO procvalidstru WITH 1',.T.
     DO addOptionButton WITH 'fTrans',5,'Совокупность',.check1.Top,.Option4.Left+.Option4.Width+10,'dimstru(2)',0,'DO procvalidstru WITH 2',.T.        
     DO addcombomy WITH 'fTrans',1,.Option4.Left,.check2.Top,dHeight,.Option4.Width+.Option5.Width+10 ,.F.,'','datagrup.name',6,.F.,.F.,.F.,.T. 
     
     DO addOptionButton WITH 'fTrans',7,'бюджет',.check1.Top,.comboBox1.Left+.comboBox1.Width+20,'dimbg(1)',0,'DO procvaliddimbg WITH 1',.T.
     DO addOptionButton WITH 'fTrans',8,'внебюджет',.check2.Top,.Option7.Left,'dimbg(2)',0,'DO procvaliddimbg WITH 2',.T.
     DO addOptionButton WITH 'fTrans',9,'внебюджет 2',.check3.Top,.Option8.Left,'dimbg(3)',0,'DO procvaliddimbg WITH 3',.T.        
     .Shape3.Height=.Shape1.Height+5
     .Shape3.Width=.comboBox1.Width+.Option9.Width+60                    
     DO adLabMy WITH 'fTrans',4,' структура ',.lab2.Top,.Shape3.Left+5,150,0,.T.,1
         
     DO addShape WITH 'fTrans',4,.Shape3.Left+.Shape3.Width+10,.Shape1.Top,100,100,8 
     DO adLabMy WITH 'fTrans',1,'расчётный месяц  - '+monthZpl,.check1.Top,.Shape4.Left+10,150,0,.T.  
     DO adLabMy WITH 'fTrans',11,'Путь в ПО "Учёт труда и заработной платы"  - '+salPath,.check2.Top,.lab1.Left,150,0,.T.  
     DO adCheckBox WITH 'fTrans','check4','переносить стаж',.check3.Top,.lab1.Left,.check1.Width,dHeight,'logStaj',0,.T.
     DO adCheckBox WITH 'fTrans','checkTot','оклад целиком',.check4.Top,.check4.Left+.check4.Width+10,.check1.Width,dHeight,'logTotOkl',0,.T.
     
     DO adLabMy WITH 'fTrans',5,' дополнительно ',.lab4.Top,.Shape4.Left+4,150,0,.T.,1
     .Shape4.Height=.Shape1.Height
     .Shape4.Width=.Width-.Shape1.Width-.Shape3.Width-48              
 
      
     WITH .fGrid
          .Top=.Parent.Shape1.Top+.Parent.Shape1.Height+5
          .Height=.Parent.Height-.Parent.Shape1.Height-5                 
          .RecordSourceType=1     
          .RecordSource='proterror'
           .ColumnCount=0
           DO addColumnToGrid WITH 'fTrans.fGrid',6
          .Left=0          
          .ScrollBars=2
          .Column1.ControlSource='proterror.fio' 
          .Column2.ControlSource='proterror.ntab'                 
          .Column3.ControlSource='proterror.tr'       
          .Column4.ControlSource='proterror.kse'          
          .Column5.ControlSource='proterror.strerror'
          
          .Column1.Header1.Caption='Фамилия Имя Отчество'
          .Column2.Header1.Caption='Таб№'
          .Column3.Header1.Caption='Тип'
          .Column4.Header1.Caption='Объём'
          .Column5.Header1.Caption='Возможная ошибка '
          
          .Column1.Width=RetTxtWidth('Венниаминов Венниамин Венниаминович')        
          .Column2.Width=RetTxtWidth('9999999')
          .Column3.Width=RetTxtWidth('совместw')
          .Column4.Width=RetTxtWidth('wобъёмw')
          .Column5.Width=.Width-.Column1.Width-.Column2.Width-.Column3.Width-.Column4.Width-.ColumnCount-Sysmetric(5)-13          
          .Columns(.ColumnCount).Width=0  

          .Column2.Format='Z'
          .Column3.Format='Z'
          .Column4.Format='Z'
          
          .Column2.Sparse=.T.
          .Column3.Sparse=.T.
          
          .Column1.Alignment=0
          .Column2.Alignment=1          
          .Column3.Alignment=1
          .Column4.Alignment=1          
          .Column5.Alignment=0
     ENDWITH
     DO myColumnTxtBox WITH 'fTrans.fGrid.column2','txtbox2','proterror.nTab',.F.,.F.,.F.,'DO validTabJob' 
     DO myColumnTxtBox WITH 'fTrans.fGrid.column3','txtbox3','protError.tr',.F.,.F.,.F.,'DO validTrJob' 
     DO gridSizeNew WITH 'fTrans','fGrid','shapeingrid'     
      
     .Show
ENDWITH 
*************************************************************
PROCEDURE exitFromProcToZpl
ON ERROR DO erSup
fTrans.Visible=.F.
fTrans.Release
ON ERROR
*******************************************************************************************************************
PROCEDURE procChangeDir
PARAMETERS par_find
pathcopy=''
pathpeop_cx=LEFT(pathpeop,RAT('\',pathpeop))                 
pathcopy=GETDIR(pathpeop_cx,'','Укажите папку ПО "Учёт труда и заработной платы"',64)  
ON ERROR DO erSup
IF !EMPTY(pathcopy)
   labMonth=''
   salPath=''
   pathpeop=pathcopy+'people.dbf'
   IF USED('salpeop')
      SELECT salpeop
      USE 
   ENDIF    
   patharcmonth=ALLTRIM(pathcopy)+'arcmonth.mem'
   IF FILE(patharcmonth)       
      USE &pathpeop ORDER 1 ALIAS salpeop IN 0
      RESTORE FROM &patharcmonth ADDITIVE
      monthZpl=arcMonth(2)
      labMonth=monthzpl 
      salPath=ALLTRIM(pathcopy)
      SELECT salpeop
      SET FILTER TO LEFT(tabfio,7)=monthzpl        
   ENDIF
ENDIF
var_path=FULLPATH('pathpeop.mem')
SAVE TO &var_path ALL LIKE pathpeop
ON ERROR 
fTrans.lab1.Caption='расчётный месяц  - '+monthZpl
fTrans.lab11.Caption='Путь в ПО "Учёт труда и заработной платы"  - '+salPath
SELECT proterror
DELETE ALL
GO TOP
fTrans.Refresh
fSupl.Refresh
*************************************************************
PROCEDURE checkTarifToZpl
SELECT jobZpl
SET ORDER TO 2
SELECT proterror
DELETE ALL

*---------------------------- Проверка на ошибки в тарификации (1)----------------------
IF dimFault(1)
   APPEND BLANK 
   REPLACE strerror WITH 'Возможные ошибки в тарификации',kod WITH 1
   SELECT jobzpl
   GO TOP   
   DO WHILE !EOF()
      log_ap=.F.
      repstr=''
      IF !vac
         IF tabn=0
            repstr=IIF(!EMPTY(repstr),ALLTRIM(repstr)+', отсутствует табельный номер','Отсутствует табельный номер')           
            log_ap=.T.
         ENDIF     
         IF tr=0  
            repstr=IIF(!EMPTY(repstr),ALLTRIM(repstr)+', не указан тип работы','Не указан тип работы')            
            log_ap=.T.             
         ENDIF 
         IF kse=1.AND.tr#1     
            repstr=IIF(!EMPTY(repstr),ALLTRIM(repstr)+', возможно не правильно указан тип работы ('+'совмес-во - '+STR(jobzpl.kse,4,2)+'ст.)',;
                   'Возможно не правильно указан тип работы ('+'совмес-во - '+STR(jobzpl.kse,4,2)+'ст.)')                      
            log_ap=.T.       
         ENDIF
         IF tr=1.AND.kse<0.7
            repstr=IIF(!EMPTY(repstr),ALLTRIM(repstr)+', возможно не правильно указан тип работы ('+'основная - '+STR(jobzpl.kse,4,2)+'ст.)',;
                   'Возможно не правильно указан тип работы ('+'основная - '+STR(jobzpl.kse,4,2)+'ст.)')              
            log_ap=.T.               
         ENDIF
         IF 'вакантн'$LOWER(fio)
            repstr=IIF(!EMPTY(repstr),ALLTRIM(repstr)+', возможно должность является вакантной',;
                    'Возможно должность является вакантной')                                      
            log_ap=.T.   
         ENDIF 
      ELSE
         IF !'вакантн'$LOWER(fio)
            repstr=IIF(!EMPTY(repstr),ALLTRIM(repstr)+', возможно должность не является вакантной',;
                    'Возможно должность не является вакантной')                             
            log_ap=.T.    
         ENDIF 
      ENDIF 
      IF log_ap
         SELECT proterror
         APPEND BLANK         
         REPLACE ntab WITH jobzpl.tabn,fio WITH jobzpl.fio,strerror WITH repstr   
         REPLACE kodEr WITH 1,tr WITH jobZpl.tr,kse WITH jobZpl.kse,kp WITH jobZpl.kp,kd WITH jobZpl.kd,kodpeop WITH jobzpl.kodpeop                         
         SELECT jobzpl  
      ENDIF         
      SKIP 
   ENDDO 
ENDIF   
*------------------------------------Проверка на соответстыие тарификация-зарплата (2)
IF dimFault(2).AND.USED('salpeop')
   SELECT proterror
   APPEND BLANK 
   REPLACE strerror WITH 'Возможные несоответствия между тарификацией и зарплатой',kod WITH 2
   SELECT salpeop
   SET ORDER TO 1
   SELECT jobzpl
   GO TOP
   DO WHILE !EOF()  
      IF !vac 
         SELECT salpeop
         SEEK STR(jobzpl.tabn,5)
         DO CASE
            CASE !FOUND()
                 SELECT proterror 
                 APPEND BLANK
                 REPLACE ntab WITH jobzpl.tabn,fio WITH jobzpl.fio,strerror WITH 'Сотрудников с таким табельным номером нет в списке зарплаты'
                 REPLACE kodEr WITH 2,tr WITH jobZpl.tr,kse WITH jobZpl.kse,kp WITH jobZpl.kp,kd WITH jobZpl.kd,kodpeop WITH jobzpl.kodpeop                    
                 SELECT jobzpl
            CASE FOUND().AND.LOWER(ALLTRIM(jobzpl.fio))#LOWER(ALLTRIM(SUBSTR(salpeop.tabFio,13,30)))+' '+LOWER(ALLTRIM(SUBSTR(salpeop.tabFio,43,25)))+' '+LOWER(ALLTRIM(SUBSTR(salpeop.tabFio,68,25)))
                 SELECT proterror 
                 APPEND BLANK
                 REPLACE ntab WITH jobzpl.tabn,fio WITH jobzpl.fio,strerror WITH 'Возможно данный табельный номер ('+LTRIM(SUBSTR(salpeop.tabFio,8,5))+') принадлежит '+ALLTRIM(SUBSTR(salpeop.tabFio,13,30))+' '+ALLTRIM(SUBSTR(salpeop.tabFio,43,25))+' '+ALLTRIM(SUBSTR(salpeop.tabFio,68,25))
                 REPLACE kodEr WITH 2,tr WITH jobZpl.tr,kse WITH jobZpl.kse,kp WITH jobZpl.kp,kd WITH jobZpl.kd,kodpeop WITH jobzpl.kodpeop          
         ENDCASE         
         SELECT jobzpl
      ENDIF    
      SKIP   
   ENDDO
ENDIF   
*------------------------------------Проверка на соответствие зарплата-тарификация (3)
IF dimFault(3).AND.USED('salpeop')
   SELECT proterror
   APPEND BLANK
   REPLACE strerror WITH 'Возможные несоответствия между зарплатой и тарификацией',kod WITH 3
   SELECT salpeop
   GO TOP
   DO WHILE !EOF()
      SELECT jobzpl
      SEEK ROUND(VAL(SUBSTR(salpeop.tabFio,8,5)),0)   
      DO CASE
         CASE !FOUND()
              SELECT proterror 
              APPEND BLANK
              REPLACE ntab WITH VAL(SUBSTR(salpeop.tabFio,8,5)),fio WITH ALLTRIM(SUBSTR(salpeop.tabFio,13,30))+' '+ALLTRIM(SUBSTR(salpeop.tabFio,43,25))+' '+ALLTRIM(SUBSTR(salpeop.tabFio,68,25)),strerror WITH 'Сотрудников с таким табельным номером нет в списке тарификации'
              SELECT salpeop
         CASE FOUND().AND.LOWER(ALLTRIM(jobzpl.fio))#LOWER(ALLTRIM(SUBSTR(salpeop.tabFio,13,30)))+' '+LOWER(ALLTRIM(SUBSTR(salpeop.tabFio,43,25)))+' '+LOWER(ALLTRIM(SUBSTR(salpeop.tabFio,68,25)))
              SELECT proterror 
              APPEND BLANK
              REPLACE ntab WITH VAL(SUBSTR(salpeop.tabFio,8,5)),fio WITH ALLTRIM(SUBSTR(salpeop.tabFio,13,30))+' '+ALLTRIM(SUBSTR(salpeop.tabFio,43,25))+' '+ALLTRIM(SUBSTR(salpeop.tabFio,68,25)),;
                      strerror WITH 'Возможно данный табельный номер ('+LTRIM(SUBSTR(salpeop.tabFio,8,5))+') принадлежит '+ALLTRIM(jobzpl.fio)
      ENDCASE 
      SELECT salpeop
      SKIP
   ENDDO
ENDIF   
SELECT proterror
GO TOP 
fTrans.Refresh  
*************************************************************
PROCEDURE newTarifToZpl
*******************************************************************************************************************
PROCEDURE procvalidstru
PARAMETERS par_ch
STORE 0 TO dimstru
dimstru(par_ch)=1
fTrans.comboBox1.Enabled=IIF(par_ch=1,.F.,.T.)
fTrans.Refresh
*******************************************************************************************************************
PROCEDURE procvaliddimbg
PARAMETERS par1
STORE 0 TO dimbg
dimbg(par1)=1
fTrans.Refresh
*******************************************************************************************************************
PROCEDURE setupPathZpl
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Дополнительные настройки'
     .MinButton=.F.
     .MaxButton=.F.
     .backColor=RGB(255,255,255)
     DO addShape WITH 'fSupl',1,10,10,100,100,8 
     DO addContFormNew WITH 'fSupl','contPath',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('WПуть - зарплатаW'),dHeight,' путь -  зарплата',0,.F.,'DO procChangeDir' 
     DO adtBox WITH 'fSupl',11,.contPath.Left+.contPath.Width-1,.contPath.Top,250,dHeight,'pathPeop',.F.,.F.,0 
          
        
     .Shape1.Width=.contPath.Width+.txtBox11.Width+40
     .Shape1.Height=.txtBox11.Height+40           
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20    
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*********************************************************************************************************************************************************
PROCEDURE formNewTarif
IF !USED('salpeop')
   RETURN
ENDIF
logTrans=.F.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Icon='money.ico'
     .Caption='Перенос окладов в ПО "Учёт труда и заработной платы"'
     .Width=400
     .Height=200
     
     DO addShape WITH 'fSupl',1,20,20,dHeight,RetTxtWidth('WWДля подтверждения намерений поставьте отметкуWW'),8   
     DO adLabMy WITH 'fSupl',1,'Для подтверждения намерений поставьте отметку',.Shape1.Top+20,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fSupl',2,"в окошке 'подтверждение намерений'",.Lab1.Top+.Lab1.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     .Shape1.Height=.lab1.Height*2+45       
      DO adCheckBox WITH 'fSupl','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'logTrans',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wпереносw')*2-20)/2,.check1.Top+.check1.Height+20,RetTxtWidth('wпереносw'),dHeight+3,'перенос','DO newTarifToZpl'
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fSupl.Release'
     .Height=.Shape1.Height+.cont1.Height+.check1.Height+80
     .Width=.Shape1.Width+40  
     
         
     DO addcontlabel WITH 'fSupl','cont3',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WвозвратW'))/2,.cont1.Top,RetTxtWidth('WвозвратW'),dHeight+5,'возврат','fSupl.Release','возврат' 
     .cont3.Visible=.F.
     
      DO addShape WITH 'fSupl',2,.Shape1.Left,.cont1.Top,.cont1.Height,.Shape1.Width,8
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
     DO addShape WITH 'fSupl',3,.Shape2.Left,.Shape2.Top,.Shape2.Height,50,8
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.               
      DO adLabMy WITH 'fSupl',25,'100%',.Shape2.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab25.Visible=.F.
      DO adLabMy WITH 'fSupl',24,'Ход выполнения',.check1.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab24.Visible=.F.
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show     
**********************************************************************************************************************************************************
*                      Непосредственно перенос тарифных величин в зарплату
**********************************************************************************************************************************************************
PROCEDURE newTarifToZpl
IF !logTrans
   RETURN 
ENDIF
*fSupl.Release
SELECT jobZpl
DO CASE
   CASE dimstru(1)=1
        SET FILTER TO 
   CASE dimstru(2)=1
        SET FILTER TO ','+LTRIM(STR(kp))+','$datagrup.sostav1                
ENDCASE
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .check1.Visible=.F.
     .Shape2.Visible=.T.
     .Shape3.Visible=.T.
     .lab24.Visible=.T.  
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape3.Width=1
ENDWITH
SELECT salpeop
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec
GO TOP
DO WHILE !EOF()   
   varStaj=''
   SELECT jobZpl
   SEEK VAL(SUBSTR(salpeop.tabfio,8,5))
   IF FOUND()
      *IF tr=1
      *    varStaj=staj_tar
      *ENDIF
      SELECT salpeop        
      SELECT tarifzp
      REPLACE sumrep WITH 0,sumrep1 WITH 0,logsb WITH .F. ALL
      SELECT jobZpl
      SCAN WHILE jobZpl.tabn=VAL(SUBSTR(salpeop.tabfio,8,5))
          * varStaj=''
           IF tr=1
              varStaj=staj_tar
           ENDIF
           SELECT tarifzp
           GO TOP
           DO WHILE !EOF()
              IF tr=jobZpl.tr
                 trfrep=ALLTRIM(trf)
                 trfrep1=ALLTRIM(trf1)
                 IF sumrep#0
                    REPLACE logsb WITH .T.
                 ENDIF
                 *REPLACE sumrep WITH sumrep+&trfrep
                 *REPLACE sumrep1 WITH sumrep1+&trfrep1
                 REPLACE sumrep WITH &trfrep
                 REPLACE sumrep1 WITH &trfrep1
              ENDIF
              SKIP
           ENDDO
           SELECT jobZpl
      ENDSCAN
      SELECT tarifzp
      GO TOP  

      DO WHILE !EOF() 
         repzpl=ALLTRIM(zpl) 
         IF logTotOkl
            IF sumrep#0
               SELECT salpeop  
               DO CASE
                  CASE dimbg(1)=1    &&бюджет
                       IF tarifZp.logSb
                          stroklad=PADL(ALLTRIM(STR(tarifzp.sumrep1,8,2)),8,' ')+SUBSTR(&repzpl,9)           
                       ELSE 
                          stroklad=PADL(ALLTRIM(STR(tarifzp.sumrep,8,2)),8,' ')+SUBSTR(&repzpl,9)                                                                                              
                       ENDIF 
                  CASE dimbg(2)=1    &&внебюджет  
                       IF tarifZp.logSb   
                          stroklad=LEFT(&repzpl,8)+PADL(ALLTRIM(STR(tarifzp.sumrep1,8,2)),8,' ')+SUBSTR(&repzpl,17)                                                                        
                       ELSE
                          stroklad=LEFT(&repzpl,8)+PADL(ALLTRIM(STR(tarifzp.sumrep,8,2)),8,' ')+SUBSTR(&repzpl,17)                               
                       ENDIF   
                  CASE dimbg(3)=1   &&внебюджет2 
                       IF tarifZp.logSb                                                                         
                          stroklad=LEFT(&repzpl,16)+PADL(ALLTRIM(STR(tarifzp.sumrep1,8,2)),8,' ')                            
                       ELSE
                          stroklad=LEFT(&repzpl,16)+PADL(ALLTRIM(STR(tarifzp.sumrep,8,2)),8,' ')                            
                       ENDIF   
               ENDCASE  
               REPLACE &repzpl WITH stroklad   
            ENDIF     
         ELSE 
            *repzpl=ALLTRIM(zpl)
            *reptrf=ALLTRIM(trf)
            IF sumrep#0
               SELECT salpeop  
               DO CASE
                  CASE dimbg(1)=1    &&бюджет
                       stroklad=PADL(ALLTRIM(STR(tarifzp.sumrep,8,2)),8,' ')+SUBSTR(&repzpl,9)                                                                                              
                  CASE dimbg(2)=1    &&внебюджет                                                                           
                       stroklad=LEFT(&repzpl,8)+PADL(ALLTRIM(STR(tarifzp.sumrep,8,2)),8,' ')+SUBSTR(&repzpl,17)                               
                  CASE dimbg(3)=1   &&внебюджет2 
                       stroklad=LEFT(&repzpl,16)+PADL(ALLTRIM(STR(tarifzp.sumrep,8,2)),8,' ')                            
               ENDCASE  
               REPLACE &repzpl WITH stroklad            
            ENDIF                           
         ENDIF  
         IF !EMPTY(varStaj).AND.logStaj
            SELECT salpeop  
            yst=VAL(LEFT(varStaj,2))
            mst=VAL(substr(varStaj,4,2))
            mst=mst/100
            repst=yst+mst
            REPLACE st_yemn WITH repst              
         ENDIF     
         SELECT tarifzp
         SKIP
      ENDDO      
   ENDIF    
   SELECT salpeop
   one_pers=one_pers+1
   pers_ch=one_pers/max_rec*100
   fSupl.shape3.Visible=.T.
   fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
   fSupl.Shape3.Width=fSupl.shape2.Width/100*pers_ch  
   SKIP
ENDDO       
SELECT salpeop
=SYS(2002)
=INKEY(2)
WITH fSupl          
     .shape2.Visible=.F.
     .shape3.Visible=.F.
     .lab25.Visible=.F.
     .lab24.Caption='Перенос выполнен'   
     .cont3.Visible=.T.
ENDWITH  
********************************************************************************************************************************************************
*                          Процедура присвоения табельных номеров из зарплаты в тарификацию
********************************************************************************************************************************************************
PROCEDURE tabzptotrf
logTab=.F.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .BackColor=RGB(255,255,255)
     .Caption='Присвоение табельных номеров'    
     .Width=700
     DO adLabMy WITH 'fSupl',1,'Внимание! Данная процедура позволяет присвоить списку сотрудников в ПО "Тарификация и штатное расписание"',10,10,700,2,.T.    
     .Width=.lab1.Width+20
     .lab1.Left=(.Width-.lab1.Width)/2
     DO adLabMy WITH 'fSupl',2,'табельные номера из списка сотрудников ПО "Сальдо"',.lab1.Top+.lab1.Height,.lab1.Left,.lab1.Width,2,.T.  
     DO adLabMy WITH 'fSupl',3,'100% точность присвоения не гарантируется! После выполнения сверьте списки.',.lab2.Top+.lab2.Height,.lab1.Left,.lab1.Width,2,.T. 
     
     DO adCheckBox WITH 'fSupl','check1','подтверждение выполнения',.lab3.Top+.lab3.Height+10,0,150,dHeight,'logTab',0
     .check1.Left=(.Width-.check1.Width)/2
     
     DO addcontlabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('WприсвоитьW')*2-20)/2,.check1.Top+.check1.Height+20,RetTxtWidth('WприсвоитьW'),dHeight+3,'присвоить','DO proctabtotarif'      
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отказ','fSupl.Release'                       
     
     DO addcontlabel WITH 'fSupl','cont3',(.Width-RetTxtWidth('WвозвратW'))/2,.cont1.Top,RetTxtWidth('WвозвратW'),dHeight+5,'возврат','fSupl.Release','возврат' 
     .cont3.Visible=.F.
     
      DO addShape WITH 'fSupl',2,5,.cont1.Top,.cont1.Height,.Width-10,8
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
     DO addShape WITH 'fSupl',3,.Shape2.Left,.Shape2.Top,.Shape2.Height,50,8
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.               
      DO adLabMy WITH 'fSupl',25,'100%',.Shape2.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab25.Visible=.F.
      DO adLabMy WITH 'fSupl',24,'Ход выполнения',.check1.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab24.Visible=.F.
    
        
     .Height=.lab1.Height*3+.check1.height+.cont1.Height+50
     .AutoCenter=.T.
ENDWITH
fSupl.Show
******************************************************************************************************************************************************
PROCEDURE proctabtotarif
IF !logTab
   RETURN 
ENDIF 
SELECT people
ordOld=SYS(21)
SET ORDER TO 1
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .check1.Visible=.F.
     .Shape2.Visible=.T.
     .Shape3.Visible=.T.
     .lab24.Visible=.T.  
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape3.Width=1
ENDWITH 
STORE 0 TO max_rec,one_pers,pers_ch
SELECT datjob
SET FILTER TO 
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,fio) ALL
COUNT TO max_rec
oldord=SYS(21)
SET ORDER TO 2

GO TOP
kvo_ch=0
DO WHILE !EOF()
   newtabn=0   
   SELECT salpeop
   IF SEEK(STR(datjob.tabn,5),'salpeop',1).AND.LOWER(ALLTRIM(datjob.fio))=LOWER(ALLTRIM(SUBSTR(tabfio,13,30))+' '+ALLTRIM(SUBSTR(tabfio,43,25))+' '+ALLTRIM(SUBSTR(tabfio,68,25)))     
   ELSE
      LOCATE FOR LOWER(ALLTRIM(datjob.fio))=LOWER(ALLTRIM(SUBSTR(tabfio,13,30))+' '+ALLTRIM(SUBSTR(tabfio,43,25))+' '+ALLTRIM(SUBSTR(tabfio,68,25)))
      IF FOUND()
         newtabn=ROUND(VAL(SUBSTR(salpeop.tabFio,8,5)),0)       
         kvo_ch=kvo_ch+1
      ENDIF
      SELECT datjob
      REPLACE tabn WITH newtabn
      SELECT people
      SEEK datjob.kodpeop
      REPLACE tabn WITH newtabn
      SELECT jobzpl
      REPLACE tabn WITH newtabn FOR kodpeop=datjob.kodpeop
      
   ENDIF
   
   SELECT datjob    
   SKIP
   one_pers=one_pers+1
   pers_ch=one_pers/max_rec*100
   fSupl.shape3.Visible=.T.
   fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
   fSupl.Shape3.Width=fSupl.shape2.Width/100*pers_ch  
ENDDO
SELECT jobzpl
=SYS(2002)
=INKEY(2)
WITH fSupl          
     .shape2.Visible=.F.
     .shape3.Visible=.F.
     .lab25.Visible=.F.
     .lab24.Caption='Присвоение выполнено'   
     .cont3.Visible=.T.
ENDWITH  
******************************************************************************************************************************************************
PROCEDURE procfSuplRefresh
fSupl.SetAll('Visible',.T.,'myContLabel')
fSupl.cont5.Visible=.F.
fSupl.lab6.Visible=.F.
******************************************************************************************************************************************************
PROCEDURE readTrfZpl
SELECT datJob
SET FILTER TO 
SELECT protError
WITH fTrans
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .SetAll('Enabled',.F.,'myOptionButton')
     .SetAll('Enabled',.F.,'myCheckBox')
     .SetAll('Enabled',.F.,'ComboMy')
     .butRet.Visible=.T.
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.F. 
     .fGrid.Column2.Enabled=.T. 
     .fGrid.Column3.Enabled=.T. 
ENDWITH

******************************************************************************************************************************************************
PROCEDURE exitFromReatTrfZpl
WITH fTrans
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Enabled',.T.,'myOptionButton')
     .SetAll('Enabled',.T.,'myCheckBox')
     .SetAll('Enabled',.T.,'ComboMy')
     .butRet.Visible=.F.     
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T. 
     
ENDWITH
******************************************************************************************************************************************************
PROCEDURE validTabJob
IF INLIST(protError.kodEr,1,2)
   SELECT datjob
   LOCATE FOR kodPeop=protError.kodPeop.AND.kp=protError.kp.AND.kd=protError.kd.AND.tr=protError.tr.AND.kse=proterror.kse
   REPLACE tabn WITH protError.nTab
   SELECT protError
ENDIF
******************************************************************************************************************************************************
PROCEDURE validTrJob
IF INLIST(protError.kodEr,1,2)
   SELECT datjob
   LOCATE FOR kodPeop=protError.kodPeop.AND.kp=protError.kp.AND.kd=protError.kd.AND.kse=proterror.kse
   REPLACE tr WITH protError.tr
   SELECT protError
ENDIF