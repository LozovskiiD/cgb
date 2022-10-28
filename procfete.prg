* Структура полей fete1-fete12
* 1-2   - кол-во праздничных дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата за час (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-34 - норма времени (оперативное поле - opnorm)
* 35-36 - кол-во должностей2 (оперативное поле - oppost2)
* 37-41 - кол-во часов работы в день2 (оперативное поле - ophday2)
* 42-43 - кол-во дней 2 (оперативное поле - odnight2)
SELECT rasp
SET FILTER TO 
SELECT people
SET FILTER TO 
SELECT sprpodr 
SET FILTER TO 
oldOrdPodr=SYS(21)
SET ORDER TO 2
SELECT datJob
SET FILTER TO 
SET ORDER TO 2
IF USED('curSprTime')
   SELECT curSprTime
   USE
ENDIF 
IF USED('curFetekat')
   SELECT curFeteKat
   USE
ENDIF
IF USED('curTarJob')
   SELECT curTarJob
   USE 
ENDIF
SELECT * FROM datJob INTO CURSOR curTarJob READWRITE
SELECT curTarJob
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SET ORDER TO 1

dCount=varDtar
RESTORE FROM logFt ADDITIVE
DIMENSION dim_fete(12)
STORE 0 TO dim_fete

SELECT * FROM sprkat INTO CURSOR curFeteKat READWRITE
ALTER TABLE curFeteKat ADD COLUMN sumTot N(9,2)
SELECT curFeteKat
INDEX ON kod TAG T1
SELECT sprtime
=AFIELDS(arSprTime,'sprtime')
CREATE CURSOR cursprtime FROM ARRAY arSprTime
SELECT cursprtime
APPEND FROM sprtime
INDEX ON name TAG T1
IF !USED('fete')
   USE fete ORDER 1 IN 0 
ENDIF
SELECT fete
DIMENSION dim_fete(12),dim_fetestr(12)
STORE 0 TO dim_fete,kvoDayFete
STORE '' TO dim_fetestr,strtime
GO TOP 
DO WHILE !EOF()     
   dim_fete(ROUND((datafet-INT(datafet))*100,0))=dim_fete(ROUND((datafet-INT(datafet))*100,0))+1
   dim_fetestr(ROUND((datafet-INT(datafet))*100,0))=dim_fetestr(ROUND((datafet-INT(datafet))*100,0))+'   '+LTRIM(STR(datafet))+' - '+ALLTRIM(comment)
   SKIP 
ENDDO 
SELECT rasp
SET RELATION TO kd INTO sprdolj,ntime INTO sprtime ADDITIVE
curnamepodr=''
kpdop=0
kpdop=IIF(rasp->kp=0,1,rasp->kp)
log_read=.F.
log_count=.F.
month_ch=1
repFete=''
*---Добавляем вакансии
SELECT datjob
ordOld=SYS(21)
SET FILTER TO 
SELECT rasp
GO TOP
DO WHILE !EOF()
   IF rasp.kse#0
      SELECT datjob
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
         SELECT curTarJob
         APPEND BLANK
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nKfVac,kse WITH kse_cx,vac WITH .T.
         REPLACE tokl WITH varBaseSt*curTarJob.namekf,mtokl WITH tokl*kse                                          
         SELECT curTarJob  
      ENDIF                                 
   ENDIF
   SELECT rasp
   SKIP
ENDDO
SELECT datJob
SET ORDER TO &ordOld  

SELECT rasp
SET FILTER TO kp=kpdop
DO repstrfete WITH 'rasp',.T.,'month_ch'
SELECT cursprpodr
LOCATE FOR kod=kpdop
IF FOUND()
   curnamepodr=ALLTRIM(cursprpodr.name)  
ELSE
   GO TOP
   curnamepodr=cursprpodr.name 
ENDIF 
fForm=Createobject('FORMSPR')
WITH fForm
     .Caption='Расчёт расходов на дополнительную оплату за работу в государственные праздники и праздничные дни'
     .procExit='DO exitFromProcFete'     
     DO addButtonOne WITH 'fForm','menuCont1',10,5,'редакция','pencil.ico','DO readFete',39,RetTxtWidth('загрузить из')+44,'редакция'  
     DO addButtonOne WITH 'fForm','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'расчёт','calculate.ico','DO procCountFete',39,.menucont1.Width,'расчёт'   
     DO addButtonOne WITH 'fForm','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delFete',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fForm','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico','DO prnFeteForm',39,.menucont1.Width,'печать'   
     DO addButtonOne WITH 'fForm','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'в Excel','excel.ico','DO excelFeteForm',39,.menucont1.Width,'в Excel'  
     DO addButtonOne WITH 'fForm','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'календарь','ical.ico','DO proccalendar',39,.menucont1.Width,'календарь'
     DO addButtonOne WITH 'fForm','menuCont7',.menucont6.Left+.menucont6.Width+3,5,'загрузить из','get.ico','DO formdownloadfete',39,.menucont1.Width,'загрузить из предыдущего периода'   
     DO addButtonOne WITH 'fForm','menuCont8',.menucont7.Left+.menucont7.Width+3,5,'возврат','undo.ico','DO exitFromProcFete',39,.menucont1.Width,'возврат'       
               
     DO addComboMy WITH 'fform',1,(fForm.Width-715)/2,fForm.menucont1.Top+fForm.menucont1.Height+15,dHeight,500,.T.,'curnamepodr','sprpodr.name',6,.F.,;
        'DO validPodFete',.F.,.T.  
     .comboBox1.DisplayCount=20         
     DO addComboMy WITH 'fform',2,fForm.comboBox1.Left+fForm.comboBox1.Width+15,fForm.comboBox1.Top,dHeight,200,.T.,'month_ch','dim_month',5,.F.,;
        'DO validMonthFete',.F.,.T.   
     .comboBox2.DisplayCount=12                         
     WITH .fGrid    
          .Top=fForm.ComboBox1.Top+fForm.ComboBox1.Height+5          
          .Height=fForm.Height-.Top-dHeight
          .Width=fForm.Width
          .RecordSource='rasp'
          DO addColumnToGrid WITH 'fForm.fGrid',14
          .RecordSourceType=1
          .ColumnCount=14            
          .Column1.ControlSource='" "+sprdolj->name'
          .Column2.ControlSource='rasp.opdnight'
          .Column3.ControlSource='ROUND(rasp.ophday,2)'
          .Column4.ControlSource='rasp.oppost'             
          .Column5.ControlSource='rasp.opdnight2'                  
          .Column6.ControlSource='ROUND(rasp.ophday2,2)'
          .Column7.ControlSource='rasp.oppost2' 
          .Column8.ControlSource='rasp.optoth'                                                           
          .Column9.ControlSource='rasp.srtofete'          
          .Column10.ControlSource='sprtime.name'          
          .Column11.ControlSource='rasp.opnorm'
          .Column12.ControlSource='rasp.opzph'                    
          .Column13.ControlSource='rasp.opsumtot'
                   
          .Column2.Width=RettxtWidth('дней2 ')
          .Column3.Width=RettxtWidth('999999')
          .Column4.Width=RetTxtWidth('сот.')
          .Column5.Width=.Column2.Width          
          .Column6.Width=.Column3.Width
          .Column7.Width=.Column4.Width 
          .Column8.Width=RetTxtWidth('9999999')                            
          .Column9.Width=RetTxtWidth('999999999')
          .Column10.Width=RetTxtWidth('wсредний медпww')
          .Column11.Width=RetTxtWidth('9999999')
          .Column12.Width=RetTxtWidth(' час.окл. ')          
          .Column13.Width=RetTxtWidth('999999999')
                    
          .Columns(.ColumnCount).Width=0   
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-;
                         .Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-.Column12.Width-.Column13.Width-SYSMETRIC(5)-13-.ColumnCount          
          .Column1.Header1.Caption='должность'
          .Column2.Header1.Caption='дней'
          .Column3.Header1.Caption='ч-д'
          .Column4.Header1.Caption='ст.'
          .Column5.Header1.Caption='дней2'
          .Column6.Header1.Caption='ч-д2'
          .Column7.Header1.Caption='ст2'        
          .Column8.Header1.Caption='ч-м'  
          .Column9.Header1.Caption='ср.окл'
          .Column10.Header1.Caption='время'
          .Column11.Header1.Caption='норма'
          .Column12.Header1.Caption='час.окл.'                                          
          .Column13.Header1.Caption='сумма.м.' 
          .SetAll('Alignment',1,'ColumnMy')            
          .Column1.Alignment=0         
          .Column10.Alignment=0
          .Column14.Alignment=0
          .Column2.Format='Z'
          .Column3.Format='Z'
          .Column4.Format='Z'
          .Column5.Format='Z'                    
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'
          .Column9.Format='Z'
          .Column11.Format='Z'
          .Column12.Format='Z'
          .Column13.Format='Z'                                  
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')           
          .colNesInf=2              
     ENDWITH
     DO gridSizeNew WITH 'fForm','fGrid','shapeingrid'  
     DO adLabMy WITH 'fForm',1,dim_fetestr(month_ch),fForm.Height-dHeight,0,fForm.Width,2,.F.,0   
     .combobox1.DisplayCount=MIN(RECCOUNT('sprpodr'),(fForm.Height-fForm.combobox1.Top-fForm.combobox1.Height)/fForm.fGrid.Rowheight)      
ENDWITH
SELECT rasp
GO TOP
fForm.Show
********************************************************************************************************************************************************
PROCEDURE exitFromProcFete
SELECT sprpodr 
SET ORDER TO &oldOrdPodr
SELECT rasp
SET FILTER TO 
SET RELATION TO
SELECT datJob
SET FILTER TO 
fForm.Release
*******************************************************************************************************************************************************
*                             Процедура замены данных оперативных полей
*******************************************************************************************************************************************************
PROCEDURE repstrfete
PARAMETERS parDbf,parScope,parMonth
*parDbf - rasp или curNight
*parScope - текущая запись или все
*parMonth - переменная периода
* 1-2   - кол-во праздничных дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата за час (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-34 - норма времени (оперативное поле - opnorm)
* 35-36 - кол-во должностей2 (оперативное поле - oppost2)
* 37-41 - кол-во часов работы в день2 (оперативное поле - ophday2)
* 42-43 - кол-во дней 2 (оперативное поле - odnight2)

SELECT &parDbf
repFete='fete'+LTRIM(STR(&parMonth))
kvoDayFete=dim_fete(&parMonth)
*kvoDayFete=IIF(logFt,dim_fete(&parMonth),VAL(SUBSTR(&repFete,1,2)))
IF parScope    
   REPLACE opdnight WITH IIF(logFt,kvoDayFete,VAL(SUBSTR(&repFete,1,2)));
           ophday WITH VAL(SUBSTR(&repFete,3,5)),;
           optoth WITH VAL(SUBSTR(&repFete,8,3)),;
           oppost WITH VAL(SUBSTR(&repFete,11,2)),;         
           opzph WITH VAL(SUBSTR(&repFete,13,8)),;
           opsumtot WITH VAL(SUBSTR(&repFete,21,8)),;                     
           opnorm WITH VAL(SUBSTR(&repFete,29,6)),; 
           oppost2 WITH VAL(SUBSTR(&repFete,35,2)),; 
           ophday2 WITH  VAL(SUBSTR(&repFete,37,5)),;
           opdnight2 WITH VAL(SUBSTR(&repFete,42,2)) ALL           
ELSE   
   REPLACE opdnight WITH IIF(logFt,kvoDayFete,VAL(SUBSTR(&repFete,1,2))),;
           ophday WITH VAL(SUBSTR(&repFete,3,5)),;
           optoth WITH VAL(SUBSTR(&repFete,8,3)),;
           oppost WITH VAL(SUBSTR(&repFete,11,2)),;         
           opzph WITH VAL(SUBSTR(&repFete,13,8)),;
           opsumtot WITH VAL(SUBSTR(&repFete,21,8)),;                    
           opnorm WITH VAL(SUBSTR(&repFete,29,6)),;
           oppost2 WITH VAL(SUBSTR(&repFete,35,2)),; 
           ophday2 WITH  VAL(SUBSTR(&repFete,37,5)),;
           opdnight2 WITH VAL(SUBSTR(&repFete,42,2))          
ENDIF         
*************************************************************************************************************************
PROCEDURE procvalidtime
SELECT rasp
REPLACE ntime WITH cursprtime.kod,opnorm WITH EVALUATE('cursprtime.t'+LTRIM(STR(month_ch)))
strtime=cursprtime.name
*KEYBOARD '{TAB}'    
SELECT rasp
fForm.Refresh
************************************************************************************************************************
PROCEDURE procgottime
SELECT cursprtime
LOCATE FOR kod=sprtime->kod
strtime=sprtime.name
fForm.fGrid.Column10.cbotime.ControlSource='strtime' 
nrec=RECNO()
GO TOP 
COUNT WHILE RECNO()#nrec TO varnrec
fForm.fGrid.column10.cbotime.DisplayCount=MAX(fForm.fGrid.RelativeRow,fForm.fGrid.RowsGrid-fForm.fGrid.RelativeRow)
fForm.fGrid.column10.cbotime.DisplayCount=MIN(fForm.fGrid.column10.cbotime.DisplayCount,RECCOUNT())
fForm.fGrid.Column10.cbotime.varCtrlSource=varnrec+1 
SELECT rasp
********************************************************************************************************************************************************
PROCEDURE validPodFete
SELECT sprpodr
kpdop=sprpodr.kod
curnamepodr=fForm.ComboBox1.Value
SELECT rasp
SET FILTER TO kp=kpdop
DO repstrfete WITH 'rasp',.T.,'month_ch'
*fForm.fGrid.Column8.ControlSource="EVALUATE('sprtime.t'+LTRIM(STR(month_ch)))"
GO TOP
fForm.Refresh
********************************************************************************************************************************************************
PROCEDURE validMonthFete
month_ch=fForm.comboBox2.Value
DO repstrfete WITH 'rasp',.T.,'month_ch'
fForm.lab1.Caption=dim_fetestr(month_ch)
SELECT rasp 
GO TOP
fForm.Refresh
*******************************************************************************************************************************************************
*                             редактиоованиек праздничных (новый вариант)
*******************************************************************************************************************************************************
PROCEDURE readFete
fSupl=CREATEOBJECT('FORMSUPL')
SELECT rasp
raspRec=RECNO()
raspKp=kp
SELECT curTarJob
STORE 0 TO kvo_peop,sumto 
SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
     kvo_peop=kvo_peop+kse
     sumto=sumto+mtokl
ENDSCAN
IF rasp.dtf#0
   SEEK STR(rasp.kp,3)+STR(rasp.dtf,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.dtf
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
ENDIF
SELECT rasp       
newSrOkl=ROUND(IIF(kvo_peop#0,sumto/kvo_peop,0),2)

newTime=ntime     && вид времени
newNorm=opnorm    && норма времени
*newSrOkl=srTofete && средний оклад

day1=opdnight     && дней 1
hour1=ophday      && часов в день 1
sotr1=oppost      && сотрудников 1
day2=opdnight2    && дней 2
hour2=ophday2     && часов в день 2
sotr2=oppost2     && сотрудников 2
totsumnew=opsumtot
newDtf=dtf
logDtf=IIF(dtf>0,.T.,.F.)

saypodr=IIF(SEEK(kpdop,'sprpodr',1),ALLTRIM(sprpodr.name),'')
saydol=IIF(SEEK(rasp.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')
sayDolSupl=IIF(SEEK(rasp.dtf,'sprdolj',1),ALLTRIM(sprdolj.name),'')
SELECT * FROM sprdolj WHERE SEEK(STR(raspKp,3)+STR(kod,3),'rasp',2) INTO CURSOR curSuplDol READWRITE
SELECT rasp
GO raspRec

SELECT curSprTime
LOCATE FOR kod=rasp.ntime
strtime=curSprTime.name

WITH fSupl
     .Caption='Ввод-редакция затрат на оплату работы в праздничные дни'
     .procExit='DO writeFete WITH .F.'
     DO adLabMy WITH 'fSupl',1,saypodr,10,0,fSupl.Width,2,.F.,0
     DO adLabMy WITH 'fSupl',2,saydol,.lab1.Top+.lab1.Height,2,fSupl.Width,2,.F.,0  
     DO addShape WITH 'fSupl',1,10,.lab2.Top+.lab2.Height,dHeight,300,8                    
       
       
     DO adTBoxAsCont WITH 'fsupl','txtTime',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('wПраздничных днейw'),dHeight,'время',0,1
     DO addComboMy WITH 'fSupl',1,.txtTime.Left+.txtTime.Width-1,.txtTime.Top,dHeight,250,.T.,'strTime','ALLTRIM(curSprTime.name)',6,.F.,'DO validTime',.F.,.T.
       
     DO adTBoxAsCont WITH 'fsupl','txtNorm',.txtTime.Left,.txtTime.Top+.txtTime.Height-1,.txtTime.Width,dHeight,'норма',0,1           
     DO adTboxNew WITH 'fSupl','boxNorm',.txtNorm.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNorm',.F.,.F.
     
     DO adTBoxAsCont WITH 'fsupl','txtOklad',.txtTime.Left,.txtNorm.Top+.txtNorm.Height-1,.txtTime.Width,dHeight,'средний оклад',0,1     
     DO adTboxNew WITH 'fSupl','boxOklad',.txtOklad.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSrOkl',.F.,.F.,0
          
     DO adTBoxAsCont WITH 'fsupl','txtDay1',.txtTime.Left,.txtOklad.Top+.txtOklad.Height-1,.txtTime.Width,dHeight,'праздничных дней 1',0,1
     DO adTboxNew WITH 'fSupl','boxDay1',.txtDay1.Top,.txtDay1.Left+.txtDay1.Width-1,.comboBox1.Width,dHeight,'day1',.F.,.T.
        
     DO adTBoxAsCont WITH 'fsupl','txtHour1',.txtDay1.Left,.txtDay1.Top+.txtDay1.Height-1,.txtDay1.Width,dHeight,'часов в день 1',0,1
     DO adTboxNew WITH 'fSupl','boxHour1',.txtHour1.Top,.boxDay1.Left,.boxDay1.Width,dHeight,'hour1',.F.,.T.,0
     
     DO adTBoxAsCont WITH 'fsupl','txtSotr1',.txtDay1.Left,.txtHour1.Top+.txtHour1.Height-1,.txtDay1.Width,dHeight,'сотрудников1',0,1     
     DO adTboxNew WITH 'fSupl','boxSot1',.txtSotr1.Top,.boxDay1.Left,.boxDay1.Width,dHeight,'sotr1',.F.,.T.,0
     
     DO adTBoxAsCont WITH 'fsupl','txtDay2',.txtDay1.Left,.txtSotr1.Top+.txtSotr1.Height-1,.txtDay1.Width,dHeight,'праздничных дней 2',0,1
     DO adTboxNew WITH 'fSupl','boxDay2',.txtDay2.Top,.boxDay1.Left,.boxDay1.Width,dHeight,'day2',.F.,.T.,0
          
     DO adTBoxAsCont WITH 'fsupl','txtHour2',.txtDay2.Left,.txtDay2.Top+.txtDay2.Height-1,.txtDay1.Width,dHeight,'часов в день 2',0,1
     DO adTboxNew WITH 'fSupl','boxHour2',.txtHour2.Top,.boxDay1.Left,.boxDay1.Width,dHeight,'hour2',.F.,.T.,0
     
     DO adTBoxAsCont WITH 'fsupl','txtSotr2',.txtDay1.Left,.txtHour2.Top+.txtHour2.Height-1,.txtDay1.Width,dHeight,'сотрудников 2',0,1     
     DO adTboxNew WITH 'fSupl','boxSot2',.txtSotr2.Top,.boxDay1.Left,.boxDay1.Width,dHeight,'sotr2',.F.,.T.,0                                            
                                                     
     .Shape1.Height=.txtDay1.Height*9+20-9
     .Shape1.Width=.txtDay1.Width+.boxDay1.Width+20                                         

     DO addShape WITH 'fSupl',2,10,.Shape1.Top+.Shape1.Height+10,100,.Shape1.Width,8                                           
     DO adCheckBox WITH 'fSupl','check1','объединить с ',.Shape2.Top+10,.Shape2.Left,.Shape2.Width,dHeight,'logdtf',0,.T.,'DO validCheckSupl'
     DO addComboMy WITH 'fSupl',2,.Shape2.Left+10,.check1.Top+.check1.Height+10,dHeight,.Shape2.Width-20,IIF(newDtf>0,.T.,.F.),'sayDolSupl','ALLTRIM(curSuplDol.name)',6,.F.,'DO validSuplDol',.F.,.T.
     .check1.Left=.Shape2.Left+(.Shape2.Width-.check1.Width)/2
     .Shape2.Height=.check1.Height+.comboBox2.Height+30

     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,;
      .Shape2.Top+.Shape2.Height+20,RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать','DO writeFete WITH .T.'    
    
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Возврат','DO writeFete WITH .F.'
     .Width=.txtDay1.Width+.boxDay1.Width+40
     .Height=.lab1.Height*2+.Shape1.Height+.Shape2.Height+.cont1.Height+50
     .lab1.Width=.Width
     .lab2.Width=.Width
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validCheckSupl
fSupl.comboBox2.Enabled=IIF(logDtf,.T.,.F.)
IF !logDtf
   newDtf=0
   STORE 0 TO kvo_peop,sumto 
   SELECT curTarJob
   STORE 0 TO kvo_peop,sumto 
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
   SELECT rasp 
   newSrOkl=ROUND(IIF(kvo_peop#0,sumto/kvo_peop,0),2)  
   sayDolSupl=''
   fSupl.comboBox2.ControlSource='sayDolSupl'
   fSupl.Refresh
   
ENDIF 
*************************************************************************************************************************
PROCEDURE validSuplDol
newDtf=curSuplDol.kod
IF newDtf#0
   SELECT curTarJob
   STORE 0 TO kvo_peop,sumto 
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
   SEEK STR(rasp.kp,3)+STR(newDtf,3)
   SCAN WHILE kp=rasp.kp.AND.kd=newDtf
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
   SELECT rasp  
   newSrOkl=ROUND(IIF(kvo_peop#0,sumto/kvo_peop,0),2)  
   fSupl.Refresh
ENDIF
*************************************************************************************************************************
PROCEDURE validtime
newTime=curSprtime.kod
newnorm=EVALUATE('cursprtime.t'+LTRIM(STR(month_ch)))
strtime=cursprtime.name
KEYBOARD '{TAB}'  
fSupl.Refresh

*******************************************************************************************************************************************************
*            Запись праздничных
*******************************************************************************************************************************************************
PROCEDURE writeFete
PARAMETERS par1
SELECT rasp
oldrec=RECNO()
IF par1   
   REPLACE opdnight WITH day1,opdnight2 WITH day2,ophday WITH hour1,ophday2 WITH hour2,oppost WITH sotr1,oppost2 WITH sotr2,;
           opnorm WITH newNorm,srToFete WITH newSrOkl,opzph WITH IIF(opnorm#0,srtofete/opnorm*IIF(month_ch=1,1,1),0),nTime WITH newTime,;
           optoth WITH (opdnight*ophday)+(opdnight2*ophday2),opsumtot WITH (opdnight*ophday*opzph*oppost)+(opdnight2*ophday2*opzph*oppost2),dtf WITH newDtf   
   IF optoth=0
      REPLACE srtofete WITH 0,opzph WITH 0
   ENDIF   
   repFete='fete'+LTRIM(STR(month_ch))    
   repStrFete=STR(opdnight,2)+STR(ophday,5,2)+STR(optoth,3)+STR(oppost,2)+STR(opzph,8,2)+STR(opsumtot,8,2)+;
              STR(opnorm,6,2)+STR(oppost2,2)+STR(ophday2,5,2)+STR(opdnight2,2)
   REPLACE &repFete WITH repStrFete
   oldrec=RECNO()
   fForm.Refresh                                
   DO repstrfete WITH 'rasp',.F.,'month_ch'    
ENDIF    
fSupl.Release
GO oldrec
fForm.Refresh
********************************************************************************************************************************************************
PROCEDURE delFete
fdel=CREATEOBJECT('FORMMY')
log_del=.F.
DIMENSION dim_del(5)
STORE 0 TO dim_del
dim_del(1)=1
dim_del(4)=1
WITH fDel
     .Caption='Удаление'
     .BackColor=RGB(255,255,255)   
     DO addShape WITH 'fDel',1,10,10,100,100,8  
     DO addShape WITH 'fDel',2,10,10,100,100,8      
     DO addOptionButton WITH 'fdel',1,'очистить выбранную запись',fdel.Shape1.Top+10,fdel.Shape1.Left+15,'dim_del(1)',0,"DO otmdimdel WITH 1",.T.
     DO addOptionButton WITH 'fdel',2,'удалить подразделение',fdel.Option1.Top+fdel.Option1.Height+10,fdel.Option1.Left,'dim_del(2)',0,"DO otmdimdel WITH 2",.T.
     DO addOptionButton WITH 'fdel',3,'удалить все',fdel.Option2.Top+fdel.Option2.Height+10,fdel.Option1.Left,'dim_del(3)',0,"DO otmdimdel WITH 3",.T.
     .Shape1.Height=.Option1.height*3+40
     .Shape1.Width=.Option1.Width+30    
     .Shape2.Top=.Shape1.Top+.Shape1.Height+15
     .Shape2.Width=.Shape1.Width       
     DO addOptionButton WITH 'fdel',4,'удаление за текущий период',fdel.Shape2.Top+10,fdel.Option1.Left,'dim_del(4)',0,"DO otmdimdel WITH 4,.T.",.T.
     DO addOptionButton WITH 'fdel',5,'удаление до конца года',fdel.Option4.Top+fDel.Option4.Height+10,fdel.Option1.Left,'dim_del(5)',0,"DO otmdimdel WITH 5,.T.",.T.
     .Shape2.Height=.Option2.Height*2+30
     .Shape2.Width=.Shape1.Width         
             
     DO adCheckBox WITH 'fdel','check1','подтверждение удаления',fdel.Shape2.Top+fdel.Shape2.Height+10,fdel.Shape1.Left,150,dHeight,'log_del',0
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+5,fdel.check1.Top+fdel.check1.Height+15,;
       (fdel.shape1.Width-20)/2,dHeight+3,'Выполнение','DO delrecRashFete'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+10,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .check1.Left=(.Width-.check1.Width)/2
     .Height=.Shape1.Height+.Shape2.Height+fDel.cont1.Height+fdel.check1.Height+65
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
***********************************************************************************************************************************************
PROCEDURE otmdimdel
PARAMETERS par1,par2
DO CASE
   CASE !par2
        FOR i=1 TO 3
            dim_del(i)=IIF(i=par1,1,0)
        ENDFOR
   CASE par2
        FOR i=4 TO 5
            dim_del(i)=IIF(i=par1,1,0)
        ENDFOR
ENDCASE      
fdel.Refresh
*****************************************************************************************************************************************************
*                         Непосредственно удаление информации по праздничным
*****************************************************************************************************************************************************
PROCEDURE delrecRashFete
IF !log_del 
   RETURN 
ENDIF
SELECT rasp
DO CASE
   CASE dim_del(1)=1    
        SELECT rasp          
        DO delRecFeteOne     
   CASE dim_del(2)=1        
        SELECT rasp        
        GO TOP
        DO WHILE !EOF()          
           DO delRecFeteOne
           SKIP
        ENDDO                             
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()           
           DO delRecFeteOne
           SKIP
        ENDDO                                 
ENDCASE
fDel.Release
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
fForm.Refresh
*****************************************************************************************************************************************************
PROCEDURE delRecFeteOne
DO CASE
   CASE dim_del(4)=1 
        REPLACE srtofete WITH 0,dtf WITH 0  
        rep_ch='fete'+LTRIM(STR(month_ch))         
        REPLACE &rep_ch WITH ' ' 
        DO repstrfete WITH 'rasp',.F.,'month_ch'   
   CASE dim_del(5)=1
        REPLACE srtofete WITH 0,dtf WITH 0 
        FOR i=1 TO 12      
            rep_ch='fete'+LTRIM(STR(i))
            REPLACE &rep_ch WITH ''                     
        ENDFOR
ENDCASE    
********************************************************************************************************************************************************
PROCEDURE procCountFete
fdel=CREATEOBJECT('FORMMY')
log_del=.F.
DIMENSION dim_del(5)
STORE 0 TO dim_del
dim_del(1)=1
dim_del(2)=0
dim_del(3)=0
dim_del(4)=0
dim_del(5)=1
log_srzp=.F.
WITH fdel
     .Caption='Расчёт расходов на дополнительную оплату за работу в государственные праздники и праздничные дни'
     .BackColor=RGB(255,255,255)   
     DO addShape WITH 'fdel',1,10,10,dHeight,50,8     
     DO adLabMy WITH 'fdel',1,'Дата отсчёта',fdel.Shape1.Top+10,fdel.Shape1.Left+15,150,0,.T.
     DO adtbox WITH 'fdel',1,fdel.lab1.Left+fdel.lab1.Width+10,fdel.Shape1.Top+10,RetTxtWidth('99999999999'),dHeight,'dCount','Z',.T.,1
     fdel.lab1.Top=fdel.txtbox1.Top+(fdel.txtbox1.Height-fdel.lab1.Height)
     DO addOptionButton WITH 'fdel',1,'расчет по выбранной должности',fdel.txtbox1.Top+fdel.txtbox1.Height+10,fdel.Shape1.Left+15,'dim_del(1)',0,"DO otmdimdel WITH 1",.T.
     DO addOptionButton WITH 'fdel',2,'расчёт по подразделению',fdel.Option1.Top+fdel.Option1.Height+10,fdel.Option1.Left,'dim_del(2)',0,"DO otmdimdel WITH 2",.T.
     DO addOptionButton WITH 'fdel',3,'расчёт по организации',fdel.Option2.Top+fdel.Option2.Height+10,fdel.Option1.Left,'dim_del(3)',0,"DO otmdimdel WITH 3",.T.
     .Shape1.Height=.Option1.height*4+60    
     .Shape1.Width=.Option1.Width+30
     
     DO addShape WITH 'fdel',4,10,fdel.Shape1.Top+fdel.Shape1.Height+10,dHeight,fdel.Shape1.Width,8    
     DO addOptionButton WITH 'fdel',4,'расчёт за текущий период',fdel.Shape4.Top+10,fdel.Option1.Left,'dim_del(4)',0,"DO otmdimdel WITH 4,.T.",.T.
     DO addOptionButton WITH 'fdel',5,'расчёт до конца года',fdel.Option4.Top+fDel.Option4.Height+10,fdel.Option1.Left,'dim_del(5)',0,"DO otmdimdel WITH 5,.T.",.T.
     .Shape4.Height=.Option4.Height*2+30
     .Shape4.Width=.Shape1.Width
     
     DO adCheckBox WITH 'fdel','check3','расставлять праздничные дни',fdel.Shape4.Top+fdel.Shape4.Height+10,fdel.Shape1.Left,150,dHeight,'logFt',0,.F.,'DO validLogFt'     
     DO adCheckBox WITH 'fdel','check2','подтверждение выполнения',.check3.Top+.check3.Height+10,.Shape1.Left,150,dHeight,'log_del',0         
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('WВыполнениеW')*2-20)/2,fdel.check2.Top+fdel.check2.Height+15,;
       RetTxtWidth('WВыполнениеW'),dHeight+3,'Выполнение','DO countFete'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+20,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'         
        
      DO adLabMy WITH 'fdel',4,'Ход выполнения',fdel.check2.Top+fdel.check2.Height+5,fdel.Shape1.Left,fdel.Shape1.Width,2,.F.
     .lab4.Visible=.F.        
     DO addShape WITH 'fdel',2,fdel.Shape1.Left,fdel.lab4.Top+fdel.lab4.Height+5,dHeight,fdel.Shape1.Width
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
          
     DO addShape WITH 'fdel',3,fdel.Shape2.Left,fdel.Shape2.Top,dHeight,0
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.             
        
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape4.Height+fdel.cont1.Height+fdel.check2.Height*2+80
     .lab1.left=.Shape1.Left+(.shape1.Width-.lab1.Width-.txtbox1.Width-10)/2
     .txtbox1.Left=.lab1.Left+.lab1.Width+10
     .check2.Left=(.Width-.check2.Width)/2
     .check3.Left=(.Width-.check3.Width)/2
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
***************************************************************************************************************************************************
PROCEDURE validLogFt
pathLogFt=FULLPATH('logFt.mem')
SAVE TO &pathLogFt ALL LIKE logFt
****************************************************************************************************************************************************
*                           Непосредственно процедра общего расчёта расходов по Указу
****************************************************************************************************************************************************
PROCEDURE countFete
IF !log_del
   fDel.Release
   RETURN
ENDIF
log_count=.T.
SELECT rasp
recpeop=RECNO()
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec 
IF max_rec=0
   RETURN
ENDIF
GO recpeop
fdel.cont1.Visible=.F.
fdel.cont2.Visible=.F.
fdel.lab4.Visible=.T. 
fdel.Shape2.Visible=.T.
fdel.Shape3.Visible=.T.
=SYS(2002)
DO CASE
   CASE dim_del(1)=1
        max_rec=1 
        GO recpeop        
        DO countOneRecFete      
        one_pers=one_pers+1
        pers_ch=one_pers/max_rec*100
        fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch 
   CASE dim_del(2)=1
        SELECT rasp
        GO TOP        
        DO WHILE !EOF()
           DO countOneRecFete
           SELECT rasp
           SKIP 
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch               
        ENDDO
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO
        COUNT TO max_rec 
        GO TOP        
        DO WHILE !EOF()
           DO countOneRecFete
           SELECT rasp
           SKIP
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch  
        ENDDO
ENDCASE
SELECT rasp
SET FILTER TO kp=kpdop
DO repstrfete WITH 'rasp',.T.,'month_ch' 
GO recpeop
=INKEY(1)
fDel.Visible=.F.
fdel.Release
DO createFormNew WITH .T.,'Общий расчёт',RetTxtWidth('WWРасчёт выполнен!WW',dFontName,dFontSize+1),'130',;
      RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
      'Расчёт выполнен!',.F.,.T. 
=SYS(2002,1) 
log_count=.F.
fForm.Refresh 
****************************************************************************************************************************************************
*                                 Расчет расходов по одной записи
****************************************************************************************************************************************************
PROCEDURE countOneRecFete
* 1-2   - кол-во праздничных дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата за час (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-34 - норма времени (оперативное поле - opnorm)
* 35-36 - кол-во должностей2 (оперативное поле - oppost2)
* 37-41 - кол-во часов работы в день2 (оперативное поле - ophday2)
* 42-43 - кол-во дней 2 (оперативное поле - odnight2)
SELECT curTarJob
STORE 0 TO kvo_peop,sumto 
SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
     kvo_peop=kvo_peop+kse
     sumto=sumto+mtokl
ENDSCAN
IF rasp.dtf#0
   SEEK STR(rasp.kp,3)+STR(rasp.dtf,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.dtf
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
ENDIF
SELECT rasp
REPLACE srtofete WITH IIF(kvo_peop#0,sumto/kvo_peop,0)
DO CASE
   CASE dim_del(4)=1        
        DO repstrfete WITH 'rasp',.F.,'month_ch'                    
        REPLACE opnorm WITH EVALUATE('sprtime.t'+LTRIM(STR(month_ch))),opzph WITH IIF(opnorm#0,srtofete/opnorm*IIF(month_ch=1,1,1),0),;
                optoth WITH (opdnight*ophday)+(opdnight2*ophday2),opsumtot WITH (opdnight*ophday*opzph*oppost)+(opdnight2*ophday2*opzph*oppost2)
        repStrFete=STR(opdnight,2)+STR(ophday,5,2)+STR(optoth,3)+STR(oppost,2)+STR(opzph,8,2)+STR(opsumtot,8,2)+STR(opnorm,6,2)+STR(oppost2,2)+STR(ophday2,5,2)+STR(opdnight2,2)   
        REPLACE &repFete WITH repStrFete  
   CASE dim_del(5)=1              
        FOR i=1 TO 12           
            DO repstrfete WITH 'rasp',.F.,'i'                                
            IF i>=MONTH(dCount)
               *REPLACE opnorm WITH EVALUATE('sprtime.t'+LTRIM(STR(month_ch))),opzph WITH IIF(opnorm#0,srtofete/opnorm,0),;
               *        optoth WITH opdnight*ophday,opsumtot WITH opzph*optoth*oppost
               REPLACE opnorm WITH EVALUATE('sprtime.t'+LTRIM(STR(i))),opzph WITH IIF(opnorm#0,srtofete/opnorm*IIF(i=1,1,1),0),;
                       optoth WITH (opdnight*ophday)+(opdnight2*ophday2),opsumtot WITH (opdnight*ophday*opzph*oppost)+(opdnight2*ophday2*opzph*oppost2)
               repStrFete=STR(opdnight,2)+STR(ophday,5,2)+STR(optoth,3)+STR(oppost,2)+STR(opzph,8,2)+STR(opsumtot,8,2)+STR(opnorm,6,2)+STR(oppost2,2)+STR(ophday2,5,2)+STR(opdnight2,2)  
               REPLACE &repFete WITH repStrFete                                                                    
            ELSE                
               REPLACE &repFete WITH ''
            ENDIF            
        ENDFOR
ENDCASE
********************************************************************************************************************************************************
PROCEDURE prnFeteForm
fSupl=CREATEOBJECT('FORMSUPL')    
=ACOPY(dim_month,dim_monthprn)
DIMENSION dim_monthprn(13)
dim_monthprn(13)='год' 
month_prn=month_ch
WITH fSupl
     .procexit='DO exitFromFetePrn'
     .logExit=.T.
     .MaxButton=.F.
     .MinButton=.F.
     .Caption='Расчёт расходов на дополнительную оплату за работу в государственные праздники и праздничные дни'
     .BackColor=RGB(255,255,255)   
     DO addShape WITH 'fSupl',1,20,20,100,100,8 
     DO addComboMy WITH 'fSupl',2,fSupl.Shape1.Left+20,fSupl.Shape1.Top+10,dHeight,400,.T.,'month_prn','dim_monthprn',5,.F.,.F.,.F.,.T. 
     .comboBox2.DisplayCount=13
     .Shape1.Height=dHeight+20
     .Shape1.Width=.comboBox2.Width+40
     
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.F.     
     
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('wпросмотрw')*3-40)/2,.Shape91.Top+.Shape91.Height+20,RetTxtWidth('wпросмотрw'),dHeight+5,'печать','DO prnFete WITH .T.' ,'Печать ведомости'
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'просмотр','DO prnFete','Предварительный просмотр и печать ведомости'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','Do exitFromFetePrn','Выход из печати'    
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+80     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
********************************************************************************************************************************************************
PROCEDURE prnFete 
PARAMETERS par1
IF USED('curNight')
   SELECT curNight
   USE
ENDIF 
SELECT curFeteKat
REPLACE sumtot WITH 0 ALL
opsumtotcx=0
flt_month='rasp.fete'+LTRIM(STR(month_prn))
strnorm_month=''

IF month_prn<13
   SELECT sprtime
   GO TOP
   DO WHILE !EOF()
      strnorm_month=strnorm_month+ALLTRIM(name)+' - '+STR(EVALUATE('t'+LTRIM(STR(month_prn))),6,2)+','
      SKIP
   ENDDO
   strnorm_month=LEFT(strnorm_month,LEN(strnorm_month)-1)
   SELECT rasp
   SELECT * FROM rasp WHERE !EMPTY(&flt_month) INTO CURSOR curNight READWRITE
   SELECT curNight
   IF RECCOUNT()=0
      SELECT rasp
      RETURN
   ENDIF
   REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL                     
   REPLACE nameD WITH IIF(SEEK(curNight.kd,'sprdolj',1),sprdolj.name,'') ALL  
   INDEX ON STR(np,3)+STR(nd,3) TAG T1       
   GO TOP 
   nd_new=1
   kpold=kp
   DO WHILE !EOF()
      DO repstrfete WITH 'curNight',.F.,'month_prn'
      REPLACE nd WITH nd_new
      IF SEEK(curNight.kat,'curFeteKat',1)
         REPLACE curFeteKat.sumTot WITH curFeteKat.sumTot+opsumtot        
      ENDIF
      opsumtotcx=opsumtotcx+opsumtot
      SKIP
      nd_new=nd_new+1
      IF kp#kpold
         nd_new=1
         kpold=kp
      ENDIF
      
   ENDDO
   SELECT curFeteKat
   SCAN ALL
        SELECT curNight
        APPEND BLANK
        REPLACE np WITH 999,named WITH curFeteKat.name,opsumtot WITH curFeteKat.sumtot
        SELECT curFeteKat
   ENDSCAN
   SELECT curNight
   APPEND BLANK
   REPLACE np WITH 999,nameD WITH 'Итого',opsumtot WITH opsumtotcx  
   DELETE FOR opsumtot=0   
   GO TOP
ELSE
   SELECT * FROM rasp WHERE !EMPTY(fete1).OR.!EMPTY(fete2).OR.!EMPTY(fete3).OR.!EMPTY(fete4).OR.;
                            !EMPTY(fete5).OR.!EMPTY(fete6).OR.!EMPTY(fete7).OR.!EMPTY(fete8).OR.;         
                            !EMPTY(fete9).OR.!EMPTY(fete10).OR.!EMPTY(fete11).OR.!EMPTY(fete12) INTO CURSOR curNight READWRITE
  
   REPLACE nameD WITH IIF(SEEK(curNight.kd,'sprdolj',1),sprdolj.name,'') ALL   
   REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL                     
   REPLACE opdnight WITH 0,ophday WITH 0,optoth WITH 0,oppost WITH 0,opzph WITH 0,opsumtot WITH 0,opdnight WITH  0,opnorm WITH 0  ALL                  
   INDEX ON STR(np,3)+STR(nd,3) TAG T1  
   GO TOP
   DO WHILE !EOF()
      kvo_month=0
      FOR i=1 TO 12
          repFete='fete'+LTRIM(STR(i))
          IF VAL(SUBSTR(&repFete,21,8))>0
             kvo_month=kvo_month+1
             REPLACE opdnight WITH opdnight+VAL(SUBSTR(&repFete,1,2)),;
                     ophday WITH ophday+VAL(SUBSTR(&repFete,3,5)),;
                     optoth WITH optoth+VAL(SUBSTR(&repFete,8,3)),;
                     oppost WITH VAL(SUBSTR(&repFete,11,2)),;         
                     opzph WITH VAL(SUBSTR(&repFete,13,8)),;
                     opsumtot WITH opsumtot+VAL(SUBSTR(&repFete,21,8)),;           
                     opdnight WITH IIF(opdnight=0,kvoDayFete,opdnight),;
                     opnorm WITH VAL(SUBSTR(&repFete,29,6))                               
          ENDIF
      ENDFOR
      IF kvo_month#0
         REPLACE ophday WITH ophday/kvo_month,opzph WITH opsumtot/(optoth*oppost)
      ENDIF
      SKIP
   ENDDO 
   DELETE FOR opsumtot=0
   GO TOP
   nd_new=1
   kpold=kp
   DO WHILE !EOF()   
      REPLACE nd WITH nd_new
      IF SEEK(curNight.kat,'curFeteKat',1)
         REPLACE curFeteKat.sumTot WITH curFeteKat.sumTot+opsumtot        
      ENDIF
      opsumtotcx=opsumtotcx+opsumtot
      SKIP
      nd_new=nd_new+1
      IF kp#kpold
         nd_new=1
         kpold=kp
      ENDIF     
   ENDDO
   SELECT curFeteKat
   SCAN ALL
        SELECT curNight
        APPEND BLANK
        REPLACE np WITH 999,named WITH curFeteKat.name,opsumtot WITH curFeteKat.sumtot
        SELECT curFeteKat
   ENDSCAN
   SELECT curNight  
   APPEND BLANK
   REPLACE np WITH 999,nameD WITH 'Итого',opsumtot WITH opsumtotcx  
   DELETE FOR opsumtot=0   
   GO TOP
ENDIF   

IF par1
   FOR i=1 TO kvo_page
       SELECT curNight
       GO TOP       
       DO CASE
          CASE dimCht(1)=1
               IF month_prn<13
                  Report Form repfete NOCONSOLE TO PRINTER RANGE page_beg,page_end          
               ELSE 
                  Report Form repfetetot NOCONSOLE TO PRINTER RANGE page_beg,page_end          
               ENDIF 
          CASE dimCht(2)=1                       
               FOR c_range=page_beg TO page_end         
                   IF MOD(c_range,2)=0
                      IF month_prn<13
                         Report Form repfete RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ELSE 
                         Report Form repfetetot RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ENDIF 
                   ENDIF  
                   IF EOF()
                      EXIT 
                   ENDIF  
               ENDFOR                                
          CASE dimCht(3)=1     
               FOR c_range=page_beg TO page_end         
                   IF MOD(c_range,2)#0
                      IF month_prn<13
                         Report Form repfete RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ELSE 
                         Report Form repfetetot RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ENDIF 
                   ENDIF  
                   IF EOF()
                      EXIT 
                   ENDIF  
               ENDFOR             
       ENDCASE                 
   ENDFOR    
ELSE
   IF month_prn<13
      DO previewrep WITH 'repfete','Расчёт расходов на дополнительную оплату за работу в государственные праздники и праздничные дни' 
   ELSE 
      DO previewrep WITH 'repfetetot','Расчёт расходов на дополнительную оплату за работу в государственные праздники и праздничные дни'  
   ENDIF   
ENDIF   
SELECT rasp
********************************************************************************************************************************************************
PROCEDURE formdownloadfete
fSupl=CREATEOBJECT('FORMSUPL')
IF USED('curDatShtat')
   SELECT curDatShtat
   USE
ENDIF
SELECT * FROM datshtat WHERE lUse INTO CURSOR curDatShtat READWRITE
SELECT curDatShtat
ALTER TABLE curDatShtat ADD COLUMN nameSupl C(70)
INDEX ON real TAG T1 DESCENDING
REPLACE namesupl WITH IIF(real,DTOC(DATE()),DTOC(dTarif))+' '+ALLTRIM(fullName) ALL
LOCATE FOR ALLTRIM(pathtarif)=pathTarSupl
strDate=namesupl

WITH fSupl
     .Caption='Загрузка расчета из другого периода'
     procexut='DO returnDownloadFete'
     DO addshape WITH 'fSupl',1,20,20,150,380,8 
     DO adtBoxAsCont WITH 'fSupl','contDate',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('wвыберите дата тарификацииw'),dHeight,'выберите дату тарификации',2,1     
     DO addcombomy WITH 'fSupl',1,.contDate.Left+.contdate.Width-1,.contDate.Top,dHeight,300,.T.,'strDate','ALLTRIM(curDatShtat.namesupl)',6,'',.F.,.F.,.T.
     .comboBox1.DisplayCount=15
     .Shape1.Width=.contDate.Width+.comboBox1.Width+40
     .Shape1.height=.comboBox1.Height+40
     
     DO addButtonOne WITH 'fSupl','btnAply',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wвыполнитьw')*2)-20)/2,.Shape1.Top+.Shape1.Height+25,'выполнить','','DO downloadfete',39,RetTxtWidth('wвыполнитьw'),'выполнить'  
     DO addButtonOne WITH 'fSupl','btnRet',.btnAply.Left+.btnAply.Width+3,.btnAply.Top,'возврат','','DO returnDownloadFete',39,.btnAply.Width,'возврат'   

     DO addButtonOne WITH 'fSupl','btnOk',.shape1.Left+(.Shape1.Width-RetTxtWidth('wвозвратw'))/2,.btnAply.Top,'возврат','','DO returnDownloadFete',39,RetTxtWidth('wвозвратw'),'удаление' 
     .btnOk.Visible=.F.
     
      DO adLabMy WITH 'fSupl',25,'загрузка выполнена',.Shape1.Top+3,.Shape1.Left,.Shape1.Width,2,.F.,0
     .lab25.Top=.Shape1.Top+.Shape1.Height
     .lab25.Visible=.F.
               
     .Width=.Shape1.Width+40
     .Height=.btnAply.Top+.btnAply.Height+25
ENDWITH 
DO pasteimage WITH 'fSupl'
fSupl.Show
********************************************************************************************************************************************************
PROCEDURE downloadfete
pathfete=pathmain+'\'+ALLTRIM(curdatshtat.pathtarif)+'\rasp.dbf'
IF !FILE(pathfete).OR.ALLTRIM(curdatshtat.pathtarif)=pathTarSupl
    RETURN
ENDIF
SELECT 0
USE &pathfete ALIAS sourcefete
SET ORDER TO 2
SELECT rasp
SET FILTER TO 
REPLACE fete1 WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'sourcefete',2),sourcefete.fete1,''),fete2 WITH sourcefete.fete2,fete3 WITH sourcefete.fete3,fete4 WITH sourcefete.fete4,fete5 WITH sourcefete.fete5,;
fete6 WITH sourcefete.fete6,fete7 WITH sourcefete.fete7,fete8 WITH sourcefete.fete8,fete9 WITH sourcefete.fete9,fete10 WITH sourcefete.fete10,;
fete11 WITH sourcefete.fete11,fete12 WITH sourcefete.fete12,ntime WITH sourcefete.ntime, dtf WITH sourcefete.dtf ALL 
SELECT sourcefete
USE
SELECT rasp
DO validPodFete
WITH fSupl
     .btnAply.Visible=.F.
     .btnRet.Visible=.F.
     .lab25.Visible=.T.
     .btnOk.Visible=.T.
ENDWITH 
********************************************************************************************************************************************************
PROCEDURE returnDownloadFete
SELECT rasp
fForm.fGrid.Columns(fForm.fGrid.ColumnCount).SetFocus
fSupl.Release
********************************************************************************************************************************************************
PROCEDURE exitFromFetePrn
SELECT rasp
GO TOP
fSupl.Release
********************************************************************************************************************************************************
PROCEDURE setupFete