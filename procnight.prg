* Структура полей night1-night12
* 1-2   - кол-во дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата в час с учётом процентов1 (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-36 - зарплата в час по норме (оперативное поле - opzpnorm)
* 37-42 - часов по первому проценту (оперативное поле - hourp)
* 43-48 - часов по второму проценту (оперативное поле - hourp2)
* 49-56 - зарплата в час с учётом процентов2 (оперативное поле - opzph2)
* 57-58 - кол-во должностей 2 (оперативное поле - oppost2)
*RESTORE FROM dim_night ADDITIVE
*var_path=FULLPATH('dim_night.mem')

dCount=varDtar
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

SELECT * FROM sprkat INTO CURSOR curFeteKat READWRITE
ALTER TABLE curFeteKat ADD COLUMN sumTot N(9,2)
SELECT curFeteKat
INDEX ON kod TAG T1
SELECT * FROM sprtime INTO CURSOR curSprtime READWRITE
SELECT cursprtime
INDEX ON kod TAG T1

SELECT rasp
SET FILTER TO 
SET RELATION TO kd INTO sprdolj,ntime INTO sprtime ADDITIVE
GO TOP
curnamepodr=''
kvoDayMonth=0
PUBLIC kpDop
STORE 0 TO kpdop
kpdop=IIF(rasp.kp=0,1,rasp.kp)
log_read=.F.
log_count=.F.
month_ch=1
repNight=''

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
DO repstrnight WITH 'rasp',.T.,'month_ch'
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
     .Caption='Расчёт расходов на доплату в ночное время или в ночную смену'
     .procExit='DO exitFromProcNight'
     
     DO addButtonOne WITH 'fForm','menuCont1',10,5,'редакция','pencil.ico','DO readnight',39,RetTxtWidth('wзагрузть изw')+44,'редакция'  
     DO addButtonOne WITH 'fForm','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'расчёт','calculate.ico','DO procCountNight',39,.menucont1.Width,'расчёт'   
     DO addButtonOne WITH 'fForm','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delNight',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fForm','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico','DO prnNightForm',39,.menucont1.Width,'печать'   
    * DO addButtonOne WITH 'fForm','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'в Excel','excel.ico','DO excelNightForm',39,.menucont1.Width,'в Excel'  
     DO addButtonOne WITH 'fForm','menuCont6',.menucont4.Left+.menucont4.Width+3,5,'календарь','ical.ico','DO proccalendar',39,.menucont1.Width,'календарь'  
     DO addButtonOne WITH 'fForm','menuCont7',.menucont6.Left+.menucont6.Width+3,5,'загрузить из','get.ico','DO formdownloadnight',39,.menucont1.Width,'загрузить из предыдущего периода'   
     DO addButtonOne WITH 'fForm','menuCont8',.menucont7.Left+.menucont7.Width+3,5,'возврат','undo.ico','DO exitFromProcNight',39,.menucont1.Width,'возврат'  
     
     DO addButtonOne WITH 'fForm','menuexit1',10,5,'возврат','undo.ico','DO exitReadNight',39,RetTxtWidth('возврат')+44,'вовзрат' 
     .menuexit1.Visible=.F.
       
     DO addComboMy WITH 'fform',1,(.Width-715)/2,.menucont1.Top+.menucont1.Height+15,dHeight,500,.T.,'curnamepodr','cursprpodr.name',6,.F.,;
        'DO validPodNight',.F.,.T.               
     DO addComboMy WITH 'fform',2,.comboBox1.Left+.comboBox1.Width+15,.comboBox1.Top,dHeight,200,.T.,'month_ch','dim_month',5,.F.,;
        'DO validMonthNight',.F.,.T.   
     .comboBox2.DisplayCount=12                         
     WITH .fGrid    
          .Top=fForm.ComboBox1.Top+fForm.ComboBox1.Height+5          
          .Height=fForm.Height-.Top
          .Width=fForm.Width
          .RecordSource='rasp'
          DO addColumnToGrid WITH 'fForm.fGrid',18
          .RecordSourceType=1        
          .Column1.ControlSource='" "+sprdolj->name'
          .Column2.ControlSource='rasp.opdnight'     
          .Column3.ControlSource='rasp.ophday'
          .Column4.ControlSource='rasp.optoth'                                                                                               
          .Column5.ControlSource='rasp.srtonight'          
          .Column6.ControlSource='sprtime.name'          
          .Column7.ControlSource="EVALUATE('sprtime.t'+LTRIM(STR(month_ch)))"
          .Column8.ControlSource='rasp.opzpnorm'                    
          .Column9.ControlSource='rasp.oppost'   
          .Column10.ControlSource='rasp.persnight'
          .Column11.ControlSource='rasp.hourp'
          .Column12.ControlSource='rasp.opzph'
          .Column13.ControlSource='rasp.oppost2' 
          .Column14.ControlSource='rasp.persnight2'          
          .Column15.ControlSource='rasp.hourp2'          
          .Column16.ControlSource='rasp.opzph2'
          .Column17.ControlSource='rasp.opsumtot'
                    
          .Column2.Width=RettxtWidth('99999')
          .Column3.Width=RettxtWidth('99999')
          .Column4.Width=RetTxtWidth('сотр.')                                       
          .Column5.Width=RetTxtWidth('999999999')
          .Column6.Width=RetTxtWidth('wсредний медперсоw')
          .Column7.Width=RetTxtWidth('9999999')
          .Column8.Width=RetTxtWidth(' час.окл. ')          
          .Column9.Width=RetTxtWidth('99999')  
          .Column10.Width=RetTxtWidth('9999')
          .Column11.Width=RetTxtWidth('9999999')
          .Column12.Width=RetTxtWidth('99999999')
          .Column13.Width=RetTxtWidth('99999')  
          .Column14.Width=RetTxtWidth('9999')
          .Column15.Width=RetTxtWidth('9999999')
          .Column16.Width=RetTxtWidth('99999999')
          .Column17.Width=RetTxtWidth('99999999')
          
          .Columns(.ColumnCount).Width=0   
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-;
                         .Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-;
                         .Column12.Width-.Column13.Width-.Column14.Width-.Column15.Width-.Column16.Width-.Column17.Width-SYSMETRIC(5)-13-.ColumnCount          
          .Column1.Header1.Caption='должность'
          .Column2.Header1.Caption='дней'
          .Column3.Header1.Caption='ч-д'
          .Column4.Header1.Caption='ч-м'       
          .Column5.Header1.Caption='ср.окл'
          .Column6.Header1.Caption='время'
          .Column7.Header1.Caption='норма'
          .Column8.Header1.Caption='в час'
          .Column9.Header1.Caption='сотр'  
          .Column10.Header1.Caption='%'
          .Column11.Header1.Caption='часов'
          .Column12.Header1.Caption='час.окл.' 
          .Column13.Header1.Caption='сотр2' 
          .Column14.Header1.Caption='%2'
          .Column15.Header1.Caption='часов'
          .Column16.Header1.Caption='час.окл.'                                                  
          .Column17.Header1.Caption='сумма.м.'             
        
          .Column2.Format='Z'
          .Column3.Format='Z'
          .Column4.Format='Z'
          .Column5.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'
          .Column9.Format='Z'
          .Column10.Format='Z'  
          .Column11.Format='Z'  
          .Column12.Format='Z'
          .Column13.Format='Z'  
          .Column14.Format='Z'  
          .Column15.Format='Z' 
          .Column16.Format='Z'
          .Column17.Format='Z'
          .Column7.Bound=.T.                          
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .SetAll('Alignment',1,'ColumnMy')  
          .Column1.Alignment=0         
          .Column6.Alignment=0
          .colNesInf=2              
     ENDWITH
     DO gridSizeNew WITH 'fForm','fGrid','shapeingrid',.F.,.T.          
     FOR i=1 TO fForm.fGrid.columnCount       
         fForm.fGrid.Columns(i).fontname=dFontName
         fForm.fGrid.Columns(i).fontSize=dFontSize      
         fForm.fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fForm.fGrid.RecordSource)#fForm.fGrid.curRec,dBackColor,dynBackColor)'
         fForm.fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fForm.fGrid.RecordSource)#fForm.fGrid.curRec,dForeColor,dynForeColor)'                                                              
         fForm.fGrid.Columns(i).Text1.SelectedForeColor=dynForeColor
         fForm.fGrid.Columns(i).Text1.SelectedBackColor=dynBackColor
         fForm.fGrid.Columns(i).Header1.ForeColor=dForeColor        
         fForm.fGrid.Columns(i).Header1.BackColor=headerBackColor 
         fForm.fGrid.Columns(i).Header1.FontName=dFontName        
         fForm.fGrid.Columns(i).Header1.FontSize=dFontSize                 
     ENDFOR 
     .combobox1.DisplayCount=MIN(RECCOUNT('sprpodr'),(fForm.Height-fForm.combobox1.Top-fForm.combobox1.Height)/fForm.fGrid.Rowheight)                          
ENDWITH
SELECT rasp
GO TOP
fForm.Show
********************************************************************************************************************************************************
PROCEDURE exitFromProcNight
*SAVE TO &var_path ALL LIKE dim_night
SELECT rasp
SET FILTER TO 
SET RELATION TO
SELECT people
SET FILTER TO 
SELECT datJob
SET FILTER TO
fForm.Release
*************************************************************************************************************************
PROCEDURE procvalidtime
SELECT rasp
REPLACE ntime WITH cursprtime.kod
KEYBOARD '{TAB}'    
fForm.Refresh
************************************************************************************************************************
PROCEDURE procgottime
SELECT cursprtime
LOCATE FOR kod=sprtime->kod
nrec=RECNO()
GO TOP 
COUNT WHILE RECNO()#nrec TO varnrec
fForm.fGrid.column6.cbotime.DisplayCount=MAX(fForm.fGrid.RelativeRow,fForm.fGrid.RowsGrid-fForm.fGrid.RelativeRow)
fForm.fGrid.column6.cbotime.DisplayCount=MIN(fForm.fGrid.column6.cbotime.DisplayCount,RECCOUNT())
fForm.fGrid.Column6.cbotime.varCtrlSource=varnrec+1 
SELECT rasp
*******************************************************************************************************************************************************
*                             Процедура замены данных оперативных полей
*******************************************************************************************************************************************************
PROCEDURE repstrnight
PARAMETERS parDbf,parScope,parMonth
*parDbf - rasp или curNight
*parScope - текущая запись или все
*parMonth - переменная периода
* 1-2   - кол-во дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата в час с учётом процентов (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-36 - зарплата в час по норме (оперативное поле - opzpnorm)
* 37-42 - часов по первому проценту (оперативное поле - hourp)
* 43-48 - часов по второму проценту (оперативное поле - hourp2)
* 49-56 - зарплата в час с учётом процентов1 (оперативное поле - opzph2)
* 57-58 - кол-во должностей2 (оперативное поле - oppost2)
SELECT &parDbf
repNight='night'+LTRIM(STR(&parMonth))
*kvoDayMonth=IIF(&parMonth=2,IIF(MOD(YEAR(dim_night(1)),4)=0,29,28),IIF(INLIST(&parMonth,4,6,9,11),30,31))
kvoDayMonth=IIF(&parMonth=2,IIF(INLIST(YEAR(dCount),2020,2024,2028,2032,2036,2040),29,28),IIF(INLIST(&parMonth,4,6,9,11),30,31))  
IF parScope
   REPLACE opdnight WITH VAL(SUBSTR(&repNight,1,2)),;
           ophday WITH VAL(SUBSTR(&repNight,3,5)),;
           optoth WITH VAL(SUBSTR(&repNight,8,3)),;
           oppost WITH VAL(SUBSTR(&repNight,11,2)),;          
           opzph WITH VAL(SUBSTR(&repNight,13,8)),;
           opsumtot WITH VAL(SUBSTR(&repNight,21,8)),;
           opzpnorm WITH VAL(SUBSTR(&repNight,29,8)),;
           hourp WITH VAL(SUBSTR(&repNight,37,6)),;
           hourp2 WITH VAL(SUBSTR(&repNight,43,6)),;
           opzph2 WITH VAL(SUBSTR(&repNight,49,8)),;
           opdnight WITH kvoDayMonth,;
           oppost2 WITH VAL(SUBSTR(&repNight,57,2))  ALL
ELSE 
   REPLACE opdnight WITH VAL(SUBSTR(&repNight,1,2)),;
           ophday WITH VAL(SUBSTR(&repNight,3,5)),;
           optoth WITH VAL(SUBSTR(&repNight,8,3)),;
           oppost WITH VAL(SUBSTR(&repNight,11,2)),;         
           opzph WITH VAL(SUBSTR(&repNight,13,8)),;
           opsumtot WITH VAL(SUBSTR(&repNight,21,8)),;
           opzpnorm WITH VAL(SUBSTR(&repNight,29,8)),;
           hourp WITH VAL(SUBSTR(&repNight,37,6)),;
           hourp2 WITH VAL(SUBSTR(&repNight,43,6)),;
           opzph2 WITH VAL(SUBSTR(&repNight,49,8)),;
           opdnight WITH IIF(opdnight=0,kvoDayMonth,opdnight),;
           oppost2 WITH VAL(SUBSTR(&repNight,57,2)) 
ENDIF            
********************************************************************************************************************************************************
PROCEDURE validPodNight
SELECT cursprpodr
kpdop=cursprpodr.kod
curnamepodr=fForm.ComboBox1.Value
SELECT rasp
SET FILTER TO kp=kpdop
DO repstrnight WITH 'rasp',.T.,'month_ch'
fForm.fGrid.Column7.ControlSource="EVALUATE('sprtime.t'+LTRIM(STR(month_ch)))"
GO TOP
fForm.Refresh
********************************************************************************************************************************************************
PROCEDURE validMonthNight
month_ch=fForm.comboBox2.Value
DO repstrnight WITH 'rasp',.T.,'month_ch'
SELECT rasp 
GO TOP
fForm.Refresh
********************************************************************************************************************************************************
*                                                  Редактирование ночных (новых вариант)
********************************************************************************************************************************************************
PROCEDURE readNight
IF USED('curSuplDol')
   SELECT curSuplDol
   USE
ENDIF
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
IF rasp.dtn#0
   SEEK STR(rasp.kp,3)+STR(rasp.dtn,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.dtn
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
ENDIF
SELECT rasp       

newTime=ntime                                        && вид времени
newNorm=EVALUATE('sprtime.t'+LTRIM(STR(month_ch)))   && норма времени
newOpdNight=opdnight                                 && дней в месяце
newOpHDay=ophday                                     && часов в месяц
newOpToth=optoth                                     && часов в месяц
newSrOkl=ROUND(IIF(kvo_peop#0,sumto/kvo_peop,0),2)   && средний оклад
newOpzpNorm=opzpnorm                                 && з/п в час
newOpPost=oppost                                    && сотрудников
newPersNight=persNight                               && %
newHourp=hourp                                       && часов
newOpzph=opZph                                       && часовой оклад
newOpPost2=oppost2                                  && сотрудников2
newPersNight2=persNight2                             && %2
newHourp2=hourp2                                     && часов2
newOpzph2=opZph2                                     && часовой оклад2
newOpSumTot=opSumTot                                 && на месяц
newDtn=dtn
logDtn=IIF(dtn>0,.T.,.F.)


saypodr=IIF(SEEK(kpdop,'sprpodr',1),ALLTRIM(sprpodr.name),'')
saydol=IIF(SEEK(rasp.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')
sayDolSupl=IIF(SEEK(rasp.dtn,'sprdolj',1),ALLTRIM(sprdolj.name),'')
SELECT * FROM sprdolj WHERE SEEK(STR(raspKp,3)+STR(kod,3),'rasp',2) INTO CURSOR curSuplDol READWRITE
SELECT rasp
GO raspRec

SELECT curSprTime
LOCATE FOR kod=rasp.ntime
strtime=curSprTime.name

WITH fSupl
     .Caption='Ввод-редакция затрат на оплату работы в ночное время'
     .procExit='DO writeNight WITH .F.'
     DO adLabMy WITH 'fSupl',1,saypodr,10,0,fSupl.Width,2,.F.,0
     DO adLabMy WITH 'fSupl',2,saydol,.lab1.Top+.lab1.Height,2,fSupl.Width,2,.F.,0  
     DO addShape WITH 'fSupl',1,10,.lab2.Top+.lab2.Height,dHeight,300,8                   
      
     DO adTBoxAsCont WITH 'fsupl','txtTime',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('wПраздничных днейw'),dHeight,'время',0,1
     DO addComboMy WITH 'fSupl',1,.txtTime.Left+.txtTime.Width-1,.txtTime.Top,dHeight,250,.T.,'strTime','ALLTRIM(curSprTime.name)',6,.F.,'DO validTime',.F.,.T.
       
     DO adTBoxAsCont WITH 'fsupl','txtNorm',.txtTime.Left,.txtTime.Top+.txtTime.Height-1,.txtTime.Width,dHeight,'норма',0,1           
     DO adTboxNew WITH 'fSupl','boxNorm',.txtNorm.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNorm',.F.,.F.
     
     DO adTBoxAsCont WITH 'fsupl','txtOklad',.txtTime.Left,.txtNorm.Top+.txtNorm.Height-1,.txtTime.Width,dHeight,'средний оклад',0,1     
     DO adTboxNew WITH 'fSupl','boxOklad',.txtOklad.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSrOkl','Z',.F.,0     
          
     DO adTBoxAsCont WITH 'fsupl','txtDayTot',.txtTime.Left,.txtOklad.Top+.txtOklad.Height-1,.txtTime.Width,dHeight,'дней в месяце',0,1
     DO adTboxNew WITH 'fSupl','boxDayTot',.txtDayTot.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpDNight','Z',.T.,.F.,.F.,'DO validRashNight'
     
     DO adTBoxAsCont WITH 'fsupl','txtHourDay',.txtTime.Left,.txtDayTot.Top+.txtDayTot.Height-1,.txtTime.Width,dHeight,'часов в день',0,1
     DO adTboxNew WITH 'fSupl','boxHourDay',.txtHourDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpHDay','Z',.T.,.F.,.F.,'DO validRashNight'
     
     DO adTBoxAsCont WITH 'fsupl','txtHourTot',.txtTime.Left,.txtHourDay.Top+.txtHourDay.Height-1,.txtTime.Width,dHeight,'часов в месяц',0,1
     DO adTboxNew WITH 'fSupl','boxHourTot',.txtHourTot.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpToth','Z',.F.
     
     DO adTBoxAsCont WITH 'fsupl','txtZpHour',.txtTime.Left,.txtHourTot.Top+.txtHourTot.Height-1,.txtTime.Width,dHeight,'з/п в час',0,1
     DO adTboxNew WITH 'fSupl','boxZpHour',.txtZpHour.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpzpNorm','Z',.F.,0
                  
     DO adTBoxAsCont WITH 'fsupl','txtSotr',.txtTime.Left,.txtZpHour.Top+.txtZpHour.Height-1,.txtTime.Width,dHeight,'сотрудников',0,1
     DO adTboxNew WITH 'fSupl','boxSotr',.txtSotr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpPost','Z',.T.,0,.F.,'DO validRashNight'        
     
     DO adTBoxAsCont WITH 'fsupl','txtPers',.txtTime.Left,.txtSotr.Top+.txtSotr.Height-1,.txtTime.Width,dHeight,'%',0,1
     DO adTboxNew WITH 'fSupl','boxPers',.txtPers.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPersNight','Z',.T.,0,.F.,'DO validRashNight'
     
     DO adTBoxAsCont WITH 'fsupl','txtHour1',.txtTime.Left,.txtPers.Top+.txtPers.Height-1,.txtTime.Width,dHeight,'часов',0,1
     DO adTboxNew WITH 'fSupl','boxHour1',.txtHour1.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newHourp','Z',.T.,0,.F.,'DO validRashNight'
     
     DO adTBoxAsCont WITH 'fsupl','txtOpZph',.txtTime.Left,.txtHour1.Top+.txtHour1.Height-1,.txtTime.Width,dHeight,'часов. окл.',0,1
     DO adTboxNew WITH 'fSupl','boxOpZoh',.txtOpZph.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpZph','Z',.F.,0
     
     DO adTBoxAsCont WITH 'fsupl','txtSotr2',.txtTime.Left,.txtOpZph.Top+.txtOpZph.Height-1,.txtTime.Width,dHeight,'сотрудников2',0,1
     DO adTboxNew WITH 'fSupl','boxSotr2',.txtSotr2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOppost2','Z',.T.,0,.F.,'DO validRashNight'        
     
     DO adTBoxAsCont WITH 'fsupl','txtPers2',.txtTime.Left,.txtSotr2.Top+.txtSotr2.Height-1,.txtTime.Width,dHeight,'%2',0,1
     DO adTboxNew WITH 'fSupl','boxPers2',.txtPers2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPersNight2','Z',.T.,0,.F.,'DO validRashNight'
     
     DO adTBoxAsCont WITH 'fsupl','txtHour2',.txtTime.Left,.txtPers2.Top+.txtPers2.Height-1,.txtTime.Width,dHeight,'часов2',0,1
     DO adTboxNew WITH 'fSupl','boxHour2',.txtHour2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newHourp2','Z',.T.,0,.F.,'DO validRashNight'
     
     DO adTBoxAsCont WITH 'fsupl','txtOpZph2',.txtTime.Left,.txtHour2.Top+.txtHour2.Height-1,.txtTime.Width,dHeight,'часов. окл.2',0,1
     DO adTboxNew WITH 'fSupl','boxOpZoh2',.txtOpZph2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpZph2','Z',.F.,0
     
     DO adTBoxAsCont WITH 'fsupl','txtSumTot',.txtTime.Left,.txtOpZph2.Top+.txtOpZph2.Height-1,.txtTime.Width,dHeight,'на месяц',0,1
     DO adTboxNew WITH 'fSupl','boxSumTot',.txtSumTot.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOpSumTot','Z',.F.,0
                                                                                             
     .Shape1.Height=.txtTime.Height*16+20-16
     .Shape1.Width=.txtTime.Width+.comboBox1.Width+20 
     
     DO addShape WITH 'fSupl',2,10,.Shape1.Top+.Shape1.Height+10,100,.Shape1.Width,8                                           
     DO adCheckBox WITH 'fSupl','check1','объединить с ',.Shape2.Top+10,.Shape2.Left,.Shape2.Width,dHeight,'logdtn',0,.T.,'DO validCheckSupl'
     DO addComboMy WITH 'fSupl',2,.Shape2.Left+10,.check1.Top+.check1.Height+10,dHeight,.Shape2.Width-20,IIF(newDtn>0,.T.,.F.),'sayDolSupl','ALLTRIM(curSuplDol.name)',6,.F.,'DO validSuplDol',.F.,.T.
     .check1.Left=.Shape2.Left+(.Shape2.Width-.check1.Width)/2
     .Shape2.Height=.check1.Height+.comboBox2.Height+30
          
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,;
        .Shape2.Top+.Shape2.Height+20,RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать','DO writeNight WITH .T.'    
    
    
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Возврат','DO writeNight WITH .F.'
     .Width=.Shape1.Width+20
     .Height=.lab1.Height*2+.Shape1.Height+.Shape2.Height+.cont1.Height+50
     .lab1.Width=.Width
     .lab2.Width=.Width
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validCheckSupl
fSupl.comboBox2.Enabled=IIF(logDtn,.T.,.F.)
IF !logDtn
   newDtn=0
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
   DO validRashNight
   sayDolSupl=''
   fSupl.comboBox2.ControlSource='sayDolSupl'
   fSupl.Refresh
   
ENDIF 
*************************************************************************************************************************
PROCEDURE validSuplDol
newDtn=curSuplDol.kod
IF newDtn#0
   SELECT curTarJob
   STORE 0 TO kvo_peop,sumto 
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
    ENDSCAN

   SEEK STR(rasp.kp,3)+STR(newDtn,3)
   SCAN WHILE kp=rasp.kp.AND.kd=newDtn
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
   SELECT rasp  
   newSrOkl=ROUND(IIF(kvo_peop#0,sumto/kvo_peop,0),2)      
   DO validRashNight
   fSupl.Refresh
ENDIF
*************************************************************************************************************************
PROCEDURE validtime
newTime=curSprtime.kod
newnorm=EVALUATE('cursprtime.t'+LTRIM(STR(month_ch)))
strtime=cursprtime.name
KEYBOARD '{TAB}'  
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE validRashNight

IF EMPTY(newHourp2).AND.EMPTY(newHourp)
   newHourp=newOptoth
ENDIF
newOpZpNorm=ROUND(newSrOkl/newNorm,2)
newOpToth=ROUND(newOpdNight*newOphDay,0)
newOpZph=ROUND(newOpZpNorm/100*newPersNight,2)
newOpZph2=ROUND(newOpZpNorm/100*newPersNight2,2)
newOpSumTot=ROUND((newOpZph*newHourp*newOppost)+(newOpZph2*NewOpPost2*newHourp2),2)
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE writeNight
PARAMETERS par1
SELECT rasp
oldrec=RECNO()
IF par1   
   REPLACE opdNight WITH newOpdNight,ophday WITH newOphDay,optoth WITH newOpToth,srtonight WITH newSrOkl,ntime WITH newTime,;
           opzpnorm WITH newOpzpNorm,oppost WITH newoppost,persnight WITH newPersNight,hourp WITH newHourp,opZph WITH newOpZph,;
           oppost2 WITH newoppost2,persnight2 WITH newPersNight2,hourp2 WITH newHourp2,opZph2 WITH newOpZph2,opsumTot WITH newopSumTot,dtn WITH newDtn  
   repNight='night'+LTRIM(STR(month_ch))        
   repStrNight=STR(opdnight,2)+STR(ophday,5,2)+STR(optoth,3)+STR(oppost,2)+;
               STR(opzph,8,2)+STR(opsumtot,8,2)+STR(opzpnorm,8,2)+STR(hourp,6,2)+STR(hourp2,6,2)+STR(opzph2,8,2)+STR(oppost2,2)
   IF opSumTot#0             
      REPLACE &repNight WITH repStrNight
   ENDIF   
ENDIF    

GO oldrec
fForm.fGrid.Columns(fForm.fGrid.ColumnCount).SetFocus
fForm.Refresh
fSupl.Release
********************************************************************************************************************************************************
PROCEDURE delNight
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
       (fdel.shape1.Width-20)/2,dHeight+3,'Выполнение','DO delrecRashNight'
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
PROCEDURE delrecRashNight
IF !log_del 
   RETURN 
ENDIF
SELECT rasp
DO CASE
   CASE dim_del(1)=1                   
        DO delRecNightOne
   CASE dim_del(2)=1        
        SELECT rasp        
        GO TOP
        DO WHILE !EOF()          
           DO delRecNightOne
           SKIP
        ENDDO                             
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()           
           DO delRecNightOne
           SKIP
        ENDDO                                 
ENDCASE
fDel.Release
SELECT rasp
SET FILTER TO kp=kpdop
DO repstrnight WITH 'rasp',.T.,'month_ch'
GO TOP
fForm.Refresh
*****************************************************************************************************************************************************
PROCEDURE delRecNightOne
SELECT rasp
DO CASE
   CASE dim_del(4)=1 
        REPLACE srtonight WITH 0  
        rep_ch='night'+LTRIM(STR(month_ch))
        DO repstrnight WITH 'rasp',.F.,'month_ch'      
        REPLACE &rep_ch WITH ' ' 
   CASE dim_del(5)=1
        REPLACE srtonight WITH 0
        FOR i=1 TO 12      
            rep_ch='night'+LTRIM(STR(i))
            REPLACE &rep_ch WITH ''                     
        ENDFOR
ENDCASE    
********************************************************************************************************************************************************
PROCEDURE procCountNight
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
     DO adtbox WITH 'fdel',1,fdel.lab1.Left+fdel.lab1.Width+10,fdel.Shape1.Top+10,RetTxtWidth('99999999999'),dHeight,'dCount','Z',.T.,1,
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
         
     DO adCheckBox WITH 'fdel','check2','подтверждение выполнения',fdel.Shape4.Top+fdel.Shape4.Height+10,fdel.Shape1.Left,150,dHeight,'log_del',0         
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('WВыполнениеW')*2-20)/2,fdel.check2.Top+fdel.check2.Height+15,;
       RetTxtWidth('WВыполнениеW'),dHeight+3,'Выполнение','DO countNight'
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
     .Height=.Shape1.Height+.Shape4.Height+fdel.cont1.Height+fdel.check2.Height+80
     .lab1.left=.Shape1.Left+(.shape1.Width-.lab1.Width-.txtbox1.Width-10)/2
     .txtbox1.Left=.lab1.Left+.lab1.Width+10
     .check2.Left=(.Width-.check2.Width)/2
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
****************************************************************************************************************************************************
*                           Непосредственно процедра общего расчёта расходов по Указу
****************************************************************************************************************************************************
PROCEDURE countNight
IF !log_del
   fDel.Release
   RETURN
ENDIF
monthOld=month_ch
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
        DO countOneRecNight      
        one_pers=one_pers+1
        pers_ch=one_pers/max_rec*100
        fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch 
   CASE dim_del(2)=1
        SELECT rasp
        GO TOP        
        DO WHILE !EOF()
           DO countOneRecNight
           SELECT rasp
           SKIP 
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch               
        ENDDO
        GO recPeop
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO
        COUNT TO max_rec 
        GO TOP        
        DO WHILE !EOF()
           DO countOneRecNight
           SELECT rasp
           SKIP
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch  
        ENDDO
        GO recPeop
ENDCASE
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
=INKEY(1)
fDel.Visible=.F.
fdel.Release
DO createFormNew WITH .T.,'Общий расчёт',RetTxtWidth('WWРасчёт выполнен!WW',dFontName,dFontSize+1),'130',;
      RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
      'Расчёт выполнен!',.F.,.T. 
=SYS(2002,1) 
log_count=.F.
month_ch=monthOld
DO repstrnight WITH 'rasp',.T.,'month_ch'
GO TOP
fForm.fGrid.Columns(fForm.fGrid.ColumnCount).SetFocus
fForm.Refresh 
****************************************************************************************************************************************************
*                                 Расчет расходов по одной записи
****************************************************************************************************************************************************
PROCEDURE countOneRecNight
* 1-2   - кол-во дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата в час с учётом процентов (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-36 - зарплата в час по норме (оперативное поле - opzpnorm)
* 37-42 - часов по первому проценту
* 43-48 - часов по второму проценту
SELECT curTarJob
STORE 0 TO kvo_peop,sumto 
SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
SCAN WHILE kp=rasp.kp.AND.kd=rasp.kd
     kvo_peop=kvo_peop+kse
     sumto=sumto+mtokl
ENDSCAN
IF rasp.dtn#0
   SEEK STR(rasp.kp,3)+STR(rasp.dtn,3)
   SCAN WHILE kp=rasp.kp.AND.kd=rasp.dtn
        kvo_peop=kvo_peop+kse
        sumto=sumto+mtokl
   ENDSCAN
ENDIF
SELECT rasp
REPLACE srtonight WITH IIF(kvo_peop#0,sumto/kvo_peop,0)
DO CASE
   CASE dim_del(4)=1    
        DO repstrnight WITH 'rasp',.F.,'month_ch'                      
        REPLACE opzpnorm WITH IIF(EVALUATE('sprtime.t'+LTRIM(STR(month_ch)))#0,srtonight/EVALUATE('sprtime.t'+LTRIM(STR(month_ch))),0)
        IF EMPTY(hourp2).AND.EMPTY(hourp)
           REPLACE hourp WITH optoth
        ENDIF
        REPLACE optoth WITH opdnight*ophday,;
        opzph WITH opzpnorm/100*persnight,;
        opzph2 WITH opzpnorm/100*persnight2,;     
        opsumtot WITH (opzph*hourp*oppost)+(opzph2*oppost2*hourp2)
        repStrNight=STR(opdnight,2)+STR(ophday,5,2)+STR(optoth,3)+STR(oppost,2)+;
                    STR(opzph,8,2)+STR(opsumtot,8,2)+STR(opzpnorm,8,2)+STR(hourp,6,2)+STR(hourp2,6,2)+STR(opzph2,8,2)+STR(oppost2,2) 
        REPLACE &repNight WITH repStrNight       
   CASE dim_del(5)=1              
        FOR i=1 TO 12
            DO repstrnight WITH 'rasp',.F.,'i'                     
            IF i>=MONTH(dCount)
               REPLACE opzpnorm WITH IIF(EVALUATE('sprtime.t'+LTRIM(STR(i)))#0,srtonight/EVALUATE('sprtime.t'+LTRIM(STR(i))),0)
               IF EMPTY(hourp2).AND.EMPTY(hourp)
                  REPLACE hourp WITH optoth
               ENDIF
               REPLACE optoth WITH hourp+hourp2
             *  REPLACE ophday WITH optoth/opdnight
               REPLACE optoth WITH opdnight*ophday,;
                       opzph WITH opzpnorm/100*persnight,;
                       opzph2 WITH opzpnorm/100*persnight2,;
                       opsumtot WITH (opzph*hourp*oppost)+(opzph2*oppost2*hourp2)
               repStrNight=STR(opdnight,2)+STR(ophday,5,2)+STR(optoth,3)+STR(oppost,2)+;
                           STR(opzph,8,2)+STR(opsumtot,8,2)+STR(opzpnorm,8,2)+STR(hourp,6,2)+STR(hourp2,6,2)+STR(opzph2,8,2)+STR(oppost2,2)  
               REPLACE &repNight WITH repStrNight                                                                      
            ELSE                
               REPLACE &repNight WITH ''
            ENDIF            
        ENDFOR
ENDCASE
********************************************************************************************************************************************************
PROCEDURE prnNightForm
fSupl=CREATEOBJECT('FORMSUPL')  
=ACOPY(dim_month,dim_monthprn)
DIMENSION dim_monthprn(13)
dim_monthprn(13)='год' 
month_prn=month_ch
logWord=.F.
WITH fSupl
     .procexit='DO exitFromNightPrn'
     .logExit=.T.
     .MaxButton=.F.
     .MinButton=.F.
     .Caption='Расчёт расходов на доплату в ночное время или в ночную смену'     
     DO addShape WITH 'fSupl',1,20,20,100,100,8     
     DO addComboMy WITH 'fSupl',2,fSupl.Shape1.Left+20,fSupl.Shape1.Top+20,dHeight,350,.T.,'month_prn','dim_monthprn',5,.F.,.F.,.F.,.T. 
     .comboBox2.DisplayCount=13
     .Shape1.Height=dHeight+40
     .Shape1.Width=.comboBox2.Width+40
     
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.F.  

     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnNight WITH .T.' ,'Печать ведомости'
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Просмотр','DO prnNight','Предварительный просмотр и печать ведомости'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','Do exitFromNightPrn','Выход из печати'    
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+80     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
********************************************************************************************************************************************************
PROCEDURE prnNight 
* 1-2   - кол-во дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата в час с учётом процентов1 (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-36 - зарплата в час по норме (оперативное поле - opzpnorm)
* 37-42 - часов по первому проценту (оперативное поле - hourp)
* 43-48 - часов по второму проценту (оперативное поле - hourp2)
* 49-56 - зарплата в час с учётом процентов1 (оперативное поле - opzph2)
PARAMETERS par1
IF USED('curNight')
   SELECT curNight
   USE
ENDIF 
SELECT curFeteKat
REPLACE sumtot WITH 0 ALL
opsumtotcx=0
IF month_prn<13
   flt_month='rasp.night'+LTRIM(STR(month_prn))
   strnorm_month=''
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
 
   REPLACE nameD WITH IIF(SEEK(curNight.kd,'sprdolj',1),sprdolj.name,'') ALL  
   REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
 
   INDEX ON STR(np,3)+STR(nd,3) TAG T1      
   DO repstrnight WITH 'curNight',.T.,'month_prn' 
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
ELSE       
   SELECT rasp  
   SELECT * FROM rasp WHERE !EMPTY(night1).OR.!EMPTY(night2).OR.!EMPTY(night3).OR.!EMPTY(night4).OR.;
                            !EMPTY(night5).OR.!EMPTY(night6).OR.!EMPTY(night7).OR.!EMPTY(night8).OR.;         
                            !EMPTY(night9).OR.!EMPTY(night10).OR.!EMPTY(night11).OR.!EMPTY(night12) INTO CURSOR curNight READWRITE
   SELECT curNight
   IF RECCOUNT()=0
      SELECT rasp
      RETURN
   ENDIF
   REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
   REPLACE nameD WITH IIF(SEEK(curNight.kd,'sprdolj',1),sprdolj.name,'') ALL    
   INDEX ON STR(np,3)+STR(nd,3) TAG T1 
   REPLACE opdNight WITH 0,ophday WITH 0,optoth WITH 0,oppost WITH 0,opzph WITH 0,;
           opsumtot WITH 0,opzpnorm WITH 0,opsumtot WITH 0,hourp WITH 0,hourp2 WITH 0,opzph2 WITH 0 ALL
   GO TOP 
   DO WHILE !EOF() 
      kvo_per=0       
      totnorm=0
      FOR i=1 TO 12          
          nightper='night'+LTRIM(STR(i))
          IF VAL(SUBSTR(&nightPer,21,8))>0
             kvo_per=kvo_per+1
          
             REPLACE opdnight WITH opdnight+VAL(SUBSTR(&nightPer,1,2)),;
                     ophday WITH ophday+VAL(SUBSTR(&nightPer,3,5)),;
                     optoth WITH optoth+VAL(SUBSTR(&nightPer,8,3)),;
                     oppost WITH oppost+VAL(SUBSTR(&nightPer,11,2)),;                   
                     opsumtot WITH opsumtot+VAL(SUBSTR(&nightPer,21,8)),;                               
                     hourp WITH hourp+VAL(SUBSTR(&nightPer,37,6)),;
                     hourp2 WITH hourp2+VAL(SUBSTR(&nightPer,43,6)),;
                     opzph2 WITH opzph2+VAL(SUBSTR(&nightPer,49,8))
                    totnorm=totnorm+EVALUATE('sprtime.t'+LTRIM(STR(i)))
                   * totnorm=totnorm+IIF(SEEK(curNight.ntime,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(i))),0)
                   *  IF totnorm=0
                   *     gkhkjhk
                   *  ENDIF 
          ENDIF         
      ENDFOR      
      IF kvo_per>0
         IF hourp2=0
            REPLACE ophday WITH ophday/kvo_per,oppost WITH oppost/kvo_per,opzph WITH opsumtot/(oppost*opdnight*ophday),;
                    opzpnorm WITH opzph/persnight*100     
                                
         ELSE 
            *REPLACE ophday WITH ophday/kvo_per,oppost WITH oppost/kvo_per,opzph WITH opsumtot/(oppost*opdnight*ophday),;
            *        opzph2 WITH opzph2/kvo_per,opzph WITH (opsumtot-(oppost*opzph2*hourp2))/(oppost*hourp),opzpnorm WITH opzph/persnight*100   
            totstrtonight=srtonight*kvo_per 
            REPLACE ophday WITH ophday/kvo_per,oppost WITH oppost/kvo_per,opzpnorm WITH IIF(totnorm#0,totstrtonight/totnorm,0),opzph WITH opzpnorm/100*persnight,opzph2 WITH opzpnorm/100*persnight2               
         ENDIF           
                 
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
IF par1.AND.!logWord
   FOR i=1 TO kvo_page
       SELECT curNight
       GO TOP       
       DO CASE
          CASE dimCht(1)=1
               IF month_prn<13
                  Report Form repnight NOCONSOLE TO PRINTER RANGE page_beg,page_end          
               ELSE 
                  Report Form repnighttot NOCONSOLE TO PRINTER RANGE page_beg,page_end          
               ENDIF 
          CASE dimCht(2)=1                       
               FOR c_range=page_beg TO page_end         
                   IF MOD(c_range,2)=0
                      IF month_prn<13
                         Report Form repnight RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ELSE 
                         Report Form repnighttot RANGE c_range,c_range NOCONSOLE TO PRINTER   
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
                         Report Form repnight RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ELSE 
                         Report Form repnighttot RANGE c_range,c_range NOCONSOLE TO PRINTER   
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
      IF par1.AND.logWord

      ELSE
         DO previewrep WITH 'repnight','Расчёт расходов на доплату в ночное время или в ночную смену' 
      ENDIF
   ELSE 
      IF par1.AND.logWord
      ELSE
         DO previewrep WITH 'repnighttot','Расчёт расходов на доплату в ночное время или в ночную смену' 
      ENDIF
   ENDIF   
ENDIF   
SELECT rasp

********************************************************************************************************************************************************
PROCEDURE exitFromNightPrn
SELECT rasp
GO TOP
fSupl.Release

********************************************************************************************************************************************************
PROCEDURE excelNightForm
fSupl=CREATEOBJECT('FORMSUPL') 
month_prn=month_ch
WITH fSupl
     .procexit='DO exitFromNightPrn'
     .logExit=.T.
     .MaxButton=.F.
     .MinButton=.F.
     .Caption='Расчёт оплаты ночных часов'     
     DO addShape WITH 'fSupl',1,20,20,100,100,8     
     DO addComboMy WITH 'fSupl',2,fSupl.Shape1.Left+20,fSupl.Shape1.Top+20,dHeight,350,.T.,'month_prn','dim_month',5,.F.,.F.,.F.,.T. 
     .comboBox2.DisplayCount=13
     .Shape1.Height=dHeight+40
     .Shape1.Width=.comboBox2.Width+40
          
     DO adLabMy WITH 'fSupl',24,'Ход выполнения',.Shape1.Top+.Shape1.Height+5,.Shape1.Left,.Shape1.Width,2,.F.,1

     DO addShape WITH 'fSupl',11,.Shape1.Left,.lab24.Top+.lab24.Height+5,dHeight,.Shape1.Width
     .Shape11.BackStyle=0
  
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,dHeight,0
     .Shape12.BackStyle=1
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+2,.Shape11.Left,.Shape11.Width,2,.F.,0
     .lab25.Visible=.F. 


          
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WПросмотрW')*2-20)/2,.Shape11.Top+.Shape11.Height+20,RetTxtWidth('WПросмотрW'),dHeight+5,'в Excel','DO exportNight' ,'Экспорт в Excel'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','Do exitFromNightExcel','Выход'    
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.Shape11.Height+.lab24.Height+.cont1.Height+70     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
********************************************************************************************************************************************************
PROCEDURE exportNight 
* 1-2   - кол-во дней в месяце  (оперативное поле - opdnight)
* 3-7   - кол-во часов работы в день (оперативное поле - ophday)
* 8-10  - часов работы в месяц (оперативное поле - optoth)
* 11-12 - кол-во должностей (оперативное поле - oppost)
* 13-20 - зарплата в час с учётом процентов1 (оперативное поле - opzph)
* 21-28 - зарплата за месяц (оперативное поле - opsumtot)
* 29-36 - зарплата в час по норме (оперативное поле - opzpnorm)
* 37-42 - часов по первому проценту (оперативное поле - hourp)
* 43-48 - часов по второму проценту (оперативное поле - hourp2)
* 49-56 - зарплата в час с учётом процентов1 (оперативное поле - opzph2)
fSupl.lab25.Visible=.T.
fSupl.lab25.Caption='0%'
SELECT * FROM curTarJob INTO CURSOR datjobExcel READWRITE
SELECT 	datjobExcel
INDEX ON STR(kp,3)+STR(kd,3)+STR(kv,1) TAG T1
CREATE CURSOR curExcel (kp N(3),kd N(3),namedol C(120),kodkat N(2),nkat C(30),tarokl N(8),kodnorma N(3),hnorma N(6,2),stavka N(8),pers N(3),oplata N(8))
SELECT curExcel
INDEX ON kp TAG T1
flt_month='rasp.night'+LTRIM(STR(month_prn))    
SELECT rasp
SELECT * FROM rasp WHERE !EMPTY(&flt_month) INTO CURSOR curNight READWRITE
SELECT curNight
IF RECCOUNT()=0
   SELECT rasp
   RETURN
ENDIF
INDEX ON STR(np,3)+STR(nd,3) TAG T1
DO repstrnight WITH 'curNight',.T.,'month_prn' 
DELETE FOR opsumtot=0
GO TOP 
nd_new=1
kpold=kp
DO WHILE !EOF()           
   REPLACE nd WITH nd_new    	
   SKIP
   nd_new=nd_new+1
   IF kp#kpold
      nd_new=1
      kpold=kp
   ENDIF
ENDDO
GO TOP
DO WHILE !EOF()
   SELECT datjobExcel
   IF SEEK(STR(curNight.kp,3)+STR(curNight.kd,3)+STR(1,1))
      SELECT curExcel 
      IF curNight.persnight#0
         APPEND BLANK
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,kodkat WITH 1,nkat WITH 'высшей категории', tarokl WITH datjobExcel.tokl,;
                 pers WITH curNight.persnight,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers
      ENDIF             
      IF curNight.persnight2#0
          APPEND BLANK
          REPLACE kp WITH curNight.kp,kd WITH curNight.kd,kodkat WITH 1,nkat WITH 'высшей категории', tarokl WITH datjobExcel.tokl,;
                  pers WITH curNight.persnight2,kodnorma WITH curnight.ntime,;
                  namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                  stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers
      ENDIF         
  
      SELECT datjobExcel    
   ENDIF
   
   IF SEEK(STR(curNight.kp,3)+STR(curNight.kd,3)+STR(2,1))
      SELECT curExcel     
      IF curNight.persnight#0
         APPEND BLANK
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,kodkat WITH 2,nkat WITH 'первой категории', tarokl WITH datjobExcel.tokl,pers WITH curNight.persnight,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers    
      ENDIF           
      IF curNight.persnight2#0
         APPEND BLANK 
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,kodkat WITH 2,nkat WITH 'первой категории', tarokl WITH datjobExcel.tokl,pers WITH curNight.persnight2,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers    
      ENDIF           
      SELECT datjobExcel    
   ENDIF
   
   IF SEEK(STR(curNight.kp,3)+STR(curNight.kd,3)+STR(3,1))
      SELECT curExcel
      IF curNight.persNight#0
         APPEND BLANK
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,kodkat WITH 3,nkat WITH 'второй категории', tarokl WITH datjobExcel.tokl,pers WITH curNight.persnight,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers    
      ENDIF 
      IF curNight.persNight2#0
         APPEND BLANK
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,kodkat WITH 3,nkat WITH 'второй категории', tarokl WITH datjobExcel.tokl,pers WITH curNight.persnight2,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers     
      ENDIF            
      SELECT datjobExcel    
   ENDIF
   
   IF SEEK(STR(curNight.kp,3)+STR(curNight.kd,3)+STR(0,1))
      SELECT curExcel
      IF curNight.persNight#0
         APPEND BLANK
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,nkat WITH 'без категории',tarokl WITH datjobExcel.tokl,pers WITH curNight.persnight,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers    
      ENDIF 
      IF curNight.persNight2#0
         APPEND BLANK           
         REPLACE kp WITH curNight.kp,kd WITH curNight.kd,nkat WITH 'без категории',tarokl WITH datjobExcel.tokl,pers WITH curNight.persnight2,kodnorma WITH curnight.ntime,;
                 namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),hnorma WITH IIF(SEEK(kodnorma,'sprtime',1),EVALUATE('sprtime.t'+LTRIM(STR(month_prn))),0),;
                 stavka WITH IIF(hnorma#0,tarokl/hnorma,0),oplata WITH stavka/100*pers    
      ENDIF    
      SELECT datjobExcel    
   ENDIF
   
   SELECT curNight
   SKIP
ENDDO
SELECT curexcel
GO TOP

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
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=27
     .Columns(3).ColumnWidth=15
     .Columns(4).ColumnWidth=7
     .Columns(5).ColumnWidth=5
     .Columns(6).ColumnWidth=6
     .Columns(7).ColumnWidth=4
     .Columns(8).ColumnWidth=6
     .Columns(9).ColumnWidth=7
         
     rowcx=3     
    .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                            
     rowcx=rowcx+1   
     
        
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select         
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Расчет'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
     rowcx=rowcx+1
            
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select         
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='оплаты ночных часов за '+dim_month(month_prn)
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
     rowcx=rowcx+1                                    
                                      
     .Range(.Cells(rowcx,1),.Cells(rowcx+2,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='№ п/п'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx+2,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Наименование должности'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH           
                 
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx+2,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='категория'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH       
        
      .Range(.Cells(rowcx,4),.Cells(rowcx+2,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='тарифный оклад'                    
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,5),.Cells(rowcx+2,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='месячная норма часов'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,6),.Cells(rowcx+2,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='часовая ставка'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,7),.Cells(rowcx+2,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='%'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                 
      
      .Range(.Cells(rowcx,8),.Cells(rowcx+2,8)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='оплата за час работы в ночное время'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH   
      
      .Range(.Cells(rowcx,9),.Cells(rowcx+2,9)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='примечание'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                             
      rowcx=rowcx+3      
      .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      rowtop=numberRow         
      SELECT curExcel
      STORE 0 TO max_rec,one_pers,pers_ch
      COUNT TO max_rec
      GO TOP
      kpold=0
      numnew=1
      fSupl.Shape12.Width=0
      fSupl.Shape12.Visible=.T. 
      SCAN ALL
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,9)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curExcel.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              numnew=1
              kpold=kp
           ENDIF 
          .Cells(numberRow,1).Value=numnew                                       
           .Cells(numberRow,2).Value=curExcel.namedol                                       
           .Cells(numberRow,3).Value=curExcel.nkat   
           .Cells(numberRow,4).Value=curExcel.tarokl                           
           .Cells(numberRow,5).NumberFormat='000.00'           
           .Cells(numberRow,5).Value=curExcel.hnorma
           .Cells(numberRow,6).Value=curExcel.stavka                                    
           .Cells(numberRow,7).Value=curExcel.pers     
           .Cells(numberRow,8).Value=curExcel.oplata  
           numnew=numnew+1                                  
           numberRow=numberRow+1
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
           fSupl.Shape12.Width=fSupl.Shape11.Width/100*pers_ch 
                   
      ENDSCAN                                 
      .Range(.Cells(rowcx-3,1),.Cells(numberRow-1,9)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
          
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,9)).Select
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
fSupl.Shape12.Visible=.F.  
fSupl.lab25.Visible=.F.          
objExcel.Visible=.T.

********************************************************************************************************************************************************
PROCEDURE formdownloadnight
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
     procexit='DO returnDownLoadNight'
     DO addshape WITH 'fSupl',1,20,20,150,380,8 
     DO adtBoxAsCont WITH 'fSupl','contDate',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('wвыберите дата тарификацииw'),dHeight,'выберите дату тарификации',2,1     
     DO addcombomy WITH 'fSupl',1,.contDate.Left+.contdate.Width-1,.contDate.Top,dHeight,300,.T.,'strDate','ALLTRIM(curDatShtat.namesupl)',6,'',.F.,.F.,.T.
     .comboBox1.DisplayCount=15
     .Shape1.Width=.contDate.Width+.comboBox1.Width+40
     .Shape1.height=.comboBox1.Height+40
     
     DO addButtonOne WITH 'fSupl','btnAply',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wвыполнитьw')*2)-20)/2,.Shape1.Top+.Shape1.Height+25,'выполнить','','DO downloadnight',39,RetTxtWidth('wвыполнитьw'),'выполнить'  
     DO addButtonOne WITH 'fSupl','btnRet',.btnAply.Left+.btnAply.Width+3,.btnAply.Top,'возврат','','DO returnDownLoadNight',39,.btnAply.Width,'возврат'   

     DO addButtonOne WITH 'fSupl','btnOk',.shape1.Left+(.Shape1.Width-RetTxtWidth('wвозвратw'))/2,.btnAply.Top,'возврат','','DO returnDownLoadNight',39,RetTxtWidth('wвозвратw'),'возврат' 
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
PROCEDURE downloadnight
pathnight=pathmain+'\'+ALLTRIM(curdatshtat.pathtarif)+'\rasp.dbf'
IF !FILE(pathnight).OR.ALLTRIM(curdatshtat.pathtarif)=pathTarSupl
   SELECT rasp  
   RETURN
ENDIF
SELECT 0
USE &pathnight ALIAS sourcenight
SET ORDER TO 2
SELECT rasp
SET FILTER TO 
REPLACE night1 WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'sourcenight',2),sourcenight.night1,''),persnight WITH sourcenight.persnight,persnight2 WITH sourcenight.persnight2,night2 WITH sourcenight.night2,night3 WITH sourcenight.night3,night4 WITH sourcenight.night4,night5 WITH sourcenight.night5,;
night6 WITH sourcenight.night6,night7 WITH sourcenight.night7,night8 WITH sourcenight.night8,night9 WITH sourcenight.night9,night10 WITH sourcenight.night10,;
night11 WITH sourcenight.night11,night12 WITH sourcenight.night12,ntime WITH sourcenight.ntime, dtf WITH sourcenight.dtf ALL 
SELECT sourcenight
USE
SELECT rasp
DO validPodnight
WITH fSupl
     .btnAply.Visible=.F.
     .btnRet.Visible=.F.
     .lab25.Visible=.T.
     .btnOk.Visible=.T.
ENDWITH 
********************************************************************************************************************************************************
PROCEDURE returnDownloadNight
SELECT rasp
fForm.fGrid.Columns(fForm.fGrid.ColumnCount).SetFocus
fSupl.Release
********************************************************************************************************************************************************
PROCEDURE exitFromNightExcel
SELECT rasp
GO TOP
fSupl.Release

********************************************************************************************************************************************************
PROCEDURE setupNight
