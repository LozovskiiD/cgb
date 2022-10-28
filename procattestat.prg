RESTORE FROM dim_vr ADDITIVE
RESTORE FROM dim_vr1 ADDITIVE 
RESTORE FROM setupvr ADDITIVE
var_path=FULLPATH('setupvr.mem')

SELECT people
SET FILTER TO 
SELECT datjob
SET FILTER TO 
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
SELECT rasp
SET FILTER TO 
REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),namekat WITH IIF(SEEK(kat,'sprkat',1),sprkat.name,'') ALL
GO TOP
SELECT * FROM sprtime INTO CURSOR curSprTime READWRITE
SELECT curSprTime
APPEND BLANK
INDEX ON name TAG T1
kpdop=IIF(rasp.kp=0,1,rasp.kp)
curnamepodr=IIF(SEEK(kpdop,'sprpodr',1),sprpodr.name,'')
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
fshtat=CREATEOBJECT('FORMMY')
WITH fshtat
     .Caption='Надбавка за вредность'
     .procexit='DO exitvred'
     .AddProperty('ksedol',0)
     DO addcontmenu WITH 'fshtat','menucont1',10,5,'должность','briefcase.ico','DO procDolVred'
     DO addcontmenu WITH 'fshtat','menucont2',.menucont1.Left+.menucont1.Width+3,5,'персонал','group.ico','DO procPersVred'
     DO addcontmenu WITH 'fshtat','menucont3',.menucont2.Left+.menucont2.Width+3,5,'печать','print1.ico','DO procvredprn'
     DO addcontmenu WITH 'fshtat','menucont4',.menucont3.Left+.menucont3.Width+3,5,'расчёт','calculate.ico',"DO createForm WITH .T.,'Расчет',RetTxtWidth('WWУдалить выбранную запись?WW',dFontName,dFontSize+1),;
        '130',RetTxtWidth('WWНетWW',dFontName,dFontSize+1),'Да','Нет',.F.,'DO countvred','nFormMes.Release',.F.,'Выполнить расчет?'"        
     DO addcontmenu WITH 'fshtat','menucont5',.menucont4.Left+.menucont4.Width+3,5,'настройка','setup.ico','DO procsetupvred'
     DO addcontmenu WITH 'fshtat','menucont6',.menucont5.Left+.menucont5.Width+3,5,'возврат','undo.ico','Do exitvred'    
        
     
     DO addComboMy WITH 'fshtat',1,10,.menucont1.Top+.menucont1.Height+15,dHeight,500,.T.,'curnamepodr','curSprPodr.name',6,.F.,'DO validPodVred',.F.,.T.  
     .comboBox1.Left=(.Width-.comboBox1.Width)/2
     .comboBox1.DisplayCount=15
     
     .AddObject('grdShtat','GridMy')
     WITH .grdshtat
          .Top=.Parent.comboBox1.Top+.Parent.comboBox1.Height+5
          .Left=0
          .Width=.Parent.Width
          .Height=.Parent.Height/2
          .ScrollBars=2
          .RecordSource='rasp'
          .ColumnCount=7  
          .Column1.ControlSource='rasp.nd'
          .Column1.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH(' 123 ',dFontName,dFontSize)
          .Column1.Header1.Caption='№'
          .Column2.ControlSource='rasp.named'
          .Column2.Alignment=0
          .Column2.Header1.Caption='Должность'
          .Column3.ControlSource='rasp.namekat'
          .Column3.ControlSource='rasp.kse'
          .Column3.Format='Z'
          .Column3.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH(' 123.99 ',dFontName,dFontSize)
          .Column3.Header1.Caption='К-во'
          .Column4.ControlSource='rasp.kfvr'
          .Column4.Format='Z'
          .Column4.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH(' 123.99 ',dFontName,dFontSize)
          .Column4.Header1.Caption='кфт1'     
          .Column5.ControlSource='rasp.pkfvr'
          .Column5.Format='Z'
          .Column5.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH(' 123.99 ',dFontName,dFontSize)
          .Column5.Header1.Caption='кфт2'         
          
          .Column6.ControlSource="IIF(SEEK(rasp.vtime,'sprtime',1),sprtime.name,'')"     
          .Column6.Width=RetTxtWidth('Wмедицинские сестрыWW')
          .Column6.Header1.Caption='время'
          .Column6.Alignment=0
          
 
          .Columns(.ColumnCount).Width=0
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-13-SYSMETRIC(15)-.ColumnCount                     
          .ProcAfterRowColChange='DO peopleincursor'  
          .SetAll('BOUND',.T.,'Column')   
     ENDWITH
     DO gridSize WITH 'fshtat','grdshtat','shapeingrid'   
   
     DO mycolumnTxtBox WITH 'fshtat.grdshtat.Column4','tBox4'
     WITH .grdshtat.Column4.tbox4
         .procForValid='DO validkftvr'
     ENDWITH
     DO mycolumnTxtBox WITH 'fshtat.grdshtat.Column5','tBox5'

     .AddObject('grdpers','GridMy')
     WITH .grdpers
          .Top=.Parent.grdshtat.Top+.Parent.grdshtat.Height+5
          .Left=0
          .Width=.Parent.Width
          .Height=.Parent.Height-.Top
          .ScrollBars=2
          .ColumnCount=8
          .RecordSource='datjob' 
          .Column1.ControlSource='datjob.kodpeop'
          .Column2.ControlSource='" "+datjob.fio'        
          .Column3.ControlSource="IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.name,'')"
          .Column4.ControlSource="IIF(SEEK(datjob.vtime,'sprtime',1),sprtime.name,'')"     
          .column5.ControlSource='datjob.kse'
          .Column6.ControlSource='datjob.pkfvr'
          .Column7.ControlSource='datjob.sumvr'   
          .Column1.Header1.Caption='код'    
          .Column2.Header1.Caption='Фамилия Имя Отчество'
          .Column3.Header1.Caption='Должность'
          .Column4.Header1.Caption='Время'
          .Column5.Header1.Caption='Объем' 
          .Column6.Header1.Caption='Кфт.'     
          .Column7.Header1.Caption='сумма в час'
          .Column1.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH(' 12345 ',dFontName,dFontSize)   
          .Column4.Width=RetTxtWidth('Wмедицинские сестрыWW')
          .Column5.Width=RetTxtWidth('999.999')
          .Column6.Width=RetTxtWidth('9999999')
          .Column7.Width=RetTxtWidth('9999999999')
          .Column2.Width=(.Width-.Column1.Width-.Column4.Width-.Column5.Width-.column6.Width-.Column7.Width)/2
          .Column3.Width=.Width-.Column1.Width-.Column2.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-SYSMETRIC(15)-13-.ColumnCount 
          .Columns(.ColumnCount).Width=0
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0
          .Column5.Alignment=1
          .Column6.Alignment=1
          .Column7.Alignment=1
          .Column6.Format='Z'
          .Column7.Format='Z'
     ENDWITH

    
     DO gridSize WITH 'fshtat','grdpers','shapeingrid1'  
     .grdpers.Height=.grdpers.height-.grdpers.RowHeight 
     DO adtBox WITH 'fshtat',1,10,.Grdpers.Top+.grdpers.Height-1,.Grdpers.column1.Width+2,fshtat.grdpers.RowHeight
     DO adtBox WITH 'fshtat',2,.TxtBox1.Left+.TxtBox1.Width-1,.TxtBox1.Top,.Grdpers.column2.Width+2,.TxtBox1.Height
     DO adtBox WITH 'fshtat',3,.TxtBox2.Left+.TxtBox2.Width-1,.TxtBox1.Top,.Grdpers.column3.Width+2,.TxtBox1.Height
     DO adtBox WITH 'fshtat',4,.TxtBox3.Left+.TxtBox3.Width-1,.TxtBox1.Top,.Grdpers.column4.Width+2,.TxtBox1.Height
     DO adtBox WITH 'fshtat',5,.TxtBox4.Left+.TxtBox4.Width-1,.TxtBox1.Top,.Grdpers.column5.Width+2,.TxtBox1.Height,'fshtat.ksedol','Z'
     DO adtBox WITH 'fshtat',6,.TxtBox5.Left+.TxtBox5.Width-1,.TxtBox1.Top,.Grdpers.column6.Width+2,.TxtBox1.Height
     DO adtBox WITH 'fshtat',7,.TxtBox6.Left+.TxtBox6.Width-1,.TxtBox1.Top,.Grdpers.column7.Width+2,.TxtBox1.Height
     .txtbox4.ForeColor=IIF(.ksedol#rasp->kse,objcolorsos,dForeColor)
     .TxtBox1.Enabled=.F.
     .TxtBox2.Enabled=.F.
     .TxtBox3.Enabled=.F.
     .TxtBox4.Enabled=.F.  
	 .grdShtat.Column1.SetFocus
ENDWITH
fshtat.Show
********************************************************************************************************************************************************
PROCEDURE validPodVred
SELECT curSprPodr
kpdop=curSprPodr.kod
curnamepodr=fShtat.ComboBox1.Value
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
fShtat.Refresh
**************************************************************************************************************************
*                                      Выбор персонала по должности
**************************************************************************************************************************
PROCEDURE peopleincursor
SELECT datjob
SET FILTER TO kp=rasp.kp.AND.kd=rasp.kd 
SUM kse TO fshtat.ksedol
fshtat.txtbox4.DisabledForeColor=IIF(fshtat.ksedol#rasp->kse,objcolorsos,dForeColor)
GO TOP
fshtat.grdpers.SetAll('ForeColor',dForeColor,'Header')
fshtat.grdpers.SetAll('BackColor',headerBackColor,'Header') 
SELECT rasp
fshtat.Refresh()
************************************************************************************************************************
PROCEDURE procDolVred
fSupl=CREATEOBJECT('FORMSUPL')
SELECT rasp
kft1=kfvr
kft2=pkfvr
newTime=vTime
newVrMain=rasp.vrMain
*newTime=ntime     && вид времени
saypodr=IIF(SEEK(kpdop,'sprpodr',1),ALLTRIM(sprpodr.name),'')
saydol=IIF(SEEK(rasp.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')

SELECT curSprTime
LOCATE FOR kod=rasp.vtime
strtime=curSprTime.name

WITH fSupl
     .Caption='Редактирование должности'
     DO adLabMy WITH 'fSupl',1,saypodr,10,0,fSupl.Width,2,.F.,0
     DO adLabMy WITH 'fSupl',2,saydol,.lab1.Top+.lab1.Height,2,fSupl.Width,2,.F.,0  
     DO addShape WITH 'fSupl',1,10,.lab2.Top+.lab2.Height,dHeight,300,8                    
     
     
     DO adTBoxAsCont WITH 'fsupl','txtkft1',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('WwВремяWw'),dHeight,'КФт1',0,1          
     DO adTboxNew WITH 'fSupl','boxKft1',.txtKft1.Top,.txtKft1.Left+.txtKft1.Width-1,250,dHeight,'Kft1',.F.,.T.,.F.,.F.,'DO validKft1'  
     
     DO adTBoxAsCont WITH 'fsupl','txtKft2',.txtKft1.Left,.txtKft1.Top+.txtKft1.Height-1,.txtKft1.Width,dHeight,'Кфт2',0,1           
     DO adTboxNew WITH 'fSupl','boxkft2',.txtKft2.Top,.boxKft1.Left,.boxKft1.Width,dHeight,'Kft2','Z',.T.,.F.,'9.99'
       
     DO adTBoxAsCont WITH 'fsupl','txtTime',.txtKft1.Left,.txtKft2.Top+.txtKft2.Height-1,.txtKft1.Width,dHeight,'Время',0,1      
     DO addComboMy WITH 'fSupl',1,.txtTime.Left+.txtTime.Width-1,.txtTime.Top,dHeight,250,.T.,'strTime','ALLTRIM(curSprTime.name)',6,.F.,'DO validTimeDol',.F.,.T.
       
                              
                                                     
     .Shape1.Height=.txtKft1.Height*3+20-3
     .Shape1.Width=.txtKft1.Width+.boxKft1.Width+20  
     DO adCheckBox WITH 'fSupl','check1','особая отметка',.Shape1.Top+.Shape1.Height+10,.Shape1.Left,150,dHeight,'newVrMain',0                                       
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
 
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,;
      .check1.Top+.check1.Height+20,RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать','DO writeDolVred'    
    
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Возврат','fSupl.Release'
     .Width=.txtkft1.Width+.boxkft1.Width+40
     .Height=.lab1.Height*2+.Shape1.Height+.cont1.Height+.check1.Height+50
     .lab1.Width=.Width
     .lab2.Width=.Width
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validkft1
kft2=IIF(ASCAN(dim_vr,kft1)=0,0,dim_vr1(ASCAN(dim_vr,kft1)))    
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE validtimedol
newTime=curSprtime.kod
strtime=cursprtime.name
KEYBOARD '{TAB}'  
fSupl.Refresh

*************************************************************************************************************************
PROCEDURE writeDolVred
fSupl.Release
SELECT rasp
REPLACE kfvr WITH kft1,pkfvr WITH kft2,vTime WITH newTime,vrmain WITH newVrMain

SELECT datjob
IF setupvr(1)
   IF setupvr(2)
      REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr*kse,vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0.AND.kse=1.AND.!vac   
   ELSE 
      REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr*kse,vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0.AND.IIF(rasp.vrMain,kfvr#0,kse=1)   
   ENDIF    
ELSE 
   IF setupvr(2)
      REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr,vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0.AND.!vac   
   ELSE  
      REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr,vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0 
   ENDIF 
ENDIF

SELECT datjob
GO TOP
DO WHILE !EOF()
   sum_tot=0
   FOR i=1 TO 12
       rep_cx='vr'+LTRIM(STR(i))
       repTime='sprtime.t'+LTRIM(STR(i))
       REPLACE &rep_cx WITH sumvr*IIF(SEEK(vTime,'sprtime',1),&repTime,0)
       sum_tot=sum_tot+&rep_cx
   ENDFOR
   REPLACE sumvrtot WITH sum_tot
   SELECT datjob
   SKIP
ENDDO
SELECT rasp
*************************************************************************************************************************
PROCEDURE procPersVred
fSupl=CREATEOBJECT('FORMSUPL')
SELECT datjob
kft1=pkfvr
newTime=vTime
newLatt=latt
*newTime=ntime     && вид времени
sayfam=ALLTRIM(datjob.fio)
saydol=IIF(SEEK(datjob.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')

SELECT curSprTime
LOCATE FOR kod=datjob.vtime
strtime=curSprTime.name

WITH fSupl
     .Caption='Редактирование персонала'
     DO adLabMy WITH 'fSupl',1,sayfam,10,0,fSupl.Width,2,.F.,0
     DO adLabMy WITH 'fSupl',2,saydol,.lab1.Top+.lab1.Height,2,fSupl.Width,2,.F.,0  
     DO addShape WITH 'fSupl',1,10,.lab2.Top+.lab2.Height,dHeight,300,8                    
     
     
     DO adTBoxAsCont WITH 'fsupl','txtkft1',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('WwВремяWw'),dHeight,'КФт',0,1          
     DO adTboxNew WITH 'fSupl','boxKft1',.txtKft1.Top,.txtKft1.Left+.txtKft1.Width-1,250,dHeight,'Kft1',.F.,.T.,.F.,.F.,.F.     
    
       
     DO adTBoxAsCont WITH 'fsupl','txtTime',.txtKft1.Left,.txtKft1.Top+.txtKft1.Height-1,.txtKft1.Width,dHeight,'Время',0,1      
     DO addComboMy WITH 'fSupl',1,.txtTime.Left+.txtTime.Width-1,.txtTime.Top,dHeight,250,.T.,'strTime','ALLTRIM(curSprTime.name)',6,.F.,'DO validTimePers',.F.,.T.
     
     DO adCheckBox WITH 'fSupl','check1','менее 1 ставки',.txtTime.Top+.txtTime.Height+10,.Shape1.Left,150,dHeight,'newLatt',0,.F.
     .check1.Enabled=IIF(datjob.kse=1,.F.,.T.)                        
                                                     
     .Shape1.Height=.txtKft1.Height*2+.check1.Height+30-3
     .Shape1.Width=.txtKft1.Width+.boxKft1.Width+20                                         
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,;
      .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать','DO writePersVred'    
    
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Возврат','fSupl.Release'
     .Width=.txtkft1.Width+.boxkft1.Width+40
     .Height=.lab1.Height*2+.Shape1.Height+.cont1.Height+40
     .lab1.Width=.Width
     .lab2.Width=.Width
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validTimePers
newTime=curSprtime.kod
strtime=cursprtime.name
KEYBOARD '{TAB}'  
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE writePersVred
fSupl.Release
SELECT datjob
REPLACE pkfvr WITH kft1,vTime WITH newTime,sumvr WITH varBaseSt/100*pkfvr*kse,latt WITH newlatt
sum_tot=0
FOR i=1 TO 12
    rep_cx='vr'+LTRIM(STR(i))
    repTime='sprtime.t'+LTRIM(STR(i))
    REPLACE &rep_cx WITH sumvr*IIF(SEEK(datjob.vTime,'sprtime',1),&repTime,0)   
    sum_tot=sum_tot+&rep_cx
ENDFOR
REPLACE sumvrtot WITH sum_tot
*************************************************************************************************************************
PROCEDURE procsetupvred
fsetup=CREATEOBJECT('FORMSUPL')
WITH fsetup
      DO addShape WITH 'fSetup',1,20,20,dHeight,380,8              
     DO adCheckBox WITH 'fsetup','check1','исключать лиц с объёмом работы менее 1.0 ставки',.Shape1.Top+10,.Shape1.Left+10,150,dHeight,'setupvr(1)',0
     DO adCheckBox WITH 'fsetup','check2','исключать вакантные',.check1.Top+.check1.Height+10,.Check1.Left,150,dHeight,'setupvr(2)',0  
     .Shape1.Height=.check1.Height*2+30 
     .Shape1.Width=.check1.Width+20        
     .Caption='настройки'    
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+40
     .procexit='SAVE TO &var_path ALL LIKE setupvr'
ENDWITH
SELECT rasp
DO pasteImage WITH 'fsetup'
fsetup.Show
*************************************************************************************************************************
PROCEDURE procvredprn
IF USED('peoplevred')
   SELECT peoplevred
   USE
ENDIF
DIMENSION sum_podr(6),sum_tot(6),dim_vred(12),dimpodr_vred(12),ksekat_podr(6),dim_sum(12)
STORE 0 TO sum_tot,dim_kat,sum_podr,numrecrep,ksekat_podr,numpage,dim_vred,dimpodr_vred,dim_sum
ksekat_podr=0
SELECT rasp
rrec=RECNO()
SELECT datjob
oldrec=RECNO()
SELECT * FROM datjob WHERE sumvr#0 INTO CURSOR peoplevred READWRITE
ALTER TABLE peoplevred ADD COLUMN nIt N(1)
ALTER TABLE peoplevred ADD COLUMN npp N(2)
ALTER TABLE peoplevred ADD COLUMN hourTot N(7,2)
ALTER TABLE peoplevred ALTER COLUMN kse N(7,2)
ALTER TABLE peoplevred ALTER COLUMN sumVrTot N(8,2)
SELECT peoplevred
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE hourTot WITH IIF(SEEK(vTime,'sprtime',1),sprtime.t1+sprtime.t2+sprtime.t3+sprtime.t4+sprtime.t5+sprtime.t6+sprtime.t7+sprtime.t8+sprtime.t9+sprtime.t10+sprtime.t11+sprtime.t12,0) ALL 
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
GO TOP 
kpOld=kp
nppNew=1
DO WHILE !EOF()
   REPLACE npp WITH nppNew  
   FOR i=1 TO 12
       vrcx='vr'+LTRIM(STR(i))
       dim_sum(i)=dim_sum(i)+&vrcx
   ENDFOR  
   SKIP 
   nppNew=nppNew+1
   IF kp#kpOld
      nppNew=1
      kpOld=kp
   ENDIF
ENDDO
kpold=kp
npold=np
STORE 0 TO ksePodr,ksetot,sumVrTot_podr,sumvrtot_tot
SUM kse,sumvrtot TO ksetot,sumvrtot_tot FOR nIt=0
APPEND BLANK
REPLACE np WITH 999,FIO WITH 'Итого',kse WITH ksetot,sumVrtot WITH sumvrtot_tot,nIt WITH 9
FOR i=1 TO 12
    APPEND BLANK
    REPLACE np WITH 999,nd WITH i,sumvrtot WITH dim_sum(i),fio WITH dim_month(i)
ENDFOR
SELECT sprpodr
GO TOP
DO WHILE !EOF()
   SELECT peoplevred
   SUM kse,sumvrtot TO ksePodr,sumVrTot_podr FOR kp=sprpodr.kod.AND.nIt=0   
   IF sumVrTot_podr#0
     APPEND BLANK
     REPLACE np WITH sprpodr.np,kp WITH sprpodr.kod,nd WITH 99,fio WITH 'Всего',kse WITH ksePodr,sumVrtot WITH sumvrtot_podr,nIt WITH 1  
   ENDIF
   SELECT sprpodr
   SKIP   
ENDDO 
SELECT peoplevred
GO TOP
DO printreport WITH 'repvred1','надбавка за вредность','peoplevred'
SELECT datjob
GO oldrec
SELECT rasp
GO rrec

*************************************************************************************************************************
*                  Процедура общего расчета  надбавки за вредность
*************************************************************************************************************************
PROCEDURE countvred
nformmes.Visible=.F.
nformmes.Release
SELECT datjob
SET FILTER TO 
SELECT rasp
nrec=RECNO()
SET FILTER TO 
SCAN ALL
     SELECT datjob
     REPLACE kfvr WITH 0,pkfvr WITH 0,sumvr WITH 0 FOR kp=rasp.kp.AND.kd=rasp.kd.AND.!rasp.vrMain
     SELECT rasp
ENDSCAN

SELECT rasp
GO TOP
DO WHILE !EOF()
   IF rasp.kfvr>0.AND.rasp.pkfvr>0
      SELECT datjob              
      IF setupvr(1)
         IF setupvr(2)
            REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr*IIF(kse=1.OR.latt,kse,0),vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0.AND.kse=1.AND.!vac
         ELSE
            REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr**IIF(kse=1.OR.latt,kse,0),vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0.AND.IIF(rasp.vrMain,kfvr#0,kse=1)  
         ENDIF   
      ELSE
         IF setupvr(2)
            REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr,vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0.AND.!vac 
         ELSE   
            REPLACE kfvr WITH rasp.kfvr,pkfvr WITH rasp.pkfvr,sumvr WITH varBaseSt/100*pkfvr,vtime WITH rasp.vTime FOR kp=rasp->kp.AND.kd=rasp->kd.AND.pcharw=0 
         ENDIF 
      ENDIF           
   ENDIF   
   SELECT rasp 
   SKIP
ENDDO 
SELECT datjob

GO TOP
DO WHILE !EOF()
   IF pkfvr#0
      sum_tot=0
      DO CASE
         CASE setupvr(1)
              FOR i=1 TO 12                 
                  rep_cx='vr'+LTRIM(STR(i))
                  repTime='sprtime.t'+LTRIM(STR(i))         
                  REPLACE &rep_cx WITH IIF(kse=1.OR.latt,sumvr*IIF(SEEK(datjob.vTime,'sprtime',1),&repTime,0),0)         
                  sum_tot=sum_tot+&rep_cx                          
              ENDFOR
              REPLACE sumvrtot WITH sum_tot
         CASE !setupvr(1)
              FOR i=1 TO 12
                  rep_cx='vr'+LTRIM(STR(i))
                  repTime='sprtime.t'+LTRIM(STR(i))         
                  REPLACE &rep_cx WITH sumvr*IIF(SEEK(datjob.vTime,'sprtime',1),&repTime,0)         
                  sum_tot=sum_tot+&rep_cx      
              ENDFOR
              REPLACE sumvrtot WITH sum_tot
      ENDCASE
      
   ELSE      
      FOR i=1 TO 12
          rep_cx='vr'+LTRIM(STR(i))
          repTime='sprtime.t'+LTRIM(STR(i))         
          REPLACE &rep_cx WITH 0                 
      ENDFOR
      REPLACE sumvrtot WITH 0,vtime WITH 0 
   ENDIF    
   SELECT datjob
   SKIP
ENDDO
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
GO nrec
DO createFormNew WITH .T.,'Общий расчёт',RetTxtWidth('WWРасчёт выполнен!WW',dFontName,dFontSize+1),'130',;
      RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
      'Расчёт выполнен!',.F.,.T. 
*************************************************************************************************************************
*                              Выход из развернутого штатного
*************************************************************************************************************************
PROCEDURE exitvred
SELECT people
SET FILTER TO 
SELECT datjob
SET FILTER TO 
SELECT rasp
SET FILTER TO 
GO TOP
fshtat.Release
