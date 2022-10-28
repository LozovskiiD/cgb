* наименование                                        %                сумма          на месяц
*за стаж                                             stpr              stsum          mstsum
*за контракт                                         pkont             konts          mkonts
*за руководство структурным подразделением           prstp             rstp           mrstp
*руковордителям интернов                             print             srint          msrint
*за старшинство                                      pstar             star           mstar

RESTORE FROM sum_pr ADDITIVE
RESTORE FROM dimConstVac ADDITIVE

USE datset IN 0

USE datshtat ORDER 1 IN 0

USE sprkat ORDER 1 IN 0

USE sprkval ORDER 1 IN 0

USE sprkoef ORDER 1 IN 0

USE sprtype ORDER 1 IN 0

SELECT * FROM sprtype INTO CURSOR curSprType ORDER BY kod
SELECT * FROM sprkat INTO CURSOR curSprKat ORDER BY kod
SELECT * FROM sprkval INTO CURSOR curSprKval READWRITE ORDER BY kod
SELECT cursprkval
APPEND BLANK
REPLACE name WITH 'без категории'
SELECT datshtat
LOCATE FOR ALLTRIM(pathTarif)=ALLTRIM(datset.pathlast).AND.lUse
IF !FOUND()
   LOCATE FOR real
ELSE 
   pathPeople=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\people.dbf'  
   IF !FILE(pathpeople)   
      SELECT datset
      REPLACE pathlast WITH 'REAL' 
      SELECT datshtat
      LOCATE FOR real
   ENDIF
ENDIF    
pathtarif=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\'+';'+pathmain+';'+pathsupl
pathcur=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\' && путь к каталогу тарификации с которой сейчас работают
pathTarSupl=ALLTRIM(datShtat.pathtarif) && каталог тарификации с которой сейчас работают   

varDtar=dTarif
varBaseSt=baseSt
varNmzp=nMzp
IF datShtat.real  
   varDtar=DATE()
   REPLACE datshtat.dtarif WITH varDtar
ENDIF  
*tarDateSay='тарификация на '+DTOC(varDTar)+' '+ALLTRIM(datshtat.fullname)+' (изменить - двойной щелчок мыши)'
tarDateSay='тарификация на '+DTOC(varDTar)+IIF(datshtat.real,' текущая ','')+' (изменить - двойной щелчок мыши)  базовая ставка - '+LTRIM(STR(varBaseSt,8,2))+'  мин.з.п - '+LTRIM(STR(varNmzp,8,2))
SET PATH TO &pathtarif
USE people ORDER 1 IN 0
USE rasp ORDER 1 IN 0
USE datjob ORDER 4 IN 0
USE sprpodr ORDER 1 IN 0
USE tarfond ORDER 1 IN 0
USE sprtime ORDER 1 IN 0
USE sprdolj ORDER 1 IN 0

SELECT * FROM sprdolj INTO CURSOR curSprDolj READWRITE ORDER BY name

USE boss IN 0
=AFIELDS(arFltJob,'datJob')
CREATE CURSOR curFltDatJob FROM ARRAY arFltJob
SELECT curFltDatJob
INDEX ON kodPeop TAG T1
SELECT * FROM sprpodr INTO CURSOR curSprPodr READWRITE ORDER BY name
SELECT * FROM tarfond WHERE nblock=2 INTO CURSOR curSupFond READWRITE 
SELECT curSupFond
INDEX ON num TAG T1
SELECT datJob
SET RELATION TO tr INTO curSprType, kat INTO curSprKat  ADDITIVE
PUBLIC peopRec,datJobRec  &&указатели записи в people,datJob
STORE 0 TO peopRec,datJobRec
nameJobKval=''
namevtime=''
strperstaj=''
fltJob=''
sostavflt=''
cLogMol=''
curRow=0
curRowJob=0
frmTop=CREATEOBJECT('Formtop')
WITH frmTop
     .AddProperty('log_fl',.F.)
     .Addproperty('filter_ch','')
     .Addproperty('filter_peop','')
     .Caption='Тарификация'
     .procExit='DO quitTarif'
     USE datmenu IN 0
     SELECT datmenu
     GO TOP
     DO addButtonOne WITH 'frmTop','menuContTop',10,3,'главное меню','topbottom.ico','DO menuRecieve',39,RetTxtWidth('главное меню')+44,'главное меню'  
     m_ch=1
     leftCont=.menuContTop.Left+.menuContTop.Width+10
     topCont=3
     DO WHILE !EOF()
        namecont='menucont'+LTRIM(STR(m_ch))
        DO addcontico WITH 'frmTop',namecont,leftCont,topCont,datmenu.mIco,datmenu.mproc,ALLTRIM(datmenu.nmenu),40,40
        SKIP
        m_ch=m_ch+1
        leftCont=frmTop.&namecont..Left+frmTop.&namecont..Width+5
     ENDDO        
     topObj=.menucont1.Top+.menucont1.Height+5
     DO addContFormNew WITH 'frmTop','contDateTarif',0,topObj,.Width,dHeight,tarDateSay,1,.F.,'DO choiceDateTar',.F.,.T. 
     topObj=.contDateTarif.Top+.contDateTarif.Height-1
     width_obj=.Width/4*3  && ширина части для личной карточки
     width_obj=ROUND(width_obj/2,0)
     widthLab=ROUND(width_obj/3,0)
     widthtxt=width_obj-widthlab
     width_obj=widthLab*4-3 
     width_obj=widthLab*2+widthtxt*2-3
     labLeft=0
     txtleft=lableft+widthLab-1
     DO adTboxAsCont WITH 'frmTop','txtname',0,topObj,width_obj,dHeight,ALLTRIM(people.fio)+IIF(!EMPTY(people.primtxt),' ('+ALLTRIM(people.primtxt)+') ',''),2,1,.T.      
     .AddObject('grdPers','GridMyNew')  
     WITH .grdPers      
          .Top=.Parent.txtName.Top
          .Left=.Parent.txtName.Left+.Parent.txtName.Width+5
          .Height=.Parent.Height-.Top
          .Width=.Parent.Width-.Left
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='people'
           DO addColumnToGrid WITH 'frmTop.grdPers',3
          .Column1.ControlSource='people.num'
          .Column2.ControlSource='" "+people.fio'       
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Фамилия Имя Отчество'  
          .Column1.Width=RettxtWidth(' 1234 ')      
          .Columns(.ColumnCount).Width=0          
          .Column2.Width=.Width-.column1.width-SYSMETRIC(5)-13-.ColumnCount 
          .Column1.Alignment=1        
          .Column2.Alignment=0
          .Column1.Movable=.F.         
          .colNesInf=2    
          *.procAfterRowColChange='DO cardTarifScr'  
          .procAfterRowColChange='DO changerowpers'  
         .SetAll('BOUND',.F.,'Column')  
         .Visible=.T.                        
    ENDWITH       
    DO gridSizeNew WITH 'frmTop','grdpers','shapeingrid',.F.,.T. 
    DO MyColumntxtBox WITH 'frmTop.grdPers.Column3','tbox3','',.F.,.F.,''
    .grdPers.Column3.tbox3.procForKeyPress='DO keyPresGridPers'
    DO addcontmy WITH 'frmTop','cont1',.grdPers.Left+13,.grdPers.Top+2,.grdPers.Column1.Width-3,.grdPers.HeaderHeight-3,'',"DO clickCont WITH 'frmTop','frmTop.cont1','people',1"
    .cont1.SpecialEffect=1   
    DO addcontmy WITH 'frmTop','cont2',.cont1.Left+.grdPers.Column1.Width+2,.grdPers.Top+2,.grdPers.Column2.Width-4,.grdPers.HeaderHeight-3,'',"DO clickCont WITH 'frmTop','frmTop.cont2','people',2"
    .AddObject('grdJob','GridMyNew') 
    WITH .grdJob
         .Top=.Parent.txtName.Top+.Parent.txtName.Height-1
         .Left=.Parent.txtName.Left       
         .Width=.Parent.txtName.Width
         .ScrollBars=2       	           
         .RecordSourceType=1     
         .RecordSource='datJob'
         DO addColumnToGrid WITH 'frmTop.grdJob',7
         .Column1.ControlSource="IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.name,'')"
         .Column2.ControlSource="IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.namework,'')"         
         .Column3.ControlSource='datjob.kse'
         .Column4.ControlSource='datjob.tr'
         *.Column5.ControlSource="IIF(SEEK(datjob.kat,'sprkat',1),sprkat.name,'')"         
         .Column5.ControlSource='curSprKat.name'
         .Column1.Header1.Caption='подразделение' 
         .Column2.Header1.Caption='должность'
         .Column3.Header1.Caption='объём'
         .Column4.Header1.Caption='тип'
         .Column5.Header1.Caption='персонал'
         .Column6.Header1.Caption='к'           
       
         .Column3.Width=RetTxtWidth('999.999')  
         .Column4.Width=RetTxtWidth('внеш.совм.')                               
         .Column5.Width=RetTxtWidth('wперсонал')
         .Column6.Width=RetTxtWidth('wкw')     
         
         .Columns(.ColumnCount).Width=0
         .Column1.Width=(.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width)/2
         .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-SYSMETRIC(5)-13-.ColumnCount
         
         .Column6.AddObject('checkColumn6','checkContainer')
         .Column6.checkColumn6.AddObject('checkLkv','checkMy')
         .Column6.CheckColumn6.checkLkv.Visible=.T.
         .Column6.CheckColumn6.checkLkv.Caption=''
         .Column6.CheckColumn6.checkLkv.Left=6
         .Column6.CheckColumn6.checkLkv.Top=3
         .Column6.CheckColumn6.checkLkv.BackStyle=0
         .Column6.CheckColumn6.checkLkv.ControlSource='datjob.lkv'         
         .Column6.CheckColumn6.checkLkv.procValid='DO validLkv' 
         .Column6.CheckColumn6.checkLkv.Left=(.Column6.Width-SYSMETRIC(15))/2                                                                                         
         .column6.CurrentControl='checkColumn6'
         *.procAfterRowColChange='DO changeJob'
         .procAfterRowColChange='DO changeRowJob'
         .Column1.Alignment=0
         .Column2.Alignment=0
         .Column3.Alignment=0
         .Column4.Alignment=0
         .Column5.Alignment=0
    ENDWITH         
    DO gridSizeNew WITH 'frmTop','grdJob','shapeingrid1',.F.,.T. 
    topObj=.grdJob.Top+.grdJob.Height-1
    DO adTboxAsCont WITH 'frmTop','txtJob',0,topObj,.txtName.Width,dHeight,'',2,1,.T.  
    SELECT tarfond
    labNum=0
    kvoStr=0
    labLeft=0
    txtleft=lableft+widthLab-1
    SCAN ALL
         IF nBlock=1.AND.!EMPTY(tarfond.fname)
            sayField=tarfond.rec               
            labNum=labNum+1
            namecont='lab'+LTRIM(STR(labnum))
            DO adTBoxAsCont WITH 'frmTop',namecont,labLeft,topObj,widthLab,dHeight,tarfond.rec,1,1 
            .&nameCont..Visible=.F.                                
            proc_obj=procObj
            IF !EMPTY(proc_Obj)
               &proc_Obj
               nObj=tarfond.nameObj 
               frmTop.&nObj..Visible=.F.
               frmTop.&nObj..Alignment=0
               topObj=topObj+dHeight-1                                                       
            ENDIF 
            kvoStr=kvoStr+1
            labLeft=IIF(kvoStr=2,0,txtLeft+widthtxt-1) 
            txtLeft=labLeft+widthlab-1                                   
            topObj=IIF(kvostr=2,topObj+Dheight-1,topObj)
            kvoStr=IIF(kvoStr=2,0,kvoStr) 
         ENDIF 
     ENDSCAN
     DO adTBoxAsCont WITH 'frmTop','maxCont',labLeft,topObj,widthLab,dHeight,'',1,1
     .maxCont.Visible=.F. 
     DO adTBoxAsCont WITH 'frmTop','maxCont1',labLeft,topObj,widthtxt,dHeight,'',1,0
     .maxCont1.Visible=.F. 
     DO adTboxAsCont WITH 'frmTop','txtOklad',0,topObj,width_obj,dHeight,'Образование оклада, доплаты',2,1    
     topObj=topObj+.txtOklad.Height-1
     .AddObject('grdOklad','GridMyNew')   
     WITH .grdOklad
          .Top=topObj
          .Left=.Parent.txtName.Left       
          .Width=.Parent.txtName.Width
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curSupFond'
          DO addColumnToGrid WITH 'frmTop.grdOklad',6
          .Column1.ControlSource='curSupFond.rec'
          .Column2.ControlSource='ALLTRIM(curSupFond.sPers)'         
          .Column3.ControlSource='ALLTRIM(curSupFond.sname)'
          .Column4.ControlSource='curSupFond.sMname'
          .Column5.ControlSource='curSupFond.primf'
          .Column1.Header1.Caption='тариф' 
          .Column2.Header1.Caption='%'
          .Column3.Header1.Caption='сумма'
          .Column4.Header1.Caption='на месяц'
          .Column5.Header1.Caption='примечание'
          .Column1.Width=RetTxtWidth('wза руководство структурным подразделениемw')
          .Column2.Width=RetTxtWidth('WобъёмW')
          .Column3.Width=RetTxtWidth('999999999999')
          .Column4.Width=.Column3.Width
          .Columns(.ColumnCount).Width=0
          .Column5.Width=.Width-.Column1.Width-.Column2.Width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column2.Alignment=2
          .Column3.Alignment=1
          .Column4.Alignment=1
           DO gridSizeNew WITH 'frmTop','grdOklad','shapeingrid2',.F.,.T. 
     ENDWITH    
     .grdPers.Columns(.grdPers.ColumnCount).SetFocus
     SELECT people
     GO TOP
     DO cardTarifScr
ENDWITH
SELECT people
frmTop.Show
READ EVENTS
************************************************************************************************************************
PROCEDURE quitTarif
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl   
     .Caption='Окончание работы'
     .Width=340     
     DO adLabMy WITH 'fSupl',1,'желаете закончить работу?',20,20,300,2 
     DO addcontlabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('wзакончитьw')*2-20)/2,.lab1.Top+.lab1.Height+10,RetTxtWidth('wзакончитьw'),dHeight+3,'закончить','DO endwork'
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fSupl.Release'     
     .Height=.lab1.Height+.cont1.Height+60        
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE endWork
CLEAR EVENTS
fSupl.Release
frmTop.Release
CLEAR EVENTS
************************************************************************************************************************
PROCEDURE menuRecieve
PARAMETERS parObj
IF !USED('datmenu')
   USE datmenu IN 0   
ENDIF
SELECT datmenu 
COUNT TO maxBar
DIMENSION dim_proc(maxBar)
GO TOP
CurTotHeight=35
row_pop=CurTotHeight/FONTMETRIC(1,dFontName,dFontSize)
col_pop=frmTop.menuContTop.LEFT/FONTMETRIC(6,dFontName,dFontSize)
DEFINE POPUP menutop FROM row_pop,col_pop SHORTCUT MARGIN FONT dFontName,dFontSize  COLOR SCHEME 4
numbar=0
DO WHILE !EOF()   
   numbar=numbar+1
   DEFINE BAR numbar OF menuTop PROMPT ' '+datmenu.nmenu PICTURE ALLTRIM(datmenu.mico) FONT dFontName,dFontSize
   dim_proc(numbar)=mproc 
   SKIP
ENDDO
SELECT datmenu
GO TOP
numbar=0
DO WHILE !EOF()  
   numbar=numbar+1     
   ON SELECTION BAR numbar OF menuTop DO choiceFromMenuRecieve
   SKIP
ENDDO  
ACTIVATE POPUP menuTop
**************************************************************************************************************************
PROCEDURE choiceFromMenuRecieve
IF !EMPTY(dim_proc(BAR()))
   &dim_proc(BAR())
ENDIF 
**************************************************************************************************************************
*                                  Смена даты тарификации
**************************************************************************************************************************
PROCEDURE choiceDateTar
IF USED('curDatShtat')
   SELECT curDatShtat
   USE
ENDIF
oldNumPeop=people.num
SELECT * FROM datshtat WHERE lUse INTO CURSOR curDatShtat READWRITE
SELECT curDatShtat
ALTER TABLE curDatShtat ADD COLUMN nameSupl C(70)
INDEX ON real TAG T1 DESCENDING
REPLACE namesupl WITH IIF(real,DTOC(DATE()),DTOC(dTarif))+' '+ALLTRIM(fullName) ALL
LOCATE FOR ALLTRIM(pathtarif)=pathTarSupl

varDtarSupl=varDtar
varBaseStSupl=varBaseSt
strDate=namesupl
*pathTarSupl=ALLTRIM(curDatShtat.pathtarif)
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl    
     .Caption='Выбор даты тарификации'
     DO addshape WITH 'fSupl',1,20,20,150,380,8 
     DO adtBoxAsCont WITH 'fSupl','contDate',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('wвыберите дата тарификацииw'),dHeight,'выберите дату тарификации',2,1     
     DO addcombomy WITH 'fSupl',1,.contDate.Left+.contdate.Width-1,.contDate.Top,dHeight,300,.T.,'strDate','ALLTRIM(curDatShtat.namesupl)',6,'','DO validDtarif',.F.,.T.
     .comboBox1.DisplayCount=15
     .Shape1.Width=.contDate.Width+.comboBox1.Width+40
     .Shape1.height=.comboBox1.Height+40
     *-----------------------------Кнопка применить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприменитьw')*2)-20)/2,;
       .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприменитьw'),dHeight+5,'Применить','DO procChangeTarif'
     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','fSupl.Release','Возврат'
     .Width=.Shape1.Width+40    
          
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.cont1.Height+60     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validDtarif
varDtarSupl=curDatShtat.dtarif
varBaseStSupl=curDatShtat.baseSt
varNmzp=curdatShtat.nmzp
*************************************************************************************************************************
PROCEDURE procChangeTarif
fSupl.Visible=.F.
SELECT datset
REPLACE pathlast WITH curDatShtat.pathtarif

varDtar=varDtarSupl
IF curDatShtat.real
   varDtar=DATE()
ENDIF  
varBaseSt=varBaseStSupl
varNmzp=curDatShtat.nmzp
pathtarif=pathmain+'\'+ALLTRIM(curdatshtat.pathtarif)+'\'+';'+pathmain+';'+pathsupl
pathTarSupl=ALLTRIM(curDatShtat.pathtarif)
pathcur=pathmain+'\'+ALLTRIM(curdatshtat.pathtarif)+'\' && путь к каталогу тарификации с которой сейчас работают
tarDateSay='тарификация на '+DTOC(varDTar)+IIF(curdatshtat.real,' текущая ','')+' (изменить - двойной щелчок мыши)  базовая ставка - '+LTRIM(STR(varBaseSt,8,2))+'  мин.з.п - '+LTRIM(STR(varNmzp,8,2))
frmTop.contDateTarif.ContLabel.Caption=tarDateSay
SELECT people
USE
SELECT rasp
USE
SELECT sprpodr
USE
SELECT datjob
USE
SELECT tarfond
USE
SELECT boss
USE
SELECT sprdolj
USE
SELECT sprtime
USE
SELECT curSupFond
SET FILTER TO 
DELETE ALL
IF USED('datagrup')
   SELECT datagrup
   USE 
ENDIF
IF USED('curReadTarfond')   
   SELECT curReadTarFond
   USE
ENDIF 

SET PATH TO &pathtarif
USE people ORDER 1 IN 0
SELECT people
SEEK oldNumPeop
USE rasp ORDER 1 IN 0
USE datjob ORDER 4 IN 0
USE sprpodr ORDER 1 IN 0
USE tarfond ORDER 1 IN 0
USE boss IN 0
USE sprtime ORDER 1 IN 0
USE sprdolj IN 0 ORDER 1
SELECT * FROM sprdolj INTO CURSOR curSprDolj READWRITE ORDER BY name
SELECT * FROM sprpodr INTO CURSOR curSprPodr READWRITE ORDER BY name
SELECT curSupFond
DELETE ALL
APPEND FROM tarfond for nBlock=2
*SELECT tarfond
*SCAN ALL
*     IF nBlock=2
*        SCATTER TO dimSup
*        SELECT curSupFond
*        APPEND BLANK
*        GATHER FROM dimSup       
*     ENDIF
*     SELECT tarfond
*ENDSCAN
*SELECT curSupFond
curRow=0
curRowJob=0
SELECT datJob
SET RELATION TO tr INTO curSprType, kat INTO curSprKat ADDITIVE
WITH frmTop
     WITH .grdPers          
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='people'
           DO addColumnToGrid WITH 'frmTop.grdPers',3
          .Column1.ControlSource='people.num'
          .Column2.ControlSource='" "+people.fio'       
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Фамилия Имя Отчество'  
          .Column1.Width=RettxtWidth(' 1234 ')      
          .Columns(.ColumnCount).Width=0          
          .Column2.Width=.Width-.column1.width-SYSMETRIC(5)-13-.ColumnCount 
          .Column1.Alignment=1        
          .Column2.Alignment=0
          .Column1.Movable=.F.         
          .colNesInf=2    
          .procAfterRowColChange='DO cardTarifScr'  
         .SetAll('BOUND',.F.,'Column')  
         .Visible=.T.             
     ENDWITH
     DO gridSizeNew WITH 'frmTop','grdpers',.F.,.F.,.T.
     WITH .grdJob      
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='datJob'
          DO addColumnToGrid WITH 'frmTop.grdJob',7
          .Column1.ControlSource="IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.name,'')"
          .Column2.ControlSource="IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.namework,'')"         
          .Column3.ControlSource='datjob.kse'
          .Column4.ControlSource='datjob.tr'
          *.Column5.ControlSource="IIF(SEEK(datjob.kat,'sprkat',1),sprkat.name,'')"   
          .Column5.ControlSource='curSprKat.name'             
          .Column6.ControlSource='datjob.lkv'
          .Column6.AddObject('checkColumn6','checkContainer')
          .Column6.checkColumn6.AddObject('checkLkv','checkMy')
          .Column6.CheckColumn6.checkLkv.ControlSource='datjob.lkv' 
          .Column6.CheckColumn6.checkLkv.procValid='DO validLkv' 
          *.procAfterRowColChange='DO changeJob'
          .procAfterRowColChange='DO changeRowJob'
    ENDWITH        
    DO gridSizeNew WITH 'frmTop','grdJob',.F.,.F.,.T.   
    .grdPers.Columns(.grdPers.ColumnCount).SetFocus  
    .Refresh    
ENDWITH
fSupl.Visible=.F.
************************************************************************************************************************
PROCEDURE keyPresGridPers
DO CASE
   CASE LASTKEY()=6 &&ctrl+F
        DO formForSearsh          
   CASE LASTKEY()=147 &&ctrl+Del 
        DO formDeletePeople  
ENDCASE
**************************************************************************************************************************
PROCEDURE changerowpers
IF curRow#frmTop.grdPers.curRec
   DO cardTarifScr
   curRow=frmTop.grdPers.curRec 
ENDIF
**************************************************************************************************************************
*   Выввод на экран информации по сотруднику
***************************************************************************************************************************
PROCEDURE cardTarifScr
SELECT people
WITH frmTop
     SELECT people
     peopRec=RECNO()
     nameSay=LTRIM(STR(people.num))+'  '+ALLTRIM(people.fio)+IIF(!EMPTY(people.primtxt),' ('+ALLTRIM(people.primtxt)+') ','')
     .txtName.ControlSource='nameSay'
     DO actualStajToday WITH 'people','date_in','DATE()'
     *DO actualstaj  
     SELECT datjob
     IF EMPTY(fltJob)
        SET FILTER TO datjob.kodpeop=people.num.AND.EMPTY(dateout)
     ELSE 
        SET FILTER TO datjob.kodpeop=people.num.AND.&fltJob     
     ENDIF    
     COUNT TO maxJob
     GO TOP     
     WITH .grdJob
          .RecordSourceType=1     
          .RecordSource='datjob'
          .Height=.headerHeight+.RowHeight*(maxJob+1)                          
         * .ScrollBars=2
          .ColumnCount=7          
          .RecordSource='datjob'
          .Column1.ControlSource="IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.name,'')"
          .Column2.ControlSource="IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.namework,'')"
          .Column3.ControlSource='datjob.kse'
          .Column4.ControlSource='curSprType.name'
          *.Column5.ControlSource="IIF(SEEK(datjob.kat,'sprkat',1),sprkat.name,'')"
          .Column5.ControlSource='curSprKat.name'
          .Column6.ControlSource='datJob.lkv'
                   
          .Column1.Header1.Caption='подразделение' 
          .Column2.Header1.Caption='должность'
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'
          .Column5.Header1.Caption='персонал'
          .Column6.Header1.Caption='к'           
          
          .Column3.Width=RetTxtWidth('999.999')  
          .Column4.Width=RetTxtWidth('внеш.совм.')                               
          .Column5.Width=RetTxtWidth('wперсонал')
          .Column6.Width=RetTxtWidth('wкw')     
         
          .Columns(.ColumnCount).Width=0
          .Column1.Width=(.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width)/2
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-SYSMETRIC(5)-13-.ColumnCount
         
         * .Column6.AddObject('checkColumn6','checkContainer')
         * .Column6.checkColumn6.AddObject('checkMy','checkBox')
          .Column6.CheckColumn6.checkLkv.Visible=.T.
          .Column6.CheckColumn6.checkLkv.Caption=''
          .Column6.CheckColumn6.checkLkv.Left=6
          .Column6.CheckColumn6.checkLkv.Top=3
          .Column6.CheckColumn6.checkLkv.BackStyle=0
          .Column6.CheckColumn6.checkLkv.ControlSource='datjob.lkv'         
          .Column6.CheckColumn6.checkLkv.Left=(.Column6.Width-SYSMETRIC(15))/2        
          .column6.CurrentControl='checkColumn6'
          *.procAfterRowColChange='DO changeJob'
           .procAfterRowColChange='DO changeRowJob'
          .Column1.Alignment=0
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0
          .Column5.Alignment=0
          .SetAll('Enabled',.F.,'Column') 
          .Column6.Enabled=.T.   
          .Columns(.ColumnCount).Enabled=.T.   
          .Column6.Sparse=.F. 
     ENDWITH    
     DO gridSizeNew WITH 'frmTop','grdJob',.F.,.T.,.T. 
     SELECT datjob
     *DO gridSizeNew WITH 'frmTop','grdJob',.F.,.F.,.T.   
     .ShapeInGrid1.Height=.grdJob.Height
     .grdJob.Columns(.grdJob.ColumnCount).SetFocus
     .grdPers.Columns(.grdPers.ColumnCount).SetFocus     
     .Refresh
ENDWITH
**************************************************************************************************************************
PROCEDURE changeRowJob
IF curRowJob#frmTop.grdJob.curRec
   DO changeJob
   curRowJob=frmTop.grdJob.curRec 
ENDIF
**************************************************************************************************************************
PROCEDURE changeJob
WITH frmTop    
     *DO perstajone 
     topObj=.grdJob.Top+.grdJob.Height-1
     .txtJob.Top=topObj
     nameJobPeop=IIF(SEEK(datjob.kp,'sprpodr',1),ALLTRIM(sprpodr.name)+'  ','')+IIF(SEEK(datjob.kd,'sprdolj',1),ALLTRIM(sprdolj.namework)+'  ','')+IIF(datJob.kse#0,ALLTRIM(STR(datjob.kse,5,2))+' ','')+ALLTRIM(curSprType.name)
     nameJobKval=IIF(SEEK(datJob.kv,'sprkval',1),sprkval.name,'')
     namevtime=IIF(SEEK(datjob.vtime,'sprtime',1),sprtime.name,'')
     cLogMol=IIF(datjob.lmol,'да','')
      strperstaj=IIF(!EMPTY(datjob.per_date),DTOC(datjob.per_date)+' -  '+LTRIM(STR(datjob.st_per))+'%','')
     .txtJob.ControlSource='nameJobPeop'
     topObj=.txtJob.Top+.txtJob.Height-1 
     IF datShtat.real
        ON ERROR DO ersup
        *DO repNadJob
        ON ERROR 
     ENDIF  
     SELECT tarfond   
     labNum=0
     kvoStr=0
     labLeft=0
     txtLeft=labLeft+widthLab-1
     kvoSay=0
     .maxCont.Visible=.F.
     .maxCont1.Visible=.F.
     SCAN ALL
         IF nBlock=1.AND.!EMPTY(tarfond.fname)         
            sayField=tarfond.fname
            labNum=labNum+1
            namecont='lab'+LTRIM(STR(labnum))
            .&nameCont..Visible=IIF(!EMPTY(&sayField),.T.,.F.) 
            nObj=tarfond.nameObj   
            frmTop.&nObj..Visible=.F.                               
            IF !EMPTY(&sayField)  
               .&nameCont..Visible=.T.                     
               nObj=tarfond.nameObj 
               frmTop.&nameCont..Left=labLeft
               frmTop.&nameCont..Width=widthLab
               frmTop.&nameCont..Top=topObj                                 
               lableft=labLeft+widthLab-1                
               frmTop.&nObj..Left=txtLeft
               frmTop.&nObj..Width=widthtxt
               frmTop.&nObj..Top=topObj
               frmTop.&nObj..Visible=.T.                     
               kvoStr=kvoStr+1                                                                       
               labLeft=IIF(kvoStr=2,0,txtLeft+widthtxt-1) 
               txtLeft=labLeft+widthlab-1                                   
               topObj=IIF(kvostr=2,topObj+Dheight-1,topObj)
               kvoStr=IIF(kvoStr=2,0,kvoStr)  
               kvoSay=kVoSay+1
              
            ENDIF                
         ENDIF 
     ENDSCAN 
     IF kvoSay#0.AND.MOD(kvoSay,2)=1
        .maxCont.Visible=.T.
        .maxCont.Top=topObj
        .maxCont.Left=labLeft  
        .maxCont1.Visible=.T.
        .maxCont1.Left=txtLeft
        .maxCont1.Top=topObj  
         topObj=topObj+dHeight-1                         
     ENDIF  
     .txtOklad.Top=topObj
     topObj=topObj+.txtOklad.Height-1       
     .grdOklad.Top=topObj
     .grdOklad.Height=.Height-.grdOklad.Top 
     .ShapeInGrid2.Top=.grdOklad.Top
     .ShapeInGrid2.Height=.grdOklad.Height 
ENDWITH 
SELECT curSupFond
SET FILTER TO 
REPLACE spers WITH '',sname WITH '' all
GO TOP
SCAN ALL
     reppl1=fpers
     reppl2=fname
     reppl3=ALLTRIM(sayokl)
     reppl4=ALLTRIM(sayoklm)
     IF !EMPTY(reppl1)
        IF logspers
            REPLACE spers WITH IIF(!EMPTY(&reppl1),ALLTRIM(STR(&reppl1,6,2))+IIF(logp,'%',''),'')          
        ELSE 
           REPLACE spers WITH IIF(!EMPTY(&reppl1),ALLTRIM(STR(&reppl1))+IIF(logp,'%',''),'')
        ENDIF            
     ENDIF      
     IF !EMPTY(reppl2)   
        IF EMPTY(&reppl2)
           REPLACE sname WITH ''
        ELSE 
           IF !EMPTY(reppl3)
              REPLACE sname WITH &reppl3 
           ELSE 
              REPLACE sname WITH LTRIM(STR(&reppl2)) 
           ENDIF 
        ENDIF 
     ENDIF 
     IF !EMPTY(reppl4)
        REPLACE sMname WITH &reppl4       
     ENDIF
ENDSCAN
SET FILTER TO !EMPTY(sPers).OR.!EMPTY(sName)
GO TOP
SELECT datjob
frmTop.Refresh
**************************************************************************************************************************
PROCEDURE validLkv
SELECT datjob
REPLACE kv WITH IIF(lkv,people.kval,0)
DO dopl_sum WITH .T.
DO changeJob 
SELECT datjob
GO curRowJob
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus
**************************************************************************************************************************
*                       Подстановка набавок и доплат для тарификации
**************************************************************************************************************************
PROCEDURE repNadJob 
SELECT datjob
REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in,pkont WITH IIF(tr=1,people.pkont,0)
DO CASE 
   CASE datjob.lkv.AND.people.kval#0
        REPLACE kv WITH people.kval,nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' №'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' от ' +DTOC(people.dkval),''),;
                pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,0)
   OTHERWISE
        REPLACE kv WITH 0,nPrik WITH '',pkat WITH 0   
ENDCASE 
DO CASE
   CASE kv=0
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf,namekf)
   CASE kv=1
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf3,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf3,namekf)
   CASE kv=2
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf2,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf2,namekf)
   CASE kv=3
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf1,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf1,namekf)
ENDCASE
SELECT rasp
nOldRaspOrd=SYS(21)
SET ORDER TO 2
SEEK STR(datjob.kp,3)+STR(datjob.kd,3)
REPLACE datjob.pkf WITH rasp.pkf
SELECT tarfond
GO TOP
SCAN ALL
     IF !EMPTY(plrep).AND.ltar
        IF !EMPTY(procrepnad)
           procrep=ALLTRIM(procrepnad)
           DO &procrep 
        ELSE 
           repjob=ALLTRIM(plrep)
           repjob1='rasp.'+ALLTRIM(plrep)
           SELECT datjob 
           REPLACE &repjob WITH &repjob1
        ENDIF    
     ENDIF
     IF ALLTRIM(LOWER(tarfond.plrep))='pkat'.AND.rasp.pkat#0.AND.datjob.kv=0
        SELECT datjob
        REPLACE pkat WITH 5       
     ENDIF
     SELECT tarfond
ENDSCAN
SELECT rasp
SET ORDER TO &nOldRaspOrd
SELECT datjob

**************************************************************************************************************************
PROCEDURE repnadvto
SELECT datjob
DO CASE 
   CASE kv=0
        REPLACE pvto WITH rasp.pvto4
   CASE kv=1
        REPLACE pvto WITH rasp.pvto1
   CASE kv=2
        REPLACE pvto WITH rasp.pvto2
   CASE kv=3
        REPLACE pvto WITH rasp.pvto3
ENDCASE 
SELECT tarfond

**************************************************************************************************************************
*                  Ввод нового сотрудника
**************************************************************************************************************************
PROCEDURE newPeopTarif
CREATE CURSOR curDolPodr (kd N(3),namedolj C(100),kat N(2)) 
SELECT curDolPodr
*INDEX ON namedolj TAG T1
SELECT people
log_ord=SYS(21)
SET ORDER TO 4
SET DELETED OFF
GO BOTTOM 
newNid=nid+1
SET DELETED ON
SET ORDER TO 1
GO BOTTOM 
new_num=num+1

STORE '' TO new_fio,str_type,str_podr,str_dolj,new_primtxt,strNewKat
STORE 0 TO new_podr,newTabn,new_dolj,new_tr,new_kat
date_innew=CTOD('  .  .    ')
staj_innew=''
new_kse=0.00
logRead=.T.
logVac=.F.
SET ORDER TO &log_ord
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl    
     .Caption='Ввод нового сотрудника'    
     DO addShape WITH 'fSupl',1,10,10,dHeight,300,8     
     .logExit=.T.      
     DO adTBoxAsCont WITH 'fsupl','txtName',.Shape1.Left+10,.Shape1.Top+20,RetTxtWidth(' Фамилия Имя Отчество '),dHeight,'Фамилия Имя Отчество',1,1
     DO addtxtboxmy WITH 'fSupl',1,.txtName.Left+.txtName.Width-1,.txtName.Top,450,.F.,'new_fio',0,"DO unosimbol WITH 'new_fio'"     
     
     DO adTBoxAsCont WITH 'fsupl','txtTabn',.txtName.Left,.txtName.Top+.txtName.Height-1,.txtName.Width,dHeight,'Табельный номер',1,1
     DO addtxtboxmy WITH 'fSupl',11,.txtBox1.Left,.txtTabn.Top,.txtBox1.Width,.F.,'newTabn',0,.F.,'Z'    
         
     DO adTboxAsCont WITH 'fSupl','txtPodr',.txtname.Left,.txtTabn.Top+.txtTabn.Height-1,.txtName.Width,dHeight,'Подразделение',1,1
     DO addComboMy WITH 'fSupl',1,.txtBox1.Left,.txtPodr.Top,dheight,.txtBox1.Width,.T.,'str_podr','cursprpodr.name',6,.F.,'DO validPodrPeop',.F.,.T.      
                 
     DO adTBoxAsCont WITH 'fSupl','txtDolj',.txtName.Left,.txtPodr.Top+.txtPodr.Height-1,.txtName.Width,dHeight,'Должность',1,1
     DO addComboMy WITH 'fSupl',2,.txtBox1.Left,.txtDolj.Top,dheight,.txtBox1.Width,.T.,'str_dolj','curDolPodr.namedolj',6,.F.,'DO validDoljPeop',.F.,.T.           
     
     DO adTBoxAsCont WITH 'fSupl','txtPers',.txtName.Left,.txtDolj.Top+.txtDolj.Height-1,.txtName.Width,dHeight,'Персонал',1,1
     DO adTBoxAsCont WITH 'fSupl','txtPers1',.txtBox1.Left,.txtPers.Top,.txtBox1.Width,dHeight,strNewkat,0,0     
      
     DO adTBoxAsCont WITH 'fSupl','txtKse',.txtName.Left,.txtPers.Top+.txtPers.Height-1,.txtName.Width,dHeight,'объём',1,1
     DO addSpinnerMy WITH 'fSupl','spinKse',.txtKse.Left+.txtKse.Width-1,.txtKse.Top,dheight,RetTxtWidth('9999999999'),'new_kse',0.25,.F.,0,1.5
                  
     DO adTBoxAsCont WITH 'fSupl','txtType',.spinKse.Left+.spinKse.Width-1,.txtKse.Top,RetTxtWidth('wтип работыw'),dHeight,'тип',2,1                                             
     DO addComboMy WITH 'fSupl',3,.txtType.Left+.txtType.Width-1,.txtType.Top,dheight,.txtPodr.Width+.comboBox1.Width-.txtKse.Width-.spinKse.Width-.txtType.Width+2,;
         .T.,'str_type','sprType.name',6,.F.,'new_tr=sprType.kod',.F.,.T. 
     
     DO adTboxAsCont WITH 'fSupl','txtDateIn',.txtName.Left,.txtKse.Top+.txtKse.Height-1,.txtName.Width,dHeight,'Дата приема',1,1    
     DO addtxtboxmy WITH 'fSupl',3,.txtBox1.Left,.txtDateIn.Top,.txtBox1.Width,.F.,'date_innew',0,.F.    
           
     DO adTboxAsCont WITH 'fSupl','txtStajIn',.txtName.Left,.txtDateIn.Top+.txtDateIn.Height-1,.txtName.Width,dHeight,'стаж на дату приема',1,1         
     DO addtxtboxmy WITH 'fSupl',4,.txtBox1.Left,.txtStajIn.Top,.txtBox1.Width,.F.,'staj_innew',0,.F.    
     .txtBox4.InputMask='99-99-99'
     
     DO adTboxAsCont WITH 'fSupl','txtPrim',.txtName.Left,.txtStajIn.Top+.txtStajIn.Height-1,.txtName.Width,dHeight,'Примечание',1,1     
     DO addtxtboxmy WITH 'fSupl',2,.txtBox1.Left,.txtPrim.Top,.txtBox1.Width,.F.,'new_primtxt',0 
       
     DO adCheckBox WITH 'fSupl','check1','Вакантная',.txtPrim.Top+.txtPrim.Height+20,.Shape1.Left,150,dHeight,'logVac',0,.F.,'DO validCheckVac'                                                    
     DO adCheckBox WITH 'fSupl','check2','После записи перейти к дальнейшему редактированию',.check1.Top+.check1.Height+10,.Shape1.Left,150,dHeight,'logRead',0      
       
     .Shape1.Height=.txtBox1.height*8+.check1.Height*2+70
     .Shape1.Width=.txtname.Width+.txtBox1.Width-1+20 
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     .check2.Left=.Shape1.Left+(.Shape1.Width-.check2.Width)/2                
    
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WWЗаписатьWW')*2-20)/2,.Shape1.Top+.Shape1.Height+10,;
     RetTxtWidth('WWЗаписатьWW'),dHeight+3,'Записать','DO writeNewPeopTarif WITH .T.'
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','DO writeNewPeopTarif WITH .F.'    
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+30+.cont1.Height       
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE validCheckVac
new_fio=IIF(logvac,'Вакантная','')
*fSupl.txtBox1.ControlSource='newFio'
fSupl.txtBox1.Enabled=IIF(logVac,.F.,.T.)
fSupl.txtBox1.Refresh
***********************************************************************************************************************
PROCEDURE validPodrPeop
new_podr=curSprPodr.kod
SELECT curDolPodr
DELETE ALL
SELECT rasp
SET FILTER TO kp=new_podr
SCAN ALL    
     SELECT curDolPodr
     APPEND BLANK
     REPLACE kd WITH rasp.kd,kat WITH rasp.kat,namedolj WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,'') 
     SELECT rasp
ENDSCAN
fSupl.ComboBox2.RowSource='curDolPodR.namedolj'
fSupl.ComboBox2.DisplayCount=RECCOUNT('curDolPodr')
fSupl.ComboBox2.RowSourceType=6
fsupl.comboBox2.ProcForValid='DO validDoljPeop'
KEYBOARD '{TAB}'
**************************************************************************************************************************
PROCEDURE validDoljPeop
new_dolj=curDolPodr.kd
new_kat=curDolPodr.kat
strNewKat=IIF(SEEK(new_kat,'sprkat',1),ALLTRIM(sprkat.name),'')
fSupl.txtPers1.ControlSource='strNewkat'
fSupl.Refresh
KEYBOARD '{TAB}'
************************************************************************************************************************
*
************************************************************************************************************************
PROCEDURE writeNewPeopTarif
PARAMETERS parLog
IF !parLog
   SELECT rasp
   SET FILTER TO
   SELECT people   
   GO peoprec
   fSupl.Release
ELSE 
   IF EMPTY(new_fio)
      RETURN
   ENDIF 
   SELECT people
   APPEND BLANK
   newrec=RECNO()
   REPLACE num WITH new_num,nid WITH newNid,fio WITH new_fio,tabn WITH newTabn,primtxt WITH new_primtxt,vac WITH logVac,date_in WITH date_innew,staj_in WITH staj_innew
   IF new_podr#0.AND.new_dolj#0
      SELECT datJob
      SET FILTER TO 
      ordJobOld=SYS(21)
      SET ORDER TO 7
      SET DELETED OFF
      GO BOTTOM 
      newNidJob=nid+1
      SET DELETED ON
      SET ORDER TO &ordJobOld   
      APPEND BLANK
      REPLACE kodpeop WITH new_num, kp WITH new_podr,kd WITH new_dolj,kse WITH new_kse,tr WITH new_tr,kat WITH new_kat,tabn WITH newTabn,;
              date_in WITH date_innew,staj_in WITH staj_innew,nidpeop WITH people.nid,nid WITH newNidJob 
   ENDIF 
   SELECT people       
   frmTop.Refresh 
   frmTop.grdPers.Columns(frmTop.grdPers.columnCount).SetFocus
   fSupl.Visible=.F.
   fSupl.Release
   IF logRead
      DO readPeople
   ENDIF
ENDIF 
************************************************************************************************************************
*                                     Редактирование сведений по сотруднику                  
************************************************************************************************************************
PROCEDURE readPeople
SELECT datJob
datJobRec=RECNO()
IF !USED('curSprKoef')
   SELECT * FROM sprkoef INTO CURSOR curSprKoef READWRITE
   ALTER TABLE curSprKoef ADD COLUMN nkf C(20)
ENDIF    
SELECT curSprKoef
REPLACE nkf WITH STR(kod,2)+' - '+STR(name,5,2) ALL
INDEX ON kod TAG T1
SET ORDER TO 1
IF !USED('curReadTarfond')
   SELECT * FROM tarfond  WHERE logR INTO CURSOR curReadTarFond READWRITE
   ALTER TABLE curReadTarfond ADD COLUMN  fieldValue C(100)
   SELECT curReadTarFond
   INDEX ON num TAG T1
ELSE 
   SELECT curReadTarfond 
   REPLACE fieldValue WITH '' ALL 
ENDIF    
SELECT curReadTarfond
SCAN ALL
     repValue=ALLTRIM(curReadTarFond.firead)
     IF !EMPTY(&repValue)
        DO CASE      
            CASE logKf                 
                 REPLACE fieldValue WITH IIF(datjob.kf#0,'  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,2),'')     
            CASE !EMPTY(nbase)
                 selBase=ALLTRIM(nbase)
                 SELECT &selBase                                
                 LOCATE FOR kod=&repValue
                 SELECT curReadTarfond
                 REPLACE fieldValue WITH '  '+&selBase..name         
            CASE har='c'                               
                 REPLACE fieldValue WITH '  '+&repValue             
                      
            CASE har='d'
                 REPLACE fieldValue WITH '  '+DTOC(&repValue)
            CASE har='n'      
                 REPLACE fieldValue WITH '  '+LTRIM(STR(&repValue))
            CASE har='N'      
                 REPLACE fieldValue WITH '  '+LTRIM(STR(&repValue,6,2)) 
            CASE har='l'
                 REPLACE fieldValue WITH '  '+'да'           
            OTHERWISE 
                 repSayOkl=ALLTRIM(curReadTarFond.sayOkl)
                 REPLACE fieldValue WITH '  '+&repSayOkl    
           
        ENDCASE
     ENDIF 
ENDSCAN
fSupl=CREATEOBJECT('FORMSUPL')
WITH fsupl
     .procExit='DO exitFromReadPeop'
     .Caption='Редактирование'
     .logExit=.T.
     DO addListBoxMy WITH 'fSupl',1,20,20,600,800
     .AddObject('lstLine','LINE')
     WITH .listBox1
          .RowSource='curReadTarfond.rec,fieldValue'           
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='300,500' 
          .logExit=.T.
          .procForClick='DO clickListBoxRead'         
          .procForKeyPress='DO KeyPressListBoxRead'          
          .procForRightClick='DO procDelTarifPeop'                    
     ENDWITH
     DO adtBoxAsCont WITH 'fSupl','cont1',.listBox1.Left,.listBox1.Top+.listBox1.Height-1,174,dHeight,'',0,0   
     DO adTboxNew WITH 'fSupl','boxread',.cont1.Top,.cont1.Left+.Cont1.Width-1,.listBox1.Width-.cont1.Width+1,dHeight,'',.F.,.T.,0,.F.,'DO validBoxRead',.F.,.F.,'DO keyPressBoxRead'   
     .boxRead.procForLostFocus='DO lostBoxread' 
     maskBoxRead=.boxRead.InputMask 
     varTxtDel='Удаление - правая клавиша мыши или клавиша Del'    
     DO adTboxAsCont WITH 'fSupl','txtDel',.cont1.Left,.cont1.Top,.listBox1.Width,dHeight,vartxtDel,2,1 
     
     DO addListBoxMy WITH 'fSupl',2,20,20,.listBox1.Height+.cont1.Height-1,.ListBox1.Width
     WITH .listBox2
          .Visible=.F.
          .ColumnCount=2
          .ColumnWidths='20,780'
          .RowSourceType=6
          .ColumnLines=.F.
     ENDWITH
     *----кнопки для выбора объекта редактирования и возврат --------------------------------------------------------------------------
     DO addContLabel WITH 'fSupl','butExit',.listBox1.Left+(.listBox1.Width-RetTxtWidth('WвозвратW'))/2,.cont1.Top+.cont1.Height+10,;
       RetTxtWidth('WвыходW'),dHeight+3,'возврат','DO exitFromReadPeop'       
       
     *------------------------ кнопки удвлить,записать и возврат
     DO addContLabel WITH 'fSupl','butDel',.listBox1.Left+(.listBox1.Width-RetTxtWidth('WЗаписатьW')*3-30)/2,.ButExit.Top,;
        RetTxtWidth('WудалитьW'),dHeight+3,'удалить','DO procDelTarif' 
        
     DO addContLabel WITH 'fSupl','butSave',.butDel.Left+.ButDel.Width+10,.butExit.Top,;
       .butDel.Width,dHeight+3,'записать','DO procObjReadEnabled WITH .F.,.T.'             
     DO addContLabel WITH 'fSupl','butReturn',.butSave.Left+.butSave.Width+20,.butExit.Top,;
       .butDel.Width,dHeight+3,'возврат','DO procObjReadEnabled WITH .F.,.T.'    
     .ButSave.Left=.listBox1.Left+(.listBox1.Width-.butSave.Width-.butReturn.Width-10)/2
     .butReturn.Left=.butSave.Left+.butSave.Width+10
     .butDel.Left=.butSave.Left                
       
     .butDel.Visible=.F.
     .butSave.Visible=.F.
     .butReturn.Visible=.F.
     
     .lstLine.Top=.listBox1.Top
     .lstLine.Height=.listBox1.Height
     .lstLine.Left=.ListBox1.Left+300+3
     .lstLine.Width=0    
     .lstLine.Visible=.T.  
     .Width=.listBox1.Width+40
     .Height=.listBox1.height+.cont1.Height+.ButExit.Height+50  
     .Autocenter=.T.
ENDWITH
*DO pasteImage WITH 'fSupl'
fSupl.Show
**************************************************************************************************************************
PROCEDURE exitFromReadPeop
SELECT people
IF people.vac
   REPLACE fio WITH 'Вакантная',staj_in WITH '',vac WITH .T. 
   DO unosimbol WITH 'people.fio',.T.
ELSE 
 * REPLACE date_in WITH datjob.date_in,staj_in WITH datJob.staj_in   
ENDIF
DO dopl_sum 
frmTop.Refresh
ON ERROR DO erSup
SELECT datJob
*REPLACE vac WITH people.vac FOR kodpeop=people.num
GO datJobRec
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus   
DO changeJob 
fSupl.Release
ON ERROR 
**************************************************************************************************************************
PROCEDURE clickListBoxRead
IF EMPTY(curReadTarFond.procRead)
   RETURN 
ENDIF
procForRead=curReadTarfond.procRead
&procForRead
**************************************************************************************************************************
PROCEDURE keyPressListBoxRead
DO CASE  
   CASE LASTKEY()=13   
        DO clickListBoxRead       
   CASE LASTKEY()=7
        DO procDelTarifPeop   
   CASE LASTKEY()=27
        DO exitFromReadPeop   
ENDCASE
**************************************************************************************************************************
PROCEDURE procVacRead
SELECT people
REPLACE vac WITH IIF(people.vac,.F.,.T.) 
SELECT curReadTarFond
REPLACE fieldValue WITH '  '+ IIF(people.vac,'да','')   
**************************************************************************************************************************
*     Редактирование процентов и сумм
**************************************************************************************************************************
PROCEDURE procTotRead
PARAMETERS parRead,parvac
IF !EMPTY(parvac)
   IF &parvac
      RETURN
   ENDIF    
ENDIF
DO procObjReadEnabled WITH 2      
WITH fSupl
     ctrlCont1=curReadTarfond.rec
     .Cont1.ControlSource='ctrlCont1'     
     .boxRead.ControlSource=parRead
     .boxRead.InputMask= maskBoxRead 
     .boxRead.Enabled=.T.
     .boxRead.Alignment=0
     .boxRead.SetFocus    
ENDWITH
**************************************************************************************************************************
*     Редактирование процентов и сумм
**************************************************************************************************************************
PROCEDURE procTarRead1
PARAMETERS parRead,parvac
IF datjob.kodpeop=0
   RETURN 
ENDIF
IF !EMPTY(parvac)
   IF &parvac
      RETURN
   ENDIF    
ENDIF
DO procObjReadEnabled WITH 2      
WITH fSupl
     ctrlCont1=curReadTarfond.rec
     .Cont1.ControlSource='ctrlCont1'     
     .boxRead.ControlSource=parRead
     .boxRead.InputMask= maskBoxRead 
     .boxRead.Enabled=.T.
     .boxRead.Alignment=1
     .boxRead.SetFocus    
ENDWITH
**************************************************************************************************************************
*     Редактирование процентов и сумм
**************************************************************************************************************************
PROCEDURE procTarRead
PARAMETERS parRead,parmask
IF datjob.kodpeop=0
   RETURN 
ENDIF
*IF !EMPTY(parMask)
*   IF &parmask
*      RETURN
*   ENDIF    
*ENDIF
DO procObjReadEnabled WITH 2      
WITH fSupl
     ctrlCont1=curReadTarfond.rec
     .Cont1.ControlSource='ctrlCont1'     
     .boxRead.ControlSource=parRead
     IF !EMPTY(parMask)
        .boxRead.InputMask=parMask 
     ELSE 
        .boxRead.InputMask= maskBoxRead 
     ENDIF 
     .boxRead.Enabled=.T.
     .boxRead.Alignment=0
     .boxRead.SetFocus    
ENDWITH
**************************************************************************************************************************
PROCEDURE validBoxRead
SELECT curReadTarFond
DO CASE
   CASE LOWER(ALLTRIM(curReadTarFond.fname))='datjob.namekf'
        SELECT datjob
        REPLACE kf WITH 0
        SELECT curReadTarFond
        REPLACE fieldValue WITH '  '+IIF(fSupl.BoxRead.Value#0,LTRIM(STR(fSupl.BoxRead.Value,6,2)),'')
        rrec=RECNO()
        LOCATE FOR logKf
        REPLACE fieldValue WITH ''
        GO rrec 
   CASE curReadTarFond.har='c'
        REPLACE fieldValue WITH '  '+fSupl.BoxRead.Value             
   CASE curReadTarfond.har='n'     
        REPLACE fieldValue WITH '  '+IIF(fSupl.BoxRead.Value#0,LTRIM(STR(fSupl.BoxRead.Value)),'')
   CASE curReadTarfond.har='N'        
        REPLACE fieldValue WITH '  '+IIF(fSupl.BoxRead.Value#0,LTRIM(STR(fSupl.BoxRead.Value,6,2)),'')
   CASE curReadTarfond.har='d'
        REPLACE fieldValue WITH '  '+IIF(!EMPTY(fSupl.BoxRead.Value),DTOC(fSupl.BoxRead.Value),'')
ENDCASE   
DO procObjReadEnabled WITH .F.,.T.
**********************************************************************************************************************
PROCEDURE lostBoxRead
boxctrl=''
fSupl.boxRead.ControlSource='boxctrl'
***********************************************************************************************************************
PROCEDURE keyPressBoxRead
**************************************************************************************************************************
*     Редактирование стажа на дату приема
**************************************************************************************************************************
PROCEDURE procReadStajIn
PARAMETERS parRead
IF datjob.kodpeop=0
   RETURN 
ENDIF
DO procObjReadEnabled WITH 2      
WITH fSupl
     ctrlCont1=curReadTarfond.rec
     .Cont1.ControlSource='ctrlCont1'     
     .boxRead.ControlSource='datjob.staj_in'
     .boxRead.InputMask='99-99-99'
     .boxRead.Enabled=.T.
     .boxRead.SetFocus    
ENDWITH
**************************************************************************************************************************
PROCEDURE procKvalRead
IF datjob.kodpeop=0
   RETURN 
ENDIF
SELECT cursprkval
REPLACE fl WITH .F.,otm WITH '' ALL
LOCATE FOR kod=datjob.kv
REPLACE fl WITH .T.,otm WITH ' • '
DO procObjReadEnabled WITH 1
WITH fSupl
     .listBox2.ControlSource=''
     .listBox2.rowSource='cursprkval.otm,name' 
     .listBox2.procForClick='Do clickKval'  
     .listBox2.procForKeyPress="Do keyPressListBox2 WITH 'DO clickKval'"
     .listBox2.SetFocus()
ENDWITH
*****************************************************************************************************************************
PROCEDURE clickKval
curTarRec=RECNO('curReadTarFond')
newPkval=cursprkval.doplkat
SELECT cursprkval
rrec=RECNO()
GO rrec
SELECT datjob
REPLACE kv WITH cursprkval.kod
IF INLIST(datJob.kat,1,2)
   REPLACE pkat WITH IIF(kv=0,5,newPkval) 
   SELECT curReadTarFond
   LOCATE FOR LOWER(ALLTRIM(fiRead))='datjob.pkat'
   REPLACE fieldValue WITH '  '+LTRIM(STR(datJob.pkat))
   GO curTarRec
   SELECT datjob
   IF datjob.kv>0
      REPLACE lkv WITH .T.
   ENDIF
   DO CASE
      CASE datJob.kv=0
           IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf,namekf WITH sprdolj.namekf
           ENDIF      
           SELECT curReadTarFond 
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
      CASE datJob.kv=3
          IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf1,namekf WITH sprdolj.namekf1 
           ENDIF      
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
      CASE datJob.kv=2
           IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf2,namekf WITH sprdolj.namekf2 
           ENDIF      
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
      CASE datJob.kv=1
           IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf3,namekf WITH sprdolj.namekf3 
           ENDIF      
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
   ENDCASE
ENDIF
SELECT curReadTarFond
REPLACE fieldValue WITH '  '+cursprkval.name

DO procObjReadEnabled WITH .F.,.T.

*************************************************************************************************************************
PROCEDURE keyPressListBox2
PARAMETERS parProc
DO CASE
   CASE LASTKEY()=27
        DO procObjReadEnabled WITH .F.,.T.
   CASE LASTKEY()=13
        &parProc
ENDCASE   
**************************************************************************************************************************
PROCEDURE procTkfRead
IF datjob.kodpeop=0
   RETURN 
ENDIF
SELECT curSprKoef
SET ORDER TO 1
REPLACE fl WITH .F.,otm WITH '' ALL
LOCATE FOR kod=datjob.kf
REPLACE fl WITH .T.,otm WITH ' • '
DO procObjReadEnabled WITH 1
WITH fSupl     
     .listBox2.rowSource='curSprKoef.otm,nkf'
     .listBox2.ControlSource=''     
     .listBox2.procForClick='DO ClickKf'
     .listBox2.procForKeyPress="Do keyPressListBox2 WITH 'DO clickKf'"
ENDWITH
**************************************************************************************************************************
PROCEDURE clickKf
SELECT datjob
REPLACE kf WITH curSprKoef.kod,namekf WITH curSprKoef.name
SELECT curReadTarFond
REPLACE fieldValue WITH '  '+curSprKoef.nkf
rrec=RECNO()
LOCATE FOR LOWER(ALLTRIM(curReadTarFond.fname))='datjob.namekf'.AND.!logKf
REPLACE fieldValue WITH '  '+LTRIM(STR(curSprKoef.name,6,2))
GO rrec 

GO rrec
DO procObjReadEnabled WITH .F.,.T.

**************************************************************************************************************************
PROCEDURE procPodrPeople
IF datjob.kodpeop=0
   RETURN 
ENDIF
SELECT curSprPodr
REPLACE fl WITH .F.,otm WITH '' ALL
LOCATE FOR kod=datjob.kp
REPLACE fl WITH .T.,otm WITH ' • '
DO procObjReadEnabled WITH 1
WITH fSupl
     .listBox2.ControlSource=''
     .listBox2.rowSource='cursprpodr.otm,name' 
     .listBox2.procForClick='Do clickPodr'   
     .listBox2.procForKeyPress="Do keyPressListBox2 WITH 'DO clickPodr'"
     .listBox2.SetFocus()
ENDWITH
*************************************************************************************************************************
PROCEDURE clickPodr
SELECT datjob
REPLACE kp WITH curSprPodr.kod
SELECT curReadTarFond
REPLACE fieldValue WITH '  '+curSprpodr.name
 
DO procObjReadEnabled WITH .F.,.T.
**************************************************************************************************************************
PROCEDURE procDoljPeople
IF datjob.kodpeop=0
   RETURN 
ENDIF
IF USED('curSuplDolj')
   SELECT curSuplDolj
   USE
ENDIF
SELECT * FROM rasp WHERE kp=datjob.kp INTO CURSOR curSuplDolj ORDER BY nd READWRITE
ALTER TABLE curSuplDolj ADD COLUMN name C(150)
ALTER TABLE curSuplDolj ADD COLUMN otm C(3)
ALTER TABLE curSuplDolj ADD COLUMN fl L
SELECT curSuplDolj
REPLACE name WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,'') ALL 
INDEX ON nd TAG T1
LOCATE FOR kd=datjob.kd
REPLACE fl WITH .T.,otm WITH ' • '

DO procObjReadEnabled WITH 1
WITH fSupl
     .listBox2.ControlSource=''
     .listBox2.rowSource='curSuplDolj.otm,name' 
     .listBox2.procForClick='Do clickDolj'   
     .listBox2.procForKeyPress='Do keyPressListDolj'
     .listBox2.SetFocus()
ENDWITH
**************************************************************************************************************************
PROCEDURE clickDolj
SELECT datjob
REPLACE kd WITH curSuplDolj.kd,kat WITH curSuplDolj.kat
curTarRec=RECNO('curReadTarFond')
IF INLIST(datJob.kat,1,2).AND.datjob.namekf=0
   DO CASE
      CASE datJob.kv=0
          IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf,namekf WITH sprdolj.namekf 
           ENDIF      
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
      CASE datJob.kv=3
           IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf1,namekf WITH sprdolj.namekf1 
           ENDIF                 
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
      CASE datJob.kv=2
           IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf2,namekf WITH sprdolj.namekf2 
           ENDIF      
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
      CASE datJob.kv=1
           IF SEEK(datJob.kd,'sprdolj',1)
              REPLACE kf WITH sprdolj.kf3,namekf WITH sprdolj.namekf3 
           ENDIF      
           SELECT curReadTarFond
           LOCATE FOR logKf
           REPLACE fieldValue WITH '  '+LTRIM(STR(datjob.kf))+' - '+STR(datjob.namekf,5,3)  
           GO curTarRec
   ENDCASE
ENDIF

SELECT curReadTarFond
REPLACE fieldValue WITH '  '+curSuplDolj.name
DO procObjReadEnabled WITH .F.,.T.
**************************************************************************************************************************
PROCEDURE keyPressListDolj
**************************************************************************************************************************
PROCEDURE procTrRead
IF datjob.kodpeop=0
   RETURN 
ENDIF
SELECT sprType
REPLACE fl WITH .F.,otm WITH '' ALL
LOCATE FOR kod=datjob.tr
REPLACE fl WITH .T.,otm WITH ' • '
DO procObjReadEnabled WITH 1
WITH fSupl
     .listBox2.ControlSource=''
     .listBox2.rowSource='sprType.otm,name' 
     .listBox2.procForClick='Do clickType'  
     .listBox2.procForKeyPress="Do keyPressListBox2 WITH 'DO clickType'"
     .listBox2.SetFocus()
ENDWITH
*****************************************************************************************************************************
PROCEDURE clickType
SELECT sprType
rrec=RECNO()
*REPLACE fl WITH IIF(fl,.F.,.T.)
*REPLACE otm WITH IIF(fl,' • ','')
GO rrec
SELECT datJob
REPLACE tr WITH sprType.kod
SELECT curReadTarFond
REPLACE fieldValue WITH '  '+sprType.name
DO procObjReadEnabled WITH .F.,.T.
**************************************************************************************************************************
PROCEDURE procKseRead
IF datjob.kodpeop=0
   RETURN 
ENDIF
DO procObjReadEnabled WITH 2      
WITH fSupl     
     ctrlCont1=curReadTarfond.rec
     .Cont1.ControlSource='ctrlCont1'     
     .boxRead.ControlSource='datJob.kse'     
     .boxRead.Enabled=.T.
     .boxRead.SetFocus    
ENDWITH
**************************************************************************************************************************
PROCEDURE procLmolRead
SELECT datjob
REPLACE lmol WITH IIF(datjob.lmol,.F.,.T.) 
SELECT curReadTarFond
REPLACE fieldValue WITH '  '+ IIF(datjob.lmol,'да','')   
**************************************************************************************************************************
PROCEDURE procDelTarifPeop
IF EMPTY(curReadTarfond.fieldValue)
   RETURN 
ENDIF 
WITH fSupl 
     .listBox1.Enabled=.F.     
     .butExit.Visible=.F.            
     .boxRead.Enabled=.F.
     .butDel.Visible=.T.
     .butReturn.Visible=.T.    
     .txtDel.Visible=.T.
      nameDel='Удаляемый реквизит - '+ALLTRIM(curReadTarfond.rec)
     .txtDel.ControlSource='nameDel'     
ENDWITH 
**************************************************************************************************************************
PROCEDURE procDelTarif
SELECT curReadTarfond
delFi=ALLTRIM(fiRead)
REPLACE fieldValue WITH ''
DO CASE 
   CASE curReadTarFond.har='c'
       * SELECT people
        REPLACE &delFi WITH ''
   CASE curReadTarFond.har='n'
        REPLACE &delFi WITH 0
   CASE curReadTarFond.har='d'
        REPLACE &delFi WITH CTOD('  .  .    ') 
   CASE EMPTY(curReadTarFond.har)
        REPLACE &delFi WITH 0      
ENDCASE 
IF LOWER(ALLTRIM(curReadTarfond.firead))='datjob.kv'
   REPLACE datjob.lkv WITH .F.
ENDIF
DO procObjReadEnabled WITH .F.,.T.
**************************************************************************************************************************
PROCEDURE procObjReadEnabled
PARAMETERS par1,par2
*par1=1 -меню
*par1=2 - boxread
*par2=.T. отключить объекты редактирования
DO CASE
   CASE par2=.F..AND.par1=1   
        WITH fSupl
             .listBox1.Visible=.F. 
             .ListBox2.Enabled=.T.         
             .butExit.Visible=.F.
             .listBox2.Visible=.T.
             .boxRead.Enabled=.F.
             .butSave.Visible=.T.
             .butReturn.Visible=.T.
*             .listBox1.procForClick='DO clickListBoxRead'    
        ENDWITH 
   CASE par2=.F..AND.par1=2   
        WITH fSupl             
             .butExit.Visible=.F.          
             .butSave.Visible=.T.
             .butReturn.Visible=.T.
             .boxRead.Enabled=.T.  
             .txtDel.Visible=.F.               
 *            .listBox1.procForClick='DO clickListBoxRead'    
        ENDWITH      
   CASE par2=.T.   
        WITH fSupl
             .listBox1.Visible=.T.
             .ListBox1.Enabled=.T.          
             .butExit.Visible=.T.
             .listBox2.Visible=.F.
             .boxRead.Enabled=.F.
             .butDel.Visible=.F.
             .butSave.Visible=.F.
             .butReturn.Visible=.F.  
             .txtDel.Visible=.T.          
  *           .listBox1.procForClick='DO clickListBoxRead'    
              KEYBOARD '{TAB}' 
        ENDWITH 
ENDCASE      
************************************************************************************************************************
PROCEDURE actualstaj1
IF datjob.staj_in#people.staj_in
   REPLACE datjob.staj_in WITH people.staj_in
ENDIF
IF !EMPTY(datjob.date_In).AND.!datjob.vac
   SELECT datjob   
   currentStaj=''
   dMbeg=0
   dMEnd=0 
   newDBeg=date_In    
   *newDEnd=varDtar-1
   newDEnd=varDtar
   y_stIn=IIF(EMPTY(staj_in),0,ROUND(VAL(LEFT(staj_in,2)),0))
   m_stIn=IIF(EMPTY(staj_in),0,VAL(SUBSTR(staj_in,4,2)))
   d_stIn=IIF(EMPTY(staj_in),0,VAL(SUBSTR(staj_in,7,2)))      
   
   y_st=0
   m_st=0
   d_st=0
   
   y_new=0
   m_new=0
   d_new=0      
   
   dayMonthBeg=0
   dayMonthEnd=0 
   IF MONTH(newDBeg)=2 
      dayMonthBeg=IIF(MOD(YEAR(newDBeg),4)=0,29,28)
   ELSE
      dayMonthBeg=IIF(INLIST(MONTH(newDBeg),1,3,5,7,8,10,12),31,30) &&кол-во дней в начальном месяце  
   ENDIF
   IF MONTH(newDend)=2 
      dayMonthEnd=IIF(DAY(varDtar)=1.AND.MONTH(varDtar)=3,30,IIF(MOD(YEAR(newDEnd),4)=0,29,28))
   ELSE 
      dayMonthEnd=IIF(INLIST(MONTH(newDEnd),1,3,5,7,8,10,12),31,30)  &&кол-во дней в конечном месяце
   ENDIF
   
 *-----считаем дни
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMBeg=DAY(newDEnd)-DAY(newDBeg)+1
   ELSE 
      dMbeg=dayMonthBeg-DAY(newDBeg)+1
   ENDIF   
   IF dMBeg=dayMonthBeg
      m_new=m_new+1
      dMbeg=0
   ENDIF  
   d_new=d_new+dMBeg

   dMEnd=0
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMEnd=0
   ELSE
      dMEnd=DAY(newDEnd)+IIF(DAY(varDtar)=1.AND.MONTH(varDtar)=3,2,0)
   ENDIF    
   IF dMEnd=dayMonthEnd
      m_new=m_new+1
      dMEnd=0
   ENDIF   
   d_new=d_New+dMEnd

 *-------считаем месяцы 
   mEbeg=0
   mYEnd=0
   IF YEAR(newDBeg)=YEAR(newDEnd)
      m_new=m_new+MONTH(newDEnd)-MONTH(newDBeg)-1
      m_new=IIF(m_new<0,0,m_new)
   ELSE 
      mYbeg=12-MONTH(newDBeg)
      mYEnd=MONTH(newDEnd)-1
      m_new=m_new+mYbeg+mYEnd
   ENDIF   
 
 
 *------------считаем годы------
   IF YEAR(newDBeg)=YEAR(newDEnd)
      y_new=0
   ELSE 
      y_new=YEAR(newDEnd)-YEAR(newDBeg)-1   
   ENDIF 

   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   newYst=y_new
   newMst=m_new
   newDst=d_new
   stajOrg=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')    
   
   
   y_new=y_new+y_stIn
   m_new=m_new+m_stIn
   d_new=d_new+d_stIn
   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   IF m_new<0
      m_new=0
      y_new=IIF(y_new=0,0,y_new-1)
   ENDIF  
   currentStaj=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')
   REPLACE staj_tar WITH currentStaj      
ELSE
   IF datjob.vac
      replace staj_tar WITH dimConstVac(1,2)
   ENDIF                           
ENDIF  
SELECT datJob
rep_st=0
y_st=IIF(EMPTY(staj_tar),0,VAL(LEFT(staj_tar,2)))
m_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,4,2)))
d_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,7,2)))
FOR i=1 TO 8
    IF y_st>=sum_pr(i,1).AND.y_st<sum_pr(i,2)
       rep_st=sum_pr(i,3)
       EXIT
    ENDIF
    IF y_st>=sum_pr(i,1).AND.sum_pr(i,2)=0
       rep_st=sum_pr(i,3)
       EXIT
    ENDIF
ENDFOR
REPLACE stpr WITH rep_st



************************************************************************************************************************
PROCEDURE actualstaj
PARAMETERS parBase,pardate,parEnd,parVar,parStart
IF datjob.staj_in#people.staj_in
   REPLACE datjob.staj_in WITH people.staj_in
ENDIF
IF !EMPTY(&parDate)
   SELECT &parBase
   currentStaj=''
   dMbeg=0
   dMEnd=0 
   newDBeg=&pardate   
   *newDEnd=&parEnd-1
   newDEnd=&parEnd
   y_stIn=IIF(EMPTY(staj_in).OR.parStart,0,ROUND(VAL(LEFT(staj_in,2)),0))
   m_stIn=IIF(EMPTY(staj_in).OR.parStart,0,VAL(SUBSTR(staj_in,4,2)))
   d_stIn=IIF(EMPTY(staj_in).OR.parStart,0,VAL(SUBSTR(staj_in,7,2)))      
   
   y_st=0
   m_st=0
   d_st=0
   
   y_new=0
   m_new=0
   d_new=0  
  
   dayMonthBeg=0
   dayMonthEnd=0 
   IF MONTH(newDBeg)=2 
      dayMonthBeg=IIF(MOD(YEAR(newDBeg),4)=0,29,28)
   ELSE
      dayMonthBeg=IIF(INLIST(MONTH(newDBeg),1,3,5,7,8,10,12),31,30) &&кол-во дней в начальном месяце  
   ENDIF
   IF MONTH(newDend)=2 
      dayMonthEnd=IIF(DAY(&parDate)=1.AND.MONTH(&parDate)=3,30,IIF(MOD(YEAR(newDEnd),4)=0,29,28))
   ELSE 
      dayMonthEnd=IIF(INLIST(MONTH(newDEnd),1,3,5,7,8,10,12),31,30)  &&кол-во дней в конечном месяце
   ENDIF
   
 *-----считаем дни
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMBeg=DAY(newDEnd)-DAY(newDBeg)+1
   ELSE 
      dMbeg=dayMonthBeg-DAY(newDBeg)+1
   ENDIF   
   IF dMBeg=dayMonthBeg
      m_new=m_new+1
      dMbeg=0
   ENDIF  
   d_new=d_new+dMBeg
 
   dMEnd=0
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMEnd=0
   ELSE
      dMEnd=DAY(newDEnd)+IIF(DAY(&parDate)=1.AND.MONTH(&parDate)=3,2,0)
   ENDIF    
   IF dMEnd=dayMonthEnd
      m_new=m_new+1
      dMEnd=0
   ENDIF   
   d_new=d_New+dMEnd
 *-------считаем месяцы 
   mEbeg=0
   mYEnd=0
   IF YEAR(newDBeg)=YEAR(newDEnd)
      m_new=m_new+MONTH(newDEnd)-MONTH(newDBeg)-1
      m_new=IIF(m_new<0,0,m_new)
   ELSE 
      mYbeg=12-MONTH(newDBeg)
      mYEnd=MONTH(newDEnd)-1
      m_new=m_new+mYbeg+mYEnd
   ENDIF 
 *------------считаем годы------
   IF YEAR(newDBeg)=YEAR(newDEnd)
      y_new=0
   ELSE 
      y_new=YEAR(newDEnd)-YEAR(newDBeg)-1   
   ENDIF 

   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   newYst=y_new
   newMst=m_new
   newDst=d_new
   stajOrg=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')    
   
   
   y_new=y_new+y_stIn
   m_new=m_new+m_stIn
   d_new=d_new+d_stIn
   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   IF m_new<0
      m_new=0
      y_new=IIF(y_new=0,0,y_new-1)
   ENDIF
   currentStaj=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')
   REPLACE staj_tar WITH currentStaj  
ELSE
   IF datjob.vac
      replace staj_tar WITH dimConstVac(1,2)
   ENDIF  
ENDIF  
SELECT datJob
rep_st=0
y_st=IIF(EMPTY(staj_tar),0,VAL(LEFT(staj_tar,2)))
m_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,4,2)))
d_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,7,2)))
FOR i=1 TO 8
    IF y_st>=sum_pr(i,1).AND.y_st<sum_pr(i,2)
       rep_st=sum_pr(i,3)
       EXIT
    ENDIF
    IF y_st>=sum_pr(i,1).AND.sum_pr(i,2)=0
       rep_st=sum_pr(i,3)
       EXIT
    ENDIF
ENDFOR
REPLACE stpr WITH rep_st


************************************************************************************************************************
PROCEDURE actualstajtoday
PARAMETERS parBase,pardate,parEnd

IF !EMPTY(&parDate)
   SELECT &parBase
   currentStaj=''
   dMbeg=0
   dMEnd=0   
   newDBeg=&pardate   
   newDEnd=&parEnd
   y_stIn=IIF(EMPTY(staj_in),0,ROUND(VAL(LEFT(staj_in,2)),0))
   m_stIn=IIF(EMPTY(staj_in),0,VAL(SUBSTR(staj_in,4,2)))
   d_stIn=IIF(EMPTY(staj_in),0,VAL(SUBSTR(staj_in,7,2)))      
   
   y_st=0
   m_st=0
   d_st=0
   
   y_new=0
   m_new=0
   d_new=0  
  
   dayMonthBeg=0
   dayMonthEnd=0 
   IF MONTH(newDBeg)=2 
      dayMonthBeg=IIF(MOD(YEAR(newDBeg),4)=0,29,28)
   ELSE
      dayMonthBeg=IIF(INLIST(MONTH(newDBeg),1,3,5,7,8,10,12),31,30) &&кол-во дней в начальном месяце  
   ENDIF
   IF MONTH(newDBeg)=2 
      dayMonthEnd=IIF(MOD(YEAR(newDEnd),4)=0,29,28)
   ELSE
      dayMonthEnd=IIF(INLIST(MONTH(newDEnd),1,3,5,7,8,10,12),31,30)  &&кол-во дней в конечном месяце
   ENDIF
 *-----считаем дни
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMBeg=DAY(newDEnd)-DAY(newDBeg)+1
   ELSE 
      dMbeg=dayMonthBeg-DAY(newDBeg)+1
   ENDIF   
   IF dMBeg=dayMonthBeg
      m_new=m_new+1
      dMbeg=0
   ENDIF  
   d_new=d_new+dMBeg
 
   dMEnd=0
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMEnd=0
   ELSE
      dMEnd=DAY(newDEnd)
   ENDIF    
   IF dMEnd=dayMonthEnd
      m_new=m_new+1
      dMEnd=0
   ENDIF   
   d_new=d_New+dMEnd
 *-------считаем месяцы 
   mEbeg=0
   mYEnd=0
   IF YEAR(newDBeg)=YEAR(newDEnd)
      m_new=m_new+MONTH(newDEnd)-MONTH(newDBeg)-1
      m_new=IIF(m_new<0,0,m_new)
   ELSE 
      mYbeg=12-MONTH(newDBeg)
      mYEnd=MONTH(newDEnd)-1
      m_new=m_new+mYbeg+mYEnd
   ENDIF 
 *------------считаем годы------
   IF YEAR(newDBeg)=YEAR(newDEnd)
      y_new=0
   ELSE 
      y_new=YEAR(newDEnd)-YEAR(newDBeg)-1   
   ENDIF 

   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   newYst=y_new
   newMst=m_new
   newDst=d_new
   stajOrg=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')    
   
   
   y_new=y_new+y_stIn
   m_new=m_new+m_stIn
   d_new=d_new+d_stIn
   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   IF m_new<0
      m_new=0
      y_new=IIF(y_new=0,0,y_new-1)
   ENDIF
   currentStaj=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')
   REPLACE staj_today WITH currentStaj   
ELSE                           
ENDIF  
************************************************************************************************************************
PROCEDURE actualstajtoday1
IF !EMPTY(people.date_In).AND.!people.vac
   SELECT people
   currentStaj=''
   dBeg=date_In   
   dEnd=DATE()
   y_st=IIF(EMPTY(staj_in),0,ROUND(VAL(LEFT(staj_in,2)),0))
   m_st=IIF(EMPTY(staj_in),0,VAL(SUBSTR(staj_in,4,2)))
   d_st=IIF(EMPTY(staj_in),0,VAL(SUBSTR(staj_in,7,2)))      
   
   y_new=0
   m_new=0
   d_new=0                    
   DO CASE      
      CASE DAY(dEnd)>=DAY(dBeg)
           dayMonth=IIF(INLIST(MONTH(dBeg),1,3,5,7,8,10,12),31,IIF(MONTH(dBeg)#2,30,28)) &&кол-во дней в месяце  
           dayMonthEnd=IIF(INLIST(MONTH(dBeg),1,3,5,7,8,10,12),31,IIF(MONTH(dBeg)#2,30,28)) &&кол-во дней в конечном месяце            
           d_new=IIF(DAY(dEnd)-DAY(dBeg)+1=dayMonth,0,DAY(dEnd)-DAY(dBeg)+1)                                   
           m_new=IIF(DAY(dEnd)-DAY(dBeg)+1=dayMonth,1,m_new)                        
          
           IF MONTH(dEnd)>=MONTH(dBeg)
              m_new=m_new+(MONTH(dEnd)-MONTH(dBeg))
              y_new=YEAR(dEnd)-YEAR(dBeg)   
           ELSE              
              m_new=m_new+(12-MONTH(dBeg))+IIF(DAY(dEnd)=dayMonthEnd,MONTH(dEnd),MONTH(dEnd))
              y_new=YEAR(dEnd)-YEAR(dBeg)-1   
           ENDIF                                       
      CASE DAY(dEnd)<DAY(dBeg)
           dayMonth=IIF(INLIST(MONTH(dBeg),1,3,5,7,8,10,12),31,IIF(MONTH(dBeg)#2,30,28)) &&кол-во дней в начальном месяце    
           dayMonthEnd=IIF(INLIST(MONTH(dBeg),1,3,5,7,8,10,12),31,IIF(MONTH(dBeg)#2,30,28)) &&кол-во дней в конечном месяце                               
           d_new=IIF(d_new=dayMonth,0,dayMonth-DAY(dBeg)+1)          
           d_new=d_new+DAY(dEnd)                   
           IF MONTH(dEnd)>=MONTH(dBeg)
              m_new=m_new+(MONTH(dEnd)-MONTH(dBeg)-1)
              y_new=YEAR(dEnd)-YEAR(dBeg)   
           ELSE
              m_new=m_new+(12-MONTH(dBeg))+IIF(DAY(dEnd)=dayMonthEnd,MONTH(dEnd),MONTH(dEnd)-1)
              y_new=YEAR(dEnd)-YEAR(dBeg)-1    
           ENDIF                                                               
   ENDCASE
   y_new=y_new+y_st
   m_new=m_new+m_st
   d_new=d_new+d_st-1
   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   currentStaj=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')
   REPLACE staj_today WITH currentStaj  
ELSE
   IF people.vac 
      replace staj_today WITH ''
   ENDIF                           
ENDIF   
**********************************************************************************************************
*             Расчет переходящего стажа по одному человеку
**********************************************************************************************************
PROCEDURE perstajoneold
SELECT datjob
REPLACE st_per WITH 0,per_date WITH CTOD('  /  /  '),per_sum WITH 0 
STORE 0 TO d_new,m_new,y_new
y_st=IIF(EMPTY(staj_tar),0,ROUND(VAL(LEFT(staj_tar,2)),0))
m_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,4,2)))
d_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,7,2)))   
FOR i=1 TO 8
    IF y_st=sum_pr(i,2)-1
       *m_new=11-m_st+MONTH(varDtar)
       *d_new=30-d_st+DAY(varDtar)
       
       m_new=MONTH(varDtar)+11-m_st
       IF ALLTRIM(staj_in)='00-00-00'
          d_new=DAY(date_in) 
          m_new=MONTH(date_in)
       ELSE
          d_new=DAY(varDtar)+30-d_st      
          IF d_new>30
             d_new=d_new-30
             m_new=m_new+1
          ENDIF      
       ENDIF      
                         
       IF m_new<13                   
          date_cx=STR(d_new,2)+'.'+STR(m_new,2)+'.'+STR(YEAR(varDtar),4)
          IF d_new>28.AND.m_new=2
             date_cx='01.03.'+STR(YEAR(varDtar),4)
          ENDIF
          IF d_new=31.AND.INLIST(m_new,4,6,9,11)
             date_cx='01.'+STR(m_new+1,2)+'.'+STR(YEAR(varDtar),4)
          ENDIF                 
          REPLACE st_per WITH sum_pr(i+1,3),per_date WITH CTOD(date_cx)
                  r_sum=varBaseSt*st_per/100                   
          REPLACE per_sum WITH r_sum
        *  DO dayfond                
       ENDIF
       EXIT
    ENDIF
ENDFOR   




**********************************************************************************************************
*             Расчет переходящего стажа по одному человеку
**********************************************************************************************************
PROCEDURE perstajone1
SELECT datjob
REPLACE st_per WITH 0,per_date WITH CTOD('  /  /  '),per_sum WITH 0 
STORE 0 TO d_new,m_new,y_new
y_st=IIF(EMPTY(staj_tar),0,ROUND(VAL(LEFT(staj_tar,2)),0))
m_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,4,2)))
d_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,7,2)))   
FOR i=1 TO 8
    IF y_st=sum_pr(i,2)-1
       *m_new=11-m_st+MONTH(varDtar)
       *d_new=30-d_st+DAY(varDtar)
       
       m_new=MONTH(varDtar)+11-m_st
       IF ALLTRIM(staj_in)='00-00-00'
          d_new=DAY(date_in) 
          m_new=MONTH(date_in)
       ELSE
          d_new=DAY(varDtar)+30-d_st      
          IF d_new>30
             d_new=d_new-30
             m_new=m_new+1
          ENDIF      
       ENDIF      
                         
       IF m_new<13                   
          date_cx=STR(d_new,2)+'.'+STR(m_new,2)+'.'+STR(YEAR(varDtar),4)
          IF d_new>28.AND.m_new=2
             date_cx='01.03.'+STR(YEAR(varDtar),4)
          ENDIF
          IF d_new=31.AND.INLIST(m_new,4,6,9,11)
             date_cx='01.'+STR(m_new+1,2)+'.'+STR(YEAR(varDtar),4)
          ENDIF                 
          REPLACE st_per WITH sum_pr(i+1,3),per_date WITH CTOD(date_cx)
                  r_sum=varBaseSt*st_per/100                   
          REPLACE per_sum WITH r_sum
        *  DO dayfond                
       ENDIF
       EXIT
    ENDIF
ENDFOR   

**********************************************************************************************************
*             Расчет переходящего стажа по одному человеку
**********************************************************************************************************
PROCEDURE perstajone
SELECT datjob
REPLACE st_per WITH 0,per_date WITH CTOD('  /  /  '),per_sum WITH 0 
STORE 0 TO d_new,m_new,y_new, d_rest,y_rest

y_st=IIF(EMPTY(staj_tar),0,ROUND(VAL(LEFT(staj_tar,2)),0))
m_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,4,2)))
d_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,7,2)))
y_new=YEAR(varDtar)
IF INLIST (y_st,4,9,14)
    DO CASE 
       CASE ALLTRIM(staj_in)='00-00-00'            
            *d_new=DAY(date_in-1) 
            *m_new=MONTH(date_in-1)
            
            d_new=DAY(date_in) 
            m_new=MONTH(date_in)

            m_rest=12-m_st
            y_new=y_new+IIF(MONTH(varDtar)+m_rest>12,1,0)         
        OTHERWISE 
            d_rest=IIF(d_st=0,0,31-d_st)
            d_new=IIF((d_rest+DAY(varDtar))<30,d_rest+DAY(varDtar),d_rest+DAY(varDtar)-30)
            m_rest=12-m_st-IIF(d_rest>0,1,0)
            m_new=m_rest+MONTH(varDtar)+IIF((d_rest+DAY(varDtar))>=30,1,0)
            y_new=y_new+IIF(MONTH(varDtar)+m_rest>12,1,0)
    ENDCASE    
    date_cx=IIF(y_new==YEAR(varDtar),STR(d_new,2)+'.'+STR(m_new,2)+'.'+STR(YEAR(varDtar),4),'  .  .    ') 
    REPLACE per_date WITH CTOD(date_cx) 
    
    y_st=IIF(EMPTY(staj_tar),0,VAL(LEFT(staj_tar,2)))
    m_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,4,2)))
    d_st=IIF(EMPTY(staj_tar),0,VAL(SUBSTR(staj_tar,7,2)))
    FOR i=1 TO 8
        IF y_st>=sum_pr(i,1).AND.y_st<sum_pr(i,2)
           rep_st=sum_pr(i,3)
           EXIT
        ENDIF
        IF y_st>=sum_pr(i,1).AND.sum_pr(i,2)=0
           rep_st=sum_pr(i,3)
           EXIT
        ENDIF
    ENDFOR
    REPLACE stpr WITH rep_st, st_per WITH sum_pr(i+1,3),per_sum WITH varBaseSt*st_per/100    
ENDIF 
************************************************************************************************************************
*                      расчет оклада и доплат по одному человеку
************************************************************************************************************************
PROCEDURE dopl_sum  
PARAMETERS par_ref
SELECT sprdolj
SEEK datjob.kd
SELECT people
IF !EMPTY(dmol).AND.dmol<varDtar
   REPLACE dmol WITH CTOD('  .  .    ' )
   SELECT datjob
   REPLACE pmols WITH 0
ENDIF
SELECT datjob
repVr='vr'+LTRIM(STR(MONTH(varDtar)))
REPLACE patt WITH pkfvr,satt WITH &repVr,matt WITH &repVr
IF EMPTY(date_in)
   REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in 
ENDIF
IF INLIST(kat,1,2,5,7)
   IF lkv.AND.kv>0
      REPLACE pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,5)
   ELSE 
      REPLACE pkat WITH 5
   ENDIF 
ELSE 
   IF !sprdolj.lspis
      REPLACE pkat WITH 0   
   ELSE    
      REPLACE pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,5)
   ENDIF    
ENDIF
DO actualstaj WITH 'datjob','datjob.date_in','varDtar'
DO perstajone
IF datshtat.real
   SELECT rasp
   nOldRaspOrd=SYS(21)
   SET ORDER TO 2
   SEEK STR(datjob.kp,3)+STR(datjob.kd,3)
ENDIF 
SELECT tarfond
SET FILTER TO 
totsumf=0
totsumfm=0
totfondprn=0
SCAN ALL
     IF datShtat.real.AND.!EMPTY(procrepnad)                
        ON ERROR DO ersup
        procrep=ALLTRIM(tarfond.procrepnad)
        DO &procrep   
        SELECT tarfond
        ON ERROR 
     ENDIF      
     IF !EMPTY(formula)
        new_sum=sum_f
        new_msum=0
        pole=fname                 
        r_sum=ALLTRIM(tarfond.formula)  
        r_sum1=ALLTRIM(tarfond.formula1) 
        SELECT datJob 
        DO CASE
            CASE !EMPTY(tarfond.proccount)
                 procForCount=ALLTRIM(tarfond.proccount)                           
                 DO &procForCount
            CASE !EMPTY(tarfond.formula)
                 IF !EMPTY(tarfond.sum_f)                    
                    REPLACE &new_sum WITH &r_sum 
                    IF !EMPTY(tarfond.sum_fm)                  
                       new_msum=tarfond.sum_fm                  
                       REPLACE &new_msum WITH IIF(tarfond.logkse,&new_sum*datjob.kse,&new_sum)  
                       IF !EMPTY(tarfond.formula1)
                          REPLACE &new_sum WITH &r_sum1  
                       ENDIF
                    ENDIF
                 ELSE
                    SELECT datjob
                    REPLACE &pole WITH &r_sum
                 ENDIF                                
           ENDCASE                 
     ENDIF  
     SELECT datjob   
     totsumf=IIF(!EMPTY(tarfond.sum_f),totsumf+EVALUATE(ALLTRIM(tarfond.sum_f)),totsumf)
     totfondprn=IIF(tarfond.logfprn,totfondprn+EVALUATE(ALLTRIM(tarfond.sum_fm)),totfondprn)
     totsumfm=IIF(!EMPTY(tarfond.sum_fm),totsumfm+EVALUATE(ALLTRIM(tarfond.sum_fm)),totsumfm)
     
     IF tarfond.logit
        REPLACE total WITH totsumf,msf WITH totsumfm,fdprn WITH totfondprn
     ENDIF
     SELECT tarfond            
ENDSCAN
GO TOP
IF !par_ref
   ON ERROR DO ersup
   frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus    
   ON ERROR
ENDIF   
************************************************************************************************************************
PROCEDURE countmols
IF people.mols.AND.!EMPTY(people.dmol).AND.!datjob.lmol
   SELECT datjob
   REPLACE pmols WITH IIF(INLIST(kat,1,5),30,20)
ELSE
   REPLACE pmols WITH 0           
ENDIF 
IF !EMPTY(tarfond.sum_f)                    
   REPLACE &new_sum WITH &r_sum 
   IF !EMPTY(tarfond.sum_fm)                  
      new_msum=tarfond.sum_fm                  
      REPLACE &new_msum WITH IIF(tarfond.logkse,&new_sum*datjob.kse,&new_sum)  
      IF !EMPTY(tarfond.formula1)
         REPLACE &new_sum WITH &r_sum1  
      ENDIF
   ENDIF
ELSE
   SELECT datjob
   REPLACE &pole WITH &r_sum
ENDIF 
**************************************************************************************************************************
*                       Подстановка набавок и доплат для тарификации
**************************************************************************************************************************
PROCEDURE repNadJob 
SELECT datjob
REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in,pkont WITH IIF(tr=1,people.pkont,0)
DO CASE 
   CASE datjob.lkv.AND.people.kval#0
        REPLACE kv WITH people.kval,nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' №'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' от ' +DTOC(people.dkval),''),;
                pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,0)
   OTHERWISE
        REPLACE kv WITH 0,nPrik WITH '',pkat WITH 0   
ENDCASE 
DO CASE
   CASE kv=0
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf,namekf)
   CASE kv=1
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf3,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf3,namekf)
   CASE kv=2
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf2,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf2,namekf)
   CASE kv=3
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf1,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf1,namekf)
ENDCASE
SELECT rasp
nOldRaspOrd=SYS(21)
SET ORDER TO 2
SEEK STR(datjob.kp,3)+STR(datjob.kd,3)
REPLACE datjob.pkf WITH rasp.pkf
SELECT tarfond
GO TOP
SCAN ALL
     IF !EMPTY(plrep).AND.ltar
        IF !EMPTY(procrepnad)
           procrep=ALLTRIM(procrepnad)
           DO &procrep 
        ELSE 
           repjob=ALLTRIM(plrep)
           repjob1='rasp.'+ALLTRIM(plrep)
           SELECT datjob 
           REPLACE &repjob WITH &repjob1
        ENDIF    
     ENDIF
     IF ALLTRIM(LOWER(tarfond.plrep))='pkat'.AND.rasp.pkat#0.AND.datjob.kv=0
        SELECT datjob
        REPLACE pkat WITH 5       
     ENDIF
     SELECT tarfond
ENDSCAN
SELECT rasp
SET ORDER TO &nOldRaspOrd
SELECT datjob

**************************************************************************************************************************
PROCEDURE repnadvto
SELECT datjob
DO CASE 
   CASE kv=0
        REPLACE pvto WITH rasp.pvto4
   CASE kv=1
        REPLACE pvto WITH rasp.pvto1
   CASE kv=2
        REPLACE pvto WITH rasp.pvto2
   CASE kv=3
        REPLACE pvto WITH rasp.pvto3
ENDCASE 
SELECT tarfond
***********************************************************************************************************************
PROCEDURE reppzdrav
SELECT datjob
REPLACE pzdrav WITH rasp.pzdrav1
************************************************************************************************************************
PROCEDURE repnadmain
repjob=ALLTRIM(tarfond.plrep)
repjob1='rasp.'+ALLTRIM(tarfond.plrep)
SELECT datjob 
REPLACE &repjob WITH &repjob1
************************************************************************************************************************
PROCEDURE repfrompeop
SELECT datjob
REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in,pkont WITH IIF(tr=1,people.pkont,0)
DO CASE 
   CASE datjob.lkv.AND.people.kval#0
        REPLACE kv WITH people.kval,nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' №'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' от ' +DTOC(people.dkval),''),;
                pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,0)
   OTHERWISE
        REPLACE kv WITH 0,nPrik WITH '',pkat WITH 0   
ENDCASE 
DO CASE
   CASE kv=0
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf,namekf)
   CASE kv=1
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf3,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf3,namekf)
   CASE kv=2
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf2,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf2,namekf)
   CASE kv=3
        REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf1,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf1,namekf)
ENDCASE
************************************************************************************************************************
*                            основная для места работы и должности
************************************************************************************************************************
PROCEDURE formForJobTarif
fSupl=CREATEOBJECT('FORMSUPL')
DIMENSION dimSelectJob(4)
STORE 0 TO dimSelectJob
dimSelectJob(1)=1
SELECT datJob
ON ERROR DO erSup
oldRecJob=RECNO()
COUNT TO maxJob
IF maxJob#0
   GO oldRecJob
ENDIF   
ON ERROR 
WITH fSupl 
     .Caption='Место работы должность'
     DO addShape WITH 'fSupl',1,20,20,10,10,8
     DO addOptionButton WITH 'fSupl',1,'добавить новую запись',.Shape1.Top+20,.Shape1.Left+10,'dimSelectJob(1)',0,'DO procSelectJob WITH 1',.T. 
     DO addOptionButton WITH 'fSupl',2,'редактировать текущую',.Option1.Top+.Option1.Height+10,.Option1.Left,'dimSelectJob(2)',0,'DO procSelectJob WITH 2',IIF(maxJob=0,.F.,.T.)    
     DO addOptionButton WITH 'fSupl',3,'удалить текущую запись',.Option2.Top+.Option2.Height+10,.Option2.Left,'dimSelectJob(3)',0,'DO procSelectJob WITH 3',IIF(maxJob=0,.F.,.T.)  
     DO addOptionButton WITH 'fSupl',4,'"передать" текущую запись',.Option3.Top+.Option3.Height+10,.Option2.Left,'dimSelectJob(4)',0,'DO procSelectJob WITH 4',IIF(maxJob=0,.F.,.T.)  
     .Shape1.Height=.Option1.Height*4+70
     .Shape1.Width=.Option4.Width+20
     *-----------------------------Кнопка приступить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприступитьw')*2)-20)/2,;
        .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприступитьw'),dHeight+5,'приступить','DO procRunJob'

     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'отмена','fSupl.Release','отмена'    
    .Width=.Shape1.Width+40
    .Height=.Shape1.Height+.cont1.Height+60
ENDWITH

DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE procSelectJob
PARAMETERS par1
STORE 0 TO dimSelectJob
dimSelectJob(par1)=1
fSupl.Refresh
************************************************************************************************************************
PROCEDURE procRunJob
fSupl.Visible=.F.
fSupl.Release
DO CASE
   CASE dimSelectJob(1)=1
        DO inputRecInJob WITH .T.        
   CASE dimSelectJob(2)=1
        DO inputRecInJob WITH .F.
   CASE dimSelectJob(3)=1
        DO deleteFromJob WITH 'DO delRecFromJob'
   CASE dimSelectJob(4)=1
        DO formChangeRecJob     
ENDCASE
***********************************************************************************************************************
PROCEDURE deleteFromJob
PARAMETERS parproc
fdel=CREATEOBJECT('FORMSUPL')
log_del=.F.
WITH fDel  
     .Caption='Удаление'    
     DO addShape WITH 'fDel',1,20,20,100,RetTxtWidth('wпоставьте птичку в окошке, расположенном нижеw'),8         
     DO adLabMy WITH 'fDel',1,'для подтверждения ваших намерений',fDel.Shape1.Top+10,fDel.Shape1.Left+5,.Shape1.Width-10,2 
     DO adLabMy WITH 'fDel',2,'поставьте птичку в окошке, расположенном ниже',.lab1.Top+.lab1.Height,fDel.Shape1.Left+5,.lab1.Width,2                                   
     DO adCheckBox WITH 'fdel','check1','подтверждение удаления',.lab2.Top+.lab2.Height+10,.Shape1.Left,150,dHeight,'log_del',0    
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     .Shape1.Height=.check1.Height+.lab1.Height*2+30
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('wУдалитьw')*2-20)/2,fdel.check1.Top+fdel.check1.Height+20,;
        RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','DO delRecFromJob'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+20,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'     
     .Width=.Shape1.Width+40   
     .Height=.Shape1.Height+.cont1.Height+60     
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
************************************************************************************************************************
PROCEDURE delRecFromJob
IF !log_del
   RETURN
ENDIF
fDel.Release
SELECT datJob
DELETE
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
frmTop.Refresh

**************************************************************************************************************************
PROCEDURE inputRecInJob
PARAMETERS par1
logNewRec=par1
SELECT sprkoef
REPLACE namesup WITH STR(kod,2)+' - '+STR(name,5,3) ALL

STORE 0 TO jobRec,kseCurrent,kseTotal
STORE '' TO str_type,str_podr,str_dolj,str_kval
fSupl=CREATEOBJECT('FORMSUPL')
SELECT * FROM rasp INTO CURSOR curDolPodr READWRITE 
SELECT curDolPodr
REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,''),kf WITH sprdolj.kf,namekf WITH sprdolj.namekf,kf1 WITH sprdolj.kf1,namekf1 WITH sprdolj.namekf1,;
        kf2 WITH sprdolj.kf2,namekf2 WITH sprdolj.namekf2,kf3 WITH sprdolj.kf3,namekf3 WITH sprdolj.namekf3 ALL 
INDEX ON nd TAG T1
new_tr=IIF(par1,0,datjob.tr)
new_podr=IIF(par1,0,datjob.kp)
new_dolj=IIF(par1,0,datjob.kd)
new_kse=IIF(par1,0.00,datjob.kse)
new_kat=IIF(par1,0,datJob.kat)
newStrKat=IIF(SEEK(new_kat,'sprkat',1),sprkat.name,'')
newPkat=IIF(par1,0,datjob.pkat)

new_kval=IIF(par1,0,datJob.kv)
new_tabn=IIF(par1,0,datJob.tabn)

new_kf=IIF(par1,0,datJob.kf)
new_nameKf=IIF(par1,0,datJob.namekf)
newNid=IIF(par1,0,datjob.nid)
log_doubl=.T.
IF !par1
   SELECT curDolPodr  
   SET FILTER TO kp=new_podr  
ENDIF 

str_type=IIF(par1,'',IIF(SEEK(datjob.tr,'sprtype',1),sprtype.name,''))
str_podr=IIF(par1,'',IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.name,''))
str_dolj=IIF(par1,'',IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.namework,''))
str_kval=IIF(par1,'',IIF(SEEK(datjob.kv,'sprkval',1),sprkval.name,''))

str_kf=IIF(par1,'',IIF(!EMPTY(datjob.namekf),STR(datjob.kf,2)+' - '+STR(datjob.namekf,5,3),''))
IF par1
   date_inch=datjob.date_in
   staj_inch=datjob.staj_in
   staj_tarch=datjob.staj_tar
   SELECT datJob
   oldJobRec=RECNO()  
   GO TOP  
   new_tabn=IIF(tabn#0,tabn,people.tabn)
   IF maxJob#0
      GO oldJobRec
   ENDIF   
ENDIF
WITH fSupl    
     .Caption=''
      DO adTboxAsCont WITH 'fSupl','txtPodr',10,10,RetTxtWidth('wтабельный номерw'),dHeight,'подразделение',1,1
      DO addComboMy WITH 'fSupl',1,.txtPodr.Left+.txtPodr.Width-1,.txtPodr.Top,dheight,450,.T.,'str_podr','cursprpodr.name',6,.F.,'DO validPodrInJob',.F.,.T.      
           
      DO adTBoxAsCont WITH 'fSupl','txtDolj',.txtPodr.Left,.txtPodr.Top+.txtPodr.Height-1,.txtPodr.Width,dHeight,'должность',1,1
      DO addComboMy WITH 'fSupl',2,.comboBox1.Left,.txtDolj.Top,dheight,.comboBox1.Width,.T.,'str_dolj','ALLTRIM(curDolPodr.named)',6,.F.,'DO validDoljInJob',.F.,.T.           
      .comboBox1.DisplayCount=15
      
      DO adTBoxAsCont WITH 'fSupl','txtKat',.txtPodr.Left,.txtDolj.Top+.txtDolj.Height-1,.txtPodr.Width,dHeight,'персонал',1,1  
      DO adtboxnew WITH 'fSupl','boxKat',.txtKat.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newStrKat',.F.,.F.,0
      
      DO adTBoxAsCont WITH 'fSupl','txtKval',.txtPodr.Left,.txtKat.Top+.txtKat.Height-1,.txtPodr.Width,dHeight,'категория',1,1   
      DO addComboMy WITH 'fSupl',4,.comboBox1.Left,.txtKval.Top,dheight,.comboBox1.Width,.T.,'str_kval','curSprKval.name',6,.F.,'DO validKatInJob',.F.,.T.    
      
      DO adTBoxAsCont WITH 'fSupl','txtTr',.txtPodr.Left,.txtKval.Top+.txtKval.Height-1,.txtPodr.Width,dHeight,'разряд-кфт',1,1   
      DO addComboMy WITH 'fSupl',14,.comboBox1.Left,.txtTr.Top,dheight,.comboBox1.Width,.T.,'str_kf','sprkoef.namesup',6,.F.,'DO validKfInJob',.F.,.T.           
              
      DO adTBoxAsCont WITH 'fSupl','txtKse',.txtPodr.Left,.txtTr.Top+.txtTr.Height-1,.txtPodr.Width,dHeight,'объём',1,1
      DO addSpinnerMy WITH 'fSupl','spinKse',.txtKse.Left+.txtKse.Width-1,.txtKse.Top,dheight,RetTxtWidth('9999999999'),'new_kse',0.25,.F.,0,1.5
         
      DO adTBoxAsCont WITH 'fSupl','txtType',.spinKse.Left+.spinKse.Width-1,.txtKse.Top,RetTxtWidth('wтип работыw'),dHeight,'тип',2,1                                             
      DO addComboMy WITH 'fSupl',3,.txtType.Left+.txtType.Width-1,.txtType.Top,dheight,.txtPodr.Width+.comboBox1.Width-.txtKse.Width-.spinKse.Width-.txtType.Width+2,;
         .T.,'str_type','sprType.name',6,.F.,'new_tr=sprType.kod',.F.,.T.                                                
                               
      DO adTBoxAsCont WITH 'fSupl','txtTabn',.txtPodr.Left,.txtType.Top+.txtType.Height-1,.txtPodr.Width,dHeight,'табельный номер',1,1   
      DO addtxtboxmy WITH 'fSupl',11,.comboBox1.Left,.txtTabn.Top,.comboBox1.Width,.F.,'new_tabn',0,.F.    
      
      DO adTBoxAsCont WITH 'fSupl','txtDateIn',.txtPodr.Left,.txtTabn.Top+.txtTabn.Height-1,.txtPodr.Width,dHeight,'дата приема',1,1   
      DO addtxtboxmy WITH 'fSupl',12,.comboBox1.Left,.txtDateIn.Top,.comboBox1.Width/2-RetTxtWidth('стаж на дату приемаv'),.F.,'date_inch',0,.F.    
      DO adTBoxAsCont WITH 'fSupl','txtStajIn',.txtBox12.Left+.txtBox12.Width-1,.txtDateIn.Top,RetTxtWidth('стаж на дату приемаv'),dHeight,'стаж на дату приема',1,1   
      DO addtxtboxmy WITH 'fSupl',13,.txtStajIn.Left+.txtStajIn.Width-1,.txtDateIn.Top,.comboBox1.Width-.txtBox12.Width-.txtStajIn.Width+2,.F.,'staj_inch',0,.F.    
            
      .Width=.txtPodr.Width+.comboBox1.Width+19 
      
      IF par1
         DO adCheckBox WITH 'fSupl','check1','скопировать сведения о месте работы',.txtDateIn.Top+.txtDateIn.Height+10,0,150,dHeight,'log_doubl',0        
         .check1.Left=(.Width-.check1.Width)/2
         .check1.Enabled=.F.
      ENDIF 
            
      DO addcontlabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,IIF(par1,.check1.Top+.check1.Height+20,.txtDateIn.Top+.txtDateIn.Height+20),;
         RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать','DO beforeWriteRecInJob'
      DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','DO exitFromInputJob'             
                     
      .Height=.txtPodr.height*8+.cont1.Height+40+IIF(par1,.check1.Height+10,0)     
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
***********************************************************************************************************************
PROCEDURE exitFromInputJob
SELECT datJob
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus   
fSupl.Release 
***********************************************************************************************************************
PROCEDURE validKfInJob
new_kf=sprkoef.kod
new_namekf=sprkoef.name
***********************************************************************************************************************
PROCEDURE validPodrInJob
new_podr=curSprPodr.kod
IF logNewRec.AND.new_podr#0.AND.new_dolj#0
   SELECT datjob
   LOCATE FOR kp=new_podr.AND.kd=new_dolj
   log_doubl=IIF(FOUND(),.T.,.F.)
ELSE
   log_doubl=.F.   
ENDIF    
IF logNewRec
   fSupl.check1.Enabled=IIF(log_doubl,.T.,.F.)
ENDIF     
SELECT curDolPodr
SET FILTER TO kp=new_podr
fSupl.ComboBox2.RowSource='curDolPodR.named'
fSupl.ComboBox2.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
fSupl.ComboBox2.RowSourceType=6
fsupl.comboBox2.ProcForValid='DO validDoljInJob'
KEYBOARD '{TAB}'
SELECT datjob
**************************************************************************************************************************
PROCEDURE beforeWriteRecInJob
fSupl.Visible=.F.
SELECT rasp
SET FILTER TO 
SELECT datjob
IF new_podr=0.OR.new_dolj=0.OR.new_kse=0
   RETURN
ENDIF
SELECT datjob
IF logNewRec.AND.log_doubl
   LOCATE FOR kodpeop=people.num.AND.kp=new_podr.AND.kd=new_dolj 
   SCATTER TO dimNewRec
ENDIF 
IF par1
   SET FILTER TO 
   ordJobOld=SYS(21)
   SET ORDER TO 7
   SET DELETED OFF
   GO BOTTOM 
   newNid=nid+1
   SET DELETED ON
   SET ORDER TO &ordJobOld
   APPEND BLANK
   IF logNewRec.AND.log_doubl
      GATHER FROM dimNewRec      
   ENDIF
   REPLACE date_in WITH date_inch,staj_in WITH staj_inch,staj_tar WITH staj_tarch      
ENDIF    
REPLACE nid WITH newNid,kodpeop WITH people.num,kp WITH new_podr,kd WITH new_dolj,kse WITH new_kse,tr WITH new_tr,;
        kat WITH new_kat,vac WITH people.vac,kv WITH new_kval,tabn WITH new_tabn,kf WITH new_kf,namekf WITH new_namekf,nidpeop WITH people.nid
datJobRec=RECNO()      
IF kat=1.OR.kat=2
   REPLACE pkat WITH IIF(kv=0,5,newPkat)
ELSE
   REPLACE pKat WITH 0   
ENDIF        
IF people.tabn=0
   SELECT people
   REPLACE tabn WITH new_tabn
ENDIF
fSupl.Release
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus    
SELECT datJob
GO datJobRec
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus    
**************************************************************************************************************************
PROCEDURE validDoljInJob
new_dolj=curDolPodr.kd
new_kat=curDolPodr.kat
newStrKat=IIF(SEEK(new_kat,'sprkat',1),sprkat.name,'')
new_kf=IIF(curDolPodr.kfvac#0,curDolPodr.kfvac,curDolPodr.kf)
new_namekf=IIF(curDolPodr.kfvac#0,curDolPodr.nkfvac,curDolPodr.namekf)
str_kf=STR(new_kf,2)+' - '+STR(new_namekf,5,3)
fSupl.comboBox14.Controlsource='str_kf'
IF logNewRec.AND.new_podr#0.AND.new_dolj#0
   SELECT datjob
   LOCATE FOR kp=new_podr.AND.kd=new_dolj
   log_doubl=IIF(FOUND(),.T.,.F.)
ELSE   
   log_doubl=.F.
ENDIF 
IF logNewRec 
   fSupl.check1.Enabled=IIF(log_doubl,.T.,.F.)
ENDIF    
fSupl.boxKat.Refresh
KEYBOARD '{TAB}'
**************************************************************************************************************************
PROCEDURE validKatInjob
new_kval=curSprKval.kod
newPkat=curSprKval.doplkat
DO CASE   
   CASE new_kval=1
        new_kf=curDolPodr.kf3
        new_namekf=curDolPodr.namekf3
   CASE new_kval=2
        new_kf=curDolPodr.kf2
        new_namekf=curDolPodr.namekf2    
   CASE new_kval=3
        new_kf=curDolPodr.kf1
        new_namekf=curDolPodr.namekf1
   OTHERWISE 
        new_kf=curDolPodr.kf
        new_namekf=curDolPodr.namekf
ENDCASE
str_kf=STR(new_kf,2)+' - '+STR(new_namekf,5,3)
fSupl.comboBox14.Controlsource='str_kf'
fSupl.Refresh
***********************************************************************************************************************
PROCEDURE formChangeRecJob   
=AFIELDS(arFio,'people')
CREATE CURSOR curFindFio FROM ARRAY arFio 
SELECT curFindFio
INDEX ON fio TAG T1
SET ORDER TO 1
fSupl=CREATEOBJECT('FORMSUPL')
newFio=''
newKodPeop=0
newTabPeop=0
newvac=.F.
WITH fSupl
     .Caption='Передача должности'    
      DO adtBoxNew WITH 'fSupl','boxFio',10,10,500,dHeight,'newFio',.F.,.T.,.F.,.F.,"DO unosimbol WITH 'newfio'"
     .boxFio.procForChange='DO changeFio'
     
     
     .Width=.boxFio.Width+20     
     
     *---------------------------------Кнопка записать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','butSave',(.Width-RetTxtWidth('WзаписатьW')*2-30)/2,.boxFio.Top+.boxFio.Height+30,;
     RetTxtWidth('WзаписатьW'),dHeight+5,'записать','DO saveChangeJob' 
     *---------------------------------Кнопка возврат--------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','butExit',.butSave.Left+.butSave.Width+30,.butSave.Top,.butSave.Width,dHeight+5,'возврат','fSupl.Release'                       
     .Height=.boxFio.Height+.butSave.Height+70
     
     DO addListBoxMy WITH 'fSupl',1,.boxFio.Left,.boxFio.Top+.boxFio.Height-1,300,.boxFio.Width
     WITH .listBox1
          .RowSource='curFindFio.fio'           
          .RowSourceType=6
          .ColumnCount=1   
*          .ColumnWidths='280,420'      
          .Visible=.F.  
          .Height=.Parent.Height-.Top-5
          .procForValid='DO validListBoxFio'
          .procForLostFocus='DO lostFocusListBoxFio'       
     ENDWITH
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************
PROCEDURE changeFio
WITH fSupl
     .listBox1.Visible=.T.         
ENDWITH 
Local lcFind
lcFind=fSupl.boxFio.Text 
SELECT curFindFio
ZAP
APPEND FROM people FOR LOWER(ALLTRIM(lcFind))=LEFT(LOWER(fio),LEN(ALLTRIM(lcFind)))
GO TOP
IF RECCOUNT('curFindFio')#0.AND.LEN(ALLTRIM(lcFind))>2
   WITH fSupl.listBox1 
        .RowSource='curFindFio.fio'      
   ENDWITH 
ELSE
   fSupl.listBox1.Visible=.F.   
ENDIF 

**************************************************************************
PROCEDURE validListBoxFio
WITH fSupl
     .listBox1.Visible=.F.
     newFio=curFindFio.fio
     newKodPeop=curFindFio.num
     newTabPeop=curFindFio.tabn
     newvac=curFindFio.vac
     .butSave.SetFocus   
     .Refresh
ENDWITH
**************************************************************************
PROCEDURE lostFocusListBoxFio
WITH fSupl
     .listBox1.Visible=.F. 
ENDWITH
**************************************************************************
PROCEDURE keyPressFio
IF LASTKEY()=9.AND.fSupl.listBox1.Visible=.T.
   fSupl.listBox1.SetFocus
ENDIF
****************************************************************************
PROCEDURE saveChangeJob
fSupl.Release
IF newKodPeop#0
   SELECT datJob
   REPLACE kodpeop WITH newKodPeop,vac WITH newvac,tabn WITH newTabPeop
   SELECT people
   LOCATE FOR num=newKodPeop
   frmTop.grdPers.Columns(frmTop.grdPers.Columncount).SetFocus 
   frmTop.Refresh
ENDIF
**************************************************************************************************************************
PROCEDURE formForSearsh
fPoisk=CREATEOBJECT('FORMMY')
WITH fPoisk
     .BackColor=RGB(255,255,255)
     DO addShape WITH 'fPoisk',1,10,10,dHeight,300,8     
     .logExit=.T.  
      find_ch=''
      DO adLabMy WITH 'fpoisk',1,'код или ФИО сотрудника' ,fpoisk.Shape1.Top+10,fpoisk.Shape1.Left+10,250,2
      DO addtxtboxmy WITH 'fpoisk',1,fpoisk.Shape1.Left+10,fpoisk.Shape1.Top+fpoisk.lab1.Height+10,280,.F.,'find_ch'
      .Shape1.Height=.lab1.Height+.txtBox1.Height+30
      fpoisk.txtBox1.procForkeyPress='DO keyPressFind'
      DO addcontlabel WITH 'fpoisk','cont1',fpoisk.Shape1.Left+(.Shape1.Width-RetTxtWidth('wОтменаw')*2-20)/2,fpoisk.Shape1.Top+fpoisk.Shape1.Height+10,;
      RetTxtWidth('wОтменаw'),dHeight+3,'Поиск','DO searshCard'
      DO addcontlabel WITH 'fpoisk','cont2',fpoisk.Cont1.Left+fpoisk.Cont1.Width+20,fpoisk.Cont1.Top,;
      fpoisk.Cont1.Width,dHeight+3,'Отмена','Fpoisk.Release'
     .Caption='Поиск'
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+30+.lab1.Height
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fpoisk'
fpoisk.Show
*************************************************************************************************************************
*                Непосредственно поиск личной карточки
*************************************************************************************************************************
PROCEDURE SearshCard
IF EMPTY(find_ch)
   RETURN
ENDIF
find_ch=ALLTRIM(find_ch)
SELECT people
oldrec=RECNO()
log_ord=SYS(21)
IF TYPE(find_ch)='N' 
   SET ORDER TO 1
   IF SEEK(VAL(find_ch))
      fPoisk.Release
   ELSE 
      find_ch=''
      fPoisk.Refresh
      SET ORDER TO &log_ord
      GO oldrec
      RETURN      
   ENDIF         
ELSE   
   SET ORDER TO 2
   DO unosimbol WITH 'find_ch'  
   IF SEEK(find_ch)
      fPoisk.Release
   ELSE 
      find_ch=''
      fPoisk.Refresh
      SET ORDER TO &log_ord
      GO oldrec
      RETURN   
   ENDIF     
ENDIF
frmTop.grdPers.Columns(frmTop.grdPers.Columncount).SetFocus 
frmTop.Refresh
************************************************************************************************************************
PROCEDURE keyPressFind
DO CASE
   CASE LASTKEY()=27
        fpoisk.Release
   CASE LASTKEY()=13
        find_ch=fpoisk.TxtBox1.Value           
        DO searshCard  
ENDCASE 
********************************************************************************************************************************************************
*                                                      Удаление сотрудника
*********************************************************************************************************************************************************
PROCEDURE formDeletePeople
fSupl=CREATEOBJECT('FORMSUPL')
log_del=.F.
WITH fSupl   
     .Caption='Удаление сотрудника'
     DO addShape WITH 'fSupl',1,10,10,dHeight,400,8 
     DO adLabMy WITH 'fSupl',1,ALLTRIM(STR(people.num))+'  '+ALLTRIM(people.Fio) ,.Shape1.Top+20,.Shape1.Left+10,.Shape1.Width-20,2   
     DO adLabMy WITH 'fSupl',2,'Для подтверждения намерений поставьте',.lab1.Top+.lab1.Height+10,.lab1.Left,.lab1.Width,2
     DO adLabMy WITH 'fSupl',3,'птичку в окошке "подтверждение намерений"',.lab2.Top+.lab2.Height,.lab1.Left,.lab1.Width,2
     .Shape1.Height=.lab1.Height*3+40
      DO adCheckBox WITH 'fSupl','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'log_del',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wУдалитьw')*2-20)/2,.check1.Top+.check1.Height+20,;
     RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','DO deletePeople'     
    
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fSupl.Release'
     .Height=.Shape1.Height+.check1.Height+.cont1.Height+60
     .Width=.Shape1.Width+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
*             Непосредственно удаление сотрудника и записей, связанных с ним
*************************************************************************************************************************
PROCEDURE deletePeople
IF !log_del
   RETURN
ENDIF
SELECT datjob
DELETE FOR kodpeop=people.num
SELECT people
DELETE
IF EOF()
   GO BOTTOM
ENDIF
fSupl.Release
frmTop.grdPers.Columns(frmTop.grdPers.Columncount).SetFocus 
frmTop.Refresh
*-------------------------------------------------------------------------------------------------------------------------
*                    Процедуры для справочнго материала
*-------------------------------------------------------------------------------------------------------------------------
*********************************************************************************************************************************************************
*                                                   Справочник подразделений
*********************************************************************************************************************************************************
PROCEDURE procpodr
SELECT sprpodr
oldOrdPodr=SYS(21)
SET ORDER TO 3
GO TOP
fdolj=CREATEOBJECT('Formspr')
namenew=''
namernew=''
primnew=''
namernewold=''
log_ap=.F.
WITH fdolj    
     .Caption='Справочник подразделений'  
     .ProcExit='DO exitFromProcPodr'  
     DO addButtonOne WITH 'fDolj','menuCont1',10,5,'редакция','pencil.ico',"Do readspr WITH 'fdolj','Do readpodr WITH .F.'",39,RetTxtWidth('удаление')+44,'редакция'  
     DO addButtonOne WITH 'fDolj','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'возврат','undo.ico','DO exitFromProcPodr',39,.menucont1.Width,'возврат'             
     DO addmenureadspr WITH 'fdolj','DO writePodr WITH .T.','DO writePodr WITH .F.' 
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Parent.menucont1.Height-5                 
          .RecordSourceType=1     
          .RecordSource='sprpodr'
           DO addColumnToGrid WITH 'fDolj.fGrid',5
          .Column1.ControlSource='sprpodr.kod'
          .Column2.ControlSource='sprpodr.np'
          .Column3.ControlSource='" "+sprpodr.name'
          .Column4.ControlSource='" "+sprpodr.namework' 
                       
          .Column1.Width=RettxtWidth(' 1234 ')
          .Column2.Width=.Column1.Width                  
          .Column3.Width=(.Width-.Column1.Width-.Column1.Width)/2         
          .Column4.Width=.Width-.column1.width-.Column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount    
          .Columns(.ColumnCount).Width=0               
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='№'
          .Column3.Header1.Caption='Наименование'
          .Column4.Header1.Caption='Примечание'                       
          .Column1.Alignment=1
          .Column2.Alignment=1           
          .Column3.Alignment=0                  
          .colNesInf=2      
          .Visible=.T.         
     ENDWITH
     DO gridSizeNew WITH 'fdolj','fGrid','shapeingrid'     
     .fGrid.Column1.Text1.ToolTipText='xccvzsvzxcv'   
     DO addtxtboxmy WITH 'fdolj',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
     .txtbox1.Enabled=.F.
     DO addtxtboxmy WITH 'fdolj',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,1
     .txtbox2.Enabled=.F.
     DO addtxtboxmy WITH 'fdolj',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fdolj',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')
     DO addcontmy WITH 'fdolj','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont1','sprpodr',1"
     DO addcontmy WITH 'fdolj','cont2',.cont1.Left+.fGrid.Column1.Width+2,.fGrid.Top+2,.fGrid.Column2.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprpodr',3"
     .cont2.SpecialEffect=1   
     DO addcontmy WITH 'fdolj','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprpodr',2" 
     DO addcontmy WITH 'fdolj','cont4',.cont3.Left+.fGrid.Column3.Width+1,.fGrid.Top+2,.fGrid.Column4.Width-4,.fGrid.HeaderHeight-3,'' 
     SELECT sprpodr  
     GO TOP 
     .Show
ENDWITH
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readpodr
PARAMETERS parlog
SELECT sprpodr
log_ap=.F.
WITH fDolj
     IF parlog
        .fGrid.GridMyAppendBlank(2,'kod','name')   
        log_ap=.T.       
     ENDIF
     .fGrid.columns(.fGrid.columnCount).SetFocus
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1) 
     .nrec=RECNO()
     nameNew=sprpodr.name     
     primNew=sprpodr.namework
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .txtBox3.Left=.txtBox2.Left+.txtBox2.Width-1   
     .txtBox4.Left=.txtBox3.Left+.txtBox3.Width-1   
     .txtbox1.ControlSource='sprpodr.kod'
     .txtbox2.ControlSource='sprpodr.np'
     .txtbox3.ControlSource='nameNew'
     .txtbox4.ControlSource='primNew'   
        
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .SetAll('Visible',.T.,'MyTxtBox')
     .fGrid.Enabled=.F.
     .Refresh
     .txtbox3.SetFocus      
ENDWITH      
IF parlog
   KEYBOARD '{TAB}'
ENDIF   
************************************************************************************************************************
PROCEDURE writepodr
PARAMETERS par_log
WITH fDolj
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     SELECT sprpodr
     IF par_log  
        REPLACE name WITH namenew,namework WITH primnew    
     ELSE   
        IF log_ap
           DELETE
        ENDIF
     ENDIF    
     .SetAll('Visible',.F.,'MyTxtBox')
     .fGrid.Enabled=.T.
     SELECT sprpodr
     .fGrid.GridUpdate
     GO .nrec
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     GO .nrec
ENDWITH      
*************************************************************************************************************************
PROCEDURE delpodr
fdolj.Setall('BorderWidth',0,'Mycontmenu')
IF SEEK(STR(sprpodr.kod,3),'rasp',1) 
   fdolj.fGrid.GridNoDelRec   
ELSE 
  fdolj.fGrid.GridDelRec('fdolj.fGrid','sprpodr') 
ENDIF   
**************************************************************************************************************************
PROCEDURE exitFromProcPodr
SELECT sprpodr
SET ORDER TO &oldOrdPodr
fDolj.Visible=.F.
fDolj.Release
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник должностей
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE procdolj
STORE 0 TO kodNew,katNew,trNew,trNew1,trNew2,trNew3,oldRec
STORE '' TO namenew,nameKfNew,nameKfNew1,nameKfNew2,nameKfNew3,strKat
log_ap=.F.
SELECT * FROM sprkat INTO CURSOR curSupKat
SELECT sprdolj
SET ORDER TO 1
SET RELATION TO kat INTO sprkat ADDITIVE 
GO TOP
fdolj=CREATEOBJECT('Formspr')
WITH fdolj    
     .Caption='Справочник должностей'  
     .ProcExit='DO procOutSprDolj'    
     DO addButtonOne WITH 'fDolj','menuCont1',10,5,'новая','pencila.ico','Do readDolj WITH .T.',39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fDolj','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico','Do readDolj WITH .F.',39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fDolj','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO deldolj',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fDolj','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico',"DO printreport WITH 'repdolj','справочник должностей','sprdolj'",39,.menucont1.Width,'печать'   
     DO addButtonOne WITH 'fDolj','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'возврат','undo.ico','DO procOutSprDolj',39,.menucont1.Width,'возврат'     
      WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Parent.menucont1.Height-5       
          .RecordSourceType=1     
          .RecordSource='sprdolj'
          DO addColumnToGrid WITH 'fDolj.fGrid',5 
          .Column1.ControlSource='sprdolj.kod'
          .Column2.ControlSource='" "+sprdolj.namework'
          .Column3.ControlSource='" "+sprdolj.name'
          .Column4.ControlSource='" "+sprkat.name'     
          .Column1.Width=RettxtWidth(' 1234 ')     
          .Column4.Width=RettxtWidth(' СРЕДНИЙ МЕДПЕРСОНАЛ ')  
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование сокращенное (рабочее)'
          .Column3.Header1.Caption='Наименование полное'
          .Column4.Header1.Caption='Персонал'    
          .Columns(.ColumnCount).Width=0
          .Column2.Width=(.Width-.column1.width-.column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount)/2
          .Column3.Width=.Width-.column1.width-.column2.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Movable=.F. 
          .Column1.Alignment=1           
          .Column4.Alignment=0
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.         
     ENDWITH
     DO gridSizeNew WITH 'fdolj','fGrid','shapeingrid'   
          

     DO addcontmy WITH 'fdolj','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont1','sprdolj',1,4"
     .cont1.SpecialEffect=1   
     DO addcontmy WITH 'fdolj','cont2',.cont1.Left+.fGrid.Column1.Width+2,.fGrid.Top+2,.fGrid.Column2.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprdolj',2,5"
     DO addcontmy WITH 'fdolj','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-4,.fGrid.HeaderHeight-3,'' 
     DO addcontmy WITH 'fdolj','cont4',.cont3.Left+.fGrid.Column3.Width+1,.fGrid.Top+2,.fGrid.Column4.Width-4,.fGrid.HeaderHeight-3,'' 
     WITH .fGrid
          FOR ch=1 TO 3
              obj_ch='.Column'+LTRIM(STR(ch))+'.Header1'          
              obj_col='fdolj.cont'+LTRIM(STR(ch))
              &obj_ch..FontSize=IIF(FONTMETRIC(6,dFontName,dFontSize)* TXTWIDTH(&obj_ch..Caption,dFontName,dFontSize)>&obj_col..Width,dFontSize-1,dFontSize)
          ENDFOR
     ENDWITH 
     SELECT sprdolj  
ENDWITH
fDolj.SHow

*************************************************************************************************************************
PROCEDURE readDolj
PARAMETERS par1
SELECT sprdolj
oldRec=RECNO()
log_ap=par1
kodNew=kod
IF log_ap
   oldOrd=SYS(21)
   SET ORDER TO 1
   GO BOTTOM
   kodNew=kod+1
   SET ORDER TO &oldOrd 
ENDIF
nameNew=IIF(log_ap,'',sprdolj.name)
nameWNew=IIF(log_ap,'',sprdolj.namework)
trNew=IIF(log_ap,0,sprdolj.kf)
namekfNew=IIF(log_ap,0.000,sprdolj.namekf)
trNew1=IIF(log_ap,0,sprdolj.kf1)
namekfNew1=IIF(log_ap,0.000,sprdolj.namekf1)
trNew2=IIF(log_ap,0,sprdolj.kf2)
namekfNew2=IIF(log_ap,0.000,sprdolj.namekf2)
trNew3=IIF(log_ap,0,sprdolj.kf3)
namekfNew3=IIF(log_ap,0.000,sprdolj.namekf3)
katNew=IIF(log_ap,0,sprdolj.kat)
strKat=IIF(SEEK(katNew,'sprkat',1),sprkat.name,'')
namemNew=IIF(log_ap,'',sprdolj.namem)
logSexNew=IIF(log_ap,.F.,sprdolj.logSex)
lAvtPers=.T.
lAvtVac=.T.

fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption=IIF(log_ap,'Новая запись','Редактирование')+' ('+LTRIM(STR(kodnew))+')'    
     DO adtBoxAsCont WITH 'fSupl','contNameW',10,10,400,dHeight,'наименование сокращенное (рабочее)',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxNw',.contNameW.Top+.contNameW.Height-1,.contNameW.Left,.contNameW.Width,dHeight,'nameWNew',.F.,.T.,.F.,.F.   
     
     DO adtBoxAsCont WITH 'fSupl','contKat',.contNameW.Left+.contNameW.Width-1,.contNameW.Top,RetTxtWidth('средний медицинский персоналw'),dHeight,'персонал',2,1           
     DO addComboMy WITH 'fSupl',2,.contKat.Left,.txtBoxNw.Top,dHeight,.contKat.Width,.T.,'strKat','curSupKat.name',6,.F.,'katnew=curSupKat.kod',.F.,.T.    
          
     DO adtBoxAsCont WITH 'fSupl','contDol',.contNameW.Left,.txtBoxNw.Top+.txtBoxNw.Height-1,.contNamew.Width+.contKat.Width-1,dHeight,'наименование должности - полное',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxDol',.contDol.Top+.contDol.Height-1,.contDol.Left,.contDol.Width,dHeight,'nameNew',.F.,.T.,.F.,.F.                           
     
     DO adtBoxAsCont WITH 'fSupl','contNameM',.contNameW.Left,.txtBoxDol.Top+.txtBoxDol.Height-1,.contDol.Width-RetTxtWidth('W!W'),dHeight,'наименование для мужчин',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxM',.contNameM.Top+.contNameM.Height-1,.contNameM.Left,.contNameM.Width,dHeight,'namemNew',.F.,.T.,.F.,.F.   
     
     DO adtBoxAsCont WITH 'fSupl','contSex',.contNameM.Left+.contNameM.Width-1,.contNameM.Top,RetTxtWidth('W!W')+1,dHeight,'!',2,1
     DO adtBoxNew WITH 'fSupl','txtBoxSex',.txtBoxM.Top,.contSex.Left,.contSex.Width,dHeight,'',.F.,.T.,.F.,.F.    
     .txtBoxSex.Enabled=.F.        
     DO adCheckBox WITH 'fSupl','check1','',.txtBoxM.Top+2,.txtBoxSex.Left+5,.contSex.Width,dHeight,'logSexNew',0
     .check1.Left=.txtBoxSex.Left+(.txtBoxSex.Width-.check1.Width)/2+1

     DO adtBoxAsCont WITH 'fSupl','contDkat',10,.txtBoxM.Top+.txtBoxM.Height+20,RetTxtWidth('WВысшая категорияW'),dHeight,'без категории',2,1     
     DO adtBoxAsCont WITH 'fSupl','contDkat1',.contDkat.Left+.contDkat.Width-10,.contDkat.Top,.contDkat.Width,dHeight,'вторая категория',2,1     
     DO adtBoxAsCont WITH 'fSupl','contDkat2',.contDkat1.Left+.contDkat1.Width-10,.contDkat.Top,.contDkat.Width,dHeight,'первая категория',2,1     
     DO adtBoxAsCont WITH 'fSupl','contDkat3',.contDkat2.Left+.contDkat2.Width-10,.contDkat.Top,.contDkat.Width,dHeight,'высшая категория',2,1             
                  
     .Width=.contDol.Width+20
        
     .contDkat.Left=(.Width-.contDkat.Width*4+3)/2
     .contDkat1.Left=.contDkat.Left+.contDkat.Width-1
     .contDkat2.Left=.contDkat1.Left+.contDkat1.Width-1
     .contDkat3.Left=.contDkat2.Left+.contDkat2.Width-1        
       
     DO adtBoxAsCont WITH 'fSupl','contTr',.contDkat.Left,.contDKat.Top+.contDKat.Height-1,RetTxtWidth('Wразряд'),dHeight,'разряд',2,1     
     DO adtBoxAsCont WITH 'fSupl','contKft',.contTr.Left+.contTr.Width-1,.contTr.Top,.contDkat.Width-.contTr.Width+1,dHeight,'кфт',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxTr',.contTr.Top+.contTr.Height-1,.contTr.Left,.contTr.Width,dHeight,'trNew','Z',.T.,.F.,.F.,'DO validTrNew WITH 1'                           
     DO adtBoxNew WITH 'fSupl','txtBoxKf',.txtBoxTr.Top,.contKft.Left,.contKft.Width,dHeight,'nameKfNew','Z',.T.,.F.,.F.
         
     DO adtBoxAsCont WITH 'fSupl','contTr1',.contDkat1.Left,.contTr.Top,.contTr.Width,dHeight,'разряд',2,1     
     DO adtBoxAsCont WITH 'fSupl','contKft1',.contTr1.Left+.contTr1.Width-1,.contTr.Top,.contKft.Width,dHeight,'кфт',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxTr1',.txtBoxTr.Top,.contTr1.Left,.txtBoxTr.Width,dHeight,'trNew1','Z',.T.,.F.,.F.,'DO validTrNew WITH 2'                           
     DO adtBoxNew WITH 'fSupl','txtBoxKf1',.txtBoxTr.Top,.contKft1.Left,.contKft1.Width,dHeight,'nameKfNew1','Z',.T.,.F.,.F.
         
     DO adtBoxAsCont WITH 'fSupl','contTr2',.contDkat2.Left,.contTr.Top,.contTr.Width,dHeight,'разряд',2,1     
     DO adtBoxAsCont WITH 'fSupl','contKft2',.contTr2.Left+.contTr2.Width-1,.contTr.Top,.contKft.Width,dHeight,'кфт',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxTr2',.txtBoxTr.Top,.contTr2.Left,.txtBoxTr.Width,dHeight,'trNew2','Z',.T.,.F.,.F.,'DO validTrNew WITH 3'                           
     DO adtBoxNew WITH 'fSupl','txtBoxKf2',.txtBoxTr.Top,.contKft2.Left,.contKft2.Width,dHeight,'nameKfNew2','Z',.T.,.F.,.F.
         
     DO adtBoxAsCont WITH 'fSupl','contTr3',.contDkat3.Left,.contTr.Top,.contTr.Width,dHeight,'разряд',2,1     
     DO adtBoxAsCont WITH 'fSupl','contKft3',.contTr3.Left+.contTr3.Width-1,.contTr.Top,.contKft.Width,dHeight,'кфт',2,1     
     DO adtBoxNew WITH 'fSupl','txtBoxTr3',.txtBoxTr.Top,.contTr3.Left,.txtBoxTr.Width,dHeight,'trNew3','Z',.T.,.F.,.F.,'DO validTrNew WITH 4'                           
     DO adtBoxNew WITH 'fSupl','txtBoxKf3',.txtBoxTr.Top,.contKft3.Left,.contKft3.Width,dHeight,'nameKfNew3','Z',.T.,.F.,.F.  
     
     DO adCheckBox WITH 'fSupl','checkPers','замена для персонала',.txtBoxTr.Top+.txtBoxTr.Height+10,10,150,dHeight,'lAvtPers',0,.T.,.F.
     DO adCheckBox WITH 'fSupl','checkVac','замена для вакансий',.checkPers.Top,10,150,dHeight,'lAvtVac',0,.T.,.F.
     .checkPers.Left=(.Width-.checkPers.Width-.checkVac.Width-10)/2
     .checkVac.Left=.checkPers.Left+.checkPers.Width+10
        
     *-----------------------------Кнопка приступить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',(.Width-(RetTxtWidth('wзаписатьw')*2)-30)/2,.checkPers.Top+.checkPers.Height+10,RetTxtWidth('wзаписатьw'),dHeight+5,'записать','DO writeDolj WITH .T.'

     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'возврат','DO writeDolj WITH .F.','возврат' 
    .Height=dHeight*8+.cont1.Height+.checkPers.Height+80
    .Autocenter=.T.    
    .WindowState=0
ENDWITH
fSupl.Show
*************************************************************************************************************************
PROCEDURE validTrNew
PARAMETERS par1
DO CASE 
   CASE par1=1
        nameKfNew=IIF(SEEK(trNew,'sprkoef',1),sprkoef.name,namekfNew)
   CASE par1=2
        nameKfNew1=IIF(SEEK(trNew1,'sprkoef',1),sprkoef.name,namekfNew1)
   CASE par1=3
        nameKfNew2=IIF(SEEK(trNew2,'sprkoef',1),sprkoef.name,namekfNew2)
   CASE par1=4
        nameKfNew3=IIF(SEEK(trNew3,'sprkoef',1),sprkoef.name,namekfNew3)
ENDCASE 
*************************************************************************************************************************
PROCEDURE procvalidkat
SELECT sprdolj
katNew=curSupKat.kod
KEYBOARD '{TAB}'    
************************************************************************************************************************
PROCEDURE procgotkat
SELECT curSupKat
LOCATE FOR kod=sprkat->kod
nrec=RECNO()
GO TOP 
COUNT WHILE RECNO()#nrec TO varnrec
fdolj.combobox3.DisplayCount=MAX(fdolj.fGrid.RelativeRow,fdolj.fGrid.RowsGrid-fdolj.fGrid.RelativeRow)
fdolj.combobox3.DisplayCount=MIN(fdolj.combobox3.DisplayCount,RECCOUNT())
SELECT sprdolj
************************************************************************************************************************
PROCEDURE writedolj
PARAMETERS par_log
IF par_log
   SELECT sprdolj
   IF log_ap
      APPEND BLANK
      REPLACE kod WITH kodNew
      oldRec=RECNO()
   ENDIF
   REPLACE name WITH nameNew,kat WITH katNew,kf WITH trNew,nameKf WITH nameKfNew,kf1 WITH trNew1,nameKf1 WITH nameKfNew1,kf2 WITH trNew2,nameKf2 WITH nameKfNew2,kf3 WITH trNew3,nameKf3 WITH nameKfNew3,;
           namework WITH nameWNew,namem WITH namemNew,logSex WITH logSexNew
   IF lAvtVac
      SELECT rasp
      oldtag=SYS(21)
      SET ORDER TO 3
      SEEK sprdolj.kod
      SCAN WHILE kd=sprdolj.kod           
           DO CASE 
              CASE kv=1
                   REPLACE kfvac WITH IIF(sprdolj.kf3#0,sprdolj.kf3,sprdolj.kf),nkfvac WITH IIF(sprdolj.namekf3#0,sprdolj.namekf3,sprdolj.namekf)
              CASE kv=2
                   REPLACE kfvac WITH IIF(sprdolj.kf2#0,sprdolj.kf2,sprdolj.kf),nkfvac WITH IIF(sprdolj.namekf2#0,sprdolj.namekf2,sprdolj.namekf)
              CASE kv=3                 
                   REPLACE kfvac WITH IIF(sprdolj.kf1#0,sprdolj.kf1,sprdolj.kf),nkfvac WITH IIF(sprdolj.namekf1#0,sprdolj.namekf1,sprdolj.namekf)
              OTHERWISE
                   REPLACE kfvac WITH sprdolj.kf,nkfvac WITH sprdolj.namekf
           ENDCASE            
      ENDSCAN       
      SET ORDER TO &oldtag
   ENDIF        
   IF lAvtPers
      SELECT datjob
      SET FILTER TO kd=sprdolj.kod
      SCAN ALL
           DO CASE 
              CASE kv=1
                   REPLACE kf WITH IIF(sprdolj.kf3#0,sprdolj.kf3,sprdolj.kf),namekf WITH IIF(sprdolj.namekf3#0,sprdolj.namekf3,sprdolj.namekf)
              CASE kv=2
                   REPLACE kf WITH IIF(sprdolj.kf2#0,sprdolj.kf2,sprdolj.kf),namekf WITH IIF(sprdolj.namekf2#0,sprdolj.namekf2,sprdolj.namekf)
              CASE kv=3                 
                   REPLACE kf WITH IIF(sprdolj.kf1#0,sprdolj.kf1,sprdolj.kf),namekf WITH IIF(sprdolj.namekf1#0,sprdolj.namekf1,sprdolj.namekf)
              OTHERWISE
                   REPLACE kf WITH sprdolj.kf,namekf WITH sprdolj.namekf
           ENDCASE            
      ENDSCAN       
      SET FILTER TO 
   ENDIF     
   SELECT sprdolj        
ENDIF 
GO oldRec
fSupl.Release
fDolj.Refresh
*************************************************************************************************************************
PROCEDURE deldolj
SELECT sprdolj
SELECT rasp
LOCATE FOR kd=sprdolj.kod
IF FOUND()
   SELECT sprdolj
   fdolj.fGrid.GridNoDelRec   
ELSE 
  SELECT sprdolj
  fdolj.fGrid.GridDelRec('fdolj.fGrid','sprdolj') 
ENDIF   
**************************************************************************************************************************
PROCEDURE procOutSprDolj
fDolj.Release
SELECT * FROM sprdolj INTO CURSOR curSprDolj READWRITE ORDER BY name
SELECT sprdolj
SET RELATION TO 
SET ORDER TO 1
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник категорий персонала
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE prockat
SELECT sprkat
SET ORDER TO 1
GO TOP
fkat=CREATEOBJECT('Formspr')
WITH fkat   
     .Caption='Справочник производственных категорий персонала' 
     .ProcExit='fkat.fGrid.GridReturn'  
     DO addButtonOne WITH 'fKat','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fKat','Do readkat WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fKat','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fKat','Do readkat WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fKat','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delkat',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fKat','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','fKat.fGrid.GridReturn',39,.menucont1.Width,'возврат'                       
     
     DO addmenureadspr WITH 'fkat',"DO writeSprNew WITH 'fkat','fkat.fGrid','sprkat'","DO exitWriteSpr WITH 'fkat','fkat.fGrid'"
     
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5                    
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fKat.fGrid',5
          .RecordSource='sprkat'
          .Column1.ControlSource='sprkat.kod'
          .Column2.ControlSource='" "+sprkat.name'    
          .Column3.ControlSource='" "+sprkat.namefull'
          .Column4.ControlSource='" "+sprkat.namefull1'
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование'
          .Column3.Header1.Caption='Для ведомостей'
          .Column4.Header1.Caption='Для штатного'     
          .Column1.Width=RettxtWidth(' 1234 ')
          .Column2.Width=(.Width-.column1.Width)/3 
          .Column3.Width=.Column2.Width   
          .Column4.Width=.Width-.column1.Width-.column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0  
          .Column1.Alignment=1    
          .colNesInf=2   
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')         
     ENDWITH  
     DO gridSizeNew WITH 'fkat','fGrid','shapeingrid'   
     DO addtxtboxmy WITH 'fkat',1,1,1,fkat.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fkat',2,1,1,fkat.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fkat',3,1,1,fkat.fGrid.Column3.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fkat',4,1,1,fkat.fGrid.Column4.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
     DO addcontmy WITH 'fkat','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,''
     .cont1.SpecialEffect=1          
     DO addcontmy WITH 'fkat','cont2',.cont1.Left+.fGrid.Column1.Width+1,.fGrid.Top+2,.fGrid.Column2.Width-3,.fGrid.HeaderHeight-3,''
     DO addcontmy WITH 'fkat','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-3,.fGrid.HeaderHeight-3,''
     DO addcontmy WITH 'fkat','cont4',.cont3.Left+.fGrid.Column3.Width+1,.fGrid.Top+2,.fGrid.Column4.Width-3,.fGrid.HeaderHeight-3,''
ENDWITH
fkat.Show
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readkat
PARAMETERS parlog
SELECT sprkat
IF parlog
   fkat.fGrid.GridMyAppendBlank(1,'kod','name')   
ENDIF
fkat.SetAll('Visible',.T.,'MyTxtBox')
fKat.fGrid.columns(fKat.fGrid.columnCount).SetFocus
lineTop=fkat.fGrid.Top+fkat.fGrid.HeaderHeight+fkat.fGrid.RowHeight*(IIF(fkat.fGrid.RelativeRow<=0,1,fkat.fGrid.RelativeRow)-1)
fkat.nrec=RECNO()
SCATTER TO fkat.dim_ap
fkat.txtBox1.Left=fkat.fGrid.Left+10
fkat.txtBox2.Left=fkat.txtbox1.Left+fkat.txtbox1.Width-1
fkat.txtBox3.Left=fkat.txtbox2.Left+fkat.txtbox2.Width-1
fkat.txtBox4.Left=fkat.txtbox3.Left+fkat.txtbox2.Width-1
fkat.txtbox1.ControlSource='fkat.dim_ap(1)'
fkat.txtbox2.ControlSource='fkat.dim_ap(2)'
fkat.txtbox3.ControlSource='fkat.dim_ap(3)'
fkat.txtbox4.ControlSource='fkat.dim_ap(4)'

fkat.SetAll('Top',linetop,'MyTxtBox')
fkat.SetAll('Height',fkat.fGrid.RowHeight+1,'MyTxtBox')
fkat.SetAll('BackStyle',1,'MyTxtBox')
fkat.txtbox1.Enabled=.F.
fkat.fGrid.Enabled=.F.
fkat.txtbox2.SetFocus
*************************************************************************************************************************
PROCEDURE delkat
SELECT rasp
LOCATE FOR kat=sprkat->kod
IF FOUND()
   SELECT sprkat
   fkat.fGrid.GridNoDelRec 
ELSE 
   SELECT sprkat
   fkat.fGrid.GridDelRec('fkat.fGrid','sprkat')
ENDIF   
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник квалификационных категорий персонала
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE prockval
fkval=CREATEOBJECT('Formspr')
SELECT sprkval
SET ORDER TO 1
GO TOP
WITH fkval        
     .Caption='Справочник квалификационных категорий персонала'
     .ProcExit='DO exitProcKval'   
     .AddProperty('doplkatold',0)    
     DO addButtonOne WITH 'fKval','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fkval','Do readkval WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fKval','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fkval','Do readkval WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fKval','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delkval',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fKval','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO exitProcKval',39,.menucont1.Width,'возврат'             
     DO addmenureadspr WITH 'fkval',"DO writeSprNew WITH 'fkval','fkval.fGrid','sprkval','reppersforkat'","DO exitWriteSpr WITH 'fkval','fkval.fGrid'"
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5               
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fKval.fGrid',4 
          .RecordSource='sprkval'
          .Column1.ControlSource='sprkval.kod'
          .Column2.ControlSource='" "+sprkval.name'    
          .Column3.ControlSource='sprkval.doplkat'
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование квалификационной категории'
          .Column3.Header1.Caption='%'
          .Column1.Width=RettxtWidth(' 1234 ') 
          .Column3.Width=RettxtWidth(' 1234 ')   
          .Column4.Width=0
          .Column2.Width=.Width-.column1.Width-.Column3.Width-SYSMETRIC(5)-13-4               
          .Column3.Header1.Caption='%'
          .Column3.Format='Z'
          .Column1.Alignment=1
          .Column3.Alignment=1
          .colNesInf=2    
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')        
     ENDWITH 
     DO gridSizeNew WITH 'fkval','fGrid','shapeingrid'   
     DO addtxtboxmy WITH 'fkval',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fkval',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fkval',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,1     
     .SetAll('Visible',.F.,'MyTxtBox')  
     DO addcontmy WITH 'fkval','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fkval','fkval.cont1','sprkval',1"
     .cont1.SpecialEffect=1        
     DO addcontmy WITH 'fkval','cont2',.cont1.Left+.fGrid.Column1.Width+1,.fGrid.Top+2,.fGrid.Column2.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fkval','fkval.cont2','sprkval',2"
     DO addcontmy WITH 'fkval','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-3,.fGrid.HeaderHeight-3,''
ENDWITH
fkval.Show
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readkval
PARAMETERS parlog
SELECT sprkval
IF parlog
   fkval.fGrid.GridMyAppendBlank(1,'kod','name')   
ENDIF
fkval.SetAll('Visible',.T.,'MyTxtBox')
fKval.fGrid.columns(fKval.fGrid.columnCount).SetFocus
lineTop=fkval.fGrid.Top+fkval.fGrid.HeaderHeight+fkval.fGrid.RowHeight*(IIF(fkval.fGrid.RelativeRow<=0,1,fkval.fGrid.RelativeRow)-1)
fkval.nrec=RECNO()
SCATTER TO fkval.dim_ap
fkval.txtBox1.Left=fkval.fGrid.Left+10
fkval.txtBox2.Left=fkval.txtbox1.Left+fkval.txtbox1.Width-1
fkval.txtBox3.Left=fkval.txtbox2.Left+fkval.txtbox2.Width-1
fkval.txtbox1.ControlSource='fkval.dim_ap(1)'
fkval.txtbox2.ControlSource='fkval.dim_ap(2)'
fkval.txtbox3.ControlSource='fkval.dim_ap(4)'
fkval.SetAll('Top',linetop,'MyTxtBox')
fkval.SetAll('Height',fkval.fGrid.RowHeight+1,'MyTxtBox')
fkval.SetAll('BackStyle',1,'MyTxtBox')
fkval.txtbox1.Enabled=.F.
fkval.fGrid.Enabled=.F.
fkval.txtbox2.SetFocus
*************************************************************************************************************************
PROCEDURE delkval
SELECT datjob
SET FILTER TO 
LOCATE FOR kv=sprkval.kod
IF FOUND() 
   SELECT sprkval
   fkval.fGrid.GridNoDelRec  
ELSE 
   SELECT sprkval
   fkval.fGrid.GridDelRec('fkval.fGrid','sprkval')  
ENDIF   
*************************************************************************************************************************
*             Процедура замены доплаты за категорию в персонале при изменении % в спр-ке
*************************************************************************************************************************
PROCEDURE reppersforkat
SELECT datjob
SET FILTER TO
REPLACE pkat WITH sprkval.doplkat FOR kv=sprkval.kod
SELECT sprkval
*************************************************************************************************************************
PROCEDURE exitProcKval
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
SELECT sprkval
fkval.fGrid.GridReturn
*************************************************************************************************************************
*                                       Тарификационная комиссия
*************************************************************************************************************************
PROCEDURE procboss
newOffice=datshtat.office
newAdres=datshtat.adres
newBoss=datshtat.boss
SELECT boss
GO TOP
fboss=CREATEOBJECT('Formspr')
WITH fboss  
     .Caption='Тарификационная комиссия'    
     .procExit='DO returnboss'    
     DO addButtonOne WITH 'fBoss','menureadrekv',10,5,'записать','pencil.ico','DO writerekv WITH .T.',39,RetTxtWidth('удаление')+44,'записать'  
     DO addButtonOne WITH 'fBoss','menuexitrekv',.menureadrekv.Left+.menureadrekv.Width+5,5,'возврат','undo.ico','DO writerekv WITH .F.',39,RetTxtWidth('удаление')+44,'возврат'            
     .menureadrekv.Visible=.F.
     .menuexitrekv.Visible=.F.     
     DO addButtonOne WITH 'fBoss','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fBoss','Do readboss WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fBoss','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fBoss','Do readboss WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fBoss','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delboss',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fBoss','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'реквизиты','rekv.ico','DO inputrekv',39,.menucont1.Width,'реквизиты'  
     DO addButtonOne WITH 'fBoss','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'возврат','undo.ico','DO returnboss',39,.menucont1.Width,'возврат'  
     
     DO addmenureadspr WITH 'fboss','DO writeboss WITH .T.','DO writeboss WITH .F.'         

     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Parent.menucont1.Height-110             
          .RecordSourceType=1
          .RecordSource='boss'
           DO addColumnToGrid WITH 'fBoss.fGrid',3 
          .Column1.Header1.Caption='Наименование должности'
          .Column2.Header1.Caption='Фамилия'
          .Column1.ControlSource='" "+boss.doljn'
          .Column2.ControlSource='" "+boss.fam'    
          .Column1.Width=.Width/3 
          .Column2.Width=.Width-.column1.Width-SYSMETRIC(5)-13-3        
          .Column3.Width=0
          .colNesInf=2     
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')        
     ENDWITH 
     DO adtbox WITH 'fboss',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,0
     DO adtbox WITH 'fboss',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')
     DO gridSizeNew WITH 'fboss','fGrid','shapeingrid'    
     DO adLabMy WITH 'fboss',11,'Реквизиты организации',.fGrid.Top+.fGrid.Height+10,0,.Width,2
     WITH .Lab11
          .FontSize=dFontSize+2
          .FontBold=.T.   
          .Visible=.T.     
     ENDWITH 
     DO addShape WITH 'fboss',1,5,.Height-70,60,.Width-10,8    
     DO adtbox WITH 'fboss',11,11,.Shape1.Top+25,(.Shape1.Width-24)/3,dHeight,'newOffice',.F.,0
     DO adtbox WITH 'fboss',12,.txtBox11.Left+.TxtBox11.Width+6,.txtBox11.Top,.TxtBox11.Width,dHeight,'newAdres',.F.,0
     DO adtbox WITH 'fboss',13,.txtBox12.Left+.TxtBox12.Width+6,.txtBox11.Top,.TxtBox11.Width,dHeight,'newBoss',.F.,0
     .SetAll('SpecialEffect',1,'MytxtBox')
     DO adLabMy WITH 'fboss',1,'Наименование организации',.Shape1.Top+3,.TxtBox11.Left,.txtBox11.Width,2
     DO adLabMy WITH 'fboss',2,'Адрес организации',.lab1.Top,.TxtBox12.Left,.txtBox12.Width,2
     DO adLabMy WITH 'fboss',3,'Руководитель организации',.lab1.Top,.TxtBox13.Left,.txtBox13.Width,2
ENDWITH
fboss.Show
*************************************************************************************************************************
PROCEDURE readboss
PARAMETERS parlog
SELECT boss
WITH fBoss
     .menureadrekv.Visible=.F.
     .menuexitrekv.Visible=.F.
     .nrec=RECNO()
     IF parlog        
        SET DELETED OFF
        LOCATE FOR DELETED()
        IF FOUND()
           RECALL
           BLANK
        ELSE
           APPEND BLANK
        ENDIF
        SET DELETED ON 
        .nrec=RECNO()
        .fGrid.Refresh
     ENDIF
     .fGrid.columns(.fGrid.columnCount).SetFocus
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .SetAll('Visible',.T.,'MyTxtBox')
     GO .nrec
     SCATTER TO fboss.dim_ap
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .txtbox1.ControlSource='fboss.dim_ap(1)'
     .txtbox2.ControlSource='fboss.dim_ap(2)'     
     .txtbox1.Top=lineTop
     .txtbox2.Top=linetop
     .txtbox1.Height=.fGrid.RowHeight+1
     .txtbox2.Height=.fGrid.RowHeight+1
     .txtbox1.Enabled=.T.
     .txtbox2.Enabled=.T.
     .txtbox1.BackStyle=1
     .txtbox2.BackStyle=1
     .fGrid.Enabled=.F.
     .txtbox1.SetFocus
ENDWITH    
**************************************************************************************************************************
PROCEDURE writeboss
PARAMETERS par_log
WITH fBoss
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     .menureadrekv.Visible=.F.
     .menuexitrekv.Visible=.F.
     .txtbox1.Visible=.F.
     .txtbox2.Visible=.F.
     .fGrid.Enabled=.T.
     .SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     IF par_log
        SELECT boss
        GATHER FROM fboss.dim_ap 
     ENDIF   
     SELECT boss
     DO WHILE .T.
        LOCATE FOR EMPTY(doljn).AND.EMPTY(fam)
        IF FOUND()                              
           DELETE
        ELSE
           EXIT
        ENDIF
     ENDDO    
     .Refresh
     GO .nrec
     .fGrid.SetFocus  
ENDWITH      
************************************************************************************************************************
PROCEDURE writerekv 
PARAMETERS par_log
WITH fBoss
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     .menureadrekv.Visible=.F.
     .menuexitrekv.Visible=.F.
     .txtbox1.Visible=.F.
     .txtbox2.Visible=.F.
     .fGrid.Enabled=.T.
     .SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .txtbox11.Enabled=.F.
     .txtbox12.Enabled=.F.
     .txtbox13.Enabled=.F.
     .txtbox11.BackStyle=0
     .txtbox12.BackStyle=0
     .txtbox13.BackStyle=0
ENDWITH    
IF par_log
   SELECT datshtat
   REPLACE office WITH newOffice,adres WITH newAdres,boss WITH newBoss  
   SELECT boss   
ENDIF
fboss.refresh
************************************************************************************************************************
PROCEDURE inputrekv
WITH fBoss
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .menureadrekv.Visible=.T.
     .menuexitrekv.Visible=.T.
     .txtbox11.Enabled=.T.
     .txtbox12.Enabled=.T.
     .txtbox13.Enabled=.T.
     .fGrid.Enabled=.F.
     .txtbox11.BackStyle=1
     .txtbox12.BackStyle=1
     .txtbox13.BackStyle=1
     .txtbox11.SetFocus
ENDWITH 
*************************************************************************************************************************
PROCEDURE delboss
fboss.Setall('BorderWidth',0,'Mycontmenu') 
SELECT boss
fboss.fGrid.GridDelRec('fboss.fGrid','boss') 
*************************************************************************************************************************
PROCEDURE returnboss
SELECT boss
DELETE FOR EMPTY(fam).AND.EMPTY(doljn)
fboss.Release
*-------------------------------------------------------------------------------------------------------------------------
*                                                Тарифная сетка
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE proctarnet
SELECT sprkoef
GO TOP
STORE 0 TO newKod,newName
STORE '' TO newForPrn,newPrim
logAp=.F.
fnet=CREATEOBJECT('Formspr')
WITH fnet   
     .Caption='Тарифная сетка'
     .AddProperty('korkfold',0)    
     .procExit='DO returnnet'     
     DO addButtonOne WITH 'fNet','menuCont1',10,5,'новая','pencila.ico','DO readNet WITH .T.',39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fNet','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico','DO readNet WITH .F.',39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fNet','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delNet',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fNet','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO returnNet',39,.menucont1.Width,'возврат'           
     DO addButtonOne WITH 'fNet','butSave',10,5,'записать','pencil.ico','DO writeNet WITH .T.',39,RetTxtWidth('удаление')+44,'записать'      
     .butSave.Visible=.F.

     DO addButtonOne WITH 'fNet','butRet',.butSave.Left+.butSave.Width+3,5,'возврат','undo.ico','DO writeNet WITH .F.',39,RetTxtWidth('удаление')+44,'возврат'  
    .butRet.Visible=.F.     
     WITH .fGrid
          .Top=fnet.menucont1.Top+fnet.menucont1.Height+5    
          .Height=fnet.Height-fnet.menucont1.Height-5                     
          .RecordSourceType=1
          .RecordSource='sprkoef'
          DO addColumnToGrid WITH 'fNet.fGrid',5 
          .Column1.ControlSource='sprkoef.kod'
          .Column2.ControlSource='sprkoef.name' 
          .Column3.ControlSource='sprkoef.forprn'  
          .Column4.ControlSource='sprkoef.prim'   
          .Column1.Width=RettxtWidth(' 123 ')    
          .Column2.Width=RettxtWidth(' 12345.999 ')    
          .Column3.Width=RetTxtWidth(' Печать ')
          .Columns(.ColumnCount).Width=0
          .Column4.Width=.Width-.column1.Width-.Column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount    
          .Column1.Header1.Caption='Кф.'
          .Column2.Header1.Caption='Значение'
          .Column3.Header1.Caption='Печать'
          .Column4.Header1.Caption='Примечание'
          .Column1.Enabled=.F.
          .Column2.Format='Z'
          .Column1.Alignment=1
          .Column2.Alignment=1
          .Column3.Alignment=0
          .Column4.Alignment=0
          .colNesInf=2   
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')        
     ENDWITH    
     DO addtxtboxmy WITH 'fnet',1,1,1,fnet.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fnet',2,1,1,fnet.fGrid.Column2.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fnet',3,1,1,fnet.fGrid.Column3.Width+2,.F.,.F.,0 
     DO addtxtboxmy WITH 'fnet',4,1,1,fnet.fGrid.Column4.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
ENDWITH     
DO gridSizeNew WITH 'fnet','fGrid','shapeingrid'   
fnet.Show
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readnet
PARAMETERS parlog
logAp=parLog
SELECT sprkoef
WITH fNet
     IF parlog
        SELECT sprkoef
        GO BOTTOM
        newKod=kod+1
        APPEND BLANK
        REPLACE kod WITH newKod
     ENDIF
     newKod=IIF(parLog,newKod,kod)
     newname=IIF(parLog,00.00,name)
     newForPrn=IIF(parLog,'',forPrn)
     newPrim=IIF(parLog,'',prim)
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butSave.Visible=.T.
     .butRet.Visible=.T.
     .Refresh
     .SetAll('Visible',.T.,'MyTxtBox')
     fNet.fGrid.columns(fNet.fGrid.columnCount).SetFocus
     lineTop=fnet.fGrid.Top+fnet.fGrid.HeaderHeight+fnet.fGrid.RowHeight*(IIF(fnet.fGrid.RelativeRow<=0,1,fnet.fGrid.RelativeRow)-1)
     .nrec=RECNO()
     .txtBox1.Left=fnet.fGrid.Left+10
     .txtBox2.Left=fnet.txtbox1.Left+fnet.txtbox1.Width-1
     .txtBox3.Left=fnet.txtbox2.Left+fnet.txtbox2.Width-1
     .txtBox4.Left=fnet.txtbox3.Left+fnet.txtbox3.Width-1
     .txtbox1.ControlSource='newKod'
     .txtbox2.ControlSource='newName'
     .txtbox3.ControlSource='newForPrn'
     .txtbox4.ControlSource='newPrim'
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',fnet.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
    * .txtbox1.Enabled=.F.
     .fGrid.Enabled=.F.
     .txtbox2.SetFocus
ENDWITH
*************************************************************************************************************************
PROCEDURE writeNet
PARAMETERS par1
IF par1
   REPLACE name WITH newName,forPrn WITH newForPrn,prim WITH newPrim,kod WITH newKod
ELSE
   IF logAp
      DELETE
   ENDIF
ENDIF
WITH fNet           
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .SetAll('Visible',.F.,'MyTxtBox')
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'columnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.     
     .Refresh
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
ENDWITH     
*************************************************************************************************************************
PROCEDURE delnet
fnet.Setall('BorderWidth',0,'Mycontmenu') 
SELECT datjob
LOCATE FOR kf=sprkoef.kod
IF FOUND()
   SELECT sprkoef
   fnet.fGrid.GridNoDelRec    
ELSE
   fnet.fGrid.GridDelRec('fnet.fGrid','sprkoef')    
ENDIF
*************************************************************************************************************************
PROCEDURE returnnet
DIMENSION dimRet(2)
dimRet(1)=1
dimRet(2)=0
logNet=.F.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl     
     .Caption='Выход'
     DO addOptionButton WITH 'fSupl',1,'выход без замены',20,20,'dimRet(1)',0,'DO procDimRet WITH 1',.T.   
     DO addOptionButton WITH 'fSupl',2,'выход с заменой',.Option1.Top+.Option1.Height+20,.Option1.Left,'dimRet(2)',0,'DO procDimRet WITH 2',.T. 
   
     DO adCheckBox WITH 'fSupl','checkNet','замена тарифных коэффициентов',.Option2.Top+.Option2.Height+10,10,150,dHeight,'logNet',0,.T.,.F.
     .Width=.checkNet.Width+60
     .Option1.Left=(.Width-.Option1.Width)/2
     .Option2.Left=(.Width-.Option2.Width)/2
     .checkNet.Left=(.Width-.checkNet.Width)/2
      *-----------------------------Кнопка выход---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',(.Width-(RetTxtWidth('wотменаw')*2)-15)/2,;
       .checkNet.Top+.checkNet.Height+10,RetTxtWidth('wотменаw'),dHeight+5,'выход','DO procReturnNet'

     *---------------------------------Кнопка отмена-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'отмена','fSupl.Release'        
    .Height=.option1.Height*2+.checkNet.Height+.cont1.Height+70
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show 
************************************************************************************************************************
PROCEDURE procReturnNet
IF dimRet(2)=1
   SELECT datjob
   SET FILTER TO 
   SELECT sprkoef
   DELETE FOR EMPTY(name)
   GO TOP
   DO WHILE !EOF()
      SELECT datjob
      IF logNet
         REPLACE namekf WITH sprkoef.name FOR kf=sprkoef.kod      
      ENDIF 
      SELECT sprkoef
      SKIP
   ENDDO
   SELECT datjob
   SET FILTER TO kodpeop=people.num
ENDIF
fSupl.Release
fNet.Release
*************************************************************************************************************************
PROCEDURE procDimRet
PARAMETERS par1
STORE 0 TO dimRet
dimRet(par1)=1
fSupl.Refresh
**************************************************************************************************************************
PROCEDURE supEr
********************************************************************************************************************************************************
*                                                 Основная для спр-ка норм времени
********************************************************************************************************************************************************
PROCEDURE procnormtime
fNorm=CREATEOBJECT('Formspr')
SELECT sprtime
GO TOP
WITh fNorm
     .Caption='Справочник норм рабочего времени'  
     .procExit='fNorm.fGrid.GridReturn'
     DO addButtonOne WITH 'fNorm','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fNorm','Do readSprTime WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fNorm','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fNorm','Do readSprTime WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fNorm','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delTime',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fNorm','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','fNorm.fGrid.GridReturn',39,.menucont1.Width,'возврат'                 
     DO addmenureadspr WITH 'fNorm',"DO writesprnew WITH 'fNorm','fNorm.fGrid','sprtime'","DO exitwrite WITH 'fNorm','fNorm.fGrid'"
     WITH .fGrid     
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5               
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fNorm.fGrid',16                                                     
          .RecordSource='sprtime'
          .Column1.ControlSource='sprtime->kod'
          .Column2.ControlSource='" "+sprtime.name' 
          .Column3.ControlSource='sprtime.t1'  
          .Column4.ControlSource='sprtime.t2'
          .Column5.ControlSource='sprtime.t3'
          .Column6.ControlSource='sprtime.t4'
          .Column7.ControlSource='sprtime.t5'
          .Column8.ControlSource='sprtime.t6'
          .Column9.ControlSource='sprtime.t7'
          .Column10.ControlSource='sprtime.t8'
          .Column11.ControlSource='sprtime.t9'
          .Column12.ControlSource='sprtime.t10'
          .Column13.ControlSource='sprtime.t11'
          .Column14.ControlSource='sprtime.t12'
          .Column15.ControlSource='sprtime.ntot'
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование' 
          .Column3.Header1.Caption='1'      
          .Column4.Header1.Caption='2'
          .Column5.Header1.Caption='3'
          .Column6.Header1.Caption='4'
          .Column7.Header1.Caption='5'
          .Column8.Header1.Caption='6'
          .Column9.Header1.Caption='7'
          .Column10.Header1.Caption='8'
          .Column11.Header1.Caption='9'
          .Column12.Header1.Caption='10'
          .Column13.Header1.Caption='11'
          .Column14.Header1.Caption='12'
          .Column15.Header1.Caption='всего'
          .Column1.Width=RettxtWidth(' 1234 ')
          .Column3.Width=RetTxtWidth('999999999')
          .Column4.Width=.Column3.Width
          .Column5.Width=.Column3.Width
          .Column6.Width=.Column3.Width
          .Column7.Width=.Column3.Width
          .Column8.Width=.Column3.Width
          .Column9.Width=.Column3.Width
          .Column10.Width=.Column3.Width
          .Column11.Width=.Column3.Width
          .Column12.Width=.Column3.Width
          .Column13.Width=.Column3.Width
          .Column14.Width=.Column3.Width 
          .Column15.Width=.Column3.Width                                       
          .Column2.Width=.Width-.column1.Width-.Column3.Width*13-SYSMETRIC(5)-13-.ColumnCount        
          .Columns(.ColumnCount).Width=0
          .Column3.Format='Z'
          .Column4.Format='Z'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'
          .Column9.Format='Z'
          .Column10.Format='Z'
          .Column11.Format='Z'
          .Column12.Format='Z'
          .Column13.Format='Z'
          .Column14.Format='Z' 
          .Column15.Format='Z' 
          .SetAll('Alignment',1,'ColumnMy')
          .Column2.Alignment=0                            
          .colNesInf=2   
          .SetAll('Movable',.F.,'Columnmy') 
          .SetAll('BOUND',.F.,'Columnmy')
     ENDWITH
     DO gridSizeNew WITH 'fNorm','fGrid','shapeingrid'  
     DO addtxtboxmy WITH 'fNorm',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fNorm',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fNorm',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',5,1,1,.fGrid.Column5.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',6,1,1,.fGrid.Column6.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',7,1,1,.fGrid.Column7.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',8,1,1,.fGrid.Column8.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',9,1,1,.fGrid.Column9.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',10,1,1,.fGrid.Column10.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',11,1,1,.fGrid.Column11.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',12,1,1,.fGrid.Column12.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',13,1,1,.fGrid.Column13.Width+2,.F.,.F.,1,'DO validntot','Z'
     DO addtxtboxmy WITH 'fNorm',14,1,1,.fGrid.Column14.Width+2,.F.,.F.,1,'DO validntot','Z'  
     DO addtxtboxmy WITH 'fNorm',15,1,1,.fGrid.Column15.Width+2,.F.,.F.,1,'DO validntot','Z'  
     .SetAll('Visible',.F.,'MyTxtBox')
ENDWITH 
fNorm.Show
*************************************************************************************************************************
PROCEDURE validntot
WITH fNorm
     .dim_ap(15)=.dim_ap(3)+.dim_ap(4)+.dim_ap(5)+.dim_ap(6)+.dim_ap(7)+.dim_ap(8)+.dim_ap(9)+.dim_ap(10)+.dim_ap(11)+.dim_ap(12)+.dim_ap(13)+.dim_ap(14)
     .txtBox15.Refresh
ENDWITH 
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readSprTime
PARAMETERS parlog
SELECT sprtime
WITH fNorm
     IF parlog
       .fGrid.GridMyAppendBlank(1,'kod','name')   
     ENDIF     
     .SetAll('Visible',.T.,'MyTxtBox')
     .fGrid.columns(.fGrid.columnCount).SetFocus
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .nrec=RECNO()
     SCATTER TO .dim_ap  
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .txtBox3.Left=.txtbox2.Left+.txtbox2.Width-1
     .txtBox4.Left=.txtbox3.Left+.txtbox3.Width-1
     .txtBox5.Left=.txtbox4.Left+.txtbox4.Width-1
     .txtBox6.Left=.txtbox5.Left+.txtbox5.Width-1
     .txtBox7.Left=.txtbox6.Left+.txtbox6.Width-1
     .txtBox8.Left=.txtbox7.Left+.txtbox7.Width-1
     .txtBox9.Left=.txtbox8.Left+.txtbox8.Width-1
     .txtBox10.Left=.txtbox9.Left+.txtbox9.Width-1
     .txtBox11.Left=.txtbox10.Left+.txtbox10.Width-1
     .txtBox12.Left=.txtbox11.Left+.txtbox11.Width-1
     .txtBox13.Left=.txtbox12.Left+.txtbox12.Width-1
     .txtBox14.Left=.txtbox13.Left+.txtbox13.Width-1
     .txtBox15.Left=.txtbox14.Left+.txtbox14.Width-1

     .txtbox1.ControlSource='fNorm.dim_ap(1)'
     .txtbox2.ControlSource='fNorm.dim_ap(2)'
     .txtbox3.ControlSource='fNorm.dim_ap(3)'
     .txtbox4.ControlSource='fNorm.dim_ap(4)'
     .txtbox5.ControlSource='fNorm.dim_ap(5)'
     .txtbox6.ControlSource='fNorm.dim_ap(6)'
     .txtbox7.ControlSource='fNorm.dim_ap(7)'
     .txtbox8.ControlSource='fNorm.dim_ap(8)'
     .txtbox9.ControlSource='fNorm.dim_ap(9)'
     .txtbox10.ControlSource='fNorm.dim_ap(10)'
     .txtbox11.ControlSource='fNorm.dim_ap(11)'
     .txtbox12.ControlSource='fNorm.dim_ap(12)'
     .txtbox13.ControlSource='fNorm.dim_ap(13)'
     .txtbox14.ControlSource='fNorm.dim_ap(14)'
     .txtbox15.ControlSource='fNorm.dim_ap(15)'
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',fNorm.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .txtbox1.Enabled=.F.
     .fGrid.Enabled=.F.
     .txtbox2.SetFocus
ENDWITH 
*************************************************************************************************************************
PROCEDURE delTime
fNorm.Setall('BorderWidth',0,'Mycontmenu')
fNorm.fGrid.GridDelRec('fNorm.fGrid','sprtime')   
*-----------------------------------------Вспомогательные процедуры-------------------------------------------------------
*******************************************************************************************************************
*        Форма для общего расчёта
*******************************************************************************************************************
PROCEDURE formTotalCount
fSupl=CREATEOBJECT('FORMSUPL')
logAvt=.F.
SELECT datShtat
LOCATE FOR ALLTRIM(pathtarif)=pathTarSupl
WITH fSupl   
     .Caption='Общий расчёт'
     .procexit='DO saveTarDate'
     DO addShape WITH 'fSupl',1,10,10,dHeight,300,8   
     DO adTBoxAsCont WITH 'fsupl','txtDate',.Shape1.Left+10,.Shape1.Top+20,RetTxtWidth('WWдата тар.WW'),dHeight,'дата тар.',2,1
     DO adTBoxAsCont WITH 'fsupl','txtStav',.txtDate.Left+.txtDate.Width-1,.txtDate.Top,.txtDate.Width,dHeight,'базовая ст.',2,1
     DO adTBoxAsCont WITH 'fsupl','txtMzp',.txtStav.Left+.txtStav.Width-1,.txtDate.Top,.txtDate.Width,dHeight,'мин.зп.',2,1
         
     DO addtxtboxmy WITH 'fSupl',1,.txtDate.Left,.txtDate.Top+.txtDate.Height-1,.txtDate.Width,.F.,'datshtat.dtarif',2,'DO validvartar'     
     DO addtxtboxmy WITH 'fSupl',2,.txtStav.Left,.txtBox1.Top,.txtStav.Width,.F.,'datshtat.baseSt',2,'DO validvartar','Z'             
     DO addtxtboxmy WITH 'fSupl',3,.txtMzp.Left,.txtBox1.Top,.txtMzp.Width,.F.,'datshtat.nmzp',2,'DO validvartar','Z'   
                      
     .Shape1.Width=.txtDate.Width*3+20
     *IF datShtat.real
     *     DO adCheckBox WITH 'fSupl','checkAvt','автоматически подставлять надбавки и доплаты',.txtBox1.Top+.txtBox1.Height+10,10,150,dHeight,'logAvtNad',0,.T.
     *    .checkAvt.Left=.Shape1.Left+(.Shape1.Width-.checkAvt.Width)/2
     *    .Shape1.Height=.txtDate.Height+.txtBox1.Height+.checkAvt.Height+50
     *ELSE     
       .Shape1.Height=.txtDate.Height+.txtBox1.Height+40
     *ENDIF 
     *----------------Кнопка для выполнения-------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WвыполнениеW')*2-20)/2,.Shape1.Top+.Shape1.Height+25,RetTxtWidth('WвыполнениеW'),dHeight+5,'выполнение','DO totalsum','Выполнить расчет'
     *---------------Кнопка для отказа------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+20,.cont1.Top,.cont1.Width,dHeight+5,'отказ','DO saveTardate','отказ'
     
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
     .Height=.Shape1.Height+.cont1.Height+50
     .Width=.Shape1.Width+20
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
**************************************************************************************************************************
PROCEDURE validvartar
varDtar=dtarif
varBaseSt=baseSt
varNmzp=nmzp
**************************************************************************************************************************
PROCEDURE saveTarDate
SELECT datShtat
LOCATE FOR ALLTRIM(pathtarif)=pathTarSupl
varDtar=dtarif
varBaseSt=baseSt
varNmzp=nmzp
SELECT datjob
tarDateSay='тарификация на '+DTOC(varDTar)+IIF(datshtat.real,' текущая ','')+' (изменить - двойной щелчок мыши)  базовая ставка - '+LTRIM(STR(varBaseSt,8,2))+'  мин.з.п - '+LTRIM(STR(varNmzp,8,2))
frmTop.contDateTarif.ContLabel.Caption=tarDateSay
fSupl.Release
**************************************************************************************************************************
*    Непосредственно процедура общего расчёта
*************************************************************************************************************************
PROCEDURE totalsum
SELECT datShtat
LOCATE FOR ALLTRIM(pathtarif)=pathTarSupl
REPLACE dTarif WITH vardtar,basest WITH varBaseSt
SELECT datjob
tarDateSay='тарификация на '+DTOC(varDTar)+IIF(datshtat.real,' текущая ','')+' (изменить - двойной щелчок мыши)  базовая ставка - '+LTRIM(STR(varBaseSt,8,2))+'  мин.з.п - '+LTRIM(STR(varNmzp,8,2))
frmTop.contDateTarif.ContLabel.Caption=tarDateSay
SELECT people
peopRec=RECNO()
ordPeop=SYS(21)
SET ORDER TO 1
SELECT datJob
*SET FILTER TO 

IF EMPTY(fltJob)
   SET FILTER TO 
ELSE 
    SET FILTER TO &fltJob     
ENDIF 

STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .Shape2.Visible=.T.
     .Shape3.Visible=.T.
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape3.Width=1
ENDWITH
SCAN ALL       
     SELECT people
     SEEK datjob.kodpeop
     DO dopl_sum WITH .T.
     one_pers=one_pers+1
     pers_ch=one_pers/max_rec*100
     fSupl.shape3.Visible=.T.
     fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
     fSupl.Shape3.Width=fSupl.shape2.Width/100*pers_ch  
ENDSCAN
SELECT people
SET ORDER TO &ordPeop
GO peopRec
=SYS(2002)
=INKEY(2)
WITH fSupl          
     .shape2.Visible=.F.
     .shape3.Visible=.F.
     .lab25.Caption='Расчёт выполнен' 
     .lab25.Top=.Shape1.Top+.Shape1.Height+10
     .cont3.Top=.lab25.Top+.lab25.Height    
     .cont3.Visible=.T.
ENDWITH  

*******************************************************************************************************************
*                                   Форма для итогов
*******************************************************************************************************************
PROCEDURE procItog
IF USED('datJobItog')
   SELECT datJobItog
   USE
ENDIF
IF USED('curItog')
   SELECT curItog
   USE
ENDIF
IF USED('curPrnTarFond')
   SELECT curPrnTarFond
   USE
ENDIF
IF USED('curKatItog')
   SELECT curKatItog
   USE
ENDIF
SELECT * FROM sprkat INTO CURSOR curKatItog READWRITE
ALTER TABLE curKatItog ADD COLUMN kse N(7,2)
ALTER TABLE curKatItog ADD COLUMN sumOkl N(10,2)
ALTER TABLE curKatItog ADD COLUMN avOkl N(7,2)
*ALTER TABLE curKatItog ADD COLUMN kse N(7,2)
SELECT curKatItog
APPEND BLANK
REPLACE kod WITH 99, name WITH 'всего'

SELECT * FROM tarfond WHERE tarfond.vac.AND.!EMPTY(tarfond.persved) INTO CURSOR curPrnTarFond READWRITE 
SELECT curPrnTarFond
INDEX ON num TAG T1
GO TOP
num_cx=0
DO WHILE !EOF()
   num_cx=num_cx+1
   REPLACE num WITH num_cx
   SKIP   
ENDDO
log_vac=.T.

SELECT * FROM datjob WHERE SEEK(datjob.kodpeop,'people',1).AND.SEEK(STR(kp,3)+STR(kd,3),'rasp',2) INTO CURSOR datJobItog  READWRITE
IF !EMPTY(fltJob)
   SELECT datJobItog
   SET FILTER TO &fltJob
ENDIF
SELECT num,rec,sum_f,sum_fm FROM tarfond WHERE !EMPTY(sum_fm) INTO CURSOR curItog READWRITE
ALTER TABLE curItog ADD COLUMN sumStav N(12,2)
ALTER TABLE curItog ADD COLUMN sumStavKse N(12,2)
ALTER TABLE curItog ADD COLUMN sumVac N(12,2)
ALTER TABLE curItog ADD COLUMN sumVacKse N(12,2)
INDEX ON num TAG T1
SELECT curItog
APPEND BLANK
REPLACE rec WITH 'оъём',sum_fm WITH 'kse',num WITH 98
APPEND BLANK
REPLACE rec WITH 'всего',sum_fm WITH 'msf',num WITH 99

SELECT datJobItog
DELETE FOR date_in>varDTar
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SET ORDER TO 1
SELECT rasp
SCAN ALL
     rKse=rasp.kse
     SELECT datJobItog
     SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
     DO  WHILE rasp.kp=datjobItog.kp.AND.rasp.kd=datJobItog.kd.AND.!EOF()         
        * IF !vac
            rKse=rKse-datJobItog.kse
        * ENDIF    
         SKIP
     ENDDO
     IF rKse#0        
        DO CASE
           CASE rKse<=1
                kvovac=1
           CASE MOD(rKse,1)=0     
                kvovac=INT(rKse)
           CASE MOD(rKse,1)>0     
                kvovac=INT(rKse)+1    
        ENDCASE               
        kvokse=rKse  
        ksevac=0
        FOR i=1 TO kvovac
            ksevac=IIF(kvokse<=1,kvokse,1)
            SELECT datJobItog
            APPEND BLANK       
            REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kse WITH kseVac,vac WITH .T.,kat WITH rasp.kat,;
                    namekf WITH rasp.nkfvac,pkf WITH rasp.pkf,kf WITH rasp.kfvac     
             SELECT curPrnTarFond
             GO TOP
             DO WHILE !EOF()
                rep_r=ALLTRIM(persved)
                rep_r1='rasp.'+ALLTRIM(persved)
                SELECT datJobItog 
                REPLACE &rep_r WITH &rep_r1         
                SELECT curPrnTarFond
                SKIP
             ENDDO                                    
             SELECT datJobItog                                
             tar_ok=0
             tar_ok=varBaseSt*datJobItog.namekf*IIF(pkf#0,datJobItog.pkf,1)      
             REPLACE tokl WITH tar_ok,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2)
             totsumf=tokl
             totsumfm=mtokl
             SELECT tarfond
             SET FILTER TO !EMPTY(countvac)
             GO TOP
             DO WHILE !EOF()
                new_sum=sum_f
                new_msum=sum_fm
                SELECT datJobItog
                r_sum=EVAL(tarfond.countvac)     
                IF !EMPTY(tarfond.sum_f) 
                   REPLACE &new_sum WITH r_sum  
                   *REPLACE &new_msum WITH IIF(tarfond.logkse,&new_sum*kse,&new_sum)     
                   REPLACE &new_msum WITH IIF(tarfond.logkse,r_sum*kse,&new_sum)     
                   totsumf=IIF(!EMPTY(tarfond.sum_f),totsumf+EVALUATE(ALLTRIM(tarfond.sum_f)),totsumf)
                   totsumfm=IIF(!EMPTY(tarfond.sum_fm),totsumfm+EVALUATE(ALLTRIM(tarfond.sum_fm)),totsumfm)
                ENDIF     
                SELECT tarfond
                SKIP
             ENDDO
             SET FILTER TO
             SELECT datJobItog
             REPLACE total WITH totsumf,msf WITH totsumfm        
             kvokse=kvokse-1
        ENDFOR        
     ENDIF
     SELECT rasp
ENDSCAN
SELECT datJobItog
DELETE FOR tokl=0
DELETE FOR date_in>varDTar
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DO countItog
SELECT curItog
GO TOP
fItog=CREATEOBJECT('FORMSUPL')
WITH fItog
     .Caption='Итоги'    
     .Width=800
     .Height=600
     .AddObject('grdItog','GridMyNew')
     WITH .grdItog
          .Top=0
          .Left=0
          .Width=.Parent.Width        
          .Height=.rowHeight*15
          .scrollBars=2          
          .RecordSourceType=1
          .RecordSource='curItog'
          DO addColumnToGrid WITH 'fItog.grdItog',6
          .Column1.ControlSource='curItog.rec'
          .Column2.ControlSource='curItog.sumStav'
          .Column3.ControlSource='curItog.sumStavKse'
          .Column4.ControlSource='curItog.sumVac'
          .Column5.ControlSource='curItog.sumVacKse'
          .Column1.Header1.Caption='доплата/надбавка'                    
          .Column2.Header1.Caption='на ставку'                    
          .Column3.Header1.Caption='на месяц'                    
          .Column4.Header1.Caption='на ставку'                    
          .Column5.Header1.Caption='на месяц'                    
          .Column2.Width=RetTxtWidth('999999999999')                    
          .Column3.Width=.Column2.Width   
          .Column4.Width=.Column2.Width   
          .Column5.Width=.Column2.Width   
          .Column2.Format='Z'
          .Column3.Format='Z'                 
          .Column4.Format='Z'
          .Column5.Format='Z'
          .SetAll('Alignment',1,'ColumnMy')
          .Column1.Alignment=0
          .Column6.Width=0
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount
          DO gridSizeNew WITH 'fItog','grdItog','shapeingrid'         
     ENDWITH       
   
     .AddObject('grdOklad','GridMyNew')
     DO adTBoxAsCont WITH 'fItog','txtOklad',0,.grdItog.Top+.grdItog.Height-1,.grdItog.Width,dHeight,'суммы окладов и средние оклады',2,1
     WITH .grdOklad
          .Top=.Parent.txtOklad.Top+.Parent.txtOklad.Height-1
          .Left=0
          .Width=.Parent.Width        
          .Height=.rowHeight*(RECCOUNT('curkatitog')+1)
          .scrollBars=2          
          .RecordSourceType=1
          .RecordSource='curKatItog'
          DO addColumnToGrid WITH 'fItog.grdOklad',5
          .Column1.ControlSource='curKatItog.name'
          .Column2.ControlSource='curKatItog.kse'
          .Column3.ControlSource='curKatItog.sumOkl'
          .Column4.ControlSource='curKatItog.avOkl'
          .Column1.Header1.Caption='персонал'                    
          .Column2.Header1.Caption='к-во'                    
          .Column3.Header1.Caption='сумма'                    
          .Column4.Header1.Caption='средн.'                    
          .Column2.Width=RetTxtWidth('99999/99')                    
          .Column3.Width=RetTxtWidth('9999999999')   
          .Column4.Width=.Column3.Width   
          .Column5.Width=0
          .Column2.Format='Z'
          .Column3.Format='Z'                 
          .Column4.Format='Z'
          .SetAll('Alignment',1,'ColumnMy')
          .Column1.Alignment=0 
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount       
          
          DO gridSizeNew WITH 'fItog','grdOklad','shapeingrid1'         
     ENDWITH     
             
     DO adCheckBox WITH 'fItog','checkVac','учитывать вакантные',.grdOklad.Top+.grdOklad.Height+10,10,150,dHeight,'log_Vac',0,.T.,'DO countItog'
     .checkVac.Left=(.Width-.checkVac.Width)/2
     
     *----------------Кнопка для выполнения-------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fItog','cont1',(.Width-RetTxtWidth('Wпечать суммыW')*3-20)/2,.checkVac.Top+.checkVac.Height+10,RetTxtWidth('Wпечать суммыW'),dHeight+5,'печать суммы',"DO printreport WITH 'repitog','итоги','curItog'",'печать'
     *----------------Кнопка для выполнения-------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fItog','cont2',.cont1.Left+.cont1.Width+10,.cont1.Top,.cont1.Width,.cont1.Height,'печать оклады',"DO printreport WITH 'repitogkat','итоги','curKatItog'",'печать'
     *---------------Кнопка для отказа------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fItog','cont3',.cont2.Left+.cont2.Width+10,.cont1.Top,.cont1.Width,.cont1.Height,'возврат','fItog.Release','возврат'
     
     
     .Height=.grdItog.height+.grdOklad.Height+.txtOklad.Height+.checkvac.Height+.cont1.Height+40
     .Autocenter=.T.
ENDWITH
fItog.Show
**************************************************************************************************************************
PROCEDURE countItog
SELECT curItog
REPLACE sumStav WITH 0,sumStavKse WITH 0,sumVac WITH 0,sumVacKse WITH 0 ALL
SCAN ALL
     sumForStav=sum_f
     sumForStavKse=sum_fm
     IF !EMPTY(sumForStav)
        SELECT datJobItog
        IF log_vac
           SUM &sumForStav TO sum1  
        ELSE 
           SUM &sumForStav TO sum1 FOR !vac
        ENDIF 
        SUM &sumForStav TO sum3 FOR vac
        SELECT curItog
        REPLACE sumStav WITH sum1,sumVac WITH IIF(log_vac,sum3,0)
     ENDIF 
     IF !EMPTY(sumForStavKse)
        SELECT datJobItog
        IF log_vac
           SUM &sumForStavKse TO sum2
        ELSE
           SUM &sumForStavKse TO sum2 FOR !vac
        ENDIF        
        SUM &sumForStavKse TO sum4 FOR vac
        SELECT curItog
        REPLACE sumStavKse WITH sum2,sumVacKse WITH IIF(log_vac,sum4,0)
     ENDIF      
     SELECT curItog
ENDSCAN
SELECT datJobItog
IF log_vac
   SUM total,msf TO sum1,sum11 
   SUM total,msf TO sum2,sum22 FOR vac
   
ELSE
   SUM total,msf TO sum1,sum11 FOR !vac
ENDIF
SELECT curItog
GO BOTTOM
REPLACE sumStav WITH sum1,sumVac WITH IIF(log_vac,sum2,0)
REPLACE sumStavKse WITH sum11,sumVacKse WITH IIF(log_vac,sum22,0)
GO TOP 
SELECT curkatItog
SCAN ALL
     SELECT datJobItog
     SUM mtOkl,kse TO mToklch,ksekat FOR IIF(log_vac,kat=curKatItog.kod,kat=curKatItog.kod.AND.!vac)
         
     SELECT curKatItog
     REPLACE  sumOkl WITH mToklch,kse WITH kseKat,avOkl WITH IIF(ksekat#0,sumOkl/ksekat,0)
ENDSCAN
SUM kse, sumOkl TO ksech,oklCh
GO BOTTOM
REPLACE kod WITH 99, name WITH 'всего',kse WITH ksech,sumOkl WITH oklCh,avOkl WITH IIF(kse#0,sumOkl/kse,0)
GO TOP 
*-------------------------------------------------------------------------------------------------------------------------
*                                              Процедура для фильтра - новый вариант
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE procFilterNew
PARAMETERS parFrm,parNum
* parFrm - родительская форма
* parNum - 1-персонал, 2-ускоренная замена
numProc=parNum
parentform=parfrm
filter_ch=''
filter_peop=''
sostavflt=''
IF USED('curFltPodr')
   SELECT curFltPodr
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprpodr INTO CURSOR curFltPodr READWRITE
SELECT curFltPodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1
ord_podr=SYS(21)

IF USED('curFltDolj')
   SELECT curFltDolj
   USE
ENDIF 
SELECT kod,namework,fl,otm FROM sprdolj INTO CURSOR curFltDolj ORDER BY namework READWRITE
SELECT curFltDolj
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltKat')
   SELECT curFltKat
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprkat INTO CURSOR curFltkat ORDER BY name READWRITE
SELECT curFltKat
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltKval')
   SELECT curFltKval
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprkval INTO CURSOR curFltkval ORDER BY name READWRITE
SELECT curFltKval
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltType')
   SELECT curFltType
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprType INTO CURSOR curFltType ORDER BY name READWRITE
SELECT curFltType
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltKoef')
   SELECT curFltKoef
   USE
ENDIF 
SELECT kod,name,fl,otm,nameSup FROM sprKoef INTO CURSOR curFltKoef READWRITE
SELECT curFltKoef
REPLACE fl WITH .F.,otm WITH '' ALL

REPLACE nameSup WITH STR(kod,2)+' - '+STR(name,5,3) ALL 
INDEX ON kod TAG T1

IF USED('curFltGrup')
   SELECT curFltGrup
   USE
ENDIF 
SELECT name,fl,sostav1 FROM datagrup INTO CURSOR curFltGrup READWRITE
SELECT curFltGrup
REPLACE fl WITH .F. ALL
ALTER TABLE curFltGrup ADD COLUMN otm C(1)

CREATE CURSOR curFltVac (name C(20),forflt C(20), fl L,otm C(1)) 
APPEND BLANK 
REPLACE name WITH 'Вакантная',forflt WITH 'vac=.T.'
APPEND BLANK 
REPLACE name WITH 'Занятая',forflt WITH '!vac'


CREATE CURSOR curFltMol (name C(20),forflt C(20), fl L,otm C(1)) 
APPEND BLANK 
REPLACE name WITH 'да',forflt WITH '!EMPTY(dmol)'
APPEND BLANK 
REPLACE name WITH 'нет',forflt WITH 'EMPTY(dmol)'

CREATE CURSOR curFltPrim (name C(20),forflt C(20), fl L,otm C(1)) 
APPEND BLANK 
REPLACE name WITH 'Только с примечаниями',forflt WITH '!EMPTY(dopsvstr)'
APPEND BLANK 
REPLACE name WITH 'Только без примечаний',forflt WITH 'EMPTY(dopsvstr)'

IF USED('curFltBase')
   SELECT curFltBase
   USE
ENDIF 
SELECT * FROM fltBase INTO CURSOR curFltBase READWRITE
SELECT fltBase
ALTER TABLE curFltBase ADD COLUMN sostSup M
REPLACE strFlt WITH '',sayfl WITH '' ALL
rrec=0
frmFlt=CREATEOBJECT('FORMSUPL')
WITH frmFlt   
     .Caption='Выбор сведений'
     .procExit='DO exitFromFrmFlt'
     DO addListBoxMy WITH 'frmFlt',1,20,20,400,600  
     .AddObject('lstLine','LINE')
     WITH .listBox1
          .RowSource='curFltBase.namefL,sayFl'           
          .RowSourceType=6
          .ColumnCount=2
          
          .ColumnWidths='170,450' 
          .ColumnLines=.F.
          .procForValid='DO validListFlt'
          .procForKeyPress='DO KeyPressListFlt'       
     ENDWITH   
     .lstLine.Top=.listBox1.Top
     .lstLine.Height=.listBox1.Height
     .lstLine.Left=.ListBox1.Left+170+3
     .lstLine.Width=0    
     .lstLine.Visible=.T.      
    
     DO addcontlabel WITH 'frmFlt','cont1',.listBox1.Left+(.listBox1.Width-RetTxtWidth('WВыполнитьW')*3-20)/2,.listBox1.Top+.listBox1.Height+10,;
     RetTxtWidth('WВыполнитьW'),dHeight+3,'Выполнить','DO procFilterPeop'
     DO addcontlabel WITH 'frmFlt','cont2',.Cont1.Left+.Cont1.Width+10,.Cont1.Top,;
     .Cont1.Width,dHeight+3,'Сброс','DO filterRelease'    
     DO addcontlabel WITH 'frmFlt','cont3',.Cont2.Left+.Cont2.Width+10,.Cont1.Top,;
     .Cont1.Width,dHeight+3,'Возврат','DO exitFromFrmFlt'  
     
     DO addcontlabel WITH 'frmFlt','contReturn',.listBox1.Left+(.listBox1.Width-RetTxtWidth('WВозвратW'))/2,.listBox1.Top+.listBox1.Height+10,;
     RetTxtWidth('WВозвратW'),dHeight+3,'Возврат','DO objFrmFltVisible WITH .F.'
     .contReturn.Visible=.F.       
          
     DO addListBoxMy WITH 'frmFlt',2,20,20,.listBox1.Height,.listBox1.Width    
     WITH .listbox2
          .Visible=.F.
          .ColumnCount=2
          .ColumnWidths='20,600'
          .RowSourceType=6
          .ControlSource=''
          .procForKeyPress='DO KeyPressListFlt2'
     ENDWITH   
     .AddObject('lstLine1','LINE')
     .lstLine1.Top=.listBox2.Top
     .lstLine1.Height=.listBox2.Height
     .lstLine1.Left=.ListBox2.Left+20+3
     .lstLine1.Width=0    
     .lstLine1.Visible=.F.
     .Width=.listBox1.Width+40
     .Height=.ListBox1.Height+.cont1.Height+50
     DO pasteImage WITH 'frmFlt'
     .Show    
ENDWITH 
**************************************************************************************************************************
PROCEDURE exitFromFrmFlt
frmFlt.Visible=.F.
frmFlt.Release
**************************************************************************************************************************
PROCEDURE filterRelease
SELECT curFltPodr
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltDolj
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltKat
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltKval
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltType
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltKoef
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltGrup
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltVac
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltMol
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltBase
REPLACE strFlt WITH '',sayFl WITH '' ALL
GO TOP
SELECT people
SET FILTER TO 
frmFlt.Refresh
*******************************************************************************************************************************
PROCEDURE objFrmFltVisible
PARAMETERS par1
WITH frmFlt
     IF par1 
        .listBox1.Visible=.F.
        .listBox2.Visible=.T.  
        .SetAll('Visible',.F.,'mycontlabel')
        .SetAll('Visible',.F.,'mycommandbutton')
        .lstLine1.Visible=.T.
        .contReturn.Visible=.T. 
     ELSE
        .listBox2.Visible=.F.
        .listBox1.Visible=.T.
        .SetAll('Visible',.T.,'mycontlabel')
        .SetAll('Visible',.T.,'mycommandbutton')
        .lstLine1.Visible=.F.
        .contReturn.Visible=.F. 
        .Refresh
     ENDIF   
ENDWITH  
*************************************************************************************************************************
PROCEDURE validListFlt
IF EMPTY(curFltBase.procflt)
   RETURN 
ENDIF
procSelect=curFltBase.procflt
&procSelect
********************************************************************************************************************************
PROCEDURE fltPodr
DO objFrmFltVisible WITH .T.
WITH frmFlt      
     .listBox2.rowSource='curFltPodr.otm,name'    
     .listBox2.ControlSource=''   
     .listBox2.procForClick="Do procSelPodr WITH 'curFltPodr'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltPodr','name'"
     .listBox2.procForKeyPress='Do keyPressListPodr'
ENDWITH
********************************************************************************************************************************
PROCEDURE returnFromFltPodr
PARAMETERS parBase,parName
DO objFrmFltVisible WITH .F.
SELECT &parBase
repFlt=''
repSayFlt=''
sostavFlt=''
GO TOP
logRep=.F.
DO CASE
   CASE curFltBase.priznak=1 
        SCAN ALL
             IF fl
                repflt=repflt+','+LTRIM(STR(kod))+','
                repSayFlt=repSayFlt+','+ALLTRIM(&parName)
                logRep=.T.        
             ENDIF
        ENDSCAN
        SELECT curFltBase
        repFlt=IIF(logRep,'"'+repFlt+'"','')
        REPLACE sayFl WITH repSayFlt,strFlt WITH repFlt
   CASE curFltBase.priznak=2     
        SCAN ALL
             IF fl
                repflt=ALLTRIM(forFlt)
                repSayFlt=ALLTRIM(&parName)
                logRep=.T.    
                EXIT    
             ENDIF
        ENDSCAN
        SELECT curFltBase
        REPLACE sayFl WITH repSayFlt,strFlt WITH repFlt
   CASE curFltBase.priznak=3     
        SCAN ALL
             IF fl
                repflt='sostavflt'
                sostavFlt=sostav1
                repSayFlt=ALLTRIM(&parName)
                logRep=.T.    
                EXIT    
             ENDIF
        ENDSCAN
        SELECT curFltBase
        REPLACE sayFl WITH repSayFlt,strFlt WITH repFlt    
ENDCASE        
********************************************************************************************************************************
PROCEDURE procSelPodr
PARAMETERS parBase
SELECT &parBase
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,'•','')
GO rrec
frmFlt.listBox2.SetFocus
frmFlt.lstLine1.Visible=.T.
*************************************************************************************************************************
PROCEDURE keyPressListPodr
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltPodr' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltDolj
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltDolj.otm,namework'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltDolj'"  
     .listBox2.procForKeyPress='Do keyPressListDolj' 
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltDolj','namework'"
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListDolj
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltDolj','namework'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltDolj' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltPers
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltKat.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltKat'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltKat','name'" 
     .listBox2.procForKeyPress='Do keyPressListPers' 
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListPers
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltKat','name'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltKat' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltKval
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltKval.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltKval'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltKval','name'" 
     .listBox2.procForKeyPress='Do keyPressListKval' 
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListKval
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltKval','name'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltKval' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltTr
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltType.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltType'"
     .listBox2.procForKeyPress='Do keyPressListType'
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltType','name'"
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListType
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltType','name'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltType' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltVac
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltVac.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltVac'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltVac','name'"
     .listBox2.procForKeyPress='Do keyPressListVac'
ENDWITH
**************************************************************************************************************************
PROCEDURE fltMol
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltMol.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltMol'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltMol','name'"
     .listBox2.procForKeyPress='Do keyPressListVac'
ENDWITH
**************************************************************************************************************************
PROCEDURE fltPrim
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltPrim.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltPrim'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltPrim','name'"
     .listBox2.procForKeyPress='Do keyPressListVac'
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListType
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltVac','name'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltVac' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltKoef
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltKoef.otm,nameSup'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltKoef'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltKoef','nameSup'"
     .listBox2.procForKeyPress='Do keyPressListKoef'
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListKoef
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltKoef','nameSup'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltKoef' 
ENDCASE   
**************************************************************************************************************************
PROCEDURE fltGrup
DO objFrmFltVisible WITH .T.
WITH frmFlt        
     .listBox2.rowSource='curFltGrup.otm,name'  
     .listBox2.procForClick="Do procSelPodr WITH 'curFltGrup'"
     .contReturn.procForClick="DO returnFromFltPodr WITH 'curFltGrup','name'"
     .listBox2.procForKeyPress='Do keyPressListGrup'
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressListGrup
DO CASE
   CASE LASTKEY()=27
        DO returnFromFltPodr WITH 'curFltGrup','name'     
   CASE LASTKEY()=13
        Do procSelPodr WITH 'curFltGrup' 
ENDCASE 
*************************************************************************************************************************
PROCEDURE keyPressListFlt
DO CASE
   CASE LASTKEY()=13
        IF EMPTY(curFltBase.procflt)
           RETURN 
        ENDIF
        procSelect=curFltBase.procflt
        &procSelect
ENDCASE
**************************************************************************************************************************
PROCEDURE keyPressListFlt2
**************************************************************************************************************************
PROCEDURE procFilterPeop
frmFlt.Visible=.F.
SELECT curFltBase
filter_ch=''
filter_peop=''
GO TOP 
DO WHILE !EOF()
   IF !EMPTY(strFlt).AND.npeop=1
      filter_ch=filter_ch+ALLTRIM(namepl)+ALLTRIM(strFlt)+'.AND.'
   ENDIF  
   IF !EMPTY(strFlt).AND.npeop=2
      filter_peop=filter_peop+ALLTRIM(namepl)+ALLTRIM(strFlt)+'.AND.'  
   ENDIF    
   SKIP
   
ENDDO
frmFlt.Release
DO CASE 
   CASE !EMPTY(filter_ch).AND.!EMPTY(filter_peop)  &&по datjob и people
         DO CASE
            CASE numProc=1     
                 frmToplog_fl=.T. 
                 frmTop.filter_ch=filter_ch
                 filter_ch=SUBSTR(filter_ch,1,LEN(filter_ch)-5)+'.AND.EMPTY(dateout)'  
                 frmTop.filter_ch=filter_ch 
                 fltJob=filter_ch                     
                 SELECT curFltDatJob
                 DELETE ALL
                 APPEND FROM datJob FOR EVALUATE(frmTop.filter_ch) 
                 filter_peop=SUBSTR(filter_peop,1,LEN(filter_peop)-5)  
                 frmTop.filter_peop=filter_peop    
                 SELECT people
                 SET FILTER TO SEEK(num,'curFltDatjob',1).AND.EVALUATE(&parentForm..filter_peop)
                 GO TOP       
                 frmTop.Refresh   
                 frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus                  
            CASE numProc=2 
                 &parentform..log_fl=.T. 
                 &parentform..filter_ch=filter_ch   
                 SELECT datJob
                 filter_ch=SUBSTR(filter_ch,1,LEN(filter_ch)-5)+'.AND.EMPTY(dateout)'  
                 &parentForm..filter_ch=filter_ch     
                 SET FILTER TO EVALUATE(&parentForm..filter_ch)
                 GO TOP   
                 &parentForm..Refresh          
        ENDCASE 
   CASE !EMPTY(filter_ch).AND.EMPTY(filter_peop)   &&только по datjob
        DO CASE
           CASE numProc=1     
                frmToplog_fl=.T. 
                frmTop.filter_ch=filter_ch
                filter_ch=SUBSTR(filter_ch,1,LEN(filter_ch)-5)+'.AND.EMPTY(dateout)'   
                frmTop.filter_ch=filter_ch 
                fltJob=filter_ch                     
                SELECT curFltDatJob
                DELETE ALL
                APPEND FROM datJob FOR EVALUATE(frmTop.filter_ch)            
                SELECT people
                SET FILTER TO SEEK(num,'curFltDatjob',1)
                GO TOP       
                frmTop.Refresh   
                frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus       
           CASE numProc=2 
                &parentform..log_fl=.T. 
                &parentform..filter_ch=filter_ch   
                SELECT datJob
                filter_ch=SUBSTR(filter_ch,1,LEN(filter_ch)-5)+'.AND.EMPTY(dateout)'  
                &parentForm..filter_ch=filter_ch     
                SET FILTER TO EVALUATE(&parentForm..filter_ch)
                GO TOP   
                &parentForm..Refresh          
        ENDCASE 
   CASE EMPTY(filter_ch).AND.!EMPTY(filter_peop)   &&только по people
        DO CASE
           CASE numProc=1     
                frmToplog_fl=.T. 
                filter_peop=SUBSTR(filter_peop,1,LEN(filter_peop)-5)  
                frmTop.filter_peop=filter_peop 
                SELECT people
                SET FILTER TO EVALUATE(&parentForm..filter_peop)
                GO TOP       
                frmTop.Refresh   
                frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus     
           CASE numProc=2 
                &parentform..log_fl=.T. 
                &parentform..filter_peop=filter_peop   
                SELECT people
                filter_peop=SUBSTR(filter_peop,1,LEN(filter_peop)-5)  
                &parentForm..filter_peop=filter_peop     
                SET FILTER TO EVALUATE(&parentForm..filter_peop)
                GO TOP  
                SELECT datjob 
                SET FILTER TO SEEK(kodpeop,'people',1)
                GO TOP
                &parentForm..Refresh          
        ENDCASE 
   CASE EMPTY(filter_ch).AND.EMPTY(filter_peop)      &&нет фильтра
        DO CASE 
           CASE numProc=1
                fltJob=''  
                SELECT datjob
                SET FILTER TO 
                SELECT people
                fltRec=RECNO()
                SET FILTER TO 
                ON ERROR DO erSup
                IF fltRec#0
                   GO fltRec
                ENDIF
                ON ERROR 
                frmTop.Refresh  
                frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
           CASE numProc=2
                SELECT datJob
                SET FILTER TO 
                GO TOP   
                &parentForm..Refresh             
        ENDCASE  
ENDCASE 








*IF !EMPTY(filter_ch)   
*   DO CASE
*      CASE numProc=1     
*           frmToplog_fl=.T. 
*           frmTop.filter_ch=filter_ch
*           filter_ch=SUBSTR(filter_ch,1,LEN(filter_ch)-5)  
*           frmTop.filter_ch=filter_ch 
*           fltJob=filter_ch                     
*           SELECT curFltDatJob
*           DELETE ALL
*           APPEND FROM datJob FOR EVALUATE(frmTop.filter_ch)            
*           SELECT people
*           SET FILTER TO SEEK(num,'curFltDatjob',1)
*           GO TOP       
*           frmTop.Refresh   
*           frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus       
*      CASE numProc=2 
*           &parentform..log_fl=.T. 
*           &parentform..filter_ch=filter_ch   
*           SELECT datJob
*           filter_ch=SUBSTR(filter_ch,1,LEN(filter_ch)-5)  
*           &parentForm..filter_ch=filter_ch     
*           SET FILTER TO EVALUATE(&parentForm..filter_ch)
*           GO TOP   
*           &parentForm..Refresh          
*   ENDCASE 
*ELSE
   *DO CASE 
   *   CASE numProc=1
   *        fltJob=''  
   *        SELECT datjob
   *        SET FILTER TO 
   *        SELECT people
   *        fltRec=RECNO()
   *        SET FILTER TO 
   *        ON ERROR DO erSup
   *        IF fltRec#0
   *           GO fltRec
   *        ENDIF
   *        ON ERROR 
   *        frmTop.Refresh  
   *        frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
   *   CASE numProc=2
   *        SELECT datJob
   *        SET FILTER TO 
   *        GO TOP   
   *        &parentForm..Refresh             
  * ENDCASE     
*ENDIF 
**************************************************************************************************************************
PROCEDURE otmforfilter
REPLACE fl WITH IIF(fl,.F.,.T.)
**************************************************************************************************************************
PROCEDURE gotfocusfilter
PARAMETERS parbase,parobj,parstr
SELECT &parbase
IF parstr
   &parobj..Rowsource="IIF(&parbase->fl,' • ','   ')+' '+STR(&parbase->name,7,2)"
ELSE 
   &parobj..Rowsource="IIF(&parbase->fl,' • ','   ')+' '+&parbase->name"
ENDIF 
&parobj..DisplayCount=(SYSMETRIC(2)-formflt.Top-formflt.Height)/FONTMETRIC(1,dFontname,dFontSize)
*********************************************************************************************************************
*                                             Выход из фильтра
*********************************************************************************************************************
PROCEDURE exitfilter
SELECT sprpodr
SET ORDER TO &ord_podr
SELECT sprdolj
SET ORDER TO &ord_dolj
&parentform..log_fl=.F.
&parentform..filter_ch=''
formflt.Release
*********************************************************************************************************************
*                                             Тарификация - штатное (новый вариант)
*********************************************************************************************************************
PROCEDURE shtatPeopleOld
IF USED('curDatGrup')
   SELECT curDatGrup
   USE
ENDIF
IF USED('curShtatJob')
   SELECT curShtatJob
   USE
ENDIF
SELECT * FROM datagrup INTO CURSOR curDatGrup READWRITE
SELECT curDatgrup
APPEND BLANK
REPLACE name WITH 'Организация'

SELECT * FROM datjob INTO CURSOR curShtatJob READWRITE
SELECT curShtatJob
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
IF USED('peopsh')
   SELECT peopsh
   USE
ENDIF
DIMENSION dimOptVac(2)
dimOptVac(1)=1
dimOptVac(2)=0 
SELECT * FROM sprpodr INTO CURSOR peopsh READWRITE
ALTER TABLE peopsh ADD COLUMN r N(8,2)
ALTER TABLE peopsh ADD COLUMN p N(8,2)
ALTER TABLE peopsh ADD COLUMN rp N(8,2)
ALTER TABLE peopsh ADD COLUMN r1 N(8,2)
ALTER TABLE peopsh ADD COLUMN p1 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp1 N(8,2)
ALTER TABLE peopsh ADD COLUMN r2 N(8,2)
ALTER TABLE peopsh ADD COLUMN p2 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp2 N(8,2)
ALTER TABLE peopsh ADD COLUMN r3 N(8,2)
ALTER TABLE peopsh ADD COLUMN p3 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp3 N(8,2)
ALTER TABLE peopsh ADD COLUMN r4 N(8,2)
ALTER TABLE peopsh ADD COLUMN p4 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp4 N(8,2)
ALTER TABLE peopsh ADD COLUMN r5 N(8,2)
ALTER TABLE peopsh ADD COLUMN p5 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp5 N(8,2)
ALTER TABLE peopsh ADD COLUMN r6 N(8,2)
ALTER TABLE peopsh ADD COLUMN p6 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp6 N(8,2)
ALTER TABLE peopsh ADD COLUMN r7 N(8,2)
ALTER TABLE peopsh ADD COLUMN p7 N(8,2)
ALTER TABLE peopsh ADD COLUMN rp7 N(8,2)

SELECT peopsh
INDEX ON np TAG T1

SELECT sprkat
COUNT TO max_kat
DIMENSION name_kat(max_kat),kod_kat(max_kat)
GO TOP
FOR i=1 TO max_kat
    name_kat(i)=name
    kod_kat(i)=kod
    SKIP  
ENDFOR
DIMENSION dim_kat(max_kat),dim_tot(max_kat+1,2)
DO inCurPeopSh WITH .F.
SELECT peopsh
GO TOP
fShtat=CREATEOBJECT('FORMMY')
=SYS(2002)
WITH fShtat    
     .Caption='Персонал-штатное'
     .procexit='fShtat.Release'  
     DO addcontmenu WITH 'fshtat','menucont1',10,5,'печать','print1.ico',"DO printreport WITH 'peopsh','тарификация-штатное','peopsh'"
     DO addcontmenu WITH 'fshtat','menucont2',.menucont1.Left+.menucont1.Width+3,5,'возврат','undo.ico','fShtat.Release'   
    
     DO addOptionButton WITH 'fShtat',1,'отображать занятые',.menucont1.Top,.menuCont2.Left+.menuCont2.Width+20,'dimOptVac(1)',0,'DO procValOptVac WITH 1',.T.   
     DO addOptionButton WITH 'fShtat',2,'отображать вакансии',.Option1.Top,.Option1.Left+.Option1.Width+10,'dimOptVac(2)',0,'DO procValOptVac WITH 2',.T.
     .Option1.Top=.menuCont1.Top+(.menuCont1.Height-.Option1.Height)/2
     .Option2.Top=.Option1.Top
     
     DO addcombomy WITH 'fShtat',1,.option2.Left+.option2.Width+10,.Option1.Top,dHeight,250,.T.,'','curDatGrup.name',6,'','DO validGrupShtatOld',.F.,.T.
     .AddObject('Grdrasp','GridMy')        
     WITH fshtat.grdrasp
          .Top=fshtat.menucont1.Top+fshtat.menucont1.Height+5
          .Left=0
          .Width=fshtat.Width
          .Height=fshtat.height-.Top
          .ScrollBars=2
          .ColumnCount=max_kat*2+4
          .RecordSource='peopsh'
          .Column1.ControlSource='peopsh.name'
          .Column1.Alignment=0
          .Column1.Header1.Caption='Подразделение'
          .Column2.ControlSource='peopsh.r'
          .Column3.ControlSource='peopsh.p'  
          .Column3.DynamicForeColor='IIF(peopsh->r#peopsh->p,objcolorsos,dForeColor)'  
          .Column3.Alignment=1
          n_ch=1
          FOR i=4 TO .ColumnCount-1  
           *   column_ch='.Column'+LTRIM(STR(i)) 
              IF MOD(i,2)=0
                 p_ch='peopsh.r'+LTRIM(STR(n_ch))         
                 .Columns(i).ControlSource=p_ch
                 .Columns(i).Alignment=1
              ELSE
                 r_ch='peopsh.p'+LTRIM(STR(n_ch))
                 .Columns(i).ControlSource=r_ch 
                 .Columns(i).DynamicForeColor='IIF(&r_ch#&p_ch,objcolorsos,dForeColor)'
                 .Columns(i).Alignment=1
                 n_ch=n_ch+1          
             ENDIF
          ENDFOR  
          .SetAll('Width',FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH('9999999 ',dFontName,dFontSize),'Column')
          .Columns(.ColumnCount).Width=0
          .SetAll('Format','Z','Column') 
          .SetAll('Enabled',.F.,'Column')            
          .Column1.Width=.Width-.Column2.Width*(.ColumnCount-2)-13-.ColumnCount-SYSMETRIC(5)     
     ENDWITH
     DO gridSize WITH 'fshtat','grdrasp','shapeingrid' 
     
     DO mycolumnTxtBox WITH 'fshtat.grdrasp.Column1','tBox1',.F.,.F.
     labwidth=.grdrasp.Column2.Width*2-1
     labheight=.grdrasp.Headerheight+1
     DO adLabMy WITH 'fshtat',1,'всего',.grdrasp.Top+1,.grdrasp.Left+13+.grdrasp.Column1.Width,labwidth,2,.F.,1
     lableft=.lab1.Left+labwidth+2
     SELECT sprkat
     GO TOP
     FOR i=1 TO max_kat    
         labcaption=LOWER(ALLTRIM(sprkat->name))    
         IF TXTWIDTH(labcaption,dFontName,dFontSize)*FONTMETRIC(6,dFontname,dFontsize)>labwidth
            DO WHILE TXTWIDTH(labcaption,dFontName,dFontSize)*FONTMETRIC(6,dFontname,dFontsize)>labwidth 
               labcaption=LEFT(labcaption,LEN(labcaption)-1)
            ENDDO    
         ENDIF  
         DO adLabMy WITH 'fshtat',i+1,labcaption,.grdrasp.Top+1,lableft,labwidth,2,.F.,1
         lableft=lableft+labwidth+3
         SKIP
     ENDFOR
     SELECT peopsh
     .grdrasp.Height=.grdrasp.Height-.grdrasp.RowHeight
     .SetAll('BackColor',.grdrasp.Column1.Header1.backColor,'Labelmy')
     .SetAll('Height',.grdrasp.Headerheight,'LabelMy')
     DO addcontmy WITH 'fshtat','cont11',10,.Grdrasp.Top+.grdrasp.Height-1,.Grdrasp.column1.Width+2,.grdrasp.RowHeight,''
     .cont11.SpecialEffect=2

     lableft=.cont11.Left+.cont11.Width-1
     labwidth=.grdrasp.Column2.Width
     nlab=12 
     FOR i=1 TO max_kat+1 
         labCaption=IIF(dim_tot(i,1)#0,' '+STR(dim_tot(i,1),6,2)+' ','')
         obj_lab=''
         obj_cont=''
         namecont='cont'+LTRIM(STR(nlab)) 
         DO addcontmy WITH 'fshtat',namecont,lableft,.cont11.Top,labwidth+2,.cont11.Height,labCaption
         fshtat.&obj_cont..SpecialEffect=2 
         fshtat.&obj_cont..ContLabel.FontSize=dFontSize         
         lableft=lableft+labwidth+1
         nlab=nlab+1
         obj_lab=''
         labCaption=IIF(dim_tot(i,2)#0,' '+STR(dim_tot(i,2),6,2)+' ','')
         namecont='cont'+LTRIM(STR(nlab))
         DO addcontmy WITH 'fshtat',namecont,lableft,.cont11.Top,labwidth+2,.cont11.Height,labCaption
         fshtat.&obj_cont..SpecialEffect=2  
         nlab=nlab+1
         lableft=lableft+labwidth+1
         IF dim_tot(i,1)#dim_tot(i,2)       
            fshtat.&obj_cont..ContLabel.ForeColor=objcolorsos
            fshtat.&obj_cont..ContLabel.FontSize=dFontSize
         ENDIF
     ENDFOR    
ENDWITH    
fShtat.Show
***********************************************************************************************
PROCEDURE validGrupShtatOld
SELECT peopsh
IF EMPTY(curDatGrup.sostav1)
   SET FILTER TO 
ELSE 
   SET FILTER TO ','+LTRIM(STR(kod))+','$curDatGrup.sostav1
ENDIF
GO TOP
fShtat.Refresh
***********************************************************************************************
PROCEDURE procValOptVac
PARAMETERS par1
STORE 0 TO dimOptVac
dimOptVac(par1)=1
DO inCurPeopSh WITH .T.
SELECT peopsh
GO TOP 
fShtat.Refresh
***********************************************************************************************
PROCEDURE inCurPeopSh
PARAMETERS par1
SELECT peopsh
REPLACE p WITH IIF(dimOptVac(1)=1,0,r),p1 WITH 0,p2 WITH 0,p3 WITH 0,p4 WITH 0,p5 WITH 0,p6 WITH 0,p7 WITH 0 ALL
SELECT rasp
SET FILTER TO kp#0.AND.kd#0
GO TOP
kp_ch=kp
STORE 0 TO dim_kat
STORE 0.00 TO ksetot,dim_tot
DO WHILE !EOF()
   IF rasp.kat#0
      dim_kat(ASCAN(kod_kat,rasp->kat))=dim_kat(ASCAN(kod_kat,rasp->kat))+kse      
   ENDIF   
   ksetot=ksetot+kse    
   SELECT rasp
   SKIP
   IF kp_ch#kp
      SELECT peopsh
      LOCATE FOR kod=kp_ch
      REPLACE r WITH ksetot
      FOR i=1 TO max_kat
          ksepeop=0 
          SELECT curShtatJob
          SUM kse TO ksepeop FOR kp=kp_ch.AND.kat=kod_kat(i)
          SELECT peopsh    
          r_rep='r'+LTRIM(STR(i))
          p_rep='p'+LTRIM(STR(i))
          REPLACE &r_rep WITH dim_kat(i)
          DO CASE 
             CASE dimOptVac(1)=1
                  REPLACE &p_rep WITH ksepeop,p WITH p+ksepeop               
             CASE dimOptVac(2)=1
                  REPLACE &p_rep WITH &r_rep-ksepeop,p WITH p-ksepeop             
          ENDCASE   
      ENDFOR      
      SELECT rasp 
      kp_ch=kp
      STORE 0 TO dim_kat,ksetot
   ENDIF
ENDDO
IF par1
    WITH fShtat.grdRasp
         DO CASE
            CASE dimOptVac(1)=1            
                 .Column3.DynamicForeColor='IIF(peopsh->r#peopsh->p,objcolorsos,dForeColor)'  
                 FOR i=4 TO .ColumnCount-1                     
                     IF MOD(i,2)#0                 
                       .Columns(i).DynamicForeColor='IIF(&r_ch#&p_ch,objcolorsos,dForeColor)'                              
                     ENDIF
                 ENDFOR  
            CASE dimOptVac(2)=1    
                 .Column3.DynamicForeColor='objcolorsos'  
                 FOR i=4 TO .ColumnCount-1                     
                     IF MOD(i,2)#0                 
                       .Columns(i).DynamicForeColor='objcolorsos'                                    
                     ENDIF
                 ENDFOR  
         ENDCASE
    ENDWITH
ENDIF
SELECT peopsh
nlab=12
FOR i=1 TO max_kat+1         
    IF i=1
       SUM r,p TO dim_tot(1,1),dim_tot(1,2)     
    ELSE
       r_rep='r'+LTRIM(STR(i-1))
       p_rep='p'+LTRIM(STR(i-1))
       SUM &r_rep,&p_rep TO dim_tot(i,1),dim_tot(i,2) 
    ENDIF    
    IF par1
       labCaption=IIF(dim_tot(i,1)#0,' '+STR(dim_tot(i,1),6,2)+' ','')
       namecont='cont'+LTRIM(STR(nlab)) 
       fshtat.&namecont..ContLabel.Caption=labCaption
       nlab=nlab+1
       labCaption=IIF(dim_tot(i,2)#0,' '+STR(dim_tot(i,2),6,2)+' ','')
       namecont='cont'+LTRIM(STR(nlab))
       fshtat.&namecont..ContLabel.Caption=labCaption   
       DO CASE
          CASE dimOptVac(1)=1
               IF dim_tot(i,1)#dim_tot(i,2)
                  fshtat.&namecont..ContLabel.ForeColor=objcolorsos
                  fshtat.&namecont..ContLabel.FontSize=dFontSize              
               ENDIF       
          CASE dimOptVac(2)=1
               fshtat.&namecont..ContLabel.ForeColor=objcolorsos
       ENDCASE 
       nlab=nlab+1
    ENDIF  
ENDFOR   
GO TOP
*********************************************************************************************************************
*                                             Тарификация - штатное (новый вариант)
*********************************************************************************************************************
PROCEDURE shtatpeople
SELECT datJob
SET FILTER TO 
IF USED('curDatGrup')
   SELECT curDatGrup
   USE
ENDIF
IF USED('curShtatJob')
   SELECT curShtatJob
   USE
ENDIF
IF USED('peopsh')
   SELECT peopsh
   USE
ENDIF
DIMENSION dim_tot(4)
STORE 0 TO dim_tot

SELECT * FROM datagrup INTO CURSOR curDatGrup READWRITE
SELECT curDatgrup
APPEND BLANK
REPLACE name WITH 'Организация'

SELECT * FROM datjob INTO CURSOR curShtatJob READWRITE
SELECT curShtatJob
REPLACE dekotp WITH IIF(SEEK(kodpeop,'people',1),people.dekotp,dekotp) ALL
DELETE FOR tr=4
IF datShtat.Real
   DELETE FOR dekotp
   DELETE FOR dateBeg>varDTar
   DELETE FOR !EMPTY(dateOut).AND.dateOut<varDTar
ENDIF

SELECT * FROM curShtatJob INTO CURSOR curdjob READWRITE
SELECT curdjob
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
ALTER TABLE curShtatJob ADD COLUMN avtVac L
SELECT curShtatJob
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SELECT rasp
GO TOP
DO WHILE !EOF()
   IF rasp.kse#0
      SELECT curdjob
      SET ORDER TO 1      
      SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
      kse_cx=rasp.kse
      DO WHILE rasp.kp=curdjob.kp.AND.rasp.kd=curdjob.kd.AND.!EOF()
      *   IF date_in>varDtar
      *      ELSE 
         IF datShtat.Real            
            *O CASE
            *  CASE SEEK(kodpeop,'people',1).AND.(!EMPTY(people.date_Out).OR.people.dekotp)
            *  CASE SEEK(kodpeop,'people',1).AND.(!EMPTY(people.date_Out).OR.people.dekotp)
            *  CASE !EMPTY(datJob.dateOut).AND.datJob.dateOut<varDtar
            *  CASE dekOtp
            *  CASE datJob.dateBeg>varDtar
            *  CASE datJob.tr=4
            
             * CASE datJob.dekotp          
             * OTHERWISE 
                   kse_cx=kse_cx-curdjob.kse                   
            *NDCASE
         ELSE    
            kse_cx=kse_cx-curdjob.kse
         ENDIF             
         SKIP
      ENDDO     
      IF kse_cx>0          
         SELECT curShtatJob
         APPEND BLANK
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH 'Вакантная', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,kse WITH kse_cx,avtVac WITH .T.  
      ENDIF
   ENDIF    
   SELECT rasp
   SKIP
ENDDO  
SELECT * FROM sprpodr INTO CURSOR peopsh READWRITE
ALTER TABLE peopsh ADD COLUMN r N(8,2)
ALTER TABLE peopsh ADD COLUMN ksetot N(8,2)
ALTER TABLE peopsh ADD COLUMN stot N(8,2)
ALTER TABLE peopsh ADD COLUMN vactot N(8,2)
ALTER TABLE peopsh ADD COLUMN pOutKf N(8,2)
ALTER TABLE peopsh ADD COLUMN pVOutKf N(8,2)
ALTER TABLE peopsh ADD COLUMN pprim MEMO
ALTER TABLE peopsh ADD COLUMN logf L
ALTER TABLE peopsh ADD COLUMN logfVac L

SELECT peopsh
INDEX ON np TAG T1
SELECT peopsh
SCAN ALL
     kseRasp=0
     SELECT rasp
     SUM kse FOR kp=peopsh.kod TO kseRasp
     SELECT peopsh 
     REPLACE r WITH kseRasp
     SELECT curShtatJob   
     SEEK STR(peopsh.kod,3)   
     ksePeop=0
     ksevac=0
     kseWithOutKf=0
     kseVacWithOutKf=0
     txtPPrim=''
     logRepF=.F.
     logRepFVac=.F.
     DO WHILE kp=peopsh.kod
        ksePeop=ksePeop+IIF(!avtVac,kse,0)
        kseVac=kseVac+IIF(avtVac.AND.namekf#0,kse,0)
        IF namekf=0.AND.!avtVac
           kseWithOutKf=kseWithOutKf+IIF(!avtVac,kse,0)
           logRepF=.T.
           txtPprim=ALLTRIM(fio)+' '+LTRIM(STR(kse,5,2))+' '+'не указан тарифный коэффициент'+CHR(13)
           SELECT peopsh
           REPLACE pprim WITH ALLTRIM(pprim)+txtPprim
           SELECT curShtatJob
        ENDIF 
        IF namekf=0.AND.avtVac
           kseVacWithOutKf=kseVacWithOutKf+IIF(avtVac,kse,0)
           logRepFVac=.T.
           txtPprim='Вакансия - авто - '+IIF(SEEK(kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' - не указан тарифный коэффициент'+CHR(13)
           SELECT peopsh
           REPLACE pprim WITH ALLTRIM(pprim)+txtPprim
           SELECT curShtatJob
        ENDIF
        SKIP 
     ENDDO    
     SELECT peopsh
     REPLACE stot WITH ksePeop,vacTot WITH ksevac,kseTot WITH stot+vacTot,pOutKf WITH kseWithOutkf,pVOutKf WITH kseWithOutkf,logF WITH logRepF,logFVac WITH logRepFVac
     
ENDSCAN 
SUM r,ksetot,stot,vactot TO dim_tot(1),dim_tot(2),dim_tot(3),dim_tot(4)
GO TOP
fShtat=CREATEOBJECT('FORMMY')
=SYS(2002)
WITH fShtat     
     .Caption='Персонал-штатное'
     .procexit='DO exitShtatPeople' 
      
     DO addButtonOne WITH 'fshtat','menuCont1',10,5,'печать','print1.ico','DO formPrintError',39,RetTxtWidth('возвратw')+44,'печать'  
     DO addButtonOne WITH 'fshtat','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'ошибка','warning.ico','DO viewError',39,.menucont1.Width,'ошибка - просмотр'  
     DO addButtonOne WITH 'fshtat','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'возврат','undo.ico','DO exitShtatPeople',39,.menucont1.Width,'возврат'  
     DO addcombomy WITH 'fShtat',1,.menuCont3.Left+.menuCont3.Width+20,.menuCont1.Top,dHeight,250,.T.,'','curDatGrup.name',6,'','DO validGrupShtat',.F.,.T.
    .AddObject('Grdrasp','GridMy')        
     WITH fshtat.grdrasp
          .Top=fshtat.menucont1.Top+fshtat.menucont1.Height+5
          .Left=0
          .Width=fshtat.Width
          .Height=fshtat.height-.Top
          .ScrollBars=2
          .ColumnCount=6
          .RecordSource='peopsh'
          .Column1.ControlSource='peopsh.name'
          .Column1.Alignment=0
          .Column1.Header1.Caption='Подразделение'
          .Column2.Header1.Caption='Ш.р.'
          .Column3.Header1.Caption='Перс.'
          .Column4.Header1.Caption='Сотр.'
          .Column5.Header1.Caption='Вак'
          .Column2.ControlSource='peopsh.r'
          .Column3.ControlSource='peopsh.ksetot'  
          .Column4.ControlSource='peopsh.stot'  
          .Column5.ControlSource='peopsh.vactot' 
        
          .Column3.Alignment=1   
          .Columns(.ColumnCount).Width=0
          .Column2.Width=RetTxtWidth('999999999')
          .Column3.Width=.Column2.Width
          .Column4.Width=.Column2.Width          
          .Column5.Width=.Column2.Width
          .Column2.Format='Z'
          .Column3.Format='Z'
          .Column4.Format='Z'
          .Column5.Format='Z'
          .Column3.Alignment=1
          .Column4.Alignment=1
          .Column5.Alignment=1
          .SetAll('Enabled',.F.,'Column')            
          .Column1.Width=.Width-.Column2.Width*4-13-.ColumnCount-SYSMETRIC(5)     
     ENDWITH
     .grdrasp.Height=.grdrasp.Height-.grdrasp.RowHeight
     DO gridSize WITH 'fshtat','grdrasp','shapeingrid' 
     .grdRasp.Column3.DynamicBackColor='IIF(peopsh.r#peopsh.ksetot,RGB(255,0,0),IIF(RECNO(fShtat.grdRasp.RecordSource)#fShtat.grdRasp.curRec,dBackColor,dynBackColor))'
     .grdRasp.Column4.DynamicBackColor='IIF(peopsh.logF,RGB(251,198,0),IIF(RECNO(fShtat.grdRasp.RecordSource)#fShtat.grdRasp.curRec,dBackColor,dynBackColor))'
     .grdRasp.Column5.DynamicBackColor='IIF(peopsh.logFVac,RGB(251,198,0),IIF(RECNO(fShtat.grdRasp.RecordSource)#fShtat.grdRasp.curRec,dBackColor,dynBackColor))'
           
     DO adtBoxNew WITH 'fshtat','txtBox1',.grdRasp.Top+.grdRasp.Height-1,.grdRasp.Left,.grdRasp.Column1.Width+12,dHeight,'',.F.,.F.,.F.,.F.
     DO adtBoxNew WITH 'fshtat','txtBox2',.txtBox1.Top,.txtBox1.Left+.txtBox1.Width-1,.grdRasp.Column2.Width+2,dHeight,'dim_tot(1)','Z',.F.,.F.,.F.
     DO adtBoxNew WITH 'fshtat','txtBox3',.txtBox1.Top,.txtBox2.Left+.txtBox2.Width-1,.grdRasp.Column3.Width+2,dHeight,'dim_tot(2)','Z',.F.,.F.,.F.
     DO adtBoxNew WITH 'fshtat','txtBox4',.txtBox1.Top,.txtBox3.Left+.txtBox3.Width-1,.grdRasp.Column3.Width+2,dHeight,'dim_tot(3)','Z',.F.,.F.,.F.
     DO adtBoxNew WITH 'fshtat','txtBox5',.txtBox1.Top,.txtBox4.Left+.txtBox2.Width-1,.grdRasp.Column4.Width+2,dHeight,'dim_tot(4)','Z',.F.,.F.,.F.    
     IF dim_tot(1)#dim_tot(2)  
        .txtBox3.Enabled=.F.     
        .txtBox3.DisabledBackColor=RGB(255,0,0)
        .txtBox3.BackStyle=1
     ENDIF
     LOCATE FOR logF
     IF FOUND()
        .txtBox4.Enabled=.F.     
        .txtBox4.DisabledBackColor=RGB(251,198,0)
        .txtBox4.BackStyle=1
     ENDIF        
     LOCATE FOR logFvAC
     IF FOUND()
        .txtBox5.Enabled=.F.     
        .txtBox5.DisabledBackColor=RGB(251,198,0)
        .txtBox5.BackStyle=1
     ENDIF        
     SELECT peopsh  
ENDWITH    
fShtat.Show
***********************************************************************************************
PROCEDURE validGrupShtat
SELECT peopsh
IF EMPTY(curDatGrup.sostav1)
   SET FILTER TO 
ELSE 
   SET FILTER TO ','+LTRIM(STR(kod))+','$curDatGrup.sostav1
ENDIF
SUM r,ksetot,stot,vactot TO dim_tot(1),dim_tot(2),dim_tot(3),dim_tot(4)
WITH fshtat
     .setAll('Enabled',.F.,'myTxtBox')
     .setAll('DisabledBackColor',ObjBackColor,'myTxtBox')
     .setAll('BackStyle',0,'myTxtBox')
     .txtBox2.ContRolSource='dim_tot(1)'
     .txtBox3.ContRolSource='dim_tot(2)'
     .txtBox4.ContRolSource='dim_tot(3)'
     .txtBox5.ContRolSource='dim_tot(4)'
     IF dim_tot(1)#dim_tot(2)  
        .txtBox3.Enabled=.F.     
        .txtBox3.DisabledBackColor=RGB(255,0,0)
        .txtBox3.BackStyle=1      
     ENDIF
     LOCATE FOR logF
     IF FOUND()
        .txtBox4.Enabled=.F.     
        .txtBox4.DisabledBackColor=RGB(251,198,0)
        .txtBox4.BackStyle=1
     ENDIF        
     LOCATE FOR logFvAC
     IF FOUND()
        .txtBox5.Enabled=.F.     
        .txtBox5.DisabledBackColor=RGB(251,198,0)
        .txtBox5.BackStyle=1
     ENDIF      
ENDWITH 
GO TOP
fShtat.Refresh
*********************************************************************************************************************
PROCEDURE formPrintError
kvo_page=1
page_beg=1
page_end=999
term_ch=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl    
     .Caption='Печать'
     .Width=400
     DO adSetupPrnToForm WITH 10,10,400,.F.,.F.     
     .Width=.Shape91.Width+20
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnError WITH .T.' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'Просмотр','DO prnError WITH .F.'
      *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','fSupl.Release','Выход из печати'
     .Height=.Shape91.Height+.cont1.Height+50
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*********************************************************************************************************************
PROCEDURE prnError
PARAMETERS par1
term_ch=par1
CREATE CURSOR curError (nameError memo, kp N(3))
SELECT peopsh
oldRec=RECNO()
SCAN ALL
     IF peopSh.logF.OR.peopSh.logFVac
        SELECT curError
        APPEND BLANK
        REPLACE nameError WITH peopsh.pprim,kp WITH peopsh.kod
     ENDIF 
     SELECT peopsh
ENDSCAN
SELECT curError 
GO TOP
DO procForPrintAndPreview WITH 'repError','возможные ошибки в штатном расписании и персонале',term_ch
SELECT peopsh
GO oldRec
*********************************************************************************************************************
PROCEDURE viewError
SELECT peopSh
IF peopSh.logF.OR.peopSh.logFVac
   fSupl=CREATEOBJECT('FORMSUPL')
   WITH fSupl        
        .Caption='Возможные ошибки'
        .Width=500
        .Height=300  
        DO adeditbox WITH 'fSupl','boxError',0,0,.Width,.Height,'peopsh.pprim',.F.,0,.F.,.F.,.F.,.F.      
   ENDWITH   
   DO pasteImage WITH 'fSupl'
   fSupl.Show
ENDIF
*********************************************************************************************************************
PROCEDURE exitShtatPeople
SELECT peopSh
SET FILTER TO
USE
SELECT curShtatJob
USE
fShtat.Release
SELECT people
