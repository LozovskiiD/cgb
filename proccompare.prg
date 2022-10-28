IF USED('olddatjob')
   SELECT olddatjob
   USE
ENDIF 
IF USED('oldpeople')
   SELECT oldpeople
   USE
ENDIF 
IF USED('curJobNew')
   SELECT curJobNew
   USE
ENDIF
IF USED('curJobOld')
   SELECT curJobOld
   USE
ENDIF

IF USED('dcompare')
   SELECT dcompare
   USE
ENDIF

IF !USED('comFond')
   USE comFond ORDER 1 IN 0
ENDIF

IF !USED('comprn')
   USE comprn ORDER 1 IN 0
ENDIF

IF USED('curFltPodr')
   SELECT curFltPodr
   USE
ENDIF

IF USED('fondold')
   SELECT fondold
   USE
ENDIF

SELECT * FROM sprpodr INTO CURSOR curFltPodr ORDER BY name READWRITE
SELECT curFltPodr
REPLACE fl WITH .F., otm WITH '' ALL

SELECT * FROM comFond INTO CURSOR curFondNew ORDER BY num READWRITE

USE dcompare ORDER 1 IN 0
DIMENSION dcom_old(FCOUNT('dcompare'))

SELECT datjob
SET FILTER TO
SET ORDER TO 6
formulaold=''
formulanew=''
fltpodr=''

=AFIELDS(arPeop,'people')
CREATE CURSOR curSuplPeop FROM ARRAY arPeop
SELECT curSuplPeop
INDEX ON fio TAG t1

=AFIELDS(dimjob,'datjob')
CREATE CURSOR curjobnew FROM ARRAY dimjob
pathcompare=pathmain+'\'+ALLTRIM(datset.pathcomp)+'\'
pathpeopold=pathcompare+'people.dbf'
pathjobold=pathcompare+'datjob.dbf'
pathfondold=pathcompare+'comfond.dbf'

IF ALLTRIM(datset.pathlast)#ALLTRIM(datset.pathcomp)
   IF FILE(pathpeopold)
      USE &pathpeopold ALIAS oldpeople IN 0 
   ENDIF 
   IF FILE(pathjobold)
      USE &pathjobold ALIAS olddatjob IN 0 ORDER 7 &&nid
   ENDIF 
   
   IF FILE(pathfondold)
      USE &pathfondold ALIAS fondOld IN 0 
   ENDIF 
   SELECT * FROM fondOld INTO CURSOR curFondOld ORDER BY num READWRITE   
ENDIF    
SELECT datshtat
LOCATE FOR ALLTRIM(pathtarif)==ALLTRIM(datset.pathlast)
varNmzp=nmzp
cround=IIF(vround#0,vround,1)

=AFIELDS(dimjob,'olddatjob')
CREATE CURSOR curjobold FROM ARRAY dimjob
SELECT curjobold
APPEND FROM DBF ('olddatjob') FOR kodpeop=people.num
INDEX ON STR(tr,1)+STR(kse,5,2) TAG T1
SET ORDER TO 1
GO TOP
=AFIELDS(dimjob,'datjob')
CREATE CURSOR curjobnew FROM ARRAY dimjob
SELECT curjobnew
APPEND FROM datjob FOR kodpeop=people.num
INDEX ON STR(tr,1)+STR(kse,5,2) TAG T1
SET RELATION TO tr INTO sprtype,kd INTO sprdolj,kp INTO sprpodr
GO TOP
oldFioPeop=people.fio
newFioPeop=people.fio
oldNum=people.num
newNum=people.num
logAp=.F.
STORE 0 TO sumItOld,sumItNew,newBdpl

SELECT dcompare 
SEEK curJobNew.nid
&& sumItOld старая сумма (mtokl+mstsum+mslwork+mkonts+mprem+mozd+matt)
&& sumItNew новая сумма (mtokl+mstsum+mslwork+mkonts+mprem+mozd+matt+mbdopl)
&& newBdpl новая базовая доплата
SELECT curFondOld
SCAN ALL      
     IF !EMPTY(sum_f)
        formulaold=formulaold+'+'+ALLTRIM(sum_f)        
     ENDIF
ENDSCAN
formulaold=SUBSTR(formulaold,2)
GO TOP

SELECT curFondNew
APPEND BLANK
REPLACE num WITH 90,rec WITH 'Разница',fname WITH 'difsum'
APPEND BLANK
REPLACE num WITH 91,rec WITH 'Новая надбавка за сложность и напряженность',fpers WITH 'newpsl',fname WITH 'newmsl'
SCAN ALL  
     IF !EMPTY(sum_f)
        formulanew=formulanew+'+'+ALLTRIM(sum_f)        
     ENDIF
ENDSCAN
formulanew=SUBSTR(formulanew,2)
GO TOP
fCompare=CREATEOBJECT('FORMSUPL')
tarcomparesay='тарификация для сравнения на '+ALLTRIM(datset.cdcomp)+' (изменить - двойной щелчок мыши)'
fltsay='фильтр - организация(изменить - двойной щелчок мыши)'
newcompareSay=''
WITH fCompare
     .Icon='money.ico'
     .Caption='сравнительная таблица'
     .procForClick='DO lostFocusPeop'
     .procexit='DO exitcompare'
     .Width=frmTop.Width
     .Height=frmTop.Height
     DO addContFormNew WITH 'fCompare','txtPath',0,0,(.Width-10)/2,dHeight,tarcomparesay,2,.F.,'DO choiceCompareTar',.F.,.F.            
     DO addContFormNew WITH 'fCompare','txtNew',.txtPath.Left+.txtPath.Width+11,0,.txtPath.Width,dHeight,fltsay,2,.F.,'DO fltcompare',.F.,.F.            
     DO adtBoxAsCont WITH 'fCompare','contOld',0,.txtPath.Top+.txtPath.Height-1,.txtPath.Width,dHeight,'действующие условия',2,0   
     DO adtBoxAsCont WITH 'fCompare','contNew',.txtNew.Left,.contOld.Top,.txtNew.Width,dHeight,'новые условия',2,0   
     DO adTboxNew WITH 'fCompare','tBoxFioOld',.contOld.Top+.contOld.Height-1,.contOld.Left,.contOld.Width-RetTxtWidth('w...')-2,dHeight,'oldFioPeop',.F.,.T.,0      
     DO adtboxnew WITH 'fCompare','boxFreeOld',.tBoxFioOld.Top,.tBoxFioOld.Left+.tBoxFioOld.Width-1,.contold.Width-.tBoxFioOld.Width+1,dheight,'',.F.,.T.   
     DO addButtonOne WITH 'fCompare','butOld',.boxFreeOld.Left+1,.boxFreeOld.Top+2,'','sbdn.ico','DO selectPeopOld',.boxFreeOld.Height-4,RetTxtWidth('w...')-1,'' 
   
     DO adTboxNew WITH 'fCompare','tBoxFioNew',.tBoxFioOld.Top,.contNew.Left,.contNew.Width-RetTxtWidth('w...')-2,dHeight,'newFioPeop',.F.,.T.,0 
     .tBoxFioNew.procforChange='DO changePeopNew'     
     DO adtboxnew WITH 'fCompare','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.contNew.Width-.tBoxFioNew.Width+1,dheight,'',.F.,.T.   
     DO addButtonOne WITH 'fCompare','butNew',.boxFreeNew.Left+1,.boxFreeNew.Top+2,'','sbdn.ico','DO selectPeopNew',.boxFreeNew.Height-4,RetTxtWidth('w...')-1,'' 
     
     .AddObject('grdNew','GridMyNew')      
     WITH .grdNew      
          .Top=.Parent.tBoxFioNew.Top+.Parent.tBoxFioNew.Height-1
          .Left=.Parent.tBoxFioNew.Left
          .Height=.rowHeight*5
          .Width=.Parent.contNew.Width
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curJobNew'
           DO addColumnToGrid WITH 'fCompare.grdNew',5
          .Column1.ControlSource="IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'')" 
          .Column2.ControlSource="IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')"    
          .Column3.ControlSource='curJobNew.kse'
          .Column4.ControlSource="IIF(SEEK(tr,'sprtype',1),sprtype.name,'')"
          
          
          .Column1.Header1.Caption='подразделение'
          .Column2.Header1.Caption='должность'  
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'  
          .Column3.Width=RettxtWidth('99999.99')      
          .Column4.Width=RettxtWidth('онсновнаяw')      
          .Columns(.ColumnCount).Width=0          
          .Column1.Width=(.Width-.column3.width-.Column4.Width)/2
          .Column2.Width=.Width-.column1.width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount 
          .Column1.Alignment=0        
          .Column2.Alignment=0
          .Column1.Movable=.F.         
          .colNesInf=2    
          .procAfterRowColChange='DO compareScr'  
         .SetAll('BOUND',.F.,'Column')  
         .Visible=.T.                        
    ENDWITH       
    DO gridSizeNew WITH 'fCompare','grdNew','shapeingrid',.F.,.T. 
    
    .AddObject('grdOld','GridMyNew')      
    WITH .grdOld      
          .Top=.Parent.grdNew.Top
          .Left=0
          .Height=.Parent.grdNew.Height
          .Width=.Parent.grdNew.Width
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curJobOld'
           DO addColumnToGrid WITH 'fCompare.grdOld',5
          
          .Column1.ControlSource="IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'')" 
          .Column2.ControlSource="IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')"    
          .Column3.ControlSource='curJobOld.kse'
          .Column4.ControlSource="IIF(SEEK(tr,'sprtype',1),sprtype.name,'')"
                 
          .Column1.Header1.Caption='подразделение'
          .Column2.Header1.Caption='должность'  
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'  
          .Column3.Width=RettxtWidth('99999.99')      
          .Column4.Width=RettxtWidth('онсновнаяw')      
          .Columns(.ColumnCount).Width=0          
          .Column1.Width=(.Width-.column3.width-.Column4.Width)/2
          .Column2.Width=.Width-.column1.width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount 
          .Column1.Alignment=0        
          .Column2.Alignment=0
          .Column1.Movable=.F.         
          .colNesInf=2    
          .procAfterRowColChange='DO compareOldscr'  
         .SetAll('BOUND',.F.,'Column')  
         .Visible=.T.                        
    ENDWITH       
    DO gridSizeNew WITH 'fCompare','grdOld','shapeingrid1',.F.,.T. 
        
    .AddObject('grd1','GridMyNew')      
    WITH .grd1      
         .Top=.Parent.grdNew.Top+.Parent.grdNew.Height-1
         .Left=0
         .Height=.Parent.Height-.Top-dHeight*3
         .Width=.Parent.grdNew.Width
         .ScrollBars=2      	           
      
         .RecordSource='curFondOld'
         
         DO addColumnToGrid WITH 'fCompare.grd1',4
         
         .Column1.ControlSource='ALLTRIM(curfondold.rec)'
         .Column2.ControlSource='curfondold.say1'
         .Column3.ControlSource='curfondold.say2'               
         .Column1.Header1.Caption='строка'
         .Column2.Header1.Caption='%'  
         .Column3.Header1.Caption='сумма'
         .Columns(.ColumnCount).Width=0          
         .Column3.Width=RetTxtWidth('999999999')
         .Column2.Width=.Column3.Width
         .Column1.Width=.Width-.column2.width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount 
         .Column1.Alignment=1        
         .Column2.Format='Z'
         .Column3.Format='Z'
         .Column2.Alignment=1
         .Column3.Alignment=1
         .colNesInf=2      
         .procAfterRowColChange='DO oldchange'
               
         DO myColumnTxtBox WITH 'fCompare.grd1.column2','tbox3','curfondold.say1',.F.,.F.,.F.,'DO validGrd1'       
         DO myColumnTxtBox WITH 'fCompare.grd1.column3','tbox3','curfondold.say2',.F.,.F.,.F.,'DO validGrd12'       
         .Column3.Sparse=.T.  
         .Column2.Sparse=.T.        
         .SetAll('BOUND',.F.,'ColumnMy')  
         .Visible=.T.                        
    ENDWITH       
  
    DO gridSizeNew WITH 'fCompare','grd1','shapeingrid3',.F.,.T. 
   
   
    .AddObject('grd2','GridMyNew')      
    WITH .grd2
         .Top=.Parent.grd1.Top
         .Left=.Parent.grdNew.Left
         .Height=.Parent.grd1.Height
         .Width=.Parent.grdNew.Width
         .ScrollBars=2       	           
         .RecordSourceType=1     
         .RecordSource='curFondNew'
          DO addColumnToGrid WITH 'fCompare.grd2',4
          
         .Column1.ControlSource='curfondnew.say1'
         .Column2.ControlSource='curfondnew.say2'
         .Column3.ControlSource='ALLTRIM(curfondnew.rec)'
         
         .Column1.Header1.Caption='%'  
         .Column2.Header1.Caption='сумма'               
         .Column3.Header1.Caption='строка'
         
         .Columns(.ColumnCount).Width=0          
         .Column1.Width=RetTxtWidth('999999999')
         .Column2.Width=.Column2.Width
         .Column3.Width=.Width-.column1.width-.Column2.Width-SYSMETRIC(5)-13-.ColumnCount 
         .Column3.Alignment=0        
         .Column1.Format='Z'
         .Column2.Format='Z'
         .Column1.Alignment=1
         .Column2.Alignment=1
         .colNesInf=2    
         .procAfterRowColChange='DO newchange'
         DO myColumnTxtBox WITH 'fCompare.grd2.column1','txtbox1','curFondNew.say1',.F.,.F.,.F.,'DO validGrd2' 
         .Column1.Sparse=.T.
         .SetAll('BOUND',.F.,'ColumnMy')  
         .Visible=.T.                        
    ENDWITH       
      
    DO gridSizeNew WITH 'fCompare','grd2','shapeingrid4',.F.,.T.          

    
    DO addButtonOne WITH 'fCompare','butRead',(.Width-RetTxtWidth('wнастройкиw')*5-40)/2,.grd1.Top+.grd1.Height+20,'расчёт','','DO readCompare',39,RetTxtWidth('wнастройкиw'),'расчёт'  
    DO addButtonOne WITH 'fCompare','butPrn',.butRead.Left+.butRead.Width+10,.butRead.Top,'печать','','DO formPrnCompare',39,.butRead.Width,'печать'  
    DO addButtonOne WITH 'fCompare','butTot',.butPrn.Left+.butPrn.Width+10,.butRead.Top,'перерасчет','','DO formTotCompare',39,.butRead.Width,'перерасчёт'  
    DO addButtonOne WITH 'fCompare','butSet',.butTot.Left+.butTot.Width+10,.butRead.Top,'настройки','','DO setupCompare',39,.butRead.Width,'настройки'  
    DO addButtonOne WITH 'fCompare','butRet',.butSet.Left+.butSet.Width+10,.butRead.Top,'возврат','','DO exitCompare',39,.butRead.Width,'возврат'  

    DO addButtonOne WITH 'fCompare','butSave',(.Width-RetTxtWidth('wзаписатьw')*2-10)/2,.butRead.Top,'записать','','DO SaveCompare WITH .T.',39,RetTxtWidth('wзаписатьw'),'записать'  
    DO addButtonOne WITH 'fCompare','butRef',.butSave.Left+.butSave.Width+10,.butSave.Top,'отказ','','DO SaveCompare WITH .F.',39,.butSave.Width,'отказ'  
    .butSave.Visible=.F.
    .butRef.Visible=.F.
          
    DO addListBoxMy WITH 'fCompare',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.contnew.Width  
    WITH .listBox2
         .RowSource='curSuplpeop.fio'         
         .RowSourceType=2
         .Visible=.F.        
         .procForDblClick='DO validListPeopNew'
         .procForLostFocus='DO lostFocusPeop'
         .Height=.Parent.Height-.Top
         .Enabled=.T.
    ENDWITH           
    DO compareScr
    SELECT curJobNew
    GO TOP
    .grdOld.Columns(.grdOld.ColumnCount).SetFocus   
    .grdNew.Columns(.grdNew.ColumnCount).SetFocus             
    .Autocenter=.T.
ENDWITH
fCompare.Show
**************************************************************************************************************
PROCEDURE choiceCompareTar
IF USED('curDatShtat')
   SELECT curDatShtat
   USE
ENDIF
SELECT * FROM datshtat INTO CURSOR curDatShtat READWRITE
SELECT curDatShtat
DELETE FOR ALLTRIM(pathtarif)==ALLTRIM(datset.pathlast)
ALTER TABLE curDatShtat ADD COLUMN nameSupl C(70)
INDEX ON real TAG T1 DESCENDING
REPLACE namesupl WITH IIF(real,DTOC(DATE()),DTOC(dTarif))+' '+ALLTRIM(fullName) ALL
LOCATE FOR ALLTRIM(pathtarif)==ALLTRIM(datset.pathcomp)
strDate=namesupl
varPathSupl=datset.pathcomp
varCsupl=datset.cdcomp
varStOld=datset.bstold
*varBaseStSupl=varBaseSt
*strDate=namesupl
fCompare1=CREATEOBJECT('FORMSUPL')
WITH fCompare1
     .Icon='money.ico'
     .Caption='Выбор даты тарификации'
     DO addshape WITH 'fCompare1',1,20,20,150,380,8 
     DO adtBoxAsCont WITH 'fCompare1','contDate',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('wвыберите тарификацию для сравненияw'),dHeight,'выберите тарификацию для сравнения',2,1     
     DO addcombomy WITH 'fCompare1',1,.contDate.Left+.contdate.Width-1,.contDate.Top,dHeight,300,.T.,'strDate','ALLTRIM(curDatShtat.namesupl)',6,'','DO validDcompare',.F.,.T.
     .comboBox1.DisplayCount=15
     .Shape1.Width=.contDate.Width+.comboBox1.Width+40
     .Shape1.height=.comboBox1.Height+40
     *-----------------------------Кнопка применить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fCompare1','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприменитьw')*2)-20)/2,;
       .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприменитьw'),dHeight+5,'Применить','DO procChangeCompare'
     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fCompare1','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','fCompare1.Release','Возврат'
     .Width=.Shape1.Width+40    
          
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.cont1.Height+60 
        
ENDWITH
DO pasteImage WITH 'fCompare1'
fCompare1.Show
*************************************************************************************
PROCEDURE validDcompare
varPathSupl=curDatShtat.pathTarif
varCsupl=DTOC(curDatShtat.dtarif)
varStOld=curdatshtat.basest
*varBaseStSupl=curDatShtat.baseSt
*************************************************************************************
PROCEDURE procChangeCompare
SELECT datset
REPLACE pathcomp WITH varPathSupl,cdcomp WITH varCsupl,bstold WITH varStOld
tarcomparesay='тарификация для сравнения на '+ALLTRIM(datset.cdcomp)+' (изменить - двойной щелчок мыши)'
fCompare.txtPath.ContLabel.Caption=tarcomparesay
IF USED('olddatjob')
   SELECT olddatjob
   USE
ENDIF 
IF USED('oldpeople')
   SELECT oldpeople
   USE
ENDIF 
pathcompare=pathmain+'\'+ALLTRIM(datset.pathcomp)+'\'
pathpeopold=pathcompare+'people.dbf'
pathjobold=pathcompare+'datjob.dbf'
IF FILE(pathpeopold)
   USE &pathpeopold ALIAS oldpeople IN 0
ENDIF 
IF FILE(pathjobold)
   USE &pathjobold ALIAS olddatjob ORDER 7 IN 0
ENDIF 
fCompare.Refresh
fCompare1.Release
********************************************************************************************
PROCEDURE exitCompare
SELECT datjob
SET ORDER TO 4
SELECT dcompare
USE
SELECT comprn
USE
IF USED('olddatjob')
   SELECT olddatjob
   USE
ENDIF 
IF USED('oldpeople')
   SELECT oldpeople
   USE
ENDIF 
IF USED('curJobNew')
   SELECT curJobNew
   USE
ENDIF
IF USED('curJobOld')
   SELECT curJobOld
   USE
ENDIF

IF USED('comFond')
   SELECT comFond
   USE
ENDIF

IF USED('fondOld')
   SELECT fondOld
   USE
ENDIF

IF USED('curFondOld')
   SELECT curFondOld
   USE
ENDIF

IF USED('curFondNew')
   SELECT curFondNew
   USE
ENDIF

SELECT people
fCompare.Release
*********************************************************
PROCEDURE selectPeopNew
SELECT curSuplPeop
ZAP
IF EMPTY(fltpodr)
   APPEND FROM people
ELSE 
   APPEND FROM people FOR SEEK(num,'datjob',1)
ENDIF    
WITH fCompare     
     SELECT curSuplPeop
     LOCATE FOR fio=fCompare.tBoxFioNew.Text
     IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.  
        .listBox2.SetFocus            
     ENDIF      
ENDWITH 
**********************************************************
PROCEDURE changePeopNew
WITH fCompare
     IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=fCompare.tBoxFioNew.Text 
SELECT curSuplPeop
ZAP
IF EMPTY(fltpodr)
   APPEND FROM people FOR LEFT(LOWER(fio),LEN(ALLTRIM(lcValue)))=LOWER(ALLTRIM(lcValue))
ELSE 
   APPEND FROM people FOR SEEK(num,'datjob',1).AND.LEFT(LOWER(fio),LEN(ALLTRIM(lcValue)))=LOWER(ALLTRIM(lcValue))
ENDIF    
WITH fCompare.listBox2  
     .RowSource='curSuplpeop.fio'         
     .Visible=IIF(RECCOUNT('curSuplPeop')=0,.F.,.T.)      
ENDWITH 

**********************************************************
PROCEDURE validListPeopNew
fCompare.listBox2.Visible=.F.
newFioPeop=curSuplPeop.fio
newNum=curSuplPeop.num
IF SEEK(newNum,'oldpeople',1)
   oldNum=newNum
   oldFioPeop=oldpeople.fio
ELSE 
   oldNum=0
   oldFioPeop=''    
ENDIF 
SELECT curjobold
ZAP
APPEND FROM DBF ('olddatjob') FOR kodpeop=oldNum
GO TOP

SELECT curjobnew
ZAP
APPEND FROM datjob FOR kodpeop=newNum
GO TOP
SELECT curJobOld
*LOCATE FOR kp=curJobNew.kp.AND.kd=curJobNew.kd.AND.tr=curJobNew.tr.AND.kse=curJobNew.kse
fCompare.Refresh
fCompare.grdOld.Columns(fCompare.grdOld.ColumnCount).SetFocus   
fCompare.grdNew.Columns(fCompare.grdNew.ColumnCount).SetFocus 
**********************************************************          
PROCEDURE lostFocusPeop
WITH fCompare
     ON ERROR DO erSup  
     .listBox1.Visible=.F.
     .listBox2.Visible=.F.    
     .tBoxFioNew.controlSource='newFioPeop'
     .tBoxFioNew.Refresh 
     ON ERROR  
ENDWITH
*********************************************************
PROCEDURE selectPeopOld
*********************************************************
PROCEDURE compareScr
SELECT dcompare
SET ORDER TO 1
SEEK curjobnew.nid
SELECT curFondNew
SCAN ALL    
     rep1=IIF(!EMPTY(fpers),'dcompare.'+ALLTRIM(fpers),'')
     rep2=IIF(!EMPTY(fname),'dcompare.'+ALLTRIM(fname),'') 
     REPLACE say1 WITH IIF(!EMPTY(rep1),EVALUATE(rep1),0)
     REPLACE say2 WITH IIF(!EMPTY(rep2),EVALUATE(rep2),0)     
ENDSCAN
SELECT curFondNew
GO TOP
fCompare.Refresh
***********************************************************
PROCEDURE compareOldScr
SELECT curJobOld
sumItOld=&formulaold
REPLACE msf WITH sumitold
SELECT curFondOld
REPLACE say1 WITH 0,say2 WITH 0 ALL
SCAN ALL    
     rep1=''
     rep2=''
     rep1=IIF(!EMPTY(curfondold.fpers),'curjobold.'+ALLTRIM(fpers),'')
     rep2=IIF(!EMPTY(curfondold.fname),'curjobold.'+ALLTRIM(fname),'')
     REPLACE say1 WITH IIF(!EMPTY(rep1),EVALUATE(rep1),0)
     REPLACE say2 WITH IIF(!EMPTY(rep2),&rep2,0)     
ENDSCAN
GO TOP 
*********************************************************
PROCEDURE validPslOld
REPLACE curJobOld.mslwork WITH curJobold.mtokl*curJobOld.pSlWork/100,curJobOld.slwork WITH curJobold.tokl*curJobOld.pSlWork/100
*sumItOld=curJobOld.mtokl+curJobOld.mstsum+curJobOld.mkonts+curJobOld.mslwork+curJobOld.mprem+curJobOld.mmat+curJobOld.mozd+curJobOld.mHigh
sumItOld=curJobOld.mtokl+curJobOld.mstsum+curJobOld.mkonts+curJobOld.mslwork+curJobOld.mprem+curJobOld.mmat+curJobOld.mozd
newBdpl=varNmzp*curjobnew.kse-curjobnew.mtokl-curjobnew.mStSum-curjobnew.mKonts
newBdpl=IIF(newBdpl<0,0,newBdpl)
REPLACE dcompare.mbdpl WITH newBdpl
sumItNew=dcompare.mtokl+dcompare.mstsum+dcompare.mkonts+dcompare.mprem+dcompare.mmat+dcompare.mozd+dcompare.mbdpl
REPLACE dcompare.newpsl WITH 0,dcompare.newmsl WITH 0
REPLACE dcompare.itog WITH sumItNew,dcompare.difsum WITH dcompare.itog-sumitOld,dcompare.mbdpl WITH newBdpl
DO CASE
   CASE mbdpl>0.AND.difSum>0
        REPLACE dcompare.newpsl WITH 0,dcompare.newmsl WITH 0                       
   CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0
        reppsl=0    
        DO CASE
           CASE cround=1 &&без округления
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
           CASE cround=2 &&по правилам 
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
           CASE cround=3 &&в большую сторону
                reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
        ENDCASE
        IF curjobold.pslwork=0
           reppsl=0
        ENDIF
        REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
   CASE difsum<0.AND.dcompare.mtokl#0 
        reppsl=0
        DO CASE
           CASE cround=1 &&без округления
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
           CASE cround=2 &&по правилам 
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
           CASE cround=3 &&в большую сторону
                reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
        ENDCASE
        REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100   
                       
ENDCASE 
fCompare.Refresh
***********************************************************
PROCEDURE validPhlOld
REPLACE curJobOld.mhigh WITH curJobold.mtokl*curJobOld.pHigh/100,curJobOld.shigh WITH curJobold.tokl*curJobOld.pHigh/100
sumItOld=curJobOld.mtokl+curJobOld.mstsum+curJobOld.mkonts+curJobOld.mslwork+curJobOld.mprem+curJobOld.mmat+curJobOld.mozd+curJobOld.mHigh
*newBdpl=varNmzp*curjobold.kse-curjobold.mtokl-curjobold.mStSum-curjobold.mKonts
newBdpl=varNmzp*curjobnew.kse-curjobnew.mtokl-curjobnew.mStSum-curjobnew.mKonts
*newBdpl=varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts
newBdpl=IIF(newBdpl<0,0,newBdpl)
REPLACE dcompare.mbdpl WITH newBdpl
sumItNew=dcompare.mtokl+dcompare.mstsum+dcompare.mkonts+dcompare.mprem+dcompare.mmat+dcompare.mozd+dcompare.mbdpl
REPLACE dcompare.newpsl WITH 0,dcompare.newmsl WITH 0
REPLACE dcompare.itog WITH sumItNew,dcompare.difsum WITH dcompare.itog-sumitOld,dcompare.mbdpl WITH newBdpl
DO CASE
   CASE mbdpl>0.AND.difSum>0
        REPLACE dcompare.newpsl WITH 0,dcompare.newmsl WITH 0                       
   CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0
        REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
   CASE difsum<0.AND.dcompare.mtokl#0 
        REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100                  
ENDCASE 
fCompare.Refresh
***********************************************************
PROCEDURE validPslNew
SELECT dcompare
REPLACE dcompare.mslwork WITH dcompare.mtokl*dcompare.pSlWork/100,dcompare.slwork WITH dcompare.tokl*dcompare.pSlWork/100
newBdpl=varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts
*newBdpl=varNmzp*curjobold.kse-curjobold.mtokl-curjobold.mStSum-curjobold.mKonts
*newBdpl=varNmzp*curjobnew.kse-curjobnew.mtokl-curjobnew.mStSum-curjobnew.mKonts
newBdpl=IIF(newBdpl<0,0,newBdpl)
REPLACE dcompare.mbdpl WITH newBdpl
sumItNew=dcompare.mtokl+dcompare.mstsum+dcompare.mkonts+dcompare.mprem+dcompare.mmat+dcompare.mozd+dcompare.mbdpl
REPLACE newpsl WITH 0,newmsl WITH 0
REPLACE itog WITH sumItNew,difsum WITH itog-sumitOld,mbdpl WITH newBdpl
DO CASE
   CASE mbdpl>0.AND.difSum>0
        REPLACE newpsl WITH 0,newmsl WITH 0                       
   CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0
   
        REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
   CASE difsum<0.AND.dcompare.mtokl#0 
        REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100                  
ENDCASE       
WITH fCompare
     .tOld217.Refresh
     .tOld222.Refresh
     .tOld821.Refresh
     .tOld822.Refresh
     .tOld823.Refresh
     .Refresh
ENDWITH 
***********************************************************
PROCEDURE readCompare
WITH fCompare
     SELECT dcompare
     SEEK curJobNew.nid 
     logAp=IIF(!FOUND(),.T.,.F.)
     .butRead.Visible=.F.
     .butPrn.Visible=.F.
     .butTot.Visible=.F.
     .butSet.Visible=.F.
     .butRet.Visible=.F.
     .butSave.Visible=.T.
     .butRef.Visible=.T.
     .grd1.Enabled=.T.
     .grd1.Column1.Enabled=.F.
   *  .grd1.Column3.Enabled=.F.
     .grd1.Column4.Enabled=.F.
     
     .grd2.Enabled=.T.
     .grd2.Column2.Enabled=.F.
     .grd2.Column3.Enabled=.F.
     .grd2.Column4.Enabled=.F.
      
     .butNew.Enabled=.F.
     .tBoxFioNew.Enabled=.F.
     .grdOld.Enabled=.F.         
     .grdNew.Enabled=.F.      
     IF logAp
        APPEND BLANK
        REPLACE kodpeop WITH curjobnew.kodpeop,nidjob WITH curJobNew.nid,nidpeop WITH curJobNew.nidpeop,kp WITH curJobNew.kp,kd WITH curJobNew.kd,bst WITH varBaseSt,kse WITH curJobNew.kse,;
                nidold WITH curjobold.nid   
        SELECT curFondNew      
        SCAN ALL
             IF num<90
                rep1=IIF(!EMPTY(fpers),ALLTRIM(curFondNew.fpers),'')
                rep2=IIF(!EMPTY(fname),ALLTRIM(curFondNew.fname),'') 
                rep11=IIF(!EMPTY(fpers),'curjobnew.'+ALLTRIM(curFondNew.fpers),'')
                rep12=IIF(!EMPTY(fname),'curjobnew.'+ALLTRIM(curFondNew.fname),'') 
                SELECT dcompare
                IF !EMPTY(rep1)
                   REPLACE &rep1 WITH IIF(!EMPTY(rep11),EVALUATE(rep11),0)
                ENDIF
                IF !EMPTY(rep2)
                   REPLACE &rep2 WITH IIF(!EMPTY(rep12),EVALUATE(rep12),0)
                ENDIF
         
             ENDIF  
             SELECT curfondnew                          
        ENDSCAN 
     
        SELECT curfondnew
        GO TOP
                
        DO countbdpl
        DO compareScr
        GO TOP                                    
     ELSE      
        SELECT curfondnew
        GO TOP
        DO countbdpl
        DO compareScr
        SELECT dcompare
        SCATTER TO dcom_old
        SELECT  curfondnew
        GO TOP
     ENDIF  
     .Refresh
ENDWITH 
***********************************************************
PROCEDURE saveCompare
PARAMETERS par1
WITH fCompare
     .butRead.Visible=.T.
     .butPrn.Visible=.T.
     .butTot.Visible=.T.
     .butSet.Visible=.T.
     .butRet.Visible=.T.
     .butSave.Visible=.F.
     .butRef.Visible=.F.
    
     .butNew.Enabled=.T.
     .tBoxFioNew.Enabled=.T.   
     .grdOld.Enabled=.T.    
     .grdOld.SetAll('Enabled',.F.,'ColumnMy')
     .grdOld.Columns(.grdOld.ColumnCount).Enabled=.T.
     .grdNew.Enabled=.T.    
     .grdNew.SetAll('Enabled',.F.,'ColumnMy')
     .grdNew.Columns(.grdNew.ColumnCount).Enabled=.T.
     
     .grd1.Enabled=.T.    
     .grd1.SetAll('Enabled',.F.,'ColumnMy')
     .grd1.Columns(.grd1.ColumnCount).Enabled=.T.
     
     .grd2.Enabled=.T.    
     .grd2.SetAll('Enabled',.F.,'ColumnMy')
     .grd2.Columns(.grd2.ColumnCount).Enabled=.T.            
ENDWITH 
IF par1
   SELECT olddatjob
   SET ORDER TO 7
   SEEK curJobOld.nid
   SELECT curFondOld
   mrec=RECNO()
   SCAN ALL
        IF logr
           rep1=ALLTRIM(fpers)
           rep2=ALLTRIM(fname)
           
           SELECT curjobold   
           REPLACE &rep1 WITH curFondOld.say1,&rep2 WITH curFondOld.say2
           SELECT olddatjob   
           REPLACE &rep1 WITH curFondOld.say1,&rep2 WITH curFondOld.say2
           SELECT curfondold    
        ENDIF
   ENDSCAN
   LOCATE FOR ALLTRIM(LOWER(fname))='msf'
   newMsf=say2
   
   *LOCATE FOR ALLTRIM(LOWER(fname))='mslwork'
   GO mrec
   SELECT curJobOld
   REPLACE msf WITH newMsf  
   SELECT olddatjob
   
   
   SELECT datjob
   SET ORDER TO 7
   SEEK dcompare.nidjob
   SELECT curJobNew
   REPLACE mbdpl WITH dcompare.mbdpl,pslwork WITH dcompare.newpsl,mSlWork WITH dcompare.newmsl
   SELECT curFondNew
   SCAN ALL
        IF logr
           rep1=ALLTRIM(fpers)
           rep2=ALLTRIM(fname)
           
           SELECT curjobNew   
           REPLACE &rep1 WITH curFondnew.say1,&rep2 WITH curFondnew.say2
           SELECT datjob   
           REPLACE &rep1 WITH curFondNew.say1,&rep2 WITH curFondNew.say2
           SELECT curfondnew
        ENDIF
   ENDSCAN
   
   SELECT datjob   
   REPLACE mbdpl WITH dcompare.mbdpl,pslwork WITH dcompare.newpsl,mSlWork WITH dcompare.newmsl 
   SET ORDER TO 6
ELSE
   SELECT dcompare
   IF logAp
      DELETE 
   ELSE
      GATHER FROM dcom_old  
   ENDIF    
ENDIF
DO compareoldscr  
DO comparescr
fCompare.grdNew.Columns(fCompare.grdNew.ColumnCount).SetFocus 
***********************************************************
PROCEDURE formPrnCompare
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl 
     .Caption='Печать'     
     DO adSetupPrnToForm WITH 20,20,350,.F.,.T.      
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnCompare WITH .T.' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'Просмотр','DO prnCompare WITH .F.'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','fSupl.Release','Выход из печати' 
     .Height=.Shape91.Height+.cont1.Height+60
     .Width=.Shape91.Width+40    
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***********************************************************
PROCEDURE prnCompare
PARAMETERS par1
SELECT dcompare
SELECT * FROM comprn INTO CURSOR curprn ORDER BY num READWRITE
SELECT curPrn
SCAN ALL
     DO CASE     
        CASE ALLTRIM(LOWER(fname))='mslwork'        
             REPLACE say1 WITH curjobold.pslwork,say2 WITH curjobold.mslwork
             REPLACE say3 WITH dcompare.newpsl,say4 WITH dcompare.newmsl
        CASE num=20 && базовая ставка
             SELECT curPrn
             REPLACE say2 WITH datset.bstold
             REPLACE say4 WITH dcompare.bst
        CASE num=25 && тарифный разряд
             SELECT curPrn
             REPLACE say1 WITH IIF(curjobold.kf<19,curjobold.kf,0)
             REPLACE say3 WITH IIF(dcompare.kf<19,dcompare.kf,0)
        CASE num=26 && тарифный разряд
             SELECT curPrn
             REPLACE say1 WITH curjobold.namekf
             REPLACE say3 WITH dcompare.namekf     
        CASE num=27 && повышающий коэффициент
             SELECT curPrn
             REPLACE say1 WITH curjobold.pkf
             REPLACE say3 WITH dcompare.pkf            
        CASE num=85 
             SELECT curPrn
             REPLACE say4 WITH dcompare.difsum
        CASE num=70
             SELECT curFondOld
             LOCATE FOR num=71
             SELECT curprn
             REPLACE say1 WITH curfondold.say1,say2 WITH curfondold.say2        
        CASE num=71                 
             SELECT curFondNew
             LOCATE FOR num=71
             SELECT curprn
             REPLACE say3 WITH curfondnew.say1,say4 WITH curfondnew.say2    
        CASE num=72                 
             SELECT curFondNew
             LOCATE FOR num=72
             SELECT curprn
             REPLACE say3 WITH IIF(dcompare.mtokl#0,ROUND(dcompare.mbdpl/curjobnew.mtokl*100,0),0),say4 WITH curfondnew.say2       
              
        CASE num=80 
             REPLACE say2 WITH curjobold.msf
             REPLACE say4 WITH dcompare.msf+dcompare.newmsl               
        OTHERWISE 
             SELECT curFondOld
             LOCATE FOR num=curprn.num
             SELECT curprn
             REPLACE say1 WITH curfondold.say1,say2 WITH curfondold.say2
             SELECT curFondNew
             LOCATE FOR num=curprn.num
             SELECT curprn
             REPLACE say3 WITH curfondnew.say1,say4 WITH curfondnew.say2
     ENDCASE
     SELECT curprn   
ENDSCAN
DELETE FOR num<70.AND.say1=0.AND.say2=0.AND.say3=0.AND.say4=0
GO TOP
DO procForPrintAndPreview WITH 'reptbl','Сравнительная таблица',par1,'compareToExcel'
***********************************************************
PROCEDURE compareToExcel
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
     .Columns(1).ColumnWidth=50
     .Columns(2).ColumnWidth=10
     .Columns(3).ColumnWidth=10
     .Columns(4).ColumnWidth=10
     .Columns(5).ColumnWidth=10
            
   
     .Range(.Cells(1,1),.Cells(1,5)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=IIF(SEEK(dcompare.nidpeop,'people',4),ALLTRIM(people.fio),'')
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
       
         
     .Range(.Cells(2,1),.Cells(4,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=IIF(SEEK(dcompare.kp,'sprpodr',1),ALLTRIM(sprpodr.name),'')+' '+IIF(SEEK(dcompare.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' '+LTRIM(STR(dcompare.kse,4,2))
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
                     
     .Range(.Cells(2,2),.Cells(3,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Действующие условия оплаты труда'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH 
      
     .Range(.Cells(2,4),.Cells(3,5)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Новые условия оплаты труда с 01 июля 2021'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH    
      .cells(4,2).Value='размер'
      .cells(4,3).Value='сумма'
      .cells(4,4).Value='размер'
      .cells(4,5).Value='сумма'
   
      rowcx=5  
      SELECT curPrn
      STORE 0 TO max_rec,one_pers,pers_ch
      COUNT TO max_rec
      GO TOP
   
      SCAN ALL          
           .Cells(rowcx,1).Value=ALLTRIM(rec)                                    
           .Cells(rowcx,2).Value=IIF(say1#0,say1,'')                                             
           .Cells(rowcx,3).Value=IIF(say2#0,say2,'')                                             
           .Cells(rowcx,4).Value=IIF(say3#0,say3,'')                                             
           .Cells(rowcx,5).Value=IIF(say4#0,say4,'')                                             
           rowcx=rowcx+1
                   
      ENDSCAN                                 
      .Range(.Cells(2,1),.Cells(rowcx-1,5)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
     
      
          
      .Range(.Cells(1,1),.Cells(rowcx-1,8)).Select
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

objExcel.Visible=.T.
************************************************************
PROCEDURE validGrd1
IF curFondOld.logr
   SELECT curFondOld
   mrec=RECNO() 
   repfrm=ALLTRIM(curFondOld.formula) 
   IF !EMPTY(repfrm)
      IF say1<1
         REPLACE say1 WITH 0        
      ENDIF
      REPLACE say2 WITH &repfrm     
      fcompare.Refresh
      SUM say2 TO say2tot FOR !EMPTY(sum_f)
      GO BOTTOM
      REPLACE say2 WITH say2tot    
      GO mrec          
      SELECT dcompare
      newbdpl=0
      IF curjobnew.tr=1
         newBdpl=varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts
         *newBdpl=varNmzp*curjobold.kse-curjobold.mtokl-curjobold.mStSum-curjobold.mKonts
         newBdpl=IIF(newBdpl>0,newBdpl,0)
      ENDIF    
      REPLACE mbdpl WITH newBdpl
      sumItNew=&formulanew
      REPLACE dcompare.msf WITH sumItNew,difsum WITH msf-say2tot  
      DO CASE
         CASE mbdpl>0.AND.difSum>0
              REPLACE newpsl WITH 0,newmsl WITH 0                       
         CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0         
              reppsl=0
              DO CASE
                 CASE cround=1 &&без округления
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                 CASE cround=2 &&по правилам 
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                 CASE cround=3 &&в большую сторону
                      reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
              ENDCASE
              IF curjobold.pslwork=0
                 reppsl=0
              ENDIF
              REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         
         
             * REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         CASE difsum<0.AND.dcompare.mtokl#0 
              reppsl=0
              DO CASE
                 CASE cround=1 &&без округления
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                 CASE cround=2 &&по правилам 
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                 CASE cround=3 &&в большую сторону
                      reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
              ENDCASE
              REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
     *         REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100            
      ENDCASE       
      DO comparescr
      SELECT curfondold
   ENDIF
ENDIF
***********************************************************
PROCEDURE validGrd12
IF curFondOld.logr.AND.ALLTRIM(LOWER(curfondOld.fname))='mslwork'
   SELECT curFondOld
   mrec=RECNO() 
   IF say2#0.AND.curjobold.mtokl#0
      REPLACE say1 WITH say2/curjobold.mtokl*100   
   ENDIF
   *   fcompare.Refresh
   SUM say2 TO say2tot FOR !EMPTY(sum_f)
   GO BOTTOM
   REPLACE say2 WITH say2tot    
   GO mrec          
   SELECT dcompare
   newbdpl=0
   IF curjobnew.tr=1
      newBdpl=varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts
      *newBdpl=varNmzp*curjobold.kse-curjobold.mtokl-curjobold.mStSum-curjobold.mKonts
      newBdpl=IIF(newBdpl>0,newBdpl,0)
   ENDIF    
   REPLACE mbdpl WITH newBdpl
   sumItNew=&formulanew
   REPLACE dcompare.msf WITH sumItNew,difsum WITH msf-say2tot  
   DO CASE
      CASE mbdpl>0.AND.difSum>0          
           REPLACE newpsl WITH 0,newmsl WITH 0                       
      CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0
           reppsl=0
           DO CASE
              CASE cround=1 &&без округления
                   reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
              CASE cround=2 &&по правилам 
                   reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
              CASE cround=3 &&в большую сторону
                   reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
           ENDCASE
           IF curjobold.pslwork=0
              reppsl=0
           ENDIF 
           REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
           *REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
      CASE difsum<0.AND.dcompare.mtokl#0 
           reppsl=0
           DO CASE
              CASE cround=1 &&без округления
                   reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
              CASE cround=2 &&по правилам 
                   reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
              CASE cround=3 &&в большую сторону
                   reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
           ENDCASE
           REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
           *REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100            
   ENDCASE       
   DO comparescr
   SELECT curfondold
  
ENDIF
***********************************************************
PROCEDURE validGrd2
IF curFondNew.logr
   SELECT curFondNew
   mrec=RECNO() 
   repfrm=ALLTRIM(curFondNew.formula) 
   IF !EMPTY(repfrm)
      rep1=ALLTRIM(fpers)
      rep2=ALLTRIM(fname)
      REPLACE say2 WITH &repfrm
      SUM say2 TO say2tot FOR !EMPTY(sum_f)
      LOCATE FOR num=80
      REPLACE say2 WITH say2tot    
      GO mrec        
      SELECT curfondold
      mrecold=RECNO()
      GO BOTTOM 
      msfold=say2
      GO mrecold  
      SELECT dcompare
      REPLACE &rep1 WITH curfondnew.say1,&rep2 WITH curfondnew.say2,msf WITH say2tot,difsum WITH msf-msfold
      
      DO CASE
         CASE mbdpl>0.AND.difSum>0
              REPLACE newpsl WITH 0,newmsl WITH 0                       
         CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0         
              reppsl=0
              DO CASE
                 CASE cround=1 &&без округления
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                 CASE cround=2 &&по правилам 
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                 CASE cround=3 &&в большую сторону
                      reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
              ENDCASE
              IF curjobold.pslwork=0
                 reppsl=0
              ENDIF 
              REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         
         
             * REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         CASE difsum<0.AND.dcompare.mtokl#0 
              reppsl=0
              DO CASE
                 CASE cround=1 &&без округления
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                 CASE cround=2 &&по правилам 
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                 CASE cround=3 &&в большую сторону
                      reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
              ENDCASE
              REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
     *         REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100            
      ENDCASE       
      DO comparescr
      SELECT curfondnew
      GO mrec
      
      
      SELECT curfondnew
      
       *fcompare.Refresh    
   ENDIF
ENDIF
********************************************




PROCEDURE validGrd2222
IF curFondNew.logr
   SELECT curFondNew
   mrec=RECNO() 
   repfrm=ALLTRIM(curFondNew.formula) 
   IF !EMPTY(repfrm)
      IF say1<1
         REPLACE say1 WITH 0        
      ENDIF
      REPLACE say2 WITH &repfrm     
      *fcompare.Refresh
      SUM say2 TO say2tot FOR !EMPTY(sum_f)
      GO BOTTOM
      REPLACE say2 WITH say2tot    
      GO mrec          
      SELECT dcompare
      newbdpl=0
      IF curjobnew.tr=1
         *newBdpl=varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts
         newBdpl=varNmzp*curjobold.kse-curjobold.mtokl-curjobold.mStSum-curjobold.mKonts
         newBdpl=IIF(newBdpl>0,newBdpl,0)
      ENDIF    
      REPLACE mbdpl WITH newBdpl
      sumItNew=&formulanew
      REPLACE dcompare.msf WITH sumItNew,difsum WITH msf-say2tot  
      DO CASE
         CASE mbdpl>0.AND.difSum>0
              REPLACE newpsl WITH 0,newmsl WITH 0                       
         CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0         
              reppsl=0
              DO CASE
                 CASE cround=1 &&без округления
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                 CASE cround=2 &&по правилам 
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                 CASE cround=3 &&в большую сторону
                      reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
              ENDCASE
              REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         
              IF curjobold.pslwork=0
                 reppsl=0
              ENDIF
             * REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         CASE difsum<0.AND.dcompare.mtokl#0 
              reppsl=0
              DO CASE
                 CASE cround=1 &&без округления
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                 CASE cround=2 &&по правилам 
                      reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                 CASE cround=3 &&в большую сторону
                      reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
              ENDCASE
              REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
     *         REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100            
      ENDCASE       
      DO comparescr
      SELECT curfondnew
      GO mrec
   ENDIF
ENDIF



***********************************************************
PROCEDURE oldchange
IF curFondOld.logr
   fcompare.grd1.column2.tBox3.ReadOnly=.F. 
   fcompare.grd1.column3.tBox3.ReadOnly=.F. 
ELSE    
   fcompare.grd1.column2.tBox3.ReadOnly=.T.
   fcompare.grd1.column3.tBox3.ReadOnly=.T. 
ENDIF
***********************************************************
PROCEDURE newchange
IF curFondNew.logr
   fcompare.grd2.column1.txtBox1.ReadOnly=.F. 
ELSE    
   fcompare.grd2.column1.txtBox1.ReadOnly=.T.
ENDIF
***********************************************************************************
PROCEDURE countbdpl
SELECT dcompare
newbdpl=0
IF curjobnew.tr=1
   newBdpl=varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts
   *newBdpl=varNmzp*curjobold.kse-curjobold.mtokl-curjobold.mStSum-curjobold.mKonts
   newBdpl=IIF(newBdpl>0,newBdpl,0)
ENDIF    
REPLACE mbdpl WITH newBdpl
sumItNew=&formulanew
REPLACE dcompare.msf WITH sumItNew,difsum WITH msf-curJobOld.msf
DO CASE
   CASE mbdpl>0.AND.difSum>0
        REPLACE newpsl WITH 0,newmsl WITH 0                       
   CASE mbdpl=0.AND.difsum>0.AND.dcompare.mtokl#0
        reppsl=0
        DO CASE
           CASE cround=1 &&без округления
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
           CASE cround=2 &&по правилам 
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
           CASE cround=3 &&в большую сторону
                reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
        ENDCASE
        IF curjobold.pslwork=0
           reppsl=0
        ENDIF
        REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
        *REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
   CASE difsum<0.AND.dcompare.mtokl#0 
        reppsl=0
        DO CASE
           CASE cround=1 &&без округления
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
           CASE cround=2 &&по правилам 
                reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
           CASE cround=3 &&в большую сторону
                reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
        ENDCASE
        REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
   *     REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
ENDCASE  
*********************************************************
PROCEDURE fltcompare
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Фильтр'
     .Icon='kone.ico'
     .Width=500
     .Height=600 
     DO addListBoxMy WITH 'fSupl',1,0,0,500,.Width  
     WITH .listBox1
          .ColumnCount=2
          .ColumnWidths='40,460' 
          .ColumnLines=.F.
          .RowSource='curFltPodr.otm,name'         
          .RowSourceType=2
          .ControlSource=''
          .Visible=.T. 
          .procForClick='DO clickListPodr'       
          .procForKeyPress='DO KeyPressListPodr'                
          .Enabled=.T.
     ENDWITH  
     *-----------------------------Кнопка принять---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont11',.listBox1.Left+(.listBox1.Width-(RetTxtWidth('wпринять')*2)-15)/2,.listBox1.Top+.listBox1.Height+10,RetTxtWidth('wпринятьw'),dHeight+5,'принять','DO returnToCompare WITH .T.'
     *---------------------------------Кнопка сброс-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont12',.cont11.Left+.cont11.Width+15,.Cont11.Top,.Cont11.Width,dHeight+5,'сброс','DO returnToCompare WITH .F.'
     .Height=.listBox1.Height+.cont11.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE clickListPodr
SELECT curFltPodr
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListPodr
DO CASE
   CASE LASTKEY()=27      
   CASE LASTKEY()=13
        Do clickListPodr
ENDCASE 
************************************************************************************************************************
PROCEDURE returnToCompare
PARAMETERS par1
IF !par1
   SELECT curfltpodr
   REPLACE fl WITH .F.,otm WITH '' ALL 
   fltpodr=''
   SELECT datjob
   SET FILTER TO 
   fltsay='фильтр - организация (изменить - двойной щелчок мыши)'
   fCompare.txtNew.ContLabel.Caption=fltSay
ELSE
   fltpodr=''
   SELECT curFltPodr
   SCAN ALL
        IF fl
           fltpodr=fltpodr+','+LTRIM(STR(kod))       
        ENDIF      
   ENDSCAN    
   IF !EMPTY(fltpodr)
      fltpodr=fltpodr+','
      SELECT datjob
      SET FILTER TO ','+LTRIM(STR(kp))+','$fltpodr  
      fltsay='фильтр - подразделение (изменить - двойной щелчок мыши)'   
      fCompare.txtNew.ContLabel.Caption=fltSay
   ELSE   
      fltsay='фильтр - организация (изменить - двойной щелчок мыши)'
      fCompare.txtNew.ContLabel.Caption=fltSay
      SELECT datjob
      SET FILTER TO
   ENDIF
ENDIF
fSupl.Release
***********************************************************************
PROCEDURE setupCompare
fSupl=CREATEOBJECT('FORMSUPL')

DIMENSION dim_cround(3)
STORE 0 TO dim_cround
dim_cround(cround)=1
WITH fSupl
     .Caption='настройки'
     .icon='main.ico'
     .Width=400
     DO addshape WITH 'fSupl',1,20,20,150,380,8 
     DO adlabMy WITH 'fSupl',1,' Округление ',.Shape1.Top-10,.Shape1.Left+10,300,0,.T.,1
     DO addOptionButton WITH 'fSupl',1,'без округления',.Shape1.Top+20,.Shape1.Left+20,'dim_cround(1)',0,'DO procSelectRound WITH 1',.T. 
     DO addOptionButton WITH 'fSupl',2,'по правилам',.Option1.Top+.Option1.Height+10,.Option1.Left,'dim_cround(2)',0,'DO procSelectRound WITH 2',.T.
     DO addOptionButton WITH 'fSupl',3,'в большую сторону',.Option2.Top+.Option2.Height+10,.Option2.Left,'dim_cround(3)',0,'DO procSelectRound WITH 3',.T.
     .Shape1.Height=.option1.Height*3+50
     .Shape1.Width=.Option3.Width+40
     .Width=.Shape1.Width+40
     DO addButtonOne WITH 'fSupl','butRet',(.Width-RetTxtWidth('wвозврат'))/2,.Shape1.Top+.Shape1.Height+20,'возврат','','fSupl.Release',39,RetTxtWidth('wвозвратw'),'возврат'  
     .Height=.Shape1.Height+.butRet.Height+60
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************
PROCEDURE procSelectRound
PARAMETERS par1
STORE 0 TO dim_cround
dim_cround(par1)=1
fSupl.Refresh
REPLACE datshtat.vround WITH par1
cround=par1
**************************************************************
PROCEDURE formTotCompare
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='перерасчёт'
     .Icon='money.ico'
     .Width=350
     DO adLabMy WITH 'fSupl',1,'выполнить перерасчёт ?',20,0,.Width,2,.F.,0
     
      DO addShape WITH 'fSupl',2,10,10,30,.Width-20,8
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
     DO addShape WITH 'fSupl',3,.Shape2.Left,.Shape2.Top,.Shape2.Height,50,8
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.               
      DO adLabMy WITH 'fSupl',25,'100%',.Shape2.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab25.Visible=.F.
          
     DO addButtonOne WITH 'fSupl','butYes',(.Width-RetTxtWidth('wвозвратw')*2-20)/2,.lab1.Top+.lab1.Height+10,'да','','DO countTotCompare',39,RetTxtWidth('wвозвратw'),'расчёт'  
     DO addButtonOne WITH 'fSupl','butNo',.butYes.Left+.butYes.Width+10,.butYes.Top,'вовзрат','','fSupl.Release',39,.butYes.Width,'возврат' 
     DO addButtonOne WITH 'fSupl','butRet',(.Width-RetTxtWidth('wвозвратw'))/2,.butYes.Top,'возврат','','fSupl.Release',39,RetTxtWidth('wвозвратw'),'возврат'  
     .butRet.Visible=.F.
     .Height=.lab1.Height+.butYes.Height+50
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************
PROCEDURE countTotCompare
WITH fSupl
     .Shape2.Visible=.T.
     .Shape3.Visible=.T.
     .lab25.Visible=.T.
     .butYes.Visible=.F.
     .butNo.Visible=.F.
     .lab1.Visible=.F.
ENDWITH
SELECT datjob
SET ORDER TO 7
SELECT dcompare
oldcomrec=RECNO()
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec
SCAN ALL
     IF SEEK(dcompare.nidold,'olddatjob',7)
        sitold=0
        newbdpl=0
        SELECT curfondold
        SCAN ALL
             rep1=fname
             IF !EMPTY(sum_f)
                rep1='olddatjob.'+ALLTRIM(sum_f)
                sitold=sitold+&rep1
             ENDIF
        ENDSCAN
        newBdpl=IIF(dcompare.tr=1,varNmzp*dcompare.kse-dcompare.mtokl-dcompare.mStSum-dcompare.mKonts,0)   
        newBdpl=IIF(newBdpl<=0,0,newBdpl)
        SELECT datjob
        SEEK dcompare.nidjob
        REPLACE mbdpl WITH newBdpl             
        SELECT curfondnew
        SCAN ALL             
             SELECT curfondnew
             SCAN ALL                 
                SELECT curFondNew            
                IF num<90
                   rep1=IIF(!EMPTY(fpers),ALLTRIM(curFondNew.fpers),'')
                   rep2=IIF(!EMPTY(fname),ALLTRIM(curFondNew.fname),'') 
                   rep11=IIF(!EMPTY(fpers),'datjob.'+ALLTRIM(curFondNew.fpers),'')
                   rep12=IIF(!EMPTY(fname),'datjob.'+ALLTRIM(curFondNew.fname),'') 
                   SELECT dcompare
                   IF !EMPTY(rep1)
                      REPLACE &rep1 WITH IIF(!EMPTY(rep11),EVALUATE(rep11),0)
                   ENDIF
                   IF !EMPTY(rep2)
                      REPLACE &rep2 WITH IIF(!EMPTY(rep12),EVALUATE(rep12),0)
                   ENDIF
                ENDIF
                SELECT curfondnew                          
             ENDSCAN
        ENDSCAN
        SELECT dcompare
        sumItNew=&formulanew                        
        REPLACE msf WITH sumItNew,difsum WITH sumItNew-sitold
        DO CASE
           CASE mbdpl>0.AND.difSum>0
                REPLACE newpsl WITH 0,newmsl WITH 0                       
            CASE mbdpl=0.AND.difsum>0.AND.mtokl#0         
                 reppsl=0
                 DO CASE
                    CASE cround=1 &&без округления
                         reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)
                    CASE cround=2 &&по правилам 
                         reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                    CASE cround=3 &&в большую сторону
                         reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
                 ENDCASE
                 IF olddatjob.pslwork=0
                    reppsl=0
                 ENDIF
                 REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
         
         
             * REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
            CASE difsum<0.AND.mtokl#0          
                 reppsl=0
                 DO CASE
                    CASE cround=1 &&без округления
                         reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,2)   
                    CASE cround=2 &&по правилам 
                         reppsl=ROUND(ABS(difSum)/dcompare.mtokl*100,0)   
                    CASE cround=3 &&в большую сторону
                         reppsl=INT(ABS(difSum)/dcompare.mtokl*100)+1   
                 ENDCASE
                 REPLACE dcompare.newpsl WITH reppsl,dcompare.newmsl WITH dcompare.mtokl*dcompare.newpSl/100  
     *         REPLACE dcompare.newpsl WITH ABS(difSum)/dcompare.mtokl*100,newmsl WITH dcompare.mtokl*dcompare.newpSl/100            
         ENDCASE
         REPLACE datjob.pslwork WITH dcompare.newpsl,datjob.mslwork WITH dcompare.newmsl            
     ENDIF
     SELECT dcompare
     one_pers=one_pers+1
     pers_ch=one_pers/max_rec*100
     fSupl.shape3.Visible=.T.
     fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
     fSupl.Shape3.Width=fSupl.shape2.Width/100*pers_ch 
ENDSCAN
GO oldcomrec
SELECT datjob
SET ORDER TO 6
fSupl.butRet.Visible=.T.