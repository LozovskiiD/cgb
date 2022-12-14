fOrdStart=CREATEOBJECT('FORMSUPL')
IF !USED('boss')
   USE boss IN 0
ENDIF

=AFIELDS(arOrder,'peoporder')
CREATE CURSOR curPeopOrder FROM ARRAY arOrder
SELECT curPeopOrder
INDEX ON npp TAG T1
SELECT datOrder 
SET ORDER TO 2
SET FILTER TO typeord=1
GO TOP 
DIMENSION dim_type(4)
STORE 0 TO dim_type
dim_type(1)=1
newKodOrder=0   && ??? ???????  
typeOrdNew=1  && ??? ??????? 
numOrdNew=0   && ????? ???????
dateOrdNew=CTOD('  .  .    ')
strOrdNew='?' && ??????????? ???????
logNew=.F.
logNewRec=.F.
logRead=.F.
newSlink=''
str_ini=''
objFocus=''  && ?????? ??? ???????? ?????? ????? ?????? ?? listbox
parPadejDol=1
parPadejFio=1
newLogApp=.T.
logDatJob=.F.
SELECT datOrder
GO TOP
WITH fOrdStart
     .Caption='???????'  
     DO addshape WITH 'fOrdStart',1,10,20,150,450,8
     shapeWidth=390
     .AddObject('grdOrd','gridMynew') 
     DO addOptionButton WITH 'fOrdStart',1,'????????????????',.Shape1.Top+20,.Shape1.Left+20,'dim_type(1)',0,"DO procValTypeOrd WITH 1",.T.     
     DO addOptionButton WITH 'fOrdStart',2,'????????',.Option1.Top,.Option1.Left+.Option1.Width+10,'dim_type(2)',0,"DO procValTypeOrd WITH 2",.T.     
     DO addOptionButton WITH 'fOrdStart',3,'??????',.Option1.Top,.Option2.Left+.Option2.Width+10,'dim_type(3)',0,"DO procValTypeOrd WITH 3",.T.
     DO addOptionButton WITH 'fOrdStart',4,'????????????',.Option1.Top+.Option1.Height+10,.Option3.Left+.Option3.Width+10,'dim_type(4)',0,"DO procValTypeOrd WITH 4",.T.
     DO adtBoxAsCont WITH 'fOrdStart','contNum',10,.Option4.Top+.Option4.Height+10,RetTxtWidth('W????? ???????W'),dHeight,'????? ???????',2,1  
     DO adtBoxAsCont WITH 'fOrdStart','contDate',.contNum.Left+.contNum.Width-1,.contNum.Top,RetTxtWidth('W???? ???????W'),dHeight,'???? ???????',2,1  
     DO adTboxNew WITH 'fOrdStart','tBox1',.contNum.Top+.contNum.Height-1,.contNum.Left,.contNum.Width,dHeight,'numOrdNew','Z',.T.,2 
    
     DO adTboxNew WITH 'fOrdStart','tBox3',.tBox1.Top,.contDate.Left,.contDate.Width,dHeight,'dateOrdNew',.F.,.T.,2 
     
     .Shape1.Width=.Option1.Width+.Option2.Width+.Option3.Width+60
     .Shape1.Height=.option1.Height*2+.tBox1.Height*2+60
     .Option4.Left=.Shape1.Left+(.Shape1.Width-.Option4.Width)/2
     
     .contNum.Left=.Shape1.Left+(.Shape1.Width-.contNum.Width-.contDate.Width+1)/2
     .contDate.Left=.contNum.Left+.contNum.Width-1
     .tBox1.Left=.contNum.Left
    * .tBox2.Left=.tBox1.Left+.tBox1.Width-1
     .tBox3.Left=.contDate.Left
                 
     WITH .grdOrd
          .Top=.Parent.Shape1.Top+.Parent.Shape1.Height+10     
          .Left=.Parent.Shape1.Left
          .Width=.Parent.Shape1.Width
          .Height=300
          .RecordSourceType=1
          .scrollBars=2   
          .ColumnCount=0         
          .colNesInf=2                         
          DO addColumnToGrid WITH 'fOrdStart.grdOrd',7       
          .RecordSource='datOrder'                                         
          .Column1.ControlSource='datOrder.numOrd'
          .Column2.ControlSource='datOrder.strOrd'
          .Column3.ControlSource='datOrder.dateOrd'
          .Column4.ControlSource='datorder.cordname'   
          .Column5.ControlSource='datorder.nkvo'  
          .Column6.ControlSource='datorder.lotm'  
          .Column1.Width=RetTxtWidth('999999')      
          .Column2.Width=RetTxtWidth('w?w')       
          .Column3.Width=RetTxtWidth('w99/99/9999')
          .Column5.Width=RetTxtWidth('999w') 
          .Column6.Width=RetTxtWidth('w!w') 
          .Columns(.ColumnCount).Width=0    
          .Column4.Width=.Width-.column1.Width-.Column2.Width-.Column3.Width-.Column5.Width-.Column6.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Column1.Header1.Caption='?'
          .Column2.Header1.Caption='???'
          .Column3.Header1.Caption='????'
          .Column4.Header1.Caption='????????????'
          .Column5.Header1.Caption='?'
          .Column6.Header1.Caption='!'
          .Column1.Alignment=2  
          .Column2.Alignment=2  
          .Column3.Alignment=0
          .Column4.Alignment=0
          .Column4.Alignment=5
          .Column4.Alignment=6
          .Column1.Format='Z'
          .Column2.Format='Z'
          .Column5.Format='Z'
          .Column6.AddObject('checkColumn6','checkContainer')
          .Column6.checkColumn6.AddObject('checkMy','checkBox')
          .Column6.CheckColumn6.checkMy.Visible=.T.
          .Column6.CheckColumn6.checkMy.Caption=''
          .Column6.CheckColumn6.checkMy.Left=5
          .Column6.CheckColumn6.checkMy.BackStyle=0
          .Column6.CheckColumn6.checkMy.ControlSource='datorder.lotm'                                                                                                  
          .column6.CurrentControl='checkColumn6'
          .Column6.Sparse=.F.                    
          
          .procAfterRowColChange='DO changeRowOrder'                                                              
          .SetAll('Enabled',.F.,'ColumnMy') 
          .Columns(.ColumnCount).Enabled=.T.          
     ENDWITH 
     DO gridSizeNew WITH 'fOrdStart','grdOrd','shapeingrid',.T.,.F.
     FOR i=1 TO .grdOrd.columnCount 
         .grdOrd.Columns(i).Backcolor=fOrdStart.BackColor           
         .grdOrd.Columns(i).DynamicBackColor='IIF(RECNO(fOrdStart.grdOrd.RecordSource)#fOrdStart.grdOrd.curRec,fOrdStart.BackColor,dynBackColor)'
         .grdOrd.Columns(i).DynamicForeColor='IIF(RECNO(fOrdStart.grdOrd.RecordSource)#fOrdStart.grdOrd.curRec,dForeColor,dynForeColor)'        
     ENDFOR            
     
     DO addButtonOne WITH 'fOrdStart','butStart',.shape1.Left+(.Shape1.Width-(RetTxtWidth('w??????????w')*3)-20)/2,.grdOrd.Top+.grdOrd.Height+20,'??????????','','DO formforOrder',dHeight+5,RetTxtWidth('w??????????w'),'????? ??????'  
     DO addButtonOne WITH 'fOrdStart','butRead',.butStart.Left+.butStart.Width+10,.butStart.Top,'?????-????','','DO changeNumOrd',.butStart.Height,.butStart.Width,'????????????? ?????-????'  
     DO addButtonOne WITH 'fOrdStart','butRet',.butRead.Left+.butRead.Width+10,.butStart.Top,'???????','','fOrdStart.Release',.butStart.Height,.butStart.Width,'???????'      


     DO addButtonOne WITH 'fOrdStart','butSave',.shape1.Left+(.Shape1.Width-(RetTxtWidth('w????????w')*2)-10)/2,.butStart.Top,'????????','','DO saveNumOrder WITH .T.',.butStart.Height,RetTxtWidth('w????????w'),'????????'  
     DO addButtonOne WITH 'fOrdStart','butReatRead',.butSave.Left+.butSave.Width+10,.butSave.Top,'?????','','DO saveNumOrder WITH .F.',.butSave.Height,.butSave.Width,'?????'       
     .butSave.Visible=.F.
     .butReatRead.Visible=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.grdOrd.Height+.butStart.Height+70  
     
ENDWITH
DO pasteImage WITH 'fOrdStart'
fOrdStart.Show

********************************************************************************************************************************
PROCEDURE changeRowOrder
numOrdNew=datOrder.numOrd
strOrdNew=datOrder.strOrd
STORE 0 TO dim_type
dim_type(datOrder.typeOrd)=1
typeOrdNew=datOrder.typeOrd
dateOrdNew=datOrder.dateOrd
fOrdStart.Refresh
********************************************************************************************************************************
PROCEDURE procValTypeOrd
PARAMETERS par1
STORE 0 TO dim_type
dim_type(par1)=1
SELECT datorder
DO CASE
   CASE dim_type(1)=1
        typeOrdNew=1
        strOrdNew='?'        
        SET FILTER TO typeord=1
   CASE dim_type(2)=1
        typeOrdNew=2
        strOrdNew='?' 
        SET FILTER TO typeord=2     
   CASE dim_type(3)=1
        typeOrdNew=3
        strOrdNew='?'  
        SET FILTER TO typeord=3
   CASE dim_type(4)=1
        typeOrdNew=4
        strOrdNew='??'         
        SET FILTER TO typeord=4
ENDCASE
GO TOP
fOrdStart.Refresh
fOrdStart.grdOrd.Columns(fOrdStart.grdOrd.ColumnCount).SetFocus
********************************************************************************************************************************
PROCEDURE changenumord
PUBLIC numOrd_cx,dateOrd_cx
numOrd_cx=datorder.numOrd
dateOrd_cx=datorder.dateOrd
WITH fOrdStart
     .SetAll('Visible',.F.,'MyCommandButton')
     .butSave.Visible=.T.
     .butReatRead.Visible=.T.
     .grdOrd.Enabled=.F.
     .tBox1.ControlSource='numOrd_cx'
     .tBox3.ControlSource='dateOrd_cx'
     .tBox1.SetFocus
ENDWITH
********************************************************************************************************************************
PROCEDURE saveNumOrder
PARAMETERS par1
IF par1
   numOrdNew=numOrd_cx
   dateOrdNew=dateOrd_cx
   SELECT datOrder
   REPLACE numord WITH numOrd_cx,dateOrd WITH dateOrd_cx
   SELECT peoporder
   REPLACE dord WITH datOrder.dateOrd,nord WITH LTRIM(STR(datorder.numord)+'-'+ALLTRIM(datorder.strord)) FOR kord=datorder.kod
   SELECT datjob
   REPLACE dordin WITH datOrder.dateOrd,nordin WITH LTRIM(STR(datorder.numord)+'-'+ALLTRIM(datorder.strord)) FOR kordin=datorder.kod
   REPLACE dordout WITH datOrder.dateOrd,nordout WITH LTRIM(STR(datorder.numord)+'-'+ALLTRIM(datorder.strord)) FOR kordout=datorder.kod
   SELECT datotp
   REPLACE dord WITH datOrder.dateOrd,nord WITH LTRIM(STR(datorder.numord)+'-'+ALLTRIM(datorder.strord)),osnov WITH '??.? '+LTRIM(STR(datOrder.numOrd))+'-'+ALLTRIM(datOrder.strOrd)+' ?? '+DTOC(datOrder.dateOrd) FOR kord=datorder.kod
ELSE 
ENDIF 
WITH fOrdStart
     .SetAll('Visible',.T.,'MyCommandButton')
     .butSave.Visible=.F.
     .butReatRead.Visible=.F.
     .grdOrd.Enabled=.T.
     .grdOrd.SetAll('Enabled',.F.,'ColumnMy')
     .grdOrd.Columns(.grdOrd.ColumnCount).Enabled=.T.  
     .tBox1.ControlSource='numOrdNew'
     .tBox3.ControlSource='dateOrdNew'
ENDWITH
********************************************************************************************************************************
PROCEDURE formForOrder
IF numOrdNew=0
   RETURN
ENDIF
IF USED('curdolpodr')
   SELECT curdolpodr
   USE
ENDIF
IF USED('cursprorder')
   SELECT cursprorder
   USE
ENDIF
IF !USED('txtorder')
   USE txtorder IN 0
ENDIF
IF !USED('fete')
   USE fete IN 0 ORDER 1 
ENDIF
CREATE CURSOR curSupOrder (kod N(1),name C(20))
DO CASE
   CASE typeOrdNew=1 
   CASE typeOrdNew=2 
        SELECT curSupOrder
        APPEND BLANK
        REPLACE kod WITH 1,name WITH '???????'
        APPEND BLANK
        REPLACE kod WITH 2,name WITH '?????????'
        APPEND BLANK
        REPLACE kod WITH 3,name WITH '????????'        
        APPEND BLANK
        REPLACE kod WITH 4,name WITH '???????'
        APPEND BLANK
        REPLACE kod WITH 5,name WITH '??????????'
                
   CASE typeOrdNew=3 
   CASE typeOrdNew=4 
ENDCASE
*SELECT people
*oldpeoprec=RECNO()
readNidPeop=people.nid

SELECT txtOrder
REPLACE txtPrn WITH ''
SELECT * FROM rasp INTO CURSOR curDolPodr READWRITE 
ALTER TABLE curDolPodr ADD COLUMN strVac C(6)
ALTER TABLE curDolPodr ADD COLUMN name C(100)

SELECT curDolPodr
REPLACE name WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL 

INDEX ON nd TAG T1
=AFIELDS(arOrdJob,'datJob')
CREATE CURSOR  curOrdJob FROM ARRAY arOrdJob

=AFIELDS(arPeop,'people')
CREATE CURSOR curSuplPeop FROM ARRAY arPeop
SELECT curSuplPeop
INDEX ON fio TAG t1

fOrdStart.Visible=.F.
SELECT * FROM sprorder WHERE typeord=typeOrdNew INTO CURSOR curSprOrder READWRITE
SELECT curSprOrder
INDEX ON kod TAG T1
SET ORDER TO 1
SELECT datOrder
SET FILTER TO 
SET ORDER TO 2
SEEK STR(YEAR(dateOrdNew),4)+STR(typeOrdNew,1)+STR(numOrdNew,5)
logNew=IIF(FOUND(),.F.,.T.)
IF !FOUND()
   SET ORDER TO 1 
   GO BOTTOM 
   newKodOrder=kod+1 
ELSE
   newKodOrder=kod   
ENDIF 
DIMENSION dim_agree(2)
dim_agree(1)='??????????, ????????:'
dim_agree(2)='???????????, ????????:'

SELECT peoporder
SEEK newKodOrder
DIMENSION dim_nkse(2)
dim_nkse(1)='??????? ???????'
dim_nkse(2)='?????????'
newKodPrik=IIF(logNew,0,peoporder.supOrd)
strPrik=IIF(SEEK(newKodPrik,'curSprOrder',1),curSprOrder.nameOrd,'')
newKodPeop=IIF(logNew,0,peoporder.kodpeop)
newNidPeop=IIF(logNew,0,peoporder.nidpeop)
newFioPeop=IIF(logNew,'',peoporder.fiopeop)
newDateBeg=IIF(logNew,CTOD('  .  .    '),peoporder.dateBeg)
newDateEnd=IIF(logNew,CTOD('  .  .    '),peoporder.dateEnd)
newKp=IIF(logNew,0,peoporder.kp)
newKd=IIF(logNew,0,peoporder.kd)
newKse=IIF(logNew,1.00,peoporder.kse)
newTr=IIF(logNew,1,peoporder.tr)
newNpp=0
STORE .F. TO logNewRec,logDek,logPer,logPers
STORE CTOD('  .  .    ') TO newDateUvol,newPerBeg,newPerEnd,newDKont,repdorder,newPbegOld,newPbegNew
STORE 0 TO newDayOtp,newDDop,newKprich,newPKont,newNidJob,newDayKomp,newSrok,parSave,parTr,newOrdSupl,newSkp,newSKd,nidJobNew,logTxt,newTabZam,vSup,newNid,kodsex,logperevod,oldNidJob,nuvol,oldSovmJob,newVDog,newDnorm,dayCost
STORE '' TO newNPodr,newNDolj,strType,newStrPrich,procForOsnov,newNKont,strSrok,newSpodr,newSDolj,newFioZam,newNSchool,newNKurs,newPlace,newFam,nkse,repnorder
newLotm=datorder.lOtm
newOsnov=SPACE(100)
*APPEND FROM peoporder FOR kord=newKodOrder

DO appendFromPeoporder

frmOrd=CREATEOBJECT('FORMSUPL')
WITH frmOrd
     .Caption='??????' 
     .Width=900
     .Height=700
     .procForClick='DO lostFocusPeop'
     .procExit='DO procExitOrder'
     DO adTboxAsCont WITH 'frmOrd','ordTop',0,0,.Width,dHeight,'',2,1 
     DO adTboxAsCont WITH 'frmOrd','ordNum',0,.ordTop.Top,RetTxtWidth('W?????? ? '),dHeight,'?????? ? ',2,1 
     DO adTboxNew WITH 'frmOrd','tBox1',.ordTop.Top,.ordTop.Left,RetTxtWidth('9999999'),dHeight,'numOrdNew','Z',.F.,0 
     DO adTboxAsCont WITH 'frmOrd','ordType',0,.ordTop.Top,RetTxtWidth('w?w'),dHeight,strOrdNew,2,1      
     DO adTboxAsCont WITH 'frmOrd','ordDate',0,.ordTop.Top,RetTxtWidth('W????'),dHeight,'????',2,1  
     DO adTboxNew WITH 'frmOrd','tBox2',.ordTop.Top,.ordTop.Left,RetTxtWidth('99/99/999999'),dHeight,'dateOrdNew',.F.,.T.,0 

     .tBox1.BackStyle=1
     .tBox2.BackStyle=1          
      
     .ordNum.left=.ordTop.Left+(.ordTop.Width-.ordNum.Width-.ordType.Width-.ordDate.Width-.tBox1.Width-.tBox2.Width)/2-1
     .tBox1.Left=.ordNum.Left+.ordNum.Width-1
     .ordType.Left=.tBox1.Left+.tBox1.Width-1
     .ordDate.Left=.ordType.Left+.ordType.Width-1  
     .tBox2.Left=.ordDate.Left+.ordDate.Width-1
     DO adCheckBox WITH 'frmOrd','checkClose','????????? ??????',.ordTop.Top,.tBox2.Left+.tBox2.Width+5,150,dHeight,'newLotm',1,.F.,'DO validcheckClose' 
     .checkClose.Top=.ordTop.Top+(.ordTop.Height-.checkClose.Height)/2
     .checkClose.Left=.tBox2.Left+.tBox2.Width+(.ordTop.Width-.tBox2.Left-.tBox2.Width-.checkClose.Width)/2
     .checkClose.BackStyle=0
     .AddObject('grdPers','gridMynew')     
     WITH .grdPers
          .Top=.Parent.ordTop.Top+.Parent.ordTop.Height-1     
          .Left=0
          .Width=300
          .Height=.Parent.Height-.Top
          .RecordSourceType=1
          .scrollBars=2   
          .ColumnCount=0         
          .colNesInf=2                         
          DO addColumnToGrid WITH 'frmOrd.grdPers',4       
          .RecordSource='curPeopOrder'                                         
          .Column1.ControlSource='curPeopOrder.npp'
          .Column2.ControlSource='curPeopOrder.kodpeop'
          .Column3.ControlSource='curPeopOrder.fiopeop'                  
          .Column1.Width=RetTxtWidth('999')      
          .Column2.Width=RetTxtWidth('99999')         
          .Columns(.ColumnCount).Width=0    
          .Column3.Width=.Width-.column1.Width-.Column2.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Column1.Header1.Caption='?'
          .Column2.Header1.Caption='???'
          .Column3.Header1.Caption='???'
          .Column1.Alignment=2  
          .Column2.Alignment=2  
          .Column3.Alignment=0  
          .Column1.Format='Z'
          .Column2.Format='Z'
          .procAfterRowColChange='DO changeRowPeopOrder'                                                              
          .SetAll('Enabled',.F.,'ColumnMy') 
          .Columns(.ColumnCount).Enabled=.T.          
     ENDWITH 
     .ordNum.left=.grdPers.Left+.grdPers.Width-1
     .tBox1.Left=.ordNum.Left+.ordNum.Width-1
     .ordType.Left=.tBox1.Left+.tBox1.Width-1
     .ordDate.Left=.ordType.Left+.ordType.Width-1  
     .tBox2.Left=.ordDate.Left+.ordDate.Width-1
     DO gridSizeNew WITH 'frmOrd','grdPers','shapeingrid',.T.,.F.
     FOR i=1 TO .grdPers.columnCount 
         .grdPers.Columns(i).Backcolor=frmOrd.BackColor           
         .grdPers.Columns(i).DynamicBackColor='IIF(RECNO(frmOrd.grdPers.RecordSource)#frmOrd.grdPers.curRec,frmOrd.BackColor,dynBackColor)'
         .grdPers.Columns(i).DynamicForeColor='IIF(RECNO(frmOrd.grdPers.RecordSource)#frmOrd.grdPers.curRec,dForeColor,dynForeColor)'        
     ENDFOR
     IF typeOrdNew=2
        DO adTboxAsCont WITH 'frmOrd','ordSupPrik',.ordNum.Left,.grdPers.Top,RetTxtWidth('W??????? ??? ????????W'),dHeight,'??? ???????',1,1 
        DO addcombomy WITH 'frmOrd',99,.ordSupPrik.Left+.ordSupPrik.Width-1,.ordSupPrik.Top,dHeight,.Width-.grdPers.Width-.ordSupPrik.Width+2,IIF(logNewRec,.T.,.F.),'strSupPrik','ALLTRIM(curSupOrder.name)',6,'','DO procSupOrd',.F.,.T.       
        .comboBox99.Enabled=.F.
        .comboBox99.DisplayCount=20
        .comboBox99.DisabledForeColor=RGB(1,0,0)
        
        DO adTboxAsCont WITH 'frmOrd','ordPrik',.ordNum.Left,.ordSupPrik.Top+.ordSupPrik.Height-1,.ordSupPrik.Width,dHeight,'??????????',1,1 
        DO addcombomy WITH 'frmOrd',1,.ordPrik.Left+.ordPrik.Width-1,.ordPrik.Top,dHeight,.Width-.grdPers.Width-.ordPrik.Width+2,IIF(logNewRec,.T.,.F.),'strPrik','ALLTRIM(curSprOrder.nameOrd)',6,'','DO procPrikOrd',.F.,.T.       
        .comboBox1.Enabled=.F.
        .comboBox1.DisplayCount=20
        .comboBox1.DisabledForeColor=RGB(1,0,0)
     ELSE  
        DO adTboxAsCont WITH 'frmOrd','ordPrik',.ordNum.Left,.grdPers.Top,RetTxtWidth('W??????? ??? ????????W'),dHeight,'??? ???????',1,1 
        DO addcombomy WITH 'frmOrd',1,.ordPrik.Left+.ordPrik.Width-1,.ordPrik.Top,dHeight,.Width-.grdPers.Width-.ordPrik.Width+2,IIF(logNewRec,.T.,.F.),'strPrik','ALLTRIM(curSprOrder.nameOrd)',6,'','DO procPrikOrd',.F.,.T.       
        .comboBox1.Enabled=.F.
        .comboBox1.DisplayCount=20
        .comboBox1.DisabledForeColor=RGB(1,0,0)       
     ENDIF 
     
     
     DO addButtonOne WITH 'frmOrd','butNew',.ordPrik.Left,.Height-59,'?????','','DO newRecInOrder WITH .T.',39,RetTxtWidth('?????????w'),'????? ??????'  
     DO addButtonOne WITH 'frmOrd','butRead',.ordPrik.Left,.butNew.Top,'????????','','DO newRecInOrder WITH .F.',39,.butNew.Width,'????????'  
     DO addButtonOne WITH 'frmOrd','butDel',.ordPrik.Left,.butNew.Top,'????????','','DO delFromOrder',39,.butNew.Width,'????????'      
     DO addButtonOne WITH 'frmOrd','butPrn',.ordPrik.Left,.butNew.Top,'??????','','DO formPrnOrder',39,.butNew.Width,'??????'
     DO addButtonOne WITH 'frmOrd','butSearch',.ordPrik.Left,.butNew.Top,'?????','','DO procSearchOrder',39,.butNew.Width,'?????' 
     DO addButtonOne WITH 'frmOrd','butRet',.ordPrik.Left,.butNew.Top,'???????','','DO procExitOrder',39,.butNew.Width,'???????' 
     
     .butNew.Left=.ordPrik.Left+(.ordPrik.Width+.comboBox1.Width-.butNew.Width*6-25)/2
     .butRead.Left=.butNew.Left+.butNew.Width+5
     .butDel.Left=.butRead.Left+.butRead.Width+5     
     .butPrn.Left=.butDel.Left+.butDel.Width+5
     .butSearch.Left=.butPrn.Left+.butPrn.Width+5
     .butRet.Left=.butSearch.Left+.butSearch.Width+5
     
     DO addButtonOne WITH 'frmOrd','butTxt',.ordPrik.Left,.butNew.Top,'?????','','DO formTextOrder',39,.butNew.Width,'???????????? ?????' 
     DO addButtonOne WITH 'frmOrd','butSaveRec',.ordPrik.Left,.butNew.Top,'????????','','DO saveRecOrder WITH .T.',39,.butNew.Width,'????????' 
     DO addButtonOne WITH 'frmOrd','butRetRead',.ordPrik.Left,.butNew.Top,'???????','','DO saveRecOrder WITH .F.',39,.butNew.Width,'???????' 
     .butTxt.Visible=.F.
     .butSaveRec.Visible=.F.
     .butRetRead.Visible=.F.          
     .butTxt.Left=.ordPrik.Left+(.ordPrik.Width+.comboBox1.Width-.butTxt.Width*3-10)/2
     .butSaveRec.Left=.butTxt.Left+.butTxt.Width+5
     .butRetRead.Left=.butSaveRec.Left+.butSaveRec.Width+5
     
     DO addButtonOne WITH 'frmOrd','butRetNew',.ordPrik.Left,.butNew.Top,'???????','','DO procButRetNew',39,.butNew.Width,'???????' 
     .butRetNew.Left=.ordPrik.Left+(.ordPrik.Width+.comboBox1.Width-.butRetNew.Width)/2
     .butRetNew.Visible=.F.     
     DO addButtonOne WITH 'frmOrd','butDelRec',.ordPrik.Left,.butNew.Top,'???????','','DO delRecOrder WITH .T.',39,.butNew.Width,'??????? ??????' 
     DO addButtonOne WITH 'frmOrd','butDelRet',.ordPrik.Left,.butNew.Top,'???????','','DO delRecOrder WITH .F.',39,.butNew.Width,'???????' 
     .butDelRec.Visible=.F.
     .butDelRet.Visible=.F.
     .butDelRec.Left=.ordPrik.Left+(.ordPrik.Width+.comboBox1.Width-.butDelRec.Width*2-10)/2
     .butDelRet.Left=.butDelRec.Left+.butDelRec.Width+10     
     
       * ??????? ??? ??????  
     DO adTboxNew WITH 'frmOrd','tBoxSearch',.butNew.Top,.ordPrik.Left+5,200,dHeight,'search_cx',.F.,.F.,0         
     .tBoxSearch.procForkeyPress='DO keyPressSearch'
     .tBoxSearch.Visible=.F.
     DO addButtonOne WITH 'frmOrd','butSearchRec',.tBoxSearch.Left+.tBoxSearch.Width+5,.butNew.Top,'?????','','DO searchRecOrder',39,.butNew.Width,'????? ??????' 
     DO addButtonOne WITH 'frmOrd','butSearchNext',.butSearchRec.Left,.butNew.Top,'?????','','DO searchNextRecOrder',39,.butNew.Width,'?????? ?????' 
     DO addButtonOne WITH 'frmOrd','butSearchRet',.butSearchRec.Left+.butSearchRec.Width+5,.butNew.Top,'???????','','DO searchRetOrder',39,.butNew.Width,'???????' 
     .tBoxSearch.Top=.butSearchRec.Top+(.butSearchRec.Height-.tBoxSearch.Height)/2
     .butSearchRec.Visible=.F.
     .butSearchNext.Visible=.F.
     .butSearchRet.Visible=.F.
     
    .Autocenter=.T.      
     .grdPers.columns(.grdPers.ColumnCount).SetFocus
     IF EMPTY(dateOrdNew)
        .tBox2.SetFocus
     ENDIF
     .Show
     
ENDWITH
********************************************************************************************************************************
PROCEDURE procsupOrd
SELECT curSprOrder
SET FILTER TO kod1=curSupOrder.kod.OR.kod=401
********************************************************************************************************************************
PROCEDURE procExitOrder
SELECT peoporder
IF logNewRec.AND.kodpeop=0
   DELETE 
ENDIF
frmOrd.Visible=.F.
frmOrd.Release
SELECT people
IF readNidPeop#0
   oldOrd=SYS(21)
   SET ORDER TO 4
   SEEK readNidPeop
   SET ORDER TO &oldOrd
ENDIF 
DO changeRowGrdPers
*fOrdStart.release
********************************************************************************************************************************
PROCEDURE procButRetNew
WITH frmOrd
     logNewRec=.F.
     logRead=.F.
     .butNew.Visible=.T.
     .butRead.Visible=.T.
     .butDel.Visible=.T.
     .butPrn.Visible=.T.
     .butSearch.Visible=.T.     
     .butRet.Visible=.T.
     .butRetNew.Visible=.F.
     .grdPers.Enabled=.T.
     .grdPers.SetAll('Enabled',.F.,'ColumnMy')
     .grdPers.Columns(.grdPers.ColumnCount).Enabled=.T.
     .grdPers.Columns(.grdPers.ColumnCount).SetFocus   
     DO changeRowPeopOrder
     .Refresh    
ENDWITH
********************************************************************************************************************************
PROCEDURE procPrikOrd
ON ERROR DO erSup
logPer=.F.
newKodPrik=curSprOrder.kod
newprocOrd=ALLTRIM(curSprOrder.procord)
procForOsnov=ALLTRIM(cursprorder.osnov)
frmOrd.butRetNew.Visible=.F.
ON ERROR 
IF !EMPTY(newProcOrd)
   &newprocOrd
ENDIF    
********************************************************************************************************************************
PROCEDURE validOrdernum 
* parPadejDol=1 ??????????? - ???? - people.fior,sprdolj.namer
* parPadejDol=2 ????????? - ???? - people.fiod,sprdolj.named
* parPadejDol=3 ??????????? ??? ???????????? - ??? - people.fiov,sprdolj.namet
oldKodpeop=newKodpeop
newKodPeop=IIF(SEEK(newKodPeop,'people',1),people.num,0)
newNidPeop=people.nid
newFioPeop=people.fio
*oldPeopRec=people.RECNO()
kodSex=IIF(people.sex#0,people.sex,1)
str_ini=''
DO procfioini WITH 'people.fiot' 
IF !EMPTY(procForOsnov) 
   newOsnov=&procForOsnov
ENDIF 
newNidJob=IIF(SEEK(newKodPeop,'people',1),people.nidjob,0)
newNdolj=IIF(SEEK(newKd,'sprdolj',1),sprdolj.name,'') 
DO CASE
   CASE logNewRec
        IF newnidJob>0      
           newKp=IIF(SEEK(newnidjob,'datjob',7),datJob.kp,0)
           newKd=datJob.kd
           oldNidJob=datjob.nid
           newNpodr=IIF(SEEK(newKp,'sprpodr',1),sprpodr.name,'')           
        ELSE 
           SELECT datJob
           LOCATE FOR kodpeop=newKodPeop.AND.INLIST(tr,1,3).AND.EMPTY(dateOut)
           newKp=datJob.kp
           newKd=datJob.kd
           oldNidJob=datjob.nid
           newNpodr=IIF(SEEK(newKp,'sprpodr',1),sprpodr.name,'')
        
           SELECT curPeopOrder
        ENDIF   
        
       
   CASE !logNewRec  
        IF newKodPeop#oldKodPeop
           newKp=IIF(SEEK(newnidjob,'datjob',7),datJob.kp,0)
           newKd=datJob.kd
           newNpodr=IIF(SEEK(newKp,'sprpodr',1),sprpodr.name,'')
         
        ELSE   
           newNpodr=curPeopOrder.npodr
           newNdolj=curPeopOrder.ndolj
           DO padejInOrder
        ENDIF   
ENDCASE 
IF logNewRec.AND.logPer
   SELECT * FROM datOtp WHERE nidpeop=newNidpeop INTO CURSOR curOrdOtp ORDER BY perBeg DESCENDING READWRITE 
   SELECT curOrdOtp
   DELETE FOR kodotp>3   
   GO TOP
   LOCATE FOR !EMPTY(perEnd) && ???????????? ??? ???????? ?? ?????? ????????? ? ????? ?????????? ???????
   IF kodotp=1
      newPerBeg=IIF(!EMPTY(perEnd),perEnd+1,newPerBeg)
      newPerEnd=IIF(!EMPTY(newPerBeg),IIF(MOD(YEAR(newPerBeg),4)=0,newPerBeg+365-1,newPerBeg+365),newPerEnd)
   ELSE 
      newPerBeg=IIF(!EMPTY(perBeg),perBeg,newPerBeg)
      newPerEnd=IIF(!EMPTY(newPerBeg),IIF(MOD(YEAR(newPerBeg),4)=0,newPerBeg+365-1,newPerBeg+365),newPerEnd)
   ENDIF 
   
   IF newKodPrik=51
      SELECT datOtp
      ordOtpold=SYS(21)
      SET ORDER TO 6
      SEEK newNidPeop
      daycx=0
      SCAN WHILE nidPeop=newNidPeop
           IF kodotp=6       
              DO CASE
                 CASE begotp>=newPerBeg.AND.endotp<=newPerEnd
                      daycx=daycx+kvoday                  
                 CASE begotp<newPerBeg.AND.endotp<newPerEnd.AND.endotp=>newPerBeg
                      daycx=daycx+(endotp-newPerBeg)+1
                 CASE endotp>newPerEnd.AND.begOtp<=newPerEnd.AND.begOtp>newPerBeg
                      daycx=daycx+(endotp-newPerBeg)+1     
              ENDCASE
           ENDIF           
      ENDSCAN
      IF daycx>14
         dayCost=daycx-14
         frmord.tboxdolj2.Refresh
      ENDIF
      SET ORDER TO &ordOtpOld
   ENDIF    
   frmOrd.tBoxPerBeg.Refresh
   frmOrd.tBoxPerEnd.Refresh
   SELECT curPeopOrder   
ENDIF
IF logDatJob 
   ON ERROR DO erSup
   SELECT curOrdJob
   ZAP    
   APPEND FROM datjob FOR kodpeop=newkodpeop
   DO CASE
      CASE newKodprik=106
           DELETE FOR tr#2
      CASE newKodprik=107
           DELETE FOR tr#4
      CASE newKodPrik=70
           DELETE FOR tr#3  
      CASE newKodPrik=115
           DELETE FOR tr#5                           
   ENDCASE   
   REPLACE npord WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.nameord,sprpodr.name) ,ndord WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namer,sprdolj.name) ALL

   IF logNewRec
      DELETE FOR !EMPTY(dateOut) 
      GO TOP 
   ELSE 
      DELETE FOR !EMPTY(dateOut).AND.dateOut<dateOut
      GO TOP   
   ENDIF 
   oldSovmJob=IIF(logNewRec,nid,oldnid)   
   newSkp=kp
   newSPodr=npord
   newSKd=kd
   newSDolj=ALLTRIM(ndord)
   newKse=kse
   nidJobNew=curOrdJob.nid
   newTr=curOrdJob.tr    
   strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
   frmOrd.tBoxPodrNew.Refresh
   frmOrd.comboBox5.ControlSource='newSDolj'
   frmOrd.comboBox5.Refresh     
   ON ERROR 
   SELECT curpeoporder
 
ENDIF 
frmOrd.Refresh
*******************************************************************************************************************************
PROCEDURE newRecInOrder
PARAMETERS parNewRec
logNewRec=parNewRec
logRead=.T.
IF logNewRec
   DO removeObjectKadr
   strPrik='' 
   SELECT curPeopOrder  
   newNpp=npp+1
   WITH frmOrd
        .comboBox1.ControlSource='strprik'
        .butNew.Visible=.F.
        .butRead.Visible=.F.
        .butDel.Visible=.F.
        .butPrn.Visible=.F.   
        .butSearch.Visible=.F.  
        .butRet.Visible=.F.
        .butTxt.Visible=IIF(logNewRec,.F.,.T.)     
        .butSaveRec.Visible=IIF(logNewRec,.F.,.T.)     
        .butRetRead.Visible=IIF(logNewRec,.F.,.T.)     
        .butRetNew.Visible=IIF(logNewRec,.T.,.F.)
        .grdPers.Enabled=.F.
        .comboBox1.Enabled=.T. 
        IF typeOrdNew=2
           .comboBox99.Enabled=.T.      
        ENDIF    
   ENDWITH  
   IF EMPTY(dateOrdNew)
      .tBox2.SetFocus
   ELSE 
      IF typeOrdNew=2
         .comboBox99.Enabled=.T.      
      ELSE
         .comboBox1.SetFocus
      ENDIF    
   ENDIF
ELSE 
   newNpp=curpeoporder.npp
   procread=IIF(SEEK(curPeopOrder.supOrd,'sprorder',1),sprorder.procord,'')
   WITH frmOrd
        .butNew.Visible=.F.
        .butRead.Visible=.F.
        .butDel.Visible=.F.
        .butPrn.Visible=.F. 
        .butSearch.Visible=.F.    
        .butRet.Visible=.F.
        .butTxt.Visible=IIF(logNewRec,.F.,.T.)     
        .butSaveRec.Visible=IIF(logNewRec,.F.,.T.)     
        .butRetRead.Visible=IIF(logNewRec,.F.,.T.)     
        .butRetNew.Visible=IIF(logNewRec,.T.,.F.)
        .grdPers.Enabled=.F.
        .comboBox1.Enabled=.T.     
        IF typeOrdNew=2
           .comboBox99.Enabled=.T.      
        ENDIF 
        IF EMPTY(dateOrdNew)
           .tBox2.SetFocus
        ELSE 
           IF typeOrdNew=2
              .comboBox99.Enabled=.T.      
           ELSE
              .comboBox1.SetFocus
           ENDIF              
        ENDIF
        &procRead
   ENDWITH
ENDIF
********************************************************************************************************************************
*       ????? ?? ???????? ?????????
********************************************************************************************************************************
PROCEDURE priemkont 
PARAMETERS par1
parPadejDol=3
parPadejFio=1
DO newNidOrder
DO removeObjectKadr
DO resetVarTot
newPkont=IIF(logNewRec,0,curpeoporder.pkont)
newDdop=IIF(logNewRec,0,curpeoporder.ddop)
objFocus='frmOrd.tBoxBeg'
newLogApp=.T.
newNpp=0
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder    
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
         
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',2,.comboBox1.Left,.ordPodr.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodr',.F.,.T.  
     .comboBox2.Visible=IIF(par1,.T.,.F.)
     .comboBox2.DisplayCount=15
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     .tBoxPodr.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',3,.comboBox1.Left,.ordDol.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDolj',.F.,.T.       
     WITH .comboBox3         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     .tBoxDolj.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)  
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordKse.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
     .tBoxKont.Alignment=0
       
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
     .tBoxDay.Alignment=0
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0    
      
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
     IF par1 
        ON ERROR DO erSup     
        .comboBox2.SetFocus
        .comboBox3.SetFocus
        .comboBox11.SetFocus
        .tBoxFio.SetFocus
        ON ERROR         
     ENDIF
ENDWITH
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ?????? ?? ???????? ?????????
********************************************************************************************************************************
PROCEDURE textPriemKont
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')

cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol

SELECT txtOrder
REPLACE txtprn WITH '??????? '+ALLTRIM(cOrdFio)+' ? '+strDateBeg+' ?? ?????? ?? ???????? ????????? ?? '+strDateEnd+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? ??????? ???????? ???????? ?????????? , ????????? ?????????????? ???? ??????????????:'+CHR(13)+;
        '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
        '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)       
       
********************************************************************************************************************************
PROCEDURE saveRecPriemKont
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPriemKont	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,pkont WITH newPkont,ddop WITH newDDop,osnov WITH newOsnov,;
        txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
    
DO saveDimOrd  
DO saveKadrOrder
********************************************************************************************************************************
*       ????? ?? ???????? ????????? (?????????????)
********************************************************************************************************************************
PROCEDURE priemkontmany 
PARAMETERS par1
parPadejDol=3
parPadejFio=1
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot

PUBLIC newKp2,newNpodr2,newKd2,newNDolj2,newKse2,newKp3,newNpodr3,newKd3,newNDolj3,newKse3,newKp4,newNpodr4,newKd4,newNDolj4,newKse4
newKp2=IIF(logNewRec,0,curpeoporder.kp2)
newNPodr2=IIF(logNewRec,'',curpeoporder.npodr2)
newKd2=IIF(logNewRec,0,curpeoporder.kd2)
newNDolj2=IIF(logNewRec,'',curpeoporder.ndolj2)
newKse2=IIF(logNewRec,0.00,curpeoporder.kse2)

newKp3=IIF(logNewRec,0,curpeoporder.kp3)
newNPodr3=IIF(logNewRec,'',curpeoporder.npodr3)
newKd3=IIF(logNewRec,0,curpeoporder.kd3)
newNDolj3=IIF(logNewRec,'',curpeoporder.ndolj3)
newKse3=IIF(logNewRec,0.00,curpeoporder.kse3)

newKp4=IIF(logNewRec,0,curpeoporder.kp4)
newNPodr4=IIF(logNewRec,'',curpeoporder.npodr4)
newKd4=IIF(logNewRec,0,curpeoporder.kd4)
newNDolj4=IIF(logNewRec,'',curpeoporder.ndolj4)
newKse4=IIF(logNewRec,0.00,curpeoporder.kse4)

newPkont=IIF(logNewRec,0,curpeoporder.pkont)
newDdop=IIF(logNewRec,0,curpeoporder.ddop)
objFocus='frmOrd.tBoxBeg'
newLogApp=.T.
newNpp=0
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder    
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
         
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',2,.comboBox1.Left,.ordPodr.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodr',.F.,.T.  
     .comboBox2.Visible=IIF(par1,.T.,.F.)
     .comboBox2.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     .tBoxPodr.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',3,.comboBox1.Left,.ordDol.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDolj',.F.,.T.       
     WITH .comboBox3         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     .tBoxDolj.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)  
     
      *****2
     
     DO adTboxAsCont WITH 'frmOrd','ordPodr2',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',22,.comboBox1.Left,.ordPodr2.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr2','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodrMany WITH 2',.F.,.T.  
     .comboBox22.Visible=IIF(par1,.T.,.F.)
     .comboBox22.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordPodr2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr2',.F.,.F.,0  
     .tBoxPodr2.Visible=IIF(par1,.F.,.T.) 
     
     
     DO adTboxAsCont WITH 'frmOrd','ordDol2',.ordPrik.Left,.ordPodr2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',32,.comboBox1.Left,.ordDol2.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj2','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDoljMany WITH 2',.F.,.T.       
     WITH .comboBox32         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj2',.ordDol2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj2',.F.,.F.,0  
     .tBoxDolj2.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse2',.ordPrik.Left,.ordDol2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse2',.comboBox1.Left,.ordKse2.Top,dheight,RetTxtWidth('9999999999'),'newKse2',0.25,.F.,0,1.5 
     .spinKse2.Enabled=IIF(par1,.T.,.F.)   
     
     
     *****3
     DO adTboxAsCont WITH 'frmOrd','ordPodr3',.ordPrik.Left,.ordKse2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',23,.comboBox1.Left,.ordPodr3.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr3','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodrMany WITH 3',.F.,.T.  
     .comboBox23.Visible=IIF(par1,.T.,.F.)
     .comboBox23.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr3',.ordPodr3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr3',.F.,.F.,0  
     .tBoxPodr3.Visible=IIF(par1,.F.,.T.) 
          
     DO adTboxAsCont WITH 'frmOrd','ordDol3',.ordPrik.Left,.ordPodr3.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',33,.comboBox1.Left,.ordDol3.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj3','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDoljMany WITH 3',.F.,.T.       
     WITH .comboBox33         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj3',.ordDol3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj3',.F.,.F.,0  
     .tBoxDolj3.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse3',.ordPrik.Left,.ordDol3.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse3',.comboBox1.Left,.ordKse3.Top,dheight,RetTxtWidth('9999999999'),'newKse3',0.25,.F.,0,1.5 
     .spinKse3.Enabled=IIF(par1,.T.,.F.)   
               
      *****43
     DO adTboxAsCont WITH 'frmOrd','ordPodr4',.ordPrik.Left,.ordKse3.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',24,.comboBox1.Left,.ordPodr4.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr4','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodrMany WITH 4',.F.,.T.  
     .comboBox24.Visible=IIF(par1,.T.,.F.)
     .comboBox24.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr4',.ordPodr4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr4',.F.,.F.,0  
     .tBoxPodr4.Visible=IIF(par1,.F.,.T.) 
          
     DO adTboxAsCont WITH 'frmOrd','ordDol4',.ordPrik.Left,.ordPodr4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',34,.comboBox1.Left,.ordDol4.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj4','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDoljMany WITH 4',.F.,.T.       
     WITH .comboBox34         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj4',.ordDol4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj4',.F.,.F.,0  
     .tBoxDolj4.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse4',.ordPrik.Left,.ordDol4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse4',.comboBox1.Left,.ordKse4.Top,dheight,RetTxtWidth('9999999999'),'newKse4',0.25,.F.,0,1.5 
     .spinKse4.Enabled=IIF(par1,.T.,.F.)               
          
     **?????????
     
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordKse4.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
     .tBoxKont.Alignment=0
       
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
     .tBoxDay.Alignment=0
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0 
          
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
     IF par1  
        ON ERROR DO erSup    
        .comboBox2.SetFocus
        .comboBox3.SetFocus
        .comboBox11.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE procPriemPodrMany
PARAMETERS parmany
SELECT datjob
SET ORDER TO 2
DO CASE
   CASE parMany=2
        newKp2=curSprPodr.kod  
        newNPodr2=curSprPodr.name 
        SELECT curDolPodr
        SET FILTER TO kp=newKp2
        ksesup=0 
        DO sumVacKse
        frmOrd.ComboBox32.RowSource='curDolPodR.name'
        WITH .comboBox32         
             .DisplayCount=15
             .ColumnCount=3
             .ColumnWidths='0,50,500'
             .RowSource="curDolPodr.name,strVac,name"
        ENDWITH 
        frmOrd.ComboBox32.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
        frmOrd.ComboBox32.RowSourceType=6
        frmOrd.comboBox32.ProcForValid='DO procPriemDoljmany WITH 2'
   CASE parMany=3
        newKp3=curSprPodr.kod  
        newNPodr3=curSprPodr.name 
        SELECT curDolPodr
        SET FILTER TO kp=newKp3
        ksesup=0 
        DO sumVacKse
        frmOrd.ComboBox33.RowSource='curDolPodR.name'
        WITH .comboBox33         
             .DisplayCount=15
             .ColumnCount=3
             .ColumnWidths='0,50,500'
             .RowSource="curDolPodr.name,strVac,name"
        ENDWITH 
        frmOrd.ComboBox33.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
        frmOrd.ComboBox33.RowSourceType=6
        frmOrd.comboBox33.ProcForValid='DO procPriemDoljmany WITH 3'
   CASE parMany=4
        newKp4=curSprPodr.kod  
        newNPodr4=curSprPodr.name 
        SELECT curDolPodr
        SET FILTER TO kp=newKp4
        ksesup=0 
        DO sumVacKse
        frmOrd.ComboBox34.RowSource='curDolPodR.name'
        WITH .comboBox34         
             .DisplayCount=15
             .ColumnCount=3
             .ColumnWidths='0,50,500'
             .RowSource="curDolPodr.name,strVac,name"
        ENDWITH 
        frmOrd.ComboBox34.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
        frmOrd.ComboBox34.RowSourceType=6
        frmOrd.comboBox34.ProcForValid='DO procPriemDoljmany WITH 4'
ENDCASE
KEYBOARD '{TAB}'
************************************************************************************************************************
PROCEDURE sumVacKse
ksesup=0     
SELECT datjob
SET ORDER TO 2
SEEK STR(curDolPodr.kp,3)+STR(curDolPodr.kd,3)
SCAN WHILE kp=curdolpodr.kp.AND.kd=curdolpodr.kd
     DO CASE
        CASE dekotp
        CASE !EMPTY(dateOut).AND.dateOut<=DATE()
        CASE dateBeg>DATE()
        OTHERWISE
             ksesup=ksesup+kse
     ENDCASE        
ENDSCAN
SELECT curDolPodr
REPLACE strVac WITH IIF(kse-ksesup=0,'',STR(kse-ksesup,6,2))
********************************************************************************************************************************
PROCEDURE procPriemDoljMany
PARAMETERS parMany
DO CASE
   CASE parMany=2
        newKd2=curDolPodr.kd  
        newNDolj2=curDolPodr.name
        frmOrd.ComboBox32.ControlSource='newNDolj2'
        frmOrd.comboBox32.Refresh
   CASE parMany=3
        newKd3=curDolPodr.kd  
        newNDolj3=curDolPodr.name
        frmOrd.ComboBox33.ControlSource='newNDolj3'
        frmOrd.comboBox33.Refresh     
   CASE parMany=4
        newKd4=curDolPodr.kd  
        newNDolj4=curDolPodr.name
        frmOrd.ComboBox34.ControlSource='newNDolj4'
        frmOrd.comboBox34.Refresh     
ENDCASE
KEYBOARD '{TAB}'
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ?????? ?? ???????? ????????? (?????????????)
********************************************************************************************************************************
PROCEDURE textPriemKontMany
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
kodSex=IIF(SEEK(newKodPeop,'people',1).AND.people.sex#0,people.sex,1)
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol

strmany=' (?? ??? '
IF newKse#0
   strmany=strmany+STR(newKse,4,2)+' ????????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' '
ENDIF

IF newKse2#0 
   cOrdPodr2=''
   cOrdDol2=''
   cOrdPodr2=IIF(SEEK(newKp2,'sprpodr',1),ALLTRIM(sprpodr.nameord),ALLTRIM(newNPodr2))
   IF SEEK(newKd2,'sprdolj',1)
      IF sprdolj.logSex.AND.kodSex=1
         DO CASE
            CASE parPadejDol=1
                 cOrdDol2=IIF(!EMPTY(sprdolj.namerm),sprdolj.namerm,newNDolj)
            CASE parPadejDol=2
                 cOrdDol2=IIF(!EMPTY(sprdolj.namedm),sprdolj.namedm,newNDolj)                                
            CASE parPadejDol=3
                 cOrdDol2=IIF(!EMPTY(sprdolj.nametm),sprdolj.nametm,newNDolj)     
         ENDCASE
      ELSE        
         DO CASE
            CASE parPadejDol=1
                 cOrdDol2=IIF(!EMPTY(sprdolj.namer),sprdolj.namer,newNDolj)
            CASE parPadejDol=2
                 cOrdDol2=IIF(!EMPTY(sprdolj.named),sprdolj.named,newNDolj)                                
            CASE parPadejDol=3
                 cOrdDol2=IIF(!EMPTY(sprdolj.namet),sprdolj.namet,newNDolj)     
         ENDCASE             
      ENDIF
   ELSE
      cOrdDol2=newNDolj2
   ENDIF    
   strmany=strmany+STR(newKse2,4,2)+' ????????? '+LOWER(cOrdDol2)+' '+LOWER(cOrdPodr2)+' '
ENDIF

IF newKse3#0
   cOrdPodr3=''
   cOrdDol3=''
   cOrdPodr3=IIF(SEEK(newKp3,'sprpodr',1),ALLTRIM(sprpodr.nameord),ALLTRIM(newNPodr3))
   IF SEEK(newKd3,'sprdolj',1)
      IF sprdolj.logSex.AND.kodSex=1
         DO CASE
            CASE parPadejDol=1
                 cOrdDol3=IIF(!EMPTY(sprdolj.namerm),sprdolj.namerm,newNDolj)
            CASE parPadejDol=2
                 cOrdDol3=IIF(!EMPTY(sprdolj.namedm),sprdolj.namedm,newNDolj)                                
            CASE parPadejDol=3
                 cOrdDol3=IIF(!EMPTY(sprdolj.nametm),sprdolj.nametm,newNDolj)     
         ENDCASE
      ELSE        
         DO CASE
            CASE parPadejDol=1
                 cOrdDol3=IIF(!EMPTY(sprdolj.namer),sprdolj.namer,newNDolj)
            CASE parPadejDol=2
                 cOrdDol3=IIF(!EMPTY(sprdolj.named),sprdolj.named,newNDolj)                                
            CASE parPadejDol=3
                 cOrdDol3=IIF(!EMPTY(sprdolj.namet),sprdolj.namet,newNDolj)     
         ENDCASE             
      ENDIF
   ELSE
      cOrdDol3=newNDolj3
   ENDIF       
   strmany=strmany+STR(newKse3,4,2)+' ????????? '+LOWER(cOrdDol3)+' '+LOWER(cOrdPodr3)+' '
ENDIF

IF newKse4#0
   cOrdPodr4=''
   cOrdDol4=''
   cOrdPodr4=IIF(SEEK(newKp4,'sprpodr',1),ALLTRIM(sprpodr.nameord),ALLTRIM(newNPodr3))
   IF SEEK(newKd4,'sprdolj',1)
      IF sprdolj.logSex.AND.kodSex=1
         DO CASE
            CASE parPadejDol=1
                 cOrdDol4=IIF(!EMPTY(sprdolj.namerm),sprdolj.namerm,newNDolj)
            CASE parPadejDol=2
                 cOrdDol4=IIF(!EMPTY(sprdolj.namedm),sprdolj.namedm,newNDolj)                                
            CASE parPadejDol=3
                 cOrdDol4=IIF(!EMPTY(sprdolj.nametm),sprdolj.nametm,newNDolj)     
         ENDCASE
      ELSE        
         DO CASE
            CASE parPadejDol=1
                 cOrdDol4=IIF(!EMPTY(sprdolj.namer),sprdolj.namer,newNDolj)
            CASE parPadejDol=2
                 cOrdDol4=IIF(!EMPTY(sprdolj.named),sprdolj.named,newNDolj)                                
            CASE parPadejDol=3
                 cOrdDol4=IIF(!EMPTY(sprdolj.namet),sprdolj.namet,newNDolj)     
         ENDCASE             
      ENDIF
   ELSE
      cOrdDol4=newNDolj4
   ENDIF 
   strmany=strmany+STR(newKse4,4,2)+' ????????? '+LOWER(cOrdDol4)+' '+LOWER(cOrdPodr4)+' '
ENDIF
strmany=strmany+')'
SELECT txtOrder
REPLACE txtprn WITH '??????? '+ALLTRIM(cordFio)+' ? '+strDateBeg+' ?? ?????? ?? ???????? ????????? ?? '+strDateEnd+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+strmany+' ? ??????? ???????? ???????? ?????????? , ????????? ?????????????? ???? ??????????????:'+CHR(13)+;
        '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
        '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)       
       
********************************************************************************************************************************
PROCEDURE saveRecPriemKontMany
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPriemKontMany	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,pkont WITH newPkont,ddop WITH newDDop,osnov WITH newOsnov,;
        txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,sex WITH kodsex,;
        kp2 WITH newKp2,npodr2 WITH newNPodr2,kd2 WITH newKd2,ndolj2 WITH newNDolj2,kse2 WITH newKse2,;
        kp3 WITH newKp3,npodr3 WITH newNPodr3,kd3 WITH newKd3,ndolj3 WITH newNDolj3,kse3 WITH newKse3,;
        kp4 WITH newKp4,npodr4 WITH newNPodr4,kd4 WITH newKd4,ndolj4 WITH newNDolj4,kse4 WITH newKse4,dord WITH repdorder,nord WITH repnorder            
    
DO saveDimOrd  
DO saveKadrOrder
********************************************************************************************************************************
*       ????? ?? ???????? ???????? ????????????????
********************************************************************************************************************************
PROCEDURE priemvn
PARAMETERS par1
IF !USED('sprtot')
   USE sprtot IN 0   
ENDIF
IF !USED('curSrokDog')
   SELECT kod,name FROM sprtot WHERE sprtot.kspr=25 INTO CURSOR curSrokDog READWRITE && ?????? ??? ????? ???????? ??????????
   SELECT curSrokDog
   INDEX ON kod TAG T1
ENDIF  
parPadejDol=3
parPadejFio=1
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newKse=IIF(logNewRec,0.00,curpeoporder.kse)
newTr=IIF(logNewRec,3,curpeoporder.tr)
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
newVdog=IIF(logNewRec,0,curPeopOrder.vdog)
strVdog=IIF(SEEK(newVdog,'curSrokDog',1),curSrokDog.name,'')
newTabZam=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,1,5)))
newFioZam=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl,6,60))

newLogApp=.T.
newNpp=0
objFocus='frmOrd.comboBox22'
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr2',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'???????? ???????',1,1 
     DO addcombomy WITH 'frmOrd',22,.comboBox1.Left,.ordPodr2.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'strVDog','ALLTRIM(curSrokDog.name)',6,'','DO validVidDog',.F.,.T.  
     .comboBox22.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordPodr2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'strVDog',.F.,.F.,0  
     .tBoxPodr2.Visible=IIF(par1,.F.,.T.) 
     
     
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordPodr2.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0  
            
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2         
             
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',2,.comboBox1.Left,.ordPodr.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodr',.F.,.T.  
     .comboBox2.Visible=IIF(par1,.T.,.F.)
     .comboBox2.DisplayCount=15
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     .tBoxPodr.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',3,.comboBox1.Left,.ordDol.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj','ALLTRIM(curDolPodr.named)',6,.F.,'DO procPriemDolj',.F.,.T.       
     WITH .comboBox3         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     .tBoxDolj.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)  
     
     DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'??? (???? ????????)',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioZam',.F.,IIF(par1.AND.newVDog=2,.T.,.F.),0      
     .tBoxFioNew.procforChange='DO changePeopDek'   
     DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectPeopDek',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlntNew.Enabled=IIF(par1.AND.newVDog=2,.T.,.F.)
     
                 
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordFioNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
   
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10'  
     DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
     WITH .listBox2
          .RowSource='curSuplpeop.fior'         
          .RowSourceType=2
          .Visible=.F.        
          .procForDblClick='DO validListDek'
          .procForLostFocus='DO lostFocusPeop2'
          .Height=.Parent.Height-.Top
          .Enabled=IIF(newVdog=2,.T.,.F.)
     ENDWITH 
           
     IF par1   
        ON ERROR DO erSup   
        .comboBox2.SetFocus
        .comboBox3.SetFocus
        .comboBox11.SetFocus
        .comboBox22.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE validVidDog
newVDog=curSrokDog.kod
DO CASE
   CASE newVdog=2
        frmOrd.listBox2.Enabled=.T.
        frmOrd.tBoxFioNew.Enabled=.T.
        frmOrd.butKlntNew.Enabled=.T.
   OTHERWISE 
        frmOrd.listBox2.Enabled=.F.
        frmOrd.tBoxFioNew.Enabled=.F.  
        frmOrd.butKlntNew.Enabled=.F.
ENDCASE
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ??????  ???????? ????????????
********************************************************************************************************************************
PROCEDURE textPriemVn
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol

SELECT txtOrder
DO CASE
   CASE newVdog=1 
        REPLACE txtprn WITH '??????? '+ALLTRIM(cOrdFio)+' ? '+strDateBeg+' ?? '+strDateEnd+' ?? ?????? ? ??????? ???????? ???????????????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ?? '+STR(newKse,4,2)+;
                ' ??????? ??????? ? ??????? ???????? ???????? ??????????'+CHR(13)+CHR(13)+;
                '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;               
                '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
                '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
                '?????????: '+ALLTRIM(newOsnov)       
                
   CASE newVdog=2 
        REPLACE txtprn WITH '??????? '+ALLTRIM(cOrdFio)+' ? '+strDateBeg+' ?? ?????? ? ??????? ???????? ???????????????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ?? '+STR(newKse,4,2)+;
                ' ??????? ??????? ?? ?????? ????????? ????????? ????????? '+ALLTRIM(newFioZam) +CHR(13)+CHR(13)+;
                '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;               
                '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
                '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
                '?????????: '+ALLTRIM(newOsnov)  
   CASE newVdog=3 
        REPLACE txtprn WITH '??????? '+ALLTRIM(cOrdFio)+' ? '+strDateBeg+' ?? ?????? ? ??????? ???????? ???????????????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ?? '+STR(newKse,4,2)+;
                ' ??????? ??????? ? ??????? ???????? ???????? ??????????'+CHR(13)+CHR(13)+;
                '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;               
                '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
                '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
                '?????????: '+ALLTRIM(newOsnov)                                
ENDCASE        
********************************************************************************************************************************
PROCEDURE saveRecPriemVn
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPriemVn	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidpeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,;
        txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,vdog WITH newVdog,varSupl WITH IIF(newVDog=2,PADR(ALLTRIM(STR(newTabZam)),5,' ')+PADR(ALLTRIM(newFioZam),60,' '),''),;
        dord WITH repdorder,nord WITH repnorder,sex WITH kodsex                      
DO saveDimOrd  
DO saveKadrOrder
********************************************************************************************************************************
*       ????? ?? ???????? ????????? ????????
********************************************************************************************************************************
PROCEDURE priemDog 
PARAMETERS par1,par2
nuvol=par2

**par2=1  ???????
**par2=2  ?? ????? ?????????? ????????? ?????????
**par2=3  ?? ?????????????? ????

parPadejDol=3
parPadejFio=1
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newPkont=IIF(logNewRec,0,curpeoporder.pkont)
newDdop=IIF(logNewRec,0,curpeoporder.ddop)
DO CASE
   CASE logNewRec.AND.nuvol=4
        newTr=3 
   CASE logNewRec
        newTr=1
   OTHERWISE 
        newTr=curpeoporder.tr           
ENDCASE
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
newTabZam=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,1,5)))
newFioZam=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl,6,60))
        
objFocus='frmOrd.tBoxBeg'
newLogApp=.T.
newNpp=0
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder    
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1.AND.nuvol=1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
         
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',2,.comboBox1.Left,.ordPodr.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodr',.F.,.T.  
     .comboBox2.Visible=IIF(par1,.T.,.F.)
     .comboBox2.DisplayCount=15
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     .tBoxPodr.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',3,.comboBox1.Left,.ordDol.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDolj',.F.,.T.       
     WITH .comboBox3         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     .tBoxDolj.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)  
   
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
     .tBoxDay.Alignment=0
               
     IF INLIST(nuvol,2,4)
        DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'??? (???? ????????)',1,1      
        DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioZam',.F.,IIF(par1,.T.,.F.),0      
        .tBoxFioNew.procforChange='DO changePeopDek'   
        DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
        DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectPeopDek',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
        .butKlntNew.Enabled=IIF(par1,.T.,.F.)
     ENDIF          
         
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,IIF(INLIST(nuvol,2,4),.ordFioNew.Top,.ordDay.Top)+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0 
               
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
     
     IF INLIST(nuvol,2,4)
        DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
        WITH .listBox2
             .RowSource='curSuplpeop.fior'         
             .RowSourceType=2
             .Visible=.F. 
             .Height=.Parent.Height-.Top       
             .procForDblClick='DO validListDek'
             .procForLostFocus='DO lostFocusPeop2'
        ENDWITH 
     ENDIF
     
     IF par1      
        ON ERROR DO erSup
        .comboBox2.SetFocus
        .comboBox3.SetFocus
        .comboBox11.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ?????? ?? ???????? ?????????
********************************************************************************************************************************
PROCEDURE textPriemDog
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
SELECT txtOrder
DO CASE
   CASE nuvol=1
        REPLACE txtprn WITH '??????? '+cOrdFio+' ? '+strDateBeg+' ?? '+strDateEnd+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? ??????? ???????? ???????? ??????????.'+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)       
   CASE nuvol=2
        REPLACE txtprn WITH '??????? '+cOrdFio+' ?? ?????? ? '+strDateBeg+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+;
        ', ? ??????? ???????? ???????? ??????????, ?? ?????? ?????????? ????????? ????????? '+ALLTRIM(newFioZam)+', ??????????? ? ?????????? ??????? "?? ????? ?? ???????? ?? ?????????? ?? ???????? 3-? ???".'+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)  
   CASE nuvol=3
        REPLACE txtprn WITH '??????? '+cOrdFio+' ?? ?????? ? '+strDateBeg+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? ??????? ???????? ???????? ??????????.'+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)          
   CASE nuvol=4
        REPLACE txtprn WITH '??????? '+cOrdFio+' ?? ?????? ? ??????? ???????? ???????????????? ? '+strDateBeg+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+;
        ', ? ??????? ???????? ???????? ??????????, ?? ?????? ?????????? ????????? ????????? '+ALLTRIM(newFioZam)+', ??????????? ? ?????????? ??????? "?? ????? ?? ???????? ?? ?????????? ?? ???????? 3-? ???."'+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)        
ENDCASE

       
********************************************************************************************************************************
PROCEDURE saveRecPriemDog
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPriemDog
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,pkont WITH newPkont,ddop WITH newDDop,osnov WITH newOsnov,;
        txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,varSupl WITH IIF(INLIST(nuvol,2,4),PADR(ALLTRIM(STR(newTabZam)),5,' ')+PADR(ALLTRIM(newFioZam),60,' '),''),;
        dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
    
DO saveDimOrd  
DO saveKadrOrder
********************************************************************************************************************************
*       ????? ?? ???????? ????????
********************************************************************************************************************************
PROCEDURE priemperevod
PARAMETERS par1
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
IF !USED('curPerevod')
   SELECT kod,name FROM sprtot WHERE sprtot.kspr=26 INTO CURSOR curPerevod READWRITE && ?????? ??? ????? ?????? ?? ???????? ????????
   SELECT curPerevod
   INDEX ON kod TAG T1
ENDIF  
parPadejDol=3
parPadejFio=1
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newKse=IIF(logNewRec,0.00,curpeoporder.kse)
newTr=IIF(logNewRec,1,curpeoporder.tr)
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
newVdog=IIF(logNewRec,0,curPeopOrder.vdog)
strVdog=IIF(SEEK(newVdog,'curPerevod',1),curPerevod.name,'')
newTabZam=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,1,5)))
newFioZam=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl,6,60))
newPlace=IIF(logNewRec,'',ALLTRIM(curpeoporder.varsupl2))
newpKont=IIF(logNewRec,0,curpeoporder.pkont)
newpDDop=IIF(logNewRec,0,curpeoporder.ddop)
newLogApp=.T.
newNpp=0
objFocus='frmOrd.comboBox22'
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr2',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?? ????????',1,1 
     DO addcombomy WITH 'frmOrd',22,.comboBox1.Left,.ordPodr2.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'strVDog','ALLTRIM(curPerevod.name)',6,'','DO validVidPerevod',.F.,.T.  
     .comboBox22.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordPodr2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'strVDog',.F.,.F.,0  
     .tBoxPodr2.Visible=IIF(par1,.F.,.T.) 
     
     
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordPodr2.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0  
            
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2         
             
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',2,.comboBox1.Left,.ordPodr.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodr',.F.,.T.  
     .comboBox2.Visible=IIF(par1,.T.,.F.)
     .comboBox2.DisplayCount=15
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     .tBoxPodr.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',3,.comboBox1.Left,.ordDol.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDolj',.F.,.T.       
     WITH .comboBox3         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     .tBoxDolj.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)  
     
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
     .tBoxKont.Alignment=0
       
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
     .tBoxDay.Alignment=0
     
     DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'??? (???? ????????)',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioZam',.F.,IIF(par1.AND.newVDog=2,.T.,.F.),0      
     .tBoxFioNew.procforChange='DO changePeopDek'   
     DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectPeopDek',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlntNew.Enabled=IIF(par1.AND.newVDog=2,.T.,.F.)
     
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordFioNew.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPlace',.F.,IIF(par1,.T.,.F.),0 
     .tBoxDoljNew.maxLength=150
                 
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
   
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10'  
     DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
     WITH .listBox2
          .RowSource='curSuplpeop.fior'         
          .RowSourceType=2
          .Visible=.F.     
          .Height=.Parent.Height-.Top   
          .procForDblClick='DO validListDek'
          .procForLostFocus='DO lostFocusPeop2'
          .Enabled=IIF(newVdog=2,.T.,.F.)
     ENDWITH 
           
     IF par1 
        ON ERROR DO erSup     
        .comboBox2.SetFocus
        .comboBox3.SetFocus
        .comboBox11.SetFocus
        .comboBox22.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR 
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE validVidPerevod
newVDog=curPerevod.kod
DO CASE
   CASE newVdog=2
        frmOrd.listBox2.Enabled=.T.
        frmOrd.tBoxFioNew.Enabled=.T.
        frmOrd.butKlntNew.Enabled=.T.
   OTHERWISE 
        frmOrd.listBox2.Enabled=.F.
        frmOrd.tBoxFioNew.Enabled=.F.  
        frmOrd.butKlntNew.Enabled=.F.
ENDCASE
KEYBOARD '{TAB}'
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ??????  ???????? ????????????
********************************************************************************************************************************
PROCEDURE textPriemPerevod
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
SELECT txtOrder
DO CASE
   CASE newVdog=1 
        REPLACE txtprn WITH '??????? '+cOrdFio+' ? '+strDateBeg+' ?? ?????? ?? ???????? ????????? ?? '+strDateEnd+' '+;
        LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? ??????? ???????? ???????? ??????????, ? ??????? ???????? ?? '+ALLTRIM(newPlace)+', ????????? ?????????????? ???? ??????????????:'+CHR(13)+;
        '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
        '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
        '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
        '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
        '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
        '?????????: '+ALLTRIM(newOsnov)
   CASE newVdog=2 
        REPLACE txtprn WITH '??????? '+cOrdFio+' ? '+strDateBeg+' ?? ?????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ?? '+STR(newKse,4,2)+;
                ' ??????? ??????? ??????? ???????? ?? '+ALLTRIM(newPlace)+' ?? ?????? ?????????? ????????? ????????? '+ALLTRIM(newFioZam) +CHR(13)+CHR(13)+'?????????:'+newOsnov+CHR(13)+' ? ???????? '+dim_agree(kodsex)
   CASE newVdog=3 
        REPLACE txtprn WITH '??????? '+cOrdFio+' ? '+strDateBeg+' ?? ?????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ?? '+STR(newKse,4,2)+;
                ' ??????? ??????? ? ??????? ???????? ???????? ?????????? ??????? ???????? ?? '+ALLTRIM(newPlace)+CHR(13)+'?????????:'+newOsnov+CHR(13)+CHR(13)+' ? ???????? '+dim_agree(kodsex)                                
ENDCASE        
********************************************************************************************************************************
PROCEDURE saveRecPriemPerevod
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPriemPerevod	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,pkont WITH newPkont,ddop WITH newddop;
        txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,vdog WITH newVdog,varSupl WITH IIF(newVDog=2,PADR(ALLTRIM(STR(newTabZam)),5,' ')+PADR(ALLTRIM(newFioZam),60,' '),''),;
        varSupl2 WITH newPlace,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex                      
DO saveDimOrd  
DO saveKadrOrder
********************************************************************************************************************************
PROCEDURE procKontZakl
PARAMETERS par1,par2
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
IF !USED('curSrok')
   SELECT kod,name FROM sprtot WHERE sprtot.kspr=7 INTO CURSOR curSrok READWRITE && ?????? ??? ?????? ?????????? ?????????
   SELECT curSrok
   INDEX ON kod TAG T1
ENDIF  
*par2=1  ????????? ????????
*par2=2  ????????? ???????? ?? ????? ??????????? ????????
*par2=2  ???????? ???????? ?? ????? ??????????? ????????
parSave=par2
parPadejDol=3
parPadejFio=3
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newNKont=IIF(logNewRec,SPACE(10),SUBSTR(curPeopOrder.varSupl,1,10))
newdKont=IIF(logNewRec,CTOD('  .  .    '),CTOD(SUBSTR(curPeopOrder.varSupl,11,10)))
strSrok=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl,21,30))
newSrok=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,51,2)))
newPkont=IIF(logNewRec,0,curpeoporder.pkont)
newDdop=IIF(logNewRec,0,curpeoporder.ddop)
objFocus='frmOrd.tBoxNKont'
newLogApp=.T.
logDatJob=.F.
SELECT txtorder
newNpp=0
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0 
     
     
     DO adTboxAsCont WITH 'frmOrd','ordNKont',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'???????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxNKont',.ordNKont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKont',.F.,IIF(par1,.T.,.F.),0      
     DO adTboxAsCont WITH 'frmOrd','ordDKont',.ordPrik.Left,.ordNKont.Top,RetTxtWidth('w???? ?????????w'),dHeight,'???? ?????????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDKont',.ordNKont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newdKont',.F.,IIF(par1,.T.,.F.),0  
     .tBoxNKont.Width=(.comboBox1.Width-.ordDKont.Width)/2
     .ordDKont.Left=.tBoxNKont.left+.tBoxNKont.Width-1
     .tBoxDKont.Left=.ordDKont.Left+.ordDKont.Width-1
     .tBoxDKont.Width=.comboBox1.Width-.tBoxNKont.Width-.ordDKont.Width+2
      
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.ordPrik.Left,.ordNKont.Top+dHeight-1,.ordPrik.Width,dHeight,'?? ????',1,1                                             
     DO addComboMy WITH 'frmOrd',11,.comboBox1.Left,.ordTr.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'strSrok','ALLTRIM(curSrok.name)',6,.F.,'newSrok=curSrok.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strSrok',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)           
         
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
  
   
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
     .tBoxKont.Alignment=0
       
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
     .tBoxDay.Alignment=0
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0 
          
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+dHeight+10' 
     IF par1      
        ON ERROR DO erSup
        .comboBox11.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR 
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE saveRecKontZakl
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textKontZakl	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,osnov WITH newOsnov,;
        txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,pkont WITH newPkont,ddop WITH newDDop,varsupl WITH PADR(ALLTRIM(newNKont),10,' ')+PADR(DTOC(newDKont),10,' ')+;
        PADR(ALLTRIM(strSrok),30,' ')+STR(newSrok,2),nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd  
DO saveKadrOrder      
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ?????? ?? ???????? ?????????
********************************************************************************************************************************
PROCEDURE textKontZakl
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol

SELECT txtOrder
DO CASE
   CASE parSave=1
        REPLACE txtprn WITH '????????? ? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ????? ???????? ???????? ?????? ?? '+ALLTRIM(strSrok)+' ? '+;
        strDateBeg+' ?? '+strDateEnd+', ????????? ?????????????? ???? ??????????????:'+CHR(13)+;
                '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
                '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
                '?????????: '+ALLTRIM(newOsnov)          
    CASE parSave=2   
         REPLACE txtprn WITH '????????? ? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ????? ???????? ???????? ?????? ?? '+ALLTRIM(strSrok)+' ? '+;
                  strDateBeg+' ?? '+strDateEnd+', ?? ????? ??????????? ????????'+CHR(13)+CHR(13)+;                
                 '?????????: '+ALLTRIM(newOsnov)          
    CASE parSave=3 
         REPLACE txtprn WITH '???????? ? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ????? ???????? ???????? ?????? ?? '+ALLTRIM(strSrok)+' ? '+;
                 strDateBeg+' ?? '+strDateEnd+', ?? ????? ??????????? ????????'+CHR(13)+CHR(13)+;                
                '?????????: '+ALLTRIM(newOsnov)   
 ENDCASE     
********************************************************************************************************************************
*                                        ?????????? ? ????? ? ?????????? ????? ?????????
********************************************************************************************************************************
PROCEDURE uvoldog 
PARAMETERS par1,par2
nuvol=par2   && ??????? ??????????
parPadejDol=IIF(nuvol=7,3,1)
parPadejFio=IIF(nuvol=7,3,1)
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newDateUvol=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.dateUvol)
newSLink=ALLTRIM(curSprOrder.slink)
newPerBeg=IIF(logNewRec,CTOD('  .  .    '),curPeopOrder.perBeg)
newPerEnd=IIF(logNewRec,CTOD('  .  .    '),curPeopOrder.perEnd)
newDayKomp=IIF(logNewRec,0,curpeoporder.DayOtp)
newNpp=0
newLogApp=.T.
newPlace=IIF(logNewRec,'',curPeopOrder.varSupl)
objFocus='frmOrd.tBoxBeg'
newKse=IIF(logNewRec,0,curpeoporder.kse)
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)   
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateUvol',.F.,IIF(par1,.T.,.F.),0 
     
 
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPlace',.F.,IIF(par1.AND.nuvol=5,.T.,.F.),0 
 
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxLink',.ordKse.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newsLink',.F.,.F.,0  
     
     DO adTboxAsCont WITH 'frmOrd','ordKomp',.ordPrik.Left,.ordKse.Top+dHeight-1,.ordPrik.Width,dHeight,'???????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxKomp',.ordKomp.Top,.comboBox1.Left,.comboBox1.Width/5,dHeight,'newDayKomp','Z',IIF(par1,.T.,.F.),0 
     .tBoxKomp.InputMask='99' 
     .tBoxKomp.Alignment=0
     DO adTboxAsCont WITH 'frmOrd','ordPerBeg',.tBoxKomp.Left+.tBoxKomp.Width-1,.ordKomp.Top,.tBoxKomp.Width,dHeight,'?????? ?',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxPerBeg',.ordKomp.Top,.ordPerBeg.Left+.ordPerBeg.Width-1,.tBoxKomp.Width,dHeight,'newPerBeg',.F.,IIF(par1,.T.,.F.),0 
     DO adTboxAsCont WITH 'frmOrd','ordPerEnd',.tBoxPerBeg.Left+.tBoxPerBeg.Width-1,.ordKomp.Top,.tBoxKomp.Width,dHeight,'?????? ??',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxPerEnd',.ordKomp.Top,.ordPerEnd.Left+.ordPerEnd.Width-1,.comboBox1.Width-.tBoxKomp.Width*4+4,dHeight,'newPerEnd',.F.,IIF(par1,.T.,.F.),0 
          
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordKomp.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0       
    
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10'
     IF par1   
        ON ERROR DO erSup     
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE saveRecUvolDog
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textUvolDog	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateUvol WITH newDateUvol,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
        dayOtp WITH newDayKomp,perBeg WITH newPerBeg,perEnd WITH newPerEnd,osnov WITH newOsnov,nid WITH newNid,varSupl WITH newPlace,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex        
DO saveDimOrd   
DO saveKadrOrder          
********************************************************************************************************************************
PROCEDURE textUvolDog
strDateUvol=IIF(!EMPTY(newDateUvol),dateToString('newDateUvol',.T.),'') 
strPerBeg=IIF(!EMPTY(newPerBeg),dateToString('newPerBeg',.T.),'') 
strPerEnd=IIF(!EMPTY(newPerEnd),dateToString('newPerEnd',.T.),'') 
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
SELECT txtOrder
DO CASE
   CASE nuvol=1
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ? ????? ? ?????????? ????? ????????? ????????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=2
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ?? ??????? ?????????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=3
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ?? ??????? ??? ???????????? ?????? '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=4
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ? ????? ???????????? ????????? ?? ?????????? ??????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=5                 
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ? ??????? ???????? ? '+ALLTRIM(newPlace)+','+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)       
   CASE nuvol=6        && ?? ?????????? ????????         
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ?????????? ????  ? ?????? '+strDateUvol+;
                ' ? ????? ? ??????????????? ?????????? ????????? ?????????? ????????? ????????, ??????????????? ??????????? ?????? ??????,'+ALLTRIM(newSlink)+CHR(13)+;
                '? ???????? ????????? ??????? ? ??????? ?????????????? ???????? ?????????.'+CHR(13)+;
                '????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)                                   
   CASE nuvol=7        && ? ????? ?? ???????            
        REPLACE txtprn WITH '?????????? ???????? ? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ? ???? ?? ??????? ?????????, '+ALLTRIM(newSlink)+CHR(13)+;
                '????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)   
   CASE nuvol=8        && ?? ????????????????
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+cOrdFio+' '+strDateUvol+' ??????? ? ?????? ?? ???????????????? ?? 0,5 ??????? ??????? (???????????? ?? ????? ), ? ????? ? ?????????? ????? ????????? ????????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)                       
   CASE nuvol=9
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ? ????? ? ???????? ?? ??????? ??????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ?????????:'+CHR(13)+;
                '1. ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                strPerBeg+' ?? '+strPerEnd+'.'+CHR(13)+;
                '2. ???????? ??????? ? ??????? ?????????????? ???????? ?????????, ? ???????????? ?? ??????? 48 ????????? ??????? ?????????? ????????.'+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)                                                
   CASE nuvol=10   
        REPLACE txtprn WITH '??????? '+cOrdFio+', ? ?????? ?? ???????????????? ?? 0,5 ????????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' '+strDateUvol+' ? ????? ? ??????? ?? ?????? ?????????, ??? ???????? ??? ?????? ????? ???????? ????????, '+ALLTRIM(newSlink)+CHR(13)+;    
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)    
   CASE nuvol=11   
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+strDateUvol+', ? ?????? ?? ???????????????? ?? 0,5 ??????? ??????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' '+strDateUvol+' ? ????? ? ???, ?????? ?? ???????????????? ????? ??? ????????? ????????, '+ALLTRIM(newSlink)+CHR(13)+;    
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)                  
   CASE nuvol=12
        REPLACE txtprn WITH '??????? '+cOrdFio+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+;
                 ' ? ????? ?? ??????????? ? ???????? ???? ????????? ????, ??????? ???????? ??????? ? ?????????, ???????????? ??????????? ??????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ?????????? ????????????? ??????.'+CHR(13)+'?????????: '+newOsnov  
   CASE nuvol=13
        REPLACE txtprn WITH '??????? '+cOrdFio+' '+strDateUvol+', ? ?????? ?? ???????????????? ?? 0,5 ????????? '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ?? ??????? ?????????, '+ALLTRIM(newSlink)+CHR(13)+;    
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)      
                 
                                                
ENDCASE                  
*******************************************************************************************************************************
*       ??????? ?? ????? ??????????
********************************************************************************************************************************
PROCEDURE perevodots
PARAMETERS par1,par2
logTxt=par2
parPadejDol=1
parPadejFio=1
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
IF !USED('curSrok')
   SELECT kod,name FROM sprtot WHERE sprtot.kspr=7 INTO CURSOR curSrok READWRITE && ?????? ??? ?????? ?????????? ?????????
   SELECT curSrok
   INDEX ON kod TAG T1
ENDIF  

CREATE CURSOR  curOrdJob1 FROM ARRAY arOrdJob
IF par1
   SELECT curOrdJob
   ZAP
   APPEND FROM datjob FOR INLIST(tr,1,2,3)
   DELETE FOR !EMPTY(dateOut)
   REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fiot,''),ndOrd WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namer,''),npOrd WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'') ALL
   GO TOP 
ENDIF
logperevod=par2
SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newKse=IIF(logNewRec,000.00,curPeopOrder.kse)
newTabZam=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,1,5)))
newFioZam=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl,6,60))

newSKp=IIF(logNewRec,0,curPeoporder.kp2)
newSPodr=IIF(logNewRec,'',curPeoporder.npodr2)
newSKd=IIF(logNewRec,0,curpeoporder.kd2)
newSDolj=IIF(logNewRec,0,curpeoporder.ndolj2)

newKse=IIF(logNewRec,1.00,curpeoporder.kse)
newTr=IIF(logNewRec,1,curpeoporder.tr)
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')

DO CASE
   CASE logPerevod=1
       * newSPodr=IIF(logNewRec,0,SUBSTR(curPeoporder.varSupl2,1,100))
       * newSDolj=IIF(logNewRec,0,ALLTRIM(SUBSTR(curPeoporder.varSupl2,101,100)))
       * newTabZam=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl2,1,5)))
       * newFioZam=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl2,6,60))
   CASE logPerevod=2  
        newNKont=IIF(logNewRec,SPACE(10),SUBSTR(curPeopOrder.varSupl,1,10))
        newdKont=IIF(logNewRec,CTOD('  .  .    '),CTOD(SUBSTR(curPeopOrder.varSupl,11,10)))
        strSrok=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl,21,30))
        newSrok=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,51,2)))
        newPkont=IIF(logNewRec,0,curpeoporder.pkont)
        newDdop=IIF(logNewRec,0,curpeoporder.ddop)  
       
ENDCASE
objFocus='frmOrd.tBoxBeg'
oldNidJob=IIF(logNewRec,0,curPeopOrder.oldnid)
newNidJob=IIF(logNewRec,0,curPeopOrder.nidJob )
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.   
     DO CASE
        CASE logPerevod=1            
             DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'??? (???? ????????)',1,1      
             DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioZam',.F.,IIF(par1,.T.,.F.),0      
             .tBoxFioNew.procforChange='DO changePeopZam'   
             DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
             DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectPeopZam',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
             .butKlntNew.Enabled=IIF(par1,.T.,.F.)          
             DO adTboxAsCont WITH 'frmOrd','ordPodrNew',.ordPrik.Left,.ordFioNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (???)',1,1 
             DO adTboxNew WITH 'frmOrd','tBoxPodrNew',.ordPodrNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSPodr',.F.,IIF(par1,.T.,.F.),0           
             DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordPodrNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (????)',1,1 
             DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSDolj',.F.,IIF(par1,.T.,.F.),0 
             DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
             DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,100
             .spinKse.Enabled=IIF(par1,.T.,.F.)
             DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
             DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
             .comboBox11.Visible=IIF(par1,.T.,.F.)
             DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
             .tBoxTr.Visible=IIF(par1,.F.,.T.) 
                 
        CASE logPerevod=2         
             DO adTboxAsCont WITH 'frmOrd','ordPodrNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (????)',1,1 
             DO addcombomy WITH 'frmOrd',4,.comboBox1.Left,.ordPodrNew.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPerPodr',.F.,.T.  
             .comboBox4.Visible=IIF(par1,.T.,.F.)
             .comboBox4.DisplayCount=17
             DO adTboxNew WITH 'frmOrd','tBoxPodrNew',.ordPodrNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSPodr',.F.,.F.,0  
             .tBoxPodrNew.Visible=IIF(par1,.F.,.T.)   
     
             DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordPodrNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (????)',1,1 
             DO addComboMy WITH 'frmOrd',5,.comboBox1.Left,.ordDolNew.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPerDolj',.F.,.T.       
             WITH .comboBox5         
                  .DisplayCount=15
                  .ColumnCount=3
                  .ColumnWidths='0,50,500'
                  .RowSource="curDolPodr.name,strVac,name"
                  .Visible=IIF(par1,.T.,.F.)
             ENDWITH     
             DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSDolj',.F.,.F.,0  
             .tBoxDoljNew.Visible=IIF(par1,.F.,.T.) 
     
             DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
             DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5           
             .spinKse.Enabled=IIF(par1,.T.,.F.)   
             DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
             DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
            .comboBox11.Visible=IIF(par1,.T.,.F.)
             DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
            .tBoxTr.Visible=IIF(par1,.F.,.T.)  
        
         
             DO adTboxAsCont WITH 'frmOrd','ordNKont',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'???????? ?',1,1   
             DO adTboxNew WITH 'frmOrd','tBoxNKont',.ordNKont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKont',.F.,IIF(par1,.T.,.F.),0      
             DO adTboxAsCont WITH 'frmOrd','ordDKont',.ordPrik.Left,.ordNKont.Top,RetTxtWidth('w???? ?????????w'),dHeight,'???? ?????????',1,1   
             DO adTboxNew WITH 'frmOrd','tBoxDKont',.ordNKont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newdKont',.F.,IIF(par1,.T.,.F.),0  
             .tBoxNKont.Width=(.comboBox1.Width-.ordDKont.Width)/2
             .ordDKont.Left=.tBoxNKont.left+.tBoxNKont.Width-1
             .tBoxDKont.Left=.ordDKont.Left+.ordDKont.Width-1
             .tBoxDKont.Width=.comboBox1.Width-.tBoxNKont.Width-.ordDKont.Width+2      
             DO adTBoxAsCont WITH 'frmOrd','ordKomp',.ordPrik.Left,.ordNKont.Top+dHeight-1,.ordPrik.Width,dHeight,'?? ????',1,1                                             
             DO addComboMy WITH 'frmOrd',3,.comboBox1.Left,.ordKomp.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'strSrok','ALLTRIM(curSrok.name)',6,.F.,'newSrok=curSrok.kod',.F.,.T. 
             .comboBox3.Visible=IIF(par1,.T.,.F.)
             DO adTboxNew WITH 'frmOrd','tBoxLink',.ordKomp.Top,.comboBox3.Left,.comboBox3.Width,dHeight,'strSrok',.F.,.F.,0  
             .tBoxLink.Visible=IIF(par1,.F.,.T.)                    
             DO adTboxAsCont WITH 'frmOrd','ordDol2',.ordPrik.Left,.ordKomp.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
             DO adTboxNew WITH 'frmOrd','tBoxDolj2',.ordDol2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
             DO adTboxAsCont WITH 'frmOrd','ordDol3',.ordPrik.Left,.ordDol2.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
             DO adTboxNew WITH 'frmOrd','tBoxDolj3',.ordDol3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
             .tBoxDolj2.Width=(.comboBox1.Width-.ordDol2.Width)/2
             .ordDol3.Left=.tBoxDolj2.left+.tBoxDolj2.Width-1
             .tBoxDolj3.Left=.ordDol3.Left+.ordDol3.Width-1
             .tBoxDolj3.Width=.comboBox1.Width-.tBoxDolj2.Width-.ordDol3.Width+2
  

             DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordDol2.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
             DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
             .tBoxKont.Alignment=0
       
             DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
             DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
             .tBoxDay.Alignment=0 
                      
     ENDCASE  
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,IIF(logPerevod=1,.ordKse.Top,.ordDay.Top)+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
     IF logPerevod=1
        DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
        WITH .listBox2
             .RowSource='curOrdJob1.fio,ndord'               
             .RowSourceType=2
             .ColumnCount=2 
             .columnWidths='250,400'        
             .Visible=.F.    
             .Height=.Parent.Height-.Top    
             .procForDblClick='DO validListZam'
             .procForLostFocus='DO lostFocusPeop2'
        ENDWITH  
     ENDIF            
     IF par1           
        IF logPerevod=2
           .comboBox3.SetFocus
           .comboBox4.SetFocus
           .comboBox5.SetFocus
           .comboBox11.SetFocus
        ENDIF
        ON ERROR DO erSup
        .tBoxFio.SetFocus
        ON ERROR    
            
     ENDIF
ENDWITH
*************************************************************************************************************************
PROCEDURE textPerevodOts
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')  
*strPerEnd=IIF(!EMPTY(newPerEnd),dateToString('newPerEnd',.T.),'') 

cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
SELECT txtOrder
DO CASE 
   CASE logperevod=1 && ?? ????? ??????????
        REPLACE txtprn WITH '????????? '+ALLTRIM(cOrdFio)+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', c '+strDateBeg+' '+LOWER(ALLTRIM(newSDolj))+' '+LOWER(ALLTRIM(newSPodr))+;
                ' ?? '+STR(newKse,4,2)+' ?????????, ? ??????? ???????? ???????? ??????????, ?? ????? ?????????? ????????? ????????? '+ALLTRIM(newFioZam)+;
                ' ??????????? ? ?????????? ??????? "?? ????? ?? ???????? ?? ?????????? ?? ???????? 3-? ???".'+CHR(13)+CHR(13)+;
                '?????????:'+newOsnov+CHR(13)+' ? ???????? '+dim_agree(kodsex)
   CASE logperevod=2  && ?? ????? ?????????? ? ??????????? ????????? 
        strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')  
        REPLACE txtprn WITH '????????? '+ALLTRIM(cOrdFio)+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ?????????? ?? ????????? ????????, ???????????? ?? ????? ?????????? ????????? ????????? '+;
                ' ?? ????????? ???????? c??????? ???????? ???????? ?????????? ? ????????? ? ??? ???????? ?????? ?? '+ALLTRIM(strSrok)+' ? '+;
                 strDateBeg+' ?? '+strDateEnd+', ????????? ?????????????? ???? ??????????????:'+CHR(13)+;
                '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
                '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
                '?????????:'+newOsnov+CHR(13)+' ? ???????? '+dim_agree(kodsex)
ENDCASE        
********************************************************************************************************************************
PROCEDURE saveRecPerevodOts
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPerevodOts	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
DO CASE
   CASE logPerevod=1
        REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
                kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,;
                txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,nidjob WITH newNidJob,dord WITH repdorder,nord WITH repnorder    
                REPLACE varsupl2 WITH PADR(ALLTRIM(STR(newSKp)),3,' ')+PADR(ALLTRIM(newSpodr),100,' ')+PADR(ALLTRIM(STR(newSKd)),3,' ')+PADR(ALLTRIM(newSdolj),100,' '),varSupl WITH PADR(ALLTRIM(STR(newTabZam)),5,' ')+PADR(ALLTRIM(newFioZam),60,' '),sex WITH kodsex                 
   CASE logPerevod=2
        REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd;
                kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,;
                kp2 WITH newSkp,npodr2 WITH newSPodr,kd2 WITH newSkd,ndolj2 WITH newSDolj,;
                txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,nidjob WITH newNidJob,pkont WITH newPkont,ddop WITH newDDop,oldnid WITH oldnidjob;
                varsupl WITH PADR(ALLTRIM(newNKont),10,' ')+PADR(DTOC(newDKont),10,' ')+PADR(ALLTRIM(strSrok),30,' ')+STR(newSrok,2),dord WITH repdorder,nord WITH repnorder,sex WITH kodsex                     
ENDCASE        
DO saveDimOrd  
DO saveKadrOrder

*******************************************************************************************************************************
*                                                 ??????? 
********************************************************************************************************************************
PROCEDURE perevod
PARAMETERS par1,par2
logTxt=par2
parPadejDol=1
parPadejFio=1
CREATE CURSOR  curOrdJob1 FROM ARRAY arOrdJob
IF par1
   SELECT curOrdJob
   ZAP
   APPEND FROM datjob FOR INLIST(tr,1,2,3)
   DELETE FOR !EMPTY(dateOut)
   REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fiot,''),ndOrd WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namer,''),npOrd WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'') ALL
   GO TOP 
ENDIF
logperevod=par2
SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newKse=IIF(logNewRec,000.00,curPeopOrder.kse)
newFioZam=IIF(logNewRec,SPACE(50),ALLTRIM(SUBSTR(curpeoporder.varsupl,1,60)))
newKse=IIF(logNewRec,1.00,curpeoporder.kse)
newPkont=IIF(logNewRec,0,curPeopOrder.pkont)
newDDop=IIF(logNewRec,0,curPeopOrder.ddop)
newNidJob=IIF(logNewRec,0,curPeopOrder.nidjob)
newSkp=IIF(logNewRec,0,curPeopOrder.nkp) 
newSkd=IIF(logNewRec,0,curPeopOrder.nkd)
newSPodr=IIF(logNewRec,0,SUBSTR(curPeoporder.varSupl2,1,100))
newSDolj=IIF(logNewRec,0,ALLTRIM(SUBSTR(curPeoporder.varSupl2,101,100)))

DO CASE  
   CASE logPerevod=4
       
ENDCASE
objFocus='frmOrd.tBoxBeg'
oldNidJob=IIF(logNewRec,0,curpeoporder.oldnid)
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     
     
     DO adTboxAsCont WITH 'frmOrd','ordPodrNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (????)',1,1 
     DO addcombomy WITH 'frmOrd',4,.comboBox1.Left,.ordPodrNew.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSPodr','ALLTRIM(curSprPodr.name)',6,'','DO procPerPodr',.F.,.T.  
     .comboBox4.Visible=IIF(par1,.T.,.F.)
     .comboBox4.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodrNew',.ordPodrNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSPodr',.F.,.F.,0  
     .tBoxPodrNew.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordPodrNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (????)',1,1 
     DO addComboMy WITH 'frmOrd',5,.comboBox1.Left,.ordDolNew.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPerDolj',.F.,.T.       
     WITH .comboBox5         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH     
     DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSDolj',.F.,.F.,0  
     .tBoxDoljNew.Visible=IIF(par1,.F.,.T.)  
     
     
     
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5           
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)  
         
     DO CASE   
        CASE logPerevod=3
             DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordKse.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
        CASE logPerevod=4             
             DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordKse.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
             DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
             .tBoxKont.Alignment=0
       
             DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
             DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
             .tBoxDay.Alignment=0
             DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     ENDCASE        
     
    DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
     IF logPerevod=1
        DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
        WITH .listBox2
             .RowSource='curOrdJob1.fio,ndord'               
             .RowSourceType=2
             .ColumnCount=2 
             .columnWidths='250,400'    
             .Height=.Parent.Height-.Top    
             .Visible=.F.        
             .procForDblClick='DO validListZam'
             .procForLostFocus='DO lostFocusPeop2'
        ENDWITH 
     ENDIF            
     IF par1 
        ON ERROR DO erSup          
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
*******************************************************************************************************************************
*                                                 ??????? ?????????????
********************************************************************************************************************************
PROCEDURE perevodmany
PARAMETERS par1,par2
logTxt=par2
parPadejDol=1
parPadejFio=1
CREATE CURSOR  curOrdJob1 FROM ARRAY arOrdJob
IF par1
   SELECT curOrdJob
   ZAP
   APPEND FROM datjob FOR INLIST(tr,1,2,3)
   DELETE FOR !EMPTY(dateOut)
   REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fiot,''),ndOrd WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namer,''),npOrd WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'') ALL
   GO TOP 
ENDIF
logperevod=par2
SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newKse=IIF(logNewRec,000.00,curPeopOrder.kse)

PUBLIC newKp2,newNpodr2,newKd2,newNDolj2,newKse2,newKp3,newNpodr3,newKd3,newNDolj3,newKse3,newKp4,newNpodr4,newKd4,newNDolj4,newKse4
newKp2=IIF(logNewRec,0,curpeoporder.kp2)
newNPodr2=IIF(logNewRec,'',curpeoporder.npodr2)
newKd2=IIF(logNewRec,0,curpeoporder.kd2)
newNDolj2=IIF(logNewRec,'',curpeoporder.ndolj2)
newKse2=IIF(logNewRec,0.00,curpeoporder.kse2)

newKp3=IIF(logNewRec,0,curpeoporder.kp3)
newNPodr3=IIF(logNewRec,'',curpeoporder.npodr3)
newKd3=IIF(logNewRec,0,curpeoporder.kd3)
newNDolj3=IIF(logNewRec,'',curpeoporder.ndolj3)
newKse3=IIF(logNewRec,0.00,curpeoporder.kse3)

newKp4=IIF(logNewRec,0,curpeoporder.kp4)
newNPodr4=IIF(logNewRec,'',curpeoporder.npodr4)
newKd4=IIF(logNewRec,0,curpeoporder.kd4)
newNDolj4=IIF(logNewRec,'',curpeoporder.ndolj4)
newKse4=IIF(logNewRec,0.00,curpeoporder.kse4)

newPkont=IIF(logNewRec,0,curPeopOrder.pkont)
newDDop=IIF(logNewRec,0,curPeopOrder.ddop)
newNidJob=IIF(logNewRec,0,curPeopOrder.nidjob)

oldNidJob=IIF(logNewRec,0,curpeoporder.oldnid)
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     
     
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.ordPodr.Left,.ordBeg.Top+dHeight-1,.comboBox1.Width,dHeight,'??? ??????',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.comboBox1.Left,.ordTr.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)
     
  *******2
     
     DO adTboxAsCont WITH 'frmOrd','ordPodr2',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (????)',1,1 
     DO addcombomy WITH 'frmOrd',22,.comboBox1.Left,.ordPodr2.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr2','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodrMany WITH 2',.F.,.T.  
     .comboBox22.Visible=IIF(par1,.T.,.F.)
     .comboBox22.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordPodr2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr2',.F.,.F.,0  
     .tBoxPodr2.Visible=IIF(par1,.F.,.T.) 
     
     
     DO adTboxAsCont WITH 'frmOrd','ordDol2',.ordPrik.Left,.ordPodr2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',32,.comboBox1.Left,.ordDol2.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj2','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDoljMany WITH 2',.F.,.T.       
     WITH .comboBox32         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj2',.ordDol2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj2',.F.,.F.,0  
     .tBoxDolj2.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse2',.ordPrik.Left,.ordDol2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse2',.comboBox1.Left,.ordKse2.Top,dheight,RetTxtWidth('9999999999'),'newKse2',0.25,.F.,0,1.5 
     .spinKse2.Enabled=IIF(par1,.T.,.F.)   
     
     
     *****3
     DO adTboxAsCont WITH 'frmOrd','ordPodr3',.ordPrik.Left,.ordKse2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',23,.comboBox1.Left,.ordPodr3.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr3','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodrMany WITH 3',.F.,.T.  
     .comboBox23.Visible=IIF(par1,.T.,.F.)
     .comboBox23.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr3',.ordPodr3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr3',.F.,.F.,0  
     .tBoxPodr3.Visible=IIF(par1,.F.,.T.) 
          
     DO adTboxAsCont WITH 'frmOrd','ordDol3',.ordPrik.Left,.ordPodr3.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',33,.comboBox1.Left,.ordDol3.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj3','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDoljMany WITH 3',.F.,.T.       
     WITH .comboBox33         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj3',.ordDol3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj3',.F.,.F.,0  
     .tBoxDolj3.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse3',.ordPrik.Left,.ordDol3.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse3',.comboBox1.Left,.ordKse3.Top,dheight,RetTxtWidth('9999999999'),'newKse3',0.25,.F.,0,1.5 
     .spinKse3.Enabled=IIF(par1,.T.,.F.)   
               
      *****43
     DO adTboxAsCont WITH 'frmOrd','ordPodr4',.ordPrik.Left,.ordKse3.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO addcombomy WITH 'frmOrd',24,.comboBox1.Left,.ordPodr4.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNPodr4','ALLTRIM(curSprPodr.name)',6,'','DO procPriemPodrMany WITH 4',.F.,.T.  
     .comboBox24.Visible=IIF(par1,.T.,.F.)
     .comboBox24.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodr4',.ordPodr4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr4',.F.,.F.,0  
     .tBoxPodr4.Visible=IIF(par1,.F.,.T.) 
          
     DO adTboxAsCont WITH 'frmOrd','ordDol4',.ordPrik.Left,.ordPodr4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO addComboMy WITH 'frmOrd',34,.comboBox1.Left,.ordDol4.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNDolj4','ALLTRIM(curDolPodr.name)',6,.F.,'DO procPriemDoljMany WITH 4',.F.,.T.       
     WITH .comboBox34         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH           
     DO adTboxNew WITH 'frmOrd','tBoxDolj4',.ordDol4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj4',.F.,.F.,0  
     .tBoxDolj4.Visible=IIF(par1,.F.,.T.)   
     DO adTboxAsCont WITH 'frmOrd','ordKse4',.ordPrik.Left,.ordDol4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse4',.comboBox1.Left,.ordKse4.Top,dheight,RetTxtWidth('9999999999'),'newKse4',0.25,.F.,0,1.5 
     .spinKse4.Enabled=IIF(par1,.T.,.F.)   
             
          
     **?????????             
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordKse4.Top+dHeight-1,.ordPrik.Width,dHeight,'% ?? ????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPkont','Z',IIF(par1,.T.,.F.),0,'999'     
     .tBoxKont.Alignment=0
       
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'?????.??????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxDay',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop','Z',IIF(par1,.T.,.F.),0,'99'       
     .tBoxDay.Alignment=0        
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0 
         
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
             
     IF par1       
        ON ERROR DO erSup    
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE procPerPodr
newSKp=curSprPodr.kod  
newSPodr=curSprPodr.name 
SELECT curDolPodr
SET FILTER TO kp=newSKp
SCAN ALL
     ksesup=0 
     DO sumvackse
     SELECT curDolPodr
     *REPLACE strVac WITH IIF(kse-ksesup=0,'',STR(kse-ksesup,6,2))
ENDSCAN
frmOrd.ComboBox5.RowSource='curDolPodR.name'
WITH .comboBox5         
     .DisplayCount=15
     .ColumnCount=3
     .ColumnWidths='0,50,500'
     .RowSource="curDolPodr.name,strVac,name"
ENDWITH 
frmOrd.ComboBox5.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
frmOrd.ComboBox5.RowSourceType=6
frmOrd.comboBox5.ProcForValid='DO procPerDolj'
KEYBOARD '{TAB}'

********************************************************************************************************************************
PROCEDURE procPerDolj
newSKd=curDolPodr.kd  
newSDolj=curDolPodr.name
frmOrd.ComboBox5.ControlSource='newSDolj'
frmOrd.comboBox5.Refresh
KEYBOARD '{TAB}'
*************************************************************************************************************************
PROCEDURE textPerevod
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')  
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol


cFioZam=''
cPodrZam=''
cDolZam=''
DO procOrdNameDolZam WITH 'newFioZam','newSDolj','newSPodr',2,3


SELECT txtOrder
DO CASE 
   
   CASE logperevod=3 && ?? ?????? ??????
        REPLACE txtprn WITH '????????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', c '+strDateBeg+' '+LOWER(cDolzam)+' '+LOWER(cPodrZam)+;
               ', ? ??????? ???????? ???????? ??????????.'+CHR(13)+CHR(13)+'?????????:'+newOsnov+CHR(13)+' ? ???????? '+dim_agree(kodsex)
   CASE logperevod=4  && ?? ???????? ????????? 
        strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')  
        REPLACE txtprn WITH '????????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ? '+;
                 strDateBeg+' ?? ???????? ????????? ?? '+strDateEnd+' '+LOWER(cDolZam)+' '+LOWER(cPodrZam)+', ? ??????? ???????? ???????? ??????????, ????????? ?????????????? ???? ??????????????:'+CHR(13)+;
                '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
                '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
                '?????????:'+newOsnov+CHR(13)+CHR(13)+' ? ???????? '+dim_agree(kodsex) 
   CASE logPerevod=10   
        strmany=' (?? ??? '      
        IF newKse2#0
           strmany=strmany+STR(newKse2,4,2)+' ????????? '+LOWER(ALLTRIM(newNDolj2))+' '+LOWER(ALLTRIM(newNPodr2))+' '
        ENDIF
        IF newKse3#0
           strmany=strmany+STR(newKse3,4,2)+' ????????? '+LOWER(ALLTRIM(newNDolj3))+' '+LOWER(ALLTRIM(newNPodr3))+' '
        ENDIF
        IF newKse4#0
           strmany=strmany+STR(newKse4,4,2)+' ????????? '+LOWER(ALLTRIM(newNDolj4))+' '+LOWER(ALLTRIM(newNPodr4))
        ENDIF
        strmany=strmany+')'
        SELECT txtOrder
        REPLACE txtprn WITH '????????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ? '+;
                strDateBeg+' ?? ???????? ????????? ?? '+strDateEnd+' '+ALLTRIM(LOWER(newNDolj2))+' '+ALLTRIM(LOWER(newNpodr2))+strMany+', ? ??????? ???????? ???????? ??????????, ????????? ?????????????? ???? ??????????????:'+CHR(13)+;       
                '     ???????? ?? ?????? ?? ???????? ????????? - '+LTRIM(STR(newPkont))+' %;'+CHR(13)+;
                '     ?????????????? ??????????????? ?????????????? ????????? ??????? - '+LTRIM(STR(newDDop))+' ??????????? ????.'+CHR(13)+CHR(13)+;
                '     ? ???????? '+dim_agree(kodsex)+CHR(13)+;
                '     ? ????????? ??????????? ????????? ?????????? ? '+CHR(13)+;
                '     ???????????? ????????? '+dim_agree(kodsex)+CHR(13)+CHR(13)+;
                '?????????: '+ALLTRIM(newOsnov)                             
ENDCASE        
********************************************************************************************************************************
PROCEDURE saveRecPerevod
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textPerevod	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
DO CASE
   CASE logPerevod=3 
        REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
                kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,;
                txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,nidjob WITH newNidJob,nkp WITH newSkp,nkd WITH newSkd,oldnid WITH oldnidjob    
                REPLACE varSupl2 WITH PADR(ALLTRIM(newSpodr),100,' ')+PADR(ALLTRIM(newSdolj),100,' '),dord WITH repdorder,nord WITH repnorder,sex WITH kodsex 
   CASE logPerevod=4 
        REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
                kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,;
                txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,nidjob WITH newNidJob,nkp WITH newSkp,nkd WITH newSkd,;
                pkont WITH newPkont,ddop WITH newDDop,oldnid WITH oldnidjob,sex WITH kodsex   
                REPLACE varSupl2 WITH PADR(ALLTRIM(newSpodr),100,' ')+PADR(ALLTRIM(newSdolj),100,' '),dord WITH repdorder,nord WITH repnorder
   CASE logPerevod=10
        REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd;
                kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,;
                txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,nid WITH newNid,nidjob WITH newNidJob,pkont WITH newPkont,ddop WITH newDDop,oldnid WITH oldnidjob,;
                kp2 WITH newKp2,npodr2 WITH newNPodr2,kd2 WITH newKd2,ndolj2 WITH newNDolj2,kse2 WITH newKse2,;
                kp3 WITH newKp3,npodr3 WITH newNPodr3,kd3 WITH newKd3,ndolj3 WITH newNDolj3,kse3 WITH newKse3,;
                kp4 WITH newKp4,npodr4 WITH newNPodr4,kd4 WITH newKd4,ndolj4 WITH newNDolj4,kse4 WITH newKse4,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex                    
ENDCASE        
DO saveDimOrd  
DO saveKadrOrder
********************************************************************************************************************************
*                                        ???????????? ???????
********************************************************************************************************************************
PROCEDURE procTime
PARAMETERS par1,par2
nuvol=par2   && 
parPadejDol=2
parPadejFio=2
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newKodPrik=IIF(!par1,curpeoporder.supOrd,newKodprik)

newNpp=0
newLogApp=.T.
objFocus='frmOrd.tBoxBeg'
newNidJob=IIF(logNewRec,0,curPeopOrder.nidjob)
oldNidJob=IIF(logNewRec,0,curPeopOrder.oldnid)
newKse=IIF(logNewRec,IIF(nuvol=1,1.00,0.00),curPeopOrder.kse)
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)   
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0  
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0  
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0 
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5           
     .spinKse.Enabled=IIF(par1,.T.,.F.) 
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.) 
         
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordKse.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0       
        
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10'
     IF par1     
        ON ERROR DO erSup   
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE saveRecTime
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textTime	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
        osnov WITH newOsnov,nid WITH newNid,oldnid WITH oldnidjob,nidJob WITH newNidJob,kse WITH newKse,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd   
DO saveKadrOrder          
********************************************************************************************************************************
PROCEDURE textTime
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'') 
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
SELECT txtOrder
DO CASE
   CASE nuvol=1
        REPLACE txtprn WITH '?????????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ? '+strDateBeg+;
                ' ?????? ??????? ????? c ??????? ???????? ???????? ??????????. '+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=2
         REPLACE txtprn WITH '?????????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ? '+strDateBeg+;
                ' ???????? ??????? ????? '+' ?? ??????? '+LTRIM(STR(newKse,4,2)) +' ?????????, c ??????? ???????? ???????? ??????????. '+CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)      
   CASE nuvol=3
        REPLACE txtprn WITH '??????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' '+strDateUvol+' ?? ??????? ??? ???????????? ?????? '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=4
        REPLACE txtprn WITH '??????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' '+strDateUvol+' ? ????? ???????????? ????????? ?? ?????????? ??????, '+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)        
   CASE nuvol=5                 
        REPLACE txtprn WITH '??????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' '+strDateUvol+' ? ??????? ???????? ? '+ALLTRIM(newPlace)+','+ALLTRIM(newSlink)+CHR(13)+;
                '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)       
   CASE nuvol=6        && ?? ?????????? ????????         
        REPLACE txtprn WITH '??????? '+cOrdFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', ?????????? ????  ? ?????? '+strDateUvol+;
                ' ? ????? ? ??????????????? ?????????? ????????? ?????????? ????????? ????????, ??????????????? ??????????? ?????? ??????,'+ALLTRIM(newSlink)+CHR(13)+;
                '? ???????? ????????? ??????? ? ??????? ?????????????? ???????? ?????????.'+CHR(13)+;
                '????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)                                   
   CASE nuvol=7        && ? ????? ?? ???????     
   
        REPLACE txtprn WITH '?????????? ???????? ? '+ordFio+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+', '+strDateUvol+' ? ???? ?? ??????? ?????????, '+ALLTRIM(newSlink)+', ? ????? ?? ??????? ??????????.'+CHR(13)+;
                '????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+ CHR(13)+CHR(13)+'?????????: '+newOsnov+CHR(13)                            
ENDCASE        
*************************************************************************************************************************        
PROCEDURE saveKadrOrder
SELECT datJob
oldOrdJob=SYS(21)
SELECT peoporder
SEEK curpeoporder.nid
SELECT curPeopOrder
oldRec=RECNO()
IF logApp     
   SELECT datJob
   DO CASE
      CASE INLIST(curPeopOrder.supOrd,1,20,21,22,23)    && ????? ?? ???????? ?????????, ????????? ????????
           IF SEEK(curPeopOrder.kodpeop,'people',1)
              SELECT people
              REPLACE pkont WITH curPeopOrder.pkont,date_in WITH curPeopOrder.dateBeg,dayotp WITH 24,daykont WITH curPeopOrder.ddop,totday WITH people.dayotp+people.daykont,;
                      begdog WITH curPeopOrder.dateBeg,enddog WITH curPeopOrder.dateEnd,dordin WITH datorder.dateOrd,nordin WITH ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)  
              DO CASE
                 CASE curpeoporder.supOrd=1
                      REPLACE dog WITH 1
                 CASE curpeoporder.supord=20
                      REPLACE dog WITH 3
                 OTHERWISE
                      REPLACE dog WITH 2,enddog WITH CTOD('  .  .    ')
              ENDCASE                    
              SELECT datjob        
           ENDIF                     
           IF !SEEK(curPeopOrder.nidjob,'datjob',7) 
              SET DELETED OFF 
              SET ORDER TO 7
              GO BOTTOM
              newNidch=nid+1
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid
                      IF INLIST(curpeoporder.supOrd,21,23)
                         REPLACE kdek WITH VAL(SUBSTR(curpeoporder.varSupl,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl,6,60)
                      ENDIF 
              REPLACE curpeoporder.nidjob WITH newNidch,peoporder.nidjob WITH newNidch  
           ELSE 
              REPLACE kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,kdek WITH 0,fiodek WITH ''
                      IF INLIST(curpeoporder.supOrd,21,23)
                         REPLACE kdek WITH VAL(SUBSTR(curpeoporder.varSupl,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl,6,60)
                      ENDIF                                               
           ENDIF
           DO repnadjob
      CASE curPeopOrder.supOrd=2    && ????? ???????? ???????????????? 
           IF SEEK(curPeopOrder.kodpeop,'people',1)
              SELECT people
              REPLACE date_in WITH curPeopOrder.dateBeg,dayotp WITH 24,daykont WITH curPeopOrder.ddop,totday WITH people.dayotp+people.daykont,;
                      begdog WITH curPeopOrder.dateBeg,enddog WITH curPeopOrder.dateEnd,dog WITH IIF(curpeoporder.vdog=1,3,2),dordin WITH datorder.dateOrd,nordin WITH ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)
              SELECT datjob        
           ENDIF          
           IF !SEEK(curPeopOrder.kodpeop,'datjob',1) 
              SET DELETED OFF 
              SET ORDER TO 7
              GO BOTTOM
              newNidCh=nid+1
              
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidCh, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid
              REPLACE curpeoporder.nidjob WITH newNidch,peoporder.nidjob WITH newNidch                                                                                                 
           ELSE
              REPLACE kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid
           ENDIF
           IF SEEK(curPeopOrder.kodpeop,'people',1)
              SELECT people
              REPLACE lvn WITH .T.,date_in WITH curPeopOrder.dateBeg
              SELECT datjob
              IF curpeoporder.vDog=2
                 REPLACE kdek WITH VAL(SUBSTR(curpeoporder.varSupl,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl,6,60)
              ENDIF   
           ENDIF
           DO repnadjob
      CASE curPeopOrder.supOrd=3      &&??????? ? ??????????? ?????????        
           newNidCh=curpeoporder.nidjob
           SELECT datjob
           SET ORDER TO 7
           SEEK oldNidJob
           REPLACE dateOut WITH newDateBeg,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid                       
           IF curPeopOrder.nidjob=0.OR.!SEEK(curPeopOrder.nidJob,'datjob',7)
              SET DELETED OFF 
              GO BOTTOM
              newNidCh=nid+1              
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop             
           ELSE           
              SEEK(curPeopOrder.nidJob)           
           ENDIF          
           REPLACE kp WITH curPeopOrder.kp2,kd WITH curPeopOrder.kd2,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                   dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                   kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
           DO repnadjob
                   
           REPLACE curpeoporder.nidjob WITH newNidch,peoporder.nidjob WITH newNidch         
           SELECT people
           oldPeopOrd=SYS(21)
           SET ORDER TO 1
           SEEK curPeopOrder.kodpeop
                REPLACE dog WITH 1,pkont WITH curpeoporder.pkont,daykont WITH curPeopOrder.ddop,numdog WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,1,10)),ALLTRIM(SUBSTR(curPeopOrder.varSupl,1,10)),people.numdog),;
                        strtime WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,21,30)),ALLTRIM(SUBSTR(curPeopOrder.varSupl,21,30)),people.strtime),begDog WITH curPeopOrder.dateBeg,enddog WITH curPeopOrder.dateEnd                                 
           ON ERROR DO erSup
           REPLACE ddog WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,11,10)),CTOD(SUBSTR(curPeopOrder.varSupl,11,10)),people.ddog)                                 
           REPLACE ktime WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,51,2)),VAL(SUBSTR(curPeopOrder.varSupl,51,20)),people.ktime)                                 
           ON ERROR             
           SET ORDER TO &oldPeopOrd 
      CASE curPeopOrder.supOrd=4      && ??????? ?? ????? ?????????  
           SELECT datjob
           SET ORDER TO 7
           SEEK oldNidJob
           REPLACE dateOut WITH newDateBeg,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid                 
           IF curPeopOrder.nidjob=0.OR.!SEEK(curPeopOrder.nidjob,'datjob',7)
              SET DELETED OFF 
              GO BOTTOM
              newNidCh=nid+1              
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop
           ELSE
              SEEK(curPeopOrder.nidJob) 
           ENDIF                   
           REPLACE kp WITH VAL(SUBSTR(curPeopOrder.varsupl2,1,3)),kd WITH VAL(SUBSTR(curPeopOrder.varsupl2,104,3)),kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                   dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                   kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,kdek WITH VAL(SUBSTR(curpeoporder.varSupl,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl,6,60),;
                   lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
           DO repnadjob        
           REPLACE curpeoporder.nidjob WITH datjob.nid,peoporder.nidjob WITH datjob.nid           
                    
      CASE curpeopOrder.supOrd=8   && ????? ?? ???????? ????????? ?????????????
           IF SEEK(curPeopOrder.kodpeop,'people',1)
              SELECT people 
              REPLACE pkont WITH curPeopOrder.pkont,date_in WITH curPeopOrder.dateBeg,dayotp WITH 24,daykont WITH curPeopOrder.ddop,;
                      totday WITH dayotp+daykont,dordin WITH datorder.dateOrd,nordin WITH ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)
              SELECT datjob       
           ENDIF 
           IF !SEEK(curPeopOrder.nidjob,'datjob',7) 
              SET DELETED OFF 
              SET ORDER TO 7
              GO BOTTOM
              newNidch=nid+1
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.) 
              DO repnadjob        
              REPLACE curpeoporder.nidjob WITH newNidch,peoporder.nidjob WITH newNidch                     
           ELSE 
              REPLACE kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                       kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
              DO repnadjob                 
           ENDIF           
           IF curpeoporder.kp2#0.AND.curpeoporder.kd2#0
              IF !SEEK(curPeopOrder.nidjob2,'datjob',7) 
                 SET DELETED OFF 
                 SET ORDER TO 7
                 GO BOTTOM
                 newNidch=nid+1
                 SET DELETED ON 
                 APPEND BLANK
                 REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp2,kd WITH curPeopOrder.kd2,kse WITH curPeopOrder.kse2,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                 DO repnadjob                
                 REPLACE curpeoporder.nidjob2 WITH newNidch,peoporder.nidjob2 WITH newNidch    
              ELSE 
                 REPLACE kp WITH curPeopOrder.kp2,kd WITH curPeopOrder.kd2,kse WITH curPeopOrder.kse2,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)                                                         
                 DO repnadjob                
              ENDIF 
           ENDIF  
           
           IF curpeoporder.kp3#0.AND.curpeoporder.kd3#0
              IF !SEEK(curPeopOrder.nidjob3,'datjob',7) 
                 SET DELETED OFF 
                 SET ORDER TO 7
                 GO BOTTOM
                 newNidch=nid+1
                 SET DELETED ON 
                 APPEND BLANK
                 REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp3,kd WITH curPeopOrder.kd3,kse WITH curPeopOrder.kse3,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                 DO repnadjob                
                 REPLACE curpeoporder.nidjob3 WITH newNidch,peoporder.nidjob3 WITH newNidch    
              ELSE 
                 REPLACE kp WITH curPeopOrder.kp3,kd WITH curPeopOrder.kd3,kse WITH curPeopOrder.kse3,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)                                                         
                 DO repnadjob                
              ENDIF 
           ENDIF  
             
           IF curpeoporder.kp4#0.AND.curpeoporder.kd4#0
              IF !SEEK(curPeopOrder.nidjob4,'datjob',7) 
                 SET DELETED OFF 
                 SET ORDER TO 7
                 GO BOTTOM
                 newNidch=nid+1
                 SET DELETED ON 
                 APPEND BLANK
                 REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp4,kd WITH curPeopOrder.kd4,kse WITH curPeopOrder.kse4,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                 DO repnadjob                
                 REPLACE curpeoporder.nidjob4 WITH newNidch,peoporder.nidjob4 WITH newNidch    
              ELSE 
                 REPLACE kp WITH curPeopOrder.kp4,kd WITH curPeopOrder.kd4,kse WITH curPeopOrder.kse4,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid ,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)                                                        
                 DO repnadjob                
              ENDIF 
           ENDIF   
      CASE curpeopOrder.supOrd=9 
           IF SEEK(curPeopOrder.kodpeop,'people',1)
              SELECT people
              REPLACE pkont WITH curPeopOrder.pkont,date_in WITH curPeopOrder.dateBeg,dayotp WITH 24,daykont WITH curPeopOrder.ddop,totday WITH people.dayotp+people.daykont,;
                      begdog WITH curPeopOrder.dateBeg,enddog WITH curPeopOrder.dateEnd,dordin WITH datorder.dateOrd,nordin WITH ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord) 
              DO CASE
                 CASE curpeoporder.vDog=1
                      REPLACE dog WITH 1                 
                 OTHERWISE
                      REPLACE dog WITH 2
              ENDCASE                    
              SELECT datjob        
           ENDIF                     
           IF !SEEK(curPeopOrder.nidjob,'datjob',7) 
              SET DELETED OFF 
              SET ORDER TO 7
              GO BOTTOM
              newNidch=nid+1
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                      IF curpeoporder.vdog=2
                         REPLACE kdek WITH VAL(SUBSTR(curpeoporder.varSupl,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl,6,60)
                      ENDIF 
              REPLACE curpeoporder.nidjob WITH newNidch,peoporder.nidjob WITH newNidch  
           ELSE 
              REPLACE kp WITH curPeopOrder.kp,kd WITH curPeopOrder.kd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,kdek WITH 0,fiodek WITH '',lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                      IF curpeoporder.vdog=2
                         REPLACE kdek WITH VAL(SUBSTR(curpeoporder.varSupl,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl,6,60)
                      ENDIF                                               
           ENDIF 
           DO repnadjob   
      CASE curPeopOrder.supOrd=33  &&??????? ?????????????
           IF SEEK(curPeopOrder.kodpeop,'people',1) &&??????? ?? ???????? ?????????
              SELECT people
              REPLACE pkont WITH curpeoporder.pkont,daykont WITH curPeopOrder.ddop           
           ENDIF
           SELECT datjob
           SET ORDER TO 7
           SEEK curPeopOrder.oldNid       
           REPLACE dateOut WITH newDateBeg,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid
       
           IF !SEEK(curPeopOrder.nidjob,'datjob',7) 
              SET DELETED OFF 
              SET ORDER TO 7
              GO BOTTOM
              newNidch=nid+1
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp2,kd WITH curPeopOrder.kd2,kse WITH curPeopOrder.kse2,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.) 
              DO repnadjob                
              REPLACE curpeoporder.nidjob WITH newNidch,peoporder.nidjob WITH newNidch                     
           ELSE 
              REPLACE kp WITH curPeopOrder.kp2,kd WITH curPeopOrder.kd2,kse WITH curPeopOrder.kse2,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                       kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)                                                        
              DO repnadjob                 
           ENDIF               
                     
           IF curpeoporder.kp3#0.AND.curpeoporder.kd3#0
              IF !SEEK(curPeopOrder.nidjob3,'datjob',7) 
                 SET DELETED OFF 
                 SET ORDER TO 7
                 GO BOTTOM
                 newNidch=nid+1
                 SET DELETED ON 
                 APPEND BLANK
                 REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp3,kd WITH curPeopOrder.kd3,kse WITH curPeopOrder.kse3,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                 DO repnadjob                
                 REPLACE curpeoporder.nidjob3 WITH newNidch,peoporder.nidjob3 WITH newNidch    
              ELSE 
                 REPLACE kp WITH curPeopOrder.kp3,kd WITH curPeopOrder.kd3,kse WITH curPeopOrder.kse3,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)                                                         
                 DO repnadjob                
              ENDIF 
           ENDIF  
             
           IF curpeoporder.kp4#0.AND.curpeoporder.kd4#0
              IF !SEEK(curPeopOrder.nidjob4,'datjob',7) 
                 SET DELETED OFF 
                 SET ORDER TO 7
                 GO BOTTOM
                 newNidch=nid+1
                 SET DELETED ON 
                 APPEND BLANK
                 REPLACE nid WITH newNidch, kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kp WITH curPeopOrder.kp4,kd WITH curPeopOrder.kd4,kse WITH curPeopOrder.kse4,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
                 DO repnadjob                
                 REPLACE curpeoporder.nidjob4 WITH newNidch,peoporder.nidjob4 WITH newNidch    
              ELSE 
                 REPLACE kp WITH curPeopOrder.kp4,kd WITH curPeopOrder.kd4,kse WITH curPeopOrder.kse4,tr WITH curPeopOrder.tr,;
                         dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                         kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)                                                         
                 DO repnadjob                
              ENDIF 
           ENDIF            
      CASE INLIST(curPeopOrder.supOrd,31,32)      && ??????? 
           IF SEEK(curPeopOrder.kodpeop,'people',1).AND.curPeopOrder.supOrd=32  &&??????? ?? ???????? ?????????
              SELECT people
              REPLACE pkont WITH curpeoporder.pkont,daykont WITH curPeopOrder.ddop           
           ENDIF 
         
           SELECT datjob
           SET ORDER TO 7
           SEEK curPeopOrder.oldNid
           REPLACE dateOut WITH newDateBeg,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid
           IF curPeopOrder.nidjob=0.OR.!SEEK(curPeopOrder.nidjob,'datjob',7)
              SET DELETED OFF 
              GO BOTTOM
              newNidCh=nid+1              
              SET DELETED ON 
              APPEND BLANK
              REPLACE nid WITH newNidch,kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop
           ELSE
              SEEK(curPeopOrder.nidJob) 
           ENDIF
           REPLACE kp WITH curPeopOrder.nkp,kd WITH curPeopOrder.nkd,kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                   dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                   kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid,lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
           DO repnadjob                
           REPLACE curpeoporder.nidjob WITH datjob.nid,peoporder.nidjob WITH datjob.nid            
      CASE INLIST(curPeopOrder.supOrd,10,41,42,43,44,45,46,47,48,50,70)     && ??????????
           SELECT datJob
           oldOrdJob=SYS(21)
           SET ORDER TO 1
           SEEK curPeopOrder.kodpeop          
           IF FOUND()
              SCAN WHILE koDpeop=curPeopOrder.kodpeop
                   IF EMPTY(dateOut)                              
                      REPLACE dateOut WITH newDateUvol,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid                   
                   ENDIF
              ENDSCAN             
              SELECT datjobout
              APPEND FROM datjob FOR nidpeop=curpeoporder.nidpeop
              SELECT datjob
              DELETE FOR nidpeop=curpeoporder.nidpeop
           ENDIF   
           SELECT datjob
           SET ORDER TO &oldOrdJob
           IF SEEK(curpeoporder.kodpeop,'people',1)
              SELECT people
              REPLACE loguv WITH .T.,date_out WITH newDateUvol,dordout WITH datorder.dateOrd,nordout WITH ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)                      
              SCATTER TO dimout
              SELECT peopout
              APPEND BLANK
              GATHER FROM dimout
              SELECT people
              DELETE
           ENDIF 
           SELECT people
      CASE INLIST(curPeopOrder.supOrd,49)  
           SELECT datJob
           oldOrdJob=SYS(21)
           SET ORDER TO 1
           SEEK curPeopOrder.kodpeop          
           IF FOUND()
              SCAN WHILE koDpeop=curPeopOrder.kodpeop
                   IF EMPTY(dateOut)                              
                      REPLACE dateOut WITH newDateUvol,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid                   
                   ENDIF
              ENDSCAN                        
           ENDIF   
           SELECT datjob
           SET ORDER TO &oldOrdJob           
           SELECT people   
      CASE INLIST(curPeopOrder.supOrd,17,18,19)    && ?????????? ? ????????? ????????
           SELECT people
           oldPeopOrd=SYS(21)
           SET ORDER TO 1
           SEEK curPeopOrder.kodpeop
                REPLACE dog WITH 1,pkont WITH curpeoporder.pkont,daykont WITH curPeopOrder.ddop,numdog WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,1,10)),ALLTRIM(SUBSTR(curPeopOrder.varSupl,1,10)),people.numdog),;
                        strtime WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,21,30)),ALLTRIM(SUBSTR(curPeopOrder.varSupl,21,30)),people.strtime),begDog WITH curPeopOrder.dateBeg,enddog WITH curPeopOrder.dateEnd                                 
           ON ERROR DO erSup
           REPLACE ddog WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,11,10)),CTOD(SUBSTR(curPeopOrder.varSupl,11,10)),people.ddog)                                 
           REPLACE ktime WITH IIF(!EMPTY(SUBSTR(curPeopOrder.varSupl,51,2)),VAL(SUBSTR(curPeopOrder.varSupl,51,20)),people.ktime)                                 
           ON ERROR             
           SET ORDER TO &oldPeopOrd    
      CASE INLIST(curPeopOrder.supord,341,342)  
           SELECT datjob
           SET ORDER TO 7
           SEEK curPeopOrder.oldNid
           SCATTER TO dimnew           
           REPLACE dateOut WITH newDateBeg,dordout WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder,nidout WITH curpeoporder.nid
           IF curPeopOrder.nidjob=0.OR.!SEEK(curPeopOrder.nidjob,'datjob',7)
              SET DELETED OFF 
              GO BOTTOM
              newNidCh=nid+1              
              SET DELETED ON 
              APPEND BLANK
              GATHER FROM dimnew
              REPLACE nid WITH newNidch
           ELSE
              SEEK(curPeopOrder.nidJob)
              newNidCh=curPeopOrder.nidJob 
           ENDIF           
           GATHER FROM dimnew           
           REPLACE nid WITH newNidch,tr WITH 1,dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                   kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),nidin WITH curpeoporder.nid, kse WITH curPeopOrder.kse                                               
           REPLACE dateOut WITH CTOD('  .  .    '),dordout WITH CTOD('  .  .    '),nordout WITH '',kordOut WITH 0,nidout WITH 0
           REPLACE curpeoporder.nidjob WITH datjob.nid,peoporder.nidjob WITH datjob.nid      
   ENDCASE
ENDIF
SELECT curPeopOrder
ON ERROR DO erSup
SELECT datJob
SET ORDER TO &oldOrdJob
SELECT curPeopOrder
GO oldRec
ON ERROR 
DO appendFromPeopOrder
LOCATE FOR nid=newNid 

********************************************************************************************************************************
PROCEDURE procPriemPodr
newKp=curSprPodr.kod  
newNPodr=curSprPodr.name 
SELECT datJob
SET ORDER TO 2
SELECT curDolPodr
SET FILTER TO kp=newKp
SCAN ALL    
     ksesup=0 
     DO sumVacKse
     SELECT curDolPodr
     REPLACE strVac WITH IIF(kse-ksesup=0,'',STR(kse-ksesup,6,2))
ENDSCAN
frmOrd.ComboBox3.RowSource='curDolPodR.name'
WITH .comboBox3         
     .DisplayCount=15
     .ColumnCount=3
     .ColumnWidths='0,50,500'
     .RowSource="curDolPodr.name,strVac,name"
ENDWITH 
frmOrd.ComboBox3.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
frmOrd.ComboBox3.RowSourceType=6
frmOrd.comboBox3.ProcForValid='DO procPriemDolj'
KEYBOARD '{TAB}'

********************************************************************************************************************************
PROCEDURE procPriemDolj
newKd=curDolPodr.kd  
newNDolj=curDolPodr.name
frmOrd.ComboBox3.ControlSource='newNDolj'
frmOrd.comboBox3.Refresh
KEYBOARD '{TAB}'
********************************************************************************************************************************
PROCEDURE formTextOrder
kodSex=IIF(SEEK(newKodPeop,'people',1).AND.people.sex#0,people.sex,1)
procMakeText=IIF(SEEK(newKodPrik,'sprorder',1),sprorder.proctext,'')
IF !EMPTY(procMakeText)
   &procMakeText
   frmOrd.editOrder.SetFocus
ENDIF

*******************************************************************************************************************************
PROCEDURE saveRecOrder
PARAMETERS par1
IF USED('sprtot')
   SELECT sprtot
   SET FILTER TO 
ENDIF
SELECT curPeopOrder
SET ORDER TO 
IF par1      
   SELECT datOrder
   SET ORDER TO 1
   SEEK newKodOrder
   IF !FOUND()
      APPEND BLANK 
      REPLACE kod WITH newKodOrder
   ENDIF
   REPLACE typeOrd WITH typeOrdNew,numOrd WITH numOrdNew,strOrd WITH strOrdNew,dateOrd WITH dateOrdNew 
   repnorder=LTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)
   repdorder=datorder.dateOrd
   SELECT curPeopOrder     
   procSave=IIF(SEEK(newKodPrik,'sprorder',1),sprorder.procsave,'')
   IF !EMPTY(procsave)
      &procSave
   ENDIF
   ordRec=RECNO()
   nppcx=1
   SCAN ALL
        REPLACE npp WITH nppcx
        nppcx=nppcx+1 
   ENDSCAN 
   SELECT datorder
   REPLACE datOrder.nkvo WITH IIF(nppcx<=0,0,nppcx-1)
   SELECT curPeopOrder
   GO ordRec 
   SET ORDER TO 1 
    IF logNewRec 
       DO newRecInOrder WITH .T.   
       RETURN 
   ENDIF 
ELSE 
   IF logNewRec.AND.SEEK(newNid,'peoporder',1)
      SELECT peoporder
      DELETE
      SELECT curPeopOrder  
   ENDIF         
ENDIF     
logRead=.F.
DO removeObjectKadr
strPrik=''
newKodprik=0
logNewRec=.F.
SELECT curPeopOrder
WITH frmOrd
     .grdPers.Enabled=.T.
     .grdPers.SetAll('Enabled',.F.,'ColumnMy')
     .grdPers.Columns(.grdPers.ColumnCount).Enabled=.T.     
     .comboBox1.ControlSource='strPrik'
     .comboBox1.SetFocus
     .grdPers.columns(.grdPers.columnCount).SetFocus     
     .SetAll('Visible',.F.,'MyCommandButton')     
     .butNew.Visible=.T.
     .butRead.Visible=.T.
     .butDel.Visible=.T.
     .butPrn.Visible=.T.
     .butSearch.Visible=.T. 
     .butRet.Visible=.T.        
     .Refresh        
ENDWITH 
       

********************************************************************************************************************************
PROCEDURE removeObjectKadr
WITH frmOrd
     ON ERROR DO erSup
     .RemoveObject('ordtabn')
     .RemoveObject('ordFio')
     .RemoveObject('ordFioNew')
     .RemoveObject('ordPodr')
     .RemoveObject('ordPodr2')
     .RemoveObject('ordPodr3')
     .RemoveObject('ordPodr4')
          
     .RemoveObject('ordPodrNew')
     .RemoveObject('ordDol')
     .RemoveObject('ordDol2')
     .RemoveObject('ordDol3')
     .RemoveObject('ordDol4')
     
     .RemoveObject('ordDolNew')
     .RemoveObject('ordBeg')
     .RemoveObject('ordEnd')
     .RemoveObject('ordPerBeg')
     .RemoveObject('ordPerEnd')
     .RemoveObject('ordDayOtp')
     .RemoveObject('ordDDop')
     .RemoveObject('ordKse')
     .RemoveObject('ordKse2')
     .RemoveObject('ordKse3')
     .RemoveObject('ordKse4')
     
     .RemoveObject('ordTr')
     .RemoveObject('ordPkont')
     .RemoveObject('ordDay')
     .RemoveObject('ordOsnov')  
     .RemoveObject('ordSupl')
     .RemoveObject('ordKomp')
     .RemoveObject('ordPerBeg')
     .RemoveObject('ordPerEnd')
     .RemoveObject('ordNKont')
     .RemoveObject('ordDKont')
     .RemoveObject('ordNkse')
       
     .RemoveObject('tBoxTabn')
     .RemoveObject('tBoxFio')
     .RemoveObject('tBoxFioNew')
     .RemoveObject('tBoxDayOtp')
     .RemoveObject('tBoxDDop')
     .RemoveObject('tBoxBeg')
     .RemoveObject('tBoxEnd')
     .RemoveObject('tBoxPerBeg')
     .RemoveObject('tBoxPerEnd')
     .RemoveObject('tBoxPodr')
     .RemoveObject('tBoxPodr2')
     .RemoveObject('tBoxPodr3')
     .RemoveObject('tBoxPodr4')
     
     .RemoveObject('tBoxPodrNew')     
     .RemoveObject('tBoxDolj')
     .RemoveObject('tBoxDolj2')
     .RemoveObject('tBoxDolj3')
     .RemoveObject('tBoxDolj4')
     
     .RemoveObject('tBoxDoljNew')
     .RemoveObject('tBoxTr')
     .RemoveObject('tBoxLink')
     .RemoveObject('tBoxKont')
     .RemoveObject('tBoxDay')
     .RemoveObject('tBoxOsnov')    
     .RemoveObject('comboBox2')
     .RemoveObject('comboBox22')
     .RemoveObject('comboBox23')
     .RemoveObject('comboBox24')
     
     .RemoveObject('comboBox3')
     .RemoveObject('comboBox32')
     .RemoveObject('comboBox33')
     .RemoveObject('comboBox34')
     .RemoveObject('comboBox40')
          
     .RemoveObject('comboBox4')
     .RemoveObject('comboBox5')
     .RemoveObject('comboBox11')
     .RemoveObject('spinKse')
     .RemoveObject('spinKse2')
     .RemoveObject('spinKse3')
     .RemoveObject('spinKse4')
     
     .RemoveObject('editOrder') 
     .RemoveObject('listBox1')
     .RemoveObject('listBox2')
     .RemoveObject('boxFree')
     .RemoveObject('boxFreeNew')
     .RemoveObject('butKlnt')
     .RemoveObject('butKlntNew')
     .RemoveObject('check1')
     .RemoveObject('comboBox21')
     .RemoveObject('tBoxKomp')
     .RemoveObject('tBoxPerBeg')
     .RemoveObject('tBoxPerEnd')
     .RemoveObject('tBoxNKont')
     .RemoveObject('tBoxDKont')
     .RemoveObject('tBoxNkse')
     
     
     ON ERROR 
     .Refresh
ENDWITH
********************************************************************************************************************************
PROCEDURE changeRowPeopOrder
procView=IIF(SEEK(curPeopOrder.supOrd,'sprorder',1),sprorder.procview,'')
IF !EMPTY(procView)
   &procView
   frmOrd.grdPers.Columns(frmOrd.grdPers.columnCount).SetFocus
   frmOrd.Refresh
ENDIF
*******************************************************************************************************************************
PROCEDURE delFromOrder
SELECT curPeopOrder
IF npp=0
   RETURN
ENDIF
WITH frmOrd
     logNewRec=.F.
     .SetAll('Visible',.F.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')    
     .butDelRec.Visible=.T.
     .butDelRet.Visible=.T.
     .grdPers.Enabled=.F.
ENDWITH
*******************************************************************************************************************************
PROCEDURE delRecOrder
PARAMETERS par1
IF par1
   IF SEEK(curPeopOrder.nid,'peoporder',1)     
      procDelOrdJob=IIF(SEEK(newKodPrik,'sprorder',1),sprorder.procdel,'')
      IF !EMPTY(procDelOrdJob)
         &procDelOrdJob
      ENDIF
      SELECT peoporder
      DELETE
      DO appendFromPeopOrder     
      SET ORDER TO 1
      GO TOP
   ENDIF     
ENDIF
WITH frmOrd
     logNewRec=.F.
     .butNew.Visible=.T.
     .butRead.Visible=.T.
     .butDel.Visible=.T.
     .butPrn.Visible=.T.  
     .butSearch.Visible=.T.  
     .butRet.Visible=.T.
     .butRetNew.Visible=.F.
     .butDelRec.Visible=.F.
     .butDelRet.Visible=.F.
     .grdPers.Enabled=.T.
     .grdPers.SetAll('Enabled',.F.,'ColumnMy')
     .grdPers.Columns(.grdPers.ColumnCount).Enabled=.T.
     .grdPers.Columns(.grdPers.ColumnCount).SetFocus 
     DO changeRowPeopOrder
     .Refresh    
ENDWITH
********************************************************************************************************************************
PROCEDURE delSovm
IF SEEK(curPeopOrder.nidJob,'datjob',7).AND.datjob.kordin=curpeoporder.kord.AND.datjob.kodpeop=curpeoporder.kodpeop
   SELECT datJob 
   DELETE 
ENDIF 
********************************************************************************************************************************
*                                    ???????? ??????
********************************************************************************************************************************
PROCEDURE procotptrud
PARAMETERS par1,parper
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF 
logPer=parper 
SELECT kod,name,otm,fl FROM sprtot WHERE sprtot.kspr=23 INTO CURSOR curSuplOrd ORDER BY kod   && ?????? ??? ?????????????? ?????? ?? ???????
parPadejDol=2
parPadejFio=2
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newDayOtp=IIF(logNewRec,0,curpeoporder.DayOtp)
newDDop=IIF(logNewRec,0,VAL(SUBSTR(curpeoporder.varsupl,1,2)))
newDNorm=IIF(logNewRec,0,VAL(SUBSTR(curpeoporder.varsupl,3,2)))
newPerBeg=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.perbeg)
newPerEnd=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.perend)

newPbegOld=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.pbegold)   
newPbegNew=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.pbegnew) 
dayCost=IIF(logNewRec,0,curpeoporder.dcost) 

newLogApp=.T.
newOrdSupl=IIF(logNewRec,SPACE(100),curpeoporder.ordSupl)
newNpp=0
logDatJob=.F.
objFocus='frmOrd.tBoxDayOtp'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     DO adTboxAsCont WITH 'frmOrd','ordDayOtp',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'???-?? ????',1,1   
  
     DO adTboxNew WITH 'frmOrd','tBoxDayotp',.ordDayOtp.Top,.comboBox1.Left,RetTxtWidth('99999'),dHeight,'newDayOtp',.F.,IIF(par1,.T.,.F.),0,'99'  
     .tBoxDayOtp.Alignment=0  
     
     DO adTboxAsCont WITH 'frmOrd','ordDDop',.tBoxDayOtp.Left+.tBoxDayOtp.Width-1,.ordDayOtp.Top,(.comboBox1.Width-.tBoxDayOtp.Width*3+4)/2,dHeight,'???.????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDdop',.ordDDop.Top,.ordDDop.Left+.ordDDop.Width-1,.tBoxDayOtp.Width,dHeight,'newDdop',.F.,IIF(par1,.T.,.F.),0,'99' 
     .tBoxDDop.Alignment=0                
     
     DO adTboxAsCont WITH 'frmOrd','ordPKont',.tBoxDDop.Left+.tBoxDDop.Width-1,.ordDayOtp.Top,.ordDDop.Width,dHeight,'??????.',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordDDop.Top,.ordPKont.Left+.ordPkont.Width-1,.ComboBox1.Width-.tBoxDayOtp.Width*2-.ordDDop.Width*2+4,dHeight,'newDNorm',.F.,IIF(par1,.T.,.F.),0,'99' 
     .tBoxKont.Alignment=0     
        
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDayOtp.Top+dHeight-1,.ordPrik.Width,dHeight,'???????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.,"DO validPerOtp WITH 'newDateBeg','newDateEnd','newDayOtp',.T."        
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
     boxTop=.ordBeg.Top+dHeight-1    
     IF newKodPrik=51  && ??? ???????? ???????
        DO adTboxAsCont WITH 'frmOrd','orddol2',.ordPrik.Left,boxTop,.ordPrik.Width,dHeight,'???????? ??????? ??? ??? ',1,1   
        DO adTboxNew WITH 'frmOrd','tBoxDolj2',boxTop,.comboBox1.Left,.tBoxDayOtp.Width,dHeight,'dayCost',.F.,IIF(par1,.T.,.F.),0,.F.        
        DO adTboxAsCont WITH 'frmOrd','orddol3',.tBoxDolj2.Left+.tBoxDolj2.Width-1,boxTop,RetTxtWidth('w??w'),dHeight,'?',1,1   
        DO adTboxNew WITH 'frmOrd','tBoxDolj3',boxTop,.orddol3.Left+.orddol3.Width-1,(.comboBox1.Width-.tBoxDolj2.Width-.ordDol3.Width*2)/2,dHeight,'newPbegOld',.F.,IIF(par1,.T.,.F.),0,.F.        
        DO adTboxAsCont WITH 'frmOrd','orddol4',.tBoxDolj3.Left+.tBoxDolj3.Width-1,boxTop,.ordDol3.Width,dHeight,'??',1,1   
        DO adTboxNew WITH 'frmOrd','tBoxDolj4',boxTop,.orddol4.Left+.orddol4.Width-1,.comboBox1.Width-.tBoxDolj2.Width-.ordDol3.Width*2-.tBoxDolj3.Width+4,dHeight,'newPbegNew',.F.,IIF(par1,.T.,.F.),0,.F.        
        boxTop=.ordDol2.Top+dHeight-1    
     ENDIF
     
     DO adTboxAsCont WITH 'frmOrd','ordPerBeg',.ordPrik.Left,boxTop,.ordPrik.Width,dHeight,'?? ?????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxPerBeg',.ordPerBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPerBeg',.F.,IIF(par1,.T.,.F.),0,.F.,'newPerEnd=newPerBeg+365'      
     DO adTboxAsCont WITH 'frmOrd','ordPerEnd',.ordPrik.Left,.ordPerBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxPerEnd',.ordPerBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPerEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxPerBeg.Width=(.comboBox1.Width-.ordPerEnd.Width)/2
     .ordPerEnd.Left=.tBoxPerBeg.left+.tBoxPerBeg.Width-1
     .tBoxPerEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxPerEnd.Width=.comboBox1.Width-.tBoxPerBeg.Width-.ordPerEnd.Width+2      
     DO adTboxAsCont WITH 'frmOrd','ordSupl',.ordPrik.Left,.ordPerBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'??????????',1,1          
     DO addcombomy WITH 'frmOrd',21,.comboBox1.Left,.ordSupl.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newOrdSupl','ALLTRIM(curSuplOrd.name)',6,.F.,'newOrdSupl=curSuplOrd.name',.F.,.T. 
     .comboBox21.DisabledForeColor=RGB(1,0,0)     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordSupl.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1            
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
    
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
        IF par1  
        ON ERROR DO erSup         
        .tBoxFio.SetFocus
        ON ERROR         
     ENDIF
ENDWITH
********************************************************************************************************************************
PROCEDURE procOrdNameDol
cOrdFio=ALLTRIM(newFioPeop)
DO CASE
   CASE parPadejFio=1
        cOrdFio=IIF(SEEK(newKodPeop,'people',1).AND.!EMPTY(people.fior),ALLTRIM(people.fior),cOrdFio)  
   CASE parPadejFio=2
        cOrdFio=IIF(SEEK(newKodPeop,'people',1).AND.!EMPTY(people.fiod),ALLTRIM(people.fiod),cOrdFio)  
   CASE parPadejFio=3
        cOrdFio=IIF(SEEK(newKodPeop,'people',1).AND.!EMPTY(people.fiov),ALLTRIM(people.fiov),cOrdFio)  
ENDCASE 
cOrdPodr=IIF(SEEK(newKp,'sprpodr',1),ALLTRIM(sprpodr.nameord),ALLTRIM(newNPodr))
newNDolj=ALLTRIM(newNDolj)
IF SEEK(newKd,'sprdolj',1)
   IF sprdolj.logSex.AND.kodSex=1
      DO CASE
         CASE parPadejDol=1
              cOrdDol=IIF(!EMPTY(sprdolj.namerm),ALLTRIM(sprdolj.namerm),newNDolj)
         CASE parPadejDol=2
              cOrdDol=IIF(!EMPTY(sprdolj.namedm),ALLTRIM(sprdolj.namedm),newNDolj)                                
         CASE parPadejDol=3
              cOrdDol=IIF(!EMPTY(sprdolj.nametm),ALLTRIM(sprdolj.nametm),newNDolj)     
      ENDCASE
   ELSE        
      DO CASE
         CASE parPadejDol=1
              cOrdDol=IIF(!EMPTY(sprdolj.namer),ALLTRIM(sprdolj.namer),newNDolj)
         CASE parPadejDol=2
              cOrdDol=IIF(!EMPTY(sprdolj.named),ALLTRIM(sprdolj.named),newNDolj)                                
         CASE parPadejDol=3
              cOrdDol=IIF(!EMPTY(sprdolj.namet),ALLTRIM(sprdolj.namet),newNDolj)     
      ENDCASE             
   ENDIF
ELSE
   cOrdDol=ALLTRIM(newNDolj)
ENDIF
SELECT txtOrder
********************************************************************************************************************************
PROCEDURE procOrdNameDolZam
PARAMETERS parfio,pardol,parpodr,padFio,padDol
*parFio ?????????? ??? ???
*parDol ?????????? ??? ?????????
*parPodr ?????????? ??? ?????????????
*padFio ????? ???????
*padDol ????? ?????????
parFio=ALLTRIM(newFioZam)
kSexZam=kodSex
nkPeop=newKodPeop
*nkPeop=IIF(newTabZam#0,newTabZam,newKodPeop)
nkPeop=IIF(newTabZam#0,newTabZam,0)
DO CASE
   CASE padFio=1
        cFioZam=IIF(SEEK(nkPeop,'people',1).AND.!EMPTY(people.fior),ALLTRIM(people.fior),parFio)  
        kSexZam=people.sex
   CASE padFio=2
        cFioZam=IIF(SEEK(nkPeop,'people',1).AND.!EMPTY(people.fiod),ALLTRIM(people.fiod),parFio)  
        kSexZam=people.sex
   CASE padFio=3
        cFioZam=IIF(SEEK(nkPeop,'people',1).AND.!EMPTY(people.fiov),ALLTRIM(people.fiov),parFio)  
        kSexZam=people.sex
   CASE padFio=4
        cFioZam=IIF(SEEK(nkPeop,'people',1).AND.!EMPTY(people.fiot),ALLTRIM(people.fiot),parFio)  
        kSexZam=people.sex     
ENDCASE 
cPodrZam=IIF(SEEK(newsKp,'sprpodr',1),ALLTRIM(sprpodr.nameord),ALLTRIM(&parpodr))
cDolZam=ALLTRIM(&parDol)
IF SEEK(newSkd,'sprdolj',1)
   IF sprdolj.logSex.AND.kSexZam=1
      DO CASE
         CASE padDol=1
              cDolZam=IIF(!EMPTY(sprdolj.namerm),ALLTRIM(sprdolj.namerm),&parDol)
         CASE padDol=2
              cDolZam=IIF(!EMPTY(sprdolj.namedm),ALLTRIM(sprdolj.namedm),&parDol)                                
         CASE padDol=3
              cDolZam=IIF(!EMPTY(sprdolj.nametm),ALLTRIM(sprdolj.nametm),&parDol)   
         CASE padDol=4
              cDolZam=IIF(!EMPTY(sprdolj.namevm),ALLTRIM(sprdolj.namevm),&parDol)         
      ENDCASE
   ELSE        
      DO CASE
         CASE padDol=1
              cDolZam=IIF(!EMPTY(sprdolj.namer),ALLTRIM(sprdolj.namer),&parDol)            
         CASE padDol=2
              cDolZam=IIF(!EMPTY(sprdolj.named),ALLTRIM(sprdolj.named),&parDol)                                
         CASE padDol=3
              cDolZam=IIF(!EMPTY(sprdolj.namet),ALLTRIM(sprdolj.namet),&parDol)  
         CASE padDol=4
              cDolZam=IIF(!EMPTY(sprdolj.namev),ALLTRIM(sprdolj.namev),&parDol)                  
      ENDCASE             
   ENDIF
ELSE
   parDol=ALLTRIM(parDol)
ENDIF
SELECT txtOrder
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ? ???????? ???????
********************************************************************************************************************************
PROCEDURE textOtpTrud
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')

strPerBeg=IIF(!EMPTY(newPerBeg),dateToString('newPerBeg',.T.),'')
strPerEnd=IIF(!EMPTY(newPerEnd),dateToString('newPerEnd',.T.),'')

cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol

SELECT txtOrder
IF newKodPrik=51.AND.dayCost#0
   REPLACE txtprn WITH UPPER(LEFT(ALLTRIM(cOrdDol),1))+LOWER(SUBSTR(ALLTRIM(cOrdDol),2))+' '+LOWER(cOrdPodr)+' '+cOrdFio+':'+CHR(13)+;
           '1.'+'???????? ??????? ??? ?? '+LTRIM(STR(dayCost))+' ??????????? ???? (c '+DTOC(newPbegOld)+' ?. ?? '+DTOC(newPbegNew)+;
           ' ?.) ? ????? ? ??????????????? ? ??????? ???????? ???? ??????? ??? ?????????? ?????????? ????? ? ??????? ??????????? 14 ??????????? ????.'+CHR(13)+;
           '2. ???????????? ???????? ?????? ?? '+ALLTRIM(STR(newDayOtp))+IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ????',' ??????????? ???'))

    IF !EMPTY(newDDop)
       REPLACE txtprn WITH txtprn+' (?? ??? '+ALLTRIM(STR(newDDop))+IIF(newDDop>4,' ??????????? ????',IIF(newDDop=1,' ??????????? ???? ',' ??????????? ??? '))+' ?? ????????'+IIF(newDNorm#0,',',')')                     
    ENDIF
    IF !EMPTY(newDNorm)
       REPLACE txtprn WITH txtprn+IIF(newDDop=0,' (?? ??? ','') +ALLTRIM(STR(newDNorm))+IIF(newDNorm>4,' ??????????? ????',IIF(newDNorm=1,' ??????????? ???? ?? ??????????????? ??????? ????)',' ??????????? ??? ?? ??????????????? ??????? ????)'))
    ENDIF
        
    REPLACE txtPrn WITH txtPrn+' c '+strDateBeg+' ?? '+strDateEnd+' ?? ??????? ?????? c '+strPerBeg+' ?? '+strPerEnd+CHR(13)+;
           IIF(EMPTY(newOrdSupl),CHR(13)+'?????????: '+newOsnov,'3.'+ALLTRIM(newOrdSupl)+'.'+CHR(13)+'?????????: '+newOsnov) 
ELSE
    REPLACE txtprn WITH UPPER(LEFT(ALLTRIM(cOrdDol),1))+LOWER(SUBSTR(ALLTRIM(cOrdDol),2))+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+IIF(EMPTY(newOrdSupl),'',':'+CHR(13)+'1.')+IIF(logPer,'???????????? ???????? ?????? ?? ','???????????? ????? ????????? ??????? ?? ')+ALLTRIM(STR(newDayOtp))+;
            IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ????',' ??????????? ???'))
    IF !EMPTY(newDDop)
       REPLACE txtprn WITH txtprn+' (?? ??? '+ALLTRIM(STR(newDDop))+IIF(newDDop>4,' ??????????? ????',IIF(newDDop=1,' ??????????? ???? ',' ??????????? ??? '))+' ?? ????????'+IIF(newDNorm#0,',',')')                     
    ENDIF
    IF !EMPTY(newDNorm)
       REPLACE txtprn WITH txtprn+IIF(newDDop=0,' (?? ??? ','') +ALLTRIM(STR(newDNorm))+IIF(newDNorm>4,' ??????????? ????',IIF(newDNorm=1,' ??????????? ???? ?? ??????????????? ??????? ????)',' ??????????? ??? ?? ??????????????? ??????? ????)'))
    ENDIF
        
    REPLACE txtPrn WITH txtPrn+' c '+strDateBeg+' ?? '+strDateEnd+' ?? ??????? ?????? c '+strPerBeg+' ?? '+strPerEnd+CHR(13)+;
           IIF(EMPTY(newOrdSupl),CHR(13)+'?????????: '+newOsnov,'2.'+ALLTRIM(newOrdSupl)+'.'+CHR(13)+'?????????: '+newOsnov) 
ENDIF         
********************************************************************************************************************************
PROCEDURE saveRecOtptrud
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textOtpTrud 
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,supord WITH newKodprik,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,dayOtp WITH newDayOtp,supord WITH newKodprik,;
        perBeg WITH newPerBeg,perEnd WITH newPerEnd,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,kodOtp WITH IIF(SEEK(supOrd,'curSprOrder',1),curSprOrder.kodotp,0),;
        ordSupl WITH newOrdSupl,varSupl WITH PADR(ALLTRIM(STR(newDDop)),2,' ')+PADR(ALLTRIM(STR(newDNorm)),2,' '),nid WITH newNid,dord WITH repdorder,nord WITH repnorder,;
        dcost WITH daycost,pBegOld WITH newPbegOld,pBegNew WITH newPbegNew,sex WITH kodsex
DO saveDimOrd   
DO saveotpOrder    
********************************************************************************************************************************
*                                   ?????????? ??????
********************************************************************************************************************************
PROCEDURE procotpsoc
PARAMETERS par1,par2
nUvol=par2
IF nUvol=2
   IF !USED('sprtot')	
       USE sprtot IN 0
   ENDIF
   SELECT sprtot
   SET FILTER TO kspr=31
ENDIF 
parPadejDol=2
parPadejFio=2
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newDayOtp=IIF(logNewRec,0,curpeoporder.DayOtp)
newKprich=IIF(logNewRec,0,curpeoporder.kprich)
newStrPrich=IIF(logNewRec,'',curpeoporder.strprich)
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxDayOtp'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     DO adTboxAsCont WITH 'frmOrd','ordDayOtp',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'???-?? ????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDayotp',.ordDayOtp.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDayOtp',.F.,IIF(par1,.T.,.F.),0,'99' 
     .tBoxDayOtp.Alignment=0
     
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDayOtp.Top+dHeight-1,.ordPrik.Width,dHeight,IIF(nuvol=1,'???????????? ?','???????? ?'),1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.,;
        IIF(nuvol=1,"DO validPerOtp WITH 'newDateBeg','newDateEnd','newDayOtp',.F.","DO validPerOtp WITH 'newDateBeg','newDateEnd','newDayOtp',.T.")
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2    
     
     DO adTboxAsCont WITH 'frmOrd','ordPerBeg',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'???????',1,1   
     
     IF nUvol=1
        DO addComboMy WITH 'frmOrd',11,.comboBox1.Left,.ordPerBeg.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newStrPrich','ALLTRIM(curPrichOtp.name)',6,.F.,'DO validPrichSoc WITH 1',.F.,.T. 
     ELSE 
        DO addComboMy WITH 'frmOrd',11,.comboBox1.Left,.ordPerBeg.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newStrPrich','ALLTRIM(sprtot.name)',6,.F.,'DO validPrichSoc WITH 2',.F.,.T. 
     ENDIF    
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordPerBeg.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'newStrPrich',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)           
              
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordPerBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0       
   
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
     
          
     IF par1      
        ON ERROR DO erSup
        .comboBox11.SetFocus     
        .tBoxFio.SetFocus        
        ON ERROR 
     ENDIF
ENDWITH
*************************************************************************************************************************
PROCEDURE validPerOtp
PARAMETERS parBeg,parEnd,parDay,parTrud
* parBeg - ?????? ???????
* parEnd - ????????? ???????
* parDay - ???? ???????
* parVid
&parEnd=&parBeg+&parDay-1
IF parTrud
   SELECT fete
   varForDay=&parDay
   varOtpCur=&parBeg   
   DO WHILE .T.
      LOCATE FOR MONTH(varOtpCur)=mfete.AND.dfete=DAY(varOtpCur)
      IF FOUND()
         &parEnd=&parEnd+1              
      ENDIF  
      varOtpCur=varOtpCur+1
      IF varOtpCur>&parEnd
        EXIT 
      ENDIF
   ENDDO
 
   *SELECT fete
   *LOCATE FOR MONTH(&parEnd)=mfete.AND.DAY(&parEnd)=dFete
   *IF FOUND()      
   *   &parEnd=&parEnd+1 
   *ENDIF 
   *LOCATE FOR MONTH(&parEnd)=mfete.AND.DAY(&parEnd)=dFete
   *IF FOUND()      
   *   &parEnd=&parEnd+1 
   *ENDIF 
   *LOCATE FOR MONTH(&parEnd)=mfete.AND.DAY(&parEnd)=dFete
   *IF FOUND()      
   *   &parEnd=&parEnd+1       
   *ENDIF 
ENDIF  

SELECT curPeopOrder
********************************************************************************************************************************
PROCEDURE validPrichSoc
PARAMETERS par1
DO CASE
   CASE par1=1
        newKPrich=curPrichOtp.kod
   CASE par1=2
        newKPrich=sprtot.kod     
ENDCASE         
KEYBOARD '{TAB}'
********************************************************************************************************************************
*                         ???????????? ????????? ????? ??????? ? ?????????? ???????
********************************************************************************************************************************
PROCEDURE textOtpSoc
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')

cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol

SELECT txtOrder
DO CASE
   CASE nuvol=1
        REPLACE txtprn WITH  UPPER(LEFT(ALLTRIM(cOrdDol),1))+LOWER(SUBSTR(ALLTRIM(cOrdDol),2))+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' ???????????? ?????????? ?????? ??? ?????????? ?????????? ????? ?? '+ALLTRIM(STR(newDayOtp))+' ??????????? ???? '+;
                ' c '+strDateBeg+' ?? '+strDateEnd+' '+ALLTRIM(newStrPrich)+'.'+CHR(13)+CHR(13)+'?????????: '+ newOsnov         
   CASE nuvol=2
        REPLACE txtprn WITH  UPPER(LEFT(ALLTRIM(cOrdDol),1))+LOWER(SUBSTR(ALLTRIM(cOrdDol ),2))+' '+LOWER(cordPodr)+' '+ALLTRIM(cordFio)+' ???????? ???????? ?????? ?? '+ALLTRIM(STR(newDayOtp))+' ??????????? ???? '+;
                ' c '+strDateBeg+' ?? '+strDateEnd+' '+ALLTRIM(newStrPrich)+'.'+CHR(13)+CHR(13)+'?????????: '+ newOsnov                         
ENDCASE        
********************************************************************************************************************************
PROCEDURE saveRecOtpSoc
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textOtpSoc
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,supord WITH newKodprik,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,dayOtp WITH newDayOtp,supord WITH newKodprik,;
        kPrich WITH newKprich,strPrich WITH newStrPrich,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,kodOtp WITH IIF(SEEK(supOrd,'curSprOrder',1),curSprOrder.kodotp,0),nid WITH newNid,;
        dord WITH repdorder,nord WITH repnorder,sex WITH kodsex        
DO saveDimOrd  
DO saveotpOrder               
********************************************************************************************************************************
*                                   ???? ??????, ?????????, ?? 3-???, ?? ???????????? ? ?????
********************************************************************************************************************************
PROCEDURE procotpsup
PARAMETERS par1,par2
vSup=par2
** vSup=1 - ???? ??????
** vSup=2,7 - ????????? ????
** vSup=3 - ?? ???????????? ? ?????
** vSup=4 - ?? ????? ?? ???????? ?? 3 ???

parPadejDol=IIF(vSup#7,2,1)
parPadejFio=IIF(vSup#7,2,1)
PUBLIC newKp2,newNpodr2
newKp2=0
newNpodr2=''
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newDayOtp=IIF(logNewRec,IIF(INLIST(vSup,1,2,7),1,0),curpeoporder.DayOtp)
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxBeg'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     DO adTboxAsCont WITH 'frmOrd','ordDayOtp',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'???-?? ????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDayotp',.ordDayOtp.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDayOtp',.F.,IIF(par1,.T.,.F.),0,'99' 
     .tBoxDayOtp.Alignment=0
     IF INLIST(vSup,3,4,5,10)
        .tBoxDayOtp.Enabled=.F. 
     ENDIF
     
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDayOtp.Top+dHeight-1,.ordPrik.Width,dHeight,IIF(INLIST(vSup,5,10),'???????? ?','???????????? ?'),1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.,IIF(INLIST(vSup,3,4,5,10),'',"DO validPerOtp WITH 'newDateBeg','newDateEnd','newDayOtp',.F.")        
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0 
      IF INLIST(vSup,5,10)
        .tBoxEnd.Enabled=.F. 
     ENDIF
      
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
  
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
          
     IF par1 
        ON ERROR DO erSup       
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH  
********************************************************************************************************************************
PROCEDURE textOtpSup
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cordFio=''
cOrdDol=''
cOrdPodr=''
DO procOrdNameDol
SELECT txtOrder
DO CASE
   CASE vSup=1
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' ???????????? ???? ?????????????? ????????? ?? ?????? ???? '+strDateBeg+;
        ' , ??? ??????, ????????????? ????? ? ???????? ?? 16 ???, ? ??????? ? ??????? ???????? ????????  ?????????.'+CHR(13)+CHR(13)+'?????????: '+newOsnov
   CASE vSup=2
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+', '+strdateBeg+' ???????????? ????????? ?? ?????? ???? ? ????? ?? ?????? ????? ? ?? ???????????.'+;
        CHR(13)+CHR(13)+'?????????: '+newOsnov
   CASE vSup=3
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' c '+strDateBeg+' ?? '+strDateEnd+' ???????????? ?????????? ?????? ?? ???????????? ? ?????.'+;
        CHR(13)+CHR(13)+'?????????: '+newOsnov
   CASE vSup=4
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' ???????????? ?????????? "?????? ?? ????? ?? ???????? ?? ?????????? ?? ???????? ???? ???"'+' c '+strDateBeg+' ?? '+strDateEnd+;
        CHR(13)+CHR(13)+'?????????: '+newOsnov
   CASE vSup=5
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' c '+strDateBeg+' ???????? ?????????? "?????? ?? ????? ?? ???????? ?? ?????????? ?? ???????? ???? ???", ? ????? ? ?????? ? ?????? ?? ???????????? ? ?????.'+;
        CHR(13)+'?????????: '+newOsnov
   CASE vSup=6
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' c '+strDateBeg+' ???????? ?????????? "?????? ?? ????? ?? ???????? ?? ?????????? ?? ???????? ???? ???", ? ????? ? ?????? ? ?????? ?? ???????????? ? ?????.'+;
        CHR(13)+'?????????: '+newOsnov     
    CASE vSup=7
        REPLACE txtprn WITH '?????????? '+ALLTRIM(newFioPeop)+', '+LOWER(ALLTRIM(cOrdDol))+' '+LOWER(cOrdPodr)+', ?? 1 ??????????? ????, '+strdateBeg+', ?? ?????? ??? ?????????? ????????? ???????.'+;
        CHR(13)+CHR(13)+'?????????: '+newOsnov     
   CASE vSup=10
   
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' c '+strDateBeg+' ?????????? ? ?????? ?? ??????????? "??????? ?? ????? ?? ???????? ?? ?????????? ?? ???????? ???? ???",? ??????? ???????? ???????? ??????????.'+;
        CHR(13)+'?????????: '+newOsnov
        
ENDCASE

********************************************************************************************************************************        
PROCEDURE saveRecOtpSup   
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textOtpSup
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,supord WITH newKodprik,dateBeg WITH newDateBeg,dateEnd WITH newDateEnd,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,dayOtp WITH newDayOtp,supord WITH newKodprik,;
        osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,kodOtp WITH IIF(SEEK(supOrd,'curSprOrder',1),curSprOrder.kodotp,0),nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex      
DO saveDimOrd  
DO saveotpOrder               
******************************************************************************************************************************************************        
PROCEDURE saveOtpOrder        
SELECT datOtp
SET ORDER TO 5
IF curpeoporder.supord=501
   SEEK curPeopOrder.nidJob
   IF FOUND()
      DELETE
   ENDIF
ELSE 
   SEEK curPeopOrder.nid
   IF !FOUND()
      APPEND BLANK
      REPLACE kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kord WITH newKodOrder,nidord WITH curPeopOrder.nid
   ENDIF 
   REPLACE kodOtp WITH curpeopOrder.kodOtp,begOtp WITH curPeopOrder.dateBeg,endotp WITH curPeopOrder.dateEnd,kvoday WITH curPeopOrder.dayOtp,dayKont WITH curPeopOrder.ddop,;              
           perBeg WITH curPeopOrder.perBeg,perEnd WITH curPeopOrder.perEnd,kprich WITH curPeopOrder.kPrich,txtPrich WITH curPeopOrder.strPrich,;
           nameotp WITH IIF(SEEK(kodotp,'curSprOtp',1),curSprOtp.name,''),osnov WITH '??.? '+LTRIM(STR(datOrder.numOrd))+'-'+datOrder.strOrd+' ?? '+DTOC(datOrder.dateOrd)      
   IF vSup=4.AND.curpeoporder.supord=62.AND.SEEK(curPeopOrder.kodpeop,'people',1)
      REPLACE people.dekOtp WITH .T.,people.bdekotp WITH curPeopOrder.dateBeg,people.ddekotp WITH curPeopOrder.dateEnd  
   ENDIF    
ENDIF 
SELECT datOtp
SET ORDER TO 1
ON ERROR DO erSup
DO appendFromPeopOrder
ON ERROR
LOCATE FOR nid=newNid
********************************************************************************************************************************
PROCEDURE delotpord
IF SEEK(curPeopOrder.nid,'datotp',5).AND.datotp.kord=curpeoporder.kord.AND.datotp.kodpeop=curpeoporder.kodpeop
   SELECT datotp 
   DELETE 
ENDIF 
********************************************************************************************************************************
PROCEDURE selectPeopOrd
SELECT curSuplPeop
ZAP
APPEND FROM people
WITH frmOrd
     .listBox1.RowSource='curSuplpeop.fio'                    
     str_ini=''
     DO procfioini WITH 'curSuplPeop.fio'    
     SELECT curSuplPeop
     LOCATE FOR fio=frmOrd.tBoxFio.Text
     IF .listBox1.Visible=.F.
        .listBox1.Visible=.T.  
        .listBox1.SetFocus            
     ENDIF 
ENDWITH 
*****************************************************************************************************************
PROCEDURE changePeopOrd
WITH frmOrd 
     IF .listBox1.Visible=.F.
        .listBox1.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=frmOrd.tBoxFio.Text 
SELECT curSuplPeop
ZAP
APPEND FROM people FOR LEFT(LOWER(fio),LEN(ALLTRIM(lcValue)))=LOWER(ALLTRIM(lcValue))
WITH frmOrd.listBox1
     .RowSource='curSuplpeop.fio'                        
     .Visible=IIF(RECCOUNT('curSuplPeop')=0,.F.,.T.)      
ENDWITH 
************************************************************************************************************************
PROCEDURE validListPeop
oldKodpeop=newKodpeop
newKodpeop=curSuplPeop.num
newNidPeop=curSuplPeop.nid
kodSex=IIF(curSuplPeop.sex#0,curSuplPeop.Sex,1)
newFioPeop=curSuplPeop.fio
str_ini=''
DO procfioini WITH 'curSuplPeop.fiot' 
IF !EMPTY(procForOsnov) 
   newOsnov=&procForOsnov
ENDIF    
newNidJob=curSuplPeop.nidjob
DO CASE
   CASE logNewRec
        IF newnidJob>0      
           newKp=IIF(SEEK(newnidjob,'datjob',7),datJob.kp,0)
           newKd=datJob.kd
           oldNidJob=datjob.nid
           newNpodr=IIF(SEEK(newKp,'sprpodr',1),sprpodr.name,'')
           newNdolj=IIF(SEEK(newKd,'sprdolj',1),sprdolj.name,'')
        ELSE            
           SELECT datjob
           jobOrdOld=SYS(21)
           SET ORDER TO 1
           SEEK newKodpeop
           SCAN WHILE kodpeop=newkodpeop
                IF INLIST(tr,1,3).AND.EMPTY(dateOut)
                   newKp=datJob.kp
                   newKd=datJob.kd
                   oldNidJob=datjob.nid
                   newNpodr=IIF(SEEK(newKp,'sprpodr',1),sprpodr.name,'')                                
                   EXIT 
                ENDIF
           ENDSCAN     
           SET ORDER TO &jobOrdOld
           SELECT curPeopOrder        
           newNdolj=IIF(SEEK(newKd,'sprdolj',1),sprdolj.name,'')           
        ENDIF 
   CASE !logNewRec  
        IF newKodPeop#oldKodPeop
           SELECT datjob
           jobOrdOld=SYS(21)
           SET ORDER TO 1
           SEEK newKodpeop
           SCAN WHILE kodpeop=newkodpeop
                IF INLIST(tr,1,3).AND.EMPTY(dateOut)
                   newKp=datJob.kp
                   newKd=datJob.kd
                   oldNidJob=datjob.nid
                   newNpodr=IIF(SEEK(newKp,'sprpodr',1),sprpodr.name,'')                                
                   EXIT 
                ENDIF
           ENDSCAN     
           SET ORDER TO &jobOrdOld
           SELECT curPeopOrder        
           newNdolj=IIF(SEEK(newKd,'sprdolj',1),sprdolj.name,'')   
        ELSE   
           newNpodr=curPeopOrder.npodr
           newNdolj=curPeopOrder.ndolj
  *         DO padejInOrder  
        ENDIF   
ENDCASE 
frmOrd.tBoxFio.ControlSource='newFioPeop'
frmOrd.listBox1.Visible=.F.
frmOrd.tBoxFio.Refresh
frmOrd.tBoxOsnov.ControlSource='newOsnov'
IF logNewRec.AND.logPer
   SELECT * FROM datOtp WHERE nidpeop=newNidpeop INTO CURSOR curOrdOtp ORDER BY perBeg DESCENDING READWRITE 
   SELECT curOrdOtp
   DELETE FOR kodotp>3   
   GO TOP
   LOCATE FOR !EMPTY(perEnd) && ???????????? ??? ???????? ?? ?????? ?????????
   IF kodotp=1
      newPerBeg=IIF(!EMPTY(perEnd),perEnd+1,newPerBeg)
      newPerEnd=IIF(!EMPTY(newPerBeg),IIF(MOD(YEAR(newPerBeg),4)=0,newPerBeg+365-1,newPerBeg+365),newPerEnd)
   ELSE 
      newPerBeg=IIF(!EMPTY(perBeg),perBeg,newPerBeg)
      newPerEnd=IIF(!EMPTY(newPerBeg),IIF(MOD(YEAR(newPerBeg),4)=0,newPerBeg+365-1,newPerBeg+365),newPerEnd)
   ENDIF  
   
   IF newKodPrik=51
      SELECT datOtp
      ordOtpold=SYS(21)
      SET ORDER TO 6
      SEEK newNidPeop
      daycx=0
      SCAN WHILE nidPeop=newNidPeop
           IF kodotp=6       
              DO CASE
                 CASE begotp>=newPerBeg.AND.endotp<=newPerEnd
                      daycx=daycx+kvoday                  
                 CASE begotp<newPerBeg.AND.endotp<newPerEnd.AND.endotp=>newPerBeg
                      daycx=daycx+(endotp-newPerBeg)+1
                 CASE endotp>newPerEnd.AND.begOtp<=newPerEnd.AND.begOtp>newPerBeg
                      daycx=daycx+(endotp-newPerBeg)+1     
              ENDCASE
           ENDIF           
      ENDSCAN
      IF daycx>14
         dayCost=daycx-14
         frmord.tboxdolj2.Refresh
      ENDIF
      SET ORDER TO &ordOtpOld
      SELECT curPeopOrder   
   ENDIF
           
   frmOrd.tBoxPerBeg.Refresh
   frmOrd.tBoxPerEnd.Refresh
   SELECT curPeopOrder  
ENDIF
IF logDatJob 
   ON ERROR DO erSup
   SELECT curOrdJob
   ZAP    
   APPEND FROM datjob FOR kodpeop=newkodpeop
   SELECT curOrdJob
   DO CASE
      CASE newKodprik=106
           DELETE FOR tr#2
      CASE newKodprik=107
           DELETE FOR tr#4
      CASE newKodPrik=115
           DELETE FOR tr#5      
   ENDCASE   
   REPLACE npord WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.name,sprpodr.name) ,ndord WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namer,sprdolj.name) ALL
   IF logNewRec
      DELETE FOR !EMPTY(dateOut) 
      GO TOP 
   ELSE 
      DELETE FOR !EMPTY(dateOut).AND.dateOut<dateOut
      GO TOP   
   ENDIF    
   newSkp=kp
   newSPodr=ALLTRIM(npord)
   newSKd=kd
   newSDolj=ALLTRIM(ndord)
   newKse=kse  
   oldSovmJob=IIF(logNewRec,nid,oldnid)   
   nidJobNew=curOrdJob.nid 
   newTr=curOrdJob.tr     
   strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
   frmOrd.tBoxPodrNew.Refresh
   frmOrd.comboBox5.ControlSource='newSDolj'
   frmOrd.comboBox5.SetFocus     
   frmOrd.SpinKse.SetFocus
   ON ERROR 
   SELECT curpeoporder
ENDIF 
DO CASE
   CASE logNewRec.AND.INLIST(newKodPrik,18,19)
        newPkont=curSuplPeop.pkont
        newddop=curSuplPeop.daykont
        newNkont=curSuplPeop.numdog
        newDkont=curSuplPeop.ddog
        newSrok=curSuplPeop.ktime
        strSrok=curSuplPeop.strTime
        newDateBeg=IIF(!EMPTY(curSuplPeop.endDog),curSuplPeop.endDog+1,CTOD('  .  .    '))
        IF !EMPTY(newDateBeg).AND.BETWEEN(newSrok,1,5)
           newDateEnd=CTOD(LEFT(DTOC(newDateBeg),6)+STR(YEAR(newDateBeg)+newSrok,4))-1
        ENDIF    
ENDCASE
frmOrd.Refresh
ON ERROR DO erSup
&objFocus..SetFocus
ON ERROR 
************************************************************************************************************************          
PROCEDURE lostFocusPeop
WITH frmOrd  
     ON ERROR DO erSup  
     .listBox1.Visible=.F. 
      newFioPeop=IIF(SEEK(newKodPeop,'people',1),people.fio,'')                      
     .tBoxFio.controlSource='newFioPeop'
     .tBoxFio.Refresh 
     ON ERROR  
ENDWITH
************************************************************************************************************************
PROCEDURE padejInOrder
DO CASE
   CASE parPadejDol=1
        newNDolj=IIF(!EMPTY(sprdolj.namer),sprdolj.namer,newNDolj)
   CASE parPadejDol=2
        newNDolj=IIF(!EMPTY(sprdolj.named),sprdolj.named,newNDolj)                                
   CASE parPadejDol=3
        newNDolj=IIF(!EMPTY(sprdolj.namet),sprdolj.namet,newNDolj)     
ENDCASE 
************************************************************************************************************************
PROCEDURE addObjectOrder
WITH frmOrd
     .comboBox1.Enabled=IIF(par1,.T.,.F.)
     IF typeOrdNew=2
       .comboBox99.Enabled=IIF(par1,.T.,.F.)      
     ENDIF 
     DO adTboxAsCont WITH 'frmOrd','ordTabn',.ordPrik.Left,.ordPrik.Top+.ordPrik.Height-1,.ordPrik.Width,dHeight,'?????',1,1
     DO adTboxNew WITH 'frmOrd','tBoxTabn',.ordTabn.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newKodPeop',.F.,IIF(par1,.T.,.F.),0,'99999','DO validOrderNum'  
     .tBoxTabn.Alignment=0
     DO adTboxAsCont WITH 'frmOrd','ordFio',.ordPrik.Left,.ordTabn.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ??? ????????',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFio',.ordFio.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioPeop',.F.,IIF(par1,.T.,.F.),0      
      .tBoxFio.procforChange='DO changePeopOrd'   
     DO adtboxnew WITH 'frmOrd','boxFree',.tBoxFio.Top,.tBoxFio.Left+.tBoxFio.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlnt',.tBoxFio.Left+.tBoxFio.Width+1,.tBoxFio.Top+2,'','sbdn.ico','DO selectPeopOrd',.tBoxFio.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlnt.Enabled=IIF(par1,.T.,.F.)
ENDWITH      
************************************************************************************************************************
PROCEDURE addObjectOrder2
PARAMETERS topEditBox
WITH frmOrd
     .AddObject('editOrder','MyEditBox')      
     WITH .editOrder
          .Visible=.T.          
          .ControlSource='txtOrder.txtPrn'
          .Left=.Parent.ordTabn.Left
          .Width=.Parent.ordTabn.Width+.Parent.comboBox1.Width-1
          .Top=&topEditBox
          .Height=.Parent.butSaveRec.Top-.Top-10 
          .Enabled=IIF(logRead,.T.,.F.)  
     ENDWITH      
     .butTxt.Visible=IIF(par1,.T.,.F.)
     .butSaveRec.Visible=IIF(par1,.T.,.F.)
     .butRetRead.Visible=IIF(par1,.T.,.F.) 
     
     DO addListBoxMy WITH 'frmOrd',1,.tBoxFio.Left,.tBoxFio.Top+dHeight-1,300,.combobox1.Width  
     WITH .listBox1          
          .RowSource='curSuplPeop.fio'                                   
          .RowSourceType=2
          .ColumnCount=1         
          .Visible=.F. 
          .Height=.Parent.Height-.Top         
          .procForDblClick='DO validListPeop'
          .procForLostFocus='DO lostFocusPeop'
     ENDWITH 
ENDWITH      
*****************************************************************************************************************
*PROCEDURE rightClickPeopOrd
*****************************************************************************************************************
*               ????? ??? ?????? ???????
*****************************************************************************************************************
PROCEDURE formPrnOrder
fSupl=CREATEOBJECT('FORMSUPL')
logform=.F.
logKmMandat=.T. &&??? ?????? ???????????????? ?????????????
logKmOrder=.T.  &&??? ?????? ??????? ?? ?????????????
WITH fSupl
     .Caption='?????? ???????'
     DO adSetupPrnToForm WITH 20,20,400,.T.,.F.
     logWord=.T.
     
     IF newKodprik=201
        DO adCheckBox WITH 'fSupl','checkOrd','??????',.Shape91.Top+.Shape91.Height+10,.Shape91.Left,150,dHeight,'logKmOrder',0           
        DO adCheckBox WITH 'fSupl','checkMandat','?????????????',.checkOrd.Top,.Shape91.Left,150,dHeight,'logKmMandat',0           
        .checkOrd.Left=.Shape91.Left+(.Shape91.Width-.checkOrd.Width-.checkMandat.Width-20)/2 
        .checkMandat.Left=.checkOrd.Left+.checkOrd.Width+20
        
        DO adCheckBox WITH 'fSupl','check1','??????? ????? ???? Word',.checkOrd.Top+.checkOrd.Height+10,.Shape91.Left,150,dHeight,'logForm',0           
     ELSE 
        DO adCheckBox WITH 'fSupl','check1','??????? ????? ???? Word',.Shape91.Top+.Shape91.Height+10,.Shape91.Left,150,dHeight,'logForm',0           
     ENDIF 
    .check1.Left=.Shape91.Left+(.Shape91.Width-.check1.Width)/2 
    
    
     DO addButtonOne WITH 'fSupl','butPrn',.Shape91.Left+(.Shape91.Width-RetTxtWidth('w????????w')*3-20)/2,.check1.Top+.check1.Height+20,'??????','','DO prnOrder WITH .T.',39,RetTxtWidth('w????????w'),'?????? ???????' 
     DO addButtonOne WITH 'fSupl','butView',.butPrn.Left+.butPrn.Width+10,.butPrn.Top,'????????','','DO prnOrder WITH .F.',39,.butPrn.Width,'????????' 
     DO addButtonOne WITH 'fSupl','butRet',.butView.Left+.butView.Width+10,.butPrn.Top,'???????','','fSupl.Release',39,.butPrn.Width,'???????' 
     .Width=.Shape91.Width+40
     .Height=.butPrn.Height+.butPrn.Top+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*****************************************************************************************************************
*                           ??????????????? ?????? ???????
*****************************************************************************************************************
PROCEDURE prnOrder
PARAMETERS par1
strDatePrik=IIF(!EMPTY(datorder.dateord),dateToString('datorder.dateord',.T.),'')
SELECT curpeoporder
IF par1
   DO procForPrintAndPreview WITH 'repOrder','??????',.T.,'orderToWord'
ELSE 
   DO procForPrintAndPreview WITH 'repOrder','??????',.F.
ENDIF 
******************************************************************************************************************
PROCEDURE orderToWord   
IF EMPTY(datorder.pathor)
   pathOrdWord=ALLTRIM(datset.pathword)+LTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)+'_'+STR(YEAR(datorder.dateord),4)+'.doc'
ELSE    
   pathOrdWord=ALLTRIM(datorder.pathor)
ENDIF
#DEFINE wdWindowStateMaximize 1

#DEFINE wdBorderTop -1           &&??????? ??????? ?????? ???????
#DEFINE wdBorderLeft -2          &&????? ??????? ?????? ???????
#DEFINE wdBorderBottom -3        &&?????? ??????? ?????? ???????
#DEFINE wdBorderRight -4         &&?????? ??????? ?????? ???????
#DEFINE wdBorderHorizontal -5    &&?????????????? ????? ???????
#DEFINE wdBorderVertical -6      &&?????????????? ????? ???????
#DEFINE wdLineStyleSingle 1      && ????? ????? ??????? ?????? (? ????? ?????? ???????)
#DEFINE wdLineStyleNone 0        && ????? ???????????
#DEFINE wdAlignParagraphRight 2


IF FILE(pathOrdWord).AND.!logForm   
   objWord=CREATEOBJECT('WORD.APPLICATION')
   nameDoc=objWord.Documents.Open(pathOrdWord) 
   objWord.WindowState=wdWindowStateMaximize
   objWord.Visible=.T. 
ELSE
   pathOrdWord=ALLTRIM(datset.pathword)+LTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)+'_'+STR(YEAR(datorder.dateord),4)+'.doc'
   objWord=CREATEOBJECT('WORD.APPLICATION')
   #DEFINE cr CHR(13)  
   nameDoc=objWord.Documents.Add()  
   nameDoc.ActiveWindow.View.ShowAll=0   
   objWord.WindowState=wdWindowStateMaximize     
   objWord.Selection.pageSetup.Orientation=0
   objWord.Selection.pageSetup.TopMargin=10
   objWord.Selection.pageSetup.BottomMargin=10
   docRef=GETOBJECT('','word.basic')
   strDatePrik=IIF(!EMPTY(datorder.dateord),dateToString('datorder.dateord',.T.),'')
   WITH docRef
        .Insert(cr)
        .Font('Times New Roman',14)
        .LeftPara 
        .Insert('?????????? ???????????????')
        docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ?????? 
        .Insert(cr)
        .Font('Times New Roman',14)
        .LeftPara 
        .Insert('????????? ???????????')
         docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ?????
        .Insert(cr)
        .Font('Times New Roman',14)
        .LeftPara 
        .Insert('????????? ????????')
        docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ?????
        .Insert(cr)     
        .Insert(cr)
        .Font('Times New Roman',14)
        .LeftPara
        .Insert('??????')
        .LeftPara        
        .Insert(cr)
        .Insert(cr)
        nameDoc.Tables.add(objWord.Selection.range,1,2)        
        ordTable1=nameDoc.Tables(1)      
        WITH ordTable1
             .Columns(1).Width=240
             .Columns(2).Width=240    
             .cell(1,1).Range.Text=strDatePrik+' ? '+ALLTRIM(STR(datorder.numord))+'-'+datorder.strord
             .cell(1,1).Range.Select                   
             docRef.LeftPara  && ???????????? ?????
             docRef.Font('Times New Roman',14)
             docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ??????           
             .cell(1,2).Range.Text='?. ?????' 
             .cell(1,2).Range.Select
             docRef.RightPara  && ???????????? ?????
             docRef.Font('Times New Roman',14) 
             docRef.CloseParaAbove     
             docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ??????            
             docRef.LineDown               
        ENDWITH     
        .Insert(cr)     
        .ClearFormatting        
        .LeftPara
        .Font('Times New Roman',14)
        .Insert('??????????:')
        .LeftPara
        SELECT curpeoporder
        COUNT TO maxRows
        SCAN ALL
             .Insert(cr)
             .Font('Times New Roman',14)          
             .JustifyPara
             .Insert(LTRIM(STR(curPeopOrder.npp))+'. '+txtord)            
             .CloseParaAbove     
             .CloseParaBelow  &&??????? ?????? ???????? ????? ?????? 
             .Insert(cr)    
        ENDSCAN
        .Insert(cr)
        .Font('Times New Roman',14)
        .Insert(SPACE(7)+ALLTRIM(boss.dolboss)+SPACE(40)+ALLTRIM(boss.fioboss))
        .LeftPara  
       
         namedoc.Sections.Add
        .InsertNewPage &&????????? ????? ????????
        .Insert(cr)
        .Insert(cr)
   
        nameDoc.Tables.add(objWord.Selection.range,1,4)        
        ordTable2=nameDoc.Tables(2)      
        WITH ordTable2
             .Columns(1).Width=160
             .cell(1,1).Select
             docRef.CenterPara 
             docRef.Font('Times New Roman',12) 
             .cell(1,1).Range.Text='?????????'
             .Columns(2).Width=120  
             .Cell(1,2).Select
             docRef.CenterPara  
             docRef.Font('Times New Roman',12)   
             .cell(1,2).Range.Text='???????'        
             .Cell(1,3).Select
             docRef.CenterPara  
             docRef.Font('Times New Roman',12)   
             .cell(1,3).Range.Text='?.?.?.'
             .Columns(3).Width=120        
             .Cell(1,4).Select
             docRef.CenterPara  
             docRef.Font('Times New Roman',12)   
             .cell(1,4).Range.Text='????'        
             .Columns(4).Width=80
             .Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderVertical).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderRight).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderLeft).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderBottom).LineStyle=wdLineStyleSingle                       
             docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ??????           
             .Rows.Add  
             .cell(2,1).Range.Text='????????? ??' 
             .cell(2,1).Select
             docRef.LeftPara 
             .cell(2,3).Range.Text='?.?.???????' 
             .cell(2,4).Select
             docRef.LeftPara         
             .Rows.Add  
             .cell(3,1).Range.Text='??????? ????????????' 
             .cell(3,1).Range.Select
             .cell(3,3).Range.Text='?.?.????????' 
             .cell(3,4).Range.Select
             docRef.RightPara  
             docRef.Font('Times New Roman',14) 
             docRef.CloseParaAbove     
             docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ??????            
             docRef.LineDown     
        ENDWITH                     
        IF datorder.typeord=3.AND.maxRows>0
           .Insert(cr)                    
           .Font('Times New Roman',14)
           .Insert('??????????, ????????:')  
           .Insert(cr) 
           nameDoc.Tables.add(objWord.Selection.range,1,4)        
           ordTable3=nameDoc.Tables(3)          
           WITH ordTable3
                .Columns(1).Width=180
                .Columns(2).Width=100 
                .Columns(3).Width=120         
                .Columns(4).Width=80                 
                .cell(1,1).Select
                docRef.CenterPara 
                docRef.Font('Times New Roman',12) 
                .cell(1,1).Range.Text='?????????'                              
                .Cell(1,2).Select              
                docRef.CenterPara  
                docRef.Font('Times New Roman',12)   
                .cell(1,2).Range.Text='???????'      
                .Cell(1,3).Select
                docRef.CenterPara  
                docRef.Font('Times New Roman',12)   
                .cell(1,3).Range.Text='?.?.?.'                             
                .Columns(4).Width=80
                .Cell(1,4).Select              
                docRef.CenterPara  
                docRef.Font('Times New Roman',12)   
                .cell(1,4).Range.Text='????'       
                .Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
                .Borders(wdBorderVertical).LineStyle=wdLineStyleSingle 
                .Borders(wdBorderRight).LineStyle=wdLineStyleSingle 
                .Borders(wdBorderLeft).LineStyle=wdLineStyleSingle 
                .Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
                .Borders(wdBorderBottom).LineStyle=wdLineStyleSingle                       
                FOR i=1 TO maxRows
                    .Rows.Add 
                    .cell(i+1,4).Select 
                    docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ??????           
                    docRef.CloseParaAbove     
                    docRef.CloseParaBelow  &&??????? ?????? ???????? ????? ??????           
                ENDFOR                      
                docRef.LineDown    
           ENDWITH                  
        ENDIF  
        .Insert(cr)     
        .ClearFormatting        
   ENDWITH   
   namedoc.saveAs(pathOrdWord)  
   SELECT datOrder
   REPLACE pathOr WITH pathOrdWord
   objWord.Visible=.T.       
ENDIF 
IF logKmMandat 
   SELECT curpeoporder
   GO TOP 
   LOCAL loWord, loDoc 
 
   ***** ????, ??????????? ? ??????? ?????????
   * dkont-???? ??????????
   * fio-??? ? ???????????? ??????
   * fio2-??? ? ???????????? ??????
   * dol-?????????
   * podr-?????????????
   * srok-????
   * period-?????? ? ??, ??????????? ? ??????? -  ? "01" ????? 1999?. ?? "01" ????? 1999?.

   objWord=CREATEOBJECT('WORD.APPLICATION')
   pathdot=ALLTRIM(datset.pathword)+'kmmandat.dot'       
   nameDoc=objWord.Documents.Add(pathdot)   
 
   
   * ??????????? ??????????? ???????? ? ????  
   IF TYPE([nameDoc.formFields("cpFio")])="O"
      nameDoc.FormFields("cpFio").Result=ALLTRIM(curpeoporder.fiopeop)
   ENDIF
   IF TYPE([nameDoc.formFields("cJob")])="O"
      nameDoc.FormFields("cJob").Result=LOWER(ALLTRIM(curpeoporder.ndolj))+' '+LOWER(ALLTRIM(curpeoporder.npodr))+' '+ALLTRIM(boss.office)
   ENDIF
   IF TYPE([nameDoc.formFields("cCity")])="O"
      nameDoc.FormFields("cCity").Result=ALLTRIM(varsupl2)+' '+ALLTRIM(LEFT(varsupl,100))     
       
   ENDIF
   IF TYPE([nameDoc.formFields("nDays")])="O"
      nameDoc.FormFields("nDays").Result=LTRIM(STR(curpeoporder.dayotp))
   ENDIF
   IF TYPE([nameDoc.formFields("cAim")])="O"
      nameDoc.FormFields("cAim").Result=ALLTRIM(curpeoporder.npodr2)
   ENDIF
   
   IF TYPE([nameDoc.formFields("cPrik")])="O"
      nameDoc.FormFields("cPrik").Result='"'+STR(DAY(curpeoporder.dord),2)+'" '+ALLTRIM(month_prn(MONTH(curpeoporder.dord)))+' '+STR(YEAR(curpeoporder.dord),4) +' ?.  ? '+ALLTRIM(curpeoporder.nord)
   ENDIF


  objWord.Visible=.T.
ENDIF 
*************************************************************************************************************************
*                   ????????? ??????? ??? ?? ??????? ? ????????
*************************************************************************************************************************
PROCEDURE procFioIni
PARAMETERS par1
*par1 - ????????
str_fio=&par1  
str_ini=ALLTRIM(LEFT(str_fio,AT(' ',str_fio)))
str_fio=ALLTRIM(SUBSTR(str_fio,AT(' ',str_fio)))
str_ini=str_ini+' '+LEFT(str_fio,1)+'.'
str_fio=ALLTRIM(SUBSTR(str_fio,AT(' ',str_fio)))
str_ini=str_ini+LEFT(str_fio,1)+'.'    
*************************************************************************************************************************
*                              ????????? ????????????????
*************************************************************************************************************************
PROCEDURE procSovmIn
PARAMETERS par1,par2,parPers,paruvol
nuvol=paruvol
logDek=par2
logPers=parPers
parPadejDol=IIF(paruvol=1,2,1)
parPadejFio=IIF(paruvol=1,2,1)
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot

IF parPers
   newKse=IIF(logNewRec,0.00,curpeoporder.kse)
   newTr=IIF(logNewRec,4,curpeoporder.tr)
ELSE 
   newKse=IIF(logNewRec,0.25,curpeoporder.kse)
   newTr=IIF(logNewRec,2,curpeoporder.tr)
ENDIF    
IF nuvol=2
   newTr=IIF(logNewRec,5,curpeoporder.tr)   
ENDIF
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
objFocus='frmOrd.tBoxBeg'
newLogApp=.T.
newNpp=0
newSKp=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,1,3)))
newSPodr=IIF(logNewRec,0,SUBSTR(curPeoporder.varSupl,4,100))
newSKd=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,104,3)))
newSDolj=IIF(logNewRec,0,SUBSTR(curPeoporder.varSupl,107,100))
newTabZam=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl2,1,5)))
newFioZam=IIF(logNewRec,'',SUBSTR(curPeoporder.varSupl2,6,60))
logDatJob=.F.
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder    
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
             
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (??????.)',1,1     
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNPodr',.F.,IIF(par1,.T.,.F.),0  
       
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (??????.)',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNdolj',.F.,IIF(par1,.T.,.F.),0  
     
     
     DO adTboxAsCont WITH 'frmOrd','ordPodrNew',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (???????.)',1,1 
     DO addcombomy WITH 'frmOrd',4,.comboBox1.Left,.ordPodrNew.Top,dHeight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSPodr','ALLTRIM(curSprPodr.name)',6,'','DO procSovmPodr',.F.,.T.  
     .comboBox4.Visible=IIF(par1,.T.,.F.)
     .comboBox4.DisplayCount=17
     DO adTboxNew WITH 'frmOrd','tBoxPodrNew',.ordPodrNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSPodr',.F.,.F.,0  
     .tBoxPodrNew.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordPodrNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (???????.)',1,1 
     DO addComboMy WITH 'frmOrd',5,.comboBox1.Left,.ordDolNew.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSDolj','ALLTRIM(curDolPodr.name)',6,.F.,'DO procSovmDolj',.F.,.T.       
     WITH .comboBox5         
          .DisplayCount=15
          .ColumnCount=3
          .ColumnWidths='0,50,500'
          .RowSource="curDolPodr.name,strVac,name"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH    
      DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSDolj',.F.,.F.,0  
     .tBoxDoljNew.Visible=IIF(par1,.F.,.T.)   
     
     
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=IIF(par1,.T.,.F.)   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1                                             
     DO addComboMy WITH 'frmOrd',11,.ordTr.Left+.ordTr.Width-1,.ordTr.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordTr.Width+2,IIF(par1,.T.,.F.),'strType','curSprType.name',6,.F.,'newTr=curSprType.kod',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.) 
     
     IF logDek
        DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'??? (???? ????????)',1,1      
        DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioZam',.F.,IIF(par1,.T.,.F.),0      
        .tBoxFioNew.procforChange='DO changePeopDek'   
        DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
        DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectPeopDek',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
        .butKlntNew.Enabled=IIF(par1,.T.,.F.)
     ENDIF    
                    
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,IIF(!logDek,.ordTr.Top+dHeight-1,.ordFioNew.Top+dHeight-1),.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0 
                
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
     IF logDek
        DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
        WITH .listBox2
             .RowSource='curSuplpeop.fio'         
             .RowSourceType=2
             .Visible=.F.  
             .Height=.Parent.Height-.Top      
             .procForDblClick='DO validListDek'
             .procForLostFocus='DO lostFocusPeop2'
        ENDWITH 
     ENDIF
     IF par1       
        ON ERROR DO erSup
        .comboBox4.SetFocus
        .comboBox5.SetFocus
        .comboBox11.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH
************************************************************************************************************************          
PROCEDURE lostFocusPeop2
WITH frmOrd  
     ON ERROR DO erSup  
     .listBox2.Visible=.F.    
     ON ERROR  
ENDWITH
***********************************************************************************************************************
PROCEDURE selectPeopDek
SELECT curSuplPeop
ZAP
APPEND FROM people
WITH frmOrd
     .listBox2.RowSource='curSuplPeop.fio'                     
      IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.  
        .listBox2.SetFocus            
     ENDIF 
ENDWITH 
*****************************************************************************************************************
PROCEDURE changePeopDek
WITH frmOrd 
     IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=frmOrd.tBoxFioNew.Text 
SELECT curSuplPeop
ZAP
APPEND FROM people FOR LEFT(LOWER(fio),LEN(ALLTRIM(lcValue)))=LOWER(ALLTRIM(lcValue))
WITH frmOrd.listBox2
     .RowSource='curSuplPeop.fio'                    
     .Visible=IIF(RECCOUNT('curSuplPeop')=0,.F.,.T.)      
ENDWITH 
************************************************************************************************************************
PROCEDURE validListDek
*padejFio=sprorder.fiop
newFioZam=curSuplPeop.fior
newTabZam=curSuplPeop.num
WITH frmOrd
     .tBoxFioNew.ControlSource='newFioZam'
     .listBox2.Visible=.F.
     .tBoxFioNew.Refresh
     .Refresh
 ENDWITH 

*************************************************************************************************************************
PROCEDURE saveRecSovmIn
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textSovmIn
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp
       REPLACE varSupl WITH PADR(ALLTRIM(STR(newSKp)),3,' ')+PADR(ALLTRIM(newSpodr),100,' ')+PADR(ALLTRIM(STR(newSKd)),3,' ')+PADR(ALLTRIM(newSdolj),100,' '),varSupl2 WITH PADR(ALLTRIM(STR(newTabZam)),5,' ')+PADR(ALLTRIM(newFioZam),60,' '),nid WITH newNid,;
       dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd 
DO saveAdmOrder             
*************************************************************************************************************************
PROCEDURE textSovmIn
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
cordFio=''
cOrdDol=''
cOrdPodr=''
DO procOrdNameDol

cFioZam=''
cPodrZam=''
cDolZam=''
DO procOrdNameDolZam WITH 'newFioZam','newSDolj','newSPodr',IIF(nuvol=2,3,1),4
str_ini=''
DO procfioini WITH 'cFioZam'
str_zam=str_ini

str_ini=''
DO procfioini WITH 'cOrdFio' 

SELECT txtOrder
IF logPers
   REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
           ' ?????????? ??????? ? ??????? '+ALLTRIM(STR(newKse*100))+'% ?????? ?? ??????????? ???? ???????????? ?? ????????? '+;
           LOWER(ALLTRIM(cDolZam))+' '+LOWER(ALLTRIM(cPodrZam))+' ??? ???????????? ?? ????? ???????? ?????? ? ??????? ????????????? ????????????????? ????????????????? ???????? ???'+;
           IIF(logDek,' ?? ????? ?????????? ????????? ????????? '+str_zam+', ???????????? ? ?????????? ??????? "?? ????? ?? ???????? ?? ?????????? ?? ???????? 3-? ???".','.')+CHR(13)+'?????????: '+ALLTRIM(newOsnov)        
ELSE 
   IF nuvol=2
           REPLACE txtprn WITH ALLTRIM(str_ini)+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? '+strDateBeg+;
                   ' ??????? ?? ???????????????? ?? '+STR(newKse,4,2)+' ??????? ??????? '+LOWER(cDolZam)+' '+LOWER(cPodrZam)+;
                   ', ? ??????? ???????? ???????? ??????????.'+CHR(13)+'?????????: '+ALLTRIM(newOsnov)      
   ELSE
      REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
              ' ????????? ?????? ????? ????????????????? ???????? ??????? ?? ???????? ?????? ? ???????? 900 ????? ? ??? ?? '+STR(newKse,4,2)+;
              ' ????????? '+LOWER(ALLTRIM(cDolZam))+' '+LOWER(ALLTRIM(cPodrZam))+' ?? ????????? ??????? ??????, ? ??????? ?? ?????????? ???????????? ?????'+;
              IIF(logDek,' ?? ????? ?????????? ????????? ????????? '+str_zam+', ???????????? ? ?????????? ??????? "?? ????? ?? ???????? ?? ?????????? ?? ???????? 3-? ???"','.')+CHR(13)+'?????????: '+ALLTRIM(newOsnov)        
   ENDIF        
ENDIF         
********************************************************************************************************************************
PROCEDURE procSovmPodr
newSKp=curSprPodr.kod  
newSPodr=curSprPodr.name

SELECT datJob
SET ORDER TO 2
SELECT curDolPodr
SET FILTER TO kp=newSKp
SCAN ALL    
     ksesup=0 
     DO sumVacKse
     SELECT curDolPodr
     REPLACE strVac WITH IIF(kse-ksesup=0,'',STR(kse-ksesup,6,2))
ENDSCAN

frmOrd.ComboBox5.RowSource='curDolPodR.name'
WITH .comboBox5         
     .DisplayCount=15
     .ColumnCount=3
     .ColumnWidths='0,50,500'
     .RowSource="curDolPodr.name,strVac,name"
ENDWITH 
frmOrd.ComboBox5.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
frmOrd.ComboBox5.RowSourceType=6
frmOrd.comboBox5.ProcForValid='DO procSovmDolj'
KEYBOARD '{TAB}'

********************************************************************************************************************************
PROCEDURE procSovmDolj
newSKd=curDolPodr.kd  
newSDolj=curDolPodr.name
frmOrd.ComboBox5.ControlSource='newSDolj'
frmOrd.comboBox5.Refresh
KEYBOARD '{TAB}'
*************************************************************************************************************************
*                         ?????? ??????????? ????????????????
*************************************************************************************************************************
PROCEDURE procSovmOut
PARAMETERS par1,parPers,parUvol
logPers=parPers
parPadejDol=1
parPadejFio=1
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
SELECT curPeopOrder
newDateBeg=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.dateBeg)
newTr=IIF(logNewRec,IIF(logPers,4,2),curpeoporder.tr)
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
objFocus='frmOrd.tBoxBeg'
newLogApp=.T.
nuvol=paruvol
newNpp=0
newSKp=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,1,3)))
newSPodr=IIF(logNewRec,0,SUBSTR(curPeoporder.varSupl,4,100))
newSKd=IIF(logNewRec,0,VAL(SUBSTR(curPeoporder.varSupl,104,3)))
newSDolj=IIF(logNewRec,0,ALLTRIM(SUBSTR(curPeoporder.varSupl,107,100)))

oldsovmjob=IIF(logNewRec,0,curPeoporder.oldnid)

newSLink=ALLTRIM(curSprOrder.slink)
newPerBeg=IIF(logNewRec,CTOD('  .  .    '),curPeopOrder.perBeg)
newPerEnd=IIF(logNewRec,CTOD('  .  .    '),curPeopOrder.perEnd)
newDayKomp=IIF(logNewRec,0,curpeoporder.DayOtp)
nidJobNew=IIF(logNewRec,0,curPeopOrder.nidJob)
logDatJob=.T.

SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     .comboBox1.Enabled=IIF(par1,.T.,.F.)
     
     DO adTboxAsCont WITH 'frmOrd','ordFio',.ordPrik.Left,.ordPrik.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ??? ????????',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFio',.ordFio.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioPeop',.F.,IIF(par1,.T.,.F.),0      
      .tBoxFio.procforChange='DO changePeopOrd'   
     DO adtboxnew WITH 'frmOrd','boxFree',.tBoxFio.Top,.tBoxFio.Left+.tBoxFio.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlnt',.tBoxFio.Left+.tBoxFio.Width+1,.tBoxFio.Top+2,'','sbdn.ico','DO selectPeopOrd',.tBoxFio.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlnt.Enabled=IIF(par1,.T.,.F.)
     
     DO adTboxAsCont WITH 'frmOrd','ordTabn',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1
     DO adTboxNew WITH 'frmOrd','tBoxTabn',.ordTabn.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newKodPeop',.F.,IIF(par1,.T.,.F.),0,'99999','DO validOrderNum'  
     .tBoxTabn.Alignment=0
               
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordTabn.Top+dHeight-1,.ordPrik.Width,dHeight,IIF(nuvol=3,'??????? ?','?????????? ?'),1,1   
    * DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordTabn.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0      
             
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (??????.)',1,1     
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNPodr',.F.,IIF(par1,.T.,.F.),0  
       
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (??????.)',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNdolj',.F.,IIF(par1,.T.,.F.),0  
                   
     DO adTboxAsCont WITH 'frmOrd','ordPodrNew',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (???????.)',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodrNew',.ordPodrNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSPodr',.F.,.T.,0  
     .tBoxPodrNew.Visible=.T.  
     
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordPodrNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (???????.)',1,1 
     DO addComboMy WITH 'frmOrd',5,.comboBox1.Left,.ordDolNew.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newSDolj','ALLTRIM(curOrdJob.ndord)',6,.F.,'DO procSovmSelect',.F.,.T.       
     WITH .comboBox5         
          .DisplayCount=15
          .ColumnCount=4
          .ColumnWidths='0,50,250,250'
          .RowSource="curOrdJob.ndord,kse,ndord,npord"
          .Visible=IIF(par1,.T.,.F.)
     ENDWITH     
      DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSDolj',.F.,.F.,0  
     .tBoxDoljNew.Visible=IIF(par1,.F.,.T.)   
     
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,1.5 
     .spinKse.Enabled=.T.   
     DO adTBoxAsCont WITH 'frmOrd','ordTr',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w???w'),dHeight,'???',2,1  
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordTr.Top,.ordTr.Left+.ordTr.Width-1,.comboBox1.Width-.spinKse.Width-.ordTr.Width+1,dHeight,'strType',.F.,.F.,0  
     .tBoxTr.Visible=.T.                
         
     IF nuvol=3
       * DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'??????',1,1  
       * DO adTboxNew WITH 'frmOrd','tBoxLink',.ordKse.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newsLink',.F.,.F.,0  
     
        DO adTboxAsCont WITH 'frmOrd','ordKomp',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'???????????',1,1  
        DO adTboxNew WITH 'frmOrd','tBoxKomp',.ordKomp.Top,.comboBox1.Left,.comboBox1.Width/5,dHeight,'newDayKomp','Z',IIF(par1,.T.,.F.),0 
        .tBoxKomp.InputMask='99' 
        .tBoxKomp.Alignment=0
        DO adTboxAsCont WITH 'frmOrd','ordPerBeg',.tBoxKomp.Left+.tBoxKomp.Width-1,.ordKomp.Top,.tBoxKomp.Width,dHeight,'?????? ?',1,1  
        DO adTboxNew WITH 'frmOrd','tBoxPerBeg',.ordKomp.Top,.ordPerBeg.Left+.ordPerBeg.Width-1,.tBoxKomp.Width,dHeight,'newPerBeg',.F.,IIF(par1,.T.,.F.),0 
        DO adTboxAsCont WITH 'frmOrd','ordPerEnd',.tBoxPerBeg.Left+.tBoxPerBeg.Width-1,.ordKomp.Top,.tBoxKomp.Width,dHeight,'?????? ??',1,1  
        DO adTboxNew WITH 'frmOrd','tBoxPerEnd',.ordKomp.Top,.ordPerEnd.Left+.ordPerEnd.Width-1,.comboBox1.Width-.tBoxKomp.Width*4+4,dHeight,'newPerEnd',.F.,IIF(par1,.T.,.F.),0 
     ENDIF
         
     
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,IIF(nuvol=3,.ordKomp.Top+dHeight-1,.ordTr.Top+dHeight-1),.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0 
       
     DO addObjectOrder2 WITH 'frmOrd.ordosnov.Top+frmOrd.ordosnov.Height+10' 
     IF par1     
        ON ERROR DO erSup
        .comboBox5.SetFocus   
        .tBoxFio.SetFocus   
        ON ERROR      
     ENDIF
ENDWITH
************************************************************************************************************************
PROCEDURE procSovmSelect
newSkp=curOrdJob.kp
newSPodr=curOrdJob.npord
newSKd=curOrdJob.kd
newSDolj=curOrdJob.ndord
oldSovmJob=curOrdJob.nid
nidJobNew=curOrdJob.nid
newKse=curOrdJob.kse
newTr=curOrdJob.tr
WITH frmOrd
     .tBoxPodrnew.ControlSource='newSPodr'
     .spinKse.ControlSource='newKse'
     .Refresh
ENDWITH 

*************************************************************************************************************************
PROCEDURE saveRecSovmOut
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textSovmOut
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF

*   REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
*        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newkse,supord WITH newKodprik,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
*        dayOtp WITH newDayKomp,perBeg WITH newPerBeg,perEnd WITH newPerEnd,osnov WITH newOsnov,nid WITH newNid,varSupl WITH newPlace,dord WITH repdorder,nord WITH repnorder 
 
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,tr WITH newTr,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp
        REPLACE varSupl WITH PADR(ALLTRIM(STR(newSKp)),3,' ')+PADR(ALLTRIM(newSpodr),100,' ')+PADR(ALLTRIM(STR(newSKd)),3,' ')+PADR(ALLTRIM(newSdolj),100,' '),nid WITH newNid,oldNid WITH oldSovmJob,;
        dord WITH repdorder,nord WITH repnorder,dayOtp WITH newDayKomp,perBeg WITH newPerBeg,perEnd WITH newPerEnd,sex WITH kodsex      
         
DO saveDimOrd 
DO saveAdmOrder       
*************************************************************************************************************************
PROCEDURE textSovmOut
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')

strPerBeg=IIF(!EMPTY(newPerBeg),dateToString('newPerBeg',.T.),'') 
strPerEnd=IIF(!EMPTY(newPerEnd),dateToString('newPerEnd',.T.),'') 
cOrdFio=''
cOrdDol=''
cOrdPodr=''
DO procOrdNameDol

cFioZam=''
cPodrZam=''
cDolZam=''
DO procOrdNameDolZam WITH 'newFioZam','newSDolj','newSPodr',2,1

str_ini=''
DO procfioini WITH 'cOrdFio' 
SELECT txtOrder
IF nuvol=3
   REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(cOrdFio)+' '+strDateBeg+' ??????? ? ?????? ?? ???????????????? ?? '+LTRIM(STR(newkse,5,2))+' ????????? ?? ??????? ?????????, '+ALLTRIM(newSlink)+CHR(13)+;
           '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
                 strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)                       
          
*   REPLACE txtprn WITH ALLTRIM(newNDolj)+' '+LOWER(ALLTRIM(newNPodr))+' '+ALLTRIM(newFioPeop)+' '+strDateBeg+' ??????? ? ?????? ?? ???????????????? ?? 0,5 ??????? ??????? (???????????? ?? ????? ), ? ????? ? ?????????? ????? ????????? ????????, '+ALLTRIM(newSlink)+CHR(13)+;
 *               '??????????? ????????? ???????? ??????????? ?? '+LTRIM(STR(newDayKomp))+ IIF(newDayKomp>4,' ??????????? ????',IIF(newDayKomp=1,' ??????????? ????',' ??????????? ???'))+' ????????????????? ????????? ??????? ?? ?????? ?????? ? '+;
 *                strPerBeg+' ?? '+strPerEnd+'.'+ CHR(13)+'?????????: '+newOsnov+CHR(13)+'? ???????? '+dim_agree(kodsex)             
ELSE 
   IF logPers
      REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
              ' ?????????? ?? ??????? '+ALLTRIM(STR(newKse*100))+'% ?? ?????????? ?? ????????? '+LOWER(cDolzam)+' '+LOWER(cPodrZam)+'.'
              REPLACE txtprn WITH ALLTRIM(txtprn)+CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
   ELSE 
      REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
              ' ?????????? ?? ?????????? ?????? ????? ????????????????? ???????? ??????? ?? ???????? ?????? ?? '+STR(newKse,4,2)+;
              ' ?????? ?? ????????? '+LOWER(cDolZam)+' '+LOWER(cPodrZam)+'.'
              REPLACE txtprn WITH ALLTRIM(txtprn)+CHR(13)+'?????????: '+ALLTRIM(newOsnov)                     
   ENDIF         
ENDIF 
*************************************************************************************************************************
PROCEDURE proczamest
PARAMETERS par1,par2
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
IF !USED('curPrichZam')
   SELECT kod,name FROM sprtot WHERE sprtot.kspr=30 INTO CURSOR curPrichZam  && ?????? ??? ?????? ???????????????
ENDIF 
IF !USED('curRMat')
   CREATE CURSOR curRMat (kod N(1),name C(75))
   SELECT curRMat
   APPEND BLANK
   APPEND BLANK
   REPLACE name WITH '? ?????? ???????????? ???????????????? ? ??????????? ???? ??????-????????'
 ENDIF
logTxt=par2
parPadejDol=IIF(INLIST(logTxt,3,4),1,2)
parPadejFio=IIF(INLIST(logTxt,3,4),1,2)
CREATE CURSOR  curOrdJob1 FROM ARRAY arOrdJob
IF par1
   SELECT curOrdJob
   ZAP
   APPEND FROM datjob FOR INLIST(tr,1,2,3)
   DELETE FOR !EMPTY(dateOut)
   *REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fior,''),ndOrd WITH STR(kse,4,2)+' '+IIF(SEEK(kd,'sprdolj',1),sprdolj.namer,'') ALL
   IF logTxt=3
      REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,''),ndOrd WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),npOrd WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.nameord,'') ALL
   ELSE 
      REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,''),ndOrd WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),npOrd WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.nameord,'') ALL
   ENDIF 
   GO TOP 
ENDIF
SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr
DO resetVarTot
newKse=IIF(logNewRec,000.00,curPeopOrder.kse)
newFioZam=IIF(logNewRec,SPACE(50),ALLTRIM(SUBSTR(curpeoporder.varsupl,1,60)))
newStrPrich=IIF(logNewRec,SPACE(50),ALLTRIM(SUBSTR(curpeoporder.varsupl,61,60)))
nKse=IIF(logNewRec,SPACE(20),ALLTRIM(SUBSTR(curpeoporder.varsupl,121,20)))
newNKurs=IIF(logNewRec,SPACE(50),ALLTRIM(SUBSTR(curpeoporder.varsupl,141,75)))
newSPodr=IIF(logNewRec,0,SUBSTR(curPeoporder.varSupl2,1,100))
newSDolj=IIF(logNewRec,0,ALLTRIM(SUBSTR(curPeoporder.varSupl2,101,100)))
newTabZam=IIF(logNewRec,0,VAL(ALLTRIM(SUBSTR(curpeoporder.varsupl2,201,5))))
newSkp=IIF(logNewRec,0,VAL(ALLTRIM(SUBSTR(curpeoporder.varsupl2,206,5))))
newSkd=IIF(logNewRec,0,VAL(ALLTRIM(SUBSTR(curpeoporder.varsupl2,211,5))))
newLogApp=.F.
newNpp=0

str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxBeg'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     
     DO CASE   
        CASE logtxt=1
             DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????? ?',1,1   
        CASE logtxt=2
             DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
        CASE logtxt=3
             DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
        CASE logtxt=4
             DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?',1,1                          
         
     ENDCASE
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.
     
     DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'??? (???? ????????)',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioZam',.F.,IIF(par1,.T.,.F.),0      
     .tBoxFioNew.procforChange='DO changePeopZam'   
     DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectPeopZam',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlntNew.Enabled=IIF(par1,.T.,.F.)
     
     DO adTboxAsCont WITH 'frmOrd','ordPodrNew',.ordPrik.Left,.ordFioNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????????? (???)',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodrNew',.ordPodrNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSPodr',.F.,IIF(par1,.T.,.F.),0  
         
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordPodrNew.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? (????)',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSDolj',.F.,IIF(par1,.T.,.F.),0  

     
     DO adTboxAsCont WITH 'frmOrd','ordKse',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1  
     DO addSpinnerMy WITH 'frmOrd','spinKse',.comboBox1.Left,.ordKse.Top,dheight,RetTxtWidth('9999999999'),'newKse',0.25,.F.,0,100
     .spinKse.Enabled=IIF(par1,.T.,.F.)
     IF par2=1
        DO adTboxAsCont WITH 'frmOrd','ordNKse',.spinKse.Left+.spinKse.Width-1,.ordKse.Top,RetTxtWidth('w??/?????.w'),dHeight,'??/????.',1,1
        DO addComboMy WITH 'frmOrd',40,.ordNkse.Left+.ordNkse.Width-1,.ordNkse.Top,dheight,.comboBox1.Width-.spinKse.Width-.ordNkse.Width+2,IIF(par1,.T.,.F.),'nkse','dim_nkse',5,.F.
        .comboBox40.Visible=IIF(par1,.T.,.F.)  
        DO adTboxNew WITH 'frmOrd','tBoxNkse',.ordNkse.Top,.combobox40.Left,.comboBox40.Width,dHeight,'nKse',.F.,.F.,0  
       .tBoxNkse.Visible=IIF(par1,.F.,.T.)           
     ENDIF
     
          
     DO adTboxAsCont WITH 'frmOrd','ordPerBeg',.ordPrik.Left,.ordKse.Top+dHeight-1,.ordPrik.Width,dHeight,'???????',1,1   
     
     DO addComboMy WITH 'frmOrd',11,.comboBox1.Left,.ordPerBeg.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newStrPrich','curPrichZam.name',6,.F.,'newStrPrich=curPrichZam.name',.F.,.T. 
     .comboBox11.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxTr',.ordPerBeg.Top,.comboBox11.Left,.comboBox11.Width,dHeight,'newStrPrich',.F.,.F.,0  
     .tBoxTr.Visible=IIF(par1,.F.,.T.)           
              
     DO adTboxAsCont WITH 'frmOrd','ordPodr4',.ordPrik.Left,.ordPerBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'???.???????????????',1,1 
     DO addComboMy WITH 'frmOrd',23,.comboBox1.Left,.ordPodr4.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNKurs','curRMat.name',6,.F.,'newNKurs=curRMat.name',.F.,.T. 
     .comboBox23.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxPodr4',.ordPodr4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKurs',.F.,IIF(par1,.F.,.T.),0  
     .tBoxPodr4.Visible=IIF(par1,.F.,.T.) 
             
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordPodr4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
     
     DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
     WITH .listBox2
          .RowSource='curOrdJob1.fio,ndord'               
          .RowSourceType=2
          .ColumnCount=2 
          .columnWidths='250,400' 
          .Height=.Parent.Height-.Top       
          .Visible=.F.        
          .procForDblClick='DO validListZam'
          .procForLostFocus='DO lostFocusPeop2'
     ENDWITH                
     IF par1 
        IF par2=1
           .combobox40.SetFocus        
        ENDIF 
        ON ERROR DO erSup              
        .combobox11.SetFocus
        .combobox23.SetFocus
        .tBoxFio.SetFocus 
        ON ERROR   
     ENDIF
ENDWITH
************************************************************************************************************************          
PROCEDURE lostFocusPeop2
WITH frmOrd  
     ON ERROR DO erSup  
     .listBox2.Visible=.F. 
     ON ERROR  
ENDWITH
***********************************************************************************************************************
PROCEDURE selectPeopZam
SELECT curOrdJob1
ZAP
APPEND FROM DBF ('curOrdJob')
WITH frmOrd
     .listBox2.RowSource='curOrdJob1.fio,ndord'                     
      IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.  
        .listBox2.SetFocus            
     ENDIF 
ENDWITH 
*****************************************************************************************************************
PROCEDURE changePeopZam
WITH frmOrd 
     IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=frmOrd.tBoxFioNew.Text 
SELECT curOrdJob1
ZAP
APPEND FROM DBF ('curOrdJob') FOR LEFT(LOWER(fio),LEN(ALLTRIM(lcValue)))=LOWER(ALLTRIM(lcValue))
WITH frmOrd.listBox2
     .RowSource='curOrdJob1.fio,ndord'                    
     .Visible=IIF(RECCOUNT('curOrdJob')=0,.F.,.T.)      
ENDWITH 
************************************************************************************************************************
PROCEDURE validListZam
*padejFio=sprorder.fiop
newTabZam=curOrdJob1.kodpeop
newFioZam=curOrdJob1.fio
newSDolj=curOrdJob1.ndord
newSPodr=curOrdJob1.npord
newSkp=curOrdJob1.kp
newSkd=curOrdJob1.kd
WITH frmOrd
     .tBoxFioNew.ControlSource='newFioZam'
     .tBoxPodrNew.ControlSource='newSPodr'
     .tBoxDoljNew.ControlSource='newSDolj'
     .listBox2.Visible=.F.
     .tBoxFioNew.Refresh
     IF logPerevod>0
        newKse=curOrdJob1.Kse
        newTr=curOrdJob1.Tr
     ENDIF
     .Refresh
     
ENDWITH 
*&objFocus..SetFocus

*************************************************************************************************************************
PROCEDURE saverecZamest
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textZamest
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,kse WITH newKse,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp
        REPLACE varSupl2 WITH PADR(ALLTRIM(newSpodr),100,' ')+PADR(ALLTRIM(newSdolj),100,' ')+STR(newTabZam,5)+STR(newSkp,5)+STR(newSkd,5);
        varSupl WITH PADR(ALLTRIM(newFioZam),60,' ')+PADR(ALLTRIM(newStrPrich),60,' ')+PADR(ALLTRIM(nKse),20,' ')+PADR(ALLTRIM(newNKurs),75,' '),;
        nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd   
DO saveAdmOrder       
*************************************************************************************************************************
PROCEDURE textZamest
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')

cFioZam=''
cPodrZam=''
cDolZam=''
DO procOrdNameDolZam WITH 'newFioZam','newSDolj','newSPodr',4,IIF(logtxt=3,3,4)
str_ini=''
DO procfioini WITH 'cFioZam'
str_zam=str_ini

cOrdFio=''
cOrdDol=''
cOrdPodr=''
DO procOrdNameDol

str_ini=''
DO procfioini WITH 'cOrdFio' 

SELECT txtOrder
DO CASE
   CASE logTxt=1
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
                ' ????????? ?????? ????? ????????????????? ???????? ??????? ?? ???????? ?????? ?? '+STR(newKse,4,2)+IIF(EMPTY(nkse),' ??????? ??????? ',' '+ALLTRIM(nkse)+' ')+LOWER(cDolZam)+' '+LOWER(cPodrZam)+;
                ', ? ??????? ?? ?????????? ???????????? ?????, '+ALLTRIM(newStrPrich)+' '+str_zam+CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
   CASE logTxt=2
        REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPOdr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
                ' ?????????? ??????? ? ??????? '+LTRIM(STR(newKse*100))+'% ?????? ?? ????????? '+LOWER(cDolZam)+' '+LOWER(cPodrZam)+;
                ' ?? ?????? ?????????? ???????????? ???????? ?????????????? ????????? ? ??????? ????????????? ????????????????? ????????????????? ???????? ???, '+;
                ALLTRIM(newStrPrich)+' '+str_zam+CHR(13)+'?????????: '+ALLTRIM(newOsnov)
   CASE logtxt=3
        REPLACE txtprn WITH ALLTRIM(str_ini)+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? '+strDateBeg+;
                ' ?????????'+IIF(!EMPTY(newNKurs),', '+ALLTRIM(newNKurs),'')+' ?? ?????? '+LOWER(cDolZam)+' '+LOWER(cPodrZam)+;
                ', ? ??????? ???????? ???????? ??????????, '+ALLTRIM(newStrPrich)+' '+str_zam+CHR(13)+'?????????: '+ALLTRIM(newOsnov)                          
   CASE logtxt=4
        REPLACE txtprn WITH ALLTRIM(str_ini)+' '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+' ? '+strDateBeg+;
                ' ??????? ?? ???????????????? ?? '+STR(newKse,4,2)+' ??????? ??????? '+LOWER(cDolZam)+' '+LOWER(cPodrZam)+;
                ', ? ??????? ???????? ???????? ??????????, '+ALLTRIM(newStrPrich)+' '+str_zam+CHR(13)+'?????????: '+ALLTRIM(newOsnov)               
ENDCASE        
*************************************************************************************************************************
PROCEDURE procOrdKurs
PARAMETERS par1	
IF !USED('datkurs')
   USE datkurs ORDER 2 IN 0
ENDIF
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
SELECT kod,name FROM sprtot WHERE kspr=22 INTO CURSOR curTypeKurs ORDER BY kod
SELECT curTypeKurs
GO TOP

parPadejDol=1
parPadejFio=1
=AFIELDS(arSchool,'datKurs')
CREATE CURSOR curSchool FROM ARRAY arSchool
SELECT nameschool FROM datKurs INTO CURSOR curSchoolSupl DISTINCT ORDER BY nameschool

SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot
newNSchool=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl,1,100)))
newNKurs=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl,101,150)))

PUBLIC newKp2,newNpodr2
newKp2=IIF(logNewRec,1,curpeoporder.kp2)
newNPodr2=IIF(logNewRec,curTypeKurs.name,curpeoporder.npodr2)

newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxBeg'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
        
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.
     
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
          
     DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?????????',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newNSchool',.F.,IIF(par1,.T.,.F.),0    
     .tBoxFioNew.procforChange='DO changeSchool'   
     DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectSchool',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlntNew.Enabled=IIF(par1,.T.,.F.)
         
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordFioNew.Top+dHeight-1,.ordPrik.Width,dHeight,'???????????? ??????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKurs',.F.,IIF(par1,.T.,.F.),0        
     
     DO adTboxAsCont WITH 'frmOrd','ordTr',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'??? ??????',1,1          
     DO addComboMy WITH 'frmOrd',22,.comboBox1.Left,.ordTr.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNpodr2','curTypeKurs.name',6,.F.,'newKp2=curTypeKurs.kod',.F.,.T. 
     .comboBox22.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordTr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr2',.F.,.F.,0  
     .tBoxPodr2.Visible=IIF(par1,.F.,.T.) 
              
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
     DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
     WITH .listBox2
          .RowSource='curSchool.namekurs'               
          .RowSourceType=2
          .ColumnCount=1 
          .Height=.Parent.Height-.Top
          .Visible=.F.        
          .procForDblClick='DO validListSchool'
          .procForLostFocus='DO lostFocusPeop2'
     ENDWITH               
     IF par1 
        ON ERROR DO erSup        
        .comboBox22.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH	
*************************************************************************************************************************
PROCEDURE saveOKurs
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textKurs
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newdateEnd;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp
        REPLACE varSupl WITH PADR(ALLTRIM(newNSchool),100,' ')+PADR(ALLTRIM(newNKurs),150,' '),nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex,;
        curpeoporder.kp2 WITH newKp2,npodr2 WITH newNPodr2
DO saveDimOrd  
DO saveAdmOrder  
*************************************************************************************************************************
PROCEDURE textKurs
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
str_ini=''
DO procfioini WITH 'cOrdFio' 
SELECT txtOrder
REPLACE txtprn WITH UPPER(LEFT(cOrdDol,1))+LOWER(SUBSTR(cOrdDol,2))+' '+LOWER(cOrdPodr)+' '+ALLTRIM(str_ini)+' ? '+strDateBeg+' ?? '+strDateEnd+;
                ' ????????? ? '+ALLTRIM(newNSchool)+' ?? ????? ????????? ???????????? ?? ????????? '+ALLTRIM(newNKurs)+' ? ??????????? ??????? ?????????? ????? ?? ????? ??????.'+CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
******************************************************************************************************************************************************        
PROCEDURE saveAdmOrder
SELECT datJob
oldOrdJob=SYS(21)
SELECT curPeopOrder
IF logApp           
   SELECT datJob
   DO CASE
      CASE INLIST(curPeopOrder.supOrd,101,102,104,105,113)    && ????????? ????????????????+??????????? ???? ????????????  ? ?.?. ?????? ??????????? (???????????? ? datjob)                                       
           IF curPeopOrder.nidJob=0
              SET DELETED OFF 
              SET ORDER TO 7
              GO BOTTOM
              newNidCx=nid+1
              SET DELETED ON 
              APPEND BLANK
              REPLACE kodpeop WITH curpeoporder.kodpeop,nid WITH newNidCx,nidPeop WITH curPeopOrder.nidPeop                         
              REPLACE curPeopOrder.nidjob WITH newNidCx                           
              SELECT peoporder
              SEEK curpeoporder.nid
             * SEEK curpeoporder.kodpeop+curpeoporder.kord
              REPLACE peoporder.nidjob WITH newNidCx
              SELECT datjob 
           ENDIF
           IF SEEK(curPeopOrder.nidJob,'datJob',7)
              REPLACE kp WITH VAL(SUBSTR(curPeopOrder.varSupl,1,3)),kd WITH VAL(SUBSTR(curPeopOrder.varSupl,104,3)),kse WITH curPeopOrder.kse,tr WITH curPeopOrder.tr,;
                      dordin WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,dateBeg WITH curPeopOrder.dateBeg,;
                      kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0),lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)
              DO repnadjob       
           ENDIF
           IF INLIST(curPeopOrder.supOrd,102,105)              
              REPLACE kdek WITH VAL(SUBSTR(curpeoporder.varSupl2,1,5)),fiodek WITH SUBSTR(curPeoporder.varSupl2,6,60)
           ENDIF                  
           SELECT datJob                                                                    
      CASE INLIST(curPeopOrder.supOrd,106,107)    && ?????? ????????????????
           SELECT datjob
           SET ORDER TO 7
           SEEK curPeopOrder.oldNid
           REPLACE dateOut WITH curPeopOrder.dateBeg,dordOut WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder
           kprep=kp
           kdrep=kd
           kserep=datjob.kse-curpeoporder.kse
           trrep=IIF(curPeopOrder.supOrd=106,2,4)
           *-----???????? ?????? ????????????????
           IF kserep#0          
              IF curPeopOrder.nidjob=0                 
                 SET DELETED OFF 
                 SET ORDER TO 7
                 GO BOTTOM
                 newNidCx=nid+1
                 SET DELETED ON 
                 APPEND BLANK
                 REPLACE kodpeop WITH curpeoporder.kodpeop,nid WITH newNidCx,nidPeop WITH curPeopOrder.nidPeop                
              ELSE 
                 SEEK curpeoporder.nidjob    
                 newNidcx=curpeoporder.nidjob                    
              ENDIF 
              REPLACE kp WITH kprep,kd WITH kdrep,kse WITH kserep,tr WITH trrep,dateBeg WITH curPeopOrder.dateBeg,dordIn WITH dateOrdNew,nordin WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordIn WITH newKodOrder,;
              lkv WITH IIF(SEEK(curpeoporder.kodpeop,'people',1).AND.people.kval>0,.T.,.F.)              
              REPLACE curPeopOrder.nidjob WITH newNidCx                           
              SELECT peoporder            
              SEEK curpeoporder.nid
              REPLACE peoporder.nidjob WITH newNidCx
              SELECT datjob 
           ENDIF
      CASE curPeopOrder.supOrd=115  &&?????????? ?? ????????????????
           SELECT datjob
           SET ORDER TO 7
           SEEK curPeopOrder.oldnid         
           REPLACE dateOut WITH curPeopOrder.dateBeg,dordOut WITH dateOrdNew,nordout WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,kordOut WITH newKodOrder
      CASE curPeopOrder.supOrd=141    && ?????
           SELECT datKurs
           SET ORDER TO 3
           SEEK curPeopOrder.nid
           IF !FOUND()
              APPEND BLANK
              REPLACE kodpeop WITH curPeopOrder.kodPeop,nidPeop WITH curPeopOrder.nidPeop,kord WITH curpeoporder.kord,supord WITH curpeoporder.supord,nidOrd WITH curPeopOrder.nid
           ENDIF     
           REPLACE perBeg WITH curPeopOrder.dateBeg,perEnd WITH curPeopOrder.dateEnd,nameschool WITH SUBSTR(curPeopOrder.varsupl,1,100),;
           namekurs WITH SUBSTR(curPeopOrder.varsupl,101),kord WITH curpeoporder.kord,supord WITH curpeoporder.supord,dord WITH dateOrdNew,nord WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,;
           ntype WITH IIF(INLIST(curpeoporder.kp2,1,2,3),curpeoporder.kp2,0)
      CASE curpeopOrder.supOrd=170
           SELECT people          
           IF SEEK(curPeopOrder.kodpeop,'people',1)
              REPLACE dekotp WITH .F.,ddekotp WITH CTOD('  .  .    ')
           ENDIF  
      CASE curpeoporder.supOrd=180 && ????????? ???????
           IF SEEK(curpeoporder.kodpeop,'people',1) 
              SELECT people
              REPLACE oldfio WITH SUBSTR(curpeoporder.varsupl,51,50),fio WITH ALLTRIM(SUBSTR(curpeoporder.varsupl,1,50))+' '+ALLTRIM(SUBSTR(ALLTRIM(people.fio),AT(' ',ALLTRIM(people.fio))))
              new_pfio=ALLTRIM(LEFT(ALLTRIM(fio),AT(' ',ALLTRIM(fio))))
              new_pname=SUBSTR(fio,AT(' ',ALLTRIM(fio))+1)
              new_pname=ALLTRIM(LEFT(ALLTRIM(new_pname),AT(' ',ALLTRIM(new_pname))))
              new_potch=ALLTRIM(SUBSTR(fio,RAT(' ',ALLTRIM(fio))))
              new_who=''
              new_whom=''
              new_whomv=''
              new_whomp=''
              DO procpadej WITH 'new_pfio','new_pname','new_potch','new_who','new_whom','new_whomv','new_whomp'
              REPLACE fior WITH new_who,fiod WITH new_whom,fiov WITH new_whomv,fiot WITH new_whomp
           ENDIF 
                  
   ENDCASE
ENDIF
SELECT curPeopOrder
SELECT datJob
SET ORDER TO &oldOrdJob
DO appendFromPeopOrder
LOCATE FOR nid=newNid
*************************************************************************************************************************
PROCEDURE procOrdKm
PARAMETERS par1	
IF !USED('datkurs')
   USE datkurs ORDER 2 IN 0
ENDIF
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
SELECT kod,name FROM sprtot WHERE kspr=22 INTO CURSOR curTypeKurs ORDER BY kod
parPadejDol=1
parPadejFio=1
=AFIELDS(arSchool,'datKurs')
CREATE CURSOR curSchool FROM ARRAY arSchool
SELECT nameschool FROM datKurs INTO CURSOR curSchoolSupl DISTINCT ORDER BY nameschool

SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr 
DO resetVarTot

newdayOtp=IIF(logNewRec,0,curpeoporder.dayotp)
newNSchool=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl,1,100)))
newNKurs=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl,101,150)))
newCity=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl2,1,100)))

PUBLIC newKp2,newNpodr2
newKp2=IIF(logNewRec,0,curpeoporder.kp2)
newNPodr2=IIF(logNewRec,'',curpeoporder.npodr2)

newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxKont'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ???? (?????)',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newCity',.F.,IIF(par1,.T.,.F.),0,
     DO adTboxAsCont WITH 'frmOrd','ordDayOtp',.ordPrik.Left,.ordPKont.Top+dHeight-1,.ordPrik.Width,dHeight,'???-?? ????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDayotp',.ordDayOtp.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDayOtp',.F.,IIF(par1,.T.,.F.),0,'99' 
     .tBoxDayOtp.Alignment=0
        
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDayOtp.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.,'DO validDayKm'
     
     DO adTboxAsCont WITH 'frmOrd','ordEnd',.ordPrik.Left,.ordBeg.Top,RetTxtWidth('??w'),dHeight,'??',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxEnd',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateEnd',.F.,IIF(par1,.T.,.F.),0  
     .tBoxBeg.Width=(.comboBox1.Width-.ordEnd.Width)/2
     .ordEnd.Left=.tBoxBeg.left+.tBoxBeg.Width-1
     .tBoxEnd.Left=.ordEnd.Left+.ordEnd.Width-1
     .tBoxEnd.Width=.comboBox1.Width-.tBoxBeg.Width-.ordEnd.Width+2
              
     DO adTboxAsCont WITH 'frmOrd','ordFioNew',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'??????? ?????????',1,1      
     DO adTboxNew WITH 'frmOrd','tBoxFioNew',.ordFioNew.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newNSchool',.F.,IIF(par1,.T.,.F.),0    
     .tBoxFioNew.procforChange='DO changeSchool'   
     DO adtboxnew WITH 'frmOrd','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFio.Width+1,dheight,'',.F.,IIF(par1,.T.,.F.)   
     DO addButtonOne WITH 'frmOrd','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectSchool',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlntNew.Enabled=IIF(par1,.T.,.F.)
     
     DO adTboxAsCont WITH 'frmOrd','ordDolNew',.ordPrik.Left,.ordFioNew.Top+dHeight-1,.ordPrik.Width,dHeight,'???????????? ??????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDoljNew',.ordDolNew.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKurs',.F.,IIF(par1,.T.,.F.),0        
     
     DO adTboxAsCont WITH 'frmOrd','ordTr',.ordPrik.Left,.ordDolNew.Top+dHeight-1,.ordPrik.Width,dHeight,'??? ????????????',1,1          
     DO addComboMy WITH 'frmOrd',22,.comboBox1.Left,.ordTr.Top,dheight,.comboBox1.Width,IIF(par1,.T.,.F.),'newNpodr2','curTypeKurs.name',6,.F.,'newKp2=curTypeKurs.kod',.F.,.T. 
     .comboBox22.Visible=IIF(par1,.T.,.F.)
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordTr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr2',.F.,.F.,0  
     .tBoxPodr2.Visible=IIF(par1,.F.,.T.) 
              
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordTr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
     DO addListBoxMy WITH 'frmOrd',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
     WITH .listBox2
          .RowSource='curSchool.namekurs'               
          .RowSourceType=2
          .ColumnCount=1 
          .Height=.Parent.Height-.Top
          .Visible=.F.        
          .procForDblClick='DO validListSchool'
          .procForLostFocus='DO lostFocusPeop2'
     ENDWITH 
                   
     IF par1   
        ON ERROR DO erSup      
        .comboBox22.SetFocus
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH	

***********************************************************************************************************************
PROCEDURE validDayKm
newDateEnd=newDateBeg+newDayOtp-1
frmOrd.tBoxEnd.Refresh
KEYBOARD '{TAB}'
***********************************************************************************************************************
PROCEDURE selectSchool
SELECT curSchool
ZAP
APPEND FROM DBF ('curSchoolSupl')
WITH frmOrd
     .listBox2.RowSource='curSchool.nameSchool'                     
      IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.  
        .listBox2.SetFocus            
     ENDIF 
ENDWITH 
*****************************************************************************************************************
PROCEDURE changeSchool
WITH frmOrd 
     IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=frmOrd.tBoxFioNew.Text 
SELECT curSchool
ZAP
APPEND FROM DBF ('curSchoolSupl') FOR LOWER(ALLTRIM(lcValue))$LOWER(nameschool)
WITH frmOrd.listBox2
     .RowSource='curSchool.nameSchool'                    
     .Visible=IIF(RECCOUNT('curSchool')=0,.F.,.T.)      
ENDWITH 
************************************************************************************************************************
PROCEDURE validListSchool
newNSchool=curSchool.nameschool
WITH frmOrd
     .tBoxFioNew.ControlSource='newNSchool'
     .listBox2.Visible=.F.
     .tBoxFioNew.Refresh  
     .tBoxDoljNew.SetFocus
ENDWITH 
*************************************************************************************************************************
PROCEDURE saveKm
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textKm	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,dateEnd WITH newdateEnd;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp
        REPLACE varSupl WITH PADR(ALLTRIM(newNSchool),100,' ')+PADR(ALLTRIM(newNKurs),150,' '),varSupl2 WITH newCity,dayOtp WITH newdayOtp, nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex,;
        curpeoporder.kp2 WITH newKp2,npodr2 WITH newNPodr2       
DO saveDimOrd  
DO saveKmOrder 
*************************************************************************************************************************
PROCEDURE textKm
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
str_ini=''
DO procfioini WITH 'cOrdFio' 
SELECT txtOrder
DO CASE
   CASE newKp2=1 &&????? ???????? ????????????
        REPLACE txtprn WITH '1. ????????? '+ALLTRIM(cOrdFio)+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+ ' ? '+ALLTRIM(newCity)+;
                ' '+ALLTRIM(newNSchool)+' ?? '+LTRIM(STR(newDayOtp))+IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ????',' ??????????? ???'))+' ? '+strDateBeg+' ?? '+strDateEnd+;
                ' ?? ????????? ???????????? ?? ????????? '+ALLTRIM(newNKurs)+' ? ??????????? ??????? ?????????? ????? ?? ????? ??????.'+CHR(13)+;
                '2. ?????? ??????????????? ???????? ?????????? ? ???????????? ? ??????????? ?????????????????.'+CHR(13)+ CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
   CASE newKp2=2 &&??????????????
        REPLACE txtprn WITH '1. ????????? '+ALLTRIM(cOrdFio)+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+ ' ? '+ALLTRIM(newCity)+;
                ' '+ALLTRIM(newNSchool)+' ?? '+LTRIM(STR(newDayOtp))+IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ????',' ??????????? ???'))+' ? '+strDateBeg+' ?? '+strDateEnd+;
                ' ?? ?????????????? ?? ????????? '+ALLTRIM(newNKurs)+' ? ??????????? ??????? ?????????? ????? ?? ????? ??????.'+CHR(13)+;
                '2. ?????? ??????????????? ???????? ?????????? ? ???????????? ? ??????????? ?????????????????.'+CHR(13)+ CHR(13)+'?????????: '+ALLTRIM(newOsnov)              
   CASE newKp2=5 &&??????????? ??????????
        REPLACE txtprn WITH '1. ????????? '+ALLTRIM(cOrdFio)+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+ ' ? '+ALLTRIM(newCity)+;
                ' '+ALLTRIM(newNSchool)+' ?? '+LTRIM(STR(newDayOtp))+IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ????',' ??????????? ???'))+' ? '+strDateBeg+' ?? '+strDateEnd+;
                ' ?? ????????? ???????????? ?? ????????? '+ALLTRIM(newNKurs)+' ? ??????????? ??????? ?????????? ????? ?? ????? ??????.'+CHR(13)+;
                '2. ?????? ??????????????? ???????? ?????????? ? ???????????? ? ??????????? ?????????????????.'+CHR(13)+ CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
   CASE newKp2=6   &&????? ????????????????? ????????   
        REPLACE txtprn WITH '1. ????????? '+ALLTRIM(cOrdFio)+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+ ' ? '+ALLTRIM(newCity)+;
                ' '+ALLTRIM(newNSchool)+' ??? ????? ????????????????? ???????? '+;
                ' ?? '+LTRIM(STR(newDayOtp))+IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ???? ',' ??????????? ??? '))+IIF(newDayOtp=1,strDateBeg,'? '+strDateBeg+' ?? '+strDateEnd)+;
                ' ? ??????????? ??????? ?????????? ????? ?? ????? ??????.'+CHR(13)+;
                '2. ?????? ??????????????? ???????? ?????????? ? ???????????? ? ??????????? ?????????????????.'+CHR(13)+ CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
   OTHERWISE 
        REPLACE txtprn WITH '1. ????????? '+ALLTRIM(cOrdFio)+', '+LOWER(cOrdDol)+' '+LOWER(cOrdPodr)+ ' ? '+ALLTRIM(newCity)+;
                ' '+ALLTRIM(newNSchool)+' ?? '+LTRIM(STR(newDayOtp))+IIF(newDayOtp>4,' ??????????? ????',IIF(newDayOtp=1,' ??????????? ????',' ??????????? ???'))+' ? '+strDateBeg+' ?? '+strDateEnd+;
                ' ?? ????????? ???????????? ?? ????????? '+ALLTRIM(newNKurs)+' ? ??????????? ??????? ?????????? ????? ?? ????? ??????.'+CHR(13)+;
                '2. ?????? ??????????????? ???????? ?????????? ? ???????????? ? ??????????? ?????????????????.'+CHR(13)+ CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
      
ENDCASE 
*************************************************************************************************************************        
PROCEDURE saveKmOrder 
SELECT curPeopOrder
DO CASE
   CASE curPeopOrder.supOrd=201
        IF INLIST(curpeoporder.kp2,1,2,3)    && ???????????? ?? ????? -  ????????????          
           SELECT datKurs
           SET ORDER TO 3
           SEEK curPeopOrder.nid
           IF !FOUND()
              APPEND BLANK
              REPLACE kodpeop WITH curPeopOrder.kodPeop,nidpeop WITH curPeopOrder.nidpeop kord WITH curpeoporder.kord,supord WITH curpeoporder.supord,nidOrd WITH curPeopOrder.nid
           ENDIF     
           REPLACE perBeg WITH curPeopOrder.dateBeg,perEnd WITH curPeopOrder.dateEnd,nameschool WITH SUBSTR(curPeopOrder.varsupl,1,100),;
                   namekurs WITH SUBSTR(curPeopOrder.varsupl,101,50),kord WITH curpeoporder.kord,supord WITH curpeoporder.supord,dord WITH dateOrdNew,nord WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,;
                   ntype WITH curpeoporder.kp2
        ELSE              
           IF SEEK(curpeoporder.nid,'datkurs',3)
              SELECT datkurs
              DELETE
           ENDIF
           SELECT curPeopOrder
        ENDIF  
   *CASE curPeopOrder.supOrd=201    && ???????????? ?? ????? - ????????????????
   *     SELECT datKurs
   *     SET ORDER TO 3
   *     SEEK curPeopOrder.nid
   *     IF !FOUND()
   *        APPEND BLANK
   *        REPLACE kodpeop WITH curPeopOrder.kodPeop,nidpeop WITH curPeopOrder.nidpeop kord WITH curpeoporder.kord,supord WITH curpeoporder.supord,nidOrd WITH curPeopOrder.nid
   *     ENDIF     
   *     REPLACE perBeg WITH curPeopOrder.dateBeg,perEnd WITH curPeopOrder.dateEnd,nameschool WITH SUBSTR(curPeopOrder.varsupl,1,100),;
   *             namekurs WITH SUBSTR(curPeopOrder.varsupl,101,50),kord WITH curpeoporder.kord,supord WITH curpeoporder.supord,dord WITH dateOrdNew,nord WITH LTRIM(STR(numOrdNew))+'-'+strOrdNew,;
   *             ntype WITH 1
   *        
ENDCASE
DO appendFromPeopOrder
LOCATE FOR nid=newNid
*************************************************************************************************************************
PROCEDURE procOrdFam
PARAMETERS par1	
parPadejDol=2
parPadejFio=2

SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr
DO resetVarTot 
newFam=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl,1,50)))
oldFam=IIF(logNewRec,SPACE(100),ALLTRIM(SUBSTR(curpeoporder.varsupl,51,50)))
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxKont'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     
     DO adTboxAsCont WITH 'frmOrd','ordPkont',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????? ???????',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxKont',.ordPkont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newFam',.F.,IIF(par1,.T.,.F.),0,
        
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordPkont.Top+dHeight-1,.ordPrik.Width,dHeight,'???????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.
     
              
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
                      
     IF par1   
        ON ERROR DO erSup        
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH	
*************************************************************************************************************************        
PROCEDURE saveRecFam
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textFam	
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
oldFio_cx=IIF(SEEK(newKodpeop,'people',1),ALLTRIM(LEFT(people.fio,AT(' ',people.fio))),'')
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
        varSupl WITH PADR(ALLTRIM(newFam),50,' ')+PADR(ALLTRIM(oldFio_cx),50,' '),nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd  
DO saveAdmOrder 
*************************************************************************************************************************
PROCEDURE textFam
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
oldFio_cx=IIF(SEEK(newKodpeop,'people',1),ALLTRIM(LEFT(people.fio,AT(' ',people.fio))),'')
str_ini=''
DO procfioini WITH 'cOrdFio' 
SELECT txtOrder
REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+ ' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
        '???????? ? ??????? ? ???? ??????????? ?????????? ??????? '+ALLTRIM(oldFio_cx)+' ?? '+ALLTRIM(newFam)+' ? ????? ? ???????????????' +CHR(13)+'?????????: '+ALLTRIM(newOsnov) 
*************************************************************************************************************************
PROCEDURE procOrdGraph
PARAMETERS par1	
parPadejDol=2
parPadejFio=2

SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr
DO resetVarTot 
newNKont=IIF(logNewRec,SPACE(20),ALLTRIM(SUBSTR(curpeoporder.varsupl,1,20)))     && ??????
newNschool=IIF(logNewRec,SPACE(20),ALLTRIM(SUBSTR(curpeoporder.varsupl,21,20)))  && ????
newNkurs=IIF(logNewRec,SPACE(20),ALLTRIM(SUBSTR(curpeoporder.varsupl,41,20)))    && ?????????
newStrPrich=IIF(logNewRec,'?? ???????? ???????????????',ALLTRIM(SUBSTR(curpeoporder.varsupl,61,60)))    && ???????
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxBeg'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
             
     DO adTboxAsCont WITH 'frmOrd','ordBeg',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordDol.Width,dHeight,'???????? ?',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxBeg',.ordBeg.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDateBeg',.F.,IIF(par1,.T.,.F.),0,.F.
     
     DO adTboxAsCont WITH 'frmOrd','ordDay',.ordPrik.Left,.ordBeg.Top+dHeight-1,.ordPrik.Width,dHeight,'???????',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxday',.ordDay.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newStrPrich',.F.,IIF(par1,.T.,.F.),0,
     
     DO adTboxAsCont WITH 'frmOrd','ordPodr2',.ordPrik.Left,.ordDay.Top+dHeight-1,.ordPrik.Width,dHeight,'?????? ???????? ???????',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxPodr2',.ordPodr2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKont',.F.,IIF(par1,.T.,.F.),0,
     
     DO adTboxAsCont WITH 'frmOrd','ordPodr3',.ordPrik.Left,.ordPodr2.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ???????',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxPodr3',.ordPodr3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNSchool',.F.,IIF(par1,.T.,.F.),0,
     
     DO adTboxAsCont WITH 'frmOrd','ordPodr4',.ordPrik.Left,.ordPodr3.Top+dHeight-1,.ordPrik.Width,dHeight,'????????? ???????? ???????',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxPodr4',.ordPodr4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNKurs',.F.,IIF(par1,.T.,.F.),0,
              
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordPodr4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10' 
                      
     IF par1  
        ON ERROR DO erSup         
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH	
*************************************************************************************************************************        
PROCEDURE saveRecGraph
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textGraph
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
oldFio_cx=IIF(SEEK(newKodpeop,'people',1),ALLTRIM(LEFT(people.fio,AT(' ',people.fio))),'')
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,dateBeg WITH newDateBeg,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
        varSupl WITH PADR(ALLTRIM(newNKont),20,' ')+PADR(ALLTRIM(newNSchool),20,' ')+PADR(ALLTRIM(newNKurs),20,' ')+PADR(ALLTRIM(newStrPrich),20,' '),nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd  
DO saveAdmOrder 
*************************************************************************************************************************
PROCEDURE textGraph
strDateBeg=IIF(!EMPTY(newDateBeg),dateToString('newDateBeg',.T.),'')
strDateEnd=IIF(!EMPTY(newDateEnd),dateToString('newDateEnd',.T.),'')
cOrdFio=''
cOrdPodr=''
cOrdDol=''
DO procOrdNameDol
str_ini=''
DO procfioini WITH 'cOrdFio' 
SELECT txtOrder
REPLACE txtprn WITH cOrdDol+' '+LOWER(cOrdPodr)+ ' '+ALLTRIM(str_ini)+' ? '+strDateBeg+;
        '???????? ?????? ?????? '+ALLTRIM(newStrPrich)+':'+CHR(13)+;
        '?????? ???????? ??????? - '+ALLTRIM(newNkont)+CHR(13)+;
        '????????? ??????? - '+ALLTRIM(newnSchool)+CHR(13)+;
        '????????? ???????? ??????? - '+ALLTRIM(newNKurs)+CHR(13)+'?????????: '+ALLTRIM(newOsnov)        
*************************************************************************************************************************
PROCEDURE procCustom
PARAMETERS par1	
parPadejDol=2
parPadejFio=2

SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr
DO resetVarTot 
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.editOrder'
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder
              
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     .ordOsnov.Visible=.F.
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     .tBoxOsnov.Visible=.F.
     DO addObjectOrder2 WITH 'frmOrd.ordFio.Top+dHeight+10' 
                      
     IF par1      
        ON ERROR DO erSup     
        .tBoxFio.SetFocus        
        ON ERROR
     ENDIF
ENDWITH	
*************************************************************************************************************************
PROCEDURE saveRecCustom
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,supord WITH newKodprik,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
        nid WITH newNid,dord WITH repdorder,nord WITH repnorder,sex WITH kodsex
DO saveDimOrd 
DO saveAdmOrder 
*************************************************************************************************************************
PROCEDURE textCustom
*************************************************************************************************************************        
PROCEDURE newNidOrder
IF logNewRec
   SELECT peoporder
   =FLOCK()
   SET DELETED OFF 
   GO BOTTOM 
   newNid=nid+1  
   APPEND BLANK
   REPLACE nid WITH newNid
   FLUSH 
   SET DELETED ON
   UNLOCK  
   SELECT curPeopOrder
ELSE
   newNid=curpeopOrder.nid   
ENDIF
*************************************************************************************************************************        
PROCEDURE repNppOrder
nppNew=1
SET ORDER TO 
SCAN ALL
     REPLACE npp WITH nppNew
     IF SEEK(curpeoporder.nid,'peoporder',1)
        REPLACE peoporder.npp WITH curpeoporder.npp
     ENDIF
     SELECT curpeoporder
     nppNew=nppNew+1
ENDSCAN
SET ORDER TO 1
GO TOP
*************************************************************************************************************************
PROCEDURE saveDimOrd
readNidPeop=curpeoporder.nidpeop
SCATTER TO dimSave        
SELECT peopOrder
SEEK newNid
GATHER FROM dimSave
REPLACE txtOrd WITH curPeopOrder.txtOrd           
SELECT peopOrder
**************************************************************************************************************************
PROCEDURE resetvartot
newKodPrik=IIF(!par1,curpeoporder.supOrd,newKodprik)
strPrik=IIF(SEEK(newKodPrik,'curSprOrder',1),curSprOrder.nameOrd,'')
newKodPeop=IIF(logNewRec,0,curpeoporder.kodpeop)
newFioPeop=IIF(logNewRec,'',curpeoporder.fiopeop)
newNidPeop=IIF(logNewRec,0,curpeoporder.nidpeop)
newDateBeg=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.dateBeg)
newDateEnd=IIF(logNewRec,CTOD('  .  .    '),curpeoporder.dateEnd)
newKp=IIF(logNewRec,0,curpeoporder.kp)
newNPodr=IIF(logNewRec,'',curpeoporder.npodr)
newKd=IIF(logNewRec,0,curpeoporder.kd)
newNDolj=IIF(logNewRec,'',curpeoporder.ndolj)
newKse=IIF(logNewRec,1.00,curpeoporder.kse)
newTr=IIF(logNewRec,1,curpeoporder.tr)
strType=IIF(SEEK(newTr,'sprtype',1),sprtype.name,'')
newOsnov=IIF(logNewRec,SPACE(100),curpeoporder.osnov)
kodsex=IIF(logNewRec,0,curpeoporder.sex)
********************************************************************************************
PROCEDURE validCheckclose
REPLACE datorder.lotm WITH newLOtm
********************************************************************************************
PROCEDURE cancelOrder
PARAMETERS par1,par2	
parPadejDol=2
parPadejFio=2

SELECT curPeopOrder
DO newNidOrder
DO removeObjectKadr
DO resetVarTot 
newPKont=IIF(logNewRec,0,VAL(SUBSTR(curpeoporder.varsupl,1,3)))
newDDop=IIF(logNewRec,0,VAL(SUBSTR(curpeoporder.varsupl,4,5)))
newPerBeg=IIF(logNewRec,CTOD('  .  .    '),CTOD(SUBSTR(curpeoporder.varsupl,9,10)))
newLogApp=.T.
newNpp=0
str_ini=''
logDatJob=.F.
objFocus='frmOrd.tBoxDolj2'
nidJobNew=IIF(logNewRec,0,curpeoporder.nidjob)
SELECT txtorder
REPLACE txtprn WITH IIF(logNewRec,'',curpeoporder.txtord)  
SELECT curpeoporder
WITH frmOrd 
     DO addObjectOrder 
     DO adTboxAsCont WITH 'frmOrd','ordPodr',.ordPrik.Left,.ordFio.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxPodr',.ordPodr.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNpodr',.F.,.F.,0           
     DO adTboxAsCont WITH 'frmOrd','ordDol',.ordPrik.Left,.ordPodr.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1 
     DO adTboxNew WITH 'frmOrd','tBoxDolj',.ordDol.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNDolj',.F.,.F.,0   
     
     DO adTboxAsCont WITH 'frmOrd','ordDol2',.ordPrik.Left,.ordDol.Top+dHeight-1,.ordPrik.Width,dHeight,'????????',1,1         
     DO adTboxNew WITH 'frmOrd','tBoxDolj2',.ordDol2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'NewPkont',.F.,IIF(par1,.T.,.F.),0,.F.
     .tBoxDolj2.InputMask='999' 
     .tBoxDolj2.Alignment=0
        
     DO adTboxAsCont WITH 'frmOrd','ordDol3',.ordPrik.Left,.ordDol2.Top+dHeight-1,.ordPrik.Width,dHeight,'?????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDolj3',.ordDol3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDDop',.F.,IIF(par1,.T.,.F.),0,.F.
     .tBoxDolj3.InputMask='999' 
     .tBoxDolj3.Alignment=0
     
     DO adTboxAsCont WITH 'frmOrd','ordDol4',.ordPrik.Left,.ordDol3.Top+dHeight-1,.ordPrik.Width,dHeight,'????',1,1   
     DO adTboxNew WITH 'frmOrd','tBoxDolj4',.ordDol4.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newPerBeg',.F.,IIF(par1,.T.,.F.),0,.F.
                   
     DO adTboxAsCont WITH 'frmOrd','ordOsnov',.ordPrik.Left,.ordDol4.Top+dHeight-1,.ordPrik.Width,dHeight,'?????????',1,1  
     DO adTboxNew WITH 'frmOrd','tBoxOsnov',.ordOsnov.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newOsnov',.F.,IIF(par1,.T.,.F.),0  
     
     DO addObjectOrder2 WITH 'frmOrd.ordOsnov.Top+frmOrd.ordOsnov .Height+10'                       
  *   .tBoxDolj2.SetFocus        
ENDWITH	
********************************************************************************************
PROCEDURE textCancel
fNewOrder=LTRIM(STR(newDDop))+'-?'
IF SEEK(STR(newnidPeop,5)+DTOS(newPerBeg)+fNewOrder+STR(newPkont,3),'peoporder',3)
   nidJobNew=peoporder.nid
   SELECT txtOrder
   REPLACE txtprn WITH '???????? '+LTRIM(STR(newPkont))+' ??????? '+fNewOrder+' ?? '+DTOC(newPerBeg)+' ?.: "'+ALLTRIM(peoporder.txtord)+'" ????????.'+CHR(13)+'?????????: '+ALLTRIM(newOsnov)        
ENDIF
********************************************************************************************
PROCEDURE saveRecCancel
IF LEN(ALLTRIM(txtOrder.txtprn))=0
   DO textCancel
ENDIF
SELECT curPeopOrder
IF logNewRec
   APPEND BLANK
   REPLACE npp WITH newNpp
ENDIF
REPLACE kodPeop WITH newKodPeop,nidPeop WITH newNidPeop,fiopeop WITH newFioPeop,typeOrd WITH typeOrdNew,kord WITH newKodOrder,supord WITH newKodprik,;
        kp WITH newKp,npodr WITH newNPodr,kd WITH newKd,ndolj WITH newNDolj,supord WITH newKodprik,osnov WITH newOsnov,txtOrd WITH txtOrder.txtprn,logapp WITH newLogApp,;
        varSupl WITH PADR(ALLTRIM(STR(newPkont)),3,' ')+PADR(ALLTRIM(STR(newDDop)),5,' ')+PADR(DTOC(newPerBeg),10,' '),;
        nid WITH newNid,dord WITH repdorder,nord WITH repnorder,nidjob WITH nidjobnew,sex WITH kodsex       
DO saveDimOrd   
DO saveotpOrder 
**********************************************************
PROCEDURE appendFromPeopOrder
SELECT curPeopOrder
DELETE ALL
SELECT peoporder
SET ORDER TO 4
SEEK newKodOrder
SCAN WHILE kord=newKodOrder
     SCATTER TO dimApOrd
     SELECT curpeoporder
     APPEND BLANK
     GATHER FROM dimApOrd
     REPLACE txtord WITH  peoporder.txtord 
     SELECT peoporder
ENDSCAN
SET ORDER TO 1
SELECT curpeoporder
DO repNppOrder
****************************************************************************************************************************************************
PROCEDURE dateToString
PARAMETERS parDate,parStrYear
repVar=LTRIM(STR(DAY(&parDate)))+' '+month_prn(MONTH(&parDate))+' '+STR(YEAR(&parDate),4)+IIF(parStrYear,' ?.','')
RETURN repVar
***********************************************************
PROCEDURE procSearchOrder
search_cx=SPACE(50)
WITH frmOrd
     logNewRec=.F.
     .SetAll('Visible',.F.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')    
     .butSearchRec.Visible=.T.
     .butSearchRet.Visible=.T.
     .tBoxSearch.Visible=.T.
     .tBoxSearch.Enabled=.T.
     *.grdPers.Enabled=.F.
ENDWITH
**********************************************************
PROCEDURE searchRecOrder
IF EMPTY(search_cx)
   RETURN
ENDIF
DO unosimbol WITH 'search_cx',.F.,.F.           
   
SELECT curPeopOrder
LOCATE FOR LOWER(ALLTRIM(search_cx))$LOWER(fiopeop)
IF FOUND()
   DO changeRowPeopOrder
   frmOrd.butSearchRec.Visible=.F.
   frmOrd.butSearchNext.Visible=.T.
ELSE 
   frmOrd.tBoxSearch.SetFocus   
ENDIF 
**********************************************************
PROCEDURE searchRetOrder
WITH frmOrd
     logNewRec=.F.
     .butSearchRec.Visible=.F.
     .butSearchNext.Visible=.F.
     .butSearchRet.Visible=.F.
     .tBoxSearch.Visible=.F.
     .tBoxSearch.Enabled=.F.
     .butNew.Visible=.T.
     .butRead.Visible=.T.
     .butDel.Visible=.T.
     .butPrn.Visible=.T.  
     .butSearch.Visible=.T.    
     .butRet.Visible=.T.
     .grdPers.Enabled=.T.
ENDWITH
************************************************************************************************************************
PROCEDURE searchNextRecOrder
SELECT curpeoporder
SKIP
LOCATE REST FOR LOWER(ALLTRIM(search_cx))$LOWER(fiopeop)
IF FOUND()
   DO changeRowPeopOrder 
ELSE 
   frmOrd.butSearchRec.Visible=.T.
   frmOrd.butSearchNext.Visible=.F.   
   frmOrd.tBoxSearch.SetFocus   
ENDIF 
************************************************************************************************************************
PROCEDURE keyPressSearch
DO CASE
   CASE LASTKEY()=27
        DO searchRetOrder
   CASE LASTKEY()=13        
        DO searchRecOrder
ENDCASE 
