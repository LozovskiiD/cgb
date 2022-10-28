IF datjob.mtokl=0
   RETURN
ENDIF
IF USED('peopold')
   SELECT peopold
   USE
ENDIF
IF USED('fondold')
   SELECT fondold
   USE
ENDIF
IF USED('curnew')
   SELECT curnew
   USE
ENDIF
IF USED('curold')
   SELECT curold
   USE
ENDIF
min1=41
logFree=.F.
logSave=.T.
IF !FILE('pathold.mem')
   newpathold=GETDIR('','','Укажите путь к каталогу',64)
   pathold=IIF(!EMPTY(newpathold),newpathold,pathold)
   var_path=FULLPATH('tarif.fxp')    
   var_pathold=LEFT(var_path,LEN(var_path)-9)+'pathold'  
   SAVE TO &var_pathold ALL LIKE pathold     
   FLUSH  
ENDIF
SELECT datJob
SCATTER TO dimOld
persnum=0
persHighNum=0
RESTORE FROM pathold ADDITIVE
peopold=pathold+'people.dbf'
raspold=pathold+'rasp.dbf'
podrold=pathold+'sprpodr.dbf'
doljold=pathold+'sprdolj.dbf'
fondold=pathold+'tarfond.dbf'
*koefold=pathold+'sprkoef.dbf'
USE &peopold ALIAS peopold ORDER 1 IN 0
SELECT peopold
GO BOTTOM
numFree=num+1
USE &fondold ALIAS fondold IN 0
*USE &koefold ALIAS koefOld IN 0
SELECT * FROM fondold WHERE !EMPTY(snew) INTO CURSOR curold READWRITE
ALTER TABLE curold ADD COLUMN colpers N(8,2)
ALTER TABLE curold ADD COLUMN colsum N(12,2)
SELECT curold
INDEX ON num TAG T1
GO BOTTOM
oldForm=ALLTRIM(formula)
newForm='mtokl+mstsum+mkonts+mvto+mkat+mamb+msmp+mchir+mint+mmain+mmrek+mmain1+mfap+mglav+mmain2+mols+mosob+mslwork+mboss+mhigh+msel+mcharw+matt+mprem+mmat+mozd'
formsup='tdk+dop6s+dop7s+dop2s+dop3s+dop1s+dop4s+dop13s+dop8s+dop9s+dop10s+fs632+st_sum+dop12s'
formsup1='msf+nad632s+sumkat+svto+smol+spom+skach+sprosn+sprsup+sumvr9'

GO TOP
SELECT * FROM tarfond WHERE nblock=2 INTO CURSOR curnew READWRITE
ALTER TABLE curnew ADD COLUMN colpers N(7,2)
ALTER TABLE curnew ADD COLUMN colsum N(10,2)
SELECT curnew
INDEX ON num TAG T1
LOCATE FOR ALLTRIM(fpers)='datjob.kf'
REPLACE sum_fm WITH 'namekf'
GO BOTTOM
REPLACE sum_fm WITH 'msf'
fNewOld=CREATEOBJECT('FORMSUPL')
podrname=IIF(SEEK(datjob.kp,'sprpodr',1),ALLTRIM(sprpodr.name),'')
doljname=IIF(SEEK(datjob.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' '+LTRIM(STR(datjob.kse,5,2))

WITH fNewOld
     .Caption='Сравнение действующих и новых условий оплаты труда - '+ALLTRIM(people.fio)
     .Icon='money.ico'
     .Width=800
     .Height=750
      DO adTboxAsCont WITH 'fNewOld','txtpodr',0,0,.Width,dHeight,podrname,2,1,.T.      
      DO adTboxAsCont WITH 'fNewOld','txtdolj',0,.txtpodr.Top+.txtpodr.Height-1,.Width,dHeight,doljname,2,1,.T.      
      DO adTboxAsCont WITH 'fNewOld','txtreal',0,.txtdolj.Top+.txtDolj.Height-1,.Width/2,dHeight,'действующие условия оплаты',2,1,.T.      
      DO adTboxAsCont WITH 'fNewOld','txtnew',.txtReal.Left+.txtReal.Width-1,.txtReal.Top,.Width-.txtReal.Width+1,dHeight,'новые условия оплаты',2,1,.T.      
      
      .AddObject('grdOld','GridMyNew')   
      WITH .grdOld
          .Top=.Parent.txtReal.Top+.Parent.txtReal.Height-1
          .Left=.Parent.txtReal.Left       
          .Width=.Parent.txtReal.Width
          .height=.Parent.Height-.Parent.txtpodr.Height*3
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curOld'
          DO addColumnToGrid WITH 'fNewOld.grdOld',4
          .Column1.ControlSource='curOld.rec'
          .Column2.ControlSource='curOld.colpers'
          .Column3.ControlSource='curOld.colsum'
          .Column1.Header1.Caption='тариф' 
          .Column2.Header1.Caption='%'
          .Column3.Header1.Caption='сумма'
          .Column2.Width=RetTxtWidth('W124.W')
          .Column3.Width=RetTxtWidth('999999999999')
          .Columns(.ColumnCount).Width=0
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Alignment=0
          .Column2.Alignment=1
          .Column3.Alignment=1
          .Column1.Enabled=.F.
          .Column2.Enabled=.T.
          .Column3.Enabled=.T.
          .Column4.Enabled=.F.
          .Column2.Sparse=.T.
          .Column3.Sparse=.T.
          .Column2.Format='Z'          
          .Column3.Format='Z'
           DO gridSizeNew WITH 'fNewOld','grdOld','shapeingrid',.T. 
     ENDWITH   
     DO myColumnTxtBox WITH 'fNewOld.grdOld.column2','txtbox2','curOld.colPers',.F.,.F.,.F.,'DO validColPersOld' 
     DO myColumnTxtBox WITH 'fNewOld.grdOld.column3','txtbox3','curOld.colSum',.F.,.F.,.F.,'DO validColSumOld' 
      
      .AddObject('grdNew','GridMyNew')   
      WITH .grdNew
          .Top=.Parent.txtNew.Top+.Parent.txtNew.Height-1
          .Left=.Parent.txtNew.Left       
          .Width=.Parent.txtNew.Width
          .height=.Parent.Height-.Parent.txtpodr.Height*3
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curNew'
          DO addColumnToGrid WITH 'fNewOld.grdNew',4
          .Column1.ControlSource='curNew.rec'
          .Column2.ControlSource='curNew.colpers'
          .Column3.ControlSource='curNew.colsum'
          .Column1.Header1.Caption='тариф' 
          .Column2.Header1.Caption='%'
          .Column3.Header1.Caption='сумма'
          .Column2.Width=RetTxtWidth('W124.W')
          .Column3.Width=RetTxtWidth('999999999999')
          .Columns(.ColumnCount).Width=0
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Alignment=0
          .Column2.Alignment=1
          .Column3.Alignment=1
          .Column2.Format='Z'
          .Column3.Format='Z'
          .Column1.Enabled=.F.
          .Column2.Enabled=.T.
          .Column3.Enabled=.F.
          .Column4.Enabled=.F.
          .Column2.Sparse=.T.
           DO gridSizeNew WITH 'fNewOld','grdNew','shapeingrid2',.T.
     ENDWITH 
     DO myColumnTxtBox WITH 'fNewOld.grdNew.column2','txtbox2','curNew.colPers',.F.,.F.,.F.,'DO validColPers' 
     
     DO adTboxAsCont WITH 'fNewOld','txtdif',0,.grdOld.Top+.grdOld.Height-1,.Width,dHeight,'bcbcb',2,1,.T.   
     DO adCheckBox WITH 'fNewOld','check1','автоматически записывать при выходе',.txtDif.Top+.txtDif.Height+10,0,150,dHeight,'logSave',0,.F.
     .check1.Left=(.Width-.check1.Width)/2
     *----кнопки поиск --------------------------------------------------------------------------
     DO addContLabel WITH 'fNewOld','butFind',(.Width-RetTxtWidth('WсохранитьW')*3-20)/2,.check1.Top+.check1.Height+10, RetTxtWidth('WсохранитьW'),dHeight+3,'поиск','DO findnewold'              
     *------------------------ кнопка печать
     DO addContLabel WITH 'fNewOld','butPrn',.butFind.Left+.butFind.Width+10,.butFind.Top,.butFind.Width,dHeight+3,'печать','DO oldnewprn' 
     *------------------------ кнопки возврат
     DO addContLabel WITH 'fNewOld','butRet',.butPrn.Left+.butPrn.Width+10,.butFind.Top,.butFind.Width,dHeight+3,'возврат','DO oldnewret' 
     
     .Height=.Height+.butFind.Height+.txtDif.Height+.check1.Height+30
       
     .Autocenter=.T.     
ENDWITH
DO repsumoldnew
fNewOld.Show
*********************************************************************************************
PROCEDURE validColPersOld
IF curOld.logNew.AND.!logFree
 * repSumNew=datjob.msf-datjob.
   oldrec=RECNO()   
   repPers=pNew
   repSum=sNew
   repForm=formula
   SELECT peopold
   REPLACE &repSum WITH IIF(!curOld.logNewSum,0,&repSum),&repPers WITH 0
   REPLACE &repPers WITH curOld.colPers,&repsum WITH &repForm,sFond WITH &oldForm   
   repOldSum='peopold.'+repSum
   SELECT curOld
   REPLACE colsum WITH &repOldSum 
   GO BOTTOM
   REPLACE colSum WITH peopold.sfond  
   GO oldRec
   sumtotold=peopold.sfond
   
   
   sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh   
   IF datjob.mslwork<0 
      sumtotnew=datjob.msf+ABS(datjob.mslwork)
   ENDIF
   
   
   
   
  * sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh  
    
   sumdif=sumtotold-sumtotnew
   persNum=sumdif/datjob.mtokl*100
         
  
   IF persNum<=200   
      persdif=LTRIM(STR(persNum,8,2))+'%'   
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
      REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'
      REPLACE colpers WITH 0,colsum WITH 0 
      GO BOTTOM 
      REPLACE colsum WITH sumTotNew+datjob.mtokl*persNum/100     
      SELECT datjob
      REPLACE pslwork WITH persnum,mslwork WITH mtokl*pSlWork/100,pHigh WITH 0,sHigh WITH 0,mHigh WITH 0,msf WITH &newForm    
   ELSE 
      persHigh=persNum-200
      persNum=200
      persdif=LTRIM(STR(persNum,8,2))+'%'   
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
      REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
      sumWork=colsum
      
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'     
      REPLACE colpers WITH persHigh,colsum WITH datjob.mtokl*persHigh/100 
      sumhigh=colsum      
      
      SELECT datJob
      REPLACE pHigh WITH persHigh,mHigh WITH sumHigh,mslWork WITH sumWork,pSlWork WITH persNum,slWork WITH sumWork,msf WITH &newForm            
      SELECT curNew
      GO BOTTOM 
      REPLACE colsum WITH datjob.msf
   ENDIF   
  
   SELECT curOld
   
   namedif='за сложность и напряженность - '+persDif
   fNewOld.txtDif.ControlSource='namedif' 
   fNewOld.txtDif.Refresh
   SELECT curOld
ENDIF   
IF logFree
   SELECT curOld
   oldrec=RECNO()   
   DO CASE
      CASE ALLTRIM(curold.f_name)='tdk'
      CASE ALLTRIM(curold.f_name)='sumvr9'
      CASE ALLTRIM(curold.f_name)='pkach'
      CASE ALLTRIM(curold.f_name)='pprsup' 
      CASE ALLTRIM(curold.f_name)='pkat'
           SELECT peopold
           REPLACE pkat WITH curOld.colPers
           REPLACE sumkat WITH tdk*pkat/100,msf WITH &formSup,sFond WITH &formsup1
           
           SELECT curOld
           REPLACE colsum WITH peopold.sumkat          
           *GO BOTTOM
           *REPLACE colSum WITH peopold.sfond  
           *GO oldRec  
      CASE ALLTRIM(curold.f_name)='pvto'  
           SELECT peopold
           REPLACE pvto WITH curOld.colPers
           REPLACE svto WITH tdk*pvto/100,msf WITH &formSup,sFond WITH &formsup1
           SELECT curOld
           REPLACE colsum WITH peopold.svto 
             
      CASE ALLTRIM(curold.f_name)='pmol'
           SELECT peopold
           REPLACE pmol WITH curOld.colPers
           REPLACE smol WITH min1/100*pmol,msf WITH &formSup,sFond WITH &formsup1
           SELECT curOld
           REPLACE colsum WITH peopold.smol 
      
      CASE !EMPTY(curold.formula)              
           oldrec=RECNO()   
           repPers=f_name
           repSum=sum_f
           repForm=formula
           SELECT peopold
           REPLACE &repPers WITH 0
           REPLACE &repPers WITH curOld.colPers,&repsum WITH &repform,msf WITH &formSup,sFond WITH &formsup1
           repOldSum='peopold.'+repSum
           SELECT curOld
           REPLACE colsum WITH &repOldSum 
           GO BOTTOM
           REPLACE colSum WITH peopold.sfond  
           GO oldRec  
   ENDCASE
   SELECT curOld
   oldRec=RECNO()
   LOCATE FOR ALLTRIM(curold.f_name)='total'
   REPLACE colsum WITH peopold.msf
   LOCATE FOR ALLTRIM(curold.f_name)='sfond'  
   REPLACE colsum WITH peopold.sfond            
   GO oldRec
   sumtotold=peopold.sfond
   sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh      
   IF datjob.mslwork<0   
      sumtotnew=datjob.msf+ABS(datjob.mslwork)
   ENDIF    
   *sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh   
   sumdif=sumtotold-sumtotnew
   persNum=sumdif/datjob.mtokl*100
   IF persNum<=200   
      persdif=LTRIM(STR(persNum,8,2))+'%'   
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
      REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'
      REPLACE colpers WITH 0,colsum WITH 0 
      GO BOTTOM 
      REPLACE colsum WITH sumTotNew+datjob.mtokl*persNum/100     
      SELECT datjob
      REPLACE pslwork WITH persnum,mslwork WITH mtokl*pSlWork/100,pHigh WITH 0,sHigh WITH 0,mHigh WITH 0,msf WITH &newForm
   ELSE 
      persHigh=persNum-200
      persNum=200
      persdif=LTRIM(STR(persNum,8,2))+'%'   
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
      REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
      sumWork=colsum
      
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'     
      REPLACE colpers WITH persHigh,colsum WITH datjob.mtokl*persHigh/100 
      sumhigh=colsum      
    
      SELECT datJob
      REPLACE pHigh WITH persHigh,mHigh WITH sumHigh,mslWork WITH sumWork,pSlWork WITH persNum,slWork WITH sumWork,msf WITH &newForm            
      SELECT curNew
      GO BOTTOM 
      REPLACE colsum WITH datjob.msf
   ENDIF    
   
   namedif='за сложность и напряженность - '+persDif
   fNewOld.txtDif.ControlSource='namedif' 
   fNewOld.txtDif.Refresh
   SELECT curOld   

ENDIF
*********************************************************************************************
PROCEDURE validColSumOld
IF !logFree
   IF curOld.logNewSum.AND.colPers=0
    * repSumNew=datjob.msf-datjob.
      oldrec=RECNO()   
      repPers=pNew
      repSum=sNew
      repForm=formula
      SELECT peopold
      REPLACE &repPers WITH 0
      REPLACE &repPers WITH curOld.colPers,&repsum WITH curold.colsum,sFond WITH &oldForm   
      repOldSum='peopold.'+repSum
      SELECT curOld
      REPLACE colsum WITH &repOldSum 
      GO BOTTOM
      REPLACE colSum WITH peopold.sfond  
      GO oldRec
      sumtotold=peopold.sfond
      sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh  
      
      IF datjob.mslwork<0   
         sumtotnew=datjob.msf+ABS(datjob.mslwork)
      ENDIF     
           
   *  sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh   
      sumdif=sumtotold-sumtotnew
      persNum=sumdif/datjob.mtokl*100
      IF persNum<=200   
         persdif=LTRIM(STR(persNum,8,2))+'%'   
         SELECT curNew
         LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
         REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
         LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'
         REPLACE colpers WITH 0,colsum WITH 0 
         GO BOTTOM 
         REPLACE colsum WITH sumTotNew+datjob.mtokl*persNum/100     
         SELECT datjob
         REPLACE pslwork WITH persnum,mslwork WITH mtokl*pSlWork/100,pHigh WITH 0,sHigh WITH 0,mHigh WITH 0,msf WITH &newForm
      ELSE 
         persHigh=persNum-200
         persNum=200
         persdif=LTRIM(STR(persNum,8,2))+'%'   
         SELECT curNew
         LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
         REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
         sumWork=colsum
      
         LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'     
         REPLACE colpers WITH persHigh,colsum WITH datjob.mtokl*persHigh/100 
         sumhigh=colsum      
      
         SELECT datJob
         REPLACE pHigh WITH persHigh,mHigh WITH sumHigh,mslWork WITH sumWork,pSlWork WITH persNum,slWork WITH sumWork,msf WITH &newForm            
         SELECT curNew
         GO BOTTOM 
         REPLACE colsum WITH datjob.msf
      ENDIF    
   
      namedif='за сложность и напряженность - '+persDif
      fNewOld.txtDif.ControlSource='namedif' 
      fNewOld.txtDif.Refresh
      SELECT curOld
   ELSE   
      repSum=ALLTRIM(curOld.sNew)
      SELECT peopold
      repsumsh=&repsum
      SELECT curold
      IF colSum#repsumsh
         REPLACE colsum WITH repsumsh
      ENDIF    
   ENDIF              
ELSE 
   DO CASE
      CASE ALLTRIM(curold.f_name)='tdk'
           oldRec=RECNO()
           SELECT peopold
           REPLACE tdk WITH curold.colsum
           REPLACE msf WITH &formSup,sFond WITH &formsup1  
           
      CASE ALLTRIM(curold.f_name)='pkach'
           SELECT peopold
           REPLACE skach WITH curold.colsum,sFond WITH &formsup1                      
      CASE ALLTRIM(curold.f_name)='pprsup'
           SELECT peopold
           REPLACE sprsup WITH curold.colsum,sFond WITH &formsup1  
             
      CASE ALLTRIM(curold.f_name)='sumvr9'
           SELECT peopold
           REPLACE sprsup WITH curold.colsum,sFond WITH &formsup1 
   ENDCASE
   SELECT curOld
   oldRec=RECNO()
   LOCATE FOR ALLTRIM(curold.f_name)='total'
   REPLACE colsum WITH peopold.msf
   LOCATE FOR ALLTRIM(curold.f_name)='sfond'  
   REPLACE colsum WITH peopold.sfond            
   GO oldRec    
   sumtotold=peopold.sfond
   sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh      
   IF datjob.mslwork<0   
      sumtotnew=datjob.msf+ABS(datjob.mslwork)
   ENDIF 
 *  sumtotnew=datjob.msf-IIF(datJob.mSlWork>0,datJob.mSlWork,ABS(datJob.mslWork))-datjob.mHigh   
   sumdif=sumtotold-sumtotnew
   persNum=sumdif/datjob.mtokl*100
   IF persNum<=200   
      persdif=LTRIM(STR(persNum,8,2))+'%'   
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
      REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'
      REPLACE colpers WITH 0,colsum WITH 0 
      GO BOTTOM 
      REPLACE colsum WITH sumTotNew+datjob.mtokl*persNum/100     
      SELECT datjob
      REPLACE pslwork WITH persnum,mslwork WITH mtokl*pSlWork/100,pHigh WITH 0,sHigh WITH 0,mHigh WITH 0,msf WITH &newForm
   ELSE 
      persHigh=persNum-200
      persNum=200
      persdif=LTRIM(STR(persNum,8,2))+'%'   
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
      REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
      sumWork=colsum
      
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'     
      REPLACE colpers WITH persHigh,colsum WITH datjob.mtokl*persHigh/100 
      sumhigh=colsum      
    
      SELECT datJob
      REPLACE pHigh WITH persHigh,mHigh WITH sumHigh,mslWork WITH sumWork,pSlWork WITH persNum,slWork WITH sumWork,msf WITH &newForm            
      SELECT curNew
      GO BOTTOM 
      REPLACE colsum WITH datjob.msf
   ENDIF    
   
   namedif='за сложность и напряженность - '+persDif
   fNewOld.txtDif.ControlSource='namedif' 
   fNewOld.txtDif.Refresh
   SELECT curOld
ENDIF    
*********************************************************************************************
PROCEDURE validColPers
IF logNew
   SELECT curNew
   newrec=RECNO() 
   repPers=firead
   repSumSt=sum_f
   repSumM='datjob.'+sum_fm
   repForm=formula
   SELECT datjob
   REPLACE &repSumSt WITH 0,&repPers WITH 0,&repSumM WITH 0
   REPLACE &repPers WITH curNew.colPers,&repsumSt WITH &repForm,&repsumM WITH &repSumSt*kse,msf WITH &newForm   
   SELECT curNew
   REPLACE colsum WITH &repSumM 
   GO BOTTOM
   REPLACE colSum WITH datJob.msf  
   GO newRec   
   persdif=LTRIM(STR(datJob.pSlWork,8,2))+'%'     
   namedif='за сложность и напряженность - '+persDif
   fNewOld.txtDif.ControlSource='namedif' 
   fNewOld.txtDif.Refresh         
   SELECT curNew
   GO newRec  
ENDIF
*********************************************************************************************
PROCEDURE repsumoldnew
STORE 0 TO sumtotold,sumtotnew,sumdif,sumHigh
SELECT people
SELECT peopold
LOCATE FOR tab=datjob.tabn.AND.kp=datjob.kp.AND.kse=datjob.kse.AND.kd=datjob.kd.AND.tr=datjob.tr
IF !FOUND()
   LOCATE FOR ALLTRIM(name)=ALLTRIM(people.fio).AND.kp=datjob.kp.AND.kse=datjob.kse.AND.kd=datjob.kd.AND.tr=datjob.tr
ENDIF
logFree=IIF(FOUND(),.F.,.T.)
IF logFree
   SELECT peopold 
   APPEND BLANK   
   REPLACE num WITH numFree
ENDIF  
IF peopOld.lFree
   logFree=.T.
ENDIF  
SELECT curold
SCAN ALL
     IF !EMPTY(pnew)
         IF 'peopold'$pnew
            repcol=ALLTRIM(pNew)              
         ELSE
            repcol='peopold.'+ALLTRIM(pNew)              
         ENDIF
        REPLACE colpers  WITH &repcol
     ENDIF    
     IF !EMPTY(snew) 
        IF 'peopold'$sNew
            repcol=ALLTRIM(sNew)              
         ELSE 
            repcol='peopold.'+ALLTRIM(snew)
         ENDIF               
        REPLACE colsum  WITH &repcol
     ENDIF    
ENDSCAN
sumtotold=peopold.sfond
GO TOP
SELECT curnew
SCAN ALL
     IF !EMPTY(fpers)        
         repcol=ALLTRIM(fpers)       
         REPLACE colpers WITH &repcol         
     ENDIF     
     IF !EMPTY(sum_fm)
         repcol='datjob.'+ALLTRIM(sum_fm)              
         REPLACE colsum  WITH &repcol         
     ENDIF    
      
     SELECT curnew
ENDSCAN
sumtotnew=datjob.msf
IF datjob.pslwork=0
   sumdif=sumtotold-sumtotnew
   persNum=sumdif/datjob.mtokl*100
   IF persNum>200
      persNum=200
      sumWork=datjob.mtokl/100*persnum
      persdif=LTRIM(STR(persnum,8,2))+'%'   
      
      sumHigh=(sumDif-sumWork)
      persHighNum=sumHigh/datjob.mtokl*100
      
      SELECT curNew
      LOCATE FOR LOWER(ALLTRIM(sum_fm))='mhigh'     
      REPLACE colpers WITH persHighNum,colsum WITH sumHigh
      SELECT datJob
      REPLACE pHigh WITH persHighNum,mHigh WITH sumHigh,mslWork WITH sumWork,pSlWork WITH persNum,slWork WITH sumWork,msf WITH &newForm
      
   ELSE 
      persdif=LTRIM(STR(persnum,8,2))+'%'   
      SELECT datJob
      REPLACE mslWork WITH sumDif,pSlWork WITH persNum,slWork WITH sumdif,msf WITH &newForm
   ENDIF
      
   SELECT curNew
   LOCATE FOR LOWER(ALLTRIM(sum_fm))='mslwork'
   REPLACE colpers WITH persNum,colsum WITH datjob.mtokl*persNum/100 
   GO BOTTOM 
   REPLACE colsum WITH datJob.msf    

ELSE 
   persnum=datjob.pslwork
   persdif=LTRIM(STR(persnum,8,2))+'%' 
ENDIF    
GO TOP
namedif='за сложность и напряженность -'+persdif
fNewOld.txtDif.ControlSource='namedif'
fNewOld.txtDif.Refresh
**********************************************************************************************
PROCEDURE findnewold
fNewOld.Visible=.F.
DO formForSearsh
fNewOld.Visible=.T.
podrname=IIF(SEEK(datjob.kp,'sprpodr',1),ALLTRIM(sprpodr.name),'')
doljname=IIF(SEEK(datjob.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' '+LTRIM(STR(datjob.kse,5,2))
fNewOld.txtPodr.ControlSource='podrName'
fNewOld.txtDolj.ControlSource='doljName'
fNewOld.Caption='Сравнение действующих и новых условий оплаты труда - '+ALLTRIM(people.fio)
DO repsumoldnew
fNewOld.Refresh
**********************************************************************************************
PROCEDURE oldnewprn
CREATE CURSOR oldNewPrn (recOld C(70),pOld N(3),sOld N(12,2),recNew C(100),pNew N(6,2),sNew N(12,2))
SELECT curOld
SCAN ALL
     SELECT oldNewPrn
     APPEND BLANK
     REPLACE recOld WITH curOld.rec,pOld WITH curOld.colpers,sOld WITH curOld.colSum
     SELECT curOld 
ENDSCAN 
GO TOP 
SELECT curNew
GO TOP 
SCAN ALL
     SELECT oldNewPrn
     LOCATE FOR EMPTY(recNew)
     IF !FOUND()
        APPEND BLANK
     ENDIF    
     REPLACE recNew WITH curNew.rec,pNew WITH curNew.colpers,sNew WITH curNew.colSum     
     SELECT curNew
ENDSCAN
SELECT curNew
GO TOP
SELECT oldNewPrn
GO TOP 
DO printreport WITH 'newoldrep','hjkl','oldnewprn'
**********************************************************************************************
PROCEDURE oldnewret
SELECT peopold
IF logFree
   IF logsave
      IF tdk#0
         REPLACE name WITH people.fio
         REPLACE kp WITH datjob.kp,kd WITH datjob.kd,tr WITH datjob.tr,kse WITH datjob.kse,tab WITH datjob.tabn,lFree WITH .T.,kv WITH datjob.kv    
      ENDIF
      IF tdk=0
        * DELETE
      ENDIF
   ELSE 
      IF !lFree
         DELETE
      ENDIF
   ENDIF
ENDIF
USE
SELECT fondold
USE
SELECT curnew
USE
SELECT curold
USE
*SELECT koefOld
*USE

IF !logSave
   SELECT datjob
   GATHER FROM dimOld
ENDIF   
SELECT people
*frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
fNewOld.Release


