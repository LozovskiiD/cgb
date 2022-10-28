logUvol=.F.
logNo=.F.
STORE CTOD('  .  .   ') TO dBeg,dEnd,dBegWork
DIMENSION  dim_dek(3)
dim_dek(1)=1   &&включать б/л
dim_dek(2)=0   &&включать д/о+б/л
dim_dek(3)=0   &&исключать д/о+б/л
DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='выплата на оздоровление'
     .Icon='kone.ico'
     .procexit='Do exitHealth' 
     DO procObjFlt     
     .checkKat.Enabled=.F.
     
     DO addshape WITH 'fSupl',21,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
       
     DO addOptionButton WITH 'fSupl',11,'включать б/л',.shape21.Top+10,.shape21.Left+20,'dim_dek(1)',0,"DO procValOption WITH 'fSupl','dim_dek',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'включать д/о+б/л',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_dek(2)',0,"DO procValOption WITH 'fSupl','dim_dek',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'исключать д/о+б/л',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_dek(3)',0,"DO procValOption WITH 'fSupl','dim_dek',3",.T. 
          
     .Option11.Left=.Shape1.Left+(.shape21.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10 
     .Option13.Left=.Option12.Left+.Option12.Width+10 

     .shape21.Height=.Option11.Height+20
     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.shape21.Top+.shape21.Height+10,150,.Shape1.Width,8     
     
     DO adLabMy WITH 'fSupl',1,'начало отпуска за период с',.Shape2.Top+10,.Shape2.Left,.Shape1.Width,0,.T.,1  
     DO adLabMy WITH 'fSupl',2,' по ',.Shape2.Top+10,.Shape2.Left,.Shape1.Width,0,.T.,1  
     
     DO adtbox WITH 'fSupl',1,.Shape1.Left+20,.Shape2.Top+20,RetTxtWidth('99/99/99999'),dHeight,'dBeg',.F.,.T.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dEnd',.F.,.T.,.F.
       
     .lab1.Left=.Shape2.Left+(.Shape2.Width-.lab1.Width-.txtBox1.Width-.lab2.Width-.txtBox2.Width-15)/2
     .txtBox1.Left=.lab1.Left+.lab1.Width+5
     .lab2.Left=.txtBox1.Left+.txtBox1.Width+5
     .txtBox2.Left=.lab2.Left+.lab2.Width+5
     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+5)
     .lab2.Top=.lab1.Top     
     
     DO adLabMy WITH 'fSupl',3,'исключать принятых после',.txtBox1.Top+.txtBox1.Height+10,.Shape2.Left,.Shape1.Width,0,.T.,1  
     DO adtbox WITH 'fSupl',3,.Shape2.Left+20,.txtBox1.Top+.txtBox1.Height+10,RetTxtWidth('99/99/99999'),dHeight,'dBegWork',.F.,.T.,.F.
     .lab3.Left=.Shape2.Left+(.Shape2.Width-.lab3.Width-.txtBox3.Width-5)/2
     .txtBox3.Left=.lab3.Left+.lab3.Width+5  
     .lab3.Top=.txtBox3.Top+(.txtBox3.Height-.lab3.Height+5)
            
     DO adCheckBox WITH 'fSupl','checkNo','список получиших',.txtBox3.Top+.txtBox3.Height+10,.Shape2.Left,150,dHeight,'logNo',0,.T.      
     DO adCheckBox WITH 'fSupl','checkUvol','включать уволенных',.checkNo.Top+.checkNo.Height+10,.Shape2.Left,150,dHeight,'logUvol',0,.T.  
     .checkNo.Left=.Shape2.Left+(.Shape2.Width-.checkNo.Width)/2   
     .checkUvol.Left=.Shape2.Left+(.Shape2.Width-.checkUvol.Width)/2   
     .Shape2.Height=.txtBox1.Height*2+.checkUvol.Height*2+70
     
     DO adSetupPrnToForm WITH .Shape2.Left,.Shape2.Top+.Shape2.Height+10,.Shape2.Width,.F.,.T. 
     
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH 
      
     DO adButtonPrnToForm WITH 'DO prnSpisHealth WITH 1','DO prnSpisHealth WITH 2','DO exitHealth',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width  
                      
     .Height=.Shape1.Height+.shape21.Height+.Shape2.Height+.Shape91.Height+.butPrn.Height+70
     .Width=.Shape1.Width+20
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitHealth
SELECT dopPodr
USE
SELECT dopKat
USE
SELECT dopDolj
USE
IF USED('allpeople')
   SELECT allpeople
   USE
ENDIF    
SELECT people
frmTop.Refresh  
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus           
fSupl.Release
******************************************************************************************************************************
PROCEDURE prnSpisHealth
PARAMETERS par1
IF USED('curprn')
   SELECT curprn
   USE
ENDIF 

IF USED('curprn')
   SELECT curprn
   USE
ENDIF 

IF USED('curJobAge')
   SELECT curJobAge
   USE
ENDIF 

IF USED('curOtpBol')
   SELECT curOtpBol
   USE
ENDIF
*SELECT * FROM peoporder WHERE supord=60.AND.dateSpis>=dateBeg.AND.dateSpis<=dateEnd INTO CURSOR curOtpBol READWRITE
SELECT * FROM peoporder WHERE supord=60 INTO CURSOR curOtpBol READWRITE
SELECT curOtpBol
DELETE FOR dateBeg>dEnd
DELETE FOR dateEnd<dBeg
INDEX ON nidpeop TAG T1

SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 

SELECT * FROM people INTO CURSOR curprn READWRITE
SELECT curprn
IF logUvol
   APPEND FROM peopout FOR date_out>=dBeg
ENDIF 

ALTER TABLE curprn ADD COLUMN npp N(4)
ALTER TABLE curprn ADD COLUMN kp N(3)
ALTER TABLE curprn ADD COLUMN kd N(3)
ALTER TABLE curprn ADD COLUMN npodr c(100)
ALTER TABLE curprn ADD COLUMN ndolj c(100)
SELECT * FROM peoporder WHERE dateBeg>=dBeg.AND.dateBeg<=dEnd.AND.!EMPTY(ordSupl).AND.INLIST(supOrd,51,52) INTO CURSOR curvp READWRITE
SELECT curvp
INDEX ON nidpeop TAG T1
SELECT curprn
IF logNo
   DELETE FOR !SEEK(nid,'curvp',1)
ELSE 
   DELETE FOR SEEK(nid,'curvp',1)
   IF !EMPTY(dBegWork)
      DELETE FOR date_in>dBegWork
   ENDIF
ENDIF 
SELECT curprn
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,npodr WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),ndolj WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL


DO CASE
   CASE dim_dek(1)=1
        DELETE FOR dekOtp.AND.!SEEK(nid,'curOtpBol',1)
   CASE dim_dek(1)=2
        
   CASE dim_dek(3)=1
        DELETE FOR dekOtp
ENDCASE

IF kvoPodr>0
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
IF kvoDolj>0
   DELETE FOR !(','+LTRIM(STR(kd))+','$fltDolj)   
ENDIF
IF kvoKat>0
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltKat)
ENDIF

SELECT curprn
INDEX ON fio TAG T1
nppcx=0
SCAN ALL
     nppcx=nppcx+1
     REPLACE npp WITH nppcx
ENDSCAN
GO TOP
DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'repspishealth','список лиц получивших выплату на оздоровление',.T.,'spisHealthToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'repspishealth','список лиц получивших выплату на оздоровление',.F.
ENDCASE 
*******************************************************************************
PROCEDURE spisHealthToExcel
DO startPrnToExcel WITH 'fSupl'  
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=4
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=8
     .Columns(4).ColumnWidth=40
     .Columns(5).ColumnWidth=40
        
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО сотрудника'              
     .cells(2,3).Value='принят'
     .cells(2,4).Value='прдразд.'
     .cells(2,5).Value='должность'
    
     numberRow=3 
     SELECT curPrn
     DO storezeropercent
     SCAN ALL             
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio
          .cells(numberRow,3).Value=IIF(!EMPTY(date_In),DTOC(date_In),'')
          .cells(numberRow,4).Value=npodr
          .cells(numberRow,5).Value=ndolj
          DO fillpercent WITH 'fSupl'  
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(2,1),.Cells(numberRow-1,5)).Select
    WITH objExcel.Selection
  *       .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH      
    .Cells(2,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'
*ON ERROR
objExcel.Visible=.T.