IF !USED('dataward')
    USE dataward ORDER 1 IN 0
ELSE
   SELECT dataward
   SET ORDER TO 1
ENDIF
DIMENSION dim_ord(2)
dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим
STORE CTOD('  .  .    ') TO dBeg_cx,dEnd_cx

DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список награжденных'
     .Icon='kone.ico'
     .procexit='Do exitend' 
     DO procObjFlt     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8             
     DO adLabMy WITH 'fSupl',1,'период',.Shape2.Top+10,.Shape2.Left+5,.Shape2.Width-10,2,.F.,1 
     
     DO adtbox WITH 'fSupl',1,.Shape2.Left+20,.lab1.Top+.lab1.Height+10,RetTxtWidth('99/99/99999'),dHeight,'dBeg_cx',.F.,.T.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dEnd_cx',.F.,.T.,.F.
       
     .txtBox1.Left=.Shape2.Left+(.Shape2.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10   
               
     DO addOptionButton WITH 'fSupl',1,'режим алфавитный',.txtBox1.Top+.txtBox1.Height+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'режим штатный',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
    
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-10)/2
     .Option2.Left=.Option1.Left+.Option1.Width+10   
         
     .Shape2.Height=.lab1.height+.txtBox1.Height+.Option1.Height+40
     
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
     DO adButtonPrnToForm WITH 'DO prnAward WITH 1','DO prnAward WITH 2','DO exitEnd',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width 
     .Height=.butPrn.Top+.butPrn.Height+10
     .Width=.Shape1.Width+20
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitEnd
SELECT dopPodr
USE
SELECT dopKat
USE
SELECT dopDolj
USE
SELECT people
frmTop.Refresh  
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus           
fSupl.Release
******************************************************************************************************************************
PROCEDURE prnAward
PARAMETERS par1
IF USED('curprn')
   SELECT curprn
   USE
ENDIF 
SELECT dataward
IF USED('kontjob')
   SELECT kontjob
   USE
ENDIF 
SELECT * FROM dataward INTO CURSOR curprn READWRITE
DO CASE 
   CASE !EMPTY(dBeg_cx).AND.!EMPTY(dEnd_cx)
        DELETE FOR daward<dBeg_cx.OR.daward>dEnd_cx
   CASE !EMPTY(dBeg_cx).AND.EMPTY(dEnd_cx)
        DELETE FOR daward<dBeg_cx
   CASE EMPTY(dBeg_cx).AND.!EMPTY(dEnd_cx)
        DELETE FOR daward>dEnd_cx
ENDCASE 
ALTER TABLE curprn ADD COLUMN npp N(3)
ALTER TABLE curprn ADD COLUMN kp N(3)
ALTER TABLE curprn ADD COLUMN cnamep C(100)
ALTER TABLE curprn ADD COLUMN kd N(3)
ALTER TABLE curprn ADD COLUMN cnamed C(100)
ALTER TABLE curprn ADD COLUMN kat N(2)
ALTER TABLE curprn ADD COLUMN cFio C(60)
SELECT curprn
INDEX ON cfio TAG T1
INDEX ON STR(kp,3)+cfio TAG T2
INDEX ON nidpeop TAG T3
SELECT * FROM datjob WHERE SEEK(nidpeop,'curprn',3) INTO CURSOR kontJob READWRITE
SELECT kontJob
DELETE FOR !EMPTY(dateout)
DELETE FOR !INLIST(tr,1,3)
INDEX ON nidpeop TAG T1
SELECT curprn
SCAN ALL
     SELECT kontJob
     SEEK curprn.nidpeop
     SELECT curprn
     REPLACE kp WITH kontJob.kp,kd WITH kontJob.kd,cnamep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),cnamed WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),;
     kat WITH kontJob.kat,cFio WITH IIF(SEEK(nidpeop,'people',4),people.fio,'')
ENDSCAN
SELECT curprn
IF kvoPodr>0
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
IF kvoDolj>0
   DELETE FOR !(','+LTRIM(STR(kd))+','$fltDolj)   
ENDIF
IF kvoKat>0
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltKat)
ENDIF

DO CASE
   CASE dim_ord(1)=1
        SET ORDER TO 1
   CASE dim_ord(2)=1
        SET ORDER TO 2  
ENDCASE         
SELECT curprn
nppcx=0
SCAN ALL
     nppcx=nppcx+1
     REPLACE npp WITH nppcx
ENDSCAN
GO TOP
DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'repaward','список награжденных',.T.,'awardToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'repaward','список награжденных',.F.,.F.
ENDCASE 
********************************************************************************************************
PROCEDURE awardToExcel
DO startPrnToExcel WITH 'fSupl'   
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
maxColumn=9
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=6
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=40
     .Columns(4).ColumnWidth=45           
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=20
     .Columns(7).ColumnWidth=20
     .Columns(8).ColumnWidth=20
     .Columns(9).ColumnWidth=15
     
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО сотрудника'             
     .cells(2,3).Value='подразд.'
     .cells(2,4).Value='должность'
     .cells(2,5).Value='дата нагр.'
     .cells(2,6).Value='вид награды'
     .cells(2,7).Value='кем'
     .cells(2,8).Value='повод'
     .cells(2,9).Value='приказ'
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     objexcel.selection.HorizontalAlignment=xlCenter

     numberRow=3 
     SELECT curPrn
     DO storezeropercent
     SCAN ALL             
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=cfio        
          .cells(numberRow,3).Value=cnamep
          .cells(numberRow,4).Value=ALLTRIM(cnamed)
          .cells(numberRow,5).Value=IIF(!EMPTY(daward),DTOC(daward),'')
          .cells(numberRow,6).Value=ALLTRIM(cnaward)
          .cells(numberRow,7).Value=ALLTRIM(cobject)
          .cells(numberRow,8).Value=ALLTRIM(cprim)
          .cells(numberRow,9).Value=ALLTRIM(nord)+IIF(!EMPTY(dord),' '+DTOC(dord),'')
          DO fillpercent WITH 'fSupl' 
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(2,1),.Cells(numberRow-1,maxColumn)).Select
    WITH objExcel.Selection
         .VerticalAlignment=1
  *       .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .WrapText=.T.
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH      
    .Cells(2,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'  
objExcel.Visible=.T.