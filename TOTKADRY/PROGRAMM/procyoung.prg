DIMENSION dim_ord(2)
dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим
DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список молодых специалистов'
     .Icon='kone.ico'
     .procexit='Do exitYoung' 
     DO procObjFlt     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
        
     DO addOptionButton WITH 'fSupl',1,'режим алфавитный',.Shape2.Top+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'режим штатный',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20   
     
     .Shape2.Height=.Option1.Height+20  
     
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
     DO adButtonPrnToForm WITH 'DO prnSpisYoung WITH 1','DO prnSpisYoung WITH 2','Do exitYoung',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width  
     .Height=.butPrn.Top+.butPrn.Height+20
     .Width=.Shape1.Width+20
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitYoung
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
PROCEDURE prnSpisYoung
PARAMETERS par1
IF USED('curprn')
   SELECT curprn
   USE
ENDIF 
IF USED('youngjob')
   SELECT youngjob
   USE
ENDIF 
SELECT * FROM people WHERE mols INTO CURSOR curprn READWRITE
ALTER TABLE curprn ADD COLUMN npp N(3)
ALTER TABLE curprn ADD COLUMN kp N(3)
ALTER TABLE curprn ADD COLUMN cnamep C(100)
ALTER TABLE curprn ADD COLUMN kd N(3)
ALTER TABLE curprn ADD COLUMN cnamed C(100)
ALTER TABLE curprn ADD COLUMN cprim C(50)
ALTER TABLE curprn ADD COLUMN kat N(2)
SELECT curprn
INDEX ON fio TAG T1
INDEX ON STR(kp,3)+fio TAG T2
INDEX ON num TAG T3
SELECT * FROM datjob WHERE SEEK(kodpeop,'curprn',3).AND.tr=1 INTO CURSOR youngJob READWRITE
SELECT youngJob
DELETE FOR !EMPTY(dateout)
INDEX ON kodpeop TAG T1
SELECT curprn
SCAN ALL
     SELECT youngJob
     SEEK curprn.num
     SELECT curprn
     REPLACE kp WITH youngJob.kp,kd WITH youngJob.kd,cnamep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),cnamed WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),cprim WITH IIF(dekotp,'д/о',''),kat WITH youngJob.kat
ENDSCAN
DO applyflt
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
        DO procForPrintAndPreview WITH 'repspisyoung','список молодых специалистов',.T.,'spisYoungToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'repspisyoung','список молодых сспециалистов',.F.,'spisYoungToExcel'
ENDCASE 
*************************************************************************************************************************
PROCEDURE spisYoungToExcel
DO startPrnToExcel WITH 'fSupl'   
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=6
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=8
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=40
     .Columns(6).ColumnWidth=40
     .Columns(7).ColumnWidth=8
     
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО сотрудника'              
     .cells(2,3).Value='принят'
     .cells(2,4).Value='мол. спец до'
     .cells(2,5).Value='подразд.'
     .cells(2,6).Value='должность'
     .cells(2,7).Value='прим.'

     numberRow=3 
     SELECT curPrn
     DO storezeropercent
     SCAN ALL             
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio
          .cells(numberRow,3).Value=IIF(!EMPTY(date_in),DTOC(date_in),'')
          .cells(numberRow,4).Value=IIF(!EMPTY(dmol),DTOC(dmol),'')
          .cells(numberRow,5).Value=cnamep
          .cells(numberRow,6).Value=ALLTRIM(cnamed)
          .cells(numberRow,7).Value=ALLTRIM(cprim)
          DO fillpercent WITH 'fSupl'
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(2,1),.Cells(numberRow-1,7)).Select
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
objExcel.Visible=.T.