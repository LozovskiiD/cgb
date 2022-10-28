**********************************************************************************************************************************
DIMENSION dim_knt(3)
dim_knt(1)=1 && все
dim_knt(2)=0 && прошли
dim_knt(3)=0 && подлежат


dBeg=CTOD('  .  .   ')
dEnd=CTOD('  .  .    ')
DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сотрудников, прошедших аттестацию' 
     DO procObjFlt     
        
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8        
      
     DO adtbox WITH 'fSupl',1,.Shape1.Left+20,.Shape2.Top+20,RetTxtWidth('99/99/99999'),dHeight,'dBeg',.F.,.T.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dEnd',.F.,.T.,.F.
       
     .txtBox1.Left=.Shape1.Left+(.Shape1.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10
     
     DO addOptionButton WITH 'fSupl',11,'все',.txtBox1.Top+.txtBox1.Height+10,.Shape2.Left+20,'dim_knt(1)',0,"DO procValOption WITH 'fSupl','dim_knt',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'прошли',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_knt(2)',0,"DO procValOption WITH 'fSupl','dim_knt',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'подлежат',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_knt(3)',0,"DO procValOption WITH 'fSupl','dim_knt',3",.T. 
     .Option11.Left=.Shape2.Left+(.Shape2.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10
     .Option13.Left=.Option12.Left+.Option12.Width+10                      
     .Shape2.Height=.txtBox1.Height+.Option11.Height+50                 
     
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape2.Top+.Shape2.Height+10,.Shape1.Width,.F.,.T.
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH 
     DO adButtonPrnToForm WITH 'DO prnAt WITH .T.','DO prnAt WITH .F.','fSupl.Release',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width 
                              
     .Width=.Shape1.Width+20
     .Height=.butPrn.Top+.butPrn.Height+20
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************************
PROCEDURE prnAt
PARAMETERS parLog
IF dBeg>dEnd
   RETURN
ENDIF 
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curjobage')
   SELECT curjobage
   USE
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 

SELECT * FROM people INTO CURSOR curPrn READWRITE

*SELECT * FROM people WHERE datts>=dBeg.AND.datts<=dEnd INTO CURSOR curPrn READWRITE
*DELETE FOR EMPTY(datts)
ALTER TABLE curPrn ADD COLUMN npp N(4)
ALTER TABLE curPrn ADD COLUMN np N(4)
ALTER TABLE curPrn ADD COLUMN kp N(3)
ALTER TABLE curPrn ADD COLUMN kd N(3)
ALTER TABLE curPrn ADD COLUMN kat N(2)
ALTER TABLE curPrn ADD COLUMN npodr C(100)
ALTER TABLE curPrn ADD COLUMN ndol C(100)

SELECT curprn
INDEX ON STR(np,3)+STR(KD,3)+DTOS(datts) TAG t1 
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,npodr WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),np WITH sprpodr.np,;
        ndol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',1),rasp.kat,curJobAge.kat) ALL
DO CASE 
   CASE dim_knt(1)=1
        REPLACE datts WITH CTOD('  .  .    ') FOR datts<dBeg.OR.datts>dEnd
   CASE dim_knt(2)=1
        DELETE FOR datts<dBeg.OR.datts>dend        
   CASE dim_knt(3)=1
        DELETE FOR EMPTY(datts)     
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
        
nppcx=0
kpcx=kp
SCAN ALL
     IF kpcx#kp
        nppcx=1
        kpcx=kp 
     ENDIF 
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO procForPrintAndPreview WITH 'repAt','',parLog,'repAtToExcel'
*******************************************************************************************************************************************
PROCEDURE repatToExcel
DO startPrnToExcel WITH 'fSupl'  
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=50   
     .Columns(4).ColumnWidth=10
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО сотрудника'              
     .cells(2,3).Value='должность'
     .cells(2,4).Value='дата аттестации'                                    
     .Range(.Cells(1,1),.Cells(1,4)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .WrapText=.T.
          .Value='Аттестация на соответствие занимаемой должности c '+DTOC(dBeg)+' по '+DTOC(dEnd)
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,4)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     numberRow=3
     yearMonth=''
     SELECT curprn
     DO storezeropercent
     kpcx=0
     SCAN ALL        
          IF kp#kpcx 
             kpcx=kp         
             .Range(.Cells(numberRow,1),.Cells(numberRow,4)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlCenter
                  .WrapText=.T.
                  .Value=npodr
                  .Interior.ColorIndex=35
                   numberRow=numberRow+1 
             ENDWITH       
          ENDIF
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio
          .cells(numberRow,3).Value=ndol         
          .cells(numberRow,4).Value=IIF(!EMPTY(datts),DTOC(datts),'')  
          DO fillpercent WITH 'fSupl' 
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,4)).Select
    WITH objExcel.Selection
         .VerticalAlignment = xlTop 
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .WrapText=.T.
         .Font.Name='Times New Roman'   
         .Font.Size=10
    ENDWITH 
    .Range(.Cells(1,1),.Cells(1,4)).Select
ENDWITH
DO endPrnToExcel WITH 'fSupl'                
objExcel.Visible=.T. 