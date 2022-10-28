DIMENSION dim_ord(2)
dim_ord(1)=1 &&  календарный режим
dim_ord(2)=0 && штатный режим
dateStart=DATE()
DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Веломость по переходящему стажу'
     .Icon='kone.ico'
     .procexit='fSupl.Release' 
     DO procObjFlt     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
     
     DO adLabMy WITH 'fSupl',1,'Дата отсчёта',.Shape2.Top+20,.Shape2.Left+5,100,2,.T.,1 
     DO adtbox WITH 'fSupl',1,.lab1.Left+.lab1.Width+10,.Shape2.Top+10,RetTxtWidth('99/99/99999'),dHeight,'dateStart',.F.,.T.,.F.
      
     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+2)  
     .lab1.Left=.Shape2.Left+(.Shape2.Width-.lab1.Width-.txtBox1.Width-10)/2
     .txtBox1.Left=.lab1.Left+.lab1.Width+10        
     .Shape2.Height=.txtBox1.Height+20       
     DO adSetupPrnToForm WITH .Shape2.Left,.Shape2.Top+.Shape2.Height+10,.Shape2.Width,.F.,.T. 
     DO adButtonPrnToForm WITH 'DO prnSpisStajPer WITH 1','DO prnSpisStajPer WITH 2','fsupl.Release'     
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH 
     *-----------------------------Кнопка принять---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont11',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wпринять')*2)-15)/2,.butPrn.Top,RetTxtWidth('wпринятьw'),.butPrn.Height,'принять','DO returnToPrn WITH .T.'
     .cont11.Visible=.F.
     *---------------------------------Кнопка сброс-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont12',.cont11.Left+.cont11.Width+15,.Cont11.Top,.Cont11.Width,.butPrn.Height,'сброс','DO returnToPrn WITH .F.'     
     .cont12.Visible=.F.      
      DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width    
     .Height=.Shape1.Height+.Shape2.Height+.Shape91.Height+.butPrn.Height+60
     .Width=.Shape1.Width+20
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************
PROCEDURE prnSpisStajPer
PARAMETERS par1
IF USED('curprn')
   SELECT curprn
   USE
ENDIF
IF USED('youngjob')
   SELECT youngjob
   USE
ENDIF 
SELECT * FROM people WHERE !dekotp INTO CURSOR curPrn READWRITE 
SELECT curPrn
APPEND FROM peopout FOR date_out>dateStart.AND.!dekotp

ALTER TABLE curPrn ADD COLUMN kp N(3)
ALTER TABLE curPrn ADD COLUMN namep C(100)
ALTER TABLE curPrn ADD COLUMN kd N(3)
ALTER TABLE curPrn ADD COLUMN named C(100)
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curprn ADD COLUMN kat N(2)

INDEX ON DTOS(dperst) TAG T1
INDEX ON STR(kp,3)+STR(kd,3)+DTOS(dperst) TAG T2
INDEX ON num TAG T3
*SET ORDER TO 1

SELECT * FROM datjob WHERE SEEK(kodpeop,'curprn',3).AND.INLIST(tr,1,3) INTO CURSOR youngJob READWRITE
SELECT youngJob
INDEX ON kodpeop TAG T1
SELECT curprn
SET RELATION TO num INTO youngJob ADDITIVE
REPLACE dperst WITH CTOD('  .  .    .')
SELECT curprn
SCAN ALL   
     REPLACE kp WITH youngJob.kp,kd WITH youngJob.kd,namep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),kat WITH youngJob.kat
     DO actualStajToday WITH 'curPrn','curPrn.date_in','dateStart'
     DO perStajOne WITH 'curprn.staj_today','dateStart','curPrn'
ENDSCAN
SELECT curPrn 
DELETE FOR EMPTY(dPerst)
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
        DO procForPrintAndPreview WITH 'repstper','список молодых специалистов',.T.,'stPerToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'repstper','список молодых специалистов',.F.,'stPerToExcel'
ENDCASE
*********************************************************************************************************
PROCEDURE stPerToExcel
DO startPrnToExcel WITH 'fSupl'    
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=6
     .Columns(2).ColumnWidth=6
     .Columns(3).ColumnWidth=30
     .Columns(4).ColumnWidth=40
     .Columns(5).ColumnWidth=40
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8
     .Columns(8).ColumnWidth=8
     .Range(.Cells(1,1),.Cells(1,4)).Select
     WITH objExcel.Selection
          .MergeCells=.T.             
          .VerticalAlignment=1   
          .HorizontalAlignment=xlCenter
          .WrapText=.T.
          .Value='Список по переходящему стажу по состоянию на '+DTOC(dateStart)
          .Font.Name='Times New Roman'
          .Font.Size=10
     ENDWITH  
     
     .cells(2,1).Value='№'              
     .cells(2,2).Value='код'
     .cells(2,3).Value='ФИО сотрудника'   
     .cells(2,4).Value='подразд.'
     .cells(2,5).Value='должность'
     .cells(2,6).Value='стаж'
     .cells(2,7).Value='дата пер.'
     .cells(2,8).Value='%'      
     numberRow=3 
     SELECT curPrn
     DO storezeropercent
     SCAN ALL             
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=num
          .cells(numberRow,3).Value=fio
          .cells(numberRow,4).Value=namep
          .cells(numberRow,5).Value=ALLTRIM(named)
          .cells(numberRow,6).NumberFormat = "General"
          .cells(numberRow,6).Value=LEFT(staj_today,2)+'_'+SUBSTR(staj_today,4,2)+'_'+SUBSTR(staj_today,7,2)       
          .cells(numberRow,7).Value=DTOC(dperst)
          DO fillpercent WITH 'fSupl' 
          numberRow=numberRow+1         
     ENDSCAN
     .Range(.Cells(2,1),.Cells(numberRow-1,8)).Select
     WITH objExcel.Selection  
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin
          .Font.Name='Times New Roman'
          .Font.Size=10
     ENDWITH   
     .Range(.Cells(2,1),.Cells(2,8)).Select    
     objExcel.Selection.HorizontalAlignment=xlCenter       
     .Cells(2,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'        

objExcel.Visible=.T.
