*log_term=.T.
IF !USED('sprtot')
   USE sprtot ORDER 1 IN 0
ENDIF 
SELECT kod,name,namesp FROM sprtot WHERE sprtot.kspr=10 INTO CURSOR curSupGrup READWRITE &&группы воинского учёта
SELECT curSupGrup
APPEND BLANK
REPLACE name WITH '- все -'
cStrGrup=curSupGrup.name
nKodGrup=curSupGrup.kod
INDEX ON kod TAG T1 
logWord=.F.
term_ch=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl         
     .Caption='Список военнообязанных'
     
     DO addshape WITH 'fSupl',1,10,10,150,400,8     
     DO adtBoxAsCont WITH 'fSupl','contGru',.Shape1.Top+10,.Shape1.Left+10,RetTxtWidth('Wгруппв учетаW'),dHeight,'группа учёта',0,1   
     DO addComboMy WITH 'fSupl',1,.contGru.Left+.contGru.Width-1,.contGru.Top,dHeight,250,.T.,'cStrGrup','ALLTRIM(curSupGrup.name)',6,'DO focusSpisGrup','nKodGrup=curSupGrup.kod',.F.,.T.     
     .Shape1.Height=.contGru.Height+20
     .Shape1.Width=.contGru.Width+.comboBox1.Width+20
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.F.,.T. 
     DO adButtonPrnToForm WITH 'DO prnSpisArmy WITH .T.','DO prnSpisArmy WITH .F.','fsupl.Release'
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width 
     .Width=.Shape1.Width+20
     .Height=.butPrn.Top+.butPrn.Height+10
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************************
PROCEDURE focusSpisGrup
****************************************************************************************************************************
PROCEDURE prnSpisArmy
PARAMETERS parLog
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curJobAge')
   SELECT curJobAge
   USE
ENDIF
IF !USED('datarmy')
   USE datarmy IN 0
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 
SELECT * FROM people WHERE SEEK(nid,'datarmy',2).AND.EMPTY(datarmy.datesn) INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN kp N(3)
ALTER TABLE curPrn ADD COLUMN namep C(100)
ALTER TABLE curPrn ADD COLUMN kd N(3)
ALTER TABLE curPrn ADD COLUMN named C(100)
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN zv C(50)
ALTER TABLE curPrn ADD COLUMN grupu N(1)
ALTER TABLE curPrn ADD COLUMN rik C(30)
ALTER TABLE curPrn ADD COLUMN profil C(20)
ALTER TABLE curPrn ADD COLUMN vus C(130)
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,namep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') all
REPLACE zv WITH IIF(SEEK(nid,'datarmy',2),datarmy.zv,''),grupu WITH datarmy.grupu,rik WITH datarmy.rik,vus WITH datarmy.vus,profil WITH datarmy.profil ALL

IF nKodGrup#0
   DELETE FOR grupu#nKodGrup
ENDIF
INDEX ON fio TAG T1
nppcx=1
SCAN ALL
     SELECT curPrn      
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO procForPrintAndPreview WITH 'repspisarmy','',parLog,'repspisArmyToExcel'
********************************************************************************************************************************
PROCEDURE repSpisArmyToExcel
ON ERROR DO erSup
DO startPrnToExcel WITH 'fSupl'   
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=12
     .Columns(4).ColumnWidth=15
     .Columns(5).ColumnWidth=25
     .Columns(6).ColumnWidth=35     
     .Columns(7).ColumnWidth=12 
     .Columns(8).ColumnWidth=20
     .Columns(9).ColumnWidth=12       
     .Columns(10).ColumnWidth=15  
     .Columns(11).ColumnWidth=35  
     
         
     .cells(2,1).Value='№'                                     
     .cells(2,2).Value='ФИО'              
     .cells(2,3).Value='дата.рож.'
     .cells(2,4).Value='звание'
     .cells(2,5).Value='военкомат'   
      
     .cells(2,6).Value='адрес'
     .cells(2,7).Value='телефон'    
     .cells(2,8).Value='учебное заведение'    
     .cells(2,9).Value='дата окончания'  
             
     .cells(2,10).Value='ВУС, профиль'    
     .cells(2,11).Value='должность, подр.'  
       
     .Range(.Cells(1,1),.Cells(1,11)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter      
          .WrapText=.T.
          .Value='Список военнообязанных'          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,11)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter      

     numberRow=3
     SELECT curPrn
     DO storezeropercent
     SCAN ALL        
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=ALLTRIM(fio)
          .cells(numberRow,3).Value=IIF(!EMPTY(age),DTOC(age),'')
          .cells(numberRow,4).Value=ALLTRIM(zv)
          .cells(numberRow,5).Value=ALLTRIM(rik)
          .cells(numberRow,6).Value=ALLTRIM(ppreb)
          .cells(numberRow,7).Value=ALLTRIM(telhome)
          .cells(numberRow,8).Value=ALLTRIM(school)
          .cells(numberRow,9).Value=IIF(!EMPTY(endeduc),DTOC(endeduc),'')        
          .cells(numberRow,10).Value=ALLTRIM(vus)+' '+ALLTRIM(profil)
          .cells(numberRow,11).Value=ALLTRIM(named)+' '+ALLTRIM(namep)
          
          DO fillpercent WITH 'fSupl'
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,11)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .VerticalAlignment=1   
         .HorizontalAlignment=xlLeft
         .WrapText=.T.
         .Font.Name='Times New Roman'   
         .Font.Size=10
    ENDWITH 
    .Range(.Cells(1,1),.Cells(2,11)).Select
    objExcel.Selection.HorizontalAlignment=xlCenter
    .Range(.Cells(1,1),.Cells(1,11)).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'  
ON ERROR            
objExcel.Visible=.T. 