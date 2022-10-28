birth1=CTOD('  .  .   ')
birth2=CTOD('  .  .    ')
STORE .F. TO l50,l60,l70
DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сотрудников'        
     DO procObjFlt   
     DO addshape WITH 'fSupl',2,.Shape1.left,.Shape1.top+.Shape1.height+10,150,.shape1.Width,8     
      
     DO adtbox WITH 'fSupl',1,.Shape2.Left+20,.Shape2.Top+20,RetTxtWidth('99/99/99999'),dHeight,'birth1',.F.,.T.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'birth2',.F.,.T.,.F.
       
     .txtBox1.Left=.Shape2.Left+(.Shape2.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10
     
     DO adCheckBox WITH 'fSupl','check1','50 лет',.txtBox1.Top+.txtBox1.Height+10,.Shape1.Left,150,dHeight,'l50',0    
     DO adCheckBox WITH 'fSupl','check2','60 лет',.check1.Top,.Shape1.Left,150,dHeight,'l60',0    
     DO adCheckBox WITH 'fSupl','check3','70 лет',.check1.Top,.Shape1.Left,150,dHeight,'l70',0    
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width-.check2.Width-.check3.Width-20)/2
     .check2.Left=.check1.Left+.check1.Width+10
     .check3.Left=.check2.Left+.check2.Width+10       
     .Shape2.Height=.txtBox1.Height+.check1.Height+50        
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
     DO adButtonPrnToForm WITH 'DO prnBirthDay WITH .T.','DO prnBirthDay WITH .F.','fsupl.Release',.T.  
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width  
     .Height=.butPrn.Top+.butPrn.Height+20
     .Width=.Shape1.Width+20
     
     
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************************
PROCEDURE prnBirthDay
PARAMETERS parLog
IF birth2<birth1.OR.(EMPTY(birth1).AND.EMPTY(birth2))
   RETURN
ENDIF 
IF USED('curAge')
   SELECT curAge
   USE
ENDIF
IF USED('curJobAge')
   SELECT curJobAge
   USE
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON kodpeop TAG T1

month1=MONTH(birth1)
month2=MONTH(birth2)
day1=DAY(birth1)
day2=DAY(birth2)
SELECT  * FROM people WHERE MONTH(age)>=month1.AND.MONTH(age)<=month2.AND.EMPTY(date_out) INTO CURSOR curAge READWRITE
ALTER TABLE curAge ADD COLUMN npp N(4)
ALTER TABLE curAge ADD COLUMN kp N(3)
ALTER TABLE curAge ADD COLUMN npodr C(100)
ALTER TABLE curAge ADD COLUMN kd N(3)
ALTER TABLE curAge ADD COLUMN ndol C(100)
ALTER TABLE curAge ADD COLUMN years N(3)
ALTER TABLE curAge ADD COLUMN kat N(2)
ALTER TABLE curAge ADD COLUMN cprim C(50)

SELECT curAge
SCAN ALL
     SELECT curJobAge
     SEEK curAge.num
     SELECT curAge
     REPLACE kp WITH curJobAge.kp,kd WITH curJobAge.kd,npodr WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),ndol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),cprim WITH IIF(dekotp,'д/о',''),kat WITH curJobAge.kat
ENDSCAN
DO applyflt
DO CASE
   CASE MONTH(birth1)=MONTH(birth2)
        DELETE FOR MONTH(age)#MONTH(birth1)
        DELETE FOR DAY(age)<day1  
        DELETE FOR DAY(age)>day2
   CASE MONTH(birth1)#MONTH(birth2)    
        DELETE FOR MONTH(age)<MONTH(birth1)  
        DELETE FOR MONTH(age)>MONTH(birth2)
        DELETE FOR MONTH(age)=month1.AND.DAY(age)<day1
        DELETE FOR MONTH(age)=month2.AND.DAY(age)>day2
ENDCASE
INDEX ON STR(MONTH(age),2)+STR(DAY(age),2) TAG T1
IF l50.OR.l60.OR.l70
   SCAN ALL
        IF l50
           agecx=STR(DAY(age),2)+'.'+STR(MONTH(age),2)+'.'+STR(YEAR(age)+50,4)
           IF CTOD(agecx)>=birth1.AND.CTOD(agecx)<=birth2
              REPLACE years WITH 50
           ENDIF
        ENDIF 
        IF l60
           agecx=STR(DAY(age),2)+'.'+STR(MONTH(age),2)+'.'+STR(YEAR(age)+60,4)
           IF CTOD(agecx)>=birth1.AND.CTOD(agecx)<=birth2
              REPLACE years WITH 60
           ENDIF
        ENDIF 
        IF l70
           agecx=STR(DAY(age),2)+'.'+STR(MONTH(age),2)+'.'+STR(YEAR(age)+70,4)
           IF CTOD(agecx)>=birth1.AND.CTOD(agecx)<=birth2
              REPLACE years WITH 70
           ENDIF
        ENDIF 
   ENDSCAN 
   DELETE FOR years=0
ELSE
   SCAN ALL
        IF YEAR(birth1)=YEAR(birth2)
           REPLACE years WITH YEAR(birth1)-YEAR(age)
        ENDIF  
   ENDSCAN    
ENDIF 
nppcx=1
SCAN ALL
     SELECT curJobAge
     SEEK STR(curAge.num,4)
     ndolcx=IIF(SEEK(curJobAge.kd,'sprdolj',1),sprdolj.name,'')
     npodrcx=IIF(SEEK(curJobAge.kp,'sprpodr',1),sprpodr.namework,'') 
     SELECT curAge
     REPLACE npp WITH nppcx,ndol WITH ndolcx,npodr WITH npodrcx
     nppcx=nppcx+1
ENDSCAN
SELECT curAge
GO TOP
DO procForPrintAndPreview WITH 'repBirthDay','',parLog,'repBirthDayToExcel'
********************************************************************************************************************************
PROCEDURE repBirthDayToExcel
DO startPrnToExcel WITH 'fSupl'       
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=30
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=30
     .Columns(4).ColumnWidth=15
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=10
     .cells(2,1).Value='ФИО сотрудника'              
     .cells(2,2).Value='подразделение'
     .cells(2,3).Value='должность'
     .cells(2,4).Value='день рождения'                                    
     .cells(2,5).Value='лет'                                    
     .cells(2,6).Value='дата приема'                                    
     .Range(.Cells(1,1),.Cells(1,6)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .WrapText=.T.
          .Value='Список сотрудников'
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,5)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter

     numberRow=3
     yearMonth=''
     SELECT curAge
     DO storezeropercent
     SCAN ALL        
          .cells(numberRow,1).Value=fio
          .cells(numberRow,2).Value=npodr
          .cells(numberRow,3).Value=ndol
          .cells(numberRow,4).Value=LTRIM(STR(DAY(curAge.age)))+' '+month_prn(MONTH(curAge.age))
          .cells(numberRow,5).Value=years
          .cells(numberRow,6).Value=IIF(!EMPTY(date_in),DTOC(date_in),'')
          DO fillpercent WITH 'fSupl'
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,6)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'   
         .Font.Size=10
    ENDWITH 
    .Range(.Cells(1,1),.Cells(1,6)).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'               
objExcel.Visible=.T. 