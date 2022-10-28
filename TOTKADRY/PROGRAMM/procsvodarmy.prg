fSupl=CREATEOBJECT('FORMSUPL')
log_term=.T.
logWord=.F.
kvo_page=1
page_beg=1
page_end=999
term_ch=.T.
dateBeg=DATE()
WITH fSupl
     .Caption='Отчёт по военнобязанным'
     .procExit='DO exitSvodArmy'
     DO addshape WITH 'fSupl',1,20,20,150,400,8     
     DO adlabMy WITH 'fSupl',1,'по состоянию на',.Shape1.Top+20,20,100,0,.T.  
     DO adtbox WITH 'fSupl',1,.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('99/99/99999'),dHeight,'dateBeg',.F.,.T.,.F.
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.txtBox1.Width-.lab1.Width-10)/2        
     .txtBox1.Left=.lab1.Left+.lab1.Width+10
     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+5)
     .Shape1.Height=.txtBox1.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.F.
      
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnWar WITH .T.' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'Просмотр','DO prnWar WITH .F.'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','DO exitSvodArmy','Выход из печати' 
     .Height=.shape1.Height+.Shape91.Height+.cont1.Height+80
     .Width=.Shape91.Width+40    
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************************************************************								
PROCEDURE prnWar
PARAMETERS parLog
IF EMPTY(dateBeg)
   RETURN 
ENDIF 

SELECT rik FROM datarmy DISTINCT INTO CURSOR curSupRik READWRITE 
SELECT curSupRik

CREATE CURSOR curPrn (priznak N(1),namePriznak C(30),kPriznak N(2),totpeop N(4),nameprim C(50))
SELECT kod,name FROM sprtot WHERE sprtot.kspr=14 INTO CURSOR curSupZv READWRITE &&воинские звания
SELECT curSupZv
INDEX ON kod TAG T1 
SELECT curSupZv
SCAN ALL
     SELECT curPrn
     APPEND BLANK
     REPLACE priznak WITH 1,namePriznak WITH curSupZv.name,kPriznak WITH curSupZv.kod,nameprim WITH 'воинское звание'
     SELECT curSupZv
ENDSCAN
SELECT curSupRik
SCAN ALL
     SELECT curPrn
     APPEND BLANK
     REPLACE priznak WITH 2,namePriznak WITH curSupRik.rik,nameprim WITH 'военкомат'
     SELECT curSupRik
ENDSCAN
SELECT * FROM datArmy WHERE !datarmy.snu INTO CURSOR curArmy READWRITE
SELECT curArmy
SCAN ALL
     SELECT curPrn
     LOCATE FOR priznak=1.AND.kpriznak=curArmy.kzv
     IF FOUND()
        REPLACE totPeop WITH totPeop+1
     ENDIF
     LOCATE FOR priznak=2.AND.ALLTRIM(namepriznak)=ALLTRIM(curarmy.rik)
     IF FOUND()
        REPLACE totPeop WITH totPeop+1
     ENDIF
     SELECT curArmy      
ENDSCAN 

SELECT curprn
GO TOP
DO procForPrintAndPreview WITH 'repArmy','',parLog,'repKontMntToExcel'
*****************************************************************************************************************************
PROCEDURE repKontMntToExcel
#DEFINE xlCenter -4108            
*#DEFINE xlLeft -4131  
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
     .Columns(1).ColumnWidth=20
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=10
     .Columns(4).ColumnWidth=10
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=10
     .Columns(7).ColumnWidth=10
     .Columns(8).ColumnWidth=10
     .Columns(9).ColumnWidth=10
     .Columns(10).ColumnWidth=10
     .Columns(11).ColumnWidth=10
     .Columns(12).ColumnWidth=10
     .Range(.Cells(2,1),.Cells(4,1)).Select
     WITH objExcel.Selection      
          .MergeCells=.T.            
          .HorizontalAlignment= -4108   
          .WrapText=.T.        
          .Value='Общее количество работающих в учреждении'   
     ENDWITH  
     .Range(.Cells(2,2),.Cells(4,2)).Select 
     WITH objExcel.Selection      
          .MergeCells=.T.     
          .HorizontalAlignment= -4108                       
          .WrapText=.T.  
          .Value='Всего заключено контрактов'
     ENDWITH            
    .Range(.Cells(2,3),.Cells(3,5)).Select     
     WITH objExcel.Selection            
          .MergeCells=.T.     
          .HorizontalAlignment= -4108                 
          .WrapText=.T.  
          .Value='Срок заключения контрактов'
     ENDWITH          
    
     .Range(.Cells(2,6),.Cells(2,12)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='Меры материального стимулирования труда'
     ENDWITH       
     .Range(.Cells(3,6),.Cells(3,9)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='повышение тарифной ставки (оклада)'
     ENDWITH  
     .Range(.Cells(3,10),.Cells(3,12)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108                             
          .Value='предоставление дополнительного поощрительного отпуска'
          .WrapText=.T.
     ENDWITH          
     .cells(4,3).Value='на 1 год'
     .cells(4,3).WrapText=.T.    
     .cells(4,4).Value='от 1 до 3 лет'  
     .cells(4,4).WrapText=.T.    
     .cells(4,5).Value='от 3 до 5 лет'                                                                                                       
     .cells(4,5).WrapText=.T.    
     .cells(4,6).Value='1%'    
     .cells(4,6).WrapText=.T.  
     .cells(4,7).Value='от 1% до 10%'  
     .cells(4,7).WrapText=.T.  
     .cells(4,8).Value='от 10% до 30%'
     .cells(4,8).WrapText=.T.  
     .cells(4,9).Value='от 30% до 50%'    
     .cells(4,9).WrapText=.T.  
     .cells(4,10).Value='1 день'  
     .cells(4,10).WrapText=.T.  
     .cells(4,11).Value='от 2-х до 4-х дней'
     .cells(4,11).WrapText=.T.  
     .cells(4,12).Value='5 дней'
     .cells(4,12).WrapText=.T.  
    .Range(.Cells(4,3),.Cells(4,12)).Select
    objExcel.Selection.HorizontalAlignment= -4108 
     numberRow=5
     SELECT curPrn
     SCAN ALL         
          .cells(numberRow,1).Value=totPeop
          .cells(numberRow,2).Value=totKont
          .cells(numberRow,3).Value=srok1     
          .cells(numberRow,4).Value=srok3     
          .cells(numberRow,5).Value=srok5     
          .cells(numberRow,6).Value=pers1
          .cells(numberRow,7).Value=pers10
          .cells(numberRow,8).Value=pers30     
          .cells(numberRow,9).Value=pers50     
          .cells(numberRow,10).Value=day1  
          .cells(numberRow,11).Value=day4     
          .cells(numberRow,12).Value=day5  
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,12)).Select
    objExcel.Selection.VerticalAlignment= -4160    
    objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
    objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
    objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
    objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
    objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
    objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
    numberRow=numberRow+1
    .cells(numberRow,1).Value='пенсионеры'
    .cells(numberRow,2).Value=kontpens
    numberRow=numberRow+1
    .cells(numberRow,1).Value='инвалиды'
    .cells(numberRow,2).Value=kontinv    
    .Range(.Cells(1,1),.Cells(numberRow,12)).Select
    objExcel.Selection.Font.Name='Times NewRoman'    
    objExcel.Selection.Font.Size=10
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
objExcel.Visible=.T.  
********************************************************************************************
PROCEDURE exitSvodArmy
fSupl.Release