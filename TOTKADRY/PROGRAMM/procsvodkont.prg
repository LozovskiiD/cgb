fSupl=CREATEOBJECT('FORMSUPL')
log_term=.T.
logWord=.F.
kvo_page=1
page_beg=1
page_end=999
term_ch=.T.
dateBeg=DATE()
WITH fSupl
     .Caption='Мониторинг контрактной формы найма'
     DO addshape WITH 'fSupl',1,20,20,150,400,8     
     DO adlabMy WITH 'fSupl',1,'по состоянию на',.Shape1.Top+20,20,100,0,.T.  
     DO adtbox WITH 'fSupl',1,.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('99/99/99999'),dHeight,'dateBeg',.F.,.T.,.F.
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.txtBox1.Width-.lab1.Width-10)/2        
     .txtBox1.Left=.lab1.Left+.lab1.Width+10
     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+5)
     .Shape1.Height=.txtBox1.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.T.
      
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnMonitoring WITH .T.' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'Просмотр','DO prnMonitoring WITH .F.'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','fSupl.Release','Выход из печати' 
     .Height=.shape1.Height+.Shape91.Height+.cont1.Height+80
     .Width=.Shape91.Width+40    
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************************************************************								
PROCEDURE prnMonitoring
PARAMETERS parLog
IF EMPTY(dateBeg)
   RETURN 
ENDIF 
DIMENSION dimMnt(15)
STORE 0 TO dimMnt
* 1 всего
* 2 контрактов
* 3 на 1
* 4 1-3
* 5 3-5
* 6 1% 
* 7 1%-10№
* 8 10%-30%
* 9 30%-50%
* 10 1д
* 11 2-4 дн
* 12 5дн.
SELECT * FROM people INTO CURSOR curCardMnt READWRITE
SELECT curCardMnt
*DELETE FOR dateIn>datebeg.OR.dateUv>dateBeg
SET FILTER TO !EMPTY(date_out)
DELETE FOR date_out<dateBeg
SET FILTER TO 
DELETE FOR date_In>dateBeg
DELETE FOR !SEEK(curCardMnt.num,'datjob',1)
CREATE CURSOR curPrn (totpeop N(4),totKont N(4),srok1 N(4),srok3 N(4),srok5 N(4),pers1 N(4),pers10 N(4),pers30 N(4),pers50 N(4),day1 N(4),day4 N(4),day5 N(4))
STORE 0 TO kontpens,kontinv
APPEND BLANK 
SELECT curCardMnt
SCAN ALL
     dimMnt(1)=dimMnt(1)+1
     IF dog=1
        dimMnt(2)=dimMnt(2)+1
        **   сроки заключения
        DO CASE
           CASE kTime=2  
                IF timeDog<=12
                   dimMnt(3)=dimMnt(3)+1    
                ENDIF           
           CASE timeDog=1         
                dimMnt(3)=dimMnt(3)+1                 
           CASE timeDog>1.AND.timeDog<=3
                dimMnt(4)=dimMnt(4)+1         
           CASE timeDog>3.AND.timeDog<=5
                dimMnt(5)=dimMnt(5)+1  
                                             
        ENDCASE  
        **  повышение по контракту
        DO CASE
           CASE pkont<=1
                dimMnt(6)=dimMnt(6)+1         
           CASE pkont>1.AND.pkont<=10
                dimMnt(7)=dimMnt(7)+1
           CASE pkont>10.AND.pkont<=30
                dimMnt(8)=dimMnt(8)+1
           CASE pkont>30.AND.pkont<=50                 
                dimMnt(9)=dimMnt(9)+1
        ENDCASE          
        **  дополнительный отпуск
        DO CASE
           CASE dayKont=1       
                dimMnt(10)=dimMnt(10)+1
           CASE dayKont>=2.AND.dayKont<=4
                dimMnt(11)=dimMnt(11)+1
           CASE dayKont=5 
                dimMnt(12)=dimMnt(12)+1
        ENDCASE
        kontpens=IIF(pens,kontpens+1,kontpens)
        kontinv=IIF(inv,kontinv+1,kontinv)
     ENDIF   
ENDSCAN 
SELECT curPrn
REPLACE totPeop WITH dimMnt(1),totkont WITH dimMnt(2),srok1 WITH dimMnt(3),srok3 WITH dimMnt(4),srok5 WITH dimMnt(5),pers1 WITH dimMnt(6),pers10 WITH dimMnt(7),;
        pers30 WITH dimMnt(8),pers50 WITH dimMnt(9),day1 WITH dimMnt(10),day4 WITH dimMnt(11),day5 WITH dimMnt(12)    
DO procForPrintAndPreview WITH 'repKontMnt','',parLog,'repKontMntToExcel'
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