DIMENSION dim_ved(2),dim_ord(2),dim_opt(2)
dim_ved(1)=1
dim_ved(2)=0
dim_ord(1)=1
dim_ord(2)=0
dim_opt(1)=1 && в Excel
dim_opt(2)=0 && в Word
dateSpis=DATE()
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Icon='kone.ico'  
     .Caption='Контрольный список специалистов'
     DO addShape WITH 'fSupl',1,20,20,10,400,8
     
     DO adLabMy WITH 'fSupl',1,'дата ',.Shape1.Top+20,.Shape1.Left,.Shape1.Width,0,.T.,1 
     DO adTboxNew WITH 'fSupl','boxBeg',.Shape1.Top+20,.Shape1.Left,RetTxtWidth('99/99/99999'),dHeight,'dateSpis',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3     
     .Shape1.Height=.boxBeg.Height+40  
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxBeg.Width-10)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+10
     DO addShape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Height,.Shape1.Width,8
     
     DO addOptionButton WITH 'fSupl',1,'список специалистов',.Shape2.Top+20,.Shape2.Left+20,'dim_ved(1)',0,"DO procValOption WITH 'fSupl','dim_ved',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'список должностей',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ved(2)',0,"DO procValOption WITH 'fSupl','dim_ved',2",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20
     
     DO addOptionButton WITH 'fSupl',3,'по должностям',.Option1.Top+.Option1.Height+20,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',4,'по алфавиту',.Option3.Top,.Option3.Left,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     
     .Option3.Left=.Shape2.Left+(.Shape2.Width-.Option3.Width-.Option4.Width-20)/2
     .Option4.Left=.Option3.Left+.Option3.Width+20
     .Shape2.Height=.Option1.Height*2+60
     
     DO addOptionButton WITH 'fSupl',5,'направить в Excel',.Shape2.Top+.Shape2.Height+20,.Shape2.Left+20,'dim_opt(1)',0,"DO procValOption WITH 'fSupl','dim_opt',1",.T. 
     DO addOptionButton WITH 'fSupl',6,'направить в Word',.Option5.Top,.Option3.Left,'dim_opt(2)',0,"DO procValOption WITH 'fSupl','dim_opt',2",.T. 
     .Option5.Left=.Shape2.Left+(.Shape2.Width-.Option5.Width-.Option6.Width-20)/2
     .Option6.Left=.Option5.Left+.Option5.Width+20
     
     DO addButtonOne WITH 'fSupl','butPrn',.Shape2.Left+(.Shape2.Width-RetTxtWidth('wсформировать')*2-20)/2,.Option5.Top+.Option5.Height+20,'сформировать','','DO prntopspis WITH .T.',39,RetTxtWidth('wсформироватьw'),'формирование отчёта' 
     DO addButtonOne WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+10,.butPrn.Top,'возврат','','fSupl.Release',39,.butPrn.Width,'возврат' 
     
     DO addShape WITH 'fSupl',11,.Shape2.Left,.butPrn.Top,.butPrn.Height,.Shape2.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape2.Left,.Shape2.Width,2,.F.,0 
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.  
     .Height=.Shape1.Height+.Shape2.Height+.butPrn.Height+.Option5.Height+90
     .Width=.Shape1.Width+40     
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************
PROCEDURE prntopspis
PARAMETERS par1
DO CASE
   CASE dim_ved(1)=1
        DO prnFamSpis
   CASE dim_ved(2)=1
        DO prnDolSpis
ENDCASE
****************************************************************************************************************
*                                Печать в разрезе сотрудников
****************************************************************************************************************
PROCEDURE prnFamSpis
IF !USED('boss')
   USE boss IN 0 
ENDIF
IF !USED('datkurs')
   USE datkurs IN 0 
ENDIF
IF USED('curJobspis')
   SELECT curJobSpis
   USE
ENDIF 
IF USED('curPeopList')
   SELECT curPeopList
   USE
ENDIF 
SELECT * FROM datKurs INTO CURSOR curdatKurs READWRITE
SELECT curDatKurs
INDEX ON STR(kodpeop,5)+DTOS(perBeg) TAG T1 DESC

SELECT * FROM datjob INTO CURSOR curJobSpis READWRITE
SELECT curJobSpis
APPEND FROM datjobout
SELECT curJobSpis
DELETE FOR tr#1
DELETE FOR dateBeg>dateSpis
DELETE FOR !EMPTY(dateout).AND.dateout<dateSpis
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) ALL
REPLACE kat WITH 1 FOR SEEK(kd,'sprdolj',1).AND.sprdolj.lspis
DELETE FOR !INLIST(kat,1,5)
*INDEX ON kodpeop TAG T1
INDEX ON nidpeop TAG T1

SELECT * FROM people INTO CURSOR curPeopList READWRITE
SELECT curPeopList
APPEND FROM peopout
ALTER TABLE curPeopList ADD COLUMN named C(100)
ALTER TABLE curPeopList ADD COLUMN kd N(3)
ALTER TABLE curPeopList ADD COLUMN kurs C(250)

DELETE FOR !SEEK(nid,'curJobspis',1)
REPLACE kd WITH IIF(SEEK(nid,'curJobSpis',1),curJobSpis.kd,0) ALL
REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL
REPLACE kurs WITH IIF(SEEK(STR(num,5),'curdatKurs',1),ALLTRIM(curDatKurs.namekurs)+' '+ALLTRIM(curDatKurs.nameschool)+' '+IIF(!EMPTY(curDatKurs.perBeg),DTOC(curDatKurs.perBeg)+' по '+DTOC(curDatKurs.perEnd),''),'') ALL
DO CASE 
   CASE dim_ord(1)=1       
        INDEX ON named TAG T1
   CASE dim_ord(2)=1
        INDEX ON fio TAG T1
ENDCASE 
SET ORDER TO 1

WITH fSupl
     .butPrn.Visible=.F.
     .butRet.Visible=.F.
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH   
IF dim_opt(1)=1
   #DEFINE xlCenter -4108            
   #DEFINE xlLeft -4131  
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
   ON ERROR DO ersup
   WITH excelBook.Sheets(1)
        .Columns(1).ColumnWidth=5
        .Columns(2).ColumnWidth=30
        .Columns(3).ColumnWidth=10
        .Columns(4).ColumnWidth=8
        .Columns(5).ColumnWidth=30
        .Columns(6).ColumnWidth=30             
        .Columns(7).ColumnWidth=30              
        .Columns(8).ColumnWidth=12             
        .Columns(9).ColumnWidth=8                
        .Range(.Cells(2,1),.Cells(2,9)).Select  
        WITH objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='Контрольный список специалистов с высшим медицинским образованием'   
             .Font.Name='Times New Roman'
             .Font.Size=11 
        ENDWITH        
        .Range(.Cells(3,1),.Cells(3,9)).Select  
        WITH objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value=ALLTRIM(boss.office)+' на '+DTOC(dateSpis)         
             .Font.Name='Times New Roman'
             .Font.Size=11 
        ENDWITH        
        numberrow=4
        rowtop=numberrow
        .Cells(numberrow,1).Value='№'
        .Cells(numberrow,2).Value='Фамилия Имя Отчество'
        .Cells(numberrow,3).Value='Пол'
        .Cells(numberrow,4).Value='Год рождения'
        .Cells(numberrow,5).Value='Когда и где окончил учебное заведение'
        .Cells(numberrow,6).Value='Занимаемая должность'
        .Cells(numberrow,7).Value='Усовершенствование специализации'
        .Cells(numberrow,8).Value='Категория'
        .Cells(numberrow,9).Value='Срок окончания контракта'
        .Range(.Cells(numberrow,1),.Cells(numberrow,9)).Select  
        WITH objExcel.Selection         
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1          
             .WrapText=.T.
        ENDWITH   
        numberrow=numberrow+1              
        SELECT curPeopList
        DO storezeropercent
        nppcx=1    
        GO TOP
        DO WHILE !EOF()           
           .cells(numberrow,1).Value=nppcx               
           .cells(numberrow,2).Value=ALLTRIM(fio)      
           .cells(numberrow,3).Value=IIF(sex=1,'муж.','жен.')
           .cells(numberrow,4).Value=STR(YEAR(age),4)
           .cells(numberrow,5).Value=ALLTRIM(school)+IIF(!EMPTY(dkval),' '+DTOC(dkval),'')
           .cells(numberrow,6).Value=ALLTRIM(named)
           .cells(numberrow,7).Value=ALLTRIM(kurs) 
           .cells(numberrow,8).Value=IIF(SEEK(kval,'sprkval',1),ALLTRIM(sprkval.name)+IIF(!EMPTY(dkval),' '+DTOC(dkval),''),'')
           .cells(numberrow,9).Value=IIF(!EMPTY(enddog),DTOC(enddog),'')
           
           numberrow=numberrow+1     
           nppcx=nppcx+1
           DO fillpercent WITH 'fSupl'
           SKIP           
        ENDDO          
       .Range(.Cells(rowtop,1),.Cells(numberRow-1,9)).Select
       WITH objExcel.Selection
*  *       .Borders(xlEdgeLeft).Weight=xlThin
            .VerticalAlignment=1
            .Borders(xlEdgeTop).Weight=xlThin            
            .Borders(xlEdgeBottom).Weight=xlThin
            .Borders(xlEdgeRight).Weight=xlThin
            .Borders(xlInsideVertical).Weight=xlThin
            .Borders(xlInsideHorizontal).Weight=xlThin
            .WrapText=.T.
            .Font.Name='Times New Roman'
            .Font.Size=11           
       ENDWITH       
       .Cells(2,1).Select
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
   ON ERROR 
   objExcel.Visible=.T.    
ELSE 
   #DEFINE wdBorderTop -1           &&верхняя граница ячейки таблицы
   #DEFINE wdBorderLeft -2          &&левая граница ячейки таблицы
   #DEFINE wdBorderBottom -3        &&нижняя граница ячейки таблицы
   #DEFINE wdBorderRight -4         &&правая граница ячейки таблицы
   #DEFINE wdBorderHorizontal -5    &&горизонтальные линии таблицы
   #DEFINE wdBorderVertical -6      &&горизонтальные линии таблицы
   #DEFINE wdLineStyleSingle 1      && стиль линии границы ячейки (в данно случае обычная)
   #DEFINE wdLineStyleNone 0        && линия отсутствует
   #DEFINE wdAlignParagraphRight 2

   objWord=CREATEOBJECT('WORD.APPLICATION')
   #DEFINE cr CHR(13)
   nameDoc=objWord.Documents.Add()  
   nameDoc.ActiveWindow.View.ShowAll=0        
   objWord.Selection.pageSetup.Orientation=1
   objWord.Selection.pageSetup.LeftMargin=30
   objWord.Selection.pageSetup.RightMargin=20
   objWord.Selection.pageSetup.TopMargin=10
   objWord.Selection.pageSetup.BottomMargin=10
   docRef=GETOBJECT('','word.basic')
   WITH docRef
        .Insert(cr)
        .Font('Times New Roman',12)
        .CenterPara 
        .Insert('Контрольный список специалистов с высшим медицинским образованием')
        .Insert(cr)
        .Font('Times New Roman',12)
        .CenterPara
        .Insert('УЗ "Брестская центральная городская больница" на '+DTOC(dateSpis))
        .Insert(cr)
        .clearFormatting
        .Font('Times New Roman',12)
        .centerPara   
        nameDoc.Tables.add(objWord.Selection.range,1,9)
        ordTable1=nameDoc.Tables(1) 
        WITH ordTable1
             .Columns(1).Width=30
             .cell(1,1).Range.Select                   
             *docRef.CenterPara  && выравниеваем слева
             docRef.Font('Times New Roman',11)
             .cell(1,1).Range.Text='№ п.п'            
             .Columns(2).Width=150 
             .cell(1,2).Range.Select     
             docRef.Font('Times New Roman',11)  
             .cell(1,2).Range.Text='Фамилия Имя Отчество'             
             .Columns(3).Width=40
             .cell(1,3).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,3).Range.Text='Пол'           
             .Columns(4).Width=40
             .cell(1,4).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,4).Range.Text='Год рождения'            
             .Columns(5).Width=140
             .cell(1,5).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,5).Range.Text='Когда и где окончил учебное заведение'            
             .Columns(6).Width=100
             .cell(1,6).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,6).Range.Text='Занимаемая должность'           
             .Columns(7).Width=140
             .cell(1,7).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,7).Range.Text='Усовершенствование специализации'         
             .Columns(8).Width=60
             .cell(1,8).Range.Select     
             docRef.Font('Times New Roman',11)          
             .cell(1,8).Range.Text='Категория'
             .Columns(9).Width=60
             .cell(1,9).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,9).Range.Text='Срок окончания контракта'           
             .Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderVertical).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderRight).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderLeft).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderBottom).LineStyle=wdLineStyleSingle                    
             docRef.CloseParaBelow  &&Удаляем лишний интервал после абзаца            
             docRef.LineDown 
             rowcx=2
             nppcx=1          
             DO storezeropercent        
             SCAN ALL 
                  .Rows.Add  
                  .cell(rowcx,1).Range.Text=nppcx               
                  docRef.CenterPara  && выравниеваем слева
                  .cell(rowcx,2).Range.Select                
                  docRef.LeftPara  && выравниеваем слева  
                  .cell(rowcx,2).Range.Text=ALLTRIM(fio) 
                  .cell(rowcx,3).Range.Text=IIF(sex=1,'муж.','жен.')
                  .cell(rowcx,4).Range.Text=STR(YEAR(age),4)
                  .cell(rowcx,5).Range.Text=ALLTRIM(school)+IIF(!EMPTY(dkval),' '+DTOC(dkval),'')
                  .cell(rowcx,5).Range.Select                
                   docRef.LeftPara  && выравниеваем слева  
                  .cell(rowcx,6).Range.Text=named
                  .cell(rowcx,6).Range.Select                               
                   docRef.LeftPara  && выравниеваем слева 
                  .cell(rowcx,7).Range.Text=ALLTRIM(kurs) 
                  .cell(rowcx,7).Range.Select                
                  docRef.LeftPara  && выравниеваем слева                 
                  .cell(rowcx,8).Range.Text=IIF(SEEK(kval,'sprkval',1),ALLTRIM(sprkval.name)+IIF(!EMPTY(dkval),' '+DTOC(dkval),''),'')
                  .cell(rowcx,9).Range.Text=IIF(!EMPTY(enddog),DTOC(enddog),'')
                  rowcx=rowcx+1
                  nppcx=nppcx+1
                  DO fillpercent WITH 'fSupl'
             ENDSCAN              
        ENDWITH   
   ENDWITH  
   objWord.Visible=.T.       
 ENDIF   
 WITH fSupl
      .butPrn.Visible=.T.
      .butRet.Visible=.T.
      .Shape11.Visible=.F.
      .Shape12.Visible=.F.
      .lab25.Visible=.F.
ENDWITH     

****************************************************************************************************************
*                                Печать в разрезе должностей
****************************************************************************************************************
PROCEDURE prnDolSpis
IF !USED('boss')
   USE boss IN 0
ENDIF 
IF USED('curJobspis')
   SELECT curJobSpis
   USE
ENDIF 
IF USED('curPeopList')
   SELECT curPeopList
   USE
ENDIF 
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF 
SELECT * FROM sprdolj INTO CURSOR curprn READWRITE
ALTER TABLE curPrn ADD COLUMN nTot N(3)
ALTER TABLE curPrn ADD COLUMN nvtot N(3)
ALTER TABLE curPrn ADD COLUMN n1tot N(3)
ALTER TABLE curPrn ADD COLUMN n2tot N(3)
ALTER TABLE curPrn ADD COLUMN n3tot N(3)
ALTER TABLE curPrn ADD COLUMN nkat N(1)
REPLACE nkat WITH kat ALL
REPLACE nkat WITH 1 FOR lSpis
INDEX ON STR(nkat,1)+name TAG T1

SELECT * FROM datjob INTO CURSOR curJobSpis READWRITE
SELECT curJobSpis
APPEND FROM datjobout
DELETE FOR tr#1
DELETE FOR dateBeg>dateSpis
DELETE FOR !EMPTY(dateout).AND.dateout<dateSpis
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) ALL
REPLACE kat WITH 1 FOR SEEK(kd,'sprdolj',1).AND.sprdolj.lspis
DELETE FOR !INLIST(kat,1,2,5,7)
INDEX ON kd TAG T1
*INDEX ON kodpeop TAG T2
INDEX ON nidpeop TAG T2
SET ORDER TO 1

SELECT * FROM people INTO CURSOR curPeopList READWRITE
SELECT curPeopList
APPEND FROM peopout
DELETE FOR !SEEK(nid,'curJobspis',2)

SELECT curJobSpis
SET ORDER TO 2
SELECT curPeopList
SCAN ALL
     SELECT curJobSpis
     SEEK curPeopList.nid
     IF !EOF()
        SKIP     
        LOCATE REST FOR nidpeop=curPeopList.nid
        IF FOUND()        
           DELETE
        ENDIF
     ENDIF
     SELECT curPeopList   
ENDSCAN

SELECT curJobSpis
SET ORDER TO 1
SELECT curprn
SCAN ALL
     SELECT curJobSpis
     SEEK curPrn.kod
     STORE 0 TO nTot_cx,nvToT_cx,n1tot_cx,n2Tot_cx,n3Tot_cx    
     SCAN WHILE kd=curPrn.kod
          nTot_cx=nTot_cx+1
          nvTot_cx=IIF(kv=1,nvTot_cx+1,nvTot_cx)
          n1Tot_cx=IIF(kv=2,n1Tot_cx+1,n1Tot_cx)
          n2Tot_cx=IIF(kv=3,n2Tot_cx+1,n2Tot_cx)
          n3Tot_cx=IIF(!INLIST(kv,1,2,3),n3Tot_cx+1,n3Tot_cx)
     ENDSCAN     
     SELECT curprn
     REPLACE nTot WITH nTot_cx,nvTot WITH nvTot_cx,n1Tot WITH n1Tot_cx,n2Tot WITH n2Tot_cx,n3Tot WITH n3Tot_cx
ENDSCAN
DELETE FOR ntot=0
WITH fSupl
     .butPrn.Visible=.F.
     .butRet.Visible=.F.
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH   
IF dim_opt(1)=1
   #DEFINE xlCenter -4108            
   #DEFINE xlLeft -4131  
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
        .Columns(1).ColumnWidth=50
        .Columns(2).ColumnWidth=10
        .Columns(3).ColumnWidth=10
        .Columns(4).ColumnWidth=10
        .Columns(5).ColumnWidth=10
        .Columns(6).ColumnWidth=10              
        .Range(.Cells(2,1),.Cells(2,6)).Select  
        WITH objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='Контрольный список должностей'   
             .Font.Name='Times New Roman'
             .Font.Size=11 
        ENDWITH        
        .Range(.Cells(3,1),.Cells(3,6)).Select  
        WITH objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value=ALLTRIM(boss.office)+' на '+DTOC(dateSpis)         
             .Font.Name='Times New Roman'
             .Font.Size=11 
        ENDWITH        
        numberrow=4
        rowtop=numberrow
        .Cells(numberrow,1).Value='наименование должности'
        .Cells(numberrow,2).Value='всего'
        .Cells(numberrow,3).Value='выс.к.'
        .Cells(numberrow,4).Value='1 кат.'
        .Cells(numberrow,5).Value='2 кат.'
        .Cells(numberrow,6).Value='без кат.'
        .Range(.Cells(numberrow,1),.Cells(numberrow,6)).Select  
        WITH objExcel.Selection         
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1          
        ENDWITH   
        numberrow=numberrow+1              
        SELECT curPrn
        DO storezeropercent  
        kat_cx=0
        GO TOP
        DO WHILE !EOF()
           IF kat_cx#nkat
              kat_cx=nkat
              STORE 0 TO nTot_cx,nvToT_cx,n1tot_cx,n2Tot_cx,n3Tot_cx         
              .cells(numberrow,1).Value=IIF(SEEK(kat_cx,'sprkat',1),sprkat.name,'')  
              .cells(numberrow,1).Font.Bold=.T.        
              numberrow=numberrow+1      
           ENDIF 
           .cells(numberrow,1).Value=name              
           .cells(numberrow,2).Value=IIF(nTot#0,nTot,'')       
           .cells(numberrow,3).Value=IIF(nvTot#0,nvTot,'')       
           .cells(numberrow,4).Value=IIF(n1Tot#0,n1Tot,'')       
           .cells(numberrow,5).Value=IIF(n2Tot#0,n2Tot,'')       
           .cells(numberrow,6).Value=IIF(n3Tot#0,n3Tot,'')                                               
          
           numberrow=numberrow+1      
           nTot_cx=nTot_cx+nTot
           nvTot_cx=nvTot_cx+nvTot
           n1Tot_cx=n1Tot_cx+n1Tot
           n2Tot_cx=n2Tot_cx+n2Tot
           n3Tot_cx=n3Tot_cx+n3Tot                                   
           DO fillpercent WITH 'fSupl'
           SKIP
           IF kat_cx#nkat
              .cells(numberrow,1).Value='итого'            
              .cells(numberrow,2).Value=IIF(nTot_cx#0,nTot_cx,'')       
              .cells(numberrow,3).Value=IIF(nvTot_cx#0,nvTot_cx,'')       
              .cells(numberrow,4).Value=IIF(n1Tot_cx#0,n1Tot_cx,'')       
              .cells(numberrow,5).Value=IIF(n2Tot_cx#0,n2Tot_cx,'')       
              .cells(numberrow,6).Value=IIF(n3Tot_cx#0,n3Tot_cx,'')            
            
              .cells(numberrow,1).Font.Bold=.T.
              .cells(numberrow,2).Font.Bold=.T.
              .cells(numberrow,3).Font.Bold=.T.
              .cells(numberrow,4).Font.Bold=.T.
              .cells(numberrow,5).Font.Bold=.T.
              .cells(numberrow,6).Font.Bold=.T.
              *.Font.Size=11
              numberrow=numberrow+1      
           ENDIF
        ENDDO          
       .Range(.Cells(rowtop,1),.Cells(numberRow-1,6)).Select
       WITH objExcel.Selection
*  *       .Borders(xlEdgeLeft).Weight=xlThin
            .Borders(xlEdgeTop).Weight=xlThin            
            .Borders(xlEdgeBottom).Weight=xlThin
            .Borders(xlEdgeRight).Weight=xlThin
            .Borders(xlInsideVertical).Weight=xlThin
            .Borders(xlInsideHorizontal).Weight=xlThin
            .WrapText=.T.
            .Font.Name='Times New Roman'
            .Font.Size=11           
       ENDWITH       
       .Cells(2,1).Select
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
ELSE 
   #DEFINE wdBorderTop -1           &&верхняя граница ячейки таблицы
   #DEFINE wdBorderLeft -2          &&левая граница ячейки таблицы
   #DEFINE wdBorderBottom -3        &&нижняя граница ячейки таблицы
   #DEFINE wdBorderRight -4         &&правая граница ячейки таблицы
   #DEFINE wdBorderHorizontal -5    &&горизонтальные линии таблицы
   #DEFINE wdBorderVertical -6      &&горизонтальные линии таблицы
   #DEFINE wdLineStyleSingle 1      && стиль линии границы ячейки (в данно случае обычная)
   #DEFINE wdLineStyleNone 0        && линия отсутствует 
   #DEFINE wdAlignParagraphRight 2
   #DEFINE wdAlignParagraphJustify 2

   objWord=CREATEOBJECT('WORD.APPLICATION')
   #DEFINE cr CHR(13)
   nameDoc=objWord.Documents.Add()  
   nameDoc.ActiveWindow.View.ShowAll=0        
   objWord.Selection.pageSetup.Orientation=0
   objWord.Selection.pageSetup.LeftMargin=30
   objWord.Selection.pageSetup.RightMargin=20
   objWord.Selection.pageSetup.TopMargin=15
   objWord.Selection.pageSetup.BottomMargin=15
   docRef=GETOBJECT('','word.basic')

   WITH docRef
        .Insert(cr)
        .Font('Times New Roman',12)
        .CenterPara 
        .Insert('Контрольный список должностей')
        .Insert(cr)
        .Font('Times New Roman',12)
        .CenterPara
        .Insert('УЗ "Брестская центральная городская больница" на '+DTOC(dateSpis))
        .Insert(cr)
        .clearFormatting
        .Font('Times New Roman',12)
        .centerPara   
        nameDoc.Tables.add(objWord.Selection.range,1,6)
        ordTable1=nameDoc.Tables(1) 
        WITH ordTable1
             .Columns(1).Width=250
             .cell(1,1).Range.Select                   
             *docRef.CenterPara  && выравниеваем слева
             docRef.Font('Times New Roman',11)
             .cell(1,1).Range.Text='наименование должности'
          
             .Columns(2).Width=50
             .cell(1,2).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,2).Range.Text='всего' 
             
             .Columns(3).Width=50 
             .cell(1,3).Range.Select     
             docRef.Font('Times New Roman',11)  
             .cell(1,3).Range.Text='выс.к.'
          
             .Columns(4).Width=50
             .cell(1,4).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,4).Range.Text='1 кат.'
          
             .Columns(5).Width=50
             .cell(1,5).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,5).Range.Text='2 кат.' 
                              
             .Columns(6).Width=50
             .cell(1,6).Range.Select     
             docRef.Font('Times New Roman',11)
             .cell(1,6).Range.Text='без кат.' 
                   
             .Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderVertical).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderRight).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderLeft).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
             .Borders(wdBorderBottom).LineStyle=wdLineStyleSingle    
                
             docRef.CloseParaBelow  &&Удаляем лишний интервал после абзаца            
             docRef.LineDown 
             rowcx=2     
          
             SELECT curPrn
             DO storezeropercent  
             kat_cx=0
             GO TOP
             DO WHILE !EOF()
                IF kat_cx#nkat
                   kat_cx=nkat
                   STORE 0 TO nTot_cx,nvToT_cx,n1tot_cx,n2Tot_cx,n3Tot_cx    
                   .Rows.Add 
                   .cell(rowcx,1).Range.Text=IIF(SEEK(kat_cx,'sprkat',1),sprkat.name,'')  
                   .cell(rowcx,1).Range.Select    
                   .cell(rowcx,1).Range.Font.Bold=.T.
                   .cell(rowcx,1).Range.Font.Size=13
                   docRef.LeftPara  && выравниеваем слева  
                   rowcx=rowcx+1  
                ENDIF 
                .Rows.Add  
                .cell(rowcx,1).Range.Text=name              
                .cell(rowcx,1).Range.Select  
                .cell(rowcx,1).Range.Font.Bold=.F.
                .cell(rowcx,1).Range.Font.Size=11
                docRef.LeftPara  && выравниеваем слева  
                .cell(rowcx,2).Range.Text=nTot 
                .cell(rowcx,3).Range.Text=IIF(nvTot#0,nvTot,'')
                .cell(rowcx,4).Range.Text=IIF(n1Tot#0,n1Tot,'')
                .cell(rowcx,5).Range.Text=IIF(n2Tot#0,n2Tot,'')
                .cell(rowcx,6).Range.Text=IIF(n3Tot#0,n3Tot,'')  
  
                nTot_cx=nTot_cx+nTot
                nvTot_cx=nvTot_cx+nvTot
                n1Tot_cx=n1Tot_cx+n1Tot
                n2Tot_cx=n2Tot_cx+n2Tot
                n3Tot_cx=n3Tot_cx+n3Tot                           
                rowcx=rowcx+1          
                DO fillpercent WITH 'fSupl'
                SKIP
                IF kat_cx#nkat
                   .Rows.Add  
                   .cell(rowcx,1).Range.Text='итого'            
                   .cell(rowcx,1).Range.Select
                   .cell(rowcx,1).Range.Font.Bold=.T.                
                   docRef.LeftPara  && выравниеваем слева  
                   .cell(rowcx,2).Range.Text=nTot_cx
                   .cell(rowcx,3).Range.Text=IIF(nvTot_cx#0,nvTot_cx,'')
                   .cell(rowcx,4).Range.Text=IIF(n1Tot_cx#0,n1Tot_cx,'')
                   .cell(rowcx,5).Range.Text=IIF(n2Tot_cx#0,n2Tot_cx,'')
                   .cell(rowcx,6).Range.Text=IIF(n3Tot_cx#0,n3Tot_cx,'')
                   rowcx=rowcx+1    
                ENDIF
             ENDDO              
        ENDWITH   
   ENDWITH  
   objWord.Visible=.T.       
ENDIF    
WITH fSupl
     .butPrn.Visible=.T.
     .butRet.Visible=.T.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH     
