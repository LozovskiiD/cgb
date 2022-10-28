***************************************************************************************************************************************************************
DIMENSION dim_ord(2),dim_tr(3)
dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим

dim_tr(1)=1  &&все
dim_tr(2)=0  &&основной
dim_tr(3)=0  &&вн.совм.
DO procDimFlt

STORE CTOD('  .  .    ') TO date_Beg,date_End
topSay=''
lOut=.F.
lIn=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='—писок при€тых, уволенных'  
     DO procObjFlt       
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8    
     
     DO adLabMy WITH 'fSupl',1,'период с ',.Shape2.Top+20,.Shape2.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxBeg',.Shape2.Top+20,.Shape1.Left,RetTxtWidth('99/99/99999'),dHeight,'date_Beg',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3
     
     DO adLabMy WITH 'fSupl',2,' по ',.lab1.Top,.Shape2.Left,.Shape2.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxEnd',.boxBeg.Top,.Shape2.Left,.boxBeg.Width,dHeight,'date_End',.F.,.T.,0
     .lab1.Left=.Shape2.Left+(.Shape2.Width-.lab1.Width-.boxBeg.Width-.lab2.Width-.boxEnd.Width-30)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+10
     .lab2.Left=.boxBeg.Left+.boxBeg.Width+10
     .boxEnd.Left=.lab2.Left+.lab2.Width+10
     
     DO addOptionButton WITH 'fSupl',21,'все',.boxBeg.Top+.boxBeg.Height+10,.Shape2.Left+20,'dim_tr(1)',0,"DO procValOption WITH 'fSupl','dim_tr',1",.T. 
     DO addOptionButton WITH 'fSupl',22,'основн.',.Option21.Top,.Option21.Left+.Option21.Width+10,'dim_tr(2)',0,"DO procValOption WITH 'fSupl','dim_tr',2",.T. 
     DO addOptionButton WITH 'fSupl',23,'вн. совм.',.Option21.Top,.Option21.Left+.Option21.Width+10,'dim_tr(3)',0,"DO procValOption WITH 'fSupl','dim_tr',3",.T. 
     .Option21.Left=.Shape2.Left+(.Shape2.Width-.Option21.Width-.Option22.Width-.Option23.Width-20)/2
     .Option22.Left=.Option21.Left+.Option21.Width+10 
     .Option23.Left=.Option22.Left+.Option22.Width+10      
     
     
     DO addOptionButton WITH 'fSupl',1,'по алфавиту',.option21.Top+.option21.Height+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'по дате',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20 
     .Shape2.Height=.Option21.height*2+.boxBeg.Height+40         
     DO adSetupPrnToForm WITH .Shape2.Left,.Shape2.Top+.Shape2.Height+10,.Shape2.Width,.F.,.T.
     DO adButtonPrnToForm WITH 'DO prnSpisInMove WITH 1','DO prnSpisInMove WITH 2','fsupl.Release',.T.
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH      
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width
     .Width=.Shape1.Width+20
     .Height=.butPrn.Top+.butPrn.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************
PROCEDURE prnSpisInMove
PARAMETERS par1
IF USED('curInJob')
   SELECT curInJob
   USE
ENDIF
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
CREATE CURSOR curPrn (npp N(3),kodpeop N(5),fio C(60),kpold N(3),kdold N(3),kseOld N(6,2), kpnew N(3),kdnew N(3),ksenew N(6,2),pold C(200),pNew C(200),dMove D,tr N(1),cOrder C(20))
topday='—писок сотрудников, переведенных за период с '+DTOC(date_beg)+' по '+DTOC(date_end)

SELECT * FROM datjob WHERE INLIST(tr,1,3).AND.!EMPTY(dateout) INTO CURSOR curoutJob READWRITE
DELETE FOR dateout<date_beg                                      &&удал€ем записи с окончанием работы меньше начальной даты
DELETE FOR dateout>date_end                                      &&удал€ем записи с окончанием работы больше конечной даты
DELETE FOR SEEK(nidpeop,'people',4).AND.people.date_out=curoutJob.dateout             &&удал€ем записи где дата увольнени€ совпадает с датой увольнени€ в people
INDEX ON DTOS(dateout)+STR(nidpeop,5) TAG T1


SELECT * FROM datjob WHERE INLIST(tr,1,3) INTO CURSOR curinJob READWRITE
DELETE FOR !EMPTY(datebeg).AND.datebeg<date_beg                                      &&удал€ем записи с началом работы меньше начальной даты
DELETE FOR !EMPTY(datebeg).AND.datebeg>date_end                                      &&удал€ем записи с началом работы больше конечной даты
DELETE FOR SEEK(nidpeop,'people',4).AND.people.date_in=curinjob.datebeg              &&удал€ем записи где дата "перевода" совпадает с датой перевода
REPLACE fio WITH IIF(SEEK(nidpeop,'people',4),people.fio,'') ALL
SELECT curinjob
SCAN ALL
     SELECT curPrn
     APPEND BLANK
     REPLACE fio WITH curInjob.fio,kodpeop WITH curinjob.kodpeop,dMove WITH curinJob.datebeg,kpNew WITH curinjob.kp,kdnew WITH curinjob.kd,ksenew WITH curinjob.kse,tr WITH curinjob.tr;
             pnew WITH IIF(SEEK(kpnew,'sprpodr',1),ALLTRIM(sprpodr.namework),'')+' '+IIF(SEEK(kdnew,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' '+LTRIM(STR(ksenew,6,2)),cOrder WITH ALLTRIM(curinjob.nordin)+' '+DTOC(curinjob.dordin) 
     SELECT curoutJob
     SEEK DTOS(curinjob.datebeg)+STR(curinjob.nidpeop,5)
     SELECT curprn
     REPLACE kpOld WITH curoutJob.kp,kdold WITH curoutJob.kd,kseold WITH curoutJob.kse,;
             pold WITH IIF(SEEK(kpold,'sprpodr',1),ALLTRIM(sprpodr.namework),'')+' '+IIF(SEEK(kdold,'sprdolj',1),ALLTRIM(sprdolj.name),'')+' '+LTRIM(STR(kseold,6,2))        
     SELECT curinjob
ENDSCAN

SELECT curPrn
DO CASE
   CASE dim_tr(2)=1
        DELETE FOR tr#1
   CASE dim_tr(3)=1
        DELETE FOR tr#3
ENDCASE
DO CASE
   CASE dim_ord(1)=1
        INDEX ON fio TAG T1
   CASE dim_ord(2)=1
        INDEX ON dmove TAG T1
ENDCASE
nppcx=1
SCAN ALL
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'repinmove','список сотрудников',.T.,'spisinmoveToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'repinmove','список сотрудников',.F.,'spisinmoveToExcel'        
ENDCASE
***********************************************************************************************
PROCEDURE spisinmoveToExcel
DO startPrnToExcel WITH 'fSupl'
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)
maxColumn=7
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30  
     .Columns(3).ColumnWidth=80
     .Columns(4).ColumnWidth=80    
     .Columns(5).ColumnWidth=11 
     .Columns(6).ColumnWidth=17  
     .Columns(7).ColumnWidth=17  
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment= -4108
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=11
           .Value=topsay
      ENDWITH   
      .cells(2,1).Value='є'
      .Cells(2,2).Value='‘»ќ'  
      .Cells(2,3).Value='ќткуда' 
      .Cells(2,4).Value=' уда' 
      .Cells(2,5).Value='тип работы' 
      .Cells(2,6).Value='дата перевода' 
      .Cells(2,7).Value='приказ'       
      .Range(.Cells(2,1),.Cells(2,maxColumn)).Select          
      objExcel.Selection.HorizontalAlignment=xlCenter
      SELECT curPrn
      DO storezeropercent
      numberRow=3
       SCAN ALL         
           .cells(numberRow,1).Value=npp
           .cells(numberRow,2).Value=fio
           .cells(numberRow,3).Value=pold
           .cells(numberRow,4).Value=pnew       
           .cells(numberRow,5).Value=IIF(SEEK(tr,'sprtype',1),ALLTRIM(sprtype.name),'')
           .cells(numberRow,6).Value=IIF(!EMPTY(dmove),DTOC(dmove),'')
           .cells(numberRow,7).Value=cOrder
           DO fillpercent WITH 'fSupl'
           numberRow=numberRow+1
      ENDSCAN
      .Range(.Cells(2,1),.Cells(numberRow-1,maxColumn)).Select
      WITH objExcel.Selection
           .Borders(xlEdgeLeft).Weight=xlThin
           .Borders(xlEdgeTop).Weight=xlThin            
           .Borders(xlEdgeBottom).Weight=xlThin
           .Borders(xlEdgeRight).Weight=xlThin
           .Borders(xlInsideVertical).Weight=xlThin
           .Borders(xlInsideHorizontal).Weight=xlThin
           .VerticalAlignment=1
           .Font.Name='Times New Roman'   
           .Font.Size=11
           .WrapText=.T.
      ENDWITH   
      .Cells(1,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'
objExcel.Visible=.T.