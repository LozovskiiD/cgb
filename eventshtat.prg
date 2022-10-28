*******************************************************************************************************************************************************
*                                                          Печать изменений численности
*******************************************************************************************************************************************************
terminal=.T.
varDshtatPrn=varDshtat
varRegPrn=varRegShtat                          
varStavkaPrn=varStavka 
strRegionPrn=varStrRegion
strDate=DTOC(dimshtat(1))
dateBeg=CTOD('  .  .    ')
dateEnd=CTOD('  .  .    ')
STORE '' TO fltPodr,fltDolj,fltPrn,yearSign
yearSign=STR(YEAR(dimshtat(1)),4)
IF USED('curFltPodr')
   SELECT curFltPodr
   USE
ENDIF
SELECT * FROM sprpodr WHERE dShtat=varDPodr INTO CURSOR curFltPodr READWRITE
SELECT curFltPodr
INDEX ON np TAG T1
SET ORDER TO 1
SET FILTER TO reg=varRegPrn

IF USED('curFltDolj')
   SELECT curFltDolj
   USE
ENDIF
SELECT * FROM sprdolj INTO CURSOR curFltDolj READWRITE
SELECT curFltDolj
INDEX ON name TAG T1
SET ORDER TO 1
numIzm=0
SELECT * FROM sprregion WHERE SEEK(STR(sprregion.kod,3),'datshtat',1) INTO CURSOR curPrnRegion ORDER BY namereg READWRITE
SELECT reg,dshtat FROM datshtat  DISTINCT INTO CURSOR curDateRaspPrn
INDEX ON dshtat TAG T1
SET FILTER TO reg=varRegPrn

fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Мероприятия по штатному расписанию'          
     DO addshape WITH 'fSupl',2,20,20,150,400,8 
     DO adtBoxNew WITH 'fSupl','txtBeg',.Shape2.Top+20,.Shape2.Left,RetTxtWidth('99/99/99999'),dHeight,'dateBeg',.F.,.T.,.F.,.F.,'' 
     DO adtBoxNew WITH 'fSupl','txtEnd',.txtBeg.Top,.Shape2.Left,.txtBeg.Width,dHeight,'dateEnd',.F.,.T.,.F.,.F.,'' 
     .txtBeg.Left=.Shape2.Left+(.shape2.Width-.txtBeg.Width*2-20)/2
     .txtEnd.Left=.txtBeg.Left+.txtBeg.Width+20
     .Shape2.Height=.txtBeg.Height+40
     DO adLabMy WITH 'fSupl',13,' Период с - по ',.Shape2.Top-10,.Shape2.Left,100,0,.T.,1
     .lab13.Left=.Shape2.Left+(.Shape2.Width-.lab13.Width)/2     
                
    
     DO adSetupPrnToForm WITH .Shape2.Left,.Shape2.Top+.Shape2.Height+20,.Shape2.Width,.F.,.T.
     *-----------------------------Кнопка печать---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape2.Left+(.Shape2.Width-(RetTxtWidth('wпросмотрw')*3)-30)/2,;
       .Shape91.Top+.Shape91.Height+20,RetTxtWidth('wпросмотрw'),dHeight+5,'Печать','DO procRepEvent WITH .T.'

     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Просмотр','DO procRepEvent WITH .F.'
     
     .SetAll('ForeColor',RGB(0,0,128),'CheckBox')  
     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.cont2.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','fSupl.Release','Возврат'   
     
     DO addShape WITH 'fSupl',11,.Shape2.Left,.cont1.Top,.cont1.Height,.Shape2.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+2,.Shape11.Left,.Shape11.Width,2,.F.,0
     .lab25.Visible=.F.    
             
     
          
    .Width=.Shape2.Width+40
    .Height=.Shape2.Height+.Shape91.Height+.cont1.Height+80      
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************************************************************
PROCEDURE procRepEvent
PARAMETERS parTerm
headEvent='Мероприятия по штптному расписанию за период с '+DTOC(dateBeg)+' по '+DTOC(dateEnd)+' '+ALLTRIM(strRegionPrn )
CREATE CURSOR curEventPrn (kd N(3),namedol C(100),kseBeg N(6,2),vacBeg N(6,2),spisBeg N(6,2),fondBeg N(10,2),kprich N(2),nprich C(50),;
                           kseEnd N(6,2),vacEnd N(6,2),spisEnd N(6,2),fondEnd N(10,2),;
                           kseTot N(6,2),vacTot N(6,2),spisTot N(6,2),fondTot N(10,2))
SELECT curEventPrn
INDEX ON STR(kd,3)+STR(kprich,2) TAG T1
SELECT rasp
SELECT * FROM datShtat WHERE reg=varRegPrn INTO CURSOR curSuplDate ORDER BY dShtat
SELECT curSuplDate
GO BOTTOM 
DO WHILE !BOF()
   varSup1=dShtat
   IF dShtat<=dateBeg   
      EXIT 
   ENDIF 
   SKIP-1
ENDDO
GO BOTTOM 
DO WHILE !BOF()
   varSup2=dShtat
   IF dShtat<=dateEnd   
      EXIT 
   ENDIF 
   SKIP-1
ENDDO
SELECT curSupldate 
USE 

IF USED('curAnalizBeg')
   SELECT curAnalizBeg
   USE
ENDIF

SELECT * FROM rasp WHERE reg=VarRegPrn.AND.dshtat=varSup1 INTO CURSOR curAnalizBeg READWRITE && Курсор для начального штатного
SELECT curAnalizBeg
DELETE FOR dateIn>=dateBeg
DELETE FOR !EMPTY(dateOut).AND.dateOut<dateBeg

*DELETE FOR !EMPTY(dateIn).AND.dateIn<dateBeg


ALTER TABLE curAnalizBeg ADD COLUMN named c(100)
INDEX ON STR(np,3)+STR(nd,3) TAG T1
INDEX ON kod TAG T2
SET ORDER TO 2
*DELETE FOR suplop=1.AND.dateop<=dateBeg
IF USED('curAnalizEnd')
   SELECT curAnalizEnd
   USE
ENDIF
SELECT * FROM rasp WHERE reg=VarRegPrn.AND.dshtat=varSup2 INTO CURSOR curAnalizEnd READWRITE  && Курсор для конечного штатного

*DELETE FOR suplop=1.AND.dateop<=dateBeg
*DELETE FOR suplop=2.AND.dateOp>dateEnd

SELECT curAnalizEnd
DELETE FOR dateIn>dateEnd
DELETE FOR !EMPTY(dateOut).AND.dateOut<dateBeg                         &&удаляем записи 'исключить' до даты dateLast
DELETE FOR !EMPTY(dateOut).AND.dateOut>=dateBeg.AND.dateOut<=dateEnd   &&удаляем записи 'исключить' за период

ALTER TABLE curAnalizEnd ADD COLUMN named c(100)
INDEX ON STR(np,3)+STR(nd,3) TAG T1
INDEX ON kod TAG T2
SET ORDER TO 2
SELECT curAnalizBeg    
SCAN ALL
     SELECT curAnalizEnd
     SEEK curAnalizBeg.kod
     DO CASE
        CASE !FOUND()  &&вообще не найдено
             SELECT curEventPrn
             SEEK STR(curAnalizBeg.kd,3)+STR(curAnalizBeg.kprichout,2)
             IF !FOUND()
                APPEND BLANK 
                REPLACE kd WITH curAnalizBeg.kd,kPrich WITH curAnalizBeg.KprichOut
                
             ENDIF
             
             REPLACE ksebeg WITH kseBeg+curAnalizBeg.kse,vacBeg WITH IIF(curAnalizBeg.vac,vacBeg+curAnalizbeg.kse,vacBeg),fondBeg WITH fondBeg+curAnalizBeg.msf
            
        CASE FOUND().AND.curAnalizBeg.kse#curAnalizEnd.Kse     
             SELECT curEventPrn
             SEEK STR(curAnalizBeg.kd,3)+STR(curAnalizBeg.kprichout,2)
             IF !FOUND()
                APPEND BLANK 
                REPLACE kd WITH curAnalizBeg.kd,kPrich WITH curAnalizBeg.KprichOut
             ENDIF
             REPLACE ksebeg WITH kseBeg+curAnalizBeg.kse,vacBeg WITH IIF(curAnalizBeg.vac,vacBeg+curAnalizbeg.kse,vacBeg),kseEnd WITH curAnalizEnd.kse,fondBeg WITH fondBeg+curAnalizBeg.msf
             
     ENDCASE
     SELECT curAnalizBeg
ENDSCAN
SELECT curAnalizEnd                           
SCAN ALL
     IF !SEEK(kod,'curAnalizBeg',2)
        SELECT curEventPrn
        SEEK STR(curAnalizEnd.kd,3)+STR(curAnalizEnd.kprichin,2)
        IF !FOUND()       
           APPEND BLANK 
           REPLACE kd WITH curAnalizEnd.kd,kPrich WITH curAnalizEnd.KprichIn
        ENDIF
         
        REPLACE kseEnd WITH kseEnd+curAnalizEnd.kse,vacEnd WITH IIF(curAnalizEnd.vac,vacEnd+curAnalizEnd.kse,vacEnd),fondEnd WITH fondEnd+curAnalizEnd.msf
     ENDIF 
     SELECT curAnalizEnd
ENDSCAN
SELECT curEventPrn
REPLACE namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''), nPrich WITH IIF(SEEK(kprich,'sprreason',1),sprreason.namereason,''),kseTot WITH kseBeg-kseEnd,fondTot WITH fondBeg-fondEnd ALL
REPLACE vacTot WITH vacBeg-vacEnd,spisBeg WITH kseBeg-vacBeg,spisEnd WITH kseEnd-vacEnd,spisTot WITH ksetot-vacTot ALL
SUM kseBeg,kseEnd,kseTot,vacBeg,vacEnd,vacTot,fondBeg,fondEnd,fondTot TO kseBeg_cx,kseEnd_cx,kseTot_cx,vacBeg_cx,vacEnd_cx,vacTot_cx,fondBeg_cx,fondEnd_cx,fondTot_cx
APPEND BLANK
REPLACE kd WITH 999,namedol WITH 'всего',kseBeg WITH kseBeg_cx,kseEnd WITH kseEnd_cx,kseTot WITH kseTot_cx,vacBeg WITH vacBeg_cx,vacEnd WITH vacEnd_cx,vacTot WITH vacTot_cx,;
        fondBeg WITH fondBeg_cx,fondEnd WITH fondEnd_cx,fondTot WITH fondTot_cx,spisBeg WITH kseBeg-vacBeg,spisEnd WITH kseEnd-vacEnd,spisTot WITH ksetot-vacTot 
GO TOP
DO CASE
   CASE parTerm=.T.
        DO procForPrintAndPreview WITH 'repevent','мероприятия по штатному',.T.,'eventToExcel' 
   OTHERWISE
        DO procForPrintAndPreview WITH 'repevent','мероприятия по штатному'
ENDCASE
******************************************************************************************
PROCEDURE eventToExcel
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F.
     .shape11.Visible=.T.     
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape12.Width=1
ENDWITH  
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
STORE 0 TO max_rec,one_pers,pers_ch,page_Hcx

WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .PageSetup.LeftMargin=0
     .PageSetup.RightMargin=0
     .PageSetup.TopMargin=0
     .PageSetup.BottomMargin=0           
     rowcx=1  
     .Columns(1).ColumnWidth=25
     .Columns(2).ColumnWidth=10 
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
     .Columns(13).ColumnWidth=10 
     .Columns(14).ColumnWidth=20               
          
     .Range(.Cells(1,1),.Cells(1,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=HeadEvent        
     ENDWITH  
     rowcx=rowcx+1
     .Range(.Cells(2,2),.Cells(2,5)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='На начало периода'         
     ENDWITH 
     
     .Range(.Cells(2,6),.Cells(2,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='На конец периода'        
     ENDWITH  
     
     .Range(.Cells(2,10),.Cells(2,13)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='экономический эффект'        
     ENDWITH  
     
     .cells(3,2).Value='штатная численность, ед.'
     .cells(3,3).Value='в т.ч. вакансии, ед.'
     .cells(3,4).Value='средне.спис.численность, ед. '
     .cells(3,5).Value='месячный фонд'
     
     .cells(3,6).Value='штатная численность, ед.'
     .cells(3,7).Value='в т.ч. вакансии, ед.'
     .cells(3,8).Value='средне.спис.численность, ед. '
     .cells(3,9).Value='месячный фонд'
     
     .cells(3,10).Value='штатная численность, ед.'
     .cells(3,11).Value='в т.ч. вакансии, ед.'
     .cells(3,12).Value='средне.спис.численность, ед. '
     .cells(3,13).Value='месячный фонд'  
         
     .Range(.Cells(3,2),.Cells(3,13)).Select      
     objExcel.Selection.WrapText=.T.
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     
     .Range(.Cells(2,1),.Cells(3,1)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='должность, профессия'         
     ENDWITH   
     
    .Range(.Cells(2,14),.Cells(3,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='причина'        
     ENDWITH            
     rowcx=4    
     SELECT curEventPrn   
     COUNT TO max_rec
  
     SCAN ALL             
          .cells(rowcx,1).Value=ALLTRIM(namedol) 
          .cells(rowcx,1).WrapText=.T.        
          .cells(rowcx,2).Value=IIF(kseBeg#0,kseBeg,'')       
          .cells(rowcx,3).Value=IIF(vacBeg#0,vacBeg,'')
          .cells(rowcx,4).Value=IIF(spisBeg#0,spisBeg,'')
          .cells(rowcx,5).Value=IIF(fondBeg#0,fondBeg,'')
          .cells(rowcx,6).Value=IIF(kseEnd#0,kseEnd,'')       
          .cells(rowcx,7).Value=IIF(vacEnd#0,vacEnd,'')
          .cells(rowcx,8).Value=IIF(spisEnd#0,spisEnd,'')
          .cells(rowcx,9).Value=IIF(fondEnd#0,fondEnd,'')          
          .cells(rowcx,10).Value=IIF(kseTot#0,kseTot,'')       
          .cells(rowcx,11).Value=IIF(vacTot#0,vacTot,'')
          .cells(rowcx,12).Value=IIF(spisTot#0,spisTot,'')
          .cells(rowcx,13).Value=IIF(fondTot#0,fondTot,'')
          .cells(rowcx,14).Value=ALLTRIM(nprich)
          .cells(rowcx,14).WrapText=.T. 
         
 
            one_pers=one_pers+1
            pers_ch=one_pers/max_rec*100
            fSupl.shape12.Visible=.T.
            fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
            fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch           
           rowcx=rowcx+1        
                
     ENDSCAN   
         
     .Range(.Cells(1,1),.Cells(rowcx-1,14)).Select
     objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
     objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
     objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
     objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
     objExcel.Selection.VerticalAlignment=1
     objExcel.Selection.Font.Name='Times New Roman'   
     objExcel.Selection.Font.Size=9
     

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
WITH fSupl
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.
     .shape11.Visible=.F.
     .shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH   
