fSupl=CREATEOBJECT('FORMSUPL')
monthrep=MONTH(DATE())
yearrep=YEAR(DATE())
WITH fSupl
     .Icon='kone.ico'  
     .Caption='Отчет о численности работающих'
     .Width=400
     .Height=400
     DO addShape WITH 'fSupl',1,20,20,10,10,8
     DO adtBoxAsCont WITH 'fSupl','contMonth',.Shape1.Left+30,.Shape1.Top+20,RetTxtWidth('WWWсентябрьWW'),dHeight,'месяц',2,1
     DO addComboMy WITH 'fSupl',1,.contMonth.Left,.contMonth.Top+.contMonth.Height-1,dheight,.contMonth.Width,.T.,'monthRep','dim_month',5,.F.,.F.,.F.,.T.  
     .comboBox1.DisplayCount=12
     DO adtBoxAsCont WITH 'fSupl','contYear',.contMonth.Left+.contMonth.Width-1,.contMonth.Top,.contMonth.Width,dHeight,'год',2,1
     DO adTboxNew WITH 'fSupl','tBoxYear',.comboBox1.Top,.contYear.Left,.contYear.Width,dHeight,'yearRep','Z',.T.,0 
     .Shape1.Width=.contMonth.Width*2+60
     .Shape1.Height=.contMonth.Height*2+40
     DO addButtonOne WITH 'fSupl','butPrn',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wприступитьw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,'приступить','','DO createemployees',39,RetTxtWidth('wприступитьw'),'формирование отчета' 
     DO addButtonOne WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+10,.butPrn.Top,'возврат','','fSupl.Release',39,.butPrn.Width,'возврат' 
     
     DO addShape WITH 'fSupl',11,.Shape1.Left,.butPrn.Top,.butPrn.Height,.Shape1.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape1.Left,.Shape1.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.  
     
     .Height=.Shape1.Height+.butPrn.Height+60
     .Width=.Shape1.Width+40
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************************************
PROCEDURE createemployees
dBeg=CTOD('01.'+STR(monthrep,2)+'.'+STR(yearrep,4))
dEnd=IIF(INLIST(monthrep,1,3,5,7,8,10,12),31,IIF(monthrep=2,28,30))
dEnd=CTOD(STR(dEnd,2)+'.'+STR(monthrep,2)+'.'+STR(yearrep,4))
beglist=dBeg
endlist=dEnd
dayMonth=IIF(INLIST(monthrep,1,3,5,7,8,10,12),31,IIF(monthrep=2,28,30))
IF yearrep=0
   RETURN
ENDIF
IF USED('curblist')
   SELECT curblist
   USE
ENDIF

IF USED('curbotp')
   SELECT curbotp
   USE
ENDIF

IF USED('curdotp')
   SELECT curdotp
   USE
ENDIF

IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curEmPeop')
   SELECT curEmPeop
   USE
ENDIF
IF USED('curEmJob')
   SELECT curEmJob
   USE
ENDIF


WITH fSupl
     .butPrn.Visible=.F.
     .butRet.Visible=.F.
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH   

SELECT * FROM datjob INTO CURSOR curEmjob READWRITE
SELECT curEmjob
APPEND FROM datJobOut
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) ALL 
DELETE FOR tr#1
DELETE FOR !EMPTY(dateOut).AND.dateOut<dBeg
*DELETE FOR !EMPTY(dateOut).AND.dateOut>dEnd
DELETE FOR dateBeg>dEnd
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL 
INDEX ON STR(nidpeop,5)+STR(tr,1) TAG T1

*SELECT * FROM people WHERE !lvn INTO CURSOR curEmPeop READWRITE 
SELECT * FROM people INTO CURSOR curEmPeop READWRITE 
SELECT curEmPeop
APPEND FROM peopout 
DELETE FOR !SEEK(STR(nid,5),'curEmJob',1)

SELECT curEmPeop
DELETE FOR !EMPTY(date_out).AND.date_out<dBeg
DELETE FOR date_in>dEnd
ALTER TABLE curEmPeop ADD COLUMN kat N(1)
INDEX ON date_in TAG T1
INDEX ON date_out TAG T2
SET ORDER TO 1
SELECT * FROM datalist INTO CURSOR curblist READWRITE
DELETE FOR curblist.dend<dBeg   &&удаляем записи где дата окончания б/л меньше начала периода
DELETE FOR curblist.dBeg>dEnd   &&удаляем записи где дата начала б/л больше окончания периода 
INDEX ON nidpeop TAG T1

SELECT * FROM datotp WHERE kodotp=6 INTO CURSOR curbotp READWRITE
DELETE FOR curbotp.endotp<dBeg   &&удаляем записи где дата окончания отпуска меньше начала периода
DELETE FOR curbotp.begotp>dEnd   &&удаляем записи где дата начала отпуска больше окончания периода 
INDEX ON nidpeop TAG T1

SELECT * FROM datotp WHERE kodotp=8 INTO CURSOR curdotp READWRITE
DELETE FOR curdotp.endotp<dBeg   &&удаляем записи где дата окончания отпуска меньше начала периода
DELETE FOR curdotp.begotp>dEnd   &&удаляем записи где дата начала отпуска больше окончания периода 
INDEX ON nidpeop TAG T1
SELECT curEmPeop

REPLACE kat WITH IIF(SEEK(STR(nid,5)+STR(1,1),'curEmJob',1),curEmJob.kat,0) ALL

CREATE CURSOR curPrn (nDay N(2),n1 N(4),n1In N(4),n1Out N(4),n1Spis N(4),n1Sh N(4),n1Dek N(4),n1Bol N(4),n1Otp N(4),n1Av N(4),n2 N(4),n2In N(4),n2Out N(4),n2Spis N(4),n2Sh N(4),n2Dek N(4),n2Bol N(4),n2Otp N(4),n2Av N(4),;
              n3 N(4),n3In N(4),n3Out N(4),n3Spis N(4),n3Sh N(4),n3Dek N(4),n3Bol N(4),n3Otp N(4),n3Av N(4),n4 N(4),n4In N(4),n4Out N(4),n4Spis N(4),n4Sh N(4),n4Dek N(4),n4Bol N(4),n4Otp N(4),n4Av N(4),;
              nt N(4),ntIn N(4),ntOut N(4),ntSpis N(4),ntSh N(4),ntDek N(4),ntBol N(4),ntOtp N(4),ntAv N(4))              
FOR i=1 TO IIF(INLIST(monthrep,1,3,5,7,8,10,12),31,IIF(monthrep=2,28,30))
    SELECT curprn
    APPEND BLANK
    REPLACE nDay WITH i           
ENDFOR         
INDEX ON nDay TAG T1
SELECT curEmPeop
******** на начало периода**************
kvobeg=0
COUNT TO kvobeg FOR date_in<dBeg
COUNT TO kvobeg1 FOR date_in<dBeg.AND.INLIST(kat,1,5)


COUNT TO kvobeg2 FOR date_in<dBeg.AND.INLIST(kat,2,7)
COUNT TO kvobeg3 FOR date_in<dBeg.AND.kat=3
COUNT TO kvobeg4 FOR date_in<dBeg.AND.kat=4
SELECT curPrn
GO TOP
REPLACE nt WITH kvoBeg,n1 WITH kvobeg1,n2 WITH kvobeg2,n3 WITH kvobeg3,n4 WITH kvobeg4
****
SELECT curEmPeop
SCAN ALL
     ** приянтые
     IF date_in>=dBeg.AND.date_in<=dEnd
        SELECT curprn
        SEEK DAY(curEmpeop.date_in)
        REPLACE ntIn WITH ntIn+1
        DO CASE
           CASE INLIST(curEmpeop.kat,1,5)
                REPLACE n1In WITH n1In+1           
           CASE INLIST(curEmpeop.kat,2,7)
                REPLACE n2In WITH n2In+1                
           CASE curEmpeop.kat=3
                REPLACE n3In WITH n3In+1    
           CASE curEmpeop.kat=4
                REPLACE n4In WITH n4In+1                           
        ENDCASE        
        SELECT curEmPeop
     ENDIF 
     ** уволенные
     IF date_out>=dBeg.AND.date_out<=dEnd
        SELECT curprn
        SEEK DAY(curEmpeop.date_out)
        REPLACE ntOut WITH ntOut+1
        DO CASE
           CASE INLIST(curEmpeop.kat,1,5)
                REPLACE n1Out WITH n1Out+1           
           CASE INLIST(curEmpeop.kat,2,7)
                REPLACE n2Out WITH n2Out+1                
           CASE curEmpeop.kat=3
                REPLACE n3Out WITH n3Out+1    
           CASE curEmpeop.kat=4
                REPLACE n4Out WITH n4Out+1                           
        ENDCASE
        SELECT curEmPeop
     ENDIF   
     ** больничные листы
    * IF SEEK(curEmpeop.nid,'curBlist',1)
    *    SELECT curBList
    *    SET FILTER TO nidpeop=curEmpeop.nid
    *    sch=0
    *    fch=0
    *    SCAN ALL
    *         DO CASE
    *            CASE curBlist.dBeg>=begList.AND.curBlist.dEnd<=endlist  &&дата начала>=начала периода и дата окончания<=окончания периода
    *                 sch=DAY(curBlist.dBeg)
    *                 fch=DAY(curBList.dEnd)
    *            CASE curBlist.dBeg<begList.AND.curBlist.dEnd<=endlist.AND.curBList.dEnd>=begList   &&дата начала<начала периода и дата окончания<=окончания периода
     *                sch=1
    *                 fch=DAY(curBList.dEnd)                    
    *            CASE curBlist.dBeg<begList.AND.curBlist.dEnd>endlist   &&дата начала<начала периода и дата окончания>окончания периода
    *                 sch=1
    *                 fch=dayMonth
    *            CASE curBlist.dBeg>=begList.AND.curBlist.dEnd>endlist   &&дата начала>=началу периода и дата окончания>окончания периода
    *                 sch=DAY(curBlist.dbeg)
    *                 fch=dayMonth
    *         ENDCASE
    *         SELECT curprn
    *         SEEK sch
     *        SCAN WHILE nday<=fch
    *              REPLACE ntBol WITH ntBol+1
    *              DO CASE
    *                 CASE INLIST(curEmpeop.kat,1,5)
    *                      REPLACE n1Bol WITH n1Bol+1           
    *                 CASE INLIST(curEmpeop.kat,2,7)
    *                      REPLACE n2Bol WITH n2Bol+1                
    *                 CASE curEmpeop.kat=3
    *                      REPLACE n3Bol WITH n3Bol+1    
    *                 CASE curEmpeop.kat=4
    *                      REPLACE n4Bol WITH n4Bol+1                           
    *              ENDCASE
    *         ENDSCAN             
    *         SELECT curBlist
    *    ENDSCAN 
    *    SET FILTER TO 
    * ENDIF
     SELECT curEmPeop
     ** отпуска за свой счёт
     IF SEEK(curEmpeop.nid,'curBotp',1)
        SELECT curBotp
        SET FILTER TO nidpeop=curEmpeop.nid
        sch=0
        fch=0
        SCAN ALL
             DO CASE
                CASE curBotp.begOtp>=begList.AND.curBotp.endOtp<=endlist  &&дата начала>=начала периода и дата окончания<=окончания периода
                     sch=DAY(curBotp.begOtp)
                     fch=DAY(curBotp.endOtp)
                CASE curBotp.begOtp<begList.AND.curBotp.endOtp<=endlist.AND.curBotp.endOtp>=begList   &&дата начала<начала периода и дата окончания<=окончания периода
                     sch=1
                     fch=DAY(curBotp.endOtp)                    
                CASE curBotp.begOtp<begList.AND.curBotp.endOtp>endlist   &&дата начала<начала периода и дата окончания>окончания периода
                     sch=1
                     fch=dayMonth
                CASE curBotp.begOtp>=begList.AND.curBotp.endOtp>endlist   &&дата начала>=началу периода и дата окончания>окончания периода
                     sch=DAY(curBotp.begOtp)
                     fch=dayMonth
             ENDCASE
             SELECT curprn
             SEEK sch
             SCAN WHILE nday<=fch
                  REPLACE ntOtp WITH ntOtp+1
                  DO CASE
                     CASE INLIST(curEmpeop.kat,1,5)
                          REPLACE n1Otp WITH n1Otp+1           
                     CASE INLIST(curEmpeop.kat,2,7)
                          REPLACE n2Otp WITH n2Otp+1                
                     CASE curEmpeop.kat=3
                          REPLACE n3Otp WITH n3Otp+1    
                     CASE curEmpeop.kat=4
                          REPLACE n4Otp WITH n4Otp+1                           
                  ENDCASE
             ENDSCAN             
             SELECT curBotp
        ENDSCAN 
        SET FILTER TO 
     ENDIF
     SELECT curEmpeop
    
    ** отпуск по уходу
    IF curEmpeop.dekotp
       sch=0
       fch=0
       DO CASE 
          CASE EMPTY(curEmPeop.bdekotp).AND.EMPTY(curEmPeop.ddekotp)
               sch=1
               fch=dayMonth
          CASE !EMPTY(curEmPeop.bdekotp).AND.!EMPTY(curEmPeop.ddekotp)     
               sch=IIF(bdekotp<beglist,1,IIF(bdekotp>endlist,0,DAY(bdekotp)))
               fch=IIF(ddekotp<beglist,0,IIF(BETWEEN(ddekotp,beglist,endlist),DAY(ddekotp),dayMonth))         
          CASE !EMPTY(curEmPeop.bdekotp).AND.EMPTY(curEmPeop.ddekotp)
               sch=IIF(bdekotp<beglist,1,IIF(bdekotp>endlist,0,DAY(bdekotp)))
               fch=dayMonth
          CASE EMPTY(curEmPeop.bdekotp).AND.!EMPTY(curEmPeop.ddekotp)          
               sch=1
               fch=IIF(ddekotp<beglist,0,IIF(BETWEEN(ddekotp,beglist,endlist),DAY(ddekotp),dayMonth))
               
       ENDCASE 
       SELECT curprn
       SEEK sch
       SCAN WHILE nday<=fch
            REPLACE ntDek WITH ntDek+1
            DO CASE
               CASE INLIST(curEmpeop.kat,1,5)
                    REPLACE n1Dek WITH n1Dek+1           
               CASE INLIST(curEmpeop.kat,2,7)
                    REPLACE n2Dek WITH n2Dek+1                
               CASE curEmpeop.kat=3
                    REPLACE n3Dek WITH n3Dek+1    
               CASE curEmpeop.kat=4
                    REPLACE n4Dek WITH n4Dek+1                           
            ENDCASE
       ENDSCAN                    
     ENDIF 
     *---неполный рабочий день
     SELECT curEmJob
     SEEK STR(curEmPeop.nid,5)
     ksecx=0
     SCAN WHILE nidpeop=curEmPeop.nid
          ksecx=ksecx+kse
     ENDSCAN    
     IF ksecx<1
        SELECT curprn
        SCAN ALL        
             REPLACE ntSh WITH ntSh+1
             DO CASE
                CASE INLIST(curEmpeop.kat,1,5)
                     REPLACE n1Sh WITH n1Sh+1           
                CASE INLIST(curEmpeop.kat,2,7)
                     REPLACE n2Sh WITH n2Sh+1                
                CASE curEmpeop.kat=3
                     REPLACE n3Sh WITH n3Sh+1    
                CASE curEmpeop.kat=4
                     REPLACE n4Sh WITH n4Sh+1                           
            ENDCASE
        ENDSCAN     
     ENDIF   
     SELECT curEmPeoP   
ENDSCAN
SELECT curprn
GO TOP
n1rep=n1
n2rep=n2
n3rep=n3
n4rep=n4
ntrep=nt
DO WHILE !EOF()
   *REPLACE n1 WITH n1rep,n1Spis WITH n1rep+n1In-n1Out,n2 WITH n2rep,n2Spis WITH n2rep+n2In-n2Out,n3 WITH n3rep,n3Spis WITH n3rep+n3In-n3Out,;
   *        n4 WITH n4rep,n4Spis WITH n4rep+n4In-n4Out,nt WITH ntrep,ntSpis WITH ntrep+ntIn-ntOut,n1Av WITH n1Spis-n1Dek-n1Otp-n1bol,;
   *        n2Av WITH n2Spis-n2Dek-n2Otp-n2bol,n3Av WITH n3Spis-n3Dek-n3Otp-n3bol,n4Av WITH n4Spis-n4Dek-n4Otp-n4bol,ntAv WITH ntSpis-ntDek-ntOtp-ntbol       
   REPLACE n1 WITH n1rep,n1Spis WITH n1rep+n1In-n1Out,n2 WITH n2rep,n2Spis WITH n2rep+n2In-n2Out,n3 WITH n3rep,n3Spis WITH n3rep+n3In-n3Out,;
           n4 WITH n4rep,n4Spis WITH n4rep+n4In-n4Out,nt WITH ntrep,ntSpis WITH ntrep+ntIn-ntOut,n1Av WITH n1Spis-n1Dek-n1Otp,;
           n2Av WITH n2Spis-n2Dek-n2Otp,n3Av WITH n3Spis-n3Dek-n3Otp,n4Av WITH n4Spis-n4Dek-n4Otp,ntAv WITH ntSpis-ntDek-ntOtp       
   n1rep=n1Spis
   n2rep=n2Spis
   n3rep=n3Spis
   n4rep=n4Spis
   ntrep=ntSpis
   SKIP
ENDDO

APPEND BLANK
REPLACE nday WITH 32
SUM n1In,n2In,n3In,n4In,ntIn,n1Out,n2Out,n3Out,n4Out,ntOut TO n1In_cx,n2In_cx,n3In_cx,n4In_cx,ntIn_cx,n1Out_cx,n2Out_cx,n3Out_cx,n4Out_cx,ntOut_cx
SUM n1,n2,n3,n4,nt TO n1_cx,n2_cx,n3_cx,n4_cx,nt_cx
SUM n1Spis,n2Spis,n3Spis,n4Spis,ntSpis TO n1Spis_cx,n2Spis_cx,n3Spis_cx,n4Spis_cx,ntSpis_cx
SUM n1Av,n2Av,n3Av,n4Av,ntAv TO n1Av_cx,n2Av_cx,n3Av_cx,n4Av_cx,ntAv_cx
GO BOTTOM
REPLACE n1In WITH n1In_cx,n2In WITH n2In_cx,n3In WITH n3In_cx,n4In WITH n4In_cx,ntIn WITH ntIn_cx,n1Out WITH n1Out_cx,n2Out WITH n2Out_cx,n3Out WITH n3Out_cx,n4Out WITH n4Out_cx,ntOut WITH ntOut_cx
APPEND BLANK
REPLACE nday WITH 33
REPLACE n1 WITH n1_cx/dayMonth,n2 WITH n2_cx/dayMonth,n3 WITH n3_cx/dayMonth,n4 WITH n4_cx/dayMonth,nt WITH nt_cx/dayMonth
REPLACE n1Spis WITH n1Spis_cx/dayMonth,n2Spis WITH n2Spis_cx/dayMonth,n3Spis WITH n3Spis_cx/dayMonth,n4Spis WITH n4Spis_cx/dayMonth,ntSpis WITH ntSpis_cx/dayMonth
REPLACE n1Av WITH n1Av_cx/dayMonth,n2Av WITH n2Av_cx/dayMonth,n3Av WITH n3Av_cx/dayMonth,n4Av WITH n4Av_cx/dayMonth,ntAv WITH ntAv_cx/dayMonth
DO startPrnToExcel WITH 'fSupl' 
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     FOR i=1 TO 46
        .Columns(i).ColumnWidth=5
     ENDFOR  
     
                       
     .Range(.Cells(1,1),.Cells(1,46)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=10
          .Value='Отчет о численности работающих по УЗ «Брестская центральная городская больница» за '+ALLTRIM(dim_month(monthrep))+' '+STR(yearrep,4)+'г.'
     ENDWITH  
     
     .Range(.Cells(2,2),.Cells(2,10)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlTop    
          .WrapText=.T.
          .Value='Врачи (физические лица без внешних совместителей)'
     ENDWITH  
     
     .Range(.Cells(2,11),.Cells(2,19)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlTop    
          .WrapText=.T.
          .Value='Средний медперсонал (физические лица без внешних совместителей)'
     ENDWITH  
     
     .Range(.Cells(2,20),.Cells(2,28)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlTop    
          .WrapText=.T.
          .Value='Младший медперсонал (физические лица без внешних совместителей)'
     ENDWITH  
     
     .Range(.Cells(2,29),.Cells(2,37)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlTop    
          .Value='Прочий персонал (физические лица без внешних совместителей)'
          .WrapText=.T.
     ENDWITH       
     
     
     .Range(.Cells(2,38),.Cells(2,46)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlTop    
          .WrapText=.T.
          .Value='Всего (физические лица без внешних совместителей)'
     ENDWITH            
     
     .Range(.Cells(2,1),.Cells(3,1)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .Orientation=90
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlCenter          
          .WrapText=.T.
          .Value='Дата'
     ENDWITH      
     
     .cells(3,2).Value='Врачи'     
     .cells(3,3).Value='Принято'
     .cells(3,4).Value='Уволено'
     .cells(3,5).Value='Списочная численность'                  
     .cells(3,6).Value='в том числе с неполныым рабочим днём'
     .cells(3,7).Value='Д/О'
     .cells(3,8).Value='Больничные листы'
     .cells(3,9).Value='Отпуск без сохранения зарплаты'                  
     .cells(3,10).Value='Среднесписочная численность'                       
     
     .cells(3,11).Value='Средний медперсонал'
     .cells(3,12).Value='Принято'
     .cells(3,13).Value='Уволено'
     .cells(3,14).Value='Списочная численность'                  
     .cells(3,15).Value='в том числе с неполныым рабочим днём'
     .cells(3,16).Value='Д/О'
     .cells(3,17).Value='Больничные листы'
     .cells(3,18).Value='Отпуск без сохранения зарплаты'                  
     .cells(3,19).Value='Среднесписочная численность'                       
     
     .cells(3,20).Value='Младший медперсонал'
     .cells(3,21).Value='Принято'
     .cells(3,22).Value='Уволено'
     .cells(3,23).Value='Списочная численность'                  
     .cells(3,24).Value='в том числе с неполныым рабочим днём'
     .cells(3,25).Value='Д/О'
     .cells(3,26).Value='Больничные листы'
     .cells(3,27).Value='Отпуск без сохранения зарплаты'                  
     .cells(3,28).Value='Среднесписочная численность'                       
     
     .cells(3,29).Value='Прочий персонал'
     .cells(3,30).Value='Принято'
     .cells(3,31).Value='Уволено'
     .cells(3,32).Value='Списочная численность'                  
     .cells(3,33).Value='в том числе с неполныым рабочим днём'
     .cells(3,34).Value='Д/О'
     .cells(3,35).Value='Больничные листы'
     .cells(3,36).Value='Отпуск без сохранения зарплаты'                  
     .cells(3,37).Value='Среднесписочная численность'                       
     
     .cells(3,38).Value='ВСЕГО'
     .cells(3,39).Value='Принято'
     .cells(3,40).Value='Уволено'
     .cells(3,41).Value='Списочная численность'                  
     .cells(3,42).Value='в том числе с неполныым рабочим днём'
     .cells(3,43).Value='Д/О'
     .cells(3,44).Value='Больничные листы'
     .cells(3,45).Value='Отпуск без сохранения зарплаты'                  
     .cells(3,46).Value='Среднесписочная численность'                                           
     
     .Range(.Cells(3,2),.Cells(3,46)).Select
     WITH objExcel.Selection
          .Orientation=90
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=xlCenter             
          .WrapText=.T.      
     ENDWITH      
     
     numberRow=4
     yearMonth=''
     SELECT curprn
     DO storezeropercent
     SCAN ALL        
          .cells(numberRow,1).Value=IIF(nDay<32,nDay,IIF(nDay=32,'итого','ср.'))
          .cells(numberRow,2).Value=IIF(n1#0,n1,'')
          .cells(numberRow,3).Value=IIF(n1In#0,n1In,'')
          .cells(numberRow,4).Value=IIF(n1Out#0,n1Out,'')
          .cells(numberRow,5).Value=IIF(n1Spis#0,n1Spis,'')
          .cells(numberRow,6).Value=IIF(n1Sh#0,n1Sh,'')
          .cells(numberRow,7).Value=IIF(n1Dek#0,n1Dek,'')
          .cells(numberRow,8).Value=IIF(n1Bol#0,n1Bol,'')
          .cells(numberRow,9).Value=IIF(n1Otp#0,n1Otp,'')
          .cells(numberRow,10).Value=IIF(n1Av#0,n1Av,'')
           
          .cells(numberRow,11).Value=IIF(n2#0,n2,'')          
          .cells(numberRow,12).Value=IIF(n2In#0,n2In,'')
          .cells(numberRow,13).Value=IIF(n2Out#0,n2Out,'')
          .cells(numberRow,14).Value=IIF(n2Spis#0,n2Spis,'')
          .cells(numberRow,15).Value=IIF(n2Sh#0,n2Sh,'')
          .cells(numberRow,16).Value=IIF(n2Dek#0,n2Dek,'')
          .cells(numberRow,17).Value=IIF(n2Bol#0,n2Bol,'')
          .cells(numberRow,18).Value=IIF(n2Otp#0,n2Otp,'')
          .cells(numberRow,19).Value=IIF(n2Av#0,n2Av,'')       
        
          .cells(numberRow,20).Value=IIF(n3#0,n3,'')      
          .cells(numberRow,21).Value=IIF(n3In#0,n3In,'')
          .cells(numberRow,22).Value=IIF(n3Out#0,n3Out,'')  
          .cells(numberRow,23).Value=IIF(n3Spis#0,n3Spis,'')
          .cells(numberRow,24).Value=IIF(n3Sh#0,n3Sh,'')
          .cells(numberRow,25).Value=IIF(n3Dek#0,n3Dek,'')
          .cells(numberRow,26).Value=IIF(n3Bol#0,n3Bol,'')
          .cells(numberRow,27).Value=IIF(n3Otp#0,n3Otp,'')
          .cells(numberRow,28).Value=IIF(n3Av#0,n3Av,'')    
          
          .cells(numberRow,29).Value=IIF(n4#0,n4,'')
          .cells(numberRow,30).Value=IIF(n4In#0,n4In,'')
          .cells(numberRow,31).Value=IIF(n4Out#0,n4Out,'')
          .cells(numberRow,32).Value=IIF(n4Spis#0,n4Spis,'')
          .cells(numberRow,33).Value=IIF(n4Sh#0,n4Sh,'')
          .cells(numberRow,34).Value=IIF(n4Dek#0,n4Dek,'')
          .cells(numberRow,35).Value=IIF(n4Bol#0,n4Bol,'')
          .cells(numberRow,36).Value=IIF(n4Otp#0,n4Otp,'')
          .cells(numberRow,37).Value=IIF(n4Av#0,n4Av,'')    

          .cells(numberRow,38).Value=IIF(nt#0,nt,'')
          .cells(numberRow,39).Value=IIF(ntIn#0,ntIn,'')
          .cells(numberRow,40).Value=IIF(ntOut#0,ntOut,'')          
          .cells(numberRow,41).Value=IIF(ntSpis#0,ntSpis,'')
          .cells(numberRow,42).Value=IIF(ntSh#0,ntSh,'')
          .cells(numberRow,43).Value=IIF(ntDek#0,ntDek,'')
          .cells(numberRow,44).Value=IIF(ntBol#0,ntBol,'')
          .cells(numberRow,45).Value=IIF(ntOtp#0,ntOtp,'')
          .cells(numberRow,46).Value=IIF(ntAv#0,ntAv,'')        
           DO fillpercent WITH 'fSupl'
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(2,1),.Cells(numberRow-1,46)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'   
         .Font.Size=10
         .WrapText=.T.
    ENDWITH 
    .Range(.Cells(1,1),.Cells(1,46)).Select
ENDWITH 
ON ERROR DO erSup
DO endPrnToExcel WITH 'fSupl' 
ON ERROR 
objExcel.Visible=.T. 