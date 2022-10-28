dateBook=DATE()
STORE '' TO fltPodr,fltType,fltDol,fltPers,tableItem

DIMENSION dimOption(5),dim_dek(3)
STORE .F. TO dimOption

logVac=.F.

dim_dek(1)=1   &&включать д/о
dim_dek(2)=0   &&исключать  д/о
dim_dek(3)=0   &&только  д/о

SELECT * FROM sprtype INTO CURSOR dopType READWRITE
SELECT dopType
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON kod TAG T1
SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1
SELECT * FROM sprkat INTO CURSOR dopKat READWRITE
SELECT dopKat
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON kod TAG T1
SELECT * FROM sprdolj INTO CURSOR dopDolj READWRITE
SELECT dopDolj
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Штатная книга'   
     DO addshape WITH 'fSupl',1,10,10,150,400,8 
     DO adlabMy WITH 'fSupl',1,' Дата ',.Shape1.Top+20,.Shape1.Left+20,300,0,.T.,1
     DO adTboxNew WITH 'fSupl','boxDate',.Shape1.Top+20,.lab1.Left+.lab1.Width+10,RetTxtWidth('99/99/999999'),dHeight,'dateBook',.F.,.T.,0
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxDate.Width-10)/2
     .boxDate.Left=.lab1.Left+.lab1.Width+10
     DO adCheckBox WITH 'fSupl','checkPodr','Подразделение',.boxDate.Top+.boxdate.Height+10,.Shape1.Left+5,150,dHeight,'dimOption(1)',0,.T.,"DO validCheckItem WITH 'dopPodr.otm,name','dopPodr','DO returnToPrnPodr WITH .T.','DO returnToPrnPodr WITH .F.'"   
     DO adCheckBox WITH 'fSupl','checkDol','Должность',.CheckPodr.Top,.Shape1.Left+5,150,dHeight,'dimOption(3)',0,.T.,"DO validCheckItem WITH 'dopDolj.otm,name','dopDolj','DO returnToPrnDol WITH .T.','DO returnToPrnDol WITH .F.'" 
     DO adCheckBox WITH 'fSupl','checkType','Тип работы',.checkPodr.Top+.checkPodr.Height+10,.Shape1.Left+5,150,dHeight,'dimOption(2)',0,.T.,"DO validCheckItem WITH 'dopType.otm,name','dopType','DO returnToPrnType WITH .T.','DO returnToPrnType WITH .F.'" 
     DO adCheckBox WITH 'fSupl','checkPers','Персонал',.checkType.Top,.Shape1.Left+5,150,dHeight,'dimOption(4)',0,.T.,"DO validCheckItem WITH 'dopKat.otm,name','dopKat','DO returnToPrnPers WITH .T.','DO returnToPrnPers WITH .F.'" 
     .checkPodr.Left=.Shape1.Left+(.Shape1.Width-.checkPodr.Width-.checkDol.Width-30)/2
     .checkDol.Left=.checkPodr.Left+.checkPodr.Width+30
     
     .checkType.Left=.Shape1.Left+(.Shape1.Width-.checkType.Width-.checkPers.Width-30)/2
     .checkPers.Left=.checkType.Left+.checkType.Width+30
     
     DO addOptionButton WITH 'fSupl',11,'вкючать д/о',.checkType.Top+.checkType.Height+10,.Shape1.Left+20,'dim_dek(1)',0,"DO procValOption WITH 'fSupl','dim_dek',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'исключать д/о',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_dek(2)',0,"DO procValOption WITH 'fSupl','dim_dek',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'только д/о',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_dek(3)',0,"DO procValOption WITH 'fSupl','dim_dek',3",.T. 
     
     .Option11.Left=.Shape1.Left+(.Shape1.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10 
     .Option13.Left=.Option12.Left+.Option12.Width+10 
     .Shape1.Height=.boxDate.Height+.checkPodr.Height*2+.Option11.Height+70
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.F.,.T.
     DO adButtonPrnToForm WITH 'DO bookprn WITH 1','DO bookprn WITH 2','fSupl.Release',.T.          
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape91.Height+20,.Shape1.Width  
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width   
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH  
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape91.Height+.butPrn.Height+60
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE returnToPrnPodr
PARAMETERS parRet
kvoPodr=0
IF parRet
   SELECT dopPodr
   fltPodr=''
   SCAN ALL
        IF fl 
           fltPodr=fltPodr+','+LTRIM(STR(kod))+','          
           kvoPodr=kvoPodr+1
        ENDIF 
   ENDSCAN
ELSE 
   fltPodr=''
   SELECT dopPodr
   REPLACE otm WITH '',fl WITH .F. ALL
    dimOption(1)=.F.
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
   *  .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
     .listBox1.Visible=.F. 
     dimOption(1)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='Подразделение'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE returnToPrnType
PARAMETERS parRet
kvoType=0
IF parRet
   SELECT dopType
   fltType='' 
   SCAN ALL
        IF fl 
           fltType=fltType+','+LTRIM(STR(kod))+','          
           kvoType=kvoType+1
        ENDIF 
   ENDSCAN
ELSE 
   fltType=''
   SELECT dopType
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
  *   .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')    
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .listBox1.Visible=.F. 
     dimOption(2)=IIF(kvoType>0,.T.,.F.)
     .checkType.Caption='Тип работы'+IIF(kvoType#0,'('+LTRIM(STR(kvoType))+')','') 
     .Refresh
ENDWITH 
***********************************************************************************************************************
PROCEDURE returnToPrnDol
PARAMETERS parRet
kvoDol=0
IF parRet
   SELECT dopDolj
   fltDol='' 
   SCAN ALL
        IF fl 
           fltDol=fltDol+','+LTRIM(STR(kod))+','          
           kvoDol=kvoDol+1
        ENDIF 
   ENDSCAN
ELSE 
   fltDol=''
   SELECT dopDolj
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
  *   .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     .lab25.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .listBox1.Visible=.F. 
     dimOption(3)=IIF(kvoDol>0,.T.,.F.)
     .checkDol.Caption='Должность'+IIF(kvoDol#0,'('+LTRIM(STR(kvoDol))+')','') 
     .Refresh
ENDWITH 
***********************************************************************************************************************
PROCEDURE returnToPrnPers
PARAMETERS parRet
kvoPers=0
IF parRet
   SELECT dopKat
   fltPers='' 
   SCAN ALL
        IF fl 
           fltPers=fltPers+','+LTRIM(STR(kod))+','          
           kvoPers=kvoPers+1
        ENDIF 
   ENDSCAN
ELSE 
   fltPers=''
   SELECT dopKat
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
  *   .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')   
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
     .cont11.Visible=.F.
     .cont12.Visible=.F.
     
     .listBox1.Visible=.F. 
     dimOption(4)=IIF(kvoPers>0,.T.,.F.)
     .checkPers.Caption='Персонал'+IIF(kvoPers#0,'('+LTRIM(STR(kvoPers))+')','') 
     .Refresh
ENDWITH 
*********************************************************************************************************
PROCEDURE validCheckItem
PARAMETERS parSource,parClick,parReturnTrue,parReturnFalse
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox') 
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.
     .listBox1.Visible=.T.     
     tableItem=parClick
     .listBox1.RowSource=parSource    
     .listBox1.procForClick='DO clickListItem WITH tableItem'
     .listBox1.procForKeyPress="DO keyPressItem WITH 'DO clickListItem WITH tableItem'"
  
     .cont11.procForClick=parReturnTrue
     .cont12.procForClick=parReturnFalse
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressItem
PARAMETERS parProc
DO CASE
   CASE LASTKEY()=27     
   CASE LASTKEY()=13
        &parProc
ENDCASE  
*************************************************************************************************************************
PROCEDURE clickListItem
PARAMETERS parTbl
SELECT &parTbl
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*********************************************************************************************************
PROCEDURE bookPrn
PARAMETERS parTerm
IF USED('curPrn')
   SELECT curPrn
   USE   
ENDIF
IF USED('curPeople')
   SELECT curPeople
   USE   
ENDIF
IF USED('curTarPeople')
   SELECT curTarPeople
   USE   
ENDIF
SELECT * FROM people INTO CURSOR curPeople READWRITE
SELECT curPeople
APPEND FROM peopout
INDEX ON nid TAG T1
INDEX ON num TAG T2
SET ORDER TO 1

SELECT * FROM datJob INTO CURSOR curTarPeople READWRITE 
SELECT curTarPeople
APPEND FROM datjobout

REPLACE dateuv WITH IIF(SEEK(nidpeop,'curPeople',1),curpeople.date_out,dateuv) ALL
DELETE FOR dateBeg>dateBook && дата начала больше даты книги
DELETE FOR !EMPTY(dateOut).AND.EMPTY(dateuv).AND.dateOut<=dateBook  &&дата окончания<=дате книги при пустой дате увольнения
DELETE FOR !EMPTY(dateuv).AND.dateuv<dateBook  &&дата увольнения<даты книги при заполненной дате увольнения

*DELETE FOR dateOut>=dateBook.AND.EMPTY(dateuv) &&дата окончания>=дате книги при пустой дате увольнения
*DELETE FOR !EMPTY(dateuv).AND.dateuv>dateBook  &&дата увольнения>даты книги при заполненной дате увольнения
IF !EMPTY(fltType)
   DELETE FOR !(','+LTRIM(STR(tr))+','$fltType)
ENDIF


REPLACE fio WITH IIF(SEEK(nidpeop,'curPeople',1),curPeople.fio,'') ALL 
REPLACE staj_in WITH IIF(SEEK(nidpeop,'curPeople',1),curPeople.staj_in,'') ALL 
REPLACE date_in WITH IIF(SEEK(nidpeop,'curPeople',1),curPeople.date_in,CTOD('  .  .    ')) ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE dekotp WITH IIF(SEEK(nidpeop,'curPeople',1),curPeople.dekotp,dekotp) ALL 

INDEX ON STR(np,3)+STR(nd,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 2

DO CASE
   CASE dim_dek(2)=1 
        DELETE FOR SEEK(nidpeop,'curPeople',1).AND.curPeople.dekOtp        
   CASE dim_dek(3)=1
        DELETE FOR SEEK(nidpeop,'curPeople',1).AND.!curPeople.dekOtp
ENDCASE
SCAN ALL
     DO actualStajToday WITH 'curTarPeople','curTarPeople.date_in','dateBook-1' 
     IF lkv
        REPLACE kv WITH IIF(SEEK(nidpeop,'curPeople',1),curPeople.kval,0)
     ELSE
        REPLACE kv WITH 0  
     ENDIF   
ENDSCAN
CREATE CURSOR curPrn (np N(3),nd N(3),kp N(3),kd N(3),named C(150), fio C(70),kse N(7,2),tr N(1),nametr C(15),kat N(1),logP L,Kpp N(3),KodKpp N(3),KodKpp1 N(3),npp N(3),vac L,nvac N(1),kodpeop N(5),kv N(1),dkv C(60), pkont N(3),staj_today C(10),primtxt C (100))
kserasp_kat=0
kserasp_cx=0
fltch=''
SELECT rasp      
SELECT * FROM rasp INTO CURSOR currasp READWRITE      
SELECT currasp
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1
IF !EMPTY(fltPodr)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
kseRasp=0
kseVac=0
ksePeop=0
SCAN ALL
     kseRasp=kse
     ksePeop=0
     SELECT curPrn
     APPEND BLANK
     REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kse WITH curRasp.kse,kd WITH curRasp.kd,kp WITH curRasp.kp,logp WITH .T.,kat WITH currasp.kat
     SELECT curTarPeople 
     IF SEEK(STR(curRasp.kp,3)+STR(curRasp.kd,3))  
        DO WHILE kp=curRasp.kp.AND.kd=curRasp.kd
                 SELECT curPrn
                 APPEND BLANK 
                 REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kp WITH curRasp.kp,kd WITH curRasp.kd,kat WITH curRasp.kat,kse WITH curTarPeople.kse,tr WITH curTarPeople.tr,;
                         kodpeop WITH curTarPeople.kodpeop,fio WITH curTarPeople.fio,nametr WITH IIF(SEEK(tr,'sprtype',1),sprtype.name,''),KodKpp WITH curRasp.KodKpp,KodKpp1 WITH curRasp.KodKpp1,;
                         pkont WITH IIF(tr=1,IIF(SEEK(curTarPeople.nidpeop,'curPeople',1),curPeople.pkont,0),0),kv WITH curTarPeople.kv,dkv WITH IIF(SEEK(kv,'sprkval',1),sprkval.name,''),staj_today WITH curTarPeople.staj_today,;
                         primTxt WITH IIF(SEEK(curTarPeople.nidpeop,'curPeople',1).AND.curPeople.dekOtp,'д/отп.'+IIF(!EMPTY(curPeople.bdekotp),' с '+DTOC(curPeople.bDekotp),'')+IIF(!EMPTY(curPeople.ddekotp),' до '+DTOC(curPeople.dDekotp),''),'')           
                         IF !EMPTY(curTarPeople.kdek)
                            REPLACE primtxt WITH ALLTRIM(primtxt)+'д.о '+ALLTRIM(curTarPeople.fiodek)
                         ENDIF
                 IF !curTarPeople.dekotp        
                    ksePeop=ksePeop+IIF(curprn.tr=4,0,curPrn.kse)              
                
                 ENDIF   
                
                 SELECT curTarPeople
                 SKIP
        ENDDO          
     ENDIF
     IF kseRasp-ksePeop>0.AND.EMPTY(fltType)  
        SELECT curPrn
        APPEND BLANK            
        REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kp WITH curRasp.kp,kat WITH curRasp.kat,kd WITH curRasp.kd,kse WITH kseRasp-ksePeop,fio WITH 'Вакантная',;
                KodKpp WITH curRasp.KodKpp,KodKpp1 WITH curRasp.KodKpp1,tr WITH 1,vac WITH .T. 
     ENDIF
     SELECT curRasp
ENDSCAN
SELECT curPrn
IF !EMPTY(fltPers)
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltPers)
ENDIF
REPLACE named WITH IIF(SEEK(curPrn.kd,'sprdolj',1),sprdolj.name,'') ALL
REPLACE nvac WITH 1 FOR vac
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
INDEX ON kp TAG T2
SET ORDER TO 1
SELECT curPrn
GO TOP
IF parTerm=1
   DO procForPrintAndPreview WITH 'repBook','штатная кгига',.T.,'shtatBookToExcel'
ELSE 
   DO procForPrintAndPreview WITH 'repBook','штатная кгига',.F. 
ENDIF 
*********************************************************************************************************
PROCEDURE shtatBookToExcel
DO startPrnToExcel WITH 'fSupl' 
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)

WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=6
     .Columns(2).ColumnWidth=40
     .Columns(3).ColumnWidth=10
     .Columns(4).ColumnWidth=10
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=10
     .Columns(7).ColumnWidth=10
     .Columns(8).ColumnWidth=15
     
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО'              
     .cells(2,3).Value='объем'
     .cells(2,4).Value='тип работы'
     .cells(2,5).Value='категория'
     .cells(2,6).Value='% контракт'
     .cells(2,7).Value='стаж'
     .cells(2,8).Value='примечание'
     SELECT curPrn 
     DO storezeropercent         
     numberRow=3
     kpOld=0
     kdOld=0
     SCAN ALL
          IF kpOld#kp
             kpOld=kp
             .Range(.Cells(numberRow,1),.Cells(numberRow,8)).Select
             With objExcel.Selection                   
                   .MergeCells=.T.
                   .HorizontalAlignment=xlCenter
                   .WrapText=.T.
                   .Value=IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'')
                   .Interior.Color=RGB(123,123,123)
             ENDWITH   
             numberRow=numberRow+1
          ENDIF 
          IF logP
             kdOld=kd
             .Range(.Cells(numberRow,1),.Cells(numberRow,2)).Select
             With objExcel.Selection                   
                   .MergeCells=.T.
                   .HorizontalAlignment=xlLeft
                   .WrapText=.T.
                   .Value=named                   
             ENDWITH
             .cells(numberRow,3).Value=IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kse,'')
             .Range(.Cells(numberRow,1),.Cells(numberRow,8)).Select
             objExcel.Selection.Interior.Color=RGB(208,202,198)
          ELSE    
             .cells(numberRow,1).Value=kodpeop
             .cells(numberRow,2).Value=fio
             .cells(numberRow,3).Value=IIF(kse#0,LTRIM(STR(kse,7,2)),'')
             .cells(numberRow,4).Value=nametr
             .cells(numberRow,5).Value=dkv
             .cells(numberRow,6).Value=IIF(pkont#0,pkont,'')
             .cells(numberRow,7).Value=ALLTRIM(staj_today)
             .cells(numberRow,8).Value=ALLTRIM(primtxt)
              DO fillpercent WITH 'fSupl'
           ENDIF      
           numberRow=numberRow+1
     ENDSCAN
     .Range(.Cells(2,1),.Cells(numberRow-1,8)).Select
     WITH objExcel.Selection
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .VerticalAlignment=1   
         .WrapText=.T.
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH   
    .Range(.Cells(2,1),.Cells(2,8)).Select  
    objExcel.Selection.HorizontalAlignment=xlCenter
    .Cells(2,1).Select
ENDWITH
DO endPrnToExcel WITH 'fSupl'
         
objExcel.Visible=.T.
 
 