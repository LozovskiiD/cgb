DIMENSION dimOption(5),dim_ord(2), dim_dek(6),dim_log(4)
STORE .F. TO dimOption,dim_log
*dimOption(1) - подразделение
*dimOption(2) - должность
*dimOption(3) - тип работы
*dimOption(4) - персонал
*dimOption(5) - категори€

dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим

dim_dek(1)=1   &&включать д/о+б/л
dim_dek(2)=0   &&исключать  д/о+б/л
dim_dek(3)=0   &&только  д/о+б/л

dim_dek(4)=0   &&включать б/л
dim_dek(5)=0   &&исключать  б/л
dim_dek(6)=0   &&только  б/л



logNewFile=.F.
dateSpis=DATE()
USE setupExcel IN 0

SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprdolj INTO CURSOR dopDolj READWRITE
SELECT dopDolj
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprkat INTO CURSOR dopKat READWRITE
SELECT dopKat
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprtype INTO CURSOR dopType READWRITE
SELECT dopType
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprkval INTO CURSOR dopKval READWRITE
SELECT dopKval
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

STORE .F. TO onlyPodr,onlyDol
STORE '' TO fltch,fltPodr,fltKat,fltType,fltDolj,fltKval
STORE 0 TO kvoPodr,kvoDolj,kvoKat,kvoType,totKvo,nAge,kvoKval

fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='—писок сотрудников'
     .Icon='kone.ico'
     .procexit='Do exitprn'
     DO addshape WITH 'fSupl',1,10,10,150,450,8
     
     DO adCheckBox WITH 'fSupl','checkPodr','подразделенение',.Shape1.Top+10,.Shape1.Left+5,150,dHeight,'dimOption(1)',0,.T.,'DO validPodrPrn'          
     DO adCheckBox WITH 'fSupl','checkDolj','должность',.checkPodr.Top,.Shape1.Left,150,dHeight,'dimOption(2)',0,.T.,'DO validDoljPrn'  
     DO adCheckBox WITH 'fSupl','checkKval','категори€',.checkPodr.Top,.Shape1.Left,150,dHeight,'dimOption(5)',0,.T.,'DO validKvalPrn'  
     .checkPodr.Left=.Shape1.Left+(.Shape1.Width-.checkPodr.Width-.checkDolj.Width-.checkKval.Width-30)/2
     .checkDolj.Left=.checkPodr.Left+.checkPodr.Width+15
     .checkKval.Left=.checkDolj.Left+.checkDolj.Width+15
     
     DO adCheckBox WITH 'fSupl','checkType','тип работы',.checkPodr.Top+.checkPodr.Height+10,.checkPodr.Left,150,dHeight,'dimOption(3)',0,.T.,'DO validTypePrn'                           
     DO adCheckBox WITH 'fSupl','checkKat','персонал',.checkType.Top,.checkDolj.Left,150,dHeight,'dimOption(4)',0,.T.,'DO validKatPrn' 
     .checkType.Left=.Shape1.Left+(.Shape1.Width-.checkType.Width-.checkKat.Width-15)/2
     .checkKat.Left=.checkType.Left+.checkType.Width+15
     
     DO addOptionButton WITH 'fSupl',11,'включать д/о+б/л',.checkType.Top+.checkType.Height+10,.Shape1.Left+20,'dim_dek(1)',0,"DO procValOption WITH 'fSupl','dim_dek',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'исключать д/о+б/л',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_dek(2)',0,"DO procValOption WITH 'fSupl','dim_dek',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'только д/о+б/л',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_dek(3)',0,"DO procValOption WITH 'fSupl','dim_dek',3",.T. 
          
     .Option11.Left=.Shape1.Left+(.Shape1.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10 
     .Option13.Left=.Option12.Left+.Option12.Width+10 
     
     DO addOptionButton WITH 'fSupl',14,'включать б/л',.Option11.Top+.Option11.Height+10,.Option11.Left,'dim_dek(4)',0,"DO procValOption WITH 'fSupl','dim_dek',4",.T. 
     DO addOptionButton WITH 'fSupl',15,'исключать б/л',.Option14.Top,.Option12.Left,'dim_dek(5)',0,"DO procValOption WITH 'fSupl','dim_dek',5",.T. 
     DO addOptionButton WITH 'fSupl',16,'только б/л',.Option14.Top,.Option13.Left,'dim_dek(6)',0,"DO procValOption WITH 'fSupl','dim_dek',6",.T. 
          
     DO adLabMy WITH 'fSupl',11,'возраст до (включительно)',.Option14.Top+.Option14.Height+10,.Shape1.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxAge',.Option14.Top+.Option14.Height+10,.Shape1.Left,RetTxtWidth('9999'),dHeight,'nAge',.F.,.T.,0
     .lab11.Left=.Shape1.Left+(.Shape1.Width-.lab11.Width-.boxAge.Width-10)/2
     .boxAge.Left=.lab11.Left+.lab11.Width+10
     .lab11.Top=.boxAge.Top+(.boxAge.Height-.lab11.Height)+3
     
     .Shape1.Height=.checkPodr.Height*2+.Option11.Height*2+.boxAge.Height+60
     
     DO addshape WITH 'fSupl',3,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8
     DO adCheckBox WITH 'fSupl','checkPens','пенсионер',.Shape3.Top+10,.Shape1.Left+5,150,dHeight,'dim_log(1)',0,.T.
     DO adCheckBox WITH 'fSupl','checkInv','инвалид',.checkPens.Top,.checkPens.Left,150,dHeight,'dim_log(2)',0,.T.
     DO adCheckBox WITH 'fSupl','checkMany','многодетный',.checkPens.Top+.checkPens.Height+10,.checkPens.Left,150,dHeight,'dim_log(3)',0,.T.
     DO adCheckBox WITH 'fSupl','checkChaes','"чернобылец"',.checkMany.Top,.checkPens.Left,150,dHeight,'dim_log(4)',0,.T.
     .checkPens.Left=.Shape1.Left+(.Shape1.Width-.checkPens.Width-.checkInv.Width-10)/2
     .checkInv.Left=.checkPens.Left+.checkPens.Width+10
     .checkMany.Left=.Shape1.Left+(.Shape1.Width-.checkMany.Width-.checkChaes.Width-30)/2
     .checkChaes.Left=.checkMany.Left+.checkMany.Width+10
     .Shape3.Height=.checkPens.Height*2+30
     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape3.Top+.Shape3.Height+10,150,.Shape1.Width,8     
     DO addOptionButton WITH 'fSupl',1,'режим алфавитный',.Shape2.Top+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'режим штатный',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20 
     DO adCheckBox WITH 'fSupl','checkFile','каждое подразделение отдельно',.Option1.Top+.Option1.Height+10,.Shape2.Left+5,150,dHeight,'logNewFile',0,.T.,
     .checkFile.Left=.Shape2.Left+(.Shape2.Width-.checkFile.Width)/2
     
     DO adLabMy WITH 'fSupl',1,'сформировать на дату',.checkFile.Top+.checkFile.Height+10,.Shape2.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxDate',.checkFile.Top+.checkFile.Height+10,.Shape2.Left,RetTxtWidth('99/99/99999'),dHeight,'dateSpis',.F.,.T.,0
     .lab1.Left=.Shape2.Left+(.Shape2.Width-.lab1.Width-.boxDate.Width-10)/2
     .boxDate.Left=.lab1.Left+.lab1.Width+10
     .lab1.Top=.boxDate.Top+(.boxDate.Height-.lab1.Height)+3     
     .Shape2.Height=.Option1.Height+.checkFile.Height+.boxDate.height+40  
     
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape2.Top+.Shape2.Height+10,.Shape1.Width,.F.,.T.
     
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape3.Height+.Shape91.Height+30,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH 
     
     DO addListBoxMy WITH 'fSupl',2,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape3.Height+.Shape91.Height+30,.Shape1.Width  
     WITH .listBox2                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''
          .RowSource='setupexcel.otm,nameflt'             
          .procForClick='DO clickListExcel'
          .procForKeyPress='DO KeyPressListExcel' 
          .Visible=.F.     
     ENDWITH     
    
     *--------------------------------- нопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('просмотрW')*4-30)/2,.Shape91.Top+.Shape91.Height+10,RetTxtWidth('просмотрW'),dHeight+5,'печать','DO prnSpisPeop WITH 1','печать ведомости'
     *--------------------------------- нопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'просмотр','DO prnSpisPeop WITH 2','предварительный просмотр и печать ведомости'   
     *--------------------------------- нопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'дл€ Excel','Do procSetExcel','дл€ Excel'
          *--------------------------------- нопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont4',.cont3.Left+.Cont3.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'возврат','Do exitprn','возврат'
     
     
     *----------------------------- нопка прин€ть---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont11',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприн€ть')*2)-15)/2,.cont1.Top,RetTxtWidth('wприн€тьw'),.cont1.Height,'прин€ть','DO returnToPrn WITH .T.'
     .cont11.Visible=.F.
     *--------------------------------- нопка сброс-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont12',.cont11.Left+.cont11.Width+15,.Cont11.Top,.Cont11.Width,.cont1.Height,'сброс','DO returnToPrn WITH .F.'
     .cont12.Visible=.F.
     

     *----------------------------- нопка прин€ть---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont13',.shape1.Left+(.Shape1.Width-RetTxtWidth('wвозврат'))/2,.cont1.Top,RetTxtWidth('wвозвратw'),.cont1.Height,'возврат','DO returnExcel'
     .cont13.Visible=.F.
     
     DO addShape WITH 'fSupl',11,.Shape91.Left,.cont1.Top,.cont1.Height,.Shape91.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape91.Left,.Shape91.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.  
         
     .Height=.Shape1.Height+.Shape2.Height+.Shape3.Height+.Shape91.Height+.cont1.Height+60
     .Width=.Shape1.Width+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitPrn
SELECT setupExcel
USE
SELECT dopPodr
USE
SELECT dopKat
USE
SELECT dopType
USE
SELECT dopDolj
USE
SELECT people
frmTop.Refresh  
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus           
fSupl.Release
*******************************************************************************************************
PROCEDURE validPodrPrn
dimOption(1)=.T. 
WITH fSupl   
     .SetAll('Visible',.F.,'MyCommandButton') 
     .SetAll('Visible',.F.,'MyContLabel')  
     .cont11.Visible=.T.
     .cont12.Visible=.T.  
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopPodr.otm,name'  
     .listBox1.procForClick='DO clickListPodr'
     .listBox1.procForKeyPress='DO KeyPressListPodr' 
     .cont11.procForClick='DO returnToPrnPodr WITH .T.'
     .cont12.procForClick='DO returnToPrnPodr WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListPodr
SELECT dopPodr
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' Х ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListPodr
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListPodr 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnPodr
PARAMETERS parRet
kvoPodr=0
IF parRet
   SELECT dopPodr
   fltPodr=''
   onlyPodr=.F.
   SCAN ALL
        IF fl 
           fltPodr=fltPodr+','+LTRIM(STR(kod))+','
           onlyPodr=.T.
           kvoPodr=kvoPodr+1
        ENDIF 
   ENDSCAN
ELSE 
   strPodr=''
   onlyPodr=.F.
   SELECT dopPodr
   REPLACE otm WITH '',fl WITH .F. ALL
    dimOption(2)=.F.
   GO TOP
ENDIF 
WITH fSupl   
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.  
     .cont13.Visible=.F.  
     .listBox1.Visible=.F.  
         
     dimOption(1)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='подразделение'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH 
*******************************************************************************************************
PROCEDURE validDoljPrn
dimOption(2)=.T. 
WITH fSupl  
     .SetAll('Visible',.F.,'MyCommandButton')       
     .SetAll('Visible',.F.,'MyContLabel')  
     .cont11.Visible=.T.
     .cont12.Visible=.T.    
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopDolj.otm,name'  
     .listBox1.procForClick='DO clickListDolj'
     .listBox1.procForKeyPress='DO KeyPressListDol' 
     .cont11.procForClick='DO returnToPrnDolj WITH .T.'
     .cont12.procForClick='DO returnToPrnDolj WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListDolj
SELECT dopDolj
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' Х ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListDolj
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListDolj 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnDolj
PARAMETERS parRet
kvoDolj=0
IF parRet
   SELECT dopDolj
   fltDolj=''
   onlyDolj=.F.
   SCAN ALL
        IF fl 
           fltDolj=fltDolj+','+LTRIM(STR(kod))+','
           onlyDolj=.T.
           kvoDolj=kvoDolj+1
        ENDIF 
   ENDSCAN
ELSE 
   strDolj=''
   onlyDolj=.F.
   SELECT dopDolj
   REPLACE otm WITH '',fl WITH .F. ALL
    dimOption(2)=.F.
   GO TOP
ENDIF 
WITH fSupl 
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F. 
     .cont13.Visible=.F.   
     .listBox1.Visible=.F.   
     dimOption(2)=IIF(kvoDolj>0,.T.,.F.)
     .checkDolj.Caption='должность'+IIF(kvoDolj#0,'('+LTRIM(STR(kvoDolj))+')','') 
     .Refresh
ENDWITH 
*******************************************************************************************************
PROCEDURE validTypePrn
dimOption(3)=.T. 
WITH fSupl     
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.  
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopType.otm,name'  
     .listBox1.procForClick='DO clickListType'
     .listBox1.procForKeyPress='DO KeyPressListType' 
     .cont11.procForClick='DO returnToPrnType WITH .T.'
     .cont12.procForClick='DO returnToPrnType WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListType
SELECT dopType
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' Х ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListType
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListType 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnType
PARAMETERS parRet
kvoType=0
IF parRet
   SELECT dopType
   fltType=''
   onlyType=.F.
   SCAN ALL
        IF fl 
           fltType=fltType+','+LTRIM(STR(kod))+','
           onlyType=.T.
           kvoType=kvoType+1
        ENDIF 
   ENDSCAN
ELSE 
   strType=''
   onlyType=.F.
   SELECT dopType
   REPLACE otm WITH '',fl WITH .F. ALL
    dimOption(2)=.F.
   GO TOP
ENDIF 
WITH fSupl   
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.  
     .cont13.Visible=.F.  
     .listBox1.Visible=.F.     
     dimOption(3)=IIF(kvoType>0,.T.,.F.)
     .checkType.Caption='тип работы'+IIF(kvoType#0,'('+LTRIM(STR(kvoType))+')','') 
     .Refresh
ENDWITH 
*******************************************************************************************************
PROCEDURE validKatPrn
dimOption(4)=.T. 
*dimOption(2)=.F.
*dimOption(4)=.F.
WITH fSupl     
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.  
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopkat.otm,name'  
     .listBox1.procForClick='DO clickListkat'
     .listBox1.procForKeyPress='DO KeyPressListkat' 
     .cont11.procForClick='DO returnToPrnkat WITH .T.'
     .cont12.procForClick='DO returnToPrnkat WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListkat
SELECT dopkat
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' Х ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListkat
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListkat 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnkat
PARAMETERS parRet
kvokat=0
IF parRet
   SELECT dopkat
   fltkat=''
   onlykat=.F.
   SCAN ALL
        IF fl 
           fltkat=fltkat+','+LTRIM(STR(kod))+','
           onlykat=.T.
           kvokat=kvokat+1
        ENDIF 
   ENDSCAN
ELSE 
   strkat=''
   onlykat=.F.
   SELECT dopkat
   REPLACE otm WITH '',fl WITH .F. ALL
    dimOption(2)=.F.
   GO TOP
ENDIF 
WITH fSupl    
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F. 
     .cont13.Visible=.F.   
     .listBox1.Visible=.F.      
     dimOption(4)=IIF(kvokat>0,.T.,.F.)
     .checkkat.Caption='персонал'+IIF(kvokat#0,'('+LTRIM(STR(kvokat))+')','') 
     .Refresh
ENDWITH 
*******************************************************************************************************
PROCEDURE validKvalPrn
dimOption(5)=.T. 
WITH fSupl     
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.  
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopkval.otm,name'  
     .listBox1.procForClick='DO clickListkval'
     .listBox1.procForKeyPress='DO KeyPressListkval' 
     .cont11.procForClick='DO returnToPrnkval WITH .T.'
     .cont12.procForClick='DO returnToPrnkval WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListkval
SELECT dopkval
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' Х ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListkval
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListkval
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnkval
PARAMETERS parRet
kvokat=0
IF parRet
   SELECT dopkval
   fltkval=''
   onlykval=.F.
   SCAN ALL
        IF fl 
           fltkval=fltkval+','+LTRIM(STR(kod))+','
           onlykval=.T.
           kvokval=kvokval+1
        ENDIF 
   ENDSCAN
ELSE 
   strkval=''
   onlykval=.F.
   kvokval=0
   SELECT dopkval
   REPLACE otm WITH '',fl WITH .F. ALL
   dimOption(5)=.F.
   GO TOP
ENDIF 
WITH fSupl    
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F. 
     .cont13.Visible=.F.   
     .listBox1.Visible=.F.      
     dimOption(5)=IIF(kvokval>0,.T.,.F.)
     .checkkval.Caption='категори€'+IIF(kvokval#0,'('+LTRIM(STR(kvokval))+')','') 
     .Refresh
ENDWITH 
*******************************************************************************************************
PROCEDURE procSetExcel
WITH fSupl     
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont13.Visible=.T.  
     .listBox2.Visible=.T.
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListExcel
SELECT setupExcel
lrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' Х ','')
GO lrec
fSupl.listBox2.SetFocus
GO lrec
*************************************************************************************************************************
PROCEDURE keyPressListExcel
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListExcel 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnExcel
PARAMETERS parRet
WITH fSupl   
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.  
     .cont13.Visible=.F.
     .listBox2.Visible=.F.  
     .Refresh
ENDWITH 
***********************************************************************************************************
PROCEDURE prnSpisPeop
PARAMETERS par1
IF USED('curDatJob')
   SELECT curDatJob
   USE
ENDIF
IF USED('curPeople')
   SELECT curPeople
   USE
ENDIF
IF USED('curKurs')
   SELECT curKurs
   USE
ENDIF
IF USED('curOtpBol')
   SELECT curOtpBol
   USE
ENDIF
SELECT * FROM peoporder WHERE supord=60.AND.dateSpis>=dateBeg.AND.dateSpis<=dateEnd INTO CURSOR curOtpBol READWRITE
SELECT curOtpBol
INDEX ON nidpeop TAG T1

SELECT * FROM datKurs INTO CURSOR curKurs READWRITE 
SELECT curKurs
INDEX ON STR(kodpeop,5)+DTOS(perBeg) TAG T1 DESCENDING

SELECT * FROM datJob INTO CURSOR curDatJob READWRITE
SELECT curDatJob
APPEND FROM datjobout
DELETE FOR dateBeg>dateSpis
DELETE FOR dateOut<dateSpis.AND.!EMPTY(dateOut)
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) ALL
REPLACE dekotp WITH IIF(SEEK(kodpeop,'people',1),people.dekotp,.F.) ALL
INDEX ON STR(kodpeop,4)+STR(kse,5,2) DESCENDING TAG T1
INDEX ON nidpeop TAG T2
SET ORDER TO 1
IF kvoPodr>0
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
IF kvoDolj>0
   DELETE FOR !(','+LTRIM(STR(kd))+','$fltDolj)   
ENDIF
IF kvoKat>0
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltKat)
ENDIF
IF kvoType>0
   DELETE FOR !(','+LTRIM(STR(tr))+','$fltType)
ELSE
   DELETE FOR !INLIST(tr,1)   
ENDIF

DO CASE
   CASE dim_dek(2)=1
        DELETE FOR dekOtp
   CASE dim_dek(3)=1
        DELETE FOR !dekOtp
   CASE dim_dek(4)=1
        DELETE FOR dekOtp.AND.!SEEK(nidpeop,'curOtpBol',1)
   CASE dim_dek(5)=1
        DELETE FOR SEEK(nidpeop,'curOtpBol',1)
   CASE dim_dek(6)=1
        DELETE FOR !SEEK(nidpeop,'curOtpBol',1)
ENDCASE

SELECT * FROM people INTO CURSOR curPeople READWRITE 
SELECT curpeople
APPEND FROM peopout
ALTER TABLE curPeople ADD COLUMN np N(3)
ALTER TABLE curPeople ADD COLUMN nd N(3)
ALTER TABLE curPeople ADD COLUMN kp N(3)
ALTER TABLE curPeople ADD COLUMN kpSup N(3)	
ALTER TABLE curPeople ADD COLUMN kd N(3)
ALTER TABLE curPeople ADD COLUMN kse N(6,2)
ALTER TABLE curPeople ADD COLUMN tr N(1)
ALTER TABLE curPeople ADD COLUMN named C(100)
ALTER TABLE curPeople ADD COLUMN namep C(100)
ALTER TABLE curPeople ADD COLUMN npp N(4)
ALTER TABLE curPeople ADD COLUMN begKurs D
ALTER TABLE curPeople ADD COLUMN endKurs D
ALTER TABLE curPeople ADD COLUMN nAgep N(2)

SELECT curPeople
DELETE FOR !SEEK(nid,'curdatJob',2)
SCAN ALL
     IF SEEK(STR(num,4),'curDatJob',1)
        REPLACE kp WITH curDatJob.kp,kd WITH curDatJob.kd,kse WITH curdatJob.kse,tr WITH curDatJob.tr,kpSup WITH curDatJob.kp
        REPLACE kpSup WITH IIF(SEEK(kp,'sprpodr',1).AND.sprpodr.kodkpp>0,sprpodr.kodkpp,kpSup)
        DO actualStajToday WITH 'curPeople','curPeople.date_in','datespis'
        SELECT curPeople
     ENDIF 
     IF nAge>0.AND.!EMPTY(curPeople.age)
        nAge_cx=0
        DO CASE
           CASE MONTH(curPeople.age)<MONTH(dateSpis)
                nAge_cx=YEAR(dateSpis)-YEAR(curPeople.age)
           CASE MONTH(curPeople.age)=MONTH(dateSpis)     
                nAge_cx=YEAR(dateSpis)-YEAR(curPeople.age)-IIF(DAY(curPeople.age)>DAY(dateSpis),1,0)
           CASE MONTH(curPeople.age)>MONTH(dateSpis)
                nAge_cx=YEAR(dateSpis)-YEAR(curPeople.age)-1
        ENDCASE
        IF nAge_cx<=nAge
           REPLACE nAgep WITH nAge_cx         
        ENDIF           
     ENDIF      
ENDSCAN 
IF nAge>0
   DELETE FOR nAgep=0 
ENDIF
str_log=''
IF dim_log(1)
   str_log='pens'
ENDIF
IF dim_log(2)
   str_log=IIF(!EMPTY(str_log),'('+str_log+'.AND.inv','inv')
ENDIF
IF dim_log(3)
   str_log=IIF(!EMPTY(str_log),'('+str_log+'.AND.mchild','mchild')
ENDIF
IF dim_log(4)
   str_log=IIF(!EMPTY(str_log),'('+str_log+'.AND.chaes','chaes')
ENDIF
str_log=IIF('AND'$str_log,str_log+')',str_log)
IF !EMPTY(str_log)
   DELETE FOR !&str_log
ENDIF    

IF kvoKval>0
   DELETE FOR !(','+LTRIM(STR(kval))+','$fltKval)
ENDIF

REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE begKurs WITH IIF(SEEK(STR(num,5),'curKurs',1),curKurs.perBeg,begKurs) ALL
REPLACE endKurs WITH IIF(SEEK(STR(num,5),'curKurs',1),curKurs.perEnd,endKurs) ALL

DO CASE 
   CASE dim_ord(1)=1
        INDEX ON fio TAG T1
   CASE dim_ord(2)=1
        INDEX ON STR(np,3)+STR(nd,3) TAG T1
ENDCASE 
USE repspispodr.frx IN 0
SELECT repspispodr
LOCATE FOR objtype=9.AND.ALLTRIM(LOWER(comment))='kpsup'
REPLACE pagebreak WITH IIF(logNewFile,.T.,.F.)

USE 
SELECT curPeople
COUNT TO totKvo
nppcx=1
kpcx=0
SCAN ALL 
     IF kpcx#kp.AND.dim_ord(2)=1
        kpcx=kp
        nppcx=1
     ENDIF
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN 
GO TOP
DO CASE 
   CASE par1=1
        IF dim_ord(1)=1
           DO procForPrintAndPreview WITH 'repspislist','список сотрудников',.T.,'spisPodrToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'repspispodr','список сотрудников',.T.,'spisPodrToExcel'
        ENDIF    
   CASE par1=2
        IF dim_ord(1)=1
           DO procForPrintAndPreview WITH 'repspislist','список сотрудников',.F.,'spisPodrToExcel'
        ELSE 
           DO procForPrintAndPreview WITH 'repspispodr','список сотрудников',.F.,'spisPodrToExcel'
        ENDIF    
ENDCASE 
***********************************************************************************************************
PROCEDURE spisPodrToExcel
SELECT * FROM setupExcel WHERE setupExcel.fl INTO CURSOR curSetupExcel    
SELECT curPeople
COUNT TO maxRow
GO TOP 
new_num=1
GO TOP
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH   
*.lab15.Caption=' ‘ормируетс€ файл - '+ALLTRIM(nameSpisFlt)
 *       DO procObjectFrmResult WITH 1                 
 #DEFINE xlCenter -4108            
 #DEFINE xlLeft -4131              
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
 kvoColumn=RECCOUNT('curSetupExcel')+2
 WITH excelBook.Sheets(1)
      .Columns(1).ColumnWidth=5
      .Columns(2).ColumnWidth=49
      IF RECCOUNT('curSetupExcel')>0 
         SELECT curSetupExcel
         GO TOP
         FOR i=3 TO kvoColumn
             .Columns(i).ColumnWidth=curSetupExcel.widthCol
             .cells(2,i).Value=ALLTRIM(curSetupExcel.nameFlt)
             .cells(2,i).HorizontalAlignment= -4108
             .cells(2,i).WrapText=.T.
             SKIP                      
         ENDFOR
      ENDIF                       
      .Range(.Cells(1,1),.Cells(1,kvoColumn)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment= -4108
           *.VerticalAlignment=1
           .WrapText=.T.
           .Value='—писок сотрудников'
      ENDWITH   
      .cells(2,1).Value='є'
      .cells(2,1).HorizontalAlignment= -4108
      .Cells(2,2).Value='‘»ќ'                 
      .Cells(2,2).HorizontalAlignment= -4108                                   
      SELECT curPeople
      DO storezeropercent    
      numberRow=3
      kpcx=0
      SCAN ALL
           IF kpcx#kp.AND.dim_ord(2)=1
              kpcx=kp
              .Range(.Cells(numberRow,1),.Cells(numberRow,kvoColumn)).Select
              With objExcel.Selection
                   .MergeCells=.T.
                   .HorizontalAlignment= -4108
                   .WrapText=.T.
                   .Value=IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'')
                   .Interior.ColorIndex=35
               ENDWITH     
               numberRow=numberRow+1
           ENDIF
           .cells(numberRow,1).Value=curPeople.npp
           .cells(numberRow,2).Value=curPeople.fio
           SELECT curSetupExcel
           GO TOP
           FOR i=3 TO  RECCOUNT('curSetupExcel')+2
               IF !EMPTY(strtoved)
                  valuecell=strtoved
                
                  IF 'staj_today'$LOWER(valuecell)
                    * .cells(numberRow,i).Value=&valuecell
                    .cells(numberRow,i).numberFormat='@'
                  ENDIF
                  .cells(numberRow,i).Value=&valuecell
               ENDIF 
               SKIP 
           ENDFOR 
           SELECT curPeople
           DO fillpercent WITH 'fSupl'
           numberRow=numberRow+1
      ENDSCAN
      .Range(.Cells(1,1),.Cells(numberRow-1,kvoColumn)).Select
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
      .Range(.Cells(1,1),.Cells(1,kvoColumn)).Select
 ENDWITH 
 #UNDEFINE xlInsideHorizontal 
 WITH fSupl
      .cont1.Visible=.T.
      .cont2.Visible=.T.
      .cont3.Visible=.T.
      .cont4.Visible=.T.
      .Shape11.Visible=.F.
      .Shape12.Visible=.F.      
      .lab25.Visible=.F.      
ENDWITH               
objExcel.Visible=.T.
 
 
