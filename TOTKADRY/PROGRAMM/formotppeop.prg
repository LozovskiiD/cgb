IF !USED('sprtot')
   USE sprtot IN 0   
ENDIF 
SELECT * FROM sprtot WHERE kspr=4 INTO CURSOR typeotp READWRITE
SELECT typeotp
INDEX ON kod TAG T1
DIMENSION dimReport(3)
dimReport(1)=1 && уход в отпуск
dimReport(2)=0 &&  выход из отпуска
dimReport(3)=0 && нахождение в отпуске
STORE CTOD('  .  .    ') TO dBeg_cx,dEnd_cx

lTypeOtp=.F.
kvootp=0
onlyotp=.F.
fltotp=''

DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список находящихся в отпусках'
     .Icon='kone.ico'
     .procexit='Do exitotp' 
     DO procObjFlt     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
        
     DO adLabMy WITH 'fSupl',1,'период',.Shape2.Top+10,.Shape2.Left+5,.Shape2.Width-10,2,.F.,1 
     
     DO adtbox WITH 'fSupl',1,.Shape2.Left+20,.lab1.Top+.lab1.Height+10,RetTxtWidth('99/99/99999'),dHeight,'dBeg_cx',.F.,.T.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dEnd_cx',.F.,.T.,.F.
       
     .txtBox1.Left=.Shape2.Left+(.Shape2.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10  
     DO adCheckBox WITH 'fSupl','checkType','вид отпуска',.txtBox1.Top+.txtBox1.Height+10,.Shape1.Left+5,150,dHeight,'lTypeOtp',0,.T.,"DO validCheckItem WITH 'typeotp.otm,name','typeotp','DO returnToTypeOtp WITH .T.','DO returnToTypeOtp WITH .F.'"         
     
     
     DO addOptionButton WITH 'fSupl',1,'уход в отпуск',.checkType.Top+.checkType.Height+10,.Shape1.Left+10,'dimReport(1)',0,"DO procvaloption WITH 'fSupl','dimReport',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'выход из отпуска',.Option1.Top,.Option1.Left,'dimReport(2)',0,"DO procvaloption WITH 'fSupl','dimReport',2",.T. 
     DO addOptionButton WITH 'fSupl',3,'нахождение в отпуске',.Option1.Top,.Option2.Left,'dimReport(3)',0,"DO procvaloption WITH 'fSupl','dimReport',3",.T. 
     .Option1.Left=.Shape1.Left+(.Shape1.Width-.Option1.Width-.Option2.Width-.Option3.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+10
     .Option3.Left=.Option2.Left+.Option2.Width+10
                
     .Shape2.Height=.lab1.height+.checkType.Height+.txtBox1.Height+.Option1.Height+60
     .checkType.Left=.Shape1.Left+(.Shape1.Width-.checkType.Width)/2
     
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
     DO adButtonPrnToForm WITH 'DO prnOtpPeop WITH 1','DO prnOtpPeop WITH 2','DO exitotp',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width 
     .Height=.Shape1.Height+.Shape2.Height+.Shape91.Height+.butPrn.Height+60
     .Width=.Shape1.Width+20
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitotp
SELECT dopPodr
USE
SELECT dopKat
USE
SELECT dopDolj
USE
SELECT typeotp
USE
SELECT people
frmTop.Refresh  
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus           
fSupl.Release
******************************************************************************************************************************
PROCEDURE returnToTypeOtp
PARAMETERS parRet
kvootp=0
IF parRet
   SELECT typeotp
   fltotp=''
   onlyotp=.F.
   SCAN ALL
        IF fl 
           fltotp=fltotp+','+LTRIM(STR(kod))+','
           onlyotp=.T.
           kvootp=kvootp+1
        ENDIF 
   ENDSCAN
ELSE 
   strotp=''
   onlyotp=.F.
   SELECT typeotp
   REPLACE otm WITH '',fl WITH .F. ALL
    lTypeOtp=.F.
   GO TOP
ENDIF 
WITH fSupl    
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
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
     lTypeOtp=IIF(kvootp>0,.T.,.F.)
     .checkType.Caption='вид отпуска'+IIF(kvootp#0,'('+LTRIM(STR(kvootp))+')','') 
     .Refresh
ENDWITH 
******************************************************************************************************************************
PROCEDURE prnOtpPeop
PARAMETERS par1
IF EMPTY(dBeg_cx).OR.EMPTY(dEnd_cx)
   RETURN
ENDIF
IF USED('curprn')
   SELECT curprn
   USE
ENDIF 
IF USED('kontjob')
   SELECT kontjob
   USE
ENDIF 
IF USED('curotppeop')
   SELECT curotppeop
   USE
ENDIF
SELECT * FROM people INTO CURSOR curpeopotp READWRITE
SELECT curpeopotp
APPEND FROM peopout FOR date_out>dbeg_cx
*ALTER TABLE curpeopotp ADD COLUMN npp N(3)
*ALTER TABLE curpeopotp ADD COLUMN kp N(3)
*ALTER TABLE curpeopotp ADD COLUMN cnamep C(100)
*ALTER TABLE curpeopotp ADD COLUMN kd N(3)
*ALTER TABLE curpeopotp ADD COLUMN cnamed C(100)
*ALTER TABLE curpeopotp ADD COLUMN kat N(2)

SELECT curpeopotp
INDEX ON fio TAG T1
INDEX ON nid TAG T2
SET ORDER TO 1

SELECT * FROM datotp INTO CURSOR curprn READWRITE 
SELECT curprn
DO CASE 
   CASE dimReport(1)=1
        DELETE FOR !BETWEEN(begotp,dBeg_cx,dEnd_cx)
   CASE dimReport(2)=1
        DELETE FOR !BETWEEN(endotp,dBeg_cx,dEnd_cx)
   CASE dimReport(3)=1
        DELETE FOR !(BETWEEN(begotp,dBeg_cx,dEnd_cx).OR.BETWEEN(endotp,dBeg_cx,dEnd_cx)).OR.(begotp<=dBeg_cx.AND.endotp>=dEnd_cx)
ENDCASE 
ALTER TABLE curprn ADD COLUMN fio C(60)
ALTER TABLE curprn ADD COLUMN npp N(4)
ALTER TABLE curprn ADD COLUMN kp N(3)
ALTER TABLE curprn ADD COLUMN cnamep C(100)
ALTER TABLE curprn ADD COLUMN kd N(3)
ALTER TABLE curprn ADD COLUMN cnamed C(100)
ALTER TABLE curprn ADD COLUMN kat N(2)

*DELETE FOR !BETWEEN(begotp,dBeg_cx,dEnd_cx).AND.!BETWEEN(endotp,dBeg_cx,dEnd_cx)
INDEX ON nidpeop TAG T1
INDEX ON fio TAG T2
SET ORDER TO 1
IF kvootp>0
   DELETE FOR !(','+LTRIM(STR(kodotp))+','$fltOtp)
ENDIF 

SELECT * FROM datjob WHERE SEEK(nidpeop,'curpeopotp',2) INTO CURSOR kontJob READWRITE
SELECT kontJob
DELETE FOR !EMPTY(dateout)
DELETE FOR INLIST(tr,4,6)
INDEX ON nidpeop TAG T1

SELECT curprn
SCAN ALL
     SELECT kontJob
     SEEK curprn.nidpeop
     SELECT curprn
     REPLACE kp WITH kontJob.kp,kd WITH kontJob.kd,cnamep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),cnamed WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),kat WITH kontJob.kat
ENDSCAN

IF kvoPodr>0
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
IF kvoDolj>0
   DELETE FOR !(','+LTRIM(STR(kd))+','$fltDolj)   
ENDIF
IF kvoKat>0
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltKat)
ENDIF

DELETE FOR !SEEK(nidpeop,'curpeopotp',2)
REPLACE fio WITH IIF(SEEK(nidpeop,'curpeopotp',2),curpeopotp.fio,'') ALL 
SET ORDER TO 2

nppcx=0
SCAN ALL
     nppcx=nppcx+1
     REPLACE npp WITH nppcx
ENDSCAN
GO TOP

DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'reppeopotp','сотрудники, находящиеся в отпуске',.T.,'peopotpToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'reppeopotp','сотрудники, находящиеся в отпуске',.F.,'peopotpToExcel'
ENDCASE 
********************************************************************************************************
PROCEDURE peopotpToExcel
DO startPrnToExcel WITH 'fSupl'      
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
maxColumn=7
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=6
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=40
     .Columns(4).ColumnWidth=45           
     .Columns(5).ColumnWidth=8 
     .Columns(6).ColumnWidth=8     
     .Columns(7).ColumnWidth=20   
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО сотрудника'             
     .cells(2,3).Value='подразд.'
     .cells(2,4).Value='должность'
     .cells(2,5).Value='начало'
     .cells(2,6).Value='оконч.'
     .cells(2,7).Value='вид отпуска'
     numberRow=3 
     SELECT curPrn
     DO storezeropercent
     SCAN ALL             
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio        
          .cells(numberRow,3).Value=cnamep
          .cells(numberRow,4).Value=ALLTRIM(cnamed)
          .cells(numberRow,5).Value=IIF(!EMPTY(begotp),DTOC(begotp),'')
          .cells(numberRow,6).Value=IIF(!EMPTY(endotp),DTOC(endotp),'')
          .cells(numberRow,7).Value=IIF(SEEK(kodotp,'typeotp',1),typeotp.name,'')
          DO fillpercent WITH 'fSupl' 
          numberRow=numberRow+1         
     ENDSCAN
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter  
     .Range(.Cells(2,1),.Cells(numberRow-1,maxColumn)).Select
     WITH objExcel.Selection  
          .wrapText=.T.
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin 
          .VerticalAlignment=xlTop      
          .Font.Name='Times New Roman'
          .Font.Size=10
     ENDWITH      
     .Cells(2,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'            
objExcel.Visible=.T.