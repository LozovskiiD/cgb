DIMENSION dim_ord(3)
dim_ord(1)=1 && ���� ���������
dim_ord(2)=0 &&  ���������� �����
dim_ord(3)=0 && ������� �����
STORE CTOD('  .  .    ') TO dBeg_cx,dEnd_cx

DIMENSION dim_knt(3)
dim_knt(1)=1 && ��������
dim_knt(2)=0 &&  �������
dim_knt(3)=0 && ��������+�������


DO procDimFlt
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='���� ��������� ����������'
     .Icon='kone.ico'
     .procexit='Do exitend' 
     DO procObjFlt     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
        
     DO adLabMy WITH 'fSupl',1,'������ ���������',.Shape2.Top+10,.Shape2.Left+5,.Shape2.Width-10,2,.F.,1 
     
     DO adtbox WITH 'fSupl',1,.Shape2.Left+20,.lab1.Top+.lab1.Height+10,RetTxtWidth('99/99/99999'),dHeight,'dBeg_cx',.F.,.T.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dEnd_cx',.F.,.T.,.F.
       
     .txtBox1.Left=.Shape2.Left+(.Shape2.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10   
     DO addOptionButton WITH 'fSupl',11,'��������',.txtBox1.Top+.txtBox1.Height+10,.Shape2.Left+20,'dim_knt(1)',0,"DO procValOption WITH 'fSupl','dim_knt',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'�������',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_knt(2)',0,"DO procValOption WITH 'fSupl','dim_knt',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'��������+�������',.Option11.Top,.Option12.Left+.Option12.Width+20,'dim_knt(3)',0,"DO procValOption WITH 'fSupl','dim_knt',3",.T. 
     .Option11.Left=.Shape2.Left+(.Shape2.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10    
     .Option13.Left=.Option12.Left+.Option12.Width+10    
               
     DO addOptionButton WITH 'fSupl',1,'���� ���������',.Option11.Top+.Option11.Height+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'����� ����������',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     DO addOptionButton WITH 'fSupl',3,'����� �������',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(3)',0,"DO procValOption WITH 'fSupl','dim_ord',3",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-.Option3.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+10   
     .Option3.Left=.Option2.Left+.Option2.Width+10  
      
     .Shape2.Height=.lab1.height+.txtBox1.Height+.Option1.Height*2+50
     
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
     DO adButtonPrnToForm WITH 'DO prnEndKont WITH 1','DO prnEndKont WITH 2','DO exitEnd',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width 
     .Height=.Shape1.Height+.Shape2.Height+.Shape91.Height+.butPrn.Height+60
     .Width=.Shape1.Width+20
     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitEnd
SELECT dopPodr
USE
SELECT dopKat
USE
SELECT dopDolj
USE
SELECT people
frmTop.Refresh  
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus           
fSupl.Release
******************************************************************************************************************************
PROCEDURE prnEndKont
PARAMETERS par1
IF USED('curprn')
   SELECT curprn
   USE
ENDIF 
IF USED('kontjob')
   SELECT kontjob
   USE
ENDIF 
SELECT * FROM people INTO CURSOR curprn READWRITE
ALTER TABLE curprn ADD COLUMN npp N(3)
ALTER TABLE curprn ADD COLUMN kp N(3)
ALTER TABLE curprn ADD COLUMN cnamep C(100)
ALTER TABLE curprn ADD COLUMN kd N(3)
ALTER TABLE curprn ADD COLUMN cnamed C(100)
ALTER TABLE curprn ADD COLUMN cprim C(50)
ALTER TABLE curprn ADD COLUMN kat N(2)
SELECT curprn
INDEX ON enddog TAG T1
INDEX ON fio TAG T2
INDEX ON STR(kp,3)+fio TAG T3
INDEX ON num TAG T4
SELECT curprn
DO CASE
   CASE dim_knt(1)=1   
        DELETE FOR dog#1
   CASE dim_knt(2)=1
        DELETE FOR dog#3
   CASE dim_knt(3)=1
        DELETE FOR !INLIST(dog,1,3)
ENDCASE
DELETE FOR !BETWEEN(enddog,dBeg_cx,dEnd_cx)

SELECT * FROM datjob WHERE SEEK(kodpeop,'curprn',4) INTO CURSOR kontJob READWRITE
SELECT kontJob
DELETE FOR !EMPTY(dateout)
DELETE FOR INLIST(tr,4,6)
DO CASE
   CASE dim_knt(1)=1   
        DELETE FOR tr#1
   CASE dim_knt(2)=1
        DELETE FOR !INLIST(tr,1,3)
   CASE dim_knt(3)=1
        DELETE FOR !INLIST(tr,1,3)
ENDCASE

INDEX ON kodpeop TAG T1
SELECT curprn
SCAN ALL
     SELECT kontJob
     SEEK curprn.num
     SELECT curprn
     REPLACE kp WITH kontJob.kp,kd WITH kontJob.kd,cnamep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),cnamed WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),cprim WITH IIF(dekotp,'�/�',''),kat WITH kontJob.kat
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

DO CASE
   CASE dim_ord(1)=1
        SET ORDER TO 1
   CASE dim_ord(2)=1
        SET ORDER TO 2
   CASE dim_ord(3)=1
        SET ORDER TO 3     
ENDCASE
SELECT curprn
nppcx=0
SCAN ALL
     nppcx=nppcx+1
     REPLACE npp WITH nppcx
ENDSCAN
GO TOP
DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'rependkont','��������� ����������',.T.,'endKontToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'rependkont','��������� ����������',.F.,'endKontToExcel'
ENDCASE 
********************************************************************************************************
PROCEDURE endKontToExcel
DO startPrnToExcel WITH 'fSupl'      
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=6
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=40
     .Columns(4).ColumnWidth=45           
     .Columns(5).ColumnWidth=10     
     .cells(2,1).Value='�'              
     .cells(2,2).Value='��� ����������'             
     .cells(2,3).Value='�������.'
     .cells(2,4).Value='���������'
     .cells(2,5).Value='���� ���������'
     .cells(2,5).WrapText=.T.
     numberRow=3 
     SELECT curPrn
     DO storezeropercent
     SCAN ALL             
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio        
          .cells(numberRow,3).Value=cnamep
          .cells(numberRow,4).Value=ALLTRIM(cnamed)
          .cells(numberRow,5).Value=IIF(!EMPTY(enddog),DTOC(enddog),'')
          DO fillpercent WITH 'fSupl' 
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(2,1),.Cells(numberRow-1,5)).Select
    WITH objExcel.Selection  
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin        
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH      
    .Cells(2,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'            
objExcel.Visible=.T.