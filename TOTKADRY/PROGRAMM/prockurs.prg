*******************************************************************************************************
DIMENSION dim_opt(2),dim_ord(2)
dim_opt(1)=1
dim_opt(2)=0
dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим
dateBeg=CTOD('  .  .   ')
dateEnd=CTOD('  .  .   ')
DO procDimFlt
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
SELECT kod,name,otm,fl FROM sprtot WHERE kspr=22 INTO CURSOR dopType READWRITE 
SELECT dopType
INDEX ON kod TAG T1
PUBLIC fltType,kvoType,kvoDolj,kvoPodr,kvoKat
fltType=''
STORE 0 TO kvoType,kvoDolj,kvoPodr,kvoKat
DIMENSION dimOption(4)
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сотрудников, прошедших курсы повышения квалификации'   
     DO procObjFlt    
     DO adCheckBox WITH 'fSupl','checkType','вид подготовки',.checkPodr.Top+.checkPodr.Height+10,.checkDolj.Left,150,dHeight,'dimOption(4)',0,.T.,"DO validCheckItem WITH 'dopType.otm,name','dopType','DO returnToPrnType WITH .T.','DO returnToPrnType WITH .F.'"    
     .checkType.Left=.Shape1.Left+(.Shape1.Width-.checkType.Width)/2
     .Shape1.Height=.checkType.Height*2+30
     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
     DO addOptionButton WITH 'fSupl',11,'последнее проходжение',.Shape2.Top+10,.Shape2.Left+20,'dim_opt(1)',0,'DO procValOptKurs WITH 1',.T. 
     DO addOptionButton WITH 'fSupl',12,'прохождение за период',.Option11.Top+.Option11.Height+10,.Option11.Left,'dim_opt(2)',0,'DO procValOptKurs WITH 2',.T. 
     .Option11.Left=.Shape2.Left+(.Shape2.Width-.Option11.Width)/2
     .Option12.Left=.Shape2.Left+(.Shape2.Width-.Option12.Width)/2   

     DO adtbox WITH 'fSupl',1,.Shape2.Left+20,.Option12.Top+.Option12.Height+10,RetTxtWidth('99/99/99999'),dHeight,'dateBeg',.F.,.F.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dateEnd',.F.,.F.,.F.
       
     .txtBox1.Left=.Shape2.Left+(.Shape1.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10             
     .Shape2.Height=.txtBox1.Height+.Option11.Height+60  
          
     DO addshape WITH 'fSupl',91,.Shape1.Left,.Shape2.Top+.Shape2.Height+20,150,.Shape1.Width,8
     DO addOptionButton WITH 'fSupl',1,'режим алфавитный',.Shape91.Top+10,.Shape91.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'режим штатный',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     .Option1.Left=.Shape91.Left+(.Shape91.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20         
     .Shape91.Height=.Option1.Height+20 
     
     DO addButtonOne WITH 'fSupl','butPrn',.Shape91.Left+(.Shape91.Width-RetTxtWidth('wсформироватьw')*2-15)/2,.Shape91.Top+.Shape91.Height+20,'сформировать','','DO prnKurs WITH .T.',39,RetTxtWidth('wсформироватьw'),'сформировать' 
     DO addButtonOne WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+15,.butPrn.Top,'возврат','','fSupl.Release',39,.butPrn.Width,'возврат'  
     
     DO addButtonOne WITH 'fSupl','cont11',.shape91.Left+(.Shape91.Width-(RetTxtWidth('wпринять')*2)-15)/2,.butPrn.Top,'принять','','DO returnToPrn',39,RetTxtWidth('wпринятьw'),'принять' 
     DO addButtonOne WITH 'fSupl','cont12',.cont11.Left+.cont11.Width+15,.butPrn.Top,'сброс','','DO returnToPrn',39,.cont11.Width,'сброс' 
        .cont11.Visible=.F.
        .cont12.Visible=.F.   
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width  
     
     .Height=.butPrn.Top+.butPrn.Height+20
     .Width=.Shape1.Width+20
     
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+40,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH   
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************************
PROCEDURE exitProcKurs
RELEASE ltType,kvoType,kvoDolj,kvoPodr,kvoKat
fSupl.Release
****************************************************************************************************************************
PROCEDURE procValOptKurs
PARAMETERS par1
STORE 0 TO dim_opt
dim_opt(par1)=1
fSupl.txtBox1.Enabled=IIF(dim_opt(1)=1,.F.,.T.)
fSupl.txtBox2.Enabled=IIF(dim_opt(1)=1,.F.,.T.)
fSupl.REfresh
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
     dimOption(4)=IIF(kvoType>0,.T.,.F.)
     .checkType.Caption='вид подготовки'+IIF(kvoType#0,'('+LTRIM(STR(kvoType))+')','') 
     .Refresh
ENDWITH  
****************************************************************************************************************************
PROCEDURE prnKurs
PARAMETERS parLog
IF dateBeg>dateEnd
   RETURN
ENDIF 
IF !USED('datKurs')
   USE datKurs IN 0
ENDIF 
IF USED('kursPrn')
   SELECT kursPrn
   USE
ENDIF
IF USED('curKurs')
   SELECT curKurs
   USE
ENDIF
IF USED('kursJob')
   SELECT kursJob
   USE
ENDIF
SELECT * FROM datJob INTO CURSOR kursJob READWRITE
SELECT kursJob
DELETE FOR INLIST(tr,4,2)
DELETE FOR !EMPTY(dateOut)
INDEX ON STR(nidpeop,5)+STR(tr,1) TAG T1
DO CASE
   CASE dim_opt(1)=1 
        SELECT * FROM datKurs INTO CURSOR curKurs READWRITE 
        SELECT curKurs
        INDEX ON STR(nidpeop,5)+DTOS(perBeg) TAG T1 DESCENDING              
        SELECT num,fio,kval,dkval,nkval,nid FROM people INTO CURSOR kursPrn READWRITE 
        SELECT kursPrn
        ALTER TABLE kursPrn ADD COLUMN npp N(6)
        ALTER TABLE kursPrn ADD COLUMN kodpeop N(5)
        ALTER TABLE kursPrn ADD COLUMN perBeg D
        ALTER TABLE kursPrn ADD COLUMN perEnd D
        ALTER TABLE kursPrn ADD COLUMN nameKurs C(100)
        ALTER TABLE kursPrn ADD COLUMN nameSchool C(100)
        ALTER TABLE kursPrn ADD COLUMN ntype N(2)
        ALTER TABLE kursPrn ADD COLUMN khours N(4)
        ALTER TABLE kursPrn ADD COLUMN kp N(3)
        ALTER TABLE kursPrn ADD COLUMN kd N(3)
        ALTER TABLE kursPrn ADD COLUMN np N(3)
        ALTER TABLE kursPrn ADD COLUMN nd N(3)
        ALTER TABLE kursPrn ADD COLUMN kat N(3)
        ALTER TABLE kursPrn ADD COLUMN ckval C(70)   
        ALTER table kursPrn ADD COLUMN nidpeop N(5)               
        REPLACE kodpeop WITH num, ckval WITH IIF(SEEK(kval,'sprkval',1),ALLTRIM(sprkval.name)+' ','')+IIF(!EMPTY(dkval),DTOC(dkval)+' ','')+ALLTRIM(nkval),nidpeop WITH nid ALL                                
        SELECT kursPrn       
        SET RELATION TO STR(nid,5) INTO curKurs ADDITIVE
        SCAN ALL             
             REPLACE perBeg WITH curKurs.perBeg,perEnd WITH curKurs.perEnd,nameKurs WITH curKurs.nameKurs,nameSchool WITH curKurs.nameSchool,ntype WITH curKurs.ntype,khours WITH curKurs.khours;
             kp WITH IIF(SEEK(STR(nid,5),'kursJob'),kursJob.kp,0),kd WITH kursJob.kd, kat WITH kursjob.kat             
        ENDSCAN 
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
        REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
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
           DELETE FOR !(','+LTRIM(STR(nType))+','$fltType)
        ENDIF  
        DELETE FOR EMPTY(perBeg).AND.EMPTY(perEnd)
        DO CASE
           CASE dim_ord(1)=1
                INDEX ON fio TAG T1
           CASE dim_ord(2)=1
                INDEX ON STR(np,3)+STR(nd,3)+fio TAG T1  
                DELETE FOR kp=0
        ENDCASE               
   CASE dim_opt(2)=1
        SELECT * FROM datKurs INTO CURSOR kursPrn READWRITE
        DELETE FOR perBeg<dateBeg.AND.perEnd<dateBeg
        DELETE FOR perBeg>dateEnd.AND.perEnd>dateEnd
        ALTER TABLE kursPrn ADD COLUMN npp N(6)
        ALTER TABLE kursPrn ADD COLUMN fio C(100)
        ALTER TABLE kursPrn ADD COLUMN kp N(3)
        ALTER TABLE kursPrn ADD COLUMN kd N(3)
        ALTER TABLE kursPrn ADD COLUMN np N(3)
        ALTER TABLE kursPrn ADD COLUMN nd N(3)
        ALTER TABLE kursPrn ADD COLUMN kat N(3)
        ALTER TABLE kursPrn ADD COLUMN ckval C(70)
        SELECT kursPrn   
        SET RELATION TO STR(nidpeop,5) INTO kursJob ADDITIVE
        REPLACE fio WITH IIF(SEEK(nidpeop,'people',4),people.fio,''),ckval WITH IIF(SEEK(people.kval,'sprkval',1),ALLTRIM(sprkval.name)+' ','')+IIF(!EMPTY(people.dkval),DTOC(people.dkval)+' ','')+ALLTRIM(people.nkval) ALL
      
        *REPLACE kp WITH IIF(SEEK(STR(nidpeop,5),'kursJob'),kursJob.kp,0),kd WITH kursJob.kd,kat WITH kursJob.kat ALL 
        REPLACE kp WITH kursJob.kp,kd WITH kursJob.kd,kat WITH kursJob.kat ALL 
        
        REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
        REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL        
        DELETE FOR EMPTY(fio)
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
           DELETE FOR !(','+LTRIM(STR(nType))+','$fltType)
        ENDIF  
        DO CASE
           CASE dim_ord(1)=1
                INDEX ON fio+DTOS(perBeg) TAG T1                
           CASE dim_ord(2)=1
                INDEX ON STR(np,3)+STR(nd,3)+fio TAG T1    
        ENDCASE    
ENDCASE
nppcx=1
SCAN ALL
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO repKursToExcel
********************************************************************************************************************************
PROCEDURE repKursToExcel
DO startPrnToExcel WITH 'fSupl'
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     maxColumn=10
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=40
     .Columns(4).ColumnWidth=40
     .Columns(5).ColumnWidth=12
     .Columns(6).ColumnWidth=12
     .Columns(7).ColumnWidth=45
     .Columns(8).ColumnWidth=45
     .Columns(9).ColumnWidth=6
     .Columns(10).ColumnWidth=20
     
     .cells(2,1).Value='№'  
     .cells(2,2).Value='ФИО сотрудника'  
     .cells(2,3).Value='Должность' 
     .cells(2,4).Value='Категория'
     .cells(2,5).Value='начало'
     .cells(2,6).Value='окончание'
     .cells(2,7).Value='наименование'
     .cells(2,8).Value='место прхождения'
     .cells(2,9).Value='часы'
     .cells(2,10).Value='вид подготовки'
                 
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter        
          .WrapText=.T.
          .Value='Список сотрудников, проходивших курсы повышения калификации'
          .Interior.ColorIndex=35
     ENDWITH                                       
     .Range(.Cells(2,1),.Cells(1,maxColumn)).Select                                                                    
     objExcel.Selection.HorizontalAlignment=xlCenter        
     numberRow=3
     yearMonth=''
     SELECT kursPrn
     DO storezeropercent
     kpcx=0
     nppcx=1
     nidpeopcx=0
     SCAN ALL              
          IF kpcx#kp.AND.dim_ord(2)=1
             kpcx=kp
             nppcx=1
             .Range(.Cells(numberRow,1),.Cells(numberRow,maxColumn)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlCenter
                  .WrapText=.T.
                  .Value=IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'')
                  .Interior.ColorIndex=35
              ENDWITH     
              numberRow=numberRow+1
          ENDIF
          IF dim_ord(1)=1
             .cells(numberRow,1).Value=IIF(nidpeopcx#nidpeop,nppcx,'')
             .cells(numberRow,2).Value=IIF(nidpeopcx#nidpeop,fio,'')
             .cells(numberRow,3).Value=IIF(nidpeopcx#nidpeop,IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),'')
             .cells(numberRow,4).Value=IIF(nidpeopcx#nidpeop,ckval,'')
              nppcx=IIF(nidpeopcx#nidpeop,nppcx+1,nppcx)  
             nidpeopcx=IIF(nidpeopcx#nidpeop,nidpeop,nidpeopcx)
          ELSE 
             .cells(numberRow,1).Value=nppcx
             .cells(numberRow,2).Value=fio
             .cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')
             .cells(numberRow,4).Value=ckval
              nppcx=nppcx+1  
          ENDIF                          
          .cells(numberRow,5).Value=IIF(!EMPTY(perBeg),perBeg,'')
          .cells(numberRow,6).Value=IIF(!EMPTY(perEnd),perEnd,'')
          .cells(numberRow,7).Value=nameKurs
          .cells(numberRow,8).Value=nameSchool
          .cells(numberRow,9).Value=IIF(kHours#0,kHours,'')
          .cells(numberRow,10).Value=IIF(SEEK(ntype,'doptype',1),doptype.name,'')
          numberRow=numberRow+1          
          DO fillpercent WITH 'fSupl'
     ENDSCAN  
     .Range(.Cells(1,1),.Cells(numberRow-1,maxColumn)).Select
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
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
ENDWITH 
ON ERROR DO erSup   
DO endPrnToExcel WITH 'fSupl'               
ON ERROR 
objExcel.Visible=.T. 