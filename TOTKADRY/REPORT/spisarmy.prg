*log_term=.T.
DIMENSION dim_opt(3)
dim_opt(1)=1
dim_opt(2)=0
dim_opt(2)=0
IF !USED('sprtot')
   USE sprtot ORDER 1 IN 0
ENDIF 
IF !USED('boss')
   USE boss IN 0
ENDIF
SELECT kod,name,namesp FROM sprtot WHERE sprtot.kspr=10 INTO CURSOR curSupGrup READWRITE &&группы воинского учёта
SELECT curSupGrup
APPEND BLANK
REPLACE name WITH '- все -'
cStrGrup=curSupGrup.name
nKodGrup=curSupGrup.kod
INDEX ON kod TAG T1 
logWord=.F.
term_ch=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl         
     .Caption='Список военнообязанных'     
     DO addshape WITH 'fSupl',1,10,10,150,400,8     
     DO adtBoxAsCont WITH 'fSupl','contGru',.Shape1.Top+10,.Shape1.Left+10,RetTxtWidth('Wгруппв учетаW'),dHeight,'группа учёта',0,1   
     DO addComboMy WITH 'fSupl',1,.contGru.Left+.contGru.Width-1,.contGru.Top,dHeight,250,.T.,'cStrGrup','ALLTRIM(curSupGrup.name)',6,'DO focusSpisGrup','nKodGrup=curSupGrup.kod',.F.,.T.     
     .Shape1.Height=.contGru.Height+20
     .Shape1.Width=.contGru.Width+.comboBox1.Width+20
     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8     
     DO addOptionButton WITH 'fSupl',1,'форма 1',.Shape2.Top+10,.Shape2.Left+20,'dim_opt(1)',0,"DO procValOption WITH 'fSupl','dim_opt',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'форма 2',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_opt(2)',0,"DO procValOption WITH 'fSupl','dim_opt',2",.T. 
     DO addOptionButton WITH 'fSupl',3,'форма 3',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_opt(3)',0,"DO procValOption WITH 'fSupl','dim_opt',3",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-.Option3.Width-40)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20 
     .Option3.Left=.Option2.Left+.Option2.Width+20 
     .Shape2.Height=.Option1.Height+20
     
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape2.Top+.Shape2.Height+10,.Shape1.Width,.F.,.T. 
     DO adButtonPrnToForm WITH 'DO prnSpisArmy WITH .T.','DO prnSpisArmy WITH .F.','fsupl.Release'
     .Width=.Shape1.Width+20
     .Height=.Shape91.Height+.Shape1.Height+.Shape2.Height+.butPrn.Height+70
     
     DO addShape WITH 'fSupl',11,.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape91.Left,.Shape91.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.       
     
     
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************************
PROCEDURE focusSpisGrup
****************************************************************************************************************************
PROCEDURE prnSpisArmy
PARAMETERS parLog

IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curJobAge')
   SELECT curJobAge
   USE
ENDIF
IF !USED('datarmy')
   USE datarmy IN 0
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 
SELECT * FROM people WHERE SEEK(nid,'datarmy',2).AND.EMPTY(datarmy.datesn) INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN kp N(3)
ALTER TABLE curPrn ADD COLUMN namep C(100)
ALTER TABLE curPrn ADD COLUMN kd N(3)
ALTER TABLE curPrn ADD COLUMN named C(100)
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN zv C(50)
ALTER TABLE curPrn ADD COLUMN grupu N(1)
ALTER TABLE curPrn ADD COLUMN rik C(30)
ALTER TABLE curPrn ADD COLUMN profilvus C(30)
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,namep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') all
REPLACE zv WITH IIF(SEEK(nid,'datarmy',2),datarmy.zv,''),grupu WITH datarmy.grupu,rik WITH datarmy.rik profilvus WITH ALLTRIM(ALLTRIM(datarmy.profil)+' '+ALLTRIM(datarmy.vus)) ALL
IF nKodGrup#0
   DELETE FOR grupu#nKodGrup
ENDIF
INDEX ON fio TAG T1
nppcx=1
SCAN ALL
     SELECT curPrn      
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO CASE
   CASE dim_opt(1)=1 
        DO procForPrintAndPreview WITH 'repspisarmy','',parLog,'repspisArmyToExcel'
   CASE dim_opt(2)=1     
        DO procForPrintAndPreview WITH 'repspisarmy','',parLog,'repspisArmyToExcel2'
   CASE dim_opt(3)=1     
        DO procForPrintAndPreview WITH 'repspisarmy','',parLog,'repspisArmyToExcel3'     
ENDCASE 
********************************************************************************************************************************
PROCEDURE repSpisArmyToExcel
ON ERROR DO erSup
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
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
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=20
     .Columns(4).ColumnWidth=50
     .Columns(5).ColumnWidth=50     
     .Columns(6).ColumnWidth=20     
     .cells(2,1).Value='№'                                     
     .cells(2,2).Value='ФИО'              
     .cells(2,3).Value='звание'
     .cells(2,4).Value='военкомат'    
     .cells(2,5).Value='адрес'
     .cells(2,6).Value='телефон'    
     .Range(.Cells(1,1),.Cells(1,6)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='Список военнообязанных'          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,6)).Select
     objExcel.Selection.HorizontalAlignment= -4108         

     numberRow=3
     SELECT curPrn
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     SCAN ALL        
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=ALLTRIM(fio)
          .cells(numberRow,3).Value=ALLTRIM(zv)
          .cells(numberRow,4).Value=ALLTRIM(rik)
          .cells(numberRow,5).Value=ALLTRIM(ppreb)
          .cells(numberRow,6).Value=ALLTRIM(telhome)
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch 
          numberRow=numberRow+1         
     ENDSCAN
     .Range(.Cells(1,1),.Cells(numberRow-1,6)).Select
     WITH objExcel.Selection
          .Borders(xlEdgeLeft).Weight=xlThin
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin
          .VerticalAlignment=1   
          .HorizontalAlignment=xlLeft
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=10
     ENDWITH 
     .Range(.Cells(1,1),.Cells(2,6)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     .Range(.Cells(1,1),.Cells(1,6)).Select
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
 WITH fSupl
      .SetAll('Visible',.T.,'myCommandButton')    
      .SetAll('Visible',.T.,'myContLabel')    
      .Shape11.Visible=.F.
      .Shape12.Visible=.F.      
      .lab25.Visible=.F.      
ENDWITH   
ON ERROR            
objExcel.Visible=.T. 
********************************************************************************************************************************
PROCEDURE repSpisArmyToExcel2
ON ERROR DO erSup
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
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
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=30
     .Columns(2).ColumnWidth=6
     .Columns(3).ColumnWidth=30
     .Columns(4).ColumnWidth=15
     .Columns(5).ColumnWidth=15    
     .Columns(6).ColumnWidth=15     
     .Columns(7).ColumnWidth=40     
     .Columns(8).ColumnWidth=15     
     .Columns(9).ColumnWidth=15     
      maxColumn=9  
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='Список'          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .WrapText=.T.
          .Value='военнообязанных, работающих '          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     .Range(.Cells(3,1),.Cells(3,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='в '+ALLTRIM(boss.office)
     ENDWITH  
     .Range(.Cells(3,1),.Cells(3,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter        
          
     .cells(4,1).Value='Фамилия, Имя Отчество'                                     
     .cells(4,2).Value='Год рождения'              
     .cells(4,3).Value='Состав (профиль) и номер ВУС'
     .cells(4,4).Value='Воинское звание'    
     .cells(4,5).Value='Годность к военной службе'
     .cells(4,6).Value='Приписан или не приписан к воиеским частям или командам'        
     .cells(4,7).Value='Занимаемая должность и наименование подразделения организации'    
     .cells(4,8).Value='Решение военного комиссара'    
     .cells(4,9).Value='Военкомат'    
             
     numberRow=5
     SELECT curPrn
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     SCAN ALL        
          .cells(numberRow,1).Value=ALLTRIM(fio)
          .cells(numberRow,2).Value=IIF(!EMPTY(age),STR(YEAR(age),4),'')
          .cells(numberRow,3).Value=ALLTRIM(profilvus)
          .cells(numberRow,4).Value=ALLTRIM(zv)      
          .cells(numberRow,7).Value=ALLTRIM(named)+' '+ALLTRIM(namep)
          .cells(numberRow,9).Value=ALLTRIM(rik)
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch
          numberRow=numberRow+1         
     ENDSCAN
     .Range(.Cells(4,1),.Cells(numberRow-1,maxColumn)).Select
     WITH objExcel.Selection
          .Borders(xlEdgeLeft).Weight=xlThin
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin
          .VerticalAlignment=1   
          .HorizontalAlignment=xlLeft                   
     ENDWITH 
     
     .Range(.Cells(1,1),.Cells(numberRow-1,maxColumn)).Select
     WITH objExcel.Selection                 
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=10
     ENDWITH 
     
     .Range(.Cells(4,1),.Cells(4,maxColumn)).Select
     WITH objExcel.Selection          
          .HorizontalAlignment=xlCenter
     ENDWITH 
     
     .Range(.Cells(1,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
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
 WITH fSupl
      .SetAll('Visible',.T.,'myCommandButton')    
      .SetAll('Visible',.T.,'myContLabel')    
      .Shape11.Visible=.F.
      .Shape12.Visible=.F.      
      .lab25.Visible=.F.      
ENDWITH     
ON ERROR            
objExcel.Visible=.T. 
********************************************************************************************************************************
PROCEDURE repSpisArmyToExcel3
ON ERROR DO erSup
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
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
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=30
     .Columns(2).ColumnWidth=6
     .Columns(3).ColumnWidth=30
     .Columns(4).ColumnWidth=15
     .Columns(5).ColumnWidth=15    
     .Columns(6).ColumnWidth=40     
     .Columns(7).ColumnWidth=15     
     .Columns(8).ColumnWidth=15     
     .Columns(9).ColumnWidth=15     
      maxColumn=9  
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='Список'          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .WrapText=.T.
          .Value='военнообязанных, работающих '          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     .Range(.Cells(3,1),.Cells(3,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='в '+ALLTRIM(boss.office)
     ENDWITH  
     .Range(.Cells(3,1),.Cells(3,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     .Range(.Cells(4,1),.Cells(4,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='на которых необходимо оформить отсрочки от призыва на военную службу по'
     ENDWITH  
     .Range(.Cells(4,1),.Cells(4,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
     .Range(.Cells(5,1),.Cells(5,maxColumn)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='по мобилизации и в военное время в соответствии с Единым перечнем'
     ENDWITH  
     .Range(.Cells(5,1),.Cells(5,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     
          
     .cells(6,1).Value='Фамилия, Имя Отчество'                                     
     .cells(6,2).Value='Год рождения'              
     .cells(6,3).Value='Состав (профиль) и номер ВУС'
     .cells(6,4).Value='Воинское звание'    
     .cells(6,5).Value='Годность к военной службе'
     .cells(6,6).Value='Занимаемая должность'    
     .cells(6,7).Value='Номер части, раздела и пункта Единого перечня, по которым военнообязанный подлежит бронированию'        
     .cells(6,8).Value='Примечание'    
     .cells(6,9).Value='Военкомат'   
             
     numberRow=7
     SELECT curPrn
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     SCAN ALL        
          .cells(numberRow,1).Value=ALLTRIM(fio)
          .cells(numberRow,2).Value=IIF(!EMPTY(age),STR(YEAR(age),4),'')
          .cells(numberRow,3).Value=ALLTRIM(profilvus)
          .cells(numberRow,4).Value=ALLTRIM(zv)      
          .cells(numberRow,6).Value=ALLTRIM(named)+' '+ALLTRIM(namep)
          .cells(numberRow,9).Value=ALLTRIM(rik)
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch
          numberRow=numberRow+1         
     ENDSCAN
     .Range(.Cells(6,1),.Cells(numberRow-1,maxColumn)).Select
     WITH objExcel.Selection
          .Borders(xlEdgeLeft).Weight=xlThin
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin
          .VerticalAlignment=1   
          .HorizontalAlignment=xlLeft                   
     ENDWITH 
     
     .Range(.Cells(1,1),.Cells(numberRow-1,maxColumn)).Select
     WITH objExcel.Selection                 
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=10
     ENDWITH 
     
     .Range(.Cells(6,1),.Cells(6,maxColumn)).Select
     WITH objExcel.Selection          
          .HorizontalAlignment=xlCenter
     ENDWITH 
     
     .Range(.Cells(1,1),.Cells(2,maxColumn)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
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
 WITH fSupl
      .SetAll('Visible',.T.,'myCommandButton')    
      .SetAll('Visible',.T.,'myContLabel')    
      .Shape11.Visible=.F.
      .Shape12.Visible=.F.      
      .lab25.Visible=.F.      
ENDWITH     
ON ERROR            
objExcel.Visible=.T. 