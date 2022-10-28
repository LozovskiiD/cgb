IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
CREATE CURSOR curPrn (npp N(2),kod N(3),nfio C(50),kp N(3),kd N(3),ndolj C(100),perBeg D,perEnd D,begOtp D,endotp D,days N(3),nameotp C(40))
STORE CTOD('  .  .    ') TO dBeg,dEnd

fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Отчет по отпускам'
     .Width=400
     DO addshape WITH 'fSupl',1,10,10,150,400,8 
     DO adlabMy WITH 'fSupl',1,'Период с',.Shape1.Left+20,.Shape1.Top+20,100,0,.T.  
     DO adtbox WITH 'fSupl',1,.lab1.Left+.lab1.Width+20,.lab1.Top,RetTxtWidth('99/99/99999'),dHeight,'dBeg',.F.,.T.,.F.
     DO adlabMy WITH 'fSupl',2,'по',20,20,100,0,.T.  
     DO adtbox WITH 'fSupl',2,.lab1.Left+.lab1.Width+20,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'dEnd',.F.,.T.,.F.
     
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.lab2.Width-.txtBox1.Width*2-30)/2        
     .txtBox1.Left=.lab1.Left+.lab1.Width+10
     .lab2.Left=.txtBox1.Left+.txtBox1.Width+10
     .txtBox2.Left=.lab2.Left+.lab2.Width+10
     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+5)
     .lab2.Top=.lab1.Top      
     .Shape1.Height=.txtBox1.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.F.,.T.         
     
     DO adLabMy WITH 'fSupl',24,'ход выполнения',.Shape91.Top+.Shape91.Height+5,.Shape1.Left,.Shape1.Width,2,.F.,1  
     .lab24.Visible=.F.
     DO addShape WITH 'fSupl',11,.Shape1.Left,.lab24.Top+.lab24.Height,dHeight,.Shape1.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,dHeight,0,8
     .Shape12.BackStyle=1
     .Shape12.Visible=.F.      
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+2,.lab24.Left,.lab24.Width,2,.F.,0
     .lab25.Visible=.F.      
     
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WПросмотрW')*3-20)/2,.Shape91.Top+.Shape91.Height+20,RetTxtWidth('WПросмотрW'),dHeight+5,'печать','DO procreportotp WITH 1' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+5,'просмотр','DO procreportotp WITH 2'
     *---------------------------------Кнопка возврат-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+10,.Cont1.Top,.Cont1.Width,dHeight+5,'возврат','fSupl.Release'
     .Width=.Shape1.Width+20
     
     .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+60    
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE procreportotp
PARAMETERS par1
SELECT datotp
SET FILTER TO 
SELECT * FROM pcard WHERE !otmuv INTO CURSOR curCardOtp 
SELECT curCardOtp
SCAN ALL
     SELECT datOtp
     SEEK STR(curCardOtp.num,4)
     IF FOUND()
        DO WHILE kodpeop=curCardOtp.num
           SELECT curPrn
           APPEND BLANK
           REPLACE kod WITH curcardOtp.num,nFio WITH curCardOtp.fio,perBeg WITH datOtp.perBeg,perEnd WITH datOtp.perEnd,begOtp WITH datOtp.begOtp,endOtp WITH datOtp.endOtp,days WITH datOtp.kvoDay,nameotp WITH datotp.nameotp
           SELECT datJob
           LOCATE FOR kodpeop=curPrn.kod.AND.tr=1.AND.EMPTY(dateout)
           SELECT curPrn
           REPLACE kp WITH datJob.kp,kd WITH datJob.kd
           SELECT datOtp
           SKIP
        ENDDO 
     ENDIF
     SELECT curCardOtp
ENDSCAN
SELECT curprn
DO CASE
   CASE !EMPTY(dBeg).AND.EMPTY(dEnd)
        DELETE FOR begotp>dBeg    
   CASE EMPTY(dBeg).AND.!EMPTY(dEnd)
        DELETE FOR begotp<dEnd
   CASE dEnd>dBeg
        DELETE FOR begotp<dBeg
        DELETE FOR begotp>dEnd        
ENDCASE
REPLACE ndolj WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL
INDEX ON STR(kp,3)+nfio+DTOS(begOtp) TAG T1
SET ORDER TO 1
nppcx=1
numOld=0
SCAN ALL
     IF numOld#kod
        nppcx=1 
        numOld=kod      
     ENDIF
     REPLACE npp WITH nppcx          
     nppcx=nppcx+1
ENDSCAN
DO CASE
   CASE par1=1
        DO procForPrintAndPreview WITH 'repotp','Список сотрудников',.T.,'repotpexcel'              
   CASE par1=2
        DO procForPrintAndPreview WITH 'repotp','Список сотрудников',.F.           
ENDCASE
*******************************************************************************
PROCEDURE repotpexcel
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F. 
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH  
#DEFINE xlCenter -4108            
*#DEFINE xlLeft -4131  
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
     .Columns(1).ColumnWidth=40
     .Columns(2).ColumnWidth=40
     .Columns(3).ColumnWidth=40
     .Columns(4).ColumnWidth=12
     .Columns(5).ColumnWidth=12
     .Columns(6).ColumnWidth=12
     .Columns(7).ColumnWidth=12
     .Columns(8).ColumnWidth=7
     
     .cells(2,1).Value='ФИО сотрудника'              
     .cells(2,2).Value='должность'
     .cells(2,3).Value='отпуск'
     .cells(2,4).Value='период с'                                    
     .cells(2,5).Value='период по'                                    
     .cells(2,6).Value='начало'
     .cells(2,7).Value='оконч.'                                    
     .cells(2,8).Value='дней'    
     numberRow=3 
     SELECT curPrn
     STORE 0 TO max_rec,one_pers,pers_ch
     COUNT TO max_rec
     kpOld=0
     SCAN ALL             
          IF kp#kpOld
             .Range(.Cells(numberRow,1),.Cells(numberRow,8)).Select
             WITH objExcel.Selection
                  .MergeCells=.T.
                  .Value=IIF(SEEK(kp,'sprpodr',1),sprpodr.name,'')
                  .Interior.ColorIndex=37
                  kpOld=kp
                  .WrapText=.T.
             ENDWITH
             numberRow=numberRow+1         
          ENDIF
          .cells(numberRow,1).Value=IIF(npp=1,nfio,'')
          .cells(numberRow,2).Value=IIF(npp=1,ndolj,'')
          .cells(numberRow,3).Value=nameotp
          .cells(numberRow,4).Value=IIF(!EMPTY(perBeg),DTOC(perBeg),'')
          .cells(numberRow,5).Value=IIF(!EMPTY(perEnd),DTOC(perEnd),'')
          .cells(numberRow,6).Value=IIF(!EMPTY(begOtp),DTOC(begOtp),'')
          .cells(numberRow,7).Value=IIF(!EMPTY(endOtp),DTOC(endOtp),'')
          .cells(numberRow,8).Value=days
          
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch    
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,8)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'
         .Font.Size=10
    ENDWITH      
    .Range(.Cells(1,1),.Cells(1,3)).Select
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
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T. 
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH              
*ON ERROR
objExcel.Visible=.T. 
*********************************************
PROCEDURE erSup