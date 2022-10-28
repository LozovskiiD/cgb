IF !USED('dataList')
   USE dataList ORDER 1 IN 0
ENDIF 
SELECT dataList
SET FILTER TO dataList.kodPeop=pCard.num
daysTot=0
SUM days TO daysTot
daysTot=ROUND(daysTot,0)
GO TOP
nrec=0
log_ap=.F.
newdBeg=CTOD('  .  .   ')
newdEnd=CTOD('  .  .   ')
newPrim=SPACE(50)
newDays=0
GO TOP
frmFam=CREATEOBJECT('FORMSUPL')
WITH frmFam
     .Caption='больничные листы'
     .procExit='DO exitFromList'
     
     DO addmenureadspr WITH 'frmFam','DO writeList WITH .T.','DO writeList WITH .F.'
     DO addcontmenu WITH 'frmFam','menucont1',10,5,'новая','pencila.ico','DO readList WITH .T.'
     DO addcontmenu WITH 'frmFam','menucont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico','DO readList WITH .F.'
     DO addcontmenu WITH 'frmFam','menucont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO formDelRecList'
     DO addcontmenu WITH 'frmFam','menucont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO exitFromList'  
     
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Left=0
          .Width=600
          .Height=300
          .ScrollBars=2          
          .ColumnCount=5        
          .RecordSourceType=1     
          .RecordSource='datalist'
          .Column1.ControlSource='dataList.dBeg'
          .Column2.ControlSource='dataList.dEnd'
          .Column3.ControlSource='datalist.days'
          .Column4.ControlSource='" "+dataList.primList'       
          .Column1.Width=RettxtWidth(' дата рожд. ')
          .Column2.Width=.column1.Width                       
          .Column4.Width=.Width-.column1.width-.column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount       
           .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='начало'
          .Column2.Header1.Caption='окончание'
          .Column3.Header1.Caption='дни'  
          .Column4.Header1.Caption='примечание'                 
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0           
          .Column3.Alignment=1         
          .Column4.Alignment=0
          .Column3.Format='Z'
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'frmFam','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(frmFam.fGrid.RecordSource)#frmfam.fGrid.curRec,frmFam.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(frmFam.fGrid.RecordSource)#frmFam.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR  
     
     DO addtxtboxmy WITH 'frmFam',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,0,'DO validDaysList'  
     DO addtxtboxmy WITH 'frmFam',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0,'DO validDaysList'
     DO addtxtboxmy WITH 'frmFam',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'frmFam',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')                     
     DO adtBoxNew WITH 'frmFam','tBox11',.fGrid.Top+.fGrid.Height-1,.fGrid.Left,.fGrid.Column1.Width+12,.fGrid.RowHeight,'',.F.,.F.,0 
     DO adtBoxNew WITH 'frmFam','tBox12',.tBox11.Top,.tBox11.Left+.tBox11.Width-1,.fGrid.Column2.Width+2,.tBox11.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'frmFam','tBox13',.tBox11.Top,.tBox12.Left+.tBox12.Width-1,.fGrid.Column3.Width+2,.tBox11.Height,'daysTot','Z',.F.,0
     DO adtBoxNew WITH 'frmFam','tBox14',.tBox11.Top,.tBox13.Left+.tBox13.Width-1,.fGrid.Column4.Width+2,.tBox11.Height,'',.F.,.F.,0           
     .Width=.fGrid.Width
     .Height=.menucont1.Height+.fgrid.Height+10+.tBox11.Height          
     
     DO pasteImage WITH 'frmFam'
     .Show
ENDWITH 
************************************************************************************************************************
PROCEDURE readList
PARAMETERS par1
SELECT dataList
IF par1     
   APPEND BLANK     
   REPLACE kodpeop WITH pCard.num
ENDIF
frmFam.Refresh
log_ap=IIF(par1,.T.,.F.)
newdBeg=IIF(par1,CTOD('  .  .    '),dBeg)
newdEnd=IIF(par1,CTOD('  .  .    '),dEnd)
newPrim=IIF(par1,SPACE(50),primList)
newDays=IIF(par1,0,days)
nrec=RECNO()
WITH frmFam
     .SetAll('Visible',.F.,'mymenucont')
     .menuread.Visible=.T.
     .menuexit.Visible=.T.

     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width-1
     .txtBox3.Left=.txtBox2.Left+.txtBox2.Width-1
     .txtBox4.Left=.txtBox3.Left+.txtBox3.Width-1
     .txtBox1.ControlSource='newdBeg'
     .txtbox2.ControlSource='newdEnd'
     .txtbox3.ControlSource='newDays'
     .txtbox4.ControlSource='newPrim'
     .txtBox1.Top=lineTop
     .txtBox2.Top=lineTop
     .txtBox3.Top=lineTop
     .txtBox4.Top=lineTop
     .txtBox1.Height=.fGrid.RowHeight+1
     .txtBox2.Height=.txtBox1.Height
     .txtBox3.Height=.txtBox1.Height
     .txtBox4.Height=.txtBox1.Height
     .SetAll('BackStyle',1,'MyTxtBox')     
     .SetAll('Visible',.T.,'MyTxtBox')                   
     .fGrid.Enabled=.F.
     .txtBox1.SetFocus
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE validDaysList
IF !EMPTY(newDBeg).AND.!EMPTY(newDEnd)
   newDays=newDEnd-newDBeg+1
   frmFam.txtBox3.Refresh
ENDIF
************************************************************************************************************************
PROCEDURE writeList
PARAMETERS par1
WITH frmFam
     .SetAll('Visible',.T.,'mymenucont')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     SELECT dataList
     IF par1
        REPLACE dBeg WITH newdBeg,dEnd WITH newdEnd,primList WITH newPrim,days WITH newDays
     ELSE
        IF log_ap
           DELETE
        ENDIF           
     ENDIF   
     SUM days TO daysTot
     daysTot=ROUND(daysTot,0)     
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'ComboMy') 
     .tBox11.Visible=.T.
     .tBox12.Visible=.T.
     .tBox13.Visible=.T.
     .tBox14.Visible=.T.        
     .fGrid.Enabled=.T.    
     GO nrec
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH  
frmFam.Refresh 
frmFam.fGrid.Columns(frmFam.fGrid.columnCount).SetFocus  
*************************************************************************************************************************
PROCEDURE formDelRecList
log_del=.F.
fDel=CREATEOBJECT('formsupl')
WITH fDel
     .Caption='Удаление записи'           
     DO addShape WITH 'fdel',1,20,20,dHeight,350,8   
     DO adLabMy WITH 'fdel',1,'Удалить выбранную запись?',.Shape1.Top+10,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fdel',2,'Для подтверждения намерений поставьте отметку',.Lab1.Top+.Lab1.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fdel',3,"в окошке 'подтверждение намерений'",.Lab2.Top+.Lab2.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     .Shape1.Height=.lab1.Height*3+30       
      DO adCheckBox WITH 'fDel','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'log_del',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wwудалитьww')*2-20)/2,.check1.Top+.check1.Height+20,;
       RetTxtWidth('wwудалитьww'),dHeight+3,'удалить','DO delRecList'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fdel.Release'
     .Height=.Shape1.Height+.cont1.Height+.check1.Height+80  
     .Width=.Shape1.Width+40 
ENDWITH
DO pasteImage WITH 'fDel'
fDel.Show
*************************************************************************************************************************
PROCEDURE delRecList
IF !log_del
   RETURN
ENDIF
fDel.Release
SELECT dataList
DELETE
SUM days TO daysTot
daysTot=ROUND(daysTot,0) 
GO TOP
frmFam.Refresh

*************************************************************************************************************************
PROCEDURE exitFromList
frmFam.Visible=.F.
SELECT dataList
USE
SELECT pCard
frmFam.Release