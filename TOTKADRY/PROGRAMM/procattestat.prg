IF !USED('dataat')
   USE dataat IN 0 ORDER 1
ENDIF 
CREATE CURSOR curSprAt (kod N(1),nameat C(30))
SELECT curSprAt
APPEND BLANK
REPLACE kod WITH 1,nameat WITH 'соответствует'
APPEND BLANK
REPLACE kod WITH 2,nameat WITH 'не полностью соответствует'
APPEND BLANK
REPLACE kod WITH 3,nameat WITH 'не соответствует'
newDateAt=CTOD('  .  .    ')
newDateNext=CTOD('  .  .    ')
newPrikaz=''
newItog=0
newStrItog=''
newSrok=0
nrec=0
logAp=.F.
SELECT dataat
SET FILTER TO dataat.kodPeop=pCard.num
GO TOP
frmAt=CREATEOBJECT('FORMSUPL')
WITH frmAt
     .Caption='аттестация'
     .procExit='DO exitFromProcAttestat'
      
     DO addmenureadspr WITH 'frmAt','DO writeAt WITH .T.','DO writeAt WITH .F.'
     DO addcontmenu WITH 'frmAt','menucont1',10,5,'новая','pencila.ico','DO readAt WITH .T.'
     DO addcontmenu WITH 'frmAt','menucont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico','DO readAt WITH .F.'
     DO addcontmenu WITH 'frmAt','menucont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico',"frmAt.fGrid.GridDelRec('frmAt.fGrid','dataat')"
     DO addcontmenu WITH 'frmAt','menucont4',.menucont3.Left+.menucont3.Width+3,5,'выход','undo.ico','DO exitFromProcAttestat'
      .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Left=0
          .Width=650
          .Height=300
          .ScrollBars=2          
          .ColumnCount=6        
          .RecordSourceType=1     
          .RecordSource='dataat'
          .Column1.ControlSource='dataat.dateAt'
          .Column2.ControlSource='dataat.prikaz'
          .Column3.ControlSource='dataat.stritog'
          .Column4.ControlSource='dataat.srok' 
          .Column5.ControlSource='dataat.datenext'        
           
          .Column1.Width=RettxtWidth(' дата рожд. ')
          .Column2.Width=RetTxtWidth('Wприказ №1999 от 15.12.2222W')
          .Column4.Width=RetTxtWidth('срокw')
          .Column5.Width=.column1.Width 
          .Columns(.ColumnCount).Width=0                     
          .Column3.Width=.Width-.column1.width-.column2.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount       
          
          .Column1.Header1.Caption='дата'
          .Column2.Header1.Caption='приказ'
          .Column3.Header1.Caption='итог'  
          .Column4.Header1.Caption='срок'                 
          .Column5.Header1.Caption='след.'                 
          .Column1.Movable=.F. 
          .Column2.Alignment=0           
          .Column3.Alignment=0        
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'frmAt','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(frmAt.fGrid.RecordSource)#frmAt.fGrid.curRec,frmAt.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(frmAt.fGrid.RecordSource)#frmAt.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR    
     .Width=.fGrid.Width
     .Height=.menucont1.Height+.fgrid.Height+10          
     DO addtxtboxmy WITH 'frmAt',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'frmAt',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     DO addComboMy WITH 'frmAt',1,1,1,frmAt.fGrid.rowHeight+1,frmAt.fGrid.Column3.Width+2,.T.,'strItog','curSprAt.nameat',6,'DO gotFocusAt','DO validAt'
     .comboBox1.Visible=.F.
     
     DO addtxtboxmy WITH 'frmAt',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,1,'DO validSrok'
     DO addtxtboxmy WITH 'frmAt',5,1,1,.fGrid.Column5.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')                     
       
     .Height=.menucont1.Height+.fGrid.Height+10
     .Width=.fGrid.Width
ENDWITH
DO pasteImage WITH 'frmAt'
frmAt.Show
****************************************************************************
PROCEDURE exitFromProcAttestat
frmAt.Visible=.F.
*SELECT curSprAt
*USE
SELECT dataat
USE
SELECT pCard
frmAt.Release
***************************************************************************
PROCEDURE readAt
PARAMETERS par1
SELECT dataAt
IF par1     
   APPEND BLANK     
   REPLACE kodpeop WITH pCard.num
ENDIF
frmAt.Refresh
logAp=IIF(par1,.T.,.F.)
newDateAt=IIF(par1,CTOD('  .  .    '),dateAt)
newDateNext=IIF(par1,CTOD('  .  .    '),dateNext)
newPrikaz=IIF(par1,SPACE(20),prikaz)
newSrok=IIF(par1,3,srok)
newItog=IIF(par1,0,itog)
newStrItog=IIF(par1,'',strItog)
nrec=RECNO()
WITH frmAt
     .SetAll('Visible',.F.,'mymenucont')
     .menuread.Visible=.T.
     .menuexit.Visible=.T.
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width-1
     .comboBox1.Left=.txtBox2.Left+.txtBox2.Width-1
     .txtBox4.Left=.comboBox1.Left+.comboBox1.Width-1
     .txtBox5.Left=.txtBox4.Left+.txtBox4.Width-1
     .txtBox1.ControlSource='newDateAt'
     .txtbox2.ControlSource='newPrikaz'
     .comboBox1.ControlSource='newStrItog'
     .txtbox4.ControlSource='newSrok'
     .txtbox5.ControlSource='newDateNext'
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Top',linetop,'comboMy')
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')     
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'comboMy')
     .fGrid.Enabled=.F.
     .txtBox1.SetFocus
ENDWITH 
***************************************************************************
PROCEDURE gotFocusAt
***************************************************************************
PROCEDURE validAt
newItog=curSprAt.kod
newStrItog=curSprAt.nameat
KEYBOARD '{TAB}'
***************************************************************************
PROCEDURE validSrok
newDateNext=CTOD(STR(DAY(newDateAt),2)+'.'+STR(MONTH(newDateAt),2)+'.'+STR(YEAR(newDateAt)+newSrok,4))
frmAt.txtBox5.Refresh
***************************************************************************
PROCEDURE writeAt
PARAMETERS par1
WITH frmAt
     .SetAll('Visible',.T.,'mymenucont')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     SELECT dataAt
     IF par1
        REPLACE dateAt WITH newDateAt,prikaz WITH newPrikaz,itog WITH newItog,srok WITH newSrok,dateNext WITH newDateNext,strItog WITH newStrItog
     ELSE
        IF logAp
           DELETE
        ENDIF           
     ENDIF   
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'ComboMy')
     .SetAll('Visible',.F.,'ComboMy')     
     .fGrid.Enabled=.T.    
     GO nrec
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH  
frmAt.Refresh 
frmAt.fGrid.Columns(frmAt.fGrid.columnCount).SetFocus  