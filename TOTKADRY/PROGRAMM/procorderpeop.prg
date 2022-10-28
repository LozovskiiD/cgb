PARAMETERS parUv
IF USED('curListOrder')
   SELECT curListOrder
   USE
ENDIF
IF !parUv
   SELECT * FROM peoporder WHERE nidpeop=people.nid INTO CURSOR curListOrder READWRITE
   INSERT INTO curListOrder SELECT * FROM pordarc WHERE nidpeop=people.nid
ELSE 
   SELECT * FROM peoporder WHERE nidpeop=peopout.nid INTO CURSOR curListOrder READWRITE
   INSERT INTO curListOrder SELECT * FROM pordarc WHERE nidpeop=peopout.nid

ENDIF    
*---------------------------------------------------------
*--------------------------------------------------------
ALTER TABLE curListOrder ADD COLUMN strOrd C(1)
ALTER TABLE curListOrder ADD COLUMN nameOrd C(80)

*SELECT * FROM pordarc WHERE nidpeop=people.nid INTO CURSOR curlistOrder2
*SELECT curListOrder
*APPEND FROM DBF('curlistorder2')
*-----
*APPEND  FROM pordarc FOR nidpeop=people.nid


REPLACE strOrd WITH IIF(SEEK(kord,'datorder',1),datorder.strOrd,strOrd) ALL
REPLACE nameOrd WITH IIF(SEEK(supOrd,'sprorder',1),sprorder.nameord,nameord) ALL
*INDEX ON dateOrd TAG T1 DESCENDING 
INDEX ON dOrd TAG T1 DESCENDING 
GO TOP
WITH oPageOrd
*     .Caption='приказы'  
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=nParent.Width
 *         .Height=.Parent.menuCont1.Top-60
          .ScrollBars=2          
          .ColumnCount=6        
          .RecordSourceType=1     
          .RecordSource='curListOrder'
          .Column1.ControlSource='curListOrder.dord'
          .Column2.ControlSource='curListOrder.nOrd'
          .Column3.ControlSource='curListOrder.strOrd'
          .Column4.ControlSource='curListOrder.npp'       
          .Column5.ControlSource='" "+curListOrder.nameOrd'       
          .Column1.Width=RettxtWidth('99/99/99999')
          .Column2.Width=RetTxtWidth('номерw')                       
          .Column3.Width=RetTxtWidth('wтип')
          .Column4.Width=RetTxtWidth('п.w')
          .Column5.Width=.Width-.column1.width-.column2.Width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount       
           .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='дата'
          .Column2.Header1.Caption='номер'
          .Column3.Header1.Caption='тип'  
          .Column4.Header1.Caption='п.'                 
          .Column5.Header1.Caption='содержание'                 
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0           
          .Column3.Alignment=2         
          .Column4.Alignment=2                   
          .Column5.Alignment=0
          .Column2.Format='Z'
          .Column4.Format='Z'
          .colNesInf=2      
          .procAfterRowColChange='DO changeRowOrd'
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'oPageOrd','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(oPageOrd.fGrid.RecordSource)#oPageOrd.fGrid.curRec,oPageOrd.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(oPageOrd.fGrid.RecordSource)#oPageOrd.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR 
     .AddObject('editOrder','MyEditBox')      
     WITH .editOrder
          .Visible=.T.          
          .ControlSource='curListOrder.txtOrd'
          .Left=.Parent.fGrid.Left
          .Width=.Parent.fGrid.Width
          .Top=.Parent.fGrid.Top+.Parent.fGrid.Height+5
          .Height=100
          .Enabled=.T.  
          .ReadOnly=.T.
     ENDWITH 
     IF !parUv
        DO addButtonOne WITH 'oPageOrd','butRead',.editOrder.Left+(.editOrder.Width-RetTxtWidth('wредакцияw')*2-10)/2,.editOrder.Top+.editOrder.Height+20,'редакция','','DO readOrder',39,RetTxtWidth('wредакцияw'),'редакция'
        DO addButtonOne WITH 'oPageOrd','butDel',.butRead.Left+.butRead.Width+10,.butRead.Top,'удаление','','DO delOrder',39,.butRead.Width,'удаление'
        DO addButtonOne WITH 'oPageOrd','butRetRead',.editOrder.Left+(.editOrder.Width-.butRead.Width)/2,.butRead.Top,'возврат','','DO retReadOrder',39,.butRead.Width,'возврат'
        .butRetRead.Visible=.F.
   
        DO addButtonOne WITH 'oPageOrd','butDelRec',.editOrder.Left+(.editOrder.Width-.butRead.Width*2-10)/2,.butRead.Top,'удалить','','DO delRecOrder WITH .T.',39,.butRead.Width,'удалить'
        DO addButtonOne WITH 'oPageOrd','butRetDel',.butDelRec.Left+.butDelRec.Width+10,.butRead.Top,'возврат','','DO delRecOrder WITH .F.',39,.butRead.Width,'вовзрат'
        .butDelRec.Visible=.F.
        .butRetDel.Visible=.F.
    ENDIF     
    .Refresh
ENDWITH 
*****************************************************************
PROCEDURE changeRowOrd
oPageOrd.editOrder.ControlSource='curListOrder.txtOrd'
*****************************************************************
PROCEDURE readOrder
WITH oPageOrd
     .butRead.Visible=.F.
     .butDel.Visible=.F.
     .butRetRead.Visible=.T.
     .editOrder.ReadOnly=.F.
ENDWITH 

*****************************************************************
PROCEDURE retReadOrder
WITH oPageOrd
     .butRead.Visible=.T.
     .butDel.Visible=.T.
     .butRetRead.Visible=.F.
     .editOrder.ReadOnly=.T.
ENDWITH 

*****************************************************************
PROCEDURE delOrder
WITH oPageOrd
     .butRead.Visible=.F.
     .butDel.Visible=.F.
     .butDelRec.Visible=.T.
     .butRetDel.Visible=.T.
ENDWITH 
*****************************************************************
PROCEDURE delRecOrder
PARAMETERS par1
IF par1
   IF SEEK(curListOrder.nid,'peoporder',2)
      SELECT peoporder
      DELETE
   ENDIF
   SELECT curListOrder
   DELETE
   GO TOP 
ENDIF
WITH oPageOrd
     .butRead.Visible=.T.
     .butDel.Visible=.T.
     .butDelRec.Visible=.F.
     .butRetDel.Visible=.F.
     .Refresh
ENDWITH 