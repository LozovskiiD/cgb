**************************************************************************************************************************
*                             Основная для групп печати
**************************************************************************************************************************
IF USED('datagrup')
   SELECT datagrup
   USE
ENDIF
USE datagrup IN 0
fgrup=CREATEOBJECT('FORMSPR')
WITH fgrup 
     .Caption='Группы печати' 
     .AddProperty('Grupname','')
     .AddProperty('Grupold',0)
     .procexit='Do returngrup'   
     
     DO addButtonOne WITH 'fGrup','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fgrup','Do readgrup WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fGrup','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fgrup','Do readgrup WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fGrup','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico',"DO createFormNew WITH .T.,'Удаление',RetTxtWidth('WWУдалить группу?WW',dFontName,dFontSize+1),;
           '130',RetTxtWidth('WНетW',dFontName,dFontSize+1),'Да','Нет',.F.,'DO delgrup','nFormMes.Release',.F.,'Удалить группу?',.F.,.T." ,39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fGrup','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','Do returngrup',39,.menucont1.Width,'возврат'             
     DO addmenureadspr WITH 'fgrup','DO writegrup WITH .T.','DO writegrup WITH .F.'
     
     SELECT rasp  
     REPLACE log_gr WITH .F. ALL
     SELECT datagrup
     GO TOP
*--------------------------Grid для справочника групп-------------------------------------------------------------
     WITH .fgrid
          .Left=0
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+10
          .Width=.Parent.Width
          .Height=(.Parent.height-.Top)/2-50                 
          .RecordSource='datagrup'   
          DO addColumnToGrid WITH 'fGrup.fGrid',4
          .Column1.Header1.Caption='Группа'
          .Column2.Header1.Caption='Учреждение'
          .Column3.Header1.Caption='Адрес'
          .Column1.Width=.Width/4 
          .Column2.Width=(.Width-.column1.Width)/2
          .Column3.Width=.Width-.Column1.Width-.column2.Width-SYSMETRIC(5)-13-4
          .Column4.Width=0
          .Column1.ControlSource='" "+datagrup->name'
          .Column2.ControlSource='" "+datagrup.firma'
          .Column3.ControlSource='" "+datagrup->adres'
          .SetAll('BOUND',.F.,'Column')
          .Column1.Sparse=.T.       
          .procAfterRowColChange='DO procotmgrup'  
     ENDWITH
     DO addtxtboxmy WITH 'fgrup',1,1,1,.fgrid.Column1.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fgrup',2,1,1,.fgrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fgrup',3,1,1,.fgrid.Column3.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')
     DO MyColumntxtBox WITH 'fgrup.fgrid.Column1','tbox1','datagrup->name'
     DO MyColumntxtBox WITH 'fgrup.fgrid.Column2','tbox2','datagrup->firma'
     DO MyColumntxtBox WITH 'fgrup.fgrid.Column3','tbox3','datagrup->adres'
     DO gridSizeNew WITH 'fgrup','fgrid','shapeingrid' 
*----------------------------------------------наименование группы--------------------------------------------------------
     DO adLabMy WITH 'fgrup',1,ALLTRIM(datagrup.name),.fgrid.Top+.fgrid.Height+12,0,fgrup.Width,2  
     WITH .Lab1
          .FontSize=dFontSize+1
          .FontBold=.T.
          .Alignment=2 
     ENDWITH     
*-------------------------------------------------------------------------------------------------------------------------
     SELECT * FROM sprpodr WHERE SEEK(STR(kod,3),'rasp',1) INTO CURSOR curpodr READWRITE
     ALTER TABLE curpodr ADD COLUMN log_gr L
     SELECT curpodr
     INDEX ON np TAG t1
     SELECT datagrup
     LOCATE FOR name=fgrup.grupName
     SELECT curpodr
     REPLACE log_gr WITH .T. FOR ','+LTRIM(STR(kod))+','$datagrup->sostav1
     GO TOP
*--------------------------------Grid для состава группы-------------------------------------------------------------------
     .AddObject('grdgrup','GridMyNew')
     WITH .grdgrup
          .Left=0
          .Top=.Parent.lab1.Top+.Parent.lab1.Height+12
          .Width=.Parent.Width
          .Height=.Parent.Height-.Top
          .ScrollBars=2
          .RecordSource='curpodr' 
          DO addColumnToGrid WITH 'fGrup.grdGrup',3
          .Column1.Width=30
          .Column2.Width=.Width-.Column1.Width-SYSMETRIC(5)-13-3
          .Column3.Width=0
          .Column1.Alignment=2
          .Column2.Alignment=0   
          .Column1.ControlSource='curpodr.log_gr'
          .Column2.ControlSource='" "+curpodr.name'  
          .Column1.Header1.Caption=''
          .Column2.Header1.Caption='Подразделение'
          .Column2.ReadOnly=.T.                                
          .Column1.AddObject('checkColumn1','checkContainer')
          .Column1.checkColumn1.AddObject('checkMy','checkMy')
          .Column1.CheckColumn1.checkMy.Visible=.T.
          .Column1.CheckColumn1.checkMy.Caption=''
          .Column1.CheckColumn1.checkMy.Left=5
          .Column1.CheckColumn1.checkMy.Top=3
          .Column1.CheckColumn1.checkMy.BackStyle=0        
          .Column1.CheckColumn1.checkMy.ControlSource='curpodr.log_gr'                                                                                                  
          .column1.CurrentControl='checkColumn1'       
          .SetAll('Enabled',.F.,'ColumnMy')
          .Column1.Enabled=.T. 
          .Column1.Sparse=.F. 
          .Columns(.ColumnCount).Enabled=.T.          
     ENDWITH  
     DO gridSizeNew WITH 'fgrup','grdgrup','shapeingrid1',.T.   
     SELECT datagrup
ENDWITH 
fgrup.Show
**************************************************************************************************************************
*                          Создание новой группы печати
**************************************************************************************************************************
PROCEDURE readgrup
PARAMETERS parlog
SELECT datagrup
IF parlog  
   SELECT datagrup
   fgrup.fgrid.SetFocus
   fgrup.fgrid.GridUpdate        
   SET DELETED OFF
   LOCATE FOR DELETED()
   IF FOUND()
      RECALL
      BLANK
   ELSE
      APPEND BLANK
   ENDIF
   SET DELETED ON
   fgrup.fgrid.Refresh  
ENDIF
WITH fGRup
     .fGrid.columns(.fGrid.columnCount).SetFocus
     SELECT datagrup
     .SetAll('Visible',.T.,'MyTxtBox')
     .nrec=RECNO()
     SCATTER TO .dim_ap  
     .txtBox1.Left=.fgrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .txtBox3.Left=.txtbox2.Left+.txtbox2.Width-1
     .txtbox1.ControlSource='fgrup.dim_ap(1)'
     .txtbox2.ControlSource='fgrup.dim_ap(2)'
     .txtbox3.ControlSource='fgrup.dim_ap(3)'
     lineTop=.fgrid.Top+.fgrid.HeaderHeight+.fgrid.RowHeight*(IIF(.fgrid.RelativeRow<=0,1,.fgrid.RelativeRow)-1)
   
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',fgrup.fgrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .fgrid.Enabled=.F.
     .grdGrup.Column1.Enabled=.T.    
     DO procotmgrup
    * .Refresh
     .txtbox1.SetFocus  
ENDWITH       
**************************************************************************************************************************************************
*                            Процедура записи группы печати
**************************************************************************************************************************************************
PROCEDURE writegrup
PARAMETERS par_log
IF par_log
   SELECT curpodr
   nrec=RECNO()
   GO TOP
   newsost=','
   DO WHILE !EOF()
      newsost=IIF(log_gr,newsost+LTRIM(STR(kod))+',',newsost)   
      SKIP
   ENDDO
   fgrup.dim_ap(4)=newsost
   SELECT datagrup
   GATHER FROM fgrup.dim_ap  
   REPLACE sostav1 WITH newsost
   SELECT datagrup    
ENDIF   
WITH fGrup
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.     
     .SetAll('Visible',.F.,'MyTxtBox')
     .grdGrup.Column1.Enabled=.F.
     .fgrid.Enabled=.T.
     .fgrid.SetAll('Enabled',.F.,'ColumnMy')
     .fgrid.Columns(fgrup.fgrid.ColumnCount).Enabled=.T.     
     .fgrid.GridUpdate
     GO .nrec  
     SELECT datagrup
     GO .nrec
     DO procotmgrup
ENDWITH 
**************************************************************************************************************************
*                          Удаление группы печати
**************************************************************************************************************************
PROCEDURE delgrup
nFormMes.Release
SELECT datagrup
DELETE    
GO TOP  
SELECT curpodr
REPLACE log_gr WITH .F. ALL
REPLACE log_gr WITH .T. FOR ','+LTRIM(STR(kp))+','$datagrup->sostav1
fgrup.Refresh
*************************************************************************************************************************
*                        Обработка события при выборе группы печати
*************************************************************************************************************************
PROCEDURE procotmgrup
fgrup.Lab1.Caption=ALLTRIM(datagrup.name)
SELECT curpodr
REPLACE log_gr WITH .F. ALL
REPLACE log_gr WITH .T. FOR ','+LTRIM(STR(kod))+','$datagrup.sostav1
GO TOP
fgrup.Refresh
**************************************************************************************************************************
*               Выход из групп печати
**************************************************************************************************************************
PROCEDURE returngrup
SELECT curpodr
USE
SELECT datagrup
USE
SELECT people
fgrup.Release