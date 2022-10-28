SELECT datShtat
GO TOP
cNewFullName=''
logFl=.F.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сохраненных тарификаций'  
     .Width=600
     .Height=400
     .AddObject('fGrid','GridMyNew')  
     WITH .fGrid     
          .Top=0
          .Left=0
          .Height=.Parent.Height
          .Width=.Parent.Width
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='datshtat'
           DO addColumnToGrid WITH 'fSupl.fGrid',5
          .Column1.ControlSource='dtarif'
          .Column2.ControlSource='" "+fullname'  
          .Column3.ControlSource='" "+pathtarif'   
          .Column4.ControlSource='luse'
          .Column1.Header1.Caption='дата'
          .Column2.Header1.Caption='наименование' 
          .Column3.Header1.Caption='каталог' 
          .Column4.Header1.Caption='!' 
          .Column1.Width=RettxtWidth('99/99/99999')      
          .Column3.Width=RettxtWidth('TAR99/99/9999w')
          .Column4.Width=RettxtWidth('W!W')
          .Columns(.ColumnCount).Width=0          
          .Column2.Width=.Width-.column1.width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount 
          .Column1.Alignment=1        
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column1.Movable=.F.         
          .colNesInf=2    
          .SetAll('BackColor',RGB(255,255,255),'ColumnMy')    
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.             
          
        *  .Column4.ReadOnly=.T.                                
          .Column4.AddObject('checkColumn4','checkContainer')
          .Column4.checkColumn4.AddObject('checkMy','checkMy')
          .Column4.CheckColumn4.checkMy.Visible=.T.
          .Column4.CheckColumn4.checkMy.Caption=''
          .Column4.CheckColumn4.checkMy.Left=5
          .Column4.CheckColumn4.checkMy.Top=3
          .Column4.CheckColumn4.checkMy.BackStyle=0        
          .Column4.CheckColumn4.checkMy.ControlSource='luse'                                                                                                  
          .column4.CurrentControl='checkColumn4'       
          .SetAll('Enabled',.F.,'ColumnMy')
          .Column4.Enabled=.T. 
          .Column4.Sparse=.F. 
          .Columns(.ColumnCount).Enabled=.T.  
    ENDWITH      
    DO gridSizeNew WITH 'fSupl','fGrid','shapeingrid',.T.
    FOR i=1 TO .fGrid.columnCount        
        .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fSupl.fGrid.RecordSource)#fSupl.fGrid.curRec,fSupl.BackColor,dynBackColor)'
        .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fSupl.fGrid.RecordSource)#fSupl.fGrid.curRec,dForeColor,dynForeColor)'        
    ENDFOR 
    DO addtxtboxmy WITH 'fSupl',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
    .txtbox1.Enabled=.F.
    DO addtxtboxmy WITH 'fSupl',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
    .SetAll('Visible',.F.,'MyTxtBox')  
    DO addButtonOne WITH 'fSupl','butRead',(.Width-RetTxtWidth('wудалениеw')*3-30)/2,.fGrid.Top+.fGrid.Height+20,'редакция','','DO readSpisTar',39,RetTxtWidth('wудалениеw'),'редакция'  
    DO addButtonOne WITH 'fSupl','butDel',.butRead.Left+.butRead.Width+10,.butRead.Top,'удаление','','DO delSpisTar',39,.butRead.Width,'удаление'  
    DO addButtonOne WITH 'fSupl','butExit',.butDel.Left+.butDel.Width+10,.butRead.Top,'возврат','','fSupl.Release',39,.butRead.Width,'возврат'  
    DO addButtonOne WITH 'fSupl','butSave',(.Width-RetTxtWidth('wзаписатьw')*2-20)/2,.butRead.Top,'записать','','DO saveSpisTar WITH .T.',39,RetTxtWidth('wзаписатьw'),'записать'  
    DO addButtonOne WITH 'fSupl','butRetRead',.butSave.Left+.butSave.Width+10,.butSave.Top,'возврат','','DO saveSpisTar WITH .F.',39,.butRead.Width,'возврат' 
    .butSave.Visible=.F.
    .butRetRead.Visible=.F.
    
    
    DO adCheckBox WITH 'fSupl','check1','удалить содержимое каталога',.fGrid.Top+.fGrid.Height+5,.fGrid.Left,150,dHeight,'logFl',0,.F.
    .check1.Left=.fGrid.Left+(.fGrid.Width-.check1.Width)/2
    .check1.Visible=.F.
    DO addButtonOne WITH 'fSupl','butDelRec',(.Width-RetTxtWidth('wудалитьw')*2-20)/2,.check1.Top+.check1.Height+5,'удалить','','DO delRecSpisTar WITH .T.',35,RetTxtWidth('wудалитьw'),'удалить'  
    DO addButtonOne WITH 'fSupl','butRetDel',.butDelRec.Left+.butDelRec.Width+10,.butDelRec.Top,'возврат','','DO delRecSpisTar WITH .F.',35,.butRead.Width,'возврат' 
    .butDelRec.Visible=.F.
    .butRetDel.Visible=.F.
    .Height=.fGrid.Height+.butRead.Height+40
    
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************************
PROCEDURE readSpisTar
IF datShtat.Real
   RETURN
ENDIF
cNewFullName=fullname
WITH fSupl
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
     .fGrid.Enabled=.F.
     .setAll('Visible',.F.,'myCommandButton')
     .butSave.Visible=.T.
     .butRetRead.Visible=.T.
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .txtbox1.ControlSource='datshtat.dTarif'
     .txtbox2.ControlSource='cNewFullName'
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .txtbox1.Visible=.T.
     .txtbox2.Visible=.T.
     .txtbox2.SetFocus 
ENDWITH 
***************************************************************************************
PROCEDURE saveSpisTar
PARAMETERS par1
IF par1
   REPLACE fullname WITH cNewFullName
ENDIF
WITH fSupl
     .setAll('Visible',.T.,'myCommandButton')
     .butSave.Visible=.F.
     .butRetRead.Visible=.F.
     .butDelRec.Visible=.F.
     .butRetDel.Visible=.F.
     .SetAll('Visible',.F.,'MyTxtBox')
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'columnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .Refresh
ENDWITH
***************************************************************************************
PROCEDURE delSpisTar
IF datShtat.Real
   RETURN
ENDIF
logFl=.F.
cNewFullName=fullname
WITH fSupl
     .fGrid.Enabled=.F.
     .setAll('Visible',.F.,'myCommandButton')
     .butDelRec.Visible=.T.
     .butRetDel.Visible=.T.
     .check1.Visible=.T.   
ENDWITH 
***************************************************************************************
PROCEDURE delRecSpisTar
PARAMETERS par1
IF par1  
   SELECT datshtat
   IF logFl
      pathDel=pathmain+'\'+ALLTRIM(datshtat.pathtarif)   
      *llForce = .T.  
      lObjDel=CREATEOBJECT('Scripting.FileSystemObject')  
      lObjDel.DeleteFolder(pathDel,.T.)  
      RELEASE lObjDel      
   ENDIF
   DELETE 
ENDIF
WITH fSupl
     .setAll('Visible',.T.,'myCommandButton')
     .butSave.Visible=.F.
     .butRetRead.Visible=.F.
     .butDelRec.Visible=.F.
     .butRetDel.Visible=.F.
     .check1.Visible=.F.
     .SetAll('Visible',.F.,'MyTxtBox')
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'columnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .Refresh
ENDWITH
