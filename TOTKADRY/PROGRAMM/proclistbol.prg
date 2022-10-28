PARAMETERS paruv
PUBLIC logUv,newPrimOtp
newPrimOtp=''
logUv=parUv
IF !USED('dataList')
   USE dataList ORDER 2 IN 0
ENDIF 
SELECT dataList
IF !parUv
   SET FILTER TO dataList.nidpeop=people.nid
ELSE 
   SET FILTER TO dataList.nidpeop=peopout.nid
ENDIF    
PUBLIC newdBeg,newdEnd,newPrim,newDays,nrec
daysTot=0
SUM days TO daysTot
daysTot=ROUND(daysTot,0)
GO TOP
nrec=0
log_ap=.F.
newdBeg=CTOD('  .  .   ')
newdEnd=CTOD('  .  .   ')
newPrim=SPACE(50)
newPrimOtp=SPACE(50)
newDays=0
GO TOP
WITH oPageBol
     .Caption='больничные листы'         
     DO addButtonOne WITH 'oPageBol','menuCont1',10,nParent.Height-dHeight-40,'нова€','','DO readList WITH .T.',39,RetTxtWidth('справочникw'),'нова€'
     DO addButtonOne WITH 'oPageBol','menuCont2',.menucont1.Left+.menucont1.Width+3,.menucont1.Top,'редакци€','','DO readList WITH .F.',39,.menucont1.Width,'редакци€'
     DO addButtonOne WITH 'oPageBol','menuCont3',.menucont2.Left+.menucont2.Width+3,.menucont1.Top,'удаление','','DO delList',39,.menucont1.Width,'удаление'   
     .SetAll('Enabled',IIF(!parUv,.T.,.F.),'myCommandButton')     
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=nParent.Width
          .Height=.Parent.menuCont1.Top-60
          .ScrollBars=2          
          .ColumnCount=6       
          .RecordSourceType=1     
          .RecordSource='datalist'
          .Column1.ControlSource='dataList.dBeg'
          .Column2.ControlSource='dataList.dEnd'
          .Column3.ControlSource='datalist.days'
          .Column4.ControlSource='" "+dataList.primList'       
          .Column5.ControlSource='" "+dataList.primOtp'       
          .Column1.Width=RettxtWidth('окончаниеw')
          .Column2.Width=.column1.Width
          .Column3.Width=RetTxtWidth('wдниw')                       
          .Column4.Width=(.Width-.column1.width-.column2.Width-.Column3.Width)/2
          .Column5.Width=.Width-.column1.width-.column2.Width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount       
          .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='начало'
          .Column2.Header1.Caption='окончание'
          .Column3.Header1.Caption='дни'  
          .Column4.Header1.Caption='номер б/листа'                 
          .Column5.Header1.Caption='отпуск'                 
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0           
          .Column3.Alignment=1         
          .Column4.Alignment=0
          .Column5.Alignment=0
          .Column3.Format='Z'
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'opageBol','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(opageBol.fGrid.RecordSource)#oPageBol.fGrid.curRec,opageBol.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(oPageBol.fGrid.RecordSource)#oPageBol.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR       
     DO addtxtboxmy WITH 'opageBol',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,0,'DO validDaysList'  
     DO addtxtboxmy WITH 'opageBol',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0,'DO validDaysListEnd'
     DO addtxtboxmy WITH 'opageBol',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'opageBol',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'opageBol',5,1,1,.fGrid.Column5.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')                     
     DO adtBoxNew WITH 'opageBol','tBox11',.fGrid.Top+.fGrid.Height-1,.fGrid.Left,.fGrid.Column1.Width+12,.fGrid.RowHeight,'',.F.,.F.,0 
     DO adtBoxNew WITH 'opageBol','tBox12',.tBox11.Top,.tBox11.Left+.tBox11.Width-1,.fGrid.Column2.Width+2,.tBox11.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBol','tBox13',.tBox11.Top,.tBox12.Left+.tBox12.Width-1,.fGrid.Column3.Width+2,.tBox11.Height,'daysTot','Z',.F.,0
     DO adtBoxNew WITH 'opageBol','tBox14',.tBox11.Top,.tBox13.Left+.tBox13.Width-1,.fGrid.Column4.Width+2,.tBox11.Height,'',.F.,.F.,0   
     DO adtBoxNew WITH 'opageBol','tBox15',.tBox11.Top,.tBox14.Left+.tBox14.Width-1,.fGrid.Column5.Width+2,.tBox11.Height,'',.F.,.F.,0   
     
     .menucont1.Left=(.fGrid.Width-.menucont1.Width-.menucont2.Width-.menucont3.Width-20)/2
     .menucont2.Left=.menucont1.Left+.menucont1.Width+10                   
     .menucont3.Left=.menucont2.Left+.menucont2.Width+10           
     .setAll('Top',.tBox11.Top+.tBox11.height+20,'mymenuCont')   
     .setAll('Top',.tBox11.Top+.tBox11.height+20,'myCommandButton') 
     .setAll('Top',.tBox11.Top+.tBox11.height+20,'myContLabel')     
     IF !paruv
        *--------------------------------- нопка записать-------------------------------------------------------------------------       
        DO addButtonOne WITH 'oPageBol','butSave',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WзаписатьW')*2-20)/2,.menucont1.Top,'записать','','DO writeList WITH .T.',39,RetTxtWidth('wзаписатьw'),'записать'
        *--------------------------------- нопка возврат при редакции--------------------------------------------------------------                                                     
        DO addButtonOne WITH 'oPageBol','butRet',.butSave.Left+.butsave.Width+20,.butSave.Top,'возврат','','DO writeList WITH .F.',39,.butsave.Width,'возврат'
        .butSave.Visible=.F.
        .butRet.Visible=.F.  
        *--------------------------------- нопка удалить-------------------------------------------------------------------------
        DO addButtonOne WITH 'oPageBol','butDel',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WудалитьW')*2-20)/2,.menucont1.Top,'удалить','','DO delRecList WITH .T.',39,RetTxtWidth('wудалитьw'),'удалить'
        *--------------------------------- нопка возврат при удалении-------------------------------------------------------------------------                                            
        DO addButtonOne WITH 'oPageBol','butDelRet',.butDel.Left+.butDel.Width+20,.butDel.Top,'возврат','','DO delRecList WITH .F.',39,.butDel.Width,'возврат'
        .butDel.Visible=.F.
        .butDelRet.Visible=.F.    
     ENDIF 
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE readList
PARAMETERS par1
IF logUv
   RETURN
ENDIF
SELECT dataList
IF datalist.kodpeop=0.AND.!par1
   RETURN 
ENDIF
IF par1     
   APPEND BLANK     
   REPLACE nidpeop WITH people.nid,kodpeop WITH people.num
ENDIF
opageBol.Refresh
log_ap=IIF(par1,.T.,.F.)
newdBeg=IIF(par1,CTOD('  .  .    '),dBeg)
newdEnd=IIF(par1,CTOD('  .  .    '),dEnd)
newPrim=IIF(par1,SPACE(50),primList)
newPrimOtp=IIF(par1,SPACE(50),primOtp)
newDays=IIF(par1,0,days)
nrec=RECNO()
WITH oPageBol
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butSave.Visible=.T.
     .butRet.Visible=.T.
     .fGrid.Columns(.fGrid.columnCount).SetFocus
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width-1
     .txtBox3.Left=.txtBox2.Left+.txtBox2.Width-1
     .txtBox4.Left=.txtBox3.Left+.txtBox3.Width-1
     .txtBox5.Left=.txtBox4.Left+.txtBox4.Width-1
     .txtBox1.ControlSource='newdBeg'
     .txtbox2.ControlSource='newdEnd'
     .txtbox3.ControlSource='newDays'
     .txtbox4.ControlSource='newPrim'
     .txtbox5.ControlSource='newPrimOtp'
     .txtBox1.Top=lineTop
     .txtBox2.Top=lineTop
     .txtBox3.Top=lineTop
     .txtBox4.Top=lineTop
     .txtBox5.Top=lineTop
     .txtBox1.Height=.fGrid.RowHeight+1
     .txtBox2.Height=.txtBox1.Height
     .txtBox3.Height=.txtBox1.Height
     .txtBox4.Height=.txtBox1.Height
     .txtBox5.Height=.txtBox1.Height
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
   opageBol.txtBox3.Refresh
ENDIF
************************************************************************************************************************
PROCEDURE validDaysListEnd
IF !EMPTY(newDBeg).AND.!EMPTY(newDEnd)
   newPrimOtp=''
   SELECT datotp
   oldOtpRec=RECNO()
   oldOtpOrd=SYS(21)
   SET ORDER TO 6
   SEEK people.nid
   SCAN WHILE nidpeop=people.nid
        IF (BETWEEN(newdBeg,begOtp,endOtp).OR.BETWEEN(newdEnd,begOtp,endOtp)).AND.INLIST(kodotp,1,2,3,4,5,7,9)
           newPrimOtp=ALLTRIM(newprimOtp)+ALLTRIM(nameotp)+' '+DTOC(begOtp)+' - '+DTOC(endOtp)
           newPrimOtp=ALLTRIM(newPrimOtp)         
        ENDIF 
   ENDSCAN            
   SET ORDER TO &oldOtpOrd
   GO oldOtpRec
   SELECT datalist
   newDays=newDEnd-newDBeg+1
   opageBol.txtBox3.Refresh
   opageBol.txtBox5.Refresh
ENDIF
************************************************************************************************************************
PROCEDURE writeList
PARAMETERS par1
WITH oPageBol
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     SELECT dataList
     IF par1
        REPLACE dBeg WITH newdBeg,dEnd WITH newdEnd,primList WITH newPrim,days WITH newDays,primOtp WITH newPrimOtp
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
     .tBox15.Visible=.T.        
     .fGrid.Enabled=.T.    
     GO nrec
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH  
oPageBol.Refresh 
oPageBol.fGrid.Columns(oPageBol.fGrid.columnCount).SetFocus  
*************************************************************************************************************************
PROCEDURE delList
IF logUv
   RETURN
ENDIF
IF datalist.kodpeop=0
   RETURN 
ENDIF
WITH oPageBol
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butDel.Visible=.T.
     .butDelRet.Visible=.T.
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE delRecList
PARAMETERS par1
IF par1
   SELECT dataList
   DELETE
   SUM days TO daysTot
   daysTot=ROUND(daysTot,0) 
   GO TOP
ENDIF   
WITH oPageBol
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .Refresh
ENDWITH