PARAMETERS parUv
IF !USED('jobBook')
   USE jobBook ORDER 1 IN 0
ENDIF
PUBLIC newOrg,newYst,newMst,newDst,newStrSt,newDBeg,newDEnd,currentstaj
IF !parUv
   DO actualStajToday WITH 'people','people.date_in','DATE()'
   SELECT * FROM jobBook WHERE nidpeop=people.nid INTO CURSOR curJobBooK READWRITE 
   SELECT curJobBook
   newDateIn=people.date_in
   newDateOut=people.date_out
ELSE 
   DO actualStajToday WITH 'peopout','peopout.date_in','peopout.date_out'
   SELECT * FROM jobBook WHERE nidpeop=peopout.nid INTO CURSOR curJobBooK READWRITE 
   SELECT curJobBook
   newDateIn=peopout.date_in
   newDateOut=peopout.date_out 
ENDIF 
WITH oPageBook
     DO addButtonOne WITH 'oPageBook','menuCont1',10,0,'нова€','','DO readJobBook WITH .T.',39,RetTxtWidth('редакци€w')+44,'нова€'
     DO addButtonOne WITH 'oPageBook','menuCont2',.menucont1.Left+.menucont1.Width+3,.menucont1.Top,'редакци€','','DO readJobBook WITH .F.',39,.menucont1.Width,'редакци€'
     DO addButtonOne WITH 'oPageBook','menuCont3',.menucont2.Left+.menucont2.Width+3,.menucont1.Top,'удаление','','DO delJobBook',39,.menucont1.Width,'удаление'  
     .SetAll('Enabled',IIF(!parUv,.T.,.F.),'myCommandButton')   
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=nParent.Width
          .Height=fPersCard.Height-.Parent.menuCont1.Height-dHeight*3-SYSMETRIC(4)-60
          .ScrollBars=2          
          .ColumnCount=5        
          .RecordSourceType=1   
          .RecordSource='curJobBook'
          .Column1.ControlSource='" "+nameorg'       
          .Column2.ControlSource='dBeg'
          .Column3.ControlSource='dEnd'
          .Column4.ControlSource='strst'
          .Column2.Width=RettxtWidth('99/99/99999')          
          .Column3.Width=.Column2.Width
          .Column4.Width=.Column2.Width 
          .Column1.Width=.Width-.column2.width-.column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount 
          .Columns(.ColumnCount).Width=0 
          .Column1.Header1.Caption='место работы' 
          .Column2.Header1.Caption='с'
          .Column3.Header1.Caption='по'
          .Column4.Header1.Caption='стаж' 
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0           
          .Column3.Alignment=0
          .colNesInf=2    
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH    
     DO gridSize WITH 'fPersCard.pagePeop.mPage7','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fPersCard.pagePeop.mPage7.fGrid.RecordSource)#fPersCard.pagePeop.mPage7.fGrid.curRec,fPersCard.pagePeop.mPage7.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fPersCard.pagePeop.mPage7.fGrid.RecordSource)#fPersCard.pagePeop.mPage7.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR    
           
     DO addtxtboxmy WITH 'oPageBook',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'oPageBook',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'oPageBook',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,0,'DO countStajBook'
     DO addtxtboxmy WITH 'oPageBook',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
     .menucont1.Left=(.fGrid.Width-.menucont1.Width-.menucont2.Width-.menucont3.Width-20)/2
     .menucont2.Left=.menucont1.Left+.menucont1.Width+10                   
     .menucont3.Left=.menucont2.Left+.menucont2.Width+10  
                      
     hs1=' прин€т (изменить двойной щелчок мыши)+стаж в организации'
     hs2=' стаж всего'
     hs3=' стаж до приема'
     hs4=' уволен (изменить двойной щелчок мыши)'
     DO addContFormNew WITH 'oPageBook','tBox111',.fGrid.Left,.fGrid.Top+.fGrid.Height-1,.fGrid.Column1.Width+12,.fGrid.RowHeight,hs3,0,.F.,.F.,.F.,.F. 
     DO adtBoxNew WITH 'opageBook','tBox112',.tBox111.Top,.tBox111.Left+.tBox111.Width-1,.fGrid.Column2.Width+2,.tBox111.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox113',.tBox111.Top,.tBox112.Left+.tBox112.Width-1,.fGrid.Column3.Width+2,.tBox111.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox114',.tBox111.Top,.tBox113.Left+.tBox113.Width-1,.fGrid.Column4.Width+2,.tBox111.Height,IIF(!parUv,'people.staj_in','peopout.staj_in'),.F.,.F.,0   
     
     
     DO addContFormNew WITH 'oPageBook','tBox11',.fGrid.Left,.tBox111.Top+.tBox111.Height-1,.fGrid.Column1.Width+12,.fGrid.RowHeight,hs1,0,.F.,IIF(!parUv,'DO readDateIn',.F.),.F.,.F. 
 
     DO adtBoxNew WITH 'opageBook','tBox12',.tBox11.Top,.tBox11.Left+.tBox11.Width-1,.fGrid.Column2.Width+2,.tBox11.Height,IIF(!parUv,'people.date_in','peopout.date_in'),.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox13',.tBox11.Top,.tBox12.Left+.tBox12.Width-1,.fGrid.Column3.Width+2,.tBox11.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox14',.tBox11.Top,.tBox13.Left+.tBox13.Width-1,.fGrid.Column4.Width+2,.tBox11.Height,'stajOrg',.F.,.F.,0   
     
     *DO adtBoxNew WITH 'opageBook','tBox21',.tBox11.Top+.tBox11.Height-1,.fGrid.Left,.fGrid.Column1.Width+12,.fGrid.RowHeight,'hs2',.F.,.F.,0 
     DO addContFormNew WITH 'oPageBook','tBox21',.fGrid.Left,.tBox11.Top+.tBox11.Height-1,.fGrid.Column1.Width+12,.fGrid.RowHeight,hs2,0,.F.,.F.,.F.
     DO adtBoxNew WITH 'opageBook','tBox22',.tBox21.Top,.tBox11.Left+.tBox11.Width-1,.fGrid.Column2.Width+2,.tBox11.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox23',.tBox21.Top,.tBox12.Left+.tBox12.Width-1,.fGrid.Column3.Width+2,.tBox11.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox24',.tBox21.Top,.tBox13.Left+.tBox13.Width-1,.fGrid.Column4.Width+2,.tBox11.Height,IIF(!parUv,'people.staj_today','peopout.staj_today'),.F.,.F.,0   
   
     DO addContFormNew WITH 'oPageBook','tBox31',.fGrid.Left,.tBox21.Top+.tBox21.Height-1,.fGrid.Column1.Width+12,.fGrid.RowHeight,hs4,0,.F.,IIF(!parUv,'DO readDateOut',.F.),.F.,.F. 
     DO adtBoxNew WITH 'opageBook','tBox32',.tBox31.Top,.tBox11.Left+.tBox11.Width-1,.fGrid.Column2.Width+2,.tBox11.Height,IIF(!parUv,'people.date_out','peopout.date_out'),.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox33',.tBox31.Top,.tBox12.Left+.tBox12.Width-1,.fGrid.Column3.Width+2,.tBox11.Height,'',.F.,.F.,0
     DO adtBoxNew WITH 'opageBook','tBox34',.tBox31.Top,.tBox13.Left+.tBox13.Width-1,.fGrid.Column4.Width+2,.tBox11.Height,'',.F.,.F.,0   
    
     *--------------------------------- нопка записать-------------------------------------------------------------------------       
     DO addButtonOne WITH 'oPageBook','butSave',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WзаписатьW')*2-20)/2,.menucont1.Top,'записать','','DO writeJobBook WITH .T.',39,RetTxtWidth('wзаписатьw'),'записать'
     *--------------------------------- нопка возврат при редакции--------------------------------------------------------------                                                     
     DO addButtonOne WITH 'oPageBook','butRet',.butSave.Left+.butsave.Width+20,.butSave.Top,'возврат','','DO writeJobBook WITH .F.',39,.butsave.Width,'возврат'
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     
     *--------------------------------- нопка удалить-------------------------------------------------------------------------
     DO addButtonOne WITH 'oPageBook','butSaveDate',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WудалитьW')*2-20)/2,.menucont1.Top,'записать','','DO writeDateIn WITH .T.',39,RetTxtWidth('wудалитьw'),'удалить'
     *--------------------------------- нопка возврат при удалении-------------------------------------------------------------------------                                            
     DO addButtonOne WITH 'oPageBook','butRetDate',.butSaveDate.Left+.butSaveDate.Width+20,.butSaveDate.Top,'возврат','','DO writeDateIn WITH .F.',39,.butSaveDate.Width,'возврат'
       
     .butSaveDate.Visible=.F.
     .butRetDate.Visible=.F. 
     *--------------------------------- нопка удалить-------------------------------------------------------------------------
     DO addButtonOne WITH 'oPageBook','butDel',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WудалитьW')*2-20)/2,.menucont1.Top,'удалить','','DO delRecJobBook WITH .T.',39,RetTxtWidth('wудалитьw'),'удалить'
     *--------------------------------- нопка возврат при удалении-------------------------------------------------------------------------                                            
     DO addButtonOne WITH 'oPageBook','butDelRet',.butDel.Left+.butDel.Width+20,.butDel.Top,'возврат','','DO delRecJobBook WITH .F.',39,.butDel.Width,'возврат'
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.    
         
     .setAll('Top',.fGrid.Top+.fGrid.height+dHeight*4+20,'mymenuCont')   
     .setAll('Top',.fGrid.Top+.fGrid.height+dHeight*4+20,'myCommandButton') 
     .setAll('Top',.fGrid.Top+.fGrid.height+dHeight*4+20,'myContLabel') 
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
ENDWITH
**************************************************************************************************************************
PROCEDURE readDateIn
newDateIn=people.date_in
WITH oPageBook
     .SetAll('Visible',.F.,'MyCommandButton')
     .SetAll('Visible',.F.,'mymenucont')
     .butSaveDate.Visible=.T.
     .butRetDate.Visible=.T.
     .butSaveDate.procForClick='DO writeDateIn WITH .T.'
     .butRetDate.procForClick='DO writeDateIn WITH .F.'
     .fGrid.Enabled=.F.
     .tBox12.Enabled=.T.
     .tBox12.ControlSource='newDateIn'
     .tBox12.SetFocus
ENDWITH 
**************************************************************************************************************************
PROCEDURE writeDateIn
PARAMETERS par1
IF par1
   SELECT people
   REPLACE date_in WITH newDateIn
   SELECT curJobBook   
ENDIF
WITH oPageBook
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butSaveDate.Visible=.F.
     .butRetDate.Visible=.F.
     .fGrid.Enabled=.F.
     .tBox12.Enabled=.F.
     .tBox12.ControlSource='people.date_in'
     .fGrid.Enabled=.T. 
     DO countTotStaj  
     *SELECT curjobBook 
     *GO famRec
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .Refresh
ENDWITH 
**************************************************************************************************************************
PROCEDURE readDateOut
newDateOut=people.date_out
WITH oPageBook
     .SetAll('Visible',.F.,'MyCommandButton')
     .SetAll('Visible',.F.,'mymenucont')
     .butSaveDate.Visible=.T.
     .butRetDate.Visible=.T.
     .butSaveDate.procForClick='DO writeDateOut WITH .T.'
     .butRetDate.procForClick='DO writeDateOut WITH .F.'
     .fGrid.Enabled=.F.
     .tBox32.Enabled=.T.
     .tBox32.ControlSource='newDateOut'
     .tBox32.SetFocus
ENDWITH 
**************************************************************************************************************************
PROCEDURE writeDateOut
PARAMETERS par1
IF par1
   SELECT people
   REPLACE date_out WITH newDateout
   SELECT curJobBook   
ENDIF
WITH oPageBook
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butSaveDate.Visible=.F.
     .butRetDate.Visible=.F.
     .fGrid.Enabled=.F.
     .tBox32.Enabled=.F.
     .tBox32.ControlSource='people.date_out'
   *  .tBox12.ControlSource='people.date_in'
     .fGrid.Enabled=.T. 
*     DO countTotStaj  
     *SELECT curjobBook 
     *GO famRec
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .Refresh
ENDWITH 
**************************************************************************************************************************
PROCEDURE readJobBook
PARAMETERS par1
SELECT curJobBook
IF curJobBook.kodpeop=0.AND.!par1
   RETURN 
ENDIF
IF par1     
   APPEND BLANK     
   REPLACE kodpeop WITH people.num,nidpeop WITH people.nid 
ENDIF
fPersCard.pagePeop.mPage7.Refresh
log_ap=IIF(par1,.T.,.F.)
newdBeg=IIF(par1,CTOD('  .  .    '),dBeg)
newdEnd=IIF(par1,CTOD('  .  .    '),dEnd)
newOrg=IIF(par1,'',nameOrg)
newStrSt=IIF(par1,'',strSt)
newYst=yst
newMst=mst
newDst=dst
newStrSt=strSt
famRec=RECNO()
WITH oPageBook    
     .fGrid.Refresh
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
     
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butSave.Visible=.T.
     .butRet.Visible=.T.

     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width-1
     .txtBox3.Left=.txtBox2.Left+.txtBox2.Width-1
     .txtBox4.Left=.txtBox3.Left+.txtBox3.Width-1     
     .txtbox1.ControlSource='newOrg'
     .txtBox2.ControlSource='newdBeg'
     .txtbox3.ControlSource='newdEnd'
     .txtBox4.ControlSource='newStrSt'
     .txtBox1.Top=lineTop 
     .txtBox2.Top=lineTop 
     .txtBox3.Top=lineTop 
     .txtBox4.Top=lineTop 
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .SetAll('Visible',.T.,'MyTxtBox')
     .fGrid.Enabled=.F.
     .txtBox1.SetFocus
ENDWITH 
************************************************************************************************************************
PROCEDURE countStajBook
lognmonth=.F.
IF !EMPTY(newdBeg).AND.!EMPTY(newdEnd).AND.newdEnd=>newdBeg        
   currentStaj=''
   currentDate=newdEnd 
   dMbeg=0
   dMEnd=0     
   y_new=0
   m_new=0
   d_new=0 
   dayMonthBeg=0
   dayMonthEnd=0 
   IF MONTH(newDBeg)=2 
      dayMonthBeg=IIF(MOD(YEAR(newDBeg),4)=0,29,28)
   ELSE
      dayMonthBeg=IIF(INLIST(MONTH(newDBeg),1,3,5,7,8,10,12),31,30) &&кол-во дней в начальном мес€це  
   ENDIF
   IF MONTH(newDBeg)=2 
      dayMonthEnd=IIF(MOD(YEAR(newDEnd),4)=0,29,28)
   ELSE
      dayMonthEnd=IIF(INLIST(MONTH(newDEnd),1,3,5,7,8,10,12),31,30)  &&кол-во дней в конечном мес€це
   ENDIF
 *-----считаем дни
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMBeg=DAY(newDEnd)-DAY(newDBeg)+1
      IF DAY(newDBeg)=1.AND.DAY(newDEnd)=dayMonthEnd
         lognmonth=.T.
      ENDIF
   ELSE       
      dMbeg=dayMonthBeg-DAY(newDBeg)+1      
   ENDIF  
   IF dMBeg=dayMonthBeg
      m_new=m_new+1
      dMbeg=0    
   ENDIF   
   d_new=d_new+dMBeg 
   dMEnd=0   
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMEnd=0
   ELSE
      dMEnd=DAY(newDEnd)
   ENDIF    
   IF dMEnd=dayMonthEnd      
      m_new=m_new+1      
      dMEnd=0
   ENDIF  
   d_new=d_New+dMEnd
 *-------считаем мес€цы 
   mYbeg=0
   mYEnd=0
  
   IF YEAR(newDBeg)=YEAR(newDEnd)  
      IF lognmonth=.T.
      ELSE     
          m_new=m_new+MONTH(newDEnd)-MONTH(newDBeg)-1
          m_new=IIF(m_new<0,0,m_new)
      ENDIF     
   ELSE 
      mYbeg=12-MONTH(newDBeg)
      mYEnd=MONTH(newDEnd)-1
      
      m_new=m_new+mYbeg+mYEnd
   ENDIF    
   *------------считаем годы------
   IF YEAR(newDBeg)=YEAR(newDEnd)
      y_new=0
   ELSE 
      y_new=YEAR(newDEnd)-YEAR(newDBeg)-1   
   ENDIF 

   IF d_new>=30     
      d_new=d_new-30
      m_new=m_new+1     
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   newYst=y_new
   newMst=m_new
   newDst=d_new
   stajOrg=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')  
   newStrSt=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')  
ELSE    
ENDIF
oPageBook.Refresh
************************************************************************************************************************
PROCEDURE writeJobBook
PARAMETERS par1
WITH oPageBook
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butSaveDate.Visible=.F.
     .butRetDate.Visible=.F.
     SELECT curJobBook   
     ON ERROR DO erSup  
     GO famRec
     ON ERROR
     IF par1
        REPLACE dBeg WITH newdBeg,dEnd WITH newdEnd,nameorg WITH newOrg,ySt WITH newYst,mSt WITH newMst,dSt WITH newDst,strSt WITH newStrSt
        SELECT jobBook
        DELETE FOR nidPeop=people.nid
        APPEND FROM DBF('curJobBook')
        SELECT curjobBook
     ELSE
        IF log_ap     
           DELETE               
        ENDIF           
     ENDIF   
     .txtBox1.Visible=.F.
     .txtBox2.Visible=.F.
     .txtBox3.Visible=.F.
     .txtBox4.Visible=.F.
     .SetAll('Visible',.F.,'ComboMy')     
     .fGrid.Enabled=.T. 
     DO countTotStaj  
     SELECT curjobBook 
     ON ERROR DO ersup
     GO famRec
     ON ERROR 
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH  
oPageBook.Refresh 
ON ERROR DO erSup
GO famRec
ON ERROR 
oPageBook.fGrid.Columns(oPageBook.fGrid.columnCount).SetFocus  
**************************************************************************************************************************
PROCEDURE delJobBook
SELECT curJobBook
IF curJobBook.kodpeop=0
   RETURN 
ENDIF
WITH oPageBook
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butDel.Visible=.T.
     .butDelRet.Visible=.T.
     .fGrid.Enabled=.F.
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE delRecJobBook
PARAMETERS par1
SELECT curJobBook
IF par1
   DELETE
   SELECT jobBook
   DELETE FOR nidPeop=people.nid
   APPEND FROM DBF('curJobBook')
   SELECT curjobBook
   DO countTotStaj
ENDIF
WITH oPageBook
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butSaveDate.Visible=.F.
     .butRetDate.Visible=.F.
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
     .Refresh
ENDWITH
************************************************************************************************************************
PROCEDURE countTotStaj
SELECT curJobBook
SUM dst,mst,yst TO d_tot,m_tot,y_tot FOR kodpeop=people.num
IF d_tot>=30
   m_tot=m_tot+INT(d_tot/30)
   d_tot=d_tot-30*INT(d_tot/30)
ENDIF
IF m_tot>=12
   y_tot=y_tot+INT(m_tot/12)
   m_tot=m_tot-12*INT(m_tot/12)
ENDIF
SELECT people
REPLACE staj_in WITH PADL(ALLTRIM(STR(y_tot)),2,'0')+'-'+PADL(ALLTRIM(STR(m_tot)),2,'0')+'-'+PADL(ALLTRIM(STR(d_tot)),2,'0')
DO actualStajToday WITH 'people','people.date_in','DATE()'