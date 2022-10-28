*-------------------------------------------------------------------------------------------------------------------------
*                         Процедуры для штатного расписания
*-------------------------------------------------------------------------------------------------------------------------
****************************************************************************************************************************
*                                         Основная (постоение окна и т.д.)
****************************************************************************************************************************
PARAMETERS par1
SELECT people
oldPeopRec=RECNO()
SELECT datjob
SET FILTER TO 
STORE 0 TO ksedol,numrec,kdold,kpold,ndold,npOld,kse_new,kse_totnew,kse_stacnew,kse_amnew,kse_sknew,kse_mjrnew,kd_new,kat_new,nd_new,numpodr_new,fGridHeight,totItog,totVac,totClvac
STORE .F. TO log_ap,log_rep,logApPodr
STORE '' TO oldnamepodr,strdol,strdolold,strkat,strkatold
log_spis=.F.

SELECT * FROM datjob INTO CURSOR curSupJob READWRITE
SELECT curSupJob
REPLACE dekotp WITH IIF(SEEK(kodpeop,'people',1),people.dekotp,.F.) ALL 
IF !par1
   DELETE FOR dateBeg>DATE()
   DELETE FOR !EMPTY(dateOut).AND.dateOut<=DATE()
ELSE 
   DELETE FOR dateBeg>varDtar
   DELETE FOR !EMPTY(dateOut).AND.dateOut<varDtar   
ENDIF 
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
CREATE CURSOR curRaspPeople (kodpeop N(4),fio C(60),kd N(3),nameDol C(80),kp N(3),kse N(6,2),tr N(1),nameTr C(20),primtxt C(50), dmol D, nid N(5))

=AFIELDS(arJob,'curSupJob')                           && CURSOR для справочника категорий персонала (comboBox)
CREATE CURSOR cursppeople FROM ARRAY arJob
SELECT cursppeople
INDEX ON kodpeop TAG T1

=AFIELDS(arKat,'sprkat')                           && CURSOR для справочника категорий персонала (comboBox)
CREATE CURSOR curSupKat FROM ARRAY arKat
APPEND FROM sprkat

SELECT * FROM sprdolj INTO CURSOR curDoljPeop ORDER BY name READWRITE   && CURSOR  для справочника должностей (персонал)       
SELECT curDoljPeop 

=AFIELDS(arRasp,'rasp')                          &&CURSOR  для развернутого штатного
CREATE CURSOR curRaspTot FROM ARRAY arRasp

ALTER TABLE currasptot ADD COLUMN kse_peop N (7,2)
ALTER TABLE currasptot ADD COLUMN ksepeopcl N (7,2)
SELECT curRaspTot
INDEX ON STR(np,3)+STR(nd,3) TAG T1
INDEX ON kp TAG T2
SET ORDER TO 1

DIMENSION dim_ap(FCOUNT('curRaspTot'))
SELECT sprdolj
SET RELATION TO
SELECT rasp
SET RELATION TO kd INTO sprdolj,kat INTO sprkat ADDITIVE
GO TOP 
*--------------------------------Формирование курсора для штатного расписания---------------------------------------------------
SELECT sprpodr
SET ORDER TO 3
GO TOP
totItog=0
DO WHILE !EOF()
   SELECT currasptot
   APPEND BLANK 
   REPLACE kodkpp WITH sprpodr.kodkpp,named WITH IIF(kodkpp=0,sprpodr.name,'     '+sprpodr.name),kp WITH sprpodr.kod,log_s WITH .T.,log_sp WITH .T.,np WITH sprpodr.np,col1 WITH '+',kpsupl WITH sprpodr.kpsup
   kprec=RECNO()
   SELECT rasp
   SEEK STR(sprpodr.kod,3)
   ksePodr=0
   SCAN WHILE kp=sprpodr.kod
       * SCATTER TO dim_rasp
        SELECT currasptot
        APPEND BLANK
      *  GATHER FROM dim_rasp     
                         
        REPLACE named WITH IIF(!EMPTY(sprdolj.namework),sprdolj.namework,sprdolj.name),kse WITH rasp.kse,kat WITH rasp.kat,nd WITH rasp.nd,namekat WITH sprkat.name,;
                kp WITH rasp.kp,kd WITH rasp.kd,kse_tot WITH rasp.kse_tot,kse_am WITH rasp.kse_am,;
                kse_STAC WITH rasp.kse_stac,kse_sk WITH rasp.kse_sk,kse_mjr WITH rasp.kse_mjr,np WITH sprpodr.np,kodkpp WITH sprpodr.kodkpp,kpsupl WITH IIF(kodkpp=0,kp,kodkpp)                
        ksePodr=ksePodr+kse        
        totItog=totItog+kse
        SELECT rasp       
   ENDSCAN	
   SELECT curRaspTot
   *LOCATE FOR kp=sprpodr.kod.AND.log_s
   GO kprec
   REPLACE kse WITH ksePodr
   SELECT sprpodr
   SKIP 
ENDDO   
SELECT currasptot
SET ORDER TO 1
SCAN ALL
     SELECT curSupJob
     SUM kse TO ksesup FOR IIF(currasptot.log_sp,kp=currasptot.kp.AND.kd#0.AND.!dekotp,kp=currasptot.kp.AND.kd=currasptot.kd.AND.!dekotp)
     
     SUM kse TO ksesupcl FOR IIF(currasptot.log_sp,kp=currasptot.kp.AND.kd#0.AND.INLIST(tr,1,2,3,5).AND.!dekotp,kp=currasptot.kp.AND.kd=currasptot.kd.AND.INLIST(tr,1,2,3,5).AND.!dekotp)
     SELECT currasptot
     REPLACE kse_peop WITH ksesup,ksepeopcl WITH ksesupcl,kse_vac WITH kse-kse_peop,ksevaccl WITH kse-ksepeopcl
ENDSCAN
SET FILTER TO log_sp
SELECT sprpodr
SCAN ALL
     kse_ch=0 
     ksevac_ch=0
     ksevaccl_ch=0
     IF kodkpp#0
        SELECT currasptot
        SEEK STR(sprpodr.np,3)
        kse_ch=kse
        ksevac_ch=kse_vac
        ksevaccl_ch=ksevaccl
        LOCATE FOR kp=sprpodr.kodkpp
        REPLACE kse WITH kse+kse_ch,kse_vac WITH kse_vac+ksevac_ch,ksevaccl WITH ksevaccl+ksevaccl_ch 
     ENDIF
     SELECT sprpodr     
ENDSCAN
SELECT currasptot
SET FILTER TO log_sp
GO TOP
fRasp=CREATEOBJECT('FORMMY')
=SYS(2002)
WITH fRasp  
     .procExit='DO exitTotalShtat'
     DO addButtonOne WITH 'fRasp','menuCont1',10,5,'подразд.','pencila.ico','Do newpodrshtat WITH .T.',39,RetTxtWidth('текст-шапка')+44,'подразделение'    
     DO addButtonOne WITH 'fRasp','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'должн.','pencila.ico','Do readShtat WITH .T.',39,.menucont1.Width,'должность'   
     DO addButtonOne WITH 'fRasp','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'редакция','pencil.ico','Do readShtat WITH .F.',39,.menucont1.Width,'редакция'       
     DO addButtonOne WITH 'fRasp','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'удаление','pencild.ico','Do deleteFromShtat',39,.menucont1.Width,'удаление' 
     DO addButtonOne WITH 'fRasp','menuCont6',.menucont4.Left+.menucont4.Width+3,5,'нумерация','numbers.ico','DO procNumPodr',39,.menucont1.Width,'нумерация'   
     DO addButtonOne WITH 'fRasp','menuCont7',.menucont6.Left+.menucont6.Width+3,5,'поиск','find.ico','Do formFindRasp',39,.menucont1.Width,'поиск'       
     DO addButtonOne WITH 'fRasp','menuCont8',.menucont7.Left+.menucont7.Width+3,5,'возврат','undo.ico','Do exittotalshtat',39,.menucont1.Width,'возврат'       
          
     .Caption='Штатное расписание' 
     .AddObject('fGrid','gridMyNew')
     WITH .fGrid        
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Top-dHeight          
          .Width=.Parent.Width/2   
           DO addColumnToGrid WITH 'fRasp.fGrid',5

          .RecordSourceType=1
          .RecordSource='currasptot'
          .ScrollBars=2
          .Column1.RemoveObject('Header1')
          .Columns(1).AddObject('Header1','HeaderMy')
          .Columns(1).Header1.procForClick='DO procFilterDolTot' 
          .Column1.Sparse=.F.            
          .Column1.ControlSource='currasptot.nd'
          .Column2.ControlSource='currasptot.named'
          .Column3.ControlSource='currasptot.kse'                  
          .Column4.ControlSource='currasptot.ksevaccl'                                         
          .Column1.Header1.Caption='+'
          .Column2.Header1.Caption='Наименование' 
          .Column3.Header1.Caption='к-во'         
          .Column4.Header1.Caption='+/-'
                           
          .Column1.Width=RettxtWidth(' 1234 ')   
          .Column3.Width=RetTxtWidth('999999.99')        
          .Column4.Width=.Column3.Width
          .Columns(.ColumnCount).Width=0
          *.Column2.Width=.Width-.column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-SYSMETRIC(5)-13-.ColumnCount        
          .Column2.Width=.Width-.column1.Width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount        
          .Column3.Alignment=1             
          .Column4.Alignment=1
          .Column5.Alignment=1       
          .Column4.Format='Z'
          .Columns(.ColumnCount-2).Format='Z'        
          .Column1.Alignment=2         
          .Column2.Alignment=0
          .Columns(.ColumnCount-1).Alignment=0
          .colNesInf=2             
          .Column1.AddObject('cont1','checkContainer')                        
          .Column1.cont1.AddObject('tBox1','textBoxAsCont')                           
          .Column1.cont1.tBox1.Left=0
          .Column1.cont1.tBox1.Width=.Column1.Width
          .Column1.cont1.tBox1.Height=.rowHeight
          .Column1.cont1.tBox1.BackStyle=0
          .Column1.cont1.tBox1.Alignment=2
          .Column1.cont1.tBox1.BorderStyle=0
          .Column1.cont1.tBox1.ControlSource='curRaspTot.col1'                                                                                                          
          .Column1.cont1.ProcForClick='DO procFilterDolPodr' 
          .Column1.cont1.tBox1.Enabled=.F.                                                                                      
          .Column1.AddObject('cont2','checkContainer')                        
          .Column1.cont2.AddObject('tBox1','MyTxtBox')
          .Column1.cont2.tBox1.Visible=.T.             
          .Column1.cont2.tBox1.Left=0
          .Column1.cont2.tBox1.Width=.Column1.Width
          .Column1.cont2.tBox1.Height=.rowHeight
          .Column1.cont2.tBox1.BackStyle=0
          .Column1.cont2.tBox1.BorderStyle=0
          .Column1.cont2.tBox1.Enabled=.F.
          .Column1.cont2.tBox1.ControlSource='curRaspTot.nd'  
          .column1.DynamicCurrentControl="IIF(curRaspTot.log_s,'cont1','cont2')"
          .column1.Enabled=.T.           
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')     
          .Column1.Sparse=.F.  
          .ProcAfterRowColChange='DO peopleincursor'    
     ENDWITH    
     DO gridSize WITH 'fRasp','fGrid','shapeingrid',.T.,.F. 
     DO adTboxAsCont WITH 'fRasp','txtitog1',.fGrid.Left,.fGrid.Top+.fGrid.height-1,.fGrid.column1.Width+12,dHeight,'',1,1
     DO adTboxAsCont WITH 'fRasp','txtitog2',.txtitog1.Left+.txtItog1.Width-1,.txtItog1.Top,.fGrid.column2.Width+2,dHeight,'всего',0,1
     DO adTboxAsCont WITH 'fRasp','txtitog3',.txtitog2.Left+.txtItog2.Width-1,.txtItog1.Top,.fGrid.column3.Width+2,dHeight,totItog,0,1
     DO adTboxAsCont WITH 'fRasp','txtitog4',.txtitog3.Left+.txtItog3.Width-1,.txtItog1.Top,.fGrid.column4.Width+2,dHeight,totVac,0,1
     DO adTboxAsCont WITH 'fRasp','txtitog5',.txtitog4.Left+.txtItog4.Width-1,.txtItog1.Top,.fGrid.column5.Width+2,dHeight,totClvac,0,1
     
     fGridHeight=.fGrid.Height
     FOR i=1 TO .fGrid.ColumnCount         
         .fGrid.Columns(i).fontname=dFontName
         .fGrid.Columns(i).fontSize=dFontSize      
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,dBackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicFontBold='IIF(log_s,.T.,.F.)'                                       
         .fGrid.Columns(i).DynamicBackColor='IIF(!currasptot.log_s,IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,dBackColor,dynBackColor),;
                                                   IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,headerBackColor,dynBackColor))'                                                                                                                                                                                                      
                                                   
         .fGrid.Columns(i).Resizable=.F.           
         .fGrid.Columns(i).Text1.SelectedForeColor=dynForeColor
         .fGrid.Columns(i).Text1.SelectedBackColor=dynBackColor
         .fGrid.Columns(i).Header1.ForeColor=dForeColor
         .fGrid.Columns(i).Header1.Alignment=2
         .fGrid.Columns(i).Header1.FontName=dFontName
         .fGrid.Columns(i).Header1.FontSize=dFontSize
         .fGrid.Columns(i).Header1.BackColor=headerBackColor
         .fGrid.Columns(i).Enabled=IIF(i=1,.T.,.F.)                           
     ENDFOR    
     
     .fGrid.Column4.DynamicBackColor='IIF(currasptot.ksevaccl<0,RGB(255,0,0),IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,dBackColor,dynBackColor))'  
     
     
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .fGrid.Column1.AddObject('contColumn1','checkContainer')  
     .fGrid.Column1.contColumn1.AddObject('labLog','textBoxAsCont')     
     WITH .fGrid.Column1.contColumn1.lablog  
          .Visible=.T.   
          .Height=fRasp.fGrid.rowHeight
          .Width=fRasp.fGrid.Column1.Width
          .Top=0 
          .Left=0
          .BackStyle=0
          .BorderStyle=0
          .Alignment=2
          .ControlSource='currasptot.col1' 
          .Enabled=.F.       
      ENDWITH   
     .fGrid.column1.CurrentControl='contColumn1' 
     .fGrid.Column1.contColumn1.ProcForKeyPress='DO procKeyPresGrid'          
     .fGrid.Column1.ContColumn1.ProcForClick='DO procFilterDolPodr'  
     .AddObject('gridPeop','gridMyNew')
     WITH .gridPeop
          .Top=fRasp.fGrid.Top
          .Left=fRasp.fGrid.Left+fRasp.fGrid.Width+10
          .Height=fRasp.fGrid.Height
          .Width=fRasp.Width-fRasp.fGrid.Width        
          *.ColumnCount=7       
          .RecordSourceType=1
          .RecordSource='curRaspPeople'
          .ScrollBars=2
           DO addColumnToGrid WITH 'fRasp.gridPeop',8
          .Column1.ControlSource='curRaspPeople.kodPeop'
          .Column2.ControlSource='curRaspPeople.fio'
          .Column3.ControlSource='curRaspPeople.nameDol' 
          .Column4.ControlSource='curRaspPeople.kse'                                   
          .Column5.ControlSource='curRaspPeople.nameTr'
          .Column6.ControlSource='curRaspPeople.primtxt'
          .Column7.ControlSource=''
                 
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Фамилия имя Отчество'
          .Column3.Header1.Caption='должность'
          .Column4.Header1.Caption='объём'   
          .Column5.Header1.Caption='тип' 
          .Column6.Header1.Caption='примечание'    
          .Column7.header1.Caption='!'
                    
          .Column1.Width=RettxtWidth('w1234w')
          .Column4.Width=RettxtWidth('99999')
          .Column5.Width=RetTxtWidth('внеш.совм.')
          .Column6.Width=RetTxtWidth('декретный отпуск на время')
          .column7.Width=.rowHeight+4
          .Columns(.ColumnCount).Width=0
          .Column4.Alignment=1   
          .Column5.Alignment=0
          .Column6.Alignment=0
          .Column4.Format='Z'                
          .Column3.Width=(.Width-.column1.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width)/2
          .Column2.Width=.Width-.column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-SYSMETRIC(5)-13-.ColumnCount        
                         
          .Column2.Alignment=0         
          .Column3.Alignment=0  
          .SetAll('Enabled',.F.,'Column')         
          .colNesInf=2   
          
          .Column7.ReadOnly=.T.                                
          .Column7.AddObject('checkColumn7','checkContainer')
          .Column7.checkColumn7.AddObject('butEdit','myCommandButton')
          .Column7.CheckColumn7.butEdit.Visible=.T.
          .Column7.CheckColumn7.butEdit.Caption=''
          .Column7.CheckColumn7.butEdit.Left=2
          .Column7.CheckColumn7.butEdit.Top=2
          .Column7.CheckColumn7.butEdit.Height=.rowHeight-4
          .Column7.CheckColumn7.butEdit.Width=.column7.Width-4
          .Column7.CheckColumn7.butEdit.procForClick='DO formPeopInfoRasp'
         * .Column8.CheckColumn8.checkMy.BackStyle=0        
          *.Column8.CheckColumn8.checkMy.ControlSource='lreadec'                                                                                                  
          .column7.CurrentControl='checkColumn7'       
          .SetAll('Enabled',.F.,'ColumnMy')
          .Column7.Enabled=.T. 
          .Column7.Sparse=.F. 
          
          
          
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')   
        *  .Visible=.F.          
     ENDWITH     
     DO gridSizeNew WITH 'fRasp','gridPeop','shapeingrid1',.T. 
     
     DO MyColumntxtBox WITH 'fRasp.fgrid.Columns(frasp.fGrid.ColumnCount)','tbox4','',.F.,.F.,''        
     .fGrid.Columns(.fGrid.ColumnCount).tbox4.procForKeyPress='DO keyPresGridShtat' 
     SELECT currasptot    
ENDWITH
fRasp.Show
**************************************************************************************************************************
PROCEDURE procClickColumn1
SELECT currasptot
IF log_s
   dolrec=RECNO()
   kodpodr=kp
   SET FILTER TO 
   REPLACE log_sp WITH IIF(log_sp,.F.,.T.) FOR kp=kodpodr.AND.RECNO()#dolrec
   SET FILTER TO log_sp
   GO dolrec
   fRasp.Refresh
ENDIF 
**************************************************************************************************************************
PROCEDURE procFilterDolTot
log_spis=IIF(log_spis,.F.,.T.)
SELECT currasptot
SET FILTER TO 
IF log_spis 
   SELECT rasp
   REPLACE log_sp WITH .T. ALL
   SELECT currasptot
   REPLACE log_sp WITH .T. ALL
   fRasp.fGrid.Column1.Header1.Caption='-'
   REPLACE col1 WITH IIF(log_s,'-','') ALL   
ELSE
   SELECT rasp
   REPLACE log_sp WITH .F. ALL
   SELECT currasptot
   REPLACE log_sp WITH IIF(log_s,.T.,.F.) ALL 
   REPLACE col1 WITH IIF(log_s,'+','') ALL 
   fRasp.fGrid.Column1.Header1.Caption='+'  
ENDIF   
SET FILTER TO log_sp
GO TOP 
fRasp.Refresh
***************************************************************************************************************************
PROCEDURE procFilterDolPodr
SELECT currasptot
IF !log_s
   RETURN 
ENDIF
dolrec=RECNO()
kodpodr=kp
SET FILTER TO 
REPLACE log_sp WITH IIF(log_sp,.F.,.T.) FOR kp=kodpodr.AND.RECNO()#dolrec
LOCATE FOR kp=kodpodr.AND.log_s
REPLACE col1 WITH IIF(col1='-','+','-')
SET FILTER TO log_sp
GO dolrec
fRasp.Refresh
**************************************************************************************************************************
PROCEDURE keyPresShtat
DO CASE
   CASE LASTKEY()=27
       * DO writeRasp WITH .F.        
   CASE LASTKEY()=23 &&ctrl+W  
       * DO writeRasp WITH .T.        
ENDCASE
**************************************************************************************************************************
PROCEDURE keyPresGridShtat
DO CASE
   CASE LASTKEY()=14 &&ctrl+N
       * Do readshtat WITH .T.
   CASE LASTKEY()=18 &&ctrl+R
       * Do readshtat WITH .F.
   CASE LASTKEY()=147 &&ctrl+Del 
       *  DO deleteFromShtat
ENDCASE

*************************************************************************************************************************
*                     Новое подразделение в штатном расписании
*************************************************************************************************************************
PROCEDURE newPodrShtat
PARAMETERS parLog
logApPodr=parLog
SELECT sprpodr
SET ORDER TO 1
GO BOTTOM
kodNewPodr=kod+1      
SET ORDER TO 3
SELECT currasptot
oldrec=RECNO()
kodPodrOld=kp
numPodrOld=np
nameNewPodr=SPACE(100)
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Icon='money.ico'
     .Caption='Ввод нового подразделения'
     DO addShape WITH 'fSupl',1,20,20,dHeight,380,8              
     DO adtBoxAsCont WITH 'fSupl','contNum',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('wкодw'),dHeight,'код',2,1     
     DO adtBoxAsCont WITH 'fSupl','contName',.contNum.Left+.contNum.Width-1,.contNum.Top,450,dHeight,'наименование подразделения',2,1  
     DO adtBoxNew WITH 'fSupl','txtBoxNum',.contNum.Top+.contNum.Height-1,.contNum.Left,.contNum.Width,dHeight,'kodNewPodr','Z',.F.,.F.,.F.
     DO adtBoxNew WITH 'fSupl','txtBoxname',.txtBoxNum.Top,.contName.Left,.contName.Width,dHeight,'nameNewPodr',.F.,.T.,.F.,.F.                      
     .Shape1.Height=dHeight*2-1+40
     .Shape1.Width=.ContNum.Width+.ContName.Width+40           
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wwзаписатьww')*2-20)/2,.Shape1.Top+.Shape1.Height+20,;
        RetTxtWidth('wwзаписатьww'),dHeight+3,'записать','DO writePodrInRasp WITH .T.'
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','DO writePodrInRasp WITH .T.'
     .Height=.Shape1.Height+.cont1.Height+60  
     .Width=.Shape1.Width+40
 ENDWITH
 DO pasteImage WITH 'fSupl'
 fSupl.Show       
*************************************************************************************************************************
*                      Запись нового подразделения в штатное расписание
*************************************************************************************************************************
PROCEDURE writePodrInRasp
PARAMETERS par1
fSupl.Release
logApPodr=.F.
IF !par1
   RETURN
ENDIF
IF EMPTY(nameNewPodr)
   RETURN 
ENDIF
SELECT sprpodr
REPLACE np WITH np+1 FOR np>curRaspTot.np
APPEND BLANK
REPLACE kod WITH kodNewPodr,name WITH nameNewPodr,np WITH numPodrOld+1,namework WITH nameNewPodr
SELECT curRaspTot
SET FILTER TO 
REPLACE np WITH np+1 FOR np>numPodrOld
IF RECCOUNT('currasptot')#0
   GO oldrec
ENDIF    
APPEND BLANK           
numrec=RECNO()
REPLACE log_s WITH .T.,log_sp WITH .T.,np WITH numPodrOld+1,kp WITH kodNewPodr,col1 WITH '+',named WITH nameNewPodr
SET FILTER TO log_sp
fRasp.Refresh  
*************************************************************************************************************************
PROCEDURE readShtat
PARAMETERS par1
log_ap=par1
ndOld=0
SELECT curRaspTot
numRec=RECNO()
newNd=0
kpold=curRaspTot.kp
kpsuplold=curRaspTot.kpsupl
npOld=curRaspTot.np
IF log_ap   
   kse_new=00.00
   nd_new=nd
   kd_new=kd
   kat_new=kat
   strdol=''      
   strkat=''
ELSE 
   IF !log_s     
      kse_new=kse             
      nd_new=nd
      ndOld=nd
      kd_new=kd
      kat_new=kat
      strdol=currasptot.named   
      strkat=currasptot.namekat 
   ELSE
   ENDIF 
ENDIF 
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     IF !log_ap.AND.curRaspTot.log_s 
        oldNamePodr=currasptot.named
        .Caption='Редактирование подразделения'
        DO adtBoxAsCont WITH 'fSupl','contName',10,10,550,dHeight,'наименование подразделения',2,1     
        DO adtBoxNew WITH 'fSupl','boxName',.contName.Top+.contName.Height-1,.contName.Left,.contName.Width,dHeight,'oldNamePodr',.F.,.T.,.F.,.F.                          
        .Width=.contName.Width+20    
        butTop=.boxName.Top+.boxName.Height+20      
     ELSE
        IF log_ap
           SELECT curRaspTot           
           oldrec=RECNO()
           SET FILTER TO 
           REPLACE log_sp WITH .T. FOR kp=kpold
           SET FILTER TO log_sp           
           LOCATE FOR kp=kpold.AND.log_s
           SCAN WHILE kp=kpold
                nd_new=nd
           ENDSCAN
           nd_new=nd_new+1                  
           ndOld=nd_new
           GO oldrec
           fRasp.Refresh              
        ENDIF        
        .Icon='money.ico'
        .Caption=IIF(log_ap,'Новая должность','Редактирование должности')
        DO adtBoxAsCont WITH 'fSupl','contNum',10,10,RetTxtWidth('w№w'),dHeight,'№',2,1     
        DO adtBoxAsCont WITH 'fSupl','contDol',.contNum.Left+.contNum.Width-1,.contNum.Top,RetTxtWidth('WWWнаименование должности наименование должностиWW'),dHeight,'наименование должности',2,1     
        DO adtBoxAsCont WITH 'fSupl','contKse',.contDol.Left+.contDol.Width-1,.contNum.Top,RetTxtWidth('99999999.99'),dHeight,'к-во',2,1
        DO adtBoxAsCont WITH 'fSupl','contKat',.contKse.Left+.contKse.Width-1,.contNum.Top,RetTxtWidth('средний медицинский персоналw'),dHeight,'персонал',2,1           
        DO adtBoxNew WITH 'fSupl','txtBox1',.contNum.Top+.contNum.Height-1,.contNum.Left,.contNum.Width,dHeight,'nd_new','Z',.T.,.F.,.F.,'DO validTxtBoxNd'                           
        DO addComboMy WITH 'fSupl',1,.contDol.Left,.txtBox1.Top,dHeight,.contDol.Width,.T.,'strDol','curSprDolj.namework',6,.F.,'DO procValidDoljNew',.F.,.T.
        .comboBox1.DisplayCount=15
        DO addSpinnerMy WITH  'fSupl','spinKse',.contKse.Left,.txtBox1.Top,dHeight,.contKse.Width,'kse_new',0.25
        DO addComboMy WITH 'fSupl',2,.contKat.Left,.txtBox1.Top,dHeight,.contKat.Width,.T.,'strKat','curSupKat.name',6,.F.,'kat_new=curSupKat.kod',.F.,.T. 
        .Width=.contNum.Width+.contDol.Width+.contKse.Width+.contKat.Width-3+20
        butTop=.comboBox1.Top+.comboBox1.Height+20       
     ENDIF    
     *-----------------------------Кнопка записать-----------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',(.Width-(RetTxtWidth('wзаписатьw')*2)-30)/2,butTop,RetTxtWidth('wзаписатьw'),dHeight+5,'записать','DO writeRasp WITH .T.'

     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'возврат','DO writeRasp WITH .F.','возврат' 
     .Height=IIF(!log_ap.AND.curRaspTot.log_s,dHeight*2+.cont1.Height+50,dHeight*2+.cont1.Height+50)
     .Autocenter=.T.    
     .WindowState=0
ENDWITH
*DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validTxtBoxNd
log_Rep=IIF(nd_new#ndOld,.T.,.F.)
*************************************************************************************************************************
*                 Обработка события при редактировании номера должности
*************************************************************************************************************************
PROCEDURE validcolnd
SELECT currasptot
drec=RECNO()
LOCATE FOR kp=kpold.AND.nd=nd_new.AND.drec#RECNO()
log_rep=IIF(FOUND(),.T.,.F.) 
GO drec
*************************************************************************************************************************
PROCEDURE procValidDoljNew
SELECT curRaspTot
kd_new=curSprDolj.kod
kat_new=curSprDolj.kat
strkat=IIF(SEEK(kat_new,'sprkat',1),sprkat.name,'')
fSupl.Refresh
KEYBOARD '{TAB}'
*************************************************************************************************************************
*                       Процедура удаления "дырок" подразделении
*************************************************************************************************************************
PROCEDURE deleteHoleInPodr
SELECT currasptot
LOCATE FOR kp=kpold
newnd=1
DO WHILE kp=kpold.AND.!EOF()
   IF !log_s
      REPLACE nd WITH newnd      
      IF SEEK(STR(currasptot.kp,3)+STR(currasptot.kd,3),'rasp',2)
         REPLACE rasp.nd WITH currasptot.nd
      ENDIF
      SELECT currasptot
      newnd=newnd+1
   ENDIF    
   SKIP   
ENDDO
SELECT currasptot

*************************************************************************************************************************
*                       Процедура удаления "дырок" подразделении
*************************************************************************************************************************
PROCEDURE deleteHoleInShtat
SELECT sprpodr
SET ORDER TO 3
GO TOP
newnp=1
DO WHILE !EOF()   
   REPLACE np WITH newnp  
   newnp=newnp+1     
   SKIP   
ENDDO
*************************************************************************************************************************
*              Процедура записи изменений в штатном расписании
*************************************************************************************************************************
PROCEDURE writerasp
PARAMETERS par_log
fSupl.Visible=.F.
IF !par_log
   fSupl.Release
   RETURN 
ENDIF
IF par_log
   DO CASE
      CASE currasptot.log_s.AND.!log_Ap          
           SELECT sprpodr
           LOCATE FOR kod=currasptot.kp
           REPLACE name WITH oldnamepodr,namework WITH name       
           SELECT currasptot
           GO numrec
           REPLACE named WITH oldnamepodr
      CASE !currasptot.log_s.OR.(curRaspTot.log_s.AND.log_ap) 
           IF log_rep
              SELECT curRaspTot
              SET FILTER TO kp=kpold.AND.!log_s
              IF nd_new>ndOld
                 SET ORDER TO 
                 REPLACE nd WITH nd-1 FOR nd<=nd_new                 
              ENDIF    
              IF nd_new<ndOld
                 SET ORDER TO 
                 REPLACE nd WITH nd+1 FOR nd>=nd_new                             
              ENDIF
           ENDIF
           SET FILTER TO 
           SET FILTER TO log_sp
           SET ORDER TO 1
           GO numrec               
           SELECT rasp
           IF log_ap
              APPEND BLANK
              REPLACE kp WITH kpold,kpsupl WITH kpsuplold
           ELSE
              LOCATE FOR kp=currasptot.kp.AND.kd=currasptot.kd
           ENDIF   
           REPLACE kd WITH kd_new,kse WITH kse_new,nd WITH nd_new,kat WITH kat_new
           IF !log_ap.AND.kat_new#currasptot.kat
              SELECT datjob
              REPLACE kat WITH kat_new FOR kp=rasp.kp.AND.kd=rasp.kd
           ENDIF          
           SELECT currasptot
           IF log_ap
              APPEND BLANK 
              REPLACE kp WITH kpOld,np WITH npOld,log_sp WITH .T.,kpsupl WITH kpsuplold
              numrec=RECNO()
           ENDIF
           GO numRec
           REPLACE kd WITH kd_new,kse WITH kse_new,nd WITH nd_new,named WITH strdol,kat WITH kat_new,namekat WITH strkat,kse_vac WITH kse-kse_peop,ksevaccl WITH kse-ksepeopcl                            
           DO deleteHoleInPodr    
           SELECT rasp
           SET ORDER TO 2
           SELECT currasptot
           LOCATE FOR kp=kpold
           SCAN WHILE kp=kpold
                IF !log_s
                   SELECT rasp 
                   SEEK STR(currasptot.kp,3)+STR(currasptot.kd,3)
                   REPLACE nd WITH currasptot.nd
                   SELECT currasptot
                ENDIF    
           ENDSCAN
           SELECT rasp
           SET ORDER TO 1
           SELECT currasptot
           GO numrec           
   ENDCASE 
   SELECT curRaspTot
   kpOld=curRaspTot.kp
   kpSuplOld=curRaspTot.kpsupl
   numrec=RECNO()
   IF kpOld=kpSuplOld
      SELECT rasp
      SUM kse TO ksePodr FOR kpsupl=kpsuplold
      SELECT curRaspTot
      LOCATE FOR kp=kpOld.AND.log_s
      REPLACE kse WITH ksePodr,kse_vac WITH kse-kse_peop,ksevaccl WITH kse-ksepeopcl       
   ELSE
      SELECT rasp
      SUM kse TO ksePodr FOR kpsupl=kpsuplold
      SELECT curRaspTot
      LOCATE FOR kp=kpSuplOld.AND.log_s
      REPLACE kse WITH ksePodr,kse_vac WITH kse-kse_peop,ksevaccl WITH kse-ksepeopcl     
      SELECT rasp
      SUM kse TO ksepodr1 FOR kp=kpold
      SELECT curRaspTot
      LOCATE FOR kp=kpOld.AND.log_s
      REPLACE kse WITH ksePodr1
      
      *,kse_vac WITH kse-kse_peop,ksevaccl WITH kse-ksepeopcl       
   ENDIF  
   SUM kse TO totItog FOR log_s.AND.kodkpp=0  
   GO numRec     
   fSupl.Release
ENDIF
SELECT curRaspTot
fRasp.fGrid.Column1.SetFocus
fRasp.txtItog3.ControlSource='totItog'
fRasp.Refresh
**************************************************************************************************************************
*                       Построение формы процедуры удаления из штатного
**************************************************************************************************************************
PROCEDURE deleteFromShtat
SELECT datjob
SET FILTER TO 
IF currasptot.log_s
   logDelpeop=IIF(SEEK(STR(curRaspTot.kp,3),'datjob',2),.T.,.F.)
ELSE
   logDelpeop=IIF(SEEK(STR(curRaspTot.kp,3)+STR(curRaspTot.kd,3),'datjob',2),.T.,.F.)
ENDIF
SELECT currasptot
log_del=.F.
fDel=CREATEOBJECT('formsupl')
WITH fDel
     .Icon='money.ico'
     .Caption='Удаление из штатного расписания'
     .Width=400
     .Height=200
     
     DO addShape WITH 'fdel',1,20,20,dHeight,RetTxtWidth('WВ выбранном подразделении находятся сотрудники!W'),8   
     DO adLabMy WITH 'fdel',1,IIF(log_s,'Удаление подразделения','Удаление должности'),.Shape1.Top+10,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fdel',2,'Для подтверждения намерений поставьте птичку',.Lab1.Top+.Lab1.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fdel',3,"в окошке 'подтверждение намерений'",.Lab2.Top+.Lab2.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     IF logDelPeop
        DO adLabMy WITH 'fdel',4,IIF(log_s,'В выбранном подразделении находятся сотрудники!','На выбранной должности находятся сотрудники!'),.Lab3.Top+.Lab3.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
        .Shape1.Height=.lab1.Height*4+35       
     ELSE   
        .Shape1.Height=.lab1.Height*3+30       
     ENDIF   
      DO adCheckBox WITH 'fDel','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'log_del',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wwудалитьww')*2-20)/2,.check1.Top+.check1.Height+20,RetTxtWidth('wwудалитьww') ,dHeight+3,'удалить','DO delrecFromShtat'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fdel.Release'
     .Height=.Shape1.Height+.cont1.Height+.check1.Height+80
     .Width=.Shape1.Width+40   
ENDWITH
DO pasteImage WITH 'FDel'
fDel.Show
**************************************************************************************************************************
PROCEDURE delrecFromShtat
IF !log_del
   RETURN 
ENDIF
fDel.Visible=.F.
fDel.Release
kpdel=currasptot.kp
DO CASE
   CASE currasptot.log_s.AND.currasptot.kd=0
        IF logDelPeop
           SELECT datjob
           REPLACE kp WITH 0 FOR kp=curRasptot.kp 
        ENDIF
        SELECT rasp
        DELETE FOR kp=currasptot.kp   
        SELECT currasptot
        kpold=kp        
        SELECT sprpodr
        SET ORDER TO 1
        SEEK currasptot.kp        
        DELETE         
        SET ORDER TO 3   && NP
        DO deleteHoleInShtat 
        SELECT curRaspTot
        SET FILTER TO 
        DELETE FOR kp=kpold
        SET FILTER TO log_sp
        GO TOP 
        fRasp.Refresh       
   CASE !currasptot.log_s  
        IF logDelPeop
           SELECT datjob
           REPLACE kd WITH 0 FOR kp=curRasptot.kp.AND.kd=currasptot.kd 
        ENDIF 
        kpsuplold=currasptot.kpsupl
        kpold=currasptot.kp
        ksepodr=0
        ksepodr1=0
        SELECT rasp
        IF SEEK(STR(currasptot.kp,3)+STR(currasptot.kd,3),'rasp',2)
           DELETE
        ENDIF        
        *SUM kse TO ksePodr FOR kp=currasptot.kp
        SUM kse TO ksePodr FOR kpsupl=currasptot.kpsupl
        SELECT currasptot
        DELETE                        
        LOCATE FOR kp=kpOld.AND.log_s
        REPLACE kse WITH ksePodr                              
        DO deleteHoleInPodr       
ENDCASE
SELECT currasptot
numrec=RECNO()
SUM kse TO totItog FOR log_s.AND.kodkpp=0  
GO numRec  
fRasp.fGrid.Column1.SetFocus
fRasp.txtItog3.ControlSource='totItog'
fRasp.Refresh
**************************************************************************************************************************
*                               Процедура перенумерации подразделений
**************************************************************************************************************************
PROCEDURE procNumPodr  
fPodr=CREATEOBJECT('FORMSPR')
log_kpp=1
log_kpp1=0
log_repnum=.F.
SELECT sprpodr
SELECT * FROM sprpodr INTO CURSOR curSprPodr READWRITE
SELECT curRaspTot
SET ORDER TO 1
SELECT curSprPodr
INDEX ON np TAG T1
GO TOP
npold=0
fpodr=CREATEOBJECT('FORMSPR')
WITH fpodr
     .Icon='money.ico'
     .Caption='Нумерация подразделений'
     .procexit='DO exitreppodr'
     primold=''      
     DO addButtonOne WITH 'fPodr','menuCont1',10,5,'редакция','pencil.ico','Do readnumpodr',39,RetTxtWidth('редакция')+44,'редакция'          
     DO addButtonOne WITH 'fPodr','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'возврат','undo.ico','Do exitreppodr',39,.menucont1.Width,'возврат'  
     DO addmenureadspr WITH 'fpodr','DO writenumpodr WITH .T.','DO writenumpodr WITH .F.'
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Left=0
          .Width=fpodr.Width
          .Height=fpodr.Height-fpodr.menucont1.Height-5        
          .ScrollBars=2              
          .RecordSourceType=1
          .RecordSource='curSprPodr' 
           DO addColumnToGrid WITH 'fPodr.fGrid',6
          .Column1.ControlSource='curSprPodr.np'    
          .Column2.ControlSource='IIF(curSprPodr.kpp1,"            ",IIF(curSprPodr.kpp,"      ","  "))+cursprpodr->name'
          .Column3.ControlSource='curSprPodr.primhead'
          .Column4.ControlSource="IIF(curSprPodr.kpp,'•','')" 
          .Column5.ControlSource="IIF(curSprPodr.kpp1,'•','')"    
     
          .Column1.Width=FONTMETRIC(6,dFontName,dFontSize)* TXTWIDTH(' 1234 ',dFontName,dFontSize)
          .Column4.Width=.Column1.Width
          .Column5.Width=.Column4.Width
          .Column2.Width=(.Width-.Column1.Width-.Column4.Width-.Column5.Width)/2
          .Column3.Width=.Width-.column1.Width-.Column2.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount          
          .Columns(.ColumnCount).Width=0   
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Наименование подразделения'          
          .Column3.Header1.Caption='Наименование совокупности'
          .Column4.Header1.Caption=''
          .Column5.Header1.Caption=''
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=2
          .Column5.Alignment=2
           .colNesInf=2    
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')        
     ENDWITH 
     DO gridSizeNew WITH 'fpodr','fGrid','shapeingrid' 
     FOR i=1 TO fpodr.fGrid.ColumnCount                                               
         fPodr.fGrid.Columns(i).DynamicBackColor='IIF(curSprPodr.kpp.OR.curSprPodr.kpp1,IIF(RECNO(fPodr.fGrid.RecordSource)#fPodr.fGrid.curRec,RGB(255,255,255),dynBackColor),;
                                                  IIF(RECNO(fPodr.fGrid.RecordSource)#fPodr.fGrid.curRec,dBackColor,dynBackColor))'              
     ENDFOR      
          
     DO myColumnTxtBox WITH 'fpodr.fGrid.column6','txtbox6','',.F.
     .fGrid.column6.txtbox6.procForKeyPress='DO repotmpodr' 
     DO adtbox WITH 'fpodr',1,1,1,fpodr.fGrid.Column1.Width+2,.F.,.F.,0,.T.
     DO adtBox WITH 'fPodr',3,1,1,fPodr.fGrid.Column3.Width+2,.F.,.F.,0,.T. 
     .SetAll('Visible',.F.,'MyTxtBox')
     .txtbox1.procForValid='DO validnp'
   
      DO addOptionButton WITH 'fpodr',1,'',.menucont1.Top,.fGrid.Column1.Width+.fGrid.Column2.Width+.fGrid.Column3.Width+13,'log_kpp',0,'DO proclogkpp WITH 1',.T.
     .Option1.Left=.Option1.Left+(.fGrid.Column4.Width-.Option1.Width)/2+3
     .Option1.Top=.fGrid.Top+(.fGrid.HeaderHeight-.Option1.Height)/2
     .Option1.ToolTipText='Редактировать итоги 1'
     DO addOptionButton WITH 'fpodr',2,'',.Option1.Top,.Option1.Left+.fGrid.Column4.Width+1,'log_kpp1',0,'DO proclogkpp WITH 2',.T. 
     .Option2.ToolTipText='Редактировать итоги 2'     
ENDWITH 
fPodr.Show
*******************************************************************************************
PROCEDURE proclogkpp
PARAMETERS parch
DO CASE
   CASE parch=1
        log_kpp=1
        log_kpp1=0
        fpodr.fGrid.column6.txtbox6.procForKeyPress='DO repotmpodr' 
   CASE parch=2
        log_kpp=0
        log_kpp1=1
        fpodr.fGrid.column6.txtbox6.procForKeyPress='DO repotmpodr1' 
ENDCASE
fpodr.Refresh
*******************************************************************************************
PROCEDURE repotmpodr
IF LASTKEY()=13   
   REPLACE kpp WITH IIF(kpp,.F.,.T.),kodkpp WITH IIF(!kpp,0,kodkpp)   
ENDIF
*******************************************************************************************
PROCEDURE repotmpodr1
IF LASTKEY()=13   
   REPLACE kpp1 WITH IIF(kpp1,.F.,.T.),kodkpp1 WITH IIF(!kpp1,0,kodkpp1)
ENDIF
**************************************************************************************************************************
PROCEDURE readnumpodr
log_repnum=.F.
SELECT curSprPodr
WITH fpodr
     .fGrid.Columns(.fGrid.columnCount).SetFocus
     .fGrid.Enabled=.F.
     .SetAll('Enabled',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .menuread.Visible=.T.
     .menuexit.Visible=.T.
     .SetAll('Visible',.T.,'MyTxtBox')
     npold=cursprpodr.np
     primold=cursprpodr.primhead
     .txtBox1.Left=.fGrid.Left+10
     .txtBox3.Left=.txtBox1.Left+.txtBox1.Width+.fGrid.Column2.Width
     .txtBox1.ControlSource='cursprpodr.np'     
     .txtBox3.ControlSource='cursprpodr.primhead'
     .txtBox1.Top=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtBox3.Top=.txtBox1.Top
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .txtbox1.SetFocus
ENDWITH 

**************************************************************************************************************************
*                         Контроль при редакции номера подразделения
**************************************************************************************************************************
PROCEDURE validnp
IF cursprpodr.np#npold
   SELECT cursprpodr
   prec=RECNO()
   npnew=cursprpodr.np
   LOCATE FOR np=npnew.AND.RECNO()#prec
   IF FOUND()
      GO prec         
      log_repnum=.T.                     
      SELECT cursprpodr       
   ELSE
      GO prec
      log_repum=.F.      
   ENDIF
ENDIF   

************************************************************************************************************************
*        Запись (отказ от) изменений при редактировании нумерации подразделений
************************************************************************************************************************
PROCEDURE writenumpodr
PARAMETERS par_log
SELECT cursprpodr
ordOld=SYS(21)
nrec=RECNO()
IF par_log
   DO CASE
      CASE log_repnum=.T.           
           SELECT cursprpodr
           nrec=RECNO()
           npnew=cursprpodr.np
           SET ORDER TO  
           DO CASE
              CASE cursprpodr.np>npold        
                   REPLACE np WITH 0
                   REPLACE np WITH np-1 FOR np<=npnew
                   GO nrec
                   REPLACE np WITH npnew  
                   SET ORDER TO &ordOld     
              CASE cursprpodr.np<npold                                   
                   REPLACE np WITH 0
                   REPLACE np WITH np+1 FOR np>=npnew
                   GO nrec
                   REPLACE np WITH npnew                   
                   SET ORDER TO &ordOld
           ENDCASE
           
      CASE log_repnum=.F.
           REPLACE np WITH npold
   ENDCASE
   GO TOP
   numnew=1
   DO WHILE !EOF()
      REPLACE np WITH numnew
      SKIP
      numnew=numnew+1
   ENDDO
   COUNT TO max_podr
   GO TOP            
   SELECT cursprpodr           
   GO nrec
ELSE 
   REPLACE np WITH npold,primhead WITH primold   
ENDIF

WITH fPodr
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     .SetAll('Enabled',.T.,'MyOptionButton')
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH 
fpodr.Refresh
********************************************************************************************************************
PROCEDURE proclogrep
PARAMETERS par_log
log_repnum=par_log
nformmes.Release
************************************************************************************************************************
*                 Удаление "дырок" при перенумерации подразделений
************************************************************************************************************************
PROCEDURE delholepodr
GO TOP
np_new=1
kp_old=kod
DO WHILE !EOF()
   REPLACE np WITH np_new
   SKIP
   IF kod#kp_old
      np_new=np_new+1
      kp_old=kod
   ENDIF
ENDDO

*************************************************************************************************************************
*                              Выход из нумерации подразделений
*************************************************************************************************************************
PROCEDURE exitreppodr
fPodr.Visible=.F.
SELECT curSprPodr
GO TOP
kodRepKpp=kod
kodrepKpp1=kod
SCAN ALL     
     IF kpp 
        IF !kpp1
           kodRepKpp1=kod
        ENDIF    
        REPLACE kodKpp WITH kodRepKpp,kpsup WITH kodKpp
        IF kpp1
           REPLACE kodKpp1 WITH kodRepKpp1
        ENDIF 
     ELSE
        kodRepKpp=kod
        kodRepKpp1=kod
        REPLACE kodkpp WITH 0,kodkpp1 WITH 0,kpsup WITH kod  
     ENDIF 
     IF SEEK(cursprpodr.kod,'sprpodr',1)
        SCATTER TO dim_podr
        SELECT sprpodr
        GATHER FROM dim_podr       
     ENDIF
     SELECT curSprPodr
ENDSCAN
SELECT curRaspTot
SET FILTER TO 
DELETE ALL
SELECT sprpodr
SET ORDER TO 3
GO TOP
DO WHILE !EOF()
   SELECT currasptot
   APPEND BLANK 
   REPLACE kodkpp WITH sprpodr.kodkpp,named WITH IIF(kodkpp=0,sprpodr.name,'      '+sprpodr.name),kp WITH sprpodr.kod,log_s WITH .T.,log_sp WITH .T.,np WITH sprpodr.np,col1 WITH '+'
   SELECT rasp
   SEEK STR(sprpodr.kod,3)
   ksePodr=0
   SCAN WHILE kp=sprpodr.kod
        SELECT currasptot
        APPEND BLANK
        REPLACE named WITH IIF(!EMPTY(sprdolj.namework),sprdolj.namework,sprdolj.name),kse WITH rasp.kse,kat WITH rasp.kat,nd WITH rasp.nd,namekat WITH sprkat.name,;
                kp WITH sprpodr.kod,kd WITH sprdolj.kod,kse_tot WITH rasp.kse_tot ,kse_am WITH rasp.kse_am,;
                kse_stac WITH rasp.kse_stac,kse_sk WITH rasp.kse_sk,kse_mjr WITH rasp.kse_mjr,np WITH sprpodr.np,kodkpp WITH sprpodr.kodkpp,kpsupl WITH IIF(kodkpp=0,kp,kodkpp)
        ksePodr=ksePodr+kse        
        SELECT rasp       
   ENDSCAN	
   SELECT curRaspTot
   LOCATE FOR kp=sprpodr.kod.AND.log_s
   REPLACE kse WITH ksePodr
   SELECT sprpodr
   SKIP 
ENDDO   
SET ORDER TO 1
SELECT rasp
REPLACE kpsupl WITH IIF(SEEK(kp,'sprpodr',1).AND.sprpodr.kodkpp=0,sprpodr.kod,sprpodr.kodkpp) ALL 
SELECT currasptot
SET ORDER TO 1
SCAN ALL
     SELECT curSupJob     
     SUM kse TO ksesup FOR IIF(currasptot.log_sp,kp=currasptot.kp.AND.kd#0.AND.!dekotp,kp=currasptot.kp.AND.kd=currasptot.kd.AND.!dekotp)
     SUM kse TO ksesupcl FOR IIF(currasptot.log_sp,kp=currasptot.kp.AND.kd#0.AND.INLIST(tr,1,2,3,5).AND.!dekotp,kp=currasptot.kp.AND.kd=currasptot.kd.AND.INLIST(tr,1,2,3,5).AND.!dekotp)
     SELECT currasptot
     REPLACE kse_peop WITH ksesup,ksepeopcl WITH ksesupcl,kse_vac WITH kse-kse_peop,ksevaccl WITH kse-ksepeopcl     
ENDSCAN

SET FILTER TO log_sp
SELECT sprpodr
SCAN ALL
     kse_ch=0 
     ksevac_ch=0
     ksevaccl_ch=0
     IF kodkpp#0
        SELECT currasptot
        SEEK STR(sprpodr.np,3)
        kse_ch=kse
        ksevac_ch=kse_vac
        ksevaccl_ch=ksevaccl
        LOCATE FOR kp=sprpodr.kodkpp
        REPLACE kse WITH kse+kse_ch,kse_vac WITH kse_vac+ksevac_ch,ksevaccl WITH ksevaccl+ksevaccl_ch 
     ENDIF
     SELECT sprpodr     
ENDSCAN
SELECT currasptot
SET FILTER TO log_sp
GO TOP
fRasp.Refresh
**************************************************************************************************************************************
PROCEDURE exittotalshtat
SELECT rasp
SET RELATION TO 
SELECT people
GO oldPeopRec
frmTop.grdPers.columns(frmTop.grdPers.ColumnCount).SetFocus
fRasp.Release
************************************************************************************************************************
PROCEDURE procnumold
REPLACE nd WITH frasp.ndold
nFormMes.Release
************************************************************************************************************************
PROCEDURE repnum
frasp.log_rep=.T.
nFormMes.Release
*************************************************************************************************************************
PROCEDURE exitrepnum
IF !nFormMes.log_rep
   REPLACE nd WITH frasp.ndold
ENDIF
**************************************************************************************************************************
*                                      Выбор персонала по должности
**************************************************************************************************************************
PROCEDURE peopleincursor
SELECT curSpPeople
ZAP
IF currasptot.log_s
    
   APPEND FROM DBF('curSupJob') FOR kp=currasptot.kp
ELSE
   APPEND FROM DBF('curSupJob') FOR kp=currasptot.kp.AND.kd=currasptot.kd
ENDIF
SELECT curRaspPeople
DELETE ALL 
SELECT curSpPeople 
DELETE FOR !EMPTY(dateOut).AND.dateOut<DATE()
DELETE FOR dateOut>DATE()
DELETE FOR dateBeg>DATE()     
SCAN ALL     
     SELECT curRaspPeople
     APPEND BLANK 
     REPLACE kodPeop WITH curSpPeople.kodpeop,kse WITH curSpPeople.kse,tr WITH curSpPeople.tr,kp WITH curSpPeople.kp,kd WITH curSpPeople.kd,;
             fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,''),namedol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),nameTr WITH IIF(SEEK(tr,'sprtype',1),sprtype.name,''),;
             primtxt WITH IIF(curSpPeople.dekOtp,'д/о',''),primtxt WITH IIF(!EMPTY(curSppeople.fiodek),'д/о '+ALLTRIM(curSpPeople.fiodek),primtxt),dmol WITH people.dmol ,nid WITH curSpPeople.nid
     SELECT curSpPeople
ENDSCAN 
SELECT curRaspPeople
GO TOP
WITH frasp    
     .gridPeop.SetAll('ForeColor',dForeColor,'Header')
     .gridPeop.SetAll('BackColor',headerBackColor,'Header') 
     SELECT currasptot
     .gridPeop.Refresh()  
  
ENDWITH
**************************************************************************************************************************
*                                               Поиск в штатном расписании
**************************************************************************************************************************
PROCEDURE formFindRasp 
SELECT * FROM sprpodr  WHERE SEEK(kod,'currasptot',2) INTO CURSOR curSearchPodr ORDER BY name READWRITE
fSupl=CREATEOBJECT('FORMSUPL')
strSearch=''
strFind=''
kpSearch=0
WITH fSupl
     .Caption='Поиск'
     DO addShape WITH 'fSupl',1,10,20,dHeight,500,8          
     DO addcombomy WITH 'fSupl',1,.Shape1.Left+10,.Shape1.Top+20,dHeight,.Shape1.Width-20,.T.,'strSearch','curSearchPodr.name',6,'','kpSearch=curSearchPodr.kod',.F.,.T.  
          
     .comboBox1.procForRightClick='DO selectSearchPodr' 
     DO addcombomy WITH 'fSupl',11,.comboBox1.Left,.comboBox1.Top,dHeight,.comboBox1.Width,.T.,'strFind','curSearchPodr.name',6,'','kpSearch=curSearchPodr.kod'
     WITH .combobox11
          .SpecialEffect=1 
          .Style=0 
          .procForDropDown='Do dropDownCombo11'        
          .procForMouseDown='DO mouseDown11'                
          .Visible=.F.  
     ENDWITH     
                  
     .Shape1.Height=.comboBox1.Height+40     
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wОтменаw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,;
     RetTxtWidth('wОтменаw'),dHeight+3,'Найти','DO procSearchPodr'           
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fSupl.Release'     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+60
ENDWITH       
DO pasteImage WITH 'fSupl'
fSupl.Show 
********************************************************************************************************************************
PROCEDURE selectSearchPodr
WITH fSupl
     .combobox1.Visible=.F.   
     SET CURSOR ON  
     strfind=''
     .combobox11.Visible=.T.
     .ComboBox11.ControlSource='strfind'
     .combobox11.SetFocus  
ENDWITH      
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE dropDownCombo11
strfind=''
strfind=ALLTRIM(fSupl.combobox11.Text)
IF !EMPTY(strfind)
    SELECT curSearchPodr
    SET FILTER TO UPPER(strfind)$UPPER(curSearchPodr.name)
    fSupl.comboBox11.RowSource='curSearchPodr.name'
    COUNT ALL TO nCount
    GO TOP
    IF nCount#0       
       *fornds.ComboBox22.DisplayCount=nCount      
    ELSE
       SET FILTER TO
       fSupl.combobox11.Visible=.F.
       fSupl.ComboBox1.Visible=.T.  
       KEYBOARD '{ENTER}'
    ENDIF   
ELSE
   fSupl.combobox11.Visible=.F.
   fSupl.ComboBox1.Visible=.T.    
ENDIF
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE mouseDown11
strfind=''
strfind=ALLTRIM(fSupl.combobox11.Text)
IF !EMPTY(strfind)
    SELECT curSearchPodr
    SET FILTER TO UPPER(strfind)$UPPER(curSearchPodr.name)
    fSupl.comboBox11.RowSource='curSearchPodr.name'
    COUNT ALL TO nCount
    GO TOP
    IF nCount#0       
       *fornds.ComboBox22.DisplayCount=nCount  
    ELSE
       SET FILTER TO
       fSupl.combobox11.Visible=.F.
       fSupl.ComboBox1.Visible=.T.  
       KEYBOARD '{ENTER}'
    ENDIF   
ELSE
   fSupl.combobox11.Visible=.F.
   fSupl.ComboBox1.Visible=.T.    
ENDIF

********************************************************************************************************************************
PROCEDURE  procSearchPodr
SELECT curRaspTot
LOCATE FOR kp=kpSearch  
fSupl.Release
fRasp.Refresh 
*************************************************************************************************************************
PROCEDURE validTrNew
PARAMETERS par1
DO CASE
   CASE par1=1    && без категории
        nameKfNew=IIF(SEEK(trNew,'sprkoef',1),sprkoef.name,0)
        fSupl.txtBoxTr.Refresh
   CASE par1=2    && вторая категория
        nameKfNew1=IIF(SEEK(trNew1,'sprkoef',1),sprkoef.name,0)
        fSupl.txtBoxTr1.Refresh
   CASE par1=3    && первая категория
        nameKfNew2=IIF(SEEK(trNew2,'sprkoef',1),sprkoef.name,0)
        fSupl.txtBoxTr2.Refresh
   CASE par1=4    && высшая категория
        nameKfNew3=IIF(SEEK(trNew3,'sprkoef',1),sprkoef.name,0)
        fSupl.txtBoxTr3.Refresh
ENDCASE
***************************************************************************************************************************
PROCEDURE formPeopInfoRasp
SELECT * FROM tarfond WHERE nblock=2 INTO CURSOR curInfoFond READWRITE 
SELECT curInfoFond
INDEX ON num TAG T1
SELECT datjob
SET FILTER TO 
oldOrdJob=SYS(21)
SET ORDER TO 7
DIMENSION diminfo(7,2)
diminfo(1,1)='категория'
diminfo(2,1)='дата присвоения'
diminfo(3,1)='спец. по аттестации'
diminfo(4,1)='дата приема'
diminfo(5,1)='стаж на начало'
diminfo(6,1)='стаж'
diminfo(7,1)='молодой специалист до'

diminfo(1,2)='infokval'
diminfo(2,2)='datjob.nprik'
diminfo(3,2)='datjob.sp'
diminfo(4,2)='datjob.date_in'
diminfo(5,2)='datjob.staj_in'
diminfo(6,2)='datjob.staj_tar'
diminfo(7,2)='curRaspPeople.dmol'

SELECT * FROM datjob WHERE kodpeop=curRaspPeople.kodpeop INTO CURSOR curjobinfopeop READWRITE ORDER BY tr
SELECT * FROM tarfond WHERE nblock=2 INTO CURSOR curInfoFond READWRITE 
SELECT curInfoFond
ALTER TABLE curInfoFond ALTER COLUMN sMname N(8,2)
INDEX ON num TAG T1
SELECT curjobinfopeop
GO TOP 
curRow=0
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .procExit='DO exitPeopleInfoRasp'
     .Caption='('+LTRIM(STR(curRaspPeople.kodpeop))+')   '+ALLTRIM(curRaspPeople.fio)
     .Width=800
     .Height=frmTop.Height-200
     .AddObject('grdJob','GridMyNew')   
     WITH .grdJob
          .Top=0
          .Left=0
          .Width=.Parent.Width
          .Height=.headerHeight+.RowHeight*(RECCOUNT('curjobinfopeop')+1) 
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curjobinfopeop'
          DO addColumnToGrid WITH 'fSupl.grdJob',6
          .Column1.ControlSource="IIF(SEEK(curjobinfopeop.kp,'sprpodr',1),sprpodr.name,'')"
          .Column2.ControlSource="IIF(SEEK(curjobinfopeop.kd,'sprdolj',1),sprdolj.namework,'')"         
          .Column3.ControlSource='curjobinfopeop.kse'
          .Column4.ControlSource="IIF(SEEK(curjobinfopeop.tr,'sprtype',1),sprtype.name,'')"           
          .Column1.Header1.Caption='подразделение' 
          .Column2.Header1.Caption='должность'
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'
          .Column5.Header1.Caption='к'           
       
          .Column3.Width=RetTxtWidth('999.999')  
          .Column4.Width=RetTxtWidth('внеш.совм.')                               
          .Column5.Width=RetTxtWidth('wкw')     
         
          .Columns(.ColumnCount).Width=0
          .Column1.Width=(.Width-.Column3.Width-.Column4.Width-.Column5.Width)/2
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount
         
          .Column5.AddObject('checkColumn5','checkContainer')
          .Column5.checkColumn5.AddObject('checkMy','checkMy')
          .Column5.CheckColumn5.checkMy.Visible=.T.
          .Column5.CheckColumn5.checkMy.Caption=''
          .Column5.CheckColumn5.checkMy.Left=6
          .Column5.CheckColumn5.checkMy.Top=3
          .Column5.CheckColumn5.checkMy.BackStyle=0
          .Column5.CheckColumn5.checkMy.ControlSource='curjobinfopeop.lkv'  
          .Column5.CheckColumn5.checkmy.Left=(.Column5.Width-SYSMETRIC(15))/2                                                                                         
          .column5.CurrentControl='checkColumn5'         
          .procAfterRowColChange='DO changeRowInfoRasp'
          .Column1.Alignment=0
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0
          .SetAll('Enabled',.F.,'Column') 
          .Columns(.ColumnCount).Enabled=.T.            
          .Column5.Sparse=.F.             
     ENDWITH        
     DO gridSizeNew WITH 'fSupl','grdJob','shapeingrid',.F.,.T.
     FOR i=1 TO .grdJob.columnCount      
         .grdJob.Columns(i).Backcolor=.BackColor    
         .grdJob.Columns(i).DynamicBackColor='IIF(RECNO(fSupl.grdJob.RecordSource)#fSupl.grdJob.curRec,fSupl.BackColor,dynBackColor)'
         .grdJob.Columns(i).DynamicForeColor='IIF(RECNO(fSupl.grdJob.RecordSource)#fSupl.grdJob.curRec,dForeColor,dynForeColor)'        
     ENDFOR   
     topCx=.grdJob.Top+.grdJob.Height-1
     DO adTboxAsCont WITH 'fSupl','txtJob',0,topCx,fSupl.Width,dHeight,'',2,1,.T.  
     topCx=.txtJob.Top+.txtJob.Height-1 
     .AddObject('grdOklad','GridMyNew')   
      WITH .grdOklad
          .Top=topCx
          .Left=0       
          .Width=.Parent.Width
          .Height=.Parent.height
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='curInfoFond'       
          DO addColumnToGrid WITH 'fSupl.grdOklad',4
          .Column1.ControlSource='curInfoFond.rec'
          .Column2.ControlSource='ALLTRIM(curInfoFond.sPers)'          
          .Column3.ControlSource='curInfoFond.sMname'
          .Column1.Header1.Caption='тариф' 
          .Column2.Header1.Caption='%'
          .Column3.Header1.Caption='сумма'
          .Column2.Width=RetTxtWidth('WобъёмW')
          .Column3.Width=RetTxtWidth('999999999999')
          .Columns(.ColumnCount).Width=0
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column2.Alignment=2
          .Column3.Alignment=1        
           DO gridSizeNew WITH 'fSupl','grdOklad','shapeingrid1',.F.,.T. 
     ENDWITH  
     FOR i=1 TO .grdOklad.columnCount   
         .grdOklad.Columns(i).Backcolor=.BackColor        
         .grdOklad.Columns(i).DynamicBackColor='IIF(RECNO(fSupl.grdOklad.RecordSource)#fSupl.grdOklad.curRec,fSupl.BackColor,dynBackColor)'
         .grdOklad.Columns(i).DynamicForeColor='IIF(RECNO(fSupl.grdOklad.RecordSource)#fSupl.grdOklad.curRec,dForeColor,dynForeColor)'        
     ENDFOR   
   .AutoCenter=.T.
ENDWITH
SELECT curjobinfopeop
LOCATE FOR nid=curRaspPeople.nid
fSupl.Show
***************************************************************************************************************************
PROCEDURE exitPeopleInfoRasp
SELECT curInfoFond
USE
SELECT tarfond
SET FILTER TO 
SELECT curjobinfopeop
USE 
SELECT datJob
SET FILTER TO 
SET ORDER TO &oldOrdJob
SELECT currasptot
***************************************************************************************************************************
PROCEDURE changeRowInfoRasp  
IF curRow#fSupl.grdJob.curRec
   DO scrPeopleInfoRasp
   curRow=fSupl.grdJob.curRec 
ENDIF
***************************************************************************************************************************
PROCEDURE scrPeopleInfoRasp
SELECT datjob
SEEK curjobinfopeop.nid
WITH fSupl    
     topCx=.grdJob.Top+.grdJob.Height-1
     .txtJob.Top=topCx
     widthInf=.Width/4+1
     leftInf=0
     leftInf1=leftInf+widthInf-1   
     infokval=IIF(SEEK(datjob.kv,'sprkval',1),sprkval.name,'')   
     topCx=.txtJob.Top+.txtJob.Height-1      
     ON ERROR DO erSup       
        FOR i=1 TO 8
            oRemove1='tNameInf'+LTRIM(STR(i))
            oRemove2='tInf'+LTRIM(STR(i))
            .RemoveObject(oRemove1)
            .RemoveObject(oRemove2)
        ENDFOR        
     ON ERROR
     kvoStr=0
     kvoSay=0
     FOR i=1 TO 7
         IF !EMPTY(&diminfo(i,2))
            namecont='tNameInf'+LTRIM(STR(i))
            namecont1='tInf'+LTRIM(STR(i))
            DO adTBoxAsCont WITH 'fSupl',namecont,leftInf,topCx, widthInf,dHeight,diminfo(i,1),1,1
            DO adTBoxAsCont WITH 'fSupl',namecont1,leftInf1,topCx, widthInf-IIF(kvoStr=0,0,1),dHeight,&diminfo(i,2),0,0
            kvoStr=kvoStr+1                                                                       
            leftInf=IIF(kvoStr=2,0,leftInf1+widthInf-1) 
            leftInf1=leftInf+widthInf-1                                   
            topCx=IIF(kvostr=2,topCx+Dheight-1,topCx)
            kvoStr=IIF(kvoStr=2,0,kvoStr)  
            kvoSay=kVoSay+1          
         ENDIF 
     ENDFOR
     IF kvoSay#0.AND.MOD(kvoSay,2)=1     
        namecont='tNameInf8'
        namecont1='tInf8'
        DO adTBoxAsCont WITH 'fSupl',namecont,leftInf,topCx, widthInf,dHeight,'',1,1
        DO adTBoxAsCont WITH 'fSupl',namecont1,leftInf1,topCx, widthInf-1,dHeight,'',0,0
        topCx=topCx+dHeight-1                         
     ENDIF  
     .grdOklad.Top=topCx   
     SELECT tarfond
     labNum=0
     kvoStr=0
ENDWITH 
SELECT curInfoFond
RECALL ALL 
SET FILTER TO
REPLACE spers WITH '',sname WITH '' all
GO TOP
SCAN ALL
     reppl1=fpers
     reppl2=fname
     reppl3=ALLTRIM(sayokl)
     reppl4=IIF(!EMPTY(sayoklm),ALLTRIM(sayoklm),IIF(!EMPTY(sayokl),ALLTRIM(sayokl),''))
     IF !EMPTY(reppl1)
        IF logspers
            REPLACE spers WITH IIF(!EMPTY(&reppl1),ALLTRIM(STR(&reppl1,6,2))+IIF(logp,'%',''),'')          
        ELSE 
           REPLACE spers WITH IIF(!EMPTY(&reppl1),ALLTRIM(STR(&reppl1))+IIF(logp,'%',''),'')
        ENDIF            
     ENDIF      
     IF !EMPTY(reppl2)   
        IF EMPTY(&reppl2)
           REPLACE sname WITH ''
        ELSE 
           IF !EMPTY(reppl3)
              REPLACE sname WITH &reppl3 
           ELSE 
              REPLACE sname WITH LTRIM(STR(&reppl2)) 
           ENDIF 
        ENDIF 
     ENDIF 
     IF !EMPTY(reppl4)
        REPLACE sMname WITH VAL(&reppl4)
     ENDIF
ENDSCAN
SET FILTER TO !EMPTY(sPers).OR.!EMPTY(sname)
GO TOP 
fSupl.Refresh