**********************************************************************************************************************
*                                       Воинский учёт
**********************************************************************************************************************
PUBLIC newUch,newKatU,newKzv,newZv,newVus,newRik,newUchet,strGrup,strZv,strRik,newNTicket,newProfil,newRzp,newDateUch,newDateSn,newPrichSn,newPrimZv,newSpprim,newTelWork
IF !USED('datArmy') 
   USE datArmy ORDER 1 IN 0
ENDIF 
IF !USED('sprtot')
   USE sprtot ORDER 1 IN 0
ENDIF 
SELECT primzv DISTINCT FROM datarmy WHERE !EMPTY(datarmy.primzv) INTO CURSOR curPrimZv READWRITE

SELECT kod,name,namesp FROM sprtot WHERE sprtot.kspr=10 INTO CURSOR curSupGrup READWRITE &&группы воинского учёта
SELECT curSupGrup
INDEX ON kod TAG T1 
SELECT kod,name,namesp FROM sprtot WHERE sprtot.kspr=14 INTO CURSOR curSupZv READWRITE &&воинские звания
SELECT curSupZv
INDEX ON kod TAG T1 
INDEX ON name TAG T2
SET ORDER TO 2

newUch=IIF(SEEK(people.nid,'datarmy',2),datarmy.grupu,0)
strGrup=IIF(SEEK(newUch,'curSupGrup',1),curSupGrup.name,'')
newKatU=IIF(SEEK(people.nid,'datarmy',2),datarmy.katu,0)
newKzv=IIF(SEEK(people.nid,'datarmy',2),datarmy.kzv,'')
newzv=IIF(SEEK(people.nid,'datarmy',2),datarmy.zv,'')
strZv=IIF(SEEK(people.nid,'datarmy',2),datarmy.zv,'')
newRzp=IIF(SEEK(people.nid,'datarmy',2),datarmy.rzp,0)
newProfil=IIF(SEEK(people.nid,'datarmy',2),datarmy.profil,'')
newDateUch=IIF(SEEK(people.nid,'datarmy',2),datarmy.dateUch,CTOD('  .  .    '))
newDateSn=IIF(SEEK(people.nid,'datarmy',2),datarmy.dateSn,CTOD('  .  .    '))
newPrichSn=IIF(SEEK(people.nid,'datarmy',2),datarmy.prichsn,'')
newVus=IIF(SEEK(people.nid,'datarmy',2),datarmy.vus,'')
newRik=IIF(SEEK(people.nid,'datarmy',2),datarmy.rik,'')
strRik=IIF(SEEK(people.nid,'datarmy',2),datarmy.rik,'')
newUchet=IIF(SEEK(people.nid,'datarmy',2),datarmy.uchet,'')
newPrimZv=IIF(SEEK(people.nid,'datarmy',2),datarmy.primzv,'')
newNTicket=IIF(SEEK(people.nid,'datarmy',2),datarmy.nTicket,'')
newSpPrim=IIF(SEEK(people.nid,'datarmy',2),datarmy.spprim,'')
newTelWork=people.telwork

SELECT rik FROM datarmy DISTINCT INTO CURSOR curSupRik READWRITE 
SELECT curSupRik
INDEX ON rik TAG T1
SELECT people
WITH oPage8     
     DO adtBoxAsCont WITH 'oPage8','cont1',10,10,RetTxtWidth('WНаименование организацииW'),dHeight,'Группа учёта',0,1   
     DO addComboMy WITH 'oPage8',1,.cont1.Left+.cont1.Width-1,.cont1.Top,dHeight,nParent.Width-.cont1.Width-20,.T.,'strGrup','curSupGrup.name',6,'DO procGotFocusSupGrup','DO procValidSupGrup',.F.,.T.     
     DO adtBoxAsCont WITH 'oPage8','cont4',10,.cont1.Top+.cont1.Height-1,.cont1.Width,dHeight,'Воинское звание',0,1 
     DO addComboMy WITH 'oPage8',3,.comboBox1.Left,.cont4.Top,dheight,.comboBox1.Width/2,.T.,'strZv','curSupZv.name',6,.F.,'DO procValidZv',.F.,.T. 
*     DO adTboxNew WITH 'oPage8','tBox113',.cont4.Top,.comboBox3.Left+.combobox3.Width-1,.comboBox1.Width-.comboBox3.Width+1,dHeight,'newPrimZv',.F.,.T.,0
     DO addComboMy WITH 'oPage8',113,.comboBox3.Left+.combobox3.Width-1,.cont4.Top,dheight,.comboBox1.Width-.comboBox3.Width+1,.T.,'newPrimZv','curPrimZv.primzv',6,'DO gotFocusPrimzv','DO validPrimZv',.F.,.T.  
     WITH .comboBox113       
          .Style=0      
     ENDWITH
      
     DO adtBoxAsCont WITH 'oPage8','cont11',10,.cont4.Top+.cont4.Height-1,.cont1.Width,dHeight,'Разряд запаса',0,1 
     DO adTboxNew WITH 'oPage8','tBox11',.cont11.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newRzp','Z',.T.,0,'9' 
     .tBox11.Alignment=0    
     DO adtBoxAsCont WITH 'oPage8','cont5',10,.cont11.Top+.cont11.Height-1,.cont1.Width,dHeight,'Военно-учётная спец-ть №',0,1 
     DO adTboxNew WITH 'oPage8','tBox5',.cont5.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newVus',.F.,.T.,0 
     
     DO adtBoxAsCont WITH 'oPage8','cont12',10,.cont5.Top+.cont5.Height-1,.cont1.Width,dHeight,'Профиль',0,1 
     DO adTboxNew WITH 'oPage8','tBox12',.cont12.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newProfil',.F.,.T.,0
     
     DO adtBoxAsCont WITH 'oPage8','cont2',10,.cont12.Top+.cont12.Height-1,.cont1.Width,dHeight,'Категория запаса',0,1 
     DO adTboxNew WITH 'oPage8','tBox2',.cont2.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newKatU','Z',.T.,0,'99' 
     .tBox2.Alignment=0
     DO adtBoxAsCont WITH 'oPage8','cont3',10,.cont2.Top+.cont2.Height-1,.cont1.Width,dHeight,'Дата приема на учет',0,1 
     DO adTboxNew WITH 'oPage8','tBox3',.cont3.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newdateuch',.F.,.T.,0
     
     DO adtBoxAsCont WITH 'oPage8','cont13',10,.cont3.Top+.cont3.Height-1,.cont1.Width,dHeight,'Дата снятия',0,1 
     DO adTboxNew WITH 'oPage8','tBox13',.cont13.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newDatesn',.F.,.T.,0
     
     DO adtBoxAsCont WITH 'oPage8','cont14',10,.cont13.Top+.cont13.Height-1,.cont1.Width,dHeight,'Основание снятия',0,1 
     DO adTboxNew WITH 'oPage8','tBox14',.cont14.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newprichsn',.F.,.T.,0
     
     DO adtBoxAsCont WITH 'oPage8','cont8',10,.cont14.Top+.cont14.Height-1,.cont1.Width,dHeight,'Cпецучёт №',0,1 
     DO adTboxNew WITH 'oPage8','tBox8',.cont8.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newUchet',.F.,.T.,0 
     
     DO adtBoxAsCont WITH 'oPage8','cont7',10,.cont8.Top+.cont8.Height-1,.cont1.Width,dHeight,'Комиссариат',0,1
     DO addComboMy WITH 'opage8',5,.comboBox1.Left,.cont7.Top,dheight,.comboBox1.Width,.T.,'strRik','curSupRik.rik',6,'DO gotFocusRik','DO procValidRik',.F.,.T.  
     WITH .comboBox5
          .Style=0
     ENDWITH             
    DO adtBoxAsCont WITH 'oPage8','cont10',10,.cont7.Top+.cont7.Height-1,.cont1.Width,dHeight,'№ военного билета',0,1 
    DO adTboxNew WITH 'oPage8','tBox10',.cont10.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newNTicket',.F.,.T.,0  
    
    DO adtBoxAsCont WITH 'oPage8','contSpPrim',10,.cont10.Top+.cont10.Height-1,.cont1.Width,dHeight,'особые отметки',0,1 
    DO adTboxNew WITH 'oPage8','tBoxSpprim',.contSpprim.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newSpprim',.F.,.T.,0                     
    
    DO adtBoxAsCont WITH 'oPage8','contTel',10,.contSpPrim.Top+.contSpprim.Height-1,.cont1.Width,dHeight,'Рабочий телефон',0,1 
     DO adTboxNew WITH 'oPage8','tBoxTel',.contTel.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newTelWork',.F.,.T.,0                     
     *
     *DO adtBoxAsCont WITH 'oPage8','cont8',10,.cont7.Top+.cont7.Height-1,.cont1.Width,dHeight,'Cпецучёт №',0,1 
     *DO adTboxNew WITH 'oPage8','tBox8',.cont8.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newUchet',.F.,.T.,0 
     *DO adtBoxAsCont WITH 'oPage8','cont9',10,.cont8.Top+.cont8.Height-1,.cont1.Width,dHeight,'Моб предписание\бронь',0,1 
     *DO addComboMy WITH 'oPage8',9,.comboBox1.Left,.cont9.Top,dHeight,.comboBox1.Width,.T.,'strMob','curSupMob.name',6,'DO procGotFocusSupMob','DO procValidSupMob',.F.,.T.
    
     *.Width=.cont1.Width+.comboBox1.Width+21
   *  DO adCheckBox WITH 'oPage8','checkSn','снят с учёта',.cont10.Top+.cont10.Height+10,.cont1.Left,150,dHeight,'newSnU',0         
   *  .checkSn.Left=(nParent.Width-.checkSn.Width)/2 
      *-----------------------------Кнопка сохранить---------------------------------------------------------------------------
     DO addcontlabel WITH 'oPage8','butSave',(nParent.Width-RetTxtWidth('wсохранитьw')*3-20)/2,.contTel.Top+.contTel.Height+20,RetTxtWidth('wсправочникиw'),dHeight+5,'сохранить','DO writeArmy'
     
     *---------------------------------Кнопка справочники ---------------------------------------------------------------------
     DO addcontlabel WITH 'oPage8','butSpr',.butSave.Left+.butSave.Width+10,.butSave.Top,.butSave.Width,dHeight+5,'справочники','DO procSprArmy','справочники'
     
     *---------------------------------Кнопка карточка ---------------------------------------------------------------------
     DO addcontlabel WITH 'oPage8','butCard',.butSpr.Left+.butSpr.Width+10,.butSave.Top,.butSave.Width,dHeight+5,'карточка','DO procCardArmy','личная карточка'
     .Refresh            
ENDWITH 

*******************************************************************************************************************************
PROCEDURE procGotFocusSupGrup
*******************************************************************************************************************************
PROCEDURE procValidSupGrup
newUch=curSupGrup.kod
strGrup=curSupGrup.name
*******************************************************************************************************************************
PROCEDURE procValidZv
newKzv=curSupZv.kod
newZv=curSupZv.name
KEYBOARD '{TAB}'
oPage8.Refresh

***********************************************************************************************************************
PROCEDURE gotFocusRik
=SYS(2002,1)

***********************************************************************************************************************
PROCEDURE procValidRik
IF EMPTY(oPage8.comboBox5.DisplayValue)=.F..AND.EMPTY(oPage8.comboBox5.Value)=.T.   
   SELECT curSupRik
   APPEND BLANK
   REPLACE rik WITH oPage8.comboBox5.DisplayValue 
   oPage8.comboBox5.Requery()  
   newRik=oPage8.comboBox5.DisplayValue  
   strRik=oPage8.comboBox5.DisplayValue   
ENDIF
newRik=oPage8.comboBox5.DisplayValue  
strRik=oPage8.comboBox5.DisplayValue 
oPage8.ComboBox5.ControlSource='strRik'
oPage8.Refresh
***********************************************************************************************************************
PROCEDURE gotFocusPrimZv
=SYS(2002,1)

***********************************************************************************************************************
PROCEDURE validPrimZv
IF EMPTY(oPage8.comboBox113.DisplayValue)=.F..AND.EMPTY(oPage8.comboBox113.Value)=.T.   
   SELECT curPrimZv
   APPEND BLANK
   REPLACE primzv WITH oPage8.comboBox113.DisplayValue 
   oPage8.comboBox113.Requery()  
   newPrimZv=oPage8.comboBox113.DisplayValue  
   *strRik=oPage8.comboBox5.DisplayValue   
ENDIF
newPrimZv=oPage8.comboBox113.DisplayValue  
*strRik=oPage8.comboBox5.DisplayValue 
oPage8.ComboBox113.ControlSource='newPrimZv'
oPage8.Refresh
*******************************************************************************************************************************
PROCEDURE readArmy
WITH oPage8
     .butRead.Visible=.F.
     .butExit.Visible=.F.
     .butSave.Visible=.T.
     .butReturn.Visible=.T.  
     .tBox4.Visible=.T.
     .tBox7.Visible=.T.
     .SetAll('Enabled',.T.,'myTxtBox')
     .SetAll('Enabled',.T.,'comboMy')
ENDWITH 
*******************************************************************************************************************************
PROCEDURE writeArmy
WITH oPage8     
     SELECT datArmy
     IF !SEEK(people.nid,'datarmy',2)
        APPEND BLANK
        REPLACE kodpeop WITH people.num,nidpeop WITH people.nid
     ENDIF
     *ON ERROR DO erSup
     REPLACE grupU WITH newUch,katU WITH newKatU,kzv WITH newKzv,zv WITH newzv,vus WITH newVus,riK WITH newRik,;
             uchet WITH newUchet,nTicket WITH newNTicket,rzp WITH newRzp,profil WITH newProfil,;
             dateUch WITH newDateUch,dateSn WITH newDateSn,prichSn WITH newPrichSn,primzv WITH newPrimZv,spPrim WITH newSpprim    
     REPLACE people.telwork WITH newTelWork      
     newUch=datarmy.grupu
     newKatU=datarmy.katu   
     newKzv=datarmy.kzv
     newzv=datarmy.zv
     newRzp=datarmy.rzp
     newProfil=datarmy.profil
     newVus=datarmy.vus
     newDateUch=datarmy.dateUch
     newDateSn=datarmy.dateSn
     newPrichSn=datarmy.prichsn
     newRik=datarmy.rik  
     newUchet=datarmy.uchet
     strGrup=IIF(SEEK(newUch,'curSupGrup',1),curSupGrup.name,'')
  
     strZv=datarmy.zv
     strRik=datarmy.rik
     newNTicket=datarmy.nTicket
     newPrimZv=datarmy.primZv  
     newSpprim=datarmy.spPrim  
     newTelWork=people.telWork
    .Refresh
    *ON ERROR 
ENDWITH 
*************************************************************************************************************************
PROCEDURE procSprArmy
kodSupSpr=10
datSup='curSupGrup'
newSprKod=0
newSprName=''
oldSprKod=0
oldSprname=''
frmSupl=CREATEOBJECT('FORMSUPL')
DIMENSION dimArmyProc(5)
STORE 0 TO dimArmyProc
dimArmyProc(1)=1
WITH frmSupl
     .Caption='Справочники для воинского учёта'
     .MinButton=.F.
     .MaxButton=.F.   
     DO addshape WITH 'frmSupl',1,10,20,150,390,8               
     .procexit='DO exitFromProcSprArmy'
     shapeWidth=390
     DO addOptionButton WITH 'frmSupl',1,'группа учёта',.Shape1.Top+20,.Shape1.Left+20,'dimArmyProc(1)',0,"DO procSelectOptionArmy WITH 1,10,'curSupGrup'",.T.     
   *  DO addOptionButton WITH 'frmSupl',2,'состав',.Option1.Top+.Option1.Height+10,.Option1.Left,'dimArmyProc(2)',0,"DO procSelectOptionArmy WITH 2,11,'curSupSostav'",.T. 
     DO addOptionButton WITH 'frmSupl',3,'звание',.Option1.Top+.Option1.Height+10,.Option1.Left,'dimArmyProc(3)',0,"DO procSelectOptionArmy WITH 3,14,'curSupZv'",.T. 
  *   DO addOptionButton WITH 'frmSupl',4,'категория годности',.Option3.Top+.Option3.Height+10,.Option1.Left,'dimArmyProc(4)',0,"DO procSelectOptionArmy WITH 4,12,'curSupGoden'",.T. 
  *   DO addOptionButton WITH 'frmSupl',5,'моб.предписание',.Option4.Top+.Option4.Height+10,.Option1.Left,'dimArmyProc(5)',0,"DO procSelectOptionArmy WITH 5,13,'curSupMob'",.T. 
     .Shape1.Height=.Option1.Height*2+50
     .Shape1.Width=RetTxtWidth('WWWмобилизационное предписаниеWWW')
     *-----------------------------Кнопка приступить---------------------------------------------------------------------------
     DO addcontlabel WITH 'frmSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприступитьw')*2)-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприступитьw'),dHeight+5,'приступить','DO procRunSprArmy'
     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'frmSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'отмена','DO exitFromProcSprArmy','отмена'
     .SetAll('ForeColor',RGB(0,0,128),'CheckBox')  
     *-------------параметры формы---------------------------------------------------------------------------------------------      
    .Width=.Shape1.Width+20
    .Height=.Shape1.Height+.cont1.Height+60         
ENDWITH
DO pasteImage WITH 'frmSupl'
frmSupl.Show
**************************************************************************************************************************
PROCEDURE procSelectOptionArmy
PARAMETERS par1,par2,par3
STORE 0 TO dimArmyProc
dimArmyProc(par1)=1
kodSupSpr=par2
datSup=par3
frmSupl.Refresh
**************************************************************************************************************************
PROCEDURE procRunSprArmy
frmSupl.Release
newSprKod=0
newSprName=''
oldSprKod=0
oldSprname=''
logNewSpr=.F.
sprRec=0
frmSpr=CREATEOBJECT('FORMSUPL')
WITH frmSpr      
     .Caption=''
     .procExit='DO exitProcSelectOptionArmy'
     .Width=600
     .Height=400
     *DO addmenureadspr WITH 'frmSpr','DO writeSprArmy WITH .T.','DO writeSprArmy WITH .F.'
     DO addcontmenu WITH 'frmSpr','menucont1',10,5,'новая','pencila.ico',"Do readspr WITH 'frmSpr','Do readSprArmy WITH .T.'"
     DO addcontmenu WITH 'frmSpr','menucont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'frmSpr','Do readSprArmy WITH .F.'"
     DO addcontmenu WITH 'frmSpr','menucont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO formDelSprArmy'         
     DO addcontmenu WITH 'frmSpr','menucont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO exitProcSelectOptionArmy'
     
      DO addmenureadspr WITH 'frmSpr','DO writeSprArmy WITH .T.','DO writeSprArmy WITH .F.'  
     .AddObject('fGrid','gridMyNew')
     SELECT &datSup
     WITH .fGrid
          .ColumnCount=0 
           DO addColumnToGrid WITH 'frmSpr.fGrid',3
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Parent.menucont1.Height-5    
          .Width=.Parent.Width             
          .RecordSourceType=1                 
          .RecordSource='&datSup'              
          .backColor=RGB(255,255,255)
          .Column1.ControlSource='kod'
          .Column2.ControlSource='name'   
          .Column1.Width=RettxtWidth(' 1234 ')         
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование'        
          .Columns(.ColumnCount).Width=0
          .ScrollBars=2
          .Column2.Width=.Width-.Column1.width-SYSMETRIC(5)-13-.ColumnCount           
          .Column1.Alignment=1
          .Column2.Alignment=0        
          .colNesInf=2     
          .Visible=.T.                 
     ENDWITH
     DO gridSizeNew WITH 'frmSpr','fGrid','shapeingrid'  
     FOR i=1 TO .fGrid.columnCount  
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(frmSpr.fGrid.RecordSource)#frmSpr.fGrid.curRec,frmSpr.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(frmSpr.fGrid.RecordSource)#frmSpr.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR  
     DO addtxtboxmy WITH 'frmSpr',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
     .txtbox1.Enabled=.F.
     DO addtxtboxmy WITH 'frmSpr',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0  
     .SetAll('Visible',.F.,'MyTxtBox')  
     DO addcontmy WITH 'frmSpr','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'frmSpr','frmSpr.cont1','&datSup',1"
     .cont1.SpecialEffect=1   
     DO addcontmy WITH 'frmSpr','cont2',.cont1.Left+.fGrid.Column1.Width+2,.fGrid.Top+2,.fGrid.Column2.Width-4,.fGrid.HeaderHeight-3,''
     .Height=.menuCont1.Height+.fGrid.Height+10 
ENDWITH 
SELECT &datSup
GO TOP
DO pasteImage WITH 'frmSpr'
frmSpr.Show
**************************************************************************************************************************
PROCEDURE readSprArmy
PARAMETERS parLog
WITH frmSpr
     SELECT &datSup   
     logNewSpr=IIF(parLog,.T.,.F.) 
     oldSprName=IIF(parLog,'',name)
     oldSprKod=IIF(parLog,0,kod)
     IF parLog
        oldOrd=SYS(21)
        SET ORDER TO 1
        GO BOTTOM
        newSprKod=kod+1
        APPEND BLANK
        REPLACE kod WITH newSprKod        
        SET ORDER TO &oldOrd
     ENDIF     
     sprRec=RECNO()
     newSprKod=kod
     newSprName=name
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1    
     .txtbox1.ControlSource='newSprKod'
     .txtbox2.ControlSource='newSprName' 
     .Refresh    
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')   
     .SetAll('Visible',.T.,'MyTxtBox')       
     .fGrid.Enabled=.F.
     .txtbox2.SetFocus
ENDWITH  
****************************************************************************************************************************
PROCEDURE writeSprArmy   
PARAMETERS par_log
WITH frmSpr
     .SetAll('Visible',.T.,'mymenucont')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     IF par_log
        SELECT &datSup
        REPLACE kod WITH newSprKod,name WITH newSprName     
        SELECT sprtot               
        IF logNewSpr
           APPEND BLANK
           REPLACE kSpr WITH kodSupSpr
        ELSE
           LOCATE FOR kSpr=kodSupSpr.AND.kod=oldSprKod           
        ENDIF   
        REPLACE name WITH newSprName
        SELECT &datSup
     ELSE
        IF logNewSpr
           DELETE
        ENDIF   
     ENDIF    
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'ComboMy')
     .SetAll('Visible',.F.,'MySpinner')
     .fGrid.Enabled=.T.
     SELECT &datSup     
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     GO sprRec
     .Refresh
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
ENDWITH    
**************************************************************************************************************************
PROCEDURE exitProcSelectOptionArmy
frmSpr.Release
**************************************************************************************************************************
PROCEDURE exitFromProcSprArmy
frmSupl.Release
**************************************************************************************************************************
PROCEDURE procCardArmy
DIMENSION dim_opt(2)
dim_opt(1)=1
dim_opt(2)=0
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Печать личной карточки'
     .Icon='kone.ico'
     .procExit='DO exitCardArmy'
     DO addshape WITH 'fSupl',1,10,10,150,400,8
     DO addOptionButton WITH 'fSupl',1,'1-я страница',.Shape1.Top+20,20,'dim_opt(1)',0,"DO procValOption WITH 'fSupl','dim_opt',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'2-я страница',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_opt(2)',0,"DO procValOption WITH 'fSupl','dim_opt',2",.T. 
     .Option1.Left=.Shape1.Left+(.Shape1.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20 
     .Shape1.Height=.Option1.Height+40
     
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.T.,.F.
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WпросмотрW')*3-30)/2,.Shape91.Top+.Shape91.Height+20,RetTxtWidth('WпросмотрW'),dHeight+3,'печать','DO cardArmyPrn WITH 1','печать личной карточик'
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+3,'просмотр','DO cardArmyPrn WITH 2','печать личной карточик'
     *---------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+15,.Cont1.Top,.Cont1.Width,dHeight+3,'возврат','DO exitCardArmy','возврат'
     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape91.Height+.Cont1.Height+60      
     DO pasteImage WITH 'fSupl'
     .Show 
ENDWITH
***************************************************************************************************************************
PROCEDURE cardArmyPrn
PARAMETERS parterm
IF !USED('boss')
   USE boss IN 0
ENDIF 
DIMENSION dimArmyFam(5)
STORE '' TO dimArmyfam,fioWife,armWork,armSpec,ctelok
SELECT boss
ctelok=ALLTRIM(telok)
USE
SELECT curjobsupl
armrec=RECNO()
LOCATE FOR tr=1
kparm=curJobsupl.kp
kdarm=curJobsupl.kd
armWork=IIF(SEEK(kparm,'sprpodr',1),sprpodr.name,'')
armDol=IIF(SEEK(kdarm,'sprdolj',1),sprdolj.name,'')
ON ERROR DO erSup
GO armRec
ON ERROR 
SELECT datArmy
SEEK people.num
fioArmy=ALLTRIM(LEFT(people.fio,AT(' ',people.fio)))
strFio=ALLTRIM(SUBSTR(ALLTRIM(people.fio),AT(' ',people.fio)))            
nameArmy=LEFT(strFio,AT(' ',strFio))
otchArmy=ALLTRIM(SUBSTR(ALLTRIM(people.fio),RAT(' ',ALLTRIM(people.fio))))
SELECT datFam
SET ORDER TO 3
LOCATE FOR kodpeop=people.num.AND.kfam>2
i=1
SCAN WHILE kodpeop=people.num.AND.kfam<5
     dimArmyFam(i)=IIF(SEEK(kfam,'curSprFam',1),ALLTRIM(curSprFam.name)+' ','')+ALLTRIM(datfam.nfio)+' '+IIF(!EMPTY(dBirth),DTOC(dBirth),'')
     i=i+1
ENDSCAN
fioWife=IIF(SEEK(STR(people.num,4)+STR(1,2),'datFam',3).OR.SEEK(STR(people.num,4)+STR(2,2),'datFam',3),ALLTRIM(datfam.nfio),'')+' '+IIF(!EMPTY(dBirth),DTOC(dBirth),'')
SET ORDER TO 1
SELECT datArmy
CREATE CURSOR curCardArm (namea C(100))
SELECT curCardArm
APPEND BLANK
IF parTerm=1
   DO CASE
      CASE dim_opt(1)=1    
           DO procForPrintAndPreview WITH 'repCardArmy1','личная карточка',.T.,'cardArmyToWord'
      CASE dim_opt(2)=1     
           DO procForPrintAndPreview WITH 'repCardArmy2','личная карточка',.T.,'cardArmyToWord'
   ENDCASE
ELSE 
   DO CASE
      CASE dim_opt(1)=1
           DO procForPrintAndPreview WITH 'repCardArmy1','личная карточка',.F.
      CASE dim_opt(2)=1     
           DO procForPrintAndPreview WITH 'repCardArmy2','личная карточка',.F.
   ENDCASE
ENDIF 
***************************************************************************************************************************
PROCEDURE cardArmyToWord
#DEFINE wdBorderTop -1           &&верхняя граница ячейки таблицы
#DEFINE wdBorderLeft -2          &&левая граница ячейки таблицы
#DEFINE wdBorderBottom -3        &&нижняя граница ячейки таблицы
#DEFINE wdBorderRight -4         &&правая граница ячейки таблицы
#DEFINE wdBorderHorizontal -5    &&горизонтальные линии таблицы
#DEFINE wdBorderVertical -6      &&горизонтальные линии таблицы
#DEFINE wdLineStyleSingle 1      && стиль линии границы ячейки (в данно случае обычная)
#DEFINE wdLineStyleNone 0        && линия отсутствует
#DEFINE wdAlignParagraphRight 2
ON ERROR DO erSup
objWord=CREATEOBJECT('WORD.APPLICATION')
#DEFINE cr CHR(13)
nameDoc=objWord.Documents.Add()  
nameDoc.ActiveWindow.View.ShowAll=0        
objWord.Selection.pageSetup.Orientation=0
objWord.Selection.pageSetup.LeftMargin=30
objWord.Selection.pageSetup.RightMargin=20
objWord.Selection.pageSetup.TopMargin=10
objWord.Selection.pageSetup.BottomMargin=10
docRef=GETOBJECT('','word.basic')
WITH docRef
     .Insert(cr)
     .Insert(cr)
     .Insert(cr)
     .Font('Times New Roman',12)
     .LeftPara   
     nameDoc.Tables.add(objWord.Selection.range,1,5)     
     
     ordTable1=nameDoc.Tables(1) 
     WITH ordTable1
          .Columns(1).Width=80
          .Columns(2).Width=50
          .Columns(3).Width=260
          .Columns(4).Width=50
          .Columns(5).Width=100
          
          .cell(1,1).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,1).Range.Text=LEFT(ALLTRIM(people.fio),1)
          docRef.CloseParaBelow 
          *.Columns(1).Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
          *.Columns(1).Borders(wdBorderVertical).LineStyle=wdLineStyleSingle 
          .cell(1,1).Borders(wdBorderRight).LineStyle=wdLineStyleSingle 
          .cell(1,1).Borders(wdBorderLeft).LineStyle=wdLineStyleSingle 
         
          *.Columns(1).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle  
          
          .cell(1,5).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,5).Range.Text=IIF(SEEK(datarmy.grupu,'curSupGrup',1),curSupGrup.nameSp,'')
          docRef.CloseParaBelow 
          .cell(1,5).Borders(wdBorderRight).LineStyle=wdLineStyleSingle 
          .cell(1,5).Borders(wdBorderLeft).LineStyle=wdLineStyleSingle   
                         
          .cell(1,2).Range.Select     
          docRef.CloseParaBelow
          .cell(1,3).Range.Select     
          docRef.CloseParaBelow
          .cell(1,4).Range.Select     
          docRef.CloseParaBelow
          .cell(1,5).Range.Select     
          docRef.CloseParaBelow          
          .Rows.Add
          .cell(2,5).Range.Select     
        *  .cell(2,5).Range.Text=ALLTRIM(datarmy.rik)
                   
          docRef.CloseParaBelow          
          .cell(2,3).Range.Select   
          docRef.CenterPara
          docRef.Font('Times New Roman',11)
          docRef.Bold
          .cell(2,3).Range.Text='ЛИЧНАЯ КАРТОЧКА'
          .Rows.Add
          .cell(1,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          .cell(3,1).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle  
          .cell(1,5).Borders(wdBorderTop).LineStyle=wdLineStyleSingle  
          .cell(3,5).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle  
          docRef.CloseParaBelow   
          .cell(3,3).Range.Select
          .Rows.Add
          .cell(4,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          .cell(4,1).Borders(wdBorderLeft).LineStyle=wdLineStyleNone 
          .cell(4,1).Borders(wdBorderRight).LineStyle=wdLineStyleNone 
          .cell(4,1).Borders(wdBorderBottom).LineStyle=wdLineStyleNone
          .cell(4,5).Borders(wdBorderTop).LineStyle=wdLineStyleSingle   
          .cell(4,5).Borders(wdBorderLeft).LineStyle=wdLineStyleNone 
          .cell(4,5).Borders(wdBorderRight).LineStyle=wdLineStyleNone
          .cell(4,5).Borders(wdBorderBottom).LineStyle=wdLineStyleNone  
          .cell(4,1).Range.Text='(Первая буква фамилии)'
          .cell(4,5).Range.Select
          docRef.Font('Times New Roman',11)
          .cell(4,5).Range.Text='(Группа учета (О,ОГБ,ПСС,П))' 
          .cell(4,5).Range.Select         
          docRef.CloseParaBelow             
          docRef.LineDown    
     ENDWITH 
           
     .Insert(cr)
     .Font('Times New Roman',12)
     .Bold
     .CenterPara   
     .Insert('ПЕРСОНАЛЬНЫЕ ДАННЫЕ')
     
     .Insert(cr)
     .Font('Times New Roman',12)
     .LeftPara     
     nameDoc.Tables.add(objWord.Selection.range,1,2)     
    
     ordTable2=nameDoc.Tables(2) 
     WITH ordTable2
          .Columns(1).Width=300
          .Columns(2).Width=240
                  
          .cell(1,1).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,1).Range.Text='Фамилия'
          .cell(1,1).Borders(wdBorderRight).LineStyle=wdLineStyleSingle                             
          .cell(1,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          docRef.CloseParaBelow 
          
          .cell(1,2).Range.Select   
          docRef.Font('Times New Roman',11)          
          .cell(1,2).Range.Text=fioArmy
          .cell(1,2).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          docRef.CloseParaBelow 
          .Rows.Add
          .cell(2,1).Range.Text='Собственное имя' 
          .cell(2,1).Range.Select 
          .cell(2,2).Range.Text=nameArmy
          docRef.CloseParaBelow                             
          .Rows.Add          
          .cell(3,1).Range.Text='Отчество (если таковое имеется)'                                     
          .cell(3,2).Range.Text=otchArmy
          .Rows.Add          
          .cell(4,1).Range.Text='Дата рождения'  
          .cell(4,2).Range.Text=IIF(!EMPTY(people.age),DTOC(people.age),'')                         
          .Rows.Add          
          .cell(5,1).Range.Text='Место рождения' 
          .cell(5,2).Range.Text=ALLTRIM(people.placeborn)
          .Rows.Add          
          .cell(6,1).Range.Text='Идентификационный номер'   
          .cell(6,2).Range.Text=ALLTRIM(people.pnum)
          .Rows.Add          
          .cell(7,1).Range.Text='Место жительства'             
          .cell(7,2).Range.Text=ALLTRIM(people.preg)
          .Rows.Add
          .cell(8,1).Range.Text='Место пребывания'             
          .cell(8,2).Range.Text=ALLTRIM(people.ppreb)      
          .Rows.Add
          .cell(9,1).Select
          .cell(9,1).Split(1,2)          
          .cell(9,1).Range.Text='Образование'             
          .cell(9,2).Range.Text='уровень основного образования'
          .cell(9,3).Range.Text=IIF(SEEK(people.educ,'cureducation',1),cureducation.name,'')
          .Rows.Add
          .cell(10,2).Range.Text='учреждения образования и годы их окончания'
          .cell(10,3).Select
          .cell(10,3).Split(1,2)          
          .cell(10,3).Range.Text='специальности (профессии)'                                    
          .cell(10,4).Range.Text='присвоенные квалификации'                                    
          .Rows.Add
          .cell(11,2).Range.Text=ALLTRIM(people.school)+' '+IIF(!EMPTY(people.endeduc),DTOC(people.endeduc),'')
          .cell(11,3).Range.Text=ALLTRIM(people.specd)
          .cell(11,4).Range.Text=ALLTRIM(people.kvald)
          .Cell(9,1).Merge(.Cell(11,1))
          .Rows.Add
          .cell(12,1).Range.Text='Семья'   
          .cell(12,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle           
          .cell(12,2).Range.Text='семейное положение'
          .Cell(12,3).Merge(.Cell(12,4))
          .cell(12,3).Range.Text=IIF(SEEK(people.family,'curfamily',1),curfamily.name,'')
          .Rows.Add
        *  SELECT datfam
         * SET ORDER TO 3
          .cell(13,2).Range.Text='супруга (супруг)'
          .cell(13,3).Range.Text=fioWife
          .Rows.Add
          .cell(14,2).Range.Text='дети/ родители, если гражданин холост (не замужем ) и не имеет детей'
         
          .cell(14,3).Range.Text=dimArmyFam(1)
          .Rows.Add
          .cell(15,3).Range.Text=dimArmyFam(2)
          .Rows.Add
          .cell(16,3).Range.Text=dimArmyFam(3)
          .Rows.Add
          .cell(17,3).Range.Text=dimArmyFam(4)  
          .Rows.Add
          .cell(18,3).Range.Text=dimArmyFam(5)
          .Rows.Add
          .cell(19,2).Range.Text='место жительства близких родственников, которые не проживают совместно с гражданином'
          .Cell(14,2).Merge(.Cell(18,2))
          .Cell(12,1).Merge(.Cell(19,1))
       *   SET ORDER TO 1
          SELECT people
          .Rows.Add
          
          .cell(20,1).Range.Text='Работа (учёба)'  
          .cell(20,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle            
          .cell(20,2).Range.Text='структурное подразделение'
          .cell(20,3).Range.Text=armWork  
          .cell(20,2).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          .Rows.Add
          .cell(21,2).Range.Text='должность (профессия)'
          .cell(21,3).Range.Text=armDol
          .Cell(20,1).Merge(.Cell(21,1))
          .Rows.Add
          .cell(22,1).Range.Text='Номера телефонов'
          .cell(22,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          .cell(22,2).Range.Text='рабочий'
          .cell(22,3).Range.Text=IIF(!EMPTY(people.telwork),people.telwork,ctelok)
          .Rows.Add
          .cell(23,2).Range.Text='домашний'
          .cell(23,3).Range.Text=people.telHome
          docRef.CloseParaBelow  
          .Rows.Add
          .cell(24,2).Range.Text='мобильный'
          .cell(24,3).Range.Text=people.telMob
          .cell(24,1).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(24,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(24,3).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .Cell(22,1).Merge(.Cell(24,1))
          .cell(24,3).Range.Select 
          docRef.CloseParaBelow  
          docRef.LineDown    
     ENDWITH 
     .Insert(cr)
     .Font('Times New Roman',12)
     .Bold
     .CenterPara   
     .Insert('ПЕРВИЧНЫЕ ДАННЫЕ ВОИНСКОГО УЧЁТА') 
     docRef.CloseParaBelow  
     .Insert(cr)
     .LeftPara
     nameDoc.Tables.add(objWord.Selection.range,1,2)     
     
     ordTable3=nameDoc.Tables(3) 
     WITH ordTable3
          .Columns(1).Width=300
          .Columns(2).Width=240
                  
          .cell(1,1).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,1).Range.Text='Группа учёта '
          .cell(1,1).Borders(wdBorderRight).LineStyle=wdLineStyleSingle                             
          .cell(1,1).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          docRef.CloseParaBelow 
          
          .cell(1,2).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,2).Range.Text=IIF(SEEK(datarmy.grupu,'curSupGrup',1),ALLTRIM(curSupGrup.name),'')+'  '+ALLTRIM(datarmy.rik)
          .cell(1,2).Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          docRef.CloseParaBelow 
          .Rows.Add
          .cell(2,1).Range.Text='Воинское звание' 
          .cell(2,1).Range.Select   
           docRef.CloseParaBelow     
          .cell(2,2).Range.Text=IIF(SEEK(datarmy.kzv,'curSupZv',1),ALLTRIM(curSupZv.name)+' '+ALLTRIM(newPrimZv),'') 
          .cell(2,2).Range.Select   
           docRef.CloseParaBelow                          
          .Rows.Add          
          .cell(3,1).Range.Text='Разряд запаса'
          .cell(3,2).Range.Text=newRzp
          .Rows.Add          
          .cell(4,1).Range.Text='Номер военно-учетной специальности'  
          .cell(4,2).Range.Text=newVus
          .Rows.Add          
          .cell(5,1).Range.Text='Профиль' 
          .cell(5,2).Range.Text=newProfil
          .Rows.Add          
          .cell(6,1).Range.Text='Категория запаса' 
          .cell(6,2).Range.Text=newKatu
          .Rows.Add          
          .cell(7,1).Range.Text='Дата приема гражданина на воинский учет'             
          .cell(7,2).Range.Text=IIF(!EMPTY(NewDateUch),DTOC(newDateUch),'')
          .Rows.Add
          .cell(8,1).Range.Text='Дата и основание снятия (исключения) гражданина с воинского учета'             
          .cell(8,2).Range.Text=IIF(!EMPTY(newDateSn),DTOC(newDateSn)+' ','')+ALLTRIM(newPrichsn)
          .Rows.Add
           .cell(9,1).Range.Text='Состоит на специальном учете'             
          .cell(9,2).Range.Text=ALLTRIM(newuchet)
          .cell(9,1).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(9,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(9,2).Select
          docRef.CloseParaBelow  
          docRef.LineDown   
          .cell(9,2).Select
          docRef.CloseParaBelow  
          docRef.LineDown     
     ENDWITH 
     
     .Insert(cr)
     .Insert(cr)
     .Insert(cr)
     .Insert(cr)
     .Insert(cr)  
     .Font('Times New Roman',12)
     .Bold
     .CenterPara   
     .Insert('ОСОБЫЕ ОТМЕТКИ') 
     docRef.CloseParaBelow  
     .Insert(cr)
     .LeftPara
     nameDoc.Tables.add(objWord.Selection.range,1,1)     
     ordTable4=nameDoc.Tables(4) 
     WITH ordTable4               
          .Columns(1).Width=540  
          .cell(1,1).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,1).Range.Text=ALLTRIM(newSpprim)
          docRef.CloseParaBelow 
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add   
          .Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          .Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
          .Borders(wdBorderBottom).LineStyle=wdLineStyleSingle   
          .cell(9,1).Select
          docRef.CloseParaBelow  
          docRef.LineDown   
          docRef.CloseParaBelow 
     ENDWITH
     
     .Insert(cr)
     .Font('Times New Roman',12)
     .Bold
     .CenterPara   
     .Insert('ОТМЕТКИ О СВЕРКЕ ДАННЫХ') 
     docRef.CloseParaBelow  
     .Insert(cr)
     .CenterPara
     nameDoc.Tables.add(objWord.Selection.range,1,3)     
     ordTable5=nameDoc.Tables(5) 
     WITH ordTable5
          .Columns(1).Width=100                         
          .Columns(2).Width=340                         
          .Columns(3).Width=100   
          .cell(1,1).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,1).Range.Text='Дата сверки'        
          .cell(1,1).Borders(wdBorderRight).LineStyle=wdLineStyleSingle        
          docRef.CloseParaBelow           
          .cell(1,2).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,2).Range.Text='С кем или чем проводилась сверка (наименование органа, уполномоченного проводить сверки или документа )'     
          .cell(1,2).Borders(wdBorderRight).LineStyle=wdLineStyleSingle      
           docRef.CloseParaBelow           
          .cell(1,3).Range.Select   
          docRef.Font('Times New Roman',11)
          .cell(1,3).Range.Text='Подпись, инициалы, фамилия лица проводившего сверку'       
          docRef.CloseParaBelow 
          
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add 
          .Rows.Add 
          .Rows.Add 
          .Rows.Add 
          .Rows.Add 
          .Rows.Add 
          .Rows.Add    
          .Borders(wdBorderTop).LineStyle=wdLineStyleSingle 
          .Borders(wdBorderHorizontal).LineStyle=wdLineStyleSingle 
          .Borders(wdBorderBottom).LineStyle=wdLineStyleSingle    
          .cell(15,3).Select
          docRef.CloseParaBelow  
          docRef.LineDown  
     ENDWITH
     .Insert(cr)
     .Insert(cr)
     .LeftPara
     docRef.Font('Times New Roman',11)
     .Insert(REPLICATE('_',45))
     .Insert(cr)
     docRef.Font('Times New Roman',11)
     .Insert('(подпись, инициалы, фамилия должностного лица')
     docRef.CloseParaBelow  
     .Insert(cr)
     docRef.Font('Times New Roman',11)
     .Insert('ответственного за ведение военно-учётной работы)')     
ENDWITH  
ON ERROR 
objWord.Visible=.T.       
***************************************************************************************************************************
PROCEDURE exitCardArmy
fSupl.Visible=.F.
*fSupl.Release