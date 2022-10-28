***********************************************************************************************************************
PARAMETERS parUv
PUBLIC dimKnt(3)
STORE 0 TO dimKnt
dimKnt(1)=1
IF parUv
   SELECT * FROM datJobout WHERE datjobout.nidpeop=peopout.nid INTO CURSOR curJobCard ORDER BY tr READWRITE 
ELSE  
   SELECT * FROM datJob WHERE datjob.kodpeop=people.num.AND.EMPTY(dateOut) INTO CURSOR curJobCard ORDER BY tr READWRITE 
ENDIF  
PUBLIC newTotDay,newdayOtp,newDayKont,newDayVred,newDayNorm,newKTime,newVidDog,newNumDog,newdDog,newSrok,newbegDog,newEndDog,newPkont
newVidDog=IIF(!parUv,people.dog,peopout.dog)      && тип договора (контракт,трудовой договор,срочный трудовой договор)*
strVid=IIF(SEEK(IIF(!parUv,people.dog,peopout.dog),'sprdog',1),sprdog.name,'')
strSrok=IIF(SEEK(IIF(!parUv,people.kTime,peopOut.kTime),'cursrok',1),cursrok.name,'')
IF !parUv
   SELECT people
ELSE
   SELECT peopout
ENDIF
newNumDog=numDog
newDDog=ddog
newkTime=kTime     && код времени
newBegDog=begDog   && начало контракта
newEndDog=endDog   && окончание контракта
newPkont=pKont     && процент по контракту


newTotDay=totDay   && всего дней отпуска
newDayOtp=dayOtp   && основной отпуск
newDayKont=dayKont && поощрительный отпуск
newDayVred=dayVred && за вредность
newDayNorm=dayNorm && за ненормированный

PUBLIC newKval,strKval,newDKval,newNordKval,newDordKval,newNkval
newKval=kval
strKval=IIF(SEEK(newKval,'sprkval',1),sprkval.name,'')
newDKval=dkval
newNordKval=nordKval
newDordKval=dordKval
newNkval=nKval

WITH oPage5
     DO addShape WITH 'oPage5',1,10,10,20,nParent.Width-20,8 
     DO adTboxAsCont WITH 'opage5','txtVid',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('Wтрудовой договор'),dHeight,'вид договора',2,1
     DO  addComboMy WITH 'opage5',1,.txtVid.Left,.txtVid.Top+.txtVid.Height-1,dheight,.txtVid.Width,.T.,'strVid','ALLTRIM(sprDog.name)',6,.F.,'newVidDog=sprdog.kod',.F.,.T.                    
     
     DO adTboxAsCont WITH 'opage5','txtNum',.txtVid.Left+.txtVid.Width-1,.txtVid.Top,RetTxtWidth('wномерw'),dHeight,'номер',2,1
     DO adTboxNew WITH 'opage5','boxNum',.comboBox1.Top,.txtNum.Left,.txtNum.Width,dHeight,'newNumDog',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'opage5','txtDdog',.txtNum.Left+.txtNum.Width-1,.txtVid.Top,RetTxtWidth('99/99/999999'),dHeight,'дата',2,1
     DO adTboxNew WITH 'opage5','boxDDog',.comboBox1.Top,.txtDDog.Left,.txtDDog.Width,dHeight,'newDDog',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'opage5','txtSrok',.txtDDog.Left+.txtDDog.Width-1,.txtVid.Top,RetTxtWidth('не определ. срокw'),dHeight,'срок',2,1
     DO  addComboMy WITH 'opage5',2,.txtSrok.Left,.comboBox1.Top,dheight,.txtSrok.Width,.T.,'strSrok','ALLTRIM(curSrok.name)',6,.F.,'newKTime=cursrok.kod',.F.,.T.    
     
     DO adTboxAsCont WITH 'opage5','txtBegDog',.txtSrok.Left+.txtSrok.Width-1,.txtVid.Top,RetTxtWidth('99/99/999999'),dHeight,'начало',2,1
     DO adTboxNew WITH 'opage5','boxBegDog',.comboBox1.Top,.txtBegDog.Left,.txtBegDog.Width,dHeight,'newBegDog',.F.,.T.,0,.F.
      
     DO adTboxAsCont WITH 'opage5','txtEndDog',.txtBegDog.Left+.txtBegDog.Width-1,.txtVid.Top,.txtBegDog.Width,dHeight,'окончание',2,1
     DO adTboxNew WITH 'opage5','boxEndDog',.comboBox1.Top,.txtEndDog.Left,.txtEndDog.Width,dHeight,'newEndDog',.F.,.T.,0,.F.      
     
     DO adTboxAsCont WITH 'opage5','txtPkont',.txtEndDog.Left+.txtEndDog.Width-1,.txtVid.Top,(.Shape1.Width-.txtEndDog.Left-.txtEndDog.Width)/6,dHeight,'% по контр.',2,1
     DO adTboxNew WITH 'opage5','boxPkont',.comboBox1.Top,.txtPKont.Left,.txtPKont.Width,dHeight,'newPkont','Z',.T.,0,.F.
     .boxPkont.InputMask='999'  
     
     DO adTboxAsCont WITH 'opage5','txtDayOtp',.txtPkont.Left+.txtPKont.Width-1,.txtVid.Top,.txtPKont.Width,dHeight,'отп. дни',2,1     
     DO adTboxNew WITH 'opage5','boxDayOtp',.comboBox1.Top,.txtDayOtp.Left,.txtDayOtp.Width,dHeight,'newDayOtp','Z',.T.,0,.F.,"DO validtotday WITH 'opage5','newTotDay=newDayOtp+newDayKont+newDayNorm+newDayVred'"
     .boxDayOtp.InputMask='999'
     
     DO adTboxAsCont WITH 'opage5','txtDayKont',.txtDayOtp.Left+.txtDayOtp.Width-1,.txtVid.Top,.txtPKont.Width,dHeight,'поощр отпуск',2,1
     DO adTboxNew WITH 'opage5','boxDayKont',.comboBox1.Top,.txtDayKont.Left,.txtDayKont.Width,dHeight,'newDayKont','Z',.T.,0,.F.,"DO validtotday WITH 'opage5','newTotDay=newDayOtp+newDayKont+newDayNorm+newDayVred'"     
     .boxDayKont.InputMask='999'
      
     DO adTboxAsCont WITH 'opage5','txtDayNorm',.txtDayKont.Left+.txtDayKont.Width-1,.txtVid.Top,.txtPKont.Width,dHeight,'за ненор.',2,1
     DO adTboxNew WITH 'opage5','boxDayNorm',.comboBox1.Top,.txtDayNorm.Left,.txtDayNorm.Width,dHeight,'newDayNorm','Z',.T.,0,.F.,"DO validtotday WITH 'opage5','newTotDay=newDayOtp+newDayKont+newDayNorm+newDayVred'"
     .boxDayNorm.InputMask='999'
     
     DO adTboxAsCont WITH 'opage5','txtDayVred',.txtDayNorm.Left+.txtDayNorm.Width-1,.txtVid.Top,.txtPKont.Width,dHeight,'за вред.',2,1
     DO adTboxNew WITH 'opage5','boxDayVred',.comboBox1.Top,.txtDayVred.Left,.txtDayVred.Width,dHeight,'newDayVred','Z',.T.,0,.F.,"DO validtotday WITH 'opage5','newTotDay=newDayOtp+newDayKont+newDayNorm+newDayVred'"
     .boxDayNorm.InputMask='999'
        
     DO adTboxAsCont WITH 'opage5','txtDayTot',.txtDayVred.Left+.txtDayVred.Width-1,.txtVid.Top,.txtPKont.Width,dHeight,'всего дней',1,1
     DO adTboxNew WITH 'opage5','boxDayTot',.comboBox1.Top,.txtDayTot.Left,.txtDayTot.Width,dHeight,'newTotDay','Z',.T.,0,.F.
     .boxDayTot.InputMask='999'  
   
     .Shape1.Height=.txtVid.Height*2+20         
     *--------------------------------Кнопка сохранить-------------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','cont1',.Shape1.Left+(.Shape1.Width-(RetTxtWidth('wсохранитьw')*3)-20)/2,.Shape1.Top+.Shape1.Height+15,RetTxtWidth('WсохранитьW'),dHeight+5,'сохранить','DO saveKontrakt'
     *---------------------------------Кнопка удалить --------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','cont2',.cont1.Left+.cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+5,'удалить','DO delKont'                    
     *---------------------------------Кнопка печать --------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','cont3',.cont2.Left+.cont2.Width+10,.Cont1.Top,.Cont1.Width,dHeight+5,'печать','DO formKontraktPrn'       
     
     
     *--------------------------------Кнопка удалить (при удалени)-------------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','butDel',.Shape1.Left+(.Shape1.Width-(RetTxtWidth('wудалитьw')*2)-10)/2,.cont1.Top,RetTxtWidth('WудалитьW'),dHeight+5,'удалить','DO delinfokontrakt WITH .T.'
     *---------------------------------Кнопка удалить --------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','butRet',.ButDel.Left+.butDel.Width+10,.butDel.Top,.Cont1.Width,dHeight+5,'возврат','DO delinfokontrakt WITH .F.'                    
     .butDel.Visible=.F.     
     .butRet.Visible=.F.     
     
     DO addShape WITH 'oPage5',2,.Shape1.Left,.cont1.Top+.cont1.Height+15,100,.Shape1.Width,8 
     
     DO adTboxAsCont WITH 'opage5','txtKat',.Shape2.Left+10,.Shape2.Top+10,RetTxtWidth('Wприсвоенная катогория'),dHeight,'категория',2,1
     DO  addComboMy WITH 'opage5',3,.txtKat.Left,.txtKat.Top+.txtKat.Height-1,dheight,.txtKat.Width,.T.,'strKval','ALLTRIM(curSprKval.name)',6,.F.,'newKval=curSprKval.kod',.F.,.T. 
      
     
     DO adTboxAsCont WITH 'opage5','txtDKat',.txtKat.Left+.txtKat.Width-1,.txtKat.Top,RetTxtWidth('Wдата присвоения'),dHeight,'дата присвоения',2,1
     DO adTboxNew WITH 'opage5','boxDKat',.comboBox3.Top,.txtDKat.Left,.txtDKat.Width,dHeight,'newDKval',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'opage5','txtNord',.txtDKat.Left+.txtDKat.Width-1,.txtKat.Top,RetTxtWidth('W№ приказа'),dHeight,'№ приказа',2,1
     DO adTboxNew WITH 'opage5','boxNOrd',.comboBox3.Top,.txtNOrd.Left,.txtNOrd.Width,dHeight,'newNOrdKval',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'opage5','txtDord',.txtNord.Left+.txtNord.Width-1,.txtKat.Top,RetTxtWidth('Wдата приказа'),dHeight,'дата приказа',2,1
     DO adTboxNew WITH 'opage5','boxDOrd',.comboBox3.Top,.txtDOrd.Left,.txtDOrd.Width,dHeight,'newDOrdKval',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'opage5','txtSpec',.txtDord.Left+.txtDord.Width-1,.txtKat.Top,.Shape2.Width-.txtDord.Left-.txtDord.Width,dHeight,'специальность',2,1
     DO  addComboMy WITH 'opage5',4,.txtSpec.Left,.combobox3.Top,dheight,.txtSpec.Width,.T.,'newNKval','ALLTRIM(sprspec.name)',6,.F.,'newNKval=sprspec.name',.F.,.T.  
     .AddObject('grdJob','gridMynew') 
     WITH .grdJob
          .RecordSourceType=1
          .Height=.headerHeight+.RowHeight*4                          
          .Width=.Parent.Shape2.Width-20
          .Left=.Parent.txtKat.Left
          .Top=.Parent.comboBox3.Top+.Parent.comboBox3.Height-1
          .ScrollBars=2
          .ColumnCount=0
           DO addColumnToGrid WITH 'oPage5.grdJob',6          
          .RecordSource='curJobCard'
          .Column1.ControlSource="IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,'')"
          .Column2.ControlSource="IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')"
          .Column3.ControlSource='kse'
          .Column4.ControlSource="IIF(SEEK(tr,'sprtype',1),sprtype.name,'')"
          .Column5.ControlSource='lkv'
          .Column3.Width=RetTxtWidth('999.999')  
          .Column4.Width=RetTxtWidth('внеш.совм.')  
          .Column5.Width=RetTxtWidth('w!w')                             
          .Columns(.ColumnCount).Width=0
          .Column2.Width=(.Width-.Column3.width-.Column4.Width-.Column5.Width)/2
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='подразделение'
          .Column2.Header1.Caption='должность'
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'         
          .Column5.Header1.Caption='!'      
                             
          .Column1.Alignment=0
          .Column2.Alignment=0   
          .Column4.Alignment=0  
          .Column5.Alignment=0  
          
          
          .Column5.AddObject('checkColumn5','checkContainer')
          .Column5.checkColumn5.AddObject('checkMy','checkMy')
          .Column5.CheckColumn5.checkMy.Visible=.T.
          .Column5.CheckColumn5.checkMy.Caption=''
          .Column5.CheckColumn5.checkMy.Left=5
          .Column5.CheckColumn5.checkMy.Top=3
          .Column5.CheckColumn5.checkMy.BackStyle=0
          IF parUv
             .Column5.CheckColumn5.checkMy.Enabled=.F.
          ENDIF 
          .Column5.CheckColumn5.checkMy.ControlSource='curJobCard.lkv'                                                                                                  
          .column5.CurrentControl='checkColumn5'
          .column5.checkColumn5.checkMy.procValid='DO validjobkf'
          .SetAll('Enabled',.F.,'ColumnMy')
          .Column5.Enabled=.T. 
          .Column5.Sparse=.F. 
                   
          .Columns(.ColumnCount).Enabled=.T.      
     ENDWITH
     DO gridSizeNew WITH 'oPage5','grdJob','shapeingrid',.T.
     FOR i=1 TO .grdJob.columnCount 
         .grdJob.Columns(i).Backcolor=oPage5.BackColor           
         .grdJob.Columns(i).DynamicBackColor='IIF(RECNO(oPage5.grdJob.RecordSource)#oPage5.grdJob.curRec,oPage5.BackColor,dynBackColor)'
         .grdJob.Columns(i).DynamicForeColor='IIF(RECNO(oPage5.grdJob.RecordSource)#oPage5.grdJob.curRec,dForeColor,dynForeColor)'        
     ENDFOR
     
     DO adTboxAsCont WITH 'opage5','txtDatt',.Shape2.Left,.grdJob.Top+.grdJob.Height+10,RetTxtWidth('wдата аттестации на соответствиеw'),dHeight,'дата аттестации на соответствие',2,1
     DO adTboxNew WITH 'opage5','boxDatt',.txtDatt.Top,.txtNum.Left,RetTxtWidth('99/99/999999'),dHeight,'people.datts',.F.,.T.,0,.F.
     .txtDatt.Left=.Shape1.Left+(.Shape1.Width-.txtDatt.Width-.boxDatt.Width)/2
     .boxDatt.Left=.txtDatt.Left+.txtDatt.Width-1
     .Shape2.Height=.txtKat.Height*3+.grdJob.Height+30  
     *--------------------------------Кнопка сохранить-------------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','butSkat',.Shape2.Left+(.Shape1.Width-(RetTxtWidth('wсохранитьw')*2)-20)/2,.Shape2.Top+.Shape2.Height+15,RetTxtWidth('WсохранитьW'),dHeight+5,'сохранить','DO saveKval'
     *---------------------------------Кнопка удалить ------------------------------------------------------------------------
     DO addcontlabel WITH 'opage5','butDKat',.butSkat.Left+.butSkat.Width+20,.butSkat.Top,.butSkat.Width,dHeight+5,'удалить','DO delKvalPeop'   
     .SetAll('Alignment',2,'myTxtBox')
     .Refresh
     IF parUv
        .SetAll('Enabled',.F.,'MyCommandButton')
        .SetAll('Enabled',.F.,'MyTxtBox')
        .SetAll('Enabled',.F.,'ComboMy')
        .SetAll('DisabledForeColor',RGB(1,0,0),'ComboMy')
     ENDIF
 ENDWITH                 
***********************************************************************************************************************
PROCEDURE saveKval
SELECT people
REPLACE kval WITH newKval,dKval WITH newDKval,nordKval WITH newNordKval,dordKval WITH newDOrdKval,nKval WITH newNKval
SELECT curJobCard
SCAN ALL
     DO validJobkf   
     SELECT curJobCard
ENDSCAN
GO TOP
SELECT people
***********************************************************************************************************************
PROCEDURE validJobkf
IF SEEK(curJobCard.nid,'datjob',7)
   REPLACE datjob.lkv WITH curJobCard.lkv,datjob.kv WITH people.kval                   
   IF datjob.lkv  
      REPLACE datjob.nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' №'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' от '+DTOC(people.dkval),''),;
              datjob.pkat WITH IIF(INLIST(datjob.kat,1,2,5,7).AND.SEEK(kv,'sprkval',1),sprkval.doplkat,0) 
      DO CASE 
         CASE datjob.kv=1 && высшая категория
              REPLACE datjob.kf WITH IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.kf3,datjob.kf),datjob.namekf WITH sprdolj.namekf3
         CASE datjob.kv=2 && первая категория 
              REPLACE datjob.kf WITH IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.kf2,datjob.kf),datjob.namekf WITH sprdolj.namekf2
         CASE datjob.kv=3 && вторая категория
              REPLACE datjob.kf WITH IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.kf1,datjob.kf),datjob.namekf WITH sprdolj.namekf1
         CASE datjob.kv=0 && без категории
              REPLACE datjob.kf WITH IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.kf,datjob.kf),datjob.namekf WITH sprdolj.namekf     
      ENDCASE
   ELSE      
      REPLACE datjob.kf WITH IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.kf,datjob.kf),datjob.namekf WITH sprdolj.namekf,datjob.nprik WITH ' ',datjob.pkat WITH IIF(INLIST(datjob.kat,1,2,5,7),5,0)                               
   ENDIF
ENDIF
SELECT curJobCard
***********************************************************************************************************************
PROCEDURE delKvalPeop
SELECT people
newKval=0
newDKval=CTOD('   .  .    ')
newDOrdKval=CTOD('   .  .    ')
newNOrdKval=''
*REPLACE kval WITH 0,dkval WITH CTOD('  .  .    '),nordKval WITH '',dordKval WITH CTOD('  .  .    '),nkval WITH ''

SELECT people 
***********************************************************************************************************************
PROCEDURE saveKontrakt
repTime=''
*parDod - код документа
*parTime - код времени
*parSrok - срок
*parBeg - начало
*pareEnd - окончание 
*parFio - фамилия
*parStr - срок(неопределенный)
*parPers - процент по контракту
*parDayOtp - дней отпуска
*parDayKont - поощрительный отпуск
*parDayVred - отпуск за вредностьб
*parDayNorm - отпуск за ненормированный рабочий день
*parTotday  - всего дней
SELECT people
REPLACE dog WITH newVidDog,begdog WITH newBegDog,enddog WITH newEndDog,ktime WITH newKTime,strtime WITH strsrok,;
        dayOtp WITH newDayOtp,dayKont WITH newDayKont,dayNorm WITH newDayNorm,dayVred WITH newDayVred,totDay WITH newTotDay,pKont WITH newPkont,numDog WITH newNumDog,ddog WITH newdDog
*DO changeRowGrdPers        
***********************************************************************************************************************
PROCEDURE delKont   
WITH oPage5             
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F.
     .SetAll('Enabled',.F.,'myTxtBox')
     .SetAll('Enabled',.F.,'comboMy')
     .butDel.Visible=.T.
     .butRet.Visible=.T.
     .Refresh
ENDWITH  

*************************************************************************************************************************
PROCEDURE delInfoKontrakt    
PARAMETERS par1
IF par1
   SELECT people
   REPLACE dog WITH 0,begdog WITH CTOD('  .  .    '),enddog WITH CTOD('  .  .    '),ktime WITH 0,strtime WITH '',timedog WITH 0,;
           pkont WITH 0,dayOtp WITH 0,dayKont WITH 0,dayVred WITH 0,dayNorm WITH 0,totDay WITH 0,numDog WITH '',dDog WITH CTOD(' .  .    ')
   DO changeRowGrdPers          
ENDIF 
newVidDog=people.dog      && тип договора (контракт,трудовой договор,срочный трудовой договор)
strVid=IIF(SEEK(people.dog,'sprdog',1),sprdog.name,'')
strSrok=IIF(SEEK(people.kTime,'cursrok',1),cursrok.name,'')

newNumDog=people.numDog
newDDog=people.ddog
newSrok=people.timeDog    && срок договора
newkTime=people.kTime     && код времени
newBegDog=people.begDog   && начало контракта
newEndDog=people.endDog   && окончание контракта
newPkont=people.pKont     && процент по контракту


newTotDay=people.totDay   && всего дней отпуска
newDayOtp=people.dayOtp   && основной отпуск
newDayKont=people.dayKont && поощрительный отпуск
newDayVred=people.dayVred && за вредность
newDayNorm=people.dayNorm && за ненормированный
WITH oPage5             
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.
     .SetAll('Enabled',.T.,'myTxtBox')
     .SetAll('Enabled',.T.,'comboMy')
     .butDel.Visible=.F.
     .butRet.Visible=.F.
     .Refresh
ENDWITH     
****************************************************************************************************************************************************
PROCEDURE formKontraktPrn
** kontrakt.dot -шаблон контракта
DIMENSION dimTkont(3)
STORE 0 TO dimTkont
dimTKont(1)=1
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Печать контракта'
     .Icon='kone.ico'
     DO addShape WITH 'fSupl',1,10,10,20,400,8 
     DO addOptionButton WITH 'fSupl',11,'контракт днев.',.Shape1.Top+20,.Shape1.Left+20,'dimTKont(1)',0,"DO procValOption WITH 'fSupl','dimTkont',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'контракт смен.',.Option11.Top,.Option11.Left+.Option11.Width+20,'dimTkont(2)',0,"DO procValOption WITH 'fSupl','dimTkont',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'трудовой дог-р',.Option11.Top,.Option11.Left+.Option11.Width+20,'dimTkont(3)',0,"DO procValOption WITH 'fSupl','dimTkont',3",.T. 
     
     .Option11.Left=.Shape1.Left+(.Shape1.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10 
     .Option13.Left=.Option12.Left+.Option12.Width+10 
     .Shape1.Height=.Option11.Height+40
     *--------------------------------Кнопка сохранить-------------------------------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','butPrn',.Shape1.Left+(.Shape1.Width-(RetTxtWidth('wвозвратw')*2)-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wвозвратw'),dHeight+5,'печать','DO prnkontrakt'
     *---------------------------------Кнопка удалить ------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+20,.butPrn.Top,.butPrn.Width,dHeight+5,'возврат','fSupl.Release'          
     .Height=.Shape1.Height+.butPrn.Height+50
     .Width=.Shape1.Width+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show

********************************************************************************************************************
PROCEDURE prnkontrakt
PARAMETERS pardkont,parfio,parfio2,pardol,parpodr,parsrok,parktime,parbeg,parend,parPers,parOsnday,parKontDay,parVredDay,parNormDay,parTotDay
*pardkont - дата заключения
*parfio - ФИО (кто)
*parfio2 - ФИО с кем
*pardol - должность
*parpodr - подразделение
*parsrok - срок
*parktime - наименование времени
*parbeg - начало
*parend - окончание
*parPers - процент надбавки за контракт
*parOsnDay - основной отпуск
*parKontDay - поощрительный отпуск
*parVredDay -  за вредность
*parNormday - за ненормированный рабочий день
*parTotDay - всего дней

LOCAL loWord, loDoc 
***** Поля, заполняемые в шаблоне контракта
* dkont-дата заключения
* fio-ФИО в именительном падеже
* fio2-ФИО в именительном падеже
* dol-должность
* podr-прдразделение
* srok-срок
* period-период с по, указывается в формате -  с "01" месяц 1999г. по "01" месяц 1999г.

objWord=CREATEOBJECT('WORD.APPLICATION')
DO CASE
   CASE dimTkont(1)=1
        pathdot=ALLTRIM(datset.pathword)+'kontrakt.dot'
   CASE dimTkont(2)=1
        pathdot=ALLTRIM(datset.pathword)+'kontrakt2.dot'     
  CASE dimTkont(3)=1
        pathdot=ALLTRIM(datset.pathword)+'kontrakt3.dot'          
ENDCASE

nameDoc=objWord.Documents.Add(pathdot)   
* Подставляем необходимые значения в поля 
IF TYPE([nameDoc.formFields("tday")])="O"
   nameDoc.FormFields("tday").Result=LTRIM(STR(DAY(people.ddog)))
ENDIF

IF TYPE([nameDoc.formFields("tmonth")])="O"
   nameDoc.FormFields("tmonth").Result=IIF(!EMPTY(people.ddog),ALLTRIM(month_prn(MONTH(people.ddog))),'')
ENDIF

IF TYPE([nameDoc.formFields("tyear")])="O"
   nameDoc.FormFields("tyear").Result=STR(YEAR(people.ddog),4) 
ENDIF
  
IF TYPE([nameDoc.formFields("tfio")])="O"
   nameDoc.FormFields("tfio").Result=ALLTRIM(people.fio)
ENDIF
IF TYPE([nameDoc.formFields("tfiod")])="O"
   IF dimTkont(3)#1
       nameDoc.FormFields("tfiod").Result=IIF(!EMPTY(people.fiov),ALLTRIM(people.fiov),ALLTRIM(people.fio))
   ELSE   
      nameDoc.FormFields("tfiod").Result=IIF(!EMPTY(people.fior),ALLTRIM(people.fior),ALLTRIM(people.fio))     
   ENDIF 
ENDIF
SELECT curjobsupl
oldrecsupl=RECNO()
LOCATE FOR tr=1
nkpodr=curjobsupl.kp
nkdol=curjobsupl.kd
ON ERROR DO ersup

   GO oldrecsupl
ON ERROR 
SELECT people
*nkpodr=IIF(SEEK(STR(people.num,4)+STR(1,1),'datjob',4).OR.SEEK(STR(people.num,4)+STR(3,1),'datjob',4),datjob.kp,0)
*nkdol=IIF(SEEK(STR(people.num,4)+STR(1,1),'datjob',4).OR.SEEK(STR(people.num,4)+STR(3,1),'datjob',4),datjob.kd,0)

cnpodr=IIF(SEEK(nkpodr,'sprpodr',1),IIF(!EMPTY(sprpodr.namek),LOWER(sprpodr.namek),LOWER(sprpodr.nameord)),'')
IF dimTKont(3)=1
   cndol=IIF(SEEK(nkdol,'sprdolj',1),IIF(!EMPTY(sprdolj.namet),LOWER(sprdolj.namet),LOWER(sprdolj.name)),'')
   cnpodr=IIF(SEEK(nkpodr,'sprpodr',1),sprpodr.namework,'')
ELSE 
   cndol=IIF(SEEK(nkdol,'sprdolj',1),IIF(!EMPTY(sprdolj.namet),LOWER(sprdolj.namet),LOWER(sprdolj.name)),'')
ENDIF 
IF TYPE([nameDoc.formFields("tdol")])="O"
   nameDoc.FormFields("tdol").Result=ALLTRIM(cndol)
ENDIF

IF TYPE([nameDoc.formFields("tpodr")])="O"
   nameDoc.FormFields("tpodr").Result=ALLTRIM(cnpodr)
ENDIF

IF TYPE([nameDoc.formFields("tsrok")])="O"
   nameDoc.FormFields("tsrok").Result=ALLTRIM(STR(people.ktime))
ENDIF

strDateBeg=IIF(!EMPTY(people.begdog),dateToString('people.begdog',.T.),'')
IF TYPE([nameDoc.formFields("tbeg")])="O"
   nameDoc.FormFields("tbeg").Result=strDatebeg
ENDIF
strDateEnd=IIF(!EMPTY(people.enddog),dateToString('people.enddog',.T.),'')
IF TYPE([nameDoc.formFields("tend")])="O"
   nameDoc.FormFields("tend").Result=strDateEnd
ENDIF

IF TYPE([nameDoc.formFields("pkont")])="O"
   nameDoc.FormFields("pkont").Result=LTRIM(STR(people.pkont))+'%'
ENDIF

IF TYPE([nameDoc.formFields("otptot")])="O"
   nameDoc.FormFields("otptot").Result=LTRIM(STR(people.totday))   
ENDIF

IF TYPE([nameDoc.formFields("otposn")])="O"
   nameDoc.FormFields("otposn").Result=LTRIM(STR(people.dayotp))   
ENDIF

IF TYPE([nameDoc.formFields("otpkont")])="O"
   nameDoc.FormFields("otpkont").Result=LTRIM(STR(people.daykont))   
ENDIF

IF TYPE([nameDoc.formFields("otpvred")])="O"
   nameDoc.FormFields("otpvred").Result=LTRIM(STR(people.dayvred))   
ENDIF

IF TYPE([nameDoc.formFields("strdbeg")])="O"
   nameDoc.FormFields("strdbeg").Result=IIF(!EMPTY(people.begdog), STR(DAY(people.begdog),2)+' '+month_prn(MONTH(people.begdog))+' '+STR(YEAR(people.begdog),4),'')   
ENDIF

IF TYPE([nameDoc.formFields("strdend")])="O"
   nameDoc.FormFields("strdend").Result=IIF(!EMPTY(people.enddog), STR(DAY(people.enddog),2)+' '+month_prn(MONTH(people.enddog))+' '+STR(YEAR(people.enddog),4),'')   
ENDIF

IF TYPE([nameDoc.formFields("absemp")])="O"
   nameDoc.FormFields("absemp").Result=''  
ENDIF
objWord.Visible=.T.
***************************************************************************************************************************************************
PROCEDURE exitKontraktPrn
WITH oPage5            
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.     
     .SetAll('Enabled',.T.,'myTxtBox')
     .SetAll('Enabled',.T.,'comboMy')      
     .butPrn.Visible=.F.
     .butRetPrn.Visible=.F.           
     .Option1.Visible=.F.
     .Option2.Visible=.F.
     .Option3.Visible=.F.
     .Refresh
ENDWITH 
***************************************************************************************************************************************************
PROCEDURE procValOption
PARAMETERS parFrm,parDim,parNum
STORE 0 TO &parDim
&parDim(parNum)=1
&parFrm..Refresh
************************************************************************************************************************
PROCEDURE validtotday
PARAMETERS parFrm,parVar
*parFrm - форма
*parvar - переменная
&parVar
&parFrm..Refresh     