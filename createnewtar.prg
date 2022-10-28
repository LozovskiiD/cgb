IF !('REAL'$UPPER(datset.pathlast))
   RETURN
ENDIF
fSupl=CREATEOBJECT('FORMSUPL')
newDateTar=CTOD(' .  .    ')
logNewTar=.F.
newPrim=''
labname=''
varOffice=datShtat.office
varBoss=datshtat.boss
varAdres=datshtat.adres
WITH fSupl
     .Caption='Создание штатного расписания'   
     DO addShape WITH 'fSupl',1,20,20,dHeight,300,8   
     DO adtBoxAsCont WITH 'fSupl','contDate',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('Wсоздать тарификацияю наW'),dHeight,'создать тарификацию на',1,1 
     DO addtxtboxmy WITH 'fSupl',1,.contDate.Left+.contDate.Width-1,.contDate.Top,RetTxtWidth('99/99/999999'),.F.,'newDateTar'
     DO adtBoxAsCont WITH 'fSupl','contName',.contDate.Left,.contDate.Top+.contDate.Height-1,.contDate.Width+.txtBox1.Width-1,dHeight,'примечание',2,1 
     DO addtxtboxmy WITH 'fSupl',2,.contName.Left,.contName.Top+.contName.Height-1,.contName.Width,.F.,'newPrim'
     DO adCheckBox WITH 'fSupl','checkNew','подтверждение намерений',.txtBox2.Top+.txtBox2.Height+20,0,150,dHeight,'logNewTar',0      
     .Shape1.Width=.contDate.Width+.txtBox1.Width+40
     .Shape1.Height=.contDate.Height*3+.checkNew.Height+60
     .checkNew.Left=.Shape1.Left+(.Shape1.Width-.checkNew.Width)/2  
    
     *-----------------------------Кнопка применить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприменитьw')*2)-20)/2,;
     .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприменитьw'),dHeight+5,'Применить','DO procBeforeCreate'
   
     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','fSupl.Release','Возврат'  
     
     .Width=.Shape1.Width+40
     DO adLabMy WITH 'fSupl',1,labname,.Shape1.Top+.Shape1.Height+10,.Shape1.Left,.Shape1.Width,2    
     .lab1.Visible=.F.                    
     
     DO addcontlabel WITH 'fSupl','contNew',(.Width-RetTxtWidth('wперезаписать')*3-20)/2,.lab1.Top+.lab1.Height,RetTxtWidth('wперезаписать'),dHeight+3,'добавить','DO asktarifnew WITH 1'         
     DO addcontlabel WITH 'fSupl','contRew',.ContNew.Left+.ContNew.Width+10,.ContNew.Top,.ContNew.Width,dHeight+3,'перезаписать','DO asktarifnew WITH 2'         
     DO addcontlabel WITH 'fSupl','contRet',.ContRew.Left+.ContRew.Width+10,.ContNew.Top,.ContNew.Width,dHeight+3,'возврат','fSupl.Release'  
         
     .contNew.Visible=.F.
     .contRew.Visible=.F.
     .contRet.Visible=.F.   
     
     
     DO addShape WITH 'fSupl',11,.Shape1.Left,.cont1.Top,.cont1.Height,.Shape1.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape1.Left,.Shape1.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.   
     .Height=.Shape1.Height+.cont1.Height+60       
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***********************************************************************************************************************************
PROCEDURE procBeforeCreate
IF !logNewTar.OR.EMPTY(newDateTar)
   RETURN 
ENDIF
SELECT datshtat
LOCATE FOR !real.AND.dtarif=newDateTar
reppathtar='TAR'+DTOC(newDateTar)
IF !FOUND()
   pathcopy=pathmain+'\TAR'+DTOC(newDateTar)  
   DO createtar WITH .T.
   fSupl.Release
ELSE 
   WITH fSupl   
        .lab1.Visible=.T.
        labname='тарификация на '+DTOC(newDatetar)+' уже создана!'
        .lab1.Caption=labname
        .lab1.Width=.Shape1.Width
        .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width)/2
        .cont1.Visible=.F.
        .cont2.Visible=.F.
        .contNew.Visible=.T.
        .contRew.Visible=.T.
        .contRet.Visible=.T.
   ENDWITH    
ENDIF
*******************************************************************************************************************************************
*
*******************************************************************************************************************************************
PROCEDURE asktarifnew
PARAMETERS par1
*1 - новая
*2 - перезаписать
DO CASE 
   CASE par1=1
        ncx=1
        DO WHILE .T.
           nametarsup='TAR'+DTOC(newDateTar)+'_'+LTRIM(STR(ncx))
           LOCATE FOR ALLTRIM(pathtarif)=nametarsup
           IF !FOUND()
               EXIT
           ENDIF
           ncx=ncx+1
        ENDDO
        pathcopy=pathmain+'\'+nametarsup 
        reppathtar=nametarsup
        DO createtar WITH .T.
       *fSupl.Release
   CASE par1=2
        pathcopy=pathmain+'\TAR'+DTOC(newDateTar)  
        reppathtar='TAR'+DTOC(newDateTar)
        DO createtar WITH .F.
        *fSupl.Release
ENDCASE 
***********************************************************************************************************************************
*                    Непосредственно создание  тарификации
***********************************************************************************************************************************
PROCEDURE createtar
PARAMETERS lognew
WITH fSupl
     .lab1.Visible=.F.
     .SetAll('Visible',.F.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%'     
ENDWITH   

IF lognew
   MKDIR &pathcopy
ELSE
   SELECT datshtat
   LOCATE FOR ALLTRIM(pathtarif)=reppathtar
ENDIF    
tarcopy=pathcur+'*.*'
RUN XCOPY /Y &tarcopy &pathcopy >nul 
SELECT datshtat
IF lognew
   APPEND BLANK
ENDIF    
REPLACE dtarif WITH newDatetar,pathtarif WITH reppathtar, basest WITH varBaseSt,dcreate WITH DATETIME(),fullname WITH newPrim,;
        office WITH varOffice,boss WITH varBoss,adres WITH varAdres,luse WITH .T.,nmzp WITH varNmzp

varDtar=newDateTar
varDBaseSt=datShtat.baseSt
pathtarif=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\'+';'+pathmain+';'+pathsupl
pathTarSupl=ALLTRIM(DatShtat.pathtarif)
pathcur=pathmain+'\'+ALLTRIM(datshtat.pathtarif)+'\' && путь к каталогу тарификации с которой сейчас работают

tarDateSay='тарификация на '+DTOC(varDTar)+' '+ALLTRIM(datShtat.fullname)+' (изменить - двойной щелчок мыши)'
frmTop.contDateTarif.ContLabel.Caption=tarDateSay
SELECT people
USE
SELECT rasp
USE
SELECT sprpodr
USE
SELECT datjob
USE
SELECT tarfond
USE
SELECT boss
USE
SELECT sprdolj
USE
SELECT curSupFond
SET FILTER TO 
DELETE ALL
IF USED('datagrup')
   SELECT datagrup
   USE 
ENDIF
SET PATH TO &pathtarif
USE people ORDER 1 IN 0
USE peopout IN 0
USE datjobout IN 0
SELECT people
APPEND FROM peopout FOR date_out>=newDateTar
REPLACE date_out WITH CTOD('  .  .   ') FOR date_out>=newDatetar
DELETE FOR !EMPTY(date_out).AND.date_out<newDateTar    && удаляем записи где дата увольнения меньще даты тарификации
DELETE FOR date_in>newDateTar    && удаляем записи где дата приема больше даты тарификации
DELETE FOR dekotp.AND.bdekotp<newDateTar

USE rasp ORDER 2 IN 0
USE datjob ORDER 4 IN 0
SELECT datJob
APPEND FROM datjobout FOR dateOut>=newDatetar
DELETE FOR kd=0
REPLACE dateuv WITH IIF(SEEK(kodpeop,'people',1),people.date_out,dateuv) ALL
DELETE FOR dateBeg>newDateTar    && удаляем записи где дата начала работы больше даты тарификации
DELETE FOR !EMPTY(dateOut).AND.EMPTY(dateuv).AND.dateOut<=newDateTar    && удаляем записи где дата окончания<=работы меньще даты тарификации при пустой дате увольнения
DELETE FOR !EMPTY(dateuv).AND.dateuv<newDateTar  &&дата увольнения<даты тарификации при заполненной дате увольнения
*DELETE FOR dateOut>=newDateTar.AND.EMPTY(dateuv) &&дата окончания>=дате тарификации при пустой дате увольнения



*DELETE FOR dateBeg>dateBook && дата начала больше даты книги
*DELETE FOR !EMPTY(dateOut).AND.EMPTY(dateuv).AND.dateOut<=dateBook  &&дата окончания<=дате книги при пустой дате увольнения
*ELETE FOR !EMPTY(dateuv).AND.dateuv<dateBook  &&дата увольнения<даты книги при заполненной дате увольнения
*DELETE FOR dateOut>=dateBook.AND.EMPTY(dateuv) &&дата окончания>=дате книги при пустой дате увольнения
*DELETE FOR !EMPTY(dateuv).AND.dateuv>dateBook  &&дата увольнения>даты книги при заполненной дате увольнения

DELETE FOR INLIST(tr,4,6)                  && удаляем записи по совмещению% tr=4 и замещению (на время отпуска, больничного и т.д.) tr=6
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) FOR kat=0
SELECT people
DELETE FOR !SEEK(num,'datjob',1)
SELECT datjob
DELETE FOR !SEEK(kodpeop,'people',1)
REPLACE dateout WITH CTOD('  .  .    ') ALL
SELECT peopout
USE
SELECT datjobout
USE
USE sprpodr ORDER 1 IN 0
USE tarfond ORDER 1 IN 0
USE boss IN 0
USE sprdolj IN 0 ORDER 1
SELECT * FROM sprdolj INTO CURSOR curSprDolj READWRITE ORDER BY name
SELECT * FROM sprpodr INTO CURSOR curSprPodr READWRITE ORDER BY name
SELECT tarfond
SCAN ALL
     IF nBlock=2
        SCATTER TO dimSup
        SELECT curSupFond
        APPEND BLANK
        GATHER FROM dimSup       
     ENDIF
     SELECT tarfond
ENDSCAN
SELECT tarfond
SELECT * FROM tarfond WHERE ltar.AND.!EMPTY(plrep) INTO CURSOR curnewfond

SELECT datJob
SET RELATION TO STR(kp,3)+STR(kd,3) INTO rasp ADDITIVE
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec
SCAN ALL
     SELECT people 
     SEEK datjob.kodpeop
   *  SELECT datjob
   *  REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in,pkont WITH IIF(tr=1,people.pkont,0),dekotp WITH .F.
   *  DO CASE 
   *     CASE datjob.lkv.AND.people.kval#0
   *          REPLACE kv WITH people.kval,nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' №'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' от ' +DTOC(people.dkval),''),;
   *                  pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,0)
   *     OTHERWISE
   *          REPLACE kv WITH 0,nPrik WITH '',pkat WITH 0   
   *  ENDCASE 
   *  DO CASE
   *     CASE kv=0
   *          REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf,kf),namekf WITH sprdolj.namekf
   *     CASE kv=1
   *          REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf3,kf),namekf WITH sprdolj.namekf3
   *     CASE kv=2
   *          REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf2,kf),namekf WITH sprdolj.namekf2
   *     CASE kv=3
   *          REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf1,kf),namekf WITH sprdolj.namekf1
   *  ENDCASE   
     DO repNadJob 
     SELECT datJob
     one_pers=one_pers+1
     pers_ch=one_pers/max_rec*100
     fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
     fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch
ENDSCAN
SELECT curnewfond
USE
SELECT tarfond
SET FILTER TO 
SELECT rasp 
SET ORDER TO 1
SELECT datjob
SET RELATION TO
SET RELATION TO tr INTO curSprType ADDITIVE

GO TOP 
SELECT people
GO TOP
fSupl.Visible=.F.
WITH frmTop
     WITH .grdPers          
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='people'
           DO addColumnToGrid WITH 'frmTop.grdPers',3
          .Column1.ControlSource='people.num'
          .Column2.ControlSource='" "+people.fio'       
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Фамилия Имя Отчество'  
          .Column1.Width=RettxtWidth(' 1234 ')      
          .Columns(.ColumnCount).Width=0          
          .Column2.Width=.Width-.column1.width-SYSMETRIC(5)-13-.ColumnCount 
          .Column1.Alignment=1        
          .Column2.Alignment=0
          .Column1.Movable=.F.         
          .colNesInf=2    
          .procAfterRowColChange='DO cardTarifScr'  
         .SetAll('BOUND',.F.,'Column')  
         .Visible=.T.             
     ENDWITH
     DO gridSizeNew WITH 'frmTop','grdpers',.F.,.F.,.T.
     WITH .grdJob      
          .ScrollBars=2       	           
          .RecordSourceType=1     
          .RecordSource='datJob'
          DO addColumnToGrid WITH 'frmTop.grdJob',7
          .Column1.ControlSource="IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.name,'')"
          .Column2.ControlSource="IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.name,'')"         
          .Column3.ControlSource='datjob.kse'
          .Column4.ControlSource='datjob.tr'
          .Column5.ControlSource="IIF(SEEK(datjob.kat,'sprkat',1),sprkat.name,'')"    
          
          .Column1.Header1.Caption='подразделение' 
          .Column2.Header1.Caption='должность'
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'
          .Column5.Header1.Caption='персонал'
          .Column6.Header1.Caption='к' 
          
          .Column3.Width=RetTxtWidth('999.999')  
          .Column4.Width=RetTxtWidth('внеш.совм.')                               
          .Column5.Width=RetTxtWidth('wперсонал')
          .Column6.Width=RetTxtWidth('wкw')     
         
          .Columns(.ColumnCount).Width=0
          .Column1.Width=(.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width)/2
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-SYSMETRIC(5)-13-.ColumnCount
                            
          .Column6.AddObject('checkColumn6','checkContainer')
          .Column6.checkColumn6.AddObject('checkLkv','checkMy')
          .Column6.CheckColumn6.checkLkv.Visible=.T.
          .Column6.CheckColumn6.checkLkv.Caption=''
          .Column6.CheckColumn6.checkLkv.Left=6
          .Column6.CheckColumn6.checkLkv.Top=3
          .Column6.CheckColumn6.checkLkv.BackStyle=0
          .Column6.CheckColumn6.checkLkv.ControlSource='datjob.lkv'         
          .Column6.CheckColumn6.checkLkv.procValid='DO validLkv' 
          .Column6.CheckColumn6.checkLkv.Left=(.Column6.Width-SYSMETRIC(15))/2                                                                                         
          .Column6.CurrentControl='checkColumn6'                   
          
          .procAfterRowColChange='DO changeJob'
          .Column1.Alignment=0
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0
          .Column5.Alignment=0          
          .procAfterRowColChange='DO changeJob'
    ENDWITH        
    DO gridSizeNew WITH 'frmTop','grdJob',.F.,.F.,.T.   
    .grdPers.Columns(.grdPers.ColumnCount).SetFocus  
    .Refresh    
ENDWITH
**************************************************************************************************************************
*                       Подстановка набавок и доплат для тарификации
**************************************************************************************************************************
PROCEDURE repNadJob 
*SELECT rasp
*SEEK STR(datjob.kp,3)+STR(datjob.kd,3)
SELECT curnewfond
GO TOP
SCAN ALL
     IF !EMPTY(procrepnad)
        procrep=ALLTRIM(procrepnad)
        DO &procrep 
     ELSE 
        repjob=ALLTRIM(plrep)
        repjob1='rasp.'+ALLTRIM(plrep)
        SELECT datjob 
        REPLACE &repjob WITH &repjob1
     ENDIF 
     IF ALLTRIM(LOWER(curnewfond.plrep))='pkat'.AND.rasp.pkat#0.AND.datjob.kv=0
        SELECT datjob
        REPLACE pkat WITH 5       
     ENDIF
     SELECT curnewfond
ENDSCAN
SELECT datjob 
**************************************************************************************************************************
PROCEDURE repnadvto
SELECT datjob
DO CASE 
   CASE kv=0
        REPLACE pvto WITH rasp.pvto4
   CASE kv=1
        REPLACE pvto WITH rasp.pvto1
   CASE kv=2
        REPLACE pvto WITH rasp.pvto2
   CASE kv=3
        REPLACE pvto WITH rasp.pvto3
ENDCASE 
SELECT curnewfond
************************************************************************************************************************
PROCEDURE repnadmain
repjob=ALLTRIM(curnewfond.plrep)
repjob1='rasp.'+ALLTRIM(curnewfond.plrep)
SELECT datjob 
REPLACE &repjob WITH &repjob1


