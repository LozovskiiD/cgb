IF USED('peopold')
   SELECT peopold
   USE
ENDIF
IF USED('curnew')
   SELECT curnew
   USE
ENDIF
log_count=.F.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl   
     .Caption='Расчёт надбавки за сложность и напряжённость работы'
     DO addshape WITH 'fSupl',1,20,20,150,450,8 
     DO adLabMy WITH 'fSupl',1,'Выполнение данной процедуры приведет к расчёту',.Shape1.Top+10,.Shape1.Left+5,.Shape1.Width-10,2 
     DO adLabMy WITH 'fSupl',2,'надбавки за сложность и напряжённость работы.',.lab1.Top+.lab1.Height,.Shape1.Left+5,.lab1.Width,2                                   
     DO adLabMy WITH 'fSupl',3,'Другие надбавки и дрплаты затронуты не будут.',.lab2.Top+.lab2.Height,.Shape1.Left+5,.lab1.Width,2   
     
     DO adLabMy WITH 'fSupl',4,'для подтверждения ваших намерений',.lab3.Top+.lab3.Height+10,.Shape1.Left+5,.Shape1.Width-10,2 
     DO adLabMy WITH 'fSupl',5,'поставьте птичку в окошке, расположенном ниже',.lab4.Top+.lab4.Height,.Shape1.Left+5,.lab1.Width,2                                   
     DO adCheckBox WITH 'fSupl','check1','подтвердить расчёт',.lab5.Top+.lab5.Height+10,.Shape1.Left,150,dHeight,'log_count',0    
     
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     .Shape1.Height=.check1.Height+.lab1.Height*5+40
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wвозвратw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wвозвратw'),dHeight+3,'расчёт','DO countnadsl'
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fSupl.Release'     
     
     DO addcontlabel WITH 'fSupl','cont3',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wвозвратw'))/2,.cont1.Top,RetTxtWidth('wвозвратw'),dHeight+3,'возврат','fSupl.Release'
     .cont3.Visible=.F.     
      DO addShape WITH 'fSupl',2,.Shape1.Left,.cont1.Top,.cont1.Height,.Shape1.Width,8
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
     DO addShape WITH 'fSupl',3,.Shape2.Left,.Shape2.Top,.Shape2.Height,50,8
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.               
      DO adLabMy WITH 'fSupl',25,'100%',.Shape2.Top+2,.Shape2.Left,.Shape2.Width,2,.F.,0
     .lab25.Visible=.F.       
     
     
     .Width=.Shape1.Width+40   
     .Height=.Shape1.Height+.cont1.Height+60                                  
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************
PROCEDURE countnadsl
IF !log_count
   RETURN 
ENDIF 
IF !FILE('pathold.mem')
   newpathold=GETDIR('','','Укажите путь к каталогу',64)
   pathold=IIF(!EMPTY(newpathold),newpathold,pathold)
   var_path=FULLPATH('tarif.fxp')    
   var_pathold=LEFT(var_path,LEN(var_path)-9)+'pathold'  
   SAVE TO &var_pathold ALL LIKE pathold     
   FLUSH  
ENDIF
IF !FILE('pathold.mem')
   RETURN
ELSE
   RESTORE FROM pathold ADDITIVE   
ENDIF
STORE 0 TO max_rec,one_pers,pers_ch,sumtotold,sumtotnew
peopold=pathold+'people.dbf'
USE &peopold ALIAS peopold IN 0
SELECT people
oldRecPeop=RECNO()
SELECT datJob
SET FILTER TO 
COUNT TO max_rec
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .Shape2.Visible=.T.
     .Shape3.Visible=.T.
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape3.Width=1
ENDWITH

SCAN ALL
     SELECT peopold
     LOCATE FOR tab=datjob.tabn.AND.kp=datjob.kp.AND.kse=datjob.kse.AND.tr=datjob.tr
     IF !FOUND()
        LOCATE FOR ALLTRIM(name)=ALLTRIM(people.fio).AND.kp=datjob.kp.AND.kse=datjob.kse.AND.tr=datjob.tr
     ENDIF
     sumtotold=peopold.sfond && фонд по старым условиям
     SELECT datJob     
     sumtotnew=msf-mslwork-mhigh && сумма но новым условиям без надбавки за сложность и напряжённость
     IF sumtotold#0.AND.sumtotnew#0.AND.mtokl#0
        sumdif=sumtotold-sumtotnew
        persNum=sumdif/datjob.mtokl*100  
        IF persNum<=200
           REPLACE mslWork WITH sumDif,pSlWork WITH persNum,slWork WITH sumdif
        ELSE 
           persHigh=persnum-200
           persNum=200
           REPLACE pSlWork WITH persNum,pHigh WITH persHigh
        ENDIF   
     ENDIF 
     one_pers=one_pers+1
     pers_ch=one_pers/max_rec*100
     fSupl.shape3.Visible=.T.
     fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
     fSupl.Shape3.Width=fSupl.shape2.Width/100*pers_ch 
ENDSCAN
=SYS(2002)
=INKEY(2)
GO oldRecPeop
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus    
WITH fSupl           
     .shape2.Visible=.F.
     .shape3.Visible=.F.
     .lab25.Caption='Расчёт выполнен' 
     .lab25.Top=.Shape1.Top+.Shape1.Height+10
     .cont3.Top=.lab25.Top+.lab25.Height    
     .cont3.Visible=.T.
ENDWITH  
