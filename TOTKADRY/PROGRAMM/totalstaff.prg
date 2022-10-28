STORE DATE() TO date_Beg
STORE '' TO fltch,fltpodr
logbdj=.F.
DIMENSION dimOption(3)
STORE .F. TO dimOption
*dimOption(1) - организация
*dimOption(2) - совокупность
*dimOption(3) - подразделение
dimOption(1)=.T.

IF !USED('datagrup')
   USE datagrup IN 0
ENDIF 
IF USED('katStaff')
   SELECT katStaff
   USE 
ENDIF

SELECT * FROM datagrup INTO CURSOR dopGroup READWRITE
SELECT datagrup 
USE
SELECT dopGroup
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprkat INTO CURSOR katStaff READWRITE
ALTER TABLE katStaff ADD COLUMN totpeop N(6)
ALTER TABLE katStaff ADD COLUMN totalive N(6)
ALTER TABLE katStaff ADD COLUMN totdek N(6)
ALTER TABLE katStaff ADD COLUMN totvn N(6)
ALTER TABLE katStaff ADD COLUMN totwom N(6)
ALTER TABLE katStaff ADD COLUMN totwomwork N(6)

ALTER TABLE katStaff ADD COLUMN kseshtat N(8,2)
ALTER TABLE katStaff ADD COLUMN ksepeop N(8,2)
ALTER TABLE katStaff ADD COLUMN ksevac N(8,2)
ALTER TABLE katStaff ADD COLUMN totkont N(6)

APPEND BLANK
REPLACE name WITH 'Всего'
INDEX ON kod TAG T1
GO TOP
fStaff=CREATEOBJECT('FORMSUPL')

WITH fStaff
     .Caption='Общее кол-во сотрудников'
     .Icon='kone.ico' 
     DO addshape WITH 'fStaff',1,10,10,150,900,8        
     
     DO adCheckBox WITH 'fStaff','checkTot','организация',.Shape1.Top+10,.Shape1.Left+5,150,dHeight,'dimOption(1)',0,.T.,'DO validCheckTot'          
     DO adCheckBox WITH 'fStaff','checkSov','совокупность',.checkTot.Top,.Shape1.Left,150,dHeight,'dimOption(2)',0,.T.,'DO validCheckgroup'
     DO adCheckBox WITH 'fStaff','checkPodr','подразделение',.checkTot.Top,.Shape1.Left,150,dHeight,'dimOption(3)',0,.T.,'DO validCheckpodr'
     .checkTot.Left=.Shape1.Left+(.Shape1.Width-.checkTot.Width-.checkPodr.Width-.checkSov.Width-40)/2
     .checkSov.Left=.checkTot.Left+.checkTot.Width+20
     .checkPodr.Left=.checkSov.Left+.checkSov.Width+20
     
     DO adLabMy WITH 'fStaff',1,'дата ',.checkTot.Top+.checkTot.Height+10,.Shape1.Left,.Shape1.Width,0,.T.,1 
     DO adTboxNew WITH 'fStaff','boxBeg',.checkTot.Top+.checkTot.Height+10,.Shape1.Left,RetTxtWidth('99/99/99999'),dHeight,'date_Beg',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3     
     
     .AddObject('grdPers','gridMynew')     
     WITH .grdPers
          .Top=.Parent.boxBeg.Top+.Parent.BoxBeg.Height+10     
          .Left=.Parent.Shape1.Left+10
          .Width=.Parent.Shape1.Width-20
         * .Height=RECCOUNT('katStaff')*.rowgridheights
         .Height=400
          .RecordSourceType=1
          .scrollBars=2   
          .ColumnCount=0         
          .colNesInf=2                        
        
          DO addColumnToGrid WITH 'fStaff.grdPers',12      
          .RecordSource='katStaff'                                         
          .Column1.ControlSource='katStaff.name'
          .Column2.ControlSource='katStaff.totpeop'                  
          .Column3.ControlSource='katStaff.totAlive'                  
          .Column4.ControlSource='katStaff.totDek'                  
          .Column5.ControlSource='katStaff.totVn'                  
          .Column6.ControlSource='katStaff.totWom'                  
          .Column7.ControlSource='katStaff.totWomWork'
          .Column8.ControlSource='katStaff.totkont'                
          .Column9.ControlSource='katStaff.kseshtat'                
          .Column10.ControlSource='katStaff.ksepeop'                
          .Column11.ControlSource='katStaff.ksevac'                
          
          .Column2.Width=RetTxtWidth('Wвн.совм')         
          .Column3.Width=.Column2.Width
          .Column4.Width=.Column2.Width
          .Column5.Width=.Column2.Width
          .Column6.Width=.Column2.Width
          .Column7.Width=.Column2.Width
          .Column8.Width=.Column2.Width
          .Column9.Width=.Column2.Width
          .Column10.Width=.Column2.Width
          .Column11.Width=.Column2.Width          
          .Columns(.ColumnCount).Width=0    
          .Column1.Width=.Width-.column2.Width*10-SYSMETRIC(5)-13-.ColumnCount   
          .Column1.Header1.Caption='персонал'
          .Column2.Header1.Caption='всего'
          .Column3.Header1.Caption='раб.'
          .Column4.Header1.Caption='декр.'
          .Column5.Header1.Caption='вн.сов.'
          .Column6.Header1.Caption='женщ.'
          .Column7.Header1.Caption='жен.раб.'          
          .Column8.Header1.Caption='контракт'
          .Column9.Header1.Caption='по штатн.'
          .Column10.Header1.Caption='занято'
          .Column11.Header1.Caption='вакант'
          
          .SetAll('Alignment',1,'ColumnMy')
          .Column1.Alignment=0    
          .Column2.Format='Z'
          .Column3.Format='Z'
          .Column4.Format='Z'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'
          .Column9.Format='Z'
          .Column10.Format='Z'          
          .Column11.Format='Z'          
          
          .SetAll('Enabled',.F.,'ColumnMy') 
          .Columns(.ColumnCount).Enabled=.T.          
     ENDWITH 
     DO gridSizeNew WITH 'fStaff','grdPers','shapeingrid',.T.,.F. 
     FOR i=1 TO .grdPers.columnCount 
         .grdPers.Columns(i).Backcolor=fStaff.BackColor           
         .grdPers.Columns(i).DynamicBackColor='IIF(RECNO(fStaff.grdPers.RecordSource)#fStaff.grdPers.curRec,fStaff.BackColor,dynBackColor)'
         .grdPers.Columns(i).DynamicForeColor='IIF(RECNO(fStaff.grdPers.RecordSource)#fStaff.grdPers.curRec,dForeColor,dynForeColor)'        
     ENDFOR 
          
     .Shape1.Height=.checkTot.Height+.boxBeg.Height+.grdPers.Height+40  
     .Shape1.Width=.grdPers.Width+20
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxBeg.Width-10)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+10     
     DO addButtonOne WITH 'fStaff','butCount',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wвозвратw')*3-20)/2,.Shape1.Top+.Shape1.Height+20,'расчёт','','DO countTotalStaff',39,RetTxtWidth('wвозвратw'),'расчёт' 
     DO addButtonOne WITH 'fStaff','butPrn',.butCount.Left+.butCount.Width+10,.butCount.Top,'печать','','DO formPrnStaff',.butCount.Height,.butCount.Width,'расчёт'  
     DO addButtonOne WITH 'fStaff','butRet',.butPrn.Left+.butPrn.Width+10,.butCount.Top,'возврат','','fStaff.Release',.butCount.Height,.butCount.Width,'вовзрат'   
     
     DO addListBoxMy WITH 'fStaff',1,.Shape1.Left,.Shape1.Top,.Shape1.Height,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH   
     
     DO addButtonOne WITH 'fStaff','butSave',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wприянтьw')*2-15)/2,.butCount.Top,'принять','','DO returnToStaff WITH .T.',39,RetTxtWidth('wпринятьw'),'принять'  
     DO addButtonOne WITH 'fStaff','butRetSave',.butSave.Left+.butSave.Width+15,.butSave.Top,'сброс','','DO returnToStaff WITH .F.',.butsave.Height,.butSave.Width,'сброс'      
     .butSave.Visible=.F.
     .butRetSave.Visible=.F.
     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.butCount.Height+60
ENDWITH
DO pasteImage WITH 'fStaff'
fStaff.Show
************************************************************************************************************
PROCEDURE validStruPrn
PARAMETERS par1
STORE .F. TO dimOption
dimOption(par1)=.T.
DO CASE
   CASE dimOption(1)=.T. 
   CASE dimOption(2)=.T. 
   CASE dimOption(3)=.T. 
ENDCASE
fStaff.Refresh
*************************************************************************************************************************
PROCEDURE validCheckTot
dimOption(1)=.T.
dimOption(2)=.F.
dimOption(3)=.F.
fStaff.checkPodr.Caption='подразделение'
fStaff.checkSov.Caption='совокупность'
fStaff.Refresh
*************************************************************************************************************************
PROCEDURE validCheckGroup 
dimOption(1)=.F.
dimOption(2)=.T. 
dimOption(3)=.F.
WITH fStaff
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')  
     .SetAll('Visible',.F.,'MyCommandButton')    
     .butSave.Visible=.T.
     .butRetSave.Visible=.T.
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopGroup.otm,name'  
     .listBox1.procForClick='DO clickListGroup'
     .listBox1.procForKeyPress='DO KeyPressListGroup' 
     .butSave.procForClick='DO returnToPrnGroup WITH .T.'
     .butRetSave.procForClick='DO returnToPrnGroup WITH .F.'
     .checkPodr.Caption='подразделение'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListGroup
SELECT dopGroup
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
nameGroup=IIF(fl,dopGroup.name,'')
GO rrec
fStaff.listBox1.SetFocus
GO rrec
*fStaff.listBox1.Refresh
*************************************************************************************************************************
PROCEDURE keyPressListGroup
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListGroip 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnGroup
PARAMETERS parRet
kvoGroup=0
IF parRet
   SELECT dopGroup
   LOCATE FOR fl
   IF FOUND()
      dimOption(2)=.T.
      dimOption(1)=.F.
      dimOption(3)=.F.
      nameGroup=dopGroup.name
      logbdj=dopgroup.logb
   ELSE 
      dimOption(1)=.T.
      dimOption(2)=.F.
      dimOption(3)=.F.
      nameGroup=''     
   ENDIF  
ELSE 
   nameGroup=''
   SELECT dopGroup
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
   dimOption(1)=.T.
   dimOption(2)=.F.
   dimOption(3)=.F.
   logBdj=.F.
ENDIF 
WITH fStaff
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel') 
     .SetAll('Visible',.T.,'MyCommandButton')    
     .butSave.Visible=.F.
     .butRetSave.Visible=.F.
     .listBox1.Visible=.F.  
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE validCheckPodr
dimOption(1)=.F. 
dimOption(2)=.F.
*dimOption(3)=.F.
WITH fStaff
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .butSave.Visible=.T.
     .butRetSave.Visible=.T.
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopPodr.otm,name'  
     .listBox1.procForClick='DO clickListPodr'
     .listBox1.procForKeyPress='DO KeyPressListPodr' 
     .butSave.procForClick='DO returnToPrnPodr WITH .T.'
     .butRetSave.procForClick='DO returnToPrnPodr WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListPodr
SELECT dopPodr
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fStaff.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListPodr
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListPodr 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnPodr
PARAMETERS parRet
kvoPodr=0
IF parRet
   SELECT dopPodr
   fltPodr=''
   onlyPodr=.F.
   SCAN ALL
        IF fl 
           fltPodr=fltPodr+','+LTRIM(STR(kod))+','
           onlyPodr=.T.
           kvoPodr=kvoPodr+1
        ENDIF 
   ENDSCAN
ELSE 
   strPodr=''
   onlyPodr=.F.
   SELECT dopPodr
   REPLACE otm WITH '',fl WITH .F. ALL
   dimOption(1)=.T.
   dimOption(2)=.F.
   dimOption(3)=.F.
   GO TOP
ENDIF 
WITH fStaff
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .butSave.Visible=.F.
     .butRetSave.Visible=.F.
     .listBox1.Visible=.F.  
      dimOption(3)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='подразделение'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH 
************************************************************************************************************
PROCEDURE countTotalStaff
PARAMETERS par1
IF USED('curStafjob')
   SELECT curStafjob
   USE 
ENDIF
IF USED('curStaff')
   SELECT curStaff
   USE 
ENDIF
SELECT * FROM datjob INTO CURSOR curStafjob READWRITE
SELECT curStafJob
APPEND FROM datjobout
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) ALL

*INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING
*INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2)+STR(kp,3) TAG T2 DESCENDING
INDEX ON STR(nidpeop,5)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING
INDEX ON STR(nidpeop,5)+STR(tr,1)+STR(kse,5,2)+STR(kp,3) TAG T2 DESCENDING
SET ORDER TO 1

SELECT * FROM curStafJob INTO CURSOR curShtatJob READWRITE
SELECT curStafJob
DELETE FOR !INLIST(tr,1,3)
DELETE FOR !EMPTY(dateOut).AND.dateOut<date_Beg
DELETE FOR dateBeg>date_beg
SELECT * FROM curStafJob INTO CURSOR curStafJob1 READWRITE
SELECT curStafJob1
*INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING
*INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2)+STR(kp,3) TAG T2 DESCENDING
INDEX ON STR(nidpeop,5)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING
INDEX ON STR(nidpeop,5)+STR(tr,1)+STR(kse,5,2)+STR(kp,3) TAG T2 DESCENDING

SELECT curShtatJob
REPLACE dekotp WITH IIF(SEEK(kodpeop,'people',1),people.dekotp,.F.) ALL
DELETE FOR tr=4
DELETE FOR !EMPTY(dateOut).AND.dateOut<date_Beg
*DELETE FOR !EMPTY(dateOut).AND.dateOut>date_Beg
DELETE FOR dateBeg>date_beg
DELETE FOR dekotp

SELECT curStafJob
DO CASE
   CASE dimOption(2)=.T.  &&группа печати
        fltch=ALLTRIM(dopGroup.sostav1)
        SET FILTER TO ','+LTRIM(STR(kp))+','$fltch  
        SELECT * FROM rasp WHERE ','+LTRIM(STR(kp))+','$fltch INTO CURSOR curRasp
        SELECT curShtatJob
        DELETE FOR !(','+LTRIM(STR(kp))+','$fltch)
   CASE dimOption(3)=.T.  &&подразделение
        fltch=fltpodr
        SET FILTER TO ','+LTRIM(STR(kp))+','$fltch  
        SELECT * FROM rasp WHERE ','+LTRIM(STR(kp))+','$fltch INTO CURSOR curRasp
        SELECT curShtatJob        
        DELETE FOR !(','+LTRIM(STR(kp))+','$fltch)       
   OTHERWISE      
        SELECT * FROM rasp INTO CURSOR curRasp
ENDCASE

IF dimOption(2)=.T.
   SELECT curStafJob1
   DELETE FOR SEEK(STR(nidpeop,5)+STR(tr,1)+STR(kse,5,2)+STR(kp,3),'curStafJob',2)  
   SELECT curStafJob  
   SCAN ALL
        IF SEEK(STR(nidpeop,5),'curStafJob1',1).AND.IIF(logbdj,curStafJob1.kse>curStafJob.kse,curStafJob1.kse>=curStafJob.kse)           
           DELETE 
        ENDIF       
        SELECT curStafJob
   ENDSCAN 
ENDIF 

SELECT * FROM people INTO CURSOR curStaff READWRITE
APPEND FROM peopout
ALTER TABLE curstaff ADD COLUMN kat N(1)
DELETE FOR !EMPTY(date_out).AND.date_Out<date_beg
DELETE FOR date_In>date_beg
REPLACE kat WITH IIF(SEEK(STR(nid,5),'curStafjob',1),curStafJob.kat,0) ALL
REPLACE lvn WITH IIF(SEEK(STR(nid,5)+STR(3,1),'curstafjob',1),.T.,lvn) ALL
DELETE FOR !SEEK(STR(nid,5),'curStafjob',1)
INDEX ON num TAG T1

COUNT TO tot_Peop FOR !lvn                                     && всего с декретчиками и без внешних совм.
COUNT TO tot_Alive FOR !lvn.AND.!dekOtp                        && "живых"
COUNT TO tot_Dek FOR dekotp.AND.!lvn.AND.bdekOtp<date_beg      && декретчики
COUNT TO tot_Vn FOR lvn                                        && внешние совм.

COUNT TO tot_Wom FOR !lvn.AND.sex=2                            && всего женщин
COUNT TO tot_WomWork FOR !lvn.AND.sex=2.AND.!dekOtp            && всего женщин "живых"
COUNT TO tot_kont FOR dog=1.AND.enddog>=date_beg
SELECT curRasp
SUM kse TO tot_kse

SELECT curShtatJob
SUM kse TO tot_ksepeop  
 
SELECT katStaff
LOCATE FOR kod=0
REPLACE totpeop WITH tot_peop,totalive WITH tot_alive,totdek WITH tot_dek,totvn WITH tot_vn,totwom WITH tot_wom,;
        totwomwork WITH tot_womwork,kseshtat WITH tot_kse,ksepeop WITH tot_ksepeop,ksevac WITH kseshtat-ksepeop,totkont WITH tot_kont
SKIP
SCAN WHILE !EOF()
     SELECT curStaff
     COUNT TO tot_Peop FOR !lvn.AND.kat=katStaff.kod                                     && всего с декретчиками и без внешних совм.
     COUNT TO tot_Alive FOR !lvn.AND.!dekOtp.AND.kat=katStaff.kod                        && "живых"
     COUNT TO tot_Dek FOR dekotp.AND.!lvn.AND.kat=katStaff.kod.AND.bDekOtp<date_beg      && декретчики
     COUNT TO tot_Vn FOR lvn.AND.kat=katStaff.kod                   					 && внешние совм.
     COUNT TO tot_Wom FOR !lvn.AND.sex=2.AND.kat=katStaff.kod                            && всего женщин
     COUNT TO tot_WomWork FOR !lvn.AND.sex=2.AND.!dekOtp.AND.kat=katStaff.kod            && всего женщин "живых"
     COUNT TO tot_kont FOR dog=1.AND.enddog>=date_beg.AND.kat=katStaff.kod               && с контрактами
     
     SELECT katStaff     
     REPLACE totpeop WITH tot_peop,totalive WITH tot_alive,totdek WITH tot_dek,totvn WITH tot_vn,totwom WITH tot_wom,totwomwork WITH tot_womwork,totkont WITH tot_kont
     
     SELECT curRasp
     SUM kse TO kse_cx FOR kat=katstaff.kod
     SELECT curShtatJob
     SUM kse TO ksepeop_cx FOR kat=katstaff.kod 
     
     SELECT katStaff
     REPLACE kseshtat WITH kse_cx,ksepeop WITH ksepeop_cx,ksevac WITH kseshtat-ksepeop
ENDSCAN
SELECT katStaff
GO TOP
fStaff.Refresh
******************************************************************************************
PROCEDURE formPrnStaff
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     DO adSetupPrnToForm WITH 10,10,400,.F.,.F.      
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnTotStaff WITH .T.' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'Просмотр','DO prnTotStaff WITH .F.'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','fSupl.Release','Выход из печати' 
      
                              
     .Width=.Shape91.Width+20
     .Height=.Shape91.Height+.cont1.Height+50
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*****************************************************************************************
PROCEDURE prnTotStaff
PARAMETERS parLog
SELECT katstaff
GO TOP
DO procForPrintAndPreview WITH 'reptotstaff','',parLog,.F.