**************************************************************************************************************************************************
*
***************************************************************************************************************************************************
IF USED('curVacTarFond')
   SELECT curVacTarFond
   USE
ENDIF
newKval=0
vacRec=0
left_ch=0
SELECT * FROM sprkval INTO CURSOR curVacKval READWRITE
SELECT curVacKval
APPEND BLANK
INDEX ON kod TAG T1
SELECT * FROM sprkoef INTO CURSOR curVacKoef ORDER BY kod

SELECT num,rec,plrep,plrepvac FROM tarfond WHERE tarfond.vac INTO CURSOR curVacTarFond READWRITE 
ALTER TABLE curVacTarFond ADD COLUMN vacValue N(6,2)
SELECT curVacTarFond
INDEX ON num TAG T1

SELECT * FROM sprdolj INTO CURSOR curVacDolj
SELECT curVacDolj
INDEX ON kod TAG T1

SELECT * FROM rasp INTO CURSOR curVacRasp READWRITE
SELECT curVacRasp
REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL 
SELECT sprpodr
GO TOP 
DO WHILE !EOF()
   SELECT curVacRasp
   REPLACE np WITH sprpodr.np FOR kp=sprpodr.kod
   APPEND BLANK
   REPLACE named WITH sprpodr.name,kp WITH sprpodr.kod,log_s WITH .T.,np WITH sprpodr.np
   SELECT curVacRasp  
   SELECT sprpodr
   SKIP
ENDDO 
SELECT rasp
oldOrd=SYS(21)
SET ORDER TO 2
SELECT curVacRasp
INDEX ON STR(np,3)+STR(nd,3) TAG T1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET RELATION TO STR(kp,3)+STR(kd,3) INTO rasp ADDITIVE 
SET ORDER TO 1
GO TOP
filtervac_ch=''
fRasp=CREATEOBJECT('FORMMY')
WITH fRasp       
     DO addButtonOne WITH 'fRasp','menuCont1',10,5,'фильтр','filter1.ico','DO formFltVac',39,RetTxtWidth('wдополнительно')+44,'фильтр'  
     DO addButtonOne WITH 'fRasp','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'автозамена','replace.ico','Do formRepTarifVac' ,39,.menucont1.Width,'автозамена'       
     DO addButtonOne WITH 'fRasp','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'поиск','find.ico','Do formFindVac' ,39,.menucont1.Width,'поиск'      
     DO addButtonOne WITH 'fRasp','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'дополнительно','setup.ico','Do constvac' ,39,.menucont1.Width,'дополнительно'      
     DO addButtonOne WITH 'fRasp','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'возврат','undo.ico','Do exitFromSetupVac' ,39,.menucont1.Width,'возврат'      
     .Caption='Настройка для вакансий'
     .procexit='DO exitFromSetupVac'   
     .AddObject('fGrid','GridMy')
     WITH .fGrid
          .Top=fRasp.menucont1.Top+fRasp.menucont1.Height+5
          .Height=fRasp.Height-.Parent.menucont1.Height-5  
          .Width=fRasp.Width/2     
          .ColumnCount=3
          .RecordSourceType=1
          .RecordSource='curVacRasp'
          .ScrollBars=2
          .Column1.ControlSource='curVacRasp.nd'
          .Column2.ControlSource='curVacRasp.named' 
          .Column3.ControlSource='sprkval.name'
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Наименование'           
          .Column1.Width=RettxtWidth('999' )
          .Columns(.ColumnCount).Width=0
          SELECT curVacRasp
          .Columns(.ColumnCount).Width=0
          .Column2.Width=.Width-.column1.Width-SYSMETRIC(5)-13-.ColumnCount                                   
          .Column2.Alignment=0
          .Column3.Alignment=0
          .ProcAfterRowColChange='DO znVac'
          .colNesInf=2   
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')             
     ENDWITH    
     DO gridSize WITH 'fRasp','fGrid','shapeingrid' 
     
     FOR i=1 TO .fGrid.ColumnCount         
         .fGrid.Columns(i).fontname=dFontName
         .fGrid.Columns(i).fontSize=dFontSize      
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,dBackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,dForeColor,dynForeColor)'          
         .fGrid.Columns(i).DynamicBackColor='IIF(!curVacRasp.log_s,IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,dBackColor,dynBackColor),;
                                             IIF(RECNO(fRasp.fGrid.RecordSource)#fRasp.fGrid.curRec,headerBackColor,dynBackColor))'    
         .fGrid.Columns(i).Resizable=.F.           
         .fGrid.Columns(i).Text1.SelectedForeColor=dynForeColor
         .fGrid.Columns(i).Text1.SelectedBackColor=dynBackColor
         .fGrid.Columns(i).Header1.ForeColor=dForeColor
         .fGrid.Columns(i).Header1.Alignment=2
         .fGrid.Columns(i).Header1.FontName=dFontName
         .fGrid.Columns(i).Header1.FontSize=dFontSize
         .fGrid.Columns(i).Header1.BackColor=headerBackColor
     ENDFOR 
     strKval='' 
         
     DO adTBoxAsCont WITH 'fRasp','txtKat',.fGrid.Left+.fGrid.Width+10,.fGrid.Top,(.Width-.fGrid.Width-10)/4,dHeight,'категория',2,1
     DO adTBoxAsCont WITH 'fRasp','txtRaz',.txtKat.Left+.txtKat.Width-1,.txtKat.Top,.txtKat.Width,dHeight,'тарифный разряд',2,1 
     DO adTBoxAsCont WITH 'fRasp','txtKf',.txtRaz.Left+.txtRaz.Width-1,.txtKat.Top,.txtKat.Width,dHeight,'тарифный кфт.',2,1 
     DO adTBoxAsCont WITH 'fRasp','txtPkf',.txtKf.Left+.txtKf.Width-1,.txtKat.Top,.txtKat.Width+3,dHeight,'пов/пониж. кфт.',2,1 
     
     DO addComboMy WITH 'fRasp',1,.txtKat.Left,.txtKat.Top+.txtKat.Height-1,dHeight,.txtKat.Width,.T.,'strKval','curVacKval.name',6,.F.,'DO procKvalVac',.F.,.T. 
     DO adTboxNew WITH 'fRasp','boxRaz',.comboBox1.Top,.txtRaz.Left,.txtRaz.Width,dHeight,'curVacRasp.kfvac','Z',.T.,1,.F.,'DO validKoefVac',.F.,.F.,.F. 
     DO adTboxNew WITH 'fRasp','boxKf',.comboBox1.Top,.txtKf.Left,.txtKf.Width,dHeight,'curVacRasp.nkfvac','Z',.T.,1,.F.,'DO validNameKf',.F.,.F.,.F.
     DO adTboxNew WITH 'fRasp','boxPkf',.comboBox1.Top,.txtPkf.Left,.txtPkf.Width,dHeight,'curVacRasp.pkf','Z',.T.,1,.F.,'DO validPkfVac',.F.,.F.,.F.
     DO adTBoxAsCont WITH 'fRasp','txtNad',.txtKat.Left,.comboBox1.Top+.comboBox1.Height-1,.Width-.fGrid.Width-10,dHeight,'доплаты и надбавки',2,1       
     
     .AddObject('gridVac','GridMyNew')
     WITH .gridVac
          .Top=.Parent.txtNad.Top+.Parent.txtNad.Height-1
          .Left=.Parent.fGrid.Left+.Parent.fGrid.Width-1
          .Width=.Parent.txtNad.Width
          .Height=.Parent.fGrid.Height-.Parent.txtKf.height*4
          .ScrollBars=2
          .RecordSource='curVacTarFond'
          DO addColumnToGrid WITH 'fRasp.gridVac',3
          .Column1.ControlSource='curVacTarfond.rec'
          .Column2.ControlSource='curVacTarfond.vacValue'
         *.Column2.ControlSource=EVALUATE('rasp.'+ALLTRIM(curVacTarfond.plrep))
          .Column1.Enabled=.F.
          .Column1.Header1.Caption='Доплата/надбавка'
          .Column2.Header1.Caption='Значение'
          .Column2.Width=RetTxtWidth('wЗначение')
          .Column2.Format='Z'
          .Column3.Width=0
          .Column1.Width=.Width-.Column2.Width-SYSMETRIC(5)-13-.ColumnCount
          .setAll('Bound',.F.,'columnMy')
          .Column3.Enabled=.F.
          .Column2.Sparse=.T.          
     ENDWITH
     DO gridSizeNew WITH 'fRasp','gridVac','shapeingrid1',.T. 
     DO myColumnTxtBox WITH 'fRasp.gridVac.column2','txtbox2','curVacTarfond.vacValue',.F.,.F.,.F.,'DO validvacValue' 
     fRasp.gridVac.column2.txtbox2.Alignment=0
     
ENDWITH
SELECT curVacRasp
GO TOP
fRasp.Show
**************************************************************************************************************************
PROCEDURE procKvalVac
SELECT curVacRasp
strKval=curVacKval.name
SELECT rasp
*LOCATE FOR kp=curVacRasp.kp.AND.kd=curVacRasp.kd
REPLACE kv WITH curVacKval.kod 
SELECT curVacRasp
REPLACE kv WITH curVacKval.kod 
SELECT curVacTarFond
**************************************************************************************************************************************************
PROCEDURE validvacValue
SELECT curVacTarFond
repFiRasp=ALLTRIM(curVacTarFond.plRepvac)
SELECT curVacRasp
REPLACE &repfiRasp WITH curVacTarFond.vacValue
SELECT rasp
*LOCATE FOR kp=curVacRasp.kp.AND.kd=curVacRasp.kd
REPLACE &repfiRasp WITH curVacTarFond.vacValue
SELECT curVacTarFond
*************************************************************************************************************************************************
PROCEDURE znVac
strKval=IIF(SEEK(curVacRasp.kv,'curVacKval',1),ALLTRIM(curVacKval.name),'')
SELECT curVacTarFond
IF curVacRasp.log_s
   REPLACE vacValue WITH 0 ALL
ELSE    
   SCAN ALL
        repzn='curVacRasp.'+ALLTRIM(plrepvac)
        REPLACE vacValue WITH &repZn
   ENDSCAN    
ENDIF   
SELECT curVacTarFond
GO TOP
*fRasp.Refresh
WITH fRasp    
*     .comboBox1.Enabled=IIF(curVacRasp.log_s,.F.,.T.)
*     .boxRaz.Enabled=IIF(curVacRasp.log_s,.F.,.T.)
*     .boxKf.Enabled=IIF(curVacRasp.log_s,.F.,.T.)  
*     .gridVac.Enabled=IIF(curVacRasp.log_s,.F.,.T.)    
*     .gridVac.Column1.Enabled=.F.
     .gridVac.Column2.Enabled=IIF(curVacRasp.log_s,.F.,.T.) 
     .gridVac.Column3.Enabled=IIF(curVacRasp.log_s,.T.,.F.) 
*     .gridVac.Column2.ControlSource='curVacTarFond.vacValue' 
ENDWITH
fRasp.Refresh
**************************************************************************************************************************
PROCEDURE validKoefvac
SELECT curVacRasp
IF SEEK(curVacRasp.kfvac,'sprkoef',1)
   REPLACE nkfvac WITH sprkoef.name  
ENDIF
SELECT rasp
*LOCATE FOR kp=curVacRasp.kp.AND.kd=curVacRasp.kd
REPLACE kfvac WITH curVacrasp.kfvac,nkfvac WITH curVacRasp.nkfvac
SELECT curVacRasp
**************************************************************************************************************************
PROCEDURE validNameKf
SELECT curVacRasp
SELECT rasp
*LOCATE FOR kp=curVacRasp.kp.AND.kd=curVacRasp.kd
REPLACE nkfvac WITH curVacRasp.nkfvac
SELECT curVacRasp
**************************************************************************************************************************
PROCEDURE validPkfVac
SELECT curVacRasp
SELECT rasp
*LOCATE FOR kp=curVacRasp.kp.AND.kd=curVacRasp.kd
REPLACE pkf WITH curVacRasp.pkf
SELECT curVacRasp

**************************************************************************************************************************************************
PROCEDURE filtrVac
**************************************************************************************************************************************************
PROCEDURE constVac
=ACOPY(dimConstVac,dimConstOld)
fsupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
    * .procExit='DO saveDimConst'
     DO addshape WITH 'fSupl',1,20,20,150,380,8 
     DO adtBoxAsCont WITH 'fSupl','contDate',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('w'+dimConstVac(1,1)+'w'),dHeight,dimConstVac(1,1),2,1 
     DO addtxtboxmy WITH 'fSupl',1,.contDate.Left+.contDate.Width-1,.contDate.Top,RetTxtWidth('W99-99-99W'),.F.,'dimConstVac(1,2)',0,.F.     
     DO adtBoxAsCont WITH 'fSupl','contPers',.contDate.Left,.contDate.Top+.contDate.Height-1,.contDate.Width,dHeight,dimConstVac(2,1),2,1  
     DO addtxtboxmy WITH 'fSupl',2,.txtBox1.Left,.contPers.Top,.txtBox1.Width,.F.,'dimConstVac(2,2)',0,.F.   
     .txtBox1.Enabled=.F.
     .txtBox2.Enabled=.F.  
     .Shape1.Width=.contDate.Width+.txtBox1.Width+40
     .Shape1.Height=.contDate.Height*2+40
     .Width=.Shape1.Width+40  
     *-----------------------------Кнопка изменить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wизменитьw')*2)-20)/2,;
       .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wизменитьw'),dHeight+5,'изменить','DO procChangeConstVac'
     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','fSupl.Release','Возврат'
     
      *-----------------------------Кнопка сохранить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','contSave',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wсохранитьw')*2)-20)/2,;
       .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wсохранитьw'),dHeight+5,'сохранить','DO saveChangeConstVac WITH .T.'
     *---------------------------------Кнопка вовзрат --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','contRet',.contSave.Left+.contSave.Width+20,.ContSave.Top,.ContSave.Width,dHeight+5,'Возврат','DO saveChangeConstVac WITH .F.','Возврат'
     .contSave.Visible=.F.
     .contRet.Visible=.F.
     .Width=.Shape1.Width+40             
     
     .Height=.Shape1.Height+.cont1.Height+60
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************************************************************************************
PROCEDURE procChangeConstvac
WITH fSupl
     .txtBox1.Enabled=.T.
     .txtBox2.Enabled=.T.
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .contSave.Visible=.T.
     .contRet.Visible=.T.
ENDWITH 
***************************************************************************************************************************************************
PROCEDURE saveChangeConstvac
PARAMETERS par1
WITH fSupl
     .txtBox1.Enabled=.F.
     .txtBox2.Enabled=.F.
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .contSave.Visible=.F.
     .contRet.Visible=.F.
     IF par1
        fPath=FULLPATH('dimConstVac.mem')
        SAVE TO &fPath ALL LIKE dimConstVac ADDITIVE
     ELSE 
        =ACOPY(dimConstOld,dimConstVac)
        .txtBox1.ControlSource='dimConstVac(1,2)'
        .txtBox2.ControlSource='dimConstVac(2,2)'
        .Refresh
     ENDIF 
ENDWITH 
***************************************************************************************************************************************************
PROCEDURE saveDimConst
fPath=FULLPATH('dimConstVac.mem')
SAVE TO &fPath ALL LIKE dimConstVac ADDITIVE
***************************************************************************************************************************************************
PROCEDURE formFltVac
filtervac_ch=''
sostavfltvac=''
IF USED('curFltPodr')
   SELECT curFltPodr
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprpodr INTO CURSOR curFltPodr READWRITE
SELECT curFltPodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1
ord_podr=SYS(21)

IF USED('curFltDolj')
   SELECT curFltDolj
   USE
ENDIF 
SELECT kod,namework,fl,otm FROM sprdolj INTO CURSOR curFltDolj ORDER BY namework READWRITE
SELECT curFltDolj
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltKat')
   SELECT curFltKat
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprkat INTO CURSOR curFltkat ORDER BY name READWRITE
SELECT curFltKat
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltKval')
   SELECT curFltKval
   USE
ENDIF 
SELECT kod,name,fl,otm FROM sprkval INTO CURSOR curFltkval ORDER BY name READWRITE
SELECT curFltKval
REPLACE fl WITH .F.,otm WITH '' ALL

IF USED('curFltGrup')
   SELECT curFltGrup
   USE
ENDIF 
SELECT name,fl,sostav1 FROM datagrup INTO CURSOR curFltGrup READWRITE
SELECT curFltGrup
REPLACE fl WITH .F. ALL
ALTER TABLE curFltGrup ADD COLUMN otm C(1)

IF USED('curFltBase')
   SELECT curFltBase
   USE
ENDIF 
SELECT * FROM fltBase WHERE lvac INTO CURSOR curFltBase READWRITE
SELECT fltBase
ALTER TABLE curFltBase ADD COLUMN sostSup M
REPLACE strFlt WITH '',sayfl WITH '' ALL
rrec=0
frmFlt=CREATEOBJECT('FORMSUPL')
WITH frmFlt
     .Caption='Фильтр'
     .Icon='money.ico'   
     .procExit='DO exitFromFltvac'
     DO addListBoxMy WITH 'frmFlt',1,20,20,400,600  
     .AddObject('lstLine','LINE')
     WITH .listBox1
          .RowSource='curFltBase.namefL,sayFl'           
          .RowSourceType=6
          .ColumnCount=2
          
          .ColumnWidths='170,450' 
          .ColumnLines=.F.
          .procForValid='DO validListFlt'
          .procForKeyPress='DO KeyPressListFlt'       
     ENDWITH   
     .lstLine.Top=.listBox1.Top
     .lstLine.Height=.listBox1.Height
     .lstLine.Left=.ListBox1.Left+170+3
     .lstLine.Width=0    
     .lstLine.Visible=.T.      
    
     DO addcontlabel WITH 'frmFlt','cont1',.listBox1.Left+(.listBox1.Width-RetTxtWidth('WВыполнитьW')*3-20)/2,.listBox1.Top+.listBox1.Height+10,;
     RetTxtWidth('WВыполнитьW'),dHeight+3,'Выполнить','DO procFilterVac'
     DO addcontlabel WITH 'frmFlt','cont2',.Cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+3,'Сброс','DO filterVacRelease'    
     DO addcontlabel WITH 'frmFlt','cont3',.Cont2.Left+.Cont2.Width+10,.Cont1.Top, .Cont1.Width,dHeight+3,'Возврат','DO exitFromFltVac'  
     
     DO addcontlabel WITH 'frmFlt','contReturn',.listBox1.Left+(.listBox1.Width-RetTxtWidth('WВозвратW'))/2,.listBox1.Top+.listBox1.Height+10,;
     RetTxtWidth('WВозвратW'),dHeight+3,'Возврат','DO objFrmFltVisible WITH .F.'
     .contReturn.Visible=.F.       
          
     DO addListBoxMy WITH 'frmFlt',2,20,20,.listBox1.Height,.listBox1.Width    
     WITH .listbox2
          .Visible=.F.
          .ColumnCount=2
          .ColumnWidths='20,600'
          .RowSourceType=6
          .ControlSource=''
          .procForKeyPress='DO KeyPressListFlt2'
     ENDWITH   
     .AddObject('lstLine1','LINE')
     .lstLine1.Top=.listBox2.Top
     .lstLine1.Height=.listBox2.Height
     .lstLine1.Left=.ListBox2.Left+20+3
     .lstLine1.Width=0    
     .lstLine1.Visible=.F.
     .Width=.listBox1.Width+40
     .Height=.ListBox1.Height+.cont1.Height+50
     DO pasteImage WITH 'frmFlt'
     .Show    
     .Width=600
     .Height=600
ENDWITH 
**************************************************************************************************************************
PROCEDURE procFilterVac
frmFlt.Visible=.F.
SELECT curFltBase
filtervac_ch=''
GO TOP 
DO WHILE !EOF()
   IF !EMPTY(strFlt)
      filtervac_ch=filtervac_ch+ALLTRIM(namepl)+ALLTRIM(strFlt)+'.AND.'
   ENDIF    
   SKIP
ENDDO
frmFlt.Release
IF !EMPTY(filtervac_ch)   
   SELECT curvacrasp
   filtervac_ch=SUBSTR(filtervac_ch,1,LEN(filtervac_ch)-5)  
   SET FILTER TO EVALUATE(filtervac_ch)
   GO TOP   
   fRasp.Refresh          
ELSE
   SELECT curvacrasp
   SET FILTER TO 
   GO TOP   
   fRasp.Refresh        
ENDIF 
**************************************************************************************************************************
PROCEDURE filterVacRelease
SELECT curFltPodr
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltDolj
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltKat
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltKval
REPLACE fl WITH .F.,otm WITH '' ALL
*SELECT curFltKoef
*REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltGrup
REPLACE fl WITH .F.,otm WITH '' ALL
SELECT curFltBase
REPLACE strFlt WITH '',sayFl WITH '' ALL
GO TOP
SELECT curVacRasp
SET FILTER TO 
frmFlt.Refresh
***************************************************************************************************************************************************
PROCEDURE exitfromFltvac
frmFlt.Visible=.F.
frmFlt.Release
***************************************************************************************************************************************************
PROCEDURE formRepTarifVac
SELECT num,rec,plrepvac FROM tarfond WHERE !EMPTY(plrepvac) INTO CURSOR curRepTarif ORDER BY num
repVac_cx=''
newZnVac=0.00
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Автозамена тарифов'
     DO addShape WITH 'fSupl',1,10,10,400,50,8 
     DO adTBoxAsCont WITH 'fSupl','txtTarif',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('WWчто меняемWW'),dHeight,'что меняем',1,1
     DO addcombomy WITH 'fSupl',1,.txtTarif.Left+.txttarif.Width-1,.txtTarif.Top,dHeight,300,.T.,'','curRepTarif.rec',6,'','repVac_cx=curReptarif.plrepvac',.F.,.T.
     DO adTBoxAsCont WITH 'fSupl','txtZn',.txtTarif.Left,.txtTarif.Top+.txtTarif.Height-1,.txtTarif.Width,dHeight,'значение',1,1
     DO adTboxNew WITH 'fSupl','boxZn',.txtZn.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'newZnvac',.F.,.T.,0,.F.,'',.F.,.F.
     
     .Shape1.Height=.txtTarif.Height*2+20
     .Shape1.Width=.txtTarif.Width+.comboBox1.Width+20
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WзаменаW')*2-20)/2,.Shape1.Top+.Shape1.Height+10,RetTxtWidth('WзаменаW'),dHeight+5,'замена','DO reptarifVac','выполнить замену'
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'отказ','fSupl.Release','отказ'
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+30
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************************************************************************************
PROCEDURE repTarifVac
IF EMPTY(repVac_cx)
   RETURN 
ENDIF
SELECT curVacRasp
SCAN ALL
     REPLACE &repVac_cx WITH newZnVac
     IF ALLTRIM(LOWER(repVac_cx))='kfvac'  
        REPLACE nkfVac WITH IIF(SEEK(curVacRasp.kfVac,'sprkoef',1),sprkoef.name,0)
     ENDIF  
     SELECT rasp
     LOCATE FOR kp=curVacRasp.kp.AND.kd=curVacRasp.kd
     REPLACE &repVac_cx WITH newZnVac
     IF ALLTRIM(LOWER(repVac_cx))='kfvac'  
        REPLACE nkfVac WITH curVacRasp.nKfVac
     ENDIF  
     SELECT curVacRasp
ENDSCAN 
fSupl.Release
GO TOP 
fRasp.Refresh     
***************************************************************************************************************************************************
PROCEDURE formFindVac
fSupl=CREATEOBJECT('FORMSUPL')
SELECT curVacRasp
oldRecVac=RECNO()
strfPodr=''
fnPodr=0
=AFIELDS(arPodr,'sprpodr')
CREATE CURSOR curFindPodr FROM ARRAY arPodr 
SELECT curFindPodr
INDEX ON name TAG t1

WITH fSupl
     .Caption='поиск'
     .Icon='find.ico'
     .Width=400   
     .procExit='DO goPodrFind WITH .F.' 
     .procForClick='DO lostFocusPodrFind'          
     DO addShape WITH 'fSupl',1,10,10,60,250,8   
     DO adtBoxascont WITH 'fSupl','contPodr',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('WWнесколько последовательных символов наименования подразденленияWW'),dheight,'несколько последовательных символов наименования подразделения',2,1       
                 
     DO adtboxnew WITH 'fSupl','boxPodr',.contPodr.Top+.contPodr.Height-1,.contPodr.Left,.contPodr.Width-RetTxtWidth('w...')-1,dheight,'strfPodr',.F.,.T. 
     .boxPodr.procforChange='DO changePodrFind'  
     .boxpodr.procForRightClick='DO rightClickPodrFind'
 
     DO adtboxnew WITH 'fSupl','boxFree',.BoxPodr.Top,.boxPodr.Left+.boxPodr.Width-1,.contPodr.Width-.boxPodr.Width+1,dheight,'',.F.,.F.   
     DO addconticonew WITH 'fSupl','butPodr',.boxPodr.Left+.boxPodr.Width+1,.boxPodr.Top+2,'sbdn.ico',RetTxtWidth('w...')-1,.boxPodr.height-4,16,16,'DO selectPodrFind'    
                 
     .Shape1.Width=.contPodr.Width+40
     .Shape1.Height=.contPodr.Height*2+40            
                     
     .Width=.Shape1.Width+20
     *---------------------------------Кнопка поиск-----------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','contFind',(.Width-RetTxtWidth('Wперейтиw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,;
        RetTxtWidth('Wперейтиw'),dHeight+5,'перейти','DO goPodrFind WITH .T.' 
     *---------------------------------Кнопка возврат-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','contClear',.contFind.Left+.contFind.Width+10,.contFind.Top,.contFind.Width,dHeight+5,'возврат','DO goPodrFind WITH .F.'
     
     .Height=.Shape1.Height+.contFind.Height+50           
     
     DO addListBoxMy WITH 'fSupl',1,.boxPodr.Left,.boxPodr.Top+.boxPodr.Height-1,300,.contPodr.Width  
     WITH .listBox1
          .RowSource='curFindPodr.name'           
          .RowSourceType=2
          .ColumnCount=1         
          .Visible=.F.        
          .Height=.Parent.Height-.Parent.boxPodr.Top
          .procForValid='DO validListPodrFind'
          .procForLostFocus='DO lostFocusPodrFind'
          .procForKeyPress='DO KeyPressListPodrFind'        
     ENDWITH     
     .boxPodr.SetFocus
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show  
************************************************************************************************************************
PROCEDURE selectPodrFind
SELECT curFindPodr
ZAP
APPEND FROM sprpodr
WITH fSupl
    .listBox1.RowSource='curFindPodr.name'   
     SELECT curFindPodr
     LOCATE FOR name=fSupl.boxPodr.Text
     IF .listBox1.Visible=.F.
        .listBox1.Visible=.T.  
        .listBox1.SetFocus        
     ENDIF 
ENDWITH 
************************************************************************************************************************
PROCEDURE changePodrFind  
WITH fSupl 
     IF .listBox1.Visible=.F.
        .listBox1.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=fSupl.boxPodr.Text 
SELECT curFindPodr
ZAP
APPEND FROM sprpodr FOR LOWER(ALLTRIM(lcValue))$LOWER(name)
WITH fSupl.listBox1
     .RowSource='curFindPodr.name'        
ENDWITH 
************************************************************************************************************************
PROCEDURE rightClickPodrFind  

************************************************************************************************************************
PROCEDURE validListPodrFind
oldKodPodr=fnPodr
fnPodr=curFindPodr.kod
strfPodr=curFindPodr.name
fSupl.BoxPodr.ControlSource='strfPodr'
fSupl.listBox1.Visible=.F.
fSupl.BoxPodr.Refresh
*SELECT curVacRasp
************************************************************************************************************************          
PROCEDURE lostFocusPodrFind
WITH fSupl     
     .listBox1.Visible=.F.          
*     strfPodr=IIF(SEEK(fnPodr,'srpodr',1),sprpodr.name,'')  
*     .boxPodr.controlSource='strfpodr'
     .boxPodr.Refresh   
ENDWITH
************************************************************************************************************************
PROCEDURE KeyPressListPodrFind
IF LASTKEY()=27     
   fSupl.listBox1.Visible=.T.        
ENDIF   
************************************************************************************************************************
PROCEDURE validSfPodrFind
oldKodPodr=fnPodr
fnPodr=curFindPodr.kod
***********************************************************************************************************************************
PROCEDURE goPodrFind
PARAMETERS par1
fSupl.Visible=.F.
fSupl.Release
SELECT curVacRasp
IF par1   
   LOCATE FOR kp=fnPodr
   fRasp.Refresh
   fRasp.fGrid.Columns(fRasp.fGrid.ColumnCount).SetFocus
ELSE
   GO oldRecVac
ENDIF    
***************************************************************************************************************************************************
PROCEDURE exitFromSetupVac
SELECT rasp
SET ORDER TO &oldOrd
SELECT people
fRasp.Release