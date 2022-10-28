PARAMETERS parUv
IF !USED('datkurs')
    USE datkurs ORDER 5 IN 0
ELSE
   SELECT datkurs 
   SET ORDER TO 5    
ENDIF
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF 

SELECT kod,name FROM sprtot WHERE sprtot.kspr=22 INTO CURSOR kursType READWRITE   && виды подготовки
SELECT kursType
INDEX ON kod TAG T1

SELECT namekurs FROM datKurs INTO CURSOR curNameKurs DISTINCT ORDER BY namekurs
SELECT nameschool FROM datKurs INTO CURSOR curNameSchool DISTINCT ORDER BY nameschool
SELECT datKurs
*SET FILTER TO kodPeop=people.num
IF !parUv
   SET FILTER TO nidpeop=people.nid
ELSE
   SET FILTER TO nidpeop=peopout.nid
ENDIF  
GO TOP
nrec=0
logAp=.F.
new_kType=0
new_nType=''
new_perBeg=CTOD('  .  .    ')
new_perEnd=CTOD('  .  .    ')
new_nameKurs=''
new_nameSchool=''
new_ndoc=''
new_ddoc=CTOD('  .  .    ')
new_ndip=''
new_ddip=CTOD('  .  .    ')
WITH oPageKurs   
     .AddObject('fGrid','GridMy')       
     DO addButtonOne WITH 'oPageKurs','menuCont1',10,nParent.Height-dHeight-40,'новая','','DO readKursCard WITH .T.',39,RetTxtWidth('справочникw'),'новая'
     DO addButtonOne WITH 'oPageKurs','menuCont2',.menucont1.Left+.menucont1.Width+3,.menucont1.Top,'редакция','','DO readKursCard WITH .F.',39,.menucont1.Width,'редакция'
     DO addButtonOne WITH 'oPageKurs','menuCont3',.menucont2.Left+.menucont2.Width+3,.menucont1.Top,'удаление','','DO delKurs',39,.menucont1.Width,'удаление' 
     .SetAll('Enabled',IIF(!parUv,.T.,.F.),'myCommandButton')           
     WITH .fGrid       
          .Width=nParent.Width
          .Height=.Parent.menuCont1.Top-60
          .Left=0       
          .ScrollBars=2          
          .ColumnCount=9        
          .RecordSourceType=1     
          .RecordSource='datKurs'
          .Column1.ControlSource='datKurs.perBeg'
          .Column2.ControlSource='datKurs.perEnd'
          .Column3.ControlSource='datKurs.nameSchool'
          .Column4.ControlSource='datKurs.nameKurs'       
          .Column5.ControlSource='datKurs.khours'
          .Column6.ControlSource='datKurs.nOrd'
          .Column7.ControlSource='datKurs.dord'
          .Column8.ControlSource="IIF(SEEK(nType,'kursType',1),kursType.name,'')"
          
          .Column1.Width=RetTxtWidth('99/99/999999')
          .Column2.Width=RetTxtWidth('99/99/999999')               
          .Column5.Width=RetTxtWidth('часовw')               
          .Column6.Width=RetTxtWidth('9999-А')
          .Column7.Width=RetTxtWidth('99/99/999999')
          .Column8.Width=RetTxtWidth('wсдача квалификационногоw')
         
          .Column3.Width=(.Width-.column1.width-.Column2.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width)/2
          .Column4.Width=.Width-.column1.width-.Column2.Width-.column3.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-SYSMETRIC(5)-13-.ColumnCount       
          .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='начало'
          .Column2.Header1.Caption='окон.'
          .Column3.Header1.Caption='учреждение'
          .Column4.Header1.Caption='наименование курса'                
          .Column5.Header1.Caption='часов' 
          .Column6.Header1.Caption='пр №' 
          .Column7.Header1.Caption='дата пр.' 
          .Column8.Header1.Caption='вид обучения' 
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0           
          .Column5.Alignment=2                    
          .Column6.Alignment=0                    
          .Column7.Alignment=0  
          .Column8.Alignment=0  
          .Column5.Format='Z'                 
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')            
          .Visible=.T.     
     ENDWITH   
     DO gridSize WITH 'fPersCard.pagePeop.mPage6','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fPersCard.pagePeop.mPage6.fGrid.RecordSource)#fPersCard.pagePeop.mPage6.fGrid.curRec,fPersCard.pagePeop.mPage6.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fPersCard.pagePeop.mPage6.fGrid.RecordSource)#fPersCard.pagePeop.mPage6.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR           
     .SetAll('Visible',.F.,'MyTxtBox') 
     .setAll('Top',.fGrid.Top+.fGrid.height+20,'mymenuCont')    
      .setAll('Top',.fGrid.Top+.fGrid.height+20,'myCommandButton') 
     .menucont1.Left=(.fGrid.Width-.menucont1.Width-.menucont2.Width-.menucont3.Width-20)/2
     .menucont2.Left=.menucont1.Left+.menucont1.Width+10                   
     .menucont3.Left=.menucont2.Left+.menucont2.Width+10 
     
     *---------------------------------Кнопка удалить-------------------------------------------------------------------------
     DO addcontlabel WITH 'oPageKurs','butDel',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WудалитьW')*2-20)/2,.menucont1.Top,RetTxtWidth('WудалитьW'),.menucont1.Height,'удалить','DO delRecKurs WITH .T.' ,'удалить'
     *---------------------------------Кнопка возврат при удалении-------------------------------------------------------------------------                                            
     DO addcontlabel WITH 'oPageKurs','butDelRet',.butDel.Left+.butDel.Width+20,.butDel.Top,.butDel.Width,.butDel.Height,'возврат','DO delRecKurs WITH .F.'
    .butDel.Visible=.F.
    .butDelRet.Visible=.F.
     
             
ENDWITH
***********************************************************************************************************************
PROCEDURE readKursCard
PARAMETERS par1
IF datkurs.kodpeop=0.AND.!par1
   RETURN
ENDIF
SELECT datKurs
logAp=par1
new_perBeg=IIF(logAp,CTOD('  .  .    '),perBeg)
new_perEnd=IIF(logAp,CTOD('  .  .    '),perEnd)
new_nameKurs=IIF(logAp,'',nameKurs)
new_nameSchool=IIF(logAp,'',nameSchool)
new_nDoc=IIF(logAp,'',nord)
new_dDoc=IIF(logAp,CTOD('  .  .    '),dord)
new_nKhours=IIF(logAp,0,kHours)
new_kType=IIF(logAp,0,nType)
new_nType=IIF(SEEK(new_kType,'kursType',1),kursType.name,'')
formRead=CREATEOBJECT('FORMSUPL')
WITH formRead
     .Caption=IIF(logAp,'курсы-новая запись','курсы-редактирование')
     .procExit='DO exitFromReadKurs'
     DO addshape WITH 'formRead',1,10,10,150,400,8
     DO adTBoxAsCont WITH 'formRead','txtBeg',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('Wнаименование курсовW'),dHeight,'начало',2,1      
     DO adtboxnew WITH 'formRead','boxBeg',.txtBeg.Top,.txtBeg.Left+.txtBeg.Width-1,450,dheight,'new_perBeg',.F.,.T.  

     DO adTBoxAsCont WITH 'formRead','txtEnd',.txtBeg.Left,.txtBeg.Top+.txtBeg.Height-1,.txtBeg.Width,dHeight,'окончание',2,1 

     DO adtboxnew WITH 'formRead','boxEnd',.txtEnd.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_perEnd',.F.,.T.  

     DO adTBoxAsCont WITH 'formRead','txtName',.txtBeg.Left,.txtEnd.Top+.txtEnd.Height-1,.txtBeg.Width,dHeight,'наименование',2,1 
     DO addComboMy WITH 'formRead',1,.boxBeg.Left,.txtName.Top,dheight,.boxBeg.Width,.T.,'','curNameKurs.namekurs',6,.F.,'DO validNameKurs',.F.,.T. 
     .comboBox1.procForDropDown='DO dropDownNameKurs'     
     .comboBox1.procForLostFocus='DO lostFocusNameKurs' 
     .comboBox1.procForKeyPress='DO KeyPressNameKurs'
     DO adtboxnew WITH 'formRead','boxNaim',.txtName.Top,.boxBeg.Left,.boxBeg.Width-21,dheight,'new_nameKurs',.F.,.T.,0  
     .boxNaim.BackStyle=1          
 
     DO adTBoxAsCont WITH 'formRead','txtSchool',.txtBeg.Left,.txtName.Top+.txtName.Height-1,.txtBeg.Width,dHeight,'учреждение',2,1 
     DO addComboMy WITH 'formRead',2,.boxBeg.Left,.txtSchool.Top,dheight,.boxBeg.Width,.T.,'','curNameSchool.nameSchool',6,.F.,'DO validNameSchool',.F.,.T.  
     .comboBox2.procForDropDown='DO dropDownNameSchool'     
     .comboBox2.procForLostFocus='DO lostFocusNameSchool' 
     .comboBox2.procForKeyPress='DO KeyPressNameSchool'
     DO adtboxnew WITH 'formRead','boxSchool',.txtSchool.Top,.boxBeg.Left,.boxBeg.Width-21,dheight,'new_nameSchool',.F.,.T.     
     .boxSchool.BackStyle=1    
     
     DO adTBoxAsCont WITH 'formRead','txtHours',.txtBeg.Left,.txtSchool.Top+.txtSchool.Height-1,.txtBeg.Width,dHeight,'часов',2,1 
     DO adtboxnew WITH 'formRead','boxHours',.txtHours.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_nKhours','Z',.T.
     .boxHours.InputMask='999'
     .boxHours.Alignment=0
     
     DO adTBoxAsCont WITH 'formRead','txtNdoc',.txtBeg.Left,.txtHours.Top+.txtHours.Height-1,.txtBeg.Width,dHeight,'№ приказ',2,1 
     DO adtboxnew WITH 'formRead','boxNdoc',.txtNdoc.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_nDoc',.F.,.T. 
     
     DO adTBoxAsCont WITH 'formRead','txtDdoc',.txtBeg.Left,.txtNdoc.Top+.txtNdoc.Height-1,.txtBeg.Width,dHeight,'дата приказа',2,1 
     DO adtboxnew WITH 'formRead','boxDdoc',.txtDdoc.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_dDoc',.F.,.T. 
     
     DO adTBoxAsCont WITH 'formRead','txtType',.txtBeg.Left,.txtDdoc.Top+.txtDdoc.Height-1,.txtBeg.Width,dHeight,'вид обучения',2,1 
     DO addComboMy WITH 'formRead',3,.boxBeg.Left,.txtType.Top,dheight,.boxBeg.Width,.T.,'new_Ntype','kursType.name',6,.F.,'DO validKursType',.F.,.T. 
       
     .Shape1.Height=.txtBeg.Height*8+40
     .Shape1.Width=.txtBeg.Width+.boxBeg.Width+40
     
     *-----------------------------Кнопка записать---------------------------------------------------------------------------
     DO addcontlabel WITH 'formRead','cont1',.shape1.Left+(.shape1.Width-(RetTxtWidth('wзаписатьw'))*2-30)/2,;
        .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wзаписатьw'),dHeight+5,'записать','DO writeKurs'
          
     *---------------------------------Кнопка отказ --------------------------------------------------------------------------
     DO addcontlabel WITH 'formRead','cont2',.cont1.Left+.cont1.Width+30,.Cont1.Top,.Cont1.Width,dHeight+5,'отказ','DO exitFromReadKurs','отказ'  
     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+50
ENDWITH  
DO pasteImage WITH 'formRead'
formRead.Show
***********************************************************************************************************************
PROCEDURE validKursType
new_kType=kursType.kod
new_nType=kursType.name
KEYBOARD '{TAB}'   
***********************************************************************************************************************
PROCEDURE validNameKurs
new_nameKurs=curNameKurs.nameKurs
formRead.comboBox1.Width=formRead.boxBeg.Width
formRead.comboBox1.Left=formRead.boxBeg.Left
KEYBOARD '{TAB}'
formRead.Refresh
***********************************************************************************************************************
PROCEDURE dropDownNameKurs
formRead.comboBox1.Left=formRead.boxBeg.Left
*formRead.comboBox1.Width=objWidth
***********************************************************************************************************************
PROCEDURE lostFocusNameKurs
formRead.comboBox1.Width=formRead.boxBeg.Width
formRead.comboBox1.Left=formRead.boxBeg.Left
formRead.Refresh
***********************************************************************************************************************
PROCEDURE keyPressNameKurs
IF LASTKEY()=27
   WITH formRead.comboBox1
        .Width=.Parent.boxBeg.Width
        .Left=.Parent.boxBeg.Left 
   ENDWITH           
   KEYBOARD '{TAB}'   
ENDIF 
***********************************************************************************************************************
PROCEDURE validNameSchool
new_nameSchool=curNameSchool.nameSchool
formRead.comboBox2.Width=formRead.boxBeg.Width
formRead.comboBox2.Left=formRead.boxBeg.Left
KEYBOARD '{TAB}'
formRead.Refresh
***********************************************************************************************************************
PROCEDURE dropDownNameSchool
formRead.comboBox2.Left=formRead.boxBeg.Left
***********************************************************************************************************************
PROCEDURE lostFocusNameSchool
formRead.comboBox2.Width=formRead.boxBeg.Width
formRead.comboBox2.Left=formRead.boxBeg.Left
formRead.Refresh
***********************************************************************************************************************
PROCEDURE keyPressNameSchool
IF LASTKEY()=27
   WITH formRead.comboBox2
        .Width=.Parent.boxBeg.Width
        .Left=.Parent.boxBeg.Left 
   ENDWITH           
   KEYBOARD '{TAB}'   
ENDIF 
************************************************************************************************************************
PROCEDURE writeKurs
formRead.Release
SELECT datkurs
IF logAp
   APPEND BLANK
   REPLACE nidPeop WITH people.nid,kodPeop WITH people.num
ENDIF 
REPLACE perBeg WITH new_perBeg,perEnd WITH new_perEnd,nameKurs WITH new_nameKurs,nameSchool WITH new_nameSchool,;
        dord WITH new_ddoc,nord WITH new_ndoc,kHours WITH new_nKhours,nType WITH new_kType
oPageKurs.Refresh
************************************************************************************************************************
PROCEDURE delKurs
SELECT datKurs
IF kodpeop=0
   RETURN
ENDIF
WITH oPageKurs
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butDel.Visible=.T.
     .butDelRet.Visible=.T.
     .fGrid.Enabled=.F.
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE delRecKurs
PARAMETERS par1
SELECT datKurs
IF par1
   DELETE
ENDIF
WITH oPageKurs
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
ENDWITH 
************************************************************************************************************************
PROCEDURE exitFromReadKurs
formRead.Release