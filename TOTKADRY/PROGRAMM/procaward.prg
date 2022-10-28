IF !USED('dataward')
    USE dataward ORDER 1 IN 0
ELSE
   SELECT dataward
   SET ORDER TO 1
ENDIF

SELECT cnaward FROM dataward INTO CURSOR curNameaward DISTINCT ORDER BY cnaward
SELECT cobject FROM dataward INTO CURSOR curNameObject DISTINCT ORDER BY cobject
SELECT dataward
*SET FILTER TO kodPeop=people.num
SET FILTER TO nidpeop=people.nid
GO TOP
nrec=0
logAp=.F.
newKtype=0
new_nType=''
new_daward=CTOD('  .  .    ')

new_cnaward=''
new_cobject=''
new_nord=''
new_dord=CTOD('  .  .    ')
new_cprim=''

WITH oPageAward   
     .AddObject('fGrid','GridMy')       
     DO addButtonOne WITH 'oPageAward','menuCont1',10,nParent.Height-dHeight-40,'новая','','DO readAward WITH .T.',39,RetTxtWidth('справочникw'),'новая'
     DO addButtonOne WITH 'oPageAward','menuCont2',.menucont1.Left+.menucont1.Width+3,.menucont1.Top,'редакция','','DO readAward WITH .F.',39,.menucont1.Width,'редакция'
     DO addButtonOne WITH 'oPageAward','menuCont3',.menucont2.Left+.menucont2.Width+3,.menucont1.Top,'удаление','','DO delAward',39,.menucont1.Width,'удаление'          
     WITH .fGrid       
          .Width=nParent.Width
          .Height=.Parent.menuCont1.Top-60
          .Left=0       
          .ScrollBars=2          
          .ColumnCount=7        
          .RecordSourceType=1     
          .RecordSource='datAward'
          .Column1.ControlSource='datAward.daward'
          .Column2.ControlSource='datAward.cnaward'
          .Column3.ControlSource='datAward.cobject'
          .Column4.ControlSource='datAward.nOrd'
          .Column5.ControlSource='datAward.dord'
          .Column6.ControlSource='datAward.cprim'       
          
          .Column1.Width=RetTxtWidth('99/99/9999')
          .Column2.Width=RetTxtWidth('wwпочетная грамотаww')                     
          .Column4.Width=RetTxtWidth('9999-А')
          .Column5.Width=RetTxtWidth('99/99/9999')
         
          .Column3.Width=(.Width-.column1.width-.Column2.Width-.Column4.Width-.Column5.Width)/2
          .Column6.Width=.Width-.column1.width-.Column2.Width-.column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount       
          .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='дата'
          .Column2.Header1.Caption='награда'
          .Column3.Header1.Caption='от кого'
          .Column4.Header1.Caption='пр №' 
          .Column5.Header1.Caption='дата пр.' 
          .Column6.Header1.Caption='примечание'                
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0           
          .Column5.Alignment=0                    
          .Column6.Alignment=0                    
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')            
          .Visible=.T.     
     ENDWITH   
     DO gridSize WITH 'fPersCard.pagePeop.mPage11','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fPersCard.pagePeop.mPage11.fGrid.RecordSource)#fPersCard.pagePeop.mPage11.fGrid.curRec,fPersCard.pagePeop.mPage11.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fPersCard.pagePeop.mPage11.fGrid.RecordSource)#fPersCard.pagePeop.mPage11.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR           
     .SetAll('Visible',.F.,'MyTxtBox') 
     .setAll('Top',.fGrid.Top+.fGrid.height+20,'mymenuCont')    
      .setAll('Top',.fGrid.Top+.fGrid.height+20,'myCommandButton') 
     .menucont1.Left=(.fGrid.Width-.menucont1.Width-.menucont2.Width-.menucont3.Width-20)/2
     .menucont2.Left=.menucont1.Left+.menucont1.Width+10                   
     .menucont3.Left=.menucont2.Left+.menucont2.Width+10 
     
     *---------------------------------Кнопка удалить-------------------------------------------------------------------------
     DO addcontlabel WITH 'oPageAward','butDel',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WудалитьW')*2-20)/2,.menucont1.Top,RetTxtWidth('WудалитьW'),.menucont1.Height,'удалить','DO delRecAward WITH .T.' ,'удалить'
     *---------------------------------Кнопка возврат при удалении-------------------------------------------------------------------------                                            
     DO addcontlabel WITH 'oPageAward','butDelRet',.butDel.Left+.butDel.Width+20,.butDel.Top,.butDel.Width,.butDel.Height,'возврат','DO delRecAward WITH .F.'
    .butDel.Visible=.F.
    .butDelRet.Visible=.F.    
             
ENDWITH
***********************************************************************************************************************
PROCEDURE readAward
PARAMETERS par1
IF dataward.kodpeop=0.AND.!par1
   RETURN
ENDIF
SELECT datAward
logAp=par1
new_daward=IIF(logAp,CTOD('  .  .    '),daward)
new_cnaward=IIF(logAp,'',cnaward)
new_cobject=IIF(logAp,'',cobject)
new_nord=IIF(logAp,'',nord)
new_dord=IIF(logAp,CTOD('  .  .    '),dord)
new_cprim=IIF(logAp,'',cprim)

formRead=CREATEOBJECT('FORMSUPL')
WITH formRead
     .Caption=IIF(logAp,'награды-новая запись','награды-редактирование')
     .procExit='DO exitFromReadAward'
     DO addshape WITH 'formRead',1,10,10,150,400,8
     DO adTBoxAsCont WITH 'formRead','txtBeg',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('Wнаименование курсовW'),dHeight,'дата',1,1      
     DO adtboxnew WITH 'formRead','boxBeg',.txtBeg.Top,.txtBeg.Left+.txtBeg.Width-1,450,dheight,'new_daward',.F.,.T.      
   
     DO adTBoxAsCont WITH 'formRead','txtName',.txtBeg.Left,.txtBeg.Top+.txtBeg.Height-1,.txtBeg.Width,dHeight,'награда',1,1 
     DO addComboMy WITH 'formRead',1,.boxBeg.Left,.txtName.Top,dheight,.boxBeg.Width,.T.,'','curNameaward.cnaward',6,.F.,'DO validnaward',.F.,.T. 
     .comboBox1.procForDropDown='DO dropDownNaward'     
     .comboBox1.procForLostFocus='DO lostFocusnAward' 
     .comboBox1.procForKeyPress='DO KeyPressnAward'
     DO adtboxnew WITH 'formRead','boxNaim',.txtName.Top,.boxBeg.Left,.boxBeg.Width-21,dheight,'new_cnaward',.F.,.T.,0  
     .boxNaim.BackStyle=1          
 
     DO adTBoxAsCont WITH 'formRead','txtSchool',.txtBeg.Left,.txtName.Top+.txtName.Height-1,.txtBeg.Width,dHeight,'от кого',1,1 
     DO addComboMy WITH 'formRead',2,.boxBeg.Left,.txtSchool.Top,dheight,.boxBeg.Width,.T.,'','curNameobject.cobject',6,.F.,'DO validcobject',.F.,.T.  
     .comboBox2.procForDropDown='DO dropDowncobject'     
     .comboBox2.procForLostFocus='DO lostFocuscobject' 
     .comboBox2.procForKeyPress='DO KeyPresscobject'
     DO adtboxnew WITH 'formRead','boxSchool',.txtSchool.Top,.boxBeg.Left,.boxBeg.Width-21,dheight,'new_cobject',.F.,.T.     
     .boxSchool.BackStyle=1    
     
     DO adTBoxAsCont WITH 'formRead','txtNdoc',.txtBeg.Left,.txtSchool.Top+.txtSchool.Height-1,.txtBeg.Width,dHeight,'№ приказ',1,1 
     DO adtboxnew WITH 'formRead','boxNdoc',.txtNdoc.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_nord',.F.,.T. 
     
     DO adTBoxAsCont WITH 'formRead','txtDdoc',.txtBeg.Left,.txtNdoc.Top+.txtNdoc.Height-1,.txtBeg.Width,dHeight,'дата приказа',1,1 
     DO adtboxnew WITH 'formRead','boxDdoc',.txtDdoc.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_dord',.F.,.T. 
  
     DO adTBoxAsCont WITH 'formRead','txtprim',.txtBeg.Left,.txtddoc.Top+.txtddoc.Height-1,.txtBeg.Width,dHeight,'примечание',1,1 
     DO adtboxnew WITH 'formRead','boxprim',.txtprim.Top,.boxBeg.Left,.boxBeg.Width,dheight,'new_cprim',.F.,.T. 
     
       
     .Shape1.Height=.txtBeg.Height*7+40
     .Shape1.Width=.txtBeg.Width+.boxBeg.Width+40
     
     *-----------------------------Кнопка записать---------------------------------------------------------------------------
     DO addcontlabel WITH 'formRead','cont1',.shape1.Left+(.shape1.Width-(RetTxtWidth('wзаписатьw'))*2-30)/2,;
        .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wзаписатьw'),dHeight+5,'записать','DO writeAward'
          
     *---------------------------------Кнопка отказ --------------------------------------------------------------------------
     DO addcontlabel WITH 'formRead','cont2',.cont1.Left+.cont1.Width+30,.Cont1.Top,.Cont1.Width,dHeight+5,'отказ','DO exitFromReadAward','отказ'  
     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+50
ENDWITH  
DO pasteImage WITH 'formRead'
formRead.Show
***********************************************************************************************************************
PROCEDURE validNAward
new_cnaward=curNameAward.cnaward
formRead.comboBox1.Width=formRead.boxBeg.Width
formRead.comboBox1.Left=formRead.boxBeg.Left
KEYBOARD '{TAB}'
formRead.Refresh
***********************************************************************************************************************
PROCEDURE dropDownNaward
formRead.comboBox1.Left=formRead.boxBeg.Left
*formRead.comboBox1.Width=objWidth
***********************************************************************************************************************
PROCEDURE lostFocusNaward
formRead.comboBox1.Width=formRead.boxBeg.Width
formRead.comboBox1.Left=formRead.boxBeg.Left
formRead.Refresh
***********************************************************************************************************************
PROCEDURE keyPressNaward
IF LASTKEY()=27
   WITH formRead.comboBox1
        .Width=.Parent.boxBeg.Width
        .Left=.Parent.boxBeg.Left 
   ENDWITH           
   KEYBOARD '{TAB}'   
ENDIF 
***********************************************************************************************************************
PROCEDURE validCobject
new_cobject=curNameObject.cobject
formRead.comboBox2.Width=formRead.boxBeg.Width
formRead.comboBox2.Left=formRead.boxBeg.Left
KEYBOARD '{TAB}'
formRead.Refresh
***********************************************************************************************************************
PROCEDURE dropDowncObject
formRead.comboBox2.Left=formRead.boxBeg.Left
***********************************************************************************************************************
PROCEDURE lostFocuscObject
formRead.comboBox2.Width=formRead.boxBeg.Width
formRead.comboBox2.Left=formRead.boxBeg.Left
formRead.Refresh
***********************************************************************************************************************
PROCEDURE keyPresscObject
IF LASTKEY()=27
   WITH formRead.comboBox2
        .Width=.Parent.boxBeg.Width
        .Left=.Parent.boxBeg.Left 
   ENDWITH           
   KEYBOARD '{TAB}'   
ENDIF 
************************************************************************************************************************
PROCEDURE writeAward
formRead.Release
SELECT dataward
IF logAp
   APPEND BLANK
   REPLACE nidPeop WITH people.nid,kodPeop WITH people.num
ENDIF 
REPLACE daward WITH new_daward,cnaward WITH new_cnaward,cobject WITH new_cobject,;
        dord WITH new_dord,nord WITH new_nord,cprim WITH new_cprim
oPageAward.Refresh
************************************************************************************************************************
PROCEDURE delAward
SELECT datAward
IF kodpeop=0
   RETURN
ENDIF
WITH oPageAward
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butDel.Visible=.T.
     .butDelRet.Visible=.T.
     .fGrid.Enabled=.F.
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE delRecAward
PARAMETERS par1
SELECT datAward
IF par1
   DELETE
ENDIF
WITH oPageAward
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
PROCEDURE exitFromReadAward
formRead.Release