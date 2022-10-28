fuvol=CREATEOBJECT('Formspr')
SELECT peopout
SET ORDER TO 2
GO TOP
log_ap=.F.
WITH fuvol
     .Icon='kone.ico'
     .Caption='Уволенные'  
     .ProcExit='DO exitFromProcuvol'  
      DO addcontico WITH 'fUvol','menucont1',10,5,'find.ico','DO findUvol','поиск',39,39      
      DO addcontico WITH 'fUvol','menucont2',.menucont1.Left+.menucont1.Width+10,5,'user.ico','DO cardUvol','л/карточка',39,39      
      DO addcontico WITH 'fUvol','menucont3',.menucont2.Left+.menucont2.Width+10,5,'script.ico','DO formsprav WITH .T.','справка',39,39      
      DO addcontico WITH 'fUvol','menucont4',.menucont3.Left+.menucont3.Width+10,5,'user_go.ico','DO formRestoreUser','восстановить',39,39      
      DO addcontico WITH 'fUvol','menucont5',.menucont4.Left+.menucont4.Width+10,5,'undo.ico','DO exitFromProcuvol','восстановить',39,39              
      WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Parent.menucont1.Height-5                           
          .RecordSourceType=1     
          .RecordSource='peopout'
          DO addColumnToGrid WITH 'fuvol.fGrid',4
          .Column1.ControlSource='peopout.num'
          .Column2.ControlSource='peopout.fio'
          .Column3.ControlSource='peopout.date_out'    
          .Column1.Width=RetTxtWidth(' 1234 ')
          .Column3.Width=RetTxtWidth('99/99/99999')
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0               
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='ФИО'
          .Column3.Header1.Caption='дата ув.'
          .Column1.Alignment=1
          .colNesInf=2      
          .Visible=.T.         
     ENDWITH
     DO gridSizeNew WITH 'fuvol','fGrid','shapeingrid'     
     DO addcontmy WITH 'fuvol','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fuvol','fuvol.cont1','peopout',1"
     DO addcontmy WITH 'fuvol','cont2',.cont1.Left+.fGrid.Column1.Width+2,.fGrid.Top+2,.fGrid.Column2.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fuvol','fuvol.cont2','peopout',2"
     .cont2.SpecialEffect=1   
     DO addcontmy WITH 'fuvol','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-4,.fGrid.HeaderHeight-3,''
     SELECT peopout
     GO TOP 
     .Show
ENDWITH
***************************************************************************
PROCEDURE carduvol
new_nfio=''
strSex=IIF(SEEK(peopout.sex,'cursex',1),cursex.name,'')
strDocum=IIF(SEEK(peopout.viddoc,'curdocum',1),curdocum.name,'')
newVidDoc=peopout.viddoc
new_dBirth=CTOD('  .  .    ')
IF !USED('sprtot')
   USE sprtot IN 0
ENDIF
IF !USED('curSrok')
   SELECT kod,name FROM sprtot WHERE sprtot.kspr=7 INTO CURSOR curSrok READWRITE && курсор для сроков заключения контракта
   SELECT curSrok
   INDEX ON kod TAG T1
ENDIF    
IF USED('suplEducation')
   SELECT suplEducation
   USE
ENDIF 
SELECT * FROM curEducation INTO CURSOR suplEducation READWRITE
SELECT suplEducation
INDEX ON kod TAG T1
strEduc=IIF(SEEK(peopout.educ,'suplEducation',1),suplEducation.name,'')

strSchool=peopout.school
strSpecd=peopout.specd
strKvald=peopout.kvald
newKodEduc=peopout.educ


strFamCard=IIF(SEEK(peopout.family,'curFamily',1),curFamily.name,'')
strVid=IIF(SEEK(peopout.dog,'sprdog',1),sprdog.name,'')
strSrok=peopout.strtime
SELECT peopout

fPersCard=CREATEOBJECT('FORMSUPL')
WITH fPersCard
     .Icon='kone.ico'
     .Caption='Личная карточка'
     .procExit='DO exitCardUvol'
     .Height=900
     .Width=1100    
     DO addPageFrame WITH 'fPersCard','pagePeop',3,0,0,.Width,.Height,.T.
     WITH .pagePeop
          .AddObject('mpage1','myPage')         
          WITH .mpage1
               .BackColor=RGB(255,255,255)
               nParent=.Parent
               oPage1=.Parent.mPage1
               .Caption='общие сведения'  
               .AddObject('formImage','myImage')               
               WITH .formImage     
                    .BorderStyle=1
                    .BackStyle=1
                    .Stretch=1                   
                    .Left=5             
                    .Top=5
                    .Width=100
                    .Height=150
                    .Visible=.T.
                    IF !EMPTY(peopout.pFoto)
                       pathmem='d:\datPic'+LTRIM(STR(peopout.num))+'.pic'
                       COPY MEMO peopout.pFoto TO &pathmem                            
                       .Picture=pathmem                       
                       widthTxt=widthTxtShort                   
                    ELSE 
                      .Picture='nofhoto.jpg'                                                           
                    ENDIF 
               ENDWITH  
               DO addShape WITH 'oPage1',1,.formImage.Left+.formImage.Width+10,.formImage.Top,100,nParent.Width-.formImage.Width-25,8                         
               
               DO adtBoxAsCont WITH 'oPage1','contFio',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('wнациональность'),dHeight,'ФИО',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFio',.contFio.Top,.contFio.Left+.contFio.Width-1,(.Shape1.Width-.contFio.Width*2-18)/2,dHeight,'peopout.fio',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contSex',.contFio.Left,.contFio.Top+.contFio.Height-1,.contFio.Width,dHeight,'пол',1,1 
               DO adTboxNew WITH 'oPage1','tBoxSex',.contSex.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'strsex',.F.,.F.,0   

                     
               DO adtBoxAsCont WITH 'oPage1','contTab',.contFio.Left,.contSex.Top+.contSex.Height-1,.contFio.Width,dHeight,'таб.номер',1,1 
               DO adTboxNew WITH 'oPage1','tBoxTab',.contTab.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'peopout.tabn',.F.,.F.,0   
               .tBoxTab.InputMask='99999'
               .tBoxtab.Alignment=0
               
               DO adtBoxAsCont WITH 'oPage1','contPnum',.contFio.Left,.contTab.Top+.contTab.Height-1,.contFio.Width,dHeight,'личный номер',1,1 
               DO adTboxNew WITH 'oPage1','tBoxPnum',.contPnum.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'peopout.pnum',.F.,.F.,0 
                              
               DO adtBoxAsCont WITH 'oPage1','contAge',.contFio.Left,.contPnum.Top+.contPnum.Height-1,.contFio.Width,dHeight,'дата рождения',1,1 
               DO adTboxNew WITH 'oPage1','tBoxAge',.contAge.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'peopout.Age',.F.,.F.,0 
                              
               DO adtBoxAsCont WITH 'oPage1','contPBirth',.tBoxFio.Left+.tBoxFio.Width-1,.contFio.Top,RetTxtWidth('место жительства'),dHeight,'место рождения',1,1 
               DO adTboxNew WITH 'oPage1','tBoxPBirth',.contPBirth.Top,.contPBirth.Left+.contPBirth.Width-1,.Shape1.Width-.contFio.Width-.tBoxFio.Width-.contPBirth.Width-18,dHeight,'peopout.placeborn',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contReg',.contPBirth.Left,.contPBirth.Top+.contPBirth.Height-1,.contPBirth.Width,dHeight,'зарегистрирован',1,1 
               DO adTboxNew WITH 'oPage1','tBoxReg',.contReg.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'peopout.pReg',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contLive',.contPBirth.Left,.contReg.Top+.contReg.Height-1,.contPBirth.Width,dHeight,'проживает',1,1 
               DO adTboxNew WITH 'oPage1','tBoxLive',.contLive.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'peopout.ppreb',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contPhone',.contPBirth.Left,.contLive.Top+.contLive.Height-1,.contPBirth.Width,dHeight,'телефон дом.',1,1 
               DO adTboxNew WITH 'oPage1','tBoxPhone',.contPhone.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'peopout.telhome',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contFree',.contPBirth.Left,.contPhone.Top+.contPhone.Height-1,.contPBirth.Width,dHeight,'телефон моб.',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFree',.contFree.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'peopout.telmob',.F.,.F.,0 
                                         
               
               .Shape1.Height=.contFio.Height*5+20
               *-------место рождения
               DO addShape WITH 'oPage1',2,.formImage.Left,.formImage.Top+.formImage.Height+5,100,(nParent.Width-25)/2,8 
           
               DO adtBoxAsCont WITH 'oPage1','contDoc',.Shape2.Left+10,.Shape2.Top+10,RetTxtWidth('место жительства'),dHeight,'документ',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDoc',.contDoc.Top,.contDoc.Left+.contDoc.Width-1,.Shape2.Width-.contDoc.Width-19,dHeight,'strDocum',.F.,.F.,0 
           
               DO adtBoxAsCont WITH 'oPage1','contDNum',.contDoc.Left,.contDoc.Top+.contDoc.Height-1,.contDoc.Width,dHeight,'номер',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDNum',.contDNum.Top,.tBoxDoc.Left,.tBoxDoc.Width,dHeight,'peopout.nDoc',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contDIn',.contDoc.Left,.contDnum.Top+.contDNum.Height-1,.contDoc.Width,dHeight,'дата выдачи',1,1
               DO adTboxNew WITH 'oPage1','tBoxDIn',.contDIn.Top,.tBoxDoc.Left,.tBoxDoc.Width,dHeight,'peopout.ddoc',.F.,.F.,0 
                
               DO adtBoxAsCont WITH 'oPage1','contDWho',.contDoc.Left,.contDIn.Top+.contDIn.Height-1,.contDoc.Width,dHeight,'кем выдан',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDWho',.contDWho.Top,.tBoxDoc.Left,.tBoxDoc.Width,dHeight,'peopout.vdoc',.F.,.F.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contDSrok',.contDoc.Left,.contDWho.Top+.contDWho.Height-1,.contDoc.Width,dHeight,'срок действия',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDStok',.contDSrok.Top,.tBoxDoc.Left,.tBoxDoc.Width,dHeight,'peopout.srokdoc',.F.,.F.,0 
                              
               .Shape2.Height=.contDoc.Height*5+20
               
               *-------остальное   
               DO addShape WITH 'oPage1',3,.Shape2.Left+.Shape2.Width+10,.Shape2.Top,.Shape2.Height,.Shape2.Width,8                         
                     
                                              
               DO adtBoxAsCont WITH 'oPage1','contEduc',.Shape3.Left+10,.Shape3.Top+10,RetTxtWidth('семейное положениеw'),dHeight,'образование',1,1 
               DO adTboxNew WITH 'oPage1','tBoxEduc',.contEduc.Top,.contEduc.Left+.contEduc.Width-1,.Shape3.Width-.contEduc.Width-18,dHeight,'strEduc',.F.,.F.,0 
               DO adtBoxAsCont WITH 'oPage1','contFam',.contEduc.Left,.contEduc.Top+.contEduc.Height-1,.contEduc.Width,dHeight,'семейное положение',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFam',.contFam.Top,.tBoxEduc.Left,.tBoxEduc.Width,dHeight,'strFamCard',.F.,.F.,0 
               DO adtBoxAsCont WITH 'oPage1','contMol',.contEduc.Left,.contFam.Top+.contFam.Height-1,.contEduc.Width,dHeight,'молодой специалист',1,1  
               oPage1.AddObject('boxMols','CONTAINER')
               WITH .boxMols
                    .BackColor=RGB(255,255,255)
                    .Visible=.T.
                    .Top=oPage1.contMol.Top
                    .Left=oPage1.tBoxEduc.Left
                    .Width=oPage1.tBoxEduc.Width
                    .Height=dHeight
                    .AddObject('check1','myCheckBox')
                    WITH .check1
                         .Caption=''
                         .Left=5
                         .Top=2
                         .Visible=.T.
                         .contRolSource='peopout.mols'
                         .BackStyle=0
                         .AutoSize=.T.
                    ENDWITH 
                    .AddObject('txtBox1','myTxtBox')                   
                    WITH .txtBox1
                         .Left=opage1.boxmols.check1.Left+opage1.boxMols.check1.Width+10
                         .Top=0
                          .Height=opage1.boxMols.Height
                         .Width=RetTxtWidth('99/99/99999')
                         .ControlSource='peopout.dmols'
                    ENDWITH
               ENDWITH
                 
               DO adtBoxAsCont WITH 'oPage1','contDek',.contEduc.Left,.contMol.Top+.contMol.Height-1,.contEduc.Width,dHeight,'декретный отпуск',1,1
               oPage1.AddObject('contBox','CONTAINER')
               WITH .contBox
                    .BackColor=RGB(255,255,255)
                    .Visible=.T.
                    .Top=oPage1.contDek.Top
                    .Left=oPage1.tBoxEduc.Left
                    .Width=oPage1.tBoxEduc.Width
                    .Height=dHeight
                    .AddObject('check1','myCheckBox')
                    WITH .check1
                         .Caption=''
                         .Left=5
                         .Top=2
                         .Visible=.T.
                         .contRolSource='peopout.dekOtp'
                         .BackStyle=0
                         .AutoSize=.T.  
                                                
                    ENDWITH  
                    .AddObject('txtBox1','myTxtBox')                   
                    WITH .txtBox1
                         .Left=opage1.contBox.check1.Left+opage1.contBox.check1.Width+10
                         .Top=0
                         .Height=opage1.contBox.Height
                         .Width=RetTxtWidth('99/99/99999')
                         .ControlSource='peopout.ddekotp'
                    ENDWITH                  
               ENDWITH
               DO adtBoxAsCont WITH 'oPage1','contFree1',.contEduc.Left,.contDek.Top+.contDek.Height-1,.contEduc.Width,dHeight,'',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFree1',.contFree1.Top,.tBoxEduc.Left,.tBoxEduc.Width,dHeight,'',.F.,.F.,0                          
               DO addShape WITH 'oPage1',4,.Shape2.Left,.Shape2.Top+.Shape2.Height+10,100,.Shape2.Width+.Shape3.Width+10,8 
               DO adCheckBox WITH 'oPage1','checkVn','внешний совместитель',.Shape4.Top+10,.Shape2.Left+10,150,dHeight,'peopout.lvn',0   
               DO adCheckBox WITH 'oPage1','checkUnion','член профсоюза',.checkVn.Top,.Shape2.Left+10,150,dHeight,'peopout.Union',0   
               DO adCheckBox WITH 'oPage1','checkPens','пенсионер',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'peopout.Pens',0   
               DO adCheckBox WITH 'oPage1','checkInv','инвалид',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'peopout.inv',0   
               DO adCheckBox WITH 'oPage1','checkMany','многодетный',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'peopout.mchild',0   
               DO adCheckBox WITH 'oPage1','checkAes','ЧАЭС',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'peopout.chaes',0   
               .Setall('Enabled',.F.,'myCheckBox')
               .checkVn.Left=.Shape4.Left+(.Shape4.Width-.checkVn.Width-.checkUnion.Width-.checkPens.Width-.checkInv.Width-.checkMany.Width-.checkAes.Width-50)/2
               .checkUnion.Left=.checkVn.Left+.checkVn.Width+10
               .checkPens.Left=.checkUnion.Left+.checkUnion.Width+10
               .checkInv.Left=.checkPens.Left+.checkPens.Width+10
               .checkMany.Left=.checkInv.Left+.checkInv.Width+10
               .checkAes.Left=.checkMany.Left+.checkMany.Width+10
               .Shape4.height=.checkUnion.Height+20                                         
               fPersCard.Height=.Shape1.Height+.Shape2.Height+.Shape4.Height+110          
               fPersCard.pagePeop.Height=fPersCard.Height
               fPersCard.Refresh               
               
          ENDWITH
          .Parent.Autocenter=.T.
          .AddObject('mpage2','myPage')      
          .AddObject('mpage3','myPage')
          .AddObject('mpage4','myPage')
          .AddObject('mpage5','myPage')
          .AddObject('mpage6','myPage')
          .AddObject('mpage7','myPage')
          .AddObject('mpage8','myPage')
          .AddObject('mpage9','myPage')
          .AddObject('mpage10','myPage')
          WITH .mpage2
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opage2=.Parent.mPage2
               .Caption='образование'  
               DO procEducationCard WITH .T.
               .SetAll('DisabledForeColor',RGB(1,0,0),'comboMy')
          ENDWITH
          WITH .mpage3     
               nParent=.Parent
               .Caption='состав семьи' 
               opage3=.Parent.mPage3 
               .BackColor=RGB(255,255,255)
               DO procFamilyCard WITH .T.                                                            
          ENDWITH
          WITH .mpage4
               nParent=.Parent
               .Caption='назначения и перемещ.'  
               opageSup=.Parent.mPage4
               .BackColor=RGB(255,255,255)
               SELECT * FROM datJobout WHERE datjobout.nidpeop=peopout.nid INTO CURSOR curHistory READWRITE
               SELECT curHistory
               INDEX ON dateBeg TAG T1 DESCENDING
               GO TOP
               .AddObject('grdJob','gridMyNew')     
               WITH .grdJob          
                    .ColumnCount=0
                    DO addColumnToGrid WITH 'oPageSup.grdJob',9
                    .Top=0
                    .Width=fPersCard.Width
                    .Height=fPersCard.Height
                    .Left=0
                    .ScrollBars=2      
                    .RecordSourceType=1
                    .RecordSource='curHistory'
                    .backColor=RGB(255,255,255)
                    .Column1.ControlSource="IIF(SEEK(curHistory.kp,'sprpodr',1),sprpodr.namework,'')"
                    .Column2.ControlSource="IIF(SEEK(curHistory.kd,'sprdolj',1),sprdolj.name,'')"
                    .Column3.ControlSource='curHistory.kse'
                    .Column4.ControlSource="IIF(SEEK(curHistory.tr,'sprtype',1),sprtype.name,'')"         
                    .Column5.ControlSource="IIF(!EMPTY(curHistory.dateBeg),curHistory.dateBeg,'')"
                    .Column6.ControlSource='curHistory.nordIn'
                    .Column7.ControlSource="IIF(!EMPTY(curHistory.dateout),curHistory.dateOut,'')"        
                    .Column8.ControlSource='curHistory.nordOut'
                    .Column3.Width=RetTxtWidth('999.999')   
                    .Column4.Width=RetTxtWidth('внеш.совмес.') 
                    .Column5.Width=RetTxtWidth('99/99/99999')                              
                    .Column6.Width=RetTxtWidth('приказw')
                    .Column7.Width=.Column5.Width 
                    .Column8.Width=.Column6.Width 
                    .Columns(.ColumnCount).Width=0
                    .Column2.Width=(.Width-.Column3.width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width)/2
                    .Column1.Width=.Width-.Column2.width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-SYSMETRIC(5)-13-.ColumnCount
                    .Column1.Header1.Caption='подразделение'
                    .Column2.Header1.Caption='должность'
                    .Column3.Header1.Caption='объём'
                    .Column4.Header1.Caption='тип'          
                    .Column5.Header1.Caption='принят'
                    .Column6.Header1.Caption='приказ'          
                    .Column7.Header1.Caption='уволен'
                    .Column8.Header1.Caption='приказ'
                    .Column1.Alignment=0
                    .Column2.Alignment=0
                    .Column4.Alignment=0    
                    .Column5.Alignment=0
                    .Column6.Alignment=0      
                    .SetAll('BOUND',.F.,'Column')  
                    .Visible=.T.   
               ENDWITH 
               * FOR i=1 TO .grdJob.columnCount        
               *     .grdJob.Columns(i).DynamicBackColor='IIF(RECNO(fJob.grdJob.RecordSource)#fJob.grdJob.curRec,fJob.BackColor,dynBackColor)'
               *     .grdJob.Columns(i).DynamicForeColor='IIF(RECNO(fJob.grdJob.RecordSource)#fJob.grdJob.curRec,dForeColor,dynForeColor)'        
               * ENDFOR                            
               DO gridSizeNew WITH 'oPageSup','grdJob','shapeingrid',.F.,.F.                  
          ENDWITH         
          WITH .mpage5
               nParent=.Parent
               opage5=.Parent.mPage5
               .BackColor=RGB(255,255,255)
               .Caption='контр.,катег.'  
               DO procKontrakt WITH .T.
          ENDWITH    
          WITH .mpage6
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageKurs=.Parent.mPage6
               .Caption='курсы'  
               DO procKursPCard WITH .T.
          ENDWITH 
          WITH .mpage7
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageBook=.Parent.mPage7
               .Caption='трудовая книжка'  
               DO procJobBook WITH .T.
          ENDWITH 
          WITH .mpage8
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opage8=.Parent.mPage8
               .Caption='воиский учёт'                 
               
               IF !USED('datArmy') 
                  USE datArmy ORDER 1 IN 0
               ENDIF 
               IF !USED('sprtot')
                   USE sprtot ORDER 1 IN 0
               ENDIF 
               SELECT kod,name,namesp FROM sprtot WHERE sprtot.kspr=10 INTO CURSOR curSupGrup READWRITE &&группы воинского учёта
               SELECT curSupGrup
               INDEX ON kod TAG T1               
               newUch=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.grupu,0)
               strGrup=IIF(SEEK(newUch,'curSupGrup',1),curSupGrup.name,'')
               newKatU=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.katu,0)
               newKzv=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.kzv,'')
               newzv=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.zv,'')
               strZv=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.zv,'')
               newRzp=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.rzp,0)
               newProfil=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.profil,'')
               newDateUch=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.dateUch,CTOD('  .  .    '))
               newDateSn=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.dateSn,CTOD('  .  .    '))
               newPrichSn=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.prichsn,'')
               newVus=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.vus,'')
               newRik=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.rik,'')
               strRik=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.rik,'')
               newUchet=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.uchet,'')

               newNTicket=IIF(SEEK(STR(peopout.num,4),'datarmy',1),datarmy.nTicket,'')
               SELECT rik FROM datarmy DISTINCT INTO CURSOR curSupRik READWRITE 
               SELECT curSupRik
               INDEX ON rik TAG T1
               SELECT peopout
               WITH oPage8     
                    DO adtBoxAsCont WITH 'oPage8','cont1',10,10,RetTxtWidth('WНаименование организацииW'),dHeight,'Группа учёта',0,1   
                    DO adTboxNew WITH 'oPage8','tBoxGr',.cont1.Top,.cont1.Left+.cont1.Width-1,nParent.Width-.cont1.Width-20,dHeight,'strGrup',.F.,.F.,0 
                    
                    DO adtBoxAsCont WITH 'oPage8','cont4',10,.cont1.Top+.cont1.Height-1,.cont1.Width,dHeight,'Воинское звание',0,1 
                    DO adTboxNew WITH 'oPage8','tBoxZv',.cont4.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'strZv',.F.,.F.,0 
                    
                    DO adtBoxAsCont WITH 'oPage8','cont11',10,.cont4.Top+.cont4.Height-1,.cont1.Width,dHeight,'Разряд запаса',0,1 
                    DO adTboxNew WITH 'oPage8','tBox11',.cont11.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newRzp','Z',.F.,0,'9' 
                    .tBox11.Alignment=0    
                    
                    DO adtBoxAsCont WITH 'oPage8','cont5',10,.cont11.Top+.cont11.Height-1,.cont1.Width,dHeight,'Военно-учётная спец-ть №',0,1 
                    DO adTboxNew WITH 'oPage8','tBox5',.cont5.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newVus',.F.,.F.,0 
     
                    DO adtBoxAsCont WITH 'oPage8','cont12',10,.cont5.Top+.cont5.Height-1,.cont1.Width,dHeight,'Профиль',0,1 
                    DO adTboxNew WITH 'oPage8','tBox12',.cont12.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newProfil',.F.,.F.,0
     
                    DO adtBoxAsCont WITH 'oPage8','cont2',10,.cont12.Top+.cont12.Height-1,.cont1.Width,dHeight,'Категория запаса',0,1 
                    DO adTboxNew WITH 'oPage8','tBox2',.cont2.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newKatU','Z',.F.,0,'99' 
                    .tBox2.Alignment=0
                    
                    DO adtBoxAsCont WITH 'oPage8','cont3',10,.cont2.Top+.cont2.Height-1,.cont1.Width,dHeight,'Дата приема на учет',0,1 
                    DO adTboxNew WITH 'oPage8','tBox3',.cont3.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newdateuch',.F.,.F.,0
     
                    DO adtBoxAsCont WITH 'oPage8','cont13',10,.cont3.Top+.cont3.Height-1,.cont1.Width,dHeight,'Дата снятия',0,1 
                    DO adTboxNew WITH 'oPage8','tBox13',.cont13.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newDatesn',.F.,.F.,0
     
                    DO adtBoxAsCont WITH 'oPage8','cont14',10,.cont13.Top+.cont13.Height-1,.cont1.Width,dHeight,'Основание снятия',0,1 
                    DO adTboxNew WITH 'oPage8','tBox14',.cont14.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newprichsn',.F.,.F.,0
     
                    DO adtBoxAsCont WITH 'oPage8','cont8',10,.cont14.Top+.cont14.Height-1,.cont1.Width,dHeight,'Cпецучёт №',0,1 
                    DO adTboxNew WITH 'oPage8','tBox8',.cont8.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newUchet',.F.,.F.,0 
     
                    DO adtBoxAsCont WITH 'oPage8','cont7',10,.cont8.Top+.cont8.Height-1,.cont1.Width,dHeight,'Комиссариат',0,1                   
                    DO adTboxNew WITH 'oPage8','tBoxRk',.cont7.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'strRik',.F.,.F.,0  
                   
                          
                    DO adtBoxAsCont WITH 'oPage8','cont10',10,.cont7.Top+.cont7.Height-1,.cont1.Width,dHeight,'№ военного билета',0,1 
                    DO adTboxNew WITH 'oPage8','tBox10',.cont10.Top,.tBoxGr.Left,.tBoxGr.Width,dHeight,'newNTicket',.F.,.F.,0      
              
                 
                   .Refresh            
               ENDWITH 
               
             
          ENDWITH     
          WITH .mpage9
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageBol=.Parent.mPage9
               .Caption='больничные листы'  
               DO procListBol WITH .T.
          ENDWITH 
          WITH .mpage10
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageOrd=.Parent.mPage10
               .Caption='приказы'  
               DO procOrderPeop WITH .T.
          ENDWITH       
     ENDWITH   
ENDWITH
fPersCard.Show
***************************************************************************
PROCEDURE finduvol
newKp=0
fPoisk=CREATEOBJECT('FORMSUPL')
WITH fPoisk
     .Icon='kone.ico'
     .Caption='Поиск'   
     DO addShape WITH 'fPoisk',1,10,10,dHeight,450,8     
     .logExit=.T.  
     find_ch=''
     DO adLabMy WITH 'fpoisk',1,'код или ФИО сотрудника' ,.Shape1.Top+10,.Shape1.Left+10,.Shape1.Width-20,2
     DO addtxtboxmy WITH 'fpoisk',1,.Shape1.Left+10,.Shape1.Top+.lab1.Height+10,.Shape1.Width-20,.F.,'find_ch'
     .Shape1.Height=.lab1.Height+.txtBox1.Height+30
     .txtBox1.procForkeyPress='DO keyPressFind'
     DO addContLabel WITH 'fpoisk','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wОтменаw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wОтменаw'),dHeight+3,'Поиск','DO searshUvol'
     DO addContLabel WITH 'fpoisk','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','Fpoisk.Release'          
     .Width=.Shape1.Width+20  
     .Height=.Shape1.Height+.cont1.Height+50 
ENDWITH     
DO pasteImage WITH 'fpoisk'
fpoisk.Show
*************************************************************************************************************************
*                Непосредственно поиск личной карточки
*************************************************************************************************************************
PROCEDURE SearshUvol
IF EMPTY(find_ch)
   RETURN
ENDIF
find_ch=ALLTRIM(find_ch)        
SELECT peopout
oldrec=RECNO()
log_ord=SYS(21)
IF TYPE(find_ch)='N' 
   SET ORDER TO 1
   IF SEEK(VAL(find_ch))
      SET ORDER TO &log_ord
      fPoisk.Release
   ELSE 
      find_ch=''
      fPoisk.Refresh
      SET ORDER TO &log_ord
      GO oldrec
      RETURN      
   ENDIF         
ELSE   
   SET ORDER TO 2
   DO unosimbol WITH 'find_ch',.F.,.F.           
   IF SEEK(find_ch)
      SET ORDER TO &log_ord
      fPoisk.Release
   ELSE 
      find_ch=''
      fPoisk.Refresh
      SET ORDER TO &log_ord
      GO oldrec
      RETURN   
   ENDIF     
ENDIF
*DO changeRowGrdPers
fUvol.fGrid.Columns(fUvol.fGrid.ColumnCount).SetFocus
************************************************************************************************************************
PROCEDURE keyPressFind
DO CASE
   CASE LASTKEY()=27
        fpoisk.Release
   CASE LASTKEY()=13
        find_ch=fpoisk.TxtBox1.Value           
        DO searshUvol  
ENDCASE 
***************************************************************************
PROCEDURE exitfromprocUvol
fUvol.Visible=.F.
SELECT peopout

***************************************************************************
PROCEDURE exitCardUvol
fPersCard.Visible=.F.
***************************************************************************
PROCEDURE formRestoreUser
fSupl=CREATEOBJECT('FORMMY')
logConfirm=.F.
WITH fSupl
     .Caption='Восстановление уволенного'
     .Width=520
     DO addShape WITH 'fSupl',1,10,10,100,500,8       
     DO adLabMy WITH 'fSupl',1,'Данная процедура переместить выбранного сотрудника из списка' ,.Shape1.Top+10,.Shape1.Left+10,.Shape1.Width-20,2                  
     DO adLabMy WITH 'fSupl',2,'уволенных в список работающих с восстановлением всех данных.' ,.lab1.Top+.lab1.Height,.Shape1.Left+10,.Shape1.Width-20,2                  
     .Shape1.Height=.lab1.Height*2+20
     DO adCheckBox WITH 'fSupl','checkConfirm','подтверждение намерений',.Shape1.Top+.Shape1.Height+10,.Shape1.Left+10,150,dHeight,'logConfirm',0    
     .checkConfirm.Left=.Shape1.Left+(.Shape1.Width-.checkConfirm.Width)/2
     DO addButtonOne WITH 'fSupl','butRestore',(.Width-RetTxtWidth('wвосстановитьw')*2-10)/2,.checkConfirm.Top+.checkConfirm.Height+10,'восстановить','','DO restoreUser',39,RetTxtWidth('wвосстановитьw'),'восстановить'
     DO addButtonOne WITH 'fSupl','butReturn',.butRestore.Left+.butRestore.Width+10,.butRestore.Top,'возврат','','fSupl.Release',39,.butRestore.Width,'возврат'               
     .Height=.butRestore.Top+.butRestore.Height+10
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************
PROCEDURE restoreUser
IF !logConfirm
   RETURN 
ENDIF 
IF USED('curRestore')
   SELECT curRestore
   USE
ENDIF 
SELECT peopout
oldnum=num
oldNidOut=nid
IF SEEK(peopout.num,'people',1)
   SELECT people
   oldOrd=SYS(21)
   oldRec=RECNO()
   SET ORDER TO 1
   GO BOTTOM
   numRestore=num+1
   SET ORDER TO &oldOrd
   GO oldRec
ELSE 
   numRestore=num
ENDIF 
SELECT people
APPEND FROM peopout FOR num=oldnum
REPLACE num WITH numRestore,date_out WITH CTOD('  .  .    '),nordout WITH '',dordout WITH CTOD('  .  .    ')

SELECT * FROM datjobout WHERE nidpeop=peopout.nid INTO CURSOR curRestore READWRITE 
SELECT curRestore
REPLACE kodpeop WITH numRestore,dateuv WITH CTOD('  .  .    ') ALL
REPLACE dateout WITH CTOD('  .  .    '), dordout WITH CTOD('  .  .    '), nordout WITH '', kordout WITH 0, nidout WITH 0 FOR dateout=peopout.date_out
SELECT datjob
APPEND FROM DBF('curRestore')
SELECT peopout
DELETE 
SELECT datjobout
DELETE FOR nidpeop=oldNidOut
SELECT peopout
fSupl.Release
fuvol.Refresh
