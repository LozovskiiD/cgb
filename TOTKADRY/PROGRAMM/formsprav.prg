**************************************************************************************************************************
*                                 Выписка справок и характеристики
**************************************************************************************************************************
PARAMETERS parUv
DIMENSION dimOpt(4)
STORE 0 TO dimOpt(4)
dimOpt(1)=1
IF USED('curjobsprav')
   SELECT curjobsprav
   USE
ENDIF 
fSupl=CREATEOBJECT('FORMSUPL')
IF !USED('boss')
   USE boss IN 0
ENDIF 
CREATE CURSOR curSprav (txtsprav M,txtChar M)
SELECT curSprav
APPEND BLANK
cAdr='по месту требования'
cNspr=''
cOrd=''
dDspr=DATE()
dSprsost=DATE()
cStajWork=''
cFio=''
cNpodr=''
cNdol=''
cStaj=''
cEduc=''
cDateIn=''
cNpodrcahr=''
cNdolchar=''
cAgeChar=''
DO textSprav WITH 1
IF !parUv
   SELECT people
ELSE 
   SELECT peopout
ENDIF    
WITH fSupl
     .Caption='Справка о месте работы и заниимаемой должности, характеристика с места работы'  
     DO addshape WITH 'fSupl',1,20,20,150,400,8 
        
     DO addOptionButton WITH 'fSupl',11,'о месте работы',.Shape1.Top+10,.Shape1.Left+20,'dimOpt(1)',0,'DO procValidSprav WITH 1',.T. 
     DO addOptionButton WITH 'fSupl',12,'о периоде работы',.Option11.Top,.Option11.Left+.Option11.Width+20,'dimOpt(2)',0,'DO procValidSprav WITH 2',.T. 
     
     DO addOptionButton WITH 'fSupl',13,'об отпуске по уходу',.Option11.Top,.Shape1.Left+20,'dimOpt(3)',0,'DO procValidSprav WITH 3',.T. 
     DO addOptionButton WITH 'fSupl',14,'характеристика',.Option11.Top,.Option11.Left+.Option11.Width+20,'dimOpt(4)',0,'DO procValidSprav WITH 4',.T. 
             
     .Shape1.Height=.Option11.Height+20
       
     DO addShape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+20,300,.Shape1.Width,8  
     DO adtBoxAsCont WITH 'fSupl','contDv',.Shape1.Left+10,.Shape2.Top+20,RetTxtWidth('WДата выдачиW'),dHeight,'дата выдачи',2,1         
     DO adTboxNew WITH 'fSupl','tBoxDv',.contDv.Top+.contDv.Height-1,.contDv.Left,.contDv.Width,dHeight,'dDspr',.F.,.T.,0     
     DO adtBoxAsCont WITH 'fSupl','contNum',.contDv.Left+.contDv.Width-1,.contDv.Top,RetTxtWidth('w№ справки'),dHeight,'№ спр-ки',2,1   
     DO adTboxNew WITH 'fSupl','tBoxNum',.tBoxDv.Top,.contNum.Left,.contNum.Width,dHeight,'cNspr',.F.,.T.,0    
     
        
     DO adtBoxAsCont WITH 'fSupl','cont1',.contNum.Left+.contNum.Width-1,.contDv.Top,RetTxtWidth('WДата рожленияW'),dHeight,'дата рождения',2,1         
     DO adTboxNew WITH 'fSupl','tBox1',.tBoxDv.Top,.cont1.Left,.cont1.Width,dHeight,'cAgeChar',.F.,.T.,0     
     DO adtBoxAsCont WITH 'fSupl','cont11',.cont1.Left+.cont1.Width-1,.cont1.Top,RetTxtWidth('w99 лет 99 месяцев 99 деньw'),dHeight,'стаж',2,1   
     DO adTboxNew WITH 'fSupl','tBox11',.tBox1.Top,.cont11.Left,.cont11.Width,dHeight,'cStaj',.F.,.T.,0     
      
      DO adtBoxAsCont WITH 'fSupl','cont2',.cont11.Left+.cont11.Width-1,.cont1.Top,RetTxtWidth('wдля предоставления по месту требования и w'),dHeight,'адресат',2,1   
      DO adTboxNew WITH 'fSupl','tBox2',.tBox1.Top,.cont2.Left,.cont2.Width,dHeight,'cAdr',.F.,.T.,0     
      DO adtBoxAsCont WITH 'fSupl','cont3',.contDv.Left,.tBox1.Top+.tBox1.Height-1,.contDv.Width+.contNum.Width+.cont1.Width+.cont11.Width+.cont2.Width-4,dHeight,'кем работает',2,1   
      .AddObject('editSprav','MyEditBox')      
      WITH .editSprav
          .Visible=.T.          
          .ControlSource='curSprav.txtSprav'
          .Left=.Parent.contDv.Left
          .Width=.Parent.cont3.Width
          .Top=.Parent.cont3.Top+.Parent.cont3.Height-1
          .Height=dHeight*2
          .Enabled=.T.  
     ENDWITH
     
     
     DO adtBoxAsCont WITH 'fSupl','cont31',.contDv.Left,.editSprav.Top+.editSprav.Height-1,.cont3.Width,dHeight,'образование и проч.',2,1   
      .AddObject('editChar','MyEditBox')      
      WITH .editChar
           .Visible=.T.          
           .ControlSource='curSprav.txtChar'
           .Left=.Parent.cont31.Left
           .Width=.Parent.cont3.Width
           .Top=.Parent.cont31.Top+.Parent.cont31.Height-1
           .Height=dHeight*2
           .Enabled=.T.  
     ENDWITH
     
     
     DO adtBoxAsCont WITH 'fSupl','cont4',.contDv.Left,.editChar.Top+.editChar.Height-1,RetTxtWidth('WДата выдачиW'),dHeight,'состояние на',2,1         
     DO adTboxNew WITH 'fSupl','tBox4',.cont4.Top+.cont4.Height-1,.cont4.Left,.cont4.Width,dHeight,'dSprsost',.F.,.T.,0     
     DO adtBoxAsCont WITH 'fSupl','cont5',.cont4.Left+.cont4.Width-1,.cont4.Top,RetTxtWidth('Wподпись должностьW'),dHeight,'подписть должность',2,1         
     DO adTboxNew WITH 'fSupl','tBox5',.tbox4.Top,.cont5.Left,.cont5.Width,dHeight,'boss.cspravdol',.F.,.T.,0  
     DO adtBoxAsCont WITH 'fSupl','cont6',.cont5.Left+.cont5.Width-1,.cont4.Top,.cont3.Width-.cont4.Width-.cont5.Width+2,dHeight,'подписть ФИО',2,1         
     DO adTboxNew WITH 'fSupl','tBox6',.tbox4.Top,.cont6.Left,.cont6.Width,dHeight,'boss.cspravfio',.F.,.T.,0      
     .Shape2.Width=.cont3.Width+20     
     .Shape2.Height=.cont1.Height*4+.tBox1.Height*2+.editSprav.Height+.editChar.Height+40
     .Shape1.Width=.Shape2.Width
     .Option11.Left=.Shape1.Left+(.Shape1.Width-.Option11.Width-.Option12.Width-.Option13.Width-.Option14.Width-30)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10   
     .Option13.Left=.Option12.Left+.Option12.Width+10   
     .Option14.Left=.Option13.Left+.Option13.Width+20   
     .Width=.Shape1.Width+40      
     
     DO addButtonOne WITH 'fSupl','butPrn',.Shape2.Left+(.Shape2.Width-RetTxtWidth('WПросмотрW')*2-20)/2,.Shape2.Top+.Shape2.Height+20,'печать','','DO prnsprav',39,RetTxtWidth('WПросмотрW'),'печать'
     DO addButtonOne WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+20,.butPrn.Top,'возврат','','fSupl.Release',39,.butPrn.Width,'возврат' 
        
     
     .Height=.Shape1.Height+.Shape2.Height+.butPrn.Height*2+80 
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE procValidSprav
PARAMETERS par1
STORE 0 TO dimOpt
dimOpt(par1)=1
DO textSprav WITH par1
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE textSprav
PARAMETERS parDoc
cFio=IIF(!parUv,ALLTRIM(people.fio),ALLTRIM(peopout.fio))
llvn=IIF(!parUv,people.lvn,peopout.lvn)
IF !parUv
   SELECT * FROM datjob WHERE nidpeop=people.nid.AND.EMPTY(dateOut) INTO CURSOR curjobsprav 
ELSE 
   SELECT * FROM datjobout WHERE nidpeop=peopout.nid INTO CURSOR curjobsprav
ENDIF    
SELECT curjobsprav
IF !llvn
   IF !paruv
      LOCATE FOR tr=1.AND.EMPTY(dateOut)
      IF !FOUND()
         GO TOP 
      ENDIF 
   ELSE 
      LOCATE FOR tr=1.AND.dateout=peopout.date_out           
   ENDIF 
ELSE
   LOCATE FOR tr=3
ENDIF

DO CASE
   CASE parDoc=1   && с места работы      
        cNpodr=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
        cNdol=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namet),ALLTRIM(sprdolj.namet),ALLTRIM(sprdolj.name)),'')  
        
        IF !parUv         
           cOrd=IIF(SEEK(people.nid,'datjob',8).AND.!EMPTY(datjob.dOrdin).AND.datjob.date_in=people.date_in, ' (приказ № '+ALLTRIM(datjob.nOrdin)+' от '+DTOC(datjob.dOrdIn)+'г.)','') 
           IF EMPTY(cOrd)
             * cOrd=IIF(SEEK(STR(people.nid,5),'peoporder',3), ' (приказ № '+ALLTRIM(peoporder.nOrd)+' от '+DTOC(peoporder.dOrd)+'г.)','')                  
           ENDIF 
        ELSE          
           cOrd=IIF(SEEK(peopout.nid,'datjobout',8).AND.!EMPTY(datjobout.dOrdin).AND.datjobout.date_in=peoplout.date_in, ' (приказ № '+ALLTRIM(datjobout.nOrdin)+' от '+DTOC(datjobout.dOrdIn)+'г.)','') 
           IF EMPTY(cOrd)
              *cOrd=IIF(SEEK(STR(peopout.nid,5),'peoporder',3), '(приказ № '+ALLTRIM(peoporder.nOrd)+' от '+DTOC(peoporder.dOrd)+'г.)','')      
           ENDIF         
        ENDIF    
                    
        SELECT curSprav
        REPLACE txtsprav WITH LOWER(cNdol)+' '+LOWER(cNpodr)+' c '+IIF(!EMPTY(IIF(!paruv,people.date_in,peopout.date_in)),LTRIM(STR(DAY(IIF(!paruv,people.date_in,peopout.date_in))))+' '+;
                ALLTRIM(month_prn(MONTH(IIF(!paruv,people.date_in,peopout.date_in))))+' '+STR(YEAR(IIF(!paruv,people.date_in,peopout.date_in)),4)+' года','')+cOrd+' по настоящее время.'
        REPLACE txtChar WITH ''
   CASE parDoc=2   &&  о периоде работы&&  о периоде работы
        cNpodr=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
        cNdol=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namet),ALLTRIM(sprdolj.namet),ALLTRIM(sprdolj.name)),'')  
        dordin_cx=CTOD('  .  .    ')
        nordin_cx=''
        
        dordout_cx=CTOD('  .  .    ')
        nordout_cx=''
        ON ERROR DO erSup
        IF !parUv    
           cDateIn=people.date_in
           IF!EMPTY(people.dordin)
             dordin_cx=people.dordin
             nordin_cx=people.nordin
           ELSE   
              dordin_cx=IIF(SEEK(people.nid,'datjob',8).AND.!EMPTY(datjob.dOrdin).AND.datjob.date_in=people.date_in,datjob.dOrdIn,dordin_cx) 
              nordin_cx=ALLTRIM(datjob.nOrdin)  
           ENDIF 
           IF EMPTY(dordin_cx)
             * dordin_cx=IIF(SEEK(STR(people.nid,5),'peoporder',3),peoporder.dOrd,'')
             * nordin_cx=ALLTRIM(peoporder.nOrd)  
           ENDIF 
        ELSE 
           cDateIn=peopout.date_in 
           IF !EMPTY(peopout.dordin)
              dordin_cx=peopout.dordin
              nordin_cx=ALLTRIM(peopout.nordin)
           ELSE   
              dordin_cx=IIF(SEEK(peopout.nid,'datjobout',8).AND.!EMPTY(datjobout.dOrdin).AND.datjobout.date_in=peopout.date_in,datjobout.dOrdIn,dordin_cx) 
              nordin_cx=ALLTRIM(datjobout.nOrdin)           
           ENDIF 
           IF EMPTY(dordin_cx)
              *dordin_cx=IIF(SEEK(STR(peopout.nid,5),'peoporder',3),peoporder.dOrd,'')
              *nordin_cx=ALLTRIM(peoporder.nOrd)  
           ENDIF
           *dordout_cx=peopout.dordout
           *nordout_cx=ALLTRIM(peopout.nordout)                    
        ENDIF    
        cDateIn='с '+LTRIM(STR(DAY(cDateIn)))+' '+ALLTRIM(month_prn(MONTH(cDateIn)))+' '+STR(YEAR(cDateIn),4)+' года'              
        SELECT curSprav
        REPLACE txtsprav WITH 'был(а) принят(а) '+LOWER(cNdol)+' '+LOWER(cNpodr)+' в учреждение здравоохранения «Брестская центральная городская больница»'        
        REPLACE txtchar WITH 'приказом от '+IIF(!EMPTY(dordin_cx),'«'+LTRIM(STR(DAY(dordin_cx)))+'» '+month_prn(MONTH(dordin_cx))+' '+STR(YEAR(dordin_cx),4)+'г. ','')+'№ '+nOrdin_cx               
        IF parUv
           REPLACE txtchar WITH txtchar+' и '+'«'+LTRIM(STR(DAY(peopout.date_out)))+'» '+ALLTRIM(month_prn(MONTH(peopout.date_out)))+' '+STR(YEAR(peopout.date_out),4)+'г. уволен(а) приказом № '+ALLTRIM(peopout.nordout)  
           REPLACE txtchar WITH txtchar+' от '+'«'+LTRIM(STR(DAY(peopout.dordout)))+'» '+ALLTRIM(month_prn(MONTH(peopout.dordout)))+' '+STR(YEAR(peopout.dordout),4)+'г.'
        ENDIF        
        ON ERROR
   CASE parDoc=3   &&  об декретном отпуске
   CASE parDoc=4   &&  характеристика
        IF !parUv           
           cNpodr=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
           cNdol=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namework),ALLTRIM(sprdolj.namework),ALLTRIM(sprdolj.namework)),'')
           cAgeChar=IIF(!EMPTY(people.age),DTOC(people.age)+'.','')
           DO actualStajToday WITH 'people','people.date_in','DATE()','cStajWork',.T.
           yspr=''
           mspr=''
           dspr=''
           yspr=LTRIM(LEFT(cStajWork,2))
           mspr=ALLTRIM(SUBSTR(cStajWork,4,2))
           dspr=LTRIM(RIGHT(cStajWork,2))           
           ON ERROR DO erSup
           IF !EMPTY(VAL(yspr))      
              yspr=LTRIM(STR(VAL(yspr)))
              DO CASE 
                 CASE BETWEEN(VAL(yspr),5,20)
                      yspr=yspr+' лет'                   
                 CASE RIGHT(yspr,1)='1'
                      yspr=yspr+' год'                    
                 CASE BETWEEN(VAL(RIGHT(yspr,1)),2,4)
                      yspr=yspr+' года'     
                 OTHERWISE 
                      yspr=yspr+' лет'     
              ENDCASE 
           ELSE
              yspr=''           
           ENDIF     
           IF !EMPTY(VAL(mspr))      
              DO CASE 
                 CASE BETWEEN(VAL(mspr),5,12)
                      mspr=mspr+' месяцев'                                                      
                 CASE BETWEEN(VAL(RIGHT(mspr,1)),2,4)
                      mspr=mspr+' месяца'     
                 OTHERWISE 
                      mspr=mspr+' месяц'     
              ENDCASE         
           ELSE
              mspr=''              
           ENDIF 
           
           IF !EMPTY(VAL(dspr))      
              DO CASE 
                 CASE BETWEEN(VAL(dspr),5,20).OR.BETWEEN(VAL(dspr),25,30)
                      dspr=dspr+' дней'                   
                 CASE BETWEEN(VAL(dspr),2,4).OR.BETWEEN(VAL(dspr),22,24) 
                      dspr=dspr+' дня'                        
                 CASE RIGHT(dspr,1)='1'
                      dspr=dspr+' день'                                 
              ENDCASE         
           ELSE
              dspr=''              
           ENDIF
           ON ERROR 
           cStaj=ALLTRIM(yspr+' '+mspr+' '+dspr)
           cStaj=ALLTRIM(yspr)
        ELSE
           cAgeChar=IIF(!EMPTY(peopout.age),DTOC(peopout.age)+'.','')
           DO actualStajToday WITH 'peopout','peopout.date_in','peopout.date_out','cStajWork',.T.
            yspr=''
           mspr=''
           dspr=''
           yspr=LTRIM(LEFT(cStajWork,2))
           mspr=ALLTRIM(SUBSTR(cStajWork,4,2))
           dspr=LTRIM(RIGHT(cStajWork,2))           
           ON ERROR DO erSup
           IF !EMPTY(VAL(yspr))      
              yspr=LTRIM(STR(VAL(yspr)))
              DO CASE 
                 CASE BETWEEN(VAL(yspr),5,20)
                      yspr=yspr+' лет'                   
                 CASE RIGHT(yspr,1)='1'
                      yspr=yspr+' год'                    
                 CASE BETWEEN(VAL(RIGHT(yspr,1)),2,4)
                      yspr=yspr+' года'     
                 OTHERWISE 
                      yspr=yspr+' лет'     
              ENDCASE 
           ELSE
              yspr=''           
           ENDIF     
           IF !EMPTY(VAL(mspr))      
              DO CASE 
                 CASE BETWEEN(VAL(mspr),5,12)
                      mspr=mspr+' месяцев'                                                      
                 CASE BETWEEN(VAL(RIGHT(mspr,1)),2,4)
                      mspr=mspr+' месяца'     
                 OTHERWISE 
                      mspr=mspr+' месяц'     
              ENDCASE         
           ELSE
              mspr=''              
           ENDIF 
           
           IF !EMPTY(VAL(dspr))      
              DO CASE 
                 CASE BETWEEN(VAL(dspr),5,20).OR.BETWEEN(VAL(dspr),25,30)
                      dspr=dspr+' дней'                   
                 CASE BETWEEN(VAL(dspr),2,4).OR.BETWEEN(VAL(dspr),22,24) 
                      dspr=dspr+' дня'                        
                 CASE RIGHT(dspr,1)='1'
                      dspr=dspr+' день'                                 
              ENDCASE         
           ELSE
              dspr=''              
           ENDIF
           ON ERROR 
           cStaj=ALLTRIM(yspr+' '+mspr+' '+dspr)
           cStaj=ALLTRIM(yspr)        
        ENDIF 
        cEduc=IIF(SEEK(IIF(!parUv,people.educ,peopout.educ),'cureducation',1),cureducation.name,'')  
        SELECT curSprav
        REPLACE txtSprav WITH LOWER(cNdol)+' '+LOWER(cNpodr)+' '+cstaj,txtChar WITH cEduc
   
ENDCASE
*************************************************************************************************************************
PROCEDURE prnSprav
PARAMETERS parLog
SELECT curSprav
DO CASE
   CASE dimOpt(1)=1
        DO spravToWord
   CASE dimOpt(2)=1
        DO spravPerToWord     
   CASE dimOpt(4)=1     
        DO charToWord
ENDCASE         
*************************************************************************************************************************
PROCEDURE spravToWord
#DEFINE wdWindowStateMaximize 1

#DEFINE wdBorderTop -1           &&верхняя граница ячейки таблицы
#DEFINE wdBorderLeft -2          &&левая граница ячейки таблицы
#DEFINE wdBorderBottom -3        &&нижняя граница ячейки таблицы
#DEFINE wdBorderRight -4         &&правая граница ячейки таблицы
#DEFINE wdBorderHorizontal -5    &&горизонтальные линии таблицы
#DEFINE wdBorderVertical -6      &&горизонтальные линии таблицы
#DEFINE wdLineStyleSingle 1      && стиль линии границы ячейки (в данно случае обычная)
#DEFINE wdLineStyleNone 0        && линия отсутствует
#DEFINE wdAlignParagraphRight 2
#DEFINE wdAlignParagraphJustify 2
pathdot=ALLTRIM(datset.pathword)+'sprav.dot'
objWord=CREATEOBJECT('WORD.APPLICATION')
#DEFINE cr CHR(13)  
nameDoc=objWord.Documents.Add(pathdot)  
nameDoc.ActiveWindow.View.ShowAll=0   
objWord.WindowState=wdWindowStateMaximize     
objWord.Selection.pageSetup.Orientation=0
objWord.Selection.pageSetup.LeftMargin=40
objWord.Selection.pageSetup.RightMargin=40
objWord.Selection.pageSetup.TopMargin=30
objWord.Selection.pageSetup.BottomMargin=30
docRef=GETOBJECT('','word.basic')
strDatePrik=IIF(!EMPTY(datorder.dateord),dateToString('datorder.dateord',.T.),'')
WITH docRef
     namedoc.tables(1).cell(1,2).Range.Select                
     namedoc.tables(1).cell(6,3).Range.Select                
     docRef.CloseParaBelow  &&Удаляем лишний интервал после абзаца            
     docRef.LineDown                       
 
     namedoc.tables(2).cell(1,1).Range.Text=DTOC(ddspr)                
     namedoc.tables(2).cell(1,3).Range.Text=ALLTRIM(cnspr)
    
     namedoc.tables(3).cell(1,3).Range.Text=ALLTRIM(cadr) 
     namedoc.tables(3).cell(3,1).Range.Text=IIF(!parUv,ALLTRIM(people.fio),ALLTRIM(peopout.fio)) 
     namedoc.tables(3).cell(9,2).Range.Text=ALLTRIM(cursprav.txtsprav)
      
     namedoc.tables(3).cell(15,1).Range.Text='Справка выдана по состоянию на '+DTOC(ddspr)+'г.' 
     namedoc.tables(3).cell(18,1).Range.Text=ALLTRIM(boss.cspravdol)  
     namedoc.tables(3).cell(18,5).Range.Text=ALLTRIM(boss.cspravfio)  
ENDWITH   
objWord.Visible=.T.    
*************************************************************************************************************************
PROCEDURE spravPerToWord
#DEFINE wdWindowStateMaximize 1
#DEFINE wdBorderTop -1           &&верхняя граница ячейки таблицы
#DEFINE wdBorderLeft -2          &&левая граница ячейки таблицы
#DEFINE wdBorderBottom -3        &&нижняя граница ячейки таблицы
#DEFINE wdBorderRight -4         &&правая граница ячейки таблицы
#DEFINE wdBorderHorizontal -5    &&горизонтальные линии таблицы
#DEFINE wdBorderVertical -6      &&горизонтальные линии таблицы
#DEFINE wdLineStyleSingle 1      && стиль линии границы ячейки (в данно случае обычная)
#DEFINE wdLineStyleNone 0        && линия отсутствует
#DEFINE wdAlignParagraphRight 2
#DEFINE wdAlignParagraphJustify 2
pathdot=ALLTRIM(datset.pathword)+'spravper.dotx'
objWord=CREATEOBJECT('WORD.APPLICATION')
#DEFINE cr CHR(13)  
nameDoc=objWord.Documents.Add(pathdot)  
nameDoc.ActiveWindow.View.ShowAll=0   
objWord.WindowState=wdWindowStateMaximize     
objWord.Selection.pageSetup.Orientation=0
objWord.Selection.pageSetup.LeftMargin=40
objWord.Selection.pageSetup.RightMargin=40
objWord.Selection.pageSetup.TopMargin=30
objWord.Selection.pageSetup.BottomMargin=30
docRef=GETOBJECT('','word.basic')
strDatePrik=IIF(!EMPTY(datorder.dateord),dateToString('datorder.dateord',.T.),'')
WITH docRef
     namedoc.tables(1).cell(1,2).Range.Select                
     namedoc.tables(1).cell(6,3).Range.Select                
     docRef.CloseParaBelow  &&Удаляем лишний интервал после абзаца            
     docRef.LineDown                       
 
     namedoc.tables(2).cell(1,1).Range.Text=DTOC(ddspr)                
     namedoc.tables(2).cell(1,3).Range.Text=ALLTRIM(cnspr)
    
     namedoc.tables(3).cell(1,3).Range.Text=ALLTRIM(cadr) 
     namedoc.tables(3).cell(3,1).Range.Text=ALLTRIM(people.fio) 
     namedoc.tables(3).cell(5,1).Range.Text=cDateIn
     namedoc.tables(3).cell(5,2).Range.Select
     docRef.LineDown 
     docRef.CloseParaBelow    
     .Insert(cr)     
     .LeftPara
     .Font('Times New Roman',12)
     .Insert(ALLTRIM(cursprav.txtsprav))
      docRef.CloseParaBelow 
     .Insert(cr)    
     .LeftPara
     .Font('Times New Roman',12)
     .Insert(ALLTRIM(cursprav.txtchar))  
     .Insert(cr)  
     .Insert(cr)  
     .LeftPara
     .Font('Times New Roman',12)
     .Insert('Дополнительные сведения')  
     .Insert(cr)
     .Font('Times New Roman',12)    
     .Insert('     Справка выдана по состоянию на '+DTOC(ddspr)+'г.')  
     .Insert(cr)    
     .Font('Times New Roman',12)
     .Insert('     Срок действия справки - бессрочно')
      
     .Insert(cr)  
     .Insert(cr)   
     .Insert(cr)   
     nameDoc.Tables.add(objWord.Selection.range,2,5)
     ordTable4=nameDoc.Tables(4) 
     WITH ordTable4
          .Columns(1).Width=150
          .Columns(2).Width=40
          .Columns(3).Width=150
          .Columns(4).Width=40
          .Columns(5).Width=150
          .cell(1,1).Range.Select 
          docRef.CenterPara    
          docRef.Font('Times New Roman',12)
          .cell(1,1).Range.Text=ALLTRIM(boss.cspravdol)  
          docRef.CloseParaBelow 
          .cell(1,3).Range.Select     
          docRef.Font('Times New Roman',12)
          docRef.CloseParaBelow 
          .cell(1,5).Range.Select
          docRef.CenterPara         
          docRef.Font('Times New Roman',12)
          .cell(1,5).Range.Text=ALLTRIM(boss.cspravfio)  
           docRef.CloseParaBelow 
          .cell(1,1).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(1,3).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(1,5).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          .cell(2,1).Range.Select 
          docRef.CenterPara    
          docRef.Font('Times New Roman',6)
          .cell(2,1).Range.Text='(руководитель)'
          docRef.CloseParaBelow 
          .cell(2,3).Range.Select 
          docRef.CenterPara    
          docRef.Font('Times New Roman',6)
          .cell(2,3).Range.Text='(подпись)'
          docRef.CloseParaBelow 
          .cell(2,5).Range.Select 
          docRef.CenterPara    
          docRef.Font('Times New Roman',6)
          .cell(2,5).Range.Text='(инициалы, фамилия)'
          docRef.CloseParaBelow 
     ENDWITH         
      *  .LeftPara
    * namedoc.tables(3).cell(7,1).Range.Text=ALLTRIM(cursprav.txtsprav)
      
    
     *namedoc.tables(4).cell(1,1).Range.Text=ALLTRIM(boss.cspravdol)  
     *namedoc.tables(4).cell(1,5).Range.Text=ALLTRIM(boss.cspravfio)  
ENDWITH   
objWord.Visible=.T.    
*************************************************************************************************************************
PROCEDURE charToWord
#DEFINE wdWindowStateMaximize 1
pathdot=ALLTRIM(datset.pathword)+'char.dotx'
objWord=CREATEOBJECT('WORD.APPLICATION')
nameDoc=objWord.Documents.Add(pathdot)  
objWord.WindowState=wdWindowStateMaximize   
*nameDoc.ActiveWindow.View.ShowAll=0   
IF TYPE([nameDoc.formFields("cdate")])="O"
        nameDoc.FormFields("cdate").Result=DTOC(dDspr)
ENDIF
IF TYPE([nameDoc.formFields("cfio")])="O"
        nameDoc.FormFields("cfio").Result=cFio
ENDIF
IF TYPE([nameDoc.formFields("cage")])="O"
        nameDoc.FormFields("cage").Result=cAgeChar
ENDIF
IF TYPE([nameDoc.formFields("ceduc")])="O"
        nameDoc.FormFields("ceduc").Result=ALLTRIM(cursprav.txtchar)
ENDIF
IF TYPE([nameDoc.formFields("ccharw")])="O"
        nameDoc.FormFields("ccharw").Result=ALLTRIM(cursprav.txtsprav)
ENDIF   
IF TYPE([nameDoc.formFields("cstaj")])="O"
        nameDoc.FormFields("cstaj").Result=cstaj
ENDIF 
IF TYPE([nameDoc.formFields("cdolboss")])="O"
        nameDoc.FormFields("cdolboss").Result=ALLTRIM(boss.cspravdol)
ENDIF
IF TYPE([nameDoc.formFields("cfboss")])="O"
        nameDoc.FormFields("cfboss").Result=ALLTRIM(boss.cspravfio)
ENDIF
objWord.Visible=.T.              