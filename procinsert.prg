PARAMETERS parTask
** parTask 1 - кадры
** parTask 1 - тарификация
nTask=parTask
dStaj=DATE()
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Вкладыш'
     .Icon='kone.ico'
     DO addshape WITH 'fSupl',1,10,10,150,450,8    
     
     DO adLabMy WITH 'fSupl',1,'стаж на ',.Shape1.Top+20,.Shape1.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxBeg',.Shape1.Top+20,.Shape1.Left,RetTxtWidth('99/99/99999'),dHeight,'dStaj',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3
      
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxBeg.Width-5)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+5
     .Shape1.Height=.boxBeg.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.T.,.F.
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('просмотрW')*3-20)/2,.Shape91.Top+.Shape91.Height+20,RetTxtWidth('просмотрW'),dHeight+5,'печать','DO prnInsert WITH 1','печать'
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'просмотр','DO prnInsert WITH 2','предварительный просмотр'   
     *---------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'возврат','fSupl.Release','возврат'     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+60
     
     .Width=.Shape1.Width+20   
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***************************************************************************
PROCEDURE prnInsert
PARAMETERS par1
STORE '' TO fioIns,podrIns,dolIns,stajIns,kvalIns,kontraktIns,stajReal,kseIns,zdravIns,oklIns,stIns,katIns,chirIns,intIns,mainIns,main2Ins,molsIns,osobIns,charwIns,vtoIns
IF EMPTY(dStaj)
   RETURN
ENDIF
SELECT people
inRec=RECNO()
fioIns=ALLTRIM(people.fio)

*kontraktIns=IIF(people.pkont#0,LTRIM(STR(people.pkont))+'%','')+IIF(nTask=2.AND.datjob.pkont#0,' - '+LTRIM(STR(datjob.mkonts,10,2)),'')
kontraktIns=IIF(nTask=1.AND.people.pkont#0,LTRIM(STR(people.pkont))+'%',IIF(nTask=2.AND.datjob.pkont#0,LTRIM(STR(datjob.pkont))+'%'+' - '+LTRIM(STR(datjob.mkonts,10,2)),''))
ON ERROR DO erSup
   zdravIns=IIF(nTask=1,'',IIF(datjob.pzdrav#0,LTRIM(STR(datjob.pzdrav))+'%','')+IIF(nTask=2.AND.datjob.pzdrav#0,' - '+LTRIM(STR(datjob.mzdrav,10,2)),''))
ON ERROR
 
IF IIF(nTask=1,curJobSupl.lkv,datjob.lkv).AND.!EMPTY(people.kval)
   kvalIns=IIF(SEEK(people.kval,'sprkval',1),ALLTRIM(sprkval.name),'')+'   приказ '+ALLTRIM(people.nordkval)+'  '+IIF(!EMPTY(people.dkval),DTOC(people.dkval),'')
ENDIF 


DO CASE 
   CASE nTask=1
        podrIns=IIF(SEEK(curJobsupl.kp,'sprpodr',1),sprpodr.namework,'')
        dolIns=IIF(SEEK(curJobsupl.kd,'sprdolj',1),sprdolj.name,'')
        kseIns=LTRIM(STR(curJobSupl.kse,4,2))              
 
        IF curJobSupl.lkv.AND.!EMPTY(people.kval)
           kvalIns=IIF(SEEK(people.kval,'sprkval',1),ALLTRIM(sprkval.name),'')+'   приказ '+ALLTRIM(people.nordkval)+'  '+IIF(!EMPTY(people.dkval),DTOC(people.dkval),'')
        ENDIF 
        DO actualStajToday WITH 'people','people.date_in','dStaj','stajIns'
        stajIns=ALLTRIM(stajIns)+' на '+DTOC(dStaj)
        DO perStajOne WITH 'stajIns','Dstaj'
   CASE nTask=2
        podrIns=IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.namework,'')
        dolIns=IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.name,'')
        kseIns=LTRIM(STR(datjob.kse,4,2))
        DO actualStajToday WITH 'people','people.date_in','dStaj'
        stajIns=ALLTRIM(people.staj_today)+' на '+DTOC(dStaj)
        oklIns=LTRIM(STR(datjob.mtokl,10,2))
        stIns=LTRIM(STR(datjob.stPr,3))+' % - '+LTRIM(STR(datjob.mstsum,10,2))
        katIns=IIF(datjob.pkat#0,LTRIM(STR(datjob.pkat,3))+' % - '+LTRIM(STR(datjob.mkat,10,2)),'')
        vtoIns=IIF(datjob.pvto#0,LTRIM(STR(datjob.pvto,3))+' % - '+LTRIM(STR(datjob.pvto,10,2)),'')
        chirIns=IIF(datjob.pchir#0,LTRIM(STR(datjob.pchir,3))+' % - '+LTRIM(STR(datjob.mchir,10,2)),'')
        intIns=IIF(datjob.pint#0,LTRIM(STR(datjob.pint,3))+' % - '+LTRIM(STR(datjob.mint,10,2)),'')
        mainIns=IIF(datjob.pmain#0,LTRIM(STR(datjob.pmain,3))+' % - '+LTRIM(STR(datjob.mmain,10,2)),'')
        main2Ins=IIF(datjob.pmain2#0,LTRIM(STR(datjob.pmain2,3))+' % - '+LTRIM(STR(datjob.mmain2,10,2)),'')
        molsIns=IIF(datjob.pmols#0,LTRIM(STR(datjob.pmols,3))+' % - '+LTRIM(STR(datjob.mmols,10,2)),'')
        osobIns=IIF(datjob.posob#0,LTRIM(STR(datjob.posob,3))+' % - '+LTRIM(STR(datjob.mosob,10,2)),'')
        charwIns=IIF(datjob.pcharw#0,LTRIM(STR(datjob.pcharw,3))+' % - '+LTRIM(STR(datjob.mcharw,10,2)),'')
        DO perStajOne 
ENDCASE 

IF !EMPTY(IIF(nTask=1,people.dPerSt,datjob.per_date))
   stajIns=ALLTRIM(stajIns)+' переходящий - '+IIF(nTask=1,DTOC(people.dPerSt),DTOC(datjob.per_date)) 
ENDIF

IF people.mols
   stajIns=ALLTRIM(stajIns)+' молодой спец. до - '+DTOC(people.dmol) 
ENDIF
DO CASE 
   CASE par1=1 
        DO procForPrintAndPreview WITH 'repinsert','вкладыш',.T.,'insertToWord'
   CASE par1=2 
        DO procForPrintAndPreview WITH 'repinsert','вкладыш',.F.,'insertToWord'        
ENDCASE
SELECT people
GO inRec

**************************************************************************
PROCEDURE insertToWord
#DEFINE wdBorderTop -1           &&верхняя граница ячейки таблицы
#DEFINE wdBorderLeft -2          &&левая граница ячейки таблицы
#DEFINE wdBorderBottom -3        &&нижняя граница ячейки таблицы
#DEFINE wdBorderRight -4         &&правая граница ячейки таблицы
#DEFINE wdBorderHorizontal -5    &&горизонтальные линии таблицы
#DEFINE wdBorderVertical -6      &&горизонтальные линии таблицы
#DEFINE wdLineStyleSingle 1      && стиль линии границы ячейки (в данно случае обычная)
#DEFINE wdLineStyleNone 0        && линия отсутствует
#DEFINE wdAlignParagraphRight 2

objWord=CREATEOBJECT('WORD.APPLICATION')
#DEFINE cr CHR(13)
nameDoc=objWord.Documents.Add()  
nameDoc.ActiveWindow.View.ShowAll=0        
objWord.Selection.pageSetup.Orientation=0
objWord.Selection.pageSetup.LeftMargin=30
objWord.Selection.pageSetup.RightMargin=20
objWord.Selection.pageSetup.TopMargin=20
objWord.Selection.pageSetup.BottomMargin=10
docRef=GETOBJECT('','word.basic')

WITH docRef
     .Insert(cr)
     .Font('Times New Roman',12)
     .CenterPara 
     .Bold
     .Insert('РАСЧЁТ ОКЛАДА (заполняется на каждого вновь поступившего (переведенного) на работу)')
     .Insert(cr)
     .Font('Times New Roman',12)
     .LeftPara
     nameDoc.Tables.add(objWord.Selection.range,1,2)
     ordTable1=nameDoc.Tables(1) 
    
     WITH ordTable1
          .Columns(1).Width=380
          .cell(1,1).Range.Select   
          docRef.Font('Times New Roman',11)
          .Columns(2).Width=140 
          .cell(1,2).Range.Select     
          docRef.Font('Times New Roman',11)    
          
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
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
          .Rows.Add
              
          .Cell(1,1).Merge(.Cell(1,2))
          .Cell(1,1).Range.Text=fioIns
          .Cell(1,1).Select
          docRef.CenterPara
          docRef.Bold
          .cell(1,1).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(2,1).Width=230   
          .Cell(2,1).Range.Text='Наименование структурного подразделения'
         
          .Cell(2,2).Width=290
          .Cell(2,2).Range.Text=podrIns
          .cell(2,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(3,1).Width=150
          .Cell(3,1).Range.Text='Наименование должности'
          .Cell(3,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(3,2).Width=370
          .Cell(3,2).Range.Text=dolIns
          .Cell(3,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(3,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(4,1).Width=160
          .Cell(4,1).Range.Text='Квалификационная категория'
          .Cell(4,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(4,2).Width=360
          .Cell(4,2).Range.Text=kvalIns
          .Cell(4,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(4,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(5,1).Width=210
          .Cell(5,1).Range.Text='Стаж работы в бюджетных организациях'
          .Cell(5,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(5,2).Width=310
          .Cell(5,2).Range.Text=stajIns
          .Cell(5,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          .Cell(5,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          
          .Cell(6,1).Width=70
          .Cell(6,1).Range.Text='ОКЛАД'
          .Cell(6,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(6,1).Select  
          docRef.Bold  
          .Cell(6,2).Width=450
          .Cell(6,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .cell(6,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(7,1).Range.Text='Надбавка за стаж работы в бюджетных организациях'        
          .Cell(7,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(7,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .cell(7,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(8,1).Range.Text='Надбавка за контракт Декрет ПРБ №29 от 26.07.1999' 
          .Cell(8,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(8,2).Range.Text=kontraktIns   
          .cell(8,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(9,1).Range.Text='Пост. 52 п.3 Надбавка за ВТО'
          .Cell(9,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(9,2).Range.Text=vtoIns
          .cell(9,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(10,1).Range.Text='Пост. 52 п.4.1 Надбавка за специфику работы в сфере здравоохранения'  
          .Cell(10,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(10,2).Range.Text=katIns
          .cell(10,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(11,1).Range.Text='Пост. 52 п.4.5 Надбавка врачам-специалистам хирургического профиля 40%'  
          .Cell(11,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(11,2).Range.Text=chirIns
          .cell(11,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(12,1).Range.Text='Пост. 52 п.4.6 Надбавка врачам-интернам, провизорам-интернам 25%'  
          .Cell(12,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(12,2).Range.Text=intIns
          .cell(12,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(13,1).Range.Text='Пост. 52 п.5 Надбавка за работу в сфере здравоохранения'  
          .Cell(13,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(13,2).Range.Text=zdravIns
          .cell(13,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(14,1).Range.Text='Пост. 52 п.6.1 Доплата за РОРФ'  
          .Cell(14,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(14,2).Range.Text=mainIns
          .cell(14,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(15,1).Range.Text='Пост. 52 п.6.6 Надбавка за "старший"'  
          .Cell(15,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(15,2).Range.Text=main2Ins
          .cell(15,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(16,1).Range.Text='Пост. 53 п.3   Надбавка молодым специалистам'  
          .Cell(16,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(16,2).Range.Text=molsIns
          .cell(16,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(17,1).Range.Text='Пост. 53 п.4   Доплата за особенности профессиональной деятельности'  
          .Cell(17,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(17,2).Range.Text=osobIns
          .cell(17,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(18,1).Range.Text='Пост. 53 п.9   Доплата за особый характер труда'  
          .Cell(18,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(18,2).Range.Text=charwIns
          .cell(18,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(19,1).Range.Text='ОБЪЁМ РАБОТЫ ПО ДАННОЙ ДОЛЖНОСТИ'  
          .cell(19,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          .Cell(19,2).Range.Text=kseIns   
          
          .cell(20,1).Range.Text='Специалист по кадрам'  
          .cell(20,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(21,1).Range.Text='Экономист'  
          .cell(21,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle   
          
        *  docRef.CloseParaBelow  &&Удаляем лишний интервал после абзаца            
          docRef.LineDown 
     ENDWITH   
ENDWITH     
objWord.Visible=.T.       