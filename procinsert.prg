PARAMETERS parTask
** parTask 1 - �����
** parTask 1 - �����������
nTask=parTask
dStaj=DATE()
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='�������'
     .Icon='kone.ico'
     DO addshape WITH 'fSupl',1,10,10,150,450,8    
     
     DO adLabMy WITH 'fSupl',1,'���� �� ',.Shape1.Top+20,.Shape1.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxBeg',.Shape1.Top+20,.Shape1.Left,RetTxtWidth('99/99/99999'),dHeight,'dStaj',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3
      
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxBeg.Width-5)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+5
     .Shape1.Height=.boxBeg.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.T.,.F.
     *---------------------------------������ ������-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('��������W')*3-20)/2,.Shape91.Top+.Shape91.Height+20,RetTxtWidth('��������W'),dHeight+5,'������','DO prnInsert WITH 1','������'
     *---------------------------------������ ���������������� ���������-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'��������','DO prnInsert WITH 2','��������������� ��������'   
     *---------------------------------������ ����� �� ����� ������----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+10,.Cont1.Top,.Cont1.Width,.cont1.Height,'�������','fSupl.Release','�������'     
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
   kvalIns=IIF(SEEK(people.kval,'sprkval',1),ALLTRIM(sprkval.name),'')+'   ������ '+ALLTRIM(people.nordkval)+'  '+IIF(!EMPTY(people.dkval),DTOC(people.dkval),'')
ENDIF 


DO CASE 
   CASE nTask=1
        podrIns=IIF(SEEK(curJobsupl.kp,'sprpodr',1),sprpodr.namework,'')
        dolIns=IIF(SEEK(curJobsupl.kd,'sprdolj',1),sprdolj.name,'')
        kseIns=LTRIM(STR(curJobSupl.kse,4,2))              
 
        IF curJobSupl.lkv.AND.!EMPTY(people.kval)
           kvalIns=IIF(SEEK(people.kval,'sprkval',1),ALLTRIM(sprkval.name),'')+'   ������ '+ALLTRIM(people.nordkval)+'  '+IIF(!EMPTY(people.dkval),DTOC(people.dkval),'')
        ENDIF 
        DO actualStajToday WITH 'people','people.date_in','dStaj','stajIns'
        stajIns=ALLTRIM(stajIns)+' �� '+DTOC(dStaj)
        DO perStajOne WITH 'stajIns','Dstaj'
   CASE nTask=2
        podrIns=IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.namework,'')
        dolIns=IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.name,'')
        kseIns=LTRIM(STR(datjob.kse,4,2))
        DO actualStajToday WITH 'people','people.date_in','dStaj'
        stajIns=ALLTRIM(people.staj_today)+' �� '+DTOC(dStaj)
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
   stajIns=ALLTRIM(stajIns)+' ����������� - '+IIF(nTask=1,DTOC(people.dPerSt),DTOC(datjob.per_date)) 
ENDIF

IF people.mols
   stajIns=ALLTRIM(stajIns)+' ������� ����. �� - '+DTOC(people.dmol) 
ENDIF
DO CASE 
   CASE par1=1 
        DO procForPrintAndPreview WITH 'repinsert','�������',.T.,'insertToWord'
   CASE par1=2 
        DO procForPrintAndPreview WITH 'repinsert','�������',.F.,'insertToWord'        
ENDCASE
SELECT people
GO inRec

**************************************************************************
PROCEDURE insertToWord
#DEFINE wdBorderTop -1           &&������� ������� ������ �������
#DEFINE wdBorderLeft -2          &&����� ������� ������ �������
#DEFINE wdBorderBottom -3        &&������ ������� ������ �������
#DEFINE wdBorderRight -4         &&������ ������� ������ �������
#DEFINE wdBorderHorizontal -5    &&�������������� ����� �������
#DEFINE wdBorderVertical -6      &&�������������� ����� �������
#DEFINE wdLineStyleSingle 1      && ����� ����� ������� ������ (� ����� ������ �������)
#DEFINE wdLineStyleNone 0        && ����� �����������
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
     .Insert('���ר� ������ (����������� �� ������� ����� ������������ (�������������) �� ������)')
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
          .Cell(2,1).Range.Text='������������ ������������ �������������'
         
          .Cell(2,2).Width=290
          .Cell(2,2).Range.Text=podrIns
          .cell(2,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(3,1).Width=150
          .Cell(3,1).Range.Text='������������ ���������'
          .Cell(3,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(3,2).Width=370
          .Cell(3,2).Range.Text=dolIns
          .Cell(3,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(3,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(4,1).Width=160
          .Cell(4,1).Range.Text='���������������� ���������'
          .Cell(4,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(4,2).Width=360
          .Cell(4,2).Range.Text=kvalIns
          .Cell(4,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(4,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .Cell(5,1).Width=210
          .Cell(5,1).Range.Text='���� ������ � ��������� ������������'
          .Cell(5,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(5,2).Width=310
          .Cell(5,2).Range.Text=stajIns
          .Cell(5,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          .Cell(5,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          
          .Cell(6,1).Width=70
          .Cell(6,1).Range.Text='�����'
          .Cell(6,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(6,1).Select  
          docRef.Bold  
          .Cell(6,2).Width=450
          .Cell(6,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .cell(6,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(7,1).Range.Text='�������� �� ���� ������ � ��������� ������������'        
          .Cell(7,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .Cell(7,2).Borders(wdBorderTop).LineStyle=wdLineStyleNone
          .cell(7,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(8,1).Range.Text='�������� �� �������� ������ ��� �29 �� 26.07.1999' 
          .Cell(8,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(8,2).Range.Text=kontraktIns   
          .cell(8,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(9,1).Range.Text='����. 52 �.3 �������� �� ���'
          .Cell(9,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(9,2).Range.Text=vtoIns
          .cell(9,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(10,1).Range.Text='����. 52 �.4.1 �������� �� ��������� ������ � ����� ���������������'  
          .Cell(10,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(10,2).Range.Text=katIns
          .cell(10,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(11,1).Range.Text='����. 52 �.4.5 �������� ������-������������ �������������� ������� 40%'  
          .Cell(11,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(11,2).Range.Text=chirIns
          .cell(11,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(12,1).Range.Text='����. 52 �.4.6 �������� ������-��������, ����������-�������� 25%'  
          .Cell(12,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(12,2).Range.Text=intIns
          .cell(12,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(13,1).Range.Text='����. 52 �.5 �������� �� ������ � ����� ���������������'  
          .Cell(13,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(13,2).Range.Text=zdravIns
          .cell(13,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(14,1).Range.Text='����. 52 �.6.1 ������� �� ����'  
          .Cell(14,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(14,2).Range.Text=mainIns
          .cell(14,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(15,1).Range.Text='����. 52 �.6.6 �������� �� "�������"'  
          .Cell(15,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(15,2).Range.Text=main2Ins
          .cell(15,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle 
          
          .cell(16,1).Range.Text='����. 53 �.3   �������� ������� ������������'  
          .Cell(16,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(16,2).Range.Text=molsIns
          .cell(16,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(17,1).Range.Text='����. 53 �.4   ������� �� ����������� ���������������� ������������'  
          .Cell(17,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(17,2).Range.Text=osobIns
          .cell(17,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(18,1).Range.Text='����. 53 �.9   ������� �� ������ �������� �����'  
          .Cell(18,1).Borders(wdBorderTop).LineStyle=wdLineStyleNone 
          .Cell(18,2).Range.Text=charwIns
          .cell(18,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(19,1).Range.Text='��ڨ� ������ �� ������ ���������'  
          .cell(19,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          .Cell(19,2).Range.Text=kseIns   
          
          .cell(20,1).Range.Text='���������� �� ������'  
          .cell(20,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle
          
          .cell(21,1).Range.Text='���������'  
          .cell(21,2).Borders(wdBorderBottom).LineStyle=wdLineStyleSingle   
          
        *  docRef.CloseParaBelow  &&������� ������ �������� ����� ������            
          docRef.LineDown 
     ENDWITH   
ENDWITH     
objWord.Visible=.T.       