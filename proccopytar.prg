fSupl=CREATEOBJECT('FORMSUPL')
repFullName='��������' 
labname='������� ����� ����������� �� '+DTOC(vardtar)+'?'
labname2='����������� �� '+DTOC(vardtar)+' ��� ����������� !'
WITH fSupl
     .Caption='������� � ����'
     .Width=RetTxtWidth('w������������w')*4+40
     DO adLabMy WITH 'fSupl',1,labname,20,10,.Width-20,2
     DO adLabMy WITH 'fSupl',2,labname2,20,10,.Width-20,2
     .lab2.Visible=.F.
     DO adLabMy WITH 'fSupl',3,'������������ �����',.lab1.Top+.lab1.Height,10,.Width-20,2
     DO adTboxNew WITH 'fSupl','boxFullName',.lab3.Top+.lab3.Height,10,.Width-20,dHeight,'repFullName',.F.,.T.,0,.F.
     
     DO addcontlabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('w�������w')*2-20)/2,.boxFullName.Top+.boxFullName.Height+20,RetTxtWidth('w�������w'),dHeight+3,'�������','DO archivtarif'         
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'�������','fSupl.Release'      
     
     DO addcontlabel WITH 'fSupl','contNew',(.Width-RetTxtWidth('w������������w')*3-20)/2,.cont1.Top,RetTxtWidth('w������������w'),dHeight+3,'��������','DO archivtarifrew WITH 1'         
     DO addcontlabel WITH 'fSupl','contRew',.ContNew.Left+.ContNew.Width+10,.Cont1.Top,.ContNew.Width,dHeight+3,'������������','DO archivtarifrew WITH 2'         
     DO addcontlabel WITH 'fSupl','contRet',.ContRew.Left+.ContRew.Width+10,.Cont1.Top,.ContNew.Width,dHeight+3,'�������','fSupl.Release'      
     .contNew.Visible=.F.
     .contRew.Visible=.F.
     .contRet.Visible=.F.
     
     .Height=.lab1.Height*2+.boxFullName.Height+.cont1.Height+70
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************************************************************************
*
*******************************************************************************************************************************************
PROCEDURE archivtarif
SELECT datshtat
LOCATE FOR !real.AND.dtarif=vardtar
reppathtar='TAR'+DTOC(vardtar)
IF !FOUND()
   pathcopy=pathmain+'\TAR'+DTOC(vardtar)  
   DO createarchiv WITH .T.
   fSupl.Release
ELSE 
   WITH fSupl   
        .lab1.Visible=.F.
        .lab2.Visible=.T.
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
PROCEDURE archivtarifrew
PARAMETERS par1
*1 - �����
*2 - ������������
DO CASE 
   CASE par1=1
        ncx=1
        DO WHILE .T.
           nametarsup='TAR'+DTOC(vardtar)+'_'+LTRIM(STR(ncx))
           LOCATE FOR ALLTRIM(pathtarif)=nametarsup
           IF !FOUND()
               EXIT
           ENDIF
           ncx=ncx+1
        ENDDO
        pathcopy=pathmain+'\'+nametarsup 
        reppathtar=nametarsup
        DO createarchiv WITH .T.
        fSupl.Release
   CASE par1=2
        pathcopy=pathmain+'\TAR'+DTOC(vardtar)  
        reppathtar='TAR'+DTOC(vardtar)
        DO createarchiv WITH .F.
        fSupl.Release
ENDCASE 
*******************************************************************************************************************************************
*
*******************************************************************************************************************************************
PROCEDURE createarchiv
PARAMETERS lognew
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
REPLACE dtarif WITH varDtar,pathtarif WITH reppathtar, basest WITH varBaseSt,dcreate WITH DATETIME(),fullname WITH repFullName