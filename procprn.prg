*log_prim - ������ ����������
*log_sr  - ������� �����
*log_kom - ��������������� ��������
*log_it - ����� �� ���������� ��������� ����� ���������
*avt_vac - �������� �������������
*vacst - �������� �� ����� 1 ������
*log_vac - ��������� ���������
logKom=.F. &&��� ������ ��������
kpdop=0
margTop=000.00
nlspace=000.00

IF !USED('fete')
   USE fete ORDER 1 IN 0 
ENDIF
SELECT people
fltrec=RECNO()
SET FILTER TO 
IF USED('curSpis')
   SELECT curSpis
   USE
ENDIF
IF USED('nadBase')
   SELECT nadbase
   USE    
ENDIF
USE nadBase ORDER 1 IN 0 
IF USED('formbase')
   SELECT formbase
   USE 
ENDIF
USE formbase IN 0
SELECT formbase
SET FILTER TO log_prn
GO TOP

*SELECT rec,persved,sumved,countvac,logkse,primf FROM tarfond WHERE !EMPTY(sumved) INTO CURSOR curnadbav ORDER BY num

DIMENSION dim_boss(20,2)
STORE '' TO dim_boss
SELECT boss
i=1
SCAN ALL 
     dim_boss(i,1)=doljn
     dim_boss(i,2)=fam
     i=i+1
ENDSCAN 

SELECT sprtype
SET ORDER TO 1
COUNT TO max_tr
DIMENSION kod_tr(max_tr),name_tr(max_tr)
STORE '' TO name_tr
STORE 0 TO kod_tr
GO TOP
FOR i=1 TO max_tr
    name_tr(i)=name
    kod_tr(i)=kod
    SKIP  
ENDFOR

SELECT sprkat
COUNT TO max1_kat
max_kat=6
DIMENSION name_kat(10),kod_kat(10),dim_kat(10),name1_kat(10),name2_kat(10)
STORE '' TO name_kat,name1_kat,fltch,fltpodr,fltkat,fltType,spisnum,fltDol
STORE 0 TO kod_kat,dim_kat
GO TOP
FOR i=1 TO max1_kat
    name_kat(i)=name
    name1_kat(i)=IIF(!EMPTY(namefull),namefull,name)
    name2_kat(i)=IIF(!EMPTY(namefull1),namefull1,name)
    kod_kat(i)=kod
    SKIP  
ENDFOR


DIMENSION dimOption(6)
STORE .F. TO dimOption
*dimOption(1) - ������ ������
*dimOption(2) - �������������
*dimOption(3) - ��������
*dimOption(4) - ������
*dimOption(5) - ��� ������
*dimOption(6) - ���������

SELECT * FROM sprdolj INTO CURSOR dopDol READWRITE
SELECT dopDol
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM datagrup INTO CURSOR dopGroup READWRITE
SELECT dopGroup
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprkat INTO CURSOR dopKat READWRITE
SELECT dopKat
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprtype INTO CURSOR dopType READWRITE
SELECT dopType
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM datJob INTO CURSOR curSpis READWRITE
ALTER TABLE curSpis ADD COLUMN otm C(3)
ALTER TABLE curSpis ADD COLUMN fl L
ALTER TABLE curSpis ADD COLUMN ksefio C(70)
SELECT curSpis
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE kseFio WITH STR(kse,4,2)+'   '+fio ALL
INDEX ON fio TAG T1

SELECT datset
IF nstru#0   
   DO CASE 
       CASE nstru=1.AND.!EMPTY(cnamegr)
            dimOption(nstru)=.T.
            SELECT dopgroup
            LOCATE FOR ALLTRIM(LOWER(name))==ALLTRIM(LOWER(datset.cnamegr))
            REPLACE fl WITH .T.,otm WITH ' � ' 
   ENDCASE 
ENDIF 

onlyVac=.F.
excludeVac=.F.
terminal=1
nameGroup=''
nameForm=ALLTRIM(formbase.name)
namenadbav=ALLTRIM(nadBase.nadbav)
persnadbav=ALLTRIM(nadBase.persved)
sumnadbav=ALLTRIM(nadBase.sumved)
sumOklSt=ALLTRIM(nadBase.oklSt)
nsupproc=nadBase.supproc
persFormula=ALLTRIM(nadBase.countvac)
persKse=nadBase.logKse
headNadbav=ALLTRIM(nadBase.nHead)

dateTar=varDTar
begTar=CTOD('  .  .    ')
endTar=CTOD('  .  .    ')
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Addproperty('colprn(10)','')
     .Caption='���������'
     .procexit='Do exitprn'
     DO addshape WITH 'fSupl',1,10,10,150,500,8 
     DO adTboxAsCont WITH 'fSupl','txtVed',.Shape1.Left+10,.Shape1.Top+10,.Shape1.Width-RetTxtWidth('w99/99/9999w')-19,dHeight,'���������',2,1
     DO adTboxAsCont WITH 'fSupl','txtDate',.txtVed.Left+.txtVed.Width-1,.txtVed.Top,RetTxtWidth('w99/99/9999w'),dHeight,'����',2,1
     
     
     DO addcombomy WITH 'fSupl',1,.txtVed.Left,.txtVed.Top+.txtved.Height-1,dHeight,.txtVed.Width,.T.,'nameform','ALLTRIM(formbase.name)',6,.F.,'Do validformbase',.F.,.T.             
     WITH .comboBox1         && ���������   
          .nDisplayCount=18
     ENDWITH
     DO adTboxNew WITH 'fSupl','boxDate',.comboBox1.Top,.txtDate.Left,.txtDate.Width,dHeight,'dateTar',.F.,.T.,0
     DO adTboxAsCont WITH 'fSupl','txtNad',.txtVed.Left,.comboBox1.Top+.comboBox1.Height-1,.txtVed.Width+.txtDate.Width-1,dHeight,'��������/�������',2,1
     DO addcombomy WITH 'fSupl',2,.txtVed.Left,.txtNad.Top+.txtNad.Height-1,dHeight,.txtNad.Width,.T.,'namenadbav','ALLTRIM(nadbase.nadbav)',6,.F.,'Do validformnadbav',.F.,.T.             
     .comboBox2.Enabled=.F.
     .comboBox2.DisplayCount=18
     DO adTboxAsCont WITH 'fSupl','txtPeriod',.txtVed.Left,.comboBox2.Top+.comboBox2.Height-1,.txtNad.Width-.txtDate.Width*2+2,dHeight,'������',0,1
     DO adTboxNew WITH 'fSupl','boxBeg',.txtPeriod.Top,.txtPeriod.Left+.txtPeriod.Width-1,.txtDate.Width,dHeight,'begTar',.F.,.F.,0
     DO adTboxNew WITH 'fSupl','boxEnd',.txtPeriod.Top,.boxBeg.Left+.boxBeg.Width-1,.txtDate.Width,dHeight,'endTar',.F.,.F.,0
     
     .Shape1.Height=.txtVed.Height*3+.comboBox1.Height*2+20
     
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8 
     
     DO adCheckBox WITH 'fSupl','checkGroup','������ ������',.Shape2.Top+10,.Shape2.Left+5,150,dHeight,'dimOption(1)',0,.T.,'DO validCheckGroup'          
     DO adCheckBox WITH 'fSupl','checkPodr','�������������',.checkGroup.Top+.checkGroup.Height+10,.Shape2.Left+5,150,dHeight,'dimOption(2)',0,.T.,'DO validCheckPodr'  
             
     DO adCheckBox WITH 'fSupl','checkPers','��������',.checkGroup.Top,.Shape2.Left+5,150,dHeight,'dimOption(3)',0,.T.,'DO validCheckPers' 
     DO adCheckBox WITH 'fSupl','checkSpis','������',.checkPodr.Top,.Shape2.Left+5,150,dHeight,'dimOption(4)',0,.T.,'DO validChecKSpis'
     
     DO adCheckBox WITH 'fSupl','checkType','��� ������',.checkGroup.Top,.Shape2.Left+5,150,dHeight,'dimOption(5)',0,.T.,'DO validCheckType' 
     DO adCheckBox WITH 'fSupl','checkDol','���������',.checkPodr.Top,.checkType.Left,150,dHeight,'dimOption(6)',0,.T.,'DO validCheckDol' 
     
     .Shape2.Height=.checkPodr.Height*2+30
     .checkGroup.Left=.Shape2.Left+(.Shape1.Width-.checkGroup.Width-.CheckPers.Width-.checkType.Width-60)/2
     .checkPodr.Left=.checkGroup.Left
     .checkPers.Left=.CheckGroup.Left+.checkGroup.Width+30
     .checkSpis.Left=.checkPers.Left
     .checkType.Left=.checkPers.Left+.checkPers.Width+30
     .checkDol.Left=.checkType.Left
      
     DO addshape WITH 'fSupl',4,.Shape1.Left,.Shape2.Top+.Shape2.Height+10,150,.Shape1.Width,8 
     DO adlabMy WITH 'fSupl',16,' ������������� ',.Shape4.Top-10,.Shape1.Left+10,300,0,.T.,1
     *----------------CheckBox ��� ������ ������ �� ���������� � ������� ��������� ���������---------------------------------
     DO adCheckBox WITH 'fSupl','check1','���� �� ����������',.Shape4.Top+10,.Shape4.Left+(.Shape4.Width-RetTxtWidth('W�������� �������������W')*2-30)/2,200,dHeight,'formbase.log_it',0,.T.  
     *----------------CheckBox �������������� ����������� ��������-----------------------------------------------------------     
     DO adCheckBox WITH 'fSupl','checkAvt','�������� �������������',.check1.Top+.check1.Height+10,.check1.Left,200,dHeight,'formbase.avt_vac',0,.T.   
     *----------------CheckBox ��� ��������� �� ����� ������-----------------------------------------------------------     
     DO adCheckBox WITH 'fSupl','checkSt','��������� �� ����� 1.00',.checkAvt.Top+.checkAvt.Height+10,.check1.Left,200,dHeight,'formbase.vacst',0,.T.  
     *----------------CheckBox ��� ������ �� ���� ������-----------------------------------------------------------     
     DO adCheckBox WITH 'fSupl','checkItr','���� �� ���� ������',.checkSt.Top+.checkSt.Height+10,.check1.Left,200,dHeight,'formbase.litr',0,.T.           
     *-----------------------------CheckBox ��� ��������� (����������)�������� � ���������------------------------------------  
     DO adCheckBox WITH 'fSupl','check5','��������� ���������',.check1.Top,.check1.Left+RetTxtWidth('W�������� �������������W')+10,200,dHeight,'excludeVac',0,.T.         
     *-----------------------------CheckBox ��� ������ ���������----------------------------------------------------------------------------  
     DO adCheckBox WITH 'fSupl','checkOnlyVac','������ ���������',.checkAvt.Top,.check5.Left,200,dHeight,'onlyVac',0,.T.  
     *-----------------------------CheckBox ��� ��������� (����������)��������������� ������� � ���������------------------------------------  
     DO adCheckBox WITH 'fSupl','check7','��������',.checkSt.Top,.check5.Left,200,dHeight,'formbase.log_kom',0,.T.           
     .Shape4.Height=dHeight*4+50   
        
         
     DO adSetupPrnToForm WITH .Shape2.Left,.Shape4.Top+.Shape4.Height+10,.Shape1.Width,.F.,.T.
     DO adTBoxAsCont WITH 'fSupl','txtMarg',.Shape91.Left,.Shape91.Top+.Shape91.Height+10,RetTxtWidth('W��������� "�����"'),dHeight,'��������� "�����"',2,1
     DO addSpinnerMy WITH 'fSupl','spinMarg',.txtMarg.Left+.txtMarg.Width-1,.txtMarg.Top,dheight,RetTxtWidth('999999999'),'margTop',0.5,.F.,-3,15
     .txtMarg.Left=.Shape91.Left+(.Shape91.Width-.txtMarg.Width-.spinMarg.Width-1)/2
     .spinMarg.Left=.txtMarg.Left+.txtMarg.Width-1
     
     DO adTBoxAsCont WITH 'fSupl','txtSpace',.Shape91.Left,.txtMarg.Top,RetTxtWidth('W����������� ��������'),dHeight,'����������� ��������',2,1
     DO addSpinnerMy WITH 'fSupl','spinSpace',.txtSpace.Left+.txtSpace.Width-1,.txtSpace.Top,dheight,RetTxtWidth('999999999'),'nlSpace',0.1,.F.,-1,2
     .txtMarg.Left=.Shape91.Left+(.Shape91.Width-.txtMarg.Width-.spinMarg.Width-.txtSpace.Width-.spinSpace.Width-12)/2
     .spinMarg.Left=.txtMarg.Left+.txtMarg.Width-1
     .txtSpace.Left=.spinMarg.Left+.spinMarg.Width+10
     .spinSpace.Left=.txtSpace.Left+.txtSpace.Width-1
     
    
     DO adButtonPrnToForm WITH 'DO vedprn WITH 2','DO vedprn WITH 1','Do exitprn',.T.,'fSupl'
     .butPrn.Top=.txtSpace.Top+.txtSpace.Height+15
     .butView.Top=.butPrn.Top
     .butRet.Top=.butPrn.Top
     .cont11.Top=.butPrn.Top
     .cont12.Top=.butPrn.Top                               
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width
     
      DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.butPrn.Top-.Shape1.Top-15,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH   
           
     
     .Width=.Shape1.Width+20
     .Height=.butPrn.Top+.butPrn.height+10
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************
PROCEDURE exitPrn
SELECT dopDol
USE
SELECT dopPodr
USE
SELECT dopKat
USE
SELECT dopType
USE
SELECT dopGroup
USE
IF USED('fondprn')
   SELECT fondprn
   USE
ENDIF 
SELECT people
IF !EMPTY(fltjob)
   SET FILTER TO SEEK(num,'curFltDatjob',1)
ENDIF
ON ERROR DO erSup
IF fltRec#0
   GO fltRec
ENDIF
ON ERROR 
frmTop.Refresh  
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus           
fSupl.Release
******************************************************************************************************************************
PROCEDURE validformbase
WITH fSupl
     wordOption=.F.
     nameform=.comboBox1.Value
     .comboBox2.Enabled=IIF(formbase.logCmb,.T.,.F.)
     .boxBeg.Enabled=IIF(formbase.logTime,.T.,.F.)
     .boxEnd.Enabled=IIF(formbase.logTime,.T.,.F.)
     .Refresh
ENDWITH 
******************************************************************************************************************************
PROCEDURE validformnadbav
persnadbav=ALLTRIM(nadbase.persved)
sumnadbav=ALLTRIM(nadbase.sumved)
persFormula=ALLTRIM(nadbase.countvac)
headNadbav=ALLTRIM(nadbase.nHead)
nsupproc=nadbase.supproc
persKse=nadBase.logKse
sumOklSt=ALLTRIM(nadBase.oklSt)
*************************************************************************************************************************
PROCEDURE validCheckGroup
dimOption(2)=.F.
dimOption(4)=.F.
dimOption(1)=.T. 
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')  
     .SetAll('Visible',.F.,'MyCommandButton')    
     .cont11.Visible=.T.
     .cont12.Visible=.T. 
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopGroup.otm,name'  
     .listBox1.procForClick='DO clickListGroup'
     .listBox1.procForKeyPress='DO KeyPressListGroup' 
     .cont11.procForClick='DO returnToPrnGroup WITH .T.'
     .cont12.procForClick='DO returnToPrnGroup WITH .F.'
     .checkPodr.Caption='�������������'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListGroup
SELECT dopGroup
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' � ','')
nameGroup=IIF(fl,dopGroup.name,'')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
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
      dimOption(1)=.T.
      nameGroup=dopGroup.name
   ELSE 
      dimOption(1)=.F.
      nameGroup=''
   ENDIF  
ELSE 
   nameGroup=''
   SELECT dopGroup
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
   dimOption(1)=.F.
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel') 
     .SetAll('Visible',.T.,'MyCommandButton')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.  
     .listBox1.Visible=.F.
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE validCheckPodr
dimOption(1)=.F. 
dimOption(2)=.T.
dimOption(4)=.F.
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T. 
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopPodr.otm,name'  
     .listBox1.procForClick='DO clickListPodr'
     .listBox1.procForKeyPress='DO KeyPressListPodr' 
     .cont11.procForClick='DO returnToPrnPodr WITH .T.'
     .cont12.procForClick='DO returnToPrnPodr WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListPodr
SELECT dopPodr
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' � ','')
GO rrec
fSupl.listBox1.SetFocus
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
    dimOption(2)=.F.
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.   
     .listBox1.Visible=.F.  
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
     dimOption(2)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='�������������'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE validCheckDol
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopDol.otm,name'  
     .listBox1.procForClick='DO clickListDol'
     .listBox1.procForKeyPress='DO KeyPressListDol' 
     .cont11.procForClick='DO returnToPrnDol WITH .T.'
     .cont12.procForClick='DO returnToPrnDol WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListDol
SELECT dopDol
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' � ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListDol
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListDol 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnDol
PARAMETERS parRet
kvoDol=0
IF parRet
   SELECT dopDol
   fltDol=''
   onlyDol=.F.
   SCAN ALL
        IF fl 
           fltDol=fltDol+','+LTRIM(STR(kod))+','
           onlyDol=.T.
           kvoDol=kvoDol+1
        ENDIF 
   ENDSCAN
ELSE 
   strDol=''
   fltDol=''
   onlyDol=.F.
   SELECT dopDol
   REPLACE otm WITH '',fl WITH .F. ALL
   dimOption(6)=.F.
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')
     .SetAll('Visible',.T.,'MyCommandButton')      
     .cont11.Visible=.F.
     .cont12.Visible=.F.    
     .listBox1.Visible=.F.  
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
     dimOption(6)=IIF(kvoDol>0,.T.,.F.)
     .checkDol.Caption='���������'+IIF(kvoDol#0,'('+LTRIM(STR(kvoDol))+')','') 
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE validCheckPers
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel') 
     .SetAll('Visible',.F.,'MyCommandButton')        
     .cont11.Visible=.T.
     .cont12.Visible=.T.    
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopKat.otm,name'  
     .listBox1.procForClick='DO clickListPers'
     .listBox1.procForKeyPress='DO KeyPressListPers' 
     .cont11.procForClick='DO returnToPrnPers WITH .T.'
     .cont12.procForClick='DO returnToPrnPers WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListPers
SELECT dopKat
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' � ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListPers
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListPers 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnPers
PARAMETERS parRet
kvoKat=0
IF parRet
   SELECT dopKat
   fltKat=''
   SCAN ALL
        IF fl 
           fltKat=fltKat+','+LTRIM(STR(kod))+','        
           kvoKat=kvoKat+1
        ENDIF 
   ENDSCAN  
ELSE  
   fltKat=''
   strPodr=''
   onlyPodr=.F.
   SELECT dopKat
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')  
     .SetAll('Visible',.T.,'MyCommandButton')     
     .cont11.Visible=.F.
     .cont12.Visible=.F.    
     .listBox1.Visible=.F.  
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
      dimOption(3)=IIF(kvokat=0,.F.,.T.)
     .checkPers.Caption='��������'+IIF(kvoKat#0,'('+LTRIM(STR(kvoKat))+')','') 
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE validCheckSpis
dimOption(1)=.F.
dimOption(2)=.F.
dimOption(3)=.F.
dimOption(4)=.T. 
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')      
     .SetAll('Visible',.F.,'MyCommandButton')    
     .cont11.Visible=.T.
     .cont12.Visible=.T.   
     .listBox1.Visible=.T.
     .listBox1.RowSource='curSpis.otm,ksefio'  
     .listBox1.procForClick='DO clickListSpis'
     .listBox1.procForKeyPress='DO KeyPressListSpis' 
     .cont11.procForClick='DO returnToPrnSpis WITH .T.'
     .cont12.procForClick='DO returnToPrnSpis WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE returnToPrnSpis
PARAMETERS parRet
kvoSpis=0
IF parRet
   SELECT curSpis
   fltPodr=''
   onlyPodr=.F.
   SCAN ALL
        IF fl 
           fltPodr=fltPodr+','+LTRIM(STR(kodpeop))+','
           onlyPodr=.T.
           kvoSpis=kvoSpis+1
        ENDIF 
   ENDSCAN
ELSE 
   strPodr=''
   onlyPodr=.F.
   SELECT curSpis
   REPLACE otm WITH '',fl WITH .F. ALL
    dimOption(4)=.F.
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')  
     .SetAll('Visible',.T.,'MyCommandButton')    
     .cont11.Visible=.F.
     .cont12.Visible=.F.   
     .listBox1.Visible=.F.  
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
     dimOption(4)=IIF(kvoSpis#0,.T.,.F.)
     .checkSpis.Caption='������'+IIF(kvoSpis#0,'('+LTRIM(STR(kvoSpis))+')','') 
     .Refresh
ENDWITH 
*************************************************************************************************************************
PROCEDURE clickListSpis
SELECT curSpis
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' � ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListSpis
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListSpis
ENDCASE   
*************************************************************************************************************************
PROCEDURE validCheckType
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'MyCheckBox')
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel') 
     .SetAll('Visible',.F.,'MyCommandButton')        
     .cont11.Visible=.T.
     .cont12.Visible=.T.   
     .listBox1.Visible=.T.
     .listBox1.RowSource='dopType.otm,name'  
     .listBox1.procForClick='DO clickListType'
     .listBox1.procForKeyPress='DO KeyPressListType' 
     .cont11.procForClick='DO returnToPrnType WITH .T.'
     .cont12.procForClick='DO returnToPrnType WITH .F.'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE clickListType
SELECT dopType
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' � ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
*************************************************************************************************************************
PROCEDURE keyPressListType
DO CASE
   CASE LASTKEY()=27
        *DO returnFromFltPodr WITH 'curFltPodr','name'
   CASE LASTKEY()=13
        Do clickListType 
ENDCASE   
************************************************************************************************************************
PROCEDURE returnToPrnType
PARAMETERS parRet
kvoType=0
IF parRet
   SELECT dopType
   fltType=''
   SCAN ALL
        IF fl 
           fltType=fltType+','+LTRIM(STR(kod))+','
           *onlyPodr=.T.
           kvoType=kvoType+1
        ENDIF 
   ENDSCAN
   dimOption(5)=IIF(kvoType=0,.F.,.T.)
ELSE 
   dimOption(5)=.F.
   fltType=''
   strPodr=''
   onlyPodr=.F.
   SELECT dopType
   REPLACE otm WITH '',fl WITH .F. ALL
   GO TOP
ENDIF 
WITH fSupl
     .SetAll('Visible',.T.,'LabelMy')
     .SetAll('Visible',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox')
     .SetAll('Visible',.T.,'comboMy')
     .SetAll('Visible',.T.,'shapeMy')
     .SetAll('Visible',.T.,'MyOptionButton')
     .SetAll('Visible',.T.,'MySpinner')
     .SetAll('Visible',.T.,'MyContLabel')  
     .SetAll('Visible',.T.,'MyCommandButton')     
     .cont11.Visible=.F.
     .cont12.Visible=.F.  
     .listBox1.Visible=.F.  
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.  
     .lab25.Visible=.F.
     .checkType.Caption='��� ������'+IIF(kvoType#0,'('+LTRIM(STR(kvoType))+')','') 
     .Refresh
ENDWITH 

*******************************************************************************************************************************************************
*                  ��������������� ��������� ��� �������� ������ ��� ������ � ��������� �����������
*******************************************************************************************************************************************************
PROCEDURE procForPrintAndPreview
PARAMETERS parreport,par_caption
IF logWord.AND.terminal=2 
   IF !EMPTY(formbase.procsupl)
       procToExcel=ALLTRIM(formbase.procsupl)
       DO &procToExcel
   ENDIF   
ELSE 
   IF terminal=1  
      DO previewRep WITH parreport,par_caption
   ELSE     
      SET PRINTER TO NAME name_prn(ASCAN(name_prn,nameprint))       
      DO CASE
         CASE dimCht(1)=1 
              FOR ch=1 TO kvo_page
                  Report Form &parreport RANGE page_beg, page_end NOCONSOLE TO PRINTER                   
              ENDFOR   
         CASE dimCht(2)=1                                       
              FOR ch=1 TO kvo_page    
                  FOR c_range=page_beg TO page_end
                      IF MOD(c_range,2)=0
                         Report Form &parreport RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ENDIF  
                      IF EOF()
                         EXIT 
                      ENDIF 
                  ENDFOR    
              ENDFOR    
         CASE dimCht(3)=1
              FOR ch=1 TO kvo_page
                  FOR c_range=page_beg TO page_end         
                      IF MOD(c_range,2)#0
                         Report Form &parreport RANGE c_range,c_range NOCONSOLE TO PRINTER   
                      ENDIF  
                      IF EOF()
                         EXIT 
                      ENDIF  
                  ENDFOR              
              ENDFOR    
      ENDCASE       
   ENDIF    
ENDIF   
*************************************************************************************************************************
PROCEDURE fltStructure
PARAMETERS parField,parBase
IF !EMPTY(parBase) 
   SELECT &parBase
ELSE    
   SELECT curTarJob
ENDIF    
SET FILTER TO 
IF !EMPTY(parField)
   IF excludeVac
   ENDIF
   
ELSE
   IF excludeVac
      parField='!vac'
   ENDIF
ENDIF
DO CASE
   CASE dimOption(1)       && ������  ������         
        fltch=ALLTRIM(dopGroup.sostav1)+',999,'     
        DO CASE        
           CASE !EMPTY(parField)               
                SET FILTER TO ','+LTRIM(STR(kp))+','$fltch.AND.&parField  
           CASE EMPTY(parField)
                SET FILTER TO ','+LTRIM(STR(kp))+','$fltch  
        ENDCASE               
   CASE dimOption(2)  && �������������   
        fltpodr=fltpodr+',999,'   
        DO CASE        
           CASE !EMPTY(parField)
                SET FILTER TO ','+LTRIM(STR(kp))+','$fltpodr.AND.&parField                  
              * xcbzxc
          CASE EMPTY(parField)    
               SET FILTER TO ','+LTRIM(STR(kp))+','$fltpodr  
            *   asdfzdxv
        ENDCASE               
       
*   CASE dimOption(3)  
*        DO procflt_str  
*        IF !EMPTY(parbase)
*           SELECT &parbase                              
*        ENDIF   
*        SET FILTER TO ','+LTRIM(STR(kp))+','$flt_str.AND.kp>0.AND.kd>0.AND.&par_field              
   CASE dimOption(4)    
        spisnum=''
        SELECT curSpis
        SCAN ALL
            spisnum=IIF(fl,spisnum+','+LTRIM(STR(kodpeop)),spisnum)
        ENDSCAN
        spisnum=spisnum+','
        IF !EMPTY(parBase) 
           SELECT &parBase
        ELSE    
           SELECT curTarJob
        ENDIF    
        GO TOP 
        SET FILTER TO ','+LTRIM(STR(kodpeop))+','$spisnum
    OTHERWISE
        IF !EMPTY(parField)       
           SET FILTER TO &parField
        ELSE 
        ENDIF  
ENDCASE
IF !EMPTY(fltKat)
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltkat)
ENDIF
IF !EMPTY(fltType)
   DELETE FOR !(','+LTRIM(STR(tr))+','$fltType)
ENDIF
IF !EMPTY(fltDol)
   DELETE FOR !(','+LTRIM(STR(kd))+','$fltDol)
ENDIF
*************************************************************************************************************************
*   ��������������� ��������� ��� ������������� ������� (currasp) � ������������ �������� � people
*************************************************************************************************************************
PROCEDURE selectfromrasp
SELECT rasp
IF USED('currasp')
   SELECT currasp
   USE
ENDIF      
SELECT * FROM rasp WHERE SEEK(STR(kp,3)+STR(kd,3),'curTarJob',2) INTO CURSOR currasp READWRITE 
*ALTER TABLE currasp ADD COLUMN np N(3)
*REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL     
SELECT currasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1
GO TOP
kp_old=kp
num_new=1
DO WHILE !EOF()
   REPLACE nd WITH num_new
   num_new=num_new+1
   SKIP
   IF kp#kp_old    
      kp_old=kp 
      num_new=1
   ENDIF
ENDDO
SELECT currasp
SET ORDER TO 2
SELECT curTarJob
ord_old=SYS(21)
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curTarJob.kp,3)+STR(curTarJob.kd,3)   
   SELECT curTarJob
   REPLACE nd WITH currasp->nd,np WITH currasp->np  
   SKIP 
ENDDO 
SELECT currasp
SET ORDER TO 1
GO TOP 
SELECT curTarJob
GO TOP 
**************************************************************************************************************************
*                       ����������� ������� � ������ ��� �����������
**************************************************************************************************************************
PROCEDURE repNadJob 
SELECT rasp
oldOrdRasp=SYS(21)
SET ORDER TO 2
SEEK STR(datjob.kp,3)+STR(datjob.kd,3)
SET ORDER TO &oldOrdRasp
SELECT tarfond
GO TOP
SCAN ALL
     IF !EMPTY(plrep).AND.ltar
        repjob=ALLTRIM(plrep)
        repjob1='rasp.'+ALLTRIM(plrep)
        SELECT datjob 
        REPLACE &repjob WITH &repjob1
     ENDIF
     IF ALLTRIM(LOWER(tarfond.plrep))='pkat'.AND.rasp.pkat#0.AND.datjob.kv=0
        SELECT datjob
        REPLACE pkat WITH 5       
     ENDIF
     SELECT tarfond
ENDSCAN
SELECT people
IF !EMPTY(dmol).AND.dmol<dateTar
   REPLACE dmol WITH CTOD('  .  .    ' )
   SELECT datjob
   REPLACE pmols WITH 0
ENDIF
SELECT datjob
repVr='vr'+LTRIM(STR(MONTH(dateTar)))
REPLACE patt WITH pkfvr,satt WITH &repVr,matt WITH &repVr
IF EMPTY(date_in)
   REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in 
ENDIF
DO actualstaj
DO perstajone
SELECT tarfond
SET FILTER TO 
totsumf=0
totsumfm=0
totfondprn=0
SCAN ALL
     IF !EMPTY(formula)
        new_sum=sum_f
        new_msum=0
        pole=fname                 
        r_sum=ALLTRIM(tarfond.formula)  
        r_sum1=ALLTRIM(tarfond.formula1) 
        SELECT datJob 
        DO CASE
            CASE !EMPTY(tarfond.proccount)
                 procForCount=ALLTRIM(tarfond.proccount)                           
                 DO &procForCount
            CASE !EMPTY(tarfond.formula)
                 IF !EMPTY(tarfond.sum_f)                    
                    REPLACE &new_sum WITH &r_sum 
                    IF !EMPTY(tarfond.sum_fm)                  
                       new_msum=tarfond.sum_fm                  
                       REPLACE &new_msum WITH IIF(tarfond.logkse,&new_sum*datjob.kse,&new_sum)  
                       IF !EMPTY(tarfond.formula1)
                          REPLACE &new_sum WITH &r_sum1  
                       ENDIF
                    ENDIF
                 ELSE
                    SELECT datjob
                    REPLACE &pole WITH &r_sum
                 ENDIF                                
           ENDCASE                 
     ENDIF  
     SELECT datjob   
     totsumf=IIF(!EMPTY(tarfond.sum_f),totsumf+EVALUATE(ALLTRIM(tarfond.sum_f)),totsumf)
     totfondprn=IIF(tarfond.logfprn,totfondprn+EVALUATE(ALLTRIM(tarfond.sum_fm)),totfondprn)
     totsumfm=IIF(!EMPTY(tarfond.sum_fm),totsumfm+EVALUATE(ALLTRIM(tarfond.sum_fm)),totsumfm)
     
     IF tarfond.logit
        REPLACE total WITH totsumf,msf WITH totsumfm,fdprn WITH totfondprn
     ENDIF
     SELECT tarfond            
ENDSCAN
GO TOP

************************************************************************************************************************
PROCEDURE tarifprn
PARAMETERS parType
itkat=formbase.log_it
logKom=formBase.log_kom
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF  
IF USED('curTarJob')
   SELECT curTarJob
   USE
ENDIF   

IF USED('curFondPrn')
   SELECT curFondPrn
   USE
ENDIF    
SELECT npved,nameved,persved,sumved FROM tarfond WHERE !EMPTY(nameved).AND.!EMPTY(sumved) INTO CURSOR curFondPrn READWRITE

SELECT curFondPrn
INDEX ON npved TAG T1

SELECT * FROM tarfond WHERE tarfond.vac.AND.!EMPTY(persved) INTO CURSOR curPrnTarFond READWRITE 
SELECT curPrnTarFond
INDEX ON num TAG T1
GO TOP
num_cx=0
DO WHILE !EOF()
   num_cx=num_cx+1
   REPLACE num WITH num_cx
   SKIP   
ENDDO

SELECT sprkat
* maxkat=���-�� ��������� ���������
* max_Tr=���-�� ����� ������
* sum_podr - ���� �� ���������� 
* sumPodrKat - ���� �� ���������� � ������� ���������

* sumTot - ���� �����
* sumTotTr -���� ����� �� ����� ������
* sumTotKat - ���� ����� �� ���������� ���������
* sumKpp -���� �� ����������������

COUNT TO maxKat
DIMENSION nSum(1)
STORE 0 TO maxItog
IF parType=3
   SELECT curFondPrn
   COUNT TO maxItog
   DIMENSION nSum(maxItog)
   GO TOP 
   FOR i=1 TO maxitog
       nSum(i)=ALLTRIM(sumved)
       SKIP 
   ENDFOR 
ELSE
   DO setupreport
ENDIF
DIMENSION sumTotTr(max_tr,maxItog)
DIMENSION sum_podr(maxItog),sumTot(maxItog)
DIMENSION sumPodrKat(maxKat,maxItog),sumTotKat(maxKat,maxItog)
STORE 0 TO sum_podr,sumTot,sumPodrKat,sumTotkat

STORE 0 TO sumTotTr  &&����� �� ���� ������ �����

DIMENSION sumPodrKpp(maxItog),sumPodrKpp1(maxItog)  &&����� � ����������������� � ������ ��������������
STORE 0 TO sumPodrKpp,sumPodrKpp1

DIMENSION sumPodrKatKpp(maxKat,maxItog),sumPodrKatKpp1(maxKat,maxItog)  &&����� � ����������������� � ������ �������������� �� ���������� ���������
STORE 0 TO sumPodrKatKpp,sumPodrKatKpp1

*������� ��� ������ ���� ������ � ������� ��������� ���������
FOR i=1 TO maxKat
    ntdim='dimkattr'+LTRIM(STR(i))+'('+LTRIM(STR(max_tr))+','+LTRIM(STR(maxItog))+')'
    ntstore='dimkattr'+LTRIM(STR(i))
    DIMENSION &ntdim
    STORE 0 TO &ntstore   
ENDFOR 
IF datShtat.Real.AND.parType=4
   SELECT people
   oldOrd=SYS( 21)
   SET ORDER TO 1
   SELECT datJob
   SET FILTER TO     
   REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,kat) FOR kat=0   
   SCAN ALL            
        SELECT people 
        SEEK datjob.kodpeop
        SELECT datjob
        REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in,pkont WITH IIF(tr=1,people.pkont,0),dateBeg WITH IIF(EMPTY(dateBeg),date_in,dateBeg)
        DO CASE 
           CASE datjob.lkv.AND.people.kval#0
                REPLACE kv WITH people.kval,nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' �'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' �� ' +DTOC(people.dkval),''),;
                         pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,0)
           OTHERWISE
             REPLACE kv WITH 0,nPrik WITH '',pkat WITH 0   
        ENDCASE
  
        DO CASE
           CASE kv=0
                REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf,namekf)
           CASE kv=1
                REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf3,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf3,namekf)
           CASE kv=2
                REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf2,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf2,namekf)
           CASE kv=3
                REPLACE kf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.kf1,kf),namekf WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namekf1,namekf)
        ENDCASE     
        *IF KODPEOP=638
        *   ZXVZXCV
        *ENDIF  
        *DO repNadJob 
        SELECT datJob    
   ENDSCAN
   SELECT people
   SET ORDER TO &oldOrd
ENDIF
SELECT * FROM datJob INTO CURSOR curTarJob READWRITE
SELECT curTarJob
IF datShtat.Real
   SELECT curTarJob  
   DELETE FOR tr=4                  && ������� ������ �� ���������� tr=4
   DELETE FOR kd=0.OR.kp=0
   DELETE FOR SEEK(kodpeop,'people',1).AND.people.dekotp.AND.people.bdekotp<=dateTar && ������� �����������   
    
   REPLACE dateuv WITH IIF(SEEK(kodpeop,'people',1),people.date_out,dateuv) ALL
   REPLACE dekotp WITH IIF(SEEK(kodpeop,'people',1),people.dekotp,.F.) ALL
   REPLACE date_in WITH IIF(SEEK(kodpeop,'people',1),people.date_in,date_in) ALL
   IF parType#4 
      DELETE FOR SEEK(kodpeop,'people',1).AND.!EMPTY(people.date_out).AND.people.date_out<dateTar    && ������� ������ ��� ���� ���������� ������ ���� �����������   
   ENDIF     
ENDIF 
*DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)

*   DELETE FOR SEEK(peopledate_in>newDateTar    && ������� ������ ��� ���� ������ ������ ���� �����������
*   DELETE FOR SEEK(kodpeop,'people',1).AND.people.dekotp

   *DELETE FOR dateBeg>newDateTar    && ������� ������ ��� ���� ������ ������ ������ ���� �����������
DO CASE   
   CASE parType=4
   OTHERWISE 
        DELETE FOR !EMPTY(dateOut).AND.EMPTY(dateuv).AND.dateOut<=dateTar    && ������� ������ ��� ���� ���������<=������ ������ ���� ����������� ��� ������ ���� ����������
        DELETE FOR !EMPTY(dateuv).AND.dateuv<dateTar  &&���� ����������<���� ����������� ��� ����������� ���� ����������
ENDCASE  

IF parType#4
   DELETE FOR date_in>dateTar && ������� ��������� ����� ���� �-���
ELSE 
   DELETE FOR !INLIST(tr,1,3)
   REPLACE dateuv WITH IIF(SEEK(kodpeop,'people',1),people.date_out,dateuv) ALL
   DELETE FOR dateBeg<begTar    && ������� ������ ��� ���� ������ ������ ������ ���� �����������
   DELETE FOR dateBeg>endTar    && ������� ������ ��� ���� ������ ������ ������ ���� ����������� 
   DELETE FOR !EMPTY(dateOut).AND.EMPTY(dateuv)   && ������� ������ ��� ���� ���������<=������ ������ ���� ����������� ��� ������ ���� ����������
   DELETE FOR !EMPTY(dateOut).AND.!EMPTY(dateuv).AND.dateuv<endTar   && ������� ������ ��� ���� ���������<=������ ������ ���� ����������� ��� ������ ���� ����������   
ENDIF    
ALTER TABLE curTarJob ADD COLUMN npp N(3)
ALTER TABLE curTarJob ADD COLUMN nit N(1)
ALTER TABLE curTarJob ADD COLUMN nkat C(200)
ALTER TABLE curTarJob ADD COLUMN dmol D
ALTER TABLE curTarJob ADD COLUMN sex N(1)
SELECT curTarJob
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,''),sex WITH people.sex ALL 
REPLACE dmol WITH IIF(SEEK(kodpeop,'people',1),people.dmol,dmol) ALL
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.np,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL
DELETE FOR !SEEK(kp,'sprpodr',1)
INDEX ON STR(np,3)+STR(ND,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 2
*----------------------------------- �������������� ���������� �������� ---------------------------------------------
IF formbase.avt_vac.AND.parType#4
   SELECT datjob
   ordOld=SYS(21)
   SET FILTER TO 
   SELECT rasp
   GO TOP
   DO WHILE !EOF()
      IF rasp.kse#0
         SELECT datjob
         SET ORDER TO 2
         SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
         kse_cx=rasp.kse          
         DO WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()        
            IF date_in>dateTar            
            ELSE 
               kse_cx=kse_cx-IIF(datjob.tr=4.OR.datjob.dekotp,0,datjob.kse)
            ENDIF  
            SKIP
         ENDDO                 
         IF kse_cx>0
            IF !formbase.vacst
               SELECT curTarJob
               APPEND BLANK
               REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,kse WITH kse_cx,vac WITH .T.,tr WITH 1
               SELECT curPrnTarFond
               GO TOP
               DO WHILE !EOF()
                  rep_r=ALLTRIM(persved)
                  rep_r1='rasp.'+ALLTRIM(persved)
                  SELECT curTarJob 
                  REPLACE &rep_r WITH &rep_r1         
                  SELECT curPrnTarFond
                  SKIP
               ENDDO                                               
               SELECT curTarJob            
               DO countOkladVac             
            ELSE 
               DO CASE
                  CASE kse_cx<=1
                       kvovac=1
                  CASE MOD(kse_cx,1)=0     
                       kvovac=INT(kse_cx)
                  CASE MOD(kse_cx,1)>0     
                       kvovac=INT(kse_cx)+1    
               ENDCASE               
               kvokse=kse_cx
               ksevac=0
             
               FOR i=1 TO kvovac
                   ksevac=IIF(kvokse<=1,kvokse,1)
                   SELECT curTarJob
                   APPEND BLANK
                   REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,kse WITH ksevac,vac WITH .T.,tr WITH 1
                   SELECT curPrnTarFond
                   GO TOP
                   DO WHILE !EOF()                   
                      rep_r=ALLTRIM(persved)
                      rep_r1='rasp.'+ALLTRIM(persved)
                      SELECT curTarJob 
                      REPLACE &rep_r WITH &rep_r1         
                      SELECT curPrnTarFond
                      SKIP
                   ENDDO                                                                
                   SELECT curTarJob                             
                   DO countOkladVac
                   kvokse=kvokse-1
               ENDFOR
            ENDIF    
         ENDIF       
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
   SELECT datJob
   SET ORDER TO &ordOld   
ENDIF    
*--------------------------------------------------------------------------------------------------------------------

=AFIELDS(arJob,'curTarJob')
CREATE CURSOR curPrn FROM ARRAY arJob
ALTER TABLE curPrn ADD COLUMN primhead C(150)
ALTER TABLE curPrn ALTER COLUMN kse N(8,2)
ALTER TABLE curPrn ADD COLUMN nvac N(1)
ALTER TABLE curPrn ALTER COLUMN fio C(185)
ALTER TABLE curPrn ADD COLUMN namedol C(150)
IF parType=4
   DO fltstructure WITH 'kse#0'
ELSE 
   DO fltstructure WITH 'mTokl#0'
ENDIF 
IF onlyVac
   SELECT curTarJob
   DELETE FOR !vac
ENDIF

IF excludeVac
   SELECT curTarJob
   DELETE FOR vac
ENDIF
*---������ ��� �������� ����������
IF USED('curRasp')
   SELECT curRasp
   USE
ENDIF 
SELECT rasp      
SELECT * FROM rasp WHERE SEEK(STR(kp,3)+STR(kd,3),'curTarJob',2) INTO CURSOR currasp READWRITE      
SELECT currasp
SCAN ALL
    IF SEEK(curRasp.kp,'sprpodr',1)
        REPLACE np WITH sprpodr.np,kpp WITH sprpodr.kpp,kpp1 WITH sprpodr.kpp1,kodkpp WITH sprpodr.kodkpp,kodkpp1 WITH sprpodr.kodkpp1  
    ENDIF   
ENDSCAN     
SELECT sprpodr
SCAN ALL
     IF sprpodr.kodKpp#0
        SELECT curRasp
        REPLACE kodkpp WITH sprpodr.kodkpp FOR kp=sprpodr.kodkpp
     ENDIF 
     IF sprpodr.kodKpp1#0
        SELECT curRasp
        REPLACE kodkpp1 WITH sprpodr.kodkpp1,kodkpp WITH sprpodr.kodkpp FOR kp=sprpodr.kodkpp1
        
     ENDIF 
     SELECT sprpodr
ENDSCAN   
SELECT curRasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1
GO TOP
kp_old=kp
num_new=1
DO WHILE !EOF()
   REPLACE nd WITH num_new
   IF nd=1 
      SELECT rasp
      LOCATE FOR kp=currasp.kp.AND.nd=1
      SELECT currasp
      REPLACE primhead WITH rasp.primhead
   ENDIF
   num_new=num_new+1
   SKIP
   IF kp#kp_old    
      kp_old=kp 
      num_new=1
   ENDIF
ENDDO

SELECT currasp
SET ORDER TO 2
SELECT curTarJob
ord_old=SYS(21)
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curTarJob.kp,3)+STR(curTarJob.kd,3)   
   SELECT curTarJob
   REPLACE nd WITH currasp.nd,np WITH currasp.np,kpp WITH currasp.kpp,kodkpp WITH currasp.kodkpp,kpp1 WITH currasp.kpp1,kodkpp1 WITH currasp.kodkpp1,;
   kat WITH IIF(currasp.kat#curTarJob.kat,currasp.kat,curTarJob.kat) 
   SKIP 
ENDDO 
SELECT curTarJob
SET ORDER TO 1
GO TOP
kp_cx=kp
kpp_cx=kodkpp
kpp1_cx=kodkpp1
npp_cx=0
DO WHILE !EOF()     
   SCATTER TO ac
   SELECT curPrn
   npp_cx=npp_cx+1
   APPEND BLANK
   REPLACE npp WITH npp_cx
   GATHER FROM ac                        
   log_sumdopl=.F.     
   FOR i=1 TO maxItog
       repSum=nSum(i)
       sum_podr(i)=sum_podr(i)+&repSum  
       sumTot(i)=sumTot(i)+&repsum   
       sumPodrKpp(i)=IIF(kpp_cx#0,sumPodrKpp(i)+&repSum,sumPodrKpp(i))        
       sumPodrKpp1(i)=IIF(kpp1_cx#0,sumPodrKpp1(i)+&repSum,sumPodrKpp1(i)) 
       
       IF tr#0
          sumTotTr(ASCAN(kod_tr,curprn.tr),i)=sumTotTr(ASCAN(kod_tr,curPrn.tr),i)+&repSum 
          
          IF kat#0
             ntdim='dimkattr'+LTRIM(STR(ASCAN(kod_kat,curPrn.kat)))+'('+LTRIM(STR(ASCAN(kod_tr,curPrn.tr)))+','+LTRIM(STR(i))+')'
             &ntdim=&ntdim+&repSum          
          ENDIF                     
       ENDIF   
       IF kat#0
          sumTotKat(ASCAN(kod_kat,curPrn.kat),i)=sumTotKat(ASCAN(kod_kat,curPrn.kat),i)+&repSum 
          sumPodrKat(ASCAN(kod_kat,curPrn.kat),i)=sumPodrKat(ASCAN(kod_kat,curPrn.kat),i)+&repSum        
          
          sumPodrKatKpp(ASCAN(kod_kat,curPrn.kat),i)=IIF(kpp_cx#0,sumPodrKatKpp(ASCAN(kod_kat,curPrn.kat),i)+&repSum,sumPodrKatKpp(ASCAN(kod_kat,curPrn.kat),i))        
          sumPodrKatKpp1(ASCAN(kod_kat,curPrn.kat),i)=IIF(kpp1_cx#0,sumPodrKatKpp1(ASCAN(kod_kat,curPrn.kat),i)+&repSum,sumPodrKatKpp1(ASCAN(kod_kat,curPrn.kat),i))        
       ENDIF  
   ENDFOR  
   
   SELECT curTarJob   
   SKIP
   IF kp_cx#curTarJob.kp   
      npp_cx=0         
      SELECT curPrn
      APPEND BLANK      
      REPLACE nIt WITH 1,kp WITH kp_cx,kd WITH 999,fio WITH '�����',nd WITH 700
      FOR i=1 TO maxItog
         repSum=nSum(i)	
         REPLACE &repSum WITH sum_podr(i)
      ENDFOR
      &&  ���� �� ���������� �������� � ���������
      IF itkat
         FOR i=1 TO maxKat
             APPEND BLANK
              REPLACE nIt WITH 2,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),nd WITH 700+i
              FOR h=1 TO maxItog
                  repSum=nSum(h)	
                  REPLACE &repSum WITH sumPodrKat(i,h)
              ENDFOR        
         ENDFOR              
      ENDIF
      
      IF kpp1_cx#curTarJob.kodkpp1.AND.kpp1_cx#0.AND.partype#4
         &&  ���� �� ���������� �������� � ���-����������������
         SELECT curprn
         APPEND BLANK      
         REPLACE nIt WITH 3,kp WITH kp_cx,kd WITH 999,fio WITH '�����',nd WITH 900
         FOR i=1 TO maxItog
             repSum=nSum(i)	
             REPLACE &repSum WITH sumPodrKpp1(i)
         ENDFOR
         IF itkat
            FOR i=1 TO maxKat
                APPEND BLANK
                REPLACE nIt WITH 2,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),nd WITH 900+i
                FOR h=1 TO maxItog
                    repSum=nSum(h)	
                    REPLACE &repSum WITH sumPodrKatKpp1(i,h)
                ENDFOR        
            ENDFOR              
         ENDIF 
         kpp1_cx=curTarJob.kodkpp1
         STORE 0 TO sumPodrKpp1,sumPodrKatKpp1         
      ENDIF         
     *--------
     &&  ���� �� ���������� �������� � ����������������.
     IF kpp_cx#curTarJob.kodkpp.AND.kpp_cx#0.AND.partype#4      
         SELECT curprn
         APPEND BLANK      
         REPLACE nIt WITH 5,kp WITH kp_cx,kd WITH 999,fio WITH '�����',nd WITH 800
         FOR i=1 TO maxItog
             repSum=nSum(i)	
             REPLACE &repSum WITH sumPodrKpp(i)
         ENDFOR
         
         IF itkat
            FOR i=1 TO maxKat
                APPEND BLANK
                REPLACE nIt WITH 2,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),nd WITH 800+i
                FOR h=1 TO maxItog
                    repSum=nSum(h)	
                    REPLACE &repSum WITH sumPodrKatKpp(i,h)
                ENDFOR        
            ENDFOR              
         ENDIF   
         kpp_cx=curTarJob.kodkpp
         STORE 0 TO sumPodrKpp,sumPodrKatKpp
      ENDIF      
      kpp_cx=curTarJob.kodkpp
      kpp1_cx=curTarJob.kodkpp1
      *--------    
      STORE 0 TO sum_podr,sumPodrKat
      SELECT curTarJob
      kp_cx=kp
   ENDIF       
   *--------             
   SELECT curTarJob   
ENDDO
**------����� ����-----------------
SELECT curPrn
APPEND BLANK      
REPLACE nIt WITH 7,kp WITH kp_cx,kd WITH 998,fio WITH IIF(formbase.litr,'�����','�����'),nd WITH 900,np WITH 900
FOR i=1 TO maxItog
    repSum=nSum(i)	
    REPLACE &repSum WITH sumTot(i)
ENDFOR
IF formbase.litr
   FOR i=1 TO max_tr       
       APPEND BLANK
       REPLACE nIt WITH 7,kp WITH kp_cx,kd WITH 998,fio WITH name_tr(i),nd WITH 900+i,np WITH 900
       FOR h=1 TO maxItog
           repSum=nSum(h)	
           REPLACE &repSum WITH sumTotTr(i,h)
       ENDFOR 
   ENDFOR   
ENDIF

*-----����� ���� �� ���������� ���������
SELECT curprn
FOR i=1 TO maxKat
    APPEND BLANK 
   * REPLACE nIt WITH 8,kp WITH kp_cx,kd WITH 998,kat WITH kod_kat(i),fio WITH IIF(formbase.litr,UPPER(name1_kat(i)),name1_kat(i)),nd WITH 910+i*10,np WITH 950
    REPLACE nIt WITH 8,kp WITH kp_cx,kd WITH 998,kat WITH kod_kat(i),fio WITH name1_kat(i),nd WITH 910+i*10,np WITH 950
    FOR h=1 TO maxItog
        repSum=nSum(h)	
        REPLACE &repSum WITH sumTotKat(i,h)
    ENDFOR 
    IF formbase.litr
       FOR itr=1 TO max_tr
           APPEND BLANK
           REPLACE nIt WITH 8,kp WITH kp_cx,kd WITH 998,kat WITH kod_kat(i),fio WITH name_tr(itr),nd WITH (910+i*10)+itr,np WITH 950        
           FOR isum=1 TO maxItog
               repSum=nSum(isum)	
               ntdim='dimkattr'+LTRIM(STR(i))+'('+LTRIM(STR(itr))+','+LTRIM(STR(isum))+')'
               REPLACE &repSum WITH &ntdim            
            ENDFOR  
       ENDFOR  
    ENDIF       
ENDFOR 

DELETE FOR kse=0 
*DO CASE
*   CASE parItog=1           &&��������
*        DELETE FOR nIt<7 
*   CASE parItog=2           && �� ����������
*        DELETE FOR nIt=0 
*ENDCASE
IF parType=2
   DELETE FOR nIt<7
ENDIF
IF logKom
   APPEND BLANK
   REPLACE np WITH 999,nIt WITH 9,kp WITH kp_cx
   APPEND BLANK
   REPLACE np WITH 999,nIt WITH 9,kp WITH kp_cx
   APPEND BLANK
   REPLACE np WITH 999,nprik WITH '��������:',nIt WITH 9,kp WITH kp_cx  
   SELECT boss
   SCAN ALL
        SELECT curprn
        APPEND BLANK
        REPLACE np WITH 999,nprik WITH boss.doljn,primtxt WITH boss.fam,nIt WITH 9,kp WITH kp_cx 
        SELECT boss      
   ENDSCAN 
ENDIF   
SELECT curPrn
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,np) FOR np=0
REPLACE nvac WITH 1 FOR vac
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
kpcx=0
newnpp=0
SCAN ALL  
     IF kp#kpcx 
        kpcx=kp
        newnpp=1     
     ENDIF
     REPLACE npp WITH IIF(nIt=0,newnpp,0)
     IF nIt=0
        SELECT sprdolj
        SEEK curprn.kd
        SELECT curprn
        DO CASE                          
           CASE sprdolj.logSex.AND.!vac
                REPLACE namedol WITH IIF(sex=2,sprdolj.name,IIF(!EMPTY(sprdolj.namem),sprdolj.namem,sprdolj.name))     
           OTHERWISE      
                REPLACE namedol WITH sprdolj.name
        ENDCASE 
        SELECT curPrn
     ENDIF 
     newnpp=newnpp+1
ENDSCAN     
GO TOP

DO CASE
   CASE parType=1
        DO procForPrintAndPreview WITH 'reptarnew','��������������� ������'
   CASE parType=2
        DO procForPrintAndPreview WITH 'reptarnew','��������������� ������'
   CASE parType=3
        DO tarifToExcelNew
   CASE parType=4  
        DO procForPrintAndPreview WITH 'reptarnew','��������������� ������'      
ENDCASE

SELECT people
SET ORDER TO &ord_old
SELECT rasp
************************************************************************************************************************
PROCEDURE tarifToExcel
DO startPrnToExcel WITH 'fSupl'     
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)    
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=18
     .Columns(3).ColumnWidth=20
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=5
     .Columns(6).ColumnWidth=7
     .Columns(7).ColumnWidth=5
     .Columns(8).ColumnWidth=7
     .Columns(9).ColumnWidth=7
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,8)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������ ������� (������) �� ����������(����������) ����������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                        
     rowcx=rowcx+1                          
     rowBeg=rowcx
     .Cells(rowcx,1).Value='� �.�'   
     .Cells(rowcx,2).Value='������������ ���� �����, ������������ �������������,�������, ��� �������� ���������'  
     .Cells(rowcx,3).Value='������������ ���������'
     .Cells(rowcx,4).Value='�������� ������'
     .Cells(rowcx,5).Value='�������� �����������'
     .Cells(rowcx,6).Value='�����'
     .Cells(rowcx,7).Value='����� ������'               
     .Cells(rowcx,8).Value='����� � ������ ������ ������'     
     .Range(.Cells(rowcx,1),.Cells(rowcx,8)).Select  
     WITH objExcel.Selection
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH 
     rowcx=rowcx+1                                         
     numberRow=rowcx+1  
     rowtop=numberRow         
     STORE 0 TO max_rec,one_pers,pers_ch
     SELECT curPrn
     COUNT TO max_rec
     GO TOP
     kpold=0      
     SCAN ALL
          IF kp#kpold
             .Range(.Cells(numberRow,1),.Cells(numberRow,8)).Select
             objExcel.Selection.MergeCells=.T.
             objExcel.Selection.HorizontalAlignment=xlLeft
             objExcel.Selection.VerticalAlignment=1
             objExcel.Selection.WrapText=.T.
             objExcel.Selection.Interior.ColorIndex=37
             objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
             numberRow=numberRow+1
             kpold=kp
          ENDIF 
          .Cells(numberRow,1).Value=curprn.npp   
          .Cells(numberRow,2).Value=curprn.fio   
          .Cells(numberRow,3).Value=IIF(SEEK(curprn.kd,'sprdolj',1),sprdolj.name,'')
          .Cells(numberRow,4).Value=curprn.kf   
          .Cells(numberRow,5).Value=curprn.namekf   
          .Cells(numberRow,6).Value=curprn.tokl
          .Cells(numberRow,7).Value=curprn.kse
          .Cells(numberRow,8).Value=curprn.mtokl
           numberRow=numberRow+1
          SELECT curprn   
          one_pers=one_pers+1
          pers_ch=one_pers/max_rec*100
          fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
          fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch              
     ENDSCAN           
     .Range(.Cells(rowBeg,1),.Cells(numberRow-1,8)).Select
     objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
     objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
     objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
     objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
     objExcel.Selection.VerticalAlignment=1
*          
     .Range(.Cells(rowcx,1),.Cells(numberRow-1,8)).Select
     objExcel.Selection.Font.Name='Times New Roman' 
     objExcel.Selection.Font.Size=8      
     objExcel.Selection.WrapText=.T.  
     .Cells(1,1).Select                       
ENDWITH   
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'   
objExcel.Visible=.T.
*************************************************************************************************************************
PROCEDURE itogprn

*************************************************************************************************************************
PROCEDURE vedunit
PARAMETERS parPers,parSum
*persVed=''
*sumVed=''
*parPers=persVed
*parSum=sumVed

* maxkat=���-�� ��������� ���������
* sum_podr - ���� �� ���������� 
* sumPodrKat - ���� �� ���������� � ������� ���������

* sumTot - ���� �����
* sumTotKat - ���� ����� �� ���������� ���������
* sumKpp -���� �� ����������������
itKat=formBase.log_It
logKom=formBase.log_kom

SELECT sprkat
COUNT TO maxKat
DIMENSION nSum(1)

DIMENSION sum_podr(2),sumTot(2)
DIMENSION sumPodrKat(maxKat,2),sumTotKat(maxKat,2)
STORE 0 TO sum_podr,sumTot,sumPodrKat,sumTotkat

DIMENSION sumPodrKpp(2),sumPodrKpp1(2)  &&����� � ����������������� � ������ ��������������
STORE 0 TO sumPodrKpp,sumPodrKpp1

DIMENSION sumPodrKatKpp(maxKat,2),sumPodrKatKpp1(maxKat,2)  &&����� � ����������������� � ������ �������������� �� ���������� ���������
STORE 0 TO sumPodrKatKpp,sumPodrKatKpp1

IF USED('curprn')
   SELECT curPrn
   USE   
ENDIF
SELECT datjob
SET FILTER TO
SELECT * FROM datjob INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN pprn N(6,2)
ALTER TABLE curPrn ADD COLUMN sprn N(12,2)
ALTER TABLE curPrn ADD COLUMN oklStPrn N(12,2)
ALTER TABLE curPrn ADD COLUMN nIt N(1)
ALTER TABLE curPrn ADD COLUMN kHours N(7,2)
ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
ALTER TABLE curPrn ADD COLUMN sex N(1)
ALTER TABLE curPrn ADD COLUMN namedol C(150)
SELECT curPrn
DELETE FOR tokl=0
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DELETE FOR date_in>dateTar
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,''),sex WITH people.sex ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL

mVr='vr'+LTRIM(STR(MONTH(dateTar)))
REPLACE matt WITH &mVr ALL
IF !EMPTY(persnadbav)
   REPLACE pprn WITH &persnadbav ALL
ENDIF
IF !EMPTY(sumnadbav)
   REPLACE sprn WITH &sumnadbav ALL
ENDIF
IF !EMPTY(sumOklSt)
   REPLACE oklStPrn WITH &sumOklSt ALL
ENDIF
*----------------------------------- �������������� ���������� ��������---------------------------------------------
IF formbase.avt_vac
   SELECT rasp
   GO TOP
   DO WHILE !EOF()
      IF rasp.kse#0
         SELECT datJob
         SET ORDER TO 2
         SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
         kse_cx=rasp.kse
         DO  WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()   
             IF date_in>dateTar        
             ELSE 
                kse_cx=kse_cx-datjob.kse
             ENDIF 
             SKIP    
         ENDDO
         IF kse_cx>0
            IF !formbase.vacst               
               SELECT curPrn
               APPEND BLANK
               REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH kse_cx,vac WITH .T.,;
                       np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,tr WITH 1,pkont WITH rasp.pkont
               tar_ok=0
               tar_ok=varBaseSt*curPrn.namekf*IIF(pkf#0,pkf,1)
               REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2),mkonts WITH mtokl*pkont/100,mstsum WITH mtokl*stpr/100*kse
               
               reppersrasp=IIF(persnadbav='stpr','stpr','rasp.'+persnadbav )
               REPLACE pprn WITH &reppersrasp
               REPLACE &persnadbav WITH pprn
               
               IF !EMPTY(persFormula)
                  repsprn=EVALUATE(persFormula)
                  REPLACE sprn WITH IIF(persKse,repsprn*kse,IIF(repsprn>0,repsprn,0))                                                  
               ENDIF  
               IF !EMPTY(sumOklSt)
                  REPLACE oklStPrn WITH &sumOklSt
               ENDIF                 
            ELSE 
               DO CASE
                  CASE kse_cx<=1
                       kvovac=1
                  CASE MOD(kse_cx,1)=0     
                       kvovac=INT(kse_cx)
                  CASE MOD(kse_cx,1)>0     
                       kvovac=INT(kse_cx)+1    
               ENDCASE               
               kvokse=kse_cx
               ksevac=0
               FOR i=1 TO kvovac
                   ksevac=IIF(kvokse<=1,kvokse,1)
                   SELECT curPrn
                   APPEND BLANK
                   REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH ksevac,vac WITH .T.,;
                       np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,tr WITH 1,pkont WITH rasp.pkont
                   tar_ok=0
                   tar_ok=varBaseSt*curPrn.namekf*IIF(pkf#0,pkf,1)
                   REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2),mkonts WITH mtokl*pkont/100,mstsum WITH varBaseSt*stpr/100*kse
               
                   reppersrasp=IIF(persnadbav='stpr','stpr','rasp.'+persnadbav )
                   *reppersrasp='rasp.'+persnadbav 
                    
                   REPLACE pprn WITH &reppersrasp
                   REPLACE &persnadbav WITH pprn
                   IF !EMPTY(persFormula)
                      repsprn=EVALUATE(persFormula)                   
*                      REPLACE sprn WITH IIF(persKse,repsprn*kse,repsprn)                                                                                                 
                      REPLACE sprn WITH IIF(persKse,repsprn*kse,IIF(repsprn>0,repsprn,0))
                   ENDIF 
                   IF !EMPTY(sumOklSt)
                      REPLACE oklStPrn WITH &sumOklSt
                   ENDIF                      
                   kvokse=kvokse-1
               ENDFOR
            ENDIF    
         ENDIF       
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
ENDIF  
SELECT curprn
repHours='sprtime.T'+LTRIM(STR(MONTH(dateTar)))
REPLACE kHours WITH IIF(SEEK(vTime,'sprtime',1),&repHours,0) ALL
IF LOWER(ALLTRIM(persnadbav))='patt'
   REPLACE sumvr WITH varBaseSt/100*patt,matt WITH sumvr*kHours*kse,sprn WITH matt FOR vac.AND.patt>0   
ENDIF 
DELETE FOR tokl=0
DO fltstructure WITH 'kse#0.AND.sprn#0','curprn'
IF onlyVac
   SELECT curPrn
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curPrn
   DELETE FOR vac
ENDIF
IF LOWER(ALLTRIM(persnadbav))='patt'
   DELETE FOR kse#1.AND.!latt
ENDIF 
SELECT curPrn
DELETE FOR tokl=0
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,np) ALL
SELECT sprpodr      
SCAN ALL
     SELECT curprn
     SUM kse,sprn TO ksecx,sprncx FOR kp=sprpodr.kod
     IF sprncx#0
        APPEND BLANK
        REPLACE fio WITH '�����',kp WITH sprpodr.kod,np WITH sprpodr.np,nd WITH 500,sprn WITH sprncx,kse WITH ksecx,nIt WITH 1     
     ENDIF
     IF itkat
        FOR i=1 TO maxKat
            SELECT curPrn
            SUM kse,sprn TO ksecx,sprncx FOR kp=sprpodr.kod.AND.kat=kod_kat(i)  
            APPEND BLANK
            REPLACE nIt WITH 2,kp WITH sprpodr.kod,np WITH sprpodr.np,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),nd WITH 700+i,;
                    sprn WITH sprncx,kse WITH ksecx                    
           
        ENDFOR              
      ENDIF     
     SELECT sprpodr
ENDSCAN
SELECT curPrn
DELETE FOR sprn=0
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
GO TOP
nppcx=1
kpOld=kp
DO WHILE !EOF()  
   REPLACE npp WITH nppcx
   nppcx=nppcx+1
   IF nIt=0
      SELECT sprdolj
      SEEK curprn.kd
      SELECT curprn
      DO CASE                          
         CASE sprdolj.logSex.AND.!vac
              REPLACE namedol WITH IIF(sex=2,sprdolj.name,IIF(!EMPTY(sprdolj.namem),sprdolj.namem,sprdolj.name))     
         OTHERWISE      
              REPLACE namedol WITH sprdolj.name
      ENDCASE 
      SELECT curPrn
   ENDIF 
   SKIP
   IF kpOld#kp
      kpOld=kp
      nppcx=1 
   ENDIF  
ENDDO 
SELECT curprn
SUM kse,sprn TO ksecx,sprncx FOR nIt=0
IF sprncx#0
   APPEND BLANK
   REPLACE fio WITH IIF(formbase.lItr,'�����','�����'),kp WITH 999,np WITH 999,nd WITH 500,sprn WITH sprncx,kse WITH ksecx,nIt WITH 9  
   IF formbase.litr
      FOR i=1 TO max_tr
          SUM kse,sprn TO ksecx,sprncx FOR nIt=0.AND.tr=i
          IF ksecx#0
             APPEND BLANK
             REPLACE fio WITH name_tr(i),kp WITH 999,np WITH 999,nd WITH 500,sprn WITH sprncx,kse WITH ksecx,nIt WITH 9  
          ENDIF 
      ENDFOR
   ENDIF
   FOR i=1 TO maxKat
       SELECT curPrn
       SUM kse,sprn TO ksecx,sprncx FOR kat=kod_kat(i).AND.nIt=0  
       IF ksecx#0
          APPEND BLANK
          REPLACE nIt WITH 8,kp WITH 999,np WITH 999,kd WITH 999,kat WITH kod_kat(i),fio WITH IIF(formbase.lItr,UPPER(name1_kat(i)),name1_kat(i)),nd WITH 700+i,;
                  sprn WITH sprncx,kse WITH ksecx 
          
          IF formbase.litr
             FOR ix=1 TO max_tr
                 SUM kse,sprn TO ksetr,sprntr FOR nIt=0.AND.tr=ix.AND.kat=i
                 IF ksecx#0
                    APPEND BLANK
                    REPLACE fio WITH name_tr(ix),kp WITH 999,np WITH 999,nd WITH 700+i,sprn WITH sprntr,kse WITH ksetr,nIt WITH 9  
                 ENDIF 
             ENDFOR
          ENDIF                                                
       ENDIF            
   ENDFOR     
ENDIF	
	
IF EMPTY(nSupProc)
   DO procForPrintAndPreview WITH 'repunit','��������������� ������'
ELSE    
   DO &nSupProc
ENDIF    
SELECT curPrn
USE
SELECT people
*************************************************************************************************************************
PROCEDURE nadbavToExcel
DO CASE
   CASE persnadbav='pkat'
        DO katToExcel
   CASE persnadbav='patt'   
        DO attToExcel
   OTHERWISE 
        DO unitToExcel
ENDCASE    
*************************************************************************************************************************
PROCEDURE unitToExcel
DO startPrnToExcel WITH 'fSupl'
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=25
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
         
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=headNadbav
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH 
      
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�� ��������� �� '+DTOC(dateTar)+' �.'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH     
     rowcx=rowcx+1  
                                      
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='� �/�'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���� �����, ������������ ������������� ������� ��� �������� ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH       
        
      .Range(.Cells(rowcx,4),.Cells(rowcx,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'                   
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,5),.Cells(rowcx,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����/������� ������'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,6),.Cells(rowcx,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='%'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,7),.Cells(rowcx,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH              
              
      rowcx=rowcx+1
      .cells(rowcx,1).Value='1'
      .cells(rowcx,2).Value='2'
      .cells(rowcx,3).Value='3'
      .cells(rowcx,4).Value='4'
      .cells(rowcx,5).Value='5'
      .cells(rowcx,6).Value='6'
      .cells(rowcx,7).Value='7'
  
      .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      rowtop=numberRow         
      SELECT curPrn
      DO storezeropercent
      GO TOP
      kpold=0
      SCAN ALL
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,7)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
           .Cells(numberRow,1).Value=IIF(nIt=0,curprn.npp,'')                                    
           .Cells(numberRow,2).Value=curprn.fio                                       
           .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')                                    
           .Cells(numberRow,4).Value=kse                                      
           .Cells(numberRow,4).NumberFormat='0.00'               
           .Cells(numberRow,5).Value=IIF(nIt=0,curprn.oklstprn,'')                                      
           .Cells(numberRow,6).Value=IIF(nIt=0,LTRIM(STR(pprn,6,2))+'%','')
           .Cells(numberRow,7).Value=sprn  
           .Cells(numberRow,7).NumberFormat='0.00'  
           numberRow=numberRow+1
           DO fillpercent WITH 'fSupl'
                   
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,7)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
      
      IF logKom
         numberRow=numberRow+1
         FOR i=1 TO 12
             .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,1)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH 
              
             .Range(.Cells(numberRow,5),.Cells(numberRow,7)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,2)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH                    
             numberRow=numberRow+1
         ENDFOR
      ENDIF
          
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,7)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'   
objExcel.Visible=.T.
*************************************************************************************************************************
PROCEDURE katToExcel
DO startPrnToExcel WITH 'fSupl'
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=25
     .Columns(4).ColumnWidth=15
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
     .Columns(8).ColumnWidth=8
         
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,8)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,8)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=headNadbav
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH    
     rowcx=rowcx+1 
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�� ��������� �� '+DTOC(dateTar)+' �.'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH     
     rowcx=rowcx+1   
                                      
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='� �/�'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���� �����, ������������ ������������� ������� ��� �������� ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH 
      
     .Range(.Cells(rowcx,4),.Cells(rowcx,4)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH         
        
      .Range(.Cells(rowcx,5),.Cells(rowcx,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'                   
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,6),.Cells(rowcx,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����/������� ������'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,7),.Cells(rowcx,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='%'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,8),.Cells(rowcx,8)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH              
              
      rowcx=rowcx+1
      .cells(rowcx,1).Value='1'
      .cells(rowcx,2).Value='2'
      .cells(rowcx,3).Value='3'
      .cells(rowcx,4).Value='4'
      .cells(rowcx,5).Value='5'
      .cells(rowcx,6).Value='6'
      .cells(rowcx,7).Value='7'
      .cells(rowcx,8).Value='8'
  
      .Range(.Cells(rowcx,1),.Cells(rowcx,8)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      rowtop=numberRow         
      SELECT curPrn
      DO storezeropercent
      GO TOP
      kpold=0
      SCAN ALL
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,8)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
           .Cells(numberRow,1).Value=IIF(nIt=0,curprn.npp,'')                                    
           .Cells(numberRow,2).Value=curprn.fio                                       
           .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')                                    
           .Cells(numberRow,4).Value=IIF(SEEK(kv,'sprkval',1),sprkval.name,'')                                    
           .Cells(numberRow,5).Value=kse                                      
           .Cells(numberRow,5).NumberFormat='0.00'           
           .Cells(numberRow,6).Value=IIF(nIt=0,curprn.oklstprn,'')                                      
           .Cells(numberRow,7).Value=IIF(nIt=0,LTRIM(STR(pprn,6,2))+'%','')
           .Cells(numberRow,8).Value=sprn  
           .Cells(numberRow,8).NumberFormat='0.00'  
           numberRow=numberRow+1
           DO fillpercent WITH 'fSupl'
                   
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,8)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
     
      IF logKom
         numberRow=numberRow+1
        .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlLeft
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='��������'                   
              .Font.Name='Times New Roman'   
              .Font.Size=9
         ENDWITH 
         numberRow=numberRow+1
         FOR i=1 TO 12
             .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,1)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH 
              
             .Range(.Cells(numberRow,5),.Cells(numberRow,7)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,2)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH                    
             numberRow=numberRow+1
         ENDFOR
      ENDIF
          
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,8)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'         
objExcel.Visible=.T.

*************************************************************************************************************************
PROCEDURE attToExcel
DO startPrnToExcel WITH 'fSupl' 
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   

WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=25
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
     .Columns(8).ColumnWidth=8
     .Columns(9).ColumnWidth=8
         
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������ �������� �� ������� ������� ���������� �� ����������� ���������� ������� ����'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
      
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�� '+ALLTRIM(dim_month(MONTH(dateTar)))+' '+STR(YEAR(dateTar),4)+' �.'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
        
     rowcx=rowcx+1  
                                      
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='� �/�'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���� �����, ������������ ������������� ������� ��� �������� ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH      
    
        
      .Range(.Cells(rowcx,4),.Cells(rowcx,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'                   
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,5),.Cells(rowcx,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='������� ������'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,6),.Cells(rowcx,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='%'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,7),.Cells(rowcx,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='� ���'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
     .Range(.Cells(rowcx,8),.Cells(rowcx,8)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�����'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     
    .Range(.Cells(rowcx,9),.Cells(rowcx,9)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�����'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                                  
              
      rowcx=rowcx+1
      .cells(rowcx,1).Value='1'
      .cells(rowcx,2).Value='2'
      .cells(rowcx,3).Value='3'
      .cells(rowcx,4).Value='4'
      .cells(rowcx,5).Value='5'
      .cells(rowcx,6).Value='6'
      .cells(rowcx,7).Value='7'
      .cells(rowcx,8).Value='8'
      .cells(rowcx,9).Value='9'
  
      .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      rowtop=numberRow         
      SELECT curPrn
      DO storezeropercent
      GO TOP
      kpold=0
      SCAN ALL
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,9)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
           .Cells(numberRow,1).Value=IIF(nIt=0,curprn.npp,'')                                    
           .Cells(numberRow,2).Value=curprn.fio                                       
           .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')                                                                        
           .Cells(numberRow,4).Value=kse                                      
           .Cells(numberRow,4).NumberFormat='0.00'           
           .Cells(numberRow,5).Value=IIF(nIt=0,varBaseSt,'')                                      
           .Cells(numberRow,6).Value=IIF(nIt=0,LTRIM(STR(pprn,6,2))+'%','')
           .Cells(numberRow,7).Value=IIF(nIt=0,curprn.sumvr,'')                                      
           .Cells(numberRow,8).Value=IIF(nIt=0,curprn.kHours,'')  
           .Cells(numberRow,8).NumberFormat='0.00'             
           .Cells(numberRow,9).Value=sprn  
           .Cells(numberRow,9).NumberFormat='0.00'  
           numberRow=numberRow+1
           DO fillpercent WITH 'fSupl'
                   
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,9)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
      
      IF logKom
         numberRow=numberRow+1
         .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlLeft
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='��������'                   
              .Font.Name='Times New Roman'   
              .Font.Size=9
         ENDWITH 
         numberRow=numberRow+1
         FOR i=1 TO 12
             .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,1)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH 
              
             .Range(.Cells(numberRow,5),.Cells(numberRow,7)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,2)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH                    
             numberRow=numberRow+1
         ENDFOR
      ENDIF          
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,9)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'    
objExcel.Visible=.T.

*************************************************************************************************************************
PROCEDURE shtatPrn
PARAMETERS par1
year_shtat=YEAR(setuptar.datetar)
IF USED('curSuplPrn')
  SELECT curSuplPrn
  USE 
ENDIF
IF USED('curSuplKat')
  SELECT curSuplKat
  USE 
ENDIF
SELECT * FROM sprkat INTO CURSOR curSuplKat READWRITE
SELECT * FROM rasp INTO CURSOR curSuplPrn READWRITE 

ALTER TABLE curSuplPrn ADD COLUMN np N(3)
ALTER TABLE curSuplPrn ADD COLUMN named C(150)
ALTER TABLE curSuplPrn ADD COLUMN namepodr C(150)
ALTER TABLE curSuplPrn ADD COLUMN nIt N(1)
ALTER TABLE curSuplPrn ADD COLUMN logIt L

ALTER TABLE curSuplPrn ADD COLUMN vac L
SELECT curSuplPrn
REPLACE np WITH IIF(SEEK(curSuplPrn.kp,'sprpodr',1),sprpodr.np,curSuplPrn.np) ALL
REPLACE named WITH IIF(SEEK(curSuplPrn.kd,'sprdolj',1),sprdolj.name,curSuplPrn.named) ALL
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
DO fltStructure WITH .F.,'curSuplPrn'
SELECT curSuplPrn
GO TOP
ndNew=1
kpOld=kp
DO WHILE !EOF()
   REPLACE nd WITH ndNew
   SELECT curSuplPrn
   SKIP 
   ndNew=ndNew+1
   IF kp#kpOld
      kpOld=kp
      ndNew=1
   ENDIF  
ENDDO
REPLACE namepodr WITH IIF(SEEK(curSuplPrn.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.prim),sprpodr.prim,sprpodr.name),'') FOR nd=1

=AFIELDS(arRasp,'curSuplPrn')
CREATE CURSOR curPrn FROM ARRAY arRasp
SELECT curSuplPrn
SCAN ALL
     SCATTER TO dimAp
     SELECT curPrn
     APPEND BLANK 
     GATHER FROM dimAp
     SELECT curSuplPrn
ENDSCAN

SELECT curSuplPrn
GO TOP 
STORE 0 TO kseTot,totTot,stacTot,skTot,mjrTot,amTot,ksePodr,totPodr,stacPodr,mjrPodr,amPodr,skPodr
kpOld=kp
npOld=np
DO WHILE !EOF()
   kseTot=kseTot+kse 
   totTot=totTot+kse_tot
   amTot=amTot+kse_am
   skTot=skTot+kse_sk
   stacTot=stacTot+kse_stac
   mjrTot=mjrTot+kse_mjr
   
   ksePodr=ksePodr+kse
   totPodr=totPodr+kse_tot
   amPodr=amPodr+kse_am
   stacPodr=stacPodr+kse_stac
   skPodr=skPodr+kse_sk
   mjrPodr=mjrPodr+kse_mjr
   
   SELECT curSuplPrn
   SKIP
   IF kp#kpOld
      SELECT curPrn
      APPEND BLANK
      REPLACE kp WITH kpOld,np WITH npOld,nd WITH 90,named WITH '�����',kse WITH ksePodr,kse_tot WITH totPodr,kse_am WITH amPodr,;
              kse_stac WITH stacPodr,kse_sk WITH skPodr,kse_mjr WITH mjrPodr, logIt WITH .T.     
      SELECT curSuplPrn
      STORE 0 TO ksePodr,totPodr,stacPodr,mjrPodr,amPodr,skPodr
      kpOld=kp
      npOld=np
   ENDIF
ENDDO 
SELECT curPrn
APPEND BLANK
REPLACE np WITH 900,nd WITH 900,named WITH '����� �� �����������',kse WITH kseTot,kse_tot WITH totTot,kse_stac WITH stacTot,kse_am WITH amTot,;
        kse_sk WITH skTot,kse_mjr WITH mjrTot ,logIt WITH .T.       
IF formBase.log_It
   SELECT curSuplPrn
   SET FILTER TO nd=1
   SCAN ALL
        SELECT curSuplKat
        SCAN ALL 
             SELECT curPrn
             SUM kse,kse_tot,kse_stac,kse_am,kse_sk,kse_mjr TO kse_kat,totTot,stacTot,amTot,skTot,mjrTot FOR kat=curSuplKat.kod.AND.kp=curSuplPrn.kp
             IF kse_kat#0              
                APPEND BLANK
                REPLACE kp WITH curSuplPrn.kp,np WITH curSuplPrn.np,nd WITH 900,kse WITH kse_kat,kse_tot WITH totTot,kse_stac WITH stacTot,;
                        kse_am WITH amTot,kse_sk WITH skTot,kse_mjr WITH mjrTot,named WITH curSuplKat.name,logIt WITH .T.  
             ENDIF  
             SELECT curSuplKat               
       ENDSCAN
       SELECT curSuplPrn  
   ENDSCAN
ENDIF        
SELECT curSuplKat
SCAN ALL 
     SELECT curPrn
     SUM kse,kse_tot,kse_stac,kse_am,kse_sk,kse_mjr TO kse_kat,totTot,stacTot,amTot,skTot,mjrTot FOR kat=curSuplKat.kod.AND.nd<90
     IF kse_kat#0 
        SELECT curPrn
        APPEND BLANK
        REPLACE np WITH 900,nd WITH 900,kse WITH kse_kat,kse_tot WITH totTot,kse_stac WITH stacTot,kse_am WITH amTot,kse_sk WITH skTot,kse_mjr WITH mjrTot,named WITH curSuplKat.name,logIt WITH .T.  
     ENDIF  
     SELECT curSuplKat               
ENDSCAN      
SELECT curPrn  
INDEX ON STR(np,3)+STR(nd,3) TAG T1
GO TOP
DO procForPrintAndPreview WITH 'shtatRasp','������� ����������'

*************************************************************************************************************************
*                 ��������� ������ ���������� 
*************************************************************************************************************************
PROCEDURE vedprn
PARAMETERS par1
terminal=par1
SELECT formbase
procch=formbase.proc
DO &procch
*************************************************************************************************************************
PROCEDURE tarifPrnNew
itkat=formbase.log_it
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF  
IF USED('curTarJob')
   SELECT curTarJob
   USE
ENDIF    

IF USED('curFondPrn')
   SELECT curFondPrn
   USE
ENDIF    
SELECT npved,nameved,persved,sumved FROM tarfond WHERE !EMPTY(nameved).AND.!EMPTY(sumved) INTO CURSOR curFondPrn READWRITE
SELECT curFondPrn
INDEX ON npved TAG T1
COUNT TO maxSum
maxSum=maxSum+2
** ������������ ����������� �����
DIMENSION nameSum(maxSum)
STORE '' TO nameSum
SELECT curFondPrn
GO TOP
nameSum(1)='kse'
nameSum(2)='mTokl'
FOR i=3 TO maxSum
    nameSum(i)=ALLTRIM(sumVed)
    SKIP
ENDFOR 

** ����� ���� �� �����������
DIMENSION sumTot(maxSum)
STORE 0 TO sumTot
** ����� ���� �� �������������
DIMENSION sumPodr(maxSum)
STORE 0 TO sumPodr


SELECT num,rec,fname,fpers,plrep,vac FROM tarfond WHERE tarfond.vac INTO CURSOR curPrnTarFond READWRITE 
SELECT curPrnTarFond
INDEX ON num TAG T1
GO TOP
num_cx=0
DO WHILE !EOF()
   num_cx=num_cx+1
   REPLACE num WITH num_cx
   SKIP   
ENDDO

SELECT sprkat
* maxkat=���-�� ��������� ���������
* sumTot - ���� �����
* sumKpp -���� �� ����������������
COUNT TO maxKat
**����� ���� �� ���������� ���������
DIMENSION sumTotKat(maxKat,maxSum)
STORE 0 TO sumTotKat
**���� �� ���������� ��������� � ��������������
DIMENSION sumPodrKat(maxKat,maxSum)
STORE 0 TO sumPodrKat
*****


DIMENSION sumKpp(1,2),sumKpp1(1,2)
STORE 0 TO numrecrep,numpage,sumKpp,sumKpp1
DIMENSION sumKatKpp(maxKat,2),sumKatKpp1(maxKat,2) && ����� �� ���������� ��������� � �������������������� � �����������������
STORE 0 TO sumKatKpp,sumKatKpp1


SELECT * FROM datJob INTO CURSOR curTarJob READWRITE
ALTER TABLE curTarJob ADD COLUMN npp N(3)
ALTER TABLE curTarJob ADD COLUMN nit N(1)
ALTER TABLE curTarJob ADD COLUMN nkat C(200)
SELECT curTarJob
DELETE FOR !SEEK(kp,'sprpodr',1)
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.np,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL
INDEX ON STR(np,3)+STR(ND,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 2

*----------------------------------- �������������� ���������� �������� ---------------------------------------------
IF formbase.avt_vac   
   SELECT datjob
   ordOld=SYS(21)
   SET FILTER TO 
   SELECT rasp
   GO TOP
   DO WHILE !EOF()
      IF rasp.kse#0
         SELECT datjob
         SET ORDER TO 2
         SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
         kse_cx=rasp.kse
         DO  WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()
             kse_cx=kse_cx-datjob.kse
             SKIP
         ENDDO
         IF kse_cx>0
            IF !formbase.vacst
               SELECT curTarJob
               APPEND BLANK
               REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfVac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,kse WITH kse_cx,vac WITH .T.
               SELECT curPrnTarFond
               GO TOP
               DO WHILE !EOF()
                  rep_r=ALLTRIM(plrep)
                  rep_r1='rasp.'+ALLTRIM(plRep)
                  SELECT curTarJob 
                  REPLACE &rep_r WITH &rep_r1         
                  SELECT curPrnTarFond
                  SKIP
               ENDDO                                               
               SELECT curTarJob
               DO countOkladVac             
            ELSE 
               DO CASE
                  CASE kse_cx<=1
                       kvovac=1
                  CASE MOD(kse_cx,1)=0     
                       kvovac=INT(kse_cx)
                  CASE MOD(kse_cx,1)>0     
                       kvovac=INT(kse_cx)+1    
               ENDCASE               
               kvokse=kse_cx
               ksevac=0
               FOR i=1 TO kvovac
                   ksevac=IIF(kvokse<=1,kvokse,1)
                   SELECT curTarJob
                   APPEND BLANK
                   REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfVac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,kse WITH ksevac,vac WITH .T.
                   SELECT curPrnTarFond
                   GO TOP
                   DO WHILE !EOF()
                      rep_r=ALLTRIM(plrep)
                      rep_r1='rasp.'+ALLTRIM(plrep)
                      SELECT curTarJob 
                      REPLACE &rep_r WITH &rep_r1         
                      SELECT curPrnTarFond
                      SKIP
                   ENDDO                                               
                   SELECT curTarJob
                   DO countOkladVac
                   kvokse=kvokse-1
               ENDFOR
            ENDIF    
         ENDIF       
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
   SELECT datJob
   SET ORDER TO &ordOld   
ENDIF   


*--------------------------------------------------------------------------------------------------------------------
=AFIELDS(arJob,'curTarJob')
CREATE CURSOR curPrn FROM ARRAY arJob
ALTER TABLE curprn ALTER COLUMN kse N(10,2)

DO fltstructure WITH 'mTokl#0'

*---������ ��� �������� ����������
IF USED('curRasp')
   SELECT curRasp
   USE
ENDIF 
SELECT rasp      
SELECT * FROM rasp WHERE SEEK(STR(kp,3)+STR(kd,3),'curTarJob',2) INTO CURSOR currasp READWRITE      
SELECT currasp
SCAN ALL
    IF SEEK(curRasp.kp,'sprpodr',1)
        REPLACE np WITH sprpodr.np,kpp WITH sprpodr.kpp,kpp1 WITH sprpodr.kpp1,kodkpp WITH sprpodr.kodkpp,kodkpp1 WITH sprpodr.kodkpp1  
    ENDIF   
ENDSCAN     
SELECT sprpodr
SCAN ALL
     IF sprpodr.kodKpp#0
        SELECT curRasp
        REPLACE kodkpp WITH sprpodr.kodkpp FOR kp=sprpodr.kodkpp
     ENDIF 
     IF sprpodr.kodKpp1#0
        SELECT curRasp
        REPLACE kodkpp1 WITH sprpodr.kodkpp1,kodkpp WITH sprpodr.kodkpp FOR kp=sprpodr.kodkpp1
        
     ENDIF 
     SELECT sprpodr
ENDSCAN   

SELECT curRasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1
GO TOP
kp_old=kp
num_new=1
DO WHILE !EOF()
   REPLACE nd WITH num_new
   IF nd=1 
      SELECT rasp
      LOCATE FOR kp=currasp.kp.AND.nd=1
      SELECT currasp
      REPLACE primhead WITH rasp.primhead
   ENDIF
   num_new=num_new+1
   SKIP
   IF kp#kp_old    
      kp_old=kp 
      num_new=1
   ENDIF
ENDDO

SELECT currasp
SET ORDER TO 2
SELECT curTarJob
ord_old=SYS(21)
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curTarJob.kp,3)+STR(curTarJob.kd,3)   
   SELECT curTarJob
   REPLACE nd WITH currasp.nd,np WITH currasp.np,kpp WITH currasp.kpp,kodkpp WITH currasp.kodkpp,kpp1 WITH currasp.kpp1,kodkpp1 WITH currasp.kodkpp1,;
   kat WITH IIF(currasp.kat#curTarJob.kat,currasp.kat,curTarJob.kat)
   *primhead WITH currasp.primhead
   SKIP 
ENDDO 

SELECT curTarJob
SET ORDER TO 1
GO TOP
kp_cx=kp
kpp_cx=kodkpp
kpp1_cx=kodkpp1
npp_cx=0
DO WHILE !EOF()     
   SCATTER TO ac
   SELECT curPrn
   npp_cx=npp_cx+1
   APPEND BLANK
   REPLACE npp WITH npp_cx
   GATHER FROM ac                        
   log_sumdopl=.F.    
   
   FOR i=1 TO maxSum
       sumPodr(i)=sumPodr(i)+&nameSum(i)
       sumTot(i)=sumTot(i)+&nameSum(i)
       IF kat#0
          sumTotKat(ASCAN(kod_kat,curPrn.kat),i)=sumTotKat(ASCAN(kod_kat,curPrn.kat),i)+&nameSum(i)
       ENDIF    
   ENDFOR
   
   sumKpp(1,1)=IIF(kpp_cx#0,sumKpp(1,1)+kse,sumKpp(1,1))        
   sumKpp(1,2)=IIF(kpp_cx#0,sumKpp(1,2)+mTokl,sumKpp(1,2))        
           
   sumKpp1(1,1)=IIF(kpp1_cx#0,sumKpp1(1,1)+kse,sumKpp1(1,1))
   sumKpp1(1,2)=IIF(kpp1_cx#0,sumKpp1(1,2)+kse,sumKpp1(1,2))
   
                  
   IF itkat.AND.curPrn.kat#0       &&������ �� ���������� ��������� � ��������������,��� � ������
      FOR i=1 TO maxSum
          IF kat#0
             sumPodrKat(ASCAN(kod_kat,curPrn.kat),i)=sumPodrKat(ASCAN(kod_kat,curPrn.kat),i)+&nameSum(i)
          ENDIF    
      ENDFOR                  
   ENDIF
     
  
   SELECT curTarJob   
   SKIP
   IF kp_cx#curTarJob.kp   
      npp_cx=0         
      SELECT curPrn
      APPEND BLANK      
      REPLACE nIt WITH 1,kp WITH kp_cx,kd WITH 999,fio WITH '�����'
      FOR i=1 TO maxSum
          REPLACE &nameSum(i) WITH sumPodr(i)        
      ENDFOR  
      &&  ���� �� ���������� �������� � ���������
      IF itkat
         FOR i=1 TO maxKat
             APPEND BLANK 
             REPLACE nIt WITH 2,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i)
             FOR xm=1 TO maxSum
                 REPLACE &nameSum(xm) WITH sumPodrKat(i,xm)
              ENDFOR  
         ENDFOR                                      
      ENDIF
      
      IF kpp1_cx#curTarJob.kodkpp1.AND.kpp1_cx#0
         &&  ���� �� ���������� �������� � ���-����������������
         SELECT curprn
         APPEND BLANK      
         REPLACE nIt WITH 3,kp WITH kp_cx,kd WITH 999,fio WITH '�����',kse WITH sumKpp1(1,1),mTokl WITH sumKpp1(1,2)        
         IF itkat
            FOR i=1 TO maxKat
                APPEND BLANK 
                REPLACE nIt WITH 6,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),kse WITH sumKatKpp1(i,1),mtokl WITH sumKatKpp1(i,2)             
            ENDFOR      
         ENDIF   
         kpp1_cx=curTarJob.kodkpp1
         STORE 0 TO sum_kpp1,sumKpp1,sumKatKpp1
         
      ENDIF         
     *--------
     &&  ���� �� ���������� �������� � ����������������
     IF kpp_cx#curTarJob.kodkpp.AND.kpp_cx#0      
         SELECT curprn
         APPEND BLANK      
         REPLACE nIt WITH 5,kp WITH kp_cx,kd WITH 999,fio WITH '�����',kse WITH sumKpp(1,1),mTokl WITH sumKpp(1,2)
         IF itkat
            FOR i=1 TO maxKat
                APPEND BLANK 
                REPLACE nIt WITH 6,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),kse WITH sumKatKpp(i,1),mtokl WITH sumKatKpp(i,2)             
            ENDFOR      
         ENDIF          
         kpp_cx=curTarJob.kodkpp
         STORE 0 TO sum_kpp,sumKpp,sumKatKpp
      ENDIF      
      kpp_cx=curTarJob.kodkpp
      kpp1_cx=curTarJob.kodkpp1
      *--------
      STORE 0 TO sumPodr
      STORE 0 TO sumPodrKat
      SELECT curTarJob
      kp_cx=kp
   ENDIF       
   *--------             
   SELECT curTarJob   
ENDDO
*------����� ����-----------------
SELECT curPrn
APPEND BLANK      
REPLACE nIt WITH 7,kp WITH kp_cx,kd WITH 999,fio WITH '�����'
FOR i=1 TO maxSum
    REPLACE &nameSum(i) WITH sumTot(i)
ENDFOR  

*-----����� ���� �� ���������� ���������
SELECT curprn
FOR i=1 TO maxKat
    APPEND BLANK 
    REPLACE nIt WITH 8,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i)
    FOR xm=1 TO maxSum
        REPLACE &nameSum(xm) WITH sumTotKat(i,xm)
    ENDFOR  
ENDFOR 
DELETE FOR kse=0 
kpcx=0
newnpp=0
SCAN ALL  
     IF kp#kpcx 
        kpcx=kp
        newnpp=1     
     ENDIF
     REPLACE npp WITH newnpp
     newnpp=newnpp+1
ENDSCAN      
DO procForPrintAndPreview WITH 'reptarif','��������������� ������'
SELECT people
**************************************************************************************************************************
*                  ��������� ������� ������ �� ���������
**************************************************************************************************************************
PROCEDURE countOkladVac
SELECT curTarJob
tar_ok=0
tar_ok=varBaseSt*curTarJob.namekf*IIF(curTarJob.pkf#0,curTarJob.pkf,1)
REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2)
*totsumf=tokl
*totsumfm=mtokl
totsumf=0
totsumfm=0
SELECT tarfond
SET FILTER TO !EMPTY(countvac)
GO TOP
DO WHILE !EOF()
   new_sum=sum_f
   new_msum=sum_fm
   SELECT curTarJob
   r_sum=EVAL(tarfond.countvac)     
   IF !EMPTY(tarfond.sum_f) 
      REPLACE &new_sum WITH r_sum  
     * REPLACE &new_msum WITH IIF(tarfond.logkse,&new_sum*kse,&new_sum)     
      REPLACE &new_msum WITH IIF(tarfond.logkse,r_sum*kse,&new_sum)     
      totsumf=IIF(!EMPTY(tarfond.sum_fm),totsumf+EVALUATE(ALLTRIM(tarfond.sum_fm)),totsumf)
      totsumfm=IIF(!EMPTY(tarfond.sum_fm).AND.tarfond.logfprn,totsumfm+EVALUATE(ALLTRIM(tarfond.sum_fm)),totsumfm)
   ENDIF     
   SELECT tarfond
   SKIP
ENDDO
SET FILTER TO
SELECT curTarJob
REPLACE msf WITH totsumf,fdprn WITH totsumfm

*************************************************************************************************************************
PROCEDURE tarifToExcelNew
DO startPrnToExcel WITH 'fSupl'   
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)     
SELECT curFondPrn
COUNT TO maxFond

WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=3     
     .Columns(2).ColumnWidth=30        
     .Columns(3).ColumnWidth=20         
     .Columns(4).ColumnWidth=8         
     .Columns(5).ColumnWidth=8         
     .Columns(6).ColumnWidth=8         
     .Columns(7).ColumnWidth=8         
     .Columns(8).ColumnWidth=8 
      
      rowcx=1     
     .Cells(rowcx,3).Value='������� ������'  
     .Cells(rowcx,4).Value=180       
     rowcx=rowcx+1                          
     .Range(.Cells(rowcx,1),.Cells(rowcx+3,1)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='�'  
     ENDWITH     
     .Range(.Cells(rowcx,2),.Cells(rowcx+3,2)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='������������ ���� ������, ������������ �������������, �������, ���, �������� ���������'  
     ENDWITH    
          
     .Range(.Cells(rowcx,3),.Cells(rowcx+3,3)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='������������ ���������'  
     ENDWITH     
     
     
     .Range(.Cells(rowcx,4),.Cells(rowcx+3,4)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='���������'  
     ENDWITH     
          
     .Range(.Cells(rowcx,5),.Cells(rowcx+3,5)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='����'  
     ENDWITH         
     
     .Range(.Cells(rowcx,6),.Cells(rowcx+3,6)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='�.���'  
     ENDWITH     
     
    .Range(.Cells(rowcx,7),.Cells(rowcx+3,7)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='�.���'  
     ENDWITH          
      
     .Range(.Cells(rowcx,8),.Cells(rowcx+3,8)).Select
     With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=9
           .Value='�����'  
     ENDWITH     
     
     columncx=9
     **  ��������� �������� �� ��������          
     SELECT curFondPrn
     GO TOP
     DO WHILE !EOF()
        DO CASE 
           CASE !EMPTY(persved).AND.!EMPTY(sumved)
                .Columns(columncx).ColumnWidth=4         
                .Columns(columncx+1).ColumnWidth=8   
                .Range(.Cells(rowcx,columncx),.Cells(rowcx+2,columncx+1)).Select
                With objExcel.Selection
                     .MergeCells=.T.
                     .HorizontalAlignment=xlCenter
                     .VerticalAlignment=1
                     .WrapText=.T.
                     .Font.Name='Times New Roman'   
                     .Font.Size=9
                     .Value=curFondPrn.nameved  
                ENDWITH
                .cells(rowcx+3,columncx).Value='%'
                .cells(rowcx+3,columncx+1).Value='�����'
                columncx=columncx+2
                
           CASE !EMPTY(sumved).AND.EMPTY(persved)
                .Columns(columncx).ColumnWidth=8   
                .Range(.Cells(rowcx,columncx),.Cells(rowcx+2,columncx)).Select
                With objExcel.Selection
                     .MergeCells=.T.
                     .HorizontalAlignment=xlCenter
                     .VerticalAlignment=1
                     .WrapText=.T.
                     .Font.Name='Times New Roman'   
                     .Font.Size=9
                     .Value=curFondPrn.nameved  
                ENDWITH
                columncx=columncx+1
                
        ENDCASE  
        SELECT curFondprn
        SKIP 
     ENDDO   
     maxColumn=columncx-1
     rowcx=rowcx+4
     kpold=0
     SELECT curprn
     DO storezeropercent
     GO TOP
     DO WHILE !EOF()
        IF kp#kpold
           .Range(.Cells(rowcx,1),.Cells(rowcx,maxColumn)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.HorizontalAlignment=xlLeft
           objExcel.Selection.VerticalAlignment=1
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Interior.ColorIndex=37
           objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
           rowcx=rowcx+1
           kpold=kp
        ENDIF                                                    
        .Cells(rowcx,1).Value=IIF(curprn.npp#0.AND.nIt=0,curprn.npp,'')                           && ����� �� �������                    
        .Cells(rowcx,2).Value=curprn.fio                                                          && ���                    
        .Cells(rowcx,3).Value=IIF(SEEK(curprn.kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')           && ���������
                 
        .Cells(rowcx,4).Value=IIF(SEEK(curprn.kv,'sprkval',1),ALLTRIM(sprkval.name),'')           && ���������                                   
                
        .Cells(rowcx,5).NumberFormat = "General"
        .Cells(rowcx,5).Value=IIF(nIt=0,LEFT(staj_Tar,2)+'_'+SUBSTR(staj_tar,4,2)+'_'+SUBSTR(staj_tar,7,2),'')  && ����                
                                      
        .Cells(rowcx,6).Value=IIF(curprn.kf#0,curprn.kf,'')                                       && ������
                
        .Cells(rowcx,7).Value=IIF(curprn.namekf#0,curprn.namekf,'')                               && ���
        .Cells(rowcx,7).NumberFormat = "#####.00"                       
                
        .Cells(rowcx,8 ).Value=IIF(curprn.tokl#0,curprn.tokl,'')                                  && �����                  
        .Cells(rowcx,8).NumberFormat = "#####.00"  
                  
                
        SELECT curFondPrn
        GO TOP
        columncx=9
        DO WHILE !EOF()
           DO CASE 
              CASE !EMPTY(persved).AND.!EMPTY(sumved)
                   expers=ALLTRIM(persved)
                   exsum=ALLTRIM(sumved)
                   SELECT curprn                                                                                                                                                    
                   .cells(rowcx,columncx).Value=IIF(!EMPTY(&expers),LTRIM(STR(&expers))+'%','')
                   .cells(rowcx,columncx+1).Value=IIF(!EMPTY(&exsum),STR(&exsum,8,2),'')
                   columncx=columncx+2                
              CASE !EMPTY(sumved).AND.EMPTY(persved)
                   exsum=ALLTRIM(sumved)
                   SELECT curprn                                                         
                   .cells(rowcx,columncx).Value=IIF(!EMPTY(&exsum),STR(&exsum,8,2),'')
                   columncx=columncx+1                
           ENDCASE  
           SELECT curFondprn
           SKIP              
        ENDDO                                                                                                                       
        SELECT curprn
        rowcx=rowcx+1                 
        SELECT curprn
        SKIP
       DO fillpercent WITH 'fSupl'
     ENDDO  
     .Range(.Cells(2,1),.Cells(rowcx-1,maxColumn)).Select
     objExcel.Selection.Font.Name='Times New Roman' 
     objExcel.Selection.Font.Size=8      
     objExcel.Selection.WrapText=.T.  
*     .cells(6,1).HorizontalAlignment=xlRight
*     .cells(7,1).HorizontalAlignment=xlRight
     
     objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
     objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
     objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
     objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
     objExcel.Selection.VerticalAlignment=1   
     .Cells(2,1).Select      
ENDWITH
DO endPrnToExcel WITH 'fSupl'     
objExcel.Visible=.T.
*************************************************************************************************************
*       ������ ������ �����������
*************************************************************************************************************
PROCEDURE spisprn
IF USED('curTarJob')
   SELECT curTarJob
   USE
ENDIF 
itkat=formbase.log_it
=AFIELDS(arPeople,'datJob')
CREATE CURSOR curPrn FROM ARRAY arPeople
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN nit N(1)
*ALTER TABLE curPrn ADD COLUMN nkat C(200)
SELECT sprkat
* maxkat=���-�� ��������� ���������
* sumPodrTot - ���� �� ���������
* sum_podr - ���� �� ���������� � ������� ���������
* sumTot - ���� �����
* sum_podr - ���� �� ���������� � ������� ���������
COUNT TO maxKat
DIMENSION sumPodrTot(1,2),sum_podr(maxKat,2),sumTot(2,2),sum_tot(maxKat,2)
STORE 0 TO sumPodrTot,sumTot,sum_tot,sum_podr,numrecrep,numpage
SELECT * FROM datjob INTO CURSOR curTarJob READWRITE
SELECT curTarJob
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
INDEX ON STR(np,3)+STR(nd,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 2
fltch=''
*fltfield=par_field
DO fltstructure WITH 'kse#0.AND.kp#0.AND.kd#0','curTarJob'
DO selectfromrasp
SELECT curTarJob
SET ORDER TO 1
GO TOP
kp_cx=kp
kpp_cx=kodkpp
kpp1_cx=kodkpp1
repField=fSupl.colprn(8)
repSum=formbase.pl_sum
npp_cx=0
DO WHILE !EOF()     
   SCATTER TO ac
   SELECT curPrn
   npp_cx=npp_cx+1
   APPEND BLANK
   GATHER FROM ac                           
   REPLACE npp WITH npp_cx
   log_sumdopl=.F.     
   sumPodrTot(1,1)=sumPodrTot(1,1)+kse   
   sumTot(1,1)=sumTot(1,1)+kse   
   IF itKat
      sum_podr(ASCAN(kod_kat,curPrn.kat),1)=sum_podr(ASCAN(kod_kat,curPrn.kat),1)+kse
   ENDIF
   sum_tot(ASCAN(kod_kat,curPrn.kat),1)=sum_tot(ASCAN(kod_kat,curPrn.kat),1)+kse    
  
   SELECT curTarJob   
   SKIP
 
   IF kp_cx#curTarJob.kp   
      npp_cx=0         
      SELECT curPrn
      APPEND BLANK      
      REPLACE nIt WITH 1,kp WITH kp_cx,kd WITH 999,fio WITH '�����',kse WITH sumPodrTot(1,1)                                                         
      IF formbase.log_it
         FOR i=1 TO maxKat
             APPEND BLANK 
             REPLACE nIt WITH 2,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH IIF(SEEK(kat,'sprkat',1),sprkat.name,''),kse WITH sum_podr(i,1)
                                                                   
         ENDFOR      
      ENDIF      
      *--------
      STORE 0 TO sumPodrTot
      STORE 0 TO sum_podr
      SELECT curTarJob
      kp_cx=kp
   ENDIF       
   *--------             
   SELECT curTarJob   
ENDDO
*------����� ����-----------------
SELECT curPrn
APPEND BLANK      
REPLACE nIt WITH 7,kp WITH kp_cx,kd WITH 999,fio WITH '�����',kse WITH sumTot(1,1)  
*-----����� ���� �� ���������� ���������
IF formbase.log_it
   FOR i=1 TO maxKat
       APPEND BLANK 
       REPLACE nIt WITH 8,kp WITH kp_cx,kd WITH 999,kat WITH kod_kat(i),fio WITH IIF(SEEK(kat,'sprkat',1),sprkat.name,''),kse WITH sum_tot(i,1)                  
   ENDFOR
ENDIF     
SELECT curPrn
DELETE FOR kse=0
GO TOP 
DO procForPrintAndPreview WITH 'repspisoknew','������ �����������'
SELECT people 
*SET ORDER TO &ord_old
SELECT rasp

*----------------------------------------------------------------------------------------------------------------------------
*            �������� � ��������� ������ ����� ������ ��� �����������
*----------------------------------------------------------------------------------------------------------------------------
PROCEDURE setupreport

DO CASE 
   CASE datset.nformat=1  && A4
        IF !USED('fondprn')
           IF FILE(pathcur+'fondprn4.dbf')
              fUse=pathcur+'fondprn4.dbf' 
              USE &fUse ORDER 1 IN 0 ALIAS fondprn
           ELSE
              USE fondprn4 ORDER 1 IN 0 ALIAS fondprn
           ENDIF
        ENDIF 
   CASE datset.nformat=2 && A3
        IF !USED('fondprn')
           IF FILE(pathcur+'fondprn3.dbf')
              fUse=pathcur+'fondprn3.dbf' 
              USE &fUse ORDER 1 IN 0 ALIAS fondprn
           ELSE
              USE fondprn3 ORDER 1 IN 0 ALIAS fondprn
           ENDIF
        ENDIF 
ENDCASE 


COPY FILE reptardop.frt TO reptarnew.frt
COPY FILE reptardop.frx TO reptarnew.frx

SELECT 0
USE reptarnew.frx

LOCATE FOR objType=9.AND.ALLTRIM(LOWER(comment))='opheader'
REPLACE height WITH height+margTop*300
REPLACE height WITH height+margTop*300 FOR ALLTRIM(LOWER(comment))=='lhead'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='olong1'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='olong3'
REPLACE vpos WITH vpos+margTop/2*300 FOR ALLTRIM(LOWER(comment))=='headsup'

REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='wlong'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='mdetail'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='mdt3'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='itogdetail'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='itogcom'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='ldetail'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='ldetailitog'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='lpodr'
REPLACE vpos WITH vpos+margTop*300 FOR ALLTRIM(LOWER(comment))=='rpodr'


REPLACE height WITH height+nlspace*300 FOR ALLTRIM(LOWER(comment))=='ldetail'
REPLACE vpos WITH vpos+nlspace*300 FOR ALLTRIM(LOWER(comment))=='olong3'
REPLACE vpos WITH vpos+nlspace*300 FOR ALLTRIM(LOWER(comment))=='ldetailitog'
REPLACE vpos WITH vpos+nlspace*300 FOR ALLTRIM(LOWER(comment))=='itogdetail'
REPLACE vpos WITH vpos+nlspace*300 FOR ALLTRIM(LOWER(comment))=='itogcom'

LOCATE FOR objType=9.AND.objCode=0

LOCATE FOR objtype=6.AND.comment='lheadv'                  &&  ���������� ������� ������������ ����� � ����� 

leftline=hpos
leftobj=hpos+104.67*2
SELECT reptarnew                                           &&������ ����� �����
LOCATE FOR comment='lheadh'
htop=vpos  

LOCATE FOR objtype=8.AND.comment='lab1'                   && ���������� ������� ���������� � �����   
SCATTER TO dim_lab                  

LOCATE FOR objtype=8.AND.ALLTRIM(comment)='labsup'        && ���������� ���������� ������ � �������� � �����   
SCATTER TO dim_labsup                  
DELETE 

LOCATE FOR objtype=8.AND.comment='labpers'                && ���������� ���������� %, ����� � �����   
SCATTER TO dim_labpers                  
DELETE 

LOCATE FOR objtype=6.AND.comment='line1'                  && ���������� ������� ������������ ����� ��������   
SCATTER TO dim_line   

LOCATE FOR objtype=6.AND.comment='linesup'                 && ���������� ������������ ����� �������� ������ � �������� � �����   
SCATTER TO dim_linesup                  
DELETE 

LOCATE FOR objtype=6.AND.comment='lhpers'                 && ���������� �������������� ����� �������-�����
SCATTER TO dim_lhpers                  
DELETE 

LOCATE FOR objtype=6.AND.comment='lvpers'                 && ���������� ������������ ����� �������-�����
SCATTER TO dim_lvpers                  
DELETE 

LOCATE FOR objtype=6.AND.ALLTRIM(comment)=='lhead'        && ���������� ������������ ����� ��������������� ���������
SCATTER TO dim_lhead                  
DELETE 

LOCATE FOR objtype=8.AND.comment='headsup'                && ���������� ������������ ���������������� ���������
SCATTER TO dim_headsup                  

LOCATE FOR objtype=8.AND.ALLTRIM(comment)='mdetail'       && ���������� ����� detail
SCATTER TO dim_mdetail                  
DELETE 

LOCATE FOR objtype=8.AND.ALLTRIM(comment)='mdt3'          && ���������� ����� detail ��� 3-�� ����
SCATTER TO dim_detail3
DELETE 

LOCATE FOR objtype=8.AND.ALLTRIM(comment)='itogdetail'    && ���������� ����� detail-�����
SCATTER TO dim_itogdetail                  
DELETE 

LOCATE FOR objtype=6.AND.comment='ldetail'                && ���������� ������������ detail
SCATTER TO dim_ldetail                  
DELETE 

LOCATE FOR objtype=6.AND.ALLTRIM(comment)='ldetailitog'   && ���������� ������������ detail � ������
SCATTER TO dim_ldetailItog        
DELETE 

ncolch=0
SELECT fondprn
SET FILTER TO itog.AND.logVed
COUNT TO maxItog
DIMENSION nSum(maxItog)
GO TOP 
FOR i=1 TO maxitog
    nSum(i)=ALLTRIM(expr1)
    SKIP 
ENDFOR
rowTrf=1
SET FILTER TO logVed
GO TOP
DO WHILE !EOF()
   SELECT reptarnew
   IF !EMPTY(fondprn.exprlab)       
      DO CASE
         CASE !EMPTY(fondprn.primlab)
              LOCATE FOR comment=ALLTRIM(fondprn.primlab)  
              IF FOUND()                                                                                                   
                 REPLACE expr WITH ALLTRIM(fondprn.exprlab),width WITH fondprn.colwidth*104.67,hpos WITH leftobj,fontsize WITH fondprn->fhsize,fontface WITH 'Times New Roman'                      
                 SELECT fondprn
                 REPLACE hpos WITH reptarnew.hpos,lhpos WITH leftline                 
                 SELECT reptarnew
                 leftline=hpos+width+104.67*2
                 LOCATE FOR comment=ALLTRIM(fondprn.primline)                
                 REPLACE hpos WITH leftline
                 leftobj=hpos+104.67*2                
              ENDIF   
         CASE EMPTY(fondprn.primlab)
              APPEND BLANK          
              IF fondprn.log_sv.OR.fondprn.log_kv                   
                 GATHER FROM dim_labsup
              ELSE 
                 GATHER FROM dim_lab
              ENDIF    
              REPLACE expr WITH ALLTRIM(fondprn.exprlab),width WITH fondprn.colwidth*104.67,hpos WITH leftobj,fontsize WITH fondprn->fhsize,fontface WITH 'Times New Roman'                                 
              SELECT fondprn
              REPLACE hpos WITH reptarnew.hpos,lhpos WITH leftline                 
              SELECT reptarnew
              leftline=hpos+width+104.67*2
              APPEND BLANK 
              DO CASE 
                 CASE fondprn.tline=1
                      GATHER FROM dim_line
                 CASE fondprn.tline=2
                 CASE fondprn.tline=3
                      GATHER FROM dim_linesup
              ENDCASE
              REPLACE hpos WITH leftline
              leftobj=hpos+104.67*2               
              IF fondprn.ldouble
                 *------- �������������� �����
                 APPEND BLANK
                 GATHER FROM dim_lhpers
                 REPLACE hpos WITH fondprn.hpos-104.67*2,width WITH (fondprn.colWidth+4)*104.67  
                 *------- ������� %
                 APPEND BLANK
                 GATHER FROM dim_labpers    
                 REPLACE expr WITH "'%'",width WITH 104.67*IIF(fondprn.colWidth1>0,fondprn.colWidth1,20),hpos WITH fondprn.lhpos+104.67,fontsize WITH fondprn->fhsize,fontface WITH 'Times New Roman'                                     
                 *------- ������������ �����
                 APPEND BLANK
                 GATHER FROM dim_lvpers
                 REPLACE hpos WITH fondprn.hpos+104.67*IIF(fondprn.colWidth1>0,fondprn.colWidth1,20)
                 supleft=hpos
                 SELECT fondprn
                 REPLACE shpos WITH supleft
                 SELECT reptarnew
                 *------- ������� �����
                 APPEND BLANK
                 GATHER FROM dim_labpers    
                 REPLACE expr WITH "'�����'",width WITH leftline-supleft-104.67,hpos WITH supleft+104.67*2,fontsize WITH fondprn->fhsize,fontface WITH 'Times New Roman'                                                 
               ENDIF                      
      ENDCASE  
      ncolch=ncolch+1
      SELECT reptarnew
      APPEND BLANK
      GATHER FROM dim_lhead
      REPLACE hpos WITH fondprn.lhpos,comment WITH 'lhead'
      APPEND BLANK
      GATHER FROM dim_headsup
      REPLACE expr WITH "'"+LTRIM(STR(ncolch))+"'",fontsize WITH 8,fontface WITH 'Times New Roman'                                      
      REPLACE width WITH IIF(!fondprn.ldouble,fondprn.colwidth*104.67,fondprn.shpos-fondprn.lhpos-104.67*2),hpos WITH fondprn.hpos,fillchar WITH 'C',offset WITH 2  
      
      IF fondprn.ldouble
         ncolch=ncolch+1
         APPEND BLANK
         GATHER FROM dim_lhead
         REPLACE hpos WITH fondprn.shpos
         APPEND BLANK
         GATHER FROM dim_headsup
         REPLACE expr WITH "'"+LTRIM(STR(ncolch))+"'",fontsize WITH 8,fontface WITH 'Times New Roman'                                      
         REPLACE width WITH fondprn.rhpos-fondprn.shpos-104.67,hpos WITH fondprn.shpos+104.67,fillchar WITH 'C',offset WITH 2  
      ENDIF
      
      *-----detail
      DO CASE
         CASE fondprn.texpr=1
              IF !fondprn.ldouble    && ��� ��������� �������
                  APPEND BLANK 
                  GATHER FROM dim_mdetail
                  REPLACE expr WITH fondprn.expr1,width WITH fondprn.colwidth*104.67,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman',supexpr WITH 'curPrn.nIt=0'
                  IF ALLTRIM(fondprn.har1)='N'
                    REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 1     
                  ENDIF
                  *---�����--
                   APPEND BLANK 
                  GATHER FROM dim_itogdetail
                  REPLACE expr WITH fondprn.expr1,width WITH fondprn.colwidth*104.67,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman',supexpr WITH 'curPrn.nIt>0'
                  IF ALLTRIM(fondprn.har1)='N'
                    REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 1     
                  ENDIF
                  *-----             
              ELSE                   && ��������� �������   
                  APPEND BLANK 
                  GATHER FROM dim_mdetail
                  REPLACE expr WITH fondprn.expr2,width WITH fondprn.shpos-fondprn.lhpos-104.67*2,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman'           
                  REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2
                   *---�����--
                  APPEND BLANK 
                  GATHER FROM dim_itogdetail
                  REPLACE expr WITH fondprn.expr2,width WITH fondprn.shpos-fondprn.lhpos-104.67*2,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman'           
                  REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2,supexpr WITH 'curPrn.nIt>0'                                   
                  *-------------------------
                  
                  APPEND BLANK 
                  GATHER FROM dim_mdetail
                  REPLACE expr WITH fondprn.expr1,width WITH fondprn.rhpos-fondprn.shpos-104.67*2,hpos WITH fondprn.shpos+104.67,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman'
                  REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2
                  *---�����--
                  APPEND BLANK 
                  GATHER FROM dim_itogdetail
                  REPLACE expr WITH fondprn.expr1,width WITH fondprn.rhpos-fondprn.shpos-104.67*2,hpos WITH fondprn.shpos+104.67,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman'
                  REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2,supexpr WITH 'curPrn.nIt>0'
                  *------------
              ENDIF     
         CASE fondprn.texpr=2 
         CASE fondprn.texpr=3
              APPEND BLANK 
              GATHER FROM dim_mdetail
              REPLACE expr WITH fondprn.expr1,width WITH fondprn.colwidth*104.67,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman',supexpr WITH 'curPrn.nIt=0'
              IF ALLTRIM(fondprn.har1)='N'
                 REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2 
              ENDIF
              *---�����--
              APPEND BLANK 
              GATHER FROM dim_itogdetail
              REPLACE expr WITH fondprn.expr1,width WITH fondprn.colwidth*104.67,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman',supexpr WITH 'curPrn.nIt>0'
              IF ALLTRIM(fondprn.har1)='N'
                 REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2 
              ENDIF
              *-----
              APPEND BLANK 
              GATHER FROM dim_detail3
              REPLACE expr WITH fondprn.expr2,width WITH fondprn.colwidth*104.67,hpos WITH fondprn.hpos,fontsize WITH fondprn->fdsize,fontface WITH 'Times New Roman',supexpr WITH 'curPrn.nIt=0'
              IF ALLTRIM(fondprn.har2)='N'
                 REPLACE picture WITH '"@Z"',fillchar WITH 'N',stretch WITH .F.,offset WITH 2,supexpr WITH 'curPrn.nIt=0'     
              ENDIF
              IF !EMPTY(fondprn.nAl)
                 REPLACE offset WITH fondprn.nAl
              ENDIF
              
      ENDCASE 
      APPEND BLANK
      GATHER FROM dim_ldetail
      REPLACE hpos WITH fondprn.lhpos,supexpr WITH 'curPrn.nIt=0'
      vcxpos=vpos
      APPEND BLANK
      GATHER FROM dim_ldetailItog   
      REPLACE hpos WITH fondprn.lhpos,supexpr WITH 'curPrn.nIt>0.AND.curPrn.nit<9'
      *REPLACE height WITH margTop*104.67
      IF rowTrf>2.AND.rowTrf<6
         REPLACE comment WITH 'itog'         
      ENDIF
    
      IF fondprn.ldouble
         APPEND BLANK
         GATHER FROM dim_ldetail
         REPLACE hpos WITH fondprn.shpos,supexpr WITH 'curPrn.nIt=0'
         APPEND BLANK
         GATHER FROM dim_ldetailItog
         REPLACE hpos WITH fondprn.lhpos,supexpr WITH 'curPrn.nIt>0'  
      
      ENDIF      
   ENDIF
   SELECT fondprn   
   REPLACE rhpos WITH leftline    
   SKIP 
   rowTrf=rowTrf+1
ENDDO

SELECT fondprn
LOCATE FOR log_sv
begline=lhpos
LOCATE FOR log_kv
endlab=lhpos
SCAN WHILE log_kv
ENDSCAN
endline=lhpos
SELECT reptarnew
LOCATE FOR comment='lstim'
REPLACE hpos WITH begline,width WITH endline-begline
LOCATE FOR comment='labstim'
REPLACE hpos WITH begline+104.67*2,width WITH endlab-begline-104.67*4
LOCATE FOR comment='labkomp'
REPLACE hpos WITH endlab+104.67*3,width WITH endline-endlab-104.67*6
SELECT fondprn
GO TOP
begline=lhpos
GO BOTTOM 
endline=rhpos
SELECT reptarnew
DELETE FOR LOWER(ALLTRIM(comment))='mdt3'
DELETE FOR LOWER(ALLTRIM(comment))='itogdetail'
REPLACE width WITH endline-begline+104.67 FOR comment='olong'.OR.comment='lheadh'
APPEND BLANK 
GATHER FROM dim_lhead
REPLACE hpos WITH endline
APPEND BLANK 
GATHER FROM dim_ldetail
REPLACE hpos WITH endline,supexpr WITH 'nIt=0'
APPEND BLANK 
GATHER FROM dim_ldetailitog
REPLACE hpos WITH endline,supexpr WITH 'nIt>0.AND.nit<9'
REPLACE hpos WITH endline FOR comment='lpodr'

LOCATE FOR 'namedol'$LOWER(expr)
newWidthCol=width
LOCATE FOR 'sprkval'$LOWER(expr)
newWidthCol=newWidthCol+width
LOCATE FOR 'staj_tar'$LOWER(expr)
newWidthCol=newWidthCol+width
LOCATE FOR 'kf'$LOWER(expr)
newWidthCol=newWidthCol+width
REPLACE width WITH width+newWidthCol FOR 'fio'$expr.AND.'nit>0'$LOWER(supexpr)
DELETE FOR ALLTRIM(comment)='itog'.AND.'nit>0'$LOWER(supexpr)
USE
**********************************************************************************************************************************
PROCEDURE oldnewprn
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF 
IF USED('peopold')
   SELECT peopold
   USE
ENDIF   
RESTORE FROM pathOld ADDITIVE
peopold=pathold+'people.dbf'

USE &peopold ALIAS peopold IN 0
SELECT * FROM datjob INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN sumOld N(12,2)
ALTER TABLE curPrn ADD COLUMN sumMiss N(12,2)
ALTER TABLE curPrn ADD COLUMN persNew N(7,2)
ALTER TABLE curPrn ADD COLUMN sumNew N(7,2)
ALTER TABLE curPrn ADD COLUMN npp N(3)
SELECT curPrn
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.np,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL
INDEX ON STR(np,3)+STR(ND,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
DO fltstructure WITH 'mTokl#0','curprn'
SET ORDER TO 1
GO TOP
nppcx=1
kpOld=kp
DO WHILE !EOF()     
   SELECT peopold
   LOCATE FOR tab=curprn.tabn.AND.kp=curprn.kp.AND.kse=curprn.kse.AND.kd=curprn.kd.AND.tr=curprn.tr
   IF !FOUND()
      LOCATE FOR ALLTRIM(name)=ALLTRIM(curprn.fio).AND.kp=curprn.kp.AND.kse=curprn.kse.AND.kd=curprn.kd.AND.tr=curprn.tr
   ENDIF
   SELECT curPrn
   REPLACE sumOld WITH peopold.sfond,persNew WITH pSlWork+sHigh,sumMiss WITH mslwork+mHigh,sumNew WITH msf-mslwork-mHigh,npp WITH nppcx
   nppcx=nppcx+1
   SKIP
   IF kpOld#kp
      kpOld=kp
      nppcx=1 
   ENDIF     
ENDDO 
DO procForPrintAndPreview WITH 'repOldNew','��������������� ������'
SELECT peopold
USE
SELECT curPrn
USE
SELECT people
*************************************************************************************
*    ������ ��������������� ����������
*************************************************************************************
PROCEDURE agreeAddPrn
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF  
SELECT * FROM datJob INTO CURSOR curPrn READWRITE 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,np) ALL
REPLACE nD WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,nd) ALL
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 

DO fltstructure WITH 'mTokl#0','curprn'
stFio=''
st1='1.  �������� �.8 ��������� ��������� � ��������� ��������:'
st8='8. ��������� ��������������� ��������� ������� ������ �����:'
st81=''
st82=''
st84=''
st86=''
st87=''
st88=''
st83="8.3. �������� �� ���������� (������� � ����������) �����, ������� � ���������� ������� �������� ����������� ������  -  ������� ������ (�� �����);"

st85='8.5. �������� ������-������������, ��������� ������� ��������� � �������������� �������, ����������� ����������� ������ � ������������ ��������  40% ������;'

st89='8.9. ������� �� ������ �� ������� � (���) ������� �������� ����� �� ����������� ���������� ������� ����.  ��� ������������ ������� �� ������ �������� ����� (�.8.8.) ������� �� ������ �� ������� � (���) ������� �������� ����� (�.8.9.)  �� �������������;' 
st90='9.0. ������������� �������: ������, �������� �� ��������� � ������������� ������, �������� �� ������� ���������� � �����  - � ��������� �� ������  �������� ������������ ��������� �� ������ �����;'
st91='9.1. �������������� ������� �� ������������, ������������ ������ -  �������� ������������ ��������� �� ������ �����; '
st14='14. ���������� ������������� ��������� �������������� ���� �������������� �����:'
st141=''
st142='14.2. ��������������  ������ �� ������ �� ���������  _______ �����������   ���� (����������� ���). '
st2='2. � ���������  ����� �������� ��������   �������� ��� ���������.'
st3='3. ��������� �������������� ���������� �������� ������������ ������ ��������� � �������������� ���� �������� �� ��������������, ���������  � 01 ������ 2020 ����.'
st4='4. ��������� �������������� ���������� ���������� � ���� �����������, ���� �� ������� �������� � ���������,  ������  - � ����������.'
signn='_______________�.�.�������'
signr='_______________�.�.�������'
fiosign=''
SELECT curPrn
GO TOP
DO procForPrintAndPreview WITH 'agreeadd','�������������� ����������'
SELECT rasp
*************************************************
PROCEDURE procstAdd
PARAMETERS par1
strFio=LEFT(ALLTRIM(curPrn.Fio),AT(' ',ALLTRIM(curprn.fio)))
strFio=UPPER(strFio)+SUBSTR(ALLTRIM(curPrn.fio),AT(' ',curprn.fio)+1)
stFio='          ���������� ��������������� "��������� ����������� ��������� ��������" � ���� �������� ����� ������� ����� ��������� (����� - ����������), ������������ �� ��������� ������ � '+;
       strFio+' (����� - ��������)'
stFio=stFio+' �� ���������� ����� ���������� ���������� �������� �� 18.01.2019�. � 27 "�� ������ ����� ���������� ��������� �����������" � �����������,'
stFio=stFio+' �� ��������� ����� ��������� ������ 19 ��������� ������� ���������� ��������, ���������� � ������������ � ������������ �����������������, ��������� ��������� ���������� � ��������� � �������������:'
st81='8.1. �����   � ������� '+LTRIM(STR(curprn.mtokl,8,2))+' ���. �� ���� ���������� ����������. � ���������� ����� ���������� � ������������ � �����������������;'
st82='8.2. �������� �� ���� ������ � ��������� ����������� '+IIF(stpr#0,LTRIM(STR(stpr,3)),'________')+' % ������� ������;' 

st84='8.4. �������� �� ��������� ������ � ����� ���������������   ����������� � ���������������� ���������� '+IIF(pkat#0,LTRIM(STR(pkat,3)),'________')+' % ������;'

st86='8.6. ������� �� ���������� ��������������-���������������� �������'+IIF(pmain#0.OR.pmain2#0,LTRIM(STR(curprn.pmain+curprn.pmain2,3)),'________')+' % ������� ������;' 

st87='8.7. �������� �� ����������� ���������������� ������������: ���������������� ���������, ����������-�������� ���������� ������������ '+IIF(posob#0,LTRIM(STR(curprn.posob,3)),'________')+' % ������; '

st88='8.8. ������� �� ������ �������� ����� '+IIF(pcharw#0,LTRIM(STR(pcharw,3)),'________')+' % ������;'
st141='14.1. �������� �� ������ �� �������� ��������� � ������� '+IIF(pkont#0,LTRIM(STR(curprn.pkont,3))+'%;','________%;')
fiosign=SUBSTR(ALLTRIM(fio),AT(' ',ALLTRIM(fio))+1,1)+'.'+SUBSTR(ALLTRIM(fio),RAT(' ',ALLTRIM(fio))+1,1)+'.'+LEFT(ALLTRIM(fio),AT(' ',ALLTRIM(fio)))

signr='_______________'+fiosign

*************************************************************************************
*    ������ ��������������� ����������
*************************************************************************************
PROCEDURE stajperprn
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
SELECT * FROM datjob INTO CURSOR curprn READWRITE
SELECT curprn
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
DO fltstructure WITH '!EMPTY(per_date)','curprn'
SELECT curPrn
INDEX ON DTOS(per_date) TAG T1
GO TOP
DO procForPrintAndPreview WITH 'repperstaj','����������� ����'
SELECT people
**************************************************************************************
*   ������� + ��������
**************************************************************************************
PROCEDURE raspPeop
IF USED('curprn')
   SELECT curPrn
   USE
ENDIF
SELECT * FROM sprkat INTO CURSOR curSuplKat READWRITE
CREATE CURSOR curPrn (np N(3),nd N(3),kp N(3),kd N(3),named C(150), fio C(70),kse N(7,2),tr N(1),nametr C(15),kat N(1),logEnd L,logBeg L,kodKpp N(3),KodKpp1 N(3),ksesh N(6,2),npp N(3))
SELECT * FROM datjob INTO CURSOR curTarPeople READWRITE
SELECT curTarPeople
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL

INDEX ON STR(np,3)+STR(nd,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 2
itKat=formbase.log_it
kserasp_kat=0
kserasp_cx=0
fltch=''

SELECT rasp      
SELECT * FROM rasp INTO CURSOR currasp READWRITE      
SELECT currasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1
*DO fltstructure WITH 'kd>0.AND.kp>0','curTarPeople'
SELECT curRasp
kseRasp=0
kseVac=0
ksePeop=0
SCAN ALL
     kseRasp=kse
     ksePeop=0
     SELECT curTarPeople
     IF SEEK(STR(curRasp.kp,3)+STR(curRasp.kd,3))
        DO WHILE kp=curRasp.kp.AND.kd=curRasp.kd
                 SELECT curPrn
                 APPEND BLANK 
                 REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kp WITH curRasp.kp,kd WITH curRasp.kd,kat WITH curRasp.kat,kse WITH curTarPeople.kse,tr WITH curTarPeople.tr,;
                         fio WITH curTarPeople.fio,nametr WITH IIF(SEEK(tr,'sprtype',1),sprtype.name,''),KodKpp WITH curRasp.KodKpp,KodKpp1 WITH curRasp.KodKpp1              
                 ksePeop=ksePeop+curPrn.kse
                 SELECT curTarPeople
                 SKIP
        ENDDO          
     ENDIF
     IF kseRasp-ksePeop>0  
        SELECT curPrn
        APPEND BLANK 
        REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kp WITH curRasp.kp,kat WITH curRasp.kat,kd WITH curRasp.kd,kse WITH kseRasp-ksePeop,fio WITH '���������',;
                KodKpp WITH curRasp.KodKpp,KodKpp1 WITH curRasp.KodKpp1,tr WITH 1 
     ENDIF
     SELECT curRasp
ENDSCAN
SELECT curPrn
INDEX ON STR(np,3)+STR(nd,3)+fio+STR(tr,1) TAG T1
INDEX ON kp TAG T2
SET ORDER TO 1
SELECT curRasp
SCAN  ALL
      SELECT curPrn
      IF SEEK(STR(curRasp.np,3)+STR(curRasp.nd,3))
         REPLACE named WITH IIF(SEEK(curPrn.kd,'sprdolj',1),sprdolj.name,'')
      ENDIF
      SELECT curRasp
ENDSCAN
SELECT curPrn

DO fltstructure WITH 'kd>0.AND.kp>0','curPrn'
*DO CASE
*   CASE setprn(1)=1        
*        SET FILTER TO kp>0.AND.kd>0         
*   CASE setprn(1)=2          
*        fltch=ALLTRIM(datagrup->sostav1)        
*        SET FILTER TO ','+LTRIM(STR(kp))+','$fltch.AND.kp>0.AND.kd>0         
*   CASE setprn(1)=3
*        DO procflt_str  
*        SELECT curPrn                       
*        flt_Str=flt_str+',999,'
*        SET FILTER TO ','+LTRIM(STR(kp))+','$flt_str.AND.kp>0.AND.kd>0            
*   CASE setprn(1)=4                    

*ENDCASE 
kdold=0
kpold=0
nppnew=0
SCAN ALL 
     nppnew=nppnew+1 
     IF kp#kpold.OR.kd#kdold
        kdold=kd
        kpold=kp
        nppnew=1
     ENDIF
     REPLACE npp WITH nppnew
     IF npp=1 
        kseshcx=IIF(SEEK(STR(curPrn.np,3)+STR(curPrn.nd,3),'rasp',1),rasp.kse,0)
        REPLACE ksesh WITH kseshcx    
     ENDIF 
ENDSCAN
GO TOP
 kpOld=kp
kdOld=kd
DO WHILE !EOF()
   SKIP
   DO CASE
      CASE kd#kdOld
           SKIP-1
           REPLACE logEnd WITH .T.
           SKIP 
           kpOld=kp
           kdOld=kp
      CASE kp#kpOld
           SKIP-1
           REPLACE logEnd WITH .T.
           SKIP 
           kpOld=kp
           kdOld=kd   
                     
   ENDCASE
ENDDO
SELECT curRasp
SET ORDER TO 2
SET FILTER TO nd=1
GO TOP
DO WHILE !EOF()
   SELECT rasp
   SUM kse TO kserasp_cx FOR kp=curRasp.kp  
   SELECT curPrn
   SUM kse TO kse_cx FOR kp=curRasp.kp  
   IF kse_cx#0
      SELECT curprn
      APPEND BLANK
      REPLACE kp WITH curRasp.kp,kd WITH 999,kse WITH kse_cx,np WITH curRasp.np,nd WITH 70,named WITH '�����',logEnd WITH .T.,ksesh WITH kserasp_cx 
      IF itKat
         SELECT curSuplKat
         SCAN ALL 
              SELECT rasp
              SUM kse TO kserasp_kat FOR kp=curRasp.kp.AND.kat=curSuplKat.kod      
              SELECT curPrn
              SUM kse TO kse_kat FOR kp=curRasp.kp.AND.kat=curSuplKat.kod
                      
              IF kse_kat#0 
                 SELECT curPrn
                 APPEND BLANK
                 REPLACE kp WITH curRasp.kp,kd WITH 999,kse WITH kse_kat,np WITH curRasp.np,nd WITH 70+curSuplKat.kod,named WITH curSuplKat.name,logEnd WITH .T.,ksesh WITH kserasp_kat  
              ENDIF  
              SELECT curSuplKat               
         ENDSCAN     
      ENDIF        
   ENDIF
   SELECT curRasp
   SKIP     
ENDDO

*---------------��� �������������������
SET FILTER TO nd=1.AND.kodKpp1>0
SET ORDER TO 1
GO TOP
kppOld=kodKpp1
kpOld=kp
npOld=np
DO WHILE !EOF()
   kpOld=kp
   npOld=np
   SELECT currasp   
   SKIP
   IF kppOld#KodKpp1
      SELECT curPrn
      SUM kse,ksesh TO kse_cx,kserasp_cx FOR kodKpp1=kppOld.AND.nd<80
      APPEND BLANK
      REPLACE kp WITH kpOld,kd WITH 999,kse WITH kse_cx,np WITH npOld,nd WITH 80,named WITH '�����',logEnd WITH .T.,ksesh WITH kserasp_cx 
      IF itKat
         SELECT curSuplKat
         SCAN ALL               
              SELECT curPrn
              SUM kse,ksesh TO kse_kat,kserasp_kat FOR kodKpp1=kppOld.AND.kat=curSuplKat.kod
              
              IF kse_kat#0 
                 SELECT curPrn
                 APPEND BLANK
                 REPLACE kp WITH kpOld,kd WITH 999,kse WITH kse_kat,np WITH npOld,nd WITH 80+curSuplKat.kod,named WITH curSuplKat.name,logEnd WITH .T.,ksesh WITH kserasp_kat  
              ENDIF  
              SELECT curSuplKat               
         ENDSCAN 
      ENDIF 
      SELECT curRasp
      kppOld=Kodkpp1
      npOld=np      
   ENDIF
   
ENDDO
*---------------��� ����������������
SET FILTER TO nd=1.AND.kodKpp>0
GO TOP
kppOld=kodKpp
kpOld=kp
npOld=np
DO WHILE !EOF()
   kpOld=kp
   npOld=np
   SELECT currasp   
   SKIP
   IF kppOld#KodKpp
      SELECT rasp 
     * SUM kse TO kserasp_cx FOR SEEK(STR(curprn.kp,3),'rasp',2).AND.rasp.kodKpp=kppOld
      SELECT curPrn    
      SUM kse,ksesh TO kse_cx,kserasp_cx FOR kodKpp=kppOld.AND.nd<70
      APPEND BLANK
      REPLACE kp WITH kpOld,kd WITH 999,kse WITH kse_cx,np WITH npOld,nd WITH 90,named WITH '�����',logEnd WITH .T.,ksesh WITH kserasp_cx 
      IF itKat
         SELECT curSuplKat
         SCAN ALL              
              SELECT curPrn
              SUM kse,ksesh TO kse_kat,kserasp_kat FOR kodKpp=kppOld.AND.kat=curSuplKat.kod
              IF kse_kat#0 
                 SELECT curPrn
                 APPEND BLANK
                 REPLACE kp WITH kpOld,kd WITH 999,kse WITH kse_kat,np WITH npOld,nd WITH 90+curSuplKat.kod,named WITH curSuplKat.name,logEnd WITH .T.,ksesh WITH kserasp_kat   
              ENDIF  
              SELECT curSuplKat               
         ENDSCAN 
      ENDIF 
      SELECT curRasp
      kppOld=Kodkpp 
      npOld=np      
   ENDIF   
ENDDO
SELECT curPrn
SUM kse,ksesh TO kse_cx,kserasp_cx FOR kd#999
APPEND BLANK 
REPLACE kp WITH 999,kd WITH 999,kse WITH kse_cx,np WITH 999,nd WITH 90,named WITH '�����',logEnd WITH .T.,ksesh WITH kserasp_cx

SELECT curSuplKat
SCAN ALL     
     SELECT curPrn
     SUM kse,ksesh TO kse_kat,kserasp_cx FOR kat=curSuplKat.kod.AND.kd#999
     IF kse_kat#0 
        SELECT curPrn
        APPEND BLANK
        REPLACE kp WITH 999,kd WITH 999,kse WITH kse_kat,np WITH 999,nd WITH 90+curSuplKat.kod,named WITH curSuplKat.name,logEnd WITH .T.,ksesh WITH kserasp_cx  
     ENDIF  
     SELECT curSuplKat               
ENDSCAN
SELECT curPrn   
GO TOP 
DO procForPrintAndPreview WITH 'repShtatPeop',''
SELECT people 
**************************************************************************************
PROCEDURE spiszpl
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF 
SELECT * FROM datjob INTO CURSOR curPrn READWRITE
SELECT curPrn
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
INDEX ON fio+STR(tabn,5) TAG T1
SET ORDER TO 1
fltch=''
DO fltstructure WITH 'kse#0.AND.kp#0.AND.kd#0','curPrn'
SELECT curPrn
SET ORDER TO 1
GO TOP 
DO procForPrintAndPreview WITH 'repspiszpl','������ ������� ��� ������������'
SELECT people
**********************************************************************************************
PROCEDURE spiszplToExcel
DO startPrnToExcel WITH 'fSupl'           
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=8
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
         
     rowcx=3     
    .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .Value='������ ����������� ��� ��������� �������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH        
     rowcx=rowcx+1                            
                      
     .cells(rowcx,1).Value='���.�'
     .cells(rowcx,2).Value='������� ��� ��������'         
     .cells(rowcx,3).Value='�����'
     .cells(rowcx,4).Value='���'                   
     .cells(rowcx,5).Value='�����'
     .cells(rowcx,6).Value='�����*�����'                    
     .cells(rowcx,7).Value='����'              
              
     rowcx=rowcx+1
     .cells(rowcx,1).Value='1'
     .cells(rowcx,2).Value='2'
     .cells(rowcx,3).Value='3'
     .cells(rowcx,4).Value='4'
     .cells(rowcx,5).Value='5'
     .cells(rowcx,6).Value='6'
     .cells(rowcx,7).Value='7'  
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     numberRow=rowcx+1  
     rowtop=numberRow         
     SELECT curPrn
     DO storezeropercent
     GO TOP
     kpold=0 
     SCAN ALL
          .Cells(numberRow,1).Value=IIF(tabn#0,tabn,'')
          .Cells(numberRow,2).Value=fio   
          .Cells(numberRow,3).Value=kse                                      
          .Cells(numberRow,3).NumberFormat='0.00'                                               
          .Cells(numberRow,4).Value=IIF(SEEK(tr,'sprtype',1),ALLTRIM(sprtype.name),'')                                                                        
          .Cells(numberRow,5).Value=tokl                      
          .Cells(numberRow,6).Value=mtokl
          .Cells(numberRow,7).NumberFormat='@'
          .Cells(numberRow,7).Value=LEFT(curprn.staj_tar,5)
           *.Cells(numberRow,7).Value=staj_tar
        
          numberRow=numberRow+1
          DO fillpercent WITH 'fSupl'
                  
     ENDSCAN      
     .Range(.Cells(4,1),.Cells(5,7)).Select
     WITH objExcel.Selection          
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=9
          .Borders(xlEdgeLeft).Weight=xlThin
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin
          .VerticalAlignment=1
     ENDWITH  
     
     .Range(.Cells(6,1),.Cells(numberRow-1,7)).Select
      WITH objExcel.Selection          
          .HorizontalAlignment=xlLeft
          .VerticalAlignment=1
          .WrapText=.T.
          .Font.Name='Times New Roman'   
          .Font.Size=9
          .Borders(xlEdgeLeft).Weight=xlThin
          .Borders(xlEdgeTop).Weight=xlThin            
          .Borders(xlEdgeBottom).Weight=xlThin
          .Borders(xlEdgeRight).Weight=xlThin
          .Borders(xlInsideVertical).Weight=xlThin
          .Borders(xlInsideHorizontal).Weight=xlThin
          .VerticalAlignment=1
     ENDWITH                       
     .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'              
objExcel.Visible=.T.  
**************************************************************************************
PROCEDURE oldNewTablePrn
IF USED('peopold')
   SELECT peopold
   USE
ENDIF
IF USED('fondold')
   SELECT fondold
   USE
ENDIF
IF USED('curnew')
   SELECT curnew
   USE
ENDIF
IF USED('curold')
   SELECT curold
   USE
ENDIF
IF USED('oldNewPrn')
   SELECT oldNewPrn
   USE
ENDIF
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF 
IF !FILE('pathold.mem')
   RETURN  
ENDIF
RESTORE FROM pathold ADDITIVE
peopold=pathold+'people.dbf'
fondold=pathold+'tarfond.dbf'
USE &peopold ALIAS peopold IN 0
USE &fondold ALIAS fondold IN 0
SELECT * FROM fondold WHERE !EMPTY(snew) INTO CURSOR curold READWRITE
ALTER TABLE curold ADD COLUMN colpers N(8,2)
ALTER TABLE curold ADD COLUMN colsum N(12,2)
SELECT curold
INDEX ON num TAG T1

SELECT * FROM tarfond WHERE nblock=2 INTO CURSOR curnew READWRITE
ALTER TABLE curnew ADD COLUMN colpers N(7,2)
ALTER TABLE curnew ADD COLUMN colsum N(10,2)
SELECT curnew
INDEX ON num TAG T1

SELECT * FROM datjob INTO CURSOR curPrn READWRITE
SELECT curPrn
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
DO fltstructure WITH 'mtokl#0','curPrn'

CREATE CURSOR oldNewPrn (kodpeop N(5),numpp N(7),fio C(60),kp N(3),kd N(3),kse N(4,2),tr N(1),recOld C(70),pOld N(3),sOld N(12,2),recNew C(100),pNew N(6,2),sNew N(12,2))

SELECT curprn
INDEX ON STR(kp,3)+fio TAG T1
GO TOP
num_cx=0
DO WHILE !EOF()
   SELECT peopOld
   LOCATE FOR tab=curprn.tabn.AND.kp=curprn.kp.AND.kse=curprn.kse.AND.kd=curprn.kd.AND.tr=curprn.tr
   IF !FOUND()
      LOCATE FOR ALLTRIM(LOWER(name))==ALLTRIM(LOWER(curprn.fio)).AND.kp=curprn.kp.AND.kse=curprn.kse.AND.kd=curprn.kd.AND.tr=curprn.tr
   ENDIF
   IF FOUND()
      num_cx=num_cx+1
      SELECT curold
      REPLACE colPers WITH 0,colsum WITH 0 ALL
      GO TOP 
      SCAN ALL
           IF !EMPTY(pnew)
              IF 'peopold'$pnew
                 repcol=ALLTRIM(pNew)              
              ELSE
                 repcol='peopold.'+ALLTRIM(pNew)              
              ENDIF
              REPLACE colpers  WITH &repcol
           ENDIF    
           IF !EMPTY(snew) 
              IF 'peopold'$sNew
                 repcol=ALLTRIM(sNew)              
              ELSE 
                repcol='peopold.'+ALLTRIM(snew)
              ENDIF               
              REPLACE colsum  WITH &repcol
           ENDIF    
      ENDSCAN
     
      GO TOP
      SELECT curnew
      REPLACE colPers WITH 0,colsum WITH 0 ALL
      GO TOP 
      SCAN ALL
           IF !EMPTY(fpers) 
              IF 'datjob'$LOWER(fPers)
                 repcol='curprn.'+SUBSTR(ALLTRIM(fpers),8)
              ELSE        
                 repcol=ALLTRIM(fpers)       
              ENDIF 
              REPLACE colpers WITH &repcol         
           ENDIF     
           IF !EMPTY(sum_fm)
              repcol='curprn.'+ALLTRIM(sum_fm)              
              REPLACE colsum  WITH &repcol         
           ENDIF          
           SELECT curnew
      ENDSCAN
      GO BOTTOM 
      REPLACE colsum WITH curprn.msf 


     SELECT curOld
     SCAN ALL
          SELECT oldNewPrn
          APPEND BLANK
          REPLACE kodpeop WITH curprn.kodpeop,numpp WITH num_cx,kp WITH curPrn.kp,kd WITH curPrn.kd,kse WITH curPrn.kse,tr WITH curPrn.kse
          REPLACE recOld WITH curOld.rec,pOld WITH curOld.colpers,sOld WITH curOld.colSum
          SELECT curOld 
     ENDSCAN 
     GO TOP 
     SELECT curNew
     GO TOP 
     SCAN ALL
          SELECT oldNewPrn
          LOCATE FOR EMPTY(recNew).AND.numpp=num_cx
          IF !FOUND()          
             APPEND BLANK
             REPLACE kodpeop WITH curprn.kodpeop,numpp WITH num_cx,kp WITH curPrn.kp,kd WITH curPrn.kd,kse WITH curPrn.kse,tr WITH curPrn.kse
          ENDIF    
          REPLACE recNew WITH curNew.rec,pNew WITH curNew.colpers,sNew WITH curNew.colSum     
          SELECT curNew
     ENDSCAN
      
   ENDIF 
   
   SELECT curPrn
   SKIP
ENDDO
SELECT oldNewPrn
*SET ORDER TO 1
GO TOP 
DO procForPrintAndPreview WITH 'NewOldTable','������ ������� ��� ������������'
SELECT people 
**************************************************************************************************************************
PROCEDURE prnRepRaspShtat
PARAMETERS par1
IF USED('curRaspPrn')
   SELECT curRaspPrn
   USE
ENDIF
SELECT sprkat
COUNT TO max_kat
SELECT rasp
DIMENSION dim_dopl(max_kat,5)
STORE 0 TO dim_dopl
SELECT datShtat
LOCATE FOR ALLTRIM(pathTarif)=pathTarSupl
topoffice=datShtat.office1
office_say=datShtat.office
adres_say=datShtat.adres  
=AFIELDS(arRasp,'rasp') 
SELECT * FROM rasp INTO CURSOR curRaspPrn READWRITE             
ALTER TABLE curraspprn ADD COLUMN namepodr C(100)
SELECT curRaspPrn
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE named WITH IIF(SEEK(curraspprn.kd,'sprdolj',1),sprdolj.name,'') ALL 
*REPLACE kodkpp WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.kodkpp,0) ALL
*REPLACE kodkpp1 WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.kodkpp1,0) ALL
REPLACE namepodr WITH IIF(SEEK(curraspprn.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.prim),sprpodr.prim,sprpodr.name),'') FOR nd=1
 
INDEX ON STR(np,3)+STR(nd,3) TAG T1
DO fltstructure WITH 'kp#0','curraspprn'

SELECT curraspprn
SET ORDER TO 1
SELECT sprkat
COUNT TO maxExcel
DIMENSION dim_Excel(maxExcel,6),kod_excel(maxExcel)
GO TOP 
FOR i=1 TO maxExcel
    kod_Excel(i)=kod
    SKIP
ENDFOR
STORE 0 TO dim_Excel
CREATE CURSOR curExcel  FROM ARRAY arRasp
ALTER TABLE curExcel ADD COLUMN namepodr C(100)
ALTER TABLE curExcel ALTER COLUMN kse N(7,2)
ALTER TABLE curExcel ADD COLUMN numIt N(2)
IF par1=1
   ALTER TABLE curExcel ALTER COLUMN kse_tot N(7,2)
   ALTER TABLE curExcel ALTER COLUMN kse_am N(7,2)
   ALTER TABLE curExcel ALTER COLUMN kse_stac N(7,2)
   ALTER TABLE curExcel ALTER COLUMN kse_sk N(7,2)
   ALTER TABLE curExcel ALTER COLUMN kse_mjr N(7,2)
ENDIF    
SELECT curRaspPrn  
FOR i=1 TO maxExcel
    SUM kse,kse_tot,kse_am,kse_stac,kse_sk,kse_mjr TO dim_Excel(i,1),dim_Excel(i,2),dim_Excel(i,3),dim_Excel(i,4),dim_Excel(i,5),dim_Excel(i,6) FOR kat=kod_Excel(i)
ENDFOR   
GO TOP

SELECT curRaspPrn
SCAN ALL
    IF SEEK(curRaspPrn.kp,'sprpodr',1)
        REPLACE np WITH sprpodr.np,kpp WITH sprpodr.kpp,kpp1 WITH sprpodr.kpp1,kodkpp WITH sprpodr.kodkpp,kodkpp1 WITH sprpodr.kodkpp1  
    ENDIF   
ENDSCAN     

SELECT sprpodr
SCAN ALL
     IF sprpodr.kodKpp#0
        SELECT curRaspPrn
        REPLACE kodkpp WITH sprpodr.kodkpp FOR kp=sprpodr.kodkpp
     ENDIF 
     IF sprpodr.kodKpp1#0
        SELECT curRaspPrn
        REPLACE kodkpp1 WITH sprpodr.kodkpp1,kodkpp WITH sprpodr.kodkpp FOR kp=sprpodr.kodkpp1
        
     ENDIF 
     SELECT sprpodr
ENDSCAN   


SELECT curRaspPrn
GO TOP

kp_cx=kp
kpp_cx=kodkpp
kpp1_cx=kodkpp1

STORE 0 TO ksecx,totcx,amcx,staccx,skcx,mjrcx,kse_kpp,kse_kpp1,am_kpp1,am_kpp,stac_kpp,stac_kpp1,sk_kpp,sk_kpp1,mjr_kpp,mjr_kpp1,tot_kpp,tot_kpp1
STORE 0 TO ksecx_tot,totcx_tot,amcx_tot,staccx_tot,skcx_tot,mjrcx_tot
DO WHILE !EOF()         
   SCATTER TO ax
   SELECT curExcel     
   APPEND BLANK 
   GATHER FROM ax      
   ksecx=ksecx+kse
   totcx=totcx+kse_tot         
   amcx=amcx+kse_am
   staccx=staccx+kse_stac
   skcx=skcx+kse_sk
   mjrcx=mjrcx+kse_mjr         
   ksecx_tot=ksecx_tot+kse
   totcx_tot=totcx_tot+kse_tot         
   amcx_tot=amcx_tot+kse_am
   staccx_tot=staccx_tot+kse_stac
   skcx_tot=skcx_tot+kse_sk
   mjrcx_tot=mjrcx_tot+kse_mjr        
   
   kse_Kpp=IIF(kpp_cx#0,kse_kpp+kse,kse_kpp)  
   tot_kpp=IIF(kpp_cx#0,tot_kpp+kse_tot,tot_kpp)
   am_kpp=IIF(kpp_cx#0,am_kpp+kse_am,am_kpp)
   stac_kpp=IIF(kpp_cx#0,stac_kpp+kse_stac,stac_kpp)
   sk_kpp=IIF(kpp_cx#0,sk_kpp+kse_sk,sk_kpp)
   mjr_kpp=IIF(kpp_cx#0,mjr_kpp+kse_mjr,mjr_kpp)
         
   kse_Kpp1=IIF(kpp1_cx#0,kse_kpp1+kse,kse_kpp1)                
   tot_kpp1=IIF(kpp1_cx#0,tot_kpp1+kse_tot,tot_kpp1)
   am_kpp1=IIF(kpp1_cx#0,am_kpp1+kse_am,am_kpp1)
   stac_kpp1=IIF(kpp1_cx#0,stac_kpp1+kse_stac,stac_kpp1)
   sk_kpp1=IIF(kpp1_cx#0,sk_kpp1+kse_sk,sk_kpp1)
   mjr_kpp1=IIF(kpp1_cx#0,mjr_kpp1+kse_sk,mjr_kpp1)
            
   SELECT curRaspPrn
   SKIP
   IF kp#kp_cx
      SELECT curExcel 
      APPEND BLANK
      REPLACE named WITH '�����',kp WITH kp_cx,kse WITH ksecx,kse_tot WITH totcx,;
              kse_am WITH amcx,kse_stac WITH staccx,kse_sk WITH skcx,kse_mjr WITH mjrcx,numIt WITH 1                          
                            
       IF kpp1_cx#curRaspPrn.kodkpp1.AND.kpp1_cx#0
         SELECT curExcel
         APPEND BLANK      
         REPLACE named WITH '�����',kp WITH kp_cx,kse WITH kse_kpp1,kse_tot WITH tot_kpp1,;
                 kse_am WITH am_kpp1,kse_stac WITH stac_kpp1,kse_sk WITH sk_kpp1,kse_mjr WITH mjr_kpp1,numIt WITH 2        
         kpp1_cx=curRaspPrn.kodkpp1
         STORE 0 TO kse_kpp1,am_kpp1,stac_kpp1,sk_kpp1,tot_kpp1,mjr_kpp1
      ENDIF 
         
     *--------    
     IF kpp_cx#curRaspPrn.kodkpp.AND.kpp_cx#0      
         SELECT curExcel    
         APPEND BLANK      
         REPLACE named WITH '�����',kp WITH kp_cx,kse WITH kse_kpp,kse_tot WITH tot_kpp,;
                 kse_am WITH am_kpp,kse_stac WITH stac_kpp,kse_sk WITH sk_kpp,kse_mjr WITH mjr_kpp,numIt WITH 3                    
         kpp_cx=curRaspPrn.kodkpp
         STORE 0 TO kse_kpp,am_kpp,stac_kpp,sk_kpp,tot_kpp,mjr_kpp
      ENDIF      
      kpp_cx=curRaspPrn.kodkpp
      kpp1_cx=curRaspPrn.kodkpp1                                                                                                                                                                                                                              
      SELECT curraspprn            
      kp_cx=kp
      STORE 0 TO ksecx,totcx,amcx,staccx,skcx,mjrcx
   ENDIF 
ENDDO
SELECT curExcel
APPEND BLANK
REPLACE named WITH '����� �� �����������',kse WITH ksecx_tot,kse_tot WITH totcx_tot,;
        kse_am WITH amcx_tot,kse_stac WITH staccx_tot,kse_sk WITH skcx_tot,kse_mjr WITH mjrcx_tot,numIt WITH 4
                    
APPEND BLANK
REPLACE named WITH '�� ���:'
FOR i=1 TO maxExcel
    IF !EMPTY(dim_Excel(i,1))
       APPEND BLANK
       REPLACE named WITH IIF(SEEK(kod_Excel(i),'sprkat',1),sprkat.namefull,''),kse WITH dim_Excel(i,1),kse_tot WITH dim_Excel(i,2),;
       kse_am WITH dim_Excel(i,3),kse_stac WITH dim_Excel(i,4),kse_sk WITH dim_Excel(i,5),kse_mjr WITH dim_Excel(i,6)
    ENDIF       
ENDFOR
SELECT curExcel      
GO TOP
IF par1=1   && �����������
   DO procForPrintAndPreview WITH 'reprasp','������� ����������'  
ELSE    && �����
   DO procForPrintAndPreview WITH 'repraspshort','������� ����������'  
ENDIF    

SELECT rasp
************************************************************************************************************************
PROCEDURE shtatToExcel
PARAMETERS parForm
DO startPrnToExcel WITH 'fSupl'      
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167) 
DIMENSION dim_rows(1)  && ������ ��� �����, ��� ������ ����� �����������
STORE 0 TO dim_rows
SELECT curExcel
DO storezeropercent
max_rows=0 
IF parForm=1 
   WITH excelBook.Sheets(1) 
        .PageSetup.Orientation = 1      
        .Columns(1).ColumnWidth=3    
        .Columns(2).ColumnWidth=28
        .Columns(3).ColumnWidth=6
        .Columns(4).ColumnWidth=6
        .Columns(5).ColumnWidth=6
        .Columns(6).ColumnWidth=6
        .Columns(7).ColumnWidth=6
        .Columns(8).ColumnWidth=6
        .Columns(9).ColumnWidth=12 
        rowcx=3           
        IF  MEMLINES(datShtat.nhead)>0 
        FOR i=1 TO MEMLINES(datShtat.nhead)
            rowcx=rowcx+1
            .Range(.Cells(rowcx,5),.Cells(rowcx,9)).Select  
            WITH objExcel.Selection
                 .MergeCells=.T.
                 .HorizontalAlignment=xlLeft
                 .VerticalAlignment=1
                 .WrapText=.T.
                 .Value=MLINE(datShtat.nhead,i)
                 .Font.Name='Times New Roman'   
                 .Font.Size=11
            ENDWITH          
        ENDFOR 
        ENDIF 
        rowcx=rowcx+1     
        .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='�������  ����������'
             .Font.Name='Times New Roman'   
             .Font.Size=12
        ENDWITH   
        rowcx=rowcx+2
               
        .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='�� '+STR(YEAR(dateTar),4)+' ���'            
             .Font.Name='Times New Roman'   
             .Font.Size=12
        ENDWITH  
         rowcx=rowcx+2     
        .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value=office_say
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH   
        rowcx=rowcx+2     
             
        .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value=adres_say
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH   
        rowcx=rowcx+2     
                                      
                 
        .Range(.Cells(rowcx,1),.Cells(rowcx+7,1)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='� �/�'
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH          
         
        .Range(.Cells(rowcx,2),.Cells(rowcx+7,2)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='������������ ����������� ������������� � ���������� (��������� ����������)'
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH           
                 
                     
        .Range(.Cells(rowcx,3),.Cells(rowcx+7,3)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='����� ���������� (��������� ����������) �����'
             .Font.Name='Times New Roman'   
             .Font.Size=10
         ENDWITH       
           
         .Range(.Cells(rowcx,4),.Cells(rowcx+2,8)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='� ��� ����� �� �������� �������� ����������� ������ (�����)'                    
              .Font.Name='Times New Roman'   
              .Font.Size=11
         ENDWITH                                                 
           
         .Range(.Cells(rowcx,9),.Cells(rowcx+7,9)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='����������'   
              .Font.Name='Times New Roman'   
              .Font.Size=11                 
         ENDWITH                                        
           
         .Range(.Cells(rowcx+3,4),.Cells(rowcx+7,4)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='����� ��� ���� ����� ��������'                    
              .Font.Name='Times New Roman'   
              .Font.Size=8
         ENDWITH                        
                      
         .Range(.Cells(rowcx+3,5),.Cells(rowcx+7,5)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='������������ �������'              
              .Font.Name='Times New Roman'   
              .Font.Size=8
         ENDWITH                 
                  
         .Range(.Cells(rowcx+3,6),.Cells(rowcx+7,6)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='������������ �������'
              .Font.Name='Times New Roman'   
              .Font.Size=8
         ENDWITH       
                               
         .Range(.Cells(rowcx+3,7),.Cells(rowcx+7,7)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='������ ����������� ������'
              .Font.Name='Times New Roman'   
              .Font.Size=8
         ENDWITH   
                                  
         .Range(.Cells(rowcx+3,8),.Cells(rowcx+7,8)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlCenter
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='������. ��������� ����������� �����'
              .Font.Name='Times New Roman'   
              .Font.Size=8
         ENDWITH                                         
              
         rowcx=rowcx+8
         .cells(rowcx,1).Value='1'
         .cells(rowcx,2).Value='2'
         .cells(rowcx,3).Value='3'
         .cells(rowcx,4).Value='4'
         .cells(rowcx,5).Value='5'
         .cells(rowcx,6).Value='6'
         .cells(rowcx,7).Value='7'
         .cells(rowcx,8).Value='8'
         .cells(rowcx,9).Value='9'
         .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select
         objExcel.Selection.HorizontalAlignment=xlCenter
         numberRow=rowcx+1  
         rowtop=numberRow         
         SELECT curExcel          
         GO TOP
         kpold=0        
         begRow=0
         endRow=0
        * endSumTot='=����(' 
         endsumTotC=''
         endSumTotD=''
         endSumTotE=''
         endSumTotF=''
         endSumTotG=''
         endSumTotH=''
         SCAN ALL
              IF nd=1
                 IF !EMPTY(primHeadr)
                    .Range(.Cells(numberRow,1),.Cells(numberRow,9)).Select
                    objExcel.Selection.MergeCells=.T.
                    objExcel.Selection.HorizontalAlignment=xlCenter
                    objExcel.Selection.VerticalAlignment=1
                    objExcel.Selection.WrapText=.T.
                    objExcel.Selection.Value=curExcel.primheadr   
                    max_rows=max_rows+1
                    DIMENSION dim_rows(max_rows)          
                    dim_rows(max_rows)=numberRow
                    numberRow=numberRow+1
                 ENDIF  
                 .Range(.Cells(numberRow,1),.Cells(numberRow,9)).Select
                 objExcel.Selection.MergeCells=.T.
                 objExcel.Selection.HorizontalAlignment=xlCenter
                 objExcel.Selection.VerticalAlignment=1
                 objExcel.Selection.WrapText=.T.
                 objExcel.Selection.Value=curExcel.namepodr                          
                 numberRow=numberRow+1                                
                 begRow=numberRow
              ENDIF               
              .Cells(numberRow,1).Value=IIF(curExcel.nd#0,curExcel.nd,'')                                      
              .Cells(numberRow,2).Value=ALLTRIM(curExcel.named)     
              .Cells(numberRow,3).Value=curExcel.kse
              .Cells(numberRow,3).NumberFormat='0.00'                                      
              .Cells(numberRow,4).Value=IIF(curExcel.kse_tot#0,curExcel.kse_tot,'')                                     
              .Cells(numberRow,4).NumberFormat='0.00'   
              .Cells(numberRow,5).Value=IIF(curExcel.kse_am#0,curExcel.kse_am,'')                                       
              .Cells(numberRow,5).NumberFormat='0.00'   
              .Cells(numberRow,6).Value=IIF(curExcel.kse_stac#0,curExcel.kse_stac,'')                                       
              .Cells(numberRow,6).NumberFormat='0.00'                                       
              .Cells(numberRow,7).Value=IIF(curExcel.kse_sk#0,curExcel.kse_sk,'')  
              .Cells(numberRow,7).NumberFormat='0.00'                                       
              .Cells(numberRow,8).Value=IIF(curExcel.kse_mjr#0,curExcel.kse_mjr,'') 
              .Cells(numberRow,8).NumberFormat='0.00'   
              endRow=numberRow-1
              DO CASE
                 CASE numIt=1              
                      endsum='=����('+'C'+LTRIM(STR(begRow))+':'+'C'+LTRIM(STR(endRow))+')'
                      .Cells(numberRow,3).formulaLocal=endsum  
                      endSumTotC=endSumTotC+'C'+LTRIM(STR(numberRow))+';' 
                      
                      endsum='=����('+'D'+LTRIM(STR(begRow))+':'+'D'+LTRIM(STR(endRow))+')'
                      .Cells(numberRow,4).formulaLocal=endsum                 
                      endSumTotD=endSumTotD+'D'+LTRIM(STR(numberRow))+';' 
                      
                      endsum='=����('+'E'+LTRIM(STR(begRow))+':'+'E'+LTRIM(STR(endRow))+')'
                      .Cells(numberRow,5).formulaLocal=endsum
                      endSumTotE=endSumTotE+'E'+LTRIM(STR(numberRow))+';'                  
                       
                      endsum='=����('+'F'+LTRIM(STR(begRow))+':'+'F'+LTRIM(STR(endRow))+')'
                      .Cells(numberRow,6).formulaLocal=endsum                 
                      endSumTotF=endSumTotF+'F'+LTRIM(STR(numberRow))+';'                  
                      
                      endsum='=����('+'G'+LTRIM(STR(begRow))+':'+'G'+LTRIM(STR(endRow))+')'
                      .Cells(numberRow,7).formulaLocal=endsum 
                      endSumTotG=endSumTotG+'G'+LTRIM(STR(numberRow))+';'                  
                      
                      endsum='=����('+'H'+LTRIM(STR(begRow))+':'+'H'+LTRIM(STR(endRow))+')'
                      .Cells(numberRow,8).formulaLocal=endsum 
                      endSumTotH=endSumTotH+'H'+LTRIM(STR(numberRow))+';'                                        
                * CASE numIt=4    
                *      endSumTotC='=����('+endSumTotC+')'
                *      .Cells(numberRow,3).formulaLocal=endSumTotC 
                                       
                *      endSumTotD='=����('+endSumTotD+')'
                *      .Cells(numberRow,4).formulaLocal=endSumTotD   
                      
                 *     endSumTotE='=����('+endSumTotE+')'
                 *     .Cells(numberRow,5).formulaLocal=endSumTotE                  
                      
                  *    endSumTotF='=����('+endSumTotF+')'
                  *    .Cells(numberRow,6).formulaLocal=endSumTotF   
                      
                   *   endSumTotG='=����('+endSumTotG+')'
                   *   .Cells(numberRow,7).formulaLocal=endSumTotG                  
                      
                    *  endSumTotH='=����('+endSumTotH+')'
                    *  .Cells(numberRow,8).formulaLocal=endSumTotH   
                                     
              ENDCASE                                   
              numberRow=numberRow+1
              DO fillpercent WITH 'fSupl'
         ENDSCAN                                 
         .Range(.Cells(rowcx-7,1),.Cells(numberRow-1,9)).Select
         objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
         objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
         objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
         objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
         objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
         objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
         objExcel.Selection.VerticalAlignment=1
           
         .Range(.Cells(rowcx,1),.Cells(numberRow-1,9)).Select
         objExcel.Selection.Font.Name='Times New Roman' 
         objExcel.Selection.Font.Size=9      
         objExcel.Selection.WrapText=.T.           
         IF max_rows#0
            FOR i=1 TO max_rows
                .Range(.Cells(dim_rows(i),1),.Cells(dim_rows(i),9)).Select
                 objExcel.Selection.Borders(xlEdgeBottom).lineStyle=xlLineStyleNone
            ENDFOR
         ENDIF
         .Cells(1,1).Select                       
   ENDWITH    
ELSE
   WITH excelBook.Sheets(1)
        .PageSetup.Orientation = 1
        .Columns(1).ColumnWidth=3
        .Columns(2).ColumnWidth=50
        .Columns(3).ColumnWidth=6
        .Columns(4).ColumnWidth=24  
        rowcx=3 
        IF  MEMLINES(datShtat.nhead)>0
            FOR i=1 TO MEMLINES(datShtat.nhead)
                rowcx=rowcx+1
                .Range(.Cells(rowcx,3),.Cells(rowcx,4)).Select  
                WITH objExcel.Selection
                     .MergeCells=.T.
                     .HorizontalAlignment=xlLeft
                     .VerticalAlignment=1
                     .WrapText=.T.
                     .Value=MLINE(datshtat.nhead,i)
                     .Font.Name='Times New Roman'   
                     .Font.Size=11
                 ENDWITH          
            ENDFOR   
            rowcx=rowcx+1     
        ENDIF    
        .Range(.Cells(rowcx,1),.Cells(rowcx,4)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='�������  ����������'
             .Font.Name='Times New Roman'   
             .Font.Size=12
        ENDWITH   
        rowcx=rowcx+2
                    
        .Range(.Cells(rowcx,1),.Cells(rowcx,4)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='�� '+STR(YEAR(dateTar),4)+' ���'                   
             .Font.Name='Times New Roman'   
             .Font.Size=12
        ENDWITH   
        rowcx=rowcx+2     
        
        .Range(.Cells(rowcx,1),.Cells(rowcx,4)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value=office_say
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH   
        rowcx=rowcx+2     
             
        .Range(.Cells(rowcx,1),.Cells(rowcx,4)).Select         
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value=adres_say
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH   
        rowcx=rowcx+2                   
             
        .Range(.Cells(rowcx,1),.Cells(rowcx+7,1)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='� �/�'
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH          
           
        .Range(.Cells(rowcx,2),.Cells(rowcx+7,2)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='������������ ����������� ������������� � ���������� (��������� ����������)'
             .Font.Name='Times New Roman'   
             .Font.Size=11
        ENDWITH           
                    
                     
        .Range(.Cells(rowcx,3),.Cells(rowcx+7,3)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='����� ���������� (��������� ����������) �����'
             .Font.Name='Times New Roman'   
             .Font.Size=10
        ENDWITH                  
            
        .Range(.Cells(rowcx,4),.Cells(rowcx+7,4)).Select
        With objExcel.Selection
             .MergeCells=.T.
             .HorizontalAlignment=xlCenter
             .VerticalAlignment=1
             .WrapText=.T.
             .Value='����������'   
             .Font.Name='Times New Roman'   
             .Font.Size=11                 
        ENDWITH                                                                   
              
        rowcx=rowcx+8
        .cells(rowcx,1).Value='1'
        .cells(rowcx,2).Value='2'
        .cells(rowcx,3).Value='3'
        .cells(rowcx,4).Value='4'            
        .Range(.Cells(rowcx,1),.Cells(rowcx,4)).Select
        objExcel.Selection.HorizontalAlignment=xlCenter
        numberRow=rowcx+1  
        rowtop=numberRow         
        SELECT curExcel
        GO TOP
        kpold=0  
        begRow=0
        endRow=0    
        endSumTot='=����('  
        SCAN ALL
             IF nd=1
                IF !EMPTY(primHeadr)
                   .Range(.Cells(numberRow,1),.Cells(numberRow,4)).Select
                   objExcel.Selection.MergeCells=.T.
                   objExcel.Selection.HorizontalAlignment=xlCenter
                   objExcel.Selection.VerticalAlignment=1
                   objExcel.Selection.WrapText=.T.
                   objExcel.Selection.Value=curExcel.primheadr   
                   max_rows=max_rows+1
                   DIMENSION dim_rows(max_rows)          
                   dim_rows(max_rows)=numberRow
                   numberRow=numberRow+1
                ENDIF  
                .Range(.Cells(numberRow,1),.Cells(numberRow,4)).Select
                objExcel.Selection.MergeCells=.T.
                objExcel.Selection.HorizontalAlignment=xlCenter
                objExcel.Selection.VerticalAlignment=1
                objExcel.Selection.WrapText=.T.
                objExcel.Selection.Value=curExcel.namepodr                   
                numberRow=numberRow+1
                begRow=numberRow
             ENDIF 
             .Cells(numberRow,1).Value=IIF(curExcel.nd#0,curExcel.nd,'')                                      
             .Cells(numberRow,2).Value=ALLTRIM(curExcel.named)     
             .Cells(numberRow,3).Value=curExcel.kse
             .Cells(numberRow,3).NumberFormat='0.00'
             endRow=numberRow-1
             DO CASE
                CASE numIt=1
                     endsum='=����('+'C'+LTRIM(STR(begRow))+':'+'C'+LTRIM(STR(endRow))+')'
                     .Cells(numberRow,3).formulaLocal=endsum   
                     endSumTot=endSumTot+'C'+LTRIM(STR(numberRow))+';'                          
                *CASE numIt=4     
                *     endSumTot=endSumTot+')' 
                *     .Cells(numberRow,3).formulaLocal=endSumTot   
             ENDCASE    
             numberRow=numberRow+1
             one_pers=one_pers+1
             pers_ch=one_pers/max_rec*100
             fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
             fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch              
        ENDSCAN                                 
        .Range(.Cells(rowcx-7,1),.Cells(numberRow-1,4)).Select
        objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
        objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
        objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
        objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
        objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
        objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
        objExcel.Selection.VerticalAlignment=1
           
        .Range(.Cells(rowcx,1),.Cells(numberRow-1,4)).Select
        objExcel.Selection.Font.Name='Times New Roman' 
        objExcel.Selection.Font.Size=9      
        objExcel.Selection.WrapText=.T.  
        IF max_rows#0
           FOR i=1 TO max_rows
               .Range(.Cells(dim_rows(i),1),.Cells(dim_rows(i),4)).Select
                objExcel.Selection.Borders(xlEdgeBottom).lineStyle=xlLineStyleNone
           ENDFOR
        ENDIF
        .Cells(1,1).Select             
   ENDWITH      
ENDIF  
#UNDEFINE xlCenter     
#UNDEFINE xlLeft   
#UNDEFINE xlRight
#UNDEFINE xlThin                 
#UNDEFINE xlMedium             
#UNDEFINE xlDiagonalDown          
#UNDEFINE xlDiagonalUp                
#UNDEFINE xlEdgeLeft            
#UNDEFINE xlEdgeTop             
#UNDEFINE xlEdgeBottom          
#UNDEFINE xlEdgeRight           
#UNDEFINE xlInsideVertical      
#UNDEFINE xlInsideHorizontal 
#UNDEFINE xlLineStyleNone -4142  
DO endPrnToExcel WITH 'fSupl' 
objExcel.Visible=.T.

*********************************************************************************
PROCEDURE topspistar
PARAMETERS dim_ch
DO topToolPreview WITH .T.
STORE 0 TO dim_kat,dim_dopl
**********************************************************************************
*          �� ���������� ������� ���� (�������)
**********************************************************************************
PROCEDURE attYear
IF USED('peoplevred')
   SELECT peoplevred
   USE
ENDIF
IF USED('curprn')
   SELECT curprn
   USE
ENDIF
DIMENSION sum_podr(6),sum_tot(6),dim_vred(12),dimpodr_vred(12),ksekat_podr(6),dim_sum(12)
STORE 0 TO sum_tot,dim_kat,sum_podr,numrecrep,ksekat_podr,numpage,dim_vred,dimpodr_vred,dim_sum
logKom=formbase.log_kom
ksekat_podr=0
SELECT rasp
rrec=RECNO()
SELECT datjob
oldrec=RECNO()

*SELECT * FROM datjob WHERE sumvr#0 INTO CURSOR curprn READWRITE
SELECT * FROM datjob INTO CURSOR curprn READWRITE
ALTER TABLE curprn ADD COLUMN nIt N(1)
ALTER TABLE curprn ADD COLUMN hourTot N(7,2)
ALTER TABLE curprn ADD COLUMN npp N(2)
ALTER TABLE curprn ALTER COLUMN kse N(7,2)
ALTER TABLE curprn ALTER COLUMN sumVrTot N(8,2)


*----------------------------------- �������������� ���������� ��������---------------------------------------------
IF formbase.avt_vac
   SELECT datJob
   SET FILTER TO 
   ordOld=SYS(21)
   SET ORDER TO 2
   SELECT rasp
   GO TOP
   DO WHILE !EOF()
      IF rasp.kse#0.AND.rasp.patt#0    
         SELECT datJob        
         SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
         kse_cx=rasp.kse
         DO  WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()   
             IF date_in>dateTar        
             ELSE 
                kse_cx=kse_cx-datjob.kse
             ENDIF 
             SKIP    
         ENDDO       
         IF kse_cx>0
            IF !formbase.vacst               
               SELECT curPrn
               APPEND BLANK
               REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH kse_cx,vac WITH .T.,;
                       np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,patt WITH rasp.patt,sumvr WITH varBaseSt/100*patt,kfVr WITH rasp.kfVr,pkfVr WITH rasp.pkfVr
               tar_ok=0
               tar_ok=varBaseSt*curPrn.namekf*IIF(pkf#0,pkf,1)
               sum_tot=0
               FOR hi=1 TO 12
                   rep_cx='vr'+LTRIM(STR(hi))
                   repTime='sprtime.t'+LTRIM(STR(hi))
                   REPLACE &rep_cx WITH sumvr*IIF(SEEK(vTime,'sprtime',1),&repTime,0)
                   sum_tot=sum_tot+&rep_cx
               ENDFOR
               REPLACE sumvrtot WITH sum_tot 
            ELSE 
               DO CASE
                  CASE kse_cx<=1
                       kvovac=1
                  CASE MOD(kse_cx,1)=0     
                       kvovac=INT(kse_cx)
                  CASE MOD(kse_cx,1)>0     
                       kvovac=INT(kse_cx)+1    
               ENDCASE               
               kvokse=kse_cx
               ksevac=0
               FOR i=1 TO kvovac              
                   ksevac=IIF(kvokse<=1,kvokse,1)
                   SELECT curPrn
                   APPEND BLANK
                   REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH ksevac,vac WITH .T.,;
                       np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,patt WITH rasp.patt,sumvr WITH varBaseSt/100*patt,kfVr WITH rasp.kfVr,pkfVr WITH rasp.pkfVr
                   tar_ok=0
                   tar_ok=varBaseSt*curPrn.namekf*IIF(pkf#0,pkf,1)
                   sum_tot=0
                   FOR hi=1 TO 12
                       rep_cx='vr'+LTRIM(STR(hi))
                       repTime='sprtime.t'+LTRIM(STR(hi))
                       REPLACE &rep_cx WITH sumvr*IIF(SEEK(vTime,'sprtime',1),&repTime,0)
                       sum_tot=sum_tot+&rep_cx                     
                   ENDFOR
                   REPLACE sumvrtot WITH sum_tot                        
                   kvokse=kvokse-1
               ENDFOR
            ENDIF    
         ENDIF       
      ENDIF
      SELECT rasp
      SKIP   
   ENDDO
   SELECT datJob
   SET ORDER TO &ordOld 
ENDIF  
SELECT curprn

=AFIELDS(arCurPrn,'curprn')
CREATE CURSOR peoplevred FROM ARRAY arCurPrn
SELECT curprn

DO fltstructure WITH 'sumvrtot#0','curprn'
DELETE FOR vac.AND.kse#1
IF onlyVac
   SELECT curPrn
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curPrn
   DELETE FOR vac
ENDIF

REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE kfvr WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kfvr,0) ALL
REPLACE hourTot WITH IIF(SEEK(vTime,'sprtime',1),sprtime.t1+sprtime.t2+sprtime.t3+sprtime.t4+sprtime.t5+sprtime.t6+sprtime.t7+sprtime.t8+sprtime.t9+sprtime.t10+sprtime.t11+sprtime.t12,0) ALL 


SCAN ALL
     SCATTER TO dima
     SELECT peoplevred
     APPEND BLANK
     GATHER FROM dima 
     FOR i=1 TO 12
         vrcx='vr'+LTRIM(STR(i))
         dim_sum(i)=dim_sum(i)+&vrcx
     ENDFOR 
     SELECT curprn   
ENDSCAN
SELECT peoplevred
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
GO TOP 
kpOld=kp
nppNew=1
DO WHILE !EOF()
   REPLACE npp WITH nppNew   
   SKIP 
   nppNew=nppNew+1
   IF kp#kpOld
      nppNew=1
      kpOld=kp
   ENDIF
ENDDO
kpold=kp
npold=np
STORE 0 TO ksePodr,ksetot,sumVrTot_podr,sumvrtot_tot
SUM kse,sumvrtot TO ksetot,sumvrtot_tot FOR nIt=0
APPEND BLANK
REPLACE np WITH 999,FIO WITH '�����',kse WITH ksetot,sumVrtot WITH sumvrtot_tot,nIt WITH 9
SELECT sprpodr
GO TOP
DO WHILE !EOF()
   SELECT peoplevred
   SUM kse,sumvrtot TO ksePodr,sumVrTot_podr FOR kp=sprpodr.kod.AND.nIt=0   
   IF sumVrTot_podr#0
     APPEND BLANK
     REPLACE np WITH sprpodr.np,kp WITH sprpodr.kod,nd WITH 99,fio WITH '�����',kse WITH ksePodr,sumVrtot WITH sumvrtot_podr,nIt WITH 1  
   ENDIF
   SELECT sprpodr
   SKIP   
ENDDO 
SELECT peoplevred
FOR i=1 TO 12
    APPEND BLANK
    REPLACE np WITH 999,nd WITH i,sumvrtot WITH dim_sum(i),fio WITH dim_month(i)
ENDFOR
GO TOP
DO procForPrintAndPreview WITH 'repvred1','�������� �� ���������'
SELECT people
**********************************************************************************
*          �� ���������� ������� ���� (�������)
**********************************************************************************
PROCEDURE attYearToExcel
DO startPrnToExcel WITH 'fSupl'  
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=25
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
     .Columns(8).ColumnWidth=8
     .Columns(9).ColumnWidth=8
         
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1    
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������ �������� �� ������� ������� ���������� �� ����������� ���������� ������� ����'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH    
     
     rowcx=rowcx+1  
     .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�� '+STR(YEAR(dateTar),4)+' ���'     
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
      rowcx=rowcx+1                       
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='� �/�'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���� �����, ������������ ������������� ������� ��� �������� ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH      
    
        
      .Range(.Cells(rowcx,4),.Cells(rowcx,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='����� ������� �����'                   
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,5),.Cells(rowcx,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='% �������� ������ 1-�� �.'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,6),.Cells(rowcx,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='����� ������'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,7),.Cells(rowcx,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='� ���'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
     .Range(.Cells(rowcx,8),.Cells(rowcx,8)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='����� �� ���'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     
    .Range(.Cells(rowcx,9),.Cells(rowcx,9)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='����� �� ���'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                                  
              
      rowcx=rowcx+1
      .cells(rowcx,1).Value='1'
      .cells(rowcx,2).Value='2'
      .cells(rowcx,3).Value='3'
      .cells(rowcx,4).Value='4'
      .cells(rowcx,5).Value='5'
      .cells(rowcx,6).Value='6'
      .cells(rowcx,7).Value='7'
      .cells(rowcx,8).Value='8'
      .cells(rowcx,9).Value='9'
  
      .Range(.Cells(rowcx,1),.Cells(rowcx,9)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      rowtop=numberRow         
      SELECT peoplevred
      DO storezeropercent
      GO TOP
      kpold=0
      SCAN ALL
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,9)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(peoplevred.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
           .Cells(numberRow,1).Value=IIF(npp#0,npp,'')                                    
           .Cells(numberRow,2).Value=peoplevred.fio                                       
           .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')                                                                        
           .Cells(numberRow,4).Value=IIF(kfvr#0,LTRIM(STR(kfvr,3,1)),'')                                      
           .Cells(numberRow,5).Value=IIF(pkfvr#0,LTRIM(STR(pkfvr,5,2)),'')
           .Cells(numberRow,6).Value=IIF(kse#0,kse,'')
           .Cells(numberRow,6).NumberFormat='0.00'           
           .Cells(numberRow,7).Value=IIF(sumvr#0,sumvr,'')
           .Cells(numberRow,7).NumberFormat='0.00'                                                 
           .Cells(numberRow,8).Value=IIF(hourtot#0,hourtot,'')               
           .Cells(numberRow,8).NumberFormat='0.00'             
           .Cells(numberRow,9).Value=IIF(sumvrtot#0,sumvrtot,'')  
           .Cells(numberRow,9).NumberFormat='0.00'  
           numberRow=numberRow+1
           DO fillpercent WITH 'fSupl'
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,9)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
      
      IF logKom
         numberRow=numberRow+1
         .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlLeft
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='��������'                   
              .Font.Name='Times New Roman'   
              .Font.Size=9
         ENDWITH 
         numberRow=numberRow+1
         FOR i=1 TO 12
             .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,1)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH 
              
             .Range(.Cells(numberRow,5),.Cells(numberRow,7)).Select
             With objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment=xlLeft
                  .VerticalAlignment=1
                  .WrapText=.T.
                  .Value=dim_boss(i,2)                   
                  .Font.Name='Times New Roman'   
                  .Font.Size=9
              ENDWITH                    
             numberRow=numberRow+1
         ENDFOR
      ENDIF
          
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,9)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'            
      
objExcel.Visible=.T.
**********************************************************************************
*                 �� ���� ������� (� ������ ������������ c����)
**********************************************************************************
PROCEDURE yearStaj
SELECT datJob
SET FILTER TO 
SELECT rasp
SET FILTER TO 
IF USED('curPrn')
   SELECT curPrn
   USE 
ENDIF
IF USED('curTarJob')
   SELECT curTarJob
   USE
ENDIF    
SELECT * FROM sprpodr INTO CURSOR curPrn READWRITE

ALTER TABLE curPrn ADD COLUMN m1 N(10,2)
ALTER TABLE curPrn ADD COLUMN m2 N(10,2)
ALTER TABLE curPrn ADD COLUMN m3 N(10,2)
ALTER TABLE curPrn ADD COLUMN m4 N(10,2)
ALTER TABLE curPrn ADD COLUMN m5 N(10,2)
ALTER TABLE curPrn ADD COLUMN m6 N(10,2)
ALTER TABLE curPrn ADD COLUMN m7 N(10,2)
ALTER TABLE curPrn ADD COLUMN m8 N(10,2)
ALTER TABLE curPrn ADD COLUMN m9 N(10,2)
ALTER TABLE curPrn ADD COLUMN m10 N(10,2)
ALTER TABLE curPrn ADD COLUMN m11 N(10,2)
ALTER TABLE curPrn ADD COLUMN m12 N(10,2)
ALTER TABLE curPrn ADD COLUMN mtot N(12,2)
INDEX ON np TAG T1
INDEX ON kod TAG T2
SET ORDER TO 2
              
SELECT * FROM datJob INTO CURSOR curTarJob READWRITE
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DELETE FOR date_in>varDtar && ������� ��������� ����� ���� �-���
ALTER TABLE curTarJob ADD COLUMN npp N(3)
ALTER TABLE curTarJob ADD COLUMN nit N(1)
ALTER TABLE curTarJob ADD COLUMN nkat C(200)
SELECT curTarJob
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.np,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL
DELETE FOR !SEEK(kp,'sprpodr',1)
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SET ORDER TO 1
*----------------------------------- �������������� ���������� �������� ---------------------------------------------
IF formbase.avt_vac   
   SELECT datjob
   ordOld=SYS(21)
   SET FILTER TO 
   SELECT rasp
   GO TOP
   DO WHILE !EOF()
      IF rasp.kse#0
         SELECT datjob
         SET ORDER TO 2
         SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
         kse_cx=rasp.kse
         DO WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()
            IF date_in>varDtar
            ELSE 
               kse_cx=kse_cx-datjob.kse
            ENDIF  
            SKIP
         ENDDO
         IF kse_cx>0
            IF !formbase.vacst
               SELECT curTarJob
               APPEND BLANK
               REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,kse WITH kse_cx,vac WITH .T.,stpr WITH dimConstVac(2,2),mstsum WITH varBaseSt/100*stpr*kse                                                  
             
            ENDIF    
         ENDIF       
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
   SELECT datJob
   SET ORDER TO &ordOld   
ENDIF   
DO fltstructure WITH 'mStsum#0'
IF onlyVac
   SELECT curTarJob
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curTarJob
   DELETE FOR vac
ENDIF 
SELECT curTarJob
SCAN ALL
     DO CASE
        CASE EMPTY(per_date)
             SELECT curPrn
             SEEK curTarJob.kp
             FOR i=1 TO 12
                 IF i>=MONTH(dateTar)                  
                    repm='m'+LTRIM(STR(i))
                    REPLACE &repm WITH &repm+curTarJob.mstsum,mtot WITH mtot+curTarJob.mstsum
                 ENDIF   
             ENDFOR 
        CASE !EMPTY(per_date)
             SELECT curPrn
             SEEK curTarJob.kp
             FOR i=1 TO 12
                 IF i>=MONTH(dateTar)                  
                    repm='m'+LTRIM(STR(i))
                    repSumMonth=0
                    DO CASE 
                       CASE i<MONTH(curTarJob.per_date)
                            repSumMonth=curTarJob.mstsum
                       CASE i=MONTH(curTarJob.per_date)                
                            nmonth=MONTH(curTarJob.per_date)
                            year_ch=YEAR(curTarJob.per_date)
                            STORE 0 TO kvo_old,kvo_new
                            kvoday=IIF(nmonth=2,IIF(MOD(year_ch,4)=0,29,28),IIF(INLIST(nmonth,4,6,9,11),30,31))
                            oneday=DOW(CTOD('01.'+STR(nmonth,2)+'.'+STR(year_ch,4)))
                            oneday=IIF(oneday=1,7,oneday-1)
                            FOR cx=1 TO kvoday
                                oneday=DOW(CTOD(STR(cx,2)+'.'+STR(nmonth,2)+'.'+STR(year_ch,4)))
                                oneday=IIF(oneday=1,7,oneday-1)
                                IF cx<DAY(curTarJob.per_date)
                                   IF oneday<6.AND.!SEEK(cx+nmonth/100,'fete',1)
                                       kvo_old=kvo_old+1
                                   ENDIF
                                ELSE
                                   IF oneday<6.AND.!SEEK(cx+nmonth/100,'fete',1)
                                      kvo_new=kvo_new+1
                                   ENDIF
                                ENDIF
                            ENDFOR
                            sumOld=(varBaseSt/100*curTarJob.stpr*curTarJob.kse)/(kvo_old+kvo_new)*kvo_old
                            sumNew=(varBaseSt/100*curTarJob.st_per*curTarJob.kse)/(kvo_old+kvo_new)*kvo_new
                            repsumMonth=sumOld+sumNew                                                                                               
                       CASE i>MONTH(curTarJob.per_date) 
                            repSumMonth=varBaseSt/100*curTarJob.st_per*curTarJob.kse
                    ENDCASE                     
                    REPLACE &repm WITH &repm+repsummonth,mtot WITH mtot+repsummonth
                 ENDIF   
             ENDFOR               
     ENDCASE
     SELECT curTarJob
ENDSCAN
SELECT curPrn
DELETE FOR mtot=0
SET ORDER TO 1
GO TOP
DO procForPrintAndPreview WITH 'repYearStaj','������� ��������� �� �����'

**********************************************************************************
*                ����� ���� �� �������, ���� ����� ��������, ������ � ������
**********************************************************************************
PROCEDURE totalsvod
IF USED('curprn')
   SELECT curprn
   USE 
ENDIF 
IF USED('curTarJob')
   SELECT curTarJob
   USE 
ENDIF 
IF USED('curTarRasp')
   SELECT curTarRasp
   USE 
ENDIF 
IF USED('curNad')
   SELECT curNad
   USE
ENDIF
ON ERROR DO erSup
RESTORE FROM kfotp ADDITIVE
ON ERROR
lognzptot=.F.
SELECT datjob
FOR fi=1 TO FCOUNT()
    cfi=FIELD(fi)
    IF LOWER(cfi)='nzptot'
       lognzptot=.T.     
       EXIT 
    ENDIF 
ENDFOR


SELECT * FROM rasp INTO CURSOR curTarRasp
SELECT * FROM tarfond WHERE tarfond.vac.AND.!EMPTY(persved) INTO CURSOR curPrnTarFond READWRITE 
SELECT curPrnTarFond
INDEX ON num TAG T1
GO TOP
num_cx=0
DO WHILE !EOF()
   num_cx=num_cx+1
   REPLACE num WITH num_cx
   SKIP   
ENDDO
SELECT * FROM nadBase INTO CURSOR curnad WHERE logsv
SELECT datJob
SET FILTER TO 
SELECT rasp
SET FILTER TO 

SELECT * FROM datJob INTO CURSOR curTarJob READWRITE
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DELETE FOR date_in>varDtar && ������� ��������� ����� ���� �-���
SELECT curTarJob
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.np,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL
DELETE FOR !SEEK(kp,'sprpodr',1)
INDEX ON STR(kp,3)+STR(kd,3) TAG T1
SET ORDER TO 1

*----------------------------------- �������������� ���������� �������� ---------------------------------------------
IF formbase.avt_vac   
   SELECT datjob
   ordOld=SYS(21)
   SET FILTER TO 
   SELECT rasp
   GO TOP
   DO WHILE !EOF()      
      IF rasp.kse#0
         SELECT datjob
         SET ORDER TO 2
         SEEK STR(rasp.kp,3)+STR(rasp.kd,3)      
         kse_cx=rasp.kse
         DO WHILE rasp.kp=datjob.kp.AND.rasp.kd=datjob.kd.AND.!EOF()
            IF date_in>varDtar
            ELSE 
               kse_cx=kse_cx-datjob.kse
            ENDIF  
            SKIP
         ENDDO         
         IF kse_cx>0    
            IF !formbase.vacst                
               SELECT curTarJob
               APPEND BLANK
               REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,pkf WITH rasp.pkf,kse WITH kse_cx,vac WITH .T.
               IF lognzptot
                  REPLACE nzptot WITH rasp.nzptot
               ELSE                  
                  IF rasp.ksezotp#0
                     REPLACE zp1 WITH rasp.z1,zp2 WITH rasp.z2,zp3 WITH rasp.z3,zp4 WITH rasp.z4,zp5 WITH rasp.z5,zp6 WITH rasp.z6,zp7 WITH rasp.z7,zp8 WITH rasp.z8,zp9 WITH rasp.z9,zp10 WITH rasp.z10,zp11 WITH rasp.z11,zp12 WITH rasp.z12
                  ENDIF
               ENDIF      
               REPLACE zpk1 WITH rasp.zk1,zpk2 WITH rasp.zk2,zpk3 WITH rasp.zk3,zpk4 WITH rasp.zk4,zpk5 WITH rasp.zk5,zpk6 WITH rasp.zk6,zpk7 WITH rasp.zk7,zpk8 WITH rasp.zk8,zpk9 WITH rasp.zk9,zpk10 WITH rasp.zk10,zpk11 WITH rasp.zk11,zpk12 WITH rasp.zk12
               
               
               SELECT curPrnTarFond
               GO TOP
               DO WHILE !EOF()
                  rep_r=ALLTRIM(persved)
                  rep_r1='rasp.'+ALLTRIM(persved)
                  SELECT curTarJob 
                  REPLACE &rep_r WITH &rep_r1         
                  SELECT curPrnTarFond
                  SKIP
               ENDDO                                               
               SELECT curTarJob            
               DO countOkladVac 
            ELSE
               DO CASE
                  CASE kse_cx<=1
                       kvovac=1
                  CASE MOD(kse_cx,1)=0     
                       kvovac=INT(kse_cx)
                  CASE MOD(kse_cx,1)>0     
                       kvovac=INT(kse_cx)+1    
               ENDCASE               
               kvokse=kse_cx
               ksevac=0
               FOR i=1 TO kvovac
                   ksevac=IIF(kvokse<=1,kvokse,1)
                   SELECT curTarJob
                   APPEND BLANK
                   REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nKfVac,pkf WITH rasp.pkf,kse WITH ksevac,vac WITH .T.,patt WITH rasp.patt
                   IF i=1
                      IF lognzptot
                         REPLACE nzptot WITH rasp.nzptot
                      ELSE   
                         IF rasp.ksezotp#0
                            REPLACE zp1 WITH rasp.z1,zp2 WITH rasp.z2,zp3 WITH rasp.z3,zp4 WITH rasp.z4,zp5 WITH rasp.z5,zp6 WITH rasp.z6,zp7 WITH rasp.z7,zp8 WITH rasp.z8,zp9 WITH rasp.z9,zp10 WITH rasp.z10,zp11 WITH rasp.z11,zp12 WITH rasp.z12
                         ENDIF 
                      ENDIF    
                      REPLACE zpk1 WITH rasp.zk1,zpk2 WITH rasp.zk2,zpk3 WITH rasp.zk3,zpk4 WITH rasp.zk4,zpk5 WITH rasp.zk5,zpk6 WITH rasp.zk6,zpk7 WITH rasp.zk7,zpk8 WITH rasp.zk8,zpk9 WITH rasp.zk9,zpk10 WITH rasp.zk10,zpk11 WITH rasp.zk11,zpk12 WITH rasp.zk12                       
                         
                    
                   ENDIF
                   
                   SELECT curPrnTarFond
                   GO TOP
                   DO WHILE !EOF()
                      rep_r=ALLTRIM(persved)                      
                      rep_r1='rasp.'+ALLTRIM(persved)
                      SELECT curTarJob 
                      REPLACE &rep_r WITH &rep_r1         
                      SELECT curPrnTarFond                  
                      SKIP
                   ENDDO                                               
                   SELECT curTarJob   
                   DO countOkladVac
                   IF curTarJob.kse=1.AND.rasp.patt#0
                      REPLACE sumvr WITH varBaseSt/100*patt,kfVr WITH rasp.kfVr,pkfVr WITH rasp.pkfVr,vtime WITH rasp.vtime
                      sum_tot=0
                      FOR hi=1 TO 12
                          rep_cx='vr'+LTRIM(STR(hi))
                          repTime='sprtime.t'+LTRIM(STR(hi))
                          REPLACE &rep_cx WITH sumvr*IIF(SEEK(vTime,'sprtime',1),&repTime,0)
                          sum_tot=sum_tot+&rep_cx                     
                      ENDFOR
                      REPLACE sumvrtot WITH sum_tot   
                   ENDIF
                   
                   kvokse=kvokse-1
               ENDFOR
            ENDIF    
         ENDIF       
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
   SELECT datJob
   SET ORDER TO &ordOld   
ENDIF 
SELECT curTarJob
DO fltstructure WITH 'mtokl#0'
IF onlyVac
   SELECT curTarJob
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curTarJob
   DELETE FOR vac
ENDIF 
SELECT * FROM sprpodr INTO CURSOR curPrn READWRITE

ALTER TABLE curPrn ADD COLUMN m1 N(12,2)
ALTER TABLE curPrn ADD COLUMN m2 N(12,2)
ALTER TABLE curPrn ADD COLUMN m3 N(12,2)
ALTER TABLE curPrn ADD COLUMN m4 N(12,2)
ALTER TABLE curPrn ADD COLUMN m5 N(12,2)
ALTER TABLE curPrn ADD COLUMN m6 N(12,2)
ALTER TABLE curPrn ADD COLUMN m7 N(12,2)
ALTER TABLE curPrn ADD COLUMN m8 N(12,2)
ALTER TABLE curPrn ADD COLUMN m9 N(12,2)
ALTER TABLE curPrn ADD COLUMN m10 N(12,2)
ALTER TABLE curPrn ADD COLUMN m11 N(12,2)
ALTER TABLE curPrn ADD COLUMN m12 N(12,2)
ALTER TABLE curPrn ADD COLUMN mtot N(14,2)
ALTER TABLE curPrn ADD COLUMN lDel L
ALTER TABLE curPrn ADD COLUMN lHead L
ALTER TABLE curPrn ADD COLUMN nGr N(1)
INDEX ON np TAG T1
INDEX ON kod TAG T2
REPLACE nGr WITH 1 ALL
SET ORDER TO 1
*----��������� ������-------
SELECT 	curPrn
SELECT curPrn
APPEND BLANK
REPLACE name WITH '�������� �����',np WITH 0,lDel WITH .T.,lHead WITH .T.
SCAN ALL
     SELECT curTarJob
     SUM mtokl TO mtoklcx FOR kp=curPrn.kod
     SELECT curPrn
     FOR i=1 TO 12
         IF i>=MONTH(dateTar)
            repm='m'+LTRIM(STR(i))
            REPLACE &repm WITH mtoklcx,mtot WITH mtot+mtoklcx 
         ENDIF
     ENDFOR
ENDSCAN
SUM mtot,m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12 TO mtotcx,m1cx,m2cx,m3cx,m4cx,m5cx,m6cx,m7cx,m8cx,m9cx,m10cx,m11cx,m12cx
APPEND BLANK
REPLACE name WITH '�����',np WITH 499
REPLACE mtot WITH mtotcx,m1 WITH m1cx,m2 WITH m2cx,m3 WITH m3cx,m4 WITH m4cx,m5 WITH m5cx,m6 WITH m6cx,m7 WITH m7cx,m8 WITH m8cx,m9 WITH m9cx,m10 WITH m10cx,m11 WITH m11cx,m12 WITH m12cx
*-------�� ���� ������ � ���������� ������� ����
SELECT curPrn
APPEND BLANK
REPLACE name WITH '�������� � ���������� �����',np WITH 499,lDel WITH .T.,lHead WITH .T.

APPEND BLANK
REPLACE name WITH '�������� �� ���� ������ � ��������� ������������ (���� �27 �.2)',np WITH 500,nGr WITH 2
APPEND BLANK
REPLACE name WITH '������� �� ����������� ���������� ������� ����',np WITH 600,nGr WITH 2

SELECT curTarJob
SCAN ALL     
     DO CASE
        CASE EMPTY(per_date)
             SELECT curPrn
             SEEK 500
             FOR i=1 TO 12
                 IF i>=MONTH(dateTar)                  
                    repm='m'+LTRIM(STR(i))
                    REPLACE &repm WITH &repm+curTarJob.mstsum,mtot WITH mtot+curTarJob.mstsum
                 ENDIF   
             ENDFOR 
        CASE !EMPTY(per_date)
             SELECT curPrn
             SEEK 500
             FOR i=1 TO 12
                 IF i>=MONTH(dateTar)                  
                    repm='m'+LTRIM(STR(i))
                    repSumMonth=0
                    DO CASE 
                       CASE i<MONTH(curTarJob.per_date)
                            repSumMonth=curTarJob.mstsum
                       CASE i=MONTH(curTarJob.per_date)                
                            nmonth=MONTH(curTarJob.per_date)
                            year_ch=YEAR(curTarJob.per_date)
                            STORE 0 TO kvo_old,kvo_new
                            kvoday=IIF(nmonth=2,IIF(MOD(year_ch,4)=0,29,28),IIF(INLIST(nmonth,4,6,9,11),30,31))
                            oneday=DOW(CTOD('01.'+STR(nmonth,2)+'.'+STR(year_ch,4)))
                            oneday=IIF(oneday=1,7,oneday-1)
                            FOR cx=1 TO kvoday
                                oneday=DOW(CTOD(STR(cx,2)+'.'+STR(nmonth,2)+'.'+STR(year_ch,4)))
                                oneday=IIF(oneday=1,7,oneday-1)
                                IF cx<DAY(curTarJob.per_date)
                                   IF oneday<6.AND.!SEEK(cx+nmonth/100,'fete',1)
                                       kvo_old=kvo_old+1
                                   ENDIF
                                ELSE
                                   IF oneday<6.AND.!SEEK(cx+nmonth/100,'fete',1)
                                      kvo_new=kvo_new+1
                                   ENDIF
                                ENDIF
                            ENDFOR
                            sumOld=(varBaseSt/100*curTarJob.stpr*curTarJob.kse)/(kvo_old+kvo_new)*kvo_old
                            sumNew=(varBaseSt/100*curTarJob.st_per*curTarJob.kse)/(kvo_old+kvo_new)*kvo_new
                            repsumMonth=sumOld+sumNew                                                                                               
                       CASE i>MONTH(curTarJob.per_date) 
                            repSumMonth=varBaseSt/100*curTarJob.st_per*curTarJob.kse
                    ENDCASE                     
                    REPLACE &repm WITH &repm+repsummonth,mtot WITH mtot+repsummonth
                 ENDIF   
             ENDFOR               
     ENDCASE
     IF !EMPTY(curtarjob.sumvrtot)
        FOR i=MONTH(dateTar) TO 12
            repvr='curtarJob.vr'+LTRIM(STR(i))
            repm='m'+LTRIM(STR(i))
            SELECT curPrn
            SEEK 600
            REPLACE &repm WITH &repm+&repvr,mtot WITH mtot+&repvr
            
        ENDFOR 
     ENDIF
     SELECT curTarJob
ENDSCAN
*------ ������ �������� � �������
SELECT curNad
numrecnad=501
SCAN ALL    
     sumvedcx=ALLTRIM(sumved)
     SELECT curTarJob
     SUM &sumvedcx TO repsumcx     
     SELECT curPrn
     APPEND BLANK 
     REPLACE name WITH curNad.nHead,np WITH numrecnad,nGr WITH 2
     SEEK numrecnad
     FOR i=MONTH(dateTar) TO 12     
         repm='m'+LTRIM(STR(i))     
         REPLACE &repm WITH repsumcx,mtot WITH mtot+repsumcx
     ENDFOR 
     numrecnad=numrecnad+1
     SELECT curNad
ENDSCAN 
*-------������ �������� � ������
*ON ERROR DO erSup
SELECT curPrn
APPEND BLANK
REPLACE np WITH 601,name WITH '������� �� ������ �������� � ������',nGr WITH 2
*APPEND BLANK
*REPLACE np WITH 602,name WITH '������� �� ������ �������� �� �����',nGr WITH 2
IF lognzptot
   SELECT curtarjob
   SUM nzptot TO sumOtpTot 
   SELECT curPrn  
   SEEK 601
   REPLACE mtot WITH sumOtpTot

   FOR i=MONTH(dateTar) TO 12
       *sumotpcx='zp'+LTRIM(STR(i))
       sumkurscx='zpk'+LTRIM(STR(i))
       repm='m'+LTRIM(STR(i))
       SELECT curTarJob
       SUM nzptot/100*kfotp(i),&sumkurscx TO sumOtpMonth,sumKursMonth
       SELECT curPrn
       SEEK 601
       REPLACE &repm WITH sumOtpMonth,mtot WITH sumOtpTot
 *      SEEK 602
 *      REPLACE &repm WITH sumKursMonth,mtot WITH mtot+sumKursMonth 
   ENDFOR
ELSE 
   FOR i=MONTH(dateTar) TO 12
       sumotpcx='zp'+LTRIM(STR(i))
       sumkurscx='zpk'+LTRIM(STR(i))
       repm='m'+LTRIM(STR(i))
       SELECT curTarJob
       SUM &sumotpcx,&sumkurscx TO sumOtpMonth,sumKursMonth
       SELECT curPrn
       SEEK 601
       REPLACE &repm WITH sumOtpMonth,mtot WITH mtot+sumOtpMonth 
  *     SEEK 602
  *     REPLACE &repm WITH sumKursMonth,mtot WITH mtot+sumKursMonth 
   ENDFOR
ENDIF    
*ON ERROR 
*-------������ � �����������
SELECT curPrn
APPEND BLANK
REPLACE np WITH 603,name WITH '������� �� ������ � ������ �����',nGr WITH 2
APPEND BLANK
REPLACE np WITH 604,name WITH '������� �� ������ � ����������� ���',nGr WITH 2
SELECT curTarRasp
SCAN ALL
     FOR i=MONTH(dateTar) TO 12
         sumnight='curTarRasp.night'+LTRIM(STR(i))
         sumfete='curtarrasp.fete'+LTRIM(STR(i))
         repm='m'+LTRIM(STR(i))
         IF !EMPTY(&sumNight)
            sumrep=VAL(SUBSTR(&sumnight,13,8))*VAL(SUBSTR(&sumnight,37,6))*VAL(SUBSTR(&sumnight,11,2))
            sumrep1=VAL(SUBSTR(&sumnight,49,8))*VAL(SUBSTR(&sumnight,43,6))*VAL(SUBSTR(&sumnight,57,2))         
            SELECT curPrn
            SEEK 603
            REPLACE &repm WITH &repm+sumrep+sumrep1,mtot WITH mtot+sumrep+sumrep1
         ENDIF  
         IF !EMPTY(&sumFete)
            sumrep=VAL(SUBSTR(&sumFete,21,8))
            SELECT curPrn
            SEEK 604
            REPLACE &repm WITH &repm+sumrep,mtot WITH mtot+sumrep
         ENDIF      
     ENDFOR
     SELECT curTarRasp 
ENDSCAN 
SELECT curPrn
SUM mtot,m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12 TO mtotcx,m1cx,m2cx,m3cx,m4cx,m5cx,m6cx,m7cx,m8cx,m9cx,m10cx,m11cx,m12cx FOR nGr=2
APPEND BLANK
REPLACE name WITH '�����',np WITH 700
REPLACE mtot WITH mtotcx,m1 WITH m1cx,m2 WITH m2cx,m3 WITH m3cx,m4 WITH m4cx,m5 WITH m5cx,m6 WITH m6cx,m7 WITH m7cx,m8 WITH m8cx,m9 WITH m9cx,m10 WITH m10cx,m11 WITH m11cx,m12 WITH m12cx

SUM mtot,m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12 TO mtotcx,m1cx,m2cx,m3cx,m4cx,m5cx,m6cx,m7cx,m8cx,m9cx,m10cx,m11cx,m12cx FOR INLIST(nGr,1,2)
APPEND BLANK
REPLACE name WITH '�����',np WITH 701
REPLACE mtot WITH mtotcx,m1 WITH m1cx,m2 WITH m2cx,m3 WITH m3cx,m4 WITH m4cx,m5 WITH m5cx,m6 WITH m6cx,m7 WITH m7cx,m8 WITH m8cx,m9 WITH m9cx,m10 WITH m10cx,m11 WITH m11cx,m12 WITH m12cx
SELECT curPrn
DELETE FOR mtot=0.AND.!lDel
SET ORDER TO 1
GO TOP
DO procForPrintAndPreview WITH 'repSvod','������� ��������� �� ���������, �������� � ������ ��������'
**********************************************************************************
*          ������� � excel
**********************************************************************************
PROCEDURE svodToExcel
DO startPrnToExcel WITH 'fSupl'
     
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=25
     .Columns(2).ColumnWidth=10
     .Columns(3).ColumnWidth=8
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
     .Columns(8).ColumnWidth=8
     .Columns(9).ColumnWidth=8
     .Columns(10).ColumnWidth=8
     .Columns(11).ColumnWidth=8
     .Columns(12).ColumnWidth=8
     .Columns(13).ColumnWidth=8
     .Columns(14).ColumnWidth=8         
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1    
     .Range(.Cells(rowcx,1),.Cells(rowcx,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������ ��������� �� �������, ���������, �������� � ������ �������� �� '+STR(YEAR(dateTar),4)+' �.'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH        
     rowcx=rowcx+1  
                         
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ �������������, ��������, ������ � ������ ������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�����'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                                                                           
             
     .cells(rowcx,3).Value='1'
     .cells(rowcx,4).Value='2'
     .cells(rowcx,5).Value='3'
     .cells(rowcx,6).Value='4'
     .cells(rowcx,7).Value='5'
     .cells(rowcx,8).Value='6'
     .cells(rowcx,9).Value='7'
     .cells(rowcx,10).Value='8'
     .cells(rowcx,11).Value='9'
     .cells(rowcx,12).Value='10'
     .cells(rowcx,13).Value='11'
     .cells(rowcx,14).Value='12'     
  
     .Range(.Cells(rowcx,1),.Cells(rowcx,14)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter
     numberRow=rowcx+1  
     rowtop=numberRow         
     SELECT curPrn
     DO storezeropercent
     GO TOP
     kpold=0
     SCAN ALL
          IF lHead
             .Range(.Cells(numberRow,1),.Cells(numberRow,14)).Select
             objExcel.Selection.HorizontalAlignment=xlLeft
             objExcel.Selection.VerticalAlignment=1
             objExcel.Selection.WrapText=.T.
             objExcel.Selection.Interior.ColorIndex=37
          ENDIF 
          .Cells(numberRow,1).Value=name                                   
          .Cells(numberRow,2).Value=IIF(mTot#0,mTot,'')                                       
          .Cells(numberRow,2).NumberFormat='0.00'                                          
          .Cells(numberRow,3).Value=IIF(m1#0,m1,'')                                       
          .Cells(numberRow,3).NumberFormat='0.00'                                          
          .Cells(numberRow,4).Value=IIF(m2#0,m2,'')                                       
          .Cells(numberRow,4).NumberFormat='0.00'                                                    
          .Cells(numberRow,5).Value=IIF(m3#0,m3,'')                                       
          .Cells(numberRow,5).NumberFormat='0.00'                                                   
          .Cells(numberRow,6).Value=IIF(m4#0,m4,'')                                       
          .Cells(numberRow,6).NumberFormat='0.00'                                                              
          .Cells(numberRow,7).Value=IIF(m5#0,m5,'')                                       
          .Cells(numberRow,7).NumberFormat='0.00'                                                                    
          .Cells(numberRow,8).Value=IIF(m6#0,m6,'')                                       
          .Cells(numberRow,8).NumberFormat='0.00'                                                  
          .Cells(numberRow,9).Value=IIF(m7#0,m7,'')                                       
          .Cells(numberRow,9).NumberFormat='0.00'                                        
          .Cells(numberRow,10).Value=IIF(m8#0,m8,'')                                       
          .Cells(numberRow,10).NumberFormat='0.00'
          .Cells(numberRow,11).Value=IIF(m9#0,m9,'')                                       
          .Cells(numberRow,11).NumberFormat='0.00' 
          .Cells(numberRow,12).Value=IIF(m10#0,m10,'')                                       
          .Cells(numberRow,12).NumberFormat='0.00'
          .Cells(numberRow,13).Value=IIF(m11#0,m11,'')                                       
          .Cells(numberRow,13).NumberFormat='0.00' 
          .Cells(numberRow,14).Value=IIF(m12#0,M12,'')
          .Cells(numberRow,14).NumberFormat='0.00'                                          
          numberRow=numberRow+1
          DO fillpercent WITH 'fSupl' 
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,14)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1      
      
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,14)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'  
objExcel.Visible=.T.

**********************************************************************************************
*   ������� �� ���
**********************************************************************************************
PROCEDURE ppsprn
*persVed=''
*sumVed=''
*parPers=persVed
*parSum=sumVed

* maxkat=���-�� ��������� ���������
* sum_podr - ���� �� ���������� 
* sumPodrKat - ���� �� ���������� � ������� ���������

* sumTot - ���� �����
* sumTotKat - ���� ����� �� ���������� ���������
* sumKpp -���� �� ����������������
itKat=formBase.log_It
logKom=formBase.log_kom

SELECT sprkat
COUNT TO maxKat
DIMENSION nSum(1)

DIMENSION sum_podr(2),sumTot(2)
DIMENSION sumPodrKat(maxKat,2),sumTotKat(maxKat,2)
STORE 0 TO sum_podr,sumTot,sumPodrKat,sumTotkat

DIMENSION sumPodrKpp(2),sumPodrKpp1(2)  &&����� � ����������������� � ������ ��������������
STORE 0 TO sumPodrKpp,sumPodrKpp1

DIMENSION sumPodrKatKpp(maxKat,2),sumPodrKatKpp1(maxKat,2)  &&����� � ����������������� � ������ �������������� �� ���������� ���������
STORE 0 TO sumPodrKatKpp,sumPodrKatKpp1

IF USED('curprn')
   SELECT curPrn
   USE   
ENDIF
SELECT datjob
SET FILTER TO
SELECT * FROM datjob INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN pprn N(6,2)
ALTER TABLE curPrn ADD COLUMN sprn N(12,2)
ALTER TABLE curPrn ADD COLUMN nIt N(1)
ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
SELECT curPrn
DELETE FOR tokl=0
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DELETE FOR date_in>dateTar
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL


*----------------------------------- �������������� ���������� ��������---------------------------------------------
SELECT rasp
SCAN ALL
     IF ksepps#0.AND.nspps#0
        SELECT curprn
        APPEND BLANK
        REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kse WITH rasp.ksepps,fio WITH '���������',vac WITH .T.,npps WITH rasp.npps,nopps WITH rasp.nopps,nspps WITH rasp.nspps
     ENDIF 
     SELECT rasp
ENDSCAN
DO fltstructure WITH 'kse#0.AND.nspps#0','curprn'
IF onlyVac
   SELECT curPrn
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curPrn
   DELETE FOR vac
ENDIF

SELECT curPrn
DELETE FOR nspps=0
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,np) ALL
SELECT sprpodr      
SCAN ALL
     SELECT curprn
     SUM kse,nspps TO ksecx,sprncx FOR kp=sprpodr.kod
     IF sprncx#0
        APPEND BLANK
        REPLACE fio WITH '�����',kp WITH sprpodr.kod,np WITH sprpodr.np,nd WITH 500,nspps WITH sprncx,kse WITH ksecx,nIt WITH 1     
     ENDIF
     IF itkat
        FOR i=1 TO maxKat
            SELECT curPrn
            SUM kse,nspps TO ksecx,sprncx FOR kp=sprpodr.kod.AND.kat=kod_kat(i)  
            APPEND BLANK
            REPLACE nIt WITH 2,kp WITH sprpodr.kod,np WITH sprpodr.np,kd WITH 999,kat WITH kod_kat(i),fio WITH name1_kat(i),nd WITH 700+i,;
                    nspps WITH sprncx,kse WITH ksecx                    
           
        ENDFOR              
      ENDIF     
     SELECT sprpodr
ENDSCAN
SELECT curPrn
DELETE FOR nspps=0
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
GO TOP
nppcx=1
kpOld=kp
DO WHILE !EOF()  
   REPLACE npp WITH nppcx
   nppcx=nppcx+1
   SKIP
   IF kpOld#kp
      kpOld=kp
      nppcx=1 
   ENDIF  
ENDDO 
SELECT curprn
SUM kse,nspps TO ksecx,sprncx FOR nIt=0
IF sprncx#0
   APPEND BLANK
   REPLACE fio WITH IIF(formbase.lItr,'�����','�����'),kp WITH 999,np WITH 999,nd WITH 500,nspps WITH sprncx,kse WITH ksecx,nIt WITH 9  
   IF formbase.litr
      FOR i=1 TO max_tr
          SUM kse,nspps TO ksecx,sprncx FOR nIt=0.AND.tr=i
          IF ksecx#0
             APPEND BLANK
             REPLACE fio WITH name_tr(i),kp WITH 999,np WITH 999,nd WITH 500,nspps WITH sprncx,kse WITH ksecx,nIt WITH 9  
          ENDIF 
      ENDFOR
   ENDIF
   FOR i=1 TO maxKat
       SELECT curPrn
       SUM kse,nspps TO ksecx,sprncx FOR kat=kod_kat(i).AND.nIt=0  
       IF ksecx#0
          APPEND BLANK
          REPLACE nIt WITH 8,kp WITH 999,np WITH 999,kd WITH 999,kat WITH kod_kat(i),fio WITH IIF(formbase.lItr,UPPER(name1_kat(i)),name1_kat(i)),nd WITH 700+i,;
                  nspps WITH sprncx,kse WITH ksecx 
          
          IF formbase.litr
             FOR ix=1 TO max_tr
                 SUM kse,sprn TO ksetr,sprntr FOR nIt=0.AND.tr=ix.AND.kat=i
                 IF ksecx#0
                    APPEND BLANK
                    REPLACE fio WITH name_tr(ix),kp WITH 999,np WITH 999,nd WITH 700+i,nspps WITH sprntr,kse WITH ksetr,nIt WITH 9  
                 ENDIF 
             ENDFOR
          ENDIF                                                
       ENDIF            
   ENDFOR     
ENDIF	
DO procForPrintAndPreview WITH 'reppps','��������������� ������'
SELECT curPrn
USE
SELECT people
**********************************************************************************************
PROCEDURE ppsToExcel
DO startPrnToExcel WITH 'fSupl'            
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=25
     .Columns(4).ColumnWidth=8
     .Columns(5).ColumnWidth=8
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8    
     rowcx=3     
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value=datshtat.office
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������ �������� �� ������� ������ �� ���'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH  
      
     rowcx=rowcx+1
     .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�� '+ALLTRIM(dim_month(MONTH(dateTar)))+' '+STR(YEAR(dateTar),4)+' �.'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
        
     rowcx=rowcx+1  
                                      
     .Range(.Cells(rowcx,1),.Cells(rowcx,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='� �/�'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���� �����, ������������ ������������� ������� ��� �������� ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH   
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,3)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='������������ ���������'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH      
    
        
      .Range(.Cells(rowcx,4),.Cells(rowcx,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'                   
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,5),.Cells(rowcx,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='����'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,6),.Cells(rowcx,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='%'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,7),.Cells(rowcx,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='�����'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
                                      
              
      rowcx=rowcx+1
      .cells(rowcx,1).Value='1'
      .cells(rowcx,2).Value='2'
      .cells(rowcx,3).Value='3'
      .cells(rowcx,4).Value='4'
      .cells(rowcx,5).Value='5'
      .cells(rowcx,6).Value='6'
      .cells(rowcx,7).Value='7'

  
      .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      rowtop=numberRow         
      SELECT curPrn
      DO storezeropercent
      GO TOP
      kpold=0
      SCAN ALL
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,7)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(curprn.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              kpold=kp
           ENDIF 
           .Cells(numberRow,1).Value=IIF(nIt=0,curprn.npp,'')                                    
           .Cells(numberRow,2).Value=curprn.fio                                       
           .Cells(numberRow,3).Value=IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'')                                                                        
           .Cells(numberRow,4).Value=kse                                      
           .Cells(numberRow,4).NumberFormat='0.00'           
           .Cells(numberRow,5).Value=IIF(nIt=0,nopps,'')                                      
           .Cells(numberRow,6).Value=IIF(nIt=0,LTRIM(STR(nspps,6,2))+'%','')
           .Cells(numberRow,7).Value=IIF(nIt=0,curprn.nspps,'')                                      

           numberRow=numberRow+1
           DO fillperCent WITH 'fSupl'
                   
      ENDSCAN                                 
      .Range(.Cells(3,1),.Cells(numberRow-1,7)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
      
      IF logKom
         numberRow=numberRow+1
         .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
         With objExcel.Selection
              .MergeCells=.T.
              .HorizontalAlignment=xlLeft
              .VerticalAlignment=1
              .WrapText=.T.
              .Value='��������'                   
              .Font.Name='Times New Roman'   
              .Font.Size=9
         ENDWITH 
         numberRow=numberRow+1        
      ENDIF
          
      .Range(.Cells(rowcx,1),.Cells(numberRow-1,7)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
ENDWITH    
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'
objExcel.Visible=.T.
***************************************************************************************************************
PROCEDURE totalvacance
RESTORE FROM rashotp ADDITIVE
repzp1='rasp.z'+LTRIM(STR(MONTH(dateTar)))
repzp='curprn.zp'+LTRIM(STR(MONTH(dateTar)))
repd='curprn.d'+LTRIM(STR(MONTH(dateTar)))
repd1='rasp.m'+LTRIM(STR(MONTH(dateTar)))
IF USED('curprn')
   SELECT curPrn
   USE   
ENDIF

IF USED('curJobSup')
   SELECT curJobsup
   USE   
ENDIF

IF USED('curOtp')
   SELECT curOtp
   USE
ENDIF

CREATE CURSOR curotp (nstr C(15),kse N(7,2),mtokl N(10,2),mtokl1 N(10,2),mtokl2 N(10,2),mstsum N(10,2),mchir N(10,2),mkat N(10,2),mvto N(10,2),mcharw N(10,2),mmain N(10,2),mmain2 N(10,2),msupl N(10,2),mtot N(10,2),mRound N(10,2))
SELECT datjob
SET FILTER TO
SELECT * FROM datjob INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN nIt N(1)
ALTER TABLE curPrn ADD COLUMN kHours N(7,2)
ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
SELECT curPrn
DELETE FOR tokl=0
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DELETE FOR date_in>dateTar
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL

*----------------------------------- �������������� ���������� ��������---------------------------------------------
IF formbase.avt_vac
   SELECT rasp
   SET FILTER TO ksezotp>0
   GO TOP
   DO WHILE !EOF()
      IF rasp.ksezotp#0        
         SELECT curPrn
         APPEND BLANK
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH rasp.ksezotp,vac WITH .T.,;
                 np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,tr WITH 1,lokl WITH rasp.lokl,zptot WITH rasp.totzp,zpday WITH rasp.zpday,srzp WITH rasp.srzp,dotp WITH rasp.dotp,dtot WITH rasp.dotp
                 
         tar_ok=0
         tar_ok=varBaseSt*namekf*IIF(pkf#0,pkf,1)                      
         REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2),;
                 pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2             
          
         REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                 mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse        
         FOR i=1 TO 12
             repzp1='rasp.z'+LTRIM(STR(i))
             repzp='curprn.zp'+LTRIM(STR(i))
             repd='curprn.d'+LTRIM(STR(i))
             repd1='rasp.m'+LTRIM(STR(i))
             REPLACE &repzp WITH &repzp1,&repd WITH &repd1 
         ENDFOR               
         
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
ENDIF  
SELECT rasp
SET FILTER TO 
SELECT * FROM curprn INTO CURSOR curJobSupl READWRITE
SELECT curJobSupl
INDEX ON STR(kp,3)+STR(kd,3) TAG T1


SELECT curprn
DELETE FOR tokl=0
DO fltstructure WITH 'kse#0','curprn'
IF onlyVac
   SELECT curPrn
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curPrn
   DELETE FOR vac
ENDIF

IF !rashotp(4)
   SELECT curotp
   FOR i=1 TO 12
       APPEND BLANK
       REPLACE nstr WITH dim_month(i)
   ENDFOR
   SELECT curprn
   SCAN ALL
        SELECT curprn
        IF zptot>0           
           FOR i=1 TO 12
               STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvOkl1,nAvOkl2
               repzp='curprn.zp'+LTRIM(STR(i))
               repd='curprn.d'+LTRIM(STR(i))
               DO CASE 
                  CASE curprn.lokl
                       nAvOkl=&repzp
                       nAvOkl1=&repzp
                  CASE !curprn.lokl
                       nAvOkl=IIF(rashotp(1),mtokl/rashotp(2),IIF(dotp#0,mtokl/dotp,0))*&repd
                       nAvOkl2=nAvOkl
                       nAvSt=IIF(rashotp(1),mstsum/rashotp(2),IIF(dotp#0,mstsum/dotp,0))*&repd
                       nAvChir=IIF(rashotp(1),mchir/rashotp(2),IIF(dotp#0,mchir/dotp,0))*&repd
                       nAvKat=IIF(rashotp(1),mkat/rashotp(2),IIF(dotp#0,mkat/dotp,0))*&repd
                       nAvVto=IIF(rashotp(1),mvto/rashotp(2),IIF(dotp#0,mvto/dotp,0))*&repd                       
                       nAvCharw=IIF(rashotp(1),mcharw/rashotp(2),IIF(dotp#0,mcharw/dotp,0))*&repd
                       nAvMain=IIF(rashotp(1),mmain/rashotp(2),IIF(dotp#0,mmain/dotp,0))*&repd                       
                       nAvMain2=IIF(rashotp(1),mmain2/rashotp(2),IIF(dotp#0,mmain2/dotp,0))*&repd                                              
                       nAvSupl=IIF(rashotp(1),mtokl*rashotp(3)/rashotp(2),IIF(dotp#0,mtokl*rashotp(3)/dotp,0))*&repd                                                                
               ENDCASE 
               SELECT curotp
               GO i
               REPLACE mtokl WITH mtokl+nAvOkl,mtokl1 WITH mtokl1+nAvOkl1,mtokl2 WITH mtokl2+nAvOkl2,mstsum WITH mstSum+nAvSt,mchir WITH mchir+nAvChir,;
                       mkat WITH mkat+nAvKat,mvto WITH mvto+nAvVto,mcharw WITH mcharw+nAvCharw,mmain WITH mmain+nAvMain,mmain2 WITH mmain2+nAvMain2,msupl WITH msupl+nAvSupl,mtot WITH mtot+&repzp
               SELECT curprn
           ENDFOR 
        ENDIF 
   ENDSCAN  
ELSE 
   FOR i=1 TO 12 
       repzp1='rasp.z'+LTRIM(STR(i))
       repzp='curprn.zp'+LTRIM(STR(i))
       repd='curprn.d'+LTRIM(STR(i))
       repd1='rasp.m'+LTRIM(STR(i))
       SELECT curotp
       APPEND BLANK
       REPLACE nstr WITH dim_month(i)
       SELECT curprn
       STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvOkl1,nAvOkl2
       SCAN ALL
            IF &repZp#0
               SELECT curJobsupl
               SEEK STR(curprn.kp,3)+STR(curprn.kd,3)
               STORE 0 TO nkse,nAvOkl,nAvOkl1,nAvOkl2,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl
               SCAN WHILE kp=curprn.kp.AND.kd=curprn.kd
                    nkse=nkse+kse
                    nAvOkl=nAvOkl+mtokl
                    IF !curprn.lOkl
                       nAvOkl2=nAvOkl2+mtokl   
                       nAvSt=nAvSt+mstsum
                       nAvChir=nAvChir+mchir
                       nAvKat=nAvKat+mkat
                       nAvVto=nAvVto+mvto
                       nAvCharw=nAvCharw+mcharw
                       nAvMain=nAvMain+mmain
                       nAvMain2=nAvMain2+mmain2
                       nAvSupl=nAvSupl+mtokl*rashotp(3)
                    ELSE  
                       nAvOkl1=nAvOkl1+mtokl   
                    ENDIF   
               ENDSCAN           
               nAvOkl=IIF(nkse=0,0,IIF(nkse<1,nAvOkl,nAvOkl/nkse))
               nAvOkl1=IIF(nkse=0,0,IIF(nkse<1,nAvOkl1,nAvOkl1/nkse))
               nAvOkl2=IIF(nkse=0,0,IIF(nkse<1,nAvOkl2,nAvOkl2/nkse))
               
               
               nAvSt=IIF(nkse=0,0,IIF(nkse<1,nAvSt,nAvSt/nkse))
               nAvChir=IIF(nkse=0,0,IIF(nkse<1,nAvChir,nAvChir/nkse))
               nAvKat=IIF(nkse=0,0,IIF(nkse<1,nAvKat,nAvKat/nkse))
               nAvVto=IIF(nkse=0,0,IIF(nkse<1,nAvVto,nAvVto/nkse))
               nAvCharw=IIF(nkse=0,0,IIF(nkse<1,nAvCharw,nAvCharw/nkse))
               nAvMain=IIF(nkse=0,0,IIF(nkse<1,nAvMain,nAvMain/nkse))
               nAvMain2=IIF(nkse=0,0,IIF(nkse<1,nAvMain2,nAvMain2/nkse))
               nAvSupl=IIF(nkse=0,0,IIF(nkse<1,nAvSupl,nAvSupl/nkse))                    
               SELECT curotp
               GO i
               REPLACE mtokl WITH mtokl+ROUND(nAvOkl/rashotp(2),2)*curprn.kse*&repd,mstSum WITH mstsum+ROUND(nAvSt/rashotp(2),2)*curprn.kse*&repd,mChir WITH mChir+ROUND(nAvChir/rashotp(2),2)*curprn.kse*&repd,;
                       mKat WITH mKat+ROUND(nAvKat/rashotp(2),2)*curprn.kse*&repd,mVto WITH mVto+ROUND(nAvVto/rashotp(2),2)*curprn.kse*&repd,mCharw WITH mCharw+ROUND(nAvCharw/rashotp(2),2)*curprn.kse*&repd,;
                       mMain WITH mMain+ROUND(nAvMain/rashotp(2),2)*curprn.kse*&repd,mMain2 WITH mmain2+ROUND(nAvmain2/rashotp(2),2)*curprn.kse*&repd,mSupl WITH mSupl+ROUND(nAvSupl/rashotp(2),2)*curprn.kse*&repd,;
                       mtokl1 WITH mtokl1+ROUND(nAvOkl1/rashotp(2),2)*curprn.kse*&repd,mtokl2 WITH mtokl2+ROUND(nAvOkl2/rashotp(2),2)*curprn.kse*&repd                   
           
               REPLACE kse WITH kse+curprn.kse,mtot WITH mtot+&repzp 
               SELECT curprn    
            ENDIF
       ENDSCAN
   ENDFOR
ENDIF 
SELECT curotp
REPLACE mRound WITH mTot-mTokl-mStsum-mChir-mKat-mVto-mCharw-mMain-mMain2-msupl ALL
SUM mtokl,mstsum,mChir,mkat,mVto,mCharw,mMain,mMain2,mSupl,mtot,mRound,kse,mtokl1,mtokl2 TO mtokl_cx,mstsum_cx,mChir_cx,mkat_cx,mVto_cx,mCharw_cx,mMain_cx,mMain2_cx,mSupl_cx,mtot_cx,mRound_cx,kse_cx,mtokl1_cx,mtokl2_cx
APPEND BLANK
REPLACE nstr WITH '�����',mtokl WITH mtokl_cx,mStsum WITH mstsum_cx,mChir WITH mChir_cx,mKat WITH mkat_cx,mVto WITH mvto_cx,mCharw WITH mCharw_cx,mMain WITH mMain_cx,mMain2 WITH mMain2_cx,;
        mSupl WITH msupl_cx,mTot WITH mtot_cx,mRound WITH mround_cx,kse WITH kse_cx,mtokl1 WITH mtokl1_cx,mtokl2 WITH mtokl2_cx
GO TOP

DO procForPrintAndPreview WITH 'reptototp','�������� �� ������ ��������'
SELECT curotp
USE
SELECT curPrn
USE
SELECT people
***************************************************************************************************************
PROCEDURE totalcourse
RESTORE FROM rashkurs ADDITIVE
repzp1='rasp.zk'+LTRIM(STR(MONTH(dateTar)))
repzp='curprn.zpk'+LTRIM(STR(MONTH(dateTar)))
repd='curprn.d'+LTRIM(STR(MONTH(dateTar)))
repd1='rasp.m'+LTRIM(STR(MONTH(dateTar)))
IF USED('curprn')
   SELECT curPrn
   USE   
ENDIF

IF USED('curJobSup')
   SELECT curJobsup
   USE   
ENDIF

IF USED('curOtp')
   SELECT curOtp
   USE
ENDIF

CREATE CURSOR curotp (nstr C(15),kse N(7,2),mtokl N(10,2),mtokl1 N(10,2),mtokl2 N(10,2),mstsum N(10,2),mchir N(10,2),mkat N(10,2),mvto N(10,2),mcharw N(10,2),mmain N(10,2),mmain2 N(10,2),msupl N(10,2),mtot N(10,2),mRound N(10,2))
SELECT datjob
SET FILTER TO
SELECT * FROM datjob INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN nIt N(1)
ALTER TABLE curPrn ADD COLUMN kHours N(7,2)
ALTER TABLE curPrn ALTER COLUMN kse N(7,2)
SELECT curPrn
DELETE FOR tokl=0
DELETE FOR !SEEK(STR(kp,3)+STR(kd,3),'rasp',2)
DELETE FOR date_in>dateTar
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
REPLACE kat WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.kat,0) ALL

*----------------------------------- �������������� ���������� ��������---------------------------------------------
IF formbase.avt_vac
   SELECT rasp
   SET FILTER TO ksezotp>0
   GO TOP
   DO WHILE !EOF()
      IF rasp.ksezotp#0        
         SELECT curPrn
         APPEND BLANK
         REPLACE kp WITH rasp.kp,kd WITH rasp.kd,kat WITH rasp.kat,fio WITH '���������', kv WITH rasp.kv,kf WITH rasp.kfvac,nameKf WITH rasp.nkfvac,pkf WITH rasp.pkf,kse WITH rasp.ksezotp,vac WITH .T.,;
                 np WITH rasp.np,nd WITH rasp.nd,vtime WITH rasp.vtime,tr WITH 1,lkokl WITH rasp.lkokl
                 
         tar_ok=0
         tar_ok=varBaseSt*namekf*IIF(pkf#0,pkf,1)                      
         REPLACE tokl WITH tar_ok,mtokl WITH tokl*kse,staj_tar WITH dimConstVac(1,2),stpr WITH dimConstVac(2,2),;
                 pkat WITH rasp.pkat,pvto WITH rasp.pvto,pchir WITH rasp.pchir,pcharw WITH rasp.pcharw,pmain WITH rasp.pmain,pmain2 WITH rasp.pmain2             
          
         REPLACE mstsum WITH varBaseSt/100*stpr*kse,mkat WITH mtokl/100*pkat,mvto WITH mtokl/100*pvto,mchir WITH mtokl/100*pchir,;
                 mcharw WITH varBaseSt/100*pcharw*kse,mmain2 WITH varBaseSt/100*pmain2*kse,mmain WITH varBaseSt/100*pmain*kse        
         REPLACE srzpk WITH rasp.srzpk,zpdayk WITH rasp.zpdayk,dkurs WITH rasp.dkurs,zptotk WITH rasp.zptotk,lkokl WITH rasp.lkokl,pol1 WITH rasp.pol1,pol2 WITH rasp.pol2
         FOR i=1 TO 12                        
             repz='rasp.zk'+LTRIM(STR(i))
             repz1='zpk'+LTRIM(STR(i))              
             REPLACE &repz1 WITH &repz
         ENDFOR               
         
      ENDIF
      SELECT rasp
      SKIP
   ENDDO
ENDIF  
SELECT rasp
SET FILTER TO 
SELECT * FROM curprn INTO CURSOR curJobSupl READWRITE
SELECT curJobSupl
INDEX ON STR(kp,3)+STR(kd,3) TAG T1


SELECT curprn
DELETE FOR tokl=0
DO fltstructure WITH 'kse#0','curprn'
IF onlyVac
   SELECT curPrn
   DELETE FOR !vac
ENDIF
IF excludeVac
   SELECT curPrn
   DELETE FOR vac
ENDIF

IF !rashkurs(4)
   SELECT curotp
   FOR i=1 TO 12
       APPEND BLANK
       REPLACE nstr WITH dim_month(i)
   ENDFOR
   SELECT curprn
   SCAN ALL
        SELECT curprn
        polday=IIF(curprn.pol1>0.OR.curprn.pol2>0,6,12)
        polbeg=IIF(polday=12,1,IIF(pol1>0,1,7))
        IF zptotk>0           
           FOR i=polbeg TO 12
               STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvOkl1,nAvOkl2
               repzp='curprn.zpk'+LTRIM(STR(i))                             
               DO CASE 
                  CASE curprn.lkokl.AND.&repzp#0
                       nAvOkl=&repzp
                       nAvOkl1=&repzp
                  CASE !curprn.lkokl.AND.&repzp#0
                       nAvOkl=mtokl/rashkurs(1)*curprn.dkurs/polday
                       nAvOkl2=nAvOkl
                       nAvSt=mstsum/rashkurs(1)*curprn.dkurs/polday
                       nAvChir=mchir/rashkurs(1)*curprn.dkurs/polday
                       nAvKat=mkat/rashkurs(1)*curprn.dkurs/polday
                       nAvVto=mvto/rashkurs(1)*curprn.dkurs/polday
                       nAvCharw=mcharw/rashkurs(1)*curprn.dkurs/polday
                       nAvMain=mmain/rashkurs(1)*curprn.dkurs/polday
                       nAvMain2=mmain2/rashkurs(1)*curprn.dkurs/polday
                       nAvSupl=mtokl*rashkurs(3)/rashkurs(1)*curprn.dkurs/polday
               ENDCASE 
               SELECT curotp
               GO i
               REPLACE mtokl WITH mtokl+nAvOkl,mtokl1 WITH mtokl1+nAvOkl1,mtokl2 WITH mtokl2+nAvOkl2,mstsum WITH mstSum+nAvSt,mchir WITH mchir+nAvChir,;
                       mkat WITH mkat+nAvKat,mvto WITH mvto+nAvVto,mcharw WITH mcharw+nAvCharw,mmain WITH mmain+nAvMain,mmain2 WITH mmain2+nAvMain2,msupl WITH msupl+nAvSupl,mtot WITH mtot+&repzp
               SELECT curprn
           ENDFOR 
        ENDIF 
   ENDSCAN     
ELSE 
   FOR i=1 TO 12    
       repzp='curprn.zpk'+LTRIM(STR(i))
       SELECT curotp
       APPEND BLANK
       REPLACE nstr WITH dim_month(i)
       SELECT curprn
       STORE 0 TO nkse,nAvOkl,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl,nAvOkl1,nAvOkl2
       SCAN ALL
            IF &repZp#0
               SELECT curJobsupl
               SEEK STR(curprn.kp,3)+STR(curprn.kd,3)
               STORE 0 TO nkse,nAvOkl,nAvOkl1,nAvOkl2,nAvst,nAvChir,nAvKat,nAvVto,nAvCharw,nAvMain,nAvMain2,nAvSupl
               SCAN WHILE kp=curprn.kp.AND.kd=curprn.kd
                    nkse=nkse+kse
                    nAvOkl=nAvOkl+mtokl
                    IF !curprn.lkOkl
                       nAvOkl2=nAvOkl2+mtokl   
                       nAvSt=nAvSt+mstsum
                       nAvChir=nAvChir+mchir
                       nAvKat=nAvKat+mkat
                       nAvVto=nAvVto+mvto
                       nAvCharw=nAvCharw+mcharw
                       nAvMain=nAvMain+mmain
                       nAvMain2=nAvMain2+mmain2
                       nAvSupl=nAvSupl+mtokl*rashkurs(3)
                    ELSE  
                       nAvOkl1=nAvOkl1+mtokl   
                    ENDIF   
               ENDSCAN           
               nAvOkl=IIF(nkse=0,0,IIF(nkse<1,nAvOkl,nAvOkl/nkse))
               nAvOkl1=IIF(nkse=0,0,IIF(nkse<1,nAvOkl1,nAvOkl1/nkse))
               nAvOkl2=IIF(nkse=0,0,IIF(nkse<1,nAvOkl2,nAvOkl2/nkse))
               
               
               nAvSt=IIF(nkse=0,0,IIF(nkse<1,nAvSt,nAvSt/nkse))
               nAvChir=IIF(nkse=0,0,IIF(nkse<1,nAvChir,nAvChir/nkse))
               nAvKat=IIF(nkse=0,0,IIF(nkse<1,nAvKat,nAvKat/nkse))
               nAvVto=IIF(nkse=0,0,IIF(nkse<1,nAvVto,nAvVto/nkse))
               nAvCharw=IIF(nkse=0,0,IIF(nkse<1,nAvCharw,nAvCharw/nkse))
               nAvMain=IIF(nkse=0,0,IIF(nkse<1,nAvMain,nAvMain/nkse))
               nAvMain2=IIF(nkse=0,0,IIF(nkse<1,nAvMain2,nAvMain2/nkse))
               nAvSupl=IIF(nkse=0,0,IIF(nkse<1,nAvSupl,nAvSupl/nkse))                    
               polday=IIF(curprn.pol1>0.OR.curprn.pol2>0,6,12)
               SELECT curotp
               GO i
               REPLACE mtokl WITH mtokl+(ROUND(nAvOkl/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mstSum WITH mstsum+(ROUND(nAvSt/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mChir WITH mChir+(ROUND(nAvChir/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,;
                       mKat WITH mKat+(ROUND(nAvKat/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mVto WITH mVto+(ROUND(nAvVto/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mCharw WITH mCharw+(ROUND(nAvCharw/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,;
                       mMain WITH mMain+(ROUND(nAvMain/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mMain2 WITH mmain2+(ROUND(nAvmain2/rashkurs(1),2)*curprn.kse*curprn.dkurs)/12,mSupl WITH mSupl+(ROUND(nAvSupl/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,;
                       mtokl1 WITH mtokl1+(ROUND(nAvOkl1/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday,mtokl2 WITH mtokl2+(ROUND(nAvOkl2/rashkurs(1),2)*curprn.kse*curprn.dkurs)/polday                   
           
               REPLACE kse WITH kse+curprn.kse,mtot WITH mtot+&repzp 
               SELECT curprn    
            ENDIF
       ENDSCAN
   ENDFOR
ENDIF 
SELECT curotp
REPLACE mRound WITH mTot-mTokl-mStsum-mChir-mKat-mVto-mCharw-mMain-mMain2-msupl ALL
SUM mtokl,mstsum,mChir,mkat,mVto,mCharw,mMain,mMain2,mSupl,mtot,mRound,kse,mtokl1,mtokl2 TO mtokl_cx,mstsum_cx,mChir_cx,mkat_cx,mVto_cx,mCharw_cx,mMain_cx,mMain2_cx,mSupl_cx,mtot_cx,mRound_cx,kse_cx,mtokl1_cx,mtokl2_cx
APPEND BLANK
REPLACE nstr WITH '�����',mtokl WITH mtokl_cx,mStsum WITH mstsum_cx,mChir WITH mChir_cx,mKat WITH mkat_cx,mVto WITH mvto_cx,mCharw WITH mCharw_cx,mMain WITH mMain_cx,mMain2 WITH mMain2_cx,;
        mSupl WITH msupl_cx,mTot WITH mtot_cx,mRound WITH mround_cx,kse WITH kse_cx,mtokl1 WITH mtokl1_cx,mtokl2 WITH mtokl2_cx
GO TOP

DO procForPrintAndPreview WITH 'reptototp','�������� �� ������ ��������'
SELECT curotp
USE
SELECT curPrn
USE
SELECT people
***********************************************************************************************************
PROCEDURE totvacanceexcel
DO startPrnToExcel WITH 'fSupl'   
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)   
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=15
     .Columns(2).ColumnWidth=10
     .Columns(3).ColumnWidth=10
     .Columns(4).ColumnWidth=10
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=10
     .Columns(7).ColumnWidth=10   
     .Columns(8).ColumnWidth=10
     .Columns(9).ColumnWidth=10
     .Columns(10).ColumnWidth=10
     .Columns(11).ColumnWidth=10
     .Columns(12).ColumnWidth=10
     .Columns(13).ColumnWidth=10
     .Columns(14).ColumnWidth=10    
         
     rowcx=1     
     .Range(.Cells(rowcx,1),.Cells(rowcx+1,1)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�����'     
     ENDWITH  
     .Range(.Cells(rowcx,1),.Cells(rowcx+1,2)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�����'
     ENDWITH  
     .Range(.Cells(rowcx,3),.Cells(rowcx,5)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='�����'
     ENDWITH  
     
     .cells(rowcx+1,3).Value='�����'
     .cells(rowcx+1,4).Value='��� ���������'
     .cells(rowcx+1,5).Value='� ����������'
     .Range(.Cells(rowcx,6),.Cells(rowcx+1,6)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.         
          .Value='����'     
     ENDWITH  
     .Range(.Cells(rowcx,7),.Cells(rowcx+1,7)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.          
          .Value='���.����.'
     ENDWITH  
     .Range(.Cells(rowcx,8),.Cells(rowcx+1,8)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.          
          .Value='���������'     
     ENDWITH  
     .Range(.Cells(rowcx,9),.Cells(rowcx+1,9)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.         
          .Value='���'
     ENDWITH       
     .Range(.Cells(rowcx,10),.Cells(rowcx+1,10)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.          
          .Value='������ �-�'     
     ENDWITH  
     .Range(.Cells(rowcx,11),.Cells(rowcx+1,11)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.         
          .Value='���� 52 � 6.1' 
     ENDWITH  
     .Range(.Cells(rowcx,12),.Cells(rowcx+1,12)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.         
          .Value='���� 52 � 6.2'     
     ENDWITH  
     .Range(.Cells(rowcx,13),.Cells(rowcx+1,13)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.          
          .Value='�����. � ������.'
     ENDWITH   
     .Range(.Cells(rowcx,14),.Cells(rowcx+1,14)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.          
          .Value='����������'
     ENDWITH      
     .Range(.Cells(rowcx,1),.Cells(rowcx+1,14)).Select  
     WITH objExcel.Selection          
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
     ENDWITH 
  
     * .Range(.Cells(rowcx,1),.Cells(rowcx,7)).Select
     * objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+2  
      rowtop=numberRow         
      SELECT curotp
      DO storezeropercent
      GO TOP
      kpold=0
      SCAN ALL          
           .Cells(numberRow,1).Value=nstr                                    
           .Cells(numberRow,2).Value=IIF(mtot#0,mtot,'')
           .Cells(numberRow,3).Value=IIF(mtokl#0,mtokl,'')
           .Cells(numberRow,4).Value=IIF(mtokl1#0,mtokl1,'')
           .Cells(numberRow,5).Value=IIF(mtokl2#0,mtokl2,'')
           .Cells(numberRow,6).Value=IIF(mstsum#0,mstsum,'')
           .Cells(numberRow,7).Value=IIF(mchir#0,mchir,'')
           .Cells(numberRow,8).Value=IIF(mkat#0,mkat,'')
           .Cells(numberRow,9).Value=IIF(mvto#0,mvto,'')
           .Cells(numberRow,10).Value=IIF(mcharw#0,mcharw,'')
           .Cells(numberRow,11).Value=IIF(mmain#0,mmain,'')
           .Cells(numberRow,12).Value=IIF(mmain2#0,mmain2,'')
           .Cells(numberRow,13).Value=IIF(msupl#0,msupl,'')
           .Cells(numberRow,14).Value=IIF(mround#0,mround,'')           
           
           numberRow=numberRow+1
           DO fillpercent WITH 'fSupl'
                   
      ENDSCAN                                 
      .Range(.Cells(1,1),.Cells(numberRow-1,14)).Select
      WITH objExcel.Selection
           .Borders(xlEdgeLeft).Weight=xlThin
           .Borders(xlEdgeTop).Weight=xlThin            
           .Borders(xlEdgeBottom).Weight=xlThin
           .Borders(xlEdgeRight).Weight=xlThin
           .Borders(xlInsideVertical).Weight=xlThin
           .Borders(xlInsideHorizontal).Weight=xlThin
           .VerticalAlignment=1
           .Font.Name='Times New Roman' 
           .Font.Size=9      
           .WrapText=.T.  
      ENDWITH 
      .Cells(1,1).Select                       
ENDWITH     
=SYS(2002)
=INKEY(2)
DO endPrnToExcel WITH 'fSupl'
objExcel.Visible=.T.
*************************************************************************************
*    ������ ������������� ������
*************************************************************************************
PROCEDURE prnTableCompare

pathcompare=pathmain+'\'+ALLTRIM(datset.pathcomp)+'\'
*pathpeopold=pathcompare+'people.dbf'
pathjobold=pathcompare+'datjob.dbf'
*pathfondold=pathcompare+'comfond.dbf'
IF !FILE(pathjobold)
   RETURN
ELSE 
   USE &pathjobold ALIAS olddatjob IN 0 ORDER 7 &&nid
ENDIF 

IF !USED('dcompare')
   USE dcompare ORDER 1 IN 0
ENDIF
IF !USED('comfond')
   USE comfond ORDER 1 IN 0
ENDIF

IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF

SELECT * FROM comprn INTO CURSOR curcomprn READWRITE ORDER BY num
ALTER TABLE curcomprn ADD COLUMN kodpeop N(5)
ALTER TABLE curcomprn ADD COLUMN nidjob N(5)
ALTER TABLE curcomprn ADD COLUMN kp N(3)
ALTER TABLE curcomprn ADD COLUMN kd N(3)
ALTER TABLE curcomprn ADD COLUMN kse N(4,2)
ALTER TABLE curcomprn ADD COLUMN fio C(60)
ALTER TABLE curcomprn ADD COLUMN tr N(1)
ALTER TABLE curcomprn ADD COLUMN kat N(1)

=AFIELDS(arFond,'curcomprn')
CREATE CURSOR curprn FROM ARRAY arFond 

*SELECT * FROM datjob INTO CURSOR curprn READWRITE
*SELECT curprn
*REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL

*DO fltstructure WITH 'mtokl#0','dcompare'
SELECT datjob
SET FILTER TO 
SELECT dcompare
REPLACE tr WITH IIF(SEEK(nidjob,'datjob',7),datjob.tr,0) ALL
SCAN ALL
     SELECT olddatjob
     SEEK dcompare.nidold 
     
     SELECT curcomprn
     REPLACE say1 WITH 0,say2 WITH 0,say3 WITH 0,say4 WITH 0,tr WITH dcompare.tr,kp WITH dcompare.kp kd WITH dcompare.kd,kse WITH dcompare.kse,kodpeop WITH dcompare.kodpeop,nidjob WITH dcompare.nidjob ALL 
     msfold=0
     SCAN ALL          
          DO CASE     
             CASE ALLTRIM(LOWER(fname))='mslwork'        
                  REPLACE say1 WITH olddatjob.pslwork,say2 WITH olddatjob.mslwork
                  REPLACE say3 WITH dcompare.newpsl,say4 WITH dcompare.newmsl
                  msfold=msfold+olddatjob.mslwork
             CASE num=20 && ������� ������             
                  REPLACE say2 WITH datset.bstold
                  REPLACE say4 WITH dcompare.bst
             CASE num=25 && �������� ������
                  REPLACE say1 WITH IIF(olddatjob.kf<19,olddatjob.kf,0)
                  REPLACE say3 WITH IIF(dcompare.kf<19,dcompare.kf,0)
             CASE num=26 && �������� ������
                  REPLACE say1 WITH olddatjob.namekf
                  REPLACE say3 WITH dcompare.namekf     
             CASE num=27 && ���������� �����������               
                  REPLACE say1 WITH olddatjob.pkf
                  REPLACE say3 WITH dcompare.pkf            
             CASE num=85                 
                  REPLACE say4 WITH dcompare.difsum
             CASE num=70                  
                  REPLACE say1 WITH olddatjob.pprem,say2 WITH olddatjob.mprem        
                  msfold=msfold+olddatjob.mprem
             CASE num=71                                                     
                  REPLACE say3 WITH dcompare.pprem,say4 WITH dcompare.mprem    
             CASE num=72   
                  REPLACE say3 WITH IIF(dcompare.mtokl#0,ROUND(dcompare.mbdpl/dcompare.mtokl*100,0),0),say4 WITH dcompare.mbdpl       
             CASE num=80 
                  REPLACE say2 WITH msfold
                  REPLACE say4 WITH dcompare.msf+dcompare.newmsl               
             OTHERWISE 
                  ms1=IIF(!EMPTY(fpers),'olddatjob.'+ALLTRIM(fpers),'')
                  ms2=IIF(!EMPTY(fname),'olddatjob.'+ALLTRIM(fname),'')
                  ms11=IIF(!EMPTY(fpers),'dcompare.'+ALLTRIM(fpers),'')
                  ms12=IIF(!EMPTY(fname),'dcompare.'+ALLTRIM(fname),'')                 
                  msfold=msfold+EVALUATE(ms2)
                  REPLACE say1 WITH IIF(!EMPTY(ms1),EVALUATE(ms1),0),say3 WITH IIF(!EMPTY(ms11),EVALUATE(ms11),0)
                  REPLACE say2 WITH IIF(!EMPTY(ms2),EVALUATE(ms2),0),say4 WITH IIF(!EMPTY(ms12),EVALUATE(ms12),0)
                
          ENDCASE
          SELECT curcomprn
         
     ENDSCAN
     SELECT curprn
     APPEND FROM DBF('curcomprn')
     SELECT dcompare     
ENDSCAN
SELECT curprn
DELETE FOR num<70.AND.say1=0.AND.say2=0.AND.say3=0.AND.say4=0
INDEX ON STR(nidjob,5)+STR(num,2) TAG T1
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') FOR num=20
REPLACE kat WITH IIF(SEEK(nidjob,'datjob',7),datjob.kat,0) ALL
DO fltstructure WITH 'kse#0','curprn'

SET ORDER TO 1
GO TOP
DO procForPrintAndPreview WITH 'reptbltot','������������� �������'
SELECT curprn
USE
SELECT curcomprn
USE
SELECT olddatjob
USE
SELECT dcompare
USE
SELECT comfond
USE

SELECT people