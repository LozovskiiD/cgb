fSupl=CREATEOBJECT('FORMSUPL')
dBeg=CTOD('  .  .    ')
dEnd=CTOD('  .  .    ')
lCheck=.F.
lDoc=.F.
WITH fSupl
     .Caption='Перенос приказов в архив'    
     DO addshape WITH 'fSupl',1,10,10,150,500,8    
     
     DO adLabMy WITH 'fSupl',1,'период с ',.Shape1.Top+20,.Shape1.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxBeg',.Shape1.Top+20,.Shape1.Left,RetTxtWidth('99/99/99999'),dHeight,'dBeg',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3
     
     DO adLabMy WITH 'fSupl',2,' по ',.lab1.Top,.Shape1.Left,.Shape1.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxEnd',.boxBeg.Top,.Shape1.Left,.boxBeg.Width,dHeight,'dEnd',.F.,.T.,0
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxBeg.Width-.lab2.Width-.boxEnd.Width-30)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+10
     .lab2.Left=.boxBeg.Left+.boxBeg.Width+10
     .boxEnd.Left=.lab2.Left+.lab2.Width+10
     
     DO adCheckBox WITH 'fSupl','checkDoc','удалять файлы "doc"',.boxBeg.Top+.boxbeg.Height+10,.Shape1.Left,150,dHeight,'lDoc',0
     .checkDoc.Left=.Shape1.Left+(.Shape1.Width-.checkDoc.Width)/2    
   
     .Shape1.Height=.boxBeg.Height+.checkDoc.Height+50
      DO adCheckBox WITH 'fSupl','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'lCheck',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2    
              
     DO addButtonOne WITH 'fSupl','butDo',.Shape1.Left+(.Shape1.Width-RetTxtWidth('выполнениеw')*2-15)/2,.check1.Top+.check1.Height+20,'выполнение','','DO moveordertoarchive',39,RetTxtWidth('wвыполнениеw'),'выполнение' 
     DO addButtonOne WITH 'fSupl','butRet',.butDo.Left+.butDo.Width+15,.butDo.Top,'возврат','','fsupl.Release',39,.butDo.Width,'возврат' 
     
     DO addShape WITH 'fSupl',11,.Shape1.Left,.butDo.Top,.butDo.Height,.Shape1.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape1.Left,.Shape1.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.     
             
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.check1.Height+.butDo.Height+80
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************
PROCEDURE moveordertoarchive
IF !lCheck.OR.(EMPTY(dBeg).OR.EMPTY(dEnd)).OR.(dEnd<dBeg)
   RETURN 
ENDIF
WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
ENDWITH 
STORE 0 TO max_rec,one_pers,pers_ch
IF lDoc
   SELECT datorder
   COUNT TO max_rec
   SCAN ALL
        IF BETWEEN(dateOrd,dBeg,dEnd).AND.!EMPTY(pathor)
           filed=ALLTRIM(pathor)
           DELETE FILE &filed
        ENDIF
        one_pers=one_pers+1
        pers_ch=one_pers/max_rec*100
        fSupl.lab25.Caption='удаление файлов приказов - '+LTRIM(STR(pers_ch))+'%'       
        fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
   ENDSCAN 
ENDIF
fSupl.Shape12.Width=0
fSupl.lab25.Caption='Перенос в архив записей о приказах'
SELECT orderarc
APPEND FROM datorder FOR BETWEEN(dateOrd,dBeg,dEnd)
SELECT pordarc
APPEND FROM peoporder FOR BETWEEN(dOrd,dBeg,dEnd)
SELECT datorder
SET ORDER TO 1
GO BOTTOM
newkod=kod+1
DELETE FOR BETWEEN(dateOrd,dBeg,dEnd)
COUNT TO totrec
IF totrec=0
   APPEND BLANK
   REPLACE kod WITH newKod
ENDIF
SELECT peoporder
SET ORDER TO 1
GO BOTTOM 
newNid=nid+1
DELETE FOR BETWEEN(dOrd,dBeg,dEnd)
COUNT TO totrec
IF totrec=0
   APPEND BLANK
   REPLACE nid WITH newNid
ENDIF
WITH fSupl
     .SetAll('Visible',.T.,'myCommandButton')     
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
ENDWITH