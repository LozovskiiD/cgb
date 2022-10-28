fSupl=CREATEOBJECT('FORMSUPL')
IF USED('curpru2')
   SELECT curpu2
   USE 
ENDIF 
IF !USED('boss')
   USE boss IN 0
ENDIF 
dBeg=CTOD('  .  .    ')
dEnd=CTOD('  .  .    ')
kvPu=0
yearPu=YEAR(DATE())
WITH fSupl
     .Caption='создание файла для импорта в ПУ-2'
     DO addshape WITH 'fSupl',1,20,20,150,400,8   
      
     DO adTboxAsCont WITH 'fSupl','txtBeg',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('wпуть для импортаw'),dHeight,'период с',0,1
     DO adTboxNew WITH 'fSupl','boxBeg',.txtBeg.Top,.txtBeg.Left+.txtBeg.Width-1,200,dHeight,'dBeg',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'fSupl','txtEnd',.txtBeg.Left,.txtBeg.Top+.txtBeg.Height-1,.txtBeg.Width,dHeight,'период по',0,1
     DO adTboxNew WITH 'fSupl','boxEnd',.txtEnd.Top,.boxBeg.Left,.boxBeg.Width,dHeight,'dEnd',.F.,.T.,0,.F.
     
     DO adTboxAsCont WITH 'fSupl','txtKv',.txtBeg.Left,.txtEnd.Top+.txtEnd.Height-1,.txtBeg.Width,dHeight,'квартал',0,1
     DO adTboxNew WITH 'fSupl','boxKv',.txtKv.Top,.boxBeg.Left,.boxBeg.Width,dHeight,'kvPu',.F.,.T.,0,.F.
     .boxKv.Alignment=0
     .boxKv.InputMask='9'
     
     DO adTboxAsCont WITH 'fSupl','txtYear',.txtBeg.Left,.txtKv.Top+.txtKv.Height-1,.txtBeg.Width,dHeight,'год',0,1
     DO adTboxNew WITH 'fSupl','boxYear',.txtYear.Top,.boxBeg.Left,.boxBeg.Width,dHeight,'yearPu',.F.,.T.,0,.F.
     .boxYear.Alignment=0
     .boxYear.InputMask='9999'
     
     DO adTboxAsCont WITH 'fSupl','txtFile',.txtBeg.Left,.txtYear.Top+.txtYear.Height-1,.txtBeg.Width,dHeight,'имя файла',0,1
     DO adTboxNew WITH 'fSupl','boxFile',.txtFile.Top,.boxBeg.Left,.boxBeg.Width,dHeight,'datset.filepu2',.F.,.T.,0,.F.
     
     hs4=' путь для файла'
     DO addContFormNew WITH 'fSupl','txtPath',.txtBeg.Left,.txtFile.Top+.txtFile.Height-1,.txtBeg.Width,dHeight,hs4,0,.F.,'DO readPathPu2',.F.,.F. 
     DO adTboxNew WITH 'fSupl','boxPath',.txtPath.Top,.boxBeg.Left,.boxBeg.Width,dHeight,'datset.pathpu2',.F.,.F.,0,.F.
     .Shape1.Height=.txtBeg.Height*6+40
     .Shape1.Width=.txtBeg.Width+.boxBeg.Width+40
        
     
     DO addcontlabel WITH 'fSupl','butFile',.Shape1.Left+(.Shape1.Width-(RetTxtWidth('wвозвратw')*3)-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wвозвратw'),dHeight+5,'файл','DO createfilepu2'
     DO addcontlabel WITH 'fSupl','butPrn',.butFile.Left+.butFile.Width+10,.butFile.Top,.butFile.Width,dHeight+5,'печать','DO prnpu2'             
     DO addcontlabel WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+10,.butFile.Top,.butFile.Width,dHeight+5,'возврат','fSupl.Release'             
        
     DO addShape WITH 'fSupl',11,.Shape1.Left,.butFile.Top,.butFile.Height,.Shape1.Width,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+3,.Shape1.Left,.Shape1.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.     
     
     DO addcontlabel WITH 'fSupl','butOk',.butPrn.Left,.butPrn.Top,.butPrn.Width,dHeight+5,'возврат','fSupl.Release'             
     .butOk.Visible=.F.
        
        
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.butFile.Height+60
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
********************************************************************
PROCEDURE readPathPu2
newpathfile=GETDIR('','','Укажите папку для сохранения',64)
newpathfile=IIF(!EMPTY(newpathfile),newpathfile,datset.pathpu2)
REPLACE datset.pathpu2 WITH newpathfile
fSupl.Refresh
********************************************************************
PROCEDURE createfilepu2
IF (EMPTY(dBeg).OR.EMPTY(dEnd)).OR.dEnd<dBeg
   RETURN
ENDIF
IF EMPTY(datset.filepu2).OR.EMPTY(datset.pathpu2)
   RETURN
ENDIF
IF kvpu=0.OR.yearpu=0
   RETURN
ENDIF
IF USED('curpu2')
   SELECT curpu2
   USE
ENDIF 
SELECT * FROM people WHERE BETWEEN(date_in,dBeg,date_in) INTO CURSOR curpu2 READWRITE
SELECT curpu2
SCAN ALL
     DO unosimbpu2 WITH 'curPu2.pnum',.T.,.T.
ENDSCAN
INDEX ON pnum TAG T1
SELECT peopout
SET FILTER TO BETWEEN(date_out,dBeg,dEnd)
SCAN ALL
    DO unosimbpu2 WITH 'peopout.pnum',.T.,.T.
     SCATTER TO dimout
     SELECT curpu2
     SEEK peopout.pnum
     IF !FOUND()
        APPEND BLANK 
        GATHER FROM dimout        
     ELSE
        REPLACE date_out WITH peopout.date_out
     ENDIF 
     SELECT peopout
     
ENDSCAN 
SET FILTER TO 
SELECT curpu2
ALTER TABLE curpu2 ADD COLUMN dprin C(10)
ALTER TABLE curpu2 ADD COLUMN nprin C(10)
ALTER TABLE curpu2 ADD COLUMN dprout C(10)
ALTER TABLE curpu2 ADD COLUMN nprout C(10)

COUNT TO maxPu2
IF maxPu2=0
   SELECT people
   RETURN
ENDIF
CREATE CURSOR memopu2 (txtpu M)
SELECT memopu2
APPEND BLANK     
pu2file=FCREATE(ALLTRIM(datset.pathpu2)+ALLTRIM(datset.filepu2)+'.txt')

WITH fSupl
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH
SET DATE TO BRITISH 
SELECT curPu2
SELECT memopu2
REPLACE txtpu WITH '{"ver":"2.0","forma":"PU2","unpf":"'+ALLTRIM(boss.fszn)+'","unp":"'+ALLTRIM(boss.unp)+'","plat":"'+ALLTRIM(boss.office)+'","type":1,"kv":'+LTRIM(STR(kvpu))+',"year":"'+LTRIM(STR(yearpu))+'","pck":"PU2","tel":null,"ddoc":null,"data":['
SELECT curpu2
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec      
GO TOP
npu=0
SCAN ALL
     npu=npu+1
     fpu=UPPER(LEFT(ALLTRIM(fio),AT(' ',ALLTRIM(fio))))     
     nmpu=UPPER(LEFT(ALLTRIM(SUBSTR(ALLTRIM(fio),AT(' ',ALLTRIM(fio)))),1))
     opu=UPPER(LEFT(ALLTRIM(SUBSTR(ALLTRIM(fio),RAT(' ',ALLTRIM(fio)))),1))
     din=IIF(BETWEEN(date_in,dBeg,dEnd),DTOC(date_in),'  /  /    ')
     dout=IIF(BETWEEN(date_out,dBeg,dEnd),DTOC(date_out),'')
     psv=IIF(lvn.AND.BETWEEN(date_in,dBeg,dEnd),'"1"','"0"')   
     STORE '' TO nordinpu2,dordinpu2,nordoutpu2,nordoutpu2
     IF EMPTY(date_out)
        ordpu2=IIF(SEEK(curpu2.nid,'peoporder',2),peoporder.kord,0)
        nordinpu2='" "'
        dordinpu2='"  /  /    "'
        IF ordpu2#0.AND.SEEK(ordpu2,'datorder',1)
           nordinpu2='"'+ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)+'"'
           dordinpu2='"'+DTOC(datorder.dateord)+'"'          
        
        ENDIF      
   *     strpu='{"ils":"'+ALLTRIM(pnum)+'","fzl":"'+fpu+'","izl":"'+nmpu+'","ozl":"'+opu+'","r1":[{"dfr1":"'+din+'","dpr11":'+dordinpu2+',"npr11":'+nordinpu2+',"sovm":'+psv+'}],"r2":[],"r3":[]}'+IIF(npu=max_rec,'',',') 
        strpu='{"ils":"'+ALLTRIM(pnum)+'","fzl":"'+fpu+'","izl":"'+nmpu+'","ozl":"'+opu+'","r1":[{"dfr1":"'+din+'","dpr11":'+dordinpu2+',"npr11":'+nordinpu2+',"sovm":'+psv+'}],"r2":[]}'+IIF(npu=max_rec,'',',') 
     ELSE
        IF SEEK(curpu2.nid,'peoporder',2)
           SELECT peoporder
           ordpu2=kord
           SCAN WHILE nidpeop=curpu2.nid
                ordpu2=kord
           ENDSCAN
        ENDIF 
        SELECT curpu2
        nordoutpu2='" "'
        dordoutpu2='"  /  /    "'
        IF ordpu2#0.AND.SEEK(ordpu2,'datorder',1)
           nordoutpu2='"'+ALLTRIM(STR(datorder.numord))+'-'+ALLTRIM(datorder.strord)+'"'
           dordoutpu2='"'+DTOC(datorder.dateord)+'"'          
        
        ENDIF      
      *  strpu='{"ils":"'+ALLTRIM(pnum)+'","fzl":"'+fpu+'","izl":"'+nmpu+'","ozl":"'+opu+'","r1":[{"dto1":"'+dout+'"}],"r2":[],"r3":[]}'+IIF(npu=max_rec,'',',') 
      *  strpu='{"ils":"'+ALLTRIM(pnum)+'","fzl":"'+fpu+'","izl":"'+nmpu+'","ozl":"'+opu+'","r1":[{"dto1":"'+dout+'","dpr12":'+dordoutpu2+',"npr12":'+nordoutpu2+'}],"r2":[],"r3":[]}'+IIF(npu=max_rec,'',',') 
         strpu='{"ils":"'+ALLTRIM(pnum)+'","fzl":"'+fpu+'","izl":"'+nmpu+'","ozl":"'+opu+'","r1":[{"dto1":"'+dout+'","dpr12":'+dordoutpu2+',"npr12":'+nordoutpu2+'}],"r2":[]}'+IIF(npu=max_rec,'',',') 
     ENDIF     
 
     SELECT memopu2
     REPLACE txtpu WITH txtpu+strpu
     SELECT curpu2  
     one_pers=one_pers+1
     pers_ch=one_pers/max_rec*100
     fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
     fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch  
ENDSCAN 
SELECT memopu2
replace txtpu WITH txtpu+']}'
FPUTS(pu2file,memopu2.txtpu)
SET DATE TO GERMAN   
=FCLOSE(pu2file)
SELECT people
WITH fSupl
     .butOk.Visible=.T.     
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.      
     .lab25.Visible=.F.      
ENDWITH    
********************************************************************
*     Процедура перевода символьной строки в один формат
********************************************************************
PROCEDURE unosimbpu2
PARAMETERS parsimb,parrep,parUp
parsimbnew=LOWER(&parsimb)
parsimbnew=CHRTRAN(parsimbnew,"снрвак","chpbak")
str_ch=UPPER(LEFT(parsimbnew,1))+LOWER(SUBSTR(parsimbnew,2))
str_cx=str_ch
len_ch=LEN(ALLTRIM(parsimbnew))
str_ch=''         
FOR i=1 TO len_ch
    new_simb=LEFT(str_cx,1)
    IF INLIST(new_simb,' ','.','-')
       str_ch=str_ch+new_simb  
       str_cx=SUBSTR(str_cx,2)         
       new_simb=UPPER(LEFT(str_cx,1))
    ENDIF
    str_ch=str_ch+new_simb
    str_cx=SUBSTR(str_cx,2)
ENDFOR    
IF parUp
   str_ch=UPPER(str_ch)
ENDIF
IF parrep
   REPLACE &parsimb WITH str_ch
ELSE 
   &parsimb=str_ch           
ENDIF  
********************************************************************
PROCEDURE prnpu2