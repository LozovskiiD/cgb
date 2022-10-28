SET CENTURY ON
USE sprpodr ORDER 1 IN 0
USE rasp ORDER 1 IN 0
USE sprdolj ORDER 1 IN 0
USE sprkat ORDER 1 IN 0
USE sprkval ORDER 1 IN 0
USE people  ORDER 1 IN 0
USE sprschool IN 0 ORDER 2
USE datspr IN 0
USE datmenu IN 0
USE datorder ORDER 1 IN 0
USE peoporder ORDER 1 IN 0
USE sprdog ORDER 1 IN 0
USE datotp ORDER 3 IN 0
USE datjob ORDER 4 IN 0 &&str(kodpeop,4)+STR(TR,1)
USE sprtype ORDER 1 IN 0
USE sprorder IN 0
USE datset IN 0
USE sprspec ORDER 2 IN 0
USE peopout ORDER 1 IN 0
USE datjobout ORDER 1 IN 0 
USE orderarc ORDER 1 IN 0
USE pordarc ORDER 1 IN 0

USE sprtot IN 0
SELECT kod,name,otm,fl FROM sprtot WHERE sprtot.kspr=1 INTO CURSOR curDocum READWRITE   && курсор для документа удостоверяющего личность
SELECT curDocum
INDEX ON kod TAG T1
SELECT kod,name,otm,fl FROM sprtot WHERE sprtot.kspr=2 INTO CURSOR curFamily READWRITE  && курсор для видов семейного положения 
SELECT curFamily
INDEX ON kod TAG T1
SELECT kod,name,otm,fl FROM sprtot WHERE sprtot.kspr=3 INTO CURSOR curEducation READWRITE && курсоср для видов образования
SELECT curEducation
INDEX ON kod TAG T1
SELECT kod,name,otm,fl FROM sprtot WHERE sprtot.kspr=4 INTO CURSOR curSprOtp READWRITE && курсоср для видов отпусков
SELECT curSprOtp
INDEX ON kod TAG T1
SELECT kod,name,otm,fl FROM sprtot WHERE sprtot.kspr=5 INTO CURSOR curPrichOtp READWRITE && курсоср для причин ухода в отпуск
SELECT curPrichOtp
INDEX ON kod TAG T1
SELECT sprtot
USE

CREATE CURSOR curSex (kod N(1),name C(20),otm c(3),fl L)
SELECT curSex
APPEND BLANK
REPLACE kod WITH 1,name WITH 'мужской'
APPEND BLANK
REPLACE kod WITH 2,name WITH 'женский'
SELECT curSex
INDEX ON kod TAG T1

SELECT * from cardFond INTO CURSOR suplCardFond READWRITE 
SELECT * FROM datJob WHERE datjob.kodpeop=people.num INTO CURSOR curJobSupl ORDER BY tr READWRITE 

SELECT * FROM sprdolj INTO CURSOR curSprDolj READWRITE ORDER BY name
SELECT * FROM sprpodr INTO CURSOR curSprPodr READWRITE ORDER BY namework
SELECT * FROM sprtype INTO CURSOR curSprType READWRITE
SELECT curSprType
INDEX ON kod TAG T1
SELECT * FROM sprkval INTO CURSOR curSprKval READWRITE ORDER BY kod
SELECT cursprkval
APPEND BLANK
REPLACE name WITH 'без категории'

agemen=63
agewom=58

SELECT people
SCAN ALL 
     IF !EMPTY(age)
        kontAge=YEAR(DATE())-YEAR(age)
        DO CASE 
           CASE (sex=1.AND.kontage>INT(agemen)).OR.(sex=2.AND.kontage>INT(agewom)) 
                 REPLACE pens WITH .T.
           CASE (sex=1.AND.kontage=INT(agemen)).OR.(sex=2.AND.kontage=INT(agewom))                
                 IF MONTH(DATE())>MONTH(age).OR.(MONTH(DATE())=MONTH(age).AND.DAY(DATE())>=DAY(age))
                    REPLACE pens WITH .T.
                 ELSE     
                     REPLACE pens WITH .F.
                 ENDIF                 
           OTHERWISE 
                REPLACE pens WITH .F.
        ENDCASE
     ELSE    
     ENDIF    
ENDSCAN 
GO TOP
nameSay=''
stajOrg=''
frmTop=CREATEOBJECT('Formtop')
WITH frmTop   
     .procExit='DO exitfrompcard' 
     width_obj=.Width/3*2  && ширина части для личной карточки
     .Caption='Кадровый учёт'     
     SELECT datmenu
     GO TOP   
     DO addButtonOne WITH 'frmTop','menuContTop',10,3,'главное меню','topbottom.ico','DO topMenu',39,RetTxtWidth('главное меню')+44,'главное меню'
     m_ch=1    
     leftCont=.menuContTop.Left+.menuContTop.Width+10
     topCont=3
     DO WHILE !EOF()
        namecont='menucont'+LTRIM(STR(m_ch))     
        DO addcontico WITH 'frmTop',namecont,leftCont,topCont,datmenu.mIco,datmenu.mproc,ALLTRIM(datmenu.nmenu),39,39         
        SKIP
        m_ch=m_ch+1
        leftCont=frmTop.&nameCont..Left+frmTop.&nameCont..Width+5
     ENDDO
     topObj=.menucont1.Top+.menucont1.Height+5  
     DO adTboxAsCont WITH 'frmTop','txtName',0,topObj,width_obj,dHeight,nameSay,2,1 
     .txtName.FontBold=.T.
     topObj=topObj+.txtName.Height+5
     .AddObject('grdPers','gridMynew')     
     WITH .grdPers
          .Top=.Parent.txtName.Top     
          .Left=.Parent.txtName.Width+5
          .Width=.Parent.Width-.Parent.txtName.Width-5
          .Height=.Parent.Height-.Top
          .RecordSourceType=1
          .scrollBars=2   
          .ColumnCount=0         
          .colNesInf=2           
              
          SELECT people
          DO addColumnToGrid WITH 'frmTop.grdPers',3
          .Column1.RemoveObject('Header1')
          .Columns(1).AddObject('Header1','HeaderMy')
          .Column2.RemoveObject('Header1')
          .Columns(2).AddObject('Header1','HeaderMy')
          .Column1.Header1.procForClick='DO clickHeaderPers WITH 1'
          .Column2.Header1.procForClick='DO clickHeaderPers WITH 2'
          .RecordSource='people'                                         
          .Column1.ControlSource='people.num'
          .Column2.ControlSource='people.fio'                  
          .Column1.Width=RetTxtWidth('99999')         
          .Columns(.ColumnCount).Width=0    
          .Column2.Width=.Width-.column1.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Column1.Header1.Caption=IIF(SYS(21)='1','*№','№')
          .Column2.Header1.Caption=IIF(SYS(21)='2','*фио','фио')       
          .Column1.Alignment=1  
          .Column2.Alignment=0  
          .Column1.Format='Z'
          .procAfterRowColChange='DO changeRowGrdPers'                                                              
          .SetAll('Enabled',.F.,'ColumnMy') 
          .Columns(.ColumnCount).Enabled=.T.          
     ENDWITH 
     DO gridSizeNew WITH 'frmTop','grdPers','shapeingrid',.T.,.F.                      
     imageWidth=dHeight*7
     imageHeight=dHeight*7-6    
     topObj=.txtName.Top+.txtName.Height-1
    .AddObject('formImage','myImage')
     WITH .formImage     
          .BorderStyle=1
          .BackStyle=1
          .Stretch=1                   
          .Left=.Parent.txtName.Width-imageWidth               
          .Top=topObj
          .Width=imageWidth
          .Height=imageHeight
          .Visible=.F.
          .ToolTipText='Изменить - правая кнопка мыши'                  
          .procForRightClick='DO rightClickImage'
          .Visible=.T.     
     ENDWITH           
     SELECT suplCardFond
     SET FILTER TO nBlock=3.AND.logTot                                             
     GO TOP    
     SCAN ALL
          IF !EMPTY(exprsay)
             repSayCard=ALLTRIM(exprsay)
             REPLACE sayCard WITH &repSayCard
          ENDIF    
     ENDSCAN
     GO TOP
     labNum=0
     kvoStr=0
     kvoLab=0
     widthLab=RetTxtWidth('Wпоследнее место работыW')
     widthTxtShort=(.txtName.Width-imageWidth-widthLab*2)/2
     widthTxtFull=(.txtName.Width-widthLab*2+4)/2
     widthTxtLong=.txtName.Width-widthLab*2-widthTxtShort+5
     widthTxt=(.txtName.Width-imageWidth-widthLab*2)/2
     labLeft=.txtName.Left   
     txtLeft=labLeft+widthLab-1
     DO WHILE !EOF()                 
        IF !EMPTY(suplCardFond.namerec)                           
           sayField=suplCardFond.namerec               
           labNum=labNum+1
           namecont='lab'+LTRIM(STR(labnum))
           DO adTBoxAsCont WITH 'frmTop',namecont,labLeft,topObj,widthLab,dHeight,suplCardFond.namerec,1,1
              .&nameCont..Visible=.F.                    
              proc_obj=procObj
              IF !EMPTY(proc_Obj)
                 &proc_Obj
                 nObj=suplCardFond.nameObj 
                 frmTop.&nObj..Visible=.F.     
              ENDIF  
              topObj=topObj+dHeight-1                  
              kvoLab=kvoLab+1
              IF kvoLab=15                 
                 widthTxt=(.txtName.Width-widthLab-widthLab*2)/2
                 labLeft=.txtName.Left   
                 txtLeft=labLeft+widthLab-1               
              ENDIF
        ENDIF    
        SELECT suplCardFond
       SKIP
     ENDDO
     DO adTBoxAsCont WITH 'frmTop','labSupl',labLeft,topObj,widthLab,dHeight,'',1,1               
     DO adTBoxAsCont WITH 'frmTop','txtSupl',labLeft,topObj,widthtxt,dHeight,'',1,0
     .labSupl.Visible=.F.
     .txtSupl.Visible=.F.  
          
     DO addContFormNew WITH 'fRmTop','txtKont',.txtName.Left,topObj,.txtName.Width,dHeight,'условия работы',1,.F.,.F. 
     .txtKont.Visible=.T.
     SELECT suplCardFond
     SET FILTER TO nBlock=4.AND.logTot                                             
     GO TOP    
     SCAN ALL
          IF !EMPTY(exprsay)
             repSayCard=ALLTRIM(exprsay)
             REPLACE sayCard WITH &repSayCard
          ENDIF    
     ENDSCAN
     GO TOP
     kvoStr=0
     kvoLab=0
     widthLab=RetTxtWidth('Wпоследнее место работыW')   
     widthTxt=(.txtKont.Width-widthLab*2)/2
     labLeft=.txtKont.Left   
     txtLeft=labLeft+widthLab-1
     DO WHILE !EOF()                 
        IF !EMPTY(suplCardFond.namerec)                           
           sayField=suplCardFond.namerec               
           labNum=labNum+1
           namecont='lab'+LTRIM(STR(labnum))
           DO adTBoxAsCont WITH 'frmTop',namecont,labLeft,topObj,widthLab,dHeight,suplCardFond.namerec,1,1
              .&nameCont..Visible=.F.                    
              proc_obj=procObj
              IF !EMPTY(proc_Obj)
                 &proc_Obj
                 nObj=suplCardFond.nameObj 
                 frmTop.&nObj..Visible=.F.     
              ENDIF  
              topObj=topObj+dHeight-1                  
              kvoLab=kvoLab+1
              IF kvoLab=15                 
                 widthTxt=(.txtKont.Width-widthLab*2)/2
                 labLeft=.txtName.Left   
                 txtLeft=labLeft+widthLab-1               
              ENDIF
        ENDIF    
        SELECT suplCardFond
       SKIP
     ENDDO
     DO adTBoxAsCont WITH 'frmTop','labSupl1',labLeft,topObj,widthLab,dHeight,'',1,1               
     DO adTBoxAsCont WITH 'frmTop','txtSupl1',labLeft,topObj,widthtxt,dHeight,'',1,0
     .labSupl1.Visible=.F.
     .txtSupl1.Visible=.F.  
     .AddObject('grdJob','gridMynew')  
     .grdJob.ColumnCount=6     
     .grdJob.Visible=.F.
     DO gridSizeNew WITH 'frmTop','grdJob','shapeingrid1',.F.,.T. 
     topObj=.grdJob.Top+.grdJob.Height-1            
     DO addContFormNew WITH 'fRmTop','txtOtp',.txtName.Left,topObj,.txtName.Width,dHeight,'отпуска',1,.F.,'DO formForOtp' 
     .AddObject('grdOtp','gridMynew')
     .grdOtp.ColumnCount=10  
     DO gridSizeNew WITH 'frmTop','grdOtp','shapeingrid2',.F.,.F.  
     .txtOtp.Visible=.T.
     .grdOtp.Visible=.F.
     .ShapeInGrid2.Visible=.F.
ENDWITH     
frmTop.Show
DO changeRowGrdPers
READ EVENTS
***********************************************************************************************************************
PROCEDURE exitFrompcard
CLEAR EVENTS
QUIT
************************************************************************************************************************
PROCEDURE topMenu
PARAMETERS parObj
IF !USED('datmenu')
   USE datmenu IN 0   
ENDIF
SELECT datmenu 
COUNT TO maxBar
DIMENSION dim_proc(maxBar)
GO TOP
CurTotHeight=35
row_pop=CurTotHeight/FONTMETRIC(1,dFontName,dFontSize)
col_pop=frmTop.menuContTop.LEFT/FONTMETRIC(6,dFontName,dFontSize)
DEFINE POPUP menuTop FROM row_pop,col_pop SHORTCUT MARGIN FONT dFontName,dFontSize  COLOR SCHEME 4
numbar=0
DO WHILE !EOF()   
   numbar=numbar+1
   DEFINE BAR numbar OF menuTop PROMPT ' '+datmenu.nmenu PICTURE ALLTRIM(datmenu.mico) FONT dFontName,dFontSize
   dim_proc(numbar)=mproc 
   SKIP
ENDDO
SELECT datmenu
GO TOP
numbar=0
DO WHILE !EOF()  
   numbar=numbar+1     
   ON SELECTION BAR numbar OF menuTop DO choiceFromMenuTop   
   SKIP
ENDDO  
ACTIVATE POPUP menuTop
**************************************************************************************************************************
PROCEDURE choiceFromMenuTop
IF !EMPTY(dim_proc(BAR()))
   &dim_proc(BAR())
ENDIF  
******************************************************************************************************
PROCEDURE changeRowGrdPers
logPict=.F.
SELECT people
DO actualStajToday WITH 'people','people.date_in','DATE()'
WITH frmTop
     peopleRec=RECNO()       
     nameSay=ALLTRIM(people.fio)+' ('+LTRIM(STR(people.num))+')'
     .txtName.controlSource='nameSay'
               
     IF !EMPTY(people.pFoto)
        pathmem=FULLPATH('kadry.fxp')
        pathmem='"'+LEFT(pathmem,LEN(pathmem)-9)+LTRIM(STR(people.num))+'.pic'+'"'     
        COPY MEMO people.pFoto TO &pathmem                            
        .formImage.Picture=pathmem
        .formImage.Visible=.T.      
        widthTxt=widthTxtShort  
        logPict=.T.      
     ELSE 
       * .formImage.Picture='nofhoto.jpg'
        .formImage.Visible=.F.
        widthTxt=widthTxtFull   
        logPict=.F.           
     ENDIF 
     IF logPict
          DELETE FILE &pathmem
     ENDIF
    * *--------------------------------------------прочая часть
     SELECT suplCardFond
     SET FILTER TO nBlock=3.AND.logTot                                             
     GO TOP 
     kvoSay=0 &&всего заполнено   
     SCAN ALL
          IF !EMPTY(exprsay)
             repSayCard=ALLTRIM(exprsay)
             REPLACE sayCard WITH &repSayCard            
             kvosay=IIF(!EMPTY(sayCard),kvosay+1,kvosay)
          ENDIF    
     ENDSCAN    
     GO TOP    
     sayCx=0
     labNum=0
     kvoStr=0  
     kvoCol=IIF(kvoSay#3,IIF(INT(kvosay/2)*2=kvoSay,INT(kvoSay/2),INT(kvoSay/2)+1),2)   &&всего в колонке  
     kvoCol_cx=0
     labLeft=.txtName.Left
     topObj=.txtName.Top+.txtName.Height-1
     txtLeft=labLeft+widthLab-1   
     .labSupl.Visible=.F.
     .txtSupl.Visible=.F.   
     DO WHILE !EOF()   
        labNum=labNum+1  
        namecont='lab'+LTRIM(STR(labnum))
        IF !EMPTY(nameObj)  
           nObj=suplCardFond.nameObj           
        ENDIF    
        IF !EMPTY(suplCardFond.sayCard)                   
           .&nameCont..Visible=.T.  
           .&nameCont..Top=topObj                   
           .&nameCont..Left=labLeft                   
           proc_obj=procObj           
           .&nObj..Visible=.T. 
           .&nObj..Top=topObj
           .&nObj..Left=txtLeft
           .&nObj..Width=widthTxt
           .&nObj..ControlSource='suplCardFond.sayCard'
           topObj=topObj+dHeight-1   
           kvoCol_cx=kvoCol_cx+1
           sayCx=sayCx+1
           DO CASE
              CASE kvosay=1
                   topObj=.txtName.Top+.txtName.Height-1  
                   labLeft=labLeft+widthLab+widthTxt-3
                   txtLeft=labLeft+widthLab-1                                
              CASE kvoCol_cx=kvoCol.AND.sayCx#kvoSay         
                   topObj=.txtName.Top+.txtName.Height-1   
                   labLeft=labLeft+widthLab+widthTxt-3
                   txtLeft=labLeft+widthLab-1               
           ENDCASE
           IF logPict.AND.kvoCol_cx>kvoCol.AND.topObj>=.formImage.Top+.formImage.Height-1
               widthTxt=widthTxtLong 
           ENDIF                                                                     
        ELSE   
           .&nameCont..Visible=.F. 
           .&nObj..Visible=.F.                
        ENDIF          
        SELECT suplCardFond
        SKIP
     ENDDO
     
     IF kvoSay#0.AND.MOD(kvoSay,2)=1   
        .labSupl.Visible=.T. 
        .labSupl.Top=topObj
        .labSupl.Left=labLeft
           
        .txtSupl.Visible=.T. 
        .txtSupl.Top=topObj
        .txtSupl.Left=txtLeft 
        .txtSupl.Width=widthTxt 
        topObj=topObj+dHeight-1                                    
     ENDIF  
      
     IF logPict.AND.topObj<.formImage.top+.formImage.Height-1
        topObj=.formImage.top+.formImage.Height-1
        
     ENDIF
     *------часть про контракт
     
     .txtKont.Top=topObj    
     widthTxt=ROUND((.txtKont.Width-widthLab*2)/2,0)
     widthTxt1=.txtKont.Width-widthlab*2-widthTxt+3
     labLeft=.txtKont.Left   
     txtLeft=labLeft+widthLab-1     
     SELECT suplCardFond
     SET FILTER TO nBlock=4.AND.logTot                                             
     GO TOP 
     kvoSay=0 &&всего заполнено   
     SCAN ALL
          IF !EMPTY(exprsay)
             repSayCard=ALLTRIM(exprsay)
             REPLACE sayCard WITH &repSayCard            
             kvosay=IIF(!EMPTY(sayCard),kvosay+1,kvosay)
          ENDIF    
     ENDSCAN    
     GO TOP     
     sayCx=0
     kvoStr=1 
     kvoCol_cx=0
     labLeft=.txtKont.Left
     topObj=.txtKont.Top+.txtKont.Height-1
     txtLeft=labLeft+widthLab-1  
     .labSupl1.Visible=.F.
     .txtSupl1.Visible=.F.   
     DO WHILE !EOF()   
        labNum=labNum+1  
        namecont='lab'+LTRIM(STR(labnum))
        IF !EMPTY(nameObj)  
           nObj=suplCardFond.nameObj                  
        ENDIF    
        IF !EMPTY(suplCardFond.sayCard)  
           .&nameCont..Visible=.T.  
           .&nameCont..Top=topObj   
            proc_obj=procObj      
           IF kvoStr#0.AND.MOD(kvoStr,2)=1
              labLeft=.txtKont.Left 
              txtLeft=labLeft+widthLab-1                 
           ELSE
              labLeft=.txtKont.Left+widthLab+widthTxt-2  
              txtLeft=.txtKont.Left+widthLab*2+widthTxt-3                             
           ENDIF                             
           .&nameCont..Left=labLeft                   
               
           .&nObj..Visible=.T.              
           .&nObj..Top=topObj
           .&nObj..ControlSource='suplCardFond.sayCard'
           .&nObj..Left=txtLeft
           .&nObj..Width=IIF(MOD(kvoStr,2)=1,widthTxt,widthTxt1)
            topObj=IIF(MOD(kvoStr,2)=1,topObj,topObj+dHeight-1)   
            kvoCol_cx=kvoCol_cx+1
            sayCx=sayCx+1   
            kvoStr=kvoStr+1                                  
        ELSE   
           .&nameCont..Visible=.F. 
           .&nObj..Visible=.F.                          
        ENDIF          
        SELECT suplCardFond
        SKIP       
     ENDDO
     
     IF kvosay#0.AND.MOD(kvoSay,2)=1
        .labSupl1.Visible=.T. 
        .labSupl1.Top=topObj
        .labSupl1.Left=.txtKont.Left+widthLab+widthTxt-2  
        .txtSupl1.Visible=.T. 
        .txtSupl1.Top=topObj
        .txtSupl1.Left=.txtKont.Left+widthLab*2+widthTxt-3  
        .txtSupl1.Width=widthTxt1 
        topObj=topObj+dHeight-1                            
     ENDIF      
       
     SELECT * FROM datJob WHERE datjob.kodpeop=people.num.AND.EMPTY(dateOut) INTO CURSOR curJobSupl ORDER BY tr READWRITE 
     SELECT curJobsupl
     COUNT TO maxJob
     GO TOP     
     WITH .grdJob
          .Visible=.T.
          .RecordSourceType=1
          .Height=.headerHeight+.RowHeight*(maxJob+1)                          
          .Width=.Parent.txtKont.Width
          .ScrollBars=2
          .ColumnCount=0
           DO addColumnToGrid WITH 'frmTop.grdJob',7       
          .RecordSource='curJobSupl'
          .Top=topObj
          .Column1.ControlSource="IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,'')"
          .Column2.ControlSource="IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,'')"
          .Column3.ControlSource='kse'
          .Column4.ControlSource="IIF(SEEK(tr,'sprtype',1),sprtype.name,'')"
          .Column5.ControlSource='lkv'
          .Column6.ControlSource='kdek'
          .Column3.Width=RetTxtWidth('999.999')  
          .Column4.Width=RetTxtWidth('внеш.совм.')  
          .Column5.Width=RetTxtWidth('катw')                             
          .Column6.Width=RetTxtWidth('999999')                             
          .Columns(.ColumnCount).Width=0
          .Column2.Width=(.Width-.Column3.width-.Column4.Width-.Column5.Width-.Column6.Width)/2
          .Column1.Width=.Width-.Column2.width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='подразделение'
          .Column2.Header1.Caption='должность'
          .Column3.Header1.Caption='объём'
          .Column4.Header1.Caption='тип'         
          .Column5.Header1.Caption='кат'  
          .Column6.Header1.Caption=' ! '        
                             
          .Column1.Alignment=0
          .Column2.Alignment=0   
          .Column4.Alignment=0  
          .Column5.Alignment=2 
          .Column6.Alignment=0 
          .Column6.Format='Z'
          .SetAll('Enabled',.F.,'Column')
          .procAfterRowColChange='DO showDekStav' 
          .Columns(.ColumnCount).Enabled=.T.  
          .Column5.AddObject('checkColumn5','checkContainer')
          .Column5.checkColumn5.AddObject('checkMy','checkBox')
          .Column5.CheckColumn5.checkMy.Visible=.T.
          .Column5.CheckColumn5.checkMy.Caption=''
          .Column5.CheckColumn5.checkMy.Left=6
          .Column5.CheckColumn5.checkMy.Top=3
          .Column5.CheckColumn5.checkMy.BackStyle=0
          .Column5.CheckColumn5.checkMy.ControlSource='curJobSupl.lkv'  
          .Column5.CheckColumn5.checkmy.Left=(.Column5.Width-SYSMETRIC(15))/2                                                                                         
          .column5.CurrentControl='checkColumn5'
         * .SetAll('Enabled',.F.,'ColumnMy')
         * .Column5.Enabled=.T. 
          .Column5.Sparse=.F. 
          
              
     ENDWITH
     DO gridSizeNew WITH 'frmTop','grdJob',.F.,.F.,.T. 
     .ShapeInGrid1.Height=.grdJob.Height
     .ShapeInGrid1.Top=.grdJob.Top
     .ShapeInGrid1.Left=.grdJob.Left+.grdJob.Width-SYSMETRIC(5)-3
     .grdJob.Columns(.grdJob.ColumnCount).SetFocus
     .grdPers.Columns(.grdPers.ColumnCount).SetFocus
     SELECT people  
     .txtOtp.Top=.grdJob.Top+.grdJob.Height-1   
     topObj=.txtOtp.Top+.txtOtp.Height-1    
     SELECT * FROM datOtp WHERE datOtp.nidpeop=people.nid INTO CURSOR curOtpSupl ORDER BY begOtp DESC READWRITE      
     
     WITH .grdOtp
          .ColumnCount=0
          DO addColumnToGrid WITH 'frmTop.grdOtp',8  
          .Top=topObj     
          .Height=.Parent.height-.Top
          .Width=.Parent.txtName.Width         
          .Left=.Parent.txtName.Left      
          .ScrollBars=2    
          .RecordSourceType=1
          .RecordSource='curOtpSupl'
          .Column1.ControlSource='curOtpSupl.nameotp'
          .Column2.ControlSource="IIF(!EMPTY(curOtpSupl.perbeg),curOtpSupl.perbeg,'')"
          .Column3.ControlSource="IIF(!EMPTY(curOtpSupl.perEnd),curOtpSupl.perend,'')"
          .Column4.ContRolSource='curOtpSupl.kvoDay'
          .Column5.ControlSource='curOtpSupl.begotp'
          .Column6.ControlSource='curOtpSupl.endotp'       
          .Column7.ControlSource='curOtpSupl.osnov'
                     
          .Column2.Width=RetTxtWidth('99999999999')  
          .Column3.Width=.Column2.Width
          .Column4.Width=RetTxtWidth('99999')
          .Column5.Width=.Column2.Width
          .Column6.Width=.Column2.Width        
          .Column7.Width=RetTxtWidth('пр.№51 - от 07/04/99999w')                  
         
          .Columns(.ColumnCount).Width=0
          .Column1.Width=.Width-.Column2.width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='вид отпуска'
          .Column2.Header1.Caption='период с'
          .Column3.Header1.Caption='период по'
          .Column4.Header1.Caption='дней'
          .Column5.Header1.Caption='начало'
          .Column6.Header1.Caption='окончание'   
          .Column7.Header1.Caption='приказ'   
          .Column4.Format='Z'
          .Column2.Alignment=2
          .Column3.Alignment=2
          .Column4.Alignment=1
          .Column5.Alignment=2
          .Column6.Alignment=2
          .Column7.Alignment=0
          .SetAll('BOUND',.F.,'ColumnMy')  
          .SetAll('Enabled',.F.,'ColumnMy') 
          .Columns(.ColumnCount).Enabled=.T.    
          .Visible=.T.   
     ENDWITH 
     .ShapeInGrid2.Visible=.T.
     .ShapeInGrid2.Top=.grdOtp.Top
     .ShapeInGrid2.Left=.grdOtp.Left+.grdOtp.Width-SYSMETRIC(5)-3
     .ShapeInGrid2.Height=.grdOtp.Height   
     topObj=.grdOtp.Top+.grdOtp.Height-1  
ENDWITH
*--------------------------------------Личная карточка-------------------------------------------------------------
PROCEDURE clickHeaderPers
PARAMETERS par1
SELECT people
oldOrdpeople=par1
SET ORDER TO par1
frmTop.grdPers.Column1.Header1.Caption=IIF(SYS(21)='1','*№','№')
frmTop.grdPers.Column2.Header1.Caption=IIF(SYS(21)='2','*фио','фио') 
frmTop.Refresh
DO changeRowGrdPers
*************************************************************************************************************************
PROCEDURE showDekStav
*IF !EMPTY(curJobSupl.fiodek)
*    frmTop.grdJob.toolTipText='д/о '+ALLTRIM(curJobSupl.fiodek)+' ('+LTRIM(STR(curJobSupl.kdek))+')'
* ELSE  
*    frmTop.grdJob.toolTipText=IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,'')
* ENDIF 
frmTop.grdJob.toolTipText=IIF(SEEK(kp,'sprpodr',1),ALLTRIM(sprpodr.namework),'')+IIF(!EMPTY(curJobSupl.fiodek),' д/о '+ALLTRIM(curJobSupl.fiodek)+' ('+LTRIM(STR(curJobSupl.kdek))+')','')
***********************************************************************************************************************
PROCEDURE rightClickImage
IF people.num=0
   RETURN 
ENDIF 
DEFINE POPUP shortpop SHORTCUT RELATIVE FROM MROW(),MCOL() FONT dFontName,dFontSize COLOR SCHEME 4
DEFINE BAR 1 OF shortpop PROMPT 'Изменить' 
DEFINE BAR 2 OF shortpop PROMPT "\-"
DEFINE BAR 3 OF shortpop PROMPT 'Удалить'
ON SELECTION POPUP shortpop DO changePImage
ACTIVATE POPUP shortpop
**************************************************************************************************************************
PROCEDURE changePImage
men_cx=BAR()
DO CASE
   CASE men_cx=1
        newPathImage=''
        newPathImage=GETPICT()
        IF EMPTY(newPathImage)
           RETURN
        ELSE 
           _screen.Addobject('OBJPIC','IMAGE')
           _screen.ObjPic.Picture=newPathImage       
           var_Width=_screen.ObjPic.WIDTH
           var_Height=_screen.ObjPic.Height
           _screen.Removeobject('ObjPic')  
           SELECT people
           REPLACE pFoto WITH ''
           newPathImage='"'+newPathImage+'"'
           APPEND MEMO pFoto FROM &newPathImage 
        ENDIF
   CASE men_cx=3
        SELECT people
        REPLACE pFoto WITH ''     
ENDCASE 
**************************************************************************************************************************
*         Форма для ввода нового сотрудника
*************************************************************************************************************************
PROCEDURE newCard
logRecShtat=.T.
new_fio=''
new_who=''
new_whom=''
new_whomv=''
new_whomp=''
findCard=''
findWho=''
findWhom=''
findWhomv=''
findWhomp=''
findKod=0
findNid=0
logapply=.F.
log_pc=.T.
newDateIn=CTOD('  .  .    ')
=AFIELDS(arPeople,'people')
CREATE CURSOR curSuplPeople FROM ARRAY arPeople 
SELECT curSuplPeople
INDEX ON fio TAG t1

SELECT people
oldPrec=RECNO()
log_ord=SYS(21)
SET ORDER TO 1
GO BOTTOM 
new_num=num+1
kodCard=new_num
SET ORDER TO &log_ord
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl   
     .Caption='Ввод новой карточки ('+LTRIM(STR(new_num))+')'    
     DO addShape WITH 'fSupl',1,10,10,dHeight,300,8     
     .logExit=.T. 
     .procExit='DO exitNewCard' 
     .procForClick='DO lostFocusPeopleFound'
     
     DO adTBoxAsCont WITH 'fsupl','txtFio',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('WФамилия Имя ОтчествоW'),dHeight,'Фамилия Имя Отчество',1,1     
     DO addtxtboxmy WITH 'fSupl',1,.txtFio.Left+.txtFio.Width-1,.txtFio.Top,300,.F.,'new_fio',0,'DO validNewFio'
     
     DO adTBoxAsCont WITH 'fsupl','txtTabn',.txtFio.Left,.txtFio.Top+.txtFio.Height-1,.txtFio.Width,dHeight,'Номер',1,1
     DO addtxtboxmy WITH 'fSupl',2,.txtBox1.Left,.txtTabn.Top,.txtBox1.Width,.F.,'new_num',0
     .txtBox2.InputMask='99999'
     
     DO adTBoxAsCont WITH 'fsupl','txtDateIn',.txtFio.Left,.txtTabn.Top+.txtTabn.Height-1,.txtFio.Width,dHeight,'Дата приема',1,1
     DO addtxtboxmy WITH 'fSupl',3,.txtBox1.Left,.txtDateIn.Top,.txtBox1.Width,.F.,'newDateIn',0
               
     .Shape1.Height=.txtFio.height*3+20
     .Shape1.Width=.txtFio.Width+.txtBox1.Width-1+20
     
     DO addShape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,dHeight,.shape1.Width,8  
     
     DO adTBoxAsCont WITH 'fsupl','txtWho',.txtFio.Left,.Shape2.Top+10,.txtFio.Width,dHeight,'принять кого',1,1     
     DO addtxtboxmy WITH 'fSupl',11,.txtBox1.Left,.txtWho.Top,.txtBox1.Width,.F.,'new_who',0
     
     DO adTBoxAsCont WITH 'fsupl','txtWhom',.txtFio.Left,.txtWho.Top+.txtWho.Height-1,.txtFio.Width,dHeight,'предоставить кому',1,1
     DO addtxtboxmy WITH 'fSupl',12,.txtBox1.Left,.txtWhom.Top,.txtBox1.Width,.F.,'new_whom',0
     
     DO adTBoxAsCont WITH 'fsupl','txtWhomv',.txtFio.Left,.txtWhom.Top+.txtWhom.Height-1,.txtFio.Width,dHeight,'заключить с кем',1,1
     DO addtxtboxmy WITH 'fSupl',13,.txtBox1.Left,.txtWhomv.Top,.txtBox1.Width,.F.,'new_whomv',0
     
     DO adTBoxAsCont WITH 'fsupl','txtWhomp',.txtFio.Left,.txtWhomv.Top+.txtWhomv.Height-1,.txtFio.Width,dHeight,'заявление кого',1,1
     DO addtxtboxmy WITH 'fSupl',14,.txtBox1.Left,.txtWhomp.Top,.txtBox1.Width,.F.,'new_whomp',0
     .Shape2.Height=.txtWhom.Height*4+20
         
     DO adCheckBox WITH 'fSupl','checkPc','После записи перейти к личной карточке',.Shape2.Top+.Shape2.Height+20,.Shape1.Left,150,dHeight,'log_pc',0,.T.
  
     .checkPc.Left=.Shape1.Left+(.Shape1.Width-.checkPc.Width)/2 
            
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('Wкопировать изW')*3-20)/2,.checkPc.Top+.checkPc.Height+10,RetTxtWidth('Wкопировать изW'),dHeight+3,'записать','DO writeNewCard'
     DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+3,'копировать из','DO copyFromCard'  
     DO addcontlabel WITH 'fSupl','cont3',.Cont2.Left+.Cont2.Width+10,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','DO exitNewCard'  
     
      DO adLabMy WITH 'fSupl',1,'Для подтверждения намерений поставьте',.checkPc.Top,.Shape2.Left,.Shape2.Width,2     
     .lab1.Visible=.F.
     
     DO addcontlabel WITH 'fSupl','contRet',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WВозвратW'))/2,.cont1.Top,RetTxtWidth('WВозвратW'),dHeight+3,'возврат','DO ReturnToNewCard'
     .contRet.Visible=.F.    
     
     DO addcontlabel WITH 'fSupl','contApply',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WwпринятьwW')*2-10)/2,.cont1.Top,RetTxtWidth('WwпринятьW'),dHeight+3,'принять','DO applyfound WITH .T.'
     .contApply.Visible=.F.    
     
      DO addcontlabel WITH 'fSupl','retApply',.contApply.Left+.contApply.Width+10,.contApply.Top,.contApply.Width,dHeight+3,'возврат','DO applyfound WITH .F.'
     .retApply.Visible=.F.       
   
     DO adLabMy WITH 'fSupl',2,'введите несколько символов ФИО',.Shape1.Top+10,.Shape1.Left,.Shape1.Width,2     
     .lab2.Visible=.F.          
     DO addtxtboxmy WITH 'fSupl',111,.txtFio.Left,.lab2.Top+.lab2.Height,.txtFio.Width+.txtBox1.Width-RetTxtWidth('w...')-1,.F.,'findCard',0,.F.
     WITH .txtBox111
          .Visible=.F.
          *.procForKeyPress='DO pressFindNewCard WITH 1' 
           .procForClick='DO lostFocusPeopleFound'
           .procforChange='DO changePeopleFound'                
     ENDWITH 
     DO adtboxnew WITH 'fSupl','boxFree',.txtBox111.Top,.txtBox111.Left+.txtBox111.Width-1,.Shape1.Width-.txtBox111.Width-19,dheight,'',.F.,.F.     
     DO addconticonew WITH 'fSupl','butPodr',.boxFree.Left+2,.boxFree.Top+2,'sbdn.ico',RetTxtWidth('w...')-1,.boxFree.height-4,16,16,'DO selectPeopleFound'      
     .boxFree.Visible=.F.
     .butPodr.Visible=.F. 
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape2.Height+.checkPc.Height+.cont1.Height+60      
     DO addListBoxMy WITH 'fSupl',1,.txtBox111.Left,.txtBox111.Top+.txtBox111.Height-1,300,.Shape1.Width-20  
     WITH .listBox1
          .RowSource='curSuplPeople.fio'           
          .RowSourceType=2
          .ColumnCount=1         
          .Visible=.F.        
          .controlSource=''
          .Height=.Parent.Height-.Parent.txtBox111.Top
          .procForValid='DO validListBoxPeople'  
          .procForKeyPress='DO keyPressListPeople'         
          .Height=.Parent.Height-.Top
     ENDWITH                                 
        
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
***********************************************************************************************************************
PROCEDURE copyFromCard
findCard=''
findWho=''
findWhom=''
findWhomv=''
findWhomp=''
findKod=0
findNid=0
WITH fSupl
     .setAll('Visible',.F.,'textBoxAsCont')
     .setAll('Visible',.F.,'myTxtBox')
     .setAll('Visible',.F.,'myCommandButton')
     .setAll('Visible',.F.,'myContLabel')
     .Shape1.Visible=.F.
     .Shape2.Visible=.F.
  *   .lab1.Visible=.F.
     .checkPc.Visible=.F.
     .lab2.Visible=.T.
     .txtBox111.Visible=.T.
     .boxFree.Visible=.T.
     .butPodr.visible=.T.
     .contApply.Visible=.T.
     .retApply.Visible=.T.
     
ENDWITH
***********************************************************************************************************************
PROCEDURE applyFound
PARAMETERS par1
WITH fsupl
     logapply=IIF(par1,.T.,.F.)        
     .setAll('Visible',.T.,'textBoxAsCont')
     .setAll('Visible',.T.,'myTxtBox')
     .setAll('Visible',.T.,'myCommandButton')
     .setAll('Visible',.T.,'myContLabel')
     .contRet.Visible=.F.
     .contApply.Visible=.F.
     .retApply.Visible=.F.
     .lab2.Visible=.F.
     .boxFree.Visible=.F.
     .butPodr.visible=.F.
     .checkPc.Visible=.T.
     .Shape1.Visible=.T.
     .Shape2.Visible=.T.
     .txtBox111.Visible=.F.
     IF par1
        new_fio=findCard
        new_who=findWho
        new_whom=findWhom
        new_whomv=findWhomv
        new_whomp=findWhomp
        .Refresh        
     ENDIF 
ENDWITH
************************************************************************************************************************
PROCEDURE selectPeopleFound
SELECT curSuplPeople
ZAP
APPEND FROM peopout
WITH fSupl
    .listBox1.RowSource='ALLTRIM(curSuplPeople.fio)'
     SELECT curSuplPeople
     LOCATE FOR fio=fSupl.txtBox111.Text
     IF .listBox1.Visible=.F.
        .listBox1.Visible=.T.  
        .listBox1.SetFocus        
     ENDIF     
ENDWITH 
************************************************************************************************************************
PROCEDURE changePeopleFound
WITH fSupl 
     IF .listBox1.Visible=.F.
        .listBox1.Visible=.T.      
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=fSupl.txtBox111.Text 
SELECT curSuplPeople
ZAP
APPEND FROM peopout FOR LOWER(ALLTRIM(lcValue))$LOWER(fio)
WITH fSupl.listBox1
     .RowSource='ALLTRIM(curSuplPeople.fio)'  
ENDWITH 
************************************************************************************************************************
PROCEDURE lostFocusPeopleFound
WITH fSupl
     IF fSupl.Visible=.T.
        .listBox1.Visible=.F.
        .txtBox111.setFocus
        .Refresh
     ENDIF  
ENDWITH
************************************************************************************************************************
PROCEDURE validListBoxPeople
fSupl.listBox1.Visible=.F.
findCard=curSuplPeople.fio
findWho=curSuplPeople.fior
findWhom=curSuplPeople.fiod
findWhomv=curSuplPeople.fiov
findWhomp=curSuplPeople.fiot
findKod=curSuplPeople.num
findNid=curSuplPeople.nid
fSupl.Refresh
fSupl.txtBox111.SetFocus

************************************************************************************************************************
PROCEDURE keyPressListPeople
DO CASE
   CASE LASTKEY()=27
        DO lostFocusPeopleFound
   CASE LASTKEY()=13
        DO validListBoxPeople
ENDCASE
************************************************************************************************************************
PROCEDURE validNewFio
IF EMPTY(new_fio)
   RETURN
ENDIF
DO unosimbol WITH 'new_fio'
new_pfio=ALLTRIM(LEFT(ALLTRIM(new_fio),AT(' ',ALLTRIM(new_fio))))
new_pname=SUBSTR(new_fio,AT(' ',ALLTRIM(new_fio))+1)
new_pname=ALLTRIM(LEFT(ALLTRIM(new_pname),AT(' ',ALLTRIM(new_pname))))
new_potch=ALLTRIM(SUBSTR(new_fio,RAT(' ',ALLTRIM(new_fio))))
DO procpadej WITH 'new_pfio','new_pname','new_potch','new_who','new_whom','new_whomv','new_whomp'
SELECT people
fSupl.Refresh
****************************************************************************************************************************
*
**************************************************************************************************************************** 
PROCEDURE procpadej
PARAMETERS parf,parn,paro,parRod,parDat,parVin,parTv
*--Родительный падеж
strnewR1=''
strnewR2=''
strnewR3=''

*--Дательный падеж
strnewD1=''
strnewD2=''
strnewD3=''

*--Винительный падеж
strnewV1=''
strnewV2=''
strnewV3=''

*--Творительный падеж
strnewT1=''
strnewT2=''
strnewT3=''
*str_ch=&par1
strnewR1=LEFT(ALLTRIM(&parf),1)+LOWER(SUBSTR(ALLTRIM(&parf),2))
strnewR2=ALLTRIM(&parn)
strnewR3=ALLTRIM(&paro)

strnewD1=LEFT(ALLTRIM(&parf),1)+LOWER(SUBSTR(ALLTRIM(&parf),2))
strnewD2=ALLTRIM(&parn)
strnewD3=ALLTRIM(&paro)

strnewV1=LEFT(ALLTRIM(&parf),1)+LOWER(SUBSTR(ALLTRIM(&parf),2))
strnewV2=ALLTRIM(&parn)
strnewV3=ALLTRIM(&paro)

strnewT1=LEFT(ALLTRIM(&parf),1)+LOWER(SUBSTR(ALLTRIM(&parf),2))
strnewT2=ALLTRIM(&parn)
strnewT3=ALLTRIM(&paro)
IF RIGHT(ALLTRIM(&paro),2)='на' && женский род
   *----Имя  
   DO CASE
      CASE LOWER(RIGHT(ALLTRIM(&parn),2))='га'
           strnewR2=LEFT(strnewR2,LEN(strnewR2)-1)+'у'
           strnewD2=LEFT(strnewR2,LEN(strnewR2)-1)+'е'
           strnewV2=LEFT(strnewR2,LEN(strnewR2)-1)+'ой'          
      CASE LOWER(RIGHT(ALLTRIM(&parn),1))='а'
           strnewR2=LEFT(strnewR2,LEN(strnewR2)-1)+'у'
           strnewD2=LEFT(strnewR2,LEN(strnewR2)-1)+'е'
           strnewV2=LEFT(strnewR2,LEN(strnewR2)-1)+'ой'
      CASE LOWER(RIGHT(ALLTRIM(&parn),2))='ия'
           strnewR2=LEFT(strnewR2,LEN(strnewR2)-1)+'ю'
           strnewD2=LEFT(strnewR2,LEN(strnewR2)-1)+'и'
           strnewV2=LEFT(strnewR2,LEN(strnewR2)-1)+'ей'    
       CASE LOWER(RIGHT(ALLTRIM(&parn),2))='ья'
           strnewR2=LEFT(strnewR2,LEN(strnewR2)-1)+'ю'
           strnewD2=LEFT(strnewR2,LEN(strnewR2)-1)+'е'
           strnewV2=LEFT(strnewR2,LEN(strnewR2)-1)+'ей'         
   ENDCASE      
   strnewT2=UPPER(LEFT(strnewR2,1))+'.'
   *--------------------Фамилия
   DO CASE
      CASE LOWER(RIGHT(ALLTRIM(&parf),2))='ая'      
           strnewR1=LEFT(strnewR1,LEN(strnewR1)-2)+'ую'
           strnewD1=LEFT(strnewD1,LEN(strnewD1)-2)+'ой'
           strnewV1=LEFT(strnewV1,LEN(strnewV1)-2)+'ой'
           strnewT1=LEFT(strnewT1,LEN(strnewT1)-2)+'ой'
      CASE LOWER(RIGHT(ALLTRIM(&parf),2))='га'      
      CASE LOWER(RIGHT(ALLTRIM(&parf),1))='а'      
           strnewR1=LEFT(strnewR1,LEN(strnewR1)-1)+'у'
           strnewD1=LEFT(strnewD1,LEN(strnewD1)-1)+'ой'
           strnewV1=LEFT(strnewV1,LEN(strnewV1)-1)+'ой'     
           strnewT1=LEFT(strnewT1,LEN(strnewT1)-1)+'ой' 
      CASE LOWER(RIGHT(ALLTRIM(&parf),2))='ва'           
           strnewR1=LEFT(strnewR1,LEN(strnewR1)-1)+'у'
           strnewD1=LEFT(strnewD1,LEN(strnewD1)-1)+'ой'
           strnewV1=LEFT(strnewV1,LEN(strnewV1)-1)+'ой'
           strnewT1=LEFT(strnewT1,LEN(strnewT1)-1)+'ой'
      CASE LOWER(RIGHT(ALLTRIM(&parf),1))='я'           
           strnewR1=LEFT(strnewR1,LEN(strnewR1)-1)+'ю'
           strnewD1=LEFT(strnewD1,LEN(strnewD1)-1)+'е'
           strnewV1=LEFT(strnewV1,LEN(strnewV1)-1)+'ей'   
           strnewT1=LEFT(strnewT1,LEN(strnewT1)-1)+'ей'     
   ENDCASE    
ELSE                            && мужской род 
   *----Имя 
   DO CASE
      CASE INLIST(LOWER(RIGHT(ALLTRIM(&parn),1)),'в','д','с','г','р','п','т','н','б')      
           strnewR2=ALLTRIM(strnewR2)+'а'
           strnewD2=ALLTRIM(strnewD2)+'у'
           strnewV2=ALLTRIM(strnewV2)+'ом'    
          * strnewT2=LEFT(strnewT2,LEN(strnewV2)-1)+'а'       
      CASE INLIST(LOWER(RIGHT(ALLTRIM(&parn),1)),'й')
           strnewR2=LEFT(strnewR2,LEN(strnewR2)-1)+'я'
           strnewD2=LEFT(strnewD2,LEN(strnewD2)-1)+'ю'
           strnewV2=LEFT(strnewV2,LEN(strnewV2)-1)+'ем'
         *  strnewT2=LEFT(strnewT2,LEN(strnewT2)-1)+'я'
   ENDCASE  
   strnewT2=UPPER(LEFT(strnewR2,1))+'.' 
    *--------------------Фамилия
   DO CASE
      CASE LOWER(RIGHT(ALLTRIM(&parf),2))='ий'      
           strnewR1=LEFT(strnewR1,LEN(strnewR1)-2)+'ого'
           strnewD1=LEFT(strnewD1,LEN(strnewD1)-2)+'ому'
           strnewV1=LEFT(strnewV1,LEN(strnewV1)-2)+'им'
           strnewT1=LEFT(strnewT1,LEN(strnewT1)-2)+'ого'
      CASE LOWER(RIGHT(ALLTRIM(&parf),1))='а'      
          * strnewR1=LEFT(strnewR1,LEN(strnewR1)-1)+'ы'
          * strnewD1=LEFT(strnewD1,LEN(strnewD1)-1)+'е'
          * strnewV1=LEFT(strnewV1,LEN(strnewV1)-1)+'у'  
      CASE LOWER(RIGHT(ALLTRIM(&parf),1))='я'           
           strnewR1=LEFT(strnewR1,LEN(strnewR1)-1)+'ю'
           strnewD1=LEFT(strnewD1,LEN(strnewD1)-1)+'е'
           strnewV1=LEFT(strnewV1,LEN(strnewV1)-1)+'ей' 
           strnewT1=LEFT(strnewT1,LEN(strnewT1)-1)+'и' 
      CASE INLIST(LOWER(RIGHT(ALLTRIM(&parf),1)),'д','з','ш','с','т','ц','к','г','р')           
           strnewR1=strnewR1+'а'
           strnewD1=strnewD1+'у'
           strnewV1=strnewV1+'ом'                      
           strnewT1=strnewT1+'а'  
      CASE INLIST(LOWER(RIGHT(ALLTRIM(&parf),1)),'в','н')           
           strnewR1=strnewR1+'а'
           strnewD1=strnewD1+'у'
           strnewV1=strnewV1+'ым'                      
           strnewT1=strnewT1+'а'                          
      CASE LOWER(RIGHT(ALLTRIM(&parf),1))='о'                 
      OTHERWISE 
           strnewR1=strnewR1+'о'
           strnewD1=strnewD1+'о'
           strnewV1=strnewV1+'о'
           strnewT1=strnewT1+'о'
   ENDCASE
ENDIF 

**----------------------------------Отчество
strnewR3=IIF(LOWER(RIGHT(strnewR3,1))='ч',strnewR3+'a',LEFT(strnewR3,LEN(strnewR3)-1)+'у')
strnewD3=IIF(LOWER(RIGHT(strnewD3,1))='ч',strnewD3+'у',LEFT(strnewR3,LEN(strnewD3)-1)+'е')
strnewV3=IIF(LOWER(RIGHT(strnewV3,1))='ч',strnewV3+'ем',LEFT(strnewR3,LEN(strnewV3)-1)+'ой')
strnewT3=UPPER(LEFT(strnewT3,1))+'.'
str1=strnewR1+' '+strnewR2+' '+strnewR3
str2=strnewD1+' '+strnewD2+' '+strnewD3
str3=strnewV1+' '+strnewV2+' '+strnewV3
str4=strnewT1+' '+strnewT2+' '+strnewT3
&parRod=strnewR1+' '+strnewR2+' '+strnewR3
&parDat=strnewD1+' '+strnewD2+' '+strnewD3
&parVin=strnewV1+' '+strnewV2+' '+strnewV3
&parTv=strnewT1+' '+strnewT2+' '+strnewT3
************************************************************************************************************************
PROCEDURE returnToNewCard
WITH fSupl
     .SetAll('Enabled',.T.,'MyTxtBox')
     .SetAll('Visible',.T.,'MyCheckBox') 
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.
     .contRet.Visible=.F.           
  *   .lab1.Visible=.F.   
     DO CASE
        CASE SEEK(new_num,'people',1)
             .txtBox1.SetFocus
        CASE EMPTY(new_num)
             .txtBox1.SetFocus
        CASE EMPTY(new_fio)
             .txtFio.SetFocus
     ENDCASE
     GO oldPrec     
ENDWITH
************************************************************************************************************************
*
************************************************************************************************************************
PROCEDURE writeNewcard
SELECT people
logRetCard=.F.
DO CASE
   CASE EMPTY(new_fio)
        logRetCard=.T.
        fSupl.lab1.Caption='Не указана фамилия!'     
   CASE new_num=0
        fSupl.lab1.Caption='Не указан номер!'
        logRetCard=.T.
   CASE SEEK(new_num,'people',1)
        fSupl.lab1.Caption='Такой номер уже занят!'
        logRetCard=.T.
ENDCASE
IF logRetCard
   GO oldPrec 
   WITH fSupl
        .SetAll('Enabled',.F.,'MyTxtBox')
        .SetAll('Visible',.F.,'MyCheckBox') 
        .cont1.Visible=.F.
        .cont2.Visible=.F.    
        .cont3.Visible=.F.    
        .contRet.Visible=.T.  
      *  .lab1.Visible=.T.  
        .Refresh
   ENDWITH
   RETURN 
ENDIF 
fSupl.Visible=.F.
SELECT people
ordOld=SYS(21)
SET ORDER TO 4
SET DELETED OFF
GO BOTTOM 
newNid=nid+1
SET DELETED ON
SET ORDER TO &ordOld
APPEND BLANK
newrec=RECNO()
IF !logApply
    REPLACE num WITH new_num,nid WITH newNid,fio WITH new_fio,fior WITH new_who,fiod WITH new_whom,fiov WITH new_whomv,fiot WITH new_whomp,;
            sex WITH IIF(RIGHT(ALLTRIM(fio),1)='а',2,IIF(!EMPTY(fio),1,0)),date_in WITH newDateIn
ELSE 
    SELECT peopout
    LOCATE FOR nid=findNid
    SCATTER TO dim_apply
    SELECT people
    GATHER FROM dim_apply
    REPLACE num WITH new_num,nid WITH newNid,fio WITH new_fio,fior WITH new_who,fiod WITH new_whom,fiov WITH new_whomv,fiot WITH new_whomp,;
            sex WITH IIF(RIGHT(ALLTRIM(fio),1)='а',2,IIF(!EMPTY(fio),1,0)),date_in WITH newDateIn,date_out WITH CTOD('  .   .    ')            
    SELECT * FROM datfam WHERE nidpeop=findnid INTO CURSOR curNid READWRITE        
    SELECT curNid
    REPLACE kodpeop WITH new_num,nidpeop WITH newNid ALL
    SELECT datfam
    APPEND FROM DBF('curnid')
    
    SELECT * FROM jobbook WHERE nidpeop=findnid INTO CURSOR curNid READWRITE        
    SELECT curNid
    REPLACE kodpeop WITH new_num,nidpeop WITH newNid ALL
    SELECT jobbook
    APPEND FROM DBF('curnid')
    SELECT people  
ENDIF             
frmTop.Refresh
frmTop.grdPers.Columns(frmTop.grdPers.columnCount).SetFocus
fSupl.Release      
IF log_pc
   ON ERROR DO erSup
   SELECT people          
   DO readPersCard
   ON ERROR
ENDIF

***********************************************************************************************************************
PROCEDURE exitNewCard
fSupl.Release
SELECT people
GO oldPrec
***************************************************************************************************
*           Редактирование личной карточки (новый вариант)
***************************************************************************************************
PROCEDURE readPersCard
log_ap=.F.
STORE 0 TO new_kfam,famrec
new_nfio=''
strSex=IIF(SEEK(people.sex,'cursex',1),cursex.name,'')
strDocum=IIF(SEEK(people.viddoc,'curdocum',1),curdocum.name,'')
newVidDoc=people.viddoc
new_dBirth=CTOD('  .  .    ')
new_nidpeop=0
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
strEduc=IIF(SEEK(people.educ,'suplEducation',1),suplEducation.name,'')

strSchool=people.school
strSpecd=people.specd
strKvald=people.kvald
newKodEduc=people.educ

*strSchool1=people.school1
*strSpecd1=people.specd1
*strKvald1=people.kvald1

strFamCard=IIF(SEEK(people.family,'curFamily',1),curFamily.name,'')
strVid=IIF(SEEK(people.dog,'sprdog',1),sprdog.name,'')
strSrok=IIF(SEEK(people.kTime,'cursrok',1),cursrok.name,'')
SELECT people

fPersCard=CREATEOBJECT('FORMSUPL')
WITH fPersCard     
     .Caption='Личная карточка'
     .procExit='DO exitOutPerscard'
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
                    .ToolTipText='Изменить - правая кнопка мыши'                  
                    .procForRightClick='DO clickImageCard'                
                    IF !EMPTY(people.pFoto)
                       pathmem='d:\datPic'+LTRIM(STR(people.num))+'.pic'
                       COPY MEMO people.pFoto TO &pathmem                            
                       .Picture=pathmem                       
                       widthTxt=widthTxtShort                   
                    ELSE 
                      .Picture='nofhoto.jpg'                                                           
                    ENDIF 
               ENDWITH  
               DO addShape WITH 'oPage1',1,.formImage.Left+.formImage.Width+10,.formImage.Top,100,nParent.Width-.formImage.Width-25,8                         
               
               DO adtBoxAsCont WITH 'oPage1','contFio',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('wнациональность'),dHeight,'ФИО',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFio',.contFio.Top,.contFio.Left+.contFio.Width-1,(.Shape1.Width-.contFio.Width*2-18)/2,dHeight,'people.fio',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contSex',.contFio.Left,.contFio.Top+.contFio.Height-1,.contFio.Width,dHeight,'пол',1,1 
               DO addcombomy WITH 'oPage1',1,.tBoxFio.Left,.contSex.Top,dHeight,.tBoxFio.Width,.T.,'strSex','curSex.name',6,'','DO procSexCard',.F.,.T.   
                      
               DO adtBoxAsCont WITH 'oPage1','contTab',.contFio.Left,.contSex.Top+.contSex.Height-1,.contFio.Width,dHeight,'таб.номер',1,1 
               DO adTboxNew WITH 'oPage1','tBoxTab',.contTab.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'people.tabn',.F.,.T.,0   
               
               DO adtBoxAsCont WITH 'oPage1','contPnum',.contFio.Left,.contTab.Top+.contTab.Height-1,.contFio.Width,dHeight,'личный номер',1,1 
               DO adTboxNew WITH 'oPage1','tBoxPnum',.contPnum.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'people.pnum',.F.,.T.,0 
                              
               DO adtBoxAsCont WITH 'oPage1','contAge',.contFio.Left,.contPnum.Top+.contPnum.Height-1,.contFio.Width,dHeight,'дата рождения',1,1 
               DO adTboxNew WITH 'oPage1','tBoxAge',.contAge.Top,.tBoxFio.Left,.tBoxFio.Width,dHeight,'people.Age',.F.,.T.,0 
                              
               DO adtBoxAsCont WITH 'oPage1','contPBirth',.tBoxFio.Left+.tBoxFio.Width-1,.contFio.Top,RetTxtWidth('место жительства'),dHeight,'место рождения',1,1 
               DO adTboxNew WITH 'oPage1','tBoxPBirth',.contPBirth.Top,.contPBirth.Left+.contPBirth.Width-1,.Shape1.Width-.contFio.Width-.tBoxFio.Width-.contPBirth.Width-18,dHeight,'people.placeborn',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contReg',.contPBirth.Left,.contPBirth.Top+.contPBirth.Height-1,.contPBirth.Width,dHeight,'зарегистрирован',1,1 
               DO adTboxNew WITH 'oPage1','tBoxReg',.contReg.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'people.pReg',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contLive',.contPBirth.Left,.contReg.Top+.contReg.Height-1,.contPBirth.Width,dHeight,'проживает',1,1 
               DO adTboxNew WITH 'oPage1','tBoxLive',.contLive.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'people.ppreb',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contPhone',.contPBirth.Left,.contLive.Top+.contLive.Height-1,.contPBirth.Width,dHeight,'телефон дом.',1,1 
               DO adTboxNew WITH 'oPage1','tBoxPhone',.contPhone.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'people.telhome',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contFree',.contPBirth.Left,.contPhone.Top+.contPhone.Height-1,.contPBirth.Width,dHeight,'телефон моб.',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFree',.contFree.Top,.tBoxPBirth.Left,.tBoxPBirth.Width,dHeight,'people.telmob',.F.,.T.,0 
               
                          
               
               .Shape1.Height=.contFio.Height*5+20
               *-------место рождения
               DO addShape WITH 'oPage1',2,.formImage.Left,.formImage.Top+.formImage.Height+5,100,(nParent.Width-25)/2,8 
               
               DO adtBoxAsCont WITH 'oPage1','contDoc',.Shape2.Left+10,.Shape2.Top+10,RetTxtWidth('место жительства'),dHeight,'документ',1,1 
               DO addcombomy WITH 'oPage1',2,.contDoc.Left+.contDoc.Width-1,.contDoc.Top,dHeight,.Shape2.Width-.contDoc.Width-19,.T.,'strDocum','curDocum.name',6,'','DO proccardviddoc',.F.,.T.   
               
               DO adtBoxAsCont WITH 'oPage1','contDNum',.contDoc.Left,.contDoc.Top+.contDoc.Height-1,.contDoc.Width,dHeight,'номер',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDNum',.contDNum.Top,.comboBox2.Left,.comboBox2.Width,dHeight,'people.nDoc',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contDIn',.contDoc.Left,.contDnum.Top+.contDNum.Height-1,.contDoc.Width,dHeight,'дата выдачи',1,1
               DO adTboxNew WITH 'oPage1','tBoxDIn',.contDIn.Top,.comboBox2.Left,.comboBox2.Width,dHeight,'people.ddoc',.F.,.T.,0 
                
               DO adtBoxAsCont WITH 'oPage1','contDWho',.contDoc.Left,.contDIn.Top+.contDIn.Height-1,.contDoc.Width,dHeight,'кем выдан',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDWho',.contDWho.Top,.comboBox2.Left,.comboBox2.Width,dHeight,'people.vdoc',.F.,.T.,0 
               
               DO adtBoxAsCont WITH 'oPage1','contDSrok',.contDoc.Left,.contDWho.Top+.contDWho.Height-1,.contDoc.Width,dHeight,'срок действия',1,1 
               DO adTboxNew WITH 'oPage1','tBoxDStok',.contDSrok.Top,.comboBox2.Left,.comboBox2.Width,dHeight,'people.srokdoc',.F.,.T.,0 
                              
               .Shape2.Height=.contDoc.Height*5+20
               
               *-------остальное   
               DO addShape WITH 'oPage1',3,.Shape2.Left+.Shape2.Width+10,.Shape2.Top,.Shape2.Height,.Shape2.Width,8                         
                     
                                              
               DO adtBoxAsCont WITH 'oPage1','contEduc',.Shape3.Left+10,.Shape3.Top+10,RetTxtWidth('семейное положениеw'),dHeight,'образование',1,1 
               DO addcombomy WITH 'oPage1',3,.contEduc.Left+.contEduc.Width-1,.contEduc.Top,dHeight,.Shape3.Width-.contEduc.Width-18,.T.,'strEduc','suplEducation.name',6,'','DO procCardEduc',.F.,.T.
               DO adtBoxAsCont WITH 'oPage1','contFam',.contEduc.Left,.contEduc.Top+.contEduc.Height-1,.contEduc.Width,dHeight,'семейное положение',1,1 
               DO addcombomy WITH 'oPage1',4,.comboBox3.Left,.contFam.Top,dHeight,.comboBox3.Width,.T.,'strFamCard','curFamily.name',6,'','DO proccardfamily',.F.,.T.
               DO adtBoxAsCont WITH 'oPage1','contMol',.contEduc.Left,.contFam.Top+.contFam.Height-1,.contEduc.Width,dHeight,'молодой специалист',1,1  
               oPage1.AddObject('boxMols','CONTAINER')
               WITH .boxMols
                    .BackColor=RGB(255,255,255)
                    .Visible=.T.
                    .Top=oPage1.contMol.Top
                    .Left=oPage1.comboBox3.Left
                    .Width=oPage1.combobox3.Width
                    .Height=dHeight
                    .AddObject('check1','myCheckBox')
                    WITH .check1
                         .Caption=''
                         .Left=5
                         .Top=2
                         .Visible=.T.
                         .contRolSource='people.mols'
                         .BackStyle=0
                         .AutoSize=.T.
                    ENDWITH 
                    .AddObject('txtBox1','myTxtBox')                   
                    WITH .txtBox1
                         .Left=opage1.boxmols.check1.Left+opage1.boxMols.check1.Width+10
                         .Top=0
                          .Height=opage1.boxMols.Height
                         .Width=RetTxtWidth('99/99/99999')
                         .ControlSource='people.dmol'
                    ENDWITH
               ENDWITH
                 
               DO adtBoxAsCont WITH 'oPage1','contDek',.contEduc.Left,.contMol.Top+.contMol.Height-1,.contEduc.Width,dHeight,'декретный отпуск',1,1
               oPage1.AddObject('contBox','CONTAINER')
               WITH .contBox
                    .BackColor=RGB(255,255,255)
                    .Visible=.T.
                    .Top=oPage1.contDek.Top
                    .Left=oPage1.comboBox3.Left
                    .Width=oPage1.combobox3.Width
                    .Height=dHeight
                    .AddObject('check1','myCheckBox')
                    WITH .check1
                         .Caption=''
                         .Left=5
                         .Top=2
                         .Visible=.T.
                         .contRolSource='people.dekOtp'
                         .procForValid='DO validCheckDek'
                         .BackStyle=0
                         .AutoSize=.T.  
                                                
                    ENDWITH  
                    .AddObject('txtBox1','myTxtBox')                   
                    WITH .txtBox1
                         .Left=opage1.contBox.check1.Left+opage1.contBox.check1.Width+10
                         .Top=0
                         .Height=opage1.contBox.Height
                         .Width=RetTxtWidth('99/99/99999')
                         .ControlSource='people.bdekotp'
                    ENDWITH  
                    .AddObject('txtBox2','myTxtBox')                   
                    WITH .txtBox2
                         .Left=opage1.contBox.txtbox1.Left+opage1.contBox.txtbox1.Width-1
                         .Top=opage1.contBox.txtbox1.Top
                         .Height=opage1.contBox.txtBox1.Height
                         .Width=opage1.contBox.txtBox1.Width
                         .ControlSource='people.ddekotp'
                    ENDWITH  
                                    
               ENDWITH
               DO adtBoxAsCont WITH 'oPage1','contFree1',.contEduc.Left,.contDek.Top+.contDek.Height-1,.contEduc.Width,dHeight,'',1,1 
               DO adTboxNew WITH 'oPage1','tBoxFree1',.contFree1.Top,.comboBox3.Left,.comboBox3.Width,dHeight,'',.F.,.F.,0                          
               DO addShape WITH 'oPage1',4,.Shape2.Left,.Shape2.Top+.Shape2.Height+10,100,.Shape2.Width+.Shape3.Width+10,8 
               DO adCheckBox WITH 'oPage1','checkVn','внешний совместитель',.Shape4.Top+10,.Shape2.Left+10,150,dHeight,'people.lvn',0   
               DO adCheckBox WITH 'oPage1','checkUnion','член профсоюза',.checkVn.Top,.Shape2.Left+10,150,dHeight,'people.Union',0   
               DO adCheckBox WITH 'oPage1','checkPens','пенсионер',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'people.Pens',0   
               DO adCheckBox WITH 'oPage1','checkInv','инвалид',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'people.inv',0   
               DO adCheckBox WITH 'oPage1','checkMany','многодетный',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'people.mchild',0   
               DO adCheckBox WITH 'oPage1','checkAes','ЧАЭС',.checkUnion.Top,.Shape4.Left+10,150,dHeight,'people.chaes',0   
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
          WITH .mpage2
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opage2=.Parent.mPage2
               .Caption='образование'  
               DO procEducationCard
          ENDWITH  
          .AddObject('mpage3','myPage')
          WITH .mpage3     
               nParent=.Parent
               .Caption='состав семьи' 
               opage3=.Parent.mPage3 
               .BackColor=RGB(255,255,255)
               DO procFamilyCard                                                            
          ENDWITH
          .AddObject('mpage4','myPage')
          WITH .mpage4
               nParent=.Parent
               .Caption='дополнительно'  
               opageSup=.Parent.mPage4
               .BackColor=RGB(255,255,255)
               hs4='Номер карточки (изменить двойной щелчок мыши)'
               newNumCard=people.num
               WITH oPageSup
                    DO addShape WITH 'opageSup',1,5,10,100,nParent.Width-20,8                                        
                    DO adtBoxAsCont WITH 'opageSup','contRp',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('Молодой специалист до  '),dHeight,'принять кого',1,1 
                    DO adTboxNew WITH 'opageSup','tBoxRp',.contRp.Top,.contRp.Left+.contRp.Width-1,(.Shape1.Width-.contRp.Width*2-18)/2,dHeight,'people.fior',.F.,.T.,0 
                    
                    DO adtBoxAsCont WITH 'opageSup','contDp',.contRp.Left,.contRp.Top+.contRp.Height-1,.contRp.Width,dHeight,'предоставить кому',1,1 
                    DO adTboxNew WITH 'opageSup','tBoxDp',.contDp.Top,.tBoxRp.Left,.tBoxRp.Width,dHeight,'people.fiod',.F.,.T.,0 
                    
                    DO adtBoxAsCont WITH 'opageSup','contVp',.contRp.Left,.contDp.Top+.contDp.Height-1,.contRp.Width,dHeight,'заключить с кем ',1,1 
                    DO adTboxNew WITH 'opageSup','tBoxVp',.contVp.Top,.tBoxRp.Left,.tBoxRp.Width,dHeight,'people.fiov',.F.,.T.,0 
                    
                    DO adtBoxAsCont WITH 'opageSup','contVt',.contRp.Left,.contVp.Top+.contVp.Height-1,.contRp.Width,dHeight,'заявление кого ',1,1 
                    DO adTboxNew WITH 'opageSup','tBoxVt',.contVt.Top,.tBoxRp.Left,.tBoxRp.Width,dHeight,'people.fiot',.F.,.T.,0    
                                                                            
                    DO addContFormNew WITH 'oPageSup','contNum',.tBoxRp.Left+.tBoxRp.Width+10,.contRp.Top,RetTxtWidth('wНомер карточки (изменить двойной щелчок мыши)w'),dHeight,hs4,0,.F.,'DO readNumCard',.F.,.F.                   
                    DO adTboxNew WITH 'opageSup','tBoxNum',.contNum.Top,.contNum.Left+.contNum.Width-1,RetTxtWidth('999999'),dHeight,'newNumCard',.F.,.F.,0                                                 
                    
                    *--------------------------------Кнопка сохранить-------------------------------------------------------------------------------------------------
                    DO addcontlabel WITH 'oPageSup','butSave',.contNum.Left,.contNum.Top+.contNum.Height+20,RetTxtWidth('wсохранитьw'),dHeight+5,'сохранить','DO saveNumCard WITH .T.'
                    *---------------------------------Кнопка отмена ------------------------------------------------------------------------
                    DO addcontlabel WITH 'oPageSup','butRet',.butSave.Left,.butSave.Top,.butSave.Width,dHeight+5,'отмена','DO saveNumCard WITH .F.'          
                    .butSave.Left=.contNum.Left+(.contNum.Width+.tBoxNum.Width-.butSave.Width*2-20)/2
                    .butRet.Left=.butSave.Left+.butSave.Width+20
                    .butSave.Visible=.F.
                    .butRet.Visible=.F.
                    
                     .Shape1.Height=.contRp.Height*4+17                  
               ENDWITH
          ENDWITH         
          .AddObject('mpage5','myPage')          
          WITH .mpage5
               nParent=.Parent
               .Caption='контр.,катег.'  
               opage5=.Parent.mPage5
               .BackColor=RGB(255,255,255)
               DO procKontrakt
          ENDWITH    
          .AddObject('mpage6','myPage')          
          WITH .mpage6
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageKurs=.Parent.mPage6
               .Caption='курсы'  
               DO procKursPCard
          ENDWITH 
          .AddObject('mpage7','myPage')          
          WITH .mpage7
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageBook=.Parent.mPage7
               .Caption='трудовая книжка'  
               DO procJobBook
          ENDWITH 
          .AddObject('mpage8','myPage')          
          WITH .mpage8
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opage8=.Parent.mPage8
               .Caption='воиский учёт'  
               DO procArmy
          ENDWITH     
          .AddObject('mpage9','myPage')          
          WITH .mpage9
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageBol=.Parent.mPage9
               .Caption='больничные листы'  
               DO procListBol
          ENDWITH 
          .AddObject('mpage10','myPage')          
          WITH .mpage10
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageOrd=.Parent.mPage10
               .Caption='приказы'  
               DO procOrderPeop
          ENDWITH               
         .AddObject('mpage11','myPage') 
          WITH .mpage11
               nParent=.Parent
               .BackColor=RGB(255,255,255)
               opageAward=.Parent.mPage11
               .Caption='награды'  
               DO procAward
          ENDWITH      
           
     ENDWITH   
    .SHOW
ENDWITH
***********************************************************************************************************************
PROCEDURE exitOutPersCard
SELECT people
fPersCard.Visible=.F.
DO changeRowGrdPers
***********************************************************************************************************************
PROCEDURE readNumCard
WITH opageSup
     .butSave.Visible=.T.
     .butRet.Visible=.T.
     .tBoxNum.Enabled=.T.
ENDWITH
***********************************************************************************************************************
PROCEDURE saveNumCard 
PARAMETERS par1
SELECT people
oldNumRec=RECNO()
IF par1
   DO CASE
      CASE newNumCard=people.num           
           RETURN
      CASE SEEK(newNumCard,'people',1) 
           GO oldNumRec
           RETURN
      OTHERWISE
           SELECT people
           GO oldNumRec
           SELECT datjob
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT datotp
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT datfam
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT datalist
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT datkurs
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT peoporder
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT jobbook
           REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
          * SELECT datimage           
          * REPLACE kodpeop WITH newNumCard FOR kodpeop=people.num
           SELECT people
           REPLACE num WITH newNumCard         
   ENDCASE   
ENDIF
WITH opageSup
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .tBoxNum.Enabled=.F.
ENDWITH
***********************************************************************************************************************
PROCEDURE validCheckDek
SELECT datJob
SET ORDER TO 1
SEEK people.num
REPLACE dekotp WITH people.dekOtp WHILE kodpeop=people.num
SELECT people
***************************************************************************************************
PROCEDURE rightClickImage
IF people.num=0
   RETURN 
ENDIF 
***************************************************************************************************
PROCEDURE clickImageCard
DEFINE POPUP shortpop SHORTCUT RELATIVE FROM MROW(),MCOL() FONT dFontName,dFontSize COLOR SCHEME 4
DEFINE BAR 1 OF shortpop PROMPT 'Изменить' 
DEFINE BAR 2 OF shortpop PROMPT "\-"
DEFINE BAR 3 OF shortpop PROMPT 'Удалить'
ON SELECTION POPUP shortpop DO changePhoto
ACTIVATE POPUP shortpop
**************************************************************************************************************************
PROCEDURE changePhoto
men_cx=BAR()
DO CASE
   CASE men_cx=1
        newPathImage=''
        newPathImage=GETPICT()
        IF EMPTY(newPathImage)
           RETURN
        ELSE 
           _screen.Addobject('OBJPIC','IMAGE')
           _screen.ObjPic.Picture=newPathImage       
           var_Width=_screen.ObjPic.WIDTH
           var_Height=_screen.ObjPic.Height
           _screen.Removeobject('ObjPic')  
           SELECT people
           newPathImage='"'+newPathImage+'"'
           REPLACE pFoto WITH ''
           APPEND MEMO pFoto FROM &newPathImage 
        ENDIF   
   CASE men_cx=3
        SELECT people
        REPLACE pFoto WITH ''     
ENDCASE 
IF !EMPTY(people.pFoto)
   pathmem='d:\datPic'+LTRIM(STR(people.num))+'.pic'
   COPY MEMO people.pFoto TO &pathmem                            
   fPersCard.pagepeop.mpage1.formImage.Picture=pathmem                         
ELSE 
   fPersCard.pagepeop.mpage1.formImage.Picture='nofhoto.jpg'                                                           
ENDIF 
fPersCard.Refresh
*************************************************************************************************************
PROCEDURE procSexcard
SELECT people
REPLACE sex WITH cursex.kod
*************************************************************************************************************
PROCEDURE proccardviddoc
SELECT people
REPLACE viddoc WITH curDocum.kod
*************************************************************************************************************
PROCEDURE proccardfamily
SELECT people
REPLACE family WITH curFamily.kod
*************************************************************************************************************
PROCEDURE proccardeduc
SELECT people
REPLACE educ WITH suplEducation.kod
fpersCard.pagePeop.mpage2.comboBox1.Refresh
*************************************************************************************************************************
*                         Редактирование состава сведений об образованиив личной карточке (новый вариант)
*************************************************************************************************************************
PROCEDURE procEducationCard
PARAMETERS parUv
SELECT school DISTINCT FROM people WHERE !EMPTY(people.school) INTO CURSOR curSchool READWRITE
SELECT specd DISTINCT FROM people WHERE !EMPTY(people.specd) INTO CURSOR curSpecd READWRITE
SELECT kvald DISTINCT FROM people WHERE !EMPTY(people.kvald) INTO CURSOR curKvald READWRITE  
IF !parUv
   SELECT people
ELSE 
   SELECT peopout
ENDIF    

strSuplEduc=IIF(SEEK(people.Educ,'curEducation',1),ALLTRIM(curEducation.name),'')
WITH opage2   
     DO addShape WITH 'oPage2',1,20,20,100,nParent.Width-40,8                          
     DO adtBoxAsCont WITH 'oPage2','contEduc',.Shape1.Left+20,.Shape1.Top+20,RetTxtWidth('WДата выдачи дипломаW'),dHeight,'образование',1,1   
*     DO addComboMy WITH 'oPage2',1,.contEduc.Left+.contEduc.Width-1,.contEduc.Top,dheight,.Shape1.Width-.contEduc.Width-40,.T.,'strSuplEduc','suplEducation.name',6,.F.,'REPLACE people.educ WITH suplEducation.kod',.F.,.T.   
     DO addComboMy WITH 'oPage2',1,.contEduc.Left+.contEduc.Width-1,.contEduc.Top,dheight,.Shape1.Width-.contEduc.Width-40,IIF(!parUv,.T.,.F.),'strEduc','suplEducation.name',6,.F.,'DO validEducMpage2',.F.,.T.   
                          
     DO adtBoxAsCont WITH 'oPage2','contVuz',.contEduc.Left,.contEduc.Top+.contEduc.Height-1,.contEduc.Width,dHeight,'учебное заведение',1,1 
     DO addComboMy WITH 'oPage2',2,.comboBox1.Left,.contVuz.Top,dheight,.comboBox1.Width,IIF(!parUv,.T.,.F.),'strSchool','curSchool.school',6,'DO gotFocusSpecdCard','DO validSchoolCard WITH 1',.F.,.T.  
     WITH .comboBox2       
          .Style=0      
     ENDWITH
     DO adtBoxAsCont WITH 'oPage2','contSpecd',.contEduc.Left,.contVuz.Top+.contVuz.Height-1,.contEduc.Width,dHeight,'специальность',1,1 
     DO addComboMy WITH 'oPage2',3,.comboBox1.Left,.contSpecd.Top,dheight,.comboBox1.Width,IIF(!parUv,.T.,.F.),'strSpecd','curSpecd.specd',6,'DO gotFocusSpecdcard','DO validSpecdCard WITH 1',.F.,.T.  
     WITH .comboBox3
          .Style=0       
     ENDWITH
     
     DO adtBoxAsCont WITH 'oPage2','contKvald',.contEduc.Left,.contSpecd.Top+.contSpecd.Height-1,.contEduc.Width,dHeight,'квалификация',1,1 
     DO addComboMy WITH 'oPage2',4,.comboBox1.Left,.contKvald.Top,dheight,.comboBox1.Width,IIF(!parUv,.T.,.F.),'strKvald','curKvald.kvald',6,'DO gotFocusSpecdcard','DO validKvaldCard WITH 1',.F.,.T.  
     WITH .comboBox4
          .Style=0       
     ENDWITH
         
     DO adtBoxAsCont WITH 'oPage2','contEndEduc',.contEduc.Left,.contKvald.Top+.contKvald.Height-1,.contEduc.Width,dHeight,'дата окончания',1,1 
     DO adTboxNew WITH 'oPage2','tBoxEnd',.contEndEduc.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'people.endEduc',.F.,IIF(!parUv,.T.,.F.),0 
     
     DO adtBoxAsCont WITH 'oPage2','contNumDip',.contEduc.Left,.contEndEduc.Top+.contEndEduc.Height-1,.contEduc.Width,dHeight,'номер диплома',1,1
     DO adTboxNew WITH 'oPage2','tBoxNumDip',.contNumDip.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'people.numDip',.F.,IIF(!parUv,.T.,.F.),0  
     
     DO adtBoxAsCont WITH 'oPage2','contDateDip',.contEduc.Left,.contNumDip.Top+.contNumDip.Height-1,.contEduc.Width,dHeight,'дата выдачи диплома',1,1 
     DO adTboxNew WITH 'oPage2','tBoxDateDip',.contDateDip.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'people.dateDip',.F.,IIF(!parUv,.T.,.F.),0          
     .Shape1.Height=.contEduc.Height*7+40
          
     .Refresh
     .comboBox1.SetFocus
     .Refresh
ENDWITH
********************************************************************************************************************
PROCEDURE validEducMpage2
SELECT people
REPLACE educ WITH suplEducation.kod
fpersCard.pagePeop.mpage1.comboBox3.Refresh
********************************************************************************************************************
PROCEDURE validSchoolCard
PARAMETERS par1
DO CASE
   CASE par1=1
        IF EMPTY(oPage2.comboBox2.DisplayValue)=.F..AND.EMPTY(oPage2.comboBox2.Value)=.T.   
           SELECT curSchool
           APPEND BLANK
           REPLACE school WITH oPage2.comboBox2.DisplayValue 
           oPage2.comboBox2.Requery()  
           strSchool=fPersCard.pagePeop.mPage2.comboBox2.DisplayValue  	
        ENDIF
        REPLACE people.school WITH strSchool
        oPage2.ComboBox2.ControlSource='strSchool'
   CASE par1=2
        IF EMPTY(oPage2.comboBox22.DisplayValue)=.F..AND.EMPTY(oPage2.comboBox22.Value)=.T.   
           SELECT curSchool
           APPEND BLANK
           REPLACE school WITH oPage2.comboBox22.DisplayValue 
           oPage2.comboBox22.Requery()  
           strSchool1=fPersCard.pagePeop.mPage2.comboBox22.DisplayValue  	
        ENDIF
        REPLACE people.school1 WITH strSchool1
        oPage2.ComboBox22.ControlSource='strSchool1'        
ENDCASE         
oPage2.Refresh
***********************************************************************************************************************
PROCEDURE gotFocusSpecdCard
=SYS(2002,1)
***********************************************************************************************************************
PROCEDURE validSpecdcard
PARAMETERS par1
DO CASE
   CASE par1=1 
        IF EMPTY(oPage2.comboBox3.DisplayValue)=.F..AND.EMPTY(oPage2.comboBox3.Value)=.T.   
           SELECT curSpecd
           APPEND BLANK
           REPLACE specd WITH oPage2.comboBox3.DisplayValue 
           oPage2.comboBox3.Requery()  
           strSpecd=fPersCard.pagePeop.mPage2.comboBox3.DisplayValue  	
        ENDIF
        REPLACE people.specd WITH strSpecd 
        oPage2.ComboBox3.ControlSource='strSpecd'
   CASE par1=2
        IF EMPTY(oPage2.comboBox33.DisplayValue)=.F..AND.EMPTY(oPage2.comboBox33.Value)=.T.   
           SELECT curSpecd
           APPEND BLANK
           REPLACE specd WITH oPage2.comboBox33.DisplayValue 
           oPage2.comboBox33.Requery()  
           strSpecd1=fPersCard.pagePeop.mPage2.comboBox33.DisplayValue  	
        ENDIF
        SELECT people
        REPLACE specd1 WITH strSpecd1  
        oPage2.ComboBox33.ControlSource='strSpecd1'
                
ENDCASE
oPage2.Refresh
***********************************************************************************************************************
PROCEDURE validKvaldCard
PARAMETERS par1
DO CASE 
   CASE par1=1
        IF EMPTY(oPage2.comboBox4.DisplayValue)=.F..AND.EMPTY(oPage2.comboBox4.Value)=.T.   
           SELECT curKvald
           APPEND BLANK
           REPLACE kvald WITH oPage2.comboBox4.DisplayValue 
           oPage2.comboBox4.Requery()  
           strKvald=oPage2.comboBox4.DisplayValue     
        ENDIF
        REPLACE people.kvald WITH strkvald	
        oPage2.ComboBox4.ControlSource='strKvald'
   CASE par1=2
        IF EMPTY(oPage2.comboBox44.DisplayValue)=.F..AND.EMPTY(oPage2.comboBox44.Value)=.T.   
           SELECT curKvald
           APPEND BLANK
           REPLACE kvald WITH oPage2.comboBox44.DisplayValue 
           oPage2.comboBox44.Requery()  
           strKvald1=oPage2.comboBox44.DisplayValue     
        ENDIF
        SELECT people
        REPLACE kvald1 WITH strkvald1	
        oPage2.ComboBox44.ControlSource='strKvald1'       
ENDCASE 
oPage2.Refresh
***********************************************************************************************************************
PROCEDURE validProfCard
dim_school(8)=curProf.osnProf
oPage2.comboBox3.Width=objWidth-.tBox7.Width
oPage2.comboBox3.Left=.tBox7.Left+.tBox7.Width
KEYBOARD '{TAB}'
oPage2.Refresh
***************************************************************************************************
*           Редактирование состава семьи в личной карточке (новый вариант)
***************************************************************************************************
PROCEDURE procFamilyCard
PARAMETERS parUv
PUBLIC logUv
logUv=parUv
SELECT sprtot   
SELECT kod,name FROM sprtot WHERE sprtot.kspr=6 INTO CURSOR curSprFam READWRITE   && курсор для состава семьи
SELECT curSprFam
INDEX ON kod TAG T1
SET ORDER TO 1

SELECT kod,name FROM sprtot WHERE sprtot.kspr=6 INTO CURSOR curMenufam READWRITE
SELECT curMenuFam
INDEX ON name TAG T1
IF !USED('datFam')  
   USE datfam IN 0 ORDER 1
ENDIF
famRec=0
strfam=''
log_ap=.F.
new_kFam=0
new_dBirth=CTOD('  .  .    ')
new_nidPeop=0
new_nFio=''
SELECT datfam
SET RELATION TO kFam INTO curSprFam ADDITIVE
IF !parUv
   SET FILTER TO datFam.nidpeop=people.nid
ELSE 
   SET FILTER TO datFam.nidpeop=peopout.nid
ENDIF 
GO TOP
WITH oPage3     
     DO addButtonOne WITH 'oPage3','menuCont1',10,nParent.Height-dHeight-40,'новая','','DO readFamCard WITH .T.',39,RetTxtWidth('справочникw'),'новая'
     DO addButtonOne WITH 'oPage3','menuCont2',.menucont1.Left+.menucont1.Width+3,.menucont1.Top,'редакция','','DO readFamCard WITH .F.',39,.menucont1.Width,'редакция'
     DO addButtonOne WITH 'oPage3','menuCont3',.menucont2.Left+.menucont2.Width+3,.menucont1.Top,'удаление','','DO delFamily',39,.menucont1.Width,'удаление' 
     .SetAll('Enabled',IIF(!parUv,.T.,.F.),'myCommandButton')       
     .AddObject('fGrid','GRIDMY')     
     WITH .fgrid
          .Top=0
          .Left=0
          .Width=nParent.Width
          .Height=.Parent.menuCont1.Top-60
          .ScrollBars=2          
          .ColumnCount=4        
          .RecordSourceType=1     
          .RecordSource='datfam'
          .Column1.ControlSource='curSprFam.name'
          .Column2.ControlSource='" "+datfam.nfio'
          .Column3.ControlSource='" "+DTOC(datfam.dBirth)'       
          .Column3.Width=RettxtWidth(' дата рожд. ')               
          .Column1.Width=(.Width-.column1.width)/3
          .Column2.Width=.Width-.column1.width-.column3.Width-SYSMETRIC(5)-13-.ColumnCount       
           .Columns(.ColumnCount).Width=0
          .Column1.Header1.Caption='степень'
          .Column2.Header1.Caption='ФИО'
          .Column3.Header1.Caption='дата рожд.'                
          .Column1.Movable=.F. 
          .Column1.Alignment=0
          .Column2.Alignment=0           
          .Column3.Alignment=0
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.           
     ENDWITH
     DO gridSize WITH 'fPersCard.pagePeop.mPage3','fGrid','shapeingrid'
     FOR i=1 TO .fGrid.columnCount        
         .fGrid.Columns(i).DynamicBackColor='IIF(RECNO(fPersCard.pagePeop.mPage3.fGrid.RecordSource)#fPersCard.pagePeop.mPage3.fGrid.curRec,fPersCard.pagePeop.mPage3.BackColor,dynBackColor)'
         .fGrid.Columns(i).DynamicForeColor='IIF(RECNO(fPersCard.pagePeop.mPage3.fGrid.RecordSource)#fPersCard.pagePeop.mPage3.fGrid.curRec,dForeColor,dynForeColor)'        
     ENDFOR    
           
     DO addComboMy WITH 'oPage3',1,1,1,dHeight,.fgRid.Column1.Width+2,.T.,'strfam','curMenuFam.name',6,'DO gotFocuskFamCard','DO validkFamCard',.F.,.F.
     DO addtxtboxmy WITH 'oPage3',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'oPage3',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
     .menucont1.Left=(.fGrid.Width-.menucont1.Width-.menucont2.Width-.menucont3.Width-20)/2
     .menucont2.Left=.menucont1.Left+.menucont1.Width+10                   
     .menucont3.Left=.menucont2.Left+.menucont2.Width+10                                  
     IF !logUv
        *---------------------------------Кнопка записать-------------------------------------------------------------------------       
        DO addButtonOne WITH 'oPage3','butSave',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WзаписатьW')*2-20)/2,.menucont1.Top,'записать','','DO writeFamCard WITH .T.',39,RetTxtWidth('wзаписатьw'),'записать'
        *---------------------------------Кнопка возврат при редакции--------------------------------------------------------------                                                     
        DO addButtonOne WITH 'oPage3','butRet',.butSave.Left+.butsave.Width+20,.butSave.Top,'возврат','','DO writeFamCard WITH .F.',39,.butsave.Width,'возврат'
        .butSave.Visible=.F.
        .butRet.Visible=.F.          
        *---------------------------------Кнопка удалить-------------------------------------------------------------------------
        DO addButtonOne WITH 'oPage3','butDel',.fGrid.Left+(.fGrid.Width-RetTxtWidth('WудалитьW')*2-20)/2,.menucont1.Top,'удалить','','DO delRecFamily WITH .T.',39,RetTxtWidth('wудалитьw'),'удалить'
        *---------------------------------Кнопка возврат при удалении-------------------------------------------------------------------------                                            
        DO addButtonOne WITH 'oPage3','butDelRet',.butDel.Left+.butDel.Width+20,.butDel.Top,'возврат','','DO delRecFamily WITH .F.',39,.butDel.Width,'возврат'
        .butDel.Visible=.F.
        .butDelRet.Visible=.F.
     ENDIF      
     .setAll('Top',.fGrid.Top+.fGrid.height+20,'mymenuCont')   
     .setAll('Top',.fGrid.Top+.fGrid.height+20,'myCommandButton') 
     .setAll('Top',.fGrid.Top+.fGrid.height+20,'myContLabel')
ENDWITH
************************************************************************************************************************
PROCEDURE readfamCard
PARAMETERS par1
IF logUv
   RETURN
ENDIF
SELECT datfam
IF par1     
   APPEND BLANK     
   REPLACE kodpeop WITH people.num,nidpeop WITH people.nid
ENDIF
strfam=IIF(par1,'',IIF(SEEK(datFam.kfam,'curSprfam',1),curSprFam.name,''))
fPersCard.pagePeop.mPage3.Refresh
log_ap=IIF(par1,.T.,.F.)
new_kFam=IIF(par1,0,kFam)
new_nFio=IIF(par1,'',nFio)
new_nidPeop=IIF(par1,people.nid,nidPeop)
new_dBirth=IIF(par1,CTOD('  .  .    '),dBirth)
famRec=RECNO()
WITH oPage3    
     .fGrid.Refresh
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
     
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butSave.Visible=.T.
     .butRet.Visible=.T.

     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .comboBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.comboBox1.Left+.comboBox1.Width-1
     .txtBox3.Left=.txtBox2.Left+.txtBox2.Width-1
     .comboBox1.ControlSource='strFam'
     .txtbox2.ControlSource='new_nFio'
     .txtbox3.ControlSource='new_dBirth'
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .combobox1.Top=lineTop
     .combobox1.Height=.fGrid.RowHeight+1
     .combobox1.BackColor=.txtbox2.BackColor
     .SetAll('Visible',.T.,'MyTxtBox')
     .combobox1.Visible=.T.
     .fGrid.Enabled=.F.
     .comboBox1.SetFocus
ENDWITH 
***********************************************************************************************************************
PROCEDURE gotFocuskFamCard
SELECT curmenuFam
LOCATE FOR kod=datfam.kFam
oPage3.combobox1.DisplayCount=MAX(oPage3.fGrid.RelativeRow,oPage3.fGrid.RowsGrid-oPage3.fGrid.RelativeRow)
oPage3.combobox1.DisplayCount=MIN(oPage3.combobox1.DisplayCount,RECCOUNT())
SELECT datFam
************************************************************************************************************************
PROCEDURE validkFamCard
new_kFam=curMenuFam.kod
strfam=curMenufam.name
oPage3.comboBox1.ControlSource='strfam'  
KEYBOARD '{TAB}'    
************************************************************************************************************************
PROCEDURE writeFamCard
PARAMETERS par1
WITH oPage3
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     SELECT datFam
     
     GO famRec
     IF par1
        REPLACE nFio WITH new_nFio,kFam WITH new_kFam,dBirth WITH new_dBirth,nidPeop WITH new_nidpeop      
     ELSE
        IF log_ap     
           DELETE               
        ENDIF           
     ENDIF   
     .SetAll('Visible',.F.,'MyTxtBox')
     .SetAll('Visible',.F.,'ComboMy')     
     .fGrid.Enabled=.T.    
     GO famRec
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH  
oPage3.Refresh 
GO famRec
oPage3.fGrid.Columns(oPage3.fGrid.columnCount).SetFocus  
************************************************************************************************************************
PROCEDURE delFamily
IF logUv
   RETURN
ENDIF
SELECT datFam
WITH oPage3
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butDel.Visible=.T.
     .butDelRet.Visible=.T.
     .fGrid.Enabled=.F.
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE delRecFamily
PARAMETERS par1
SELECT datFam
IF par1
   DELETE
ENDIF
WITH oPage3
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.F.,'myContLabel')
     .butDel.Visible=.F.
     .butDelRet.Visible=.F.
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .fGrid.Columns(.fGrid.ColumnCount).SetFocus
ENDWITH 
************************************************************************************************************************
PROCEDURE actualstajtoday
PARAMETERS parBase,pardate,parEnd,parVar,parStart

IF !EMPTY(&parDate)
   SELECT &parBase
   currentStaj=''
   dMbeg=0
   dMEnd=0 
   newDBeg=&pardate   
   *newDEnd=&parEnd-1
   newDEnd=&parEnd
   y_stIn=IIF(EMPTY(staj_in).OR.parStart,0,ROUND(VAL(LEFT(staj_in,2)),0))
   m_stIn=IIF(EMPTY(staj_in).OR.parStart,0,VAL(SUBSTR(staj_in,4,2)))
   d_stIn=IIF(EMPTY(staj_in).OR.parStart,0,VAL(SUBSTR(staj_in,7,2)))      
   
   y_st=0
   m_st=0
   d_st=0
   
   y_new=0
   m_new=0
   d_new=0  
  
   dayMonthBeg=0
   dayMonthEnd=0 
   IF MONTH(newDBeg)=2 
      dayMonthBeg=IIF(MOD(YEAR(newDBeg),4)=0,29,28)
   ELSE
      dayMonthBeg=IIF(INLIST(MONTH(newDBeg),1,3,5,7,8,10,12),31,30) &&кол-во дней в начальном месяце  
   ENDIF
   IF MONTH(newDend)=2 
      dayMonthEnd=IIF(DAY(&parDate)=1.AND.MONTH(&parDate)=3,30,IIF(MOD(YEAR(newDEnd),4)=0,29,28))
   ELSE 
      dayMonthEnd=IIF(INLIST(MONTH(newDEnd),1,3,5,7,8,10,12),31,30)  &&кол-во дней в конечном месяце
   ENDIF
   
 *-----считаем дни
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMBeg=DAY(newDEnd)-DAY(newDBeg)+1
   ELSE 
      dMbeg=dayMonthBeg-DAY(newDBeg)+1
   ENDIF   
   IF dMBeg=dayMonthBeg
      m_new=m_new+1
      dMbeg=0
   ENDIF  
   d_new=d_new+dMBeg
 
   dMEnd=0
   IF YEAR(newDbeg)=YEAR(newDEnd).AND.MONTH(newDBeg)=MONTH(newDEnd) 
      dMEnd=0
   ELSE
      dMEnd=DAY(newDEnd)+IIF(DAY(&parDate)=1.AND.MONTH(&parDate)=3,2,0)
   ENDIF    
   IF dMEnd=dayMonthEnd
      m_new=m_new+1
      dMEnd=0
   ENDIF   
   d_new=d_New+dMEnd
 *-------считаем месяцы 
   mEbeg=0
   mYEnd=0
   IF YEAR(newDBeg)=YEAR(newDEnd)
      m_new=m_new+MONTH(newDEnd)-MONTH(newDBeg)-1
      m_new=IIF(m_new<0,0,m_new)
   ELSE 
      mYbeg=12-MONTH(newDBeg)
      mYEnd=MONTH(newDEnd)-1
      m_new=m_new+mYbeg+mYEnd
   ENDIF 
 *------------считаем годы------
   IF YEAR(newDBeg)=YEAR(newDEnd)
      y_new=0
   ELSE 
      y_new=YEAR(newDEnd)-YEAR(newDBeg)-1   
   ENDIF 

   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   newYst=y_new
   newMst=m_new
   newDst=d_new
   stajOrg=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')    
   
   
   y_new=y_new+y_stIn
   m_new=m_new+m_stIn
   d_new=d_new+d_stIn
   IF d_new>=30
      d_new=d_new-30
      m_new=m_new+1
   ENDIF     
   IF m_new>11
      y_new=y_new+1
      m_new=m_new-12
   ENDIF
   IF m_new<0
      m_new=0
      y_new=IIF(y_new=0,0,y_new-1)
   ENDIF
   currentStaj=PADL(ALLTRIM(STR(y_new)),2,'0')+'-'+PADL(ALLTRIM(STR(m_new)),2,'0')+'-'+PADL(ALLTRIM(STR(d_new)),2,'0')
   IF EMPTY(parVar)
      REPLACE staj_today WITH currentStaj   
   ELSE 
     &parVar=currentStaj
   ENDIF    
   
ELSE

ENDIF  
**********************************************************************************************************
*             Расчет переходящего стажа по одному человеку
**********************************************************************************************************
PROCEDURE perstajone1
PARAMETERS parStaj,pardstaj,parbase
IF EMPTY(parbase)
   SELECT people
ELSE  
   SELECT &parbase
ENDIF    
REPLACE dPerSt WITH CTOD('  /  /  ')
STORE 0 TO d_new,m_new,y_new
y_new=YEAR(&pardstaj)
y_st=IIF(EMPTY(&parStaj),0,ROUND(VAL(LEFT(&parStaj,2)),0))
m_st=IIF(EMPTY(&parStaj),0,VAL(SUBSTR(&parStaj,4,2)))
d_st=IIF(EMPTY(&parStaj),0,VAL(SUBSTR(&parStaj,7,2)))  
IF INLIST(y_st,4,9,14)    
   y_new=IIF((MONTH(&pardstaj)+11-m_st)>12,y_new+1,y_new)   
   m_new=MONTH(&pardstaj)+11-m_st
   IF ALLTRIM(staj_in)='00-00-00'
      d_new=DAY(date_in) 
      m_new=MONTH(date_in)
   ELSE
       d_new=DAY(&pardstaj)+31-d_st        
       IF d_new>30
          d_new=d_new-30
          m_new=m_new+1
       ENDIF      
   ENDIF 
                                 
   IF m_new<13.AND.y_new=YEAR(&pardstaj)                  
      date_cx=STR(d_new,2)+'.'+STR(m_new,2)+'.'+STR(YEAR(&pardstaj),4)
      IF d_new>28.AND.m_new=2
         date_cx='01.03.'+STR(YEAR(&pardstaj),4)
      ENDIF
      IF d_new=31.AND.INLIST(m_new,4,6,9,11)
         date_cx='01.'+STR(m_new+1,2)+'.'+STR(YEAR(&pardstaj),4)
      ENDIF    
      REPLACE dPerSt WITH CTOD(date_cx)                     
    ENDIF
ENDIF   

**********************************************************************************************************
*             Расчет переходящего стажа по одному человеку
**********************************************************************************************************
PROCEDURE perstajone
PARAMETERS parStaj,pardstaj,parbase
IF EMPTY(parbase)
   SELECT people
ELSE  
   SELECT &parbase
ENDIF    
REPLACE dPerSt WITH CTOD('  /  /  ')
STORE 0 TO d_new,m_new,y_new, d_rest,y_rest
y_st=IIF(EMPTY(&parStaj),0,ROUND(VAL(LEFT(&parStaj,2)),0))
m_st=IIF(EMPTY(&parStaj),0,VAL(SUBSTR(&parStaj,4,2)))
d_st=IIF(EMPTY(&parStaj),0,VAL(SUBSTR(&parStaj,7,2)))   
y_new=YEAR(&pardstaj)
IF INLIST (y_st,4,9,14)
    DO CASE 
       CASE ALLTRIM(staj_in)='00-00-00'            
            *d_new=DAY(date_in-1) 
            *m_new=MONTH(date_in-1)
            
            d_new=DAY(date_in) 
            m_new=MONTH(date_in)

            m_rest=12-m_st
            y_new=y_new+IIF(MONTH(&pardstaj)+m_rest>12,1,0)         
        OTHERWISE 
            d_rest=IIF(d_st=0,0,31-d_st)
            d_new=IIF((d_rest+DAY(&pardstaj))<30,d_rest+DAY(&pardstaj),d_rest+DAY(&pardstaj)-30)
            m_rest=12-m_st-IIF(d_rest>0,1,0)
            m_new=m_rest+MONTH(&pardstaj)+IIF((d_rest+DAY(&pardstaj))>=30,1,0)
            y_new=y_new+IIF(MONTH(&pardstaj)+m_rest>12,1,0)
    ENDCASE    
    date_cx=IIF(y_new==YEAR(&pardstaj),STR(d_new,2)+'.'+STR(m_new,2)+'.'+STR(YEAR(&pardstaj),4),'  .  .    ') 
    REPLACE dPerSt WITH CTOD(date_cx) 
ENDIF 


************************************************************************************************************************
PROCEDURE formDeleteCard
SELECT people 
kodPeopOld=num
fSupl=CREATEOBJECT('FORMSUPL')
log_del=.F.
WITH fSupl 
     .Caption='Удаление личной карточки'
     DO addShape WITH 'fSupl',1,10,10,dHeight,400,8 
     DO adLabMy WITH 'fSupl',1,ALLTRIM(STR(people.num))+'  '+ALLTRIM(people.fio),.Shape1.Top+20,.Shape1.Left+10,.Shape1.Width-20,2   
     DO adLabMy WITH 'fSupl',2,'Для подтверждения намерений поставьте',.lab1.Top+.lab1.Height+10,.lab1.Left,.lab1.Width,2
     DO adLabMy WITH 'fSupl',3,'птичку в окошке "подтверждение намерений"',.lab2.Top+.lab2.Height,.lab1.Left,.lab1.Width,2
     .Shape1.Height=.lab1.Height*3+40
      DO adCheckBox WITH 'fSupl','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'log_del',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     
     DO addContLabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wУдалитьw')*2-20)/2,.check1.Top+.check1.Height+20,;
     RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','DO deleteCard'     
    
     DO addContLabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','fSupl.Release'
     .Height=.Shape1.Height+.check1.Height+.cont1.Height+60
     .Width=.Shape1.Width+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE deleteCard
IF !log_del
   RETURN
ENDIF
fsupl.Release
*SELECT datOtp
*DELETE FOR kodpeop=kodPeopOld
*SELECT datorder
*DELETE FOR kodpeop=kodPeopOld
*pathImage=pathDir+'datImage.dbf'
*IF !USED('datImage')
*   USE &pathImage IN 0
*ENDIF 
*SELECT datImage
*DELETE FOR kodpeop=kodPeopOld
*USE
IF !USED('datFam')
   USE datfam IN 0 ORDER 1
ENDIF
SELECT datFam
DELETE FOR kodpeop=kodPeopOld
USE
SELECT datJob
DELETE FOR kodpeop=kodPeopOld
SELECT people
DELETE
frmTop.Refresh 
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
DO changeRowGrdPers
*************************************************************************************************************************
*                  Процедура поиска личной карточки
*************************************************************************************************************************
PROCEDURE formForSearsh
newKp=0
fPoisk=CREATEOBJECT('FORMSUPL')
WITH fPoisk    
     .Caption='Поиск'   
     DO addShape WITH 'fPoisk',1,10,10,dHeight,450,8     
     .logExit=.T.  
     find_ch=''
     DO adLabMy WITH 'fpoisk',1,'код или ФИО сотрудника' ,.Shape1.Top+10,.Shape1.Left+10,.Shape1.Width-20,2
     DO addtxtboxmy WITH 'fpoisk',1,.Shape1.Left+10,.Shape1.Top+.lab1.Height+10,.Shape1.Width-20,.F.,'find_ch'
     .Shape1.Height=.lab1.Height+.txtBox1.Height+30
     .txtBox1.procForkeyPress='DO keyPressFind'
     DO addContLabel WITH 'fpoisk','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wОтменаw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wОтменаw'),dHeight+3,'Поиск','DO searshCard'
     DO addContLabel WITH 'fpoisk','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','Fpoisk.Release'          
     .Width=.Shape1.Width+20  
     .Height=.Shape1.Height+.cont1.Height+50 
ENDWITH     
DO pasteImage WITH 'fpoisk'
fpoisk.Show
*************************************************************************************************************************
*                Непосредственно поиск личной карточки
*************************************************************************************************************************
PROCEDURE SearshCard
IF EMPTY(find_ch)
   RETURN
ENDIF
find_ch=ALLTRIM(find_ch)        
SELECT people
oldrec=RECNO()
log_ord=SYS(21)
IF TYPE(find_ch)='N' 
   SET ORDER TO 1
   IF SEEK(VAL(find_ch))
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
      fPoisk.Release
   ELSE 
      find_ch=''
      fPoisk.Refresh
      SET ORDER TO &log_ord
      GO oldrec
      RETURN   
   ENDIF     
ENDIF
DO changeRowGrdPers
frmTop.grdPers.Column3.SetFocus
************************************************************************************************************************
PROCEDURE keyPressFind
DO CASE
   CASE LASTKEY()=27
        fpoisk.Release
   CASE LASTKEY()=13
        find_ch=fpoisk.TxtBox1.Value           
        DO searshCard  
ENDCASE 
**************************************************************************************************************************
PROCEDURE procReadJob
PARAMETERS par1
* par1 - новая запись
* parSovm - совместительство
IF USED('curDolPodr')
   SELECT curDolPodr
   USE
ENDIF
logNewRec=par1

=AFIELDS(arPeop,'people')
CREATE CURSOR curSuplPeop FROM ARRAY arPeop
SELECT curSuplPeop
INDEX ON fio TAG t1

SELECT datJob
ON ERROR DO erSup
oldRecJob=RECNO()
nidRec=0
COUNT TO maxJob
IF maxJob#0
   GO oldRecJob
ENDIF  
 
fSupl=CREATEOBJECT('FORMSUPL')
SELECT * FROM rasp INTO CURSOR curDolPodr READWRITE 
ALTER TABLE curDolPodr ADD COLUMN strVac C(15)
SELECT curDolPodr
REPLACE named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.namework,'') ALL 
INDEX ON nd TAG T1

SELECT datjob
GO oldRecJob

new_tr=IIF(par1,0,curJobSupl.tr)
new_podr=IIF(par1,0,curJobSupl.kp)
new_dolj=IIF(par1,0,curJobSupl.kd)
new_kse=IIF(par1,1.00,curJobSupl.kse)
newnordin=IIF(par1,'',curJobSupl.nordin)
newdordin=IIF(par1,CTOD('  .  .    '),curJobSupl.dordin)

newkat=IIF(par1,0,curJobSupl.kat)
new_kval=IIF(par1,people.kval,curJobSupl.kv)
newNid=IIF(par1,0,curJobSupl.nid)
newDateBeg=IIF(par1,CTOD('  .  .    '),curJobSupl.dateBeg)
newDateOut=IIF(par1,CTOD('  .  .    '),curJobSupl.dateOut)
newPkat=IIF(par1,0,curJobSupl.pkat)
newnordout=IIF(par1,'',curJobSupl.nordout)
newdordout=IIF(par1,CTOD('  .  .    '),curJobSupl.dordout)
newKdek=IIF(par1,0,curJobSupl.kdek)
newFioDek=IIF(par1,'',curJobSupl.fiodek)
kseRasp=0

IF !par1
   SELECT curDolPodr  
   SET FILTER TO kp=new_podr  
ENDIF 
str_type=IIF(SEEK(new_tr,'curSprType',1),curSprType.name,'')
str_podr=IIF(par1,'',IIF(SEEK(new_podr,'sprpodr',1),sprpodr.namework,''))
str_dolj=IIF(par1,'',IIF(SEEK(new_dolj,'sprdolj',1),sprdolj.namework,''))
str_kval=IIF(par1,'',IIF(SEEK(new_kval,'sprkval',1),sprkval.name,''))

SELECT datJob
oldJobRec=RECNO()  
oldOrd=SYS(21)
SET DELETED OFF 
SET ORDER TO 7 && nid
IF par1
   GO BOTTOM
   newNid=nid+1
ELSE 
   SEEK curJobSupl.nid
   nidRec=RECNO()
ENDIF    
SET DELETED ON 
SET ORDER TO &oldOrd
IF maxJob#0
   GO oldJobRec
ENDIF 
ON ERROR 
WITH fSupl       
     .Caption=ALLTRIM(people.fio)
     .procForClick='DO lostFocusDek'
     .procExit='DO exitFromInputJob'  
      DO adTboxAsCont WITH 'fSupl','txtPodr',10,10,RetTxtWidth('wдата увольнения (назначения)w'),dHeight,'подразделение',1,1
      DO addComboMy WITH 'fSupl',1,.txtPodr.Left+.txtPodr.Width-1,.txtPodr.Top,dheight,550,.T.,'str_podr',ALLTRIM('cursprpodr.namework'),6,.F.,'DO validPodrInJob',.F.,.T.      
      .comboBox1.DisplayCount=17    
      DO adTBoxAsCont WITH 'fSupl','txtDolj',.txtPodr.Left,.txtPodr.Top+.txtPodr.Height-1,.txtPodr.Width,dHeight,'должность',1,1
      DO addComboMy WITH 'fSupl',2,.comboBox1.Left,.txtDolj.Top,dheight,.comboBox1.Width,.T.,'str_dolj','ALLTRIM(curDolPodr.named)',6,.F.,'DO validDoljInJob',.F.,.T.  
      WITH .comboBox2         
           .DisplayCount=15
           .ColumnCount=3
           .ColumnWidths='0,50,500'
           .RowSource="curDolPodr.named,strVac,named"
      ENDWITH 
      
      DO adTBoxAsCont WITH 'fSupl','txtKse',.txtPodr.Left,.txtDolj.Top+.txtDolj.Height-1,.txtPodr.Width,dHeight,'объём',1,1
      DO addSpinnerMy WITH 'fSupl','spinKse',.txtKse.Left+.txtKse.Width-1,.txtKse.Top,dheight,RetTxtWidth('9999999999'),'new_kse',0.25,.F.,0,1.5
      
      DO adTBoxAsCont WITH 'fSupl','txtType',.spinKse.Left+.spinKse.Width-1,.txtKse.Top,RetTxtWidth('wтипw'),dHeight,'тип',2,1                                             
      DO addComboMy WITH 'fSupl',3,.txtType.Left+.txtType.Width-1,.txtType.Top,dheight,RetTxtWidth('совместительство'),.T.,'str_type','curSprType.name',6,.F.,'new_tr=curSprType.kod',.F.,.T.           
      
      DO adTBoxAsCont WITH 'fSupl','txtKval',.comboBox3.Left+.comboBox3.Width-1,.txtType.Top,RetTxtWidth('категорияw'),dHeight,'категория',1,1   
      DO addComboMy WITH 'fSupl',4,.txtKval.Left+.txtKval.Width-1,.txtKval.Top,dheight,.comboBox1.Width-.spinKse.Width-.txtType.Width-.comboBox3.Width-.txtKval.Width+4,.T.,'str_kval','curSprKval.name',6,.F.,'DO validKatInJob',.F.,.T.                                   
      
      
     DO adTboxAsCont WITH 'fSupl','txtDek',.txtPodr.Left,.txtKse.Top+.txtKse.Height-1,.txtPodr.Width,dHeight,'ФИО (кого замещают)',1,1      
     DO adTboxNew WITH 'fSupl','tBoxFioNew',.txtDek.Top,.comboBox1.Left,.comboBox1.Width-RetTxtWidth('w...')-2,dHeight,'newFioDek',.F.,.T.
     
     .tBoxFioNew.procforChange='DO changeFioZam'   
     DO adtboxnew WITH 'fSupl','boxFreeNew',.tBoxFioNew.Top,.tBoxFioNew.Left+.tBoxFioNew.Width-1,.comboBox1.Width-.tBoxFioNew.Width+1,dheight,'',.F.,.T.
     DO addButtonOne WITH 'fSupl','butKlntNew',.tBoxFioNew.Left+.tBoxFioNew.Width+1,.tBoxFioNew.Top+2,'','sbdn.ico','DO selectFioZam',.tBoxFioNew.Height-4,RetTxtWidth('w...')-1,'' 
     .butKlntNew.Enabled=.T.
    
      
      DO adTBoxAsCont WITH 'fSupl','txtDateIn',.txtPodr.Left,.txtDek.Top+.txtDek.Height-1,.txtPodr.Width,dHeight,'дата приема (назначения)',1,1   
      DO addtxtboxmy WITH 'fSupl',11,.comboBox1.Left,.txtDateIn.Top,.comboBox1.Width/5,.F.,'newDateBeg',0,.F.    
      
      DO adTBoxAsCont WITH 'fSupl','txtNprik',.txtBox11.Left+.txtBox11.Width-1,.txtDateIn.Top,.txtBox11.Width,dHeight,'№ прик.',1,1   
      DO addtxtboxmy WITH 'fSupl',12,.txtNprik.Left+.txtNPrik.Width-1,.txtDateIn.Top,.txtBox11.Width,.F.,'newnordin',0,.F.    
        
      DO adTBoxAsCont WITH 'fSupl','txtDprik',.txtBox12.Left+.txtBox12.Width-1,.txtDateIn.Top,.txtNprik.Width,dHeight,'Дата прик.',1,1   
      DO addtxtboxmy WITH 'fSupl',13,.txtDprik.Left+.txtDprik.Width-1,.txtDateIn.Top,.ComboBox1.Width-.txtBox11.Width*4+4,.F.,'newdordin',0,.F.    
         
      DO adTBoxAsCont WITH 'fSupl','txtDateOut',.txtPodr.Left,.txtDateIn.Top+.txtDateIn.Height-1,.txtPodr.Width,dHeight,'дата увольнения (перевода)',1,1   
      DO addtxtboxmy WITH 'fSupl',14,.comboBox1.Left,.txtDateOut.Top,.txtBox11.Width,.F.,'newDateOut',0,.F.    
      
      DO adTBoxAsCont WITH 'fSupl','txtNprikOut',.txtNprik.Left,.txtDateOut.Top,.txtNprik.Width,dHeight,'№ прик.',1,1   
      DO addtxtboxmy WITH 'fSupl',15,.txtBox12.Left,.txtDateOut.Top,.txtBox12.Width,.F.,'newnordout',0,.F.    
      
      DO adTBoxAsCont WITH 'fSupl','txtDprikOut',.txtDprik.Left,.txtDateOut.Top,.txtDprik.Width,dHeight,'Дата прик.',1,1   
      DO addtxtboxmy WITH 'fSupl',16,.txtBox13.Left,.txtDateOut.Top,.txtBox13.Width,.F.,'newdordout',0,.F.    
                                                         
      .Width=.txtPodr.Width+.comboBox1.Width+19 
                        
      DO addcontlabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,.txtDateOut.Top+.txtDateOut.Height+20,RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать','DO beforeWriteRecInJob'
      DO addcontlabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','DO exitFromInputJob'             
       
      .Height=.txtPodr.height*6+.cont1.Height+50  
      DO addListBoxMy WITH 'fSupl',2,.tBoxFioNew.Left,.tBoxFioNew.Top+dHeight-1,300,.combobox1.Width  
      WITH .listBox2
          .RowSource='curSuplpeop.fio'               
          .RowSourceType=2               
          .Visible=.F.        
          .procForDblClick='DO validFioDek'
          .procForLostFocus='DO lostFocusDek'
          .Height=.Parent.Height-.Top
     ENDWITH  
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
***********************************************************************************************************************
PROCEDURE exitFromInputJob
SELECT curSprType
SET FILTER TO 
SELECT people
fSupl.Release 

***********************************************************************************************************************
PROCEDURE validPodrInJob
new_podr=curSprPodr.kod
IF logNewRec.AND.new_podr#0.AND.new_dolj#0
   SELECT datjob
   LOCATE FOR kp=new_podr.AND.kd=new_dolj
ENDIF    
  
SELECT curDolPodr
SET FILTER TO kp=new_podr
SCAN ALL
     SELECT datJob
     SET ORDER TO 2
     ksesup=0
     SEEK STR(curDolPodr.kp,3)+STR(curDolPodr.kd,3)
     SCAN WHILE kp=curdolpodr.kp.AND.kd=curdolpodr.kd
          DO CASE
             CASE dekotp
             CASE !EMPTY(dateOut).AND.dateOut<=DATE()
             CASE dateBeg>DATE()
             OTHERWISE
                  ksesup=ksesup+kse
          ENDCASE        
     ENDSCAN   
     SELECT curDolPodr
     REPLACE strVac WITH IIF(kse-ksesup=0,'',LTRIM(STR(kse-ksesup,6,2)))
ENDSCAN
fSupl.ComboBox2.DisplayCount=IIF(RECCOUNT('curDolPodr')<15,RECCOUNT('curDolPodr'),15)
fSupl.ComboBox2.RowSourceType=6
fsupl.comboBox2.ProcForValid='DO validDoljInJob'
KEYBOARD '{TAB}'
SELECT datjob
**************************************************************************************************************************
PROCEDURE selectFioZam
SELECT curSuplPeop
ZAP
APPEND FROM people
WITH fSupl
     .listBox2.RowSource='curSuplPeop.fio'                     
      IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.  
        .listBox2.SetFocus            
     ENDIF 
ENDWITH 
**************************************************************************************************************************
PROCEDURE changeFioZam
WITH fSupl
     IF .listBox2.Visible=.F.
        .listBox2.Visible=.T.
     ENDIF    
ENDWITH 
Local lcValue,lcOption  
lcValue=fSupl.tBoxFioNew.Text 
SELECT curSuplPeop
ZAP
APPEND FROM people FOR LEFT(LOWER(fio),LEN(ALLTRIM(lcValue)))=LOWER(ALLTRIM(lcValue))
WITH fSupl.listBox2
     .RowSource='curSuplPeop.fio'                    
     .Visible=IIF(RECCOUNT('curSuplPeop')=0,.F.,.T.)      
ENDWITH 
**************************************************************************************************************************
PROCEDURE validFioDek
newFioDek=curSuplPeop.fio
newKDek=curSuplPeop.num
WITH fSupl
     .tBoxFioNew.ControlSource='newFioDek'
     .listBox2.Visible=.F.
     .tBoxFioNew.Refresh
     .Refresh
     .txtBox11.SetFocus
 ENDWITH 
**************************************************************************************************************************
PROCEDURE lostFocusDek
WITH fSupl
     ON ERROR DO erSup  
     .listBox2.Visible=.F.  
     ON ERROR  
ENDWITH
**************************************************************************************************************************
PROCEDURE beforeWriteRecInJob
fSupl.Visible=.F.
SELECT rasp
SET FILTER TO 
SELECT datjob
IF new_podr=0.OR.new_dolj=0.OR.new_kse=0
   RETURN
ENDIF
SELECT datJob
oldOrd=SYS(21)
SET ORDER TO 2
SEEK STR(new_podr,3)+STR(new_dolj,3)
realKse=0
SCAN WHILE kp=new_podr.AND.kd=new_dolj
     IF nid#newNid
        realKse=realKse+kse
     ENDIF 
ENDSCAN
SELECT datjob
SET ORDER TO &oldOrd
IF logNewRec
   APPEND BLANK  
   REPLACE nid WITH newNid  
ELSE 
   GO nidRec    
ENDIF  
REPLACE kodpeop WITH people.num,nidpeop WITH people.nid,kp WITH new_podr,kd WITH new_dolj,kse WITH new_kse,tr WITH new_tr,kat WITH newkat,;
        kv WITH new_kval,nordin WITH newnordin,dordin WITH newdordin,dateBeg WITH newDateBeg,nordOut WITH newnordOut,dordOut WITH newdordOut,dateOut WITH newDateOut,dekotp WITH people.dekotp,;
        fiodek WITH newFiodek,kDek WITH IIF(EMPTY(fiodek),0,newKdek),date_in WITH people.date_in,staj_in WITH people.staj_in,lkv WITH IIF(kv#0,.T.,.F.)
IF INLIST(tr,1,2,3,5)
   DO repnadjob   
   SELECT datjob
ENDIF         
datJobRec=RECNO()   
SELECT rasp
LOCATE FOR kp=new_podr.AND.kd=new_dolj
fSupl.Release
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus 
SELECT curSprType
SET FILTER TO    
SELECT datJob
GO datJobRec
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus    
**************************************************************************************************************************
PROCEDURE repnadjob
USE tarfond IN 0 ORDER 1
SELECT datjob
REPLACE date_in WITH people.date_in,staj_in WITH people.staj_in,pkont WITH IIF(tr=1,people.pkont,0),kat WITH IIF(SEEK(STR(datjob.kp,3)+STR(datjob.kd,3),'rasp',2),rasp.kat,kat)
DO CASE 
   CASE datjob.lkv.AND.people.kval#0
        REPLACE kv WITH people.kval,nprik WITH IIF(!EMPTY(people.nkval),'"'+ALLTRIM(people.nkval)+'"','')+IIF(!EMPTY(people.nordkval),' №'+ALLTRIM(people.nordkval),'')+IIF(!EMPTY(people.dkval),' от ' +DTOC(people.dkval),''),;
                pkat WITH IIF(SEEK(kv,'sprkval',1),sprkval.doplkat,0)
   OTHERWISE
        REPLACE kv WITH 0,nPrik WITH '',pkat WITH IIF(INLIST(datjob.kat,1,2,5,7),5,0)        
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
SELECT rasp
nOldRaspOrd=SYS(21)
SET ORDER TO 2
SEEK STR(datjob.kp,3)+STR(datjob.kd,3)
REPLACE datjob.pkf WITH rasp.pkf
SELECT tarfond
GO TOP
SCAN ALL    
     IF !EMPTY(plrep).AND.ltar
        repjob=ALLTRIM(plrep)
        repjob1='rasp.'+ALLTRIM(plrep)
        SELECT datjob 
        REPLACE &repjob WITH &repjob1          
     ENDIF  
     SELECT tarfond
ENDSCAN
SELECT rasp
SET ORDER TO &nOldRaspOrd
SELECT tarfond
USE
SELECT datjob
**************************************************************************************************************************
PROCEDURE validDoljInJob
new_dolj=curDolPodr.kd
newKat=curDolPodr.kat
kseRasp=curDolPodr.kse
IF logNewRec.AND.new_podr#0.AND.new_dolj#0
   SELECT datjob
   LOCATE FOR kp=new_podr.AND.kd=new_dolj  
ENDIF 
KEYBOARD '{TAB}'
**************************************************************************************************************************
PROCEDURE validKatInjob
new_kval=curSprKval.kod
newPkat=curSprKval.doplkat
KEYBOARD '{TAB}'
fSupl.Refresh
***********************************************************************************************************************
PROCEDURE formDelJob
fdel=CREATEOBJECT('FORMSUPL')
log_del=.F.
WITH fDel    
     .Caption='Удаление'    
     DO addShape WITH 'fDel',1,20,20,100,RetTxtWidth('wпоставьте птичку в окошке, расположенном нижеw'),8         
     DO adLabMy WITH 'fDel',1,'для подтверждения ваших намерений',fDel.Shape1.Top+10,fDel.Shape1.Left+5,.Shape1.Width-10,2 
     DO adLabMy WITH 'fDel',2,'поставьте птичку в окошке, расположенном ниже',.lab1.Top+.lab1.Height,fDel.Shape1.Left+5,.lab1.Width,2                                   
     DO adCheckBox WITH 'fdel','check1','подтверждение удаления',.lab2.Top+.lab2.Height+10,.Shape1.Left,150,dHeight,'log_del',0    
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     .Shape1.Height=.check1.Height+.lab1.Height*2+30
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('wУдалитьw')*2-20)/2,fdel.check1.Top+fdel.check1.Height+20,;
        RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','DO delRecFromJob'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+20,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'     
     .Width=.Shape1.Width+40   
     .Height=.Shape1.Height+.cont1.Height+60     
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
************************************************************************************************************************
PROCEDURE delRecFromJob
IF !log_del
   RETURN
ENDIF
fDel.Release
SELECT datJob
oldOrd=SYS(21)
SET ORDER TO 7
SEEK curJobSupl.nid
DELETE
SET ORDER TO &oldOrd
SELECT curJobSupl
DELETE
GO TOP
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus  
DO changeRowGrdPers
**********************************************************************************************************************
*                                                                  история
**********************************************************************************************************************
PROCEDURE procJobHistory
SELECT * FROM datJob WHERE datjob.kodpeop=people.num INTO CURSOR curHistory READWRITE
SELECT curHistory
INDEX ON dateBeg TAG T1 DESCENDING
GO TOP
fJob=CREATEOBJECT('FORMSUPL')
WITH fJob
     .procexit='DO exitHistoryJob'     
     .Caption='История назначений и перемещений - '+ALLTRIM(people.Fio)
     .AddObject('grdJob','gridMyNew')     
     WITH .grdJob          
          .ColumnCount=0
          DO addColumnToGrid WITH 'fJob.grdJob',9
          .Top=0
          .Width=frmTop.Width-frmTop.grdPers.Width
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
     FOR i=1 TO .grdJob.columnCount        
         .grdJob.Columns(i).DynamicBackColor='IIF(RECNO(fJob.grdJob.RecordSource)#fJob.grdJob.curRec,fJob.BackColor,dynBackColor)'
         .grdJob.Columns(i).DynamicForeColor='IIF(RECNO(fJob.grdJob.RecordSource)#fJob.grdJob.curRec,dForeColor,dynForeColor)'        
     ENDFOR                            
     DO gridSizeNew WITH 'fJob','grdJob','shapeingrid',.F.,.F.
     DO addcontlabel WITH 'fJob','butRead',10,.grdJob.Top+.grdJob.Height+20,RetTxtWidth('wвосстановитьw'),dHeight+3,'восстановить','DO formReadHistory'  
     DO addcontlabel WITH 'fJob','butDel',.butRead.Left+.butRead.Width+10,.butRead.Top,.butRead.Width,dHeight+3,'удаление','DO delFromHistory'  
     DO addcontlabel WITH 'fJob','butRet',.butDel.Left+.butDel.Width+10,.butRead.Top,.butRead.Width,.butRead.Height,'возврат','DO exitHistoryJob'   
        
     .Width=.grdJob.Width
     .butRead.Left=(.Width-.butRead.Width*3-20)/2    
     .butDel.Left=.butRead.Left+.butRead.Width+10     
     .butRet.Left=.butDel.Left+.butDel.Width+10     
     .Height=.grdJob.Height+.butRead.Height+40
ENDWITH
DO pasteImage WITH 'fJob'
fJob.Show   
*********************************************************************************************************************************
PROCEDURE exitHistoryJob
fJob.Release
frmTop.grdJob.Columns(frmTop.grdJob.ColumnCount).SetFocus  
DO changeRowGrdPers  
*********************************************************************************************************************************
PROCEDURE formReadHistory
IF EMPTY(curHistory.dateOut)
   RETURN
ENDIF
fdel=CREATEOBJECT('FORMSUPL')
log_del=.F.
WITH fDel   
     .Caption='Восстановить запись их архива'    
     DO addShape WITH 'fDel',1,20,20,100,RetTxtWidth('wпоставьте птичку в окошке, расположенном нижеw'),8         
     DO adLabMy WITH 'fDel',1,'для подтверждения ваших намерений',fDel.Shape1.Top+10,fDel.Shape1.Left+5,.Shape1.Width-10,2 
     DO adLabMy WITH 'fDel',2,'поставьте птичку в окошке, расположенном ниже',.lab1.Top+.lab1.Height,fDel.Shape1.Left+5,.lab1.Width,2                                   
     DO adCheckBox WITH 'fdel','check1','подтверждение удаления',.lab2.Top+.lab2.Height+10,.Shape1.Left,150,dHeight,'log_del',0    
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     .Shape1.Height=.check1.Height+.lab1.Height*2+30
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('wвосстановитьw')*2-20)/2,fdel.check1.Top+fdel.check1.Height+20,;
        RetTxtWidth('wвосстановитьw'),dHeight+3,'восстановить','DO restoreRecHistoryJob'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+20,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'отмена','fdel.Release'     
     .Width=.Shape1.Width+40   
     .Height=.Shape1.Height+.cont1.Height+60     
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
*********************************************************************************************************************************
PROCEDURE restoreRecHistoryJob
IF !log_del
   RETURN
ENDIF
fDel.Visible=.F.
fDel.Release
SELECT datJob
ordOld=SYS(21)
SET ORDER TO 7
SEEK curHistory.nid         
REPLACE dateOut WITH CTOD('  .  .    '),nordout WITH '',dordout WITH CTOD('  .  .    '),nidout WITH 0,kordout WITH 0
SET ORDER TO &ordOld
SELECT curHistory
DELETE 
fJob.Refresh  
fJob.grdJob.Columns(fJob.grdJob.ColumnCount).SetFocus  

*********************************************************************************************************************************
PROCEDURE delFromHistory
SELECT curHistory
oldRec=RECNO()
COUNT TO maxHistory
IF maxHistory=0
   RETURN 
ENDIF 
GO oldRec
logJob=IIF(EMPTY(dateout),.T.,.F.)
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl    
     .Caption='Удаление записи'
     .Width=400   
     IF !logJob  
        DO adLabMy WITH 'fSupl',1,'Удалить выбранную запись?',20,0,.Width,2 
        DO addContLabel WITH 'fSupl','cont1',(.Width-RetTxtWidth('wУдалитьw')*2-20)/2,.lab1.Top+.lab1.Height+10,RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','DO procDelRecFromHistory WITH .T.'
        DO addContLabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','DO procDelRecFromHistory' 
        .Height=.lab1.Height+.cont1.Height+60
     ELSE 
        DO addShape WITH 'fSupl',1,10,10,100,400,8         
        DO adLabMy WITH 'fSupl',1,'Внимание!',.Shape1.Top+20,.Shape1.Left,.Shape1.Width,2 
        DO adLabMy WITH 'fSupl',2,'В настоящее время выбранная должность',.lab1.Top+.lab1.Height,.Shape1.Left,.Shape1.Width,2 
        DO adLabMy WITH 'fSupl',3,'закреплена за сотрудником!',.lab2.Top+.lab2.Height,.Shape1.Left,.Shape1.Width,2 
        .Shape1.Height=.lab1.Height*3+40
        DO addContLabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wУдалитьw')*2-20)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','DO procDelRecFromHistory WITH .T.'
        DO addContLabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','DO procDelRecFromHistory' 
        .Height=.Shape1.Height+.cont1.Height+60
        .Width=.Shape1.Width+20
     ENDIF   
     DO pasteImage WITH 'fSupl'    
     .Show 
ENDWITH 
********************************************************************************************************************************
PROCEDURE procDelRecFromHistory
PARAMETERS par1
fSupl.Visible=.F.
fSupl.Release
IF par1
   SELECT datJob
   ordOld=SYS(21)
   SET ORDER TO 7
   SEEK curHistory.nid         
   DELETE 
   SET ORDER TO &ordOld
   SELECT curHistory
   DELETE 
   fJob.Refresh  
   fJob.grdJob.Columns(fJob.grdJob.ColumnCount).SetFocus  
ENDIF    
*-------------------------------------------------------------------------------------------------------------------------
*                    Процедуры для справочнго материала
*-------------------------------------------------------------------------------------------------------------------------
*********************************************************************************************************************************************************
*                                                   Справочник подразделений
*********************************************************************************************************************************************************
PROCEDURE procpodr
SELECT sprpodr
oldOrdPodr=SYS(21)
SET ORDER TO 3
GO TOP
fdolj=CREATEOBJECT('Formspr')
namenew=''
nameOrdNew=''
nameKNew=''
namernew=''
primnew=''
namernewold=''
log_ap=.F.
WITH fdolj  
     .Caption='Справочник подразделений'  
     .ProcExit='DO exitFromProcPodr'  
     DO addButtonOne WITH 'fDolj','menuCont1',10,5,'редакция','pencil.ico',"Do readspr WITH 'fdolj','Do readpodr WITH .F.'",39,RetTxtWidth('удаление')+44,'редакция'  
     DO addButtonOne WITH 'fDolj','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'возврат','undo.ico','DO exitFromProcPodr',39,.menucont1.Width,'возврат'             
     DO addmenureadspr WITH 'fdolj','DO writePodr WITH .T.','DO writePodr WITH .F.' 
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Height=.Parent.Height-.Parent.menucont1.Height-5                 
          .RecordSourceType=1     
          .RecordSource='sprpodr'
           DO addColumnToGrid WITH 'fDolj.fGrid',6
          .Column1.ControlSource='sprpodr.kod'
          .Column2.ControlSource='sprpodr.np'
          .Column3.ControlSource='" "+sprpodr.namework'
          .Column4.ControlSource='" "+sprpodr.nameord' 
          .Column5.ControlSource='" "+sprpodr.namek' 
                       
          .Column1.Width=RettxtWidth(' 1234 ')
          .Column2.Width=.Column1.Width                  
          .Column3.Width=(.Width-.Column1.Width-.Column2.Width)/3
          .Column4.Width=.Column3.Width
          .Column5.Width=.Width-.column1.width-.Column2.Width-.Column3.Width-.Column4.Width-SYSMETRIC(5)-13-.ColumnCount    
          .Columns(.ColumnCount).Width=0               
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='№'
          .Column3.Header1.Caption='Наименование'
          .Column4.Header1.Caption='Наименование в приказах'                       
          .Column5.Header1.Caption='Работает где'                       
          .Column1.Alignment=1
          .Column2.Alignment=1           
          .Column3.Alignment=0                  
          .Column4.Alignment=0
          .Column5.Alignment=0
          .colNesInf=2      
          .Visible=.T.         
     ENDWITH
     DO gridSizeNew WITH 'fdolj','fGrid','shapeingrid'     
     .fGrid.Column1.Text1.ToolTipText='xccvzsvzxcv'   
     DO addtxtboxmy WITH 'fdolj',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
     .txtbox1.Enabled=.F.
     DO addtxtboxmy WITH 'fdolj',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,1
     .txtbox2.Enabled=.F.
     DO addtxtboxmy WITH 'fdolj',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fdolj',4,1,1,.fGrid.Column4.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fdolj',5,1,1,.fGrid.Column5.Width+2,.F.,.F.,0     
     .SetAll('Visible',.F.,'MyTxtBox')
     DO addcontmy WITH 'fdolj','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont1','sprpodr',1"
     DO addcontmy WITH 'fdolj','cont2',.cont1.Left+.fGrid.Column1.Width+2,.fGrid.Top+2,.fGrid.Column2.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprpodr',3"
     .cont2.SpecialEffect=1   
     DO addcontmy WITH 'fdolj','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprpodr',2" 
     DO addcontmy WITH 'fdolj','cont4',.cont3.Left+.fGrid.Column3.Width+1,.fGrid.Top+2,.fGrid.Column4.Width-4,.fGrid.HeaderHeight-3,'' 
     SELECT sprpodr  
     GO TOP 
     .Show
ENDWITH
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readpodr
PARAMETERS parlog
SELECT sprpodr
log_ap=.F.
WITH fDolj
     IF parlog
        .fGrid.GridMyAppendBlank(2,'kod','name')   
        log_ap=.T.       
     ENDIF
     .fGrid.columns(.fGrid.columnCount).SetFocus
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1) 
     .nrec=RECNO()
     nameNew=sprpodr.namework     
     nameordNew=sprpodr.nameord
     nameKNew=sprpodr.namek
     .txtBox1.Left=.fGrid.Left+10
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .txtBox3.Left=.txtBox2.Left+.txtBox2.Width-1   
     .txtBox4.Left=.txtBox3.Left+.txtBox3.Width-1   
     .txtBox5.Left=.txtBox4.Left+.txtBox4.Width-1   
     .txtbox1.ControlSource='sprpodr.kod'
     .txtbox2.ControlSource='sprpodr.np'
     .txtbox3.ControlSource='nameNew'
     .txtbox4.ControlSource='nameordNew' 
     .txtbox5.ControlSource='nameKNew' 
        
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .SetAll('Visible',.T.,'MyTxtBox')
     .fGrid.Enabled=.F.
     .Refresh
     .txtbox3.SetFocus
      
ENDWITH      
IF parlog
   KEYBOARD '{TAB}'
ENDIF   
************************************************************************************************************************
PROCEDURE writepodr
PARAMETERS par_log
WITH fDolj
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     SELECT sprpodr
     IF par_log  
        REPLACE namework WITH namenew,nameord WITH nameordNew,namek WITH nameKNew
     ELSE   
        IF log_ap
           DELETE
        ENDIF
     ENDIF    
     .SetAll('Visible',.F.,'MyTxtBox')
     .fGrid.Enabled=.T.
     SELECT sprpodr
     .fGrid.GridUpdate
     GO .nrec
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     GO .nrec
ENDWITH      
*************************************************************************************************************************
PROCEDURE delpodr
fdolj.Setall('BorderWidth',0,'Mycontmenu')
IF SEEK(STR(sprpodr.kod,3),'rasp',1) 
   fdolj.fGrid.GridNoDelRec   
ELSE 
  fdolj.fGrid.GridDelRec('fdolj.fGrid','sprpodr') 
ENDIF   
**************************************************************************************************************************
PROCEDURE exitFromProcPodr
SELECT sprpodr
SET ORDER TO &oldOrdPodr
fDolj.Visible=.F.
fDolj.Release
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник должностей
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE procdolj
IF !USED('sprpadej')
   USE sprpadej IN 0
ENDIF
kodnew=0
namenew=''
strkat=''
katnew=0
log_ap=.F.
SELECT sprdolj
SET ORDER TO 1
GO TOP
fdolj=CREATEOBJECT('Formspr')
WITH fdolj    
     .Caption='Справочник должностей'  
     .ProcExit='DO procOutSprDolj'     
     DO addButtonOne WITH 'fDolj','menuCont1',10,5,'новая','pencila.ico','DO formReadDolj WITH .T.',39,RetTxtWidth('удаление')+44,'новая'   
     DO addButtonOne WITH 'fDolj','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico','DO formReadDolj WITH .F.',39,RetTxtWidth('удаление')+44,'редакция'   
     DO addButtonOne WITH 'fDolj','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO deldolj',39,.menucont2.Width,'удаление'   
     DO addButtonOne WITH 'fDolj','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico',"DO printreport WITH 'repdolj','справочник должностей','sprdolj'",39,.menucont2.Width,'печать'   
     DO addButtonOne WITH 'fDolj','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'возврат','undo.ico','DO procOutSprDolj',39,.menucont2.Width,'возврат'     
     WITH .fGrid
          .Top=.Parent.menucont2.Top+.Parent.menucont2.Height+5
          .Height=.Parent.Height-.Parent.menucont2.Height-5       
          .RecordSourceType=1     
          .RecordSource='sprdolj'
          DO addColumnToGrid WITH 'fDolj.fGrid',6 
          .Column1.ControlSource='sprdolj.kod'
          .Column2.ControlSource='" "+sprdolj.name'
          .Column3.ControlSource='" "+sprdolj.namer'     
          .Column4.ControlSource='" "+sprdolj.named'     
          .Column5.ControlSource='" "+sprdolj.namet'    
         
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование'
          .Column3.Header1.Caption='Кого'    
          .Column4.Header1.Caption='Кому'    
          .Column5.Header1.Caption='Кем'    
          .Column1.Width=RettxtWidth(' 1234 ')     
          .Column5.Width=(.Width-.column1.width-SYSMETRIC(5)-13-.ColumnCount)/4
          .Column4.Width=.Column5.Width
          .Column3.Width=.Column5.Width   
          .Columns(.ColumnCount).Width=0
          .Column2.Width=.Width-.column1.Width-.column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount    
          .Column1.Movable=.F. 
          .Column1.Alignment=1           
          .Column2.Alignment=0
          .colNesInf=2      
          .SetAll('BOUND',.F.,'Column')  
          .Visible=.T.         
     ENDWITH
     DO gridSizeNew WITH 'fdolj','fGrid','shapeingrid'   
     DO addcontmy WITH 'fdolj','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont1','sprdolj',1,4"
     .cont1.SpecialEffect=1   
     DO addcontmy WITH 'fdolj','cont2',.cont1.Left+.fGrid.Column1.Width+2,.fGrid.Top+2,.fGrid.Column2.Width-4,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprdolj',2,5"
     DO addcontmy WITH 'fdolj','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-4,.fGrid.HeaderHeight-3,'' 
     DO addcontmy WITH 'fdolj','cont4',.cont3.Left+.fGrid.Column3.Width+1,.fGrid.Top+2,.fGrid.Column4.Width-4,.fGrid.HeaderHeight-3,'' 
     DO addcontmy WITH 'fdolj','cont5',.cont4.Left+.fGrid.Column4.Width+1,.fGrid.Top+2,.fGrid.Column5.Width-4,.fGrid.HeaderHeight-3,'' 
     SELECT sprdolj  
ENDWITH
fdolj.Show
*************************************************************************************************************************
PROCEDURE formReadDolj
PARAMETERS par1
logAp=IIF(par1,.T.,.F.)

newName=IIF(par1,SPACE(150),sprdolj.name)
newNameW=IIF(par1,SPACE(150),sprdolj.namework)
newNameR=IIF(par1,SPACE(150),sprdolj.namer)
newNameD=IIF(par1,SPACE(150),sprdolj.named)
newNameT=IIF(par1,SPACE(150),sprdolj.namet)
newNameV=IIF(par1,SPACE(150),sprdolj.namev)

newNamem=IIF(par1,SPACE(150),sprdolj.namem)
newNameRm=IIF(par1,SPACE(150),sprdolj.namerm)
newNameDm=IIF(par1,SPACE(150),sprdolj.namedm)
newNameTm=IIF(par1,SPACE(150),sprdolj.nametm)
newNameVm=IIF(par1,SPACE(150),sprdolj.namevm)
newLogSex=IIF(par1,.F.,sprdolj.logSex)

fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Редактирование'     
     DO addShape WITH 'fSupl',1,10,10,20,500,8 
     DO adTBoxAsCont WITH 'fsupl','txtName',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('Wнаименование - рабочее  W'),dHeight,'наименование - полное',1,1 
     DO adtboxnew WITH 'fSupl','boxName',.txtName.Top,.txtName.Left+.txtName.Width-1,500,dheight,'newName',.F.,.T.,0,.F.,'DO validDoljpadej'  
     
     DO adTBoxAsCont WITH 'fsupl','txtNameW',.txtName.Left,.txtName.Top+.txtName.Height-1,.txtName.Width,dHeight,'наименование - рабочее',1,1 
     DO adtboxnew WITH 'fSupl','boxNameW',.txtNameW.Top,.txtName.Left+.txtName.Width-1,500,dheight,'newNameW',.F.,.T.,0,.F.
        
     DO adTBoxAsCont WITH 'fsupl','txtNameR',.txtName.Left,.txtNameW.Top+.txtNameW.Height-1,.txtName.Width,dHeight,'перевести кого',1,1
     DO adtboxnew WITH 'fSupl','boxNameR',.txtNameR.Top,.boxName.Left,.boxName.Width,dheight,'newNameR',.F.,.T.,0,.F.  

     DO adTBoxAsCont WITH 'fsupl','txtNameD',.txtName.Left,.txtNamer.Top+.txtNamer.Height-1,.txtName.Width,dHeight,'предоставить кому',1,1
     DO adtboxnew WITH 'fSupl','boxNameD',.txtNameD.Top,.boxName.Left,.boxName.Width,dheight,'newNameD',.F.,.T.,0,.F.  
     
     DO adTBoxAsCont WITH 'fsupl','txtNameT',.txtName.Left,.txtNameD.Top+.txtNameD.Height-1,.txtName.Width,dHeight,'перевести кем',1,1
     DO adtboxnew WITH 'fSupl','boxNameT',.txtNameT.Top,.boxName.Left,.boxName.Width,dheight,'newNameT',.F.,.T.,0,.F.          
     
     DO adTBoxAsCont WITH 'fsupl','txtNameV',.txtName.Left,.txtNameT.Top+.txtNameT.Height-1,.txtName.Width,dHeight,'на должность кого',1,1
     DO adtboxnew WITH 'fSupl','boxNameМ',.txtNameV.Top,.boxName.Left,.boxName.Width,dheight,'newNameV',.F.,.T.,0,.F.          
     
     
     .Shape1.Width=.txtName.Width+.boxName.Width+20
     .Shape1.Height=.txtName.Height*6+20
     DO adCheckBox WITH 'fSupl','checkSex','двойное наименование ',.Shape1.Top+.Shape1.Height+10,.Shape1.Left,250,dHeight,'newLogSex',0,.T.  
     .checkSex.Left=.Shape1.Left+(.Shape1.Width-.checkSex.Width)/2
     
     DO addShape WITH 'fSupl',2,.Shape1.Left,.checkSex.Top+.checkSex.Height+10,100,.Shape1.Width,8  
     
     DO adTBoxAsCont WITH 'fsupl','txtNamem',.txtName.Left,.Shape2.Top+10,.txtName.Width,dHeight,'наименование',1,1 
     DO adtboxnew WITH 'fSupl','boxNamem',.txtNamem.Top,.boxName.Left,.boxName.Width,dheight,'newNamem',.F.,.T.,0,.F.
        
     DO adTBoxAsCont WITH 'fsupl','txtNameRm',.txtName.Left,.txtNamem.Top+.txtNamem.Height-1,.txtName.Width,dHeight,'перевести кого',1,1
     DO adtboxnew WITH 'fSupl','boxNameRm',.txtNameRm.Top,.boxName.Left,.boxName.Width,dheight,'newNameRm',.F.,.T.,0,.F.  

     DO adTBoxAsCont WITH 'fsupl','txtNameDm',.txtName.Left,.txtNameRm.Top+.txtNameRm.Height-1,.txtName.Width,dHeight,'предоставить кому',1,1
     DO adtboxnew WITH 'fSupl','boxNameDm',.txtNameDm.Top,.boxName.Left,.boxName.Width,dheight,'newNameDm',.F.,.T.,0,.F.  
     
     DO adTBoxAsCont WITH 'fsupl','txtNameTm',.txtName.Left,.txtNameDm.Top+.txtNameDm.Height-1,.txtName.Width,dHeight,'перевести кем',1,1
     DO adtboxnew WITH 'fSupl','boxNameTm',.txtNameTm.Top,.boxName.Left,.boxName.Width,dheight,'newNameTm',.F.,.T.,0,.F.          
     
     DO adTBoxAsCont WITH 'fsupl','txtNameVm',.txtName.Left,.txtNameTm.Top+.txtNameTm.Height-1,.txtName.Width,dHeight,'на должность кого',1,1
     DO adtboxnew WITH 'fSupl','boxNameVm',.txtNameVm.Top,.boxName.Left,.boxName.Width,dheight,'newNameVm',.F.,.T.,0,.F.       
     
     .Shape2.Width=.Shape1.Width
     .Shape2.Height=.txtName.Height*5+20
     
     DO addContLabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wзаписатьw')*2-20)/2,.Shape2.Top+.Shape2.Height+20,;
        RetTxtWidth('wзаписатьw'),dHeight+3,'записать','DO writeRecSprDolj'
     DO addContLabel WITH 'fSupl','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fSupl.Release'          
     .Width=.Shape1.Width+20     
     .Height=.Shape1.Height+.Shape2.Height+.checkSex.Height+.cont1.Height+90
     DO pasteImage WITH 'fSupl'      
ENDWITH
fSupl.Show

*************************************************************************************************************************
PROCEDURE validDoljPadej
strPadej=ALLTRIM(newName)
atPadej=AT(' ',strPadej)
atPadej1=AT('-',strPadej)
repRp=''
repDp=''
repTp=''
DO CASE
   CASE INLIST(RIGHT(strPadej,1),'г','з','ч','р','т','к')
        repRp=strPadej+'а' 
        repDp=strPadej+'у' 
        repTp=strPadej+'ом' 
   CASE RIGHT(strPadej,2)='а)'
        repRp=LEFT(strPadej,LEN(strPadej)-2)+'у)' 
        repDp=LEFT(strPadej,LEN(strPadej)-2)+'е)' 
        repTp=LEFT(strPadej,LEN(strPadej)-2)+'ой)' 
   CASE INLIST(RIGHT(strPadej,1),'а')
        repRp=LEFT(strPadej,LEN(strPadej)-1)+'у' 
        repDp=LEFT(strPadej,LEN(strPadej)-1)+'е' 
        repTp=LEFT(strPadej,LEN(strPadej)-1)+'ой'      
             
*   CASE RIGHT(strPadej,1)='а'     
*        repRp=LEFT(strPadej,LEN(strPadej)-1)+'у' 
*        repDp=LEFT(strPadej,LEN(strPadej)-1)+'е' 
*        repTpp=LEFT(strPadej,LEN(strPadej)-1)+'ой'
*   CASE RIGHT(strPadej,1)='а)'     
*        repRp=LEFT(strPadej,LEN(strPadej)-2)+'у)' 
*        repDp=LEFT(strPadej,LEN(strPadej)-2)+'е)' 
*        repTpp=LEFT(strPadej,LEN(strPadej)-2)+'ой)'     
   OTHERWISE 
        repRp=strPadej
        repDp=strPadej
        repTp=strPadej
ENDCASE

IF  ' '$strPadej
    stringSup=ALLTRIM(LEFT(strPadej,AT(' ',strPadej)-1))
    repRp1=LEFT(repRp,AT(' ',repRp)-1)
    repDp1=LEFT(repDp,AT(' ',repDp)-1)
    repTp1=LEFT(repTp,AT(' ',repTp)-1)
   
    
    repRpRight=SUBSTR(repRp,AT(' ',strPadej))
    repDpRight=SUBSTR(repDp,AT(' ',strPadej))
    repTpRight=SUBSTR(repTp,AT(' ',strPadej))
       
    DO CASE
        CASE INLIST(RIGHT(stringSup,1),'г','з','ч','р','т','к')
             repRp=stringSup+'а'+repRpRight 
             repDp=stringSup+'у'+repDpRight  
             repTp=stringSup+'ом'++repTpRight 
         CASE RIGHT(stringSup,1)='ь'
             repRp=LEFT(stringSup,LEN(stringSup)-1)+'я'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-1)+'ю'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-1)+'ем'+repTpRight        
        CASE RIGHT(stringSup,2)='ый'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'ого'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'ому'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'ым'+repTpRight  
        CASE RIGHT(stringSup,3)='щий'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'его'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'ему'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'им'+repTpRight      
        CASE RIGHT(stringSup,3)='кий'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'ого'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'ому'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'им'+repTpRight           
        CASE RIGHT(stringSup,2)='ая'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'ую'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'ой'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'ой'+repTpRight       
        CASE RIGHT(stringSup,2)='а)'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'у)'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'е)'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'ой)'+repTpRight  
        CASE INLIST(RIGHT(stringSup,1),'а')
             repRp=LEFT(stringSup,LEN(stringSup)-1)+'у'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-1)+'е'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-1)+'ой'+repTpRight 
        OTHERWISE     
     ENDCASE
     IF OCCURS(' ',strPadej)=2
        stringSup1=ALLTRIM(SUBSTR(strPadej,AT(' ',strPadej)+1))
        stringSup=ALLTRIM(SUBSTR(strPadej,AT(' ',strPadej)+1,AT(' ',stringSup1)))
        repRpLeft=LEFT(repRp,AT(' ',repRp)-1)
        repDpLeft=LEFT(repDp,AT(' ',repDp)-1)
        repTpLeft=LEFT(repTp,AT(' ',repTp)-1)
        repRpRight=SUBSTR(repRp,RAT(' ',repRp))
        repDpRight=SUBSTR(repDp,RAT(' ',repDp))
        repTpRight=SUBSTR(repTp,RAT(' ',repTp))
    DO CASE
        CASE INLIST(RIGHT(stringSup,1),'г','з','ч','р','т','к')
             repRp=stringSup+'а'+repRpRight 
             repDp=stringSup+'у'+repDpRight  
             repTp=stringSup+'ом'++repTpRight  
        CASE RIGHT(stringSup,2)='ый'
             repRp=repRpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ого'+repRpRight  
             repDp=repDpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ому'+repDpRight  
             repTp=repTpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ым'+repTpRight  
        CASE RIGHT(stringSup,3)='кий'
             repRp=repRpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ого'+repRpRight  
             repDp=repDpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ому'+repDpRight  
             repTp=repTpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'им'+repTpRight  
        CASE RIGHT(stringSup,3)='щий'
             repRp=repRpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'его'+repRpRight  
             repDp=repDpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ему'+repDpRight  
             repTp=repTpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'им'+repTpRight            
        CASE RIGHT(stringSup,2)='ая'
             repRp=repRpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ую'+repRpRight  
             repDp=repDpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ой'+repDpRight  
             repTp=repTpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ой'+repTpRight       
        CASE RIGHT(stringSup,2)='а)'
             repRp=repRpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'у)'+repRpRight  
             repDp=repDpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'е)'+repDpRight  
             repTp=repTpLeft+' '+LEFT(stringSup,LEN(stringSup)-2)+'ой)'+repTpRight  
        CASE INLIST(RIGHT(stringSup,1),'а')
             repRp=repRpLeft+' '+LEFT(stringSup,LEN(stringSup)-1)+'у'+repRpRight  
             repDp=repDpLeft+' '+LEFT(stringSup,LEN(stringSup)-1)+'е'+repDpRight  
             repTp=repTpLeft+' '+LEFT(stringSup,LEN(stringSup)-1)+'ой'+repTpRight 
        OTHERWISE     
     ENDCASE
     ENDIF
ENDIF 


IF  '-'$strPadej
    stringSup=ALLTRIM(LEFT(strPadej,AT('-',strPadej)-1))
    repRp1=LEFT(repRp,AT('-',repRp)-1)
    repDp1=LEFT(repDp,AT('-',repDp)-1)
    repTp1=LEFT(repTp,AT('-',repTp)-1)
   
    
    repRpRight=SUBSTR(repRp,AT('-',strPadej))
    repDpRight=SUBSTR(repDp,AT('-',strPadej))
    repTpRight=SUBSTR(repTp,AT('-',strPadej))
    
    DO CASE
        CASE INLIST(RIGHT(stringSup,1),'г','з','ч','р','т','к')
             repRp=stringSup+'а'+repRpRight 
             repDp=stringSup+'у'+repDpRight  
             repTp=stringSup+'ом'++repTpRight  
        CASE RIGHT(stringSup,2)='ый'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'ого'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'ому'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'ым'+repTpRight  
        CASE RIGHT(stringSup,1)='ь'
             repRp=LEFT(stringSup,LEN(stringSup)-1)+'я'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-1)+'ю'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-1)+'ем'+repTpRight       
        CASE RIGHT(stringSup,2)='а)'
             repRp=LEFT(stringSup,LEN(stringSup)-2)+'у)'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-2)+'е)'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-2)+'ой)'+repTpRight  
        CASE INLIST(RIGHT(stringSup,1),'а')
             repRp=LEFT(stringSup,LEN(stringSup)-1)+'у'+repRpRight  
             repDp=LEFT(stringSup,LEN(stringSup)-1)+'е'+repDpRight  
             repTp=LEFT(stringSup,LEN(stringSup)-1)+'ой'+repTpRight 
        OTHERWISE     
     ENDCASE
ENDIF 

IF EMPTY(newNameR)
   newNameR=repRp  
ENDIF
IF EMPTY(newNameD)
   newNameD=repDp
ENDIF 
IF EMPTY(newNameT)
   newNameT=repTp
ENDIF 
newNameW=newName
SELECT sprdolj
fSupl.Refresh
*************************************************************************************************************************
PROCEDURE writeRecSprDolj
IF logAp
   SELECT sprDolj
   oldRec=RECNO()
   oldOrd=SYS(21)
   SET ORDER TO 1
   GO BOTTOM
   newKod=kod+1
   APPEND BLANK
   REPLACE kod WITH newKod
ENDIF
REPLACE name WITH newName,namer WITH newNameR,named WITH newNameD,namet WITH newNameT,namework WITH newNameW,logSex WITH newLogSex,;
        namem WITH newNamem,namerm WITH newNameRm,namedm WITH newNameDm,nametm WITH newNameTm,namev WITH newNameV,namevm WITH newNameVm
fSupl.Release
fDolj.Refresh

*************************************************************************************************************************
PROCEDURE procvalidkat
SELECT sprdolj
katNew=cursprkat.kod
KEYBOARD '{TAB}'    
************************************************************************************************************************
PROCEDURE procgotkat
SELECT cursprkat
LOCATE FOR kod=sprkat->kod
nrec=RECNO()
GO TOP 
COUNT WHILE RECNO()#nrec TO varnrec
fdolj.combobox3.DisplayCount=MAX(fdolj.fGrid.RelativeRow,fdolj.fGrid.RowsGrid-fdolj.fGrid.RelativeRow)
fdolj.combobox3.DisplayCount=MIN(fdolj.combobox3.DisplayCount,RECCOUNT())
SELECT sprdolj
************************************************************************************************************************
PROCEDURE writedolj
PARAMETERS par_log
WITH fDolj
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     SELECT sprdolj
     IF par_log  
        REPLACE name WITH namenew,kat WITH katNew
     ELSE   
        IF log_ap
           DELETE
        ENDIF
     ENDIF    
     .SetAll('Visible',.F.,'MyTxtBox')
     .comboBox3.Visible=.F.
     .fGrid.Enabled=.T.
     SELECT sprdolj
     .fGrid.GridUpdate
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     GO .nrec
ENDWITH  
*************************************************************************************************************************
PROCEDURE deldolj
SELECT sprdolj
SELECT rasp
LOCATE FOR kd=sprdolj.kod
IF FOUND()
   SELECT sprdolj
   fdolj.fGrid.GridNoDelRec   
ELSE 
  SELECT sprdolj
  fdolj.fGrid.GridDelRec('fdolj.fGrid','sprdolj') 
ENDIF   
**************************************************************************************************************************
PROCEDURE procOutSprDolj
fDolj.Release
SELECT * FROM sprdolj INTO CURSOR curSprDolj READWRITE ORDER BY name
SELECT sprdolj
SET RELATION TO 
SET ORDER TO 1
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник категорий персонала
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE prockat
SELECT sprkat
SET ORDER TO 1
GO TOP
fkat=CREATEOBJECT('Formspr')
WITH fkat   
     .Caption='Справочник производственных категорий персонала' 
     .ProcExit='fkat.fGrid.GridReturn'  
     DO addButtonOne WITH 'fKat','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fKat','Do readkat WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fKat','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fKat','Do readkat WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fKat','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delkat',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fKat','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','fKat.fGrid.GridReturn',39,.menucont1.Width,'возврат'                       
     
     DO addmenureadspr WITH 'fkat',"DO writeSprNew WITH 'fkat','fkat.fGrid','sprkat'","DO exitWriteSpr WITH 'fkat','fkat.fGrid'"
     
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5                    
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fKat.fGrid',5
          .RecordSource='sprkat'
          .Column1.ControlSource='sprkat.kod'
          .Column2.ControlSource='" "+sprkat.name'    
          .Column3.ControlSource='" "+sprkat.namefull'
          .Column4.ControlSource='" "+sprkat.namefull1'
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование'
          .Column3.Header1.Caption='Для ведомостей'
          .Column4.Header1.Caption='Для штатного'     
          .Column1.Width=RettxtWidth(' 1234 ')
          .Column2.Width=(.Width-.column1.Width)/3 
          .Column3.Width=.Column2.Width   
          .Column4.Width=.Width-.column1.Width-.column2.Width-.Column3.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0  
          .Column1.Alignment=1    
          .colNesInf=2   
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')         
     ENDWITH  
     DO gridSizeNew WITH 'fkat','fGrid','shapeingrid'   
     DO addtxtboxmy WITH 'fkat',1,1,1,fkat.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fkat',2,1,1,fkat.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fkat',3,1,1,fkat.fGrid.Column3.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fkat',4,1,1,fkat.fGrid.Column4.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
     DO addcontmy WITH 'fkat','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,''
     .cont1.SpecialEffect=1          
     DO addcontmy WITH 'fkat','cont2',.cont1.Left+.fGrid.Column1.Width+1,.fGrid.Top+2,.fGrid.Column2.Width-3,.fGrid.HeaderHeight-3,''
     DO addcontmy WITH 'fkat','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-3,.fGrid.HeaderHeight-3,''
     DO addcontmy WITH 'fkat','cont4',.cont3.Left+.fGrid.Column3.Width+1,.fGrid.Top+2,.fGrid.Column4.Width-3,.fGrid.HeaderHeight-3,''
ENDWITH
fkat.Show
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readkat
PARAMETERS parlog
SELECT sprkat
IF parlog
   fkat.fGrid.GridMyAppendBlank(1,'kod','name')   
ENDIF
fkat.SetAll('Visible',.T.,'MyTxtBox')
fKat.fGrid.columns(fKat.fGrid.columnCount).SetFocus
lineTop=fkat.fGrid.Top+fkat.fGrid.HeaderHeight+fkat.fGrid.RowHeight*(IIF(fkat.fGrid.RelativeRow<=0,1,fkat.fGrid.RelativeRow)-1)
fkat.nrec=RECNO()
SCATTER TO fkat.dim_ap
fkat.txtBox1.Left=fkat.fGrid.Left+10
fkat.txtBox2.Left=fkat.txtbox1.Left+fkat.txtbox1.Width-1
fkat.txtBox3.Left=fkat.txtbox2.Left+fkat.txtbox2.Width-1
fkat.txtBox4.Left=fkat.txtbox3.Left+fkat.txtbox2.Width-1
fkat.txtbox1.ControlSource='fkat.dim_ap(1)'
fkat.txtbox2.ControlSource='fkat.dim_ap(2)'
fkat.txtbox3.ControlSource='fkat.dim_ap(3)'
fkat.txtbox4.ControlSource='fkat.dim_ap(4)'

fkat.SetAll('Top',linetop,'MyTxtBox')
fkat.SetAll('Height',fkat.fGrid.RowHeight+1,'MyTxtBox')
fkat.SetAll('BackStyle',1,'MyTxtBox')
fkat.txtbox1.Enabled=.F.
fkat.fGrid.Enabled=.F.
fkat.txtbox2.SetFocus
*************************************************************************************************************************
PROCEDURE delkat
SELECT rasp
LOCATE FOR kat=sprkat->kod
IF FOUND()
   SELECT sprkat
   fkat.fGrid.GridNoDelRec 
ELSE 
   SELECT sprkat
   fkat.fGrid.GridDelRec('fkat.fGrid','sprkat')
ENDIF   
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник квалификационных категорий персонала
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE prockval
fkval=CREATEOBJECT('Formspr')
SELECT sprkval
SET ORDER TO 1
GO TOP
WITH fkval    
     .Caption='Справочник квалификационных категорий персонала'
     .ProcExit='DO exitProcKval'   
     .AddProperty('doplkatold',0)    
     DO addButtonOne WITH 'fKval','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fkval','Do readkval WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fKval','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fkval','Do readkval WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fKval','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delkval',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fKval','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO exitProcKval',39,.menucont1.Width,'возврат'             
     DO addmenureadspr WITH 'fkval',"DO writeSprNew WITH 'fkval','fkval.fGrid','sprkval','reppersforkat'","DO exitWriteSpr WITH 'fkval','fkval.fGrid'"
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5               
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fKval.fGrid',4 
          .RecordSource='sprkval'
          .Column1.ControlSource='sprkval.kod'
          .Column2.ControlSource='" "+sprkval.name'    
          .Column3.ControlSource='sprkval.doplkat'
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование квалификационной категории'
          .Column3.Header1.Caption='%'
          .Column1.Width=RettxtWidth(' 1234 ') 
          .Column3.Width=RettxtWidth(' 1234 ')   
          .Column4.Width=0
          .Column2.Width=.Width-.column1.Width-.Column3.Width-SYSMETRIC(5)-13-4               
          .Column3.Header1.Caption='%'
          .Column3.Format='Z'
          .Column1.Alignment=1
          .Column3.Alignment=1
          .colNesInf=2    
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')        
     ENDWITH 
     DO gridSizeNew WITH 'fkval','fGrid','shapeingrid'   
     DO addtxtboxmy WITH 'fkval',1,1,1,.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fkval',2,1,1,.fGrid.Column2.Width+2,.F.,.F.,0
     DO addtxtboxmy WITH 'fkval',3,1,1,.fGrid.Column3.Width+2,.F.,.F.,1     
     .SetAll('Visible',.F.,'MyTxtBox')  
     DO addcontmy WITH 'fkval','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fkval','fkval.cont1','sprkval',1"
     .cont1.SpecialEffect=1        
     DO addcontmy WITH 'fkval','cont2',.cont1.Left+.fGrid.Column1.Width+1,.fGrid.Top+2,.fGrid.Column2.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fkval','fkval.cont2','sprkval',2"
     DO addcontmy WITH 'fkval','cont3',.cont2.Left+.fGrid.Column2.Width+1,.fGrid.Top+2,.fGrid.Column3.Width-3,.fGrid.HeaderHeight-3,''
ENDWITH
fkval.Show
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readkval
PARAMETERS parlog
SELECT sprkval
IF parlog
   fkval.fGrid.GridMyAppendBlank(1,'kod','name')   
ENDIF
fkval.SetAll('Visible',.T.,'MyTxtBox')
fKval.fGrid.columns(fKval.fGrid.columnCount).SetFocus
lineTop=fkval.fGrid.Top+fkval.fGrid.HeaderHeight+fkval.fGrid.RowHeight*(IIF(fkval.fGrid.RelativeRow<=0,1,fkval.fGrid.RelativeRow)-1)
fkval.nrec=RECNO()
SCATTER TO fkval.dim_ap
fkval.txtBox1.Left=fkval.fGrid.Left+10
fkval.txtBox2.Left=fkval.txtbox1.Left+fkval.txtbox1.Width-1
fkval.txtBox3.Left=fkval.txtbox2.Left+fkval.txtbox2.Width-1
fkval.txtbox1.ControlSource='fkval.dim_ap(1)'
fkval.txtbox2.ControlSource='fkval.dim_ap(2)'
fkval.txtbox3.ControlSource='fkval.dim_ap(4)'
fkval.SetAll('Top',linetop,'MyTxtBox')
fkval.SetAll('Height',fkval.fGrid.RowHeight+1,'MyTxtBox')
fkval.SetAll('BackStyle',1,'MyTxtBox')
fkval.txtbox1.Enabled=.F.
fkval.fGrid.Enabled=.F.
fkval.txtbox2.SetFocus
*************************************************************************************************************************
PROCEDURE delkval
SELECT datjob
SET FILTER TO 
LOCATE FOR kv=sprkval.kod
IF FOUND() 
   SELECT sprkval
   fkval.fGrid.GridNoDelRec  
ELSE 
   SELECT sprkval
   fkval.fGrid.GridDelRec('fkval.fGrid','sprkval')  
ENDIF   
*************************************************************************************************************************
*             Процедура замены доплаты за категорию в персонале при изменении % в спр-ке
*************************************************************************************************************************
PROCEDURE reppersforkat
SELECT datjob
SET FILTER TO
REPLACE pkat WITH sprkval.doplkat FOR kv=sprkval.kod
SELECT sprkval
*************************************************************************************************************************
PROCEDURE exitProcKval
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
SELECT sprkval
fkval.fGrid.GridReturn
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник специальностей по аттестации	
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE procsprspec
SELECT sprspec
GO TOP
fkat=CREATEOBJECT('Formspr')
WITH fkat    
     .Caption='Справочник специальностей по аттестации' 
     .ProcExit='DO exitSprSpec' 
     DO addButtonOne WITH 'fKat','menuCont1',10,5,'новая','pencila.ico',"Do readspr WITH 'fKat','Do readspec WITH .T.'",39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fKat','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico',"Do readspr WITH 'fKat','Do readspec WITH .F.'",39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fKat','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','DO delspec',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fKat','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO exitSprSpec',39,.menucont1.Width,'возврат'                       
     
     DO addmenureadspr WITH 'fkat',"DO writeSprNew WITH 'fkat','fkat.fGrid','sprspec'","DO exitWriteSpr WITH 'fkat','fkat.fGrid'"
     
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5                    
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fKat.fGrid',3
          .RecordSource='sprspec'
          .Column1.ControlSource='sprspec.kod'
          .Column2.ControlSource='" "+sprspec.name'    
          .Column1.Header1.Caption='Код'
          .Column2.Header1.Caption='Наименование'  
          .Column1.Width=RettxtWidth(' 1234 ')
          .Column2.Width=.Width-.column1.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0  
          .Column1.Alignment=1    
          .colNesInf=2   
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')         
     ENDWITH  
     DO gridSizeNew WITH 'fkat','fGrid','shapeingrid'   
     DO addtxtboxmy WITH 'fkat',1,1,1,fkat.fGrid.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fkat',2,1,1,fkat.fGrid.Column2.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
     DO addcontmy WITH 'fkat','cont1',.fGrid.Left+13,.fGrid.Top+2,.fGrid.Column1.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fKat','fKat.cont1','sprspec',1"
     .cont1.SpecialEffect=1          
     DO addcontmy WITH 'fkat','cont2',.cont1.Left+.fGrid.Column1.Width+1,.fGrid.Top+2,.fGrid.Column2.Width-3,.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fKat','fKat.cont2','sprspec',2"
ENDWITH
fkat.Show
*************************************************************************************************************************
PROCEDURE exitSprSpec
SELECT sprspec
SET ORDER TO 2
SELECT people
fKat.Release
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readspec
PARAMETERS parlog
SELECT sprspec
IF parlog
   fkat.fGrid.GridMyAppendBlank(1,'kod','name')   
ENDIF
fkat.SetAll('Visible',.T.,'MyTxtBox')
fKat.fGrid.columns(fKat.fGrid.columnCount).SetFocus
lineTop=fkat.fGrid.Top+fkat.fGrid.HeaderHeight+fkat.fGrid.RowHeight*(IIF(fkat.fGrid.RelativeRow<=0,1,fkat.fGrid.RelativeRow)-1)
fkat.nrec=RECNO()
SCATTER TO fkat.dim_ap
fkat.txtBox1.Left=fkat.fGrid.Left+10
fkat.txtBox2.Left=fkat.txtbox1.Left+fkat.txtbox1.Width-1
fkat.txtbox1.ControlSource='fkat.dim_ap(1)'
fkat.txtbox2.ControlSource='fkat.dim_ap(2)'

fkat.SetAll('Top',linetop,'MyTxtBox')
fkat.SetAll('Height',fkat.fGrid.RowHeight+1,'MyTxtBox')
fkat.SetAll('BackStyle',1,'MyTxtBox')
fkat.txtbox1.Enabled=.F.
fkat.fGrid.Enabled=.F.
fkat.txtbox2.SetFocus
*************************************************************************************************************************
PROCEDURE delspec
*SELECT rasp
*LOCATE FOR kat=sprkat->kod
*IF FOUND()
*   SELECT sprkat
*   fkat.fGrid.GridNoDelRec 
*ELSE 
*   SELECT sprkat
   fkat.fGrid.GridDelRec('fkat.fGrid','sprspec')
*ENDIF   
*-------------------------------------------------------------------------------------------------------------------------
*                                       ссылки на статьи в приказах
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE proclinkorder
SELECT sprorder
SET ORDER TO 2
GO TOP
newslink=''
fkat=CREATEOBJECT('Formspr')
WITH fkat    
     .Caption='ссылки на статьи в приказах' 
     .ProcExit='DO exitlinkorder' 
     DO addButtonOne WITH 'fKat','menuCont1',10,5,'редакция','pencil.ico','Do readLinkOrder',39,RetTxtWidth('удаление')+44,'редакция'  
     DO addButtonOne WITH 'fKat','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'возврат','undo.ico','DO exitlinkorder',39,.menucont1.Width,'возврат'                       
          
     DO addButtonOne WITH 'fKat','butSave',10,5,'записать','pencil.ico','Do writeLinkOrder WITH .T.',39,RetTxtWidth('wзаписатьw')+44,'записать'  
     DO addButtonOne WITH 'fKat','butRetRead',.butSave.Left+.butSave.Width+3,5,'возврат','undo.ico','Do writeLinkOrder WITH .F.',39,.butSave.Width,'возврат'    
     .butSave.Visible=.F.
     .butRetRead.Visible=.F.
     WITH .fGrid
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5    
          .Height=.Parent.Height-.Parent.menucont1.Height-5                    
          .RecordSourceType=1
          DO addColumnToGrid WITH 'fKat.fGrid',4
          .RecordSource='sprorder'
          .Column1.ControlSource='strord'
          .Column2.ControlSource='" "+nameord'    
          .Column3.ControlSource='" "+slink'    
          .Column1.Header1.Caption='тип'
          .Column2.Header1.Caption='наименование'  
          .Column3.Header1.Caption='ссылка'  
          .Column1.Width=RettxtWidth('wтипw')
          .Column2.Width=(.Width-.column1.Width)/2
          .Column3.Width=.Width-.column1.Width-.Column2.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0  
          .Column1.Alignment=2    
          .colNesInf=2   
          .SetAll('Movable',.F.,'Column') 
          .SetAll('BOUND',.F.,'Column')         
     ENDWITH  
     DO gridSizeNew WITH 'fkat','fGrid','shapeingrid'   
     DO addtxtboxmy WITH 'fkat',3,1,1,fkat.fGrid.Column3.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')      
ENDWITH
fkat.Show
*************************************************************************************************************************
PROCEDURE exitlinkorder
SELECT sprorder
SET ORDER TO 1
SELECT people
fKat.Release
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readlinkorder
SELECT sprorder
newslink=slink
WITH fKat
     .SetAll('Visible',.F.,'MyCommandButton')
      .butSave.Visible=.T.
     .butRetRead.Visible=.T.
     .SetAll('Visible',.T.,'MyTxtBox')
     .fGrid.columns(.fGrid.columnCount).SetFocus
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .nrec=RECNO()  
     fkat.txtBox3.Left=.fGrid.Left+.fGrid.Column1.Width+.fGrid.column2.Width+12
     fkat.txtbox3.ControlSource='newSlink'
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',fkat.fGrid.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .fGrid.Enabled=.F.
     .txtbox3.SetFocus
ENDWITH      
*************************************************************************************************************************
PROCEDURE writeLinkOrder
PARAMETERS par1
SELECT sprorder
IF par1
   REPLACE slink WITH newslink
ENDIF
WITH fkat
     .SetAll('Visible',.T.,'MyCommandButton')
     .butSave.Visible=.F.
     .butRetRead.Visible=.F.
     .SetAll('Visible',.F.,'MyTxtBox')
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH 
*-------------------------------------------------------------------------------------------------------------------------
*                                       Организация и руководство
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE procboss
IF !USED('boss')
   USE boss IN 0
ENDIF 
SELECT boss
newOffice=office
newAdress=adress
newDolBoss=dolBoss
newFioBoss=fioboss
newunp=unp
newfszn=fszn
newtelok=telok
fBoss=CREATEOBJECT('Formmy')
WITH fBoss    
     .Caption='Организация и руководство'    
     .procExit='DO returnboss WITH .F.'     
     .BackColor=RGB(255,255,255) 
      DO addShape WITH 'fboss',1,10,10,60,100,8       
      DO adtBoxAsCont WITH 'fBoss','cont1',.Shape1.Left+10,.Shape1.Top+10,RetTxtWidth('WНаименование организацииW'),dHeight,'Наименование организации',0,1         
      DO adtbox WITH 'fboss',1,.cont1.Left+.cont1.Width-1,.cont1.Top,300,dHeight,'newOffice',.F.,.T.,0
      
      DO adtBoxAsCont WITH 'fBoss','cont2',.cont1.Left,.cont1.Top+.cont1.Height-1,.cont1.Width,dHeight,'адрес',0,1        
      DO adtbox WITH 'fboss',2,.txtBox1.Left,.cont2.Top,.txtBox1.Width,dHeight,'newAdress',.F.,.T.,0
      
      DO adtBoxAsCont WITH 'fBoss','cont3',.cont1.Left,.cont2.Top+.cont2.Height-1,.cont1.Width,dHeight,'должность руководителя',0,1        
      DO adtbox WITH 'fboss',3,.txtBox1.Left,.cont3.Top,.txtBox1.Width,dHeight,'newDolBoss',.F.,.T.,0
      
      DO adtBoxAsCont WITH 'fBoss','cont4',.cont1.Left,.cont3.Top+.cont3.Height-1,.cont1.Width,dHeight,'ФИО руководителя',0,1        
      DO adtbox WITH 'fboss',4,.txtBox1.Left,.cont4.Top,.txtBox1.Width,dHeight,'newFioBoss',.F.,.T.,0
      
      DO adtBoxAsCont WITH 'fBoss','cont5',.cont1.Left,.cont4.Top+.cont4.Height-1,.cont1.Width,dHeight,'УНП',0,1        
      DO adtbox WITH 'fboss',5,.txtBox1.Left,.cont5.Top,.txtBox1.Width,dHeight,'newunp',.F.,.T.,0
      
      DO adtBoxAsCont WITH 'fBoss','cont6',.cont1.Left,.cont5.Top+.cont5.Height-1,.cont1.Width,dHeight,'номер ФСЗН',0,1        
      DO adtbox WITH 'fboss',6,.txtBox1.Left,.cont6.Top,.txtBox1.Width,dHeight,'newfszn',.F.,.T.,0
      
      DO adtBoxAsCont WITH 'fBoss','cont7',.cont1.Left,.cont6.Top+.cont6.Height-1,.cont1.Width,dHeight,'телефон ОК',0,1        
      DO adtbox WITH 'fboss',7,.txtBox1.Left,.cont7.Top,.txtBox1.Width,dHeight,'newtelok',.F.,.T.,0
            
      .SetAll('SpecialEffect',1,'MytxtBox')
      .Shape1.Width=.cont1.Width+.txtBox1.Width+20
      .Shape1.Height=.cont1.Height*7+20           
      DO addContLabel WITH 'fBoss','cont11',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WWзаписатьWW')*2-15)/2,.Shape1.Top+.Shape1.Height+15,RetTxtWidth('WWзаписатьWW'),dHeight+3,'записать','DO returnboss WITH .T.'
      DO addContLabel WITH 'fBoss','cont12',fBoss.Cont11.Left+fBoss.Cont11.Width+15,fBoss.Cont11.Top,;
         fBoss.Cont11.Width,dHeight+3,'Отмена','DO returnboss WITH .F.'
      .Width=.Shape1.Width+20     
      .Height=.Shape1.Height+.cont11.Height+50     
      .WindowState=0
      .Autocenter=.T.    
      .SetAll('BorderColor',RGB(192,192,192),'ShapeMy') 
ENDWITH
DO pasteImage WITH 'fBoss'
fBoss.Show
*************************************************************************************************************************
PROCEDURE returnboss
PARAMETERS par_log
IF par_log
   REPLACE office WITH newOffice,adress WITH newAdress,dolBoss WITH newDolBoss,fioBoss WITH newFioBoss,unp WITH newunp,fszn WITH newfszn,telok WITH newtelok
ENDIF   
SELECT boss
USE
SELECT people
fboss.Release
*-------------------------------------------------------------------------------------------------------------------------
*                                       Праздничные дни
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE procfete
IF !USED('fete')
   USE fete ORDER 1 IN 0
ENDIF
logAp=.F.
newDat=0
nrec=0
newComment=''
fsupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Праздничные дни'  
     .Width=500
     .procExit='DO exitfete'
     .AddObject('grdFete','gridMynew')     
     WITH .grdFete
          .Top=0     
          .Left=0
          .Width=.Parent.Width
          .Height=.rowHeight*12
          .RecordSourceType=1
          .scrollBars=2   
          .ColumnCount=0         
          .colNesInf=2                        
        
          DO addColumnToGrid WITH 'fSupl.grdFete',3        
           .RecordSource='fete'                                         
          .Column1.ControlSource='fete.datafet'
          .Column2.ControlSource='fete.comment'                   
          
          .Column2.Width=RetTxtWidth('Wдата(дд.мм)')                          
          .Columns(.ColumnCount).Width=0    
          .Column2.Width=.Width-.column1.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Column1.Header1.Caption='день'
          .Column2.Header1.Caption='наименование'
          
          .Column1.Alignment=2  
          .Column2.Alignment=0
          .SetAll('Enabled',.F.,'ColumnMy') 
          .Columns(.ColumnCount).Enabled=.T.          
     ENDWITH 
     DO gridSizeNew WITH 'fSupl','grdFete','shapeingrid',.T.,.F.  
     FOR i=1 TO .grdFete.columnCount 
         .grdFete.Columns(i).Backcolor=fSupl.BackColor           
         .grdFete.Columns(i).DynamicBackColor='IIF(RECNO(fSupl.grdFete.RecordSource)#fSupl.grdFete.curRec,fSupl.BackColor,dynBackColor)'
         .grdFete.Columns(i).DynamicForeColor='IIF(RECNO(fSupl.grdFete.RecordSource)#fSupl.grdFete.curRec,dForeColor,dynForeColor)'        
     ENDFOR 
     
                  
    
     DO addtxtboxmy WITH 'fSupl',1,1,1,.grdFete.Column1.Width+2,.F.,.F.,1
     DO addtxtboxmy WITH 'fSupl',2,1,1,.grdFete.Column2.Width+2,.F.,.F.,0
     .SetAll('Visible',.F.,'MyTxtBox')  
     
     DO addButtonOne WITH 'fSupl','butNew',5,.grdFete.Top+.grdFete.Height+10,'новая','','DO readFete WITH .T.',39,RetTxtWidth('wудалениеw'),'новая'  
     DO addButtonOne WITH 'fSupl','butRead',.butNew.Left+.butNew.Width+3,.butNew.Top,'редакция','','DO readFete WITH .F.',39,.butNew.Width,'редакция'   
     DO addButtonOne WITH 'fSupl','butDel',.butRead.Left+.butRead.Width+3,.butNew.Top,'удаление','','DO delfete',39,.butNew.Width,'удаление'       
     DO addButtonOne WITH 'fSupl','butExit',.butDel.Left+.butDel.Width+3,.butNew.Top,'возврат','','DO exitfete',39,.butNew.Width,'возврат'  
     
     DO addButtonOne WITH 'fSupl','butSave',5,.butNew.Top,'записать','','DO savedayfete WITH .T.',39,RetTxtWidth('wудалениеw'),'записать'  
     DO addButtonOne WITH 'fSupl','butRet',.butNew.Left+.butNew.Width+3,.butNew.Top,'возврат','','DO savedayfete WITH .F.',39,.butNew.Width,'возврат'   
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     
     
     DO addButtonOne WITH 'fSupl','butDelRec',5,.butNew.Top,'удалить','','DO delDatfete WITH .T.',39,RetTxtWidth('wудалениеw'),'удалить'  
     DO addButtonOne WITH 'fSupl','butDelRet',.butDelRec.Left+.butDelRec.Width+3,.butNew.Top,'возврат','','DO delDatFete WITH .F.',39,.butNew.Width,'возврат'   
     .butDelRec.Visible=.F.
     .butDelRet.Visible=.F.
     
     
     .butNew.Left=(.Width-.butNew.Width*4-15)/2
     .butRead.Left=.butNew.Left+.butNew.Width+5
     .butDel.Left=.butRead.Left+.butRead.Width+5
     .butExit.Left=.butDel.Left+.butDel.Width+5
     
     .butSave.Left=(.Width-.butSave.Width*2-10)/2
     .butRet.Left=.butSave.Left+.butSave.Width+10
     
     .butDelRec.Left=(.Width-.butDelRec.Width*2-10)/2
     .butDelRet.Left=.butDelRec.Left+.butDelRec.Width+10
               
     .Height=.grdFete.Height+.butNew.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
********************************************************************************************************************************************************
PROCEDURE readfete
PARAMETERS parlog
SELECT fete
IF parlog
   APPEND BLANK
ENDIF
fSupl.Refresh
nrec=RECNO()
logAp=parlog
newDat=IIF(parLog,00.00,VAL(SUBSTR(datafet,1,2))+VAL(SUBSTR(datafet,4,2))/100)
newComment=IIF(parlog,'',comment)
WITH fSupl
     .SetAll('Visible',.F.,'MyCommandButton')
     .butSave.Visible=.T.
     .butRet.Visible=.T.
     .SetAll('Visible',.T.,'MyTxtBox')
     .grdFete.columns(.grdFete.columnCount).SetFocus

     lineTop=.grdFete.Top+.grdFete.HeaderHeight+.grdFete.RowHeight*(IIF(.grdFete.RelativeRow<=0,1,.grdFete.RelativeRow)-1)
     .txtBox1.Left=.grdFete.Left+10
     .txtBox1.InputMask='99.99'
     .txtBox1.Alignment=2
     .txtBox2.Left=.txtbox1.Left+.txtbox1.Width-1
     .SetAll('Top',linetop,'MyTxtBox')
     .SetAll('Height',.grdFete.RowHeight+1,'MyTxtBox')
     .SetAll('BackStyle',1,'MyTxtBox')
     .txtBox1.ControlSource='newDat'
     .txtBox2.ControlSource='newComment'
     .grdFete.Enabled=.F.     
     .Refresh
     .txtBox1.SetFocus
ENDWITH 
*******************************************************************************************************************************************************
PROCEDURE savedayfete
PARAMETERS par1
IF logAp.AND.!par1
   DELETE
ELSE
   REPLACE comment WITH newcomment,dfete WITH INT(newDat),mfete WITH (newDat-INT(newDat))*100,datafet WITH PADL(LTRIM(STR(dfete)),2,'0')+'.'+PADL(LTRIM(STR(mfete)),2,'0') 
ENDIF
WITH fSupl
     .SetAll('Visible',.T.,'MyCommandButton')
     .SetAll('Visible',.F.,'MyTxtBox')
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butDelRec.Visible=.F.
     .butDelRet.Visible=.F.
     .grdFete.Enabled=.T.
     .grdFete.SetAll('Enabled',.F.,'ColumnMy')
     .grdFete.Column3.Enabled=.T.     
     .Refresh
ENDWITH 
*******************************************************************************************************************************************************
PROCEDURE delfete
WITH fSupl
     .SetAll('Visible',.F.,'MyCommandButton')
     .butDelRec.Visible=.T.
     .butDelRet.Visible=.T.
     .grdFete.Enabled=.F.         
ENDWITH
*******************************************************************************************************************************************************
PROCEDURE delDatFete
PARAMETERS par1
SELECT fete
IF par1 
   DELETE
ENDIF
WITH fSupl
     .SetAll('Visible',.T.,'MyCommandButton')
     .SetAll('Visible',.F.,'MyTxtBox')
     .butSave.Visible=.F.
     .butRet.Visible=.F.
     .butDelRec.Visible=.F.
     .butDelRet.Visible=.F.
     .grdFete.Enabled=.T.
     .grdFete.SetAll('Enabled',.F.,'ColumnMy')
     .grdFete.Column3.Enabled=.T.     
     .Refresh
ENDWITH 
*******************************************************************************************************************************************************
PROCEDURE exitFete
SELECT fete
USE
SELECT people
fSupl.Release
********************************************************************************************************************************************************
*                                           Процедура настройки экрана
********************************************************************************************************************************************************
PROCEDURE setupScreen
SELECT datSet
SCATTER TO dimDatSet
newPathWord=ALLTRIM(pathWord)
newPathDoc=ALLTRIM(pathDoc)
newColorUv=ALLTRIM(datSet.rgbUv)
newColorCxUv=0
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl  
     .Caption='Настройки'
     DO addShape WITH 'fSupl',1,10,20,100,500,8           
     DO adCheckBox WITH 'fSupl','checkColorUv','выделять уволенных цветом',.Shape1.Top+20,.Shape1.Left,250,dHeight,'datset.logColUv',0,.T.  
     .checkColorUv.Left=.Shape1.Left+(.Shape1.Width-.checkColorUv.Width)/2
     DO addContFormNew WITH 'fSupl','txtColorUv',.Shape1.Left+10,.checkColorUv.Top+.checkColorUv.height+10,.Shape1.Width-20,dHeight,'двойной щелчок мыши для выбора цвета',1,.F.,'DO selectColorUv' 
     .txtColorUv.BackColor=RGB(&newColorUv)   
     .Shape1.Height=.checkColorUv.Height+.txtColorUv.Height+60
      
     DO addShape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+20,100,.Shape1.Width,8  
     DO addContFormNew WITH 'fSupl','contWord',.Shape2.Left+10,.Shape2.Top+20,RetTxtWidth('Путь для шаблонов MS WordWW'),dHeight,' Путь для других документовw',0,.F.,'DO selectDirForWord'    
     DO adtBox WITH 'fSupl',1,.contWord.Left+.contWord.Width-1,.contWord.Top,.Shape1.Width-.contWord.Width-20,dHeight,'newpathword',.F.,.F.,0  
     DO addContFormNew WITH 'fSupl','contDoc',.Shape2.Left+10,.contWord.Top+.contWord.Height-1,.contWord.Width,dHeight,' Путь для других документов',0,.F.,'DO selectDirForDoc'    
     DO adtBox WITH 'fSupl',2,.txtBox1.Left,.contDoc.Top,.txtBox1.Width,dHeight,'newpathdoc',.F.,.F.,0  
     .Shape2.Height=.contWord.Height*2+40
     
     *-----------------------------Кнопка возврат---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-RetTxtWidth('wприменитьw')*2-10)/2,.Shape2.Top+.Shape2.Height+20,RetTxtWidth('wприменитьw'),dHeight+5,'применить','DO aplySetup WITH .T.'
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+10,.cont1.Top,.cont1.Width,.cont1.Height,'возврат','DO aplySetup WITH .F.'
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape2.Height+.cont1.Height+80
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************************************************************************************
PROCEDURE selectColorUv
newColorCxUv=GETCOLOR()
IF newColorCxUv#1
   r=MOD(INT(newColorCxUv),256)
   g=MOD(INT(newColorCxUv/256),256)
   b=MOD(INT(newColorCxUv/65536),256)
   newColorUv=LTRIM(STR(r))+','+LTRIM(STR(g))+','+LTRIM(STR(b))
   fSupl.txtColorUv.BackColor=RGB(&newColorUv)
ENDIF
*******************************************************************************************************************************************************
PROCEDURE aplySetup
PARAMETERS parLog
IF parLog
   SELECT datSet
   REPLACE rgbUv WITH newColorUv,pathWord WITH newPathWord,pathDoc WITH newPathDoc
ELSE 
   SELECT datSet
   GATHER FROM dimDatSet  
ENDIF 
fSupl.Release
**********************************************************************************************************
PROCEDURE selectDirForWord
newpathword=GETDIR('','','Укажите папку для сохранения',64)
newpathword=IIF(!EMPTY(newpathword),newpathword,datset.pathword)
fSupl.Refresh
**********************************************************************************************************
PROCEDURE selectDirForDoc
newpathdoc=GETDIR('','','Укажите папку для сохранения',64)
newpathdoc=IIF(!EMPTY(newpathdoc),newpathdoc,datset.pathdoc)
fSupl.Refresh
***********************************************************************************************************
PROCEDURE procShtatPeop1
dateBook=DATE()
fltPodr=''
DIMENSION dimOption(5)
STORE .F. TO dimOption
*dimOption(1) - группа печати
SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     DO addshape WITH 'fSupl',1,10,10,150,400,8 
     DO adlabMy WITH 'fSupl',1,' Дата ',.Shape1.Top+20,.Shape1.Left+20,300,0,.T.,1
     DO adTboxNew WITH 'fSupl','boxDate',.Shape1.Top+20,.lab1.Left+.lab1.Width+10,RetTxtWidth('99/99/9999'),dHeight,'dateBook',.F.,.T.,0
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.boxDate.Width-10)/2
     .boxDate.Left=.lab1.Left+.lab1.Width+10
     DO adCheckBox WITH 'fSupl','checkPodr','Подразделение',.boxDate.Top+.boxdate.Height+10,.Shape1.Left+5,150,dHeight,'dimOption(1)',0,.T.,'DO validCheckPodr'  
     .checkPodr.Left=.Shape1.Left+(.Shape1.Width-.checkPodr.Width)/2
     .Shape1.Height=.boxDate.Height+.checkPodr.Height+50
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+10,.Shape1.Width,.F.,.T.
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WпросмотрW')*3-30)/2,.Shape91.Top+.Shape91.Height+15,RetTxtWidth('WпросмотрW'),dHeight+3,'печать','DO bookprn WITH 1','печать ведомости'
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+3,'просмотр','DO bookprn WITH 2','предварительный просмотр и печать ведомости'   
     *---------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+15,.Cont1.Top,.Cont1.Width,dHeight+3,'возврат','fSupl.Release','возврат'
     
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape91.Height+20,.Shape1.Width  
     
     
     *-----------------------------Кнопка принять---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont11',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wпринять')*2)-15)/2,.cont1.Top,RetTxtWidth('wпринятьw'),dHeight+5,'принять','DO returnToBook WITH .T.'
     .cont11.Visible=.F.
     *---------------------------------Кнопка сброс-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont12',.cont11.Left+.cont11.Width+15,.Cont11.Top,.Cont11.Width,dHeight+5,'сброс','DO returnToBook WITH .F.'
     .cont12.Visible=.F.
     
     .AddObject('lstLine','LINE')
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH    
     WITH .lstLine   
          .Top=.Parent.listBox1.Top
          .Height=.Parent.listBox1.Height
          .Left=.Parent.ListBox1.Left+40+3
          .Width=0    
          .Visible=.F.
     ENDWITH 
     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+60
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE validCheckPodr
dimOption(2)=.F. 
dimOption(3)=.T.
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
     .lstLine.Visible=.T.
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
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
fSupl.lstLine.Visible=.T.
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
    dimOption(1)=.F.
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
     .lstLine.Visible=.F.
     .listBox1.Visible=.F. 
     dimOption(1)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='Подразделение'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH 
*********************************************************************************************************
PROCEDURE bookPrn
PARAMETERS parTerm
IF USED('curPrn')
   SELECT curPrn
   USE   
ENDIF
SELECT * FROM datJob INTO CURSOR curTarPeople READWRITE 
SELECT curTarPeople
DELETE FOR dateOut>dateBook
DELETE FOR !EMPTY(dateOut).AND.dateOut<dateBook
DELETE FOR dateBeg>dateBook
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL 
REPLACE staj_in WITH IIF(SEEK(kodpeop,'people',1),people.staj_in,'') ALL 
REPLACE date_in WITH IIF(SEEK(kodpeop,'people',1),people.date_in,'') ALL 
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
REPLACE nd WITH IIF(SEEK(STR(kp,3)+STR(kd,3),'rasp',2),rasp.nd,0) ALL
INDEX ON STR(np,3)+STR(nd,3)+fio TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG T2
SET ORDER TO 2
SCAN ALL
     DO actualStajToday WITH 'curTarPeople','curTarPeople.date_in','dateBook'    
ENDSCAN

CREATE CURSOR curPrn (np N(3),nd N(3),kp N(3),kd N(3),named C(150), fio C(70),kse N(7,2),tr N(1),nametr C(15),kat N(1),logP L,Kpp N(3),KodKpp N(3),KodKpp1 N(3),npp N(3),vac L,nvac N(1),kodpeop N(5),kv N(1),dkv C(60), pkont N(3),staj_today C(10),primtxt C (30))
kserasp_kat=0
kserasp_cx=0
fltch=''

SELECT rasp      
SELECT * FROM rasp INTO CURSOR currasp READWRITE      
SELECT currasp
REPLACE np WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.np,0) ALL
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1
IF !EMPTY(fltPodr)
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
kseRasp=0
kseVac=0
ksePeop=0
SCAN ALL
     kseRasp=kse
     ksePeop=0
     SELECT curPrn
     APPEND BLANK
     REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kse WITH curRasp.kse,kd WITH curRasp.kd,kp WITH curRasp.kp,logp WITH .T.
     SELECT curTarPeople 
     IF SEEK(STR(curRasp.kp,3)+STR(curRasp.kd,3))
        DO WHILE kp=curRasp.kp.AND.kd=curRasp.kd
                 SELECT curPrn
                 APPEND BLANK 
                 REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kp WITH curRasp.kp,kd WITH curRasp.kd,kat WITH curRasp.kat,kse WITH curTarPeople.kse,tr WITH curTarPeople.tr,;
                         kodpeop WITH curTarPeople.kodpeop,fio WITH curTarPeople.fio,nametr WITH IIF(SEEK(tr,'sprtype',1),sprtype.name,''),KodKpp WITH curRasp.KodKpp,KodKpp1 WITH curRasp.KodKpp1,;
                         pkont WITH curTarPeople.pkont,kv WITH curTarPeople.kv,dkv WITH IIF(SEEK(kv,'sprkval',1),sprkval.name,''),staj_today WITH curTarPeople.staj_today,;
                         primTxt WITH IIF(SEEK(curTarPeople.kodpeop,'people',1).AND.people.dekOtp,'д/отп.'+IIF(!EMPTY(people.ddekotp),' до '+DTOC(people.dDekotp),''),'')           
                 ksePeop=ksePeop+IIF(curprn.tr=4,0,curPrn.kse)
                 SELECT curTarPeople
                 SKIP
        ENDDO          
     ENDIF
     IF kseRasp-ksePeop>0  
        SELECT curPrn
        APPEND BLANK            
        REPLACE np WITH curRasp.np,nd WITH curRasp.nd,kp WITH curRasp.kp,kat WITH curRasp.kat,kd WITH curRasp.kd,kse WITH kseRasp-ksePeop,fio WITH 'Вакантная',;
                KodKpp WITH curRasp.KodKpp,KodKpp1 WITH curRasp.KodKpp1,tr WITH 1,vac WITH .T. 
     ENDIF
     SELECT curRasp
ENDSCAN
SELECT curPrn
REPLACE named WITH IIF(SEEK(curPrn.kd,'sprdolj',1),sprdolj.name,'') ALL
REPLACE nvac WITH 1 FOR vac
*INDEX ON STR(np,3)+STR(nd,3)+fio+STR(tr,1) TAG T1
INDEX ON STR(np,3)+STR(nd,3)+STR(nvac,1)+fio+STR(tr,1) TAG T1
INDEX ON kp TAG T2
SET ORDER TO 1
SELECT curPrn

*DO fltstructure WITH 'kd>0.AND.kp>0','curPrn'
SELECT curPrn
GO TOP
IF parTerm=1
   DO procForPrintAndPreview WITH 'repBook','штатная кгига',.T.
ELSE 
   DO procForPrintAndPreview WITH 'repBook','штатная кгига',.F. 
ENDIF 
************************************************************************************************************************
*                            основная для отпусков
************************************************************************************************************************
PROCEDURE formForOtp
IF RECCOUNT('curOtpSupl')#0
   frmTop.grdOtp.Columns(frmTop.grdOtp.ColumnCount).SetFocus
ENDIF   
fSupl=CREATEOBJECT('FORMSUPL')
DIMENSION dimSelectOtp(3)
STORE 0 TO dimSelectOtp
dimSelectOtp(1)=1
WITH fSupl 
     .Caption='Отпуска'
     DO addShape WITH 'fSupl',1,20,20,10,10,8
     DO addOptionButton WITH 'fSupl',1,'добавить новую запись',.Shape1.Top+10,.Shape1.Left+10,'dimSelectOtp(1)',0,'DO procSelectOtp WITH 1',.T. 
     DO addOptionButton WITH 'fSupl',2,'редактировать текущую',.Option1.Top+.Option1.Height+10,.Option1.Left,'dimSelectOtp(2)',0,'DO procSelectOtp WITH 2',IIF(reccount('curOtpSupl')=0,.F.,.T.)
     DO addOptionButton WITH 'fSupl',3,'удалить текущую запись',.Option2.Top+.Option2.Height+10,.Option2.Left,'dimSelectOtp(3)',0,'DO procSelectOtp WITH 3',IIF(reccount('curOtpSupl')=0,.F.,.T.)      
     .Shape1.Height=.Option1.Height*3+60
     .Shape1.Width=.Option3.Width+20
     *-----------------------------Кнопка приступить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wприступитьw')*2)-20)/2,;
        .Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприступитьw'),dHeight+5,'приступить','DO procRunOtp'

     *---------------------------------Кнопка отмена --------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'отмена','fSupl.Release','отмена'    
    .Width=.Shape1.Width+40
    .Height=.Shape1.Height+.cont1.Height+60
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE procSelectOtp
PARAMETERS par1
STORE 0 TO dimSelectOtp
dimSelectOtp(par1)=1
fSupl.Refresh
************************************************************************************************************************
PROCEDURE procRunOtp
fSupl.Visible=.F.
fSupl.Release
DO CASE
   CASE dimSelectOtp(1)=1
        DO inputRecInOtp WITH .T.  
   CASE dimSelectOtp(2)=1
        DO inputRecInOtp WITH .F.
   CASE dimSelectOtp(3)=1
        DO deleteFromJob WITH 'DO delRecFromOtp'
ENDCASE
************************************************************************************************************************
PROCEDURE inputRecInOtp
PARAMETERS par1
IF !USED('fete')
   USE fete IN 0 ORDER 1
ENDIF
frmOtp=CREATEOBJECT('FORMSUPL')
str_otp=IIF(par1,'',IIF(SEEK(curOtpsupl.kodotp,'curSprotp',1),curSprotp.name,''))
str_prich=IIF(par1,'',IIF(SEEK(curOtpsupl.kprich,'curPrichOtp',1),curPrichOtp.name,''))
new_kodotp=IIF(par1,0,curOtpsupl.kodotp)
new_prichotp=IIF(par1,0,curOtpsupl.kprich)
new_perBeg=IIF(par1,CTOD('  .  .    '),curOtpsupl.perbeg)
new_perEnd=IIF(par1,CTOD('  .  .    '),curOtpsupl.perend)
new_kvoDay=IIF(par1,0,curOtpsupl.kvoDay)
new_dayOtp=IIF(par1,0,curOtpsupl.dayOtp)
new_dayKont=IIF(par1,0,curOtpsupl.dayKont)
new_dayVred=IIF(par1,0,curOtpsupl.dayVred)
new_dayNorm=IIF(par1,0,curOtpsupl.dayNorm)
new_BegOtp=IIF(par1,CTOD('  .  .    '),curOtpsupl.begotp)
new_EndOtp=IIF(par1,CTOD('  .  .    '),curOtpsupl.endotp)
new_osnov=IIF(par1,'',curOtpsupl.osnov)
WITH frmOtp    
     .Caption='Отпуска'
     DO adTboxAsCont WITH 'frmOtp','txtVid',10,10,RetTxtWidth('wОсновной Отпуск (дней)w'),dHeight,'вид отпуска',1,1       
     DO adTboxAsCont WITH 'frmOtp','txtPrich',.txtVid.Left,.txtVid.Top+.txtVid.Height-1,.txtVid.Width,dHeight,'причина с',1,1              
     DO adTboxAsCont WITH 'frmOtp','txtPerBeg',.txtVid.Left,.txtPrich.Top+.txtPrich.Height-1,.txtVid.Width,dHeight,'период с',1,1 
          
     DO adTboxAsCont WITH 'frmOtp','txtPerEnd',.txtVid.Left,.txtPerBeg.Top,RetTxtWidth('wокончание'),dHeight,'период по',2,1
     
     DO adTboxAsCont WITH 'frmOtp','txtDayOtp',.txtVid.Left,.txtPerBeg.Top+.txtPerBeg.Height-1,.txtVid.Width,dHeight,'основной отпуск(дней)',1,1 
     DO adTboxAsCont WITH 'frmOtp','txtDayKont',.txtVid.Left,.txtDayOtp.Top+.txtDayOtp.Height-1,.txtVid.Width,dHeight,'поощрит.отпуск(дней)',1,1 
     DO adTboxAsCont WITH 'frmOtp','txtDayVred',.txtVid.Left,.txtDayKont.Top+.txtDayKont.Height-1,.txtVid.Width,dHeight,'за вредность (дней)',1,1 
     DO adTboxAsCont WITH 'frmOtp','txtDayNorm',.txtVid.Left,.txtDayVred.Top+.txtDayVred.Height-1,.txtVid.Width,dHeight,'за ненормир.(дней)',1,1 
     DO adTboxAsCont WITH 'frmOtp','txtDayTot',.txtVid.Left,.txtDayNorm.Top+.txtDayNorm.Height-1,.txtVid.Width,dHeight,'всего дней',1,1 
           
     DO adTboxAsCont WITH 'frmOtp','txtBeg',.txtVid.Left,.txtDayTot.Top+.txtDayTot.Height-1,.txtVid.Width,dHeight,'начало',1,1 
     DO adTboxAsCont WITH 'frmOtp','txtEnd',.txtVid.Left,.txtBeg.Top,RettxtWidth('wокончаниеw'),dHeight,'окончание',2,1
     DO adTboxAsCont WITH 'frmOtp','txtOsn',.txtVid.Left,.txtEnd.Top+.txtEnd.Height-1,.txtVid.Width,dHeight,'основание',1,1
       
     DO addComboMy WITH 'frmOtp',1,.txtVid.Left+.txtVid.Width-1,.txtVid.Top,dheight,RetTxtWidth('по семейным обстоятельствам без сохраненияw'),.T.,'str_otp','curSprotp.name',6,.F.,'DO validOtp',.F.,.T.  
     DO addComboMy WITH 'frmOtp',2,.comboBox1.Left,.txtPrich.Top,dheight,.comboBox1.Width,.T.,'str_prich','curPrichOtp.name',6,.F.,'DO validPrichOtp',.F.,.T.                           
     .comboBox2.Enabled=IIF(new_kodOtp=1,.F.,.T.)
     .comboBox2.Style=IIF(.comboBox2.Enabled=.T.,2,0)     
     DO adTboxNew WITH 'frmOtp','boxPerBeg',.txtPerBeg.Top,.comboBox1.Left,RetTxtWidth('99/99/999999'),dHeight,'new_perBeg',.F.,IIF(new_kodotp=1,.T.,.F.),0                                
     DO adTboxNew WITH 'frmOtp','boxPerEnd',.txtPerBeg.Top,.comboBox1.Left,.boxPerBeg.Width,dHeight,'new_perEnd',.F.,IIF(new_kodotp=1,.T.,.F.),0
     
     DO adTboxNew WITH 'frmOtp','boxDayOtp',.txtDayOtp.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'new_dayOtp','Z',.T.,0,.F.,'DO sumDayRecOtp'
     DO adTboxNew WITH 'frmOtp','boxDayKont',.txtDayKont.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'new_dayKont','Z',.T.,0,.F.,'DO sumDayRecOtp'
     DO adTboxNew WITH 'frmOtp','boxDayVred',.txtDayVred.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'new_dayVred','Z',.T.,0,.F.,'DO sumDayRecOtp'
     DO adTboxNew WITH 'frmOtp','boxDayNorm',.txtDayNorm.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'new_dayNorm','Z',.T.,0,.F.,'DO sumDayRecOtp'
     DO adTboxNew WITH 'frmOtp','boxDayTot',.txtDayTot.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'new_KvoDay','Z',.T.,0
      
     .txtPerEnd.Width=.comboBox1.Width-.boxPerBeg.Width*2+2
     .txtPerEnd.Left=.boxPerBeg.Left+.boxPerbeg.Width-1
     .boxPerEnd.Left=.txtPerEnd.Left+.txtPerEnd.Width-1 
          
     DO adTboxNew WITH 'frmOtp','boxBeg',.txtBeg.Top,.comboBox1.Left,.boxPerBeg.Width,dHeight,'new_BegOtp',.F.,.T.,0,.F.,;
        "DO validPerOtp WITH 'new_begotp','new_endOtp','new_kvoDay','new_kodOtp'"   
     
     .txtEnd.Width=.txtPerEnd.Width
     .txtEnd.Left=.txtPerEnd.Left
     DO adTboxNew WITH 'frmOtp','boxEnd',.txtBeg.Top,.txtEnd.Left+.txtEnd.Width-1,.boxBeg.Width,dHeight,'new_EndOtp',.F.,.T.,0    
     DO adTboxNew WITH 'frmOtp','boxOsn',.txtOsn.Top,.comboBox1.Left,.comboBox1.Width,dHeight,'new_Osnov',.F.,.T.,0 
       
     .Width=.txtVid.Width+.comboBox1.Width+20           
      
     DO addcontlabel WITH 'frmOtp','cont1',(.Width-RetTxtWidth('wЗаписатьw')*2-20)/2,.txtOsn.Top+.txtOsn.Height+20,;
        RetTxtWidth('wЗаписатьw'),dHeight+3,'Записать',IIF(par1,'DO writeRecInOtp WITH .T.','DO writeRecInOtp WITH .F.')
     DO addcontlabel WITH 'frmOtp','cont2',frmOtp.Cont1.Left+frmOtp.Cont1.Width+20,frmOtp.Cont1.Top,;
        .Cont1.Width,dHeight+3,'Отмена','frmOtp.Release'       
     .Height=.txtVid.height*10+.cont1.Height+40     
ENDWITH 
DO pasteImage WITH 'frmOtp'
frmOtp.Show
*************************************************************************************************************************
PROCEDURE validOtp
new_kodOtp=curSprotp.kod
WITH frmOtp
     .boxPerBeg.Enabled=IIF(INLIST(new_kodOtp,1,2,3,5,9),.T.,.F.)
     .boxPerEnd.Enabled=IIF(INLIST(new_kodOtp,1,2,3,3,5,9),.T.,.F.)
     .comboBox2.Enabled=IIF(new_kodOtp=1,.F.,.T.)
     .comboBox2.Style=IIF(.comboBox2.Enabled=.T.,2,0)
     .boxdayOtp.Enabled=IIF(INLIST(new_kodOtp,1,3),.T.,.F.)
     .boxdayKont.Enabled=IIF(new_kodOtp=1,.T.,.F.)
     .boxdayVred.Enabled=IIF(new_kodOtp=1,.T.,.F.)
     .boxdayNorm.Enabled=IIF(new_kodOtp=1,.T.,.F.)
     new_dayOtp=IIF(new_kodOtp=1,people.dayOtp,0)
     new_dayKont=IIF(new_kodOtp=1,people.dayKont,0)
     new_dayVred=IIF(new_kodOtp=1,people.dayVred,0)
     new_dayNorm=IIF(new_kodOtp=1,people.dayNorm,0)
     new_kvoDay=IIF(new_kodOtp=1,people.dayOtp+people.dayKont+people.dayVred+people.dayNorm,new_kvoDay)
     KEYBOARD '{TAB}'
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE validPerOtp
PARAMETERS parBeg,parEnd,parDay,parVid
* parBeg - начало отпуска
* parEnd - окончание отпуска
* parDay - дней отпуска
* parVid
&parEnd=&parBeg+&parDay-1
IF &parVid=1
   SELECT fete
   GO TOP   
   SCAN ALL
        IF YEAR(&parBeg)=YEAR(&parEnd)      
           IF CTOD(datafet+'.'+STR(YEAR(&parBeg),4))>=&parBeg..AND.CTOD(datafet+'.'+STR(YEAR(&parBeg),4))<=&parEnd        
              &parEnd=&parEnd+1
           ENDIF
        ELSE 
           IF VAL(RIGHT(datafet,2))=MONTH(&parBeg).AND.VAL(LEFT(datafet,2))>=DAY(&parBeg)
              &parEnd=&parEnd+1                                
           ENDIF
           IF VAL(RIGHT(datafet,2))=MONTH(&parEnd).AND.VAL(LEFT(datafet,2))<=DAY(&parEnd)
              &parEnd=&parEnd+1                       
           ENDIF
        ENDIF    
   ENDSCAN  
   
    
ENDIF  
SELECT people
*frmOtp.boxDay.Refresh
*************************************************************************************************************************
PROCEDURE validPrichOtp
new_prichOtp=curPrichOtp.kod
*************************************************************************************************************************
PROCEDURE writeRecInOtp
PARAMETERS par1
frmOtp.Visible=.F.
SELECT datOtp
IF par1   
   APPEND BLANK
   REPLACE kodpeop WITH people.num,nidpeop WITH people.nid
ELSE
   SEEK STR(people.num,5)   
   SCAN WHILE kodPeop=people.num
        IF kodotp=curOtpsupl.kodOtp.AND.perBeg=curOtpsupl.perbeg.AND.perEnd=curOtpsupl.perEnd.AND.kvoDay=curOtpsupl.kvoDay.AND.begOtp=curOtpsupl.begOtp.AND.endOtp=curOtpsupl.endOtp                   
           EXIT 
        ENDIF
   ENDSCAN       
ENDIF
REPLACE kodOtp WITH new_kodOtp,perBeg WITH new_PerBeg,perEnd WITH new_perEnd,kvoDay WITH new_kvoDay,;
        begOtp WITH new_begOtp,endOtp WITH new_endOtp,osnov WITH new_osnov,nameOtp WITH str_otp,kprich WITH new_prichOtp,txtprich WITH str_prich,;
        dayOtp WITH new_dayOtp,dayKont WITH new_DayKont,dayVred WITH new_dayVred,dayNorm WITH new_DayNorm
IF kPrich=3
   SELECT people
   REPLACE dekOtp WITH .T.,dekEnd WITH new_endOtp
   SELECT datOtp
ENDIF 
frmOtp.Release   
frmTop.Refresh 
SELECT people
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
DO changeRowGrdPers
**********************************************************************************************************************      \
PROCEDURE sumDayRecOtp
new_kvoDay=new_dayOtp+new_dayKont+new_dayVred+new_dayNorm
frmOtp.boxDayTot.Refresh
***********************************************************************************************************************
PROCEDURE deleteFromJob
PARAMETERS parproc
fdel=CREATEOBJECT('FORMSUPL')
log_del=.F.
WITH fDel
     .Caption='Удаление'    
     DO addShape WITH 'fDel',1,20,20,100,RetTxtWidth('wпоставьте птичку в окошке, расположенном нижеw'),8         
     DO adLabMy WITH 'fDel',1,'для подтверждения ваших намерений',fDel.Shape1.Top+10,fDel.Shape1.Left+5,.Shape1.Width-10,2 
     DO adLabMy WITH 'fDel',2,'поставьте птичку в окошке, расположенном ниже',.lab1.Top+.lab1.Height,fDel.Shape1.Left+5,.lab1.Width,2                                   
     DO adCheckBox WITH 'fdel','check1','подтверждение удаления',.lab2.Top+.lab2.Height+10,.Shape1.Left,150,dHeight,'log_del',0    
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     .Shape1.Height=.check1.Height+.lab1.Height*2+30
     DO addContLabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('wУдалитьw')*2-20)/2,fdel.check1.Top+fdel.check1.Height+20,;
        RetTxtWidth('wУдалитьw'),dHeight+3,'Удалить','&parproc'
     DO addContLabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+20,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'     
     .Width=.Shape1.Width+40   
     .Height=.Shape1.Height+.cont1.Height+60     
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
***********************************************************************************************************************
PROCEDURE delRecFromOtp
IF !log_del
   RETURN
ENDIF
fDel.Release
SELECT datOtp
SEEK STR(people.num,5)   
SCAN WHILE kodPeop=people.num
     IF kodotp=curOtpsupl.kodOtp.AND.perBeg=curOtpsupl.perbeg.AND.perEnd=curOtpsupl.perEnd.AND.kvoDay=curOtpsupl.kvoDay.AND.begOtp=curOtpsupl.begOtp.AND.endOtp=curOtpsupl.endOtp                   
        DELETE 
        EXIT 
     ENDIF
ENDSCAN   
frmTop.grdPers.Columns(frmTop.grdPers.ColumnCount).SetFocus
frmTop.Refresh    
****************************************************************************************************************************************************************************************
PROCEDURE procEndKont
log_term=.T.
logWord=.F.
kvo_page=1
page_beg=1
page_end=999
term_ch=.T.
dateBeg=CTOD('  .  .    ')
dateEnd=CTOD('  .  .    ')
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сотрудников'        
     DO addshape WITH 'fSupl',1,20,20,150,400,8
     
     DO adlabMy WITH 'fSupl',1,'Период',.Shape1.Top+20,10,100,0,.T.     
     DO adtbox WITH 'fSupl',1,.lab1.Left+.lab1.Width+10,.lab1.Top,RetTxtWidth('99/99/99999'),dHeight,'dateBeg',.F.,.T.,.F.
     .lab1.Top=.lab1.Top+(.txtBox1.Height-.lab1.Height+4)    
     
     DO adlabMy WITH 'fSupl',2,' - ',.lab1.Top,10,100,0,.T.     
     DO adtbox WITH 'fSupl',2,.lab1.Left+.lab1.Width+10,.txtBox1.Top,.txtBox1.Width,dHeight,'dateEnd',.F.,.T.,.F.
            
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.txtBox1.Width-.lab2.Width-.txtBox2.Width-15)/2
     .txtBox1.Left=.lab1.Left+.lab1.Width+5
     .lab2.Left=.txtBox1.Left+.txtBox1.Width+5
     .txtBox2.Left=.lab2.Left+.lab2.Width+5
     
           
     .Shape1.Height=.txtBox1.Height+40        
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.T.
      
     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO prnRepKont WITH .T.' 
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,;
        .Cont1.Width,dHeight+5,'Просмотр','DO prnRepKont WITH .F.'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','fSupl.Release','Выход из печати'      
                              
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+80
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************************
PROCEDURE prnRepKont
PARAMETERS parLog
CREATE CURSOR curKontrakt (fio c(50),ndog C(20),enddog D) 
SELECT curKontrakt
INDEX ON enddog TAG T1
SELECT people
oldRec=RECNO()
SCAN ALL
     IF !otmUv       
        IF IIF(!EMPTY(dateEnd),!EMPTY(enddog).AND.enddog<=dateEnd,!EMPTY(enddog))          
           SELECT  curKontrakt
           APPEND BLANK
           REPLACE fio WITH people.fio,enddog WITH people.enddog,ndog WITH IIF(SEEK(people.dog,'sprdog',1),sprdog.name,'')  
        ENDIF       
        SELECT people
     ENDIF    
ENDSCAN
SELECT curKontrakt
IF !EMPTY(dateBeg)
   DELETE FOR enddog<dateBeg
   DELETE FOR enddog>dateEnd
ENDIF
SELECT people
GO oldRec
SELECT curKontrakt
GO TOP
DO procForPrintAndPreview WITH 'repKontrakt','',parLog,'repKontToExcel'
********************************************************************************************************************************
PROCEDURE repKontToExcel
#DEFINE xlCenter -4108            
#DEFINE xlLeft -4131  
#DEFINE xlRight -4152  
#DEFINE xlThin 2                  
#DEFINE xlMedium -4138            
#DEFINE xlDiagonalDown 5          
#DEFINE xlDiagonalUp 6                 
#DEFINE xlEdgeLeft 7              
#DEFINE xlEdgeTop 8               
#DEFINE xlEdgeBottom 9            
#DEFINE xlEdgeRight 10            
#DEFINE xlInsideVertical 11         
#DEFINE xlInsideHorizontal 12        
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=40
     .Columns(2).ColumnWidth=17
     .Columns(3).ColumnWidth=12
     .cells(2,1).Value='ФИО сотрудника'              
     .cells(2,1).HorizontalAlignment= -4108             
     .cells(2,2).Value='заключён'
     .cells(2,2).HorizontalAlignment= -4108              
     .cells(2,3).Value='окончание'
     .cells(2,3).HorizontalAlignment= -4108             
              
     *.Range(.Cells(2,1),.Cells(2,5)).Select
     *objExcel.Selection.Interior.ColorIndex=35          
     .Range(.Cells(1,1),.Cells(1,3)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='Список сотрудников'
          .Interior.ColorIndex=35
     ENDWITH                                                                                                           
     numberRow=3
     yearMonth=''  
     SELECT curKontrakt
     SCAN ALL
          IF yearMonth#STR(MONTH(enddog),2)+STR(YEAR(enddog),4)        
             .Range(.Cells(numberRow,1),.Cells(numberRow,3)).Select
             WITH objExcel.Selection
                  .MergeCells=.T.
                  .HorizontalAlignment= -4108                
                  .WrapText=.T.
                  .Value=dim_month(MONTH(enddog))+' '+STR(YEAR(enddog),4)
                  .Interior.ColorIndex=37
                  numberRow=numberRow+1
             ENDWITH  
             yearMonth=STR(MONTH(enddog),2)+STR(YEAR(enddog),4)                    
          ENDIF
          .cells(numberRow,1).Value=fio
          .cells(numberRow,2).Value=ndog
          .cells(numberRow,3).Value=enddog         
          numberRow=numberRow+1         
     ENDSCAN
     .Range(.Cells(1,1),.Cells(numberRow-1,3)).Select
     objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
     objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
     objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
     objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
     objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
     .Range(.Cells(1,1),.Cells(1,3)).Select
ENDWITH 
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
objExcel.Visible=.T. 
***************************************************************************************************************************************************************
PROCEDURE formspisinout
DIMENSION dim_opt(2),dim_ord(3),dim_tr(3)
dim_opt(1)=1
dim_opt(2)=0

*dimOption(1) - подразделение
*dimOption(2) - должность
*dimOption(3) - тип работы
*dimOption(4) - персонал

dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим
dim_ord(3)=0 && по причине увольнения

dim_tr(1)=0  &&все
dim_tr(2)=1  &&основной
dim_tr(3)=0  &&вн.совм.
DO procDimFlt

STORE CTOD('  .  .    ') TO date_Beg,date_End
lOut=.F.
lIn=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список приятых, уволенных'
     .procexit='fSupl.Release' 
     DO procObjFlt 
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8    
     DO addOptionButton WITH 'fSupl',11,'принятые',.Shape2.Top+10,.Shape2.Left+20,'dim_opt(1)',0,'DO validInMove WITH 1',.T. 
     DO addOptionButton WITH 'fSupl',12,'уволенные',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_opt(2)',0,'DO validInMove WITH 2',.T. 
     .Option11.Left=.Shape2.Left+(.Shape2.Width-.Option11.Width-.Option12.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+20 
     
     DO adLabMy WITH 'fSupl',1,'период с ',.Option11.Top+.option11.Height+10,.Shape2.Left,.Shape2.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxBeg',.Option11.Top+.Option11.Height+10,.Shape2.Left,RetTxtWidth('99/99/99999'),dHeight,'date_Beg',.F.,.T.,0
     .lab1.Top=.boxBeg.Top+(.boxBeg.Height-.lab1.Height)+3
     
     DO adLabMy WITH 'fSupl',2,' по ',.lab1.Top,.Shape2.Left,.Shape2.Width,0,.T.,1  
     DO adTboxNew WITH 'fSupl','boxEnd',.boxBeg.Top,.Shape2.Left,.boxBeg.Width,dHeight,'date_End',.F.,.T.,0
     .lab1.Left=.Shape2.Left+(.Shape2.Width-.lab1.Width-.boxBeg.Width-.lab2.Width-.boxEnd.Width-30)/2
     .boxBeg.Left=.lab1.Left+.lab1.Width+10
     .lab2.Left=.boxBeg.Left+.boxBeg.Width+10
     .boxEnd.Left=.lab2.Left+.lab2.Width+10
     
     DO addOptionButton WITH 'fSupl',21,'все',.boxBeg.Top+.boxBeg.Height+10,.Shape2.Left+20,'dim_tr(1)',0,"DO procValOption WITH 'fSupl','dim_tr',1",.T. 
     DO addOptionButton WITH 'fSupl',22,'основн.',.Option21.Top,.Option21.Left+.Option21.Width+10,'dim_tr(2)',0,"DO procValOption WITH 'fSupl','dim_tr',2",.T. 
     DO addOptionButton WITH 'fSupl',23,'вн. совм.',.Option21.Top,.Option21.Left+.Option21.Width+10,'dim_tr(3)',0,"DO procValOption WITH 'fSupl','dim_tr',3",.T. 
     .Option21.Left=.Shape2.Left+(.Shape2.Width-.Option21.Width-.Option22.Width-.Option23.Width-20)/2
     .Option22.Left=.Option21.Left+.Option21.Width+10 
     .Option23.Left=.Option22.Left+.Option22.Width+10      
          
     DO addOptionButton WITH 'fSupl',1,'по алфавиту',.option21.Top+.option21.Height+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'по дате',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     DO addOptionButton WITH 'fSupl',3,'по причине увольнения',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(3)',0,"DO procValOption WITH 'fSupl','dim_ord',3",.T. 
     .Option3.Enabled=.F.
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-.Option3.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+10 
     .Option3.Left=.Option2.Left+.Option2.Width+10 
     .Shape2.Height=.Option11.height*3+.boxBeg.Height+50
          
     DO adSetupPrnToForm WITH .Shape2.Left,.Shape2.Top+.Shape2.Height+10,.Shape2.Width,.F.,.T.
     DO adButtonPrnToForm WITH 'DO prnSpisInOut WITH 1','DO prnSpisInOut WITH 2','fSupl.Release',.T.  
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH      
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width
     .Width=.Shape2.Width+20
     .Height=.butPrn.Top+.butPrn.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************
PROCEDURE validInMove
PARAMETERS par1
STORE 0 TO dim_opt
dim_opt(par1)=1
fSupl.Option3.Enabled=IIF(dim_opt(2)=1,.T.,.F.)
IF dim_ord(3)=1.AND.dim_opt(1)=1
   dim_ord(3)=0
   dim_ord(1)=1
ENDIF 
fSupl.Refresh
**************************************************
PROCEDURE prnspisInOut
PARAMETERS par1
IF !lIn.AND.!lOut
   RETURN  
ENDIF
IF USED('curInJob')
   SELECT curInJob
   USE
ENDIF
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curUvolOrd')
   SELECT curUvolOrd
   USE
ENDIF 
topsay='Список '

DO CASE 
   CASE dim_opt(1)=1  &&принятые
        SELECT * FROM datjob WHERE INLIST(tr,1,3) INTO CURSOR curinJob READWRITE
        SELECT curInjob
        APPEND FROM datjobout FOR INLIST(tr,1,3)
        INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING        
        DO CASE
           CASE dim_tr(1)=1
                DELETE FOR INLIST(tr,2,4,5)
           CASE dim_tr(2)=1
                DELETE FOR tr#1
           CASE dim_tr(3)=1
                DELETE FOR tr#3
        ENDCASE    
        INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING     
        DELETE FOR !EMPTY(date_in).AND.date_in>date_End
        DELETE FOR !EMPTY(date_in).AND.date_in<date_Beg 
   CASE dim_opt(2)=1  &&уволенные
        SELECT * FROM datjobout INTO CURSOR curinJob READWRITE
        SELECT curInjob
        DO CASE
           CASE dim_tr(1)=1
                DELETE FOR INLIST(tr,2,4,5)
           CASE dim_tr(2)=1
                DELETE FOR tr#1
           CASE dim_tr(3)=1
                DELETE FOR tr#3
        ENDCASE
        INDEX ON STR(kodpeop,4)+STR(tr,1)+STR(kse,5,2) TAG T1 DESCENDING 
        DELETE FOR EMPTY(dateout)
        DELETE FOR dateout<date_Beg
        DELETE FOR dateout>date_End
ENDCASE 

DO CASE  
   CASE dim_opt(1)=1 
        SELECT * FROM people WHERE !EMPTY(date_in) INTO CURSOR curprn READWRITE
        SELECT curprn
        APPEND FROM peopout 
        DELETE FOR date_in<date_Beg
        DELETE FOR date_in>date_End
        ALTER TABLE curprn ADD COLUMN kd N(3)
        ALTER TABLE curprn ADD COLUMN kp N(3)
        ALTER TABLE curprn ADD COLUMN npp N(3)
        ALTER TABLE curprn ADD COLUMN tr N(1)
        ALTER TABLE curprn ADD COLUMN kat N(2)
        REPLACE kp WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.kp,0) kd WITH curinjob.kd, kat WITH curinjob.kat,tr WITH curinjob.tr ALL
        *REPLACE kd WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.kd,0) ALL
        *REPLACE kp WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.kp,0) ALL
        *REPLACE tr WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.tr,0) ALL
        topsay=topsay+'принятых за период с '+DTOC(date_Beg)+' по '+DTOC(date_End)  
        DO CASE
           CASE dim_ord(1)=1
                INDEX ON fio TAG T1
           CASE dim_ord(2)=1
                INDEX ON date_in TAG T1
        ENDCASE
   CASE dim_opt(2)=1
        SELECT * FROM peoporder WHERE !EMPTY(dateuvol).AND.typeord=2 INTO CURSOR curUvolOrd READWRITE 
        SELECT curUvolOrd
        INDEX ON nidpeop TAG T1
        SELECT * FROM peopout WHERE !EMPTY(date_out) INTO CURSOR curprn READWRITE
        DELETE FOR date_out<date_Beg
        DELETE FOR date_out>date_End
        ALTER TABLE curprn ADD COLUMN kd N(3)
        ALTER TABLE curprn ADD COLUMN kp N(3)
        ALTER TABLE curprn ADD COLUMN npp N(3)
        ALTER TABLE curprn ADD COLUMN tr N(1)
        ALTER TABLE curprn ADD COLUMN kat N(2)
        ALTER TABLE curprn ADD COLUMN nkpr N(3)
        ALTER TABLE curprn ADD COLUMN cnamepr C(60)
        
        REPLACE kp WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.kp,0) kd WITH curinjob.kd, kat WITH curinjob.kat,tr WITH curinjob.tr ALL
        REPLACE nkpr WITH IIF(SEEK(nid,'curUvolOrd',1),curUvolOrd.supord,0) cnamepr WITH IIF(SEEK(nkpr,'sprorder',1),sprorder.nameord,'')ALL
        
        *REPLACE kd WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.kd,0) ALL
         *REPLACE tr WITH IIF(SEEK(STR(num,4),'curinJob',1),curinjob.tr,0) ALL
        topsay=topsay+'уволенных за период с '+DTOC(date_Beg)+' по '+DTOC(date_End)  
        DO CASE
           CASE dim_ord(1)=1
                INDEX ON fio TAG T1
           CASE dim_ord(2)=1
                INDEX ON date_out TAG T1
           CASE dim_ord(3)=1
                INDEX ON STR(nkpr,3)+DTOS(date_out) TAG T1
        ENDCASE       
ENDCASE
SELECT curprn
DO CASE
   CASE dim_tr(2)=1
        DELETE FOR tr#1
   CASE dim_tr(3)=1
        DELETE FOR tr#3
ENDCASE
DO applyflt
nppcx=1
SCAN ALL
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'repinout','список сотрудников',.T.,'spisinoutToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'repinout','список сотрудников',.F.,'spisinoutToExcel'
ENDCASE
******************************************************************************************************************
PROCEDURE spisinoutToExcel 
DO startPrnToExcel WITH 'fSupl'
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)
maxColumn=IIF(dim_opt(1)=1,9,10)
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=40  
     .Columns(3).ColumnWidth=60
     .Columns(4).ColumnWidth=60    
     .Columns(5).ColumnWidth=11 
     .Columns(6).ColumnWidth=17  
     .Columns(7).ColumnWidth=11  
     .Columns(8).ColumnWidth=11  
     .Columns(9).ColumnWidth=12  
     IF dim_opt(2)=1
        .Columns(10).ColumnWidth=30               
     ENDIF
     
     .Range(.Cells(1,1),.Cells(1,maxColumn)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .WrapText=.T.
           .Font.Name='Times New Roman'   
           .Font.Size=11
           .Value=topsay
      ENDWITH   
      .cells(2,1).Value='№'
      .Cells(2,2).Value='ФИО'                 
      .Cells(2,3).Value='Подразделение'                 
      .Cells(2,4).Value='Должность'                 
      .Cells(2,5).Value='Дата рожд'                 
      .Cells(2,6).Value='Личный номер'                 
      .Cells(2,7).Value='Дата приема'                 
      .Cells(2,8).Value='Дата увольнения'                 
      .Cells(2,9).Value='Тип работы'                 
      IF dim_opt(2)=1
         .Cells(2,10).Value='Причина увольнения'                 
      ENDIF  
      .Range(.Cells(2,1),.Cells(2,maxColumn)).Select
      objExcel.Selection.HorizontalAlignment=xlCenter
      SELECT curPrn
      DO storezeropercent
      numberRow=3
       SCAN ALL         
            .cells(numberRow,1).Value=npp
            .cells(numberRow,2).Value=fio
            .cells(numberRow,3).Value=IIF(SEEK(kp,'sprpodr',1),ALLTRIM(sprpodr.namework),'')
            .cells(numberRow,4).Value=IIF(SEEK(kd,'sprdolj',1),ALLTRIM(sprdolj.name),'')
            .cells(numberRow,5).Value=IIF(!EMPTY(age),age,'')
            .cells(numberRow,6).Value=pnum
            .cells(numberRow,7).Value=IIF(!EMPTY(date_in),DTOC(date_in),'')
            .cells(numberRow,8).Value=IIF(!EMPTY(date_out),DTOC(date_out),'')
            .cells(numberRow,9).Value=IIF(SEEK(tr,'sprtype',1),ALLTRIM(sprtype.name),'')
            IF dim_opt(2)=1
               .cells(numberRow,10).Value=cnamepr
            ENDIF
            DO fillpercent WITH 'fSupl'
            numberRow=numberRow+1
      ENDSCAN
      .Range(.Cells(2,1),.Cells(numberRow-1,maxColumn)).Select
      WITH objExcel.Selection
           .Borders(xlEdgeLeft).Weight=xlThin
           .Borders(xlEdgeTop).Weight=xlThin            
           .Borders(xlEdgeBottom).Weight=xlThin
           .Borders(xlEdgeRight).Weight=xlThin
           .Borders(xlInsideVertical).Weight=xlThin
           .Borders(xlInsideHorizontal).Weight=xlThin
           .VerticalAlignment=1
           .Font.Name='Times New Roman'   
           .Font.Size=11
           .WrapText=.T.
      ENDWITH   
      .Cells(1,1).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'
objExcel.Visible=.T.
************************************************************************************************************
*     Список находящихся в декретном отпуске+ распределение должностей
************************************************************************************************************
PROCEDURE formspisdekotp
DIMENSION dim_opt(3),dim_ord(2)
dim_opt(1)=1
dim_opt(2)=0
dim_opt(3)=0
*dimOption(1) - только список
*dimOption(2) - список+распределение
*dimOption(3) - только распределение


dim_ord(1)=1 &&  алфавитный режим
dim_ord(2)=0 && штатный режим

DO procDimFlt

STORE CTOD('  .  .    ') TO dateBeg,dateEnd
lOut=.F.
lIn=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сотрудников, находящихся в отпуске по уходу за ребенком до 3-х лет'   
    * .procexit='Do exitprn'
     DO procObjFlt       
     DO addshape WITH 'fSupl',2,.Shape1.Left,.Shape1.Top+.Shape1.Height+10,150,.Shape1.Width,8  
              
     DO addOptionButton WITH 'fSupl',11,'список',.Shape2.Top+10,.Shape2.Left+20,'dim_opt(1)',0,"DO procValOption WITH 'fSupl','dim_opt',1",.T. 
     DO addOptionButton WITH 'fSupl',12,'список+замещение',.Option11.Top,.Option11.Left+.Option11.Width+20,'dim_opt(2)',0,"DO procValOption WITH 'fSupl','dim_opt',2",.T. 
     DO addOptionButton WITH 'fSupl',13,'замещение',.Option11.Top,.Option12.Left+.Option12.Width+20,'dim_opt(3)',0,"DO procValOption WITH 'fSupl','dim_opt',3",.T. 
     .Option11.Left=.Shape2.Left+(.Shape2.Width-.Option11.Width-.Option12.Width-.Option13.Width-20)/2
     .Option12.Left=.Option11.Left+.Option11.Width+10
     .Option13.Left=.Option12.Left+.Option12.Width+10
     
     DO addOptionButton WITH 'fSupl',1,'по алфавиту',.option11.Top+.option11.Height+10,.Shape2.Left+20,'dim_ord(1)',0,"DO procValOption WITH 'fSupl','dim_ord',1",.T. 
     DO addOptionButton WITH 'fSupl',2,'по штатному расписанию',.Option1.Top,.Option1.Left+.Option1.Width+20,'dim_ord(2)',0,"DO procValOption WITH 'fSupl','dim_ord',2",.T. 
     .Option1.Left=.Shape2.Left+(.Shape2.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20 
     
     .Shape2.Height=.Option11.height+.Option2.Height+30
          
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape2.Top+.Shape2.Height+10,.Shape1.Width,.F.,.T.
     
     DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape2.Height+.Shape91.Height+20,.Shape1.Width  
     WITH .listBox1                  
          .RowSourceType=2
          .ColumnCount=2
          .ColumnWidths='40,360' 
          .ColumnLines=.F.
          .ControlSource=''          
          .Visible=.F.     
     ENDWITH 
     DO adButtonPrnToForm WITH 'DO prnSpisDekotp WITH 1','DO prnSpisDekotp WITH 2','fSupl.Release',.T.       
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width        
     
     .Width=.Shape1.Width+20
     .Height=.butPrn.Top+.butPrn.Height+20
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show  
**************************************************************************************************************
PROCEDURE prnSpisDekOtp          
PARAMETERS par1
topsay='Список сотрудников, находящихся в отпуске по уходу за ребенком до 3-х лет.'
SELECT * FROM people WHERE dekotp INTO CURSOR curdek
CREATE CURSOR curprn (kodpeop N(6),npp N(3),fiodek C(60),kp N(3),kd N(3),namep C(100),named C(100),bdekotp D,ddekotp D,np N(3),nd N(3),kat N(2))

SELECT * FROM datjob  INTO CURSOR dekJob READWRITE
SELECT dekJob
DELETE FOR !EMPTY(dateout)
DELETE FOR !INLIST(tr,1,3)
INDEX ON nidpeop TAG T1
SELECT curDek
SCAN ALL
     SELECT dekJob
     SEEK curdek.nid 
     SELECT curprn
     APPEND BLANK
     REPLACE kodpeop WITH curdek.num,fiodek WITH curdek.fio,bdekotp WITH curdek.bdekotp,ddekotp WITH curDek.ddekotp,kp WITH dekjob.kp,kd WITH dekjob.kd,;
             namep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.name,''),named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,''),kat WITH dekjob.kat
     SELECT curDek
ENDSCAN
SELECT curprn
DO applyflt
INDEX ON fiodek TAG T1
nppcx=1
SCAN ALL
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO CASE 
   CASE par1=1
        DO procForPrintAndPreview WITH 'spisdek','список сотрудников',.T.,'spisdekToExcel'
   CASE par1=2
        DO procForPrintAndPreview WITH 'spisdek','список сотрудников',.F.,'spisdekToExcel'
ENDCASE
*******************************************************************************************************
PROCEDURE spisdekToExcel
************************************************************************************************************
PROCEDURE spisaes
log_term=.T.
logWord=.F.
term_ch=.T.
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список "чернобыльцев"'      
     DO adSetupPrnToForm WITH 10,10,400,.F.,.T.
     DO adButtonPrnToForm WITH 'DO prnSpisAes WITH .T.','DO prnSpisAes WITH .F.','fsupl.Release'
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width 
     .Width=.Shape91.Width+40
     .Height=.butPrn.Top+.butPrn.Height+10
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
****************************************************************************************************************************
PROCEDURE prnSpisAes
PARAMETERS parLog
IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curJobAge')
   SELECT curJobAge
   USE
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 
SELECT * FROM people  WHERE chaes INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN kp N(3)
ALTER TABLE curPrn ADD COLUMN namep C(100)
ALTER TABLE curPrn ADD COLUMN kd N(3)
ALTER TABLE curPrn ADD COLUMN named C(100)
ALTER TABLE curPrn ADD COLUMN npp N(3)
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,namep WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),named WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') all

INDEX ON fio TAG T1
nppcx=1
SCAN ALL
     SELECT curPrn      
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO procForPrintAndPreview WITH 'repspisaes','',parLog,'repspisAesToExcel'
********************************************************************************************************************************
PROCEDURE repSpisAesToExcel
ON ERROR DO erSup
DO startPrnToExcel WITH 'fSupl'       
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=50
     .Columns(4).ColumnWidth=50     
     .cells(2,1).Value='№'                                     
     .cells(2,2).Value='ФИО'              
     .cells(2,3).Value='подразделение'
     .cells(2,4).Value='должность'    
     .Range(.Cells(1,1),.Cells(1,4)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment= -4108         
          .WrapText=.T.
          .Value='Список "чернобыльцев"'          
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,4)).Select
     objExcel.Selection.HorizontalAlignment= -4108         

     numberRow=3
     SELECT curPrn
     DO storezeropercent
     SCAN ALL        
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio
          .cells(numberRow,3).Value=namep
          .cells(numberRow,3).WrapText=.T.
          .cells(numberRow,4).Value=named 
          .cells(numberRow,4).WrapText=.T.   
          DO fillpercent WITH 'fSupl'      
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,3)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'   
         .Font.Size=10
    ENDWITH 
    .Range(.Cells(1,1),.Cells(1,3)).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl' 
ON ERROR            
objExcel.Visible=.T. 
*********************************************************************************************************************
PROCEDURE procretirement
DIMENSION dim_rt(2)
dim_rt(1)=1
dim_rt(2)=0
perBeg=CTOD('  .  .    ')
perEnd=CTOD('  .  .    ')
agemen=63.0
agewom=57.0
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список пенсионеров'        
     DO addshape WITH 'fSupl',1,20,20,150,400,8                             
          
     DO addOptionButton WITH 'fSupl',1,'работающие пенсионеры',.Shape1.Top+10,.Shape1.Left+20,'dim_rt(1)',0,'DO procValidRt WITH 1',.T. 
     
     DO addOptionButton WITH 'fSupl',2,'выход на пенсию',.Option1.Top,.Option1.Left,'dim_rt(2)',0,'DO procValidRt WITH 2',.T. 
     .Option1.Left=.Shape1.Left+(.Shape1.Width-.Option1.Width-.Option2.Width-20)/2
     .Option2.Left=.Option1.Left+.Option1.Width+20     
     
     DO adLabMy WITH 'fSupl',1,'период выхода на пенсию',.Option1.Top+.Option1.Height+10,.Shape1.Left+5,.Shape1.Width-10,2,.F.,1 
     
     DO adtbox WITH 'fSupl',1,.Shape1.Left+20,.lab1.Top+.lab1.Height+10,RetTxtWidth('99/99/99999'),dHeight,'perBeg',.F.,.F.,.F.
     DO adtbox WITH 'fSupl',2,.txtBox1.Left+.txtBox1.Width+10,.txtBox1.Top,RetTxtWidth('99/99/99999'),dHeight,'perEnd',.F.,.F.,.F.
       
     .txtBox1.Left=.Shape1.Left+(.Shape1.Width-.txtBox1.Width-.txtBox2.Width-10)/2
     .txtBox2.Left=.txtBox1.Left+.txtBox1.Width+10
     
     DO adLabMy WITH 'fSupl',2,'возраст для выхода на пенсию',.txtBox1.Top+.txtBox1.Height+10,.Shape1.Left+5,.Shape1.Width-10,2,.F.,1 
     
     DO adLabMy WITH 'fSupl',3,'мужчины',.lab2.Top+.lab2.Height,.Shape1.Left+5,100,2,.T.,1 
     DO adtbox WITH 'fSupl',3,.lab3.Left+.lab3.Width+10,.lab3.Top,RetTxtWidth('9999999'),dHeight,'agemen','Z',.F.,.F.
     .txtBox3.InputMask='99.9'
     .lab3.Top=.txtBox3.Top+(.txtBox3.Height-.lab3.Height+3)
     
     DO adLabMy WITH 'fSupl',4,'женщины',.lab3.Top,.Shape1.Left+5,100,2,.T.,1 
     DO adtbox WITH 'fSupl',4,.lab4.Left+.lab4.Width+10,.txtBox3.Top,.txtBox3.Width,dHeight,'agewom','T',.F.,.F.
     .txtBox4.InputMask='99.9'
     
     .lab3.Left=.Shape1.Left+(.Shape1.Width-.lab3.Width-.lab4.Width-.txtBox3.Width*2-40)/2
     .txtBox3.Left=.lab3.Left+.lab3.Width+10
     .lab4.Left=.txtBox3.Left+.txtBox3.Width+20
     .txtBox4.Left=.lab4.Left+.lab4.Width+10
             
     .Shape1.Height=.txtBox1.Height*2+.Option1.Height+.lab1.Height*2+70   
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.T.     
    
     DO adButtonPrnToForm WITH 'DO prnRetirment WITH .T.','DO prnRetirment WITH .F.','fSupl.Release',.T.          
    * DO addListBoxMy WITH 'fSupl',1,.Shape1.Left,.Shape1.Top,.Shape1.Height+.Shape91.Height+20,.Shape1.Width  
     DO addShapePercent WITH 'fSupl',.Shape91.Left,.butPrn.Top,.butPrn.Height,.Shape91.Width   
     
     .Width=.Shape1.Width+40
     .Height=.butPrn.Top+.butPrn.Height+10
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
*****************************************************************************************
PROCEDURE procValidRt
PARAMETERS par1
STORE 0 TO dim_rt
dim_rt(par1)=1
WITH fSupl
     .txtBox1.Enabled=IIF(dim_rt(1)=1,.F.,.T.)
     .txtBox2.Enabled=IIF(dim_rt(1)=1,.F.,.T.)
     .txtBox3.Enabled=IIF(dim_rt(1)=1,.F.,.T.)
     .txtBox4.Enabled=IIF(dim_rt(1)=1,.F.,.T.)
     .Refresh
ENDWITH
****************************************************************************************
PROCEDURE prnRetirment
PARAMETERS parLog
IF USED('curprn')
   SELECT curprn
   USE
ENDIF
IF USED('curjobage')
   SELECT curjobage
   USE
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 

DO CASE
   CASE dim_rt(1)=1 &&работающие пенсионеры
        SELECT * FROM people WHERE pens INTO CURSOR curPrn READWRITE
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ADD COLUMN kp N(3)
        ALTER TABLE curPrn ADD COLUMN kd N(3)
        ALTER TABLE curPrn ADD COLUMN npodr C(100)
        ALTER TABLE curPrn ADD COLUMN ndol C(100)
        ALTER TABLE curPrn ADD COLUMN dpens D
        INDEX ON fio TAG T1
   CASE dim_rt(2)=1
        SELECT * FROM people WHERE !pens INTO CURSOR curPrn READWRITE 
        ALTER TABLE curPrn ADD COLUMN npp N(3)
        ALTER TABLE curPrn ADD COLUMN kp N(3)
        ALTER TABLE curPrn ADD COLUMN kd N(3)
        ALTER TABLE curPrn ADD COLUMN npodr C(100)
        ALTER TABLE curPrn ADD COLUMN ndol C(100)
        ALTER TABLE curPrn ADD COLUMN dpens D
        DELETE FOR EMPTY(age)
        REPLACE dpens WITH CTOD(STR(DAY(age),2)+'.'+STR(MONTH(age),2)+'.'+STR(YEAR(age)+IIF(sex=1,INT(agemen),INT(agewom)),4)) ALL
        IF agewom-INT(agewom)#0.OR.agemen-INT(agemen)#0
           SCAN ALL
                yearpens=YEAR(dpens)
                monthpens=MONTH(dpens)+IIF(sex=1,(agemen-INT(agemen))*10,(agewom-INT(agewom))*10)
                IF monthpens>12
                   yearpens=YEAR(dpens)+1
                   monthpens=monthpens-12                               
                ENDIF                
                REPLACE dpens WITH CTOD(STR(DAY(dpens),2)+'.'+STR(monthpens,2)+'.'+STR(yearpens,4))
           ENDSCAN
        ENDIF   
        DELETE FOR dpens<perBeg.OR.dpens>perEnd  
        INDEX ON dpens TAG T1
ENDCASE
SELECT curprn
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,npodr WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),ndol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL
nppcx=1
SCAN ALL 
     REPLACE npp WITH nppcx
     nppcx=nppcx+1
ENDSCAN 
GO TOP
DO procForPrintAndPreview WITH 'retirment','',parLog,'retirmentToExcel'
****************************************************************************************
PROCEDURE retirmentToExcel
DO startPrnToExcel WITH 'fSupl' 
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=30
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=30
     .Columns(4).ColumnWidth=10
     .Columns(5).ColumnWidth=10
     .cells(2,1).Value='ФИО сотрудника'              
     .cells(2,2).Value='подразделение'
     .cells(2,3).Value='должность'
     .cells(2,4).Value='день рождения'                                    
     .cells(2,5).Value='дата выхода'                                    
     .Range(.Cells(1,1),.Cells(1,5)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter      
          .WrapText=.T.
          .Value='Список сотрудников'
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,5)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter      

     numberRow=3
     yearMonth=''
     SELECT curprn
     DO storezeropercent
     SCAN ALL        
          .cells(numberRow,1).Value=fio
          .cells(numberRow,2).Value=npodr
          .cells(numberRow,3).Value=ndol
          .cells(numberRow,4).Value=IIF(!EMPTY(age),DTOC(age),'')
          .cells(numberRow,5).Value=IIF(!EMPTY(dpens),DTOC(dpens),'')
          DO fillpercent WITH 'fSupl'
          numberRow=numberRow+1         
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,5)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .Font.Name='Times New Roman'   
         .Font.Size=10
    ENDWITH 
    .Range(.Cells(1,1),.Cells(1,5)).Select
ENDWITH 
DO endPrnToExcel WITH 'fSupl'       
objExcel.Visible=.T. 

**************************************************************************************************************************
PROCEDURE formSprav1
PARAMETERS parUv
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
dDspr=DATE()
dSprsost=DATE()
cStajWork=''
IF !parUv
   SELECT * FROM datjob WHERE nidpeop=people.nid.AND.EMPTY(dateOut) INTO CURSOR curjobsprav
   cOrd=IIF(SEEK(people.nid,'datjob',8).AND.!EMPTY(datjob.dOrdin), ' (приказ № '+ALLTRIM(datjob.nOrdin)+' от '+DTOC(datjob.dOrdIn)+'г.)','') 
   IF EMPTY(cOrd)
      cOrd=IIF(SEEK(STR(people.nid,5),'peoporder',3), ' (приказ № '+ALLTRIM(peoporder.nOrd)+' от '+DTOC(peoporder.dOrd)+'г.)','') 
   ENDIF 
   cAgeChar=IIF(!EMPTY(people.age),LTRIM(STR(DAY(people.age)))+' '+month_prn(MONTH(people.age))+' '+STR(YEAR(people.age),4)+' г.','')
   DO actualStajToday WITH 'people','people.date_in','DATE()','cStajWork',.T.
ELSE 
   SELECT * FROM datjobout WHERE nidpeop=peopout.nid INTO CURSOR curjobsprav
   cOrd=IIF(SEEK(peopout.nid,'datjobout',8).AND.!EMPTY(datjobout.dOrdin), ' (приказ № '+ALLTRIM(datjobout.nOrdin)+' от '+DTOC(datjobout.dOrdIn)+'г.)','') 
   IF EMPTY(cOrd)
      cOrd=IIF(SEEK(STR(peopout.nid,5),'peoporder',3), '(приказ № '+ALLTRIM(peoporder.nOrd)+' от '+DTOC(peoporder.dOrd)+'г.)','') 
   ENDIF 
   cAgeChar=IIF(!EMPTY(peopout.age),LTRIM(STR(DAY(peopout.age)))+' '+month_prn(MONTH(peopout.age))+' '+STR(YEAR(peopout.age),4)+' г.','')
   DO actualStajToday WITH 'peopout','peopout.date_in','peopout.date_out','cStajWork',.T.
ENDIF    
cFio=IIF(!parUv,ALLTRIM(people.fio),ALLTRIM(peopout.fio))
cNpodr=''
cNdol=''

cStaj=''
*ON ERROR DO erSup
IF !EMPTY(cStajWork)
   cStaj=IIF(VAL(SUBSTR(cStajWork,1,2))#0,SUBSTR(cStajWork,1,2),SUBSTR(cStajWork,4,2))
   cYear=IIF(VAL(SUBSTR(cStajWork,1,2))#0,.T.,.F.)
   cStaj=IIF(LEFT(cStaj,1)='0',SUBSTR(cStaj,2),cStaj)
   IF cYear
      cStaj=cStaj+' '+IIF(INLIST(RIGHT(cStaj,1),'2','3'),'года',IIF(INLIST(RIGHT(cStaj,1),'1'),'год','лет'))
   ELSE
      cStaj=cStaj+' '+IIF(INLIST(RIGHT(cStaj,1),'2','3'),'месяца',IIF(INLIST(RIGHT(cStaj,1),'1'),'месяц','месяцев'))
   ENDIF    
ENDIF
*ON ERROR 
cEduc=IIF(SEEK(IIF(!parUv,people.educ,peopout.educ),'cureducation',1),cureducation.name,'')   
cNpodrcahr=''
cNdolchar=''

DO CASE
   CASE !people.lvn
        SELECT curjobsprav
        IF !paruv
           LOCATE FOR tr=1           
           IF !FOUND()
              GO TOP
           ENDIF
        ELSE 
           LOCATE FOR tr=1.AND.dateout=peopout.date_out           
        ENDIF    
        cNpodr=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
        cNdol=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namet),ALLTRIM(sprdolj.namet),ALLTRIM(sprdolj.name)),'')
       
        cNpodrchar=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
        cNdolchar=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namework),ALLTRIM(sprdolj.namework),ALLTRIM(sprdolj.namework)),'')
        
   CASE people.lvn  
        SELECT curjobsprav
        LOCATE FOR tr=3
        IF !FOUND()
           GO TOP
        ENDIF
        cNpodr=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
        cNdol=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namet),ALLTRIM(sprdolj.namet),ALLTRIM(sprdolj.name)),'')
        cNpodrchar=IIF(SEEK(curjobsprav.kp,'sprpodr',1),IIF(!EMPTY(sprpodr.nameord),ALLTRIM(sprpodr.nameord),ALLTRIM(sprpodr.name)),'')
        cNdolchar=IIF(SEEK(curjobsprav.kd,'sprdolj',1),IIF(!EMPTY(sprdolj.namework),ALLTRIM(sprdolj.namework),ALLTRIM(sprdolj.name)),'')
ENDCASE
SELECT curSprav
REPLACE txtsprav WITH LOWER(cNdol)+' '+LOWER(cNpodr)+' c '+IIF(!EMPTY(IIF(!paruv,people.date_in,peopout.date_in)),LTRIM(STR(DAY(IIF(!paruv,people.date_in,peopout.date_in))))+' '+;
        ALLTRIM(month_prn(MONTH(IIF(!paruv,people.date_in,peopout.date_in))))+' '+STR(YEAR(IIF(!paruv,people.date_in,peopout.date_in)),4)+' года','')+cOrd+' по настоящее время.'
REPLACE txtChar WITH LOWER(cNdolchar)+' '+LOWER(cNpodrchar)+' '+cstaj
SELECT people
WITH fSupl
     .Caption='Справка о месте работы и заниимаемой должности, характеристика с места работы'  
      DO addshape WITH 'fSupl',1,20,20,150,400,8   
      DO adtBoxAsCont WITH 'fSupl','cont1',.Shape1.Left+10,.Shape1.Top+20,RetTxtWidth('WДата выдачиW'),dHeight,'дата выдачи',2,1         
      DO adTboxNew WITH 'fSupl','tBox1',.cont1.Top+.cont1.Height-1,.cont1.Left,.cont1.Width,dHeight,'dDspr',.F.,.T.,0     
      DO adtBoxAsCont WITH 'fSupl','cont11',.cont1.Left+.cont1.Width-1,.cont1.Top,RetTxtWidth('w№ справки'),dHeight,'№ спр-ки',2,1   
      DO adTboxNew WITH 'fSupl','tBox11',.tBox1.Top,.cont11.Left,.cont11.Width,dHeight,'cNspr',.F.,.T.,0     
      
      DO adtBoxAsCont WITH 'fSupl','cont2',.cont11.Left+.cont11.Width-1,.cont1.Top,RetTxtWidth('wдля предоставления по месту требования и в другие присутственные местаw'),dHeight,'адресат',2,1   
      DO adTboxNew WITH 'fSupl','tBox2',.tBox1.Top,.cont2.Left,.cont2.Width,dHeight,'cAdr',.F.,.T.,0     
      DO adtBoxAsCont WITH 'fSupl','cont3',.cont1.Left,.tBox1.Top+.tBox1.Height-1,.cont1.Width+.cont11.Width+.cont2.Width-2,dHeight,'кем работает',2,1   
      .AddObject('editSprav','MyEditBox')      
      WITH .editSprav
          .Visible=.T.          
          .ControlSource='curSprav.txtSprav'
          .Left=.Parent.cont1.Left
          .Width=.Parent.cont3.Width
          .Top=.Parent.cont3.Top+.Parent.cont3.Height-1
          .Height=dHeight*2
          .Enabled=.T.  
     ENDWITH
     DO adtBoxAsCont WITH 'fSupl','cont4',.cont1.Left,.editSprav.Top+.editSprav.Height-1,RetTxtWidth('WДата выдачиW'),dHeight,'состояние на',2,1         
     DO adTboxNew WITH 'fSupl','tBox4',.cont4.Top+.cont4.Height-1,.cont1.Left,.cont4.Width,dHeight,'dSprsost',.F.,.T.,0     
     DO adtBoxAsCont WITH 'fSupl','cont5',.cont4.Left+.cont4.Width-1,.cont4.Top,RetTxtWidth('Wподпись должностьW'),dHeight,'подписть должность',2,1         
     DO adTboxNew WITH 'fSupl','tBox5',.tbox4.Top,.cont5.Left,.cont5.Width,dHeight,'boss.cspravdol',.F.,.T.,0  
     DO adtBoxAsCont WITH 'fSupl','cont6',.cont5.Left+.cont5.Width-1,.cont4.Top,.cont3.Width-.cont4.Width-.cont5.Width+2,dHeight,'подписть ФИО',2,1         
     DO adTboxNew WITH 'fSupl','tBox6',.tbox4.Top,.cont6.Left,.cont6.Width,dHeight,'boss.cspravfio',.F.,.T.,0      
     .Shape1.Width=.cont3.Width+20
     .Shape1.Height=.cont1.Height*3+.tBox1.Height*2+.editSprav.Height+40
    .Width=.Shape1.Width+40      
     
     DO addButtonOne WITH 'fSupl','butPrn',.Shape1.Left+(.Shape1.Width-RetTxtWidth('WПросмотрW')*2-20)/2,.Shape1.Top+.Shape1.Height+20,'печать','','DO spravtoword1',39,RetTxtWidth('WПросмотрW'),'печать'
     DO addButtonOne WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+20,.butPrn.Top,'возврат','','fSupl.Release',39,.butPrn.Width,'возврат' 
    
    
     DO addShape WITH 'fSupl',2,.Shape1.Left,.butPrn.Top+.butPrn.Height+20,300,.Shape1.Width,8
     DO adtBoxAsCont WITH 'fSupl','cont1ch',.Shape2.Left+10,.Shape2.Top+20,RetTxtWidth('WWДата рожденияW'),dHeight,'дата рождения',2,1         
     DO adTboxNew WITH 'fSupl','tBox1ch',.cont1ch.Top+.cont1ch.Height-1,.cont1ch.Left,.cont1ch.Width,dHeight,'cAgeChar',.F.,.T.,0     
     DO adtBoxAsCont WITH 'fSupl','cont11ch',.cont1ch.Left+.cont1ch.Width-1,.cont1ch.Top,RetTxtWidth('W99 месяцевW '),dHeight,'стаж',2,1
     DO adTboxNew WITH 'fSupl','tBox11ch',.tBox1ch.Top,.cont11ch.Left,.cont11ch.Width,dHeight,'cStaj',.F.,.T.,0     
     DO adtBoxAsCont WITH 'fSupl','cont12ch',.cont11ch.Left+.cont11ch.Width-1,.cont1ch.Top,.cont3.Width-.cont1ch.Width-.cont11ch.Width+2,dHeight,'образование',2,1   
     DO adTboxNew WITH 'fSupl','tBox12ch',.tBox1ch.Top,.cont12ch.Left,.cont12ch.Width,dHeight,'cEduc',.F.,.T.,0     
                       
     DO adtBoxAsCont WITH 'fSupl','cont3ch',.cont1ch.Left,.tBox1ch.Top+.tBox1ch.Height-1,.cont3.Width,dHeight,'кем работает',2,1   
      .AddObject('editChar','MyEditBox')      
      WITH .editChar
           .Visible=.T.          
           .ControlSource='curSprav.txtChar'
           .Left=.Parent.cont1ch.Left
           .Width=.Parent.cont3.Width
           .Top=.Parent.cont3ch.Top+.Parent.cont3ch.Height-1
           .Height=dHeight*2
           .Enabled=.T.  
     ENDWITH
     .Shape2.Height=.cont1ch.Height*2+.tBox1ch.Height+.editChar.Height+40
     
     DO addButtonOne WITH 'fSupl','butPrnCh',.Shape2.Left+(.Shape2.Width-RetTxtWidth('WПросмотрW')*2-20)/2,.Shape2.Top+.Shape2.Height+20,'печать','','DO chartoword1',39,RetTxtWidth('WПросмотрW'),'печать'
     DO addButtonOne WITH 'fSupl','butRetCh',.butPrnCh.Left+.butPrnCh.Width+20,.butPrnCh.Top,'возврат','','fSupl.Release',39,.butPrnCh.Width,'возврат'
     
     .Height=.Shape1.Height+.Shape2.Height+.butPrn.Height*2+100     
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*************************************************************************************************************************
PROCEDURE prnSprav1
PARAMETERS parLog
SELECT curSprav
DO procForPrintAndPreview WITH 'repSprav','',parLog,'spravToWord1'
*************************************************************************************************************************
PROCEDURE spravToWord1
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
     namedoc.tables(3).cell(3,1).Range.Text=ALLTRIM(people.fio) 
     namedoc.tables(3).cell(9,2).Range.Text=ALLTRIM(cursprav.txtsprav)
      
     namedoc.tables(3).cell(15,1).Range.Text='Справка выдана по состоянию на '+DTOC(ddspr)+'г.' 
     namedoc.tables(3).cell(18,1).Range.Text=ALLTRIM(boss.cspravdol)  
     namedoc.tables(3).cell(18,5).Range.Text=ALLTRIM(boss.cspravfio)  
ENDWITH   
objWord.Visible=.T.    
*************************************************************************************************************************
PROCEDURE charToWord1
#DEFINE wdWindowStateMaximize 1
pathdot=ALLTRIM(datset.pathword)+'char.dotx'
objWord=CREATEOBJECT('WORD.APPLICATION')
nameDoc=objWord.Documents.Add(pathdot)  
objWord.WindowState=wdWindowStateMaximize   
*nameDoc.ActiveWindow.View.ShowAll=0   
IF TYPE([nameDoc.formFields("cfio")])="O"
        nameDoc.FormFields("cfio").Result=cFio
ENDIF
IF TYPE([nameDoc.formFields("cage")])="O"
        nameDoc.FormFields("cage").Result=cAgeChar
ENDIF
IF TYPE([nameDoc.formFields("ceduc")])="O"
        nameDoc.FormFields("ceduc").Result=ceduc
ENDIF
IF TYPE([nameDoc.formFields("ccharw")])="O"
        nameDoc.FormFields("ccharw").Result=ALLTRIM(cursprav.txtchar)
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
************************************************************************************************
*                           Подтверждение категории
************************************************************************************************
PROCEDURE catconfirm
dEnd=CTOD('  .  .    ')
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .Caption='Список сотрудников для подтверждения категории'        
     DO addshape WITH 'fSupl',1,20,20,150,400,8     
     DO adLabMy WITH 'fSupl',1,'Дата',.Shape1.Top+20,.Shape1.Left+5,100,2,.T.,1 
     DO adtbox WITH 'fSupl',1,.lab1.Left+.lab1.Width+10,.Shape1.Top+20,RetTxtWidth('99/99/99999'),dHeight,'dEnd',.F.,.T.,.F.

     .lab1.Top=.txtBox1.Top+(.txtBox1.Height-.lab1.Height+2)  
     .lab1.Left=.Shape1.Left+(.Shape1.Width-.lab1.Width-.txtBox1.Width-10)/2
     .txtBox1.Left=.lab1.Left+.lab1.Width+10
     .Shape1.Height=.txtBox1.Height+40
     
     DO addButtonOne WITH 'fSupl','butPrn',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wсформироватьw')*2-15)/2,.Shape1.Top+.Shape1.Height+20,'сформировать','','DO prnconfirm',39,RetTxtWidth('wсформироватьw'),'сформировать' 
     DO addButtonOne WITH 'fSupl','butRet',.butPrn.Left+.butPrn.Width+15,.butPrn.Top,'возврат','','fSupl.Release',39,.butPrn.Width,'возврат'  
    
     DO addShapePercent WITH 'fSupl',.Shape1.Left,.butPrn.Top,.butPrn.Height,.Shape1.Width       
                        
     .Width=.Shape1.Width+40          
     .Height=.butPrn.Top+.butPrn.Height+10
     
ENDWITH 
DO pasteImage WITH 'fSupl'
fSupl.Show
******************************************************************************************************************************************
PROCEDURE prnconfirm
IF EMPTY(dEnd)
   RETURN
ENDIF 

IF USED('curPrn')
   SELECT curPrn
   USE
ENDIF
IF USED('curjobage')
   SELECT curjobage
   USE
ENDIF
IF USED('curDataKurs')
   SELECT curDataKurs
   USE 
ENDIF
SELECT * FROM datjob WHERE EMPTY(dateout).AND.INLIST(tr,1,3) INTO CURSOR curJobage READWRITE
SELECT curJobAge
INDEX ON STR(kodpeop,4)+STR(kse,4,2) TAG T1 DESCENDING 

SELECT * FROM datkurs INTO CURSOR curDataKurs READWRITE
SELECT curDataKurs
INDEX ON STR(nidpeop,5)+DTOS(perBeg) TAG T1 DESCENDING

SELECT * FROM people WHERE kval#0 INTO CURSOR curPrn READWRITE
ALTER TABLE curPrn ADD COLUMN npp N(3)
ALTER TABLE curPrn ADD COLUMN kp N(3)
ALTER TABLE curPrn ADD COLUMN kd N(3)
ALTER TABLE curPrn ADD COLUMN npodr C(100)
ALTER TABLE curPrn ADD COLUMN ndol C(100)
ALTER TABLE curPrn ADD COLUMN dnewkval D
ALTER TABLE curPrn ADD COLUMN khours N(5)
SELECT curprn
*INDEX ON DTOS(datts) TAG t1 
REPLACE kp WITH IIF(SEEK(STR(num,4),'curJobAge',1),curJobAge.kp,0),kd WITH curjobAge.kd,npodr WITH IIF(SEEK(kp,'sprpodr',1),sprpodr.namework,''),ndol WITH IIF(SEEK(kd,'sprdolj',1),sprdolj.name,'') ALL
REPLACE dnewkval WITH CTOD(STR(DAY(dkval),2)+'.'+STR(MONTH(dkval),2)+'.'+STR(YEAR(dkval)+5,4)) ALL 
DELETE FOR dnewkval>dend
nppcx=1
SCAN ALL 
     REPLACE npp WITH nppcx
     SELECT curDataKurs
     SUM khours FOR nidpeop=curprn.nid.AND.perEnd>=curprn.dkval.AND.perEnd<=dend TO khours_cx
     SELECT curprn
     REPLACE khours WITH khours_cx
     nppcx=nppcx+1
ENDSCAN
GO TOP
DO startPrnToExcel WITH 'fSupl'
   
objExcel=CREATEOBJECT('EXCEL.APPLICATION')
excelBook=objExcel.workBooks.Add(-4167)  
WITH excelBook.Sheets(1)
     .Columns(1).ColumnWidth=5
     .Columns(2).ColumnWidth=30
     .Columns(3).ColumnWidth=30
     .Columns(4).ColumnWidth=30
     .Columns(5).ColumnWidth=10
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8
     .Columns(8).ColumnWidth=40
     .Columns(9).ColumnWidth=15
     .Columns(10).ColumnWidth=8
     .Columns(11).ColumnWidth=8
          
     .cells(2,1).Value='№'              
     .cells(2,2).Value='ФИО сотрудника'              
     .cells(2,3).Value='подразделение'
     .cells(2,4).Value='должность'
     .cells(2,5).Value='категория'
     .cells(2,6).Value='дата присвоения' 
     .cells(2,6).WrapText=.T.                                   
     .cells(2,7).Value='дата подтвержд.' 
     .cells(2,7).WrapText=.T.                                   
     .cells(2,8).Value='курсы'
     .cells(2,9).Value='период'                                              
     .cells(2,10).Value='часы'
     .cells(2,11).Value='часы всего'
     .cells(2,11).WrapText=.T.                                   
     .Range(.Cells(1,1),.Cells(1,10)).Select
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter       
          .WrapText=.T.
          .Value='Список сотрудников для подтверждения категории на '+DTOC(dEnd)
     ENDWITH  
     .Range(.Cells(2,1),.Cells(2,11)).Select
     objExcel.Selection.HorizontalAlignment=xlCenter       

     numberRow=3
     yearMonth=''
     SELECT curprn
     DO storezeropercent
     SCAN ALL        
          .cells(numberRow,1).Value=npp
          .cells(numberRow,2).Value=fio
          .cells(numberRow,3).Value=npodr
          .cells(numberRow,4).Value=ndol
          .cells(numberRow,5).Value=IIF(SEEK(kval,'sprkval',1),sprkval.name,'')
          .cells(numberRow,6).Value=IIF(!EMPTY(dkval),dkval,'')
          .cells(numberRow,7).Value=IIF(!EMPTY(dnewkval),dnewkval,'')
          .cells(numberRow,11).Value=IIF(khours#0,kHours,'')
          SELECT curDataKurs
          SEEK STR(curprn.nid,5)
          IF FOUND()          
             SCAN WHILE nidpeop=curprn.nid                 
                  IF perEnd>=curprn.dkval.AND.perEnd<=dend
                     .cells(numberRow,8).Value=ALLTRIM(namekurs)+' '+ALLTRIM(nameschool)
                     .cells(numberRow,8).WrapText=.T.
                     .cells(numberRow,9).Value=DTOC(perbeg)+' - '+DTOC(perend)
                     .cells(numberRow,10).Value=IIF(kHours#0,khours,'')
                     numberRow=numberRow+1         
                  ENDIF    
             ENDSCAN      
          ELSE 
             numberRow=numberRow+1         
          ENDIF 
          SELECT curprn
          DO fillpercent WITH 'fSupl'
     ENDSCAN
    .Range(.Cells(1,1),.Cells(numberRow-1,11)).Select
    WITH objExcel.Selection
         .Borders(xlEdgeLeft).Weight=xlThin
         .Borders(xlEdgeTop).Weight=xlThin            
         .Borders(xlEdgeBottom).Weight=xlThin
         .Borders(xlEdgeRight).Weight=xlThin
         .Borders(xlInsideVertical).Weight=xlThin
         .Borders(xlInsideHorizontal).Weight=xlThin
         .VerticalAlignment=1
         .Font.Name='Times New Roman'   
         .Font.Size=10
    ENDWITH 
    .Range(.Cells(1,1),.Cells(1,10)).Select
ENDWITH 
ON ERROR DO erSup
DO endPrnToExcel WITH 'fSupl'
ON ERROR
objExcel.Visible=.T. 
 *******************************************************************************
PROCEDURE procDimFlt
PUBLIC dimOption(3)
STORE .F. TO dimOption
*dimOption(1) - подразделение
*dimOption(2) - должность
*dimOption(3) - персонал

SELECT * FROM sprpodr INTO CURSOR dopPodr READWRITE
SELECT doppodr
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprdolj INTO CURSOR dopDolj READWRITE
SELECT dopDolj
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

SELECT * FROM sprkat INTO CURSOR dopKat READWRITE
SELECT dopKat
REPLACE fl WITH .F.,otm WITH '' ALL
INDEX ON name TAG T1

PUBLIC onlyPodr,onlyDol,fltch,fltPodr,fltKat,fltDolj,kvoPodr,kvoDolj,kvoKat,totKvo,tableItem
STORE .F. TO onlyPodr,onlyDol
STORE '' TO fltch,fltPodr,fltKat,fltDolj,tableItem
STORE 0 TO kvoPodr,kvoDolj,kvoKat,totKvo
************************************************************************
PROCEDURE procObjFlt
DO addshape WITH 'fSupl',1,10,10,150,500,8
WITH fSupl       
     DO adCheckBox WITH 'fSupl','checkPodr','подразделение',.Shape1.Top+10,.Shape1.Left+5,150,dHeight,'dimOption(1)',0,.T.,"DO validCheckItem WITH 'dopPodr.otm,name','dopPodr','DO returnToPrnPodr WITH .T.','DO returnToPrnPodr WITH .F.'"         
     DO adCheckBox WITH 'fSupl','checkDolj','должность',.checkPodr.Top,.Shape1.Left,150,dHeight,'dimOption(2)',0,.T.,"DO validCheckItem WITH 'dopDolj.otm,name','dopDolj','DO returnToPrnDolj WITH .T.','DO returnToPrnDolj WITH .F.'"  
     DO adCheckBox WITH 'fSupl','checkKat','персонал',.checkPodr.Top,.checkDolj.Left,150,dHeight,'dimOption(3)',0,.T.,"DO validCheckItem WITH 'dopKat.otm,name','dopKat','DO returnToPrnKat WITH .T.','DO returnToPrnKat WITH .F.'"  
     .checkPodr.Left=.Shape1.Left+(.Shape1.Width-.checkPodr.Width-.checkDolj.Width-.checkKat.Width-40)/2
     .checkDolj.Left=.checkPodr.Left+.checkPodr.Width+20                                
     .checkKat.Left=.checkDolj.Left+.checkDolj.Width+20                                
     .Shape1.Height=.CheckPodr.Height+20  
ENDWITH 
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
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
     .listBox1.Visible=.F. 
     dimOption(1)=IIF(kvoPodr>0,.T.,.F.)
     .checkPodr.Caption='подразделение'+IIF(kvoPodr#0,'('+LTRIM(STR(kvoPodr))+')','') 
     .Refresh
ENDWITH  
************************************************************************************************************************
PROCEDURE returnToPrnDolj
PARAMETERS parRet
kvoDolj=0
IF parRet
   SELECT dopDolj
   fltDolj=''
   onlyDolj=.F.
   SCAN ALL
        IF fl 
           fltDolj=fltDolj+','+LTRIM(STR(kod))+','
           onlyDolj=.T.
           kvoDolj=kvoDolj+1
        ENDIF 
   ENDSCAN
ELSE 
   strDolj=''
   onlyDolj=.F.
   SELECT dopDolj
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
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
     .listBox1.Visible=.F.    
     dimOption(2)=IIF(kvoDolj>0,.T.,.F.)
     .checkDolj.Caption='должность'+IIF(kvoDolj#0,'('+LTRIM(STR(kvoDolj))+')','') 
     .Refresh
ENDWITH 
************************************************************************************************************************
PROCEDURE returnToPrnkat
PARAMETERS parRet
kvokat=0
IF parRet
   SELECT dopkat
   fltkat=''
   onlykat=.F.
   SCAN ALL
        IF fl 
           fltkat=fltkat+','+LTRIM(STR(kod))+','
           onlykat=.T.
           kvokat=kvokat+1
        ENDIF 
   ENDSCAN
ELSE 
   strkat=''
   onlykat=.F.
   SELECT dopkat
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
     .Shape11.Visible=.F.
     .Shape12.Visible=.F.
     .lab25.Visible=.F.
     .listBox1.Visible=.F.    
     dimOption(3)=IIF(kvokat>0,.T.,.F.)
     .checkkat.Caption='персонал'+IIF(kvokat#0,'('+LTRIM(STR(kvokat))+')','') 
     .Refresh
ENDWITH 
*********************************************************************************************************
PROCEDURE validCheckItem
PARAMETERS parSource,parClick,parReturnTrue,parReturnFalse
WITH fSupl
     .SetAll('Visible',.F.,'LabelMy')
     .SetAll('Visible',.F.,'MyTxtBox') 
     .SetAll('Visible',.F.,'comboMy')
     .SetAll('Visible',.F.,'shapeMy')
     .SetAll('Visible',.F.,'MyOptionButton')
     .SetAll('Visible',.F.,'MySpinner') 
     .SetAll('Visible',.F.,'MyContLabel')   
     .SetAll('Visible',.F.,'MyCommandButton')       
     .cont11.Visible=.T.
     .cont12.Visible=.T.
     .listBox1.Visible=.T.     
     tableItem=parClick
     .listBox1.RowSource=parSource    
     .listBox1.procForClick='DO clickListItem WITH tableItem'
     .listBox1.procForKeyPress="DO keyPressItem WITH 'DO clickListItem WITH tableItem'"
  
     .cont11.procForClick=parReturnTrue
     .cont12.procForClick=parReturnFalse
     .Refresh
ENDWITH
*************************************************************************************************************************
PROCEDURE keyPressItem
PARAMETERS parProc
DO CASE
   CASE LASTKEY()=27     
   CASE LASTKEY()=13
        &parProc
ENDCASE  
*************************************************************************************************************************
PROCEDURE clickListItem
PARAMETERS parTbl
SELECT &parTbl
rrec=RECNO()
REPLACE fl WITH IIF(fl,.F.,.T.)
REPLACE otm WITH IIF(fl,' • ','')
GO rrec
fSupl.listBox1.SetFocus
GO rrec
************************************************************************************************************************
PROCEDURE applyflt
IF kvoPodr>0
   DELETE FOR !(','+LTRIM(STR(kp))+','$fltPodr)
ENDIF
IF kvoDolj>0
   DELETE FOR !(','+LTRIM(STR(kd))+','$fltDolj)   
ENDIF
IF kvoKat>0
   DELETE FOR !(','+LTRIM(STR(kat))+','$fltKat)
ENDIF