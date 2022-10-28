IF !USED('datImage')
   USE datImage IN 0
ENDIF 
SELECT * FROM datImage WHERE kodPeop=people.num INTO CURSOR curDatImage READWRITE 
GO TOP 
imRec=0
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     .procExit='DO exitDocImage'
     .Caption=ALLTRIM(people.fio)+'- сохраненные изображения документов'  
     .AddObject('supImage','myImage')      
     .AddObject('grdImage','gridMyNew')
     .scrollBars=3
     .Height=SYSMETRIC(22)
     .Width=SYSMETRIC(21) && 1024 
    
     WITH .grdImage
          .Left=0
          .Top=0          
          .Height=.Parent.Height-dHeight-35      
          .Width=450        
          .ScrollBars=2       
          .RecordSourceType=1
          .RecordSource='curDatImage'
          .ColumnCount=0
          DO addColumnToGrid WITH 'fSupl.grdImage',2
          .Column1.ControlSource='curDatImage.namedoc'    
          *.Column2.ControlSource=''         
          .Columns(.ColumnCount).Width=0
          *.Column2.Width=40
          .Column1.Width=.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='документ'    
         * .Column2.Header1.Caption=''            
          .SetAll('BOUND',.F.,'Column')  
                    
        *  .Column2.AddObject('ImagePdf','Image')
        *  WITH .Column2.ImagePdf
        *       .Picture='pdf.ico'
        *       .Visible=.T.
        *       .Width=fSupl.grdImage.Column2.Width
        *       .Height=fSupl.grdImage.rowHeight
        *       .Top=0
        *       .Left=0   
        *       .Width=16
        *       .Height=16
        *       .Stretch=1    
        *  ENDWITH 
              
        * .column2.CurrentControl='ImagePdf'
        * .column2.DynamicCurrentControl='IIF(LOWER(ALLTRIM(RIGHT(curDatImage.pathImage,3)))="pdf","fSupl.grdImage.Column2.imagePdf","") '
         
          .Visible=.T. 
         * .Column2.Sparse=.F.  
          .procAfterRowColChange='DO changeImage'            
     ENDWITH                                     
     DO gridSizeNew WITH 'fSupl','grdImage','shapeingrid' 
             
    
     supWidth=.Width-.grdImage.Width-10
     supHeight=.Height
     supLeft=.grdImage.Left+.grdImage.Width+5        
     DO addcontlabel WITH 'fSupl','contAdd',10,.grdImage.Top+.grdImage.Height+10, RetTxtWidth('WдобавитьW'),dHeight+5,'добавить','DO procAddImage'   
     DO addcontlabel WITH 'fSupl','contRead',.contAdd.Left+.contAdd.Width+10,.contAdd.Top,.contAdd.Width,dHeight+5,'редакция','DO procReadImage'  
     DO addcontlabel WITH 'fSupl','contDel',.contRead.Left+.contRead.Width+10,.contAdd.Top,.contAdd.Width,dHeight+5,'удаление','DO delDocImage','Возврат'   
     DO addcontlabel WITH 'fSupl','contRet',.contDel.Left+.contDel.Width+10,.contAdd.Top,.contAdd.Width,dHeight+5,'возврат','DO exitDocImage','Возврат' 
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
     .maxButton=.T.
     .minButton=.T.
      WITH .supImage
          .Visible=.F.
          .Width=256
          .Height=256
          .Top=(SupHeight-.Height)/2
          .Left=supLeft+(.Parent.Width-fSupl.grdImage.Width-.Width)/2         
     ENDWITH   
      DO adLabMy WITH 'fSupl',1,'два раза щёлкните мышью для открытия файла' ,.supImage.Top+.supImage.Height+10,supLeft+(.Width-supLeft-RetTxtWidth('два раза щёлкните мышью для открытия файла'))/2,400,2,.T.
      .lab1.Left=supLeft+(.Width-supLeft-.lab1.Width)/2     
      .lab1.Visible=.F.
ENDWITH 
SELECT curDatImage
*fSupl.grdImage.Columns(fSupl.grdImage.ColumnCount).SetFocus
fSupl.Show
***********************************************************************************
PROCEDURE changeImage
SELECT curDatImage
IF imRec#RECNO() 
   pathmem=FULLPATH('kadry.fxp')
   DO CASE
      CASE LOWER(RIGHT(ALLTRIM(curDatImage.pathImage),3))='doc'
           pathmem=LEFT(pathmem,LEN(pathmem)-9)+'datDoc'+LTRIM(STR(RECNO()))+'.doc'
          * COPY MEMO pDoc TO &pathmem   
           fSupl.formImage.Visible=.F.
           fSupl.SupImage.Visible=.T. 
           fSupl.lab1.Visible=.T.         
           fSupl.supImage.procForDoublClick='DO openDoc WITH 1'
          
      CASE LOWER(RIGHT(ALLTRIM(curDatImage.pathImage),3))='pdf'           
           fSupl.formImage.Visible=.F.
           fSupl.SupImage.Visible=.T.
           fSupl.lab1.Visible=.T.         
           fSupl.supImage.Picture='pdf.jpg'    
           fSupl.supImage.procForDoublClick='DO openDoc WITH 2'     
               
      OTHERWISE 
           pathmem=LEFT(pathmem,LEN(pathmem)-9)+'datPic'+LTRIM(STR(people.num))+'.pic'
           COPY MEMO curDatImage.pDoc TO &pathmem            
           fSupl.SupImage.Visible=.F.
           fSupl.lab1.Visible=.F.         
           WITH fSupl.formImage 
                .Picture=pathmem
                .Visible=.T.
                DO CASE
                   CASE curDatImage.pwidth<=supWidth.AND.curDatImage.pHeight<=supHeight
                        .Width=curDatImage.pWidth
                        .Height=curDatImage.pHeight
                        .Left=IIF(.Width=supWidth,supLeft,supLeft+(.Parent.Width-supLeft-.Width)/2)
                        .Top=IIF(.Height=supHeight,0,(.Parent.Height-.Height)/2)
                   CASE curDatImage.pHeight>supHeight.AND.curDatImage.pHeight>=curDatImage.pWidth
                        .Height=supHeight
                        .Width =.Height/(curDatImage.pHeight/curDatImage.pWidth) 
                        .Top=0
                        .Left=IIF(.Width=supWidth,supLeft,supLeft+(.Parent.Width-supLeft-.Width)/2)
                   CASE curDatImage.pWidth>supWidth.AND.curDatImage.pWidth>=curDatImage.pHeight
                        .Width=supWidth
                        .Height =.Width/(curDatImage.pWidth/curDatImage.pHeight) 
                        .Left=supLeft
                        .Top=IIF(.Height=supHeight,0,(.Parent.Height-.Height)/2)                
                   OTHERWISE 
                        .Width=supWidth
                        .Height=supHeight      
                        .Left=IIF(.Width=supWidth,supLeft,supLeft+(.Parent.Width-supLeft-.Width)/2)
                        .Top=IIF(.Height=supHeight,0,(.Parent.Height-.Height)/2)
                ENDCASE        
           ENDWITH 
           DELETE FILE &pathmem       
   ENDCASE
    fSupl.Refresh
    imRec=RECNO()         
       
  
ENDIF    
************************************************************************************
PROCEDURE procAddImage
nameNewDoc='Документ 1'
newPathImage=''
newPathImage=GETFILE('jpg','выбор файла','выбрать',0)
*newPathImage=GETPICT()
IF EMPTY(newPathImage)
   RETURN
ENDIF
fSuplNew=CREATEOBJECT('FORMSUPL')
WITH fSuplNew
     .Caption='Добавление документа'
      DO addContFormNew WITH 'fSuplNew','contf',10,10,RetTxtWidth('WНаименование док-таW'),dHeight,'файл',0,.F.,'DO selectImage'  
*      DO adtBoxAsCont WITH 'fSuplNew','contf',10,10,RetTxtWidth('WНаименование док-таW'),dHeight,'файл',0,1  
      DO adTboxNew WITH 'fSuplNew','tBoxF',.contF.Top,.contF.Left+.contF.Width-1,300,dHeight,'newPathImage',.F.,.F.,0 
      DO adtBoxAsCont WITH 'fSuplNew','contFn',.contF.Left,.contF.Top+.contF.Height-1,.contF.Width,dHeight,'наименование док-та',0,1
      DO adTboxNew WITH 'fSuplNew','tBoxFn',.contFn.Top,.tBoxF.Left,.tBoxF.Width,dHeight,'nameNewDoc',.F.,.T.,0 
      .Width=.contf.Width+.tBoxF.Width+20
      DO addcontlabel WITH 'fSuplNew','butSave',(.Width-RetTxtWidth('wсохранитьw')*2-30)/2,.contFn.Top+.contFn.Height+20,RetTxtWidth('wсохранитьw'),dHeight+5,'сохранить','DO writeImage WITH .T.'      

      DO addcontlabel WITH 'fSuplNew','butReturn',.butSave.Left+.butSave.Width+15,.butSave.Top,;
        .butSave.Width,dHeight+5,'возврат','fSuplNew.Release','возврат'
      .Height=.contF.Height*2+.butSave.Height+50  
ENDWITH 
DO pasteImage WITH 'fSuplNew'
fSuplNew.Show
************************************************************************************
PROCEDURE writeImage
PARAMETERS parAp
IF parAp
   _screen.Addobject('OBJPIC','IMAGE')
   _screen.ObjPic.Picture=newPathImage       
   var_Width=_screen.ObjPic.WIDTH
   var_Height=_screen.ObjPic.Height
   _screen.Removeobject('ObjPic')              
   SELECT datImage
   APPEND BLANK 
   REPLACE kodPeop WITH people.num,nameDoc WITH nameNewDoc
   APPEND MEMO pdoc from &newPathImage
   REPLACE pwidth WITH var_width,pheight WITH var_height,pathImage WITH newPathImage
   SELECT curDatImage 
   APPEND BLANK
   REPLACE kodPeop WITH people.num,nameDoc WITH nameNewDoc
   APPEND MEMO pdoc from &newPathImage
   REPLACE pwidth WITH var_width,pheight WITH var_height,pathImage WITH newPathImage
ELSE 
   SELECT datImage
   LOCATE FOR kodPeop=curDatImage.kodPeop.AND.nameDoc=curDatImage.nameDoc.AND.pWidth=curDatImage.pWidth.AND.pHeight=curDatImage.pHeight
   REPLACE nameDoc WITH nameNewDoc
   IF !EMPTY(newPathImage)
      _screen.Addobject('OBJPIC','IMAGE')
      _screen.OBJPIC.PICTURE=newPathImage       
      var_Width=_screen.ObjPic.Width
      var_Height=_screen.ObjPic.Height
      _screen.Removeobject('ObjPic')     
      APPEND MEMO pdoc from &newPathImage OVERWRITE
      REPLACE pwidth WITH var_width,pheight WITH var_height,pathImage WITH newPathImage      
   ENDIF 
   SELECT curDatImage 
   REPLACE nameDoc WITH nameNewDoc
   IF !EMPTY(newPathImage)    
      APPEND MEMO pdoc from &newPathImage OVERWRITE
      REPLACE pwidth WITH datImage.pWidth,pheight WITH datImage.pHeight,pathImage WITH newPathImage
   ENDIF 
ENDIF    
fSuplNew.Release
imRec=0
fSupl.Refresh
************************************************************************************
PROCEDURE procReadImage
oldPathImage=curDatImage.pathImage
newPathImage=''
oldNameDoc=curDatImage.nameDoc
oldpWidth=curDatImage.pWidth
oldpHeight=curDatImage.pHeight
nameNewDoc=curDatImage.nameDoc
fSuplNew=CREATEOBJECT('FORMSUPL')
WITH fSuplNew
     .Caption='Редактирование документа'
      DO addContFormNew WITH 'fSuplNew','contf',10,10,RetTxtWidth('WНаименование док-таW'),dHeight,'файл',0,.F.,'DO selectImage'  
      DO adTboxNew WITH 'fSuplNew','tBoxF',.contF.Top,.contF.Left+.contF.Width-1,300,dHeight,'oldPathImage',.F.,.F.,0 
      DO adtBoxAsCont WITH 'fSuplNew','contFn',.contF.Left,.contF.Top+.contF.Height-1,.contF.Width,dHeight,'наименование док-та',0,1
      DO adTboxNew WITH 'fSuplNew','tBoxFn',.contFn.Top,.tBoxF.Left,.tBoxF.Width,dHeight,'nameNewDoc',.F.,.T.,0 
      .Width=.contf.Width+.tBoxF.Width+20
      DO addcontlabel WITH 'fSuplNew','butSave',(.Width-RetTxtWidth('wсохранитьw')*2-30)/2,.contFn.Top+.contFn.Height+20,RetTxtWidth('wсохранитьw'),dHeight+5,'сохранить','DO writeImage WITH .F.'      

      DO addcontlabel WITH 'fSuplNew','butReturn',.butSave.Left+.butSave.Width+15,.butSave.Top,;
        .butSave.Width,dHeight+5,'возврат','fSuplNew.Release','возврат'
      .Height=.contF.Height*2+.butSave.Height+50  
ENDWITH 
DO pasteImage WITH 'fSuplNew'
fSuplNew.Show
************************************************************************************
PROCEDURE selectImage
newPathImage=GETFILE('jpg','выбор файла','выбрать',0)
IF !EMPTY(newPathImage)
   fSuplNew.tBoxF.ControlSource='newPathImage'
ENDIF
fSuplNew.Refresh
************************************************************************************
PROCEDURE delDocImage
log_del=.F.
fDel=CREATEOBJECT('formsupl')
WITH fDel
     .Caption='Удаление записи'           
     DO addShape WITH 'fdel',1,20,20,dHeight,350,8   
     DO adLabMy WITH 'fdel',1,'Удалить выбранную запись?',.Shape1.Top+10,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fdel',2,'Для подтверждения намерений поставьте отметку',.Lab1.Top+.Lab1.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     DO adLabMy WITH 'fdel',3,"в окошке 'подтверждение намерений'",.Lab2.Top+.Lab2.Height+5,.Shape1.Left,.Shape1.Width,2,.F.
     .Shape1.Height=.lab1.Height*3+30       
      DO adCheckBox WITH 'fDel','check1','подтверждение намерений',.Shape1.Top+.Shape1.Height+20,.Shape1.Left,150,dHeight,'log_del',0
     .check1.Left=.Shape1.Left+(.Shape1.Width-.check1.Width)/2
     DO addcontlabel WITH 'fdel','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wwудалитьww')*2-20)/2,.check1.Top+.check1.Height+20,;
       RetTxtWidth('wwудалитьww'),dHeight+3,'удалить','DO delRecImage'
     DO addcontlabel WITH 'fdel','cont2',.Cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+3,'отмена','fdel.Release'
     .Height=.Shape1.Height+.cont1.Height+.check1.Height+80  
     .Width=.Shape1.Width+40 
ENDWITH
DO pasteImage WITH 'fDel'
fDel.Show
*************************************************************************************************************************
PROCEDURE delRecImage
IF !log_del
   RETURN
ENDIF
fDel.Release
SELECT datImage
LOCATE FOR kodPeop=curDatImage.kodPeop.AND.nameDoc=curDatImage.nameDoc.AND.pWidth=curDatImage.pWidth.AND.pHeight=curDatImage.pHeight
DELETE 
SELECT curDatImage
DELETE
GO TOP
fSupl.Refresh
***********************************************************************************
PROCEDURE openDoc
PARAMETERS par1
DO CASE
   CASE par1=1
        pathWord=FULLPATH('kadry.fxp')
        pathWord=LEFT(pathWord,LEN(pathWord)-9)+'datPdf'+LTRIM(STR(RECNO()))+'.pdf'
        *objWord=CREATEOBJECT('WORD.APPLICATION')
        *pathdot=dim_word(2)+'kontrakt.dot'
        *nameDoc=objWord.Documents.Add(pathdot)          
   CASE par1=2   
        ON ERROR DO erImage    
        pathPdf=FULLPATH('kadry.fxp')
        pathPdf=LEFT(pathPdf,LEN(pathPdf)-9)+'datPdf'+LTRIM(STR(RECNO()))+'.pdf'
        COPY MEMO pDoc TO &pathPdf   
        oPdf=CREATEOBJECT("wscript.shell") 
        oPdf.RUN(pathPdf)  
        ON ERROR
ENDCASE
**********************************************************************************
PROCEDURE erImage
***********************************************************************************
PROCEDURE exitDocImage
SELECT datImage
USE
SELECT people
ON ERROR DO erImage
pathPdf=FULLPATH('kadry.fxp')
pathPdf=LEFT(pathPdf,LEN(pathPdf)-9)+'datPdf*.*'
pathPic=LEFT(pathPdf,LEN(pathPdf)-9)+'datPic*.pic'
DELETE FILE &pathPdf
DELETE FILE &pathPic
ON ERROR
fSupl.Release