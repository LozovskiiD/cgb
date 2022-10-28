************************************************************************************************************************
*                        Основная для настройки ведомости
************************************************************************************************************************
DO CASE 
   CASE datset.nformat=1  && A4
        IF !USED('fondprn4')
           IF FILE(pathcur+'fondprn4.dbf')
              fUse=pathcur+'fondprn4.dbf' 
              USE &fUse ORDER 1 IN 0 ALIAS fondprn
           ELSE
              USE fondprn4 ORDER 1 IN 0 ALIAS fondprn
           ENDIF
        ENDIF 
   CASE datset.nformat=2 && A3
        IF !USED('fondprn3')
           IF FILE(pathcur+'fondprn3.dbf')
              fUse=pathcur+'fondprn3.dbf' 
              USE &fUse ORDER 1 IN 0 ALIAS fondprn
           ELSE
              USE fondprn3 ORDER 1 IN 0 ALIAS fondprn
           ENDIF
        ENDIF 
ENDCASE 

DIMENSION dimFormat(2)
currentFormat=datset.nformat
dimFormat(1)=IIF(currentFormat=1,1,0)
dimFormat(2)=IIF(currentFormat=2,1,0)
IF !USED('formbase')
   USE formbase IN 0
ENDIF 

IF !USED('tarifset')
   USE tarifset IN 0
ENDIF 
SELECT fondprn
SET FILTER TO 
GO TOP

fdop=Createobject('Formspr')
WITH fdop
     .Caption='Настройка печати'
     .AddProperty('log_ap',.F.)
     .procexit='DO returnsetupved'
     
     DO addButtonOne WITH 'fDop','menuCont1',10,5,'новая','pencila.ico','Do inputrecvedprn WITH .T.',39,RetTxtWidth('удаление')+44,'новая'  
     DO addButtonOne WITH 'fDop','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'редакция','pencil.ico','Do inputrecvedprn WITH .F.',39,.menucont1.Width,'редакция'   
     DO addButtonOne WITH 'fDop','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','Do delcol',39,.menucont1.Width,'удаление'       
     DO addButtonOne WITH 'fDop','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'возврат','undo.ico','DO returnsetupved' ,39,.menucont1.Width,'возврат'                       
     
     DO addOptionButton WITH 'fDop',1,'Формат A4',.menucont1.Top,.menucont4.Left+.menucont4.Width+30,'dimFormat(1)',0,'DO procvalidformat WITH 1',.T.
     .Option1.Top=.menucont1.Top+(.menucont1.height-.option1.height)/2
     DO addOptionButton WITH 'fDop',21,'Формат A3',.option1.Top,.option1.Left+.option1.Width+10,'dimFormat(2)',0,'DO procvalidformat WITH 2',.T.

   
     DO addmenureadspr WITH 'fdop','DO writevedrec WITH .T.','DO writevedrec WITH .F.'
     WITH .fGrid     
          .Top=.Parent.menucont1.Top+.Parent.menucont1.Height+5
          .Left=0
          .Width=.Parent.Width/2    
          .Height=.Parent.Height-.Parent.menucont1.Height-5        
          DO addColumnToGrid WITH 'fDop.fGrid',7 
          .RecordSourceType=1
          .RecordSource='fondprn'
          .Column1.ControlSource='fondprn.exprlab'
          .Column2.ControlSource='fondprn.logved' 
          .Column3.ControlSource='fondprn.ncol' 
          .Column4.ControlSource="IIF(fondprn.log_sv,'•','')"
          .Column5.ControlSource="IIF(fondprn.log_kv,'•','')"
          .Column6.ControlSource="IIF(fondprn.itog,'•','')"
          .Column1.Alignment=0
          .Column2.Alignment=2
          .Column3.Alignment=1
          .Column4.Alignment=2
          .Column5.Alignment=2
          .Column6.Alignment=2
          .Column1.Header1.Caption='Наименование колонки'
          .Column2.Header1.Caption='!'       
          .Column3.Header1.Caption='Номер'       
          .Column4.Header1.Caption='Св'
          .Column5.Header1.Caption='Кв'
          .Column6.Header1.Caption='Итог'
          .Column2.Width=RetTxtWidth('www')
          .Column3.Width=RetTxtWidth(' 1234 ')
          .Column4.Width=RetTxtWidth(' WW ')
          .Column5.Width=RetTxtWidth(' WW ')
          .Column6.Width=RetTxtWidth(' WW ')
          .Column7.Width=0
          .SetAll('Enabled',.F.,'ColumnMy')
          .Column2.Enabled=.T.
          .Column7.Enabled=.T.
          .Column7.ReadOnly=.T.
          .Column1.Width=.Width-.column2.Width-.column3.Width-.column4.Width-.Column5.Width-.Column6.Width-SYSMETRIC(5)-13-.ColumnCount         
          .colNesInf=2    
          .SetAll('Movable',.F.,'ColumnMy') 
          .SetAll('BOUND',.F.,'ColumnMy')    
          .Column2.Sparse=.F.
          .procAfterRowColChange='DO infocol'     
     ENDWITH 
     DO gridSizeNew WITH 'fdop','fGrid','shapeingrid',.T. 
     .fGrid.Column2.AddObject('checkColumn2','checkContainer')     
     .fGrid.Column2.checkColumn2.AddObject('checkMy','checkBox')
     .fGrid.Column2.CheckColumn2.checkMy.Visible=.T.
     .fGrid.Column2.CheckColumn2.checkMy.Caption=''
     .fGrid.Column2.CheckColumn2.checkMy.Left=7
     .fGrid.Column2.CheckColumn2.checkMy.BackStyle=0
     .fGrid.Column2.CheckColumn2.checkMy.ControlSource='fondPrn.logVed'   
     .fGrid.column2.CurrentControl='checkColumn2'
     fwidth=.Width-.fGrid.Width-5 

     DO addcontform WITH 'fdop','conthead',.fGrid.Left+.fGrid.Width+5,.fGrid.Top,fwidth,dheight,ALLTRIM(fondprn->rec) 
     *------------------------Информация по настройкам--------------------------------------------------------------
     SELECT tarifset
     GO TOP
     ord_ch=1 
     obj_ch=1 
     top_ch=.conthead.Top+.conthead.Height-1            &&Верхняя координата объектов                                 
     width_ch=fwidth/3*2                                &&Ширина объектов   
     leftcont_ch=.conthead.Left                         &&Левая координата для контейнеров           
     left_ch=.conthead.Left+fwidth/3-1                  &&Левая координата для tbox и проч.           

     DO WHILE !EOF()  
       obj_cont='lCont'+LTRIM(STR(ord_ch))
       .AddObject(obj_cont,'Container')
       WITH .&obj_cont
            .Height=dHeight
            .Width=fwidth/3  
            .Top=top_ch
            .Left=Leftcont_ch
            .BackStyle=0                  
            .Visible=.T.
       ENDWITH    
       obj_lab=''     
       DO adLabMy WITH 'fdop.&obj_cont',ord_ch,tarifset->name,3,4,fdop.&obj_cont..Width-4     
       IF !EMPTY(tarifset->procadd)
          procch=tarifset->procadd
          &procch
       ENDIF      
       ord_ch=ord_ch+1  
       obj_ch=obj_ch+1
       top_ch=fdop.&obj_cont..Top+dHeight-1     
       SKIP 
    ENDDO  
    .AddObject('grdved','GridMy')
    WITH .grdved
         .Top=top_ch
         .Left=leftcont_ch
         .Width=fwidth    
         .Height=.Parent.Height-top_ch
         .FontSize=dFontSize
         .ScrollBars=2
         .ColumnCount=3     
         *DO addColumnToGrid WITH 'fDop.grdVed',3 
         .RecordSourceType=1
         .RecordSource='formbase'
         .Column1.ControlSource='formbase->name'   
         .Column2.ControlSource="IIF(formbase->log_prn,'•','')"    
         .Column1.Alignment=0 
         .Column2.Alignment=2      
         .Column1.Header1.Caption='Ведомость'
         .Column2.Header1.Caption='•'           
         .Column2.Width=RettxtWidth(' 1234 ')     
         .Column3.Width=0
         .Column1.Width=.Width-.column2.Width-SYSMETRIC(5)-13-3          
         .colNesInf=2    
         .SetAll('Movable',.F.,'Column') 
         .SetAll('BOUND',.F.,'Column')   
    ENDWITH
    DO myColumnTxtBox WITH 'fdop.grdved.column3','txtbox3','',.F.
    .grdved.column3.txtbox3.procForKeyPress='DO repotmved'
    DO gridSize WITH 'fdop','grdved','shapeingrid1' 
    
ENDWITH 
SELECT fondprn
fdop.Show
***************************************************************************************************************************
PROCEDURE procvalidformat
PARAMETERS par1
STORE 0 TO dimFormat
dimFormat(par1)=1
fDop.Refresh
IF par1#currentFormat
  SELECT datset
  REPLACE nformat WITH par1
  SELECT fondprn
  USE
  SELECT tarifset
  USE
  DO setupprn
ENDIF 
***************************************************************************************************************************
PROCEDURE selectDirForSave
PARAMETERS par1
newpathword=GETDIR('','','Укажите папку для сохранения',64)
dim_copy(par1)=IIF(!EMPTY(newpathword),newpathword,dim_copy(par1))
fDop.Refresh
**************************************************************************************************************************
PROCEDURE returnsetupved
SELECT tarifset
USE 
SAVE TO ved_set ALL LIKE ved_set
SELECT fondprn
USE
fdop.Release
****************************************************************************************************************************
*            Ввод редакция колонки в тарификации
****************************************************************************************************************************
PROCEDURE inputrecvedprn
PARAMETERS par_log
WITH fDop
     .log_ap=.F.
     .SetAll('Enabled',.T.,'Mytxtbox')
     .fGrid.Enabled=.F.
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .menuread.Visible=.T.
     .menuexit.Visible=.T.
     .contitog.SetAll('ReadOnly',.F.,'MyCheckBox')
     .contdouble.SetAll('ReadOnly',.F.,'MyCheckBox')
     SELECT fondprn
     IF par_log
        .log_ap=.T.
        APPEND BLANK 
     ENDIF
     SCATTER TO .dim_ap
     .nrec=RECNO()
     .Refresh
     .txtbox1.SetFocus
ENDWITH      
***************************************************************************************************************************
PROCEDURE writevedrec
PARAMETERS par_log
WITH fDop
     .SetAll('Enabled',.F.,'Mytxtbox')
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     .contitog.SetAll('ReadOnly',.T.,'MyCheckBox')
     .contdouble.SetAll('ReadOnly',.T.,'MyCheckBox')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     .fGrid.Columns(.fGrid.ColumnCount).ReadOnly=.T.
     .fGrid.Column2.Enabled=.T.
ENDWITH      
SELECT fondprn
GO fdop.nrec
IF !par_log    
   IF !fdop.log_ap
      GATHER FROM fdop.dim_ap   
   ELSE
      DELETE
      GO TOP
   ENDIF   
ELSE
   IF ncol#fdop.dim_ap(3)
      DO CASE 
         CASE ncol<fdop.dim_ap(3).OR.fdop.dim_ap(3)=0
              num_cx=ncol
              REPLACE ncol WITH 0
              REPLACE ncol WITH ncol+1 FOR ncol>=num_cx
              GO fdop.nrec
              REPLACE ncol WITH num_cx
              GO TOP
              num_new=1
              DO WHILE !EOF()
                 REPLACE ncol WITH num_new,expcollab WITH "'"+ALLTRIM(STR(num_new))+"'"
                 SKIP
                  num_new=num_new+1
              ENDDO
         CASE ncol>fdop.dim_ap(3)
              num_cx=ncol
              REPLACE ncol WITH 0
              REPLACE ncol WITH ncol-1 FOR ncol>=num_cx
              GO fdop.nrec
              REPLACE ncol WITH num_cx
              GO TOP
              num_new=1
              DO WHILE !EOF()
                 REPLACE ncol WITH num_new,expcollab WITH "'"+ALLTRIM(STR(num_new))+"'"
                 SKIP
                  num_new=num_new+1
              ENDDO               
      ENDCASE  
      GO fdop.nrec     
    ENDIF  
       REPLACE comlab WITH ALLTRIM(STR(ncol)),ncollab WITH ALLTRIM(STR(ncol)),expcollab WITH "'"+ALLTRIM(STR(ncol))+"'"     

ENDIF
fdop.Refresh
***************************************************************************************************************************
PROCEDURE pdoitog
fdop.AddObject('contitog','Container')
WITH fdop.contitog
     .Height=dHeight
     .Width=width_ch+1
     .Top=top_ch
     .Left=Left_ch
     .BackStyle=0                  
     .Visible=.T.
ENDWITH    
DO adcheckbox WITH 'fdop.contitog','checkido','',5,5,150,dHeight,'fondprn.log_sv',0,.T.
DO adcheckbox WITH 'fdop.contitog','checkpdo','',5,fdop.contitog.checkido.left+fdop.contitog.checkido.Width+10,150,dHeight,'fondprn.log_kv',0,.T.
DO adcheckbox WITH 'fdop.contitog','checkitog','',5,fdop.contitog.checkpdo.left+fdop.contitog.checkpdo.Width+10,150,dHeight,'fondprn.itog',0,.T.
fdop.contitog.SetAll('BackStyle',0,'MyCheckBox')
fdop.contitog.SetAll('ReadOnly',.T.,'MyCheckBox')
****************************************************************************************************************************
PROCEDURE doublecol    
fdop.AddObject('contdouble','Container')
WITH fdop.contdouble
     .Height=dHeight
     .Width=width_ch+1
     .Top=top_ch
     .Left=Left_ch
     .BackStyle=0                  
     .Visible=.T.
ENDWITH    
DO adcheckbox WITH 'fdop.contdouble','checkdouble','',5,5,150,dHeight,'fondprn.ldouble',0,.T.
fdop.contdouble.SetAll('BackStyle',0,'MyCheckBox')
fdop.contdouble.SetAll('ReadOnly',.T.,'MyCheckBox')
****************************************************************************************************************************
PROCEDURE infocol
fdop.conthead.Contlabel.Caption=ALLTRIM(fondprn->rec)
fdop.Refresh
*************************************************************************************************************************
*                Удаление колонки из тарификации
*************************************************************************************************************************
PROCEDURE delcol
SELECT fondprn 
DO createFormNew WITH .T.,'Удаление',RetTxtWidth('WWУдалить выбранную запись?WW',dFontName,dFontSize+1),;
  '130',RetTxtWidth('WWНетWW',dFontName,dFontSize+1),'Да','Нет',.F.,'DO deltarcol','nFormMes.Release',.F.,;
  'Удалить выбранную запись?',.F.,.T.

                             
**************************************************************************************************************************
*      Непосредственно удаление записи
**************************************************************************************************************************  
PROCEDURE deltarcol
SELECT fondprn
DELETE
GO TOP
num_new=1
DO WHILE !EOF()
   REPLACE ncol WITH num_new,expcollab WITH "'"+ALLTRIM(STR(num_new))+"'"
   SKIP
   num_new=num_new+1
ENDDO
nFormMes.Release
fdop.Refresh
*******************************************************************************************
PROCEDURE repotmved
IF LASTKEY()=13   
   REPLACE log_prn WITH IIF(log_prn,.F.,.T.)
ENDIF



