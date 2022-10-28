SELECT people
oldInd=SYS(21)
IF !USED('fltBase')
   USE fltbase IN 0
ENDIF
SELECT datjob
SET FILTER TO
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
GO TOP
SELECT tarfond
SET FILTER TO !EMPTY(plrep)
SELECT sprkoef
COUNT TO max_kf
GO TOP
DIMENSION name_kf(max_kf),s_kf(max_kf),kod_kf(max_kf)
FOR i=1 TO max_kf
    name_kf(i)=STR(kod,2)+' -'+STR(name,5,2)
    s_kf(i)=name
    kod_kf(i)=kod
    SKIP
ENDFOR
STORE '' TO sostavFlt
frep=CREATEOBJECT('FORMMY')
WITH frep
     .Caption='Ускоренная замена тарифов'   
     .procExit='DO procExitFastRep' 
     .AddProperty('log_fl',.F.)    
     .Addproperty('filter_ch','') 
     .Addproperty('filter_peop','')   
     .AddProperty('reptar','')     
     .AddProperty('newtar',0) 
     *-----------------------------Выбор тарифа------------------------------------------------------------------------------
     DO addcombomy WITH 'frep',1,10,5,dHeight,300,.T.,'','tarfond.rec',6,.F.,'DO tarifrefresh' 
     WITH .comboBox1     && тариф
          .SpecialEffect=1               
          .BackColor=.Parent.backColor  
          .nDisplayCount=RECCOUNT('tarfond')   
          .Visible=.T. 
     ENDWITH
         
     DO addButtonOne WITH 'fRep','menuCont1',.combobox1.Left+.combobox1.Width+3,5,'автозамена','pencila.ico','DO autorep',39,RetTxtWidth('по признаку')+44,'автозамена'  
     DO addButtonOne WITH 'fRep','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'по признаку','markreplace.ico','DO menupriz',39,.menucont1.Width,'замена по признаку'   
     DO addButtonOne WITH 'fRep','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'редакция','pencil.ico','DO readFiltr',39,.menucont1.Width,'редакция'       
     DO addButtonOne WITH 'fRep','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'поиск','find.ico','DO poiskFastRep',39,.menucont1.Width,'поиск'     
     DO addButtonOne WITH 'fRep','menuCont5',.menucont4.Left+.menucont3.Width+4,5,'фильтр','filter1.ico',"Do procFilterNew WITH 'fRep',2",39,.menucont1.Width,'фильтр'     
     DO addButtonOne WITH 'fRep','menuCont6',.menucont5.Left+.menucont5.Width+4,5,'возврат','undo.ico','DO procExitFastRep',39,.menucont1.Width,'возврат'               
          
     DO addButtonOne WITH 'fRep','butPoisk',.menucont1.Left,5,'поиск','find.ico','DO poiskFiltr',39,RetTxtWidth('возврат')+44,'поиск'       
     DO addButtonOne WITH 'fRep','butReturn',.butPoisk.Left+.butPoisk.Width+3,5,'возврат','undo.ico','DO returnFromReadFiltr',39,.butPoisk.Width,'возврат'      
     
*      .comboBox1.Top=.menucont1.Top+(.menucont1.Height-.comboBox1.Height)/2
     .butPoisk.Visible=.F.
     .butReturn.Visible=.F.
     *-----------------Grid для персонала----------------------------------------------------------------------------------
     .AddObject('grdpers','GridMy')
     WITH .grdpers
          .Top=frep.menucont1.Top+frep.menucont1.Height+5
          .Left=0
          .Width=frep.Width
          .Height=frep.height-frep.Top
          .ColumnCount=6
          .RecordSource='datjob'  
          .Column1.ControlSource='datJob.kodpeop'
         * .Column2.ControlSource="IIF(SEEK(datjob.kodpeop,'people',1),people.fio,'')"
          .Column2.ControlSource='datjob.fio'
          .Column3.ControlSource="IIF(SEEK(datjob.kp,'sprpodr',1),sprpodr.name,'')"
          .Column4.ControlSource="IIF(SEEK(datjob.kd,'sprdolj',1),sprdolj.name,'')"
          .Column5.ControlSource=''
          .Column1.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH(' 1234 ',dFontName,dFontSize)
          .Column5.Width=FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH('9999999999.99',dFontName,dFontSize)
          .Column3.Width=(.Width-.Column1.Width-.Column5.Width)/3
          .Column4.Width=.Column3.Width
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(15)-13-6
          .Column6.Width=0
          .ScrollBars=2     
          .Column1.ReadOnly=.T.
          .Column2.Enabled=.F. 
          .Column3.Enabled=.T.
          .Column4.Enabled=.T.   
          .Column2.Alignment=0
          .Column3.Alignment=0
          .Column4.Alignment=0
          .Column5.Alignment=1
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Фамилия Имя Отчество'
          .Column3.Header1.Caption='Подразделение'
          .Column4.Header1.Caption='Должность'
          .Column5.Header1.Caption='Тариф'
          .colNesInf=2           
     ENDWITH
     DO myColumnTxtBox WITH 'frep.grdpers.column5','txtbox5','',.F.
     DO gridSize WITH 'frep','grdpers','shapeingrid'   
     DO addcontmy WITH 'frep','cont1',.grdpers.Left+13,frep.grdpers.Top+2,.grdpers.Column1.Width-3,.grdpers.HeaderHeight-3,'',"DO clickCont WITH 'frep','frep.cont1','datjob',1"
     DO addcontmy WITH 'frep','cont2',.cont1.Left+.Grdpers.Column1.Width+1,.grdpers.Top+2,.grdpers.Column2.Width-3,.grdpers.HeaderHeight-3,'',"DO clickCont WITH 'frep','frep.cont2','datjob',5"
     DO addcontmy WITH 'frep','cont3',.cont2.Left+.Grdpers.Column2.Width+1,.grdpers.Top+2,.grdpers.Column3.Width-3,.grdpers.HeaderHeight-3,''   
     DO addcontmy WITH 'frep','cont4',.cont3.Left+.Grdpers.Column3.Width+1,.grdpers.Top+2,.grdpers.Column4.Width-3,.grdpers.HeaderHeight-3,''  
     DO addcontmy WITH 'frep','cont5',.cont4.Left+.Grdpers.Column4.Width+1,.grdpers.Top+2,.grdpers.Column5.Width-3,.grdpers.HeaderHeight-3,''     
     .cont1.SpecialEffect=1 
     SELECT datjob
     .Grdpers.Column1.SetFocus
ENDWITH
frep.Show
**************************************************************************************************************************
PROCEDURE procExitFastRep
fRep.Release
SELECT fltBase
USE
SELECT tarfond
SET FILTER TO 
SELECT datJob
SET FILTER TO 
SELECT people
SET FILTER TO 
SET ORDER TO &oldInd
GO peopRec
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE tarifrefresh
frep.reptar=ALLTRIM(tarfond->plrep)
frep.newtar=0
frep.grdpers.Column5.ControlSource=frep.reptar
frep.Refresh
**************************************************************************************************************************
*                    Автозамена для ускоренной замены
**************************************************************************************************************************
PROCEDURE autorep
IF EMPTY(frep.reptar)
   DO createFormNew WITH .T.,'Автозамена',RetTxtWidth('WWНе указана тарифная величина!WW',dFontName,dFontSize+1),'130',;
   RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
   'Не указана тарифная величина!',.F.,.T.   
   RETURN
ENDIF
SELECT tarfond
LOCATE FOR plrep=frep.reptar
fauto=CREATEOBJECT('FORMSUPL')
WITH fauto
     .Caption='Автозамена'
     .backColor=RGB(255,255,255)
     DO adlabmy WITH 'fauto',1,'Данная процедура позволяет автоматически заменить',10,10,FONTMETRIC(6,dFontName,dFontSize)*TXTWIDTH('Данная процедура позволяет автоматически заменить',dFontName,dFontSize),2,.T.,1
     DO adlabmy WITH 'fauto',2,'выбранную тарифную величину на новое значение',.lab1.Top+.lab1.height,.lab1.Left,.lab1.width,2,.F.,1
     DO addShape WITH 'fauto',1,10,.lab2.Top+.lab2.Height+10,400,50,8 
     DO adlabmy WITH 'fauto',3,'Тарифная величина -'+ALLTRIM(tarfond.rec),.Shape1.Top+10,.Shape1.Left+1,.lab1.width-2,2,.F.,1
     DO adlabmy WITH 'fauto',4,'Новое значение -',.lab3.Top+.lab3.Height+5,.Shape1.Left+1,.lab1.width,0,.T.,1
     .lab4.Left=.lab4.Left+(.lab1.Width-2-.lab4.Width-100)/2
     DO addtxtboxmy WITH 'fauto',1,.lab4.Left+.lab4.Width+5,.lab4.Top,100,.F.,'frep.newtar'
     WITH .txtbox1           
          .Alignment=1
          .Enabled=.T.  
          .Format='Z'
          .InputMask='99999999.99'  
     ENDWITH
     .Shape1.Width=.lab1.Width
     .Shape1.Height=.txtBox1.Height+.lab3.Height+25
     DO addcontlabel WITH 'fauto','cont1',.Shape1.Left,.Shape1.Top+fauto.Shape1.Height+10,(.Shape1.Width-10)/2,dHeight+5,'Замена','DO autorepproc','Выполнить замену'
     DO addcontlabel WITH 'fauto','cont2',.cont1.Left+.cont1.Width+5,.Cont1.Top,.Cont1.Width,dHeight+5,'Отказ','fauto.Release','Отмена'
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+fauto.cont1.Height+40+fauto.lab1.Height*2    
ENDWITH
DO pasteImage WITH 'fauto'
fauto.Show
**************************************************************************************************************************
*              Непосредственно автозамена тарифа
**************************************************************************************************************************
PROCEDURE autorepproc
fauto.Release
SELECT datjob
new_ch=frep.reptar
REPLACE &new_ch WITH frep.newtar ALL
IF ALLTRIM(frep.repTar)='datjob.kf'  
   REPLACE namekf WITH IIF(SEEK(kf,'sprkoef',1),sprkoef.name,0) ALL 
ENDIF 
frep.Refresh
**************************************************************************************************************************
*
**************************************************************************************************************************
PROCEDURE menupriz
IF EMPTY(frep.reptar)
   DO createFormNew WITH .T.,'Автозамена',RetTxtWidth('WWНе указана тарифная величина!WW',dFontName,dFontSize+1),'130',;
   RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
   'Не указана тарифная величина!',.F.,.T.   
   RETURN
ENDIF
DIMENSION dimOpt(3)
STORE 0 TO dimOpt
dimOpt(1)=1
fSupl=CREATEOBJECT('FORMSUPL')
WITH fSupl
     DO addShape WITH 'fSupl',1,20,20,400,50,8 
     DO addOptionButton WITH 'fSupl',1,'тарифный разряд',.Shape1.Top+20,.Shape1.Left+20,'dimOpt(1)',0,"DO procValOption WITH 'fSupl','dimOpt',1",.T.   
     DO addOptionButton WITH 'fSupl',2,'квалификационная категория',.Option1.Top+.Option1.Height+20,.Option1.Left,'dimOpt(2)',0,"DO procValOption WITH 'fSupl','dimOpt',2",.T. 
     DO addOptionButton WITH 'fSupl',3,'категория персонала',.Option2.Top+.Option2.Height+20,.Option1.Left,'dimOpt(3)',0,"DO procValOption WITH 'fSupl','dimOpt',3",.T. 
     .Shape1.Height=.Option1.Height*3+80
     .Shape1.Width=.Option2.Width+40
     DO addcontlabel WITH 'fSupl','cont1',.Shape1.Left+(.Shape1.Width-RetTxtWidth('wприступитьw')*2-10)/2,.Shape1.Top+.Shape1.Height+20,RetTxtWidth('wприступитьw'),dHeight+5,'приступить','DO procRepPriz'
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+5,'отказ','fSupl.Release' 
     .Width=.Shape1.Width+40
     .Height=.Shape1.Height+.cont1.Height+60    
ENDWITH
DO pasteImage WITH 'fsupl'
fSupl.Show
************************************************************************************************************************
PROCEDURE procRepPriz
fSupl.Visible=.F.
fSupl.Release
CREATE CURSOR curforrep (namer C(40), kod N(3), newzn N(9,2))  
DO CASE
   CASE dimOpt(1)=1   &&тарифный разряд
        SELECT sprkoef
        SCAN ALL
             SELECT curForRep
             APPEND BLANK
             REPLACE namer WITH STR(sprkoef.kod,2)+' - '+STR(sprkoef.name,5,3),kod WITH sprkoef.kod
             SELECT sprkoef 
        ENDSCAN    
   CASE dimOpt(2)=1   &&квалификационная категория
        SELECT sprkval
        SET ORDER TO 1
        SCAN ALL
             SELECT curForRep
             APPEND BLANK
             REPLACE namer WITH sprkval.name,kod WITH sprkval.kod
             SELECT sprkval 
        ENDSCAN 
        SELECT curForRep
        APPEND BLANK
        REPLACE namer WITH 'без категории'
   CASE dimOpt(3)=1   &&категория персонала
        SELECT sprkat
        SET ORDER TO 1
        SCAN ALL
             SELECT curForRep
             APPEND BLANK
             REPLACE namer WITH sprkat.name,kod WITH sprkat.kod
             SELECT sprkat 
        ENDSCAN 
ENDCASE
fPriz=CREATEOBJECT('FORMSUPL')
WITH fPriz
     .Width=400
     .Height=400
     .Caption='Замена по признаку'
     .AddObject('Grdpriz','GridMy')
     WITH .grdpriz
          .Top=0
          .Left=0
          .Width=.Parent.Width
          .Height=.Parent.Height
          .ColumnCount=3
          .RecordSource='curforrep'
          .ScrollBars=2
          .Column2.Width=RetTxtWidth('9999999.999')
          .Column1.Width=.Width-.Column2.Width-SYSMETRIC(15)-13-.ColumnCount
          .Columns(.ColumnCount).Width=0
          .Column1.ControlSource='curforrep.namer' 
          .Column2.ControlSource='curforrep.newzn'
          .Column1.Header1.Caption='признак'
          .Column2.Header1.Caption='значение'
          .Column1.Enabled=.F.
          .Columns(.ColumnCount).Enabled=.T.
          .Column2.Format='Z'
          .Column1.Alignment=0
     ENDWITH          
     DO myColumnTxtBox WITH 'fPriz.grdpriz.column2','txtbox2','',.F.
     DO gridSize WITH 'fPriz','grdpriz','shapeingrid',.T.
     .grdpriz.column2.txtbox2.InputMask='9999999.999'
     DO addcontlabel WITH 'fPriz','cont1',(fPriz.Width-RetTxtWidth('WWзаменаWW')*2-10)/2,.grdpriz.Top+.grdpriz.Height+20,RetTxtWidth('WWзаменаWW',dFontname,dFontSize+1),dHeight+5,'замена','DO prizrep'
     DO addcontlabel WITH 'fPriz','cont2',.cont1.Left+.cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+5,'вовзрат','fPriz.Release'
     .Height=.grdPriz.Height+.cont1.Height+40
ENDWITH
SELECT curForRep
GO TOP
DO pasteImage WITH 'fPriz'
fPriz.Show
**************************************************************************************************************************
*                           Непосредственно замена по признаку
**************************************************************************************************************************
PROCEDURE prizrep
fForRep=frep.reptar
forPlRep=''
DO CASE
   CASE dimOpt(1)=1 
        forPlRep='kf'
   CASE dimOpt(2)=1 
        forPlRep='kv'
   CASE dimOpt(3)=1 
        forPlRep='kat'                
ENDCASE
SELECT curforrep
GO TOP
DO WHILE !EOF()
   SELECT datjob
   REPLACE &fForRep WITH curforrep.newzn FOR &forPlRep=curforrep.kod 
   IF ALLTRIM(frep.repTar)='datjob.kf'  
      REPLACE namekf WITH IIF(SEEK(kf,'sprkoef',1),sprkoef.name,0) FOR &forPlRep=curforrep.kod  
   ENDIF 
   SELECT curforrep
   SKIP 
ENDDO
SELECT datJob
GO TOP 
frep.Refresh
fPriz.Release
**************************************************************************************************************************
PROCEDURE poiskFastRep
PARAMETERS par_proc
fPoisk=CREATEOBJECT('FORMSUPL')
WITH fPoisk
     DO addShape WITH 'fPoisk',1,10,10,400,50,8     
     .logExit=.T.  
     find_ch=''
     DO adLabMy WITH 'fpoisk',1,'код или ФИО сотрудника' ,.Shape1.Top+10,.Shape1.Left+10,250,2
     DO addtxtboxmy WITH 'fpoisk',1,.Shape1.Left+10,.Shape1.Top+.lab1.Height+10,250,.F.,'find_ch'
     .txtBox1.procForkeyPress='DO keyPressPoiskFast'
     WITH .Shape1     
          .Width=fpoisk.TxtBox1.Width+20
          .Height=fpoisk.TxtBox1.Height+fpoisk.lab1.Height+42
     ENDWITH
     DO addcontlabel WITH 'fpoisk','cont1',.Shape1.Left+5,.Shape1.Top+.Shape1.Height+5,(.shape1.Width-20)/2,dHeight+3,'Поиск','DO procPoiskFastRep'
     DO addcontlabel WITH 'fpoisk','cont2',.Cont1.Left+.Cont1.Width+10,.Cont1.Top,.Cont1.Width,dHeight+3,'Отмена','Fpoisk.Release'    
     .Caption='Поиск'  
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+30+fpoisk.lab1.Height
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fpoisk'
fpoisk.Show
**************************************************************************************************************************
PROCEDURE procPoiskFastRep
IF EMPTY(find_ch)
   RETURN
ENDIF
find_ch=ALLTRIM(find_ch)
SELECT datjob
oldRec=RECNO()
log_ord=SYS(21)
IF TYPE(find_ch)='N'            
   IF SEEK(VAL(find_ch),'datjob',1)         
      SET ORDER TO 1
      fpoisk.Release 
      fRep.Cont2.SpecialEffect=0      
      fRep.Cont1.SpecialEffect=1      
      fRep.Grdpers.Column1.SetFocus                             
   ELSE  
      GO oldRec
      fPoisk.txtBox1.SetFocus        
   ENDIF             
ELSE
   DO unosimbol WITH 'find_ch'   
   IF SEEK(find_ch,'datjob',5)
      fRep.cont1.SpecialEffect=0
      fRep.cont2.SpecialEffect=1  
      SET ORDER TO 5       
      fRep.Grdpers.Column3.SetFocus   
      fpoisk.Release               
   ELSE        
      GO oldRec
      fPoisk.txtBox1.SetFocus       
   ENDIF
ENDIF
************************************************************************************************************************
PROCEDURE keyPressPoiskFast
DO CASE
   CASE LASTKEY()=27
        fpoisk.Release
   CASE LASTKEY()=13
        find_ch=fpoisk.TxtBox1.Value         
        DO procPoiskFastRep
ENDCASE 
**************************************************************************************************************************
PROCEDURE readFiltr
IF EMPTY(frep.reptar)   
   RETURN
ENDIF
WITH fRep
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .butPoisk.Visible=.T.
     .butReturn.Visible=.T.
     .grdPers.Columns(.grdPers.ColumnCount).Enabled=.F.
     .grdPers.Column5.Enabled=.T.
     .grdPers.Column5.SetFocus
ENDWITH
**************************************************************************************************************************
PROCEDURE returnFromReadFiltr
WITH fRep
     .SetAll('Visible',.T.,'myCommandButton')
     .SetAll('Visible',.T.,'mymenucont')
     .butPoisk.Visible=.F.
     .butReturn.Visible=.F.
     .grdPers.Columns(.grdPers.ColumnCount).Enabled=.T.
     .grdPers.Columns(.grdPers.ColumnCount).ReadOnly=.T.
     .grdPers.Column5.Enabled=.F.
     .grdPers.Columns(.grdPers.ColumnCount).SetFocus
ENDWITH
**************************************************************************************************************************
PROCEDURE poiskFiltr