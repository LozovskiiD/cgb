DIMENSION dim_tot(13),dim_day(13)
STORE 0 TO dim_tot,dim_day
formulaotp='datjob.mtokl+datjob.mstsum+datjob.mvto+datjob.mkat+datjob.mchir+datjob.mcharw+datjob.mmain+datjob.mmain2'
IF !USED('datprn')
   USE datprn IN 0
ENDIF
RESTORE FROM rashset ADDITIVE
SELECT datJob
SET ORDER TO 2
*RESTORE FROM setprn ADDITIVE 
countDate=varDtar
dim_day(10)=50
dim_day(6)=20
*USE datprn IN 0
SELECT sprpodr 
oldOrdPodr=SYS(21)
SET ORDER TO 2
SELECT rasp
SET RELATION TO kd INTO sprdolj ADDITIVE
GO TOP
fpodr=CREATEOBJECT('FORMSPR')
PUBLIC kpdop
kpdop=0
kpdop=IIF(rasp->kp=0,1,rasp->kp)
curnamepodr=''
SELECT cursprpodr
LOCATE FOR kod=kpdop
IF FOUND()
   curnamepodr=ALLTRIM(cursprpodr.name)  
ELSE
   GO TOP
   curnamepodr=cursprpodr.name 
ENDIF 
SELECT rasp
SET FILTER TO kp=kpdop
SELECT sprpodr
LOCATE FOR kod=kpdop    
var_path=FULLPATH('rashset.mem')
WITH fpodr
     .Caption='Расчет планируемых отпусков на оплату труда, для лиц замещающих уходящих в отпуск работников'   
     .AddProperty('kpdop')   
     DO addButtonOne WITH 'fPodr','menuCont1',10,5,'редакция','pencil.ico',"Do readspr WITH 'fpodr','Do inputzam'",39,RetTxtWidth('календарь')+44,'редакция'
     DO addButtonOne WITH 'fPodr','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'ред-е пер.','group.ico','DO inputPersZam',39,.menucont1.Width,'персонал'   
     DO addButtonOne WITH 'fPodr','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'удаление','pencild.ico','Do deletefromzam',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fPodr','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'расчёт','calculate.ico','DO procCountrash',39,.menucont1.Width,'расчёт'       
     DO addButtonOne WITH 'fPodr','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'печать','print1.ico','DO printRashod',39,.menucont1.Width,'печать' 
     DO addButtonOne WITH 'fPodr','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'настройки','setup.ico','DO setupZam',39,.menucont1.Width,'настройки'  
     DO addButtonOne WITH 'fPodr','menuCont7',.menucont6.Left+.menucont6.Width+3,5,'возврат','undo.ico','DO exitFromProcOtp',39,.menucont1.Width,'возврат'       
     DO addButtonOne WITH 'fPodr','menuexit1',10,5,'возврат','undo.ico','DO exitReadPers',39,RetTxtWidth('возврат')+44,'вовзрат' 
     .menuexit1.Visible=.F.   
     
     DO addmenureadspr WITH 'fpodr','DO writezam WITH .F.','DO writezam WITH .T.'   
**************************Combo для подразделения*******************************************************************************************************
     .AddObject('ComboBox1','Combomy')
     WITH .combobox1
          .BackColor=fpodr.BackColor
          .Width=500
          .SpecialEffect=1 
          .Height=dHeight
          .Top=fpodr.menucont1.Top+fpodr.menucont1.Height+5    
          .Left=(fpodr.Width-.Width)/2
          .ControlSource='curnamepodr'        
          .procForValid='DO validpodrasch'
          .toolTipText='выбор подразделения' 
          .RowSource='sprpodr->name'
          .RowSourceType=6    
     ENDWITH
********************************************************************************************************************************************************
     WITH .fGrid    
          .Top=fpodr.ComboBox1.Top+fpodr.ComboBox1.Height+5 
          .Height=dHeight*14
          .Width=fpodr.Width/3*2
          .RecordSource='rasp'
          DO addColumnToGrid WITH 'fPodr.fGrid',10
          .RecordSourceType=1     
          .Column1.ControlSource='rasp->nd'
          .Column2.ControlSource='" "+sprdolj.name'
          .Column3.ControlSource='rasp.kpeop'
          .Column4.ControlSource='rasp.lokl'
          .Column5.ControlSource='rasp.kse'
          .Column6.ControlSource='rasp.dotp'
          .Column7.ControlSource='rasp.dzam'
          .Column8.ControlSource='rasp.srzp'
          .Column9.ControlSource='rasp.zpday'     
          .Column1.Width=RettxtWidth(' 123 ')    
          .Column3.Width=RettxtWidth('99999')
          .Column4.Width=RettxtWidth('9999')
          .Column5.Width=RettxtWidth('99999')
          .Column6.Width=.Column4.Width
          .Column7.Width=.Column4.Width
          .Column8.Width=RettxtWidth('99999999.99')
          .Column9.Width=.Column8.Width       
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Должность'      
          .Column3.Header1.Caption='Сотр.'
          .Column4.Header1.Caption='То'
          .Column5.Header1.Caption='Ш.ед.'
          .Column6.Header1.Caption='Дн.отп.'
          .Column7.Header1.Caption='Дн.зам.'
          .Column8.Header1.Caption='Ср.зп.'
          .Column9.Header1.Caption='За 1 день.'     
          .Column3.Format='Z'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'     
          .Column9.Format='Z' 
          .Column1.Alignment=1
          .Column2.Alignment=0
          .Column3.Alignment=1         
          .Column5.Alignment=1         
          .Column6.Alignment=1         
          .Column7.Alignment=1         
          
          .Column4.AddObject('checkColumn4','checkContainer')
          .Column4.checkColumn4.AddObject('checkMy','checkBox')
          .Column4.CheckColumn4.checkMy.Visible=.T.
          .Column4.CheckColumn4.checkMy.Caption=''
          .Column4.CheckColumn4.checkMy.Left=10
          .Column4.CheckColumn4.checkMy.BackStyle=0
          .Column4.CheckColumn4.checkMy.ControlSource='rasp.lokl'                                                                                                  
          .column4.CurrentControl='checkColumn4'
          .Column4.Sparse=.F.              
          .procAfterRowColChange='DO fpodrRefresh'         
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .colNesInf=2              
     ENDWITH
     DO gridSizeNew WITH 'fpodr','fGrid','shapeingrid' 
     IF .fgrid.Top+dHeight*28>.Height
  
     ENDIF
     .combobox1.DisplayCount=MIN(RECCOUNT('sprpodr'),(fpodr.Height-fpodr.combobox1.Top-fpodr.combobox1.Height)/fpodr.fGrid.Rowheight)
         
     .AddObject('checkOkl','checkContainer')
     WITH .checkOkl
          .Width=.Parent.fGrid.Column4.Width+2    
          .AddObject('checkMy','MycheckBox')
          WITH .checkMy          
               .Caption=''
               .Left=10
               .BackStyle=0
               .ControlSource='rasp.lokl' 
               .Height=dHeight  
               .Visible=.T.                                                                                               
               .procForValid='DO procCheckOkl'
          ENDWITH     
          .Visible=.F.
          .BorderWidth=1
     ENDWITH 
     
     DO adtbox WITH 'fpodr',2,1,1,fpodr.fGrid.Column5.Width+2,dHeight,.F.,'Z',.T.
     DO adtbox WITH 'fpodr',3,1,1,fpodr.fGrid.Column6.Width+2,dHeight,.F.,'Z',.T.,2,'REPLACE dzam WITH ROUND(kse*dotp,0)'
     DO adtbox WITH 'fpodr',4,1,1,fpodr.fGrid.Column7.Width+2,dHeight,.F.,'Z',.T.
     DO adtbox WITH 'fpodr',5,1,1,fpodr.fGrid.Column8.Width+2,dHeight,.F.,'Z',.T.,2,'REPLACE zpday WITH srzp/rashset(4)'
     DO adtbox WITH 'fpodr',6,1,1,fpodr.fGrid.Column9.Width+2,dHeight,.F.,'Z',.T.
     .txtbox6.inputMask='999999.99'
     .SetAll('Visible',.F.,'MyTxtBox')

     objTop=.fGrid.Top
     objLeft=.fGrid.Width+5
     objWidth=ROUND((.Width-.fGrid.Width-5)/3,0)
     DO addcontform WITH 'fpodr','cont1',objLeft,objtop,objWidth,fpodr.fGrid.HeaderHeight+2,'Месяц' 
     DO addcontform WITH 'fpodr','cont2',fpodr.cont1.Left+fpodr.cont1.Width-1,objtop,objWidth,fpodr.cont1.Height,'Дни' 
     DO addcontform WITH 'fpodr','cont3',fpodr.cont2.Left+fpodr.cont2.Width-1,objtop,objWidth,fpodr.cont1.Height,'Сумма' 
     objTop=.cont1.Top+.cont1.Height-1
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'dim_month('+LTRIM(STR(i))+')','всего')  
         DO adtbox WITH 'fpodr',i+10,objLeft,objTop,objWidth,dHeight,vctrl,'Z',.F.,0
         objTop=objTop+dHeight-1
     ENDFOR 

     objTop=.txtbox11.Top
     objLeft=.Txtbox11.Left+.Txtbox11.Width-1
     objWidth=ROUND((.Width-.fGrid.Width-5)/3,0)
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'rasp->m'+LTRIM(STR(i)),'rasp->itday')     
         repzam=IIF(i=13,'itday','z'+LTRIM(STR(i)))   
         DO adtbox WITH 'fpodr',i+30,objLeft,objTop,objWidth,dHeight,vctrl,'Z',.F.,.F.
         obj_cx='fpodr.txtbox'+LTRIM(STR(i+30))
         &obj_cx..procForLostFocus='DO sumzptot'
         &obj_cx..procForValid='REPLACE &repzam WITH &vctrl*zpday'
         objTop=objTop+dHeight-1   
     ENDFOR 
     .txtbox43.Forecolor=IIF(rasp->itday#rasp->dzam,objcolorsos,dForeColor)
     objTop=.txtbox11.Top
     objLeft=.Txtbox31.Left+.Txtbox31.Width-1
     objWidth=ROUND((.Width-.fGrid.Width-5)/3,0) 
     FOR i=1 TO 13   
         vctrl=IIF(i<13,'rasp->z'+LTRIM(STR(i)),'rasp->totzp')
         DO adtbox WITH 'fpodr',i+50,objLeft,objTop,objWidth,dHeight,vctrl,'Z',
         objTop=objTop+dHeight-1   
     ENDFOR 
     DO addcontform WITH 'fpodr','cont35',0,fpodr.fgrid.Top+fpodr.fgrid.height-1,fpodr.Width,dHeight,''
     ftop=.cont35.Top+.cont35.Height-1
     grdperstop=ftop
     fpodr.AddObject('grdpers','GridMyNew')
     WITH .grdpers
          .ScrollBars=2
          .Left=0
          .Width=fpodr.Width/2
          .Top=ftop
          .Height=fpodr.Height-.Top
          *.ColumnCount=0
          .RecordSource='datJob'
           DO addColumnToGrid WITH 'fPodr.grdPers',4
          .RecordSourceType=1
          .Column1.ControlSource='datJob.kodPeop'
          .Column2.ControlSource='datJob.fio'
          .Column3.ControlSource='datJob.kse'
          .Column1.Header1.Caption='код'
          .Column2.Header1.Caption='фамилия'
          .Column3.Header1.Caption='объём'
          .Column1.Format='Z'
          .Column3.Format='Z'
          .Column1.Width=RetTxtWidth('99999')
          .Column3.Width=RetTxtWidth('9999999')
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-SYSMETRIC(5)-13-4
          .Column4.Width=0
          .Column1.Alignment=1              
          .Column2.Alignment=0
          .Column3.Alignment=1         
          .procAfterRowColChange='DO changepeople'
     ENDWITH
     DO gridSizeNew WITH 'fpodr','grdpers','shapeingrid1' 
     objWidth=ROUND((.Width-.grdpers.Width-5)/4,0)
     objTopOld=objTop
     DO addcontform WITH 'fpodr','cont36',fpodr.grdpers.Left+fpodr.grdpers.Width+5,ftop,objWidth,fpodr.grdpers.HeaderHeight,'Месяц' 
     DO addcontform WITH 'fpodr','cont37',fpodr.cont36.Left+fpodr.cont36.Width-1,fpodr.cont36.Top,fpodr.cont36.Width,fpodr.cont36.Height,'Дни отп.' 
     DO addcontform WITH 'fpodr','cont38',fpodr.cont37.Left+fpodr.cont37.Width-1,fpodr.cont36.Top,fpodr.cont36.Width,fpodr.cont36.Height,'Дни на ст.'
     DO addcontform WITH 'fpodr','cont39',fpodr.cont38.Left+fpodr.cont38.Width-1,fpodr.cont36.Top,fpodr.cont36.Width,fpodr.cont36.Height,'Сумма'
     .cont39.Width=.Width-.cont39.Left
     objtop=ftop+dHeight-1
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'dim_month('+LTRIM(STR(i))+')','всего')  
         DO adtbox WITH 'fpodr',i+120,fpodr.cont36.Left,objtop,fpodr.cont36.Width,dHeight,vctrl,'Z',.F.,0
         objtop=objtop+dHeight-1
     ENDFOR 

     objtop=ftop+dHeight-1
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'datJob.d'+LTRIM(STR(i)),'datJob.dtot')  
         DO adtbox WITH 'fpodr',i+140,fpodr.cont37.Left,objtop,fpodr.cont37.Width,dHeight,vctrl,'Z',.F.,2
         repzam=IIF(i=13,'datJob.dsttot','datJob.dst'+LTRIM(STR(i))) 
         repzp=IIF(i=13,'datJob.zptot','datJob.zp'+LTRIM(STR(i)))  
         obj_cx='fpodr.txtbox'+LTRIM(STR(i+140))
         &obj_cx..procForLostFocus='DO sumzpdoltot'
         &obj_cx..procForValid='REPLACE &repzam WITH &vctrl*datJob.kse,&repzp WITH &repzam*rasp->zpday'
         objtop=objtop+dHeight-1
     ENDFOR 

     objtop=ftop+dHeight-1
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'datJob.dst'+LTRIM(STR(i)),'datJob.dsttot')  
         DO adtbox WITH 'fpodr',i+160,fpodr.cont38.Left,objtop,fpodr.cont38.Width,dHeight,vctrl,'Z',.F.,2
         repzam=IIF(i=13,'datJob.dsttot','datJob.dst'+LTRIM(STR(i))) 
         repzp=IIF(i=13,'datJob.zptot','datJob.zp'+LTRIM(STR(i)))  
         obj_cx='fpodr.txtbox'+LTRIM(STR(i+160))
         &obj_cx..procForLostFocus='DO sumzpdoltot'
         &obj_cx..procForValid='REPLACE &repzp WITH &repzam*rasp->zpday'
         objtop=objtop+dHeight-1  
     ENDFOR 
     objtop=ftop+dHeight-1
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'datJob.zp'+LTRIM(STR(i)),'datJob.zptot')  
         DO adtbox WITH 'fpodr',i+180,fpodr.cont39.Left,objtop,fpodr.cont39.Width,dHeight,vctrl,'Z',.F.,2
         objtop=objtop+dHeight-1  
         obj_cx='fpodr.txtbox'+LTRIM(STR(i+180))  
         &obj_cx..procForLostFocus='DO sumzpdoltot'
     ENDFOR 
     ord_ch1=1
     DO WHILE .T.
        IF objtop+dHeight>SYSMETRIC(2)     
           EXIT
        ENDIF
        obj_cont='lCont'+LTRIM(STR(ord_ch1))
        .AddObject(obj_cont,'Shape')
        WITH .&obj_cont
             .Height=dHeight
             .Width=fpodr.cont36.Width
             .Top=objTop
             .Left=fpodr.cont36.Left
             .BackStyle=0                   
             .Visible=.T.
        ENDWITH    
        ord_ch1=ord_ch1+1
        obj_cont='lCont'+LTRIM(STR(ord_ch1))
        .AddObject(obj_cont,'Shape')
        WITH .&obj_cont
             .Height=dHeight
             .Width=fpodr.Cont37.Width
             .Top=objTop
             .Left=fpodr.Cont37.Left
             .BackStyle=0                   
             .Visible=.T.
        ENDWITH       
        ord_ch1=ord_ch1+1
        obj_cont='lCont'+LTRIM(STR(ord_ch1))
        .AddObject(obj_cont,'Shape')
        WITH .&obj_cont
             .Height=dHeight
             .Width=fpodr.Cont38.Width
             .Top=objTop
             .Left=fpodr.Cont38.Left
             .BackStyle=0                   
             .Visible=.T.
        ENDWITH       
        ord_ch1=ord_ch1+1
        obj_cont='lCont'+LTRIM(STR(ord_ch1))
        .AddObject(obj_cont,'Shape')
        WITH .&obj_cont
             .Height=dHeight
             .Width=fpodr.Cont39.Width
             .Top=objTop
             .Left=fpodr.Cont39.Left
             .BackStyle=0                   
             .Visible=.T.
         ENDWITH       
         ord_ch1=ord_ch1+1
         objtop=objtop+dHeight-1         
     ENDDO 
ENDWITH 
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus    
fpodr.Show
********************************************************************************************************************************************************
PROCEDURE exitFromProcOtp
IF USED('datprn')
   SELECT datprn 
   USE 
ENDIF
SELECT sprpodr 
SET ORDER TO &oldOrdPodr
SELECT rasp
SET RELATION TO
SELECT datJob
SET ORDER TO 4
fPodr.Release
********************************************************************************************************************************************************
PROCEDURE fpodrRefresh
DO incursor
**************************************************************************************************************************
*                                      Выбор услуг по документу
**************************************************************************************************************************
PROCEDURE incursor
SELECT datJob
SET FILTER TO kd=rasp.kd.AND.kp=rasp.kp
GO TOP
WITH fpodr.grdpers 
     .Top=grdperstop
     .RecordSource='datJob'
     .RecordSourceType=1
     .Column1.ControlSource='datJob.kodPeop'
     .Column2.ControlSource='datJob.fio'
     .Column3.ControlSource='datJob.kse'
     .Column1.Width=RetTxtWidth('99999')
     .Column3.Width=RetTxtWidth('9999999')
     .Column2.Width=.Width-.Column1.Width-.Column3.Width-SYSMETRIC(5)-13-4
     .Column4.Width=0     
ENDWITH
fpodr.txtbox43.DisabledForecolor=IIF(rasp->itday#rasp->dzam,objcolorsos,dForeColor)
SELECT rasp
fpodr.Refresh()
******************************************************************************************************************************
*             Информация по сотруднику 
*****************************************************************************************************************************
PROCEDURE changepeople
SELECT rasp
fpodr.Refresh
*---------------------------------------------Процедры для расчёта, удаления и печати-------------------------------------------------------------------
********************************************************************************************************************************************************
*
********************************************************************************************************************************************************
PROCEDURE validpodrasch
SELECT sprpodr
kpdop=sprpodr.kod
curnamepodr=fpodr.ComboBox1.Value
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
fpodr.Refresh

*********************************************************************************************************************************************************
*                                     Редактирование информации по замене (должность)
*********************************************************************************************************************************************************
PROCEDURE inputzam
SELECT rasp
SCATTER TO fpodr.dim_ap  
SELECT datJob
SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
STORE 0 TO nkse,nms,srms,npeop
DO WHILE kp=rasp.kp.AND.kd=rasp.kd
   npeop=npeop+1
   nkse=nkse+datJob.kse  
   IF !rasp.lOkl
      nms=nms+&formulaOtp
   ELSE 
      nms=nms+datJob.mtokl
   ENDIF    
   SKIP 
   ENDDO
SELECT rasp 
REPLACE srzp WITH IIF(nkse<1,nms,nms/nkse),zpday WITH srzp/rashset(4),kpeop WITH npeop   
WITH fPodr
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .fGrid.Column1.SetFocus
     .menuread.Visible=.T.
     .menuexit.Visible=.T.
     .fGrid.Enabled=.F.
     .grdpers.Enabled=.F.
     .combobox1.Enabled=.F.
     .combobox1.Style=0
     .CheckOkl.Visible=.T.
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .txtbox2.Top=linetop
     .CheckOkl.Top=lineTop
     .txtbox3.Top=linetop
     .txtbox4.Top=linetop
     .txtbox5.Top=linetop
     .txtbox6.Top=linetop
     .txtbox2.Height=.fGrid.RowHeight+1
     .CheckOkl.Height=.fGrid.RowHeight+1
     .txtbox3.Height=.fGrid.RowHeight+1
     .txtbox4.Height=.fGrid.RowHeight+1
     .txtbox5.Height=.fGrid.RowHeight+1
     .txtbox6.Height=.fGrid.RowHeight+1
     .checkOkl.Left=.fGrid.Left+13+.fGrid.Column1.Width+.fgrid.Column2.Width+.fgrid.Column3.Width
     .txtBox2.Left=.checkOkl.Left+.checkOkl.Width-1
     
     .txtbox3.Left=.txtBox2.Left+.txtBox2.Width-1
     .txtbox4.Left=.txtbox3.Left+.txtbox3.Width-1
     .txtbox5.Left=.txtbox4.Left+.txtbox4.Width-1
     .txtbox6.Left=.txtbox5.Left+.txtbox5.Width-1
     .txtbox2.ControlSource='rasp->kse'
     .txtbox3.ControlSource='rasp->dotp'
     .txtbox4.ControlSource='rasp->dzam'
     .txtbox5.ControlSource='rasp->srzp'
     .txtbox6.ControlSource='rasp->zpday'
     .txtbox2.BackStyle=1
     .checkOkl.BackStyle=1
     .txtbox3.BackStyle=1
     .txtbox4.BackStyle=1
     .txtbox5.BackStyle=1
     .txtbox6.BackStyle=1
     .SetAll('Visible',.T.,'MyTxtBox')
     .Refresh
     .Txtbox2.SetFocus
ENDWITH 
SELECT rasp
*************************************************************************************************************************
PROCEDURE procCheckOkl
SELECT datJob
SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
STORE 0 TO nkse,nms,srms,npeop
DO WHILE kp=rasp.kp.AND.kd=rasp.kd
   npeop=npeop+1
   nkse=nkse+datJob.kse  
   IF !rasp.lOkl
      nms=nms+&formulaOtp
   ELSE 
      nms=nms+datJob.mtokl
   ENDIF    
   SKIP 
   ENDDO
SELECT rasp 

REPLACE srzp WITH IIF(nkse<1,nms,nms/nkse),zpday WITH srzp/rashset(4),kpeop WITH npeop   
KEYBOARD '{TAB}'
fPodr.Refresh
*************************************************************************************************************************
*                               Запись информации по замене
*************************************************************************************************************************
PROCEDURE writezam
PARAMETERS parlog
WITH fPodr
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     .menuexit1.Visible=.F.
     SELECT rasp
     IF parlog
        GATHER FROM fpodr.dim_ap
     ENDIF
     .txtbox2.Visible=.F.
     .checkOkl.Visible=.F.
     .txtbox3.Visible=.F.
     .txtbox4.Visible=.F.
     .txtbox5.Visible=.F.
     .txtbox6.Visible=.F.
     .combobox1.Enabled=.T.
     .combobox1.Style=2
     FOR i=31 TO 43
         objtxt='txtbox'+LTRIM(STR(i))
         fpodr.&objtxt..Enabled=.F. 
         fpodr.&objtxt..BackStyle=0         
     ENDFOR 
     .Grdpers.Enabled=.T.  
     .Grdpers.SetAll('Enabled',.F.,'ColumnMy')
     .Grdpers.Columns(.grdpers.ColumnCount).Enabled=.T.
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH 
IF !parlog.AND.rasp->dzam#rasp->itday
   * DO createForm WITH .T.,'Внимание!',RetTxtWidth('WWНесовпадение дней замены!WW',dFontName,dFontSize+1),'130',;
   *  RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,'Несовпадение дней замены!' 
ENDIF
******************************************************************************************************************************************************
*                             Редактирование информации по замене (персонал)
******************************************************************************************************************************************************
PROCEDURE inputperszam 
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus   
IF rasp->dzam=0.OR.rasp->zpday=0
   DO createForm WITH .T.,'Редактирование',RetTxtWidth('WНе указаны дни замены!W',dFontName,dFontSize+1),;
   '130',RetTxtWidth('WWOKWW',dFontName,dFontSize+1),'ОК',.F.,.F.,'nFormMes.Release',.F.,.F.,'Не указаны дни замены!'   
   RETURN
ENDIF  
WITH fPodr  
     .SetAll('Visible',.F.,'mymenucont')
     .SetAll('Visible',.F.,'myCommandButton')
     .menuexit1.Visible=.T.
     .combobox1.Enabled=.F.
     .combobox1.Style=0
     .fGrid.Enabled=.F.
     FOR i=141 TO 152
         objtxt='txtbox'+LTRIM(STR(i))
         fpodr.&objtxt..Enabled=.T. 
         fpodr.&objtxt..BackStyle=1   
     ENDFOR 
     FOR i=161 TO 172
         objtxt='txtbox'+LTRIM(STR(i))
         fpodr.&objtxt..Enabled=.T. 
         fpodr.&objtxt..BackStyle=1   
     ENDFOR 
     IF rashset(2)
        FOR i=181 TO 192
            objtxt='txtbox'+LTRIM(STR(i))
            fpodr.&objtxt..Enabled=.T. 
            fpodr.&objtxt..BackStyle=1   
         ENDFOR 
     ENDIF
ENDWITH 
******************************************************************************************************************************************************
*                                Выход из редактирования по персоналу
******************************************************************************************************************************************************
PROCEDURE exitreadpers
WITH fPodr
     .SetAll('Visible',.T.,'mymenucont')
     .SetAll('Visible',.T.,'myCommandButton')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     .menuexit1.Visible=.F.
     .combobox1.Enabled=.T.
     .combobox1.Style=2
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'ColumnMy')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
     FOR i=141 TO 152
         objtxt='txtbox'+LTRIM(STR(i))
         fpodr.&objtxt..Enabled=.F. 
         fpodr.&objtxt..BackStyle=0
     ENDFOR 
     FOR i=161 TO 172
         objtxt='txtbox'+LTRIM(STR(i))
         fpodr.&objtxt..Enabled=.F. 
         fpodr.&objtxt..BackStyle=0   
     ENDFOR 
     IF rashset(2)
        FOR i=181 TO 192
            objtxt='txtbox'+LTRIM(STR(i))
            fpodr.&objtxt..Enabled=.F. 
            fpodr.&objtxt..BackStyle=0   
        ENDFOR 
     ENDIF
     fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus  
ENDWITH 
*****************************************************************************************************************************************************
PROCEDURE sumzptot
SELECT rasp
STORE 0 TO tot_cx,day_cx
FOR h=1 TO 12
    tot_cx=tot_cx+EVALUATE('z'+LTRIM(STR(h)))
    day_cx=day_cx+EVALUATE('m'+LTRIM(STR(h)))
ENDFOR 
REPLACE totzp WITH tot_cx, itday WITH day_cx
fpodr.Refresh
*****************************************************************************************************************************************************
PROCEDURE sumzpdoltot
SELECT datJob
STORE 0 TO tot_cx,day_cx,dayst_cx
FOR h=1 TO 12
    tot_cx=tot_cx+EVALUATE('zp'+LTRIM(STR(h)))
    day_cx=day_cx+EVALUATE('d'+LTRIM(STR(h)))
    dayst_cx=dayst_cx+EVALUATE('dst'+LTRIM(STR(h)))
    
ENDFOR 
REPLACE zptot WITH tot_cx, dtot WITH day_cx,dsttot WITH dayst_cx
FOR i=1 TO 13  
    obj_cx='fpodr.txtbox'+LTRIM(STR(i+140))   
    &obj_cx..Refresh
    obj_ch='fpodr.txtbox'+LTRIM(STR(i+160))   
    &obj_ch..Refresh
    obj_cf='fpodr.txtbox'+LTRIM(STR(i+180))   
    &obj_cf..Refresh
ENDFOR 
sumrec=RECNO()
DIMENSION dim_dayst(13),dim_zp(13)
STORE 0 TO dim_dayst,dim_zp
FOR i=1 TO 13
    sumdim=IIF(i<13,'dst'+LTRIM(STR(i)),'dsttot')
    repdim='dim_dayst'+LTRIM(STR(i))
    sumzp=IIF(i<13,'zp'+LTRIM(STR(i)),'zptot')
    repzp='dim_zp'+LTRIM(STR(i))
    SUM &sumdim,&sumzp TO &repdim,&repzp
    SELECT rasp
    repm=IIF(i<13,'m'+LTRIM(STR(i)),'itday')
    repz=IIF(i<13,'z'+LTRIM(STR(i)),'totzp')
    IF i<13
       REPLACE &repm WITH &repdim,&repz WITH &repzp
       IF !rashset(9)
          REPLACE &repz WITH EVALUATE('m'+LTRIM(STR(i)))*zpday  &&общая сумма по должности без учета округлений          
       ENDIF   
    ENDIF                   
    obj_cx='fpodr.txtbox'+LTRIM(STR(i+30))
    &obj_cx..Refresh
    obj_ch='fpodr.txtbox'+LTRIM(STR(i+50))
    IF i=13
       &obj_cx..DisabledForecolor=IIF(rasp->itday#rasp->dzam,objcolorsos,dForeColor)
    ENDIF   
    &obj_ch..Refresh
    SELECT datJob
ENDFOR
SELECT rasp
REPLACE totzp WITH z1+z2+z3+z4+z5+z6+z7+z8+z9+z10+z11+z12
REPLACE itday WITH m1+m2+m3+m4+m5+m6+m7+m8+m9+m10+m11+m12
SELECT datJob
GO sumrec
*****************************************************************************************************************************************************
*                         Непосредственно удаление информации по замене
*****************************************************************************************************************************************************
PROCEDURE delreczam
IF !log_del 
   RETURN 
ENDIF
SELECT rasp
oldkp=kp
oldkd=kd
oldnp=np
oldnd=nd
oldkse=kse
oldkat=kat
SELECT rasp
DO CASE
   CASE dim_del(1)=1
        SELECT datJob
        FOR i=1 TO 12
            repd='d'+LTRIM(STR(i))
            repdst='dst'+LTRIM(STR(i))
            repzp='zp'+LTRIM(STR(i))
            REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
        ENDFOR                      
        replace zptot WITH 0,dtot WITH 0,dsttot WITH 0
        SELECT rasp
        FOR i=1 TO 12
            repm='m'+LTRIM(STR(i))
            repz='z'+LTRIM(STR(i))
            REPLACE &repm WITH 0,&repz WITH 0
        ENDFOR                      
        replace kpeop WITH 0,dotp WITH 0,dzam WITH 0,itday WITH 0,srzp WITH 0,zpday WITH 0,totzp WITH 0
   CASE dim_del(2)=1
        SELECT datJob
        SET FILTER TO kp=kpdop
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12
               repd='d'+LTRIM(STR(i))
               repdst='dst'+LTRIM(STR(i))
               repzp='zp'+LTRIM(STR(i))
               REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
           ENDFOR                      
           replace zptot WITH 0,dtot WITH 0,dsttot WITH 0
           SKIP
        ENDDO         
        SELECT rasp        
        GO TOP
        DO WHILE !EOF()
           SELECT rasp
           FOR i=1 TO 12
               repm='m'+LTRIM(STR(i))
               repz='z'+LTRIM(STR(i))
               REPLACE &repm WITH 0,&repz WITH 0
           ENDFOR                      
           replace kpeop WITH 0,dotp WITH 0,dzam WITH 0,itday WITH 0,srzp WITH 0,zpday WITH 0,totzp WITH 0
           SKIP
        ENDDO   
        SET FILTER TO kp=kpdop
        GO top         
   CASE dim_del(3)=1
        SELECT datJob 
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12
               repd='d'+LTRIM(STR(i))
               repdst='dst'+LTRIM(STR(i))
               repzp='zp'+LTRIM(STR(i))
               REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
           ENDFOR                      
           replace zptot WITH 0,dtot WITH 0,dsttot WITH 0
           SKIP
        ENDDO           
        SELECT rasp
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()
           SELECT rasp
           FOR i=1 TO 12
               repm='m'+LTRIM(STR(i))
               repz='z'+LTRIM(STR(i))
               REPLACE &repm WITH 0,&repz WITH 0
           ENDFOR                      
           replace kpeop WITH 0,dotp WITH 0,dzam WITH 0,itday WITH 0,srzp WITH 0,zpday WITH 0,totzp WITH 0
           SKIP
        ENDDO   
        SET FILTER TO kp=kpdop
        GO top                  
ENDCASE
fdel.Release
SELECT rasp
fpodr.Refresh
*****************************************************************************************************************************************************
*                  Форма для удаления сведений по замене
*****************************************************************************************************************************************************
PROCEDURE deletefromzam
fdel=CREATEOBJECT('FORMMY')
log_del=.F.
DIMENSION dim_del(4)
STORE 0 TO dim_del
dim_del(1)=1
WITH fdel
     .Caption='Удаление'
     .BackColor=RGB(255,255,255)
     .AddObject('Shape1','ShapeMy')
     .Shape1.Top=10
     .Shape1.Left=10
     .Shape1.Curvature=8
     .Shape1.BorderColor=RGB(192,192,192)     
ENDWITH
DO addOptionButton WITH 'fdel',1,'очистить выбранную запись',fdel.Shape1.Top+10,fdel.Shape1.Left+15,'dim_del(1)',0,"DO storedimdel WITH 1",.T.
DO addOptionButton WITH 'fdel',2,'удалить подразделение',fdel.Option1.Top+fdel.Option1.Height+10,fdel.Option1.Left,'dim_del(2)',0,"DO storedimdel WITH 2",.T.
DO addOptionButton WITH 'fdel',3,'удалить все',fdel.Option2.Top+fdel.Option2.Height+10,fdel.Option1.Left,'dim_del(3)',0,"DO storedimdel WITH 3",.T.
fdel.Shape1.Height=fdel.Option1.height*3+40
fdel.Shape1.Width=fdel.Option1.Width+30
DO adCheckBox WITH 'fdel','check1','подтверждение удаления',fdel.Shape1.Top+fdel.Shape1.Height+10,fdel.Shape1.Left,150,dHeight,'log_del',0
DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+5,fdel.check1.Top+fdel.check1.Height+15,;
   (fdel.shape1.Width-20)/2,dHeight+3,'Выполнение','DO delreczam'
DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+10,fdel.Cont1.Top,;
   fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'

WITH fdel        
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+fpodr.cont1.Height+fdel.check1.Height+50
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
***********************************************************************************************************************************************
PROCEDURE storedimdel
PARAMETERS par1
FOR i=1 TO 4
    dim_del(i)=IIF(i=par1,1,0)
ENDFOR
fdel.Refresh
*****************************************************************************************************************************************************
*                         Форма для общего расчёта сведений по замене
*****************************************************************************************************************************************************
PROCEDURE proccountrash
fdel=CREATEOBJECT('FORMMY')
log_del=.F.
DIMENSION dim_del(4)
STORE 0 TO dim_del
dim_del(1)=1
dim_del(2)=0
dim_del(3)=0
log_srzp=.F.
WITH fdel
     .Caption='Расчёт расходов'
     .BackColor=RGB(255,255,255)   
     DO addShape WITH 'fdel',1,10,10,dHeight,50,8     
     DO adLabMy WITH 'fdel',1,'Дата отсчёта',fdel.Shape1.Top+10,fdel.Shape1.Left+15,150,0,.T.
     DO adtbox WITH 'fdel',1,fdel.lab1.Left+fdel.lab1.Width+10,fdel.Shape1.Top+10,RetTxtWidth('99/99/99999'),dHeight,'varDtar','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashset'
     fdel.lab1.Top=fdel.txtbox1.Top+(fdel.txtbox1.Height-fdel.lab1.Height)
     DO addOptionButton WITH 'fdel',1,'расчет по выбранной должности',fdel.txtbox1.Top+fdel.txtbox1.Height+10,fdel.Shape1.Left+15,'dim_del(1)',0,"DO storedimdel WITH 1",.T.
     DO addOptionButton WITH 'fdel',2,'расчёт по подразделению',fdel.Option1.Top+fdel.Option1.Height+10,fdel.Option1.Left,'dim_del(2)',0,"DO storedimdel WITH 2",.T.
     DO addOptionButton WITH 'fdel',3,'расчёт по организации',fdel.Option2.Top+fdel.Option2.Height+10,fdel.Option1.Left,'dim_del(3)',0,"DO storedimdel WITH 3",.T.
     .Shape1.Height=.Option1.height*4+60    
     
     DO addShape WITH 'fdel',4,10,fdel.Shape1.Top+fdel.Shape1.Height+10,dHeight,fdel.Shape1.Width,8 
     DO adCheckBox WITH 'fdel','check1','пересчитать среднюю зарплату',fdel.Shape4.Top+10,fdel.Option1.Left,150,dHeight,'log_srzp',0
     DO adCheckBox WITH 'fdel','checkRound','учитывать округление в общей сумме по должности',fdel.check1.Top+fdel.check1.Height+10,fdel.check1.Left,150,dHeight,'rashset(9)',0,.F.,'SAVE TO &var_path ALL LIKE rashset'
     .Shape4.Height=.check1.Height*2+30
     .Shape4.Width=.checkRound.Width+30
     .Shape1.Width=.Shape4.Width   
     DO adCheckBox WITH 'fdel','check2','подтверждение выполнения',fdel.Shape4.Top+fdel.Shape4.Height+10,fdel.Shape1.Left,150,dHeight,'log_del',0    
     .check2.Left=.shape4.Left+(.shape4.Width-.check2.Width)/2
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+(.Shape1.Width-RetTxtWidth('WВыполнениеW')*2-20)/2,fdel.check2.Top+fdel.check2.Height+15,;
       RetTxtWidth('WВыполнениеW'),dHeight+3,'Выполнение','DO countrash'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+20,fdel.Cont1.Top,;
        fdel.Cont1.Width,dHeight+3,'Отмена','fdel.Release'         
        
      DO adLabMy WITH 'fdel',4,'Ход выполнения',fdel.check2.Top+fdel.check2.Height+5,fdel.Shape1.Left,fdel.Shape1.Width,2,.F.
     .lab4.Visible=.F.        
     DO addShape WITH 'fdel',2,fdel.Shape1.Left,fdel.lab4.Top+fdel.lab4.Height+5,dHeight,fdel.Shape1.Width
     .Shape2.BackStyle=0
     .Shape2.Visible=.F.
          
     DO addShape WITH 'fdel',3,fdel.Shape2.Left,fdel.Shape2.Top,dHeight,0
     .Shape3.BackStyle=1
     .Shape3.Visible=.F.             
        
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.Shape4.Height+fdel.cont1.Height+fdel.check1.Height+80
     .lab1.left=.Shape1.Left+(.shape1.Width-.lab1.Width-.txtbox1.Width-10)/2
     .txtbox1.Left=.lab1.Left+.lab1.Width+10
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fdel'
fdel.Show
*****************************************************************************************************************************************************
*                                             Непосредственно  общий расчёт расходов по замене
*****************************************************************************************************************************************************
PROCEDURE countrash
IF !log_del
   fDel.Release
   RETURN
ENDIF
SELECT rasp
recpeop=RECNO()
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec 
GO recpeop
fdel.cont1.Visible=.F.
fdel.cont2.Visible=.F.
fdel.lab4.Visible=.T. 
fdel.Shape2.Visible=.T.
fdel.Shape3.Visible=.T.
=SYS(2002)
DO CASE
   CASE dim_del(1)=1
        max_rec=1 
        DO countone
        one_pers=one_pers+1
        pers_ch=one_pers/max_rec*100
        fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch 
   CASE dim_del(2)=1
        SELECT rasp
        GO TOP        
        DO WHILE !EOF()
           DO countone
           SELECT rasp
           SKIP 
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch               
        ENDDO
   CASE dim_del(3)=1
        SELECT rasp
        SET FILTER TO
        COUNT TO max_rec 
        GO TOP        
        DO WHILE !EOF()
           DO countone
           SELECT rasp
           SKIP
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fdel.Shape3.Width=fdel.shape2.Width/100*pers_ch  
        ENDDO
ENDCASE
SELECT rasp
SET FILTER TO kp=kpdop
GO TOP
=INKEY(1)
fDel.Visible=.F.
fdel.Release
DO createFormNew WITH .T.,'Общий расчёт',RetTxtWidth('WWРасчёт выполнен!WW',dFontName,dFontSize+1),'130',;
      RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
      'Расчёт выполнен!',.F.,.T. 
=SYS(2002,1) 
fpodr.Refresh 
********************************************************************************************************************************************************
*                     Процедура расчёта расходов на замену по одной должности
********************************************************************************************************************************************************
PROCEDURE countone
SELECT rasp
DIMENSION dim_day(13),sumzp_dol(12)
STORE 0 TO dim_day,sumzp_dol
IF itday#0
   SELECT datJob
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   STORE 0 TO nkse,nms,srms,npeop
   DO WHILE kp=rasp.kp.AND.kd=rasp.kd
      npeop=npeop+1
      nkse=nkse+datJob.kse  
      IF !rasp.lOkl
         nms=nms+&formulaOtp
      ELSE 
         nms=nms+datJob.mtokl
      ENDIF        
      SKIP 
   ENDDO
   SELECT rasp 
  * REPLACE kpeop WITH npeop,srzp WITH IIF(nkse<1,nms,nms/nkse),zpday WITH srzp/rashset(4)  
   IF log_srzp
      REPLACE kpeop WITH  npeop,srzp WITH IIF(nkse<1,nms,nms/kse),zpday WITH srzp/rashset(4)
   ENDIF      
   SELECT datJob
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   DO WHILE kp=rasp->kp.AND.kd=rasp->kd
      STORE 0 TO tot_cx,day_cx,dayst_cx
      FOR h=1 TO 12
          IF h>=MONTH(countDate)             
             tot_cx=tot_cx+EVALUATE('zp'+LTRIM(STR(h)))
             day_cx=day_cx+EVALUATE('d'+LTRIM(STR(h)))
             dayst_cx=dayst_cx+EVALUATE('dst'+LTRIM(STR(h)))    
             repdst='dst'+LTRIM(STR(h))
             repzp='zp'+LTRIM(STR(h))
             REPLACE &repdst WITH EVALUATE('dst'+LTRIM(STR(h))),&repzp WITH &repdst*rasp->zpday 
             sumzp_dol(h)=sumzp_dol(h)+&repzp  && для общей суммы по должности помесячно с учетом округлений округлений            
             dim_day(h)=dim_day(h)+EVALUATE('dst'+LTRIM(STR(h)))   &&
             dim_day(13)=dim_day(13)+EVALUATE('dst'+LTRIM(STR(h))) && 
          ELSE
             repzp='zp'+LTRIM(STR(h))
             repd='d'+LTRIM(STR(h))
             repdst='dst'+LTRIM(STR(h))
             REPLACE &repzp WITH 0,&repd WITH 0,&repdst WITH 0
          ENDIF   
      ENDFOR 
      zptot_cx=zp1+zp2+zp3+zp4+zp5+zp6+zp7+zp8+zp9+zp10+zp11+zp12
      REPLACE zptot WITH zptot_cx, dtot WITH day_cx,dsttot WITH dayst_cx
      SKIP
   ENDDO  
   SELECT rasp
   STORE 0 TO tot_cx,day_cx
   FOR h=1 TO 12
       rep_cx='z'+LTRIM(STR(h))
       mon_cx='m'+LTRIM(STR(h))
       IF h>=MONTH(countDate) 
          REPLACE  &mon_cx WITH dim_day(h)  &&
          IF !rashset(9)
             REPLACE &rep_cx WITH EVALUATE('m'+LTRIM(STR(h)))*zpday  &&общая сумма по должности без учета округлений
          ELSE   
             REPLACE &rep_cx WITH sumzp_dol(h)                        &&общая сумма по должности с учетом округлений    
          ENDIF   
          tot_cx=tot_cx+EVALUATE('z'+LTRIM(STR(h)))
          day_cx=day_cx+EVALUATE('m'+LTRIM(STR(h)))
       ELSE
          REPLACE &rep_cx WITH 0,&mon_cx WITH 0    
       ENDIF   
   ENDFOR 
*   REPLACE totzp WITH tot_cx, itday WITH day_cx,dzam WITH day_cx 
   REPLACE totzp WITH tot_cx, itday WITH dim_day(13),dzam WITH day_cx    
   &&
   IF itday=0
      REPLACE srzp WITH 0, zpday WITH 0, dzam WITH 0
   ENDIF  
ENDIF

*****************************************************************************************************************************************************
PROCEDURE printrashod
fSupl=CREATEOBJECT('FORMSUPL')

SELECT datprn
GO TOP
logWord=.F.
kvo_page=1
page_beg=1
page_end=999
term_ch=.T.
procForPrn=datprn.procved
strPrn=ALLTRIM(datprn.nameved)
WITH fSupl
     .Caption='Ведомости'
     .procexit='DO exitPrintRepZam'
     DO addshape WITH 'fSupl',1,20,20,150,380,8 
     DO addComboMy WITH 'fSupl',11,.Shape1.Left+20,.Shape1.Top+20,dHeight,.Shape1.Width-40,.T.,'strprn','datprn.nameved',6,.F.,'DO selectVedZam',.F.,.T.    
     .Shape1.Height=.comboBox11.Height+40
     DO adSetupPrnToForm WITH .Shape1.Left,.Shape1.Top+.Shape1.Height+20,.Shape1.Width,.F.,.T.
   
     *-----------------------------Кнопка печать---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.shape1.Left+(.Shape1.Width-(RetTxtWidth('wпросмотрw')*3)-30)/2,;
       .Shape91.Top+.Shape91.Height+20,RetTxtWidth('wпросмотрw'),dHeight+5,'Печать','DO suplPrn WITH .T.'

    *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
    DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Просмотр','DO suplPrn WITH .F.'
    .SetAll('ForeColor',RGB(0,0,128),'CheckBox')  
    *---------------------------------Кнопка отмена --------------------------------------------------------------------------
    DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.cont2.Width+15,.Cont1.Top,.Cont1.Width,dHeight+5,'Возврат','DO exitPrintRepZam','Возврат'
    
     DO addShape WITH 'fSupl',11,.Shape1.Left,.cont1.Top,dHeight,.shape1.Width
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH 'fSupl',12,.Shape11.Left,.Shape11.Top,dHeight,0
     .Shape12.BackStyle=1
     .Shape12.Visible=.F.  
     
     DO adLabMy WITH 'fSupl',25,'100%',.Shape11.Top+2,.Shape11.Left,.Shape11.Width,2,.F.,0
     .lab25.Visible=.F.             
    
    .Width=.Shape1.Width+40
    .Height=.Shape1.Height+.Shape91.Height+.cont1.Height+80
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*********************************************************************************************
PROCEDURE exitPrintRepZam
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
fSupl.Release
*********************************************************************************************
PROCEDURE selectVedZam
procForPrn=ALLTRIM(datprn.procved)
logWord=.F.
fSupl.checkWord.ControlSource='logWord'
fSupl.checkWord.Enabled=IIF(datprn.logEx,.T.,.F.)
*********************************************************************************************
PROCEDURE suplPrn
PARAMETERS parTerm
STORE 0 TO numrecrep
term_ch=parTerm
&procForPrn

*********************************************************************************************
PROCEDURE printRepZam
PARAMETERS parLog
IF USED('curKatKurs')
   SELECT curKatKurs
   USE
ENDIF
maxKurs=8
DIMENSION dimKurs(maxKurs,2)
FOR i=1 TO 8
    dimKurs(i,1)=''
    dimKurs(i,2)=0
ENDFOR 
SELECT * FROM sprkat INTO CURSOR curKatKurs READWRITE
ALTER TABLE curKatKurs ADD COLUMN sumtot N (10,2)
SELECT curKatKurs
INDEX ON kod TAG T1
SELECT rasp
SET FILTER TO totzp#0    
SCAN ALL
     IF SEEK(rasp.kat,'curKatKurs',1)
        REPLACE curKatKurs.sumtot WITH curKatKurs.sumtot+rasp.totzp
     ENDIF 
ENDSCAN 
SELECT curKatKurs
DELETE FOR sumtot=0
COUNT TO maxKurs1
IF maxKurs1>0
   GO TOP
   FOR i=1 TO maxKurs1
       dimKurs(i,1)=name
       dimKurs(i,2)=sumtot
       SKIP
  ENDFOR 
ENDIF 
SELECT rasp        
GO TOP 
IF parTerm
   IF logWord
      DO CASE
         CASE parLog=1  
              DO repZamToExcel
         CASE parLog=2
              DO repZamNewToExcel
      ENDCASE        
   ELSE  
      FOR ch=1 TO kvo_page 
          SELECT rasp
          GO TOP    
          DO CASE 
             CASE parLog=1           
                  REPORT FORM repzam NOCONSOLE TO PRINTER RANGE page_beg,page_end          
             CASE parLog=2
                  REPORT FORM repzamnew NOCONSOLE TO PRINTER RANGE page_beg,page_end  
          ENDCASE         
      ENDFOR   
   ENDIF    
ELSE
   SELECT rasp
   GO TOP
   DO CASE 
      CASE parLog=1           
           DO previewrep WITH 'repzam',''       
      CASE parLog=2
           DO previewrep WITH 'repzamnew',''    
    ENDCASE   
   
ENDIF
SELECT rasp
SET FILTER TO kp=kpdop  
fpodr.Refresh



*****************************************************************************************************************************************************
*                                  Печать ведомости расчёта расходов по замене отпусков
*****************************************************************************************************************************************************
PROCEDURE printraspzam
PARAMETERS par_log
fpodr.fGrid.Column8.SetFocus
STORE 0 TO numrecrep
SELECT rasp
SET FILTER TO totzp#0          
GO TOP 
DO CASE
   CASE par_log=1 
        DO printreport WITH 'repzam','Расчёт расходов по замене','rasp'
   CASE par_log=2  
        DO printreport WITH 'repzamnew','Расчёт расходов по замене','rasp'
ENDCASE   
*DO printreport WITH 'repzamtot','Расчёт расходов по замене отпусков'
SELECT rasp
SET FILTER TO kp=kpdop
fpodr.Refresh
*****************************************************************************************************************************************************
PROCEDURE repZamToExcel
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
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F.   
     .shape11.Visible=.T.     
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape12.Width=1
ENDWITH        
STORE 0 TO max_rec,one_pers,pers_ch,itDaycx,totZpcx,itDayCxTot,totZpCxTot,num_cx
DIMENSION mcx(12),zcx(12),mcxtot(12),zcxtot(12)
STORE 0 TO mcx,zcx,mcxtot,zcxtot
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=18
     .Columns(3).ColumnWidth=7   
     .Columns(4).ColumnWidth=7
     .Columns(5).ColumnWidth=7
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8
     .Columns(8).ColumnWidth=8
     .Columns(9).ColumnWidth=8
     .Columns(10).ColumnWidth=8     
     .Columns(11).ColumnWidth=8
     .Columns(12).ColumnWidth=8
     .Columns(13).ColumnWidth=8
     .Columns(14).ColumnWidth=8
     .Columns(15).ColumnWidth=8
     .Columns(16).ColumnWidth=8
     .Columns(17).ColumnWidth=8
     .Columns(18).ColumnWidth=8
     .Columns(19).ColumnWidth=8
     .Columns(20).ColumnWidth=8    
     rowcx=3     
    .Range(.Cells(rowcx,1),.Cells(rowcx,20)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Расчёт планируемых расходов на оплату труда лиц, заменяющих уходящих в отпуск работников'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                                     
     rowcx=rowcx+1   
     
                                  
     .Range(.Cells(rowcx,1),.Cells(rowcx+4,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='№ п/п'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx+4,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Наименование структурных подразделений, должностей'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH           
                 
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,5)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Количество'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH       
        
      .Range(.Cells(rowcx,6),.Cells(rowcx+4,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Среднемесячный оклад'                    
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,7),.Cells(rowcx+4,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Размер оплаты работнику в день'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,8),.Cells(rowcx,19)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Распределение дней замены и расходов по месяцам'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,20),.Cells(rowcx+4,20)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Сумма расходов на все должности'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                 
                  
      .Range(.Cells(rowcx+1,3),.Cells(rowcx+4,3)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Долж.подолеж.замене'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH       
                             
      .Range(.Cells(rowcx+1,4),.Cells(rowcx+4,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Дней отпуска'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH   
                                  
      .Range(.Cells(rowcx+1,5),.Cells(rowcx+4,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Дней замены'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                                         
      
      .Range(.Cells(rowcx+1,8),.Cells(rowcx+4,8)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='1'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,9),.Cells(rowcx+4,9)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='2'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,10),.Cells(rowcx+4,10)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='3'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,11),.Cells(rowcx+4,11)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='4'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,12),.Cells(rowcx+4,12)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='5'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,13),.Cells(rowcx+4,13)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='6'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,14),.Cells(rowcx+4,14)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='7'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,15),.Cells(rowcx+4,15)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='8'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
       
      .Range(.Cells(rowcx+1,16),.Cells(rowcx+4,16)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='9'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,17),.Cells(rowcx+4,17)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='10'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,18),.Cells(rowcx+4,18)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='11'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,19),.Cells(rowcx+4,19)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='12'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
              
      rowcx=rowcx+4
     * objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      numberRow2=rowcx+2
      rowtop=numberRow         
      SELECT rasp
      COUNT TO max_rec
      GO TOP
      kpold=0     
      DO WHILE !EOF()
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,20)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(rasp.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              numberRow2=numberRow2+1
              STORE 0 TO mcx,zcx,itDayCx,totZpCx,num_cx
              kpold=kp
           ENDIF
           num_cx=num_cx+1
           .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=num_cx
           
           .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
           objExcel.Selection.MergeCells=.T.
          * objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=IIF(SEEK(rasp.kd,'sprdolj',1),sprdolj.name,'') 
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.VerticalAlignment=1
           
           .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.kse
           objExcel.Selection.NumberFormat='0.00'
              
           .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.dotp
   
           .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.dzam
           
           .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.srzp
            
           .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.zpday     
         
           .Range(.Cells(numberRow,8),.Cells(numberRow,20)).Select
           objExcel.Selection.NumberFormat='0.00'
           .Cells(numberRow,8).Value=IIF(rasp.m1#0,rasp.m1,'')
           .Cells(numberRow,9).Value=IIF(rasp.m2#0,rasp.m2,'')
           .Cells(numberRow,10).Value=IIF(rasp.m3#0,rasp.m3,'')
           .Cells(numberRow,11).Value=IIF(rasp.m4#0,rasp.m4,'')
           .Cells(numberRow,12).Value=IIF(rasp.m5#0,rasp.m5,'')
           .Cells(numberRow,13).Value=IIF(rasp.m6#0,rasp.m6,'')           
           .Cells(numberRow,14).Value=IIF(rasp.m7#0,rasp.m7,'')
           .Cells(numberRow,15).Value=IIF(rasp.m8#0,rasp.m8,'')
           .Cells(numberRow,16).Value=IIF(rasp.m9#0,rasp.m9,'')
           .Cells(numberRow,17).Value=IIF(rasp.m10#0,rasp.m10,'')
           .Cells(numberRow,18).Value=IIF(rasp.m11#0,rasp.m11,'')
           .Cells(numberRow,19).Value=IIF(rasp.m12#0,rasp.m12,'')
           .Cells(numberRow,20).Value=IIF(rasp.itday#0,rasp.itday,'')
           
           
           .Cells(numberRow2,8).Value=IIF(z1#0,z1,'')
           .Cells(numberRow2,9).Value=IIF(z2#0,z2,'')
           .Cells(numberRow2,10).Value=IIF(z3#0,z3,'')
           .Cells(numberRow2,11).Value=IIF(z4#0,z4,'')
           .Cells(numberRow2,12).Value=IIF(z5#0,z5,'')
           .Cells(numberRow2,13).Value=IIF(z6#0,z6,'')
           .Cells(numberRow2,14).Value=IIF(z7#0,z7,'')
           .Cells(numberRow2,15).Value=IIF(z8#0,z8,'')
           .Cells(numberRow2,16).Value=IIF(z9#0,z9,'')
           .Cells(numberRow2,17).Value=IIF(z10#0,z10,'')
           .Cells(numberRow2,18).Value=IIF(z11#0,z11,'')
           .Cells(numberRow2,19).Value=IIF(z12#0,z12,'')
           .Cells(numberRow2,20).Value=IIF(totzp#0,totzp,'')
           
           FOR i=1 TO 12
               mcx(i)=mcx(i)+EVALUATE('m'+LTRIM(STR(i)))
               zcx(i)=zcx(i)+EVALUATE('z'+LTRIM(STR(i)))
               mcxtot(i)=mcxtot(i)+EVALUATE('m'+LTRIM(STR(i)))
               zcxtot(i)=zcxtot(i)+EVALUATE('z'+LTRIM(STR(i)))
               
               itDaycx=itDaycx+EVALUATE('m'+LTRIM(STR(i)))
               totZpcx=totZpcx++EVALUATE('z'+LTRIM(STR(i)))
               
               itDayCxTot=itDayCxTot+EVALUATE('m'+LTRIM(STR(i)))
               totZpCxTot=totZpCxTot+EVALUATE('z'+LTRIM(STR(i)))
               
           ENDFOR  
           numberRow=numberRow+2 
           numberRow2=numberRow2+2  
   
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fSupl.shape12.Visible=.T.
           fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
           fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch 
           SKIP 
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
           
              .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Value='По отделению'
              objExcel.Selection.VerticalAlignment=1
           
              .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.            
             
              .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
          
   
              .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.             
           
              .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.             
            
              .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
                     
              .Range(.Cells(numberRow,8),.Cells(numberRow,20)).Select
              objExcel.Selection.NumberFormat='0.00'
              .Cells(numberRow,8).Value=IIF(mcx(1)#0,mcx(1),'')
              .Cells(numberRow,9).Value=IIF(mcx(2)#0,mcx(2),'')
              .Cells(numberRow,10).Value=IIF(mcx(3)#0,mcx(3),'')
              .Cells(numberRow,11).Value=IIF(mcx(4)#0,mcx(4),'')
              .Cells(numberRow,12).Value=IIF(mcx(5)#0,mcx(5),'')
              .Cells(numberRow,13).Value=IIF(mcx(6)#0,mcx(6),'')           
              .Cells(numberRow,14).Value=IIF(mcx(7)#0,mcx(7),'')
              .Cells(numberRow,15).Value=IIF(mcx(8)#0,mcx(8),'')
              .Cells(numberRow,16).Value=IIF(mcx(9)#0,mcx(9),'')
              .Cells(numberRow,17).Value=IIF(mcx(10)#0,mcx(10),'')
              .Cells(numberRow,18).Value=IIF(mcx(11)#0,mcx(11),'')
              .Cells(numberRow,19).Value=IIF(mcx(12)#0,mcx(12),'')
              .Cells(numberRow,20).Value=IIF(itDayCx#0,itDayCx,'')
                      
              .Cells(numberRow2,8).Value=IIF(zcx(1)#0,zcx(1),'')
              .Cells(numberRow2,9).Value=IIF(zcx(2)#0,zcx(2),'')
              .Cells(numberRow2,10).Value=IIF(zcx(3)#0,zcx(3),'')
              .Cells(numberRow2,11).Value=IIF(zcx(4)#0,zcx(4),'')
              .Cells(numberRow2,12).Value=IIF(zcx(5)#0,zcx(5),'')
              .Cells(numberRow2,13).Value=IIF(zcx(6)#0,zcx(6),'')
              .Cells(numberRow2,14).Value=IIF(zcx(7)#0,zcx(7),'')
              .Cells(numberRow2,15).Value=IIF(zcx(8)#0,zcx(8),'')
              .Cells(numberRow2,16).Value=IIF(zcx(9)#0,zcx(9),'')
              .Cells(numberRow2,17).Value=IIF(zcx(10)#0,zcx(10),'')
              .Cells(numberRow2,18).Value=IIF(zcx(11)#0,zcx(11),'')
              .Cells(numberRow2,19).Value=IIF(zcx(12)#0,zcx(12),'')
              .Cells(numberRow2,20).Value=IIF(totZpCx#0,totZpCx,'')
              .Range(.Cells(numberRow,1),.Cells(numberRow2,20)).Select
              objExcel.Selection.Interior.ColorIndex=36
              numberRow=numberRow+2 
              numberRow2=numberRow2+2 
           ENDIF      
                             
      ENDDO 
     
      .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
           
      .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
      objExcel.Selection.Value='По организации'
      objExcel.Selection.VerticalAlignment=1
           
      .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.            
             
      .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
           
      .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.             
           
      .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.             
          
      .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
                    
      .Range(.Cells(numberRow,8),.Cells(numberRow,20)).Select
      objExcel.Selection.NumberFormat='0.00'
      .Cells(numberRow,8).Value=IIF(mcxtot(1)#0,mcxtot(1),'')
      .Cells(numberRow,9).Value=IIF(mcxtot(2)#0,mcxtot(2),'')
      .Cells(numberRow,10).Value=IIF(mcxtot(3)#0,mcxtot(3),'')
      .Cells(numberRow,11).Value=IIF(mcxtot(4)#0,mcxtot(4),'')
      .Cells(numberRow,12).Value=IIF(mcxtot(5)#0,mcxtot(5),'')
      .Cells(numberRow,13).Value=IIF(mcxtot(6)#0,mcxtot(6),'')           
      .Cells(numberRow,14).Value=IIF(mcxtot(7)#0,mcxtot(7),'')
      .Cells(numberRow,15).Value=IIF(mcxtot(8)#0,mcxtot(8),'')
      .Cells(numberRow,16).Value=IIF(mcxtot(9)#0,mcxtot(9),'')
      .Cells(numberRow,17).Value=IIF(mcxtot(10)#0,mcxtot(10),'')
      .Cells(numberRow,18).Value=IIF(mcxtot(11)#0,mcxtot(11),'')
      .Cells(numberRow,19).Value=IIF(mcxtot(12)#0,mcxtot(12),'')
      .Cells(numberRow,20).Value=IIF(itDayCxTot#0,itDayCxTot,'')
    
                      
      .Cells(numberRow2,8).Value=IIF(zcxtot(1)#0,zcxtot(1),'')
      .Cells(numberRow2,9).Value=IIF(zcxtot(2)#0,zcxtot(2),'')
      .Cells(numberRow2,10).Value=IIF(zcxtot(3)#0,zcxtot(3),'')
      .Cells(numberRow2,11).Value=IIF(zcxtot(4)#0,zcxtot(4),'')
      .Cells(numberRow2,12).Value=IIF(zcxtot(5)#0,zcxtot(5),'')
      .Cells(numberRow2,13).Value=IIF(zcxtot(6)#0,zcxtot(6),'')
      .Cells(numberRow2,14).Value=IIF(zcxtot(7)#0,zcxtot(7),'')
      .Cells(numberRow2,15).Value=IIF(zcxtot(8)#0,zcxtot(8),'')
      .Cells(numberRow2,16).Value=IIF(zcxtot(9)#0,zcxtot(9),'')
      .Cells(numberRow2,17).Value=IIF(zcxtot(10)#0,zcxtot(10),'')
      .Cells(numberRow2,18).Value=IIF(zcxtot(11)#0,zcxtot(11),'')
      .Cells(numberRow2,19).Value=IIF(zcxtot(12)#0,zcxtot(12),'')
      .Cells(numberRow2,20).Value=IIF(totZpCxTot#0,totZpCxTot,'') 
                                  
      .Range(.Cells(3,1),.Cells(numberRow2,20)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
          
      .Range(.Cells(3,1),.Cells(numberRow2-1,20)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      *objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
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
=SYS(2002)
=INKEY(2)
WITH fSupl
     .Shape12.Visible=.F.
     .Shape11.Visible=.F.  
     .lab25.Visible=.F. 
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.
ENDWITH               
objExcel.Visible=.T.


*****************************************************************************************************************************************************
PROCEDURE repZamNewToExcel
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
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F.   
     .shape11.Visible=.T.     
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape12.Width=1
ENDWITH        
STORE 0 TO max_rec,one_pers,pers_ch,itDaycx,totZpcx,itDayCxTot,totZpCxTot,num_cx
DIMENSION mcx(12),zcx(12),mcxtot(12),zcxtot(12)
STORE 0 TO mcx,zcx,mcxtot,zcxtot
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 1
     .pageSetup.printTitleRows="$11:$11"
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=20
     .Columns(3).ColumnWidth=7   
     .Columns(4).ColumnWidth=7
     .Columns(5).ColumnWidth=7
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
          .Value='Расчёт планируемых расходов на дополнительную оплату труда работников, заменяющих уходящих в отпуск работников'
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
          .Value='по '+ALLTRIM(office)
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
          .Value='на '+STR(YEAR(varDTar),4)+' год'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                                     
     rowcx=rowcx+1    
                                  
     .Range(.Cells(rowcx,1),.Cells(rowcx+4,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='№ п/п'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx+4,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Наименование условия оказания медицинской помощи,иного вида деятельности,  структурных подразделений, должностей'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH           
                 
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,5)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Количество'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH       
        
      .Range(.Cells(rowcx,6),.Cells(rowcx+4,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Среднемесячный оклад (ставка) по списку окладов, руб.'                    
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,7),.Cells(rowcx+4,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Размер оплаты работнику в день, руб.'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH   
       .Range(.Cells(rowcx,8),.Cells(rowcx+4,8)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Сумма расходов на год на все должности'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                                           
    
                  
      .Range(.Cells(rowcx+1,3),.Cells(rowcx+4,3)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Долж., проф. подолеж.замене'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH       
                             
      .Range(.Cells(rowcx+1,4),.Cells(rowcx+4,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Дней отпуска'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH   
                                  
      .Range(.Cells(rowcx+1,5),.Cells(rowcx+4,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Дней замены'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                                         
      
      rowcx=rowcx+5
     * objExcel.Selection.HorizontalAlignment=xlCenter
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
      numberRow2=rowcx+2
      rowtop=numberRow         
      SELECT rasp
      COUNT TO max_rec
      GO TOP
      kpold=0     
      DO WHILE !EOF()
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,8)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(rasp.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              numberRow2=numberRow2+1
              STORE 0 TO mcx,zcx,itDayCx,totZpCx,num_cx
              kpold=kp
           ENDIF
           num_cx=num_cx+1
           .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=num_cx
           
           .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
           objExcel.Selection.MergeCells=.T.
          * objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=IIF(SEEK(rasp.kd,'sprdolj',1),sprdolj.name,'') 
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.VerticalAlignment=1
           
           .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.kse
           objExcel.Selection.NumberFormat='0.00'
              
           .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.dotp
   
           .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.dzam
           
           .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.NumberFormat='0.00'
           objExcel.Selection.Value=rasp.srzp
           
                       
           .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.NumberFormat='0.00'   
           objExcel.Selection.Value=rasp.zpday 
        
         
           .Range(.Cells(numberRow,8),.Cells(numberRow,8)).Select
           objExcel.Selection.NumberFormat='0.00'
           .Cells(numberRow,8).Value=IIF(rasp.itday#0,rasp.itday,'')
           .Cells(numberRow2,8).NumberFormat='#####.##'
           .Cells(numberRow2,8).Value=IIF(totzp#0,totzp,'')
         
                       
            FOR i=1 TO 12
               mcx(i)=mcx(i)+EVALUATE('m'+LTRIM(STR(i)))
               zcx(i)=zcx(i)+EVALUATE('z'+LTRIM(STR(i)))
               mcxtot(i)=mcxtot(i)+EVALUATE('m'+LTRIM(STR(i)))
               zcxtot(i)=zcxtot(i)+EVALUATE('z'+LTRIM(STR(i)))
               
               itDaycx=itDaycx+EVALUATE('m'+LTRIM(STR(i)))
               totZpcx=totZpcx++EVALUATE('z'+LTRIM(STR(i)))
               
               itDayCxTot=itDayCxTot+EVALUATE('m'+LTRIM(STR(i)))
               totZpCxTot=totZpCxTot+EVALUATE('z'+LTRIM(STR(i)))
               
           ENDFOR              
                       
           numberRow=numberRow+2 
           numberRow2=numberRow2+2  
   
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fSupl.shape12.Visible=.T.
           fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
           fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch 
           SKIP 
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
           
              .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Value='По отделению'
              objExcel.Selection.VerticalAlignment=1
           
              .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.            
             
              .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
          
   
              .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.             
           
              .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.             
            
              .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
                     
              .Range(.Cells(numberRow,8),.Cells(numberRow2,8)).Select
              objExcel.Selection.NumberFormat='0.00'
              .Cells(numberRow,8).Value=IIF(itDayCx#0,itDayCx,'')
                      
              .Cells(numberRow2,8).Value=IIF(totZpCx#0,totZpCx,'')
              .Range(.Cells(numberRow,1),.Cells(numberRow2,8)).Select
              
              objExcel.Selection.Interior.ColorIndex=36
              numberRow=numberRow+2 
              numberRow2=numberRow2+2 
           ENDIF      
                             
      ENDDO 
     
      .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
           
      .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
      objExcel.Selection.Value='По организации'
      objExcel.Selection.VerticalAlignment=1
           
      .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.            
             
      .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
           
      .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.             
           
      .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.             
          
      .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
                    
      .Range(.Cells(numberRow,8),.Cells(numberRow,80)).Select
      objExcel.Selection.NumberFormat='0.00'     
     
      .Cells(numberRow,8).Value=IIF(itDayCxTot#0,itDayCxTot,'')   
                            
      .Cells(numberRow2,8).Value=IIF(totZpCxTot#0,totZpCxTot,'')                                   
      .Range(.Cells(6,1),.Cells(numberRow2,8)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
          
      .Range(.Cells(3,1),.Cells(numberRow2-1,8)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      *objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
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
=SYS(2002)
=INKEY(2)
WITH fSupl
     .Shape12.Visible=.F.
     .Shape11.Visible=.F.  
     .lab25.Visible=.F. 
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.
ENDWITH               
objExcel.Visible=.T.
*****************************************************************************************************************************************************
*                                  Печать ведомости расчёта расходов по замене отпусков
*****************************************************************************************************************************************************
PROCEDURE printRaspZamTotNew
*fpodr.fGrid.Column8.SetFocus
IF USED('curKatKurs')
   SELECT curKatKurs
   USE
ENDIF
maxKurs=8
DIMENSION dimKurs(maxKurs,2)
FOR i=1 TO 8
    dimKurs(i,1)=''
    dimKurs(i,2)=0
ENDFOR 
SELECT * FROM sprkat INTO CURSOR curKatKurs READWRITE
ALTER TABLE curKatKurs ADD COLUMN sumtot N (10,2)
SELECT curKatKurs
INDEX ON kod TAG T1

STORE 0 TO numrecrep
SELECT rasp
=AFIELDS(arRasp,'rasp') 
CREATE CURSOR currasp FROM ARRAY arRasp
SELECT currasp
APPEND FROM rasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1  
SET FILTER TO totzp>0.AND.kp>0.AND.nd>0     
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
SET ORDER TO 2
SELECT datJob
SELECT * FROM datJob WHERE SEEK(STR(datJob.kp,3)+STR(datJob.kd,3),'currasp',2).AND.datJob.zptot>0 INTO CURSOR curpeople READWRITE  
SELECT curpeople
SET ORDER TO
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curpeople->kp,3)+STR(curpeople->kd,3)   
   SELECT curpeople
   REPLACE nd WITH currasp->nd,np WITH currasp->np  
   IF rashset(5)  
      sumrep=0
      FOR i=1 TO 13      
          IF i<13
             repd='curpeople->dst'+LTRIM(STR(i))
             REPLACE &repd WITH IIF(&repd>0.AND.&repd<1,1,ROUND(&repd,0)) 
             sumrep=sumrep+&repd
          ELSE 
             REPLACE dsttot WITH sumrep
          ENDIF 
      ENDFOR 
   ENDIF  
   SKIP 
ENDDO 
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
IF rashset(5)
   DIMENSION dim_m(12)
   STORE 0 TO dim_m
   SELECT currasp
   GO TOP
   DO WHILE !EOF()
      STORE 0 TO dim_m
      SELECT curpeople
      SEEK STR(currasp->np,3)+STR(currasp->nd,3)
      SCAN WHILE np=currasp->np.AND.nd=currasp->nd
           FOR i=1 TO 12
               dim_m(i)=dim_m(i)+EVAL('dst'+LTRIM(STR(i)))             
           ENDFOR
      ENDSCAN
      SELECT currasp
      sumrep=0
      FOR i=1 TO 12
          repd='m'+LTRIM(STR(i))
          REPLACE &repd WITH dim_m(i)
          sumrep=sumrep+&repd
      ENDFOR 
      REPLACE itday WITH sumrep
      SKIP 
   ENDDO
ENDIF    
SELECT curpeople
SCAN ALL
     IF SEEK(curPeople.kat,'curKatKurs',1)
        REPLACE curKatKurs.sumtot WITH curKatKurs.sumtot+curPeople.zpTot
     ENDIF 
ENDSCAN 
SELECT curKatKurs
DELETE FOR sumtot=0
COUNT TO maxKurs1
IF maxKurs1>0
   GO TOP
   FOR i=1 TO maxKurs1
       dimKurs(i,1)=name
       dimKurs(i,2)=sumtot
       SKIP
  ENDFOR 
ENDIF   
SELECT curPeople  
SET ORDER TO 1
GO TOP 


IF parTerm
   IF logWord
      DO repZamTotToExcel
   ELSE  
      FOR ch=1 TO kvo_page 
          SELECT curpeople
          GO TOP         
          REPORT FORM repzamtot NOCONSOLE TO PRINTER RANGE page_beg,page_end                          
      ENDFOR   
   ENDIF    
ELSE
   SELECT curpeople
   GO TOP            
   DO previewrep WITH 'repzamtot',''           
ENDIF

*****************************************************************************************************************************************************
PROCEDURE repZamTotToExcel
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
WITH fSupl
     .cont1.Visible=.F.
     .cont2.Visible=.F.
     .cont3.Visible=.F.   
     .shape11.Visible=.T.     
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
     .Shape12.Width=1
ENDWITH        
STORE 0 TO max_rec,one_pers,pers_ch,itDaycx,totZpcx,itDayCxTot,totZpCxTot,num_cx
DIMENSION mcx(12),zcx(12),mcxtot(12),zcxtot(12)
STORE 0 TO mcx,zcx,mcxtot,zcxtot
WITH excelBook.Sheets(1)
     .PageSetup.Orientation = 2
     .Columns(1).ColumnWidth=3
     .Columns(2).ColumnWidth=18
     .Columns(3).ColumnWidth=7   
     .Columns(4).ColumnWidth=7
     .Columns(5).ColumnWidth=7
     .Columns(6).ColumnWidth=8
     .Columns(7).ColumnWidth=8
     .Columns(8).ColumnWidth=8
     .Columns(9).ColumnWidth=8
     .Columns(10).ColumnWidth=8     
     .Columns(11).ColumnWidth=8
     .Columns(12).ColumnWidth=8
     .Columns(13).ColumnWidth=8
     .Columns(14).ColumnWidth=8
     .Columns(15).ColumnWidth=8
     .Columns(16).ColumnWidth=8
     .Columns(17).ColumnWidth=8
     .Columns(18).ColumnWidth=8
     .Columns(19).ColumnWidth=8
     .Columns(20).ColumnWidth=8    
     rowcx=3     
    .Range(.Cells(rowcx,1),.Cells(rowcx,20)).Select  
     WITH objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Расчёт планируемых расходов на оплату труда лиц, заменяющих уходящих в отпуск работников'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH                                     
     rowcx=rowcx+1   
     
                                  
     .Range(.Cells(rowcx,1),.Cells(rowcx+4,1)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='№ п/п'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH          
         
     .Range(.Cells(rowcx,2),.Cells(rowcx+4,2)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Наименование структурных подразделений, должностей'
          .Font.Name='Times New Roman'   
          .Font.Size=9
     ENDWITH           
                 
                     
     .Range(.Cells(rowcx,3),.Cells(rowcx,5)).Select
     With objExcel.Selection
          .MergeCells=.T.
          .HorizontalAlignment=xlCenter
          .VerticalAlignment=1
          .WrapText=.T.
          .Value='Количество'
          .Font.Name='Times New Roman'   
          .Font.Size=9
      ENDWITH       
        
      .Range(.Cells(rowcx,6),.Cells(rowcx+4,6)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Среднемесячный оклад'                    
           .Font.Name='Times New Roman'   
           .Font.Size=9
      ENDWITH                                                 
        
      .Range(.Cells(rowcx,7),.Cells(rowcx+4,7)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Размер оплаты работнику в день'   
           .Font.Name='Times New Roman'   
           .Font.Size=8                 
      ENDWITH                                        
           
      .Range(.Cells(rowcx,8),.Cells(rowcx,19)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Распределение дней замены и расходов по месяцам'                    
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                        
                      
      .Range(.Cells(rowcx,20),.Cells(rowcx+4,20)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Сумма расходов на все должности'              
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                 
                  
      .Range(.Cells(rowcx+1,3),.Cells(rowcx+4,3)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Долж.подолеж.замене'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH       
                             
      .Range(.Cells(rowcx+1,4),.Cells(rowcx+4,4)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Дней отпуска'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH   
                                  
      .Range(.Cells(rowcx+1,5),.Cells(rowcx+4,5)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='Дней замены'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH                                         
      
      .Range(.Cells(rowcx+1,8),.Cells(rowcx+4,8)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='1'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,9),.Cells(rowcx+4,9)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='2'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,10),.Cells(rowcx+4,10)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='3'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,11),.Cells(rowcx+4,11)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='4'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,12),.Cells(rowcx+4,12)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='5'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,13),.Cells(rowcx+4,13)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='6'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,14),.Cells(rowcx+4,14)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='7'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,15),.Cells(rowcx+4,15)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='8'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
       
      .Range(.Cells(rowcx+1,16),.Cells(rowcx+4,16)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='9'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,17),.Cells(rowcx+4,17)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='10'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH 
      
      .Range(.Cells(rowcx+1,18),.Cells(rowcx+4,18)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='11'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
      
      .Range(.Cells(rowcx+1,19),.Cells(rowcx+4,19)).Select
      With objExcel.Selection
           .MergeCells=.T.
           .HorizontalAlignment=xlCenter
           .VerticalAlignment=1
           .WrapText=.T.
           .Value='12'
           .Font.Name='Times New Roman'   
           .Font.Size=8
      ENDWITH  
              
      rowcx=rowcx+4
     * objExcel.Selection.HorizontalAlignment=xlCenter
      numberRow=rowcx+1  
      numberRow2=rowcx+2
      rowtop=numberRow         
      SELECT curpeople
      SET ORDER TO 1
      COUNT TO max_rec
      GO TOP
      kpold=0     
      DO WHILE !EOF()
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow,20)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.HorizontalAlignment=xlLeft
              objExcel.Selection.VerticalAlignment=1
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Interior.ColorIndex=37
              objExcel.Selection.Value=IIF(SEEK(rasp.kp,'sprpodr',1),sprpodr.name,'')                   
              numberRow=numberRow+1
              numberRow2=numberRow2+1
              STORE 0 TO mcx,zcx,itDayCx,totZpCx,num_cx
              kpold=kp
           ENDIF
           num_cx=num_cx+1
           .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=num_cx
           
           .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
           objExcel.Selection.MergeCells=.T.
          * objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=IIF(SEEK(rasp.kd,'sprdolj',1),sprdolj.name,'') 
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.VerticalAlignment=1
           
           .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.kse
           objExcel.Selection.NumberFormat='0.00'
              
           .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.dotp
   
           .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.dzam
           
           .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.srzp
            
           .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
           objExcel.Selection.MergeCells=.T.
           objExcel.Selection.WrapText=.T.
           objExcel.Selection.Value=rasp.zpday     
         
           .Range(.Cells(numberRow,8),.Cells(numberRow,20)).Select
           objExcel.Selection.NumberFormat='0.00'
           .Cells(numberRow,8).Value=IIF(rasp.m1#0,rasp.m1,'')
           .Cells(numberRow,9).Value=IIF(rasp.m2#0,rasp.m2,'')
           .Cells(numberRow,10).Value=IIF(rasp.m3#0,rasp.m3,'')
           .Cells(numberRow,11).Value=IIF(rasp.m4#0,rasp.m4,'')
           .Cells(numberRow,12).Value=IIF(rasp.m5#0,rasp.m5,'')
           .Cells(numberRow,13).Value=IIF(rasp.m6#0,rasp.m6,'')           
           .Cells(numberRow,14).Value=IIF(rasp.m7#0,rasp.m7,'')
           .Cells(numberRow,15).Value=IIF(rasp.m8#0,rasp.m8,'')
           .Cells(numberRow,16).Value=IIF(rasp.m9#0,rasp.m9,'')
           .Cells(numberRow,17).Value=IIF(rasp.m10#0,rasp.m10,'')
           .Cells(numberRow,18).Value=IIF(rasp.m11#0,rasp.m11,'')
           .Cells(numberRow,19).Value=IIF(rasp.m12#0,rasp.m12,'')
           .Cells(numberRow,20).Value=IIF(rasp.itday#0,rasp.itday,'')
           
           
           .Cells(numberRow2,8).Value=IIF(z1#0,z1,'')
           .Cells(numberRow2,9).Value=IIF(z2#0,z2,'')
           .Cells(numberRow2,10).Value=IIF(z3#0,z3,'')
           .Cells(numberRow2,11).Value=IIF(z4#0,z4,'')
           .Cells(numberRow2,12).Value=IIF(z5#0,z5,'')
           .Cells(numberRow2,13).Value=IIF(z6#0,z6,'')
           .Cells(numberRow2,14).Value=IIF(z7#0,z7,'')
           .Cells(numberRow2,15).Value=IIF(z8#0,z8,'')
           .Cells(numberRow2,16).Value=IIF(z9#0,z9,'')
           .Cells(numberRow2,17).Value=IIF(z10#0,z10,'')
           .Cells(numberRow2,18).Value=IIF(z11#0,z11,'')
           .Cells(numberRow2,19).Value=IIF(z12#0,z12,'')
           .Cells(numberRow2,20).Value=IIF(totzp#0,totzp,'')
           
           FOR i=1 TO 12
               mcx(i)=mcx(i)+EVALUATE('m'+LTRIM(STR(i)))
               zcx(i)=zcx(i)+EVALUATE('z'+LTRIM(STR(i)))
               mcxtot(i)=mcxtot(i)+EVALUATE('m'+LTRIM(STR(i)))
               zcxtot(i)=zcxtot(i)+EVALUATE('z'+LTRIM(STR(i)))
               
               itDaycx=itDaycx+EVALUATE('m'+LTRIM(STR(i)))
               totZpcx=totZpcx++EVALUATE('z'+LTRIM(STR(i)))
               
               itDayCxTot=itDayCxTot+EVALUATE('m'+LTRIM(STR(i)))
               totZpCxTot=totZpCxTot+EVALUATE('z'+LTRIM(STR(i)))
               
           ENDFOR  
           numberRow=numberRow+2 
           numberRow2=numberRow2+2  
   
           one_pers=one_pers+1
           pers_ch=one_pers/max_rec*100
           fSupl.shape12.Visible=.T.
           fSupl.lab25.Caption=LTRIM(STR(pers_ch))+'%'       
           fSupl.Shape12.Width=fSupl.shape11.Width/100*pers_ch 
           SKIP 
           IF kp#kpold
              .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
           
              .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
              objExcel.Selection.Value='По отделению'
              objExcel.Selection.VerticalAlignment=1
           
              .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.            
             
              .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
          
   
              .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.             
           
              .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.             
            
              .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
              objExcel.Selection.MergeCells=.T.
              objExcel.Selection.WrapText=.T.
                     
              .Range(.Cells(numberRow,8),.Cells(numberRow,20)).Select
              objExcel.Selection.NumberFormat='0.00'
              .Cells(numberRow,8).Value=IIF(mcx(1)#0,mcx(1),'')
              .Cells(numberRow,9).Value=IIF(mcx(2)#0,mcx(2),'')
              .Cells(numberRow,10).Value=IIF(mcx(3)#0,mcx(3),'')
              .Cells(numberRow,11).Value=IIF(mcx(4)#0,mcx(4),'')
              .Cells(numberRow,12).Value=IIF(mcx(5)#0,mcx(5),'')
              .Cells(numberRow,13).Value=IIF(mcx(6)#0,mcx(6),'')           
              .Cells(numberRow,14).Value=IIF(mcx(7)#0,mcx(7),'')
              .Cells(numberRow,15).Value=IIF(mcx(8)#0,mcx(8),'')
              .Cells(numberRow,16).Value=IIF(mcx(9)#0,mcx(9),'')
              .Cells(numberRow,17).Value=IIF(mcx(10)#0,mcx(10),'')
              .Cells(numberRow,18).Value=IIF(mcx(11)#0,mcx(11),'')
              .Cells(numberRow,19).Value=IIF(mcx(12)#0,mcx(12),'')
              .Cells(numberRow,20).Value=IIF(itDayCx#0,itDayCx,'')
                      
              .Cells(numberRow2,8).Value=IIF(zcx(1)#0,zcx(1),'')
              .Cells(numberRow2,9).Value=IIF(zcx(2)#0,zcx(2),'')
              .Cells(numberRow2,10).Value=IIF(zcx(3)#0,zcx(3),'')
              .Cells(numberRow2,11).Value=IIF(zcx(4)#0,zcx(4),'')
              .Cells(numberRow2,12).Value=IIF(zcx(5)#0,zcx(5),'')
              .Cells(numberRow2,13).Value=IIF(zcx(6)#0,zcx(6),'')
              .Cells(numberRow2,14).Value=IIF(zcx(7)#0,zcx(7),'')
              .Cells(numberRow2,15).Value=IIF(zcx(8)#0,zcx(8),'')
              .Cells(numberRow2,16).Value=IIF(zcx(9)#0,zcx(9),'')
              .Cells(numberRow2,17).Value=IIF(zcx(10)#0,zcx(10),'')
              .Cells(numberRow2,18).Value=IIF(zcx(11)#0,zcx(11),'')
              .Cells(numberRow2,19).Value=IIF(zcx(12)#0,zcx(12),'')
              .Cells(numberRow2,20).Value=IIF(totZpCx#0,totZpCx,'')
              .Range(.Cells(numberRow,1),.Cells(numberRow2,20)).Select
              objExcel.Selection.Interior.ColorIndex=36
              numberRow=numberRow+2 
              numberRow2=numberRow2+2 
           ENDIF      
                             
      ENDDO 
     
      .Range(.Cells(numberRow,1),.Cells(numberRow2,1)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
           
      .Range(.Cells(numberRow,2),.Cells(numberRow2,2)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
      objExcel.Selection.Value='По организации'
      objExcel.Selection.VerticalAlignment=1
           
      .Range(.Cells(numberRow,3),.Cells(numberRow2,3)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.            
             
      .Range(.Cells(numberRow,4),.Cells(numberRow2,4)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
           
      .Range(.Cells(numberRow,5),.Cells(numberRow2,5)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.             
           
      .Range(.Cells(numberRow,6),.Cells(numberRow2,6)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.             
          
      .Range(.Cells(numberRow,7),.Cells(numberRow2,7)).Select
      objExcel.Selection.MergeCells=.T.
      objExcel.Selection.WrapText=.T.
                    
      .Range(.Cells(numberRow,8),.Cells(numberRow,20)).Select
      objExcel.Selection.NumberFormat='0.00'
      .Cells(numberRow,8).Value=IIF(mcxtot(1)#0,mcxtot(1),'')
      .Cells(numberRow,9).Value=IIF(mcxtot(2)#0,mcxtot(2),'')
      .Cells(numberRow,10).Value=IIF(mcxtot(3)#0,mcxtot(3),'')
      .Cells(numberRow,11).Value=IIF(mcxtot(4)#0,mcxtot(4),'')
      .Cells(numberRow,12).Value=IIF(mcxtot(5)#0,mcxtot(5),'')
      .Cells(numberRow,13).Value=IIF(mcxtot(6)#0,mcxtot(6),'')           
      .Cells(numberRow,14).Value=IIF(mcxtot(7)#0,mcxtot(7),'')
      .Cells(numberRow,15).Value=IIF(mcxtot(8)#0,mcxtot(8),'')
      .Cells(numberRow,16).Value=IIF(mcxtot(9)#0,mcxtot(9),'')
      .Cells(numberRow,17).Value=IIF(mcxtot(10)#0,mcxtot(10),'')
      .Cells(numberRow,18).Value=IIF(mcxtot(11)#0,mcxtot(11),'')
      .Cells(numberRow,19).Value=IIF(mcxtot(12)#0,mcxtot(12),'')
      .Cells(numberRow,20).Value=IIF(itDayCxTot#0,itDayCxTot,'')
    
                      
      .Cells(numberRow2,8).Value=IIF(zcxtot(1)#0,zcxtot(1),'')
      .Cells(numberRow2,9).Value=IIF(zcxtot(2)#0,zcxtot(2),'')
      .Cells(numberRow2,10).Value=IIF(zcxtot(3)#0,zcxtot(3),'')
      .Cells(numberRow2,11).Value=IIF(zcxtot(4)#0,zcxtot(4),'')
      .Cells(numberRow2,12).Value=IIF(zcxtot(5)#0,zcxtot(5),'')
      .Cells(numberRow2,13).Value=IIF(zcxtot(6)#0,zcxtot(6),'')
      .Cells(numberRow2,14).Value=IIF(zcxtot(7)#0,zcxtot(7),'')
      .Cells(numberRow2,15).Value=IIF(zcxtot(8)#0,zcxtot(8),'')
      .Cells(numberRow2,16).Value=IIF(zcxtot(9)#0,zcxtot(9),'')
      .Cells(numberRow2,17).Value=IIF(zcxtot(10)#0,zcxtot(10),'')
      .Cells(numberRow2,18).Value=IIF(zcxtot(11)#0,zcxtot(11),'')
      .Cells(numberRow2,19).Value=IIF(zcxtot(12)#0,zcxtot(12),'')
      .Cells(numberRow2,20).Value=IIF(totZpCxTot#0,totZpCxTot,'') 
                                  
      .Range(.Cells(3,1),.Cells(numberRow2,20)).Select
      objExcel.Selection.Borders(xlEdgeLeft).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeTop).Weight=xlThin            
      objExcel.Selection.Borders(xlEdgeBottom).Weight=xlThin
      objExcel.Selection.Borders(xlEdgeRight).Weight=xlThin
      objExcel.Selection.Borders(xlInsideVertical).Weight=xlThin
      objExcel.Selection.Borders(xlInsideHorizontal).Weight=xlThin
      objExcel.Selection.VerticalAlignment=1
          
      .Range(.Cells(3,1),.Cells(numberRow2-1,20)).Select
      objExcel.Selection.Font.Name='Times New Roman' 
      objExcel.Selection.Font.Size=8      
      *objExcel.Selection.WrapText=.T.  
      .Cells(1,1).Select                       
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
=SYS(2002)
=INKEY(2)
WITH fSupl
     .Shape12.Visible=.F.
     .Shape11.Visible=.F.  
     .lab25.Visible=.F. 
     .cont1.Visible=.T.
     .cont2.Visible=.T.
     .cont3.Visible=.T.
ENDWITH               
objExcel.Visible=.T.



*****************************************************************************************************************************************************
*                                  Печать ведомости расчёта расходов по замене отпусков
*****************************************************************************************************************************************************
PROCEDURE printraspzamtot
fpodr.fGrid.Column8.SetFocus
STORE 0 TO numrecrep
SELECT rasp
=AFIELDS(arRasp,'rasp') 
CREATE CURSOR currasp FROM ARRAY arRasp
SELECT currasp
APPEND FROM rasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1  
SET FILTER TO totzp>0.AND.kp>0.AND.nd>0     
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
SET ORDER TO 2
SELECT datJob
SELECT * FROM datJob WHERE SEEK(STR(datJob.people.kp,3)+STR(datJob.kd,3),'currasp',2).AND.datJob.zptot>0 INTO CURSOR curpeople READWRITE  
SELECT curpeople
SET ORDER TO
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curpeople->kp,3)+STR(curpeople->kd,3)   
   SELECT curpeople
   REPLACE nd WITH currasp->nd,np WITH currasp->np  
   IF rashset(5)  
      sumrep=0
      FOR i=1 TO 13      
          IF i<13
             repd='curpeople->dst'+LTRIM(STR(i))
             REPLACE &repd WITH IIF(&repd>0.AND.&repd<1,1,ROUND(&repd,0)) 
             sumrep=sumrep+&repd
          ELSE 
             REPLACE dsttot WITH sumrep
          ENDIF 
      ENDFOR 
   ENDIF  
   SKIP 
ENDDO 
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
IF rashset(5)
   DIMENSION dim_m(12)
   STORE 0 TO dim_m
   SELECT currasp
   GO TOP
   DO WHILE !EOF()
      STORE 0 TO dim_m
      SELECT curpeople
      SEEK STR(currasp->np,3)+STR(currasp->nd,3)
      SCAN WHILE np=currasp->np.AND.nd=currasp->nd
           FOR i=1 TO 12
               dim_m(i)=dim_m(i)+EVAL('dst'+LTRIM(STR(i)))             
           ENDFOR
      ENDSCAN
      SELECT currasp
      sumrep=0
      FOR i=1 TO 12
          repd='m'+LTRIM(STR(i))
          REPLACE &repd WITH dim_m(i)
          sumrep=sumrep+&repd
      ENDFOR 
      REPLACE itday WITH sumrep
      SKIP 
   ENDDO
ENDIF    
SELECT curpeople
SET ORDER TO 1
GO TOP 
DO printreport WITH 'repzamtot','Расчёт расходов по замене отпусков','curpeople'
SELECT rasp
SET FILTER TO kp=kpdop
fpodr.Refresh
****************************************************************************************************************************************************
PROCEDURE topzam
PARAMETERS dim_ch
DO topToolPreview WITH .T.
*****************************************************************************************************************************************************
PROCEDURE toppodrforzam
PARAMETERS par_ch
STORE 0 TO numrecrep,sum_podr,ksekat_podr
*************************************************************************************************************************************************
PROCEDURE retnumreczam
PARAMETERS parnum
numrecrep=numrecrep+1
parnum=numrecrep
RETURN parnum

***************************************************************************************************************************************************
*                   Процедура для настрек по работе с заменой отпусков
***************************************************************************************************************************************************
PROCEDURE setupzam
fsetup=CREATEOBJECT('FORMMY')
WITH fsetup
     .BackColor=RGB(255,255,255)
      DO addShape WITH 'fSetup',1,10,10,dHeight,0,8      
     .procexit='DO exitfsetup'   

     DO adLabMy WITH 'fsetup',1,'Среднее кол-во дней',fsetup.Shape1.Top+10,fsetup.Shape1.Left+10,150,0,.T.
     DO adtbox WITH 'fsetup',1,fsetup.Lab1.Left+fsetup.lab1.Width+5,fsetup.lab1.Top,RetTxtWidth('9999999'),dHeight,'rashset(4)','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashset'
     
     .lab1.Top=.txtbox1.Top+(.txtbox1.Height-.lab1.Height)
     DO adLabMy WITH 'fsetup',2,'Дата отсчёта',fsetup.lab1.Top,fsetup.txtBox1.Left+fSetup.txtBox1.Width+10,150,0,.T.
     DO adtbox WITH 'fsetup',2,fsetup.lab2.Left+fSetup.lab2.Width+5,fsetup.txtbox1.Top,RetTxtWidth('99/99/999999'),dHeight,'countDate','Z',.T.,1

     DO adCheckBox WITH 'fsetup','check1','редактировать суммы - должность',fsetup.txtbox2.Top+fsetup.txtbox2.Height+10,fsetup.lab1.Left,150,dHeight,'rashset(1)',0
     DO adCheckBox WITH 'fsetup','check2','редактировать суммы - персонал',fsetup.check1.Top+fsetup.check1.Height+10,fsetup.lab1.Left,150,dHeight,'rashset(2)',0
     DO adCheckBox WITH 'fsetup','check3','округлять дни в полной ведомости',fsetup.check2.Top+fsetup.check2.Height+10,fsetup.lab1.Left,150,dHeight,'rashset(5)',0
     DO adCheckBox WITH 'fsetup','checkRound','учитывать округление в общей сумме по должности',fSetup.check3.Top+fSetup.check3.Height+10,fSetup.check1.Left,150,dHeight,'rashset(9)',0,.F.,'SAVE TO &var_path ALL LIKE rashset'
 
     .check2.ProcForValid='SAVE TO &var_path ALL LIKE rashset'
     .Shape1.Height=dHeight*5+50 
     .Shape1.Width=.lab1.Width+.txtBox1.Width+.lab2.Width+.txtBox2.Width+50       
     .Caption='настройки'
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+20
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fsetup'
fsetup.Show
**************************************************************************************************************************
PROCEDURE exitfsetup
fPodr.fGrid.Columns(fPodr.fGrid.ColumnCount).SetFocus 
SAVE TO &var_path ALL LIKE rashset
fsetup.Release
***************************************************************************************************************************
*                                    Меню для печати ведомостей
***************************************************************************************************************************
PROCEDURE menuprn
CurTotHeight=35
row_pop=CurTotHeight/FONTMETRIC(1,dFontName,dFontSize)
col_pop=fpodr.menucont5.LEFT/FONTMETRIC(6,dFontName,dFontSize)
DEFINE POPUP menuprint FROM row_pop,col_pop SHORTCUT MARGIN FONT dFontName,dFontSize  COLOR SCHEME 4
DEFINE BAR 1 OF menuprint PROMPT "Сокращенная ведомость замены"
DEFINE BAR 2 OF menuprint PROMPT "\-"
DEFINE BAR 3 OF menuprint PROMPT "Сокращенная ведомость замены(новая)"
DEFINE BAR 4 OF menuprint PROMPT "\-"
DEFINE BAR 5 OF menuprint PROMPT "Полная ведомость замены"
ON SELECTION BAR 1 OF menuprint DO printraspzam WITH 1
ON SELECTION BAR 3 OF menuprint DO printraspzam WITH 2
ON SELECTION BAR 5 OF menuprint DO printraspzamtot
ACTIVATE POPUP menuprint
