DIMENSION dim_tot(13),dim_day(13)
STORE 0 TO dim_tot,dim_day
RESTORE FROM rashset ADDITIVE
dim_day(10)=50
dim_day(6)=20
formulaKurs='datjob.mtokl+datjob.mstsum+datjob.mvto+datjob.mkat+datjob.mchir+datjob.mcharw'
formulKursFull='datjob.tokl+datjob.stsum+datjob.svto+datjob.skat+datjob.schir+datjob.charw'
countDate=varDtar
log_form=1
SELECT datJob
SET FILTER TO 
SET ORDER TO 2
SELECT * FROM sprdolj INTO CURSOR curSupDolj READWRITE
SELECT curSupDolj
INDEX ON kod TAG T1

SELECT * FROM rasp INTO CURSOR doprasp READWRITE
SELECT doprasp
REPLACE srzpk WITH 0 ALL
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET RELATION TO kp INTO cursprpodr,kd INTO curSupDolj ADDITIVE
GO TOP
DO WHILE !EOF()
   SELECT datjob
   SEEK STR(doprasp.kp,3)+STR(doprasp.kd,3)
    STORE 0 TO nms,nkse,npeop
   DO WHILE kp=doprasp.kp.AND.kd=doprasp.kd
      npeop=npeop+1
      nkse=nkse+datjob.kse  
      IF !doprasp.lOkl      
         *nms=nms+&formulaOtp
      ELSE 
         nms=nms+datjob.mtokl
      ENDIF 
      SKIP 
   ENDDO  
   SELECT doprasp
   REPLACE ksekurs WITH IIF(ksekurs=0,doprasp.kse,ksekurs),srzpk WITH IIF(nkse<1,nms,nms/nkse),zpdayk WITH srzpk/rashset(8),;
   primpodr WITH IIF(nd=1,cursprpodr->name,''),pldop WITH curSupDolj.name,kpeopk WITH npeop   
   SKIP
ENDDO 
GO TOP
CREATE CURSOR doppeople (podr C (50),dolj C(50),name C(50),kse N(5,2),msf N(8))
log_podr=.F.
SELECT rasp
SET RELATION TO kd INTO sprdolj ADDITIVE
GO TOP
DO WHILE !EOF()
   SELECT datJob
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   DO WHILE kp=rasp.kp.AND.kd=rasp.kd      
      SELECT doppeople
      APPEND BLANK
      *REPLACE name WITH datJob.fio,kse WITH datJob.kse,msf WITH &formulaOtp,dolj WITH sprdolj.name,;
      *podr WITH IIF(SEEK(rasp.kp,'sprpodr',1).AND.rasp.nd=1.AND.log_podr=.F.,sprpodr.name,'')
      SELECT datJob
      SKIP
      log_podr=.T.
   ENDDO
   SELECT rasp
   log_podr=.F.
   SKIP
ENDDO



SELECT rasp
GO TOP
PUBLIC kpdop
kpdop=0
kpdop=IIF(rasp->kp=0,1,rasp->kp)

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


SELECT rasp
fpodr=CREATEOBJECT('FORMSPR')
WITH fpodr
     .Caption='Расчет планируемых затрат на оплату труда, для лиц замещающих на курсы работников'
     DO addButtonOne WITH 'fPodr','menuCont1',10,5,'редакция','pencil.ico',"Do readspr WITH 'fpodr','Do inputzam'",39,RetTxtWidth('календарь')+44,'редакция'
     DO addButtonOne WITH 'fPodr','menuCont2',.menucont1.Left+.menucont1.Width+3,5,'удаление','pencild.ico','Do deletefromzam',39,.menucont1.Width,'удаление'   
     DO addButtonOne WITH 'fPodr','menuCont3',.menucont2.Left+.menucont2.Width+3,5,'расчёт','calculate.ico','DO procCountrash',39,.menucont1.Width,'расчёт'       
     DO addButtonOne WITH 'fPodr','menuCont4',.menucont3.Left+.menucont3.Width+3,5,'печать','print1.ico','DO printKursZam',39,.menucont1.Width,'печать' 
     DO addButtonOne WITH 'fPodr','menuCont5',.menucont4.Left+.menucont4.Width+3,5,'настройки','setup.ico','DO setupZam',39,.menucont1.Width,'настройки'  
     DO addButtonOne WITH 'fPodr','menuCont6',.menucont5.Left+.menucont5.Width+3,5,'возврат','undo.ico','DO exitFromProcKurs',39,.menucont1.Width,'возврат'       
     DO addButtonOne WITH 'fPodr','menuexit1',10,5,'возврат','undo.ico','DO exitReadPers',39,RetTxtWidth('возврат')+44,'вовзрат' 
     .menuExit1.Visible=.F.
                     
     var_path=FULLPATH('rashset.mem')
     DO addmenureadspr WITH 'fpodr','DO writezam WITH .F.','DO writezam WITH .T.'
     *DO addcontmenu WITH 'fpodr','menucont1',10,5,'ред-е дол.','read.bmp',"Do readspr WITH 'fpodr','Do inputzam'"
     *DO addcontmenu WITH 'fpodr','menucont2',.menucont1.Left+.menucont1.Width+3,5,'удаление','del.bmp','Do deletefromzam'
     *DO addcontmenu WITH 'fpodr','menucont3',.menucont2.Left+.menucont2.Width+3,5,'расчёт','count.ico','DO proccountrash'              
     *DO addcontmenu WITH 'fpodr','menucont4',.menucont3.Left+.menucont3.Width+3,5,'печать','printer.bmp','DO printkurszam'
     *DO addcontmenu WITH 'fpodr','menucont5',.menucont4.Left+.menucont4.Width+3,5,'сокр.','small.ico','DO gridVisible','перейти в полную форму'
     *DO addcontmenu WITH 'fpodr','menucont6',.menucont5.Left+.menucont5.Width+3,5,'настройки','tools.bmp','DO setupzam' 
     *DO addcontmenu WITH 'fpodr','menucont7',.menucont6.Left+.menucont6.Width+3,5,'поиск','poisk.bmp','DO poiskkurs' 
     *DO addcontmenu WITH 'fpodr','menucont8',.menucont7.Left+.menucont7.Width+3,5,'выход','exit.bmp','fpodr.Release' 
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
          .Height=(fpodr.Height-.Top-5)/2
          .Width=fpodr.Width/4*3
          .RecordSource='rasp'
           DO addColumnToGrid WITH 'fPodr.fGrid',12
          .RecordSourceType=1
*          .ColumnCount=12
          .Column1.ControlSource='rasp.nd'
          .Column2.ControlSource='" "+sprdolj.name'
          .Column3.ControlSource='rasp.kpeopk'
     
          .Column5.ControlSource='rasp.ksekurs'
          .Column6.ControlSource='rasp.dkurs'
          .Column7.ControlSource='rasp.dzamk'
          .Column8.ControlSource='rasp.pol1'
          .Column9.ControlSource='rasp.pol2'
          .Column10.ControlSource='rasp.srzpk'
          .Column11.ControlSource='rasp.zpdayk'     
          .Column1.Width=RettxtWidth(' 123 ')    
          .Column3.Width=RettxtWidth('99999')
          .Column4.Width=RettxtWidth('9999')
          .Column5.Width=RettxtWidth('9999999')
          .Column6.Width=.Column4.Width
          .Column7.Width=.Column4.Width
          .Column8.Width=RetTxtWidth(' 1 пол. ')
          .Column9.Width=RetTxtWidth(' 2 пол. ')
          .Column10.Width=RettxtWidth('99999999.99')
          .Column11.Width=.Column7.Width       
          .Columns(.ColumnCount).Width=0   
          .Column2.Width=.Width-.Column1.Width-.Column3.Width-.Column4.Width-.Column5.Width-.Column6.Width-.Column7.Width-.Column8.Width-.Column9.Width-.Column10.Width-.Column11.Width-SYSMETRIC(5)-13-.ColumnCount
          .Column1.Header1.Caption='№'
          .Column2.Header1.Caption='Должность'
          .Column3.Header1.Caption='Сотр.'
          .Column4.Header1.Caption='То'
          .Column5.Header1.Caption='Ш.ед.'
          .Column6.Header1.Caption='Дн.курс.'
          .Column7.Header1.Caption='Дн.зам.'
          .Column8.Header1.Caption='1 пол.'
          .Column9.Header1.Caption='2 пол.'
          .Column10.Header1.Caption='Ср.зп.'
          .Column11.Header1.Caption='За 1 день.'     
          .Column2.Alignment=0         
          .Column3.Format='Z'
          .Column5.Format='Z'
          .Column6.Format='Z'
          .Column7.Format='Z'
          .Column8.Format='Z'     
          .Column9.Format='Z' 
          .Column10.Format='Z'     
          .Column11.Format='Z' 
          .procAfterRowColChange='DO fpodrRefresh' 
          .Column4.AddObject('checkColumn4','checkContainer')
          .Column4.checkColumn4.AddObject('checkMy','checkBox')
          .Column4.CheckColumn4.checkMy.Visible=.T.
          .Column4.CheckColumn4.checkMy.Caption=''
          .Column4.CheckColumn4.checkMy.Left=10
          .Column4.CheckColumn4.checkMy.BackStyle=0
          .Column4.CheckColumn4.checkMy.ControlSource='rasp.lokl'                                                                                                  
          .column4.CurrentControl='checkColumn4'
          .Column4.Sparse=.F.            
          .SetAll('BOUND',.F.,'ColumnMy')       
          .SetAll('Alignment',2,'Header')  
          .colNesInf=2              
     ENDWITH
     DO gridSizeNew WITH 'fpodr','fGrid','shapeingrid' 

     .AddObject('dopGrid','gridMyNew')
     WITH .dopGrid
          .Left=0
          .Top=fpodr.fGrid.Top+fpodr.fGrid.Height+10 
          .Width=fpodr.Width
          .Height=fpodr.height-.Top
          .ScrollBars=2
        *  .ColumnCount=6
          .RecordSourceType=1
          .RecordSource='doprasp' 
           DO addColumnToGrid WITH 'fPodr.dopGrid',6
          .Column1.Header1.Caption='Подразделение'
          .Column1.ControlSource='doprasp.primpodr'
          .Column1.Alignment=0
     
          .Column2.Header1.Caption='Должность'
          .Column2.ControlSource='doprasp.pldop'
          .Column2.Alignment=0
          .Column3.Header1.Caption='Шт.ед.'
          .Column3.ControlSource='doprasp.kse'
          .Column3.Width=RettxtWidth('9999999')         
     
          .Column4.Header1.Caption='Сотр.'
          .Column4.ControlSource='doprasp.kpeopk'
          .Column4.Width=RettxtWidth('9999999')
     
          .Column5.Header1.Caption='Ср.зп'
          .Column5.ControlSource='doprasp.srzpk'
          .Column5.Width=RettxtWidth('99999999.99')
          .Column2.Width=(.Width-.Column3.Width)/2
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0   
          .procAfterRowColChange='DO fpodrRefresh'    
     ENDWITH
     DO gridsizeNew WITH 'fpodr','dopGrid','shapeingrid1'


     .AddObject('peopGrid','gridMyNew')
     WITH .peopGrid
          .Left=0
          .Top=fpodr.fGrid.Top+fpodr.fGrid.Height+10 
          .Width=fpodr.Width
          .Height=fpodr.height-.Top
          .ScrollBars=2
        *  .ColumnCount=6
          .RecordSourceType=1
          .RecordSource='doppeople' 
           DO addColumnToGrid WITH 'fPodr.peopGrid',6
          .Column1.Header1.Caption='Подразделение'
          .Column1.ControlSource='doppeople.podr'
          .Column1.Alignment=0
     
          .Column2.Header1.Caption='Должность'
          .Column2.ControlSource='doppeople.dolj'
          .Column2.Alignment=0
     
          .Column3.Header1.Caption='ФИО.'
          .Column3.ControlSource='doppeople.name'       
     
          .Column4.Header1.Caption='шт.ед.'
          .Column4.ControlSource='doppeople.kse'
          .Column4.Width=RettxtWidth('9999999')
     
          .Column5.Header1.Caption='м.фонд'
          .Column5.ControlSource='doppeople.msf'
          .Column5.Width=RettxtWidth('99999999.99')
          .Column2.Width=(.Width-.Column4.Width-.Column5.Width)/3
          .Column3.Width=.Column2.Width
          .Column1.Width=.Width-.Column2.Width-.Column3.Width-.Column4.Width-.Column5.Width-SYSMETRIC(5)-13-.ColumnCount   
          .Columns(.ColumnCount).Width=0   
          .procAfterRowColChange='DO fpodrRefresh'  
          .Visible=.F.  
     ENDWITH
     DO gridsizeNew WITH 'fpodr','peopGrid','shapeingrid2'
     .shapeingrid2.Visible=.F.


     .combobox1.DisplayCount=MIN(RECCOUNT('sprpodr'),(.Height-.combobox1.Top-.combobox1.Height)/.fGrid.Rowheight)
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
     DO adtbox WITH 'fpodr',2,1,1,.fGrid.Column5.Width+2,dHeight,.F.,'Z',.T.
     DO adtbox WITH 'fpodr',3,1,1,.fGrid.Column6.Width+2,dHeight,.F.,'Z',.T.,2,'REPLACE dzamk WITH ROUND(ksekurs*dkurs,0)'
     DO adtbox WITH 'fpodr',4,1,1,.fGrid.Column7.Width+2,dHeight,.F.,'Z',.T.,2,'DO validdaypol'
     DO adtbox WITH 'fpodr',5,1,1,.fGrid.Column8.Width+2,dHeight,.F.,'Z',.T.,2,'DO validdaypol'
     DO adtbox WITH 'fpodr',6,1,1,.fGrid.Column9.Width+2,dHeight,.F.,'Z',.T.,2,'DO validdaypol'
     DO adtbox WITH 'fpodr',7,1,1,.fGrid.Column10.Width+2,dHeight,.F.,'Z',.T.,2,'DO validsrzpk'
     DO adtbox WITH 'fpodr',8,1,1,.fGrid.Column11.Width+2,dHeight,.F.,'Z',.T.,2,'DO validdaypol'
     .SetAll('Visible',.F.,'MyTxtBox')
     objTop=.fGrid.Top
     objLeft=.fGrid.Width+5
     objWidth=ROUND((fpodr.Width-fpodr.fGrid.Width-5)/2,0)
     DO addcontform WITH 'fpodr','cont1',objLeft,objtop,objWidth,.fGrid.HeaderHeight+1,'Месяц' 
     DO addcontform WITH 'fpodr','cont2',.cont1.Left+.cont1.Width-1,objtop,objWidth,.cont1.Height,'Сумма' 
     objTop=.cont1.Top+.cont1.Height-1
     objHeight=.fGrid.RowHeight
     FOR i=1 TO 13 
         vctrl=IIF(i<13,'dim_month('+LTRIM(STR(i))+')','всего')  
         DO adtbox WITH 'fpodr',i+10,objLeft,objTop,objWidth,objHeight,vctrl,'Z',.F.,0
         objTop=objTop+objHeight-1
     ENDFOR 

     objleft1=objLeft
     objWidth1=objWidth
     objTop=.txtbox11.Top
     objLeft=.Txtbox11.Left+.Txtbox11.Width-1
     objWidth=.cont2.Width

     FOR i=1 TO 13 
         vctrl=IIF(i<13,'rasp->zk'+LTRIM(STR(i)),'rasp->totzpk')     
         repzam=IIF(i=13,'itdayk','zk'+LTRIM(STR(i)))   
         DO adtbox WITH 'fpodr',i+30,objLeft,objTop,objWidth,objHeight,vctrl,'Z',.F.,.F.
         obj_cx='fpodr.txtbox'+LTRIM(STR(i+30))
         &obj_cx..procForLostFocus='DO sumzptotk'
     *    &obj_cx..procForValid='REPLACE &repzam WITH &vctrl*zpdayk'
         objTop=objTop+objHeight-1   
     ENDFOR 

     fpodr.txtbox43.Forecolor=IIF(rasp->itdayk#rasp->dzamk,objcolorsos,dForeColor) 
     ord_ch=44
     limTop=fpodr.fGrid.Top+fpodr.fGrid.Height
     DO WHILE .T.
        IF objTop>limTop     
           EXIT
        ENDIF
        obj_cont='lCont'+LTRIM(STR(ord_ch))
        fpodr.AddObject(obj_cont,'Container')
        WITH fpodr.&obj_cont
             .Height=objHeight
             .Width=objWidth1
             .Top=objTop
             .Left=objLeft1
             .BackStyle=0                   
             .Visible=.T.
        ENDWITH    
        ord_ch=ord_ch+1
        obj_cont='lCont'+LTRIM(STR(ord_ch))
        fpodr.AddObject(obj_cont,'Container')
        WITH fpodr.&obj_cont
             .Height=objHeight
             .Width=objwidth
             .Top=objTop
             .Left=objLeft
             .BackStyle=0                   
             .Visible=.T.
        ENDWITH       
        ord_ch=ord_ch+1
        objTop=objTop+objHeight-1     
     ENDDO 
ENDWITH
fpodr.Show
********************************************************************************************************************************************************
PROCEDURE exitFromProcKurs
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
fpodr.Refresh
SELECT rasp
********************************************************************************************************************************************************
PROCEDURE gridVisible
log_form=IIF(log_form=1,2,1)
 WITH fpodr.menucont5
        .contlabel.Caption=IIF(log_form=1,'сокр.','полн.')
        .contimage.Picture=IIF(log_form=1,'small.ico','full.ico')
        .contimage.ToolTipText=IIF(log_form=1,'перейти в полную форму','перейти в сокращённую форму')
        .contlabel.ToolTipText=IIF(log_form=1,'перейти в полную форму','перейти в сокращённую форму')
        .Refresh
   ENDWITH
IF log_form=1  
   SELECT doprasp
   GO TOP
*   fpodr.menucont4.procForClick='DO printkurszam WITH 1'
   fpodr.peopGrid.Visible=.F.
   fpodr.Shapeingrid2.Visible=.F.
   fpodr.dopGrid.Visible=.T.
   fpodr.Shapeingrid1.Visible=.T.   
ELSE  
   SELECT doppeople
   GO TOP
*   fpodr.menucont4.procForClick='DO printkurszam WITH 2'
   fpodr.dopGrid.Visible=.F.
   fpodr.Shapeingrid1.Visible=.F.
   fpodr.peopGrid.Visible=.T.
   fpodr.Shapeingrid2.Visible=.T.  
ENDIF
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
*************************************************************************************************************************
PROCEDURE procCheckOkl
SELECT datjob
SEEK STR(rasp->kp,3)+STR(rasp->kd,3)
STORE 0 TO nkse,nms,srms,npeop
DO WHILE kp=rasp->kp.AND.kd=rasp->kd
   npeop=npeop+1
   nkse=nkse+datJob.kse  
   IF !rasp.lOkl
      nms=nms+&formulaOpt
   ELSE 
      nms=nms+(datjob.tOkl*datjob.kse)
   ENDIF    
   SKIP 
   ENDDO
SELECT rasp 
REPLACE ksekurs WITH IIF(ksekurs=0,rasp.kse,ksekurs),srzpk WITH IIF(nkse<1,nms,nms/nkse),zpdayk WITH srzpk/rashset(8),kpeopk WITH npeop   
KEYBOARD '{TAB}'
fPodr.Refresh

*********************************************************************************************************************************************************
*                                     Редактирование информации по замене (должность)
*********************************************************************************************************************************************************
PROCEDURE inputzam
SELECT rasp
SCATTER TO fpodr.dim_ap  
SELECT datjob
SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
STORE 0 TO nkse,nms,srms,npeop
DO WHILE kp=rasp.kp.AND.kd=rasp.kd
   npeop=npeop+1
   nkse=nkse+datjob.kse  
   IF !rasp.lOkl
      nms=nms+&formulaOtp
   ELSE 
      nms=nms+(datjob.tokl*datjob.kse)
   ENDIF  
   SKIP 
ENDDO
SELECT rasp 
REPLACE ksekurs WITH IIF(ksekurs=0,rasp.kse,ksekurs),srzpk WITH IIF(nkse<1,nms,nms/nkse),zpdayk WITH srzpk/rashset(8),kpeopk WITH npeop   
WITH fPodr
     .SetAll('Visible',.F.,'mymenucont')
     .fGrid.Column1.SetFocus
     .menuread.Visible=.T.
     .menuexit.Visible=.T.
     .fGrid.Enabled=.F.
     .combobox1.Enabled=.F.
     .combobox1.Style=0
     .CheckOkl.Visible=.T.
     lineTop=.fGrid.Top+.fGrid.HeaderHeight+.fGrid.RowHeight*(IIF(.fGrid.RelativeRow<=0,1,.fGrid.RelativeRow)-1)
     .CheckOkl.Top=lineTop
     .txtbox2.Top=linetop
     .txtbox3.Top=linetop
     .txtbox4.Top=linetop
     .txtbox5.Top=linetop
     .txtbox6.Top=linetop
     .txtbox7.Top=linetop
     .txtbox8.Top=lineTop
     .CheckOkl.Height=.fGrid.RowHeight+1
     .txtbox2.Height=.fGrid.RowHeight+1
     .txtbox3.Height=.fGrid.RowHeight+1
     .txtbox4.Height=.fGrid.RowHeight+1
     .txtbox5.Height=.fGrid.RowHeight+1
     .txtbox6.Height=.fGrid.RowHeight+1 
     .txtbox7.Height=.fGrid.RowHeight+1
     .txtbox8.Height=.fGrid.RowHeight+1
      .checkOkl.Left=.fGrid.Left+13+.fGrid.Column1.Width+.fgrid.Column2.Width+.fgrid.Column3.Width
     .txtBox2.Left=.checkOkl.Left+.checkOkl.Width-1
     .txtbox3.Left=fpodr.txtbox2.Left+fpodr.txtbox2.Width-1
     .txtbox4.Left=fpodr.txtbox3.Left+fpodr.txtbox3.Width-1
     .txtbox5.Left=fpodr.txtbox4.Left+fpodr.txtbox4.Width-1 
     .txtbox6.Left=fpodr.txtbox5.Left+fpodr.txtbox5.Width-1
     .txtbox7.Left=fpodr.txtbox6.Left+fpodr.txtbox6.Width-1
     .txtbox8.Left=fpodr.txtbox7.Left+fpodr.txtbox7.Width-1
     .txtbox2.ControlSource='rasp->ksekurs'
     .txtbox3.ControlSource='rasp->dkurs'
     .txtbox4.ControlSource='rasp->dzamk'
     .txtbox5.ControlSource='rasp->pol1' 
     .txtbox6.ControlSource='rasp->pol2'
     .txtbox7.ControlSource='rasp->srzpk'
     .txtbox8.ControlSource='rasp->zpdayk'
     .checkOkl.BackStyle=1
     .txtbox2.BackStyle=1
     .txtbox3.BackStyle=1
     .txtbox4.BackStyle=1
     .txtbox5.BackStyle=1
     .txtbox6.BackStyle=1
     .txtbox7.BackStyle=1
     .txtbox8.BackStyle=1
     FOR i=31 TO 42
         objen='fpodr.txtbox'+LTRIM(STR(i))
         &objen..Enabled=.T.
         &objen..BackStyle=1  
     ENDFOR
     .txtbox31.Enabled=.T.
     .txtbox32.Enabled=.T.
     .SetAll('Visible',.T.,'MyTxtBox')
     .Refresh
     .Txtbox2.SetFocus
ENDWITH 
SELECT rasp
****************************************************************************************************************************************************
PROCEDURE validsrzpk
REPLACE zpdayk WITH srzpk/rashset(8)
DO validdaypol
****************************************************************************************************************************************************
*                      Расчёт сумм по месяцам
****************************************************************************************************************************************************
PROCEDURE validdaypol
SELECT rasp
sum1pol=zpdayk*pol1/6
sum2pol=zpdayk*pol2/6
REPLACE zk1 WITH sum1pol,zk2 WITH sum1pol,zk3 WITH sum1pol,zk4 WITH sum1pol,zk5 WITH sum1pol,zk6 WITH sum1pol
REPLACE zk7 WITH sum2pol,zk8 WITH sum2pol,zk9 WITH sum2pol,zk10 WITH sum2pol,zk11 WITH sum2pol,zk12 WITH sum2pol
REPLACE totzpk WITH zk1+zk2+zk3+zk4+zk5+zk6+zk7+zk8+zk9+zk10+zk11+zk12

tot_cx=0
FOR  h=1 TO 12
      rep_cx='zk'+LTRIM(STR(h))
*      mon_cx='mk'+LTRIM(STR(h))
      IF h<MONTH(rashset(7))
         REPLACE &rep_cx WITH 0           
      ELSE
         tot_cx=tot_cx+ EVALUATE('zk'+LTRIM(STR(h)))  
      ENDIF        
ENDFOR 
REPLACE totzpk WITH tot_cx
fpodr.Refresh
*****************************************************************************************************************************************************
*                               Запись информации по замене
**************************************************************************************************************************
PROCEDURE writezam
PARAMETERS parlog
WITH fPodr 
     .SetAll('Visible',.T.,'mymenucont')
     .menuread.Visible=.F.
     .menuexit.Visible=.F.
     SELECT rasp
     IF parlog
        GATHER FROM .dim_ap
     ENDIF
     .checkOkl.Visible=.F.
     .txtbox2.Visible=.F.
     .txtbox3.Visible=.F.
     .txtbox4.Visible=.F.
     .txtbox5.Visible=.F.
     .txtbox6.Visible=.F.
     .txtbox7.Visible=.F.
     .txtbox8.Visible=.F.
     .combobox1.Enabled=.T.
     .combobox1.Style=2
     FOR i=31 TO 43
         objtxt='txtbox'+LTRIM(STR(i))
         fpodr.&objtxt..Enabled=.F. 
         fpodr.&objtxt..BackStyle=0        
     ENDFOR 
     .fGrid.Enabled=.T.
     .fGrid.SetAll('Enabled',.F.,'Column')
     .fGrid.Columns(.fGrid.ColumnCount).Enabled=.T.
ENDWITH 
IF !parlog.AND.rasp->dzamk#rasp->itdayk
   * DO createForm WITH .T.,'Внимание!',RetTxtWidth('WWНесовпадение дней замены!WW',dFontName,dFontSize+1),'130',;
   *  RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,'Несовпадение дней замены!' 
ENDIF
*****************************************************************************************************************************************************
PROCEDURE sumzptotk
SELECT rasp
STORE 0 TO tot_cx,day_cx
FOR h=1 TO 12
    tot_cx=tot_cx+EVALUATE('zk'+LTRIM(STR(h)))   
ENDFOR 
REPLACE totzpk WITH tot_cx
fpodr.Refresh
*****************************************************************************************************************************************************
PROCEDURE sumzpkdoltot
SELECT datjob
STORE 0 TO tot_cx,day_cx,dayst_cx
FOR h=1 TO 12
    tot_cx=tot_cx+EVALUATE('zpk'+LTRIM(STR(h)))
    day_cx=day_cx+EVALUATE('dk'+LTRIM(STR(h)))
    dayst_cx=dayst_cx+EVALUATE('dstk'+LTRIM(STR(h)))
    
ENDFOR 
REPLACE zptotk WITH tot_cx, dktot WITH day_cx,dsttotk WITH dayst_cx
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
    sumdim=IIF(i<13,'dstk'+LTRIM(STR(i)),'dsttotk')
    repdim='dim_dayst'+LTRIM(STR(i))
    sumzp=IIF(i<13,'zpk'+LTRIM(STR(i)),'zptotk')
    repzp='dim_zp'+LTRIM(STR(i))
    SUM &sumdim,&sumzp TO &repdim,&repzp
    SELECT rasp
    repm=IIF(i<13,'mk'+LTRIM(STR(i)),'itdayk')
    repz=IIF(i<13,'zk'+LTRIM(STR(i)),'totzpk')
    REPLACE &repm WITH &repdim,&repz WITH &repzp
    obj_cx='fpodr.txtbox'+LTRIM(STR(i+30))
    &obj_cx..Refresh
    obj_ch='fpodr.txtbox'+LTRIM(STR(i+50))
    IF i=13
       &obj_cx..DisabledForecolor=IIF(rasp->itdayk#rasp->dzamk,objcolorsos,dForeColor)
    ENDIF   
    &obj_ch..Refresh
    SELECT datjob
ENDFOR
SELECT datjob
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
oldkse=ksekurs
oldkat=kat
SELECT rasp
DO CASE
   CASE dim_del(1)=1
        SELECT datjob
        FOR i=1 TO 12
            repd='dk'+LTRIM(STR(i))
            repdst='dstk'+LTRIM(STR(i))
            repzp='zpk'+LTRIM(STR(i))
            REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
        ENDFOR                      
        replace zptotk WITH 0,dktot WITH 0,dsttotk WITH 0
        SELECT rasp
        FOR i=1 TO 12
            repm='mk'+LTRIM(STR(i))
            repz='zk'+LTRIM(STR(i))
            REPLACE &repm WITH 0,&repz WITH 0
        ENDFOR                      
        replace kpeopk WITH 0,dkurs WITH 0,dzamk WITH 0,itdayk WITH 0,srzpk WITH 0,zpdayk WITH 0,totzpk WITH 0
   CASE dim_del(2)=1
        SELECT datjob 
        SET FILTER TO kp=kpdop
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12
               repd='dk'+LTRIM(STR(i))
               repdst='dstk'+LTRIM(STR(i))
               repzp='zpk'+LTRIM(STR(i))
               REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
           ENDFOR                      
           replace zptotk WITH 0,dktot WITH 0,dsttotk WITH 0
           SKIP
        ENDDO         
        SELECT rasp        
        GO TOP
        DO WHILE !EOF()
           SELECT rasp
           FOR i=1 TO 12
               repm='mk'+LTRIM(STR(i))
               repz='zk'+LTRIM(STR(i))
               REPLACE &repm WITH 0,&repz WITH 0
           ENDFOR                      
           replace kpeopk WITH 0,dkurs WITH 0,dzamk WITH 0,itdayk WITH 0,srzpk WITH 0,zpdayk WITH 0,totzpk WITH 0
           SKIP
        ENDDO   
        SET FILTER TO kp=kpdop
        GO top         
   CASE dim_del(3)=1
        SELECT datjob 
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()
           FOR i=1 TO 12
               repd='dk'+LTRIM(STR(i))
               repdst='dstk'+LTRIM(STR(i))
               repzp='zpk'+LTRIM(STR(i))
               REPLACE &repd WITH 0,&repdst WITH 0,&repzp WITH 0
           ENDFOR                      
           replace zptotk WITH 0,dktot WITH 0,dsttotk WITH 0
           SKIP
        ENDDO           
        SELECT rasp
        SET FILTER TO 
        GO TOP
        DO WHILE !EOF()
           SELECT rasp
           FOR i=1 TO 12
               repm='mk'+LTRIM(STR(i))
               repz='zk'+LTRIM(STR(i))
               REPLACE &repm WITH 0,&repz WITH 0
           ENDFOR                      
           replace kpeopk WITH 0,dkurs WITH 0,dzamk WITH 0,itdayk WITH 0,srzpk WITH 0,zpdayk WITH 0,totzpk WITH 0
           SKIP
        ENDDO   
        SET FILTER TO kp=kpdop
        GO top                  
ENDCASE
fpodr.Refresh
fdel.Release
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
log_srzp=.F.
WITH fdel
     .Caption='Расчёт расходов'
     .BackColor=RGB(255,255,255)
     .AddObject('Shape1','ShapeMy')
     .Shape1.Top=10
     .Shape1.Left=10
     .Shape1.Curvature=8
     .Shape1.BorderColor=RGB(192,192,192)
     DO adLabMy WITH 'fdel',1,'Дата отсчёта',fdel.Shape1.Top+10,fdel.Shape1.Left+15,150,0,.T.
     DO adtbox WITH 'fdel',1,fdel.lab1.Left+fdel.lab1.Width+10,fdel.Shape1.Top+10,RetTxtWidth('99/99/99999'),dHeight,'rashset(7)','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashset'
     fdel.lab1.Top=fdel.txtbox1.Top+(fdel.txtbox1.Height-fdel.lab1.Height)
     DO addOptionButton WITH 'fdel',1,'расчет по выбранной должности',fdel.txtbox1.Top+fdel.txtbox1.Height+10,fdel.Shape1.Left+15,'dim_del(1)',0,"DO storedimdel WITH 1",.T.
     DO addOptionButton WITH 'fdel',2,'расчёт по подразделению',fdel.Option1.Top+fdel.Option1.Height+10,fdel.Option1.Left,'dim_del(2)',0,"DO storedimdel WITH 2",.T.
     DO addOptionButton WITH 'fdel',3,'расчёт по организации',fdel.Option2.Top+fdel.Option2.Height+10,fdel.Option1.Left,'dim_del(3)',0,"DO storedimdel WITH 3",.T.
     fdel.Shape1.Height=fdel.Option1.height*4+60
     fdel.Shape1.Width=fdel.Option1.Width+30
     DO adCheckBox WITH 'fdel','check1','пересчитать среднюю зарплату',fdel.Shape1.Top+fdel.Shape1.Height+10,fdel.Shape1.Left,150,dHeight,'log_srzp',0
     DO adCheckBox WITH 'fdel','check2','подтверждение выполнения',fdel.check1.Top+fdel.check1.Height+10,fdel.Shape1.Left,150,dHeight,'log_del',0  
     DO addcontlabel WITH 'fdel','cont1',fdel.Shape1.Left+5,fdel.check2.Top+fdel.check2.Height+15,;
        (fdel.shape1.Width-20)/2,dHeight+3,'Выполнение','DO countrash'
     DO addcontlabel WITH 'fdel','cont2',fdel.Cont1.Left+fdel.Cont1.Width+10,fdel.Cont1.Top,;
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
     .Height=.Shape1.Height+fpodr.cont1.Height+fdel.check1.Height*2+70
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
IF srzpk#0
   SELECT datJob
   SEEK STR(rasp.kp,3)+STR(rasp.kd,3)
   STORE 0 TO nkse,nms,srms,npeop
   DO WHILE kp=rasp.kp.AND.kd=rasp.kd
      npeop=npeop+1
      nkse=nkse+datJob.kse
      IF !rasp.lOkl
         nms=nms+&formulaOtp
      ELSE 
         nms=nms+(datJob.tokl*datjob.kse)
      ENDIF       
      SKIP 
   ENDDO
   SELECT rasp 
   IF log_srzp
      REPLACE kpeopk WITH npeop,srzpk WITH IIF(nkse<1,nms,nms/nkse),zpdayk WITH srzpk/rashset(8)    
   ENDIF   
*   SELECT people
*   SEEK STR(rasp->kp,3)+STR(rasp->kd,3)
 *  DO WHILE kp=rasp->kp.AND.kd=rasp->kd
 *     STORE 0 TO tot_cx,day_cx,dayst_cx                
  *   FOR h=1 TO 12
   *       IF h>=MONTH(rashset(7))                
   *          tot_cx=tot_cx+EVALUATE('zpk'+LTRIM(STR(h)))
   *          day_cx=day_cx+EVALUATE('dk'+LTRIM(STR(h)))
   *          dayst_cx=dayst_cx+EVALUATE('dstk'+LTRIM(STR(h)))    
   *          repdst='dstk'+LTRIM(STR(h))
   *          repzp='zpk'+LTRIM(STR(h))                               
   *       ELSE
   *          repzp='zpk'+LTRIM(STR(h))
   *          repd='dk'+LTRIM(STR(h))
   *          repdst='dstk'+LTRIM(STR(h))
   *          REPLACE &repzp WITH 0,&repd WITH 0,&repdst WITH 0
   *       ENDIF   
   *   ENDFOR 
   *   REPLACE zptotk WITH tot_cx, dktot WITH day_cx,dsttotk WITH dayst_cx
  *    SKIP
  * ENDDO             
 *  SELECT rasp
   STORE 0 TO tot_cx,day_cx
   
   
   sum1pol=zpdayk*pol1/6
   sum2pol=zpdayk*pol2/6
   REPLACE zk1 WITH sum1pol,zk2 WITH sum1pol,zk3 WITH sum1pol,zk4 WITH sum1pol,zk5 WITH sum1pol,zk6 WITH sum1pol
   REPLACE zk7 WITH sum2pol,zk8 WITH sum2pol,zk9 WITH sum2pol,zk10 WITH sum2pol,zk11 WITH sum2pol,zk12 WITH sum2pol
   REPLACE totzpk WITH zk1+zk2+zk3+zk4+zk5+zk6+zk7+zk8+zk9+zk10+zk11+zk12         
   FOR  h=1 TO 12
        rep_cx='zk'+LTRIM(STR(h))
*        mon_cx='mk'+LTRIM(STR(h))
        IF h<MONTH(rashset(7))
           REPLACE &rep_cx WITH 0           
        ELSE
           tot_cx=tot_cx+ EVALUATE('zk'+LTRIM(STR(h)))  
        ENDIF        
   ENDFOR 
   REPLACE totzpk WITH tot_cx
   
   
   
   *FOR h=1 TO 12
   *    rep_cx='zk'+LTRIM(STR(h))
   *    mon_cx='mk'+LTRIM(STR(h))
   *    IF h>=MONTH(rashset(7))                       
   *       REPLACE &rep_cx WITH EVALUATE('mk'+LTRIM(STR(h)))*zpdayk
   *       tot_cx=tot_cx+EVALUATE('zk'+LTRIM(STR(h)))
   *       day_cx=day_cx+EVALUATE('mk'+LTRIM(STR(h)))
   *    ELSE
   *       REPLACE &rep_cx WITH 0,&mon_cx WITH 0    
   *    ENDIF   
   *ENDFOR 
   *REPLACE totzpk WITH tot_cx, itdayk WITH day_cx,dzamk WITH day_cx     
   *IF itdayk=0
   *   REPLACE srzpk WITH 0, zpdayk WITH 0, dzamk WITH 0
   *ENDIF  
ENDIF
*****************************************************************************************************************************************************
*                                  Печать ведомости расчёта расходов по замене курсов
*****************************************************************************************************************************************************
PROCEDURE printkurszam
*PARAMETERS par_log
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
dimKurs(1,1)=''
dimKurs(1,2)=0
SELECT * FROM sprkat INTO CURSOR curKatKurs READWRITE
ALTER TABLE curKatKurs ADD COLUMN sumtot N (10,2)
SELECT curKatKurs
INDEX ON kod TAG T1

fpodr.fGrid.Column8.SetFocus
STORE 0 TO numrecrep
SELECT rasp
SET FILTER TO totzpk#0
SCAN ALL
     IF SEEK(rasp.kat,'curKatKurs',1)
        REPLACE curKatKurs.sumtot WITH curKatKurs.sumtot+rasp.totzpk 
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
DO CASE
   CASE log_form=2 
        DO printreport WITH 'repzamkurs','Расчёт расходов по замене курсов','rasp'
   CASE log_form=1    
        DO printreport WITH 'repzamkursnew','Расчёт расходов по замене курсов','rasp' 
ENDCASE        
SELECT rasp
SET FILTER TO kp=kpdop
*****************************************************************************************************************************************************
*                                  Печать ведомости расчёта расходов по замене курсов
*****************************************************************************************************************************************************
PROCEDURE 1printkurszam
fpodr.fGrid.Column8.SetFocus
STORE 0 TO numrecrep
SELECT rasp
IF rashset(6)
   SET FILTER TO totzpk#0
ELSE
   SET FILTER TO       
ENDIF  

=AFIELDS(arRasp,'rasp') 
CREATE CURSOR currasp FROM ARRAY arRasp
SELECT currasp
APPEND FROM rasp
INDEX ON STR(np,3)+STR(nd,3) TAG t1
INDEX ON STR(kp,3)+STR(kd,3) TAG t2
SET ORDER TO 1  
IF rashset(6)
    SET FILTER TO totzpk>0.AND.kp>0.AND.nd>0     
ELSE 
    SET FILTER TO 
ENDIF    
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
SELECT datjob
SELECT * FROM datjob WHERE SEEK(STR(datjob.kp,3)+STR(datjob.kd,3),'currasp',2).AND.datjob.zptotk>0 INTO CURSOR curpeople READWRITE  
SELECT curpeople
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
SET ORDER TO
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curpeople->kp,3)+STR(curpeople->kd,3)   
   SELECT curpeople
   REPLACE nd WITH currasp->nd,np WITH currasp->np     
   SKIP 
ENDDO 
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
SELECT currasp
SET RELATION TO STR(np,3)+STR(nd,3) INTO curpeople
SET SKIP TO curpeople
SELECT currasp
GO TOP
DO printreport WITH 'repzamkurs','Расчёт расходов по замене курсов','rasp'
SELECT rasp
SET FILTER TO kp=kpdop
*****************************************************************************************************************************************************
*                                  Печать ведомости расчёта расходов по замене курсов
*****************************************************************************************************************************************************
PROCEDURE printkurszamtot
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
SET FILTER TO totzpk>0.AND.kp>0.AND.nd>0     
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
SELECT datjob
*SELECT * FROM people WHERE SEEK(STR(people->kp,3)+STR(people->kd,3),'currasp',2).AND.people->zptotk>0 INTO CURSOR curpeople READWRITE  
SELECT * FROM datjob WHERE SEEK(STR(datjob.kp,3)+STR(datjob.kd,3),'currasp',2)INTO CURSOR curpeople READWRITE  
SELECT curpeople
REPLACE fio WITH IIF(SEEK(kodpeop,'people',1),people.fio,'') ALL
SET ORDER TO
GO TOP
DO WHILE !EOF()
   SELECT currasp  
   SEEK STR(curpeople->kp,3)+STR(curpeople->kd,3)   
   SELECT curpeople
   REPLACE nd WITH currasp->nd,np WITH currasp->np     
   SKIP 
ENDDO 
INDEX ON STR(np,3)+STR(nd,3) TAG T1
SET ORDER TO 1
SELECT curpeople
SET ORDER TO 1
GO TOP 
DO printreport WITH 'repzamtotkurs','Расчёт расходов по замене отпусков','curpeople'
SELECT rasp
SET FILTER TO kp=kpdop
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
     .AddObject('Shape1','ShapeMy')
     .Shape1.Top=10
     .Shape1.Left=10
     .Shape1.Curvature=8
     .Shape1.BorderColor=RGB(192,192,192)
ENDWITH
DO adLabMy WITH 'fsetup',1,'Среднее кол-во дней',fsetup.Shape1.Top+10,fsetup.Shape1.Left+10,150,0,.T.
DO adtbox WITH 'fsetup',1,fsetup.Lab1.Left+fsetup.lab1.Width+5,fsetup.lab1.Top,150,dHeight,'rashset(8)','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashset'
fsetup.txtbox1.InputMask='99.99'
fsetup.lab1.Top=fsetup.txtbox1.Top+(fsetup.txtbox1.Height-fsetup.lab1.Height)
DO adLabMy WITH 'fsetup',2,'Дата отсчёта',fsetup.lab1.Top+fsetup.lab1.Height,fsetup.lab1.Left,150,0,.T.
DO adtbox WITH 'fsetup',2,fsetup.txtbox1.Left,fsetup.txtbox1.Top+fsetup.txtbox1.Height+10,fsetup.txtbox1.Width,dHeight,'rashset(7)','Z',.T.,1,'SAVE TO &var_path ALL LIKE rashset'
fsetup.lab2.Top=fsetup.txtbox2.Top+(fsetup.txtbox2.Height-fsetup.lab2.Height)

fsetup.Shape1.Height=dHeight*2+30 
fsetup.Shape1.Width=fsetup.Lab1.Width+fsetup.txtbox1.Width+30  
WITH fsetup
     .Caption='настройки'
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+30
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fsetup'
fsetup.Show
**************************************************************************************************************************
*                                           Процедура поиска сотрудника
**************************************************************************************************************************
PROCEDURE poiskkurs
IF log_form=1
   DO createForm WITH .T.,'Поиск сотрудника',RetTxtWidth('WПоиск возможен только при полной форме!W',dFontName,dFontSize+1),'130',;
   RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,'Поиск возможен только при полной форме!'  
   RETURN 
ENDIF 
peoprec=0
fpoisk=CREATEOBJECT('FORMMY')
WITH fpoisk
     .BackColor=RGB(255,255,255)
     .AddObject('Shape1','ShapeMy')
     .Shape1.Top=10
     .Shape1.Left=10
     .Shape1.Curvature=8
     .Shape1.BorderColor=RGB(192,192,192)
ENDWITH
find_ch=''
DO adLabMy WITH 'fpoisk',1,'ФИО сотрудника',fpoisk.Shape1.Top+10,fpoisk.Shape1.Left+10,250,2
DO addtxtboxmy WITH 'fpoisk',1,fpoisk.Shape1.Left+10,fpoisk.Shape1.Top+fpoisk.lab1.Height+10,250,.F.,'find_ch'
fpoisk.txtBox1.procForkeyPress='DO keypressfind '
WITH fpoisk.Shape1     
     .Width=fpoisk.TxtBox1.Width+20
     .Height=fpoisk.TxtBox1.Height+fpoisk.lab1.Height+22
ENDWITH

DO addcontlabel WITH 'fpoisk','cont1',fpoisk.Shape1.Left+5,fpoisk.Shape1.Top+fpoisk.Shape1.Height+5,;
   (fpoisk.shape1.Width-20)/2,dHeight+3,'Поиск','DO poiskkurspeop'
DO addcontlabel WITH 'fpoisk','cont2',fpoisk.Cont1.Left+fpoisk.Cont1.Width+10,fpoisk.Cont1.Top,;
   fpoisk.Cont1.Width,dHeight+3,'Отмена','Fpoisk.Release'
DO addcontlabel WITH 'fpoisk','cont3',fpoisk.Shape1.Left+5,fpoisk.Shape1.Top+fpoisk.Shape1.Height+5,;
  (fpoisk.shape1.Width-20)/2,dHeight+3,'Продолжить','DO restpoiskkurspeop'
fpoisk.cont3.Visible=.F.  
  
WITH fpoisk
     .Caption='Поиск'
     .MinButton=.F.
     .MaxButton=.F.
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+30+fpoisk.lab1.Height
     .WindowState=0
     .AlwaysOnTop=.T.
     .AutoCenter=.T.
ENDWITH
DO pasteImage WITH 'fpoisk'
fpoisk.Show

************************************************************************************************************************
*                                  Непосредственно поиск сотрудника
************************************************************************************************************************
PROCEDURE poiskkurspeop
IF EMPTY(find_ch)   
   RETURN
ENDIF
find_ch=ALLTRIM(find_ch)
SELECT doppeople
log_ord=SYS(21)
LOCATE FOR LOWER(find_ch)$LOWER(name)
IF FOUND()    
   peoprec=RECNO()
   fpodr.peopGrid.Refresh
   fpodr.peopGrid.Columns(fpodr.peopGrid.columnCount).SetFocus
   fpoisk.TxtBox1.SetFocus  
   fpoisk.cont1.Visible=.F.
   fpoisk.cont3.Visible=.T.              
   fpoisk.txtBox1.procForkeyPress='DO restpoiskkurspeop' 
ELSE        
   fpoisk.Visible=.F.  
   DO createForm WITH .T.,'Поиск сотрудника',RetTxtWidth('WWЗапись не найдена!WW',dFontName,dFontSize+1),'130',;
   RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,'Запись не найдена!'   
   fpoisk.Release
ENDIF
************************************************************************************************************************
*          Непосредственно поиск клиента (продолжение)
************************************************************************************************************************
PROCEDURE restpoiskkurspeop
SELECT doppeople
GO peoprec
SKIP
LOCATE REST FOR LOWER(ALLTRIM(find_ch))$LOWER(name)
IF FOUND()
   SELECT doppeople
   peoprec=RECNO() 
   fpodr.peopGrid.Refresh
   fpodr.peopGrid.Columns(fpodr.peopGrid.ColumnCount).SetFocus
   fpoisk.txtBox1.SetFocus  
ELSE  
   SELECT doppeople
   GO peoprec         
   fpoisk.Release     
ENDIF
************************************************************************************************************************
PROCEDURE keyPressFind
IF LASTKEY()=13
   find_ch=fpoisk.TxtBox1.Value    
   DO poiskkurspeop  
ENDIF   
