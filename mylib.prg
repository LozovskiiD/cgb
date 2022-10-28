PARAMETERS name_proc,par_ico
SET STATUS OFF
SET STATUS BAR OFF
SET TALK OFF
SET DATE TO GERMAN
SET HELP OFF
SET ESCAPE OFF
SET CENTURY OFF
SET SAFETY OFF
SET SCOREBOARD OFF
SET MULTILOCKS ON
SET COMPATIBLE OFF
SET SYSMENU OFF
SET DELETED ON 
IF FILE('office.mem')
   RESTORE FROM office ADDITIVE
ENDIF  
RESTORE FROM fontdef ADDITIVE
RESTORE FROM pictset ADDITIVE
RESTORE FROM strmem ADDITIVE

DIMENSION dim_month(12)
dim_month(1)='январь'
dim_month(2)='февраль'
dim_month(3)='март'
dim_month(4)='апрель'
dim_month(5)='май'
dim_month(6)='июнь'
dim_month(7)='июль'
dim_month(8)='август'
dim_month(9)='сентябрь'
dim_month(10)='октябрь'
dim_month(11)='ноябрь'
dim_month(12)='декабрь'

DIMENSION month_prn(12)
month_prn(1)='января'
month_prn(2)='февраля'
month_prn(3)='марта'
month_prn(4)='апреля'
month_prn(5)='мая'
month_prn(6)='июня'
month_prn(7)='июля'
month_prn(8)='августа'
month_prn(9)='сентября'
month_prn(10)='октября'
month_prn(11)='ноября'
month_prn(12)='декабря'

DIMENSION month_prnd(12)
month_prnd(1)='январе'
month_prnd(2)='феврале'
month_prnd(3)='марте'
month_prnd(4)='апреле'
month_prnd(5)='мае'
month_prnd(6)='июне'
month_prnd(7)='июле'
month_prnd(8)='августе'
month_prnd(9)='сентябре'
month_prnd(10)='октябре'
month_prnd(11)='ноябре'
month_prnd(12)='декабре'

STORE 0 TO max_rec,one_pers,pers_ch

DIMENSION dimCht(3),dim_prn(1)
STORE 0 TO dimCht,dim_prn
page_beg=1
page_end=999
kvo_page=1 
logWord=.F.
cFormIco=IIF(!EMPTY(par_ico),par_ico,'')

PUBLIC nFormMes
dFontName=fontdef(1)
dfontSize=fontdef(2)
selBackColor=RGB(49,106,197)
selForeColor=RGB(255,255,255)
dHeight=dFontsize+IIF(SYSMETRIC(21)=800,10,12)
*dRowHeight=IIF(dFontSize=12,dHeight+1,dHeight)
dHeight=IIF(dFontSize=12,dHeight+1,dheight)
dRowHeight=dHeight
SELECT 0
USE colorset
LOCATE FOR log_scheme
dForeColor=EVALUATE(tForeColor)
dBackColor=EVALUATE(fbackColor)
headerForeColor=EVALUATE(hforecolor)
headerBackColor=EVALUATE(hbackcolor)
dynForeColor=EVALUATE(dynamFore)
dynBackColor=EVALUATE(dynamBack)
ObjBackColor=EVALUATE(tBackColor)
ObjForeColor=EVALUATE(tForeColor)
objColorSos=RGB(255,0,0)
selBackcolor=IIF(!EMPTY(selBack),EVALUATE(selBack),selBackColor)
selForeColor=IIF(!EMPTY(selFore),EVALUATE(selFore),selForeColor)
USE 
=APRINTERS(name_prn)
=APRINTERS(prndop)
FOR i=1 TO ALEN(prndop)
    prndop(i)=UPPER(prndop(i))
ENDFOR
nameprint=''
nameprint=name_prn(ASCAN(prndop,SET('PRINTER',2)))
DO &name_proc
QUIT
**************************************************************************************************************************
PROCEDURE setuppictset
STORE .F. TO pictset
pictset=BAR()
var_path=FULLPATH('pictset.mem')
SAVE TO &var_path ALL LIKE pictset
********************************************************************************************
*           Класс для создания набора форм
********************************************************************************************
DEFINE CLASS myfset AS FORMSET
       *WindowType=0
       *WindowState=2
ENDDEFINE
*********************************************************************************************
*               Класс для создания небольших (вспомогательных) форм
* Задаются основные элементы и параметры, всё остальное потом
*********************************************************************************************
DEFINE CLASS formsupl AS FORM
       DeskTop=.T.
       AlwaysOnTop=.T.
       Top=0
       Left=0   
       Icon=cFormIco  
       Height=SYSMETRIC(22)
       Width=SYSMETRIC(21) && 1024 
       WindowState=0          
       WindowType=1      
       ShowWindow=1                  
       DoCreate = .T.
       BorderStyle = 2	
       Autocenter = .T.   
       Minbutton=.F. 
       MaxButton=.F.
       Name = "Form1"
       FontName=dFontName
       FontSize=dFontSize 
       ShowTips=.T.           
       BackColor=RGB(255,255,255)
       foreColor=dForeColor
       nDsplayCount=0
       procexit=''
       procact=''   
       procInit=''
       textCopy='' 
       copyname=''                          
       logExit=.T.
       procForKeyPress='' 
       procForRightClick=''   
       procForClick=''  
       copyname=''        
       ADD OBJECT formImage AS IMAGE WITH VISIBLE=.F., STRETCH=2, LEFT=0, TOP=0, WIDTH=0,HEIGHT=0
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Activate
       IF !EMPTY(This.procact)
          procdo=This.procact
          &procdo   
       ENDIF 
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Init                  
       IF !EMPTY(This.procact)
          procdo=This.procact
          &procdo   
       ENDIF     
*-----------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE Keypress      
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl 
       IF This.logexit.AND.nKeyCode=27
          This.Release
       ENDIF                                             
       IF !EMPTY(This.procForKeyPress)
          procdo=This.procForKeyPress
          &procdo        
       ENDIF       
*------------------------------------------------------------------------------------------------------------------------                                       
       PROCEDURE RightClick
       IF !EMPTY(This.procForRightClick)
           procdo=This.procForRightClick
           &procdo   
       ENDIF   
*------------------------------------------------------------------------------------------------------------------------                                       
       PROCEDURE Click
       IF !EMPTY(This.procForClick)
           procdo=This.procForClick
           &procdo   
       ENDIF          
*------------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE QueryUnload        
       IF !EMPTY(This.procexit)     
          procdo=This.procexit
          &procdo   
       ENDIF              
ENDDEFINE
*******************************************************************************************
*   Добавленние вспомогательной формы
*******************************************************************************************
PROCEDURE adFormSupl
PARAMETERS parFrm,parCaption
*parFrm - имя создаваемой формы
*parCaption - заголовок
objFrm=parFrm
parFrm=CREATEOBJECT('FORMSUPL')
*@ 10,10 say &parfrm
*dfvsadfasd
&objFrm..Caption=parCaption
*********************************************************************************************
*               Класс для создания форм (больших,например справочники)
* Задаются основные элементы и параметры, всё остальное потом
*********************************************************************************************
DEFINE CLASS formmy AS FORM
       DeskTop=.T.
       AlwaysOnTop=.T.
       Icon=cFormIco 
       Top=0
       Left=0     
       Height=SYSMETRIC(22)
       Width=SYSMETRIC(21) && 1024 
       WindowState=2          
       WindowType=1      
       ShowWindow=1                  
       DoCreate = .T.
       BorderStyle = 2	   
       Minbutton=.F. 
       Name = "Form1"
       FontName=dFontName
       FontSize=dFontSize 
       ShowTips=.T.       
       Visible =.T. 
       BackColor=dBackColor
       nDsplayCount=0
       procexit=''
       procact=''   
       procInit=''
       textCopy='' 
       copyname=''                          
       logExit=.F.
       procForKeyPress='' 
       procForRightClick=''   
       procForClick=''
       copyname=''        
       ADD OBJECT formImage AS IMAGE WITH VISIBLE=.F., STRETCH=2, LEFT=0, TOP=0, WIDTH=0,;
                  HEIGHT=0,VISIBLE=.T.                           
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Activate
       IF !EMPTY(This.procact)
          procdo=This.procact
          &procdo   
       ENDIF 
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Init                  
       IF !EMPTY(This.procact)
          procdo=This.procact
          &procdo   
       ENDIF     
*-----------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE Keypress      
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl 
       IF This.logexit.AND.nKeyCode=27
          This.Release
       ENDIF                                             
       IF !EMPTY(This.procForKeyPress)
          procdo=This.procForKeyPress
          &procdo   
       ENDIF       
*------------------------------------------------------------------------------------------------------------------------                                       
       PROCEDURE RightClick
       IF !EMPTY(This.procForRightClick)
           procdo=This.procForRightClick
           &procdo   
       ENDIF   
*------------------------------------------------------------------------------------------------------------------------                                       
       PROCEDURE Click
       IF !EMPTY(This.procForClick)
           procdo=This.procForClick
           &procdo   
       ENDIF   
*------------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE QueryUnload
       IF !EMPTY(This.procexit)
          procdo=This.procexit
          &procdo   
       ENDIF              
ENDDEFINE

*********************************************************************************************
*               Класс для создания форм (больших,например справочники)
* Задаются основные элементы и параметры, всё остальное потом
*********************************************************************************************
DEFINE CLASS formspr AS FORM
       DeskTop=.T.
       Top=0
       Left=0     
       Icon=cFormIco 
       Height=SYSMETRIC(22)
       Width=SYSMETRIC(21) && 1024 
       WindowState=2          
       WindowType=1      
       ShowWindow=1                  
       DoCreate = .T.
       BorderStyle = 2	   
       Minbutton=.F. 
       Name = "Form1"
       FontName=dFontName
       FontSize=dFontSize 
       ShowTips=.T.       
       Visible =.T. 
       BackColor=dBackColor
       nDsplayCount=0
       procexit=''
       procact=''   
       procInit=''
       strname=''
       copyname='' 
       logExit=.F. 
       nrec=0         
       DIMENSION dim_ap(FCOUNT())                       
             
       ADD OBJECT formImage AS IMAGE WITH VISIBLE=.F., STRETCH=2, LEFT=0, TOP=0, WIDTH=0,;
                  HEIGHT=0,VISIBLE=.T.  
                  
       ADD OBJECT fGrid AS gridMyNew   WITH VISIBLE=.T., LEFT=0, WIDTH=ThisForm.Width, SCROLLBARS=2               
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Activate
                 IF !EMPTY(This.procact)
                    procdo=This.procact
                    &procdo   
                 ENDIF 
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Init                                                         
                 IF !EMPTY(This.procact)
                    procdo=This.procact
                    &procdo   
                 ENDIF                               
*------------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE QueryUnload
                 IF !EMPTY(This.procexit)
                    procdo=This.procexit
                    &procdo   
                 ENDIF 
       ENDPROC  
ENDDEFINE

DEFINE CLASS formTopSupl AS Form
       DeskTop=.T.
       AlwaysOnBottom=.F.
       Icon=cFormIco 
       Top=0
       Left=0    
       Height=SYSMETRIC(22)
       Width=SYSMETRIC(21) && 1024 
       WindowState=0       
       WindowType=0      
       ShowWindow=2   
       BackColor=RGB(255,255,255)         
       Autocenter = .T.   
       
       DoCreate = .T.
       BorderStyle = 2	   
       Minbutton=.T. 
       Name = "Form1"
       FontName=dFontName
       FontSize=dFontSize 
       ShowTips=.T.       
       Visible =.T. 
       nDsplayCount=0
       procexit=''      
       procact=''   
       procInit=''
       textCopy='' 
       copyname=''                          
       logExit=.F.
       procForKeyPress='' 
       procForRightClick=''   
       copyname=''        
       ADD OBJECT formImage AS IMAGE WITH VISIBLE=.F., STRETCH=2, LEFT=0, TOP=0, WIDTH=0,;
                  HEIGHT=0,VISIBLE=.T.     
                  
       *-------------------------------------------------------------------------------------          
       PROCEDURE QueryUnload       
       IF !EMPTY(This.procexit)
          procForDo=This.procexit        
          &procForDo
          NODEFAULT           
       ELSE      
          QUIT
       ENDIF                        
       AlwaysOnTop=.T.
            
     *  ShowWindow=1                  
                        
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Activate
       IF !EMPTY(This.procact)
          procdo=This.procact
          &procdo   
       ENDIF 
*------------------------------------------------------------------------------------------------------------------------                  
       PROCEDURE Init                  
       IF !EMPTY(This.procact)
          procdo=This.procact
          &procdo   
       ENDIF     
*-----------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE Keypress      
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl 
       IF This.logexit.AND.nKeyCode=27
          This.Release
       ENDIF                                             
       IF !EMPTY(This.procForKeyPress)
          procdo=This.procForKeyPress
          &procdo   
       ENDIF       
*------------------------------------------------------------------------------------------------------------------------                                       
       PROCEDURE RightClick
       IF !EMPTY(This.procForRightClick)
           procdo=This.procForRightClick
           &procdo   
       ENDIF    
ENDDEFINE

DEFINE CLASS formtop AS Form
       DeskTop=.T.
       AlwaysOnBottom=.F.
       Top=0
       Left=0    
       Icon=cFormIco 
       Height=SYSMETRIC(22)
       Width=SYSMETRIC(21) && 1024 
       WindowState=0          
       WindowType=0      
       ShowWindow=2   
       BackColor=dBackColor            
       *AutoCenter=.T.   
       DoCreate = .T.
       BorderStyle = 2	   
       Minbutton=.T. 
       Name = "Form1"
       FontName=dFontName
       FontSize=dFontSize 
       ShowTips=.T.       
       Visible =.T. 
       nDsplayCount=0
       procexit=''
       procForResize=''
      
       *-------------------------------------------------------------------------------------
       PROCEDURE Resize
       IF !EMPTY(This.procForResize)
          procForDo=This.procForResize        
         &procForDo         
       ENDIF      
       *-----------------------------------------------------------------------------------------------------------------------
       PROCEDURE Load                  
       *-------------------------------------------------------------------------------------
       PROCEDURE Init                
       *-------------------------------------------------------------------------------------       
       PROCEDURE Activate                
       *-------------------------------------------------------------------------------------          
       PROCEDURE QueryUnload       
       IF !EMPTY(This.procexit)
          procForDo=This.procexit        
          &procForDo
          NODEFAULT           
       ELSE      
          QUIT
       ENDIF          
ENDDEFINE

****************************************************************************************************************************
DEFINE CLASS MyPageFrame AS PageFrame
       VISIBLE=.T.
       PAGECOUNT=0
       *ADD OBJECT mpage1 AS myPage WITH procForActivate=''
       *ADD OBJECT mpage2 AS myPage WITH procForActivate=''
       *ADD OBJECT mpage3 AS myPage WITH procForActivate=''
       *ADD OBJECT mpage4 AS myPage WITH procForActivate=''
       *ADD OBJECT mpage5 AS myPage WITH procForActivate=''
       *ADD OBJECT mpage6 AS myPage WITH procForActivate=''    
       
ENDDEFINE
***************************************************************************************************************************
DEFINE CLASS myPage AS Page
       BackColor=dBackColor
       ForeColor=dForeColor
       FontName=dFontName
       FontSize=dFontSize       
       procActivate=''
      * --------------------------------------------------------------------
       PROCEDURE ACTIVATE
       IF !EMPTY(This.procActivate)
          procDo=This.procActivate
          &procDo
       ENDIF
ENDDEFINE
****************************************************************************************************************************
PROCEDURE addPageFrame
PARAMETERS parFrm,parName,parcount,partop,parleft,parWidth,parheight,parVisible
objPageFrame=parname
&parfrm..AddObject(objPageFrame,'myPageFrame')
WITH &parfrm..&objPageFrame    
    
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=parHeight    
     .pageCount=0
     .Visible=parVisible
     .SetAll('fontSize',dFontSize,'page')
     .SetAll('fontName',dFontName,'page')
ENDWITH
***************************************************************************************************************************
PROCEDURE addPageToPageFrame
LPARAMETERS parPageFrame,parPageCount
*parPageFrame - pafeGrame
*parPageCount - кол-во страниц
WITH parPageFrame
*     FOR i=1 TO parPageCount
       .ADDOBJECT('mpage3','myPage')
*         .ADDOBJECT('page'+LTRIM(STR(i)),'myPage')         
*     ENDFOR     
ENDWITH
****************************************************************************************************************************
DEFINE CLASS myTxtBoxGrd AS TextBox
       FontName=dFontName
       FontSize=dFontSize
       SelectedBackColor=selBackColor
       SelectedForeColor=selForeColor
       DisabledBackColor=objBackColor
*       BackColor=dBackColor
       ForeColor=dForeColor
       SelectOnEntry=.T.
       Visible=.T.
       BorderStyle=0
       valGotFocus=0
       parold=''
       procForValid=''
       varCtrlSource=''
       procGotFocus=''
       procLostFocus=''
       procRightClick=''
       procForDblClick=''    
       procForKeyPress=''
       procForClick=''
       logExit=.F.
       *--------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl
       This.ReadOnly=.F.     
       *-----------------------------------------------------------------------------------------------------
       PROCEDURE VALID     
       IF !EMPTY(This.procForValid)
          procDo=This.procForValid
          &procDo
       ENDIF
       *-----------------------------------------------------------------------------
       PROCEDURE GotFocus       
       IF !EMPTY(This.procGotFocus)
          procDo=This.procGotFocus
          &procDo
       ENDIF            
       *-----------------------------------------------------------------------------------------
       PROCEDURE LostFocus
       IF !EMPTY(This.procLostFocus)
          procDo=This.procLostFocus
          &procDo
       ENDIF
       *----------------------------------------------------------------------------------
       PROCEDURE RightClick
       IF !EMPTY(This.procRightClick)
          procDo=This.procRightClick
         &procDo
       ENDIF
        *------------------------------------------------------------------------------------------------
       PROCEDURE Click
       IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo
       ENDIF 
       *------------------------------------------------------------------------------------------------
       PROCEDURE DblClick
       IF ! EMPTY(THIS.procForDblClick)
          procForDo=THIS.procForDblClick
          &procForDo
       ENDIF 
       *--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl
       IF This.logexit.AND.nKeyCode=27
           ThisForm.Release
       ENDIF 
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF     
ENDDEFINE

DEFINE CLASS myTxtBox AS TextBox
       FontName=dFontName
       FontSize=dFontSize
       ForeColor=objForeColor
       BackColor=ObjBackColor
       DisabledForeColor=ObjForeColor
       DisabledBackColor=ObjBackColor
       SelectedBackColor=selBackColor
       SelectedForecolor=selForeColor
       SelectOnEntry=.T.  
       BorderStyle=1
       BackStyle=0
       Visible=.T.
       SpecialEffect=1       
       varSource=''
  *     ControlSource='varSource'
       procForGotFocus=''
       procForLostFocus=''
       procForValid=''
       procForRightClick=''
       procForClick=''
       procForKeyPress=''
       procForChange=''
       logExit=.F.
       *--------------------------------------------------------------------------
       PROCEDURE GotFocus
       SET CURSOR ON
       IF ! EMPTY(THIS.procForGotFocus)
          procForDo=THIS.procForGotFocus
          &procForDo
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE LostFocus
       IF ! EMPTY(THIS.procForLostFocus)
          procForDo=THIS.procForLostFocus
          &procForDo
       ENDIF
       SET CURSOR OFF
       *--------------------------------------------------------------------------
       PROCEDURE Valid
       IF ! EMPTY(THIS.procForValid)
          procForDo=THIS.procForValid
          &procForDo
       ENDIF 
       *------------------------------------------------------------------------------------------------
       PROCEDURE RightClick
       IF ! EMPTY(THIS.procForRightClick)
          procForDo=THIS.procForRightClick
          &procForDo
       ENDIF 
       *------------------------------------------------------------------------------------------------
       PROCEDURE Click
       IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo
       ENDIF 
       *--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl
       IF This.logexit.AND.nKeyCode=27
           ThisForm.Release
       ENDIF 
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF  
        *------------------------------------------------------------------------------------------------
       PROCEDURE interActiveChange
       IF ! EMPTY(THIS.procForChange)
          procForDo=THIS.procForChange
          &procForDo
       ENDIF        
       *-------------------------------------------------------------------------------------------------  
       *   Стандартное меню для редактирования некоторых объектов
       *-------------------------------------------------------------------------------------------------
       PROCEDURE menread
       PARAMETERS parfrm,parVar,parField
       DEFINE POPUP short SHORTCUT RELATIVE FROM MROW(),MCOL() FONT dFontName,dFontSize COLOR SCHEME 4
       DEFINE BAR 1 OF short PROMPT 'Вырезать'
       DEFINE BAR 2 OF short PROMPT "\-"
       DEFINE BAR 3 OF short PROMPT 'Копировать'  
       DEFINE BAR 4 OF short PROMPT "\-"
       DEFINE BAR 5 OF short PROMPT 'Вставить' &&SKIP FOR EMPTY(&parVar)
       *ON SELECTION POPUP short DO This.proccopy  
       ACTIVATE POPUP short
       *---------------------------------------------------------------------------------------------------
       PROCEDURE proccopy
       men_cx=BAR()
       DEACTIVATE POPUP short
       DO CASE
           CASE men_cx=1
                &parVar=This.Seltext
                newtext=LEFT(&parfield,This.SelStart)+SUBSTR(&parfield,This.SelStart+1+This.SelLength)
                REPLACE &parField WITH newtext  
                &parfrm..Refresh     
           CASE men_cx=3 
                &parVar=This.Seltext
           CASE men_cx=5
                newtext=LEFT(&parfield,This.SelStart)+&parVar+SUBSTR(&parfield,This.SelStart+1)       
                REPLACE &parField WITH newtext      
                &parfrm..Refresh   
       ENDCASE                         
ENDDEFINE



DEFINE CLASS myTextBox AS TextBox
       FontName=dFontName
       FontSize=dFontSize
       ForeColor=objForeColor
       BackColor=ObjBackColor
       DisabledForeColor=ObjForeColor
       DisabledBackColor=ObjBackColor
       SelectedBackColor=selBackColor
       SelectedForecolor=selForeColor
       SelectOnEntry=.T.  
       BorderStyle=1
       BackStyle=0
       Visible=.T.
       SpecialEffect=1       
       varSource=''
  *     ControlSource='varSource'
       procForGotFocus=''
       procForLostFocus=''
       procForValid=''
       procForRightClick=''
       procForClick=''
       procForKeyPress=''
       *--------------------------------------------------------------------------
       PROCEDURE GotFocus
       SET CURSOR ON
       IF ! EMPTY(THIS.procForGotFocus)
          procForDo=THIS.procForGotFocus
          &procForDo
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE LostFocus
       IF ! EMPTY(THIS.procForLostFocus)
          procForDo=THIS.procForLostFocus
          &procForDo
       ENDIF
       SET CURSOR OFF
       *--------------------------------------------------------------------------
       PROCEDURE Valid
       IF ! EMPTY(THIS.procForValid)
          procForDo=THIS.procForValid
          &procForDo
       ENDIF 
       *------------------------------------------------------------------------------------------------
       PROCEDURE RightClick
       IF ! EMPTY(THIS.procForRightClick)
          procForDo=THIS.procForRightClick
          &procForDo
       ENDIF 
       *------------------------------------------------------------------------------------------------
       PROCEDURE Click
       IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo
       ENDIF 
       *--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF         
       *-------------------------------------------------------------------------------------------------  
       *   Стандартное меню для редактирования некоторых объектов
       *-------------------------------------------------------------------------------------------------
       PROCEDURE menread
       PARAMETERS parfrm,parVar,parField
       DEFINE POPUP short SHORTCUT RELATIVE FROM MROW(),MCOL() FONT dFontName,dFontSize COLOR SCHEME 4
       DEFINE BAR 1 OF short PROMPT 'Вырезать'
       DEFINE BAR 2 OF short PROMPT "\-"
       DEFINE BAR 3 OF short PROMPT 'Копировать'  
       DEFINE BAR 4 OF short PROMPT "\-"
       DEFINE BAR 5 OF short PROMPT 'Вставить' &&SKIP FOR EMPTY(&parVar)
       *ON SELECTION POPUP short DO This.proccopy  
       ACTIVATE POPUP short
       *---------------------------------------------------------------------------------------------------
       PROCEDURE proccopy
       men_cx=BAR()
       DEACTIVATE POPUP short
       DO CASE
           CASE men_cx=1
                &parVar=This.Seltext
                newtext=LEFT(&parfield,This.SelStart)+SUBSTR(&parfield,This.SelStart+1+This.SelLength)
                REPLACE &parField WITH newtext  
                &parfrm..Refresh     
           CASE men_cx=3 
                &parVar=This.Seltext
           CASE men_cx=5
                newtext=LEFT(&parfield,This.SelStart)+&parVar+SUBSTR(&parfield,This.SelStart+1)       
                REPLACE &parField WITH newtext      
                &parfrm..Refresh   
       ENDCASE                         
ENDDEFINE





**********************************************************************************************
*                    Класс для создания надписей
**********************************************************************************************
DEFINE CLASS labelmy AS Label       
       Visible=.T.             
       BackStyle=0
       Forecolor=ObjForeColor
       FontName=dFontName
       FontSize=dFontSize
       ProcForClick=''
       PROCEDURE Init
                 This.BackColor=This.Parent.BackColor       
       PROCEDURE Click
                 IF !EMPTY(This.ProcForClick)
                    ProcDo=This.ProcForClick
                    &ProcDo 
                 ENDIF              
ENDDEFINE
**********************************************************************************************
*                    Класс для создания надписей (в рамке)
**********************************************************************************************
DEFINE CLASS labelAsCont AS Label       
       Visible=.T.             
       BackStyle=1
       BorderStyle=1
       Forecolor=ObjForeColor
       FontName=dFontName
       FontSize=dFontSize
       ProcForClick=''       
       PROCEDURE Init
                 This.BackColor=This.Parent.BackColor       
       PROCEDURE Click
                 IF !EMPTY(This.ProcForClick)
                    ProcDo=This.ProcForClick
                    &ProcDo 
                 ENDIF              
ENDDEFINE



*************************************************************************************************************************
*         Процедура добавления объекта "Label" к форме (новый вариант)
*************************************************************************************************************************
PROCEDURE adlabAsCont
PARAMETERS parFrm,parName,parCaption,parTop,parLeft,parWidth,parHeight,parAlignment
&& parFrm - форма
&& parName - имя элемента
&& parCaption - зоголовок
&& parTop -верх
&& parLeft -лево
&& parWidth -ширина
&& parAlignment -выравнивание
&& parAutoSize - авторазмер
&& parBackStyle - BackStyle
obj_lab=parName
&parFrm..ADDOBJECT(obj_lab,'labelmy')
WITH &parFrm..&obj_lab
     .BackColor=dBackColor
     .Caption=parCaption
     .Top=parTop
     .Left=parLeft
     .Height=parHeight
     .Width=parWidth
     .Alignment=IIF(!EMPTY(parAlignment),parAlignment,0)
     .BorderStyle=1
     .BackStyle=1
*     .AutoSize=IIF(parAutoSize,.T.,.F.)
*     .BackStyle=IIF(!EMPTY(parBackStyle),parBackStyle,0)
     .Visible=.T.
ENDWITH




*********************************************************************************************************************
*                            Класс для создания GRID
**********************************************************************************************************************
DEFINE CLASS gridmy AS GRID
       DeleteMark=.F.
       Visible=.T.
       BackColor=dBackColor
       ForeColor=dForeColor
       GridLineColor=dForecolor
       fontname=dFontName
       fontSize=dFontSize
       RowHeight=dRowHeight    
       AllowRowSizing=.F.         
       relRow=0
       logNoDefault=0
       colnesinf=0 
       rowsgrid=0
       nRelRow=0
       procBeforeRowColChange=''
       procAfterRowColChange='' 
       procForWhen=''
       fieldUpdate='' 
       CurRec=0
*-----------------------------------------------------------------------------------------------------------       
       PROCEDURE Init    
       This.BackColor=This.Parent.BackColor     
       FOR i=1 TO This.ColumnCount         
           This.Columns(i).DynamicBackColor='IIF(RECNO(This.RecordSource)#This.curRec,dBackColor,dynBackColor)'
           This.Columns(i).DynamicForeColor='IIF(RECNO(This.RecordSource)#This.curRec,dForeColor,dynForeColor)'
           This.Columns(i).Resizable=.F.
           This.Columns(i).Text1.SelectedForeColor=dynForeColor
           This.Columns(i).Text1.SelectedBackColor=dynBackColor        
           This.Columns(i).Header1.ForeColor=dForeColor
           This.Columns(i).Header1.BackColor=headerBackColor
           This.Columns(i).Header1.FontName=dFontName
           This.Columns(i).Header1.FontSize=dFontSize
           This.Columns(i).FontSize=dFontSize
           This.Columns(i).FontName=dFontName                    
       ENDFOR   
*------------------------------------------------------------------------------------------------------------
       PROCEDURE Gridsetup  
       PARAMETERS parBackColor
       FOR i=1 TO This.ColumnCount         
           This.Columns(i).DynamicBackColor='IIF(RECNO(This.RecordSource)#This.curRec,IIF(!EMPTY(parBackColor),dBackColor,parBackColor),dynBackColor)'
           This.Columns(i).DynamicForeColor='IIF(RECNO(This.RecordSource)#This.curRec,dForeColor,dynForeColor)'
           This.Columns(i).Text1.SelectedBackColor=selBackColor                  
           This.Columns(i).Header1.ForeColor=dForeColor
           This.Columns(i).Header1.BackColor=headerBackColor         
           This.Columns(i).Header1.Alignment=2
          
       ENDFOR  
*------------------------------------------------------------------------------------------------------------       
       PROCEDURE BeforeRowColChange
       LPARAMETERS nColIndex
       IF THIS.logNoDefault#0
          THIS.logNoDefault=0
          RETURN
       ENDIF
       SET CURSOR OFF      
       IF ncolindex=THIS.ColNesInf.AND.EMPTY(THIS.VALUE)
          THIS.VALUE=EVALUATE(THIS.Columns(THIS.ColNesInf).ControlSource)
       ENDIF 
       IF !EMPTY(This.procBeforeRowColChange)
           procDo=This.procBeforeRowColChange
           &procDo
       ENDIF    
*----------------------------------------------------------------------------------------------------
       PROCEDURE AfterRowColChange
       LPARAMETERS nColIndex
       This.nRelRow=This.RelativeRow
       This.curRec=RECNO(This.RecordSource)     
       This.Refresh
       SET CURSOR ON
       IF ! THIS.Columns(nColIndex).ENABLED
          KEYBOARD '{ENTER}'
       ENDIF   
       IF THIS.Columns(nColIndex).Readonly
          SET NOTIFY OFF
          SET CURSOR OFF
       ENDIF
       IF !EMPTY(This.procAfterRowColChange)
           procDo=This.procAfterRowColChange
           &procDo
       ENDIF
       SET CURSOR ON
*----------------------------------------------------------------------------------------------------       
*        Добавление строки в Grid при формировании справочников   
*----------------------------------------------------------------------------------------------------
       PROCEDURE GridmyAppendBlank 
       PARAMETERS parOrd,parKod,parName  
       basech=This.recordSource
       SELECT &basech
       This.GridUpdate  
       log_ord=SYS(21)
       SET ORDER TO parOrd
       GO TOP
       newkod=1
       DO WHILE !EOF()
          IF newkod#EVALUATE(parKod)
             EXIT
          ENDIF
          SKIP
          newkod=newkod+1   
       ENDDO              
       SET DELETED OFF
       LOCATE FOR EMPTY(EVALUATE(parName))
       IF FOUND()
          RECALL
          BLANK
       ELSE
          APPEND BLANK
       ENDIF
       SET DELETED ON
       SET ORDER TO EVALUATE(log_ord)      
       REPLACE &parKod WITH newkod
       This.Refresh
*       This.column2.SetFocus
       *--------------------------------------------------------------------------
       * Проверка на законченность ввода и непустое значение редактируемого поля
       *--------------------------------------------------------------------------
       PROCEDURE GridUpdate     
       PARAMETERS parFl 
       THIS.SetFocus  
       basech=This.recordSource    
       nrec=THIS.ActiveRow      
       IF THIS.ActiveColumn=THIS.ColNesInf.AND.EMPTY(THIS.VALUE).AND.This.ActiveColumn#0        
          THIS.VALUE=EVALUATE(THIS.COLUMNS(This.ActiveColumn).ControlSource)          
       ENDIF
       GO TOP
       THIS.Refresh       
       IF nrec#0
          GO nrec         
       ENDIF       
       SELECT &basech
       varOrd=SYS(21)
     *  SET ORDER TO 2
       DO WHILE .T.
          LOCATE FOR EMPTY(&parFl)
          IF FOUND()                              
             DELETE
          ELSE
             EXIT
          ENDIF
       ENDDO     
       IF varOrd#'0'
          SET ORDER TO EVALUATE(VarOrd)
       ENDIF   
*--------------------------------------------------------------------------
*                  Удаление строки Grid
*----------------------------------------------------------------------------------------------------
       PROCEDURE GridDelRec 
       PARAMETERS parGrid,parTable    
       DO createFormNew WITH .T.,'Удаление',RetTxtWidth('WWУдалить выбранную запись?WW',dFontName,dFontSize+1),;
       '130',RetTxtWidth('WWНетWW',dFontName,dFontSize+1),'Да','Нет',.F.,'&parGrid..GridDelOne','nFormMes.Release',.F.,;
       'Удалить выбранную запись?',.F.,.T.
*-------------------------------------------------------------------------------- -      
*               Невозмжоность удаления строки Grid
*---------------------------------------------------------------------------------
       PROCEDURE GridNoDelRec        
       DO createFormNew WITH .T.,'Удаление',RetTxtWidth('WWУдаление невозможно!WW',dFontName,dFontSize+1),'130',;
          RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
          'Удаление невозможно!',.F.,.T.                                
*----------------------------------------------------------------------------------------------------------------
*      Непосредственно удаление записи
*----------------------------------------------------------------------------------------------------------------  
       PROCEDURE GridDelOne
       SELECT (parTable)
       DELETE
       This.Refresh
       FormMy.Release
*----------------------------------------------------------------------------------------------------------------
*
*----------------------------------------------------------------------------------------------------------------         
       PROCEDURE gridReturn
       This.GridUpdate     
       ThisForm.Release
ENDDEFINE

DEFINE CLASS HeaderMy AS Header
       Fontname=dFontName
       FontSize=dFontSize
       ForeColor=dForeColor
       BackColor=headerBackColor
       ALIGNMENT=2
       Visible=.T.       
       procForClick=''
       *--------------------------------------------------------------------------
       PROCEDURE Click
       IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo       
       ENDIF 
ENDDEFINE


*********************************************************************************************************************
*                            Класс для создания GRID
**********************************************************************************************************************
DEFINE CLASS gridmynew AS GRID
       DeleteMark=.F.
       Visible=.T.
       BackColor=dBackColor
       ForeColor=dForeColor
       GridLineColor=dForecolor
       fontname=dFontName
       fontSize=dFontSize
       RowHeight=dRowHeight    
       AllowRowSizing=.F.   
       columncount=0     
       relRow=0
       logNoDefault=0
       colnesinf=0 
       rowsgrid=0
       nRelRow=0
       procBeforeRowColChange=''
       procAfterRowColChange='' 
       fieldUpdate=''  
       CurRec=0
*-----------------------------------------------------------------------------------------------------------       
       PROCEDURE Init    
       This.BackColor=This.Parent.BackColor     
       FOR i=1 TO This.ColumnCount         
           This.Columns(i).DynamicBackColor='IIF(RECNO(This.RecordSource)#This.curRec,dBackColor,dynBackColor)'
           This.Columns(i).DynamicForeColor='IIF(RECNO(This.RecordSource)#This.curRec,dForeColor,dynForeColor)'
           This.Columns(i).Resizable=.F.
           This.Columns(i).Text1.SelectedForeColor=dynForeColor
           This.Columns(i).Text1.SelectedBackColor=dynBackColor        
           This.Columns(i).Header1.ForeColor=dForeColor
           This.Columns(i).Header1.BackColor=headerBackColor
           This.Columns(i).Header1.FontName=dFontName
           This.Columns(i).Header1.FontSize=dFontSize
           This.Columns(i).FontSize=dFontSize
           This.Columns(i).FontName=dFontName                    
       ENDFOR   
*------------------------------------------------------------------------------------------------------------
       PROCEDURE Gridsetup  
       PARAMETERS parBackColor
       FOR i=1 TO This.ColumnCount         
           This.Columns(i).DynamicBackColor='IIF(RECNO(This.RecordSource)#This.curRec,IIF(!EMPTY(parBackColor),dBackColor,parBackColor),dynBackColor)'
           This.Columns(i).DynamicForeColor='IIF(RECNO(This.RecordSource)#This.curRec,dForeColor,dynForeColor)'
           This.Columns(i).Text1.SelectedBackColor=selBackColor                  
           This.Columns(i).Header1.ForeColor=dForeColor
           This.Columns(i).Header1.BackColor=headerBackColor         
           This.Columns(i).Header1.Alignment=2
          
       ENDFOR  
*------------------------------------------------------------------------------------------------------------       
       PROCEDURE BeforeRowColChange
       LPARAMETERS nColIndex
       IF THIS.logNoDefault#0
          THIS.logNoDefault=0
          RETURN
       ENDIF
       SET CURSOR OFF      
       IF ncolindex=THIS.ColNesInf.AND.EMPTY(THIS.VALUE)
          THIS.VALUE=EVALUATE(THIS.Columns(THIS.ColNesInf).ControlSource)
       ENDIF 
       IF !EMPTY(This.procBeforeRowColChange)
           procDo=This.procBeforeRowColChange
           &procDo
       ENDIF    
*----------------------------------------------------------------------------------------------------
       PROCEDURE AfterRowColChange
       LPARAMETERS nColIndex
       This.nRelRow=This.RelativeRow
       This.curRec=RECNO(This.RecordSource)     
       This.Refresh
       SET CURSOR ON
       IF ! THIS.Columns(nColIndex).ENABLED
          KEYBOARD '{ENTER}'
       ENDIF   
       IF THIS.Columns(nColIndex).Readonly
          SET NOTIFY OFF
          SET CURSOR OFF
       ENDIF
       IF !EMPTY(This.procAfterRowColChange)
           procDo=This.procAfterRowColChange
           &procDo
       ENDIF
       SET CURSOR ON
*----------------------------------------------------------------------------------------------------       
*        Добавление строки в Grid при формировании справочников   
*----------------------------------------------------------------------------------------------------
       PROCEDURE GridmyAppendBlank 
       PARAMETERS parOrd,parKod,parName
       basech=This.recordSource
       SELECT &basech
       This.GridUpdate(parName)  
       log_ord=SYS(21)
       SET ORDER TO parOrd
       GO TOP
       newkod=1
       DO WHILE !EOF()
          IF newkod#EVALUATE(parKod)
             EXIT
          ENDIF
          SKIP
          newkod=newkod+1   
       ENDDO              
       SET DELETED OFF
       LOCATE FOR EMPTY(EVALUATE(parName))
       IF FOUND()
          RECALL
          BLANK
       ELSE
          APPEND BLANK
       ENDIF
       SET DELETED ON
       SET ORDER TO EVALUATE(log_ord)      
       REPLACE &parKod WITH newkod
       This.Refresh
*       This.column2.SetFocus
       *--------------------------------------------------------------------------
       * Проверка на законченность ввода и непустое значение редактируемого поля
       *--------------------------------------------------------------------------
       PROCEDURE GridUpdate    
       PARAMETERS par1
       THIS.SetFocus  
       basech=This.recordSource    
       nrec=THIS.ActiveRow      
       IF THIS.ActiveColumn=THIS.ColNesInf.AND.EMPTY(THIS.VALUE).AND.This.ActiveColumn#0        
          THIS.VALUE=EVALUATE(THIS.COLUMNS(This.ActiveColumn).ControlSource)          
       ENDIF
       GO TOP
       THIS.Refresh       
       IF nrec#0
          GO nrec         
       ENDIF       
       SELECT &basech
       varOrd=SYS(21)
     *  SET ORDER TO 2
       DO WHILE .T.
          IF !EMPTY(par1)
             LOCATE FOR EMPTY(&par1)
          ELSE 
             LOCATE FOR EMPTY(name)
          ENDIF    
          IF FOUND()                              
             DELETE
          ELSE
             EXIT
          ENDIF
       ENDDO     
       IF varOrd#'0'
          SET ORDER TO EVALUATE(VarOrd)
       ENDIF   

*--------------------------------------------------------------------------
*                  Удаление строки Grid
*----------------------------------------------------------------------------------------------------
       PROCEDURE GridDelRec 
       PARAMETERS parGrid,parTable    
       DO createFormNew WITH .T.,'Удаление',RetTxtWidth('WWУдалить выбранную запись?WW',dFontName,dFontSize+1),;
       '130',RetTxtWidth('WWНетWW',dFontName,dFontSize+1),'Да','Нет',.F.,'&parGrid..GridDelOne','nFormMes.Release',.F.,;
       'Удалить выбранную запись?',.F.,.T.
*-------------------------------------------------------------------------------- -      
*               Невозмжоность удаления строки Grid
*---------------------------------------------------------------------------------
       PROCEDURE GridNoDelRec        
       DO createFormNew WITH .T.,'Удаление',RetTxtWidth('WWУдаление невозможно!WW',dFontName,dFontSize+1),'130',;
          RetTxtWidth('WWОКWW',dFontName,dFontSize+1),'OK',.F.,.F.,'nFormMes.Release',.F.,.F.,;
          'Удаление невозможно!',.F.,.T.                                
*----------------------------------------------------------------------------------------------------------------
*      Непосредственно удаление записи
*----------------------------------------------------------------------------------------------------------------  
       PROCEDURE GridDelOne
       SELECT (parTable)
       DELETE
       This.Refresh
       FormMy.Release
*----------------------------------------------------------------------------------------------------------------
*
*----------------------------------------------------------------------------------------------------------------         
       PROCEDURE gridReturn
       PARAMETERS parFl
       This.GridUpdate(parFl)     
       ThisForm.Release
ENDDEFINE
*------------------------------Класс для колонок в Grid--------------------------------------------------------------------------------------------
DEFINE CLASS ColumnMy AS Column
       BOUND=.F.
       Resizable=.F.
       Movable=.F.
       Fontname=dFontName
       FontSize=dFontSize
       Alignment=0
       Visible=.T.
       procForMouseMove=''
       *--------------------------------------------------------------------------
       PROCEDURE MouseMove
       LPARAMETERS nIndex, nButton, nShift, nXCoord, nYCoord
       IF This.ReadOnly=.T..OR.! This.Enabled
          ThisForm.MousePointer=1
       ENDIF   
       IF ! EMPTY(THIS.procForMouseMove)
          procForDo=THIS.procForMouseMove
          &procForDo
       ENDIF
ENDDEFINE       




************************************************************************************************************************
*                            Класс для создания ListBox
************************************************************************************************************************
DEFINE CLASS listboxmy AS ListBox
       SpecialEffect=1
       BorderStyle=2
       Visible=.T.
       SelectedBackColor=selBackColor
       SelectedItemBackColor=selBackColor
       SelectedItemForeColor=selForeColor
       ItemForeColor=dForeColor
       FontName=dFontName
       FontSize=dFontSize       
       varCtrlSource=''
       ControlSource='This.varCtrlSource'
       procForKeyPress=''
       procForClick=''
       procForDblClick=''
       procForRightClick=''
       procForGotFocus=''
       procForLostFocus=''
       procForValid=''
       procForInit=''
       logExit=.F.
       *---------------------------------------------------------------------------
       PROCEDURE Init
       IF ! EMPTY(THIS.procForInit)
          procForDo=THIS.procForInit
          &procForDo       
       ENDIF 
       *---------------------------------------------------------------------------
       PROCEDURE Click  
       IF INLIST(LASTKEY(),5,24,56,50)
       		RETURN 	
       ENDIF 
       IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo       
       ENDIF 
       *---------------------------------------------------------------------------
       PROCEDURE dblClick  
       IF ! EMPTY(THIS.procForDblClick)
          procForDo=THIS.procForDblClick
          &procForDo       
       ENDIF 
       *---------------------------------------------------------------------------
       PROCEDURE RightClick   
       IF ! EMPTY(THIS.procForRightClick)
          procForDo=THIS.procForRightClick
          &procForDo       
       ENDIF 
       *---------------------------------------------------------------------------
       PROCEDURE KeyPress   
       LPARAMETERS nKeyCode,nShiftAltCtrl
       IF This.logExit.AND.nKeyCode=27
           ThisForm.Release
       ENDIF 
       IF ! EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress
          &procForDo       
       ENDIF 
       *---------------------------------------------------------------------------
       PROCEDURE GotFocus   
       IF ! EMPTY(THIS.procForGotFocus)
          procForDo=THIS.procForGotFocus
          &procForDo       
       ENDIF  
       *---------------------------------------------------------------------------
       PROCEDURE LostFocus   
       IF ! EMPTY(THIS.procForLostFocus)
          procForDo=THIS.procForLostFocus
          &procForDo       
       ENDIF
       *---------------------------------------------------------------------------
       PROCEDURE Valid    
       IF ! EMPTY(THIS.procForValid)
          procForDo=THIS.procForValid
          &procForDo       
       ENDIF      
ENDDEFINE

************************************************************************************************************************
*                       Процедура добавления объекта ListBox в форму
************************************************************************************************************************
PROCEDURE addListboxmy
PARAMETERS parfrm,parord,parleft,partop,parheight,parwidth
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parheight - высота
* parwidth - щирина
 
obj_lbox='listbox'+LTRIM(STR(parord))
&parfrm..AddObject(obj_lbox,'ListBoxMy')
WITH &parfrm..&obj_lbox     
     .Top=partop
     .Left=parleft
     .Width=parwidth
     .Height=parheight     
ENDWITH

*********************************************************************************************************************
*                            Класс для создания ComboBox
**********************************************************************************************************************
DEFINE CLASS cbomy AS ComboBox
       BorderStyle=0
       SelectedBackColor=selBackColor
       SelectedForeColor=selForeColor
       SelectedItemBackColor=selBackColor
       SelectedItemForeColor=selForeColor 
       ItemForeColor=objForeColor      
       FontName=dFontName
       FontSize=dFontSize
*       Height=CurTotHeight
       Visible=.T.
       Style=2
       BackStyle=0
       
       lenPixBeg=0                  && --     Переменные используемые
       chr32Maxbeg=0                &&   | при построении List DropDown
       lenPix=0                     &&   |    в качестве меню Popup 
       chr32Max=0                   &&   |     (RowSourceType=9) 
       namePopCbo=''                &&   |     
       nBar=0                       && -- 
                
       logNoDefault=0               && триггерная переменная для управления обработки событий
       varCtrlSource=''             && Переменная для ControlSource
       procGotFocus=''           && User-defined Property - Команда или процедура
                                    && обрабатываемая в событии GotFocus
       procForLostFocus=''          && User-defined Property - Команда или процедура
                                    && обрабатываемая в событии LostFocus
       procForDropDown=''           && User-defined Property - Команда или процедура
                                    && обрабатываемая в событии DropDown
       procForValid=''              && User-defined Property - Команда или процедура
                                    && обрабатываемая в событии Valid
       procForChange=''       
       procForKeyPress=''                        
       *--------------------------------------------------------------------------
       PROCEDURE Valid     
       IF ! EMPTY(THIS.procForValid)
          procForDo=THIS.procForValid         
          &procForDo
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE GotFocus
       THIS.logNodefault=0
       IF ! EMPTY(THIS.procGotFocus)
          procForDo=THIS.procGotFocus
          &procForDo
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE DropDown
       THIS.logNodefault=1
       IF ! EMPTY(THIS.procForDropDown)
          procForDo=THIS.procForDropDown 
          &procForDo
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE Click
       IF THIS.logNodefault=1
          THIS.logNodefault=2
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl
       *IF THIS.logNodefault=1.AND.LASTKEY()=13
       *   THIS.logNodefault=2
       *ENDIF
       IF LASTKEY()=27
          *KEYBOARD '{TAB}'       && для выхода из ComboBox по нажатию клавиши Esc
       ENDIF                     && и переходу на следующее поле
       *--------------------------------------------------------------------------
       PROCEDURE LostFocus
       IF ! EMPTY(THIS.procForLostFocus)
          procForDo=THIS.procForLostFocus
          &procForDo
       ENDIF
       *--------------------------------------------------------------
        PROCEDURE InterActiveChange
        IF ! EMPTY(THIS.procForChange)
           procDo=THIS.procForChange
           &procDo
       ENDIF
       *----------------------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       PARAMETERS nKeyCode, nShiftAltCtrl
       IF !EMPTY(This.procForKeyPress)
          procForDo=This.procForKeyPress
          &procForDo
       ENDIF
       *-----------------------------------------------------------------------
	   *     Присвоение выбранного значения ComboBox соответствующему полю
	   *-----------------------------------------------------------------------
	   PROCEDURE CboValidPop	   
	   PARAMETERS parFieldGrd,parAliasREl,parFieldRel
       IF THIS.logNodefault=0      
          DO CASE
    	     CASE LASTKEY()=5
                  IF ! BOF()      
                     SKIP-1  
    	          ENDIF
             CASE LASTKEY()=24
		          IF ! EOF()      
	                 SKIP
	              ENDIF
             CASE LASTKEY()=13
                  KEYBOARD '{TAB}'
          ENDCASE
       ELSE 
          IF THIS.logNodefault=2
             IF BETWEEN(THIS.ListIndex,1,RECCOUNT(parAliasRel))
            	REPLACE (parFieldGrd) WITH EVALUATE(parFieldRel)
            	
                KEYBOARD '{TAB}'                && для выхода из ComboBox
             ENDIF
          ENDIF
       ENDIF
       *-----------------------------------------------------------------------
       *       Обработка события GotFocus при RowSourceType=9 (Popup)
       *-----------------------------------------------------------------------
 	   PROCEDURE CboGotFocPop
	   PARAMETERS parAlias,parGrd,parAliasGrd
	   SELECT (parAlias)
	   varCountRec=RECNO()
	   GO TOP
	   COUNT WHILE RECNO()#varCountRec TO THIS.nBar
	   This.nBar=This.nBar+1
       THIS.varCtrlSource=THIS.nBar
	   This.DisplayCount=MAX(&parGrd..RelativeRow,&parGrd..RowsGrid-&parGrd..RelativeRow)
       This.DisplayCount=MIN(This.DisplayCount,RECCOUNT())
	   SELECT (parAliasGrd)
ENDDEFINE
**************************************************************************************************************************
DEFINE CLASS combomy AS ComboBox
       BorderStyle=1      
       ForeColor=objForeColor 
       DisabledBackColor=dBackColor    
       DisabledBackColor=This.BackColor 
       SelectedBackColor=selBackColor  
       SelectedForeColor=selForeColor  
       SelectedItemBackColor=selBackColor
       SelectedItemForeColor=selForeColor
       ItemForeColor=objForeColor
       DisabledForeColor=dForeColor              
       
       FontName=dFontName
       FontSize=dFontSize       
       Visible=.T.
       Style=2
       BackStyle=0
       SpecialEffect=1
       varCtrlSource=''         
       lenPix=0
       chr32Max=0
       nDisplayCount=0
       logExit=.F.
       procForGotFocus=''
       procForLostFocus=''
       procForDropDown='' 
       procForValid=''
       procForChange=''
       procForRightClick=''
       procForClick=''
       procForKeyPress=''              
       procForkeyPress=''
       procForMouseDown=''
       *--------------------------------------------------------------------------
       PROCEDURE Init
       This.BackColor=This.Parent.BackColor
       This.DisabledBackColor=This.Parent.BackColor
       This.DisabledForeColor=dForeColor
       *--------------------------------------------------------------------------
       PROCEDURE DropDown
       IF ! EMPTY(THIS.procForDropDown)
          procForDo=THIS.procForDropDown
          &procForDo       
       ENDIF 
       *--------------------------------------------------------------------------
       PROCEDURE GotFocus
       IF ! EMPTY(THIS.procForGotFocus)
          procForDo=THIS.procForGotFocus
          &procForDo
       ELSE
          IF This.nDisplayCount#0
	         This.DisplayCount=This.nDisplayCount
	      ENDIF
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE LostFocus
        IF ! EMPTY(THIS.procForLostFocus)
          procForDo=THIS.procForLostFocus
          &procForDo       
       ENDIF
       *--------------------------------------------------------------------------       
       PROCEDURE RightClick
        IF ! EMPTY(THIS.procForRightClick)
          procForDo=THIS.procForRightClick
          &procForDo       
       ENDIF
       *--------------------------------------------------------------------------       
       PROCEDURE Click
        IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo       
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE Valid
       IF ! EMPTY(THIS.procForValid)
          procForDo=THIS.procForValid
          &procForDo
       ENDIF
       *----------------------------------------------------------------------------
       PROCEDURE InterActiveChange
        IF ! EMPTY(THIS.procForChange)
           procDo=THIS.procForChange
           &procDo
       ENDIF
       *--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl
      * IF This.logexit.AND.nKeyCode=27
      *     ThisForm.Release
      * ENDIF 
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF   
       *----------------------------------------------------------------------------------------------------------------
       PROCEDURE MouseDown 
       LPARAMETERS nButton, nShift, nXCoord, nYCoord      
       IF !EMPTY(This.procForMouseDown)
          procForDo=This.procForMouseDown
          &procForDo
       ENDIF  
ENDDEFINE
**************************************************************************************************************************
DEFINE CLASS combonew AS ComboBox
       BorderStyle=1      
       ForeColor=objForeColor 
       DisabledBackColor=dBackColor    
       SelectedBackColor=selBackColor  
       SelectedForeColor=selForeColor  
       SelectedItemBackColor=selBackColor
       ItemForeColor=objForeColor
       DisabledForeColor=objForeColor              
       
       FontName=dFontName
       FontSize=dFontSize       
       Visible=.T.
       Style=2
       BackStyle=0
       SpecialEffect=1
       varCtrlSource=''  
       ControlSource=This.varCtrlSource       
       lenPix=0
       chr32Max=0
       nDisplayCount=0
       procForGotFocus=''
       procForLostFocus=''
       procForDropDown='' 
       procForValid=''
       procForChange=''
       procForRightClick=''
       *--------------------------------------------------------------------------
       PROCEDURE Init
       This.BackColor=This.Parent.BackColor
       *--------------------------------------------------------------------------
       PROCEDURE GotFocus
       IF ! EMPTY(THIS.procForGotFocus)
          procForDo=THIS.procForGotFocus
          &procForDo
       ELSE
          IF This.nDisplayCount#0
	         This.DisplayCount=This.nDisplayCount
	      ENDIF
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE LostFocus
        IF ! EMPTY(THIS.procForLostFocus)
          procForDo=THIS.procForLostFocus
          &procForDo       
       ENDIF
       *--------------------------------------------------------------------------       
       PROCEDURE RightClick
        IF ! EMPTY(THIS.procForRightClick)
          procForDo=THIS.procForRightClick
          &procForDo       
       ENDIF
       *--------------------------------------------------------------------------
       PROCEDURE Valid
       IF ! EMPTY(THIS.procForValid)
          procForDo=THIS.procForValid
          &procForDo
       ENDIF
       *----------------------------------------------------------------------------
       PROCEDURE InterActiveChange
        IF ! EMPTY(THIS.procForChange)
           procDo=THIS.procForChange
           &procDo
       ENDIF
       
ENDDEFINE
*-----------------------------------------------------------------------
*         Назначение для Column.CurrentConrol объекта TextBox          
*-----------------------------------------------------------------------
PROCEDURE myColumnTxtBox
PARAMETERS parColumn,parTxtBox,parFname,parRead,parRightClick,parProc,parValid,parLostFocus
&parColumn..AddObject(parTxtBox,'myTxtBoxGrd')
&parColumn..CurrentControl=parTxtBox
WITH &parColumn..&parTxtBox
     .BackColor=IIF(&parColumn..Enabled=.T.,.BackColor,objBackColor)
     IF !EMPTY(parfname)
        .ControlSource=parFname     
     ENDIF  
     IF !EMPTY(parRead)
        .readOnly=parRead     
     ENDIF  
     IF !EMPTY(parRightClick)
        .procRightClick=parRightClick 
     ENDIF  
     IF !EMPTY(parProc)
        .procForKeyPress=parProc 
     ENDIF 
     IF !EMPTY(parValid)
        .procForValid=parValid
     ENDIF
     IF !EMPTY(parLostFocus)
        .procLostFocus=parLostFocus
     ENDIF
ENDWITH
*************************************************************************************************************************
*     Процедура "окультуривания" Grid (высота, ширина строк),Shape
*************************************************************************************************************************
PROCEDURE gridSize
PARAMETERS parFrm,parGrd,parObj,parLog,parHeight
* parFrm - имя формы
* parGrd - имя Grid
* parObj - имя Shape добавляемого, чтобы "оформить" правый край Grid
* parLog - устанавливает Enabled - .F. для колонок
* parHeight - устанавливает высоту Grid до конца формы
WITH &parFrm..&parGrd
     IF !parLog
        .SetAll('Enabled',.F.,'Column')
        .Columns(.ColumnCount).Enabled=.T.
        .Columns(.ColumnCount).ReadOnly=.T.
     ENDIF  
     .RowHeight=dRowHeight
     .HeaderHeight=.RowHeight
     .rowsGrid=(&parFrm..&parGrd..Height-&parFrm..&parGrd..HeaderHeight)/&parFrm..&parGrd..RowHeight
     IF !parHeight
        IF INT(.rowsGrid)-.rowsGrid#0
           .rowsGrid=IIF(ROUND(.RowsGrid,0)>.RowsGrid,INT(.rowsGrid),INT(.RowsGrid)-1) 
           .Height=(.RowsGrid*.rowHeight)+.HeaderHeight
        ENDIF      
     ENDIF    
ENDWITH  
IF !EMPTY(parObj)
   &parFrm..AddObject(parObj,'ShapeMy')
   WITH &parFrm..&parObj
        .SpecialEffect=1
        .Top=.Parent.&parGrd..Top+1
        .Left=.Parent.&parGrd..Left+.Parent.&parGrd..Width-SYSMETRIC(5)-3
        .BorderColor=RGB(255,255,255)        
        .Width=2
        .Height=.Parent.&parGrd..Height-2
        .Visible=.T.    
   ENDWITH
ENDIF
*************************************************************************************************************************
*     Процедура "окультуривания" Grid (высота, ширина строк),Shape - новый вариант
*************************************************************************************************************************
PROCEDURE gridSizeNew
PARAMETERS parFrm,parGrd,parObj,parLog,parHeight
* parFrm - имя формы
* parGrd - имя Grid
* parObj - имя Shape добавляемого, чтобы "оформить" правый край Grid
* parLog - устанавливает Enabled - .F. для колонок
* parHeight - устанавливает высоту Grid до конца формы
WITH &parFrm..&parGrd
     IF !parLog
        .SetAll('Enabled',.F.,'ColumnMy')
        .Columns(.ColumnCount).Enabled=.T.
        .Columns(.ColumnCount).ReadOnly=.T.
     ENDIF  
     .RowHeight=dRowHeight
     .HeaderHeight=.RowHeight
     .rowsGrid=(&parFrm..&parGrd..Height-&parFrm..&parGrd..HeaderHeight)/&parFrm..&parGrd..RowHeight
     IF !parHeight
        IF INT(.rowsGrid)-.rowsGrid#0
           .rowsGrid=IIF(ROUND(.RowsGrid,0)>.RowsGrid,INT(.rowsGrid),INT(.RowsGrid)-1) 
           .Height=(.RowsGrid*.rowHeight)+.HeaderHeight
        ENDIF      
     ENDIF    
ENDWITH  
IF !EMPTY(parObj)
   &parFrm..AddObject(parObj,'ShapeMy')
   WITH &parFrm..&parObj
        .SpecialEffect=1
        .Top=.Parent.&parGrd..Top+1
        .Left=.Parent.&parGrd..Left+.Parent.&parGrd..Width-SYSMETRIC(5)-3
        .BorderColor=RGB(255,255,255)        
        .Width=2
        .Height=.Parent.&parGrd..Height-2
        .Visible=.T.    
   ENDWITH
ENDIF
***********************************************************************************************
*
***********************************************************************************************
DEFINE CLASS MySpinner AS Spinner
       FontSize=dFontSize
       FontName=dFontName          
       BackColor=dBackColor
       ForeColor=objForeColor 
       SelectedBackColor=SelbackColor 
       SelectedForeColor=selForeColor
       DisabledBackColor=dBackColor  
       DisabledForeColor=objForeColor            
       SpecialEffect=1
       SpinnerLowValue=0
       KeyboardLowValue=0
       Visible=.T.
       procForValid=''
       procForKeyPress=''
       procForInterActiveChange=''
       *------------------------------------------------------------
       PROCEDURE GotFocus
       =SYS(2002,1)      
       *------------------------------------------------------------
       PROCEDURE LostFocus          
       *-------------------------------------------------------------
       PROCEDURE valid
       IF !EMPTY(This.procForValid)
           procDo=This.procForValid
           &procDo
       ENDIF
       *-------------------------------------------------------------
       PROCEDURE InterActiveChange
       IF !EMPTY(This.procForInterActiveChange)
           procDo=This.procForInterActiveChange
           &procDo
       ENDIF
       *--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl     
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF   
ENDDEFINE 
*************************************************************************************************************************
*    Процедура добавления Spinner в форму
*************************************************************************************************************************
PROCEDURE addSpinnerMy
PARAMETERS parFrm,parname,parLeft,parTop,parHeight,parWidth,parCtrlSource,parIncr,parProc,parLow,parHigh
obj_spin=parName
&parFrm..AddObject(obj_spin,'MySpinner')
WITH &parfrm..&obj_spin
     .BackColor=.Parent.BackColor 
     .DisabledBackColor=.Parent.BackColor  
     .SpecialEffect=1
     .Left=parLeft
     .Top=parTop
     .Width=parWidth
     .Height=parHeight
     .ControlSource=parCtrlSource 
     .Increment=IIF(!EMPTY(parIncr),parIncr,1)  
     .procForValid=IIF(!EMPTY(parProc),parProc,'')  
     IF !EMPTY(parLow)
        .KeyBoardLowValue=parLow 
        .SpinnerLowValue=parLow      
     ENDIF
     IF !EMPTY(parHigh)
        .KeyBoardHighValue=parHigh   
        .SpinnerHighValue=parHigh     
     ENDIF
ENDWITH
*************************************************************************************************************************
*              Процедура добавления Spinner в качестве CyrrentControl в колонку Grid
*************************************************************************************************************************
PROCEDURE addSpinnerGrd
PARAMETERS parFrm,parname,parCtrlSource,parIncr,parProc
obj_spin=parName
&parFrm..AddObject(obj_spin,'MySpinner')
WITH &parfrm..&obj_spin
     .SpecialEffect=2
     .BorderStyle=0     
     .ControlSource=parCtrlSource 
     .Increment=IIF(!EMPTY(parIncr),parIncr,1)  
     .procForValid=IIF(!EMPTY(parProc),parProc,'')  
ENDWITH
&parFrm..CurrentControl=parname
************************************************************************************************************************
*                       Процедура добавления объекта EditBox в форму для персонала
************************************************************************************************************************
PROCEDURE adeditbox
PARAMETERS parFrm,parname,parTop,parLeft,parWidth,parHeight,parCtrlSource,parRead,parAlign,parValid,parRightClick,parToolTipText,parKeyPress
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parwidth - щирина
* parheight - высота
* parCtrlSource - ControlSource
* parRead - ReadOnly
* parAlign - Выравнивание
* parValid - Процедура
* parRightClick - по правой клавише
* parToolTipText - TollTip
* parKeyPress
*obj_tbox='txtbox'+LTRIM(STR(parord))
obj_tbox=parname
&parfrm..AddObject(parname,'MyEditBox')
WITH &parfrm..&obj_tbox
     .SpecialEffect=1    
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=IIF(!EMPTY(parHeight),parHeight,dHeight)
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF        
    .Enabled=parRead    
     IF !EMPTY(parValid)   
        .procForLostFocus=parValid
     ENDIF
     IF !EMPTY(parAlign)
         .Alignment=parAlign
     ENDIF
     .toolTipText=IIF(!EMPTY(parToolTipText),parToolTipText,'')     
     IF !EMPTY(parKeyPress)    
        .procForKeyPress=parKeyPress     
     ENDIF
     *IF !EMPTY(parRightClick)   
     *   .procForRightClick="fpers.txtbox5.menread WITH 'fpers'"
      * .procForRightClick="DO menpeop WITH &parfrm,&parfrm..&obj_tbox,&parFrm..TextCopy,&parRightClick" 
     *ENDIF 
    * .DisabledBackColor=dBackColor
ENDWITH
************************************************************************************************************************
*                       Процедура добавления объекта TextBox в форму для персонала
************************************************************************************************************************
PROCEDURE adtboxnew
PARAMETERS parFrm,parname,parTop,parLeft,parWidth,parHeight,parCtrlSource,parFormat,parEnabled,parAlign,parMask,parValid,parRightClick,parToolTipText,parKeyPress
* parfrm - форма
* parname - имя эл-та
* parleft - левая граница
* partop - верхняя граница
* parwidth - щирина
* parheight - высота
* parCtrlSource - ControlSource
* parFormat - Формат
* parEnabled - Enabled
* parAlign - Выравнивание
* parMask - маска
* parValid - Процедура
* parRightClick - по правой клавише
* parToolTipText - TollTip
* parKeyPress
obj_tbox=parname
&parfrm..AddObject(obj_tbox,'MyTxtBox')
WITH &parfrm..&obj_tbox
     .SpecialEffect=1    
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=IIF(!EMPTY(parHeight),parHeight,dHeight)
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF 
     IF !EMPTY(parFormat)
         .Format=parFormat
     ENDIF     
    .Enabled=parEnabled    
     IF !EMPTY(parValid)   
        .procForValid=parValid
     ENDIF
     IF !EMPTY(parAlign)
         .Alignment=parAlign
     ENDIF
     IF !EMPTY(parMask)
        .InputMask=parMask
     ENDIF
     .toolTipText=IIF(!EMPTY(parToolTipText),parToolTipText,'')     
     IF !EMPTY(parKeyPress)    
        .procForKeyPress=parKeyPress     
     ENDIF
     *IF !EMPTY(parRightClick)   
     *   .procForRightClick="fpers.txtbox5.menread WITH 'fpers'"
      * .procForRightClick="DO menpeop WITH &parfrm,&parfrm..&obj_tbox,&parFrm..TextCopy,&parRightClick" 
     *ENDIF 
    * .DisabledBackColor=dBackColor
ENDWITH

************************************************************************************************************************
*                       Процедура добавления объекта TextBox в форму для персонала
************************************************************************************************************************
PROCEDURE adtbox
PARAMETERS parFrm,parord,parLeft,parTop,parWidth,parHeight,parCtrlSource,parFormat,parRead,parAlign,parValid,parRightClick,parToolTipText,parKeyPress
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parwidth - щирина
* parheight - высота
* parCtrlSource - ControlSource
* parFormat - Формат
* parRead - ReadOnly
* parAlign - Выравнивание
* parValid - Процедура
* parRightClick - по правой клавише
* parToolTipText - TollTip
* parKeyPress
obj_tbox='txtbox'+LTRIM(STR(parord))
&parfrm..AddObject(obj_tbox,'MyTxtBox')
WITH &parfrm..&obj_tbox
     .SpecialEffect=1    
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=IIF(!EMPTY(parHeight),parHeight,dHeight)
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF 
     IF !EMPTY(parFormat)
         .Format=parFormat
     ENDIF     
    .Enabled=parRead    
     IF !EMPTY(parValid)   
        .procForLostFocus=parValid
     ENDIF
     IF !EMPTY(parAlign)
         .Alignment=parAlign
     ENDIF
     .toolTipText=IIF(!EMPTY(parToolTipText),parToolTipText,'')     
     IF !EMPTY(parKeyPress)    
        .procForKeyPress=parKeyPress     
     ENDIF
     *IF !EMPTY(parRightClick)   
     *   .procForRightClick="fpers.txtbox5.menread WITH 'fpers'"
      * .procForRightClick="DO menpeop WITH &parfrm,&parfrm..&obj_tbox,&parFrm..TextCopy,&parRightClick" 
     *ENDIF 
    * .DisabledBackColor=dBackColor
ENDWITH

************************************************************************************************************************
*                       Процедура добавления объекта TextBox в форму
************************************************************************************************************************
PROCEDURE adtxtbox
PARAMETERS parFrm,parname,parLeft,parTop,parWidth,parHeight,parCtrlSource,parFormat,parRead,parAlign,parValid,parRightClick,parToolTipText
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parwidth - щирина
* parheight - высота
* parCtrlSource - ControlSource
* parFormat - Формат
* parRead - ReadOnly
* parAlign - Выравнивание
* parValid - Процедура
* parRightClick - по правой клавише
* parToolTipText - TollTip
*obj_tbox='txtbox'+LTRIM(STR(parord))
obj_tbox=parname
&parfrm..AddObject(obj_tbox,'MyTextBox')
WITH &parfrm..&obj_tbox
     .SpecialEffect=1    
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=IIF(!EMPTY(parHeight),parHeight,dHeight)
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF 
     IF !EMPTY(parFormat)
         .Format=parFormat
     ENDIF     
    .Enabled=parRead    
     IF !EMPTY(parValid)   
        .procForLostFocus=parValid
     ENDIF
     IF !EMPTY(parAlign)
         .Alignment=parAlign
     ENDIF
     .toolTipText=IIF(!EMPTY(parToolTipText),parToolTipText,'')  
     *IF !EMPTY(parRightClick)   
     *   .procForRightClick="fpers.txtbox5.menread WITH 'fpers'"
      * .procForRightClick="DO menpeop WITH &parfrm,&parfrm..&obj_tbox,&parFrm..TextCopy,&parRightClick" 
     *ENDIF 
    * .DisabledBackColor=dBackColor
ENDWITH

****************************************************************************************************************
*  Класс текстБокс для создания заголовков, надписей и т.д.
****************************************************************************************************************
DEFINE CLASS TextBoxAsCont AS TextBox
       FontName=dFontName
       FontSize=dFontSize
       ForeColor=objForeColor
       BackColor=ObjBackColor
       DisabledForeColor=ObjForeColor
       DisabledBackColor=HeaderBackColor     
       SelectOnEntry=.T.  
       BorderStyle=1
       *BackStyle=1
       Visible=.T.
       parlab=''
       procForRightClick=''
       procForDblClick=''
       procForClick=''
       procForKeyPress=''
       SpecialEffect=1   
       
       *------------------------------------------------------------------------------------------------
       PROCEDURE RightClick
       IF ! EMPTY(THIS.procForRightClick)
          procForDo=THIS.procForRightClick
          &procForDo
       ENDIF 
       *------------------------------------------------------------------------------------------------
       PROCEDURE Click
       IF ! EMPTY(THIS.procForClick)
          procForDo=THIS.procForClick
          &procForDo
       ENDIF 
       *--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl     
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF 
       *----------------------------------------------------------------------------------------------------
       PROCEDURE dblClick    
       IF ! EMPTY(THIS.procForDblClick)
          procForDo=THIS.procForDblClick
          &procForDo
       ENDIF 
ENDDEFINE
************************************************************************************************************************
*                  Процедура добавления объекта TextBox в форму для персонала (похож на контейнер)
************************************************************************************************************************
PROCEDURE adtboxascont
PARAMETERS parFrm,parname,parLeft,parTop,parWidth,parHeight,parCtrlSource,parAlign,parBackStyle,parBold
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parwidth - щирина
* parheight - высота
* parCtrlSource - ControlSource
* parAlign - Выравнивание
* parBackStyle=прозрачность

*obj_tbox='txtbox'+LTRIM(STR(parord))
obj_tbox=parname
&parfrm..AddObject(obj_tbox,'TextBoxAsCont')
WITH &parfrm..&obj_tbox
     .SpecialEffect=1    
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=IIF(!EMPTY(parHeight),parHeight,dHeight)     
     .Enabled=.F. 
     IF !EMPTY(parCtrlSource)
        .parlab=parCtrlSource         
        .ControlSource='&parfrm..&obj_tbox..parlab'       
     ENDIF
     IF parBold
        .FontBold=.T.
     ENDIF                    
     IF !EMPTY(parAlign)
         .Alignment=parAlign
     ENDIF     
     .DisabledBackColor=HeaderBackColor
     *IF !parBackStyle
        .BackStyle=parBackStyle
     *ELSE
        .BackStyle=parBackStyle
     *ENDIF   
ENDWITH

************************************************************************************************************************
*                       Процедура добавления объекта ComboBox в форму для персонала
************************************************************************************************************************
PROCEDURE adcbo
PARAMETERS parfrm,parord,parleft,partop,parheight,parwidth,parCtrl,parRowSource,parRowSourceType
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parheight - высота
* parwidth - щирина
 
obj_combo='combobox'+LTRIM(STR(parord))
&parfrm..AddObject(obj_combo,'ComboMy')
WITH &parfrm..&obj_combo
     .SpecialEffect=1   
     .Top=partop
     .Left=parleft
     .Width=parwidth
     .Height=parheight
     .ControlSource=parCtrl
     .RowSource=parRowSource
     .RowSourceType=parRowSourceType
ENDWITH
**********************************************************************************************
*                    Класс для создания командной кнопки
**********************************************************************************************
DEFINE CLASS mybutton AS CommandButton       
       Visible=.T.
       fontname=dFontName
       fontSize=dFontSize   
       Autosize=.F.     
       procForClick=''
       PROCEDURE Init
                 This.BackColor=This.Parent.BackColor   
       PROCEDURE Click
	             IF !EMPTY(This.ProcForClick)
                    ProcForDo=This.ProcForClick
                    &ProcForDo
                 ENDIF
    
ENDDEFINE
********************************************************************************************************************
DEFINE CLASS checkcontainer AS Container
       SpecialEffect=2
       Visible=.t.
       BackStyle=0
       BorderWidth=0
       BackColor=dBackColor
       ProcForClick='' 
       procForKeyPress=''
       logExit=.F.
       *ADD OBJECT checkMy AS checkBox WITH VISIBLE=.T., BACKSTYLE=0,;
       *			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize+1,;
       *			  Alignment=0,Caption=''  
*--------------------------------------------------------------------------
       PROCEDURE click
                 IF !EMPTY(This.ProcForClick)
                    ProcForDo=This.ProcForClick
                    &ProcForDo
                 ENDIF
*--------------------------------------------------------------------------
      * PROCEDURE checkMy.Click
      *           This.Parent.Click                
*-----------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE Keypress      
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl   
       IF This.logExit.AND.nKeyCode=27      
           ThisForm.Release
       ENDIF                                           
       IF !EMPTY(This.procForKeyPress)
          procdo=This.procForKeyPress
          &procdo   
       ENDIF       		     
ENDDEFINE
********************************************************************************************************************
PROCEDURE addCheckContainerInColumn
PARAMETERS parFrm,parcont,parname
*parFrm - форма
&parFrm..AddObject('parcont','checkcontainer')

*&parFrm..AddObject(parname,'checkMy')
&parFrm..CurrentControl=parcont
*WITH &parFrm..&parName
*     .SpecialEffect=0 
*     .Caption=''
*     .Width=18
*     .Left=5
*ENDWITH

********************************************************************************************************************
DEFINE CLASS contmy AS Container
       SpecialEffect=0
       Visible=.t.
       BackStyle=0
       BackColor=dBackColor
       ProcForClick='' 
       ProcMouseEnter=''
       ProcMouseLeave=''
       NameForm=''     
                             
       PROCEDURE click
                 IF !EMPTY(This.ProcForClick)
                    ProcForDo=This.ProcForClick
                    &ProcForDo
                 ENDIF
       PROCEDURE MouseEnter
                 LPARAMETERS nButton, nShift, nXCoord, nYCoord  
                 IF !EMPTY(This.procMouseEnter)
                    ProcForDo=This.ProcMouseEnter
                    &ProcForDo 
                 ENDIF                 
       PROCEDURE MouseLeave 
                 LPARAMETERS nButton, nShift, nXCoord, nYCoord                        
                 IF !EMPTY(This.procMouseLeave)
                    ProcForDo=This.ProcMouseLeave
                    &ProcForDo 
                 ENDIF                        
ENDDEFINE
********************************************************************************************************************
DEFINE CLASS contHead AS Container
       SpecialEffect=0
       Visible=.t.
       BackStyle=1
       BackColor=HeaderBackColor
       ProcForClick='' 
       ProcForDblClick=''
       procForRightClick=''
       ProcMouseEnter=''
       ProcMouseLeave=''
       NameForm=''
       
       ADD OBJECT ContLabel AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize+1,;
       			  Alignment=2     
*--------------------------------------------------------------------------
       PROCEDURE click
                 IF !EMPTY(This.ProcForClick)
                    ProcForDo=This.ProcForClick
                    &ProcForDo
                 ENDIF
*--------------------------------------------------------------------------                
       PROCEDURE dblclick
                 IF !EMPTY(This.ProcForDblClick)
                    ProcForDo=This.ProcForDblClick
                    &ProcForDo
                 ENDIF  
*--------------------------------------------------------------------------
       PROCEDURE Rightclick
                 IF !EMPTY(This.ProcForRightClick)
                    ProcForDo=This.ProcForRightClick
                    &ProcForDo
                 ENDIF                         
*--------------------------------------------------------------------------
       PROCEDURE ContLabel.Click
                 This.Parent.Click  
*--------------------------------------------------------------------------  
       PROCEDURE ContLabel.DblClick
                 This.Parent.DblClick  
*--------------------------------------------------------------------------  
       PROCEDURE ContLabel.RightClick
                 This.Parent.RightClick                  
*--------------------------------------------------------------------------         
       PROCEDURE MouseEnter
                 LPARAMETERS nButton, nShift, nXCoord, nYCoord  
                 IF !EMPTY(This.procMouseEnter)
                    ProcForDo=This.ProcMouseEnter
                    &ProcForDo 
                 ENDIF                 
       PROCEDURE MouseLeave 
                 LPARAMETERS nButton, nShift, nXCoord, nYCoord                        
                 IF !EMPTY(This.procMouseLeave)
                    ProcForDo=This.ProcMouseLeave
                    &ProcForDo 
                 ENDIF                        
ENDDEFINE

********************************************************************************************************************
DEFINE CLASS contHeadNew AS Container
       SpecialEffect=0
       Visible=.t.
       BackStyle=1
       BackColor=HeaderBackColor
       ProcForClick='' 
       ProcForDblClick=''
       procForRightClick=''
       ProcMouseEnter=''
       ProcMouseLeave=''
       NameForm=''
       
       ADD OBJECT ContLabel AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize+1,;
       			  Alignment=2     
*--------------------------------------------------------------------------
       PROCEDURE click
                 IF !EMPTY(This.ProcForClick)
                    ProcForDo=This.ProcForClick
                    &ProcForDo
                 ENDIF
*--------------------------------------------------------------------------                
       PROCEDURE dblclick
                 IF !EMPTY(This.ProcForDblClick)
                    ProcForDo=This.ProcForDblClick
                    &ProcForDo
                 ENDIF  
*--------------------------------------------------------------------------
       PROCEDURE Rightclick
                 IF !EMPTY(This.ProcForRightClick)
                    ProcForDo=This.ProcForRightClick
                    &ProcForDo
                 ENDIF                         
*--------------------------------------------------------------------------
       PROCEDURE ContLabel.Click
                 This.Parent.Click  
*--------------------------------------------------------------------------  
       PROCEDURE ContLabel.DblClick
                 This.Parent.DblClick  
*--------------------------------------------------------------------------  
       PROCEDURE ContLabel.RightClick
                 This.Parent.RightClick                  
*--------------------------------------------------------------------------         
       PROCEDURE MouseEnter
                 LPARAMETERS nButton, nShift, nXCoord, nYCoord  
                 IF !EMPTY(This.procMouseEnter)
                    ProcForDo=This.ProcMouseEnter
                    &ProcForDo 
                 ENDIF                 
       PROCEDURE MouseLeave 
                 LPARAMETERS nButton, nShift, nXCoord, nYCoord                        
                 IF !EMPTY(This.procMouseLeave)
                    ProcForDo=This.ProcMouseLeave
                    &ProcForDo 
                 ENDIF                        
ENDDEFINE




*************************************************************************************************************************
*                 Клаcc для создания EditBox
*************************************************************************************************************************
DEFINE CLASS MyEditBox AS EditBox
       BackStyle=0       
       Visible=.T.
       SpecialEffect=1      
       ForeColor=objForeColor
       BackColor=ObjBackColor
       DisabledForeColor=ObjForeColor
       DisabledBackColor=ObjBackColor 
       FontName=dFontName
       FontSize=dFontSize
       FontSize=dFontSize    
       ForeColor=dForeColor                 
       ProcGotFocus=''
       ProcLostFocus=''
       procForRightClick=''
       procForDblClick=''
       ProcValid=''     
       *------------------------------------------------------------------------------------------------------------------ 
       PROCEDURE GotFocus
       SET CURSOR ON 
              IF !EMPTY(This.ProcGotFocus)
          procDo=This.ProcGotFocus
          &procDo
       ENDIF
       *------------------------------------------------------------------------------------------------------------------ 
       PROCEDURE LostFocus
       SET CURSOR OFF       
       IF !EMPTY(This.ProcLostFocus)
          procDo=This.ProcLostFocus
          &procDo
       ENDIF  
       *------------------------------------------------------------------------------------------------------------------ 
       PROCEDURE Valid
       IF !EMPTY(This.ProcValid)
          procDo=This.ProcValid
          &procDo
       ENDIF   
       *------------------------------------------------------------------------------------------------------------------      
       PROCEDURE RightClick
       IF !EMPTY(This.ProcForRightClick)
          procDo=This.ProcForRightClick
          &procDo
       ENDIF    
       *------------------------------------------------------------------------------------------------------------------             
       PROCEDURE DblClick
       IF !EMPTY(This.ProcForDblClick)
          procDo=This.ProcForDblClick
          &procDo
       ENDIF      
ENDDEFINE
************************************************************************************************************************
*
************************************************************************************************************************
DEFINE CLASS checkmy AS CheckBox
       procInterActiveChange=''
       procValid=''
       Visible=.T.
       
       PROCEDURE InterActiveChange
                 IF !EMPTY(This.procInterActiveChange)
                    procdo=This.procInterActiveChange
                    &procdo  
                 ENDIF   
       PROCEDURE Valid
                 IF !EMPTY(This.procValid)
                    procdo=This.procValid
                    &procdo  
                 ENDIF             
ENDDEFINE
*----------------------------------------------------------------------------------------------
*   Назначение для Column.CurrentCongtrol объекта CheckBox
*----------------------------------------------------------------------------------------------
PROCEDURE ColumnCheckBox
PARAMETERS parcolumn,parCheckBox,parValid
&parColumn..AddObject(parCheckBox,'CheckMy')
&parColumn..CurrentControl=parCheckBox
WITH &parColumn..&parCheckBox
     .Caption=''
     .Autosize=.T.
     .Alignment=1
     .SpecialEffect=1        
     .procValid=IIF(!EMPTY(parvalid),parValid,'')      
*     .ControlSource=parFname
ENDWITH
*********************************************************************************************************************
*              Процедура добавления контейнера в качестве заголовка столбца Grid
*****************************************************************************************************
PROCEDURE addcontmy
PARAMETERS parfrm,parname,parleft,partop,parwidth,parheight,parLab,parproc,parBackStyle
obj_cont=parname
&parFrm..AddObject(obj_cont,'contHead')
WITH &parFrm..&obj_cont
     .Left=parleft
     .Top=partop
     .Width=parwidth
     .Height=parheight    
     .ContLabel.Caption=parLab
     .Contlabel.Autosize=.T.
     .Contlabel.Visible=.T.     
     .ContLabel.Left=(.Width-.ContLabel.Width)/2     
     .ContLabel.Top=(.Height-FONTMETRIC(1,dFontName,.Contlabel.FontSize))/2
     .ContLabel.Visible=.T.
     .BackStyle=IIF(EMPTY(parBackStyle),0,parBackStyle)
     IF !EMPTY(parproc)
        .procForClick=parproc 
     ENDIF     
ENDWITH

*********************************************************************************************************************
*              Процедура добавления контейнера в качестве заголовка столбца Grid
*****************************************************************************************************
PROCEDURE addcontform
PARAMETERS parfrm,parname,parleft,partop,parwidth,parheight,parLab,parproc
obj_cont=parname
&parFrm..AddObject(obj_cont,'contHead')
WITH &parFrm..&obj_cont
     .Left=parleft
     .Top=partop
     .Width=parwidth
     .Height=parheight  
     .SpecialEffect=2  
     .ContLabel.Caption=parLab
     .ContLabel.FontSize=dFontSize
     .Contlabel.Autosize=.T.
     .Contlabel.Visible=.T.     
     .ContLabel.Left=(.Width-.ContLabel.Width)/2     
     .ContLabel.Top=(.Height-FONTMETRIC(1,dFontName,.Contlabel.FontSize))/2
     .ContLabel.Visible=.T.
     IF !EMPTY(parproc)
        .procForClick=parproc 
     ENDIF     
ENDWITH
*********************************************************************************************************************
*              Процедура добавления контейнера в качестве заголовка столбца Grid
*****************************************************************************************************
PROCEDURE addcontformnew
PARAMETERS parfrm,parname,parleft,partop,parwidth,parheight,parLab,parlabAl,parproc,parDblClick,parRightClick,parBold
obj_cont=parname
&parFrm..AddObject(obj_cont,'contHeadNew')
WITH &parFrm..&obj_cont
     .Left=parleft
     .Top=partop
     .Width=parwidth
     .Height=parheight  
     .SpecialEffect=2  
     .ContLabel.Width=parWidth
     .ContLabel.Caption=parLab
     .ContLabel.FontSize=dFontSize
     IF parBold
        .ContLabel.FontBold=.T.
     ENDIF 
     .Contlabel.Autosize=.F.
     .Contlabel.Visible=.T.     
     .ContLabel.Left=2     
     .ContLabel.Top=(.Height-FONTMETRIC(1,dFontName,.Contlabel.FontSize))/2
     IF EMPTY(parlabal)
        .Contlabel.Alignment=parlabAl
     ENDIF
     .ContLabel.Visible=.T.
     IF !EMPTY(parproc)
        .procForClick=parproc 
     ENDIF    
     IF !EMPTY(parDblClick)
        .procForDblClick=parDblClick 
     ENDIF  
     IF !EMPTY(parRightClick)
        .procForRightClick=parRightClick 
     ENDIF  
ENDWITH
*************************************************************************************************************************
*         Процедура добавления объекта "Label" к форме (новый вариант)
*************************************************************************************************************************
PROCEDURE adlabmy
PARAMETERS parFrm,parOrd,parCaption,parTop,parLeft,parWidth,parAlignment,parAutoSize,parBackStyle
&& parFrm - форма
&& parord - номер эл-та
&& parCaption - зоголовок
&& parTop -верх
&& parLeft -лево
&& parWidth -ширина
&& parAlignment -выравнивание
&& parAutoSize - авторазмер
&& parBackStyle - BackStyle
obj_lab='lab'+LTRIM(STR(parord))
&parFrm..ADDOBJECT(obj_lab,'labelmy')
WITH &parFrm..&obj_lab
     .Caption=parCaption
     .Top=parTop
     .Left=parLeft
     .Height=dHeight
     .Width=parWidth
     .Alignment=IIF(!EMPTY(parAlignment),parAlignment,0)
     .AutoSize=IIF(parAutoSize,.T.,.F.)
     .BackStyle=IIF(!EMPTY(parBackStyle),parBackStyle,0)
     .Visible=.T.
ENDWITH
*************************************************************************************************************************
*         Процедура добавления объекта "Label" к форме
*************************************************************************************************************************
PROCEDURE addlabelmy
PARAMETERS parFrm,parord,parCaption,parLeft,parTop,parHeight,parWidth,parAlignment
obj_lab='lab'+LTRIM(STR(parord))
&parFrm..ADDOBJECT(obj_lab,'labelmy')
WITH &parFrm..&obj_lab
     .Caption=parCaption
     .Top=parTop
     .Left=parLeft
     .Height=parHeight
     .Width=parWidth
     .Alignment=IIF(!EMPTY(parAlignment),parAlignment,0)
ENDWITH
************************************************************************************************************************
*                       Процедура добавления объекта TextBox в форму
************************************************************************************************************************
PROCEDURE addtxtboxmy
PARAMETERS parFrm,parOrd,parLeft,parTop,parWidth,parHeight,parCtrlSource,parAlign,parValid,parFormat
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parheight - высота
* parwidth - щирина
* parctrlsource - ControlSource 
* parAlign - Выравнивание (лево-0,право-1,центр-2)
* parValid - procForValid
* parFormat - Формат 
obj_tbox='txtbox'+LTRIM(STR(parOrd))
&parfrm..AddObject(obj_tbox,'MyTxtBox')
WITH &parfrm..&obj_tbox
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=IIF(!EMPTY(parHeight),parHeight,dHeight)  
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF 
     .Alignment=IIF(!EMPTY(parAlign),parAlign,0) 
     IF !EMPTY(parFormat)
        .Format=parFormat
     ENDIF
     IF !EMPTY(parValid)
        .procForValid=parValid
     ENDIF
    * .DisabledBackColor=dBackColor
ENDWITH

************************************************************************************************************************
*                       Процедура добавления объекта ComboBox в форму
************************************************************************************************************************
PROCEDURE addcombomy
PARAMETERS parfrm,parord,parleft,partop,parheight,parwidth,parenabled,parCtrlSource,parRowSource,parSourceType,parGotFocus,parValid,parStyle,parVisible
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parheight - высота
* parwidth - щирина
* parenabled - enabled
* parCtrlSource - ControlSource
* parRowSource  - RowSource
* parSourceType - RowSourceType
* parGotFocus - GotFocus
* parValid   - Valid
* parStyle Style
* parVisible - Visible
obj_combo='combobox'+LTRIM(STR(parord))
&parfrm..AddObject(obj_combo,'ComboMy')
WITH &parfrm..&obj_combo
     .Top=partop
     .Left=parleft
     .Width=parwidth
     .Height=parheight
     .Enabled=parenabled    
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF   
     IF !EMPTY(parRowSource)
       .RowSource=parRowSource
     ENDIF
     IF !EMPTY(parSourceType)
        .RowSourceType=parSourceType
     ENDIF   
     IF !EMPTY(parGotFocus)
        .procForGotFocus=parGotFocus
     ENDIF
     IF !EMPTY(parValid)
        .procForValid=parValid
     ENDIF
     IF !EMPTY(parStyle)
         .Style=parStyle
     ENDIF
     .Visible=IIF(parVisible,.T.,.F.)
     
        
ENDWITH
*************************************************************************************************************************
*                          Класс Image 
*************************************************************************************************************************
DEFINE CLASS myimage AS Image 
       Stretch=1
       BorderStyle=0
       BackStyle=0
       Visible=.T.
       ProcForClick=''
       procForDoublClick=''
       procForRightClick=''
       **********************************************
       PROCEDURE Click
       IF !EMPTY(This.ProcForClick)
           ProcDo=This.ProcForClick
           &ProcDo           
       ENDIF 
       ***********************************************
       PROCEDURE dblClick      
       IF !EMPTY(This.ProcForDoublClick)
           ProcDo=This.ProcForDoublClick
           &ProcDo           
       ENDIF 
        ***********************************************
       PROCEDURE RightClick      
       IF !EMPTY(This.ProcForRightClick)
           ProcDo=This.ProcForRightClick
           &ProcDo           
       ENDIF 
ENDDEFINE
*************************************************************************************************************************
*         Класс для сохдания Shape
*************************************************************************************************************************
DEFINE CLASS shapemy AS Shape
       BackStyle=0
       BackColor=dBackColor
       SpecialEffect=0
       Visible=.T.
       procForClick=''
       procForDoublClick=''
       **********************************************
       PROCEDURE Click
       IF !EMPTY(This.ProcForClick)
           ProcDo=This.ProcForClick
           &ProcDo           
       ENDIF 
       ***********************************************
       PROCEDURE dblClick      
       IF !EMPTY(This.ProcForDoublClick)
           ProcDo=This.ProcForDoublClick
           &ProcDo           
       ENDIF        
ENDDEFINE 
************************************************************************************************************************
PROCEDURE addshapeingrid
PARAMETERS parfrm,parobj,pargrd
&parfrm..AddObject(parobj,'ShapeMy')
WITH &parfrm..&parobj
     .Top=.Parent.&parGrd..Top+1
     .Left=.Parent.&parGrd..Left+.Parent.&parGrd..Width-SYSMETRIC(5)-3
     .BorderColor=RGB(255,255,255)        
     .Width=2
     .Height=.Parent.&parGrd..Height-2
     .Visible=.T.    
ENDWITH

************************************************************************************************************************
*                       Процедура добавления объекта Shape в форму
************************************************************************************************************************
PROCEDURE addshape
PARAMETERS parfrm,parord,parleft,partop,parheight,parwidth,parcurv
* parfrm - форма
* parord - номер эл-та
* parleft - левая граница
* partop - верхняя граница
* parheight - высота
* parwidth - щирина
* parcurv - кривизна 
obj_shape='Shape'+LTRIM(STR(parord))
&parfrm..AddObject(obj_shape,'ShapeMy')
WITH &parfrm..&obj_Shape
     .Top=partop
     .Left=parleft
     .Width=parwidth
     .Height=parheight
      IF !EMPTY(parcurv)
          .Curvature=8
          .BorderColor=RGB(182,192,192)
      ENDIF 
     .Visible=.T.
ENDWITH
*************************************************************************************************************************
*                                Класс  для OptionButton
*************************************************************************************************************************
DEFINE CLASS myOptionButton AS OptionButton     
       ForeColor=ObjForeColor
       procInterActiveChange='' 
       procForValid=''
       procForKeyPress=''
       logExit=.F.
       FontSize=dFontSize
       FontName=dFontName
       Visible=.T.
       *----------------------------------------------------
       PROCEDURE Init
       This.BackColor=This.Parent.BackColor
       *----------------------------------------------------
       PROCEDURE InterActiveChange
       IF !EMPTY(This.procInterActiveChange)
           procDo=This.procInterActiveChange
           &procDo
       ENDIF 
       *----------------------------------------------------
       PROCEDURE Valid
       IF !EMPTY(This.procForValid)
           procDo=This.procForValid
           &procDo
       ENDIF       
       *-----------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE Keypress      
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl   
       IF This.logExit.AND.nKeyCode=27      
           ThisForm.Release
       ENDIF                                           
       IF !EMPTY(This.procForKeyPress)
          procdo=This.procForKeyPress
          &procdo   
       ENDIF       
ENDDEFINE
*************************************************************************************************************************
*         Процедура добавления OptionButton в форму
*************************************************************************************************************************
PROCEDURE addOptionButton
PARAMETERS parFrm,parOrd,parCaption,parTop,parLeft,parCtrlSource,parAlign,parValid,parEnabled
* parfrm - форма
* parord - номер эл-та
* parCaption - Заголовок
* parleft - левая граница
* partop - верхняя граница
* parctrlsource - ControlSource 
* parAlign - Выравнивание (лево-0,право-1,центр-2)
* parValid - ProcForValid
* parEnabled - Enabled
obj_Option='Option'+LTRIM(STR(parOrd))
&parfrm..AddObject(obj_Option,'MyOptionButton')
WITH &parfrm..&obj_Option
     .Caption=parCaption
     .Top=parTop
     .Left=parLeft        
     IF !EMPTY(parCtrlSource)
        .ControlSource=parCtrlSource
     ENDIF 
     IF !EMPTY(parValid)
        .procForValid=parValid 
     ENDIF
     .Alignment=IIF(!EMPTY(parAlign),parAlign,0) 
     .AutoSize=.T. 
     .Enabled=parEnabled
ENDWITH
*************************************************************************************************************************
*                                Класс  для CheckBox
*************************************************************************************************************************
DEFINE CLASS myCheckBox AS CheckBox          
       FontSize=dFontSize
       FontName=dFontName
       ForeColor=ObjForeColor
       AutoSize=.T.
       SpecialEffect=0
       procForValid=''
       procForKeyPress=''
       logExit=.F.
       Visible=.T. 
*---------------------------------------------------------------------------------------------------------------------       
       PROCEDURE Init
       This.BackColor=This.Parent.BackColor  
       This.DisabledBackColor=This.Parent.BackColor  
*----------------------------------------------------
       PROCEDURE Valid
       IF !EMPTY(This.procForValid)
           procDo=This.procForValid
           &procDo
       ENDIF 
*------------------------------------------------------
       PROCEDURE Keypress      
       LPARAMETERS nIndex,nKeyCode,nShiftAltCtrl                                                   
       IF This.logexit.AND.nKeyCode=27
           ThisForm.Release
       ENDIF 
       IF !EMPTY(This.procForKeyPress)
          procdo=This.procForKeyPress
          &procdo   
       ENDIF                        
ENDDEFINE
***********************************************************************************************************************
*                     Процедура добавления в форму объекта CheckBox
***********************************************************************************************************************
PROCEDURE AdCheckBox
PARAMETERS parfrm,parCheckBox,parCaption,parTop,parLeft,parWidth,parHeight,parCtrlSource,parAlignment,parAutoSize,parValid
* parfrm       - Форма
* parCheckBox  - Имя объекта
* parCaptiom   - заголовок
* parTop       - Верх
* parLeft      - Лево
* parWidth     - Ширина
* parHeight    - Высота
* parAlignment - выравнивание 
* parCtrlSource- ControlSource
* parAutoSize  - Авторазмер
* parValid     - Процедура
&parfrm..AddObject(parCheckBox,'MyCheckBox')
WITH &parfrm..&parCheckBox
     .Caption=parCaption
     .Top=parTop
     .Left=parLeft   
     .Width=parWidth
     .Height=parHeight
     .Alignment=parAlignment   
     .ControlSource=parCtrlSource
     IF !EMPTY(parAutoSize)
        .Autosize=parAutoSize 
      ENDIF
     .Visible=.T. 
     .procForValid=IIF(!EMPTY(parvalid),parValid,'')          
ENDWITH
*************************************************************************************************************************
*               Класс Control для "кнопок" меню
*************************************************************************************************************************
DEFINE CLASS Mycontmenu AS Container
       Height=36
       Visible=.T.
       SpecialEffect=0
       BackStyle=0
       BorderWidth=0  
       ProcForclick=''
       ProcMouseEnter=''
       ProcMouseLeave=''
       PROCEDURE CLICK                   
       IF !EMPTY(This.ProcForClick)                 
          This.BorderWidth=0     
          procdo=This.ProcForClick 
          &procdo
       ENDIF                           
       
       PROCEDURE MouseEnter
       LPARAMETERS nButton, nShift, nXCoord, nYCoord  
       This.BorderWidth=1
       IF !EMPTY(This.ProcMouseEnter)
           procdo=This.ProcMouseEnter
           &procdo 
       ENDIF
       
       PROCEDURE MouseLeave 
       LPARAMETERS nButton, nShift, nXCoord, nYCoord         
       This.BorderWidth=0
       IF !EMPTY(This.ProcMouseLeave)
          procdo=This.ProcMouseLeave
          &procdo 
       ENDIF                               
ENDDEFINE
******************************************************************************************
*   		          Процедура добавление Image к форме
******************************************************************************************
PROCEDURE addImage
LPARAMETERS parFrm,parInd,parTop,parLeft,parHeight,parWidth,parico,parForClick
name_obj='image'+LTRIM(STR(parInd))
&parFrm..ADDOBJECT(name_obj,'myImage')
WITH &parFrm..&name_obj
     .Top=parTop
     .Left=parLeft
     .Height=parHeight
     .Width=parWidth
     .Picture=parico
     .procForClick=parForClick
ENDWITH
*************************************************************************************************************************
*                      Процедура добавления control в форму
*************************************************************************************************************************
PROCEDURE addmycontmenu
PARAMETERS parfrm,parname,partop,parleft,parwidth,parpict,parlab,parproc,parico,partooltiptext
objcontmenu=parname
&parfrm..AddObject(objcontmenu,'mycontmenu')  
&parfrm..&objcontmenu..AddObject('npict','MyImage')
WITH &parfrm..&objcontmenu..npict       
     .Picture=parico   
     .Height=32
     .Width=32
     .Top=(&parfrm..&objcontmenu..Height-&parfrm..&objcontmenu..npict.Height)/2
     .Left=5
     .ProcForClick=parproc
     .ToolTipText=IIF(!EMPTY(partooltiptext),partooltiptext,'')
ENDWITH
&parfrm..&objcontmenu..AddObject('label1','Labelmy')
WITH &parfrm..&objcontmenu..label1     
     .LEFT=&parfrm..&objcontmenu..npict.Left+&parfrm..&objcontmenu..npict.Width+5
     .Top=(&parfrm..&objcontmenu..Height-&parfrm..&objcontmenu..label1.Height)/2
     .Caption=parlab
     .ProcForClick=parproc
     .Width=RetTxtWidth(parlab,dFontName,dFontSize) 
     .ToolTipText=IIF(!EMPTY(partooltiptext),partooltiptext,'')   
ENDWITH
WITH &parFrm..&objcontmenu
     .Top=partop
     .Left=parleft     
     *.Width=&parFrm..Objcontmenu..label1.Width+&parFrm..Objcontmenu..npict.Width+6
     .Width=&parfrm..&objcontmenu..label1.Width+&parfrm..&objcontmenu..npict.Width+15
ENDWITH
*************************************************************************************************************************
*            Процедура добавления объекта Grid к форме
*************************************************************************************************************************
PROCEDURE addGridmy
PARAMETERS parFrm,parOrd,parColumn,parTop,parLeft,parHeight,parWidth,parScroll
obj_grid='Grid'+LTRIM(STR(parord))
&parFrm..ADDOBJECT(obj_grid,'gridmy')
WITH &parFrm..&Obj_grid
     .ColumnCount=parColumn
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=parHeight
     .ScrollBars=parScroll
ENDWITH
********************************************************************************************
*                              Процедура переспроса
********************************************************************************************   
PROCEDURE yes_no 
PARAMETERS par1,par2,par3,par4
pfrm=NEWOBJECT('yesnoform')
pfrm.Show()
DEFINE CLASS yesnoform AS Form
       WINDOWTYPE=1
       Height=90
       Width=220      
       DoCreate=.T.
       BorderStyle=2
       Caption=par1
       MaxButton=.F.
       Minbutton=.F.      
       ShowTips=.T.
       VISIBLE =.T.  
       PROCEDURE Activate
       ENDPROC       
       PROCEDURE Init                             
                 IF !EMPTY(par3).AND.!EMPTY(par4)
                     This.Top=par3
                     This.Left=par4
                  ELSE
                     This.AutoCenter=.T.
                  ENDIF      
                  ThisForm.AddObject('labyes','Label')
                  ThisForm.AddObject('but1','Comyes')
                  ThisForm.AddObject('but2','Comno')                                        
                  WITH Thisform.labyes
                       .Caption=par2
                       .Alignment=2
                       .Left=4
                       .Top=9         
                       .Width=211             
                       .Visible=.T.
                  ENDWITH                                                                                   
       ENDPROC
       
       
       PROCEDURE keyPress       
       LPARAMETERS nKeyCode, nShiftAltCtrl
                 IF nkeyCode=27 
                    prm_ch=0
                    ThisForm.Release()                   
                 ENDIF
       ENDPROC      
       PROCEDURE Destroy               
                
       ENDPROC
                 
ENDDEFINE

DEFINE CLASS comyes AS CommandButton
       Caption='Да'
       Left=8
       Top=48                      
       Visible=.T.
       Width=100
       Height=30
       
       PROCEDURE click
                 prm_ch=1
                 log_act=.T.
                 ThisForm.RELEASE()
       ENDPROC
ENDDEFINE

DEFINE CLASS comno AS CommandButton
       Caption='Нет'
       Left=111
       Top=48                      
       Visible=.T.
       width=100
       Height=30
       
       PROCEDURE click
                 prm_ch=2
                 log_act=.T.                
                 ThisForm.RELEASE()
       ENDPROC
ENDDEFINE  
**************************************************************************************************************************
*           Вспомогательная процедура по организации окна подтверждения
**************************************************************************************************************************
PROCEDURE procok
PARAMETERS par1,par2,par3,par4
okch=NEWOBJECT('okform')
okch.Show()
DEFINE CLASS okform AS formmy     
       PROCEDURE Destroy
                 
       ENDPROC
       PROCEDURE Init                                
                 This.Height=70  
                 This.AddObject('lab1','label')
                 This.AddObject('combut','combutcl')    
                 This.Caption=IIF(!EMPTY(par4),par4,'Сообщение')
                 This.lab1.Visible=.T.
                 This.lab1.Autosize=.T.
                 This.lab1.Caption=par3
                 This.lab1.Top=5
                 This.lab1.left=20                 
                 This.Width=40+This.lab1.Width
                 widthform=This.Width
                 This.combut.Left=(widthform-100)/2
                 IF EMPTY(par1).OR.EMPTY(par2)
                    ThisForm.AutoCenter=.T.
                 ELSE
                    ThisForm.AutoCenter=.F.
                    This.Left=par1
                    This.Top=par2                  
                 ENDIF
                 
       ENDPROC
       PROCEDURE keyPress       
       LPARAMETERS nKeyCode, nShiftAltCtrl
                   IF nkeyCode=27 
                      prm_ch=0
                      ThisForm.Release()                   
                   ENDIF
       ENDPROC       
ENDDEFINE

DEFINE CLASS combutcl AS CommandButton
       Visible=.T.
       *Left=5
       Top=30
       Height=25
       Width=100
       Caption='OK'
       PROCEDURE click
                 _screen.ActiveForm.Release()
       ENDPROC
ENDDEFINE
*****************************************************************************************************
*                               Процедура построения формы при проверке номера на совпадение
**********************************************************************************************************
PROCEDURE numread
PARAMETERS par1
nfrm=NEWOBJECT('numform')
nfrm.Show()
DEFINE CLASS numform AS Form
       WINDOWTYPE=1
       ShowWindow=1
       Height=69
       Width=334      
       DoCreate=.T.
       BorderStyle=2
       Caption='Перенумерация'
       MaxButton=.F.
       Minbutton=.F.      
       ShowTips=.T.
       VISIBLE =.T.  
       PROCEDURE Activate
       ENDPROC       
       PROCEDURE Init   
                 *IF !EMPTY(par3).AND.!EMPTY(par4)
                 *   This.Top=par3
                 *   This.Left=par4
                 *ELSE
                    This.AutoCenter=.T.
                 *ENDIF      
                 ThisForm.AddObject('labyes','Label')
                 ThisForm.AddObject('comgr','Comgrnum')                                                       
                 WITH Thisform.labyes
                      .Caption='Внимание! Данный номер уже занят!' 
                      .Top=5   
                      .Left=0
                      .Width=ThisForm.Width
                      .Alignment=2 
                      *.Autosize=.T.                                         
                      .Visible=.T. 
                                
                 ENDWITH                                                                                   
       ENDPROC
       
       
       PROCEDURE keyPress       
       LPARAMETERS nKeyCode, nShiftAltCtrl
                  IF nkeyCode=27 
                    prm_ch=0
                    ThisForm.Release()                   
                 ENDIF
       ENDPROC      
       PROCEDURE Destroy               
                
       ENDPROC       
                 
ENDDEFINE

DEFINE CLASS Comgrnum AS CommandGroup
       Top=24       
       Left=10
       *Height=37
       ButtonCount=3
       Visible=.T.
       
       Command1.Left=2
       Command2.Left=106
       Command3.Left=208           
       Command1.Caption='Перенумерация' 
       Command2.Caption='Возврат' 
       Command3.Caption='Отказ'             
       Command1.ToolTipText='Перенумеровать должности' 
       Command2.ToolTipText='Вернуться к редактированию номера' 
       Command3.ToolTipText='Отказаться от редактирования'       
       PROCEDURE Init
                 This.SetAll("Width",102,"CommandButton")
                 This.SetAll("Height",30,"CommandButton")
                 This.SetAll("Top",5,"CommandButton")
                 This.SetAll("FontName",'Courier New',"Commandbutton")
                 This.AutoSize=.T.
       ENDPROC
       PROCEDURE Click
                 DO CASE
                    CASE This.Value=1
                         prm_ch=1
                    CASE This.Value=2
                         prm_ch=2
                    CASE This.Value=3
                         prm_ch=0 
                 ENDCASE
                 ThisForm.Release()
       ENDPROC
ENDDEFINE
********************************************************************************************************************
DEFINE CLASS MyCommandGroup AS CommandGroup
       procforClick=''
       Visible=.T.
       Autosize=.T.
       BorderStyle=0   
       BackStyle=0   
       procForValid=''
       PROCEDURE Init
                 backColor=This.Parent.BackColor
       PROCEDURE Valid
                 IF !EMPTY(this.procForValid)                   
                    ProcForDo=This.ProcForValid 
                    &ProcForDo
                 ENDIF         
ENDDEFINE 
*******************
PROCEDURE addCommandGroup
PARAMETERS parFrm,parName,parKvo,parTop,parLeft
objGroup=parname
&parFrm..AddObject(objGroup,'myCommandGroup')
WITH &parFrm..&objGroup
     .ButtonCount=parKvo
     .Top=parTop
     .Left=parLeft
*     .Caption=parCaption
*     .Width=parWidth
*     .Height=parHeight
*     .procForClick=parProc
*     .Visible=parVisible
      
ENDWITH
**********************************************************************************************
*                Процедура для печати промежуточных форм
**********************************************************************************************
*PROCEDURE sprprn
*PARAMETERS parproc
*DO FORM setupprn
************************************************************************************************
*                  Процедура построения меню POPUP
************************************************************************************************
PROCEDURE proc_men
PARAMETERS par1,par2,par3,par4
DEFINE POPUP short SHORTCUT RELATIVE FROM par3,par4 FONT 'courier new',9 COLOR SCHEME 4
FOR i=1 TO &par2
    DEFINE BAR i OF short PROMPT par1(i)
ENDFOR
ON SELECTION POPUP short DO procmen  
ACTIVATE POPUP short
***************************************************************************
*
***************************************************************************
PROCEDURE procmen
men_cx=BAR()
DEACTIVATE POPUP short
************************************************************************
PROCEDURE moneyToStr
PARAMETERS prnrub,prnbel,lognum
SELECT 0
USE number
mony=STR(INT(ABS(sumpodr)),12)
cents=ABS(sumpodr)-INT(ABS(sumpodr))
j=1
endsay=''
endjs=''
strcents=''
IF INT(sumpodr)=0
   fname='ноль '
ENDIF
DO WHILE j<=12
   js=SUBSTR(mony,j,1)
   IF js#' '.AND.js#'0'
      GO VAL(js)
      DO CASE
         CASE j=1.OR.j=4.OR.j=7.OR.j=10
              fname=fname+TRIM(cent)
              IF j=1.AND.SUBSTR(mony,2,1)='0'.AND.SUBSTR(mony,3,1)='0'
                 fname=fname+' '+'миллиардов'
              ENDIF
              IF j=4.AND.SUBSTR(mony,5,1)='0'.AND.SUBSTR(mony,6,1)='0'
                 fname=fname+' '+'миллионов'
              ENDIF
              IF j=7.AND.SUBSTR(mony,8,1)='0'.AND.SUBSTR(mony,9,1)='0'
                 fname=fname+' '+'тысяч'
              ENDIF
         CASE j=2.OR.j=5.OR.j=8.OR.j=11
              IF js#'1'.OR.(js='1'.AND.SUBSTR(mony,j+1,1)='0')
                 fname=fname+TRIM(dix2)
                 IF j=2.AND.SUBSTR(mony,3,1)='0'
                    fname=fname+' '+'миллиардов'
                 ENDIF
                 IF j=5.AND.SUBSTR(mony,6,1)='0'
                    fname=fname+' '+'миллионов'
                 ENDIF
                 IF j=8.AND.SUBSTR(mony,9,1)='0'
                    fname=fname+' '+'тысяч'
                 ENDIF
              ELSE
                 GO VAL(SUBSTR(mony,j+1,1))
                 fname=fname+TRIM(dix1)
                 DO CASE
                    CASE j=2
                         fname=fname+' '+'миллиардов'
                    CASE j=5
                         fname=fname+' '+'миллионов'
                    CASE j=8
                         fname=fname+' '+'тысяч'
                 ENDCASE
                 j=j+1
              ENDIF
         CASE j=3
              fname=fname+TRIM(seul)+' '+;
              IIF(VAL(js)>4,'миллиардов',IIF(js='1','миллиард','миллиарда'))
         CASE j=6
              fname=fname+TRIM(seul)+' '+;
              IIF(VAL(js)>4,'миллионов',IIF(js='1','миллион','миллиона'))
         CASE j=9
              fname=fname+TRIM(fem)+' '+;
              IIF(VAL(js)>4,'тысяч',IIF(js='1','тысяча','тысячи'))
         CASE j=12
              fname=fname+TRIM(seul)
      ENDCASE
      fname=fname+' '      
   ENDIF
   j=j+1
ENDDO
endsay=IIF(prnbel,'белорусских рублей,','рублей,')
endjs=RIGHT(ALLTRIM((fname)),2)
DO CASE
   CASE INLIST(endjs,'ре','ри','ва')
        endsay=IIF(prnbel,'белорусских рубля,','рубля,')
   CASE INLIST(endjs,'мь','ть','то','та','ят')
        endsay=IIF(prnbel,'белорусских рублей,','рублей,')
   CASE INLIST(endjs,'ин')   
        endsay=IIF(prnbel,'белорусский рубль,','рубль,')     
   CASE INLIST(endjs,'ль')   
        endsay=IIF(prnbel,'белорусских рублей,','рублей,')               
ENDCASE   
frubl=endsay
IF prnrub.AND.!EMPTY(fname)
  fname=fname+endsay   
ENDIF   
fname=UPPER(LEFT(fname,1))+SUBSTR(fname,2)
IF cents#0   
      strCents=PADL(LTRIM(STR(cents*100)),2,'0')
      cents=ROUND(cents*100,0)
      DO CASE
         CASE cents<10
              GO cents
              strcents=strcents+' '+ALLTRIM(ncent)
         CASE cents=10
              strcents='10 копеек'   
         CASE cents>10.AND.cents<20
              GO cents-10
              strcents=strcents+' копеек' 
         OTHERWISE     
              GO INT(cents/10)
              IF cents-INT(cents/10)*10#0
                 gocx=cents-INT(cents/10)*10
                 GO gocx
                 strcents=strcents+' '+ALLTRIM(ncent)
              ELSE 
                 strcents=strcents+' копеек'                 
              ENDIF   
      ENDCASE    
ELSE 
  strcents='00 копеек'    
ENDIF
IF lognum
   fnamenum=LTRIM(mony)+' '+endsay+' '+strcents 
ENDIF
fname=fname+' '+strcents
USE


************************************************************************
PROCEDURE moneyToStr1
PARAMETERS prnrub,prnbel
SELECT 0
USE number
mony=STR(INT(ABS(sumpodr)),12)
cents=ABS(sumpodr)-INT(ABS(sumpodr))
j=1
endsay=''
endjs=''
strcents=''
DO WHILE j<=12
   js=SUBSTR(mony,j,1)
   IF js#' '.AND.js#'0'
      GO VAL(js)
      DO CASE
         CASE j=1.OR.j=4.OR.j=7.OR.j=10
              fname=fname+TRIM(cent)
              IF j=1.AND.SUBSTR(mony,2,1)='0'.AND.SUBSTR(mony,3,1)='0'
                 fname=fname+' '+'миллиардов'
              ENDIF
              IF j=4.AND.SUBSTR(mony,5,1)='0'.AND.SUBSTR(mony,6,1)='0'
                 fname=fname+' '+'миллионов'
              ENDIF
              IF j=7.AND.SUBSTR(mony,8,1)='0'.AND.SUBSTR(mony,9,1)='0'
                 fname=fname+' '+'тысяч'
              ENDIF
         CASE j=2.OR.j=5.OR.j=8.OR.j=11
              IF js#'1'.OR.(js='1'.AND.SUBSTR(mony,j+1,1)='0')
                 fname=fname+TRIM(dix2)
                 IF j=2.AND.SUBSTR(mony,3,1)='0'
                    fname=fname+' '+'миллиардов'
                 ENDIF
                 IF j=5.AND.SUBSTR(mony,6,1)='0'
                    fname=fname+' '+'миллионов'
                 ENDIF
                 IF j=8.AND.SUBSTR(mony,9,1)='0'
                    fname=fname+' '+'тысяч'
                 ENDIF
              ELSE
                 GO VAL(SUBSTR(mony,j+1,1))
                 fname=fname+TRIM(dix1)
                 DO CASE
                    CASE j=2
                         fname=fname+' '+'миллиардов'
                    CASE j=5
                         fname=fname+' '+'миллионов'
                    CASE j=8
                         fname=fname+' '+'тысяч'
                 ENDCASE
                 j=j+1
              ENDIF
         CASE j=3
              fname=fname+TRIM(seul)+' '+;
              IIF(VAL(js)>4,'миллиардов',IIF(js='1','миллиард','миллиарда'))
         CASE j=6
              fname=fname+TRIM(seul)+' '+;
              IIF(VAL(js)>4,'миллионов',IIF(js='1','миллион','миллиона'))
         CASE j=9
              fname=fname+TRIM(fem)+' '+;
              IIF(VAL(js)>4,'тысяч',IIF(js='1','тысяча','тысячи'))
         CASE j=12
              fname=fname+TRIM(seul)
      ENDCASE
      fname=fname+' '
   ENDIF
   j=j+1
ENDDO
IF EMPTY(ALLTRIM(fname))
   fname='Ноль '
ENDIF 
endsay=IIF(prnbel,'белорусских рублей','рублей')
endjs=RIGHT(ALLTRIM((fname)),2)
DO CASE
   CASE INLIST(endjs,'ре','ри','ва')
        endsay=IIF(prnbel,'белорусских рубля','рубля')
   CASE INLIST(endjs,'мь','ть','то','та','ят')
        endsay=IIF(prnbel,'белорусских рублей','рублей')
   CASE INLIST(endjs,'ин')   
        endsay=IIF(prnbel,'белорусский рубль','рубль')          
ENDCASE   
frubl=endsay
IF prnrub.AND.!EMPTY(fname)
  fname=fname+endsay   
ENDIF   
fname=UPPER(LEFT(fname,1))+SUBSTR(fname,2)
IF cents#0   
      strCents=PADL(LTRIM(STR(cents*100)),2,'0')
      cents=ROUND(cents*100,0)
      DO CASE
         CASE cents<10
              GO cents
              strcents=strcents+' '+ALLTRIM(ncent)
         CASE cents=10
              strcents='10 копеек'   
         CASE cents>10.AND.cents<20         
              GO cents-10
              strcents=strcents+' копеек' 
         OTHERWISE     
              GO INT(cents/10)
              IF cents-INT(cents/10)*10#0
                 gocx=cents-INT(cents/10)*10
                 GO gocx
                 strcents=strcents+' '+ALLTRIM(ncent)
              ELSE 
                 strcents=strcents+' копеек'                 
              ENDIF   
      ENDCASE    
ELSE 
  strcents='00 копеек'    
ENDIF
fname=fname+', '+strcents
USE
************************************************************************               
*                Вывод на печать суммы прописью
************************************************************************
PROCEDURE vs37
PARAMETERS prnrub,prnbel
SELECT 0
USE number
mony=STR(INT(ROUND(ABS(sumpodr),2)),12)
j=1
endsay=''
endjs=''
DO WHILE j<=12
   js=SUBSTR(mony,j,1)
   IF js#' '.AND.js#'0'
      GO VAL(js)
      DO CASE
         CASE j=1.OR.j=4.OR.j=7.OR.j=10
              fname=fname+TRIM(cent)
              IF j=1.AND.SUBSTR(mony,2,1)='0'.AND.SUBSTR(mony,3,1)='0'
                 fname=fname+' '+'миллиардов'
              ENDIF
              IF j=4.AND.SUBSTR(mony,5,1)='0'.AND.SUBSTR(mony,6,1)='0'
                 fname=fname+' '+'миллионов'
              ENDIF
              IF j=7.AND.SUBSTR(mony,8,1)='0'.AND.SUBSTR(mony,9,1)='0'
                 fname=fname+' '+'тысяч'
              ENDIF
         CASE j=2.OR.j=5.OR.j=8.OR.j=11
              IF js#'1'.OR.(js='1'.AND.SUBSTR(mony,j+1,1)='0')
                 fname=fname+TRIM(dix2)
                 IF j=2.AND.SUBSTR(mony,3,1)='0'
                    fname=fname+' '+'миллиардов'
                 ENDIF
                 IF j=5.AND.SUBSTR(mony,6,1)='0'
                    fname=fname+' '+'миллионов'
                 ENDIF
                 IF j=8.AND.SUBSTR(mony,9,1)='0'
                    fname=fname+' '+'тысяч'
                 ENDIF
              ELSE
                 GO VAL(SUBSTR(mony,j+1,1))
                 fname=fname+TRIM(dix1)
                 DO CASE
                    CASE j=2
                         fname=fname+' '+'миллиардов'
                    CASE j=5
                         fname=fname+' '+'миллионов'
                    CASE j=8
                         fname=fname+' '+'тысяч'
                 ENDCASE
                 j=j+1
              ENDIF
         CASE j=3
              fname=fname+TRIM(seul)+' '+;
              IIF(VAL(js)>4,'миллиардов',IIF(js='1','миллиард','миллиарда'))
         CASE j=6
              fname=fname+TRIM(seul)+' '+;
              IIF(VAL(js)>4,'миллионов',IIF(js='1','миллион','миллиона'))
         CASE j=9
              fname=fname+TRIM(fem)+' '+;
              IIF(VAL(js)>4,'тысяч',IIF(js='1','тысяча','тысячи'))
         CASE j=12
              fname=fname+TRIM(seul)
      ENDCASE
      fname=fname+' '
   ENDIF
   j=j+1
ENDDO
endsay=IIF(prnbel,'белорусских рублей','рублей')
endjs=RIGHT(ALLTRIM((fname)),2)
DO CASE
   CASE INLIST(endjs,'ре','ри','ва')
        endsay=IIF(prnbel,'белорусских рубля','рубля')
   CASE INLIST(endjs,'мь','ть','то','та','ят')
        endsay=IIF(prnbel,'белорусских рублей','рублей')
   CASE INLIST(endjs,'ин')   
        endsay=IIF(prnbel,'белорусский рубль','рубль')          
ENDCASE   
frubl=endsay
IF prnrub.AND.!EMPTY(fname)
  fname=fname+endsay    
ENDIF   
fname=UPPER(LEFT(fname,1))+SUBSTR(fname,2)
USE
*************************************************************************************************************************
*         Процедура выполняемая по щелчку заголовка столбца Grid
*************************************************************************************************************************
PROCEDURE clickcont
PARAMETERS parfrm,parcont,parbase,parord1,parord2
*parord1=Acsending
*parord2=Descending
&parfrm..SetAll('SpecialEffect',0,'ContHead')       
&parcont..SpecialEffect=1
SELECT &parbase
IF EMPTY(parord2)
   SET ORDER TO parord1
ELSE
   IF VAL(SYS(21))=parord1.AND.!EMPTY(TAG(parord2))
      SET ORDER TO parord2
   ELSE
      SET ORDER TO parord1
   ENDIF
ENDIF   
GO TOP 
&parfrm..Refresh
*************************************************************************************************************************
*         Процедура выполняемая по щелчку заголовка столбца Grid
*************************************************************************************************************************
PROCEDURE adtboxascnt
PARAMETERS parfrm,parcont,parbase,parord1,parord2
*parord1=Acsending
*parord2=Descending
&parfrm..SetAll('SpecialEffect',0,'ContHead')       
&parcont..SpecialEffect=1
SELECT &parbase
IF EMPTY(parord2)
   SET ORDER TO parord1
ELSE
   IF VAL(SYS(21))=parord1.AND.!EMPTY(TAG(parord2))
      SET ORDER TO parord2
   ELSE
      SET ORDER TO parord1
   ENDIF
ENDIF   
GO TOP 
&parfrm..Refresh

*------------------------------------------------------------------------------------------------------------------------
*                       "Кнопка" "иконка" + надпись
*------------------------------------------------------------------------------------------------------------------------
*************************************************************************************************************************
*   Создание класса Container для "кнопок" ("иконка"+ надпись) меню
*************************************************************************************************************************
DEFINE CLASS mymenucont AS CONTAINER
             VISIBLE=.T.
             BACKSTYLE=1
             BORDERWIDTH=0
             Height=36 
             SPECIALEFFECT=0  
             BACKCOLOR=dBackcolor
             procForClick=''
       ADD OBJECT ContImage AS Image WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
                  Stretch=1, Height=32, Width=32, Top=1, Left=5
       ADD OBJECT ContLabel AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize,;
       			  Height=dFontSize+2,Left=This.ContImage.Left+This.ContImage.Width+3
     
       *--------------------------------------------------------------------------
       PROCEDURE Init
       This.BackColor=This.Parent.BackColor
       *--------------------------------------------------------------------------
       PROCEDURE procenabled
       PARAMETERS par_log
       IF par_log
          This.Enabled=.T.
          This.ContImage.Enabled=.T.
          This.Contlabel.Enabled=.T.
       ELSE
          This.Enabled=.F.
          This.ContImage.Enabled=.F.
          This.Contlabel.Enabled=.F.
       ENDIF    
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.CLICK
       This.Parent.Click      
       *--------------------------------------------------------------------------
       PROCEDURE ContLabel.CLICK
       This.Parent.Click            
       *--------------------------------------------------------------------------
       PROCEDURE CLICK   
       This.BackColor=This.Parent.BackColor
       This.BorderWidth=0
       IF !EMPTY(This.ProcForClick)
          ProcForDo=This.ProcForClick 
          &ProcForDo
       ENDIF         
       *--------------------------------------------------------------------------
       PROCEDURE MouseEnter
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       ThisForm.SetAll('BorderWidth',0,'mymenucont')
       ThisForm.SetAll('BackColor',This.Parent.Backcolor,'mymenucont')
       This.BorderWidth=1
       This.BackColor=RGB(255,255,255)
       *-----------------------------------------------------------------------
       PROCEDURE MouseLeave    
       LPARAMETERS nButton, nShift, nXCoord, nYCoord         
       This.BorderWidth=0 
       This.BackColor=This.Parent.BackColor
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown     
      * *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseUp     
       *--------------------------------------------------------------------------
       PROCEDURE ContLabel.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown      
      * *--------------------------------------------------------------------------
       PROCEDURE ContLabel.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.parent.MouseUp   
      * *--------------------------------------------------------------------------
       PROCEDURE MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=1  
       This.Visible=.T.        
      * *--------------------------------------------------------------------------
       PROCEDURE MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=0   
       This.Visible=.T.   
ENDDEFINE
*************************************************************************************************************************
*                         Процедура добавления "кнопки в форму"
*************************************************************************************************************************
PROCEDURE addcontmenu
PARAMETERS parfrm,parname,parleft,partop,parlab,parpict,parproc,partool
objcontmenu=parname
&parfrm..AddObject(objcontmenu,'mymenucont')  
WITH &parfrm..&objcontmenu   
     .contimage.Picture=parpict  
     .contlabel.Caption=parlab
     .contlabel.Top=(&parfrm..&objcontmenu..Height-&parfrm..&objcontmenu..contlabel.Height)/2-1     
     .contlabel.Autosize=.T.
     .contlabel.Visible=.T.
     .Top=partop
     .Left=parleft     
     .Width=&parfrm..&objcontmenu..contlabel.Width+&parfrm..&objcontmenu..contimage.Width+10  
     .procForClick=IIF(!EMPTY(parproc),parproc,'')
     IF !EMPTY(partool)
        .contimage.ToolTipText=partool
        .contlabel.ToolTipText=partool
     ENDIF
ENDWITH
*************************************************************************************************************************
*                         Процедура добавления "кнопки в форму"
*************************************************************************************************************************
PROCEDURE addButtonPictLabel
PARAMETERS parfrm,parname,parleft,partop,parlab,parpict,parproc,partool,parWidth
objcontmenu=parname
&parfrm..AddObject(objcontmenu,'mymenucont')  
WITH &parfrm..&objcontmenu   
     .contimage.Picture=parpict  
     .contlabel.Caption=parlab
     .contlabel.Top=(&parfrm..&objcontmenu..Height-&parfrm..&objcontmenu..contlabel.Height)/2-1     
     .contlabel.Autosize=.T.
     .contlabel.Visible=.T.
     .contImage.Left=(parWidth-.contImage.Width-.contLabel.Width-5)/2
     .contLabel.Left=.contImage.Left+.contImage.Width+5
     .Top=partop
     .Left=parleft     
     .Width=parWidth  
     .procForClick=IIF(!EMPTY(parproc),parproc,'')
     IF !EMPTY(partool)
        .contimage.ToolTipText=partool
        .contlabel.ToolTipText=partool
     ENDIF
ENDWITH
*************************************************************************************************************************
DEFINE CLASS icoincolumn as Container 
       VISIBLE=.T.
       BACKSTYLE=0
       BORDERWIDTH=0
       HEIGHT=30
       WIDTH=30  
       SPECIALEFFECT=0  
       procForClick=''
       ADD OBJECT ContImage AS Image WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
                  Stretch=1, Height=28, Width=28, Top=1, Left=1     
         
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.CLICK
       This.Parent.Click                        
       *--------------------------------------------------------------------------
       PROCEDURE CLICK   
       IF !EMPTY(This.ProcForClick)
          ProcForDo=This.ProcForClick 
          &ProcForDo
       ENDIF         
       *--------------------------------------------------------------------------
       PROCEDURE MouseEnter
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.BorderWidth=1
       *-----------------------------------------------------------------------
       PROCEDURE MouseLeave    
       LPARAMETERS nButton, nShift, nXCoord, nYCoord         
       This.BorderWidth=0 
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown     
      * *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseUp            
      * *--------------------------------------------------------------------------
       PROCEDURE MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=1  
       This.Visible=.T.        
      * *--------------------------------------------------------------------------
       PROCEDURE MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=0   
       This.Visible=.T.   
ENDDEFINE 
**************************************************************************************************************************
PROCEDURE adIcoInColumn
PARAMETERS parfrm,parname,parpict,parproc
*parfrm - имя формы
*parname - имя объекта
*parpict - изображние
*parproc - процедура
objcontmenu=parname
&parfrm..AddObject(objcontmenu,'icoincolumn')
WITH &parfrm..&objcontmenu   
     .contimage.Picture=parpict 
     .contimage.BackStyle=0
     .BackStyle=0
     .contimage.Visible=.F.
    * .contImage.Height=&parfrm..Height
     .contImage.Width=&parfrm..Width
     .contimage.Stretch=1    
     *Width=&parfrm.Width
     .procForClick=IIF(!EMPTY(parproc),parproc,'')
    * IF !EMPTY(partool)
    *    .contimage.ToolTipText=partool
    * ENDIF
ENDWITH
*------------------------------------------------------------------------------------------------------------------------
*                       "Кнопка"-"иконка" 
*------------------------------------------------------------------------------------------------------------------------
*************************************************************************************************************************
*         Создание класса Container для "кнопок" ("иконка")
*************************************************************************************************************************
DEFINE CLASS mycontico AS CONTAINER
             VISIBLE=.T.
             BACKSTYLE=0
             BORDERWIDTH=0
             HEIGHT=36
             WIDTH=36  
             SPECIALEFFECT=0  
             procForClick=''
       ADD OBJECT ContImage AS Image WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
                  Stretch=2, Height=32, Width=32, Top=2, Left=2     
         
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.CLICK
       This.Parent.Click                        
       *--------------------------------------------------------------------------
       PROCEDURE CLICK   
       IF !EMPTY(This.ProcForClick)
          ProcForDo=This.ProcForClick 
          &ProcForDo
       ENDIF         
       *--------------------------------------------------------------------------
       PROCEDURE MouseEnter
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.BorderWidth=1
       *-----------------------------------------------------------------------
       PROCEDURE MouseLeave    
       LPARAMETERS nButton, nShift, nXCoord, nYCoord         
       This.BorderWidth=0 
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown     
      * *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseUp            
      * *--------------------------------------------------------------------------
       PROCEDURE MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=1  
       This.Visible=.T.        
      * *--------------------------------------------------------------------------
       PROCEDURE MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=0   
       This.Visible=.T.   
ENDDEFINE
*************************************************************************************************************************
*                         Процедура добавления "кнопки-иконки" в форму
*************************************************************************************************************************
PROCEDURE addcontico
PARAMETERS parfrm,parname,parleft,partop,parpict,parproc,parTool,parWidth,parHeight
* DO addcontico WITH 'fNakl',namecont,leftCont,topCont,datmenu.mIco,datmenu.mproc,ALLTRIM(datmenu.nmenu),36,36   
DO CASE
   CASE VERSION(5)=900 
        objContmenu=parname
        &parFrm..AddObject(objContmenu,'myCommandButton')
        WITH &parFrm..&objContmenu
            .Caption=''
            .Top=parTop
            .Left=parLeft
            .Width=parWidth
            .Height=parHeight
            .Picture=parPict     
            .procForClick=parProc
            .ToolTipText=partool
        *     .Visible=parVisible
        ENDWITH 
   OTHERWISE
        objcontmenu=parname
        &parfrm..AddObject(objcontmenu,'mycontico')  
        WITH &parfrm..&objcontmenu   
            .contimage.Picture=parpict     
            .Top=partop
            .Left=parleft          
            .procForClick=IIF(!EMPTY(parproc),parproc,'')
            IF !EMPTY(partool)
               .contimage.ToolTipText=partool
            ENDIF
        ENDWITH
ENDCASE       
*------------------------------------------------------------------------------------------------------------------------
*                       "Кнопка"-"иконка" (новый вариант)
*------------------------------------------------------------------------------------------------------------------------
*************************************************************************************************************************
*         Создание класса Container для "кнопок" ("иконка")
*************************************************************************************************************************
DEFINE CLASS myconticonew AS CONTAINER
             VISIBLE=.T.
             BACKSTYLE=0
             BORDERWIDTH=0
             SPECIALEFFECT=0  
             procForClick=''
       ADD OBJECT ContImage AS Image WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
                  Stretch=2, Top=2, Left=2     
         
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.CLICK
       This.Parent.Click                        
       *--------------------------------------------------------------------------
       PROCEDURE CLICK   
       IF !EMPTY(This.ProcForClick)
          ProcForDo=This.ProcForClick 
          &ProcForDo
       ENDIF         
       *--------------------------------------------------------------------------
       PROCEDURE MouseEnter
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.BorderWidth=1
       *-----------------------------------------------------------------------
       PROCEDURE MouseLeave    
       LPARAMETERS nButton, nShift, nXCoord, nYCoord         
       This.BorderWidth=0 
       *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown     
      * *--------------------------------------------------------------------------
       PROCEDURE ContImage.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseUp            
      * *--------------------------------------------------------------------------
       PROCEDURE MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=1  
       This.Visible=.T.        
      * *--------------------------------------------------------------------------
       PROCEDURE MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=0   
       This.Visible=.T.   
ENDDEFINE
*************************************************************************************************************************
*                         Процедура добавления "кнопки-иконки" в форму
*************************************************************************************************************************
PROCEDURE addconticonew
PARAMETERS parfrm,parname,parleft,partop,parpict,parWidth,parHeight,parPictWidth,parPictHeight,parProc
* parFrm  - форма
* parname - имя
* parleft 
* partop
* parpict
* parproc
* parWidth
* parHeight
* parPictWidth
* parPictHeight
* parProc 
objcontmenu=parname
&parfrm..AddObject(objcontmenu,'myconticonew')  
WITH &parfrm..&objcontmenu   
     .Width=parWidth
     .Height=parHeight
     .Top=partop
     .Left=parleft  
     .contimage.Picture=parpict 
     .contimage.Width=parPictWidth
     .contImage.Height=parPictHeight                  
     .procForClick=IIF(!EMPTY(parproc),parproc,'')
     
ENDWITH
  
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE addButtonOne
PARAMETERS parfrm,parname,parleft,partop,parlab,parpict,parproc,parHeight,parWidth,partool
 * parfrm - имя формы
 * parname - имя объекта
 * parleft - лево
 * partop -  право
 * parlab - надпись
 * parpict - иконка
 * parproc - процедура
 * partool - toolTipText
DO CASE
   CASE VERSION(5)=900        
        objContmenu=parname
        &parFrm..AddObject(objContmenu,'myCommandButton')
        WITH &parFrm..&objContmenu
             .Caption=parLab
             .Top=parTop
             .Left=parLeft
             .Width=parWidth
             .Height=parHeight
             .Picture=parPict     
             .procForClick=parProc  
             .PicturePosition=1
             .Alignment=2   
        ENDWITH               
               
   OTHERWISE
        objcontmenu=parname  
        &parfrm..AddObject(objcontmenu,'mymenucont')  
        WITH &parfrm..&objcontmenu   
             .Height=IIF(EMPTY(parHeight),36,parHeight)
             .contImage.Height=32
             .contImage.Width=32
             .contImage.Left=2
             .contImage.Top=2     
             .contimage.Picture=parpict  
             .contlabel.Caption=parlab
             .contlabel.Top=(&parfrm..&objcontmenu..Height-&parfrm..&objcontmenu..contlabel.Height)/2-1     
             .contlabel.Autosize=.T.
             .contlabel.Visible=.T.
             .Top=partop
             .Left=parleft     
             .Width=&parfrm..&objcontmenu..contlabel.Width+&parfrm..&objcontmenu..contimage.Width+10
             .procForClick=IIF(!EMPTY(parproc),parproc,'')
             IF !EMPTY(partool)
                .contimage.ToolTipText=partool
                .contlabel.ToolTipText=partool
             ENDIF
        ENDWITH
ENDCASE  

************************************************************************************************************************
*   Процедура добавления кнопки в форму (commandButton)
************************************************************************************************************************
PROCEDURE addCommandButton
PARAMETERS parFrm,parName,parCaption,parTop,parLeft,parWidth,parHeight,parProc,parVisible
* parFrm - имя формы
* parname - имя объекта
* parCaption  - Caption
* parTop  - Top
* parLeft - Left 
* parWidth - Width
* parHeight - Height
* parProc - Click
* parVisible - Visible
objButton=parname
&parFrm..AddObject(objButton,'myButton')
WITH &parFrm..&objButton
     .Top=parTop
     .Left=parLeft
     .Caption=parCaption
     .Width=parWidth
     .Height=parHeight
     .procForClick=parProc
     .Visible=parVisible
ENDWITH
*************************************************************************************************************************
*  Посторение формы для сообщений,переспроса, настроек и т.д.
*************************************************************************************************************************
PROCEDURE formCreate
PARAMETERS parVisible,parCaption,parWidth,parHeight,parWidthCont,parCont1,parCont2,parCont3,parProc1,parProc2,parProc3,parLabel,parExit
dataOld=ALIAS()
FormMy=CREATEOBJECT('MyFormMes')
IF !USED('datapic')
   USE datapic IN 0 ORDER 1
ENDIF
SELECT datapic
widthHeight=pWidth/pHeight
SCAN ALL
     IF nShow=0.AND.pHeight>=&parHeight
        pictRec=RECNO()
        REPLACE nShow WITH  1
        EXIT 
     ENDIF
ENDSCAN
IF EOF() 
   REPLACE nShow WITH 0 ALL
   GO TOP
   REPLACE nShow WITH 1
ENDIF 
pathmem=FULLPATH('osnproc.fxp')
pathmem=LEFT(pathmem,LEN(pathmem)-11)+'datapic'+LTRIM(STR(RECNO()))+'.pic'
widthHeight=pWidth/pHeight
pictHeight=&parHeight
pictWidth=&parHeight*widthHeight
COPY MEMO pict TO &pathmem
IF !EMPTY(dataOld)
   SELECT (dataOld)
ENDIF
WITH FormMy
     .Caption=parCaption
     .logExit=parExit
     
     .FormImage.Picture=pathmem
     .FormImage.Height=pictHeight
     .FormImage.Width=pictWidth
     .FormImage.Visible=.T.
     .Width=parWidth+pictWidth
     .Height=&parHeight   
     .SetAll('Width',parWidthCont,'MyContLabel')
     IF !EMPTY(parLabel)
        .FormLabel.Caption=IIF(!EMPTY(parLabel),parLabel,'')
        .FormLabel.Width=RetTxtWidth(parLabel,dFontName,dFontSize+1)
        .FormLabel.Left=pictWidth+(parWidth-.FormLabel.Width)/2
        .FormLabel.Visible=.T.     
     ENDIF     
     DO CASE
        CASE !EMPTY(parCont1).AND.!EMPTY(parCont2).AND.!EMPTY(parCont3)
             IF VERSION(5)<=700
                DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont*3-30)/2,;
                .FormLabel.Top+.FormLabel.Height+40,parWidthCont,dHeight+5,parCont1,parProc1
                DO addcontlabel WITH 'FormMy','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,parWidthCont,dHeight+5,parCont2,parProc2 
                DO addcontlabel WITH 'FormMy','cont3',.cont2.Left+.cont2.Width+15,.Cont1.Top,parWidthCont,dHeight+5,parCont3,parProc3
             ELSE 
                DO addCommandButton WITH 'FormMy','cont1',parCont1,.FormLabel.Top+.FormLabel.Height+40,pictWidth+(parWidth-parWidthCont*3-30)/2,parWidthCont,dHeight+5,parProc1,.T.
                DO addCommandButton WITH 'FormMy','cont2',parCont2,.cont1.Top,.cont1.Left+.cont1.Width+30,parWidthCont,dHeight+5,parProc2,.T.
                DO addCommandButton WITH 'FormMy','cont3',parCont3,.cont1.Top,.cont2.Left+.cont2.Width+30,parWidthCont,dHeight+5,parProc3,.T.
             ENDIF             
        CASE !EMPTY(parCont1).AND.!EMPTY(parCont2).And.EMPTY(parCont3)
             IF VERSION(5)<=700
                DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont*2-30)/2,;
                .FormLabel.Top+.FormLabel.Height+40,parWidthCont,dHeight+5,parCont1,parProc1
                .FormLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
                .Cont1.Top=.FormLabel.Top+40
             ELSE             
                DO addCommandButton WITH 'FormMy','cont1',parCont1,.FormLabel.Top+.FormLabel.Height+40 ,pictWidth+(parWidth-parWidthCont*2-30)/2,parWidthCont,dHeight+5,parProc1,.T.
                .FormLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
                .Cont1.Top=.FormLabel.Top+40           
                DO addCommandButton WITH 'FormMy','cont2',parCont2,.cont1.Top,.cont1.Left+.cont1.Width+30,parWidthCont,dHeight+5,parProc2,.T.
             ENDIF             
        CASE !EMPTY(parCont1).AND.EMPTY(parCont2).And.EMPTY(parCont3)     
             IF VERSION(5)<=700
                DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont)/2,40,parWidthCont,dHeight+5,parCont1,parProc1            
                .FormLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
                .Cont1.Top=.FormLabel.Top+40
             ELSE 
                DO addCommandButton WITH 'FormMy','cont1',parCont1,.FormLabel.Top+.FormLabel.Height+40 ,pictWidth+(parWidth-parWidthCont*2-30)/2,parWidthCont,dHeight+5,parProc1,.T.                              
                .FormLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
                .Cont1.Top=.FormLabel.Top+40
             ENDIF 
     ENDCASE
     .AutoCenter=.T.        
ENDWITH
IF parVisible
    FormMy.Show    
ENDIF
DELETE FILE &pathmem
*------------------------------------------------------------------------------------------------------------------------
*                       "Кнопка-надпись"
*------------------------------------------------------------------------------------------------------------------------

*************************************************************************************************************************
*   Создание класса Container для "кнопок" ("надпись") меню
*************************************************************************************************************************
DEFINE CLASS mycontlabel AS CONTAINER
             VISIBLE=.T.
             BACKSTYLE=0
             BORDERWIDTH=1
             Height=30  
             SPECIALEFFECT=0  
             procForClick=''  
             procForMouseEnter=''
             procForMouseLeave=''            
       ADD OBJECT ContLabel AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize+1,;
       			  Alignment=2        
       *--------------------------------------------------------------------------
       PROCEDURE INIT
       This.BackColor=This.Parent.BackColor            
       *--------------------------------------------------------------------------
       PROCEDURE ContLabel.CLICK      
       This.Parent.Click            
       *--------------------------------------------------------------------------
       PROCEDURE CLICK   
       This.BorderWidth=1
       This.BackColor=dBackColor 
       IF !EMPTY(This.ProcForClick)
          ProcForDo=This.ProcForClick 
          &ProcForDo
       ENDIF         
       *--------------------------------------------------------------------------
       PROCEDURE MouseEnter
       LPARAMETERS nButton,nShift,nXcoord,nYcoord   
       This.BorderWidth=2
       IF !EMPTY(This.ProcForMouseEnter)
          ProcForDo=This.ProcForMouseEnter 
          &ProcForDo
       ELSE
           This.BackColor=dbackColor  
       ENDIF        
       *-----------------------------------------------------------------------
       PROCEDURE MouseLeave    
       LPARAMETERS nButton, nShift, nXCoord, nYCoord  
       This.BorderWidth=1
        IF !EMPTY(This.ProcForMouseLeave)
          ProcForDo=This.ProcForMouseLeave 
          &ProcForDo       
       ELSE
           This.BackColor=This.Parent.BackColor            
       ENDIF                   
       *--------------------------------------------------------------------------
       PROCEDURE ContLabel.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown      
      * *--------------------------------------------------------------------------
       PROCEDURE ContLabel.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.parent.MouseUp   
      * *--------------------------------------------------------------------------
       PROCEDURE MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=1  
       This.Visible=.T.        
      * *--------------------------------------------------------------------------
       PROCEDURE MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=0   
       This.Visible=.T.   
ENDDEFINE

*************************************************************************************************************************
*                         Процедура добавления "кнопки-надписи" в форму"
*************************************************************************************************************************
PROCEDURE addcontlabel
PARAMETERS parfrm,parname,parleft,partop,parwidth,parheight,parlab,parproc,partool,parFontSize
DO CASE
   CASE VERSION(5)=900
        objButton=parname
        &parFrm..AddObject(objButton,'myCommandButton')
        WITH &parfrm..&objButton      
             .Width=parWidth  
             .Top=partop
             .Left=parleft
             .Caption=parlab     
             IF !EMPTY(parHeight)
                .Height=parHeight    
             ENDIF     
             .procForClick=IIF(!EMPTY(parproc),parproc,'')
             IF !EMPTY(partool)       
                *.contlabel.ToolTipText=partool
             ENDIF
        ENDWITH
   OTHERWISE
        objcontmenu=parname
        &parfrm..AddObject(objcontmenu,'mycontlabel')  
        WITH &parfrm..&objcontmenu            
             .BackStyle=1
             .BackColor=&parFrm..BackColor
             .Width=parWidth  
             .Top=partop
             .Left=parleft
             .ContLabel.Caption=parlab       
             .ContLabel.AutoSize=.T.
             .ContLabel.Visible=.T.
             IF !EMPTY(parHeight)
                .Height=parHeight
             ELSE
                .Height=.contLabel.Height+8
             ENDIF
             IF !EMPTY(parFontSize)
                .ContLabel.FontSize=parFontSize
                .ContLabel.Top=(&parfrm..&objcontmenu..Height-FONTMETRIC(1,dFontName,parFontSize))/2          
                .ContLabel.Left=(&parfrm..&objcontmenu..Width-RetTxtWidth(parlab,dFontName,parFontSize))/2
             ELSE 
                .ContLabel.Top=(&parfrm..&objcontmenu..Height-FONTMETRIC(1,dFontName,dFontSize+1))/2          
                .ContLabel.Left=(&parfrm..&objcontmenu..Width-RetTxtWidth(parlab,dFontName,dFontSize+1))/2  
             ENDIF   
             .ContLabel.Visible=.T.                                          
             .procForClick=IIF(!EMPTY(parproc),parproc,'')
             IF !EMPTY(partool)       
                .contlabel.ToolTipText=partool
             ENDIF
        ENDWITH
ENDCASE

*************************************************************************************************************************
*   Создание класса Container для всплывающих "кнопок" ( надпись)
*************************************************************************************************************************
DEFINE CLASS mymenucontvs AS CONTAINER
             VISIBLE=.T.
             BACKSTYLE=1
             BORDERWIDTH=0
             Height=30  
             SPECIALEFFECT=0  
             BACKCOLOR=RGB(255,255,255)
             procForClick=''       
       ADD OBJECT ContLabel AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize,;
       			  Height=dFontSize+2,Left=3         
       
       *--------------------------------------------------------------------------
       PROCEDURE ContLabel.CLICK
       This.Parent.Click            
       *--------------------------------------------------------------------------
       PROCEDURE CLICK   
       This.BackColor=dBackColor
       This.BorderWidth=0
       IF !EMPTY(This.ProcForClick)
          ProcForDo=This.ProcForClick 
          &ProcForDo
       ENDIF         
       *--------------------------------------------------------------------------
       PROCEDURE MouseEnter
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       ThisForm.SetAll('BorderWidth',0,'mymenucont')
       ThisForm.SetAll('BackColor',dBackcolor,'mymenucont')
       This.BorderWidth=1
       This.BackColor=dBackColor
       *-----------------------------------------------------------------------
       PROCEDURE MouseLeave    
       LPARAMETERS nButton, nShift, nXCoord, nYCoord         
       This.BorderWidth=0 
       This.BackColor=RGB(255,255,255)  
       *--------------------------------------------------------------------------
       PROCEDURE ContLabel.MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.Parent.MouseDown      
      * *--------------------------------------------------------------------------
       PROCEDURE ContLabel.MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.parent.MouseUp   
      * *--------------------------------------------------------------------------
       PROCEDURE MouseDown
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=1  
       This.Visible=.T.        
      * *--------------------------------------------------------------------------
       PROCEDURE MouseUp
       LPARAMETERS nButton,nShift,nXcoord,nYcoord
       This.SpecialEffect=0   
       This.Visible=.T.   
ENDDEFINE
***********************************************************************************************************************
PROCEDURE sprrepprn
PARAMETERS parForm,parRep
&parForm..Visible=.F.
Report Form &parRep TO PRINTER PROMPT PREVIEW
&parForm..Visible=.T.
********************************************************************************************************
*  Класс для окна предварительного просмотра
********************************************************************************************************
DEFINE CLASS FormPrew AS Form      
       TOP=0
       LEFT=0
       Height=SYSMETRIC(22)
       Width=SYSMETRIC(21)    
       TitleBar=0
       AlwaysOnTop=.T.
       MinButton=.F.
       MaxButton=.F.
       WindowState=2
       ShowWindow=2   
       AlwaysOnTop=.T.    
ENDDEFINE
*-----------------------------------------------------------------------
*       Процeдура Drag to the top для Toolbar "Print Preview"
*   Вызов процедуры вставляется в событие On Entry полосы отчета Title
*-----------------------------------------------------------------------
PROCEDURE topToolPreview
PARAMETERS parHidePrinting
IF WEXIST("Print Preview").AND.log_tool
   nRow=MROW()
   nCol=MCOL()
   MOUSE DBLCLICK AT WLROW("Print Preview")+0.6,WLCOL("Print Preview")+0.6
   MOUSE AT nRow,nCol
   MOUSE DBLCLICK AT WLROW("Print Preview")+3,WLCOL("Print Preview")+3
   log_tool=.F.
ENDIF
*************************************************************************************************************************
*                   Процедура для печати отчетов после предварительного просмотра
*************************************************************************************************************************
PROCEDURE previewrep
PARAMETERS parReport,parCaption
log_tool=.T.
winPrew=CREATEOBJECT('FormPrew')
winPrew.Name='winqwe'
winPrew.Caption=parCaption
winPrew.Show
REPORT FORM &parReport NOCONSOLE TO PRINTER  PROMPT PREVIEW WINDOW winqwe IN WINDOW winqwe   
WinPrew.Release
*************************************************************************************************************************
*       Класс для создания формы для переспроса и т.д. 
*************************************************************************************************************************
DEFINE CLASS MyFormMes AS Form
       DeskTop=.T.
       Top=0
       Left=0     
       WindowState=0          
       WindowType=1      
       ShowWindow=1                  
       DoCreate = .T.
       BorderStyle = 2	   
       Minbutton=.F. 
       MaxButton=.F.       
       FontName=dFontName
       FontSize=dFontSize 
       BackColor=RGB(255,255,255) 
       procExit=''  
       procForKeyPress=''  
       logexit=.F.  
       ADD OBJECT FormLabel AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize+1,;
       			  Top=15,Left=0    
       ADD OBJECT FormLabel2 AS LABEL WITH VISIBLE=.T., BACKSTYLE=0, BORDERSTYLE=0,;
       			  ForeColor=ObjForeColor, FontName=dFontName, FontSize=dFontSize+1,;
       			  Top=15,Left=0             			     
       ADD OBJECT formImage AS IMAGE WITH VISIBLE=.F., STRETCH=2, LEFT=0, TOP=0, WIDTH=0,;
                   HEIGHT=0   
       PROCEDURE Init
                 nFormMes=ThisForm                                                                     
*------------------------------------------------------------------------------------------------------------------------                 
       PROCEDURE QueryUnload
                 IF !EMPTY(This.procExit)
                    procdo=This.procExit
                    &procdo   
                 ENDIF 
       ENDPROC   
*--------------------------------------------------------------------------------------------------
       PROCEDURE KeyPress
       LPARAMETERS nKeyCode, nShiftAltCtrl
       IF This.logexit.AND.nKeyCode=27
          This.Release
       ENDIF
       IF !EMPTY(THIS.procForKeyPress)
          procForDo=THIS.procForKeyPress     
          &procForDo
       ENDIF           
 
ENDDEFINE
*************************************************************************************************************************
*  Посторение формы для сообщений,переспроса, настроек и т.д.
*************************************************************************************************************************
PROCEDURE Createform
PARAMETERS parVisible,parCaption,parWidth,parHeight,parWidthCont,parCont1,parCont2,parCont3,parProc1,parProc2,parProc3,parLabel,parExit
dataOld=ALIAS()
FormMy=CREATEOBJECT('MyFormMes')
IF !USED('datapic')
   USE datapic IN 0 ORDER 1
ENDIF
SELECT datapic
widthHeight=pWidth/pHeight
SCAN ALL
     IF nShow=0.AND.pHeight>=&parHeight
        pictRec=RECNO()
        REPLACE nShow WITH  1
        EXIT 
     ENDIF
ENDSCAN
IF EOF() 
   REPLACE nShow WITH 0 ALL
   GO TOP
   REPLACE nShow WITH 1
ENDIF 
pathmem=FULLPATH('osnproc.fxp')
pathmem=LEFT(pathmem,LEN(pathmem)-11)+'datapic'+LTRIM(STR(RECNO()))+'.pic'
widthHeight=pWidth/pHeight
pictHeight=&parHeight
pictWidth=&parHeight*widthHeight
COPY MEMO pict TO &pathmem
IF !EMPTY(dataOld)
   SELECT (dataOld)
ENDIF
WITH FormMy
     .Caption=parCaption
     .logExit=parExit
     .FormImage.Picture=pathmem
     .FormImage.Height=pictHeight
     .FormImage.Width=pictWidth
     .FormImage.Visible=.T.
     .Width=parWidth+pictWidth
     .Height=&parHeight   
     .SetAll('Width',parWidthCont,'MyContLabel')
     IF !EMPTY(parLabel)
        .FormLabel.Caption=IIF(!EMPTY(parLabel),parLabel,'')
        .FormLabel.Width=RetTxtWidth(parLabel,dFontName,dFontSize+1)
        .FormLabel.Left=pictWidth+(parWidth-.FormLabel.Width)/2
        .FormLabel.Visible=.T.     
     ENDIF     
     DO CASE
        CASE !EMPTY(parCont1).AND.!EMPTY(parCont2).AND.!EMPTY(parCont3)
              DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont*3-30)/2,;
             .FormLabel.Top+.FormLabel.Height+40,parWidthCont,dHeight+5,parCont1,parProc1
             DO addcontlabel WITH 'FormMy','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,parWidthCont,dHeight+5,parCont2,parProc2 
             DO addcontlabel WITH 'FormMy','cont3',.cont2.Left+.cont2.Width+15,.Cont1.Top,parWidthCont,dHeight+5,parCont3,parProc3
        CASE !EMPTY(parCont1).AND.!EMPTY(parCont2).And.EMPTY(parCont3)
             DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont*2-30)/2,;
             .FormLabel.Top+.FormLabel.Height+40,parWidthCont,dHeight+5,parCont1,parProc1
             .FormLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
             .Cont1.Top=.FormLabel.Top+40
             DO addcontlabel WITH 'FormMy','cont2',.cont1.Left+.cont1.Width+30,.Cont1.Top,parWidthCont,dHeight+5,parCont2,parProc2
        CASE !EMPTY(parCont1).AND.EMPTY(parCont2).And.EMPTY(parCont3)                     
             DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont)/2,40,parWidthCont,dHeight+5,parCont1,parProc1            
             .FormLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
             .Cont1.Top=.FormLabel.Top+40
     ENDCASE
     .AutoCenter=.T.        
ENDWITH
IF parVisible
    FormMy.Show 
ENDIF
DELETE FILE &pathmem

*************************************************************************************************************************
*  Посторение формы для сообщений,переспроса, настроек и т.д.
*************************************************************************************************************************
PROCEDURE Createformnew
PARAMETERS parVisible,parCaption,parWidth,parHeight,parWidthCont,parCont1,parCont2,parCont3,parProc1,parProc2,parProc3,parLabel,parlabel2,parExit
dataOld=ALIAS()
FormMy=CREATEOBJECT('MyFormMes')
pictWidth=0
pictHeight=0
IF pictset>1
   IF !USED('datapic')
      USE datapic IN 0 ORDER 1
   ENDIF
   SELECT datapic
   widthHeight=pWidth/pHeight
   IF pictset>2
      SEEK pictset     
   ENDIF
   SCAN WHILE IIF(pictset>2,theme=pictset,!EOF())
        IF nShow=0.AND.pHeight>=&parHeight
           pictRec=RECNO()
           REPLACE nShow WITH  1
           EXIT 
        ENDIF
   ENDSCAN
   IF pictset>2
      IF theme#pictset 
         REPLACE nShow WITH 0 FOR theme=pictset
         SEEK pictset
         REPLACE nShow WITH 1
      ENDIF 
   ELSE 
      IF EOF() 
         REPLACE nShow WITH 0 ALL
         GO TOP
         REPLACE nShow WITH 1
      ENDIF 
   ENDIF   
   pathmem=FULLPATH('osnproc.fxp')
   pathmem=LEFT(pathmem,LEN(pathmem)-11)+'datapic'+LTRIM(STR(RECNO()))+'.pic'
   widthHeight=pWidth/pHeight
   pictHeight=&parHeight
   pictWidth=&parHeight*widthHeight
   COPY MEMO pict TO &pathmem
ENDIF

IF !EMPTY(dataOld)
   SELECT (dataOld)
ENDIF
WITH FormMy
     .Caption=parCaption
     .logExit=parExit
     IF pictset>1
        .FormImage.Picture=pathmem
        .FormImage.Height=pictHeight
        .FormImage.Width=pictWidth
        .FormImage.Visible=.T.
     ENDIF
     .Width=parWidth+pictWidth
     .Height=&parHeight   
     .SetAll('Width',parWidthCont,'MyContLabel')
     IF !EMPTY(parLabel)
        .formLabel.Caption=IIF(!EMPTY(parLabel),parLabel,'')
        .formLabel.Width=RetTxtWidth(parLabel,dFontName,dFontSize+1)
        .formLabel.Left=pictWidth+(parWidth-.FormLabel.Width)/2
        .formLabel.Visible=.T.     
     ENDIF     
     IF !EMPTY(parlabel2)
        .formLabel2.Caption=IIF(!EMPTY(parLabel2),parLabel2,'')
        .formLabel2.Width=RetTxtWidth(parLabel2,dFontName,dFontSize+1)
        .formLabel2.Left=pictWidth+(parWidth-.FormLabel2.Width)/2
        .formLabel2.Visible=.T.
        .formlabel2.Top=.formlabel.Top+.formlabel.Height+5
     ELSE                 
        .formlabel2.Top=.formlabel.Top
        .FormLabel2.Visible=.F.
     ENDIF
     DO CASE
        CASE !EMPTY(parCont1).AND.!EMPTY(parCont2).AND.!EMPTY(parCont3)
             DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont*3-30)/2,;
             .formLabel2.Top+.formLabel.Height+40,parWidthCont,dHeight+5,parCont1,parProc1
             DO addcontlabel WITH 'FormMy','cont2',.cont1.Left+.cont1.Width+15,.Cont1.Top,parWidthCont,dHeight+5,parCont2,parProc2 
             DO addcontlabel WITH 'FormMy','cont3',.cont2.Left+.cont2.Width+15,.Cont1.Top,parWidthCont,dHeight+5,parCont3,parProc3
        CASE !EMPTY(parCont1).AND.!EMPTY(parCont2).And.EMPTY(parCont3)
             DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont*2-30)/2,;
             .formLabel2.Top+.FormLabel.Height+40,parWidthCont,dHeight+5,parCont1,parProc1
             .formLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
             .formLabel2.Top=IIF(EMPTY(parlabel2),.formlabel.Top,.Formlabel.Top+.formlabel.Height+5)
             .Cont1.Top=.FormLabel2.Top+40
             DO addcontlabel WITH 'FormMy','cont2',.cont1.Left+.cont1.Width+30,.Cont1.Top,parWidthCont,dHeight+5,parCont2,parProc2
        CASE !EMPTY(parCont1).AND.EMPTY(parCont2).And.EMPTY(parCont3)                     
             DO addcontlabel WITH 'FormMy','cont1',pictWidth+(parWidth-parWidthCont)/2,40,parWidthCont,dHeight+5,parCont1,parProc1            
             .formLabel.Top=(.Height-(.FormLabel.Height+.Cont1.Height+40))/2
             .formLabel2.Top=IIF(EMPTY(parlabel2),.formlabel.Top,.Formlabel.Top+.formlabel.Height+5)
             .Cont1.Top=.FormLabel2.Top+40
     ENDCASE
     .AutoCenter=.T.        
ENDWITH
IF parVisible
    FormMy.Show 
ENDIF
DELETE FILE &pathmem
*************************************************************************************************************************
*                                  вставка изображения в форму
*************************************************************************************************************************
PROCEDURE PasteImage
PARAMETERS parFrm,parMove
dataold=''
IF pictset>1
   dataOld=ALIAS()
ENDIF  
pathmem=''
logRepShow=.F.
pictWidth=0
IF pictset>1
   IF !USED('datapic')
      USE datapic IN 0 ORDER 1
   ENDIF
   SELECT datapic
   widthHeight=pWidth/pHeight
   IF pictset>2
      SEEK pictset     
   ENDIF
   SCAN WHILE IIF(pictset>2,theme=pictset,!EOF())
        IF nShow=0.AND.pHeight>=&parFrm..Height.AND.pWidth+&parFrm..Width<SYSMETRIC(21) 
           pictRec=RECNO()
           REPLACE nShow WITH  1 
           logRepShow=.T.                   
           EXIT 
        ENDIF
   ENDSCAN
   IF pictset>2
      IF theme#pictset 
         REPLACE nShow WITH 0 FOR theme=pictset
         LOCATE FOR theme=pictset.AND.pHeight>=&parFrm..Height.AND.pWidth+&parFrm..Width<SYSMETRIC(21) 
         IF FOUND()     
            REPLACE nShow WITH 1
         ELSE  
            LOCATE FOR pHeight>=&parFrm..Height.AND.pWidth+&parFrm..Width<SYSMETRIC(21)  
            REPLACE nShow WITH 1
         ENDIF    
      ENDIF 
   ELSE 
      IF EOF() 
         LOCATE FOR pHeight>=&parFrm..Height.AND.pWidth+&parFrm..Width<SYSMETRIC(21)
         IF FOUND()
            REPLACE nShow WITH 1            
         ELSE
            REPLACE nShow WITH 0 ALL
            GO TOP        
            REPLACE nShow WITH 1
         ENDIF
      ENDIF 
   ENDIF 

   pathmem=FULLPATH('osnproc.fxp')
   pathmem=LEFT(pathmem,LEN(pathmem)-11)+'datapic'+LTRIM(STR(RECNO()))+'.pic'
   widthHeight=pWidth/pHeight
   pictHeight=&parFrm..Height
   pictWidth=&parFrm..Height*widthHeight
   IF pictWidth+&parFrm..Width>=SYSMETRIC(21)
      pictWidth=SYSMETRIC(21)-&parFrm..Width-10  
   ENDIF 
   COPY MEMO pict TO &pathmem
   SELECT datapic
   USE
ENDIF
IF !EMPTY(dataOld)
   SELECT (dataOld)
ENDIF
WITH &parfrm
     .FormImage.Picture=pathmem
     .FormImage.Height=.Height
     .FormImage.Width=pictWidth
     .FormImage.Visible=.T.
     .Width=&parFrm..Width+pictWidth
     .Autocenter=.T.    
     .WindowState=0
ENDWITH  
FOR i=1 TO &parFrm..ControlCount
    obj_move=&parFrm..Controls(i)
    IF obj_move.BaseClass#'Image'
       obj_move.Move(obj_move.Left+IIF(EMPTY(parMove),pictWidth,parMove),obj_move.Top)
    ENDIF
ENDFOR 
IF pictset>1 
   DELETE FILE &pathmem
ENDIF 
*************************************************************************************************************************
*            Процедуоа для перемещения объектов после вставки изображения в некоторых формах
*************************************************************************************************************************
PROCEDURE moveobject
PARAMETERS parFrm,parMove
&& parFrm  - Форма
&& parMove - величина на которую передвигается объект
FOR i=1 TO &parFrm..ControlCount
    obj_move=&parFrm..Controls(i)
    IF obj_move.BaseClass#'Image'
       obj_move.Move(obj_move.Left+parMove,obj_move.Top)
    ENDIF
ENDFOR
*************************************************************************************************************************
*                              Возврат ширины текста в пикселах 
*************************************************************************************************************************
PROCEDURE RetTxtWidth
PARAMETERS parObj,parFontName,parFontSize
parObj=FONTMETRIC(6,IIF(EMPTY(parFontName),dFontName,parFontName),IIF(EMPTY(parFontSize),dFontSize,parFontSize))*;
       TXTWIDTH(parObj,IIF(EMPTY(parFontName),dFontName,parFontName),IIF(EMPTY(parFontSize),dFontSize,parFontSize))
RETURN parObj


*------------------------------------------------------------------------------------------------------------------------
*         Процедура копирования-вставки текста
*------------------------------------------------------------------------------------------------------------------------
PROCEDURE copyins
PARAMETERS parFrm,parObj,parVar,parField
DEFINE POPUP short SHORTCUT RELATIVE FROM MROW(),MCOL() FONT dFontName,dFontSize COLOR SCHEME 4
DEFINE BAR 1 OF short PROMPT 'Вырезать'
DEFINE BAR 2 OF short PROMPT "\-"
DEFINE BAR 3 OF short PROMPT 'Копировать'  
DEFINE BAR 4 OF short PROMPT "\-"
DEFINE BAR 5 OF short PROMPT 'Вставить' SKIP FOR EMPTY(&parVar)
ON SELECTION POPUP short DO proccopy  
ACTIVATE POPUP short
***************************************************************************
PROCEDURE proccopy
men_cx=BAR()
DEACTIVATE POPUP short
DO CASE
   CASE men_cx=1
        &parVar=&parObj..Seltext
        newtext=LEFT(&parfield,&parobj..SelStart)+SUBSTR(&parfield,&parobj..SelStart+1+&parObj..SelLength)
        REPLACE &parField WITH newtext  
        &parfrm..Refresh     
   CASE men_cx=3 
        &parVar=&parObj..Seltext
   CASE men_cx=5
        newtext=LEFT(&parfield,&parobj..SelStart)+&parVar+SUBSTR(&parfield,&parobj..SelStart+1)
        *txtrep=ALLTRIM(&parField)+&parVar 
        REPLACE &parField WITH newtext      
        &parfrm..Refresh        
ENDCASE

**********************************************************************************************************************
PROCEDURE totalsetup
log_set=.F.
IF !USED('colorset')
   USE colorset IN 0
ENDIF 
SELECT colorset
max_sc=RECCOUNT()
DIMENSION dim_sc(max_sc)
STORE 0 TO dim_sc
LOCATE FOR log_scheme
dim_sc(RECNO())=1
COUNT TO max_pic FOR !EMPTY(procset1)
DIMENSION dim_pic(max_pic)
STORE 0 TO dim_pic
dim_pic(pictset)=1
IF !USED('sprsetup')
   SELECT 0
   USE sprsetup
ENDIF   
IF !USED('datapic')
   SELECT 0
   USE datapic ORDER 1   
ENDIF
frmsetup=CREATEOBJECT('FORMMY')
WITH frmsetup
     .Caption='Настройка'
     .procexit='DO exitfromsetup'     
ENDWITH
frmsetup.AddObject('grdset','GridMy')
WITH frmsetup.grdset
     .Top=0
     .Left=0
     .Width=frmsetup.Width     
     .Height=frmsetup.Height/2     
     .ScrollBars=2
     .ColumnCount=7
     .FontSize=dFontSize
     .RecordSourceType=1
     .RecordSource='sprsetup'
     .Column1.ControlSource='sprsetup.kod'
     .Column2.ControlSource='sprsetup.name'
     .Column3.ControlSource='sprsetup.adres'
  *   .Column4.ControlSource='sprbank.name'
     .Column4.ControlSource=''
     .Column5.ControlSource='sprsetup.rs'
     .Column6.ControlSource='sprsetup->unn' 
     .Column4.Alignment=0    
     .Column1.Width=RetTxtWidth(' 1234 ')        
     .Column5.Width=RetTxtWidth(' 1234567890123 ')    
     .Column6.Width=RetTxtWidth(' 123456789 ')    
     .Column7.Width=0 
     .Column2.Width=(.Width-.column1.width-.column5.Width-.Column6.Width-SYSMETRIC(5)-13-7)/2
     .Column3.Width=(.Width-.Column1.Width-.Column5.Width-.Column6.Width-.Column2.Width-SYSMETRIC(5)-13-7)/2
     .Column4.Width=.Width-.Column1.Width-.Column3.Width-.Column5.Width-.Column6.Width-.Column2.Width-SYSMETRIC(5)-13-7
     .Column1.Enabled=.F.
     .Column1.Movable=.F.
     .Column1.ReadOnly=.T.           
     .Column5.Format='Z'
     .Column6.Format='Z'
     .Column1.Header1.Caption='Код'
     .Column2.Header1.Caption='Наименование'
     .Column3.Header1.Caption='Адрес'
     .Column4.Header1.Caption='Банк'
     .Column5.Header1.Caption='Р/с'
     .Column6.Header1.Caption='УНП'       
     .colNesInf=2
     .rowsGrid=(.Height-.HeaderHeight)/.RowHeight
     .SetAll('Resizable',.F.,'Column') 
     .SetAll('BOUND',.F.,'Column')      
     .rowsGrid=(.Height-.HeaderHeight)/.RowHeight 
ENDWITH

DO addShape WITH 'frmSetup',1,5,10,dHeight*max_pic+(max_pic-1)*5+30,(frmsetup.Width-15)/2,8 
frmSetup.Shape1.Top=frmSetup.Height-frmSetup.Shape1.Height-20
DO addShape WITH 'frmsetup',2,frmsetup.Shape1.Left+frmsetup.Shape1.Width+5,frmsetup.Shape1.Top,frmsetup.Shape1.Height,frmsetup.Shape1.Width,8 
frmsetup.SetAll('BorderColor',RGB(192,192,192),'ShapeMy') 

leftObj=(frmsetup.Width-RetTxtWidth('Шрифт')-RetTxtWidth('Размер')-250)/2
DO adLabMy WITH 'frmsetup',1,'Шрифт ',frmsetup.grdSet.Top+frmSetup.grdSet.Height+10,LeftObj,150,0,.T.
DO addcombomy WITH 'frmsetup',1,frmsetup.Lab1.Left+frmsetup.Lab1.Width+10,frmSetup.Shape1.Top-dHeight-20,dHeight,150,.T.,'fontdef(1)','Arial,Times New Roman,Courier New',1,;
   .F.,'DO changeFontSize',.F.,.T.  
frmsetup.Lab1.Top=frmSetup.ComboBox1.Top+(frmSetup.ComboBox1.Height-frmSetup.Lab1.Height)

DO adLabMy WITH 'frmsetup',2,'Размер ',frmsetup.lab1.Top,frmsetup.ComboBox1.Left+frmsetup.combobox1.Width+20,150,0,.T.

DO addspinnermy WITH 'frmSetup','spinsize',frmSetup.lab2.Left+frmSetup.lab2.Width+15,frmsetup.ComboBox1.Top,dHeight,70,'fontdef(2)',1
WITH frmSetup.spinsize
     .spinnerHighValue=14
     .spinnerLowValue=9
     .procForInterActiveChange='DO changeFontSize'
ENDWITH 

frmsetup.grdset.Height=frmsetup.grdSet.Top+frmSetup.ComboBox1.Top-20
DO gridSize WITH 'frmsetup','grdSet','shapeingrid'  

SELECT colorset
GO TOP
ord_cx=1
topOpt=frmSetup.Shape1.Top+15
LeftOpt=frmSetup.Shape1.Left+10
DO WHILE !EOF()
   IF ord_cx=max_pic+1
      topOpt=frmSetup.Shape1.Top+15
      LeftOpt=frmSetup.Option1.Left+frmSetup.Option1.Width+20
   ENDIF
   proc_cx=ALLTRIM(procset)
   &proc_cx
   SKIP
   ord_cx=ord_cx+1
   topOpt=topOpt+dHeight+5        
ENDDO
topOpt=frmSetup.Option1.Top
LeftOpt=frmSetup.Shape2.Left+10
SELECT colorset 
GO TOP
ord_cx=21

DO WHILE !EOF()
   IF !EMPTY(procset1)
      proc_cx=ALLTRIM(procset1)
      &proc_cx
   ENDIF   
   SKIP
*   topOpt=topOpt+dHeight
ENDDO
WITH frmSetup.formImage
     SELECT datapic
     IF pictset#1
        IF pictset>2
           LOCATE FOR theme=pictset
        ENDIF
        pathmem=FULLPATH('osnproc.fxp')
        pathmem=LEFT(pathmem,LEN(pathmem)-11)+'datapic'+LTRIM(STR(RECNO()))+'.pic'       
        COPY MEMO pict TO &pathmem
        SELECT datapic          
     ENDIF   
     .Visible=.T.
     .Left=frmSetup.Option26.Left+frmSetup.Option26.Width+20
     .Top= frmSetup.Option21.Top
     .Height=frmSetup.Shape2.Height-40
     .Width=300     
     .Picture=IIF(pictset=1,'',pathmem)
ENDWITH   
frmsetup.Show
*************************************************************************************************************************
PROCEDURE setupPic 
PARAMETERS parOrd
STORE 0 TO dim_pic
log_set=.T.
dim_pic(parOrd)=1
pictset=parOrd
SELECT datapic
IF pictset#1
   IF pictset>2
      LOCATE FOR theme=pictset
   ENDIF
   pathmem=FULLPATH('osnproc.fxp')
   pathmem=LEFT(pathmem,LEN(pathmem)-11)+'datapic'+LTRIM(STR(RECNO()))+'.pic'
   COPY MEMO pict TO &pathmem
   SELECT datapic  
   frmSetup.formImage.Picture=pathmem        
ELSE
   frmSetup.formImage.Picture=''       
ENDIF   

*     .Left=frmSetup.Option26.Left+frmSetup.Option26.Width+20
*     .Top= frmSetup.Option21.Top
*     .Height=frmSetup.Shape2.Height-40
*     .Width=300     


frmSetup.Refresh
*************************************************************************************************************************
PROCEDURE setupScheme 
PARAMETERS parOrd
log_set=.T.
STORE 0 TO dim_sc
dim_sc(parOrd)=1
SELECT colorset
REPLACE log_scheme WITH .F. ALL
GO parOrd
REPLACE log_scheme WITH .T.
dForeColor=EVALUATE(tForeColor)
dBackColor=EVALUATE(fbackColor)
headerForeColor=EVALUATE(hforecolor)
headerBackColor=EVALUATE(hbackcolor)
dynForeColor=EVALUATE(dynamFore)
dynBackColor=EVALUATE(dynamBack)
ObjBackColor=EVALUATE(tBackColor)
ObjForeColor=EVALUATE(tForeColor)
objColorSos=RGB(255,0,0)
selBackcolor=IIF(!EMPTY(selBack),EVALUATE(selBack),selBackColor)
frmSetup.BackColor=dBackColor
frmSetup.SetAll('BackColor',dBackColor,'ComboMy')
frmSetup.SetAll('BackColor',dBackColor,'MyOptionButton')
frmSetup.SetAll('BackColor',dBackColor,'MySpinner')
frmSetup.Refresh
*************************************************************************************************************************
PROCEDURE changeFontSize
fontdef(2)=frmSetup.spinSize.Value
dFontSize=fontdef(2)
fontdef(1)=frmsetup.ComboBox1.Value
dFontName=fontdef(1)
log_set=.T.
WITH frmsetup.grdset
     .FontSize=fontdef(2) 
     .Top=0
     .Left=0
     .Width=frmsetup.Width          
     .Height=frmsetup.grdSet.Top+frmSetup.ComboBox1.Top-20   
     .ScrollBars=2
     .ColumnCount=7
     .FontSize=dFontSize
     .FontName=dFontName
     .RecordSourceType=1
     .RecordSource='sprsetup'
     .Column1.ControlSource='sprsetup.kod'
     .Column2.ControlSource='sprsetup.name'
     .Column3.ControlSource='sprsetup.adres'
  *   .Column4.ControlSource='sprbank.name'
     .Column4.ControlSource=''
     .Column5.ControlSource='sprsetup.rs'
     .Column6.ControlSource='sprsetup->unn' 
     .Column4.Alignment=0    
     .Column1.Width=RetTxtWidth(' 1234 ')        
     .Column5.Width=RetTxtWidth(' 1234567890123 ')    
     .Column6.Width=RetTxtWidth(' 123456789 ')    
     .Column7.Width=0 
     .Column2.Width=(.Width-.column1.width-.column5.Width-.Column6.Width-SYSMETRIC(5)-13-7)/2
     .Column3.Width=(.Width-.Column1.Width-.Column5.Width-.Column6.Width-.Column2.Width-SYSMETRIC(5)-13-7)/2
     .Column4.Width=.Width-.Column1.Width-.Column3.Width-.Column5.Width-.Column6.Width-.Column2.Width-SYSMETRIC(5)-13-7
     .Column1.Enabled=.F.
     .Column1.Movable=.F.
     .Column1.ReadOnly=.T.           
     .Column5.Format='Z'
     .Column6.Format='Z'
     .Column1.Header1.Caption='Код'
     .Column2.Header1.Caption='Наименование'
     .Column3.Header1.Caption='Адрес'
     .Column4.Header1.Caption='Банк'
     .Column5.Header1.Caption='Р/с'
     .Column6.Header1.Caption='УНП'       
     .colNesInf=2
     .rowsGrid=(.Height-.HeaderHeight)/.RowHeight
     .SetAll('Resizable',.F.,'Column') 
     .SetAll('BOUND',.F.,'Column')      
     .rowsGrid=(.Height-.HeaderHeight)/.RowHeight 
ENDWITH
DO gridSize WITH 'frmsetup','grdSet'  
frmSetup.Refresh
***************************************************************************************************************************
*                      Выход из настроек
***************************************************************************************************************************
PROCEDURE exitFromSetup
IF log_set
   DO createFormNew WITH .T.,'Настройки',RetTxtWidth('WWДля вступления изменений в силуWW',dFontName,dFontSize+1),;
     '130',RetTxtWidth('WНетW',dFontName,dFontSize+1),'ОК',.F.,.F.,;
     'nFormMes.Release',.F.,.F.,'Для вступления изменений в силу','перезапустите программу заново!',.T.
    var_path=FULLPATH('fontdef.mem')
    SAVE TO &var_path ALL LIKE fontdef
    var_path=FULLPATH('pictset.mem')
    SAVE TO &var_path ALL LIKE pictset
ENDIF     


**************************************************************************************************************************
*        Процедура добавления кнопок для редактирования справочников
**************************************************************************************************************************
PROCEDURE addmenureadspr
PARAMETERS parfrm,parwrite,parexit

DO addButtonOne WITH parfrm,'menuRead',10,5,'записать','pencil.ico',parwrite,39,RetTxtWidth('удаление')+44,'записать'  
DO addButtonOne WITH parfrm,'menuexit',&parfrm..menuread.Left+&parfrm..menuread.Width+5,5,'возврат','undo.ico',parexit,39,.menucont1.Width,'возврат'  

*DO addcontmenu WITH parfrm,'menuread',10,5,'записать','pencil.ico',parwrite
*DO addcontmenu WITH parfrm,'menuexit',&parfrm..menuread.Left+&parfrm..menuread.Width+5,5,'возврат','undo.ico',parexit
&parfrm..menuread.Visible=.F.
&parfrm..menuexit.Visible=.F.

**************************************************************************************************************************
PROCEDURE readspr
PARAMETERS parfrm,parread
&parfrm..SetAll('Visible',.F.,'mymenucont')
&parfrm..SetAll('Visible',.F.,'myCommandButton')
&parfrm..menuread.Visible=.T.
&parfrm..menuexit.Visible=.T.
&parread
**************************************************************************************************************************
PROCEDURE writespr
PARAMETERS parfrm,pargrd,parbase,parproc,parfield
&parfrm..SetAll('Visible',.T.,'mymenucont')
&parfrm..SetAll('Visible',.T.,'myCommandButton')
&parfrm..menuread.Visible=.F.
&parfrm..menuexit.Visible=.F.
SELECT &parbase
GATHER FROM &parfrm..dim_ap
IF !EMPTY(parproc)
   DO &parproc
ENDIF   
&parfrm..SetAll('Visible',.F.,'MyTxtBox')
&parfrm..SetAll('Visible',.F.,'ComboMy')
&parfrm..SetAll('Visible',.F.,'MySpinner')
&parfrm..SetAll('Visible',.F.,'checkContainer')
&pargrd..Enabled=.T.
SELECT &parbase
&pargrd..GridUpdate WITH parfield
GO &parfrm..nrec
&pargrd..SetAll('Enabled',.F.,'Column')
&pargrd..Columns(&pargrd..ColumnCount).Enabled=.T.
GO &parfrm..nrec
**************************************************************************************************************************
PROCEDURE exitwrite
PARAMETERS parfrm,pargrd
&parfrm..SetAll('Visible',.T.,'mymenucont')
&parfrm..SetAll('Visible',.T.,'myCommandButton')
&parfrm..menuread.Visible=.F.
&parfrm..menuexit.Visible=.F.
&parfrm..SetAll('Visible',.F.,'mytxtbox')
&parfrm..SetAll('Visible',.F.,'ComboMy')
&parfrm..SetAll('Visible',.F.,'MySpinner')
&parfrm..SetAll('Visible',.F.,'checkContainer')
&pargrd..Enabled=.T.
&pargrd..Setall('Enabled',.F.,'Column')
&pargrd..Setall('Enabled',.F.,'ColumnMy')
&pargrd..Columns(&pargrd..ColumnCount).Enabled=.T.
&pargrd..GridUpdate
&parfrm..Refresh
LOCATE FOR &parfrm..nrec=RECNO()
IF !FOUND()
   GO BOTTOM
ENDIF
**************************************************************************************************************************
PROCEDURE writeSprNew
PARAMETERS parfrm,pargrd,parbase,parproc
&parfrm..SetAll('Visible',.T.,'mymenucont')
&parfrm..SetAll('Visible',.T.,'myCommandButton')
&parfrm..menuread.Visible=.F.
&parfrm..menuexit.Visible=.F.
SELECT &parbase
GATHER FROM &parfrm..dim_ap
IF !EMPTY(parproc)
   DO &parproc
ENDIF   
&parfrm..SetAll('Visible',.F.,'MyTxtBox')
&parfrm..SetAll('Visible',.F.,'ComboMy')
&parfrm..SetAll('Visible',.F.,'MySpinner')
&parfrm..SetAll('Visible',.F.,'checkContainer')
&pargrd..Enabled=.T.
SELECT &parbase
&pargrd..GridUpdate
GO &parfrm..nrec
&pargrd..SetAll('Enabled',.F.,'ColumnMy')
&pargrd..Columns(&pargrd..ColumnCount).Enabled=.T.
GO &parfrm..nrec
**************************************************************************************************************************
PROCEDURE exitWriteSpr
PARAMETERS parfrm,pargrd
&parfrm..SetAll('Visible',.T.,'mymenucont')
&parfrm..SetAll('Visible',.T.,'myCommandButton')
&parfrm..menuread.Visible=.F.
&parfrm..menuexit.Visible=.F.
&parfrm..SetAll('Visible',.F.,'mytxtbox')
&parfrm..SetAll('Visible',.F.,'ComboMy')
&parfrm..SetAll('Visible',.F.,'MySpinner')
&parfrm..SetAll('Visible',.F.,'checkContainer')
&pargrd..Enabled=.T.
&pargrd..Setall('Enabled',.F.,'ColumnMy')
&pargrd..Columns(&pargrd..ColumnCount).Enabled=.T.
&pargrd..GridUpdate
&parfrm..Refresh
LOCATE FOR &parfrm..nrec=RECNO()
IF !FOUND()
   GO BOTTOM
ENDIF
****************************************************************************************************************************
*
****************************************************************************************************************************
PROCEDURE printreport
PARAMETERS parReport,parCaption,parBase,parFrmVisible,parWord,parExcel,parProc
fSupl=CREATEOBJECT('Formsupl')
nforreport=parReport
nforcaption=parCaption
parForPrint=parReport
logFrmVisible=.F.
frmVisible=''
IF !EMPTY(parFrmVisible)
   frmVisible=parFrmVisible
   &frmVisible..Visible=.F.  
   logFrmVisible=.T.
ENDIF
WITH fSupl      
     .procexit='DO exitprintreport'
     .logExit=.T.
     .Caption=parCaption
     DO adSetupPrnToForm WITH 20,20,400,IIF(parWord,.T.,.F.),IIF(parExcel,.T.,.F.)


     *---------------------------------Кнопка печать-------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',.Shape91.Left+(.Shape91.Width-RetTxtWidth('WПросмотрW')*3-40)/2,.Shape91.Top+.Shape91.Height+20,;
        RetTxtWidth('WПросмотрW'),dHeight+5,'Печать','DO printForReport' ,'Печать ведомости'
     *---------------------------------Кнопка предварительного просмотра-----------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont2',.cont1.Left+.Cont1.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Просмотр',"DO previewrep WITH nforreport,nforcaption",'Предварительный просмотр и печать ведомости'
     *-------------------------------------Кнопка выход из формы печати----------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont3',.cont2.Left+.Cont2.Width+20,.Cont1.Top,.Cont1.Width,dHeight+5,'Выход','DO exitprintreport','Выход из печати'        
      
     .Height=.Shape91.Height+.cont1.Height+60
     .Width=.Shape91.Width+40   
ENDWITH
SELECT &parBase
GO TOP
DO pasteImage WITH 'fSupl'
fSupl.cont1.SetFocus
fSupl.Show
************************************************************************************************************************************
PROCEDURE printforreport
IF logWord
   &parProc
ELSE 
   FOR i=1 TO kvo_page
       DO CASE
          CASE dimCht(1)=1
               Report Form &parForPrint NOCONSOLE TO PRINTER RANGE page_beg,page_end
          CASE dimCht(2)=1     
               FOR c_range=page_beg TO page_end         
                   IF MOD(c_range,2)=0                
                      Report Form &parForPrint RANGE c_range,c_range NOCONSOLE TO PRINTER                         
                      IF EOF()
                         EXIT 
                      ENDIF  
                   ENDIF    
               ENDFOR      
          CASE dimCht(3)=1     
               FOR c_range=page_beg TO page_end         
                   IF MOD(c_range,2)#0                
                      Report Form &parForPrint RANGE c_range,c_range NOCONSOLE TO PRINTER                         
                      IF EOF()
                         EXIT 
                      ENDIF  
                   ENDIF    
               ENDFOR      
       ENDCASE 
   ENDFOR
ENDIF    
******************************************************************************************************************************
PROCEDURE exitprintreport
IF logFrmVisible
   &frmVisible..Visible=.T.
ENDIF
fSupl.Release

***************************************************************************************************************************
*            Процедура добавления ComboBox для выбора принтера
***************************************************************************************************************************
PROCEDURE addComboPrn
PARAMETERS parFrm,parInd,parTop,parLeft,parWidth
DO addcombomy WITH parFrm,parInd,parLeft,parTop,dHeight,parWidth,.T.,'nameprint',;
   'name_prn',5,.F.,' SET PRINTER TO NAME name_prn(ASCAN(name_prn,nameprint))',.F.,.T.  
****************************************************************************************************************************
PROCEDURE mypassword
CLOSE ALL
fdop=CREATEOBJECT('FORMMY')
pasdop=''
WITH fdop 
     DO adlabmy WITH 'fdop',1,'Данная процедура предназначена для программиста',10,10,RetTxtWidth('Данная процедура предназначена для программиста'),2,.F.
     DO adlabmy WITH 'fdop',2,'или опытного ползователя.',fdop.lab1.Top+fdop.lab1.Height+3,10,RetTxtWidth('Данная процедура предназначена для программиста'),2,.F.
     DO adlabmy WITH 'fdop',3,'Пожалуйста, введите пароль',fdop.lab2.Top+fdop.lab2.Height+3,10,RetTxtWidth('Данная процедура предназначена для программиста'),2,.F.
     DO adTboxNew WITH 'fdop','txtbox1',.lab3.Top+.lab3.Height+5,.lab1.Left+(.lab1.Width-150)/2,150,dHeight,'pasdop',.F.,.T.,2,.F.,'DO validpassword',.F.,.F.,'DO exitfrompassword'
     .txtbox1.PasswordChar='*'
     .txtbox1.maxlength=8    
     .Caption='Настройка печати'
     .BackColor=RGB(255,255,255)
     .Height=.lab1.Height*3+dheight+40
     .Width=.lab1.Width+20
     .MinButton=.F.
     .MaxButton=.F.     
     .AutoCenter=.T.
     .WindowState=0    
ENDWITH 
DO pasteimage WITH 'fdop'
fdop.Show
********************************************************
PROCEDURE exitfrompassword
IF LASTKEY()=27
   fdop.Release
   RETURN TO proctop
ENDIF
**************************************************************************************************************************************************
PROCEDURE validpassword
IF ALLTRIM(pasdop)#'lozdm'
   fdop.Release
   RETURN TO proctop
ENDIF
fdop.Release
*-----------------------------------------------------------------------
*        Представление бухгалтерского счета в нормальном виде
*-----------------------------------------------------------------------
PROCEDURE mySayScore
LPARAMETERS parName,parPoint
 RETURN ALLTRIM(LEFT(parName,3))+SUBSTR(parName,5,3)
 
 
 *   RETURN ALLTRIM(LEFT(parName,3))+IIF(! EMPTY(SUBSTR(parName,5,3)),'.'+;
 *         ALLTRIM(SUBSTR(parName,5,3)),'')+IIF(! EMPTY(SUBSTR(parName,9)),'.'+;
 *         ALLTRIM(SUBSTR(parName,9)),'')*


*********************************************************************************************************************************
*    Построение окна для Excel
*********************************************************************************************************************************
PROCEDURE createWindowMsWord
oWord=GETOBJECT('C:\PROGRAM fILES\Microsoft Office\office11\word.exe',"Word.Application")
oWord.Application.Visible=.T.
oWord.Caption='Новый документ'
oWord.Documents.Add
WITH oWord.ActiveDocument.PageSetup
     .Orientation=0
ENDWITH
WITH oWord.Selection
     .TypeText('Построение окна Microsoft Word для того, чтобы пользователь имел возможностиь редактировать документ')
     .Typeparagraph
ENDWITH
oWord.documents.tables.add
*oWord.Application.Top = 65
*oWord.Application.Left = 10
*oWord.Application.Height = 470
*oWord.Application.Width = 290
*********************************************************************************************************************************
*    Построение окна для Word
*********************************************************************************************************************************
PROCEDURE createWindowMsExcel
oExcel=GETOBJECT('C:\PROGRAM fILES\Microsoft Office\office11\excel.exe',"Excel.Application")
oExcel.Application.Visible=.T.
*oExcel.Application.Top = 1
*oExcel.Application.Left = 1
*oExcel.Application.Height = SYSMETRIC(1)
*oExcel.Application.Width = SYSMETRIC(2)
*******************************************************************************************************************************************************
*                                Календарь и праздничные дни
*******************************************************************************************************************************************************
PROCEDURE proccalendar
IF !USED('fete')
   USE fete ORDER 1 IN 0 
ENDIF
fClnd=CREATEOBJECT('FORMMY')
month_ch=MONTH(DATE())
year_ch=YEAR(DATE())
newfete=''
log_fete=.F.
DIMENSION dim_day(42)
STORE '   ' TO dim_day
WITH fClnd
     .BackColor=RGB(255,255,255)    
     .Caption='Календарь'
     
     DO addcombomy WITH 'fClnd',1,20,10,dHeight,RetTxtWidth(' сентябрь ')+SYSMETRIC(5),.T.,'month_ch','dim_month',5,;
       .F.,'DO calendarscr',.F.,.T.                            
     fClnd.comboBox1.DisplayCount=12   
     
     DO addspinnermy WITH 'fClnd','spin_year',fClnd.comboBox1.Left+fClnd.comboBox1.Width+10,fClnd.comboBox1.Top,dHeight,RetTxtWidth('W9999W'),'year_ch',1,.F.,1900,3000
     fClnd.spin_year.procForInterActiveChange='DO calendarscr'                
     DO addcontmy WITH 'fClnd','contMonday',5,fClnd.comboBox1.Top+fClnd.comboBox1.Height+10,RetTxtWidth('wпонедельникw'),dHeight,'понедельник',.F.,1
     DO addcontmy WITH 'fClnd','contTuesday',fClnd.contMonday.Left+fClnd.contMonday.Width-1,fClnd.contMonday.Top,fClnd.contMonday.Width,dHeight,'вторник',.F.,1
     DO addcontmy WITH 'fClnd','contWednesday',fClnd.contTuesday.Left+fClnd.contMonday.Width-1,fClnd.contMonday.Top,fClnd.contMonday.Width,dHeight,'среда',.F.,1
     DO addcontmy WITH 'fClnd','contThursday',fClnd.contWednesday.Left+fClnd.contMonday.Width-1,fClnd.contMonday.Top,fClnd.contMonday.Width,dHeight,'четверг',.F.,1
     DO addcontmy WITH 'fClnd','contFriday',fClnd.contThursday.Left+fClnd.contMonday.Width-1,fClnd.contMonday.Top,fClnd.contMonday.Width,dHeight,'пятница',.F.,1
     DO addcontmy WITH 'fClnd','contSaturday',fClnd.contFriday.Left+fClnd.contMonday.Width-1,fClnd.contMonday.Top,fClnd.contMonday.Width,dHeight,'суббота',.F.,1
     DO addcontmy WITH 'fClnd','contSunday',fClnd.contSaturday.Left+fClnd.contMonday.Width-1,fClnd.contMonday.Top,fClnd.contMonday.Width,dHeight,'воскресенье',.F.,1                            
     leftObj=fClnd.contMonday.Left
     topObj=fClnd.contMonday.Top+dHeight-1 
     widthObj=fClnd.contMonday.Width
     heightObj=dHeight*2
     objcx=1
     kvoDay=IIF(month_ch=2,IIF(MOD(year_ch,4)=0,29,28),IIF(INLIST(month_ch,4,6,9,11),30,31))
     dayOne=DOW(CTOD('01.'+STR(month_ch,2)+'.'+STR(year_ch,4)))
     dayOne=IIF(dayOne=1,7,dayOne-1)
     day_ch=1
     FOR i=1 TO 42
         DO CASE
          CASE i<dayOne               
          CASE i>=dayOne.AND.day_ch<=kvoDay
               dim_day(i)=LTRIM(STR(day_ch))         
               day_ch=day_ch+1
          CASE i>day_ch               
       ENDCASE    
     ENDFOR   
     FOR i=1 TO 42                                 
         contname='cont'+LTRIM(STR(objcx))                                         
         DO addcontmy WITH 'fClnd',contname,leftObj,topObj,widthObj,heightObj,dim_day(i),.F.,1                        
         fClnd.&contname..procForRightClick='DO procInputFete WITH fClnd.&contname..contlabel.Caption'          
         leftObj=leftObj+widthObj-1
         objcx=objcx+1               
         IF MOD(i+1,7)=0.OR.MOD(i,7)=0.OR.SEEK(VAL(dim_day(i))+month_ch/100,'fete',1)                                  
            fClnd.&contname..contlabel.ForeColor=RGB(255,00,0)    
         ELSE
            fClnd.&contname..contlabel.ForeColor=objForeColor           
         ENDIF           
         IF SEEK(VAL(dim_day(i))+month_ch/100,'fete',1)                                         
            fClnd.&contname..contlabel.toolTipText=ALLTRIM(fete.comment)
         ELSE
            fClnd.&contname..contlabel.toolTipText=''           
         ENDIF          
         fClnd.&contname..BackColor=IIF(CTOD(dim_day(i)+'.'+STR(month_ch,2)+'.'+STR(year_ch,4))#DATE(),dBackColor,fClnd.contMonday.backColor)                             
         IF MOD(i,7)=0                                               
            leftObj=fClnd.contMonday.Left
            topObj=topObj+heightObj-1 
         ENDIF
     ENDFOR           
     .Height=heightObj*6+fClnd.ComboBox1.Height+fClnd.contMonday.Height+30
     .Width=fClnd.contMonday.Width*7+3
     .comboBox1.Left=(.Width-.comboBox1.Width-.spin_year.Width-10)/2
     .spin_year.Left=.comboBox1.Left+.comboBox1.Width+10
     .MinButton=.F.
     .MaxButton=.F.    
     .AlwaysOnTop=.T. 
     .AutoCenter=.T.
     .WindowState=0
ENDWITH
DO pasteimage WITH 'fClnd'
fClnd.Show
**********************************************************************************************************************************************************
PROCEDURE calendarscr
year_ch=fClnd.spin_year.Value
leftObj=fClnd.contMonday.Left
topObj=fClnd.contMonday.Top+dHeight-1 
widthObj=fClnd.contMonday.Width
heightObj=dHeight*2
objcx=1
kvoDay=IIF(month_ch=2,IIF(MOD(year_ch,4)=0,29,28),IIF(INLIST(month_ch,4,6,9,11),30,31))
dayOne=DOW(CTOD('01.'+STR(month_ch,2)+'.'+STR(year_ch,4)))
dayOne=IIF(dayOne=1,7,dayOne-1)
day_ch=1
STORE '   ' TO dim_day
FOR i=1 TO 42
    DO CASE
       CASE i<dayOne
       CASE i>=dayOne.AND.day_ch<=kvoDay
            dim_day(i)=LTRIM(STR(day_ch ))
            day_ch=day_ch+1
       CASE i>day_ch           
    ENDCASE    
ENDFOR 
FOR i=1 TO 42        
    contname='fClnd.cont'+LTRIM(STR(objcx))       
    &contname..contlabel.Caption=dim_day(i) 
    IF MOD(i+1,7)=0.OR.MOD(i,7)=0.OR.SEEK(VAL(dim_day(i))+month_ch/100,'fete',1)                                  
       &contname..contLabel.ForeColor=RGB(255,00,0)        
    ELSE
       &contname..contLabel.ForeColor=objForeColor           
    ENDIF  
    IF SEEK(VAL(dim_day(i))+month_ch/100,'fete',1)                                         
       &contname..contLabel.toolTipText=ALLTRIM(fete.comment)
    ELSE
       &contname..contLabel.toolTipText=''           
    ENDIF   
    &contname..BackColor=IIF(CTOD(dim_day(i)+'.'+STR(month_ch,2)+'.'+STR(year_ch,4))#DATE(),dBackColor,fClnd.contMonday.backColor)    
    &contname..Refresh
    leftObj=leftObj+widthObj-1
    objcx=objcx+1
    IF MOD(i,7)=0
       leftObj=fClnd.contMonday.Left
       topObj=topObj+heightObj-1 
     ENDIF
 ENDFOR    
***********************************************************************************************************************************************
PROCEDURE procInputFete
PARAMETERS par1
IF EMPTY(par1)
   RETURN
ENDIF
log_fete=IIF(SEEK(VAL(par1)+month_ch/100,'fete',1),.T.,.F.)
DEFINE POPUP short SHORTCUT RELATIVE FROM MROW(),MCOL() FONT dFontName,dFontSize COLOR SCHEME 4
DEFINE BAR 1 OF short PROMPT IIF(log_fete,'Редактировать','Добавить')
DEFINE BAR 2 OF short PROMPT "\-"
DEFINE BAR 3 OF short PROMPT 'Удалить'  SKIP FOR !log_fete   
ON SELECTION POPUP short DO procreadfete  
ACTIVATE POPUP short
***************************************************************************
PROCEDURE procreadfete
men_cx=BAR()
DEACTIVATE POPUP short
DO CASE
   CASE men_cx=1
        newfete=IIF(SEEK(VAL(par1)+month_ch/100,'fete',1),fete.comment,SPACE(35))
        fFete=CREATEOBJECT('FORMMY')     
        feteSay=par1+' '+IIF(INLIST(month_ch,3,8),dim_month(month_ch),LEFT(dim_month(month_ch),LEN(dim_month(month_ch))-1))+IIF(INLIST(month_ch,3,8),'а','я')+' '+STR(year_ch,4)
        WITH fFete
             .BackColor=RGB(255,255,255)    
             .Caption='Праздничный день'            
             .Width=260 
             DO adlabmy WITH 'fFete',1,feteSay,5,0,fFete.Width,2,.F.           
             DO adtbox WITH 'fFete',1,5,fFete.lab1.Top+fFete.lab1.Height+10,250,dHeight,'newfete',.F.,.T.,0             
             DO addcontlabel WITH 'fFete','cont1',(fFete.txtBox1.Width-RetTxtWidth('WЗаписатьW')*2-20)/2,fFete.txtBox1.Top+fFete.txtBox1.Height+10,;
                RetTxtWidth('WЗаписатьW'),dHeight+5,'записать','DO saveFete'          
             DO addcontlabel WITH 'fFete','cont2',fFete.cont1.Left+fFete.Cont1.Width+20,fFete.Cont1.Top,;
                fFete.Cont1.Width,dHeight+5,'отказ','fFete.Release' 
             .Height=.txtBox1.Height+.cont1.Height+.lab1.Height+40                                     
             .MinButton=.F.
             .MaxButton=.F.    
             .AlwaysOnTop=.T. 
             .AutoCenter=.T.
             .WindowState=0
        ENDWITH
        fFete.Show    
   CASE men_cx=3 
        SELECT fete
        SEEK VAL(par1)+month_ch/100
        DELETE        
ENDCASE
DO calendarscr
*****************************************************************************************************************************************************
PROCEDURE saveFete
SELECT fete
DO CASE
   CASE log_fete.AND.!EMPTY(newfete)
        SEEK VAL(par1)+month_ch/100
        REPLACE comment WITH newfete
   CASE !log_fete
        APPEND BLANK
        REPLACE comment WITH newfete,datafet WITH VAL(par1)+month_ch/100
ENDCASE
fFete.Release
*-----------------------------------------------------------------------------------------------------------------------------------------------------
*                                      Общие справочники для кадров и штатного расписания
*-----------------------------------------------------------------------------------------------------------------------------------------------------
******************************************************************************************************************************************************
*                                       Справочник подразделений
*******************************************************************************************************************************************************
PROCEDURE mypodr
CLOSE ALL
USE sprpodr ORDER 1 IN 0
USE rasp ORDER 1 IN 0
fpodr=CREATEOBJECT('Formspr')
WITH fpodr
     .Icon='spr.ico'
     .Caption='Справочник подразделений' 
     .ProcExit='fpodr.fgrid.GridReturn'    
ENDWITH
DO addmenureadspr WITH 'fpodr',"DO writespr WITH 'fpodr','fpodr.fGrid','sprpodr'","DO exitwrite WITH 'fpodr','fpodr.fGrid'"
DO addcontmenu WITH 'fpodr','menucont1',10,5,'новая','newrec.bmp',"Do readspr WITH 'fpodr','Do readmypodr WITH .T.'"
DO addcontmenu WITH 'fpodr','menucont2',fpodr.menucont1.Left+fpodr.menucont1.Width+3,5,'редакция','read.bmp',"Do readspr WITH 'fpodr','Do readmypodr WITH .F.'"
DO addcontmenu WITH 'fpodr','menucont3',fpodr.menucont2.Left+fpodr.menucont2.Width+3,5,'удаление','del.bmp','Do delMyPodr'
DO addcontmenu WITH 'fpodr','menucont4',fpodr.menucont3.Left+fpodr.menucont3.Width+3,5,'печать','printer.bmp',"DO printreport WITH 'reppodr','справочник подразделений','sprpodr'"
DO addcontmenu WITH 'fpodr','menucont5',fpodr.menucont4.Left+fpodr.menucont4.Width+3,5,'выход','exit.bmp','fpodr.fGrid.GridReturn' 
WITH fpodr.fGrid
     .Top=fpodr.menucont1.Top+fpodr.menucont1.Height+5
     .Height=fpodr.Height-fpodr.menucont1.Height-5        
     .ColumnCount=4     
     .RecordSourceType=1
     .RecordSource='sprpodr'
     .Column1.ControlSource='sprpodr->kod'
     .Column2.ControlSource='" "+sprpodr->name' 
     .Column3.ControlSource='" "+sprpodr->prim'  
     .Column1.Header1.Caption='Код'
     .Column2.Header1.Caption='Наименование' 
     .Column3.Header1.Caption='Примечание'      
     .Column1.Width=RettxtWidth(' 1234 ')
     .Column4.Width=0
     .Column2.Width=(.Width-.Column1.Width)/2
     .Column3.Width=.Width-.column1.Width-.Column2.Width-SYSMETRIC(5)-13-4        
     .colNesInf=2   
     .SetAll('Movable',.F.,'Column') 
     .SetAll('BOUND',.F.,'Column')        
ENDWITH   
DO gridSize WITH 'fpodr','fGrid','shapeingrid'  
DO addtxtboxmy WITH 'fpodr',1,1,1,fpodr.fGrid.Column1.Width+2,.F.,.F.,1
DO addtxtboxmy WITH 'fpodr',2,1,1,fpodr.fGrid.Column2.Width+2,.F.,.F.,0
DO addtxtboxmy WITH 'fpodr',3,1,1,fpodr.fGrid.Column3.Width+2,.F.,.F.,0
fpodr.SetAll('Visible',.F.,'MyTxtBox')
DO addcontmy WITH 'fpodr','cont1',fpodr.fGrid.Left+13,fpodr.fGrid.Top+2,fpodr.fGrid.Column1.Width-3,;
   fpodr.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fpodr','fpodr.cont1','sprpodr',1,3"
fpodr.cont1.SpecialEffect=1    
DO addcontmy WITH 'fpodr','cont2',fpodr.cont1.Left+fpodr.fGrid.Column1.Width+1,fpodr.fGrid.Top+2,;
   fpodr.fGrid.Column2.Width-3,fpodr.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fpodr','fpodr.cont2','sprpodr',2,4"  
DO addcontmy WITH 'fpodr','cont3',fpodr.cont2.Left+fpodr.fGrid.Column2.Width+1,fpodr.fGrid.Top+2,;
   fpodr.fGrid.Column3.Width-3,fpodr.fGrid.HeaderHeight-3,''      
fpodr.Show
*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readmypodr
PARAMETERS parlog
SELECT sprpodr
IF parlog
   fpodr.fGrid.GridMyAppendBlank(1,'kod','name')   
ENDIF
fpodr.SetAll('Visible',.T.,'MyTxtBox')
fpodr.nrec=RECNO()
SCATTER TO fpodr.dim_ap  
fpodr.txtBox1.Left=fpodr.fGrid.Left+10
fpodr.txtBox2.Left=fpodr.txtbox1.Left+fpodr.txtbox1.Width-1
fpodr.txtBox3.Left=fpodr.txtbox2.Left+fpodr.txtbox2.Width-1
fpodr.txtbox1.ControlSource='fpodr.dim_ap(1)'
fpodr.txtbox2.ControlSource='fpodr.dim_ap(2)'
fpodr.txtbox3.ControlSource='fpodr.dim_ap(3)'
lineTop=fpodr.fGrid.Top+fpodr.fGrid.HeaderHeight+fpodr.fGrid.RowHeight*(IIF(fpodr.fGrid.RelativeRow<=0,1,fpodr.fGrid.RelativeRow)-1)
fpodr.SetAll('Top',linetop,'MyTxtBox')
fpodr.SetAll('Height',fpodr.fGrid.RowHeight+1,'MyTxtBox')
fpodr.SetAll('BackStyle',1,'MyTxtBox')
fpodr.txtbox1.Enabled=.F.
fpodr.fGrid.Enabled=.F.
fpodr.txtbox2.SetFocus
*************************************************************************************************************************
PROCEDURE delMyPodr
fpodr.Setall('BorderWidth',0,'Mycontmenu')
IF SEEK(sprpodr->kod,'rasp',3) 
   fpodr.fGrid.GridNoDelRec   
ELSE 
  fpodr.fGrid.GridDelRec('fpodr.fGrid','sprpodr') 
ENDIF   
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-**-**-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
*-------------------------------------------------------------------------------------------------------------------------
*                                       Справочник должностей
*-------------------------------------------------------------------------------------------------------------------------
PROCEDURE mydolj
CLOSE ALL
USE sprdolj ORDER 1
USE rasp ORDER 1 IN 2
USE sprkat ORDER 1 IN 3
SELECT sprdolj
SET RELATION TO kat INTO sprkat
fdolj=CREATEOBJECT('Formspr')
WITH fdolj
     .Caption='Справочник должностей'  
     .ProcExit='fdolj.fGrid.GridReturn'      
ENDWITH
*DO addmenureadspr WITH 'fdolj',"DO writespr WITH 'fdolj','fdolj.fGrid','sprdolj'","DO exitwrite WITH 'fdolj','fdolj.fGrid'"
DO addmenureadspr WITH 'fdolj','DO writemydolj WITH .T.','DO writemydolj WITH .F.'
DO addcontmenu WITH 'fdolj','menucont1',10,5,'новая','newrec.bmp',"Do readspr WITH 'fdolj','Do readmydolj WITH .T.'"
DO addcontmenu WITH 'fdolj','menucont2',fdolj.menucont1.Left+fdolj.menucont1.Width+3,5,'редакция','read.bmp',"Do readspr WITH 'fdolj','Do readmydolj WITH .F.'"
DO addcontmenu WITH 'fdolj','menucont3',fdolj.menucont2.Left+fdolj.menucont2.Width+3,5,'удаление','del.bmp','DO delmydolj'
DO addcontmenu WITH 'fdolj','menucont4',fdolj.menucont3.Left+fdolj.menucont3.Width+3,5,'печать','printer.bmp',"DO printreport WITH 'repdolj','справочник должностей','sprdolj'"
DO addcontmenu WITH 'fdolj','menucont5',fdolj.menucont4.Left+fdolj.menucont4.Width+3,5,'выход','exit.bmp','fdolj.fGrid.GridReturn'  
WITH fdolj.fGrid
     .Top=fdolj.menucont1.Top+fdolj.menucont1.Height+5
     .Height=fdolj.Height-fdolj.menucont1.Height-5       
     .ColumnCount=5        
     .RecordSourceType=1     
     .RecordSource='sprdolj'
     .Column1.ControlSource='sprdolj.kod'
     .Column2.ControlSource='" "+sprdolj.name'
     .Column3.ControlSource='" "+sprkat.name'  
     .Column4.ControlSource='" "+sprdolj.prim'     
     .Column1.Width=RettxtWidth(' 1234 ')     
     .Column3.Width=RettxtWidth(' СРЕДНИЙ МЕДПЕРСОНАЛ ')    
     .Column1.Header1.Caption='Код'
     .Column2.Header1.Caption='Наименование'
     .Column3.Header1.Caption='Персонал'   
     .Column4.Header1.Caption='Примечание'   
     .Columns(.ColumnCount).Width=0
     .Column2.Width=(.Width-.column1.width-.column3.Width)/2
     .Column4.Width=.Width-.column1.width-.Column2.Width-.column3.Width-SYSMETRIC(5)-13-.ColumnCount
     .Column1.Movable=.F.            
     .Column3.Alignment=0
     .colNesInf=2      
     .SetAll('BOUND',.F.,'Column')  
     .Visible=.T.         
ENDWITH
DO gridSize WITH 'fdolj','fGrid','shapeingrid'   
DO addtxtboxmy WITH 'fdolj',1,1,1,fdolj.fGrid.Column1.Width+2,.F.,.F.,1
fdolj.txtbox1.Enabled=.F.
DO addtxtboxmy WITH 'fdolj',2,1,1,fdolj.fGrid.Column2.Width+2,.F.,.F.,0
DO addcombomy WITH 'fdolj',3,1,1,fdolj.fGrid.rowheight,fdolj.fGrid.Column3.Width+2,.T.
DO addtxtboxmy WITH 'fdolj',4,1,1,fdolj.fGrid.Column4.Width+2,.F.,.F.,0
fdolj.SetAll('Visible',.F.,'MyTxtBox')
=AFIELDS(arKat,'sprkat')
CREATE CURSOR cursprkat FROM ARRAY arKat
APPEND FROM sprkat
SELECT sprkat 
WITH fdolj.combobox3 
     .procForGotFocus='DO procGotDolKat'    
     .procForValid='DO procValidDolKat'
     .RowSourceType=6    
     .RowSource='cursprkat.name'    
     .ControlSource='fdolj.strname' 
     .Visible=.F.   
ENDWITH
DO addcontmy WITH 'fdolj','cont1',fdolj.fGrid.Left+13,fdolj.fGrid.Top+2,fdolj.fGrid.Column1.Width-3,;
   fdolj.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont1','sprdolj',1,4"
fdolj.cont1.SpecialEffect=1   
DO addcontmy WITH 'fdolj','cont2',fdolj.cont1.Left+fdolj.fGrid.Column1.Width+2,fdolj.fGrid.Top+2,;
  fdolj.fGrid.Column2.Width-4,fdolj.fGrid.HeaderHeight-3,'',"DO clickCont WITH 'fdolj','fdolj.cont2','sprdolj',2,5"
DO addcontmy WITH 'fdolj','cont3',fdolj.cont2.Left+fdolj.fGrid.Column2.Width+1,fdolj.fGrid.Top+2,fdolj.fGrid.Column3.Width-4,fdolj.fGrid.HeaderHeight-3,;
  '' 
DO addcontmy WITH 'fdolj','cont4',fdolj.cont3.Left+fdolj.fGrid.Column3.Width+1,fdolj.fGrid.Top+2,fdolj.fGrid.Column4.Width-4,fdolj.fGrid.HeaderHeight-3,;
  ''   
WITH fdolj.fGrid
     FOR ch=1 TO .ColumnCount-1
         obj_ch='.Column'+LTRIM(STR(ch))+'.Header1'          
         obj_col='fdolj.cont'+LTRIM(STR(ch))
         &obj_ch..FontSize=IIF(FONTMETRIC(6,dFontName,dFontSize)* TXTWIDTH(&obj_ch..Caption,dFontName,dFontSize)>&obj_col..Width,dFontSize-1,dFontSize)
     ENDFOR
ENDWITH 
SELECT sprdolj  
fdolj.Show

*************************************************************************************************************************
*
*************************************************************************************************************************
PROCEDURE readmydolj
PARAMETERS parlog
fdolj.strname=sprkat.name
SELECT sprdolj
IF parlog
   fdolj.fGrid.GridMyAppendBlank(1,'kod','name')
   fdolj.strname=''   
ENDIF
fdolj.nrec=RECNO()
SCATTER TO fdolj.dim_ap
fdolj.txtBox1.Left=fdolj.fGrid.Left+10
fdolj.txtBox2.Left=fdolj.txtbox1.Left+fdolj.txtbox1.Width-1
fdolj.comboBox3.Left=fdolj.txtbox2.Left+fdolj.txtbox2.Width-1
fdolj.txtBox4.Left=fdolj.comboBox3.Left+fdolj.comboBox3.Width-1
fdolj.combobox3.ControlSource='fdolj.strname'
fdolj.txtbox1.ControlSource='fdolj.dim_ap(1)'
fdolj.txtbox2.ControlSource='fdolj.dim_ap(2)'
fdolj.txtbox4.ControlSource='fdolj.dim_ap(4)'
lineTop=fdolj.fGrid.Top+fdolj.fGrid.HeaderHeight+fdolj.fGrid.RowHeight*(IIF(fdolj.fGrid.RelativeRow<=0,1,fdolj.fGrid.RelativeRow)-1)
fdolj.SetAll('Top',linetop,'MyTxtBox')
fdolj.SetAll('Height',fdolj.fGrid.RowHeight+1,'MyTxtBox')
fdolj.SetAll('BackStyle',1,'MyTxtBox')
fdolj.combobox3.Top=lineTop
fdolj.combobox3.Height=fdolj.fGrid.RowHeight+1
fdolj.combobox3.BackColor=fdolj.txtbox2.BackColor
fdolj.SetAll('Visible',.T.,'MyTxtBox')
fdolj.combobox3.Visible=.T.
fdolj.fGrid.Enabled=.F.
fdolj.txtbox2.SetFocus
IF parlog
   KEYBOARD '{TAB}'
ENDIF   
************************************************************************************************************************
PROCEDURE writemydolj
PARAMETERS par_log
fDolj.SetAll('Visible',.T.,'mymenucont')
fDolj.menuread.Visible=.F.
fDolj.menuexit.Visible=.F.
IF par_log
   SELECT sprdolj
   GATHER FROM fDolj.dim_ap
   REPLACE namework WITH ALLTRIM(name)+' '+ALLTRIM(prim)    
ENDIF    
fDolj.SetAll('Visible',.F.,'MyTxtBox')
fDolj.SetAll('Visible',.F.,'ComboMy')
fDolj.SetAll('Visible',.F.,'MySpinner')
fDolj.fGrid.Enabled=.T.
SELECT sprdolj
fDolj.fGrid.GridUpdate
GO fDolj.nrec
fDolj.fGrid.SetAll('Enabled',.F.,'Column')
fDolj.fGrid.Columns(fDolj.fGrid.ColumnCount).Enabled=.T.
GO fDolj.nrec
*************************************************************************************************************************
PROCEDURE procValidDolKat
SELECT sprdolj
fdolj.dim_ap(3)=cursprkat.kod
KEYBOARD '{TAB}'    
************************************************************************************************************************
PROCEDURE procGotDolKat
SELECT cursprkat
LOCATE FOR kod=sprkat->kod
nrec=RECNO()
GO TOP 
COUNT WHILE RECNO()#nrec TO varnrec
fdolj.combobox3.DisplayCount=MAX(fdolj.fGrid.RelativeRow,fdolj.fGrid.RowsGrid-fdolj.fGrid.RelativeRow)
fdolj.combobox3.DisplayCount=MIN(fdolj.combobox3.DisplayCount,RECCOUNT())
SELECT sprdolj
*************************************************************************************************************************
PROCEDURE delMyDolj
SELECT rasp
LOCATE FOR kd=sprdolj.kod
IF FOUND()
   fdolj.fGrid.GridNoDelRec   
ELSE 
  fdolj.fGrid.GridDelRec('fdolj.fGrid','sprdolj') 
ENDIF   


**************************************************************************************************************************
*                           Процедура перевода символьной строки в один формат
**************************************************************************************************************************
PROCEDURE unosimbol
PARAMETERS parsimb,parrep,parUp
parsimbnew=LOWER(&parsimb)
parsimbnew=CHRTRAN(parsimbnew,"qwertyuiop[]asdfghjkl;'zxcvbnm,.","йцукенгшщзхъфывапролджэячсмитьбюё")
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
*-----------------------Процедуры печати ведомостей--------
***********************************************************************************************************************
PROCEDURE adSetupPrnToForm
PARAMETERS parLeft,parTop,parWidth,parWord,parExcel,parFrm
page_beg=1
page_end=999
kvo_page=1 
logWord=.F.
nameFrm=IIF(!EMPTY(parFrm),parFrm,'fSupl')
*parLef - Лево
*parTop - Верх
*parWidth=Ширина
*parWord - направлять в Word
*parFrm - имя формы
STORE 0 TO dimCht
dimCht(1)=1
*WITH fSupl
WITH &nameFrm
     DO addshape WITH 'fSupl',91,parLeft,parTop,150,parWidth,8
     DO addComboPrn WITH 'fSupl',91,.Shape91.Top+10,.Shape91.Left+20,parWidth-40 
     DO adlabMy WITH 'fSupl',91,'Копий',.comboBox91.Top+dHeight+10,.comboBox91.Left,100,0,.T.
     DO addSpinnerMy WITH 'fSupl','spinner91',.lab91.Left+.lab91.Width+10,.lab91.Top,dHeight,FONTMETRIC(6,dFontName,dFontSize)* TXTWIDTH('W999W',dFontName,dFontSize),'kvo_page',1
    .lab91.Top=.lab91.Top+(.spinner91.Height-.lab91.Height+4)
     WITH .spinner91
         .SpecialEffect=1 
         .ForeColor=objForeColor
         .BackColor=objBackColor
         .SpinnerHighValue=99
         .SpinnerLowValue=1
         .KeyBoardHighValue=99
         .KeyBoardLowValue=1
         .Increment=1
         .Visible=.T.
    ENDWITH 
    DO adlabMy WITH 'fSupl',92,'Страницы с',.lab91.Top,.spinner91.Left+.spinner91.Width+15,100,0,.T.  
    tboxWidth=(.combobox91.Width-.spinner91.Width-.Lab91.Width-.lab92.Width-RetTxtWidth('по')-57)/2
    DO adtbox WITH 'fSupl',91,.lab92.Left+.lab92.Width+10,.spinner91.Top,RetTxtWidth('99999'),dHeight,'page_beg','Z',.T.,.F.
    DO adlabMy WITH 'fSupl',93,'по',.lab91.Top,.txtbox91.Left+.txtbox91.Width+10,100,0,.T.
    DO adtbox WITH 'fSupl',92,.lab93.Left+.lab93.Width+10,.spinner91.Top,.txtBox91.Width,dHeight,'page_end','Z',.T.,.F.
    .lab91.Left=.comboBox91.Left+(.comboBox91.Width-.lab91.Width-.spinner91.Width-.lab92.Width-.txtBox91.Width-.lab93.Width-.txtBox92.Width-55)/2
    .spinner91.Left=.lab91.Left+.lab91.Width+10
    .lab92.Left=.spinner91.Left+.spinner91.Width+15
    .txtBox91.Left=.lab92.Left+.Lab92.Width+10
    .lab93.Left=.txtBox91.Left+.txtBox91.Width+10
    .txtBox92.Left=.lab93.Left+.lab93.Width+10         
    
    DO addOptionButton WITH 'fSupl',91,'Всe',.txtbox91.Top+.txtbox91.height+10,.Shape91.Left+5,'dimCht(1)',0,'DO procDimCht WITH 1',.T.   
    DO addOptionButton WITH 'fSupl',92,'Чётные',.Option91.Top,.Option91.Left+.Option91.Width+5,'dimCht(2)',0,'DO procDimCht WITH 2',.T.   
    DO addOptionButton WITH 'fSupl',93,'Нечётные',.Option91.Top,.Option92.Left+.Option92.Width+10,'dimCht(3)',0,'DO procDimCht WITH 3',.T.   
        
    .Option91.Left=.Shape91.Left+(.Shape91.Width-.Option91.Width-.Option92.Width-.Option93.Width-20)/2
    .Option92.Left=.Option91.Left+.Option91.Width+10
    .Option93.Left=.Option92.Left+.Option92.Width+10          
    
      
    DO adCheckBox WITH 'fSupl','checkWord',IIF(parWord,'направить в MSWord','направить в Excel'),.Option91.Top+.Option91.Height+10,.Shape91.Left,150,dHeight,'logWord',0            
    .checkWord.Left=.Shape91.Left+(.Shape91.Width-.checkWord.Width)/2    
    IF parWord.OR.parExcel      
       .Shape91.Height=dHeight*2+.checkWord.Height+.Option91.Height+50  
    ELSE 
       .Shape91.Height=dHeight*2+.Option91.Height+40  
       .checkWord.Visible=.F.
    ENDIF 
ENDWITH   
***********************************************************************************************************************
PROCEDURE adButtonPrnToForm
PARAMETERS parProc1,parProc2,parProc3,parAply,parFrm
nameFrm=IIF(!EMPTY(parFrm),parFrm,'fSupl')
WITH &nameFrm
     DO addButtonOne WITH nameFrm,'butPrn',.Shape91.Left+(.Shape91.Width-RetTxtWidth('wпросмотрw')*3-20)/2,.Shape91.Top+.Shape91.Height+20,'печать','',parProc1,39,RetTxtWidth('wпросмотрw'),'печать' 
     DO addButtonOne WITH nameFrm,'butView',.butPrn.Left+.butPrn.Width+10,.butPrn.Top,'просмотр','',parProc2,39,.butPrn.Width,'просмотр' 
     DO addButtonOne WITH nameFrm,'butRet',.butView.Left+.butView.Width+10,.butPrn.Top,'возврат','',parproc3,39,.butPrn.Width,'возврат' 
     IF parAply
        DO addButtonOne WITH nameFrm,'cont11',.shape91.Left+(.Shape91.Width-(RetTxtWidth('wпринять')*2)-15)/2,.butPrn.Top,'принять','','DO returnToPrn',39,RetTxtWidth('wпринятьw'),'принять' 
        DO addButtonOne WITH nameFrm,'cont12',.cont11.Left+.cont11.Width+15,.butPrn.Top,'сброс','','DO returnToPrn',39,.cont11.Width,'сброс' 
        .cont11.Visible=.F.
        .cont12.Visible=.F.
     ENDIF    
ENDWITH
**************************************************************************************************************************
PROCEDURE returnToPrn
***************************************************************************************************************************
PROCEDURE procDimCht
PARAMETERS par1
STORE 0 TO dimCht
dimCht(par1)=1
fSupl.Refresh  
*******************************************************************************************************************************************************
PROCEDURE procOptionStructure
PARAMETERS par1
STORE 0 TO dim_prn
dim_prn(par1)=1
DO CASE
   CASE dim_prn(1)=1
        fSupl.ComboBox2.Enabled=.F.
   CASE dim_prn(2)=1 
        fSupl.ComboBox2.Enabled=.T.        
ENDCASE
fSupl.Refresh
*******************************************************************************************************************************************************
*                  Вспомогательная процедура для промотра отчета или печати с заданными параметрами
*******************************************************************************************************************************************************
PROCEDURE procForPrintAndPreview
PARAMETERS parreport,par_caption,parTerm,procExcel
IF !parTerm
   DO previewRep WITH parreport,par_caption  
ELSE  
   IF logWord.AND.!EMPTY(procExcel)
      DO &procExcel
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

*******************************************************************************************************************************************************
*                  Вспомогательная процедура для промотра отчета или печати с заданными параметрами
*******************************************************************************************************************************************************
PROCEDURE 1procForPrintAndPreview
PARAMETERS parreport,par_caption,parTerm
IF !parTerm
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
**********************************************************************************************************************************************************
DEFINE CLASS timerMy AS TIMER
       Visible=.T.                    
       ProcTimerEvent=''
       PROCEDURE timer Event                 
                 IF !EMPTY(This.procTimerEvent)
                    ProcDo=This.ProcTimerEvent
                    &ProcDo 
                 ENDIF              
ENDDEFINE

*-----------------------------------------------------------------------------------------------
*  	          Добавление колонок в Grid
*-----------------------------------------------------------------------------------------------
PROCEDURE addColumnToGrid
LPARAMETERS parGrid,parColumnCount
*parGrid - Grid
*parColumnCount - кол-во колонок

WITH &parGrid
     FOR i=1 TO parColumnCount
         .ADDOBJECT('Column'+LTRIM(STR(i)),'ColumnMy')
     ENDFOR
     FOR i=1 TO parColumnCount
         .Columns(i).DynamicBackColor='IIF(RECNO(This.RecordSource)#This.curRec,dBackColor,dynBackColor)'
         .Columns(i).DynamicForeColor='IIF(RECNO(This.RecordSource)#This.curRec,dForeColor,dynForeColor)'     

         .Columns(i).DynamicForeColor='IIF(RECNO(This.RecordSource)#This.curRec,dForeColor,dynForeColor)'
         .Columns(i).Text1.SelectedBackColor=selBackColor     
         .Columns(i).Header1.FontName=dFontName             
         .Columns(i).Header1.FontSize=dFontSize             
         .Columns(i).Header1.ForeColor=dForeColor
         .Columns(i).Header1.BackColor=headerBackColor         
         .Columns(i).Header1.Alignment=2          
     ENDFOR
     .rowsGrid=(.Height-.HeaderHeight)/.RowHeight
     .BackColor=dBackColor
ENDWITH
**********************************************************************************************
*                    Класс для создания командной кнопки
**********************************************************************************************
DEFINE CLASS myCommandButton AS CommandButton       
       Visible=.T.
       BackColor=dBackColor
       fontname=dFontName
       fontSize=dFontSize   
       Autosize=.F.     
       PicturePosition=12
       procForClick=''
       PROCEDURE Init
                * This.BackColor=This.Parent.BackColor   
       PROCEDURE Click
	             IF !EMPTY(This.ProcForClick)
                    ProcForDo=This.ProcForClick
                    &ProcForDo
                 ENDIF
    
ENDDEFINE


*****************************************************************************************************
*   Процедура добавления в форму кнопки-иконки
*****************************************************************************************************
PROCEDURE addButtonPicture
PARAMETERS parFrm,parName,parTop,parLeft,parWidth,parHeight,parPict,parProc,parToolTip


*PARAMETERS parFrm,parName,parCaption,parTop,parLeft,parWidth,parHeight,parProc,parVisible
* parFrm - имя формы
* parname - имя объекта
* parCaption  - Caption
* parTop  - Top
* parLeft - Left 
* parWidth - Width
* parHeight - Height
* parProc - Click
* parVisible - Visible
objButton=parname
&parFrm..AddObject(objButton,'myCommandButton')
WITH &parFrm..&objButton
     .Caption=''
     .Top=parTop
     .Left=parLeft
     .Width=parWidth
     .Height=parHeight
     .Picture=parPict     
     .procForClick=parProc
*     .Visible=parVisible
ENDWITH
******************************************************************************************************
*           Формирование окна для выбора процедуры
******************************************************************************************************
PROCEDURE procchoice
PARAMETERS par1,par2
IF !USED('datsupl')
   USE datsupl IN 0
ENDIF
SELECT * FROM datsupl WHERE datsupl.numproc=par1 INTO CURSOR curchoice READWRITE 
SELECT datsupl
USE
SELECT curchoice
INDEX ON num TAG T1
GO BOTTOM
maxChoice=num
GO TOP 
fSupl=CREATEOBJECT('Formsupl')
WITH fSupl          
     .Caption=IIF(!EMPTY(par2),par2,'')
     .BackColor=RGB(255,255,255) 
     DO addshape WITH 'fSupl',1,10,20,150,390,8               
     .procexit='DO exitFromProcChoice'
     DIMENSION dimproc(maxChoice)
     STORE 0 TO dimproc
     SELECT curchoice
     COUNT TO maxproc     
     dimproc(1)=1
     GO TOP
     proccx=ALLTRIM(proc)
     optiontop=.Shape1.Top+10 
     objWidth=0  
     FOR i=1 TO maxproc
          procForObj=ALLTRIM(procobj)
         &procForObj
       *  ON ERROR DO erSup        
         objForWidth='fSupl.option'+LTRIM(STR(num))
         objWidth=IIF(objWidth>=&objForWidth..Width,objWidth,&objForWidth..Width)
      *   ON ERROR 
         optiontop=optiontop+.Option1.Height+5
         SKIP
     ENDFOR    
     .Shape1.Width=objWidth+40
     .Shape1.Height=.Option1.Height*maxproc+(maxproc-1)*5+20
     *-----------------------------Кнопка приступить---------------------------------------------------------------------------
     DO addcontlabel WITH 'fSupl','cont1',fSupl.shape1.Left+(fSupl.Shape1.Width-(RetTxtWidth('Wприступить')*2)-30)/2,;
        fsupl.Shape1.Top+fSupl.Shape1.Height+20,RetTxtWidth('Wприступить'),dHeight+5,'приступить','DO procRunChoice'

    *---------------------------------Кнопка отмена --------------------------------------------------------------------------
    DO addcontlabel WITH 'fSupl','cont2',fSupl.cont1.Left+fSupl.cont1.Width+15,fSupl.Cont1.Top,;
       fSupl.Cont1.Width,dHeight+5,'отмена','DO exitFromProcChoice','Отмена'
    .SetAll('ForeColor',RGB(0,0,128),'CheckBox')  

    *-------------параметры формы----------------------------------------------------------------------------------------    
     
     .Width=.Shape1.Width+20
     .Height=.Shape1.Height+.cont1.Height+60         
   
ENDWITH
DO pasteImage WITH 'fSupl'
fSupl.Show
*******************************************************************************************************************************************************************
PROCEDURE procSelectOptionChoice
PARAMETERS par1
STORE 0 TO dimproc
dimproc(par1)=1
fSupl.Refresh
SELECT curchoice
LOCATE FOR num=par1
proccx=ALLTRIM(proc)
*********************************************************************************************************************************************************
PROCEDURE procRunChoice
fSupl.Visible=.F.
fSupl.Release 
DO &proccx
********************************************************************************************************************************************************
PROCEDURE exitFromProcChoice
fSupl.Release
***************************************************************************************************************************************************
PROCEDURE procValOption
PARAMETERS parFrm,parDim,parNum
STORE 0 TO &parDim
&parDim(parNum)=1
&parFrm..Refresh
****************************************************************************************************************************************************
PROCEDURE erSup
****************************************************************************************************************************************************
PROCEDURE dateToString
PARAMETERS parDate,parStrYear
repVar=LTRIM(STR(DAY(&parDate)))+' '+month_prn(MONTH(&parDate))+' '+STR(YEAR(&parDate),4)+IIF(parStrYear,' г.','')
RETURN repVar
****************************************************************************************************************************************************
PROCEDURE addShapePercent
PARAMETERS parFrm,parLeft,parTop,parHeight,parWidth
WITH &parFrm
     DO addShape WITH parFrm,11,parLeft,parTop,parHeight,parWidth,8
     .Shape11.BackStyle=0
     .Shape11.Visible=.F.
     DO addShape WITH parFrm,12,.Shape11.Left,.Shape11.Top,.Shape11.Height,0,8
     .Shape12.BackStyle=1
     .Shape12.BackColor=selBackColor
     .Shape12.Visible=.F.       
     DO adLabMy WITH parFrm,25,'100%',.Shape11.Top+3,.Shape11.Left,.Shape11.Width,2,.F.,0
     .lab25.Top=.Shape11.Top+(.Shape11.Height-.Lab25.Height)/2
     .lab25.Visible=.F.  
ENDWITH 
***************************************************************************
PROCEDURE startPrnToExcel
PARAMETERS parFrm
WITH &parFrm
     .SetAll('Visible',.F.,'myCommandButton')    
     .SetAll('Visible',.F.,'myContLabel')    
     .Shape11.Visible=.T.
     .Shape12.Visible=.T.
     .Shape12.Width=1 
     .lab25.Visible=.T.
     .lab25.Caption='0%' 
ENDWITH 
PUBLIC xlCenter,xlRight,xlLeft,xlTop,xlThin,xlMedium,xlDiagonalDown,xlDiagonalUp,xlEdgeLeft,xlEdgeTop,xlEdgeBottom,xlEdgeRight,xlInsideVertical,xlInsideHorizontal  
xlCenter= -4108            
xlLeft= -4131 
xlRight= -4152
xlTop=-4160             
xlThin= 2                  
xlMedium= -4138            
xlDiagonalDown= 5          
xlDiagonalUp= 6                 
xlEdgeLeft= 7              
xlEdgeTop= 8               
xlEdgeBottom= 9            
xlEdgeRight= 10            
xlInsideVertical= 11         
xlInsideHorizontal= 12  
***************************************************************************
PROCEDURE endPrnToExcel
PARAMETERS parFrm
RELEASE xlCenter,xlLeft,xlRight,xlTop,xlThin,xlMedium,xlDiagonalDown,xlDiagonalUp,xlEdgeLeft,xlEdgeTop,xlEdgeBottom,xlEdgeRight,xlInsideVertical,xlInsideHorizontal  
WITH &parFrm
      .butPrn.Visible=.T.
      .butView.Visible=.T.
      .butRet.Visible=.T.      
      .Shape11.Visible=.F.
      .Shape12.Visible=.F.      
      .lab25.Visible=.F.      
ENDWITH      
****************************************************************************
PROCEDURE storezeropercent
STORE 0 TO max_rec,one_pers,pers_ch
COUNT TO max_rec
****************************************************************************
PROCEDURE fillpercent
PARAMETERS parFrm
one_pers=one_pers+1
pers_ch=one_pers/max_rec*100
&parFrm..lab25.Caption=LTRIM(STR(pers_ch))+'%'       
&parFrm..Shape12.Width=&parFrm..shape11.Width/100*pers_ch  