VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7500
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   8415
   _ExtentX        =   14843
   _ExtentY        =   13229
   _Version        =   393216
   Description     =   "Flying Windows VB6 V. 2.4.1 is an  VB addin to have more room for working within the VB IDE."
   DisplayName     =   "Flying Windows VB6"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'
'   Flying Windows VB6 - AddIn Connect
'

'
'   Flying Windows V. 2.5.0
'
'
'   Last edit       :   10/20/2004
'
'   Started         :   January 2003, VB5 version, Light_Templer, Germany (schwepps_bitterlemon@gmx.de)
'   Conversion      :   To VB6 in September/October 2004 with adding of three new features.
'   Copyrights      :   All copyrights by Light Templer. Don't sell this as a compiled DLL!
'   Risk            :   Use at your own risk. I'm not responsible for anything ;)
'
'
'   A freeware addin for MS Visual Basic 6
'
'   Many thanks to Carlos J. Quintero (www.mztools.com) for his kick into my ass to do it by myself!
'   And much more thanks for his great freeware addin  'MZ-Tools' :
'   The BEST thing a developer could happen! Visit his site - you will be surprised!
'
'
'   ******************************
'   *  WHAT windows are flying?  *
'   ******************************
'   Implemented in Flying Windows so far:
'
'                           * Hotcorners for most used tool windows (toolbox, properties, project
'                             explorer and immediate window)
'
'                           * Empty the immediate window when ctrl key is pressed on opening by hotcorner.
'
'                           * Tooltip window for controls showing the name and the most important properties.
'
'                           * Click on a property value in the tooltip window with ctrl key pressed puts this
'                             value into the clipboard.
'
'                           * Tooltip windows are moveable on caption bar. Leave them open with ctrl key pressed when
'                             mouse leaves tooltip window. Close them with a simple mouseclick.
'
'                           * Show mouse pointers absolut screen coordinates in VB IDEs title bar in pixel.
'
'                           * With an open source code pane moving the mouse to the left border of the screen opens
'                             the coresponding designer window (same as pressing  <Shift-F7> ).
'
'                           * The size of comboboxes of a code pane are increased to show much more values without
'                             boring scrolling. The same effect is on File Open/Save dialogs combo boxes.
'
'
'   --- NEW WITH UPDATE 1 ---
'
'                           * Moving the mouse to the right border of the screen increases the topmost code window
'                             to full VB IDE client area size (but doesn't maxmize it!).
'
'                           * Moving the mouse to the right border of the screen AND hold the <Ctrl> key pressed
'                             closes the topmost window which has VB IDE as parent window. This can be a code,
'                             a designer or any other window (not dialog!): e.g. Object Browser, Watch Window, ...
'
'
'   --- NEW WITH UPDATE 2 ---
'
'                           * With an open designer code window moving the mouse to the left border of the screen opens
'                             the coresponding code window (same as pressing <F7> ).
'
'
'   --- NEW WITH UPDATE 3 (V. 2.5.0) ---
'
'                           * Now your CODE is flying, too ;) ! I have added a powerfull AutoComplete feature to Flying
'                             Windows:  Write 2 or 3 letters, press F-12 (function key 12) and this small keyword will
'                             replaced by a longer word or a couple of lines with code - whatever you want. A long list
'                             with abbreviations missed in native VB is included. No more writing 'End With', 'Select Case'
'                             'Private WithEvents' or a standard header block. Try 'ew' and press F-12, try 'sc' and press
'                             F-12 or as an example for the full power type '*!' and press F-12 ...
'                             The replacement is done by VBs 'SendKey' command so you can use all of its possiblities.
'
'
'           _________________________________________________________________________
'
'                 This is my first VB AddIn and my first conversion to VB6.
'             (But not my first VB proggy ;) ). Plz be kind to it.
'   ___________________________________________________________________________



'   UPDATE 1 - CHANGES/FIXES
'
'   Thx for all comments from comunity on PSC - here is what i 've changed to get better:
'
'   1 - Mouse pointer will moved from hotcorner to over the window which just appears as before. That cannot
'       be changed by design of this function. But now the mouse pointer is much closer to the hotcorner so
'       you don't have to align the mouse. The feeling is much better this way. Thx, Alaeddin Hallak.
'       For other (own) solutions: All of this calculation is done in 'SetMousePosOverWindow()'
'
'   2 - Tests for docking mode of hotcorner windows are added. The Overflow Error (Raised, when setting a docked
'       windows's Left position property) is catched and gives a long msgbox note what to do to avoid this.
'       Sorry to all for this problem in first release. I'm not using docked windows and so i didn't get this
'       problem on my system earlier.
'
'   3 - Added a configurable value to FlyingWindows options dialog to set a max value for increased combo boxes.
'
'   4 - Changed all msgboxes (error warnings) to msgbox "..." , vbExclamation + vbMsgBoxSetForeground, ... to
'       get the error in forground. Thx, Tom Pydeski.
'
'   5 - Added features for 'mouse at RIGHT screen side' event. Please look 40 lines above ;)


'   UPDATE 2 - CHANGES/FIXES
'
'   1 - Flying Windows Option Dialog is now shown in 'modal' mode.
'
'   2 - Switch in option dialog for increasing the combo boxes doesn't work. Fixed.
'
'   3 - A switch is added to option dialog to get the left screen border /right screen border functions off and on.
'
'   4 - Added additional function for 'mouse at LEFT screen side' event. Please look 43 lines above ;)
'

'   UPDATE 3 - CHANGES/FIXES
'
'   No changes - just the new main feature 'AutoText' - plz read above for details.
'



Option Explicit

' *************************************
' *  CONSTS TO CHANGE To YOUR NEEDS   *
' *************************************
Private Const AT_DATEFORMAT As String = "mm/dd/yyyy"        ' Used in 'AutoText' part
Private Const AT_TIMEFORMAT As String = "hh:nn"             ' Used in 'AutoText' part
Private Const AT_AUTHOR     As String = "'AUTHOR_NAME'"     ' Used in 'AutoText' part


' *************************************
' *           CONSTS                  *
' *************************************
Private Const COMBO_SIZE_DEFAULT As Long = 400&             ' Default vertical size of IDE combo boxes


' *************************************
' *           PUBLICS                 *
' *************************************
Public flgActivateHotCorners        As Boolean
Public flgActivateToolTips          As Boolean
Public flgActivatePosition          As Boolean
Public flgActivateLargeCombos       As Boolean
Public flgActivateHotSides          As Boolean
Public flgActivateAutoText          As Boolean
Public lMaxComboSize                As Long


' *************************************
' *        PRIVATE TYPES              *
' *************************************
Private Type POINTAPI
    pX                  As Long
    pY                  As Long
End Type

Private Type RECTAPI
    lLeft               As Long
    lTop                As Long
    lRight              As Long
    lBottom             As Long
End Type

Private Type tpAutoText
    sKey                As String               ' Keyword which is to be substituted
    sSubstCode          As String               ' New text to put into source code
End Type


' *************************************
' *            PRIVATES               *
' *************************************
Private WithEvents MenuHandler      As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents VBModeEvents     As VBBuildEvents
Attribute VBModeEvents.VB_VarHelpID = -1
Private WithEvents m_frmTimer       As frmTimer
Attribute m_frmTimer.VB_VarHelpID = -1

Private VBInstance                  As VBIDE.VBE
Private mcbMenuCommandBar           As Office.CommandBarControl
Private frmFFoptions                As frmFFoptions

Private m_objPropertiesWindow       As Window
Private m_objToolWindow             As Window
Private m_objProjectExplorerWindow  As Window
Private m_objImmediateWindow        As Window

' For Hotcorners
Private lDelayShowHotCorners        As Long         ' Display delay time in milli seconds (200)
Private lDelayRemoveHotCorners      As Long         ' Delay time in milli seconds before removing (500)

' For AutoText
Private lDelayAutoText              As Long         ' Check interval

' For ToolTips
Private lDelayShowTooltips          As Long         ' Display delay time in milli seconds (500)
Private lDelayRemoveTooltips        As Long         ' Delay time in milli seconds before removing (200)
Private frmTTips                    As frmToolTips  ' Those nice litte windows

' Calc screen dimensions in pixel
Private Screen_X                    As Long
Private Screen_Y                    As Long

Private oMainWin                    As Window
Private lwHndlVBIDEmain             As Long
Private lwHndlVBIDEMDIClient        As Long
Private IDEmode                     As vbext_VBAMode
Private arrAutoText()               As tpAutoText   ' Holds Key/SubstCode pairs from FW_AutoText.Txt



' *******************************
' *          API STUFF          *
' *******************************
Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long

Private Declare Function SetCursorPos Lib "user32" _
        (ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function GetWindowRect Lib "user32" _
        (ByVal hWnd As Long, _
         lpRect As RECTAPI) As Long

Private Declare Function OffsetRect Lib "user32" _
        (lpRect As RECTAPI, _
         ByVal lX As Long, _
         ByVal lY As Long) As Long

Private Declare Function ClientToScreen Lib "user32" _
        (ByVal hWnd As Long, _
         lpPoint As POINTAPI) As Long

Private Declare Function PtInRect Lib "user32" _
        (lpRect As RECTAPI, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Long

Private Const VK_CONTROL = &H11

Private Declare Function IsWindow Lib "user32" _
        (ByVal hWnd As Long) As Long

Private Declare Function API_GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hWnd As Long, _
         ByVal lpClassName As String, _
         ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
        (ByVal hWnd As Long, _
         ByVal lpString As String, _
         ByVal cch As Long) As Long

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
        (ByVal hWnd As Long, _
         ByVal lpString As String) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

' Show windows without activation
Private Declare Function ShowWindow Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal nCmdShow As Long) As Long

Private Const SW_SHOWNA = 8


Private Declare Function WindowFromPoint Lib "user32" _
        (ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function GetParent Lib "user32" _
        (ByVal hWnd As Long) As Long

Private Declare Function API_FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hWndParent As Long, _
         ByVal hWndFirstChild As Long, _
         ByVal lpClass As String, _
         ByVal lpCaption As String) As Long

Private Declare Function API_MoveWindow Lib "user32" Alias "MoveWindow" _
        (ByVal hWnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Private Declare Function API_SetWindowPos Lib "user32" Alias "SetWindowPos" _
        (ByVal hWnd As Long, _
         ByVal hWndInsertAfter As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal cx As Long, _
         ByVal cy As Long, _
         ByVal wFlags As Long) As Long

Private Declare Function API_FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long

Private Declare Function API_IsWindowVisible Lib "user32" Alias "IsWindowVisible" _
        (ByVal hWnd As Long) As Long

Private Declare Function API_LockWindowUpdate Lib "user32" Alias "LockWindowUpdate" _
        (ByVal hWndLock As Long) As Long


' *************************************
' *        PRIVATE ENUMS              *
' *************************************
Private Enum enFW_HC_State
    [State Idle]
    [State Win Lft Top Opened]
    [State Win Rgt Top Opened]
    [State Win Lft Btm Opened]
    [State Win Rgt Btm Opened]
    [State Dlg Opened]
End Enum

Private Enum enWindowPosition
    [WinPos LeftTop]
    [WinPos LeftBottom]
    [WinPos RightTop]
    [WinPos RightBottom]
End Enum

Private Enum enFW_TT_State
    [State TT Idle]
    [State TT Win Open]
    [State TT Win Keep Open]
End Enum
'
'
'


' Here we have this three parts:
'
'   1   PUBLIC  SUBS/FUNCS
'   2   PRIVATE SUBS/FUNCS
'   3   USUAL ADDIN HANDLING STUFF (all private)





' ***************************************************************
' *                                                             *
' *                 PUBLIC SUBS/FUNCS                           *
' *                                                             *
' ***************************************************************

Public Sub HideFrmOptions()
    
    On Error Resume Next
    
    frmFFoptions.Hide
   
End Sub

Public Sub ShowFrmOptions()
  
    On Error Resume Next
    
    frmFFoptions.Show vbModal
   
End Sub





' ***************************************************************
' *                                                             *
' *                   PRIVATE SUBS/FUNCS                        *
' *                                                             *
' ***************************************************************

Private Sub Init_Startup()
    ' Prepare to run

    Dim nEvents2    As Events2
    
    
    On Error GoTo error_handler
    
    
    ' Init hidden event handler
    Set nEvents2 = VBInstance.Events
    Set VBModeEvents = nEvents2.VBBuildEvents
    If VBModeEvents Is Nothing Then
        
        Exit Sub
    End If
    
    ' Set up start value when running VB
    IDEmode = vbext_vm_Design
    
    ' Get ref to VB IDEs main window
    Set oMainWin = VBInstance.MainWindow
    If oMainWin Is Nothing Then
        MsgBox "Couldn't get reference to object 'main window'!", vbExclamation + vbMsgBoxSetForeground, " Start aborted!"
        
        Exit Sub
    End If
    
    ' Get window handle of main window (thats easy, because correct implemented by MS)
    lwHndlVBIDEmain = oMainWin.hWnd
    If lwHndlVBIDEmain = 0 Then
        MsgBox "Couldn't get window handle of IDE's main window!", vbExclamation + vbMsgBoxSetForeground, " Start aborted!"

        Exit Sub
    End If
     
    ' Get VB IDE's "Desktop" window handle. Its a child window with class MDIClient.
    ' Needing/looking for more window handles than the main app window is ugly, because of
    ' MS has implemented the (hidden) property, but doesn't put any value into it ... ;(((
    lwHndlVBIDEMDIClient = API_FindWindowEx(lwHndlVBIDEmain, ByVal 0&, "MDIClient", vbNullString)
    
    
    ' Read options from registry
    flgActivateHotCorners = IIf(GetSetting(App.EXEName, "Settings", "HotCorners", "1") = "1", True, False)
    flgActivateToolTips = IIf(GetSetting(App.EXEName, "Settings", "ToolTips", "1") = "1", True, False)
    flgActivatePosition = IIf(GetSetting(App.EXEName, "Settings", "Position", "1") = "1", True, False)
    flgActivateLargeCombos = IIf(GetSetting(App.EXEName, "Settings", "LargeCombos", "1") = "1", True, False)
    flgActivateHotSides = IIf(GetSetting(App.EXEName, "Settings", "HotSides", "1") = "1", True, False)
    flgActivateAutoText = IIf(GetSetting(App.EXEName, "Settings", "AutoText", "1") = "1", True, False)
    lMaxComboSize = Val(GetSetting(App.EXEName, "Settings", "MaxComboSize", Format(COMBO_SIZE_DEFAULT)))
    
    ' Our timers to do something
    If m_frmTimer Is Nothing Then

        ' Prepare the timers for Flying Windows
        Set m_frmTimer = New frmTimer
    End If
    
    
    ' Our option window
    If frmFFoptions Is Nothing Then
        Set frmFFoptions = New frmFFoptions
        Set frmFFoptions.Connect = Me
    End If
    
    
    GetRefsToToolWindows
    
    ' Check the four windows (toolbox/property/immediate/project):  Floating (ok) or docked ?
    If flgActivateHotCorners = True And CheckWinMode() = False Then
        MsgBox "One or all of the four windows" + vbCrLf + _
                "(toolbox/property/immediate/project)" + vbCrLf + _
                "are not set to 'floating' (docking off)." + vbCrLf + _
                "Due to this the hotcorner function is deactivated." + vbCrLf + _
                "After changing this VB IDEs option dialog / docking tab" + vbCrLf + _
                "it can be switched on in Flying Windows config dialog." + vbCrLf + _
                "Thank you.", _
                vbExclamation + vbMsgBoxSetForeground, " AddIn Flying Windows:  Problem on StartUp"
        
        flgActivateHotCorners = False
    End If

    ' Put into addin menu
    Set mcbMenuCommandBar = AddToAddInCommandBar("Flying Windows")
    Set MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)

    ' Calc constants
    Screen_X = (Screen.Width / Screen.TwipsPerPixelX) - 1       ' Done here for speedup only,
    Screen_Y = (Screen.Height / Screen.TwipsPerPixelY) - 1      ' but should be in 'm_frmTimer_TimerElapsed()'
    
    
    ' Presets (Those values do a good job for me. But change to your needs. Values im milliseconds)
    lDelayShowHotCorners = 200
    lDelayRemoveHotCorners = 300
    lDelayAutoText = 60
    lDelayShowTooltips = 200
    lDelayRemoveTooltips = 200

    ' Activate!
    m_frmTimer.SetTimerHotcorners lDelayShowHotCorners
    m_frmTimer.SetTimerToolTips lDelayShowTooltips
    
    If flgActivateAutoText = True Then
        LoadAutoText
        m_frmTimer.SetTimerAutoText lDelayAutoText
    End If

    DoEvents

    Exit Sub


error_handler:

    MsgBox Err.Description, vbExclamation + vbMsgBoxSetForeground, " FlyingWindows:   Error in 'Init_Startup()'"

End Sub


Private Sub VBModeEvents_EnterDesignMode()
    ' Silly workarround for unfinished VBE to get the app state.
    ' This way we miss 'Break' mode, but it works with all localized VB versions (english, german, ...)
    
    ' MS VB6 MSDN AddIn Docu says: There is a property 'Mode'  (Debug.Print Application.VBE.ActiveVBProject.Mode)
    ' Maybe I'm not good enough to find it ... ;(
    
    IDEmode = vbext_vm_Design
    
End Sub

Private Sub VBModeEvents_EnterRunMode()
    ' Silly workarround for unfinished VBE to get app state. Comments look above, plz.

    IDEmode = vbext_vm_Run

End Sub

Private Sub GetRefsToToolWindows()
    ' Get references to the four VB tool windows called using the hot corners
    
    Dim oWindow     As Window
    
    On Local Error Resume Next
    
    For Each oWindow In VBInstance.Windows
            
        If m_objImmediateWindow Is Nothing And oWindow.Type = vbext_wt_Immediate Then
            Set m_objImmediateWindow = oWindow
        End If
        
        If m_objProjectExplorerWindow Is Nothing And oWindow.Type = vbext_wt_ProjectWindow Then
            Set m_objProjectExplorerWindow = oWindow
        End If
        
        If m_objToolWindow Is Nothing And oWindow.Type = vbext_wt_ToolWindow And oWindow.Caption = "" Then
            Set m_objToolWindow = oWindow
        End If
        
        If m_objPropertiesWindow Is Nothing And oWindow.Type = vbext_wt_PropertyWindow Then
            Set m_objPropertiesWindow = oWindow
        End If
            
    Next oWindow
    
End Sub

Private Function CheckWinMode() As Boolean
    ' Check window mode (docked/floating)
    ' Result: TRUE = success if all checked windows are floating.
    
    ' Here we have a good example for the consistence in Microsofts VB development process:
    ' Compare the windows classes of this four very similar windows in floating and in
    ' docked mode for a good laugh ;)
    
    
    Dim lwHndl      As Long
    Dim flgOldState As Boolean
    
    
   On Error GoTo CheckWinMode_Error
   

    ' The windows must be in 'visible' state to get them. With this call we prevent
    ' that the user sees a flickering when open/close a window.
    API_LockWindowUpdate lwHndlVBIDEMDIClient
    
    
    ' Check TOOLBOX window
    ' The window class is "DockingView" in floating mode and "ToolsPalette" when docked.
    With m_objImmediateWindow
        flgOldState = .Visible
        .Visible = True
        lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", vbNullString)  ' No caption, so this check isn't sure...
        m_objImmediateWindow.Visible = flgOldState
    End With
    If lwHndl = 0 Then
        API_LockWindowUpdate False                      ' Enable window refreshing/repainting
        
        Exit Function
    End If
    
    
    ' Check IMMEDIATE window
    ' The window class is "DockingView" in floating mode and "VbaWindow" when docked.
    With m_objImmediateWindow
        flgOldState = .Visible
        .Visible = True
        lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", m_objImmediateWindow.Caption)
        m_objImmediateWindow.Visible = flgOldState
    End With
    If lwHndl = 0 Then
        API_LockWindowUpdate False                      ' Enable window refreshing/repainting
        
        Exit Function
    End If
    
    
    ' Check PROPERTIES window
    ' The window class is "DockingView" in floating mode and "wndclass_pbrs" when docked.
    With m_objPropertiesWindow
        flgOldState = .Visible
        .Visible = True
        lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", m_objPropertiesWindow.Caption)
        m_objPropertiesWindow.Visible = flgOldState
    End With
    If lwHndl = 0 Then
        API_LockWindowUpdate False                      ' Enable window refreshing/repainting
        
        Exit Function
    End If
    
    
    ' Check PROJECT EXPLORER window
    ' The window class is "DockingView" in floating mode and "PROJECT" when docked.
    With m_objProjectExplorerWindow
        flgOldState = .Visible
        .Visible = True
        lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", m_objProjectExplorerWindow.Caption)
        m_objProjectExplorerWindow.Visible = flgOldState
    End With
    If lwHndl = 0 Then
        API_LockWindowUpdate False                      ' Enable window refreshing/repainting
        
        Exit Function
    End If
    
    API_LockWindowUpdate False                      ' Enable window refreshing/repainting
    
    ' All checks are ok
    CheckWinMode = True

    Exit Function


CheckWinMode_Error:
    
    API_LockWindowUpdate False
    
End Function


' **************************************************
' *                 HOTCORNER PART                 *
' **************************************************
Private Sub m_frmTimer_TimerHotcornersElapsed()
    ' All the stuff needed for hotcorner handling
                                                                ' set left/top corner of window plus this offset
    Static s_FW_State       As enFW_HC_State
    Static s_SavedState     As enFW_HC_State
    Static s_LastWinObj     As Window
    Static s_lwHndl         As Long
    Static s_flgNoRepeat    As Boolean                          ' Used to prevent repeated commands when mouse isn't move
                                                                ' away from right border
                                                                
    Dim MousePos            As POINTAPI
    Dim WinRECT             As RECTAPI
    Dim RectMDIClient       As RECTAPI
    Dim MainWin             As Window
    Dim lHndlTopWindow      As Long
    Dim oActiveCodeWindow   As Window
    Dim oActiveWindow       As Window
    
        
    On Local Error Resume Next
    
    
    If flgActivateHotCorners = False Then                       ' Not wanted by option setting? Leave sub...
        
        Exit Sub
    End If
    
    
    Set MainWin = VBInstance.MainWindow
    If MainWin Is Nothing Then                                  ' Just for sure ... (got trouble in special situations!)
        
        Exit Sub
    End If

    On Local Error GoTo error_handler
    
    With VBInstance

        ' If IDE in 'RUN' mode: Do nothing!
        If IDEmode <> vbext_vm_Design Then

            Exit Sub
        End If

        ' Minimized IDE? Do nothing!
        If MainWin.WindowState = vbext_ws_Minimize Then

            Exit Sub
        End If

        ' IDE in foreground ?
        lHndlTopWindow = GetForegroundWindow()
        Do While lHndlTopWindow <> lwHndlVBIDEmain
            lHndlTopWindow = GetParent(lHndlTopWindow)

            ' Top most window reached ?
            If lHndlTopWindow = 0 Then

                ' Yes!  > Leave this!
                Exit Sub
            End If
        Loop


        ' Prevent closing if user opens a dialog from your poped up window
        If s_FW_State <> [State Idle] And s_FW_State <> [State Dlg Opened] Then

            If .ActiveWindow Is Nothing Then
                ' Easy for dlgs like StdColorSelector
                s_SavedState = s_FW_State
                s_FW_State = [State Dlg Opened]

                Exit Sub
            End If

            If .ActiveWindow.Caption = MainWin.Caption Then
                ' More tricky for the others (StdFont, StdFileOpen, ...). Don't know the difference :(
                s_SavedState = s_FW_State
                s_FW_State = [State Dlg Opened]

                Exit Sub
            End If

        End If
    End With

    GetCursorPos MousePos
    With MousePos

        Select Case s_FW_State

            Case [State Idle]       ' === Wait for mouse in a screen corner

                    ' [Code could be shorter, but this long form should be the fastest when
                    '  compiled (hope so ;) ) and easier to read and modfiy.

                    ' === LEFT =========================================
                    If .pX = 0 Then

                        ' === LEFT / TOP :  Toolbox ===
                        If .pY = 0 Then
                            If m_objToolWindow Is Nothing Then              ' Sometimes call in startup goes wrong ... ;(
                                GetRefsToToolWindows                        ' Don't know why, it's hard to track!
                            End If                                          ' Then we try here ones more.
                            If m_objToolWindow Is Nothing Then              ' Just for sure ... ;)
                                s_FW_State = [State Idle]
                                
                                Exit Sub
                            End If
                            
                            ' Show window at position and set focus to it
                            With m_objToolWindow
                                .Visible = True
                                .Left = 0                                   ' Setting works with 'Toolbox' as usual.
                                .Top = 0
                                .SetFocus
                                DoEvents
                                
                                ' We need the window handle of this tool window
                                s_lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", vbNullString)
                                
'                                GetWindowRect lwHndlVBIDEMDIClient, RectMDIClient
'                                s_lwHndl = WindowFromPoint(RectMDIClient.lLeft + 10, RectMDIClient.lTop + 10)
                                SetMousePosOverWindow s_lwHndl, [WinPos LeftTop]
                                
                            End With

                            ' Longer check interval to prevent removing the window, if
                            ' you just move the mouse pointer out of the window for a very
                            ' short time
                            Call m_frmTimer.SetTimerHotcorners(lDelayRemoveHotCorners)

                            ' Store reference to last showed window
                            Set s_LastWinObj = m_objToolWindow

                            ' Set state to "Wait for window leaved by mouse pointer"
                            s_FW_State = [State Win Lft Top Opened]

                            Exit Sub
                        End If


                        ' === LEFT / Bottom :  Immediate ===
                         If .pY = Screen_Y Then
                            
                            If m_objImmediateWindow Is Nothing Then         ' Sometimes call in startup goes wrong ... ;(
                                GetRefsToToolWindows                        ' Don't know why, it's hard to track!
                            End If                                          ' Then we try here ones more.
                            If m_objImmediateWindow Is Nothing Then         ' Just for sure ... ;)
                                s_FW_State = [State Idle]
                                
                                Exit Sub
                            End If
                                                     
                            With m_objImmediateWindow
                                .Visible = True
                                s_lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", .Caption)
                                
                                GetWindowRect lwHndlVBIDEMDIClient, RectMDIClient
                                GetWindowRect s_lwHndl, WinRECT
                                
                                ' Don't know why, but simple setting of position by properties doesn't work here ;(
                                API_MoveWindow s_lwHndl, _
                                                0, _
                                                ((RectMDIClient.lBottom - RectMDIClient.lTop) - (WinRECT.lBottom - WinRECT.lTop)), _
                                                (WinRECT.lRight - WinRECT.lLeft), _
                                                (WinRECT.lBottom - WinRECT.lTop), _
                                                1
                                .SetFocus

                                ' If Ctrl key pressed:  Clear immediate window
                                If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) Then
                                    DoEvents
                                    SendKeys "^a", True         ' Select all
                                    DoEvents
                                    SendKeys "{BS}"             ' Erase selected
                                    DoEvents
                                End If
                                
                                SetMousePosOverWindow s_lwHndl, [WinPos LeftBottom]
                            End With

                            Call m_frmTimer.SetTimerHotcorners(lDelayRemoveHotCorners)      ' Comments: Same as above ;)
                            Set s_LastWinObj = m_objImmediateWindow
                            s_FW_State = [State Win Lft Btm Opened]

                            Exit Sub
                        End If

                        
                        ' **************************************************************
                        ' * Gimmick No 4:   LEFT Screen Border
                        ' **************************************************************
                        ' We are at the LEFT border of the screen, but not in a corner
                        If flgActivateHotSides = True Then
                            If .pY > 30 And .pY < Screen_Y - 30 Then            ' A little gap to top and bottom?
                                On Local Error Resume Next
                                
                                If s_flgNoRepeat = False Then               ' Prevent multiple raises
                                
                                    ' Is topmost window a designer window?
                                    Set oActiveWindow = VBInstance.ActiveWindow     ' *** Note: THIS ISN'T TRACEABLE IN VB IDE!
                                    If Not oActiveWindow Is Nothing Then            ' ***       YOU DON'T GET THE REFERENCE!
                                        If oActiveWindow.Type = vbext_wt_Designer Then
                                            s_flgNoRepeat = True
                                            Set oActiveWindow = Nothing
                                            SendKeys "{F7}"                         ' ===> Call coresponding code window
                                        
                                            Exit Sub
                                        End If
                                    End If
                                    Set oActiveWindow = Nothing
                                    
                                    ' Is topmost window a code window?
                                    Set oActiveCodeWindow = VBInstance.ActiveCodePane.Window
                                    If Not oActiveCodeWindow Is Nothing Then
                                        s_flgNoRepeat = True
                                        Set oActiveCodeWindow = Nothing
                                        SendKeys "+{F7}"                            ' ===> Call coresponding designer window
                                        
                                        Exit Sub
                                    End If
                                    Set oActiveCodeWindow = Nothing
                                    
                                End If
                                
                                Exit Sub
                            End If
                        End If
                        
                    End If


                    ' === Right ======================================
                    If .pX = Screen_X Then

                        ' === Right / Top :  Properies ===
                        If .pY = 0 Then
                            If m_objPropertiesWindow Is Nothing Then        ' Sometimes call in startup goes wrong ... ;(
                                GetRefsToToolWindows                        ' Don't know why, it's hard to track!
                            End If                                          ' Then we try here ones more.
                            If m_objPropertiesWindow Is Nothing Then        ' Just for sure ... ;)
                                s_FW_State = [State Idle]
                                
                                Exit Sub
                            End If
                            
                            With m_objPropertiesWindow
                                .Visible = True
                                s_lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", .Caption)
                                
                                GetWindowRect lwHndlVBIDEMDIClient, RectMDIClient
                                GetWindowRect s_lwHndl, WinRECT
                                
                                ' Don't know why, but simple setting of position by properties doesn't work here ;(
                                API_MoveWindow s_lwHndl, _
                                                RectMDIClient.lRight - (WinRECT.lRight - WinRECT.lLeft), _
                                                0, _
                                                (WinRECT.lRight - WinRECT.lLeft), _
                                                (WinRECT.lBottom - WinRECT.lTop), _
                                                1
                                .SetFocus
                                SetMousePosOverWindow s_lwHndl, [WinPos RightTop]
                            End With

                            Call m_frmTimer.SetTimerHotcorners(lDelayRemoveHotCorners)
                            Set s_LastWinObj = m_objPropertiesWindow
                            s_FW_State = [State Win Rgt Top Opened]

                            Exit Sub
                        End If


                        ' === RIGHT / Bottom :   Project Explorer ===
                        If .pY = Screen_Y Then
                            If m_objProjectExplorerWindow Is Nothing Then           ' Sometimes call in startup goes wrong ... ;(
                                GetRefsToToolWindows                                ' Don't know why, it's hard to track!
                            End If                                                  ' Then we try here ones more.
                            If m_objProjectExplorerWindow Is Nothing Then           ' Just for sure ... ;)
                                s_FW_State = [State Idle]
                                
                                Exit Sub
                            End If
                        
                            With m_objProjectExplorerWindow
                                .Visible = True
                                s_lwHndl = API_FindWindowEx(lwHndlVBIDEMDIClient, ByVal 0&, "DockingView", .Caption)
                                
                                GetWindowRect lwHndlVBIDEMDIClient, RectMDIClient
                                GetWindowRect s_lwHndl, WinRECT
                                
                                ' Don't know why, but simple setting of position by properties doesn't work here ;(
                                API_MoveWindow s_lwHndl, _
                                                RectMDIClient.lRight - (WinRECT.lRight - WinRECT.lLeft), _
                                                ((RectMDIClient.lBottom - RectMDIClient.lTop) - (WinRECT.lBottom - WinRECT.lTop)), _
                                                (WinRECT.lRight - WinRECT.lLeft), _
                                                (WinRECT.lBottom - WinRECT.lTop), _
                                                1
                                .SetFocus
                                SetMousePosOverWindow s_lwHndl, [WinPos RightBottom]
                            End With

                            Call m_frmTimer.SetTimerHotcorners(lDelayRemoveHotCorners)
                            Set s_LastWinObj = m_objProjectExplorerWindow
                            s_FW_State = [State Win Rgt Btm Opened]

                            Exit Sub
                        End If
                        
                        
                        ' ***************************************
                        ' * Gimmick No 5:    Right Screen Border
                        ' ***************************************
                        
                        ' We are at the RIGHT border of the screen, but not in a corner
                        If flgActivateHotSides = True Then
                            If .pY > 30 And .pY < Screen_Y - 30 Then        ' A little gap to top and bottom?
                                
                                On Error Resume Next
                                
                                If s_flgNoRepeat = False Then               ' Prevent multiple raises
                                    
                                    ' Is Ctrl key pressed?
                                    If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) Then
                                        
                                        ' Close topmost window
                                        DoEvents
                                        SendKeys "^{F4}", True
                                    
                                    Else
                                        
                                        Set oActiveCodeWindow = VBInstance.ActiveCodePane.Window
                                        If Not oActiveCodeWindow Is Nothing Then
                                        
                                            ' Increase window to fill whole IDE client area (but don't set its state to 'Max' !)
                                            With oActiveCodeWindow
                                                .Left = -5
                                                .Top = -4
                                                .Width = VBInstance.MainWindow.Width - 7
                                                .Height = VBInstance.MainWindow.Height - 102
                                                DoEvents
                                            End With
                                        End If
                                    End If
                                    s_flgNoRepeat = True            ' Prevent multiple raises
                                End If
                                
                                Exit Sub
                            End If
                        End If

                    End If

                    ' Not in a 'hot area', so reset the flag
                    s_flgNoRepeat = False



            Case [State Dlg Opened]
                    ' We poped up a window (normaly the property window) and from there,
                    ' the user has opened a dialog (like StdColorSelector, StdFontSelecor,
                    ' StdFileOpen, ...).
                    '
                    ' We check here: Is this dialog still open ? -> Yes? So keep waiting for closing!
                    ' No? (closed):  Return to saved state [State Win X Y Opened]

                    ' [It nearly doubled the time for me (for this feature of FF) to get a solution for this problem
                    ' and to write this few lines ... ;((( ]

                    If Not VBInstance.ActiveWindow Is Nothing Then
                        With VBInstance.ActiveWindow
                            If .Caption <> VBInstance.MainWindow.Caption Then
                                s_FW_State = s_SavedState
                                .SetFocus
                                DoEvents
                                s_LastWinObj.SetFocus

                                Exit Sub
                            End If
                        End With
                    End If


            Case [State Win Lft Top Opened], [State Win Lft Btm Opened], [State Win Rgt Top Opened], [State Win Rgt Btm Opened]

                    ' Get position rectangle of current opened window
                    If GetWindowRect(s_lwHndl, WinRECT) Then
                        ' Is the mouse pointer within this rectangle ?
                        If PtInRect(WinRECT, MousePos.pX, MousePos.pY) = 0 Then

                            ' Is Ctrl key pressed ?
                            If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) = True Then
                                ' Yes! Don't auto remove windows anymore. Just leave it opened on screen.
                                Call m_frmTimer.SetTimerHotcorners(lDelayShowHotCorners)
                                s_FW_State = [State Idle]

                                Exit Sub
                            End If

                            ' Is left mouse button not pressed (No dragging) ?
                            If GetAsyncKeyState(vbLeftButton) = 0 Then
                                ' No! Hide away the window
                                s_LastWinObj.Visible = False
                                Call m_frmTimer.SetTimerHotcorners(lDelayShowHotCorners)
                                s_FW_State = [State Idle]
                            End If

                        End If
                    Else
                        s_FW_State = [State Idle]
                    End If

        End Select

    End With

    Call m_frmTimer.SetTimerHotcorners(lDelayShowHotCorners)
    

    Exit Sub


error_handler:
    
    If Format(Err.Number) = "-2147467259" Then
        MsgBox "Please don't use Flying Windows hotcorners with docked windows." + vbCrLf + _
                "This raises an internal error. Please goto to VB IDEs [Options ...]" + vbCrLf + _
                "dialog in <Tools> menu. Select the docking tab there and switch of the" + vbCrLf + _
                "docking mode for the following windows:" + vbCrLf + _
                "'Immediate Window', 'Project Explorer', 'Properties Window' and 'Toolbox'." + vbCrLf + _
                "Maybe a restart of the Flying Windows addin is needed afterwards." + vbCrLf + vbCrLf + _
                "Thank you.", vbExclamation + vbMsgBoxSetForeground, _
              " Error Message from Addin 'Flying Windows' :   'Windows in docked mode found'"
    Else
        MsgBox Err.Description, vbExclamation + vbMsgBoxSetForeground, _
              " Flying Windows:  Error in 'm_frmTimer_TimerHotcornersElapsed()'"
    End If
    s_FW_State = [State Idle]
    Call m_frmTimer.SetTimerHotcorners(lDelayShowHotCorners)

End Sub


Private Sub SetMousePosOverWindow(hWnd As Long, WinPos As enWindowPosition)
    ' This puts the mouse pointer over the window. Position is a quater to enWindowPosition.
    
    Const MOUSE_GAP As Long = 20&
    
    Dim WinRECT As RECTAPI
    
    If IsWindow(hWnd) Then
        GetWindowRect hWnd, WinRECT
        With WinRECT
            Select Case WinPos
                Case [WinPos LeftTop]
                        SetCursorPos .lLeft + MOUSE_GAP, .lTop + MOUSE_GAP
                        
                Case [WinPos LeftBottom]
                        SetCursorPos .lLeft + MOUSE_GAP, .lTop + ((.lBottom - .lTop) - MOUSE_GAP)
                        
                Case [WinPos RightTop]
                        SetCursorPos .lLeft + ((.lRight - .lLeft) - MOUSE_GAP), .lTop + MOUSE_GAP
                        
                Case [WinPos RightBottom]
                        SetCursorPos .lLeft + ((.lRight - .lLeft) - MOUSE_GAP), .lTop + ((.lBottom - .lTop) - MOUSE_GAP)
            End Select
        End With
    End If
    
End Sub


' **************************************************
' *                 TOOLTIPS PART                  *
' **************************************************
Private Sub m_frmTimer_TimerToolTipsElapsed()
    ' All the stuff needed for tooltips handling and to show the mouse pointer position ...
    
    ' Folks, This was really hard work ;) !

    ' Once again:   Execution speed beats code length.
    
    
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER As Long = &H4
    Const SWP_NOOWNERZORDER As Long = &H200
    
    Static s_LastMousePos   As POINTAPI
    Static s_TT_Mode        As enFW_TT_State
    Static RECT_MouseArea   As RECTAPI
    Static lCountDown       As Long
    Static flgCBincreased   As Boolean
    
    Dim MousePos            As POINTAPI
    Dim lwHndl              As Long
    Dim lParentwHndl        As Long
    Dim sDesgnrWinCaption   As String
    Dim VBIDE_Commponent    As VBComponent
    Dim VBIDE_Form          As VBForm
    Dim VBIDE_Ctrl          As VBControl
    Dim RECT_Ctrl           As RECTAPI
    Dim RECT_ContainerCtrl  As RECTAPI
    Dim POS_LookedCtrl      As POINTAPI
    Dim lwHndlLookedCtrl    As Long
    Dim lwHndlContainedCtrl As Long
    Dim lContainedCtrls     As Long
    Dim lThndrFrmwHndl      As Long
    Dim OFFSET_ThndrFrm     As POINTAPI
    Dim sWinClassName       As String
    Dim flgAbortLoops       As Boolean
    Dim lCtrls              As Long
    Dim i                   As Long
    Dim k                   As Long
    Dim ContainerCtrl       As VBControl
    Dim DummyChkCtrl        As VBControl
    Dim lwHndlContainerCtrl As Long
    Dim ContainedCtrl       As VBControl
    Dim lScaleMode          As Long
    Dim lOffsetX            As Long
    Dim lOffsetY            As Long
    Dim lDeep               As Long
    Dim sClassName          As String
    Dim lHndlTopWindow      As String
    Dim sTitle              As String
    Dim MainWin             As Window
    Dim RECT_Win            As RECTAPI
    
    
    On Local Error Resume Next
    
    Set MainWin = VBInstance.MainWindow
    If MainWin Is Nothing Then                      ' In special situations I got trouble without this ...
                       
        Exit Sub
    End If
    
    On Error GoTo error_handler

    ' Right mode to do something ?
    If s_TT_Mode = [State TT Idle] Then

        ' If IDE in 'RUN' mode: Do nothing!
        If IDEmode <> vbext_vm_Design Then

            Exit Sub
        End If

        ' Minimized IDE? Do nothing!
        If MainWin.WindowState = vbext_ws_Minimize Then

            Exit Sub
        End If

        ' IDE at foreground ? No: Do nothing!
        lHndlTopWindow = GetForegroundWindow()
        Do While lHndlTopWindow <> lwHndlVBIDEmain
            lHndlTopWindow = GetParent(lHndlTopWindow)

            ' Top most window reached ?
            If lHndlTopWindow = 0 Then

                ' Yes!  > Leave this!
                Exit Sub
            End If
        Loop

    End If

    ' Get mouse X/Y
    GetCursorPos MousePos

    
    ' Little add-on 1 :  Show mouse' absolut screen coordinates in VB IDEs title bar in pixel
    If flgActivatePosition = True Then
        sTitle = MainWin.Caption
        sTitle = Left$(sTitle, InStr(31, sTitle, "]")) & "        " & MousePos.pX & " / " & MousePos.pY & vbNullChar
        SetWindowText lwHndlVBIDEmain, sTitle
    End If
    
    
    ' Little add-on 2 :  Increase dropdown part of comboboxes of current code windows to show more values
    ' (This was a hard one, too. The API call 'MoveWindow()' leaded me into terrible labyrinths of
    ' workarrounds and problems ... At the end I check it out with 'SetWindowPos() with success.  )
    If flgActivateLargeCombos = True Then
        
        On Error Resume Next
        If Not VBInstance.ActiveCodePane.Window Is Nothing Then
            
            ' Look for combobox
            lwHndl = API_FindWindow("ComboLBox", vbNullString)          ' ComboLBox is an interessting class - not VB standard ...
            If lwHndl <> 0 Then                                         ' and this windows parent's handle is 128 always ...
                If API_IsWindowVisible(lwHndl) Then
                
                    ' Not yet increased?
                    If flgCBincreased = False Then
                        GetWindowRect lwHndl, RECT_Win
                        
                        API_SetWindowPos lwHndl, 0, 0, 0, RECT_Win.lRight - RECT_Win.lLeft, lMaxComboSize, _
                                SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
                                            
                        flgCBincreased = True
                    End If
                Else
                    flgCBincreased = False
                End If
            End If
        Else
            flgCBincreased = False
        End If
        On Error GoTo error_handler
    End If
   
    
    
    ' Tooltips deactivated by options? Leave sub ...
    If flgActivateToolTips = False Then
    
        Exit Sub
    End If
    

    ' Tooltip Window already open ?
    If s_TT_Mode = [State TT Win Open] Then

        ' Mouse still in same area as before ?
        If PtInRect(RECT_MouseArea, MousePos.pX, MousePos.pY) = 0 Then

            ' Pressing Ctrl key prevents from removing
            If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) = True Then
                s_TT_Mode = [State TT Win Keep Open]

                Exit Sub
            End If

            ' No. Remove tooltip window
            frmTTips.UnloadTT
            Call m_frmTimer.SetTimerToolTips(lDelayShowTooltips)
            s_TT_Mode = [State TT Idle]

        End If

        Exit Sub
    End If


    ' So s_TT_Mode = [State TT Win Keep Open] or [State TT Idle] ...
    With MousePos

        ' Change in X ?
        If .pX <> s_LastMousePos.pX Then
            s_LastMousePos.pX = .pX
            s_LastMousePos.pY = .pY
            lCountDown = 2

            Exit Sub
        End If

        ' Change in Y ?
        If .pY <> s_LastMousePos.pY Then
            s_LastMousePos.pX = .pX
            s_LastMousePos.pY = .pY
            lCountDown = 2

            Exit Sub
        End If


        ' ### Here we know:  The mouse pointer is at the same position for mostly one timer cycle

        ' Prevent another check for the same position
        lCountDown = lCountDown - 1
        If lCountDown < 1 Then
            lCountDown = 1

            Exit Sub
        End If


        ' Find the right MS VB 'DesignerWindow' by 'Window Classname', than look for the control
        ' (Controls without a wHndl (Line, Label, Image) forced me to rewrite this whole stuff more than once ... :((( NO FUN! )
        lwHndl = WindowFromPoint(.pX, .pY)

        lParentwHndl = lwHndl
        sWinClassName = GetClassNameByWHndl(lParentwHndl)

        lDeep = -1
        Do
            If sWinClassName = "ThunderForm" Or _
                    sWinClassName = "ThunderMDIForm" Or _
                    sWinClassName = "ThunderUserControl" Or _
                    sWinClassName = "ThunderPropertyPage" Then
                    
                ' We need this to get the pixel offsets for control coordinates
                lThndrFrmwHndl = lParentwHndl
            End If

            lParentwHndl = GetParent(lParentwHndl)

            ' Top most window reached ?
            If lParentwHndl = 0 Then

                ' Yes!  > Leave this!
                Exit Sub
            End If

            sWinClassName = GetClassNameByWHndl(lParentwHndl)
            lDeep = lDeep + 1

        Loop While sWinClassName <> "DesignerWindow"

        ' lDeep = 0 > Mouse over the form or over a control without a wHndl

        ' We are over a 'DesignerWindow'
        sDesgnrWinCaption = GetWindowsNameByWHndl(lParentwHndl)

        ' Look for this 'DesignerWindow' Window in VBE IDE Window Collection
        flgAbortLoops = False
        For Each VBIDE_Commponent In VBInstance.ActiveVBProject.VBComponents
            With VBIDE_Commponent
                If .Type = vbext_ct_VBForm Or _
                        .Type = vbext_ct_VBMDIForm Or _
                        .Type = vbext_ct_UserControl Or _
                        .Type = vbext_ct_PropPage Then
                        
                    If .HasOpenDesigner = True Then
                        Set VBIDE_Form = .Designer
                        Call m_frmTimer.SetTimerToolTips(0)
                        With VBIDE_Form
                            If .Parent.DesignerWindow.Caption = sDesgnrWinCaption Then
                                ' Here we should have the WINDOW we are looking for

                                ' Get screen coordinates of client area (offset in pixel)
                                ClientToScreen lThndrFrmwHndl, OFFSET_ThndrFrm

                                lCtrls = .VBControls.Count


                                ' ### CASE 1:  Mouse over a ctrl without a wHndl (Line, Shape, Image, Label) NOT nested
                                If lwHndl = lThndrFrmwHndl Then
                                    For i = lCtrls To 1 Step -1    ' from top to bottom to get top ctrls first (nesteds ...)

                                        Set VBIDE_Ctrl = .VBControls.Item(i)
                                        sClassName = VBIDE_Ctrl.ClassName

                                        Select Case sClassName

                                        Case "Label", "Image", "Shape", "Line", "OLE", "CMDialogWndClass"

                                                RECT_Ctrl = GetCtrlsCoordinates(VBIDE_Ctrl, VBIDE_Form)
                                                OffsetRect RECT_Ctrl, OFFSET_ThndrFrm.pX, OFFSET_ThndrFrm.pY
                                                If PtInRect(RECT_Ctrl, MousePos.pX, MousePos.pY) Then
                                                    s_TT_Mode = ShowTT(VBIDE_Ctrl, lParentwHndl, lDeep, _
                                                            MousePos.pX, MousePos.pY, RECT_MouseArea)

                                                    Exit Sub
                                                End If

                                        End Select

                                    Next i

                                    Call m_frmTimer.SetTimerToolTips(lDelayShowTooltips)

                                    Exit Sub    ' We must be over empty ground of the DesignerWindow: Search is over!
                                End If



                                ' ### CASE 2:  Mouse over a standard ctrl with a wHndl (not Line, Image or Label, ...)
                                '              which is NOT a container

                                ' There are ctrls (like combos, ...) with sub-windows. Get the true handle
                                lwHndlLookedCtrl = lwHndl
'                                Do While Left$(GetClassNameByWHndl(lwHndlLookedCtrl), 7) <> "Thunder"
'                                    lwHndlLookedCtrl = GetParent(lwHndlLookedCtrl)
'                                Loop   '######## Something like this... very hard to get all! #####

                                ClientToScreen lwHndlLookedCtrl, POS_LookedCtrl
                                ClipLftTopToParent lwHndlLookedCtrl, POS_LookedCtrl, lThndrFrmwHndl

                                For i = lCtrls To 1 Step -1    ' from top to bottom to get top ctrls first (nesteds ...)
                                    Set VBIDE_Ctrl = .VBControls.Item(i)
                                    If VBIDE_Ctrl.Container Is VBIDE_Form Then
                                        sClassName = VBIDE_Ctrl.ClassName

                                        Select Case sClassName

                                        Case "Label", "Image", "Line", "OLE", "Menu", "CMDialogWndClass"
                                                ' Do nothing

                                        Case Else

                                                If VBIDE_Ctrl.ContainedVBControls.Count = 0 Then
                                                    RECT_Ctrl = GetCtrlsCoordinates(VBIDE_Ctrl, VBIDE_Form)
                                                    OffsetRect RECT_Ctrl, OFFSET_ThndrFrm.pX, OFFSET_ThndrFrm.pY

                                                    ' Compare coordinates (using ABS(delta) < 4 to equalize rounding diffs)
                                                    If Abs(POS_LookedCtrl.pX - RECT_Ctrl.lLeft) < 4 And _
                                                            Abs(POS_LookedCtrl.pY - RECT_Ctrl.lTop) < 4 Then

                                                        s_TT_Mode = ShowTT(VBIDE_Ctrl, lParentwHndl, lDeep, _
                                                                MousePos.pX, MousePos.pY, RECT_MouseArea)

                                                        Exit Sub
                                                    End If
                                                End If
                                        End Select

                                    End If
                                Next i



                                ' ### CASE 3:   Mouse must be over container ctrl with contents
                                '               (Let's get recursive and do the same job again ;) )

                                For i = lCtrls To 1 Step -1
                                    Set VBIDE_Ctrl = .VBControls.Item(i)
                                    With VBIDE_Ctrl
                                        lContainedCtrls = .ContainedVBControls.Count
                                        If lContainedCtrls > 0 And .Container Is VBIDE_Form Then    ' Container ontop of form ?

                                            ' Is the mouse over this container ?
                                            RECT_ContainerCtrl = GetCtrlsCoordinates(VBIDE_Ctrl, VBIDE_Form)
                                            OffsetRect RECT_ContainerCtrl, OFFSET_ThndrFrm.pX, OFFSET_ThndrFrm.pY
                                            If PtInRect(RECT_ContainerCtrl, MousePos.pX, MousePos.pY) Then

                                                ' Searched ctrl must be this container or a (sub-)child of him

                                                lwHndlContainerCtrl = WindowFromPoint(RECT_ContainerCtrl.lLeft, _
                                                        RECT_ContainerCtrl.lTop)

                                                For k = lCtrls To 1 Step -1

                                                    Set ContainedCtrl = VBIDE_Form.VBControls.Item(k)
                                                    If IsChildOf(ContainedCtrl, VBIDE_Ctrl, VBIDE_Form) Then

                                                        RECT_Ctrl = GetCtrlsCoordinates(ContainedCtrl, VBIDE_Form)
                                                        ' Clip to containers left top border
                                                        With RECT_Ctrl
                                                            If .lLeft < 0 Then
                                                                .lLeft = 1
                                                            End If
                                                            If .lTop < 0 Then
                                                                .lTop = 1
                                                            End If
                                                        End With

                                                        On Error Resume Next
                                                        Set ContainerCtrl = ContainedCtrl
                                                        Do Until ContainerCtrl.Container Is VBIDE_Ctrl
                                                            
                                                            Set DummyChkCtrl = Nothing
                                                            On Error Resume Next
                                                            Set DummyChkCtrl = ContainerCtrl.Container ' A form isn't a container...
                                                            If DummyChkCtrl Is Nothing Then
                                                            
                                                                Exit Do
                                                            End If
                                                            
                                                            Set ContainerCtrl = ContainerCtrl.Container
                                                            With ContainerCtrl.Properties
                                                                lOffsetX = .Item("left")
                                                                lOffsetY = .Item("top")
                                                            End With

                                                            ' Unit conversion neccessary ?
                                                            lScaleMode = vbTwips
                                                            On Error Resume Next
                                                            lScaleMode = ContainerCtrl.Container.Properties.Item("ScaleMode")
                                                            On Error GoTo error_handler
                                                            If lScaleMode <> vbPixels And lScaleMode <> 0 Then
                                                                  ' Debug.Print "Yep: " & lOffsetX, lScaleMode
                                                                lOffsetX = frmTimer.ScaleX(lOffsetX, lScaleMode, vbPixels)
                                                                lOffsetY = frmTimer.ScaleY(lOffsetY, lScaleMode, vbPixels)
                                                            End If

                                                            ' Add containers position as offset
                                                            OffsetRect RECT_Ctrl, lOffsetX, lOffsetY

                                                        Loop
                                                        On Error GoTo error_handler

                                                        OffsetRect RECT_Ctrl, RECT_ContainerCtrl.lLeft + 2, RECT_ContainerCtrl.lTop + 2

                                                        sClassName = ContainedCtrl.ClassName
                                                        If sClassName = "Label" Or sClassName = "Image" Or sClassName = "Line" Then

                                                            ' #  CASE 3.1:   Mouse over a nested ctrl without a wHndl

                                                            ' Mouse over ctrls rectangle ?
                                                            If PtInRect(RECT_Ctrl, MousePos.pX, MousePos.pY) Then
                                                                s_TT_Mode = ShowTT(ContainedCtrl, lParentwHndl, lDeep + 1, _
                                                                        MousePos.pX, MousePos.pY, RECT_MouseArea)

                                                                Exit Sub
                                                            End If

                                                        ElseIf sClassName <> "Menu" Then

                                                            ' CASE 3.2 :  Mouse over a nested standard ctrl with a wHndl

                                                            ' Reusing first part of CASE 2 ! (POS_LookedCtrl)

                                                            ' Compare coordinates (Using ABS(delta) < 4 to equalize rounding diffs)

                                                            lwHndlContainedCtrl = WindowFromPoint(RECT_Ctrl.lLeft, RECT_Ctrl.lTop)

                                                            If lwHndlContainedCtrl <> lwHndlContainerCtrl Then
                                                                If Abs(POS_LookedCtrl.pX - RECT_Ctrl.lLeft) < 4 And _
                                                                        Abs(POS_LookedCtrl.pY - RECT_Ctrl.lTop) < 4 Then

    '                                                            If lwHndlContainedCtrl = lwHndl Then
                                                                    s_TT_Mode = ShowTT(ContainedCtrl, lParentwHndl, lDeep, _
                                                                            MousePos.pX, MousePos.pY, RECT_MouseArea)

                                                                    Exit Sub
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next k

                                                ' It 's not a (sub-)child, so it must be the container by itself
                                                s_TT_Mode = ShowTT(VBIDE_Ctrl, lParentwHndl, lDeep, _
                                                    MousePos.pX, MousePos.pY, RECT_MouseArea)

                                                Exit Sub
                                            End If  '  PtInRect(RECT_ContainerCtrl, MousePos.pX, MousePos.pY)

                                        End If
                                    End With
                                Next i
                            End If
                        End With
                    End If
                End If
            End With
        Next VBIDE_Commponent
    End With                                ' Hey: That's about 16 closings ! ;) My new record!
                                            ' It is the ugliest, but maybe best code I've ever written...
                                            ' MS's VB IDE baricades are very high! :(
                                            
    Call m_frmTimer.SetTimerToolTips(lDelayShowTooltips)

    Exit Sub


error_handler:

    MsgBox Err.Description, vbExclamation + vbMsgBoxSetForeground, "  Flying Windows:  Error in 'm_frmTimer_TimerToolTipsElapsed()'"
    s_TT_Mode = [State TT Idle]
    Call m_frmTimer.SetTimerToolTips(lDelayShowTooltips)
    
End Sub


Private Function ShowTT(ByRef VBIDE_Ctrl As VBControl, _
                        ByVal lParentwHndl As Long, _
                        ByVal lDeep As Long, _
                        ByVal lMouseX As Long, _
                        ByVal lMouseY As Long, _
                        ByRef RECT_MouseArea As RECTAPI) As enFW_TT_State

    ' Displays a new ToolTips window


    Set frmTTips = New frmToolTips
    With frmTTips
        Set .CtrlToDisplay = VBIDE_Ctrl
        Set .RefToVBinst = VBInstance
        .lOwnerWHndl = lParentwHndl
        .lNested = IIf(lDeep > 0, lDeep - 1, 0)

        .Left = frmTTips.ScaleX(lMouseX + 20, vbPixels, vbTwips)
        .Top = frmTTips.ScaleY(lMouseY + 20, vbPixels, vbTwips)

        ' Show tooltip window, but don't deactivate main form (change focus)
        ShowWindow .hWnd, SW_SHOWNA

        ' Prepare for check: Mouse moved only a little
        With RECT_MouseArea
            .lLeft = lMouseX - 6
            .lRight = lMouseX + 6
            .lTop = lMouseY - 4
            .lBottom = lMouseY + 4
        End With

        Call m_frmTimer.SetTimerToolTips(lDelayRemoveTooltips)


    End With

    ShowTT = [State TT Win Open]

End Function


Private Function GetWindowsNameByWHndl(ByVal hWnd As Long) As String
    ' Result: Windows name (caption) for a given hwnd
    '         vbNullString for invalid hwnd

    Dim sBuffer As String
    Dim lResult As Long


    If IsWindow(hWnd) = 0 Then

        Exit Function
    End If

    sBuffer = String$(256, vbNullChar)
    lResult = GetWindowText(hWnd, sBuffer, 255)
    If lResult Then
        GetWindowsNameByWHndl = Left$(sBuffer, lResult)
    End If

End Function


Private Function GetCtrlsCoordinates(TheCtrl As VBControl, VBIDE_Form As VBForm) As RECTAPI

    Dim lSwap           As Long
    Dim lScaleMode      As Long

    On Local Error Resume Next      ' Very important here ... :(

    With TheCtrl
        Select Case .ClassName

        Case "Menu"
                
                Exit Function


        Case "Line"
                GetCtrlsCoordinates.lLeft = .Properties.Item("X1")
                GetCtrlsCoordinates.lTop = .Properties.Item("Y1")
                GetCtrlsCoordinates.lRight = .Properties.Item("X2")
                GetCtrlsCoordinates.lBottom = .Properties.Item("Y2")

                ' Need to sort ?
                With GetCtrlsCoordinates
                    If .lLeft > .lRight Then
                        lSwap = .lLeft
                        .lLeft = .lRight
                        .lRight = lSwap
                    End If
                    If .lTop > .lBottom Then
                        lSwap = .lTop
                        .lTop = .lBottom
                        .lBottom = lSwap
                    End If
                End With


        Case Else
                GetCtrlsCoordinates.lLeft = .Properties.Item("left")
                GetCtrlsCoordinates.lTop = .Properties.Item("top")

                ' There are ctrls without (width and height, e.g. Timers, ... ;( ) For this we have "On Error Resume Next"
                GetCtrlsCoordinates.lRight = GetCtrlsCoordinates.lLeft + .Properties.Item("width")
                GetCtrlsCoordinates.lBottom = GetCtrlsCoordinates.lTop + .Properties.Item("height")
                
        End Select


        ' Handle scalemode problem (frames and MDI forms have no scalemode ...)
        lScaleMode = vbTwips
        If .Container Is VBIDE_Form Then
            lScaleMode = VBIDE_Form.Parent.Properties.Item("ScaleMode")
        Else
            lScaleMode = .Container.Properties.Item("ScaleMode")
        End If

        If lScaleMode = vbPixels Then
            
            Exit Function
        End If

        With GetCtrlsCoordinates
            .lLeft = frmTimer.ScaleX(.lLeft, lScaleMode, vbPixels)
            .lTop = frmTimer.ScaleY(.lTop, lScaleMode, vbPixels)
            .lRight = frmTimer.ScaleX(.lRight, lScaleMode, vbPixels)
            .lBottom = frmTimer.ScaleY(.lBottom, lScaleMode, vbPixels)
        End With

    End With

End Function


Private Sub ClipLftTopToParent(lwHndl As Long, ByRef POS_Ctrl As POINTAPI, lwHndlForm As Long)

    Dim lwHndlParent    As Long
    Dim POS_Parent      As POINTAPI

    lwHndlParent = GetParent(lwHndl)
    If lwHndlParent = lwHndlForm Then         ' Do nothing with ctrls directly on the form

        Exit Sub
    End If

    ClientToScreen lwHndlParent, POS_Parent

    With POS_Parent
        If POS_Ctrl.pX < .pX Then
            POS_Ctrl.pX = .pX
        End If

        If POS_Ctrl.pY < .pY Then
            POS_Ctrl.pY = .pY
        End If
    End With

End Sub


Private Function GetClassNameByWHndl(ByVal hWnd As Long) As String
    ' Result: Class name for a given hwnd
    '         "" for invalid hwnd

    Dim sBuffer As String
    Dim lResult As Long


    If IsWindow(hWnd) = 0 Then

        Exit Function
    End If

    sBuffer = String$(255, vbNullChar)
    lResult = API_GetClassName(hWnd, sBuffer, 255)
    If lResult Then
        GetClassNameByWHndl = Left$(sBuffer, lResult)
    End If

End Function


Private Function IsChildOf(TheCtrl As VBControl, ContainerCtrl As VBControl, TheForm As VBForm) As Boolean
    ' Result: TRUE, if 'TheCtrl' is a child (or sub-child) of 'ContainerCtrl'

    Dim ParentCtrl As VBControl

    On Error GoTo error_handler

    IsChildOf = False

    If Not TheCtrl.Container Is TheForm Then
        Set ParentCtrl = TheCtrl.Container
        Do
            If ParentCtrl = ContainerCtrl Then
                IsChildOf = True

                Exit Function
            End If

            If Not ParentCtrl.Container Is TheForm Then
                Set ParentCtrl = ParentCtrl.Container
            Else

                Exit Do
            End If
        Loop
    End If

    On Error GoTo 0

    Exit Function


error_handler:

End Function





' **************************************************
' *                 AUTOTEXT PART                  *
' **************************************************
Private Sub m_frmTimer_TimerAutoText()
    ' All the stuff needed to autocomplete text.
    ' This sub handles the calling: When the F12 (Function Key No 12)
    ' is pressed AND released the AutoText sub will be called.
        
    Const VK_F12 As Long = &H7B
    
    Static flgKeysPressed As Boolean
    
    If IDEmode = vbext_vm_Run Then
        
        Exit Sub
    End If
    
    If flgKeysPressed = False Then
        
        ' Is F12 pressed?
        If GetAsyncKeyState(VK_F12) And &H8000 Then                 ' Of course you can change this key to a different one.
            
            ' Set trigger
            flgKeysPressed = True
                
        End If
    Else
    
        ' Is F12 released?
        If Not (GetAsyncKeyState(VK_F12) And &H8000) Then
                
            ' Disable timer - no more key checks
            m_frmTimer.SetTimerAutoText 0
            
            ' Do the job
            DoAutoText
            
            ' Reset trigger
            flgKeysPressed = False
            
            ' Enable timer - we start waiting again
            m_frmTimer.SetTimerAutoText lDelayAutoText
        End If
    
    End If

End Sub


Private Sub DoAutoText()
    
    Dim lStartLine      As Long
    Dim lEndLine        As Long
    Dim lStartCol       As Long
    Dim lEndCol         As Long
    Dim sCurLine        As String
    Dim sWordLeft       As String           ' The word left to the current cursor
    Dim lEntries        As Long
    Dim i               As Long
    Dim sSubstCode      As String

    On Error GoTo DoAutoText_Error
   

    ' In code windows only!
    If VBInstance.ActiveWindow.Type <> vbext_wt_CodeWindow Then
    
        Exit Sub
    End If

    With VBInstance.ActiveCodePane
        .GetSelection lStartLine, lStartCol, lEndLine, lEndCol
        
        ' === The CHECK PART - Any miss leads to a 'Leave Sub'
        
        ' If something is selected or to near to left -> leave!
        If lStartLine <> lEndLine Or lStartCol <> lEndCol Or lStartCol < 2 Then
        
            Exit Sub
        End If
    
        ' Lets get the word left to current cursors position
        sCurLine = .CodeModule.Lines(lStartLine, 1)
        
        ' We don't handle empty lines
        If Trim$(sCurLine) = "" Then
        
            Exit Sub
        End If

        If Mid$(sCurLine, lStartCol, 1) <> " " And Len(sCurLine) > lStartCol Then
            
            Exit Sub
        End If
        
        ' Cut off right part
        sWordLeft = Left$(sCurLine, lStartCol - 1)
        
        ' Cursor must be immediatly behind a word, not behind a <space>
        If Right$(" " + sWordLeft, 1) = " " Then
        
            Exit Sub
        End If
        
        ' Cut off from start of word to beginning of line
        sWordLeft = Trim$(" " + Mid$(sWordLeft, InStrRev(sWordLeft, " ") + 1))
        
        ' We have a word?
        If sWordLeft = "" Then
        
            Exit Sub
        End If
        
        
        ' === The ACTION part - All checks succeeded
        
        ' Search for this key word in AutoText substitution list
        lEntries = UBound(arrAutoText)
        For i = 1 To lEntries
            If arrAutoText(i).sKey = sWordLeft Then
                On Error GoTo AutoTextSK_Error
                
                ' Select (key)word left to cursor
                SendKeys "+{LEFT " & Len(sWordLeft) & "}", True
                DoEvents
                
                ' Get substitution string and replace keywords (if there) with current values
                sSubstCode = arrAutoText(i).sSubstCode
                
                With VBInstance.ActiveVBProject
                    sSubstCode = Replace$(sSubstCode, "%AUTHOR%", AT_AUTHOR)
                    sSubstCode = Replace$(sSubstCode, "%DATE%", Format(Now, AT_DATEFORMAT))
                    sSubstCode = Replace$(sSubstCode, "%TIME%", Format(Now, AT_TIMEFORMAT))
                    sSubstCode = Replace$(sSubstCode, "%PROJECTDESCRIPTION%", .Description)
                    sSubstCode = Replace$(sSubstCode, "%PROJECTFILENAME%", .FileName)
                    sSubstCode = Replace$(sSubstCode, "%PROJECTNAME%", .Name)
                End With
                  
                ' NOW substitute the keyword with the new code
                SendKeys sSubstCode, True
                DoEvents

                Exit For
            End If
        Next i
        
    End With


    Exit Sub


DoAutoText_Error:
    
    Err.Clear
    
    Exit Sub
    
    
AutoTextSK_Error:
    
    MsgBox "Error in your file 'FW_AutoText.Txt'  at Section  [" + sWordLeft + "]" + vbCrLf + _
            "Replacement definition raises an error in SendKeys() !" + vbCrLf + _
            "Please open Flying Windows Option dialog an edit your file. Thank you.", _
            vbExclamation + vbMsgBoxSetForeground, _
            " Addin Flying Windows VB6:  Definition Error in Config File"

    Err.Clear

End Sub


Public Sub LoadAutoText()
    ' Load and parse contents of file 'FW_AutoText.Txt' into array

    Dim fHndl       As Integer
    Dim lCounter    As Long
    Dim lLineNo     As Long
    Dim sPathFName  As String
    Dim sLine       As String
    Dim sTrimedLine As String
    
    
    On Local Error GoTo LoadAutoText_Error
    
        
    sPathFName = App.Path + "\FW_AutoText.txt"
    
    ' If there is no definition file -> disable this feature
    If Dir$(sPathFName) = "" Then
        flgActivateAutoText = False
                
        Exit Sub
    End If
    
    Erase arrAutoText
    
    fHndl = FreeFile
    Open sPathFName For Input As #fHndl
        Do While EOF(fHndl) = False
            Line Input #fHndl, sLine
            lLineNo = lLineNo + 1
            
            sTrimedLine = Trim$(sLine)
            
            Select Case True
            
                Case sTrimedLine = ""
                        ' == Just ignore empty lines ==
                        
                Case Left$(sTrimedLine, 1) = ";"
                        ' == Just ignore comment lines ==
            
                Case Left$(sTrimedLine, 1) = "[" And Right$(sTrimedLine, 1) = "]" And Len(sTrimedLine) > 2
                        ' == Start a new entry ==
                        lCounter = lCounter + 1
                        ReDim Preserve arrAutoText(1 To lCounter)
                        arrAutoText(lCounter).sKey = Mid$(sLine, 2, Len(sLine) - 2)
                        
                Case lCounter > 0
                        ' == Add to current entry ==
                        arrAutoText(lCounter).sSubstCode = arrAutoText(lCounter).sSubstCode + sLine
                        
                Case Else
                        MsgBox "Error in line #" & lLineNo & "in file 'FW_AutoText.Txt' : " & vbCrLf & "[" & sLine & "]", _
                                vbExclamation + vbMsgBoxSetForeground, " Error in StartUp of AddIn 'FlyingWindows VB6' :"
                
                End Select
                
        Loop
    Close #fHndl
    
    ' No entries -> disable feature
    If lCounter < 1 Then
        flgActivateAutoText = False
    End If
    
    Exit Sub


LoadAutoText_Error:

    MsgBox "Error [" & Err.Description & "] when loading 'FW_AutoText.Txt'", _
            vbExclamation + vbMsgBoxSetForeground, " Error in StartUp of AddIn 'FlyingWindows VB6' :"

    Close #fHndl
    flgActivateAutoText = False
    
End Sub





' ***************************************************************
' *                                                             *
' *             ADDIN HANDLING STUFF (all private)              *
' *                                                             *
' ***************************************************************


Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                        ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                        ByVal AddInInst As Object, _
                                        custom() As Variant)
    ' Add Flying Windows VB6 to VB IDE
    
    On Error GoTo error_handler
    
    ' Tests
    If ConnectMode = ext_cm_External Then
        MsgBox "FlyingWindows can only run when enabled from VB'S 'Add-Ins' menu", vbExclamation + vbMsgBoxSetForeground, _
                " Error init FlyingWindows VB6"
        
        Exit Sub
    End If
    
    ' Get refs
    Set VBInstance = Application
        
    ' All inits
    Init_Startup


    Exit Sub
    
    
error_handler:
    
    MsgBox Err.Description, vbExclamation + vbMsgBoxSetForeground, " Error on startup AddIn Flying Windows VB6 :"
    
End Sub


Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                            custom() As Variant)
    
    ' Remove AddIn FlyingWindows
    
    
    Dim i As Long

    On Error Resume Next
    
    ' Stop timers loops
    Call m_frmTimer.SetTimerHotcorners(0)
    Call m_frmTimer.SetTimerToolTips(0)

    ' Clean up
    Set m_objPropertiesWindow = Nothing
    Set m_objImmediateWindow = Nothing
    Set m_objProjectExplorerWindow = Nothing
    Set m_objToolWindow = Nothing
    
    ' Write options to registry
    SaveSetting App.EXEName, "Settings", "HotCorners", IIf(flgActivateHotCorners = True, "1", "0")
    SaveSetting App.EXEName, "Settings", "ToolTips", IIf(flgActivateToolTips = True, "1", "0")
    SaveSetting App.EXEName, "Settings", "Position", IIf(flgActivatePosition = True, "1", "0")
    SaveSetting App.EXEName, "Settings", "LargeCombos", IIf(flgActivateLargeCombos = True, "1", "0")
    SaveSetting App.EXEName, "Settings", "HotSides", IIf(flgActivateHotSides = True, "1", "0")
    SaveSetting App.EXEName, "Settings", "AutoText", IIf(flgActivateAutoText = True, "1", "0")
    SaveSetting App.EXEName, "Settings", "MaxComboSize", Format(lMaxComboSize)
    
    ' Remove menubar entry
    mcbMenuCommandBar.Delete
    
    ' Remove option form
    Me.HideFrmOptions
    Unload frmFFoptions
    Set frmFFoptions = Nothing

    ' Not leaving any tooltip forms or so ...
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next i
    
    Set VBInstance = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

    ' Nothing yet

End Sub


Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, _
                                handled As Boolean, _
                                CancelDefault As Boolean)
    
    ' Show FlyingWindows option form
    Me.ShowFrmOptions
    
End Sub


Private Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    
    Dim cbMenu              As Object
    Dim sClipboardText      As String
    
    
    On Error GoTo AddToAddInCommandBarErr
    
    
    ' Look for Add-Ins Menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        
        Exit Function
    End If
    
    ' Add menu entry
    Set AddToAddInCommandBar = cbMenu.Controls.Add(1)
    AddToAddInCommandBar.Caption = sCaption                 ' Set menu text
    With Clipboard
        sClipboardText = .GetText
        .SetData frmFFoptions.picForMenu.Image          ' Set menu picture
        AddToAddInCommandBar.PasteFace
        .Clear
        .SetText sClipboardText
    End With
    
    
    Exit Function
    
    
AddToAddInCommandBarErr:

End Function


' #*#

