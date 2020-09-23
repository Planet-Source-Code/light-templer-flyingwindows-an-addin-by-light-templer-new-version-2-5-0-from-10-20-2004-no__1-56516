VERSION 5.00
Begin VB.Form frmToolTips 
   Appearance      =   0  'Flat
   BackColor       =   &H00FCDCD3&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2040
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3450
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1545
      TabIndex        =   2
      Top             =   240
      Width           =   465
   End
   Begin VB.Label lblKey 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   -15
      TabIndex        =   1
      Top             =   240
      Width           =   345
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   600
   End
End
Attribute VB_Name = "frmToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmToolTips
'

Option Explicit

Public CtrlToDisplay    As VBControl
Public lOwnerWHndl      As Long
Public lNested          As Long
Public RefToVBinst      As VBIDE.VBE


' Sorry, too much formwide vars here, but it was a little too slow ... ;(
Private lPixelXinTwips  As Long
Private lPixelYinTwips  As Long

Private lMaxWidthKeys   As Long
Private lMaxWidthValues As Long
Private lHeightForAll   As Long


' for changing the owner (parent window)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, _
         ByVal nIndex As Long, _
         ByVal wNewLong As Long) As Long

Private Const GWL_HWNDPARENT = (-8)

        
' For moving form without titlebar
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         lParam As Any) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

' misc
Private Declare Function GetParent Lib "user32" _
        (ByVal hWnd As Long) As Long

Private Declare Function IsWindow Lib "user32" _
        (ByVal hWnd As Long) As Long
        
Private Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Long
'
'
'

Public Sub UnloadTT()

    Unload Me

End Sub

Private Sub Form_Load()
        
    Dim i               As Long
    Dim lProps          As Long
    Dim l2PixelXinTwips As Long     ' All for speed ... :)
    Dim sIndex          As String
    Dim lAdjWidth       As Long
    
    On Local Error GoTo error_handler
    
    
    ' Just for sure ...
    If CtrlToDisplay Is Nothing Then
    
        Unload Me
    End If
    
    lPixelXinTwips = Screen.TwipsPerPixelX
    lPixelYinTwips = Screen.TwipsPerPixelY
    l2PixelXinTwips = 2 * lPixelXinTwips
    
    ' Headline
    With lblCaption
        .Caption = " " + CtrlToDisplay.Properties.Item("Name") + " "
        
        ' Ctrl with index?
        sIndex = CtrlToDisplay.Properties.Item("index")
        If sIndex <> "-1" Then
            .Caption = .Caption + "(" + sIndex + ") "
        End If
        
        lHeightForAll = .Height + lPixelYinTwips
        .Height = lHeightForAll
        .Left = -lPixelXinTwips
    End With
        
    ' First Prop
    With lblKey(0)
        .Top = lblCaption.Height - (2 * lPixelYinTwips)
        .Caption = " Type "
        lMaxWidthKeys = .Width
    End With
    With lblValue(0)
        .Top = lblKey(0).Top
        .Caption = " " + CtrlToDisplay.ClassName + " "
        lMaxWidthValues = .Width
    End With
    
           
    If lblValue(0).Caption = " Line " Then   ' Maybe some day I 'll (or you??? ;) ) extend FF to show line props ...
        addProp "X1", "X1"
        addProp "Y1", "Y1"
        addProp "X2", "X2"
        addProp "Y2", "Y2"
    Else
        addProp "Left", "left"
        addProp "Top", "top"
        addProp "Width", "width"
        addProp "Height", "height"
        addProp "Enabled", "enabled"
    End If
        
    ' Probs for most of ctrls
    addProp "Visible", "visible"
    addProp "TabIndex", "tabindex"

    ' "Own" Property
    addProp "Nested", vbNullString, CStr(lNested)
        
    ' Adjust caption width and maybe width of key/value labels
    lAdjWidth = lMaxWidthKeys + lMaxWidthValues - lPixelXinTwips
    If lblCaption.Width < lAdjWidth Then
        lblCaption.Width = lAdjWidth
    Else
        lAdjWidth = (lblCaption.Width - lAdjWidth) / 2
        lMaxWidthKeys = lMaxWidthKeys + lAdjWidth
        lMaxWidthValues = lMaxWidthValues + lAdjWidth
    End If
    
    ' Sizing columns
    lProps = lblKey.uBound
    For i = 0 To lProps
        With lblKey(i)
            .Width = lMaxWidthKeys
            .Height = lHeightForAll
        End With
        
        With lblValue(i)
            .Left = lMaxWidthKeys - l2PixelXinTwips
            .Width = lMaxWidthValues
            .Height = lHeightForAll
        End With
    Next i
    
    
    ' Form adjust
    Me.Width = lblCaption.Width
    Me.Height = lblCaption.Height + ((lProps + 1) * (lHeightForAll - lPixelYinTwips))

    ' CHANGE OWNER. - NOT POSSIBLE IF FLYING WINDOWS RUNS IN IDE MODE (uncompiled to a DLL)
    ' Better comment this 5 lines out to develop or don't click to a form while a tooltip window is open ...
    If IsWindow(lOwnerWHndl) Then
        lOwnerWHndl = GetParent(lOwnerWHndl)
        lOwnerWHndl = GetParent(lOwnerWHndl)
        lOwnerWHndl = SetWindowLong(Me.hWnd, GWL_HWNDPARENT, lOwnerWHndl)
    End If
    

    Exit Sub
    
    
error_handler:

    MsgBox Err.Description, vbExclamation, " Flying Windows:  Error in 'frmToolTips: Form_Load()'"
    
    Unload Me
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' CHANGE OWNER. - NOT POSSIBLE IF FLYING WINDOWS RUNS IN IDE MODE (uncompiled to a DLL)
    ' For developing comment this 3 lines out !
    If IsWindow(lOwnerWHndl) Then
        Call SetWindowLong(Me.hWnd, GWL_HWNDPARENT, lOwnerWHndl)
    End If
    
End Sub


Private Sub addProp(sKey As String, sItemName As String, Optional sValue As String = vbNullString)
    ' Add a key and a value label

    Dim lIndex  As Long
    
    On Error Resume Next
    
    If sValue = vbNullString Then
        With CtrlToDisplay.Properties
            sValue = .Item(sItemName)            ' on error resume ...
            If sValue = vbNullString Then        ' This property does not exist for this ctrl ...
                On Error GoTo 0
                    
                Exit Sub
            End If
        End With
    End If
    
    lIndex = lblKey.Count
    
    Load lblKey(lIndex)
    Load lblValue(lIndex)

    With lblKey(lIndex)
        .Caption = " " + sKey + " "
        .Top = lblKey(lIndex - 1).Top + lHeightForAll - lPixelYinTwips
        .Visible = True
        If .Width > lMaxWidthKeys Then
            lMaxWidthKeys = .Width
        End If
    End With
    
    With lblValue(lIndex)
        .Caption = " " + sValue + " "
        .Top = lblKey(lIndex).Top
        .Visible = True
        If .Width > lMaxWidthValues Then
            lMaxWidthValues = .Width
        End If
    End With
    
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Move a 'keep open' form arround on yellow title
    ' If control key is pressed: Put name into clipboard and close the tooltip form
    
    Dim lPos    As Long
    Dim sName   As String
    
    Const VK_CONTROL = &H11
    
    If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) = True Then
        sName = Trim$(lblCaption.Caption)
        lPos = InStr(1, sName, " ")
        If lPos Then
            sName = Left$(sName, lPos - 1) + Mid$(sName, lPos + 1)
        End If
        Clipboard.Clear
        Clipboard.SetText sName
        Beep
        
        Exit Sub
    End If
    
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
    ' (Re)set VB IDE to "active window"
    RefToVBinst.MainWindow.SetFocus

End Sub

Private Sub lblKey_Click(Index As Integer)
    
    Unload Me
    
End Sub

Private Sub lblValue_Click(Index As Integer)
    
    Const VK_CONTROL = &H11
    
    If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) = True Then
        Clipboard.Clear
        Clipboard.SetText Trim$(lblValue(Index).Caption)
        Beep
        
        Exit Sub
    End If

    Unload Me
    
End Sub


' #*#
