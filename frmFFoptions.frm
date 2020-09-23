VERSION 5.00
Begin VB.Form frmFFoptions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   " Config Flying Windows VB6"
   ClientHeight    =   4380
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFFoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picForMenu 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   405
      Picture         =   "frmFFoptions.frx":548A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Top             =   3855
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Activate "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   270
      TabIndex        =   2
      Top             =   728
      Width           =   5715
      Begin VB.CommandButton btnEditAutoText 
         Appearance      =   0  '2D
         BackColor       =   &H00F9C19F&
         Caption         =   "Edit AutoText List"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         Style           =   1  'Grafisch
         TabIndex        =   15
         Top             =   2415
         Width           =   1530
      End
      Begin VB.CheckBox chkAutoText 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Auto Complete Text with F-12 Key"
         Height          =   285
         Left            =   390
         TabIndex        =   14
         Top             =   2415
         Value           =   1  'Aktiviert
         Width           =   3285
      End
      Begin VB.CheckBox chkHotSides 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Hot Screen Sides (Left/Right Screen Border)"
         Height          =   330
         Left            =   390
         TabIndex        =   13
         Top             =   2010
         Value           =   1  'Aktiviert
         Width           =   4260
      End
      Begin VB.TextBox txtComboSize 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5085
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Values in Pixel from 100 to 999"
         Top             =   1605
         Width           =   465
      End
      Begin VB.CheckBox chkIncCombos 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Increase comboboxes     New max Size (100...999)"
         Height          =   345
         Left            =   390
         TabIndex        =   11
         Top             =   1590
         Value           =   1  'Aktiviert
         Width           =   4620
      End
      Begin VB.CheckBox chkPosition 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Show Mouse Position in Pixels"
         Height          =   345
         Left            =   390
         TabIndex        =   5
         Top             =   1200
         Value           =   1  'Aktiviert
         Width           =   3000
      End
      Begin VB.CheckBox chkToolTips 
         BackColor       =   &H00E0E0E0&
         Caption         =   " ToolTips for Controls"
         Height          =   345
         Left            =   390
         TabIndex        =   4
         Top             =   795
         Value           =   1  'Aktiviert
         Width           =   2580
      End
      Begin VB.CheckBox chkHotCorners 
         BackColor       =   &H00E0E0E0&
         Caption         =   " HotCorners for IDE Windows"
         Height          =   345
         Left            =   390
         TabIndex        =   3
         Top             =   390
         Value           =   1  'Aktiviert
         Width           =   3195
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4365
         Picture         =   "frmFFoptions.frx":A914
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.CommandButton btnCancel 
      Appearance      =   0  '2D
      BackColor       =   &H00F9C19F&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4635
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   3825
      Width           =   1335
   End
   Begin VB.CommandButton btnOk 
      Appearance      =   0  '2D
      BackColor       =   &H00F9C19F&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3090
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   3825
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version  2.5.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   9
      Top             =   3885
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Flying Windows VB6  -  (C) by Light Templer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F9C19F&
      Height          =   364
      Index           =   2
      Left            =   403
      TabIndex        =   8
      Top             =   208
      Width           =   4914
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Flying Windows VB6  -  (C) by Light Templer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   364
      Index           =   1
      Left            =   442
      TabIndex        =   7
      Top             =   221
      Width           =   4914
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Flying Windows VB6  -  (C) by Light Templer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   364
      Index           =   0
      Left            =   377
      TabIndex        =   6
      Top             =   208
      Width           =   4914
   End
End
Attribute VB_Name = "frmFFoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmFFoptions.frm
'

Option Explicit


' *************************************
' *           PUBLICS                 *
' *************************************
Public Connect As Connect
'
'
'


' *************************************
' *        PRIVATE SUBS/FUNCS         *
' *************************************
Private Sub btnOk_Click()
    
    With Connect
    
        ' Checkboxes to app vars
        .flgActivateHotCorners = (chkHotCorners.Value = 1)
        .flgActivateToolTips = (chkToolTips.Value = 1)
        .flgActivatePosition = (chkPosition.Value = 1)
        .flgActivateLargeCombos = (chkIncCombos.Value = 1)
        .flgActivateHotSides = (chkHotSides.Value = 1)
        .flgActivateAutoText = (chkAutoText.Value = 1)
        .lMaxComboSize = Val(txtComboSize.Text)

        If .flgActivateAutoText = True Then
            ' Reload the AutoText file
            Screen.MousePointer = vbHourglass
            Connect.LoadAutoText
            Screen.MousePointer = vbDefault
        End If
        
        .HideFrmOptions
    End With
    
End Sub

Private Sub btnCancel_Click()
    
    Connect.HideFrmOptions
    
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Vers. " & App.Major & "." & App.Minor & "." & App.Revision
    With Connect
    
        ' App vars to checkboxes
        chkHotCorners.Value = IIf(.flgActivateHotCorners = True, 1, 0)
        chkToolTips.Value = IIf(.flgActivateToolTips = True, 1, 0)
        chkPosition.Value = IIf(.flgActivatePosition = True, 1, 0)
        chkIncCombos.Value = IIf(.flgActivateLargeCombos = True, 1, 0)
        chkHotSides.Value = IIf(.flgActivateHotSides = True, 1, 0)
        chkAutoText.Value = IIf(.flgActivateAutoText = True, 1, 0)
        txtComboSize.Text = Format(.lMaxComboSize)
    End With

End Sub


Private Sub txtComboSize_Validate(Cancel As Boolean)
    ' Check bounds
    
    Dim lVal    As Long
    
    lVal = Val("0" + txtComboSize.Text)
    If lVal < 100 Or lVal > 999 Then
        Cancel = True
    End If

End Sub

Private Sub btnEditAutoText_Click()
    ' Edit FW_AutoText.Txt with windows notepad. Simple, but efficency ;)
    
    Shell "notepad " + App.Path + "\FW_AutoText.Txt", vbNormalFocus
    
End Sub


' #*#

