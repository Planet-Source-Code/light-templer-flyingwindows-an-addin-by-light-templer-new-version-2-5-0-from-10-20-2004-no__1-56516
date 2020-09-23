VERSION 5.00
Begin VB.Form frmAddIn 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   " Config Flying Windows VB6"
   ClientHeight    =   3510
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
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
      Height          =   2115
      Left            =   270
      TabIndex        =   2
      Top             =   750
      Width           =   4905
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Show Mouse Position in Pixels"
         Height          =   345
         Left            =   390
         TabIndex        =   5
         Top             =   1410
         Value           =   1  'Aktiviert
         Width           =   3195
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Control ToolTips"
         Height          =   345
         Left            =   390
         TabIndex        =   4
         Top             =   900
         Value           =   1  'Aktiviert
         Width           =   2580
      End
      Begin VB.CheckBox Check1 
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
         Left            =   3810
         Picture         =   "frmAddIn.frx":548A
         Top             =   405
         Width           =   720
      End
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  '2D
      BackColor       =   &H00F9C19F&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3945
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
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
      Left            =   2445
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
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
      Height          =   360
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   210
      Width           =   4920
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
      Height          =   360
      Index           =   1
      Left            =   270
      TabIndex        =   7
      Top             =   225
      Width           =   4920
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
      Height          =   360
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   4920
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub OKButton_Click()
    MsgBox "AddIn operation on: " & VBInstance.FullName
End Sub
