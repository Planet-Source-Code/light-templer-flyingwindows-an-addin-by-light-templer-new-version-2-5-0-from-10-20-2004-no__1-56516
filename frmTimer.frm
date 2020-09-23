VERSION 5.00
Begin VB.Form frmTimer 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   0  'Kein
   Caption         =   "frmTimer"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer TimerAutoText 
      Left            =   1815
      Top             =   315
   End
   Begin VB.Timer TimerHotcorners 
      Left            =   270
      Top             =   315
   End
   Begin VB.Timer TimerToolTips 
      Left            =   1042
      Top             =   315
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmTimer.frm
'

Option Explicit

Public Event TimerHotcornersElapsed()
Public Event TimerToolTipsElapsed()
Public Event TimerAutoText()
'
'
'

Friend Sub SetTimerHotcorners(ByVal iMilliSeconds As Integer)

   TimerHotcorners.Interval = iMilliSeconds
   
End Sub

Friend Sub SetTimerToolTips(ByVal iMilliSeconds As Integer)

   TimerToolTips.Interval = iMilliSeconds
   
End Sub

Friend Sub SetTimerAutoText(ByVal iMilliSeconds As Integer)

   TimerAutoText.Interval = iMilliSeconds
   
End Sub


Private Sub TimerHotcorners_Timer()

   RaiseEvent TimerHotcornersElapsed    ' Catched in 'Connect'

End Sub

Private Sub TimerToolTips_Timer()

    RaiseEvent TimerToolTipsElapsed     ' Catched in 'Connect'

End Sub

Private Sub TimerAutoText_Timer()

    RaiseEvent TimerAutoText            ' Catched in 'Connect'

End Sub


' #*#

