VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2640
      Top             =   2640
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   375
      LargeChange     =   2000
      Left            =   6600
      Max             =   10000
      Min             =   -10000
      SmallChange     =   500
      TabIndex        =   4
      Top             =   1320
      Width           =   5415
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      LargeChange     =   2000
      Left            =   960
      Max             =   10000
      Min             =   -10000
      SmallChange     =   500
      TabIndex        =   3
      Top             =   1320
      Width           =   5415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   2000
      Left            =   960
      Max             =   30000
      Min             =   1000
      SmallChange     =   80
      TabIndex        =   1
      Top             =   720
      Value           =   11025
      Width           =   5415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      LargeChange     =   2000
      Left            =   6600
      Max             =   30000
      Min             =   1000
      SmallChange     =   80
      TabIndex        =   2
      Top             =   720
      Value           =   11025
      Width           =   5415
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3255
      LargeChange     =   1200
      Left            =   360
      Max             =   -10000
      SmallChange     =   200
      TabIndex        =   0
      Top             =   600
      Value           =   -2800
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "SetPan(1)  ( sets the tone to the left or right channel )"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "SetPan(0)  ( sets the tone to the left or right channel )"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Right Frequency"
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Left Frequency"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Volume"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    calcWave.make330Hz
    Init_DX7 (Form1.Hwnd)

    HScroll1_Scroll    'Set directSound Frequency
    HScroll2_Scroll    'Set directSound Frequency
    VScroll1_Change    'Set directSound Volume
        
    DSBWRITE 0, myByte(): DSBWRITE 1, myByte() 'Play it
End Sub
Private Sub HScroll1_Change()
    DSB(0).SetFrequency HScroll1.Value: Label2 = Fix(HScroll1.Value / 11025 * O2F) & " Hz"
End Sub
Private Sub HScroll2_Change()
    DSB(1).SetFrequency HScroll2.Value: Label3 = Fix(HScroll2.Value / 11025 * O2F) & " Hz"
End Sub
Private Sub HScroll3_Change()
    DSB(0).SetPan HScroll3.Value: Label4 = HScroll3.Value
End Sub
Private Sub HScroll4_Scroll()
    DSB(1).SetPan HScroll4.Value: Label5 = HScroll4.Value
End Sub
Private Sub Timer1_Timer()
    gfw = GetForegroundWindow
    If gfw > 0 Then DS.SetCooperativeLevel gfw, DSSCL_NORMAL
End Sub

Private Sub VScroll1_Scroll(): VScroll1_Change: End Sub
Private Sub HScroll1_Scroll(): HScroll1_Change: End Sub
Private Sub HScroll2_Scroll(): HScroll2_Change: End Sub
Private Sub HScroll3_Scroll(): HScroll3_Change: End Sub
Private Sub Form_Unload(Cancel As Integer): Term_DX7: End Sub
Private Sub VScroll1_Change()
    DSB(0).SetVolume VScroll1.Value: DSB(1).SetVolume VScroll1.Value
    Label1.Caption = "Volume" & vbCrLf & VScroll1.Value
End Sub
