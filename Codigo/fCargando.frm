VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form fCargando 
   Caption         =   "Cargando AnimViewer [MMS]"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "fCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub Timer1_Timer()

    bRunning = True
    Timer1.Enabled = False

End Sub
