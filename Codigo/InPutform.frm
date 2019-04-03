VERSION 5.00
Begin VB.Form InPutform 
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "InPutform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum eParse
    ep_Animacion
End Enum
Private PT As Byte


Private Sub Command1_Click(Index As Integer)
If Index = 1 Then
Unload Me
ElseIf Index = 0 Then
    Select Case PT
    
        Case eParse.ep_Animacion
            If Combo1.ListIndex = -1 Then
                MsgBox "Selecciona una animación."
                Exit Sub
            End If
            Call fIndexador.ParseAcFr(fIndexador.qFr, Combo1.ListIndex + 1)
            Select Case fIndexador.qFr
                Case 1
                    fIndexador.Label5.ForeColor = vbWhite
                    fIndexador.Label5.Caption = "Norte: " & fIndexador.GetAcFr(fIndexador.qFr)
                Case 2
                    fIndexador.Label6.ForeColor = vbWhite
                    fIndexador.Label6.Caption = "Este: " & fIndexador.GetAcFr(fIndexador.qFr)
                Case 3
                    fIndexador.Label7.ForeColor = vbWhite
                    fIndexador.Label7.Caption = "Sur: " & fIndexador.GetAcFr(fIndexador.qFr)
                Case 4
                    fIndexador.Label13.ForeColor = vbWhite
                    fIndexador.Label13.Caption = "Oeste: " & fIndexador.GetAcFr(fIndexador.qFr)
            End Select
            
    End Select
    Unload Me
End If
End Sub

Public Sub Parse(ByVal Mensaje As String, ByVal Tipo As Byte)
Dim i As Long
    lbl.Caption = Mensaje
    
    Select Case Tipo
        Case eParse.ep_Animacion
            For i = 1 To Num_NwAnim
                Combo1.AddItem NewAnimationData(i).Desc & " (" & i & ")"
            Next i
    End Select
    InPutform.Show vbModal
End Sub

