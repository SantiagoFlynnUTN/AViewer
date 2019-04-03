VERSION 5.00
Begin VB.Form frmMapHandler 
   Caption         =   "Map Handler"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   638
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   937
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Indexacion"
      Height          =   8415
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   11895
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3835
         Left            =   120
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   37
         Top             =   4320
         Width           =   3835
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7685
         Left            =   4080
         ScaleHeight     =   512
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   36
         Top             =   480
         Width           =   7685
      End
      Begin VB.ListBox lTempEstatic 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   2480
         Width           =   2055
      End
      Begin VB.ListBox lTempIndex 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   400
         Width           =   2055
      End
      Begin VB.Label lbltiDinamic 
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   35
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Dinamica:"
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblTeReplace 
         Caption         =   "0"
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Reemplazar:"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblTeHeight 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblTeWidth 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblTeTop 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblTeLeft 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblteIndex 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Height:"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Width:"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Top:"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Left:"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Indice:"
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbltiReplace 
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbltiEstatic 
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbltiGrafico 
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lbltiTemp 
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbltiIndex 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Reemplazar:"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Estatica:"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Grafico:"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Estatic Temporal:"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Indice:"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblnumEstatic 
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Numero de estatic temporales:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label lblnumindex 
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de indices temporales:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   975
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   11775
      Begin VB.Label lblPath 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label1 
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdAnalizar 
      Caption         =   "Analizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   4380
      Left            =   120
      Pattern         =   "*.MapTemp*"
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMapHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UltimoSelecto As Integer

Private Sub CmdAnalizar_Click()

 
modMapHandler.Analizar File1.Path & "\" & File1.List(File1.ListIndex), lTempIndex, lTempEstatic, lblnumindex, lblnumEstatic

CmdAnalizar.Enabled = False

End Sub

Private Sub File1_Click()
    If File1.ListIndex > -1 Then
        If File1.ListIndex <> UltimoSelecto Then
            CmdAnalizar.Enabled = True
            UltimoSelecto = File1.ListIndex
            lblPath.Caption = File1.Path & "\" & File1.List(File1.ListIndex)
        End If
    End If
        
End Sub

Private Sub Form_Load()
UltimoSelecto = -1
File1.Path = App.Path & "\TempMapas\"
End Sub

Private Sub lblTeHeight_Click()
Dim Valor As Integer
    '>>>> Modifica Height.
    If lTempEstatic.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el nuevo Height para el estatic.", "Estatic - Height", modMapHandler.Estatic_rHeight(lTempEstatic.ListIndex + 1)))
        modMapHandler.Estatic_wHeight lTempEstatic.ListIndex + 1, Valor
        lblTeHeight.Caption = Valor
    End If
End Sub

Private Sub lblTeLeft_Click()
Dim Valor As Integer
    '>>>> Modifica left.
    If lTempEstatic.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el nuevo LEFT para el estatic.", "Estatic - Left", modMapHandler.Estatic_rLeft(lTempEstatic.ListIndex + 1)))
        modMapHandler.Estatic_wLeft lTempEstatic.ListIndex + 1, Valor
        lblTeLeft.Caption = Valor
    End If
    
    
End Sub

Private Sub lblTeTop_Click()
Dim Valor As Integer
    '>>>> Modifica top.
    If lTempEstatic.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el nuevo TOP para el estatic.", "Estatic - Top", modMapHandler.Estatic_rTop(lTempEstatic.ListIndex + 1)))
        modMapHandler.Estatic_wTop lTempEstatic.ListIndex + 1, Valor
        lblTeTop.Caption = Valor
    End If
End Sub

Private Sub lblTeWidth_Click()
Dim Valor As Integer
    '>>>> Modifica w.
    If lTempEstatic.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el nuevo WIDTH para el estatic.", "Estatic - Width", modMapHandler.Estatic_rWidth(lTempEstatic.ListIndex + 1)))
        modMapHandler.Estatic_wWidth lTempEstatic.ListIndex + 1, Valor
        lblTeWidth.Caption = Valor
    End If
End Sub

Private Sub lbltiDinamic_Click()
Dim Valor As Integer
    If lTempIndex.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el numero de Dinamic", "Dinamic", modMapHandler.Index_rDinamic(lTempIndex.ListIndex + 1)))
        lbltiDinamic.Caption = Valor
    End If
End Sub

Private Sub lbltiEstatic_Click()
Dim Valor As Integer
    If lTempIndex.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el numero de estatic", "Estatic", modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1)))
        modMapHandler.Index_wEstatic lTempIndex.ListIndex + 1, Valor
        lbltiEstatic.Caption = Valor
    End If
Dim S As String
    If modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1) > 0 Then
        If modMapHandler.Index_rTemp(lTempIndex.ListIndex + 1) = 1 Then
            S = " Left: " & modMapHandler.Estatic_rLeft(0, modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1)) & vbCrLf & _
                " Top: " & modMapHandler.Estatic_rTop(0, modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1)) & vbCrLf & _
                " Width: " & modMapHandler.Estatic_rWidth(0, modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1)) & vbCrLf & _
                " Height: " & modMapHandler.Estatic_rHeight(0, modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1))
        
        Else
            With EstaticData(modMapHandler.Index_rEstatic(lTempIndex.ListIndex + 1))
            S = " Left: " & .L & vbCrLf & _
            " Top: " & .T & vbCrLf & _
            " Width: " & .W & vbCrLf & vbCrLf & _
            " Height: " & .H
            
            
            End With
        
        End If
                    lbltiEstatic.ToolTipText = S

    End If
    
End Sub

Private Sub lbltiGrafico_Click()
Dim Valor As Integer
    If lTempIndex.ListIndex > -1 Then
        Valor = Val(InputBox("Escribe el numero de Grafico", "Estatic", modMapHandler.Index_rGrafico(lTempIndex.ListIndex + 1)))
        modMapHandler.Index_wGrafico lTempIndex.ListIndex + 1, Valor
        lbltiGrafico.Caption = Valor
    End If
End Sub

Private Sub lbltiTemp_Click()
Dim Valor As Integer
    If lTempIndex.ListIndex > -1 Then
        Valor = Val(InputBox("1=Estatic Temporal / 0= Estatic No temporal", "Temp Estatic", modMapHandler.Index_rTemp(lTempIndex.ListIndex + 1)))
        lbltiTemp.Caption = Valor
    End If
    
End Sub

Private Sub lTempEstatic_Click()
    'Muestra la info del estatic.
    If lTempEstatic.ListIndex > -1 Then
        modMapHandler.VerEstatic_Info lTempEstatic.ListIndex + 1, lblteIndex, lblTeLeft, lblTeTop, lblTeWidth, lblTeHeight, lblTeReplace
        
    End If
End Sub

Private Sub lTempIndex_Click()
    If lTempIndex.ListIndex > -1 Then
        modMapHandler.VerIndex_Info lTempIndex.ListIndex + 1, lbltiIndex, lbltiDinamic, lbltiEstatic, lbltiTemp, lbltiGrafico, lbltiReplace
        
        Picture1.Cls
        Picture2.Cls
        modMapHandler.VerIndex_Grafico lTempIndex.ListIndex + 1, Picture1
        modMapHandler.VerIndex_Preview lTempIndex.ListIndex + 1, Picture2
        
    End If
End Sub
