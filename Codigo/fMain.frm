VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Main"
   ClientHeight    =   8610
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   574
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   923
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Animar"
      Height          =   195
      Left            =   11640
      TabIndex        =   74
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Height          =   495
      Left            =   13080
      TabIndex        =   73
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Height          =   495
      Left            =   11640
      TabIndex        =   72
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Height          =   855
      Left            =   12480
      TabIndex        =   71
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Height          =   855
      Left            =   12480
      TabIndex        =   70
      Top             =   5880
      Width           =   495
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   7
      ItemData        =   "fMain.frx":0000
      Left            =   120
      List            =   "fMain.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin AnimViewer.lvButtons_H czoom 
      Height          =   375
      Left            =   3000
      TabIndex        =   46
      Top             =   7800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "Zoom"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   6
      ItemData        =   "fMain.frx":0011
      Left            =   120
      List            =   "fMain.frx":0018
      Style           =   2  'Dropdown List
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton CmdSAve 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Indices"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Viewer"
      Height          =   7950
      Left            =   3720
      TabIndex        =   1
      Top             =   675
      Width           =   7800
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Animaciones"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   2
      Left            =   2700
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Modeling"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   3
      Left            =   4230
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Particulas"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   4
      Left            =   5760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Fx"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   5
      Left            =   7305
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Meditacion"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   6
      Left            =   8835
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Cabezas"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin AnimViewer.lvButtons_H cOpc 
      Height          =   615
      Index           =   7
      Left            =   10320
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "Cascos"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   2
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame fCabezas 
      Caption         =   "Cabezas"
      Height          =   3255
      Left            =   120
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1200
         TabIndex        =   49
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1200
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   960
         TabIndex        =   39
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   960
         TabIndex        =   35
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Offset Lat:"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Offset Y:"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Offset Ojos:"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame fModeling 
      Caption         =   "Modeling"
      Height          =   3015
      Left            =   120
      TabIndex        =   51
      Top             =   4800
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command6 
         Caption         =   "Ver"
         Height          =   255
         Left            =   3000
         TabIndex        =   66
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ver"
         Height          =   255
         Left            =   3000
         TabIndex        =   65
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ver"
         Height          =   255
         Left            =   3000
         TabIndex        =   64
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ver"
         Height          =   255
         Left            =   3000
         TabIndex        =   63
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ver"
         Height          =   255
         Left            =   3000
         TabIndex        =   62
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox cArma 
         Height          =   315
         Left            =   840
         TabIndex        =   61
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cEscudo 
         Height          =   315
         Left            =   840
         TabIndex        =   60
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cCasco 
         Height          =   315
         Left            =   840
         TabIndex        =   59
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cCabeza 
         Height          =   315
         Left            =   840
         TabIndex        =   58
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cBody 
         Height          =   315
         Left            =   840
         TabIndex        =   57
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Casco:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Arma:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Escudo:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Cabeza:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Cuerpo:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frAnim 
      Caption         =   "Animaciones"
      Height          =   5535
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1320
         TabIndex        =   68
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Nueva"
         Height          =   315
         Left            =   2520
         TabIndex        =   67
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Inicial:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Grafico:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "W-H:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Col-Fil:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Num Frames:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Velocidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Velocidad de mov:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label lblFrame 
      Height          =   375
      Left            =   11640
      TabIndex        =   75
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblfr 
      Caption         =   "FR:"
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label lblVel 
      Caption         =   "VELOCIDAD:"
      Height          =   255
      Left            =   960
      TabIndex        =   30
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label lblfps 
      Caption         =   "FPS:"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Menu mFile 
      Caption         =   "Archivo"
      Begin VB.Menu m_a_freeGraficos 
         Caption         =   "Liberar Graficos"
      End
      Begin VB.Menu m_a_Exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mGrafica 
      Caption         =   "Grafica"
      Begin VB.Menu m_g_NeglectNegro 
         Caption         =   "Neglect Negro"
         Checked         =   -1  'True
      End
      Begin VB.Menu m_g_ColorFondo 
         Caption         =   "Color de Fondo"
         Begin VB.Menu m_CF_Negro 
            Caption         =   "Negro"
         End
         Begin VB.Menu m_CF_Azul 
            Caption         =   "Azul"
         End
         Begin VB.Menu m_CF_Rojo 
            Caption         =   "Rojo"
         End
         Begin VB.Menu m_CF_Blanco 
            Caption         =   "Blanco"
         End
      End
   End
   Begin VB.Menu mAnim 
      Caption         =   "Animacion"
      Begin VB.Menu m_a_Intervalo 
         Caption         =   "Intervalo"
      End
      Begin VB.Menu m_a_Vel 
         Caption         =   "Velocidad"
      End
      Begin VB.Menu m_a_verframe 
         Caption         =   "Ver Frame"
      End
      Begin VB.Menu m_a_Pausa 
         Caption         =   "Pausa"
      End
      Begin VB.Menu m_a_TestBody 
         Caption         =   "Cuerpo de prueba"
         Begin VB.Menu m_c_elegir 
            Caption         =   "Cuerpo:"
         End
         Begin VB.Menu m_c_activado 
            Caption         =   "Activado"
         End
      End
      Begin VB.Menu m_z_TestHead 
         Caption         =   "Cabeza de Prueba"
         Begin VB.Menu m_z_elegir 
            Caption         =   "Cabeza:"
         End
         Begin VB.Menu m_z_Activado 
            Caption         =   "Activado"
         End
      End
   End
   Begin VB.Menu m_i_Open 
      Caption         =   "Indexador"
   End
   Begin VB.Menu m_mapHandler 
      Caption         =   "MapHandler"
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cArma_Click()
    sWeapon = cArma.ListIndex
    If sWeapon > 0 Then MostrarVer 4
    
End Sub

Private Sub cBody_Click()
    sBody = fMain.cBody.ListIndex
    If sBody > 0 Then MostrarVer 0
End Sub

Private Sub cCabeza_Click()
    sHead = cCabeza.ListIndex
    If sHead > 0 Then MostrarVer 1
End Sub

Private Sub cCasco_Click()
       sHelmet = cCasco.ListIndex
    If sHelmet > 0 Then MostrarVer 2
    
End Sub

Private Sub cEscudo_Click()
    sShield = cEscudo.ListIndex
    If sShield > 0 Then MostrarVer 3
End Sub

Private Sub CmdSAve_Click()
    Select Case UltimaOpcion
    
        Case eOpciones.Animaciones_op
            GuardarAnim IxS
    
    End Select
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then NHeadData(sHead).Raza = Combo1.ListIndex

End Sub

Private Sub Combo2_Click()
    
    NHeadData(sHead).Genero = Combo2.ListIndex + 1

End Sub

Private Sub Command1_Click()
If Modeling_Type = 2 Then
    WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & sHead, "OFFSET_DIBUJO", CStr(NHeadData(sHead).OffsetDibujoY)
    WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & sHead, "OFFSET_OJOS", CStr(NHeadData(sHead).OffsetOjos)
    WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & sHead, "RAZA", CStr(NHeadData(sHead).Raza)
    WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & sHead, "GENERO", CStr(NHeadData(sHead).Genero)
    WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & sHead, "DESC", CStr(NHeadData(sHead).Desc)
    
    fMain.cCabeza.List(fMain.cCabeza.ListIndex) = "[" & GetRaza(NHeadData(sHead).Raza) & IIf(NHeadData(sHead).Raza > 0, GetLastLetter(NHeadData(sHead).Genero), vbNullString) & "] - " & NHeadData(sHead).Desc & "(" & sHead & ")"
ElseIf Modeling_Type = 3 Then
    WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "HELMET" & IxS, "OFFSET_DIBUJO", CStr(NHelmetData(IxS).OffsetDibujoY)
    WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "HELMET" & IxS, "DESC", CStr(NHelmetData(IxS).Desc)
    WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "HELMET" & IxS, "ALPHA", CStr(NHelmetData(IxS).Alpha)
    WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "HELMET" & IxS, "OFFSET_LAT", CStr(NHelmetData(IxS).OffsetLat)
        
    fMain.cCasco.List(fMain.cCasco.ListIndex) = NHelmetData(sHelmet).Desc & "(" & sHelmet & ")"
ElseIf Modeling_Type = 4 Then
    WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(sShield), "OverWriteGrafico", CStr(nShieldDATA(sShield).OverWriteGrafico)
    WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(sShield), "DESC", CStr(nShieldDATA(sShield).Desc)
    WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(sShield), "ALPHA", CStr(nShieldDATA(sShield).Alpha)

    fMain.cEscudo.List(fMain.cEscudo.ListIndex) = nShieldDATA(sShield).Desc & "(" & sShield & ")"
    
ElseIf Modeling_Type = 5 Then
    WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(sWeapon), "OverWriteGrafico", CStr(nWeaponData(sWeapon).OverWriteGrafico)
    WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(sWeapon), "DESC", CStr(nWeaponData(sWeapon).Desc)
    WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(sWeapon), "ALPHA", CStr(nWeaponData(sWeapon).Alpha)
    fMain.cArma.List(fMain.cArma.ListIndex) = nWeaponData(sWeapon).Desc & "(" & sWeapon & ")"
ElseIf Modeling_Type = 1 Then
    fMain.cBody.List(fMain.cBody.ListIndex) = nBodyData(sBody).Desc & "(" & sBody & ")"
    WriteVar App.Path & "\RES\INDEX\NewBody.dat", CStr(sBody), "DESC", nBodyData(sBody).Desc
    WriteVar App.Path & "\RES\INDEX\NewBody.dat", CStr(sBody), "OverWriteGrafico", CStr(nBodyData(sBody).OverWriteGrafico)
    WriteVar App.Path & "\RES\INDEX\NewBody.dat", CStr(sBody), "OffsetY", CStr(nBodyData(sBody).OffsetY)
End If

End Sub

Private Sub Command10_Click()

If UltimaOpcion = eOpciones.Animaciones_op Then
    animCounter = animCounter - 1
    Exit Sub
End If

If acHeading = E_Heading.WEST Then AcFrm = AcFrm + 1

acHeading = E_Heading.WEST
End Sub

Private Sub Command11_Click()
If UltimaOpcion = eOpciones.Animaciones_op Then
    animCounter = animCounter + 1
    Exit Sub
End If

If acHeading = E_Heading.EAST Then AcFrm = AcFrm + 1

acHeading = E_Heading.EAST
End Sub

Private Sub Command2_Click()
    Label12.Visible = True
    Text9.Visible = True
    Text8.Visible = True
    Label8.Visible = True
    Label8.Caption = "Offset Y"
    Label13.Caption = "Grafico:"
    Label13.Visible = True
    Text11.Visible = True
    Combo1.Visible = False
    Combo2.Visible = False
    fCabezas.Visible = True
    Text10.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Modeling_Type = 1
    Text8.Text = nBodyData(sBody).OffsetY
    Text11.Text = nBodyData(sBody).OverWriteGrafico
    Text9.Text = nBodyData(sBody).Desc
    
End Sub

Private Sub Command3_Click()
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Escudos"
            fMain.Text8.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = False
            fMain.Combo2.Visible = False
            fMain.Label9.Caption = "Transparencia:"
            fMain.Text11.Visible = True
            fMain.Text8.Visible = False
            Label8.Visible = False
            Label10.Visible = False
            Label11.Visible = False
                        Label9.Visible = True
            Text10.Visible = True
            fMain.Label13.Visible = True
            Modeling_Type = 5
            Label13.Caption = "Grafico:"
            Text10.Text = nWeaponData(sWeapon).Alpha
            Text9.Text = nWeaponData(sWeapon).Desc
            Text11.Text = nWeaponData(sWeapon).OverWriteGrafico
End Sub

Private Sub Command4_Click()
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Escudos"
            fMain.Text8.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = False
            fMain.Combo2.Visible = False
            fMain.Label9.Caption = "Transparencia:"
            fMain.Text11.Visible = True
            fMain.Label13.Visible = True
            fMain.Text8.Visible = False
            Label8.Visible = False
            Label10.Visible = False
            Label11.Visible = False
            Label9.Visible = True
            Text10.Visible = True
            Modeling_Type = 4
            Label13.Caption = "Grafico:"
            Text10.Text = nShieldDATA(sShield).Alpha
            Text9.Text = nShieldDATA(sShield).Desc
            Text11.Text = nShieldDATA(sShield).OverWriteGrafico
            
            
End Sub

Private Sub Command5_Click()
            Label13.Caption = "Offset Lat:"
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Cascos"
            fMain.Text8.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = False
            fMain.Combo2.Visible = False
            fMain.Label9.Caption = "Transparencia:"
            fMain.Text11.Visible = True
            fMain.Label13.Visible = True
            Modeling_Type = 3
            fMain.Text8.Visible = True
            Label9.Visible = True
            Text10.Visible = True
            
            Label10.Visible = False
            Label11.Visible = False
            Label8.Visible = True
End Sub
Private Sub MostrarVer(ByVal quemuestro As Byte)

    Select Case quemuestro
        Case 0
            Label12.Visible = True
    Text9.Visible = True
    Text8.Visible = True
    Label8.Visible = True
    Label8.Caption = "Offset Y"
    Label13.Caption = "Grafico:"
    Label13.Visible = True
    Text11.Visible = True
    Combo1.Visible = False
    Combo2.Visible = False
    fCabezas.Visible = True
    Text10.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Modeling_Type = 1
    Text8.Text = nBodyData(sBody).OffsetY
    Text11.Text = nBodyData(sBody).OverWriteGrafico
    Text9.Text = nBodyData(sBody).Desc
    Case 1
        Modeling_Type = 2
    Label13.Caption = "Offset Lat:"
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Cabezas"
            fMain.Combo1.AddItem "Fantasma"
            fMain.Combo1.AddItem "Humano"
            fMain.Combo1.AddItem "Elfo"
            fMain.Combo1.AddItem "Elfo Oscuro"
            fMain.Combo1.AddItem "Gnomo"
            fMain.Combo1.AddItem "Enano"
            fMain.Combo2.AddItem "Hombre"
            fMain.Combo2.AddItem "Mujer"
            fMain.Combo1.ListIndex = -1
            fMain.Combo2.ListIndex = -1
            fMain.Text8.Text = vbNullString
            fMain.Text10.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = True
            fMain.Combo2.Visible = True
            fMain.Label9.Caption = "Offset Ojos:"
            Label9.Visible = True
            Text10.Visible = True
            Text10.Text = NHeadData(sHead).OffsetOjos
            
            fMain.Text11.Visible = False
            fMain.Label13.Visible = False
                    fMain.Text8.Visible = True
                    fMain.Text8.Text = NHeadData(sHead).OffsetDibujoY
                    fMain.Text9.Text = NHeadData(sHead).Desc
                    fMain.Combo1.ListIndex = NHeadData(sHead).Raza
                    fMain.Combo2.ListIndex = NHeadData(sHead).Genero - 1
                    
Label8.Visible = True
            Label10.Visible = True
            Label11.Visible = True

Case 2
            Label13.Caption = "Offset Lat:"
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Cascos"
            fMain.Text8.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = False
            fMain.Combo2.Visible = False
            fMain.Label9.Caption = "Transparencia:"
            fMain.Text11.Visible = True
            fMain.Label13.Visible = True
            Modeling_Type = 3
            fMain.Text8.Visible = True
            Label9.Visible = True
            Text10.Visible = True
            
            Label10.Visible = False
            Label11.Visible = False
            Label8.Visible = True
Case 3
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Escudos"
            fMain.Text8.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = False
            fMain.Combo2.Visible = False
            fMain.Label9.Caption = "Transparencia:"
            fMain.Text11.Visible = True
            fMain.Label13.Visible = True
            fMain.Text8.Visible = False
            Label8.Visible = False
            Label10.Visible = False
            Label11.Visible = False
            Label9.Visible = True
            Text10.Visible = True
            Modeling_Type = 4
            Label13.Caption = "Grafico:"
            Text10.Text = nShieldDATA(sShield).Alpha
            Text9.Text = nShieldDATA(sShield).Desc
            Text11.Text = nShieldDATA(sShield).OverWriteGrafico
            
        Case 4
                    fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Escudos"
            fMain.Text8.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = False
            fMain.Combo2.Visible = False
            fMain.Label9.Caption = "Transparencia:"
            fMain.Text11.Visible = True
            fMain.Text8.Visible = False
            Label8.Visible = False
            Label10.Visible = False
            Label11.Visible = False
                        Label9.Visible = True
            Text10.Visible = True
            fMain.Label13.Visible = True
            Modeling_Type = 5
            Label13.Caption = "Grafico:"
            Text10.Text = nWeaponData(sWeapon).Alpha
            Text9.Text = nWeaponData(sWeapon).Desc
            Text11.Text = nWeaponData(sWeapon).OverWriteGrafico
        
    End Select
    
End Sub
Private Sub Command6_Click()
    Modeling_Type = 2
    Label13.Caption = "Offset Lat:"
            fMain.fCabezas.Visible = True
            fMain.fCabezas.Caption = "Cabezas"
            fMain.Combo1.AddItem "Fantasma"
            fMain.Combo1.AddItem "Humano"
            fMain.Combo1.AddItem "Elfo"
            fMain.Combo1.AddItem "Elfo Oscuro"
            fMain.Combo1.AddItem "Gnomo"
            fMain.Combo1.AddItem "Enano"
            fMain.Combo2.AddItem "Hombre"
            fMain.Combo2.AddItem "Mujer"
            fMain.Combo1.ListIndex = -1
            fMain.Combo2.ListIndex = -1
            fMain.Text8.Text = vbNullString
            fMain.Text10.Text = vbNullString
            fMain.Text9.Text = vbNullString
            fMain.Combo1.Visible = True
            fMain.Combo2.Visible = True
            fMain.Label9.Caption = "Offset Ojos:"
            Label9.Visible = True
            Text10.Visible = True
            Text10.Text = NHeadData(sHead).OffsetOjos
            
            fMain.Text11.Visible = False
            fMain.Label13.Visible = False
                    fMain.Text8.Visible = True
                    fMain.Text8.Text = NHeadData(sHead).OffsetDibujoY
                    fMain.Text9.Text = NHeadData(sHead).Desc
                    fMain.Combo1.ListIndex = NHeadData(sHead).Raza
                    fMain.Combo2.ListIndex = NHeadData(sHead).Genero - 1
                    
Label8.Visible = True
            Label10.Visible = True
            Label11.Visible = True
End Sub

Private Sub Command7_Click()
    Num_NwAnim = Num_NwAnim + 1
    ReDim Preserve NewAnimationData(1 To Num_NwAnim)
    Text1.Text = vbNullString
    Text2.Text = vbNullString
    Text3.Text = vbNullString
    Text4.Text = vbNullString
    Text5.Text = vbNullString
    Text6.Text = vbNullString
    Text7.Text = vbNullString
    Text8.Text = vbNullString
    IxS = Num_NwAnim
    fMain.Lista(UltimaOpcion).AddItem "(" & IxS & ")"
    fMain.Lista(UltimaOpcion).ListIndex = fMain.Lista(UltimaOpcion).ListCount - 1
    
End Sub

Private Sub Command8_Click()
If acHeading = E_Heading.SOUTH Then AcFrm = AcFrm + 1

acHeading = E_Heading.SOUTH
End Sub

Private Sub Command9_Click()
If acHeading = E_Heading.NORTH Then AcFrm = AcFrm + 1
                                                    
acHeading = E_Heading.NORTH
End Sub

Private Sub cOpc_Click(Index As Integer)
    If Index <> UltimaOpcion Then
        If UltimaOpcion <> -1 Then
            CerrarOpcion UltimaOpcion
        End If
        AbrirOpcion Index
        UltimaOpcion = Index
        Else
        CerrarOpcion UltimaOpcion
        UltimaOpcion = -1
    End If
End Sub

Private Sub czoom_Click()
    Zooming = czoom.value
    If Zooming Then
        SetZoom 0
    End If
    Debug.Print Zooming
    
End Sub

Private Sub Form_Click()
    Me.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        VelMod = VelMod + 0.01
        fMain.lblVel.Caption = "VELOCIDAD: " & VelMod
    ElseIf KeyCode = 2 Then
        VelMod = VelMod - 0.01
        fMain.lblVel.Caption = "VELOCIDAD: " & VelMod
    End If
    If KeyCode = vbKeyRight Then
    acHeading = E_Heading.EAST
ElseIf KeyCode = vbKeyLeft Then
    acHeading = E_Heading.WEST
ElseIf KeyCode = vbKeyUp Then
    acHeading = E_Heading.NORTH
ElseIf KeyCode = vbKeyDown Then
    acHeading = E_Heading.SOUTH
End If
End Sub

Private Sub Form_Load()
Dim p As Long
    UltimaOpcion = -1
    For p = 0 To cOpc.UBound
        cOpc(p).value = False
    Next p
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
bRunning = False
End Sub

Private Sub Frame1_Click()
    Me.SetFocus

End Sub
Public Sub SetZoom(ByVal BTN As Integer)
Dim aW As Integer
Dim aH As Integer
Dim zW As Integer
Dim zH As Integer
Dim ExV As Integer

ExV = ZoomV
If BTN = vbRightButton Then
    ZoomV = ZoomV - 1
    If ZoomV <= 1 Then
        ZoomV = 1
        ZoomBufferRect.left = MainBufferRect.left
        ZoomBufferRect.top = MainBufferRect.top
        ZoomBufferRect.Bottom = MainBufferRect.Bottom
        ZoomBufferRect.Right = MainBufferRect.Right
        ZoomX = 0
        ZoomY = 0
        
        Exit Sub
    End If
ElseIf BTN = vbLeftButton Then
    ZoomV = ZoomV + 1
ElseIf BTN = 0 Then
    If ZoomV = 0 Then
        ZoomV = 1
        ZoomBufferRect.left = MainBufferRect.left
        ZoomBufferRect.top = MainBufferRect.top
        ZoomBufferRect.Bottom = MainBufferRect.Bottom
        ZoomBufferRect.Right = MainBufferRect.Right
        ZoomX = 0
        ZoomY = 0
    End If
Exit Sub

End If



aW = MainBufferRect.Right - MainBufferRect.left
aH = MainBufferRect.Bottom - MainBufferRect.top

zW = aW / ZoomV
zH = aH / ZoomV
ZoomX = (ZoomX + ((fX / Screen.TwipsPerPixelX) / ExV)) - (zW * 0.5)
ZoomY = (ZoomY + ((fY / Screen.TwipsPerPixelY) / ExV)) - (zH * 0.5)


ZoomBufferRect.left = ZoomX
ZoomBufferRect.Right = ZoomX + zW

ZoomBufferRect.top = ZoomY
ZoomBufferRect.Bottom = ZoomY + zH



End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fX = X
    fY = Y
    
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Zooming Then
        SetZoom Button
    End If
End Sub

Private Sub Lista_Click(Index As Integer)
    If LenB(ReadField(2, Lista(Index).List(Lista(Index).ListIndex), Asc("("))) > 0 Then
        IxS = Val(ReadField(2, Lista(Index).List(Lista(Index).ListIndex), Asc("(")))
    End If
    Select Case UltimaOpcion
    
        Case eOpciones.Animaciones_op
        If IxS > 0 Then
            Text2.Text = Replace(CStr(NewAnimationData(IxS).Velocidad), ",", ".")
            Text3.Text = NewAnimationData(IxS).NumFrames
            Text4.Text = NewAnimationData(IxS).Columnas & "-" & NewAnimationData(IxS).Filas
            Text5.Text = NewAnimationData(IxS).Width & "-" & NewAnimationData(IxS).Height
            Text6.Text = NewAnimationData(IxS).Grafico
            Text7.Text = NewAnimationData(IxS).Initial
            Text12.Text = NewAnimationData(IxS).Desc
        End If
        

    End Select
    
End Sub

Private Sub m_a_Exit_Click()
    bRunning = False
End Sub

Private Sub m_a_freeGraficos_Click()
    SurfaceDB8.Release
End Sub

Private Sub m_a_Intervalo_Click()
    anim_Intervalo = Val(InputBox("Escribe el intervalo en segundos."))
    
End Sub

Private Sub m_a_Pausa_Click()
    m_a_Pausa.Checked = Not m_a_Pausa.Checked
    Pausa = m_a_Pausa.Checked
End Sub

Private Sub m_a_Vel_Click()
    VelMod = Val(InputBox("Escribe el modificador de velocidad."))
    fMain.lblVel.Caption = "VELOCIDAD: " & VelMod
End Sub

Private Sub m_a_verframe_Click()
    m_a_verframe.Checked = Not m_a_verframe.Checked
    VerFrame = m_a_verframe.Checked
    fMain.lblfr.Caption = animCounter
End Sub

Private Sub m_c_activado_Click()
    m_c_activado.Checked = Not m_c_activado.Checked
    bBodyTest = m_c_activado.Checked
End Sub

Private Sub m_c_elegir_Click()
Dim Cuerpo As Integer
    Cuerpo = InputBox("Selecciona el cuerpo de prueba.", "TestBody", num_test_body)
    If Cuerpo > 0 And Cuerpo <= NumNewBodys Then
        num_test_body = Cuerpo
    End If
    fMain.m_c_elegir.Caption = "Cuerpo Prueba: " & num_test_body
End Sub

Private Sub m_CF_Azul_Click()

    lColorFondo = D3DColorXRGB(0, 0, 255)
End Sub

Private Sub m_CF_Blanco_Click()

    lColorFondo = D3DColorXRGB(255, 255, 255)
End Sub

Private Sub m_CF_Negro_Click()
    lColorFondo = D3DColorXRGB(0, 0, 0)
End Sub

Private Sub m_CF_Rojo_Click()

    lColorFondo = D3DColorXRGB(255, 0, 0)
End Sub

Private Sub m_g_NeglectNegro_Click()
    m_g_NeglectNegro.Checked = Not m_g_NeglectNegro.Checked
    bNeglectNegro = m_g_NeglectNegro.Checked
End Sub

Private Sub m_i_Open_Click()
fIndexador.Show
End Sub

Private Sub m_mapHandler_Click()
    frmMapHandler.Show
End Sub

Private Sub m_z_Activado_Click()

    m_z_Activado.Checked = Not m_z_Activado.Checked
    bHeadTest = m_z_Activado.Checked
    
End Sub

Private Sub m_z_elegir_Click()
Dim Cuerpo As Integer
    Cuerpo = InputBox("Selecciona la cabeza de prueba.", "TestHead", num_test_head)
    If Cuerpo > 0 And Cuerpo <= Num_Heads Then
        num_test_head = Cuerpo
    End If
    fMain.m_z_elegir.Caption = "Cabeza Prueba: " & num_test_head
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
VelMov = Val(Text1)
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Modeling_Type = 2 Then
        NHeadData(sHead).OffsetOjos = Val(Text10.Text)
        ElseIf Modeling_Type = 3 Then
        NHelmetData(sHelmet).Alpha = Val(Text10.Text)
        ElseIf Modeling_Type = 4 Then
        nShieldDATA(sShield).Alpha = Val(Text10.Text)
        ElseIf Modeling_Type = 5 Then
        nWeaponData(sWeapon).Alpha = Val(Text10.Text)
        
        End If

    End If
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then

        If Modeling_Type = 3 Then
            NHelmetData(sHelmet).OffsetLat = Val(Text11.Text)
        ElseIf Modeling_Type = 4 Then
            nShieldDATA(sShield).OverWriteGrafico = Val(Text11.Text)
        ElseIf Modeling_Type = 5 Then
            nWeaponData(sWeapon).OverWriteGrafico = Val(Text11.Text)
        ElseIf Modeling_Type = 1 Then
            nBodyData(sBody).OverWriteGrafico = Val(Text11.Text)
            
        End If

    End If
End Sub

Private Sub Text12_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If IxS > 0 Then
    NewAnimationData(IxS).Desc = (Text12)
    fMain.Lista(UltimaOpcion).List(fMain.Lista(UltimaOpcion).ListIndex) = NewAnimationData(IxS).Desc & " (" & IxS & ")"
End If
End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If IxS > 0 Then
    NewAnimationData(IxS).Velocidad = Val(Text2)
End If
End If

End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim p As Long
Dim K As Long
Dim grafcounter As Integer

If IxS > 0 Then
    With NewAnimationData(IxS)
    
    .NumFrames = Val(Text3)
    If .NumFrames > 0 Then
        ReDim Preserve NewAnimationData(IxS).Indice(1 To .NumFrames)
        If .NumFrames = 0 Or .Columnas = 0 Or .Filas = 0 Then Exit Sub
            If .Initial = 0 Then .Initial = 1
        K = .Initial - 1
        If K >= (CInt(.Columnas) * CInt(.Filas)) Then
            K = K Mod (CInt(.Columnas) * CInt(.Filas))
        End If
        ReDim .Indice(1 To .NumFrames)
        grafcounter = .Grafico
        For p = 1 To .NumFrames
            K = K + 1
            .Indice(p).X = (((K - 1) Mod .Columnas) * .Width)
            .Indice(p).Y = ((Int((K - 1) / .Columnas)) * .Height)
            .Indice(p).Grafico = grafcounter
            If (K Mod (CInt(.Columnas) * CInt(.Filas))) = 0 And ((K + 1) - .Initial) < .NumFrames Then
                grafcounter = grafcounter + 1
                K = 0
            End If
        Next p
        End If
    End With
End If
End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If IxS > 0 Then
    NewAnimationData(IxS).Columnas = Val(ReadField(1, Text4, Asc("-")))
    NewAnimationData(IxS).Filas = Val(ReadField(2, Text4, Asc("-")))
    
End If
End If
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If IxS > 0 Then
    NewAnimationData(IxS).Width = Val(ReadField(1, Text5, Asc("-")))
    NewAnimationData(IxS).Height = Val(ReadField(2, Text5, Asc("-")))
    
End If
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim p As Long
Dim K As Long
Dim grafcounter As Integer

If IxS > 0 Then

        With NewAnimationData(IxS)
        If .NumFrames = 0 Then Exit Sub
        If .Columnas = 0 Or .Filas = 0 Then Exit Sub
    .Grafico = Val(Text6)
        If .Initial = 0 Then .Initial = 1
    K = .Initial - 1
    If K >= (CInt(.Columnas) * CInt(.Filas)) Then
        K = K Mod (CInt(.Columnas) * CInt(.Filas))
    End If
    ReDim .Indice(1 To .NumFrames)
    grafcounter = .Grafico
    For p = 1 To .NumFrames
        K = K + 1
        .Indice(p).X = (((K - 1) Mod .Columnas) * .Width)
        .Indice(p).Y = ((Int((K - 1) / .Columnas)) * .Height)
        .Indice(p).Grafico = grafcounter
        If (K Mod (CInt(.Columnas) * CInt(.Filas))) = 0 And ((K + 1) - .Initial) < .NumFrames Then
            grafcounter = grafcounter + 1
            K = 0
        End If
    Next p
    End With
    
    End If
End If

End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
Dim K As Long
Dim p As Long
Dim grafcounter As Long
If IxS > 0 Then
With NewAnimationData(IxS)
    If .NumFrames = 0 Then Exit Sub
    .Initial = Val(Text7)
    If .Initial = 0 Then .Initial = 1
    K = .Initial - 1
    If K >= (CInt(.Columnas) * CInt(.Filas)) Then
        K = K Mod (CInt(.Columnas) * CInt(.Filas))
    End If

    grafcounter = .Grafico
        ReDim .Indice(1 To .NumFrames)
    For p = 1 To .NumFrames
        K = K + 1
        .Indice(p).X = (((K - 1) Mod .Columnas) * .Width)
        .Indice(p).Y = ((Int((K - 1) / .Columnas)) * .Height)
        .Indice(p).Grafico = grafcounter
        If (K Mod (CInt(.Columnas) * CInt(.Filas))) = 0 And ((K + 1) - .Initial) < .NumFrames Then
            grafcounter = grafcounter + 1
            K = 0
        End If
    Next p
    End With
End If
End If

End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Modeling_Type = 2 Then
            NHeadData(sHead).OffsetDibujoY = Val(Text8.Text)
        ElseIf Modeling_Type = 3 Then
            NHelmetData(sHelmet).OffsetDibujoY = Val(Text8.Text)
        ElseIf Modeling_Type = 1 Then
            nBodyData(sBody).OffsetY = Val(Text8.Text)
        End If
    End If
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Modeling_Type = 2 Then
        NHeadData(sHead).Desc = Text9.Text
        ElseIf Modeling_Type = 3 Then
        NHelmetData(sHelmet).Desc = Text9.Text
        ElseIf Modeling_Type = 4 Then
            nShieldDATA(sShield).Desc = Text9.Text
        ElseIf Modeling_Type = 5 Then
            nWeaponData(sWeapon).Desc = Text9.Text
        ElseIf Modeling_Type = 1 Then
            nBodyData(sBody).Desc = Text9.Text
        End If
    End If
End Sub
