VERSION 5.00
Begin VB.Form fIndexador 
   Caption         =   "Indexador"
   ClientHeight    =   8550
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11820
   LinkTopic       =   "Form2"
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fAnim 
      Caption         =   "Nueva Animacion"
      Height          =   2655
      Left            =   120
      TabIndex        =   34
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Frame fCabezas 
      Caption         =   "Cabezas"
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command4 
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Numero:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliquea sobre el label de la direccion y luego sobre el grafico."
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label13 
         Caption         =   "Oeste:"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Sur:"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Este:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Norte:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6960
      TabIndex        =   21
      Text            =   "32"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      Text            =   "32"
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Indice"
      Height          =   2535
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton cmdIndex 
         Caption         =   "Indexar"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdReci 
         Caption         =   "Reciclado"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2080
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Index:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Grafico"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Indice:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Grafico:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox Actual 
      Height          =   7800
      Left            =   3960
      ScaleHeight     =   7740
      ScaleWidth      =   7620
      TabIndex        =   0
      Top             =   600
      Width           =   7680
   End
   Begin VB.Label Label4 
      Caption         =   "Estandar H:"
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Estandar W:"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mInd 
      Caption         =   "Indexar"
      Begin VB.Menu m_NuevoCasco 
         Caption         =   "Nuevo Casco"
      End
      Begin VB.Menu m_n_Cuerpo 
         Caption         =   "Nuevo Cuerpo"
      End
      Begin VB.Menu miCabeza 
         Caption         =   "Nueva Cabeza"
      End
      Begin VB.Menu m_n_Escudo 
         Caption         =   "Nuevo Escudo"
      End
      Begin VB.Menu m_n_Arma 
         Caption         =   "Nueva Arma"
      End
   End
End
Attribute VB_Name = "fIndexador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private aGrafico As Integer
Private agraficow As Integer
Private agraficoh As Integer
Private atx As Integer
Private nfx As Integer
Private nfy As Integer
Private aty As Integer
Private tw As Integer
Private th As Integer
Dim tX As Integer
Dim tY As Integer
Dim iX As Integer
Dim iY As Integer
Private tNumIndex As Integer
Private tAncho As Integer
Public Indexando As Byte

Public AcInd As Integer
Private AcFr(1 To 4) As Integer
Public qFr As Byte
Private qIndexo As Byte

Private tAlto As Integer
Public Function GetAcFr(ByVal Index As Integer) As Integer
    GetAcFr = AcFr(Index)
End Function
Public Sub ParseAcFr(ByVal Index As Integer, ByVal value As Integer)
    AcFr(Index) = value
    
End Sub

Private Sub Actual_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim z As Integer



    iX = (x / Screen.TwipsPerPixelX)
    iY = (y / Screen.TwipsPerPixelY)
    Actual.Cls
    If iX < agraficow And iY < agraficoh Then
        dxEngine.DibujareEnHwnd3 Actual.hwnd, aGrafico, 0, 0, True, 0, 0
        If Button = vbLeftButton Then
            For z = 1 To numNewIndex
                If NewIndexData(z).OverWriteGrafico = aGrafico Then
                    If NewIndexData(z).Estatic = 0 Then Exit Sub
                    With EstaticData(NewIndexData(z).Estatic)
                        If iX >= .L And iX < .L + .W Then
                            If iY >= .T And iY < .T + .H Then
                                'Es este el indice.
                                Actual.ForeColor = vbWhite
                                Actual.Line ((.L * Screen.TwipsPerPixelX), (.T * Screen.TwipsPerPixelY))-(((.L + .W) * Screen.TwipsPerPixelX), (.T * Screen.TwipsPerPixelY))
                                Actual.Line ((.L * Screen.TwipsPerPixelX), ((.T + .H) * Screen.TwipsPerPixelY))-(((.L + .W) * Screen.TwipsPerPixelX), ((.T + .H) * Screen.TwipsPerPixelY))
                                Actual.Line (((.L + .W) * Screen.TwipsPerPixelX), (.T * Screen.TwipsPerPixelY))-(((.L + .W) * Screen.TwipsPerPixelX), ((.T + .H) * Screen.TwipsPerPixelY))
                                Actual.Line ((.L * Screen.TwipsPerPixelX), (.T * Screen.TwipsPerPixelY))-((.L * Screen.TwipsPerPixelX), ((.T + .H) * Screen.TwipsPerPixelY))
                                If Indexando > 0 Then
                                    If Indexando = 1 Then
                                        If qFr = 1 Then
                                            Label5.Caption = "Norte: " & z
                                            Label5.ForeColor = vbWhite
                                        ElseIf qFr = 2 Then
                                            Label6.Caption = "Este: " & z
                                            Label6.ForeColor = vbWhite
                                        ElseIf qFr = 3 Then
                                            Label7.Caption = "Sur: " & z
                                            Label7.ForeColor = vbWhite
                                        ElseIf qFr = 4 Then
                                            Label13.Caption = "Oeste: " & z
                                            Label8.ForeColor = vbWhite
                                        End If
                                        AcFr(qFr) = z
                                
                                    End If
                            
                                End If
                                Exit For
                        
                        
                            End If
                        End If
            
        
                    End With
                End If
            Next z
            tw = Val(Text3.Text)
            th = Val(Text4.Text)
            If z > numNewIndex Then
                'No esta indexado.
                tX = iX \ tw
                tY = iY \ th
                tX = tX * tw
                tY = tY * th
    
                Text2.Text = 0
                Text7.Text = tX
                Text8.Text = tY
                Text9.Text = tw
                Text10.Text = th
                setcelda
                Text11.Text = z
                cmdIndex.Enabled = True
                cmdReci.Enabled = True
        
            Else
                If NewIndexData(z).Estatic > 0 Then
                    Text2.Text = z
                    Text7.Text = EstaticData(NewIndexData(z).Estatic).L
                    Text8.Text = EstaticData(NewIndexData(z).Estatic).T
                    Text9.Text = EstaticData(NewIndexData(z).Estatic).W
                    Text10.Text = EstaticData(NewIndexData(z).Estatic).H
                    Text11.Text = z
                    cmdIndex.Enabled = False
                    cmdReci.Enabled = False
                End If
            End If
        ElseIf Button = vbRightButton Then

        End If
    End If
End Sub

Private Sub cmdIndex_Click()
    Dim p As Long
    If Val(Text11.Text) = numNewIndex + 1 Then
        numNewIndex = numNewIndex + 1
        ReDim Preserve NewIndexData(1 To numNewIndex)
        WriteVar App.Path & "\RES\INDEX\NewIndex.dat", "INIT", "NUM", CStr(numNewIndex)
    ElseIf Val(Text11.Text) > numNewIndex + 1 Then
        MsgBox "Estás dejando indices sin usar."
        Exit Sub
    End If


    For p = 1 To numNewEstatic
        If EstaticData(p).L = Val(Text7.Text) Then
            If EstaticData(p).T = Val(Text8.Text) Then
                If EstaticData(p).W = Val(Text9.Text) Then
                    If EstaticData(p).H = Val(Text10.Text) Then
                        Exit For
                    End If
                End If
            End If
        End If
    Next p
    If p > numNewEstatic Then
        'No esta indexado.
        numNewEstatic = numNewEstatic + 1
        ReDim Preserve EstaticData(1 To numNewEstatic)
        EstaticData(p).L = Val(Text7.Text)
        EstaticData(p).T = Val(Text8.Text)
        EstaticData(p).W = Val(Text9.Text)
        EstaticData(p).H = Val(Text10.Text)
        WriteVar App.Path & "\RES\INDEX\NewEstatics.dat", "INIT", "NUM", CStr(numNewEstatic)
        WriteVar App.Path & "\RES\INDEX\NewEstatics.dat", CStr(p), "Left", CStr(EstaticData(p).L)
        WriteVar App.Path & "\RES\INDEX\NewEstatics.dat", CStr(p), "Top", CStr(EstaticData(p).T)
        WriteVar App.Path & "\RES\INDEX\NewEstatics.dat", CStr(p), "Width", CStr(EstaticData(p).W)
        WriteVar App.Path & "\RES\INDEX\NewEstatics.dat", CStr(p), "Height", CStr(EstaticData(p).H)
    End If

    NewIndexData(Val(Text11.Text)).Estatic = p
    NewIndexData(Val(Text11.Text)).OverWriteGrafico = Val(Text1.Text)
    
    WriteVar App.Path & "\RES\INDEX\NewIndex.dat", Text11.Text, "Estatica", CStr(p)
    WriteVar App.Path & "\RES\INDEX\NewIndex.dat", Text11.Text, "OverWriteGrafico", Val(Text1.Text)
    WriteVar App.Path & "\RES\INDEX\NewIndex.dat", Text11.Text, "Desc", InputBox("Escribe descripcion")
    
    Text2.Text = (Text11.Text)
End Sub

Private Sub cmdReci_Click()
    Dim p As Long
    Text11.Text = "..."
    DoEvents
    For p = 5 To numNewIndex
        If UCase$(left$(GetVar(App.Path & "\RES\INDEX\NewIndex.dat", CStr(p), "OverWriteGrafico"), 1)) = "R" Then
            Text11.Text = p
            Exit For
        End If
    Next p
End Sub



Private Sub Command1_Click()
    'Busca grafico y le pone en el Actual.
    If Val(Text1.Text) > 0 Then
        'Esta escrito un grafico.
        aGrafico = Val(Text1.Text)
    ElseIf Val(Text2.Text) > 0 Then
        'Esta escrito un index
        aGrafico = NewIndexData(Val(Text2.Text)).OverWriteGrafico
    End If
    Actual.Cls
    dxEngine.DibujareEnHwnd3 Actual.hwnd, aGrafico, 0, 0, True, agraficow, agraficoh

    
End Sub

Private Sub Command2_Click()
    Indexando = 0
    qIndexo = 0
    fCabezas.Visible = False
    Label5.Caption = "Norte: "
    Label5.ForeColor = vbWhite
    
    Label6.Caption = "Este: "
    Label6.ForeColor = vbWhite
    Label7.Caption = "Sur: "
    Label7.ForeColor = vbWhite
    Label13.Caption = "Oeste: "
    Label13.ForeColor = vbWhite
End Sub


Private Sub Command3_Click()
    If qIndexo = 1 Then
        If AcInd <= Num_Heads Then

            NHeadData(AcInd).Frame(1) = AcFr(1)
            NHeadData(AcInd).Frame(2) = AcFr(2)
            NHeadData(AcInd).Frame(3) = AcFr(3)
            NHeadData(AcInd).Frame(4) = AcFr(4)


        Else
            ReDim Preserve NHeadData(1 To Num_Heads + 1)
            Num_Heads = Num_Heads + 1
            WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "INIT", "NUM", CStr(Num_Heads)

            NHeadData(AcInd).Frame(1) = AcFr(1)
            NHeadData(AcInd).Frame(2) = AcFr(2)
            NHeadData(AcInd).Frame(3) = AcFr(3)
            NHeadData(AcInd).Frame(4) = AcFr(4)
            NHeadData(AcInd).Raza = 1
            NHeadData(AcInd).Genero = 1
        End If

        WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & AcInd, "NORTH", CStr(AcFr(1))
        WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & AcInd, "EAST", CStr(AcFr(2))
        WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & AcInd, "SOUTH", CStr(AcFr(3))
        WriteVar App.Path & "\RES\INDEX\NewHeads.dat", "HEAD" & AcInd, "WEST", CStr(AcFr(4))

        Label5.Caption = "Norte: "
        Label5.ForeColor = vbWhite

        Label6.Caption = "Este: "
        Label6.ForeColor = vbWhite
        Label7.Caption = "Sur: "
        Label7.ForeColor = vbWhite
        Label13.Caption = "Oeste: "
        Label13.ForeColor = vbWhite
        fCabezas.Visible = False
        
    ElseIf qIndexo = 2 Then
        If AcInd <= Num_Helmets Then

            NHelmetData(AcInd).Frame(1) = AcFr(1)
            NHelmetData(AcInd).Frame(2) = AcFr(2)
            NHelmetData(AcInd).Frame(3) = AcFr(3)
            NHelmetData(AcInd).Frame(4) = AcFr(4)

        Else
            ReDim Preserve NHelmetData(1 To Num_Helmets + 1)
            Num_Helmets = Num_Helmets + 1
            WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "INIT", "NUM", CStr(Num_Helmets)

            NHelmetData(AcInd).Frame(1) = AcFr(1)
            NHelmetData(AcInd).Frame(2) = AcFr(2)
            NHelmetData(AcInd).Frame(3) = AcFr(3)
            NHelmetData(AcInd).Frame(4) = AcFr(4)
        End If

        WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "Helmet" & AcInd, "NORTH", CStr(AcFr(1))
        WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "Helmet" & AcInd, "EAST", CStr(AcFr(2))
        WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "Helmet" & AcInd, "SOUTH", CStr(AcFr(3))
        WriteVar App.Path & "\RES\INDEX\NewHelmets.dat", "Helmet" & AcInd, "WEST", CStr(AcFr(4))

        Label5.Caption = "Norte: "
        Label5.ForeColor = vbWhite

        Label6.Caption = "Este: "
        Label6.ForeColor = vbWhite
        Label7.Caption = "Sur: "
        Label7.ForeColor = vbWhite
        Label13.Caption = "Oeste: "
        Label13.ForeColor = vbWhite
        fCabezas.Visible = False
        
    ElseIf qIndexo = 3 Then
        If AcInd <= NumNewBodys Then

            nBodyData(AcInd).mMovement(1) = AcFr(1)
            nBodyData(AcInd).mMovement(2) = AcFr(2)
            nBodyData(AcInd).mMovement(3) = AcFr(3)
            nBodyData(AcInd).mMovement(4) = AcFr(4)


        Else
            ReDim Preserve nBodyData(1 To NumNewBodys + 1)
            NumNewBodys = NumNewBodys + 1
            WriteVar App.Path & "\RES\INDEX\newBody.dat", "INIT", "NUM", CStr(NumNewBodys)
            nBodyData(AcInd).mMovement(1) = AcFr(1)
            nBodyData(AcInd).mMovement(2) = AcFr(2)
            nBodyData(AcInd).mMovement(3) = AcFr(3)
            nBodyData(AcInd).mMovement(4) = AcFr(4)
        End If

        WriteVar App.Path & "\RES\INDEX\Newbody.dat", CStr(AcInd), "MOV1", CStr(AcFr(1))
        WriteVar App.Path & "\RES\INDEX\NewBody.dat", CStr(AcInd), "MOV2", CStr(AcFr(2))
        WriteVar App.Path & "\RES\INDEX\NewBody.dat", CStr(AcInd), "MOV3", CStr(AcFr(3))
        WriteVar App.Path & "\RES\INDEX\NewBody.dat", CStr(AcInd), "MOV4", CStr(AcFr(4))

        Label5.Caption = "Norte: "
        Label5.ForeColor = vbWhite

        Label6.Caption = "Este: "
        Label6.ForeColor = vbWhite
        Label7.Caption = "Sur: "
        Label7.ForeColor = vbWhite
        Label13.Caption = "Oeste: "
        Label13.ForeColor = vbWhite
        fCabezas.Visible = False
    ElseIf qIndexo = 4 Then
        If AcInd <= NumNewShields Then

            nShieldDATA(AcInd).mMovimiento(1) = AcFr(1)
            nShieldDATA(AcInd).mMovimiento(2) = AcFr(2)
            nShieldDATA(AcInd).mMovimiento(3) = AcFr(3)
            nShieldDATA(AcInd).mMovimiento(4) = AcFr(4)


        Else
            ReDim Preserve nShieldDATA(1 To NumNewShields + 1)
            NumNewShields = NumNewShields + 1
            WriteVar App.Path & "\RES\INDEX\NwShields.dat", "INIT", "NUM", CStr(NumNewShields)
            nShieldDATA(AcInd).mMovimiento(1) = AcFr(1)
            nShieldDATA(AcInd).mMovimiento(2) = AcFr(2)
            nShieldDATA(AcInd).mMovimiento(3) = AcFr(3)
            nShieldDATA(AcInd).mMovimiento(4) = AcFr(4)




        End If

        WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(AcInd), "MOV1", CStr(AcFr(1))
        WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(AcInd), "MOV2", CStr(AcFr(2))
        WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(AcInd), "MOV3", CStr(AcFr(3))
        WriteVar App.Path & "\RES\INDEX\NwShields.dat", CStr(AcInd), "MOV4", CStr(AcFr(4))

        Label5.Caption = "Norte: "
        Label5.ForeColor = vbWhite

        Label6.Caption = "Este: "
        Label6.ForeColor = vbWhite
        Label7.Caption = "Sur: "
        Label7.ForeColor = vbWhite
        Label13.Caption = "Oeste: "
        Label13.ForeColor = vbWhite
        fCabezas.Visible = False
    ElseIf qIndexo = 5 Then
        If AcInd <= NumNewWeapons Then

            nWeaponData(AcInd).mMovimiento(1) = AcFr(1)
            nWeaponData(AcInd).mMovimiento(2) = AcFr(2)
            nWeaponData(AcInd).mMovimiento(3) = AcFr(3)
            nWeaponData(AcInd).mMovimiento(4) = AcFr(4)


        Else
            ReDim Preserve nWeaponData(1 To NumNewWeapons + 1)
            NumNewWeapons = NumNewWeapons + 1
            WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", "INIT", "NUM", CStr(NumNewWeapons)
            nWeaponData(AcInd).mMovimiento(1) = AcFr(1)
            nWeaponData(AcInd).mMovimiento(2) = AcFr(2)
            nWeaponData(AcInd).mMovimiento(3) = AcFr(3)
            nWeaponData(AcInd).mMovimiento(4) = AcFr(4)




        End If

        WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(AcInd), "MOV1", CStr(AcFr(1))
        WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(AcInd), "MOV2", CStr(AcFr(2))
        WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(AcInd), "MOV3", CStr(AcFr(3))
        WriteVar App.Path & "\RES\INDEX\NwWeapons.dat", CStr(AcInd), "MOV4", CStr(AcFr(4))

        Label5.Caption = "Norte: "
        Label5.ForeColor = vbWhite

        Label6.Caption = "Este: "
        Label6.ForeColor = vbWhite
        Label7.Caption = "Sur: "
        Label7.ForeColor = vbWhite
        Label13.Caption = "Oeste: "
        Label13.ForeColor = vbWhite
        fCabezas.Visible = False

    End If

End Sub









Private Sub Command4_Click()
    
    If qIndexo = 4 Then
        Text5.Text = NumNewShields + 1
        AcInd = NumNewShields + 1
        Command3.Enabled = True
    
        If MsgBox("¿Deseas utilizar la animación standard para Escudos?", vbOKCancel) = vbOK Then
            AcFr(1) = Standard_Escudo_North
            AcFr(2) = Standard_Escudo_East
            AcFr(3) = Standard_Escudo_South
            AcFr(4) = Standard_Escudo_West
        End If
    ElseIf qIndexo = 5 Then
        Text5.Text = NumNewWeapons + 1
        AcInd = NumNewWeapons + 1
        If MsgBox("¿Deseas utilizar la animación standard para Armas?", vbOKCancel) = vbOK Then
            AcFr(1) = Standard_Arma_North
            AcFr(2) = Standard_Arma_East
            AcFr(3) = Standard_Arma_South
            AcFr(4) = Standard_Arma_West
        End If
        Command3.Enabled = True
    ElseIf qIndexo = 3 Then
        Text5.Text = NumNewBodys + 1
        AcInd = NumNewBodys + 1
        Command3.Enabled = True
        If MsgBox("¿Deseas utilizar la animación standard para cuerpos?", vbOKCancel) = vbOK Then
            If MsgBox("¿Deseas utilizar animacion pequeña?", vbYesNo) = vbYes Then
                If MsgBox("¿Deseas utilizar la animacion pequeña de arriba?", vbYesNo) = vbYes Then
                    AcFr(1) = Standard_Cuerpo_North_Small
                    AcFr(2) = Standard_Cuerpo_East_Small
                    AcFr(3) = Standard_Cuerpo_South_Small
                    AcFr(4) = Standard_Cuerpo_West_Small
                Else
                    AcFr(1) = Standard_Cuerpo_North_Small2
                    AcFr(2) = Standard_Cuerpo_East_Small2
                    AcFr(3) = Standard_Cuerpo_South_Small2
                    AcFr(4) = Standard_Cuerpo_West_Small2
            
                End If
        
            Else
                AcFr(1) = Standard_Cuerpo_North
                AcFr(2) = Standard_Cuerpo_East
                AcFr(3) = Standard_Cuerpo_South
                AcFr(4) = Standard_Cuerpo_West
            End If
        End If

    End If

End Sub

Private Sub Label13_Click()
    If qIndexo <= 2 Then
        qFr = 4
        Label13.ForeColor = vbYellow
    ElseIf qIndexo >= 3 And qIndexo <= 5 Then
        qFr = 4
        Call InPutform.Parse("Selecciona la animación deseada.", 0)


    End If
End Sub

Private Sub Label5_Click()
    If qIndexo <= 2 Then
        qFr = 1
        Label5.ForeColor = vbYellow
    ElseIf qIndexo >= 3 And qIndexo <= 5 Then
        qFr = 1
        Call InPutform.Parse("Selecciona la animación deseada.", 0)

    End If
End Sub

Private Sub Label6_Click()
    If qIndexo <= 2 Then
        qFr = 2
        Label6.ForeColor = vbYellow
    ElseIf qIndexo >= 3 And qIndexo <= 5 Then
        qFr = 2
        Call InPutform.Parse("Selecciona la animación deseada.", 0)


    End If
End Sub

Private Sub Label7_Click()
    If qIndexo <= 2 Then
        qFr = 3
        Label7.ForeColor = vbYellow
    ElseIf qIndexo >= 3 And qIndexo <= 5 Then
        qFr = 3
        Call InPutform.Parse("Selecciona la animación deseada.", 0)

    End If
End Sub

Private Sub m_n_Arma_Click()
    qIndexo = 5
    fCabezas.Visible = True
    fCabezas.Caption = "Armas"
    Indexando = 1
    Command3.Enabled = False
End Sub

Private Sub m_n_Cuerpo_Click()
    qIndexo = 3
    fCabezas.Visible = True
    fCabezas.Caption = "Cuerpos"
    Indexando = 1
    Command3.Enabled = False
    
End Sub

Private Sub m_n_Escudo_Click()
    qIndexo = 4
    Text5.Text = vbNullString
    fCabezas.Visible = True
    fCabezas.Caption = "Escudos"
    Command3.Enabled = False
    Indexando = 1
End Sub

Private Sub m_NuevoCasco_Click()
    fCabezas.Visible = True
    fCabezas.Caption = "Cascos"
    Text5.Text = Num_Helmets + 1
    AcInd = Num_Helmets + 1
    Indexando = 1
    qIndexo = 2
End Sub

Private Sub miCabeza_Click()
    fCabezas.Visible = True
    fCabezas.Caption = "Cabezas"
    Text5.Text = Num_Heads + 1
    AcInd = Num_Heads + 1
    Indexando = 1
    qIndexo = 1
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        dxEngine.DibujareEnHwnd3 Actual.hwnd, aGrafico, 0, 0, True, agraficow, agraficoh
        setcelda
    End If
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then


        AcInd = Val(Text5.Text)
        Select Case qIndexo
            Case 1 ' Cabezas
                If AcInd > Num_Heads Then
                    AcInd = Num_Heads + 1
                    Text5.Text = Num_Heads + 1
                Else
                    AcInd = Val(Text5.Text)
                    AcFr(1) = NHeadData(AcInd).Frame(1)
                    AcFr(2) = NHeadData(AcInd).Frame(2)
                    AcFr(3) = NHeadData(AcInd).Frame(3)
                    AcFr(4) = NHeadData(AcInd).Frame(4)
                    Label5.Caption = "Norte: " & AcFr(1)
                    Label6.Caption = "Este: " & AcFr(2)
                    Label7.Caption = "Sur: " & AcFr(3)
                    Label13.Caption = "Oeste: " & AcFr(4)
                End If
            Case 2 'Cascos
                If AcInd > Num_Helmets Then
                    AcInd = Num_Helmets + 1
                    Text5.Text = Num_Helmets + 1
                Else
                    AcInd = Val(Text5.Text)
                    AcFr(1) = NHelmetData(AcInd).Frame(1)
                    AcFr(2) = NHelmetData(AcInd).Frame(2)
                    AcFr(3) = NHelmetData(AcInd).Frame(3)
                    AcFr(4) = NHelmetData(AcInd).Frame(4)
                    Label5.Caption = "Norte: " & AcFr(1)
                    Label6.Caption = "Este: " & AcFr(2)
                    Label7.Caption = "Sur: " & AcFr(3)
                    Label13.Caption = "Oeste: " & AcFr(4)
                End If
        
            Case 3 'Cuerpos
                If AcInd > NumNewBodys Then
                    AcInd = NumNewBodys + 1
                    Text5.Text = AcInd
                Else
                    AcFr(1) = nBodyData(AcInd).mMovement(1)
                    AcFr(2) = nBodyData(AcInd).mMovement(2)
                    AcFr(3) = nBodyData(AcInd).mMovement(3)
                    AcFr(4) = nBodyData(AcInd).mMovement(4)
                    Label5.Caption = "Norte: " & AcFr(1)
                    Label6.Caption = "Este: " & AcFr(2)
                    Label7.Caption = "Sur: " & AcFr(3)
                    Label13.Caption = "Oeste: " & AcFr(4)
                End If
            Case 4 'Escudos
                If AcInd > NumNewShields Then
                    AcInd = NumNewShields + 1
                    Text5.Text = AcInd
                Else
                    AcFr(1) = nShieldDATA(AcInd).mMovimiento(1)
                    AcFr(2) = nShieldDATA(AcInd).mMovimiento(2)
                    AcFr(3) = nShieldDATA(AcInd).mMovimiento(3)
                    AcFr(4) = nShieldDATA(AcInd).mMovimiento(4)
                    Label5.Caption = "Norte: " & AcFr(1)
                    Label6.Caption = "Este: " & AcFr(2)
                    Label7.Caption = "Sur: " & AcFr(3)
                    Label13.Caption = "Oeste: " & AcFr(4)
                End If
            Case 5 'Armas
                If AcInd > NumNewWeapons Then
                    AcInd = NumNewWeapons + 1
                    Text5.Text = AcInd
                Else
                    AcFr(1) = nWeaponData(AcInd).mMovimiento(1)
                    AcFr(2) = nWeaponData(AcInd).mMovimiento(2)
                    AcFr(3) = nWeaponData(AcInd).mMovimiento(3)
                    AcFr(4) = nWeaponData(AcInd).mMovimiento(4)
                    Label5.Caption = "Norte: " & AcFr(1)
                    Label6.Caption = "Este: " & AcFr(2)
                    Label7.Caption = "Sur: " & AcFr(3)
                    Label13.Caption = "Oeste: " & AcFr(4)
                End If
        End Select
        Command3.Enabled = True
        
    End If


End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        dxEngine.DibujareEnHwnd3 Actual.hwnd, aGrafico, 0, 0, True, agraficow, agraficoh

        setcelda
    End If
End Sub
Sub setcelda()

    th = Val(Text10.Text)
    tw = Val(Text9.Text)
    tX = Val(Text7.Text)
    tY = Val(Text8.Text)
        
    Actual.ForeColor = vbRed
    Actual.Line (tX * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)-((tX + tw) * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)
    Actual.Line (tX * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)-((tX + tw) * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)
    Actual.Line (tX * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)-(tX * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)
    Actual.Line ((tX + tw) * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)-((tX + tw) * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)

End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        dxEngine.DibujareEnHwnd3 Actual.hwnd, aGrafico, 0, 0, True, agraficow, agraficoh
    
        setcelda
    End If
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        dxEngine.DibujareEnHwnd3 Actual.hwnd, aGrafico, 0, 0, True, agraficow, agraficoh
   
        setcelda
    End If
End Sub
