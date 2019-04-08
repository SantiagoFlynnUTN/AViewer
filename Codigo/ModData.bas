Attribute VB_Name = "ModData"
Option Explicit

Public Const Standard_Cuerpo_North As Integer = 86
Public Const Standard_Cuerpo_East As Integer = 88
Public Const Standard_Cuerpo_South As Integer = 85
Public Const Standard_Cuerpo_West As Integer = 87

Public Const Standard_Cuerpo_North_Small As Integer = 232
Public Const Standard_Cuerpo_East_Small As Integer = 234
Public Const Standard_Cuerpo_South_Small As Integer = 231
Public Const Standard_Cuerpo_West_Small As Integer = 233

Public Const Standard_Cuerpo_North_Small2 As Integer = 236
Public Const Standard_Cuerpo_East_Small2 As Integer = 238
Public Const Standard_Cuerpo_South_Small2 As Integer = 235
Public Const Standard_Cuerpo_West_Small2 As Integer = 237

Public Const Standard_Escudo_North As Integer = 240
Public Const Standard_Escudo_East As Integer = 242
Public Const Standard_Escudo_South As Integer = 239
Public Const Standard_Escudo_West As Integer = 241
Public Const Standard_Arma_North As Integer = 240
Public Const Standard_Arma_East As Integer = 242
Public Const Standard_Arma_South As Integer = 239
Public Const Standard_Arma_West As Integer = 241

Public num_test_body As Integer
Public Const Test_body_def As Integer = 11
Public num_test_head As Integer
Public Const Test_Head_Def As Integer = 5


Public Type tNewMunicionData
    mMovimiento(1 To 4) As Integer
    Alpha As Byte
    OverWriteGrafico As Integer
    Desc As String
End Type
Public Type tNewCapa
    mMovimiento(1 To 4) As Integer
    
    Alpha As Byte
    Desc As String
    aOverWriteGrafico As Integer
    pOverWriteGrafico As Integer
End Type
Public Type tNewShieldData
    mMovimiento(1 To 4) As Integer
    Alpha As Byte
    OverWriteGrafico As Integer
    Desc As String
End Type
Public Type tNewWeaponData
    mMovimiento(1 To 4) As Integer
    Alpha As Byte
    OverWriteGrafico As Integer
    Desc As String
End Type
Public nShieldDATA() As tNewShieldData
Public nWeaponData() As tNewWeaponData
Public nMunicionData() As tNewMunicionData
Public nCapaData() As tNewCapa
Public NumNewShields As Integer
Public NumNewWeapons As Integer
Public NumNewM As Integer
Public NumNewCapas As Integer



Public Type tNewIndice
    x As Integer
    y As Integer
    Grafico As Integer
End Type
Public Type tNewBody
    mMovement(1 To 4) As Integer
    Reposo(1 To 4) As Integer
    Attack(1 To 4) As Integer
    Death(1 To 4) As Integer
    Attacked(1 To 4) As Integer
    
    bAtacado As Boolean
    bReposo As Boolean
    bAtaque As Boolean
    bDeath As Boolean
    bContinuo As Boolean
    OverWriteGrafico As Integer
    OffsetY As Integer
    Capa As Integer
    Desc As String
End Type
Public Type tNewBodyData
    mMovement(1 To 4) As Integer
    Reposo(1 To 4) As Integer
    Attack(1 To 4) As Integer
    Death(1 To 4) As Integer
    Attacked(1 To 4) As Integer
    
    bAtacado As Boolean
    bReposo As Boolean
    bAtaque As Boolean
    bDeath As Boolean
    bContinuo As Boolean
    
    Capa As Integer
End Type
Public nBodyData() As tNewBody
Public NumNewBodys As Integer
Private Type tConfigClient
    CargaDinamica As Boolean
    CargaEstaticaBinaria As Boolean
    CambiarResolucion   As Boolean
    PreguntarChangeRes  As Boolean
    CuentasOwnerHD      As Long
    LimitarFPS          As Boolean
End Type

Public CConfig As tConfigClient
Public Type tNewIndex
    Estatic As Integer ' Info de estatica
    Dinamica As Integer ' Animacion
    OverWriteGrafico As Integer ' Grafico
End Type
Public Type tNewEstatic
    W As Integer
    H As Integer
    L As Integer
    T As Integer
    th As Single
    tw As Single
End Type
Public Type tNewAnimation
    Desc As String
    Numero As Integer
    Grafico As Integer
    NumFrames As Byte
    Filas As Byte
    Columnas As Byte
    Indice() As tNewIndice
    Width As Integer
    Height As Integer
    IndiceCounter As Single
    Velocidad As Single
    TileWidth As Single
    TileHeight As Single
    
    Romboidal As Byte
    Direction As Integer
    
    OffsetX As Integer
    OffsetY As Integer
    Initial As Integer
End Type
Public Type tnHead
    Frame(1 To 4) As Integer
    OffsetDibujoY As Integer
    OffsetOjos As Integer
    Raza As Byte
    Genero As Byte
    Desc As String
End Type
Public Type tnHelmets
    Frame(1 To 4) As Integer
    OffsetDibujoY As Integer
    OffsetLat As Integer
    Alpha As Byte
    Desc As String
End Type
Public NHelmetData() As tnHelmets
Public Num_Helmets As Integer
Public NHeadData() As tnHead
Public Num_Heads As Integer
Public NewAnimationData() As tNewAnimation
Public Num_NwAnim As Integer
Public NewIndexData() As tNewIndex
Public numNewIndex As Integer
Public EstaticData() As tNewEstatic
Public numNewEstatic As Integer
Public Sub Load_NewEstatics()
    Dim S As String
    Dim i As Long
    Dim z As Long

    S = App.Path & "\RES\INDEX\NewEstatics.dat"


    If numNewEstatic > 0 Then
        ReDim EstaticData(1 To numNewEstatic)
        For i = 1 To numNewEstatic
            With EstaticData(i)
                .L = Val(GetVar(S, CStr(i), "Left"))
                .T = Val(GetVar(S, CStr(i), "Top"))
                .W = Val(GetVar(S, CStr(i), "Width"))
                .H = Val(GetVar(S, CStr(i), "Height"))
                .tw = .W / 32
                .th = .H / 32
        
            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i
    End If
End Sub
Public Sub GuardarAnim(ByVal i As Integer)
    Dim S As String
    Dim K As String

    S = App.Path & "\RES\INDEX\NewAnim.dat"
    K = "ANIMACION" & i
    If i <= Num_NwAnim Then

        WriteVar S, "NW_ANIM", "NUM", CStr(Num_NwAnim)
        WriteVar S, K, "Grafico", CStr(NewAnimationData(i).Grafico)
        WriteVar S, K, "NumeroFrames", CStr(NewAnimationData(i).NumFrames)
        WriteVar S, K, "Columnas", CStr(NewAnimationData(i).Columnas)
        WriteVar S, K, "Filas", CStr(NewAnimationData(i).Filas)
        WriteVar S, K, "Ancho", CStr(NewAnimationData(i).Width)
        WriteVar S, K, "Alto", CStr(NewAnimationData(i).Height)
        WriteVar S, K, "Velocidad", CStr(NewAnimationData(i).Velocidad)
        WriteVar S, K, "Inicial", CStr(NewAnimationData(i).Initial)
        WriteVar S, K, "Desc", NewAnimationData(i).Desc

    End If

End Sub
Public Sub Load_NewAnimation()
    Dim S As String
    Dim i As Long
    Dim p As Long
    Dim K As Integer
    Dim grafcounter As Integer
    Dim j As String

    If Num_NwAnim < 1 Then Exit Sub

    ReDim NewAnimationData(1 To Num_NwAnim)

    For i = 1 To Num_NwAnim


        fCargando.PB.value = fCargando.PB.value + 1
        S = App.Path & "\RES\INDEX\NewAnim.dat"


        With NewAnimationData(i)
            .Numero = i
            .Grafico = Val(GetVar(S, "ANIMACION" & i, "Grafico"))
            .Columnas = Val(GetVar(S, "ANIMACION" & i, "Columnas"))
            .Filas = Val(GetVar(S, "ANIMACION" & i, "Filas"))
            .Height = Val(GetVar(S, "ANIMACION" & i, "Alto"))
            .Width = Val(GetVar(S, "ANIMACION" & i, "Ancho"))
            .NumFrames = Val(GetVar(S, "ANIMACION" & i, "NumeroFrames"))
            j = (GetVar(S, "ANIMACION" & i, "Velocidad"))
            .Velocidad = Val(Replace(j, ",", "."))
            .TileWidth = .Width / 32
            .TileHeight = .Height / 32
            .Romboidal = Val(GetVar(S, "ANIMACION" & i, "AnimacionRomboidal"))
            .OffsetX = Val(GetVar(S, "ANIMACION" & i, "OffsetX"))
            .OffsetY = Val(GetVar(S, "ANIMACION" & i, "OffsetY"))
            .Initial = Val(GetVar(S, "ANIMACION" & i, "Inicial"))
            .Desc = GetVar(S, "ANIMACION" & i, "Desc")
            If .NumFrames > 0 Then
                ReDim .Indice(1 To .NumFrames) As tNewIndice
                grafcounter = .Grafico
                If .Initial = 0 Then .Initial = 1
                K = .Initial - 1
                If K >= (CInt(.Columnas) * CInt(.Filas)) Then
                    K = K Mod (CInt(.Columnas) * CInt(.Filas))
                End If


                For p = 1 To .NumFrames
                    K = K + 1
                    .Indice(p).x = (((K - 1) Mod .Columnas) * .Width)
                    .Indice(p).y = ((Int((K - 1) / .Columnas)) * .Height)
                    .Indice(p).Grafico = grafcounter
                    If (K Mod (CInt(.Columnas) * CInt(.Filas))) = 0 And ((K + 1) - .Initial) < .NumFrames Then
                        grafcounter = grafcounter + 1
                        K = 0
                    End If
                Next p
            End If
            If NewAnimationData(i).Grafico > 0 Then
                fMain.Lista(eOpciones.Animaciones_op).AddItem NewAnimationData(i).Desc & " (" & i & ")"
            End If

        End With
    Next i



End Sub

Public Sub Load_NewIndex()
    Dim S As String
    Dim i As Long
    Dim z As Long

    S = App.Path & "\RES\INDEX\NewIndex.dat"

    If numNewIndex > 0 Then
        ReDim NewIndexData(1 To numNewIndex)
        For i = 1 To numNewIndex
            With NewIndexData(i)
                .Dinamica = Val(GetVar(S, CStr(i), "Dinamica"))
                .Estatic = Val(GetVar(S, CStr(i), "Estatica"))
                .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
        
                If .OverWriteGrafico <> 0 Then
                    If .Dinamica = 0 And .Estatic > 0 Then
                        fMain.Lista(eOpciones.Indices_op).AddItem i & "-" & "[ESTATICO]"
                    ElseIf .Estatic = 0 And .Dinamica > 0 Then
                        fMain.Lista(eOpciones.Indices_op).AddItem i & "-" & "[ANIM]"
                    End If
                End If
        
            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i
    End If
End Sub
Public Sub Load_NewHeads()
    Dim S As String
    S = App.Path & "\RES\INDEX\NewHeads.dat"

    Dim i As Long


    If Num_Heads > 0 Then
        ReDim NHeadData(1 To Num_Heads)
        For i = 1 To Num_Heads
        
            NHeadData(i).Frame(E_Heading.NORTH) = Val(GetVar(S, "HEAD" & i, "NORTH"))
            NHeadData(i).Frame(E_Heading.EAST) = Val(GetVar(S, "HEAD" & i, "EAST"))
            NHeadData(i).Frame(E_Heading.SOUTH) = Val(GetVar(S, "HEAD" & i, "SOUTH"))
            NHeadData(i).Frame(E_Heading.WEST) = Val(GetVar(S, "HEAD" & i, "WEST"))
            NHeadData(i).OffsetDibujoY = Val(GetVar(S, "HEAD" & i, "OFFSET_DIBUJO"))
            NHeadData(i).OffsetOjos = Val(GetVar(S, "HEAD" & i, "OFFSET_OJOS"))
            NHeadData(i).Raza = Val(GetVar(S, "HEAD" & i, "RAZA"))
            NHeadData(i).Genero = Val(GetVar(S, "HEAD" & i, "GENERO"))
            NHeadData(i).Desc = GetVar(S, "HEAD" & i, "DESC")
            fCargando.PB.value = fCargando.PB.value + 1
        
        Next i
    End If
End Sub
Public Sub Load_NewHelmets()
    Dim S As String
    S = App.Path & "\RES\INDEX\NewHelmets.dat"

    Dim i As Long


    If Num_Helmets > 0 Then
        ReDim NHelmetData(1 To Num_Helmets)
        For i = 1 To Num_Helmets
        
            NHelmetData(i).Frame(E_Heading.NORTH) = Val(GetVar(S, "helmet" & i, "NORTH"))
            NHelmetData(i).Frame(E_Heading.EAST) = Val(GetVar(S, "helmet" & i, "EAST"))
            NHelmetData(i).Frame(E_Heading.SOUTH) = Val(GetVar(S, "helmet" & i, "SOUTH"))
            NHelmetData(i).Frame(E_Heading.WEST) = Val(GetVar(S, "helmet" & i, "WEST"))
            NHelmetData(i).OffsetDibujoY = Val(GetVar(S, "helmet" & i, "OFFSET_DIBUJO"))
            NHelmetData(i).Alpha = Val(GetVar(S, "helmet" & i, "ALPHA"))
            NHelmetData(i).OffsetLat = Val(GetVar(S, "helmet" & i, "OFFSET_LAT"))
        
            NHelmetData(i).Desc = GetVar(S, "helmet" & i, "DESC")
            fCargando.PB.value = fCargando.PB.value + 1
        
        Next i
    End If
End Sub
Public Sub Load_NewBodys()
    Dim S As String
    Dim i As Long
    Dim z As Long
    Dim x As Integer
    S = App.Path & "\RES\INDEX\NewBody.dat"

    If NumNewBodys > 0 Then
        ReDim nBodyData(1 To NumNewBodys)
        For i = 1 To NumNewBodys

            With nBodyData(i)

                .bAtaque = IIf(Val(GetVar(S, CStr(i), "Ataque")), True, False)
                .bContinuo = IIf(Val(GetVar(S, CStr(i), "Continuo")), True, False)
                .bReposo = IIf(Val(GetVar(S, CStr(i), "Reposo")), True, False)
                .bAtacado = IIf(Val(GetVar(S, CStr(i), "Atacado")), True, False)
                .bDeath = IIf(Val(GetVar(S, CStr(i), "Muerte")), True, False)
                .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
                .OffsetY = Val(GetVar(S, CStr(i), "OffsetY"))
                .Desc = GetVar(S, CStr(i), "Desc")
                If .bAtaque Then
                    For z = 1 To 4
                        x = Val(GetVar(S, CStr(i), "Ataque" & z))
                        .Attack(z) = x 'NewAnimationData(X)
                    Next z
                End If
                If .bAtacado Then
                    For z = 1 To 4
                        x = Val(GetVar(S, CStr(i), "Atacado" & z))
                        .Attacked(z) = x ' NewAnimationData(X)
                    Next z
                End If
        
                If .bReposo Then
                    For z = 1 To 4
                        x = Val(GetVar(S, CStr(i), "Reposo" & z))
                        .Reposo(z) = x 'xNewAnimationData(X)

                    Next z
                End If
                If .bDeath Then
                    For z = 1 To 4
                        x = Val(GetVar(S, CStr(i), "Muerte" & z))
                        .Death(z) = x ' NewAnimationData(X)

                    Next z
                End If

                For z = 1 To 4
        
                    x = Val(GetVar(S, CStr(i), "Mov" & z))
                    If x > 0 Then .mMovement(z) = x ' NewAnimationData(X)
                Next z

            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i
    End If

End Sub
Public Sub Load_NewShields()



    Dim S As String
    Dim i As Long
    Dim z As Long
    S = App.Path & "\RES\INDEX\NwShields.dat"
    If NumNewShields > 0 Then
        ReDim nShieldDATA(1 To NumNewShields)
        For i = 1 To NumNewShields

            With nShieldDATA(i)

                .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
                .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
                For z = 1 To 4
        
                    .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
                Next z
                .Desc = GetVar(S, CStr(i), "Desc")
            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i
    End If
End Sub

Public Sub Load_NewWeapons()

    Dim S As String
    Dim i As Long
    Dim z As Long

    S = App.Path & "\RES\INDEX\NwWeapons.dat"


    If NumNewWeapons > 0 Then
        ReDim nWeaponData(1 To NumNewWeapons)
        For i = 1 To NumNewWeapons

            With nWeaponData(i)

                .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
                .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
        
                For z = 1 To 4
        
                    .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
                Next z
                .Desc = GetVar(S, CStr(i), "Desc")
            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i
    End If
End Sub
Public Sub Load_NwMuniciones()

Dim S As String
Dim i As Long
Dim z As Long
S = App.Path & "\RES\INDEX\NwMuniciones.dat"


    If NumNewM > 0 Then
        ReDim nMunicionData(1 To NumNewM)
        For i = 1 To NumNewM

            With nMunicionData(i)

                .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
                .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
        
                For z = 1 To 4
        
                    .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
                Next z
                .Desc = GetVar(S, CStr(i), "Desc")
            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i

    End If
End Sub

Public Sub Load_NwCapas()
    Dim S As String
    Dim i As Long
    Dim z As Long
    Dim NumNewM As Integer
    S = App.Path & "\RES\INDEX\NwCapa.dat"

    If NumNewCapas > 0 Then
        ReDim nCapaData(1 To NumNewCapas)
        For i = 1 To NumNewCapas

            With nCapaData(i)

                .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
                .aOverWriteGrafico = Val(GetVar(S, CStr(i), "aOverWriteGrafico"))
                .pOverWriteGrafico = Val(GetVar(S, CStr(i), "pOverWriteGrafico"))
                For z = 1 To 4
        
                    .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))

                Next z
                .Desc = GetVar(S, CStr(i), "Desc")
            End With
            fCargando.PB.value = fCargando.PB.value + 1
            DoEvents
        Next i


    End If
End Sub
