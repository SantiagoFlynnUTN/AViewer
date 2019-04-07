Attribute VB_Name = "ModMain"
Option Explicit
Public Enum eOpciones
    Indices_op
    Animaciones_op
    Modeling_op
    Particulas_op
    Fx_op
    Meditaciones_op
End Enum
Public Modeling_Type As Byte
Public LastCheckKeys As Long
Public Enum E_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum
Public IxS As Integer
Public Movement_Engine_Counter As Long
Public ShowFPS As Long
Public MEE As Long ' Main Timer
Public UltimaOpcion As Integer
Public FpsLastCheck As Long
Public FPS As Integer

Public Sub CerrarOpcion(ByVal Index As Integer)
    IxS = 0
    With fMain
        .cOpc(Index).value = False
        .Lista(Index).Visible = False
    End With
    Select Case Index
    
        Case eOpciones.Animaciones_op
            fMain.CmdSAve.Visible = False
            fMain.frAnim.Visible = False
        Case eOpciones.Modeling_op
            fMain.fModeling.Visible = False
            fMain.cBody.Clear
            fMain.cCabeza.Clear
            fMain.cCasco.Clear
            fMain.cEscudo.Clear
            fMain.cArma.Clear
            fMain.fCabezas.Visible = False
            
            
    End Select
End Sub
Public Sub AbrirOpcion(ByVal Index As Integer)
    Dim i As Long
    With fMain
        .Lista(Index).Visible = True
    
    
    
    
        Select Case Index
            Case eOpciones.Indices_op
            Case eOpciones.Animaciones_op
                .CmdSAve.Visible = True
                .frAnim.Visible = True
            Case eOpciones.Modeling_op
                .fModeling.Visible = True
                .cCasco.AddItem "Ningun Casco"
                If Num_Helmets > 0 Then
                    For i = 1 To Num_Helmets
                        fMain.cCasco.AddItem NHelmetData(i).Desc & "(" & i & ")"
                    Next i
                End If
                .cCabeza.AddItem "Ninguna Cabeza"
                If Num_Heads > 0 Then
                    For i = 1 To Num_Heads
                        fMain.cCabeza.AddItem NHeadData(i).Desc & "(" & i & ")"
                    Next i
                End If
                .cBody.AddItem "Ningun Cuerpo"
                If NumNewBodys > 0 Then
                    For i = 1 To NumNewBodys
                        fMain.cBody.AddItem nBodyData(i).Desc & "(" & i & ")"
                    Next i
                End If
                .cArma.AddItem "Ningun Arma"
                If NumNewWeapons > 0 Then
                    For i = 1 To NumNewWeapons
                        fMain.cArma.AddItem nWeaponData(i).Desc & "(" & i & ")"
                    Next i
            
                End If
                .cEscudo.AddItem "Ningun Escudo"
                If NumNewShields > 0 Then
                    For i = 1 To NumNewShields
                        fMain.cEscudo.AddItem nShieldDATA(i).Desc & "(" & i & ")"
                    Next i
                End If
            Case eOpciones.Particulas_op
            Case eOpciones.Fx_op
            Case eOpciones.Meditaciones_op

            

                
        End Select
    End With
    
    

End Sub


Sub Main()

    fCargando.Show
    DoEvents

    'Cargamos la data que define el max del pb.
    Num_Heads = Val(GetVar(App.Path & "\RES\INDEX\NewHeads.dat", "INIT", "NUM"))
    Num_Helmets = Val(GetVar(App.Path & "\RES\INDEX\NewHelmets.dat", "INIT", "NUM"))
    numNewIndex = Val(GetVar(App.Path & "\RES\INDEX\NewIndex.dat", "INIT", "num"))
    numNewEstatic = Val(GetVar(App.Path & "\RES\INDEX\NewEstatics.dat", "INIT", "num"))
    Num_NwAnim = Val(GetVar(App.Path & "\RES\INDEX\NewAnim.dat", "NW_ANIM", "NUM"))
    NumNewBodys = Val(GetVar(App.Path & "\RES\INDEX\NewBody.dat", "INIT", "num"))
    NumNewShields = Val(GetVar(App.Path & "\RES\INDEX\Nwshields.dat", "INIT", "num"))
    NumNewWeapons = Val(GetVar(App.Path & "\RES\INDEX\NwWeapons.dat", "INIT", "num"))
    NumNewM = Val(GetVar(App.Path & "\RES\INDEX\NwMunicion.dat", "INIT", "num"))
    NumNewCapas = Val(GetVar(App.Path & "\RES\INDEX\NwCapa.dat", "INIT", "num"))

    fCargando.PB.Max = 5 + numNewIndex + numNewEstatic + Num_NwAnim + Num_Heads + Num_Helmets + NumNewShields + NumNewWeapons + NumNewM + NumNewCapas + NumNewBodys
    '/

    fCargando.lblestado.Caption = "Iniciando Engine..."
    DoEvents
    dxEngine.Engine_Init
    fCargando.PB = fCargando.PB + 5
    DoEvents

    fCargando.lblestado.Caption = "Cargando Estaticos..."
    DoEvents
    ModData.Load_NewEstatics

    fCargando.lblestado.Caption = "Cargando Indices..."
    DoEvents
    ModData.Load_NewIndex
    DoEvents

    fCargando.lblestado.Caption = "Cargando animaciones..."
    DoEvents
    ModData.Load_NewAnimation
    DoEvents

    fCargando.lblestado.Caption = "Cargando Heads..."
    DoEvents
    ModData.Load_NewHeads
    DoEvents

    fCargando.lblestado.Caption = "Cargando cuerpos..."
    DoEvents
    ModData.Load_NewBodys
    DoEvents

    fCargando.lblestado.Caption = "Cargando cascos..."
    DoEvents
    ModData.Load_NewHelmets
    DoEvents

    fCargando.lblestado.Caption = "Cargando escudos..."
    DoEvents
    ModData.Load_NewShields
    DoEvents

    fCargando.lblestado.Caption = "Cargando armas..."
    DoEvents
    ModData.Load_NewWeapons
    DoEvents

    fCargando.lblestado.Caption = "Cargando municiones..."
    DoEvents
    ModData.Load_NwMuniciones
    DoEvents

    fCargando.lblestado.Caption = "Cargando capas..."
    DoEvents
    ModData.Load_NwCapas
    DoEvents
    num_test_body = Test_body_def
    fMain.m_c_elegir.Caption = "Cuerpo Prueba: " & num_test_body
    num_test_head = Test_Head_Def
    fMain.m_z_elegir.Caption = "Cabeza Prueba: " & num_test_head
    'Terminamos la carga.
    fCargando.Timer1.Enabled = True

    Do Until bRunning
        DoEvents
    Loop
    Dibujo = True
    VelMod = 1
    acHeading = 3
    fMain.Show
    Unload fCargando
    MainLoop
End Sub

Public Sub MainLoop()

    Do Until bRunning = False
        
        DoEvents
        
        dxEngine.RenderMain
        
        DoEvents
        
        
        If anim_Intervalo > 0 And Anim_Stoped Then
            If CumplioIntervalo(anim_Counter, anim_Intervalo * 1000, True) Then
                Dibujo = True
                Anim_Stoped = False
            End If
        End If
        
        MEE = PasoTiempo(Movement_Engine_Counter)
        If CumplioIntervalo(FpsLastCheck, 1000, True) Then
            fMain.lblfps.Caption = "FPS : " & FPS
            FPS = 0
        End If
        If CumplioIntervalo(LastCheckKeys, 150, True) Then
            CheckKeys
        End If
        FPS = FPS + 1
        
    Loop


    CerrarPrograma
    End
    
End Sub
Public Sub CerrarPrograma()

    Set dx = Nothing
    Set D3DDevice = Nothing
    Set D3DX = Nothing
    Set D3d = Nothing

End Sub

Public Function GetRaza(ByVal i As Byte) As String

    Select Case i
        Case 0
            GetRaza = "Fantasma"
        Case 1
            GetRaza = "Human"
        Case 2
            GetRaza = "Elf"
        Case 3
            GetRaza = "Elfo Oscur"
        Case 5
            GetRaza = "Enan"
        Case 4
            GetRaza = "Gnom"
    End Select
End Function
Public Function GetLastLetter(ByVal i As Byte) As String

    If i = 1 Then
        GetLastLetter = "o"
    ElseIf i = 2 Then
        GetLastLetter = "a"
    End If

    
End Function
Public Sub CheckKeys()

    Select Case UltimaOpcion
    End Select
End Sub
