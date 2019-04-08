Attribute VB_Name = "dxEngine"
Option Explicit
Public Zooming As Boolean
Public ZoomX As Integer
Public ZoomY As Integer
Public ZoomV As Integer
Public fX As Integer
Public fY As Integer
Public AcFrm As Single
Public sBody As Integer
Public sShield As Integer
Public sHead As Integer
Public sHelmet As Integer
Public sWeapon As Integer
Public sCapa As Integer
Public sMunicion As Integer



Public bBodyTest As Boolean
Public bHeadTest As Boolean
Public Pausa As Boolean
Public VerFrame As Boolean
Public LastCounter As Integer
Public VelMod As Single
Public anim_Intervalo As Single
Public anim_Counter As Long
Public Anim_Stoped As Boolean
Public Dibujo As Boolean
Public VelMov As Single
Public animCounter As Single
Public lColorFondo As Long
Public SurfaceDB8 As clsSurfaceManDynDX8
Public MainViewRect As RECT
Public MainBufferRect As RECT
Public ZoomBufferRect As RECT
Public dx As DirectX8
Public D3d As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public bNeglectNegro As Boolean
Private MainViewLeft As Integer
Private MainViewTop As Integer
Private TilePixelWidth As Integer
Private TilePixelHeight As Integer
Private WindowTileHeight As Integer
Private WindowTileWidth As Integer
Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer
Public ScreenWidth As Integer
Public ScreenHeight As Integer
Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer
Private TileBufferOffset As Integer
Private TileBufferSize As Integer
Public indexList(0 To 5)    As Integer
Public ibQuad               As DxVBLibA.Direct3DIndexBuffer8
Public vbQuadIdx            As DxVBLibA.Direct3DVertexBuffer8
Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public Type D3D8Textures
    texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Public acHeading As Byte

Dim dimeTex As Long
Dim Tex As Direct3DTexture8
Dim D3DbackBuffer As Direct3DSurface8
Dim zTarget As Direct3DSurface8
Dim stencil As Direct3DSurface8
Dim superTex As Direct3DSurface8
Public bRunning As Boolean
Private bump_map_supported As Boolean
Private Const engineBaseSpeed As Single = 0.02
Private FramesPerSecCounter As Integer
Private ScrollPixelsPerFrameX As Integer
Private ScrollPixelsPerFrameY As Integer
Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo ErrHandler:

    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB8 = New clsSurfaceManDynDX8
    
    Set dx = New DirectX8
    Set D3d = dx.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = fMain.ScaleWidth
        .BackBufferHeight = fMain.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = fMain.Frame1.hwnd
        
    End With
    DispMode.Format = D3DFMT_X8R8G8B8
    If D3d.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3d.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.Format = DispModeBK.Format
        End If
    Else
        bump_map_supported = False
        DispMode.Format = DispModeBK.Format
    End If
    Set D3DDevice = D3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, fMain.Frame1.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
      MainViewRect.top = 14
      MainViewRect.left = 4
      MainViewRect.Bottom = MainViewRect.top + 512
      MainViewRect.Right = MainViewRect.left + 512
      'lean1**
      MainBufferRect.top = 0
      MainBufferRect.left = 0
      MainBufferRect.Bottom = MainViewRect.Bottom - MainViewRect.top
      MainBufferRect.Right = MainViewRect.Right - MainViewRect.left
      
    '// Index Buffer
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
    
    Set ibQuad = D3DDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX32, D3DPOOL_SYSTEMMEM)
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
    
    Set vbQuadIdx = D3DDevice.CreateVertexBuffer(28 * 4, 0, FVF, D3DPOOL_SYSTEMMEM)

      
    MainViewTop = MainViewRect.top
    MainViewLeft = MainViewRect.left
    TilePixelWidth = 32
    TilePixelHeight = 32
    WindowTileHeight = ((MainViewRect.Bottom - MainViewRect.top) / 32)
    WindowTileWidth = ((MainViewRect.Right - MainViewRect.left) / 32)
    ScreenWidth = MainViewRect.Right - MainViewRect.left
    ScreenHeight = MainViewRect.Bottom - MainViewRect.top
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    TileBufferOffset = ((10 - TileBufferSize) * 32)
    D3DDevice.SetVertexShader FVF
    
    
D3DDevice.SetRenderState D3DRS_LIGHTING, False

'D3DDevice.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(150, 150, 150) 'The ambient value is a hex-RRGGBB code

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB8.Init(D3DX, D3DDevice, General_Get_Free_Ram_BytesEX)

    
   
    FPS = 60
    FramesPerSecCounter = 60
    
    ScrollPixelsPerFrameX = 9
    ScrollPixelsPerFrameY = 9

    

    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    'motion blur
    Set D3DbackBuffer = D3DDevice.GetRenderTarget
    Set zTarget = D3DDevice.GetDepthStencilSurface
    Set stencil = D3DDevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    Set Tex = D3DX.CreateTexture(D3DDevice, dimeTex, dimeTex, 1, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
    Set superTex = Tex.GetSurfaceLevel(0)


    
     

    
Exit Sub
ErrHandler:
bRunning = False
End Sub

Public Sub RenderMain()
On Error GoTo errr

Dim CC As Integer

1 D3DDevice.BeginScene

6 D3DDevice.Clear 1, MainBufferRect, D3DCLEAR_TARGET, lColorFondo, 1#, 0

Select Case UltimaOpcion

    Case eOpciones.Indices_op
    
    
        If IxS > 0 Then
        
            Draw_Index IxS, 0, 0, , bNeglectNegro


        
        End If

    Case eOpciones.Animaciones_op
        If IxS > 0 Then
        
           If Dibujo Then Draw_Anim IxS, animCounter, 250, 250, , ShouldAnim, , bNeglectNegro, VelMov

        End If
        
    Case eOpciones.Modeling_op
        'fMain.lblFrame.Caption = AcFrm
        'If Dibujo Then
         '       If sBody > 0 Then
          '          Draw_Anim nBodyData(sBody).mMovement(acHeading), AcFrm, 250, 250, nBodyData(sBody).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
          '          Cc = sBody
          '      ElseIf bBodyTest Then
          '          Draw_Anim nBodyData(num_test_body).mMovement(acHeading), AcFrm, 250, 250, nBodyData(num_test_body).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
          '          Cc = bBodyTest
          '      End If
          '      If sHead > 0 Then
          '          Draw_Index NHeadData(sHead).Frame(acHeading), 250, 250 + nBodyData(Cc).OffsetY + NHeadData(sHead).OffsetDibujoY, , bNeglectNegro
          '
          '      ElseIf bHeadTest Then
          '          Draw_Index NHeadData(num_test_head).Frame(acHeading), 250, 250 + nBodyData(num_test_body).OffsetY + NHeadData(num_test_head).OffsetDibujoY, , bNeglectNegro
          '      End If
          '      If sHelmet > 0 Then
          '          If acHeading = E_Heading.EAST Then
          '              Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 + NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(Cc).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
          '          ElseIf acHeading = E_Heading.WEST Then
          '              Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 - NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(Cc).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
          '
          '          Else
          '              Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250, 250 + nBodyData(Cc).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
          '          End If
          '      End If
          '      If sShield > 0 Then
          '          Draw_Anim nShieldDATA(sShield).mMovimiento(acHeading), AcFrm, 250, 250, nShieldDATA(sShield).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
          '
           '     End If
           '     If sWeapon > 0 Then
            '        Draw_Anim nWeaponData(sWeapon).mMovimiento(acHeading), AcFrm, 250, 250, nWeaponData(sWeapon).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov

           '     End If
          
        'End If
        RenderModeling
End Select
If Zooming Then
5     D3DDevice.Present ZoomBufferRect, MainViewRect, 0, ByVal 0
Else
3     D3DDevice.Present MainBufferRect, MainViewRect, 0, ByVal 0
End If
4 D3DDevice.EndScene

Exit Sub
errr:

Debug.Print Err.Description & "_" & Erl
D3DDevice.EndScene
End Sub

Public Sub RenderModeling()
Dim CC As Integer

        fMain.lblFrame.Caption = AcFrm
        If Dibujo Then
        Select Case acHeading
            Case 1 'N E S W

                If sWeapon > 0 Then
                    Draw_Anim nWeaponData(sWeapon).mMovimiento(acHeading), AcFrm, 250, 250, nWeaponData(sWeapon).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov

                End If
                If sShield > 0 Then
                    Draw_Anim nShieldDATA(sShield).mMovimiento(acHeading), AcFrm, 250, 250, nShieldDATA(sShield).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                End If
                            If sBody > 0 Then
                    Draw_Anim nBodyData(sBody).mMovement(acHeading), AcFrm, 250, 250, nBodyData(sBody).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = sBody
                ElseIf bBodyTest Then
                    Draw_Anim nBodyData(num_test_body).mMovement(acHeading), AcFrm, 250, 250, nBodyData(num_test_body).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = bBodyTest
                End If
                If sHead > 0 Then
                    Draw_Index NHeadData(sHead).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHeadData(sHead).OffsetDibujoY, , bNeglectNegro
                
                ElseIf bHeadTest Then
                    Draw_Index NHeadData(num_test_head).Frame(acHeading), 250, 250 + nBodyData(num_test_body).OffsetY + NHeadData(num_test_head).OffsetDibujoY, , bNeglectNegro
                End If
                If sHelmet > 0 Then
                    If acHeading = E_Heading.EAST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 + NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    ElseIf acHeading = E_Heading.WEST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 - NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    
                    Else
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    End If
                End If
                If sMunicion > 0 Then
                
                    Draw_Anim nMunicionData(sMunicion).mMovimiento(acHeading), AcFrm, 250, 250, nMunicionData(sMunicion).OverWriteGrafico, ShouldAnim, nMunicionData(sMunicion).Alpha, bNeglectNegro, VelMov
                End If
            Case 2

                If sShield > 0 Then
                    Draw_Anim nShieldDATA(sShield).mMovimiento(acHeading), AcFrm, 250, 250, nShieldDATA(sShield).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                End If
                            If sBody > 0 Then
                    Draw_Anim nBodyData(sBody).mMovement(acHeading), AcFrm, 250, 250, nBodyData(sBody).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = sBody
                ElseIf bBodyTest Then
                    Draw_Anim nBodyData(num_test_body).mMovement(acHeading), AcFrm, 250, 250, nBodyData(num_test_body).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = bBodyTest
                End If
                If sHead > 0 Then
                    Draw_Index NHeadData(sHead).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHeadData(sHead).OffsetDibujoY, , bNeglectNegro
                
                ElseIf bHeadTest Then
                    Draw_Index NHeadData(num_test_head).Frame(acHeading), 250, 250 + nBodyData(num_test_body).OffsetY + NHeadData(num_test_head).OffsetDibujoY, , bNeglectNegro
                End If
                If sMunicion > 0 Then
                
                    Draw_Anim nMunicionData(sMunicion).mMovimiento(acHeading), AcFrm, 250, 250, nMunicionData(sMunicion).OverWriteGrafico, ShouldAnim, nMunicionData(sMunicion).Alpha, bNeglectNegro, VelMov
                End If
                If sHelmet > 0 Then
                    If acHeading = E_Heading.EAST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 + NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    ElseIf acHeading = E_Heading.WEST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 - NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    
                    Else
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    End If
                
                End If
                If sWeapon > 0 Then
                    Draw_Anim nWeaponData(sWeapon).mMovimiento(acHeading), AcFrm, 250, 250, nWeaponData(sWeapon).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov

                End If
            Case 3
            
                If sMunicion > 0 Then
                
                    Draw_Anim nMunicionData(sMunicion).mMovimiento(acHeading), AcFrm, 250, 250, nMunicionData(sMunicion).OverWriteGrafico, ShouldAnim, nMunicionData(sMunicion).Alpha, bNeglectNegro, VelMov
                End If
    If sBody > 0 Then
                    Draw_Anim nBodyData(sBody).mMovement(acHeading), AcFrm, 250, 250, nBodyData(sBody).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = sBody
                ElseIf bBodyTest Then
                    Draw_Anim nBodyData(num_test_body).mMovement(acHeading), AcFrm, 250, 250, nBodyData(num_test_body).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = bBodyTest
                End If
                If sHead > 0 Then
                    Draw_Index NHeadData(sHead).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHeadData(sHead).OffsetDibujoY, , bNeglectNegro
                
                ElseIf bHeadTest Then
                    Draw_Index NHeadData(num_test_head).Frame(acHeading), 250, 250 + nBodyData(num_test_body).OffsetY + NHeadData(num_test_head).OffsetDibujoY, , bNeglectNegro
                End If
                If sHelmet > 0 Then
                    If acHeading = E_Heading.EAST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 + NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    ElseIf acHeading = E_Heading.WEST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 - NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    
                    Else
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    End If
                End If
                If sWeapon > 0 Then
                    Draw_Anim nWeaponData(sWeapon).mMovimiento(acHeading), AcFrm, 250, 250, nWeaponData(sWeapon).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov

                End If
                If sShield > 0 Then
                    Draw_Anim nShieldDATA(sShield).mMovimiento(acHeading), AcFrm, 250, 250, nShieldDATA(sShield).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                End If
                        
            Case 4
                If sWeapon > 0 Then
                    Draw_Anim nWeaponData(sWeapon).mMovimiento(acHeading), AcFrm, 250, 250, nWeaponData(sWeapon).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov

                End If
                            If sBody > 0 Then
                    Draw_Anim nBodyData(sBody).mMovement(acHeading), AcFrm, 250, 250, nBodyData(sBody).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = sBody
                ElseIf bBodyTest Then
                    Draw_Anim nBodyData(num_test_body).mMovement(acHeading), AcFrm, 250, 250, nBodyData(num_test_body).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                    CC = bBodyTest
                End If
                If sHead > 0 Then
                    Draw_Index NHeadData(sHead).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHeadData(sHead).OffsetDibujoY, , bNeglectNegro
                
                ElseIf bHeadTest Then
                    Draw_Index NHeadData(num_test_head).Frame(acHeading), 250, 250 + nBodyData(num_test_body).OffsetY + NHeadData(num_test_head).OffsetDibujoY, , bNeglectNegro
                End If
                If sMunicion > 0 Then
                
                    Draw_Anim nMunicionData(sMunicion).mMovimiento(acHeading), AcFrm, 250, 250, nMunicionData(sMunicion).OverWriteGrafico, ShouldAnim, nMunicionData(sMunicion).Alpha, bNeglectNegro, VelMov
                End If
                If sHelmet > 0 Then
                    If acHeading = E_Heading.EAST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 + NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    ElseIf acHeading = E_Heading.WEST Then
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250 - NHelmetData(sHelmet).OffsetLat, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    
                    Else
                        Draw_Index NHelmetData(sHelmet).Frame(acHeading), 250, 250 + nBodyData(CC).OffsetY + NHelmetData(sHelmet).OffsetDibujoY, NHelmetData(sHelmet).Alpha, bNeglectNegro
                    End If
                
                End If
                If sShield > 0 Then
                    Draw_Anim nShieldDATA(sShield).mMovimiento(acHeading), AcFrm, 250, 250, nShieldDATA(sShield).OverWriteGrafico, ShouldAnim, , bNeglectNegro, VelMov
                End If
            
        End Select
    End If
    
End Sub
Private Function Engine_FToDW(F As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, F
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function
Private Sub Draw_Index(ByVal nIndex As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Alpha_Blend As Byte = 0, Optional ByVal NeglectNegro As Boolean = False)
    Dim d3dtextures As D3D8Textures
    Dim z As Long
    Dim light_value(0 To 3) As Long
    Dim verts(3) As TLVERTEX
    
    Set d3dtextures.texture = SurfaceDB8.GetTexture(NewIndexData(nIndex).OverWriteGrafico, d3dtextures.texwidth, d3dtextures.texheight)
    D3DDevice.SetTexture 0, d3dtextures.texture
    

    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    
    With EstaticData(NewIndexData(nIndex).Estatic)
    jx = .L
    jy = .T
    jw = .W
    jh = .H
    If .tw <> 1 Then
        z = -.tw * 16 + 16
        X = X + z
    End If
    
    If .th <> 1 Then
        z = -.th * 32 + 32
        Y = Y + z
    End If
    
    
    End With
    
    
    If d3dtextures.texwidth = 0 Or d3dtextures.texheight = 0 Then Exit Sub
    
        With verts(2)
            .X = X
            .Y = Y + jh
            .tu = jx / (d3dtextures.texwidth)
            .tv = (jy + jh) / (d3dtextures.texheight)
            .rhw = 1
            .color = -1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = jx / (d3dtextures.texwidth)
            .tv = jy / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        
        With verts(3)
            .X = X + jw
            .Y = Y + jh
            .tu = (jx + jw) / (d3dtextures.texwidth)
            .tv = (jy + jh) / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        
        With verts(1)
            .X = X + jw
            .Y = Y
            .tu = (jx + jw) / (d3dtextures.texwidth)
            .tv = jy / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        If Alpha_Blend > 0 Or NeglectNegro Then

            If Not NeglectNegro Then
                D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                'D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                verts(0).color = D3DColorARGB(Alpha_Blend, 255, 255, 255)
                verts(1).color = D3DColorARGB(Alpha_Blend, 255, 255, 255)
                verts(2).color = D3DColorARGB(Alpha_Blend, 255, 255, 255)
                verts(3).color = D3DColorARGB(Alpha_Blend, 255, 255, 255)
                
    
            Else
                D3DDevice.SetRenderState D3DRS_DESTBLEND, 2
                D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            End If

        End If
    
        'D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, verts(0), 28
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28

    
        If Alpha_Blend > 0 Or NeglectNegro Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTOP_SELECTARG1)
            Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTOP_DISABLE)
        
            
        End If
End Sub
Private Sub Draw_Anim(ByVal nIndex As Integer, ByRef Counter As Single, ByVal X As Integer, ByVal Y As Integer, Optional ByVal oGrafico As Integer = -1, Optional ByVal Animar As Boolean = True, Optional ByVal Alpha_Blend As Byte = 0, Optional ByVal NeglectNegro As Boolean = False, Optional ByVal Vel As Single)
On Error GoTo erdrawanim
    Dim d3dtextures As D3D8Textures
    Dim light_value(0 To 3) As Long
    Dim verts(3) As TLVERTEX
    Dim i As Integer
    Dim Grf As Integer
    Dim z As Long
    
    With NewAnimationData(nIndex)
        If Animar And .NumFrames > 1 And .Filas <> 0 And .Columnas <> 0 Then
            If Vel <> 0 Then
                Counter = Counter + ((MEE * (0.02 * VelMod)) * (.NumFrames / (20 - Vel)) * .Velocidad)
            Else
                Counter = Counter + ((MEE * (0.002 * VelMod)) * .NumFrames * .Velocidad)
            End If
            If Counter > .NumFrames Then
                If anim_Intervalo > 0 Then
                    Anim_Stoped = True
                    Dibujo = False
                    Counter = 1
                    Exit Sub
                Else
                    Counter = Counter Mod .NumFrames
                End If
            End If
            If Counter < 1 Then Counter = 1
            i = Counter
        Else
            If .NumFrames = 0 Or .Filas = 0 Or .Columnas = 0 Then Exit Sub
            If Counter > .NumFrames Or Counter < 1 Then Counter = 1
            
            i = Counter

        End If

        
        
        If VerFrame Then
        If i <> LastCounter Then
                fMain.lblfr.Caption = "FR: " & i
                LastCounter = i
        End If
        End If
        If oGrafico > 0 Then
            Grf = (.Indice(i).Grafico - .Indice(1).Grafico) + oGrafico
        Else
            Grf = .Indice(i).Grafico
        End If
        
    
    Set d3dtextures.texture = SurfaceDB8.GetTexture(Grf, d3dtextures.texwidth, d3dtextures.texheight)
    D3DDevice.SetTexture 0, d3dtextures.texture
    
    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    
    jx = .Indice(i).X
    jy = .Indice(i).Y
    jh = .Height
    jw = .Width
    
    If .TileWidth <> 1 Then
        z = -.TileWidth * 16 + 16
        X = X + z
    End If
    
    If .TileHeight <> 1 Then
        z = -.TileHeight * 32 + 32
        Y = Y + z
    End If

    
    End With
    
    If d3dtextures.texwidth = 0 Or d3dtextures.texheight = 0 Then Exit Sub
    
        With verts(2)
            .X = X
            .Y = Y + jh
            .tu = jx / (d3dtextures.texwidth - 1)
            .tv = (jy + jh) / (d3dtextures.texheight - 1)
            .rhw = 1
            .color = -1
            If .tu > 1 Then .tu = 1
            If .tv > 1 Then .tv = 1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = jx / (d3dtextures.texwidth - 1)
            .tv = jy / (d3dtextures.texheight - 1)
            .rhw = 1
            .color = -1
            If .tu > 1 Then .tu = 1
            If .tv > 1 Then .tv = 1
        End With
        
        With verts(3)
            .X = X + jw
            .Y = Y + jh
            .tu = (jx + jw) / (d3dtextures.texwidth - 1)
            .tv = (jy + jh) / (d3dtextures.texheight - 1)
            .rhw = 1
            .color = -1
            If .tu > 1 Then .tu = 1
            If .tv > 1 Then .tv = 1
        End With
        
        With verts(1)
            .X = X + jw
            .Y = Y
            .tu = (jx + jw) / (d3dtextures.texwidth - 1)
            .tv = jy / (d3dtextures.texheight - 1)
            .rhw = 1
            .color = -1
            If .tu > 1 Then .tu = 1
            If .tv > 1 Then .tv = 1
        End With
        If Alpha_Blend > 0 Or NeglectNegro Then

            If Not NeglectNegro Then
                D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            Else
                D3DDevice.SetRenderState D3DRS_DESTBLEND, 2
                D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            End If

        End If
    
        D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, verts(0), 28


    
        If Alpha_Blend > 0 Or NeglectNegro Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTOP_SELECTARG1)
            Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTOP_DISABLE)
        
            
        End If
        Exit Sub
erdrawanim:


End Sub
Public Function ShouldAnim() As Boolean

If Pausa Then
    ShouldAnim = False
    Exit Function
End If
If fMain.Check1.value = 0 Then
    Exit Function
End If

ShouldAnim = True

End Function
Public Sub DibujareEnHwnd3(ByVal PIC As Long, ByVal Graf As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal PRESENTO As Boolean, ByRef W As Integer, ByRef H As Integer)

Dim DestRect As RECT
Dim tX As Byte
Dim tY As Byte
Dim src_rect As RECT
Dim d3dtextures As D3D8Textures
Dim light_value(0 To 3) As Long
Dim verts(3) As TLVERTEX

Set d3dtextures.texture = SurfaceDB8.GetTexture(Graf, d3dtextures.texwidth, d3dtextures.texheight)
D3DDevice.SetTexture 0, d3dtextures.texture
W = d3dtextures.texwidth
H = d3dtextures.texheight


  X = X
  Y = Y

   src_rect.top = 0
   src_rect.left = 0
   src_rect.Right = d3dtextures.texwidth
   src_rect.Bottom = d3dtextures.texheight
    
    
   DestRect.top = Y
   DestRect.left = X
   DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
   DestRect.Right = X + src_rect.Right - src_rect.left
   If src_rect.Bottom <= 0 Or src_rect.Right <= 0 Or src_rect.left = src_rect.Right Or src_rect.top = src_rect.Bottom Then Exit Sub
   
   
   D3DDevice.Clear 1, DestRect, D3DCLEAR_TARGET, &H0, ByVal 0, 0
   D3DDevice.BeginScene



        With verts(2)
            .X = X
            .Y = Y + d3dtextures.texheight
            .tu = 0 / (d3dtextures.texwidth)
            .tv = (0 + d3dtextures.texheight) / (d3dtextures.texheight)
            .rhw = 1
            .color = -1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = 0 / (d3dtextures.texwidth)
            .tv = 0 / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        
        With verts(3)
            .X = X + d3dtextures.texwidth
            .Y = Y + d3dtextures.texheight
            .tu = (0 + d3dtextures.texwidth) / (d3dtextures.texwidth)
            .tv = (0 + d3dtextures.texheight) / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        
        With verts(1)
            .X = X + d3dtextures.texwidth
            .Y = Y
            .tu = (0 + d3dtextures.texwidth) / (d3dtextures.texwidth)
            .tv = 0 / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
    
   

  
  

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28



   D3DDevice.EndScene
   

   
   If PRESENTO Then D3DDevice.Present src_rect, DestRect, PIC, ByVal 0


End Sub
Public Sub DibujareEnHwndIndex(ByVal PIC As Long, ByVal Grafico As Integer, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal PRESENTO As Boolean, ByRef Wi As Integer, ByRef Hi As Integer)

Dim DestRect As RECT
Dim tX As Byte
Dim tY As Byte
Dim src_rect As RECT
Dim d3dtextures As D3D8Textures
Dim light_value(0 To 3) As Long
Dim verts(3) As TLVERTEX

Set d3dtextures.texture = SurfaceDB8.GetTexture(Grafico, d3dtextures.texwidth, d3dtextures.texheight)
D3DDevice.SetTexture 0, d3dtextures.texture
Wi = d3dtextures.texwidth
Hi = d3dtextures.texheight


  X = X
  Y = Y

   src_rect.top = T
   src_rect.left = L
   src_rect.Right = L + W
   src_rect.Bottom = T + H
    
    
   DestRect.top = Y
   DestRect.left = X
   DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
   DestRect.Right = X + src_rect.Right - src_rect.left
   If src_rect.Bottom <= 0 Or src_rect.Right <= 0 Or src_rect.left = src_rect.Right Or src_rect.top = src_rect.Bottom Then Exit Sub
   
   
   D3DDevice.Clear 1, DestRect, D3DCLEAR_TARGET, &H0, ByVal 0, 0
   D3DDevice.BeginScene



        With verts(2)
            .X = X
            .Y = Y + W
            .tu = L / (d3dtextures.texwidth)
            .tv = (T + H) / (d3dtextures.texheight)
            .rhw = 1
            .color = -1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = L / (d3dtextures.texwidth)
            .tv = T / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        
        With verts(3)
            .X = X + W
            .Y = Y + H
            .tu = (L + W) / (d3dtextures.texwidth)
            .tv = (T + H) / (d3dtextures.texheight)
            .rhw = 1
            .color = -1

        End With
        
        With verts(1)
            .X = X + W
            .Y = Y
            .tu = (L + W) / (d3dtextures.texwidth)
            .tv = 0 / (H)
            .rhw = 1
            .color = -1

        End With
    
   

  
  

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28



   D3DDevice.EndScene
   

   
   If PRESENTO Then D3DDevice.Present src_rect, DestRect, PIC, ByVal 0


End Sub
