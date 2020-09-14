Attribute VB_Name = "Module1"
Public Const STAT_MAXELV As Byte = 60
Public Const MAX_HABILIDADES As Byte = 48
Public Num_Med As Integer
Public NumHerr As Byte
Public NumCarp As Byte
Public NumSastr As Byte
Public Type tCrafts
    Tipo As Byte
    Mat1 As Integer
    Mat2 As Integer
    Mat3 As Integer
    ProfesionNivel As Byte
    Item As Integer
    Version As Byte
End Type
Public numero_buffs As Byte
Public Type tBuffData
    Duracion As Integer
    Intervalo As Integer
    Nombre As String
    dFX As Byte
    dEfecto As Byte
    Tipo As Byte
    grhindex As Integer
End Type
Public buff_data() As tBuffData

Public AuraDATA() As tAura
Public nAura As Integer
Public Type tAura
    grhindex As Integer
    r As Byte
    g As Byte
    b As Byte
    A As Byte
    OffsetX As Integer
    OffsetY As Integer
    Giratoria As Byte
    Velocidad As Single
    Tipo As Byte
End Type
Public TotalStreams As Integer
Public StreamData() As Stream
Public SPOTLIGHTS_COLORES() As Long
Public NUM_SPOTLIGHTS_COLORES As Byte

Public NUM_SPOTLIGHTS_ANIMATION As Byte
Public SPOTLIGHTS_ANIMATION() As Integer
'RGB Type
Private Type RGB
    r As Long
    g As Long
    b As Long
End Type
 
Private Type Stream
    Name As String
    NumOfParticles As Long
    NumTrueParticles As Long
    NumGrhs As Long
    id As Long
    X1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    Spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Public cHerreria() As tCrafts
Public cSastreria() As tCrafts
Public cCarpinteria() As tCrafts

Public Num_Fx As Integer
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpFileName As String) As Long
                                                   
Public NumCanje As Integer
Private Type tCanje
    Nombre As String
    Info As String
    Tipo As Byte
    nItems As Byte
    Items() As Integer
    Cant() As Integer
    vGema As Integer
    vMM As Integer
    Descuento As Byte
    Version As Byte
End Type
Public Canjes() As tCanje


                                                   
Public Type tDecorDrop
   
   ObjIndex    As Integer
   DropChance  As Long
   
End Type
Public Type tNewIndex
    Estatic As Integer ' Info de estatica
    Dinamica As Integer ' Animacion
    OverWriteGrafico As Integer ' Grafico
End Type
Public Type tNewHelmet
    Alpha As Byte
    mMovimiento(1 To 4) As Integer
    OffsetY As Integer
    OffsetLat As Integer
End Type


Public Type tnHead
    Frame(1 To 4) As Integer
    OffsetDibujoY As Integer
    OffsetOjos As Integer
    Raza As Byte
    Genero As Byte
End Type
Public NHeadData() As tnHead

Public Type tNewEstatic
    w As Integer
    h As Integer
    L As Integer
    t As Integer
    TW As Single
    TH As Single
End Type
Public NumEstatics As Integer
Public EstaticData() As tNewEstatic
Public NewIndexData() As tNewIndex
Public NumNewIndex As Integer

Public Type tNewFX
    Animacion As Integer
    OffsetY As Integer
    OffsetX As Integer
    Alpha As Byte
    Color As Long
    Rombo As Byte
    Particula As Byte
    Life As Integer
    AnimInicial As Integer
    AnimFinal As Integer
    ParaleloInicial As Integer
    ParaleloStart As Byte
End Type
Public Type tNewShieldData
    mMovimiento(1 To 4) As Integer
    Alpha As Byte
    OverWriteGrafico As Integer
End Type
Public Type tNewWeaponData
    mMovimiento(1 To 4) As Integer
    Alpha As Byte
    OverWriteGrafico As Integer
End Type
Public nShieldDATA() As tNewShieldData
Public nWeaponData() As tNewWeaponData

Public FxData() As tNewFX
Public MedData() As tNewFX


Public Type tFonts
    lStart As Long
    lSize As Long
    nT As Integer
    Data() As Byte
End Type
Public nFonts As Integer
Public Fonts() As tFonts

Public Type tNewBodyData
    mMovement(1 To 4) As Integer
    Reposo(1 To 4) As Integer
    Attack(1 To 4) As Integer
    Death(1 To 4) As Integer
    Attacked(1 To 4) As Integer
    OverWriteGrafico As Integer
    bAtacado As Boolean
    bReposo As Boolean
    bAtaque As Boolean
    bDeath As Boolean
    bContinuo As Boolean
    OffsetY As Integer
    Capa As Integer
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
ByRef Destination As Any, ByRef source As Any, ByVal numbytes As Long)
Public Type tNewIndice
    X As Integer
    Y As Integer
    Grafico As Integer
End Type
Private nBodyData() As tNewBodyData
Public Type tNewAnimation
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
    TipoAnimacion As Byte
    
    Romboidal As Byte
    Direction As Integer
    
    OffsetX As Integer
    OffsetY As Integer
    Initial As Integer
End Type
Public NewAnimationData() As tNewAnimation
Private Type tArchivoI
    Le As Long
    St As Long
    Da() As Byte
End Type

Public IXAR() As tArchivoI


Public Type tpos
    X As Integer
    Y As Integer
End Type

Public MapaData() As tpos
Public NroMapas As Integer


Public Type thab
   Grafico As Integer
End Type

Public num_npcs_h As Integer
Public num_npcs_nh As Integer
Public Habilidades(1 To MAX_HABILIDADES) As thab

Public Type tNpcHostile
    Body As Integer
    Head As Integer
    MAX_HP As Long
    Nombre As String
    Snd1 As Byte
    Snd2 As Byte
    
End Type

Public Type tQuests
    Nombre As String
    Tipo As Byte
    nTargets As Byte
    Targets() As Integer
    TargetsCant() As Integer
    Oro As Long
    Exp As Long
    Item() As Integer
    Cant() As Integer
    numritems As Byte
    Desc As String
    TipoReco As Byte
    Puntos As Long
    
    
End Type
Public Quests() As tQuests
Public nQuest As Integer

Public hostiles() As tNpcHostile

Public Type tNpcNoHostile
    Body As Integer
    Head As Integer
    MAX_HP As Long
    NPCTYPE As Byte
    Nombre As String
    Desc As String
End Type

Public nHostiles() As tNpcNoHostile


Public Type sd
    Nombre As String
    fx As Integer
    loops As Integer
    Sound As Byte
    Tipo As Byte
    Manareq As Integer
    Skills As Byte
    Libro As Integer
    
    Info As String
    castermsg As String
    magicwords As String
    propiomsg As String
    targetmsg As String
    CasterFx As Integer
    CasterLoop As Integer
    
End Type

Public Spells() As sd
Public ns As Byte

Private Type tGraphicButtonData
    size As RECT
        
    NormalSurfaceNum As Integer
    SelectSurfaceNum As Integer
    PressSurfaceNum As Integer
    
    Normal_Rojo As Byte
    Normal_Verde As Byte
    Normal_Azul As Byte
    
    Sel_Rojo As Byte
    Sel_Verde As Byte
    Sel_Azul As Byte
    
    Press_Rojo As Byte
    Press_Verde As Byte
    Press_Azul As Byte
    
    
    HandIco As Boolean
    Sound As Boolean
    Caption As Integer
End Type

Private Type tGraphicTextBox
    SurfaceNum As Integer
    X As Integer
    Y As Integer
    OffsetX As Integer
    OffsetY As Integer
    texto As String
    fColor As Long
    Selecto As Boolean
    Muestro As Boolean
    Centrar As Boolean
    TipoTexto As Byte
    w As Integer
    h As Integer
End Type

Private Type tGraphicCheckData
    Tipo_Check As Byte '1 TICK, 2 NUMERIC
    X As Integer
    Y As Integer
    w As Integer
    h As Integer
    Caption As Byte
    ColorR As Byte
    ColorG As Byte
    ColorB As Byte
    ColorA As Byte
    SurfaceNum As Long
    CheckSurface As Long
    IniciaVisible As Byte
    min As Byte
    max As Byte
End Type


Private Type tGraphicTextBoxData
    SurfaceNum As Integer
    X As Integer
    Y As Integer
    OffsetX As Integer
    OffsetY As Integer
    fR As Byte
    fG As Byte
    fb As Byte
    fA As Byte
    IniciaVisible As Byte
    Centrar As Byte
        w As Integer
    h As Integer
    TipoTexto As Byte
End Type


Private Type tGraphicTextosData
    texto As Integer
    X As Integer
    Y As Integer
    r As Byte
    g As Byte
    b As Byte
    A As Byte
    IniciaVisible As Byte
    Centrar As Byte
    MaxWidth As Integer
End Type

Private Type tGraphicFormData
    TieneSpecial As Boolean
    SurfaceNum As Integer
    GraficosX As Byte
    GraficosY As Byte
    Buttons() As tGraphicButtonData
    num_Buttons As Byte
    ScreenX As Integer
    ScreenY As Integer
    Draw_Stage As Byte
    AlphaValue As Byte
    Num_Textos As Byte
    Textos() As tGraphicTextosData
    Num_TextBox As Byte
    TextBox() As tGraphicTextBoxData
    Width As Integer
    Height As Integer
    Num_Checks As Byte
    Checks() As tGraphicCheckData
End Type

Public FD() As tGraphicFormData
Public NUM_FD As Byte

Public Type tChirimbolo_Data
    Graf_Index As Integer
    Tipo As Byte  ' 0 Desaparece cuando se le avisa, 1 desaparece por conteo, 2 desaparece por accion
    Tiempo As Integer
End Type

Public Type tNobleItem
    Numero As Integer
    NumItems_Requeridos As Byte
    Items_Requeridos() As Integer
    cantItems_Requeridos() As Integer
End Type

Public Type tNobledata
    NumItems As Byte
    Items() As tNobleItem
End Type
Public Nobleza_Data As tNobledata


Public NHelmetData() As tNewHelmet
Public NumNewHelmet As Integer
Public Type tNewMunicionData
    mMovimiento(1 To 4) As Integer
    Alpha As Byte
    OverWriteGrafico As Integer
End Type
Public Type tNewCapa
    mMovimiento(1 To 4) As Integer
    
    Alpha As Byte
    
    aOverWriteGrafico As Integer
    pOverWriteGrafico As Integer
End Type
Public nMunicionData() As tNewMunicionData
Public nCapaData() As tNewCapa
Public Type tDecor
   
   MaxHP          As Long           ' Cuanta vida tiene el decor
   Respawn        As Long           ' Cada cuanto respawnea
   VALUE          As Single         ' Modificaro del objeto que da
   DecorGrh(1 To 5)  As Integer   ' Graficos
   Atacable       As Byte           ' SI pueden atacarse
   clave          As Integer        ' Para las puertas?
   Objeto()       As tDecorDrop        '
   CantObjetos    As Byte
   DecorType      As Byte
   EstadoDefault  As Byte           ' Cual es el estado default del decor
   TileH          As Byte
   TileW          As Byte
   OffX           As Integer
   OffY           As Integer
   TileTransY     As Byte
   Sombra       As Byte
End Type
Public Cantdecordata As Integer
Public DecoData() As tDecor

' Marian
Public EluTable(1 To STAT_MAXELV)   As Long
Public StaTable(1 To STAT_MAXELV)   As Integer

Public num_Chirimbolos_data As Byte
Public Chirimbolos_Data() As tChirimbolo_Data
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Public Sub Compilar_NewEstatics()
Dim s As String
Dim i As Long
Dim z As Long

s = App.path & "\RES\Index\NewEstatics.dat"

NumEstatics = Val(GetVar(s, "INIT", "num"))

If NumEstatics > 0 Then
ReDim EstaticData(1 To NumEstatics)
For i = 1 To NumEstatics
    With EstaticData(i)
        .L = Val(GetVar(s, CStr(i), "Left"))
        .t = Val(GetVar(s, CStr(i), "Top"))
        .w = Val(GetVar(s, CStr(i), "Width"))
        .h = Val(GetVar(s, CStr(i), "Height"))
        .TW = .w / 32
        .TH = .h / 32
    End With
Next i

Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwEstatics.IND" For Binary Access Write Lock Write As K
    Put K, , NumEstatics
    For i = 1 To NumEstatics
        With EstaticData(i)
            Put K, , .L
            Put K, , .t
            Put K, , .w
            Put K, , .h
            Put K, , .TW
            Put K, , .TH
                   
        End With
    Next i
Close K
End If
End Sub
Public Sub Compilar_NewIndex()
Dim s As String
Dim i As Long
Dim z As Long

s = App.path & "\RES\Index\NewIndex.dat"

NumNewIndex = Val(GetVar(s, "INIT", "num"))

If NumNewIndex > 0 Then
ReDim NewIndexData(1 To NumNewIndex)
For i = 1 To NumNewIndex
    With NewIndexData(i)
        .Dinamica = Val(GetVar(s, CStr(i), "Dinamica"))
        .Estatic = Val(GetVar(s, CStr(i), "Estatica"))
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
    End With
Next i


Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwIndex.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewIndex
    For i = 1 To NumNewIndex
        With NewIndexData(i)
            Put K, , .Dinamica
            Put K, , .Estatic
            Put K, , .OverWriteGrafico
       
        End With
    Next i
Close K
End If
End Sub
Sub loadsd()
Dim t As Long
ns = Val(GetVar(App.path & "\ENCODE\hechizos.dat", "INIT", "NumeroHechizos"))
ReDim Spells(1 To ns)
Dim F As String
F = App.path & "\ENCODE\hechizos.dat"
Dim i As Integer
For t = 1 To ns
Dim p As Byte

    
    If Val(GetVar(F, "HECHIZO" & t, "SUBEHP")) = 2 Then
    p = 0
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEHP")) = 1 Then
    p = 1
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEMANA")) = 2 Then
    p = 2
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEMANA")) = 1 Then
    p = 3
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBESTA")) = 2 Then
    p = 4
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBESTA")) = 1 Then
    p = 5
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEHAM")) = 2 Then
    p = 6
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEHAM")) = 1 Then
    p = 7
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBESED")) = 2 Then
    p = 8
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBESED")) = 1 Then
    p = 9
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEAG")) = 2 Then
    p = 10
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEAG")) = 1 Then
        If Val(GetVar(F, "HECHIZO" & t, "SUBEFU")) = 1 Then
            p = 15
        Else
            p = 11
        End If
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEFU")) = 2 Then
    p = 12
    ElseIf Val(GetVar(F, "HECHIZO" & t, "SUBEFU")) = 1 Then
    p = 13
    Else
    p = 14
    End If
    
    Spells(t).Tipo = p
    i = Val(GetVar(F, "HECHIZO" & t, "wav"))
    If i > 255 Or i < 0 Then i = 1
    Spells(t).Sound = i
    Spells(t).fx = Val(GetVar(F, "HECHIZO" & t, "Fxgrh"))
    Spells(t).loops = Val(GetVar(F, "HECHIZO" & t, "loops"))
    Spells(t).magicwords = GetVar(F, "HECHIZO" & t, "PalabrasMagicas")
    Spells(t).propiomsg = GetVar(F, "HECHIZO" & t, "PropioMsg")
    Spells(t).targetmsg = GetVar(F, "HECHIZO" & t, "TargetMsg")
    Spells(t).castermsg = GetVar(F, "HECHIZO" & t, "HechizeroMsg")
    Spells(t).Info = GetVar(F, "HECHIZO" & t, "DESC")
    Spells(t).Manareq = Val(GetVar(F, "HECHIZO" & t, "ManaRequerido"))
    Spells(t).Skills = Val(GetVar(F, "HECHIZO" & t, "MinSkill"))
    Spells(t).Nombre = GetVar(F, "HECHIZO" & t, "Nombre")
    Spells(t).Libro = Val(GetVar(F, "HECHIZO" & t, "Libro"))
    Spells(t).CasterFx = Val(GetVar(F, "HECHIZO" & t, "CasterFX"))
    Spells(t).CasterLoop = Val(GetVar(F, "HECHIZO" & t, "CasterLoop"))
    
    
    
    

Next t

End Sub
Sub WriteVar(ByVal file As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal VALUE As String)
      '*****************************************************************
      'Writes a var to a text file
      '*****************************************************************
      writeprivateprofilestring Main, Var, VALUE, file
End Sub
Sub Load_NoblezaData()
Dim s As String
Dim i As Long
Dim p As Long
s = App.path & "\ENCODE\Nobleza.dat"


Nobleza_Data.NumItems = Val(GetVar(s, "INIT", "NUM"))

ReDim Nobleza_Data.Items(1 To Nobleza_Data.NumItems)



For i = 1 To Nobleza_Data.NumItems

With Nobleza_Data.Items(i)

    .NumItems_Requeridos = Val(GetVar(s, "OBJ" & i, "CantItem"))
    .Numero = Val(GetVar(s, "OBJ" & i, "ObjIndexRecompensa"))
    ReDim .Items_Requeridos(1 To .NumItems_Requeridos)
    ReDim .cantItems_Requeridos(1 To .NumItems_Requeridos)
    For p = 1 To .NumItems_Requeridos
        .Items_Requeridos(p) = Val(GetVar(s, "OBJ" & i, "ObjIndex" & p))
        .cantItems_Requeridos(p) = Val(GetVar(s, "OBJ" & i, "Cantidad" & p))
    Next p
End With
Next i










End Sub


Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim L As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function
Sub LoadTables()

   Dim i As Long
   Dim tmpStr As String
   
   For i = 1 To STAT_MAXELV
      EluTable(i) = CLng(GetVar(App.path & "\ENCODE\Tables.dat", "EluTable", CStr(i)))
   Next i
   
   For i = 1 To STAT_MAXELV
      StaTable(i) = CInt(GetVar(App.path & "\ENCODE\Tables.dat", "StaTable", CStr(i)))
   Next i
   
      
End Sub

Sub Load_FD()

Dim s As String
Dim i As Long
Dim X As Long
s = App.path & "\ENCODE\FORMULARIO_DATA.dat"

NUM_FD = Val(GetVar(s, "INIT", "NUM"))

ReDim FD(1 To NUM_FD)

For i = 1 To NUM_FD

With FD(i)
        .TieneSpecial = IIf(Val(GetVar(s, "FD" & i, "Special")) = 1, True, False)
    .SurfaceNum = Val(GetVar(s, "FD" & i, "Grafico"))
    .ScreenX = Val(GetVar(s, "FD" & i, "X"))
    .ScreenY = Val(GetVar(s, "FD" & i, "Y"))
    
    .GraficosX = Val(GetVar(s, "FD" & i, "GRAFICOSENX"))
    .GraficosY = Val(GetVar(s, "FD" & i, "GRAFICOSENY"))
    If .GraficosX = 0 Then .GraficosX = 1
    If .GraficosY = 0 Then .GraficosY = 1
    .Num_Checks = Val(GetVar(s, "FD" & i, "NumChecks"))
    .Draw_Stage = Val(GetVar(s, "FD" & i, "NivelDibujo"))
    .num_Buttons = Val(GetVar(s, "FD" & i, "NumeroBotones"))
    .AlphaValue = Val(GetVar(s, "FD" & i, "AlphaValue"))
    .Num_Textos = Val(GetVar(s, "FD" & i, "NumeroTextos"))
    .Num_TextBox = Val(GetVar(s, "FD" & i, "NumTextBox"))
    .Height = Val(GetVar(s, "FD" & i, "Height"))
    .Width = Val(GetVar(s, "FD" & i, "Width"))
    
    If .AlphaValue = 0 Then .AlphaValue = 255
    If .num_Buttons > 0 Then
    ReDim .Buttons(1 To .num_Buttons)
    
    For X = 1 To .num_Buttons
        
        With FD(i).Buttons(X)
        
            .NormalSurfaceNum = Val(GetVar(s, "FD" & i, "Btn_" & X & "_NormalGrafico"))
            .SelectSurfaceNum = Val(GetVar(s, "FD" & i, "Btn_" & X & "_SelectGrafico"))
            .PressSurfaceNum = Val(GetVar(s, "FD" & i, "Btn_" & X & "_PressGrafico"))
            
            .size.Top = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Top"))
            .size.Left = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Left"))
            .size.Right = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Width"))
            .size.bottom = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Height"))
            
            .Sound = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Sound"))
            .HandIco = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Hand"))
            .Caption = Val(GetVar(s, "FD" & i, "Btn_" & X & "_Caption"))
            
            .Normal_Rojo = Val(GetVar(s, "FD" & i, "Btn_" & X & "_NormalRojo"))
            .Normal_Verde = Val(GetVar(s, "FD" & i, "Btn_" & X & "_NormalVerde"))
            .Normal_Azul = Val(GetVar(s, "FD" & i, "Btn_" & X & "_NormalAzul"))
                        
            .Sel_Rojo = Val(GetVar(s, "FD" & i, "Btn_" & X & "_SelRojo"))
            .Sel_Verde = Val(GetVar(s, "FD" & i, "Btn_" & X & "_SelVerde"))
            .Sel_Azul = Val(GetVar(s, "FD" & i, "Btn_" & X & "_SelAzul"))
            
            .Press_Rojo = Val(GetVar(s, "FD" & i, "Btn_" & X & "_PressRojo"))
            .Press_Verde = Val(GetVar(s, "FD" & i, "Btn_" & X & "_PressVerde"))
            .Press_Azul = Val(GetVar(s, "FD" & i, "Btn_" & X & "_PressAzul"))
        End With
    Next X
    End If
    If .Num_Textos > 0 Then
    ReDim FD(i).Textos(1 To .Num_Textos)
    For X = 1 To .Num_Textos
    
        With FD(i).Textos(X)
            .texto = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Texto"))
            
            .X = Val(GetVar(s, "FD" & i, "Txt_" & X & "_X"))
            
            .Y = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Y"))
            
            .MaxWidth = Val(GetVar(s, "FD" & i, "Txt_" & X & "_MaxWidth"))
            .IniciaVisible = Val(GetVar(s, "FD" & i, "Txt_" & X & "_IniciaVisible"))
            
            
            .A = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Alpha"))
            
            .r = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Rojo"))
            
            .g = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Verde"))
                    
            
            .b = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Azul"))
            
            .Centrar = Val(GetVar(s, "FD" & i, "Txt_" & X & "_Centrar"))
                    
        End With
    Next X
    End If
    If .Num_TextBox > 0 Then
        ReDim .TextBox(1 To .Num_TextBox)
        
        For X = 1 To .Num_TextBox
            .TextBox(X).SurfaceNum = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_Surface"))
            .TextBox(X).Centrar = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_Centrar"))
            .TextBox(X).TipoTexto = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_TipoTexto"))
            .TextBox(X).X = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_x"))
            .TextBox(X).Y = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_y"))
            .TextBox(X).OffsetX = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_Offsetx"))
            .TextBox(X).OffsetY = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_Offsety"))
            .TextBox(X).IniciaVisible = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_IniciaVisible"))
            .TextBox(X).fA = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_AFont"))
            .TextBox(X).fR = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_RFont"))
            .TextBox(X).fG = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_GFont"))
            .TextBox(X).fb = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_BFont"))
            .TextBox(X).w = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_W"))
            .TextBox(X).h = Val(GetVar(s, "FD" & i, "TxtB_" & X & "_H"))
            
        Next X
    End If
    
        If .Num_Checks > 0 Then
        ReDim .Checks(1 To .Num_Checks)
        For X = 1 To .Num_Checks
            .Checks(X).Caption = Val(GetVar(s, "FD" & i, "Chk_" & X & "_Caption"))
            .Checks(X).X = Val(GetVar(s, "FD" & i, "Chk_" & X & "_X"))
            .Checks(X).Y = Val(GetVar(s, "FD" & i, "Chk_" & X & "_Y"))
            .Checks(X).w = Val(GetVar(s, "FD" & i, "Chk_" & X & "_W"))
            .Checks(X).h = Val(GetVar(s, "FD" & i, "Chk_" & X & "_H"))
            .Checks(X).min = Val(GetVar(s, "FD" & i, "Chk_" & X & "_MIN"))
            .Checks(X).max = Val(GetVar(s, "FD" & i, "Chk_" & X & "_MAX"))
            .Checks(X).Tipo_Check = Val(GetVar(s, "FD" & i, "Chk_" & X & "_TipoCheck"))
            .Checks(X).CheckSurface = Val(GetVar(s, "FD" & i, "Chk_" & X & "_CheckSurface"))
            .Checks(X).SurfaceNum = Val(GetVar(s, "FD" & i, "Chk_" & X & "_Grafico"))
            .Checks(X).IniciaVisible = Val(GetVar(s, "FD" & i, "Chk_" & X & "_IniciaVisible"))
            .Checks(X).ColorA = Val(GetVar(s, "FD" & i, "Chk_" & X & "_A"))
            .Checks(X).ColorR = Val(GetVar(s, "FD" & i, "Chk_" & X & "_R"))
            .Checks(X).ColorG = Val(GetVar(s, "FD" & i, "Chk_" & X & "_G"))
            .Checks(X).ColorB = Val(GetVar(s, "FD" & i, "Chk_" & X & "_B"))
        Next X
    End If
    
    
End With

Next i



End Sub
Public Sub Load_HabilidadesData()

Dim i As Long



For i = 1 To MAX_HABILIDADES

    Habilidades(i).Grafico = Val(GetVar(App.path & "\ENCODE\Habilidades.dat", "HAB" & i, "Grafico"))

Next i

End Sub
Public Sub Load_NpcnoHostiles()

Dim s As String
Dim i As Long
s = App.path & "\ENCODE\npcs-hostiles.dat"

num_npcs_h = Val(GetVar(s, "init", "numnpcs")) + 1

ReDim hostiles(0 To num_npcs_h)
For i = 500 To num_npcs_h
With hostiles(i - 499)

    .Body = Val(GetVar(s, "NPC" & i, "body"))
    .Head = Val(GetVar(s, "NPC" & i, "Head"))
    .MAX_HP = Val(GetVar(s, "NPC" & i, "MaxHP"))
    .Snd1 = Val(GetVar(s, "NPC" & i, "Snd1"))
    .Snd2 = Val(GetVar(s, "NPC" & i, "Snd2"))
    .Nombre = GetVar(s, "NPC" & i, "Name")
    
    
End With
Next i
End Sub
Public Sub Load_NpcHostiles()

Dim s As String
Dim i As Long
s = App.path & "\ENCODE\npcs.dat"

num_npcs_nh = Val(GetVar(s, "init", "numnpcs"))

ReDim nHostiles(1 To num_npcs_nh)
For i = 1 To num_npcs_nh
With nHostiles(i)


    
        .Body = Val(GetVar(s, "NPC" & i, "body"))
    .Head = Val(GetVar(s, "NPC" & i, "Head"))
    .MAX_HP = Val(GetVar(s, "NPC" & i, "MaxHP"))
    .Nombre = GetVar(s, "NPC" & i, "Name")
    .Desc = GetVar(s, "NPC" & i, "Desc")
    .NPCTYPE = Val(GetVar(s, "NPC" & i, "NpcType"))
    
End With
Next i
End Sub
Public Sub LoadQuest()
Dim s As String
Dim i As Long
Dim z As Long
s = App.path & "\ENCODE\Quests.dat"

nQuest = Val(GetVar(s, "MAIN", "NUMQUESTS"))
ReDim Quests(1 To nQuest)

For i = 1 To nQuest

    Quests(i).Tipo = Val(GetVar(s, CStr(i), "Tipo"))
    
    Quests(i).nTargets = Val(GetVar(s, CStr(i), "numero_targets"))
    Quests(i).Oro = Val(GetVar(s, CStr(i), "recompensa_oro"))
    Quests(i).Puntos = Val(GetVar(s, CStr(i), "recompensa_puntos"))
    Quests(i).Exp = Val(GetVar(s, CStr(i), "recompensa_exp"))
    Quests(i).numritems = Val(GetVar(s, CStr(i), "recompensa_numero_items"))
    Quests(i).TipoReco = Val(GetVar(s, CStr(i), "recompensa_tipo"))
    Quests(i).Nombre = GetVar(s, CStr(i), "nombre")
    
    
    If Quests(i).numritems > 0 Then
        ReDim Quests(i).Item(1 To Quests(i).numritems)
        ReDim Quests(i).Cant(1 To Quests(i).numritems)
        For z = 1 To Quests(i).numritems
            Quests(i).Item(z) = Val(GetVar(s, CStr(i), "recompensa_item_tipo" & z))
            Quests(i).Cant(z) = Val(GetVar(s, CStr(i), "recompensa_item_cant" & z))
            
        
        Next z
    End If
    Quests(i).Desc = GetVar(s, CStr(i), "desc")
    
    If Quests(i).nTargets > 0 Then
    ReDim Quests(i).Targets(1 To Quests(i).nTargets)
    ReDim Quests(i).TargetsCant(1 To Quests(i).nTargets)
    For z = 1 To Quests(i).nTargets
        Quests(i).Targets(z) = Val(GetVar(s, CStr(i), "target_tipo" & z))
        Quests(i).TargetsCant(z) = Val(GetVar(s, CStr(i), "target_cant" & z))
    Next z
    End If
    
    
    

    
    

Next i
End Sub
Public Sub SaveQuests(ByVal F As Integer)

Dim i As Long
Dim z As Long

    Put F, , nQuest
    
    For i = 1 To nQuest
        Put F, , Quests(i).Tipo
        Put F, , Quests(i).nTargets
        For z = 1 To Quests(i).nTargets
            Put F, , Quests(i).Targets(z)
            Put F, , Quests(i).TargetsCant(z)
        Next z
        Put F, , Quests(i).TipoReco
        Select Case Quests(i).TipoReco
            Case 1
                Put F, , Quests(i).Oro
            Case 2
                Put F, , Quests(i).numritems
                For z = 1 To Quests(i).numritems
                    Put F, , Quests(i).Item(z)
                    Put F, , Quests(i).Cant(z)
                Next z
            Case 3
                Put F, , Quests(i).Oro
                Put F, , Quests(i).numritems
                For z = 1 To Quests(i).numritems
                    Put F, , Quests(i).Item(z)
                    Put F, , Quests(i).Cant(z)
                Next z
            Case 4
            Case 5
                Put F, , Quests(i).Puntos
            Case 6
                Put F, , Quests(i).Puntos
                Put F, , Quests(i).Oro
        End Select
        Put F, , Quests(i).Exp
    
    Next i
End Sub
Sub LoadMapaData()

Dim i As Integer
Dim F As Integer
Dim s As String
F = FreeFile

Open App.path & "\ENCODE\Mapadata.txt" For Input As #F
    Line Input #F, s
    
    NroMapas = Val(s)
    
    ReDim MapaData(1 To NroMapas)
Do Until EOF(F)
i = i + 1

    Line Input #F, s
    
    MapaData(i).X = Val(Readfield(1, s, 44))
    MapaData(i).Y = Val(Readfield(2, s, 44))
    
Loop
Close #F


End Sub
Function Readfield(ByVal pos As Integer, ByVal texto As String, ByVal separador As Byte) As String
'caserita
Dim i As Long
Dim t As Long
Dim L As Long
Dim K As Long
Dim s As String

L = Len(texto)


Do Until i = L
    i = i + 1
    If Asc(mid$(texto, i, 1)) = separador Then
        K = K + 1
        If K = pos Then
            Readfield = s
            Exit Do
        Else
            s = vbNullString
        End If
    Else
        s = s & mid$(texto, i, 1)
    End If
Loop
Readfield = s
End Function


Public Sub Compilar_Archivo_Index()
    Compilar_NwANim
    Compilar_NwBody
    Compilar_NewFx
    Compilar_NwShield
    Compilar_NwWeapons
    Compilar_NewEstatics
    Compilar_NewIndex
    Compilar_NwHelmet
    Compilar_NwMunicion
    Compilar_NwCapa
    Compilar_NwHeads
    Cargar_Archivo_Index
    Dim K As Integer
    K = FreeFile
    Dim p As Integer
    Open App.path & "\RES\OUTPUT\INDEX.BIN" For Binary Access Write Lock Write As K
        For p = 1 To 11
            Put #K, , IXAR(p).St
            Put #K, , IXAR(p).Le
        Next p
        For p = 1 To 11
            Put #K, , IXAR(p).Da
        Next p
    Close K
    
    MsgBox "compilación exitosa."

End Sub
Public Sub Cargar_Archivo_Index()


Dim path As String

Dim i As Integer

Dim UltimoS As Long
Dim UltimoL As Long
ReDim IXAR(1 To 11)

path = App.path & "\RES\INDEX\"
UltimoS = (8 * 11) - 1


i = FreeFile
Open path & "NewFx.BIN" For Binary Access Read Lock Read As i
    IXAR(1).Le = LOF(i)
    ReDim IXAR(1).Da(0 To IXAR(1).Le - 1)
    Get i, , IXAR(1).Da
Close i

IXAR(1).St = UltimoS + UltimoL
UltimoS = IXAR(1).St
UltimoL = IXAR(1).Le

i = FreeFile
Open path & "NwAnim.IND" For Binary Access Read Lock Read As i
    IXAR(2).Le = LOF(i)
    ReDim IXAR(2).Da(0 To IXAR(2).Le - 1)
    Get i, , IXAR(2).Da
Close i

IXAR(2).St = UltimoS + UltimoL
UltimoS = IXAR(2).St
UltimoL = IXAR(2).Le

i = FreeFile
Open path & "NwBody.IND" For Binary Access Read Lock Read As i
    IXAR(3).Le = LOF(i)
    ReDim IXAR(3).Da(0 To IXAR(3).Le - 1)
    Get i, , IXAR(3).Da
Close i

IXAR(3).St = UltimoS + UltimoL
UltimoS = IXAR(3).St
UltimoL = IXAR(3).Le

i = FreeFile
Open path & "NwShields.IND" For Binary Access Read Lock Read As i
    IXAR(4).Le = LOF(i)
    If IXAR(4).Le > 0 Then
    ReDim IXAR(4).Da(0 To IXAR(4).Le - 1)
    Get i, , IXAR(4).Da
    End If
Close i

IXAR(4).St = UltimoS + UltimoL
UltimoS = IXAR(4).St
UltimoL = IXAR(4).Le

i = FreeFile
Open path & "NwWeapons.IND" For Binary Access Read Lock Read As i
    IXAR(5).Le = LOF(i)
    If IXAR(5).Le > 0 Then
    ReDim IXAR(5).Da(0 To IXAR(5).Le - 1)
    Get i, , IXAR(5).Da
    End If
Close i

IXAR(5).St = UltimoS + UltimoL
UltimoS = IXAR(5).St
UltimoL = IXAR(5).Le

i = FreeFile
Open path & "NwIndex.IND" For Binary Access Read Lock Read As i
    IXAR(6).Le = LOF(i)
    If IXAR(6).Le > 0 Then
    ReDim IXAR(6).Da(0 To IXAR(6).Le - 1)
    Get i, , IXAR(6).Da
    End If
Close i

IXAR(6).St = UltimoS + UltimoL
UltimoS = IXAR(6).St
UltimoL = IXAR(6).Le

i = FreeFile
Open path & "NwEstatics.IND" For Binary Access Read Lock Read As i
    IXAR(7).Le = LOF(i)
    If IXAR(7).Le > 0 Then
    ReDim IXAR(7).Da(0 To IXAR(7).Le - 1)
    Get i, , IXAR(7).Da
    End If
Close i

IXAR(7).St = UltimoS + UltimoL
UltimoS = IXAR(7).St
UltimoL = IXAR(7).Le

i = FreeFile
Open path & "NwHelmets.IND" For Binary Access Read Lock Read As i
    IXAR(8).Le = LOF(i)
    If IXAR(8).Le > 0 Then
    ReDim IXAR(8).Da(0 To IXAR(8).Le - 1)
    Get i, , IXAR(8).Da
    End If
Close i

IXAR(8).St = UltimoS + UltimoL
UltimoS = IXAR(8).St
UltimoL = IXAR(8).Le

i = FreeFile
Open path & "NwMunicion.IND" For Binary Access Read Lock Read As i
    IXAR(9).Le = LOF(i)
    If IXAR(9).Le > 0 Then
    ReDim IXAR(9).Da(0 To IXAR(9).Le - 1)
    Get i, , IXAR(9).Da
    End If
Close i

IXAR(9).St = UltimoS + UltimoL
UltimoS = IXAR(9).St
UltimoL = IXAR(9).Le

i = FreeFile
Open path & "NwCapas.IND" For Binary Access Read Lock Read As i
    IXAR(10).Le = LOF(i)
    If IXAR(10).Le > 0 Then
    ReDim IXAR(10).Da(0 To IXAR(10).Le - 1)
    Get i, , IXAR(10).Da
    End If
Close i

IXAR(10).St = UltimoS + UltimoL
UltimoS = IXAR(10).St
UltimoL = IXAR(10).Le

i = FreeFile
Open path & "NwHeads.IND" For Binary Access Read Lock Read As i
    IXAR(11).Le = LOF(i)
    If IXAR(11).Le > 0 Then
    ReDim IXAR(11).Da(0 To IXAR(11).Le - 1)
    Get i, , IXAR(11).Da
    End If
Close i

IXAR(11).St = UltimoS + UltimoL
UltimoS = IXAR(11).St
UltimoL = IXAR(11).Le


End Sub
Public Sub Compilar_NwANim()

Dim s As String
Dim i As Long
Dim p As Long
Dim K As Long
Dim GrafCounter As Integer
Dim num_nwanim As Integer

s = App.path & "\RES\INDEX\NewAnim.dat"


num_nwanim = Val(GetVar(s, "NW_ANIM", "NUM"))

If num_nwanim < 1 Then Exit Sub

ReDim NewAnimationData(1 To num_nwanim)

For i = 1 To num_nwanim


With NewAnimationData(i)
    .Grafico = Val(GetVar(s, "ANIMACION" & i, "Grafico"))
    .Columnas = Val(GetVar(s, "ANIMACION" & i, "Columnas"))
    .Filas = Val(GetVar(s, "ANIMACION" & i, "Filas"))
    .Height = Val(GetVar(s, "ANIMACION" & i, "Alto"))
    .Width = Val(GetVar(s, "ANIMACION" & i, "Ancho"))
    .NumFrames = Val(GetVar(s, "ANIMACION" & i, "NumeroFrames"))
    .Velocidad = Val(GetVar(s, "ANIMACION" & i, "Velocidad"))
    .TileWidth = .Width / 32
    .TileHeight = .Height / 32
    .Romboidal = Val(GetVar(s, "ANIMACION" & i, "AnimacionRomboidal"))
    .OffsetX = Val(GetVar(s, "ANIMACION" & i, "OffsetX"))
    .OffsetY = Val(GetVar(s, "ANIMACION" & i, "OffsetY"))
    .Initial = Val(GetVar(s, "ANIMACION" & i, "Inicial"))
    .TipoAnimacion = Val(GetVar(s, "ANIMACION" & i, "TipoAnimacion"))
    
    If .NumFrames > 0 Then
    ReDim .Indice(1 To .NumFrames) As tNewIndice
    GrafCounter = .Grafico
    If .Initial = 0 Then .Initial = 1
    K = .Initial - 1
    If K >= CInt(.Columnas) * CInt(.Filas) Then
        K = K Mod (CInt(.Columnas) * CInt(.Filas))
    End If
    For p = 1 To .NumFrames
        K = K + 1
        .Indice(p).X = (((K - 1) Mod .Columnas) * .Width)
        .Indice(p).Y = ((Int((K - 1) / .Columnas)) * .Height)
        .Indice(p).Grafico = GrafCounter
        If (K Mod (CInt(.Columnas) * CInt(.Filas))) = 0 And ((K + 1) - .Initial) < .NumFrames Then
            GrafCounter = GrafCounter + 1
            K = 0
        End If
    Next p
    End If
End With
Next i

Dim z As Integer
z = FreeFile

Open App.path & "\RES\INDEX\NwAnim.IND" For Binary Access Write Lock Write As z
    
    Put z, , num_nwanim
    
    For p = 1 To num_nwanim

        With NewAnimationData(p)
        
            Put z, , .Grafico
            Put z, , .Filas
            Put z, , .Columnas
            Put z, , .Height
            Put z, , .Width
            Put z, , .NumFrames
            Put z, , .Velocidad
            Put z, , .TileWidth
            Put z, , .TileHeight
            Put z, , .Romboidal
            Put z, , .OffsetX
            Put z, , .OffsetY
            Put z, , .TipoAnimacion
            
            For K = 1 To .NumFrames
            
                Put z, , .Indice(K).Grafico
                Put z, , .Indice(K).X
                Put z, , .Indice(K).Y
            Next K

        End With
    
    Next p

Close z
End Sub
Public Sub Compilar_NwShield()

Dim s As String
Dim i As Long
Dim z As Long
Dim NumNewShields As Integer
s = App.path & "\RES\INDEX\Nwshields.dat"

NumNewShields = Val(GetVar(s, "INIT", "num"))

If NumNewShields > 0 Then
ReDim nShieldDATA(1 To NumNewShields)
For i = 1 To NumNewShields

    With nShieldDATA(i)

        .Alpha = Val(GetVar(s, CStr(i), "Alpha"))
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(s, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i



Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwShields.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewShields
    For i = 1 To NumNewShields
        With nShieldDATA(i)
            Put K, , .Alpha
            Put K, , .OverWriteGrafico
            For z = 1 To 4
                
                Put K, , .mMovimiento(z)
                
            Next z
       
        End With
    Next i
Close K
End If

End Sub
Public Sub Compilar_NwHelmet()
    
Dim s As String
Dim i As Long
Dim z As Long

s = App.path & "\RES\INDEX\NewHelmets.dat"

NumNewHelmet = Val(GetVar(s, "INIT", "num"))

If NumNewHelmet > 0 Then
ReDim NHelmetData(1 To NumNewHelmet)
For i = 1 To NumNewHelmet

    With NHelmetData(i)

        .Alpha = Val(GetVar(s, "HELMET" & CStr(i), "Alpha"))
        .OffsetY = Val(GetVar(s, "HELMET" & CStr(i), "OFFSET_DIBUJO"))
        .OffsetLat = Val(GetVar(s, "HELMET" & CStr(i), "OFFSET_LAT"))
        
        .mMovimiento(1) = Val(GetVar(s, "HELMET" & CStr(i), "NORTH"))
        .mMovimiento(2) = Val(GetVar(s, "HELMET" & CStr(i), "EAST"))
                .mMovimiento(3) = Val(GetVar(s, "HELMET" & CStr(i), "SOUTH"))
                .mMovimiento(4) = Val(GetVar(s, "HELMET" & CStr(i), "WEST"))
                
       
    End With
Next i



Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwHelmets.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewHelmet
    For i = 1 To NumNewHelmet
        With NHelmetData(i)
            Put K, , .Alpha
            Put K, , .OffsetY
            Put K, , .OffsetLat
            For z = 1 To 4
                
                Put K, , .mMovimiento(z)
                
            Next z
       
        End With
    Next i
Close K
End If

End Sub
Public Sub Compilar_NwWeapons()

Dim s As String
Dim i As Long
Dim z As Long
Dim NumNewWeapons As Integer
s = App.path & "\RES\INDEX\NwWeapons.dat"

NumNewWeapons = Val(GetVar(s, "INIT", "num"))

If NumNewWeapons > 0 Then
ReDim nWeaponData(1 To NumNewWeapons)
For i = 1 To NumNewWeapons

    With nWeaponData(i)

        .Alpha = Val(GetVar(s, CStr(i), "Alpha"))
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
        
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(s, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i



Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwWeapons.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewWeapons
    For i = 1 To NumNewWeapons
        With nWeaponData(i)
            Put K, , .Alpha
            Put K, , .OverWriteGrafico
            For z = 1 To 4
                
                Put K, , .mMovimiento(z)
                
            Next z
       
        End With
    Next i
Close K
End If

End Sub
Public Sub Compilar_NwMunicion()

Dim s As String
Dim i As Long
Dim z As Long
Dim NumNewM As Integer
s = App.path & "\RES\INDEX\NwMunicion.dat"

NumNewM = Val(GetVar(s, "INIT", "num"))

If NumNewM > 0 Then
ReDim nMunicionData(1 To NumNewM)
For i = 1 To NumNewM

    With nMunicionData(i)

        .Alpha = Val(GetVar(s, CStr(i), "Alpha"))
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
        
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(s, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i

End If

Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwMunicion.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewM
    If NumNewM > 0 Then
    For i = 1 To NumNewM
        With nMunicionData(i)
            Put K, , .Alpha
            Put K, , .OverWriteGrafico
            For z = 1 To 4
                
                Put K, , .mMovimiento(z)
                
            Next z
       
        End With
    Next i
    End If
Close K

End Sub
Public Sub Compilar_NwCapa()

Dim s As String
Dim i As Long
Dim z As Long
Dim NumNewM As Integer
s = App.path & "\RES\INDEX\NwCapa.dat"

NumNewM = Val(GetVar(s, "INIT", "num"))

If NumNewM > 0 Then
ReDim nCapaData(1 To NumNewM)
For i = 1 To NumNewM

    With nCapaData(i)

        .Alpha = Val(GetVar(s, CStr(i), "Alpha"))
        .aOverWriteGrafico = Val(GetVar(s, CStr(i), "aOverWriteGrafico"))
        .pOverWriteGrafico = Val(GetVar(s, CStr(i), "pOverWriteGrafico"))
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(s, CStr(i), "Mov" & z))

        Next z

    End With
Next i


End If

Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwCapa.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewM
    If NumNewM > 0 Then
    For i = 1 To NumNewM
        With nCapaData(i)
            Put K, , .Alpha
            Put K, , .aOverWriteGrafico
            Put K, , .pOverWriteGrafico
            For z = 1 To 4
                
                Put K, , .mMovimiento(z)
                
            Next z
       
        End With
    Next i
    End If
Close K

End Sub
Public Sub Compilar_NwHeads()

Dim s As String
Dim i As Long
Dim z As Long
Dim NumNewM As Integer
s = App.path & "\RES\INDEX\NewHeads.dat"

NumNewM = Val(GetVar(s, "INIT", "num"))

If NumNewM > 0 Then
ReDim NHeadData(1 To NumNewM)
For i = 1 To NumNewM
    With NHeadData(i)
        .Raza = Val(GetVar(s, "HEAD" & CStr(i), "RAZA"))
        .OffsetDibujoY = Val(GetVar(s, "HEAD" & CStr(i), "OFFSET_DIBUJO"))
        .OffsetOjos = Val(GetVar(s, "HEAD" & CStr(i), "OFFSET_OJOS"))
        .Genero = Val(GetVar(s, "HEAD" & CStr(i), "GENERO"))
        .Frame(2) = Val(GetVar(s, "HEAD" & CStr(i), "EAST"))
        .Frame(1) = Val(GetVar(s, "HEAD" & CStr(i), "NORTH"))
        .Frame(3) = Val(GetVar(s, "HEAD" & CStr(i), "SOUTH"))
        .Frame(4) = Val(GetVar(s, "HEAD" & CStr(i), "WEST"))
    End With
Next i
End If

Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwHeads.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewM
    If NumNewM > 0 Then
    For i = 1 To NumNewM
        With NHeadData(i)
            Put K, , .OffsetDibujoY
            Put K, , .OffsetOjos
            Put K, , .Raza
            Put K, , .Genero
            For z = 1 To 4
                
                Put K, , .Frame(z)
                
            Next z
       
        End With
    Next i
    End If
Close K

End Sub

Public Sub Compilar_NwBody()

Dim s As String
Dim i As Long
Dim z As Long
Dim NumNewBodys As Integer
s = App.path & "\RES\INDEX\NewBody.dat"

NumNewBodys = Val(GetVar(s, "INIT", "num"))

If NumNewBodys > 0 Then
ReDim nBodyData(1 To NumNewBodys)
For i = 1 To NumNewBodys

    With nBodyData(i)

        .bAtaque = IIf(Val(GetVar(s, CStr(i), "Ataque")), True, False)
        .bContinuo = IIf(Val(GetVar(s, CStr(i), "Continuo")), True, False)
        .bReposo = IIf(Val(GetVar(s, CStr(i), "Reposo")), True, False)
        .bAtacado = IIf(Val(GetVar(s, CStr(i), "Atacado")), True, False)
        .bDeath = IIf(Val(GetVar(s, CStr(i), "Muerte")), True, False)
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
        .OffsetY = Val(GetVar(s, CStr(i), "OffsetY"))
        .Capa = Val(GetVar(s, CStr(i), "Capa"))
        If .bAtaque Then
            For z = 1 To 4
                .Attack(z) = Val(GetVar(s, CStr(i), "Ataque" & z))
                

            Next z
        End If
        If .bAtacado Then
            For z = 1 To 4
                .Attacked(z) = Val(GetVar(s, CStr(i), "Atacado" & z))
            Next z
        End If
        
        If .bReposo Then
            For z = 1 To 4
            .Reposo(z) = Val(GetVar(s, CStr(i), "Reposo" & z))
        

            Next z
        End If
        If .bDeath Then
            For z = 1 To 4
            .Death(z) = Val(GetVar(s, CStr(i), "Muerte" & z))
                

            Next z
        End If

        For z = 1 To 4
        
            .mMovement(z) = Val(GetVar(s, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i



Dim K As Integer
K = FreeFile

Open App.path & "\RES\INDEX\NwBody.IND" For Binary Access Write Lock Write As K
    Put K, , NumNewBodys
    For i = 1 To NumNewBodys
        With nBodyData(i)
            Put K, , .bContinuo
            Put K, , .bReposo
            Put K, , .bAtaque
            Put K, , .bAtacado
            Put K, , .bReposo
            Put K, , .bDeath
            Put K, , .OverWriteGrafico
            Put K, , .OffsetY
            Put K, , .Capa
            For z = 1 To 4
                
                Put K, , .mMovement(z)
                
            Next z
            If .bReposo Then
                For z = 1 To 4
                    Put K, , .Reposo(z)
                Next z
            End If
            If .bAtacado Then
                For z = 1 To 4
                    Put K, , .Attacked(z)
                Next z
            End If
            If .bAtaque Then
                For z = 1 To 4
                    Put K, , .Attack(z)
                Next z
            End If
            If .bDeath Then
                For z = 1 To 4
                    Put K, , .Death(z)
                Next z
            End If
                
    
        End With
    Next i
Close K
End If

End Sub
Public Sub GenerarFonts()
Dim s As String
Dim p As Long
Dim K As Integer
Dim ll As Long
s = App.path & "\FONTS\"

nFonts = Val(GetVar(s & "Fonts.dat", "INIT", "numFonts"))
ReDim Fonts(1 To nFonts)
ll = 2 + nFonts * 10
For p = 1 To nFonts
K = FreeFile

    Open s & p & ".dat" For Binary Access Read Lock Read As #K
        Fonts(p).lSize = LOF(K)
        ReDim Fonts(p).Data(0 To Fonts(p).lSize - 1)
        Get K, , Fonts(p).Data
        Fonts(p).lStart = ll
        ll = ll + Fonts(p).lSize
    Close #K
    Fonts(p).nT = Val(GetVar(s & "Fonts.dat", "FONT" & p, "Textura"))
Next p
For p = 1 To 20

    Debug.Print Fonts(1).Data(p)
Next p

K = FreeFile
Open App.path & "\OUTPUT\Fonts.bin" For Binary Access Write Lock Write As #K
    Put #K, , nFonts
    For p = 1 To nFonts
        Put #K, , Fonts(p).lStart
        Put #K, , Fonts(p).lSize
        Put #K, , Fonts(p).nT
    Next p
    For p = 1 To nFonts
        Put #K, , Fonts(p).Data
    Next p
Close #K
MsgBox "Compilación exitosa"
End Sub
Public Sub Compilar_NewFx()

Dim s As String
Dim t As Long
Dim r As Byte
Dim V As Byte
Dim A As Byte
s = App.path & "\RES\INDEX\NewFxs.dat"

Num_Fx = Val(GetVar(s, "INIT", "NumFx"))
Num_Med = Val(GetVar(s, "INIT", "NumMeditaciones"))


If Num_Fx > 0 Then ReDim FxData(1 To Num_Fx)
If Num_Med > 0 Then ReDim MedData(1 To Num_Med)

If Num_Fx > 0 Then

    For t = 1 To Num_Fx

        With FxData(t)

        
            .Animacion = Val(GetVar(s, "FX" & t, "Anim"))
            .OffsetX = Val(GetVar(s, "FX" & t, "OffsetX"))
            .OffsetY = Val(GetVar(s, "FX" & t, "OffsetY"))
            .Alpha = Val(GetVar(s, "FX" & t, "Alpha"))
            .Rombo = Val(GetVar(s, "FX" & t, "Rombo"))
            .Particula = Val(GetVar(s, "FX" & t, "Particula"))
            .Life = Val(GetVar(s, "FX" & t, "Life"))
            r = Val(GetVar(s, "FX" & t, "Rojo"))
            V = Val(GetVar(s, "FX" & t, "Verde"))
            A = Val(GetVar(s, "FX" & t, "Azul"))
            .AnimFinal = Val(GetVar(s, "FX" & t, "AnimFinal"))
            .AnimInicial = Val(GetVar(s, "FX" & t, "AnimInicial"))
            .ParaleloInicial = Val(GetVar(s, "FX" & t, "AnimInitParalelo"))
            .ParaleloStart = Val(GetVar(s, "FX" & t, "InitParaleloStart"))
            If r = 0 And V = 0 And A = 0 Then
                .Color = 0
            Else
                .Color = D3DColorARGB(255, r, V, A)
            End If
        End With
    
    Next t
End If

If Num_Med > 0 Then

    For t = 1 To Num_Med

        With MedData(t)
            
            .Animacion = Val(GetVar(s, "MED" & t, "Anim"))
            .OffsetX = Val(GetVar(s, "MED" & t, "OffsetX"))
            .OffsetY = Val(GetVar(s, "MED" & t, "OffsetY"))
            .Alpha = Val(GetVar(s, "MED" & t, "Alpha"))
            .Rombo = Val(GetVar(s, "MED" & t, "Rombo"))
            .Particula = Val(GetVar(s, "MED" & t, "Particula"))
            r = Val(GetVar(s, "MED" & t, "Rojo"))
            V = Val(GetVar(s, "MED" & t, "Verde"))
            A = Val(GetVar(s, "MED" & t, "Azul"))
            .Life = Val(GetVar(s, "MED" & t, "Life"))
            If r = 0 And V = 0 And A = 0 Then
                .Color = 0
            Else
                .Color = D3DColorARGB(255, r, V, A)
            End If
            .AnimFinal = Val(GetVar(s, "MED" & t, "AnimFinal"))
            .AnimInicial = Val(GetVar(s, "MED" & t, "AnimInicial"))
            .ParaleloInicial = Val(GetVar(s, "MED" & t, "AnimInitParalelo"))
                        .ParaleloStart = Val(GetVar(s, "MED" & t, "InitParaleloStart"))
        End With
    Next t
End If

Dim F As Integer

F = FreeFile


Open App.path & "\RES\INDEX\NewFX.BIN" For Binary Access Write Lock Write As #F
    Put #F, , Num_Fx
    Put #F, , Num_Med
    Put #F, , FxData
    Put #F, , MedData
Close #F

End Sub
Sub LoadDecorData()
'***************************************************
'Author: Marian
'Last Modification: -
'***************************************************
On Error GoTo ErrHandler

      Dim Decor   As Long
      Dim tmpStr  As String
      Dim LoopC   As Long
      Dim LoopD   As Long
      Dim tmpStr2 As String
      
      Dim L As String
    
      L = App.path & "\Encode\Decor.dat"
    
      'obtiene el numero de obj
      Cantdecordata = Val(GetVar(L, "INIT", "NumDecors"))
        
      ReDim DecoData(1 To Cantdecordata) As tDecor
      
      For Decor = 1 To Cantdecordata
      
         With DecoData(Decor)
            
            .DecorType = Val(GetVar(L, "DECOR" & Decor, "DecorType"))
            .MaxHP = Val(GetVar(L, "DECOR" & Decor, "MaxHP"))
            .Respawn = Val(GetVar(L, "DECOR" & Decor, "Respawn"))
            
            tmpStr = GetVar(L, "DECOR" & Decor, "DecorGrh")
            
            For LoopC = 1 To 5
               .DecorGrh(LoopC) = Val(Readfield(LoopC, tmpStr, 45))
            Next LoopC
            
            .Atacable = Val(GetVar(L, "DECOR" & Decor, "Atacable"))
            .clave = Val(GetVar(L, "DECOR" & Decor, "Clave"))
            .CantObjetos = Val(GetVar(L, "DECOR" & Decor, "CantObjetos"))
            
            If .CantObjetos > 0 Then
               ReDim .Objeto(1 To .CantObjetos) As tDecorDrop
               
               For LoopD = 1 To .CantObjetos

                  tmpStr2 = GetVar(L, "DECOR" & Decor, "Objeto" & LoopD)
                  .Objeto(LoopD).ObjIndex = Val(Readfield(1, tmpStr2, 45))
                  .Objeto(LoopD).DropChance = Val(Readfield(2, tmpStr2, 45))
                  
               Next LoopD
               
            End If
            
            .EstadoDefault = Val(GetVar(L, "DECOR" & Decor, "EstadoDefault"))
            .TileH = Val(GetVar(L, "DECOR" & Decor, "TileH"))
            .TileW = Val(GetVar(L, "DECOR" & Decor, "TileW"))
            .OffX = Val(GetVar(L, "DECOR" & Decor, "OffX"))
            .OffY = Val(GetVar(L, "DECOR" & Decor, "OffY"))
            .TileTransY = Val(GetVar(L, "DECOR" & Decor, "TileTransY"))
            .Sombra = Val(GetVar(L, "DECOR" & Decor, "NoSombra"))
            
               
         
         End With
         
      Next Decor
      

      Exit Sub
      
ErrHandler:
   MsgBox "ERROR DECORDATA: " & Err.Description
End Sub
Public Sub LoadCraft()
Dim K As Long
Dim s As String

'Cargamos herreria, sastreria y carpinteria.
s = App.path & "\ENCODE\Herreria.dat"

NumHerr = Val(GetVar(s, "INIT", "NUM"))
ReDim cHerreria(1 To NumHerr)
For K = 1 To NumHerr
    With cHerreria(K)
        .Item = Val(GetVar(s, K, "ITEM"))
        .Tipo = Val(GetVar(s, K, "TIPO"))
        .ProfesionNivel = Val(GetVar(s, K, "NIVEL"))
        .Mat1 = Val(GetVar(s, K, "BRONCE"))
        .Mat2 = Val(GetVar(s, K, "PLATA"))
        .Mat3 = Val(GetVar(s, K, "ORO"))
        .Version = Val(GetVar(s, K, "VER"))
    End With
Next K
s = App.path & "\ENCODE\Sastreria.dat"
NumSastr = Val(GetVar(s, "INIT", "NUM"))
ReDim cSastreria(1 To NumSastr)
For K = 1 To NumSastr
    With cSastreria(K)
        .Item = Val(GetVar(s, K, "ITEM"))
        .Tipo = Val(GetVar(s, K, "TIPO"))
        .ProfesionNivel = Val(GetVar(s, K, "NIVEL"))
        .Mat1 = Val(GetVar(s, K, "PIEL1"))
        .Mat2 = Val(GetVar(s, K, "PIEL2"))
        .Mat3 = Val(GetVar(s, K, "PIEL3"))
        .Version = Val(GetVar(s, K, "VER"))
    End With
Next K
s = App.path & "\ENCODE\Carpinteria.dat"
NumCarp = Val(GetVar(s, "INIT", "NUM"))
ReDim cCarpinteria(1 To NumCarp)
For K = 1 To NumCarp
    With cCarpinteria(K)
        .Item = Val(GetVar(s, K, "ITEM"))
        .Tipo = Val(GetVar(s, K, "TIPO"))
        .ProfesionNivel = Val(GetVar(s, K, "NIVEL"))
        .Mat1 = Val(GetVar(s, K, "Madera"))
        .Mat2 = Val(GetVar(s, K, "Madera2"))
        .Mat3 = Val(GetVar(s, K, "MARFIL"))
        .Version = Val(GetVar(s, K, "VER"))
    End With
Next K

End Sub
Public Sub LoadPremios()
Dim K As Long
Dim s As String
Dim J As Long
s = App.path & "\ENCODE\Canje.dat"



NumCanje = Val(GetVar(s, "INIT", "NumItems"))
ReDim Canjes(1 To NumCanje)
For K = 1 To NumCanje

    With Canjes(K)
    

        .Nombre = GetVar(s, "PREMIO" & K, "Nombre")
        .Info = GetVar(s, "PREMIO" & K, "Info")
        .Descuento = Val(GetVar(s, "PREMIO" & K, "Descuento"))
        .vGema = Val(GetVar(s, "PREMIO" & K, "Valor_Gemas"))
        .vMM = Val(GetVar(s, "PREMIO" & K, "Valor_MM"))
        .nItems = Val(GetVar(s, "PREMIO" & K, "NumObjs"))
        .Tipo = Val(GetVar(s, "PREMIO" & K, "Tipo"))
        .Version = Val(GetVar(s, "PREMIO" & K, "Version"))
        ReDim .Cant(1 To .nItems)
        ReDim .Items(1 To .nItems)
        For J = 1 To .nItems
            .Items(J) = Val(Readfield(1, GetVar(s, "PREMIO" & K, "OBJ" & J), Asc("-")))
            .Cant(J) = Val(Readfield(2, GetVar(s, "PREMIO" & K, "OBJ" & J), Asc("-")))
        Next J



    End With
    
Next K

End Sub
Sub EscribirAuras(ByVal FF As Integer)
      Dim i As Long
   On Error GoTo EscribirAuras_Error

2         Put FF, , nAura
4         For i = 1 To nAura
6             Put FF, , AuraDATA(i).Tipo
8             Put FF, , AuraDATA(i).grhindex
10            Put FF, , AuraDATA(i).Giratoria
12            Put FF, , AuraDATA(i).Velocidad
14            Put FF, , AuraDATA(i).OffsetX
16            Put FF, , AuraDATA(i).OffsetY
18            Put FF, , AuraDATA(i).A
20            Put FF, , AuraDATA(i).r
22            Put FF, , AuraDATA(i).g
24            Put FF, , AuraDATA(i).b
          
          
26        Next i

    Exit Sub

EscribirAuras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EscribirAuras in line:" & Erl

End Sub
Sub CargarAuras() 'sub de testeo
      Dim path As String
      Dim i As Long
   On Error GoTo CargarAuras_Error

2     path = App.path & "\ENCODE\AURA.DAT"

4     nAura = Val(GetVar(path, "INIT", "NumAuras"))

6     ReDim AuraDATA(0 To nAura) As tAura

8     For i = 1 To nAura

10        With AuraDATA(i)
12            .grhindex = Val(GetVar(path, "Aura" & i, "GrhIndex"))
14            .r = Val(GetVar(path, "Aura" & i, "Rojo"))
16            .g = Val(GetVar(path, "Aura" & i, "Verde"))
18            .b = Val(GetVar(path, "Aura" & i, "Azul"))
20            .Giratoria = Val(GetVar(path, "Aura" & i, "Giratoria"))
22            .OffsetX = Val(GetVar(path, "Aura" & i, "OffsetX"))
24            .OffsetY = Val(GetVar(path, "Aura" & i, "Offset"))
26            .A = Val(GetVar(path, "Aura" & i, "Alpha"))
28            .Velocidad = Val(GetVar(path, "Aura" & i, "Vel"))
30        End With

32    Next i

    Exit Sub

CargarAuras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarAuras in line:" & Erl

End Sub
Sub CargarParticulas() 'sub de testeo
       
      Dim StreamFile As String
      Dim LoopC As Long
      Dim i As Long
      Dim GrhListing As String
      Dim TempSet As String
      Dim ColorSet As Long

   On Error GoTo CargarParticulas_Error

2     StreamFile = App.path & "\ENCODE\Particles.ini"
4     TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
       
      'resize StreamData array
6     ReDim StreamData(1 To TotalStreams) As Stream
       
          'fill StreamData array with info from Particles.ini
8         For LoopC = 1 To TotalStreams
10            StreamData(LoopC).Name = GetVar(StreamFile, Val(LoopC), "Name")
12            StreamData(LoopC).NumOfParticles = GetVar(StreamFile, Val(LoopC), "NumOfParticles")
14            StreamData(LoopC).NumTrueParticles = StreamData(LoopC).NumOfParticles
              
16            StreamData(LoopC).X1 = GetVar(StreamFile, Val(LoopC), "X1")
18            StreamData(LoopC).y1 = GetVar(StreamFile, Val(LoopC), "Y1")
20            StreamData(LoopC).x2 = GetVar(StreamFile, Val(LoopC), "X2")
22            StreamData(LoopC).y2 = GetVar(StreamFile, Val(LoopC), "Y2")
24            StreamData(LoopC).Angle = GetVar(StreamFile, Val(LoopC), "Angle")
26            StreamData(LoopC).vecx1 = GetVar(StreamFile, Val(LoopC), "VecX1")
28            StreamData(LoopC).vecx2 = GetVar(StreamFile, Val(LoopC), "VecX2")
30            StreamData(LoopC).vecy1 = GetVar(StreamFile, Val(LoopC), "VecY1")
32            StreamData(LoopC).vecy2 = GetVar(StreamFile, Val(LoopC), "VecY2")
34            StreamData(LoopC).life1 = GetVar(StreamFile, Val(LoopC), "Life1")
36            StreamData(LoopC).life2 = GetVar(StreamFile, Val(LoopC), "Life2")
38            StreamData(LoopC).friction = GetVar(StreamFile, Val(LoopC), "Friction")
40            StreamData(LoopC).Spin = GetVar(StreamFile, Val(LoopC), "Spin")
42            StreamData(LoopC).spin_speedL = GetVar(StreamFile, Val(LoopC), "Spin_SpeedL")
44            StreamData(LoopC).spin_speedH = GetVar(StreamFile, Val(LoopC), "Spin_SpeedH")
46            StreamData(LoopC).AlphaBlend = GetVar(StreamFile, Val(LoopC), "AlphaBlend")
48            StreamData(LoopC).gravity = GetVar(StreamFile, Val(LoopC), "Gravity")
50            StreamData(LoopC).grav_strength = GetVar(StreamFile, Val(LoopC), "Grav_Strength")
52            StreamData(LoopC).bounce_strength = GetVar(StreamFile, Val(LoopC), "Bounce_Strength")
54            StreamData(LoopC).XMove = GetVar(StreamFile, Val(LoopC), "XMove")
56            StreamData(LoopC).YMove = GetVar(StreamFile, Val(LoopC), "YMove")
58            StreamData(LoopC).move_x1 = GetVar(StreamFile, Val(LoopC), "move_x1")
60            StreamData(LoopC).move_x2 = GetVar(StreamFile, Val(LoopC), "move_x2")
62            StreamData(LoopC).move_y1 = GetVar(StreamFile, Val(LoopC), "move_y1")
64            StreamData(LoopC).move_y2 = GetVar(StreamFile, Val(LoopC), "move_y2")
66            StreamData(LoopC).life_counter = GetVar(StreamFile, Val(LoopC), "life_counter")
68            StreamData(LoopC).Speed = Val(GetVar(StreamFile, Val(LoopC), "Speed"))
70            StreamData(LoopC).grh_resize = Val(GetVar(StreamFile, Val(LoopC), "resize"))
72            StreamData(LoopC).grh_resizex = Val(GetVar(StreamFile, Val(LoopC), "rx"))
74            StreamData(LoopC).grh_resizey = Val(GetVar(StreamFile, Val(LoopC), "ry"))
76            StreamData(LoopC).NumGrhs = GetVar(StreamFile, Val(LoopC), "NumGrhs")
             
78            ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
80            GrhListing = GetVar(StreamFile, Val(LoopC), "Grh_List")
             
82            For i = 1 To StreamData(LoopC).NumGrhs
84                StreamData(LoopC).grh_list(i) = Readfield(Str(i), GrhListing, 44)
86            Next i
88            StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
90            For ColorSet = 1 To 4
92                TempSet = GetVar(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
94                StreamData(LoopC).colortint(ColorSet - 1).r = Readfield(1, TempSet, 44)
96                StreamData(LoopC).colortint(ColorSet - 1).g = Readfield(2, TempSet, 44)
98                StreamData(LoopC).colortint(ColorSet - 1).b = Readfield(3, TempSet, 44)
100           Next ColorSet
102       Next LoopC

    Exit Sub

CargarParticulas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarParticulas in line:" & Erl & " en particula : " & LoopC
       
End Sub
Public Sub EscribirParticulas(ByVal FF As Integer)
      Dim LoopC As Long
      Dim i As Long
      Dim ColorSet As Long


2     Put FF, , TotalStreams
       
          'fill StreamData array with info from Particles.ini
4         For LoopC = 1 To TotalStreams
6             Put FF, , StreamData(LoopC).NumOfParticles
              
8             Put FF, , StreamData(LoopC).X1
10            Put FF, , StreamData(LoopC).y1 '= GetVar(StreamFile, Val(loopc), "Y1")
12            Put FF, , StreamData(LoopC).x2 '= GetVar(StreamFile, Val(loopc), "X2")
14            Put FF, , StreamData(LoopC).y2 '= GetVar(StreamFile, Val(loopc), "Y2")
16            Put FF, , StreamData(LoopC).Angle '= GetVar(StreamFile, Val(loopc), "Angle")
18            Put FF, , StreamData(LoopC).vecx1 '= GetVar(StreamFile, Val(loopc), "VecX1")
20            Put FF, , StreamData(LoopC).vecx2 '= GetVar(StreamFile, Val(loopc), "VecX2")
22            Put FF, , StreamData(LoopC).vecy1 '= GetVar(StreamFile, Val(loopc), "VecY1")
24            Put FF, , StreamData(LoopC).vecy2 '= GetVar(StreamFile, Val(loopc), "VecY2")
26            Put FF, , StreamData(LoopC).life1 '= GetVar(StreamFile, Val(loopc), "Life1")
28            Put FF, , StreamData(LoopC).life2 '= GetVar(StreamFile, Val(loopc), "Life2")
30            Put FF, , StreamData(LoopC).friction '= GetVar(StreamFile, Val(loopc), "Friction")
32            Put FF, , StreamData(LoopC).Spin '= GetVar(StreamFile, Val(loopc), "Spin")
34            Put FF, , StreamData(LoopC).spin_speedL '= GetVar(StreamFile, Val(loopc), "Spin_SpeedL")
36            Put FF, , StreamData(LoopC).spin_speedH '= GetVar(StreamFile, Val(loopc), "Spin_SpeedH")
38            Put FF, , StreamData(LoopC).AlphaBlend '= GetVar(StreamFile, Val(loopc), "AlphaBlend")
40            Put FF, , StreamData(LoopC).gravity '= GetVar(StreamFile, Val(loopc), "Gravity")
42            Put FF, , StreamData(LoopC).grav_strength '= GetVar(StreamFile, Val(loopc), "Grav_Strength")
44            Put FF, , StreamData(LoopC).bounce_strength '= GetVar(StreamFile, Val(loopc), "Bounce_Strength")
46            Put FF, , StreamData(LoopC).XMove '= GetVar(StreamFile, Val(loopc), "XMove")
48            Put FF, , StreamData(LoopC).YMove '= GetVar(StreamFile, Val(loopc), "YMove")
50            Put FF, , StreamData(LoopC).move_x1 '= GetVar(StreamFile, Val(loopc), "move_x1")
52            Put FF, , StreamData(LoopC).move_x2 '= GetVar(StreamFile, Val(loopc), "move_x2")
54            Put FF, , StreamData(LoopC).move_y1 '= GetVar(StreamFile, Val(loopc), "move_y1")
56            Put FF, , StreamData(LoopC).move_y2 '= GetVar(StreamFile, Val(loopc), "move_y2")
58            Put FF, , StreamData(LoopC).life_counter '= GetVar(StreamFile, Val(loopc), "life_counter")
60            Put FF, , StreamData(LoopC).Speed '= Val(GetVar(StreamFile, Val(loopc), "Speed"))
62            Put FF, , StreamData(LoopC).grh_resize '= Val(GetVar(StreamFile, Val(loopc), "resize"))
64            Put FF, , StreamData(LoopC).grh_resizex '= Val(GetVar(StreamFile, Val(loopc), "rx"))
66            Put FF, , StreamData(LoopC).grh_resizey '= Val(GetVar(StreamFile, Val(loopc), "ry"))
68            Put FF, , StreamData(LoopC).NumGrhs '= GetVar(StreamFile, Val(loopc), "NumGrhs")
             
            
70           For i = 1 To StreamData(LoopC).NumGrhs
72               Put FF, , StreamData(LoopC).grh_list(i)
74           Next i
             
76            For ColorSet = 1 To 4
78                Put FF, , StreamData(LoopC).colortint(ColorSet - 1).r
80                Put FF, , StreamData(LoopC).colortint(ColorSet - 1).g
82                Put FF, , StreamData(LoopC).colortint(ColorSet - 1).b
84            Next ColorSet
86        Next LoopC

End Sub
Public Sub EscribirBuffdataBin(ByVal FF As Integer)

      Dim i As Long

2         Put FF, , numero_buffs
          
4         For i = 1 To numero_buffs

6             Put FF, , buff_data(i).Tipo
8             Put FF, , buff_data(i).Intervalo
10            Put FF, , buff_data(i).Duracion
12            Put FF, , buff_data(i).dFX
14            Put FF, , buff_data(i).dEfecto
16            Put FF, , buff_data(i).grhindex
          
18        Next i
End Sub
Public Sub CargarBuffData()
      Dim F As String
      Dim i As Long
2         F = App.path & "\ENCODE\Buffs.dat"
4         numero_buffs = Val(GetVar(F, "MAIN", "NUMBUFFS"))
6         ReDim buff_data(1 To numero_buffs)
8         For i = 1 To numero_buffs
              
10            buff_data(i).Tipo = Val(GetVar(F, CStr(i), "Tipo"))
12            buff_data(i).Intervalo = Val(GetVar(F, CStr(i), "Intervalo"))
14            buff_data(i).Duracion = Val(GetVar(F, CStr(i), "Duracion"))
16            buff_data(i).dFX = Val(GetVar(F, CStr(i), "Fx"))
18            buff_data(i).dEfecto = Val(GetVar(F, CStr(i), "Efecto"))
20            buff_data(i).grhindex = Val(GetVar(F, CStr(i), "GrhIndex"))
              
22        Next i

End Sub
Public Sub SPOTLIGHTS_Escribir(ByVal FF As Integer)
          'CARGA BINARIA DE SPOTLIGHTS DESDE EFECTOS.BIN
      Dim i As Long
2         Put FF, , NUM_SPOTLIGHTS_COLORES
4         For i = 1 To NUM_SPOTLIGHTS_COLORES
6             Put #FF, , SPOTLIGHTS_COLORES(i)
8         Next i
10        Put FF, , NUM_SPOTLIGHTS_ANIMATION
12        For i = 1 To NUM_SPOTLIGHTS_ANIMATION
14            Put #FF, , SPOTLIGHTS_ANIMATION(i)
16        Next i
End Sub
Public Sub SPOTLIGHTS_LOADDAT() 'sub de testeo
      Dim s As String
      Dim i As Long
      Dim A As Byte
      Dim r As Byte
      Dim g As Byte
      Dim b As Byte
2     s = App.path & "\ENCODE\SPOTLIGHTS.DAT"
4     NUM_SPOTLIGHTS_COLORES = Val(GetVar(s, "COLORES", "NUM_COLORES"))
6     If NUM_SPOTLIGHTS_COLORES > 0 Then
8     ReDim SPOTLIGHTS_COLORES(1 To NUM_SPOTLIGHTS_COLORES)
10        For i = 1 To NUM_SPOTLIGHTS_COLORES
12            A = Val(GetVar(s, "COLOR" & i, "A"))
14            r = Val(GetVar(s, "COLOR" & i, "R"))
16            g = Val(GetVar(s, "COLOR" & i, "G"))
18            b = Val(GetVar(s, "COLOR" & i, "B"))
20            SPOTLIGHTS_COLORES(i) = D3DColorARGB(A, r, g, b)
22        Next i
24    End If

26    NUM_SPOTLIGHTS_ANIMATION = Val(GetVar(s, "ANIMACIONES", "NUM_ANIM"))
28    If NUM_SPOTLIGHTS_ANIMATION > 0 Then
30    ReDim SPOTLIGHTS_ANIMATION(1 To NUM_SPOTLIGHTS_ANIMATION)
32    For i = 1 To NUM_SPOTLIGHTS_ANIMATION
34        SPOTLIGHTS_ANIMATION(i) = Val(GetVar(s, "ANIM" & i, "Indice"))
36    Next i
38    End If

End Sub

