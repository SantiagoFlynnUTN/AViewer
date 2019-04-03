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
    R As Byte
    G As Byte
    B As Byte
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
Dim S As String
Dim i As Long
Dim z As Long

S = App.PATH & "\RES\Index\NewEstatics.dat"

NumEstatics = Val(GetVar(S, "INIT", "num"))

If NumEstatics > 0 Then
ReDim EstaticData(1 To NumEstatics)
For i = 1 To NumEstatics
    With EstaticData(i)
        .L = Val(GetVar(S, CStr(i), "Left"))
        .t = Val(GetVar(S, CStr(i), "Top"))
        .w = Val(GetVar(S, CStr(i), "Width"))
        .h = Val(GetVar(S, CStr(i), "Height"))
        .TW = .w / 32
        .TH = .h / 32
    End With
Next i

Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwEstatics.IND" For Binary Access Write Lock Write As K
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
Dim S As String
Dim i As Long
Dim z As Long

S = App.PATH & "\RES\Index\NewIndex.dat"

NumNewIndex = Val(GetVar(S, "INIT", "num"))

If NumNewIndex > 0 Then
ReDim NewIndexData(1 To NumNewIndex)
For i = 1 To NumNewIndex
    With NewIndexData(i)
        .Dinamica = Val(GetVar(S, CStr(i), "Dinamica"))
        .Estatic = Val(GetVar(S, CStr(i), "Estatica"))
        .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
    End With
Next i


Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwIndex.IND" For Binary Access Write Lock Write As K
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
ns = Val(GetVar(App.PATH & "\ENCODE\hechizos.dat", "INIT", "NumeroHechizos"))
ReDim Spells(1 To ns)
Dim F As String
F = App.PATH & "\ENCODE\hechizos.dat"
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
Dim S As String
Dim i As Long
Dim p As Long
S = App.PATH & "\ENCODE\Nobleza.dat"


Nobleza_Data.NumItems = Val(GetVar(S, "INIT", "NUM"))

ReDim Nobleza_Data.Items(1 To Nobleza_Data.NumItems)



For i = 1 To Nobleza_Data.NumItems

With Nobleza_Data.Items(i)

    .NumItems_Requeridos = Val(GetVar(S, "OBJ" & i, "CantItem"))
    .Numero = Val(GetVar(S, "OBJ" & i, "ObjIndexRecompensa"))
    ReDim .Items_Requeridos(1 To .NumItems_Requeridos)
    ReDim .cantItems_Requeridos(1 To .NumItems_Requeridos)
    For p = 1 To .NumItems_Requeridos
        .Items_Requeridos(p) = Val(GetVar(S, "OBJ" & i, "ObjIndex" & p))
        .cantItems_Requeridos(p) = Val(GetVar(S, "OBJ" & i, "Cantidad" & p))
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
      EluTable(i) = CLng(GetVar(App.PATH & "\ENCODE\Tables.dat", "EluTable", CStr(i)))
   Next i
   
   For i = 1 To STAT_MAXELV
      StaTable(i) = CInt(GetVar(App.PATH & "\ENCODE\Tables.dat", "StaTable", CStr(i)))
   Next i
   
      
End Sub

Sub Load_FD()

Dim S As String
Dim i As Long
Dim X As Long
S = App.PATH & "\ENCODE\FORMULARIO_DATA.dat"

NUM_FD = Val(GetVar(S, "INIT", "NUM"))

ReDim FD(1 To NUM_FD)

For i = 1 To NUM_FD

With FD(i)
        .TieneSpecial = IIf(Val(GetVar(S, "FD" & i, "Special")) = 1, True, False)
    .SurfaceNum = Val(GetVar(S, "FD" & i, "Grafico"))
    .ScreenX = Val(GetVar(S, "FD" & i, "X"))
    .ScreenY = Val(GetVar(S, "FD" & i, "Y"))
    
    .GraficosX = Val(GetVar(S, "FD" & i, "GRAFICOSENX"))
    .GraficosY = Val(GetVar(S, "FD" & i, "GRAFICOSENY"))
    If .GraficosX = 0 Then .GraficosX = 1
    If .GraficosY = 0 Then .GraficosY = 1
    .Num_Checks = Val(GetVar(S, "FD" & i, "NumChecks"))
    .Draw_Stage = Val(GetVar(S, "FD" & i, "NivelDibujo"))
    .num_Buttons = Val(GetVar(S, "FD" & i, "NumeroBotones"))
    .AlphaValue = Val(GetVar(S, "FD" & i, "AlphaValue"))
    .Num_Textos = Val(GetVar(S, "FD" & i, "NumeroTextos"))
    .Num_TextBox = Val(GetVar(S, "FD" & i, "NumTextBox"))
    .Height = Val(GetVar(S, "FD" & i, "Height"))
    .Width = Val(GetVar(S, "FD" & i, "Width"))
    
    If .AlphaValue = 0 Then .AlphaValue = 255
    If .num_Buttons > 0 Then
    ReDim .Buttons(1 To .num_Buttons)
    
    For X = 1 To .num_Buttons
        
        With FD(i).Buttons(X)
        
            .NormalSurfaceNum = Val(GetVar(S, "FD" & i, "Btn_" & X & "_NormalGrafico"))
            .SelectSurfaceNum = Val(GetVar(S, "FD" & i, "Btn_" & X & "_SelectGrafico"))
            .PressSurfaceNum = Val(GetVar(S, "FD" & i, "Btn_" & X & "_PressGrafico"))
            
            .size.Top = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Top"))
            .size.Left = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Left"))
            .size.Right = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Width"))
            .size.bottom = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Height"))
            
            .Sound = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Sound"))
            .HandIco = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Hand"))
            .Caption = Val(GetVar(S, "FD" & i, "Btn_" & X & "_Caption"))
            
            .Normal_Rojo = Val(GetVar(S, "FD" & i, "Btn_" & X & "_NormalRojo"))
            .Normal_Verde = Val(GetVar(S, "FD" & i, "Btn_" & X & "_NormalVerde"))
            .Normal_Azul = Val(GetVar(S, "FD" & i, "Btn_" & X & "_NormalAzul"))
                        
            .Sel_Rojo = Val(GetVar(S, "FD" & i, "Btn_" & X & "_SelRojo"))
            .Sel_Verde = Val(GetVar(S, "FD" & i, "Btn_" & X & "_SelVerde"))
            .Sel_Azul = Val(GetVar(S, "FD" & i, "Btn_" & X & "_SelAzul"))
            
            .Press_Rojo = Val(GetVar(S, "FD" & i, "Btn_" & X & "_PressRojo"))
            .Press_Verde = Val(GetVar(S, "FD" & i, "Btn_" & X & "_PressVerde"))
            .Press_Azul = Val(GetVar(S, "FD" & i, "Btn_" & X & "_PressAzul"))
        End With
    Next X
    End If
    If .Num_Textos > 0 Then
    ReDim FD(i).Textos(1 To .Num_Textos)
    For X = 1 To .Num_Textos
    
        With FD(i).Textos(X)
            .texto = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Texto"))
            
            .X = Val(GetVar(S, "FD" & i, "Txt_" & X & "_X"))
            
            .Y = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Y"))
            
            .MaxWidth = Val(GetVar(S, "FD" & i, "Txt_" & X & "_MaxWidth"))
            .IniciaVisible = Val(GetVar(S, "FD" & i, "Txt_" & X & "_IniciaVisible"))
            
            
            .A = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Alpha"))
            
            .R = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Rojo"))
            
            .G = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Verde"))
                    
            
            .B = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Azul"))
            
            .Centrar = Val(GetVar(S, "FD" & i, "Txt_" & X & "_Centrar"))
                    
        End With
    Next X
    End If
    If .Num_TextBox > 0 Then
        ReDim .TextBox(1 To .Num_TextBox)
        
        For X = 1 To .Num_TextBox
            .TextBox(X).SurfaceNum = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_Surface"))
            .TextBox(X).Centrar = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_Centrar"))
            .TextBox(X).TipoTexto = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_TipoTexto"))
            .TextBox(X).X = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_x"))
            .TextBox(X).Y = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_y"))
            .TextBox(X).OffsetX = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_Offsetx"))
            .TextBox(X).OffsetY = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_Offsety"))
            .TextBox(X).IniciaVisible = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_IniciaVisible"))
            .TextBox(X).fA = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_AFont"))
            .TextBox(X).fR = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_RFont"))
            .TextBox(X).fG = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_GFont"))
            .TextBox(X).fb = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_BFont"))
            .TextBox(X).w = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_W"))
            .TextBox(X).h = Val(GetVar(S, "FD" & i, "TxtB_" & X & "_H"))
            
        Next X
    End If
    
        If .Num_Checks > 0 Then
        ReDim .Checks(1 To .Num_Checks)
        For X = 1 To .Num_Checks
            .Checks(X).Caption = Val(GetVar(S, "FD" & i, "Chk_" & X & "_Caption"))
            .Checks(X).X = Val(GetVar(S, "FD" & i, "Chk_" & X & "_X"))
            .Checks(X).Y = Val(GetVar(S, "FD" & i, "Chk_" & X & "_Y"))
            .Checks(X).w = Val(GetVar(S, "FD" & i, "Chk_" & X & "_W"))
            .Checks(X).h = Val(GetVar(S, "FD" & i, "Chk_" & X & "_H"))
            .Checks(X).min = Val(GetVar(S, "FD" & i, "Chk_" & X & "_MIN"))
            .Checks(X).max = Val(GetVar(S, "FD" & i, "Chk_" & X & "_MAX"))
            .Checks(X).Tipo_Check = Val(GetVar(S, "FD" & i, "Chk_" & X & "_TipoCheck"))
            .Checks(X).CheckSurface = Val(GetVar(S, "FD" & i, "Chk_" & X & "_CheckSurface"))
            .Checks(X).SurfaceNum = Val(GetVar(S, "FD" & i, "Chk_" & X & "_Grafico"))
            .Checks(X).IniciaVisible = Val(GetVar(S, "FD" & i, "Chk_" & X & "_IniciaVisible"))
            .Checks(X).ColorA = Val(GetVar(S, "FD" & i, "Chk_" & X & "_A"))
            .Checks(X).ColorR = Val(GetVar(S, "FD" & i, "Chk_" & X & "_R"))
            .Checks(X).ColorG = Val(GetVar(S, "FD" & i, "Chk_" & X & "_G"))
            .Checks(X).ColorB = Val(GetVar(S, "FD" & i, "Chk_" & X & "_B"))
        Next X
    End If
    
    
End With

Next i



End Sub
Public Sub Load_HabilidadesData()

Dim i As Long



For i = 1 To MAX_HABILIDADES

    Habilidades(i).Grafico = Val(GetVar(App.PATH & "\ENCODE\Habilidades.dat", "HAB" & i, "Grafico"))

Next i

End Sub
Public Sub Load_NpcnoHostiles()

Dim S As String
Dim i As Long
S = App.PATH & "\ENCODE\npcs-hostiles.dat"

num_npcs_h = Val(GetVar(S, "init", "numnpcs")) + 1

ReDim hostiles(0 To num_npcs_h)
For i = 500 To num_npcs_h
With hostiles(i - 499)

    .Body = Val(GetVar(S, "NPC" & i, "body"))
    .Head = Val(GetVar(S, "NPC" & i, "Head"))
    .MAX_HP = Val(GetVar(S, "NPC" & i, "MaxHP"))
    .Snd1 = Val(GetVar(S, "NPC" & i, "Snd1"))
    .Snd2 = Val(GetVar(S, "NPC" & i, "Snd2"))
    .Nombre = GetVar(S, "NPC" & i, "Name")
    
    
End With
Next i
End Sub
Public Sub Load_NpcHostiles()

Dim S As String
Dim i As Long
S = App.PATH & "\ENCODE\npcs.dat"

num_npcs_nh = Val(GetVar(S, "init", "numnpcs"))

ReDim nHostiles(1 To num_npcs_nh)
For i = 1 To num_npcs_nh
With nHostiles(i)


    
        .Body = Val(GetVar(S, "NPC" & i, "body"))
    .Head = Val(GetVar(S, "NPC" & i, "Head"))
    .MAX_HP = Val(GetVar(S, "NPC" & i, "MaxHP"))
    .Nombre = GetVar(S, "NPC" & i, "Name")
    .Desc = GetVar(S, "NPC" & i, "Desc")
    .NPCTYPE = Val(GetVar(S, "NPC" & i, "NpcType"))
    
End With
Next i
End Sub
Public Sub LoadQuest()
Dim S As String
Dim i As Long
Dim z As Long
S = App.PATH & "\ENCODE\Quests.dat"

nQuest = Val(GetVar(S, "MAIN", "NUMQUESTS"))
ReDim Quests(1 To nQuest)

For i = 1 To nQuest

    Quests(i).Tipo = Val(GetVar(S, CStr(i), "Tipo"))
    
    Quests(i).nTargets = Val(GetVar(S, CStr(i), "numero_targets"))
    Quests(i).Oro = Val(GetVar(S, CStr(i), "recompensa_oro"))
    Quests(i).Puntos = Val(GetVar(S, CStr(i), "recompensa_puntos"))
    Quests(i).Exp = Val(GetVar(S, CStr(i), "recompensa_exp"))
    Quests(i).numritems = Val(GetVar(S, CStr(i), "recompensa_numero_items"))
    Quests(i).TipoReco = Val(GetVar(S, CStr(i), "recompensa_tipo"))
    Quests(i).Nombre = GetVar(S, CStr(i), "nombre")
    
    
    If Quests(i).numritems > 0 Then
        ReDim Quests(i).Item(1 To Quests(i).numritems)
        ReDim Quests(i).Cant(1 To Quests(i).numritems)
        For z = 1 To Quests(i).numritems
            Quests(i).Item(z) = Val(GetVar(S, CStr(i), "recompensa_item_tipo" & z))
            Quests(i).Cant(z) = Val(GetVar(S, CStr(i), "recompensa_item_cant" & z))
            
        
        Next z
    End If
    Quests(i).Desc = GetVar(S, CStr(i), "desc")
    
    If Quests(i).nTargets > 0 Then
    ReDim Quests(i).Targets(1 To Quests(i).nTargets)
    ReDim Quests(i).TargetsCant(1 To Quests(i).nTargets)
    For z = 1 To Quests(i).nTargets
        Quests(i).Targets(z) = Val(GetVar(S, CStr(i), "target_tipo" & z))
        Quests(i).TargetsCant(z) = Val(GetVar(S, CStr(i), "target_cant" & z))
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
Dim S As String
F = FreeFile

Open App.PATH & "\ENCODE\Mapadata.txt" For Input As #F
    Line Input #F, S
    
    NroMapas = Val(S)
    
    ReDim MapaData(1 To NroMapas)
Do Until EOF(F)
i = i + 1

    Line Input #F, S
    
    MapaData(i).X = Val(Readfield(1, S, 44))
    MapaData(i).Y = Val(Readfield(2, S, 44))
    
Loop
Close #F


End Sub
Function Readfield(ByVal pos As Integer, ByVal texto As String, ByVal separador As Byte) As String
'caserita
Dim i As Long
Dim t As Long
Dim L As Long
Dim K As Long
Dim S As String

L = Len(texto)


Do Until i = L
    i = i + 1
    If Asc(mid$(texto, i, 1)) = separador Then
        K = K + 1
        If K = pos Then
            Readfield = S
            Exit Do
        Else
            S = vbNullString
        End If
    Else
        S = S & mid$(texto, i, 1)
    End If
Loop
Readfield = S
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
    Open App.PATH & "\RES\OUTPUT\INDEX.BIN" For Binary Access Write Lock Write As K
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


Dim PATH As String

Dim i As Integer

Dim UltimoS As Long
Dim UltimoL As Long
ReDim IXAR(1 To 11)

PATH = App.PATH & "\RES\INDEX\"
UltimoS = (8 * 11) - 1


i = FreeFile
Open PATH & "NewFx.BIN" For Binary Access Read Lock Read As i
    IXAR(1).Le = LOF(i)
    ReDim IXAR(1).Da(0 To IXAR(1).Le - 1)
    Get i, , IXAR(1).Da
Close i

IXAR(1).St = UltimoS + UltimoL
UltimoS = IXAR(1).St
UltimoL = IXAR(1).Le

i = FreeFile
Open PATH & "NwAnim.IND" For Binary Access Read Lock Read As i
    IXAR(2).Le = LOF(i)
    ReDim IXAR(2).Da(0 To IXAR(2).Le - 1)
    Get i, , IXAR(2).Da
Close i

IXAR(2).St = UltimoS + UltimoL
UltimoS = IXAR(2).St
UltimoL = IXAR(2).Le

i = FreeFile
Open PATH & "NwBody.IND" For Binary Access Read Lock Read As i
    IXAR(3).Le = LOF(i)
    ReDim IXAR(3).Da(0 To IXAR(3).Le - 1)
    Get i, , IXAR(3).Da
Close i

IXAR(3).St = UltimoS + UltimoL
UltimoS = IXAR(3).St
UltimoL = IXAR(3).Le

i = FreeFile
Open PATH & "NwShields.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwWeapons.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwIndex.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwEstatics.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwHelmets.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwMunicion.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwCapas.IND" For Binary Access Read Lock Read As i
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
Open PATH & "NwHeads.IND" For Binary Access Read Lock Read As i
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

Dim S As String
Dim i As Long
Dim p As Long
Dim K As Long
Dim GrafCounter As Integer
Dim num_nwanim As Integer

S = App.PATH & "\RES\INDEX\NewAnim.dat"


num_nwanim = Val(GetVar(S, "NW_ANIM", "NUM"))

If num_nwanim < 1 Then Exit Sub

ReDim NewAnimationData(1 To num_nwanim)

For i = 1 To num_nwanim

With NewAnimationData(i)
    .Grafico = Val(GetVar(S, "ANIMACION" & i, "Grafico"))
    .Columnas = Val(GetVar(S, "ANIMACION" & i, "Columnas"))
    .Filas = Val(GetVar(S, "ANIMACION" & i, "Filas"))
    .Height = Val(GetVar(S, "ANIMACION" & i, "Alto"))
    .Width = Val(GetVar(S, "ANIMACION" & i, "Ancho"))
    .NumFrames = Val(GetVar(S, "ANIMACION" & i, "NumeroFrames"))
    .Velocidad = Val(GetVar(S, "ANIMACION" & i, "Velocidad"))
    .TileWidth = .Width / 32
    .TileHeight = .Height / 32
    .Romboidal = Val(GetVar(S, "ANIMACION" & i, "AnimacionRomboidal"))
    .OffsetX = Val(GetVar(S, "ANIMACION" & i, "OffsetX"))
    .OffsetY = Val(GetVar(S, "ANIMACION" & i, "OffsetY"))
    .Initial = Val(GetVar(S, "ANIMACION" & i, "Inicial"))
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

Open App.PATH & "\RES\INDEX\NwAnim.IND" For Binary Access Write Lock Write As z
    
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

Dim S As String
Dim i As Long
Dim z As Long
Dim NumNewShields As Integer
S = App.PATH & "\RES\INDEX\Nwshields.dat"

NumNewShields = Val(GetVar(S, "INIT", "num"))

If NumNewShields > 0 Then
ReDim nShieldDATA(1 To NumNewShields)
For i = 1 To NumNewShields

    With nShieldDATA(i)

        .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
        .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i



Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwShields.IND" For Binary Access Write Lock Write As K
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
    
Dim S As String
Dim i As Long
Dim z As Long

S = App.PATH & "\RES\INDEX\NewHelmets.dat"

NumNewHelmet = Val(GetVar(S, "INIT", "num"))

If NumNewHelmet > 0 Then
ReDim NHelmetData(1 To NumNewHelmet)
For i = 1 To NumNewHelmet

    With NHelmetData(i)

        .Alpha = Val(GetVar(S, "HELMET" & CStr(i), "Alpha"))
        .OffsetY = Val(GetVar(S, "HELMET" & CStr(i), "OFFSET_DIBUJO"))
        .OffsetLat = Val(GetVar(S, "HELMET" & CStr(i), "OFFSET_LAT"))
        
        .mMovimiento(1) = Val(GetVar(S, "HELMET" & CStr(i), "NORTH"))
        .mMovimiento(2) = Val(GetVar(S, "HELMET" & CStr(i), "EAST"))
                .mMovimiento(3) = Val(GetVar(S, "HELMET" & CStr(i), "SOUTH"))
                .mMovimiento(4) = Val(GetVar(S, "HELMET" & CStr(i), "WEST"))
                
       
    End With
Next i



Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwHelmets.IND" For Binary Access Write Lock Write As K
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

Dim S As String
Dim i As Long
Dim z As Long
Dim NumNewWeapons As Integer
S = App.PATH & "\RES\INDEX\NwWeapons.dat"

NumNewWeapons = Val(GetVar(S, "INIT", "num"))

If NumNewWeapons > 0 Then
ReDim nWeaponData(1 To NumNewWeapons)
For i = 1 To NumNewWeapons

    With nWeaponData(i)

        .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
        .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
        
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i



Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwWeapons.IND" For Binary Access Write Lock Write As K
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

Dim S As String
Dim i As Long
Dim z As Long
Dim NumNewM As Integer
S = App.PATH & "\RES\INDEX\NwMunicion.dat"

NumNewM = Val(GetVar(S, "INIT", "num"))

If NumNewM > 0 Then
ReDim nMunicionData(1 To NumNewM)
For i = 1 To NumNewM

    With nMunicionData(i)

        .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
        .OverWriteGrafico = Val(GetVar(S, CStr(i), "OverWriteGrafico"))
        
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i

End If

Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwMunicion.IND" For Binary Access Write Lock Write As K
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

Dim S As String
Dim i As Long
Dim z As Long
Dim NumNewM As Integer
S = App.PATH & "\RES\INDEX\NwCapa.dat"

NumNewM = Val(GetVar(S, "INIT", "num"))

If NumNewM > 0 Then
ReDim nCapaData(1 To NumNewM)
For i = 1 To NumNewM

    With nCapaData(i)

        .Alpha = Val(GetVar(S, CStr(i), "Alpha"))
        .aOverWriteGrafico = Val(GetVar(S, CStr(i), "aOverWriteGrafico"))
        .pOverWriteGrafico = Val(GetVar(S, CStr(i), "pOverWriteGrafico"))
        For z = 1 To 4
        
            .mMovimiento(z) = Val(GetVar(S, CStr(i), "Mov" & z))

        Next z

    End With
Next i


End If

Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwCapa.IND" For Binary Access Write Lock Write As K
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

Dim S As String
Dim i As Long
Dim z As Long
Dim NumNewM As Integer
S = App.PATH & "\RES\INDEX\NewHeads.dat"

NumNewM = Val(GetVar(S, "INIT", "num"))

If NumNewM > 0 Then
ReDim NHeadData(1 To NumNewM)
For i = 1 To NumNewM
    With NHeadData(i)
        .Raza = Val(GetVar(S, "HEAD" & CStr(i), "RAZA"))
        .OffsetDibujoY = Val(GetVar(S, "HEAD" & CStr(i), "OFFSET_DIBUJO"))
        .OffsetOjos = Val(GetVar(S, "HEAD" & CStr(i), "OFFSET_OJOS"))
        .Genero = Val(GetVar(S, "HEAD" & CStr(i), "GENERO"))
        .Frame(2) = Val(GetVar(S, "HEAD" & CStr(i), "EAST"))
        .Frame(1) = Val(GetVar(S, "HEAD" & CStr(i), "NORTH"))
        .Frame(3) = Val(GetVar(S, "HEAD" & CStr(i), "SOUTH"))
        .Frame(4) = Val(GetVar(S, "HEAD" & CStr(i), "WEST"))
    End With
Next i
End If

Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwHeads.IND" For Binary Access Write Lock Write As K
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

Dim S As String
Dim i As Long
Dim z As Long
Dim NumNewBodys As Integer
S = App.PATH & "\RES\INDEX\NewBody.dat"

NumNewBodys = Val(GetVar(S, "INIT", "num"))

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
        .Capa = Val(GetVar(S, CStr(i), "Capa"))
        If .bAtaque Then
            For z = 1 To 4
                .Attack(z) = Val(GetVar(S, CStr(i), "Ataque" & z))
                

            Next z
        End If
        If .bAtacado Then
            For z = 1 To 4
                .Attacked(z) = Val(GetVar(S, CStr(i), "Atacado" & z))
            Next z
        End If
        
        If .bReposo Then
            For z = 1 To 4
            .Reposo(z) = Val(GetVar(S, CStr(i), "Reposo" & z))
        

            Next z
        End If
        If .bDeath Then
            For z = 1 To 4
            .Death(z) = Val(GetVar(S, CStr(i), "Muerte" & z))
                

            Next z
        End If

        For z = 1 To 4
        
            .mMovement(z) = Val(GetVar(S, CStr(i), "Mov" & z))
        
        Next z

    End With
Next i



Dim K As Integer
K = FreeFile

Open App.PATH & "\RES\INDEX\NwBody.IND" For Binary Access Write Lock Write As K
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
Dim S As String
Dim p As Long
Dim K As Integer
Dim ll As Long
S = App.PATH & "\FONTS\"

nFonts = Val(GetVar(S & "Fonts.dat", "INIT", "numFonts"))
ReDim Fonts(1 To nFonts)
ll = 2 + nFonts * 10
For p = 1 To nFonts
K = FreeFile

    Open S & p & ".dat" For Binary Access Read Lock Read As #K
        Fonts(p).lSize = LOF(K)
        ReDim Fonts(p).Data(0 To Fonts(p).lSize - 1)
        Get K, , Fonts(p).Data
        Fonts(p).lStart = ll
        ll = ll + Fonts(p).lSize
    Close #K
    Fonts(p).nT = Val(GetVar(S & "Fonts.dat", "FONT" & p, "Textura"))
Next p
For p = 1 To 20

    Debug.Print Fonts(1).Data(p)
Next p

K = FreeFile
Open App.PATH & "\OUTPUT\Fonts.bin" For Binary Access Write Lock Write As #K
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

Dim S As String
Dim t As Long
Dim R As Byte
Dim V As Byte
Dim A As Byte
S = App.PATH & "\RES\INDEX\NewFxs.dat"

Num_Fx = Val(GetVar(S, "INIT", "NumFx"))
Num_Med = Val(GetVar(S, "INIT", "NumMeditaciones"))


If Num_Fx > 0 Then ReDim FxData(1 To Num_Fx)
If Num_Med > 0 Then ReDim MedData(1 To Num_Med)

If Num_Fx > 0 Then

    For t = 1 To Num_Fx

        With FxData(t)

        
            .Animacion = Val(GetVar(S, "FX" & t, "Anim"))
            .OffsetX = Val(GetVar(S, "FX" & t, "OffsetX"))
            .OffsetY = Val(GetVar(S, "FX" & t, "OffsetY"))
            .Alpha = Val(GetVar(S, "FX" & t, "Alpha"))
            .Rombo = Val(GetVar(S, "FX" & t, "Rombo"))
            .Particula = Val(GetVar(S, "FX" & t, "Particula"))
            .Life = Val(GetVar(S, "FX" & t, "Life"))
            R = Val(GetVar(S, "FX" & t, "Rojo"))
            V = Val(GetVar(S, "FX" & t, "Verde"))
            A = Val(GetVar(S, "FX" & t, "Azul"))
            .AnimFinal = Val(GetVar(S, "FX" & t, "AnimFinal"))
            .AnimInicial = Val(GetVar(S, "FX" & t, "AnimInicial"))
            .ParaleloInicial = Val(GetVar(S, "FX" & t, "AnimInitParalelo"))
            .ParaleloStart = Val(GetVar(S, "FX" & t, "InitParaleloStart"))
            If R = 0 And V = 0 And A = 0 Then
                .Color = 0
            Else
                .Color = D3DColorARGB(255, R, V, A)
            End If
        End With
    
    Next t
End If

If Num_Med > 0 Then

    For t = 1 To Num_Med

        With MedData(t)
            
            .Animacion = Val(GetVar(S, "MED" & t, "Anim"))
            .OffsetX = Val(GetVar(S, "MED" & t, "OffsetX"))
            .OffsetY = Val(GetVar(S, "MED" & t, "OffsetY"))
            .Alpha = Val(GetVar(S, "MED" & t, "Alpha"))
            .Rombo = Val(GetVar(S, "MED" & t, "Rombo"))
            .Particula = Val(GetVar(S, "MED" & t, "Particula"))
            R = Val(GetVar(S, "MED" & t, "Rojo"))
            V = Val(GetVar(S, "MED" & t, "Verde"))
            A = Val(GetVar(S, "MED" & t, "Azul"))
            .Life = Val(GetVar(S, "MED" & t, "Life"))
            If R = 0 And V = 0 And A = 0 Then
                .Color = 0
            Else
                .Color = D3DColorARGB(255, R, V, A)
            End If
            .AnimFinal = Val(GetVar(S, "MED" & t, "AnimFinal"))
            .AnimInicial = Val(GetVar(S, "MED" & t, "AnimInicial"))
            .ParaleloInicial = Val(GetVar(S, "MED" & t, "AnimInitParalelo"))
                        .ParaleloStart = Val(GetVar(S, "MED" & t, "InitParaleloStart"))
        End With
    Next t
End If

Dim F As Integer

F = FreeFile


Open App.PATH & "\RES\INDEX\NewFX.BIN" For Binary Access Write Lock Write As #F
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
    
      L = App.PATH & "\Encode\Decor.dat"
    
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
               
         
         End With
         
      Next Decor
      

      Exit Sub
      
ErrHandler:
   Stop
End Sub
Public Sub LoadCraft()
Dim K As Long
Dim S As String

'Cargamos herreria, sastreria y carpinteria.
S = App.PATH & "\ENCODE\Herreria.dat"

NumHerr = Val(GetVar(S, "INIT", "NUM"))
ReDim cHerreria(1 To NumHerr)
For K = 1 To NumHerr
    With cHerreria(K)
        .Item = Val(GetVar(S, K, "ITEM"))
        .Tipo = Val(GetVar(S, K, "TIPO"))
        .ProfesionNivel = Val(GetVar(S, K, "NIVEL"))
        .Mat1 = Val(GetVar(S, K, "BRONCE"))
        .Mat2 = Val(GetVar(S, K, "PLATA"))
        .Mat3 = Val(GetVar(S, K, "ORO"))
        .Version = Val(GetVar(S, K, "VER"))
    End With
Next K
S = App.PATH & "\ENCODE\Sastreria.dat"
NumSastr = Val(GetVar(S, "INIT", "NUM"))
ReDim cSastreria(1 To NumSastr)
For K = 1 To NumSastr
    With cSastreria(K)
        .Item = Val(GetVar(S, K, "ITEM"))
        .Tipo = Val(GetVar(S, K, "TIPO"))
        .ProfesionNivel = Val(GetVar(S, K, "NIVEL"))
        .Mat1 = Val(GetVar(S, K, "PIEL1"))
        .Mat2 = Val(GetVar(S, K, "PIEL2"))
        .Mat3 = Val(GetVar(S, K, "PIEL3"))
        .Version = Val(GetVar(S, K, "VER"))
    End With
Next K
S = App.PATH & "\ENCODE\Carpinteria.dat"
NumCarp = Val(GetVar(S, "INIT", "NUM"))
ReDim cCarpinteria(1 To NumCarp)
For K = 1 To NumCarp
    With cCarpinteria(K)
        .Item = Val(GetVar(S, K, "ITEM"))
        .Tipo = Val(GetVar(S, K, "TIPO"))
        .ProfesionNivel = Val(GetVar(S, K, "NIVEL"))
        .Mat1 = Val(GetVar(S, K, "Madera"))
        .Mat2 = Val(GetVar(S, K, "Madera2"))
        .Mat3 = Val(GetVar(S, K, "MARFIL"))
        .Version = Val(GetVar(S, K, "VER"))
    End With
Next K

End Sub
Public Sub LoadPremios()
Dim K As Long
Dim S As String
Dim J As Long
S = App.PATH & "\ENCODE\Canje.dat"



NumCanje = Val(GetVar(S, "INIT", "NumItems"))
ReDim Canjes(1 To NumCanje)
For K = 1 To NumCanje

    With Canjes(K)
    

        .Nombre = GetVar(S, "PREMIO" & K, "Nombre")
        .Info = GetVar(S, "PREMIO" & K, "Info")
        .Descuento = Val(GetVar(S, "PREMIO" & K, "Descuento"))
        .vGema = Val(GetVar(S, "PREMIO" & K, "Valor_Gemas"))
        .vMM = Val(GetVar(S, "PREMIO" & K, "Valor_MM"))
        .nItems = Val(GetVar(S, "PREMIO" & K, "NumObjs"))
        .Tipo = Val(GetVar(S, "PREMIO" & K, "Tipo"))
        .Version = Val(GetVar(S, "PREMIO" & K, "Version"))
        ReDim .Cant(1 To .nItems)
        ReDim .Items(1 To .nItems)
        For J = 1 To .nItems
            .Items(J) = Val(Readfield(1, GetVar(S, "PREMIO" & K, "OBJ" & J), Asc("-")))
            .Cant(J) = Val(Readfield(2, GetVar(S, "PREMIO" & K, "OBJ" & J), Asc("-")))
        Next J



    End With
    
Next K

End Sub
