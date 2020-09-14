VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compilador BinData & Objetos"
   ClientHeight    =   5175
   ClientLeft      =   10575
   ClientTop       =   6225
   ClientWidth     =   6270
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6270
   Begin VB.CommandButton Command10 
      Caption         =   "Compilar Efectos.bin"
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Modificar Grafico"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cargar NI y ESTatic"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Buscar Grafico en NI"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar en Estatic"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generar Maps.Bin"
      Height          =   480
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generar Fonts"
      Height          =   480
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compilar INDEX.IND"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   960
      TabIndex        =   2
      Top             =   4560
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar BinData"
      Height          =   480
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdCompilar 
      Caption         =   "Generar ObjNames"
      Height          =   480
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MasterFileName As String = "obj.dat"
Private Const NameFileName As String = "ItemNames.Dat"
Private Const BinFileName As String = "ItemInfo.bin"
Private Type tModRaza
    Fuerza As Integer
    Agilidad As Integer
    Suerte As Integer
    Inteligencia As Integer
    Constitucion As Integer
End Type
Private ModifRaza(1 To 5) As tModRaza
Private MasterFilePath   As String
Private NameFilePath     As String
Private BinFilePath      As String
Private Type tCardinal
    X As Byte
    Y As Byte
End Type
Private Type tMapExtra
    MapaGrafico As Byte
    SombrasAmbientales As Integer
    GraficoMiniMapa As Integer
    SaltosFila(1 To 100) As Integer
    INPos() As tCardinal
End Type
Private MapExtra(1 To 160) As tMapExtra
Private Enum eOBJType
      otUseOnce = 1
      otWeapon = 2
      otArmadura = 3
      otArboles = 4
      otGuita = 5
      otPuertas = 6
      otContenedores = 7         ' No se usa
      otCarteles = 8
      otLlaves = 9
      otForos = 10
      otPociones = 11
      otBebidas = 13
      otLeña = 14
      otFogata = 15
      otESCUDO = 16
      otCASCO = 17
      otAnillo = 18
      otTeleport = 19
      otMueble = 20
      otJoyas = 21
      otYacimiento = 22
      otMinerales = 23
      otPergaminos = 24
      otAuras = 25
      otInstrumentos = 26
      otYunque = 27
      otFragua = 28
      otBarcos = 31
      otFlechas = 32
      otBotellaVacia = 33
      otBotellaLlena = 34
      otManchas = 35             ' No se usa
      otArbolElfico = 36
      otMochilas = 37            ' No se usa
      otGema = 38
      otYacimientoPez = 39
      otMapa = 40                ' Marian16?
      otCualquiera = 255
End Enum


Private Type ObjData

      Name As String 'Nombre del obj
    
      OBJType As Byte 'Tipo enum que determina cuales son las caract del obj
    
      grhindex As Integer ' Indice del grafico que representa el obj
      GrhSecundario As Integer
      
      SkSastreria As Byte
      
      PielL As Integer
      PielO As Integer
      PielB As Integer
      
      ItemGM As Byte
      ItemLevel As Byte
      NoFundible As Byte
      
      'Solo contenedores
      MaxItems As Integer
      Apuñala As Byte
      Acuchilla As Byte
    
      HechizoIndex As Integer
    
      ForoID As String
    
      MinHp As Integer ' Minimo puntos de vida
      MaxHP As Integer ' Maximo puntos de vida
    
      MineralIndex As Integer
      LingoteInex As Integer
    
      proyectil As Byte
      Trabajo_Tipo As Byte
      
      Municion As Integer
    
      Crucial As Byte
      Newbie As Byte
    
      'Puntos de Stamina que da
      MinSta As Integer ' Minimo puntos de stamina
    
      'Pociones
      TipoPocion As Byte
      MaxModificador As Integer
      MinModificador As Integer
      DuracionEfecto As Long
      MinSkill As Integer
      LingoteIndex As Integer
    
      MinHIT As Integer 'Minimo golpe
      MaxHIT As Integer 'Maximo golpe
    
      MinHam As Integer
      MinSed As Integer
    
      Def As Integer
      MinDef As Integer ' Armaduras
      MaxDef As Integer ' Armaduras
    
      Ropaje As Integer 'Indice del grafico del ropaje
    
      WeaponAnim As Integer ' Apunta a una anim de armas
      WeaponRazaEnanaAnim As Integer
      ShieldAnim As Integer ' Apunta a una anim de escudo
      CascoAnim As Integer
    
      Valor As Long     ' Precio
    
      Cerrada As Integer
      Llave As Byte
      clave As Long 'si clave=llave la puerta se abre o cierra
    
      Radio As Integer ' Para teleps: El radio para calcular el random de la pos destino
    
      Guante As Byte ' Indica si es un guante o no.
    
      IndexAbierta As Integer
      IndexCerrada As Integer
      IndexCerradaLlave As Integer
    
      RazaEnana As Byte
      RazaDrow As Byte
      RazaElfa As Byte
      RazaGnoma As Byte
      RazaHumana As Byte
    
      Mujer As Byte
      Hombre As Byte
    
      Envenena As Byte
      Paraliza As Byte
    
      Agarrable As Byte
    
      LingH As Integer
      LingO As Integer
      LingP As Integer
      Madera As Integer
      NoDecraft As Byte
      MaderaElfica As Integer
    
      SkHerreria As Integer
      SkCarpinteria As Integer
    
      texto As String
    
      'Clases que no tienen permitido usar este obj
      'ClaseProhibida(1 To NUMCLASES) As eClass
      CP1 As Byte
      CP2 As Byte
      
      NoVendible As Byte
      SoulBound As Byte
      
      MinMagicHit As Integer
      MaxMagicHit As Integer
      
      Marfil As Integer
    
      Mat1 As Integer
      Mat2 As Integer
      Mat3 As Integer
      
      NivelProf As Byte
      TipoCraft As Byte
      
      
      Snd1 As Integer
      Snd2 As Integer
      Snd3 As Integer
    
      AlianzaEnlistado As Integer
      HordaEnlistado As Integer
    
      NoSeCae As Integer
    
      StaffPower As Integer
      StaffDamageBonus As Integer
      DefensaMagicaMax As Integer
      DefensaMagicaMin As Integer
      Refuerzo As Byte
    
      Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
      NoLog As Byte 'es un objeto que esta prohibido loguear?
    
      Upgrade As Integer
    
      MenuIndex As Byte
        
      EfectoObjeto As Byte
      Aura As Byte
      Sombra As Byte
      RP As Byte
      GP As Byte
End Type

Private ObjData()                          As ObjData


Private Sub cmdCompilar_Click()
   
   
   GenerarArchivos
   
   
End Sub

Sub GenerarArchivos()

On Error GoTo ErrHandler

   Dim LoopC            As Long
   Dim MasterFile       As Integer
   Dim NameFile         As Integer
   Dim BinFile          As Integer

   
   MasterFilePath = App.path & "\ENCODE\" & MasterFileName
   NameFilePath = App.path & "\OUTPUT\" & NameFileName
   
   If Not FileExist(MasterFilePath) Then
      MsgBox MasterFileName & " no existe"
   End If
   
   If Not LoadOBJData() Then
      MsgBox "Error cargando archivo de Objetos"
   End If
      
   NameFile = FreeFile
   
   Open NameFilePath For Output As NameFile
      For LoopC = 1 To UBound(ObjData)
            Print #NameFile, ObjData(LoopC).Name
      Next LoopC
   Close NameFile
      
   MsgBox NameFileName & " generado"
   
   Exit Sub
   
ErrHandler:
   MsgBox "Error generando " & NameFileName
End Sub


Function FileExist(ByVal file As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
      '*****************************************************************
      'Se fija si existe el archivo
      '*****************************************************************

      FileExist = LenB(Dir$(file, FileType)) <> 0
End Function

Function LoadOBJData() As Boolean

On Error GoTo ErrHandler

      Dim Object As Long
      Dim Leer   As clsIniManager
      Dim NumObjDatas      As Integer
      Set Leer = New clsIniManager
    
      Call Leer.Initialize(MasterFilePath)
    
      'obtiene el numero de obj
      NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))

      ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
      'Llena la lista

      For Object = 1 To NumObjDatas

            With ObjData(Object)
                   .NoDecraft = Val(Leer.GetValue("OBJ" & Object, "NoDecraft"))
                   
                  .Name = Leer.GetValue("OBJ" & Object, "Name")
                  .Sombra = Val(Leer.GetValue("OBJ" & Object, "Sombra"))
                  'Pablo (ToxicWaste) Log de Objetos.
                  .Log = Val(Leer.GetValue("OBJ" & Object, "Log"))
                  .NoLog = Val(Leer.GetValue("OBJ" & Object, "NoLog"))
                  '07/09/07
                   .Trabajo_Tipo = Val(Leer.GetValue("OBJ" & Object, "TipoTrabajo"))
                  .grhindex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
                    
                  If .grhindex = 0 Then
                        .grhindex = .grhindex
                  End If
            
                  .OBJType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
                  .Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
                  
                  .ItemLevel = Val(Leer.GetValue("OBJ" & Object, "LVL"))
                  
                  .ItemGM = Val(Leer.GetValue("OBJ" & Object, "GM"))
                  
                  .NoFundible = Val(Leer.GetValue("OBJ" & Object, "NoFundible"))
            
                  Select Case .OBJType

                        Case eOBJType.otArmadura
                              .AlianzaEnlistado = Val(Leer.GetValue("OBJ" & Object, "AlianzaEnlistado"))
                              .HordaEnlistado = Val(Leer.GetValue("OBJ" & Object, "HordaEnlistado"))
                              .Aura = Val(Leer.GetValue("OBJ" & Object, "AURA"))
                        Case eOBJType.otESCUDO
                              .ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                              .Aura = Val(Leer.GetValue("OBJ" & Object, "AURA"))
                        Case eOBJType.otCASCO
                              .CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                              .AlianzaEnlistado = Val(Leer.GetValue("OBJ" & Object, "AlianzaEnlistado"))
                              .HordaEnlistado = Val(Leer.GetValue("OBJ" & Object, "HordaEnlistado"))
                               .Aura = Val(Leer.GetValue("OBJ" & Object, "AURA"))
                        Case eOBJType.otWeapon
                              .WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
                              .Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                              .Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
                              .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                              .MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                              .proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                              .Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
                              .StaffPower = Val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                              .StaffDamageBonus = Val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                              .Refuerzo = Val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                              .AlianzaEnlistado = Val(Leer.GetValue("OBJ" & Object, "AlianzaEnlistado"))
                              .HordaEnlistado = Val(Leer.GetValue("OBJ" & Object, "HordaEnlistado"))
                    
                              .WeaponRazaEnanaAnim = Val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                              .Aura = Val(Leer.GetValue("OBJ" & Object, "AURA"))
                        Case eOBJType.otInstrumentos
                              .Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
                              .Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
                              .Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))

                              
                              'Pablo (ToxicWaste)
                              .AlianzaEnlistado = Val(Leer.GetValue("OBJ" & Object, "AlianzaEnlistado"))
                              .HordaEnlistado = Val(Leer.GetValue("OBJ" & Object, "HordaEnlistado"))
                
                        Case eOBJType.otMinerales
                              .MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                              .IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                              .IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                              .IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                        Case otPociones
                              .TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                              .MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                              .MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                              .DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                        Case eOBJType.otBarcos
                              .MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                              .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                              .MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                        Case eOBJType.otFlechas
                              .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                              .MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                              .Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
                              .Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    
                        Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                              .MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                              .MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    
                        Case eOBJType.otTeleport
                              .Radio = Val(Leer.GetValue("OBJ" & Object, "Radio"))
                              
                  End Select
            
                  ' Menues desplegables p/objeto
                  
                  .SkSastreria = Val(Leer.GetValue("OBJ" & Object, "SkillSastre"))
            
                  .Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
                  .HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
                  .LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
                  .MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
                  .MaxHP = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
                  .MinHp = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
                  .Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
                  .Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
                  .MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
                  .MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
                  .MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
                  .MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
                  .Def = (.MinDef + .MaxDef) / 2
            
                  .RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
                  .RazaDrow = Val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
                  .RazaElfa = Val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
                  .RazaGnoma = Val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
                  .RazaHumana = Val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
                  .Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
            
                  .Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
                  .Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))

                  If .Cerrada = 1 Then
                        .Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
                        .clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
                  End If
            
                  'Puertas y llaves
                  .clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
            
                  .texto = Leer.GetValue("OBJ" & Object, "Texto")
                  .GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
                  .Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
                  .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
                  .Acuchilla = Val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
                  .Guante = Val(Leer.GetValue("OBJ" & Object, "Guante"))
            
                   .CP1 = Val(Leer.GetValue("OBJ" & Object, "CP1"))
                   .CP2 = Val(Leer.GetValue("OBJ" & Object, "CP2"))
                   .NoVendible = Val(Leer.GetValue("OBJ" & Object, "NoVendible"))
                   .MinMagicHit = Val(Leer.GetValue("OBJ" & Object, "MinDañoMagico"))
                   .MaxMagicHit = Val(Leer.GetValue("OBJ" & Object, "MaxDañoMagico"))
                  .Marfil = Val(Leer.GetValue("OBJ" & Object, "Piedras"))
            

                  
                  
                  .RP = Val(Leer.GetValue("OBJ" & Object, "RP"))
                  .GP = Val(Leer.GetValue("OBJ" & Object, "GP"))
                  .DefensaMagicaMax = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
                  .DefensaMagicaMin = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
           
 
                  
                  'Bebidas
                  .MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
            
                  .NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
                  .Upgrade = Val(Leer.GetValue("OBJ" & Object, "Upgrade"))
                  
                  .NivelProf = Val(Leer.GetValue("OBJ" & Object, "NivelProf"))
                  
                  
                                   .PielL = Val(Leer.GetValue("OBJ" & Object, "PielL"))
                  .PielO = Val(Leer.GetValue("OBJ" & Object, "PielO"))
                  .PielB = Val(Leer.GetValue("OBJ" & Object, "PielB"))
                   .LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
                .LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
                 .LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))

                     .Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
                     .MaderaElfica = Val(Leer.GetValue("OBJ" & Object, "Madera2"))
                     .Marfil = Val(Leer.GetValue("OBJ" & Object, "Marfil"))

                  If .LingH > 0 Or .LingP > 0 Or .LingO > 0 Then
                    .Mat1 = .LingH
                    .Mat2 = .LingP
                    .Mat3 = .LingO
                    .TipoCraft = 1
                  End If
                  If .Madera > 0 Or .MaderaElfica > 0 Or .Marfil > 0 Then
                    .Mat1 = .Madera
                    .Mat2 = .MaderaElfica
                    .Mat3 = .Marfil
                    .TipoCraft = 3
                  End If
                  If .PielL > 0 Or .PielO > 0 Or .PielB > 0 Then
                    .Mat1 = .PielL
                    .Mat2 = .PielO
                    .Mat3 = .PielB
                    .TipoCraft = 2
                  End If
                  
                  
                  
            End With

      Next Object
    
      Set Leer = Nothing
    
      LoadOBJData = True
    
      Exit Function
      
ErrHandler:
MsgBox "ERROR LOADOBJDATA:  " & Err.Description & "_" & Object
End Function

Private Sub Command1_Click()

CargarChirimbolitos
'Load_NoblezaData
Load_ModifRaza
Load_FD
LoadTables
Load_HabilidadesData
loadsd
Module1.Load_NpcHostiles
Module1.Load_NpcnoHostiles
Module1.LoadQuest
Module1.LoadMapaData
Module1.LoadDecorData
Module1.LoadPremios
Module1.LoadCraft

Dim b As Long
Dim F As Integer
F = FreeFile

Dim i As Long
If FileExist(App.path & "\OUTPUT\BinData.bin", vbArchive) Then Kill App.path & "\OUTPUT\BinData.bin"
Open App.path & "\OUTPUT\BinData.Bin" For Binary Access Write Lock Write As #F

    Put #F, , num_Chirimbolos_data
    For i = 1 To num_Chirimbolos_data
        Put #F, , Chirimbolos_Data(i).Tipo
        Put #F, , Chirimbolos_Data(i).Graf_Index
        Put #F, , Chirimbolos_Data(i).Tiempo
    
    Next i


   Dim LoopC            As Long
   Dim MasterFile       As Integer
   Dim BinFile          As Integer

   
   MasterFilePath = App.path & "\ENCODE\" & MasterFileName

   If Not FileExist(MasterFilePath) Then
      MsgBox MasterFileName & " no existe"
   End If
   
   If Not LoadOBJData() Then
      MsgBox "Error cargando archivo de Objetos"
   End If
         
      Put #F, , CInt(UBound(ObjData))
      For LoopC = 1 To UBound(ObjData)
         Put #F, , ObjData(LoopC).grhindex
         Put #F, , ObjData(LoopC).OBJType
         Put #F, , ObjData(LoopC).MaxHIT
         Put #F, , ObjData(LoopC).MinHIT
         Put #F, , ObjData(LoopC).MaxDef
         Put #F, , ObjData(LoopC).MinDef
         Put #F, , ObjData(LoopC).DefensaMagicaMax
         Put #F, , ObjData(LoopC).DefensaMagicaMin
         Put #F, , ObjData(LoopC).MaxMagicHit
         Put #F, , ObjData(LoopC).MinMagicHit
         Put #F, , ObjData(LoopC).Valor
         Put #F, , ObjData(LoopC).Mat1
         Put #F, , ObjData(LoopC).Mat2
         Put #F, , ObjData(LoopC).Mat3
         Put #F, , ObjData(LoopC).TipoCraft
         Put #F, , ObjData(LoopC).NivelProf
         Put #F, , ObjData(LoopC).NoDecraft
         Put #F, , ObjData(LoopC).Sombra
         Put #F, , ObjData(LoopC).CP1
         Put #F, , ObjData(LoopC).CP2
         Put #F, , ObjData(LoopC).RP
         Put #F, , ObjData(LoopC).GP
         Put #F, , ObjData(LoopC).proyectil
         Put #F, , ObjData(LoopC).Trabajo_Tipo
         Put #F, , ObjData(LoopC).Newbie
         Put #F, , ObjData(LoopC).ItemLevel
         Put #F, , ObjData(LoopC).ItemGM
         Put #F, , ObjData(LoopC).NoVendible
         Put #F, , ObjData(LoopC).SoulBound

         
      Next LoopC

      'Put #F, , Nobleza_Data.NumItems
      'For LoopC = 1 To Nobleza_Data.NumItems
      '  Put #F, , Nobleza_Data.Items(LoopC).Numero
      '  Put #F, , Nobleza_Data.Items(LoopC).NumItems_Requeridos
      '  For i = 1 To Nobleza_Data.Items(LoopC).NumItems_Requeridos
      '      Put #F, , Nobleza_Data.Items(LoopC).Items_Requeridos(i)
      '      Put #F, , Nobleza_Data.Items(LoopC).cantItems_Requeridos(i)
      '  Next i
      'Next LoopC


    For b = 1 To 5
    
        Put #F, , ModifRaza(b).Fuerza
        Put #F, , ModifRaza(b).Agilidad
        Put #F, , ModifRaza(b).Suerte
        Put #F, , ModifRaza(b).Inteligencia
        Put #F, , ModifRaza(b).Constitucion
        
    
    Next b
    
    Put #F, , NUM_FD
    
    For b = 1 To NUM_FD
        With FD(b)
            Put #F, , .TieneSpecial
            Put #F, , .SurfaceNum
            Put #F, , .GraficosX
            Put #F, , .GraficosY
            Put #F, , .num_Buttons
            Put #F, , .Width
            Put #F, , .Height
            If .num_Buttons > 0 Then
            For i = 1 To .num_Buttons
            
                Put #F, , .Buttons(i).size.Top
                Put #F, , .Buttons(i).size.Left
                Put #F, , .Buttons(i).size.Right
                Put #F, , .Buttons(i).size.bottom
                
                Put #F, , .Buttons(i).NormalSurfaceNum
        
                Put #F, , .Buttons(i).SelectSurfaceNum
                Put #F, , .Buttons(i).PressSurfaceNum
                
                
                Put #F, , .Buttons(i).HandIco
        
                Put #F, , .Buttons(i).Sound
                Put #F, , .Buttons(i).Caption
                
                Put #F, , .Buttons(i).Normal_Rojo
                Put #F, , .Buttons(i).Normal_Verde
                Put #F, , .Buttons(i).Normal_Azul
                
                Put #F, , .Buttons(i).Sel_Rojo
                Put #F, , .Buttons(i).Sel_Verde
                Put #F, , .Buttons(i).Sel_Azul
                
                Put #F, , .Buttons(i).Press_Rojo
                Put #F, , .Buttons(i).Press_Verde
                Put #F, , .Buttons(i).Press_Azul
                
            Next i
            End If
            Put #F, , .Num_Textos
            If .Num_Textos > 0 Then
            For i = 1 To .Num_Textos
            
                Put #F, , .Textos(i).texto
                Put #F, , .Textos(i).X
                Put #F, , .Textos(i).Y
                Put #F, , .Textos(i).r
                Put #F, , .Textos(i).g
                Put #F, , .Textos(i).b
                Put #F, , .Textos(i).A
                Put #F, , .Textos(i).IniciaVisible
                Put #F, , .Textos(i).Centrar
                Put #F, , .Textos(i).MaxWidth
            
            Next i
            End If
            
            Put #F, , .ScreenX
            Put #F, , .ScreenY
            Put #F, , .Draw_Stage
            Put #F, , .AlphaValue
            Put #F, , .Num_TextBox
            
            If .Num_TextBox > 0 Then
                For i = 1 To .Num_TextBox
                Put #F, , .TextBox(i).SurfaceNum
                Put #F, , .TextBox(i).X
                Put #F, , .TextBox(i).Y
                Put #F, , .TextBox(i).OffsetX
                Put #F, , .TextBox(i).OffsetY
                Put #F, , .TextBox(i).fA
                Put #F, , .TextBox(i).fR
                Put #F, , .TextBox(i).fG
                Put #F, , .TextBox(i).fb
                Put #F, , .TextBox(i).IniciaVisible
                Put #F, , .TextBox(i).Centrar
                Put #F, , .TextBox(i).TipoTexto
                Put #F, , .TextBox(i).w
                Put #F, , .TextBox(i).h
                
                Next i
            End If
            
            Put #F, , .Num_Checks
            If .Num_Checks > 0 Then
                For i = 1 To .Num_Checks
                    Put #F, , .Checks(i).X
                    Put #F, , .Checks(i).Y
                    Put #F, , .Checks(i).w
                    Put #F, , .Checks(i).h
                    Put #F, , .Checks(i).SurfaceNum
                    Put #F, , .Checks(i).CheckSurface
                    Put #F, , .Checks(i).Tipo_Check
                    Put #F, , .Checks(i).min
                    Put #F, , .Checks(i).max
                    Put #F, , .Checks(i).IniciaVisible
                    Put #F, , .Checks(i).Caption
                    Put #F, , .Checks(i).ColorA
                    Put #F, , .Checks(i).ColorR
                    Put #F, , .Checks(i).ColorG
                    Put #F, , .Checks(i).ColorB
                Next i
            End If
            
        End With
    
    
    Next b
   
   ' EluTable
   For b = 1 To STAT_MAXELV
      Put #F, , EluTable(b)
   Next b
   
   ' StaTable
   For b = 1 To STAT_MAXELV
      Put #F, , StaTable(b)
   Next b
   
   For b = 1 To MAX_HABILIDADES
   
    Put #F, , Habilidades(b).Grafico
   
   Next b
   
   
   Put #F, , ns
   
   For b = 1 To ns
    
    Put #F, , Spells(b).fx
    Put #F, , Spells(b).loops
    Put #F, , Spells(b).Tipo
    Put #F, , Spells(b).Sound
    Put #F, , Spells(b).Manareq
    Put #F, , Spells(b).Skills
    Put #F, , Spells(b).Libro
    Put #F, , Spells(b).CasterFx
    Put #F, , Spells(b).CasterLoop
   Next b
   
   Put #F, , num_npcs_h
   For b = 1 To num_npcs_h
        Put #F, , hostiles(b).Body
        Put #F, , hostiles(b).Head
        Put #F, , hostiles(b).MAX_HP
        Put #F, , hostiles(b).Snd1
        Put #F, , hostiles(b).Snd2
   Next b
   
   Put #F, , num_npcs_nh
   For b = 1 To num_npcs_nh
        Put #F, , nHostiles(b).Body
        Put #F, , nHostiles(b).Head
        Put #F, , nHostiles(b).MAX_HP
        Put #F, , nHostiles(b).NPCTYPE
   Next b
   
   Module1.SaveQuests F
   
   Put #F, , NroMapas
   For b = 1 To NroMapas
    Put #F, , MapaData(b).X
    Put #F, , MapaData(b).Y
   Next b
    Put #F, , Cantdecordata
    For b = 1 To Cantdecordata
        Put #F, , DecoData(b).DecorType
        Put #F, , DecoData(b).EstadoDefault
        Put #F, , DecoData(b).DecorGrh(1)
        Put #F, , DecoData(b).DecorGrh(2)
        Put #F, , DecoData(b).DecorGrh(3)
        Put #F, , DecoData(b).DecorGrh(4)
        Put #F, , DecoData(b).DecorGrh(5)
        Put #F, , DecoData(b).MaxHP
        Put #F, , DecoData(b).TileW
        Put #F, , DecoData(b).TileH
        Put #F, , DecoData(b).Atacable
        Put #F, , DecoData(b).OffX
        Put #F, , DecoData(b).OffY
        Put #F, , DecoData(b).TileTransY
        Put #F, , DecoData(b).Sombra
        
    Next b
    
    Put #F, , NumCanje
    For b = 1 To NumCanje
        Put #F, , Canjes(b).Tipo
        Put #F, , Canjes(b).vGema
        Put #F, , Canjes(b).vMM
        Put #F, , Canjes(b).nItems
        For i = 1 To Canjes(b).nItems
            Put #F, , Canjes(b).Items(i)
            Put #F, , Canjes(b).Cant(i)
        Next i
        Put #F, , Canjes(b).Version
    Next b
    
    Put #F, , NumHerr
    For b = 1 To NumHerr
        Put #F, , cHerreria(b).Item
        Put #F, , cHerreria(b).Tipo
        Put #F, , cHerreria(b).ProfesionNivel
        Put #F, , cHerreria(b).Mat1
        Put #F, , cHerreria(b).Mat2
        Put #F, , cHerreria(b).Mat3
        Put #F, , cHerreria(b).Version
    Next b
    Put #F, , NumSastr
    For b = 1 To NumSastr
        Put #F, , cSastreria(b).Item
        Put #F, , cSastreria(b).Tipo
        Put #F, , cSastreria(b).ProfesionNivel
        Put #F, , cSastreria(b).Mat1
        Put #F, , cSastreria(b).Mat2
        Put #F, , cSastreria(b).Mat3
        Put #F, , cSastreria(b).Version
    Next b
    Put #F, , NumCarp
    For b = 1 To NumCarp
        Put #F, , cCarpinteria(b).Item
        Put #F, , cCarpinteria(b).Tipo
        Put #F, , cCarpinteria(b).ProfesionNivel
        Put #F, , cCarpinteria(b).Mat1
        Put #F, , cCarpinteria(b).Mat2
        Put #F, , cCarpinteria(b).Mat3
        Put #F, , cCarpinteria(b).Version
    Next b

Close #F


   Dim xk As Integer
   xk = FreeFile
   
   Open App.path & "\OUTPUT\hechizosmensajes.txt" For Output As xk
   
    For b = 1 To ns
        
        Print #xk, Spells(b).magicwords
        Print #xk, Spells(b).propiomsg
        Print #xk, Spells(b).targetmsg
        Print #xk, Spells(b).castermsg
        Print #xk, Spells(b).Info
        Print #xk, Spells(b).Nombre
    Next b
   Close xk
   
   
   xk = FreeFile
   
   Open App.path & "\OUTPUT\npcs.txt" For Output As xk
   
    For b = 1 To num_npcs_h - 500
        Print #xk, hostiles(b).Nombre
    Next b
    For b = 1 To num_npcs_nh
        Print #xk, nHostiles(b).Nombre
        Print #xk, nHostiles(b).Desc
    Next b
   Close xk

   xk = FreeFile
   
   Open App.path & "\OUTPUT\quests.txt" For Output As xk
   
    For b = 1 To nQuest
        Print #xk, Quests(b).Desc & "#"
        Print #xk, Quests(b).Nombre
        
    Next b

   Close xk
   
   
   xk = FreeFile

    Open App.path & "\OUTPUT\Premios.txt" For Output As xk
        
        For b = 1 To NumCanje
            Print #xk, Canjes(b).Nombre
            Print #xk, Canjes(b).Info
        Next b
    Close xk
MsgBox "Compilacion exitosa."

Exit Sub

ErrHandler:
   MsgBox "Error generando BinData.Bin. Error : " & Err.Description
End Sub
Private Sub CargarChirimbolitos()
Dim s As String
Dim i As Long
s = App.path & "\ENCODE\Chirimbolos.dat"


num_Chirimbolos_data = Val(GetVar(s, "INIT", "NUM"))
ReDim Chirimbolos_Data(1 To num_Chirimbolos_data)
For i = 1 To num_Chirimbolos_data
    Chirimbolos_Data(i).Graf_Index = Val(GetVar(s, CStr(i), "Graf_Index"))
    Chirimbolos_Data(i).Tiempo = Val(GetVar(s, CStr(i), "Tiempo"))
    Chirimbolos_Data(i).Tipo = Val(GetVar(s, CStr(i), "Tipo"))

Next i


End Sub

Sub Load_ModifRaza()

   ModifRaza(1).Fuerza = 1
   ModifRaza(1).Agilidad = 1
   ModifRaza(1).Suerte = 2
   ModifRaza(1).Inteligencia = 0
   ModifRaza(1).Constitucion = 2
   
   ModifRaza(2).Fuerza = -1
   
   ModifRaza(2).Agilidad = 2
   ModifRaza(2).Suerte = 1
   ModifRaza(2).Inteligencia = 2
   ModifRaza(2).Constitucion = 0
   
   
   ModifRaza(3).Fuerza = 2
   
   ModifRaza(3).Agilidad = 0
   ModifRaza(3).Suerte = 0
   ModifRaza(3).Inteligencia = 1
   ModifRaza(3).Constitucion = 1
   
   
   ModifRaza(4).Fuerza = -3
   ModifRaza(4).Agilidad = 3
   ModifRaza(4).Suerte = 0
   ModifRaza(4).Inteligencia = 3
   ModifRaza(4).Constitucion = -1
   
   
   ModifRaza(5).Fuerza = 3
   ModifRaza(5).Agilidad = -2
   ModifRaza(5).Suerte = -1
   ModifRaza(5).Inteligencia = -5
   ModifRaza(5).Constitucion = 3

End Sub

Private Sub Command10_Click()
2        On Error GoTo Command10_Click_Error

4     CargarAuras
6     CargarParticulas
8     CargarBuffData
10    SPOTLIGHTS_LOADDAT
      Dim FF As Integer
12    FF = FreeFile

14    If FileExist(App.path & "\RES\OUTPUT\Efectos.bin", vbNormal) Then Kill App.path & "\RES\OUTPUT\Efectos.bin"

16    Open App.path & "\RES\OUTPUT\Efectos.bin" For Binary Access Write Lock Write As #FF
          ' Escribimos auras.
18        EscribirAuras FF
          ' Escribimos particulas
20        EscribirParticulas FF
          'Escribimos Buffs
22        EscribirBuffdataBin FF
          'Escribimos SpotLights
24        SPOTLIGHTS_Escribir FF
26    Close #FF

28    MsgBox "Efectos.Bin compilado exitosamente."

30        Exit Sub

Command10_Click_Error:

32        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command10_Click in line:" & Erl

End Sub

Private Sub Command2_Click()

   Unload Me
   End

End Sub

Private Sub Command3_Click()
    Module1.Compilar_Archivo_Index
End Sub

Private Sub Command4_Click()
    GenerarFonts
    
End Sub

Private Sub Command5_Click()
Dim F As Integer
Dim p As Long
Dim NI(1 To 160) As Integer
Dim Data() As Byte

For p = 1 To 160


    MapExtra(p).GraficoMiniMapa = Val(GetVar(App.path & "\MAPAS\Mapas.dat", "MAPA" & p, "Grafico_mini_Mapa"))
    MapExtra(p).SombrasAmbientales = Val(GetVar(App.path & "\MAPAS\Mapas.dat", "MAPA" & p, "Sombras_Ambientales"))
    MapExtra(p).MapaGrafico = Val(GetVar(App.path & "\MAPAS\Mapas.dat", "MAPA" & p, "MapaGrafico"))
    
F = FreeFile
If FileExist(App.path & "\MAPAS\" & p & ".int", vbNormal) Then
Open App.path & "\MAPAS\" & p & ".int" For Binary Access Read Lock Read As #F

    Get #F, , NI(p)
    If NI(p) > 0 Then
        Get #F, , MapExtra(p).SaltosFila
        ReDim MapExtra(p).INPos(1 To NI(p))
        Get #F, , MapExtra(p).INPos
        Debug.Print MapExtra(p).INPos(1).X
    End If
Close #F
End If


Next p

F = FreeFile
Dim nummap As Integer
nummap = 160
Open App.path & "\RES\OUTPUT\MAPAS.BIN" For Binary Access Write Lock Write As #F

ReDim Data(0 To (LenB(NI(1)) * 160) - 1)

CopyMemory Data(0), NI(1), LenB(NI(1)) * 160
Put #F, , nummap
Put #F, , Data()


For p = 1 To nummap

    Put #F, , MapExtra(p).GraficoMiniMapa
    Put #F, , MapExtra(p).SombrasAmbientales
    Put #F, , MapExtra(p).MapaGrafico
    If NI(p) > 0 Then

        ReDim Data(0 To (LenB(MapExtra(p).SaltosFila(1)) * 100) - 1)
        CopyMemory Data(0), MapExtra(p).SaltosFila(1), LenB(MapExtra(p).SaltosFila(1)) * 100
        Put #F, , Data
        ReDim Data(0 To (2 * NI(p)) - 1)
        CopyMemory Data(0), MapExtra(p).INPos(1), 2 * NI(p)
        Put #F, , Data
    End If
Next p

Close #F
MsgBox "OK"

End Sub

Private Sub Command6_Click()
Dim s As String
Dim i As Long
Dim z As Long

s = App.path & "\Index\NewEstatics.dat"

NumEstatics = Val(GetVar(s, "INIT", "num"))

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

Dim w As Integer
Dim h As Integer
Dim t As Integer
Dim L As Integer

s = frmMain.Text1.Text
L = Val(Readfield(1, s, Asc("-")))
t = Val(Readfield(2, s, Asc("-")))
w = Val(Readfield(3, s, Asc("-")))
h = Val(Readfield(4, s, Asc("-")))



For i = 1 To NumEstatics

    If EstaticData(i).L = L Then
        If EstaticData(i).t = t Then
            If EstaticData(i).w = w Then
                If EstaticData(i).h = h Then
                    MsgBox i
                    Exit For
                End If
            End If
        End If
    End If
        

Next i

If i > NumEstatics Then
MsgBox "No esta. Se agrego. " & NumEstatics + 1
NumEstatics = NumEstatics + 1
s = App.path & "\Index\NewEstatics.dat"
WriteVar s, "INIT", "NUM", NumEstatics

WriteVar s, NumEstatics, "Left", CStr(L)
WriteVar s, NumEstatics, "Top", CStr(t)
WriteVar s, NumEstatics, "Width", CStr(w)
WriteVar s, NumEstatics, "Height", CStr(h)
End If



End Sub

Private Sub Command7_Click()
Dim p As Long
Dim s As String
Dim i As Long
Dim z As Long

s = App.path & "\Index\NewIndex.dat"

NumNewIndex = Val(GetVar(s, "INIT", "num"))

If NumNewIndex > 0 Then
ReDim NewIndexData(1 To NumNewIndex)
For i = 1 To NumNewIndex
    With NewIndexData(i)
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
        If .OverWriteGrafico = Val(frmMain.Text1) Then
            MsgBox i
            Exit For
        End If
    End With
Next i
If i > NumNewIndex Then MsgBox "no se encontro"
End If
End Sub

Private Sub Command8_Click()
Dim s As String
Dim i As Long
Dim z As Long

s = App.path & "\Index\NewIndex.dat"

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
End If
s = App.path & "\Index\NewEstatics.dat"

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
End If
MsgBox "OK"
End Sub

Private Sub Command9_Click()
Dim i As Long
Dim p As Long
Dim og As Integer
Dim ml As Integer
Dim mt As Integer
Dim mg As Integer


Dim nl As Integer
Dim nT As Integer


If Val(frmMain.Text2.Text) > 0 And LenB(frmMain.Text1.Text) > 0 Then
og = Val(frmMain.Text2.Text)
ml = Val(Readfield(1, frmMain.Text1.Text, Asc("-")))
mt = Val(Readfield(2, frmMain.Text1.Text, Asc("-")))
mg = Val(Readfield(3, frmMain.Text1.Text, Asc("-")))

    For i = 1 To NumNewIndex
        If NewIndexData(i).OverWriteGrafico = og Then
                nl = EstaticData(NewIndexData(i).Estatic).L + ml
                nT = EstaticData(NewIndexData(i).Estatic).t + mt
                If nl > 0 Or nT > 0 Then
                    For p = 1 To NumEstatics
                        
                        If EstaticData(p).L = nl Then
                        If EstaticData(p).t = nT Then
                        If EstaticData(p).w = EstaticData(NewIndexData(i).Estatic).w Then
                        If EstaticData(p).h = EstaticData(NewIndexData(i).Estatic).h Then
                            Exit For
                        End If
                        End If
                        End If
                        End If
                    Next p
                    If p > NumEstatics Then
                        Stop
                    Else
                        NewIndexData(i).Estatic = p
                        WriteVar App.path & "\Index\NewIndex.dat", CStr(i), "Estatica", CStr(p)
                    End If
                End If
                If mg > 0 And mg <> NewIndexData(i).OverWriteGrafico Then
                    NewIndexData(i).OverWriteGrafico = mg
                    WriteVar App.path & "\Index\NewIndex.dat", CStr(i), "OverWriteGrafico", CStr(mg)
                End If
        End If
    Next i
End If
MsgBox "OK"
End Sub

