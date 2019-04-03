VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDynDX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const HASH_TABLE_SIZE As Long = 337


Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    texture As Direct3DTexture8
    Size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode
Private Surfaces_Entrys() As SURFACE_ENTRY_DYN
Private mD3D As D3DX8
Private device As Direct3DDevice8
Private ResourcePath As String
Private mMaxEntries As Integer
Private mCantidadGraficos As Integer
'<Modified by: Lean at 10/20/2014-01:33:37>
Public mFreeMemoryBytes As Double
'</Modified by: Lean at 10/20/2014-01:33:37 >
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Private Declare Function SetDIBitsToDevice _
'                Lib "GDI32.dll" (ByVal hdc As Long, _
'                                 ByVal XDest As Long, _
'                                 ByVal YDest As Long, _
'                                 ByVal dwWidth As Long, _
'                                 ByVal dwHeight As Long, _
'                                 ByVal xSrc As Long, _
'                                 ByVal ySrc As Long, _
'                                 ByVal uStartScan As Long, _
'                                 ByVal CScanLine As Long, _
'                                 lpBits As Any, _
'                                 lpbmi As BITMAPINFO, _
'                                 ByVal fuColorUse As Long) As Long

Private Sub Class_Terminate()
    
    Dim I As Long
    Dim J As Long
    
    'Destroy every surface in memory
    For I = 0 To HASH_TABLE_SIZE - 1
        With TexList(I)
            For J = 1 To .surfaceCount
                Set .SurfaceEntry(J).texture = Nothing
            Next J
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next I
End Sub

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
If FileName = 0 Then

Exit Function
End If

    Dim I As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For I = 1 To .surfaceCount
            If .SurfaceEntry(I).FileName = FileName Then
                .SurfaceEntry(I).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(I).texture_width
                textheight = .SurfaceEntry(I).texture_height
                Set GetTexture = .SurfaceEntry(I).texture
                Exit Function
            End If
        Next I
    End With
    'Not in memory, load it!
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)

End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Double) As Boolean
    
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    ResourcePath = App.Path & "\Graficos\"
    If CConfig.CargaDinamica = False Then
        ReDim Surfaces_Entrys(1 To 32000)
    End If

    
    Init = True
End Function

Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Long, ByRef textheight As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index As Integer
    Index = Archivo Mod HASH_TABLE_SIZE
    'lean1*
    Dim Data()     As Byte

    Dim d As Long

    With TexList(Index)

        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .FileName = Archivo
            
            'Ultimo acceso
            .UltimoAcceso = GetTickCount



            Set .texture = mD3D.CreateTextureFromFileEx(device, App.Path & "\RES\GRAFICOS\" & .FileName & ".PNG", 0, 0, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
            

                

            .texture.GetLevelDesc 0, surface_desc
        
            .texture_width = surface_desc.Height
            .texture_height = surface_desc.Width
            .Size = surface_desc.Size
            texwidth = .texture_width
            textheight = .texture_height
            texture_info.Height = surface_desc.Height
            texture_info.Width = surface_desc.Width
            Set CrearGrafico = .texture
            mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.Size
        End With
    End With
    
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp" & Err.Description
End Function
Public Function CrearGraficoEstatico(ByVal Archivo As Integer, ByRef Data() As Byte) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index As Integer
    Index = Archivo Mod HASH_TABLE_SIZE
    'lean1*
    
    Dim d As Long
    
        With Surfaces_Entrys(Archivo)
            'Nombre
            .FileName = Archivo
            
            'Ultimo acceso
            .UltimoAcceso = GetTickCount

             Set .texture = mD3D.CreateTextureFromFileInMemoryEx(device, Data(0), CLng(UBound(Data) + 1), 0, 0, 1, 0, D3DFMT_A8R8G8B8 _
             , D3DPOOL_DEFAULT, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
            
               
            .texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.Width
            .texture_height = texture_info.Height
            .Size = surface_desc.Size
            mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.Size
        End With

    
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp" & Err.Description
End Function
Public Sub Release()
    Dim surface_desc As D3DSURFACE_DESC
Dim I As Long
Dim X As Long
For X = 0 To HASH_TABLE_SIZE - 1
For I = 1 To TexList(X).surfaceCount
    If TexList(X).SurfaceEntry(I).FileName <> 0 Then
    If TexList(X).SurfaceEntry(I).Size > 0 Then Call TexList(X).SurfaceEntry(I).texture.GetLevelDesc(0, surface_desc)
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.Size
    Set TexList(X).SurfaceEntry(I).texture = Nothing
    TexList(X).SurfaceEntry(I).FileName = 0
    End If
Next I
TexList(X).surfaceCount = 0
Next X
End Sub
Private Function RemoveLRU() As Boolean
    
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim I As Long
    Dim J As Long
    Dim surface_desc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For I = 0 To HASH_TABLE_SIZE - 1
        With TexList(I)
            For J = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(J).UltimoAcceso Then
                    LRUi = I
                    LRUj = J
                    LRUtime = .SurfaceEntry(J).UltimoAcceso
                End If
            Next J
        End With
    Next I
    
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).texture.GetLevelDesc(0, surface_desc)
    
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With TexList(LRUi)
            For J = LRUj To .surfaceCount - 1
                .SurfaceEntry(J) = .SurfaceEntry(J + 1)
            Next J
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
    
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.Size
End Function

