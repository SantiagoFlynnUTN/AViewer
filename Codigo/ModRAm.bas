Attribute VB_Name = "ModRAm"
Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type


Private Type INT64
    LoPart As Long
    HiPart As Long
End Type

Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ulTotalPhys As INT64
    ulAvailPhys As INT64
    ulTotalPageFile As INT64
    ulAvailPageFile As INT64
    ulTotalVirtual As INT64
    ulAvailVirtual As INT64
    ulAvailExtendedVirtual As INT64
End Type

Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Private pUdtMemStatusEX As MEMORYSTATUSEX
Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)


Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Sub WriteVar(file As String, Main As String, Var As String, value As String)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub
Public Function General_Get_Free_Ram_BytesEX() As Double
    pUdtMemStatusEX.dwLength = Len(pUdtMemStatusEX)
    GlobalMemoryStatusEx pUdtMemStatusEX
    General_Get_Free_Ram_BytesEX = CLargeInt(pUdtMemStatusEX.ulAvailPhys.LoPart, pUdtMemStatusEX.ulAvailPhys.HiPart)
End Function


Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double
    Dim dblAns As Double
    dblAns = (Bytes / 1024) / 1024
    General_Bytes_To_Megabytes = Format(dblAns, "###,###,##0.00")
End Function

Public Function General_Get_Free_Ram() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function


Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    Dim dblLo As Double
    Dim dblHi As Double

    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If

    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If

    CLargeInt = dblLo + dblHi * 2 ^ 32

End Function

Private Function NumberInKB(ByVal vNumber As Currency) As String
    Dim strReturn As String

    Select Case vNumber
        Case Is < 1024 ^ 1
            strReturn = CStr(vNumber) & " bytes"

        Case Is < 1024 ^ 2
            strReturn = CStr(Round(vNumber / 1024, 1)) & " KB"

        Case Is < 1024 ^ 3
            strReturn = CStr(Round(vNumber / 1024 ^ 2, 2)) & " MB"

        Case Is < 1024 ^ 4
            strReturn = CStr(Round(vNumber / 1024 ^ 3, 2)) & " GB"
    End Select

    NumberInKB = strReturn

End Function
Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function
Function ReadField(ByVal Pos As Integer, _
    ByRef Text As String, _
    ByVal SepASCII As Byte) As String


    Dim i          As Long
    Dim lastPos    As Long
    Dim CurrentPos As Long
    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)

    Next i
    
    If CurrentPos = 0 Then
        ReadField = Mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = Mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If

End Function


Public Function PasoTiempo(Optional ByRef Counter As Long = -1) As Long
    Static Contador As Long
    Dim tl As Long
    tl = GetTickCount
    If Counter <> -1 Then
        If Counter = 0 Then
            PasoTiempo = 0
        Else
            PasoTiempo = tl - Counter
        End If
        Counter = tl
    Else
        PasoTiempo = Contador - tl
        Contador = tl
    
    End If


End Function
Public Function CumplioIntervalo(ByRef LastCheck As Long, ByVal Intervalo As Long, Optional ByVal Update As Boolean) As Boolean


    Dim tl As Long
    tl = GetTickCount

    If tl - LastCheck >= Intervalo Then
        CumplioIntervalo = True
        If Update Then
            LastCheck = tl
        End If
    End If

End Function
