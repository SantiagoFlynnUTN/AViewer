Attribute VB_Name = "modMapHandler"
Public Type ttNewEstatic
    W As Integer
    H As Integer
    L As Integer
    T As Integer
    th As Single
    tw As Single
    Replace As Integer
End Type
Public Type ttNewIndex
    Estatic As Integer ' Info de estatica
    temp As Byte
    Dinamica As Integer ' Animacion
    OverWriteGrafico As Integer ' Grafico
    Replace As Integer
End Type
Private TempIndex() As ttNewIndex
Private TempEstatic() As ttNewEstatic
Private nTempIndex As Integer
Private nTempEstatic As Integer
Dim TI() As Integer
Dim TE() As Integer
Dim NTI As Integer
Dim NTE As Integer


Public Sub Analizar(ByVal Path As String, ByRef ListIndex As ListBox, ByRef ListEstatic As ListBox, ByRef lblnti As Label, ByRef lblnte As Label)
    Dim x As Long
    Dim y As Long


    Dim K As Long
    'Cargamos los indices.
    Call Cargar_Temporal_Dats(left$(Path, Len(Path) - 7) & "TempIndex", TI, TE, NTI, NTE)

        
    'Cargamos los listbox
    ListIndex.Clear
    ListEstatic.Clear
    For K = 1 To NTI
        ListIndex.AddItem TI(K) & " - [" & TempIndex(TI(K)).OverWriteGrafico & "]"
    Next K
    For K = 1 To NTE
        ListEstatic.AddItem TE(K)
    Next K
    ListIndex.ListIndex = -1
    ListEstatic.ListIndex = -1

    lblnti.Caption = NTI
    lblnte.Caption = NTE


End Sub
Public Sub VerEstatic_Info(ByVal Indice As Integer, ByRef lblIndex As Label, ByRef lblLeft As Label, ByRef lblTop As Label, ByRef lblwidth As Label, ByRef lblheight As Label, ByRef lblRep As Label)

    If Indice <= NTE Then
        If TE(Indice) <= nTempEstatic Then
            With TempEstatic(TE(Indice))
                lblIndex.Caption = TE(Indice)
                lblLeft.Caption = .L
                lblTop.Caption = .T
                lblwidth.Caption = .W
                lblheight.Caption = .H
                lblRep.Caption = .Replace
            End With
        End If
    End If
End Sub
Public Sub VerIndex_Grafico(ByVal Indice As Integer, ByRef Picture As PictureBox, Optional ByVal GraficoW As Integer, Optional ByVal GraficoH As Integer, Optional ByVal Index As Integer)
    If Index = 0 Then
        If Indice <= NTI Then
            If TempIndex(TI(Indice)).OverWriteGrafico > 0 Then
                dxEngine.DibujareEnHwnd3 Picture.hwnd, TempIndex(TI(Indice)).OverWriteGrafico, 0, 0, True, GraficoW, GraficoH
            End If
        End If
    End If
End Sub
Public Sub VerIndex_Preview(ByVal Indice As Integer, ByRef Picture As PictureBox, Optional ByVal GraficoW As Integer, Optional ByVal GraficoH As Integer, Optional ByVal Index As Integer)
    If Index = 0 Then
        If Indice <= NTI Then
            If TempIndex(TI(Indice)).OverWriteGrafico > 0 Then
                If TempIndex(TI(Indice)).temp Then
                    dxEngine.DibujareEnHwndIndex Picture.hwnd, TempIndex(TI(Indice)).OverWriteGrafico, TempEstatic(TempIndex(TI(Indice)).Estatic).L, TempEstatic(TempIndex(TI(Indice)).Estatic).T, TempEstatic(TempIndex(TI(Indice)).Estatic).W, TempEstatic(TempIndex(TI(Indice)).Estatic).H, 0, 0, True, GraficoW, GraficoH
                Else
                    dxEngine.DibujareEnHwndIndex Picture.hwnd, TempIndex(TI(Indice)).OverWriteGrafico, EstaticData(TempIndex(TI(Indice)).Estatic).L, EstaticData(TempIndex(TI(Indice)).Estatic).T, EstaticData(TempIndex(TI(Indice)).Estatic).W, EstaticData(TempIndex(TI(Indice)).Estatic).H, 0, 0, True, GraficoW, GraficoH
                
                End If
            End If
        End If
    End If
End Sub

Public Sub VerIndex_Info(ByVal Indice As Integer, ByRef lblIndex As Label, ByRef lblDinamic As Label, ByRef lblEstatic As Label, ByRef lblTEmp As Label, ByRef lblGrafico As Label, ByRef lblRep As Label)
    If Indice <= NTI Then
        If TI(Indice) <= nTempIndex Then
            With TempIndex(TI(Indice))
                lblIndex.Caption = TI(Indice)
                lblDinamic.Caption = TempIndex(TI(Indice)).Dinamica
                lblEstatic.Caption = TempIndex(TI(Indice)).Estatic
                lblGrafico.Caption = TempIndex(TI(Indice)).OverWriteGrafico
                lblTEmp.Caption = TempIndex(TI(Indice)).temp
                lblRep.Caption = TempIndex(TI(Indice)).Replace
            End With
        End If
    End If

End Sub
Public Sub Cargar_Temporal_Dats(ByVal Path As String, ByRef TI() As Integer, ByRef TE() As Integer, ByVal NumTempIndex As Integer, ByVal NumTempEstatic As Integer)
    Dim K As Long
    Dim LTE As Integer
    Dim LTI As Integer
    'Cargamos la cantidad
    NTI = Val(GetVar(Path, "INIT", "NumTI"))
    NTE = Val(GetVar(Path, "INIT", "NumTE"))

    'Redimencionamos
    ReDim TI(1 To NTI)
    ReDim TE(1 To NTE)
    
    'Cargamos los indices temporales usados.
    For K = 1 To NTI
        TI(K) = Val(GetVar(Path, K, "Index"))
        If TI(K) > LTI Then LTI = TI(K)
    Next K
    For K = 1 To NTE
        TE(K) = Val(GetVar(Path, "e" & K, "Index"))
        If TE(K) > LTE Then LTE = TE(K)
    Next K
    
    'Igualamos el maximo TE y TI y redimencionamos.
    nTempEstatic = LTE
    nTempIndex = LTI
    ReDim TempIndex(1 To nTempIndex)
    ReDim TempEstatic(1 To nTempEstatic)
    
    'Cargamos los indices temporales.
    For K = 1 To NTI
        With TempIndex(TI(K))
            .OverWriteGrafico = Val(GetVar(Path, K, "OverWriteGrafico"))
            .Dinamica = Val(GetVar(Path, K, "Dinamica"))
            .Estatic = Val(GetVar(Path, K, "Estatica"))
            .temp = Val(GetVar(Path, K, "Temp"))
            .Replace = Val(GetVar(Path, K, "Replace"))
        End With
    Next K
    'Cargamos las estaticas temporales.
    For K = 1 To NTE
        With TempEstatic(TE(K))
            .L = Val(GetVar(Path, "e" & K, "Left"))
            .T = Val(GetVar(Path, "e" & K, "Top"))
            .W = Val(GetVar(Path, "e" & K, "Width"))
            .H = Val(GetVar(Path, "e" & K, "Height"))
            .Replace = Val(GetVar(Path, "e" & K, "Replace"))
        End With
    Next K
    
End Sub
'/////////////////////SUB Y FUNCTIONS DE LECTURA Y MODIFICACION DE ESTATICS/////////////////
Public Function Estatic_rLeft(ByVal Indice As Integer, Optional ByVal Index As Integer) As Integer
    If Index = 0 Then
        If Indice <= NTE And Indice > 0 Then
            Estatic_rLeft = TempEstatic(TE(Indice)).L
        End If
    ElseIf Index <= nTempEstatic Then
        Estatic_rLeft = TempEstatic(Index).L
    End If
    
End Function
Public Function Estatic_rTop(ByVal Indice As Integer, Optional ByVal Index As Integer) As Integer
    If Index = 0 Then
        If Indice <= NTE And Indice > 0 Then
            Estatic_rTop = TempEstatic(TE(Indice)).T
        End If
    ElseIf Index <= nTempEstatic Then
        Estatic_rTop = TempEstatic(Index).T
    End If
End Function
Public Function Estatic_rWidth(ByVal Indice As Integer, Optional ByVal Index As Integer) As Integer
    If Index = 0 Then
        If Indice <= NTE And Indice > 0 Then
            Estatic_rWidth = TempEstatic(TE(Indice)).W
        End If
    ElseIf Index <= nTempEstatic Then
        Estatic_rWidth = TempEstatic(Index).W
    End If
End Function
Public Function Estatic_rHeight(ByVal Indice As Integer, Optional ByVal Index As Integer) As Integer
    If Index = 0 Then
        If Indice <= NTE And Indice > 0 Then
            Estatic_rHeight = TempEstatic(TE(Indice)).H
        End If
    ElseIf Index <= nTempEstatic Then
        Estatic_rHeight = TempEstatic(Index).H
    End If
End Function
Public Sub Estatic_wLeft(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTE And Indice > 0 Then
        TempEstatic(TE(Indice)).L = Valor
    End If
End Sub
Public Sub Estatic_wTop(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTE And Indice > 0 Then
        TempEstatic(TE(Indice)).T = Valor
    End If
End Sub
Public Sub Estatic_wWidth(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTE And Indice > 0 Then
        TempEstatic(TE(Indice)).W = Valor
    End If
End Sub
Public Sub Estatic_wHeight(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTE And Indice > 0 Then
        TempEstatic(TE(Indice)).H = Valor
    End If
End Sub
'//////////////////////////////SUBS DE LETURA Y MODIFICACION DE TEMPINDEX.
Public Function Index_rEstatic(ByVal Indice As Integer) As Integer
    If Indice <= NTI And Indice > 0 Then
        Index_rEstatic = TempIndex(TI(Indice)).Estatic
    End If
End Function
Public Sub Index_wEstatic(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTI And Indice > 0 Then
        TempIndex(TI(Indice)).Estatic = Valor
    End If
End Sub
Public Function Index_rTemp(ByVal Indice As Integer) As Integer
    If Indice <= NTI And Indice > 0 Then
        Index_rTemp = TempIndex(TI(Indice)).temp
    End If
End Function
Public Sub Index_wTemp(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTI And Indice > 0 Then
        TempIndex(TI(Indice)).temp = Valor
    End If
End Sub
Public Function Index_rDinamic(ByVal Indice As Integer) As Integer
    If Indice <= NTI And Indice > 0 Then
        Index_rDinamic = TempIndex(TI(Indice)).Dinamica
    End If
End Function
Public Sub Index_wDinamic(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTI And Indice > 0 Then
        TempIndex(TI(Indice)).Dinamica = Valor
    End If
End Sub
Public Function Index_rGrafico(ByVal Indice As Integer) As Integer
    If Indice <= NTI And Indice > 0 Then
        Index_rGrafico = TempIndex(TI(Indice)).OverWriteGrafico
    End If
End Function
Public Sub Index_wGrafico(ByVal Indice As Integer, ByVal Valor As Integer)
    If Indice <= NTI And Indice > 0 Then
        TempIndex(TI(Indice)).OverWriteGrafico = Valor
    End If
End Sub
