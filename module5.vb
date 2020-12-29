'Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Option Explicit
Public Conexion As ADODB.Connection
Public miBase As String
Public CadenaConexion As String
Public rsConciliacion As ADODB.Recordset
Public rsVariacion As ADODB.Recordset
Sub ConectarBase()
miBase = "D:\PRUEBAS INFORMES\BASEtRABAJO.accdb"
CadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0; " & "data source=" & miBase & ";"
If Len(Dir(miBase)) = 0 Then
    MsgBox "La base que intenta conectar no se encuentra disponible", vbCritical
    Exit Sub
End If
Set Conexion = New ADODB.Connection
    If Conexion.State = 1 Then
        Conexion.Close
    End If
        Conexion.Open (CadenaConexion)
End Sub
Sub ConsultarPagos()
Dim Sql As String
Dim Criterio  As String
Dim CriterioFechaInicial, CriterioFechaFinal As String
Dim limpiarDatos, lngCampos, i As Long
    limpiarDatos = Sheets("PAGOS").Range("A" & Rows.Count).End(xlUp).Row
    Sheets("PAGOS").Range("A2:D" & limpiarDatos).ClearContents
    Call ConectarBase
    Set rsConciliacion = New ADODB.Recordset
    If rsConciliacion.State = 1 Then
        rsConciliacion.Close
    End If
    Criterio = Sheets("CONTABILIZADOS").Cells(1, 5)
    CriterioFechaInicial = Mid(Sheets("CONTABILIZADOS").Cells(2, 5), 4, 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(2, 5), 2) & "/" & Right(Sheets("CONTABILIZADOS").Cells(2, 5), 4)
    CriterioFechaFinal = Mid(Sheets("CONTABILIZADOS").Cells(3, 5), 4, 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(3, 5), 2) & "/" & Right(Sheets("CONTABILIZADOS").Cells(3, 5), 4)
    Sql = "SELECT [Cédula], [Valor Transmitido],[Fecha de Transmisión], [ALIAS] FROM Conciliacion WHERE Alias = '" & Criterio & "' AND [Fecha de Transmisión] BETWEEN #" & CriterioFechaInicial & "# AND  #" & CriterioFechaFinal & "# "
    rsConciliacion.Open Sql, Conexion
    Sheets("PAGOS").Cells(2, 1).CopyFromRecordset rsConciliacion
    lngCampos = rsConciliacion.Fields.Count
    For i = 0 To lngCampos - 1
         Sheets("PAGOS").Cells(1, i + 1).Value = rsConciliacion.Fields(i).Name
    Next
    limpiarDatos = Sheets("PAGOS").Range("A" & Rows.Count).End(xlUp).Row
    For i = 2 To limpiarDatos
                Sheets("PAGOS").Cells(i, 1).Value = CLng(Sheets("PAGOS").Cells(i, 1))
    Next i
    rsConciliacion.Close
    Set rsConciliacion = Nothing
    Conexion.Close
    Set Conexion = Nothing
End Sub
Sub abrirVariacion()
 Dim Criterio, Sql As String
 Dim limpiarDatos, j As Long
    Call ConectarBase
    Set rsVariacion = New ADODB.Recordset
                Sheets("VARIACION").Range("A4:AV50000").ClearContents
                Criterio = Sheets("CONTABILIZADOS").Cells(4, 2)
                Sql = "SELECT * FROM [VALIDACION_CUOTAS] WHERE ALIAS = '" & Criterio & "'  AND DIFERENCIA >1000  AND [CUOTA TOTAL] IS NOT NULL "
                rsVariacion.Open Sql, Conexion
                Sheets("VARIACION").Cells(4, 1).CopyFromRecordset rsVariacion
                limpiarDatos = Sheets("VARIACION").Range("A" & Rows.Count).End(xlUp).Row
                For j = 4 To limpiarDatos
                    Sheets("VARIACION").Cells(j, 1).Value = CLngLng(Sheets("VARIACION").Cells(j, 1))
                Next j
    rsVariacion.Close
    Set rsVariacion = Nothing
    Conexion.Close
    Set Conexion = Nothing
    Sheets("VARIACION").Visible = xlSheetVisible
End Sub
Sub cruceNovedades()
Dim largNovedades, i As Long
    largNovedades = Sheets("NOVEDADES").Range("A" & Rows.Count).End(xlUp).Row
'    For i = 4 To largNovedades
'        Sheets("NOVEDADES").Cells(i, 11) = CDate(Sheets("NOVEDADES").Cells(i, 11).Value)
'    Next i
    Sheets("NOVEDADES").Select
    For i = 4 To largNovedades
        With Sheets("NOVEDADES")
            .Cells(i, 16) = Application.VLookup(.Cells(i, 1), Sheets("DINAMICA").Range("A:B"), 2, 0)
             If IsNumeric(Cells(i, 16)) Then
                .Cells(i, 17) = .Cells(i, 16).Value - .Cells(i, 8).Value
             End If
             Select Case Len(.Cells(i, 3))
             Case 17
                If Left(.Cells(i, 3), 4) = "0013" Then
                    .Cells(i, 15) = Right(.Cells(i, 3), 13)
                        
                End If
             Case 18
                If Left(.Cells(i, 3), 4) = "0013" Then
                    .Cells(i, 15) = Right(.Cells(i, 3), 14)
                End If
            End Select
            
            .Cells(i, 18) = Application.VLookup(.Cells(i, 15), Sheets("CONTABILIZADOS").Range("B:AN"), 39, 0)
            .Cells(i, 19) = Application.VLookup(.Cells(i, 15), Sheets("CONTABILIZADOS").Range("B:AS"), 44, 0)
            If IsNumeric(Cells(i, 19)) Then
              .Cells(i, 20) = .Cells(i, 19).Value - .Cells(i, 8).Value
            End If
            .Cells(i, 21) = Application.VLookup(.Cells(i, 15), Sheets("CONTABILIZADOS").Range("B:AS"), 3, 0)
            .Cells(i, 22) = Application.VLookup(.Cells(i, 15), Sheets("CONTABILIZADOS").Range("B:AS"), 4, 0)
            .Cells(i, 23) = Application.VLookup(.Cells(i, 15), Sheets("Restructurados").Range("I:J"), 2, 0)
            
        End With
    Next i
End Sub
Sub organizarLibranzas()
Dim CantidadContab, CantidadActivos, CantidadCancel As Long
Dim i, j, X As Long
Dim usuario, FecCargue As String
Dim stre As String
Dim idap As String
Dim sBuffer As String
Dim lSize As Long
Dim consulSQL As String
Dim MismoMes, MesSiguiente, MesSubsiguiente, MesFechaFin As Long
Dim path_Bd As String
Dim cnn As New ADODB.Connection
Dim recSet As New ADODB.Recordset
Dim strDB, strSQL As String
Dim strTabla As String
Dim Criterio, FECHA  As String
Dim ColumnaCancelados, limpiarDatos  As Long
Dim bBien As Boolean
Dim finActivos, finalCancel As Long
Dim ColumnaVariacion As Long
Dim hoja As Worksheet
Dim CANNN As String

If Date < DateValue("14/03/2021") Then
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    
    Application.StatusBar = "Analizando Cra Libranzas"
    MismoMes = Sheets("CONTABILIZADOS").Cells(3, 2)
    MesSiguiente = DateAdd("M", 1, MismoMes) '( Application.EDate(Sheets("CONTABILIZADOS").Cells(3, 2), 1)
    MesSubsiguiente = DateAdd("m", 2, MismoMes) ' Application.EDate(Sheets("CONTABILIZADOS").Cells(3, 2), 2)
    'sBuffer = Space$(255)
    'lSize = Len(sBuffer)
    'MIO'Call GetUserName(sBuffer, lSize)
    'MIO'usuario = LCase(Left$(sBuffer, lSize - 1))
'If usuario = "CP00308" Or usuario = "cp00308" Or usuario = "cp00307" Or usuario = "cp00304" Or usuario = "cp00317" Or usuario = "cp00318" Then '
   
 '   Sheets("CONTABILIZADOS").Cells(4, 5) = usuario
    FecCargue = Format(Now(), "m/d/yyyy h:N:S")
    Sheets("CONTABILIZADOS").Cells(6, 2) = FecCargue
    For Each hoja In ThisWorkbook.Worksheets
            If hoja.FilterMode Then
            hoja.ShowAllData
        End If
    Next hoja
       Sheets("ACTIVOS").Range("AT2:BZ50000").ClearContents
    Sheets("CANCELADOS").Range("AT2:AZ50000").ClearContents
    Sheets("CONTABILIZADOS").Range("AT9:AT50000").ClearContents
    Sheets("VARIACION").Range("A4:AZ50000").ClearContents
    Sheets("NOVEDADES").Range("A4:AZ50000").ClearContents
    Sheets("Restructurados").Range("A2:AZ50000").ClearContents
    
    CantidadContab = Sheets("CONTABILIZADOS").Range("A" & Rows.Count).End(xlUp).Row
    Sheets("ACTIVOS").Range("A2:BZ50000").ClearContents
    Sheets("CANCELADOS").Range("A2:BA50000").ClearContents
    
    For X = 9 To CantidadContab
        With Sheets("CONTABILIZADOS")
            If .Cells(X, 40).Value = "0" Then
                CantidadActivos = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row
                Sheets("CONTABILIZADOS").Rows(X).Copy
                Sheets("ACTIVOS").Rows(CantidadActivos + 1).PasteSpecial xlPasteAll
             ElseIf .Cells(X, 40).Value = "1" Then
                CantidadCancel = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
                Sheets("CONTABILIZADOS").Rows(X).Copy
                Sheets("CANCELADOS").Rows(CantidadCancel + 1).PasteSpecial xlPasteAll
            End If
        MesFechaFin = Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2)
              If Left(Sheets("CONTABILIZADOS").Cells(X, 39), 4) Mod 4 = 0 And MesFechaFin = 2 Then
                    Sheets("CONTABILIZADOS").Cells(X, 46) = "29/" & Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(X, 39), 4)
                ElseIf Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "02" Then
                        Sheets("CONTABILIZADOS").Cells(X, 46) = "28/" & Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(X, 39), 4)
                ElseIf Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "01" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "03" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "05" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "07" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "08" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "10" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "12" Then
                    Sheets("CONTABILIZADOS").Cells(X, 46) = "31/" & Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(X, 39), 4)
                ElseIf Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "04" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "06" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "09" Or Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) = "11" Then
                    Sheets("CONTABILIZADOS").Cells(X, 46) = "30/" & Right(Sheets("CONTABILIZADOS").Cells(X, 39), 2) & "/" & Left(Sheets("CONTABILIZADOS").Cells(X, 39), 4)
              End If
        End With
    Next X
    Call OrdenarActivos
    Call OrdenarCancelados
    '/////////////////////////////////////////////////
    'consulta de pagos
     If Sheets("CONTABILIZADOS").Cells(1, 5).Value <> "" Then
            Call ConsultarPagos
     End If
'//////////////////////////////////
'ANALIZAMOS ACTIVOS
CantidadActivos = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row
finalCancel = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("ACTIVOS").Select
With Sheets("ACTIVOS")
        For i = 2 To CantidadActivos
             If Sheets("CONTABILIZADOS").Cells(2, 2).Value = "TODOS" Then
                Select Case Sheets("CONTABILIZADOS").Cells(1, 2).Value
                    Case "MISMO MES"
                        If Month(MismoMes) = Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MismoMes) = Year(Sheets("ACTIVOS").Cells(i, 37)) Then
                            If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False)) Then
                                'If (Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) <= 15 And (Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) >= -15 Then
                                If (Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) = 0 Then
                                   Sheets("ACTIVOS").Cells(i, 46) = "RETANQUEO"
                                   Else
                                      Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                End If
                                Else
                                      Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                             End If
                            ElseIf (Month(MismoMes) > Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MismoMes) = Year(Sheets("ACTIVOS").Cells(i, 37))) Or Year(MismoMes) > Year(Cells(i, 37)) Then
                                   Sheets("ACTIVOS").Cells(i, 46) = "REPORTAR"
                            Else
                               Sheets("ACTIVOS").Cells(i, 46) = "NO REPORTAR"
                        End If
                    Case "MES SIGUIENTE"
                        If Month(MesSiguiente) = Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MesSiguiente) = Year(Cells(i, 37)) Then
                            If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False)) Then
                                'If (Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) <= 15 And (Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) >= -15 Then
                                If (Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) = 0 Then
                                   Sheets("ACTIVOS").Cells(i, 46) = "RETANQUEO"
                                  Else
                                  Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                End If
                              Else
                                      Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                             End If
                        ElseIf (Month(MesSiguiente) > Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MesSiguiente) = Year(Sheets("ACTIVOS").Cells(i, 37))) Or Year(MesSiguiente) > Year(Sheets("ACTIVOS").Cells(i, 37)) Then
                                   Sheets("ACTIVOS").Cells(i, 46) = "REPORTAR"
                        Else
                               Sheets("ACTIVOS").Cells(i, 46) = "NO REPORTAR"
                       End If
                    Case "MES SUBSIGUIENTE"
                        If Month(MesSubsiguiente) = Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MesSubsiguiente) = Year(Sheets("ACTIVOS").Cells(i, 37)) Then
                            If Not IsError(Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False)) Then
                                'If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) <= 15 And (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - .Cells(i, 15)) >= -15 Then
                                If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) = 0 Then
                                   Sheets("ACTIVOS").Cells(i, 46) = "RETANQUEO"
                                   Else
                                      .Cells(i, 46) = "NUEVO"
                                End If
                                Else
                                      Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                             End If
                           ElseIf (Month(MesSubsiguiente) > Month(Cells(i, 37)) And Year(MesSubsiguiente) = Year(Sheets("ACTIVOS").Cells(i, 37))) Or Year(MesSubsiguiente) > Year(Sheets("ACTIVOS").Cells(i, 37)) Then
                                   Sheets("ACTIVOS").Cells(i, 46) = "REPORTAR"
                            Else
                               Sheets("ACTIVOS").Cells(i, 46) = "NO REPORTAR"
                        End If
                        Sheets("VARIACION").Visible = xlHidden
                 End Select
            ElseIf Sheets("CONTABILIZADOS").Cells(2, 2).Value = "NUEVOS" Then
                    Select Case Sheets("CONTABILIZADOS").Cells(1, 2).Value
                        Case "MISMO MES"
                            If Month(MismoMes) = Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MismoMes) = Year(Sheets("ACTIVOS").Cells(i, 37)) Then
                                  If Not IsError(Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False)) Then
                                          'If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) <= 15 And (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) >= -15 Then
                                          If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) = 0 Then
                                             Sheets("ACTIVOS").Cells(i, 46) = "RETANQUEO"
                                             Else
                                                Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                          End If
                                   Else
                                      Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                End If
                             Else
                                Sheets("ACTIVOS").Cells(i, 46) = "NO REPORTAR"
                             End If
                         Case "MES SIGUIENTE"
                            If Month(MesSiguiente) = Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MesSiguiente) = Year(Sheets("ACTIVOS").Cells(i, 37)) Then
                                    If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False)) Then
                                          'If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) <= 15 And (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) >= -15 Then
                                          If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) = 0 Then
                                             Sheets("ACTIVOS").Cells(i, 46) = "RETANQUEO"
                                             Else
                                                Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                          End If
                                     Else
                                         Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                     End If
                             Else
                                Sheets("ACTIVOS").Cells(i, 46) = "NO REPORTAR"
                             End If
                         Case "MES SUBSIGUIENTE"
                            If Month(MesSubsiguiente) = Month(Sheets("ACTIVOS").Cells(i, 37)) And Year(MesSubsiguiente) = Year(Cells(i, 37)) Then
                                If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False)) Then
                                          'If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) <= 15 And (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) >= -15 Then
                                          If (Application.VLookup(Cells(i, 1), Sheets("CANCELADOS").Range("A2:AG" & finalCancel), 33, False) - Sheets("ACTIVOS").Cells(i, 15)) = 0 Then
                                             Sheets("ACTIVOS").Cells(i, 46) = "RETANQUEO"
                                             Else
                                                Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                          End If
                                Else
                                      Sheets("ACTIVOS").Cells(i, 46) = "NUEVO"
                                End If
                             Else
                                Sheets("ACTIVOS").Cells(i, 46) = "NO REPORTAR"
                             End If
                    End Select
             End If
             Sheets("ACTIVOS").Cells(i, 47) = Application.RoundUp(Cells(i, 26), 0)
             MesFechaFin = Right(Sheets("ACTIVOS").Cells(i, 39), 2)
              If Left(Sheets("ACTIVOS").Cells(i, 39), 4) Mod 4 = 0 And MesFechaFin = 2 Then
                Sheets("ACTIVOS").Cells(i, 48) = "29/" & Right(Sheets("ACTIVOS").Cells(i, 39), 2) & "/" & Left(Sheets("ACTIVOS").Cells(i, 39), 4)
                ElseIf Right(Cells(i, 39), 2) = "02" Then
                        Sheets("ACTIVOS").Cells(i, 48) = "28/" & Right(Sheets("ACTIVOS").Cells(i, 39), 2) & "/" & Left(Sheets("ACTIVOS").Cells(i, 39), 4)
                ElseIf Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "01" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "03" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "05" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "07" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "08" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "10" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "12" Then
                    Sheets("ACTIVOS").Cells(i, 48) = "31/" & Right(Sheets("ACTIVOS").Cells(i, 39), 2) & "/" & Left(Sheets("ACTIVOS").Cells(i, 39), 4)
                ElseIf Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "04" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "06" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "09" Or Right(Sheets("ACTIVOS").Cells(i, 39), 2) = "11" Then
                    Sheets("ACTIVOS").Cells(i, 48) = "30/" & Right(Sheets("ACTIVOS").Cells(i, 39), 2) & "/" & Left(Sheets("ACTIVOS").Cells(i, 39), 4)
              End If
         Next i
    End With
'/////////////////////
'ACTUALIZAMOS TABLA DINAMICA
     Sheets("DINAMICA").Visible = xlSheetVisible
    Sheets("DINAMICA").Select
    ActiveSheet.PivotTables("Tabla dinámica1").PivotCache.Refresh
      With Sheets("CANCELADOS")
    Sheets("CANCELADOS").Select
        For j = 2 To finalCancel
            Sheets("CANCELADOS").Cells(j, 46) = Application.VLookup(Sheets("CANCELADOS").Cells(j, 1), Sheets("ACTIVOS").Range("A:A"), 1, False)
            Sheets("CANCELADOS").Cells(j, 47) = Application.VLookup(Sheets("CANCELADOS").Cells(j, 1), Sheets("DINAMICA").Range("A:B"), 2, False)
            Sheets("CANCELADOS").Cells(j, 48) = Application.VLookup(Sheets("CANCELADOS").Cells(j, 1), Sheets("VALIDA CANCELADOS").Range("C:C"), 1, False)
        Next j
    End With
         '////////////////////
        '//////////////////// trae restructurados
         If Sheets("CONTABILIZADOS").Cells(2, 2).Value = "NUEVOS" Then
                path_Bd = "D:\PRUEBAS INFORMES\BASEtRABAJO.accdb"
                cnn.Provider = "Microsoft.ACE.OLEDB.12.0"
                cnn.Properties("Data Source") = path_Bd
                cnn.Properties("Jet OLEDB:Database Password") = ""
                cnn.Open
                Criterio = Sheets("CONTABILIZADOS").Cells(4, 2)
                strSQL = "SELECT * FROM [DEFINITIVA_RESTRUCTURADOS] WHERE [CONVENIO] = '" & Criterio & "'"
                recSet.Open strSQL, cnn
                limpiarDatos = Sheets("Restructurados").Range("A" & Rows.Count).End(xlUp).Row
                
                Sheets("Restructurados").Cells(2, 1).CopyFromRecordset recSet
'
                ColumnaVariacion = Sheets("Restructurados").Range("A" & Rows.Count).End(xlUp).Row
                With Sheets("Restructurados")
                    For i = 2 To ColumnaVariacion
                        .Cells(i, 2).Value = CLng(.Cells(i, 2))
                        .Cells(i, 3).Value = CLngLng(.Cells(i, 3))
                        .Cells(i, 4).Value = CLngLng(.Cells(i, 4))
                    Next i
                End With
                recSet.Close: Set recSet = Nothing
                cnn.Close: Set cnn = Nothing
                Sheets("Restructurados").Visible = xlSheetVisible
                Call validaRestructurados
                FECHA = Right(Sheets("CONTABILIZADOS").Range("B3").Value, 4) & Mid(Sheets("CONTABILIZADOS").Range("B3").Value, 4, 2)
                
        Else
                Sheets("Restructurados").Visible = xlSheetVeryHidden
        End If
    
    ' conexion a variacion de coutas
    '//////////////////////////////////////////////
     If Sheets("CONTABILIZADOS").Cells(2, 2).Value = "NUEVOS" Then
            Call abrirVariacion
            ColumnaVariacion = Sheets("VARIACION").Range("A" & Rows.Count).End(xlUp).Row
                For i = 4 To ColumnaVariacion
                With Sheets("VARIACION")
                    Sheets("VARIACION").Cells(i, 11) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:AN"), 39, 0)
                    Sheets("VARIACION").Cells(i, 12) = Application.VLookup(Sheets("VARIACION").Cells(i, 4), Sheets("VALIDA CANCELADOS").Range("E:F"), 2, 0)
                    If Not IsError(.Cells(i, 12)) Then
                        Sheets("VARIACION").Cells(i, 12) = Val(Sheets("VARIACION").Cells(i, 12).Value)
                        Sheets("VARIACION").Cells(i, 13) = Val(.Cells(i, 12) - .Cells(i, 7))
                        Else
                         Sheets("VARIACION").Cells(i, 13) = "NO"
                    End If
                        Sheets("VARIACION").Cells(i, 14) = Application.VLookup(Sheets("VARIACION").Cells(i, 4), Sheets("DINAMICA").Range("A:B"), 2, 0)
                    Sheets("DINAMICA").Visible = xlSheetHidden
                    If Not IsError(.Cells(i, 14)) Then
                        Sheets("VARIACION").Cells(i, 15) = Sheets("VARIACION").Cells(i, 14) - Sheets("VARIACION").Cells(i, 7)
                        Else
                            Sheets("VARIACION").Cells(i, 15) = "NO"
                    End If
                        Sheets("VARIACION").Cells(i, 17) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:E"), 4, 0)
                        Sheets("VARIACION").Cells(i, 18) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:AT"), 45, 0)
                        Sheets("VARIACION").Cells(i, 19) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:W"), 22, 0)
                        Sheets("VARIACION").Cells(i, 20) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("Restructurados").Range("I:J"), 2, 0)
                        Sheets("VARIACION").Cells(i, 21) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:Q"), 16, 0)
                        Sheets("VARIACION").Cells(i, 22) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:Z"), 25, 0)
                        Sheets("VARIACION").Cells(i, 23) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:P"), 15, 0)
                        Sheets("VARIACION").Cells(i, 24) = Application.VLookup(Sheets("VARIACION").Cells(i, 1), Sheets("CONTABILIZADOS").Range("B:O"), 14, 0)
                        
                End With
                Next i
        Else
            Sheets("VARIACION").Visible = xlSheetHidden
        End If

End If
                path_Bd = "D:\PRUEBAS INFORMES\BASEtRABAJO.accdb"
                cnn.Provider = "Microsoft.ACE.OLEDB.12.0"
                cnn.Properties("Data Source") = path_Bd
                cnn.Properties("Jet OLEDB:Database Password") = ""
                cnn.Open
                Criterio = Sheets("CONTABILIZADOS").Cells(4, 2).Value
                strSQL = "SELECT * FROM [novedades] WHERE [ALIAS] like '%" & Criterio & "%'"
                recSet.Open strSQL, cnn
                limpiarDatos = Sheets("NOVEDADES").Range("A" & Rows.Count).End(xlUp).Row
                
                Sheets("NOVEDADES").Cells(4, 1).CopyFromRecordset recSet
                ColumnaVariacion = Sheets("NOVEDADES").Range("A" & Rows.Count).End(xlUp).Row
                For i = 4 To ColumnaVariacion
                    Sheets("NOVEDADES").Cells(i, 1).Value = Val(Sheets("NOVEDADES").Cells(i, 1))
                    Sheets("NOVEDADES").Cells(i, 8).Value = Val(Sheets("NOVEDADES").Cells(i, 8))
                Next i
                recSet.Close: Set recSet = Nothing
                cnn.Close: Set cnn = Nothing
                
                Sheets("NOVEDADES").Visible = xlSheetVisible
   
        Call cruceNovedades
        FECHA = Right(Sheets("CONTABILIZADOS").Range("B3").Value, 4) & Mid(Sheets("CONTABILIZADOS").Range("B3").Value, 4, 2)
        
     Call ValidacionDatosCancelados
    Sheets("DINAMICA").Visible = xlSheetHidden
    Call validaRestructurados
    
    Hoja1.Range("ax2:ay600").ClearContents
With Sheets("ACTIVOS")
.Select
    For i = 2 To CantidadActivos
       If UCase(Trim(.Cells(i, 46))) = "NUEVO" Or UCase(Trim(.Cells(i, 46))) = "REPORTAR" Or UCase(Trim(.Cells(i, 46))) = "RETANQUEO" Or UCase(Trim(.Cells(i, 46))) = "ACTUALIZACION" Or UCase(Trim(.Cells(i, 46))) = "VARIACION COUTA" Then
            If Not IsError(Application.VLookup(.Cells(i, 1), Sheets("VALIDA CANCELADOS").Range("e:f"), 2, 0)) Then
                .Cells(i, 50) = Application.VLookup(.Cells(i, 1), Sheets("VALIDA CANCELADOS").Range("e:f"), 2, 0)
                .Cells(i, 51) = .Cells(i, 45) - .Cells(i, 50)
            End If
       End If
    Next i
End With
Call Insolvencias
    Sheets("ACTIVOS").Range("AW1").AutoFilter Field:=49, Criteria1:="#N/A"
    Sheets("ACTIVOS").Range("AT1").AutoFilter Field:=46, Criteria1:="<>*NO REPORTAR*"
    
 '///////////////////////////////////REALIZAR IF DE CANCELADOS PARA CALCULAR LAS REGLAS//////////////////////////////////
 
 
'CANNN = MsgBox("¿Cancelados Del Periodo?", vbYesNo + vbQuestion, "Opplus Analisis Cancelados")

'If CANNN = vbNo Then
If Sheets("CONTABILIZADOS").Range("F6") <> "SI" Then
Sheets("CANCELADOS").Range("AT1").AutoFilter Field:=46, Criteria1:="#N/A"
Sheets("CANCELADOS").Range("AV1").AutoFilter Field:=48, Criteria1:="#N/A"
Sheets("CANCELADOS").Range("AU1").AutoFilter Field:=47, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>#N/A"

Else

End If
    'Sheets("VARIACION").Range("T1").AutoFilter Field:=20, Criteria1:="#N/A", Operator:=xlOr, Criteria2:=FECHA
    'Sheets("VARIACION").Range("T1").AutoFilter Field:=20, Criteria1:=FECHA
    'Sheets("VARIACION").Range("M1").AutoFilter Field:=13, Criteria1:="<>0"
    Sheets("VARIACION").Range("K2").AutoFilter Field:=11, Criteria1:="<>1"
    Sheets("Restructurados").Range("J1").AutoFilter Field:=10, Criteria1:=FECHA
    Call VALIDACIONDATOS
    Sheets("CONTABILIZADOS").Select
    MsgBox "ANALISIS FINALIZADO", vbInformation, "Opplus"
    Application.ScreenUpdating = True
    
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
    'Else
'MIO'Application.DisplayAlerts = False
'MIO'ThisWorkbook.Close
'MIO'End If
    
End Sub
Sub validaRestructurados()
Dim largActivos, i As Long
    Sheets("ACTIVOS").Select
    largActivos = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row
    'FECHA = Right(Sheets("CONTABILIZADOS").Range("B3").Value, 4) & Mid(Sheets("CONTABILIZADOS").Range("B3").Value, 4, 2)
     For i = 2 To largActivos
        With Sheets("ACTIVOS")
            .Cells(i, 49) = Application.VLookup(.Cells(i, 2), Sheets("Restructurados").Range("I:J"), 2, 0) ' "RESTRUCTURADO"
        End With
    Next i
End Sub
Sub OrdenarActivos()
Dim LargoActivos As Long
    Sheets("ACTIVOS").Select
    LargoActivos = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row
    Selection.End(xlToLeft).Select
    Range("A1:AS1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ACTIVOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ACTIVOS").Sort.SortFields.Add Key:=Range("A2:A" & LargoActivos) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ACTIVOS").Sort.SortFields.Add Key:=Range("O2:O" & LargoActivos) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ACTIVOS").Sort
        .SetRange Range("A1:AS" & LargoActivos)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F2").Select
End Sub
Sub OrdenarCancelados()
Dim LargoCancelado As Long
    Sheets("CANCELADOS").Select
    LargoCancelado = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
    Range("A1:AS1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("CANCELADOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CANCELADOS").Sort.SortFields.Add Key:=Range( _
        "A2:A" & LargoCancelado), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("CANCELADOS").Sort.SortFields.Add Key:=Range( _
        "AG2:AG" & LargoCancelado), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("CANCELADOS").Sort
        .SetRange Range("A1:AS" & LargoCancelado)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub VALIDACIONDATOS()

    Sheets("ACTIVOS").Range("AT1").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=NOVEDADES!$XFD$1:$XFD$8"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub ValidacionDatosCancelados()

'
Sheets("CANCELADOS").Select
    Range("AW2").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=NOVEDADES!$XFB$1:$XFB$2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

