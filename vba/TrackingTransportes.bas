Option Explicit

' CONFIGURACIÓN:
' 1) Cambia SERVER_NAME por el nombre de tu instancia SQL Server.
' 2) Usa autenticación de Windows. Excel debe ejecutarse con un usuario autorizado en SQL.
Private Const SERVER_NAME As String = "localhost"
Private Const DATABASE_NAME As String = "TrackingTransportes"

Private Function GetConnection() As Object
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Provider=MSOLEDBSQL;Server=" & SERVER_NAME & _
        ";Database=" & DATABASE_NAME & ";Trusted_Connection=Yes;Encrypt=Optional;"
    cn.Open
    Set GetConnection = cn
End Function

Private Function SheetValue(ByVal sheetName As String, ByVal cellAddress As String) As String
    SheetValue = Trim$(CStr(ThisWorkbook.Worksheets(sheetName).Range(cellAddress).Value))
End Function

Private Function SqlDateOrNull(ByVal valueText As String) As String
    If Len(valueText) = 0 Then
        SqlDateOrNull = "NULL"
    Else
        SqlDateOrNull = "'" & Format$(CDate(valueText), "yyyy-mm-dd hh:nn:ss") & "'"
    End If
End Function

Public Sub RegistrarUnidad()
    On Error GoTo ErrHandler
    Dim cn As Object, cmd As Object
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Registro")
    If SheetValue("Registro", "B4") = "" Or SheetValue("Registro", "B5") = "" Then
        MsgBox "Completa como mínimo la placa y el tipo de unidad.", vbExclamation
        Exit Sub
    End If

    Set cn = GetConnection()
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "INSERT INTO dbo.Unidades (Placa,TipoUnidad,Marca,Modelo,CapacidadTon,Estado) VALUES (?,?,?,?,?,?)"
    cmd.Parameters.Append cmd.CreateParameter("@Placa", 202, 1, 15, SheetValue("Registro", "B4"))
    cmd.Parameters.Append cmd.CreateParameter("@Tipo", 202, 1, 40, SheetValue("Registro", "B5"))
    cmd.Parameters.Append cmd.CreateParameter("@Marca", 202, 1, 50, SheetValue("Registro", "B6"))
    cmd.Parameters.Append cmd.CreateParameter("@Modelo", 202, 1, 50, SheetValue("Registro", "B7"))
    If IsNumeric(ws.Range("B8").Value) And ws.Range("B8").Value <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@Capacidad", 131, 1, , CDbl(ws.Range("B8").Value))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Capacidad", 131, 1, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 202, 1, 20, IIf(SheetValue("Registro", "B9") = "", "Disponible", SheetValue("Registro", "B9")))
    cmd.Execute
    cn.Close
    MsgBox "Unidad registrada correctamente en SQL Server.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "No se pudo registrar la unidad: " & Err.Description, vbCritical
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub

Public Sub RegistrarViaje()
    On Error GoTo ErrHandler
    Dim cn As Object, cmd As Object, rs As Object
    Dim unidadID As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Registro")
    If SheetValue("Registro", "B12") = "" Or SheetValue("Registro", "B13") = "" Or _
       SheetValue("Registro", "B14") = "" Or SheetValue("Registro", "B15") = "" Or _
       SheetValue("Registro", "B16") = "" Or SheetValue("Registro", "B17") = "" Then
        MsgBox "Completa código, placa, conductor, origen, destino y fecha de salida.", vbExclamation
        Exit Sub
    End If

    Set cn = GetConnection()
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT UnidadID FROM dbo.Unidades WHERE Placa=?"
    cmd.Parameters.Append cmd.CreateParameter("@Placa", 202, 1, 15, SheetValue("Registro", "B13"))
    Set rs = cmd.Execute
    If rs.EOF Then
        MsgBox "La placa no existe en SQL. Registra primero la unidad.", vbExclamation
        rs.Close: cn.Close
        Exit Sub
    End If
    unidadID = CLng(rs.Fields("UnidadID").Value)
    rs.Close

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "INSERT INTO dbo.Viajes (CodigoViaje,UnidadID,Conductor,DocumentoConductor,Origen,Destino,Cliente,FechaSalida,ETA,Estado,Observaciones,UsuarioRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
    cmd.Parameters.Append cmd.CreateParameter("@Codigo", 202, 1, 30, SheetValue("Registro", "B12"))
    cmd.Parameters.Append cmd.CreateParameter("@UnidadID", 3, 1, , unidadID)
    cmd.Parameters.Append cmd.CreateParameter("@Conductor", 202, 1, 120, SheetValue("Registro", "B14"))
    cmd.Parameters.Append cmd.CreateParameter("@Documento", 202, 1, 20, SheetValue("Registro", "B18"))
    cmd.Parameters.Append cmd.CreateParameter("@Origen", 202, 1, 150, SheetValue("Registro", "B15"))
    cmd.Parameters.Append cmd.CreateParameter("@Destino", 202, 1, 150, SheetValue("Registro", "B16"))
    cmd.Parameters.Append cmd.CreateParameter("@Cliente", 202, 1, 120, SheetValue("Registro", "B19"))
    cmd.Parameters.Append cmd.CreateParameter("@Salida", 135, 1, , CDate(ws.Range("B17").Value))
    If IsDate(ws.Range("B20").Value) Then
        cmd.Parameters.Append cmd.CreateParameter("@ETA", 135, 1, , CDate(ws.Range("B20").Value))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ETA", 135, 1, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 202, 1, 25, IIf(SheetValue("Registro", "B21") = "", "Programado", SheetValue("Registro", "B21")))
    cmd.Parameters.Append cmd.CreateParameter("@Obs", 202, 1, 500, SheetValue("Registro", "B22"))
    cmd.Parameters.Append cmd.CreateParameter("@Usuario", 202, 1, 100, Environ$("Username"))
    cmd.Execute
    cn.Close
    MsgBox "Viaje registrado correctamente en SQL Server.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "No se pudo registrar el viaje: " & Err.Description, vbCritical
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub

Public Sub RegistrarEventoTracking()
    On Error GoTo ErrHandler
    Dim cn As Object, cmd As Object, rs As Object
    Dim viajeID As Long
    If SheetValue("Tracking", "B4") = "" Or SheetValue("Tracking", "B5") = "" Or SheetValue("Tracking", "B6") = "" Then
        MsgBox "Completa código de viaje, ubicación y estado.", vbExclamation
        Exit Sub
    End If
    Set cn = GetConnection()
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "SELECT ViajeID FROM dbo.Viajes WHERE CodigoViaje=?"
    cmd.Parameters.Append cmd.CreateParameter("@Codigo", 202, 1, 30, SheetValue("Tracking", "B4"))
    Set rs = cmd.Execute
    If rs.EOF Then
        MsgBox "No se encontró el código de viaje.", vbExclamation
        rs.Close: cn.Close
        Exit Sub
    End If
    viajeID = CLng(rs.Fields("ViajeID").Value)
    rs.Close

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "INSERT INTO dbo.EventosTracking (ViajeID,FechaEvento,Ubicacion,Estado,Kilometraje,Observacion,RegistradoPor) VALUES (?,?,?,?,?,?,?)"
    cmd.Parameters.Append cmd.CreateParameter("@ViajeID", 3, 1, , viajeID)
    If IsDate(ThisWorkbook.Worksheets("Tracking").Range("B7").Value) Then
        cmd.Parameters.Append cmd.CreateParameter("@Fecha", 135, 1, , CDate(ThisWorkbook.Worksheets("Tracking").Range("B7").Value))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Fecha", 135, 1, , Now)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Ubicacion", 202, 1, 180, SheetValue("Tracking", "B5"))
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 202, 1, 25, SheetValue("Tracking", "B6"))
    If IsNumeric(ThisWorkbook.Worksheets("Tracking").Range("B8").Value) And SheetValue("Tracking", "B8") <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@Km", 3, 1, , CLng(ThisWorkbook.Worksheets("Tracking").Range("B8").Value))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Km", 3, 1, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Obs", 202, 1, 500, SheetValue("Tracking", "B9"))
    cmd.Parameters.Append cmd.CreateParameter("@Usuario", 202, 1, 100, Environ$("Username"))
    cmd.Execute
    ' Mantener el estado actual del viaje actualizado
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = 1
    cmd.CommandText = "UPDATE dbo.Viajes SET Estado=? WHERE ViajeID=?"
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 202, 1, 25, SheetValue("Tracking", "B6"))
    cmd.Parameters.Append cmd.CreateParameter("@ViajeID", 3, 1, , viajeID)
    cmd.Execute
    cn.Close
    MsgBox "Evento de tracking guardado.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "No se pudo guardar el tracking: " & Err.Description, vbCritical
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub

Public Sub ConsultarTracking()
    On Error GoTo ErrHandler
    Dim cn As Object, rs As Object, ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("Consultas")
    Set cn = GetConnection()
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT ViajeID,CodigoViaje,Placa,Conductor,Cliente,Origen,Destino,FechaSalida,ETA,EstadoActual,UltimaActualizacion,UltimaUbicacion,Kilometraje,Observacion FROM dbo.vw_TrackingActual ORDER BY FechaSalida DESC", cn, 3, 1
    ws.Range("A5:N10000").ClearContents
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(4, i + 1).Value = rs.Fields(i).Name
        ws.Cells(4, i + 1).Font.Bold = True
    Next i
    If Not rs.EOF Then ws.Range("A5").CopyFromRecordset rs
    rs.Close: cn.Close
    ws.Columns("A:N").AutoFit
    MsgBox "Consulta completada. Revisa la hoja Consultas.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Error al consultar SQL Server: " & Err.Description, vbCritical
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
End Sub

Public Sub LimpiarFormulario()
    With ThisWorkbook.Worksheets("Registro")
        .Range("B4:B9,B12:B22").ClearContents
    End With
    With ThisWorkbook.Worksheets("Tracking")
        .Range("B4:B9").ClearContents
    End With
    MsgBox "Formulario limpio.", vbInformation
End Sub

Public Sub ProbarConexionSQL()
    On Error GoTo ErrHandler
    Dim cn As Object
    Set cn = GetConnection()
    MsgBox "Conexión correcta con la base de datos " & DATABASE_NAME, vbInformation
    cn.Close
    Exit Sub
ErrHandler:
    MsgBox "No se pudo conectar. Verifica SQL Server, el servidor configurado y el controlador OLE DB. Detalle: " & Err.Description, vbCritical
End Sub
