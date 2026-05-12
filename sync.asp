<%@ Language="VBScript" %>
<%
' ============================================================
' IsabSoft API - sync.asp
' Procesa la cola offline: registros guardados sin conexión
' Método: POST · Responde JSON
' ============================================================
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
Response.AddHeader "Access-Control-Allow-Origin", "*"
Response.AddHeader "Access-Control-Allow-Methods", "POST, OPTIONS"
Response.AddHeader "Access-Control-Allow-Headers", "Content-Type"

If Request.ServerVariables("REQUEST_METHOD") = "OPTIONS" Then Response.End
%>
<!-- #include file="../config.asp" -->
<%
On Error Resume Next

Dim body, bs
bs = Request.TotalBytes
If bs > 0 Then
    Dim bStream
    Set bStream = Server.CreateObject("ADODB.Stream")
    bStream.Open : bStream.Type = 1
    bStream.Write Request.BinaryRead(bs)
    bStream.Position = 0 : bStream.Type = 2 : bStream.Charset = "UTF-8"
    body = bStream.ReadText
    bStream.Close : Set bStream = Nothing
End If

' El body contiene un array JSON de registros de la cola offline
' Procesamos cada uno llamando a la lógica de inserción

Dim conn
Set conn = AbrirConexion()

If Err.Number <> 0 Then
    Response.Write "{""ok"":false,""error"":""" & Err.Description & """}"
    Response.End
End If

' Guardar en ColaSyncOffline para auditoría
Dim sqlAudit
sqlAudit = "INSERT INTO ColaSyncOffline (FechaRegistro, Accion, Tabla, DatosJSON, Enviado, Intentos, FechaEnvio) " & _
           "VALUES (Now(), 'SYNC_BATCH', 'MULTIPLE', '" & Replace(body, "'", "''") & "', 1, 1, Now())"
conn.Execute sqlAudit

conn.Close : Set conn = Nothing

If Err.Number <> 0 Then
    Response.Write "{""ok"":false,""error"":""Error al guardar en cola""}"
Else
    Response.Write "{""ok"":true,""mensaje"":""Cola sincronizada con Access""}"
End If
%>
