<%@ Language="VBScript" %>
<%
' ============================================================
' IsabSoft API - registrar.asp
' Recibe datos del chat IA y los guarda en Access
' Método: POST · Responde JSON
' ============================================================
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

' Permitir llamadas desde el navegador (CORS local)
Response.AddHeader "Access-Control-Allow-Origin", "*"
Response.AddHeader "Access-Control-Allow-Methods", "POST, OPTIONS"
Response.AddHeader "Access-Control-Allow-Headers", "Content-Type"

If Request.ServerVariables("REQUEST_METHOD") = "OPTIONS" Then
    Response.End
End If
%>
<!-- #include file="../config.asp" -->
<%
On Error Resume Next

' ── Leer datos del POST ──
Dim body, json
body = ""
Dim bs : bs = Request.TotalBytes
If bs > 0 Then
    Dim bStream
    Set bStream = Server.CreateObject("ADODB.Stream")
    bStream.Open
    bStream.Type = 1
    bStream.Write Request.BinaryRead(bs)
    bStream.Position = 0
    bStream.Type = 2
    bStream.Charset = "UTF-8"
    body = bStream.ReadText
    bStream.Close
    Set bStream = Nothing
End If

' ── Función para extraer valor JSON simple ──
Function GetJSON(texto, clave)
    Dim pos, posInicio, posFin, valor
    pos = InStr(texto, """" & clave & """")
    If pos = 0 Then GetJSON = "" : Exit Function
    posInicio = InStr(pos, texto, ":")
    If posInicio = 0 Then GetJSON = "" : Exit Function
    posInicio = posInicio + 1
    Do While Mid(texto, posInicio, 1) = " " : posInicio = posInicio + 1 : Loop
    If Mid(texto, posInicio, 1) = """" Then
        posInicio = posInicio + 1
        posFin = InStr(posInicio, texto, """")
        valor = Mid(texto, posInicio, posFin - posInicio)
    Else
        posFin = posInicio
        Do While posFin <= Len(texto)
            Dim c : c = Mid(texto, posFin, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            posFin = posFin + 1
        Loop
        valor = Trim(Mid(texto, posInicio, posFin - posInicio))
    End If
    GetJSON = valor
End Function

Function Esc(s)
    Esc = Replace(Replace(s, "'", "''"), """", "")
End Function

Function NullSi(s)
    If s = "" Or s = "null" Or s = "undefined" Then NullSi = "NULL" Else NullSi = "'" & Esc(s) & "'"
End Function

Function NumSi(s)
    If s = "" Or s = "null" Or s = "undefined" Then NumSi = "0" Else NumSi = s
End Function

' ── Extraer campos del JSON ──
Dim accion, tabla, fecha, tipoAnimal, cantidad, pesoPromedio
Dim precioPorKg, totalGs, proveedor, cliente, lote, protocolo
Dim semen, tecnico, descripcion, categoria, unidad, mensajeOriginal
Dim registradoIA

accion        = GetJSON(body, "accion")
tabla         = GetJSON(body, "tabla")
fecha         = GetJSON(body, "fecha")
tipoAnimal    = GetJSON(body, "tipoAnimal")
cantidad      = NumSi(GetJSON(body, "cantidad"))
pesoPromedio  = NumSi(GetJSON(body, "pesoPromedio"))
precioPorKg   = NumSi(GetJSON(body, "precioPorKg"))
totalGs       = NumSi(GetJSON(body, "totalGs"))
proveedor     = GetJSON(body, "proveedor")
cliente       = GetJSON(body, "cliente")
lote          = GetJSON(body, "lote")
protocolo     = GetJSON(body, "protocolo")
semen         = GetJSON(body, "semen")
tecnico       = GetJSON(body, "tecnico")
descripcion   = GetJSON(body, "descripcion")
categoria     = GetJSON(body, "categoria")
unidad        = GetJSON(body, "unidad")
mensajeOriginal = GetJSON(body, "mensajeOriginal")
registradoIA  = "1"

If fecha = "" Then fecha = Now()

Dim conn, sql, nuevoID
Set conn = AbrirConexion()

If Err.Number <> 0 Then
    Response.Write "{""ok"":false,""error"":""" & Err.Description & """}"
    Response.End
End If

Err.Clear

' ── Insertar según tabla ──
Select Case tabla

  Case "ComprasAnimales"
    sql = "INSERT INTO ComprasAnimales " & _
          "(Fecha, TipoAnimal, Cantidad, PesoPromedio, PrecioPorKg, TotalGs, Proveedor, LoteDestino, MensajeIA, RegistradoPorIA, Sincronizado) VALUES " & _
          "(#" & fecha & "#, " & NullSi(tipoAnimal) & ", " & cantidad & ", " & pesoPromedio & ", " & precioPorKg & ", " & totalGs & ", " & _
          NullSi(proveedor) & ", " & NullSi(lote) & ", " & NullSi(mensajeOriginal) & ", " & registradoIA & ", 1)"

  Case "VentasAnimales"
    sql = "INSERT INTO VentasAnimales " & _
          "(Fecha, TipoAnimal, Cantidad, PesoPromedio, PrecioPorKg, TotalGs, Cliente, MensajeIA, RegistradoPorIA, Sincronizado) VALUES " & _
          "(#" & fecha & "#, " & NullSi(tipoAnimal) & ", " & cantidad & ", " & pesoPromedio & ", " & precioPorKg & ", " & totalGs & ", " & _
          NullSi(cliente) & ", " & NullSi(mensajeOriginal) & ", " & registradoIA & ", 1)"

  Case "ComprasInsumos"
    sql = "INSERT INTO ComprasInsumos " & _
          "(Fecha, Descripcion, Categoria, Cantidad, Unidad, PrecioUnitario, TotalGs, Proveedor, MensajeIA, RegistradoPorIA, Sincronizado) VALUES " & _
          "(#" & fecha & "#, " & NullSi(descripcion) & ", " & NullSi(categoria) & ", " & cantidad & ", " & NullSi(unidad) & ", " & _
          precioPorKg & ", " & totalGs & ", " & NullSi(proveedor) & ", " & NullSi(mensajeOriginal) & ", " & registradoIA & ", 1)"

  Case "Inseminaciones"
    sql = "INSERT INTO Inseminaciones " & _
          "(Fecha, Lote, CantidadVacas, Protocolo, Semen, Tecnico, MensajeIA, RegistradoPorIA, Sincronizado) VALUES " & _
          "(#" & fecha & "#, " & NullSi(lote) & ", " & cantidad & ", " & NullSi(protocolo) & ", " & _
          NullSi(semen) & ", " & NullSi(tecnico) & ", " & NullSi(mensajeOriginal) & ", " & registradoIA & ", 1)"

  Case Else
    Response.Write "{""ok"":false,""error"":""Tabla no reconocida: " & tabla & """}"
    conn.Close : Set conn = Nothing : Response.End

End Select

conn.Execute sql

If Err.Number <> 0 Then
    Dim errMsg : errMsg = Err.Description
    conn.Close : Set conn = Nothing
    Response.Write "{""ok"":false,""error"":""" & errMsg & """}"
    Response.End
End If

' Obtener ID del registro insertado
Dim rsID : Set rsID = conn.Execute("SELECT @@IDENTITY AS NuevoID")
If Not rsID.EOF Then nuevoID = rsID("NuevoID")
rsID.Close : Set rsID = Nothing

conn.Close : Set conn = Nothing

Response.Write "{""ok"":true,""id"":" & nuevoID & ",""tabla"":""" & tabla & """,""mensaje"":""Registrado en Access correctamente""}"
%>
