<%@ Language="VBScript" %>
<%
' ============================================================
' IsabSoft - Instalador: crea todas las tablas en Access
' Accedé a esta página UNA SOLA VEZ para crear la BD
' URL: http://localhost/isabsoft/instalador.asp
' ============================================================
%>
<!-- #include file="config.asp" -->
<%
Dim conn, sql, errores, creadas
errores = ""
creadas = 0

On Error Resume Next

Set conn = AbrirConexion()

If Err.Number <> 0 Then
    Response.Write "<h2 style='color:red'>Error al conectar: " & Err.Description & "</h2>"
    Response.Write "<p>Verificá que la ruta de la BD en config.asp sea correcta.</p>"
    Response.End
End If

Err.Clear

' ── 1. ANIMALES (stock general) ──
sql = "CREATE TABLE Animales (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Fecha DATETIME, " & _
      "Tipo TEXT(50), " & _
      "Raza TEXT(100), " & _
      "Cantidad INTEGER, " & _
      "PesoPromedio DOUBLE, " & _
      "Lote TEXT(50), " & _
      "Estado TEXT(30), " & _
      "Establecimiento TEXT(100), " & _
      "Observaciones MEMO, " & _
      "Sincronizado INTEGER DEFAULT 0, " & _
      "FechaSync DATETIME, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "Animales: " & Err.Description & "<br>" : Err.Clear

' ── 2. COMPRAS DE ANIMALES ──
sql = "CREATE TABLE ComprasAnimales (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Fecha DATETIME, " & _
      "TipoAnimal TEXT(50), " & _
      "Raza TEXT(100), " & _
      "Cantidad INTEGER, " & _
      "PesoPromedio DOUBLE, " & _
      "PrecioPorKg DOUBLE, " & _
      "TotalGs DOUBLE, " & _
      "Proveedor TEXT(150), " & _
      "RucProveedor TEXT(30), " & _
      "TelefonoProveedor TEXT(30), " & _
      "LoteDestino TEXT(50), " & _
      "FechaIngreso DATETIME, " & _
      "FormaPago TEXT(50), " & _
      "Observaciones MEMO, " & _
      "RegistradoPorIA INTEGER DEFAULT 0, " & _
      "MensajeIA MEMO, " & _
      "Sincronizado INTEGER DEFAULT 0, " & _
      "FechaSync DATETIME, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "ComprasAnimales: " & Err.Description & "<br>" : Err.Clear

' ── 3. VENTAS DE ANIMALES ──
sql = "CREATE TABLE VentasAnimales (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Fecha DATETIME, " & _
      "TipoAnimal TEXT(50), " & _
      "Cantidad INTEGER, " & _
      "PesoPromedio DOUBLE, " & _
      "PrecioPorKg DOUBLE, " & _
      "TotalGs DOUBLE, " & _
      "Cliente TEXT(150), " & _
      "RucCliente TEXT(30), " & _
      "TelefonoCliente TEXT(30), " & _
      "Destino TEXT(150), " & _
      "FormaPago TEXT(50), " & _
      "FechaDespacho DATETIME, " & _
      "Transportista TEXT(100), " & _
      "Observaciones MEMO, " & _
      "RegistradoPorIA INTEGER DEFAULT 0, " & _
      "MensajeIA MEMO, " & _
      "Sincronizado INTEGER DEFAULT 0, " & _
      "FechaSync DATETIME, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "VentasAnimales: " & Err.Description & "<br>" : Err.Clear

' ── 4. COMPRAS DE INSUMOS ──
sql = "CREATE TABLE ComprasInsumos (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Fecha DATETIME, " & _
      "Descripcion TEXT(200), " & _
      "Categoria TEXT(80), " & _
      "Cantidad DOUBLE, " & _
      "Unidad TEXT(30), " & _
      "PrecioUnitario DOUBLE, " & _
      "TotalGs DOUBLE, " & _
      "Proveedor TEXT(150), " & _
      "Observaciones MEMO, " & _
      "RegistradoPorIA INTEGER DEFAULT 0, " & _
      "MensajeIA MEMO, " & _
      "Sincronizado INTEGER DEFAULT 0, " & _
      "FechaSync DATETIME, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "ComprasInsumos: " & Err.Description & "<br>" : Err.Clear

' ── 5. INSEMINACIONES ──
sql = "CREATE TABLE Inseminaciones (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Fecha DATETIME, " & _
      "Lote TEXT(50), " & _
      "CantidadVacas INTEGER, " & _
      "Protocolo TEXT(100), " & _
      "Semen TEXT(150), " & _
      "PajuelasUsadas INTEGER, " & _
      "Tecnico TEXT(100), " & _
      "FechaDiagnostico DATETIME, " & _
      "ResultadoDiagnostico TEXT(30), " & _
      "TasaPrenez DOUBLE, " & _
      "Observaciones MEMO, " & _
      "RegistradoPorIA INTEGER DEFAULT 0, " & _
      "MensajeIA MEMO, " & _
      "Sincronizado INTEGER DEFAULT 0, " & _
      "FechaSync DATETIME, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "Inseminaciones: " & Err.Description & "<br>" : Err.Clear

' ── 6. COLA SYNC OFFLINE ──
sql = "CREATE TABLE ColaSyncOffline (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "FechaRegistro DATETIME, " & _
      "Accion TEXT(50), " & _
      "Tabla TEXT(50), " & _
      "DatosJSON MEMO, " & _
      "MensajeOriginal MEMO, " & _
      "Enviado INTEGER DEFAULT 0, " & _
      "Intentos INTEGER DEFAULT 0, " & _
      "FechaEnvio DATETIME, " & _
      "Error MEMO)"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "ColaSyncOffline: " & Err.Description & "<br>" : Err.Clear

' ── 7. CLIENTES ──
sql = "CREATE TABLE Clientes (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Nombre TEXT(150), " & _
      "RUC TEXT(30), " & _
      "Telefono TEXT(30), " & _
      "Email TEXT(100), " & _
      "Direccion MEMO, " & _
      "Ciudad TEXT(80), " & _
      "Activo INTEGER DEFAULT 1, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "Clientes: " & Err.Description & "<br>" : Err.Clear

' ── 8. PROVEEDORES ──
sql = "CREATE TABLE Proveedores (" & _
      "ID AUTOINCREMENT PRIMARY KEY, " & _
      "Nombre TEXT(150), " & _
      "RUC TEXT(30), " & _
      "Telefono TEXT(30), " & _
      "Email TEXT(100), " & _
      "Direccion MEMO, " & _
      "Ciudad TEXT(80), " & _
      "Activo INTEGER DEFAULT 1, " & _
      "FechaCreacion DATETIME DEFAULT NOW())"
conn.Execute sql
If Err.Number = 0 Then creadas = creadas + 1 Else errores = errores & "Proveedores: " & Err.Description & "<br>" : Err.Clear

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>IsabSoft · Instalador</title>
<style>
body{font-family:'Segoe UI',sans-serif;background:#f0f4f9;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}
.box{background:#fff;border-radius:20px;padding:40px;max-width:500px;width:100%;box-shadow:0 8px 32px rgba(0,0,0,.12);}
.logo{font-size:1.5rem;font-weight:900;color:#1565c0;margin-bottom:6px;}
.logo em{font-style:normal;color:#1e88e5;}
h2{font-size:1.1rem;margin:20px 0 16px;color:#333;}
.ok{background:#e8f5e9;border-left:4px solid #2e7d32;border-radius:8px;padding:12px 16px;margin-bottom:8px;font-size:.88rem;color:#1b5e20;}
.err{background:#fce4ec;border-left:4px solid #c62828;border-radius:8px;padding:12px 16px;margin-bottom:8px;font-size:.88rem;color:#b71c1c;}
.total{background:#e3f2fd;border-radius:12px;padding:16px;text-align:center;margin-top:20px;font-weight:700;color:#1565c0;font-size:1rem;}
.btn{display:block;width:100%;padding:14px;background:#1565c0;color:#fff;border:none;border-radius:30px;font-size:1rem;font-weight:800;cursor:pointer;margin-top:16px;text-align:center;text-decoration:none;}
.warn{background:#fff8e1;border-radius:10px;padding:14px;font-size:.82rem;color:#e65100;margin-top:16px;border-left:4px solid #f57c00;}
</style>
</head>
<body>
<div class="box">
  <div class="logo">Isab<em>Soft</em></div>
  <p style="color:#888;font-size:.8rem;">Instalador de base de datos · v2.0</p>
  <h2>Resultado de instalación</h2>

  <% If errores = "" Then %>
    <div class="ok">✅ Animales — creada correctamente</div>
    <div class="ok">✅ ComprasAnimales — creada correctamente</div>
    <div class="ok">✅ VentasAnimales — creada correctamente</div>
    <div class="ok">✅ ComprasInsumos — creada correctamente</div>
    <div class="ok">✅ Inseminaciones — creada correctamente</div>
    <div class="ok">✅ ColaSyncOffline — creada correctamente</div>
    <div class="ok">✅ Clientes — creada correctamente</div>
    <div class="ok">✅ Proveedores — creada correctamente</div>
    <div class="total">🎉 <%=creadas%> tablas creadas exitosamente en Access</div>
  <% Else %>
    <div class="err">⚠️ Algunas tablas ya existen o hubo errores:<br><br><%=errores%></div>
    <div class="total"><%=creadas%> tablas procesadas</div>
  <% End If %>

  <div class="warn">
    ⚠️ <strong>Importante:</strong> Esta página solo debe ejecutarse una vez.<br>
    Después de instalar, eliminá o renombrá este archivo por seguridad.
  </div>
  <a href="index.html" class="btn">→ Ir a la aplicación</a>
</div>
</body>
</html>
